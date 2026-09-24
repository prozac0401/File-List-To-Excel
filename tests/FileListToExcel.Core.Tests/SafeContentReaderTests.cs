using System.Security.Cryptography;
using System.Text;
using FileListToExcel.Core;
using Xunit;

namespace FileListToExcel.Core.Tests;

public sealed class SafeContentReaderTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "FileListToExcel.SafeRead.Tests", Guid.NewGuid().ToString("N"));
    public SafeContentReaderTests() => Directory.CreateDirectory(directory);
    public void Dispose() { try { Directory.Delete(directory, true); } catch (IOException) { } }
    private string Source(string relative, byte[] contents)
    {
        string path = Path.Combine(directory, relative);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, contents);
        return path;
    }

    [Fact]
    public async Task ReadOnlySourceHashesWithoutChangingContentOrWriteTime()
    {
        byte[] contents = Encoding.UTF8.GetBytes("읽기 전용 원본 파일\n");
        string source = Source("읽기 전용.txt", contents);
        var attributes = File.GetAttributes(source);
        DateTime modified = File.GetLastWriteTimeUtc(source);
        File.SetAttributes(source, attributes | FileAttributes.ReadOnly);
        try
        {
            var hash = await new FileContentHasher().FullAsync(source, CancellationToken.None);
            Assert.Equal(Convert.ToHexString(SHA256.HashData(contents)), hash.Hash);
            Assert.Equal(contents.Length, hash.BytesRead);
            Assert.Equal(modified, File.GetLastWriteTimeUtc(source));
            Assert.Equal(contents, File.ReadAllBytes(source));
            Assert.True(File.GetAttributes(source).HasFlag(FileAttributes.ReadOnly));
        }
        finally { File.SetAttributes(source, attributes); }
    }

    [Fact]
    public async Task NativeReaderSupportsLongUnicodePaths()
    {
        string relative = Path.Combine(Enumerable.Repeat(new string('가', 48), 7).Append("원본 😀.bin").ToArray());
        byte[] contents = Encoding.UTF8.GetBytes("long unicode source");
        string source = Source(relative, contents);
        Assert.True(source.Length > 260);
        var hash = await new FileContentHasher().FullAsync(source, CancellationToken.None);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(contents)), hash.Hash);
    }

    [Fact]
    public async Task SourceAlreadyOpenForWritingIsNotRead()
    {
        string source = Source("writer.bin", [1, 2, 3, 4]);
        using var writer = new FileStream(source, FileMode.Open, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete);
        await Assert.ThrowsAsync<IOException>(() => new FileContentHasher().FullAsync(source, CancellationToken.None));
    }

    [Fact]
    public async Task OfflineSourceIsRejectedBeforeContentRead()
    {
        string source = Source("offline.bin", [1, 2, 3, 4]);
        var attributes = File.GetAttributes(source);
        try
        {
            File.SetAttributes(source, attributes | FileAttributes.Offline);
            Assert.True(File.GetAttributes(source).HasFlag(FileAttributes.Offline));
            var error = await Assert.ThrowsAsync<IOException>(() => new FileContentHasher().FullAsync(source, CancellationToken.None));
            Assert.Contains("Cloud-only", error.Message);
        }
        finally { File.SetAttributes(source, attributes); }
    }

    [Fact]
    public async Task SelectedFileInsideDirectoryJunctionIsRejected()
    {
        string target = Path.Combine(directory, "target");
        string source = Source(Path.Combine("target", "source.bin"), [1, 2, 3, 4]);
        string junction = Path.Combine(directory, "junction");
        var start = new System.Diagnostics.ProcessStartInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "System32", "cmd.exe"))
        {
            UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true
        };
        foreach (string arg in new[] { "/c", "mklink", "/J", junction, target }) start.ArgumentList.Add(arg);
        using var process = System.Diagnostics.Process.Start(start)!;
        await process.WaitForExitAsync();
        Assert.True(process.ExitCode == 0, await process.StandardError.ReadToEndAsync());
        try
        {
            var error = await Assert.ThrowsAsync<IOException>(() => new FileContentHasher().FullAsync(Path.Combine(junction, "source.bin"), CancellationToken.None));
            Assert.Contains("Reparse Point", error.Message);
            Assert.Equal([1, 2, 3, 4], File.ReadAllBytes(source));
        }
        finally { Directory.Delete(junction); }
    }

    [Fact]
    public async Task CancellationDuringLargeReadClosesSourceHandle()
    {
        string source = Path.Combine(directory, "large.bin");
        using (var file = new FileStream(source, FileMode.CreateNew, FileAccess.Write)) file.SetLength(256L * 1024 * 1024);
        using var cancellation = new CancellationTokenSource();
        var work = Task.Run(() => new FileContentHasher().FullAsync(source, cancellation.Token));
        bool reading = false;
        for (int attempt = 0; attempt < 500 && !work.IsCompleted; attempt++)
        {
            try { using var probe = new FileStream(source, FileMode.Open, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete); }
            catch (IOException) { reading = true; break; }
            await Task.Delay(1);
        }
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => work);
        Assert.True(reading, "The read handle must have opened before cancellation.");
        using var released = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.None);
        Assert.Equal(256L * 1024 * 1024, released.Length);
    }
}
