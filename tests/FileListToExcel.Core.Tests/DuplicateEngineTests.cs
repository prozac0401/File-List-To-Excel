using System.Security.Cryptography;
using FileListToExcel.Core;
using Xunit;
using Xunit.Abstractions;

namespace FileListToExcel.Core.Tests;

public sealed class DuplicateEngineTests(ITestOutputHelper output) : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "FileListToExcel.DuplicateTests", Guid.NewGuid().ToString("N"));
    public void Dispose() { try { if (Directory.Exists(directory)) Directory.Delete(directory, true); } catch (IOException) { } }

    private string FileAt(string name, byte[] contents)
    {
        string path = Path.Combine(directory, name);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, contents);
        return path;
    }
    private static FileEntry[] Entries(params string[] paths) => new FileScanner().Enumerate(new(paths, ScanMode.SelectedItems))
        .Select(item => item.Entry ?? throw new IOException(item.Error!.Message)).ToArray();
    private static byte[] Data(int size, int seed)
    {
        var bytes = new byte[size];
        new Random(seed).NextBytes(bytes);
        return bytes;
    }
    private void Stats(DuplicateResult result) => output.WriteLine($"files={result.Stats.ScannedFiles}; size candidates={result.Stats.SizeCandidates}; quick candidates={result.Stats.QuickCandidates}; quick hashes={result.Stats.QuickHashCount}; full hashes={result.Stats.FullHashCount}; bytes read={result.Stats.BytesRead}; elapsed={result.Stats.Elapsed.TotalMilliseconds:F2}ms; groups={result.Groups.Count}");

    [Fact]
    public async Task SameContentDifferentNamesAndExtensionsFormsOneGroup()
    {
        byte[] bytes = Data(4096, 1);
        string first = FileAt("원본 한글.txt", bytes), second = FileAt("README", bytes);
        var result = await new DuplicateDetector().DetectAsync(Entries(first, second, first));
        var group = Assert.Single(result.Groups);
        Assert.Equal(2, group.Files.Count);
        Assert.Equal(4096, group.RecoverableBytes);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(bytes)), group.FullSha256);
        Assert.Equal(2, result.Stats.FullHashCount); Assert.Equal(0, result.Stats.QuickHashCount);
        Assert.Equal(8192, result.Stats.BytesRead);
        Assert.Equal(bytes, File.ReadAllBytes(first)); Assert.Equal(bytes, File.ReadAllBytes(second));
        Stats(result);
    }

    [Fact]
    public async Task SameSizeDifferentContentsDoNotMatch()
    {
        var result = await new DuplicateDetector().DetectAsync(Entries(FileAt("first", Data(4096, 1)), FileAt("second", Data(4096, 2))));
        Assert.Empty(result.Groups); Assert.Equal(2, result.Stats.FullHashCount); Stats(result);
    }

    [Fact]
    public async Task DifferentSizesNeverOpenContents()
    {
        string first = FileAt("one", [1]), second = FileAt("two", [1, 2]);
        using var lockFirst = new FileStream(first, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var lockSecond = new FileStream(second, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        var result = await new DuplicateDetector(hasher: new ForbiddenHasher()).DetectAsync(Entries(first, second));
        Assert.Empty(result.Groups); Assert.Empty(result.Errors); Assert.Equal(0, result.Stats.SizeCandidates);
        Assert.Equal(0, result.Stats.BytesRead); Assert.Equal(0, result.Stats.FullHashCount); Stats(result);
    }

    [Fact]
    public async Task ZeroByteFilesGroupWithoutReadingContents()
    {
        string first = FileAt("empty1", []), second = FileAt("empty2", []), third = FileAt("empty3", []);
        using var fileLock = new FileStream(first, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        var result = await new DuplicateDetector(hasher: new ForbiddenHasher()).DetectAsync(Entries(first, second, third));
        var group = Assert.Single(result.Groups); Assert.Equal(3, group.Files.Count);
        Assert.Equal(Convert.ToHexString(SHA256.HashData([])), group.FullSha256);
        Assert.Equal(0, result.Stats.FullHashCount); Assert.Equal(0, result.Stats.BytesRead); Assert.Empty(result.Errors);
        Stats(result);
    }

    [Fact]
    public async Task OneMebibyteFilesAreFullyHashedExactlyOnce()
    {
        byte[] bytes = Data(FileContentHasher.SmallFileLimit, 44);
        var result = await new DuplicateDetector().DetectAsync(Entries(FileAt("first", bytes), FileAt("second", bytes)));
        Assert.Single(result.Groups); Assert.Equal(2, result.Stats.FullHashCount);
        Assert.Equal(0, result.Stats.QuickHashCount); Assert.Equal(2L * bytes.Length, result.Stats.BytesRead); Stats(result);
    }

    [Fact]
    public async Task LargeQuickFingerprintsPruneDifferentContentsBeforeFullHash()
    {
        byte[] duplicate = Data(2 * 1024 * 1024, 5);
        var result = await new DuplicateDetector().DetectAsync(Entries(FileAt("first", duplicate), FileAt("copy", duplicate), FileAt("different", Data(duplicate.Length, 6))));
        Assert.Equal(2, Assert.Single(result.Groups).Files.Count);
        Assert.Equal(3, result.Stats.SizeCandidates); Assert.Equal(2, result.Stats.QuickCandidates);
        Assert.Equal(3, result.Stats.QuickHashCount); Assert.Equal(2, result.Stats.FullHashCount);
        Assert.Equal(3L * 3 * FileContentHasher.QuickBlockSize + 2L * duplicate.Length, result.Stats.BytesRead); Stats(result);
    }

    [Fact]
    public async Task SampleCollisionMustStillPassFullSha256()
    {
        byte[] first = Data(2 * 1024 * 1024, 7), changed = (byte[])first.Clone();
        changed[256 * 1024] ^= 0xff; // Outside the start, centered middle, and end samples.
        var result = await new DuplicateDetector().DetectAsync(Entries(FileAt("first", first), FileAt("changed", changed)));
        Assert.Empty(result.Groups); Assert.Equal(2, result.Stats.QuickCandidates);
        Assert.Equal(2, result.Stats.QuickHashCount); Assert.Equal(2, result.Stats.FullHashCount); Stats(result);
    }

    [Fact]
    public async Task ReferenceModeOnlyReportsItsGroupAndPrunesOtherSizes()
    {
        byte[] bytes = Data(2 * 1024 * 1024, 77);
        string reference = FileAt("reference", bytes), copy = FileAt("copy", bytes);
        var result = await new DuplicateDetector().DetectAsync(Entries(reference, copy,
            FileAt("different", Data(bytes.Length, 78)), FileAt("unrelated1", [1]), FileAt("unrelated2", [1])), reference);
        Assert.Equal(2, Assert.Single(result.Groups).Files.Count); Assert.Equal(reference, result.ReferenceFile!.Entry.AbsolutePath);
        Assert.Equal(3, result.Stats.SizeCandidates); Assert.Equal(2, result.Stats.FullHashCount); Stats(result);
    }

    [Fact]
    public async Task MissingOrLockedFileDoesNotDiscardAccessibleDuplicateGroup()
    {
        string first = FileAt("first", [1, 2]), second = FileAt("copy", [1, 2]), missing = FileAt("missing", [1, 2]), locked = FileAt("locked", [1, 2]);
        var entries = Entries(first, second, missing, locked); File.Delete(missing);
        using var fileLock = new FileStream(locked, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        var result = await new DuplicateDetector().DetectAsync(entries);
        Assert.Equal(2, Assert.Single(result.Groups).Files.Count); Assert.Equal(2, result.Errors.Count); Stats(result);
    }

    [Fact]
    public async Task ChangedSinceEnumerationOrDuringHashIsRejected()
    {
        string first = FileAt("first", [1, 2]), second = FileAt("copy", [1, 2]), stale = FileAt("stale", [1, 2]);
        var entries = Entries(first, second, stale); File.SetLastWriteTimeUtc(stale, DateTime.UtcNow.AddMinutes(-1));
        var result = await new DuplicateDetector(hasher: new MutatingHasher(first)).DetectAsync(entries);
        Assert.Empty(result.Groups); Assert.Equal(2, result.Errors.Count);
        Assert.All(result.Errors, error => Assert.Equal("ChangedDuringScan", error.ErrorType));
    }

    [Fact]
    public async Task CancellationDuringHashClosesStreamAndReturnsNoPartialResult()
    {
        using var cancellation = new CancellationTokenSource();
        string first = FileAt("first", Data(1024 * 1024, 4)), second = FileAt("copy", Data(1024 * 1024, 4));
        var entries = Entries(first, second);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new DuplicateDetector(hasher: new CancellingHasher(cancellation)).DetectAsync(entries, cancellationToken: cancellation.Token));
        using var fileLock = new FileStream(first, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        Assert.True(fileLock.CanWrite);
    }

    [Fact]
    public async Task AlreadyCanceledScanStartsNoHashWork()
    {
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new DuplicateDetector(hasher: new ForbiddenHasher()).DetectAsync([], cancellationToken: cancellation.Token));
    }

    [Fact]
    public async Task EmptyInputAndSingleFileNeedNoHash()
    {
        var detector = new DuplicateDetector(hasher: new ForbiddenHasher());
        Assert.Empty((await detector.DetectAsync([])).Groups);
        var result = await detector.DetectAsync(Entries(FileAt("only", [4, 5])));
        Assert.Empty(result.Groups); Assert.Equal(0, result.Stats.BytesRead);
    }

    [Fact]
    public async Task ScanTimeCloudAttributesAreHonoredWithoutOpeningTheFile()
    {
        var entries = Entries(FileAt("first", [1, 2]), FileAt("second", [1, 2]));
        entries[0] = entries[0] with { Attributes = entries[0].Attributes | LocalFileContentPolicy.RecallOnOpen };
        var result = await new DuplicateDetector(hasher: new ForbiddenHasher()).DetectAsync(entries);
        Assert.Empty(result.Groups); Assert.Equal("CloudOnly", Assert.Single(result.Skipped).ErrorType);
    }

    [Fact]
    public async Task WorkerLimitIsTwoEvenWithManyCandidates()
    {
        var hasher = new ConcurrentHasher();
        var paths = Enumerable.Range(0, 12).Select(i => FileAt($"file{i}", [1, 2])).ToArray();
        var result = await new DuplicateDetector(hasher: hasher).DetectAsync(Entries(paths));
        Assert.Equal(12, Assert.Single(result.Groups).Files.Count); Assert.Equal(2, hasher.Maximum);
        Assert.Throws<ArgumentOutOfRangeException>(() => new DuplicateDetector(workerCount: 3));
    }

    private sealed class ForbiddenHasher : IFileContentHasher
    {
        public Task<ContentHash> QuickAsync(string path, long size, CancellationToken token) => throw new InvalidOperationException("Contents must not be opened.");
        public Task<ContentHash> FullAsync(string path, CancellationToken token) => throw new InvalidOperationException("Contents must not be opened.");
    }
    private sealed class MutatingHasher(string pathToChange) : IFileContentHasher
    {
        private readonly FileContentHasher inner = new();
        public Task<ContentHash> QuickAsync(string path, long size, CancellationToken token) => inner.QuickAsync(path, size, token);
        public async Task<ContentHash> FullAsync(string path, CancellationToken token)
        {
            var result = await inner.FullAsync(path, token);
            if (path == pathToChange) File.SetLastWriteTimeUtc(path, DateTime.UtcNow.AddMinutes(-2));
            return result;
        }
    }
    private sealed class CancellingHasher(CancellationTokenSource cancellation) : IFileContentHasher
    {
        private readonly FileContentHasher inner = new();
        public Task<ContentHash> QuickAsync(string path, long size, CancellationToken token) => inner.QuickAsync(path, size, token);
        public async Task<ContentHash> FullAsync(string path, CancellationToken token)
        {
            var result = await inner.FullAsync(path, token); cancellation.Cancel(); return result;
        }
    }
    private sealed class ConcurrentHasher : IFileContentHasher
    {
        private readonly FileContentHasher inner = new();
        private int active;
        public int Maximum;
        public Task<ContentHash> QuickAsync(string path, long size, CancellationToken token) => inner.QuickAsync(path, size, token);
        public async Task<ContentHash> FullAsync(string path, CancellationToken token)
        {
            int count = Interlocked.Increment(ref active);
            Interlocked.Exchange(ref Maximum, Math.Max(Maximum, count));
            try { await Task.Delay(10, token); return await inner.FullAsync(path, token); }
            finally { Interlocked.Decrement(ref active); }
        }
    }
}
