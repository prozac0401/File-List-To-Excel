using System.Diagnostics;
using System.Security.AccessControl;
using System.Security.Principal;
using FileListToExcel.Core;
using Xunit;

namespace FileListToExcel.Core.Tests;

public sealed class DuplicateScanTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "FileListToExcel.DuplicateScanTests", Guid.NewGuid().ToString("N"));
    public DuplicateScanTests() => Directory.CreateDirectory(directory);
    public void Dispose() => Directory.Delete(directory, true);
    private string FileAt(string name, string value = "same")
    {
        var path = Path.GetFullPath(Path.Combine(directory, name));
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, value);
        return path;
    }

    [Fact]
    public async Task MultipleAndOverlappingFoldersAreOneSearchWithoutDuplicateRows()
    {
        FileAt("A/original"); FileAt("B/복사본.txt"); FileAt("A/nested/other", "different");
        var result = await new DuplicateScanService().ScanAsync(new([Path.Combine(directory, "A"), Path.Combine(directory, "B"), directory], DuplicateMode.Folders));
        Assert.Equal(3, result.Stats.ScannedFiles);
        Assert.Equal(2, Assert.Single(result.Groups).Files.Count);
        Assert.Empty(result.Errors);
    }

    [Fact]
    public async Task ReferenceSearchUsesParentTreeAndSelectedFilesDoNotIncludeSiblings()
    {
        var reference = FileAt("scope/reference"); var copy = FileAt("scope/deep/copy"); FileAt("outside", "same");
        var result = await new DuplicateScanService().ScanAsync(new([reference], DuplicateMode.Matches));
        Assert.Equal(reference, result.ReferenceFile!.Entry.AbsolutePath);
        Assert.Equal(2, result.Stats.ScannedFiles);
        Assert.Contains(Assert.Single(result.Groups).Files, file => file.Entry.AbsolutePath == copy);
        var selected = await new DuplicateScanService().ScanAsync(new([reference, FileAt("only-selected", "else")], DuplicateMode.SelectedFiles));
        Assert.Equal(2, selected.Stats.ScannedFiles); Assert.Empty(selected.Groups);
    }

    [Fact]
    public async Task MissingReferenceCannotReportOtherDuplicateGroups()
    {
        FileAt("one"); FileAt("two");
        var result = await new DuplicateScanService().ScanAsync(new([Path.Combine(directory, "missing")], DuplicateMode.Matches));
        Assert.Empty(result.Groups); Assert.Null(result.ReferenceFile); Assert.Single(result.Errors);
    }

    [Fact]
    public async Task EmptyFoldersAndTenLevelsOfLongUnicodePathsWork()
    {
        Assert.Empty((await new DuplicateScanService().ScanAsync(new([directory], DuplicateMode.Folders))).Groups);
        var nested = string.Join(Path.DirectorySeparatorChar, Enumerable.Repeat(new string('가', 30), 10));
        var longPath = FileAt(Path.Combine(nested, "한글 😀 #10%")); FileAt("copy");
        var result = await new DuplicateScanService().ScanAsync(new([directory], DuplicateMode.Folders));
        Assert.Contains(Assert.Single(result.Groups).Files, file => file.Entry.AbsolutePath == longPath);
        Assert.Empty(result.Errors);
    }

    [Fact]
    public async Task JunctionCycleAndExplicitJunctionRootAreSkipped()
    {
        if (!OperatingSystem.IsWindows()) return;
        FileAt("one"); FileAt("two");
        var junction = Path.Combine(directory, "cycle");
        var start = new ProcessStartInfo("cmd.exe") { UseShellExecute = false, CreateNoWindow = true, RedirectStandardError = true, RedirectStandardOutput = true };
        foreach (var argument in new[] { "/c", "mklink", "/J", junction, directory }) start.ArgumentList.Add(argument);
        using var process = Process.Start(start)!;
        await process.WaitForExitAsync(); Assert.Equal(0, process.ExitCode);
        try
        {
            var result = await new DuplicateScanService().ScanAsync(new([directory], DuplicateMode.Folders));
            Assert.Equal(2, result.Stats.ScannedFiles); Assert.Single(result.Groups);
            Assert.Equal("ReparsePoint", Assert.Single(result.Skipped).ErrorType);
            var explicitRoot = await new DuplicateScanService().ScanAsync(new([junction], DuplicateMode.Folders));
            Assert.Empty(explicitRoot.Groups); Assert.Single(explicitRoot.Skipped);
        }
        finally { Directory.Delete(junction); }
    }

    [Fact]
    public async Task AccessDeniedFolderDoesNotDiscardReadableDuplicates()
    {
        if (!OperatingSystem.IsWindows()) return;
        FileAt("one"); FileAt("two"); FileAt("denied/hidden");
        var folder = new DirectoryInfo(Path.Combine(directory, "denied"));
        var original = folder.GetAccessControl(); var restricted = folder.GetAccessControl();
        restricted.AddAccessRule(new FileSystemAccessRule(WindowsIdentity.GetCurrent().User!, FileSystemRights.ListDirectory, AccessControlType.Deny));
        try
        {
            folder.SetAccessControl(restricted);
            var result = await new DuplicateScanService().ScanAsync(new([directory], DuplicateMode.Folders));
            Assert.Single(result.Groups); Assert.Contains(result.Errors, error => error.Path == folder.FullName);
        }
        finally
        {
            original.SetSecurityDescriptorBinaryForm(original.GetSecurityDescriptorBinaryForm(), AccessControlSections.Access);
            folder.SetAccessControl(original);
        }
    }

    [Fact]
    public async Task CancellationWhileFindingFilesStopsBeforeHashing()
    {
        for (int i = 0; i < 100; i++) FileAt($"{i}");
        using var cancellation = new CancellationTokenSource();
        var progress = new CallbackProgress(value => { if (value.Stage == DuplicateStage.FindingFiles && value.FileCount > 0) cancellation.Cancel(); });
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new DuplicateScanService().ScanAsync(new([directory], DuplicateMode.Folders), progress, cancellation.Token));
    }

    private sealed class CallbackProgress(Action<DuplicateProgress> callback) : IProgress<DuplicateProgress>
    {
        public void Report(DuplicateProgress value) => callback(value);
    }
}
