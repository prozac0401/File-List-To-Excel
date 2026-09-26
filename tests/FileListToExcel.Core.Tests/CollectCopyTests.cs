using System.Diagnostics;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using FileListToExcel.Core;
using Xunit;

namespace FileListToExcel.Core.Tests;

public sealed class CollectCopyTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "FileListToExcel.Collect.Tests", Guid.NewGuid().ToString("N"));
    private readonly List<string> reports = [];
    public CollectCopyTests() => Directory.CreateDirectory(directory);
    public void Dispose()
    {
        foreach (string report in reports)
            if (report.StartsWith(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)) File.Delete(report);
        try { Directory.Delete(directory, true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
    }
    private string Source(string name, byte[]? bytes = null)
    {
        string path = Path.Combine(directory, name);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, bytes ?? Encoding.UTF8.GetBytes("읽기 전용 원본\n"));
        return path;
    }
    private static CollectRow Row(string path, string kind = "file")
    {
        string id = Guid.NewGuid().ToString("D");
        return new(id, path, JsonSerializer.Serialize(new { kind, absolutePath = path,
            sizeBytes = kind == "file" ? new FileInfo(path).Length.ToString() : null,
            modifiedUtcTicks = File.GetLastWriteTimeUtc(path).Ticks.ToString(), itemId = id }));
    }
    private async Task<CopyResult> Copy(CopyPlan plan, IProgress<CopyProgress>? progress = null, CancellationToken token = default)
    {
        string destination = Path.Combine(directory, "delivery"); Directory.CreateDirectory(destination);
        var result = await new CopyService().CopyAsync(plan, destination, progress, token);
        if (result.ReportPath is not null) reports.Add(result.ReportPath);
        Assert.True(result.ReportError is null, result.ReportError);
        return result;
    }

    [Theory]
    [InlineData(@"C:\safe\file.txt")]
    [InlineData(@"\\server\share\한글 file.txt")]
    [InlineData(@"C:\safe\=SUM(1).xlsx")]
    public void OrdinaryWindowsPathsAccepted(string path) => Assert.Equal(path, CollectPathPolicy.Validate(path));

    [Theory]
    [InlineData("relative.txt")]
    [InlineData(@"C:relative.txt")]
    [InlineData(@"\rooted.txt")]
    [InlineData(@"\\?\C:\file.txt")]
    [InlineData(@"\\.\PhysicalDrive0")]
    [InlineData(@"C:\a.txt:Zone.Identifier")]
    [InlineData(@"C:\a\..\b.txt")]
    [InlineData(@"C:\a\CON.txt")]
    [InlineData(@"C:\a\file. ")]
    [InlineData(@"C:\a\file.")]
    [InlineData(@"C:\a\\file")]
    [InlineData("https://server/file")]
    [InlineData(@"\\server\share")]
    public void AmbiguousDeviceAndStreamPathsRejected(string path) => Assert.Throws<ArgumentException>(() => CollectPathPolicy.Validate(path));

    [Fact]
    public async Task CollisionsDuplicatePathsAndIdenticalBytesRemainDistinctAndReportIsPrivate()
    {
        string a = Source(Path.Combine("a", "same.txt")), b = Source(Path.Combine("b", "same.txt"));
        byte[] before = File.ReadAllBytes(a); DateTime stamp = File.GetLastWriteTimeUtc(a);
        var plan = new CopyPlanner().BuildPlan([Row(a), Row(b), Row(a)]);
        Assert.Equal(2, plan.EligibleFiles.Count); Assert.Equal(1, plan.DuplicateCount);
        Assert.Equal(["same.txt", "same (2).txt"], plan.EligibleFiles.Select(x => x.DestinationName));
        var result = await Copy(plan);
        Assert.Equal(2, result.CopiedCount); Assert.Equal(0, result.FailedCount);
        Assert.Equal(before, File.ReadAllBytes(a)); Assert.Equal(stamp, File.GetLastWriteTimeUtc(a));
        Assert.Equal(2, Directory.GetFiles(result.OutputDirectory!).Length);
        Assert.False(result.ReportPath!.StartsWith(result.OutputDirectory!, StringComparison.OrdinalIgnoreCase));
        using var report = JsonDocument.Parse(File.ReadAllText(result.ReportPath));
        Assert.Equal(3, report.RootElement.GetProperty("outcomes").GetArrayLength());
        Assert.Equal(before, File.ReadAllBytes(Path.Combine(result.OutputDirectory!, "same (2).txt")));
    }

    [Fact]
    public async Task DotfileNameCollisionsArePreserved()
    {
        var plan = new CopyPlanner().BuildPlan([Row(Source(Path.Combine("a", ".env"))), Row(Source(Path.Combine("b", ".env")))]);
        var result = await Copy(plan);
        Assert.Equal(2, result.CopiedCount);
        Assert.True(File.Exists(Path.Combine(result.OutputDirectory!, ".env (2)")));
    }

    [Fact]
    public async Task EmptyUnicodeFormulaLikeAndExtensionlessNamesCopy()
    {
        string empty = Source("empty", []), formula = Source("=SUM(한글 공백).txt");
        var result = await Copy(new CopyPlanner().BuildPlan([Row(empty), Row(formula)]));
        Assert.Equal(2, result.CopiedCount);
        Assert.Equal(0, new FileInfo(Path.Combine(result.OutputDirectory!, "empty")).Length);
        Assert.True(File.Exists(Path.Combine(result.OutputDirectory!, "=SUM(한글 공백).txt")));
    }

    [Fact]
    public async Task ReadOnlyAndZoneIdentifierArePreserved()
    {
        string source = Source("downloaded.txt");
        const string zone = "[ZoneTransfer]\r\nZoneId=3\r\n";
        File.WriteAllText(source + ":Zone.Identifier", zone);
        var attributes = File.GetAttributes(source);
        File.SetAttributes(source, attributes | FileAttributes.ReadOnly);
        CopyResult? result = null;
        try
        {
            result = await Copy(new CopyPlanner().BuildPlan([Row(source)]));
            Assert.True(result.CopiedCount == 1, JsonSerializer.Serialize(result.Outcomes));
            string copy = Path.Combine(result.OutputDirectory!, "downloaded.txt");
            Assert.Equal(zone, File.ReadAllText(copy + ":Zone.Identifier"));
            Assert.True(File.GetAttributes(copy).HasFlag(FileAttributes.ReadOnly));
            Assert.True(File.GetAttributes(source).HasFlag(FileAttributes.ReadOnly));
        }
        finally
        {
            File.SetAttributes(source, attributes);
            if (result?.CopiedCount == 1) File.SetAttributes(Path.Combine(result.OutputDirectory!, "downloaded.txt"), attributes);
        }
    }

    [Fact]
    public async Task LongPathIsCopiedToFlatNewFolder()
    {
        string relative = Path.Combine(Enumerable.Repeat(new string('가', 48), 6).Append("원본 😀.bin").ToArray());
        string source = Source(relative);
        Assert.True(source.Length > 260);
        var result = await Copy(new CopyPlanner().BuildPlan([Row(source)]));
        Assert.Equal(1, result.CopiedCount);
        Assert.Equal(File.ReadAllBytes(source), File.ReadAllBytes(Path.Combine(result.OutputDirectory!, "원본 😀.bin")));
    }

    [Fact]
    public void MissingChangedFolderOfflineAndMismatchedRowsAreExcluded()
    {
        string missing = Source("missing"), changed = Source("changed"), offline = Source("offline"), valid = Source("valid");
        var missingRow = Row(missing); File.Delete(missing);
        var changedRow = Row(changed); File.AppendAllText(changed, "changed");
        var offlineRow = Row(offline); var attributes = File.GetAttributes(offline);
        File.SetAttributes(offline, attributes | FileAttributes.Offline);
        try
        {
            var plan = new CopyPlanner().BuildPlan([missingRow, changedRow, Row(directory, "directory"), offlineRow, Row(valid) with { DisplayPath = changed }]);
            Assert.Empty(plan.EligibleFiles); Assert.Equal(5, plan.ExcludedCount);
            Assert.Contains(plan.Outcomes, x => x.Code == "ChangedSinceList");
            Assert.Contains(plan.Outcomes, x => x.Code == "CloudOnly");
            Assert.Contains(plan.Outcomes, x => x.Code == "PathMismatch");
        }
        finally { File.SetAttributes(offline, attributes); }
    }

    [Fact]
    public void ExistingWriterPreventsContentOpen()
    {
        string source = Source("writer.bin"); var row = Row(source);
        using var writer = new FileStream(source, FileMode.Open, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete);
        var plan = new CopyPlanner().BuildPlan([row]);
        Assert.Empty(plan.EligibleFiles); Assert.Equal("InUse", Assert.Single(plan.Outcomes).Code);
    }

    [Fact]
    public void DuplicateOrCorruptIdentityDoesNotBecomeAFileOperation()
    {
        string source = Source("valid"); var row = Row(source);
        Assert.Throws<ArgumentException>(() => new CopyPlanner().BuildPlan([row, row]));
        var plan = new CopyPlanner().BuildPlan([row with { ItemId = Guid.NewGuid().ToString("D") }]);
        Assert.Empty(plan.EligibleFiles); Assert.Equal("ItemIdMismatch", Assert.Single(plan.Outcomes).Code);
        Assert.Empty(new CopyPlanner().BuildPlan([row with { SourceRecord = "{broken}" }]).EligibleFiles);
    }

    [Fact]
    public async Task ChangeAndReplacementAfterPlanningAreExcluded()
    {
        string changed = Source("changed"), replaced = Source("replaced");
        var plan = new CopyPlanner().BuildPlan([Row(changed), Row(replaced)]);
        var modified = File.GetLastWriteTimeUtc(replaced); byte[] bytes = File.ReadAllBytes(replaced);
        File.AppendAllText(changed, "new data");
        string old = replaced + ".old"; File.Move(replaced, old); File.WriteAllBytes(replaced, bytes); File.SetLastWriteTimeUtc(replaced, modified);
        var result = await Copy(plan);
        Assert.Equal(0, result.CopiedCount);
        Assert.Contains(result.Outcomes, x => x.Code == "ChangedSinceList");
        Assert.Contains(result.Outcomes, x => x.Code == "SourceReplaced");
        Assert.Empty(Directory.GetFiles(result.OutputDirectory!));
    }

    [Fact]
    public async Task CancellationKeepsCompletedCopiesAndRemovesOnlyCurrentCopy()
    {
        string small = Source("small.txt"), large = Source("large.bin", []);
        using (var stream = new FileStream(large, FileMode.Open, FileAccess.Write)) stream.SetLength(128L * 1024 * 1024);
        DateTime modified = File.GetLastWriteTimeUtc(large);
        var plan = new CopyPlanner().BuildPlan([Row(small), Row(large)]);
        using var cancellation = new CancellationTokenSource();
        var progress = new ImmediateProgress(value => { if (value.FileName == "large.bin" && value.TransferredBytes > plan.EligibleFiles[0].Source.SizeBytes) cancellation.Cancel(); });
        var result = await Copy(plan, progress, cancellation.Token);
        Assert.True(result.Cancelled); Assert.Equal(1, result.CopiedCount); Assert.Equal(1, result.CancelledCount);
        Assert.True(File.Exists(Path.Combine(result.OutputDirectory!, "small.txt")));
        Assert.False(File.Exists(Path.Combine(result.OutputDirectory!, "large.bin")));
        Assert.Equal(128L * 1024 * 1024, new FileInfo(large).Length); Assert.Equal(modified, File.GetLastWriteTimeUtc(large));
        using var released = new FileStream(large, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
    }

    [Fact]
    public async Task CopyCallbackCannotModifyOrReplaceLockedSource()
    {
        string source = Source("locked.bin", new byte[4 * 1024 * 1024]); bool writerBlocked = false, renameBlocked = false;
        var result = await Copy(new CopyPlanner().BuildPlan([Row(source)]), new ImmediateProgress(value =>
        {
            if (value.CompletedCount > 0) return;
            try { using var writer = new FileStream(source, FileMode.Open, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete); }
            catch (IOException) { writerBlocked = true; }
            try { File.Move(source, source + ".replaced"); } catch (IOException) { renameBlocked = true; }
        }));
        Assert.Equal(1, result.CopiedCount); Assert.True(writerBlocked); Assert.True(renameBlocked);
    }

    [Fact]
    public async Task DestinationCancellationAndInvalidDestinationProducePerRowReports()
    {
        var plan = new CopyPlanner().BuildPlan([Row(Source("one")), Row(Source("two"))]);
        var cancelled = await new CopyService().RecordCancelledAsync(plan);
        if (cancelled.ReportPath is not null) reports.Add(cancelled.ReportPath);
        Assert.Null(cancelled.OutputDirectory); Assert.Equal(2, cancelled.CancelledCount); Assert.NotNull(cancelled.ReportPath);
        var invalid = await new CopyService().CopyAsync(plan, Path.Combine(directory, "does-not-exist"));
        if (invalid.ReportPath is not null) reports.Add(invalid.ReportPath);
        Assert.Equal(2, invalid.FailedCount); Assert.Null(invalid.OutputDirectory); Assert.NotNull(invalid.ReportPath);
    }

    [Fact]
    public async Task JunctionAncestorIsRejectedForSourceAndDestination()
    {
        string target = Path.Combine(directory, "target"); string source = Source(Path.Combine("target", "source.bin"));
        string junction = Path.Combine(directory, "junction");
        var start = new ProcessStartInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "System32", "cmd.exe"))
            { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
        foreach (string arg in new[] { "/c", "mklink", "/J", junction, target }) start.ArgumentList.Add(arg);
        using var process = Process.Start(start)!; await process.WaitForExitAsync();
        Assert.True(process.ExitCode == 0, await process.StandardError.ReadToEndAsync());
        try
        {
            var blocked = new CopyPlanner().BuildPlan([Row(Path.Combine(junction, "source.bin"))]);
            Assert.Empty(blocked.EligibleFiles); Assert.Equal("ReparsePoint", Assert.Single(blocked.Outcomes).Code);
            var result = await new CopyService().CopyAsync(new CopyPlanner().BuildPlan([Row(source)]), junction);
            if (result.ReportPath is not null) reports.Add(result.ReportPath);
            Assert.Equal(1, result.FailedCount); Assert.Equal("ReparsePoint", Assert.Single(result.Outcomes).Code);
            Assert.Single(Directory.GetFiles(target));
        }
        finally { Directory.Delete(junction); }
    }

    [Fact]
    public async Task ACompetingDestinationFileIsNeitherOverwrittenNorDeleted()
    {
        string first = Source("first.bin"), second = Source("second.bin");
        bool injected = false;
        var result = await Copy(new CopyPlanner().BuildPlan([Row(first), Row(second)]), new ImmediateProgress(_ =>
        {
            if (injected) return;
            string output = Assert.Single(Directory.GetDirectories(Path.Combine(directory, "delivery")));
            File.WriteAllText(Path.Combine(output, "second.bin"), "another owner's output");
            injected = true;
        }));
        Assert.Equal(1, result.CopiedCount); Assert.Equal(1, result.FailedCount);
        Assert.Equal("DestinationExists", result.Outcomes[1].Code);
        Assert.Equal("another owner's output", File.ReadAllText(Path.Combine(result.OutputDirectory!, "second.bin")));
    }

    [Fact]
    public async Task ConcurrentOperationsUseDifferentNewFolders()
    {
        string source = Source("shared.bin"); var plan = new CopyPlanner().BuildPlan([Row(source)]);
        var results = await Task.WhenAll(Copy(plan), Copy(plan));
        Assert.All(results, result => Assert.Equal(1, result.CopiedCount));
        Assert.NotEqual(results[0].OutputDirectory, results[1].OutputDirectory);
    }

    [Fact]
    public async Task PlanningCancellationAndNoEligiblePlanAreRecordedWithoutDestination()
    {
        var row = Row(Source("one"));
        using var token = new CancellationTokenSource(); token.Cancel();
        var cancelledPlan = new CopyPlanner().BuildPlan([row], token.Token);
        Assert.Equal(CopyStatus.Cancelled, Assert.Single(cancelledPlan.Outcomes).Status);
        var recorded = await new CopyService().RecordPlanAsync(cancelledPlan);
        if (recorded.ReportPath is not null) reports.Add(recorded.ReportPath);
        Assert.True(recorded.Cancelled); Assert.Null(recorded.OutputDirectory); Assert.NotNull(recorded.ReportPath);
        var invalidPlan = new CopyPlanner().BuildPlan([row with { SourceRecord = "{}" }]);
        var excluded = await new CopyService().RecordPlanAsync(invalidPlan);
        if (excluded.ReportPath is not null) reports.Add(excluded.ReportPath);
        Assert.False(excluded.Cancelled); Assert.Equal(1, excluded.ExcludedCount); Assert.NotNull(excluded.ReportPath);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SourceAclDenyIsRecordedAndReadableSiblingStillCopies(bool denyAfterPlanning)
    {
        if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("ACL-deny acceptance requires Windows; this environment is not verified.");
        string denied = Source("acl-denied.txt"), readable = Source("readable.txt");
        byte[] before = SHA256.HashData(File.ReadAllBytes(denied));
        DateTime modified = File.GetLastWriteTimeUtc(denied);
        var rows = new[] { Row(denied), Row(readable) };
        CopyPlan? plan = denyAfterPlanning ? new CopyPlanner().BuildPlan(rows) : null;
        var info = new FileInfo(denied);
        var original = info.GetAccessControl(); var restricted = info.GetAccessControl();
        restricted.AddAccessRule(new FileSystemAccessRule(WindowsIdentity.GetCurrent().User!, FileSystemRights.ReadData, AccessControlType.Deny));
        try
        {
            info.SetAccessControl(restricted);
            // Verify the fixture really denies this token; elevated/bypass behavior cannot silently pass.
            Assert.Throws<UnauthorizedAccessException>(() => { using var probe = File.OpenRead(denied); });
            plan ??= new CopyPlanner().BuildPlan(rows);
            var result = await Copy(plan);
            Assert.Equal(1, result.CopiedCount);
            var deniedOutcome = Assert.Single(result.Outcomes, row => row.SourcePath == denied);
            Assert.Equal("AccessDenied", deniedOutcome.Code);
            Assert.Equal(denyAfterPlanning ? CopyStatus.Failed : CopyStatus.Excluded, deniedOutcome.Status);
            Assert.False(File.Exists(Path.Combine(result.OutputDirectory!, "acl-denied.txt")));
            Assert.True(File.Exists(Path.Combine(result.OutputDirectory!, "readable.txt")));
            Assert.NotNull(result.ReportPath);
        }
        finally
        {
            original.SetSecurityDescriptorBinaryForm(original.GetSecurityDescriptorBinaryForm(), AccessControlSections.Access);
            info.SetAccessControl(original);
        }
        Assert.Equal(before, SHA256.HashData(File.ReadAllBytes(denied)));
        Assert.Equal(modified, File.GetLastWriteTimeUtc(denied));
    }

    [Fact]
    public async Task DestinationAclDenyRecordsEveryFileAndLeavesSourcesUnchanged()
    {
        if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("ACL-deny acceptance requires Windows; this environment is not verified.");
        string first = Source("one.txt"), second = Source("two.txt");
        byte[] firstHash = SHA256.HashData(File.ReadAllBytes(first)), secondHash = SHA256.HashData(File.ReadAllBytes(second));
        DateTime firstStamp = File.GetLastWriteTimeUtc(first), secondStamp = File.GetLastWriteTimeUtc(second);
        var plan = new CopyPlanner().BuildPlan([Row(first), Row(second)]);
        var destination = Directory.CreateDirectory(Path.Combine(directory, "acl-denied-destination"));
        var original = destination.GetAccessControl(); var restricted = destination.GetAccessControl();
        restricted.AddAccessRule(new FileSystemAccessRule(WindowsIdentity.GetCurrent().User!, FileSystemRights.CreateDirectories, AccessControlType.Deny));
        try
        {
            destination.SetAccessControl(restricted);
            Assert.Throws<UnauthorizedAccessException>(() => Directory.CreateDirectory(Path.Combine(destination.FullName, "permission-probe")));
            var result = await new CopyService().CopyAsync(plan, destination.FullName);
            if (result.ReportPath is not null) reports.Add(result.ReportPath);
            Assert.Null(result.ReportError); Assert.NotNull(result.ReportPath);
            Assert.Null(result.OutputDirectory); Assert.Equal(2, result.FailedCount); Assert.Equal(0, result.CopiedCount);
            Assert.All(result.Outcomes, row => Assert.Equal("AccessDenied", row.Code));
            Assert.Empty(Directory.GetFileSystemEntries(destination.FullName));
        }
        finally
        {
            original.SetSecurityDescriptorBinaryForm(original.GetSecurityDescriptorBinaryForm(), AccessControlSections.Access);
            destination.SetAccessControl(original);
        }
        Assert.Equal(firstHash, SHA256.HashData(File.ReadAllBytes(first)));
        Assert.Equal(secondHash, SHA256.HashData(File.ReadAllBytes(second)));
        Assert.Equal(firstStamp, File.GetLastWriteTimeUtc(first)); Assert.Equal(secondStamp, File.GetLastWriteTimeUtc(second));
    }

    private sealed class ImmediateProgress(Action<CopyProgress> callback) : IProgress<CopyProgress>
    {
        public void Report(CopyProgress value) => callback(value);
    }
}
