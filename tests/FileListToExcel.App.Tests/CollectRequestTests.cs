using System.Text.Json;
using FileListToExcel.App;
using FileListToExcel.Core;
using Xunit;

namespace FileListToExcel.App.Tests;

public sealed class CollectRequestTests
{
    [Fact]
    public void OwnedRequestIsConsumedOnceAndNoHeadlessCopyOptionExists()
    {
        var path = CreateRequest(out _);
        try
        {
            var options = CommandLine.Parse(["--collect-request", path]);
            Assert.NotNull(options.CollectRequest);
            Assert.False(options.NoOpen);
            Assert.Single(options.CollectRequest.Rows);
            Assert.False(File.Exists(path));
            Assert.ThrowsAny<Exception>(() => CommandLine.Parse(["--collect-request", path]));
        }
        finally { File.Delete(path); }
    }

    [Theory]
    [InlineData("--no-open")]
    [InlineData("--output")]
    [InlineData("--files")]
    public void CopyCannotSkipInteractiveDestination(string option)
    {
        var path = CreateRequest(out _);
        try
        {
            Assert.Throws<ArgumentException>(() => CommandLine.Parse(["--collect-request", path, option]));
            Assert.True(File.Exists(path));
        }
        finally { File.Delete(path); }
    }

    [Theory]
    [InlineData("version")]
    [InlineData("id")]
    [InlineData("table")]
    [InlineData("path")]
    [InlineData("name")]
    [InlineData("duplicate")]
    [InlineData("unknown")]
    public void MalformedRequestsRemainUntouched(string change)
    {
        var path = CreateRequest(out var json);
        try
        {
            json = change switch
            {
                "version" => json.Replace("\"version\":1", "\"version\":2"),
                "id" => json.Replace(Path.GetFileNameWithoutExtension(path), Guid.NewGuid().ToString("D")),
                "table" => json.Replace("FileListTable1", "OrdinaryTable"),
                "path" => json.Replace("\"displayPath\":\"C:\\\\fixtures\\\\sample.txt\"", "\"displayPath\":\"C:\\\\elsewhere.txt\""),
                "name" => json.Replace("\"displayName\":\"sample.txt\"", "\"displayName\":\"wrong.txt\""),
                "duplicate" => json.Replace("\"version\":1", "\"version\":1,\"version\":1"),
                _ => json.Replace("\"version\":1", "\"version\":1,\"execute\":\"no\"")
            };
            File.WriteAllText(path, json);
            Assert.ThrowsAny<Exception>(() => CollectRequestReader.Read(path));
            Assert.Equal(json, File.ReadAllText(path));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void ExternalGuidFileCannotBeConsumed()
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".json");
        File.WriteAllText(path, "{}");
        try
        {
            Assert.Throws<ArgumentException>(() => CollectRequestReader.Read(path));
            Assert.True(File.Exists(path));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void HardlinkedRequestCannotBeConsumed()
    {
        var path = CreateRequest(out _);
        var link = Path.Combine(CommandLine.RequestDirectory, Guid.NewGuid() + ".json");
        try
        {
            if (!CreateHardLink(link, path, IntPtr.Zero)) throw new IOException("hardlink fixture failed");
            Assert.Throws<IOException>(() => CollectRequestReader.Read(path));
            Assert.True(File.Exists(path));
            Assert.True(File.Exists(link));
        }
        finally { File.Delete(link); File.Delete(path); }
    }

    [Theory]
    [InlineData(@"C:\", @"C:\")]
    [InlineData(@"\\server\share\", "share")]
    public void DirectoryRootRowsAreAcceptedThenExcludedWithoutFilesystemAccess(string directoryRoot, string displayName)
    {
        Directory.CreateDirectory(CommandLine.RequestDirectory);
        var id = Guid.NewGuid(); var itemId = Guid.NewGuid().ToString("D");
        string record = JsonSerializer.Serialize(new { kind = "directory", itemId, absolutePath = directoryRoot,
            sizeBytes = (string?)null, modifiedUtcTicks = "639257184000000000" });
        string json = JsonSerializer.Serialize(new { mode = "collect-files", version = 1, requestId = id,
            workbookId = Guid.NewGuid(), tableName = "FileListTable1",
            rows = new[] { new { itemId, displayPath = directoryRoot, displayName, sourceRecord = record } } });
        string path = Path.Combine(CommandLine.RequestDirectory, id + ".json");
        File.WriteAllText(path, json);
        try
        {
            var request = CollectRequestReader.Read(path);
            Assert.False(File.Exists(path));
            var plan = new CopyPlanner().BuildPlan(request.Rows);
            Assert.Empty(plan.EligibleFiles);
            var excluded = Assert.Single(plan.Outcomes);
            Assert.Equal(CopyStatus.Excluded, excluded.Status);
            Assert.Equal("NotFile", excluded.Code);
            Assert.Equal(directoryRoot, excluded.SourcePath);
        }
        finally { File.Delete(path); }
    }

    private static string CreateRequest(out string json)
    {
        Directory.CreateDirectory(CommandLine.RequestDirectory);
        var id = Guid.NewGuid();
        var itemId = Guid.NewGuid().ToString("D");
        var record = JsonSerializer.Serialize(new { kind = "file", itemId, absolutePath = @"C:\fixtures\sample.txt", sizeBytes = "123", modifiedUtcTicks = "639257184000000000" });
        json = JsonSerializer.Serialize(new { mode = "collect-files", version = 1, requestId = id, workbookId = Guid.NewGuid(), tableName = "FileListTable1",
            rows = new[] { new { itemId, displayPath = @"C:\fixtures\sample.txt", displayName = "sample.txt", sourceRecord = record } } });
        var path = Path.Combine(CommandLine.RequestDirectory, id + ".json");
        File.WriteAllText(path, json);
        return path;
    }

    [System.Runtime.InteropServices.DllImport("kernel32.dll", EntryPoint = "CreateHardLinkW", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)]
    [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
    private static extern bool CreateHardLink(string link, string target, IntPtr security);
}
