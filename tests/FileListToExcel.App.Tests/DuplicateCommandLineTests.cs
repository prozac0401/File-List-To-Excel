using System.Text.Json;
using FileListToExcel.App;
using FileListToExcel.Core;
using Xunit;

namespace FileListToExcel.App.Tests;

public sealed class DuplicateCommandLineTests
{
    [Theory]
    [InlineData("--matches", DuplicateMode.Matches)]
    [InlineData("--duplicates", DuplicateMode.Folders)]
    [InlineData("--duplicate-files", DuplicateMode.SelectedFiles)]
    public void NewModesPreserveSelectionAndDiagnosticFlags(string option, DuplicateMode expected)
    {
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        var parsed = CommandLine.Parse([option, Path.GetFullPath("C:/한글 공백/존재하지 않아도 됨.txt"), "--no-open", "--no-cache", "--output", output]);
        Assert.Null(parsed.Request); Assert.NotNull(parsed.DuplicateRequest);
        Assert.Equal(expected, parsed.DuplicateRequest.Mode);
        Assert.Equal(Path.GetFullPath("C:/한글 공백/존재하지 않아도 됨.txt"), Assert.Single(parsed.DuplicateRequest.Paths));
        Assert.True(parsed.NoOpen); Assert.True(parsed.NoCache); Assert.Equal(output, parsed.OutputPath);
    }

    [Theory]
    [InlineData("--duplicates", DuplicateMode.Folders)]
    [InlineData("--duplicate-files", DuplicateMode.SelectedFiles)]
    public void MultiplePathsAreOneDuplicateScope(string option, DuplicateMode expected)
    {
        var parsed = CommandLine.Parse([option, Path.GetFullPath("C:/자료 2025"), Path.GetFullPath("D:/백업 😀")]);
        Assert.Equal(expected, parsed.DuplicateRequest!.Mode);
        Assert.Equal(new[] { Path.GetFullPath("C:/자료 2025"), Path.GetFullPath("D:/백업 😀") }, parsed.DuplicateRequest.Paths);
        Assert.False(parsed.NoOpen); Assert.False(parsed.NoCache); Assert.Null(parsed.Request);
    }

    [Fact]
    public void MatchesRequiresExactlyOnePathAndOtherModesRequireAtLeastOne()
    {
        foreach (string flag in new[] { "--matches", "--duplicates", "--duplicate-files" })
            Assert.Throws<ArgumentException>(() => CommandLine.Parse([flag]));
        Assert.Throws<ArgumentException>(() => CommandLine.Parse(["--matches", "C:/first", "C:/second"]));
    }

    [Fact]
    public void AllListingAndDuplicateModesAreMutuallyExclusiveIncludingRepeatedModes()
    {
        string[] modes = ["--files", "--folder", "--recursive", "--matches", "--duplicates", "--duplicate-files"];
        foreach (string first in modes)
        foreach (string second in modes)
            Assert.Throws<ArgumentException>(() => CommandLine.Parse([first, "C:/selected", second]));
    }

    [Fact]
    public void LiteralOptionLookingNameRemainsASelectedFile()
    {
        var parsed = CommandLine.Parse(["--duplicate-files", "--", "--matches", "--no-cache", "--no-open"]);
        Assert.Equal(DuplicateMode.SelectedFiles, parsed.DuplicateRequest!.Mode);
        Assert.Equal(new[] { "--matches", "--no-cache", "--no-open" }.Select(Path.GetFullPath), parsed.DuplicateRequest.Paths);
        Assert.False(parsed.NoOpen); Assert.False(parsed.NoCache);
    }

    [Theory]
    [InlineData("matches", DuplicateMode.Matches)]
    [InlineData("duplicates", DuplicateMode.Folders)]
    [InlineData("duplicate-files", DuplicateMode.SelectedFiles)]
    public void ValidDuplicateShellRequestIsConsumedExactlyOnce(string mode, DuplicateMode expected)
    {
        string path = WriteRequest(mode, [Path.GetFullPath("C:/한글 # 😀/선택.txt")]);
        try
        {
            var parsed = CommandLine.Parse(["--request", path, "--no-open"]);
            Assert.Null(parsed.Request); Assert.Equal(expected, parsed.DuplicateRequest!.Mode);
            Assert.Equal(Path.GetFullPath("C:/한글 # 😀/선택.txt"), Assert.Single(parsed.DuplicateRequest.Paths));
            Assert.False(File.Exists(path));
            Assert.ThrowsAny<IOException>(() => CommandLine.Parse(["--request", path]));
        }
        finally { File.Delete(path); }
    }

    [Theory]
    [InlineData("matches")]
    [InlineData("duplicates")]
    [InlineData("duplicate-files")]
    public void InvalidSelectionPayloadIsPreserved(string mode)
    {
        foreach (string[] paths in new[] { Array.Empty<string>(), new[] { "relative" }, new[] { " " } })
        {
            string path = WriteRequest(mode, paths);
            try
            {
                string original = File.ReadAllText(path);
                Assert.Throws<ArgumentException>(() => CommandLine.Parse(["--request", path]));
                Assert.Equal(original, File.ReadAllText(path));
            }
            finally { File.Delete(path); }
        }
    }

    [Fact]
    public void InvalidMatchesCountAndUnknownRequestModeAreNotConsumed()
    {
        foreach (var payload in new[] { (Mode: "matches", Paths: new[] { Path.GetFullPath("C:/first"), Path.GetFullPath("C:/second") }), (Mode: "unknown", Paths: new[] { Path.GetFullPath("C:/first") }) })
        {
            string path = WriteRequest(payload.Mode, payload.Paths);
            try
            {
                Assert.Throws<ArgumentException>(() => CommandLine.Parse(["--request", path]));
                Assert.True(File.Exists(path));
            }
            finally { File.Delete(path); }
        }
    }

    [Theory]
    [InlineData("--matches")]
    [InlineData("--duplicates")]
    [InlineData("--duplicate-files")]
    public void ShellRequestCannotBeCombinedWithExplicitModeOrPaths(string option)
    {
        string path = WriteRequest("duplicates", [Path.GetFullPath("C:/selected")]);
        try
        {
            Assert.Throws<ArgumentException>(() => CommandLine.Parse(["--request", path, option, Path.GetFullPath("C:/second")]));
            Assert.True(File.Exists(path));
            Assert.Throws<ArgumentException>(() => CommandLine.Parse(["--request", path, Path.GetFullPath("C:/second")]));
            Assert.True(File.Exists(path));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void DuplicateOutputCannotOverwriteExistingWorkbookOrUseOtherExtension()
    {
        string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try
        {
            File.WriteAllText(output, "user contents");
            Assert.Throws<IOException>(() => CommandLine.Parse(["--duplicates", Path.GetFullPath("C:/folder"), "--output", output]));
            Assert.Equal("user contents", File.ReadAllText(output));
            Assert.Throws<ArgumentException>(() => CommandLine.Parse(["--matches", Path.GetFullPath("C:/file"), "--output", output + ".csv"]));
        }
        finally { File.Delete(output); }
    }

    [Fact]
    public void MoreThanOneHundredThousandSelectionsIsRejected()
    {
        string path = WriteRequest("duplicate-files", Enumerable.Repeat(Path.GetFullPath("C:/file"), 100001).ToArray());
        try
        {
            Assert.Throws<ArgumentException>(() => CommandLine.Parse(["--request", path]));
            Assert.True(File.Exists(path));
        }
        finally { File.Delete(path); }
    }

    [Theory]
    [InlineData("--help")]
    [InlineData("--version")]
    public void InformationOptionsDoNotStartADuplicateScan(string flag)
    {
        var parsed = CommandLine.Parse(["--duplicates", flag]);
        Assert.Null(parsed.Request); Assert.Null(parsed.DuplicateRequest);
        Assert.True(parsed.Help || parsed.Version);
    }

    private static string WriteRequest(string mode, string[] paths)
    {
        Directory.CreateDirectory(CommandLine.RequestDirectory);
        string path = Path.Combine(CommandLine.RequestDirectory, Guid.NewGuid() + ".json");
        File.WriteAllText(path, JsonSerializer.Serialize(new { mode, paths }));
        return path;
    }
}
