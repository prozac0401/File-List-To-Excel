using FileListToExcel.App;
using FileListToExcel.Core;
using System.Text.Json;
using Xunit;

namespace FileListToExcel.App.Tests;

public sealed class CommandLineTests
{
    [Theory]
    [InlineData("--files", ScanMode.SelectedItems)]
    [InlineData("--folder", ScanMode.DirectChildren)]
    [InlineData("--recursive", ScanMode.Recursive)]
    public void PreservesSelectedPathsAndMode(string option, ScanMode expected)
    {
        var parsed = CommandLine.Parse([option, @"C:\한글 공백\one.txt", @"D:\two.txt", "--no-open"]);
        Assert.Equal(expected, parsed.Request!.Mode);
        Assert.Equal(2, parsed.Request.Paths.Count);
        Assert.Equal(@"C:\한글 공백\one.txt", parsed.Request.Paths[0]);
        Assert.True(parsed.NoOpen);
    }

    [Fact]
    public void LiteralOptionLookingFilenameDoesNotEnableHeadlessMode()
    {
        var parsed = CommandLine.Parse(["--files", "--", "--no-open"]);
        Assert.False(parsed.NoOpen);
        Assert.Equal(Path.GetFullPath("--no-open"), Assert.Single(parsed.Request!.Paths));
    }

    [Fact]
    public void RejectsInvalidArguments()
    {
        string[][] cases =
        [
            ["--files"],
            ["--folder", "C:/", "--recursive", "C:/"],
            ["--unknown"],
            ["--files", "C:/file", "--output"],
            ["--files", "C:/file", "--output", "C:/file.csv"]
        ];
        foreach (var args in cases) Assert.ThrowsAny<ArgumentException>(() => CommandLine.Parse(args));
    }
    [Fact]
    public void ExistingOutputIsNeverOverwritten()
    {
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try
        {
            File.WriteAllText(output, "keep");
            Assert.Throws<IOException>(() => CommandLine.Parse(["--files", output, "--output", output]));
            Assert.Equal("keep", File.ReadAllText(output));
        }
        finally { File.Delete(output); }
    }

    [Fact]
    public void RequestOutsideOwnedDirectoryIsNeverReadOrDeleted()
    {
        var source = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".json");
        try
        {
            File.WriteAllText(source, "private user data");
            Assert.Throws<ArgumentException>(() => CommandLine.Parse(["--request", source]));
            Assert.Equal("private user data", File.ReadAllText(source));
        }
        finally { File.Delete(source); }
    }

    [Fact]
    public void ValidShellRequestIsConsumedOnceAndPreservesUnicode()
    {
        Directory.CreateDirectory(CommandLine.RequestDirectory);
        var request = Path.Combine(CommandLine.RequestDirectory, Guid.NewGuid() + ".json");
        try
        {
            File.WriteAllText(request, JsonSerializer.Serialize(new { mode = "recursive", paths = new[] { @"C:\한글 폴더", @"D:\자료" } }));
            var parsed = CommandLine.Parse(["--request", request, "--no-open"]);
            Assert.Equal(ScanMode.Recursive, parsed.Request!.Mode);
            Assert.Equal(@"C:\한글 폴더", parsed.Request.Paths[0]);
            Assert.False(File.Exists(request));
        }
        finally { File.Delete(request); }
    }

    [Fact]
    public void InvalidShellRequestDoesNotDeleteAnything()
    {
        Directory.CreateDirectory(CommandLine.RequestDirectory);
        var request = Path.Combine(CommandLine.RequestDirectory, Guid.NewGuid() + ".json");
        try
        {
            File.WriteAllText(request, "{\"mode\":\"files\",\"paths\":[\"relative\"]}");
            Assert.Throws<ArgumentException>(() => CommandLine.Parse(["--request", request]));
            Assert.True(File.Exists(request));
        }
        finally { File.Delete(request); }
    }
}
