using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FileListToExcel.Core;
using Xunit;

namespace FileListToExcel.Core.Tests;

public sealed class CoreAcceptanceTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "FileListToExcel.Tests", Guid.NewGuid().ToString("N"));
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public CoreAcceptanceTests() => Directory.CreateDirectory(directory);
    public void Dispose() { try { Directory.Delete(directory, true); } catch (IOException) { } }
    private string FileAt(string name, string contents = "abc")
    {
        string path = Path.Combine(directory, name);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, contents);
        return path;
    }
    private List<ScanItem> Scan(ScanMode mode, params string[] paths) => new FileScanner().Enumerate(new(paths, mode)).ToList();
    private static List<FileEntry> Entries(IEnumerable<ScanItem> items) => items.Where(x => x.Entry is not null).Select(x => x.Entry!).ToList();

    [Fact]
    public void SelectedFilesIncludeExactlyTwentyDistinctSelections()
    {
        string[] files = Enumerable.Range(0, 100).Select(i => FileAt($"선택 {i}.{(i % 2 == 0 ? "txt" : "csv")}")).ToArray();
        var result = Entries(Scan(ScanMode.SelectedItems, [.. files.Take(20), files[0]]));
        Assert.Equal(20, result.Count);
        Assert.Equal(files.Take(20), result.Select(x => x.AbsolutePath));
        Assert.All(result, x => Assert.Equal(3, x.SizeBytes));
    }

    [Fact]
    public void DirectChildrenIncludeFoldersWithoutTheirContents()
    {
        FileAt("README"); FileAt("한글 하위/안쪽.txt");
        var result = Entries(Scan(ScanMode.DirectChildren, directory));
        Assert.Equal(2, result.Count);
        Assert.Contains(result, x => x.Name == "README" && x.Extension == "");
        Assert.Contains(result, x => x.Name == "한글 하위" && x.IsDirectory && x.SizeBytes is null);
        Assert.All(result, x => Assert.Equal(1, x.Depth));
    }

    [Fact]
    public void RecursiveSelectionPreservesIndependentRootsAndRelativeDepth()
    {
        FileAt("root A/child/first.txt"); FileAt("root B/child/second.txt");
        string rootA = Path.Combine(directory, "root A"), rootB = Path.Combine(directory, "root B");
        var result = Entries(Scan(ScanMode.Recursive, rootA, rootB));
        Assert.Equal(4, result.Count);
        var first = Assert.Single(result, x => x.Name == "first.txt");
        Assert.Equal(rootA, first.RootPath);
        Assert.Equal(Path.Combine("child", "first.txt"), first.RelativePath);
        Assert.Equal(2, first.Depth);
        Assert.Equal(rootB, Assert.Single(result, x => x.Name == "second.txt").RootPath);
    }

    [Fact]
    public void MissingAndInvalidSelectionsDoNotDiscardValidItems()
    {
        string existing = FileAt("exists.txt");
        var result = Scan(ScanMode.SelectedItems, Path.Combine(directory, "missing.txt"), "bad\0path", existing);
        Assert.Equal(2, result.Count(x => x.Error is not null));
        Assert.Equal(existing, Assert.Single(Entries(result)).AbsolutePath);
    }

    [Fact]
    public void EmptyFolderHasNoInventedEntries()
    {
        Assert.Empty(Scan(ScanMode.DirectChildren, directory));
        Assert.Empty(Scan(ScanMode.Recursive, directory));
    }

    [Fact]
    public void HiddenAndReadOnlyFilesAreNotFilteredOut()
    {
        string item = FileAt("숨김.txt");
        var previous = File.GetAttributes(item);
        try
        {
            File.SetAttributes(item, previous | FileAttributes.Hidden | FileAttributes.ReadOnly);
            var entry = Assert.Single(Entries(Scan(ScanMode.DirectChildren, directory)));
            Assert.Equal("숨김.txt", entry.Name);
            Assert.True(entry.Attributes.HasFlag(FileAttributes.ReadOnly));
        }
        finally { File.SetAttributes(item, previous); }
    }

    [Fact]
    public void CancellationStopsScanAndDoesNotPublishIncompleteWorkbook()
    {
        FileAt("exists.txt");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => ScanCanceled());
        void ScanCanceled() => new FileScanner().Enumerate(new([directory], ScanMode.Recursive), cancellationToken: cancellation.Token).ToList();
        string destination = Path.Combine(directory, "cancelled.xlsx");
        using var during = new CancellationTokenSource();
        IEnumerable<ScanItem> Midstream()
        {
            yield return new(Entry("ok.txt"));
            during.Cancel();
            yield return new(Entry("later.txt"));
        }
        Assert.Throws<OperationCanceledException>(() => new WorkbookWriter().Write(destination, Midstream(), during.Token));
        Assert.False(File.Exists(destination));
        Assert.Empty(Directory.GetFiles(directory, "*.partial"));
    }

    [Fact]
    public void WorkbookIsSchemaValidAndUsesTypedCellsTablesFiltersFreezeAndSafeLinks()
    {
        string source = FileAt("=HYPERLINK & 한글 # _x0020_.txt");
        string destination = Path.Combine(directory, "out.xlsx");
        var items = Scan(ScanMode.SelectedItems, source);
        items.Add(new(Error: new(source, "AccessDenied", "읽기 오류 < & >")));
        ExportResult result = new WorkbookWriter().Write(destination, items);
        Assert.Equal(1, result.EntryCount); Assert.Equal(1, result.ErrorCount);
        using (var document = SpreadsheetDocument.Open(destination, false))
        {
            var errors = new OpenXmlValidator().Validate(document).Select(x => $"{x.Path?.XPath}: {x.Description}").ToArray();
            Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors));
        }
        using var zip = ZipFile.OpenRead(destination);
        var sheet = Read(zip, "xl/worksheets/sheet1.xml");
        Assert.Empty(sheet.Descendants(Main + "f"));
        Assert.Equal("inlineStr", (string?)Cell(sheet, "A2").Attribute("t"));
        Assert.Contains("_x005F_x0020_", Cell(sheet, "A2").Value);
        Assert.Equal("3", Cell(sheet, "C2").Element(Main + "v")?.Value);
        Assert.Null(Cell(sheet, "C2").Attribute("t"));
        Assert.Equal("2", (string?)Cell(sheet, "E2").Attribute("s"));
        Assert.True(double.Parse(Cell(sheet, "E2").Value, System.Globalization.CultureInfo.InvariantCulture) > 40000);
        Assert.Equal("frozen", (string?)sheet.Descendants(Main + "pane").Single().Attribute("state"));
        Assert.Single(sheet.Descendants(Main + "hyperlink"));
        var table = Read(zip, "xl/tables/table1.xml");
        Assert.NotNull(table.Descendants(Main + "autoFilter").SingleOrDefault());
        Assert.Equal(14, table.Descendants(Main + "tableColumn").Count());
        string rels = Read(zip, "xl/worksheets/_rels/sheet1.xml.rels").ToString();
        Assert.Contains("%23", rels); Assert.Contains("TargetMode=\"External\"", rels);
        Assert.Contains("Errors", Read(zip, "xl/workbook.xml").ToString());
    }

    [Fact]
    public void EmptyWorkbookRemainsValidAndFilterable()
    {
        string destination = Path.Combine(directory, "empty.xlsx");
        Assert.Equal(0, new WorkbookWriter().Write(destination, []).EntryCount);
        using var document = SpreadsheetDocument.Open(destination, false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
    }

    [Fact]
    public void HyperlinkLimitSplitsSheetsWithoutDroppingRows()
    {
        string destination = Path.Combine(directory, "large.xlsx");
        var row = Entry("sample.txt");
        var items = Enumerable.Range(0, WorkbookWriter.RowsPerSheet + 1).Select(_ => new ScanItem(row));
        var result = new WorkbookWriter().Write(destination, items);
        Assert.Equal(65001, result.EntryCount); Assert.Equal(2, result.FileSheetCount);
        using var zip = ZipFile.OpenRead(destination);
        Assert.Equal(65000, CountElements(zip, "xl/worksheets/sheet1.xml", "hyperlink"));
        Assert.Equal(1, CountElements(zip, "xl/worksheets/sheet2.xml", "hyperlink"));
        Assert.Equal("A1:N65001", (string?)Read(zip, "xl/tables/table1.xml").Root?.Attribute("ref"));
        Assert.Equal("A1:N2", (string?)Read(zip, "xl/tables/table2.xml").Root?.Attribute("ref"));
    }

    [Fact]
    public void ExistingWorkbookIsNeverOverwritten()
    {
        string destination = FileAt("existing.xlsx", "preserved");
        Assert.Throws<IOException>(() => new WorkbookWriter().Write(destination, [new(Entry("sample.txt"))]));
        Assert.Equal("preserved", File.ReadAllText(destination));
        Assert.Empty(Directory.GetFiles(directory, "*.partial"));
    }

    [Fact]
    public void TenThousandRealFilesReportProgressAndKeepMetadataAccurate()
    {
        for (int i = 0; i < 10000; i++) FileAt($"bulk/{i:D5}.txt");
        var reports = new List<ScanProgress>();
        var entries = new FileScanner().Enumerate(new([Path.Combine(directory, "bulk")], ScanMode.Recursive), new SynchronousProgress(reports)).ToList();
        Assert.Equal(10000, entries.Count); Assert.All(entries, x => Assert.Null(x.Error));
        Assert.True(reports.Count > 100); Assert.Equal(10000, reports[^1].ItemCount);
    }

    [Fact]
    public void OverlongHyperlinkIsReportedWhileThePathIsPreserved()
    {
        var entry = Entry("long.txt") with { AbsolutePath = Path.Combine(directory, new string('a', 2200), "long.txt") };
        string destination = Path.Combine(directory, "long.xlsx");
        var result = new WorkbookWriter().Write(destination, [new(entry)]);
        Assert.Equal(1, result.EntryCount); Assert.Equal(1, result.ErrorCount);
        using var zip = ZipFile.OpenRead(destination);
        var sheet = Read(zip, "xl/worksheets/sheet1.xml");
        Assert.Equal(entry.AbsolutePath, Cell(sheet, "I2").Value); Assert.Empty(sheet.Descendants(Main + "hyperlink"));
    }

    [Fact]
    public void InvalidXmlCharactersAndFormulaLookingNamesAreLiteralText()
    {
        var entry = Entry("+cmd _x000A_\u0001 😀") with { RelativePath = "@SUM(A1)", ParentName = "-1+2", RootPath = "=1+1" };
        string destination = Path.Combine(directory, "literal.xlsx");
        new WorkbookWriter().Write(destination, [new(entry)]);
        using var document = SpreadsheetDocument.Open(destination, false);
        Assert.Empty(new OpenXmlValidator().Validate(document));
        using var zip = ZipFile.OpenRead(destination);
        var sheet = Read(zip, "xl/worksheets/sheet1.xml");
        Assert.Empty(sheet.Descendants(Main + "f"));
        Assert.Contains("_x005F_x000A_", Cell(sheet, "A2").Value);
        Assert.Contains("_x0001_", Cell(sheet, "A2").Value);
    }


    [Fact]
    public void DirectoryLinkCycleIsListedButNeverTraversed()
    {
        string target = Path.Combine(directory, "tree");
        Directory.CreateDirectory(target);
        File.WriteAllText(Path.Combine(target, "safe.txt"), "safe");
        string link = Path.Combine(target, "cycle");
        if (OperatingSystem.IsWindows())
        {
            var start = new System.Diagnostics.ProcessStartInfo("cmd.exe") { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
            foreach (string arg in new[] { "/c", "mklink", "/J", link, target }) start.ArgumentList.Add(arg);
            using var process = System.Diagnostics.Process.Start(start)!;
            process.WaitForExit();
            Assert.True(process.ExitCode == 0, process.StandardError.ReadToEnd());
        }
        else Directory.CreateSymbolicLink(link, target);
        try
        {
            var result = Scan(ScanMode.Recursive, target);
            Assert.Equal(2, Entries(result).Count);
            Assert.Contains(result, x => x.Entry?.Name == "cycle");
            Assert.Contains(result, x => x.Error?.ErrorType == "LinkNotTraversed");
        }
        finally { Directory.Delete(link); }
    }

    [Fact]
    public void FileContentLockedAgainstReadingStillAllowsMetadataScan()
    {
        string path = FileAt("locked.txt", "content must not be read");
        using var locked = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        var result = Scan(ScanMode.SelectedItems, path);
        Assert.Null(Assert.Single(result).Error);
        Assert.Equal(24, result[0].Entry!.SizeBytes);
    }

    [Fact]
    public void DeletedItemDuringEnumerationIsIsolated()
    {
        string first = FileAt("first.txt"), deleted = FileAt("deleted.txt"), last = FileAt("last.txt");
        using var scan = new FileScanner().Enumerate(new([first, deleted, last], ScanMode.SelectedItems)).GetEnumerator();
        Assert.True(scan.MoveNext()); Assert.Equal(first, scan.Current.Entry!.AbsolutePath);
        File.Delete(deleted);
        Assert.True(scan.MoveNext()); Assert.NotNull(scan.Current.Error);
        Assert.True(scan.MoveNext()); Assert.Equal(last, scan.Current.Entry!.AbsolutePath);
        Assert.False(scan.MoveNext());
    }

    [Fact]
    public void AccessDeniedSubdirectoryDoesNotDiscardAccessibleSibling()
    {
        if (!OperatingSystem.IsWindows()) return;
        string denied = Path.Combine(directory, "denied");
        Directory.CreateDirectory(denied);
        string visible = FileAt("visible.txt");
        var info = new DirectoryInfo(denied);
        var original = System.IO.FileSystemAclExtensions.GetAccessControl(info);
        var changed = System.IO.FileSystemAclExtensions.GetAccessControl(info);
        var identity = System.Security.Principal.WindowsIdentity.GetCurrent().User!;
        changed.AddAccessRule(new System.Security.AccessControl.FileSystemAccessRule(identity,
            System.Security.AccessControl.FileSystemRights.ListDirectory,
            System.Security.AccessControl.AccessControlType.Deny));
        try
        {
            System.IO.FileSystemAclExtensions.SetAccessControl(info, changed);
            var result = Scan(ScanMode.Recursive, directory);
            Assert.Contains(result, x => x.Entry?.AbsolutePath == visible);
            Assert.Contains(result, x => x.Error?.Path == denied);
        }
        finally
        {
            original.SetSecurityDescriptorBinaryForm(original.GetSecurityDescriptorBinaryForm(), System.Security.AccessControl.AccessControlSections.Access);
            System.IO.FileSystemAclExtensions.SetAccessControl(info, original);
        }
    }


    [Fact]
    public void LongNestedPathsAndUnicodeNamesRemainIntact()
    {
        string relative = string.Join(Path.DirectorySeparatorChar, Enumerable.Repeat(new string('가', 35), 9)) + Path.DirectorySeparatorChar + "공백 있는 파일 😀.txt";
        string path = FileAt(relative);
        var entries = Entries(Scan(ScanMode.Recursive, directory));
        var file = Assert.Single(entries, x => !x.IsDirectory);
        Assert.Equal(path, file.AbsolutePath); Assert.Equal(relative, file.RelativePath); Assert.Equal(10, file.Depth);
    }

    [Fact]
    public void Excel1900DateBoundaryUsesTheCorrectSerial()
    {
        var entry = Entry("old.txt") with { CreatedAt = new DateTime(1900, 1, 1), ModifiedAt = new DateTime(1900, 3, 1) };
        string destination = Path.Combine(directory, "dates.xlsx");
        new WorkbookWriter().Write(destination, [new(entry)]);
        using var zip = ZipFile.OpenRead(destination);
        var sheet = Read(zip, "xl/worksheets/sheet1.xml");
        Assert.Equal("61", Cell(sheet, "E2").Value); Assert.Equal("1", Cell(sheet, "F2").Value);
    }


    [Fact]
    public void OutputInsideSelectedFolderNeverListsItsOwnStagingWorkbook()
    {
        string source = FileAt("source.txt");
        string output = Path.Combine(directory, "output.xlsx");
        var result = new WorkbookWriter().Write(output, new FileScanner().Enumerate(new([directory], ScanMode.Recursive)));
        Assert.Equal(1, result.EntryCount);
        using var zip = ZipFile.OpenRead(output);
        var sheet = Read(zip, "xl/worksheets/sheet1.xml");
        Assert.Equal(source, Cell(sheet, "I2").Value);
    }


    [Fact]
    public void HyperlinksPreserveUnicodeAndEscapeReservedFilenameCharacters()
    {
        var entry = Entry("선택 자료 😀 # 100% question?.txt");
        string output = Path.Combine(directory, "unicode-links.xlsx");
        new WorkbookWriter().Write(output, [new(entry)]);
        using (var document = SpreadsheetDocument.Open(output, false))
            Assert.Empty(new OpenXmlValidator().Validate(document));
        using var zip = ZipFile.OpenRead(output);
        XNamespace relationship = "http://schemas.openxmlformats.org/package/2006/relationships";
        var links = Read(zip, "xl/worksheets/_rels/sheet1.xml.rels");
        string target = (string)links.Descendants(relationship + "Relationship").Single(x => (string?)x.Attribute("TargetMode") == "External").Attribute("Target")!;
        Assert.Contains("선택", target); Assert.Contains("자료", target); Assert.Contains("😀", target);
        Assert.Contains("%23", target); Assert.Contains("%25", target); Assert.Contains("%3F", target); Assert.Contains("%20", target);
        Assert.DoesNotContain("#", target); Assert.DoesNotContain("?", target); Assert.DoesNotContain(" ", target);
        Assert.Equal(entry.AbsolutePath, new Uri(target).LocalPath);
    }


    [Theory]
    [InlineData("한글 #10% 😀.txt")]
    [InlineData("😀 #10% 한글.txt")]
    [InlineData("한😀글%#😀.txt")]
    [InlineData("원본 %F0%9F%98%80 😀.txt")]
    public void ScannedKoreanHashPercentEmojiFilenameProducesValidHyperlink(string name)
    {
        string source = FileAt(name, "special hyperlink target");
        string output = Path.Combine(directory, "combined-unicode.xlsx");
        new WorkbookWriter().Write(output, new FileScanner().Enumerate(new([source], ScanMode.SelectedItems)));
        using (var document = SpreadsheetDocument.Open(output, false))
            Assert.Empty(new OpenXmlValidator().Validate(document));
        using var zip = ZipFile.OpenRead(output);
        XNamespace relationship = "http://schemas.openxmlformats.org/package/2006/relationships";
        string target = (string)Read(zip, "xl/worksheets/_rels/sheet1.xml.rels").Descendants(relationship + "Relationship").Single(x => (string?)x.Attribute("TargetMode") == "External").Attribute("Target")!;
        Assert.Contains("😀", target);
        Assert.Contains("%25", target);
        if (name.Contains('#')) Assert.Contains("%23", target);
        if (name.Contains("한글", StringComparison.Ordinal)) Assert.Contains("한글", target);
        Assert.Equal(source, new Uri(target).LocalPath);
    }

    private FileEntry Entry(string name) => new(name, "txt", false, 42, new DateTime(2025, 1, 2, 3, 4, 5), new DateTime(2025, 2, 3, 4, 5, 6), "root", Path.Combine(directory, name), name, 1, FileAttributes.Normal, directory);
    private static XDocument Read(ZipArchive zip, string name) { using var stream = zip.GetEntry(name)!.Open(); return XDocument.Load(stream); }
    private static XElement Cell(XDocument sheet, string reference) => sheet.Descendants(Main + "c").Single(x => (string?)x.Attribute("r") == reference);
    private static int CountElements(ZipArchive zip, string part, string localName)
    {
        using var stream = zip.GetEntry(part)!.Open(); using var reader = System.Xml.XmlReader.Create(stream);
        int count = 0; while (reader.Read()) if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.LocalName == localName) count++;
        return count;
    }
    private sealed class SynchronousProgress(List<ScanProgress> values) : IProgress<ScanProgress>
    {
        public void Report(ScanProgress value) => values.Add(value);
    }
}
