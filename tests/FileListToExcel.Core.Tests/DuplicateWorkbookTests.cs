using System.Collections;
using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FileListToExcel.Core;
using Xunit;

namespace FileListToExcel.Core.Tests;

public sealed class DuplicateWorkbookTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "FileListToExcel.DuplicateWorkbookTests", Guid.NewGuid().ToString("N"));
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace PackageRel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly string Hash = new('A', 64);
    public DuplicateWorkbookTests() => Directory.CreateDirectory(directory);
    public void Dispose() { try { Directory.Delete(directory, true); } catch (IOException) { } }
    private DuplicateFile File(string name, long size = 4096) => new(new FileEntry(name, Path.GetExtension(name).TrimStart('.'), false,
        size, new DateTime(2025, 1, 2, 3, 4, 5), new DateTime(2025, 2, 3, 4, 5, 6), "root",
        Path.Combine(directory, name), name, 1, FileAttributes.Normal, directory), Hash, Hash);
    private static DuplicateResult Result(IReadOnlyList<DuplicateGroup> groups, DuplicateFile? reference = null,
        IReadOnlyList<ScanError>? errors = null, IReadOnlyList<ScanError>? skipped = null) => new(groups, reference,
            new(groups.Sum(group => group.Files.Count), groups.Sum(group => group.Files.Count), groups.Sum(group => group.Files.Count), 0, 0, 0, 0, TimeSpan.FromMilliseconds(12)),
            errors ?? [], skipped ?? []);
    private static XDocument Read(ZipArchive zip, string name) { using var stream = zip.GetEntry(name)!.Open(); return XDocument.Load(stream); }
    private static XElement Cell(XDocument sheet, string address) => sheet.Descendants(Main + "c").Single(cell => (string?)cell.Attribute("r") == address);
    private static string[] Sheets(ZipArchive zip) => Read(zip, "xl/workbook.xml").Descendants(Main + "sheet").Select(sheet => (string)sheet.Attribute("name")!).ToArray();
    private static void SchemaValid(string path)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        var errors = new OpenXmlValidator().Validate(document).Select(error => $"{error.Path?.XPath}: {error.Description}").ToArray();
        Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors));
    }

    [Fact]
    public async Task RealDuplicateResultKeepsTypedCellsSortedGroupsAndSafeUnicodeLinks()
    {
        string special = "한글 # 100% 😀.txt";
        string[] names = ["small1", "small2", special, "largeCopy"];
        foreach (string name in names) System.IO.File.WriteAllBytes(Path.Combine(directory, name), name.StartsWith("small", StringComparison.Ordinal) ? [1] : [1, 2, 3, 4]);
        var result = await new DuplicateScanService().ScanAsync(new([directory], DuplicateMode.Folders));
        string destination = Path.Combine(directory, "duplicates.xlsx");
        var exported = new WorkbookWriter().WriteDuplicates(destination, result);
        Assert.Equal(4, exported.EntryCount); Assert.Equal(1, exported.FileSheetCount);
        SchemaValid(destination);
        using var zip = ZipFile.OpenRead(destination);
        Assert.Equal(new[] { "Duplicates", "Summary", "Errors", "Skipped" }, Sheets(zip));
        var sheet = Read(zip, "xl/worksheets/sheet1.xml");
        Assert.Equal("1", Cell(sheet, "A2").Value); Assert.Equal("2", Cell(sheet, "B2").Value);
        Assert.Equal("4", Cell(sheet, "C2").Value); Assert.Equal("4", Cell(sheet, "F2").Value);
        Assert.Null(Cell(sheet, "C2").Attribute("t")); Assert.Null(Cell(sheet, "F2").Attribute("t"));
        Assert.Equal("2", (string?)Cell(sheet, "H2").Attribute("s"));
        Assert.True(double.Parse(Cell(sheet, "H2").Value, CultureInfo.InvariantCulture) > 40000);
        Assert.Equal(64, Cell(sheet, "K2").Value.Length);
        Assert.Equal("frozen", (string?)sheet.Descendants(Main + "pane").Single().Attribute("state"));
        Assert.Equal("A2", (string?)sheet.Descendants(Main + "pane").Single().Attribute("topLeftCell"));
        Assert.Equal(4, sheet.Descendants(Main + "hyperlink").Count());
        Assert.All(sheet.Descendants(Main + "hyperlink"), link => Assert.StartsWith("D", (string)link.Attribute("ref")!));
        var table = Read(zip, "xl/tables/table1.xml");
        Assert.Equal("A1:M5", (string?)table.Root?.Attribute("ref")); Assert.Single(table.Descendants(Main + "autoFilter"));
        Assert.Equal(13, table.Descendants(Main + "tableColumn").Count());
        var rels = Read(zip, "xl/worksheets/_rels/sheet1.xml.rels");
        var target = rels.Descendants(PackageRel + "Relationship").Select(rel => (string?)rel.Attribute("Target")).Single(value => value?.Contains("😀", StringComparison.Ordinal) == true)!;
        Assert.Contains("%23", target); Assert.Contains("%25", target); Assert.Contains("%20", target);
        Assert.Equal(Path.Combine(directory, special), new Uri(target).LocalPath);
        Assert.Empty(sheet.Descendants(Main + "f"));
        var summary = Read(zip, "xl/worksheets/sheet2.xml");
        Assert.Contains(summary.Descendants(Main + "c"), cell => cell.Attribute("t") is null && cell.Element(Main + "v")?.Value == "5");
    }

    [Fact]
    public void MatchesShowsReferenceMetadataAndExcludesReferenceFromRows()
    {
        DuplicateFile reference = File("기준 파일.txt"), match = File("same bytes other name.txt");
        string destination = Path.Combine(directory, "matches.xlsx");
        var exported = new WorkbookWriter().WriteDuplicates(destination, Result([new(1, [reference, match])], reference), matches: true);
        Assert.Equal(1, exported.EntryCount); SchemaValid(destination);
        using var zip = ZipFile.OpenRead(destination);
        Assert.Equal(new[] { "Matches", "Summary", "Errors", "Skipped" }, Sheets(zip));
        var sheet = Read(zip, "xl/worksheets/sheet1.xml");
        var metadata = sheet.Descendants(Main + "row").Where(row => (int)row.Attribute("r")! <= 3).Select(row => row.Value).ToArray();
        Assert.Contains(metadata, value => value.Contains(reference.Entry.AbsolutePath, StringComparison.Ordinal));
        Assert.Contains(metadata, value => value.Contains(Hash, StringComparison.Ordinal));
        Assert.Contains(metadata, value => value.Contains("4 KB", StringComparison.Ordinal));
        Assert.Equal(match.Entry.Name, Cell(sheet, "A6").Value);
        Assert.Equal("4096", Cell(sheet, "C6").Value); Assert.Null(Cell(sheet, "C6").Attribute("t"));
        Assert.Equal("2", (string?)Cell(sheet, "E6").Attribute("s")); Assert.Equal(Hash, Cell(sheet, "H6").Value);
        Assert.Equal(match.Entry.AbsolutePath, Cell(sheet, "G6").Value);
        Assert.DoesNotContain(sheet.Descendants(Main + "row").Where(row => (int)row.Attribute("r")! > 5), row => row.Value.Contains(reference.Entry.AbsolutePath, StringComparison.Ordinal));
        Assert.Equal("A6", (string?)sheet.Descendants(Main + "pane").Single().Attribute("topLeftCell"));
        Assert.Equal("5", (string?)sheet.Descendants(Main + "pane").Single().Attribute("ySplit"));
        Assert.Equal("A6", (string?)sheet.Descendants(Main + "hyperlink").Single().Attribute("ref"));
        Assert.Equal("A5:J6", (string?)Read(zip, "xl/tables/table1.xml").Root?.Attribute("ref"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EmptyResultsKeepValidFilterableSheetsWithoutInventedMatches(bool matches)
    {
        string destination = Path.Combine(directory, "empty.xlsx");
        var reference = matches ? File("reference") with { FullSha256 = null, QuickHash = null } : null;
        var exported = new WorkbookWriter().WriteDuplicates(destination, Result([], reference), matches);
        Assert.Equal(0, exported.EntryCount); SchemaValid(destination);
        using var zip = ZipFile.OpenRead(destination);
        Assert.Equal(4, Sheets(zip).Length); Assert.Contains("Errors", Sheets(zip)); Assert.Contains("Skipped", Sheets(zip));
        var sheet = Read(zip, "xl/worksheets/sheet1.xml"); Assert.Empty(sheet.Descendants(Main + "hyperlink"));
        Assert.Single(Read(zip, "xl/tables/table1.xml").Descendants(Main + "autoFilter"));
    }

    [Fact]
    public void ErrorsSkippedAndFormulaLookingNamesRemainLiteralAndSeparate()
    {
        DuplicateFile first = File("=HYPERLINK _x000A_한글.txt"), second = File("+cmd.txt");
        var result = Result([new(1, [first, second])], errors: [new(first.Entry.AbsolutePath, "AccessDenied", "<권한> & 오류")],
            skipped: [new(second.Entry.AbsolutePath, "CloudOnly", "건너뜀: Cloud-only 파일")]);
        string destination = Path.Combine(directory, "issues.xlsx");
        var exported = new WorkbookWriter().WriteDuplicates(destination, result);
        Assert.Equal(1, exported.ErrorCount); SchemaValid(destination);
        using var zip = ZipFile.OpenRead(destination);
        var main = Read(zip, "xl/worksheets/sheet1.xml");
        Assert.Empty(main.Descendants(Main + "f"));
        Assert.Contains(main.Descendants(Main + "c"), cell => cell.Value.Contains("_x005F_x000A_", StringComparison.Ordinal));
        var errors = Read(zip, "xl/worksheets/sheet3.xml"); var skipped = Read(zip, "xl/worksheets/sheet4.xml");
        Assert.Contains("AccessDenied", errors.ToString()); Assert.Contains("CloudOnly", skipped.ToString());
        Assert.DoesNotContain("CloudOnly", errors.ToString()); Assert.Empty(errors.Descendants(Main + "f"));
    }

    [Fact]
    public void UnrepresentableLinkKeepsPathAndAddsAnError()
    {
        var first = File("first") with { Entry = File("first").Entry with { AbsolutePath = Path.Combine(directory, new string('a', 2200), "first") } };
        var second = File("copy");
        string destination = Path.Combine(directory, "longlink.xlsx");
        var exported = new WorkbookWriter().WriteDuplicates(destination, Result([new(1, [first, second])]));
        Assert.Equal(2, exported.EntryCount); Assert.Equal(1, exported.ErrorCount); SchemaValid(destination);
        using var zip = ZipFile.OpenRead(destination);
        var sheet = Read(zip, "xl/worksheets/sheet1.xml");
        Assert.Contains(sheet.Descendants(Main + "c"), cell => cell.Value == first.Entry.AbsolutePath);
        Assert.Single(sheet.Descendants(Main + "hyperlink"));
        Assert.Contains("HyperlinkUnavailable", Read(zip, "xl/worksheets/sheet3.xml").ToString());
    }

    [Fact]
    public void LargeDuplicateGroupSplitsAtHyperlinkLimitWithoutLosingRows()
    {
        var files = Enumerable.Range(0, WorkbookWriter.RowsPerSheet + 1).Select(index => File($"file{index:D5}.txt", 1)).ToArray();
        string destination = Path.Combine(directory, "large.xlsx");
        var exported = new WorkbookWriter().WriteDuplicates(destination, Result([new(1, files)]));
        Assert.Equal(65001, exported.EntryCount); Assert.Equal(2, exported.FileSheetCount);
        using var zip = ZipFile.OpenRead(destination);
        Assert.Equal(new[] { "Duplicates", "Duplicates_2", "Summary", "Errors", "Skipped" }, Sheets(zip));
        Assert.Equal(65000, Count(zip, "xl/worksheets/sheet1.xml", "hyperlink"));
        Assert.Equal(1, Count(zip, "xl/worksheets/sheet2.xml", "hyperlink"));
        Assert.Equal("A1:M65001", (string?)Read(zip, "xl/tables/table1.xml").Root?.Attribute("ref"));
        Assert.Equal("A1:M2", (string?)Read(zip, "xl/tables/table2.xml").Root?.Attribute("ref"));
    }

    [Fact]
    public void CancellationAndExistingDestinationNeverPublishIncompleteOrReplaceUserWork()
    {
        string cancelled = Path.Combine(directory, "cancelled.xlsx");
        using var cancellation = new CancellationTokenSource();
        var files = new CancelingList<DuplicateFile>([File("first"), File("second"), File("third")], cancellation);
        var result = Result([new(1, files)]);
        Assert.ThrowsAny<OperationCanceledException>(() => new WorkbookWriter().WriteDuplicates(cancelled, result, cancellationToken: cancellation.Token));
        Assert.False(System.IO.File.Exists(cancelled)); Assert.Empty(Directory.GetFiles(directory, "*.partial"));
        string existing = Path.Combine(directory, "existing.xlsx"); System.IO.File.WriteAllText(existing, "user workbook preserved");
        Assert.Throws<IOException>(() => new WorkbookWriter().WriteDuplicates(existing, Result([])));
        Assert.Equal("user workbook preserved", System.IO.File.ReadAllText(existing)); Assert.Empty(Directory.GetFiles(directory, "*.partial"));
    }

    private static int Count(ZipArchive zip, string part, string name)
    {
        using var stream = zip.GetEntry(part)!.Open(); using var reader = System.Xml.XmlReader.Create(stream);
        int count = 0; while (reader.Read()) if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.LocalName == name) count++;
        return count;
    }
    private sealed class CancelingList<T>(IReadOnlyList<T> items, CancellationTokenSource cancellation) : IReadOnlyList<T>
    {
        public int Count => items.Count;
        public T this[int index] => items[index];
        public IEnumerator<T> GetEnumerator()
        {
            for (int index = 0; index < items.Count; index++)
            {
                if (index == 1) cancellation.Cancel();
                yield return items[index];
            }
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
