using System.Globalization;
using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FileListToExcel.Core;
using Xunit;

namespace FileListToExcel.Core.Tests;

public sealed class FileCollectContractTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "FileListToExcel.ContractTests", Guid.NewGuid().ToString("N"));
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Custom = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
    public FileCollectContractTests() => Directory.CreateDirectory(directory);
    public void Dispose() { try { Directory.Delete(directory, true); } catch (IOException) { } }
    private FileEntry Entry(string name = "한글 = 수식 아님.txt") => new(name, "txt", false, 9007199254740993,
        new DateTime(2026, 9, 24, 1, 2, 3, DateTimeKind.Utc), new DateTime(2026, 9, 24, 2, 3, 4, DateTimeKind.Utc).AddTicks(1234),
        "ContractTests", Path.Combine(directory, name), name, 0, FileAttributes.Normal, directory);
    private static XDocument Read(ZipArchive zip, string part) { using var stream = zip.GetEntry(part)!.Open(); return XDocument.Load(stream); }
    private static string Cell(XDocument sheet, string address) => sheet.Descendants(Main + "c").Single(cell => (string?)cell.Attribute("r") == address).Value;
    private static Dictionary<string, string> Properties(ZipArchive zip) => Read(zip, "docProps/custom.xml").Descendants(Custom + "property").ToDictionary(p => (string)p.Attribute("name")!, p => p.Value);
    private static void SchemaValid(string path)
    {
        using var workbook = SpreadsheetDocument.Open(path, false);
        Assert.Empty(new OpenXmlValidator().Validate(workbook));
    }

    [Fact]
    public void ListWorkbookRegistersOnlyItsOwnTableAndKeepsExactMetadataOnEachRow()
    {
        string path = Path.Combine(directory, "list.xlsx");
        FileEntry first = Entry(), second = Entry("second.txt");
        new WorkbookWriter().Write(path, [new(first), new(second), new(Error: new("bad", "Missing", "Unavailable"))]);
        SchemaValid(path);
        using var zip = ZipFile.OpenRead(path);
        var props = Properties(zip);
        Assert.True(FileCollectContract.IsSupportedWorkbook(props["FLT.Product"], props["FLT.ActionSchemaVersion"], props["FLT.WorkbookId"]));
        Assert.Equal("files", props["FLT.Table.FileListTable1"]);
        Assert.False(props.ContainsKey("FLT.Table.FileListTable2"));
        Assert.All(props.Values, value => Assert.InRange(value.Length, 1, 255));
        var table = Read(zip, "xl/tables/table1.xml");
        Assert.Equal("FileListTable1", (string?)table.Root?.Attribute("name"));
        Assert.Equal("A1:N3", (string?)table.Root?.Attribute("ref"));
        Assert.Equal(new[] { FileCollectContract.ItemIdColumn, FileCollectContract.SourceRecordColumn },
            table.Descendants(Main + "tableColumn").TakeLast(2).Select(c => (string)c.Attribute("name")!));
        var sheet = Read(zip, "xl/worksheets/sheet1.xml");
        Assert.Equal(new[] { "13", "14" }, sheet.Descendants(Main + "col").Where(c => (string?)c.Attribute("hidden") == "1").Select(c => (string)c.Attribute("min")!));
        var ids = new HashSet<string>();
        foreach (var (entry, row) in new[] { (first, 2), (second, 3) })
        {
            string id = Cell(sheet, $"M{row}"), json = Cell(sheet, $"N{row}");
            Assert.True(ids.Add(id));
            Assert.True(FileCollectContract.TryParseSourceRecord(json, out var source, out var error), error);
            Assert.Equal(id, source!.ItemId); Assert.Equal("file", source.Kind);
            Assert.Equal(Cell(sheet, $"I{row}"), source.AbsolutePath);
            Assert.Equal(entry.SizeBytes, source.SizeBytes); Assert.Equal(entry.ModifiedAt.Ticks, source.ModifiedUtcTicks);
            using var raw = JsonDocument.Parse(json);
            Assert.Equal(JsonValueKind.String, raw.RootElement.GetProperty("sizeBytes").ValueKind);
            Assert.Equal(JsonValueKind.String, raw.RootElement.GetProperty("modifiedUtcTicks").ValueKind);
        }
        Assert.Empty(sheet.Descendants(Main + "f"));
        Assert.DoesNotContain(Read(zip, "xl/tables/table2.xml").Descendants(Main + "tableColumn"), c => ((string?)c.Attribute("name"))?.StartsWith("__FLT_", StringComparison.Ordinal) == true);
        Assert.DoesNotContain(zip.Entries, e => e.FullName.Contains("vba", StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData(false, "duplicates", "L2", "M2", "J2", 13)]
    [InlineData(true, "matches", "I6", "J6", "G6", 10)]
    public void DuplicateAndMatchRowsShareTheContractButSummaryAndIssuesDoNot(bool matches, string role,
        string idCell, string sourceCell, string pathCell, int columns)
    {
        string path = Path.Combine(directory, "result.xlsx");
        var first = new DuplicateFile(Entry("a.txt"), null, new string('A', 64));
        var second = new DuplicateFile(Entry("b.txt"), null, new string('A', 64));
        var result = new DuplicateResult([new(1, [first, second])], matches ? first : null,
            new(2, 2, 0, 0, 2, 0, 0, TimeSpan.Zero), [], []);
        new WorkbookWriter().WriteDuplicates(path, result, matches);
        SchemaValid(path);
        using var zip = ZipFile.OpenRead(path);
        var props = Properties(zip);
        Assert.Equal(role, props["FLT.Table.FileListTable1"]);
        Assert.Single(props.Keys, key => key.StartsWith("FLT.Table.", StringComparison.Ordinal));
        Assert.Equal(columns, Read(zip, "xl/tables/table1.xml").Descendants(Main + "tableColumn").Count());
        var sheet = Read(zip, "xl/worksheets/sheet1.xml");
        Assert.True(FileCollectContract.TryParseSourceRecord(Cell(sheet, sourceCell), out var source, out _));
        Assert.Equal(Cell(sheet, idCell), source!.ItemId); Assert.Equal(Cell(sheet, pathCell), source.AbsolutePath);
        Assert.Equal(2, sheet.Descendants(Main + "col").Count(c => (string?)c.Attribute("hidden") == "1"));
        foreach (int part in new[] { 2, 3, 4 })
            Assert.DoesNotContain(Read(zip, $"xl/tables/table{part}.xml").Descendants(Main + "tableColumn"), c => ((string?)c.Attribute("name"))?.StartsWith("__FLT_", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void DirectoryRowsRemainInspectableWithoutBecomingRecursiveCopyRequests()
    {
        var entry = Entry() with { IsDirectory = true, SizeBytes = null };
        string json = FileCollectContract.CreateSourceRecord(entry, Guid.NewGuid().ToString("D"), out var error);
        Assert.Null(error);
        Assert.True(FileCollectContract.TryParseSourceRecord(json, out var source, out _));
        Assert.Equal("directory", source!.Kind); Assert.Null(source.SizeBytes);
        Assert.Equal(entry.AbsolutePath, source.AbsolutePath);
    }

    [Fact]
    public void OversizedSourceRecordIsExplicitlyUnavailableAndListingStillExists()
    {
        var entry = Entry() with { AbsolutePath = Path.Combine(directory, new string('a', 33_000), "file.txt") };
        string path = Path.Combine(directory, "long.xlsx");
        var result = new WorkbookWriter().Write(path, [new(entry)]);
        Assert.Equal(1, result.EntryCount); Assert.Equal(2, result.ErrorCount); // hyperlink plus copy contract
        SchemaValid(path);
        using var zip = ZipFile.OpenRead(path);
        var sheet = Read(zip, "xl/worksheets/sheet1.xml");
        string json = Cell(sheet, "N2");
        Assert.InRange(json.Length, 1, FileCollectContract.MaxSourceRecordCharacters);
        Assert.True(FileCollectContract.TryParseSourceRecord(json, out var source, out _));
        Assert.Equal("unavailable", source!.Kind); Assert.Equal("SourceRecordTooLong", source.Reason);
        Assert.Empty(source.AbsolutePath); Assert.Equal(Cell(sheet, "M2"), source.ItemId);
        Assert.Contains("FileCollectUnavailable", Read(zip, "xl/worksheets/sheet2.xml").ToString());
    }

    [Fact]
    public void MissingFileSizeProducesUnavailableMetadataInsteadOfInventingZero()
    {
        string json = FileCollectContract.CreateSourceRecord(Entry() with { SizeBytes = null }, Guid.NewGuid().ToString("D"), out var error);
        Assert.Equal("MissingFileSize", error);
        Assert.True(FileCollectContract.TryParseSourceRecord(json, out var source, out _));
        Assert.Equal("unavailable", source!.Kind);
    }

    [Fact]
    public void StructuralParserRejectsDamagedAmbiguousAndUnboundedRecords()
    {
        string good = FileCollectContract.CreateSourceRecord(Entry(), Guid.NewGuid().ToString("D"), out _);
        string[] bad = ["", "{}", "[]", "null", good + "{}", new string('x', 32768),
            good.Replace("\"kind\":\"file\"", "\"kind\":\"file\",\"kind\":\"directory\"", StringComparison.Ordinal),
            good.Replace("\"kind\":\"file\"", "\"kind\":\"link\"", StringComparison.Ordinal),
            good.Replace("\"kind\":\"file\"", "\"kind\":\"file\",\"execute\":\"bad\"", StringComparison.Ordinal),
            good.Replace("\"9007199254740993\"", "9007199254740993", StringComparison.Ordinal),
            good.Replace("\"9007199254740993\"", "\"-1\"", StringComparison.Ordinal),
            good.Replace("\"9007199254740993\"", "\"9223372036854775808\"", StringComparison.Ordinal),
            good.Replace("\"9007199254740993\"", "\"1e3\"", StringComparison.Ordinal),
            good.Replace(Entry().ModifiedAt.Ticks.ToString(CultureInfo.InvariantCulture), "3155378976000000000", StringComparison.Ordinal),
            good.Replace("\"itemId\":", "\"missingId\":", StringComparison.Ordinal)];
        foreach (string json in bad)
        {
            Assert.False(FileCollectContract.TryParseSourceRecord(json, out var source, out var error));
            Assert.Null(source); Assert.False(string.IsNullOrEmpty(error));
        }
    }

    [Theory]
    [InlineData("FileListToExcel", "0")]
    [InlineData("FileListToExcel", "2")]
    [InlineData("FileListToExcel", "1.0")]
    [InlineData("OtherProduct", "1")]
    public void UnsupportedVersionsOrProductsCannotOptIn(string product, string version) =>
        Assert.False(FileCollectContract.IsSupportedWorkbook(product, version, Guid.NewGuid().ToString("D")));

    [Fact]
    public void TableRolesAndIdsAreStrictAndDoNotInferFromWorksheetOrHeaderNames()
    {
        Assert.True(FileCollectContract.IsSupportedTable("FileListTable2", "files"));
        Assert.False(FileCollectContract.IsSupportedTable("Files", "files"));
        Assert.False(FileCollectContract.IsSupportedTable("FileListTable1", "summary"));
        Assert.False(FileCollectContract.IsSupportedTable("FileListTable0", "files"));
        Assert.False(FileCollectContract.IsSupportedTable("FileListTable01", "files"));
        Assert.False(FileCollectContract.IsSupportedWorkbook("FileListToExcel", "1", Guid.Empty.ToString("D")));
        Assert.False(FileCollectContract.IsSupportedWorkbook("FileListToExcel", "1", Guid.NewGuid().ToString("N")));
        Assert.False(FileCollectContract.IsValidId(" " + Guid.NewGuid().ToString("D")));
        Assert.False(FileCollectContract.IsValidId(Guid.NewGuid().ToString("D") + " "));
    }

    [Fact]
    public void MetadataFromScannerMatchesTheActualUtcTimestampExactly()
    {
        string sourcePath = Path.Combine(directory, "source.txt");
        File.WriteAllText(sourcePath, "source");
        var entry = new FileScanner().Enumerate(new([sourcePath], ScanMode.SelectedItems)).Single().Entry!;
        string json = FileCollectContract.CreateSourceRecord(entry, Guid.NewGuid().ToString("D"), out _);
        Assert.True(FileCollectContract.TryParseSourceRecord(json, out var source, out _));
        Assert.Equal(new FileInfo(sourcePath).Length, source!.SizeBytes);
        Assert.Equal(File.GetLastWriteTimeUtc(sourcePath).Ticks, source.ModifiedUtcTicks);
    }
}
