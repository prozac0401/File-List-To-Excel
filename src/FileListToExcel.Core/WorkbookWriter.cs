using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml;

namespace FileListToExcel.Core;

/// <summary>Streams OOXML without Office automation. Each Files sheet holds at most 65,000 links.</summary>
public sealed class WorkbookWriter
{
    public const int RowsPerSheet = 65_000;
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly string[] Headers = ["이름", "종류", "크기", "크기(표시)", "수정일", "생성일", "상위 폴더", "상대경로", "전체경로", "깊이", "속성", "기준 폴더"];
    private static readonly double[] Widths = [36, 12, 18, 16, 22, 22, 25, 45, 65, 10, 28, 55];
    private static readonly XmlWriterSettings XmlSettings = new() { Encoding = new UTF8Encoding(false), CloseOutput = true, CheckCharacters = true };

    public ExportResult Write(string destinationPath, IEnumerable<ScanItem> items, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);
        ArgumentNullException.ThrowIfNull(items);
        string destination = Path.GetFullPath(destinationPath);
        string parent = Path.GetDirectoryName(destination)!;
        Directory.CreateDirectory(parent);
        string staging = Path.Combine(parent, $".{Path.GetFileName(destination)}.{Guid.NewGuid():N}.partial");
        string errorSpool = Path.Combine(Path.GetTempPath(), $"FileListToExcel-errors-{Guid.NewGuid():N}.jsonl");
        long entries = 0, errors = 0;
        int fileSheets = 0;
        try
        {
            using (var errorStream = new FileStream(errorSpool, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.DeleteOnClose))
            using (var errorWriter = new StreamWriter(errorStream, new UTF8Encoding(false), 4096, leaveOpen: true))
            using (var output = new FileStream(staging, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create))
            {
                var sheets = new List<SheetInfo>();
                SheetBuilder? sheet = null;
                bool OwnTemporaryPath(string path) => string.Equals(path, staging, OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal)
                    || string.Equals(path, errorSpool, OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);
                void LogError(ScanError error)
                {
                    errorWriter.WriteLine(JsonSerializer.Serialize(error));
                    errors++;
                }
                SheetBuilder NextFilesSheet()
                {
                    fileSheets++;
                    var info = new SheetInfo(sheets.Count + 1, fileSheets == 1 ? "Files" : $"Files_{fileSheets}", Headers);
                    sheets.Add(info);
                    return new SheetBuilder(archive, info, Widths);
                }
                try
                {
                    sheet = NextFilesSheet();
                    foreach (var item in items)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        // The destination may live inside a selected tree. Do not export our own work files.
                        if (item.Error is not null && !OwnTemporaryPath(item.Error.Path)) LogError(item.Error);
                        if (item.Entry is not { } entry || OwnTemporaryPath(entry.AbsolutePath)) continue;
                        if (sheet.Rows == RowsPerSheet) { sheet.Dispose(); sheet = NextFilesSheet(); }
                        string? link = FileLink(entry.AbsolutePath);
                        if (link is null) LogError(new(entry.AbsolutePath, "HyperlinkUnavailable", "The full path is preserved, but Excel cannot represent this path as a hyperlink (invalid URI or more than 2,079 URI characters)."));
                        sheet.WriteEntry(entry, link);
                        entries++;
                    }
                }
                finally { sheet?.Dispose(); }
                errorWriter.Flush();
                if (errors > 0)
                {
                    errorStream.Position = 0;
                    using var reader = new StreamReader(errorStream, Encoding.UTF8, true, 4096, leaveOpen: true);
                    int errorSheet = 0;
                    SheetBuilder? errorBuilder = null;
                    try
                    {
                        while (reader.ReadLine() is { } line)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            if (errorBuilder is null || errorBuilder.Rows == RowsPerSheet)
                            {
                                errorBuilder?.Dispose();
                                errorSheet++;
                                var info = new SheetInfo(sheets.Count + 1, errorSheet == 1 ? "Errors" : $"Errors_{errorSheet}", ["경로", "오류", "상세"]);
                                sheets.Add(info);
                                errorBuilder = new SheetBuilder(archive, info, [70, 28, 100]);
                            }
                            errorBuilder.WriteError(JsonSerializer.Deserialize<ScanError>(line)!);
                        }
                    }
                    finally { errorBuilder?.Dispose(); }
                }
                cancellationToken.ThrowIfCancellationRequested();
                WriteStyles(archive);
                WritePackage(archive, sheets);
            }
            cancellationToken.ThrowIfCancellationRequested();
            // Never replace a user's existing workbook; incomplete work never has an .xlsx extension.
            File.Move(staging, destination, overwrite: false);
            return new(destination, entries, errors, fileSheets);
        }
        finally
        {
            try { if (File.Exists(staging)) File.Delete(staging); } catch (IOException) { }
            try { if (File.Exists(errorSpool)) File.Delete(errorSpool); } catch (IOException) { }
        }
    }

    private sealed record SheetInfo(int Id, string Name, string[] Headers);

    private sealed class SheetBuilder : IDisposable
    {
        private readonly ZipArchive archive;
        private readonly SheetInfo info;
        private readonly XmlWriter writer;
        private readonly List<string> links = [];
        private bool disposed;
        public int Rows { get; private set; }

        public SheetBuilder(ZipArchive archive, SheetInfo info, double[] widths)
        {
            this.archive = archive; this.info = info;
            writer = Xml(archive, $"xl/worksheets/sheet{info.Id}.xml");
            writer.WriteStartElement("worksheet", Main);
            writer.WriteAttributeString("xmlns", "r", null, Rel);
            writer.WriteStartElement("sheetViews", Main);
            writer.WriteStartElement("sheetView", Main); writer.WriteAttributeString("workbookViewId", "0");
            writer.WriteStartElement("pane", Main);
            Attr(writer, "ySplit", "1", "topLeftCell", "A2", "activePane", "bottomLeft", "state", "frozen");
            writer.WriteEndElement(); writer.WriteEndElement(); writer.WriteEndElement();
            writer.WriteStartElement("cols", Main);
            for (int i = 0; i < widths.Length; i++)
            {
                writer.WriteStartElement("col", Main);
                Attr(writer, "min", (i + 1).ToString(CultureInfo.InvariantCulture), "max", (i + 1).ToString(CultureInfo.InvariantCulture), "width", widths[i].ToString(CultureInfo.InvariantCulture), "customWidth", "1");
                writer.WriteEndElement();
            }
            writer.WriteEndElement(); writer.WriteStartElement("sheetData", Main);
            writer.WriteStartElement("row", Main); writer.WriteAttributeString("r", "1");
            for (int i = 0; i < info.Headers.Length; i++) TextCell(writer, Cell(i, 1), info.Headers[i]);
            writer.WriteEndElement();
        }

        public void WriteEntry(FileEntry entry, string? link)
        {
            int row = ++Rows + 1;
            writer.WriteStartElement("row", Main); writer.WriteAttributeString("r", row.ToString(CultureInfo.InvariantCulture));
            TextCell(writer, Cell(0, row), entry.Name, link is null ? 0 : 3);
            TextCell(writer, Cell(1, row), entry.IsDirectory ? "폴더" : entry.Extension);
            if (entry.SizeBytes is { } size) NumberCell(writer, Cell(2, row), size, 1);
            TextCell(writer, Cell(3, row), entry.SizeBytes is { } bytes ? HumanSize(bytes) : string.Empty);
            DateCell(writer, Cell(4, row), entry.ModifiedAt);
            DateCell(writer, Cell(5, row), entry.CreatedAt);
            TextCell(writer, Cell(6, row), entry.ParentName);
            TextCell(writer, Cell(7, row), entry.RelativePath);
            TextCell(writer, Cell(8, row), entry.AbsolutePath);
            NumberCell(writer, Cell(9, row), entry.Depth);
            TextCell(writer, Cell(10, row), entry.Attributes.ToString());
            TextCell(writer, Cell(11, row), entry.RootPath);
            writer.WriteEndElement();
            links.Add(link ?? string.Empty);
        }

        public void WriteError(ScanError error)
        {
            int row = ++Rows + 1;
            writer.WriteStartElement("row", Main); writer.WriteAttributeString("r", row.ToString(CultureInfo.InvariantCulture));
            TextCell(writer, Cell(0, row), error.Path); TextCell(writer, Cell(1, row), error.ErrorType); TextCell(writer, Cell(2, row), error.Message);
            writer.WriteEndElement();
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            writer.WriteEndElement(); // sheetData
            if (links.Any(link => link.Length > 0))
            {
                writer.WriteStartElement("hyperlinks", Main);
                for (int i = 0; i < links.Count; i++)
                {
                    if (links[i].Length == 0) continue;
                    writer.WriteStartElement("hyperlink", Main); writer.WriteAttributeString("ref", $"A{i + 2}");
                    writer.WriteAttributeString("r", "id", Rel, $"link{i + 1}"); writer.WriteEndElement();
                }
                writer.WriteEndElement();
            }
            writer.WriteStartElement("tableParts", Main); writer.WriteAttributeString("count", "1");
            writer.WriteStartElement("tablePart", Main); writer.WriteAttributeString("r", "id", Rel, "table");
            writer.WriteEndElement(); writer.WriteEndElement(); writer.WriteEndElement(); writer.Dispose();
            using (var rels = Xml(archive, $"xl/worksheets/_rels/sheet{info.Id}.xml.rels"))
            {
                rels.WriteStartElement("Relationships", PackageRel);
                Relationship(rels, "table", "table", $"../tables/table{info.Id}.xml");
                for (int i = 0; i < links.Count; i++)
                    if (links[i].Length > 0) Relationship(rels, $"link{i + 1}", "hyperlink", links[i], external: true);
                rels.WriteEndElement();
            }
            using var table = Xml(archive, $"xl/tables/table{info.Id}.xml");
            string range = $"A1:{(char)('A' + info.Headers.Length - 1)}{Math.Max(Rows, 1) + 1}";
            table.WriteStartElement("table", Main);
            Attr(table, "id", info.Id.ToString(CultureInfo.InvariantCulture), "name", $"FileListTable{info.Id}", "displayName", $"FileListTable{info.Id}", "ref", range, "totalsRowShown", "0");
            table.WriteStartElement("autoFilter", Main); table.WriteAttributeString("ref", range); table.WriteEndElement();
            table.WriteStartElement("tableColumns", Main); table.WriteAttributeString("count", info.Headers.Length.ToString(CultureInfo.InvariantCulture));
            for (int i = 0; i < info.Headers.Length; i++)
            {
                table.WriteStartElement("tableColumn", Main); Attr(table, "id", (i + 1).ToString(CultureInfo.InvariantCulture), "name", info.Headers[i]); table.WriteEndElement();
            }
            table.WriteEndElement();
            table.WriteStartElement("tableStyleInfo", Main); Attr(table, "name", "TableStyleMedium2", "showFirstColumn", "0", "showLastColumn", "0", "showRowStripes", "1", "showColumnStripes", "0");
            table.WriteEndElement(); table.WriteEndElement();
        }
    }

    private static void WritePackage(ZipArchive archive, List<SheetInfo> sheets)
    {
        using (var content = Xml(archive, "[Content_Types].xml"))
        {
            const string ns = "http://schemas.openxmlformats.org/package/2006/content-types";
            content.WriteStartElement("Types", ns);
            content.WriteStartElement("Default", ns); Attr(content, "Extension", "rels", "ContentType", "application/vnd.openxmlformats-package.relationships+xml"); content.WriteEndElement();
            content.WriteStartElement("Default", ns); Attr(content, "Extension", "xml", "ContentType", "application/xml"); content.WriteEndElement();
            void Override(string name, string type)
            {
                content.WriteStartElement("Override", ns); Attr(content, "PartName", name, "ContentType", $"application/vnd.openxmlformats-officedocument.spreadsheetml.{type}+xml"); content.WriteEndElement();
            }
            Override("/xl/workbook.xml", "sheet.main"); Override("/xl/styles.xml", "styles");
            foreach (var sheet in sheets) { Override($"/xl/worksheets/sheet{sheet.Id}.xml", "worksheet"); Override($"/xl/tables/table{sheet.Id}.xml", "table"); }
            content.WriteEndElement();
        }
        using (var rels = Xml(archive, "_rels/.rels"))
        {
            rels.WriteStartElement("Relationships", PackageRel); Relationship(rels, "workbook", "officeDocument", "xl/workbook.xml"); rels.WriteEndElement();
        }
        using (var workbook = Xml(archive, "xl/workbook.xml"))
        {
            workbook.WriteStartElement("workbook", Main); workbook.WriteAttributeString("xmlns", "r", null, Rel);
            workbook.WriteStartElement("sheets", Main);
            foreach (var sheet in sheets)
            {
                workbook.WriteStartElement("sheet", Main); Attr(workbook, "name", sheet.Name, "sheetId", sheet.Id.ToString(CultureInfo.InvariantCulture));
                workbook.WriteAttributeString("r", "id", Rel, $"sheet{sheet.Id}"); workbook.WriteEndElement();
            }
            workbook.WriteEndElement(); workbook.WriteEndElement();
        }
        using var wbRels = Xml(archive, "xl/_rels/workbook.xml.rels");
        wbRels.WriteStartElement("Relationships", PackageRel);
        foreach (var sheet in sheets) Relationship(wbRels, $"sheet{sheet.Id}", "worksheet", $"worksheets/sheet{sheet.Id}.xml");
        Relationship(wbRels, "styles", "styles", "styles.xml"); wbRels.WriteEndElement();
    }

    private static void WriteStyles(ZipArchive archive)
    {
        using var writer = Xml(archive, "xl/styles.xml");
        writer.WriteStartElement("styleSheet", Main);
        writer.WriteStartElement("numFmts", Main); writer.WriteAttributeString("count", "1");
        writer.WriteStartElement("numFmt", Main); Attr(writer, "numFmtId", "164", "formatCode", "yyyy-mm-dd hh:mm:ss"); writer.WriteEndElement(); writer.WriteEndElement();
        writer.WriteStartElement("fonts", Main); writer.WriteAttributeString("count", "2");
        for (int i = 0; i < 2; i++)
        {
            writer.WriteStartElement("font", Main);
            if (i == 1) writer.WriteElementString("u", Main, "");
            writer.WriteStartElement("sz", Main); writer.WriteAttributeString("val", "11"); writer.WriteEndElement();
            if (i == 1) { writer.WriteStartElement("color", Main); writer.WriteAttributeString("rgb", "FF0563C1"); writer.WriteEndElement(); }
            writer.WriteStartElement("name", Main); writer.WriteAttributeString("val", "맑은 고딕"); writer.WriteEndElement();
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
        writer.WriteStartElement("fills", Main); writer.WriteAttributeString("count", "2");
        foreach (var pattern in new[] { "none", "gray125" })
        {
            writer.WriteStartElement("fill", Main); writer.WriteStartElement("patternFill", Main); writer.WriteAttributeString("patternType", pattern); writer.WriteEndElement(); writer.WriteEndElement();
        }
        writer.WriteEndElement(); writer.WriteStartElement("borders", Main); writer.WriteAttributeString("count", "1"); writer.WriteElementString("border", Main, ""); writer.WriteEndElement();
        writer.WriteStartElement("cellStyleXfs", Main); writer.WriteAttributeString("count", "1");
        writer.WriteStartElement("xf", Main); Attr(writer, "numFmtId", "0", "fontId", "0", "fillId", "0", "borderId", "0"); writer.WriteEndElement(); writer.WriteEndElement();
        writer.WriteStartElement("cellXfs", Main); writer.WriteAttributeString("count", "4");
        foreach (var format in new[] { (0, 0), (3, 0), (164, 0), (0, 1) })
        {
            writer.WriteStartElement("xf", Main); Attr(writer, "numFmtId", format.Item1.ToString(CultureInfo.InvariantCulture), "fontId", format.Item2.ToString(CultureInfo.InvariantCulture), "fillId", "0", "borderId", "0", "xfId", "0");
            if (format.Item1 > 0) writer.WriteAttributeString("applyNumberFormat", "1");
            if (format.Item2 > 0) writer.WriteAttributeString("applyFont", "1");
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
        writer.WriteStartElement("cellStyles", Main); writer.WriteAttributeString("count", "1"); writer.WriteStartElement("cellStyle", Main); Attr(writer, "name", "Normal", "xfId", "0", "builtinId", "0"); writer.WriteEndElement(); writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static XmlWriter Xml(ZipArchive archive, string name) => XmlWriter.Create(archive.CreateEntry(name, CompressionLevel.Fastest).Open(), XmlSettings);
    private static string Cell(int column, int row) => $"{(char)('A' + column)}{row}";
    private static void Attr(XmlWriter writer, params string[] attributes) { for (int i = 0; i < attributes.Length; i += 2) writer.WriteAttributeString(attributes[i], attributes[i + 1]); }
    private static void Relationship(XmlWriter writer, string id, string type, string target, bool external = false)
    {
        writer.WriteStartElement("Relationship", PackageRel); Attr(writer, "Id", id, "Type", $"{Rel}/{type}", "Target", target);
        if (external) writer.WriteAttributeString("TargetMode", "External"); writer.WriteEndElement();
    }
    private static void TextCell(XmlWriter writer, string cell, string text, int style = 0)
    {
        writer.WriteStartElement("c", Main); Attr(writer, "r", cell, "t", "inlineStr");
        if (style > 0) writer.WriteAttributeString("s", style.ToString(CultureInfo.InvariantCulture));
        writer.WriteStartElement("is", Main); writer.WriteStartElement("t", Main); writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
        writer.WriteString(ExcelString(text)); writer.WriteEndElement(); writer.WriteEndElement(); writer.WriteEndElement();
    }
    private static void NumberCell(XmlWriter writer, string cell, double number, int style = 0)
    {
        writer.WriteStartElement("c", Main); writer.WriteAttributeString("r", cell);
        if (style > 0) writer.WriteAttributeString("s", style.ToString(CultureInfo.InvariantCulture));
        writer.WriteElementString("v", Main, number.ToString("R", CultureInfo.InvariantCulture)); writer.WriteEndElement();
    }
    private static void DateCell(XmlWriter writer, string cell, DateTime date)
    {
        // Excel cannot display pre-1900 metadata as a normal date. Preserve that uncommon value as text.
        if (date < new DateTime(1900, 1, 1)) TextCell(writer, cell, date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
        else
        {
            // OLE dates include day 0; Excel inserts a fictitious 1900-02-29.
            double serial = date.ToOADate();
            if (date < new DateTime(1900, 3, 1)) serial--;
            NumberCell(writer, cell, serial, 2);
        }
    }
    private static string HumanSize(long bytes)
    {
        string[] units = ["B", "KB", "MB", "GB", "TB", "PB", "EB"];
        double size = bytes; int unit = 0;
        while (size >= 1024 && unit < units.Length - 1) { size /= 1024; unit++; }
        return $"{size.ToString(unit == 0 ? "0" : "0.##", CultureInfo.InvariantCulture)} {units[unit]}";
    }
    private static string? FileLink(string path)
    {
        try
        {
            if (!Path.IsPathFullyQualified(path)) return null;
            string fullPath = Path.GetFullPath(path);
            if (OperatingSystem.IsWindows())
            {
                if (fullPath.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase)) fullPath = @"\\" + fullPath[8..];
                else if (fullPath.StartsWith(@"\\?\", StringComparison.Ordinal)) fullPath = fullPath[4..];
                if (fullPath.StartsWith(@"\\.\", StringComparison.Ordinal)) return null;
                fullPath = fullPath.Replace('\\', '/');
            }
            // Build an IRI from filesystem characters directly. Uri(path) can reinterpret literal percent
            // sequences and corrupt an emoji following %, while Excel decodes escaped UTF-8 using ANSI.
            // Preserve Unicode; percent-escape ASCII filename punctuation, including %, #, ?, and spaces.
            var target = new StringBuilder(fullPath.Length + 8);
            target.Append(fullPath.StartsWith("//", StringComparison.Ordinal) ? "file:" : fullPath.StartsWith('/') ? "file://" : "file:///");
            foreach (char c in fullPath)
            {
                if (c >= 128 || char.IsAsciiLetterOrDigit(c) || c is '/' or ':' or '-' or '.' or '_' or '~') target.Append(c);
                else target.Append('%').Append(((int)c).ToString("X2", CultureInfo.InvariantCulture));
            }
            string result = target.ToString();
            XmlConvert.VerifyXmlChars(result);
            return result.Length <= 2079 ? result : null;
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or XmlException) { return null; }
    }
    private static string ExcelString(string value)
    {
        if (value.Length > 32767) value = value[..(char.IsHighSurrogate(value[32766]) ? 32766 : 32767)];
        value = Regex.Replace(value, "_x[0-9A-Fa-f]{4}_", match => "_x005F_" + match.Value[1..], RegexOptions.CultureInvariant);
        var text = new StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++)
        {
            char c = value[i];
            if (char.IsHighSurrogate(c) && i + 1 < value.Length && char.IsLowSurrogate(value[i + 1])) { text.Append(c).Append(value[++i]); continue; }
            if (!XmlConvert.IsXmlChar(c) || char.IsSurrogate(c)) text.Append($"_x{(int)c:X4}_"); else text.Append(c);
        }
        return text.ToString();
    }
}
