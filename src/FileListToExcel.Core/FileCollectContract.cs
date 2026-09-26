using System.Globalization;
using System.Text.Json;

namespace FileListToExcel.Core;

/// <summary>Untrusted, versioned row metadata. This contract identifies format, never provenance or permission.</summary>
public sealed record FileCollectSourceRecord(string Kind, string AbsolutePath, long? SizeBytes,
    long ModifiedUtcTicks, string ItemId, string? Reason);

public static class FileCollectContract
{
    public const string Product = "FileListToExcel";
    public const string ProductProperty = "FLT.Product";
    public const string SchemaVersionProperty = "FLT.ActionSchemaVersion";
    public const string SchemaVersion = "1";
    public const string WorkbookIdProperty = "FLT.WorkbookId";
    public const string TablePropertyPrefix = "FLT.Table.";
    public const string ItemIdColumn = "__FLT_ItemId";
    public const string SourceRecordColumn = "__FLT_SourceRecord";
    public const string PathColumn = "전체경로";
    public const string NameColumn = "이름";
    public const int MaxSourceRecordCharacters = 32_767;
    public const int MaxSelectedFiles = 10_000;

    public static bool IsValidId(string? value) => value is { Length: 36 } && Guid.TryParseExact(value, "D", out var id) && id != Guid.Empty;

    public static bool IsSupportedWorkbook(string? product, string? schemaVersion, string? workbookId) =>
        product == Product && schemaVersion == SchemaVersion && IsValidId(workbookId);

    public static bool IsSupportedTable(string? tableName, string? role)
    {
        const string prefix = "FileListTable";
        if (role is not ("files" or "duplicates" or "matches") || tableName is null ||
            !tableName.StartsWith(prefix, StringComparison.Ordinal)) return false;
        var suffix = tableName.AsSpan(prefix.Length);
        return suffix.Length > 0 && suffix[0] != '0' && int.TryParse(suffix, NumberStyles.None,
            CultureInfo.InvariantCulture, out int number) && number > 0;
    }

    /// <summary>Uses captured metadata only; never probes or opens the source while generating a workbook.</summary>
    public static string CreateSourceRecord(FileEntry entry, string itemId, out string? unavailableReason)
    {
        ArgumentNullException.ThrowIfNull(entry);
        if (!IsValidId(itemId)) throw new ArgumentException("An item ID must be a nonempty UUID.", nameof(itemId));
        unavailableReason = null;
        if (!entry.IsDirectory && (entry.SizeBytes is null or < 0)) unavailableReason = "MissingFileSize";
        string json = JsonSerializer.Serialize(new
        {
            kind = entry.IsDirectory ? "directory" : "file",
            absolutePath = entry.AbsolutePath,
            sizeBytes = entry.IsDirectory ? null : entry.SizeBytes?.ToString(CultureInfo.InvariantCulture),
            modifiedUtcTicks = entry.ModifiedAt.ToUniversalTime().Ticks.ToString(CultureInfo.InvariantCulture),
            itemId
        });
        if (json.Length > MaxSourceRecordCharacters) unavailableReason = "SourceRecordTooLong";
        if (unavailableReason is null) return json;
        // Never truncate the technical record into a valid-looking but different path.
        return JsonSerializer.Serialize(new
        {
            kind = "unavailable", absolutePath = "", sizeBytes = (string?)null,
            modifiedUtcTicks = "0", itemId, reason = unavailableReason
        });
    }

    /// <summary>Strict structural parsing only. The caller must check displayed fields and revalidate the filesystem.</summary>
    public static bool TryParseSourceRecord(string? json, out FileCollectSourceRecord? record, out string? error)
    {
        record = null;
        error = "행의 원본 정보가 없거나 손상되었습니다. 새 파일목록을 생성하세요.";
        if (string.IsNullOrEmpty(json) || json.Length > MaxSourceRecordCharacters) return false;
        try
        {
            using var document = JsonDocument.Parse(json, new JsonDocumentOptions { MaxDepth = 4 });
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object) return false;
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (var property in root.EnumerateObject())
            {
                if (!names.Add(property.Name) || property.Name is not
                    ("kind" or "absolutePath" or "sizeBytes" or "modifiedUtcTicks" or "itemId" or "reason")) return false;
            }
            string? Text(string name) => root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String ? value.GetString() : null;
            string? kind = Text("kind"), path = Text("absolutePath"), ticksText = Text("modifiedUtcTicks"), itemId = Text("itemId"), reason = Text("reason");
            if (kind is not ("file" or "directory" or "unavailable") || path is null || !IsValidId(itemId) ||
                !TryUnsignedLong(ticksText, out long ticks) || ticks > DateTime.MaxValue.Ticks ||
                !root.TryGetProperty("sizeBytes", out var sizeValue)) return false;
            long? size = null;
            if (sizeValue.ValueKind == JsonValueKind.String)
            {
                if (!TryUnsignedLong(sizeValue.GetString(), out long parsedSize)) return false;
                size = parsedSize;
            }
            else if (sizeValue.ValueKind != JsonValueKind.Null) return false;
            if (kind == "file" && (size is null || string.IsNullOrWhiteSpace(path))) return false;
            if (kind == "directory" && (size is not null || string.IsNullOrWhiteSpace(path))) return false;
            if (kind == "unavailable" && (size is not null || path.Length != 0 || string.IsNullOrWhiteSpace(reason))) return false;
            if (kind != "unavailable" && names.Contains("reason")) return false;
            if (names.Contains("reason") && reason is null) return false;
            record = new(kind, path, size, ticks, itemId!, reason);
            error = null;
            return true;
        }
        catch (JsonException) { return false; }
    }

    private static bool TryUnsignedLong(string? value, out long number) =>
        long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out number) && number >= 0;
}
