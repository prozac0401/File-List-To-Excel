using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.Win32.SafeHandles;
using FileListToExcel.Core;

namespace FileListToExcel.App;

public sealed record CollectRequest(Guid RequestId, Guid WorkbookId, string TableName, IReadOnlyList<CollectRow> Rows);

/// <summary>Consumes only a validated, owned GUID request. Metadata is format identification, never authorization.</summary>
public static class CollectRequestReader
{
    public const int MaximumBytes = 33_554_432;
    public const int MaximumRows = 10_000;

    public static CollectRequest Read(string path)
    {
        if (!Path.IsPathFullyQualified(path)) throw new ArgumentException("요청 파일에는 절대 경로가 필요합니다.");
        path = Path.GetFullPath(path);
        if (!string.Equals(Path.GetDirectoryName(path), CommandLine.RequestDirectory, StringComparison.OrdinalIgnoreCase)
            || !FileCollectContract.IsValidId(Path.GetFileNameWithoutExtension(path)) || !Guid.TryParseExact(Path.GetFileNameWithoutExtension(path), "D", out var fileId)
            || Path.GetExtension(path) != ".json")
            throw new ArgumentException("제품 소유 GUID 요청 파일만 사용할 수 있습니다.");
        var guards = new List<SafeFileHandle>();
        try
        {
            // Pin each ancestor against rename/delete. Never traverse a redirected request directory.
            var ancestors = new Stack<string>();
            for (string? p = Path.GetDirectoryName(path); p != null; p = Path.GetDirectoryName(p)) ancestors.Push(p);
            while (ancestors.TryPop(out var directory)) guards.Add(OpenChecked(directory, true));
            using var handle = OpenChecked(path, false);
            using var stream = new FileStream(handle, FileAccess.Read);
            if (stream.Length is <= 0 or > MaximumBytes) throw new ArgumentException("요청 파일 크기 한도를 초과했습니다.");
            using var json = JsonDocument.Parse(stream, new JsonDocumentOptions { MaxDepth = 12 });
            var request = Parse(json.RootElement, fileId);
            // Delete by the SAME validated handle, not by a pathname that can be replaced after close.
            byte delete = 1;
            if (!SetFileInformationByHandle(handle, 4, ref delete, 1))
                throw new IOException("요청 파일을 일회성으로 소비하지 못했습니다.", new Win32Exception(Marshal.GetLastWin32Error()));
            return request;
        }
        finally { foreach (var guard in guards) guard.Dispose(); }
    }

    private static CollectRequest Parse(JsonElement root, Guid expectedId)
    {
        CheckProperties(root, ["mode", "version", "requestId", "workbookId", "tableName", "rows"]);
        if (root.GetProperty("mode").GetString() != "collect-files" || root.GetProperty("version").GetInt32() != 1)
            throw new ArgumentException("지원하지 않는 복사 요청 버전입니다.");
        if (!FileCollectContract.IsValidId(root.GetProperty("requestId").GetString()) || !Guid.TryParseExact(root.GetProperty("requestId").GetString(), "D", out var requestId) || requestId != expectedId
            || !FileCollectContract.IsValidId(root.GetProperty("workbookId").GetString())
            || !Guid.TryParseExact(root.GetProperty("workbookId").GetString(), "D", out var workbookId))
            throw new ArgumentException("요청/통합문서 식별자가 유효하지 않습니다.");
        var table = root.GetProperty("tableName").GetString();
        if (!FileCollectContract.IsSupportedTable(table, "files"))
            throw new ArgumentException("지원되는 결과표가 아닙니다.");
        var rows = root.GetProperty("rows");
        if (rows.ValueKind != JsonValueKind.Array || rows.GetArrayLength() is < 1 or > MaximumRows)
            throw new ArgumentException("선택은 1개 이상 10,000행 이하여야 합니다.");
        var result = new List<CollectRow>(rows.GetArrayLength());
        var ids = new HashSet<Guid>();
        foreach (var row in rows.EnumerateArray())
        {
            CheckProperties(row, ["itemId", "displayPath", "sourceRecord"], ["displayName"]);
            var id = row.GetProperty("itemId").GetString();
            var display = row.GetProperty("displayPath").GetString();
            var record = row.GetProperty("sourceRecord").GetString();
            if (!Guid.TryParseExact(id, "D", out var itemId) || !ids.Add(itemId))
                throw new ArgumentException("선택 행의 항목 식별자가 손상되었거나 중복되었습니다.");
            if (display is null || display.Length > 32767 || record is null || record.Length > 32767)
                throw new ArgumentException("선택 행의 경로/원본 정보가 손상되었습니다.");
            if (!FileCollectContract.TryParseSourceRecord(record, out var source, out var reason)
                || source!.ItemId != id || (source.Kind != "unavailable" && !string.Equals(display, source.AbsolutePath, StringComparison.Ordinal)))
                throw new ArgumentException(reason ?? "표시된 경로와 원본 정보가 다릅니다. 새 목록을 생성하세요.");
            if (row.TryGetProperty("displayName", out var name) && source.Kind == "file"
                && !string.Equals(name.GetString(), Path.GetFileName(display.TrimEnd('\\', '/')), StringComparison.Ordinal))
                throw new ArgumentException("표시된 파일명과 원본 정보가 다릅니다. 새 목록을 생성하세요.");
            result.Add(new(id!, display, record));
        }
        return new(requestId, workbookId, table!, result.AsReadOnly());
    }

    private static void CheckProperties(JsonElement value, string[] required, string[]? optional = null)
    {
        if (value.ValueKind != JsonValueKind.Object) throw new ArgumentException("요청 형식이 유효하지 않습니다.");
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var property in value.EnumerateObject())
            if (!names.Add(property.Name) || (!required.Contains(property.Name) && !(optional?.Contains(property.Name) ?? false)))
                throw new ArgumentException("요청에 중복되거나 지원하지 않는 필드가 있습니다.");
        if (required.Any(p => !names.Contains(p))) throw new ArgumentException("필수 요청 필드가 없습니다.");
    }

    private static SafeFileHandle OpenChecked(string path, bool directory)
    {
        const uint Read = 0x80000000, Delete = 0x10000, ReadAttributes = 0x80;
        var handle = CreateFile(path, directory ? ReadAttributes : Read | Delete, 1u,
            IntPtr.Zero, 3, 0x00200000 | 0x00100000 | (directory ? 0x02000000u : 0u), IntPtr.Zero);
        if (handle.IsInvalid) { handle.Dispose(); throw new IOException("요청 파일/폴더를 열 수 없습니다.", new Win32Exception(Marshal.GetLastWin32Error())); }
        if (!GetFileInformationByHandle(handle, out var info)
            || (info.Attributes & (0x400u | 0x1000u | 0x40000u | 0x400000u)) != 0
            || ((info.Attributes & 0x10) != 0) != directory || (!directory && info.NumberOfLinks != 1))
        { handle.Dispose(); throw new IOException("링크/클라우드/손상된 요청 파일은 사용할 수 없습니다."); }
        return handle;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct FileInfo
    {
        public uint Attributes;
        public System.Runtime.InteropServices.ComTypes.FILETIME CreationTime, AccessTime, WriteTime;
        public uint Volume, SizeHigh, SizeLow, NumberOfLinks, IndexHigh, IndexLow;
    }
    [DllImport("kernel32.dll", EntryPoint = "CreateFileW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateFile(string name, uint access, uint share, IntPtr security, uint disposition, uint flags, IntPtr template);
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetFileInformationByHandle(SafeFileHandle handle, out FileInfo info);
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetFileInformationByHandle(SafeFileHandle handle, int infoClass, ref byte info, uint size);
}
