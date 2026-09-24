namespace FileListToExcel.Core;

public enum ScanMode { SelectedItems, DirectChildren, Recursive }

public sealed record ScanRequest(IReadOnlyList<string> Paths, ScanMode Mode);

public sealed record FileEntry(
    string Name, string Extension, bool IsDirectory, long? SizeBytes,
    DateTime CreatedAt, DateTime ModifiedAt, string ParentName,
    string AbsolutePath, string RelativePath, int Depth,
    FileAttributes Attributes, string RootPath);

public sealed record ScanError(string Path, string ErrorType, string Message);

public sealed record ScanItem(FileEntry? Entry = null, ScanError? Error = null);

public sealed record ScanProgress(long ItemCount, long ErrorCount, string CurrentPath);

public sealed record ExportResult(string OutputPath, long EntryCount, long ErrorCount, int FileSheetCount);
