namespace FileListToExcel.Core;

/// <summary>A value snapshot of one visible Excel table row, never an Excel COM object.</summary>
public sealed record CollectRow(string ItemId, string DisplayPath, string SourceRecord);
public enum CopyStatus { Copied, Excluded, Failed, Cancelled, Duplicate }
public sealed record CopyOutcome(int RowIndex, string ItemId, string SourcePath, string? DestinationPath,
    CopyStatus Status, string Code, string Message);
public sealed record CopyProgress(string FileName, int CompletedCount, int TotalCount, long TransferredBytes, long TotalBytes);

public sealed class PlannedCopyFile
{
    internal PlannedCopyFile(int rowIndex, CollectRow row, FileCollectSourceRecord source, CopyNative.FileStamp stamp, string destinationName)
        => (RowIndex, Row, Source, Stamp, DestinationName) = (rowIndex, row, source, stamp, destinationName);
    public int RowIndex { get; }
    public CollectRow Row { get; }
    public FileCollectSourceRecord Source { get; }
    public string DestinationName { get; }
    internal CopyNative.FileStamp Stamp { get; }
}

public sealed class CopyPlan
{
    internal CopyPlan(int selectedCount, IEnumerable<PlannedCopyFile> files, IEnumerable<CopyOutcome> outcomes)
    {
        SelectedCount = selectedCount;
        EligibleFiles = Array.AsReadOnly(files.ToArray());
        Outcomes = Array.AsReadOnly(outcomes.ToArray());
        TotalBytes = EligibleFiles.Aggregate(0L, (total, item) => item.Stamp.Length > long.MaxValue - total ? long.MaxValue : total + item.Stamp.Length);
    }
    public int SelectedCount { get; }
    public IReadOnlyList<PlannedCopyFile> EligibleFiles { get; }
    public IReadOnlyList<CopyOutcome> Outcomes { get; }
    public long TotalBytes { get; }
    public int DuplicateCount => Outcomes.Count(x => x.Status == CopyStatus.Duplicate);
    public int ExcludedCount => Outcomes.Count(x => x.Status == CopyStatus.Excluded);
}

public sealed record CopyResult(string? OutputDirectory, string? ReportPath, IReadOnlyList<CopyOutcome> Outcomes, bool Cancelled,
    string? ReportError = null)
{
    public int CopiedCount => Outcomes.Count(x => x.Status == CopyStatus.Copied);
    public int FailedCount => Outcomes.Count(x => x.Status == CopyStatus.Failed);
    public int ExcludedCount => Outcomes.Count(x => x.Status is CopyStatus.Excluded or CopyStatus.Duplicate);
    public int CancelledCount => Outcomes.Count(x => x.Status == CopyStatus.Cancelled);
}

internal sealed class CollectSafetyException(string code, string message) : IOException(message)
{
    public string Code { get; } = code;
}
