namespace FileListToExcel.Core;

public enum DuplicateStage { FindingFiles, SizeGrouping, QuickFingerprint, FullHash, Completed }

public sealed record DuplicateProgress(DuplicateStage Stage, long FileCount, long SizeCandidates,
    long QuickCandidates, long HashedFiles, string CurrentPath);

public sealed record DuplicateFile(FileEntry Entry, string? QuickHash, string? FullSha256);

public sealed record DuplicateGroup(int GroupId, IReadOnlyList<DuplicateFile> Files)
{
    public long SizeBytes => Files[0].Entry.SizeBytes ?? 0;
    public long RecoverableBytes => checked(SizeBytes * (Files.Count - 1));
    public string FullSha256 => Files[0].FullSha256!;
}

public sealed record DuplicateStatistics(long ScannedFiles, long SizeCandidates, long QuickCandidates,
    long QuickHashCount, long FullHashCount, long CacheHitCount, long BytesRead, TimeSpan Elapsed)
{
    public long HashedFiles => QuickHashCount + FullHashCount;
}

public sealed record DuplicateResult(IReadOnlyList<DuplicateGroup> Groups, DuplicateFile? ReferenceFile,
    DuplicateStatistics Stats, IReadOnlyList<ScanError> Errors, IReadOnlyList<ScanError> Skipped)
{
    public long DuplicateFileCount => Groups.Sum(group => (long)group.Files.Count);
    public long RecoverableBytes => Groups.Sum(group => group.RecoverableBytes);
}

/// <summary>UTC ticks retain the native 100 ns timestamp precision on Windows.</summary>
public sealed record HashCacheKey(string AbsolutePath, long SizeBytes, long ModifiedUtcTicks, string Algorithm);
public sealed record HashCacheValue(string? QuickHash, string? FullSha256, DateTime CheckedAtUtc);

/// <summary>Implementations must be safe for calls from up to two hash workers. Failures are optional-cache failures.</summary>
public interface IDuplicateHashCache
{
    bool TryGet(HashCacheKey key, out HashCacheValue? value);
    void Store(HashCacheKey key, HashCacheValue value);
}

public sealed record ContentHash(string Hash, long BytesRead);

public interface IFileContentHasher
{
    Task<ContentHash> QuickAsync(string path, long sizeBytes, CancellationToken cancellationToken);
    Task<ContentHash> FullAsync(string path, CancellationToken cancellationToken);
}

/// <summary>Called with freshly read metadata before opening any contents, including cache hits.</summary>
public interface IFileContentPolicy
{
    ScanError? GetSkipReason(string path, FileAttributes attributes);
}

public sealed class LocalFileContentPolicy : IFileContentPolicy
{
    // Windows FILE_ATTRIBUTE_RECALL_ON_OPEN and RECALL_ON_DATA_ACCESS are not exposed by FileAttributes.
    public const FileAttributes RecallOnOpen = (FileAttributes)0x00040000;
    public const FileAttributes RecallOnDataAccess = (FileAttributes)0x00400000;

    public ScanError? GetSkipReason(string path, FileAttributes attributes)
    {
        if ((attributes & (FileAttributes.Offline | RecallOnOpen | RecallOnDataAccess)) != 0)
            return new(path, "CloudOnly", "건너뜀: Cloud-only 파일 (로컬에 내용 없음)");
        if ((attributes & FileAttributes.ReparsePoint) != 0)
            return new(path, "ReparsePoint", "건너뜀: 파일 Reparse Point를 따라가지 않습니다.");
        return null;
    }
}
