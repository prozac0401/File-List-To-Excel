using System.Collections.Concurrent;
using System.Diagnostics;
using System.Security.Cryptography;

namespace FileListToExcel.Core;

/// <summary>Reduces candidates by size, then sampled fingerprints, then full SHA-256. Never changes source files.</summary>
public sealed class DuplicateDetector
{
    private readonly IDuplicateHashCache? cache;
    private readonly IFileContentHasher hasher;
    private readonly IFileContentPolicy contentPolicy;
    private readonly int workerCount;
    private static readonly StringComparer PathComparer = OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
    private static readonly string EmptyHash = Convert.ToHexString(SHA256.HashData([]));

    public DuplicateDetector(IDuplicateHashCache? cache = null, IFileContentHasher? hasher = null,
        IFileContentPolicy? contentPolicy = null, int workerCount = 2)
    {
        if (workerCount is < 1 or > 2) throw new ArgumentOutOfRangeException(nameof(workerCount), "Use one or two disk workers.");
        this.cache = cache;
        this.hasher = hasher ?? new FileContentHasher();
        this.contentPolicy = contentPolicy ?? new LocalFileContentPolicy();
        this.workerCount = workerCount;
    }

    public async Task<DuplicateResult> DetectAsync(IReadOnlyList<FileEntry> entries, string? referencePath = null,
        IProgress<DuplicateProgress>? progress = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(entries);
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        var errors = new ConcurrentBag<ScanError>();
        var skipped = new ConcurrentBag<ScanError>();
        var states = entries.Where(entry => !entry.IsDirectory).DistinctBy(entry => entry.AbsolutePath, PathComparer)
            .Select(entry => new Candidate(entry)).ToArray();
        Candidate? reference = referencePath is null ? null : states.FirstOrDefault(state => PathComparer.Equals(state.Entry.AbsolutePath, referencePath));
        if (referencePath is not null && reference is null) throw new ArgumentException("The reference file must be present in the scanned entries.", nameof(referencePath));
        long sizeCandidates = 0, quickCandidates = 0, quickHashes = 0, fullHashes = 0, cacheHits = 0, bytesRead = 0;
        var progressGate = new object();
        void Report(DuplicateStage stage, string path = "")
        {
            lock (progressGate)
                progress?.Report(new(stage, states.LongLength, sizeCandidates, quickCandidates, quickHashes + fullHashes, path));
        }
        bool Validate(Candidate state)
        {
            if (state.Failed) return false;
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                var info = new FileInfo(state.Entry.AbsolutePath);
                info.Refresh();
                if (!info.Exists) throw new FileNotFoundException("The file no longer exists.", state.Entry.AbsolutePath);
                var skip = contentPolicy.GetSkipReason(state.Entry.AbsolutePath, state.Entry.Attributes)
                    ?? contentPolicy.GetSkipReason(state.Entry.AbsolutePath, info.Attributes);
                if (skip is not null) { skipped.Add(skip); state.Failed = true; return false; }
                if (info.Length != state.Entry.SizeBytes || info.LastWriteTimeUtc.Ticks != state.Entry.ModifiedAt.ToUniversalTime().Ticks)
                {
                    errors.Add(new(state.Entry.AbsolutePath, "ChangedDuringScan", "검사 중 파일 크기 또는 수정시간이 변경되었습니다."));
                    state.Failed = true;
                    return false;
                }
                return true;
            }
            catch (Exception ex) when (IsFileError(ex))
            {
                state.Failed = true;
                errors.Add(new(state.Entry.AbsolutePath, ex is UnauthorizedAccessException ? "AccessDenied" : ex.GetType().Name, ex.Message));
                return false;
            }
        }
        void MarkCacheHit(Candidate state)
        {
            if (state.CacheUsed) return;
            state.CacheUsed = true;
            Interlocked.Increment(ref cacheHits);
        }
        void ReadCache(Candidate state)
        {
            try
            {
                if (cache?.TryGet(state.Key, out var value) == true && value is not null)
                {
                    state.QuickHash = ValidHash(value.QuickHash) ? value.QuickHash!.ToUpperInvariant() : null;
                    state.FullHash = ValidHash(value.FullSha256) ? value.FullSha256!.ToUpperInvariant() : null;
                }
            }
            catch (Exception ex) when (IsCacheError(ex)) { /* Cache availability must never prevent a scan. */ }
        }
        void StoreCache(Candidate state)
        {
            try { cache?.Store(state.Key, new(state.QuickHash, state.FullHash, DateTime.UtcNow)); }
            catch (Exception ex) when (IsCacheError(ex)) { /* The verified in-memory result remains usable. */ }
        }
        async ValueTask Hash(Candidate state, bool full, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            if (!Validate(state)) return;
            Report(full ? DuplicateStage.FullHash : DuplicateStage.QuickFingerprint, state.Entry.AbsolutePath);
            try
            {
                bool small = state.Entry.SizeBytes <= FileContentHasher.SmallFileLimit;
                if (full || small)
                {
                    if (state.FullHash is not null) MarkCacheHit(state);
                    else
                    {
                        Interlocked.Increment(ref fullHashes);
                        var hash = await hasher.FullAsync(state.Entry.AbsolutePath, token).ConfigureAwait(false);
                        state.FullHash = hash.Hash;
                        Interlocked.Add(ref bytesRead, hash.BytesRead);
                    }
                    if (small) state.QuickHash = state.FullHash;
                }
                else if (state.QuickHash is not null) MarkCacheHit(state);
                else
                {
                    Interlocked.Increment(ref quickHashes);
                    var hash = await hasher.QuickAsync(state.Entry.AbsolutePath, state.Entry.SizeBytes!.Value, token).ConfigureAwait(false);
                    state.QuickHash = hash.Hash;
                    Interlocked.Add(ref bytesRead, hash.BytesRead);
                }
                token.ThrowIfCancellationRequested();
                if (Validate(state)) StoreCache(state);
            }
            catch (Exception ex) when (IsFileError(ex))
            {
                // Check metadata even when a concurrent truncation caused an early EOF/read failure.
                if (Validate(state))
                {
                    errors.Add(new(state.Entry.AbsolutePath, ex is FileChangedDuringScanException ? "ChangedDuringScan" :
                        ex is UnauthorizedAccessException ? "AccessDenied" : ex.GetType().Name, ex.Message));
                    state.Failed = true;
                }
            }
        }
        Report(DuplicateStage.SizeGrouping);
        foreach (var state in states) Validate(state);
        var eligible = states.Where(state => !state.Failed && state.Entry.SizeBytes >= 0);
        if (reference is not null) eligible = reference.Failed ? [] : eligible.Where(state => state.Entry.SizeBytes == reference.Entry.SizeBytes);
        var candidates = eligible.GroupBy(state => state.Entry.SizeBytes!.Value).Where(group => group.Count() > 1).SelectMany(group => group).ToArray();
        sizeCandidates = candidates.LongLength;
        foreach (var state in candidates)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (state.Entry.SizeBytes == 0) state.QuickHash = state.FullHash = EmptyHash;
            else ReadCache(state);
        }
        var options = new ParallelOptions { MaxDegreeOfParallelism = workerCount, CancellationToken = cancellationToken };
        await Parallel.ForEachAsync(candidates.Where(state => state.Entry.SizeBytes > 0), options,
            (state, token) => Hash(state, false, token)).ConfigureAwait(false);
        var quickGroups = candidates.Where(state => !state.Failed).GroupBy(state => (state.Entry.SizeBytes, state.QuickHash))
            .Where(group => group.Count() > 1 && (reference is null || group.Contains(reference))).ToArray();
        var finalists = quickGroups.SelectMany(group => group).ToArray();
        quickCandidates = finalists.LongLength;
        Report(DuplicateStage.FullHash);
        await Parallel.ForEachAsync(finalists.Where(state => state.Entry.SizeBytes > FileContentHasher.SmallFileLimit), options,
            (state, token) => Hash(state, true, token)).ConfigureAwait(false);
        // A file can change after its quick/full/cache step while another candidate is still being read.
        foreach (var state in finalists) Validate(state);
        cancellationToken.ThrowIfCancellationRequested();
        var groups = finalists.Where(state => !state.Failed && state.FullHash is not null)
            .GroupBy(state => (state.Entry.SizeBytes, state.FullHash))
            .Where(group => group.Count() > 1 && (reference is null || group.Contains(reference)))
            .Select(group => new DuplicateGroup(0, group.Select(state => state.ToFile()).OrderBy(file => file.Entry.RelativePath, PathComparer).ToArray()))
            .OrderByDescending(group => group.RecoverableBytes).ThenBy(group => group.Files[0].Entry.AbsolutePath, PathComparer)
            .Select((group, index) => group with { GroupId = index + 1 }).ToArray();
        timer.Stop();
        Report(DuplicateStage.Completed);
        return new(groups, reference?.ToFile(), new(states.LongLength, sizeCandidates, quickCandidates, quickHashes, fullHashes,
            cacheHits, bytesRead, timer.Elapsed), errors.OrderBy(error => error.Path, PathComparer).ToArray(), skipped.OrderBy(error => error.Path, PathComparer).ToArray());
    }

    private static bool ValidHash(string? hash) => hash is { Length: 64 } && hash.All(Uri.IsHexDigit);
    private static bool IsFileError(Exception ex) => ex is IOException or UnauthorizedAccessException or
        ArgumentException or NotSupportedException or System.Security.SecurityException;
    private static bool IsCacheError(Exception ex) => IsFileError(ex) || ex is InvalidOperationException or System.Data.Common.DbException;

    private sealed class Candidate(FileEntry entry)
    {
        public FileEntry Entry { get; } = entry;
        public HashCacheKey Key { get; } = new(entry.AbsolutePath, entry.SizeBytes ?? 0, entry.ModifiedAt.ToUniversalTime().Ticks, FileContentHasher.Algorithm);
        public string? QuickHash { get; set; }
        public string? FullHash { get; set; }
        public bool Failed { get; set; }
        public bool CacheUsed { get; set; }
        public DuplicateFile ToFile() => new(Entry, Failed ? null : QuickHash, Failed ? null : FullHash);
    }
}
