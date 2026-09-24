using System.Collections.Concurrent;
using System.Security.Cryptography;
using FileListToExcel.Core;
using Xunit;

namespace FileListToExcel.Core.Tests;

public sealed class DuplicateSafetyRegressionTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "FileListToExcel.DuplicateSafetyTests", Guid.NewGuid().ToString("N"));
    public DuplicateSafetyRegressionTests() => Directory.CreateDirectory(directory);
    public void Dispose() { try { Directory.Delete(directory, true); } catch (IOException) { } }
    private string FileAt(string relative, byte[] contents)
    {
        string path = Path.Combine(directory, relative);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, contents);
        return path;
    }
    private static FileEntry[] Entries(params string[] paths) => new FileScanner().Enumerate(new(paths, ScanMode.SelectedItems))
        .Select(item => item.Entry ?? throw new IOException(item.Error!.Message)).ToArray();
    private static HashCacheKey Key(FileEntry entry) => new(entry.AbsolutePath, entry.SizeBytes!.Value,
        entry.ModifiedAt.ToUniversalTime().Ticks, FileContentHasher.Algorithm);

    [Fact]
    public async Task LockedReferenceCannotTurnIntoUnrelatedDuplicateResults()
    {
        string reference = FileAt("reference", [1, 2, 3, 4]);
        FileAt("unrelated-one", [9, 9, 9, 9]); FileAt("unrelated-two", [9, 9, 9, 9]);
        using var held = new FileStream(reference, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        var result = await new DuplicateScanService().ScanAsync(new([reference], DuplicateMode.Matches));
        Assert.Empty(result.Groups);
        Assert.Contains(result.Errors, error => error.Path == reference);
        Assert.Null(result.ReferenceFile!.FullSha256);
    }

    [Fact]
    public async Task OfflineReferenceCannotTurnIntoUnrelatedDuplicateResults()
    {
        string reference = FileAt("reference", [1, 2, 3, 4]);
        FileAt("unrelated-one", [9, 9, 9, 9]); FileAt("unrelated-two", [9, 9, 9, 9]);
        var original = File.GetAttributes(reference);
        try
        {
            File.SetAttributes(reference, original | FileAttributes.Offline);
            var result = await new DuplicateScanService().ScanAsync(new([reference], DuplicateMode.Matches));
            Assert.Empty(result.Groups);
            Assert.Contains(result.Skipped, error => error.Path == reference && error.ErrorType == "CloudOnly");
            Assert.Equal(0, result.Stats.BytesRead);
            Assert.Equal(0, result.Stats.FullHashCount);
        }
        finally { File.SetAttributes(reference, original); }
    }

    [Fact]
    public async Task ReorderingOverlappingRootsKeepsDistinctFilesAndDuplicateTotals()
    {
        FileAt("A/one", [1, 2]); FileAt("A/deep/two", [1, 2]); FileAt("B/three", [1, 2]);
        FileAt("B/unique", [1, 2, 3]);
        var roots = new[] { Path.Combine(directory, "A"), directory, Path.Combine(directory, "B"), Path.Combine(directory, "A", "deep") };
        var first = await new DuplicateScanService().ScanAsync(new(roots, DuplicateMode.Folders));
        var second = await new DuplicateScanService().ScanAsync(new(roots.Reverse().ToArray(), DuplicateMode.Folders));
        Assert.Equal(4, first.Stats.ScannedFiles); Assert.Equal(first.Stats.ScannedFiles, second.Stats.ScannedFiles);
        Assert.Equal(3, first.DuplicateFileCount); Assert.Equal(first.DuplicateFileCount, second.DuplicateFileCount);
        Assert.Equal(4, first.RecoverableBytes); Assert.Equal(first.RecoverableBytes, second.RecoverableBytes);
        Assert.Equal(Assert.Single(first.Groups).Files.Select(file => file.Entry.AbsolutePath).Order(),
            Assert.Single(second.Groups).Files.Select(file => file.Entry.AbsolutePath).Order());
        Assert.Empty(first.Errors); Assert.Empty(second.Errors);
    }

    [Fact]
    public async Task CachedFileChangedWhileAnotherCandidateHashesIsRejected()
    {
        byte[] contents = [1, 2, 3, 4];
        string cachedPath = FileAt("cached", contents), trigger = FileAt("trigger", contents), copy = FileAt("copy", contents);
        var entries = Entries(cachedPath, trigger, copy);
        var cache = new MemoryCache(cachedPath);
        string hash = Convert.ToHexString(SHA256.HashData(contents));
        cache.Values[Key(entries[0])] = new(hash, hash, DateTime.UtcNow);
        var hasher = new MutatingHasher(trigger, () =>
        {
            Assert.True(Volatile.Read(ref cache.CompletedCachedStep));
            File.WriteAllBytes(cachedPath, [8, 8, 8, 8, 8]);
        });
        var result = await new DuplicateDetector(cache, hasher).DetectAsync(entries);
        Assert.Equal(2, Assert.Single(result.Groups).Files.Count);
        Assert.DoesNotContain(result.Groups.SelectMany(group => group.Files), file => file.Entry.AbsolutePath == cachedPath);
        Assert.Contains(result.Errors, error => error.Path == cachedPath && error.ErrorType == "ChangedDuringScan");
        Assert.Equal(1, result.Stats.CacheHitCount);
        Assert.False(cache.TryGet(Key(Entries(cachedPath)[0]), out _));
    }

    [Fact]
    public async Task EmptyFileChangedWhileOtherCandidatesHashIsRejected()
    {
        string changed = FileAt("empty-changed", []), empty = FileAt("empty-stable", []);
        string trigger = FileAt("trigger", [1, 2]), copy = FileAt("copy", [1, 2]);
        var entries = Entries(changed, empty, trigger, copy);
        var result = await new DuplicateDetector(hasher: new MutatingHasher(trigger, () => File.WriteAllBytes(changed, [3])))
            .DetectAsync(entries);
        var group = Assert.Single(result.Groups);
        Assert.Equal(2, group.SizeBytes);
        Assert.DoesNotContain(group.Files, file => file.Entry.AbsolutePath == changed || file.Entry.AbsolutePath == empty);
        Assert.Contains(result.Errors, error => error.Path == changed && error.ErrorType == "ChangedDuringScan");
    }

    [Fact]
    public async Task CancellationInsideRealFullHashDoesNotCacheAnIncompleteDigest()
    {
        string first = Path.Combine(directory, "large-first.bin"), second = Path.Combine(directory, "large-second.bin");
        foreach (string path in new[] { first, second })
            using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write)) file.SetLength(128L * 1024 * 1024);
        var entries = Entries(first, second);
        string database = Path.Combine(directory, "hash_cache.sqlite");
        var fullStarted = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var cancellation = new CancellationTokenSource();
        string interrupted;
        using (var cache = new SqliteHashCache(database))
        {
            Assert.True(cache.IsAvailable, cache.LastError);
            var progress = new CallbackProgress(value =>
            {
                if (value.Stage == DuplicateStage.FullHash && value.CurrentPath.Length > 0) fullStarted.TrySetResult(value.CurrentPath);
            });
            var work = Task.Run(() => new DuplicateDetector(cache).DetectAsync(entries, progress: progress, cancellationToken: cancellation.Token));
            interrupted = await fullStarted.Task.WaitAsync(TimeSpan.FromSeconds(10));
            await Task.Delay(1);
            bool readHandleObserved = false;
            for (int attempt = 0; attempt < 500 && !work.IsCompleted; attempt++)
            {
                try { using var probe = new FileStream(interrupted, FileMode.Open, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete); }
                catch (IOException) { readHandleObserved = true; break; }
                await Task.Delay(1);
            }
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => work);
            Assert.True(readHandleObserved, "Cancellation must occur after the full-read handle is open.");
            Assert.True(cache.TryGet(Key(entries.Single(entry => entry.AbsolutePath == interrupted)), out var saved));
            Assert.NotNull(saved!.QuickHash); Assert.Null(saved.FullSha256);
            foreach (string path in new[] { first, second })
            {
                using var released = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
                Assert.Equal(128L * 1024 * 1024, released.Length);
            }
        }
        using var reopened = new SqliteHashCache(database);
        Assert.True(reopened.IsAvailable, reopened.LastError);
        var resumed = await new DuplicateDetector(reopened).DetectAsync(entries);
        Assert.Equal(2, Assert.Single(resumed.Groups).Files.Count);
        Assert.True(resumed.Stats.FullHashCount >= 1);
        Assert.Empty(resumed.Errors);
    }

    private sealed class MemoryCache(string observedPath) : IDuplicateHashCache
    {
        public ConcurrentDictionary<HashCacheKey, HashCacheValue> Values { get; } = new();
        public bool CompletedCachedStep;
        public bool TryGet(HashCacheKey key, out HashCacheValue? value) => Values.TryGetValue(key, out value);
        public void Store(HashCacheKey key, HashCacheValue value)
        {
            Values[key] = value;
            if (key.AbsolutePath == observedPath) Volatile.Write(ref CompletedCachedStep, true);
        }
    }
    private sealed class MutatingHasher(string trigger, Action change) : IFileContentHasher
    {
        private readonly FileContentHasher inner = new();
        private int changed;
        public Task<ContentHash> QuickAsync(string path, long size, CancellationToken token) => inner.QuickAsync(path, size, token);
        public Task<ContentHash> FullAsync(string path, CancellationToken token)
        {
            if (path == trigger && Interlocked.Exchange(ref changed, 1) == 0) change();
            return inner.FullAsync(path, token);
        }
    }
    private sealed class CallbackProgress(Action<DuplicateProgress> callback) : IProgress<DuplicateProgress>
    {
        public void Report(DuplicateProgress value) => callback(value);
    }
}
