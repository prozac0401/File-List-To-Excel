using FileListToExcel.Core;
using Microsoft.Data.Sqlite;
using Xunit;
using Xunit.Abstractions;

namespace FileListToExcel.Core.Tests;

public sealed class HashCacheTests(ITestOutputHelper output) : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "FileListToExcel.CacheTests", Guid.NewGuid().ToString("N"));
    private string CachePath => Path.Combine(directory, "cache", "hash_cache.sqlite");
    public void Dispose() { try { if (Directory.Exists(directory)) Directory.Delete(directory, true); } catch (IOException) { } }
    private string FileAt(string name, byte[] contents)
    {
        string path = Path.Combine(directory, name);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, contents);
        return path;
    }
    private static FileEntry[] Entries(params string[] paths) => new FileScanner().Enumerate(new(paths, ScanMode.SelectedItems))
        .Select(item => item.Entry ?? throw new IOException(item.Error!.Message)).ToArray();
    private static HashCacheKey Key(FileEntry entry) => new(entry.AbsolutePath, entry.SizeBytes!.Value, entry.ModifiedAt.ToUniversalTime().Ticks, FileContentHasher.Algorithm);
    private static byte[] Data(int size, int seed)
    {
        var bytes = new byte[size]; new Random(seed).NextBytes(bytes); return bytes;
    }
    private void Stats(string run, DuplicateResult result) => output.WriteLine($"{run}: files={result.Stats.ScannedFiles}; size candidates={result.Stats.SizeCandidates}; quick candidates={result.Stats.QuickCandidates}; quick hashes={result.Stats.QuickHashCount}; full hashes={result.Stats.FullHashCount}; cache hits={result.Stats.CacheHitCount}; bytes read={result.Stats.BytesRead}; elapsed={result.Stats.Elapsed.TotalMilliseconds:F2}ms; groups={result.Groups.Count}");
    private static readonly string HashA = new('A', 64);
    private static readonly string HashB = new('B', 64);

    [Fact]
    public async Task SecondScanPersistsAndReusesEveryCandidateWithoutReadingContents()
    {
        byte[] bytes = Data(2 * 1024 * 1024, 101);
        string[] paths = [FileAt("원본 한국어 😀.bin", bytes), FileAt("다른이름", bytes), FileAt("distinct", Data(bytes.Length, 102))];
        DuplicateResult first;
        using (var cache = new SqliteHashCache(CachePath))
        {
            Assert.True(cache.IsAvailable, cache.LastError);
            first = await new DuplicateDetector(cache).DetectAsync(Entries(paths));
        }
        using var reopened = new SqliteHashCache(CachePath);
        var second = await new DuplicateDetector(reopened).DetectAsync(Entries(paths));
        Assert.Single(first.Groups); Assert.Single(second.Groups);
        Assert.Equal(3, first.Stats.QuickHashCount); Assert.Equal(2, first.Stats.FullHashCount);
        Assert.Equal(3, second.Stats.CacheHitCount); Assert.Equal(0, second.Stats.QuickHashCount);
        Assert.Equal(0, second.Stats.FullHashCount); Assert.Equal(0, second.Stats.BytesRead);
        Assert.Equal(first.Groups[0].FullSha256, second.Groups[0].FullSha256);
        Stats("cold", first); Stats("warm", second);
    }

    [Fact]
    public async Task SameSizeModificationInvalidatesOnlyTheChangedFile()
    {
        string first = FileAt("first", [1, 2]), second = FileAt("copy", [1, 2]);
        using var cache = new SqliteHashCache(CachePath);
        await new DuplicateDetector(cache).DetectAsync(Entries(first, second));
        DateTime previous = File.GetLastWriteTimeUtc(first);
        File.WriteAllBytes(first, [2, 1]); File.SetLastWriteTimeUtc(first, previous.AddSeconds(2));
        var result = await new DuplicateDetector(cache).DetectAsync(Entries(first, second));
        Assert.Empty(result.Groups); Assert.Equal(1, result.Stats.FullHashCount); Assert.Equal(1, result.Stats.CacheHitCount);
        Stats("same-size-modified", result);
    }

    [Fact]
    public async Task SizeModificationCannotReuseAnOldHash()
    {
        string first = FileAt("first", [1, 2]), second = FileAt("copy", [1, 2]);
        using var cache = new SqliteHashCache(CachePath);
        var oldEntry = Entries(first)[0];
        await new DuplicateDetector(cache).DetectAsync(Entries(first, second));
        File.WriteAllBytes(first, [1, 2, 3]); File.SetLastWriteTimeUtc(first, oldEntry.ModifiedAt.ToUniversalTime());
        var newEntry = Entries(first)[0]; Assert.False(cache.TryGet(Key(newEntry), out _));
        string third = FileAt("newCopy", [1, 2, 3]);
        var result = await new DuplicateDetector(cache).DetectAsync(Entries(first, second, third));
        Assert.Equal(3, Assert.Single(result.Groups).SizeBytes); Assert.Equal(2, result.Stats.FullHashCount);
    }

    [Fact]
    public async Task DeletedDatabaseIsRecreatedWithoutChangingResults()
    {
        string first = FileAt("first", [1, 2]), second = FileAt("copy", [1, 2]);
        using (var cache = new SqliteHashCache(CachePath)) await new DuplicateDetector(cache).DetectAsync(Entries(first, second));
        File.Delete(CachePath);
        using var recreated = new SqliteHashCache(CachePath);
        var result = await new DuplicateDetector(recreated).DetectAsync(Entries(first, second));
        Assert.True(recreated.IsAvailable, recreated.LastError); Assert.Single(result.Groups);
        Assert.Equal(0, result.Stats.CacheHitCount); Assert.Equal(2, result.Stats.FullHashCount);
    }

    [Fact]
    public async Task CorruptDatabaseIsPreservedAndRebuiltSafely()
    {
        FileAt("cache/hash_cache.sqlite", Data(1024, 32));
        using var cache = new SqliteHashCache(CachePath);
        Assert.True(cache.IsAvailable, cache.LastError); Assert.NotNull(cache.RecoveredDatabasePath);
        Assert.True(File.Exists(cache.RecoveredDatabasePath));
        var result = await new DuplicateDetector(cache).DetectAsync(Entries(FileAt("first", [1, 2]), FileAt("copy", [1, 2])));
        Assert.Single(result.Groups); Assert.Empty(result.Errors);
        using var verify = Open(CachePath);
        using var command = verify.CreateCommand(); command.CommandText = "PRAGMA integrity_check";
        Assert.Equal("ok", command.ExecuteScalar());
        command.CommandText = "PRAGMA journal_mode"; Assert.Equal("wal", command.ExecuteScalar());
    }

    [Fact]
    public void InvalidSchemaIsQuarantinedAndRecreated()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(CachePath)!);
        using (var database = Open(CachePath))
        {
            using var command = database.CreateCommand(); command.CommandText = "CREATE TABLE hashes (broken TEXT)"; command.ExecuteNonQuery();
        }
        using var cache = new SqliteHashCache(CachePath);
        Assert.True(cache.IsAvailable, cache.LastError); Assert.NotNull(cache.RecoveredDatabasePath);
        var key = new HashCacheKey(Path.Combine(directory, "source"), 2, 123, FileContentHasher.Algorithm);
        cache.Store(key, new(HashA, HashB, DateTime.UtcNow)); Assert.True(cache.TryGet(key, out _));
    }

    [Fact]
    public async Task UnavailableCacheDirectoryDoesNotFailTheScan()
    {
        string blockedParent = FileAt("blocked", [1]);
        using var cache = new SqliteHashCache(Path.Combine(blockedParent, "hash_cache.sqlite"));
        Assert.False(cache.IsAvailable); Assert.NotNull(cache.LastError);
        var result = await new DuplicateDetector(cache).DetectAsync(Entries(FileAt("first", [1, 2]), FileAt("copy", [1, 2])));
        Assert.Single(result.Groups); Assert.Empty(result.Errors); Assert.Equal(2, result.Stats.FullHashCount);
    }

    [Fact]
    public async Task ConcurrentWorkersStoreAtomicRowsAndPreserveFullHashOnQuickUpdate()
    {
        using var cache = new SqliteHashCache(CachePath);
        await Parallel.ForEachAsync(Enumerable.Range(0, 100), async (index, token) =>
        {
            await Task.Yield();
            var key = new HashCacheKey(Path.Combine(directory, $"file{index}"), index, index + 1000, FileContentHasher.Algorithm);
            cache.Store(key, new(HashA, HashB, DateTime.UtcNow));
            cache.Store(key, new(HashA, null, DateTime.UtcNow));
            Assert.True(cache.TryGet(key, out var stored)); Assert.Equal(HashB, stored!.FullSha256);
        });
        using var verify = Open(CachePath);
        using var command = verify.CreateCommand(); command.CommandText = "SELECT COUNT(*) FROM hashes";
        Assert.Equal(100L, command.ExecuteScalar());
    }

    [Fact]
    public void NewFileVersionDoesNotInheritPreviousFullHashAndPathCaseIsNormalized()
    {
        using var cache = new SqliteHashCache(CachePath);
        var key = new HashCacheKey(Path.Combine(directory, "한글 NAME"), 2, 1000, FileContentHasher.Algorithm);
        cache.Store(key, new(HashA.ToLowerInvariant(), HashB.ToLowerInvariant(), DateTime.UtcNow));
        var alternate = OperatingSystem.IsWindows() ? key with { AbsolutePath = key.AbsolutePath.ToUpperInvariant() } : key;
        Assert.True(cache.TryGet(alternate, out var same)); Assert.Equal(HashA, same!.QuickHash); Assert.Equal(HashB, same.FullSha256);
        var changed = key with { ModifiedUtcTicks = 2000 };
        cache.Store(changed, new(HashA, null, DateTime.UtcNow));
        Assert.False(cache.TryGet(key, out _)); Assert.True(cache.TryGet(changed, out var next)); Assert.Null(next!.FullSha256);
    }

    [Fact]
    public void InvalidHashesAreIgnoredAndFutureSchemaIsPreserved()
    {
        using (var cache = new SqliteHashCache(CachePath))
        {
            var key = new HashCacheKey(Path.Combine(directory, "source"), 1, 1, FileContentHasher.Algorithm);
            cache.Store(key, new("bogus", new string('Z', 64), DateTime.UtcNow));
            Assert.False(cache.TryGet(key, out _));
        }
        using (var database = Open(CachePath))
        {
            using var command = database.CreateCommand(); command.CommandText = "PRAGMA user_version=999"; command.ExecuteNonQuery();
        }
        using var future = new SqliteHashCache(CachePath);
        Assert.False(future.IsAvailable); Assert.Null(future.RecoveredDatabasePath); Assert.NotNull(future.LastError);
    }

    [Fact]
    public async Task CancellationKeepsEarlierCommittedCacheRowsValid()
    {
        string first = FileAt("first", [1, 2]), second = FileAt("copy", [1, 2]);
        using (var cache = new SqliteHashCache(CachePath))
        {
            await new DuplicateDetector(cache).DetectAsync(Entries(first, second));
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new DuplicateDetector(cache).DetectAsync(Entries(first, second), cancellationToken: cancellation.Token));
        }
        using var reopened = new SqliteHashCache(CachePath);
        var result = await new DuplicateDetector(reopened).DetectAsync(Entries(first, second));
        Assert.Single(result.Groups); Assert.Equal(2, result.Stats.CacheHitCount); Assert.Equal(0, result.Stats.BytesRead);
    }

    [Fact]
    public void CacheOwnsOnlyItsDatabaseCompanionsAndRecoveryArtifacts()
    {
        using var cache = new SqliteHashCache(CachePath);
        Assert.True(cache.OwnsPath(CachePath)); Assert.True(cache.OwnsPath(CachePath + "-wal"));
        Assert.True(cache.OwnsPath(CachePath + "-shm")); Assert.True(cache.OwnsPath(CachePath + "-journal"));
        Assert.True(cache.OwnsPath(CachePath + ".corrupt-20260924"));
        Assert.False(cache.OwnsPath(CachePath + ".xlsx")); Assert.False(cache.OwnsPath(Path.Combine(directory, "different", "hash_cache.sqlite")));
    }

    [Fact]
    public void CacheInsideDirectoryLinkIsRejectedBeforeWriting()
    {
        string target = Path.Combine(directory, "target"), link = Path.Combine(directory, "link");
        Directory.CreateDirectory(target);
        string original = FileAt("target/hash_cache.sqlite", [1, 2, 3, 4]);
        if (OperatingSystem.IsWindows())
        {
            var start = new System.Diagnostics.ProcessStartInfo("cmd.exe") { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
            foreach (string argument in new[] { "/c", "mklink", "/J", link, target }) start.ArgumentList.Add(argument);
            using var process = System.Diagnostics.Process.Start(start)!; process.WaitForExit();
            Assert.True(process.ExitCode == 0, process.StandardError.ReadToEnd());
        }
        else Directory.CreateSymbolicLink(link, target);
        try
        {
            using var cache = new SqliteHashCache(Path.Combine(link, "hash_cache.sqlite"));
            Assert.False(cache.IsAvailable); Assert.Null(cache.RecoveredDatabasePath);
            Assert.Equal(new byte[] { 1, 2, 3, 4 }, File.ReadAllBytes(original));
        }
        finally { Directory.Delete(link); }
    }

    [Fact]
    public async Task LockedCacheFallsBackWithoutQuarantiningTheDatabase()
    {
        using (var initialized = new SqliteHashCache(CachePath)) Assert.True(initialized.IsAvailable, initialized.LastError);
        using var blocker = Open(CachePath);
        using var transaction = blocker.BeginTransaction();
        using var cache = new SqliteHashCache(CachePath);
        Assert.Null(cache.RecoveredDatabasePath);
        var result = await new DuplicateDetector(cache).DetectAsync(Entries(FileAt("first", [1, 2]), FileAt("copy", [1, 2])));
        Assert.Single(result.Groups); Assert.Empty(result.Errors);
        Assert.False(cache.IsAvailable); Assert.NotNull(cache.LastError);
    }

    [Fact]
    public async Task ChangedDuringHashDoesNotPoisonCache()
    {
        string first = FileAt("first", [1, 2]), second = FileAt("copy", [1, 2]);
        var entries = Entries(first, second);
        using var cache = new SqliteHashCache(CachePath);
        var result = await new DuplicateDetector(cache, new ChangingHasher(first)).DetectAsync(entries);
        Assert.Empty(result.Groups); Assert.Equal("ChangedDuringScan", Assert.Single(result.Errors).ErrorType);
        Assert.False(cache.TryGet(Key(entries[0]), out _));
        Assert.True(cache.TryGet(Key(entries[1]), out _));
    }

    private sealed class ChangingHasher(string changedPath) : IFileContentHasher
    {
        private readonly FileContentHasher inner = new();
        public Task<ContentHash> QuickAsync(string path, long size, CancellationToken token) => inner.QuickAsync(path, size, token);
        public async Task<ContentHash> FullAsync(string path, CancellationToken token)
        {
            var result = await inner.FullAsync(path, token);
            if (path == changedPath) File.SetLastWriteTimeUtc(path, DateTime.UtcNow.AddMinutes(-3));
            return result;
        }
    }

    private static SqliteConnection Open(string path)
    {
        var connection = new SqliteConnection(new SqliteConnectionStringBuilder { DataSource = path, Pooling = false }.ToString());
        connection.Open(); return connection;
    }
}
