using System.Diagnostics;
using System.Text.Json;
using FileListToExcel.Core;

if (!args.Contains("--run", StringComparer.Ordinal))
{
    Console.WriteLine("Opt-in: dotnet run --project tests/FileListToExcel.Benchmarks -c Release -- --run [--output result.json] [--fixture-root new-directory]");
    Console.WriteLine("Writes approximately 4.21 GiB of real test files and retains them for inspection. Never scans user files.");
    return 2;
}
string Option(string name, string fallback)
{
    int index = Array.IndexOf(args, name);
    return index >= 0 ? index + 1 < args.Length ? args[index + 1] : throw new ArgumentException(name + " requires a value.") : fallback;
}
string fixtureRoot = Path.GetFullPath(Option("--fixture-root", Path.Combine("artifacts", "benchmarks", Guid.NewGuid().ToString("N"))));
string outputPath = Path.GetFullPath(Option("--output", Path.Combine("artifacts", "duplicate-performance.json")));
if (Directory.Exists(fixtureRoot)) throw new IOException("The fixture root must be a new directory; existing data is never replaced.");
const int SmallCount = 5000, QuickCount = 100;
const long LargeSize = 2L * 1024 * 1024 * 1024 + 64 * 1024;
long expectedBytes = (long)SmallCount * (SmallCount + 1) / 2 + QuickCount * 2L * 1024 * 1024 + 2 * LargeSize;
var drive = new DriveInfo(Path.GetPathRoot(fixtureRoot)!);
if (drive.AvailableFreeSpace < expectedBytes + 512L * 1024 * 1024) throw new IOException("Insufficient space for the retained benchmark fixtures.");
Directory.CreateDirectory(fixtureRoot);
string source = Path.Combine(fixtureRoot, "source"), smallRoot = Path.Combine(source, "small"), quickRoot = Path.Combine(source, "quick"), largeRoot = Path.Combine(source, "large");
Directory.CreateDirectory(smallRoot); Directory.CreateDirectory(quickRoot); Directory.CreateDirectory(largeRoot);
var preparation = Stopwatch.StartNew();
Console.WriteLine($"Fixture: {fixtureRoot}");
Console.WriteLine($"Preparing {SmallCount + QuickCount + 2} real files; {expectedBytes:N0} bytes.");
for (int size = 1; size <= SmallCount; size++)
{
    byte[] bytes = new byte[size]; new Random(size).NextBytes(bytes);
    File.WriteAllBytes(Path.Combine(smallRoot, $"한글-{size:D5}.bin"), bytes);
    if (size % 1000 == 0) Console.WriteLine($"Prepared {size:N0}/{SmallCount:N0} small files.");
}
for (int index = 0; index < QuickCount; index++)
{
    byte[] bytes = new byte[2 * 1024 * 1024];
    int seed = index < 96 ? index : index < 98 ? 777 : 888;
    new Random(seed).NextBytes(bytes);
    if (index == 99) bytes[256 * 1024] ^= 0xff; // Same sampled fingerprint, different full SHA-256.
    File.WriteAllBytes(Path.Combine(quickRoot, $"quick-{index:D3}.bin"), bytes);
}
string largeA = Path.Combine(largeRoot, "two-GiB-original.bin"), largeB = Path.Combine(largeRoot, "two-GiB-copy.bin");
byte[] block = new byte[FileContentHasher.StreamBufferSize]; new Random(20260924).NextBytes(block);
await using (var first = new FileStream(largeA, FileMode.CreateNew, FileAccess.Write, FileShare.Read, 1, FileOptions.Asynchronous | FileOptions.SequentialScan))
await using (var second = new FileStream(largeB, FileMode.CreateNew, FileAccess.Write, FileShare.Read, 1, FileOptions.Asynchronous | FileOptions.SequentialScan))
{
    for (long written = 0; written < LargeSize; written += block.Length)
    {
        int length = (int)Math.Min(block.Length, LargeSize - written);
        await first.WriteAsync(block.AsMemory(0, length));
        await second.WriteAsync(block.AsMemory(0, length));
        if ((written + length) % (256L * 1024 * 1024) == 0) Console.WriteLine($"Prepared {written + length:N0} bytes in each large file.");
    }
    await first.FlushAsync(); await second.FlushAsync();
}
preparation.Stop();
var snapshots = Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories).Select(path => new FileInfo(path))
    .Select(info => new Snapshot(info.FullName, info.Length, info.LastWriteTimeUtc.Ticks, info.Attributes)).ToArray();
Require(snapshots.Length == 5102, "Fixture file count differs.");
Require(snapshots.Sum(item => item.Bytes) == expectedBytes, "Fixture byte count differs.");
Require(snapshots.All(item => (item.Attributes & (FileAttributes.SparseFile | FileAttributes.Compressed)) == 0), "Fixture must be materialized and uncompressed.");
Console.WriteLine($"Preparation finished in {preparation.Elapsed.TotalSeconds:F2}s. Starting cold scan.");
string cachePath = Path.Combine(fixtureRoot, "cache", "hash_cache.sqlite");
var cold = await MeasureScan("cold", source, cachePath);
Console.WriteLine("Starting persistent-cache warm scan.");
var warm = await MeasureScan("warm", source, cachePath);
Require(cold.Stats.ScannedFiles == 5102 && cold.Stats.SizeCandidates == 102 && cold.Stats.QuickCandidates == 6, "Candidate stages differ from controlled fixture.");
Require(cold.Stats.QuickHashCount == 102 && cold.Stats.FullHashCount == 6 && cold.Groups == 2, "Cold hash results differ from expected.");
Require(warm.Stats.CacheHitCount == 102 && warm.Stats.BytesRead == 0 && warm.Stats.FullHashCount == 0 && warm.Stats.QuickHashCount == 0 && warm.Groups == 2, "Warm scan failed to reuse persisted cache.");
Require(cold.MaximumActiveHashOperations == 2, "The two-worker bound was not observed.");
var scanCancellation = TestScanCancellation(source);
var hashCancellation = await TestLargeHashCancellation(largeA, largeB);
foreach (var snapshot in snapshots)
{
    var current = new FileInfo(snapshot.Path);
    Require(current.Length == snapshot.Bytes && current.LastWriteTimeUtc.Ticks == snapshot.ModifiedUtcTicks && current.Attributes == snapshot.Attributes, "Source metadata changed: " + snapshot.Path);
}
var report = new
{
    measuredAtUtc = DateTime.UtcNow, machine = Environment.MachineName,
    os = System.Runtime.InteropServices.RuntimeInformation.OSDescription,
    runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
    fixtureRoot, sourceRoot = source, cachePath, fileCount = snapshots.Length, totalBytes = expectedBytes,
    largeFileBytes = LargeSize, preparationMilliseconds = preparation.Elapsed.TotalMilliseconds,
    fixtureIsSparse = false, fixtureIsCompressed = false,
    workerLimit = 2, cold, warm, scanCancellation, hashCancellation, sourceMetadataUnchanged = true,
    notes = "Elapsed timings use this machine and include OS file-cache effects. Warm cache persisted across disposed/reopened SQLite connections. Peak working set is sampled every 10 ms. Fixtures are retained for UI checks. Full hash counts are attempted operations."
};
Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
await File.WriteAllTextAsync(outputPath, JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine($"Report: {outputPath}");
return 0;

static async Task<ScanMeasurement> MeasureScan(string label, string source, string cachePath)
{
    GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
    using var process = Process.GetCurrentProcess(); process.Refresh();
    long initialWorkingSet = process.WorkingSet64, peakWorkingSet = initialWorkingSet;
    long initialManagedBytes = GC.GetTotalMemory(false), initialAllocatedBytes = GC.GetTotalAllocatedBytes();
    var hasher = new TrackingHasher();
    var stages = new List<string>();
    using var monitorStop = new CancellationTokenSource();
    var monitor = Task.Run(async () =>
    {
        try
        {
            while (true)
            {
                monitorStop.Token.ThrowIfCancellationRequested();
                process.Refresh(); InterlockedMax(ref peakWorkingSet, process.WorkingSet64);
                await Task.Delay(10, monitorStop.Token);
            }
        }
        catch (OperationCanceledException) { }
    });
    var timer = Stopwatch.StartNew();
    DuplicateResult result;
    using (var cache = new SqliteHashCache(cachePath))
    {
        Require(cache.IsAvailable, cache.LastError ?? "Cache unavailable.");
        var scanItems = new FileScanner().Enumerate(new([source], ScanMode.Recursive, SkipAllReparsePoints: true)).ToArray();
        Require(scanItems.All(item => item.Error is null), "Unexpected fixture enumeration error.");
        var entries = scanItems.Where(item => item.Entry?.IsDirectory == false).Select(item => item.Entry!).ToArray();
        result = await new DuplicateDetector(cache, hasher).DetectAsync(entries, progress: new ProgressSink(value =>
        {
            string stage = value.Stage.ToString();
            lock (stages) if (stages.Count == 0 || stages[^1] != stage) stages.Add(stage);
        }));
    }
    timer.Stop(); monitorStop.Cancel(); await monitor;
    Require(result.Errors.Count == 0 && result.Skipped.Count == 0, "Unexpected scan issue.");
    long finalManagedBytes = GC.GetTotalMemory(false);
    var measurement = new ScanMeasurement(timer.Elapsed.TotalMilliseconds, result.Stats, result.Groups.Count,
        result.DuplicateFileCount, result.RecoverableBytes, initialWorkingSet, peakWorkingSet,
        Math.Max(0, peakWorkingSet - initialWorkingSet), initialManagedBytes, finalManagedBytes,
        GC.GetTotalAllocatedBytes() - initialAllocatedBytes, hasher.Maximum, stages.ToArray());
    Console.WriteLine($"{label}: {measurement.WallMilliseconds:F2}ms; candidates {result.Stats.SizeCandidates}->{result.Stats.QuickCandidates}; quick/full {result.Stats.QuickHashCount}/{result.Stats.FullHashCount}; cache {result.Stats.CacheHitCount}; read {result.Stats.BytesRead:N0}B; peak {peakWorkingSet:N0}B; groups {result.Groups.Count}");
    return measurement;
}

static object TestScanCancellation(string source)
{
    using var cancellation = new CancellationTokenSource();
    long reported = 0; bool cancelled = false;
    var timer = Stopwatch.StartNew();
    try
    {
        new FileScanner().Enumerate(new([source], ScanMode.Recursive, SkipAllReparsePoints: true),
            new ScanProgressSink(progress => { reported = progress.ItemCount; if (reported >= 128) cancellation.Cancel(); }), cancellation.Token).ToArray();
    }
    catch (OperationCanceledException) { cancelled = true; }
    timer.Stop(); Require(cancelled && reported < 5102, "Enumeration cancellation did not stop early.");
    return new { cancelled, reportedItems = reported, elapsedMilliseconds = timer.Elapsed.TotalMilliseconds };
}

static async Task<object> TestLargeHashCancellation(string first, string second)
{
    using var cancellation = new CancellationTokenSource();
    var hasher = new CancellationHasher(cancellation);
    var entries = new FileScanner().Enumerate(new([first, second], ScanMode.SelectedItems)).Select(item => item.Entry!).ToArray();
    bool cancelled = false, resultProduced = false;
    var timer = Stopwatch.StartNew();
    try { await new DuplicateDetector(hasher: hasher).DetectAsync(entries, cancellationToken: cancellation.Token); resultProduced = true; }
    catch (OperationCanceledException) { cancelled = true; }
    timer.Stop();
    double cancellationLatency = hasher.RequestedAt == 0 ? -1 : Stopwatch.GetElapsedTime(hasher.RequestedAt).TotalMilliseconds;
    bool handlesReleased;
    using (var a = new FileStream(first, FileMode.Open, FileAccess.Read, FileShare.None))
    using (var b = new FileStream(second, FileMode.Open, FileAccess.Read, FileShare.None)) handlesReleased = a.CanRead && b.CanRead;
    Require(cancelled && hasher.Started > 0 && !resultProduced && handlesReleased, "Large-file cancellation failed.");
    Console.WriteLine($"Cancel during {hasher.Started} large full streams: {cancellationLatency:F2}ms after request; handles released.");
    return new { cancelled, resultProduced, startedFullStreams = hasher.Started, handlesReleased, elapsedMilliseconds = timer.Elapsed.TotalMilliseconds, cancellationLatencyMilliseconds = cancellationLatency };
}

static void Require(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
static void InterlockedMax(ref long target, long value)
{
    long old;
    do { old = Interlocked.Read(ref target); if (value <= old) return; } while (Interlocked.CompareExchange(ref target, value, old) != old);
}

sealed record Snapshot(string Path, long Bytes, long ModifiedUtcTicks, FileAttributes Attributes);
sealed record ScanMeasurement(double WallMilliseconds, DuplicateStatistics Stats, int Groups, long DuplicateFiles, long RecoverableBytes,
    long InitialWorkingSetBytes, long PeakWorkingSetBytes, long WorkingSetGrowthBytes, long InitialManagedBytes, long FinalManagedBytes,
    long AllocatedBytes, int MaximumActiveHashOperations, string[] Stages);
sealed class ProgressSink(Action<DuplicateProgress> callback) : IProgress<DuplicateProgress> { public void Report(DuplicateProgress value) => callback(value); }
sealed class ScanProgressSink(Action<ScanProgress> callback) : IProgress<ScanProgress> { public void Report(ScanProgress value) => callback(value); }
sealed class TrackingHasher : IFileContentHasher
{
    private readonly FileContentHasher inner = new();
    private int active, maximum;
    public int Maximum => maximum;
    public Task<ContentHash> QuickAsync(string path, long size, CancellationToken token) => Track(() => inner.QuickAsync(path, size, token));
    public Task<ContentHash> FullAsync(string path, CancellationToken token) => Track(() => inner.FullAsync(path, token));
    private async Task<ContentHash> Track(Func<Task<ContentHash>> operation)
    {
        int count = Interlocked.Increment(ref active), observed;
        do { observed = maximum; } while (count > observed && Interlocked.CompareExchange(ref maximum, count, observed) != observed);
        try { return await operation(); } finally { Interlocked.Decrement(ref active); }
    }
}
sealed class CancellationHasher(CancellationTokenSource cancellation) : IFileContentHasher
{
    private readonly FileContentHasher inner = new();
    public int Started;
    public long RequestedAt;
    public Task<ContentHash> QuickAsync(string path, long size, CancellationToken token) => inner.QuickAsync(path, size, token);
    public async Task<ContentHash> FullAsync(string path, CancellationToken token)
    {
        if (Interlocked.Increment(ref Started) == 1)
            _ = Task.Run(async () => { await Task.Delay(50); Interlocked.Exchange(ref RequestedAt, Stopwatch.GetTimestamp()); cancellation.Cancel(); });
        return await inner.FullAsync(path, token);
    }
}
