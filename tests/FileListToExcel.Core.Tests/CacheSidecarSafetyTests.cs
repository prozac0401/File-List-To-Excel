using FileListToExcel.Core;
using Xunit;

namespace FileListToExcel.Core.Tests;

public sealed class CacheSidecarSafetyTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "FileListToExcel.CacheSidecarTests", Guid.NewGuid().ToString("N"));
    public CacheSidecarSafetyTests() => Directory.CreateDirectory(directory);
    public void Dispose() => Directory.Delete(directory, true);

    [Theory]
    [InlineData("-wal")]
    [InlineData("-shm")]
    [InlineData("-journal")]
    public void OfflineSidecarDisablesCacheWithoutWritingItsContents(string suffix)
    {
        string database = Path.Combine(directory, "hash_cache.sqlite");
        string sidecar = database + suffix;
        byte[] original = [0, 1, 2, 3, 4, 5, 250, 251, 252, 253];
        File.WriteAllBytes(sidecar, original);
        DateTime modified = File.GetLastWriteTimeUtc(sidecar);
        var attributes = File.GetAttributes(sidecar);
        try
        {
            File.SetAttributes(sidecar, attributes | FileAttributes.Offline);
            Assert.True(File.GetAttributes(sidecar).HasFlag(FileAttributes.Offline));
            using var cache = new SqliteHashCache(database);
            Assert.False(cache.IsAvailable);
            Assert.Contains("sidecar", cache.LastError);
            Assert.False(File.Exists(database));
            Assert.Null(cache.RecoveredDatabasePath);
            Assert.Equal(modified, File.GetLastWriteTimeUtc(sidecar));
        }
        finally { File.SetAttributes(sidecar, attributes); }
        Assert.Equal(original, File.ReadAllBytes(sidecar));
    }

    [Theory]
    [InlineData("-wal")]
    [InlineData("-shm")]
    [InlineData("-journal")]
    public void DirectoryAtSidecarPathDisablesCacheBeforeDatabaseCreation(string suffix)
    {
        string database = Path.Combine(directory, "hash_cache.sqlite");
        string sidecar = database + suffix;
        Directory.CreateDirectory(sidecar);
        string sentinel = Path.Combine(sidecar, "preserved.txt");
        File.WriteAllText(sentinel, "preserved sidecar directory");
        using var cache = new SqliteHashCache(database);
        Assert.False(cache.IsAvailable);
        Assert.Contains("sidecar", cache.LastError);
        Assert.False(File.Exists(database));
        Assert.Null(cache.RecoveredDatabasePath);
        Assert.Equal("preserved sidecar directory", File.ReadAllText(sentinel));
    }
}
