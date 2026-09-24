using FileListToExcel.Core;
using Xunit;

namespace FileListToExcel.Core.Tests;

public sealed class DuplicateCacheIntegrationTests
{
    [Fact]
    public async Task CacheInsideSelectedTreeIsNotTreatedAsSourceData()
    {
        var directory = Path.Combine(Path.GetTempPath(), "FileListToExcel.CacheIntegration", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            File.WriteAllText(Path.Combine(directory, "first"), "same content");
            File.WriteAllText(Path.Combine(directory, "second"), "same content");
            using var cache = new SqliteHashCache(Path.Combine(directory, "hash_cache.sqlite"));
            var service = new DuplicateScanService(cache);
            var cold = await service.ScanAsync(new([directory], DuplicateMode.Folders));
            var warm = await service.ScanAsync(new([directory], DuplicateMode.Folders));
            Assert.Equal(2, cold.Stats.ScannedFiles); Assert.Equal(2, warm.Stats.ScannedFiles);
            Assert.Single(warm.Groups); Assert.Equal(0, warm.Stats.FullHashCount);
            Assert.Equal(2, warm.Stats.CacheHitCount); Assert.Equal(0, warm.Stats.BytesRead);
            Assert.Contains(warm.Skipped, item => item.ErrorType == "InternalCache");
        }
        finally { Directory.Delete(directory, true); }
    }
}
