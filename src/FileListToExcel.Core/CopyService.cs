using System.Text.Json;
using System.Text.Json.Serialization;

namespace FileListToExcel.Core;

public sealed class CopyService
{
    public Task<CopyResult> CopyAsync(CopyPlan plan, string destinationParent, IProgress<CopyProgress>? progress = null,
        CancellationToken token = default) => Task.Run(() => Copy(plan, destinationParent, progress, token));

    public Task<CopyResult> RecordPlanAsync(CopyPlan plan) => Task.Run(() =>
    {
        if (plan.EligibleFiles.Count != 0) throw new ArgumentException("복사 대상이 남아 있는 계획은 완료 기록만 할 수 없습니다.", nameof(plan));
        return Complete(null, plan.Outcomes, plan.Outcomes.Any(x => x.Status == CopyStatus.Cancelled));
    });

    /// <summary>Records an explicit cancellation of the destination dialog, without creating an output folder.</summary>
    public Task<CopyResult> RecordCancelledAsync(CopyPlan plan) => Task.Run(() =>
    {
        var outcomes = plan.Outcomes.Concat(plan.EligibleFiles.Select(item => Outcome(item, null,
            CopyStatus.Cancelled, "Cancelled", "목적지 선택 또는 복사가 취소되었습니다."))).OrderBy(x => x.RowIndex).ToArray();
        return Complete(null, outcomes, true);
    });

    private static CopyResult Copy(CopyPlan plan, string destinationParent, IProgress<CopyProgress>? progress, CancellationToken token)
    {
        ArgumentNullException.ThrowIfNull(plan);
        var outcomes = plan.Outcomes.ToList();
        string? outputDirectory = null;
        int completed = 0;
        long transferred = 0;
        try
        {
            if (token.IsCancellationRequested)
                return Complete(null, outcomes.Concat(plan.EligibleFiles.Select(item => Outcome(item, null,
                    CopyStatus.Cancelled, "Cancelled", "복사가 시작되기 전에 취소되었습니다."))).OrderBy(x => x.RowIndex).ToArray(), true);
            if (plan.EligibleFiles.Count == 0) return Complete(null, outcomes, outcomes.Any(x => x.Status == CopyStatus.Cancelled));
            string parent = CollectPathPolicy.Validate(destinationParent, directory: true);
            using var parentGuard = CopyNative.LockDirectoryChain(parent);
            CopyNative.RequireMetadataSupport(parentGuard);
            outputDirectory = CopyNative.CreateOutputDirectory(parent, parentGuard);
            using var outputGuard = CopyNative.LockDirectoryChain(outputDirectory);
            foreach (var item in plan.EligibleFiles)
            {
                string destination = Path.Combine(outputDirectory, item.DestinationName);
                if (token.IsCancellationRequested)
                {
                    outcomes.Add(Outcome(item, null, CopyStatus.Cancelled, "Cancelled", "새 파일 복사를 시작하지 않고 취소했습니다."));
                    continue;
                }
                long fileTransferred = 0;
                try
                {
                    using var source = CopyNative.OpenSource(item.Source.AbsolutePath);
                    CopyNative.CopyFile(item, destination, source, outputGuard, bytes =>
                    {
                        fileTransferred = Math.Max(fileTransferred, bytes);
                        progress?.Report(new(item.DestinationName, completed, plan.EligibleFiles.Count,
                            SaturatingAdd(transferred, fileTransferred), plan.TotalBytes));
                    }, token);
                    outcomes.Add(Outcome(item, destination, CopyStatus.Copied, "Copied",
                        Path.GetFileName(item.Source.AbsolutePath) == item.DestinationName ? "복사 완료" : "이름 충돌을 구분하여 복사 완료"));
                }
                catch (OperationCanceledException)
                {
                    outcomes.Add(Outcome(item, destination, CopyStatus.Cancelled, "Cancelled", "진행 중 복사가 취소되었습니다. OS/네트워크 호출에서는 반영이 늦어질 수 있습니다."));
                }
                catch (Exception ex) when (CopyNative.IsFileError(ex))
                {
                    string code = CopyNative.ErrorCode(ex);
                    var status = code is "ChangedSinceList" or "SourceReplaced" or "CloudOnly" or "ReparsePoint" or "NotFile" ? CopyStatus.Excluded : CopyStatus.Failed;
                    outcomes.Add(Outcome(item, destination, status, code, ex.Message));
                }
                transferred = SaturatingAdd(transferred, fileTransferred);
                completed++;
                progress?.Report(new(item.DestinationName, completed, plan.EligibleFiles.Count, transferred, plan.TotalBytes));
            }
        }
        catch (Exception ex) when (CopyNative.IsFileError(ex))
        {
            var recorded = outcomes.Select(x => x.RowIndex).ToHashSet();
            foreach (var item in plan.EligibleFiles.Where(item => !recorded.Contains(item.RowIndex)))
                outcomes.Add(Outcome(item, null, CopyStatus.Failed, CopyNative.ErrorCode(ex), "목적지 준비 실패: " + ex.Message));
        }
        return Complete(outputDirectory, outcomes.OrderBy(x => x.RowIndex).ToArray(), token.IsCancellationRequested || outcomes.Any(x => x.Status == CopyStatus.Cancelled));
    }

    private static long SaturatingAdd(long a, long b) => b > long.MaxValue - a ? long.MaxValue : a + b;
    private static CopyOutcome Outcome(PlannedCopyFile item, string? destination, CopyStatus status, string code, string message)
        => new(item.RowIndex, item.Row.ItemId, item.Source.AbsolutePath, destination, status, code, message);

    private static CopyResult Complete(string? output, IReadOnlyList<CopyOutcome> outcomes, bool cancelled)
    {
        try
        {
            string report = CollectReportWriter.Write(output, outcomes, cancelled);
            return new(output, report, outcomes, cancelled);
        }
        catch (Exception ex) when (CopyNative.IsFileError(ex))
        {
            return new(output, null, outcomes, cancelled, "내부 결과 보고서를 기록하지 못했습니다: " + ex.Message);
        }
    }
}

/// <summary>Private local JSON report; never placed in the user's delivery folder.</summary>
internal static class CollectReportWriter
{
    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true, Converters = { new JsonStringEnumConverter() } };
    internal static string Write(string? output, IReadOnlyList<CopyOutcome> outcomes, bool cancelled)
    {
        string local = CopyNative.LocalAppDataPath();
        using var localGuard = CopyNative.LockDirectoryChain(local);
        string product = Path.Combine(local, "FileListToExcel");
        Directory.CreateDirectory(product);
        product = CopyNative.ResolveLocalReportDirectory(product, local);
        using var productGuard = CopyNative.LockDirectoryChain(product);
        string reports = Path.Combine(product, "Reports");
        Directory.CreateDirectory(reports);
        reports = CopyNative.ResolveLocalReportDirectory(reports, local);
        using var reportGuard = CopyNative.LockDirectoryChain(reports);
        string path = Path.Combine(reports, "collect_" + DateTime.UtcNow.ToString("yyyyMMdd_HHmmss") + "_" + Guid.NewGuid().ToString("N") + ".json");
        using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        JsonSerializer.Serialize(stream, new { product = "FileListToExcel", reportSchemaVersion = 1,
            createdUtc = DateTime.UtcNow, outputDirectory = output, cancelled, outcomes }, Options);
        stream.Flush(flushToDisk: true);
        return path;
    }
}
