using System.Diagnostics;

namespace FileListToExcel.Core;

public enum DuplicateMode { Matches, Folders, SelectedFiles }
public sealed record DuplicateRequest(IReadOnlyList<string> Paths, DuplicateMode Mode);

/// <summary>Resolves Explorer selections and reuses the metadata scanner with the stricter duplicate policy.</summary>
public sealed class DuplicateScanService(IDuplicateHashCache? cache = null)
{
    public async Task<DuplicateResult> ScanAsync(DuplicateRequest request,
        IProgress<DuplicateProgress>? progress = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(request.Paths);
        if (!Enum.IsDefined(request.Mode) || request.Paths.Count == 0)
            throw new ArgumentException("중복 검사 대상이 없습니다.", nameof(request));
        if (request.Mode == DuplicateMode.Matches && request.Paths.Count != 1)
            throw new ArgumentException("같은 파일 찾기는 파일 하나를 선택하세요.", nameof(request));
        var timer = Stopwatch.StartNew();
        var errors = new List<ScanError>();
        var skipped = new List<ScanError>();
        var files = new Dictionary<string, FileEntry>(OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);
        var roots = new List<string>();
        string? referencePath = null;
        foreach (var path in request.Paths)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                var fullPath = Path.TrimEndingDirectorySeparator(Path.GetFullPath(path));
                var attributes = File.GetAttributes(fullPath);
                bool directory = (attributes & FileAttributes.Directory) != 0;
                if (directory != (request.Mode == DuplicateMode.Folders))
                {
                    errors.Add(new(fullPath, "InvalidSelection", request.Mode == DuplicateMode.Folders ? "폴더를 선택하세요." : "파일을 선택하세요."));
                    continue;
                }
                var checkPath = directory ? fullPath : Path.GetDirectoryName(fullPath)!;
                var reason = DuplicatePathPolicy.CheckDirectoryChain(checkPath);
                if (reason is not null) { AddIssue(reason); continue; }
                if (request.Mode == DuplicateMode.Matches)
                {
                    referencePath = fullPath;
                    roots.Add(Path.GetDirectoryName(fullPath)!);
                }
                else roots.Add(fullPath);
            }
            catch (Exception ex) when (DuplicatePathPolicy.IsFileError(ex)) { errors.Add(DuplicatePathPolicy.Error(path, ex)); }
        }
        var scanMode = request.Mode == DuplicateMode.SelectedFiles ? ScanMode.SelectedItems : ScanMode.Recursive;
        var scannerProgress = new CallbackProgress(value => progress?.Report(new(DuplicateStage.FindingFiles,
            files.Count, 0, 0, 0, value.CurrentPath)));
        foreach (var item in new FileScanner().Enumerate(new(roots, scanMode, SkipAllReparsePoints: true), scannerProgress, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (item.Error is { } error) AddIssue(error);
            if (item.Entry is { IsDirectory: false } entry)
            {
                if (cache is SqliteHashCache diskCache && diskCache.OwnsPath(entry.AbsolutePath))
                {
                    skipped.Add(new(entry.AbsolutePath, "InternalCache", "검사 도구의 해시 캐시 파일을 제외했습니다."));
                    continue;
                }
                files.TryAdd(entry.AbsolutePath, entry);
            }
        }
        // A missing/inaccessible reference must not turn a matches request into a folder duplicate scan.
        if (request.Mode == DuplicateMode.Matches && (referencePath is null || !files.ContainsKey(referencePath)))
        {
            if (referencePath is not null && !errors.Any(error => error.Path == referencePath))
                errors.Add(new(referencePath, "ReferenceUnavailable", "기준 파일을 검사할 수 없습니다."));
            return new([], null, new(files.Count, 0, 0, 0, 0, 0, 0, timer.Elapsed), errors, skipped);
        }
        int workers = roots.Any(path => path.StartsWith(@"\\", StringComparison.Ordinal)) ? 1 : 2;
        var result = await new DuplicateDetector(cache, workerCount: workers).DetectAsync(files.Values.ToArray(), referencePath, progress, cancellationToken).ConfigureAwait(false);
        timer.Stop();
        return result with
        {
            Stats = result.Stats with { Elapsed = timer.Elapsed },
            Errors = errors.Concat(result.Errors).ToArray(),
            Skipped = skipped.Concat(result.Skipped).ToArray()
        };

        void AddIssue(ScanError issue)
        {
            if (issue.ErrorType is "ReparsePoint" or "CloudOnly" or "LinkNotTraversed") skipped.Add(issue);
            else errors.Add(issue);
        }
    }

    private sealed class CallbackProgress(Action<ScanProgress> callback) : IProgress<ScanProgress>
    {
        public void Report(ScanProgress value) => callback(value);
    }
}

internal static class DuplicatePathPolicy
{
    public static ScanError? CheckDirectoryChain(string path)
    {
        try
        {
            for (string? current = path; current is not null; current = Path.GetDirectoryName(Path.TrimEndingDirectorySeparator(current)))
            {
                var attributes = File.GetAttributes(current);
                if ((attributes & (FileAttributes.Offline | LocalFileContentPolicy.RecallOnOpen | LocalFileContentPolicy.RecallOnDataAccess)) != 0)
                    return new(path, "CloudOnly", "건너뜀: Cloud-only 폴더를 열거하지 않습니다.");
                if ((attributes & FileAttributes.ReparsePoint) != 0)
                    return new(path, "ReparsePoint", "건너뜀: Junction, Symbolic Link 및 기타 Reparse Point 폴더를 따라가지 않습니다.");
            }
            return null;
        }
        catch (Exception ex) when (IsFileError(ex)) { return Error(path, ex); }
    }

    public static bool IsFileError(Exception ex) => ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException or System.Security.SecurityException;
    public static ScanError Error(string path, Exception ex) => new(path, ex is UnauthorizedAccessException ? "AccessDenied" : ex.GetType().Name, ex.Message);
}
