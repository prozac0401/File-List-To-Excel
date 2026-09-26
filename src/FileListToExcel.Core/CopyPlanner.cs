using System.Text.RegularExpressions;

namespace FileListToExcel.Core;

public sealed class CopyPlanner
{
    public const int MaximumFiles = 10_000;

    /// <summary>Call only after the user's explicit copy command; this performs filesystem I/O.</summary>
    public CopyPlan BuildPlan(IReadOnlyList<CollectRow> rows, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(rows);
        if (rows.Count > MaximumFiles) throw new ArgumentException("한 번에 선택 가능한 행은 10,000개입니다.", nameof(rows));
        var files = new List<PlannedCopyFile>();
        var outcomes = new List<CopyOutcome>();
        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // A duplicated identity means a damaged table, not a duplicated path selection.
        foreach (var row in rows)
            if (row is null || !FileCollectContract.IsValidId(row.ItemId) || !ids.Add(row.ItemId))
                throw new ArgumentException("항목 ID가 없거나 중복되어 결과표를 사용할 수 없습니다.", nameof(rows));
        for (int index = 0; index < rows.Count; index++)
        {
            var row = rows[index];
            if (token.IsCancellationRequested)
            {
                outcomes.Add(new(index, row.ItemId, row.DisplayPath, null, CopyStatus.Cancelled, "Cancelled", "사전 검증이 취소되었습니다."));
                continue;
            }
            try
            {
                if (!FileCollectContract.TryParseSourceRecord(row.SourceRecord, out var source, out var error) || source is null)
                    throw new CollectSafetyException("InvalidRecord", error ?? "원본 정보가 손상되었습니다.");
                if (!string.Equals(row.ItemId, source.ItemId, StringComparison.OrdinalIgnoreCase))
                    throw new CollectSafetyException("ItemIdMismatch", "항목 ID와 원본 정보가 다릅니다. 새 목록을 만드세요.");
                if (source.Kind != "file")
                    throw new CollectSafetyException("NotFile", source.Reason ?? "폴더 및 복사 불가 항목은 제외합니다.");
                if (!string.Equals(row.DisplayPath, source.AbsolutePath, StringComparison.Ordinal))
                    throw new CollectSafetyException("PathMismatch", "전체경로와 원본 정보가 다릅니다. 새 목록을 만드세요.");
                string path = CollectPathPolicy.Validate(source.AbsolutePath);
                if (!paths.Add(path))
                {
                    outcomes.Add(new(index, row.ItemId, path, null, CopyStatus.Duplicate, "DuplicatePath", "같은 원본 경로를 한 번만 복사합니다."));
                    continue;
                }
                using var file = CopyNative.OpenSource(path);
                var stamp = file.ReadStamp();
                VerifyGeneratedStamp(source, stamp);
                string name = AllocateName(Path.GetFileName(path), names);
                files.Add(new(index, row, source, stamp, name));
            }
            catch (Exception ex) when (CopyNative.IsFileError(ex))
            {
                outcomes.Add(new(index, row.ItemId, row.DisplayPath, null, CopyStatus.Excluded, CopyNative.ErrorCode(ex), ex.Message));
            }
        }
        return new(rows.Count, files, outcomes);
    }

    internal static void VerifyGeneratedStamp(FileCollectSourceRecord source, CopyNative.FileStamp stamp)
    {
        if (source.SizeBytes != stamp.Length || source.ModifiedUtcTicks != stamp.ModifiedUtcTicks)
            throw new CollectSafetyException("ChangedSinceList", "목록 생성 후 변경됨: 파일 크기 또는 UTC 수정시간이 다릅니다.");
    }

    internal static string AllocateName(string name, HashSet<string> used)
    {
        if (used.Add(name)) return name;
        string extension = Path.GetExtension(name), stem = Path.GetFileNameWithoutExtension(name);
        if (stem.Length == 0) { stem = name; extension = ""; }
        for (int suffix = 2; ; suffix++)
        {
            string tag = $" ({suffix})";
            // Windows components are limited to 255 UTF-16 code units. Never split a surrogate pair.
            int keep = 255 - extension.Length - tag.Length;
            if (keep < 1) throw new CollectSafetyException("CollisionNameTooLong", "확장자가 너무 길어 충돌 이름을 안전하게 만들 수 없습니다.");
            string shortened = stem[..Math.Min(stem.Length, keep)];
            if (shortened.Length > 0 && char.IsHighSurrogate(shortened[^1])) shortened = shortened[..^1];
            string candidate = shortened + tag + extension;
            if (used.Add(candidate)) return candidate;
        }
    }
}

/// <summary>Only canonical DOS drive and UNC paths, with no aliases, devices or stream suffixes.</summary>
public static class CollectPathPolicy
{
    public static string Validate(string? path, bool directory = false)
    {
        if (string.IsNullOrWhiteSpace(path) || path.Length > 32_000 || path.Any(c => c < 32 || c is '/' or '"' or '<' or '>' or '|' or '?' or '*'))
            throw new ArgumentException("지원하지 않는 Windows 절대 경로입니다.");
        bool drive = path.Length >= 3 && char.IsAsciiLetter(path[0]) && path[1] == ':' && path[2] == '\\';
        bool unc = path.StartsWith(@"\\", StringComparison.Ordinal) && !path.StartsWith(@"\\.\", StringComparison.Ordinal);
        if (!drive && !unc) throw new ArgumentException("로컬 드라이브 또는 UNC 절대 경로만 지원합니다.");
        if (path[(drive ? 2 : 0)..].Contains(':')) throw new ArgumentException("대체 데이터 스트림을 원본 경로로 지정할 수 없습니다.");
        string tail = path[(drive ? 3 : 2)..];
        if (directory) tail = tail.TrimEnd('\\');
        var parts = tail.Split('\\');
        if (drive && directory && tail.Length == 0) return path[..3];
        if (unc && parts.Length < (directory ? 2 : 3)) throw new ArgumentException("UNC 서버와 공유 및 파일 경로가 필요합니다.");
        foreach (string part in parts)
        {
            if (part.Length is 0 or > 255 || part is "." or ".." || part.EndsWith(' ') || part.EndsWith('.'))
                throw new ArgumentException("빈 경로 요소, 상대 요소 또는 끝 공백/마침표는 지원하지 않습니다.");
            string stem = part.Split('.')[0];
            if (Regex.IsMatch(stem, @"^(CON|PRN|AUX|NUL|CLOCK\$|CONIN\$|CONOUT\$|COM[1-9¹²³]|LPT[1-9¹²³])$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                throw new ArgumentException("예약된 Windows 장치 이름은 사용할 수 없습니다.");
        }
        if (!directory && drive && tail.Length == 0) throw new ArgumentException("파일 경로가 필요합니다.");
        return directory ? Path.TrimEndingDirectorySeparator(path) : path;
    }
}
