using FileListToExcel.Core;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace FileListToExcel.App;

public sealed record AppOptions(ScanRequest? Request, string? OutputPath, bool NoOpen, bool Help, bool Version,
    DuplicateRequest? DuplicateRequest = null, bool NoCache = false);

public static class CommandLine
{
    public static AppOptions Parse(string[] args)
    {
        ScanMode? mode = null;
        DuplicateMode? duplicateMode = null;
        var paths = new List<string>();
        string? output = null, requestFile = null;
        bool noOpen = false, help = false, version = false, literal = false, noCache = false;
        for (int i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (literal) { paths.Add(arg); continue; }
            switch (arg)
            {
                case "--": literal = true; break;
                case "--files": SetMode(ScanMode.SelectedItems); break;
                case "--folder": SetMode(ScanMode.DirectChildren); break;
                case "--recursive": SetMode(ScanMode.Recursive); break;
                case "--matches": SetDuplicateMode(DuplicateMode.Matches); break;
                case "--duplicates": SetDuplicateMode(DuplicateMode.Folders); break;
                case "--duplicate-files": SetDuplicateMode(DuplicateMode.SelectedFiles); break;
                case "--no-cache": noCache = true; break;
                case "--no-open": noOpen = true; break;
                case "--help": case "-h": help = true; break;
                case "--version": version = true; break;
                case "--output":
                    if (output != null) throw new ArgumentException("--output 옵션은 한 번만 지정하세요.");
                    output = Next(); break;
                case "--request":
                    if (requestFile != null) throw new ArgumentException("--request 옵션은 한 번만 지정하세요.");
                    requestFile = Next(); break;
                default:
                    if (arg.StartsWith("--", StringComparison.Ordinal)) throw new ArgumentException($"알 수 없는 옵션: {arg}");
                    paths.Add(arg); break;
            }
            string Next() => ++i < args.Length ? args[i] : throw new ArgumentException($"{arg} 뒤에 값을 입력하세요.");
        }
        if (help || version) return new(null, null, noOpen, help, version);
        ScanRequest? request = null;
        DuplicateRequest? duplicateRequest = null;
        if (requestFile != null)
        {
            if (mode != null || duplicateMode != null || paths.Count != 0) throw new ArgumentException("요청 파일과 경로 옵션을 함께 지정할 수 없습니다.");
            (request, duplicateRequest) = ReadShellRequest(requestFile);
        }
        else
        {
            if ((mode == null && duplicateMode == null) || paths.Count == 0) throw new ArgumentException("탐색기에서 파일 또는 폴더를 선택해 우클릭 메뉴로 실행하세요. CLI 사용법: --help");
            var fullPaths = paths.Select(Path.GetFullPath).ToArray();
            if (duplicateMode is { } duplicate) duplicateRequest = new(fullPaths, duplicate);
            else request = new(fullPaths, mode!.Value);
        }
        var inputs = request?.Paths ?? duplicateRequest!.Paths;
        if (inputs.Count > 100_000) throw new ArgumentException("한 번에 선택할 수 있는 최대 항목 수는 100,000개입니다.");
        ValidateDuplicate(duplicateRequest);
        if (output != null)
        {
            output = Path.GetFullPath(output);
            if (!string.Equals(Path.GetExtension(output), ".xlsx", StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("출력 파일 확장자는 .xlsx여야 합니다.");
            if (File.Exists(output)) throw new IOException("출력 파일이 이미 있습니다. 기존 파일을 덮어쓰지 않습니다.");
        }
        return new(request, output, noOpen, false, false, duplicateRequest, noCache);

        void SetMode(ScanMode value)
        {
            if (mode != null || duplicateMode != null) throw new ArgumentException("목록/중복 검사 명령 중 하나만 지정하세요.");
            mode = value;
        }
        void SetDuplicateMode(DuplicateMode value)
        {
            if (mode != null || duplicateMode != null) throw new ArgumentException("목록/중복 검사 명령 중 하나만 지정하세요.");
            duplicateMode = value;
        }
    }

    public static string RequestDirectory => Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "FileListToExcel", "Requests"));

    private static (ScanRequest? Scan, DuplicateRequest? Duplicate) ReadShellRequest(string path)
    {
        path = Path.GetFullPath(path);
        if (!string.Equals(Path.GetDirectoryName(path), RequestDirectory, StringComparison.OrdinalIgnoreCase)
            || !Guid.TryParseExact(Path.GetFileNameWithoutExtension(path), "D", out _)
            || !string.Equals(Path.GetExtension(path), ".json", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("유효한 탐색기 요청 파일이 아닙니다.");
        if ((File.GetAttributes(RequestDirectory) & FileAttributes.ReparsePoint) != 0 || (File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0)
            throw new IOException("요청 파일은 링크일 수 없습니다.");
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
        if (stream.Length is <= 0 or > 33_554_432) throw new ArgumentException("요청 파일 크기가 유효하지 않습니다.");
        var payload = JsonSerializer.Deserialize<ShellRequest>(stream) ?? throw new ArgumentException("요청 파일이 비어 있습니다.");
        if (payload.Paths is not { Length: > 0 and <= 100_000 } || payload.Paths.Any(p => string.IsNullOrWhiteSpace(p) || !Path.IsPathFullyQualified(p)))
            throw new ArgumentException("요청 경로가 유효하지 않습니다.");
        (ScanRequest? Scan, DuplicateRequest? Duplicate) request = payload.Mode switch
        {
            "files" => (new(payload.Paths, ScanMode.SelectedItems), null),
            "folder" => (new(payload.Paths, ScanMode.DirectChildren), null),
            "recursive" => (new(payload.Paths, ScanMode.Recursive), null),
            "matches" => (null, new(payload.Paths, DuplicateMode.Matches)),
            "duplicates" => (null, new(payload.Paths, DuplicateMode.Folders)),
            "duplicate-files" => (null, new(payload.Paths, DuplicateMode.SelectedFiles)),
            _ => throw new ArgumentException("요청 작업 유형이 유효하지 않습니다.")
        };
        ValidateDuplicate(request.Duplicate);
        stream.Dispose();
        File.Delete(path); // Only our validated, one-use request; never an input source file.
        return request;
    }

    private static void ValidateDuplicate(DuplicateRequest? request)
    {
        if (request is { Mode: DuplicateMode.Matches, Paths.Count: not 1 })
            throw new ArgumentException("같은 파일 찾기는 파일 하나를 선택하세요.");
    }

    private sealed record ShellRequest([property: JsonPropertyName("mode")] string Mode, [property: JsonPropertyName("paths")] string[] Paths);
}
