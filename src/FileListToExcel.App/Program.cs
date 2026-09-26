using FileListToExcel.Core;
using System.Diagnostics;
using System.Text.Json;

namespace FileListToExcel.App;

internal static class Program
{
    internal const string ProductName = "File List to Excel";
    internal const string HelpText = "탐색기에서 파일/폴더를 선택한 뒤 '더 많은 옵션 표시' → Excel/중복 메뉴를 사용하세요.\n\nCLI:\nFileListToExcel.exe --files <paths...>\nFileListToExcel.exe --folder <paths...>\nFileListToExcel.exe --recursive <paths...>\nFileListToExcel.exe --matches <file>\nFileListToExcel.exe --duplicates <folders...>\nFileListToExcel.exe --duplicate-files <files...>\n\n자동화: --no-open --output <새 파일.xlsx>\n캐시 없이 검사: --no-cache\n기존 파일과 원본은 변경하지 않습니다.";

    [STAThread]
    private static int Main(string[] args)
    {
        bool headless = args.TakeWhile(arg => arg != "--").Contains("--no-open", StringComparer.Ordinal);
        try
        {
            var options = CommandLine.Parse(args);
            headless = options.NoOpen;
            if (options.Help || options.Version)
            {
                var message = options.Version ? $"File List to Excel {typeof(Program).Assembly.GetName().Version?.ToString(3)}" : HelpText;
                if (headless) Console.WriteLine(message); else MessageBox.Show(message, ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }
            if (options.CollectRequest is { } collectRequest)
            {
                ApplicationConfiguration.Initialize();
                using var collect = new CollectWindow(collectRequest);
                Application.Run(collect);
                return collect.ExitCode;
            }
            if (headless)
            {
                using var cancellation = new CancellationTokenSource();
                ConsoleCancelEventHandler handler = (_, e) => { e.Cancel = true; cancellation.Cancel(); };
                Console.CancelKeyPress += handler;
                try
                {
                    var result = Export(options, null, cancellation.Token);
                    Console.WriteLine(JsonSerializer.Serialize(result));
                    return 0;
                }
                finally { Console.CancelKeyPress -= handler; }
            }
            ApplicationConfiguration.Initialize();
            using var context = new ExportContext(options);
            Application.Run(context);
            return context.ExitCode;
        }
        catch (OperationCanceledException) { return 2; }
        catch (Exception ex)
        {
            ReportError(ex, headless);
            return 1;
        }
    }

    internal static ExportResult Export(AppOptions options, IProgress<ScanProgress>? progress, CancellationToken cancellation)
    {
        var destination = options.OutputPath;
        if (destination == null)
        {
            var directory = Path.Combine(Path.GetTempPath(), "FileListToExcel");
            Directory.CreateDirectory(directory);
            var prefix = options.DuplicateRequest is null ? "FileList" : options.DuplicateRequest.Mode == DuplicateMode.Matches ? "Matches" : "Duplicates";
            destination = Path.Combine(directory, $"{prefix}_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}.xlsx");
        }
        if (options.DuplicateRequest is { } duplicateRequest)
        {
            using var cache = options.NoCache ? null : new SqliteHashCache();
            var duplicateProgress = new DuplicateProgressAdapter(progress);
            var result = new DuplicateScanService(cache).ScanAsync(duplicateRequest, duplicateProgress, cancellation).GetAwaiter().GetResult();
            cancellation.ThrowIfCancellationRequested();
            progress?.Report(new(result.Stats.ScannedFiles, result.Errors.Count, "Excel 결과 저장 중…"));
            return new WorkbookWriter().WriteDuplicates(destination, result, duplicateRequest.Mode == DuplicateMode.Matches, cancellation);
        }
        var scanner = new FileScanner();
        return new WorkbookWriter().Write(destination, scanner.Enumerate(options.Request!, progress, cancellation), cancellation);
    }

    internal static void ReportError(Exception ex, bool headless)
    {
        try
        {
            var logs = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "FileListToExcel", "Logs");
            Directory.CreateDirectory(logs);
            File.WriteAllText(Path.Combine(logs, "last-error.log"), $"{DateTimeOffset.Now:O}\n{ex}");
        }
        catch (Exception logError) when (logError is IOException or UnauthorizedAccessException) { }
        var message = $"Excel 결과를 만들지 못했습니다.\n\n{ex.Message}";
        if (headless) Console.Error.WriteLine(message);
        else MessageBox.Show(message, ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
    }

    private sealed class DuplicateProgressAdapter(IProgress<ScanProgress>? target) : IProgress<DuplicateProgress>
    {
        public void Report(DuplicateProgress value) => target?.Report(new(value.FileCount, 0, value.CurrentPath, value));
    }
}

internal sealed class ExportContext : ApplicationContext
{
    private readonly CancellationTokenSource cancellation = new();
    private readonly System.Windows.Forms.Timer timer = new() { Interval = 100 };
    private readonly Stopwatch elapsed = Stopwatch.StartNew();
    private readonly Task<ExportResult> work;
    private ScanProgress? progress;
    private ProgressWindow? window;
    private readonly bool duplicateJob;
    internal int ExitCode { get; private set; }

    internal ExportContext(AppOptions options)
    {
        duplicateJob = options.DuplicateRequest is not null;
        work = Task.Run(() => Program.Export(options, new CallbackProgress(p => Volatile.Write(ref progress, p)), cancellation.Token));
        timer.Tick += Tick;
        timer.Start();
    }

    private void Tick(object? sender, EventArgs args)
    {
        if (!work.IsCompleted)
        {
            if (elapsed.ElapsedMilliseconds >= 700)
            {
                if (window == null)
                {
                    window = new ProgressWindow(cancellation, duplicateJob);
                    window.Show();
                }
                window.UpdateProgress(Volatile.Read(ref progress));
            }
            return;
        }
        timer.Stop();
        window?.Complete();
        window?.Close();
        try
        {
            var result = work.GetAwaiter().GetResult();
            cancellation.Token.ThrowIfCancellationRequested();
            WorkbookLauncher.Open(result.OutputPath);
        }
        catch (OperationCanceledException) { ExitCode = 2; }
        catch (Exception ex) { ExitCode = 1; Program.ReportError(ex, false); }
        ExitThread();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) { timer.Dispose(); window?.Dispose(); cancellation.Dispose(); }
        base.Dispose(disposing);
    }
    private sealed class CallbackProgress(Action<ScanProgress> callback) : IProgress<ScanProgress>
    {
        public void Report(ScanProgress value) => callback(value);
    }
}
