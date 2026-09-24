using FileListToExcel.Core;
using System.Diagnostics;
using System.Text.Json;

namespace FileListToExcel.App;

internal static class Program
{
    internal const string ProductName = "File List to Excel";
    internal const string HelpText = "탐색기에서 파일/폴더를 선택한 뒤 '더 많은 옵션 표시' → Excel 메뉴를 사용하세요.\n\nCLI:\nFileListToExcel.exe --files <paths...>\nFileListToExcel.exe --folder <paths...>\nFileListToExcel.exe --recursive <paths...>\n\n자동화: --no-open --output <새 파일.xlsx>\n기존 파일은 덮어쓰지 않습니다.";

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
                var message = options.Version ? "File List to Excel 1.0.0" : HelpText;
                if (headless) Console.WriteLine(message); else MessageBox.Show(message, ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
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
            destination = Path.Combine(directory, $"FileList_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}.xlsx");
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
        var message = $"파일 목록을 만들지 못했습니다.\n\n{ex.Message}";
        if (headless) Console.Error.WriteLine(message);
        else MessageBox.Show(message, ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
    internal int ExitCode { get; private set; }

    internal ExportContext(AppOptions options)
    {
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
                    window = new ProgressWindow(cancellation);
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
