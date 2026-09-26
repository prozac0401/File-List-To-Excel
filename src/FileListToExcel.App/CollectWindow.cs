using System.Diagnostics;
using FileListToExcel.Core;

namespace FileListToExcel.App;

internal sealed class CollectWindow : Form
{
    private readonly CollectRequest request;
    private readonly CancellationTokenSource cancellation = new();
    private readonly Label status = new() { AutoSize = false, Dock = DockStyle.Fill, Padding = new Padding(18), Text = "선택한 파일을 확인하고 있습니다…", AutoEllipsis = true };
    private readonly ProgressBar progressBar = new() { Dock = DockStyle.Bottom, Height = 20, Style = ProgressBarStyle.Marquee };
    private readonly FlowLayoutPanel actions = new() { Dock = DockStyle.Bottom, Height = 54, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(8) };
    private readonly Button cancel = new() { Text = "취소", AutoSize = true };
    private bool complete;
    internal int ExitCode { get; private set; }

    internal CollectWindow(CollectRequest request)
    {
        this.request = request;
        Text = "파일목록 · 선택한 파일 복사";
        ClientSize = new Size(660, 235);
        MinimumSize = new Size(550, 250);
        StartPosition = FormStartPosition.CenterScreen;
        actions.Controls.Add(cancel);
        Controls.Add(status);
        Controls.Add(progressBar);
        Controls.Add(actions);
        cancel.Click += (_, _) => RequestCancel();
        FormClosing += (_, e) => { if (!complete) { e.Cancel = true; RequestCancel(); } };
        Shown += async (_, _) => await RunAsync();
    }

    private void RequestCancel()
    {
        cancellation.Cancel();
        cancel.Enabled = false;
        status.Text = "취소 요청 중… 현재 네트워크/OS 호출이 끝날 때까지 시간이 걸릴 수 있습니다.";
    }

    private async Task RunAsync()
    {
        try
        {
            var plan = await Task.Run(() => new CopyPlanner().BuildPlan(request.Rows, cancellation.Token));
            CopyResult result;
            if (cancellation.IsCancellationRequested)
                result = await new CopyService().RecordCancelledAsync(plan);
            else if (plan.EligibleFiles.Count == 0)
                result = await new CopyService().RecordPlanAsync(plan);
            else
            {
                status.Text = $"복사 가능 {plan.EligibleFiles.Count:N0}개 · 예상 {plan.TotalBytes:N0} bytes\n제외 {plan.ExcludedCount:N0}개 · 같은 원본 중복 {plan.DuplicateCount:N0}개";
                var destination = CollectFolderPicker.Pick(plan, Handle);
                if (destination is null)
                    result = await new CopyService().RecordCancelledAsync(plan);
                else
                {
                    var progress = new Progress<CopyProgress>(value =>
                    {
                        if (cancellation.IsCancellationRequested) return;
                        status.Text = $"{value.FileName}\n\n완료 {value.CompletedCount:N0}/{value.TotalCount:N0}개\n전송 {value.TransferredBytes:N0}/{value.TotalBytes:N0} bytes";
                        progressBar.Style = ProgressBarStyle.Continuous;
                        progressBar.Value = value.TotalBytes <= 0 ? 0 : (int)Math.Clamp((double)value.TransferredBytes / value.TotalBytes * 100, 0, 100);
                    });
                    result = await new CopyService().CopyAsync(plan, destination, progress, cancellation.Token);
                }
            }
            ShowResult(result, plan.SelectedCount);
        }
        catch (Exception error)
        {
            complete = true;
            ExitCode = 1;
            status.Text = $"파일 복사를 마치지 못했습니다.\n\n{error.Message}";
            ShowClose();
        }
    }

    private void ShowResult(CopyResult result, int selected)
    {
        complete = true;
        ExitCode = result.Cancelled ? 2 : result.FailedCount > 0 ? 1 : 0;
        progressBar.Style = ProgressBarStyle.Continuous;
        progressBar.Value = 100;
        status.Text = $"{selected:N0}개 선택 · {result.CopiedCount:N0}개 복사 완료\n제외 {result.ExcludedCount:N0}개 · 실패 {result.FailedCount:N0}개" +
            (result.Cancelled ? "\n취소됨: 완료된 복사본은 유지했습니다." : "") +
            (string.IsNullOrEmpty(result.OutputDirectory) ? "" : $"\n\n{result.OutputDirectory}") +
            (result.ReportError is null ? "" : $"\n내부 결과 기록 실패: {result.ReportError}");
        ShowClose();
        if (!string.IsNullOrEmpty(result.OutputDirectory))
            AddAction("모은 폴더 열기", () => Process.Start(new ProcessStartInfo("explorer.exe") { UseShellExecute = true, ArgumentList = { result.OutputDirectory } }));
        if (result.Outcomes.Count > 0)
            AddAction("작업 결과 보기", () => ShowReport(result));
    }

    private void ShowReport(CopyResult result)
    {
        using var report = new Form { Text = "작업 결과 · 내부용 원본 경로 포함", Width = 1000, Height = 500, StartPosition = FormStartPosition.CenterParent };
        report.Controls.Add(CreateReportGrid(result));
        report.ShowDialog(this);
    }

    internal static DataGridView CreateReportGrid(CopyResult result)
    {
        var grid = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, DataSource = result.Outcomes.ToList() };
        grid.DataBindingComplete += (_, _) =>
        {
            foreach (DataGridViewColumn column in grid.Columns)
                column.HeaderText = column.DataPropertyName switch
                {
                    nameof(CopyOutcome.RowIndex) => "선택 순번", nameof(CopyOutcome.ItemId) => "항목 ID",
                    nameof(CopyOutcome.SourcePath) => "원본 경로", nameof(CopyOutcome.DestinationPath) => "복사 경로",
                    nameof(CopyOutcome.Status) => "결과", nameof(CopyOutcome.Code) => "코드",
                    nameof(CopyOutcome.Message) => "내용", _ => column.HeaderText
                };
        };
        grid.CellFormatting += (_, e) =>
        {
            if (e.ColumnIndex >= 0 && grid.Columns[e.ColumnIndex].DataPropertyName == nameof(CopyOutcome.RowIndex) && e.Value is int index)
            { e.Value = (index + 1).ToString(System.Globalization.CultureInfo.CurrentCulture); e.FormattingApplied = true; }
            else if (e.Value is CopyStatus outcome)
            {
                e.Value = outcome switch { CopyStatus.Copied => "복사 완료", CopyStatus.Excluded => "제외",
                    CopyStatus.Failed => "실패", CopyStatus.Cancelled => "취소", CopyStatus.Duplicate => "중복 선택", _ => outcome.ToString() };
                e.FormattingApplied = true;
            }
        };
        return grid;
    }
    private void ShowClose()
    {
        actions.Controls.Clear();
        AddAction("닫기", Close);
    }
    private void AddAction(string text, Action action)
    {
        var button = new Button { Text = text, AutoSize = true };
        button.Click += (_, _) => { try { action(); } catch (Exception error) { MessageBox.Show(this, error.Message, Text); } };
        actions.Controls.Add(button);
    }
    protected override void Dispose(bool disposing)
    {
        if (disposing) cancellation.Dispose();
        base.Dispose(disposing);
    }
}
