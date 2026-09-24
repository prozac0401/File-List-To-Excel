using FileListToExcel.Core;

namespace FileListToExcel.App;

internal sealed class ProgressWindow : Form
{
    private readonly CancellationTokenSource cancellation;
    private readonly bool duplicateJob;
    private readonly Label count = new() { AutoSize = false, Dock = DockStyle.Top, Height = 30 };
    private readonly Label path = new() { AutoSize = false, Dock = DockStyle.Fill, AutoEllipsis = true };
    private readonly Button cancel = new() { Text = "취소", AutoSize = true, Anchor = AnchorStyles.Right };
    private bool complete;

    internal ProgressWindow(CancellationTokenSource cancellation, bool duplicateJob = false)
    {
        this.cancellation = cancellation;
        this.duplicateJob = duplicateJob;
        Text = Program.ProductName;
        ClientSize = duplicateJob ? new Size(620, 300) : new Size(520, 205);
        if (duplicateJob) count.Height = 86;
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = new Font("맑은 고딕", 10);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false; MinimizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        Padding = new Padding(24);
        var title = new Label { Text = duplicateJob ? "중복 파일 검사 중…" : "파일 목록을 Excel로 만드는 중…", AutoSize = false, Height = 34, Dock = DockStyle.Top, Font = new Font(Font, FontStyle.Bold) };
        var bar = new ProgressBar { Dock = DockStyle.Bottom, Height = 5, Style = ProgressBarStyle.Marquee, MarqueeAnimationSpeed = 30 };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 43, FlowDirection = FlowDirection.RightToLeft };
        buttons.Controls.Add(cancel);
        Controls.Add(path); Controls.Add(count); Controls.Add(title); Controls.Add(buttons); Controls.Add(bar);
        cancel.Click += (_, _) => Cancel();
        FormClosing += (_, e) => { if (!complete) { e.Cancel = true; Cancel(); } };
        AccessibleName = duplicateJob ? "중복 파일 검사 진행 상황" : "파일 목록 생성 진행 상황";
    }

    private void Cancel()
    {
        cancellation.Cancel();
        cancel.Enabled = false;
        count.Text = "취소 중… 현재 파일 시스템 작업이 끝나면 종료합니다.";
    }

    internal void UpdateProgress(ScanProgress? value)
    {
        if (cancellation.IsCancellationRequested || value == null) return;
        if (value.Duplicate is { } duplicate)
        {
            string phase = duplicate.Stage switch
            {
                DuplicateStage.FindingFiles => "파일 찾기",
                DuplicateStage.SizeGrouping => "크기로 후보 추리기",
                DuplicateStage.QuickFingerprint => "빠른 지문 검사",
                DuplicateStage.FullHash => "전체 SHA-256 검사",
                _ => "검사 완료"
            };
            count.Text = $"{phase}\n발견: {duplicate.FileCount:N0}개    크기 후보: {duplicate.SizeCandidates:N0}개\n지문 통과: {duplicate.QuickCandidates:N0}개    Hash 계산: {duplicate.HashedFiles:N0}회";
        }
        else count.Text = duplicateJob
            ? $"Excel 결과 저장 중…\n검사 파일: {value.ItemCount:N0}개    오류: {value.ErrorCount:N0}개"
            : $"수집: {value.ItemCount:N0}개    오류: {value.ErrorCount:N0}개";
        path.Text = value.CurrentPath;
    }

    internal void Complete() => complete = true;
}
