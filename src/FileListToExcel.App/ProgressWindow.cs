using FileListToExcel.Core;

namespace FileListToExcel.App;

internal sealed class ProgressWindow : Form
{
    private readonly CancellationTokenSource cancellation;
    private readonly Label count = new() { AutoSize = false, Dock = DockStyle.Top, Height = 30 };
    private readonly Label path = new() { AutoSize = false, Dock = DockStyle.Fill, AutoEllipsis = true };
    private readonly Button cancel = new() { Text = "취소", AutoSize = true, Anchor = AnchorStyles.Right };
    private bool complete;

    internal ProgressWindow(CancellationTokenSource cancellation)
    {
        this.cancellation = cancellation;
        Text = Program.ProductName;
        ClientSize = new Size(520, 205);
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = new Font("맑은 고딕", 10);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false; MinimizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        Padding = new Padding(24);
        var title = new Label { Text = "파일 목록을 Excel로 만드는 중…", AutoSize = false, Height = 34, Dock = DockStyle.Top, Font = new Font(Font, FontStyle.Bold) };
        var bar = new ProgressBar { Dock = DockStyle.Bottom, Height = 5, Style = ProgressBarStyle.Marquee, MarqueeAnimationSpeed = 30 };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 43, FlowDirection = FlowDirection.RightToLeft };
        buttons.Controls.Add(cancel);
        Controls.Add(path); Controls.Add(count); Controls.Add(title); Controls.Add(buttons); Controls.Add(bar);
        cancel.Click += (_, _) => Cancel();
        FormClosing += (_, e) => { if (!complete) { e.Cancel = true; Cancel(); } };
        AccessibleName = "파일 목록 생성 진행 상황";
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
        count.Text = $"수집: {value.ItemCount:N0}개    오류: {value.ErrorCount:N0}개";
        path.Text = value.CurrentPath;
    }

    internal void Complete() => complete = true;
}
