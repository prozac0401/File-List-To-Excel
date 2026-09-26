using System.Reflection;
using System.Windows.Forms;
using FileListToExcel.App;
using FileListToExcel.Core;
using Xunit;

namespace FileListToExcel.App.Tests;

public sealed class CollectReportGridTests
{
    [Fact]
    public async Task RealStaGridFormatsOneBasedRowsAndEveryOutcomeWithoutDataErrors()
    {
        var completed = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var thread = new Thread(() =>
        {
            try
            {
                var statuses = new[] { CopyStatus.Copied, CopyStatus.Excluded, CopyStatus.Failed, CopyStatus.Cancelled, CopyStatus.Duplicate };
                var rows = statuses.Select((status, index) => new CopyOutcome(index, Guid.NewGuid().ToString("D"),
                    @"C:\fixtures\=literal.txt", status == CopyStatus.Copied ? @"C:\delivery\=literal.txt" : null,
                    status, "Code", "상세 결과")).ToArray();
                var result = new CopyResult(null, null, rows, true, "보고서 저장 실패 fixture");
                var window = typeof(CommandLine).Assembly.GetType("FileListToExcel.App.CollectWindow", throwOnError: true)!;
                var factory = window.GetMethod("CreateReportGrid", BindingFlags.NonPublic | BindingFlags.Static)!;
                using var grid = (DataGridView)factory.Invoke(null, [result])!;
                using var host = new Form { BindingContext = new BindingContext() };
                var errors = new List<Exception>();
                grid.DataError += (_, error) => { if (error.Exception is not null) errors.Add(error.Exception); error.ThrowException = false; };
                host.Controls.Add(grid);
                _ = host.Handle;
                _ = grid.Handle;
                Assert.Equal(statuses.Length, grid.Rows.Count);
                Assert.Equal("선택 순번", grid.Columns[nameof(CopyOutcome.RowIndex)]!.HeaderText);
                Assert.Equal("결과", grid.Columns[nameof(CopyOutcome.Status)]!.HeaderText);
                var expectedStatuses = new[] { "복사 완료", "제외", "실패", "취소", "중복 선택" };
                for (int index = 0; index < grid.Rows.Count; index++)
                {
                    // Exercise WinForms' real formatting pipeline, including its FormattedValueType check.
                    foreach (DataGridViewCell cell in grid.Rows[index].Cells) _ = cell.FormattedValue;
                    Assert.Equal((index + 1).ToString(System.Globalization.CultureInfo.CurrentCulture),
                        Assert.IsType<string>(grid.Rows[index].Cells[nameof(CopyOutcome.RowIndex)].FormattedValue));
                    Assert.Equal(expectedStatuses[index],
                        Assert.IsType<string>(grid.Rows[index].Cells[nameof(CopyOutcome.Status)].FormattedValue));
                    Assert.Equal(index, rows[index].RowIndex);
                }
                Assert.Empty(errors);
                completed.SetResult();
            }
            catch (Exception error) { completed.SetException(error); }
        }) { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        await completed.Task.WaitAsync(TimeSpan.FromSeconds(30));
    }
}
