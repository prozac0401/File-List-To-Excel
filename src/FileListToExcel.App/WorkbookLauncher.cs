using Microsoft.Win32;
using System.ComponentModel;
using System.Diagnostics;

namespace FileListToExcel.App;

internal static class WorkbookLauncher
{
    internal static void Open(string path)
    {
        foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
        foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            try
            {
                using var registry = RegistryKey.OpenBaseKey(hive, view);
                using var key = registry.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe");
                var executable = (key?.GetValue(null) as string)?.Trim('"');
                if (string.IsNullOrEmpty(executable) || !Path.IsPathFullyQualified(executable) || !File.Exists(executable)) continue;
                var start = new ProcessStartInfo(executable) { UseShellExecute = false };
                start.ArgumentList.Add(path);
                using var process = Process.Start(start);
                if (process != null) return;
            }
            catch (Exception ex) when (ex is Win32Exception or IOException or UnauthorizedAccessException or System.Security.SecurityException) { }
        }
        try
        {
            using var process = Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Win32Exception)
        {
            MessageBox.Show($"Excel 파일이 생성되었습니다.\n연결된 앱을 찾지 못했습니다. Excel 또는 .xlsx 호환 앱에서 아래 파일을 열어 주세요.\n\n{path}", Program.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
