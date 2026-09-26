using System.Runtime.InteropServices;
using FileListToExcel.Core;

namespace FileListToExcel.App;

/// <summary>One native destination dialog; the final button explicitly authorizes copying.</summary>
internal static class CollectFolderPicker
{
    internal static string? Pick(CopyPlan plan, IntPtr owner)
    {
        var dialog = (IFileDialog)new FileOpenDialog();
        try
        {
            dialog.SetOptions(0x20 | 0x40 | 0x800 | 0x1000); // folders, filesystem, existing path/file
            dialog.SetTitle("선택한 파일 복사 · 목적 폴더");
            dialog.SetOkButtonLabel("여기에 복사");
            var customize = (IFileDialogCustomize)dialog;
            customize.AddText(99, $"선택 {plan.SelectedCount:N0}행 · 복사 가능 {plan.EligibleFiles.Count:N0}개 · 예상 {plan.TotalBytes:N0} bytes · 제외 {plan.ExcludedCount:N0}개 · 중복 {plan.DuplicateCount:N0}개");
            customize.AddText(100, "선택한 폴더 아래 새 '모은파일_날짜_시간_식별자' 폴더에 복사합니다.");
            customize.AddText(101, "원본은 변경하지 않습니다. 내부 원본 경로 보고서는 이 폴더에 넣지 않습니다.");
            int hr = dialog.Show(owner);
            if (hr == unchecked((int)0x800704C7)) return null;
            Marshal.ThrowExceptionForHR(hr);
            dialog.GetResult(out var item);
            try
            {
                item.GetDisplayName(0x80058000, out var name);
                try { return Marshal.PtrToStringUni(name); }
                finally { Marshal.FreeCoTaskMem(name); }
            }
            finally { Marshal.ReleaseComObject(item); }
        }
        finally { Marshal.ReleaseComObject(dialog); }
    }

    [ComImport, Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
    private class FileOpenDialog { }
    [ComImport, Guid("42F85136-DB7E-439C-85F1-E4075D135FC8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileDialog
    {
        [PreserveSig] int Show(IntPtr parent);
        void SetFileTypes(uint count, IntPtr filters);
        void SetFileTypeIndex(uint index);
        void GetFileTypeIndex(out uint index);
        void Advise(IntPtr sink, out uint cookie);
        void Unadvise(uint cookie);
        void SetOptions(uint options);
        void GetOptions(out uint options);
        void SetDefaultFolder(IShellItem item);
        void SetFolder(IShellItem item);
        void GetFolder(out IShellItem item);
        void GetCurrentSelection(out IShellItem item);
        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string name);
        void GetFileName(out IntPtr name);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string title);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string label);
        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string label);
        void GetResult(out IShellItem item);
        void AddPlace(IShellItem item, uint placement);
        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string extension);
        void Close(int result);
        void SetClientGuid(ref Guid guid);
        void ClearClientData();
        void SetFilter(IntPtr filter);
    }
    [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem
    {
        void BindToHandler(IntPtr context, ref Guid handler, ref Guid iid, out IntPtr value);
        void GetParent(out IShellItem item);
        void GetDisplayName(uint kind, out IntPtr value);
        void GetAttributes(uint mask, out uint attributes);
        void Compare(IShellItem other, uint hint, out int order);
    }
    [ComImport, Guid("E6FDD21A-163F-4975-9C8C-A69F1BA37034"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileDialogCustomize
    {
        void EnableOpenDropDown(uint id);
        void AddMenu(uint id, [MarshalAs(UnmanagedType.LPWStr)] string text);
        void AddPushButton(uint id, [MarshalAs(UnmanagedType.LPWStr)] string text);
        void AddComboBox(uint id);
        void AddRadioButtonList(uint id);
        void AddCheckButton(uint id, [MarshalAs(UnmanagedType.LPWStr)] string text, [MarshalAs(UnmanagedType.Bool)] bool check);
        void AddEditBox(uint id, [MarshalAs(UnmanagedType.LPWStr)] string text);
        void AddSeparator(uint id);
        void AddText(uint id, [MarshalAs(UnmanagedType.LPWStr)] string text);
    }
}
