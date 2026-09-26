using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace FileListToExcel.Core;

/// <summary>Windows-only, fail-closed primitives. Shares no write/delete access to sources or ancestor directories.</summary>
internal static class CopyNative
{
    private const uint ReadAttributes = 0x80, GenericRead = 0x80000000, DeleteAccess = 0x10000;
    private const uint OpenNoRecall = 0x00100000, OpenReparsePoint = 0x00200000, BackupSemantics = 0x02000000;
    internal readonly record struct FileStamp(long Length, long ModifiedUtcTicks, uint Volume, ulong FileId, FileAttributes Attributes)
    {
        public bool SameIdentity(FileStamp other) => Volume == other.Volume && FileId == other.FileId;
        public bool SameFile(FileStamp other) => SameIdentity(other) && Length == other.Length && ModifiedUtcTicks == other.ModifiedUtcTicks;
    }

    internal sealed class DirectoryGuard : IDisposable
    {
        private readonly List<(string Path, SafeFileHandle Handle)> handles = [];
        public SafeFileHandle Leaf => handles[^1].Handle;
        public void Add(string path)
        {
            var handle = Open(path, ReadAttributes, true);
            try { CheckHandle(handle, path, true); handles.Add((path, handle)); }
            catch { handle.Dispose(); throw; }
        }
        public void Verify()
        {
            foreach (var item in handles) CheckHandle(item.Handle, item.Path, true);
        }
        public void Dispose()
        {
            for (int i = handles.Count - 1; i >= 0; i--) handles[i].Handle.Dispose();
            handles.Clear();
        }
    }

    internal sealed class SourceLease(DirectoryGuard directories, SafeFileHandle handle, string path) : IDisposable
    {
        public FileStamp ReadStamp() { directories.Verify(); return CheckHandle(handle, path, false); }
        public void Dispose() { handle.Dispose(); directories.Dispose(); }
    }

    internal static DirectoryGuard LockDirectoryChain(string path)
    {
        EnsureWindows();
        path = CollectPathPolicy.Validate(path, directory: true);
        var chain = new Stack<string>();
        for (string? current = path; current is not null; current = Path.GetDirectoryName(Path.TrimEndingDirectorySeparator(current))) chain.Push(current);
        var guard = new DirectoryGuard();
        try { while (chain.TryPop(out var next)) guard.Add(next); return guard; }
        catch { guard.Dispose(); throw; }
    }

    internal static SourceLease OpenSource(string path)
    {
        path = CollectPathPolicy.Validate(path);
        var directories = LockDirectoryChain(Path.GetDirectoryName(path)!);
        try
        {
            // Reuse duplicate detection's conservative cloud/reparse policy before any content open.
            CheckAttributes(path, File.GetAttributes(path), false);
            var handle = Open(path, GenericRead, false);
            try { CheckHandle(handle, path, false); return new(directories, handle, path); }
            catch { handle.Dispose(); throw; }
        }
        catch { directories.Dispose(); throw; }
    }

    private static SafeFileHandle Open(string path, uint access, bool directory, uint sharing = (uint)FileShare.Read)
    {
        EnsureWindows();
        var handle = CreateFile(NativePath(path), access, sharing, IntPtr.Zero, 3,
            OpenNoRecall | OpenReparsePoint | (directory ? BackupSemantics : 0), IntPtr.Zero);
        if (!handle.IsInvalid) return handle;
        int error = Marshal.GetLastWin32Error(); handle.Dispose(); throw Error(error);
    }

    private static FileStamp CheckHandle(SafeFileHandle handle, string path, bool directory)
    {
        if (GetFileType(handle) != 1) throw new CollectSafetyException("NotDiskFile", "일반 디스크 파일/폴더만 지원합니다.");
        var stamp = ReadStamp(handle);
        CheckAttributes(path, stamp.Attributes, directory);
        string actual = FinalPath(handle);
        if (!string.Equals(Path.TrimEndingDirectorySeparator(actual), Path.TrimEndingDirectorySeparator(path), StringComparison.OrdinalIgnoreCase))
            throw new CollectSafetyException("PathChanged", $"열린 파일/폴더의 실제 경로가 요청 경로와 다릅니다. 요청: {path}, 실제: {actual}");
        return stamp;
    }

    private static string FinalPath(SafeFileHandle handle)
    {
        var final = new StringBuilder(32_768);
        uint count = GetFinalPathNameByHandle(handle, final, (uint)final.Capacity, 0);
        if (count == 0 || count >= final.Capacity) throw Error(Marshal.GetLastWin32Error());
        string actual = final.ToString();
        if (actual.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase)) actual = @"\\" + actual[8..];
        else if (actual.StartsWith(@"\\?\", StringComparison.Ordinal)) actual = actual[4..];
        return actual;
    }

    internal static string ResolveLocalReportDirectory(string path, string localRoot)
    {
        // An MSIX-hosted process may transparently virtualize newly created AppData directories.
        // Resolve only our report directory, require it to stay beneath LocalAppData, then pin its
        // complete physical chain as usual. This never accepts redirected user source/destination paths.
        using var handle = Open(path, ReadAttributes, true);
        CheckAttributes(path, ReadStamp(handle).Attributes, true);
        string actual = FinalPath(handle);
        if (!actual.StartsWith(Path.TrimEndingDirectorySeparator(localRoot) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            throw new CollectSafetyException("ReportOutsideLocalAppData", "내부 보고서 폴더가 LocalAppData 외부로 연결되어 있습니다.");
        using var physical = LockDirectoryChain(actual);
        if (!ReadStamp(handle).SameIdentity(ReadStamp(physical.Leaf)))
            throw new CollectSafetyException("ReportPathChanged", "내부 보고서 폴더가 검증 중 변경되었습니다.");
        return actual;
    }

    private static FileStamp ReadStamp(SafeFileHandle handle)
    {
        if (!GetFileInformationByHandle(handle, out var info)) throw Error(Marshal.GetLastWin32Error());
        ulong id = ((ulong)info.FileIndexHigh << 32) | info.FileIndexLow;
        if (id == 0) throw new CollectSafetyException("IdentityUnavailable", "파일 식별자를 확인할 수 없는 파일시스템입니다.");
        long fileTime = ((long)info.LastWriteTime.dwHighDateTime << 32) | (uint)info.LastWriteTime.dwLowDateTime;
        return new(((long)info.SizeHigh << 32) | info.SizeLow, DateTime.FromFileTimeUtc(fileTime).Ticks,
            info.VolumeSerialNumber, id, (FileAttributes)info.Attributes);
    }

    private static void CheckAttributes(string path, FileAttributes attributes, bool directory)
    {
        if (new LocalFileContentPolicy().GetSkipReason(path, attributes) is { } issue)
            throw new CollectSafetyException(issue.ErrorType, issue.Message);
        if (attributes.HasFlag(FileAttributes.Directory) != directory)
            throw new CollectSafetyException("NotFile", directory ? "폴더가 아닌 경로입니다." : "일반 파일이 아닌 항목입니다.");
    }

    internal static string CreateOutputDirectory(string parent, DirectoryGuard guard)
    {
        guard.Verify();
        for (int attempt = 0; attempt < 10; attempt++)
        {
            string path = Path.Combine(parent, $"모은파일_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}");
            if (CreateDirectory(NativePath(path), IntPtr.Zero)) return path;
            int error = Marshal.GetLastWin32Error();
            if (error != 183) throw Error(error);
        }
        throw new IOException("새 작업 폴더 이름을 만들 수 없습니다.");
    }

    internal static void RequireMetadataSupport(DirectoryGuard destination)
    {
        if (!GetVolumeInformationByHandle(destination.Leaf, null, 0, out _, out _, out uint flags, null, 0))
            throw Error(Marshal.GetLastWin32Error());
        // Fail closed on filesystems that can silently discard ADS/MOTW or security metadata.
        if ((flags & 0x00040008) != 0x00040008)
            throw new CollectSafetyException("DestinationMetadataUnsupported", "대상 파일시스템이 ADS/보안 속성 보존을 지원하지 않습니다. NTFS 등 지원 폴더를 선택하세요.");
    }

    internal static void CopyFile(PlannedCopyFile item, string destination, SourceLease source, DirectoryGuard output,
        Action<long> report, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        output.Verify();
        var before = source.ReadStamp();
        CopyPlanner.VerifyGeneratedStamp(item.Source, before);
        if (!item.Stamp.SameFile(before)) throw new CollectSafetyException("SourceReplaced", "사전 검증 이후 원본 파일이 변경 또는 교체되었습니다.");
        FileStamp? owned = null;
        SafeFileHandle? ownedHandle = null;
        Exception? callbackError = null;
        ProgressRoutine callback = (total, transferred, streamSize, streamTransferred, stream, reason, sourceHandle, targetHandle, data) =>
        {
            string validationStage = "source-handle";
            try
            {
                using var actualSource = new SafeFileHandle(sourceHandle, ownsHandle: false);
                using var actualTarget = new SafeFileHandle(targetHandle, ownsHandle: false);
                var sourceStamp = ReadStamp(actualSource);
                CheckAttributes(item.Source.AbsolutePath, sourceStamp.Attributes, false);
                if (!before.SameIdentity(sourceStamp) || !before.SameFile(source.ReadStamp()))
                    throw new CollectSafetyException("SourceChangedDuringCopy", "복사 중 원본 파일 변경을 감지했습니다.");
                validationStage = "destination-handle";
                // Alternate-stream destination handles may be write-only and deny metadata queries.
                // Bind ownership using the first (default) stream handle, then retain a separate
                // metadata handle allowing CopyFileEx writes but denying deletion/replacement.
                if (ownedHandle is null)
                {
                    var targetStamp = ReadStamp(actualTarget);
                    owned = targetStamp;
                    ownedHandle = Open(destination, ReadAttributes, false, (uint)(FileShare.Read | FileShare.Write));
                    if (!targetStamp.SameIdentity(CheckHandle(ownedHandle, destination, false)))
                        throw new CollectSafetyException("OutputChanged", "복사 중 대상 파일이 교체되었습니다.");
                }
                else if (!owned!.Value.SameIdentity(CheckHandle(ownedHandle, destination, false)))
                    throw new CollectSafetyException("OutputChanged", "복사 중 대상 파일이 교체되었습니다.");
                validationStage = "directory-chain";
                output.Verify();
                report(transferred);
                return token.IsCancellationRequested ? 1u : 0u;
            }
            catch (Exception ex) { callbackError = new CollectSafetyException("CopyCallbackValidationFailed", $"복사 핸들 검증 실패 ({validationStage}, stream {stream}): {ex.Message}"); return 1; }
        };
        // COPY_FILE_FAIL_IF_EXISTS | COPY_FILE_COPY_SYMLINK. No decrypted-destination or source-write flags.
        bool succeeded = CopyFileEx(NativePath(item.Source.AbsolutePath), NativePath(destination), callback, IntPtr.Zero, IntPtr.Zero, 0x801);
        int nativeError = Marshal.GetLastWin32Error();
        GC.KeepAlive(callback);
        try
        {
            if (callbackError is not null) throw callbackError;
            if (!succeeded)
            {
                token.ThrowIfCancellationRequested();
                throw Error(nativeError);
            }
            if (!before.SameFile(source.ReadStamp())) throw new CollectSafetyException("SourceChangedDuringCopy", "복사 중 원본 파일 변경을 감지했습니다.");
            output.Verify();
            var after = ownedHandle is null ? throw new CollectSafetyException("OutputOwnershipUnavailable", "복사본의 소유권을 확인할 수 없습니다.") : CheckHandle(ownedHandle, destination, false);
            if (owned is null || !owned.Value.SameIdentity(after) || after.Length != before.Length || after.ModifiedUtcTicks != before.ModifiedUtcTicks)
                throw new CollectSafetyException("OutputVerificationFailed", "복사본의 파일 식별자, 크기 또는 수정시간 검증에 실패했습니다.");
            if (before.Attributes.HasFlag(FileAttributes.Encrypted) && !after.Attributes.HasFlag(FileAttributes.Encrypted))
                throw new CollectSafetyException("EncryptionNotPreserved", "암호화 속성이 보존되지 않아 성공으로 처리하지 않습니다.");
        }
        catch (Exception ex) when (IsFileError(ex) || ex is OperationCanceledException)
        {
            ownedHandle?.Dispose();
            string? cleanup = owned is null && !succeeded && nativeError is 80 or 183 ? null : TryDeleteOwned(destination, owned);
            if (cleanup is not null)
                throw new CollectSafetyException("IncompleteOutputCleanupFailed", ex.Message + " 자체 미완료 복사본 정리 실패: " + cleanup);
            throw;
        }
        finally { ownedHandle?.Dispose(); }
    }

    private static string? TryDeleteOwned(string path, FileStamp? owned)
    {
        // No callback means no proven ownership. Never delete by a guessed name.
        if (owned is null)
        {
            try
            {
                _ = File.GetAttributes(path);
                return "대상 항목은 있으나 생성 소유권을 확인할 수 없어 삭제하지 않았습니다. 결과 경로를 확인하세요.";
            }
            catch (FileNotFoundException) { return null; }
            catch (DirectoryNotFoundException) { return null; }
            catch (Exception ex) when (IsFileError(ex)) { return ex.Message; }
        }
        try
        {
            using var handle = Open(path, DeleteAccess | ReadAttributes, false);
            var current = CheckHandle(handle, path, false);
            if (!owned.Value.SameIdentity(current)) return "대상 식별자가 달라 파일을 삭제하지 않았습니다.";
            // Handle-based deletion, including read-only copies, never path-based attribute mutation.
            uint flags = 0x1 | 0x2 | 0x10;
            if (!SetFileInformationByHandle(handle, 21, ref flags, sizeof(uint))) throw Error(Marshal.GetLastWin32Error());
            return null;
        }
        catch (Win32Exception ex) when (ex.NativeErrorCode is 2 or 3) { return null; }
        catch (Exception ex) when (IsFileError(ex)) { return ex.Message; }
    }

    internal static string LocalAppDataPath()
    {
        // Resolve OS package virtualization explicitly; never relax source/destination identity checks.
        var folder = new Guid("F1B32785-6FBA-4FCF-9D55-7B8E7F157091");
        int result = SHGetKnownFolderPath(ref folder, 0x00040000, IntPtr.Zero, out IntPtr value);
        if (result < 0) throw new IOException("LocalAppData 위치 확인 실패", Marshal.GetExceptionForHR(result));
        try { return Marshal.PtrToStringUni(value) ?? throw new IOException("LocalAppData 경로가 없습니다."); }
        finally { Marshal.FreeCoTaskMem(value); }
    }

    internal static bool IsFileError(Exception ex) => ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException or Win32Exception or System.Security.SecurityException;
    internal static string ErrorCode(Exception ex) => ex switch
    {
        CollectSafetyException safe => safe.Code,
        Win32Exception native => native.NativeErrorCode switch { 2 or 3 => "NotFound", 5 => "AccessDenied", 32 or 33 => "InUse", 80 or 183 => "DestinationExists", 112 => "DiskFull", _ => "Win32_" + native.NativeErrorCode },
        UnauthorizedAccessException => "AccessDenied", FileNotFoundException or DirectoryNotFoundException => "NotFound", _ => ex.GetType().Name
    };
    private static Exception Error(int code) => new Win32Exception(code);
    private static void EnsureWindows() { if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("안전한 파일 복사는 Windows에서만 지원합니다."); }
    private static string NativePath(string path) => path.StartsWith(@"\\", StringComparison.Ordinal) ? @"\\?\UNC\" + path[2..] : @"\\?\" + path;

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHGetKnownFolderPath(ref Guid folder, uint flags, IntPtr token, out IntPtr path);

    [StructLayout(LayoutKind.Sequential)]
    private struct HandleInformation
    {
        public uint Attributes;
        public System.Runtime.InteropServices.ComTypes.FILETIME CreationTime, LastAccessTime, LastWriteTime;
        public uint VolumeSerialNumber, SizeHigh, SizeLow, NumberOfLinks, FileIndexHigh, FileIndexLow;
    }
    private delegate uint ProgressRoutine(long total, long transferred, long streamSize, long streamTransferred, uint stream, uint reason, IntPtr source, IntPtr destination, IntPtr data);
    [DllImport("kernel32.dll", EntryPoint = "CreateFileW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateFile(string path, uint access, uint sharing, IntPtr security, uint disposition, uint flags, IntPtr template);
    [DllImport("kernel32.dll", SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetFileInformationByHandle(SafeFileHandle handle, out HandleInformation info);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern uint GetFileType(SafeFileHandle handle);
    [DllImport("kernel32.dll", EntryPoint = "GetFinalPathNameByHandleW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetFinalPathNameByHandle(SafeFileHandle handle, StringBuilder path, uint capacity, uint flags);
    [DllImport("kernel32.dll", EntryPoint = "CreateDirectoryW", CharSet = CharSet.Unicode, SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreateDirectory(string path, IntPtr security);
    [DllImport("kernel32.dll", EntryPoint = "CopyFileExW", CharSet = CharSet.Unicode, SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CopyFileEx(string source, string destination, ProgressRoutine progress, IntPtr data, IntPtr cancel, uint flags);
    [DllImport("kernel32.dll", EntryPoint = "GetVolumeInformationByHandleW", CharSet = CharSet.Unicode, SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetVolumeInformationByHandle(SafeFileHandle handle, StringBuilder? name, uint capacity, out uint serial, out uint componentLength, out uint flags, StringBuilder? filesystem, uint filesystemCapacity);
    [DllImport("kernel32.dll", SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetFileInformationByHandle(SafeFileHandle handle, int informationClass, ref uint value, uint size);
}
