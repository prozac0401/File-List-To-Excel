using System.Buffers;
using System.Buffers.Binary;
using System.Security.Cryptography;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace FileListToExcel.Core;

/// <summary>Reads only source contents. Each worker owns a bounded pooled buffer and closes its stream on cancellation.</summary>
public sealed class FileContentHasher : IFileContentHasher
{
    public const int SmallFileLimit = 1024 * 1024;
    public const int QuickBlockSize = 64 * 1024;
    public const int StreamBufferSize = 4 * 1024 * 1024;
    public const string Algorithm = "SHA256/quick-v1-size-le64-first-middle-last-64KiB";

    public async Task<ContentHash> QuickAsync(string path, long sizeBytes, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (sizeBytes <= SmallFileLimit) return await FullAsync(path, cancellationToken).ConfigureAwait(false);
        await using var stream = Open(path, FileOptions.RandomAccess);
        var before = ReadStamp(stream, path);
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        byte[] size = new byte[sizeof(long)];
        BinaryPrimitives.WriteInt64LittleEndian(size, sizeBytes);
        hash.AppendData(size);
        byte[] buffer = ArrayPool<byte>.Shared.Rent(QuickBlockSize);
        long bytesRead = 0;
        try
        {
            foreach (long offset in new[] { 0L, (sizeBytes - QuickBlockSize) / 2, sizeBytes - QuickBlockSize })
            {
                cancellationToken.ThrowIfCancellationRequested();
                stream.Position = offset;
                await stream.ReadExactlyAsync(buffer.AsMemory(0, QuickBlockSize), cancellationToken).ConfigureAwait(false);
                hash.AppendData(buffer, 0, QuickBlockSize);
                bytesRead += QuickBlockSize;
            }
            VerifyUnchanged(stream, path, before);
            return new(Convert.ToHexString(hash.GetHashAndReset()), bytesRead);
        }
        finally { ArrayPool<byte>.Shared.Return(buffer); }
    }

    public async Task<ContentHash> FullAsync(string path, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        await using var stream = Open(path, FileOptions.SequentialScan);
        var before = ReadStamp(stream, path);
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        byte[] buffer = ArrayPool<byte>.Shared.Rent(StreamBufferSize);
        long bytesRead = 0;
        try
        {
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int count = await stream.ReadAsync(buffer.AsMemory(0, StreamBufferSize), cancellationToken).ConfigureAwait(false);
                if (count == 0) break;
                hash.AppendData(buffer, 0, count);
                bytesRead += count;
            }
            cancellationToken.ThrowIfCancellationRequested();
            VerifyUnchanged(stream, path, before);
            return new(Convert.ToHexString(hash.GetHashAndReset()), bytesRead);
        }
        finally { ArrayPool<byte>.Shared.Return(buffer); }
    }

    private static FileStream Open(string path, FileOptions options)
    {
        string absolutePath = Path.GetFullPath(path);
        // Check every ancestor too: OPEN_REPARSE_POINT applies only to the final path component.
        if (Path.GetDirectoryName(absolutePath) is { } parent && DuplicatePathPolicy.CheckDirectoryChain(parent) is { } issue)
            throw new IOException(issue.Message);
        CheckAttributes(absolutePath, File.GetAttributes(absolutePath));
        if (!OperatingSystem.IsWindows())
            return new FileStream(absolutePath, new FileStreamOptions
            {
                Mode = FileMode.Open, Access = FileAccess.Read, Share = FileShare.Read,
                BufferSize = 1, Options = options | FileOptions.Asynchronous
            });

        // Never follow a last-component link or request cloud recall. Recheck the opened handle
        // before the first content read, closing the metadata/open race for the leaf file.
        const uint GenericRead = 0x80000000, OpenExisting = 3;
        const uint OpenNoRecall = 0x00100000, OpenReparsePoint = 0x00200000, Overlapped = 0x40000000;
        SafeFileHandle handle = CreateFile(NativePath(absolutePath), GenericRead, (uint)FileShare.Read,
            IntPtr.Zero, OpenExisting, OpenNoRecall | OpenReparsePoint | Overlapped | (uint)options, IntPtr.Zero);
        if (handle.IsInvalid)
        {
            int error = Marshal.GetLastWin32Error();
            handle.Dispose();
            throw NativeError(error);
        }
        try
        {
            if (!GetFileInformationByHandle(handle, out var info)) throw NativeError(Marshal.GetLastWin32Error());
            CheckAttributes(absolutePath, (FileAttributes)info.Attributes);
            return new FileStream(handle, FileAccess.Read, 1, isAsync: true);
        }
        catch { handle.Dispose(); throw; }
    }

    private static void CheckAttributes(string path, FileAttributes attributes)
    {
        var reason = new LocalFileContentPolicy().GetSkipReason(path, attributes);
        if (reason is not null) throw new IOException(reason.Message);
    }

    private static string NativePath(string path)
    {
        if (path.StartsWith(@"\\.\", StringComparison.Ordinal)) throw new NotSupportedException("Device paths are not source files.");
        if (path.StartsWith(@"\\?\", StringComparison.Ordinal)) return path;
        return path.StartsWith(@"\\", StringComparison.Ordinal) ? @"\\?\UNC\" + path[2..] : @"\\?\" + path;
    }

    private static Exception NativeError(int error)
    {
        var cause = new Win32Exception(error);
        return error == 5 ? new UnauthorizedAccessException(cause.Message, cause) : new IOException(cause.Message, cause);
    }

    private static ContentStamp ReadStamp(FileStream stream, string path)
    {
        if (!OperatingSystem.IsWindows()) return new(stream.Length, File.GetLastWriteTimeUtc(path).Ticks);
        if (!GetFileInformationByHandle(stream.SafeFileHandle, out var info)) throw NativeError(Marshal.GetLastWin32Error());
        CheckAttributes(path, (FileAttributes)info.Attributes);
        return new(((long)info.SizeHigh << 32) | info.SizeLow,
            ((long)info.LastWriteTime.dwHighDateTime << 32) | (uint)info.LastWriteTime.dwLowDateTime);
    }

    private static void VerifyUnchanged(FileStream stream, string path, ContentStamp before)
    {
        if (ReadStamp(stream, path) != before)
            throw new FileChangedDuringScanException("검사 중 열린 파일의 크기 또는 수정시간이 변경되었습니다.");
    }

    private readonly record struct ContentStamp(long Length, long ModifiedTime);

    [StructLayout(LayoutKind.Sequential)]
    private struct HandleInformation
    {
        public uint Attributes;
        public System.Runtime.InteropServices.ComTypes.FILETIME CreationTime, LastAccessTime, LastWriteTime;
        public uint VolumeSerialNumber, SizeHigh, SizeLow, NumberOfLinks, FileIndexHigh, FileIndexLow;
    }

    [DllImport("kernel32.dll", EntryPoint = "CreateFileW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateFile(string fileName, uint desiredAccess, uint shareMode,
        IntPtr securityAttributes, uint creationDisposition, uint flagsAndAttributes, IntPtr templateFile);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetFileInformationByHandle(SafeFileHandle handle, out HandleInformation information);
}

internal sealed class FileChangedDuringScanException(string message) : IOException(message);
