using System.ComponentModel;
using System.Runtime.InteropServices;

namespace FileListToExcel.Core;

/// <summary>Reads file-system metadata only. Never opens file contents or follows directory links.</summary>
public sealed class FileScanner
{
    private static readonly EnumerationOptions Options = new()
    {
        AttributesToSkip = 0,
        IgnoreInaccessible = false,
        ReturnSpecialDirectories = false,
        RecurseSubdirectories = false
    };

    public IEnumerable<ScanItem> Enumerate(ScanRequest request, IProgress<ScanProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(request.Paths);
        if (!Enum.IsDefined(request.Mode)) throw new ArgumentOutOfRangeException(nameof(request));
        long count = 0, errors = 0;
        string current = string.Empty;
        var seen = new HashSet<string>(OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);
        foreach (string input in request.Paths)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? fullPath = null;
            ScanError? pathError = null;
            try
            {
                if (string.IsNullOrWhiteSpace(input)) throw new ArgumentException("The selected path is empty.");
                fullPath = Path.TrimEndingDirectorySeparator(Path.GetFullPath(input));
            }
            catch (Exception ex) when (IsPathError(ex)) { pathError = Error(input ?? string.Empty, ex); }
            if (pathError is not null)
            {
                errors++;
                progress?.Report(new(count, errors, input ?? string.Empty));
                yield return new(Error: pathError);
                continue;
            }
            if (!seen.Add(fullPath!)) continue;
            var rootItem = ReadEntry(fullPath!, Path.GetDirectoryName(fullPath!) ?? fullPath!, 1);
            if (rootItem.Error is not null)
            {
                errors++;
                progress?.Report(new(count, errors, fullPath!));
                yield return rootItem;
                continue;
            }
            if (request.Mode == ScanMode.SelectedItems || !rootItem.Entry!.IsDirectory)
            {
                count++;
                current = fullPath!;
                progress?.Report(new(count, errors, current));
                yield return rootItem;
                continue;
            }
            var stack = new Stack<(string Path, int Depth)>();
            stack.Push((fullPath!, 1));
            while (stack.TryPop(out var directory))
            {
                cancellationToken.ThrowIfCancellationRequested();
                current = directory.Path;
                progress?.Report(new(count, errors, current));
                // Check again immediately before traversal in case a directory was replaced during scanning.
                var traversalError = CheckTraversal(directory.Path);
                if (traversalError is not null)
                {
                    errors++;
                    yield return new(Error: traversalError);
                    continue;
                }
                IEnumerator<string>? children = null;
                ScanError? openError = null;
                try { children = Directory.EnumerateFileSystemEntries(directory.Path, "*", Options).GetEnumerator(); }
                catch (Exception ex) when (IsPathError(ex)) { openError = Error(directory.Path, ex); }
                if (openError is not null)
                {
                    errors++;
                    yield return new(Error: openError);
                    continue;
                }
                using (children)
                {
                    while (true)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        string? child = null;
                        ScanError? nextError = null;
                        try { if (children!.MoveNext()) child = children.Current; }
                        catch (Exception ex) when (IsPathError(ex)) { nextError = Error(directory.Path, ex); }
                        if (nextError is not null)
                        {
                            errors++;
                            yield return new(Error: nextError);
                            break;
                        }
                        if (child is null) break;
                        var item = ReadEntry(child, fullPath!, directory.Depth);
                        current = child;
                        if (item.Error is not null) errors++; else count++;
                        if ((count + errors) % 64 == 0) progress?.Report(new(count, errors, current));
                        yield return item;
                        if (request.Mode == ScanMode.Recursive && item.Entry?.IsDirectory == true)
                            stack.Push((child, directory.Depth + 1));
                    }
                }
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        progress?.Report(new(count, errors, current));
    }

    private static ScanItem ReadEntry(string path, string root, int depth)
    {
        try
        {
            var attributes = File.GetAttributes(path);
            bool directory = (attributes & FileAttributes.Directory) != 0;
            FileSystemInfo info = directory ? new DirectoryInfo(path) : new FileInfo(path);
            // Explicit refresh ensures a deleted entry produces an error instead of zero/default metadata.
            info.Refresh();
            if (!info.Exists) throw new FileNotFoundException("The selected item no longer exists.", path);
            return new(new FileEntry(info.Name, directory ? "폴더" : info.Extension.TrimStart('.'), directory,
                directory ? null : ((FileInfo)info).Length, info.CreationTime, info.LastWriteTime,
                Path.GetFileName(Path.GetDirectoryName(path)) ?? string.Empty, path,
                Path.GetRelativePath(root, path), depth, info.Attributes, root));
        }
        catch (Exception ex) when (IsPathError(ex)) { return new(Error: Error(path, ex)); }
    }

    private static ScanError? CheckTraversal(string path)
    {
        try
        {
            var attributes = File.GetAttributes(path);
            if ((attributes & FileAttributes.ReparsePoint) == 0) return null;
            bool nameSurrogate;
            if (OperatingSystem.IsWindows())
            {
                // FindFirstFile reports the reparse tag without opening/downloading cloud file contents.
                IntPtr handle = FindFirstFile(path, out var data);
                if (handle == new IntPtr(-1)) throw new Win32Exception(Marshal.GetLastWin32Error());
                try { nameSurrogate = (data.Reserved0 & 0x20000000) != 0; }
                finally { FindClose(handle); }
            }
            else nameSurrogate = new DirectoryInfo(path).LinkTarget is not null;
            return nameSurrogate ? new(path, "LinkNotTraversed", "Directory symbolic links and junctions are listed but are not traversed, to avoid cycles and leaving the selected tree.") : null;
        }
        catch (Exception ex) when (IsPathError(ex) || ex is Win32Exception) { return Error(path, ex); }
    }

    private static bool IsPathError(Exception ex) => ex is IOException or UnauthorizedAccessException or
        ArgumentException or NotSupportedException or System.Security.SecurityException;

    private static ScanError Error(string path, Exception ex) => new(path, ex.GetType().Name, ex.Message);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct FindData
    {
        public uint Attributes;
        public System.Runtime.InteropServices.ComTypes.FILETIME CreationTime, LastAccessTime, LastWriteTime;
        public uint SizeHigh, SizeLow, Reserved0, Reserved1;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string FileName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)] public string AlternateFileName;
    }

    [DllImport("kernel32.dll", EntryPoint = "FindFirstFileW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr FindFirstFile(string fileName, out FindData findData);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool FindClose(IntPtr handle);
}
