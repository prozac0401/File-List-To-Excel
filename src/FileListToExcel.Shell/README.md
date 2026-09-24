# Explorer integration

This x64 native C++ COM DLL adds classic Explorer context-menu commands.
On Windows 11 the commands appear in **Show more options** (Shift+F10).
It does not load .NET, enumerate directory contents, or write a workbook inside
Explorer. Each invocation starts the adjacent helper once and returns.

## MSI registration

The MSI owns these per-user registry entries and removes them on uninstall:

- `HKCU\Software\Classes\CLSID\{AF7A7218-8B84-4A11-9790-EB24A43EC39E}\InprocServer32`
  - default: full installed path of `FileListToExcel.Shell.dll`
  - `ThreadingModel`: `Apartment`
- `HKCU\Software\Classes\*\shellex\ContextMenuHandlers\FileListToExcel`
- `HKCU\Software\Classes\Directory\shellex\ContextMenuHandlers\FileListToExcel`
- `HKCU\Software\Classes\Directory\Background\shellex\ContextMenuHandlers\FileListToExcel`

Each handler's default value is the CLSID above. There are deliberately no
`DllRegisterServer` exports, self-registration, installer custom actions, or
machine-wide registry changes.

## Selection and command behavior

`IShellExtInit` reads all selected paths from `CF_HDROP`. A background invocation
uses the supplied absolute filesystem PIDL. Virtual folders without a filesystem
path are excluded. Explorer's cached shell attributes classify selections when
available; `GetFileAttributesW` is a fallback for other shell hosts.

| Input | Existing list commands | Added duplicate command |
| --- | --- | --- |
| One file | Exact selected item (`files`) | 같은 파일 찾기 (`matches`) |
| Two or more files | Exact selected items (`files`) | 선택 파일 중 중복 찾기 (`duplicate-files`) |
| One or more folders | Immediate children (`folder`), then descendants (`recursive`) | 중복 파일 찾기 (`duplicates`) |
| Folder background | Immediate children (`folder`), then descendants (`recursive`) | 현재 폴더 중복 파일 찾기 (`duplicates`) |
| Mixed files/folders | Exact selected items (`files`), then selected files and folder descendants (`recursive`) | Hidden |

Existing list commands remain first, with their original canonical verbs. A
duplicate command is appended only when the selection is supported and Explorer
provides enough command identifiers. Numeric command offsets map to the visible
commands; ANSI and Unicode canonical invocations also validate the selection.

A matching-file request searches the selected file's parent folder recursively
and excludes the reference file from result rows. Folder duplicate requests
combine all selected folder trees into one scope. Multiple-file requests compare
only the selected files. All enumeration, hashing, caching, and workbook creation
run in the helper; the native extension does not read file contents.

ZIP files remain files even when Explorer marks them as browsable folders.
There is no per-selected-file process and no shell command interpreter.

## Helper protocol

The installed executable must be named `FileListToExcel.exe` beside the DLL.

```text
FileListToExcel.exe --request "<request path>"
```

The UTF-8 file (without BOM) contains exactly this model:

```json
{"mode":"files","paths":["C:\\Reports\\one.txt","C:\\Reports\\two.txt"]}
```

Modes are `files`, `folder`, `recursive`, `matches`, `duplicates`, and
`duplicate-files`. Duplicate commands use the same request format and security
rules as existing list commands; the original selection is passed unchanged.

| Duplicate request mode | Canonical verb |
| --- | --- |
| `matches` | `filelisttoexcelmatches` |
| `duplicates` | `filelisttoexcelduplicates` |
| `duplicate-files` | `filelisttoexcelduplicatefiles` |

The spool is
`%LOCALAPPDATA%\FileListToExcel\Requests\<GUID-D>.json` with a lowercase GUID,
without braces. The helper validates the request and deletes it after reading.

The shell uses `CoCreateGuid`, `CREATE_NEW`, no sharing during the write, and a
protected ACL limited to the current user and SYSTEM. It rejects a reparse-point
spool directory, limits requests to 100,000 input paths and 32 MiB, and deletes
the request if process creation fails. User-selected paths are JSON data, never
command-line fragments. The only process arguments are the fixed switch and
quoted request path. Explorer is free to exit after the process starts.

## Build and test

Visual Studio 2022 C++ build tools, a Windows SDK, and CMake 3.21 or later:

```powershell
cmake -S src/FileListToExcel.Shell -B artifacts/shell -A x64
cmake --build artifacts/shell --config Release
ctest --test-dir artifacts/shell -C Release --output-on-failure
```

The project also supports a Windows x64 LLVM-MinGW UCRT toolchain. The repository
build script can select its portable fallback. The native DLL statically links
the C++ runtime and uses Windows system DLLs; no separate VC++ runtime installer
is required.

`ShellTests.exe <DLL path>` directly loads the COM server without changing the
registry. It checks menu labels, command identifiers, localized help, ANSI and
Unicode canonical verbs, COM lifetime, single/multiple file and folder
selections, mixed selections, folder backgrounds, ZIP classification, command
identifier limits, Unicode paths, JSON limits, and argument quoting. Unsupported
duplicate canonical commands are rejected for mixed or incompatible selections.

`ShellTests.exe <DLL path> --dispatch` additionally starts the test-only adjacent
`FileListToExcel.exe` request sink and checks exact JSON delivery for existing
list commands and all three duplicate modes. Coverage includes single and
multiple folders, folder backgrounds, numeric/ANSI/Unicode duplicate
invocations, and 10,000 selected paths. CTest enables this mode. The v1.1.0
suite passes **414 integration checks** with dispatch enabled.

**The test sink must never ship in the MSI**; package only the native DLL plus
the separately published real helper.

The shell tests use temporary files under their own GUID-named test directory.
They do not open Excel or modify production registry entries.
