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

| Input | First command | Second command |
| --- | --- | --- |
| Files only | Exact selected items (`files`) | None |
| Folders only | Immediate children (`folder`) | Descendants (`recursive`) |
| Folder background | Immediate children (`folder`) | Descendants (`recursive`) |
| Mixed files/folders | Exact selected items (`files`) | Selected files and folder descendants (`recursive`) |

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

Modes are `files`, `folder`, and `recursive`. The spool is
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
registry, verifies menu labels, IDs, canonical verbs, COM lifetime, file/folder/
mixed/background selections, Unicode, JSON limits, and argument quoting.

`ShellTests.exe <DLL path> --dispatch` additionally starts the test-only adjacent
`FileListToExcel.exe` request sink and verifies exact JSON delivery for seven
selection scenarios, including 10,000 selected paths. CTest enables this mode. **The test sink must never ship in
the MSI**; package only the native DLL plus the separately published real helper.

The shell tests use temporary files under their own GUID-named test directory.
They do not open Excel or modify production registry entries.
