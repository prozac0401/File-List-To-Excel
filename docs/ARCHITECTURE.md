# Architecture

The v1 specification is the product boundary. No content analysis, accounts, resident service or custom result viewer is introduced.

1. Native COM DLL implements IShellExtInit/IContextMenu in the classic Explorer menu. It receives the full CF_HDROP selection or a directory-background PIDL.
2. A single, uniquely named JSON request is created in the user's LocalAppData/FileListToExcel/Requests directory. One helper invocation receives its path, avoiding command-line limits and per-file launches.
3. The helper validates the owned request directory, GUID filename, size, mode and absolute paths, rejects reparse request files, and consumes the request once.
4. FileScanner produces a lazy stream of metadata rows/errors. No source content is opened. Directory link traversal is skipped, cloud reparse tags distinguished, and cancellation checked between operations.
5. WorkbookWriter streams ZIP/XML, encodes source strings literally, creates typed cells, links and Excel tables, and splits large sheets. A completed workbook is atomically published; cancellation removes unfinished output.
6. A WinForms message loop displays progress only after 700ms. The worker handles enumeration/writing; the UI remains responsive. Excel is launched after completion and the helper exits.

MSI installs per user under LocalAppData/Programs/FileListToExcel with HKCU COM and shell registrations. UpgradeCode remains stable; MSI owns all registration removal. No custom registration executables run during installation. Source files and generated workbooks are never MSI-owned resources.

The product uses .NET 10 LTS with self-contained Windows x64 deployment. Windows 11 modern context menu integration and ARM64 are outside this release. The classic menu is an explicitly accepted specification option.

Safety boundaries: never follow arbitrary request-file paths for deletion; never overwrite a chosen output; never interpret file names as Excel formulas; no source file modifications; no telemetry; diagnostics remain local. Enumeration may expose file paths in the workbook and local error log as required by the feature.
