# Architecture

The original file-list specification remains unchanged for list commands. Version 1.1 adds the duplicate-finder work order as a separate operation; no accounts, resident service, deletion or custom result viewer are introduced.

1. Native COM DLL implements IShellExtInit/IContextMenu in the classic Explorer menu. It receives the full CF_HDROP selection or a directory-background PIDL.
2. A single, uniquely named JSON request is created in the user's LocalAppData/FileListToExcel/Requests directory. One helper invocation receives its path, avoiding command-line limits and per-file launches.
3. The helper validates the owned request directory, GUID filename, size, mode and absolute paths, rejects reparse request files, and consumes the request once.
4. FileScanner produces a lazy stream of metadata rows/errors. No source content is opened. Directory link traversal is skipped, cloud reparse tags distinguished, and cancellation checked between operations.
5. WorkbookWriter streams ZIP/XML, encodes source strings literally, creates typed cells, links and Excel tables, and splits large sheets. A completed workbook is atomically published; cancellation removes unfinished output.
6. A WinForms message loop displays progress only after 700ms. The worker handles enumeration/writing; the UI remains responsive. Excel is launched after completion and the helper exits.

MSI installs per user under LocalAppData/Programs/FileListToExcel with HKCU COM and shell registrations. UpgradeCode remains stable; MSI owns all registration removal. No custom registration executables run during installation. Source files and generated workbooks are never MSI-owned resources.

The product uses .NET 10 LTS with self-contained Windows x64 deployment. Windows 11 modern context menu integration and ARM64 are outside this release. The classic menu is an explicitly accepted specification option.

Safety boundaries: never follow arbitrary request-file paths for deletion; never overwrite a chosen output; never interpret file names as Excel formulas; no source file modifications; no telemetry; diagnostics remain local. Enumeration may expose file paths in the workbook and local error log as required by the feature.

## Duplicate operation

The same COM selection/request path carries `matches`, `duplicates`, or `duplicate-files`. `DuplicateScanService` resolves the selection and reuses `FileScanner` with an opt-in policy that excludes every reparse directory and its ancestors. The original scanner default and list commands retain their metadata-only behavior.

`DuplicateDetector` groups metadata by size, computes a versioned SHA-256 fingerprint over length and first/middle/last 64 KiB for large candidates, and computes full streaming SHA-256 only for surviving groups. Candidates at most 1 MiB use one full hash; empty files require no content reads. Up to two workers own pooled 4 MiB buffers; UNC selections use one worker. Pre/post size and UTC timestamp checks reject changed sources. Native read-only handles refuse writers, reparse processing and recall; handle metadata is checked before content reads and afterward.

`SqliteHashCache` stores the current size, 100 ns UTC modification ticks, algorithm version, quick/full hashes and check time for each normalized path. Connections serialize access, use WAL/atomic upserts and a one-second contention timeout. Cache failure never prevents comparison. Invalid caches are quarantined; future schema versions are left intact and ignored. Cache/sidecar paths are excluded from source candidates. This is metadata-based cache reuse, not filesystem snapshot or tamper detection.

`WorkbookWriter.WriteDuplicates` adds Duplicates/Matches/Summary/Errors/Skipped using the existing SheetBuilder, OOXML package, types, escaping, styles, filters and links. Optional header offsets support reference metadata and configurable hyperlink columns; existing list defaults are unchanged. Output uses the same atomic staging/no-overwrite pattern. WinForms uses the existing delayed window/message loop and cancellation path.

## Optional Excel file collection (v1.2)

The new user-authorized optional COM integration supersedes the original add-in prohibition only for this feature. The standalone list/duplicate helper remains independent and produces macro-free workbooks. See [ADR 0003](adr/0003-excel-file-collect.md), [schema](FILE_COLLECT_SCHEMA.md), and [validation](FILE_COLLECT_VALIDATION.md). A native x86/x64 Excel shim installs temporary owned CommandBars menus, captures the connected instance’s current visible table rows, and sends a bounded collect-files request to the existing helper. Safe Win32 copy planning and progress/cancellation run outside Excel. Local reports stay outside the delivery folder.
