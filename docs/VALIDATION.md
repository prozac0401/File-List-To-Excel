# Validation for v1.0.0

## Release gate

The Windows CI workflow runs the same scripts as local development. Publication is gated on managed tests, native integration tests, MSI validation, install/repair/uninstall, and a major-upgrade test. A release tag must point to a commit already contained in main. Test evidence and installer logs are uploaded as workflow artifacts.

The suite contains **26 core tests**, **9 application tests**, and **159 native integration assertions**. Counts refer to this release, not an assurance of exhaustive coverage.

## Automated coverage

| Area | Checks |
| --- | --- |
| Selection | Exact chosen files; unselected siblings excluded; direct children; recursive folders; multiple roots; mixed file/folder selections; folder background PIDL |
| Metadata | Numeric sizes and Excel date values; creation/modified times; hidden/read-only attributes; Unicode names and paths longer than 260 characters |
| Reliability | Deleted-file race, real access-denied ACL fixture, junction cycle prevention, exclusive source-content lock (metadata remains readable), cancellation, existing-output protection |
| Excel | Open XML schema validation; table/filter/frozen header; hyperlinks; formula-like names kept as literals; Excel date boundary; 65,001-row sheet split |
| File URI edge cases | Korean, emoji, hash and percent combinations; literal percent-looking filenames; URI/path round trips; overlong-link error recording |
| Scale | 10,000 actual files with progress; 10,000 paths transferred in one native helper invocation |
| Request boundary | Owned request directory and GUID name; consumed-once request; invalid mode/path/option rejection; external request files never deleted |
| Native shell | COM lifetime; real Explorer-style data object; ZIP treated as a file; menu selection; UTF-8 JSON; command-line quoting; helper dispatch |
| MSI | ICE validation with warnings as errors; payload and registry ownership; native-x64 launch condition; installed helper export; deleted-helper and corrupted-registration repair; uninstall; workbook preservation |
| Upgrade | Synthetic 0.9.0 MSI → 1.0.0 major upgrade; ProductCode replacement; current version registration; helper execution; uninstall; preservation of workbooks by SHA-256 |

The native test harness uses a distinct capture file for each child invocation, waits for complete readable output and reports payload differences. It does not silently retry product assertions.

The upgrade fixture intentionally uses the current application payload in a lower-version MSI. It validates Windows Installer upgrade behavior, not migration from a previously shipped application or persisted user database.

## Interactive Windows 11 + Excel Desktop verification

- A directory contained 20 files; selecting 3 produced exactly 3 rows in Excel.
- Excel opened the workbook without a repair prompt and exposed a real table, filter headers, frozen first row, typed size/date cells and hyperlinks.
- Korean filenames opened their actual original file from Excel.
- The combined filename `한글 #10% 😀.txt` opened the expected source content through Excel's hyperlink command.
- A live scan exceeded 40,000 entries while the progress window remained responsive. Clicking Cancel closed the window, returned exit code 2, left no finished workbook or staging file, and did not open Excel.
- The helper exited after successfully opening Excel; no product background service or tray process is installed.

These checks caught and fixed an Excel-specific UTF-8 hyperlink decoding problem and a percent/emoji path conversion failure that schema validation alone did not catch.

## Reproduction

    dotnet test FileListToExcel.sln -c Release
    powershell -NoProfile -ExecutionPolicy Bypass -File scripts/Build.ps1 -TestInstaller

Use a clean Windows account or VM: installer tests deliberately damage only the test installation and refuse to touch a pre-existing product installation. Normal source scripts keep strict registry-repair checks enabled.

Local development observed an unexplained mismatch between a successful MSI registry-repair operation in the installer log and the registry value read by the test host after deliberately corrupting it. A diagnostic clone omitted only that corruption step; it is not a release gate. The unchanged strict test on the clean GitHub Windows runner is required before release.

## Scope not established by these checks

- Live cloud-only OneDrive accounts and remote SMB servers with real latency, disconnection and organization policies.
- Every Excel version, Office security policy, regional setting, Windows display scale or assistive technology.
- Interactive installation and fallback messages on a PC with no Excel or .xlsx-associated app.
- Concurrent installations by multiple Windows users.
- Authenticode signing/SmartScreen reputation: no publisher certificate was supplied, so release binaries are unsigned.
- ARM64 Explorer is unsupported and blocked by the MSI.
