# Duplicate Finder validation

Work order: `FileListToExcel_DuplicateFinder_WorkOrder.md` (all 25 sections).

## Baseline — before implementation

- Repository baseline: `187504855525126372a01a0ce3d7dc2543d80118`.
- Windows x64; .NET SDK 10.0.401; existing Core/App/WinForms/native COM/WiX structure retained.
- Actual baseline: Core 26 passed, App 9 passed; native shell 159 checks passed including request delivery.
- Evidence: `artifacts/duplicate-baseline.log`, `artifacts/duplicate-baseline/*.trx`.
- Existing file-list scanner reads metadata only; native shell uses per-user one-use JSON requests; writer streams OOXML; per-user MSI owns shell registration.

## Milestone evidence

Implementation and measured results are appended only after each milestone is exercised.

### Milestone 1 — comparison engine

15/15 new engine tests passed (no skips), `artifacts/test-results/milestone1.trx`. Same content/different names, same size/different contents, unique sizes, empty files, 1 MiB threshold, sampled collision, reference scope, missing/locked files, changed metadata, cancellation, cloud attribute gate and two-worker limit were exercised. Three actual 2 MiB files produced 3 size candidates → 2 quick candidates, 3 quick computations and 2 full computations, 4,784,128 content bytes read in 26.87 ms. Unique sizes read 0 bytes; three empty files grouped with 0 content bytes. Times are local fixture measurements, not guarantees.

### Milestone 2 — safe recursive scanning

7/7 scan tests and 6/6 safe-reader tests passed. Evidence: `artifacts/duplicate-milestones/milestone2-scan.trx`, `artifacts/test-results/M2-safe-reader.trx`. Real junction cycle/root, real ACL access-denied folder, overlapping roots, reference parent scope, selected-file scope, empty directory, 10 nested Unicode/long-path levels and scan cancellation were exercised. Read-only input, an already writable-open file, a real Offline attribute, ancestor junction rejection and cancellation during a 256 MiB streamed read were exercised. Native file open uses read access only, read sharing, OPEN_REPARSE_POINT and OPEN_NO_RECALL; opened-handle attributes and metadata are checked. Actual symbolic-link creation was denied by this host privilege; live OneDrive hydration was not tested.

### Platform references

The conservative cloud policy follows [Windows file attributes](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-fscc/ca28ec38-f155-4768-81d6-4bfeb8586fc9) and [CreateFile flags](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createfilew). All reparse files/directories are skipped in duplicate mode, including hydrated cloud files that still carry a reparse attribute. Ordinary file-list enumeration retains its original metadata-only cloud-directory policy.

### Milestone 3 — SQLite cache

16/16 cache and integration tests passed, `artifacts/test-results/milestone3-cache.trx`. Confirmed persistent cold/warm reuse, timestamp and size invalidation, same-size edits, deleted DB, corrupt DB and schema recovery, read/write contention, 100 concurrent stored rows, canceled/changed-file handling, cache-path junction refusal and excluding internal DB/sidecars from selected trees. Actual 3 × 2 MiB fixture: cold 31.05 ms (3 quick + 2 full calculations, 4,784,128 bytes read); warm 2.45 ms (3 cache-hit files, 0 hash calculations, 0 content bytes). Same-size changed file: one rehash + one cache hit, no false duplicate group. SQLite uses WAL, atomic upserts and a bounded one-second busy timeout with cacheless fallback.

### Milestone 4 — Excel output

8/8 duplicate workbook tests passed, `artifacts/test-results/milestone4-duplicates.trx`; all 26 existing CoreAcceptance tests also passed after the shared writer extension. Verified Open XML schema, numeric/date cells, Tables/AutoFilter/frozen headers, Duplicates name hyperlinks, Matches reference metadata/self exclusion, sorted groups, literal formula-like names, Errors/Skipped, empty results, 65,001-row splitting, cancellation cleanup and no-overwrite. The combined initial run contained a test-only expectation mismatch (4096 vs displayed 4 KB), corrected before the passing gate.

### Milestone 5 — Explorer and helper integration

31/31 App tests (9 existing + 22 new) passed, `artifacts/test-results/milestone5-app.trx`. Native Release build and 414 integration checks passed in 14.51 s, `artifacts/test-results/native-duplicates.xml`, including actual request sink dispatch for all selection cases, canonical Unicode/ANSI verbs, ZIP files, 10,000 selected paths and limited menu-ID capacity. Actual latest helper executions generated successful workbooks for --duplicates (4 rows), --matches (1 row), --duplicate-files (2 rows) and legacy --folder (5 rows), all exit 0/errors 0. Evidence `artifacts/milestone5-cli.json`.

### Milestone 6 — regression, performance and live UI

Six additional safety regression tests passed: cancellation during full hashing of two 128 MiB files, cache recovery after cancellation, locked/offline reference isolation, changed cached/zero-length files and stable overlapping-root ordering. Evidence: `artifacts/test-results/M6-safety-regressions.trx`.

The opt-in benchmark in `tests/FileListToExcel.Benchmarks` ran on DT-PROZAC, Windows 11 build 22631, .NET 10.0.12 and local NTFS. It created 5,102 files totaling 4,517,316,068 bytes: 5,000 unique small sizes, 100 × 2 MiB files (including sample collisions), and two identical **2,147,549,184-byte dense, non-sparse, non-compressed files**. Preparation took 144.02 s and is excluded from scan timings. Evidence: `artifacts/duplicate-performance.json`.

| Measurement | First scan | Cached scan |
| --- | ---: | ---: |
| Wall time including enumeration | 44,715.53 ms | 869.37 ms |
| Detector time | 42,462.65 ms | 382.91 ms |
| Size candidates → quick candidates | 102 → 6 | 102 → 6 |
| Quick / full hash computations | 102 / 6 | 0 / 0 |
| Cache-hit files | 0 | 102 |
| Content bytes read | 4,323,540,992 | 0 |
| Peak working set | 74,141,696 bytes | 75,517,952 bytes |

Both runs found 2 groups / 4 files, with 2,149,646,336 potentially recoverable bytes. At most two content workers were active. Source lengths, timestamps and attributes were unchanged. Enumeration stopped after 128 items (10.52 ms total elapsed); cancellation while both large streams were active took 24.21 ms, released handles and returned no result. Timings include OS file-cache effects and are not a guarantee of cold physical-disk performance.

Actual Excel Desktop opened Duplicates and Matches without a repair dialog. Inspected group ordering, typed size/date values, table/filter headers, reference metadata, reference self-exclusion and Summary. Clicking a Korean/hash/percent/emoji filename opened its expected original content in Notepad. The actual helper opened Excel and exited successfully.

A separate UI cancellation fixture used two sparse 16 GiB files solely to keep the progress window visible; these files were **not** used in the performance table. The displayed full-SHA-256 stage remained responsive; clicking Cancel closed the window, returned exit code 2 and left neither a completed workbook nor a partial file. Evidence: `artifacts/milestone6-ui.json`.

### Milestone 7 — installer and release verification

The clean Windows 2022 runner passed the strict full build in [PR run 35952108802](https://github.com/prozac0401/File-List-To-Excel/actions/runs/35952108802) and the matching branch run 35952103314. It ran Core 90/90 and App 31/31 with zero failures/skips, native integration, WiX validation with warnings as errors, actual per-user installation, installed old/new command exports, persistent SQLite reuse, source/workbook/cache retention, deleted-helper and corrupted-handler repair, uninstall and the synthetic 0.9.0 → 1.1.0 major upgrade. Native SQLite and all required SQLite legal files were verified in the installed payload. No helper process remained. The synthetic fixture contains current binaries in a lower-version MSI; it is not the published 1.0.0 application.

Local Release managed tests and 414 native checks passed and the MSI was built. The strict local installer test stopped when PowerShell 5.1 returned null for a context-handler registry lookup (Test-Installer.ps1 line 53); the local published-1.0 upgrade attempt stopped at the equivalent lookup before upgrade. Test installations were uninstalled. These are recorded as local failures, not counted as successful installer tests; strict checks passed unchanged on the clean runner. An additional CI gate downloads the published 1.0.0 MSI, checks its pinned SHA-256, and runs the same strict upgrade suite.

The tagged release [run 35952846805](https://github.com/prozac0401/File-List-To-Excel/actions/runs/35952846805) and main run 35952824762 both succeeded. The published **1.0.0 → 1.1.0** upgrade also passed, including sole current ProductCode, old/new exports, cache reuse, preserved source/workbook/cache and clean uninstall. [PR #1](https://github.com/prozac0401/File-List-To-Excel/pull/1) merged the application changes as `e231f33465f6fb9983bc8b40dd3c1228a3ead3d3`; tag v1.1.0 points to that commit. The public MSI was downloaded after publication and its 41,358,012 bytes verified against both SHA256SUMS and GitHub's asset digest: `14af126d51684e61f39a334db63a572c69926a016d043d217e8866ee05a33ab0`.

A separate local diagnostic compared PowerShell and explicit Registry64 reads while the package was installed. Both reported the file-handler key under `HKCU\Software\Classes\*\shellex\ContextMenuHandlers\FileListToExcel` absent, although MSI logged writing the correct path/CLSID. Folder/background/CLSID values agreed and were correct. This does not establish a PowerShell-provider bug; the local cause remains unresolved. The diagnostic installation was removed and both views confirmed no product keys/payload remained. Evidence: `artifacts/registry-installed-probe/f875337a5155482296abf9df2a03dbb0`. The unchanged strict checks passed on the clean release runner; this local failure remains a stated limitation.
