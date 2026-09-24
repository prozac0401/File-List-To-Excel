# Duplicate Finder opt-in performance check

Run from the repository root:

    .tools/dotnet/dotnet run --project tests/FileListToExcel.Benchmarks -c Release -- --run

This is intentionally outside the solution and ordinary tests. It writes 5,102 real, non-sparse files (about 4.21 GiB), including two files of 2 GiB + 64 KiB each. It tests size pruning, quick pruning, sampled collisions, full SHA-256, persistent cold/warm cache, enumeration cancellation, large-stream cancellation and released handles. It records elapsed times and sampled process working set to artifacts/duplicate-performance.json. Source fixtures remain under artifacts/benchmarks/<GUID>/source for UI validation; no user files are scanned or changed.

Optional arguments: --output <JSON path>, --fixture-root <new directory>. An existing fixture directory is rejected.
