# File Collect schema 1

## Workbook

Generated .xlsx remains macro-free. Short custom string properties: FLT.Product=FileListToExcel, FLT.ActionSchemaVersion=1, FLT.WorkbookId=nonempty UUID D format. FLT.Table.FileListTableN maps each existing supported table to files, duplicates or matches. Table names must use positive canonical decimal N; Errors/Summary/Skipped are unregistered.

Each copyable table appends hidden __FLT_ItemId and __FLT_SourceRecord columns. Each row UUID is also inside SourceRecord, so partial technical-column sorts/pastes mismatch. JSON fields are itemId, kind, absolutePath, sizeBytes, modifiedUtcTicks; sizes and UTC .NET ticks are strings, avoiding Excel floating-point loss. directory uses null size. unavailable uses empty path, null size, ticks "0" and a reason. Records exceeding 32,767 characters are never truncated: an unavailable record and FileCollectUnavailable error preserve the ordinary list.

These markers do not authenticate the author or grant filesystem access. Tampered-but-consistent records remain untrusted filesystem requests. The engine independently validates Windows paths, file kind, no-recall/reparse state, permissions and current metadata. Same size/time is not a historical content proof.

## Selection and current memory

The attached Excel instance supplies its current ActiveWorkbook, ActiveSheet and Selection. One registered table is required. DataBodyRange intersections map selected cells to table rows. Multiple/overlapping areas deduplicate row indices in sheet order; filters/manual Hidden rows are excluded. Full rows are clipped before traversal. Headers/totals alone and unrelated/multiple tables are rejected. Names locate fields after column moves. Save As and sheet rename preserve registration. No read of last-saved disk workbook substitutes for current memory; the workbook is not auto-saved.

## Request

--collect-request accepts exactly one product-owned Requests/<UUID>.json and no output/headless switches. Maximum 32 MiB, 10,000 rows, JSON depth 12. Schema:

```json
{
  "mode": "collect-files",
  "version": 1,
  "requestId": "UUID D matching the filename",
  "workbookId": "UUID D",
  "tableName": "FileListTable1",
  "rows": [{
    "itemId": "UUID D",
    "displayPath": "C:\\example\\file.txt",
    "displayName": "file.txt",
    "sourceRecord": "{...schema-1 JSON...}"
  }]
}
```

displayName is optional for the wire contract; the compiled UI emits it. Helper rejects duplicate/unknown JSON fields, inconsistent row UUID/path/name and unsupported versions. Request ancestors/file cannot be reparse/recall objects. Hardlinked request files are rejected. A validated request is deleted using its still-open handle; rejected requests are not consumed. No source strings are executed as formulas, shell commands, URLs or macros.

## Version/migration

Only schema 1 is supported. Old results without metadata and future versions require fresh generation; the add-in does not insert markers into arbitrary workbooks or guess based on filenames/headers. Renaming required Table/column identifiers or deleting technical fields disables the contract. A consistent hostile forgery is still possible and must pass the same defensive copy checks.

## Reports

Per-file statuses: Copied, Excluded, Failed, Cancelled, Duplicate. Local JSON reports use reportSchemaVersion=1 under LocalAppData/FileListToExcel/Reports, with source/destination/name mapping and reason. They are kept outside the delivery folder. The in-app result view reads these outcomes as literal strings. There is no upload, automatic opening of copied documents, automatic retry, recursive folder copy or destructive original operation.
