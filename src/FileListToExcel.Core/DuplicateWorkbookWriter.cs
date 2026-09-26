using System.Globalization;
using System.IO.Compression;

namespace FileListToExcel.Core;

public sealed partial class WorkbookWriter
{
    private static readonly string[] DuplicateHeaders = ["그룹", "그룹 파일수", "중복 절약 가능 크기", "이름", "종류", "크기", "크기(표시)", "수정일", "상대경로", "전체경로", "SHA-256", FileCollectContract.ItemIdColumn, FileCollectContract.SourceRecordColumn];
    private static readonly string[] MatchHeaders = ["이름", "종류", "크기", "크기(표시)", "수정일", "상대경로", "전체경로", "SHA-256", FileCollectContract.ItemIdColumn, FileCollectContract.SourceRecordColumn];

    public ExportResult WriteDuplicates(string destinationPath, DuplicateResult result, bool matches = false,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);
        ArgumentNullException.ThrowIfNull(result);
        cancellationToken.ThrowIfCancellationRequested();
        string destination = Path.GetFullPath(destinationPath);
        string parent = Path.GetDirectoryName(destination)!;
        Directory.CreateDirectory(parent);
        string staging = Path.Combine(parent, $".{Path.GetFileName(destination)}.{Guid.NewGuid():N}.partial");
        var errors = result.Errors.ToList();
        long rowCount = 0;
        int resultSheets = 0;
        var comparer = OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
        try
        {
            using (var output = new FileStream(staging, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create))
            {
                var sheets = new List<SheetInfo>();
                IReadOnlyList<(string Label, string Value)>? metadata = matches ?
                [
                    ("기준 파일", result.ReferenceFile?.Entry.AbsolutePath ?? "검사할 수 없음 — Errors/Skipped 확인"),
                    ("크기", result.ReferenceFile?.Entry.SizeBytes is { } size ? HumanSize(size) : "알 수 없음"),
                    ("SHA-256", result.ReferenceFile?.FullSha256 ?? "미계산 (비교 후보 없음 또는 검사 불가)")
                ] : null;
                SheetBuilder NextResultSheet()
                {
                    resultSheets++;
                    string name = matches ? "Matches" : "Duplicates";
                    var info = new SheetInfo(sheets.Count + 1, resultSheets == 1 ? name : $"{name}_{resultSheets}", matches ? MatchHeaders : DuplicateHeaders, matches ? "matches" : "duplicates");
                    sheets.Add(info);
                    return new SheetBuilder(archive, info,
                        matches ? [36, 12, 18, 16, 22, 45, 70, 68, 38, 80] : [10, 14, 24, 36, 12, 18, 16, 22, 45, 70, 68, 38, 80],
                        matches ? 0 : 3, metadata);
                }
                SheetBuilder? sheet = null;
                try
                {
                    sheet = NextResultSheet();
                    foreach (var group in result.Groups.OrderByDescending(group => group.RecoverableBytes).ThenBy(group => group.GroupId))
                    foreach (var file in group.Files.OrderBy(file => file.Entry.RelativePath, comparer).ThenBy(file => file.Entry.AbsolutePath, comparer))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (matches && comparer.Equals(file.Entry.AbsolutePath, result.ReferenceFile?.Entry.AbsolutePath)) continue;
                        if (sheet.Rows == RowsPerSheet) { sheet.Dispose(); sheet = NextResultSheet(); }
                        var link = FileLink(file.Entry.AbsolutePath);
                        if (link is null) errors.Add(new(file.Entry.AbsolutePath, "HyperlinkUnavailable", "전체경로는 보존했지만 Excel 하이퍼링크로 표현할 수 없습니다."));
                        string itemId = Guid.NewGuid().ToString("D");
                        string sourceRecord = FileCollectContract.CreateSourceRecord(file.Entry, itemId, out string? unavailableReason);
                        if (unavailableReason is not null) errors.Add(new(file.Entry.AbsolutePath, "FileCollectUnavailable", $"선택 파일 복사 불가: {unavailableReason}. 기존 목록 조회는 사용할 수 있습니다."));
                        sheet.WriteDuplicate(group, file, link, matches, itemId, sourceRecord);
                        rowCount++;
                    }
                }
                finally { sheet?.Dispose(); }

                var summaryInfo = new SheetInfo(sheets.Count + 1, "Summary", ["항목", "값"]);
                sheets.Add(summaryInfo);
                using (var summary = new SheetBuilder(archive, summaryInfo, [34, 90]))
                {
                    summary.WriteValues(["검사 파일수", result.Stats.ScannedFiles]);
                    summary.WriteValues(["중복 그룹", result.Groups.Count]);
                    summary.WriteValues([matches ? "발견한 동일 파일" : "중복 그룹 내 파일수", matches ? rowCount : result.DuplicateFileCount]);
                    summary.WriteValues(["예상 중복 용량 (byte)", result.RecoverableBytes]);
                    summary.WriteValues(["예상 중복 용량", HumanSize(result.RecoverableBytes)]);
                    summary.WriteValues(["용량 설명", "각 그룹에서 한 파일을 남긴다고 가정한 이론적 값입니다. 삭제 권고가 아닙니다."]);
                    summary.WriteValues(["건너뛴 항목", result.Skipped.Count]);
                    summary.WriteValues(["오류", errors.Count]);
                    summary.WriteValues(["크기 후보", result.Stats.SizeCandidates]);
                    summary.WriteValues(["Quick Fingerprint 통과 후보", result.Stats.QuickCandidates]);
                    summary.WriteValues(["Quick Fingerprint 계산 횟수", result.Stats.QuickHashCount]);
                    summary.WriteValues(["전체 SHA-256 계산 횟수", result.Stats.FullHashCount]);
                    summary.WriteValues(["캐시 재사용 파일수", result.Stats.CacheHitCount]);
                    summary.WriteValues(["원본 읽기량 (byte)", result.Stats.BytesRead]);
                    summary.WriteValues(["검사 시간 (초)", result.Stats.Elapsed.TotalSeconds]);
                    summary.WriteValues(["비교 방법", "크기 → Quick Fingerprint → SHA-256. 1 MiB 이하는 전체 SHA-256을 한 번 계산합니다."]);
                    summary.WriteValues(["안전 정책", "원본 변경 없음. Cloud-only 및 Reparse Point는 제외합니다."]);
                }
                WriteIssues("Errors", errors);
                WriteIssues("Skipped", result.Skipped);
                cancellationToken.ThrowIfCancellationRequested();
                WriteStyles(archive);
                WritePackage(archive, sheets);

                void WriteIssues(string name, IReadOnlyList<ScanError> issues)
                {
                    int part = 0;
                    SheetBuilder? builder = null;
                    try
                    {
                        do
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            var info = new SheetInfo(sheets.Count + 1, part == 0 ? name : $"{name}_{part + 1}", ["경로", "구분", "내용"]);
                            sheets.Add(info);
                            builder = new SheetBuilder(archive, info, [70, 28, 100]);
                            foreach (var issue in issues.Skip(part * RowsPerSheet).Take(RowsPerSheet))
                            {
                                cancellationToken.ThrowIfCancellationRequested();
                                builder.WriteError(issue);
                            }
                            builder.Dispose(); builder = null;
                            part++;
                        } while ((long)part * RowsPerSheet < issues.Count);
                    }
                    finally { builder?.Dispose(); }
                }
            }
            cancellationToken.ThrowIfCancellationRequested();
            File.Move(staging, destination, overwrite: false);
            return new(destination, rowCount, errors.Count, resultSheets);
        }
        finally
        {
            try { if (File.Exists(staging)) File.Delete(staging); } catch (IOException) { }
        }
    }

    private sealed partial class SheetBuilder
    {
        public void WriteDuplicate(DuplicateGroup group, DuplicateFile file, string? link, bool matches, string itemId, string sourceRecord)
        {
            var entry = file.Entry;
            object[] values = matches
                ? [entry.Name, entry.Extension, entry.SizeBytes ?? 0, HumanSize(entry.SizeBytes ?? 0), entry.ModifiedAt, entry.RelativePath, entry.AbsolutePath, file.FullSha256 ?? ""]
                : [group.GroupId, group.Files.Count, group.RecoverableBytes, entry.Name, entry.Extension, entry.SizeBytes ?? 0,
                    HumanSize(entry.SizeBytes ?? 0), entry.ModifiedAt, entry.RelativePath, entry.AbsolutePath, file.FullSha256 ?? ""];
            WriteValues([.. values, itemId, sourceRecord], link);
        }

        public void WriteValues(object[] values, string? link = null)
        {
            int row = ++Rows + headerRow;
            writer.WriteStartElement("row", Main); writer.WriteAttributeString("r", row.ToString(CultureInfo.InvariantCulture));
            for (int column = 0; column < values.Length; column++)
            {
                string cell = Cell(column, row);
                switch (values[column])
                {
                    case DateTime date: DateCell(writer, cell, date); break;
                    case long number: NumberCell(writer, cell, number, 1); break;
                    case int number: NumberCell(writer, cell, number, 1); break;
                    case double number: NumberCell(writer, cell, number); break;
                    default: TextCell(writer, cell, values[column].ToString() ?? "", column == linkColumn && link is not null ? 3 : 0); break;
                }
            }
            writer.WriteEndElement();
            links.Add(link ?? "");
        }
    }
}
