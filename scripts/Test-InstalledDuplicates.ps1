[CmdletBinding()]
param()

# Shared by clean-install and major-upgrade smoke tests. All source fixtures and
# workbook outputs belong to the caller's GUID-named artifacts directory.
function Invoke-InstalledWorkbook([string]$Helper, [string[]]$Arguments, [string]$OutputPath) {
    $allArguments = @($Arguments) + @('--no-open', '--output', $OutputPath)
    $quoted = @($allArguments | ForEach-Object {
        if ($_.Contains('"')) { throw 'Unexpected double quote in process argument.' }
        '"' + $_ + '"'
    })
    $process = Start-Process -FilePath $Helper -ArgumentList ($quoted -join ' ') -PassThru -WindowStyle Hidden
    try {
        if (-not $process.WaitForExit(60000)) {
            $process.Kill()
            $process.WaitForExit()
            throw 'Installed duplicate helper did not exit in 60 seconds.'
        }
        $process.Refresh()
        if ($process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
            throw "Installed duplicate helper failed: $OutputPath"
        }
    } finally { $process.Dispose() }
}

function Read-WorkbookXml($Archive, [string]$Part) {
    $entry = $Archive.GetEntry($Part)
    if ($null -eq $entry) { throw "Workbook is missing $Part" }
    $reader = New-Object IO.StreamReader($entry.Open())
    try { return [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Assert-InstalledDuplicateWorkbook([string]$Path, [string]$SheetName, [int]$ExpectedFiles, [int]$ExpectedCacheHits = -1) {
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $workbook = Read-WorkbookXml $zip 'xl/workbook.xml'
        $names = @($workbook.SelectNodes('//*[local-name()="sheet"]') | ForEach-Object { $_.GetAttribute('name') })
        if (($names -join ',') -ne "$SheetName,Summary,Errors,Skipped") { throw "Unexpected duplicate workbook sheets: $names" }
        $sheet = Read-WorkbookXml $zip 'xl/worksheets/sheet1.xml'
        $table = Read-WorkbookXml $zip 'xl/tables/table1.xml'
        $header = if ($SheetName -eq 'Matches') { 5 } else { 1 }
        $rows = @($sheet.SelectNodes('//*[local-name()="sheetData"]/*[local-name()="row"]') | Where-Object { [int]$_.GetAttribute('r') -gt $header })
        if ($rows.Count -ne $ExpectedFiles) { throw "Expected $ExpectedFiles duplicate result rows, found $($rows.Count)." }
        if ($null -eq $table.SelectSingleNode('//*[local-name()="autoFilter"]')) { throw 'Duplicate workbook has no table AutoFilter.' }
        $pane = $sheet.SelectSingleNode('//*[local-name()="pane"]')
        if ($null -eq $pane -or $pane.GetAttribute('state') -ne 'frozen') { throw 'Duplicate workbook has no frozen header.' }
        $links = @($sheet.SelectNodes('//*[local-name()="hyperlink"]'))
        if ($links.Count -ne $ExpectedFiles) { throw 'Duplicate workbook is missing file hyperlinks.' }
        $sizeColumn = if ($SheetName -eq 'Matches') { 'C' } else { 'F' }
        $dateColumn = if ($SheetName -eq 'Matches') { 'E' } else { 'H' }
        foreach ($row in $rows) {
            foreach ($column in @($sizeColumn, $dateColumn)) {
                $cell = $row.SelectSingleNode('*[local-name()="c" and @r="' + $column + $row.GetAttribute('r') + '"]')
                if ($null -eq $cell -or $cell.GetAttribute('t') -eq 'inlineStr' -or $null -eq $cell.SelectSingleNode('*[local-name()="v"]')) {
                    throw 'Duplicate size/date values are not numeric Excel values.'
                }
            }
        }
        if ($SheetName -eq 'Matches') {
            $reference = $sheet.SelectSingleNode('//*[local-name()="c" and @r="B1"]')
            if ($null -eq $reference -or [string]::IsNullOrWhiteSpace($reference.InnerText)) { throw 'Matches workbook has no reference metadata.' }
        }
        if ($ExpectedCacheHits -ge 0) {
            $summary = Read-WorkbookXml $zip 'xl/worksheets/sheet2.xml'
            # Summary rows 13 and 14 are full-hash computations and reused cache entries.
            $full = $summary.SelectSingleNode('//*[local-name()="c" and @r="B13"]/*[local-name()="v"]')
            $hits = $summary.SelectSingleNode('//*[local-name()="c" and @r="B14"]/*[local-name()="v"]')
            if ($null -eq $full -or $null -eq $hits -or [long]$full.InnerText -ne 0 -or [long]$hits.InnerText -ne $ExpectedCacheHits) {
                throw 'Installed SQLite cache did not eliminate repeated full hashes.'
            }
        }
    } finally { $zip.Dispose() }
}

function Test-InstalledDuplicateWorkbooks([string]$Helper, [string]$ArtifactsDirectory) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $installation = Split-Path -Parent $Helper
    foreach ($name in @('Microsoft.Data.Sqlite-LICENSE.txt', 'SQLitePCLRaw-LICENSE.txt', 'SQLitePCLRaw-NOTICE.txt', 'SQLite-PUBLIC-DOMAIN.html')) {
        $notice = Join-Path $installation ('licenses/sqlite/' + $name)
        if (-not (Test-Path -LiteralPath $notice -PathType Leaf)) { throw "Missing SQLite dependency notice: $notice" }
    }
    $nativeSqlite = @(Get-ChildItem -LiteralPath $installation -Filter 'e_sqlite3.dll' -File -Recurse)
    if ($nativeSqlite.Count -eq 0) { throw 'Installed payload is missing the native SQLite library.' }
    $inputDirectory = Join-Path $ArtifactsDirectory 'duplicate-input'
    $nested = Join-Path $inputDirectory 'nested'
    New-Item -ItemType Directory -Path $nested -Force | Out-Null
    $first = Join-Path $inputDirectory 'reference.txt'
    $second = Join-Path $inputDirectory 'different-name.txt'
    $third = Join-Path $nested 'copy-no-extension'
    foreach ($path in @($first, $second, $third)) { [IO.File]::WriteAllText($path, 'installed duplicate regression content') }
    [IO.File]::WriteAllText((Join-Path $inputDirectory 'unique.txt'), 'unique content')
    $sources = @(Get-ChildItem -LiteralPath $inputDirectory -File -Recurse | ForEach-Object {
        @{ Path = $_.FullName; Hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }
    })
    $duplicates = Join-Path $ArtifactsDirectory 'installed-duplicates.xlsx'
    $warm = Join-Path $ArtifactsDirectory 'installed-duplicates-cached.xlsx'
    $matches = Join-Path $ArtifactsDirectory 'installed-matches.xlsx'
    $selected = Join-Path $ArtifactsDirectory 'installed-selected-duplicates.xlsx'
    Invoke-InstalledWorkbook $Helper @('--duplicates', $inputDirectory) $duplicates
    Assert-InstalledDuplicateWorkbook $duplicates 'Duplicates' 3
    Invoke-InstalledWorkbook $Helper @('--duplicates', $inputDirectory) $warm
    Assert-InstalledDuplicateWorkbook $warm 'Duplicates' 3 3
    Invoke-InstalledWorkbook $Helper @('--matches', $first) $matches
    Assert-InstalledDuplicateWorkbook $matches 'Matches' 2
    Invoke-InstalledWorkbook $Helper @('--duplicate-files', $first, $second, '--no-cache') $selected
    Assert-InstalledDuplicateWorkbook $selected 'Duplicates' 2

    $cachePath = Join-Path $env:LOCALAPPDATA 'FileListToExcel/hash_cache.sqlite'
    if (-not (Test-Path -LiteralPath $cachePath -PathType Leaf)) { throw 'Installed helper did not create a persistent SQLite cache.' }
    foreach ($source in $sources) {
        if ((Get-FileHash -LiteralPath $source.Path -Algorithm SHA256).Hash -ne $source.Hash) { throw 'Duplicate inspection changed an input file.' }
    }
    $retained = @($sources) + @(@($duplicates, $warm, $matches, $selected, $cachePath) | ForEach-Object {
        @{ Path = $_; Hash = (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash }
    })
    Assert-NoInstalledHelper $Helper
    return ,$retained
}

function Assert-NoInstalledHelper([string]$Helper) {
    $expected = [IO.Path]::GetFullPath($Helper)
    foreach ($process in @(Get-Process -Name 'FileListToExcel' -ErrorAction SilentlyContinue)) {
        try {
            if ($process.Path -and [IO.Path]::GetFullPath($process.Path) -eq $expected) {
                throw "Installed helper remains running: $($process.Id)"
            }
        } finally { $process.Dispose() }
    }
}

function Assert-DuplicateTestRetention($RetainedFiles) {
    foreach ($item in $RetainedFiles) {
        if (-not (Test-Path -LiteralPath $item.Path -PathType Leaf) -or
            (Get-FileHash -LiteralPath $item.Path -Algorithm SHA256).Hash -ne $item.Hash) {
            throw "Maintenance removed or modified a source, workbook, or hash cache: $($item.Path)"
        }
    }
}
