[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$MsiPath)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if (-not [Environment]::Is64BitProcess) { throw 'Use 64-bit PowerShell for installer tests.' }
$repo = Split-Path -Parent $PSScriptRoot
$MsiPath = (Resolve-Path -LiteralPath $MsiPath).Path
$installDirectory = Join-Path $env:LOCALAPPDATA 'Programs/FileListToExcel'
$clsid = '{AF7A7218-8B84-4A11-9790-EB24A43EC39E}'
$registryPaths = @(
    'HKCU:\Software\FileListToExcel',
    "HKCU:\Software\Classes\CLSID\$clsid",
    'HKCU:\Software\Classes\*\shellex\ContextMenuHandlers\FileListToExcel',
    'HKCU:\Software\Classes\Directory\shellex\ContextMenuHandlers\FileListToExcel',
    'HKCU:\Software\Classes\Directory\Background\shellex\ContextMenuHandlers\FileListToExcel'
)
# Never upgrade, remove, or overwrite an existing installation during a smoke test.
if (Test-Path -LiteralPath $installDirectory) { throw "An existing installation is present at $installDirectory. Run installer tests in a clean Windows account or VM." }
foreach ($key in $registryPaths) {
    if (Test-Path -LiteralPath $key) { throw "Existing product registration found at $key. Use a clean Windows account or VM." }
}
$windowsInstaller = New-Object -ComObject WindowsInstaller.Installer
$database = $windowsInstaller.OpenDatabase($MsiPath, 0)
$view = $database.OpenView("SELECT Value FROM Property WHERE Property = 'ProductCode'")
$view.Execute()
$record = $view.Fetch()
$productCode = $record.StringData(1)
$view.Close()
foreach ($comObject in @($record, $view, $database, $windowsInstaller)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) }
$logs = Join-Path $repo ('artifacts/installer-smoke/' + [Guid]::NewGuid().ToString('N'))
$fixture = Join-Path $logs 'input'
New-Item -ItemType Directory -Path $fixture -Force | Out-Null
[IO.File]::WriteAllText((Join-Path $fixture 'one.txt'), 'one')
[IO.File]::WriteAllText((Join-Path $fixture 'two with spaces.txt'), 'two')
$output = Join-Path $logs 'installed-output.xlsx'
$installed = $false
function Run-Msi([string[]]$Arguments, [string]$Log) {
    $quoted = @($Arguments | ForEach-Object { if ($_ -match '[\s"]') { '"' + $_.Replace('"', '\"') + '"' } else { $_ } })
    $process = Start-Process -FilePath (Join-Path $env:WINDIR 'System32/msiexec.exe') -ArgumentList ($quoted -join ' ') -PassThru -Wait -WindowStyle Hidden
    if ($process.ExitCode -notin @(0, 3010)) { throw "Windows Installer failed with exit code $($process.ExitCode). See $Log" }
}
try {
    $installLog = Join-Path $logs 'install.log'
    Run-Msi @('/i', $MsiPath, '/qn', '/norestart', '/L*V', $installLog) $installLog
    $installed = $true
    $helper = Join-Path $installDirectory 'FileListToExcel.exe'
    $shell = Join-Path $installDirectory 'FileListToExcel.Shell.dll'
    foreach ($file in @($helper, $shell)) { if (-not (Test-Path -LiteralPath $file)) { throw "Missing installed payload: $file" } }
    if ((Get-Item -LiteralPath "HKCU:\Software\Classes\CLSID\$clsid\InprocServer32").GetValue('') -ne $shell) { throw 'COM registration has the wrong DLL path.' }
    if ((Get-Item -LiteralPath "HKCU:\Software\Classes\CLSID\$clsid\InprocServer32").GetValue('ThreadingModel') -ne 'Apartment') { throw 'Invalid COM threading model.' }
    foreach ($key in $registryPaths[2..4]) { if ((Get-Item -LiteralPath $key).GetValue('') -ne $clsid) { throw "Wrong handler CLSID: $key" } }
    $shellTest = Join-Path $repo 'artifacts/shell/Release/ShellTests.exe'
    if (-not (Test-Path -LiteralPath $shellTest)) { $shellTest = Join-Path $repo 'artifacts/shell/ShellTests.exe' }
    if (-not (Test-Path -LiteralPath $shellTest)) { throw 'Build native ShellTests before running installer tests.' }
    & $shellTest $shell
    if ($LASTEXITCODE -ne 0) { throw 'The installed native shell integration failed its tests.' }
    $helperArguments = '--folder "' + $fixture + '" --no-open --output "' + $output + '"'
    $helperProcess = Start-Process -FilePath $helper -ArgumentList $helperArguments -PassThru -WindowStyle Hidden
    if (-not $helperProcess.WaitForExit(60000)) { $helperProcess.Kill(); throw 'Installed helper did not exit in 60 seconds.' }
    $helperProcess.Refresh()
    if ($helperProcess.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $output)) { throw 'Installed helper failed to create an Excel workbook.' }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($output)
    try {
        foreach ($part in @('[Content_Types].xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/tables/table1.xml')) {
            if ($null -eq $zip.GetEntry($part)) { throw "Workbook is missing $part" }
        }
        $reader = New-Object IO.StreamReader($zip.GetEntry('xl/worksheets/sheet1.xml').Open())
        try { [xml]$sheet = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $rows = $sheet.SelectNodes('//*[local-name()="sheetData"]/*[local-name()="row"]')
        if ($rows.Count -ne 3) { throw "Expected two file rows and one header; found $($rows.Count) rows." }
    } finally { $zip.Dispose() }
    # Repair exercises registration and payload consistency on an already installed package.
    Remove-Item -LiteralPath $helper -Force
    Set-Item -LiteralPath $registryPaths[2] -Value '{00000000-0000-0000-0000-000000000000}'
    $repairLog = Join-Path $logs 'repair.log'
    Run-Msi @('/famus', $productCode, '/qn', '/norestart', '/L*V', $repairLog) $repairLog
    if (-not (Test-Path -LiteralPath $helper)) { throw 'Repair did not restore the missing application.' }
    $repairedHandler = (Get-Item -LiteralPath $registryPaths[2]).GetValue('')
    if ($repairedHandler -ne $clsid) { throw ('Repair did not restore the context-menu registration. Actual value: ' + $repairedHandler) }
} finally {
    if ($installed) {
        $uninstallLog = Join-Path $logs 'uninstall.log'
        Run-Msi @('/x', $productCode, '/qn', '/norestart', '/L*V', $uninstallLog) $uninstallLog
    }
}
foreach ($key in $registryPaths) { if (Test-Path -LiteralPath $key) { throw "Uninstall left product registry data: $key" } }
if (Test-Path -LiteralPath $installDirectory) { throw "Uninstall left the application directory: $installDirectory" }
if (-not (Test-Path -LiteralPath $output)) { throw 'Uninstall removed a user-generated workbook.' }
Write-Host "Installer smoke passed: install, native integration, helper, repair, uninstall, workbook retention. Logs: $logs"
