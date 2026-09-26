[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$MsiPath, [switch]$IncludeExcelIntegration,
    [switch]$ContinueAfterShellRegistrationFailure)
$ErrorActionPreference = 'Stop'
trap {
    Write-Host ($_ | Format-List * -Force | Out-String)
    Write-Host $_.ScriptStackTrace
    exit 1
}
Set-StrictMode -Version Latest
$registrationFailures = [System.Collections.Generic.List[string]]::new()
function Assert-ShellHandler([string]$Path, [string]$Expected, [string]$Stage) {
    $key = Get-Item -LiteralPath $Path
    $actual = if ($null -ne $key) { $key.GetValue('') } else { '<missing>' }
    if ($actual -ne $Expected) {
        $message = "$Stage context-menu registration mismatch: $Path; actual=$actual"
        if (-not $ContinueAfterShellRegistrationFailure) { throw $message }
        # Diagnostic continuation still fails the suite at the end. It never
        # converts a registration failure into a successful release gate.
        $registrationFailures.Add($message)
        Write-Warning $message
    }
}
if (-not [Environment]::Is64BitProcess) { throw 'Use 64-bit PowerShell for installer tests.' }
$repo = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'Test-InstalledDuplicates.ps1')
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

$excelProgId = 'FileListToExcel.ExcelAddIn'
$excelClsid = '{0B1E297C-42CC-48A4-A973-3AA8EAF26795}'
$excelRegistration = 'HKCU:\Software\Microsoft\Office\Excel\Addins\' + $excelProgId
if (Test-Path -LiteralPath $excelRegistration) { throw 'Existing Excel add-in registration found. Use a clean Windows account or VM.' }
function Assert-ExcelRegistration([bool]$Expected, [int]$ExpectedLoadBehavior = 3) {
    if ((Test-Path -LiteralPath $excelRegistration) -ne $Expected) { throw 'Unexpected optional Excel feature registration state.' }
    if ($Expected) {
        $key = Get-Item -LiteralPath $excelRegistration
        if ($key.GetValueKind('LoadBehavior') -ne 'DWord' -or $key.GetValue('LoadBehavior') -ne $ExpectedLoadBehavior) { throw 'Excel LoadBehavior type/value was not preserved correctly.' }
    }
    foreach ($architecture in @('x64', 'x86')) {
        $view = if ($architecture -eq 'x64') { [Microsoft.Win32.RegistryView]::Registry64 } else { [Microsoft.Win32.RegistryView]::Registry32 }
        $user = [Microsoft.Win32.RegistryKey]::OpenBaseKey('CurrentUser', $view)
        try {
            $key = $user.OpenSubKey('Software\Classes\CLSID\' + $excelClsid + '\InprocServer32')
            try {
                if ([bool]$key -ne $Expected) { throw ('Unexpected Excel COM registry state for ' + $architecture) }
                $dll = Join-Path $installDirectory ('FileListToExcel.ExcelAddIn.' + $architecture + '.dll')
                if ((Test-Path -LiteralPath $dll) -ne $Expected) { throw ('Unexpected optional payload state for ' + $architecture) }
                if ($Expected -and ($key.GetValue('') -ne $dll -or $key.GetValue('ThreadingModel') -ne 'Apartment')) { throw ('Incorrect Excel COM registration for ' + $architecture) }
            } finally { if ($key) { $key.Dispose() } }
        } finally { $user.Dispose() }
    }
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
$duplicateRetained = @()
function Run-Msi([string[]]$Arguments, [string]$Log) {
    $quoted = @($Arguments | ForEach-Object { if ($_ -match '[\s"]') { '"' + $_.Replace('"', '\"') + '"' } else { $_ } })
    $process = Start-Process -FilePath (Join-Path $env:WINDIR 'System32/msiexec.exe') -ArgumentList ($quoted -join ' ') -PassThru -Wait -WindowStyle Hidden
    if ($process.ExitCode -notin @(0, 3010)) { throw "Windows Installer failed with exit code $($process.ExitCode). See $Log" }
}
try {
    $installLog = Join-Path $logs 'install.log'
    $installArguments = @('/i', $MsiPath, '/qn', '/norestart', '/L*V', $installLog)
    if ($IncludeExcelIntegration) { $installArguments += 'ADDLOCAL=Complete,ExcelIntegration' }
    Run-Msi $installArguments $installLog
    $installed = $true
    Assert-ExcelRegistration ([bool]$IncludeExcelIntegration)
    $helper = Join-Path $installDirectory 'FileListToExcel.exe'
    $shell = Join-Path $installDirectory 'FileListToExcel.Shell.dll'
    foreach ($file in @($helper, $shell)) { if (-not (Test-Path -LiteralPath $file)) { throw "Missing installed payload: $file" } }
    if ((Get-Item -LiteralPath "HKCU:\Software\Classes\CLSID\$clsid\InprocServer32").GetValue('') -ne $shell) { throw 'COM registration has the wrong DLL path.' }
    if ((Get-Item -LiteralPath "HKCU:\Software\Classes\CLSID\$clsid\InprocServer32").GetValue('ThreadingModel') -ne 'Apartment') { throw 'Invalid COM threading model.' }
    foreach ($key in $registryPaths[2..4]) { Assert-ShellHandler $key $clsid 'Installed' }
    $shellTest = Join-Path $repo 'artifacts/shell/Release/ShellTests.exe'
    if (-not (Test-Path -LiteralPath $shellTest)) { $shellTest = Join-Path $repo 'artifacts/shell/ShellTests.exe' }
    if (-not (Test-Path -LiteralPath $shellTest)) { throw 'Build native ShellTests before running installer tests.' }
    $nativeLog = Join-Path $logs 'native-shell.txt'
    $nativeErrorLog = Join-Path $logs 'native-shell-error.txt'
    $nativeProcess = Start-Process -FilePath $shellTest -ArgumentList ('"' + $shell + '"') -PassThru -Wait -WindowStyle Hidden -RedirectStandardOutput $nativeLog -RedirectStandardError $nativeErrorLog
    try {
        Write-Host (Get-Content -LiteralPath $nativeLog -Raw)
        if ($nativeProcess.ExitCode -ne 0) { throw "The installed native shell integration failed with exit code $($nativeProcess.ExitCode). See $nativeErrorLog" }
    } finally { $nativeProcess.Dispose() }
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
    $duplicateRetained = Test-InstalledDuplicateWorkbooks $helper $logs
    # Repair exercises registration and payload consistency on an already installed package.
    Remove-Item -LiteralPath $helper -Force
    # A missing key is already a repair defect; do not create it manually and
    # hide that state. Corrupt it only when the exact owned key is present.
    $handlerToCorrupt = Get-Item -LiteralPath $registryPaths[2]
    if ($null -ne $handlerToCorrupt) { Set-Item -LiteralPath $registryPaths[2] -Value '{00000000-0000-0000-0000-000000000000}' }
    if ($IncludeExcelIntegration) { Set-ItemProperty -LiteralPath $excelRegistration -Name LoadBehavior -Value 2 -Type DWord }
    $repairLog = Join-Path $logs 'repair.log'
    Run-Msi @('/famus', $productCode, '/qn', '/norestart', '/L*V', $repairLog) $repairLog
    if (-not (Test-Path -LiteralPath $helper)) { throw 'Repair did not restore the missing application.' }
    Assert-ShellHandler $registryPaths[2] $clsid 'Repaired'
    Assert-ExcelRegistration ([bool]$IncludeExcelIntegration) 2
    Assert-DuplicateTestRetention $duplicateRetained
    Assert-NoInstalledHelper $helper
} catch {
    Write-Host ("Installer test failed before cleanup: " + $_.Exception.Message)
    Write-Host $_.ScriptStackTrace
    throw
} finally {
    if ($installed) {
        $uninstallLog = Join-Path $logs 'uninstall.log'
        Run-Msi @('/x', $productCode, '/qn', '/norestart', '/L*V', $uninstallLog) $uninstallLog
    }
}
Assert-ExcelRegistration $false
foreach ($key in $registryPaths) { if (Test-Path -LiteralPath $key) { throw "Uninstall left product registry data: $key" } }
if (Test-Path -LiteralPath $installDirectory) { throw "Uninstall left the application directory: $installDirectory" }
if (-not (Test-Path -LiteralPath $output)) { throw 'Uninstall removed a user-generated workbook.' }
Assert-DuplicateTestRetention $duplicateRetained
Assert-NoInstalledHelper (Join-Path $installDirectory 'FileListToExcel.exe')
if ($registrationFailures.Count -ne 0) {
    throw ("Independent install/function/repair/uninstall stages completed, but shell registration failed: " + ($registrationFailures -join '; ') + ". Logs: $logs")
}
Write-Host "Installer smoke passed: install, native integration, existing list, duplicate/matches/selection, SQLite reuse, repair, uninstall, source/workbook/cache retention, no helper process. Logs: $logs"
