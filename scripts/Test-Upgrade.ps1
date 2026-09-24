[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$PreviousMsiPath,
    [Parameter(Mandatory = $true)][string]$MsiPath
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if (-not [Environment]::Is64BitProcess) { throw 'Use 64-bit PowerShell for upgrade tests.' }

$repo = Split-Path -Parent $PSScriptRoot
$PreviousMsiPath = (Resolve-Path -LiteralPath $PreviousMsiPath).Path
$MsiPath = (Resolve-Path -LiteralPath $MsiPath).Path
$installDirectory = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'Programs/FileListToExcel'))
$expectedUpgradeCode = '{138C25B0-CA8D-47D7-AE96-8B5F7B2D69A1}'
$clsid = '{AF7A7218-8B84-4A11-9790-EB24A43EC39E}'
$registryPaths = @(
    'HKCU:\Software\FileListToExcel',
    "HKCU:\Software\Classes\CLSID\$clsid",
    'HKCU:\Software\Classes\*\shellex\ContextMenuHandlers\FileListToExcel',
    'HKCU:\Software\Classes\Directory\shellex\ContextMenuHandlers\FileListToExcel',
    'HKCU:\Software\Classes\Directory\Background\shellex\ContextMenuHandlers\FileListToExcel'
)

function Release-ComObject($Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Get-MsiMetadata($Installer, [string]$Path) {
    $database = $null
    try {
        $database = $Installer.OpenDatabase($Path, 0)
        $values = @{}
        foreach ($property in @('ProductCode', 'ProductVersion', 'ProductName', 'UpgradeCode')) {
            $view = $null
            $record = $null
            try {
                # Property names are fixed above; no user text is interpolated into SQL.
                $view = $database.OpenView("SELECT Value FROM Property WHERE Property = '$property'")
                [void]$view.Execute()
                $record = $view.Fetch()
                if ($null -eq $record) { throw "MSI is missing $($property): $Path" }
                $values[$property] = $record.StringData(1)
            } finally {
                if ($null -ne $view) { [void]$view.Close() }
                Release-ComObject $record
                Release-ComObject $view
            }
        }
        return [pscustomobject]$values
    } finally { Release-ComObject $database }
}

function Get-RelatedProductCodes($Installer, [string]$UpgradeCode) {
    $related = $null
    try {
        $related = $Installer.RelatedProducts($UpgradeCode)
        foreach ($code in $related) { [string]$code }
    } finally { Release-ComObject $related }
}

function Quote-ProcessArgument([string]$Value) {
    # Test arguments are paths, MSI product codes, and fixed switches. Windows
    # filenames cannot contain a double quote; reject it rather than interpreting it.
    if ($Value.Contains('"')) { throw 'Unexpected double quote in process argument.' }
    if ($Value -match '\s') { return '"' + $Value + '"' }
    return $Value
}

function Run-Msi([string[]]$Arguments, [string]$Log) {
    $quoted = @($Arguments | ForEach-Object { Quote-ProcessArgument $_ })
    $process = Start-Process -FilePath (Join-Path $env:WINDIR 'System32/msiexec.exe') -ArgumentList ($quoted -join ' ') -PassThru -Wait -WindowStyle Hidden
    try {
        if ($process.ExitCode -notin @(0, 3010)) {
            throw "Windows Installer failed with exit code $($process.ExitCode). See $Log"
        }
    } finally { $process.Dispose() }
}

function Assert-Registration {
    $helper = Join-Path $installDirectory 'FileListToExcel.exe'
    $shell = Join-Path $installDirectory 'FileListToExcel.Shell.dll'
    foreach ($file in @($helper, $shell)) {
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Missing installed payload: $file" }
    }
    $registeredDirectory = (Get-Item -LiteralPath $registryPaths[0]).GetValue('InstallDirectory')
    if ([IO.Path]::GetFullPath($registeredDirectory).TrimEnd('\', '/') -ne $installDirectory.TrimEnd('\', '/')) {
        throw 'The registered application directory does not match the expected per-user installation.'
    }
    $inproc = Get-Item -LiteralPath "HKCU:\Software\Classes\CLSID\$clsid\InprocServer32"
    if ($inproc.GetValue('') -ne $shell -or $inproc.GetValue('ThreadingModel') -ne 'Apartment') {
        throw 'The upgraded COM registration is invalid.'
    }
    foreach ($key in $registryPaths[2..4]) {
        if ((Get-Item -LiteralPath $key).GetValue('') -ne $clsid) { throw "Wrong handler CLSID: $key" }
    }
}

function New-TestWorkbook([string]$OutputPath) {
    $helper = Join-Path $installDirectory 'FileListToExcel.exe'
    $arguments = @('--folder', $fixture, '--no-open', '--output', $OutputPath)
    $quoted = @($arguments | ForEach-Object { Quote-ProcessArgument $_ })
    $process = Start-Process -FilePath $helper -ArgumentList ($quoted -join ' ') -PassThru -WindowStyle Hidden
    try {
        if (-not $process.WaitForExit(60000)) {
            # This is only the helper started by this test, never Explorer or Excel.
            $process.Kill()
            $process.WaitForExit()
            throw 'Installed helper did not exit in 60 seconds.'
        }
        $process.Refresh()
        if ($process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
            throw 'Installed helper failed to create an Excel workbook.'
        }
    } finally { $process.Dispose() }

    $zip = [IO.Compression.ZipFile]::OpenRead($OutputPath)
    try {
        foreach ($part in @('[Content_Types].xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/tables/table1.xml')) {
            if ($null -eq $zip.GetEntry($part)) { throw "Workbook is missing $part" }
        }
        $reader = New-Object IO.StreamReader($zip.GetEntry('xl/worksheets/sheet1.xml').Open())
        try { [xml]$sheet = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $rows = $sheet.SelectNodes('//*[local-name()="sheetData"]/*[local-name()="row"]')
        if ($rows.Count -ne 3) { throw "Expected two file rows and one header; found $($rows.Count) rows." }
    } finally { $zip.Dispose() }
}

# Never upgrade or remove a user's existing installation during regression tests.
if (Test-Path -LiteralPath $installDirectory) {
    throw "An existing installation is present at $installDirectory. Use a clean Windows account or VM."
}
foreach ($key in $registryPaths) {
    if (Test-Path -LiteralPath $key) { throw "Existing product registration found at $key. Use a clean Windows account or VM." }
}
if ($PreviousMsiPath -eq $MsiPath) { throw 'Previous and current MSI paths must be different.' }

$windowsInstaller = New-Object -ComObject WindowsInstaller.Installer
$attemptedInstall = $false
$testFailure = $null
$cleanupFailures = New-Object 'System.Collections.Generic.List[string]'
$previous = $null
$current = $null
$logs = $null
try {
    $previous = Get-MsiMetadata $windowsInstaller $PreviousMsiPath
    $current = Get-MsiMetadata $windowsInstaller $MsiPath
    foreach ($package in @($previous, $current)) {
        if ($package.UpgradeCode -ne $expectedUpgradeCode -or $package.ProductName -ne 'File List to Excel') {
            throw 'Upgrade tests accept only File List to Excel MSI packages with the expected UpgradeCode.'
        }
    }
    if ($previous.ProductCode -eq $current.ProductCode) { throw 'A major upgrade must have a new ProductCode.' }
    if ([version]$previous.ProductVersion -ge [version]$current.ProductVersion) {
        throw "Previous version $($previous.ProductVersion) must be older than current version $($current.ProductVersion)."
    }
    $existing = @(Get-RelatedProductCodes $windowsInstaller $expectedUpgradeCode)
    if ($existing.Count -ne 0) { throw 'A related MSI product is already registered. Use a clean Windows account or VM.' }

    $logs = [IO.Path]::GetFullPath((Join-Path $repo ('artifacts/upgrade-smoke/' + [Guid]::NewGuid().ToString('N'))))
    if ($logs.StartsWith($installDirectory.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'The upgrade test artifacts must be outside the application installation directory.'
    }
    $fixture = Join-Path $logs 'input'
    New-Item -ItemType Directory -Path $fixture -Force | Out-Null
    [IO.File]::WriteAllText((Join-Path $fixture 'one.txt'), 'one')
    [IO.File]::WriteAllText((Join-Path $fixture 'two with spaces.txt'), 'two')
    $userWorkbook = Join-Path $logs 'user-result-before-upgrade.xlsx'
    $upgradedWorkbook = Join-Path $logs 'upgraded-helper-result.xlsx'
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    # Set before launching so a partially completed operation is still cleaned up.
    $attemptedInstall = $true
    $previousLog = Join-Path $logs 'install-previous.log'
    Run-Msi -Arguments @('/i', $PreviousMsiPath, '/qn', '/norestart', '/L*V', $previousLog) -Log $previousLog
    Assert-Registration
    $registered = @(Get-RelatedProductCodes $windowsInstaller $expectedUpgradeCode)
    if ($registered.Count -ne 1 -or $registered[0] -ne $previous.ProductCode) {
        throw 'The previous package is not the sole registered related product.'
    }
    if ($windowsInstaller.ProductInfo($previous.ProductCode, 'VersionString') -ne $previous.ProductVersion) {
        throw 'The installed previous ProductVersion is incorrect.'
    }
    New-TestWorkbook $userWorkbook
    $userWorkbookHash = (Get-FileHash -LiteralPath $userWorkbook -Algorithm SHA256).Hash

    $upgradeLog = Join-Path $logs 'upgrade-current.log'
    Run-Msi -Arguments @('/i', $MsiPath, '/qn', '/norestart', '/L*V', $upgradeLog) -Log $upgradeLog
    Assert-Registration
    $registered = @(Get-RelatedProductCodes $windowsInstaller $expectedUpgradeCode)
    if ($registered.Count -ne 1 -or $registered[0] -ne $current.ProductCode) {
        throw 'Major upgrade did not replace the old ProductCode with only the current ProductCode.'
    }
    if ($windowsInstaller.ProductInfo($current.ProductCode, 'VersionString') -ne $current.ProductVersion) {
        throw 'The upgraded ProductVersion is incorrect.'
    }
    $oldUninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\' + $previous.ProductCode
    $newUninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\' + $current.ProductCode
    if (Test-Path -LiteralPath $oldUninstallKey) { throw 'Major upgrade left the old Apps and Features entry.' }
    if (-not (Test-Path -LiteralPath $newUninstallKey)) { throw 'The current Apps and Features entry is missing.' }
    if ((Get-Item -LiteralPath $newUninstallKey).GetValue('DisplayVersion') -ne $current.ProductVersion) {
        throw 'Apps and Features reports the wrong upgraded version.'
    }
    if ((Get-FileHash -LiteralPath $userWorkbook -Algorithm SHA256).Hash -ne $userWorkbookHash) {
        throw 'Upgrade modified a user-generated workbook.'
    }
    New-TestWorkbook $upgradedWorkbook
    $upgradedWorkbookHash = (Get-FileHash -LiteralPath $upgradedWorkbook -Algorithm SHA256).Hash
} catch {
    $testFailure = $_
} finally {
    if ($attemptedInstall) {
        # Remove only this test's two known package IDs. Never delete application
        # folders manually or uninstall any other related product discovered here.
        foreach ($package in @($current, $previous)) {
            try {
                $registered = @(Get-RelatedProductCodes $windowsInstaller $expectedUpgradeCode)
                if ($registered -contains $package.ProductCode) {
                    $uninstallLog = Join-Path $logs ('uninstall-' + $package.ProductVersion + '.log')
                    Run-Msi -Arguments @('/x', $package.ProductCode, '/qn', '/norestart', '/L*V', $uninstallLog) -Log $uninstallLog
                }
            } catch { $cleanupFailures.Add($_.Exception.Message) }
        }
        try {
            if (@(Get-RelatedProductCodes $windowsInstaller $expectedUpgradeCode).Count -ne 0) {
                throw 'Uninstall left a related product registered.'
            }
        } catch { $cleanupFailures.Add($_.Exception.Message) }
    }
    Release-ComObject $windowsInstaller
}
if ($cleanupFailures.Count -ne 0) {
    $details = $cleanupFailures -join '; '
    if ($null -ne $testFailure) { throw "Upgrade test failed: $($testFailure.Exception.Message). Cleanup also failed: $details. Logs: $logs" }
    throw "Upgrade cleanup failed: $details. Logs: $logs"
}
if ($null -ne $testFailure) { throw $testFailure }

foreach ($key in $registryPaths) {
    if (Test-Path -LiteralPath $key) { throw "Uninstall left product registry data: $key" }
}
foreach ($package in @($previous, $current)) {
    $uninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\' + $package.ProductCode
    if (Test-Path -LiteralPath $uninstallKey) { throw "Uninstall left an Apps and Features entry: $uninstallKey" }
}
if (Test-Path -LiteralPath $installDirectory) { throw "Uninstall left the application directory: $installDirectory" }
foreach ($result in @(
    @{ Path = $userWorkbook; Hash = $userWorkbookHash },
    @{ Path = $upgradedWorkbook; Hash = $upgradedWorkbookHash }
)) {
    if (-not (Test-Path -LiteralPath $result.Path -PathType Leaf) -or
        (Get-FileHash -LiteralPath $result.Path -Algorithm SHA256).Hash -ne $result.Hash) {
        throw "Uninstall removed or modified a user-generated workbook: $($result.Path)"
    }
}
Write-Host "Upgrade regression passed: $($previous.ProductVersion) -> $($current.ProductVersion), sole current ProductCode, helper, uninstall, workbook retention. Logs: $logs"
