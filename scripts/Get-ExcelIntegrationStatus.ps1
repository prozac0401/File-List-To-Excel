[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
function Get-PeMachine([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return 'missing' }
    $stream = [IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
    $reader = [IO.BinaryReader]::new($stream)
    try {
        if ($reader.ReadUInt16() -ne 0x5a4d) { return 'invalid PE' }
        $stream.Position = 0x3c
        $offset = $reader.ReadInt32()
        if ($offset -lt 0 -or $offset -gt $stream.Length - 6) { return 'invalid PE' }
        $stream.Position = $offset
        if ($reader.ReadUInt32() -ne 0x4550) { return 'invalid PE' }
        switch ($reader.ReadUInt16()) { 0x8664 { 'x64' } 0x14c { 'x86' } default { 'unsupported architecture' } }
    } finally { $reader.Dispose(); $stream.Dispose() }
}
$registrations = foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
    $userKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey('CurrentUser', $view)
    try {
        $comKey = $userKey.OpenSubKey('Software\Classes\CLSID\{0B1E297C-42CC-48A4-A973-3AA8EAF26795}\InprocServer32')
        try {
            $dll = if ($comKey) { [string]$comKey.GetValue('') } else { $null }
            [pscustomobject]@{ View = [string]$view; Registered = [bool]$comKey; Dll = $dll; BinaryArchitecture = if ($dll) { Get-PeMachine $dll } else { 'not registered' } }
        } finally { if ($comKey) { $comKey.Dispose() } }
    } finally { $userKey.Dispose() }
}
$addinPath = 'HKCU:\Software\Microsoft\Office\Excel\Addins\FileListToExcel.ExcelAddIn'
$addin = Get-Item -LiteralPath $addinPath -ErrorAction SilentlyContinue
$loadBehavior = if ($addin) { $addin.GetValue('LoadBehavior') } else { $null }
$excelPaths = @('HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', 'HKCU:\Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe')
$excel = foreach ($key in $excelPaths) {
    $entry = Get-Item -LiteralPath $key -ErrorAction SilentlyContinue
    if ($entry) {
        $path = [string]$entry.GetValue('')
        [pscustomobject]@{ Path = $path; Architecture = Get-PeMachine $path; Version = if (Test-Path -LiteralPath $path) { [Diagnostics.FileVersionInfo]::GetVersionInfo($path).FileVersion } else { $null } }
    }
}
$policyKeys = foreach ($hive in @('HKCU:', 'HKLM:')) {
    foreach ($suffix in @('Software\Policies\Microsoft\Office\16.0\Excel\Security', 'Software\Policies\Microsoft\Office\16.0\Excel\Resiliency', 'Software\Policies\Microsoft\Office\16.0\Excel\Resiliency\AddinList')) {
        $key = Join-Path $hive $suffix
        if (Test-Path -LiteralPath $key) { [pscustomobject]@{ Key = $key; Values = Get-ItemProperty -LiteralPath $key } }
    }
}
[pscustomobject]@{
    Excel = @($excel)
    OfficeAvailability = if (@($excel).Count) { 'Excel registration found; actual startup must be tested.' } else { 'Excel Desktop not found. Main file-list commands remain available.' }
    Registration = @($registrations)
    LoadBehavior = $loadBehavior
    LoadState = if ($null -eq $loadBehavior) { 'Optional component is absent or registration is incomplete.' } elseif ($loadBehavior -eq 3) { 'Configured for startup. Actual Connected state must be observed in a normal Excel start.' } elseif ($loadBehavior -eq 2 -or $loadBehavior -eq 0) { 'Disabled or previous load failed. Do not force enable; check user choice, policy, payload and Excel diagnostics.' } else { 'Existing non-default load mode preserved. Review Excel diagnostics without changing it.' }
    BinaryDiagnostics = @($registrations | ForEach-Object { if ($_.Registered -and $_.BinaryArchitecture -eq 'missing') { 'Missing add-in payload: repair product installation.' } elseif ($_.Registered -and (($_.View -eq 'Registry64' -and $_.BinaryArchitecture -ne 'x64') -or ($_.View -eq 'Registry32' -and $_.BinaryArchitecture -ne 'x86'))) { 'COM registration and binary bitness mismatch.' } })
    PolicyEvidence = @($policyKeys)
    Runtime = 'Native static C++ add-in; self-contained x64 helper. No separate VSTO/.NET runtime required.'
    Validation = 'Read-only diagnostics do not establish successful automatic load. Check a normal Excel start; do not change policy or force Connect=true.'
} | ConvertTo-Json -Depth 6
