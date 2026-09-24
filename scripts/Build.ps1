[CmdletBinding()]
param(
    [ValidatePattern('^\d+\.\d+\.\d+$')][string]$Version = '1.1.0',
    [string]$ProjectUrl = 'https://github.com/prozac0401/File-List-To-Excel',
    [string]$NativeToolchain,
    [switch]$TestInstaller
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$repo = Split-Path -Parent $PSScriptRoot
$artifacts = Join-Path $repo 'artifacts'
$publish = Join-Path $artifacts 'publish'
$nativeBuild = Join-Path $artifacts 'shell'
$release = Join-Path $artifacts 'release'
$testResults = Join-Path $artifacts 'test-results'
if ($env:OS -ne 'Windows_NT') { throw 'Build.ps1 requires 64-bit Windows.' }
if (-not [Environment]::Is64BitProcess) { throw 'Run this script in 64-bit PowerShell.' }
$dotnet = Join-Path $repo '.tools/dotnet/dotnet.exe'
if (-not (Test-Path -LiteralPath $dotnet)) { $dotnet = (Get-Command dotnet -ErrorAction Stop).Source }
# WiX 4 uses .NET 6 tooling; roll it forward onto the pinned supported SDK runtime.
$env:DOTNET_ROLL_FORWARD = 'Major'
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
$env:DOTNET_NOLOGO = '1'
function Invoke-Native([string]$Command, [string[]]$Arguments) {
    & $Command @Arguments
    if ($LASTEXITCODE -ne 0) { throw "$Command failed with exit code $LASTEXITCODE" }
}
function Reset-ArtifactDirectory([string]$Path) {
    $resolved = [IO.Path]::GetFullPath($Path)
    $boundary = [IO.Path]::GetFullPath($artifacts).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($boundary, [StringComparison]::OrdinalIgnoreCase)) { throw "Unsafe artifact path: $resolved" }
    if (Test-Path -LiteralPath $resolved) { Remove-Item -LiteralPath $resolved -Recurse -Force }
    New-Item -ItemType Directory -Path $resolved -Force | Out-Null
}
function Xml-Escape([string]$Value) { [Security.SecurityElement]::Escape($Value) }
function Payload-Id([string]$Value) {
    $hash = [Security.Cryptography.SHA256]::Create()
    try { ([BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($Value.ToLowerInvariant())))).Replace('-', '').Substring(0, 24) }
    finally { $hash.Dispose() }
}
function Payload-Guid([string]$Value) {
    $hash = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $hash.ComputeHash([Text.Encoding]::UTF8.GetBytes('FileListToExcel.Component.v1/' + $Value.ToLowerInvariant()))
        $hex = ([BitConverter]::ToString($bytes)).Replace('-', '').Substring(0, 32)
        ([guid]::ParseExact($hex, 'N')).ToString('B').ToUpperInvariant()
    } finally { $hash.Dispose() }
}
Push-Location $repo
try {
    foreach ($directory in @($publish, $release, $testResults)) { Reset-ArtifactDirectory $directory }
    Invoke-Native $dotnet @('tool', 'restore')
    Invoke-Native $dotnet @('tool', 'run', 'wix', '--', 'extension', 'add', 'WixToolset.Util.wixext/4.0.6')
    Invoke-Native $dotnet @('test', 'FileListToExcel.sln', '-c', 'Release', '--logger', 'trx', '--results-directory', $testResults, ('-p:Version=' + $Version))
    $trxFiles = @(Get-ChildItem -LiteralPath $testResults -Filter '*.trx' -File)
    if ($trxFiles.Count -lt 2) { throw 'Expected test results from both the core and app test projects.' }
    foreach ($trxFile in $trxFiles) {
        [xml]$trx = [IO.File]::ReadAllText($trxFile.FullName)
        if ([int]$trx.TestRun.ResultSummary.Counters.executed -eq 0) { throw ('No tests ran in ' + $trxFile.Name) }
    }
    Invoke-Native $dotnet @('publish', 'src/FileListToExcel.App/FileListToExcel.App.csproj', '-c', 'Release', '-r', 'win-x64', '--self-contained', 'true', '-o', $publish, ('-p:Version=' + $Version), '-p:DebugType=None', '-p:DebugSymbols=false')
    $cmakeCommand = Get-Command cmake -ErrorAction SilentlyContinue
    $cmakePortable = Get-ChildItem -LiteralPath (Join-Path $repo '.tools') -Directory -Filter 'cmake*' -ErrorAction SilentlyContinue | ForEach-Object { Join-Path $_.FullName 'bin/cmake.exe' } | Where-Object { Test-Path -LiteralPath $_ } | Sort-Object -Descending | Select-Object -First 1
    if ($cmakeCommand) { $cmake = $cmakeCommand.Source }
    elseif ($cmakePortable) { $cmake = $cmakePortable }
    else {
        $vsBase = Join-Path $env:ProgramFiles 'Microsoft Visual Studio/2022'
        $cmake = Get-ChildItem -LiteralPath $vsBase -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            Join-Path $_.FullName 'Common7/IDE/CommonExtensions/Microsoft/CMake/CMake/bin/cmake.exe'
        } | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
        if (-not $cmake) { throw 'Install CMake and Visual Studio 2022 C++ Build Tools (including a Windows SDK).' }
    }
    $ctest = Join-Path (Split-Path -Parent $cmake) 'ctest.exe'
    $nativeArguments = @('-S', 'src/FileListToExcel.Shell', '-B', $nativeBuild, '-DCMAKE_BUILD_TYPE=Release', ('-DCMAKE_RUNTIME_OUTPUT_DIRECTORY_RELEASE=' + (Join-Path $nativeBuild 'Release')))
    if (-not $NativeToolchain) {
        $portable = Get-ChildItem -LiteralPath (Join-Path $repo '.tools') -Directory -Filter 'llvm-mingw*' -ErrorAction SilentlyContinue | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'bin/x86_64-w64-mingw32-clang++.exe') } | Select-Object -ExpandProperty FullName -First 1
        if ($portable) { $NativeToolchain = $portable }
    }
    if ($NativeToolchain) {
        $toolBin = Join-Path ([IO.Path]::GetFullPath($NativeToolchain)) 'bin'
        $env:PATH = $toolBin + ';' + $env:PATH
        $nativeArguments += @('-G', 'Ninja', ('-DCMAKE_RC_COMPILER=' + (Join-Path $toolBin 'x86_64-w64-mingw32-windres.exe')), ('-DCMAKE_CXX_COMPILER=' + (Join-Path $toolBin 'x86_64-w64-mingw32-clang++.exe')))
        $ninjaCommand = Get-Command ninja -ErrorAction SilentlyContinue
        if (-not $ninjaCommand) {
            $ninja = Join-Path $repo '.tools/ninja-fast/ninja.exe'
            if (-not (Test-Path -LiteralPath $ninja)) { $ninja = Join-Path $repo '.tools/ninja/ninja.exe' }
            if (-not (Test-Path -LiteralPath $ninja)) { $ninja = Join-Path (Split-Path -Parent (Split-Path -Parent $cmake)) 'Ninja/ninja.exe' }
            if (Test-Path -LiteralPath $ninja) { $nativeArguments += '-DCMAKE_MAKE_PROGRAM=' + $ninja }
        }
    } else { $nativeArguments += @('-G', 'Visual Studio 17 2022', '-A', 'x64') }
    Invoke-Native $cmake $nativeArguments
    Invoke-Native $cmake @('--build', $nativeBuild, '--config', 'Release', '--parallel')
    Invoke-Native $ctest @('--test-dir', $nativeBuild, '-C', 'Release', '--output-on-failure', '--output-junit', (Join-Path $testResults 'native.xml'))
    $shellDll = Join-Path $nativeBuild 'Release/FileListToExcel.Shell.dll'
    if (-not (Test-Path -LiteralPath $shellDll)) { $shellDll = Join-Path $nativeBuild 'FileListToExcel.Shell.dll' }
    if (-not (Test-Path -LiteralPath $shellDll)) { throw 'Native shell DLL was not built.' }
    Copy-Item -LiteralPath $shellDll -Destination $publish
    foreach ($document in @('README.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md')) {
        $source = Join-Path $repo $document
        if (Test-Path -LiteralPath $source) { Copy-Item -LiteralPath $source -Destination $publish }
    }
    # Self-contained publish omits the runtime packages' legal documents; ship exact-version originals.
    $runtimeConfig = Get-Content -LiteralPath (Join-Path $publish 'FileListToExcel.runtimeconfig.json') -Raw | ConvertFrom-Json
    $assets = Get-Content -LiteralPath (Join-Path $repo 'src/FileListToExcel.App/obj/project.assets.json') -Raw | ConvertFrom-Json
    foreach ($framework in $runtimeConfig.runtimeOptions.includedFrameworks) {
        $relativePackage = $framework.name.ToLowerInvariant() + '.runtime.win-x64/' + $framework.version
        $runtimePackage = $assets.packageFolders.PSObject.Properties.Name | ForEach-Object { Join-Path $_ $relativePackage } | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
        if (-not $runtimePackage) { throw ('Cannot locate runtime license source for ' + $relativePackage) }
        $licenseFiles = @(Get-ChildItem -LiteralPath $runtimePackage -File | Where-Object { $_.Name -like 'LICENSE*' -or $_.Name -like 'THIRD-PARTY-NOTICES*' })
        if ($licenseFiles.Count -eq 0) { throw ('Runtime package has no license files: ' + $runtimePackage) }
        $licenseDirectory = Join-Path $publish ('licenses/' + $framework.name)
        New-Item -ItemType Directory -Path $licenseDirectory -Force | Out-Null
        $licenseFiles | Copy-Item -Destination $licenseDirectory
    }
    if ($NativeToolchain) {
        $nativeLicense = Join-Path $NativeToolchain 'LICENSE.TXT'
        if (-not (Test-Path -LiteralPath $nativeLicense)) { throw 'Portable native toolchain license is missing.' }
        $nativeLicenseDirectory = Join-Path $publish 'licenses/llvm-mingw'
        New-Item -ItemType Directory -Path $nativeLicenseDirectory -Force | Out-Null
        Copy-Item -LiteralPath $nativeLicense -Destination $nativeLicenseDirectory
    }
    $nativeNotices = Join-Path $repo 'docs/licenses/native'
    if (Test-Path -LiteralPath $nativeNotices) {
        $nativeNoticeDestination = Join-Path $publish 'licenses/native'
        New-Item -ItemType Directory -Path $nativeNoticeDestination -Force | Out-Null
        Get-ChildItem -LiteralPath $nativeNotices | Copy-Item -Destination $nativeNoticeDestination -Recurse -Force
    }
    $sqliteNotices = Join-Path $repo 'docs/licenses/sqlite'
    $sqliteNoticeDestination = Join-Path $publish 'licenses/sqlite'
    if (-not (Test-Path -LiteralPath $sqliteNotices -PathType Container)) { throw 'SQLite dependency license documents are missing.' }
    New-Item -ItemType Directory -Path $sqliteNoticeDestination -Force | Out-Null
    Get-ChildItem -LiteralPath $sqliteNotices | Copy-Item -Destination $sqliteNoticeDestination -Recurse -Force
    $wixLicenseDirectory = Join-Path $publish 'licenses/WiX'
    New-Item -ItemType Directory -Path $wixLicenseDirectory -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $repo 'installer/LICENSE-WIX.txt') -Destination $wixLicenseDirectory
    # Explicit HKCU keypaths keep per-user components repairable without ICE38/ICE64 suppressions.
    $payloadXml = New-Object Text.StringBuilder
    [void]$payloadXml.AppendLine('<?xml version="1.0" encoding="utf-8"?><Wix xmlns="http://wixtoolset.org/schemas/v4/wxs"><Fragment>')
    [void]$payloadXml.AppendLine('<DirectoryRef Id="INSTALLFOLDER">')
    $directories = @(Get-ChildItem -LiteralPath $publish -Directory -Recurse | Sort-Object FullName)
    foreach ($directory in $directories) {
        $relative = $directory.FullName.Substring($publish.Length + 1)
        $parent = Split-Path -Parent $relative
        $parentId = if ($parent) { 'D' + (Payload-Id $parent) } else { 'INSTALLFOLDER' }
        $id = 'D' + (Payload-Id $relative)
        [void]$payloadXml.AppendLine('</DirectoryRef><DirectoryRef Id="' + $parentId + '"><Directory Id="' + $id + '" Name="' + (Xml-Escape $directory.Name) + '" /></DirectoryRef><DirectoryRef Id="INSTALLFOLDER">')
    }
    [void]$payloadXml.AppendLine('</DirectoryRef><ComponentGroup Id="ApplicationPayload">')
    foreach ($file in @(Get-ChildItem -LiteralPath $publish -File -Recurse | Sort-Object FullName)) {
        $relative = $file.FullName.Substring($publish.Length + 1)
        $id = Payload-Id $relative
        $parent = Split-Path -Parent $relative
        $directoryId = if ($parent) { 'D' + (Payload-Id $parent) } else { 'INSTALLFOLDER' }
        [void]$payloadXml.AppendLine('<Component Id="C' + $id + '" Directory="' + $directoryId + '" Guid="' + (Payload-Guid ('file/' + $relative)) + '" Bitness="always64">')
        [void]$payloadXml.AppendLine('<File Id="F' + $id + '" Source="' + (Xml-Escape $file.FullName) + '" />')
        [void]$payloadXml.AppendLine('<RegistryValue Root="HKCU" Key="Software\FileListToExcel\Components" Name="' + $id + '" Type="integer" Value="1" KeyPath="yes" />')
        [void]$payloadXml.AppendLine('</Component>')
    }
    foreach ($directory in $directories) {
        $id = Payload-Id ($directory.FullName.Substring($publish.Length + 1))
        [void]$payloadXml.AppendLine('<Component Id="Remove' + $id + '" Directory="D' + $id + '" Guid="*" Bitness="always64"><RegistryValue Root="HKCU" Key="Software\FileListToExcel\Directories" Name="' + $id + '" Type="integer" Value="1" KeyPath="yes" /><RemoveFolder Id="R' + $id + '" On="uninstall" /></Component>')
    }
    [void]$payloadXml.AppendLine('</ComponentGroup></Fragment></Wix>')
    $payloadFile = Join-Path $artifacts 'Payload.wxs'
    [IO.File]::WriteAllText($payloadFile, $payloadXml.ToString(), (New-Object Text.UTF8Encoding($false)))
    $msi = Join-Path $release "FileListToExcel-$Version-win-x64.msi"
    $wixArguments = @('tool', 'run', 'wix', '--', 'build', 'installer/Product.wxs', $payloadFile, '-ext', 'WixToolset.Util.wixext/4.0.6', '-arch', 'x64', '-d', "ProductVersion=$Version", '-d', "ProjectUrl=$ProjectUrl", '-d', ('AppIcon=' + (Join-Path $repo 'src/FileListToExcel.App/app.ico')), '-o', $msi, '-wx')
    Invoke-Native $dotnet $wixArguments
    # Symbols help development but are not public installer assets.
    Get-ChildItem -LiteralPath $release -Filter '*.wixpdb' | Remove-Item -Force
    if ($TestInstaller) {
        & (Join-Path $PSScriptRoot 'Test-Installer.ps1') -MsiPath $msi
        if ([version]$Version -le [version]'0.9.0') { throw 'Upgrade validation requires a target version newer than 0.9.0.' }
        $upgradeDirectory = Join-Path $artifacts 'installer-upgrade'
        New-Item -ItemType Directory -Path $upgradeDirectory -Force | Out-Null
        $previousMsi = Join-Path $upgradeDirectory 'FileListToExcel-0.9.0-win-x64.msi'
        $previousArguments = [string[]]$wixArguments.Clone()
        for ($index = 0; $index -lt $previousArguments.Length; $index++) {
            if ($previousArguments[$index] -eq "ProductVersion=$Version") { $previousArguments[$index] = 'ProductVersion=0.9.0' }
            if ($previousArguments[$index] -eq $msi) { $previousArguments[$index] = $previousMsi }
        }
        Invoke-Native $dotnet $previousArguments
        & (Join-Path $PSScriptRoot 'Test-Upgrade.ps1') -PreviousMsiPath $previousMsi -MsiPath $msi
    }
    $hash = (Get-FileHash -LiteralPath $msi -Algorithm SHA256).Hash.ToLowerInvariant()
    [IO.File]::WriteAllText((Join-Path $release 'SHA256SUMS.txt'), "$hash  $([IO.Path]::GetFileName($msi))" + [Environment]::NewLine, (New-Object Text.UTF8Encoding($false)))
    Write-Host "Built $msi"
} finally { Pop-Location }
