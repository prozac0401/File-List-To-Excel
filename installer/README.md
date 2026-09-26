# MSI packaging

The MSI requires native Intel/AMD x64 Windows (ARM64 is rejected) and installs for the current user under %LOCALAPPDATA%\Programs\FileListToExcel. The required component registers the Explorer COM class in HKCU and adds classic Explorer context-menu handlers for files, selected folders, and folder backgrounds. Windows 11 exposes these commands under **Show more options**.

- UpgradeCode: {138C25B0-CA8D-47D7-AE96-8B5F7B2D69A1}. Keep this stable across releases.
- Shell CLSID: {AF7A7218-8B84-4A11-9790-EB24A43EC39E}.
- The application carries its .NET runtime; the native shell uses a static C/C++ runtime. End users do not install developer tools or a separate runtime.
- Files are individual Windows Installer components with stable HKCU registry key paths. The build generates Payload.wxs and validates the MSI without suppressing ICE checks.
- Uninstall removes only installed resources and empty installation directories. Generated workbooks and the reusable SQLite hash cache remain outside the install directory. No service, resident indexer, or scheduled task is installed.
- The same COM registration provides matching-file and duplicate-file commands. The self-contained payload includes Microsoft.Data.Sqlite, SQLitePCLRaw, and the Windows x64 SQLite native library; dependency licenses are shipped under licenses/sqlite.
- Major upgrades remove the previous version inside the installation transaction. Downgrades are blocked. MSI repair restores payload and registration.
- The MSI does not forcibly terminate Explorer. Windows may require signing out if Explorer holds an old shell DLL open during an upgrade or uninstall.

Build from a 64-bit Windows PowerShell session with .NET SDK 10.0.401, Visual Studio 2022 C++ build tools and Windows SDK, and CMake available:

    ./scripts/Build.ps1 -Version 1.2.0 -TestInstaller

The optional NativeToolchain parameter accepts a portable llvm-mingw installation. A matching .tools/llvm-mingw* directory is detected automatically for local development. CI uses MSVC.

WiX 4.0.6 is pinned in .config/dotnet-tools.json. Its tooling runs on the installed supported .NET runtime using major roll-forward; The WiX Util native architecture-check custom action is embedded in the MSI; its original license is packaged under licenses/WiX. Developer SDKs are not shipped to users.

Test-Installer.ps1 deliberately refuses to run over any existing installation or product registration. Use a clean Windows user or disposable VM in that case. It verifies installation, native shell behavior, existing file-list output, duplicate/match/selected-file workbooks, SQLite reuse on a second run, source preservation, repair, removal, no residual helper process, and retention of workbooks and hash cache. Logs and sample output are retained under artifacts/installer-smoke. The full validation build also creates a test-only 0.9.0 MSI in artifacts/installer-upgrade and verifies a transactional major upgrade to the target version, uninstall, and preservation of user output. Only the target MSI is released.

The build outputs an MSI and SHA256SUMS.txt under artifacts/release. The checked-in pipeline tests those exact assets before publishing a version tag whose commit is already on main. Code signing requires the publisher's signing certificate and is not fabricated by the build.

## Optional Excel component (v1.2)

The feature-tree setup page offers **Excel에서 선택한 파일 복사**. It is off by default; existing silent upgrades do not enable a feature that was absent. Choosing it installs native x86 and x64 shims and registers the matching COM views for Excel. Office bitness is independent from the Explorer extension. Save open work and restart Excel to load the new menu; the installer never kills Excel or changes Trust Center, macro, add-in-policy, certificate, or MOTW settings. There are no VBA/VBS steps or manual add-in registration. Excel Desktop is required for this optional feature; the required file-list commands still work without Excel.

Excel ProgID: `FileListToExcel.ExcelAddIn`; CLSID: `{0B1E297C-42CC-48A4-A973-3AA8EAF26795}`. These keys and the unique menu tags are independent from Explorer, roster matching, and Workspace. `LoadBehavior=3` is written for initial setup. Existing known load states, including user/Office disabled states, are normalized from MSI registry-search DWORD encoding and preserved on repair/upgrade. An unrecognized existing value fails closed to disabled. No background process monitors or restores this setting.

The native shims use static C++ runtimes, and the existing helper remains self-contained Windows x64. Per-user installation requires no separate end-user SDK, Python, VSTO, or scripting-policy change. Unsigned binaries still require the organization to allow this add-in; signing and organizational approval are separate release requirements.

`Build.ps1 -TestInstaller` now runs the existing regression once with the default main-only feature and once with the Excel feature explicitly included. `Test-Installer.ps1 -IncludeExcelIntegration` verifies both registry views/payloads, `REG_DWORD` load state, preservation of a test-disabled state during repair, and uninstall cleanup. It does not replace actual Excel startup/restart/menu/copy verification. Product source files, collected copies, user reports and generated workbooks are not installer-owned resources.

For an explicitly requested unattended installation including Excel integration, use MSI property `ADDLOCAL=Complete,ExcelIntegration`. Do not apply it to an existing user installation without their choice. Read-only registration/architecture diagnostics are available in `scripts/Get-ExcelIntegrationStatus.ps1`; a registered DLL alone is not proof of a successful Excel load.

The CAB payload uses MSZIP compression to keep repeated build/validation runs practical. This changes installer size/build time only; owned resources, hashes, component selection and installation behavior are unchanged.
