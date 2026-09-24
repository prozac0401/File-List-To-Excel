# MSI packaging

The MSI requires native Intel/AMD x64 Windows (ARM64 is rejected) and installs for the current user under %LOCALAPPDATA%\Programs\FileListToExcel. It registers one native COM class in HKCU and adds classic Explorer context-menu handlers for files, selected folders, and folder backgrounds. Windows 11 exposes these commands under **Show more options**.

- UpgradeCode: {138C25B0-CA8D-47D7-AE96-8B5F7B2D69A1}. Keep this stable across releases.
- Shell CLSID: {AF7A7218-8B84-4A11-9790-EB24A43EC39E}.
- The application carries its .NET runtime; the native shell uses a static C/C++ runtime. End users do not install developer tools or a separate runtime.
- Files are individual Windows Installer components with stable HKCU registry key paths. The build generates Payload.wxs and validates the MSI without suppressing ICE checks.
- Uninstall removes only installed resources and empty installation directories. Generated workbooks remain outside the install directory.
- Major upgrades remove the previous version inside the installation transaction. Downgrades are blocked. MSI repair restores payload and registration.
- The MSI does not forcibly terminate Explorer. Windows may require signing out if Explorer holds an old shell DLL open during an upgrade or uninstall.

Build from a 64-bit Windows PowerShell session with .NET SDK 10.0.401, Visual Studio 2022 C++ build tools and Windows SDK, and CMake available:

    ./scripts/Build.ps1 -Version 1.0.0 -TestInstaller

The optional NativeToolchain parameter accepts a portable llvm-mingw installation. A matching .tools/llvm-mingw* directory is detected automatically for local development. CI uses MSVC.

WiX 4.0.6 is pinned in .config/dotnet-tools.json. Its tooling runs on the installed supported .NET runtime using major roll-forward; The WiX Util native architecture-check custom action is embedded in the MSI; its original license is packaged under licenses/WiX. Developer SDKs are not shipped to users.

Test-Installer.ps1 deliberately refuses to run over any existing installation or product registration. Use a clean Windows user or disposable VM in that case. It verifies installation, native shell behavior, helper output, repair, removal, and retention of the generated workbook. Logs and sample output are retained under artifacts/installer-smoke. The full validation build also creates a test-only 0.9.0 MSI in artifacts/installer-upgrade and verifies a transactional major upgrade to the target version, uninstall, and preservation of user output. Only the target MSI is released.

The build outputs an MSI and SHA256SUMS.txt under artifacts/release. The checked-in pipeline tests those exact assets before publishing a version tag whose commit is already on main. Code signing requires the publisher's signing certificate and is not fabricated by the build.
