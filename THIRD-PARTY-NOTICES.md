# Third-party notices

The Windows distribution contains the Microsoft .NET runtime and Windows Desktop runtime, distributed under their upstream licenses. The installer includes the exact runtime packages’ license and third-party notice files under the installed licenses directory. Sources and license information: https://github.com/dotnet/runtime and https://github.com/dotnet/winforms.

The application writes Office Open XML using .NET XML and ZIP APIs; no Excel automation library is included. Microsoft Excel is a separately installed product and is not distributed.

The MSI includes the WiX Util 4.0.6 native architecture query custom action (Microsoft Reciprocal License with its generated-output exception). The original WiX license is included in the installed licenses directory.

Build/test tools (not bundled as application functionality): WiX Toolset 4.0.6 (Microsoft Reciprocal License), xUnit.net (Apache-2.0), Microsoft.NET.Test.Sdk (MIT), DocumentFormat.OpenXml (MIT), CMake (BSD-3-Clause), and the selected C++ compiler toolchain. Native MSVC builds use the static runtime; LLVM-MinGW builds use its static C++ runtime with Windows system libraries. Native runtime notices are preserved in docs/licenses/native in the source and included in the installed licenses directory. See respective upstream distributions for complete licenses.

Microsoft, Windows and Excel are trademarks of Microsoft. This project is independently maintained and is not affiliated with Microsoft.
