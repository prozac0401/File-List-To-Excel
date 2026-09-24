# Third-party notices

The Windows distribution contains the Microsoft .NET runtime and Windows Desktop runtime, distributed under their upstream licenses. The installer includes the exact runtime packages’ license and third-party notice files under the installed licenses directory. Sources and license information: https://github.com/dotnet/runtime and https://github.com/dotnet/winforms.

The application writes Office Open XML using .NET XML and ZIP APIs; no Excel automation library is included. Microsoft Excel is a separately installed product and is not distributed.

The duplicate finder hash cache uses Microsoft.Data.Sqlite and Microsoft.Data.Sqlite.Core 10.0.12 (MIT; Microsoft Corporation, with the upstream license attribution to .NET Foundation and Contributors), plus SQLitePCLRaw.bundle_e_sqlite3, SQLitePCLRaw.core, SQLitePCLRaw.provider.e_sqlite3 and SQLitePCLRaw.lib.e_sqlite3 2.1.12 (Apache-2.0; Copyright 2014-2024 SourceGear, LLC). The native SQLite engine is public domain. Exact upstream license text, the complete SQLitePCLRaw NOTICE, SQLite’s public-domain statement and source provenance are preserved in [docs/licenses/sqlite](docs/licenses/sqlite/README.md). Sources: [Microsoft.Data.Sqlite license at the package source commit](https://github.com/dotnet/dotnet/blob/95017c711e6afc1085133d440e42b4bd78155701/src/efcore/LICENSE.txt), [SQLitePCLRaw v2.1.12](https://github.com/ericsink/SQLitePCL.raw/tree/v2.1.12), and [SQLite copyright](https://sqlite.org/copyright.html).

The MSI includes the WiX Util 4.0.6 native architecture query custom action (Microsoft Reciprocal License with its generated-output exception). The original WiX license is included in the installed licenses directory.

Build/test tools (not bundled as application functionality): WiX Toolset 4.0.6 (Microsoft Reciprocal License), xUnit.net (Apache-2.0), Microsoft.NET.Test.Sdk (MIT), DocumentFormat.OpenXml (MIT), CMake (BSD-3-Clause), and the selected C++ compiler toolchain. Native MSVC builds use the static runtime; LLVM-MinGW builds use its static C++ runtime with Windows system libraries. Native runtime notices are preserved in docs/licenses/native in the source and included in the installed licenses directory. See respective upstream distributions for complete licenses.

Microsoft, Windows and Excel are trademarks of Microsoft. This project is independently maintained and is not affiliated with Microsoft.
