# SQLite dependency license provenance

These unmodified upstream license and notice files accompany the SQLite hash-cache dependencies. Retrieved on 2026-09-24. Package IDs and versions were verified against the restored NuGet package manifests and Core project assets.

| NuGet package | Version | Declared package license | Upstream attribution |
| --- | --- | --- | --- |
| Microsoft.Data.Sqlite | 10.0.12 | MIT | Microsoft Corporation; upstream EF Core license credits .NET Foundation and Contributors |
| Microsoft.Data.Sqlite.Core | 10.0.12 | MIT | Microsoft Corporation; upstream EF Core license credits .NET Foundation and Contributors |
| SQLitePCLRaw.bundle_e_sqlite3 | 2.1.12 | Apache-2.0 | Copyright 2014-2024 SourceGear, LLC |
| SQLitePCLRaw.core | 2.1.12 | Apache-2.0 | Copyright 2014-2024 SourceGear, LLC |
| SQLitePCLRaw.lib.e_sqlite3 | 2.1.12 | Apache-2.0 | Copyright 2014-2024 SourceGear, LLC |
| SQLitePCLRaw.provider.e_sqlite3 | 2.1.12 | Apache-2.0 | Copyright 2014-2024 SourceGear, LLC |

Microsoft.Data.Sqlite's NuGet manifests identify the source repository commit as 95017c711e6afc1085133d440e42b4bd78155701 in dotnet/dotnet. The MIT text here is copied from that exact commit's src/efcore/LICENSE.txt, without replacing its upstream copyright attribution. The matching release license is also available at [dotnet/efcore v10.0.12](https://github.com/dotnet/efcore/blob/v10.0.12/LICENSE.txt).

SQLitePCLRaw license and NOTICE are copied from the official v2.1.12 tag. The complete upstream NOTICE is retained; it includes notices for alternative providers such as SQLCipher and OpenSSL, which are not application dependencies in this build.

The native SQLite code shipped by SQLitePCLRaw.lib.e_sqlite3 is public domain, separately from the package's Apache-2.0 metadata. SQLite-PUBLIC-DOMAIN.html is an unmodified copy of SQLite's official public-domain statement; the upstream SQLite blessing is also retained in SQLitePCLRaw-NOTICE.txt.

## Exact downloaded sources

- [Microsoft.Data.Sqlite-LICENSE.txt](https://raw.githubusercontent.com/dotnet/dotnet/95017c711e6afc1085133d440e42b4bd78155701/src/efcore/LICENSE.txt) — 1115 bytes; SHA-256: `ae48df11a335dc1a615f4f938b69cba73bcf4485c4f97af49b38efb0f216353b`
- [SQLitePCLRaw-LICENSE.txt](https://raw.githubusercontent.com/ericsink/SQLitePCL.raw/v2.1.12/LICENSE.TXT) — 11358 bytes; SHA-256: `cfc7749b96f63bd31c3c42b5c471bf756814053e847c10f3eb003417bc523d30`
- [SQLite-PUBLIC-DOMAIN.html](https://sqlite.org/copyright.html) — 8351 bytes; SHA-256: `44ca9f793055c8e32fc65f65a4f5bcf813a33f5bdaaa084067dd617a4ed3cc70`
- [SQLitePCLRaw-NOTICE.txt](https://raw.githubusercontent.com/ericsink/SQLitePCL.raw/v2.1.12/NOTICE.TXT) — 9756 bytes; SHA-256: `485b276b3d2bfaa26df348e1e5c84df3648981e09b88531a6a53006f0705c24b`
