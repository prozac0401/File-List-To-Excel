# Native runtime notices

These unmodified upstream notices accompany native shell builds made with the
portable LLVM-MinGW fallback. The compiler toolchain itself is not distributed
with File List to Excel.

Audited build: LLVM-MinGW **20260922**, Clang **23.1.2**, Windows x64 UCRT.

- Release: https://github.com/mstorsjo/llvm-mingw/releases/tag/20260922
- Archive: `llvm-mingw-20260922-ucrt-x86_64.zip`
- Verified SHA-256: `e3ad77d117a4bea19a7a3b333341824d79a5a371004a10e25b8504e7b3047666`
- `LLVM-LICENSE.txt` is copied from the toolchain root `LICENSE.TXT`.
- `COPYING.MinGW-w64*.txt` and `COPYING.winpthreads.txt` are copied from
  `x86_64-w64-mingw32/share/mingw32/` in that archive.
- Component `*-LICENSE.txt` and available `*-CREDITS.txt` files are copied
  from the matching `llvmorg-23.1.2` source tag:
  https://github.com/llvm/llvm-project/tree/llvmorg-23.1.2

The native fallback statically links libc++, its ABI/unwind support, compiler
runtime support, and MinGW-w64 support as selected by the linker. The UCRT and
Win32 API DLLs are Windows components and are not copied into the MSI.
An import-table inspection of the tested DLL found only Windows system DLLs;
no additional C++ runtime DLL is needed.

MSVC builds statically link the Microsoft C++ runtime instead and do not
incorporate LLVM-MinGW runtime code. Preserve applicable toolchain notices
when changing the shipping compiler. If the portable toolchain version changes,
refresh this directory from the corresponding release and source tag.

The repository build packages this directory along with the other product
third-party notices. .NET runtime notices are collected separately from the
runtime pack used by the self-contained helper build.
