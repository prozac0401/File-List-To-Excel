// Test-only helper, built into the native test output directory.
// It never ships in the installer; the real app is published separately.
#include <windows.h>
#include <filesystem>
#include <fstream>
#include <iterator>
#include <string>

int wmain(int argc, wchar_t** argv) {
    if (argc != 3 || std::wstring(argv[1]) != L"--request") return 2;
    wchar_t capture[32768]{};
    DWORD length = GetEnvironmentVariableW(L"FILELISTTOEXCEL_TEST_CAPTURE", capture, 32768);
    if (!length || length >= 32768) return 3;
    std::filesystem::path source(argv[2]), target(capture), temporary(std::wstring(capture) + L".pending");
    std::ifstream input(source, std::ios::binary);
    if (!input) return 4;
    std::string content(std::istreambuf_iterator<char>(input), {});
    input.close();
    std::ofstream output(temporary, std::ios::binary);
    output.write(content.data(), static_cast<std::streamsize>(content.size()));
    output.close();
    if (!output) return 5;
    // Publish completion atomically so the parent never sees a partial capture.
    std::filesystem::rename(temporary, target);
    std::filesystem::remove(source);
    return 0;
}
