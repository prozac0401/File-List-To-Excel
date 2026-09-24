#pragma once
#include <windows.h>
#include <stdexcept>
#include <string>
#include <vector>

namespace filelist {
inline constexpr CLSID ShellClsid = {0xaf7a7218, 0x8b84, 0x4a11, {0x97, 0x90, 0xeb, 0x24, 0xa4, 0x3e, 0xc3, 0x9e}};
inline constexpr size_t MaximumPaths = 100000;
inline constexpr size_t MaximumRequestBytes = 32 * 1024 * 1024;

inline std::string Utf8(const std::wstring& value) {
    if (value.empty()) return {};
    int size = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, value.data(),
        static_cast<int>(value.size()), nullptr, 0, nullptr, nullptr);
    if (size == 0) throw std::runtime_error("Invalid Unicode path");
    std::string result(static_cast<size_t>(size), '\0');
    if (WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, value.data(),
        static_cast<int>(value.size()), result.data(), size, nullptr, nullptr) != size)
        throw std::runtime_error("Cannot convert path");
    return result;
}

inline std::string JsonString(const std::wstring& value) {
    std::string result = "\"";
    const auto utf8 = Utf8(value);
    constexpr char hex[] = "0123456789abcdef";
    for (unsigned char c : utf8) {
        if (c == '\\' || c == '"') { result += '\\'; result += static_cast<char>(c); }
        else if (c < 0x20) { result += "\\u00"; result += hex[c >> 4]; result += hex[c & 15]; }
        else result += static_cast<char>(c);
    }
    result += '"';
    return result;
}

inline std::string BuildRequest(const char* mode, const std::vector<std::wstring>& paths) {
    if (paths.empty() || paths.size() > MaximumPaths) throw std::runtime_error("Invalid selection size");
    std::string result = std::string("{\"mode\":\"") + mode + "\",\"paths\":[";
    for (size_t i = 0; i < paths.size(); ++i) {
        if (i) result += ',';
        result += JsonString(paths[i]);
        if (result.size() > MaximumRequestBytes - 2) throw std::runtime_error("Selection is too large");
    }
    result += "]}";
    return result;
}

// Windows CommandLineToArgvW/CRT quoting. No command shell is involved.
inline std::wstring QuoteArgument(const std::wstring& value) {
    std::wstring result = L"\"";
    size_t slashes = 0;
    for (wchar_t c : value) {
        if (c == L'\\') { ++slashes; continue; }
        if (c == L'"') result.append(slashes * 2 + 1, L'\\');
        else result.append(slashes, L'\\');
        result += c;
        slashes = 0;
    }
    result.append(slashes * 2, L'\\');
    return result + L'"';
}
}
