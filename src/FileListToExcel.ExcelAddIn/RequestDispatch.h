#pragma once
#include "../FileListToExcel.Shell/RequestCodec.h"
#include <memory>
#include <sddl.h>
#include <shlobj.h>
namespace flt
{
class Handle
{
  public:
    HANDLE value;
    explicit Handle(HANDLE handle = INVALID_HANDLE_VALUE) : value(handle)
    {
    }
    ~Handle()
    {
        if (value && value != INVALID_HANDLE_VALUE)
            CloseHandle(value);
    }
    Handle(const Handle &) = delete;
    Handle &operator=(const Handle &) = delete;
};
class Medium
{
  public:
    STGMEDIUM value{};
    ~Medium()
    {
        if (value.tymed != TYMED_NULL)
            ReleaseStgMedium(&value);
    }
};
class LocalAllocation
{
  public:
    HLOCAL value = nullptr;
    ~LocalAllocation()
    {
        if (value)
            LocalFree(value);
    }
};
template <typename T> class ComPtr
{
  public:
    T *value = nullptr;
    ~ComPtr()
    {
        if (value)
            value->Release();
    }
    T **put()
    {
        return &value;
    }
};

HRESULT LastErrorResult()
{
    DWORD error = GetLastError();
    return HRESULT_FROM_WIN32(error ? error : ERROR_GEN_FAILURE);
}
bool IsSafeDirectory(const std::wstring &path)
{
    DWORD attributes = GetFileAttributesW(path.c_str());
    return attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY) &&
           !(attributes & FILE_ATTRIBUTE_REPARSE_POINT);
}

// Pin every ancestor against rename/delete; OPEN_REPARSE_POINT prevents junction traversal.
HRESULT PinDirectory(const std::wstring &path, std::vector<std::unique_ptr<Handle>> &guards)
{
    auto guard = std::make_unique<Handle>(CreateFileW(
        path.c_str(), FILE_READ_ATTRIBUTES, FILE_SHARE_READ, nullptr, OPEN_EXISTING,
        FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_OPEN_NO_RECALL, nullptr));
    if (guard->value == INVALID_HANDLE_VALUE)
        return LastErrorResult();
    BY_HANDLE_FILE_INFORMATION info{};
    if (!GetFileInformationByHandle(guard->value, &info) ||
        !(info.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) ||
        (info.dwFileAttributes &
         (FILE_ATTRIBUTE_REPARSE_POINT | FILE_ATTRIBUTE_OFFLINE | 0x40000 | 0x400000)))
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    guards.push_back(std::move(guard));
    return S_OK;
}

// This per-user spool is never shared across users. A random CREATE_NEW file plus
// a protected DACL avoids predictable temp-file replacement and cross-user reads.
HRESULT WriteRequest(const std::string &payload, const std::wstring &requestId, std::wstring &requestPath)
{
    HANDLE rawToken = nullptr;
    if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &rawToken))
        return LastErrorResult();
    Handle token(rawToken);
    DWORD tokenBytes = 0;
    GetTokenInformation(token.value, TokenUser, nullptr, 0, &tokenBytes);
    if (!tokenBytes)
        return LastErrorResult();
    std::vector<BYTE> tokenBuffer(tokenBytes);
    if (!GetTokenInformation(token.value, TokenUser, tokenBuffer.data(), tokenBytes, &tokenBytes))
        return LastErrorResult();
    LPWSTR sidText = nullptr;
    if (!ConvertSidToStringSidW(reinterpret_cast<TOKEN_USER *>(tokenBuffer.data())->User.Sid, &sidText))
        return LastErrorResult();
    LocalAllocation sidAllocation;
    sidAllocation.value = sidText;
    std::wstring sddl = L"D:P(A;;FA;;;SY)(A;;FA;;;" + std::wstring(sidText) + L")";
    PSECURITY_DESCRIPTOR descriptor = nullptr;
    if (!ConvertStringSecurityDescriptorToSecurityDescriptorW(sddl.c_str(), SDDL_REVISION_1, &descriptor,
                                                              nullptr))
        return LastErrorResult();
    LocalAllocation descriptorAllocation;
    descriptorAllocation.value = descriptor;
    SECURITY_ATTRIBUTES security{sizeof(SECURITY_ATTRIBUTES), descriptor, FALSE};

    PWSTR appData = nullptr;
    HRESULT hr = SHGetKnownFolderPath(FOLDERID_LocalAppData, KF_FLAG_CREATE, nullptr, &appData);
    if (FAILED(hr))
        return hr;
    std::wstring directory(appData);
    CoTaskMemFree(appData);
    std::vector<std::unique_ptr<Handle>> guards;
    if (directory.size() < 3 || directory[1] != L':' || directory[2] != L'\\')
        return E_INVALIDARG;
    for (size_t end = 3; end <= directory.size(); ++end)
    {
        if (end != directory.size() && directory[end] != L'\\')
            continue;
        hr = PinDirectory(directory.substr(0, end), guards);
        if (FAILED(hr))
            return hr;
    }
    directory += L"\\FileListToExcel";
    if (!CreateDirectoryW(directory.c_str(), &security) && GetLastError() != ERROR_ALREADY_EXISTS)
        return LastErrorResult();
    hr = PinDirectory(directory, guards);
    if (FAILED(hr))
        return hr;
    directory += L"\\Requests";
    if (!CreateDirectoryW(directory.c_str(), &security) && GetLastError() != ERROR_ALREADY_EXISTS)
        return LastErrorResult();
    hr = PinDirectory(directory, guards);
    if (FAILED(hr))
        return hr;

    requestPath = directory + L"\\" + requestId + L".json";
    HRESULT writeResult = S_OK;
    {
        Handle file(CreateFileW(requestPath.c_str(), GENERIC_WRITE | DELETE, 0, &security, CREATE_NEW,
                                FILE_ATTRIBUTE_TEMPORARY, nullptr));
        if (file.value == INVALID_HANDLE_VALUE)
            return LastErrorResult();
        DWORD written = 0;
        if (!WriteFile(file.value, payload.data(), static_cast<DWORD>(payload.size()), &written, nullptr))
            writeResult = LastErrorResult();
        else if (written != payload.size())
            writeResult = HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
        if (FAILED(writeResult))
        {
            FILE_DISPOSITION_INFO disposition{TRUE};
            SetFileInformationByHandle(file.value, FileDispositionInfo, &disposition, sizeof(disposition));
        }
    }
    if (FAILED(writeResult))
        requestPath.clear();
    return writeResult;
}

} // namespace flt
