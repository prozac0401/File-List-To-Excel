#include <windows.h>
#include <shlobj.h>
#include <shellapi.h>
#include <sddl.h>
#include <atomic>
#include <cstring>
#include <cwchar>
#include <new>
#include <string>
#include <vector>
#include <utility>
#include "RequestCodec.h"

namespace {
HMODULE moduleHandle = nullptr;
std::atomic<long> objectCount{0};
std::atomic<long> serverLocks{0};

class Handle {
public:
    HANDLE value;
    explicit Handle(HANDLE handle = INVALID_HANDLE_VALUE) : value(handle) {}
    ~Handle() { if (value && value != INVALID_HANDLE_VALUE) CloseHandle(value); }
    Handle(const Handle&) = delete;
    Handle& operator=(const Handle&) = delete;
};
class Medium {
public:
    STGMEDIUM value{};
    ~Medium() { if (value.tymed != TYMED_NULL) ReleaseStgMedium(&value); }
};
class LocalAllocation {
public:
    HLOCAL value = nullptr;
    ~LocalAllocation() { if (value) LocalFree(value); }
};
template <typename T> class ComPtr {
public:
    T* value = nullptr;
    ~ComPtr() { if (value) value->Release(); }
    T** put() { return &value; }
};

HRESULT LastErrorResult() {
    DWORD error = GetLastError();
    return HRESULT_FROM_WIN32(error ? error : ERROR_GEN_FAILURE);
}
bool IsSafeDirectory(const std::wstring& path) {
    DWORD attributes = GetFileAttributesW(path.c_str());
    return attributes != INVALID_FILE_ATTRIBUTES &&
        (attributes & FILE_ATTRIBUTE_DIRECTORY) && !(attributes & FILE_ATTRIBUTE_REPARSE_POINT);
}

// This per-user spool is never shared across users. A random CREATE_NEW file plus
// a protected DACL avoids predictable temp-file replacement and cross-user reads.
HRESULT WriteRequest(const std::string& payload, std::wstring& requestPath) {
    HANDLE rawToken = nullptr;
    if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &rawToken)) return LastErrorResult();
    Handle token(rawToken);
    DWORD tokenBytes = 0;
    GetTokenInformation(token.value, TokenUser, nullptr, 0, &tokenBytes);
    if (!tokenBytes) return LastErrorResult();
    std::vector<BYTE> tokenBuffer(tokenBytes);
    if (!GetTokenInformation(token.value, TokenUser, tokenBuffer.data(), tokenBytes, &tokenBytes))
        return LastErrorResult();
    LPWSTR sidText = nullptr;
    if (!ConvertSidToStringSidW(reinterpret_cast<TOKEN_USER*>(tokenBuffer.data())->User.Sid, &sidText))
        return LastErrorResult();
    LocalAllocation sidAllocation;
    sidAllocation.value = sidText;
    std::wstring sddl = L"D:P(A;;FA;;;SY)(A;;FA;;;" + std::wstring(sidText) + L")";
    PSECURITY_DESCRIPTOR descriptor = nullptr;
    if (!ConvertStringSecurityDescriptorToSecurityDescriptorW(sddl.c_str(), SDDL_REVISION_1, &descriptor, nullptr))
        return LastErrorResult();
    LocalAllocation descriptorAllocation;
    descriptorAllocation.value = descriptor;
    SECURITY_ATTRIBUTES security{sizeof(SECURITY_ATTRIBUTES), descriptor, FALSE};

    PWSTR appData = nullptr;
    HRESULT hr = SHGetKnownFolderPath(FOLDERID_LocalAppData, KF_FLAG_CREATE, nullptr, &appData);
    if (FAILED(hr)) return hr;
    std::wstring directory(appData);
    CoTaskMemFree(appData);
    directory += L"\\FileListToExcel";
    if (!CreateDirectoryW(directory.c_str(), &security) && GetLastError() != ERROR_ALREADY_EXISTS)
        return LastErrorResult();
    if (!IsSafeDirectory(directory)) return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    directory += L"\\Requests";
    if (!CreateDirectoryW(directory.c_str(), &security) && GetLastError() != ERROR_ALREADY_EXISTS)
        return LastErrorResult();
    if (!IsSafeDirectory(directory)) return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);

    GUID id{};
    hr = CoCreateGuid(&id);
    if (FAILED(hr)) return hr;
    wchar_t idText[40]{};
    if (!StringFromGUID2(id, idText, 40)) return E_FAIL;
    std::wstring basename(idText + 1, 36);
    for (wchar_t& c : basename) if (c >= L'A' && c <= L'F') c = static_cast<wchar_t>(c + L'a' - L'A');
    requestPath = directory + L"\\" + basename + L".json";
    HRESULT writeResult = S_OK;
    {
        Handle file(CreateFileW(requestPath.c_str(), GENERIC_WRITE, 0, &security, CREATE_NEW,
            FILE_ATTRIBUTE_TEMPORARY, nullptr));
        if (file.value == INVALID_HANDLE_VALUE) return LastErrorResult();
        DWORD written = 0;
        if (!WriteFile(file.value, payload.data(), static_cast<DWORD>(payload.size()), &written, nullptr))
            writeResult = LastErrorResult();
        else if (written != payload.size()) writeResult = HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
    }
    if (FAILED(writeResult)) { DeleteFileW(requestPath.c_str()); requestPath.clear(); }
    return writeResult;
}

HRESULT LaunchHelper(HWND owner, const char* mode, const std::vector<std::wstring>& paths) {
    std::wstring requestPath;
    try {
        std::vector<wchar_t> buffer(32768);
        DWORD length = GetModuleFileNameW(moduleHandle, buffer.data(), static_cast<DWORD>(buffer.size()));
        if (!length || length >= buffer.size()) return HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME);
        std::wstring directory(buffer.data(), length);
        const size_t separator = directory.find_last_of(L"\\/");
        if (separator == std::wstring::npos) return E_FAIL;
        directory.resize(separator);
        const std::wstring executable = directory + L"\\FileListToExcel.exe";
        HRESULT hr = WriteRequest(filelist::BuildRequest(mode, paths), requestPath);
        if (FAILED(hr)) return hr;
        std::wstring command = filelist::QuoteArgument(executable) + L" --request " + filelist::QuoteArgument(requestPath);
        STARTUPINFOW startup{};
        startup.cb = sizeof(startup);
        PROCESS_INFORMATION process{};
        if (!CreateProcessW(executable.c_str(), command.data(), nullptr, nullptr, FALSE, CREATE_NO_WINDOW,
            nullptr, directory.c_str(), &startup, &process)) {
            hr = LastErrorResult();
            DeleteFileW(requestPath.c_str());
            MessageBoxW(owner, L"File List to Excel을 시작하지 못했습니다. 프로그램을 다시 설치해 주세요.",
                L"File List to Excel", MB_OK | MB_ICONERROR);
            return hr;
        }
        CloseHandle(process.hThread);
        CloseHandle(process.hProcess);
        return S_OK;
    } catch (const std::bad_alloc&) {
        if (!requestPath.empty()) DeleteFileW(requestPath.c_str());
        return E_OUTOFMEMORY;
    } catch (...) {
        if (!requestPath.empty()) DeleteFileW(requestPath.c_str());
        return E_FAIL;
    }
}

class ShellExtension final : public IShellExtInit, public IContextMenu {
    std::atomic<ULONG> references_{1};
    std::vector<std::wstring> paths_;
    bool background_ = false;
    bool hasFolder_ = false;
    bool hasFile_ = false;
    UINT commandCount_ = 0;

    const char* Mode(UINT command) const { return command == 1 ? "recursive" : (hasFolder_ && !hasFile_ ? "folder" : "files"); }
    const wchar_t* Verb(UINT command) const { return command == 1 ? L"filelisttoexcelrecursive" : L"filelisttoexcel"; }
    const wchar_t* Help(UINT command) const {
        return command == 1 ? L"선택 폴더의 하위 항목을 포함하여 Excel 목록을 만듭니다." :
            L"현재 선택한 파일 또는 폴더 내용으로 Excel 목록을 만듭니다.";
    }
public:
    ShellExtension() { ++objectCount; }
    ~ShellExtension() { --objectCount; }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** result) override {
        if (!result) return E_POINTER;
        *result = nullptr;
        if (iid == IID_IUnknown || iid == IID_IShellExtInit) *result = static_cast<IShellExtInit*>(this);
        else if (iid == IID_IContextMenu) *result = static_cast<IContextMenu*>(this);
        else return E_NOINTERFACE;
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return ++references_; }
    ULONG STDMETHODCALLTYPE Release() override {
        ULONG remaining = --references_;
        if (!remaining) delete this;
        return remaining;
    }
    HRESULT STDMETHODCALLTYPE Initialize(PCIDLIST_ABSOLUTE folder, IDataObject* data, HKEY) override {
        try {
            paths_.clear();
            std::vector<std::wstring> selected;
            background_ = hasFolder_ = hasFile_ = false;
            commandCount_ = 0;
            if (data) {
                FORMATETC format{CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
                Medium medium;
                HRESULT hr = data->GetData(&format, &medium.value);
                if (FAILED(hr)) return hr;
                if (medium.value.tymed != TYMED_HGLOBAL || !medium.value.hGlobal) return E_INVALIDARG;
                HDROP drop = reinterpret_cast<HDROP>(medium.value.hGlobal);
                UINT count = DragQueryFileW(drop, 0xFFFFFFFF, nullptr, 0);
                if (!count || count > filelist::MaximumPaths) return E_INVALIDARG;
                selected.reserve(count);
                for (UINT index = 0; index < count; ++index) {
                    UINT length = DragQueryFileW(drop, index, nullptr, 0);
                    if (!length || length >= 32768) return E_INVALIDARG;
                    std::vector<wchar_t> path(static_cast<size_t>(length) + 1);
                    if (DragQueryFileW(drop, index, path.data(), length + 1) != length) return E_FAIL;
                    selected.emplace_back(path.data(), length);
                }
                // Explorer's shell item array answers these from its cached item
                // attributes, avoiding a filesystem call for every selected file.
                ComPtr<IShellItemArray> items;
                if (SUCCEEDED(SHCreateShellItemArrayFromDataObject(data, IID_IShellItemArray,
                    reinterpret_cast<void**>(items.put())))) {
                    // ZIP files can advertise SFGAO_FOLDER as well as
                    // SFGAO_STREAM. They remain files for this product.
                    for (UINT index = 0; index < count; ++index) {
                        ComPtr<IShellItem> item;
                        SFGAOF attributes = 0;
                        if (FAILED(items.value->GetItemAt(index, item.put())) ||
                            FAILED(item.value->GetAttributes(SFGAO_FOLDER | SFGAO_STREAM, &attributes))) {
                            hasFolder_ = hasFile_ = false;
                            break;
                        }
                        if ((attributes & SFGAO_FOLDER) && !(attributes & SFGAO_STREAM)) hasFolder_ = true;
                        else hasFile_ = true;
                        if (hasFolder_ && hasFile_) break;
                    }
                }
                if (!hasFolder_ && !hasFile_) {
                    for (const auto& path : selected) {
                        DWORD attributes = GetFileAttributesW(path.c_str());
                        if (attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY))
                            hasFolder_ = true;
                        else hasFile_ = true;
                        if (hasFolder_ && hasFile_) break;
                    }
                }
            } else if (folder) {
                std::vector<wchar_t> path(32768);
                if (!SHGetPathFromIDListEx(folder, path.data(), static_cast<DWORD>(path.size()), GPFIDL_DEFAULT))
                    return E_INVALIDARG; // Virtual folders have no filesystem input.
                selected.emplace_back(path.data());
                background_ = hasFolder_ = true;
            } else return E_INVALIDARG;
            paths_ = std::move(selected);
            return S_OK;
        } catch (const std::bad_alloc&) { paths_.clear(); return E_OUTOFMEMORY; }
        catch (...) { paths_.clear(); return E_FAIL; }
    }
    HRESULT STDMETHODCALLTYPE QueryContextMenu(HMENU menu, UINT position, UINT firstId, UINT lastId, UINT flags) override {
        commandCount_ = 0;
        if ((flags & CMF_DEFAULTONLY) || paths_.empty() || firstId > lastId) return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
        const wchar_t* primary = background_ ? L"현재 폴더 목록을 Excel로" :
            (!hasFolder_ ? L"파일 목록을 Excel로" : (hasFile_ ? L"선택 항목을 Excel로" : L"폴더 내용 목록을 Excel로"));
        if (!InsertMenuW(menu, position, MF_BYPOSITION | MF_STRING, firstId, primary)) return LastErrorResult();
        commandCount_ = 1;
        if (hasFolder_ && lastId > firstId) {
            const wchar_t* recursive = background_ ? L"현재 폴더 + 하위 폴더를 Excel로" :
                (hasFile_ ? L"선택 폴더 + 하위 폴더를 Excel로" : L"하위 폴더까지 Excel로");
            if (InsertMenuW(menu, position + 1, MF_BYPOSITION | MF_STRING, firstId + 1, recursive)) commandCount_ = 2;
        }
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, commandCount_);
    }
    HRESULT STDMETHODCALLTYPE InvokeCommand(CMINVOKECOMMANDINFO* info) override {
        if (!info || info->cbSize < sizeof(CMINVOKECOMMANDINFO) || paths_.empty()) return E_INVALIDARG;
        UINT command = 0;
        bool byName = false;
        if (info->cbSize >= sizeof(CMINVOKECOMMANDINFOEX) && (info->fMask & CMIC_MASK_UNICODE)) {
            auto* extended = reinterpret_cast<CMINVOKECOMMANDINFOEX*>(info);
            if (!IS_INTRESOURCE(extended->lpVerbW)) {
                byName = true;
                if (_wcsicmp(extended->lpVerbW, Verb(0)) == 0) command = 0;
                else if (_wcsicmp(extended->lpVerbW, Verb(1)) == 0) command = 1;
                else return E_INVALIDARG;
            }
        }
        if (!byName && !IS_INTRESOURCE(info->lpVerb)) {
            byName = true;
            if (_stricmp(info->lpVerb, "filelisttoexcel") == 0) command = 0;
            else if (_stricmp(info->lpVerb, "filelisttoexcelrecursive") == 0) command = 1;
            else return E_INVALIDARG;
        }
        if (!byName) command = LOWORD(info->lpVerb);
        if (command > 1 || (command == 1 && !hasFolder_) || (!byName && command >= commandCount_)) return E_INVALIDARG;
        return LaunchHelper(info->hwnd, Mode(command), paths_);
    }
    HRESULT STDMETHODCALLTYPE GetCommandString(UINT_PTR id, UINT flags, UINT*, LPSTR result, UINT count) override {
        if (id >= commandCount_) return E_INVALIDARG;
        if (flags == GCS_VALIDATEA || flags == GCS_VALIDATEW) return S_OK;
        if (!result || !count) return E_POINTER;
        const bool verb = flags == GCS_VERBA || flags == GCS_VERBW;
        const bool unicode = flags == GCS_VERBW || flags == GCS_HELPTEXTW;
        if (!verb && flags != GCS_HELPTEXTA && flags != GCS_HELPTEXTW) return E_INVALIDARG;
        if (unicode) {
            const wchar_t* text = verb ? Verb(static_cast<UINT>(id)) : Help(static_cast<UINT>(id));
            size_t length = wcslen(text);
            if (length >= count) return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
            memcpy(result, text, (length + 1) * sizeof(wchar_t));
        } else {
            const char* text = verb ? (id == 1 ? "filelisttoexcelrecursive" : "filelisttoexcel") :
                "Create an Excel workbook from the selected filesystem items.";
            size_t length = strlen(text);
            if (length >= count) return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
            memcpy(result, text, length + 1);
        }
        return S_OK;
    }
};

class ClassFactory final : public IClassFactory {
    std::atomic<ULONG> references_{1};
public:
    ClassFactory() { ++objectCount; }
    ~ClassFactory() { --objectCount; }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** result) override {
        if (!result) return E_POINTER;
        *result = nullptr;
        if (iid != IID_IUnknown && iid != IID_IClassFactory) return E_NOINTERFACE;
        *result = static_cast<IClassFactory*>(this);
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return ++references_; }
    ULONG STDMETHODCALLTYPE Release() override {
        ULONG remaining = --references_;
        if (!remaining) delete this;
        return remaining;
    }
    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* outer, REFIID iid, void** result) override {
        if (!result) return E_POINTER;
        *result = nullptr;
        if (outer) return CLASS_E_NOAGGREGATION;
        auto* instance = new (std::nothrow) ShellExtension();
        if (!instance) return E_OUTOFMEMORY;
        HRESULT hr = instance->QueryInterface(iid, result);
        instance->Release();
        return hr;
    }
    HRESULT STDMETHODCALLTYPE LockServer(BOOL lock) override {
        if (lock) ++serverLocks;
        else {
            long current = serverLocks.load();
            while (current > 0 && !serverLocks.compare_exchange_weak(current, current - 1)) {}
        }
        return S_OK;
    }
};
}
BOOL WINAPI DllMain(HINSTANCE instance, DWORD reason, LPVOID) {
    if (reason == DLL_PROCESS_ATTACH) moduleHandle = instance;
    // Keep thread notifications enabled for the statically linked C++ runtime.
    return TRUE;
}
extern "C" HRESULT __stdcall DllCanUnloadNow() {
    return objectCount.load() == 0 && serverLocks.load() == 0 ? S_OK : S_FALSE;
}
extern "C" HRESULT __stdcall DllGetClassObject(REFCLSID clsid, REFIID iid, void** result) {
    if (!result) return E_POINTER;
    *result = nullptr;
    if (clsid != filelist::ShellClsid) return CLASS_E_CLASSNOTAVAILABLE;
    auto* factory = new (std::nothrow) ClassFactory();
    if (!factory) return E_OUTOFMEMORY;
    HRESULT hr = factory->QueryInterface(iid, result);
    factory->Release();
    return hr;
}
