#include <windows.h>
#include <shlobj.h>
#include <shellapi.h>
#include <atomic>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <iterator>
#include <stdexcept>
#include <string>
#include <vector>
#include "../RequestCodec.h"

namespace {
int checks = 0;
void Check(bool condition, const char* name) {
    ++checks;
    if (!condition) throw std::runtime_error(name);
}
void Hr(HRESULT value, const char* name) { Check(SUCCEEDED(value), name); }

class DropData final : public IDataObject {
    std::atomic<ULONG> references_{1};
    std::vector<std::wstring> paths_;
public:
    explicit DropData(std::vector<std::wstring> paths) : paths_(std::move(paths)) {}
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** result) override {
        if (!result) return E_POINTER;
        *result = nullptr;
        if (iid != IID_IUnknown && iid != IID_IDataObject) return E_NOINTERFACE;
        *result = static_cast<IDataObject*>(this);
        AddRef(); return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return ++references_; }
    ULONG STDMETHODCALLTYPE Release() override { auto remaining = --references_; if (!remaining) delete this; return remaining; }
    HRESULT STDMETHODCALLTYPE GetData(FORMATETC* format, STGMEDIUM* output) override {
        if (!format || !output) return E_POINTER;
        if (format->cfFormat != CF_HDROP || !(format->tymed & TYMED_HGLOBAL)) return DV_E_FORMATETC;
        size_t characters = 1;
        for (const auto& path : paths_) characters += path.size() + 1;
        HGLOBAL memory = GlobalAlloc(GHND, sizeof(DROPFILES) + characters * sizeof(wchar_t));
        if (!memory) return E_OUTOFMEMORY;
        auto* drop = static_cast<DROPFILES*>(GlobalLock(memory));
        if (!drop) { GlobalFree(memory); return E_OUTOFMEMORY; }
        drop->pFiles = sizeof(DROPFILES);
        drop->fWide = TRUE;
        auto* target = reinterpret_cast<wchar_t*>(reinterpret_cast<BYTE*>(drop) + sizeof(DROPFILES));
        for (const auto& path : paths_) { memcpy(target, path.c_str(), (path.size() + 1) * sizeof(wchar_t)); target += path.size() + 1; }
        *target = 0;
        GlobalUnlock(memory);
        *output = {};
        output->tymed = TYMED_HGLOBAL;
        output->hGlobal = memory;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetDataHere(FORMATETC*, STGMEDIUM*) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE QueryGetData(FORMATETC* format) override { return format && format->cfFormat == CF_HDROP ? S_OK : DV_E_FORMATETC; }
    HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc(FORMATETC*, FORMATETC* output) override { if (output) output->ptd = nullptr; return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE SetData(FORMATETC*, STGMEDIUM*, BOOL) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE EnumFormatEtc(DWORD, IEnumFORMATETC**) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE DAdvise(FORMATETC*, DWORD, IAdviseSink*, DWORD*) override { return OLE_E_ADVISENOTSUPPORTED; }
    HRESULT STDMETHODCALLTYPE DUnadvise(DWORD) override { return OLE_E_ADVISENOTSUPPORTED; }
    HRESULT STDMETHODCALLTYPE EnumDAdvise(IEnumSTATDATA**) override { return OLE_E_ADVISENOTSUPPORTED; }
};

struct Fixture {
    std::filesystem::path directory;
    Fixture() {
        wchar_t temp[32768]{};
        DWORD tempLength = GetTempPathW(32768, temp);
        Check(tempLength > 0 && tempLength < 32768, "GetTempPath");
        // Hosted Windows runners may expose TEMP through an 8.3 username
        // (RUNNER~1). Explorer resolves PIDLs to long names. Normalize the
        // fixture once so exact request comparisons use the same real path.
        DWORD longLength = GetLongPathNameW(temp, nullptr, 0);
        Check(longLength != 0, "Measure long temporary directory path");
        std::vector<wchar_t> longTemp(longLength);
        DWORD resolvedLength = GetLongPathNameW(temp, longTemp.data(), longLength);
        Check(resolvedLength > 0 && resolvedLength < longLength, "Resolve long temporary directory path");
        GUID id{};
        Hr(CoCreateGuid(&id), "Test GUID");
        wchar_t text[40]{};
        StringFromGUID2(id, text, 40);
        directory = std::filesystem::path(longTemp.data()) / (std::wstring(L"FileListToExcel-ShellTests-") + text);
        std::filesystem::create_directory(directory);
    }
    ~Fixture() {
        // Only this test's independently generated root is removed.
        std::error_code error;
        std::filesystem::remove_all(directory, error);
    }
    std::wstring File(const std::wstring& name) {
        auto path = directory / name;
        std::ofstream file(path, std::ios::binary);
        file << "test";
        return path.wstring();
    }
    std::wstring Folder(const std::wstring& name) {
        auto path = directory / name;
        std::filesystem::create_directory(path);
        return path.wstring();
    }
};
struct Extension {
    IShellExtInit* init = nullptr;
    IContextMenu* menu = nullptr;
    explicit Extension(IClassFactory* factory) {
        Hr(factory->CreateInstance(nullptr, IID_IShellExtInit, reinterpret_cast<void**>(&init)), "Create IShellExtInit");
        Hr(init->QueryInterface(IID_IContextMenu, reinterpret_cast<void**>(&menu)), "Get IContextMenu");
    }
    ~Extension() { if (menu) menu->Release(); if (init) init->Release(); }
    void Select(const std::vector<std::wstring>& paths) {
        auto* data = new DropData(paths);
        HRESULT hr = init->Initialize(nullptr, data, nullptr);
        data->Release();
        Hr(hr, "Initialize selection");
    }
    void SelectShellData(const std::vector<std::wstring>& paths) {
        std::vector<PIDLIST_ABSOLUTE> ids(paths.size(), nullptr);
        for (size_t i = 0; i < paths.size(); ++i)
            Hr(SHParseDisplayName(paths[i].c_str(), nullptr, &ids[i], 0, nullptr), "Parse native shell selection");
        IShellItemArray* items = nullptr;
        HRESULT hr = SHCreateShellItemArrayFromIDLists(static_cast<UINT>(ids.size()),
            const_cast<PCIDLIST_ABSOLUTE*>(ids.data()), &items);
        for (auto id : ids) CoTaskMemFree(id);
        Hr(hr, "Create real shell item array");
        IDataObject* data = nullptr;
        hr = items->BindToHandler(nullptr, BHID_DataObject, IID_IDataObject, reinterpret_cast<void**>(&data));
        items->Release();
        Hr(hr, "Bind real Explorer data object");
        hr = init->Initialize(nullptr, data, nullptr);
        data->Release();
        Hr(hr, "Initialize real shell data object");
    }
    void CheckMenu(int count, const wchar_t* firstLabel, const wchar_t* duplicateLabel = nullptr, const wchar_t* duplicateVerb = nullptr) {
        HMENU popup = CreatePopupMenu();
        Check(popup != nullptr, "Create menu");
        HRESULT hr = menu->QueryContextMenu(popup, 0, 10, 100, CMF_NORMAL);
        Check(SUCCEEDED(hr) && HRESULT_CODE(hr) == count, "Correct reserved command count");
        Check(GetMenuItemCount(popup) == count, "Correct menu item count");
        wchar_t label[256]{};
        GetMenuStringW(popup, 0, label, 256, MF_BYPOSITION);
        Check(std::wstring(label) == firstLabel, "Correct localized menu label");
        for (int i = 0; i < count; ++i) {
            Check(GetMenuItemID(popup, i) == static_cast<UINT>(10 + i), "Menu identifier is offset");
            wchar_t verb[64]{};
            Hr(menu->GetCommandString(static_cast<UINT_PTR>(i), GCS_VERBW, nullptr, reinterpret_cast<LPSTR>(verb), 64),
                "Unicode canonical verb");
            const wchar_t* expected = duplicateVerb && i == count - 1 ? duplicateVerb :
                (i ? L"filelisttoexcelrecursive" : L"filelisttoexcel");
            Check(std::wstring(verb) == expected, "Canonical verb matches");
            char ansiVerb[64]{};
            Hr(menu->GetCommandString(static_cast<UINT_PTR>(i), GCS_VERBA, nullptr, ansiVerb, 64), "ANSI canonical verb");
            Check(std::string(ansiVerb) == filelist::Utf8(expected), "ANSI canonical verb matches");
            wchar_t help[256]{};
            Hr(menu->GetCommandString(static_cast<UINT_PTR>(i), GCS_HELPTEXTW, nullptr, reinterpret_cast<LPSTR>(help), 256),
                "Localized command help");
            Check(wcslen(help) != 0, "Command help is not empty");
            if (duplicateLabel && i == count - 1) {
                GetMenuStringW(popup, i, label, 256, MF_BYPOSITION);
                Check(std::wstring(label) == duplicateLabel, "Correct duplicate menu label");
            }
        }
        DestroyMenu(popup);
    }
};

void TestCodec() {
    Check(filelist::JsonString(L"C:\\한글\\line\n\t\".txt") ==
        std::string("\"C:\\\\") + filelist::Utf8(L"한글") + "\\\\line\\u000a\\u0009\\\".txt\"", "JSON Unicode/control escaping");
    std::vector<std::wstring> arguments{L"", L"C:\\spaces here\\", L"a\"b", L"C:\\한글\\emoji-\U0001F600.txt"};
    for (const auto& argument : arguments) {
        auto command = std::wstring(L"test.exe ") + filelist::QuoteArgument(argument);
        int count = 0;
        LPWSTR* parsed = CommandLineToArgvW(command.c_str(), &count);
        Check(parsed && count == 2 && parsed[1] == argument, "Windows argv quoting round trip");
        LocalFree(parsed);
    }
    bool rejected = false;
    try { filelist::BuildRequest("files", {}); } catch (...) { rejected = true; }
    Check(rejected, "Reject empty request");
    rejected = false;
    try { filelist::BuildRequest("files", std::vector<std::wstring>(filelist::MaximumPaths + 1, L"x")); }
    catch (...) { rejected = true; }
    Check(rejected, "Bound selection count");
    rejected = false;
    try { filelist::BuildRequest("files", {std::wstring(filelist::MaximumRequestBytes, L'x')}); }
    catch (...) { rejected = true; }
    Check(rejected, "Bound request byte size");
}

void TestDispatch(Extension& extension, Fixture& fixture, const std::vector<std::wstring>& paths, const char* mode, UINT offset, bool canonical = false, bool ansi = false) {
    GUID captureId{};
    Hr(CoCreateGuid(&captureId), "Unique helper capture identity");
    wchar_t captureName[40]{};
    StringFromGUID2(captureId, captureName, 40);
    // Each helper gets its own completion file. Never reuse a pathname while
    // another short-lived helper process may still be finishing its exit.
    auto capture = fixture.directory / (std::wstring(captureName) + L".json");
    Check(SetEnvironmentVariableW(L"FILELISTTOEXCEL_TEST_CAPTURE", capture.c_str()) != 0, "Set sink location");
    CMINVOKECOMMANDINFOEX command{};
    command.cbSize = sizeof(command);
    command.lpVerb = MAKEINTRESOURCEA(offset);
    wchar_t canonicalVerb[64]{};
    char ansiVerb[64]{};
    if (canonical) {
        if (ansi) {
            Hr(extension.menu->GetCommandString(offset, GCS_VERBA, nullptr, ansiVerb, 64), "Read invoked ANSI verb");
            command.lpVerb = ansiVerb;
        } else {
            Hr(extension.menu->GetCommandString(offset, GCS_VERBW, nullptr, reinterpret_cast<LPSTR>(canonicalVerb), 64),
                "Read invoked Unicode verb");
            command.fMask = CMIC_MASK_UNICODE;
            command.lpVerbW = canonicalVerb;
        }
    }
    Hr(extension.menu->InvokeCommand(reinterpret_cast<CMINVOKECOMMANDINFO*>(&command)), "Invoke starts one helper");
    for (int attempt = 0; attempt < 200 && !std::filesystem::exists(capture); ++attempt) Sleep(25);
    Check(std::filesystem::exists(capture), "Helper received request");
    std::ifstream input;
    for (int attempt = 0; attempt < 200; ++attempt) {
        input.open(capture, std::ios::binary);
        if (input.is_open()) break;
        input.clear();
        Sleep(25);
    }
    Check(input.is_open(), "Completed helper capture is readable");
    std::string actual(std::istreambuf_iterator<char>(input), {});
    Check(!input.bad(), "Read complete helper capture");
    const auto expected = filelist::BuildRequest(mode, paths);
    if (actual != expected) {
        std::cerr << "Dispatch mismatch: mode=" << mode << " count=" << paths.size()
                  << " actual-bytes=" << actual.size() << " expected-bytes=" << expected.size()
                  << "\nactual=" << actual.substr(0, 500) << "\nexpected=" << expected.substr(0, 500) << "\n";
    }
    Check(actual == expected, "Helper receives exact mode and selection");
    input.close();
    SetEnvironmentVariableW(L"FILELISTTOEXCEL_TEST_CAPTURE", nullptr);
}
}

int wmain(int argc, wchar_t** argv) {
    HRESULT apartment = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(apartment)) return 2;
    HMODULE library = nullptr;
    IClassFactory* factory = nullptr;
    try {
        std::filesystem::path dll = argc > 1 ? argv[1] : L"FileListToExcel.Shell.dll";
        bool dispatch = argc > 2 && std::wstring(argv[2]) == L"--dispatch";
        library = LoadLibraryW(dll.c_str());
        Check(library != nullptr, "Load shell DLL");
        using GetFactory = HRESULT (WINAPI*)(REFCLSID, REFIID, void**);
        using CanUnload = HRESULT (WINAPI*)();
        // memcpy avoids compiler warnings about casting incompatible function types.
        auto factoryProc = GetProcAddress(library, "DllGetClassObject");
        auto unloadProc = GetProcAddress(library, "DllCanUnloadNow");
        GetFactory getFactory = nullptr;
        CanUnload canUnload = nullptr;
        static_assert(sizeof(factoryProc) == sizeof(getFactory));
        memcpy(&getFactory, &factoryProc, sizeof(getFactory));
        memcpy(&canUnload, &unloadProc, sizeof(canUnload));
        Check(getFactory && canUnload, "COM exports exist");
        Check(canUnload() == S_OK, "Initially unloadable");
        Hr(getFactory(filelist::ShellClsid, IID_IClassFactory, reinterpret_cast<void**>(&factory)), "Load class factory");
        Check(canUnload() == S_FALSE, "Factory holds module alive");
        void* unknown = nullptr;
        Check(getFactory(GUID_NULL, IID_IClassFactory, &unknown) == CLASS_E_CLASSNOTAVAILABLE, "Reject unknown CLSID");
        Check(factory->CreateInstance(reinterpret_cast<IUnknown*>(factory), IID_IShellExtInit, &unknown) == CLASS_E_NOAGGREGATION,
            "Reject aggregation");
        Fixture fixture;
        const auto first = fixture.File(L"한글 보고서.txt");
        const auto second = fixture.File(L"emoji-\U0001F600 no extension");
        fixture.File(L"unselected.txt");
        const auto folder = fixture.Folder(L"folder with spaces");
        const auto folder2 = fixture.Folder(L"another folder");
        TestCodec();
        {
            Extension extension(factory);
            Check(extension.init->Initialize(nullptr, nullptr, nullptr) == E_INVALIDARG, "Reject missing input");
            extension.Select({first});
            extension.CheckMenu(2, L"파일 목록을 Excel로", L"같은 파일 찾기", L"filelisttoexcelmatches");
            if (dispatch) {
                TestDispatch(extension, fixture, {first}, "matches", 1);
                TestDispatch(extension, fixture, {first}, "matches", 1, true);
                TestDispatch(extension, fixture, {first}, "matches", 1, true, true);
            }
            extension.Select({first, second});
            extension.CheckMenu(2, L"파일 목록을 Excel로", L"선택 파일 중 중복 찾기", L"filelisttoexcelduplicatefiles");
            char verb[64]{};
            Hr(extension.menu->GetCommandString(0, GCS_VERBA, nullptr, verb, 64), "ANSI verb");
            Check(std::string(verb) == "filelisttoexcel", "ANSI verb matches");
            Check(extension.menu->GetCommandString(5, GCS_VERBA, nullptr, verb, 64) == E_INVALIDARG, "Reject unknown verb");
            CMINVOKECOMMANDINFO bad{};
            bad.cbSize = sizeof(bad);
            bad.lpVerb = MAKEINTRESOURCEA(30);
            Check(extension.menu->InvokeCommand(&bad) == E_INVALIDARG, "Reject unknown command");
            if (dispatch) {
                TestDispatch(extension, fixture, {first, second}, "files", 0);
                TestDispatch(extension, fixture, {first, second}, "duplicate-files", 1);
                TestDispatch(extension, fixture, {first, second}, "duplicate-files", 1, true);
                TestDispatch(extension, fixture, {first, second}, "duplicate-files", 1, true, true);
            }
            bad.lpVerb = "filelisttoexcelmatches";
            Check(extension.menu->InvokeCommand(&bad) == E_INVALIDARG, "Multiple files reject matches canonical verb");
            bad.lpVerb = "filelisttoexcelduplicates";
            Check(extension.menu->InvokeCommand(&bad) == E_INVALIDARG, "Files reject folder duplicates canonical verb");
            const auto archive = fixture.File(L"archive.zip");
            extension.SelectShellData({archive});
            extension.CheckMenu(2, L"파일 목록을 Excel로", L"같은 파일 찾기", L"filelisttoexcelmatches");
            extension.SelectShellData({first, archive});
            extension.CheckMenu(2, L"파일 목록을 Excel로", L"선택 파일 중 중복 찾기", L"filelisttoexcelduplicatefiles");
            extension.SelectShellData({folder, archive});
            extension.CheckMenu(2, L"선택 항목을 Excel로");
            extension.Select({folder});
            extension.CheckMenu(3, L"폴더 내용 목록을 Excel로", L"중복 파일 찾기", L"filelisttoexcelduplicates");
            if (dispatch) {
                TestDispatch(extension, fixture, {folder}, "duplicates", 2);
                TestDispatch(extension, fixture, {folder}, "duplicates", 2, true);
                TestDispatch(extension, fixture, {folder}, "duplicates", 2, true, true);
            }
            extension.Select({folder, folder2});
            extension.CheckMenu(3, L"폴더 내용 목록을 Excel로", L"중복 파일 찾기", L"filelisttoexcelduplicates");
            if (dispatch) {
                TestDispatch(extension, fixture, {folder, folder2}, "folder", 0);
                TestDispatch(extension, fixture, {folder, folder2}, "recursive", 1);
                TestDispatch(extension, fixture, {folder, folder2}, "duplicates", 2);
            }
            extension.Select({first, folder});
            extension.CheckMenu(2, L"선택 항목을 Excel로");
            for (const char* duplicateVerb : {"filelisttoexcelmatches", "filelisttoexcelduplicates", "filelisttoexcelduplicatefiles"}) {
                bad.lpVerb = duplicateVerb;
                Check(extension.menu->InvokeCommand(&bad) == E_INVALIDARG, "Mixed selection rejects duplicate canonical commands");
            }
            if (dispatch) {
                TestDispatch(extension, fixture, {first, folder}, "files", 0);
                TestDispatch(extension, fixture, {first, folder}, "recursive", 1);
            }
            if (dispatch) {
                std::vector<std::wstring> many;
                for (int index = 0; index < 10000; ++index)
                    many.push_back((fixture.directory / (L"large-selection-" + std::to_wstring(index) + L".txt")).wstring());
                extension.Select(many);
                extension.CheckMenu(2, L"파일 목록을 Excel로", L"선택 파일 중 중복 찾기", L"filelisttoexcelduplicatefiles");
                TestDispatch(extension, fixture, many, "files", 0, true);
                TestDispatch(extension, fixture, many, "duplicate-files", 1, true);
            }
            PIDLIST_ABSOLUTE pidl = nullptr;
            Hr(SHParseDisplayName(folder.c_str(), nullptr, &pidl, 0, nullptr), "Create background PIDL");
            HRESULT initialized = extension.init->Initialize(pidl, nullptr, nullptr);
            CoTaskMemFree(pidl);
            Hr(initialized, "Initialize folder background");
            extension.CheckMenu(3, L"현재 폴더 목록을 Excel로", L"현재 폴더 중복 파일 찾기", L"filelisttoexcelduplicates");
            if (dispatch) {
                TestDispatch(extension, fixture, {folder}, "folder", 0);
                TestDispatch(extension, fixture, {folder}, "duplicates", 2);
            }
            HMENU popup = CreatePopupMenu();
            Check(HRESULT_CODE(extension.menu->QueryContextMenu(popup, 0, 1, 2, CMF_DEFAULTONLY)) == 0, "Honor default-only menus");
            Check(GetMenuItemCount(popup) == 0, "Default-only inserts no menu");
            Check(HRESULT_CODE(extension.menu->QueryContextMenu(popup, 0, 1, 1, CMF_NORMAL)) == 1, "Honor command identifier capacity");
            Check(GetMenuItemCount(popup) == 1, "One available identifier creates one menu");
            DestroyMenu(popup);
            popup = CreatePopupMenu();
            Check(HRESULT_CODE(extension.menu->QueryContextMenu(popup, 0, 1, 2, CMF_NORMAL)) == 2, "Two identifiers retain only list commands");
            wchar_t limitedVerb[64]{};
            Hr(extension.menu->GetCommandString(1, GCS_VERBW, nullptr, reinterpret_cast<LPSTR>(limitedVerb), 64), "Limited recursive verb");
            Check(std::wstring(limitedVerb) == L"filelisttoexcelrecursive", "Duplicate command does not displace recursive command");
            bad.lpVerb = MAKEINTRESOURCEA(2);
            Check(extension.menu->InvokeCommand(&bad) == E_INVALIDARG, "Unavailable numeric duplicate identifier rejected");
            DestroyMenu(popup);
        }
        factory->LockServer(TRUE);
        factory->Release();
        factory = nullptr;
        Check(canUnload() == S_FALSE, "Server lock prevents unload");
        Hr(getFactory(filelist::ShellClsid, IID_IClassFactory, reinterpret_cast<void**>(&factory)), "Reload class factory");
        factory->LockServer(FALSE);
        factory->Release();
        factory = nullptr;
        Check(canUnload() == S_OK, "No lingering COM objects");
        FreeLibrary(library);
        library = nullptr;
        CoUninitialize();
        std::cout << "PASS: " << checks << " shell integration checks" << (dispatch ? " including helper dispatch" : "") << "\n";
        return 0;
    } catch (const std::exception& error) {
        if (factory) factory->Release();
        if (library) FreeLibrary(library);
        CoUninitialize();
        std::cerr << "FAIL after " << checks << " checks: " << error.what() << "\n";
        return 1;
    }
}
