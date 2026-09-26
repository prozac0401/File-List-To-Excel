#include "../Automation.h"
#include "../Selection.h"
#include "SdiMocks.h"
#include "SelectionMocks.h"
#include <iostream>
using namespace flt;
int wmain(int argc, wchar_t **argv)
{
    if (argc < 2)
        return 2;
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    int result = 1;
    try
    {
        HMODULE dll = LoadLibraryW(argv[1]);
        if (!dll)
            throw std::runtime_error("DLL load failed");
        using GetClass = HRESULT(__stdcall *)(REFCLSID, REFIID, void **);
        using Unload = HRESULT(__stdcall *)();
        auto get =
            reinterpret_cast<GetClass>(reinterpret_cast<void *>(GetProcAddress(dll, "DllGetClassObject")));
        auto unload =
            reinterpret_cast<Unload>(reinterpret_cast<void *>(GetProcAddress(dll, "DllCanUnloadNow")));
        if (!get || !unload || unload() != S_OK)
            throw std::runtime_error("Exports/lifetime invalid");
        {
            Ptr<IClassFactory> factory;
            Check(get(AddInClsid, IID_IClassFactory, reinterpret_cast<void **>(factory.put())));
            Ptr<IDTExtensibility2> addin;
            Check(factory->CreateInstance(nullptr, ExtensibilityIid, reinterpret_cast<void **>(addin.put())));
            if (unload() != S_FALSE)
                throw std::runtime_error("DLL prematurely unloadable");
            flt::test::SdiRegression(addin.get());
            flt::test::SelectionRegression();
            Check(addin->OnAddInsUpdate(nullptr));
            Check(addin->OnDisconnection(0, nullptr));
            Check(addin->OnBeginShutdown(nullptr));
            if (SUCCEEDED(addin->OnConnection(nullptr, 0, nullptr, nullptr)))
                throw std::runtime_error("Null Excel accepted");
        }
        if (unload() != S_OK)
            throw std::runtime_error("COM lifetime leaked");
        if (!IsGuid(NewId()) || IsGuid(L"bad"))
            throw std::runtime_error("GUID validation failed");
        FreeLibrary(dll);
        std::cout << "Native COM lifecycle, request identity, SDI events and selection capture tests passed\n";
        result = 0;
    }
    catch (const std::exception &e)
    {
        std::cerr << e.what() << "\n";
    }
    CoUninitialize();
    return result;
}
