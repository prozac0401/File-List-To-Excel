#include "Automation.h"
#include "RequestDispatch.h"
#include "Selection.h"
#include <functional>
#include <memory>
#include <set>
using namespace flt;
namespace
{
HMODULE moduleHandle = nullptr;
std::atomic<long> objects{0}, locks{0};
class Sink final : public IDispatch
{
    std::atomic<ULONG> refs_{1};
    IID iid_;
    std::function<void(DISPID, DISPPARAMS *)> action_;

  public:
    Sink(IID iid, std::function<void(DISPID, DISPPARAMS *)> action) : iid_(iid), action_(std::move(action))
    {
    }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void **out) override
    {
        if (!out)
            return E_POINTER;
        *out = nullptr;
        if (iid != IID_IUnknown && iid != IID_IDispatch && iid != iid_)
            return E_NOINTERFACE;
        *out = static_cast<IDispatch *>(this);
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return ++refs_;
    }
    ULONG STDMETHODCALLTYPE Release() override
    {
        auto n = --refs_;
        if (!n)
            delete this;
        return n;
    }
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *n) override
    {
        if (n)
            *n = 0;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo **) override
    {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR *, UINT, LCID, DISPID *) override
    {
        return DISP_E_UNKNOWNNAME;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID id, REFIID, LCID, WORD, DISPPARAMS *args, VARIANT *, EXCEPINFO *,
                                     UINT *) override
    {
        AddRef(); // A callback may unadvise this sink while Invoke is still on the stack.
        try
        {
            action_(id, args);
        }
        catch (...)
        {
        }
        Release();
        return S_OK;
    }
};
struct Subscription
{
    Ptr<IConnectionPoint> point;
    DWORD cookie = 0;
    Subscription(IDispatch *source, IID iid, std::function<void(DISPID, DISPPARAMS *)> action)
    {
        Ptr<IConnectionPointContainer> container;
        Check(source->QueryInterface(IID_IConnectionPointContainer,
                                     reinterpret_cast<void **>(container.put())));
        Check(container->FindConnectionPoint(iid, point.put()));
        Ptr<Sink> sink(new Sink(iid, std::move(action)));
        Check(point->Advise(sink.get(), &cookie));
    }
    ~Subscription()
    {
        if (point && cookie)
            point->Unadvise(cookie);
    }
};
class AddIn final : public IDTExtensibility2
{
    std::atomic<ULONG> refs_{1};
    Ptr<IDispatch> app_;
    struct OwnedMenu
    {
        Ptr<IDispatch> control;
        Ptr<IUnknown> identity;
        std::set<long> windows;
    };
    struct BoundButton
    {
        Ptr<IDispatch> control;
        Ptr<IUnknown> identity;
        std::unique_ptr<Subscription> subscription;
        std::set<long> windows;
    };
    std::vector<OwnedMenu> menus_;
    std::vector<BoundButton> buttons_;
    std::unique_ptr<Subscription> applicationEvents_;
    bool refreshing_ = false;
    unsigned operationDepth_ = 0;
    bool disconnectPending_ = false;
    bool disconnecting_ = false;
    struct Operation
    {
        AddIn *owner;
        explicit Operation(AddIn *value) : owner(value)
        {
            owner->AddRef();
            ++owner->operationDepth_;
        }
        ~Operation()
        {
            if (--owner->operationDepth_ == 0 && owner->disconnectPending_)
                owner->DisconnectNow();
            owner->Release();
        }
    };
    void RequestDisconnect()
    {
        if (disconnecting_)
            return;
        if (operationDepth_)
            disconnectPending_ = true;
        else
            DisconnectNow();
    }
    HANDLE helper_ = nullptr;
    bool invoking_ = false;
    HWND Owner()
    {
        try
        {
            return reinterpret_cast<HWND>(static_cast<INT_PTR>(Call(app_.get(), L"Hwnd").Number()));
        }
        catch (...)
        {
            return nullptr;
        }
    }
    bool Busy()
    {
        if (helper_ && WaitForSingleObject(helper_, 0) == WAIT_OBJECT_0)
        {
            CloseHandle(helper_);
            helper_ = nullptr;
        }
        return invoking_ || helper_;
    }
    static Ptr<IUnknown> Identity(IDispatch *control)
    {
        Ptr<IUnknown> identity;
        Check(control->QueryInterface(IID_IUnknown, reinterpret_cast<void **>(identity.put())));
        return identity;
    }
    long CurrentWindow()
    {
        try
        {
            auto window = Obj(app_.get(), L"ActiveWindow");
            return Call(window.get(), L"Hwnd").Number();
        }
        catch (...)
        {
            return 0;
        } // Excel's startup view has no workbook window yet.
    }
    void PruneClosedWindows()
    {
        std::set<long> live;
        try
        {
            auto windows = Obj(app_.get(), L"Windows");
            long count = Call(windows.get(), L"Count").Number();
            if (count > 1024)
                return;
            for (long index = 1; index <= count; ++index)
            {
                auto window = Obj(windows.get(), L"Item", {V(index)});
                live.insert(Call(window.get(), L"Hwnd").Number());
            }
            if (live.empty())
                live.insert(0);
        }
        catch (...)
        {
            return;
        } // A transient COM rejection is not proof that a window closed.
        auto prune = [&live](auto &controls) {
            for (auto &control : controls)
                for (auto it = control.windows.begin(); it != control.windows.end();)
                    if (!live.count(*it))
                        it = control.windows.erase(it);
                    else
                        ++it;
            controls.erase(std::remove_if(controls.begin(), controls.end(),
                                          [](const auto &control) { return control.windows.empty(); }),
                           controls.end());
        };
        prune(buttons_); // Unadvise before releasing controls belonging to closed workbook windows.
        prune(menus_);
    }
    Ptr<IDispatch> OwnedControl(IDispatch *controls, const wchar_t *tag)
    {
        Ptr<IDispatch> found;
        for (long index = Call(controls, L"Count").Number(); index >= 1; --index)
        {
            auto control = Obj(controls, L"Item", {V(index)});
            if (Call(control.get(), L"Tag").Text() != tag)
                continue;
            if (!found)
                found = control;
            else
                Call(control.get(), L"Delete", DISPATCH_METHOD); // Exact product tag only.
        }
        return found;
    }
    void EnsureCurrentSurface(const wchar_t *barName, long windowId, bool enabled)
    {
        auto bars = Obj(app_.get(), L"CommandBars");
        auto bar = Obj(bars.get(), L"Item", {V(barName)});
        auto controls = Obj(bar.get(), L"Controls");
        auto menu = OwnedControl(controls.get(), MenuTag);
        if (!menu)
        {
            menu = Call(controls.get(), L"Add", DISPATCH_METHOD,
                        {V(10L), V::Missing(), V::Missing(), V::Missing(), V(true)})
                       .Obj();
            Put(menu.get(), L"Tag", V(MenuTag));
            Put(menu.get(), L"Caption", V(L"파일목록"));
        }
        auto menuIdentity = Identity(menu.get());
        auto knownMenu = std::find_if(menus_.begin(), menus_.end(), [&](const OwnedMenu &entry) {
            return entry.identity.get() == menuIdentity.get();
        });
        if (knownMenu == menus_.end())
            menus_.push_back({menu, menuIdentity, {windowId}});
        else
            knownMenu->windows.insert(windowId);

        auto children = Obj(menu.get(), L"Controls");
        auto button = OwnedControl(children.get(), ButtonTag);
        if (!button)
        {
            button = Call(children.get(), L"Add", DISPATCH_METHOD,
                          {V(1L), V::Missing(), V::Missing(), V::Missing(), V(true)})
                         .Obj();
            Put(button.get(), L"Tag", V(ButtonTag));
            Put(button.get(), L"Caption", V(L"선택한 파일 복사…"));
            Put(button.get(), L"TooltipText",
                V(L"선택한 보이는 목록 행의 실제 파일을 새 폴더에 복사합니다."));
        }
        // SDI can expose a different control COM identity in each workbook window.
        // Reacquire the active window's tagged control and bind each identity exactly once.
        Put(button.get(), L"Enabled", V(false));
        auto buttonIdentity = Identity(button.get());
        auto knownButton = std::find_if(buttons_.begin(), buttons_.end(), [&](const BoundButton &entry) {
            return entry.identity.get() == buttonIdentity.get();
        });
        if (knownButton == buttons_.end())
        {
            auto subscription = std::make_unique<Subscription>(button.get(), ButtonEventsIid,
                                                               [this](DISPID id, DISPPARAMS *) {
                                                                   Operation operation(this);
                                                                   if (!disconnectPending_ && !disconnecting_ && id == 1)
                                                                       Click();
                                                               });
            buttons_.push_back({button, buttonIdentity, std::move(subscription), {windowId}});
        }
        else
            knownButton->windows.insert(windowId);
        Put(button.get(), L"Enabled", V(enabled));
    }
    void Refresh(IDispatch *target = nullptr)
    {
        if (refreshing_ || disconnectPending_ || disconnecting_ || !app_)
            return;
        refreshing_ = true;
        struct Restore
        {
            bool &value;
            ~Restore()
            {
                value = false;
            }
        } restore{refreshing_};
        bool enabled = false;
        try
        {
            if (!Busy())
            {
                auto selected = Locate(app_.get());
                enabled =
                    !target || static_cast<bool>(Intersection(app_.get(), selected.selection.get(), target));
            }
        }
        catch (...)
        {
        }
        if (disconnectPending_)
            return;
        PruneClosedWindows();
        if (disconnectPending_)
            return;
        long windowId = CurrentWindow();
        for (const wchar_t *bar : {L"Cell", L"List Range Popup"})
            try
            {
                if (disconnectPending_)
                    break;
                EnsureCurrentSurface(bar, windowId, enabled);
            }
            catch (...)
            {
            }
    }
    void DisconnectNow()
    {
        if (disconnecting_)
            return;
        disconnecting_ = true;
        disconnectPending_ = false;
        applicationEvents_.reset();
        buttons_.clear();
        for (auto &menu : menus_)
            try
            {
                Call(menu.control.get(), L"Delete", DISPATCH_METHOD);
            }
            catch (...)
            {
            }
        menus_.clear();
        app_ = Ptr<IDispatch>();
        if (helper_)
        {
            CloseHandle(helper_);
            helper_ = nullptr;
        }
        disconnecting_ = false;
    }
    void Click()
    {
        if (disconnectPending_ || disconnecting_ || !app_ || Busy())
            return;
        invoking_ = true;
        std::wstring requestPath;
        try
        {
            auto id = NewId();
            auto payload = Capture(app_.get(), id);
            if (disconnectPending_)
                throw std::runtime_error("Excel disconnected during selection capture");
            std::vector<wchar_t> module(32768);
            auto length = GetModuleFileNameW(moduleHandle, module.data(), static_cast<DWORD>(module.size()));
            if (!length || length >= module.size())
                throw std::runtime_error("Installation path unavailable");
            std::wstring directory(module.data(), length);
            directory.resize(directory.find_last_of(L"\\/"));
            auto exe = directory + L"\\FileListToExcel.exe";
            Check(WriteRequest(payload, id, requestPath));
            std::wstring command =
                filelist::QuoteArgument(exe) + L" --collect-request " + filelist::QuoteArgument(requestPath);
            STARTUPINFOW si{};
            si.cb = sizeof(si);
            PROCESS_INFORMATION pi{};
            if (!CreateProcessW(exe.c_str(), command.data(), nullptr, nullptr, FALSE, CREATE_NO_WINDOW,
                                nullptr, directory.c_str(), &si, &pi))
                throw std::runtime_error("Helper could not start; repair installation");
            CloseHandle(pi.hThread);
            helper_ = pi.hProcess;
            requestPath.clear();
        }
        catch (const std::exception &e)
        { /* A failed dispatch may leave its random request for local diagnosis. Never delete by a replaceable
             pathname. */
            if (disconnectPending_)
            {
                invoking_ = false;
                return;
            }
            std::string error = e.what();
            int length = MultiByteToWideChar(CP_UTF8, 0, error.c_str(), -1, nullptr, 0);
            std::wstring message(static_cast<size_t>(length), L'\0');
            MultiByteToWideChar(CP_UTF8, 0, error.c_str(), -1, message.data(), length);
            MessageBoxW(Owner(),
                        (L"선택한 파일을 복사할 수 없습니다. 지원되는 새 결과표에서 보이는 데이터 행을 "
                         L"선택해 주세요.\n\n" +
                         message)
                            .c_str(),
                        L"File List to Excel", MB_OK | MB_ICONINFORMATION);
        }
        invoking_ = false;
        Refresh();
    }
    void InstallMenus()
    {
        applicationEvents_ = std::make_unique<Subscription>(
            app_.get(), ApplicationEventsIid, [this](DISPID id, DISPPARAMS *args) {
                Operation operation(this);
                if (disconnectPending_ || disconnecting_)
                    return;
                if (id == 0x618) // SheetBeforeRightClick: Target may arrive as a by-reference VARIANT.
                {
                    Ptr<IDispatch> target;
                    if (args && args->cArgs >= 3)
                    {
                        V value;
                        if (SUCCEEDED(VariantCopyInd(&value, &args->rgvarg[1])) && value.vt == VT_DISPATCH)
                            target = Ptr<IDispatch>(value.pdispVal, true);
                    }
                    Refresh(target.get());
                }
                else if (id == 0x616 || id == 0x617 || id == 0x619 || id == 0x61c || id == 0x61d ||
                         id == 0x61f || id == 0x620 || id == 0x614)
                    Refresh();
            });
        Refresh();
        // A normal start-screen launch has no workbook window/CommandBars yet.
        // Keep the application event subscription; the first workbook/window event prepares its menus.
        if (buttons_.empty() && CurrentWindow() != 0)
            throw std::runtime_error("Cannot initialize Excel context menu");
    }

  public:
    AddIn()
    {
        ++objects;
    }
    ~AddIn()
    {
        DisconnectNow();
        --objects;
    }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void **out) override
    {
        if (!out)
            return E_POINTER;
        *out = nullptr;
        if (iid != IID_IUnknown && iid != IID_IDispatch && iid != ExtensibilityIid)
            return E_NOINTERFACE;
        *out = static_cast<IDTExtensibility2 *>(this);
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return ++refs_;
    }
    ULONG STDMETHODCALLTYPE Release() override
    {
        auto n = --refs_;
        if (!n)
            delete this;
        return n;
    }
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *n) override
    {
        if (!n)
            return E_POINTER;
        *n = 0;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo **) override
    {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR *names, UINT count, LCID, DISPID *ids) override
    {
        for (UINT i = 0; i < count; ++i)
        {
            if (_wcsicmp(names[i], L"OnConnection") == 0)
                ids[i] = 1;
            else if (_wcsicmp(names[i], L"OnDisconnection") == 0)
                ids[i] = 2;
            else if (_wcsicmp(names[i], L"OnAddInsUpdate") == 0)
                ids[i] = 3;
            else if (_wcsicmp(names[i], L"OnStartupComplete") == 0)
                ids[i] = 4;
            else if (_wcsicmp(names[i], L"OnBeginShutdown") == 0)
                ids[i] = 5;
            else
                return DISP_E_UNKNOWNNAME;
        }
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID id, REFIID, LCID, WORD, DISPPARAMS *p, VARIANT *, EXCEPINFO *,
                                     UINT *) override
    {
        if (id == 1 && p && p->cArgs == 4 && p->rgvarg[3].vt == VT_DISPATCH)
            return OnConnection(p->rgvarg[3].pdispVal, 0, nullptr, nullptr);
        if (id == 2)
            return OnDisconnection(0, nullptr);
        if (id == 3)
            return OnAddInsUpdate(nullptr);
        if (id == 4)
            return OnStartupComplete(nullptr);
        if (id == 5)
            return OnBeginShutdown(nullptr);
        return DISP_E_MEMBERNOTFOUND;
    }
    HRESULT STDMETHODCALLTYPE OnConnection(IDispatch *app, LONG, IDispatch *, SAFEARRAY **) override
    {
        if (!app)
            return E_INVALIDARG;
        if (operationDepth_ || disconnecting_)
            return E_UNEXPECTED;
        Operation operation(this);
        DisconnectNow();
        try
        {
            app_ = Ptr<IDispatch>(app, true);
            InstallMenus();
            return disconnectPending_ ? E_ABORT : S_OK;
        }
        catch (...)
        {
            RequestDisconnect();
            return E_FAIL;
        }
    }
    HRESULT STDMETHODCALLTYPE OnDisconnection(LONG, SAFEARRAY **) override
    {
        Operation operation(this);
        RequestDisconnect();
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY **) override
    {
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY **) override
    {
        Operation operation(this);
        Refresh();
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY **) override
    {
        Operation operation(this);
        RequestDisconnect();
        return S_OK;
    }
};
class Factory final : public IClassFactory
{
    std::atomic<ULONG> refs_{1};

  public:
    Factory()
    {
        ++objects;
    }
    ~Factory()
    {
        --objects;
    }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void **out) override
    {
        if (!out)
            return E_POINTER;
        *out = nullptr;
        if (iid != IID_IUnknown && iid != IID_IClassFactory)
            return E_NOINTERFACE;
        *out = this;
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return ++refs_;
    }
    ULONG STDMETHODCALLTYPE Release() override
    {
        auto n = --refs_;
        if (!n)
            delete this;
        return n;
    }
    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown *outer, REFIID iid, void **out) override
    {
        if (outer)
            return CLASS_E_NOAGGREGATION;
        auto addin = new (std::nothrow) AddIn();
        if (!addin)
            return E_OUTOFMEMORY;
        auto hr = addin->QueryInterface(iid, out);
        addin->Release();
        return hr;
    }
    HRESULT STDMETHODCALLTYPE LockServer(BOOL lock) override
    {
        if (lock)
            ++locks;
        else
            --locks;
        return S_OK;
    }
};
} // namespace
extern "C" BOOL WINAPI DllMain(HINSTANCE module, DWORD reason, LPVOID)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        moduleHandle = module;
        DisableThreadLibraryCalls(module);
    }
    return TRUE;
}
extern "C" HRESULT __stdcall DllGetClassObject(REFCLSID clsid, REFIID iid, void **out)
{
    if (clsid != AddInClsid)
        return CLASS_E_CLASSNOTAVAILABLE;
    auto factory = new (std::nothrow) Factory();
    if (!factory)
        return E_OUTOFMEMORY;
    auto hr = factory->QueryInterface(iid, out);
    factory->Release();
    return hr;
}
extern "C" HRESULT __stdcall DllCanUnloadNow()
{
    return objects == 0 && locks == 0 ? S_OK : S_FALSE;
}
