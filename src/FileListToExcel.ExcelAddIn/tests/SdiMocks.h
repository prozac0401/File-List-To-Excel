#pragma once
#include "../Automation.h"
#include <functional>
#include <map>
#include <memory>
#include <olectl.h>

namespace flt::test
{
class Point final : public IConnectionPoint
{
    ULONG refs_ = 1;
    IID eventId_;
    Ptr<IDispatch> sink_;

  public:
    unsigned advises = 0, unadvises = 0;
    explicit Point(IID id) : eventId_(id)
    {
    }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID id, void **out) override
    {
        if (!out)
            return E_POINTER;
        *out = nullptr;
        if (id != IID_IUnknown && id != IID_IConnectionPoint)
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
    HRESULT STDMETHODCALLTYPE GetConnectionInterface(IID *id) override
    {
        if (!id)
            return E_POINTER;
        *id = eventId_;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetConnectionPointContainer(IConnectionPointContainer **) override
    {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE Advise(IUnknown *value, DWORD *cookie) override
    {
        if (sink_)
            return CONNECT_E_ADVISELIMIT;
        Check(value->QueryInterface(IID_IDispatch, reinterpret_cast<void **>(sink_.put())));
        ++advises;
        *cookie = 1;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Unadvise(DWORD cookie) override
    {
        if (cookie != 1 || !sink_)
            return CONNECT_E_NOCONNECTION;
        ++unadvises;
        sink_ = Ptr<IDispatch>();
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE EnumConnections(IEnumConnections **) override
    {
        return E_NOTIMPL;
    }
    void Fire(DISPID id, DISPPARAMS *args = nullptr)
    {
        DISPPARAMS empty{};
        if (sink_)
            Check(sink_->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, args ? args : &empty,
                                nullptr, nullptr, nullptr));
    }
};
class Node final : public IDispatch, public IConnectionPointContainer
{
    ULONG refs_ = 1;
    std::map<std::wstring, DISPID> ids_;
    std::map<DISPID, std::wstring> names_;

  public:
    std::map<std::wstring, V> properties;
    std::function<V(const std::wstring &, WORD, DISPPARAMS *)> action;
    Ptr<Point> point;
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID id, void **out) override
    {
        if (!out)
            return E_POINTER;
        *out = nullptr;
        if (id == IID_IUnknown || id == IID_IDispatch)
            *out = static_cast<IDispatch *>(this);
        else if (id == IID_IConnectionPointContainer && point)
            *out = static_cast<IConnectionPointContainer *>(this);
        else
            return E_NOINTERFACE;
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
            auto found = ids_.find(names[i]);
            if (found == ids_.end())
            {
                DISPID next = static_cast<DISPID>(ids_.size() + 1);
                ids_[names[i]] = next;
                names_[next] = names[i];
                ids[i] = next;
            }
            else
                ids[i] = found->second;
        }
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID id, REFIID, LCID, WORD flags, DISPPARAMS *args, VARIANT *result,
                                     EXCEPINFO *, UINT *) override
    {
        try
        {
            auto name = names_.at(id);
            V value;
            if (flags == DISPATCH_PROPERTYPUT)
            {
                V copy;
                Check(VariantCopy(&copy, &args->rgvarg[0]));
                properties[name] = std::move(copy);
            }
            else if (action)
                value = action(name, flags, args);
            else
                value = properties.at(name);
            if (result)
                Check(VariantCopy(result, &value));
            return S_OK;
        }
        catch (...)
        {
            return DISP_E_MEMBERNOTFOUND;
        }
    }
    HRESULT STDMETHODCALLTYPE EnumConnectionPoints(IEnumConnectionPoints **) override
    {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE FindConnectionPoint(REFIID id, IConnectionPoint **out) override
    {
        if (!point)
            return CONNECT_E_NOCONNECTION;
        IID expected{};
        point->GetConnectionInterface(&expected);
        if (id != expected)
            return CONNECT_E_NOCONNECTION;
        *out = point.get();
        point->AddRef();
        return S_OK;
    }
};
struct Controls
{
    Ptr<Node> dispatch{new Node()};
    std::vector<Ptr<Node>> items;
    std::vector<std::unique_ptr<Controls>> children;
    Controls()
    {
        dispatch->action = [this](const std::wstring &name, WORD, DISPPARAMS *args) -> V {
            if (name == L"Count")
                return V(static_cast<long>(items.size()));
            if (name == L"Item")
                return V(items.at(static_cast<size_t>(args->rgvarg[0].lVal - 1)).get());
            if (name == L"Add")
                return V(Add(args->rgvarg[args->cArgs - 1].lVal).get());
            throw std::runtime_error("Unexpected controls call");
        };
    }
    Ptr<Node> Add(long type)
    {
        Ptr<Node> node(new Node());
        node->properties[L"Tag"] = V(L"");
        node->properties[L"Caption"] = V(L"");
        node->properties[L"Enabled"] = V(true);
        if (type == 10)
        {
            children.push_back(std::make_unique<Controls>());
            node->properties[L"Controls"] = V(children.back()->dispatch.get());
        }
        else
            node->point = Ptr<Point>(new Point(ButtonEventsIid));
        Node *identity = node.get();
        node->action = [this, identity](const std::wstring &name, WORD, DISPPARAMS *) -> V {
            if (name == L"Delete")
            {
                items.erase(
                    std::remove_if(items.begin(), items.end(),
                                   [identity](const Ptr<Node> &item) { return item.get() == identity; }),
                    items.end());
                return V();
            }
            return identity->properties.at(name);
        };
        items.push_back(node);
        return node;
    }
    Node *Button()
    {
        return children.empty() || children.front()->items.empty() ? nullptr
                                                                   : children.front()->items.front().get();
    }
};
struct Window
{
    long id;
    Ptr<Node> object{new Node()}, bars{new Node()};
    Controls cell, list;
    Ptr<Node> cellBar{new Node()}, listBar{new Node()};
    explicit Window(long hwnd) : id(hwnd)
    {
        object->properties[L"Hwnd"] = V(hwnd);
        cellBar->properties[L"Controls"] = V(cell.dispatch.get());
        listBar->properties[L"Controls"] = V(list.dispatch.get());
        bars->action = [this](const std::wstring &name, WORD, DISPPARAMS *args) -> V {
            if (name != L"Item")
                throw std::runtime_error("Unexpected commandbars call");
            return V(std::wstring(args->rgvarg[0].bstrVal) == L"Cell" ? cellBar.get() : listBar.get());
        };
    }
    void CloneVisibleMenus()
    {
        for (Controls *controls : {&cell, &list})
        {
            auto menu = controls->Add(10);
            menu->properties[L"Tag"] = V(MenuTag);
            auto button = controls->children.back()->Add(1);
            button->properties[L"Tag"] = V(ButtonTag);
            button->properties[L"Enabled"] = V(true);
        }
    }
};
inline void SdiRegression(IDTExtensibility2 *addin)
{
    Window first(101), second(202);
    Ptr<Node> app(new Node()), windows(new Node());
    app->point = Ptr<Point>(new Point(ApplicationEventsIid));
    Window *active = &first;
    bool firstOpen = true;
    bool hasWindow = false;
    bool disconnectDuringCommandBars = false;
    windows->action = [&](const std::wstring &name, WORD, DISPPARAMS *args) -> V {
        if (name == L"Count")
            return V(!hasWindow ? 0L : (firstOpen ? 2L : 1L));
        if (name == L"Item")
            return V(firstOpen && args->rgvarg[0].lVal == 1 ? first.object.get() : second.object.get());
        throw std::runtime_error("Unexpected window call");
    };
    app->action = [&](const std::wstring &name, WORD, DISPPARAMS *) -> V {
        if (name == L"CommandBars")
        {
            if (!hasWindow)
                throw std::runtime_error("Start screen has no CommandBars");
            if (disconnectDuringCommandBars)
            {
                disconnectDuringCommandBars = false;
                Check(addin->OnDisconnection(0, nullptr));
            }
            return V(active->bars.get());
        }
        if (name == L"ActiveWindow")
        {
            if (!hasWindow)
                throw std::runtime_error("Start screen has no window");
            return V(active->object.get());
        }
        if (name == L"Windows")
            return V(windows.get());
        throw std::runtime_error("An ordinary workbook has no product metadata");
    };
    struct DisconnectGuard
    {
        IDTExtensibility2 *value;
        ~DisconnectGuard()
        {
            value->OnDisconnection(0, nullptr);
        }
    } guard{addin};
    Check(addin->OnConnection(app.get(), 0, nullptr, nullptr));
    auto require = [](bool value, const char *message) {
        if (!value)
            throw std::runtime_error(message);
    };
    require(!first.cell.Button() && app->point->advises == 1,
            "Start-screen connection must retain events without requiring an active workbook/window");
    hasWindow = true;
    app->point->Fire(0x61f);
    require(first.cell.Button() && !first.cell.Button()->properties.at(L"Enabled").Truth(),
            "Initial ordinary workbook must be disabled");
    require(first.cell.Button()->point->advises == 1, "Initial button needs one compiled callback");
    second.CloneVisibleMenus();
    active = &second;
    app->point->Fire(0x620);
    require(!second.cell.Button()->properties.at(L"Enabled").Truth(),
            "WorkbookActivate must update the new SDI control identity");
    require(second.cell.Button()->point->advises == 1, "A cloned SDI button needs its own compiled callback");
    app->point->Fire(0x614);
    app->point->Fire(0x616);
    require(second.cell.Button()->point->advises == 1,
            "Repeated window/selection events must not duplicate callbacks");
    second.cell.Button()->properties[L"Enabled"] = V(true);
    app->point->Fire(0x61c);
    require(!second.cell.Button()->properties.at(L"Enabled").Truth(),
            "SheetChange must refresh invalid metadata state");
    second.cell.Button()->properties[L"Enabled"] = V(true);
    IDispatch *target = second.object.get();
    VARIANT args[3]{};
    args[1].vt = VT_DISPATCH | VT_BYREF;
    args[1].ppdispVal = &target;
    DISPPARAMS event{args, nullptr, 3, 0};
    app->point->Fire(0x618, &event);
    require(!second.cell.Button()->properties.at(L"Enabled").Truth(),
            "By-reference right-click Target must still refresh");
    Ptr<Point> oldPoint(first.cell.Button()->point);
    firstOpen = false;
    app->point->Fire(0x620);
    require(oldPoint->unadvises == 1, "Closed workbook callback must be released");
    disconnectDuringCommandBars = true;
    app->point->Fire(0x620);
    require(second.cell.items.empty() && second.list.items.empty(),
            "Disconnect must remove only owned live menus");
    require(app->point->unadvises == 1, "Disconnect must release application event subscription");
}
} // namespace flt::test
