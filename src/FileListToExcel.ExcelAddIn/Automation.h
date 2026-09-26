#pragma once
#include <algorithm>
#include <atomic>
#include <oaidl.h>
#include <ocidl.h>
#include <stdexcept>
#include <string>
#include <vector>
#include <windows.h>

namespace flt
{
inline constexpr CLSID AddInClsid = {
    0x0b1e297c, 0x42cc, 0x48a4, {0xa9, 0x73, 0x3a, 0xa8, 0xea, 0xf2, 0x67, 0x95}};
inline constexpr IID ExtensibilityIid = {
    0xb65ad801, 0xabaf, 0x11d0, {0xbb, 0x8b, 0x00, 0xa0, 0xc9, 0x0f, 0x27, 0x44}};
inline constexpr IID ButtonEventsIid = {
    0x000c0351, 0x0000, 0x0000, {0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}};
inline constexpr IID ApplicationEventsIid = {
    0x00024413, 0x0000, 0x0000, {0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}};
inline constexpr wchar_t ProgId[] = L"FileListToExcel.ExcelAddIn";
inline constexpr wchar_t MenuTag[] = L"FileListToExcel.ExcelAddIn.Menu.v1";
inline constexpr wchar_t ButtonTag[] = L"FileListToExcel.ExcelAddIn.Collect.v1";
struct IDTExtensibility2 : IDispatch
{
    virtual HRESULT STDMETHODCALLTYPE OnConnection(IDispatch *, LONG, IDispatch *, SAFEARRAY **) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnDisconnection(LONG, SAFEARRAY **) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY **) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY **) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY **) = 0;
};
inline void Check(HRESULT hr)
{
    if (FAILED(hr))
        throw std::runtime_error("Excel automation call failed");
}
template <class T> class Ptr
{
    T *p_ = nullptr;

  public:
    Ptr() = default;
    explicit Ptr(T *p, bool retain = false) : p_(p)
    {
        if (p_ && retain)
            p_->AddRef();
    }
    Ptr(const Ptr &p) : Ptr(p.p_, true)
    {
    }
    Ptr(Ptr &&p) noexcept : p_(p.p_)
    {
        p.p_ = nullptr;
    }
    ~Ptr()
    {
        if (p_)
            p_->Release();
    }
    Ptr &operator=(Ptr p) noexcept
    {
        std::swap(p_, p.p_);
        return *this;
    }
    T *get() const
    {
        return p_;
    }
    T **put()
    {
        if (p_)
            p_->Release();
        p_ = nullptr;
        return &p_;
    }
    T *operator->() const
    {
        return p_;
    }
    explicit operator bool() const
    {
        return p_ != nullptr;
    }
};
class V : public VARIANT
{
  public:
    V()
    {
        VariantInit(this);
    }
    explicit V(long n) : V()
    {
        vt = VT_I4;
        lVal = n;
    }
    explicit V(bool b) : V()
    {
        vt = VT_BOOL;
        boolVal = b ? VARIANT_TRUE : VARIANT_FALSE;
    }
    explicit V(const wchar_t *s) : V()
    {
        vt = VT_BSTR;
        bstrVal = SysAllocString(s);
        if (!bstrVal)
            throw std::bad_alloc();
    }
    explicit V(const std::wstring &s) : V(s.c_str())
    {
    }
    explicit V(IDispatch *p) : V()
    {
        vt = VT_DISPATCH;
        pdispVal = p;
        if (p)
            p->AddRef();
    }
    V(const V &v) : V()
    {
        Check(VariantCopy(this, const_cast<V *>(&v)));
    }
    V(V &&v) noexcept
    {
        static_cast<VARIANT &>(*this) = v;
        VariantInit(&v);
    }
    V &operator=(V v) noexcept
    {
        std::swap(static_cast<VARIANT &>(*this), static_cast<VARIANT &>(v));
        return *this;
    }
    ~V()
    {
        VariantClear(this);
    }
    static V Missing()
    {
        V v;
        v.vt = VT_ERROR;
        v.scode = DISP_E_PARAMNOTFOUND;
        return v;
    }
    Ptr<IDispatch> Obj() const
    {
        if (vt != VT_DISPATCH || !pdispVal)
            throw std::runtime_error("Missing Excel object");
        return Ptr<IDispatch>(pdispVal, true);
    }
    std::wstring Text() const
    {
        if (vt != VT_BSTR)
            throw std::runtime_error("Expected text");
        return std::wstring(bstrVal, SysStringLen(bstrVal));
    }
    long Number() const
    {
        V n;
        Check(VariantChangeType(&n, const_cast<V *>(this), 0, VT_I4));
        return n.lVal;
    }
    bool Truth() const
    {
        V n;
        Check(VariantChangeType(&n, const_cast<V *>(this), 0, VT_BOOL));
        return n.boolVal == VARIANT_TRUE;
    }
};
inline V Call(IDispatch *p, const wchar_t *name, WORD flags = DISPATCH_PROPERTYGET, std::vector<V> args = {})
{
    if (!p)
        throw std::runtime_error("Disconnected Excel");
    DISPID id = 0;
    LPOLESTR n = const_cast<LPOLESTR>(name);
    Check(p->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &id));
    std::reverse(args.begin(), args.end());
    std::vector<VARIANT> raw;
    for (auto &a : args)
        raw.push_back(a);
    DISPID put = DISPID_PROPERTYPUT;
    DISPPARAMS dp{raw.data(), flags == DISPATCH_PROPERTYPUT ? &put : nullptr, static_cast<UINT>(raw.size()),
                  flags == DISPATCH_PROPERTYPUT ? 1u : 0u};
    V result;
    EXCEPINFO ex{};
    auto hr = p->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags, &dp, &result, &ex, nullptr);
    SysFreeString(ex.bstrSource);
    SysFreeString(ex.bstrDescription);
    SysFreeString(ex.bstrHelpFile);
    Check(hr);
    return result;
}
inline Ptr<IDispatch> Obj(IDispatch *p, const wchar_t *n, std::vector<V> a = {})
{
    return Call(p, n, DISPATCH_PROPERTYGET, std::move(a)).Obj();
}
inline void Put(IDispatch *p, const wchar_t *n, V value)
{
    Call(p, n, DISPATCH_PROPERTYPUT, {std::move(value)});
}
inline std::wstring Property(IDispatch *workbook, const std::wstring &name)
{
    auto props = Obj(workbook, L"CustomDocumentProperties");
    auto prop = Obj(props.get(), L"Item", {V(name)});
    return Call(prop.get(), L"Value").Text();
}
inline bool IsGuid(const std::wstring &text)
{
    if (text.size() != 36)
        return false;
    GUID id{};
    return SUCCEEDED(CLSIDFromString((L"{" + text + L"}").c_str(), &id)) && id != GUID_NULL;
}
} // namespace flt
