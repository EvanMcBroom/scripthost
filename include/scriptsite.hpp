#pragma once
#include <ActivScp.h>
#include <string>
#include <vector>

class GenericScriptSite :
    public IActiveScriptSite,
    public IActiveScriptSiteWindow {
public:
    using NamedItem = std::pair<std::wstring, IDispatch*>;

    GenericScriptSite() : referenceCount(1) { };
    GenericScriptSite(const std::vector<NamedItem>& namedItems);
    ~GenericScriptSite();

    // IUnknown
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // IActiveGenericScriptSite
    STDMETHODIMP GetLCID(LCID __RPC_FAR* plcid);
    STDMETHODIMP GetItemInfo(LPCOLESTR pstrName, DWORD dwReturnMask, IUnknown __RPC_FAR* __RPC_FAR* ppiunkItem, ITypeInfo __RPC_FAR* __RPC_FAR* ppti);
    STDMETHODIMP GetDocVersionString(BSTR __RPC_FAR* pbstrVersion);
    STDMETHODIMP OnScriptTerminate(const VARIANT __RPC_FAR* pvarResult, const EXCEPINFO __RPC_FAR* pexcepinfo);
    STDMETHODIMP OnStateChange(SCRIPTSTATE ssScriptState);
    STDMETHODIMP OnScriptError(IActiveScriptError __RPC_FAR* pscripterror);
    STDMETHODIMP OnEnterScript(void);
    STDMETHODIMP OnLeaveScript(void);

    // IActiveScriptSiteWindow
    // Optional interface to add support for graphical interaction (ex. InputBox or MsgBox)
    STDMETHODIMP GetWindow(HWND* phwnd);
    STDMETHODIMP EnableModeless(BOOL fEnable);

private:
    ULONG referenceCount;
    std::vector<NamedItem> namedItems;
};