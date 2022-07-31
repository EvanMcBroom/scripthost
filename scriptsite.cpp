#include <algorithm>
#include <iostream>
#include <scriptsite.hpp>

extern ULONG globalObjectCount;

GenericScriptSite::GenericScriptSite(const std::vector<NamedItem>& namedItems)
    : referenceCount(1) {
    InterlockedIncrement(&globalObjectCount);
    this->namedItems = namedItems;
}

GenericScriptSite::~GenericScriptSite() {
    std::for_each(namedItems.begin(), namedItems.end(), [](NamedItem& namedItem) {
        namedItem.second->Release();
        namedItem.second = nullptr;
    });
    InterlockedDecrement(&globalObjectCount);
}

HRESULT STDMETHODCALLTYPE GenericScriptSite::QueryInterface(REFIID riid, void** ppv) {
    if (IID_IUnknown == riid || IID_IActiveScriptSite == riid)
        *ppv = dynamic_cast<IActiveScriptSite*>(this);
    else if (IID_IActiveScriptSiteWindow == riid)
        *ppv = dynamic_cast<IActiveScriptSiteWindow*>(this);
    else
        return E_NOINTERFACE;
    reinterpret_cast<IUnknown*>(*ppv)->AddRef();
    return S_OK;
}

ULONG STDMETHODCALLTYPE GenericScriptSite::AddRef() {
    return InterlockedIncrement(&referenceCount);
}

ULONG STDMETHODCALLTYPE GenericScriptSite::Release() {
    ULONG count{ InterlockedDecrement(&referenceCount) };
    if (!count)
        delete this;
    return count;
}

HRESULT STDMETHODCALLTYPE GenericScriptSite::GetLCID(LCID __RPC_FAR* plcid) {
    *plcid = GetSystemDefaultLCID();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE GenericScriptSite::GetItemInfo(LPCOLESTR pstrName, DWORD dwReturnMask, IUnknown __RPC_FAR* __RPC_FAR* ppunkItem, ITypeInfo __RPC_FAR* __RPC_FAR* ppTypeInfo) {
    auto namedItem{ std::find_if(namedItems.begin(), namedItems.end(), [pstrName](NamedItem& namedItem) {
        return !namedItem.first.compare(pstrName);
    }) };
    if (namedItem != namedItems.end()) {
        if (ppunkItem) {
            *ppunkItem = nullptr;
            if (ppTypeInfo)
                *ppTypeInfo = nullptr;
            if (!(dwReturnMask & SCRIPTINFO_ITYPEINFO)) {
                if (dwReturnMask & SCRIPTINFO_IUNKNOWN) {
                    namedItem->second->QueryInterface(IID_IUnknown, (void**)ppunkItem);
                }
                return S_OK;
            }
            return E_INVALIDARG;
        }
        return E_POINTER;
    }
    return TYPE_E_ELEMENTNOTFOUND;
}

HRESULT STDMETHODCALLTYPE GenericScriptSite::GetDocVersionString(BSTR __RPC_FAR* pbstrVersion) {
    return E_NOTIMPL;
}

// May be called by the engine after script completetion and before calling OnStateChange
HRESULT STDMETHODCALLTYPE GenericScriptSite::OnScriptTerminate(const VARIANT __RPC_FAR* pvarResult, const EXCEPINFO __RPC_FAR* pexcepinfo) {
    return S_OK;
}

HRESULT STDMETHODCALLTYPE GenericScriptSite::OnStateChange(SCRIPTSTATE ssScriptState) {
    return S_OK;
}

HRESULT STDMETHODCALLTYPE GenericScriptSite::OnScriptError(IActiveScriptError __RPC_FAR* pase) {
    EXCEPINFO info;
    if (SUCCEEDED(pase->GetExceptionInfo(&info))) {
        std::wcout << info.bstrSource << L" (0x" << std::hex << info.scode << L")" << std::endl;
        DWORD context, line;
        long position;
        pase->GetSourcePosition(&context, &line, &position);
        std::wcout << L"Line " << line << L", position " << position << L": " << info.bstrDescription << std::endl;
    }
    return S_OK;
}

// Called by the engine when the script starts or continues, including for handling an event
HRESULT STDMETHODCALLTYPE GenericScriptSite::OnEnterScript() {
    return S_OK;
}

// Called by the engine when the script pauses or ends, including after handling an event
HRESULT STDMETHODCALLTYPE GenericScriptSite::OnLeaveScript() {
    return S_OK;
}

HRESULT STDMETHODCALLTYPE GenericScriptSite::GetWindow(HWND* phwnd) {
    *phwnd = GetActiveWindow();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE GenericScriptSite::EnableModeless(BOOL fEnable) {
    return S_OK;
}