#include <comfactory.hpp>
#include <scriptsite.hpp>

extern ULONG globalObjectCount;
extern ULONG globalLockCount;

GenericScriptSiteFactory::GenericScriptSiteFactory() : referenceCount(1) {
    InterlockedIncrement(&globalObjectCount);
};

GenericScriptSiteFactory::~GenericScriptSiteFactory() {
    InterlockedDecrement(&globalObjectCount);
}

STDMETHODIMP GenericScriptSiteFactory::QueryInterface(REFIID riid, LPVOID* ppv) {
    *ppv = nullptr;
    if (riid == IID_IUnknown || riid == IID_IClassFactory) {
        *ppv = this;
        reinterpret_cast<LPUNKNOWN>(*ppv)->AddRef();
        return NOERROR;
    }
    return ResultFromScode(E_NOINTERFACE);
}

STDMETHODIMP_(ULONG) GenericScriptSiteFactory::AddRef() {
    InterlockedIncrement(&referenceCount);
    return referenceCount;
}

STDMETHODIMP_(ULONG) GenericScriptSiteFactory::Release() {
    InterlockedDecrement(&referenceCount);
    if (referenceCount != 0)
        return referenceCount;
    delete this;
    return 0;
}

STDMETHODIMP GenericScriptSiteFactory::CreateInstance(IN LPUNKNOWN pUnkOuter, IN REFIID riid, OUT PPVOID ppvObj) {
    *ppvObj = nullptr;
    if (!pUnkOuter) {
        IUnknown* object;
        if ((object = reinterpret_cast<IUnknown*>(new GenericScriptSite))) {
            HRESULT result{ object->QueryInterface(riid, ppvObj) };
            if (FAILED(result))
                delete object;
            return result;
        }
        return E_OUTOFMEMORY;
    }
    return CLASS_E_NOAGGREGATION;
}

STDMETHODIMP GenericScriptSiteFactory::LockServer(IN BOOL fLock){
    (fLock) ? InterlockedIncrement(&globalLockCount) : InterlockedDecrement(&globalLockCount);
    return NOERROR;
}