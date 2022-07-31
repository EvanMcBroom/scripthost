#pragma once
#include <Windows.h>

using PPVOID = LPVOID*;

class GenericScriptSiteFactory : public IClassFactory {
public:
    GenericScriptSiteFactory();
    ~GenericScriptSiteFactory();

    // IUnknown
    STDMETHODIMP QueryInterface(REFIID riid, LPVOID* ppv);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // IClassFactory
    STDMETHODIMP CreateInstance(IN LPUNKNOWN pUnkOuter, IN REFIID riid, OUT PPVOID ppvObj);
    STDMETHODIMP LockServer(IN BOOL fLock);

private:
    ULONG referenceCount;
};