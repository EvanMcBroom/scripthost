#pragma once
#include <Windows.h>
#include <initguid.h> // Required to use DEFINE_GUID

// CLSID for our IActiveScriptingSite implementation
// {98d2b558-bc82-4b80-9213-8be8d9f5aff4}
DEFINE_GUID(CLSID_GenericScriptSite, 0x98d2b558, 0xbc82, 0x4b80, 0x92, 0x13, 0x8b, 0xe8, 0xd9, 0xf5, 0xaf, 0xf4);

STDAPI DllGetClassObject(IN REFCLSID rclsid, IN REFIID riid, OUT LPVOID* ppv);
STDAPI DllCanUnloadNow();
STDAPI DllRegisterServer(void);
STDAPI DllUnregisterServer(void);