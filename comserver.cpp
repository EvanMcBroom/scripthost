#include <comfactory.hpp>
#include <comserver.hpp>
#include <string>

namespace {
	constexpr auto GuidLength() {
		return sizeof(GUID);
	};
}

ULONG globalObjectCount{ 0 };
ULONG globalLockCount{ 0 };

STDAPI DllGetClassObject(IN REFCLSID rclsid, IN REFIID riid, OUT LPVOID* ppv) {
	if (rclsid == CLSID_GenericScriptSite) {
		auto factory{ new GenericScriptSiteFactory };
		if (factory) {
			HRESULT result{ factory->QueryInterface(riid, ppv) };
			if (FAILED(result))
				delete factory;
			return result;
		}
	}
	return E_FAIL;
}

STDAPI DllCanUnloadNow() {
	return (!globalObjectCount && !globalLockCount) ? S_OK : S_FALSE;
}

STDAPI DllRegisterServer() {
	bool successful{ false };
	// Add a key for the server
	OLECHAR guidString[GuidLength()];
	if (StringFromGUID2(CLSID_GenericScriptSite, guidString, GuidLength())) {
		std::wstring parentKeyName{ L"SOFTWARE\\CLASSES\\CLSID\\" };
		std::wstring clsidKeyName{ parentKeyName };
		clsidKeyName.append(guidString);
		HKEY hKey1;
		if (RegCreateKeyW(HKEY_LOCAL_MACHINE, clsidKeyName.data(), &hKey1) == ERROR_SUCCESS) {
			// Add a description
			auto description{ "Generic Active Scripting Site" };
			RegSetValueExW(hKey1, NULL, 0, REG_SZ, reinterpret_cast<const BYTE*>(description), (strlen(description) + 1));
			// Get the path of the current module
			HMODULE server;
			if (GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT, reinterpret_cast<LPCWSTR>(&DllRegisterServer), &server)) {
				char modulePath[MAX_PATH];
				if (!GetModuleFileNameA(server, modulePath, MAX_PATH)) {
					// Set the server location and apartment type
					HKEY hKey2;
					if (RegCreateKeyW(hKey1, L"InProcServer32", &hKey2) == ERROR_SUCCESS) {
						RegSetValueExW(hKey2, NULL, 0, REG_SZ, reinterpret_cast<BYTE*>(modulePath), (strlen(modulePath) + 1));
						auto model{ "Apartment" };
						RegSetValueExW(hKey2, L"ThreadingModel", 0, REG_SZ, reinterpret_cast<const BYTE*>(model), (strlen(model) + 1));
						RegCloseKey(hKey2);
						successful = true;
					}
				}
			}
			RegCloseKey(hKey1);
		}
	}
	return (successful) ? S_OK : S_FALSE;
}

STDAPI DllUnregisterServer() {
	OLECHAR guidString[GuidLength()];
	if (StringFromGUID2(CLSID_GenericScriptSite, guidString, GuidLength())) {
		std::wstring parentKeyName{ L"SOFTWARE\\CLASSES\\CLSID\\" };
		std::wstring clsidKeyName{ parentKeyName };
		clsidKeyName.append(guidString);
		HKEY hKey;
		if (RegOpenKeyW(HKEY_LOCAL_MACHINE, clsidKeyName.data(), &hKey)) {
			RegDeleteKeyW(hKey, L"InProcServer32");
			RegCloseKey(hKey);
			if (RegOpenKeyW(HKEY_LOCAL_MACHINE, parentKeyName.data(), &hKey)) {
				RegDeleteKeyW(hKey, guidString);
				RegCloseKey(hKey);
				return S_OK;
			}
		}
	}
	return S_FALSE;
}