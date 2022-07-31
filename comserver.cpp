#include <comfactory.hpp>
#include <comserver.hpp>
#include <string>

namespace {
	constexpr auto GuidStringLength() {
		return 38 + 1;
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
	OLECHAR guidString[GuidStringLength()];
	if (StringFromGUID2(CLSID_GenericScriptSite, guidString, GuidStringLength())) {
		std::wstring parentKeyName{ L"SOFTWARE\\CLASSES\\CLSID\\" };
		std::wstring clsidKeyName{ parentKeyName };
		clsidKeyName.append(guidString);
		HKEY key1;
		if (RegCreateKeyW(HKEY_LOCAL_MACHINE, clsidKeyName.data(), &key1) == ERROR_SUCCESS) {
			// Add a description
			auto description{ L"Generic Active Scripting Site" };
			RegSetValueExW(key1, NULL, 0, REG_SZ, reinterpret_cast<const BYTE*>(description), (lstrlenW(description) + 1) * sizeof(OLECHAR));
			// Get the path of the current module
			HMODULE server;
			if (GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT, reinterpret_cast<LPCWSTR>(&DllRegisterServer), &server)) {
				wchar_t modulePath[MAX_PATH] = { 0 };
				if (GetModuleFileNameW(server, modulePath, MAX_PATH)) {
					// Set the server location and apartment type
					HKEY key2;
					if (RegCreateKeyW(key1, L"InProcServer32", &key2) == ERROR_SUCCESS) {
						RegSetValueExW(key2, NULL, 0, REG_SZ, reinterpret_cast<BYTE*>(modulePath), (lstrlenW(modulePath) + 1) * sizeof(OLECHAR));
						auto model{ L"Apartment" };
						RegSetValueExW(key2, L"ThreadingModel", 0, REG_SZ, reinterpret_cast<const BYTE*>(model), (lstrlenW(model) + 1) * sizeof(OLECHAR));
						RegCloseKey(key2);
						successful = true;
					}
				}
			}
			RegCloseKey(key1);
		}
	}
	return (successful) ? S_OK : S_FALSE;
}

STDAPI DllUnregisterServer() {
	OLECHAR guidString[GuidStringLength()];
	if (StringFromGUID2(CLSID_GenericScriptSite, guidString, GuidStringLength())) {
		std::wstring parentKeyName{ L"SOFTWARE\\CLASSES\\CLSID\\" };
		std::wstring clsidKeyName{ parentKeyName };
		clsidKeyName.append(guidString);
		HKEY key;
		if (RegOpenKeyW(HKEY_LOCAL_MACHINE, clsidKeyName.data(), &key) == ERROR_SUCCESS) {
			RegDeleteKeyW(key, L"InProcServer32");
			RegCloseKey(key);
			if (RegOpenKeyW(HKEY_LOCAL_MACHINE, parentKeyName.data(), &key) == ERROR_SUCCESS) {
				RegDeleteKeyW(key, guidString);
				RegCloseKey(key);
				return S_OK;
			}
		}
	}
	return S_FALSE;
}