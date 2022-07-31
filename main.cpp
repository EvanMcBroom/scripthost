#include <ActivScp.h>
#include <comserver.hpp>
#include <fstream>
#include <iostream>
#include <sstream>
#include <stdio.h>
#include <vector>

// The user may have decided to not register the COM server on their machine
// Internally resolve CLSID_GenericScriptSite to account for this situation
HRESULT CreateGenericScriptSite(IActiveScriptSite** site) {
    IClassFactory* factory;
    HRESULT result{ DllGetClassObject(CLSID_GenericScriptSite, IID_IClassFactory, reinterpret_cast<void**>(&factory)) };
    if (SUCCEEDED(result)) {
        result = factory->CreateInstance(nullptr, IID_IActiveScriptSite, reinterpret_cast<void**>(site));
    }
    return result;

}

IActiveScript* InitializeEngine(wchar_t* engineProgId) {
    IActiveScript* engine{ nullptr };
    IActiveScriptSite* site;
    if (SUCCEEDED(CreateGenericScriptSite(&site))) {
        CLSID engineClsid;
        if (SUCCEEDED(CLSIDFromProgID(engineProgId, &engineClsid))) {
            if (SUCCEEDED(CoCreateInstance(engineClsid, nullptr, CLSCTX_INPROC_SERVER, IID_IActiveScript, (void**)&engine))) {
                bool successful{ false };
                IActiveScriptParse* parser;
                if (SUCCEEDED(engine->QueryInterface(IID_IActiveScriptParse, reinterpret_cast<void**>(&parser)))) {
                    if (SUCCEEDED(parser->InitNew())) {
                        if (SUCCEEDED(engine->SetScriptSite(site))) {
                            successful = true;
                        }
                    }
                }
                if (!successful) {
                    engine->Release();
                    engine = nullptr;
                }
            }
        }
        site->Release();
    }
    return engine;
}

HRESULT RunScript(IActiveScript* engine, wchar_t* scriptPath) {
    engine->AddRef();
    HRESULT result{ E_FAIL };
    IActiveScriptParse* parser;
    if (SUCCEEDED((result = engine->QueryInterface(IID_IActiveScriptParse, reinterpret_cast<void**>(&parser))))) {
        // Read the script
        std::wifstream ifs{ scriptPath };
        if (!ifs.fail()) {
            std::wstringstream ss;
            ss << ifs.rdbuf();
            std::wstring script{ ss.str() };
            // Parse the script
            EXCEPINFO exceptionInfo;
            if (SUCCEEDED((result = parser->ParseScriptText(script.data(), nullptr, nullptr, nullptr, 0, 0, 0, nullptr, &exceptionInfo)))) {
                result = engine->SetScriptState(SCRIPTSTATE_CONNECTED);
            }
        }
        else {
            result = E_FAIL;
        }
    }
    engine->Release();
    return result;
}

int wmain(int argc, wchar_t* argv[]) {
    if (argc > 2) {
        if (SUCCEEDED(CoInitialize(nullptr))) {
            IActiveScript* engine;
            if ((engine = InitializeEngine(argv[1]))) {
                for (int index{ 2 }; index < argc; index++) {
                    std::wcout << L"[+] Running: " << argv[index] << std::endl;
                    if (SUCCEEDED(RunScript(engine, argv[index]))) {
                        std::wcout << L"[*] Script completed" << std::endl;
                    }
                    else {
                        std::wcout << L"[-] Could not run script " << std::endl;
                    }
                }
                engine->Release();
            }
            else {
                std::wcout << L"[-] Unrecognized engine name" << std::endl;
            }
            CoUninitialize();
            return 0;
        }
        std::wcout << L"[-] Could not initialize COM" << std::endl;
    }
    std::wcout << L"[-] Invalid syntax:\n" << argv[0] << " <engine name> <script path>..." << std::endl;
    return 1;
}