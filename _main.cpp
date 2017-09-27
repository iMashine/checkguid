#define _WIN32_DCOM
#include <iostream>
using namespace std;
#include <comdef.h>
#include <Wbemidl.h>

#pragma comment(lib, "wbemuuid.lib")

class systemQuery{
private:
	HRESULT hres;

	// Use the IWbemServices pointer to make requests of WMI ----
	IEnumWbemClassObject* pEnumerator = NULL;

	IWbemLocator *pLoc = NULL;
	IWbemServices *pSvc = NULL;

	bool Initialize(){
		// Step 1: --------------------------------------------------
		// Initialize COM. ------------------------------------------
		hres = CoInitializeEx(0, COINIT_MULTITHREADED);
		if (FAILED(hres)) {
			cout << "Failed to initialize COM library. Error code = 0x"
				<< hex << hres << endl;
			return 1;                  // Program has failed.
		}

		// Step 2: --------------------------------------------------
		// Set general COM security levels --------------------------
		hres = CoInitializeSecurity(
			NULL,
			-1,                          // COM authentication
			NULL,                        // Authentication services
			NULL,                        // Reserved
			RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
			RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
			NULL,                        // Authentication info
			EOAC_NONE,                   // Additional capabilities 
			NULL                         // Reserved
			);

		if (FAILED(hres)) {
			cout << "Failed to initialize security. Error code = 0x"
				<< hex << hres << endl;
			CoUninitialize();
			return 1;                    // Program has failed.
		}

		// Step 3: ---------------------------------------------------
		// Obtain the initial locator to WMI -------------------------
		hres = CoCreateInstance(
			CLSID_WbemLocator,
			0,
			CLSCTX_INPROC_SERVER,
			IID_IWbemLocator, (LPVOID *)&pLoc);

		if (FAILED(hres)) {
			cout << "Failed to create IWbemLocator object."
				<< " Err code = 0x"
				<< hex << hres << endl;
			CoUninitialize();
			return 1;                 // Program has failed.
		}

		// Step 4: -----------------------------------------------------
		// Connect to the root\cimv2 namespace with
		// the current user and obtain pointer pSvc
		// to make IWbemServices calls.
		hres = pLoc->ConnectServer(
			_bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
			NULL,                    // User name. NULL = current user
			NULL,                    // User password. NULL = current
			0,                       // Locale. NULL indicates current
			NULL,                    // Security flags.
			0,                       // Authority (for example, Kerberos)
			0,                       // Context object 
			&pSvc                    // pointer to IWbemServices proxy
			);

		if (FAILED(hres)) {
			cout << "Could not connect. Error code = 0x"
				<< hex << hres << endl;
			pLoc->Release();
			CoUninitialize();
			return 1;                // Program has failed.
		}

		//cout << "Connected to ROOT\\CIMV2 WMI namespace" << endl;

		// Step 5: --------------------------------------------------
		// Set security levels on the proxy -------------------------
		hres = CoSetProxyBlanket(
			pSvc,                        // Indicates the proxy to set
			RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
			RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
			NULL,                        // Server principal name 
			RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
			RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
			NULL,                        // client identity
			EOAC_NONE                    // proxy capabilities 
			);

		if (FAILED(hres)) {
			cout << "Could not set proxy blanket. Error code = 0x"
				<< hex << hres << endl;
			pSvc->Release();
			pLoc->Release();
			CoUninitialize();
			return 1;               // Program has failed.
		}
	}

	bool state = true;

public:
	systemQuery(){
		if (!Initialize()){
			state = false;
		}
	}

	bool getState(){
		return state;
	}

	string getInfoAboutMotherBoard(){
		IWbemClassObject *pclsObj = NULL;
		ULONG uReturn = 0;

		hres = pSvc->ExecQuery(
			bstr_t("WQL"),
			bstr_t("SELECT * FROM Win32_BaseBoard"),
			WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
			NULL,
			&pEnumerator);

		if (FAILED(hres)) {
			cout << "Query for operating system name failed."
				<< " Error code = 0x"
				<< hex << hres << endl;
			pSvc->Release();
			pLoc->Release();
			CoUninitialize();
			return NULL;               // Program has failed.
		}

		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
			&pclsObj, &uReturn);

		VARIANT vtProp;
		wstring value = L"MB&";

		// Get the value of the Name property
		hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
		if (vtProp.bstrVal != NULL)
			value += vtProp.bstrVal;

		value += L"&";

		// Get the value of the Name property
		hr = pclsObj->Get(L"Product", 0, &vtProp, 0, 0);
		if (vtProp.bstrVal != NULL)
			value += vtProp.bstrVal;

		value += L"&";

		hr = pclsObj->Get(L"SerialNumber", 0, &vtProp, 0, 0);
		if (vtProp.bstrVal != NULL)
			value += vtProp.bstrVal;

		VariantClear(&vtProp);
		pclsObj->Release();

		return string(value.begin(), value.end());
	}
	
	string getInfoAboutBios(){
		IWbemClassObject *pclsObj = NULL;
		ULONG uReturn = 0;

		hres = pSvc->ExecQuery(
			bstr_t("WQL"),
			bstr_t("SELECT * FROM Win32_BIOS"),
			WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
			NULL,
			&pEnumerator);

		if (FAILED(hres)) {
			cout << "Query for operating system name failed."
				<< " Error code = 0x"
				<< hex << hres << endl;
			pSvc->Release();
			pLoc->Release();
			CoUninitialize();
			return NULL;               // Program has failed.
		}

		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
			&pclsObj, &uReturn);

		VARIANT vtProp;
		wstring value = L"BIOS&";

		// Get the value of the Name property
		hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
		if (vtProp.bstrVal != NULL)
			value += vtProp.bstrVal;

		value += L"&";

		// Get the value of the Name property
		hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
		if (vtProp.bstrVal != NULL)
			value += vtProp.bstrVal;

		value += L"&";

		// Get the value of the Name property
		hr = pclsObj->Get(L"Version", 0, &vtProp, 0, 0);
		if (vtProp.bstrVal != NULL)
			value += vtProp.bstrVal;

		value += L"&";

		hr = pclsObj->Get(L"SerialNumber", 0, &vtProp, 0, 0);
		if (vtProp.bstrVal != NULL)
			value += vtProp.bstrVal;

		VariantClear(&vtProp);
		pclsObj->Release();

		return string(value.begin(), value.end());
	}

	string getInfoAboutHDD(){
		IWbemClassObject *pclsObj = NULL;
		ULONG uReturn = 0;

		hres = pSvc->ExecQuery(
			bstr_t("WQL"),
			bstr_t("SELECT * FROM Win32_DiskDrive"),
			WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
			NULL,
			&pEnumerator);

		if (FAILED(hres)) {
			cout << "Query for operating system name failed."
				<< " Error code = 0x"
				<< hex << hres << endl;
			pSvc->Release();
			pLoc->Release();
			CoUninitialize();
			return NULL;               // Program has failed.
		}

		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
			&pclsObj, &uReturn);

		VARIANT vtProp;
		wstring value = L"HDD&";

		// Get the value of the Name property
		hr = pclsObj->Get(L"Model", 0, &vtProp, 0, 0);
		if (vtProp.bstrVal != NULL)
			value += vtProp.bstrVal;

		value += L"&";

		// Get the value of the Name property
		hr = pclsObj->Get(L"FirmwareRevision", 0, &vtProp, 0, 0);
		if (vtProp.bstrVal != NULL)
			value += vtProp.bstrVal;

		value += L"&";

		hr = pclsObj->Get(L"SerialNumber", 0, &vtProp, 0, 0);
		if (vtProp.bstrVal != NULL)
			value += vtProp.bstrVal;

		VariantClear(&vtProp);
		pclsObj->Release();

		return string(value.begin(), value.end());
	}

	string getInfoAboutCPU(){
		IWbemClassObject *pclsObj = NULL;
		ULONG uReturn = 0;

		hres = pSvc->ExecQuery(
			bstr_t("WQL"),
			bstr_t("SELECT * FROM Win32_Processor"),
			WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
			NULL,
			&pEnumerator);

		if (FAILED(hres)) {
			cout << "Query for operating system name failed."
				<< " Error code = 0x"
				<< hex << hres << endl;
			pSvc->Release();
			pLoc->Release();
			CoUninitialize();
			return NULL;               // Program has failed.
		}

		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
			&pclsObj, &uReturn);

		VARIANT vtProp;
		wstring value = L"CPU&";

		// Get the value of the Name property
		hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
		if (vtProp.bstrVal != NULL)
			value += vtProp.bstrVal;

		value += L"&";

		// Get the value of the Name property
		hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
		if (vtProp.bstrVal != NULL)
			value += vtProp.bstrVal;

		VariantClear(&vtProp);
		pclsObj->Release();

		return string(value.begin(), value.end());
	}

	~systemQuery(){
		// Cleanup
		// ========

		pSvc->Release();
		pLoc->Release();
		pEnumerator->Release();
		CoUninitialize();
	}
};

int main(){

	systemQuery sQuery;

	if (sQuery.getState()){
		cout << sQuery.getInfoAboutMotherBoard().c_str() << endl;
		cout << sQuery.getInfoAboutBios().c_str() << endl;
		cout << sQuery.getInfoAboutHDD().c_str() << endl;
		cout << sQuery.getInfoAboutCPU().c_str() << endl;
	}

	getchar();
	return 0;   // Program successfully completed.
}