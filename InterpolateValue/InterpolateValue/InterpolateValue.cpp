// InterpolateValue.cpp : Defines the exported functions for the DLL application.
//

#include "stdafx.h"
#include "MyGuids.h"
#include "MyFactory.h"

// local functions
BOOL SetRegKeyValue(
	LPTSTR pszKey,
	LPTSTR pszSubkey,
	LPTSTR pszValue);
HRESULT	MyStringFromCLSID(
	REFGUID		refGuid,
	LPTSTR		szCLSID);
// exposed functions

STDAPI DllCanUnloadNow(void)
{
	return GetObjectCount() == 0 ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(
	_In_  REFCLSID rclsid,
	_In_  REFIID   riid,
	_Out_ LPVOID   *ppv)
{
	CMyFactory		*	pMyFactory = NULL;
	HRESULT				hr = CLASS_E_CLASSNOTAVAILABLE;

	*ppv = NULL;
	if (MY_CLSID == rclsid)
	{
		// main COM object
		pMyFactory = new CMyFactory();
	}
	else
	{
		return hr;
	}
	if (NULL != pMyFactory)
	{
		objectsUp();
		hr = pMyFactory->QueryInterface(riid, ppv);
		if (FAILED(hr))
		{
			objectsDown();
			delete pMyFactory;
		}
	}
	else
		hr = E_OUTOFMEMORY;
	return hr;
}

STDAPI DllRegisterServer(void)
{
	HRESULT			hr = NOERROR;
	TCHAR			szID[MAX_PATH];
	TCHAR			szCLSID[MAX_PATH];
	TCHAR			szModulePath[MAX_PATH];
	//	WCHAR			wszModulePath[MAX_PATH];
	ITypeLib	*	pITypeLib;
	//	TCHAR			szString[MAX_PATH];
	//	TCHAR			szStr2[MAX_PATH];
	WCHAR			wszID[MAX_PATH];
	LPTSTR			szTypeLib = NULL;

	// Obtain the path to this module's executable file for later use.
	GetModuleFileName(GetOurInstance(), szModulePath, sizeof(szModulePath) / sizeof(TCHAR));

	// Create some base key strings.
	MyStringFromCLSID(MY_CLSID, szID);
	StringCchCopy(szCLSID, MAX_PATH, TEXT("CLSID\\"));
	StringCchCat(szCLSID, MAX_PATH, szID);

	// Create ProgID keys.
	SetRegKeyValue(PROGID, NULL, FRIENDLY_NAME);
	SetRegKeyValue(PROGID, TEXT("CLSID"), szID);

	// Create VersionIndependentProgID keys.
	SetRegKeyValue(VERSIONINDEPENDENTPROGID, NULL, FRIENDLY_NAME);
	SetRegKeyValue(VERSIONINDEPENDENTPROGID, TEXT("CurVer"), PROGID);
	SetRegKeyValue(VERSIONINDEPENDENTPROGID, TEXT("CLSID"), szID);
	// Create entries under CLSID.
	SetRegKeyValue(szCLSID, NULL, FRIENDLY_NAME);
	SetRegKeyValue(szCLSID, TEXT("ProgID"), PROGID);
	SetRegKeyValue(szCLSID, TEXT("VersionIndependentProgID"), VERSIONINDEPENDENTPROGID);
	SetRegKeyValue(szCLSID, TEXT("NotInsertable"), NULL);
	SetRegKeyValue(szCLSID, TEXT("InprocServer32"), szModulePath);
	// type library
	StringFromGUID2(MY_LIBID, wszID, MAX_PATH);
	SHStrDup(wszID, &szTypeLib);
	if (NULL != szTypeLib)
	{
		SetRegKeyValue(szCLSID, TEXT("TypeLib"), szTypeLib);
		CoTaskMemFree((LPVOID)szTypeLib);
	}
	// version number
	SetRegKeyValue(szCLSID, TEXT("Version"), TEXT("1.0"));
	// register the type library
	//	MultiByteToWideChar(
	//		CP_ACP,
	//		0,
	//		szModulePath,
	//		-1,
	//		wszModulePath,
	//		MAX_PATH);
	hr = LoadTypeLibEx(szModulePath, REGKIND_REGISTER, &pITypeLib);
	if (SUCCEEDED(hr))
		pITypeLib->Release();

	return hr;
}

STDAPI DllUnregisterServer(void)
{
	HRESULT  hr = S_OK;
	TCHAR    szID[MAX_PATH];
	TCHAR    szCLSID[MAX_PATH];
	TCHAR    szTemp[MAX_PATH];

	//Create some base key strings.
	MyStringFromCLSID(MY_CLSID, szID);
	StringCchCopy(szCLSID, MAX_PATH, TEXT("CLSID\\"));
	StringCchCopy(szCLSID, MAX_PATH, szID);

	// un register the entries under Version independent prog ID
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\CurVer"), VERSIONINDEPENDENTPROGID);
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\CLSID"), VERSIONINDEPENDENTPROGID);
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	RegDeleteKey(HKEY_CLASSES_ROOT, VERSIONINDEPENDENTPROGID);

	// delete the entries under prog ID
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\CLSID"), PROGID);
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	RegDeleteKey(HKEY_CLASSES_ROOT, PROGID);

	// delete the entries under CLSID
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\%s"), szCLSID, TEXT("ProgID"));
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\%s"), szCLSID, TEXT("VersionIndependentProgID"));
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\%s"), szCLSID, TEXT("NotInsertable"));
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\%s"), szCLSID, TEXT("InprocServer32"));
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	RegDeleteKey(HKEY_CLASSES_ROOT, szCLSID);

	// unregister the type library
	hr = UnRegisterTypeLib(MY_LIBID, 1, 0, 0x09, SYS_WIN32);
	return S_OK;
}



/*F+F++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function: SetRegKeyValue

Summary:  Internal utility function to set a Key, Subkey, and value
in the system Registry under HKEY_CLASSES_ROOT.

Args:     LPTSTR pszKey,
LPTSTR pszSubkey,
LPTSTR pszValue)

Returns:  BOOL
TRUE if success; FALSE if not.
------------------------------------------------------------------------F-F*/
BOOL SetRegKeyValue(
	LPTSTR pszKey,
	LPTSTR pszSubkey,
	LPTSTR pszValue)
{
	BOOL bOk = FALSE;
	LONG ec;
	HKEY hKey;
	TCHAR szKey[MAX_PATH];

	StringCchCopy(szKey, MAX_PATH, pszKey);

	if (NULL != pszSubkey)
	{
		StringCchCat(szKey, MAX_PATH, TEXT("\\"));
		StringCchCat(szKey, MAX_PATH, pszSubkey);
	}

	ec = RegCreateKeyEx(
		HKEY_CLASSES_ROOT,
		szKey,
		0,
		NULL,
		REG_OPTION_NON_VOLATILE,
		KEY_ALL_ACCESS,
		NULL,
		&hKey,
		NULL);

	if (ERROR_SUCCESS == ec)
	{
		if (NULL != pszValue)
		{
			ec = RegSetValueEx(
				hKey,
				NULL,
				0,
				REG_SZ,
				(BYTE *)pszValue,
				(lstrlen(pszValue) + 1) * sizeof(TCHAR));
		}
		if (ERROR_SUCCESS == ec)
			bOk = TRUE;
		RegCloseKey(hKey);
	}
	return bOk;
}

HRESULT	MyStringFromCLSID(
	REFGUID		refGuid,
	LPTSTR		szCLSID)
{
	HRESULT				hr;
	LPOLESTR			wszCLSID;

	hr = StringFromCLSID(refGuid, &wszCLSID);
	if (SUCCEEDED(hr))
	{
		// copy to output string
		StringCchCopy(szCLSID, MAX_PATH, wszCLSID);
		CoTaskMemFree((LPVOID)wszCLSID);
	}
	return hr;
}

