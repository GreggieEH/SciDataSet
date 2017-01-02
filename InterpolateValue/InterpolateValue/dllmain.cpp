// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"
#include <initguid.h>
#include "MyGuids.h"

ULONG		g_objects = 0;
HINSTANCE	g_hInst = NULL;

BOOL APIENTRY DllMain(HMODULE hModule,
	DWORD  ul_reason_for_call,
	LPVOID lpReserved
)
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		// store the instance handle
		g_hInst = hModule;
		// fall through
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}

// global functions
void objectsUp()
{
	InterlockedIncrement(&g_objects);
}
void objectsDown()
{
	InterlockedDecrement(&g_objects);
}

ULONG GetObjectCount()
{
	return g_objects;
}

HINSTANCE GetOurInstance()
{
	return g_hInst;
}

HRESULT GetTypeLib(ITypeLib ** ppTypeLib)
{
	HRESULT				hr;

	*ppTypeLib = NULL;
	hr = LoadRegTypeLib(MY_LIBID, 1, 0, LOCALE_SYSTEM_DEFAULT, ppTypeLib);
	if (FAILED(hr))
	{
		TCHAR				szModule[MAX_PATH];
		GetModuleFileName((HMODULE)GetOurInstance(), szModule, MAX_PATH);
		hr = LoadTypeLibEx(szModule, REGKIND_REGISTER, ppTypeLib);
	}
	return hr;
}
