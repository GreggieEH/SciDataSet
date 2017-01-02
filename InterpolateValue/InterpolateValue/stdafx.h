// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files:
#include <windows.h>
#include <Ole2.h>
#include <OCIdl.h>
#include <olectl.h>

// c run time
#include <stdio.h>
#include <tchar.h>
#include <math.h>

// shell utilities
#include <Shlwapi.h>
#include <Shobjidl.h>
#include <Shlobj.h>

// 
#include <propvarutil.h>

#include <vector>

// string safe
#include <strsafe.h>

// some utilities
#include "G:\Users\Greg\Documents\Visual Studio 2015\Projects\MyUtils\MyUtils\MyUtils.h"

// global functions
void objectsUp();
void objectsDown();
ULONG GetObjectCount();
HINSTANCE GetOurInstance();
HRESULT GetTypeLib(ITypeLib ** ppTypeLib);

// our guids
#define			MY_CLSID				CLSID_InterpolateValue
#define			MY_LIBID				LIBID_InterpolateValue
#define			MY_IID					IID_IInterpolateValue
#define			MY_IIDSINK				IID__InterpolateValue

// connection points
#define			MAX_CONN_PTS			2
#define			CONN_PT_CUSTOMSINK		0
#define			CONN_PT_PROPNOTIFY		1

// registry entry strings
#define			FRIENDLY_NAME				L"InterpolateValue"
#define			PROGID						L"Sciencetech.InterpolateValue.1"
#define			VERSIONINDEPENDENTPROGID	L"Sciencetech.InterpolateValue"