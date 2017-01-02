// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files:
#include <windows.h>
#include <Windowsx.h>
#include <ole2.h>
#include <olectl.h>

#include <shlobj.h>
#include <shlwapi.h>
#include <stdio.h>
#include <tchar.h>
#include <math.h>
#include <Propvarutil.h>

#include <vector>

#include <strsafe.h>
#include "G:\Users\Greg\Documents\Visual Studio 2015\Projects\MyUtils\MyUtils\myutils.h"
#include "resource.h"

class CServer;
CServer * GetServer();

// definitions
#define				MY_CLSID					CLSID_SciDataSet
#define				MY_LIBID					LIBID_SciDataSet
#define				MY_IID						IID_ISciDataSet
#define				MY_IIDSINK					IID__SciDataSet

#define				MAX_CONN_PTS				1
#define				CONN_PT_CUSTOMSINK			0

#define				FRIENDLY_NAME				TEXT("SciDataSet")
#define				PROGID						TEXT("Sciencetech.SciDataSet.1")
#define				VERSIONINDEPENDENTPROGID	TEXT("Sciencetech.SciDataSet")