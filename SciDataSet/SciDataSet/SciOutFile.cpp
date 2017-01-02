#include "stdafx.h"
#include "SciOutFile.h"

CSciOutFile::CSciOutFile(void) :
	m_pdisp(NULL)
{
	HRESULT				hr;
//	LPOLESTR			ProgID		= NULL;
	CLSID				clsid;
//	Utils_AnsiToUnicode(TEXT("Sciencetech.SciOutFile.1"), &ProgID);
	hr = CLSIDFromProgID(L"Sciencetech.SciOutFile.1", &clsid);
//	CoTaskMemFree((LPVOID) ProgID);
	if (SUCCEEDED(hr))
	{
		hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch,
			(LPVOID*) &(this->m_pdisp));
	}
}

CSciOutFile::~CSciOutFile(void)
{
	Utils_RELEASE_INTERFACE(this->m_pdisp);
}

BOOL CSciOutFile::GetMyObject(
	IDispatch	**	ppdisp)
{
	if (NULL != this->m_pdisp)
	{
		*ppdisp = this->m_pdisp;
		this->m_pdisp->AddRef();
		return TRUE;
	}
	else
	{
		*ppdisp = NULL;
		return FALSE;
	}

}


void CSciOutFile::SetfileName(
						LPCTSTR			szfileName)
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("fileName"), &dispid);
	Utils_SetStringProperty(this->m_pdisp, dispid, szfileName);
}

void CSciOutFile::SettimeStamp(
						VARIANT		*	timeStamp)
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("timeStamp"), &dispid);
	Utils_InvokePropertyPut(this->m_pdisp, dispid, timeStamp, 1);
}

void CSciOutFile::SetsourceFile(
						LPCTSTR			szsourceFile)
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("sourceFile"), &dispid);
	Utils_SetStringProperty(this->m_pdisp, dispid, szsourceFile);
}

void CSciOutFile::SetcalibrationFile(
						LPCTSTR			szcalibrationFile)
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("calibrationFile"), &dispid);
	Utils_SetStringProperty(this->m_pdisp, dispid, szcalibrationFile);
}

void CSciOutFile::SetanalysisType(
						LPCTSTR			szanalysisType)
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("analysisType"), &dispid);
	Utils_SetStringProperty(this->m_pdisp, dispid, szanalysisType);
}

void CSciOutFile::SetxValues(
						long			nValues,
						double		*	pxValues)
{
	DISPID			dispid;
	VARIANTARG		varg;
	double		*	arr;
	long			i;
	SAFEARRAYBOUND	sab;

	Utils_GetMemid(this->m_pdisp, TEXT("xValues"), &dispid);
	sab.lLbound		= 0;
	sab.cElements	= nValues;
	VariantInit(&varg);
	varg.vt			= VT_ARRAY | VT_R8;
	varg.parray		= SafeArrayCreate(VT_R8, 1, &sab);
	SafeArrayAccessData(varg.parray, (void**) &arr);
	for (i=0; i<nValues; i++) arr[i]	= pxValues[i];
	SafeArrayUnaccessData(varg.parray);
	Utils_InvokePropertyPut(this->m_pdisp, dispid, &varg, 1);
	VariantClear(&varg);
}

void CSciOutFile::SetyValues(
						long			nValues,
						double		*	pyValues)
{
	DISPID			dispid;
	VARIANTARG		varg;
	double		*	arr;
	long			i;
	SAFEARRAYBOUND	sab;

	Utils_GetMemid(this->m_pdisp, TEXT("yValues"), &dispid);
	sab.lLbound		= 0;
	sab.cElements	= nValues;
	VariantInit(&varg);
	varg.vt			= VT_ARRAY | VT_R8;
	varg.parray		= SafeArrayCreate(VT_R8, 1, &sab);
	SafeArrayAccessData(varg.parray, (void**) &arr);
	for (i=0; i<nValues; i++) arr[i]	= pyValues[i];
	SafeArrayUnaccessData(varg.parray);
	Utils_InvokePropertyPut(this->m_pdisp, dispid, &varg, 1);
	VariantClear(&varg);
}

void CSciOutFile::saveFile()
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("saveFile"), &dispid);
	Utils_InvokeMethod(this->m_pdisp, dispid, NULL, 0, NULL);
}

void CSciOutFile::SetComment(
	LPCTSTR			szComment)
{
	IDispatch		*	pdisp;
	DISPID				dispid;
	if (this->Get_clsIDataSet(&pdisp))
	{
		Utils_GetMemid(pdisp, L"Comment", &dispid);
		Utils_SetStringProperty(pdisp, dispid, szComment);
		pdisp->Release();
	}
}

void CSciOutFile::SetyUnits(
	LPCTSTR			szyUnits)
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("yUnits"), &dispid);
	Utils_SetStringProperty(this->m_pdisp, dispid, szyUnits);
}

void CSciOutFile::SetExtraData(
	LPCTSTR			szTitle,
	long			nArray,
	double		*	pWaves,
	double		*	pValues)
{
	IDispatch	*	pdisp;
	DISPID			dispid;
	long			i;
	VARIANTARG		avarg[3];

	if (this->Get_clsIDataSet(&pdisp))
	{
		Utils_GetMemid(pdisp, L"ExtraScanData", &dispid);
		InitVariantFromString(szTitle, &avarg[2]);
		for (i = 0; i < nArray; i++)
		{
			InitVariantFromDouble(pWaves[i], &avarg[1]);
			InitVariantFromDouble(pValues[i], &avarg[0]);
			Utils_InvokeMethod(pdisp, dispid, avarg, 3, NULL);
		}
		VariantClear(&avarg[2]);
		pdisp->Release();
	}
}

BOOL CSciOutFile::Get_clsIDataSet(
	IDispatch	**	ppdisp)
{
	HRESULT				hr;
	ITypeInfo		*	pTypeInfo = NULL;
	TYPEATTR		*	pTypeAttr;
//	DISPID				dispid;

	*ppdisp = NULL;
	hr = Utils_GetNamedTypeInfo(this->m_pdisp, L"_clsIDataSet", &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			hr = this->m_pdisp->QueryInterface(pTypeAttr->guid, (LPVOID*)ppdisp);
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		pTypeInfo->Release();
	}
	return SUCCEEDED(hr);
}
