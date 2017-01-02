#include "stdafx.h"
#include "GaussianFit.h"

CGaussianFit::CGaussianFit() :
	m_pdisp(NULL)
{
	HRESULT			hr;
	CLSID			clsid;

	hr = CLSIDFromProgID(L"Sciencetech.GaussianFit.1", &clsid);
	if (SUCCEEDED(hr)) hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (LPVOID*)&this->m_pdisp);
}

CGaussianFit::~CGaussianFit()
{
	Utils_RELEASE_INTERFACE(this->m_pdisp);
}

void CGaussianFit::SetPeakThreshold(
	double		PeakThreshold)
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, L"PeakThreshold", &dispid);
	Utils_SetDoubleProperty(this->m_pdisp, dispid, PeakThreshold);
}

BOOL CGaussianFit::SetData(
	long		nArray,
	double	*	pX,
	double	*	pY)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANTARG		avarg[2];
	VARIANT			varResult;
	SAFEARRAYBOUND	sab;
	SAFEARRAY	*	pSAX = NULL;
	SAFEARRAY	*	pSAY = NULL;
	long			i;
	double		*	arrX;
	double		*	arrY;
	BOOL			fSuccess = FALSE;
	Utils_GetMemid(this->m_pdisp, L"SetData", &dispid);
	sab.lLbound = 0;
	sab.cElements = nArray;
	pSAX = SafeArrayCreate(VT_R8, 1, &sab);
	pSAY = SafeArrayCreate(VT_R8, 1, &sab);
	SafeArrayAccessData(pSAX, (void**)&arrX);
	SafeArrayAccessData(pSAY, (void**)&arrY);
	for (i = 0; i < nArray; i++)
	{
		arrX[i] = pX[i];
		arrY[i] = pY[i];
	}
	SafeArrayUnaccessData(pSAX);
	SafeArrayUnaccessData(pSAY);
	VariantInit(&avarg[1]);
	avarg[1].vt = VT_BYREF | VT_ARRAY | VT_R8;
	avarg[1].pparray = &pSAX;
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_ARRAY | VT_R8;
	avarg[0].pparray = &pSAY;
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, avarg, 2, &varResult);
	SafeArrayDestroy(pSAX);
	SafeArrayDestroy(pSAY);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

BOOL CGaussianFit::FitPeaks()
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	Utils_GetMemid(this->m_pdisp, L"FitPeaks", &dispid);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

long CGaussianFit::GetNumPeaks()
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, L"NumPeaks", &dispid);
	return Utils_GetLongProperty(this->m_pdisp, dispid);
}

double CGaussianFit::GetPeakPosition(
	long		index)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANTARG		avarg[2];
	VARIANT			varResult;
	double			dummy = 0.0;
	double			retVal = 0.0;
	Utils_GetMemid(this->m_pdisp, L"PeakPosition", &dispid);
	InitVariantFromInt32(index, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_R8;
	avarg[0].pdblVal = &dummy;
	hr = Utils_InvokePropertyGet(this->m_pdisp, dispid, avarg, 2, &varResult);
	if (SUCCEEDED(hr)) VariantToDouble(varResult, &retVal);
	return retVal;
}

double CGaussianFit::GetPeakHeight(
	long		index)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANTARG		avarg[2];
	VARIANT			varResult;
	double			dummy = 0.0;
	double			retVal = 0.0;
	Utils_GetMemid(this->m_pdisp, L"PeakHeight", &dispid);
	InitVariantFromInt32(index, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_R8;
	avarg[0].pdblVal = &dummy;
	hr = Utils_InvokePropertyGet(this->m_pdisp, dispid, avarg, 2, &varResult);
	if (SUCCEEDED(hr)) VariantToDouble(varResult, &retVal);
	return retVal;
}

double CGaussianFit::GetPeakWidth(
	long		index)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANTARG		avarg[2];
	VARIANT			varResult;
	double			dummy = 0.0;
	double			retVal = 0.0;
	Utils_GetMemid(this->m_pdisp, L"PeakWidth", &dispid);
	InitVariantFromInt32(index, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_R8;
	avarg[0].pdblVal = &dummy;
	hr = Utils_InvokePropertyGet(this->m_pdisp, dispid, avarg, 2, &varResult);
	if (SUCCEEDED(hr)) VariantToDouble(varResult, &retVal);
	return retVal;
}

void CGaussianFit::DisplayPeaks(
	HWND		hwndParent)
{
	DISPID			dispid;
	VARIANTARG		varg;
	Utils_GetMemid(this->m_pdisp, L"DisplayPeaks", &dispid);
	InitVariantFromInt32((long)hwndParent, &varg);
	Utils_InvokeMethod(this->m_pdisp, dispid, &varg, 1, NULL);
}