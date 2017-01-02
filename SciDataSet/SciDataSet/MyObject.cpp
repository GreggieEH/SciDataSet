#include "StdAfx.h"
#include "MyObject.h"
#include "Server.h"
#include "dispids.h"
#include "MyGuids.h"
#include "IDataSet.h"
#include "DlgNewObjectSetup.h"
#include "SciOutFile.h"
#include "GaussianFit.h"

CMyObject::CMyObject(IUnknown * pUnkOuter) :
	m_pImpIDispatch(NULL),
	m_pImpIConnectionPointContainer(NULL),
	m_pImpIProvideClassInfo2(NULL),
	// outer unknown for aggregation
	m_pUnkOuter(pUnkOuter),
	// object reference count
	m_cRefs(0),
	// data set
	m_pIDataSet(NULL),
	// calibration data
	m_nCalData(0),
	m_pXCal(NULL),
	m_pYCal(NULL),
	// Gaussian fit
	m_pGaussianFit(NULL)
{
	if (NULL == this->m_pUnkOuter) this->m_pUnkOuter = this;
	for (ULONG i=0; i<MAX_CONN_PTS; i++)
		this->m_paConnPts[i]	= NULL;
}

CMyObject::~CMyObject(void)
{
	Utils_DELETE_POINTER(this->m_pImpIDispatch);
	Utils_DELETE_POINTER(this->m_pImpIConnectionPointContainer);
	Utils_DELETE_POINTER(this->m_pImpIProvideClassInfo2);
	for (ULONG i=0; i<MAX_CONN_PTS; i++)
		Utils_RELEASE_INTERFACE(this->m_paConnPts[i]);
	Utils_DELETE_POINTER(this->m_pIDataSet);
	// clear the calibration data
	this->SetCalibrationData(0, NULL, NULL);
	// gaussian fit
	Utils_DELETE_POINTER(this->m_pGaussianFit);
}

// IUnknown methods
STDMETHODIMP CMyObject::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	*ppv = NULL;
	if (IID_IUnknown == riid)
		*ppv = this;
	else if (IID_IDispatch == riid || MY_IID == riid)
		*ppv = this->m_pImpIDispatch;
	else if (IID_IConnectionPointContainer == riid)
		*ppv = this->m_pImpIConnectionPointContainer;
	else if (IID_IProvideClassInfo == riid || IID_IProvideClassInfo2 == riid)
		*ppv = this->m_pImpIProvideClassInfo2;
	if (NULL != *ppv)
	{
		((IUnknown*)*ppv)->AddRef();
		return S_OK;
	}
	else
	{
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CMyObject::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyObject::Release()
{
	ULONG			cRefs;
	cRefs = --m_cRefs;
	if (0 == cRefs)
	{
		m_cRefs++;
		GetServer()->ObjectsDown();
		delete this;
	}
	return cRefs;
}

// initialization
HRESULT CMyObject::Init()
{
	HRESULT						hr;

	this->m_pImpIDispatch					= new CImpIDispatch(this, this->m_pUnkOuter);
	this->m_pImpIConnectionPointContainer	= new CImpIConnectionPointContainer(this, this->m_pUnkOuter);
	this->m_pImpIProvideClassInfo2			= new CImpIProvideClassInfo2(this, this->m_pUnkOuter);
	this->m_pIDataSet						= new CIDataSet(this);
	if (NULL != this->m_pImpIDispatch					&&
		NULL != this->m_pImpIConnectionPointContainer	&&
		NULL != this->m_pImpIProvideClassInfo2			&&
		NULL != this->m_pIDataSet)
	{
		hr = Utils_CreateConnectionPoint(this, MY_IIDSINK,
			 &(this->m_paConnPts[CONN_PT_CUSTOMSINK]));
	}
	else hr = E_OUTOFMEMORY;
	return hr;
}

// sink events
double CMyObject::FirerequestDispersion(
									long			gratingID)
{
	VARIANTARG		avarg[2];
	double			dispersion		= 0.0;
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_BYREF | VT_I4;
	avarg[1].plVal		= &gratingID;
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_R8;
	avarg[0].pdblVal	= &dispersion;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_requestDispersion, avarg, 2);
	return dispersion;
}

short int CMyObject::FirerequestAnalysisModeIndex(
									LPCTSTR			szAnalysisMode)
{
	VARIANTARG		avarg[2];
	short int		modeIndex	= -1;
//	LPOLESTR		pOleStr		= NULL;
	BSTR			bstr		= NULL;
//	Utils_AnsiToUnicode(szAnalysisMode, &pOleStr);
	bstr				= SysAllocString(szAnalysisMode);
//	CoTaskMemFree((LPVOID) pOleStr);
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_BYREF | VT_BSTR;
	avarg[1].pbstrVal	= &bstr;
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_I2;
	avarg[0].piVal		= &modeIndex;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_requestAnalysisModeIndex, avarg, 2);
	SysFreeString(bstr);
	return modeIndex;
}

BOOL CMyObject::FirerequestAnalysisModeString(
									short int		ModeIndex,
									LPTSTR			szAnalysisMode,
									UINT			nBufferSize)
{
	VARIANTARG		avarg[2];
	BSTR			bstr		= NULL;
	BOOL			fSuccess	= FALSE;
//	LPTSTR			szString	= NULL;
	szAnalysisMode[0]	= '\0';
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_BYREF | VT_I2;
	avarg[1].piVal		= &ModeIndex;
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_BSTR;
	avarg[0].pbstrVal	= &bstr;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_requestAnalysisModeString, avarg, 2);
	if (NULL != bstr)
	{
//		Utils_UnicodeToAnsi(bstr, &szString);

		StringCchCopy(szAnalysisMode, nBufferSize, bstr);
//		CoTaskMemFree((LPVOID) szString);
		fSuccess	= TRUE;
		SysFreeString(bstr);
	}
	return fSuccess;
}

BOOL CMyObject::FirerequestOpticalTransfer(
									long		*	nValues,
									double		**	ppXValues,
									double		**	ppYValues)
{
/*
	long		i;
	double		x;
	double		y;
	*nValues = 0;
	*ppXValues = NULL;
	*ppYValues = NULL;
	*nValues = this->FireRequestOPTArraySize();
	if ((*nValues) > 0)
	{
		*ppXValues = new double[*nValues];
		*ppYValues = new double[*nValues];
		for (i = 0; i < (*nValues); i++)
		{
			this->FireRequestOPTDataValue(i, &x, &y);
			(*ppXValues)[i] = x;
			(*ppYValues)[i] = y;
		}
		return TRUE;
	}
	else
		return FALSE;
*/
	VARIANTARG			avarg[2];
	SAFEARRAY		*	pSAX		= NULL;
	SAFEARRAY		*	pSAY		= NULL;
//	long				ubx;
//	long				uby;
//	double			*	arrX;
//	double			*	arrY;
	long				i;
	BOOL				fSuccess	= FALSE;
	*nValues	= 0;
	*ppXValues	= NULL;
	*ppYValues	= NULL;
	// clear the calibration data
	this->SetCalibrationData(0, NULL, NULL);
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_BYREF | VT_ARRAY | VT_R8;
	avarg[1].pparray	= &pSAX;
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_ARRAY | VT_R8;
	avarg[0].pparray	= &pSAY;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_requestOpticalTransfer, avarg, 2);
	/*
	if (NULL != pSAX && NULL != pSAY)
	{
		SafeArrayGetUBound(pSAX, 1, &ubx);
		SafeArrayGetUBound(pSAY, 1, &uby);
		if (ubx == uby)
		{
			fSuccess	= TRUE;
			*nValues	= ubx+1;
			*ppXValues	= new double [*nValues];
			*ppYValues	= new double [*nValues];
			SafeArrayAccessData(pSAX, (void**) &arrX);
			SafeArrayAccessData(pSAY, (void**) &arrY);
			for (i=0; i<(*nValues); i++)
			{
				(*ppXValues)[i]	= arrX[i];
				(*ppYValues)[i] = arrY[i];
			}
			SafeArrayUnaccessData(pSAX);
			SafeArrayUnaccessData(pSAY);
		}
	}
	*/
	SafeArrayDestroy(pSAX);
	SafeArrayDestroy(pSAY);
	if (this->m_nCalData > 0)
	{
		fSuccess = TRUE;
		*nValues = this->m_nCalData;
		*ppXValues = new double[this->m_nCalData];
		*ppYValues = new double[this->m_nCalData];
		for (i = 0; i < this->m_nCalData; i++)
		{
			(*ppXValues)[i] = this->m_pXCal[i];
			(*ppYValues)[i] = this->m_pYCal[i];
		}
	}
	return fSuccess;
}

long CMyObject::FireRequestOPTArraySize()
{
	VARIANTARG			varg;
	long				nArray = 0;
	VariantInit(&varg);
	varg.vt = VT_BYREF | VT_I4;
	varg.plVal = &nArray;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestOPTArraySize, &varg, 1);
	return nArray;
}

void CMyObject::FireRequestOPTDataValue(
	long			index,
	double		*	x,
	double		*	y)
{
	VARIANTARG		avarg[3];
	*x = 0.0;
	*y = 0.0;
	InitVariantFromInt32(index, &avarg[2]);
	VariantInit(&avarg[1]);
	avarg[1].vt = VT_BYREF | VT_R8;
	avarg[1].pdblVal = x;
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_R8;
	avarg[0].pdblVal = y;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestOPTDataValue, avarg, 3);
}

void CMyObject::FireRequestSetupDisplay(
	LPCTSTR			component)
{
	VARIANTARG			varg;
	InitVariantFromString(component, &varg);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_RequestSetupDisplay, &varg, 1);
	VariantClear(&varg);
}

void CMyObject::SetCalibrationData(
	long			nValues,
	double		*	pX,
	double		*	pY)
{
	long			i;
	if (NULL != this->m_pXCal)
	{
		delete[] this->m_pXCal;
		this->m_pXCal = NULL;
	}
	if (NULL != this->m_pYCal)
	{
		delete[] this->m_pYCal;
		this->m_pYCal = NULL;
	}
	this->m_nCalData = 0;
	if (nValues > 0)
	{
		this->m_nCalData = nValues;
		this->m_pXCal = new double[this->m_nCalData];
		this->m_pYCal = new double[this->m_nCalData];
		for (i = 0; i < nValues; i++)
		{
			this->m_pXCal[i] = pX[i];
			this->m_pYCal[i] = pY[i];
		}
	}
}

void CMyObject::FireupdateOptFile(
	LPCTSTR			opt)
{
	VARIANTARG			varg;
	InitVariantFromString(opt, &varg);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_updateOptFile, &varg, 1);
	VariantClear(&varg);
}

// flag skip the last data point
BOOL CMyObject::GetSkipLastPoint()
{
	TCHAR			szProgID[MAX_PATH];
	BOOL			fSkipLastPoint = TRUE;
	if (this->m_pIDataSet->GetDataSetProgID(szProgID, MAX_PATH))
	{
		if (0 == lstrcmpi(szProgID, L"Sciencetech.OPTFile.1") || 0 == lstrcmpi(szProgID, L"Sciencetech.SciOutFile.1"))
		{
			fSkipLastPoint = FALSE;
		}
	}
	return fSkipLastPoint;
}

// set the AD info string
BOOL CMyObject::SetADInfoString()
{
	return this->m_pIDataSet->SetADInfoString();
}

BOOL CMyObject::FirerequestADInfoString(
	LPTSTR		*	infoString)
{
	VARIANTARG			avarg[2];
	VARIANT_BOOL		handled = VARIANT_FALSE;
	BSTR				bstr = NULL;
	BOOL				fSuccess = FALSE;
	*infoString = NULL;
	VariantInit(&avarg[1]);
	avarg[1].vt = VT_BYREF | VT_BSTR;
	avarg[1].pbstrVal = &bstr;
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_BOOL;
	avarg[0].pboolVal = &handled;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_requestADInfoString, avarg, 2);
	if (NULL != bstr && VARIANT_TRUE == handled)
	{
		fSuccess = TRUE;
		SHStrDup(bstr, infoString);
	}
	SysFreeString(bstr);
	return fSuccess;
}

BOOL CMyObject::FirerequestCalibrationScan(
	short			grating,
	LPCTSTR			filter,
	short			detector,
	long		*	nValues,
	double		**	ppXValues,
	double		**	ppYValues)
{
	VARIANTARG			avarg[3];
	long				i;
	BOOL				fSuccess = FALSE;
	*nValues = 0;
	*ppXValues = NULL;
	*ppYValues = NULL;
	// clear the calibration data
	this->SetCalibrationData(0, NULL, NULL);
	InitVariantFromInt16(grating, &avarg[2]);
	InitVariantFromString(filter, &avarg[1]);
	InitVariantFromInt16(detector, &avarg[0]);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_requestCalibrationScan, avarg, 3);
	VariantClear(&avarg[1]);
	if (this->m_nCalData > 0)
	{
		fSuccess = TRUE;
		*nValues = this->m_nCalData;
		*ppXValues = new double[this->m_nCalData];
		*ppYValues = new double[this->m_nCalData];
		for (i = 0; i < this->m_nCalData; i++)
		{
			(*ppXValues)[i] = this->m_pXCal[i];
			(*ppYValues)[i] = this->m_pYCal[i];
		}
	}
	return fSuccess;
}

double CMyObject::FirerequestCalibrationGain(
	double			wavelength)
{
	VARIANTARG		avarg[2];
	double			calibrationGainFactor = 1.0;
	InitVariantFromDouble(wavelength, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_R8;
	avarg[0].pdblVal = &calibrationGainFactor;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_requestCalibrationGain, avarg, 2);
	return calibrationGainFactor;
}

BOOL CMyObject::FirerequestWorkingDirectory(
	LPTSTR			szWorkingDirectory,
	UINT			nBufferSize)
{
	VARIANTARG		varg;
	BSTR			bstr = NULL;
	BOOL			fSuccess = FALSE;
	szWorkingDirectory[0] = '\0';
	VariantInit(&varg);
	varg.vt = VT_BYREF | VT_BSTR;
	varg.pbstrVal = &bstr;
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_requestWorkingDirectory, &varg, 1);
	if (NULL != bstr)
	{
		StringCchCopy(szWorkingDirectory, nBufferSize, bstr);
		SysFreeString(bstr);
		fSuccess = TRUE;
	}
	return fSuccess;
}

void CMyObject::FirebrowseForWorkingDirectory(
	HWND			hwndParent)
{
	VARIANTARG			varg;
	InitVariantFromInt32((long)hwndParent, &varg);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_browseForWorkingDirectory, &varg, 1);
}

void CMyObject::setWorkingDirectory(
	LPCTSTR			workingDirectory)
{
	VARIANTARG			varg;
	InitVariantFromString(workingDirectory, &varg);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_setWorkingDirectory, &varg, 1);
	VariantClear(&varg);
}

CMyObject::CImpIDispatch::CImpIDispatch(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter),
	m_pTypeInfo(NULL),
	// radiance calculation
	m_nValuesRadiance(0),
	m_pRadianceWaves(NULL),
	m_pRadianceValues(NULL),
	// optical transfer data
	m_nValuesOPT(0),
	m_pWavesOPT(NULL),
	m_pValuesOPT(NULL)
{
	this->m_szProgID[0]	= '\0';
	// scan type
	StringCchCopy(m_ScanType, MAX_PATH, L"Sample");
}

CMyObject::CImpIDispatch::~CImpIDispatch()
{
	Utils_RELEASE_INTERFACE(this->m_pTypeInfo);
	if (NULL != this->m_pRadianceWaves)
	{
		delete[] this->m_pRadianceWaves;
		this->m_pRadianceWaves = NULL;
	}
	if (NULL != this->m_pRadianceValues)
	{
		delete[] this->m_pRadianceValues;
		this->m_pRadianceValues = NULL;
	}
	this->m_nValuesRadiance = 0;
	// Optical transfer data
	if (NULL != this->m_pWavesOPT)
	{
		delete[] this->m_pWavesOPT;
		this->m_pWavesOPT = NULL;
	}
	if (NULL != this->m_pValuesOPT)
	{
		delete[] this->m_pValuesOPT;
		this->m_pValuesOPT = NULL;
	}
	this->m_nValuesOPT = 0;
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIDispatch::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIDispatch::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIDispatch::Release()
{
	return this->m_punkOuter->Release();
}

// IDispatch methods
STDMETHODIMP CMyObject::CImpIDispatch::GetTypeInfoCount( 
									PUINT			pctinfo)
{
	*pctinfo	= 1;
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIDispatch::GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo)
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;

	*ppTInfo	= NULL;
	if (0 != iTInfo) return DISP_E_BADINDEX;
	if (NULL == this->m_pTypeInfo)
	{
		hr = GetServer()->GetTypeLib(&pTypeLib);
		if (SUCCEEDED(hr))
		{
			hr = pTypeLib->GetTypeInfoOfGuid(MY_IID, &(this->m_pTypeInfo));
			pTypeLib->Release();
		}
	}
	else hr = S_OK;
	if (SUCCEEDED(hr))
	{
		*ppTInfo	= this->m_pTypeInfo;
		this->m_pTypeInfo->AddRef();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIDispatch::GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId)
{
	HRESULT					hr;
	ITypeInfo			*	pTypeInfo;
	hr = this->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgDispId);
		pTypeInfo->Release();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIDispatch::Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr) 
{
	switch (dispIdMember)
	{
	case DISPID_DataProgID:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetDataProgID(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetDataProgID(pDispParams);
		break;
	case DISPID_RadianceAvailable:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetRadianceAvailable(pVarResult);
		break;
	case DISPID_IrradianceAvailable:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetIrradianceAvailable(pVarResult);
		break;
	case DISPID_ReadDataFile:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ReadDataFile(pVarResult);
		break;
	case DISPID_WriteDataFile:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->WriteDataFile(pVarResult);
		break;
	case DISPID_GetWavelengths:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetWavelengths(pVarResult);
		break;
	case DISPID_GetSpectra:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetSpectra(pVarResult);
		break;
	case DISPID_AddValue:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->AddValue(pDispParams, pVarResult);
		break;
	case DISPID_ViewSetup:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ViewSetup(pDispParams, pVarResult);
		break;
	case DISPID_GetArraySize:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetArraySize(pVarResult);
		break;
	case DISPID_SetCurrentExp:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SetCurrentExp(pDispParams);
		break;
	case DISPID_calcRadiance:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->calcRadiance(pDispParams, pVarResult);
		break;
	case DISPID_RadianceWaves:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetRadianceWaves(pVarResult);
		break;
	case DISPID_Radiance:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetRadiance(pVarResult);
		break;
	case DISPID_setObject:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->setObject(pDispParams, pVarResult);
		break;
	case DISPID_AnalysisMode:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetAnalysisMode(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetAnalysisMode(pDispParams);
		break;
	case DISPID_FileName:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetFileName(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetFileName(pDispParams);
		break;
	case DISPID_AmCalibration:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetAmCalibration(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetAmCalibration(pDispParams);
		break;
	case DISPID_GetCreateTime:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetCreateTime(pVarResult);
		break;
	case DISPID_FormOutFile:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->FormOutFile(pDispParams, pVarResult);
		break;
	case DISPID_GetDataValue:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetDataValue(pDispParams, pVarResult);
		break;
	case DISPID_SetDataValue:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SetDataValue(pDispParams);
		break;
	case DISPID_LoadFromString:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->LoadFromString(pDispParams, pVarResult);
		break;
	case DISPID_HaveOPT:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetHaveOPT(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetHaveOPT(pDispParams);
		break;
	case DISPID_OPTWaves:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetOPTWaves(pVarResult);
		break;
	case DISPID_OPTValues:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetOPTValues(pVarResult);
		break;
	case DISPID_FormOpticalTransferFile:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->FormOpticalTransferFile(pDispParams, pVarResult);
		break;
//	case DISPID_SetCalibrationWavelengths:
//		if (0 != (wFlags & DISPATCH_METHOD))
//			return this->SetCalibrationWavelengths(pDispParams);
//		break;
	case DISPID_SetCalibrationData:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SetCalibrationData(pDispParams, pVarResult);
		break;
	case DISPID_FileNameFromPath:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->FileNameFromPath(pDispParams, pVarResult);
		break;
	case DISPID_NewObjectSetup:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->NewObjectSetup(pDispParams, pVarResult);
		break;
	case DISPID_ScanType:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetScanType(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetScanType(pDispParams);
		break;
	case DISPID_SelectCalibrationStandard:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SelectCalibrationStandard(pDispParams, pVarResult);
		break;
	case DISPID_HaveCalibrationStandard:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetHaveCalibrationStandard(pVarResult);
		break;
	case DISPID_Comment:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetComment(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetComment(pDispParams);
		break;
	case DISPID_ExtraScanData:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ExtraScanData(pDispParams);
		break;
	case DISPID_DefaultUnits:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetDefaultUnits(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetDefaultUnits(pDispParams);
		break;
	case DISPID_DefaultTitle:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetDefaultTitle(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetDefaultTitle(pDispParams);
		break;
	case DISPID_CountsAvailable:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetCountsAvailable(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetCountsAvailable(pDispParams);
	case DISPID_GetGratingScan:
		return this->m_pBackObj->m_pIDataSet->GetGratingScan(pDispParams, pVarResult);
	case DISPID_ExportToCSV:
		return this->ExportToCSV(pDispParams, pVarResult);
	case DISPID_GetADGainFactor:
		return this->GetADGainFactor(pDispParams, pVarResult);
	case DISPID_ArrayScanning:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetArrayScanning(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetArrayScanning(pDispParams);
		break;
	case DISPID_SetArrayData:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SetArrayData(pDispParams, pVarResult);
		break;
	case DISPID_NumPeaks:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetNumPeaks(pVarResult);
		break;
	case DISPID_FindPeaks:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->FindPeaks(pDispParams, pVarResult);
		break;
	case DISPID_DisplayPeaks:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->DisplayPeaks(pDispParams);
		break;
	case DISPID_PeakPosition:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetPeakPosition(pDispParams, pVarResult);
		break;
	case DISPID_PeakHeight:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetPeakHeight(pDispParams, pVarResult);
		break;
	case DISPID_PeakWidth:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetPeakWidth(pDispParams, pVarResult);
		break;
	case DISPID_IntegrateSpectra:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->IntegrateSpectra(pDispParams, pVarResult);
		break;
	case DISPID_SmoothSpectra:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SmoothSpectra(pDispParams, pVarResult);
		break;
	case DISPID_CalcAverage:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->CalcAverage(pDispParams, pVarResult);
		break;
	default:
		break;
	}
	return DISP_E_BADPARAMCOUNT;
}

HRESULT CMyObject::CImpIDispatch::GetDataProgID(
									VARIANT		*	pVarResult)
{
	TCHAR			szProgID[MAX_PATH];
	if (NULL == pVarResult) return E_INVALIDARG;
//	Utils_StringToVariant(this->m_szProgID, pVarResult);
//	InitVariantFromString(this->m_szProgID, pVarResult);
	if (this->m_pBackObj->m_pIDataSet->GetDataSetProgID(szProgID, MAX_PATH))
	{
		InitVariantFromString(szProgID, pVarResult);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetDataProgID(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
//	LPTSTR				szString		= NULL;
	CLSID				clsid;
	IDispatch		*	pdisp = NULL;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	hr = CLSIDFromProgID(varg.bstrVal, &clsid);
	if (SUCCEEDED(hr))
	{
		hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (LPVOID*) &pdisp);
	}
	if (SUCCEEDED(hr))
	{
		hr = this->m_pBackObj->m_pIDataSet->SetDataSet(pdisp) ? S_OK : E_FAIL;
		pdisp->Release();
	}
	if (SUCCEEDED(hr))
	{
//		Utils_UnicodeToAnsi(varg.bstrVal, &szString);
		StringCchCopy(this->m_szProgID, MAX_PATH, varg.bstrVal);
//		CoTaskMemFree((LPVOID) szString);
		VariantClear(&varg);
	}
	return hr;
}

HRESULT CMyObject::CImpIDispatch::GetRadianceAvailable(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pIDataSet->GetRadianceAvailable(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetIrradianceAvailable(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pIDataSet->GetIrradianceAvailable(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::ReadDataFile(
									VARIANT		*	pVarResult)
{
	BOOL				fSuccess	= FALSE;
	fSuccess = this->m_pBackObj->m_pIDataSet->ReadDataFile();
	if (NULL != pVarResult)
		InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::WriteDataFile(
									VARIANT		*	pVarResult)
{
	BOOL				fSuccess	= FALSE;
	fSuccess = this->m_pBackObj->m_pIDataSet->WriteDataFile();
	if (NULL != pVarResult)
		InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetWavelengths(
									VARIANT		*	pVarResult)
{
	long			nArray	= 0;
	double		*	pArray	= NULL;
	SAFEARRAYBOUND	sab;
	double		*	arr;
	long			i;

	if (NULL == pVarResult) return E_INVALIDARG;
	if (this->m_pBackObj->m_pIDataSet->GetWavelengths(&nArray, &pArray))
	{
		sab.lLbound			= 0;
		sab.cElements		= nArray;
		VariantInit(pVarResult);
		pVarResult->vt		= VT_ARRAY | VT_R8;
		pVarResult->parray	= SafeArrayCreate(VT_R8, 1, &sab);
		SafeArrayAccessData(pVarResult->parray, (void**) &arr);
		for (i=0; i<nArray; i++)
			arr[i]	= pArray[i];
		SafeArrayUnaccessData(pVarResult->parray);
		delete [] pArray;
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetSpectra(
									VARIANT		*	pVarResult)
{
	long			nArray	= 0;
	double		*	pArray	= NULL;
	SAFEARRAYBOUND	sab;
	double		*	arr;
	long			i;

	if (NULL == pVarResult) return E_INVALIDARG;
	if (this->m_pBackObj->m_pIDataSet->GetSpectra(&nArray, &pArray))
	{
		sab.lLbound			= 0;
		sab.cElements		= nArray;
		VariantInit(pVarResult);
		pVarResult->vt		= VT_ARRAY | VT_R8;
		pVarResult->parray	= SafeArrayCreate(VT_R8, 1, &sab);
		SafeArrayAccessData(pVarResult->parray, (void**) &arr);
		for (i=0; i<nArray; i++)
			arr[i]	= pArray[i];
		SafeArrayUnaccessData(pVarResult->parray);
		delete [] pArray;
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::AddValue(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANT			varCopy;
	double			NM; 
	double			Signal;
	BOOL			fSuccess = FALSE;
	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	VariantInit(&varCopy);
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[1]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_R8);
	if (FAILED(hr)) return hr;
	NM = varCopy.dblVal;
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[0]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_R8);
	if (FAILED(hr)) return hr;
	Signal = varCopy.dblVal;
	fSuccess = this->m_pBackObj->m_pIDataSet->AddValue(NM, Signal);
	if (NULL != pVarResult)InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::ViewSetup(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	BOOL				fSuccess	= FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	fSuccess = this->m_pBackObj->m_pIDataSet->ViewSetup((HWND) varg.lVal);
	if (!fSuccess)
	{
	}
	if (NULL != pVarResult)
		InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetArraySize(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pBackObj->m_pIDataSet->GetArraySize(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetCurrentExp(
									DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANT			varCopy;
//	LPTSTR			szString		= NULL;
	TCHAR			szFilter[MAX_PATH];
	short int		grating;
	short int		detector;
	if (3 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	VariantInit(&varCopy);
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[2]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_BSTR);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varCopy.bstrVal, &szString);
//	VariantClear(&varCopy);
	StringCchCopy(szFilter, MAX_PATH, varCopy.bstrVal);
	VariantClear(&varCopy);
//	CoTaskMemFree((LPVOID) szString);
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[1]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_I2);
	if (FAILED(hr)) return hr;
	grating = varCopy.iVal;
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[0]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_I2);
	if (FAILED(hr)) return hr;
	detector = varCopy.iVal;
	this->m_pBackObj->m_pIDataSet->SetCurrentExp(szFilter, grating, detector);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::calcRadiance(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANT			varCopy;
	BOOL			fRadiance;
//	long			nArrays;
//	double		*	pWaves		= NULL;
//	double		*	pRadiance	= NULL;
//	SAFEARRAYBOUND	sab;
//	double		*	arrX;
//	double		*	arrY;
//	long			i;
	BOOL			fSuccess = FALSE;

//	if (NULL == pVarResult) return E_INVALIDARG;
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	VariantInit(&varCopy);
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[0]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_BOOL);
	if (FAILED(hr)) return hr;
	// clear current values;
	fRadiance = VARIANT_TRUE == varCopy.boolVal;
	if (NULL != this->m_pRadianceWaves)
	{
		delete[] this->m_pRadianceWaves;
		this->m_pRadianceWaves = NULL;
	}
	if (NULL != this->m_pRadianceValues)
	{
		delete[] this->m_pRadianceValues;
		this->m_pRadianceValues = NULL;
	}
	this->m_nValuesRadiance = 0;
	fSuccess = this->m_pBackObj->m_pIDataSet->calcRadiance(fRadiance, &(this->m_nValuesRadiance), &(this->m_pRadianceWaves),
		&(this->m_pRadianceValues));
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
/*

	if (this->m_pBackObj->m_pIDataSet->calcRadiance(fRadiance, &nArrays, &pWaves, &pRadiance))
	{
		sab.lLbound			= 0;
		sab.cElements		= nArrays;
		*(pDispParams->rgvarg[0].pparray)	= SafeArrayCreate(VT_R8, 1, &sab);
		VariantInit(pVarResult);
		pVarResult->vt		= VT_ARRAY | VT_R8;
		pVarResult->parray	= SafeArrayCreate(VT_R8, 1, &sab);
		SafeArrayAccessData(*(pDispParams->rgvarg[0].pparray), (void**) &arrX);
		SafeArrayAccessData(pVarResult->parray, (void **) &arrY);
		for (i=0; i<nArrays; i++)
		{
			arrX[i]	= pWaves[i];
			arrY[i]	= pRadiance[i];
		}
		SafeArrayUnaccessData(*(pDispParams->rgvarg[0].pparray));
		SafeArrayUnaccessData(pVarResult->parray);
		// cleanup
		delete [] pWaves;
		delete [] pRadiance;
	}
	return S_OK;
*/
}

HRESULT CMyObject::CImpIDispatch::setObject(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	BOOL		fSuccess	= FALSE;
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if (VT_DISPATCH != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	fSuccess = this->m_pBackObj->m_pIDataSet->SetDataSet(pDispParams->rgvarg[0].pdispVal);
	if (NULL != pVarResult)
		InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetAnalysisMode(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt16(this->m_pBackObj->m_pIDataSet->GetAnalysisMode(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetAnalysisMode(
									DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I2, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pIDataSet->SetAnalysisMode(varg.iVal);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::FormOutFile(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	BOOL			fSuccess	= NULL;
	long			ubX;
	long			ubY;
	double		*	arrX;
	double		*	arrY;
//	LPTSTR			szString	= NULL;
	TCHAR			szCalibration[MAX_PATH];
	TCHAR			szFileExt[MAX_PATH];
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	if (4 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_ARRAY | VT_R8) != pDispParams->rgvarg[1].vt) return DISP_E_TYPEMISMATCH;
	if ((VT_BYREF | VT_ARRAY | VT_R8) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szString);
//	VariantClear(&varg);
	StringCchCopy(szCalibration, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	hr = DispGetParam(pDispParams, 1, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(szFileExt, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
//	CoTaskMemFree((LPVOID) szString);
	if (NULL == *(pDispParams->rgvarg[1].pparray)) return E_INVALIDARG;
	if (NULL == *(pDispParams->rgvarg[0].pparray)) return E_INVALIDARG;
	SafeArrayGetUBound(*(pDispParams->rgvarg[1].pparray), 1, &ubX);
	SafeArrayGetUBound(*(pDispParams->rgvarg[0].pparray), 1, &ubY);
	if (ubX == ubY)
	{
		SafeArrayAccessData(*(pDispParams->rgvarg[1].pparray), (void**) &arrX);
		SafeArrayAccessData(*(pDispParams->rgvarg[0].pparray), (void**) &arrY);
		fSuccess = this->m_pBackObj->m_pIDataSet->FormOutFile(
			szCalibration,
			szFileExt,
			ubX + 1, arrX, arrY);
		SafeArrayUnaccessData(*(pDispParams->rgvarg[1].pparray));
		SafeArrayUnaccessData(*(pDispParams->rgvarg[0].pparray));
	}
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetFileName(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	TCHAR			szFileName[MAX_PATH];
	if (this->m_pBackObj->m_pIDataSet->GetFileName(szFileName, MAX_PATH))
		InitVariantFromString(szFileName, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetFileName(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
//	LPTSTR				szFileName		= NULL;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szFileName);
//	VariantClear(&varg);
	this->m_pBackObj->m_pIDataSet->SetFileName(varg.bstrVal);
	VariantClear(&varg);

	VARIANT			varResult;
	this->GetFileName(&varResult);
	VariantClear(&varResult);
//	CoTaskMemFree((LPVOID) szFileName);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetAmCalibration(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pIDataSet->GetAmCalibration(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetAmCalibration(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pIDataSet->SetAmCalibration(VARIANT_TRUE == varg.boolVal);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetCreateTime(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	this->m_pBackObj->m_pIDataSet->GetCreateTime(pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetDataValue(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANT				varCopy;

	if (NULL == pVarResult) return E_INVALIDARG;
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	VariantInit(&varCopy);
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[0]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_R8);
	if (FAILED(hr)) return hr;
	InitVariantFromDouble(this->m_pBackObj->m_pIDataSet->GetDataValue(varCopy.dblVal), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetDataValue(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANT				varCopy;
	double				xValue;
	double				yValue;

	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	VariantInit(&varCopy);
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[1]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_R8);
	if (FAILED(hr)) return hr;
	xValue = varCopy.dblVal;
	hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[0]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varCopy, &varCopy, 0, VT_R8);
	if (FAILED(hr)) return hr;
	yValue = varCopy.dblVal;
	this->m_pBackObj->m_pIDataSet->SetDataValue(xValue, yValue);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::LoadFromString(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	BOOL				fSuccess = FALSE;
	IDispatch		*	pdispDataSet = NULL;
	IDispatch		*	pdisp		= NULL;
	DISPID				dispid;
	VARIANT				varResult;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	Utils_SetStringProperty(this, DISPID_DataProgID, L"Sciencetech.OPTFile.1");		// need to be OPTfile
	if (this->m_pBackObj->m_pIDataSet->GetDataSet(&pdispDataSet))
	{
		hr = pdispDataSet->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);			// want default interface
		pdispDataSet->Release();
	}
	else hr = E_FAIL;
	if (SUCCEEDED(hr))
	{
		Utils_GetMemid(pdisp, L"LoadFromString", &dispid);
		hr = Utils_InvokeMethod(pdisp, dispid, &varg, 1, &varResult);
		if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
		if (fSuccess)
		{
			// clear the read only flag
			Utils_GetMemid(pdisp, L"ClearReadonly", &dispid);
			Utils_InvokeMethod(pdisp, dispid, NULL, 0, NULL);
		}
		pdisp->Release();
	}
	VariantClear(&varg);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetRadianceWaves(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDoubleArray(this->m_pRadianceWaves, this->m_nValuesRadiance, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetRadiance(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDoubleArray(this->m_pRadianceValues, this->m_nValuesRadiance, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetHaveOPT(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(NULL != this->m_pValuesOPT, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetHaveOPT(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	// Optical transfer data
	if (NULL != this->m_pWavesOPT)
	{
		delete[] this->m_pWavesOPT;
		this->m_pWavesOPT = NULL;
	}
	if (NULL != this->m_pValuesOPT)
	{
		delete[] this->m_pValuesOPT;
		this->m_pValuesOPT = NULL;
	}
	this->m_nValuesOPT = 0;
	if (VARIANT_TRUE == varg.boolVal)
	{
		this->m_pBackObj->m_pIDataSet->CalculateOPT(&(this->m_nValuesOPT), &(this->m_pWavesOPT), &(this->m_pValuesOPT));
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetOPTWaves(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDoubleArray(this->m_pWavesOPT, this->m_nValuesOPT, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetOPTValues(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDoubleArray(this->m_pValuesOPT, this->m_nValuesOPT, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::FormOpticalTransferFile(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	BOOL			fSuccess = FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	fSuccess = this->m_pBackObj->m_pIDataSet->FormOpticalTransferFile(varg.bstrVal);
	VariantClear(&varg);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

/*
HRESULT CMyObject::CImpIDispatch::SetCalibrationWavelengths(
	DISPPARAMS	*	pDispParams)
{
	long			ub;

	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_R8 | VT_ARRAY) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	SafeArrayGetUBound(*(pDispParams->rgvarg[0].pparray), 1, &ub);
	return S_OK;
}
*/

HRESULT CMyObject::CImpIDispatch::SetCalibrationData(
					DISPPARAMS		*	pDispParams,
					VARIANT			*	pVarResult)
{
	long			ubx, uby;
	double		*	arrX;
	double		*	arrY;
	BOOL			fSuccess = FALSE;

	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_R8 | VT_ARRAY) != pDispParams->rgvarg[0].vt	||
		(VT_BYREF | VT_R8 | VT_ARRAY) != pDispParams->rgvarg[1].vt) return DISP_E_TYPEMISMATCH;
	SafeArrayGetUBound(*(pDispParams->rgvarg[1].pparray), 1, &ubx);
	SafeArrayGetUBound(*(pDispParams->rgvarg[0].pparray), 1, &uby);
	if (ubx == uby)
	{
		fSuccess = TRUE;
		SafeArrayAccessData(*(pDispParams->rgvarg[1].pparray), (void**)&arrX);
		SafeArrayAccessData(*(pDispParams->rgvarg[0].pparray), (void**)&arrY);
		this->m_pBackObj->SetCalibrationData(ubx + 1, arrX, arrY);
		SafeArrayUnaccessData(*(pDispParams->rgvarg[1].pparray));
		SafeArrayUnaccessData(*(pDispParams->rgvarg[0].pparray));
	}
	if (!fSuccess) this->m_pBackObj->SetCalibrationData(0, NULL, NULL);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::FileNameFromPath(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	TCHAR				szFilePath[MAX_PATH];
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(szFilePath, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
//	// remove the extension
//	PathRemoveExtension(szFilePath);
	// return the file name part
	InitVariantFromString(PathFindFileName(szFilePath), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::NewObjectSetup(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	CDlgNewObjectSetup	dlg(this->m_pBackObj);
	BOOL				fSuccess = FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	fSuccess = dlg.DoOpenDialog(GetServer()->GetInstance(), (HWND)varg.lVal);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	// ask client to update the setup
	TCHAR			szProgID[MAX_PATH];
	if (this->m_pBackObj->m_pIDataSet->GetDataSetProgID(szProgID, MAX_PATH) &&
		0 == lstrcmpi(szProgID, L"Sciencetech.OPTFile.1"))
	{
		IDispatch	*	pdispDataSet;
		IDispatch	*	pdisp;
		DISPID			dispid;
		VARIANT			varResult;
		LPTSTR			szString = NULL;

		if (this->m_pBackObj->m_pIDataSet->GetDataSet(&pdispDataSet))
		{
			hr = pdispDataSet->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
			if (SUCCEEDED(hr))
			{
				// set the time stamp
				Utils_GetMemid(pdisp, L"SetTimeStamp", &dispid);
				Utils_InvokeMethod(pdisp, dispid, NULL, 0, NULL);
				Utils_GetMemid(pdisp, L"SaveToString", &dispid);
				hr = Utils_InvokeMethod(pdisp, dispid, NULL, 0, &varResult);
				if (SUCCEEDED(hr))
				{
					hr = VariantToStringAlloc(varResult, &szString);
					if (SUCCEEDED(hr))
					{
						this->m_pBackObj->FireupdateOptFile(szString);
						CoTaskMemFree((LPVOID)szString);
					}
					VariantClear(&varResult);
				}
				pdisp->Release();
			}
			pdispDataSet->Release();
		}
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetScanType(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromString(this->m_ScanType, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetScanType(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	TCHAR				szProgID[MAX_PATH];
	IDispatch		*	pdispDataSet;
	IDispatch		*	pdisp;
	DISPID				dispid;

	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(this->m_ScanType, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	// set am calibration flag
	if (this->m_pBackObj->m_pIDataSet->GetDataSetProgID(szProgID, MAX_PATH) &&
		0 == lstrcmpi(szProgID, L"Sciencetech.OPTFile.1"))
	{
		if (this->m_pBackObj->m_pIDataSet->GetDataSet(&pdispDataSet))
		{
			hr = pdispDataSet->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
			if (SUCCEEDED(hr))
			{
				Utils_GetMemid(pdisp, L"CalibrationMeasurement", &dispid);
				Utils_SetBoolProperty(pdisp, dispid, 0 == lstrcmpi(this->m_ScanType, L"Intensity Calibration"));
				pdisp->Release();
			}
			pdispDataSet->Release();
		}
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SelectCalibrationStandard(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	HWND			hwndParent = NULL;
	BOOL			fSuccess = FALSE;
	TCHAR			szProgID[MAX_PATH];
	IDispatch	*	pdispDataSet;
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANTARG		avarg[2];
	VARIANT			varResult;
	
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	hwndParent = (HWND)varg.lVal;
	if (this->m_pBackObj->m_pIDataSet->GetDataSetProgID(szProgID, MAX_PATH) &&
		0 == lstrcmpi(szProgID, L"Sciencetech.OPTFile.1"))
	{
		if (this->m_pBackObj->m_pIDataSet->GetDataSet(&pdispDataSet))
		{
			hr = pdispDataSet->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
			if (SUCCEEDED(hr))
			{
				Utils_GetMemid(pdisp, L"Setup", &dispid);
				InitVariantFromInt32((long)hwndParent, &avarg[1]);
				InitVariantFromString(L"Calibration Standard", &avarg[0]);
				hr = Utils_InvokeMethod(pdisp, dispid, avarg, 2, &varResult);
				VariantClear(&avarg[0]);
				if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
				pdisp->Release();
			}
			pdispDataSet->Release();
		}
	}
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetHaveCalibrationStandard(
	VARIANT		*	pVarResult)
{
	TCHAR			szProgID[MAX_PATH];
	HRESULT			hr;
	IDispatch	*	pdispDataSet	= NULL;
	IDispatch	*	pdisp			= NULL;
	IDispatch	*	pdispCalibrationStandard = NULL;
	BOOL			haveCalibration = FALSE;
	DISPID			dispid;
	VARIANT			varResult;
	DISPPARAMS		dispParams;

	if (NULL == pVarResult) return E_INVALIDARG;
	if (this->m_pBackObj->m_pIDataSet->GetDataSetProgID(szProgID, MAX_PATH) && 0 == lstrcmpi(szProgID, L"Sciencetech.OPTFile.1"))
	{
		if (this->m_pBackObj->m_pIDataSet->GetDataSet(&pdispDataSet))
		{
			hr = pdispDataSet->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
			pdispDataSet->Release();
		}
		if (NULL != pdisp)
		{
			Utils_GetMemid(pdisp, L"CalibrationStandard", &dispid);
			// want the dispatch object - can't use Utils_InvokePropertyGet as it initializes varResult;
			VariantInit(&varResult);
			varResult.vt = VT_DISPATCH;
			dispParams.cArgs = 0;
			dispParams.cNamedArgs = 0;
			dispParams.rgdispidNamedArgs = NULL;
			dispParams.rgvarg = NULL;
			hr = pdisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &varResult, NULL, NULL);
			if (SUCCEEDED(hr))
			{
				if (VT_DISPATCH == varResult.vt && NULL != varResult.pdispVal)
				{
					pdispCalibrationStandard = varResult.pdispVal;
					varResult.pdispVal->AddRef();
				}
				VariantClear(&varResult);
			}
			pdisp->Release();
		}
		if (NULL != pdispCalibrationStandard)
		{
			Utils_GetMemid(pdispCalibrationStandard, L"amLoaded", &dispid);
			haveCalibration = Utils_GetBoolProperty(pdispCalibrationStandard, dispid);
			pdispCalibrationStandard->Release();
		}
	}
	InitVariantFromBoolean(haveCalibration, pVarResult);
	return S_OK;
}


HRESULT CMyObject::CImpIDispatch::GetComment(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	TCHAR			szComment[MAX_PATH];
	if (this->m_pBackObj->m_pIDataSet->GetComment(szComment, MAX_PATH))
	{
		InitVariantFromString(szComment, pVarResult);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetComment(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pIDataSet->SetComment(varg.bstrVal);
	VariantClear(&varg);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::ExtraScanData(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	TCHAR			szTitle[MAX_PATH];
	double			nm;
	double			extraValue;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(szTitle, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	hr = DispGetParam(pDispParams, 1, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	nm = varg.dblVal;
	hr = DispGetParam(pDispParams, 2, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	extraValue = varg.dblVal;
	this->m_pBackObj->m_pIDataSet->ExtraScanData(szTitle, nm, extraValue);
	return hr;
}

HRESULT CMyObject::CImpIDispatch::GetDefaultUnits(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	TCHAR			szUnits[MAX_PATH];
	if (this->m_pBackObj->m_pIDataSet->GetDefaultUnits(szUnits, MAX_PATH))
	{
		InitVariantFromString(szUnits, pVarResult);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetDefaultUnits(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pIDataSet->SetDefaultUnits(varg.bstrVal);
	VariantClear(&varg);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetDefaultTitle(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	TCHAR			szTitle[MAX_PATH];
	if (!this->m_pBackObj->m_pIDataSet->GetDefaultTitle(szTitle, MAX_PATH))
	{
		StringCchCopy(szTitle, MAX_PATH, L"Voltage");
	}
	InitVariantFromString(szTitle, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetDefaultTitle(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pIDataSet->SetDefaultTitle(varg.bstrVal);
	VariantClear(&varg);
	return S_OK;
}


HRESULT CMyObject::CImpIDispatch::GetCountsAvailable(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pIDataSet->GetCountsAvailable(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetCountsAvailable(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pIDataSet->SetCountsAvailable(VARIANT_TRUE == varg.boolVal);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::ExportToCSV(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	BOOL			fSuccess = FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	fSuccess = this->m_pBackObj->m_pIDataSet->ExportToCSV(varg.bstrVal);
	VariantClear(&varg);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetADGainFactor(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	return this->m_pBackObj->m_pIDataSet->GetADGainFactor(pDispParams, pVarResult);
}

HRESULT CMyObject::CImpIDispatch::GetArrayScanning(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pIDataSet->GetArrayScanning(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetArrayScanning(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pIDataSet->SetArrayScanning(VARIANT_TRUE == varg.boolVal);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetArrayData(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	long			ubx, uby;
	double		*	arrX;
	double		*	arrY;
	BOOL			fSuccess = FALSE;
//	CSciOutFile		outFile;
//	IDispatch	*	pdisp;

	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_R8 | VT_ARRAY) != pDispParams->rgvarg[0].vt ||
		(VT_BYREF | VT_R8 | VT_ARRAY) != pDispParams->rgvarg[1].vt) return DISP_E_TYPEMISMATCH;
	if (this->m_pBackObj->m_pIDataSet->GetArrayScanning())
	{
		SafeArrayGetUBound(*(pDispParams->rgvarg[1].pparray), 1, &ubx);
		SafeArrayGetUBound(*(pDispParams->rgvarg[0].pparray), 1, &uby);
		if (ubx == uby)
		{
//			fSuccess = TRUE;
			SafeArrayAccessData(*(pDispParams->rgvarg[1].pparray), (void**)&arrX);
			SafeArrayAccessData(*(pDispParams->rgvarg[0].pparray), (void**)&arrY);
			if (ubx == uby)
			{
				fSuccess = this->m_pBackObj->m_pIDataSet->SetArrayData(ubx + 1, arrX, arrY);
			}
//			outFile.SetxValues(ubx + 1, arrX);
	//		outFile.SetyValues(ubx + 1, arrY);
		//	if (outFile.GetMyObject(&pdisp))
			//{
				//fSuccess = this->m_pBackObj->m_pIDataSet->SetDataSet(pdisp);
//				pdisp->Release();
//			}
//			this->m_pBackObj->SetCalibrationData(ubx + 1, arrX, arrY);
			SafeArrayUnaccessData(*(pDispParams->rgvarg[1].pparray));
			SafeArrayUnaccessData(*(pDispParams->rgvarg[0].pparray));
		}
//		if (!fSuccess) this->m_pBackObj->SetCalibrationData(0, NULL, NULL);
	}
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetNumPeaks(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_NOTIMPL;
	if (NULL != this->m_pBackObj->m_pGaussianFit)
		InitVariantFromInt32(this->m_pBackObj->m_pGaussianFit->GetNumPeaks(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::FindPeaks(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	BOOL			fSuccess = FALSE;
	long			nX;
	double		*	pX;
	long			nY;
	double		*	pY;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	Utils_DELETE_POINTER(this->m_pBackObj->m_pGaussianFit);
	if (this->m_pBackObj->m_pIDataSet->GetWavelengths(&nX, &pX))
	{
		if (this->m_pBackObj->m_pIDataSet->GetSpectra(&nY, &pY))
		{
			if (nX == nY)
			{
				this->m_pBackObj->m_pGaussianFit = new CGaussianFit();
				this->m_pBackObj->m_pGaussianFit->SetPeakThreshold(varg.dblVal);
				this->m_pBackObj->m_pGaussianFit->SetData(nX, pX, pY);
				fSuccess = this->m_pBackObj->m_pGaussianFit->FitPeaks();
			}
			delete[] pY;
		}
		delete[] pX;
	}
	if (NULL != pVarResult)
		InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::DisplayPeaks(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	if (NULL != this->m_pBackObj->m_pGaussianFit)
	{
		this->m_pBackObj->m_pGaussianFit->DisplayPeaks((HWND)varg.lVal);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetPeakPosition(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	if (NULL == pVarResult) return E_INVALIDARG;
	if (NULL != this->m_pBackObj->m_pGaussianFit)
	{
		InitVariantFromDouble(this->m_pBackObj->m_pGaussianFit->GetPeakPosition(varg.lVal), pVarResult);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetPeakHeight(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	if (NULL == pVarResult) return E_INVALIDARG;
	if (NULL != this->m_pBackObj->m_pGaussianFit)
	{
		InitVariantFromDouble(this->m_pBackObj->m_pGaussianFit->GetPeakHeight(varg.lVal), pVarResult);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetPeakWidth(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	if (NULL == pVarResult) return E_INVALIDARG;
	if (NULL != this->m_pBackObj->m_pGaussianFit)
	{
		InitVariantFromDouble(this->m_pBackObj->m_pGaussianFit->GetPeakWidth(varg.lVal), pVarResult);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::IntegrateSpectra(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	long		ubx;
	long		uby;
	double	*	arrX;
	double	*	arrY;
	long		i;
	BOOL		fSuccess = FALSE;
	double		sum;

	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_R8 | VT_ARRAY) != pDispParams->rgvarg[1].vt) return DISP_E_TYPEMISMATCH;
	if ((VT_BYREF | VT_R8 | VT_ARRAY) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	SafeArrayGetLBound(*(pDispParams->rgvarg[1].pparray), 1, &ubx);
	SafeArrayGetLBound(*(pDispParams->rgvarg[0].pparray), 1, &uby);
	if (ubx == uby)
	{
		SafeArrayAccessData(*(pDispParams->rgvarg[1].pparray), (void**)&arrX);
		SafeArrayAccessData(*(pDispParams->rgvarg[0].pparray), (void**)&arrY);
		sum = 0.0;
		for (i = 0; i <= uby; i++)
		{
			sum += arrY[i];
			arrY[i] = sum;
		}
		SafeArrayUnaccessData(*(pDispParams->rgvarg[1].pparray));
		SafeArrayUnaccessData(*(pDispParams->rgvarg[0].pparray));
		fSuccess = TRUE;
	}
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);  
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SmoothSpectra(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	BOOL			fGaussian = FALSE;		// Gaussian smoothing flag
	short int		numPoints;				// number of points used = 3, 5, or 7
	long			ubx;
	long			uby;
	double		*	arrX;
	double		*	arrY;
	BOOL			fSuccess = FALSE;
	double		*	yCopy;
	long			i;

	if (4 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_ARRAY | VT_R8) != pDispParams->rgvarg[3].vt) return DISP_E_TYPEMISMATCH;
	if ((VT_BYREF | VT_ARRAY | VT_R8) != pDispParams->rgvarg[2].vt) return DISP_E_TYPEMISMATCH;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 2, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	fGaussian = VARIANT_TRUE == varg.boolVal;
	hr = DispGetParam(pDispParams, 3, VT_I2, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	numPoints = varg.iVal;
	if (numPoints != 3 && numPoints != 5 && numPoints != 7) return E_INVALIDARG;
	SafeArrayGetUBound(*(pDispParams->rgvarg[3].pparray), 1, &ubx);
	SafeArrayGetUBound(*(pDispParams->rgvarg[2].pparray), 1, &uby);
	if (ubx == uby)
	{
		fSuccess = TRUE;
		SafeArrayAccessData(*(pDispParams->rgvarg[3].pparray), (void**)&arrX);
		SafeArrayAccessData(*(pDispParams->rgvarg[2].pparray), (void**)&arrY);
		// make scratch copy
		yCopy = new double[uby + 1];
		for (i = 0; i <= uby; i++)
		{
			yCopy[i] = arrY[i];
		}
		// copy first and last
		arrY[0] = yCopy[0];
		arrY[uby] = yCopy[uby];
		if (fGaussian)
		{
			// gaussian smoothing
			if (3 == numPoints)
			{
				for (i = 1; i < uby; i++)
				{
					arrY[i] = (yCopy[i - 1] + (2.0 * yCopy[i]) + yCopy[i + 1]) / 4.0;
				}
			}
			else if (5 == numPoints)
			{
				arrY[1] = yCopy[1];
				arrY[uby - 1] = yCopy[uby - 1];
				for (i = 2; i < (uby - 1); i++)
				{
					arrY[i] = (yCopy[i - 2] + (4.0 * yCopy[i - 1]) + (6.0 * yCopy[i]) + (4.0 * yCopy[i + 1]) + yCopy[i + 2]) / 16.0;
				}
			}
			else if (7 == numPoints)
			{
				arrY[1] = yCopy[1];
				arrY[2] = yCopy[2];
				arrY[uby - 2] = yCopy[uby - 2];
				arrY[uby - 1] = yCopy[uby - 1];
				for (i = 3; i < (uby - 2); i++)
				{
					arrY[i] = (yCopy[i - 3] + (6.0 * yCopy[i - 2]) + (15.0 * yCopy[i - 1]) + (20.0 * yCopy[i]) + (15.0 * yCopy[i + 1]) + (6.0 * yCopy[i + 2]) + yCopy[i + 3]) / 64.0;
				}
			}
		}
		else
		{
			// boxcar smoothing
			if (3 == numPoints)
			{
				for (i = 1; i < uby; i++)
				{
					arrY[i] = (yCopy[i - 1] + yCopy[i] + yCopy[i + 1]) / 3.0;
				}
			}
			else if (5 == numPoints)
			{
				arrY[1] = yCopy[1];
				arrY[uby - 1] = yCopy[uby - 1];
				for (i = 2; i < (uby - 1); i++)
				{
					arrY[i] = (yCopy[i - 2] + yCopy[i - 1] + yCopy[i] + yCopy[i + 1] + yCopy[i + 2]) / 5.0;
				}

			}
			else if (7 == numPoints)
			{
				arrY[1] = yCopy[1];
				arrY[2] = yCopy[2];
				arrY[uby - 2] = yCopy[uby - 2];
				arrY[uby - 1] = yCopy[uby - 1];
				for (i = 3; i < (uby - 2); i++)
				{
					arrY[i] = (yCopy[i - 3] + yCopy[i - 2]+ yCopy[i - 1] + yCopy[i] + yCopy[i + 1] + yCopy[i + 2] + yCopy[i + 3]) / 7.0;
				}
			}
		}
		// cleanup
		delete[] yCopy;
		SafeArrayUnaccessData(*(pDispParams->rgvarg[3].pparray));
		SafeArrayUnaccessData(*(pDispParams->rgvarg[2].pparray));
	}
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::CalcAverage(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	long			nY;
	double		*	pY;
	double			average = 0.0;
	double			StanDev = 0.0;
	double			sum = 0.0;
	double			sumSq = 0.0;
	long			i;
	long			nAverage = 0;
	double			temp;

	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_R8) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	if (this->m_pBackObj->m_pIDataSet->GetWavelengths(&nY, &pY))
	{
		for (i = 0; i < nY; i++)
		{
			sum += pY[i];
			sumSq += (pY[i] * pY[i]);
			nAverage++;
		}
		if (nAverage > 0)
		{
			average = sum / nAverage;
			temp = sumSq - ((sum * sum) / nAverage);
			if (temp > 0.0)
			{
				StanDev = sqrt(temp / nAverage);
			}
			else
			{
				StanDev = 0.0;
			}
		}
		delete[] pY;
	}
	*(pDispParams->rgvarg[0].pdblVal) = StanDev;
	InitVariantFromDouble(average, pVarResult);
	return S_OK;
}


CMyObject::CImpIProvideClassInfo2::CImpIProvideClassInfo2(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIProvideClassInfo2::~CImpIProvideClassInfo2()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIProvideClassInfo2::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIProvideClassInfo2::Release()
{
	return this->m_punkOuter->Release();
}

// IProvideClassInfo method
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::GetClassInfo(
									ITypeInfo	**	ppTI) 
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;
	*ppTI		= NULL;
	hr = GetServer()->GetTypeLib(&pTypeLib);
	if (SUCCEEDED(hr))
	{
		hr = pTypeLib->GetTypeInfoOfGuid(MY_CLSID, ppTI);
		pTypeLib->Release();
	}
	return hr;
}

// IProvideClassInfo2 method
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::GetGUID(
									DWORD			dwGuidKind,  //Desired GUID
									GUID		*	pGUID)
{
	if (GUIDKIND_DEFAULT_SOURCE_DISP_IID == dwGuidKind)
	{
		*pGUID	= MY_IIDSINK;
		return S_OK;
	}
	else
	{
		*pGUID	= GUID_NULL;
		return E_INVALIDARG;	
	}
}

CMyObject::CImpIConnectionPointContainer::CImpIConnectionPointContainer(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIConnectionPointContainer::~CImpIConnectionPointContainer()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIConnectionPointContainer::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIConnectionPointContainer::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIConnectionPointContainer::Release()
{
	return this->m_punkOuter->Release();
}

// IConnectionPointContainer methods
STDMETHODIMP CMyObject::CImpIConnectionPointContainer::EnumConnectionPoints(
									IEnumConnectionPoints **ppEnum)
{
	return Utils_CreateEnumConnectionPoints(this, MAX_CONN_PTS, this->m_pBackObj->m_paConnPts,
		ppEnum);
}

STDMETHODIMP CMyObject::CImpIConnectionPointContainer::FindConnectionPoint(
									REFIID			riid,  //Requested connection point's interface identifier
									IConnectionPoint **ppCP)
{
	HRESULT					hr		= CONNECT_E_NOCONNECTION;
	IConnectionPoint	*	pConnPt	= NULL;
	*ppCP	= NULL;
	if (MY_IIDSINK == riid)
		pConnPt	= this->m_pBackObj->m_paConnPts[CONN_PT_CUSTOMSINK];
	if (NULL != pConnPt)
	{
		*ppCP = pConnPt;
		pConnPt->AddRef();
		hr = S_OK;
	}
	return hr;
}
