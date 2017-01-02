#include "stdafx.h"
#include "IDataSet.h"
#include "MyObject.h"
#include "SciOutFile.h"
#include "dispids.h"
#include "TempFile.h"

CIDataSet::CIDataSet(CMyObject * pMyObject) :
	m_pMyObject(pMyObject),
	m_pdisp(NULL),
	m_iidSink(IID_NULL),
	m_dwCookie(0),
	// file to store readings in case of crash
	m_pTempFile(NULL)
{
}

CIDataSet::~CIDataSet(void)
{
	this->SetDataSet(NULL);
	Utils_DELETE_POINTER(this->m_pTempFile);
}

BOOL CIDataSet::SetDataSet(
						IDispatch	*	pdisp)
{
	HRESULT				hr;
	IProvideClassInfo*	pProvideClassInfo;
	ITypeInfo		*	pClassInfo = NULL;;
	ITypeInfo		*	pTypeInfo = NULL;
	TYPEATTR		*	pTypeAttr;
	CImpISink		*	pSink;
	IUnknown		*	pUnkSink;
	BOOL				fSuccess	= FALSE;

	if (NULL != this->m_pdisp)
	{
		Utils_ConnectToConnectionPoint(this->m_pdisp, NULL, this->m_iidSink, FALSE,
			&(this->m_dwCookie));
		Utils_RELEASE_INTERFACE(this->m_pdisp);
	}
	if (NULL == pdisp) return FALSE;
	hr = pdisp->QueryInterface(IID_IProvideClassInfo, (LPVOID*) &pProvideClassInfo);
	if (SUCCEEDED(hr))
	{
		hr = pProvideClassInfo->GetClassInfo(&pClassInfo);
		pProvideClassInfo->Release();
	}
	if (SUCCEEDED(hr))
	{
		hr = Utils_FindImplClassName(pClassInfo, TEXT("_clsIDataSet"), &pTypeInfo) ? S_OK : E_FAIL;
		pClassInfo->Release();
	}
	if (SUCCEEDED(hr))
	{
		hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			hr = pdisp->QueryInterface(pTypeAttr->guid, (LPVOID*) &(this->m_pdisp));
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		pTypeInfo->Release();
		fSuccess = SUCCEEDED(hr);
	}
	if (SUCCEEDED(hr))
	{
		pSink	= new CImpISink(this);
		hr = pSink->QueryInterface(this->m_iidSink, (LPVOID*) &pUnkSink);
		if (SUCCEEDED(hr))
		{
			hr = Utils_ConnectToConnectionPoint(this->m_pdisp, pUnkSink, this->m_iidSink,
				TRUE, &(this->m_dwCookie));
			pUnkSink->Release();
		}
		else
			delete pSink;
	}
	return fSuccess;
}

BOOL CIDataSet::GetDataSet(
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

BOOL CIDataSet::GetRadianceAvailable()
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("RadianceAvailable"), &dispid);
	return Utils_GetBoolProperty(this->m_pdisp, dispid);
}

BOOL CIDataSet::GetIrradianceAvailable()
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("IrradianceAvailable"), &dispid);
	return Utils_GetBoolProperty(this->m_pdisp, dispid);
}

short int CIDataSet::GetAnalysisMode()
{
	DISPID			dispid;
	HRESULT			hr;
	VARIANT			varResult;
	short int		mode	= -1;
	Utils_GetMemid(this->m_pdisp, TEXT("AnalysisMode"), &dispid);
	hr = Utils_InvokePropertyGet(this->m_pdisp, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr)) VariantToInt16(varResult, &mode);
	return mode;
}

void CIDataSet::SetAnalysisMode(
						short int		AnalysisMode)
{
	DISPID			dispid;
	VARIANTARG		varg;
	Utils_GetMemid(this->m_pdisp, TEXT("AnalysisMode"), &dispid);
	InitVariantFromInt16(AnalysisMode, &varg);
	Utils_InvokePropertyPut(this->m_pdisp, dispid, &varg, 1);
}

BOOL CIDataSet::ReadDataFile()
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANT				varResult;
	BOOL				fSuccess	= FALSE;
	Utils_GetMemid(this->m_pdisp, TEXT("ReadDataFile"), &dispid);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

BOOL CIDataSet::WriteDataFile()
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANT				varResult;
	BOOL				fSuccess	= FALSE;
	this->CheckFileExtension();			// check the file extension for saving
	Utils_GetMemid(this->m_pdisp, TEXT("WriteDataFile"), &dispid);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess); 
	return fSuccess;
}

BOOL CIDataSet::GetWavelengths(
						long		*	nArray,
						double		**	ppArray)
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANT				varResult;
	long				ub;
	long				i;
	double			*	arr;
	BOOL				fSuccess		= FALSE;
	*nArray		= 0;
	*ppArray	= NULL;
	Utils_GetMemid(this->m_pdisp, TEXT("GetWavelengths"), &dispid);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr))
	{
		if ((VT_ARRAY | VT_R8) == varResult.vt && NULL != varResult.parray)
		{
			fSuccess		= TRUE;
			SafeArrayGetUBound(varResult.parray, 1, &ub);
			*nArray		= this->m_pMyObject->GetSkipLastPoint() ? ub : ub + 1;
			*ppArray	= new double [*nArray];
			SafeArrayAccessData(varResult.parray, (void**) &arr);
			for (i=0; i<(*nArray); i++)
				(*ppArray)[i]	= arr[i];
			SafeArrayUnaccessData(varResult.parray);
		}
		VariantClear(&varResult);
	}
	return fSuccess;
}

BOOL CIDataSet::GetSpectra(
						long		*	nArray,
						double		**	ppArray)
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANT				varResult;
	long				ub;
	long				i;
	double			*	arr;
	BOOL				fSuccess		= FALSE;
	*nArray		= 0;
	*ppArray	= NULL;
	Utils_GetMemid(this->m_pdisp, TEXT("GetSpectra"), &dispid);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr))
	{
		if ((VT_ARRAY | VT_R8) == varResult.vt && NULL != varResult.parray)
		{
			fSuccess		= TRUE;
			SafeArrayGetUBound(varResult.parray, 1, &ub);
			*nArray = this->m_pMyObject->GetSkipLastPoint() ? ub : ub + 1;
			*ppArray	= new double [*nArray];
			SafeArrayAccessData(varResult.parray, (void**) &arr);
			for (i=0; i<(*nArray); i++)
				(*ppArray)[i]	= arr[i];
			SafeArrayUnaccessData(varResult.parray);
		}
		VariantClear(&varResult);
	}
	return fSuccess;
}

BOOL CIDataSet::AddValue(                      
						double			NM, 
						double			Signal)
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			avarg[2];
	VARIANT				varResult;
	BOOL				fSuccess	= FALSE;
	Utils_GetMemid(this->m_pdisp, TEXT("AddValue"), &dispid);
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_BYREF | VT_R8;
	avarg[1].pdblVal	= &NM;
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_R8;
	avarg[0].pdblVal	= &Signal;
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, avarg, 2, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);

	if (NULL != this->m_pTempFile)
	{
		this->m_pTempFile->AddValue(NM, Signal);
	}

	return fSuccess;
}

BOOL CIDataSet::ViewSetup(
						HWND			hwnd)
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			varg;
	VARIANT				varResult;
	BOOL				fSuccess	= FALSE;
	Utils_GetMemid(this->m_pdisp, TEXT("ViewSetup"), &dispid);
	InitVariantFromInt32((long)hwnd, &varg);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

long CIDataSet::GetArraySize()
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANT			varResult;
	long			arraySize		= 0;
	Utils_GetMemid(this->m_pdisp, TEXT("GetArraySize"), &dispid);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr))VariantToInt32(varResult, &arraySize);
	return arraySize;
}

void CIDataSet::SetCurrentExp(
						LPCTSTR			filter, 
						short			grating,
						short			detector)
{
	DISPID				dispid;
	VARIANTARG			avarg[3];
	BSTR				bstr		= NULL;
	Utils_GetMemid(this->m_pdisp, TEXT("SetCurrentExp"), &dispid);
	bstr = SysAllocString(filter);
	VariantInit(&avarg[2]);
	avarg[2].vt			= VT_BYREF | VT_BSTR;
	avarg[2].pbstrVal	= &bstr;
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_BYREF | VT_I2;
	avarg[1].piVal		= &grating;
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_I2;
	avarg[0].piVal		= &detector;
	Utils_InvokeMethod(this->m_pdisp, dispid, avarg, 3, NULL);
	SysFreeString(bstr);

	if (NULL == this->m_pTempFile)
	{
		this->m_pTempFile = new CTempFile();
	}
	if (0 == this->GetArraySize())
	{
		TCHAR			szWorkingDirectory[MAX_PATH];
		if (this->m_pMyObject->FirerequestWorkingDirectory(szWorkingDirectory, MAX_PATH))
		{
			this->m_pTempFile->InitTempFile(szWorkingDirectory);
		}
	}
	this->m_pTempFile->SetCurrentExp(filter, grating, detector);
}

BOOL CIDataSet::calcRadiance(
						BOOL			radiance,
						long		*	nArrays,
						double		**	ppWaves,
						double		**	ppRadiance)
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			avarg[2];
	VARIANT				varResult;
	SAFEARRAY		*	pSA			= NULL;
	long				ubx;
	long				uby;
	double			*	arrX;
	double			*	arrY;
	long				i;
	BOOL				fSuccess		= FALSE;
	VARIANT_BOOL		fvarRadiance	= radiance ? VARIANT_TRUE : VARIANT_FALSE;
	*nArrays	= 0;
	*ppWaves	= NULL;
	*ppRadiance	= NULL;
	Utils_GetMemid(this->m_pdisp, TEXT("calcRadiance"), &dispid);
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_BYREF | VT_BOOL;
	avarg[1].pboolVal	= &fvarRadiance;
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_ARRAY | VT_R8;
	avarg[0].pparray	= &pSA;
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, avarg, 2, &varResult);
	if (SUCCEEDED(hr))
	{
		if (NULL != pSA && (VT_ARRAY | VT_R8) == varResult.vt && NULL != varResult.parray)
		{
			SafeArrayGetUBound(pSA, 1, &ubx);
			SafeArrayGetUBound(varResult.parray, 1, &uby);
			if (ubx == uby)
			{
				fSuccess	= TRUE;
				*nArrays	= ubx+1;
				*ppWaves	= new double [*nArrays];
				*ppRadiance	= new double [*nArrays];
				SafeArrayAccessData(pSA, (void**) &arrX);
				SafeArrayAccessData(varResult.parray, (void**) &arrY);
				for (i=0; i<(*nArrays); i++)
				{
					(*ppWaves)[i]		= arrX[i];
					(*ppRadiance)[i]	= arrY[i];
				}
				SafeArrayUnaccessData(pSA);
				SafeArrayUnaccessData(varResult.parray);
			}
		}
		SafeArrayDestroy(pSA);
		VariantClear(&varResult);
	}
	return fSuccess;
}

HRESULT	CIDataSet::GetCreateTime(
						VARIANT		*	pVarResult)
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("GetCreateTime"), &dispid);
	return Utils_InvokeMethod(this->m_pdisp, dispid, NULL, 0, pVarResult);
}

BOOL CIDataSet::GetFileName(
						LPTSTR			szFileName,
						UINT			nBufferSize)
{
	HRESULT			hr;
	DISPID			dispid;
	LPTSTR			szString		= NULL;
	BOOL			fSuccess		= FALSE;
	szFileName[0]	= '\0';
	Utils_GetMemid(this->m_pdisp, TEXT("FileName"), &dispid);
	hr = Utils_GetStringProperty(this->m_pdisp, dispid, &szString);
	if (NULL != szString)
	{
		StringCchCopy(szFileName, nBufferSize, szString);
		CoTaskMemFree((LPVOID) szString);
		fSuccess = TRUE;
	}
	return fSuccess;
}

void CIDataSet::SetFileName(
						LPCTSTR			szFileName)
{
	DISPID			dispid;
	// make sure have object
	if (NULL == this->m_pdisp)
	{
		// find the prog id for the file using ScFileOpen utility
		HRESULT			hr;
//		LPOLESTR		ProgID	= NULL;
		CLSID			clsid;
		IDispatch	*	pdispSciFileOpen = NULL;
		DISPID			dispid;
		VARIANTARG		varg;
		VARIANT			varResult;
		IDispatch	*	pdisp;

		hr = CLSIDFromProgID(L"Sciencetech.SciFileOpen.1", &clsid);
		if (SUCCEEDED(hr))
		{
			hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (LPVOID*) &pdispSciFileOpen);
		}
		if (SUCCEEDED(hr))
		{
			Utils_GetMemid(pdispSciFileOpen, TEXT("FileProgID"), &dispid);
			InitVariantFromString(szFileName, &varg);
			hr = Utils_InvokeMethod(pdispSciFileOpen, dispid, &varg, 1, &varResult);
			VariantClear(&varg);
			if (SUCCEEDED(hr))
			{
				hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*) &pdisp);
				if (SUCCEEDED(hr))
				{
					Utils_InvokePropertyPut(pdisp, DISPID_DataProgID, &varResult, 1);
					pdisp->Release();
				}
				VariantClear(&varResult);
			}
			pdispSciFileOpen->Release();
		}
	}

	Utils_GetMemid(this->m_pdisp, TEXT("FileName"), &dispid);
	Utils_SetStringProperty(this->m_pdisp, dispid, szFileName);
	TCHAR			szString[MAX_PATH];
	this->GetFileName(szString, MAX_PATH);
	BOOL			stat	= FALSE;
}

BOOL CIDataSet::GetAmCalibration()
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("amCalibration"), &dispid);
	return Utils_GetBoolProperty(this->m_pdisp, dispid);
}

void CIDataSet::SetAmCalibration(
						BOOL			AmCalibration)
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("amCalibration"), &dispid);
	Utils_SetBoolProperty(this->m_pdisp, dispid, AmCalibration);
}

BOOL CIDataSet::FormOutFile(
						LPCTSTR			szCalibration,
						LPCTSTR			szFileExt,
						long			nArray,
						double		*	pXValues,
						double		*	pYValues)
{
	CSciOutFile			outFile;
	TCHAR				szFileName[MAX_PATH];
	TCHAR				szOutFile[MAX_PATH];
	TCHAR				szAnalysis[MAX_PATH];
	VARIANT				TimeStamp;
	HRESULT				hr;
	TCHAR				szComment[MAX_PATH];
	TCHAR				szDefUnits[MAX_PATH];
	TCHAR				szTitle[MAX_PATH];
	long				nTitles;
	long				index;
	long				nExtra;
	double			*	pWaves;
	double			*	pValues;

	// form the file name
	if (!this->GetFileName(szFileName, MAX_PATH))
		return FALSE;
	StringCchCopy(szOutFile, MAX_PATH, szFileName);
	PathRenameExtension(szOutFile, szFileExt);		//			 TEXT(".out"));
	outFile.SetfileName(szOutFile);
	// set the source file
	outFile.SetsourceFile(szFileName);
	// set the calibration file
	outFile.SetcalibrationFile(szCalibration);
	// set the analysis type
	if (this->m_pMyObject->FirerequestAnalysisModeString(
		this->GetAnalysisMode(), szAnalysis, MAX_PATH))
	{
		outFile.SetanalysisType(szAnalysis);
	}
	// set the y units
	if (this->GetDefaultUnits(szDefUnits, MAX_PATH))
	{
		outFile.SetyUnits(szDefUnits);
	}
	// set the comment
	if (this->GetComment(szComment, MAX_PATH))
	{
		outFile.SetComment(szComment);
	}
	// time stamp
	this->GetCreateTime(&TimeStamp);
	outFile.SettimeStamp(&TimeStamp);
	VariantClear(&TimeStamp);

	if (0 == lstrcmpi(szFileExt, L".out"))
	{
		// add extra data
		nTitles = this->GetNumberExtraScanArrays();
		for (index = 0; index < nTitles; index++)
		{
			if (this->GetExtraScanArray(index, &nExtra, &pWaves, &pValues, szTitle, MAX_PATH))
			{
				outFile.SetExtraData(szTitle, nExtra, pWaves, pValues);
				delete[] pWaves;
				delete[] pValues;
			}
		}
	}
	// set the data
	outFile.SetxValues(nArray, pXValues);
	outFile.SetyValues(nArray, pYValues);
	// set the time stamp
	hr = this->GetCreateTime(&TimeStamp);
	if (SUCCEEDED(hr))
	{
		outFile.SettimeStamp(&TimeStamp);
		VariantClear(&TimeStamp);
	}
	// write the file
	outFile.saveFile();
	return TRUE;
}

double CIDataSet::GetDataValue(
						double			xValue)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANTARG		varg;
	VARIANT			varResult;
	double			dataValue	= 0.0;
	Utils_GetMemid(this->m_pdisp, TEXT("GetDataValue"), &dispid);
	VariantInit(&varg);
	varg.vt			= VT_BYREF | VT_R8;
	varg.pdblVal	= &xValue;
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToDouble(varResult, &dataValue);
	return dataValue;
}

void CIDataSet::SetDataValue(
						double			xValue,
						double			yValue)
{
	DISPID			dispid;
	VARIANTARG		avarg[2];
	Utils_GetMemid(this->m_pdisp, TEXT("SetDataValue"), &dispid);
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_BYREF | VT_R8;
	avarg[1].pdblVal	= &xValue;
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_R8;
	avarg[0].pdblVal	= &yValue;
	Utils_InvokeMethod(this->m_pdisp, dispid, avarg, 2, NULL);
}

// calculate the optical transfer function (if available)
BOOL CIDataSet::CalculateOPT(
	long		*	nValues,
	double		**	ppWaves,
	double		**	ppValues)
{
	if (this->GetAmCalibration())
	{
		return this->calcRadiance(FALSE, nValues, ppWaves, ppValues);
	}
	else
	{
		*nValues = 0;
		*ppWaves = NULL;
		*ppValues = NULL;
		return FALSE;
	}
}

// form the optical transfer function
BOOL CIDataSet::FormOpticalTransferFile(
	LPCTSTR			filePath)
{
	HRESULT			hr;
	ITypeInfo	*	pTypeInfo;
	TYPEATTR	*	pTypeAttr;
	IDispatch	*	pdispNamed		= NULL;
	BOOL			fSuccess = FALSE;
	DISPID			dispid;
	VARIANTARG		varg;
	VARIANT			varResult;

	hr = Utils_GetNamedTypeInfo(this->m_pdisp, L"IOPTFile", &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			hr = this->m_pdisp->QueryInterface(pTypeAttr->guid, (LPVOID*)&pdispNamed);
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		pTypeInfo->Release();
	}
	if (SUCCEEDED(hr))
	{
		Utils_GetMemid(pdispNamed, L"FormOpticalTransferFile", &dispid);
		InitVariantFromString(filePath, &varg);
		hr = Utils_InvokeMethod(pdispNamed, dispid, &varg, 1, &varResult);
		VariantClear(&varg);
		if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
		pdispNamed->Release();
	}
	return fSuccess;
}

// set the AD info string
BOOL CIDataSet::SetADInfoString()
{
	HRESULT			hr;
	TCHAR			szProgID[MAX_PATH];
	DISPID			dispid;
	VARIANTARG		varg;
	LPTSTR			infoString = NULL;
	BOOL			fSuccess = FALSE;
	IDispatch	*	pdisp;

	// check prog id
	if (!this->GetDataSetProgID(szProgID, MAX_PATH) || 0 != lstrcmpi(szProgID, L"Sciencetech.OPTFile.1"))
	{
		return FALSE;
	}
	if (this->m_pMyObject->FirerequestADInfoString(&infoString))
	{
		hr = this->m_pdisp->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
		if (SUCCEEDED(hr))
		{
			fSuccess = TRUE;
			Utils_GetMemid(pdisp, L"ADInfoString", &dispid);
			InitVariantFromString(infoString, &varg);
			Utils_InvokePropertyPut(pdisp, dispid, &varg, 1);
			VariantClear(&varg);
			pdisp->Release();
		}
		CoTaskMemFree((LPVOID) infoString);
	}
	return fSuccess;
}

// comment string
BOOL CIDataSet::GetComment(
	LPTSTR			szComment,
	UINT			nBufferSize)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	szComment[0] = '\0';
	Utils_GetMemid(this->m_pdisp, L"Comment", &dispid);
	hr = Utils_InvokePropertyGet(this->m_pdisp, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr))
	{
		hr = VariantToString(varResult, szComment, nBufferSize);
		fSuccess = SUCCEEDED(hr);
		VariantClear(&varResult);
	}
	return fSuccess;
}

void CIDataSet::SetComment(
	LPCTSTR			szComment)
{
	DISPID			dispid;
	VARIANTARG		varg;
	Utils_GetMemid(this->m_pdisp, L"Comment", &dispid);
	InitVariantFromString(szComment, &varg);
	Utils_InvokePropertyPut(this->m_pdisp, dispid, &varg, 1);
	VariantClear(&varg);
}

// extra scan data
void CIDataSet::ExtraScanData(
	BSTR			title,
	double			wave,
	double			dValue)
{
	DISPID		dispid;
	VARIANTARG	avarg[3];
	Utils_GetMemid(this->m_pdisp, L"ExtraScanData", &dispid);
	InitVariantFromString(title, &avarg[2]);
	InitVariantFromDouble(wave, &avarg[1]);
	InitVariantFromDouble(dValue, &avarg[0]);
	Utils_InvokeMethod(this->m_pdisp, dispid, avarg, 3, NULL);
	VariantClear(&avarg[2]);
}

// default units
BOOL CIDataSet::GetDefaultUnits(
	LPTSTR			szDefaultUnits,
	UINT			nBufferSize)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	Utils_GetMemid(this->m_pdisp, L"DefaultUnits", &dispid);
	hr = Utils_InvokePropertyGet(this->m_pdisp, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr))
	{
		hr = VariantToString(varResult, szDefaultUnits, nBufferSize);
		fSuccess = SUCCEEDED(hr);
		VariantClear(&varResult);
	}
	return fSuccess;
}

void CIDataSet::SetDefaultUnits(
	LPCTSTR			szDefaultUnits)
{
	DISPID			dispid;
	VARIANTARG		varg;
	Utils_GetMemid(this->m_pdisp, L"DefaultUnits", &dispid);
	InitVariantFromString(szDefaultUnits, &varg);
	Utils_InvokePropertyPut(this->m_pdisp, dispid, &varg, 1);
	VariantClear(&varg);
}

BOOL CIDataSet::GetDefaultTitle(
	LPTSTR			szDefaultTitle,
	UINT			nBufferSize)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	Utils_GetMemid(this->m_pdisp, L"DefaultTitle", &dispid);
	hr = Utils_InvokePropertyGet(this->m_pdisp, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr))
	{
		hr = VariantToString(varResult, szDefaultTitle, nBufferSize);
		fSuccess = SUCCEEDED(hr);
		VariantClear(&varResult);
	}
	return fSuccess;
}

void CIDataSet::SetDefaultTitle(
	LPCTSTR			szDefaultTitle)
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, L"DefaultTitle", &dispid);
	Utils_SetStringProperty(this->m_pdisp, dispid, szDefaultTitle);
}

// are counts available?
BOOL CIDataSet::GetCountsAvailable()
{
	return FALSE;
}

void CIDataSet::SetCountsAvailable(
	BOOL			CountsAvailable)
{

}

double CIDataSet::OnrequestDispersion(
						long			grating)
{
	return this->m_pMyObject->FirerequestDispersion(grating);
}

short int CIDataSet::OnrequestAnalysisModeIndex(
						LPCTSTR			szAnalysisMode)
{
	return this->m_pMyObject->FirerequestAnalysisModeIndex(szAnalysisMode);
}

BOOL CIDataSet::OnrequestOpticalTransfer(
						long		*	nValues,
						double		**	ppXValues,
						double		**	ppYValues)
{
	return this->m_pMyObject->FirerequestOpticalTransfer(nValues, ppXValues, ppYValues);
}

BOOL CIDataSet::OnrequestADInfoString(
	LPTSTR		*	szInfoString)
{
	return this->m_pMyObject->FirerequestADInfoString(szInfoString);
}

// check the file name for saving
void CIDataSet::CheckFileExtension()
{
	// determine what the file extension should be
	



}

// data set prog id
BOOL CIDataSet::GetDataSetProgID(
	LPTSTR			szProgID,
	UINT			nBufferSize)
{
	HRESULT				hr;
	IProvideClassInfo*	pProvideClassInfo	= NULL;
	ITypeInfo		*	pClassInfo = NULL;
	TYPEATTR		*	pClassAttr;
	LPOLESTR			pOleStr;
	BOOL				fSuccess = FALSE;

	szProgID[0] = '\0';
	if (NULL == this->m_pdisp) return FALSE;
	hr = this->m_pdisp->QueryInterface(IID_IProvideClassInfo, (LPVOID*)&pProvideClassInfo);
	if (SUCCEEDED(hr))
	{
		hr = pProvideClassInfo->GetClassInfo(&pClassInfo);
		pProvideClassInfo->Release();
	}
	if (SUCCEEDED(hr))
	{
		hr = pClassInfo->GetTypeAttr(&pClassAttr);
		if (SUCCEEDED(hr))
		{
			hr = ProgIDFromCLSID(pClassAttr->guid, &pOleStr);
			if (SUCCEEDED(hr))
			{
				StringCchCopy(szProgID, nBufferSize, pOleStr);
				fSuccess = TRUE;
				CoTaskMemFree((LPVOID)pOleStr);
			}
			pClassInfo->ReleaseTypeAttr(pClassAttr);
		}
		pClassInfo->Release();
	}
	return fSuccess;
}

// extra values
long CIDataSet::GetNumberExtraScanArrays()
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, L"NumberExtraScanArrays", &dispid);
	return Utils_GetLongProperty(this->m_pdisp, dispid);
}
BOOL CIDataSet::GetExtraScanArray(
	long			index,
	long		*	nArray,
	double		**	ppWaves,
	double		**	ppValues,
	LPTSTR			szTitle,
	UINT			nBufferSize)
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			avarg[3];
	VARIANT				varResult;
	BOOL				fSuccess = FALSE;
	long				ubx, uby;
	long				i;
	SAFEARRAY		*	pSAX	= NULL;
	SAFEARRAY		*	pSAY	= NULL;
	double			*	arrX;
	double			*	arrY;
	*nArray = 0;
	*ppWaves = NULL;
	*ppValues = NULL;
	szTitle[0] = '\0';
	Utils_GetMemid(this->m_pdisp, L"GetExtraScanArray", &dispid);
	InitVariantFromInt32(index, &avarg[2]);
	VariantInit(&avarg[1]);
	avarg[1].vt = VT_BYREF | VT_ARRAY | VT_R8;
	avarg[1].pparray = &pSAX;
	VariantInit(&avarg[0]);
	avarg[0].vt = VT_BYREF | VT_ARRAY | VT_R8;
	avarg[0].pparray = &pSAY;
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, avarg, 3, &varResult);
	if (SUCCEEDED(hr))
	{
		VariantToString(varResult, szTitle, nBufferSize);
		VariantClear(&varResult);
	}
	if (NULL != pSAX && NULL != pSAY)
	{
		SafeArrayGetUBound(pSAX, 1, &ubx);
		SafeArrayGetUBound(pSAY, 1, &uby);
		if (ubx == uby)
		{
			fSuccess = TRUE;
			*nArray = ubx + 1;
			*ppWaves = new double[*nArray];
			*ppValues = new double[*nArray];
			SafeArrayAccessData(pSAX, (void**)&arrX);
			SafeArrayAccessData(pSAY, (void**)&arrY);
			for (i = 0; i < (*nArray); i++)
			{
				(*ppWaves)[i] = arrX[i];
				(*ppValues)[i] = arrY[i];
			}
			SafeArrayUnaccessData(pSAX);
			SafeArrayUnaccessData(pSAY);
		}
	}
	SafeArrayDestroy(pSAX);
	SafeArrayDestroy(pSAY);
	return fSuccess;
}

HRESULT CIDataSet::GetGratingScan(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, L"GetGratingScan", &dispid);
	return this->m_pdisp->Invoke(dispid, IID_NULL, 0x0409, DISPATCH_METHOD,
		pDispParams, pVarResult, NULL, NULL);
}

// export to CSV file for Excel
BOOL CIDataSet::ExportToCSV(
	LPCTSTR			szFilePath)
{
	DISPID			dispid;
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	Utils_GetMemid(this->m_pdisp, L"ExportToCSV", &dispid);
	InitVariantFromString(szFilePath, &varg);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

HRESULT CIDataSet::GetADGainFactor(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, L"GetADGainFactor", &dispid);
	return this->m_pdisp->Invoke(dispid, IID_NULL, 0x0409, DISPATCH_METHOD,
		pDispParams, pVarResult, NULL, NULL);
}

// array scanning
BOOL CIDataSet::GetArrayScanning()
{
	TCHAR			szProgID[MAX_PATH];

	this->GetDataSetProgID(szProgID, MAX_PATH);
	return 0 == lstrcmpi(szProgID, L"Sciencetech.LDAData.1");
}

void CIDataSet::SetArrayScanning(
	BOOL			fArrayScanning)
{
	if (fArrayScanning && !this->GetArrayScanning())
	{
		HRESULT			hr;
		IDispatch	*	pdisp;
		TCHAR			szProgID[MAX_PATH];
		VARIANT			varResult;
		DISPID			dispid;
		BOOL			fSuccess = FALSE;
		CLSID			clsid;
		LPTSTR			szString = NULL;
		VARIANTARG		varg;

		this->GetDataSetProgID(szProgID, MAX_PATH);
		if (0 == lstrcmpi(szProgID, L"Sciencetech.OPTFile.1"))
		{
			hr = this->m_pdisp->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
			if (SUCCEEDED(hr))
			{
				Utils_GetMemid(pdisp, L"SaveToString", &dispid);
				hr = Utils_InvokeMethod(pdisp, dispid, NULL, 0, &varResult);
				if (SUCCEEDED(hr))
				{
					hr = VariantToStringAlloc(varResult, &szString);
					VariantClear(&varResult);
				}
				pdisp->Release();
			}
			if (NULL != szString)
			{
				// load and set new object
				hr = CLSIDFromProgID(L"Sciencetech.LDAData.1", &clsid);
				if (SUCCEEDED(hr)) hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (LPVOID*)&pdisp);
				if (SUCCEEDED(hr))
				{
					Utils_GetMemid(pdisp, L"InitFromString", &dispid);
					InitVariantFromString(szString, &varg);
					Utils_InvokeMethod(pdisp, dispid, &varg, 1, NULL);
					VariantClear(&varg);
					this->SetDataSet(pdisp);
					pdisp->Release();
				}
				CoTaskMemFree((LPVOID)szString);
			}
		}
	}
}

BOOL CIDataSet::SetArrayData(
	long			nArray,
	double		*	pX,
	double		*	pY)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	DISPID			dispid;
	VARIANTARG		avarg[2];
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;

	if (this->GetArrayScanning())
	{
		hr = this->m_pdisp->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
		if (SUCCEEDED(hr))
		{
			Utils_GetMemid(pdisp, L"AddArrayScan", &dispid);
			InitVariantFromDoubleArray(pX, nArray, &avarg[1]);
			InitVariantFromDoubleArray(pY, nArray, &avarg[0]);
			hr = Utils_InvokeMethod(pdisp, dispid, avarg, 2, &varResult);
			VariantClear(&avarg[1]);
			VariantClear(&avarg[0]);
			if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
			pdisp->Release();
		}
	}
	return fSuccess;
}


CIDataSet::CImpISink::CImpISink(CIDataSet * pIDataSet) :
	m_pIDataSet(pIDataSet),
	m_cRefs(0),
	m_dispidrequestDispersion(DISPID_UNKNOWN),
	m_dispidrequestAnalysisModeIndex(DISPID_UNKNOWN),
	m_dispidrequestOpticalTransfer(DISPID_UNKNOWN),
	m_dispidrequestADInfoString(DISPID_UNKNOWN),
	m_dispidrequestCalibrationScan(DISPID_UNKNOWN),
	m_dispidrequestCalibrationGain(DISPID_UNKNOWN)
{
	HRESULT				hr;
	IProvideClassInfo*	pProvideClassInfo;
	ITypeInfo		*	pClassInfo	= NULL;
	ITypeInfo		*	pTypeInfo = NULL;
	TYPEATTR		*	pTypeAttr;

	if (NULL == this->m_pIDataSet->m_pdisp) return;
	hr = this->m_pIDataSet->m_pdisp->QueryInterface(IID_IProvideClassInfo, (LPVOID*) &pProvideClassInfo);
	if (SUCCEEDED(hr))
	{
		hr = pProvideClassInfo->GetClassInfo(&pClassInfo);
		pProvideClassInfo->Release();
	}
	if (SUCCEEDED(hr))
	{
		hr = Utils_FindImplClassName(pClassInfo, TEXT("__clsIDataSet"), &pTypeInfo) ? S_OK : E_FAIL;
		pClassInfo->Release();
	}
	if (SUCCEEDED(hr))
	{
		hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			this->m_pIDataSet->m_iidSink	= pTypeAttr->guid;
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		Utils_GetMemid(pTypeInfo, L"requestDispersion", &m_dispidrequestDispersion);
		Utils_GetMemid(pTypeInfo, L"requestAnalysisModeIndex", &m_dispidrequestAnalysisModeIndex);
		Utils_GetMemid(pTypeInfo, L"requestOpticalTransfer", &m_dispidrequestOpticalTransfer);
		Utils_GetMemid(pTypeInfo, L"requestADInfoString", &m_dispidrequestADInfoString);
		Utils_GetMemid(pTypeInfo, L"requestCalibrationScan", &m_dispidrequestCalibrationScan);
		Utils_GetMemid(pTypeInfo, L"requestCalibrationGain", &m_dispidrequestCalibrationGain);
		pTypeInfo->Release();
	}
}

CIDataSet::CImpISink::~CImpISink()
{
}

// IUnknown methods
STDMETHODIMP CIDataSet::CImpISink::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IDispatch == riid || riid == this->m_pIDataSet->m_iidSink)
	{
		*ppv = this;
		this->AddRef();
		return S_OK;
	}
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CIDataSet::CImpISink::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CIDataSet::CImpISink::Release()
{
	ULONG			cRefs		= --m_cRefs;
	if (0 == cRefs) delete this;
	return cRefs;
}

// IDispatch methods
STDMETHODIMP CIDataSet::CImpISink::GetTypeInfoCount( 
									PUINT			pctinfo)
{
	*pctinfo	= 0;
	return S_OK;
}

STDMETHODIMP CIDataSet::CImpISink::GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CIDataSet::CImpISink::GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId)
{
	return E_NOTIMPL;
}

STDMETHODIMP CIDataSet::CImpISink::Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr)
{
	if (dispIdMember == this->m_dispidrequestDispersion)
	{
		if (2 == pDispParams->cArgs  && (VT_BYREF | VT_R8) == pDispParams->rgvarg[0].vt)
		{
			HRESULT			hr;
			VARIANT			varCopy;
			VariantInit(&varCopy);
			hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[1]));
			if (SUCCEEDED(hr)) hr = VariantChangeType(&varCopy, &varCopy, 0, VT_I4);
			if (SUCCEEDED(hr))
				*(pDispParams->rgvarg[0].pdblVal) = this->m_pIDataSet->OnrequestDispersion(varCopy.lVal);
		}
	}
	else if (dispIdMember == this->m_dispidrequestAnalysisModeIndex)
	{
		if (2 == pDispParams->cArgs && (VT_BYREF | VT_I2) == pDispParams->rgvarg[0].vt)
		{
			HRESULT			hr;
			VARIANTARG		varCopy;
	//		LPTSTR			szAnalysisMode	= NULL;
			VariantInit(&varCopy);
			hr = VariantCopyInd(&varCopy, &(pDispParams->rgvarg[1]));
			if (SUCCEEDED(hr)) hr = VariantChangeType(&varCopy, &varCopy, 0, VT_BSTR);
			if (SUCCEEDED(hr))
			{
	//			Utils_UnicodeToAnsi(varCopy.bstrVal, &szAnalysisMode);

				*(pDispParams->rgvarg[0].piVal) = this->m_pIDataSet->OnrequestAnalysisModeIndex(varCopy.bstrVal);
	//			CoTaskMemFree((LPVOID) szAnalysisMode);
				VariantClear(&varCopy);
			}
		}
	}
	else if (dispIdMember == this->m_dispidrequestOpticalTransfer)
	{
		if (2 == pDispParams->cArgs && 
			(VT_ARRAY | VT_BYREF | VT_R8) == pDispParams->rgvarg[1].vt	&&
			(VT_ARRAY | VT_BYREF | VT_R8) == pDispParams->rgvarg[0].vt)
		{
			long			nValues;
			double		*	pXValues;
			double		*	pYValues;
			SAFEARRAYBOUND	sab;
			double		*	arrX;
			double		*	arrY;
			long			i;
			if (this->m_pIDataSet->OnrequestOpticalTransfer(&nValues, &pXValues, &pYValues))
			{
				sab.lLbound		= 0;
				sab.cElements	= nValues;
				*(pDispParams->rgvarg[1].pparray)	= SafeArrayCreate(VT_R8, 1, &sab);
				*(pDispParams->rgvarg[0].pparray)	= SafeArrayCreate(VT_R8, 1, &sab);
				SafeArrayAccessData(*(pDispParams->rgvarg[1].pparray), (void**) &arrX);
				SafeArrayAccessData(*(pDispParams->rgvarg[0].pparray), (void**) &arrY);
				for (i=0; i<nValues; i++)
				{
					arrX[i]	= pXValues[i];
					arrY[i] = pYValues[i];
				}
				SafeArrayUnaccessData(*(pDispParams->rgvarg[1].pparray));
				SafeArrayUnaccessData(*(pDispParams->rgvarg[0].pparray));
				delete [] pXValues;
				delete [] pYValues;
			}
		}
	}
	else if (dispIdMember == this->m_dispidrequestADInfoString)
	{
		if (2 == pDispParams->cArgs && (VT_BYREF | VT_BSTR) == pDispParams->rgvarg[1].vt && (VT_BYREF | VT_BOOL) == pDispParams->rgvarg[0].vt)
		{
			LPTSTR			infoString = NULL;
			if (this->m_pIDataSet->OnrequestADInfoString(&infoString))
			{
				*(pDispParams->rgvarg[1].pbstrVal) = SysAllocString(infoString);
				*(pDispParams->rgvarg[0].pboolVal) = VARIANT_TRUE;
				CoTaskMemFree((LPVOID)infoString);
			}
		}
	}
	else if (dispIdMember == this->m_dispidrequestCalibrationScan)
	{
		this->OnrequestCalibrationScan(pDispParams);
	}
	else if (dispIdMember == this->m_dispidrequestCalibrationGain)
	{
		this->OnrequestCalibrationGain(pDispParams);
	}
	return S_OK;
}

void CIDataSet::CImpISink::OnrequestCalibrationScan(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANT			varg;
	UINT			uArgErr;
	short			grating;
	TCHAR			szFilter[MAX_PATH];
	short			detector;
	long			nValues;
	double		*	pX;
	double		*	pY;
	SAFEARRAYBOUND	sab;
	double		*	arrX;
	double		*	arrY;
	long			i;

	if (5 == pDispParams->cArgs &&
		(VT_BYREF | VT_ARRAY | VT_R8) == pDispParams->rgvarg[1].vt &&
		(VT_BYREF | VT_ARRAY | VT_R8) == pDispParams->rgvarg[0].vt)
	{
		VariantInit(&varg);
		hr = DispGetParam(pDispParams, 0, VT_I2, &varg, &uArgErr);
		if (SUCCEEDED(hr)) grating = varg.iVal;
		if (SUCCEEDED(hr)) hr = DispGetParam(pDispParams, 1, VT_BSTR, &varg, &uArgErr);
		if (SUCCEEDED(hr))
		{
			StringCchCopy(szFilter, MAX_PATH, varg.bstrVal);
			VariantClear(&varg);
		}
		if (SUCCEEDED(hr)) hr = DispGetParam(pDispParams, 2, VT_I2, &varg, &uArgErr);
		if (SUCCEEDED(hr))
		{
			detector = varg.iVal;
			if (this->m_pIDataSet->m_pMyObject->FirerequestCalibrationScan(
				grating, szFilter, detector, &nValues, &pX, &pY))
			{
				sab.lLbound = 0;
				sab.cElements = nValues;
				*(pDispParams->rgvarg[1].pparray) = SafeArrayCreate(VT_R8, 1, &sab);
				*(pDispParams->rgvarg[0].pparray) = SafeArrayCreate(VT_R8, 1, &sab);
				SafeArrayAccessData(*(pDispParams->rgvarg[1].pparray), (void**)&arrX);
				SafeArrayAccessData(*(pDispParams->rgvarg[0].pparray), (void**)&arrY);
				for (i = 0; i < nValues; i++)
				{
					arrX[i] = pX[i];
					arrY[i] = pY[i];
				}
				SafeArrayUnaccessData(*(pDispParams->rgvarg[1].pparray));
				SafeArrayUnaccessData(*(pDispParams->rgvarg[0].pparray));
				delete[] pX;
				delete[] pY;
			}
		}
	}
}

void CIDataSet::CImpISink::OnrequestCalibrationGain(
	DISPPARAMS	*	pDispParams)
{
	if (2 == pDispParams->cArgs &&
		(VT_BYREF | VT_R8) == pDispParams->rgvarg[0].vt)
	{
		HRESULT			hr;
		VARIANTARG		varg;
		UINT			uArgErr;
//		TCHAR			szGain[MAX_PATH];
		VariantInit(&varg);
		hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
		*(pDispParams->rgvarg[0].pdblVal) =
			this->m_pIDataSet->m_pMyObject->FirerequestCalibrationGain(varg.dblVal);
	}
}
