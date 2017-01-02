#pragma once

class CIDataSet;
class CGaussianFit;

class CMyObject : public IUnknown
{
public:
	CMyObject(IUnknown * pUnkOuter);
	~CMyObject(void);
	// IUnknown methods
	STDMETHODIMP				QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
	STDMETHODIMP_(ULONG)		AddRef();
	STDMETHODIMP_(ULONG)		Release();
	// initialization
	HRESULT						Init();
	// sink events
	double						FirerequestDispersion(
									long			gratingID);
	short int					FirerequestAnalysisModeIndex(
									LPCTSTR			szAnalysisMode);
	BOOL						FirerequestAnalysisModeString(
									short int		ModeIndex,
									LPTSTR			szAnalysisMode,
									UINT			nBufferSize);
	BOOL						FirerequestOpticalTransfer(
									long		*	nValues,
									double		**	ppXValues,
									double		**	ppYValues);
	long						FireRequestOPTArraySize();
	void						FireRequestOPTDataValue(
									long			index,
									double		*	x,
									double		*	y);
	void						FireRequestSetupDisplay(
									LPCTSTR			component);
	void						FireupdateOptFile(
									LPCTSTR			opt);
	BOOL						FirerequestADInfoString(
									LPTSTR		*	infoString);
	BOOL						FirerequestCalibrationScan(
									short			grating,
									LPCTSTR			filter,
									short			detector,
									long		*	nValues,
									double		**	ppXValues,
									double		**	ppYValues);
	double						FirerequestCalibrationGain(
									double			wavelength);
	BOOL						FirerequestWorkingDirectory(
									LPTSTR			szWorkingDirectory,
									UINT			nBufferSize);
	void						FirebrowseForWorkingDirectory(
									HWND			hwndParent);
	void						setWorkingDirectory(
									LPCTSTR			workingDirectory);
	// flag skip the last data point
	BOOL						GetSkipLastPoint();
	// set the AD info string
	BOOL						SetADInfoString();

protected:
	void						SetCalibrationData(
									long			nValues,
									double		*	pX,
									double		*	pY);
private:
	class CImpIDispatch : public IDispatch
	{
	public:
		CImpIDispatch(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIDispatch();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IDispatch methods
		STDMETHODIMP			GetTypeInfoCount( 
									PUINT			pctinfo);
		STDMETHODIMP			GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo);
		STDMETHODIMP			GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId);
		STDMETHODIMP			Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr); 
	protected:
		HRESULT					GetDataProgID(
									VARIANT		*	pVarResult);
		HRESULT					SetDataProgID(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetRadianceAvailable(
									VARIANT		*	pVarResult);
		HRESULT					GetIrradianceAvailable(
									VARIANT		*	pVarResult);
		HRESULT					ReadDataFile(
									VARIANT		*	pVarResult);
		HRESULT					WriteDataFile(
									VARIANT		*	pVarResult);
		HRESULT					GetWavelengths(
									VARIANT		*	pVarResult);
		HRESULT					GetSpectra(
									VARIANT		*	pVarResult);
		HRESULT					AddValue(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					ViewSetup(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetArraySize(
									VARIANT		*	pVarResult);
		HRESULT					SetCurrentExp(
									DISPPARAMS	*	pDispParams);
		HRESULT					calcRadiance(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					setObject(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetAnalysisMode(
									VARIANT		*	pVarResult);
		HRESULT					SetAnalysisMode(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetCreateTime(
									VARIANT		*	pVarResult);
		HRESULT					FormOutFile(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetFileName(
									VARIANT		*	pVarResult);
		HRESULT					SetFileName(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetAmCalibration(
									VARIANT		*	pVarResult);
		HRESULT					SetAmCalibration(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetDataValue(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					SetDataValue(
									DISPPARAMS	*	pDispParams);
		HRESULT					LoadFromString(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetRadianceWaves(
									VARIANT		*	pVarResult);
		HRESULT					GetRadiance(
									VARIANT		*	pVarResult);
		HRESULT					GetHaveOPT(
									VARIANT		*	pVarResult);
		HRESULT					SetHaveOPT(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetOPTWaves(
									VARIANT		*	pVarResult);
		HRESULT					GetOPTValues(
									VARIANT		*	pVarResult);
		HRESULT					FormOpticalTransferFile(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
//		HRESULT					SetCalibrationWavelengths(
//									DISPPARAMS	*	pDispParams);
		HRESULT					SetCalibrationData(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					FileNameFromPath(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					NewObjectSetup(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetScanType(
									VARIANT		*	pVarResult);
		HRESULT					SetScanType(
									DISPPARAMS	*	pDispParams);
		HRESULT					SelectCalibrationStandard(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetHaveCalibrationStandard(
									VARIANT		*	pVarResult);
		HRESULT					GetComment(
									VARIANT		*	pVarResult);
		HRESULT					SetComment(
									DISPPARAMS	*	pDispParams);
		HRESULT					ExtraScanData(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetDefaultUnits(
									VARIANT		*	pVarResult);
		HRESULT					SetDefaultUnits(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetDefaultTitle(
									VARIANT		*	pVarResult);
		HRESULT					SetDefaultTitle(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetCountsAvailable(
									VARIANT		*	pVarResult);
		HRESULT					SetCountsAvailable(
									DISPPARAMS	*	pDispParams);
		HRESULT					ExportToCSV(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetADGainFactor(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetArrayScanning(
									VARIANT		*	pVarResult);
		HRESULT					SetArrayScanning(
									DISPPARAMS	*	pDispParams);
		HRESULT					SetArrayData(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetNumPeaks(
									VARIANT		*	pVarResult);
		HRESULT					FindPeaks(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					DisplayPeaks(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetPeakPosition(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetPeakHeight(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetPeakWidth(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					IntegrateSpectra(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					SmoothSpectra(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					CalcAverage(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
		ITypeInfo			*	m_pTypeInfo;
		TCHAR					m_szProgID[MAX_PATH];
		// scan type
		TCHAR					m_ScanType[MAX_PATH];
		// radiance calculation
		long					m_nValuesRadiance;
		double				*	m_pRadianceWaves;
		double				*	m_pRadianceValues;
		// optical transfer data
		long					m_nValuesOPT;
		double				*	m_pWavesOPT;
		double				*	m_pValuesOPT;
	};
	class CImpIProvideClassInfo2 : public IProvideClassInfo2
	{
	public:
		CImpIProvideClassInfo2(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIProvideClassInfo2();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IProvideClassInfo method
		STDMETHODIMP			GetClassInfo(
									ITypeInfo	**	ppTI);  
		// IProvideClassInfo2 method
		STDMETHODIMP			GetGUID(
									DWORD			dwGuidKind,  //Desired GUID
									GUID		*	pGUID);       
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImpIConnectionPointContainer : public IConnectionPointContainer
	{
	public:
		CImpIConnectionPointContainer(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIConnectionPointContainer();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IConnectionPointContainer methods
		STDMETHODIMP			EnumConnectionPoints(
									IEnumConnectionPoints **ppEnum);
		STDMETHODIMP			FindConnectionPoint(
									REFIID			riid,  //Requested connection point's interface identifier
									IConnectionPoint **ppCP);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	friend CImpIDispatch;
	friend CImpIConnectionPointContainer;
	friend CImpIProvideClassInfo2;
	// data members
	CImpIDispatch			*	m_pImpIDispatch;
	CImpIConnectionPointContainer*	m_pImpIConnectionPointContainer;
	CImpIProvideClassInfo2	*	m_pImpIProvideClassInfo2;
	// outer unknown for aggregation
	IUnknown				*	m_pUnkOuter;
	// object reference count
	ULONG						m_cRefs;
	// connection points array
	IConnectionPoint		*	m_paConnPts[MAX_CONN_PTS];
	// data set
	CIDataSet				*	m_pIDataSet;
	// calibration data
	long						m_nCalData;
	double					*	m_pXCal;
	double					*	m_pYCal;
	// Gaussian fit
	CGaussianFit			*	m_pGaussianFit;
};
