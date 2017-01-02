#pragma once

class CMyObject;
class CTempFile;

class CIDataSet
{
public:
	CIDataSet(CMyObject * pMyObject);
	~CIDataSet(void);
	BOOL			SetDataSet(
						IDispatch	*	pdisp);
	BOOL			GetDataSet(
						IDispatch	**	ppdisp);
	BOOL			GetRadianceAvailable();
	BOOL			GetIrradianceAvailable();
	short int		GetAnalysisMode();
	void			SetAnalysisMode(
						short int		AnalysisMode);
	BOOL			ReadDataFile();
	BOOL			WriteDataFile();
	BOOL			GetWavelengths(
						long		*	nArray,
						double		**	ppArray);
	BOOL			GetSpectra(
						long		*	nArray,
						double		**	ppArray);
	BOOL			AddValue(                      
						double			NM, 
						double			Signal);
	BOOL			ViewSetup(
						HWND			hwnd);
	long			GetArraySize();
	void			SetCurrentExp(
						LPCTSTR			filter,
						short			grating,
						short			detector);
	BOOL			calcRadiance(
						BOOL			radiance,
						long		*	nArrays,
						double		**	ppWaves,
						double		**	ppRadiance);
	HRESULT			GetCreateTime(
						VARIANT		*	pVarResult);
	BOOL			GetFileName(
						LPTSTR			szFileName,
						UINT			nBufferSize);
	void			SetFileName(
						LPCTSTR			szFileName);
	BOOL			GetAmCalibration();
	void			SetAmCalibration(
						BOOL			AmCalibration);
	BOOL			FormOutFile(
						LPCTSTR			szCalibration,
						LPCTSTR			szFileExt,
						long			nArray,
						double		*	pXValues,
						double		*	pYValues);
	double			GetDataValue(
						double			xValue);
	void			SetDataValue(
						double			xValue,
						double			yValue);
	// data set prog id
	BOOL			GetDataSetProgID(
						LPTSTR			szProgID,
						UINT			nBufferSize);
	// calculate the optical transfer function (if available)
	BOOL			CalculateOPT(
						long		*	nValues,
						double		**	ppWaves,
						double		**	ppValues);
	// form the optical transfer function
	BOOL			FormOpticalTransferFile(
						LPCTSTR			filePath);
	// set the AD info string
	BOOL			SetADInfoString();
	// comment string
	BOOL			GetComment(
						LPTSTR			szComment,
						UINT			nBufferSize);
	void			SetComment(
						LPCTSTR			szComment);
	// extra scan data
	void			ExtraScanData(
						BSTR			title,
						double			wave,
						double			dValue);
	// default units
	BOOL			GetDefaultUnits(
						LPTSTR			szDefaultUnits,
						UINT			nBufferSize);
	void			SetDefaultUnits(
						LPCTSTR			szDefaultUnits);
	// default title
	BOOL			GetDefaultTitle(
						LPTSTR			szDefaultTitle,
						UINT			nBufferSize);
	void			SetDefaultTitle(
						LPCTSTR			szDefaultTitle);
	// are counts available?
	BOOL			GetCountsAvailable();
	void			SetCountsAvailable(
						BOOL			CountsAvailable);
	// extra values
	long			GetNumberExtraScanArrays();
	BOOL			GetExtraScanArray(
						long			index,
						long		*	nArray,
						double		**	ppWaves,
						double		**	ppValues,
						LPTSTR			szTitle,
						UINT			nBufferSize);
	HRESULT			GetGratingScan(
						DISPPARAMS	*	pDispParams,
						VARIANT		*	pVarResult);
	// export to CSV file for Excel
	BOOL			ExportToCSV(
						LPCTSTR			szFilePath);
	HRESULT			GetADGainFactor(
						DISPPARAMS	*	pDispParams,
						VARIANT		*	pVarResult);
	// array scanning
	BOOL			GetArrayScanning();
	void			SetArrayScanning(
						BOOL			fArrayScanning);
	BOOL			SetArrayData(
						long			nArray,
						double		*	pX,
						double		*	pY);
protected:
	double			OnrequestDispersion(
						long			grating);
	short int		OnrequestAnalysisModeIndex(
						LPCTSTR			szAnalysisMode);
	BOOL			OnrequestOpticalTransfer(
						long		*	nValues,
						double		**	ppXValues,
						double		**	ppYValues);
	BOOL			OnrequestADInfoString(
						LPTSTR		*	szInfoString);
	// check the file name for saving
	void			CheckFileExtension();
private:
	CMyObject	*	m_pMyObject;
	IDispatch	*	m_pdisp;
	IID				m_iidSink;
	DWORD			m_dwCookie;
	// file to store readings in case of crash
	CTempFile	*	m_pTempFile;

	class CImpISink : public IDispatch
	{
	public:
		CImpISink(CIDataSet * pIDataSet);
		~CImpISink();
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
		void					OnrequestCalibrationScan(
									DISPPARAMS	*	pDispParams);
		void					OnrequestCalibrationGain(
									DISPPARAMS	*	pDispParams);
	private:
		CIDataSet			*	m_pIDataSet;
		ULONG					m_cRefs;
		DISPID					m_dispidrequestDispersion;
		DISPID					m_dispidrequestAnalysisModeIndex;
		DISPID					m_dispidrequestOpticalTransfer;
		DISPID					m_dispidrequestADInfoString;
		DISPID					m_dispidrequestCalibrationScan;
		DISPID					m_dispidrequestCalibrationGain;
	};
	friend CImpISink;




};
