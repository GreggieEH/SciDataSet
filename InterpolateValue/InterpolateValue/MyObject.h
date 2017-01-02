#pragma once
// basic com object class

class CMyObject : public IUnknown
{
public:
	CMyObject();
	virtual ~CMyObject();
	// IUnknown methods
	STDMETHODIMP			QueryInterface(
								REFIID			riid,
								LPVOID		*	ppv);
	STDMETHODIMP_(ULONG)	AddRef();
	STDMETHODIMP_(ULONG)	Release();
	// initialize the object
	virtual HRESULT			Init();
protected:
	// hook to allow additional interfaces
	virtual HRESULT			_intQueryInterface(
								REFIID			riid,
								LPVOID		*	ppv);
	// invoker
	virtual HRESULT			InvokeHelper(
								DISPID			dispIdMember,
								WORD			wFlags,
								DISPPARAMS	*	pDispParams,
								VARIANT		*	pVarResult);
private:
	// nested classes
	// dispatch interface
	class CImpIDispatch : public IDispatch
	{
	public:
		CImpIDispatch(CMyObject * pMyObject);
		~CImpIDispatch();
		STDMETHODIMP		QueryInterface(
								REFIID			riid,
								LPVOID		*	ppv);
		STDMETHODIMP_(ULONG) AddRef();
		STDMETHODIMP_(ULONG) Release();
		// IDispatch methods
		STDMETHODIMP		GetTypeInfoCount(
								PUINT			pctinfo);
		STDMETHODIMP		GetTypeInfo(
								UINT			iTInfo,
								LCID			lcid,
								ITypeInfo	**	ppTInfo);
		STDMETHODIMP		GetIDsOfNames(
								REFIID			riid,
								OLECHAR		**  rgszNames,
								UINT			cNames,
								LCID			lcid,
								DISPID		*	rgDispId);
		STDMETHODIMP		Invoke(
								DISPID			dispIdMember,
								REFIID			riid,
								LCID			lcid,
								WORD			wFlags,
								DISPPARAMS	*	pDispParams,
								VARIANT		*	pVarResult,
								EXCEPINFO	*	pExcepInfo,
								PUINT			puArgErr);
	private:
		CMyObject		*	m_pMyObject;
	};
	class CImpIConnectionPointContainer : public IConnectionPointContainer
	{
	public:
		CImpIConnectionPointContainer(CMyObject * pBackObj);
		~CImpIConnectionPointContainer();
		// IUnknown methods
		STDMETHODIMP		QueryInterface(
								REFIID			riid,
								LPVOID		*	ppv);
		STDMETHODIMP_(ULONG) AddRef();
		STDMETHODIMP_(ULONG) Release();
		// IConnectionPointContainer methods
		STDMETHODIMP		EnumConnectionPoints(
								IEnumConnectionPoints **ppEnum);
		STDMETHODIMP		FindConnectionPoint(
								REFIID			riid,  //Requested connection point's interface identifier
								IConnectionPoint **ppCP);
	private:
		CMyObject		*	m_pBackObj;
	};
	class CImpIProvideClassInfo2 : public IProvideClassInfo2
	{
	public:
		CImpIProvideClassInfo2(CMyObject * pBackObj);
		~CImpIProvideClassInfo2();
		// IUnknown methods
		STDMETHODIMP		QueryInterface(
								REFIID			riid,
								LPVOID		*	ppv);
		STDMETHODIMP_(ULONG) AddRef();
		STDMETHODIMP_(ULONG) Release();
		// IProvideClassInfo method
		STDMETHODIMP		GetClassInfo(
								ITypeInfo	**	ppTI);
		// IProvideClassInfo2 method
		STDMETHODIMP		GetGUID(
								DWORD			dwGuidKind,  //Desired GUID
								GUID		*	pGUID);
	private:
		CMyObject		*	m_pBackObj;
	};
	friend CImpIDispatch;
	friend CImpIConnectionPointContainer;
	friend CImpIProvideClassInfo2;
	// data members
	CImpIDispatch		*	m_pImpIDispatch;
	CImpIConnectionPointContainer*	m_pImpIConnectionPointContainer;
	CImpIProvideClassInfo2*	m_pImpIProvideClassInfo2;
	// object reference count
	ULONG					m_cRefs;
	// connection points
	IConnectionPoint	*	m_paConnPts[MAX_CONN_PTS];
};

