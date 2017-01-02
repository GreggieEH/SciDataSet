#include "stdafx.h"
#include "MyObject.h"
#include "MyGuids.h"

CMyObject::CMyObject() :
	// data members
	m_pImpIDispatch(NULL),
	m_pImpIConnectionPointContainer(NULL),
	m_pImpIProvideClassInfo2(NULL),
	// object reference count
	m_cRefs(0)
{
	ULONG		i;
	for (i = 0; i < MAX_CONN_PTS; i++)
		this->m_paConnPts[i] = NULL;
}


CMyObject::~CMyObject()
{
	Utils_DELETE_POINTER(this->m_pImpIDispatch);
	Utils_DELETE_POINTER(this->m_pImpIConnectionPointContainer);
	Utils_DELETE_POINTER(this->m_pImpIProvideClassInfo2);
	ULONG		i;
	for (i = 0; i < MAX_CONN_PTS; i++)
		Utils_RELEASE_INTERFACE(this->m_paConnPts[i]);
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
	if (NULL != (*ppv))
	{
		((IUnknown *)*ppv)->AddRef();
		return S_OK;
	}
	else
	{
		return this->_intQueryInterface(riid, ppv);
	}
}
STDMETHODIMP_(ULONG) CMyObject::AddRef()
{
	return ++m_cRefs;
}
STDMETHODIMP_(ULONG) CMyObject::Release()
{
	ULONG		cRefs = --m_cRefs;
	if (0 == cRefs)
	{
		++m_cRefs;
		objectsDown();
		delete this;
	}
	return cRefs;
}
// initialize the object
HRESULT CMyObject::Init()
{
	HRESULT				hr;

	this->m_pImpIDispatch = new CImpIDispatch(this);
	this->m_pImpIConnectionPointContainer = new CImpIConnectionPointContainer(this);
	this->m_pImpIProvideClassInfo2 = new CImpIProvideClassInfo2(this);
	if (NULL != this->m_pImpIDispatch					&&
		NULL != this->m_pImpIConnectionPointContainer	&&
		NULL != this->m_pImpIProvideClassInfo2)
	{
		hr = Utils_CreateConnectionPoint(this, MY_IIDSINK, &this->m_paConnPts[CONN_PT_CUSTOMSINK]);
		if (SUCCEEDED(hr)) hr = Utils_CreateConnectionPoint(this, IID_IPropertyNotifySink, &this->m_paConnPts[CONN_PT_PROPNOTIFY]);
	}
	else hr = E_OUTOFMEMORY;
	return hr;
}
// hook to allow additional interfaces
HRESULT CMyObject::_intQueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	*ppv = NULL;
	return E_NOINTERFACE;
}

// invoker
HRESULT CMyObject::InvokeHelper(
	DISPID			dispIdMember,
	WORD			wFlags,
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	UNREFERENCED_PARAMETER(dispIdMember);
	UNREFERENCED_PARAMETER(wFlags);
	UNREFERENCED_PARAMETER(pDispParams);
	UNREFERENCED_PARAMETER(pVarResult);
	return DISP_E_MEMBERNOTFOUND;
}

CMyObject::CImpIDispatch::CImpIDispatch(CMyObject * pMyObject) :
	m_pMyObject(pMyObject)
{

}

CMyObject::CImpIDispatch::~CImpIDispatch()
{

}

STDMETHODIMP CMyObject::CImpIDispatch::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	return this->m_pMyObject->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIDispatch::AddRef()
{
	return this->m_pMyObject->AddRef();
}
STDMETHODIMP_(ULONG) CMyObject::CImpIDispatch::Release()
{
	return this->m_pMyObject->Release();
}
// IDispatch methods
STDMETHODIMP CMyObject::CImpIDispatch::GetTypeInfoCount(
	PUINT			pctinfo)
{
	*pctinfo = 1;
	return S_OK;
}
STDMETHODIMP CMyObject::CImpIDispatch::GetTypeInfo(
	UINT			iTInfo,
	LCID			lcid,
	ITypeInfo	**	ppTInfo)
{
	HRESULT			hr;
	ITypeLib	*	pTypeLib;
	*ppTInfo = NULL;
	UNREFERENCED_PARAMETER(iTInfo);
	UNREFERENCED_PARAMETER(lcid);
	hr = GetTypeLib(&pTypeLib);
	if (SUCCEEDED(hr))
	{
		hr = pTypeLib->GetTypeInfoOfGuid(MY_IID, ppTInfo);
		pTypeLib->Release();
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
	HRESULT			hr;
	ITypeInfo	*	pTypeInfo;
	UNREFERENCED_PARAMETER(riid);

	hr = this->GetTypeInfo(0, lcid, &pTypeInfo);
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
	UNREFERENCED_PARAMETER(riid);
	UNREFERENCED_PARAMETER(lcid);
	UNREFERENCED_PARAMETER(pExcepInfo);
	UNREFERENCED_PARAMETER(puArgErr);
	return this->m_pMyObject->InvokeHelper(dispIdMember, wFlags, pDispParams, pVarResult);
}

CMyObject::CImpIConnectionPointContainer::CImpIConnectionPointContainer(CMyObject * pBackObj) :
	m_pBackObj(pBackObj)
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
	return this->m_pBackObj->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIConnectionPointContainer::AddRef()
{
	return this->m_pBackObj->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIConnectionPointContainer::Release()
{
	return this->m_pBackObj->Release();
}
// IConnectionPointContainer methods
STDMETHODIMP CMyObject::CImpIConnectionPointContainer::EnumConnectionPoints(
	IEnumConnectionPoints **ppEnum)
{
	return Utils_CreateEnumConnectionPoints(this, MAX_CONN_PTS, this->m_pBackObj->m_paConnPts, ppEnum);
}
STDMETHODIMP CMyObject::CImpIConnectionPointContainer::FindConnectionPoint(
	REFIID			riid,  //Requested connection point's interface identifier
	IConnectionPoint **ppCP)
{
	*ppCP = NULL;
	if (MY_IIDSINK == riid)
		*ppCP = this->m_pBackObj->m_paConnPts[CONN_PT_CUSTOMSINK];
	else if (IID_IPropertyNotifySink == riid)
		*ppCP = this->m_pBackObj->m_paConnPts[CONN_PT_PROPNOTIFY];
	if (NULL != *ppCP)
	{
		((IUnknown*)*ppCP)->AddRef();
		return S_OK;
	}
	else
	{
		return CONNECT_E_NOCONNECTION;
	}
}


CMyObject::CImpIProvideClassInfo2::CImpIProvideClassInfo2(CMyObject * pBackObj) :
	m_pBackObj(pBackObj)
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
	return this->m_pBackObj->QueryInterface(riid, ppv);
}
STDMETHODIMP_(ULONG) CMyObject::CImpIProvideClassInfo2::AddRef()
{
	return this->m_pBackObj->AddRef();
}
STDMETHODIMP_(ULONG) CMyObject::CImpIProvideClassInfo2::Release()
{
	return this->m_pBackObj->Release();
}
// IProvideClassInfo method
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::GetClassInfo(
	ITypeInfo	**	ppTI)
{
	HRESULT			hr;
	ITypeLib	*	pTypeLib;
	*ppTI = NULL;
	hr = GetTypeLib(&pTypeLib);
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
		*pGUID = MY_IIDSINK;
		return S_OK;
	}
	else
	{
		*pGUID = GUID_NULL;
		return E_INVALIDARG;
	}
}