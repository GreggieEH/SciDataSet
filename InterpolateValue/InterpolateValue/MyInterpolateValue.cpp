#include "stdafx.h"
#include "MyInterpolateValue.h"
#include "MyGuids.h"
#include "dispids.h"

// interpolate
float interpolateValue(
	float		*	px,
	float		*	py,
	unsigned long	n,
	double			xValue);

CMyInterpolateValue::CMyInterpolateValue() :
	CMyObject(),
	// data for the numerical recipes routines
	m_nArray(0),
	m_px(NULL),
	m_py(NULL)
{
	this->m_XArray.clear();
	this->m_YArray.clear();
}

CMyInterpolateValue::~CMyInterpolateValue()
{
	this->m_XArray.clear();
	this->m_YArray.clear();
	// clear the numerical recipes data
	this->ClearNRData();
}

// sink events
void CMyInterpolateValue::FireError(
	LPCTSTR			szError)
{
	VARIANTARG			varg;
	InitVariantFromString(szError, &varg);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_Error, &varg, 1);
	VariantClear(&varg);
}

// invoker
HRESULT CMyInterpolateValue::InvokeHelper(
	DISPID			dispIdMember,
	WORD			wFlags,
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	switch (dispIdMember)
	{
	case DISPID_XArray:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetXArray(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetXArray(pDispParams);
		break;
	case DISPID_YArray:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetYArray(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetYArray(pDispParams);
		break;
	case DISPID_arraySize:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetarraySize(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetarraySize(pDispParams);
		break;
	case DISPID_XValue:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetXValue(pDispParams, pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetXValue(pDispParams);
		break;
	case DISPID_YValue:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetmYValue(pDispParams, pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetYValue(pDispParams);
		break;
	case DISPID_interpolateValue:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->interpolateValue(pDispParams, pVarResult);
		break;
	default:
		break;
	}
	return DISP_E_MEMBERNOTFOUND;
}

// properties and methods
HRESULT CMyInterpolateValue::GetXArray(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	unsigned long	nArray = this->m_XArray.size();
	if (nArray > 0)
	{
		SAFEARRAYBOUND		sab;
		double			*	arr;
		unsigned long		i;

		sab.lLbound = 0;
		sab.cElements = nArray;
		pVarResult->vt = VT_ARRAY | VT_R8;
		pVarResult->parray = SafeArrayCreate(VT_R8, 1, &sab);
		SafeArrayAccessData(pVarResult->parray, (void**)&arr);
		for (i = 0; i < nArray; i++)
			arr[i] = this->m_XArray[i];
		SafeArrayUnaccessData(pVarResult->parray);
	}
	return S_OK;
}

HRESULT CMyInterpolateValue::SetXArray(
	DISPPARAMS	*	pDispParams)
{
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_ARRAY | VT_R8) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	long		ub;
	long		i;
	double	*	arr;
	this->m_XArray.clear();
	SafeArrayGetUBound(pDispParams->rgvarg[0].parray, 1, &ub);
	SafeArrayAccessData(pDispParams->rgvarg[0].parray, (void**)&arr);
	for (i = 0; i <= ub; i++)
		this->m_XArray.push_back(arr[i]);
	SafeArrayUnaccessData(pDispParams->rgvarg[0].parray);
	this->ClearNRData();			// clear the numerical recipes data
	Utils_OnPropChanged(this, DISPID_XArray);
	return S_OK;
}

HRESULT CMyInterpolateValue::GetYArray(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	unsigned long	nArray = this->m_YArray.size();
	if (nArray > 0)
	{
		SAFEARRAYBOUND		sab;
		double			*	arr;
		unsigned long		i;

		sab.lLbound = 0;
		sab.cElements = nArray;
		pVarResult->vt = VT_ARRAY | VT_R8;
		pVarResult->parray = SafeArrayCreate(VT_R8, 1, &sab);
		SafeArrayAccessData(pVarResult->parray, (void**)&arr);
		for (i = 0; i < nArray; i++)
			arr[i] = this->m_YArray[i];
		SafeArrayUnaccessData(pVarResult->parray);
	}
	return S_OK;
}

HRESULT CMyInterpolateValue::SetYArray(
	DISPPARAMS	*	pDispParams)
{
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_ARRAY | VT_R8) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	long		ub;
	long		i;
	double	*	arr;
	this->m_YArray.clear();
	SafeArrayGetUBound(pDispParams->rgvarg[0].parray, 1, &ub);
	SafeArrayAccessData(pDispParams->rgvarg[0].parray, (void**)&arr);
	for (i = 0; i <= ub; i++)
		this->m_YArray.push_back(arr[i]);
	SafeArrayUnaccessData(pDispParams->rgvarg[0].parray);
	this->ClearNRData();			// clear the numerical recipes data
	Utils_OnPropChanged(this, DISPID_YArray);
	return S_OK;
}

HRESULT CMyInterpolateValue::GetarraySize(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromUInt32(this->m_XArray.size(), pVarResult);
	return S_OK;
}

HRESULT CMyInterpolateValue::SetarraySize(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_UI4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_XArray.resize(varg.ulVal);
	this->m_YArray.resize(varg.ulVal);
	this->ClearNRData();
	Utils_OnPropChanged(this, DISPID_arraySize);
	return S_OK;
}

HRESULT CMyInterpolateValue::GetXValue(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_UI4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	if (varg.ulVal >= 0 && varg.ulVal < this->m_XArray.size())
	{
		InitVariantFromDouble(this->m_XArray[varg.ulVal], pVarResult);
	}
	return S_OK;
}

HRESULT CMyInterpolateValue::SetXValue(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	unsigned long	index;
	double			xValue;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_UI4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	index = varg.ulVal;
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	xValue = varg.dblVal;
	if (index >= 0 && index < this->m_XArray.size())
	{
		this->m_XArray[index] = xValue;
		this->ClearNRData();
		Utils_OnPropChanged(this, DISPID_XValue);
	}
	return S_OK;
}

HRESULT CMyInterpolateValue::GetmYValue(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_UI4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	if (varg.ulVal >= 0 && varg.ulVal < this->m_YArray.size())
	{
		InitVariantFromDouble(this->m_YArray[varg.ulVal], pVarResult);
	}
	return S_OK;
}

HRESULT CMyInterpolateValue::SetYValue(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	unsigned long	index;
	double			yValue;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_UI4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	index = varg.ulVal;
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	yValue = varg.dblVal;
	if (index >= 0 && index < this->m_XArray.size())
	{
		this->m_YArray[index] = yValue;
		this->ClearNRData();
		Utils_OnPropChanged(this, DISPID_YValue);
	}
	return S_OK;
}

HRESULT CMyInterpolateValue::interpolateValue(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	double			dval;
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	if (!this->SetNRData()) return E_UNEXPECTED;
	dval = ::interpolateValue(this->m_px, this->m_py, this->m_nArray, (float)varg.dblVal);
	InitVariantFromDouble(dval, pVarResult);
	return S_OK;
}

// clear the numerical recipes data
void CMyInterpolateValue::ClearNRData()
{
	if (NULL != this->m_px)
	{
		delete[] this->m_px;
		this->m_px = NULL;
	}
	if (NULL != this->m_py)
	{
		delete[] this->m_py;
		this->m_py = NULL;
	}
	this->m_nArray = 0;

}
// set the numerical recipes data
BOOL CMyInterpolateValue::SetNRData()
{
	if (NULL != this->m_px && NULL != this->m_py && this->m_nArray > 0)
	{
		return TRUE;
	}
	if (this->m_XArray.size() == this->m_YArray.size())
	{
		unsigned long		i;
		this->m_nArray = this->m_XArray.size();
		this->m_px = new float[this->m_nArray + 1];
		this->m_py = new float[this->m_nArray + 1];
		this->m_px[0] = 0.0;
		this->m_py[0] = 0.0;
		for (i = 0; i < this->m_nArray; i++)
		{
			this->m_px[i + 1] = (float) this->m_XArray[i];
			this->m_py[i + 1] = (float) this->m_YArray[i];
		}
		return TRUE;
	}
	else
	{
		this->FireError(L"Inconsistent Array Sizes");
		return FALSE;
	}
}
