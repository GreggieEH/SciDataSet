#pragma once
#include "MyObject.h"

using namespace std;

class CMyInterpolateValue : public CMyObject
{
public:
	CMyInterpolateValue();
	virtual ~CMyInterpolateValue();
	// sink events
	void				FireError(
							LPCTSTR			szError);
protected:
	// invoker
	virtual HRESULT		InvokeHelper(
							DISPID			dispIdMember,
							WORD			wFlags,
							DISPPARAMS	*	pDispParams,
							VARIANT		*	pVarResult);
	// properties and methods
	HRESULT				GetXArray(
							VARIANT		*	pVarResult);
	HRESULT				SetXArray(
							DISPPARAMS	*	pDispParams);
	HRESULT				GetYArray(
							VARIANT		*	pVarResult);
	HRESULT				SetYArray(
							DISPPARAMS	*	pDispParams);
	HRESULT				GetarraySize(
							VARIANT		*	pVarResult);
	HRESULT				SetarraySize(
							DISPPARAMS	*	pDispParams);
	HRESULT				GetXValue(
							DISPPARAMS	*	pDispParams,
							VARIANT		*	pVarResult);
	HRESULT				SetXValue(
							DISPPARAMS	*	pDispParams);
	HRESULT				GetmYValue(
							DISPPARAMS	*	pDispParams,
							VARIANT		*	pVarResult);
	HRESULT				SetYValue(
							DISPPARAMS	*	pDispParams);
	HRESULT				interpolateValue(
							DISPPARAMS	*	pDispParams,
							VARIANT		*	pVarResult);
	// clear the numerical recipes data
	void				ClearNRData();
	// set the numerical recipes data
	BOOL				SetNRData();
private:
	// input data
	vector<double>		m_XArray;
	vector<double>		m_YArray;
	// data for the numerical recipes routines
	unsigned long		m_nArray;
	float			*	m_px;
	float			*	m_py;

};

