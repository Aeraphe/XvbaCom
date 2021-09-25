#include "XvbaCom.h"
#include "windows.h"
#include <iostream>

enum XVBA_ERROR {

	CO_CREATE_INSTANCE = -141,
	OPEN_DOCUMENT = -142,
	INVOKE = -143,
	GET_WORKBOOK = -144,

};



int XvbaShowApplication(IDispatch*& app) {

	HRESULT hr;
	VARIANT x;
	x.vt = VT_I4;
	x.lVal = 1;
	hr = XvbaInvoke(DISPATCH_PROPERTYPUT, NULL, app, L"Visible", 1, x);

	if (FAILED(hr)) {
		return hr;
	}

	return hr;
}

HRESULT XvbaCoCreateInstance(LPCOLESTR lpszProgId, IDispatch*& app) {

	CLSID clsId;

	HRESULT hr;

	hr = CoInitialize(NULL);

	if (FAILED(hr)) {

		return hr;

	}

	hr = CLSIDFromProgID(lpszProgId, &clsId);

	if (FAILED(hr)) {
		return hr;

	}


	hr = CoCreateInstance(clsId, NULL, CLSCTX_SERVER, IID_IDispatch, (void**)&app);

	if (FAILED(hr)) {
		return hr;

	}


	return hr;
}

HRESULT XvbaGetVBComponets(IDispatch*& app, IDispatch*& pVBAComponents) {

	HRESULT hr;
	IDispatch* pVbProject = (IDispatch*)NULL;

	// GetVBProject
	{
		VARIANT result;
		VariantInit(&result);
		hr = XvbaInvoke(DISPATCH_PROPERTYGET, &result, app, L"VBProject", 0);
		pVbProject = result.pdispVal;

		if (FAILED(hr)) {
			return hr;
		}
	}

	// GetDocuments
	{
		VARIANT result;
		VariantInit(&result);
		hr = XvbaInvoke(DISPATCH_PROPERTYGET, &result, pVbProject, L"VBComponents", 0);
		pVBAComponents = result.pdispVal;

		if (FAILED(hr)) {
			return hr;
		}
	}


	return hr;
}

HRESULT XvbaGetMethod(IDispatch*& pIn, IDispatch*& pOut, LPCTSTR pMenthodName) {

	HRESULT hr = 0;
	VARIANT result;
	VariantInit(&result);
	hr = XvbaInvoke(DISPATCH_PROPERTYGET, &result, pIn, pMenthodName, 0);
	pOut = result.pdispVal;

	return   hr;
}

HRESULT XvbaCall(LPCTSTR pPropToCall, IDispatch*& pIn, LPCTSTR param, IDispatch*& pOut, VOID*& valueOut, int paramType) {

	HRESULT hr = 0;

	VARIANT result;
	VariantInit(&result);

	if (!param || !param[0]) {

		hr = XvbaInvoke(DISPATCH_PROPERTYGET, &result, pIn, pPropToCall, 0);
	}
	else {
		VARIANT vProperty;
		VariantInit(&vProperty);
		//Integer
		if (paramType == 1) {
			vProperty.vt = VT_I4;
			vProperty.lVal = 1;
		}
		//String
		else {
			vProperty.vt = VT_BSTR;
			vProperty.bstrVal = SysAllocString(param);
		}

		hr = XvbaInvoke(DISPATCH_PROPERTYGET, &result, pIn, pPropToCall, 1, vProperty);

	}


	if (FAILED(hr)) {
		return hr;
	}

	if (result.iVal) {

		valueOut = &result.iVal;
	}
	else {
		valueOut = &result.bstrVal;
	}

	pOut = result.pdispVal;


	return hr;
}

HRESULT XvbaSetVal(LPCTSTR pPropToCall, IDispatch*& pIn, LPCTSTR param,  int paramType) {

	HRESULT hr = 0;


	VARIANT vProperty;
	VariantInit(&vProperty);
	//Integer
	if (paramType == 1) {
		vProperty.vt = VT_I4;
		vProperty.lVal = 1;
	}
	//String
	else {
		vProperty.vt = VT_BSTR;
		vProperty.bstrVal = SysAllocString(param);
	}

	hr = XvbaInvoke(DISPATCH_PROPERTYPUT, NULL, pIn, pPropToCall, 1, vProperty);



	if (FAILED(hr)) {
		return hr;
	}




	return hr;
}


//check if number or string
bool check_number(std::string str) {
	for (int i = 0; i < str.length(); i++)
		if (isdigit(str[i]) == false)
			return false;
	return true;
}