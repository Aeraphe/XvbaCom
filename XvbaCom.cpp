#include "XvbaCom.h"
#include "windows.h"
#include <iostream>
#include <string>
#include <comdef.h>

enum XVBA_ERROR {

	CO_CREATE_INSTANCE = -141,
	OPEN_DOCUMENT = -142,
	INVOKE = -143,
	GET_WORKBOOK = -144,

};



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
		if (FAILED(hr)) return hr;
		
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




	if (!result.iVal) {

		valueOut = &result.iVal;
	}
	if(!result.dblVal) {
		BSTR bs = SysAllocString(L"Hello");

		std::wstring* r;
		std::wstring  resp = ConvertBSTRToMBS(result.bstrVal);
	
	
		valueOut = (LPCSTR*) "OK";
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


HRESULT XvbaRelease(IDispatch*& pIn) {

	return pIn->Release();
};


std::wstring ConvertBSTRToMBS(BSTR bstr)
{
	int wslen = ::SysStringLen(bstr);
	return ConvertWCSToMBS((wchar_t*)bstr, wslen);
}

std::wstring ConvertWCSToMBS(const wchar_t* pstr, long wslen)
{
	int len = ::WideCharToMultiByte(CP_ACP, 0, pstr, wslen, NULL, 0, NULL, NULL);

	std::wstring dblstr(len, '\0');
	len = ::WideCharToMultiByte(CP_ACP, 0 /* no flags */,
		pstr, wslen /* not necessary NULL-terminated */,
		NULL, len,
		NULL, NULL /* no default char */);

	return dblstr;
}

BSTR ConvertMBSToBSTR(const std::string& str)
{
	int wslen = ::MultiByteToWideChar(CP_ACP, 0 /* no flags */,
		str.data(), str.length(),
		NULL, 0);

	BSTR wsdata = ::SysAllocStringLen(NULL, wslen);
	::MultiByteToWideChar(CP_ACP, 0 /* no flags */,
		str.data(), str.length(),
		wsdata, wslen);
	return wsdata;
}