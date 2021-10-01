#include "XvbaCom.h"
#include "windows.h"
#include <iostream>
#include <string>
#include <comdef.h>
#include <atlbase.h>
#include <stdlib.h>

#include "comutil.h"


enum XVBA_ERROR {

	CO_CREATE_INSTANCE = -141,
	OPEN_DOCUMENT = -142,
	INVOKE = -143,
	GET_WORKBOOK = -144,

};


struct InputValueStruct {
	INT32 type;
	INT32 intValue;
	char*& stringValue;
	bool boolValue;
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



HRESULT XvbaGetMethod(IDispatch*& pIn, IDispatch*& pOut, LPCTSTR pMenthodName) {

	HRESULT hr = 0;
	VARIANT result;
	VariantInit(&result);
	hr = XvbaInvoke(DISPATCH_PROPERTYGET | DISPATCH_METHOD, &result, pIn, pMenthodName, 0);
	pOut = result.pdispVal;

	return   hr;
}


HRESULT XvbaGetPropertyRef(IDispatch*& pIn, IDispatch*& pOut, LPCTSTR pMenthodName) {

	HRESULT hr = 0;
	VARIANT result;
	VariantInit(&result);
	hr = XvbaInvoke(DISPATCH_PROPERTYPUTREF, &result, pIn, pMenthodName, 0);
	pOut = result.pdispVal;

	return   hr;
}


HRESULT XvbaCall(LPCTSTR pPropToCall, IDispatch*& pIn, VOID*& param, IDispatch*& pOut, VOID*& valueOut, int paramType, int totalArrayParams) {

	HRESULT hr = 0;

	VARIANT result;
	VariantInit(&result);

	VARIANT vProperty;
	VariantInit(&vProperty);


	if (totalArrayParams > 0) {
		struct InputValueStruct* paramArrayStruct = (struct InputValueStruct*)param;
		// Allocate memory for arguments...
		VARIANT* pArgs = new VARIANT[totalArrayParams];

		for (int i = 0; i < totalArrayParams; i++) {


			//Integer
			if (paramArrayStruct->type == 1) {

				result.vt = VT_I4;
				vProperty.vt = VT_I4;
				vProperty.lVal = paramArrayStruct->intValue;

				pArgs[i] = vProperty;

				std::cout << "-Integer-->\0" << "\n";
			}
			//String
			if (paramArrayStruct->type == 0) {


				struct InputValueStruct* a = (struct InputValueStruct*)param;
				const char* bosta = (char*)&a[i].stringValue;

				_bstr_t bstrt(bosta);

				vProperty.vt = VT_BSTR;
				vProperty.vt = VT_BSTR;
				vProperty.bstrVal = bstrt;

				pArgs[i] = vProperty;

				std::cout << "-String-->\0" << "\n";
			}




		}//End For

		hr = XvbaInvoke(DISPATCH_PROPERTYGET | DISPATCH_METHOD, &result, pIn, pPropToCall, totalArrayParams, pArgs);
		if (FAILED(hr)) return hr;

	}
	else {



		hr = XvbaInvoke(DISPATCH_PROPERTYGET | DISPATCH_METHOD, &result, pIn, pPropToCall, 0);
		if (FAILED(hr)) return hr;


	}

	char* myCharArray = NULL;

	switch (result.vt)
	{
	case VT_I4:
		valueOut = (LONG*)result.lVal;
		break;
	case VT_BSTR:
		myCharArray = _com_util::ConvertBSTRToString(result.bstrVal);
		valueOut = myCharArray;
		break;

	}



	pOut = result.pdispVal;


	return hr;
}

HRESULT XvbaSetVal(LPCTSTR pPropToCall, IDispatch*& pIn, LPCTSTR param, int paramType) {

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
