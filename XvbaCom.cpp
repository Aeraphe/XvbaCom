#include "XvbaCom.h"
#include <iostream>
#include "windows.h"


HRESULT XvbaInvoke(int nType, VARIANT* pvResult, IDispatch* pDisp, LPOLESTR ptName, int cArgs...)
{
    if (!pDisp) return E_FAIL;

    va_list marker;
    va_start(marker, cArgs);

    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    DISPID dispID;
    char szName[256];

    // Convert down to ANSI
    WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);

    // Get DISPID for name passed...
    HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1,
        LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        return hr;
    }
    // Allocate memory for arguments...
    VARIANT* pArgs = new VARIANT[static_cast<unsigned __int64>(cArgs) + 1];
    // Extract arguments...
    for (int i = 0; i < cArgs; i++) {
        pArgs[i] = va_arg(marker, VARIANT);
    }

    // Build DISPPARAMS
    dp.cArgs = cArgs;
    dp.rgvarg = pArgs;

    // Handle special-case for property-puts!
    if (nType & DISPATCH_PROPERTYPUT) {
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
    }

    // Make the call!
    hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT,
        nType, &dp, pvResult, NULL, NULL);
    if (FAILED(hr)) {
        return hr;
    }

    // End variable-argument section...
    va_end(marker);
    delete[] pArgs;
    return hr;
}


HRESULT DoSomething(LPCOLESTR lpszProgId,OLECHAR* prop) {

	

	CLSID clsId;

	HRESULT hr;

	IDispatch* app = (IDispatch*)NULL;


	VARIANT x;
	x.vt = VT_I4;
	x.lVal = 1;

	DISPID propID;


	
	unsigned returnArg;

	VARIANT varTrue;
	DISPID rgDispidNamedArgs[1];
	rgDispidNamedArgs[0] = DISPID_PROPERTYPUT;
	DISPPARAMS params = { &varTrue, rgDispidNamedArgs, 1, 1 };

	varTrue.vt = VT_I4;
	varTrue.lVal = 1;

	hr = CoInitialize(NULL);
		 
	if (FAILED(hr)) {

		printf("Fail ConInitialize - [%x]\n", hr);
		goto error;
	}

	hr = CLSIDFromProgID(lpszProgId, &clsId);

	if (FAILED(hr)) {
		printf("Fail CLSIDFromProgID - [%x]\n", hr);
		goto error;
	}


	hr = CoCreateInstance(clsId, NULL, CLSCTX_SERVER, IID_IDispatch, (void**)&app);

	if (FAILED(hr)) {
		printf("Fail CLSIDFromProgID - [%x]\n", hr);
		goto error;
	}




	hr = XvbaInvoke(DISPATCH_PROPERTYPUT, NULL, app, prop, 1, x);

  hr =	app->GetIDsOfNames(IID_NULL, &prop, 1, LOCALE_SYSTEM_DEFAULT, &propID);

  if (FAILED(hr)) {
	  printf("Fail GetIDsOfNames - [%x]\n", hr);
	  goto error;
  }

  
  app->Invoke(propID, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT, &params, NULL, NULL, NULL);




	return hr;

error:
	return hr;
}



