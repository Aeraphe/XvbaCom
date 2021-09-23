#include "XvbaCom.h"
#include <iostream>
#include "windows.h"


HRESULT XvbaImportVBA(LPCTSTR szFilename) {

	
	HRESULT hr;
	IDispatch* pExcelCOM = (IDispatch*)NULL;
    IDispatch* pWorkbook = (IDispatch*)NULL;


    hr = XvbaCoCreateInstance(L"Excel.Application",  pExcelCOM);

    if (FAILED(hr)) {
        return hr;
    }

    //Open 
    hr = OpenDocument(szFilename, pExcelCOM, pWorkbook);

    //Access VBA

    if (FAILED(hr)) {
        return hr;
      
    }

    //Close COM
    //app->Release();
	
    return hr;


}



HRESULT OpenDocument(LPCTSTR szFilename,  IDispatch* app, IDispatch* pWorkbook)
{

    HRESULT hr;


    VARIANT fname;
    VariantInit(&fname);
    fname.vt = VT_BSTR;
    fname.bstrVal = SysAllocString(szFilename);
  
 
    IDispatch* pDocuments;
   

    // GetDocuments
    {
        VARIANT result;
        VariantInit(&result);
        hr = XvbaInvoke(DISPATCH_PROPERTYGET, &result, app, L"Workbooks", 0);
        pDocuments = result.pdispVal;
    }
    // OpenDocument
    {
        VARIANT result;
        VariantInit(&result);
        hr = XvbaInvoke(DISPATCH_METHOD, &result, pDocuments,L"Open", 1, fname);
        pWorkbook = result.pdispVal;
    }
    return hr;
}


HRESULT XvbaCoCreateInstance(LPCOLESTR lpszProgId, IDispatch* app) {

    CLSID clsId;

    HRESULT hr;

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


error:
    return hr;
}



HRESULT XvbaInvoke(int nType, VARIANT* pvResult, IDispatch* pDisp, LPCTSTR propertyName, int  cArgs... )
{
    if (!pDisp) return E_FAIL;

 
    va_list marker;
    va_start(marker, cArgs);

    //https://docs.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-dispparams
    DISPPARAMS dp = { NULL, NULL, 0, 0 };   
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    DISPID dispID;
    char szName[256];

    std::wstring str = propertyName;
    LPOLESTR ptName = (LPOLESTR) new wchar_t[str.length() + 1];

    // Convert down to ANSI
    WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);

    // Get DISPID for name passed...
    HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1,
        LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        return hr;
    }
    // Allocate memory for arguments...
    VARIANT* pArgs = new VARIANT[cArgs + 1];
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
    hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT,nType, &dp, pvResult, NULL, NULL);
    if (FAILED(hr)) {
        return hr;
    }

    // End variable-argument section...
    va_end(marker);
    delete[] pArgs;
    return hr;


}

