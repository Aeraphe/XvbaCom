#include "XvbaCom.h"
#include "windows.h"
#include <iostream>

enum XVBA_ERROR {

    CO_CREATE_INSTANCE = -141,
    OPEN_DOCUMENT = -142,
    INVOKE = -143,
    GET_WORKBOOK = -144,

};


int XvbaShowApplication(IDispatch * &app) {

    HRESULT hr;
    VARIANT x;
    x.vt = VT_I4;
    x.lVal = 1;
    hr = XvbaInvoke(DISPATCH_PROPERTYPUT, NULL, app, L"Visible", 1, x);

    if(FAILED(hr)){
       return hr;
    }

    return hr;
}

int XvbaOpenDocument(LPCTSTR szFilename,  IDispatch * &app, IDispatch * &pWorkbook)
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

        if (FAILED(hr)) {
            return hr;
        }
    }
  
    // OpenDocument
    {
        VARIANT result;
        VariantInit(&result);
        hr = XvbaInvoke(DISPATCH_METHOD, &result, pDocuments, L"Open", 1, fname);
        pWorkbook = result.pdispVal;

        if (FAILED(hr)) {

            return XVBA_ERROR::INVOKE;

        }
    }
    return hr;
}

HRESULT XvbaCoCreateInstance(LPCOLESTR lpszProgId, IDispatch * &app) {

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

}




