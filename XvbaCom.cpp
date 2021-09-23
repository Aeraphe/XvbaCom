#include "XvbaCom.h"
#include "XvbaInvoke.h"
#include <iostream>
#include "windows.h"



enum XVBA_ERROR {

    CO_CREATE_INSTANCE = -141,
    OPEN_DOCUMENT = -142,
    INVOKE = -143,
    GET_WORKBOOK = -144,

};

int XvbaImportVBA(LPCTSTR szFilename) {

	
	HRESULT hr;
	IDispatch *pExcelCOM = (IDispatch*)NULL;
    IDispatch *pWorkbook = (IDispatch*)NULL;
   

    try
    {
        //Create COM Instance
        hr = XvbaCoCreateInstance(L"Excel.Application", pExcelCOM);
        //Show Com
        hr = XvbaShowApplication(pExcelCOM);
        //Open Document
       // hr = XvbaOpenDocument(szFilename,pExcelCOM,pWorkbook);
        
    }
    catch (const std::exception&)
    {
        pExcelCOM->Release();
        return hr;
    }
   
	
    return hr;

}


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




