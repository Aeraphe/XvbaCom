#include "XvbaCom.h"
#include <iostream>
#include "windows.h"



int XvbaImportVBA(LPCTSTR szFilename) {


    HRESULT hr;
    IDispatch* pExcelCOM = (IDispatch*)NULL;
    IDispatch* pWorkbook = (IDispatch*)NULL;
    IDispatch* pVBAComponents = (IDispatch*)NULL;


    try
    {
        //Create COM Instance
        hr = XvbaCoCreateInstance(L"Excel.Application", pExcelCOM);
        //Show Com
        hr = XvbaShowApplication(pExcelCOM); 
        //Open Document
        hr = XvbaOpenDocument(szFilename,pExcelCOM,pWorkbook);
       //VBAComponents
        hr = XvbaGetVBComponets(pWorkbook, pVBAComponents);


    }
    catch (const std::exception&)
    {
        pExcelCOM->Release();
        return hr;
    }


    return hr;

}


