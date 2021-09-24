#include "XvbaCom.h"
#include <iostream>
#include "windows.h"



int XvbaImportVBA(LPCTSTR szExcelFileName, LPCTSTR localToExportVbaFiles) {


    HRESULT hr;
    IDispatch* pExcelCOM = (IDispatch*)NULL;
    IDispatch* pWorkbook = (IDispatch*)NULL;
    IDispatch* pVBAComponents = (IDispatch*)NULL;
    LPCTSTR pCount = (LPCTSTR)NULL;


    try
    {
        //Create COM Instance
        hr = XvbaCoCreateInstance(L"Excel.Application", pExcelCOM);
        //Show Com
        hr = XvbaShowApplication(pExcelCOM); 
        //Open Document
        hr = XvbaOpenDocument(szExcelFileName,pExcelCOM,pWorkbook);
       //VBAComponents
        hr = XvbaGetVBComponets(pWorkbook, pVBAComponents);



    }
    catch (const std::exception&)
    {
        pExcelCOM->Release();
        return hr;
    }


    return   hr;

}


