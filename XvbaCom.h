#pragma once
#include "windows.h"
#include <iostream>


//https://www.codeproject.com/Articles/2995/The-Complete-Guide-to-C-Strings-Part-I-Win32-Chara

extern "C" __declspec(dllexport) int  XvbaImportVBA(LPCTSTR szExcelFileName, LPCTSTR localToExportVbaFiles);

/*
 Imvoke methos or Prop from COMs

 @param int       ntype :    Flag enum property for define if the invoke will be for Method, insert/get Property (DISPATCH_PROPERTYGET)
 @param VARIANT   pvResult:  A pointer to the result off invoke call
 @param IDispatch pDisp :    COM that caller to invoke
 @param LPOLESTR  ptName:    Prop/Method Name
 @param int       cArgs:     Arguments to pass

 [Example]
 VARIANT x;
 x.vt = VT_I4;
 x.lVal = 1;        // 1=visible; 0=invisible;
 hr=OLEMethod(DISPATCH_PROPERTYPUT, NULL, pWApp, L"Visible", 1, x);

*/
extern "C" __declspec(dllexport) HRESULT XvbaInvoke(int nType, VARIANT * pvResult, IDispatch * pDisp, LPCTSTR ptName, int cArgs...);

extern "C" __declspec(dllexport) HRESULT XvbaCoCreateInstance(LPCOLESTR lpszProgId, IDispatch * &app);

extern "C" __declspec(dllexport)  HRESULT XvbaGetMethod(IDispatch*& pIn, IDispatch*& pOut, LPCTSTR pMenthodName);

extern "C" __declspec(dllexport) HRESULT XvbaCall(LPCTSTR pPropToCall, IDispatch * &pIn, VOID * &param, IDispatch * &pOut, VOID * &sValue, int paramType, int totalArrayParams);

extern "C" __declspec(dllexport) HRESULT XvbaSetVal(LPCTSTR pPropToCall, IDispatch * &pIn, LPCTSTR param, int paramType);

extern "C" __declspec(dllexport) HRESULT XvbaRelease(IDispatch * &pIn);

extern "C" __declspec(dllexport) HRESULT XvbaGetPropertyRef(IDispatch*& pIn, IDispatch*& pOut, LPCTSTR pMenthodName);

std::wstring ConvertBSTRToMBS(BSTR bstr);

std::wstring ConvertWCSToMBS(const wchar_t* pstr, long wslen);

BSTR ConvertMBSToBSTR(const std::string& str);

