#pragma once
#include "windows.h"


extern "C" __declspec(dllexport) HRESULT  DoSomething(LPCOLESTR lpszProgId, OLECHAR * prop);

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
extern "C" __declspec(dllexport) HRESULT XvbaInvoke(int nType, VARIANT * pvResult, IDispatch * pDisp, LPOLESTR ptName, int cArgs...);