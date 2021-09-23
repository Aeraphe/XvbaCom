#pragma once
#include "windows.h"



extern "C" __declspec(dllexport) int XvbaInvoke(int propertyTypeFlag, VARIANT * pInvokeResultVariant, IDispatch * &pDisp, LPCTSTR propertyName, int  cArgs...);
