
#include <iostream>
#include "windows.h"
#include "XvbaInvoke.h";

int XvbaInvoke(int propertyTypeFlag, VARIANT* pInvokeResultVariant, IDispatch*& pDisp, LPCTSTR propertyName, int  cArgs...)
{

    HRESULT hr;

    va_list marker;
    va_start(marker, cArgs);

    //https://docs.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-dispparams
    DISPPARAMS propertyParams = { NULL, NULL, 0, 0 };
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    DISPID propertyDISPID;
    char szName[256];

    LPOLESTR ptName = const_cast <wchar_t*>(propertyName);

    // Convert down to ANSI
    WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);

    // Get DISPID for name passed...
    hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1,
        LOCALE_USER_DEFAULT, &propertyDISPID);

    if (FAILED(hr)) {
        return -25;
    }
    // Allocate memory for arguments...
    uint8_t  variantSize = cArgs;

    VARIANT* pArgs = new VARIANT[++variantSize];
    // Extract arguments...
    for (int i = 0; i < cArgs; i++) {
        pArgs[i] = va_arg(marker, VARIANT);
    }

    // Build DISPPARAMS
    propertyParams.cArgs = cArgs;
    propertyParams.rgvarg = pArgs;

    // Handle special-case for property-puts!
    if (propertyDISPID & DISPATCH_PROPERTYPUT) {
        propertyParams.cNamedArgs = 1;
        propertyParams.rgdispidNamedArgs = &dispidNamed;
    }

    // Make the call!
    hr = pDisp->Invoke(propertyDISPID, IID_NULL, LOCALE_SYSTEM_DEFAULT, propertyTypeFlag, &propertyParams, pInvokeResultVariant, NULL, NULL);
    if (FAILED(hr)) {
        return hr;
    }

    // End variable-argument section...
    va_end(marker);
    delete[] pArgs;
    return hr;


}