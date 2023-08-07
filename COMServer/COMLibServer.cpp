// COMLibServer.cpp : Defines the entry point for the application.
//

#include "COMLibServer.h"

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    if(FAILED(CoInitialize(NULL))){ return FALSE; }
   
    HRESULT hr;

    delete new Computer();

    int testMode = 0;
    DWORD regNum;
    if (testMode == 0)
    {
      hr = CoRegisterClassObject(
        CLSID_CCOM,
        new ComputerFactory(),
        CLSCTX_LOCAL_SERVER,  //register factory for use out of process in same computer memory
        REGCLS_MULTIPLEUSE,    //allow this factory to be used multiple times
        &regNum
      );
      if (FAILED(hr)) {return FALSE;}
    }

    MSG msg = {};
    // Main message loop:
    while (GetMessage(&msg, 0, 0, 0))
    {  
      TranslateMessage(&msg);
      DispatchMessage(&msg);
    }

    hr = CoRevokeClassObject(regNum);
    if (FAILED(hr)){return FALSE;}

    CoUninitialize();

    return (int) msg.wParam;
}
