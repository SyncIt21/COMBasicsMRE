#include "TestImplementation.h"

HRESULT Computer::Create(REFIID riid, void** ppv)
{
  Computer* computer = new Computer();
  if (computer == NULL || computer->dispatch == nullptr)
    return E_OUTOFMEMORY;

  HRESULT hr = computer->QueryInterface(riid, ppv);
  if (FAILED(hr))
    delete computer;

  return hr;
}

Computer::Computer() :ref_count(0),typeInfo(nullptr),dispatch(nullptr)
{
  BSTR param1_name_1 = CStringW("button").AllocSysString();
  PARAMDATA param1 = { param1_name_1, VT_INT };
  PARAMDATA* params = { &param1 };
  BSTR method_name_1 = CStringW("click").AllocSysString();
  METHODDATA clickMethod = {
    method_name_1,
    params,
    100L,              //Unique ID for method
    0,                 //Index of method in header file
    CC_STDCALL,        //std_call convention
    1,                 //number of arguments
    DISPATCH_METHOD,   //this is a method
    VT_HRESULT         //return type is void
  };

  BSTR param1_name_2 = CStringW("amount").AllocSysString();
  param1 = { param1_name_2, VT_INT };
  params = { &param1 };
  BSTR method_name_2 = CStringW("scroll").AllocSysString();
  METHODDATA scrollMethod = {
    method_name_2,
    params,
    101L,              //Unique ID for method
    1,                 //Index of method in header file
    CC_STDCALL,        //std_call convention
    1,                 //number of arguments
    DISPATCH_METHOD,   //this is a method
    VT_HRESULT         //return type is void
  };

  BSTR param1_name_3 = CStringW("key").AllocSysString();
  param1 = { param1_name_2, VT_INT };
  params = { &param1 };
  BSTR method_name_3 = CStringW("pressKey").AllocSysString();
  METHODDATA pressMethod = {
    method_name_3,
    params,
    102L,              //Unique ID for method
    2,                 //Index of method in header file
    CC_STDCALL,        //std_call convention
    1,                 //number of arguments
    DISPATCH_METHOD,   //this is a method
    VT_HRESULT         //return type is void
  };

  BSTR param1_name_4 = CStringW("key").AllocSysString();
  param1 = { param1_name_4, VT_INT };
  params = { &param1 };
  BSTR method_name_4 = CStringW("releaseKey").AllocSysString();
  METHODDATA releaseMethod = {
    method_name_4,
    params,
    103L,              //Unique ID for method
    3,                 //Index of method in header file
    CC_STDCALL,        //std_call convention
    1,                 //number of arguments
    DISPATCH_METHOD,   //this is a method
    VT_HRESULT        //return type is void
  };

  METHODDATA methods[] = { clickMethod, scrollMethod, pressMethod, releaseMethod };
  INTERFACEDATA iData = { methods, 4 };

  HRESULT hr = CreateDispTypeInfo(
    &iData,
    LOCALE_SYSTEM_DEFAULT,
    &typeInfo
  );
  SysReleaseString(param1_name_1);
  SysReleaseString(param1_name_2);
  SysReleaseString(param1_name_3);
  SysReleaseString(param1_name_4);
  SysReleaseString(method_name_1);
  SysReleaseString(method_name_2);
  SysReleaseString(method_name_3);
  SysReleaseString(method_name_4);
  if (FAILED(hr)) { return; }

  hr = CreateStdDispatch(
    this,
    this,
    typeInfo,
    &dispatch
  );
  if (FAILED(hr)) { dispatch = nullptr; }
}


HRESULT __stdcall Computer::GetTypeInfoCount(UINT* pctinfo)
{
  return E_NOTIMPL;
}
HRESULT __stdcall Computer::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
  return E_NOTIMPL;
}
HRESULT __stdcall Computer::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
  return E_NOTIMPL;
}
HRESULT __stdcall Computer::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
  return E_NOTIMPL;
}

HRESULT Computer::click(int button) { return mouse_implementation.click(button); }
HRESULT Computer::scroll(int amount) { return mouse_implementation.scroll(amount); }
HRESULT Computer::pressKey(int key) { return keyboard_implementation.pressKey(key); }
HRESULT Computer::releaseKey(int key) { return keyboard_implementation.releaseKey(key); }

HRESULT __stdcall Computer::QueryInterface(REFIID riid, void** ppvObject)
{
  if (riid == IID_IUnknown)
    *ppvObject = (IUnknown*)this;
  else if (riid == IID_IMOUSE)
    *ppvObject = (IMouse*)&mouse_implementation;
  else if (riid == IID_IKEYBOARD)
    *ppvObject = (IKeyboard*)&keyboard_implementation;
  else if (riid == IID_IDispatch)
    return dispatch->QueryInterface(IID_IDispatch, ppvObject);
  else
  {
    *ppvObject = NULL;
    return ResultFromScode(E_NOINTERFACE);
  }
   
  ((IUnknown*)*ppvObject)->AddRef();

  return S_OK;
}

ULONG __stdcall Computer::AddRef(void)
{
  ref_count++;
  return ref_count;
}
ULONG __stdcall Computer::Release(void)
{
  ref_count--;
  if (ref_count > 0)
    return ref_count;

  delete this;
  return 0;
}

Computer::~Computer()
{
  if (typeInfo != nullptr) {
    typeInfo->Release();
  }
  if (dispatch != nullptr) {
    dispatch->Release();
  }
}


HRESULT __stdcall Computer::MouseImp::click(int button)
{
  std::cout << "Click " << button << std::endl;

  return S_OK;
}
HRESULT __stdcall Computer::MouseImp::scroll(int amount)
{
  std::cout << "Scroll " << amount << std::endl;

  return S_OK;
}
HRESULT __stdcall Computer::MouseImp::QueryInterface(REFIID riid, void** ppvObject)
{
  return E_NOTIMPL;
}
ULONG __stdcall Computer::MouseImp::AddRef(void)
{
  return 1;
}
ULONG __stdcall Computer::MouseImp::Release(void)
{
  return 1;
}

HRESULT __stdcall Computer::KeyboardImp::pressKey(int key)
{
  std::cout << "Key Pressed " << key << std::endl;

  return S_OK;
}
HRESULT __stdcall Computer::KeyboardImp::releaseKey(int key)
{
  std::cout << "Key Released " << key << std::endl;

  return S_OK;
}
HRESULT __stdcall Computer::KeyboardImp::QueryInterface(REFIID riid, void** ppvObject)
{
  return E_NOTIMPL;
}
ULONG __stdcall Computer::KeyboardImp::AddRef(void)
{
  return 1;
}
ULONG __stdcall Computer::KeyboardImp::Release(void)
{
  return 1;
}

HRESULT __stdcall ComputerFactory::QueryInterface(REFIID riid, void** ppvObject)
{
  if ((riid == IID_IClassFactory) || (riid == IID_IUnknown))
  {
    *ppvObject = this;
    AddRef();
    return S_OK;
  }

  return E_NOINTERFACE;
}
ULONG __stdcall ComputerFactory::AddRef(void)
{
  return 1;
}
ULONG __stdcall ComputerFactory::Release(void)
{
  return 1;
}

HRESULT __stdcall ComputerFactory::CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject)
{
  if (pUnkOuter != NULL) {
    return CLASS_E_NOAGGREGATION;
  }

  return Computer::Create(riid, ppvObject);
}

HRESULT __stdcall ComputerFactory::LockServer(BOOL fLock)
{
  return S_OK;
}