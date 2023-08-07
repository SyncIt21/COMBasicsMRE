#include "Containment.h"

using namespace Containment;

HRESULT Computer::Create(REFIID riid, void** ppv)
{
	Computer* computer = new Computer();
	if (computer == NULL)
		return E_OUTOFMEMORY;

	HRESULT hr = computer->QueryInterface(riid, ppv);
	if (FAILED(hr))
		delete computer;
	
	return hr;
}

Computer::Computer():ref_count(0)
{}

Computer::~Computer(){std::cout << "END" << std::endl;}
			
HRESULT __stdcall Computer::QueryInterface(REFIID riid, void** ppvObject)
{
	if (riid == IID_IUnknown)
		*ppvObject = this;
	else if (riid == IID_IMOUSE)
		*ppvObject = (IMouse*)this;
	else if (riid == IID_IKEYBOARD)
		*ppvObject = (IKeyboard*)this;

	if (NULL == *ppvObject)
		return ResultFromScode(E_NOINTERFACE);
	((IUnknown*)*ppvObject)->AddRef();

	return S_OK;
}

HRESULT Computer::click(int button) { return mouse_implementation.click(button); }
HRESULT Computer::scroll(int amount) { return mouse_implementation.scroll(amount); }
HRESULT Computer::pressKey(int key) { return keyboard_implementation.pressKey(key); }
HRESULT Computer::releaseKey(int key) { return keyboard_implementation.releaseKey(key); }
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
ULONG __stdcall Computer::MouseImp::AddRef(void) { return 0; }
ULONG __stdcall Computer::MouseImp::Release(void) { return  0; }

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
ULONG __stdcall Computer::KeyboardImp::AddRef(void) { return 0; }
ULONG __stdcall Computer::KeyboardImp::Release(void) { return  0; }

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
	ref_count++;
	return ref_count;
}
ULONG __stdcall ComputerFactory::Release(void)
{
	ref_count--;
	if (ref_count > 0) {
		return ref_count;
	}

	delete this;
	return 0;
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
