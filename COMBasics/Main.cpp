#include <objbase.h>
#include <iostream>
#include "GenerateGUID.h"
#include "Containment.h"

inline void containment_test()
{
	IMouse* mouse = nullptr;
	HRESULT hr = Containment::Computer::Create(IID_IMOUSE, (void**)&mouse);
	if (hr == NO_ERROR)
	{
		mouse->click(1);
		mouse->scroll(10);
		mouse->Release();
	}

	IKeyboard* keyboard = nullptr;
	hr = Containment::Computer::Create(IID_IKEYBOARD, (void**)&keyboard);
	if (hr == NO_ERROR)
	{
		keyboard->pressKey(5);
		keyboard->releaseKey(10);
		keyboard->Release();
	}
}

inline void in_memory_register_Test()
{
	DWORD regNum;
	HRESULT hr = CoRegisterClassObject(
		CLSID_CCOM,
		new Containment::ComputerFactory(),
		CLSCTX_INPROC_SERVER,  //register factory for use in process local memory
		REGCLS_MULTIPLEUSE,    //allow this factory to be used multiple times
		&regNum
	);

	if (hr == S_OK) {
		std::cout << "Register:Success" << std::endl;
	}
	else
	{
		if (hr == E_INVALIDARG) {
			std::cout << "Register:Invalid Arg" << std::endl;
		}
		else if (hr == E_OUTOFMEMORY) {
			std::cout << "Register:No Memory" << std::endl;
		}
		else if (hr == E_UNEXPECTED) {
			std::cout << "Register:Unexpected" << std::endl;
		}
		return;
	}

	IMouse* mouse = nullptr;
	hr = CoCreateInstance(
		CLSID_CCOM,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_IMOUSE,
		(void**)&mouse
	);
	
	if (hr == S_OK)
	{
		mouse->click(1);
		mouse->scroll(10);
		mouse->Release();
	}
	else if (hr == REGDB_E_CLASSNOTREG) {
		std::cout << "CLSID Not Registered" << std::endl;
	}
	else if (hr == CLASS_E_NOAGGREGATION) {
		std::cout << "No Aggregation Supported" << std::endl;
	}
	else if (hr == E_NOINTERFACE) {
		std::cout << "Interface Not Supported" << std::endl;
	}
	else if (hr == E_POINTER) {
		std::cout << "Pointer was null" << std::endl;
	}

	IKeyboard* keyboard = nullptr;
	hr = CoCreateInstance(
		CLSID_CCOM,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_IKEYBOARD,
		(void**)&keyboard
	);

	if (hr == S_OK)
	{
		keyboard->pressKey(5);
		keyboard->releaseKey(10);
		keyboard->Release();
	}
	else if (hr == REGDB_E_CLASSNOTREG) {
		std::cout << "CLSID Not Registered" << std::endl;
	}
	else if (hr == CLASS_E_NOAGGREGATION) {
		std::cout << "No Aggregation Supported" << std::endl;
	}
	else if (hr == E_NOINTERFACE) {
		std::cout << "Interface Not Supported" << std::endl;
	}
	else if (hr == E_POINTER) {
		std::cout << "Pointer was null" << std::endl;
	}

	hr = CoRevokeClassObject(regNum);
	if (hr == S_OK) {
		std::cout << "Unregister:Success" << std::endl;
	}
	else
	{
		if (hr == E_INVALIDARG) {
			std::cout << "Unregister:Invalid Arg" << std::endl;
		}
		else if (hr == E_OUTOFMEMORY) {
			std::cout << "Unregister:No Memory" << std::endl;
		}
		else if (hr == E_UNEXPECTED) {
			std::cout << "Unregister:Unexpected" << std::endl;
		}
	}
}

inline void local_process_create_Test()
{
	IMouse* mouse = nullptr;
	HRESULT hr = CoCreateInstance(
		CLSID_CCOM,
		NULL,
		CLSCTX_LOCAL_SERVER, //the object is registered in another process same computer
		IID_IMOUSE,
		(void**)&mouse
	);

	if (hr == S_OK)
	{
		mouse->click(1);
		mouse->scroll(10);
		mouse->Release();
	}
	else if (hr == REGDB_E_CLASSNOTREG) {
		std::cout << "CLSID Not Registered" << std::endl;
	}
	else if (hr == CLASS_E_NOAGGREGATION) {
		std::cout << "No Aggregation Supported" << std::endl;
	}
	else if (hr == E_NOINTERFACE) {
		std::cout << "Interface Not Supported" << std::endl;
	}
	else if (hr == E_POINTER) {
		std::cout << "Pointer was null" << std::endl;
	}

	IKeyboard* keyboard = nullptr;
	hr = CoCreateInstance(
		CLSID_CCOM,
		NULL,
		CLSCTX_LOCAL_SERVER, //the object is registered in another process same computer
		IID_IKEYBOARD,
		(void**)&keyboard
	);

	if (hr == S_OK)
	{
		keyboard->pressKey(5);
		keyboard->releaseKey(10);
		keyboard->Release();
	}
	else if (hr == REGDB_E_CLASSNOTREG) {
		std::cout << "CLSID Not Registered" << std::endl;
	}
	else if (hr == CLASS_E_NOAGGREGATION) {
		std::cout << "No Aggregation Supported" << std::endl;
	}
	else if (hr == E_NOINTERFACE) {
		std::cout << "Interface Not Supported" << std::endl;
	}
	else if (hr == E_POINTER) {
		std::cout << "Pointer was null" << std::endl;
	}
}

int main()
{
  HRESULT hr = CoInitialize(NULL);
	if (hr == S_OK)
	{
		std::cout << "OK" << std::endl;
	}
	else if (hr == S_FALSE)
	{
		std::cout << "Damn";
		return -1;
	}
	else if (hr == RPC_E_CHANGED_MODE)
	{
		std::cout << "Thread model changed";
		return -1;
	}


	//==============COM CODE HERE========================//
	//containment_test();

	//in_memory_register_Test();

	local_process_create_Test();
	//=====================================================//
	CoUninitialize();

	return 0;
}