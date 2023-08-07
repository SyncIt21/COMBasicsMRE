#ifndef TESTIMPLEMENTATION
#define TESTIMPLEMENTATION

#include "TestInterfaces.h"
#include <oleauto.h>
#include <atlstr.h>

class Computer : public IDispatch
{
	ULONG ref_count = 0;
	ITypeInfo* typeInfo;
	IUnknown* dispatch = nullptr;

	//IUnknown
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override;
	virtual ULONG __stdcall AddRef(void) override;
	virtual ULONG __stdcall Release(void) override;
	
	// Inherited via IDispatch
	virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) override;
	virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) override;
	virtual HRESULT __stdcall GetTypeInfoCount(UINT* pctinfo) override;
	virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr) override;


	// Static member function for creating new instances (don't use 
	// new directly). Protects against outer objects asking for 
	// interfaces other than Iunknown. 
public:
	static HRESULT Create(REFIID riid, void** ppv);

	Computer();

	virtual HRESULT click(int button);
	virtual HRESULT scroll(int amount);
	virtual HRESULT pressKey(int key);
	virtual HRESULT releaseKey(int key);

	~Computer();

private:
	class MouseImp : public IMouse
	{
	private:
		ULONG ref_count = 0;

	public:
		HRESULT __stdcall click(int button) override;

		HRESULT __stdcall scroll(int amount) override;

		virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override;

		virtual ULONG __stdcall AddRef(void) override;

		virtual ULONG __stdcall Release(void) override;
	};
	MouseImp mouse_implementation;

	class KeyboardImp : public IKeyboard
	{
	private:
		ULONG ref_count = 0;

	public:
		HRESULT __stdcall pressKey(int key) override;

		HRESULT __stdcall releaseKey(int key) override;

		virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override;

		virtual ULONG __stdcall AddRef(void) override;

		virtual ULONG __stdcall Release(void) override;
	};
	KeyboardImp keyboard_implementation;
};

class ComputerFactory : public IClassFactory
{
	ULONG ref_count = 0;

	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override;
	virtual ULONG __stdcall AddRef(void) override;
	virtual ULONG __stdcall Release(void) override;
	virtual HRESULT __stdcall CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject) override;
	virtual HRESULT __stdcall LockServer(BOOL fLock) override;
};

#endif TESTIMPLEMENTATION
