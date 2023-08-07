#ifndef CONTAINMENT
#define CONTAINMENT

#include <Unknwn.h>
#include <Unknwnbase.h>
#include <iostream>
#include "TestInterfaces.h"

/*
* In containment the lifecycle of the innerobject ends when its outer(containing) object is destroyed
*/

namespace Containment
{
	class Computer : public IMouse, public IKeyboard
	{
	private:
		ULONG ref_count = 0;

	public:
		// Static member function for creating new instances (don't use 
		// new directly). Protects against outer objects asking for 
		// interfaces other than Iunknown. 
		static HRESULT Create(REFIID riid, void** ppv);

		Computer();

		~Computer();

		virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override;

	private:
		class MouseImp : public IMouse
		{
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
	  	public:
				HRESULT __stdcall pressKey(int key) override;

				HRESULT __stdcall releaseKey(int key) override;

				virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override;

				virtual ULONG __stdcall AddRef(void) override;

				virtual ULONG __stdcall Release(void) override;
		};
		KeyboardImp keyboard_implementation;

		// Inherited via IMouse
		virtual HRESULT click(int button) override;
		virtual HRESULT scroll(int amount) override;

		// Inherited via IKeyboard
		virtual HRESULT pressKey(int key) override;
		virtual HRESULT releaseKey(int key) override;

		virtual ULONG __stdcall AddRef(void) override;
		virtual ULONG __stdcall Release(void) override;
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
};

#endif CONTAINMENT