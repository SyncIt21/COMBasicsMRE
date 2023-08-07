#ifndef TESTINTERFACES
#define TESTINTERFACES

#include <Unknwn.h>
#include "GenerateGUID.h"
#include <iostream>

class COMObject : public IUnknown
{
protected:
	ULONG ref_count = 0;

public:
	COMObject();

	virtual ULONG __stdcall AddRef(void) override;

	virtual ULONG __stdcall Release(void) override;

	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) = 0;

	virtual ~COMObject();
};

__interface IMouse : public IUnknown
{
public:
	HRESULT click(int button);
	HRESULT scroll(int amount);
};

__interface IKeyboard : public IUnknown
{
public:
	HRESULT pressKey(int key);
	HRESULT releaseKey(int key);
};

#endif TESTINTERFACES
