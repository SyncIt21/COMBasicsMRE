#include "TestInterfaces.h"

COMObject::COMObject(){}

ULONG __stdcall COMObject::AddRef(void)
{
	ref_count++;
	return ref_count;
}

ULONG __stdcall COMObject::Release(void)
{
	ref_count--;
	if (ref_count > 0)
		return ref_count;

	delete this;
	return 0;
}

COMObject::~COMObject()
{
}
