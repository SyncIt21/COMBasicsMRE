#include <Initguid.h>
#include "GenerateGUID.h"

void printAll()
{
	const GUID data[] = {
		CLSID_CCOM,
		IID_IMOUSE,
		IID_IKEYBOARD
	};

	for (int i = 0; i < 3; i++)
	{
		GUID guid = data[i];

		printf("Guid = {%08lX-%04hX-%04hX-%02hX%02hX-%02hX%02hX%02hX%02hX%02hX%02hX}\n",
			guid.Data1, guid.Data2, guid.Data3,
			guid.Data4[0], guid.Data4[1], guid.Data4[2], guid.Data4[3],
			guid.Data4[4], guid.Data4[5], guid.Data4[6], guid.Data4[7]);
	}
}

void generateAll()
{
	const std::string data[] = {
		"CLSID_CCOM",
		"IID_IMOUSE",
		"IID_IKEYBOARD"
	};

	GUID guid;
	for (int i = 0; i < 3; i++) 
	{
		if (SUCCEEDED(CoCreateGuid(&guid)))
		{
			std::string name = data[i];

			std::cout << name << std::endl;
			std::cout << "{" << std::endl;
			printf("DEFINE_GUID(%s, 0x%08lX, 0x%04hX, 0x%04hX, 0x%02hX, 0x%02hX, 0x%02hX, 0x%02hX, 0x%02hX, 0x%02hX, 0x%02hX, 0x%02hX);\n",
				name.c_str(),
				guid.Data1, guid.Data2, guid.Data3,
				guid.Data4[0], guid.Data4[1], guid.Data4[2], guid.Data4[3],
				guid.Data4[4], guid.Data4[5], guid.Data4[6], guid.Data4[7]);
			printf("Guid = {%08lX-%04hX-%04hX-%02hhX%02hhX-%02hhX%02hhX%02hhX%02hhX%02hhX%02hX}\n",
				guid.Data1, guid.Data2, guid.Data3,
				guid.Data4[0], guid.Data4[1], guid.Data4[2], guid.Data4[3],
				guid.Data4[4], guid.Data4[5], guid.Data4[6], guid.Data4[7]);
			std::cout << "}" << std::endl;
		}
	}
}
