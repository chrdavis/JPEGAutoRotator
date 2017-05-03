#pragma once

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files:
#include <windows.h>
#include <unknwn.h>
#include <shlwapi.h>
#include <atlbase.h>
#include <Shobjidl.h>
#include <Shlobj.h>

void DllAddRef();
void DllRelease();

#define INITGUID
#include <guiddef.h>

// {1AC8EAF1-9A6A-471C-B87B-B1DF2B8EA62F}
DEFINE_GUID(CLSID_JPEGAutoRotatorMenu, 0x1AC8EAF1, 0x9A6A, 0x471C, 0xB8, 0x7B, 0xB1, 0xDF, 0x2B, 0x8E, 0xA6, 0x2F);


