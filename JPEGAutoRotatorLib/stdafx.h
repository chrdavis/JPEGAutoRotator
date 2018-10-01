#pragma once

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
#include <windows.h>
#include <unknwn.h>
#include <shlwapi.h>
#include <atlbase.h>
#include <Shobjidl.h>
#include <Shlobj.h>
#pragma warning(push, 3)   // Ignore warnings from GDIPlus headers
#include <gdiplus.h>
#pragma warning(pop)
#include <strsafe.h>
#include <pathcch.h>

