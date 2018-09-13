#pragma once

// Windows Header Files:
#include <windows.h>
#include <unknwn.h>
#include <shlwapi.h>
#include <atlbase.h>
#include <Shobjidl.h>
#include <Shlobj.h>

#include "targetver.h"

#include "CppUnitTest.h"

bool GetCurrentFolderPath(_In_ UINT count, _Out_ PWSTR path);

void DllAddRef()
{
}

void DllRelease()
{
}