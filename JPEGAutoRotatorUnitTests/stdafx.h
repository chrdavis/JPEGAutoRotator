#pragma once

// Windows Header Files:
#include <windows.h>
#include <unknwn.h>
#include <shlwapi.h>
#include <atlbase.h>
#include <Shobjidl.h>
#include <Shlobj.h>
#include <PathCch.h>
#include "targetver.h"

#include "CppUnitTest.h"

// Unit test helper functions
bool GetCurrentFolderPath(_In_ UINT count, _Out_ PWSTR path);
bool GetTestFolderPath(_In_ PCWSTR testFolderPathAppend, _In_ UINT testFolderPathLen, _Out_ PWSTR testFolderPath);

bool CopyHelper(_In_ PCWSTR src, _In_ PCWSTR dest);
bool DeleteHelper(_In_ PCWSTR path);

void DllAddRef()
{
}

void DllRelease()
{
}