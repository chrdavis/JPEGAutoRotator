// stdafx.cpp : source file that includes just the standard includes
// JPEGAutoRotatorUnitTests.pch will be the pre-compiled header
// stdafx.obj will contain the pre-compiled type information

#include "stdafx.h"

// TODO: reference any additional headers you need in STDAFX.H
// and not in this file

HINSTANCE g_hInstance;

BOOL WINAPI DllMain(_In_ HINSTANCE hInstance, _In_ DWORD, _In_ void*)
{
    g_hInstance = hInstance;
    return TRUE;
}

bool GetCurrentFolderPath(_In_ UINT count, _Out_ PWSTR path)
{
    bool ret = false;
    if (GetModuleFileName(g_hInstance, path, count))
    {
        PathRemoveFileSpec(path);
        ret = true;
    }
    return ret;
}