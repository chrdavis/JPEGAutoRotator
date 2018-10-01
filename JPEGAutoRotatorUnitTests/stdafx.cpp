#include "stdafx.h"
#include <strsafe.h>

HINSTANCE g_hInst;

BOOL WINAPI DllMain(_In_ HINSTANCE hInstance, _In_ DWORD, _In_ void*)
{
    g_hInst = hInstance;
    return TRUE;
}

bool GetCurrentFolderPath(_In_ UINT count, _Out_ PWSTR path)
{
    bool ret = false;
    if (GetModuleFileName(g_hInst, path, count))
    {
        PathRemoveFileSpec(path);
        ret = true;
    }
    return ret;
}

bool GetTestFolderPath(_In_ PCWSTR testFolderPathAppend, _In_ UINT testFolderPathLen, _Out_ PWSTR testFolderPath)
{
    bool ret = false;
    if (GetCurrentFolderPath(testFolderPathLen, testFolderPath))
    {
        if (SUCCEEDED(PathCchAppend(testFolderPath, testFolderPathLen, testFolderPathAppend)))
        {
            PathCchAddBackslash(testFolderPath, testFolderPathLen);
            ret = true;
        }
    }

    return ret;
}

bool CopyHelper(_In_ PCWSTR src, _In_ PCWSTR dest)
{
    // Ensure paths are double null terminated
    wchar_t srcLocal[MAX_PATH + 2] = { 0 };
    wchar_t destLocal[MAX_PATH + 2] = { 0 };
    StringCchCopy(srcLocal, ARRAYSIZE(srcLocal), src);
    StringCchCopy(destLocal, ARRAYSIZE(destLocal), dest);
    SHFILEOPSTRUCT fileStruct = { 0 };
    fileStruct.pFrom = srcLocal;
    fileStruct.pTo = destLocal;
    fileStruct.wFunc = FO_COPY;
    fileStruct.fFlags = FOF_SILENT | FOF_NOERRORUI | FOF_NO_UI | FOF_NOCONFIRMMKDIR | FOF_NOCONFIRMATION;
    return (SHFileOperation(&fileStruct) == 0);
}

bool DeleteHelper(_In_ PCWSTR path)
{
    // Ensure paths are double null terminated
    wchar_t pathLocal[MAX_PATH + 2] = { 0 };
    StringCchCopy(pathLocal, ARRAYSIZE(pathLocal), path);
    SHFILEOPSTRUCT fileStruct = { 0 };
    fileStruct.pFrom = pathLocal;
    fileStruct.wFunc = FO_DELETE;
    fileStruct.fFlags = FOF_SILENT | FOF_NOERRORUI | FOF_NO_UI | FOF_NOCONFIRMATION;
    return (SHFileOperation(&fileStruct) == 0);
}