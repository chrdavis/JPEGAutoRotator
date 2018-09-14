#include "stdafx.h"

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

bool GetTestFolderPath(_In_ PCWSTR testFolderPathAppend, _In_ UINT testFolderPathLen, _Out_ PWSTR testFolderPath)
{
    bool ret = false;
    if (GetCurrentFolderPath(testFolderPathLen, testFolderPath))
    {
        if (SUCCEEDED(PathCchAppend(testFolderPath, testFolderPathLen, testFolderPathAppend)))
        {
            ret = true;
        }
    }

    return ret;
}

bool CopyHelper(_In_ PCWSTR src, _In_ PCWSTR dest)
{
    SHFILEOPSTRUCT fileStruct = { 0 };
    fileStruct.pFrom = src;
    fileStruct.pTo = dest;
    fileStruct.wFunc = FO_COPY;
    fileStruct.fFlags = FOF_SILENT | FOF_NOERRORUI | FOF_NO_UI | FOF_NOCONFIRMMKDIR | FOF_NOCONFIRMATION;
    return (SHFileOperation(&fileStruct) == 0);
}

bool DeleteHelper(_In_ PCWSTR path)
{
    SHFILEOPSTRUCT fileStruct = { 0 };
    fileStruct.pFrom = path;
    fileStruct.wFunc = FO_DELETE;
    fileStruct.fFlags = FOF_SILENT | FOF_NOERRORUI | FOF_NO_UI;
    return (SHFileOperation(&fileStruct) == 0);
}