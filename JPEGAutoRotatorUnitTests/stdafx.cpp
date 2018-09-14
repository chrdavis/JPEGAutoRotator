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

bool CreateTestFolder(_In_ PCWSTR testFolderPathAppend, _In_ UINT testFolderPathLen, _Out_ PWSTR testFolderPath)
{
    bool ret = false;
    if (GetCurrentFolderPath(testFolderPathLen, testFolderPath))
    {
        if (SUCCEEDED(PathCchAppend(testFolderPath, testFolderPathLen, testFolderPathAppend)))
        {
            if (PathFileExists(testFolderPath))
            {
                DeleteHelper(testFolderPath);
            }

            ret = CreateDirectory(testFolderPath, NULL);
        }
    }

    return ret;
}

bool DeleteHelper(_In_ PCWSTR path)
{
    // Delete test folder
    SHFILEOPSTRUCT fileStruct = { 0 };
    fileStruct.pFrom = path;
    fileStruct.wFunc = FO_DELETE;
    fileStruct.fFlags = FOF_SILENT | FOF_NOERRORUI | FOF_NO_UI;
    return (SHFileOperation(&fileStruct) == 0);
}