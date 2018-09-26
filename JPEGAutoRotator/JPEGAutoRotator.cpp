// JPEGAutoRotator.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include "pch.h"

bool GetCurrentFolderPath(_In_ UINT count, _Out_ PWSTR path)
{
    bool ret = false;
    if (GetModuleFileName(nullptr, path, count))
    {
        PathRemoveFileSpec(path);
        ret = true;
    }
    return ret;
}

HRESULT CreateDataObjectFromPath(_In_ PCWSTR pathToFile, _COM_Outptr_ IDataObject** dataObject)
{
    *dataObject = nullptr;

    CComPtr<IShellItem> spsi;
    HRESULT hr = SHCreateItemFromParsingName(pathToFile, nullptr /* IBindCtx */, IID_PPV_ARGS(&spsi));
    if (SUCCEEDED(hr))
    {
        hr = spsi->BindToHandler(nullptr /* IBindCtx */, BHID_DataObject, IID_PPV_ARGS(dataObject));
    }
    return hr;
}


HRESULT PerformOperation(_In_ PCWSTR pszPath)
{
    CComPtr<IDataObject> spdo;
    HRESULT hr = CreateDataObjectFromPath(pszPath, &spdo);
    if (SUCCEEDED(hr))
    {
        CComPtr<IRotationManager> sprm;
        hr = CRotationManager::s_CreateInstance(&sprm);
        if (SUCCEEDED(hr))
        {
            CComPtr<IRotationManagerDiagnostics> sprmd;
            hr = sprm->QueryInterface(IID_PPV_ARGS(&sprmd));
            if (SUCCEEDED(hr))
            {
                hr = EnumerateDataObject(spdo, sprm);
                if (SUCCEEDED(hr))
                {
                    sprm->Start();
                }
            }
        }
    }
 
    return hr;
}

void PrintUsage()
{
    // JPEGAutoRotator.exe /? <path>
    wprintf(L"\nUSAGE:\n");
    wprintf(L"\tJPEGAutoRotator [/?] [path]\n");
    wprintf(L"\n");
    wprintf(L"Options:\n");
    wprintf(L"\t/?\t\t\tDisplays this help message\n");
    wprintf(L"\tpath\t\t\tPath to a file or folder to auto-rotate\n");
    wprintf(L"\n");
    wprintf(L"Example:\n");
    wprintf(L"\t> JPEGAutoRotator c:\\users\\chris\\pictures\\mypicturetorotate.jpg\n");
}

int main()
{
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (SUCCEEDED(hr))
    {
        int cArguments;
        PWSTR *ppszArguments = CommandLineToArgvW(GetCommandLine(), &cArguments);
        if (1 == cArguments || 
            (2 == cArguments && (_wcsicmp(ppszArguments[1], L"\\?") == 0) || (_wcsicmp(ppszArguments[1], L"?") == 0) || (_wcsicmp(ppszArguments[1], L"-?") == 0) || (_wcsicmp(ppszArguments[1], L"-help") == 0) || (_wcsicmp(ppszArguments[1], L"\\help") == 0)))
        {
            // No arguments passed.  Print the usage.
            PrintUsage();
        }
        else if (2 == cArguments)
        {
            hr = PerformOperation(ppszArguments[1]);
        }

        if (FAILED(hr))
        {
            wchar_t message[200];
            StringCchPrintf(message, ARRAYSIZE(message), L"Error: 0x%x\n", hr);
            wprintf(message);
        }

        CoUninitialize();
    }
}

