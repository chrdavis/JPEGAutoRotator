#include "stdafx.h"
#include "Helpers.h"

// Iterate through the data object and add paths to the rotation manager
HRESULT EnumerateDataObject(_In_ IDataObject* pdo, _In_ IRotationManager* prm)
{
    CComPtr<IShellItemArray> spsia;
    HRESULT hr = SHCreateShellItemArrayFromDataObject(pdo, IID_PPV_ARGS(&spsia));
    if (SUCCEEDED(hr))
    {
        CComPtr<IEnumShellItems> spesi;
        hr = spsia->EnumItems(&spesi);
        if (SUCCEEDED(hr))
        {
            ULONG celtFetched = 0;
            IShellItem* psi = nullptr;
            while ((S_OK == spesi->Next(1, &psi, &celtFetched)) && (SUCCEEDED(hr)))
            {
                // Check if this is a file or folder
                SFGAOF att = 0;
                hr = psi->GetAttributes(SFGAO_FOLDER | SFGAO_FILESYSTEM | SFGAO_STREAM, &att);
                if (SUCCEEDED(hr))
                {
                    // We only care about filesystem folders.  In some cases, objects will return SFGAO_FOLDER
                    // with SFGAO_STREAM (such as zip folders).
                    if ((att == (SFGAO_FOLDER | SFGAO_FILESYSTEM)) || (!(att & SFGAO_FOLDER)))
                    {
                        // Get the path
                        PWSTR path = nullptr;
                        hr = psi->GetDisplayName(SIGDN_FILESYSPATH, &path);
                        if (SUCCEEDED(hr))
                        {
                            hr = prm->AddPath(path);
                            CoTaskMemFree(path);
                        }
                    }
                }
                psi->Release();
            }
        }
    }

    return hr;
}

UINT GetLogicalProcessorCount()
{
    SYSTEM_INFO sysinfo;
    GetSystemInfo(&sysinfo);
    return sysinfo.dwNumberOfProcessors;
}