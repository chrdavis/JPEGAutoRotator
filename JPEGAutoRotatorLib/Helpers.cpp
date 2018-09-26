#include "stdafx.h"
#include "Helpers.h"
#include "RotationManager.h"

// Iterate through the data object and add items to the rotation manager
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
            hr = ParseEnumItems(spesi, 0, prm);
        }
    }

    return hr;
}

// Just in case setup a maximum folder depth
#define MAX_ENUM_DEPTH 300

HRESULT ParseEnumItems(_In_ IEnumShellItems *pesi, _In_ UINT depth, _In_ IRotationManager* prm)
{
    HRESULT hr = E_INVALIDARG;

    // We shouldn't get this deep since we only enum the contents of
    // regular folders but adding just in case
    if (depth < MAX_ENUM_DEPTH)
    {
        hr = S_OK;

        ULONG celtFetched = 0;
        IShellItem* psi = nullptr;
        while ((S_OK == pesi->Next(1, &psi, &celtFetched)) && (SUCCEEDED(hr)))
        {
            // Check if this is a file or folder
            SFGAOF att = 0;
            hr = psi->GetAttributes(SFGAO_FOLDER | SFGAO_FILESYSTEM | SFGAO_STREAM, &att);
            if (SUCCEEDED(hr))
            {
                // We only care about filesystem folders.  In some cases, objects will return SFGAO_FOLDER
                // with SFGAO_STREAM (such as zip folders).
                if (att == (SFGAO_FOLDER | SFGAO_FILESYSTEM))
                {
                    // This is a file system folder
                    // Bind to the IShellItem for the IEnumShellItems interface
                    CComPtr<IEnumShellItems> spesiNext;
                    hr = psi->BindToHandler(NULL, BHID_EnumItems, IID_PPV_ARGS(&spesiNext));
                    if (SUCCEEDED(hr))
                    {
                        // Parse the folder contents recursively
                        hr = ParseEnumItems(spesiNext, depth + 1, prm);
                    }
                }
                else if (!(att & SFGAO_FOLDER))
                {
                    // Get the path
                    PWSTR pszPath = nullptr;
                    hr = psi->GetDisplayName(SIGDN_FILESYSPATH, &pszPath);
                    if (SUCCEEDED(hr))
                    {
                        // Check if this is in fact a JPEG so we don't add items needlessly
                        PCWSTR pszExt = PathFindExtension(pszPath);
                        if (pszExt && (StrCmpI(pszExt, L".jpeg") == 0 || StrCmpI(pszExt, L".jpg") == 0))
                        {
                            // TODO: use a factor or restructure code to allow a different IRotationItem implementation
                            CComPtr<IRotationItem> spriNew;
                            hr = CRotationItem::s_CreateInstance(pszPath, &spriNew);
                            if (SUCCEEDED(hr))
                            {
                                hr = prm->AddItem(spriNew);
                            }
                        }
                        CoTaskMemFree(pszPath);
                    }
                }
            }
            psi->Release();
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