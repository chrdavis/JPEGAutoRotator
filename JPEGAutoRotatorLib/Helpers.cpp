#include "stdafx.h"
#include "Helpers.h"
#include "RotationManager.h"
#include <strsafe.h>
#include <pathcch.h>

// Iterate through the data object and add items to the rotation manager
HRESULT EnumerateDataObject(_In_ IDataObject* pdo, _In_ IRotationManager* prm, _In_ bool enumSubFolders)
{
    CComPtr<IShellItemArray> spsia;
    HRESULT hr = SHCreateShellItemArrayFromDataObject(pdo, IID_PPV_ARGS(&spsia));
    if (SUCCEEDED(hr))
    {
        CComPtr<IEnumShellItems> spesi;
        hr = spsia->EnumItems(&spesi);
        if (SUCCEEDED(hr))
        {
            hr = ParseEnumItems(spesi, 0, prm, enumSubFolders);
        }
    }

    return hr;
}

bool IsJPEG(_In_ PCWSTR path)
{
    PCWSTR pszExt = PathFindExtension(path);
    return (pszExt && (StrCmpI(pszExt, L".jpeg") == 0 || StrCmpI(pszExt, L".jpg") == 0));
}

HRESULT AddPathToRotatonManager(_In_ IRotationManager* prm, _In_ PCWSTR path)
{
    CComPtr<IRotationItemFactory> sprif;
    HRESULT hr = prm->GetRotationItemFactory(&sprif);
    if (SUCCEEDED(hr))
    {
        // Use the rotation manager's rotation item factory to create a new IRotationItem
        CComPtr<IRotationItem> spriNew;
        if (SUCCEEDED(sprif->Create(&spriNew)))
        {
            hr = spriNew->put_Path(path);
            if (SUCCEEDED(hr))
            {
                // Add the item to the rotation manager
                hr = prm->AddItem(spriNew);
            }
        }
    }

    return hr;
}

// Just in case setup a maximum folder depth
#define MAX_ENUM_DEPTH 300

HRESULT ParseEnumItems(_In_ IEnumShellItems *pesi, _In_ UINT depth, _In_ IRotationManager* prm, _In_ bool enumSubFolders)
{
    HRESULT hr = E_INVALIDARG;

    // We shouldn't get this deep since we only enum the contents of
    // regular folders but adding just in case
    if (depth < MAX_ENUM_DEPTH)
    {
        CComPtr<IRotationItemFactory> sprif;
        hr = prm->GetRotationItemFactory(&sprif);
        if (SUCCEEDED(hr))
        {
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
                        if (enumSubFolders)
                        {
                            // This is a file system folder
                            // Bind to the IShellItem for the IEnumShellItems interface
                            CComPtr<IEnumShellItems> spesiNext;
                            hr = psi->BindToHandler(NULL, BHID_EnumItems, IID_PPV_ARGS(&spesiNext));
                            if (SUCCEEDED(hr))
                            {
                                // Parse the folder contents recursively
                                hr = ParseEnumItems(spesiNext, depth + 1, prm, enumSubFolders);
                            }
                        }
                    }
                    else if (!(att & SFGAO_FOLDER))
                    {
                        // Get the path
                        PWSTR pszPath = nullptr;
                        hr = psi->GetDisplayName(SIGDN_FILESYSPATH, &pszPath);
                        if (SUCCEEDED(hr))
                        {
                            // Check if this is in fact a JPEG so we don't add items needlessly to the rotation manager.
                            if (IsJPEG(pszPath))
                            {
                                hr = AddPathToRotatonManager(prm, pszPath);
                            }

                            CoTaskMemFree(pszPath);
                        }
                    }
                }
                psi->Release();
            }
        }
    }

    return hr;
}

__inline bool PathIsDotOrDotDot(_In_ PCWSTR pszPath)
{
    return ((pszPath[0] == L'.') && ((pszPath[1] == L'\0') || ((pszPath[1] == L'.') && (pszPath[2] == L'\0'))));
}

// Enumerate via the Find* apis and add any JPEG files to the rotation manager
bool EnumeratePath(_In_ PCWSTR path, _In_ UINT depth, _In_ IRotationManager* prm, _In_ bool enumSubfolders)
{
    bool ret = false;
    if (depth < MAX_ENUM_DEPTH)
    {
        wchar_t searchPath[MAX_PATH] = { 0 };
        wchar_t parent[MAX_PATH] = { 0 };

        StringCchCopy(searchPath, ARRAYSIZE(searchPath), path);
        StringCchCopy(parent, ARRAYSIZE(parent), path);

        if (PathIsDirectory(searchPath))
        {
            // Add wildcard to end of folder path so we can enumerate its contents
            PathCchAddBackslash(searchPath, ARRAYSIZE(searchPath));
            StringCchCat(searchPath, ARRAYSIZE(searchPath), L"*");
        }
        else
        {
            PathCchRemoveFileSpec(parent, ARRAYSIZE(parent));
        }

        CComPtr<IRotationItemFactory> sprif;
        if (SUCCEEDED(prm->GetRotationItemFactory(&sprif)))
        {
            WIN32_FIND_DATA findData = { 0 };
            HANDLE findHandle = FindFirstFile(searchPath, &findData);
            if (findHandle != INVALID_HANDLE_VALUE)
            {
                do
                {
                    if ((findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == FILE_ATTRIBUTE_DIRECTORY)
                    {
                        // Ensure the directory is no . or ..
                        if (enumSubfolders && !PathIsDotOrDotDot(findData.cFileName))
                        {
                            wchar_t pathSubFolder[MAX_PATH] = { 0 };
                            if (SUCCEEDED(PathCchCombine(pathSubFolder, ARRAYSIZE(pathSubFolder), parent, findData.cFileName)))
                            {
                                PathCchAddBackslash(pathSubFolder, ARRAYSIZE(pathSubFolder));
                                ret = EnumeratePath(pathSubFolder, ++depth, prm, enumSubfolders) || ret;
                            }
                        }
                    }
                    else
                    {
                        // Is this a JPEG file?  Check the extension.
                        if (IsJPEG(findData.cFileName))
                        {
                            wchar_t pathFile[MAX_PATH] = { 0 };
                            if (SUCCEEDED(PathCchCombine(pathFile, ARRAYSIZE(pathFile), parent, findData.cFileName)))
                            {
                                ret = (SUCCEEDED(AddPathToRotatonManager(prm, pathFile))) || ret;
                            }
                        }
                    }
                } while (FindNextFile(findHandle, &findData));

                FindClose(findHandle);
            }
        }
    }
    return ret;
}


UINT GetLogicalProcessorCount()
{
    SYSTEM_INFO sysinfo;
    GetSystemInfo(&sysinfo);
    return sysinfo.dwNumberOfProcessors;
}