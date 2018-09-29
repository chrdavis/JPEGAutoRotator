#include "stdafx.h"
#include "JPEGAutoRotatorExt.h"
#include "RotationUI.h"
#include "RotationManager.h"
#include "resource.h"

extern HINSTANCE g_hInst;

HWND g_hwndParent = 0;

CJPEGAutoRotatorMenu::CJPEGAutoRotatorMenu()
{
    DllAddRef();
}

CJPEGAutoRotatorMenu::~CJPEGAutoRotatorMenu()
{
    m_spdo = nullptr;
    DllRelease();
}

HRESULT CJPEGAutoRotatorMenu::s_CreateInstance(_In_opt_ IUnknown*, _In_ REFIID riid, _Out_ void **ppv)
{
    HRESULT hr = E_OUTOFMEMORY;
    CJPEGAutoRotatorMenu *pprm = new CJPEGAutoRotatorMenu();
    if (pprm)
    {
        hr = pprm->QueryInterface(riid, ppv);
        pprm->Release();
    }
    return hr;
}

// IShellExtInit
HRESULT CJPEGAutoRotatorMenu::Initialize(_In_opt_ PCIDLIST_ABSOLUTE, _In_ IDataObject *pdtobj, HKEY)
{
    // Cache the data object to be used later
    m_spdo = pdtobj;
    return S_OK;
}

// IContextMenu
HRESULT CJPEGAutoRotatorMenu::QueryContextMenu(HMENU hMenu, UINT index, UINT uIDFirst, UINT, UINT uFlags)
{
    HRESULT hr = E_UNEXPECTED;
    if (m_spdo)
    {
        if ((uFlags & ~CMF_OPTIMIZEFORINVOKE) && (uFlags & ~(CMF_DEFAULTONLY | CMF_VERBSONLY)))
        {
            wchar_t menuName[64] = { 0 };
            LoadString(g_hInst, _IsFolder() ? IDS_AUTOROTATEFOLDER : IDS_AUTOROTATEIMAGE, menuName, ARRAYSIZE(menuName));
            InsertMenu(hMenu, index, MF_STRING | MF_BYPOSITION, uIDFirst++, menuName);
            hr = MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 1);
        }
    }

    return hr;
}

HRESULT CJPEGAutoRotatorMenu::InvokeCommand(_In_ LPCMINVOKECOMMANDINFO pici)
{
    HRESULT hr = E_FAIL;

    if ((IS_INTRESOURCE(pici->lpVerb)) &&
        (LOWORD(pici->lpVerb) == 0))
    {
        IStream* pstrm = nullptr;
        if (SUCCEEDED(CoMarshalInterThreadInterfaceInStream(__uuidof(m_spdo), m_spdo, &pstrm)))
        {
            if (!SHCreateThread(s_RotationUIThreadProc, pstrm, CTF_COINIT | CTF_PROCESS_REF, nullptr))
            {
                pstrm->Release(); // if we failed to create the thread, then we must release the stream
            }
        }
    }

    return hr;
}

bool CJPEGAutoRotatorMenu::_IsFolder()
{
    bool isFolder = false;

    CComPtr<IShellItemArray> spsia;
    HRESULT hr = SHCreateShellItemArrayFromDataObject(m_spdo, IID_PPV_ARGS(&spsia));
    if (SUCCEEDED(hr))
    {
        CComPtr<IShellItem> spsi;
        hr = spsia->GetItemAt(0, &spsi);
        if (SUCCEEDED(hr))
        {
            PWSTR filePath = nullptr;
            hr = spsi->GetDisplayName(SIGDN_FILESYSPATH, &filePath);
            if (SUCCEEDED(hr))
            {
                if (GetFileAttributes(filePath) & FILE_ATTRIBUTE_DIRECTORY)
                {
                    isFolder = true;
                }
                else
                {
                    PCWSTR fileExt = PathFindExtension(filePath);
                    if (fileExt)
                    {
                        isFolder = false;
                    }
                }

                CoTaskMemFree(filePath);
            }
            else
            {
                isFolder = true;
            }
        }
    }

    return isFolder;
}

DWORD WINAPI CJPEGAutoRotatorMenu::s_RotationUIThreadProc(_In_ void* pData)
{
    IStream* pstrm = (IStream*)pData;
    CComPtr<IDataObject> spdo;
    if (SUCCEEDED(CoGetInterfaceAndReleaseStream(pstrm, IID_PPV_ARGS(&spdo))))
    {
        // Create the rotation manager
        CComPtr<IRotationManager> sprm;
        if (SUCCEEDED(CRotationManager::s_CreateInstance(&sprm)))
        {
            // Create the rotation UI instance and pass the rotation manager
            CComPtr<IRotationUI> sprui;
            if (SUCCEEDED(CRotationUI::s_CreateInstance(sprm, &sprui)))
            {
                if (SUCCEEDED(sprui->Initialize(spdo)))
                {
                    // Call blocks until we are done
                    sprui->Start();
                }
                
                sprui->Close();
            }
        }
    }

    return 0; // ignored
}
