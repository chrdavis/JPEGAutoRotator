#include "stdafx.h"
#include "JPEGAutoRotatorExt.h"
#include "RotationUI.h"
#include "RotationManager.h"
#include "resource.h"

static HINSTANCE g_hInst = 0;

CJPEGAutoRotatorMenu::CJPEGAutoRotatorMenu() :
    m_cRef(1),
    m_spdo(nullptr)
{
    DllAddRef();
}

CJPEGAutoRotatorMenu::~CJPEGAutoRotatorMenu()
{
    m_spdo = nullptr;
    DllRelease();
}

HRESULT CJPEGAutoRotatorMenu::s_CreateInstance(__in_opt IUnknown*, __in REFIID riid, __out void **ppv)
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
HRESULT CJPEGAutoRotatorMenu::Initialize(__in_opt PCIDLIST_ABSOLUTE, __in IDataObject *pdtobj, HKEY)
{
    // Cache the data object to be used later
    m_spdo = pdtobj;
    return S_OK;
}

// IContextMenu
HRESULT CJPEGAutoRotatorMenu::QueryContextMenu(HMENU hMenu, UINT uIndex, UINT uIDFirst, UINT, UINT uFlags)
{
    HRESULT hr = E_UNEXPECTED;
    if (m_spdo)
    {
        if ((uFlags & ~CMF_OPTIMIZEFORINVOKE) && (uFlags & ~(CMF_DEFAULTONLY | CMF_VERBSONLY)))
        {
            WCHAR szMenuName[64] = { 0 };
            LoadString(g_hInst, _IsFolder() ? IDS_AUTOROTATEFOLDER : IDS_AUTOROTATEIMAGE, szMenuName, ARRAYSIZE(szMenuName));
            InsertMenu(hMenu, uIndex, MF_STRING | MF_BYPOSITION, uIDFirst++, szMenuName);
            hr = MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 1);
        }
    }

    return hr;
}

HRESULT CJPEGAutoRotatorMenu::InvokeCommand(__in LPCMINVOKECOMMANDINFO pici)
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
            PWSTR pszFilePath = nullptr;
            hr = spsi->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath);
            if (SUCCEEDED(hr))
            {
                if (GetFileAttributes(pszFilePath) & FILE_ATTRIBUTE_DIRECTORY)
                {
                    isFolder = true;
                }
                else
                {
                    PCWSTR pszExt = PathFindExtension(pszFilePath);
                    if (pszExt)
                    {
                        isFolder = false;
                    }
                }

                CoTaskMemFree(pszFilePath);
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
                    sprui->Start();
                }
            }
        }
    }

    return 0; // ignored
}
