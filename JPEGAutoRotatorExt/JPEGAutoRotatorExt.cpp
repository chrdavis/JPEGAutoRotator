#include "stdafx.h"
#include "JPEGAutoRotatorExt.h"
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

HRESULT CJPEGAutoRotatorMenu::s_CreateInstance(__in_opt IUnknown *punkOuter, __in REFIID riid, __out void **ppv)
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
HRESULT CJPEGAutoRotatorMenu::Initialize(__in_opt PCIDLIST_ABSOLUTE pidlFolder, __in IDataObject *pdtobj, HKEY hkProgID)
{
    m_spdo = pdtobj;

    return S_OK;
}

// IContextMenu
HRESULT CJPEGAutoRotatorMenu::QueryContextMenu(HMENU hMenu, UINT uIndex, UINT uIDFirst, UINT uIDLast, UINT uFlags)
{
    HRESULT hr = E_UNEXPECTED;
    if (m_spdo)
    {
        if ((uFlags & ~CMF_OPTIMIZEFORINVOKE) && (uFlags & ~(CMF_DEFAULTONLY | CMF_VERBSONLY)))
        {
            WCHAR szMenuName[64] = { 0 };
            LoadString(g_hInst, IDS_AUTOROTATEIMAGE, szMenuName, ARRAYSIZE(szMenuName));
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
        hr = S_OK;
        // TODO: Update params
        /*RENAMEPARAMS *prp = (RENAMEPARAMS*)LocalAlloc(LPTR, sizeof(*prp));
        if (prp)
        {
            prp->hwndParent = pici->hwnd;

            if (SUCCEEDED(CoMarshalInterThreadInterfaceInStream(__uuidof(_pdo), _pdo, &(prp->pstmDataObject))))
            {
                if (!SHCreateThread(LaunchThreadProc, prp, CTF_COINIT | CTF_PROCESS_REF, NULL))
                {
                    prp->pstmDataObject->Release(); // if we failed to create the thread, then we must release the stream
                    LocalFree(prp);
                }
            }
            else
            {
                LocalFree(prp);
            }
        }*/
    }

    return hr;
}


