#include "stdafx.h"
#include "JPEGAutoRotatorExt.h"
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
    // Cache the data object to be used later
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

// TODO: can we create the progress dialog here and setup a callback on the rotation manager?
// TODO: What is the lifetime of CJPEGAutoRotatorMenu?  Is it released after we enter the thread below?
// TODO: Or can it host our callback interface and update the progress dialog?
// TODO: We should also convert the IDataObject to IShellItems and create IRenameItems.  Then add
// TODO: them to the manager here.
DWORD WINAPI CJPEGAutoRotatorMenu::s_RotationUIThreadProc(void* pData)
{
    IStream* pstrm = (IStream*)pData;

    CComPtr<IDataObject> spdo;
    if (SUCCEEDED(CoGetInterfaceAndReleaseStream(pstrm, IID_PPV_ARGS(&spdo))))
    {
        CComPtr<IRotationUI> sprui;
        if (SUCCEEDED(CRotationUI::s_CreateInstance(spdo, &sprui)))
        {
            sprui->Start();
        }
    }

    return 0; // ignored
}

// Rotation UI
CRotationUI::CRotationUI() :
    m_cRef(1),
    m_dwCookie(0)
{
}

CRotationUI::~CRotationUI()
{
}

HRESULT CRotationUI::s_CreateInstance(__in IDataObject* pdo, __deref_out IRotationUI** pprui)
{
    *pprui = nullptr;
    CRotationUI *prui = new CRotationUI();
    HRESULT hr = prui ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = prui->_Initialize(pdo);
        if (SUCCEEDED(hr))
        {
            hr = prui->QueryInterface(IID_PPV_ARGS(pprui));
        }
        prui->Release();
    }
    return hr;
}

// IRotationUI
IFACEMETHODIMP CRotationUI::Start()
{
    // Start progress dialog
    m_sppd->StartProgressDialog(nullptr, nullptr, PROGDLG_NORMAL | PROGDLG_AUTOTIME, nullptr);

    // Start operation.  Here we will block but we should get reentered in our event callback.
    // That way we can update the progress dialog, check the cancel state and notify the 
    // manager if we want to cancel.
    HRESULT hr = m_sprm->Start();

    m_sppd->StopProgressDialog();

    return hr;
}

IFACEMETHODIMP CRotationUI::Close()
{
    _Cleanup();

    return S_OK;
}

// IRotationManagerEvents
IFACEMETHODIMP CRotationUI::OnRotated(__in UINT uIndex)
{
    // TODO: Show dialog on error?  Can't block though
    return S_OK;
}

IFACEMETHODIMP CRotationUI::OnProgress(__in UINT uCompleted, __in UINT uTotal)
{
    if (m_sppd)
    {
        m_sppd->SetProgress(uCompleted, uTotal);
    }
    return S_OK;
}

IFACEMETHODIMP CRotationUI::OnCompleted()
{
    if (m_sppd)
    {
        m_sppd->StopProgressDialog();
    }
    return S_OK;
}

HRESULT CRotationUI::_Initialize(__in IDataObject* pdo)
{
    // Cache the data object for now.  We will enumerate in Start.
    m_spdo = pdo;

    // Create the rotation manager
    HRESULT hr = CRotationManager::s_CreateInstance(&m_sprm);
    if (SUCCEEDED(hr))
    {
        hr = m_sprm->Advise(this, &m_dwCookie);
    }

    if (SUCCEEDED(hr))
    {
        // Create the progress dialog we will show during the operation
        hr = CoCreateInstance(CLSID_ProgressDialog,
            nullptr,
            CLSCTX_INPROC_SERVER,
            IID_PPV_ARGS(&m_sppd));
    }

    if (FAILED(hr))
    {
        _Cleanup();
    }

    return hr;
}

void CRotationUI::_Cleanup()
{
    if (m_sprm)
    {
        // Ensure we cancel
        m_sprm->Cancel();

        // Be sure we UnAdvise to the manager is not holding a ref on us
        if (m_dwCookie != 0)
        {
            m_sprm->UnAdvise(m_dwCookie);
            m_dwCookie = 0;
        }
    }

    if (m_sppd)
    {
        m_sppd->StopProgressDialog();
    }
    
    m_sprm = nullptr;
    m_spdo = nullptr;
    m_sppd = nullptr;
}

// Iterate through the data object and add items to the rotation manager
HRESULT CRotationUI::_EnumerateDataObject()
{
    CComPtr<IShellItemArray> spsia;
    HRESULT hr = SHCreateShellItemArrayFromDataObject(m_spdo, IID_PPV_ARGS(&spsia));
    if (SUCCEEDED(hr))
    {
        CComPtr<IEnumShellItems> spesi;
        hr = spsia->EnumItems(&spesi);
        if (SUCCEEDED(hr))
        {
            ULONG celtFetched = 0;
            IShellItem *psi = nullptr;
            while ((S_OK == spesi->Next(1, &psi, &celtFetched)) && (SUCCEEDED(hr)))
            {
                SFGAOF att = 0;
                if (SUCCEEDED(psi->GetAttributes(SFGAO_FOLDER, &att)))
                {
                    // TODO: should we do this here or later?
                    // Don't bother including folders
                    if (!(att & SFGAO_FOLDER))
                    {
                        IRotationItem* priNew;
                        hr = CRotationItem::s_CreateInstance(psi, &priNew);
                        if (SUCCEEDED(hr))
                        {
                            hr = m_sprm->AddItem(priNew);
                            priNew->Release();
                        }
                    }
                }
                psi->Release();
            }
        }
    }

    return hr;
}
