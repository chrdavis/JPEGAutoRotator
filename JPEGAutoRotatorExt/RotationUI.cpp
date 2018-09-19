#include "stdafx.h"
#include "RotationUI.h"
#include "RotationManager.h"
#include "Helpers.h"
#include "resource.h"

extern HINSTANCE g_hInst;

// Rotation UI
CRotationUI::CRotationUI() :
    m_cRef(1),
    m_dwCookie(0)
{
}

CRotationUI::~CRotationUI()
{
}

HRESULT CRotationUI::s_CreateInstance(__in IRotationManager* prm, __deref_out IRotationUI** pprui)
{
    *pprui = nullptr;
    CRotationUI *prui = new CRotationUI();
    HRESULT hr = prui ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = prui->_Initialize(prm);
        if (SUCCEEDED(hr))
        {
            hr = prui->QueryInterface(IID_PPV_ARGS(pprui));
        }
        prui->Release();
    }
    return hr;
}

// IRotationUI
IFACEMETHODIMP CRotationUI::Initialize(__in IDataObject* pdo)
{
    return EnumerateDataObject(pdo, m_sprm);
}

IFACEMETHODIMP CRotationUI::Start()
{

    // Initialize progress dialog 
    WCHAR szResource[100] = { 0 };
    LoadString(g_hInst, IDS_AUTOROTATING, szResource, ARRAYSIZE(szResource));
    m_sppd->SetTitle(szResource);
    LoadString(g_hInst, IDS_CANCELING, szResource, ARRAYSIZE(szResource));
    m_sppd->SetCancelMsg(szResource, nullptr);

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
IFACEMETHODIMP CRotationUI::OnItemAdded(__in UINT)
{
    return S_OK;
}

IFACEMETHODIMP CRotationUI::OnItemProcessed(__in UINT uIndex)
{
    // Update the item in our list view
    if (m_sprm)
    {
        CComPtr<IRotationItem> spri;
        if (SUCCEEDED(m_sprm->GetItem(uIndex, &spri)))
        {
            BOOL wasRotated = FALSE;
            if (SUCCEEDED(spri->get_WasRotated(&wasRotated)) && wasRotated)
            {
                PWSTR path = nullptr;
                if (SUCCEEDED(spri->get_Path(&path)))
                {
                    // Notify the Shell to update the thumbnail
                    SHChangeNotify(SHCNE_UPDATEITEM, (SHCNF_PATH | SHCNF_FLUSHNOWAIT), path, nullptr);
                    CoTaskMemFree(path);
                }
            }
        }
    }

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

IFACEMETHODIMP CRotationUI::OnCanceled()
{
    if (m_sppd)
    {
        m_sppd->StopProgressDialog();
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

HRESULT CRotationUI::_Initialize(__in IRotationManager* prm)
{
    m_sprm = prm;

    // Subscribe to rotation manager events
    HRESULT hr = m_sprm->Advise(this, &m_dwCookie);
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
    m_sppd = nullptr;
}
