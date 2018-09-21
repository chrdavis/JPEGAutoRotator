#include "stdafx.h"
#include "RotationUI.h"
#include "RotationManager.h"
#include "Helpers.h"
#include <strsafe.h>
#include "resource.h"

extern HINSTANCE g_hInst;
extern HWND g_hwndParent;

CRotationUI::CRotationUI()
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
    m_spdo = pdo;
    return S_OK;
}

IFACEMETHODIMP CRotationUI::Start()
{
    HRESULT hr = E_FAIL;
    if (m_spdo)
    {
        // Initialize progress dialog 
        WCHAR szResource[100] = { 0 };
        LoadString(g_hInst, IDS_AUTOROTATOR, szResource, ARRAYSIZE(szResource));
        m_sppd->SetTitle(szResource);
        LoadString(g_hInst, IDS_LOADING, szResource, ARRAYSIZE(szResource));
        m_sppd->SetLine(1, szResource, FALSE, nullptr);
        LoadString(g_hInst, IDS_CANCELING, szResource, ARRAYSIZE(szResource));
        m_sppd->SetCancelMsg(szResource, nullptr);

        // Start progress dialog
        m_sppd->StartProgressDialog(nullptr, nullptr, PROGDLG_NORMAL | PROGDLG_MODAL | PROGDLG_AUTOTIME, nullptr);

        // Enumerate the data object and add all items to the rotation manager
        EnumerateDataObject(m_spdo, m_sprm);

        // Update progress dialog line to show we are now rotating
        LoadString(g_hInst, IDS_AUTOROTATING, szResource, ARRAYSIZE(szResource));
        m_sppd->SetLine(1, szResource, FALSE, nullptr);

        // Start operation.  Here we will block but we should get reentered in our event callback.
        // That way we can update the progress dialog, check the cancel state and notify the 
        // manager if we want to cancel.
        hr = m_sprm->Start();

        if (m_sppd)
        {
            m_sppd->StopProgressDialog();
        }
    }
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
    _CheckIfCanceled();
    return S_OK;
}

IFACEMETHODIMP CRotationUI::OnItemProcessed(__in UINT uIndex)
{
    _CheckIfCanceled();

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
    _CheckIfCanceled();

    if (m_sppd)
    {
        WCHAR szProgress[100] = { 0 };
        WCHAR szResource[100] = { 0 };
        LoadString(g_hInst, IDS_ROTATEPROGRESS, szResource, ARRAYSIZE(szResource));
        if (SUCCEEDED(StringCchPrintf(szProgress, ARRAYSIZE(szProgress), szResource, uCompleted, uTotal)))
        {
            m_sppd->SetLine(2, szProgress, FALSE, nullptr);
        }
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

void CRotationUI::_CheckIfCanceled()
{
    // Check if the user canceled from the progress dialog
    if (m_sprm && m_sppd && m_sppd->HasUserCancelled())
    {
        m_sprm->Cancel();
    }
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
