#include "stdafx.h"
#include "RotationUI.h"
#include <RotationManager.h>
#include <Helpers.h>
#include "resource.h"

extern HINSTANCE g_hInst;

CRotationUI::CRotationUI()
{
}

CRotationUI::~CRotationUI()
{
}

HRESULT CRotationUI::s_CreateInstance(_In_ IRotationManager* prm, _In_opt_ IDataObject* pdo, _COM_Outptr_ IRotationUI** pprui)
{
    *pprui = nullptr;
    CRotationUI *prui = new CRotationUI();
    HRESULT hr = prui ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        // Pass the rotation manager to the rotation UI so it can subscribe
        // to events
        hr = prui->_Initialize(prm, pdo);
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
    // Initialize progress dialog text
    wchar_t resource[100] = { 0 };
    LoadString(g_hInst, IDS_AUTOROTATOR, resource, ARRAYSIZE(resource));
    m_sppd->SetTitle(resource);
    LoadString(g_hInst, IDS_LOADING, resource, ARRAYSIZE(resource));
    m_sppd->SetLine(1, resource, FALSE, nullptr);
    LoadString(g_hInst, IDS_CANCELING, resource, ARRAYSIZE(resource));
    m_sppd->SetCancelMsg(resource, nullptr);

    // Start progress dialog
    m_sppd->StartProgressDialog(nullptr, nullptr, PROGDLG_NORMAL | PROGDLG_MODAL | PROGDLG_AUTOTIME, nullptr);

    // If we have a IDataObject enumerate it now and add it to the rotation manager.  Caller may have already
    // added items to the rotation manager before passing it to the rotation UI.
    if (m_spdo)
    {
        // Enumerate the data object and add all items to the rotation manager
        EnumerateDataObject(m_spdo, m_sprm);
        m_spdo = nullptr;
    }

    // Update progress dialog line to show we are now rotating
    LoadString(g_hInst, IDS_AUTOROTATING, resource, ARRAYSIZE(resource));
    m_sppd->SetLine(1, resource, FALSE, nullptr);

    // Start operation.  Here we will block but we should get reentered in our event callback.
    // That way we can update the progress dialog, check the cancel state and notify the 
    // manager if we want to cancel.
    HRESULT hr = m_sprm->Start();
    if (m_sppd)
    {
        m_sppd->StopProgressDialog();
    }

    if (SUCCEEDED(hr))
    {
        _ShowResults();
    }
    return hr;
}

IFACEMETHODIMP CRotationUI::Close()
{
    _Cleanup();
    return S_OK;
}

// IRotationManagerEvents
IFACEMETHODIMP CRotationUI::OnItemAdded(_In_ UINT)
{
    _CheckIfCanceled();
    return S_OK;
}

IFACEMETHODIMP CRotationUI::OnItemProcessed(_In_ UINT index)
{
    _CheckIfCanceled();

    // Update the item in our list view
    if (m_sprm)
    {
        CComPtr<IRotationItem> spItem;
        if (SUCCEEDED(m_sprm->GetItem(index, &spItem)))
        {
            BOOL wasRotated = FALSE;
            if (SUCCEEDED(spItem->get_WasRotated(&wasRotated)) && wasRotated)
            {
                PWSTR path = nullptr;
                if (SUCCEEDED(spItem->get_Path(&path)))
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

IFACEMETHODIMP CRotationUI::OnProgress(_In_ UINT completed, _In_ UINT total)
{
    _CheckIfCanceled();

    if (m_sppd)
    {
        // Update text in progress dialog
        wchar_t progress[100] = { 0 };
        wchar_t resource[100] = { 0 };
        LoadString(g_hInst, IDS_ROTATEPROGRESS, resource, ARRAYSIZE(resource));
        if (SUCCEEDED(StringCchPrintf(progress, ARRAYSIZE(progress), resource, completed, total)))
        {
            m_sppd->SetLine(2, progress, FALSE, nullptr);
        }
        m_sppd->SetProgress(completed, total);
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

HRESULT CRotationUI::_Initialize(_In_ IRotationManager* prm, _In_opt_ IDataObject* pdo)
{
    // Cache the rotation manager
    m_sprm = prm;

    // Cache the data object for enumeration later
    m_spdo = pdo;

    // Subscribe to rotation manager events
    HRESULT hr = m_sprm->Advise(this, &m_cookie);
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
        if (m_cookie != 0)
        {
            m_sprm->UnAdvise(m_cookie);
            m_cookie = 0;
        }
    }

    if (m_sppd)
    {
        m_sppd->StopProgressDialog();
    }

    m_sprm = nullptr;
    m_sppd = nullptr;
}

void CRotationUI::_ShowResults()
{
    if (m_sprm)
    {
        // Show Results
        wchar_t resultstitle[100] = { 0 };
        LoadString(g_hInst, IDS_ROTATIONRESULTSTITLE, resultstitle, ARRAYSIZE(resultstitle));
        wchar_t resultsdescfmt[100] = { 0 };
        LoadString(g_hInst, IDS_ROTATIONRESULTS, resultsdescfmt, ARRAYSIZE(resultsdescfmt));

        UINT itemCount = 0;
        m_sprm->GetItemCount(&itemCount);

        UINT numRotated = 0;
        for (UINT i = 0; i < itemCount; i++)
        {
            CComPtr<IRotationItem> spItem;
            m_sprm->GetItem(i, &spItem);
            BOOL wasRotated = FALSE;
            if (SUCCEEDED(spItem->get_WasRotated(&wasRotated)) && wasRotated)
            {
                numRotated++;
            }
        }

        wchar_t resultsdesc[100] = { 0 };
        StringCchPrintf(resultsdesc, ARRAYSIZE(resultsdesc), resultsdescfmt, numRotated, itemCount);

        MessageBox(NULL, resultsdesc, resultstitle, MB_ICONINFORMATION | MB_OK);
    }
}