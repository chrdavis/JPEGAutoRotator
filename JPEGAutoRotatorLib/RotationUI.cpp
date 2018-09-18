#include "stdafx.h"
#include "RotationUI.h"
#include "RotationManager.h"

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
    m_spdo = pdo;
    return _EnumerateDataObject();
}

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
IFACEMETHODIMP CRotationUI::OnAdded(__in UINT)
{
    return S_OK;
}

IFACEMETHODIMP CRotationUI::OnRotated(__in UINT uIndex)
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
            hr = _ParseEnumItems(spesi, 0);
        }
    }

    return hr;
}

// Just in case setup a maximum folder depth
#define MAX_ENUM_DEPTH 300

HRESULT CRotationUI::_ParseEnumItems(_In_ IEnumShellItems *pesi, _In_ UINT depth)
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
                        hr = _ParseEnumItems(spesiNext, depth + 1);
                    }
                }
                else if (!(att & SFGAO_FOLDER))
                {
                    // Get the path
                    PWSTR pszPath = nullptr;
                    hr = psi->GetDisplayName(SIGDN_FILESYSPATH, &pszPath);
                    if (SUCCEEDED(hr))
                    {
                        CComPtr<IRotationItem> spriNew;
                        hr = CRotationItem::s_CreateInstance(pszPath, &spriNew);
                        if (SUCCEEDED(hr))
                        {
                            hr = m_sprm->AddItem(spriNew);
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