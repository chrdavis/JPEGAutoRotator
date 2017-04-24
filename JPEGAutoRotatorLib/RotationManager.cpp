#include "stdafx.h"
#include "RotationManager.h"
#include <gdiplus.h>

extern void DllAddRef();
extern void DllRelease();

using namespace Gdiplus;

CRotationItem::CRotationItem() :
    m_cRef(1)
{
}

CRotationItem::~CRotationItem()
{
}

HRESULT CRotationItem::GetItem(__deref_out IShellItem** ppsi)
{
    *ppsi = nullptr;
    HRESULT hr = E_FAIL;
    if (m_spsi)
    {
        *ppsi = m_spsi;
        (*ppsi)->AddRef();
    }
    return hr;
}

HRESULT CRotationItem::s_CreateInstance(__in IShellItem* psi, __out CRotationItem** ppri)
{
    *ppri = nullptr;
    CRotationItem* pri = new CRotationItem();
    HRESULT hr = pri ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        pri->m_spsi = psi;
        *ppri = pri;
    }
    return hr;
}

CRotationManager::CRotationManager() :
    m_cRef(1),
    m_spdo(nullptr),
    m_hwndParent(nullptr)
{
    DllAddRef();
}

CRotationManager::~CRotationManager()
{
    m_spdo = nullptr;
    DllRelease();
}

HRESULT CRotationManager::s_CreateInstance(__in IDataObject* pdo, HWND hwnd, __deref_out CRotationManager** pprm)
{
    *pprm = nullptr;
    CRotationManager *prm = new CRotationManager();
    HRESULT hr = prm ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        // TODO: Init
        prm->m_spdo = pdo;
        prm->m_hwndParent = hwnd;
        *pprm = prm;
    }
    return hr;
}

HRESULT CRotationManager::_Init()
{
    GdiplusStartupInput gdiplusStartupInput;
    GdiplusStartup(&m_gdiplusToken, &gdiplusStartupInput, nullptr);

    return S_OK;
}

HRESULT CRotationManager::_Cleanup()
{
    GdiplusShutdown(m_gdiplusToken);
    return S_OK;
}

// Iterate through the data object and add items to the IObjectCollection
HRESULT CRotationManager::_EnumerateDataObject()
{
    // Clear any existing collection
    m_spoc = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_EnumerableObjectCollection,
        nullptr,
        CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(&m_spoc));
    if (SUCCEEDED(hr))
    {
        CComPtr<IShellItemArray> spsia;
        hr = SHCreateShellItemArrayFromDataObject(m_spdo, IID_PPV_ARGS(&spsia));
        if (SUCCEEDED(hr))
        {
            CComPtr<IEnumShellItems> spesi;
            hr = spsia->EnumItems(&spesi);
            if (SUCCEEDED(hr))
            {
                ULONG celtFetched;
                IShellItem *psi;
                while ((S_OK == spesi->Next(1, &psi, &celtFetched)) && (SUCCEEDED(hr)))
                {
                    SFGAOF att = 0;
                    if (SUCCEEDED(psi->GetAttributes(SFGAO_FOLDER, &att))) // pItem is a IShellItem*
                    {
                        // TODO: should we do this here or later?
                        // Don't bother including folders
                        if (!(att & SFGAO_FOLDER))
                        {
                            CRotationItem* priNew;
                            hr = CRotationItem::s_CreateInstance(psi, &priNew);
                            if (SUCCEEDED(hr))
                            {
                                hr = m_spoc->AddObject(priNew);
                                priNew->Release();
                            }
                        }
                    }
                    psi->Release();
                }
            }
        }
    }

    return hr;
}
