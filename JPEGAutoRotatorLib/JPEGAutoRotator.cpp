#include "stdafx.h"
#include "JPEGAutoRotator.h"

extern void DllAddRef();
extern void DllRelease();

CJPEGRotator::CJPEGRotator() :
    m_cRef(1),
    m_spdo(nullptr),
    m_hwndParent(nullptr)
{
    DllAddRef();
}

CJPEGRotator::~CJPEGRotator()
{
    m_spdo = nullptr;
    DllRelease();
}

HRESULT CJPEGRotator::s_CreateInstance(__in IDataObject* pdo, HWND hwnd, __out CJPEGRotator** pppr)
{
    *pppr = NULL;

    CJPEGRotator *ppr = new CJPEGRotator();
    HRESULT hr = ppr ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        ppr->m_spdo = pdo;
        ppr->m_hwndParent = hwnd;
        *pppr = ppr;
    }
    return hr;
}
