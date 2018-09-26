#pragma once
#include "RotationInterfaces.h"
#include "srwlock.h"

class CMockRotationItem :
    public IRotationItem,
    public IRotationItemFactory
{
public:
    CMockRotationItem() {}

    // IUnknown
    IFACEMETHODIMP QueryInterface(_In_ REFIID riid, _COM_Outptr_ void** ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CMockRotationItem, IRotationItem),
            QITABENT(CMockRotationItem, IRotationItemFactory),
            { 0, 0 },
        };
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef()
    {
        return InterlockedIncrement(&m_cRef);
    }

    IFACEMETHODIMP_(ULONG) Release()
    {
        LONG cRef = InterlockedDecrement(&m_cRef);
        if (cRef == 0)
        {
            delete this;
        }
        return cRef;
    }

    // IRotationItem
    IFACEMETHODIMP get_Path(_Outptr_ PWSTR* ppszPath)
    {
        CSRWSharedAutoLock lock(&m_lock);
        *ppszPath = nullptr;
        HRESULT hr = m_pszPath ? S_OK : E_FAIL;
        if (SUCCEEDED(hr))
        {
            hr = SHStrDup(m_pszPath, ppszPath);
        }
        return hr;
    }

    IFACEMETHODIMP put_Path(_In_ PCWSTR pszPath)
    {
        CSRWExclusiveAutoLock lock(&m_lock);
        HRESULT hr = pszPath ? S_OK : E_INVALIDARG;
        if (SUCCEEDED(hr))
        {
            CoTaskMemFree(m_pszPath);
            hr = SHStrDup(pszPath, &m_pszPath);
        }
        return hr;
    }

    IFACEMETHODIMP get_WasRotated(_Out_ BOOL* pfWasRotated)
    {
        CSRWSharedAutoLock lock(&m_lock);
        *pfWasRotated = m_fWasRotated;
        return S_OK;
    }

    IFACEMETHODIMP get_IsValidJPEG(_Out_ BOOL* pfIsValidJPEG)
    {
        CSRWSharedAutoLock lock(&m_lock);
        *pfIsValidJPEG = m_fIsValidJPEG;
        return S_OK;
    }

    IFACEMETHODIMP get_IsRotationLossless(_Out_ BOOL* pfIsRotationLossless)
    {
        CSRWSharedAutoLock lock(&m_lock);
        *pfIsRotationLossless = m_fIsRotationLossless;
        return S_OK;
    }

    IFACEMETHODIMP get_OriginalOrientation(_Out_ UINT* puOriginalOrientation)
    {
        CSRWSharedAutoLock lock(&m_lock);
        *puOriginalOrientation = m_uOriginalOrientation;
        return S_OK;
    }

    IFACEMETHODIMP get_Result(_Out_ HRESULT* phrResult)
    {
        CSRWSharedAutoLock lock(&m_lock);
        *phrResult = m_hrResult;
        return S_OK;
    }

    IFACEMETHODIMP put_Result(_In_ HRESULT hrResult)
    {
        CSRWExclusiveAutoLock lock(&m_lock);
        m_hrResult = hrResult;
        return S_OK;
    }

    IFACEMETHODIMP Rotate()
    {
        return S_OK;
    }

    // IRotationItemFactory
    IFACEMETHODIMP Create(_COM_Outptr_ IRotationItem** ppri)
    {
        return CMockRotationItem::s_CreateInstance(L"", ppri);
    }

    HRESULT s_CreateInstance(_In_ PCWSTR pszPath, _COM_Outptr_ IRotationItem** ppri)
    {
        *ppri = nullptr;
        CMockRotationItem* pri = new CMockRotationItem();
        HRESULT hr = pri ? S_OK : E_OUTOFMEMORY;
        if (SUCCEEDED(hr))
        {
            hr = pri->put_Path(pszPath);
            if (SUCCEEDED(hr))
            {
                hr = pri->QueryInterface(IID_PPV_ARGS(ppri));
            }
            pri->Release();
        }
        return hr;
    }

private:
    ~CMockRotationItem()
    {
        CoTaskMemFree(m_pszPath);
    }

private:
    long m_cRef = 1;
    CSRWLock m_lock;
    PWSTR m_pszPath = nullptr;

public:
    bool m_fWasRotated = false;
    bool m_fIsValidJPEG = false;
    bool m_fIsRotationLossless = true;
    UINT m_uOriginalOrientation = 1;
    HRESULT m_hrResult = S_FALSE;  // We init to S_FALSE which means Not Run Yet.  S_OK on success.  Otherwise an error code.
};

