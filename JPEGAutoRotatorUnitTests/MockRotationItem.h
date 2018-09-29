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
        return InterlockedIncrement(&m_refCount);
    }

    IFACEMETHODIMP_(ULONG) Release()
    {
        LONG refCount = InterlockedDecrement(&m_refCount);
        if (refCount == 0)
        {
            delete this;
        }
        return refCount;
    }

    // IRotationItem
    IFACEMETHODIMP get_Path(_Outptr_ PWSTR* path)
    {
        CSRWSharedAutoLock lock(&m_lock);
        *path = nullptr;
        HRESULT hr = m_path ? S_OK : E_FAIL;
        if (SUCCEEDED(hr))
        {
            hr = SHStrDup(m_path, path);
        }
        return hr;
    }

    IFACEMETHODIMP put_Path(_In_ PCWSTR path)
    {
        CSRWExclusiveAutoLock lock(&m_lock);
        HRESULT hr = path ? S_OK : E_INVALIDARG;
        if (SUCCEEDED(hr))
        {
            CoTaskMemFree(m_path);
            hr = SHStrDup(path, &m_path);
        }
        return hr;
    }

    IFACEMETHODIMP get_WasRotated(_Out_ BOOL* wasRotated)
    {
        CSRWSharedAutoLock lock(&m_lock);
        *wasRotated = m_wasRotated;
        return S_OK;
    }

    IFACEMETHODIMP get_IsValidJPEG(_Out_ BOOL* isValidJPEG)
    {
        CSRWSharedAutoLock lock(&m_lock);
        *isValidJPEG = m_isValidJPEG;
        return S_OK;
    }

    IFACEMETHODIMP get_IsRotationLossless(_Out_ BOOL* isRotationLossless)
    {
        CSRWSharedAutoLock lock(&m_lock);
        *isRotationLossless = m_isRotationLossless;
        return S_OK;
    }

    IFACEMETHODIMP get_OriginalOrientation(_Out_ UINT* originalOrientation)
    {
        CSRWSharedAutoLock lock(&m_lock);
        *originalOrientation = m_originalOrientation;
        return S_OK;
    }

    IFACEMETHODIMP get_Result(_Out_ HRESULT* result)
    {
        CSRWSharedAutoLock lock(&m_lock);
        *result = m_result;
        return S_OK;
    }

    IFACEMETHODIMP put_Result(_In_ HRESULT result)
    {
        CSRWExclusiveAutoLock lock(&m_lock);
        m_result = result;
        return S_OK;
    }

    IFACEMETHODIMP Load()
    {
        return S_OK;
    }

    IFACEMETHODIMP Rotate()
    {
        return S_OK;
    }

    // IRotationItemFactory
    IFACEMETHODIMP Create(_COM_Outptr_ IRotationItem** ppItem)
    {
        return CMockRotationItem::s_CreateInstance(L"", ppItem);
    }

    HRESULT s_CreateInstance(_In_ PCWSTR path, _COM_Outptr_ IRotationItem** ppItem)
    {
        *ppItem = nullptr;
        CMockRotationItem* pItem = new CMockRotationItem();
        HRESULT hr = pItem ? S_OK : E_OUTOFMEMORY;
        if (SUCCEEDED(hr))
        {
            hr = pItem->put_Path(path);
            if (SUCCEEDED(hr))
            {
                hr = pItem->QueryInterface(IID_PPV_ARGS(ppItem));
            }
            pItem->Release();
        }
        return hr;
    }

private:
    ~CMockRotationItem()
    {
        CoTaskMemFree(m_path);
    }

private:
    long m_refCount = 1;
    CSRWLock m_lock;
    PWSTR m_path = nullptr;

public:
    bool m_wasRotated = false;
    bool m_isValidJPEG = false;
    bool m_isRotationLossless = true;
    UINT m_originalOrientation = 1;
    HRESULT m_result = S_FALSE;  // We init to S_FALSE which means Not Run Yet.  S_OK on success.  Otherwise an error code.
};

