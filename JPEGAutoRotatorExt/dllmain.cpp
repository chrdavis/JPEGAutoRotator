#include "stdafx.h"
#include "JPEGAutoRotatorExt.h"

DWORD g_dwModuleRefCount = 0;
HINSTANCE g_hInst = 0;

class CJPEGAutoRotatorClassFactory : public IClassFactory
{
public:
    CJPEGAutoRotatorClassFactory(_In_ REFCLSID clsid) : 
        m_refCount(1),
        m_clsid(clsid)
    {
        DllAddRef();
    }

    // IUnknown methods
    IFACEMETHODIMP QueryInterface(_In_ REFIID riid, _COM_Outptr_ void ** ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CJPEGAutoRotatorClassFactory, IClassFactory),
            { 0 }
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

    // IClassFactory methods
    IFACEMETHODIMP CreateInstance(_In_opt_ IUnknown *punkOuter, _In_ REFIID riid, _COM_Outptr_ void** ppv)
    {
        *ppv = NULL;
        HRESULT hr;
        if (punkOuter)
        {
            hr = CLASS_E_NOAGGREGATION;
        }
        else
        {
            if (m_clsid == CLSID_JPEGAutoRotatorMenu)
            {
                hr = CJPEGAutoRotatorMenu::s_CreateInstance(punkOuter, riid, ppv);
            }
            else
            {
                hr = CLASS_E_CLASSNOTAVAILABLE;
            }
        }
        return hr;
    }

    IFACEMETHODIMP LockServer(BOOL bLock)
    {
        if (bLock)
        {
            DllAddRef();
        }
        else
        {
            DllRelease();
        }
        return S_OK;
    }

private:
    ~CJPEGAutoRotatorClassFactory()
    {
        DllRelease();
    }

    long m_refCount;
    CLSID m_clsid;
};

BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, void*)
{
    switch (dwReason)
    {
    case DLL_PROCESS_ATTACH:
        g_hInst = hInstance;
        DisableThreadLibraryCalls(hInstance);
        break;

    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

//
// Checks if there are any external references to this module
//
STDAPI DllCanUnloadNow()
{
    return (g_dwModuleRefCount == 0) ? S_OK : S_FALSE;
}

//
// DLL export for creating COM objects
//
STDAPI DllGetClassObject(_In_ REFCLSID clsid, _In_ REFIID riid, _COM_Outptr_ void **ppv)
{
    *ppv = NULL;
    HRESULT hr = E_OUTOFMEMORY;
    CJPEGAutoRotatorClassFactory *pClassFactory = new CJPEGAutoRotatorClassFactory(clsid);
    if (pClassFactory)
    {
        hr = pClassFactory->QueryInterface(riid, ppv);
        pClassFactory->Release();
    }
    return hr;
}

STDAPI DllRegisterServer()
{
    return S_OK;
}

STDAPI DllUnregisterServer()
{
    return S_OK;
}

void DllAddRef()
{
    g_dwModuleRefCount++;
}

void DllRelease()
{
    g_dwModuleRefCount--;
}

