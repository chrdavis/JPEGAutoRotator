#include "stdafx.h"
#include "JPEGAutoRotatorExt.h"

DWORD g_dwModuleRefCount = 0;
HINSTANCE g_hInst = 0;

class CJPEGAutoRotatorClassFactory : public IClassFactory
{
public:
    CJPEGAutoRotatorClassFactory(REFCLSID clsid) : 
        m_cRef(1),
        m_clsid(clsid)
    {
        DllAddRef();
    }

    // IUnknown methods
    IFACEMETHODIMP QueryInterface(REFIID riid, __deref_out void ** ppv)
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

    // IClassFactory methods
    IFACEMETHODIMP CreateInstance(__in_opt IUnknown *punkOuter, REFIID riid, __deref_out void **ppv)
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

    long m_cRef;
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
STDAPI DllGetClassObject(__in REFCLSID clsid, __in REFIID riid, __deref_out void **ppv)
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

