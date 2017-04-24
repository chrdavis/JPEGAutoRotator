#pragma once

class CRotationItem : public IUnknown
{
public:
    CRotationItem();

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, __deref_out void** ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CRotationItem, IUnknown),
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

    HRESULT GetItem(__deref_out IShellItem** ppsi);

    static HRESULT s_CreateInstance(__in IShellItem* psi, __deref_out CRotationItem** ppri);

private:
    ~CRotationItem();

private:
    CComPtr<IShellItem> m_spsi;
    long  m_cRef;
};


struct
{
    
    CComPtr<IObjectArray> m_spoa;
};

class CRotationManager : public IUnknown
{
public:
    CRotationManager();

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, __deref_out void** ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CRotationManager, IUnknown),
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

    HRESULT PerformRotation();

    static HRESULT s_CreateInstance(__in IDataObject* pdo, HWND hwnd, __deref_out CRotationManager** pprm);

private:
    HRESULT _EnumerateDataObject();
    HRESULT _Init();
    HRESULT _Cleanup();

    ~CRotationManager();

private:
    CComPtr<IDataObject> m_spdo;
    CComPtr<IObjectCollection> m_spoc;
    long  m_cRef;
    HWND  m_hwndParent;
    ULONG_PTR m_gdiplusToken;
};
