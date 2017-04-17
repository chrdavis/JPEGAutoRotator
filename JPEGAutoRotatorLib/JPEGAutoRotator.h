#pragma once

class CJPEGRotator : public IUnknown
{
public:
    CJPEGRotator();

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, __deref_out void ** ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CJPEGRotator, IUnknown),
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

    static HRESULT s_CreateInstance(__in IDataObject* pdo, HWND hwnd, __out CJPEGRotator** pppr);

private:
    ~CJPEGRotator();

private:
    CComPtr<IDataObject> m_spdo;
    long  m_cRef;
    HWND  m_hwndParent;
};
