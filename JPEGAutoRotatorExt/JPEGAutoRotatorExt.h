#pragma once
#include <RotationInterfaces.h>

class CJPEGAutoRotatorMenu : 
    public IShellExtInit,
    public IContextMenu
{
public:
    CJPEGAutoRotatorMenu();

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, __deref_out void** ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CJPEGAutoRotatorMenu, IShellExtInit),
            QITABENT(CJPEGAutoRotatorMenu, IContextMenu),
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

    // IShellExtInit
    STDMETHODIMP Initialize(__in_opt PCIDLIST_ABSOLUTE pidlFolder, __in IDataObject* pdto, HKEY hkProgID);

    // IContextMenu
    STDMETHODIMP QueryContextMenu(HMENU hMenu, UINT uIndex, UINT uIDFirst, UINT uIDLast, UINT uFlags);
    STDMETHODIMP InvokeCommand(__in LPCMINVOKECOMMANDINFO pCMI);
    STDMETHODIMP GetCommandString(UINT_PTR uID, UINT uFlags, __in_opt UINT* res, __in LPSTR pName, UINT ccMax)
    {
        return E_NOTIMPL;
    }

    static HRESULT s_CreateInstance(__in_opt IUnknown* punkOuter, __in REFIID riid, __out void** ppv);
    static DWORD WINAPI s_RotationUIThreadProc(void* pData);

private:
    ~CJPEGAutoRotatorMenu();

    long m_cRef;
    CComPtr<IDataObject> m_spdo;
};

// Simple class that implements the rotation manager events interface and shows
// the progress dialog.
class CRotationUI :
    public IRotationUI,
    public IRotationManagerEvents
{
public:
    CRotationUI();

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, __deref_out void** ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CRotationUI, IRotationUI),
            QITABENT(CRotationUI, IRotationManagerEvents),
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

    // IRotationUI
    IFACEMETHODIMP Start();
    IFACEMETHODIMP Close();

    // IRotationManagerEvents
    IFACEMETHODIMP OnRotated(__in UINT uIndex);
    IFACEMETHODIMP OnProgress(__in UINT uCompleted, __in UINT uTotal);
    IFACEMETHODIMP OnCompleted();

    static HRESULT s_CreateInstance(__in IDataObject* pdo, __deref_out IRotationUI** pprui);

private:
    ~CRotationUI();
    
    HRESULT _Initialize(__in IDataObject* pdo);
    void _Cleanup();
    HRESULT _EnumerateDataObject();

    long m_cRef;
    DWORD m_dwCookie;
    CComPtr<IDataObject> m_spdo;
    CComPtr<IRotationManager> m_sprm;
    CComPtr<IProgressDialog> m_sppd;
};
