#pragma once
#include <Shobjidl.h>

class CJPEGAutoRotatorMenu : 
    public IShellExtInit,
    public IContextMenu
{
public:
    CJPEGAutoRotatorMenu();

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, __deref_out void ** ppv)
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
    STDMETHODIMP Initialize(__in_opt PCIDLIST_ABSOLUTE pidlFolder, __in IDataObject *pdto, HKEY hkProgID);

    // IContextMenu
    STDMETHODIMP QueryContextMenu(HMENU hMenu, UINT uIndex, UINT uIDFirst, UINT uIDLast, UINT uFlags);
    STDMETHODIMP InvokeCommand(__in LPCMINVOKECOMMANDINFO pCMI);
    STDMETHODIMP GetCommandString(UINT_PTR uID, UINT uFlags, __in_opt UINT *res, __in LPSTR pName, UINT ccMax)
    {
        return E_NOTIMPL;
    }

    static HRESULT s_CreateInstance(__in_opt IUnknown *punkOuter, __in REFIID riid, __out void **ppv);
    static DWORD WINAPI s_RotationManagerThreadProc(void* pData);

private:
    ~CJPEGAutoRotatorMenu();

    long m_cRef;
    CComPtr<IDataObject> m_spdo;
};
