#pragma once
#include <ShlObj.h>
#include "RotationInterfaces.h"

// Simple class that implements the rotation manager events interface and shows
// the progress dialog.
class CRotationUI :
    public IRotationUI,
    public IRotationManagerEvents
{
public:
    CRotationUI();

    // IUnknown
    IFACEMETHODIMP QueryInterface(__in REFIID riid, __deref_out void** ppv)
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
    IFACEMETHODIMP Initialize(__in IDataObject* pdo);
    IFACEMETHODIMP Start();
    IFACEMETHODIMP Close();

    // IRotationManagerEvents
    IFACEMETHODIMP OnItemAdded(__in UINT uIndex);
    IFACEMETHODIMP OnItemProcessed(__in UINT uIndex);
    IFACEMETHODIMP OnProgress(__in UINT uCompleted, __in UINT uTotal);
    IFACEMETHODIMP OnCanceled();
    IFACEMETHODIMP OnCompleted();

    static HRESULT s_CreateInstance(__in IRotationManager* prm, __deref_out IRotationUI** pprui);

private:
    ~CRotationUI();

    HRESULT _Initialize(__in IRotationManager* prm);
    void _Cleanup();
    HRESULT _EnumerateDataObject(__in IDataObject* pdo);
    HRESULT _ParseEnumItems(_In_ IEnumShellItems *pesi, _In_ UINT depth);

    long m_cRef;
    DWORD m_dwCookie;
    CComPtr<IRotationManager> m_sprm;
    CComPtr<IProgressDialog> m_sppd;
};