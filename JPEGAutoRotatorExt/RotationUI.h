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
    IFACEMETHODIMP QueryInterface(_In_ REFIID riid, _COM_Outptr_ void** ppv)
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
    IFACEMETHODIMP Initialize(_In_ IDataObject* pdo);
    IFACEMETHODIMP Start();
    IFACEMETHODIMP Close();

    // IRotationManagerEvents
    IFACEMETHODIMP OnItemAdded(_In_ UINT uIndex);
    IFACEMETHODIMP OnItemProcessed(_In_ UINT uIndex);
    IFACEMETHODIMP OnProgress(_In_ UINT uCompleted, _In_ UINT uTotal);
    IFACEMETHODIMP OnCanceled();
    IFACEMETHODIMP OnCompleted();

    static HRESULT s_CreateInstance(_In_ IRotationManager* prm, _COM_Outptr_ IRotationUI** pprui);

private:
    ~CRotationUI();

    HRESULT _Initialize(_In_ IRotationManager* prm);
    void _Cleanup();
    void _CheckIfCanceled();

    HWND m_hwndWorker = 0;
    long m_cRef = 1;
    DWORD m_dwCookie = 0;
    CComPtr<IRotationManager> m_sprm;
    CComPtr<IProgressDialog> m_sppd;
    CComPtr<IDataObject> m_spdo;
};