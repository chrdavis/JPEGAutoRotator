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

    // IRotationUI
    IFACEMETHODIMP Initialize(_In_ IDataObject* pdo);
    IFACEMETHODIMP Start();
    IFACEMETHODIMP Close();

    // IRotationManagerEvents
    IFACEMETHODIMP OnItemAdded(_In_ UINT index);
    IFACEMETHODIMP OnItemProcessed(_In_ UINT index);
    IFACEMETHODIMP OnProgress(_In_ UINT completed, _In_ UINT total);
    IFACEMETHODIMP OnCanceled();
    IFACEMETHODIMP OnCompleted();

    static HRESULT s_CreateInstance(_In_ IRotationManager* prm, _COM_Outptr_ IRotationUI** pprui);

private:
    ~CRotationUI();

    HRESULT _Initialize(_In_ IRotationManager* prm);
    void _Cleanup();
    void _CheckIfCanceled();

    long m_refCount = 1;
    DWORD m_cookie = 0;
    CComPtr<IRotationManager> m_sprm;
    CComPtr<IProgressDialog> m_sppd;
    CComPtr<IDataObject> m_spdo;
};