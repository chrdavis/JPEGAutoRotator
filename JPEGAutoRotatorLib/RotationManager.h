#pragma once
#include <ShlObj.h>
#include "RotationInterfaces.h"

class CRotationItem : public IRotationItem
{
public:
    CRotationItem();

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, __deref_out void** ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CRotationItem, IRotationItem),
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

    IFACEMETHODIMP GetItem(__deref_out IShellItem** ppsi);
    IFACEMETHODIMP SetItem(__in IShellItem* psi);
    IFACEMETHODIMP GetResult(__out HRESULT* phrResult);
    IFACEMETHODIMP SetResult(__in HRESULT hrResult);
    IFACEMETHODIMP Rotate();

    static HRESULT s_CreateInstance(__in IShellItem* psi, __deref_out IRotationItem** ppri);

private:
    ~CRotationItem();

private:
    CComPtr<IShellItem> m_spsi;
    long m_cRef;
    HRESULT m_hrResult;  // We init to S_FALSE which means Not Run Yet.  S_OK on success.  Otherwise an error code.
};

// TODO: Consider modifying the below or making them customizable via the interface.  We will likely want to control
// TODO: these from unit tests.

// Maximum number of running worker threads
// We should never exceed the number of logical processors
#define MAX_ROTATION_WORKER_THREADS 64

// Minimum amount of work to schedule to a worker thread
#define MIN_ROTATION_WORK_SIZE 5

class CRotationManager : public IRotationManager
{
public:
    CRotationManager();

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, __deref_out void** ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CRotationManager, IRotationManager),
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

    IFACEMETHODIMP Advise(__in IRotationManagerEvents* prme, __out DWORD* pdwCookie);
    IFACEMETHODIMP UnAdvise(__in DWORD dwCookie);
    IFACEMETHODIMP Start();
    IFACEMETHODIMP Cancel();
    IFACEMETHODIMP AddItem(__in IRotationItem* pri);
    IFACEMETHODIMP GetItem(__in UINT uIndex, __deref_out IRotationItem** ppri);
    IFACEMETHODIMP GetItemCount(__out UINT* puCount);

    static HRESULT s_CreateInstance(__deref_out IRotationManager** pprm);

private:
    ~CRotationManager();

    HRESULT _Init();
    void _Cleanup();
    HRESULT _PerformRotation();
    HRESULT _CreateWorkerThreads();

    static UINT s_GetLogicalProcessorCount();
    static DWORD WINAPI s_rotationWorkerThread(__in void* pv);

private:
    struct ROTATION_WORKER_THREAD_INFO
    {
        HANDLE hWorker;
        DWORD dwThreadId;
        UINT uCompletedItems;
        UINT uTotalItems;
    };

    ROTATION_WORKER_THREAD_INFO m_workerThreadInfo[MAX_ROTATION_WORKER_THREADS];
    HANDLE m_workerThreadHandles[MAX_ROTATION_WORKER_THREADS];
    UINT m_uWorkerThreadCount;
    CComPtr<IObjectCollection> m_spoc;
    long  m_cRef;
    ULONG_PTR m_gdiplusToken;
    HANDLE m_hCancelEvent;
    HANDLE m_hStartEvent;
    CComPtr<IRotationManagerEvents> m_sprme;
    DWORD m_dwCookie;
};
