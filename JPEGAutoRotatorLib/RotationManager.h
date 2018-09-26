#pragma once
#include <ShlObj.h>
#include "RotationInterfaces.h"
#include <vector>
#include "srwlock.h"

class CRotationItem :
    public IRotationItem,
    public IRotationItemFactory
{
public:
    CRotationItem();

    // IUnknown
    IFACEMETHODIMP QueryInterface(_In_ REFIID riid, _COM_Outptr_ void** ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CRotationItem, IRotationItem),
            QITABENT(CRotationItem, IRotationItemFactory),
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

    // IRotationItem
    IFACEMETHODIMP get_Path(_Outptr_ PWSTR* ppszPath);
    IFACEMETHODIMP put_Path(_In_ PCWSTR pszPath);
    IFACEMETHODIMP get_WasRotated(_Out_ BOOL* pfWasRotated);
    IFACEMETHODIMP get_IsValidJPEG(_Out_ BOOL* pfIsValidJPEG);
    IFACEMETHODIMP get_IsRotationLossless(_Out_ BOOL* pfIsRotationLossless);
    IFACEMETHODIMP get_OriginalOrientation(_Out_ UINT* puOriginalOrientation);
    IFACEMETHODIMP get_Result(_Out_ HRESULT* phrResult);
    IFACEMETHODIMP put_Result(_In_ HRESULT hrResult);
    IFACEMETHODIMP Rotate();

    // IRotationItemFactory
    IFACEMETHODIMP Create(_COM_Outptr_ IRotationItem** ppri)
    {
        return CRotationItem::s_CreateInstance(L"", ppri);
    }

    static HRESULT s_CreateInstance(_In_ PCWSTR pszPath, _COM_Outptr_ IRotationItem** ppri);

private:
    ~CRotationItem();

private:
    long m_cRef = 1;
    PWSTR m_pszPath = nullptr;
    bool m_fWasRotated = false;
    bool m_fIsValidJPEG = false;
    bool m_fIsRotationLossless = true;
    UINT m_uOriginalOrientation = 1;
    HRESULT m_hrResult = S_FALSE;  // We init to S_FALSE which means Not Run Yet.  S_OK on success.  Otherwise an error code.
    CSRWLock m_lock;
    static UINT s_uTagOrientationPropSize;
};

// Maximum number of running worker threads
// We should never exceed the number of logical processors
#define MAX_ROTATION_WORKER_THREADS 64

// Minimum amount of work to schedule to a worker thread
#define MIN_ROTATION_WORK_SIZE 5

class CRotationManager : 
    public IRotationManager,
    public IRotationManagerEvents,
    public IRotationManagerDiagnostics
{
public:
    CRotationManager();

    // IUnknown
    IFACEMETHODIMP QueryInterface(_In_ REFIID riid, _COM_Outptr_ void** ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CRotationManager, IRotationManager),
            QITABENT(CRotationManager, IRotationManagerEvents),
            QITABENT(CRotationManager, IRotationManagerDiagnostics),
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

    // IRotationManager
    IFACEMETHODIMP Advise(_In_ IRotationManagerEvents* prme, _Out_ DWORD* pdwCookie);
    IFACEMETHODIMP UnAdvise(_In_ DWORD dwCookie);
    IFACEMETHODIMP Start();
    IFACEMETHODIMP Cancel();
    IFACEMETHODIMP Shutdown();
    IFACEMETHODIMP AddItem(_In_ IRotationItem* pri);
    IFACEMETHODIMP GetItem(_In_ UINT uIndex, _COM_Outptr_ IRotationItem** ppri);
    IFACEMETHODIMP GetItemCount(_Out_ UINT* puCount);
    IFACEMETHODIMP SetRotationItemFactory(_In_ IRotationItemFactory* prif);
    IFACEMETHODIMP GetRotationItemFactory(_In_ IRotationItemFactory** pprif);

    // IRotationManagerDiagnostics
    IFACEMETHODIMP get_MaxWorkerThreadCount(_Out_ UINT* puMaxThreadCount);
    IFACEMETHODIMP put_MaxWorkerThreadCount(_In_ UINT uMaxThreadCount);
    IFACEMETHODIMP get_WorkerThreadCount(_Out_ UINT* puThreadCount);
    IFACEMETHODIMP put_WorkerThreadCount(_In_ UINT uThreadCount);
    IFACEMETHODIMP get_MinItemsPerWorkerThread(_Out_ UINT* puMinItemsPerThread);
    IFACEMETHODIMP put_MinItemsPerWorkerThread(_In_ UINT uMinItemsPerThread);
    IFACEMETHODIMP get_ItemsPerWorkerThread(_Out_ UINT* puNumItemsPerThread);
    IFACEMETHODIMP put_ItemsPerWorkerThread(_In_ UINT uNumItemsPerThread);

    // IRotationManagerEvents
    IFACEMETHODIMP OnItemAdded(_In_ UINT uIndex);
    IFACEMETHODIMP OnItemProcessed(_In_ UINT uIndex);
    IFACEMETHODIMP OnProgress(_In_ UINT uCompleted, _In_ UINT uTotal);
    IFACEMETHODIMP OnCanceled();
    IFACEMETHODIMP OnCompleted();

    static HRESULT s_CreateInstance(_COM_Outptr_ IRotationManager** pprm);

private:
    ~CRotationManager();

    HRESULT _Init();
    void _Cleanup();
    HRESULT _PerformRotation();
    HRESULT _CreateWorkerThreads();
    HRESULT _GetWorkerThreadDimensions();

    void _ClearRotationItems();
    void _ClearEventHandlers();

    static DWORD WINAPI s_rotationWorkerThread(_In_ void* pv);

private:
    struct ROTATION_WORKER_THREAD_INFO
    {
        HANDLE hWorker;
        DWORD dwThreadId;
        UINT uCompletedItems;
        UINT uTotalItems;
    };

    struct ROTATION_MANAGER_EVENT
    {
        IRotationManagerEvents* prme;
        DWORD dwCookie;
    };

    long  m_cRef = 1;
    ROTATION_WORKER_THREAD_INFO m_workerThreadInfo[MAX_ROTATION_WORKER_THREADS];
    bool m_diagnosticsMode = false;
    HANDLE m_workerThreadHandles[MAX_ROTATION_WORKER_THREADS];
    UINT m_uWorkerThreadCount = 0;
    UINT m_uMaxWorkerThreadCount = MAX_ROTATION_WORKER_THREADS;
    UINT m_uMinItemsPerWorkerThread = 0;
    UINT m_uMaxItemsPerWorkerThread = 0;
    UINT m_uItemsPerWorkerThread = 0;
    ULONG_PTR m_gdiplusToken;
    HANDLE m_hCancelEvent = nullptr;
    HANDLE m_hStartEvent = nullptr;
    DWORD m_dwCookie = 0;
    CSRWLock m_lockEvents;
    CSRWLock m_lockItems;
    CComPtr<IRotationItemFactory> m_sprif;
    _Guarded_by_(m_lockEvents) std::vector<ROTATION_MANAGER_EVENT> m_rotationManagerEvents;
    _Guarded_by_(m_lockItems) std::vector<IRotationItem*> m_rotationItems;
};
