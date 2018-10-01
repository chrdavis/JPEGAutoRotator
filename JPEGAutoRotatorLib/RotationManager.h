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

    // IRotationItem
    IFACEMETHODIMP get_Path(_Outptr_ PWSTR* path);
    IFACEMETHODIMP put_Path(_In_ PCWSTR path);
    IFACEMETHODIMP get_WasRotated(_Out_ BOOL* wasRotated);
    IFACEMETHODIMP get_IsValidJPEG(_Out_ BOOL* isValidJPEG);
    IFACEMETHODIMP get_IsRotationLossless(_Out_ BOOL* isRotationLossless);
    IFACEMETHODIMP get_OriginalOrientation(_Out_ UINT* originalOrientation);
    IFACEMETHODIMP get_Result(_Out_ HRESULT* result);
    IFACEMETHODIMP put_Result(_In_ HRESULT result);
    IFACEMETHODIMP Load();
    IFACEMETHODIMP Rotate();

    // IRotationItemFactory
    IFACEMETHODIMP Create(_COM_Outptr_ IRotationItem** ppItem)
    {
        return CRotationItem::s_CreateInstance(L"", ppItem);
    }

    static HRESULT s_CreateInstance(_In_ PCWSTR path, _COM_Outptr_ IRotationItem** ppItem);

private:
    ~CRotationItem();

private:
    long m_refCount = 1;
    PWSTR m_path = nullptr;
    bool m_wasRotated = false;
    bool m_isValidJPEG = false;
    bool m_isRotationLossless = true;
    bool m_loaded = false;
    UINT m_originalOrientation = 1;
    HRESULT m_result = S_FALSE;  // We init to S_FALSE which means Not Run Yet.  S_OK on success.  Otherwise an error code.
    CSRWLock m_lock;
    static UINT s_tagOrientationPropSize;
};

// Maximum number of running worker threads
// We should never exceed the number of logical processors
#define MAX_ROTATION_WORKER_THREADS 64

// Minimum amount of work to schedule to a worker thread
#define MIN_ROTATION_WORK_SIZE 10

class CRotationManager : 
    public IRotationManager,
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
            QITABENT(CRotationManager, IRotationManagerDiagnostics),
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

    // IRotationManager
    IFACEMETHODIMP Advise(_In_ IRotationManagerEvents* pEvents, _Out_ DWORD* cookie);
    IFACEMETHODIMP UnAdvise(_In_ DWORD cookie);
    IFACEMETHODIMP Start();
    IFACEMETHODIMP Cancel();
    IFACEMETHODIMP Shutdown();
    IFACEMETHODIMP AddPath(_In_ PCWSTR path);
    IFACEMETHODIMP AddItem(_In_ IRotationItem* pItem);
    IFACEMETHODIMP GetItem(_In_ UINT index, _COM_Outptr_ IRotationItem** ppItem);
    IFACEMETHODIMP GetItemCount(_Out_ UINT* count);
    IFACEMETHODIMP SetRotationItemFactory(_In_ IRotationItemFactory* pItemFactory);
    IFACEMETHODIMP GetRotationItemFactory(_In_ IRotationItemFactory** ppItemFactory);

    // IRotationManagerDiagnostics
    IFACEMETHODIMP get_EnumerateSubFolders(_Out_ BOOL* enumSubFolders);
    IFACEMETHODIMP put_EnumerateSubFolders(_In_ BOOL enumSubFolders);
    IFACEMETHODIMP get_LosslessOnly(_Out_ BOOL* losslessOnly);
    IFACEMETHODIMP put_LosslessOnly(_In_ BOOL losslessOnly);
    IFACEMETHODIMP get_PreviewOnly(_Out_ BOOL* previewOnly);
    IFACEMETHODIMP put_PreviewOnly(_In_ BOOL previewOnly);
    IFACEMETHODIMP get_MaxWorkerThreadCount(_Out_ UINT* maxThreadCount);
    IFACEMETHODIMP put_MaxWorkerThreadCount(_In_ UINT maxThreadCount);
    IFACEMETHODIMP get_WorkerThreadCount(_Out_ UINT* threadCount);
    IFACEMETHODIMP put_WorkerThreadCount(_In_ UINT threadCount);
    IFACEMETHODIMP get_MinItemsPerWorkerThread(_Out_ UINT* minItemsPerThread);
    IFACEMETHODIMP put_MinItemsPerWorkerThread(_In_ UINT minItemsPerThread);
    IFACEMETHODIMP get_ItemsPerWorkerThread(_Out_ UINT* numItemsPerThread);
    IFACEMETHODIMP put_ItemsPerWorkerThread(_In_ UINT numItemsPerThread);

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

    bool _EnumeratePath(_In_ PCWSTR path, UINT depth = 0);

    bool _PathIsDotOrDotDot(_In_ PCWSTR path);
    bool _IsJPEG(_In_ PCWSTR path);

    HRESULT _OnItemAdded(_In_ UINT index);
    HRESULT _OnItemProcessed(_In_ UINT index);
    HRESULT _OnProgress(_In_ UINT completed, _In_ UINT total);
    HRESULT _OnCanceled();
    HRESULT _OnCompleted();

    static DWORD WINAPI s_rotationWorkerThread(_In_ void* pv);

private:
    struct ROTATION_WORKER_THREAD_INFO
    {
        HANDLE workerHandle;
        DWORD threadId;
        UINT processedItems;
        UINT totalItems;
    };

    struct ROTATION_MANAGER_EVENT
    {
        IRotationManagerEvents* pEvents;
        DWORD cookie;
    };

    long  m_refCount = 1;
    ROTATION_WORKER_THREAD_INFO m_workerThreadInfo[MAX_ROTATION_WORKER_THREADS];
    HANDLE m_workerThreadHandles[MAX_ROTATION_WORKER_THREADS];
    UINT m_workerThreadCount = 0;
    UINT m_maxWorkerThreadCount = MAX_ROTATION_WORKER_THREADS;
    UINT m_minItemsPerWorkerThread = MIN_ROTATION_WORK_SIZE;
    UINT m_itemsPerWorkerThread = 0;
    bool m_previewOnly = false;
    bool m_losslessOnly = false;
    bool m_enumSubFolders = true;
    ULONG_PTR m_gdiplusToken;
    HANDLE m_cancelEvent = nullptr;
    HANDLE m_startEvent = nullptr;
    DWORD m_cookie = 0;
    CSRWLock m_lockEvents;
    CSRWLock m_lockItems;
    CComPtr<IRotationItemFactory> m_spItemFactory;
    _Guarded_by_(m_lockEvents) std::vector<ROTATION_MANAGER_EVENT> m_rotationManagerEvents;
    _Guarded_by_(m_lockItems) std::vector<IRotationItem*> m_rotationItems;
};
