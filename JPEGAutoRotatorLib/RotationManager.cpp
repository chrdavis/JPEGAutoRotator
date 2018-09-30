#include "stdafx.h"
#include "RotationManager.h"
#include "helpers.h"
#include <gdiplus.h>

extern HINSTANCE g_hInst;

using namespace Gdiplus;

// Position in below table corresponds to current orientation.
// The value is the action to perform to correct it

// For a background on EXIF Orientation, see:
// http://www.impulseadventure.com/photo/exif-orientation.html

const RotateFlipType rotateFlipTable[] =
{
    RotateNoneFlipNone,
    RotateNoneFlipNone,
    RotateNoneFlipX,
    Rotate180FlipNone,
    RotateNoneFlipY,
    Rotate270FlipY,
    Rotate90FlipNone,
    Rotate90FlipY,
    Rotate270FlipNone
};

// Custom messages for rotation manager worker thread
enum
{
    ROTM_ROTI_ITEM_PROCESSED = (WM_APP + 1),  // Single rotation item processed
    ROTM_ROTI_CANCELED,                       // Rotation operation was canceled
    ROTM_ROTI_COMPLETE,                       // Worker thread completed
    ROTM_ENDTHREAD                            // End worker threads and exit manager
};

int GetEncoderClsid(PCWSTR format, CLSID* pClsid)
{
    UINT  num = 0;          // number of image encoders
    UINT  size = 0;         // size of the image encoder array in bytes

    ImageCodecInfo* pImageCodecInfo = nullptr;

    GetImageEncodersSize(&num, &size);
    if (size == 0)
    {
        return -1;  // Failure
    }

    pImageCodecInfo = (ImageCodecInfo*)(malloc(size));
    if (pImageCodecInfo == nullptr)
    {
        return -1;  // Failure
    }

    GetImageEncoders(num, size, pImageCodecInfo);

    for (UINT j = 0; j < num && j < size; ++j)
    {
        if (wcscmp(pImageCodecInfo[j].MimeType, format) == 0)
        {
            *pClsid = pImageCodecInfo[j].Clsid;
            free(pImageCodecInfo);
            return j;  // Success
        }
    }
    free(pImageCodecInfo);
    return -1;  // Failure
}

struct RotateWorkerThreadData
{
    UINT firstIndex;
    UINT lastIndex;
    DWORD managerThreadId;
    HANDLE hStartEvent;
    HANDLE cancelEvent;
    CComPtr<IRotationManager> sprm;
};

UINT CRotationItem::s_tagOrientationPropSize = 0;

CRotationItem::CRotationItem()
{
}

CRotationItem::~CRotationItem()
{
    CoTaskMemFree(m_path);
}

HRESULT CRotationItem::s_CreateInstance(_In_ PCWSTR path, _COM_Outptr_ IRotationItem** ppItem)
{
    *ppItem = nullptr;
    CRotationItem* pItem = new CRotationItem();
    HRESULT hr = pItem ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = pItem->put_Path(path);
        if (SUCCEEDED(hr))
        {
            hr = pItem->QueryInterface(IID_PPV_ARGS(ppItem));
        }
        pItem->Release();
    }
    return hr;
}

IFACEMETHODIMP CRotationItem::get_Path(_Outptr_ PWSTR* path)
{
    CSRWSharedAutoLock lock(&m_lock);
    *path = nullptr;
    HRESULT hr = m_path ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        hr = SHStrDup(m_path, path);
    }
    return hr;
}

IFACEMETHODIMP CRotationItem::put_Path(_In_ PCWSTR path)
{
    CSRWExclusiveAutoLock lock(&m_lock);
    HRESULT hr = path ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        CoTaskMemFree(m_path);
        hr = SHStrDup(path, &m_path);
    }
    return hr;
}

IFACEMETHODIMP CRotationItem::get_WasRotated(_Out_ BOOL* wasRotated)
{
    CSRWSharedAutoLock lock(&m_lock);
    *wasRotated = m_wasRotated;
    return S_OK;
}

IFACEMETHODIMP CRotationItem::get_IsValidJPEG(_Out_ BOOL* isValidJPEG)
{
    CSRWSharedAutoLock lock(&m_lock);
    *isValidJPEG = m_isValidJPEG;
    return S_OK;
}

IFACEMETHODIMP CRotationItem::get_IsRotationLossless(_Out_ BOOL* isRotationLossless)
{
    CSRWSharedAutoLock lock(&m_lock);
    *isRotationLossless = m_isRotationLossless;
    return S_OK;
}

IFACEMETHODIMP CRotationItem::get_OriginalOrientation(_Out_ UINT* originalOrientation)
{
    CSRWSharedAutoLock lock(&m_lock);
    *originalOrientation = m_originalOrientation;
    return S_OK;
}

IFACEMETHODIMP CRotationItem::get_Result(_Out_ HRESULT* result)
{
    CSRWSharedAutoLock lock(&m_lock);
    *result = m_result;
    return S_OK; 
}

IFACEMETHODIMP CRotationItem::put_Result(_In_ HRESULT result)
{
    CSRWExclusiveAutoLock lock(&m_lock);
    m_result = result;
    return S_OK;
}
IFACEMETHODIMP CRotationItem::Load()
{
    HRESULT hr = S_OK;
    CSRWExclusiveAutoLock lock(&m_lock);

    // Get image info before actual rotation
    if (!m_loaded)
    {
        hr = m_path ? S_OK : E_FAIL;
        if (SUCCEEDED(hr))
        {
            CComPtr<IStream> spstrm;
            hr = SHCreateStreamOnFile(m_path, STGM_READWRITE, &spstrm);
            if (SUCCEEDED(hr))
            {
                // Create a GDIPlus Image from the stream
                Image* pImage = Image::FromStream(spstrm);
                hr = (pImage && pImage->GetLastStatus() == Ok) ? S_OK : E_FAIL;
                if (SUCCEEDED(hr))
                {
                    GUID guidGdiplusFormat = { 0 };
                    if (pImage->GetRawFormat(&guidGdiplusFormat) == Ok &&
                        (guidGdiplusFormat == ImageFormatJPEG))
                    {
                        // Dimensions must be multiples of 8 for the rotation to be lossless
                        m_isRotationLossless = ((pImage->GetHeight() % 8 == 0) && (pImage->GetWidth() % 8 == 0));
                        m_isValidJPEG = true;

                        // GetPropertyItemSize can be costly so try to only get this once
                        if (s_tagOrientationPropSize == 0)
                        {
                            s_tagOrientationPropSize = pImage->GetPropertyItemSize(PropertyTagOrientation);
                        }

                        if (pImage->GetLastStatus() == Ok)
                        {
                            // Get the orientation property if it exists
                            const int LOCAL_BUFFER_SIZE = 32;
                            BYTE buffer[LOCAL_BUFFER_SIZE];
                            PropertyItem* pItem = (PropertyItem*)buffer;

                            if (s_tagOrientationPropSize < LOCAL_BUFFER_SIZE &&
                                pImage->GetPropertyItem(PropertyTagOrientation, s_tagOrientationPropSize, pItem) == Gdiplus::Ok &&
                                pItem->type == PropertyTagTypeShort)
                            {
                                m_originalOrientation = static_cast<UINT>(*((USHORT*)pItem->value));
                            }
                        }
                    }
                }

                if (pImage)
                {
                    delete pImage;
                    pImage = nullptr;
                }
            }
        }

        m_loaded = true;
    }
    return hr;
}

IFACEMETHODIMP CRotationItem::Rotate()
{
    CSRWExclusiveAutoLock lock(&m_lock);
    HRESULT hr = (m_loaded && m_isValidJPEG) ? S_OK : E_FAIL; 
    if (SUCCEEDED(hr))
    {
        hr = m_path ? S_OK : E_FAIL;
        if (SUCCEEDED(hr))
        {
            CComPtr<IStream> spstrm;
            hr = SHCreateStreamOnFile(m_path, STGM_READWRITE, &spstrm);
            if (SUCCEEDED(hr))
            {
                // Create a GDIPlus Image from the stream
                Image* pImage = Image::FromStream(spstrm);
                hr = (pImage && pImage->GetLastStatus() == Ok) ? S_OK : E_FAIL;
                if (SUCCEEDED(hr))
                {
                    // Ensure valid orientation range.  If not in range do nothing and return success.
                    if (m_originalOrientation <= 8 && m_originalOrientation >= 2)
                    {
                        hr = (pImage->RotateFlip(rotateFlipTable[m_originalOrientation]) == Ok) ? S_OK : E_FAIL;
                        // Remove the orientation tag
                        pImage->RemovePropertyItem(PropertyTagOrientation);
                        // Save back to original location
                        IStream_Reset(spstrm);
                        CLSID jpgClsid;
                        GetEncoderClsid(L"image/jpeg", &jpgClsid);
                        hr = (pImage->Save(spstrm, &jpgClsid, nullptr) == Ok) ? S_OK : E_FAIL;
                        if (SUCCEEDED(hr))
                        {
                            m_wasRotated = true;
                        }
                    }
                }

                if (pImage)
                {
                    delete pImage;
                    pImage = nullptr;
                }
            }
        }
    }
    return hr;
}

CRotationManager::CRotationManager()
{
    // Initialize the maximum number of worker threads to the number of cores available
    m_maxWorkerThreadCount = min(GetLogicalProcessorCount(), MAX_ROTATION_WORKER_THREADS);
}

CRotationManager::~CRotationManager()
{
    _Cleanup();
}

IFACEMETHODIMP CRotationManager::Advise(_In_ IRotationManagerEvents* pEvents, _Out_ DWORD* cookie)
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);
    m_cookie++;
    ROTATION_MANAGER_EVENT rmestruct;
    rmestruct.cookie = m_cookie;
    rmestruct.pEvents = pEvents;
    pEvents->AddRef();
    m_rotationManagerEvents.push_back(rmestruct);
    
    *cookie = m_cookie;

    return S_OK;
}

IFACEMETHODIMP CRotationManager::UnAdvise(_In_ DWORD cookie)
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<ROTATION_MANAGER_EVENT>::iterator it = m_rotationManagerEvents.begin(); it != m_rotationManagerEvents.end(); ++it)
    {
        if (it->cookie == cookie)
        {
            it->cookie = 0;
            if (it->pEvents)
            {
                it->pEvents->Release();
                it->pEvents = nullptr;
            }
            break;
        }
    }

    return S_OK;
}

IFACEMETHODIMP CRotationManager::Start()
{
    // Ensure any previous runs are cancelled
    Cancel();

    ResetEvent(m_startEvent);
    return _PerformRotation();
}

IFACEMETHODIMP CRotationManager::Cancel()
{
    SetEvent(m_startEvent);
    SetEvent(m_cancelEvent);
    return S_OK;
}

IFACEMETHODIMP CRotationManager::Shutdown()
{
    // Cancel any running operations
    Cancel();
    // Kickoff cleanup
    _Cleanup();

    return S_OK;
}

IFACEMETHODIMP CRotationManager::AddItem(_In_ IRotationItem* pItem)
{
    // Scope lock
    {
        CSRWExclusiveAutoLock lock(&m_lockItems);
        m_rotationItems.push_back(pItem);
        pItem->AddRef();
    }

    OnItemAdded(static_cast<UINT>(m_rotationItems.size() - 1));

    return S_OK;
}

IFACEMETHODIMP CRotationManager::GetItem(_In_ UINT index, _COM_Outptr_ IRotationItem** ppItem)
{
    *ppItem = nullptr;
    CSRWSharedAutoLock lock(&m_lockItems);
    HRESULT hr = E_FAIL;
    if (index < m_rotationItems.size())
    {
        *ppItem = m_rotationItems.at(index);
        (*ppItem)->AddRef();
        hr = S_OK;
    }

    return hr;
}

IFACEMETHODIMP CRotationManager::GetItemCount(_Out_ UINT* count)
{
    CSRWSharedAutoLock lock(&m_lockItems);
    *count = static_cast<UINT>(m_rotationItems.size());
    return S_OK;
}

IFACEMETHODIMP CRotationManager::SetRotationItemFactory(_In_ IRotationItemFactory* pItemFactory)
{
    m_spItemFactory = pItemFactory;
    return S_OK;
}

IFACEMETHODIMP CRotationManager::GetRotationItemFactory(_In_ IRotationItemFactory** ppItemFactory)
{
    HRESULT hr = E_FAIL;
    if (m_spItemFactory)
    {
        hr = S_OK;
        *ppItemFactory = m_spItemFactory;
        (*ppItemFactory)->AddRef();
    }
    return hr;
}

// IRotationManagerEvents
IFACEMETHODIMP CRotationManager::OnItemAdded(_In_ UINT index)
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<ROTATION_MANAGER_EVENT>::iterator it = m_rotationManagerEvents.begin(); it != m_rotationManagerEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnItemAdded(index);
        }
    }

    return S_OK;
}

IFACEMETHODIMP CRotationManager::OnItemProcessed(_In_ UINT index)
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<ROTATION_MANAGER_EVENT>::iterator it = m_rotationManagerEvents.begin(); it != m_rotationManagerEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnItemProcessed(index);
        }
    }

    return S_OK;
}

IFACEMETHODIMP CRotationManager::OnProgress(_In_ UINT completed, _In_ UINT total)
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<ROTATION_MANAGER_EVENT>::iterator it = m_rotationManagerEvents.begin(); it != m_rotationManagerEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnProgress(completed, total);
        }
    }

    return S_OK;
}

IFACEMETHODIMP CRotationManager::OnCanceled()
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<ROTATION_MANAGER_EVENT>::iterator it = m_rotationManagerEvents.begin(); it != m_rotationManagerEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnCanceled();
        }
    }

    return S_OK;
}

IFACEMETHODIMP CRotationManager::OnCompleted()
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    for (std::vector<ROTATION_MANAGER_EVENT>::iterator it = m_rotationManagerEvents.begin(); it != m_rotationManagerEvents.end(); ++it)
    {
        if (it->pEvents)
        {
            it->pEvents->OnCompleted();
        }
    }

    return S_OK;
}

// IRotationManagerDiagnostics
IFACEMETHODIMP CRotationManager::get_EnumerateSubFolders(_Out_ BOOL* enumSubFolders)
{
    *enumSubFolders = m_enumSubFolders;
    return S_OK;
}

IFACEMETHODIMP CRotationManager::put_EnumerateSubFolders(_In_ BOOL enumSubFolders)
{
    m_enumSubFolders = enumSubFolders;
    return S_OK;
}

IFACEMETHODIMP CRotationManager::get_LosslessOnly(_Out_ BOOL* losslessOnly)
{
    *losslessOnly = m_losslessOnly;
    return S_OK;
}

IFACEMETHODIMP CRotationManager::put_LosslessOnly(_In_ BOOL losslessOnly)
{
    m_losslessOnly = losslessOnly;
    return S_OK;
}

IFACEMETHODIMP CRotationManager::get_PreviewOnly(_Out_ BOOL* previewOnly)
{
    *previewOnly = m_previewOnly;
    return S_OK;
}

IFACEMETHODIMP CRotationManager::put_PreviewOnly(_In_ BOOL previewOnly)
{
    m_previewOnly = previewOnly;
    return S_OK;
}

IFACEMETHODIMP CRotationManager::get_MaxWorkerThreadCount(_Out_ UINT* maxThreadCount)
{
    *maxThreadCount = m_maxWorkerThreadCount;
    return S_OK;
}

IFACEMETHODIMP CRotationManager::put_MaxWorkerThreadCount(_In_ UINT maxThreadCount)
{
    m_maxWorkerThreadCount = maxThreadCount;
    return S_OK;
}

IFACEMETHODIMP CRotationManager::get_WorkerThreadCount(_Out_ UINT* threadCount)
{
    *threadCount = m_workerThreadCount;
    return S_OK;
}

IFACEMETHODIMP CRotationManager::put_WorkerThreadCount(_In_ UINT threadCount)
{
    m_workerThreadCount = threadCount;
    return S_OK;
}

IFACEMETHODIMP CRotationManager::get_MinItemsPerWorkerThread(_Out_ UINT* minItemsPerThread)
{
    *minItemsPerThread = m_minItemsPerWorkerThread;
    return S_OK;
}

IFACEMETHODIMP CRotationManager::put_MinItemsPerWorkerThread(_In_ UINT minItemsPerThread)
{
    m_minItemsPerWorkerThread = minItemsPerThread;
    return S_OK;
}

IFACEMETHODIMP CRotationManager::get_ItemsPerWorkerThread(_Out_ UINT* numItemsPerThread)
{
    *numItemsPerThread = m_itemsPerWorkerThread;
    return S_OK;
}

IFACEMETHODIMP CRotationManager::put_ItemsPerWorkerThread(_In_ UINT numItemsPerThread)
{
    m_itemsPerWorkerThread = numItemsPerThread;
    return S_OK;
}

HRESULT CRotationManager::s_CreateInstance(_COM_Outptr_ IRotationManager** pprm)
{
    *pprm = nullptr;
    CRotationManager *prm = new CRotationManager();
    HRESULT hr = prm ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = prm->_Init();
        if (SUCCEEDED(hr))
        {
            hr = prm->QueryInterface(IID_PPV_ARGS(pprm));
        }
        prm->Release();
    }
    return hr;
}

HRESULT CRotationManager::_PerformRotation()
{
    // Create worker threads which will message us progress and
    // completion.
    HRESULT hr = _CreateWorkerThreads();
    if (SUCCEEDED(hr))
    {
        UINT processedCount = 0;
        UINT totalItems = 0;
        UINT threadCompleteCount = 0;
        GetItemCount(&totalItems);

        ResetEvent(m_cancelEvent);

        // Signal the worker thread that they can start working. We needed to wait until we
        // were ready to process thread messages.
        SetEvent(m_startEvent);
        
        while (true)
        {
            // Check if all running threads have exited
            if (WaitForMultipleObjects(m_workerThreadCount, m_workerThreadHandles, TRUE, 0) == WAIT_OBJECT_0)
            {
                // Ensure we process remaining messages from worker threads
                if (threadCompleteCount == m_workerThreadCount)
                {
                    break;
                }
            }

            MSG msg;
            while (PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE))
            {
                if (msg.message == ROTM_ROTI_ITEM_PROCESSED)
                {
                    processedCount++;
                    OnItemProcessed(static_cast<UINT>(msg.lParam));
                    OnProgress(processedCount, totalItems);
                }
                else if (msg.message == ROTM_ROTI_CANCELED)
                {
                    OnCanceled();
                }
                else if (msg.message == ROTM_ROTI_COMPLETE)
                {
                    threadCompleteCount++;
                    // Worker thread completed
                    // Break out of the loop and check if all threads are done
                    break;
                }
                else
                {
                    TranslateMessage(&msg);
                    DispatchMessage(&msg);
                }
            }
        }

        OnCompleted();
    }

    return 0;
}

HRESULT CRotationManager::_CreateWorkerThreads()
{
    HRESULT hr = _GetWorkerThreadDimensions();
    if (SUCCEEDED(hr))
    {
        UINT firstIndex = 0;
        UINT lastIndex = m_itemsPerWorkerThread;
        UINT totalItems = 0;
        GetItemCount(&totalItems);

        // Create the worker threads
        for (UINT u = 0; SUCCEEDED(hr) && u < m_workerThreadCount; u++)
        {
            RotateWorkerThreadData* prwtd = new RotateWorkerThreadData;
            hr = prwtd ? S_OK : E_OUTOFMEMORY;
            if (SUCCEEDED(hr))
            {
                prwtd->firstIndex = firstIndex;
                prwtd->lastIndex = lastIndex;
                prwtd->managerThreadId = GetCurrentThreadId();
                prwtd->hStartEvent = m_startEvent;
                prwtd->cancelEvent = m_cancelEvent;
                prwtd->sprm = this;
                m_workerThreadInfo[u].workerHandle = CreateThread(nullptr, 0, s_rotationWorkerThread, prwtd, 0, &m_workerThreadInfo[u].threadId);
                m_workerThreadHandles[u] = m_workerThreadInfo[u].workerHandle;
                hr = (m_workerThreadInfo[u].workerHandle) ? S_OK : E_FAIL;
                if (SUCCEEDED(hr))
                {
                    // increment the indices for the next thread
                    firstIndex = ++lastIndex;
                    lastIndex = min((firstIndex + m_itemsPerWorkerThread), totalItems - 1);
                }
                else
                {
                    delete prwtd;
                }
            }
        }

        if (FAILED(hr))
        {
            Cancel();
        }
    }

    return hr;
}

HRESULT CRotationManager::_GetWorkerThreadDimensions()
{
    UINT totalItems = 0;
    HRESULT hr = GetItemCount(&totalItems);
    if (SUCCEEDED(hr))
    {
        hr = (totalItems > 0) ? S_OK : E_FAIL;
        if (SUCCEEDED(hr))
        {
            // Determine the best number of threads based on the number of items to rotate
            // and our minimum number of items per worker thread

            // Have we already determined our worker thread count?
            if (m_workerThreadCount == 0)
            {
                // How many work item sized chunks of items do we have?  Ignore remainder for now.
                // Think of a work chunk as the minimum amount of items you would put on a worker
                // thread.  Less than this would result in perf degradation due to the cost of 
                // standing up the worker thread instead of running it with other work on another thread.
                UINT totalWorkChunks = max(1, (totalItems / max(1, m_minItemsPerWorkerThread)));

                // Ensure we don't exceed m_maxWorkerThreadCount. Also ensure if we have less than
                // m_minItemsPerWorkerThread items that we default to 1 worker thread.
                m_workerThreadCount = max(1, min(totalWorkChunks, m_maxWorkerThreadCount));
            }

            // Have we already determined the number of items per worker thread?
            if (m_itemsPerWorkerThread == 0)
            {
                // Now determine the number of items per worker thread
                m_itemsPerWorkerThread = (totalItems / m_workerThreadCount);
                if ((totalItems % m_workerThreadCount) > 0)
                {
                    // Take into account the leftover items.  This will spread the remainder over
                    // some of the threads.
                    m_itemsPerWorkerThread++;
                }
            }
        }
    }

    return hr;
}

DWORD WINAPI CRotationManager::s_rotationWorkerThread(_In_ void* pv)
{
    HRESULT hr = CoInitializeEx(NULL, 0);
    if (SUCCEEDED(hr))
    {
        RotateWorkerThreadData* prwtd = reinterpret_cast<RotateWorkerThreadData*>(pv);
        if (prwtd)
        {
            // Wait to be told we can begin
            if (WaitForSingleObject(prwtd->hStartEvent, INFINITE) == WAIT_OBJECT_0)
            {
                CComPtr<IRotationManagerDiagnostics> sprmd;
                if (SUCCEEDED(prwtd->sprm->QueryInterface(IID_PPV_ARGS(&sprmd))))
                {
                    BOOL previewOnly = FALSE;
                    BOOL losslessOnly = FALSE;
                    sprmd->get_PreviewOnly(&previewOnly);
                    sprmd->get_LosslessOnly(&losslessOnly);
                    for (UINT u = prwtd->firstIndex; u <= prwtd->lastIndex; u++)
                    {
                        // Check if cancel event is signaled
                        if (WaitForSingleObject(prwtd->cancelEvent, 0) == WAIT_OBJECT_0)
                        {
                            // Canceled from manager
                            // Send the manager thread the canceled message
                            PostThreadMessage(prwtd->managerThreadId, ROTM_ROTI_CANCELED, GetCurrentThreadId(), 0);
                            break;
                        }

                        CComPtr<IRotationItem> spItem;
                        if (SUCCEEDED(prwtd->sprm->GetItem(u, &spItem)))
                        {
                            HRESULT hrWork;
                            spItem->get_Result(&hrWork);
                            // S_FALSE means we have not processed this item yet
                            if (hrWork == S_FALSE)
                            {
                                // Perform the rotation
                                hrWork = spItem->Load();
                                if (SUCCEEDED(hrWork))
                                {
                                    BOOL isLossless = FALSE;
                                    BOOL isValidJPEG = FALSE;
                                    spItem->get_IsValidJPEG(&isValidJPEG);
                                    spItem->get_IsRotationLossless(&isLossless);
                                    if (!previewOnly && isValidJPEG && (!losslessOnly || !isLossless))
                                    {
                                        UINT orientation = 0;
                                        spItem->get_OriginalOrientation(&orientation);
                                        // Does the image actually have a orientation property that 
                                        // necessitates rotation?
                                        if (orientation <= 8 && orientation >= 2)
                                        {
                                            hrWork = spItem->Rotate();
                                        }
                                    }
                                }

                                spItem->put_Result(hrWork);
                                // Send the manager thread the rotation item processed message
                                PostThreadMessage(prwtd->managerThreadId, ROTM_ROTI_ITEM_PROCESSED, GetCurrentThreadId(), u);
                            }
                        }
                    }
                }
            }

            // Send the manager thread the completion message
            PostThreadMessage(prwtd->managerThreadId, ROTM_ROTI_COMPLETE, GetCurrentThreadId(), 0);

            delete prwtd;
        }
        CoUninitialize();
    }

    return 0;
}

HRESULT CRotationManager::_Init()
{
    // Initialize GDIPlus now. All of our GDIPlus usage is in CRenameItem.
    GdiplusStartupInput gdiplusStartupInput;
    GdiplusStartup(&m_gdiplusToken, &gdiplusStartupInput, nullptr);

    ZeroMemory(m_workerThreadInfo, sizeof(m_workerThreadInfo));

    // Event used to signal worker thread that it can start
    m_startEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);
    // Event used to signal worker thread in the event of a cancel
    m_cancelEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);

    HRESULT hr = (m_startEvent && m_cancelEvent) ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        // Create the default IRotationItemFactory
        CComPtr<IRotationItem> spItem;
        hr = CRotationItem::s_CreateInstance(L"", &spItem);
        if (SUCCEEDED(hr))
        {
            hr = spItem->QueryInterface(IID_PPV_ARGS(&m_spItemFactory));
        }
    }

    if (FAILED(hr))
    {
        _Cleanup();
    }

    return hr;
}

void CRotationManager::_ClearEventHandlers()
{
    CSRWExclusiveAutoLock lock(&m_lockEvents);

    // Cleanup event handlers
    for (std::vector<ROTATION_MANAGER_EVENT>::iterator it = m_rotationManagerEvents.begin(); it != m_rotationManagerEvents.end(); ++it)
    {
        it->cookie = 0;
        if (it->pEvents)
        {
            it->pEvents->Release();
            it->pEvents = nullptr;
        }
    }

    m_rotationManagerEvents.clear();
}

void CRotationManager::_ClearRotationItems()
{
    CSRWExclusiveAutoLock lock(&m_lockItems);

    // Cleanup rotation items
    for (std::vector<IRotationItem*>::iterator it = m_rotationItems.begin(); it != m_rotationItems.end(); ++it)
    {
        IRotationItem* pItem = *it;
        pItem->Release();
    }

    m_rotationItems.clear();
}

void CRotationManager::_Cleanup()
{
    for (UINT u = 0; u < m_workerThreadCount; u++)
    {
        CloseHandle(m_workerThreadInfo[u].workerHandle);
        m_workerThreadInfo[u].workerHandle = nullptr;
    }

    CloseHandle(m_startEvent);
    m_startEvent = nullptr;

    CloseHandle(m_cancelEvent);
    m_cancelEvent = nullptr;

    // Done with GDIPlus so shutdown now
    GdiplusShutdown(m_gdiplusToken);

    _ClearEventHandlers();
    _ClearRotationItems();
}

