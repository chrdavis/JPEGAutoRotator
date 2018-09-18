#include "stdafx.h"
#include "RotationManager.h"
#include <gdiplus.h>

extern void DllAddRef();
extern void DllRelease();

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
    ROTM_ROTI_ROTATED = (WM_APP + 1),  // Single rotation item finished
    ROTM_ROTI_CANCELED,                // Rotation operation was canceled
    ROTM_ROTI_COMPLETE,                // Worker thread completed
    ROTM_ENDTHREAD                     // End worker threads and exit manager
};

int GetEncoderClsid(const WCHAR* format, CLSID* pClsid)
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
    UINT uFirstIndex;
    UINT uLastIndex;
    DWORD dwManagerThreadId;
    HANDLE hStartEvent;
    HANDLE hCancelEvent;
    CComPtr<IRotationManager> sprm;
};

CRotationItem::CRotationItem()
{
}

CRotationItem::~CRotationItem()
{
    CoTaskMemFree(m_pszPath);
}

HRESULT CRotationItem::s_CreateInstance(__in PCWSTR pszPath, __deref_out IRotationItem** ppri)
{
    *ppri = nullptr;
    CRotationItem* pri = new CRotationItem();
    HRESULT hr = pri ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = pri->put_Path(pszPath);
        if (SUCCEEDED(hr))
        {
            hr = pri->QueryInterface(IID_PPV_ARGS(ppri));
        }
        pri->Release();
    }
    return hr;
}

IFACEMETHODIMP CRotationItem::get_Path(__deref_out PWSTR* ppszPath)
{
    *ppszPath = nullptr;
    HRESULT hr = m_pszPath ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        hr = SHStrDup(m_pszPath, ppszPath);
    }
    return hr;
}

IFACEMETHODIMP CRotationItem::put_Path(__in PCWSTR pszPath)
{
    HRESULT hr = pszPath ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        CoTaskMemFree(m_pszPath);
        hr = SHStrDup(pszPath, &m_pszPath);
    }
    return hr;
}

IFACEMETHODIMP CRotationItem::get_WasRotated(__out BOOL* pfWasRotated)
{
    *pfWasRotated = m_fWasRotated;
    return S_OK;
}

IFACEMETHODIMP CRotationItem::get_IsValidJPEG(__out BOOL* pfIsValidJPEG)
{
    *pfIsValidJPEG = m_fIsValidJPEG;
    return S_OK;
}

IFACEMETHODIMP CRotationItem::get_IsRotationLossless(__out BOOL* pfIsRotationLossless)
{
    *pfIsRotationLossless = m_fIsRotationLossless;
    return S_OK;
}

IFACEMETHODIMP CRotationItem::get_OriginalOrientation(__out UINT* puOriginalOrientation)
{
    *puOriginalOrientation = m_uOriginalOrientation;
    return S_OK;
}

IFACEMETHODIMP CRotationItem::get_Result(__out HRESULT* phrResult)
{
    *phrResult = m_hrResult;
    return S_OK; 
}

IFACEMETHODIMP CRotationItem::put_Result(__in HRESULT hrResult)
{
    m_hrResult = hrResult;
    return S_OK;
}

IFACEMETHODIMP CRotationItem::Rotate()
{
    HRESULT hr = m_pszPath ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        CComPtr<IStream> spstrm;
        hr = SHCreateStreamOnFile(m_pszPath, STGM_READWRITE, &spstrm);
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
                    m_fIsRotationLossless = ((pImage->GetHeight() % 8 == 0) && (pImage->GetWidth() % 8 == 0));
                    m_fIsValidJPEG = true;
                    const int LOCAL_BUFFER_SIZE = 32;
                    BYTE buffer[LOCAL_BUFFER_SIZE];
                    PropertyItem* pItem = (PropertyItem*)buffer;
                    UINT uSize = pImage->GetPropertyItemSize(PropertyTagOrientation);
                    if (pImage->GetLastStatus() == Ok)
                    {
                        if (uSize < LOCAL_BUFFER_SIZE &&
                            pImage->GetPropertyItem(PropertyTagOrientation, uSize, pItem) == Gdiplus::Ok &&
                            pItem->type == PropertyTagTypeShort)
                        {
                            m_uOriginalOrientation = static_cast<UINT>(*((USHORT*)pItem->value));
                            // Ensure valid orientation range.  If not in range do nothing and return success.
                            if (m_uOriginalOrientation <= 8 && m_uOriginalOrientation >= 2)
                            {
                                hr = (pImage->RotateFlip(rotateFlipTable[m_uOriginalOrientation]) == Ok) ? S_OK : E_FAIL;
                                if (SUCCEEDED(hr))
                                {
                                    // Remove the orientation tag
                                    pImage->RemovePropertyItem(PropertyTagOrientation);
                                    // Save back to original location
                                    IStream_Reset(spstrm);
                                    CLSID jpgClsid;
                                    GetEncoderClsid(L"image/jpeg", &jpgClsid);
                                    hr = (pImage->Save(spstrm, &jpgClsid, nullptr) == Ok) ? S_OK : E_FAIL;
                                    if (SUCCEEDED(hr))
                                    {
                                        m_fWasRotated = true;
                                    }
                                }
                            }
                        }
                    }
                }

                delete pImage;
            }
        }
    }
    return hr;
}

CRotationManager::CRotationManager()
{
    DllAddRef();

    InitializeSRWLock(&m_lockItems);
    InitializeSRWLock(&m_lockEvents);
}

CRotationManager::~CRotationManager()
{
    _Cleanup();
    DllRelease();
}

IFACEMETHODIMP CRotationManager::Advise(__in IRotationManagerEvents* prme, __out DWORD* pdwCookie)
{
    AcquireSRWLockExclusive(&m_lockEvents);
    m_dwCookie++;
    ROTATION_MANAGER_EVENT rmestruct;
    rmestruct.dwCookie = m_dwCookie;
    rmestruct.prme = prme;
    prme->AddRef();
    m_rotationManagerEvents.push_back(rmestruct);
    
    *pdwCookie = m_dwCookie;

    ReleaseSRWLockExclusive(&m_lockEvents);
    
    return S_OK;
}

IFACEMETHODIMP CRotationManager::UnAdvise(__in DWORD dwCookie)
{
    AcquireSRWLockExclusive(&m_lockEvents);

    for (std::vector<ROTATION_MANAGER_EVENT>::iterator it = m_rotationManagerEvents.begin(); it != m_rotationManagerEvents.end(); ++it)
    {
        if (it->dwCookie == dwCookie)
        {
            it->dwCookie = 0;
            if (it->prme)
            {
                it->prme->Release();
                it->prme = nullptr;
            }
            break;
        }
    }

    ReleaseSRWLockExclusive(&m_lockEvents);

    return S_OK;
}

IFACEMETHODIMP CRotationManager::Start()
{
    // Ensure any previous runs are cancelled
    Cancel();

    ResetEvent(m_hStartEvent);
    return _PerformRotation();
}

IFACEMETHODIMP CRotationManager::Cancel()
{
    SetEvent(m_hStartEvent);
    SetEvent(m_hCancelEvent);
    return S_OK;
}

IFACEMETHODIMP CRotationManager::AddItem(__in IRotationItem* pri)
{
    AcquireSRWLockExclusive(&m_lockItems);

    m_rotationItems.push_back(pri);
    pri->AddRef();

    ReleaseSRWLockExclusive(&m_lockItems);

    OnAdded(static_cast<UINT>(m_rotationItems.size() - 1));

    return S_OK;
}

IFACEMETHODIMP CRotationManager::GetItem(__in UINT uIndex, __deref_out IRotationItem** ppri)
{
    *ppri = nullptr;
    AcquireSRWLockShared(&m_lockItems);
    HRESULT hr = E_FAIL;
    if (uIndex < m_rotationItems.size())
    {
        *ppri = m_rotationItems.at(uIndex);
        (*ppri)->AddRef();
        hr = S_OK;
    }

    ReleaseSRWLockShared(&m_lockItems);
    return hr;
}

IFACEMETHODIMP CRotationManager::GetItemCount(__out UINT* puCount)
{
    AcquireSRWLockShared(&m_lockItems);
    *puCount = static_cast<UINT>(m_rotationItems.size());
    ReleaseSRWLockShared(&m_lockItems);
    return S_OK;
}


// IRotationManagerEvents
IFACEMETHODIMP CRotationManager::OnAdded(__in UINT uIndex)
{
    AcquireSRWLockExclusive(&m_lockEvents);

    for (std::vector<ROTATION_MANAGER_EVENT>::iterator it = m_rotationManagerEvents.begin(); it != m_rotationManagerEvents.end(); ++it)
    {
        if (it->prme)
        {
            it->prme->OnAdded(uIndex);
        }
    }

    ReleaseSRWLockExclusive(&m_lockEvents);

    return S_OK;
}

IFACEMETHODIMP CRotationManager::OnRotated(__in UINT uIndex)
{
    AcquireSRWLockExclusive(&m_lockEvents);

    for (std::vector<ROTATION_MANAGER_EVENT>::iterator it = m_rotationManagerEvents.begin(); it != m_rotationManagerEvents.end(); ++it)
    {
        if (it->prme)
        {
            it->prme->OnRotated(uIndex);
        }
    }

    ReleaseSRWLockExclusive(&m_lockEvents);
    return S_OK;
}

IFACEMETHODIMP CRotationManager::OnProgress(__in UINT uCompleted, __in UINT uTotal)
{
    AcquireSRWLockExclusive(&m_lockEvents);

    for (std::vector<ROTATION_MANAGER_EVENT>::iterator it = m_rotationManagerEvents.begin(); it != m_rotationManagerEvents.end(); ++it)
    {
        if (it->prme)
        {
            it->prme->OnProgress(uCompleted, uTotal);
        }
    }

    ReleaseSRWLockExclusive(&m_lockEvents);

    return S_OK;
}

IFACEMETHODIMP CRotationManager::OnCanceled()
{
    AcquireSRWLockExclusive(&m_lockEvents);

    for (std::vector<ROTATION_MANAGER_EVENT>::iterator it = m_rotationManagerEvents.begin(); it != m_rotationManagerEvents.end(); ++it)
    {
        if (it->prme)
        {
            it->prme->OnCanceled();
        }
    }

    ReleaseSRWLockExclusive(&m_lockEvents);

    return S_OK;
}

IFACEMETHODIMP CRotationManager::OnCompleted()
{
    AcquireSRWLockExclusive(&m_lockEvents);

    for (std::vector<ROTATION_MANAGER_EVENT>::iterator it = m_rotationManagerEvents.begin(); it != m_rotationManagerEvents.end(); ++it)
    {
        if (it->prme)
        {
            it->prme->OnCompleted();
        }
    }

    ReleaseSRWLockExclusive(&m_lockEvents);

    return S_OK;
}


HRESULT CRotationManager::s_CreateInstance(__deref_out IRotationManager** pprm)
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
        UINT uCompleted = 0;
        UINT uTotalItems = static_cast<UINT>(m_rotationItems.size());

        ResetEvent(m_hCancelEvent);
        // Signal the worker thread that they can start working. We needed to wait until we
        // were ready to process thread messages.
        SetEvent(m_hStartEvent);
        bool fDone = false;
        while (!fDone)
        {
            // Check if all running threads have exited
            if (WaitForMultipleObjects(m_uWorkerThreadCount, m_workerThreadHandles, TRUE, 10) == WAIT_OBJECT_0)
            {
                fDone = true;
            }

            MSG msg;
            while(PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE))
            {
                // If we got the "operation complete" message
                if (msg.message == ROTM_ENDTHREAD)
                {
                    // Break out of the loop and end the thread
                    fDone = true;
                }
                else if (msg.message == ROTM_ROTI_ROTATED)
                {
                    uCompleted++;
                    OnRotated(static_cast<UINT>(msg.lParam));
                    OnProgress(uCompleted, uTotalItems);
                }
                else if (msg.message == ROTM_ROTI_CANCELED)
                {
                    OnCanceled();
                }
                else if (msg.message == ROTM_ROTI_COMPLETE)
                {
                    // Worker thread completed
                    // Break out of the loop and end the thread
                    fDone = true;
                }
                else
                {
                    TranslateMessage(&msg);
                    DispatchMessage(&msg);
                }
            }

            OnCompleted();
        }
    }

    return 0;
}

HRESULT CRotationManager::_CreateWorkerThreads()
{
    UINT uMaxWorkerThreads = min(s_GetLogicalProcessorCount(), MAX_ROTATION_WORKER_THREADS);
    UINT uTotalItems = static_cast<UINT>(m_rotationItems.size());

    HRESULT hr = (uTotalItems > 0) ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        // Determine the best number of threads based on the number of items to rotate
        // and our minimum number of items per worker thread

        // How many work item sized amounts of items do we have?
        UINT uTotalWorkItems = max(1, (uTotalItems / MIN_ROTATION_WORK_SIZE));

        // How many worker threads do we require, with a minimum being the max worker threads value
        UINT uIdealWorkerThreadCount = (uTotalWorkItems / uMaxWorkerThreads);
        if ((uTotalWorkItems % uMaxWorkerThreads) > 0)
        {
            uIdealWorkerThreadCount++;
        }

        // Ensure we don't exceed uMaxWorkerThreads
        m_uWorkerThreadCount = min(uIdealWorkerThreadCount, uMaxWorkerThreads);

        // Now determine the number of items per worker
        UINT uItemsPerWorker = (uTotalItems / m_uWorkerThreadCount);
        if (uItemsPerWorker == 0)
        {
            uItemsPerWorker = uTotalItems % m_uWorkerThreadCount;
        }

        UINT uFirstIndex = 0;
        UINT uLastIndex = uItemsPerWorker;

        // Create the worker threads
        for (UINT u = 0; SUCCEEDED(hr) && u < m_uWorkerThreadCount; u++)
        {
            RotateWorkerThreadData* prwtd = new RotateWorkerThreadData;
            hr = prwtd ? S_OK : E_OUTOFMEMORY;
            if (SUCCEEDED(hr))
            {
                prwtd->uFirstIndex = uFirstIndex;
                prwtd->uLastIndex = uLastIndex;
                prwtd->dwManagerThreadId = GetCurrentThreadId();
                prwtd->hStartEvent = m_hStartEvent;
                prwtd->hCancelEvent = m_hCancelEvent;
                prwtd->sprm = this;
                m_workerThreadInfo[u].hWorker = CreateThread(nullptr, 0, s_rotationWorkerThread, prwtd, 0, &m_workerThreadInfo[u].dwThreadId);
                m_workerThreadHandles[u] = m_workerThreadInfo[u].hWorker;
                hr = (m_workerThreadInfo[u].hWorker) ? S_OK : E_FAIL;
                if (SUCCEEDED(hr))
                {
                    // increment the indices for the next thread
                    uFirstIndex = uLastIndex++;
                    uLastIndex = min((uFirstIndex + uItemsPerWorker), uTotalItems - 1);
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

DWORD WINAPI CRotationManager::s_rotationWorkerThread(__in void* pv)
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
                for (UINT u = prwtd->uFirstIndex; u <= prwtd->uLastIndex; u++)
                {
                    // Check if cancel event is signaled
                    if (WaitForSingleObject(prwtd->hCancelEvent, 0) == WAIT_OBJECT_0)
                    {
                        // Cancelled from manager
                        // Send the manager thread the canceled message
                        PostThreadMessage(prwtd->dwManagerThreadId, ROTM_ROTI_CANCELED, GetCurrentThreadId(), 0);
                        break;
                    }

                    CComPtr<IRotationItem> spri;
                    if (SUCCEEDED(prwtd->sprm->GetItem(u, &spri)))
                    {
                        // Perform the rotation
                        HRESULT hrWork = spri->Rotate();
                        spri->put_Result(hrWork);
                    }

                    // Send the manager thread the rotation item completed message
                    PostThreadMessage(prwtd->dwManagerThreadId, ROTM_ROTI_ROTATED, GetCurrentThreadId(), u);
                }
            }

            // Send the manager thread the completion message
            PostThreadMessage(prwtd->dwManagerThreadId, ROTM_ROTI_COMPLETE, GetCurrentThreadId(), 0);

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
    m_hStartEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);
    // Event used to signal worker thread in the event of a cancel
    m_hCancelEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);

    HRESULT hr = (m_hStartEvent && m_hCancelEvent) ? S_OK : E_FAIL;
    if (FAILED(hr))
    {
        _Cleanup();
    }

    return hr;
}

void CRotationManager::_ClearEventHandlers()
{
    AcquireSRWLockExclusive(&m_lockEvents);

    // Cleanup event handlers
    for (std::vector<ROTATION_MANAGER_EVENT>::iterator it = m_rotationManagerEvents.begin(); it != m_rotationManagerEvents.end(); ++it)
    {
        it->dwCookie = 0;
        if (it->prme)
        {
            it->prme->Release();
            it->prme = nullptr;
        }
    }

    m_rotationManagerEvents.clear();

    ReleaseSRWLockExclusive(&m_lockEvents);
}

void CRotationManager::_ClearRotationItems()
{
    AcquireSRWLockExclusive(&m_lockItems);

    // Cleanup rotation items
    for (std::vector<IRotationItem*>::iterator it = m_rotationItems.begin(); it != m_rotationItems.end(); ++it)
    {
        IRotationItem* pri = *it;
        pri->Release();
    }

    m_rotationItems.clear();

    ReleaseSRWLockExclusive(&m_lockItems);
}

void CRotationManager::_Cleanup()
{
    for (UINT u = 0; u < m_uWorkerThreadCount; u++)
    {
        CloseHandle(m_workerThreadInfo[u].hWorker);
        m_workerThreadInfo[u].hWorker = nullptr;
    }

   

    

    CloseHandle(m_hStartEvent);
    m_hStartEvent = nullptr;

    CloseHandle(m_hCancelEvent);
    m_hCancelEvent = nullptr;

    // Done with GDIPlus so shutdown now
    GdiplusShutdown(m_gdiplusToken);
}

UINT CRotationManager::s_GetLogicalProcessorCount()
{
    SYSTEM_INFO sysinfo;
    GetSystemInfo(&sysinfo);
    return sysinfo.dwNumberOfProcessors;
}
