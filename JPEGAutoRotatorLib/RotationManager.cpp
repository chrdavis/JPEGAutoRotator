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
    ROTM_ROTI_COMPLETE,                // Worker thread completed
    ROTM_ENDTHREAD                     // End worker threads and exit manager
};

int GetEncoderClsid(const WCHAR* format, CLSID* pClsid)
{
    UINT  num = 0;          // number of image encoders
    UINT  size = 0;         // size of the image encoder array in bytes

    ImageCodecInfo* pImageCodecInfo = NULL;

    GetImageEncodersSize(&num, &size);
    if (size == 0)
        return -1;  // Failure

    pImageCodecInfo = (ImageCodecInfo*)(malloc(size));
    if (pImageCodecInfo == NULL)
        return -1;  // Failure

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
    IObjectCollection* poc;

    ~RotateWorkerThreadData()
    {
        if (poc)
        {
            poc->Release();
        }
    }
};

CRotationItem::CRotationItem() :
    m_cRef(1),
    m_pszPath(nullptr),
    m_hrResult(S_FALSE)
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
        hr = pri->SetPath(pszPath);
        if (SUCCEEDED(hr))
        {
            hr = pri->QueryInterface(IID_PPV_ARGS(ppri));
        }
        pri->Release();
    }
    return hr;
}

IFACEMETHODIMP CRotationItem::GetPath(__deref_out PWSTR* ppszPath)
{
    *ppszPath = nullptr;
    HRESULT hr = m_pszPath ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        hr = SHStrDup(m_pszPath, ppszPath);
    }
    return hr;
}

IFACEMETHODIMP CRotationItem::SetPath(__in PCWSTR pszPath)
{
    HRESULT hr = pszPath ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        CoTaskMemFree(m_pszPath);
        hr = SHStrDup(pszPath, &m_pszPath);
    }
    return hr;
}

IFACEMETHODIMP CRotationItem::GetResult(__out HRESULT* phrResult)
{
    *phrResult = m_hrResult;
    return S_OK; 
}

IFACEMETHODIMP CRotationItem::SetResult(__in HRESULT hrResult)
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
                            USHORT uOrientation = *((USHORT*)pItem->value);
                            // Ensure valid orientation range.  If not in range do nothing and return success.
                            if (uOrientation <= 8 && uOrientation >= 2)
                            {
                                hr = (pImage->RotateFlip(rotateFlipTable[uOrientation]) == Ok) ? S_OK : E_FAIL;
                                if (SUCCEEDED(hr))
                                {
                                    // Remove the orientation tag
                                    pImage->RemovePropertyItem(PropertyTagOrientation);
                                    // Save back to original location
                                    IStream_Reset(spstrm);
                                    CLSID jpgClsid;
                                    GetEncoderClsid(L"image/jpeg", &jpgClsid);
                                    hr = (pImage->Save(spstrm, &jpgClsid, nullptr) == Ok) ? S_OK : E_FAIL;
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

CRotationManager::CRotationManager() :
    m_cRef(1),
    m_dwCookie(0),
    m_uWorkerThreadCount(0),
    m_hStartEvent(nullptr),
    m_hCancelEvent(nullptr)
{
    DllAddRef();
}

CRotationManager::~CRotationManager()
{
    _Cleanup();
    DllRelease();
}

IFACEMETHODIMP CRotationManager::Advise(__in IRotationManagerEvents* prme, __out DWORD* pdwCookie)
{
    HRESULT hr = prme ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        // TODO: allow > 1 IRotationManagerEvents
        m_sprme = prme;
        m_dwCookie++;
        *pdwCookie = m_dwCookie;
    }

    return hr;
}

IFACEMETHODIMP CRotationManager::UnAdvise(__in DWORD dwCookie)
{
    if (dwCookie == m_dwCookie)
    {
        m_sprme = nullptr;
    }

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
    HRESULT hr = (pri && m_spoc) ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        hr = m_spoc->AddObject(pri);
    }
    return hr;
}

IFACEMETHODIMP CRotationManager::GetItem(__in UINT uIndex, __deref_out IRotationItem** ppri)
{
    *ppri = nullptr;
    HRESULT hr = m_spoc ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        hr = m_spoc->GetAt(uIndex, IID_PPV_ARGS(ppri));
    }
    return hr;
}

IFACEMETHODIMP CRotationManager::GetItemCount(__out UINT* puCount)
{
    *puCount = 0;
    HRESULT hr = m_spoc ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        hr = m_spoc->GetCount(puCount);
    }
    return hr;
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
        UINT uTotalItems = 0;
        m_spoc->GetCount(&uTotalItems);

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
            while (PeekMessage(&msg, nullptr, 0, 0, PM_REMOVE))
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
                    if (m_sprme)
                    {
                        m_sprme->OnRotated(msg.lParam);
                        m_sprme->OnProgress(uCompleted, uTotalItems);
                    }
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

            if (m_sprme)
            {
                m_sprme->OnCompleted();
            }
        }
    }

    return 0;
}

HRESULT CRotationManager::_CreateWorkerThreads()
{
    UINT uMaxWorkerThreads = min(s_GetLogicalProcessorCount(), MAX_ROTATION_WORKER_THREADS);
    UINT uTotalItems = 0;
    HRESULT hr = m_spoc->GetCount(&uTotalItems);
    if (SUCCEEDED(hr))
    {
        hr = (uTotalItems > 0) ? S_OK : E_FAIL;
    }

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
                prwtd->poc = m_spoc;
                prwtd->poc->AddRef();
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
                        break;
                    }

                    CComPtr<IRotationItem> spri;
                    HRESULT hrWork = prwtd->poc->GetAt(u, IID_PPV_ARGS(&spri));
                    if (SUCCEEDED(hrWork))
                    {
                        // Perform the rotation
                        hrWork = spri->Rotate();

                        spri->SetResult(hrWork);
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

    HRESULT hr = CoCreateInstance(CLSID_EnumerableObjectCollection,
        nullptr,
        CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(&m_spoc));
    if (SUCCEEDED(hr))
    {
        // Event used to signal worker thread that it can start
        m_hStartEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);
        // Event used to signal worker thread in the event of a cancel
        m_hCancelEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);
    
        hr = (m_hStartEvent && m_hCancelEvent) ? S_OK : E_FAIL;
        if (FAILED(hr))
        {
            _Cleanup();
        }
    }

    return hr;
}

void CRotationManager::_Cleanup()
{
    // Done with GDIPlus so shutdown now
    GdiplusShutdown(m_gdiplusToken);

    for (UINT u = 0; u < m_uWorkerThreadCount; u++)
    {
        CloseHandle(m_workerThreadInfo[u].hWorker);
        m_workerThreadInfo[u].hWorker = nullptr;
    }

    m_sprme = nullptr;

    CloseHandle(m_hStartEvent);
    m_hStartEvent = nullptr;

    CloseHandle(m_hCancelEvent);
    m_hCancelEvent = nullptr;
}

UINT CRotationManager::s_GetLogicalProcessorCount()
{
    SYSTEM_INFO sysinfo;
    GetSystemInfo(&sysinfo);
    return sysinfo.dwNumberOfProcessors;
}
