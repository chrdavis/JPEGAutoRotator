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
    ROTM_ROTI_UPDATE = (WM_APP + 1),   // Single rotation item finished
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

CRotationItem::CRotationItem() :
    m_cRef(1)
{
}

CRotationItem::~CRotationItem()
{
}

HRESULT CRotationItem::GetItem(__deref_out IShellItem** ppsi)
{
    *ppsi = nullptr;
    HRESULT hr = E_FAIL;
    if (m_spsi)
    {
        *ppsi = m_spsi;
        (*ppsi)->AddRef();
    }
    return hr;
}

HRESULT CRotationItem::SetItem(__in IShellItem* psi)
{
    HRESULT hr = psi ? S_OK : E_INVALIDARG;
    if (SUCCEEDED(hr))
    {
        m_spsi = psi;
    }
    return hr;
}

HRESULT CRotationItem::s_CreateInstance(__in IShellItem* psi, __out CRotationItem** ppri)
{
    *ppri = nullptr;
    CRotationItem* pri = new CRotationItem();
    HRESULT hr = pri ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = pri->SetItem(psi);
        if (SUCCEEDED(hr))
        {
            *ppri = pri;
            (*ppri)->AddRef();
        }
        pri->Release();
    }
    return hr;
}

HRESULT CRotationItem::Rotate()
{
    HRESULT hr = m_spsi ? S_OK : E_FAIL;
    if (SUCCEEDED(hr))
    {
        // Open the stream read write
        CComPtr<IBindCtx> spbc;
        hr = CreateBindCtx(0, &spbc);
        if (SUCCEEDED(hr))
        {
            BIND_OPTS bo = { sizeof(bo), 0, STGM_READWRITE, 0 };
            hr = spbc->SetBindOptions(&bo);
            if (SUCCEEDED(hr))
            {
                CComPtr<IStream> spstrm;
                hr = m_spsi->BindToHandler(spbc, BHID_Stream, IID_PPV_ARGS(&spstrm));
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
                            if (pImage->GetLastStatus() ==Ok)
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
        }
    }
    return hr;
}

CRotationManager::CRotationManager() :
    m_cRef(1),
    m_spdo(nullptr),
    m_sppd(nullptr),
    m_uWorkerThreadCount(0)
{
    DllAddRef();
}

CRotationManager::~CRotationManager()
{
    _Cleanup();
    DllRelease();
}

HRESULT CRotationManager::s_CreateInstance(__in IDataObject* pdo, __deref_out CRotationManager** pprm)
{
    *pprm = nullptr;
    CRotationManager *prm = new CRotationManager();
    HRESULT hr = prm ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = prm->_Init(pdo);
        if (SUCCEEDED(hr))
        {
            *pprm = prm;
            (*pprm)->AddRef();
        }
        prm->Release();
    }
    return hr;
}

HRESULT CRotationManager::PerformRotation()
{
    // Start progress dialog
    m_sppd->StartProgressDialog(HWND_DESKTOP, nullptr, PROGDLG_NORMAL | PROGDLG_AUTOTIME, nullptr);

    // Enumerate items
    HRESULT hr = _EnumerateDataObject();
    if (SUCCEEDED(hr))
    {
        // Create worker threads which will message us progress and
        // completion.

        MSG msg;
        while (GetMessage(&msg, nullptr, 0, 0))
        {
            // If we got the "operation complete" message
            if (msg.message == ROTM_ENDTHREAD)
            {
                // Break out of the loop and end the thread
                break;
            }
            else if (msg.message == ROTM_ROTI_UPDATE)
            {
                _UpdateProgressForWorkerThread((UINT)msg.wParam, (UINT)msg.lParam);
            }
            else if (msg.message == ROTM_ROTI_COMPLETE)
            {
                // Worker thread completed
                // Break out of the loop and end the thread
                break;
            }
            else
            {
                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }
        }
    }

    return 0;
}

HRESULT CRotationManager::_CreateWorkerThreads()
{
    UINT uMaxWorkerThreads = min(s_GetLogicalProcessorCount(), MAX_ROTATION_WORKER_THREADS);
    UINT uTotalItems = 0;
    m_spoc->GetCount(&uTotalItems);

    // Determine the best number of threads based on the number of items to rotate
    // and our minimum number of items per worker thread

    // How many work item sized amounts of items do we have?
    UINT uTotalWorkItems = (uTotalItems / MIN_ROTATION_WORK_SIZE);

    // How many worker threads do we require, with a minimum being the max worker threads value
    m_uWorkerThreadCount = min((uTotalWorkItems / uMaxWorkerThreads), uMaxWorkerThreads);
    if (m_uWorkerThreadCount == 0)
    {
        m_uWorkerThreadCount = (uTotalWorkItems % uMaxWorkerThreads);
    }

    // Create the worker threads


    return S_OK;
}

void CRotationManager::_UpdateProgressForWorkerThread(UINT uThreadId, UINT uCompleted)
{
    if (m_sppd && m_spoc)
    {
        // Get total number of rotation items
        UINT uTotal = 0;
        m_spoc->GetCount(&uTotal);

        // Get count
        m_sppd->SetProgress(uCompleted, uTotal);
    }
}

HRESULT CRotationManager::_Init(__in IDataObject* pdo)
{
    // Initialize GDIPlus now. All of our GDIPlus usage is in CRenameItem.
    GdiplusStartupInput gdiplusStartupInput;
    GdiplusStartup(&m_gdiplusToken, &gdiplusStartupInput, nullptr);

    // Cache the data object
    m_spdo = pdo;

    // Create the progress dialog we will show during the operation
    return CoCreateInstance(CLSID_ProgressDialog,
        nullptr,
        CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(&m_sppd));
}

HRESULT CRotationManager::_Cleanup()
{
    // Done with GDIPlus so shutdown now
    GdiplusShutdown(m_gdiplusToken);

    m_spdo = nullptr;
    m_sppd = nullptr;

    return S_OK;
}

// Iterate through the data object and add items to the IObjectCollection
HRESULT CRotationManager::_EnumerateDataObject()
{
    // Clear any existing collection
    m_spoc = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_EnumerableObjectCollection,
        nullptr,
        CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(&m_spoc));
    if (SUCCEEDED(hr))
    {
        CComPtr<IShellItemArray> spsia;
        hr = SHCreateShellItemArrayFromDataObject(m_spdo, IID_PPV_ARGS(&spsia));
        if (SUCCEEDED(hr))
        {
            CComPtr<IEnumShellItems> spesi;
            hr = spsia->EnumItems(&spesi);
            if (SUCCEEDED(hr))
            {
                ULONG celtFetched = 0;
                IShellItem *psi = nullptr;
                while ((S_OK == spesi->Next(1, &psi, &celtFetched)) && (SUCCEEDED(hr)))
                {
                    SFGAOF att = 0;
                    if (SUCCEEDED(psi->GetAttributes(SFGAO_FOLDER, &att))) // pItem is a IShellItem*
                    {
                        // TODO: should we do this here or later?
                        // Don't bother including folders
                        if (!(att & SFGAO_FOLDER))
                        {
                            CRotationItem* priNew;
                            hr = CRotationItem::s_CreateInstance(psi, &priNew);
                            if (SUCCEEDED(hr))
                            {
                                hr = m_spoc->AddObject(priNew);
                                priNew->Release();
                            }
                        }
                    }
                    psi->Release();
                }
            }
        }
    }

    return hr;
}

UINT CRotationManager::s_GetLogicalProcessorCount()
{
    SYSTEM_INFO sysinfo;    GetSystemInfo(&sysinfo);    return sysinfo.dwNumberOfProcessors;
}
