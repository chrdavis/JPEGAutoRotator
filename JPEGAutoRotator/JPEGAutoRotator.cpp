#include "pch.h"

// Maps EXIF orientation property values to strings
// representing the rotation to perform to normalize
PCWSTR g_rotationStrings[] =
{
    L"None",
    L"None",
    L"RotateNoneFlipX",
    L"Rotate180FlipNone",
    L"RotateNoneFlipY",
    L"Rotate270FlipY",
    L"Rotate90FlipNone",
    L"Rotate90FlipY",
    L"Rotate270FlipNone"
};

class CJPEGAutoRotator :
    public IRotationUI,
    public IRotationManagerEvents
{
public:
    CJPEGAutoRotator() {}

    // IUnknown
    IFACEMETHODIMP QueryInterface(_In_ REFIID riid, _COM_Outptr_ void** ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CJPEGAutoRotator, IRotationUI),
            QITABENT(CJPEGAutoRotator, IRotationManagerEvents),
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
    IFACEMETHODIMP Initialize(_In_ IDataObject*)
    {
        return S_OK;
    }

    IFACEMETHODIMP Start()
    {
        ULONGLONG startTime = GetTickCount64();

        HRESULT hr = _Initialize();
        if (SUCCEEDED(hr))
        {
            hr = m_sprm->Start();
        }

        m_duration = GetTickCount64() - startTime;

        if (m_previewMode || m_printStats)
        {
            m_previewMode ? _PrintPreview() : _PrintRotations();
            wprintf(L"\n");
            _PrintThreadDetails();
            wprintf(L"\n");
            _PrintDuration();
        }

        _Cleanup();
        return hr;
    }

    IFACEMETHODIMP Close()
    {
        return S_OK;
    }

    // IRotationManagerEvents
    IFACEMETHODIMP OnItemAdded(_In_ UINT)
    {
        return S_OK;
    }

    IFACEMETHODIMP OnItemProcessed(_In_ UINT index)
    {
        if (!m_silentMode && m_verbose && !m_showProgress)
        {
            CComPtr<IRotationItem> spItem;
            if (SUCCEEDED(m_sprm->GetItem(index, &spItem)))
            {
                PWSTR path = nullptr;
                if (SUCCEEDED(spItem->get_Path(&path)))
                {
                    BOOL wasRotated = FALSE;
                    spItem->get_WasRotated(&wasRotated);

                    wasRotated ? wprintf(L"Rotated: ") : wprintf(L"Did not rotate: ");
                    wprintf(path);
                    wprintf(L"\n");
                    CoTaskMemFree(path);
                }
            }
        }
        return S_OK;
    }

    IFACEMETHODIMP OnProgress(_In_ UINT completed, _In_ UINT total)
    {
        if (!m_silentMode && m_showProgress)
        {
            wchar_t buff[100];
            if (SUCCEEDED(StringCchPrintf(buff, ARRAYSIZE(buff), L"Processed %u of %u items...\r", completed, total)))
            {
                wprintf(buff);
                fflush(stdout);
            }

        }
        return S_OK;
    }

    IFACEMETHODIMP OnCanceled()
    {
        if (!m_silentMode && !m_showProgress)
        {
            wprintf(L"Operation canceled.\n");
        }
        return S_OK;
    }

    IFACEMETHODIMP OnCompleted()
    {
        if (!m_silentMode && !m_showProgress)
        {
            wprintf(L"Completed.\n");
        }
        return S_OK;
    }

private:
    ~CJPEGAutoRotator()
    {
        CoTaskMemFree(m_path);
    }

    HRESULT _Initialize()
    {
        // Parse the command line
        HRESULT hr = _ParseCommandLine() ? S_OK : E_FAIL;
        if (SUCCEEDED(hr))
        {
            // Create rotation manager && subscribe to events
            hr = CRotationManager::s_CreateInstance(&m_sprm);
            if (SUCCEEDED(hr))
            {
                hr = m_sprm->Advise(this, &m_adviseCookie);
            }
        }

        if (SUCCEEDED(hr))
        {
            // Apply any command line options to the rotation manager
            CComPtr<IRotationManagerDiagnostics> sprmd;
            hr = m_sprm->QueryInterface(IID_PPV_ARGS(&sprmd));
            if (SUCCEEDED(hr))
            {
                sprmd->put_EnumerateSubFolders(m_enumSubFolders);
                sprmd->put_LosslessOnly(m_losslessOnly);
                sprmd->put_PreviewOnly(m_previewMode);
                if (m_threadCount > 0)
                {
                    sprmd->put_WorkerThreadCount(m_threadCount);
                }

                if (m_minItemsPerThread > 0)
                {
                    sprmd->put_MinItemsPerWorkerThread(m_minItemsPerThread);
                }
            }
        }

        if (SUCCEEDED(hr))
        {
            // AddPath will enumerate all files and folders if EnumerateSubFolders is TRUE (default)
            hr = m_sprm->AddPath(m_path);
        }

        return hr;
    }

    void _Cleanup()
    {
        if (m_sprm)
        {
            m_sprm->Shutdown();
            m_sprm = nullptr;
        }
    }

    bool _ParseCommandLine()
    {
        bool printUsage = false;
        int cArguments;
        PWSTR *ppArguments = CommandLineToArgvW(GetCommandLine(), &cArguments);
        if (ppArguments)
        {
            printUsage = (cArguments < 2);
            if (!printUsage)
            {
                int i = 0;
                while (i < cArguments)
                {
                    if ((_wcsicmp(ppArguments[i], L"/?") == 0) ||
                        (_wcsicmp(ppArguments[i], L"?") == 0) ||
                        (_wcsicmp(ppArguments[i], L"-?") == 0) ||
                        (_wcsicmp(ppArguments[i], L"-Help") == 0) ||
                        (_wcsicmp(ppArguments[i], L"/Help") == 0))
                    {
                        printUsage = true;
                        break;
                    }
                    else if ((_wcsicmp(ppArguments[i], L"/Silent") == 0) ||
                        (_wcsicmp(ppArguments[i], L"-Silent") == 0))
                    {
                        m_silentMode = true;
                    }
                    else if ((_wcsicmp(ppArguments[i], L"/NoSubFolders") == 0) ||
                        (_wcsicmp(ppArguments[i], L"-NoSubFolders") == 0))
                    {
                        m_enumSubFolders = false;
                    }
                    else if ((_wcsicmp(ppArguments[i], L"/Progress") == 0) ||
                        (_wcsicmp(ppArguments[i], L"-Progress") == 0))
                    {
                        m_showProgress = true;
                    }
                    else if ((_wcsicmp(ppArguments[i], L"/LosslessOnly") == 0) ||
                        (_wcsicmp(ppArguments[i], L"-LosslessOnly") == 0))
                    {
                        m_losslessOnly = true;
                    }
                    else if ((_wcsicmp(ppArguments[i], L"/Verbose") == 0) ||
                        (_wcsicmp(ppArguments[i], L"-Verbose") == 0))
                    {
                        m_verbose = true;
                    }
                    else if ((_wcsicmp(ppArguments[i], L"/Preview") == 0) ||
                        (_wcsicmp(ppArguments[i], L"-Preview") == 0))
                    {
                        m_previewMode = true;
                    }
                    else if ((_wcsicmp(ppArguments[i], L"/Stats") == 0) ||
                        (_wcsicmp(ppArguments[i], L"-Stats") == 0))
                    {
                        m_printStats = true;
                    }
                    else if ((_wcsicmp(ppArguments[i], L"/MinItemsPerThread") == 0) ||
                        (_wcsicmp(ppArguments[i], L"-MinItemsPerThread") == 0))
                    {
                        printUsage = true;
                        if (i + 1 < cArguments)
                        {
                            m_minItemsPerThread = _wtoi(ppArguments[i + 1]);
                            if (m_minItemsPerThread > 0 && m_minItemsPerThread != ERANGE)
                            {
                                printUsage = false;
                            }
                        }

                        if (printUsage)
                        {
                            break;
                        }
                    }
                    else if ((_wcsicmp(ppArguments[i], L"/ThreadCount") == 0) ||
                        (_wcsicmp(ppArguments[i], L"-ThreadCount") == 0))
                    {
                        printUsage = true;
                        if (i + 1 < cArguments)
                        {
                            m_threadCount = _wtoi(ppArguments[i + 1]);
                            if (m_threadCount > 0 && m_threadCount != ERANGE)
                            {
                                printUsage = false;
                            }
                        }

                        if (printUsage)
                        {
                            break;
                        }
                    }
                    else if (i == (cArguments - 1))
                    {
                        // Last argument is assumed to be the path
                        SHStrDup(ppArguments[i], &m_path);
                    }
                    i++;
                }
            }

            if (printUsage)
            {
                _PrintUsage();
            }

            LocalFree(ppArguments);
        }

        return !printUsage;
    }

    void _PrintRotations()
    {
        UINT itemCount = 0;
        UINT validJPEGCount = 0;
        UINT rotatedCount = 0;
        if (m_sprm)
        {
            if (SUCCEEDED(m_sprm->GetItemCount(&itemCount)))
            {
                for (UINT i = 0; i < itemCount; i++)
                {
                    CComPtr<IRotationItem> spItem;
                    if (SUCCEEDED(m_sprm->GetItem(i, &spItem)))
                    {
                        BOOL wasRotated = FALSE;
                        BOOL isValidJPEG = FALSE;
                        if (SUCCEEDED(spItem->get_IsValidJPEG(&isValidJPEG)) && isValidJPEG)
                        {
                            validJPEGCount++;
                        }

                        if (SUCCEEDED(spItem->get_WasRotated(&wasRotated)) && wasRotated)
                        {
                            rotatedCount++;
                        }
                    }
                }
            }
        }

        wprintf(L"\n");
        wchar_t buff[100];
        StringCchPrintf(buff, ARRAYSIZE(buff), L"Number of items processed:\t%u\n", itemCount);
        wprintf(buff);
        StringCchPrintf(buff, ARRAYSIZE(buff), L"Number of valid JPEGs:\t\t%u\n", validJPEGCount);
        wprintf(buff);
        StringCchPrintf(buff, ARRAYSIZE(buff), L"Number of JPEGs rotated:\t%u\n", rotatedCount);
        wprintf(buff);
    }

    void _PrintPreview()
    {
        UINT itemCount = 0;
        UINT validJPEGCount = 0;
        UINT willRotateCount = 0;
        if (m_sprm)
        {
            wprintf(L"Path,IsValidJPEG,IsLossless,Orientation,WillRotate\n");
            if (SUCCEEDED(m_sprm->GetItemCount(&itemCount)))
            {
                for (UINT i = 0; i < itemCount; i++)
                {
                    CComPtr<IRotationItem> spItem;
                    if (SUCCEEDED(m_sprm->GetItem(i, &spItem)))
                    {
                        PWSTR path = nullptr;
                        BOOL isValidJPEG = FALSE;
                        UINT orientation = 0;
                        BOOL isLossless = FALSE;
                        spItem->get_Path(&path);
                        spItem->get_IsRotationLossless(&isLossless);
                        spItem->get_IsValidJPEG(&isValidJPEG);
                        spItem->get_OriginalOrientation(&orientation);

                        bool willRotate = false;
                        if (isValidJPEG)
                        {
                            validJPEGCount++;
                            if (isLossless || !m_losslessOnly)
                            {
                                if (orientation <= 8 && orientation >= 2)
                                {
                                    willRotateCount++;
                                    willRotate = true;
                                }
                            }
                        }

                        wchar_t detailsBuff[MAX_PATH * 2] = { 0 };
                        StringCchPrintf(detailsBuff, ARRAYSIZE(detailsBuff), L"%s,%s,%s,%u,%s\n", path, isValidJPEG ? L"true" : L"false", isLossless ? L"true" : L"false", orientation, willRotate ? L"true" : L"false");
                        wprintf(detailsBuff);
                        CoTaskMemFree(path);
                    }
                }
            }
        }

        wprintf(L"\n");
        wchar_t buff[100];
        StringCchPrintf(buff, ARRAYSIZE(buff), L"Number of items processed:\t%u\n", itemCount);
        wprintf(buff);
        StringCchPrintf(buff, ARRAYSIZE(buff), L"Number of valid JPEGs:\t\t%u\n", validJPEGCount);
        wprintf(buff);
        StringCchPrintf(buff, ARRAYSIZE(buff), L"Number of JPEGs to be rotated:\t%u\n", willRotateCount);
        wprintf(buff);
    }

    void _PrintThreadDetails()
    {
        if (m_sprm)
        {
            CComPtr<IRotationManagerDiagnostics> sprmd;
            if (SUCCEEDED(m_sprm->QueryInterface(IID_PPV_ARGS(&sprmd))))
            {
                UINT maxWorkerThreadCount = 0;
                UINT workerThreadCount = 0;
                UINT itemsPerThread = 0;
                UINT minItemsPerThread = 0;
                sprmd->get_MaxWorkerThreadCount(&maxWorkerThreadCount);
                sprmd->get_WorkerThreadCount(&workerThreadCount);
                sprmd->get_MinItemsPerWorkerThread(&minItemsPerThread);
                sprmd->get_ItemsPerWorkerThread(&itemsPerThread);
                wchar_t buff[100];
                StringCchPrintf(buff, ARRAYSIZE(buff), L"Max Worker Threads:\t%u\n", maxWorkerThreadCount);
                wprintf(buff);
                StringCchPrintf(buff, ARRAYSIZE(buff), L"Worker Threads:\t\t%u\n", workerThreadCount);
                wprintf(buff);
                StringCchPrintf(buff, ARRAYSIZE(buff), L"Min Items Per Thread:\t%u\n", minItemsPerThread);
                wprintf(buff);
                StringCchPrintf(buff, ARRAYSIZE(buff), L"Items Per Thread:\t%u\n", itemsPerThread);
                wprintf(buff);
            }
        }
    }

    void _PrintDuration()
    {
        wchar_t buff[100];
        StringCchPrintf(buff, ARRAYSIZE(buff), L"Operation took: %u ms\n", m_duration);
        wprintf(buff);
    }

    void _PrintUsage()
    {
        wprintf(L"Normalize JPEG images with EXIF orientation properties. Images are rotated based on the EXIF data\n");
        wprintf(L"and the orientation property is then removed.\n");
        wprintf(L"\n");
        wprintf(L"USAGE:\n");
        wprintf(L"  JPEGAutoRotator [Options] Path\n");
        wprintf(L"\n");
        wprintf(L"Options:\n");
        wprintf(L"  -?\t\t\tDisplays this help message\n");
        wprintf(L"  -Silent\t\tNo output to the console\n");
        wprintf(L"  -Verbose\t\tDetailed output of operation and results to the console\n");
        wprintf(L"  -NoSubFolders\t\tDo not enumerate items in sub folders of path\n");
        wprintf(L"  -LosslessOnly\t\tDo not rotate images which would lose data as a result.\n");
        wprintf(L"  -ThreadCount n\tNumber of threads used for the rotations. Default is the number of cores on the device.\n");
        wprintf(L"  -MinItemsPerThread n\tThe minimum number of items per worker thread.\n");
        wprintf(L"  -Stats\t\tPrint detailed information of the operation that was performed to the console.\n");
        wprintf(L"  -Progress\t\tPrint progress of operation to the console.\n");
        wprintf(L"  -Preview\t\tProvides a printed preview in csv format of what will be performed without modifying any images.\n");
        wprintf(L"\n");
        wprintf(L"  Path\t\t\tFull path to a file or folder to auto-rotate.  Wildcards allowed.\n");
        wprintf(L"\n");
        wprintf(L"Examples:\n");
        wprintf(L"  > JPEGAutoRotator c:\\users\\chris\\pictures\\mypicturetorotate.jpg\n");
        wprintf(L"  > JPEGAutoRotator c:\\foo\\\n");
        wprintf(L"  > JPEGAutoRotator c:\\bar\\*.jpg\n");
        wprintf(L"  > JPEGAutoRotator -Progress -Stats c:\\test\\\n");
        wprintf(L"  > JPEGAutoRotator -Preview -LosslessOnly -NoSubFolders c:\\test\\\n");
    }

    long m_refCount = 1;
    bool m_silentMode = false;
    bool m_losslessOnly = false;
    bool m_enumSubFolders = true;
    bool m_showProgress = false;
    bool m_verbose = false;
    bool m_previewMode = false;
    bool m_printStats = false;
    int m_minItemsPerThread = -1;
    int m_threadCount = -1;
    ULONGLONG m_duration = 0;
    PWSTR m_path = nullptr;
    DWORD m_adviseCookie = 0;
    CComPtr<IRotationManager> m_sprm;
};

int main()
{
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (SUCCEEDED(hr))
    {
        CJPEGAutoRotator *pRotator = new CJPEGAutoRotator();
        hr = pRotator->Start();
        pRotator->Release();
        CoUninitialize();
    }
}

