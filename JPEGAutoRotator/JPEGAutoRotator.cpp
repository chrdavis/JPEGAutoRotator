// JPEGAutoRotator.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include "pch.h"

bool GetCurrentFolderPath(_In_ UINT count, _Out_ PWSTR path)
{
    bool ret = false;
    if (GetModuleFileName(nullptr, path, count))
    {
        PathRemoveFileSpec(path);
        ret = true;
    }
    return ret;
}

bool GetTestFolderPath(_In_ PCWSTR testFolderPathAppend, _In_ UINT testFolderPathLen, _Out_ PWSTR testFolderPath)
{
    bool ret = false;
    if (GetCurrentFolderPath(testFolderPathLen, testFolderPath))
    {
        if (SUCCEEDED(PathCchAppend(testFolderPath, testFolderPathLen, testFolderPathAppend)))
        {
            ret = true;
        }
    }

    return ret;
}

bool CopyHelper(_In_ PCWSTR src, _In_ PCWSTR dest)
{
    SHFILEOPSTRUCT fileStruct = { 0 };
    fileStruct.pFrom = src;
    fileStruct.pTo = dest;
    fileStruct.wFunc = FO_COPY;
    fileStruct.fFlags = FOF_SILENT | FOF_NOERRORUI | FOF_NO_UI | FOF_NOCONFIRMMKDIR | FOF_NOCONFIRMATION;
    return (SHFileOperation(&fileStruct) == 0);
}

bool DeleteHelper(_In_ PCWSTR path)
{
    SHFILEOPSTRUCT fileStruct = { 0 };
    fileStruct.pFrom = path;
    fileStruct.wFunc = FO_DELETE;
    fileStruct.fFlags = FOF_SILENT | FOF_NOERRORUI | FOF_NO_UI;
    return (SHFileOperation(&fileStruct) == 0);
}

HRESULT CreateDataObjectFromPath(_In_ PCWSTR pathToFile, _COM_Outptr_ IDataObject** dataObject)
{
    *dataObject = nullptr;

    CComPtr<IShellItem> spsi;
    SHCreateItemFromParsingName(pathToFile, nullptr /* IBindCtx */, IID_PPV_ARGS(&spsi));
    spsi->BindToHandler(nullptr /* IBindCtx */, BHID_DataObject, IID_PPV_ARGS(dataObject));

    return S_OK;
}

constexpr auto JPEGWITHEXIFROTATION_TESTFOLDER = L"TestFiles\\JPEGWithExifRotation\\";

PCWSTR g_rgTestJPEGFiles[] =
{
    L"Landscape_1.jpg",
    L"Landscape_2.jpg",
    L"Landscape_3.jpg",
    L"Landscape_4.jpg",
    L"Landscape_5.jpg",
    L"Landscape_6.jpg",
    L"Landscape_7.jpg",
    L"Landscape_8.jpg",
    L"Portrait_1.jpg",
    L"Portrait_2.jpg",
    L"Portrait_3.jpg",
    L"Portrait_4.jpg",
    L"Portrait_5.jpg",
    L"Portrait_6.jpg",
    L"Portrait_7.jpg",
    L"Portrait_8.jpg"
};

struct WorkerThreadTestInfo
{
    UINT uNumItems;
    UINT uNumItemsPerThread;
    UINT uNumWorkerThreads;
    ULONGLONG uEnumBegin;
    ULONGLONG uEnumEnd;
    ULONGLONG uEnumDuration;
    ULONGLONG uOperationBegin;
    ULONGLONG uOperationEnd;
    ULONGLONG uOperationDuration;
    ULONGLONG uTotalDuration;
};

struct WorkerThreadResults
{
    UINT uNumItems;
    UINT uNumItemsPerThread;
    UINT uNumWorkerThreads;
    ULONGLONG uEnumBegin;
    ULONGLONG uEnumEnd;
    ULONGLONG uEnumDuration;
    ULONGLONG uOperationBegin;
    ULONGLONG uOperationEnd;
    ULONGLONG uOperationDuration;
    ULONGLONG uTotalDuration;
};

WorkerThreadTestInfo g_workerThreadTestInfo[]
{
    // uNumItems      uNumItemsPerThread        uNumWorkerThreads
    {     1,                  1,                        1,      0,0,0,0,0,0,0},
    {     10,                 2,                        5,      0,0,0,0,0,0,0},
    {     10,                 5,                        2,      0,0,0,0,0,0,0},
    {     10,                 10,                       1,      0,0,0,0,0,0,0},
    {     50,                 5,                        10,     0,0,0,0,0,0,0},
    {     50,                 10,                       5,      0,0,0,0,0,0,0},
    {     50,                 50,                       1,      0,0,0,0,0,0,0},
    {     100,                10,                       10,     0,0,0,0,0,0,0},
    {     100,                20,                       5,      0,0,0,0,0,0,0},
    {     100,                30,                       4,      0,0,0,0,0,0,0},
    {     100,                50,                       2,      0,0,0,0,0,0,0},
    {     100,                100,                      1,      0,0,0,0,0,0,0},
    {     1000,               50,                       20,     0,0,0,0,0,0,0},
    {     1000,               100,                      10,     0,0,0,0,0,0,0},
    {     1000,               250,                      4,      0,0,0,0,0,0,0},
    {     1000,               300,                      4,      0,0,0,0,0,0,0},
    {     1000,               500,                      2,      0,0,0,0,0,0,0},
    {     1000,               1000,                     1,      0,0,0,0,0,0,0},
    {     10000,              1000,                     10,     0,0,0,0,0,0,0},
    {     10000,              2000,                     5,      0,0,0,0,0,0,0},
    {     10000,              5000,                     2,      0,0,0,0,0,0,0},
    {     10000,              10000,                    1,      0,0,0,0,0,0,0},
};

#define NUM_TEST_RUNS 10

HRESULT CopyTestFiles(PCWSTR pszSourceFolder, PCWSTR pszDestFolder, UINT uNumItems)
{
    HRESULT hr = S_OK;
    for (UINT i = 0; (i < uNumItems) && SUCCEEDED(hr); i++)
    {
        PCWSTR pszSrcFilename = g_rgTestJPEGFiles[i % ARRAYSIZE(g_rgTestJPEGFiles)];
        wchar_t szSourceFilePath[MAX_PATH];
        hr = PathCchCombine(szSourceFilePath, ARRAYSIZE(szSourceFilePath), pszSourceFolder, pszSrcFilename);
        if (SUCCEEDED(hr))
        {
            wchar_t szDestFilename[MAX_PATH];
            hr = StringCchPrintf(szDestFilename, ARRAYSIZE(szDestFilename), L"TestFile_%u.jpg", i);
            if (SUCCEEDED(hr))
            {
                wchar_t szDestFilePath[MAX_PATH];
                hr = PathCchCombine(szDestFilePath, ARRAYSIZE(szDestFilePath), pszDestFolder, szDestFilename);
                if (SUCCEEDED(hr))
                {
                    if (!CopyFile(szSourceFilePath, szDestFilePath, FALSE))
                    {
                        hr = HRESULT_FROM_WIN32(GetLastError());
                    }
                }
            }
        }
    }

    return hr;
}

HRESULT PerformOperation(__in PCWSTR pszTestFolder, __in WorkerThreadTestInfo* testData, __in WorkerThreadResults* results)
{
    results->uNumItems = testData->uNumItems;
    results->uNumItemsPerThread = testData->uNumItemsPerThread;
    results->uNumWorkerThreads = testData->uNumWorkerThreads;

    CComPtr<IDataObject> spdo;
    HRESULT hr = CreateDataObjectFromPath(pszTestFolder, &spdo);
    if (SUCCEEDED(hr))
    {
        CComPtr<IRotationManager> sprm;
        hr = CRotationManager::s_CreateInstance(&sprm);
        if (SUCCEEDED(hr))
        {
            CComPtr<IRotationManagerDiagnostics> sprmd;
            hr = sprm->QueryInterface(IID_PPV_ARGS(&sprmd));
            if (SUCCEEDED(hr))
            {
                sprmd->put_ItemsPerWorkerThread(testData->uNumItemsPerThread);
                sprmd->put_WorkerThreadCount(testData->uNumWorkerThreads);

                results->uEnumBegin = GetTickCount64();
                hr = EnumerateDataObject(spdo, sprm);
                results->uEnumEnd = GetTickCount64();
                results->uEnumDuration = results->uEnumEnd - results->uEnumBegin;

                if (SUCCEEDED(hr))
                {
                    results->uOperationBegin = GetTickCount64();
                    sprm->Start();
                    results->uOperationEnd = GetTickCount64();
                    results->uOperationDuration = results->uOperationEnd - results->uOperationBegin;
                    results->uTotalDuration = results->uOperationDuration + results->uEnumDuration;
                }
            }
        }
    }
 
    return hr;
}

int main()
{
    ULONGLONG ullBegin = GetTickCount64();
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (SUCCEEDED(hr))
    {
        // Get test file source folder
        wchar_t testFolderSource[MAX_PATH] = { 0 };
        GetTestFolderPath(JPEGWITHEXIFROTATION_TESTFOLDER, ARRAYSIZE(testFolderSource), testFolderSource);

        // Get test file working folder
        wchar_t testFolderWorking[MAX_PATH] = { 0 };
        GetTestFolderPath(L"TestWorkThreadAllocations", ARRAYSIZE(testFolderWorking), testFolderWorking);

        UINT numTestResultEntries = ARRAYSIZE(g_workerThreadTestInfo) * NUM_TEST_RUNS;
        WorkerThreadResults* workerThreadTestResults = new WorkerThreadResults[numTestResultEntries];
        ZeroMemory(workerThreadTestResults, numTestResultEntries * sizeof(WorkerThreadResults));

        UINT testRun = 0;
        while (testRun < NUM_TEST_RUNS)
        {
            for (UINT i = 0; i < ARRAYSIZE(g_workerThreadTestInfo); i++)
            {
                // Ensure test folder doesn't exist
                DeleteHelper(testFolderWorking);

                CreateDirectory(testFolderWorking, nullptr);

                CopyTestFiles(testFolderSource, testFolderWorking, g_workerThreadTestInfo[i].uNumItems);

                PerformOperation(testFolderWorking, &g_workerThreadTestInfo[i], &workerThreadTestResults[(testRun * ARRAYSIZE(g_workerThreadTestInfo)) + i]);

                // Cleanup working folder
                DeleteHelper(testFolderWorking);
            }

            testRun++;
        }

        // Print test results
        for (UINT i = 0; i < numTestResultEntries; i++)
        {
            std::cout << workerThreadTestResults[i].uNumItems << " ";
            std::cout << workerThreadTestResults[i].uNumItemsPerThread << " ";
            std::cout << workerThreadTestResults[i].uNumWorkerThreads << " ";
            std::cout << workerThreadTestResults[i].uEnumDuration << " ";
            std::cout << workerThreadTestResults[i].uOperationDuration << " ";
            std::cout << workerThreadTestResults[i].uTotalDuration << "\n";
        }

        ULONGLONG ullEnd = GetTickCount64();

        std::cout << "Total Run Time: " << ullEnd - ullBegin << "\n";
        CoUninitialize();
    }
}

