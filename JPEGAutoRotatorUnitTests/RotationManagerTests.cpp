#include "stdafx.h"
#include <RotationInterfaces.h>
#include <RotationManager.h>
#include "MockRotationItem.h"
#include "MockRotationManagerEvents.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

constexpr auto JPEGWITHEXIFROTATION_TESTFOLDER = L"TestFiles\\JPEGWithExifRotation\\";

extern HINSTANCE g_hInst;

namespace JPEGAutoRotatorUnitTests
{
    struct TestFileData
    {
        PCWSTR filename;
        UINT originalOrientation;
        BOOL isValidJPEG;
        BOOL isLossless;
        BOOL shouldRotate;
    };

    // Test file names. These have all the possible exif rotation values
    TestFileData g_rgTestFileData[] =
    {
        { L"Landscape_1.jpg", 1, TRUE, FALSE, FALSE },
        { L"Landscape_2.jpg", 2, TRUE, FALSE, TRUE },
        { L"Landscape_3.jpg", 3, TRUE, FALSE, TRUE },
        { L"Landscape_4.jpg", 4, TRUE, FALSE, TRUE },
        { L"Landscape_5.jpg", 5, TRUE, FALSE, TRUE },
        { L"Landscape_6.jpg", 6, TRUE, FALSE, TRUE },
        { L"Landscape_7.jpg", 7, TRUE, FALSE, TRUE },
        { L"Landscape_8.jpg", 8, TRUE, FALSE, TRUE },
        { L"Portrait_1.jpg", 1, TRUE, FALSE, FALSE },
        { L"Portrait_2.jpg", 2, TRUE, FALSE, TRUE },
        { L"Portrait_3.jpg", 3, TRUE, FALSE, TRUE },
        { L"Portrait_4.jpg", 4, TRUE, FALSE, TRUE },
        { L"Portrait_5.jpg", 5, TRUE, FALSE, TRUE },
        { L"Portrait_6.jpg", 6, TRUE, FALSE, TRUE },
        { L"Portrait_7.jpg", 7, TRUE, FALSE, TRUE },
        { L"Portrait_8.jpg", 8, TRUE, FALSE, TRUE },
    };

    TEST_CLASS(RotationManagerTests)
    {
    public:
        TEST_CLASS_INITIALIZE(Setup)
        {
        }

        TEST_CLASS_CLEANUP(Cleanup)
        {
        }

        TEST_METHOD(RotationManagerNoItems)
        {
            CComPtr<IRotationManager> spRotationManager;
            Assert::IsTrue(CRotationManager::s_CreateInstance(&spRotationManager) == S_OK);
            Assert::IsTrue(spRotationManager->Start() == S_OK);
            Assert::IsTrue(spRotationManager->Shutdown() == S_OK);
        }

        TEST_METHOD(RotationManager1ItemInvalidPath)
        {
            CComPtr<IRotationManager> spRotationManager;
            Assert::IsTrue(CRotationManager::s_CreateInstance(&spRotationManager) == S_OK);

            CComPtr<IRotationItem> spRotationItem;
            Assert::IsTrue(CRotationItem::s_CreateInstance(L"c:\\foo.jpg", &spRotationItem) == S_OK);
            Assert::IsTrue(spRotationManager->AddItem(spRotationItem) == S_OK);

            Assert::IsTrue(spRotationManager->Start() == S_OK);

            // Verify rotation did not happen
            UINT originalOrientation = 0;
            BOOL isValidJPEG = FALSE;
            BOOL isLossless = FALSE;
            BOOL wasRotated = FALSE;
            HRESULT result = S_FALSE;
            Assert::IsTrue(spRotationItem->get_IsRotationLossless(&isLossless) == S_OK);
            Assert::IsTrue(spRotationItem->get_IsValidJPEG(&isValidJPEG) == S_OK);
            Assert::IsTrue(spRotationItem->get_WasRotated(&wasRotated) == S_OK);
            Assert::IsTrue(spRotationItem->get_OriginalOrientation(&originalOrientation) == S_OK);
            Assert::IsTrue(spRotationItem->get_Result(&result) == S_OK);
            Assert::IsTrue(originalOrientation == 1);
            Assert::IsTrue(isLossless == TRUE);
            Assert::IsTrue(isValidJPEG == FALSE);
            Assert::IsTrue(wasRotated == FALSE);
            Assert::IsTrue(result == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND));
        }

        TEST_METHOD(RotationManagerGetItemInvalidIndex)
        {
            CComPtr<IRotationManager> spRotationManager;
            Assert::IsTrue(CRotationManager::s_CreateInstance(&spRotationManager) == S_OK);

            CComPtr<IRotationItem> spRotationItem;
            Assert::IsTrue(CRotationItem::s_CreateInstance(L"c:\\foo.jpg", &spRotationItem) == S_OK);
            Assert::IsTrue(spRotationManager->AddItem(spRotationItem) == S_OK);

            CComPtr<IRotationItem> spRotationItemInvalidIndex;
            Assert::IsTrue(spRotationManager->GetItem(20, &spRotationItemInvalidIndex) == E_FAIL);
        }

        TEST_METHOD(RotationManager1ValidJPEG)
        {
            // Get test file source folder
            wchar_t testFolderSource[MAX_PATH] = { 0 };
            Assert::IsTrue(GetTestFolderPath(JPEGWITHEXIFROTATION_TESTFOLDER, ARRAYSIZE(testFolderSource), testFolderSource));

            // Get test file working folder
            wchar_t testFolderWorking[MAX_PATH] = { 0 };
            Assert::IsTrue(GetTestFolderPath(L"RotationManager1ValidJPEG", ARRAYSIZE(testFolderWorking), testFolderWorking));

            // Ensure test file working folder doesn't exist
            DeleteHelper(testFolderWorking);

            // Copy the test folder source to the working folder
            Assert::IsTrue(CopyHelper(testFolderSource, testFolderWorking));

            CComPtr<IRotationManager> spRotationManager;
            Assert::IsTrue(CRotationManager::s_CreateInstance(&spRotationManager) == S_OK);

            wchar_t testFilePath[MAX_PATH];
            Assert::IsTrue(PathCchCombine(testFilePath, ARRAYSIZE(testFilePath), testFolderWorking, g_rgTestFileData[0].filename) == S_OK);
            CComPtr<IRotationItem> spRotationItem;
            Assert::IsTrue(CRotationItem::s_CreateInstance(testFilePath, &spRotationItem) == S_OK);
            Assert::IsTrue(spRotationManager->AddItem(spRotationItem) == S_OK);

            Assert::IsTrue(spRotationManager->Start() == S_OK);

            // Verify rotation happened as it should
            UINT originalOrientation = 0;
            BOOL isValidJPEG = FALSE;
            BOOL isLossless = TRUE;
            BOOL wasRotated = FALSE;
            HRESULT result = S_FALSE;
            Assert::IsTrue(spRotationItem->get_IsRotationLossless(&isLossless) == S_OK);
            Assert::IsTrue(spRotationItem->get_IsValidJPEG(&isValidJPEG) == S_OK);
            Assert::IsTrue(spRotationItem->get_WasRotated(&wasRotated) == S_OK);
            Assert::IsTrue(spRotationItem->get_OriginalOrientation(&originalOrientation) == S_OK);
            Assert::IsTrue(spRotationItem->get_Result(&result) == S_OK);
            Assert::IsTrue(isLossless == FALSE);
            Assert::IsTrue(originalOrientation == g_rgTestFileData[0].originalOrientation);
            Assert::IsTrue(isValidJPEG == g_rgTestFileData[0].isValidJPEG);
            Assert::IsTrue(wasRotated == g_rgTestFileData[0].shouldRotate);
            Assert::IsTrue(result == S_OK);

            // Cleanup working folder
            DeleteHelper(testFolderWorking);
        }

        TEST_METHOD(RotateAllOrientationsTest)
        {
            // Get test file source folder
            wchar_t testFolderSource[MAX_PATH] = { 0 };
            Assert::IsTrue(GetTestFolderPath(JPEGWITHEXIFROTATION_TESTFOLDER, ARRAYSIZE(testFolderSource), testFolderSource));
            
            // Get test file working folder
            wchar_t testFolderWorking[MAX_PATH] = { 0 };
            Assert::IsTrue(GetTestFolderPath(L"RotateAllConfigurationsTest", ARRAYSIZE(testFolderWorking), testFolderWorking));

            // Ensure test file working folder doesn't exist
            DeleteHelper(testFolderWorking);

            // Copy the test folder source to the working folder
            Assert::IsTrue(CopyHelper(testFolderSource, testFolderWorking));

            CComPtr<IRotationManager> spRotationManager;
            Assert::IsTrue(CRotationManager::s_CreateInstance(&spRotationManager) == S_OK);
            
            for (int i = 0; i < ARRAYSIZE(g_rgTestFileData); i++)
            {
                wchar_t testFilePath[MAX_PATH];
                Assert::IsTrue(PathCchCombine(testFilePath, ARRAYSIZE(testFilePath), testFolderWorking, g_rgTestFileData[i].filename) == S_OK);
                CComPtr<IRotationItem> spRotationItem;
                Assert::IsTrue(CRotationItem::s_CreateInstance(testFilePath, &spRotationItem) == S_OK);
                Assert::IsTrue(spRotationManager->AddItem(spRotationItem) == S_OK);
            }

            Assert::IsTrue(spRotationManager->Start() == S_OK);

            // Verify rotation happened as it should
            for (int i = 0; i < ARRAYSIZE(g_rgTestFileData); i++)
            {
                CComPtr<IRotationItem> spRotationItem;
                Assert::IsTrue(spRotationManager->GetItem(i, &spRotationItem) == S_OK);
                UINT originalOrientation = 0;
                BOOL isValidJPEG = FALSE;
                BOOL isLossless = TRUE;
                BOOL wasRotated = FALSE;
                HRESULT result = S_FALSE;
                Assert::IsTrue(spRotationItem->get_IsRotationLossless(&isLossless) == S_OK);
                Assert::IsTrue(spRotationItem->get_IsValidJPEG(&isValidJPEG) == S_OK);
                Assert::IsTrue(spRotationItem->get_WasRotated(&wasRotated) == S_OK);
                Assert::IsTrue(spRotationItem->get_OriginalOrientation(&originalOrientation) == S_OK);
                Assert::IsTrue(spRotationItem->get_Result(&result) == S_OK);
                Assert::IsTrue(isLossless == FALSE);
                Assert::IsTrue(originalOrientation == g_rgTestFileData[i].originalOrientation);
                Assert::IsTrue(isValidJPEG == g_rgTestFileData[i].isValidJPEG);
                Assert::IsTrue(wasRotated == g_rgTestFileData[i].shouldRotate);
                Assert::IsTrue(result == S_OK);
            }

            // Cleanup working folder
            DeleteHelper(testFolderWorking);
        }

        TEST_METHOD(RotationManagerEventsVerification)
        {
            CMockRotationManagerEvents* mockEvents = new CMockRotationManagerEvents();
            CComPtr<IRotationManagerEvents> rotationManagerEvents;
            Assert::IsTrue(mockEvents->QueryInterface(IID_PPV_ARGS(&rotationManagerEvents)) == S_OK);
            CComPtr<IRotationManager> sprm;
            Assert::IsTrue(CRotationManager::s_CreateInstance(&sprm) == S_OK);
            DWORD cookie = 0;
            Assert::IsTrue(sprm->Advise(rotationManagerEvents, &cookie) == S_OK);
            Assert::IsTrue(cookie != 0);

            // Add some mock items to the manager
            for (int i = 0; i < 100; i++)
            {
                CComPtr<IRotationItem> spItem;
                CMockRotationItem *pmri = new CMockRotationItem();
                pmri->m_isValidJPEG = true;
                pmri->m_wasRotated = true;
                pmri->m_result = S_FALSE;
                pmri->m_originalOrientation = 2;
                pmri->QueryInterface(IID_PPV_ARGS(&spItem));
                wchar_t path[MAX_PATH] = { 0 };
                Assert::IsTrue(StringCchPrintf(path, ARRAYSIZE(path), L"foo%d.jpg", i) == S_OK);
                Assert::IsTrue(spItem->put_Path(path) == S_OK);
                Assert::IsTrue(sprm->AddItem(spItem) == S_OK);
                pmri->Release();
            }

            // Run the manager
            Assert::IsTrue(sprm->Start() == S_OK);

            // Verify events
            Assert::IsTrue(mockEvents->m_canceled == false);
            Assert::IsTrue(mockEvents->m_completed == true);
            Assert::IsTrue(mockEvents->m_itemsAdded == 100);
            Assert::IsTrue(mockEvents->m_itemsProcessed == 100);
            Assert::IsTrue(mockEvents->m_numTimeOnProgressCalled == 100);
            Assert::IsTrue(mockEvents->m_total == 100);
            Assert::IsTrue(mockEvents->m_lastCompleted == 100);

            Assert::IsTrue(sprm->UnAdvise(cookie) == S_OK);

            mockEvents->Release();
        }

        TEST_METHOD(RotationManagerEventsVerifyCancel)
        {
            class CMockRotationManagerEventsWithCancel : public CMockRotationManagerEvents
            {
            public:
                IFACEMETHODIMP OnProgress(_In_ UINT completed, _In_ UINT total)
                {
                    CMockRotationManagerEvents::OnProgress(completed, total);
                    if (completed == 1)
                    {
                        // Cancel halfway through
                        m_sprm->Cancel();
                    }
                    return S_OK;
                }

                CComPtr<IRotationManager> m_sprm;
            };

            CMockRotationManagerEventsWithCancel* mockEvents = new CMockRotationManagerEventsWithCancel();
            CComPtr<IRotationManagerEvents> rotationManagerEvents;
            Assert::IsTrue(mockEvents->QueryInterface(IID_PPV_ARGS(&rotationManagerEvents)) == S_OK);
            CComPtr<IRotationManager> sprm;
            Assert::IsTrue(CRotationManager::s_CreateInstance(&sprm) == S_OK);
            mockEvents->m_sprm = sprm;
            DWORD cookie = 0;
            Assert::IsTrue(sprm->Advise(rotationManagerEvents, &cookie) == S_OK);
            Assert::IsTrue(cookie != 0);

            // Add some mock items to the manager
            for (int i = 0; i < 100; i++)
            {
                CComPtr<IRotationItem> spItem;
                CMockRotationItem *pmri = new CMockRotationItem();
                pmri->m_isValidJPEG = true;
                pmri->m_wasRotated = true;
                pmri->m_result = S_FALSE;
                pmri->m_originalOrientation = 2;
                pmri->QueryInterface(IID_PPV_ARGS(&spItem));
                wchar_t path[MAX_PATH] = { 0 };
                Assert::IsTrue(StringCchPrintf(path, ARRAYSIZE(path), L"foo%d.jpg", i) == S_OK);
                Assert::IsTrue(spItem->put_Path(path) == S_OK);
                Assert::IsTrue(sprm->AddItem(spItem) == S_OK);
                pmri->Release();
            }

            // Run the manager
            Assert::IsTrue(sprm->Start() == S_OK);

            // Verify events
            Assert::IsTrue(mockEvents->m_canceled == true);
            Assert::IsTrue(mockEvents->m_completed == true);
            Assert::IsTrue(mockEvents->m_total == 100);

            Assert::IsTrue(sprm->UnAdvise(cookie) == S_OK);

            mockEvents->Release();
        }

        TEST_METHOD(RotationManagerLosslessOnly)
        {
            // Get test file source folder
            wchar_t testFolderSource[MAX_PATH] = { 0 };
            Assert::IsTrue(GetTestFolderPath(JPEGWITHEXIFROTATION_TESTFOLDER, ARRAYSIZE(testFolderSource), testFolderSource));

            // Get test file working folder
            wchar_t testFolderWorking[MAX_PATH] = { 0 };
            Assert::IsTrue(GetTestFolderPath(L"RotationManagerLosslessOnly", ARRAYSIZE(testFolderWorking), testFolderWorking));

            // Ensure test file working folder doesn't exist
            DeleteHelper(testFolderWorking);

            // Copy the test folder source to the working folder
            Assert::IsTrue(CopyHelper(testFolderSource, testFolderWorking));

            CComPtr<IRotationManager> spRotationManager;
            Assert::IsTrue(CRotationManager::s_CreateInstance(&spRotationManager) == S_OK);
            
            // Configure the manager for lossless only rotation
            CComPtr<IRotationManagerDiagnostics> spRotationManagerDiagnostics;
            Assert::IsTrue(spRotationManager->QueryInterface(&spRotationManagerDiagnostics) == S_OK);
            Assert::IsTrue(spRotationManagerDiagnostics->put_LosslessOnly(TRUE) == S_OK);

            wchar_t testFilePath[MAX_PATH];
            Assert::IsTrue(PathCchCombine(testFilePath, ARRAYSIZE(testFilePath), testFolderWorking, g_rgTestFileData[1].filename) == S_OK);
            CComPtr<IRotationItem> spRotationItem;
            Assert::IsTrue(CRotationItem::s_CreateInstance(testFilePath, &spRotationItem) == S_OK);
            Assert::IsTrue(spRotationManager->AddItem(spRotationItem) == S_OK);

            Assert::IsTrue(spRotationManager->Start() == S_OK);
            
            // Verify rotation did not happen since it was not lossless
            UINT originalOrientation = 0;
            BOOL isValidJPEG = FALSE;
            BOOL isLossless = TRUE;
            BOOL wasRotated = FALSE;
            HRESULT result = S_FALSE;
            Assert::IsTrue(spRotationItem->get_IsRotationLossless(&isLossless) == S_OK);
            Assert::IsTrue(spRotationItem->get_IsValidJPEG(&isValidJPEG) == S_OK);
            Assert::IsTrue(spRotationItem->get_WasRotated(&wasRotated) == S_OK);
            Assert::IsTrue(spRotationItem->get_OriginalOrientation(&originalOrientation) == S_OK);
            Assert::IsTrue(spRotationItem->get_Result(&result) == S_OK);
            Assert::IsTrue(isLossless == FALSE);
            Assert::IsTrue(originalOrientation == g_rgTestFileData[1].originalOrientation);
            Assert::IsTrue(isValidJPEG == g_rgTestFileData[1].isValidJPEG);
            Assert::IsTrue(wasRotated == FALSE);
            Assert::IsTrue(result == S_OK);

            // Cleanup working folder
            DeleteHelper(testFolderWorking);
        }

        TEST_METHOD(RotationManagerPreviewOnly)
        {
            // Get test file source folder
            wchar_t testFolderSource[MAX_PATH] = { 0 };
            Assert::IsTrue(GetTestFolderPath(JPEGWITHEXIFROTATION_TESTFOLDER, ARRAYSIZE(testFolderSource), testFolderSource));

            // Get test file working folder
            wchar_t testFolderWorking[MAX_PATH] = { 0 };
            Assert::IsTrue(GetTestFolderPath(L"RotationManagerPreviewOnly", ARRAYSIZE(testFolderWorking), testFolderWorking));

            // Ensure test file working folder doesn't exist
            DeleteHelper(testFolderWorking);

            // Copy the test folder source to the working folder
            Assert::IsTrue(CopyHelper(testFolderSource, testFolderWorking));

            CComPtr<IRotationManager> spRotationManager;
            Assert::IsTrue(CRotationManager::s_CreateInstance(&spRotationManager) == S_OK);

            // Configure the manager for preview only
            CComPtr<IRotationManagerDiagnostics> spRotationManagerDiagnostics;
            Assert::IsTrue(spRotationManager->QueryInterface(&spRotationManagerDiagnostics) == S_OK);
            Assert::IsTrue(spRotationManagerDiagnostics->put_PreviewOnly(TRUE) == S_OK);

            wchar_t testFilePath[MAX_PATH];
            Assert::IsTrue(PathCchCombine(testFilePath, ARRAYSIZE(testFilePath), testFolderWorking, g_rgTestFileData[1].filename) == S_OK);
            CComPtr<IRotationItem> spRotationItem;
            Assert::IsTrue(CRotationItem::s_CreateInstance(testFilePath, &spRotationItem) == S_OK);
            Assert::IsTrue(spRotationManager->AddItem(spRotationItem) == S_OK);

            Assert::IsTrue(spRotationManager->Start() == S_OK);

            // Verify rotation did not happen since we specified preview only
            UINT originalOrientation = 0;
            BOOL isValidJPEG = FALSE;
            BOOL isLossless = TRUE;
            BOOL wasRotated = FALSE;
            HRESULT result = S_FALSE;
            Assert::IsTrue(spRotationItem->get_IsRotationLossless(&isLossless) == S_OK);
            Assert::IsTrue(spRotationItem->get_IsValidJPEG(&isValidJPEG) == S_OK);
            Assert::IsTrue(spRotationItem->get_WasRotated(&wasRotated) == S_OK);
            Assert::IsTrue(spRotationItem->get_OriginalOrientation(&originalOrientation) == S_OK);
            Assert::IsTrue(spRotationItem->get_Result(&result) == S_OK);
            Assert::IsTrue(isLossless == FALSE);
            Assert::IsTrue(originalOrientation == g_rgTestFileData[1].originalOrientation);
            Assert::IsTrue(isValidJPEG == g_rgTestFileData[1].isValidJPEG);
            Assert::IsTrue(wasRotated == FALSE);
            Assert::IsTrue(result == S_OK);

            // Cleanup working folder
            DeleteHelper(testFolderWorking);
        }

        TEST_METHOD(EnumSubfolderOnTest)
        {
            // Get test file source folder
            wchar_t testFolderSource[MAX_PATH] = { 0 };
            Assert::IsTrue(GetTestFolderPath(JPEGWITHEXIFROTATION_TESTFOLDER, ARRAYSIZE(testFolderSource), testFolderSource));

            // Get test file working folder
            wchar_t testFolderWorking[MAX_PATH] = { 0 };
            Assert::IsTrue(GetTestFolderPath(L"EnumSubfolderOnTest", ARRAYSIZE(testFolderWorking), testFolderWorking));

            // Ensure test file working folder doesn't exist
            DeleteHelper(testFolderWorking);

            // Copy the test folder source to the working folder
            Assert::IsTrue(CopyHelper(testFolderSource, testFolderWorking));

            wchar_t testFolderWorkingSub[MAX_PATH] = { 0 };
            Assert::IsTrue(StringCchCopy(testFolderWorkingSub, ARRAYSIZE(testFolderWorkingSub), testFolderWorking) == S_OK);
            Assert::IsTrue(PathCchAppend(testFolderWorkingSub, ARRAYSIZE(testFolderWorkingSub), L"EnumSubFolderOnTestSubFolder") == S_OK);

            // Copy the test folder source to the working sub folder
            Assert::IsTrue(CopyHelper(testFolderSource, testFolderWorkingSub));

            CComPtr<IRotationManager> spRotationManager;
            Assert::IsTrue(CRotationManager::s_CreateInstance(&spRotationManager) == S_OK);

            // Pass the root test folder.  The manager will enumerate the contents and subfolder contents
            Assert::IsTrue(spRotationManager->AddPath(testFolderWorking) == S_OK);

            // Verify we have the correct number of items enumerated and added to the roaming manager
            UINT numItems = 0;
            Assert::IsTrue(spRotationManager->GetItemCount(&numItems) == S_OK);
            Assert::IsTrue(numItems == ARRAYSIZE(g_rgTestFileData) * 2);

            // Cleanup working folder
            DeleteHelper(testFolderWorking);
        }

        TEST_METHOD(EnumSubfolderOffTest)
        {
            // Get test file source folder
            wchar_t testFolderSource[MAX_PATH] = { 0 };
            Assert::IsTrue(GetTestFolderPath(JPEGWITHEXIFROTATION_TESTFOLDER, ARRAYSIZE(testFolderSource), testFolderSource));

            // Get test file working folder
            wchar_t testFolderWorking[MAX_PATH] = { 0 };
            Assert::IsTrue(GetTestFolderPath(L"EnumSubfolderOffTest", ARRAYSIZE(testFolderWorking), testFolderWorking));

            // Ensure test file working folder doesn't exist
            DeleteHelper(testFolderWorking);

            // Copy the test folder source to the working folder
            Assert::IsTrue(CopyHelper(testFolderSource, testFolderWorking));

            wchar_t testFolderWorkingSub[MAX_PATH] = { 0 };
            Assert::IsTrue(StringCchCopy(testFolderWorkingSub, ARRAYSIZE(testFolderWorkingSub), testFolderWorking) == S_OK);
            Assert::IsTrue(PathCchAppend(testFolderWorkingSub, ARRAYSIZE(testFolderWorkingSub), L"EnumSubFolderOffTestSubFolder") == S_OK);

            // Copy the test folder source to the working sub folder
            Assert::IsTrue(CopyHelper(testFolderSource, testFolderWorkingSub));

            CComPtr<IRotationManager> spRotationManager;
            Assert::IsTrue(CRotationManager::s_CreateInstance(&spRotationManager) == S_OK);

            // Configure the manager to not enumerate sub folders
            CComPtr<IRotationManagerDiagnostics> spRotationManagerDiagnostics;
            Assert::IsTrue(spRotationManager->QueryInterface(&spRotationManagerDiagnostics) == S_OK);
            Assert::IsTrue(spRotationManagerDiagnostics->put_EnumerateSubFolders(FALSE) == S_OK);

            // Pass the root test folder.  The manager will enumerate the contents and subfolder contents
            Assert::IsTrue(spRotationManager->AddPath(testFolderWorking) == S_OK);

            // Verify we have the correct number of items enumerated and added to the roaming manager
            UINT numItems = 0;
            Assert::IsTrue(spRotationManager->GetItemCount(&numItems) == S_OK);
            Assert::IsTrue(numItems == ARRAYSIZE(g_rgTestFileData));

            // Cleanup working folder
            DeleteHelper(testFolderWorking);
        }
    };
}