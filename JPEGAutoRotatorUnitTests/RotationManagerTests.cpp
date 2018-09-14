#include "stdafx.h"
#include "CppUnitTest.h"
#include <RotationInterfaces.h>
#include <RotationManager.h>
#include <pathcch.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

constexpr auto JPEGWITHEXIFROTATION_TESTFOLDER = L"TestFiles\\JPEGWithExifRotation\\";

extern HINSTANCE g_hInstance;

namespace JPEGAutoRotatorUnitTests
{
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

    TEST_CLASS(RotationManagerTests)
    {
    public:
        TEST_CLASS_INITIALIZE(Setup)
        {
        }

        TEST_CLASS_CLEANUP(Cleanup)
        {
        }
        TEST_METHOD(RotateAllConfigurationsTest)
        {
            // Get test file source folder
            wchar_t testFolderSource[MAX_PATH];
            Assert::IsTrue(GetTestFolderPath(JPEGWITHEXIFROTATION_TESTFOLDER, ARRAYSIZE(testFolderSource), testFolderSource));
            
            // Get test file working folder
            wchar_t testFolderWorking[MAX_PATH];
            Assert::IsTrue(GetTestFolderPath(L"RotateAllConfigurationsTest", ARRAYSIZE(testFolderWorking), testFolderWorking));

            // Ensure test file working folder doesn't exist
            DeleteHelper(testFolderWorking);

            // Copy the test folder source to the working folder
            Assert::IsTrue(CopyHelper(testFolderSource, testFolderWorking));

            CComPtr<IRotationManager> spRotationManager;
            Assert::IsTrue(CRotationManager::s_CreateInstance(&spRotationManager) == S_OK);
            
            for (int i = 0; i < ARRAYSIZE(g_rgTestJPEGFiles); i++)
            {
                wchar_t testFilePath[MAX_PATH];
                Assert::IsTrue(PathCchCombine(testFilePath, ARRAYSIZE(testFilePath), testFolderWorking, g_rgTestJPEGFiles[i]) == S_OK);
                CComPtr<IRotationItem> spRotationItem;
                Assert::IsTrue(CRotationItem::s_CreateInstance(testFilePath, &spRotationItem) == S_OK);
                Assert::IsTrue(spRotationManager->AddItem(spRotationItem) == S_OK);
            }

            Assert::IsTrue(spRotationManager->Start() == S_OK);

            DeleteHelper(testFolderWorking);
        }
    };
}