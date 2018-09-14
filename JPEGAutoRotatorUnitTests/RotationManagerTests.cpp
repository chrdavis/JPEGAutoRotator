#include "stdafx.h"
#include "CppUnitTest.h"
#include <RotationInterfaces.h>
#include <RotationManager.h>
#include <pathcch.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

#define JPEGWITHEXIFROTATION_UNITTEST_FOLDER L"Images\\JPEGWithExifRotation\\"

extern HINSTANCE g_hInstance;

namespace JPEGAutoRotatorUnitTests
{
    // TODO: add test files to the unit tests project
    // https://msdn.microsoft.com/en-us/library/ms182475.aspx
    PCWSTR g_rgTestFiles[] =
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
            /*ExpandEnvironmentStrings(g_pszTestOutputFolder, g_szPathOut, ARRAYSIZE(g_szPathOut));
            // Create the test directory
            CreateDirectory(g_szPathOut, nullptr);*/
        }

        TEST_CLASS_CLEANUP(Cleanup)
        {
            // Delete the test directory
            /*SHFILEOPSTRUCT fileStruct = { 0 };
            fileStruct.pFrom = g_szPathOut;
            fileStruct.wFunc = FO_DELETE;
            fileStruct.fFlags = FOF_SILENT | FOF_NOERRORUI | FOF_NO_UI;
            SHFileOperation(&fileStruct);*/
        }

        TEST_METHOD_INITIALIZE(TestSetup)
        {
            // Get current path
            wchar_t testFolderPath[MAX_PATH];
            Assert::IsTrue(GetCurrentFolderPath(ARRAYSIZE(testFolderPath), testFolderPath));
            // Append test folder path
            Assert::IsTrue(PathCchAppend(testFolderPath, ARRAYSIZE(testFolderPath), JPEGWITHEXIFROTATION_UNITTEST_FOLDER) == S_OK);
            // Copy test files
        }

        TEST_METHOD_CLEANUP(TestCleanup)
        {
            wchar_t testFolderPath[MAX_PATH];

            // Get current path
            Assert::IsTrue(GetCurrentFolderPath(ARRAYSIZE(testFolderPath), testFolderPath));

            // Append test folder path
            Assert::IsTrue(PathCchAppend(testFolderPath, ARRAYSIZE(testFolderPath), JPEGWITHEXIFROTATION_UNITTEST_FOLDER) == S_OK);

            // Delete test folder
            DeleteHelper(testFolderPath);
        }
        
        TEST_METHOD(RotateAllConfigurationsTest)
        {
            
            CComPtr<IRotationManager> spRotationManager;
            Assert::IsTrue(CRotationManager::s_CreateInstance(&spRotationManager) == S_OK);
            
            for (int i = 0; i < ARRAYSIZE(g_rgTestFiles); i++)
            {
                CComPtr<IRotationItem> spRotationItem;
                Assert::IsTrue(CRotationItem::s_CreateInstance(g_rgTestFiles[i], &spRotationItem) == S_OK);
                Assert::IsTrue(spRotationManager->AddItem(spRotationItem) == S_OK);
            }

            Assert::IsTrue(spRotationManager->Start() == S_OK);
        }
    };
}