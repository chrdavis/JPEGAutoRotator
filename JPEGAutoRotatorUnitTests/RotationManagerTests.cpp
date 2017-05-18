#include "stdafx.h"
#include "CppUnitTest.h"
#include <RotationInterfaces.h>
#include <RotationManager.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

#define SZ_ROOT_TEST_FOLDER L"C:\\EXIF TEST FILES\\"

namespace JPEGAutoRotatorUnitTests
{
    PCWSTR g_rgTestFiles[] =
    {
        SZ_ROOT_TEST_FOLDER L"Landscape_1.jpg",
        SZ_ROOT_TEST_FOLDER L"Landscape_2.jpg",
        SZ_ROOT_TEST_FOLDER L"Landscape_3.jpg",
        SZ_ROOT_TEST_FOLDER L"Landscape_4.jpg",
        SZ_ROOT_TEST_FOLDER L"Landscape_5.jpg",
        SZ_ROOT_TEST_FOLDER L"Landscape_6.jpg",
        SZ_ROOT_TEST_FOLDER L"Landscape_7.jpg",
        SZ_ROOT_TEST_FOLDER L"Landscape_8.jpg",
        SZ_ROOT_TEST_FOLDER L"Portrait_1.jpg",
        SZ_ROOT_TEST_FOLDER L"Portrait_2.jpg",
        SZ_ROOT_TEST_FOLDER L"Portrait_3.jpg",
        SZ_ROOT_TEST_FOLDER L"Portrait_4.jpg",
        SZ_ROOT_TEST_FOLDER L"Portrait_5.jpg",
        SZ_ROOT_TEST_FOLDER L"Portrait_6.jpg",
        SZ_ROOT_TEST_FOLDER L"Portrait_7.jpg",
        SZ_ROOT_TEST_FOLDER L"Portrait_8.jpg"
    };

    TEST_CLASS(RotationManagerTests)
    {
    public:
        TEST_CLASS_INITIALIZE(Setup)
        {
            /*ExpandEnvironmentStrings(g_pszTestOutputFolder, g_szPathOut, ARRAYSIZE(g_szPathOut));
            // Create the test directory
            CreateDirectory(g_szPathOut, nullptr);
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

        TEST_METHOD(TestMethod1)
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