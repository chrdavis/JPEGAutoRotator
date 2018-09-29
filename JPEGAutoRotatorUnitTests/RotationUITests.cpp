#include "stdafx.h"
#include "CppUnitTest.h"
#include <RotationInterfaces.h>
#include <RotationManager.h>
#include "MockRotationItem.h"
#include <RotationUI.h>
#include <strsafe.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace JPEGAutoRotatorUnitTests
{
    
    TEST_CLASS(RotationUITests)
    {
    public:
        TEST_CLASS_INITIALIZE(Setup)
        {
        }

        TEST_CLASS_CLEANUP(Cleanup)
        {
        }

        TEST_METHOD(VerifyRotationUICreation)
        {
            CComPtr<IRotationManager> sprm;
            Assert::IsTrue(CRotationManager::s_CreateInstance(&sprm) == S_OK);
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

            // Create the Rotation UI
            CComPtr<IRotationUI> sprui;
            Assert::IsTrue(CRotationUI::s_CreateInstance(sprm, nullptr, &sprui) == S_OK);
            // Start the operation through the UI
            Assert::IsTrue(sprui->Start() == S_OK);
        }
    };
}