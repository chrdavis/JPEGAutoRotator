#include "stdafx.h"
#include "CppUnitTest.h"
#include <RotationInterfaces.h>
#include <RotationUI.h>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace JPEGAutoRotatorUnitTests
{
    
    TEST_CLASS(RotationUITests)
    {
    public:
        TEST_CLASS_INITIALIZE(Setup)
        {
            // Create test jpeg files (and subfolders)
        }

        TEST_CLASS_CLEANUP(Cleanup)
        {
            // Delete test jpeg files (and subfolders)
        }

        TEST_METHOD(VerifyRotationUICreation)
        {
            // TODO: Need a mock rotation manager and a IDataObject of some dummy jpeg files?
            //CComPtr<IRotationUI> sprui;
            //if (SUCCEEDED(CRotationUI::s_CreateInstance(spdo, sprm, &sprui)))
            //{
            //    sprui->Start();
            //}
        }

    };
}