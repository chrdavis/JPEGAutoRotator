#pragma once
#include <ShlObj.h>
#include "RotationInterfaces.h"

class CMockRotationManagerEvents :
    public IRotationManagerEvents
{
public:
    CMockRotationManagerEvents() {}

    // IUnknown
    IFACEMETHODIMP QueryInterface(_In_ REFIID riid, _COM_Outptr_ void** ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CMockRotationManagerEvents, IRotationManagerEvents),
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

    // IRotationManagerEvents
    IFACEMETHODIMP OnItemAdded(_In_ UINT)
    {
        m_itemsAdded++;
        return S_OK;
    }

    IFACEMETHODIMP OnItemProcessed(_In_ UINT)
    {
        m_itemsProcessed++;
        return S_OK;
    }

    IFACEMETHODIMP OnProgress(_In_ UINT completed, _In_ UINT total)
    {
        m_total = total;
        m_lastCompleted = completed;
        m_numTimeOnProgressCalled++;
        return S_OK;
    }

    IFACEMETHODIMP OnCanceled()
    {
        m_canceled = true;
        return S_OK;
    }

    IFACEMETHODIMP OnCompleted()
    {
        m_completed = true;
        return S_OK;
    }

    virtual ~CMockRotationManagerEvents() {}

    bool m_canceled = false;
    bool m_completed = false;
    UINT m_total = 0;
    UINT m_lastCompleted = 0;
    UINT m_itemsAdded = 0;
    UINT m_itemsProcessed = 0;
    UINT m_numTimeOnProgressCalled = 0;
    long m_refCount = 1;
    DWORD m_cookie = 0;
};