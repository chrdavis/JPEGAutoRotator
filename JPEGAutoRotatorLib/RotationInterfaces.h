#pragma once
#include "stdafx.h"

interface __declspec(uuid("FFB76FF1-DE4B-459A-A6C7-025392CFDA53")) IRotationItem : public IUnknown
{
public:
    IFACEMETHOD(get_Path)(__deref_out PWSTR* ppszPath) = 0;
    IFACEMETHOD(put_Path)(__in PCWSTR pszPath) = 0;
    IFACEMETHOD(get_WasRotated)(__out BOOL* pfWasRotated) = 0;
    IFACEMETHOD(get_IsValidJPEG)(__out BOOL* pfIsValidJPEG) = 0;
    IFACEMETHOD(get_IsRotationLossless)(__out BOOL* pfIsRotationLossless) = 0;
    IFACEMETHOD(get_OriginalOrientation)(__out UINT* puOriginalOrientation) = 0;
    IFACEMETHOD(get_Result)(__out HRESULT* phrResult) = 0;
    IFACEMETHOD(put_Result)(__in HRESULT hrResult) = 0;
    IFACEMETHOD(Rotate)() = 0;
};

interface __declspec(uuid("1FB14328-E681-4BC2-B947-00BAC0387A01")) IRotationManagerEvents : public IUnknown
{
public:
    IFACEMETHOD(OnItemAdded)(__in UINT uIndex) = 0;
    IFACEMETHOD(OnItemProcessed)(__in UINT uIndex) = 0;
    IFACEMETHOD(OnProgress)(__in UINT uCompleted, __in UINT uTotal) = 0;
    IFACEMETHOD(OnCanceled)() = 0;
    IFACEMETHOD(OnCompleted)() = 0;
};

interface __declspec(uuid("D78A5931-DE20-4CB7-A90F-B2BB1B31BE41")) IRotationManager : public IUnknown
{
public:
    IFACEMETHOD(Advise)(__in IRotationManagerEvents* prme, __out DWORD* pdwCookie) = 0;
    IFACEMETHOD(UnAdvise)(__in DWORD dwCookie) = 0;
    IFACEMETHOD(Start)() = 0;
    IFACEMETHOD(Cancel)() = 0;
    IFACEMETHOD(Shutdown)() = 0;
    IFACEMETHOD(AddItem)(__in IRotationItem* pri) = 0;
    IFACEMETHOD(GetItem)(__in UINT uIndex, __deref_out IRotationItem** ppri) = 0;
    IFACEMETHOD(GetItemCount)(__out UINT* puCount) = 0;
    IFACEMETHOD(get_WorkerThreadCount)(__out UINT* puThreadCount) = 0;
    IFACEMETHOD(put_WorkerThreadCount)(__in UINT uThreadCount) = 0;
    IFACEMETHOD(get_ItemsPerWorkerThread)(__out UINT* numItemsPerThread) = 0;
    IFACEMETHOD(put_ItemsPerWorkerThread)(__in UINT numItemsPerThread) = 0;
};

interface __declspec(uuid("D1952AE2-93FF-4C5D-9AC5-C6DADBEDE242")) IRotationManagerDiagnostics : public IUnknown
{
public:
    IFACEMETHOD(get_MaxWorkerThreadCount)(__out UINT* puMaxThreadCount) = 0;
    IFACEMETHOD(put_MaxWorkerThreadCount)(__in UINT uMaxThreadCount) = 0;
    IFACEMETHOD(get_WorkerThreadCount)(__out UINT* puThreadCount) = 0;
    IFACEMETHOD(put_WorkerThreadCount)(__in UINT uThreadCount) = 0;
    IFACEMETHOD(get_MinItemsPerWorkerThread)(__out UINT* puMinItemsPerThread) = 0;
    IFACEMETHOD(put_MinItemsPerWorkerThread)(__in UINT uMinItemsPerThread) = 0;
    IFACEMETHOD(get_ItemsPerWorkerThread)(__out UINT* puNumItemsPerThread) = 0;
    IFACEMETHOD(put_ItemsPerWorkerThread)(__in UINT uNumItemsPerThread) = 0;
};

interface __declspec(uuid("39EEEC7B-81B4-4719-90E1-37D70059724C")) IRotationUI : public IUnknown
{
public:
    IFACEMETHOD(Initialize)(__in IDataObject* pdo) = 0;
    IFACEMETHOD(Start)() = 0;
    IFACEMETHOD(Close)() = 0;
};