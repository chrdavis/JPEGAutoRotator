#pragma once
#include "stdafx.h"

interface __declspec(uuid("FFB76FF1-DE4B-459A-A6C7-025392CFDA53")) IRotationItem : public IUnknown
{
public:
    IFACEMETHOD(get_Path)(_Outptr_ PWSTR* ppszPath) = 0;
    IFACEMETHOD(put_Path)(_In_ PCWSTR pszPath) = 0;
    IFACEMETHOD(get_WasRotated)(_Out_ BOOL* pfWasRotated) = 0;
    IFACEMETHOD(get_IsValidJPEG)(_Out_ BOOL* pfIsValidJPEG) = 0;
    IFACEMETHOD(get_IsRotationLossless)(_Out_ BOOL* pfIsRotationLossless) = 0;
    IFACEMETHOD(get_OriginalOrientation)(_Out_ UINT* puOriginalOrientation) = 0;
    IFACEMETHOD(get_Result)(_Out_ HRESULT* phrResult) = 0;
    IFACEMETHOD(put_Result)(_In_ HRESULT hrResult) = 0;
    IFACEMETHOD(Rotate)() = 0;
};

interface __declspec(uuid("1FB14328-E681-4BC2-B947-00BAC0387A01")) IRotationManagerEvents : public IUnknown
{
public:
    IFACEMETHOD(OnItemAdded)(_In_ UINT uIndex) = 0;
    IFACEMETHOD(OnItemProcessed)(_In_ UINT uIndex) = 0;
    IFACEMETHOD(OnProgress)(_In_ UINT uCompleted, _In_ UINT uTotal) = 0;
    IFACEMETHOD(OnCanceled)() = 0;
    IFACEMETHOD(OnCompleted)() = 0;
};

interface __declspec(uuid("D78A5931-DE20-4CB7-A90F-B2BB1B31BE41")) IRotationManager : public IUnknown
{
public:
    IFACEMETHOD(Advise)(_In_ IRotationManagerEvents* prme, _Out_ DWORD* pdwCookie) = 0;
    IFACEMETHOD(UnAdvise)(_In_ DWORD dwCookie) = 0;
    IFACEMETHOD(Start)() = 0;
    IFACEMETHOD(Cancel)() = 0;
    IFACEMETHOD(Shutdown)() = 0;
    IFACEMETHOD(AddItem)(_In_ IRotationItem* pri) = 0;
    IFACEMETHOD(GetItem)(_In_ UINT uIndex, _COM_Outptr_ IRotationItem** ppri) = 0;
    IFACEMETHOD(GetItemCount)(_Out_ UINT* puCount) = 0;
    IFACEMETHOD(get_WorkerThreadCount)(_Out_ UINT* puThreadCount) = 0;
    IFACEMETHOD(put_WorkerThreadCount)(_In_ UINT uThreadCount) = 0;
    IFACEMETHOD(get_ItemsPerWorkerThread)(_Out_ UINT* numItemsPerThread) = 0;
    IFACEMETHOD(put_ItemsPerWorkerThread)(_In_ UINT numItemsPerThread) = 0;
};

interface __declspec(uuid("D1952AE2-93FF-4C5D-9AC5-C6DADBEDE242")) IRotationManagerDiagnostics : public IUnknown
{
public:
    IFACEMETHOD(get_MaxWorkerThreadCount)(_Out_ UINT* puMaxThreadCount) = 0;
    IFACEMETHOD(put_MaxWorkerThreadCount)(_In_ UINT uMaxThreadCount) = 0;
    IFACEMETHOD(get_WorkerThreadCount)(_Out_ UINT* puThreadCount) = 0;
    IFACEMETHOD(put_WorkerThreadCount)(_In_ UINT uThreadCount) = 0;
    IFACEMETHOD(get_MinItemsPerWorkerThread)(_Out_ UINT* puMinItemsPerThread) = 0;
    IFACEMETHOD(put_MinItemsPerWorkerThread)(_In_ UINT uMinItemsPerThread) = 0;
    IFACEMETHOD(get_ItemsPerWorkerThread)(_Out_ UINT* puNumItemsPerThread) = 0;
    IFACEMETHOD(put_ItemsPerWorkerThread)(_In_ UINT uNumItemsPerThread) = 0;
};

interface __declspec(uuid("39EEEC7B-81B4-4719-90E1-37D70059724C")) IRotationUI : public IUnknown
{
public:
    IFACEMETHOD(Initialize)(_In_ IDataObject* pdo) = 0;
    IFACEMETHOD(Start)() = 0;
    IFACEMETHOD(Close)() = 0;
};