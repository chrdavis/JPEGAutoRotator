#pragma once
#include "stdafx.h"

interface __declspec(uuid("FFB76FF1-DE4B-459A-A6C7-025392CFDA53")) IRotationItem : public IUnknown
{
public:
    IFACEMETHOD(get_Path)(_Outptr_ PWSTR* path) = 0;
    IFACEMETHOD(put_Path)(_In_ PCWSTR path) = 0;
    IFACEMETHOD(get_WasRotated)(_Out_ BOOL* wasRotated) = 0;
    IFACEMETHOD(get_IsValidJPEG)(_Out_ BOOL* isValidJPEG) = 0;
    IFACEMETHOD(get_IsRotationLossless)(_Out_ BOOL* isRotationLossless) = 0;
    IFACEMETHOD(get_OriginalOrientation)(_Out_ UINT* originalOrientation) = 0;
    IFACEMETHOD(get_Result)(_Out_ HRESULT* result) = 0;
    IFACEMETHOD(put_Result)(_In_ HRESULT result) = 0;
    IFACEMETHOD(Load)() = 0;
    IFACEMETHOD(Rotate)() = 0;
};

interface __declspec(uuid("9A0D1E02-DAC8-4464-B483-193AEF022478")) IRotationItemFactory : public IUnknown
{
public:
    IFACEMETHOD(Create)(_COM_Outptr_ IRotationItem** ppItem) = 0;
};

interface __declspec(uuid("1FB14328-E681-4BC2-B947-00BAC0387A01")) IRotationManagerEvents : public IUnknown
{
public:
    IFACEMETHOD(OnItemAdded)(_In_ UINT index) = 0;
    IFACEMETHOD(OnItemProcessed)(_In_ UINT index) = 0;
    IFACEMETHOD(OnProgress)(_In_ UINT completed, _In_ UINT total) = 0;
    IFACEMETHOD(OnCanceled)() = 0;
    IFACEMETHOD(OnCompleted)() = 0;
};

interface __declspec(uuid("D78A5931-DE20-4CB7-A90F-B2BB1B31BE41")) IRotationManager : public IUnknown
{
public:
    IFACEMETHOD(Advise)(_In_ IRotationManagerEvents* pEvents, _Out_ DWORD* cookie) = 0;
    IFACEMETHOD(UnAdvise)(_In_ DWORD cookie) = 0;
    IFACEMETHOD(Start)() = 0;
    IFACEMETHOD(Cancel)() = 0;
    IFACEMETHOD(Shutdown)() = 0;
    IFACEMETHOD(AddPath)(_In_ PCWSTR path) = 0;
    IFACEMETHOD(AddItem)(_In_ IRotationItem* pItem) = 0;
    IFACEMETHOD(GetItem)(_In_ UINT index, _COM_Outptr_ IRotationItem** ppItem) = 0;
    IFACEMETHOD(GetItemCount)(_Out_ UINT* count) = 0;
    IFACEMETHOD(get_WorkerThreadCount)(_Out_ UINT* threadCount) = 0;
    IFACEMETHOD(put_WorkerThreadCount)(_In_ UINT threadCount) = 0;
    IFACEMETHOD(get_ItemsPerWorkerThread)(_Out_ UINT* numItemsPerThread) = 0;
    IFACEMETHOD(put_ItemsPerWorkerThread)(_In_ UINT numItemsPerThread) = 0;
    IFACEMETHOD(GetRotationItemFactory)(_In_ IRotationItemFactory** ppItemFactory) = 0;
    IFACEMETHOD(SetRotationItemFactory)(_In_ IRotationItemFactory* pItemFactory) = 0;
};

interface __declspec(uuid("D1952AE2-93FF-4C5D-9AC5-C6DADBEDE242")) IRotationManagerDiagnostics : public IUnknown
{
public:
    IFACEMETHOD(get_EnumerateSubFolders)(_Out_ BOOL* enumSubFolders) = 0;
    IFACEMETHOD(put_EnumerateSubFolders)(_In_ BOOL enumSubFolders) = 0;
    IFACEMETHOD(get_LosslessOnly)(_Out_ BOOL* losslessOnly) = 0;
    IFACEMETHOD(put_LosslessOnly)(_In_ BOOL losslessOnly) = 0;
    IFACEMETHOD(get_PreviewOnly)(_Out_ BOOL* previewOnly) = 0;
    IFACEMETHOD(put_PreviewOnly)(_In_ BOOL previewOnly) = 0;
    IFACEMETHOD(get_MaxWorkerThreadCount)(_Out_ UINT* maxThreadCount) = 0;
    IFACEMETHOD(put_MaxWorkerThreadCount)(_In_ UINT maxThreadCount) = 0;
    IFACEMETHOD(get_WorkerThreadCount)(_Out_ UINT* threadCount) = 0;
    IFACEMETHOD(put_WorkerThreadCount)(_In_ UINT threadCount) = 0;
    IFACEMETHOD(get_MinItemsPerWorkerThread)(_Out_ UINT* minItemsPerThread) = 0;
    IFACEMETHOD(put_MinItemsPerWorkerThread)(_In_ UINT minItemsPerThread) = 0;
    IFACEMETHOD(get_ItemsPerWorkerThread)(_Out_ UINT* numItemsPerThread) = 0;
    IFACEMETHOD(put_ItemsPerWorkerThread)(_In_ UINT numItemsPerThread) = 0;
};

interface __declspec(uuid("39EEEC7B-81B4-4719-90E1-37D70059724C")) IRotationUI : public IUnknown
{
public:
    IFACEMETHOD(Start)() = 0;
    IFACEMETHOD(Close)() = 0;
};
