#pragma once
#include "stdafx.h"

interface __declspec(uuid("FFB76FF1-DE4B-459A-A6C7-025392CFDA53")) IRotationItem : public IUnknown
{
public:

    IFACEMETHOD(GetPath)(__deref_out PWSTR* ppszPath) = 0;
    IFACEMETHOD(SetPath)(__in PCWSTR pszPath) = 0;
    IFACEMETHOD(GetResult)(__out HRESULT* phrResult) = 0;
    IFACEMETHOD(SetResult)(__in HRESULT hrResult) = 0;
    IFACEMETHOD(Rotate)() = 0;
};

interface __declspec(uuid("1FB14328-E681-4BC2-B947-00BAC0387A01")) IRotationManagerEvents : public IUnknown
{
public:
    IFACEMETHOD(OnAdded)(__in UINT uIndex) = 0;
    IFACEMETHOD(OnRotated)(__in UINT uIndex) = 0;
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
    IFACEMETHOD(AddItem)(__in IRotationItem* pri) = 0;
    IFACEMETHOD(GetItem)(__in UINT uIndex, __deref_out IRotationItem** ppri) = 0;
    IFACEMETHOD(GetItemCount)(__out UINT* puCount) = 0;
};

interface __declspec(uuid("39EEEC7B-81B4-4719-90E1-37D70059724C")) IRotationUI : public IUnknown
{
public:
    IFACEMETHOD(Start)() = 0;
    IFACEMETHOD(Close)() = 0;
};