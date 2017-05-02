#pragma once
#include "stdafx.h"

interface __declspec(uuid("698B7CF6-0AB5-4174-BC99-015AA7FBC7EA")) IRotationItem : public IUnknown
{
public:
    IFACEMETHOD(GetItem)(__deref_out IShellItem** ppsi) = 0;
    IFACEMETHOD(SetItem)(__in IShellItem* psi) = 0;
    IFACEMETHOD(GetResult)(__out HRESULT* phrResult) = 0;
    IFACEMETHOD(SetResult)(__in HRESULT hrResult) = 0;
    IFACEMETHOD(Rotate)() = 0;
};
