#pragma once
#include "stdafx.h"
#include "RotationInterfaces.h"

// Helpers for enumerating shell items from a data object
HRESULT EnumerateDataObject(__in IDataObject* pdo, __in IRotationManager* prm);
HRESULT ParseEnumItems(_In_ IEnumShellItems *pesi, _In_ UINT depth, __in IRotationManager* prm);