#pragma once
#include "stdafx.h"
#include "RotationInterfaces.h"

// Helpers for enumerating shell items from a data object
HRESULT EnumerateDataObject(_In_ IDataObject* pdo, _In_ IRotationManager* prm);
HRESULT ParseEnumItems(_In_ IEnumShellItems *pesi, _In_ UINT depth, _In_ IRotationManager* prm);

// Determining worker thread count
UINT GetLogicalProcessorCount();