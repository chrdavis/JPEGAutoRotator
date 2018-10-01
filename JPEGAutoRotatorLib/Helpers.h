#pragma once
#include "stdafx.h"
#include "RotationInterfaces.h"

// Helpers for enumerating shell items from a data object
HRESULT EnumerateDataObject(_In_ IDataObject* pdo, _In_ IRotationManager* prm);

// Get the core count
UINT GetLogicalProcessorCount();