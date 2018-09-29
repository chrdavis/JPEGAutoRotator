#pragma once
#include "stdafx.h"
#include "RotationInterfaces.h"

// Helpers for enumerating shell items from a data object
HRESULT EnumerateDataObject(_In_ IDataObject* pdo, _In_ IRotationManager* prm, _In_ bool enumSubFolders);
HRESULT ParseEnumItems(_In_ IEnumShellItems* pesi, _In_ UINT depth, _In_ IRotationManager* prm, _In_ bool enumSubfolders);

// Helpers for enumerating via the file system apis (FindFirst/FindNext)
bool EnumeratePath(_In_ PCWSTR search, _In_ UINT depth, _In_ IRotationManager* prm, _In_ bool enumSubfolders);

bool IsJPEG(_In_ PCWSTR path);

// Get the core count
UINT GetLogicalProcessorCount();