//
//////////////////////////////////////////////////////////////////
//
//  Windows Mbox Viewer is a free tool to view, search and print mbox mail archives.
//
// Source code and executable can be downloaded from
//  https://sourceforge.net/projects/mbox-viewer/  and
//  https://github.com/eneam/mboxviewer
//
//  Copyright(C) 2019  Enea Mansutti, Zbigniew Minciel
//
//  This program is free software; you can redistribute it and/or modify
//  it under the terms of the version 3 of GNU Affero General Public License
//  as published by the Free Software Foundation; 
//
//  This program is distributed in the hope that it will be useful,
//  but WITHOUT ANY WARRANTY; without even the implied warranty of
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU
//  Library General Public License for more details.
//
//  You should have received a copy of the GNU Library General Public
//  License along with this program; if not, write to the
//  Free Software Foundation, Inc., 51 Franklin St, Fifth Floor,
//  Boston, MA  02110 - 1301, USA.
//
//////////////////////////////////////////////////////////////////
//

#pragma once

#include "stdafx.h"
#include "Profile.h"
#include "FileUtils.h"
#include "SimpleTree.h"
#include "MainFrm.h"
#include "mboxview.h"


#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif


BOOL CProfile::m_registry = TRUE;
CString CProfile::m_configFilePath;


BOOL CProfile::DetermineConfigurationType()
{
	CString configFilePath;

	CString configFileName = L"MBoxViewer.config";

	CString processNamePath;
	CmboxviewApp::GetProcessPath(processNamePath);

	CString processFolderPath;
	FileUtils::GetFolderPath(processNamePath, processFolderPath);

	CString configFolderPath = processFolderPath + L"Config\\";
	if (FileUtils::PathDirExists(configFolderPath))
	{
		configFilePath = configFolderPath + configFileName;
		if (!FileUtils::PathFileExist(configFilePath))
			configFilePath.Empty();
		int deb = 1;
	}
	if (configFilePath.IsEmpty())
	{
		CString localAppFolder = FileUtils::GetMboxviewLocalAppPath();
		configFolderPath = localAppFolder + L"UMBoxViewer\\Config";
		if (FileUtils::PathDirExists(configFolderPath))
		{
			configFilePath = configFolderPath + configFileName;
			if (!FileUtils::PathFileExist(configFilePath))
				configFilePath.Empty();
			int deb = 1;
		}
	}
	if (configFilePath.IsEmpty())
	{
		CProfile::m_registry = TRUE;
		CProfile::m_configFilePath.Empty();
		return FALSE;
	}
	else
	{
		CProfile::m_registry = FALSE;
		CProfile::m_configFilePath = configFilePath;
		return TRUE;
	}
}


ConfigTree* CProfile::GetConfigTree()
{
	if (CMainFrame::m_configTree == 0)
		CMainFrame::m_configTree = new ConfigTree(CString(L"ROOT"));

	_ASSERTE(CMainFrame::m_configTree != 0);

	return CMainFrame::m_configTree;
}

BOOL CProfile::GetFileConfigSection(LPCWSTR registrySection, CString& fileSection)
{
	fileSection.Empty();
	CString registrySectionStr = registrySection;
	int pos = registrySectionStr.Find(L"UMBoxViewer");

	if (pos >= 0)
	{
		fileSection = registrySectionStr.Mid(pos);
		return TRUE;
	}
	else
		return FALSE;
}

BOOL CProfile::IsRegistryConfig()
{
	if (m_registry)
		return TRUE;
	else
	{
		return FALSE;
	}
}

// TODO: Define dedicated functions for non UMBoxViewer configurations. Only few needed
BOOL CProfile::IsRegistryConfig(LPCWSTR registrySection, CString& fileSection)
{
	CString configSection;
	BOOL isTreeSection = GetFileConfigSection(registrySection, fileSection);
	if (m_registry || !isTreeSection)
		return TRUE;
	else
		return FALSE;
}

// Fix return from function to avoid multiple RegCloseKey(myKey);  // FIXMEFIXME

//		prof._WriteProfileString( HKEY_LOCAL_MACHINE, "Software\\Microsoft\\Windows\\CurrentVersion\\Run", szAppName, filePath );
//		prof._DeleteProfileString(HKEY_LOCAL_MACHINE, "Software\\Microsoft\\Windows\\CurrentVersion\\Run", szAppName );


// Ambigious. May return empty string due non-existent key or empty key value
// Use BOOL CProfile::_GetProfileString(HKEY hKey, LPCWSTR section, LPCWSTR key, CString& str) instead

CString CProfile::_GetProfileString(HKEY hKey, LPCWSTR section, LPCWSTR key)
{
	CString ret;
	CString fileSection;
	if (IsRegistryConfig(section, fileSection))
	{
		ret = CProfile::_GetProfileString_registry(hKey, section, key);
	}
	else
	{
		ConfigTree* config = CProfile::GetConfigTree();
		ret = config->_GetProfileString(hKey, section, key);
	}
	return ret;
}

CString CProfile::_GetProfileString_registry(HKEY hKey, LPCWSTR section, LPCWSTR key)
{
	DWORD	size = 4096;
	BYTE	data[4096];
	data[0] = 0;
	HKEY	myKey;
	CString	result;
	LSTATUS res;

	LSTATUS sts = RegOpenKeyEx(hKey, section, NULL, KEY_READ | KEY_QUERY_VALUE, &myKey);

	if (sts == ERROR_SUCCESS)
	{
		res = RegQueryValueEx(myKey, key, NULL, NULL, data, &size);
		if (res == ERROR_SUCCESS)
		{
			_ASSERTE((data[size - 1] == 0) && (data[size - 2] == 0));
			result = (LPCWSTR)data;
			RegCloseKey(myKey);
		}
		else
		{
			DWORD err = GetLastError();
			CString errText = FileUtils::GetLastErrorAsString(res);
			RegCloseKey(myKey);
		}
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString(sts);
	}
	return result;
}

int CProfile::_GetProfileInt(HKEY hKey, LPCWSTR section, LPCWSTR key)
{
	int ret;
	CString fileSection;
	if (IsRegistryConfig(section, fileSection))
	{
		ret = CProfile::_GetProfileInt_registry(hKey, section, key);
	}
	else
	{
		ConfigTree* config = CProfile::GetConfigTree();
		ret = config->_GetProfileInt(hKey, section, key);
	}
	return ret;
}

int CProfile::_GetProfileInt_registry(HKEY hKey, LPCWSTR section, LPCWSTR key)
{
	DWORD	size = sizeof(DWORD);
	HKEY	myKey;
	DWORD	result = 0;
	LSTATUS res;

	LSTATUS sts = RegOpenKeyEx(hKey, section, NULL, KEY_READ | KEY_QUERY_VALUE, &myKey);

	if (sts == ERROR_SUCCESS)
	{
		res = RegQueryValueEx(myKey, key, NULL, NULL, (BYTE*)&result, &size);
		if (res != ERROR_SUCCESS)
		{
			DWORD err = GetLastError();
			CString errText = FileUtils::GetLastErrorAsString(res);
		}
		RegCloseKey(myKey);
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString(sts);
	}
	return (int)result;
}

BOOL CProfile::_GetProfileString(HKEY hKey, LPCWSTR section, LPCWSTR key, CString& str)
{
	BOOL ret = TRUE;
	CString fileSection;
	if (IsRegistryConfig(section, fileSection))
	{
		ret = CProfile::_GetProfileString_registry(hKey, section, key, str);
	}
	else
	{
		ConfigTree* config = CProfile::GetConfigTree();
		ret = config->_GetProfileString(hKey, section, key, str);
	}
	return ret;
}

BOOL CProfile::_GetProfileString_registry(HKEY hKey, LPCWSTR section, LPCWSTR key, CString& str)
{
	DWORD	size = 4096;
	BYTE	data[4096];
	data[0] = 0;
	HKEY	myKey;
	CString	result;

	LSTATUS res;

	LSTATUS sts = RegOpenKeyEx(hKey, section, NULL, KEY_READ | KEY_QUERY_VALUE, &myKey);

	if (sts == ERROR_SUCCESS)
	{
		res = RegQueryValueEx(myKey, key, NULL, NULL, data, &size);

		if (res == ERROR_SUCCESS)
		{
			_ASSERTE((data[size - 1] == 0) && (data[size - 2] == 0));

			result = (LPCWSTR)data;
			str = result;
			RegCloseKey(myKey);
			return TRUE;
		}
		else
		{
			DWORD err = GetLastError();
			CString errText = FileUtils::GetLastErrorAsString(res);
			RegCloseKey(myKey);
			return FALSE;
		}
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString(sts);
	}
	return FALSE;
}

BOOL CProfile::_GetProfileBinary(HKEY hKey, LPCWSTR section, LPCWSTR key, BYTE* lpData, DWORD& cbData)
{
	BOOL ret = TRUE;
	CString fileSection;
	if (IsRegistryConfig(section, fileSection))
	{
		ret = CProfile::_GetProfileBinary_registry(hKey, section, key, lpData, cbData);
	}
	else
	{
		ConfigTree* config = CProfile::GetConfigTree();
		ret = config->_GetProfileBinary(hKey, section, key, lpData, cbData);
	}
	return ret;
}

BOOL CProfile::_GetProfileBinary_registry(HKEY hKey, LPCWSTR section, LPCWSTR key, BYTE* lpData, DWORD& cbData)
{
	HKEY	myKey;

	LSTATUS res;
	LSTATUS sts = RegOpenKeyEx(hKey, section, NULL, KEY_READ | KEY_QUERY_VALUE, &myKey);

	if (sts == ERROR_SUCCESS)
	{
		res = RegQueryValueEx(myKey, key, NULL, NULL, lpData, &cbData);

		if (res == ERROR_SUCCESS) {
			RegCloseKey(myKey);
			return TRUE;
		}
		else
		{
			DWORD err = GetLastError();
			CString errText = FileUtils::GetLastErrorAsString(res);
			RegCloseKey(myKey);
			return FALSE;
		}
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString(sts);
	}
	return FALSE;
}

BOOL CProfile::_GetProfileInt(HKEY hKey, LPCWSTR section, LPCWSTR key, DWORD& intval)
{
	BOOL ret =  TRUE;
	CString fileSection;
	if (IsRegistryConfig(section, fileSection))
	{
		ret = CProfile::_GetProfileInt_registry(hKey, section, key, intval);
	}
	else
	{
		ConfigTree* config = CProfile::GetConfigTree();
		ret = config->_GetProfileInt(hKey, section, key, intval);
	}
	return ret;
}

BOOL CProfile::_GetProfileInt_registry(HKEY hKey, LPCWSTR section, LPCWSTR key, DWORD& intval)
{
	DWORD	size = sizeof(DWORD);
	HKEY	myKey;
	DWORD	result = 0;
	LSTATUS res;

	LSTATUS sts = RegOpenKeyEx(hKey, section, NULL, KEY_READ | KEY_QUERY_VALUE, &myKey);

	if (sts == ERROR_SUCCESS)
	{
		if (key)
		{
			res = RegQueryValueEx(myKey, key, NULL, NULL, (BYTE*)&result, &size);

			if (res == ERROR_SUCCESS) {
				intval = result;
				RegCloseKey(myKey);
				return TRUE;
			}
			else
			{
				DWORD err = GetLastError();
				CString errText = FileUtils::GetLastErrorAsString(res);
				RegCloseKey(myKey);
				return FALSE;
			}
		}
		else
		{
			RegCloseKey(myKey);
			return TRUE;  // FIXMEFIXME
		}
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString(sts);
	}
	return FALSE;
}


BOOL CProfile::_GetProfileInt(HKEY hKey, LPCWSTR section, LPCWSTR key, int& intval)
{
	BOOL ret = TRUE;
	CString fileSection;
	if (IsRegistryConfig(section, fileSection))
	{
		ret = CProfile::_GetProfileInt_registry(hKey, section, key, intval);
	}
	else
	{
		ConfigTree* config = CProfile::GetConfigTree();
		ret = config->_GetProfileInt(hKey, section, key, intval);
	}
	return ret;
}

BOOL CProfile::_GetProfileInt_registry(HKEY hKey, LPCWSTR section, LPCWSTR key, int& intval)
{
	DWORD dwVal = 0;
	if (key == 0)
		return FALSE;
	BOOL ret = CProfile::_GetProfileInt(hKey, section, key, dwVal);
	if (ret)
		intval = dwVal;
	return ret;
}

BOOL CProfile::_DeleteProfileString(HKEY hKey, LPCWSTR section, LPCWSTR key)
{
	BOOL ret = TRUE;
	CString fileSection;
	if (IsRegistryConfig(section, fileSection))
	{
		ret = CProfile::_DeleteProfileString_registry(hKey, section, key);
	}
	else
	{
		ConfigTree* config = CProfile::GetConfigTree();
		ret = config->_DeleteProfileString(hKey, section, key);
	}

	return ret;
}

BOOL CProfile::_DeleteProfileString_registry(HKEY hKey, LPCWSTR section, LPCWSTR key)
{
	HKEY	myKey;
	LSTATUS res;

	LSTATUS sts = RegOpenKeyEx(hKey, section, 0, KEY_ALL_ACCESS, &myKey);

	if (sts == ERROR_SUCCESS)
	{
		if (key)
			res = RegDeleteValue(myKey, key);

		if (res != ERROR_SUCCESS)
		{
			DWORD err = GetLastError();
			CString errText = FileUtils::GetLastErrorAsString(res);
		}
		RegCloseKey(myKey);
		return TRUE;  // ???
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString(sts);
		return FALSE;
	}
}

BOOL CProfile::_WriteProfileString(HKEY hKey, LPCWSTR section, LPCWSTR key, CString& value)
{
	BOOL ret = TRUE;
	CString fileSection;
	if (IsRegistryConfig(section, fileSection))
	{
		ret = CProfile::_WriteProfileString_registry(hKey, section, key, value);
	}
	else
	{
		ConfigTree* config = CProfile::GetConfigTree();
		ret = config->_WriteProfileString(hKey, section, key, value);
	}
	return ret;
}

BOOL CProfile::_WriteProfileString_registry(HKEY hKey, LPCWSTR section, LPCWSTR key, CString& value)
{
	DWORD	dwDisposition;
	HKEY	myKey;

	LSTATUS sts = RegCreateKeyEx(hKey,
		section, 0, NULL,
		REG_OPTION_NON_VOLATILE,
		KEY_WRITE, NULL, &myKey,
		&dwDisposition);

	// ERROR_SUCCESS
	if (sts == ERROR_SUCCESS)
	{
		if (key)
		{
			LSTATUS stsSet = RegSetValueEx(myKey, key, 0, REG_SZ, (CONST BYTE*)(LPCWSTR)value, value.GetLength() * 2 + 2);

			if (stsSet != ERROR_SUCCESS)
			{
				DWORD err = GetLastError();
				CString errText = FileUtils::GetLastErrorAsString(stsSet);
				RegCloseKey(myKey);
				return FALSE;
			}
		}
		RegCloseKey(myKey);
		return TRUE;
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString(sts);
		return FALSE;
	}
}

BOOL CProfile::_WriteProfileBinary(HKEY hKey, LPCWSTR section, LPCWSTR key, const BYTE* lpData, DWORD cbData)
{
	BOOL ret = TRUE;
	CString fileSection;
	if (IsRegistryConfig(section, fileSection))
	{
		ret = CProfile::_WriteProfileBinary_registry(hKey, section, key, lpData, cbData);
	}
	else
	{
		ConfigTree* config = CProfile::GetConfigTree();
		ret = config->_WriteProfileBinary(hKey, section, key, lpData, cbData);
	}
	return ret;
}

BOOL CProfile::_WriteProfileBinary_registry(HKEY hKey, LPCWSTR section, LPCWSTR key, const BYTE* lpData, DWORD cbData)
{
	DWORD	dwDisposition;
	HKEY	myKey;

	LSTATUS sts = RegCreateKeyEx(hKey,
		section, 0, NULL,
		REG_OPTION_NON_VOLATILE,
		KEY_WRITE, NULL, &myKey,
		&dwDisposition);

	// ERROR_SUCCESS
	if (sts == ERROR_SUCCESS)
	{
		if (key)
		{
			LSTATUS stsSet = RegSetValueEx(myKey, key, 0, REG_BINARY, lpData, cbData);
			if (stsSet != ERROR_SUCCESS)
			{
				DWORD err = GetLastError();
				CString errText = FileUtils::GetLastErrorAsString(stsSet);
				RegCloseKey(myKey);
				return FALSE;
			}
		}
		RegCloseKey(myKey);
		return TRUE;
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString(sts);
	}
	return FALSE;
}

BOOL CProfile::_WriteProfileInt(HKEY hKey, LPCWSTR section, LPCWSTR key, DWORD value)
{
	BOOL ret = TRUE;
	CString fileSection;
	if (IsRegistryConfig(section, fileSection))
	{
		ret = CProfile::_WriteProfileInt_registry(hKey, section, key, value);
	}
	else
	{
		ConfigTree* config = CProfile::GetConfigTree();
		ret = config->_WriteProfileInt(hKey, section, key, value);
	}
	return ret;
}

BOOL CProfile::_WriteProfileInt_registry(HKEY hKey, LPCWSTR section, LPCWSTR key, DWORD value)
{
	DWORD	dwDisposition;
	HKEY	myKey;

	LSTATUS sts = RegCreateKeyEx(hKey,
		section, 0, NULL,
		REG_OPTION_NON_VOLATILE,
		KEY_WRITE, NULL, &myKey,
		&dwDisposition);

	if (sts == ERROR_SUCCESS)
	{
		if (key)
		{
			LSTATUS stsSet = RegSetValueEx(myKey, key, 0, REG_DWORD, (CONST BYTE*) & value, sizeof(value));
			if (stsSet != ERROR_SUCCESS)
			{
				DWORD err = GetLastError();
				CString errText = FileUtils::GetLastErrorAsString(stsSet);
				RegCloseKey(myKey);
				return FALSE;
			}
		}
		RegCloseKey(myKey);
		return TRUE;
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString(sts);
	}
	return FALSE;
}
// Deletes a subkey and its values from the specified platform-specific view of the registry.
// The subkey to be deleted must not have subkeys. 
// To delete a key and all its subkeys, you need to enumerate the subkeys and delete them individually.
// To delete keys recursively, use the RegDeleteTree or SHDeleteKey function.
// 

BOOL CProfile::_DeleteValue(HKEY hKey, LPCWSTR section, LPCWSTR key)
{
	BOOL ret = TRUE;
	CString fileSection;
	if (IsRegistryConfig(section, fileSection))
	{
		ret = CProfile::_DeleteValue_registry(hKey, section, key);
	}
	else
	{
		ConfigTree* config = CProfile::GetConfigTree();
		ret = config->_DeleteValue(hKey, section, key);
	}
	return ret;
}
BOOL CProfile::_DeleteValue_registry(HKEY hKey, LPCWSTR section, LPCWSTR key)
{
	HKEY	myKey;
	if (!key)
		return FALSE;

	LSTATUS sts = RegOpenKeyEx(hKey,
		section, 0, KEY_SET_VALUE| KEY_WRITE, &myKey);

	if (sts == ERROR_SUCCESS)
	{
		LSTATUS stsDel = RegDeleteValue(myKey, key);

		if (stsDel != ERROR_SUCCESS)
		{
			DWORD err = GetLastError();
			CString errText = FileUtils::GetLastErrorAsString(stsDel);
			//RegCloseKey(myKey);
			// return FALSE  // ???
		}
		RegCloseKey(myKey);
		return TRUE;
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString(sts);
		return TRUE;  // return FALSE;  // ??
	}
}

BOOL CProfile::_DeleteKey(HKEY hKey, LPCWSTR section, LPCWSTR key, BOOL recursive)
{
	BOOL ret = TRUE;
	CString fileSection;
	if (IsRegistryConfig(section, fileSection))
	{
		ret = CProfile::_DeleteKey_registry(hKey, section, key, recursive);
	}
	else
	{
		ConfigTree* config = CProfile::GetConfigTree();
		ret = config->_DeleteKey(hKey, section, key, recursive);
	}
	return ret;
}

BOOL CProfile::_DeleteKey_registry(HKEY hKey, LPCWSTR section, LPCWSTR key, BOOL recursive)
{
	HKEY	myKey;
	if (!key)
		return FALSE;

	REGSAM samDesired = KEY_WRITE|KEY_WOW64_32KEY;

	LSTATUS sts = RegOpenKeyEx(hKey, section, 0, samDesired, &myKey);

	if (sts == ERROR_SUCCESS)
	{
		if (recursive)
		{
			//LSTATUS stsDel = RegDeleteTree(myKey, key);  //  #if _WIN32_WINNT >= 0x0600

			LSTATUS stsDel = SHDeleteKey(myKey, key);

			if (stsDel != ERROR_SUCCESS)
			{
				DWORD err = GetLastError();
				CString errText = FileUtils::GetLastErrorAsString(stsDel);
				//RegCloseKey(myKey);
				// return FALSE  // ???
			}
		}
		else
		{
			//REGSAM samDesired = KEY_WOW64_32KEY;
			//LSTATUS stsDel = RegDeleteKeyEx(myKey, key, samDesired, 0);
			//LSTATUS stsDel = RegDeleteKey(myKey, key);
			// RegDeleteKey and RegDeleteKeyEx don't work for me
			// RegDeleteValue deleted value and key
			LSTATUS stsDel = RegDeleteValue(myKey, key);
			if (stsDel != ERROR_SUCCESS)
			{
				DWORD err = GetLastError();
				CString errText = FileUtils::GetLastErrorAsString(stsDel);
				//RegCloseKey(myKey);
				// return FALSE  // ???
			}
		}
		RegCloseKey(myKey);
		return TRUE;
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString(sts);
		return TRUE;  // return FALSE;  // ??
	}
}

LSTATUS CProfile::InitRegQueryInfoKeyParams(HKEY hKey, RegQueryInfoKeyParams& params)
{
	LSTATUS retCode = RegQueryInfoKey(
		hKey,                    // key handle 
		params.achClass,                // buffer for class name 
		&params.cchClassName,           // size of class string 
		NULL,                    // reserved 
		&params.cSubKeys,               // number of subkeys 
		&params.cbMaxSubKey,            // longest subkey size 
		&params.cchMaxClass,            // longest class string 
		&params.cValues,                // number of values for this key 
		&params.cchMaxValue,            // longest value name 
		&params.cbMaxValueData,         // longest value data 
		&params.cbSecurityDescriptor,   // security descriptor 
		&params.ftLastWriteTime);       // last write time

	_ASSERTE(retCode == ERROR_SUCCESS);
	if (retCode != ERROR_SUCCESS)
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString(retCode);
		int deb = 1;
	}

	return retCode;
}

// Enumerate the subkeys, until RegEnumKeyEx fails.
LSTATUS CProfile::EnumerateAllSubKeys(HKEY hKey, LPCWSTR section)
{
	HKEY myKey;
	LSTATUS retCode = RegOpenKeyEx(hKey, section, NULL, KEY_READ| KEY_ENUMERATE_SUB_KEYS, &myKey);
	if (ERROR_SUCCESS == retCode)
	{
		RegQueryInfoKeyParams params;

		LSTATUS retCode = CProfile::InitRegQueryInfoKeyParams(myKey, params);
		if (retCode == ERROR_SUCCESS)
		{
			TRACE(L"Number of subkeys: %d\n", params.cSubKeys);
			if (params.cSubKeys > 0)
			{
				int i;
				for (i = 0; i < (int)params.cSubKeys; i++)
				{
					params.cbName = MAX_KEY_NAME_LENGTH;
					retCode = RegEnumKeyEx(myKey, i,
						params.achKey,
						&params.cbName,
						NULL,
						NULL,
						NULL,
						&params.ftLastWriteTime);

					if (retCode == ERROR_SUCCESS)
					{
						TRACE(L"(%d) %s\n", i + 1, params.achKey);
					}
					else
					{
						DWORD err = GetLastError();
						CString errText = FileUtils::GetLastErrorAsString(retCode);
					}
				}
			}
		}
		else
		{
			DWORD err = GetLastError();
			CString errText = FileUtils::GetLastErrorAsString();
		}
		RegCloseKey(myKey);
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString(retCode);
	}
	return retCode;
}

// Enumerate the key data values in Registry. Not recursive
LSTATUS CProfile::EnumerateAllSubKeyValues(HKEY hKey, LPCWSTR section)
{
	wchar_t  achValue[MAX_KEY_VALUE_LENGTH];
	DWORD cchValue = MAX_KEY_VALUE_LENGTH;
	BYTE  data[MAX_KEY_VALUE_LENGTH];
	DWORD dataSize = MAX_KEY_VALUE_LENGTH;
	DWORD dataType;

	HKEY myKey;
	LSTATUS retCode = RegOpenKeyEx(hKey, section, NULL, KEY_READ, &myKey);
	if (ERROR_SUCCESS == retCode)
	{
		RegQueryInfoKeyParams params;

		retCode = CProfile::InitRegQueryInfoKeyParams(myKey, params);
		if (retCode == ERROR_SUCCESS)
		{
			TRACE(L"Number of values: %d\n", params.cValues);

			int i;
			for (i = 0, retCode = ERROR_SUCCESS; i < (int)params.cValues; i++)
			{
				cchValue = MAX_KEY_VALUE_LENGTH;
				achValue[0] = 0;
				dataSize = MAX_KEY_VALUE_LENGTH;
				data[0] = 0;
				retCode = RegEnumValue(myKey, i,
					achValue,
					&cchValue,
					NULL,
					&dataType,
					(BYTE*)&data[0],
					&dataSize);

				if (retCode == ERROR_SUCCESS)
				{
					DWORD dval;
					CString dataValue;
					achValue[cchValue] = 0;
					if (dataType == REG_DWORD)
					{
						memcpy(&dval, data, sizeof(DWORD));
						QWORD ddval = dval;
						dataValue.Format(L"%lld", ddval);
					}
					else if (dataType == REG_SZ)
					{
						memcpy(&dval, data, sizeof(DWORD));
						QWORD ddval = dval;
						dataValue.Empty();
						dataValue.Append((wchar_t*)data, dataSize);
					}
					else if (dataType == REG_BINARY)
					{
						int k;
						CString bval;
						for (k = 0; k < (int)dataSize; k++)
						{
							if (k == 0)
								bval.Format(L"%02x", data[k]);
							else
								bval.Format(L" %02x", data[k]);
							dataValue.Append(bval);
						}
					}
					data[dataSize] = 0;
					TRACE(L"(%d) %s \"%s\" \n", i + 1, &achValue[0], (LPCWSTR)dataValue);
				}
				else
				{
					DWORD err = GetLastError();
					CString errText = FileUtils::GetLastErrorAsString(retCode);
				}
			}
		}
		else
		{
			DWORD err = GetLastError();
			CString errText = FileUtils::GetLastErrorAsString(retCode);
		}
		RegCloseKey(myKey);
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString();
	}
	return retCode;
}

// Enumerate the key data values. 
LSTATUS CProfile::CopyKey(HKEY hKey, LPCWSTR section, LPCWSTR toSection)
{
	wchar_t  achValue[MAX_KEY_VALUE_LENGTH];
	DWORD cchValue = MAX_KEY_VALUE_LENGTH;
	BYTE  data[MAX_KEY_VALUE_LENGTH];
	DWORD dataSize = MAX_KEY_VALUE_LENGTH;
	DWORD dataType;

	HKEY myKey;
	LSTATUS retCode = RegOpenKeyEx(hKey, section, NULL, KEY_READ, &myKey);
	if (ERROR_SUCCESS == retCode)
	{
		RegQueryInfoKeyParams params;

		retCode = CProfile::InitRegQueryInfoKeyParams(myKey, params);
		if (retCode == ERROR_SUCCESS)
		{
			TRACE(L"Number of values: %d\n", params.cValues);

			int i;
			for (i = 0, retCode = ERROR_SUCCESS; i < (int)params.cValues; i++)
			{
				cchValue = MAX_KEY_VALUE_LENGTH;
				achValue[0] = 0;
				dataSize = MAX_KEY_VALUE_LENGTH;
				data[0] = 0;
				retCode = RegEnumValue(myKey, i,
					achValue,
					&cchValue,
					NULL,
					&dataType,
					(BYTE*)&data[0],
					&dataSize);

				if (retCode == ERROR_SUCCESS)
				{
					DWORD dval;
					CString dataValue;
					achValue[cchValue] = 0;
					if (dataType == REG_DWORD)
					{
						memcpy(&dval, data, sizeof(DWORD));
						QWORD ddval = dval;
						dataValue.Format(L"%lld", ddval);

						CProfile::_WriteProfileInt(hKey, (LPCWSTR)toSection, &achValue[0], dval);
					}
					else if (dataType == REG_BINARY)
					{
						CProfile::_WriteProfileBinary(hKey, (LPCWSTR)toSection, &achValue[0], (BYTE*)&data[0], dataSize);
					}
					else if (dataType == REG_SZ)
					{
						dataValue.Empty();
						dataValue.Append((wchar_t*)&data[0], dataSize);
						CProfile::_WriteProfileString(hKey, (LPCWSTR)toSection, &achValue[0], dataValue);
					}
					data[dataSize] = 0;
					TRACE(L"(%d) %s \"%s\" \n", i + 1, &achValue[0], (LPCWSTR)dataValue);
				}
				else
				{
					DWORD err = GetLastError();
					CString errText = FileUtils::GetLastErrorAsString(retCode);
				}
			}
		}
		else
		{
			DWORD err = GetLastError();
			CString errText = FileUtils::GetLastErrorAsString();
		}
		RegCloseKey(myKey);
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString(retCode);
	}
	return retCode;
}

// Iterate subkeys and copy subkeys and data values
LSTATUS CProfile::CopySubKeys(HKEY hKey, LPCWSTR section, LPCWSTR toSection)
{
	HKEY myKey;
	LSTATUS retCode = RegOpenKeyEx(hKey, section, NULL, KEY_READ | KEY_ENUMERATE_SUB_KEYS, &myKey);
	if (ERROR_SUCCESS == retCode)
	{
		RegQueryInfoKeyParams params;

		LSTATUS retCode = CProfile::InitRegQueryInfoKeyParams(myKey, params);
		if (retCode == ERROR_SUCCESS)
		{
			TRACE(L"Number of subkeys: %d\n", params.cSubKeys);

			int i;
			for (i = 0; i < (int)params.cSubKeys; i++)
			{
				params.cbName = MAX_KEY_NAME_LENGTH;
				retCode = RegEnumKeyEx(myKey, i,
					params.achKey,
					&params.cbName,
					NULL,
					NULL,
					NULL,
					&params.ftLastWriteTime);

				if (retCode == ERROR_SUCCESS)
				{
					TRACE(L"(%d) %s\n", i + 1, params.achKey);
					CString subKeySection = CString(section) + L"\\" + params.achKey;
					CString subKeyToSection = CString(toSection) + L"\\" + params.achKey;
					LSTATUS retCode = CProfile::CopyKey(hKey, (LPCWSTR)subKeySection, (LPCWSTR)subKeyToSection);
					int deb = 1;
				}
				else
				{
					DWORD err = GetLastError();
					CString errText = FileUtils::GetLastErrorAsString(retCode);
				}
			}
		}
		else
		{
			DWORD err = GetLastError();
			CString errText = FileUtils::GetLastErrorAsString(retCode);
		}
		RegCloseKey(myKey);
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString(retCode);
	}
	return retCode;
}

// Not very efficient but this done during MBox Viewer version migration only
LSTATUS CProfile::CopyKeyValueList(HKEY hKey, LPCWSTR fromSection, LPCWSTR toSection, KeyFromToTable &arr)
{
	CString sval;
	DWORD dval;
	BYTE bval[MAX_KEY_VALUE_LENGTH];
	BOOL retval;
	int i;
	for (i = 0; i < arr.GetCount(); i++)
	{
		RegKeyFromToInfo &info = arr.GetAt(i);
		if (info.m_type == REG_SZ)
		{
			retval = CProfile::_GetProfileString(hKey, (LPCWSTR)fromSection, (LPCWSTR)info.m_from, sval);
			if (retval)
			{
				if (!info.m_to.IsEmpty())
					CProfile::_WriteProfileString(hKey, (LPCWSTR)toSection, (LPCWSTR)info.m_to, sval);
				else
					CProfile::_WriteProfileString(hKey, (LPCWSTR)toSection, (LPCWSTR)info.m_from, sval);
			}
		}
		else if (info.m_type == REG_DWORD)
		{
			retval = CProfile::_GetProfileInt(hKey, (LPCWSTR)fromSection, (LPCWSTR)info.m_from, dval);
			if (retval)
			{
				if (!info.m_to.IsEmpty())
					CProfile::_WriteProfileInt(hKey, (LPCWSTR)toSection, (LPCWSTR)info.m_to, dval);
				else
					CProfile::_WriteProfileInt(hKey, (LPCWSTR)toSection, (LPCWSTR)info.m_from, dval);
			}
		}
		else if (info.m_type == REG_BINARY)
		{
			DWORD blen = 0;
			bval[0] = 0;
			retval = CProfile::_GetProfileBinary(hKey, (LPCWSTR)fromSection, (LPCWSTR)info.m_from, bval, blen);
			if (retval)
			{
				if (!info.m_to.IsEmpty())
					CProfile::_WriteProfileBinary(hKey, (LPCWSTR)toSection, (LPCWSTR)info.m_to, bval, blen);
				else
					CProfile::_WriteProfileBinary(hKey, (LPCWSTR)toSection, (LPCWSTR)info.m_from, bval, blen);
			}
		}
	}
	return 0;
}

// Check if Key exists in Registry

BOOL CProfile::CheckIfKeyExists(HKEY hKey, LPCWSTR section)
{
	HKEY	myKey = 0;
	LSTATUS retCode = RegOpenKeyEx(hKey, section, NULL, KEY_READ | KEY_QUERY_VALUE, &myKey);

	if (ERROR_SUCCESS == retCode)
	{
		// No need to close the key
		return TRUE;
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString(retCode);
		return FALSE;
	}
}

