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

#include "stdafx.h"
#include "Profile.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

#pragma once
//		prof._WriteProfileString( HKEY_LOCAL_MACHINE, "Software\\Microsoft\\Windows\\CurrentVersion\\Run", szAppName, filePath );
//		prof._DeleteProfileString(HKEY_LOCAL_MACHINE, "Software\\Microsoft\\Windows\\CurrentVersion\\Run", szAppName );

CString CProfile::_GetProfileString( HKEY hKey, LPCTSTR section, LPCTSTR key )
{
	DWORD	size = 4096;
	unsigned char	data[4096];
	data[0] = 0;
	HKEY	myKey;
	CString	result;
	LSTATUS res;
	CString	path = (char *)section;
	if (ERROR_SUCCESS == RegOpenKeyEx( hKey, (LPCTSTR)path, NULL, KEY_READ|KEY_QUERY_VALUE, &myKey ) )
	{
		res = RegQueryValueEx( myKey, (LPCTSTR)key,  NULL, NULL, data, &size );
		RegCloseKey( myKey );
		if( res != ERROR_SUCCESS)
			result = (char *)("");
		else
			result = (char *)data;
	} else
		result = (char *)data;
	return result;
}

int CProfile::_GetProfileInt( HKEY hKey, LPCTSTR section, LPCTSTR key )
{
	DWORD	size = sizeof(DWORD);
	HKEY	myKey;
	DWORD		result = 0;
	LSTATUS res;
	CString	path = (char *)section;
	if (ERROR_SUCCESS == RegOpenKeyEx( hKey, (LPCTSTR)path, NULL, KEY_READ|KEY_QUERY_VALUE, &myKey ) )
	{
		res = RegQueryValueEx( myKey, (LPCTSTR)key,  NULL, NULL, (BYTE *)&result, &size );
		RegCloseKey( myKey );
	}
	return (int)result;
}

BOOL CProfile::_GetProfileString(HKEY hKey, LPCTSTR section, LPCTSTR key, CString &str)
{
	DWORD	size = 4096;
	unsigned char	data[4096];
	data[0] = 0;
	HKEY	myKey;
	CString	result = (char *)("");

	LSTATUS res;
	CString	path = (char *)section;
	if (ERROR_SUCCESS == RegOpenKeyEx(hKey, (LPCTSTR)path, NULL, KEY_READ | KEY_QUERY_VALUE, &myKey))
	{
		res = RegQueryValueEx(myKey, (LPCTSTR)key, NULL, NULL, data, &size);
		RegCloseKey(myKey);
		if (res == ERROR_SUCCESS) {
			result = (char *)data;
			str = result;
			return TRUE;
		}
	}
	return FALSE;
}

BOOL CProfile::_GetProfileBinary(HKEY hKey, LPCTSTR section, LPCTSTR key, BYTE *lpData, DWORD &cbData)
{
	HKEY	myKey;
	CString	result = (char *)("");

	LSTATUS res;
	CString	path = (char *)section;
	if (ERROR_SUCCESS == RegOpenKeyEx(hKey, (LPCTSTR)path, NULL, KEY_READ | KEY_QUERY_VALUE, &myKey))
	{
		res = RegQueryValueEx(myKey, (LPCTSTR)key, NULL, NULL, lpData, &cbData);
		RegCloseKey(myKey);
		if (res == ERROR_SUCCESS) {
			return TRUE;
		}
	}
	return FALSE;
}

BOOL CProfile::_GetProfileInt(HKEY hKey, LPCTSTR section, LPCTSTR key, DWORD &intval)
{
	DWORD	size = sizeof(DWORD);
	HKEY	myKey;
	DWORD	result = 0;
	LSTATUS res;
	CString	path = (char *)section;
	if (ERROR_SUCCESS == RegOpenKeyEx(hKey, (LPCTSTR)path, NULL, KEY_READ | KEY_QUERY_VALUE, &myKey))
	{
		res = RegQueryValueEx(myKey, (LPCTSTR)key, NULL, NULL, (BYTE *)&result, &size);
		RegCloseKey(myKey);
		if (res == ERROR_SUCCESS) {
			intval = result;
			return TRUE;
		}
	}
	return FALSE;
}

BOOL CProfile::_GetProfileInt(HKEY hKey, LPCTSTR section, LPCTSTR key, int &intval)
{
	DWORD dwVal = 0;
	if (key == 0)
		return FALSE;
	BOOL ret = _GetProfileInt(hKey, section, key, dwVal);
	if (ret)
		intval = dwVal;
	return ret;
}

BOOL CProfile::_DeleteProfileString(HKEY hKey, LPCTSTR section, LPCTSTR key)
{
	HKEY	myKey;
	CString	path = (char *)section;
	if ( ERROR_SUCCESS == RegOpenKeyEx( hKey, 
							(LPCTSTR)path, 0, KEY_ALL_ACCESS, 
							&myKey) ) 
	{
		if (key)
			LSTATUS Res = RegDeleteValue( myKey, key );
		RegCloseKey( myKey );
		return TRUE;
	} else
		return FALSE;
}

BOOL CProfile::_WriteProfileString( HKEY hKey, LPCTSTR section, LPCTSTR key, CString &value )
{
	DWORD	dwDisposition;
	HKEY	myKey;
	CString	path = (char *)section;
	LSTATUS sts = RegCreateKeyEx(hKey,
		(LPCTSTR)path, 0, NULL,
		REG_OPTION_NON_VOLATILE,
		KEY_WRITE, NULL, &myKey,
		&dwDisposition);

		// ERROR_SUCCESS
	if (sts == ERROR_SUCCESS)
	{
		if (key)
		{
			LSTATUS sts = RegSetValueEx(myKey, key, 0, REG_SZ, (CONST BYTE*)(LPCTSTR)value, value.GetLength() + 1);
			if (sts != ERROR_SUCCESS)
			{
				DWORD err = GetLastError();
			}
		}
		RegCloseKey( myKey );
		return TRUE;
	} else
		return FALSE;
}

BOOL CProfile::_WriteProfileBinary(HKEY hKey, LPCTSTR section, LPCTSTR key, const BYTE *lpData, DWORD cbData)
{
	DWORD	dwDisposition;
	HKEY	myKey;
	CString	path = (char *)section;
	LSTATUS sts = RegCreateKeyEx(hKey,
		(LPCTSTR)path, 0, NULL,
		REG_OPTION_NON_VOLATILE,
		KEY_WRITE, NULL, &myKey,
		&dwDisposition);

	// ERROR_SUCCESS
	if (sts == ERROR_SUCCESS)
	{
		if (key)
		{
			LSTATUS sts = RegSetValueEx(myKey, key, 0, REG_BINARY, lpData, cbData);
			if (sts != ERROR_SUCCESS)
			{
				DWORD err = GetLastError();
			}
		}
		RegCloseKey(myKey);
		return TRUE;
	}
	else
		return FALSE;
}


BOOL CProfile::_WriteProfileInt( HKEY hKey, LPCTSTR section, LPCTSTR key, DWORD value )
{
	DWORD	dwDisposition;
	HKEY	myKey;
	CString	path = (char *)section;
	if ( ERROR_SUCCESS == RegCreateKeyEx( hKey, 
							(LPCTSTR)path, 0, NULL, 
							REG_OPTION_NON_VOLATILE, 
							KEY_WRITE, NULL, &myKey, 
							&dwDisposition ) ) 
	{
		if (key)
		{
			LSTATUS sts = RegSetValueEx(myKey, key, 0, REG_DWORD, (CONST BYTE*)&value, sizeof(value));
			if (sts != ERROR_SUCCESS)
			{
				DWORD err = GetLastError();
			}
		}
		RegCloseKey( myKey );
		return TRUE;
	} else
		return FALSE;
}

BOOL CProfile::_DeleteKey(HKEY hKey, LPCTSTR section, LPCTSTR key)
{
	HKEY	myKey;

	if (!key)
		return FALSE;

	CString	path = (char *)section;
	if (ERROR_SUCCESS == RegOpenKeyEx(hKey,
		(LPCTSTR)path, 0, KEY_SET_VALUE, &myKey))
	{

		LSTATUS sts = RegDeleteValue(myKey, key);
		if (sts != ERROR_SUCCESS)
		{
			DWORD err = GetLastError();
		}
		else
			; // return FALSE

		RegCloseKey(myKey);
		return TRUE;
	}
	else
		return TRUE;
}
