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

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

#pragma once
//		prof._WriteProfileString( HKEY_LOCAL_MACHINE, "Software\\Microsoft\\Windows\\CurrentVersion\\Run", szAppName, filePath );
//		prof._DeleteProfileString(HKEY_LOCAL_MACHINE, "Software\\Microsoft\\Windows\\CurrentVersion\\Run", szAppName );
/*
CProfile::CProfile()
{
}

CProfile::CProfile(CString path)
{
	m_regAppKey = path;
}

CString CProfile::_GetProfileString( LPCTSTR section, LPCTSTR key )
{
	DWORD	size = 512;
	unsigned char	data[512];
	data[0] = 0;
	HKEY	myKey;
	CString	result;
	long res;
	CString	path = m_regAppKey;
	if( section && lstrlen(section) ) {
		path += "\\";
		path += (char *)section;
	}
	if( ! RegOpenKeyEx( HKEY_CURRENT_USER, (LPCTSTR)path,
				NULL, KEY_READ|KEY_QUERY_VALUE, &myKey ) ) {
		res = RegQueryValueEx( myKey, (LPCTSTR)key, 
				NULL, NULL, data, &size );
		RegCloseKey( myKey );
		if( res )
			result = "";
		else
			result = (char *)data;
	} else
		result = (char *)data;
	return result;
}

int CProfile::_GetProfileInt( LPCTSTR section, LPCTSTR key )
{
	DWORD	size = sizeof(DWORD);
	HKEY	myKey;
	DWORD		result = 0;
	CString	path = m_regAppKey;
	if( section && lstrlen(section) ) {
		path += "\\";
		path += (char *)section;
	}
	if( ! RegOpenKeyEx( HKEY_CURRENT_USER, (LPCTSTR)path,
				NULL, KEY_READ|KEY_QUERY_VALUE, &myKey ) ) {
		RegQueryValueEx( myKey, (LPCTSTR)key, 
				NULL, NULL, (BYTE *)&result, &size );
		RegCloseKey( myKey );
	}
	return (int)result;
}
*/
CString CProfile::_GetProfileString( HKEY hKey, LPCTSTR section, LPCTSTR key )
{
	DWORD	size = 4096;
	unsigned char	data[4096];
	data[0] = 0;
	HKEY	myKey;
	CString	result;
	long res;
	CString	path = (char *)section;
	if( ! RegOpenKeyEx( hKey, (LPCTSTR)path,
				NULL, KEY_READ|KEY_QUERY_VALUE, &myKey ) ) {
		res = RegQueryValueEx( myKey, (LPCTSTR)key, 
				NULL, NULL, data, &size );
		RegCloseKey( myKey );
		if( res )
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
	CString	path = (char *)section;
	if( ! RegOpenKeyEx( hKey, (LPCTSTR)path,
				NULL, KEY_READ|KEY_QUERY_VALUE, &myKey ) ) {
		RegQueryValueEx( myKey, (LPCTSTR)key, 
				NULL, NULL, (BYTE *)&result, &size );
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
	int l = result.GetLength();
	if (result.IsEmpty()) {
		int deb = 1;
	}
	long res;
	CString	path = (char *)section;
	if (!RegOpenKeyEx(hKey, (LPCTSTR)path,
		NULL, KEY_READ | KEY_QUERY_VALUE, &myKey)) {
		res = RegQueryValueEx(myKey, (LPCTSTR)key,
			NULL, NULL, data, &size);
		RegCloseKey(myKey);
		if (res == ERROR_SUCCESS) {
			result = (char *)data;
			str = result;
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
	long res;
	CString	path = (char *)section;
	if (!RegOpenKeyEx(hKey, (LPCTSTR)path,
		NULL, KEY_READ | KEY_QUERY_VALUE, &myKey)) {
		res = RegQueryValueEx(myKey, (LPCTSTR)key,
			NULL, NULL, (BYTE *)&result, &size);
		RegCloseKey(myKey);
		if (res == ERROR_SUCCESS) {
			intval = result;
			return TRUE;
		}
	}
	return FALSE;
}

/*
BOOL CProfile::_WriteProfileString( LPCTSTR section, LPCTSTR key, CString &value )
{
	DWORD	dwDisposition;
	HKEY	myKey;
	CString	path = m_regAppKey;
	if( section && lstrlen(section) ) {
		path += "\\";
		path += (char *)section;
	}
	if( ERROR_SUCCESS == RegCreateKeyEx( HKEY_CURRENT_USER, 
							(LPCTSTR)path, 0, NULL, 
							REG_OPTION_NON_VOLATILE, 
							KEY_WRITE, NULL, &myKey, 
							&dwDisposition ) ) {
		RegSetValueEx ( myKey,key,0,REG_SZ,(CONST BYTE*)(LPCTSTR)value,value.GetLength()+1);
		RegCloseKey( myKey );
		return TRUE;
	} else
		return FALSE;
}
*/
BOOL CProfile::_DeleteProfileString(HKEY hKey, LPCTSTR section, LPCTSTR key)
{
	HKEY	myKey;
	CString	path = (char *)section;
	if( ERROR_SUCCESS == RegOpenKeyEx( hKey, 
							(LPCTSTR)path, 0, KEY_ALL_ACCESS, 
							&myKey) ) {
		long Res = RegDeleteValue( myKey, key );
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
	if( ERROR_SUCCESS == RegCreateKeyEx( hKey, 
							(LPCTSTR)path, 0, NULL, 
							REG_OPTION_NON_VOLATILE, 
							KEY_WRITE, NULL, &myKey, 
							&dwDisposition ) ) {
		RegSetValueEx ( myKey,key,0,REG_SZ,(CONST BYTE*)(LPCTSTR)value,value.GetLength()+1);
		RegCloseKey( myKey );
		return TRUE;
	} else
		return FALSE;
}

/*
BOOL CProfile::_WriteProfileInt( LPCTSTR section, LPCTSTR key, DWORD value )
{
	DWORD	dwDisposition;
	HKEY	myKey;
	CString	path = m_regAppKey;
	if( section && lstrlen(section) ) {
		path += "\\";
		path += (char *)section;
	}
	if( ERROR_SUCCESS == RegCreateKeyEx( HKEY_CURRENT_USER, 
							(LPCTSTR)path, 0, NULL, 
							REG_OPTION_NON_VOLATILE, 
							KEY_WRITE, NULL, &myKey, 
							&dwDisposition ) ) {
		RegSetValueEx ( myKey,key,0,REG_DWORD,(CONST BYTE*)&value,sizeof(value));
		RegCloseKey( myKey );
		return TRUE;
	} else
		return FALSE;
}
*/
BOOL CProfile::_WriteProfileInt( HKEY hKey, LPCTSTR section, LPCTSTR key, DWORD value )
{
	DWORD	dwDisposition;
	HKEY	myKey;
	CString	path = (char *)section;
	if( ERROR_SUCCESS == RegCreateKeyEx( hKey, 
							(LPCTSTR)path, 0, NULL, 
							REG_OPTION_NON_VOLATILE, 
							KEY_WRITE, NULL, &myKey, 
							&dwDisposition ) ) {
		RegSetValueEx ( myKey,key,0,REG_DWORD,(CONST BYTE*)&value,sizeof(value));
		RegCloseKey( myKey );
		return TRUE;
	} else
		return FALSE;
}
