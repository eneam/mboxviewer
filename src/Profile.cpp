#include "stdafx.h"

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
