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
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
//  Library General Public License for more details.
//
//  You should have received a copy of the GNU Library General Public
//  License along with this program; if not, write to the
//  Free Software Foundation, Inc., 51 Franklin St, Fifth Floor,
//  Boston, MA  02110 - 1301, USA.
//
//////////////////////////////////////////////////////////////////
//

#if !defined(_PROFILE_)
#define _PROFILE_

#pragma once

#define MAX_KEY_NAME_LENGTH 255
#define MAX_KEY_VALUE_LENGTH 16383

class RegKeyFromToInfo
{
public:
	UINT m_type;
	CString m_from;
	CStringW m_to;
};

typedef CArray<RegKeyFromToInfo> KeyFromToTable;

struct RegQueryInfoKeyParams
{
	RegQueryInfoKeyParams() { SetDflts(); }
	TCHAR    achKey[MAX_KEY_NAME_LENGTH +1];   // buffer for subkey name
	DWORD    cbName;                   // size of name string 
	TCHAR    achClass[MAX_PATH+1];   // buffer for class name 
	DWORD    cchClassName;         // size of class string 
	DWORD    cSubKeys;             // number of subkeys 
	DWORD    cbMaxSubKey;          // longest subkey size 
	DWORD    cchMaxClass;          // longest class string 
	DWORD    cValues;              // number of values for key 
	DWORD    cchMaxValue;          // longest value name 
	DWORD    cbMaxValueData;       // longest value data 
	DWORD    cbSecurityDescriptor; // size of security descriptor 
	FILETIME ftLastWriteTime;      // last write time
	void SetDflts()
	{
		cbName = MAX_KEY_NAME_LENGTH;
		achClass[0] = 0;
		cchClassName = MAX_PATH;
		cSubKeys = 0;
	}
};

class CProfile {
public:
/*	CProfile();
	CProfile(CString path);
	BOOL _WriteProfileInt( LPCWSTR section, LPCWSTR key, DWORD value );
	BOOL _WriteProfileString( LPCWSTR section, LPCWSTR key, CString &value );
	int _GetProfileInt( LPCWSTR section, LPCWSTR key );
	CString _GetProfileString( LPCWSTR section, LPCWSTR key );
	void _SetProfileAppKey(CString str) { m_regAppKey = str; };
*/	
	static BOOL _DeleteProfileString(HKEY hKey, LPCWSTR section, LPCWSTR key);
	static BOOL _WriteProfileInt( HKEY hKey, LPCWSTR section, LPCWSTR key, DWORD value );
	static BOOL _WriteProfileString( HKEY hKey, LPCWSTR section, LPCWSTR key, CString &value );
	static BOOL _WriteProfileBinary(HKEY hKey, LPCWSTR section, LPCWSTR key, const BYTE *lpData, DWORD cbData);

/*	static BOOL _WriteProfileString( HKEY hKey, LPCWSTR section, LPCWSTR key, int value ) {
		CString w;
		w.Format("%d", value);
		return _WriteProfileString( hKey, section, key, w );
	}*/
	static int _GetProfileInt( HKEY hKey, LPCWSTR section, LPCWSTR key );
	static CString _GetProfileString( HKEY hKey, LPCWSTR section, LPCWSTR key );
	static BOOL _GetProfileInt(HKEY hKey, LPCWSTR section, LPCWSTR key, DWORD &intval);
	static BOOL _GetProfileInt(HKEY hKey, LPCWSTR section, LPCWSTR key, int &intval);
	static BOOL _GetProfileString(HKEY hKey, LPCWSTR section, LPCWSTR key, CString &str);
	static BOOL _GetProfileBinary(HKEY hKey, LPCWSTR section, LPCWSTR key, BYTE *lpData, DWORD &cbData);
	//
	static BOOL _DeleteValue(HKEY hKey, LPCWSTR section, LPCWSTR key);
	static BOOL _DeleteKey(HKEY hKey, LPCWSTR section, LPCWSTR key, BOOL recursive);
	//
	static LSTATUS CopyKey(HKEY hKey, LPCWSTR section, LPCWSTR toSection);
	static LSTATUS CopySubKeys(HKEY hKey, LPCWSTR section, LPCWSTR toSection);
	static LSTATUS CopyKeyValueList(HKEY hKey, LPCWSTR section, LPCWSTR toSection, KeyFromToTable& arr);
	static BOOL CheckIfKeyExists(HKEY hKey, LPCWSTR section);
	//
	static LSTATUS InitRegQueryInfoKeyParams(HKEY hKey, RegQueryInfoKeyParams& params);
	static LSTATUS EnumerateAllSubKeys(HKEY hKey, LPCWSTR section);
	static LSTATUS EnumerateAllSubKeyValues(HKEY hKey, LPCWSTR section);

private:
//	CString	m_regAppKey;
};

#endif // _PROFILE_
