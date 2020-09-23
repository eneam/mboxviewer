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

class CProfile {
public:
/*	CProfile();
	CProfile(CString path);
	BOOL _WriteProfileInt( LPCTSTR section, LPCTSTR key, DWORD value );
	BOOL _WriteProfileString( LPCTSTR section, LPCTSTR key, CString &value );
	int _GetProfileInt( LPCTSTR section, LPCTSTR key );
	CString _GetProfileString( LPCTSTR section, LPCTSTR key );
	void _SetProfileAppKey(CString str) { m_regAppKey = str; };
*/	
	static BOOL _DeleteProfileString(HKEY hKey, LPCTSTR section, LPCTSTR key);
	static BOOL _WriteProfileInt( HKEY hKey, LPCTSTR section, LPCTSTR key, DWORD value );
	static BOOL _WriteProfileString( HKEY hKey, LPCTSTR section, LPCTSTR key, CString &value );
	static BOOL _WriteProfileBinary(HKEY hKey, LPCTSTR section, LPCTSTR key, const BYTE *lpData, DWORD cbData);

/*	static BOOL _WriteProfileString( HKEY hKey, LPCTSTR section, LPCTSTR key, int value ) {
		CString w;
		w.Format("%d", value);
		return _WriteProfileString( hKey, section, key, w );
	}*/
	static int _GetProfileInt( HKEY hKey, LPCTSTR section, LPCTSTR key );
	static CString _GetProfileString( HKEY hKey, LPCTSTR section, LPCTSTR key );
	static BOOL _GetProfileInt(HKEY hKey, LPCTSTR section, LPCTSTR key, DWORD &intval);
	static BOOL _GetProfileInt(HKEY hKey, LPCTSTR section, LPCTSTR key, int &intval);
	static BOOL _GetProfileString(HKEY hKey, LPCTSTR section, LPCTSTR key, CString &str);
	static BOOL _GetProfileBinary(HKEY hKey, LPCTSTR section, LPCTSTR key, BYTE *lpData, DWORD &cbData);
	//
	static BOOL _DeleteKey(HKEY hKey, LPCTSTR section, LPCTSTR key);
private:
//	CString	m_regAppKey;
};

#endif // _PROFILE_
