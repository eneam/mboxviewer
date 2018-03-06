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
/*	static BOOL _WriteProfileString( HKEY hKey, LPCTSTR section, LPCTSTR key, int value ) {
		CString w;
		w.Format("%d", value);
		return _WriteProfileString( hKey, section, key, w );
	}*/
	static int _GetProfileInt( HKEY hKey, LPCTSTR section, LPCTSTR key );
	static CString _GetProfileString( HKEY hKey, LPCTSTR section, LPCTSTR key );
	static BOOL _GetProfileInt(HKEY hKey, LPCTSTR section, LPCTSTR key, DWORD &intval);
	static BOOL _GetProfileString(HKEY hKey, LPCTSTR section, LPCTSTR key, CString &str);
private:
//	CString	m_regAppKey;
};
