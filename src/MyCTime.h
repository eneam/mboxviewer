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


#if !defined(_MYCTIME__)
#define _MYCTIME__

#pragma once


#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif



// Microsoft CTime throws exceptions and mboxview doesn't handle well exceptions
// Reimplemented some methods as MyCTime but without exceptions
// Also most important it assumes input time is GMT and not local 
class MyCTime
{
public:
	// CTime m_CTime;  // should I just implement soime and delegate other to CTime

#if 0
	static MyCTime WINAPI GetCurrentTime() throw();
	static BOOL WINAPI IsValidFILETIME(_In_ const FILETIME& ft) throw();
#endif

	MyCTime();
	MyCTime(__time64_t time);
	MyCTime(
		int nYear,
		int nMonth,
		int nDay,
		int nHour,
		int nMin,
		int nSec,
		int nDST = -1);
#if 0
	MyCTime(
		WORD wDosDate,
		WORD wDosTime,
		int nDST = -1);
#endif
	MyCTime(
		const SYSTEMTIME& st,
		int nDST = -1);

	bool SetDateTime(
		const SYSTEMTIME& st,
		int nDST = -1);

	bool SetDate(
		const SYSTEMTIME& st,
		int nDST = -1);

	MyCTime& operator=(__time64_t time) { m_time = time; return *this;  }

#if 0
	MyCTime(
		const FILETIME& ft,
		int nDST = -1);
	MyCTime(
		const DBTIMESTAMP& dbts,
		int nDST = -1) throw();

	

	MyCTime& operator+=(MyCTimeSpan span);
	MyCTime& operator-=(MyCTimeSpan span);

	MyCTimeSpan operator-(MyCTime time);
	MyCTime operator-(MyCTimeSpan span);
	MyCTime operator+(MyCTimeSpan span);

	bool operator==(MyCTime time);
	bool operator!=(MyCTime time);
	bool operator<(MyCTime time);
	bool operator>(MyCTime time);
	bool operator<=(MyCTime time);
	bool operator>=(MyCTime time);
#endif

	struct tm* GetGmtTm(struct tm* ptm);
	struct tm* GetLocalTm(struct tm* ptm);

	bool GetAsSystemTime(SYSTEMTIME& st, BOOL gmtTime = TRUE);
#if 0
	bool GetAsDBTIMESTAMP(_Out_ DBTIMESTAMP& dbts);
#endif

	bool IsValid();

	static bool validateSystemtime(const SYSTEMTIME *sysTime);
	static bool fixSystemtime(SYSTEMTIME *sysTime);

	__time64_t GetTime();

#if 0
	int GetYear();
	int GetMonth();
	int GetDay();
	int GetHour();
	int GetMinute();
	int GetSecond();
	int GetDayOfWeek();
#endif

	// formatting using "C" strftime
#if 0
	CString Format(LPCWSTR pszFormat);
	CString FormatGmt(LPCWSTR pszFormat);
#endif
	CStringA FormatLocalTmA(CStringA &format);  // mboxview custom
	CStringA FormatGmtTmA(CStringA &format);  // mboxview custom
	//
	//
	CStringA FormatLocalTmA(CString &format);  // mboxview custom
	CStringA FormatGmtTmA(CString& format);  // mboxview custom
	//
	CString FormatLocalTm(CString& format);  // mboxview custom
	CString FormatGmtTm(CString& format);  // mboxview custom
#if 0
	CString Format(UINT nFormatID);
	CString FormatGmt(UINT nFormatID);
#endif

private:
	CString Format(LPCWSTR pszFormat);
	CString FormatGmt(LPCWSTR pszFormat);
	__time64_t MaxTime();
	//
	__time64_t m_time;
};


#endif // _MBOX_MAIL_