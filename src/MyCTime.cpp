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
#include <afxtempl.h>
#include "TextUtilsEx.h"
#include "MyCTime.h"


#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif


// Microsoft CTime generates exceptions which currently are not handled and that makes mboxview unusable
// Reimplemented subset of methods without exceptions
// Also most important it assumes input time is GMT and not local
// However, lowest/samllest Date is still January 1, 1970 which is the limitation to fix at later time
//

// time_t my_timegm(struct tm *tm) // Linux

MyCTime::MyCTime()
{
	__time64_t UTCTime = _time64(&m_time);
}

MyCTime::MyCTime(__time64_t time) :
	m_time(time)
{
}

__time64_t MyCTime::MaxTime()
{
	//23:59 : 59, December 31, 3000
	int nYear = 3000;
	int nMonth = 12;
	int nDay = 31;
	int nHour = 23;
	int nMin = 59;
	int nSec = 59;
	int nDST = -1;
	MyCTime tm(nYear, nMonth, nDay, nHour, nMin, nSec, nDST);
	__time64_t maxTime = tm.GetTime();
	return maxTime;
}

MyCTime::MyCTime(
	int nYear,
	int nMonth,
	int nDay,
	int nHour,
	int nMin,
	int nSec,
	int nDST)
{
#if 0
	_ASSERTE(nYear >= 1970);
	_ASSERTE(nMonth >= 1 && nMonth <= 12);
	_ASSERTE(nDay >= 1 && nDay <= 31);
	_ASSERTE(nHour >= 0 && nHour <= 23);
	_ASSERTE(nMin >= 0 && nMin <= 59);
	_ASSERTE(nSec >= 0 && nSec <= 59);
#endif

	struct tm atm;

	atm.tm_sec = nSec;
	atm.tm_min = nMin;
	atm.tm_hour = nHour;
	atm.tm_mday = nDay;
	atm.tm_mon = nMonth - 1;        // tm_mon is 0 based
	atm.tm_year = nYear - 1900;     // tm_year is 1900 based
	atm.tm_isdst = nDST;

	//m_time = _mktime64(&atm);  // This is how Microsoft CTime works
	m_time = _mkgmtime64(&atm);  // 

	if (m_time != -1)
		int deb = 1;
}

MyCTime::MyCTime(
	const SYSTEMTIME& sysTime,
	int nDST)
{
	if (sysTime.wYear < 1900)
	{
		__time64_t time0 = 0L;
		MyCTime timeT(time0);
		*this = timeT;
	}
	else
	{
		MyCTime timeT(
			(int)sysTime.wYear, (int)sysTime.wMonth, (int)sysTime.wDay,
			(int)sysTime.wHour, (int)sysTime.wMinute, (int)sysTime.wSecond,
			nDST);
		*this = timeT;
	}
}

bool MyCTime::SetDateTime(
	const SYSTEMTIME& sysTime,
	int nDST)
{
	if (sysTime.wYear < 1900)
	{
		__time64_t time0 = 0L;
		MyCTime timeT(time0);
		*this = timeT;
	}
	else
	{
		MyCTime timeT(
			(int)sysTime.wYear, (int)sysTime.wMonth, (int)sysTime.wDay,
			(int)sysTime.wHour, (int)sysTime.wMinute, (int)sysTime.wSecond,
			nDST);
		*this = timeT;
	}
	return true;
}

bool MyCTime::SetDate(
	const SYSTEMTIME& sysTime,
	int nDST)
{
	SYSTEMTIME st = sysTime;
	st.wHour = 0;
	st.wMinute = 0;
	st.wSecond = 0;
	st.wMilliseconds = 0;

	SetDateTime(st);
	return true;
}

struct tm* MyCTime::GetGmtTm(struct tm* ptm)
{
	// Ensure ptm is valid
	//_ASSERTE(ptm != NULL);

	if (ptm != NULL)
	{
		struct tm ptmTemp;
		errno_t err = _gmtime64_s(&ptmTemp, &m_time);

		// Be sure the call succeeded
		if (err != 0) { 
			return NULL; 
		}

		*ptm = ptmTemp;
		return ptm;
	}

	return NULL;
}

struct tm* MyCTime::GetLocalTm(struct tm* ptm)
{
	// Ensure ptm is valid
	//_ASSERTE(ptm != NULL);

	if (ptm != NULL)
	{
		struct tm ptmTemp;
		errno_t err = _localtime64_s(&ptmTemp, &m_time);

		if (err != 0)
		{
			return NULL;    // indicates that m_time was not initialized!
		}

		*ptm = ptmTemp;
		return ptm;
	}

	return NULL;
}

__time64_t MyCTime::GetTime()
{
	return(m_time);
}

CString MyCTime::Format(LPCWSTR pFormat)
{
	wchar_t szBuffer[maxTimeBufferSize];
	if (pFormat == NULL)
	{
		//return pFormat;
		szBuffer[0] = '\0';
		return szBuffer;
	}

	struct tm ptmTemp;

	if (_localtime64_s(&ptmTemp, &m_time) != 0)
	{
		szBuffer[0] = '\0';
		return szBuffer;
	}

	if (!_tcsftime(szBuffer, maxTimeBufferSize, pFormat, &ptmTemp))
	{
		szBuffer[0] = '\0';
	}

	return szBuffer;
}

CString MyCTime::FormatGmt(LPCWSTR pFormat)
{
	wchar_t szBuffer[maxTimeBufferSize];

	if (pFormat == NULL)
	{
		//return pFormat;
		szBuffer[0] = '\0';
		return szBuffer;
	}
	struct tm ptmTemp;

	if (_gmtime64_s(&ptmTemp, &m_time) != 0)
	{
		szBuffer[0] = '\0';
		return szBuffer;
	}

	if (!_tcsftime(szBuffer, maxTimeBufferSize, pFormat, &ptmTemp))
	{
		szBuffer[0] = '\0';
	}

	return szBuffer;
}

bool MyCTime::GetAsSystemTime(SYSTEMTIME& timeDest, int gmtTime)
{
	struct tm ttm;
	struct tm* ptm;

	if (gmtTime)
		ptm = GetGmtTm(&ttm);
	else
		ptm = GetLocalTm(&ttm);
	if (!ptm)
	{
		return false;
	}

	timeDest.wYear = (WORD)(1900 + ptm->tm_year);
	timeDest.wMonth = (WORD)(1 + ptm->tm_mon);
	timeDest.wDayOfWeek = (WORD)ptm->tm_wday;
	timeDest.wDay = (WORD)ptm->tm_mday;
	timeDest.wHour = (WORD)ptm->tm_hour;
	timeDest.wMinute = (WORD)ptm->tm_min;
	timeDest.wSecond = (WORD)ptm->tm_sec;
	timeDest.wMilliseconds = 0;

	return true;
}

bool MyCTime::validateSystemtime(const SYSTEMTIME *sysTime)
{
	if ((sysTime->wYear < 1970) || (sysTime->wYear > 3000))
		return false;
	if ((sysTime->wMonth < 1) || (sysTime->wMonth > 12))
		return false;
	// if ((sysTime->wDayOfWeek < 1) || (sysTime->wDayOfWeek > 7)) return false;
	if ((sysTime->wDay < 1) || (sysTime->wDay > 31))  // This seem to be sufficient
		return false;
	if ((sysTime->wHour < 0) || (sysTime->wHour > 23))
		return false;
	if ((sysTime->wMinute < 0) || (sysTime->wMinute > 59))
		return false;
	if ((sysTime->wSecond < 0) || (sysTime->wSecond > 59))
		return false;
	//if ((sysTime->wMilliseconds < 0) || (sysTime->wMilliseconds < 0)) return false;
	return true;
}

bool MyCTime::fixSystemtime(SYSTEMTIME *sysTime)
{
	if (sysTime->wYear < 1970)
	{
		sysTime->wYear = 1970;
		sysTime->wMonth = 1;
		sysTime->wDay = 1;
		sysTime->wHour = 0;
		sysTime->wMinute = 0;
		sysTime->wSecond = 0;
		sysTime->wMilliseconds = 0;

		return true;
	}
	else if (sysTime->wYear > 3000)
	{
		sysTime->wYear = 3000;

		sysTime->wMonth = 12;
		sysTime->wDay = 31;
		sysTime->wHour = 23;
		sysTime->wMinute = 59;
		sysTime->wSecond = 59;
		sysTime->wMilliseconds = 999;

		return true;
	}

	// Attempt to fix date and time
	if (sysTime->wMonth < 1)
		sysTime->wMonth = 1;
	if(sysTime->wMonth > 12)
		sysTime->wMonth = 12;

	// if ((sysTime->wDayOfWeek < 1) || (sysTime->wDayOfWeek > 7)) return false;
	if (sysTime->wDay < 1)
		sysTime->wDay = 1;
	// Each month has different number of days; Not sure we can rely on time functions to fix it
	if (sysTime->wDay > 31)
		sysTime->wDay = 31;

	if (sysTime->wHour < 0)
		sysTime->wHour = 0;
	if (sysTime->wHour > 23)
		sysTime->wHour = 23;

	if (sysTime->wMinute < 0)
		sysTime->wMinute = 0;
	if (sysTime->wMinute > 59)
		sysTime->wMinute = 59;

	if (sysTime->wSecond < 0)
		sysTime->wSecond = 0;
	if (sysTime->wSecond > 59)
		sysTime->wSecond = 59;

	if (sysTime->wMilliseconds < 0)
		sysTime->wMilliseconds = 0;
	if (sysTime->wMilliseconds > 999)
		sysTime->wMilliseconds = 999;

	return true;
}

CStringA MyCTime::FormatLocalTmA(CStringA& format)
{
	CString formatW;
	DWORD error;

	BOOL retA2W = TextUtilsEx::Ansi2WStr(format, formatW, error);
	CStringA tmA = MyCTime::FormatLocalTmA(formatW);
	return tmA;
}

CStringA MyCTime::FormatLocalTmA(CString& format)
{
	CStringA tmA;
	DWORD error;

	CString tmW = MyCTime::FormatLocalTm(format);
	BOOL retW2A = TextUtilsEx::WStr2Ansi(tmW, tmA, error);
	return tmA;
}


CStringA MyCTime::FormatGmtTmA(CStringA& format)
{
	CString formatW;
	DWORD error;

	BOOL retA2W = TextUtilsEx::Ansi2WStr(format, formatW, error);
	CStringA tmA = MyCTime::FormatGmtTmA(formatW);
	return tmA;
}

CStringA MyCTime::FormatGmtTmA(CString& format)
{
	CStringA tmA;
	DWORD error;

	CString tmW = MyCTime::FormatGmtTm(format);
	BOOL retW2A =  TextUtilsEx::WStr2Ansi(tmW, tmA, error);
	return tmA;
}

CString MyCTime::FormatLocalTm(CString &format)
{
	__time64_t maxTime = MaxTime();
	if ((m_time == 0) || (m_time == maxTime))
	{
		CString lDateTime = FormatGmtTm(format);
		return lDateTime;
	}

	SYSTEMTIME st;
	SYSTEMTIME lst;
	bool ret = GetAsSystemTime(st);
	if (SystemTimeToTzSpecificLocalTime(0, &st, &lst))
	{
		MyCTime::fixSystemtime(&lst);
		MyCTime ltt(lst);
		CString lDateTime = ltt.FormatGmt(format);
		return lDateTime;
	}
	else
	{
		CString lDateTime = FormatGmtTm(format);
		return lDateTime;
	}
}

CString MyCTime::FormatGmtTm(CString &format)
{
	CString lDateTime = FormatGmt(format);
	return lDateTime;
}

bool MyCTime::IsValid()
{
	return (m_time != -1);
}





