//
// Provided by Chilkat Software, Inc.  (http://www.chilkatsoft.com)
//
// Makers of ActiveX and .NET components:
//
// Zip Compression Component
// Email Component
// Encryption Component
// S/MIME Component
// HTML Email Component
// Character Encoding Conversion Component
// FTP Component (Free)
// Super-Fast XML Parser Component (Free)
// ASP Email Component
// ASP Super-Fast XML Parser Component (Free)
// Free SSL Component (Free)
// Real-time Compression Component (Free)
// DirectX Game Development C++ Library (Free, Open Source)
// XML Messaging Component
// EXE Stuffer Component (Embed anything in an EXE)
// Digital Certificates Component (Free)
// "Mail This Page!" utility (Free)
// Zip 2 Secure EXE (Free) for creating self-extracting EXEs
//
 
// DateParser.cpp: implementation of the DateParser class.
//
//////////////////////////////////////////////////////////////////////
 
/****************************************************************************/
/**																								**/
/** Modified 2005 by Diodia Software				(http://www.diodia.com/)	**/
/**																								**/
/****************************************************************************/
#include "stdafx.h"
 
#include "DateParser.h"
#include "TextUtilsEx.h"
#include <string.h>
#include <stdio.h>
#include <math.h>
#include <time.h>

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif
 
/*
	  5.1.  SYNTAX
 
	  date-time	=  [ day "," ] date time		  ; dd mm yy
																 ;  hh:mm:ss zzz
 
	  day			=  "Mon"  / "Tue" /  "Wed"  / "Thu"
					  /  "Fri"  / "Sat" /  "Sun"
 
	  date		  =  1*2DIGIT month 2DIGIT		  ; day month year
																 ;  e.g. 20 Jun 82
 
	  month		 =  "Jan"  /  "Feb" /  "Mar"  /  "Apr"
					  /  "May"  /  "Jun" /  "Jul"  /  "Aug"
					  /  "Sep"  /  "Oct" /  "Nov"  /  "Dec"
 
	  time		  =  hour zone						  ; ANSI and Military
 
	  hour		  =  2DIGIT ":" 2DIGIT [":" 2DIGIT]
																 ; 00:00:00 - 23:59:59
 
	  zone		  =  "UT"  / "GMT"					 ; Universal Time
																 ; North American : UT
					  /  "EST" / "EDT"					 ;  Eastern:  - 5/ - 4
					  /  "CST" / "CDT"					 ;  Central:  - 6/ - 5
					  /  "MST" / "MDT"					 ;  Mountain: - 7/ - 6
					  /  "PST" / "PDT"					 ;  Pacific:  - 8/ - 7
					  /  1ALPHA							  ; Military: Z = UT;
																 ;  A:-1; (J not used)
																 ;  M:-12; N:+1; Y:+12
					  / ( ("+" / "-") 4DIGIT )		  ; Local differential
																 ;  hours+min. (HHMM)
 
	  5.2.  SEMANTICS
 
			 If included, day-of-week must be the day implied by the date
	  specification.
 
			 Time zone may be indicated in several ways.  "UT" is Univer-
	  sal  Time  (formerly called "Greenwich Mean Time"); "GMT" is per-
	  mitted as a reference to Universal Time.  The  military  standard
	  uses  a  single  character for each zone.  "Z" is Universal Time.
	  "A" indicates one hour earlier, and "M" indicates 12  hours  ear-
	  lier;  "N"  is  one  hour  later, and "Y" is 12 hours later.  The
	  letter "J" is not used.  The other remaining two forms are  taken
	  from ANSI standard X3.51-1975.  One allows explicit indication of
	  the amount of offset from UT; the other uses  common  3-character
	  strings for indicating time zones in North America.
 
*/
 
static char *days[] = { "sun","mon","tue","wed","thu","fri","sat" };
static char *months[] = { "jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec" };

inline
bool return_false() {
	return false;
}
 
#define SKIP_WHITESPACE() while (*s == ' ' || *s == '\t') s++
#define SKIP_NON_WHITESPACE() while (*s != ' ' && *s != '\t' && *s != '\0') s++
//#define CHECK_PREMATURE_END()	 if (*s == '\0') return false
#define CHECK_PREMATURE_END()	 if (*s == '\0') return return_false();



bool DateParser::parseRFC822Date(const char *str1, SYSTEMTIME *sysTime, int dateFormatType)
{
	// dateFormatType == 1
// Date: Sun, 5 Nov 2017 22:42:43 -0600 (CST)
// dateFormatType == 2
// From 1583290388308606088@xxx Mon Nov 06 04:42:58 +0000 2017 

	if (dateFormatType != 1)
		return false;   // TODO: dateFormatType ==2 needs some work

	if (!sysTime || !str1)
		return false;
 
	int dayOfWeek = 0;
 
	char str[300],*s,*s_end;
	s = &str[0];
	//s_end = &str[300];
	strncpy(&str[0],str1,299);
	str[299] = '\0';

	int slen = istrlen(s);
	s_end = &str[slen];

	// Convert to lowercase.
	int j;
	int i = 0;
	while (str[i] != '\0')
	{
		str[i] = tolower(str[i]);
		i++;
	}

	int cnt = TextUtilsEx::ReplaceChar(str, '-', ' ');
 
	SKIP_WHITESPACE();
	CHECK_PREMATURE_END();

	// dateFormatType == 1
	// Date: Sun, 5 Nov 2017 22:42:43 -0600 (CST)
	// dateFormatType == 2
	// From 1583290388308606088@xxx Mon Nov 06 04:42:58 +0000 2017 

	if ((s_end - s) < 3) 
		return false;
 
	for (j=0; j<7; j++)
	{
		if (strncmp(s,days[j],3) == 0)
		{
			break;
		}
	}
 
	if (j < 7)
	{
		// skip days
		SKIP_NON_WHITESPACE();
		dayOfWeek = j;
		SKIP_WHITESPACE();
		CHECK_PREMATURE_END();

		if (*s == ',') 
			s++;
		SKIP_WHITESPACE();
		CHECK_PREMATURE_END();
	}
 
	// Get the day, month, and year.
	int day;
	char monthStr[300];
	int month;
	int year;
	int hour, minute, seconds;
 
	if (dateFormatType == 1)
	{
		// Parse Date field
		// Date: Sun, 5 Nov 2017 22:42:43 -0600 (CST)

		if (sscanf(s, "%d%s%d", &day, monthStr, &year) != 3)
			return false;

		// skip day
		SKIP_NON_WHITESPACE();
		SKIP_WHITESPACE();
		// skip month
		SKIP_NON_WHITESPACE();
		SKIP_WHITESPACE();
		// skip year
		SKIP_NON_WHITESPACE();
		SKIP_WHITESPACE();

		for (j = 0; j < 12; j++)
		{
			if (strncmp(monthStr, months[j], 3) == 0)
			{
				break;
			}
		}
		if (j == 12)
			return false;
		month = j;

		if (year < 1900)
		{
			if (year < 50)
			{
				year += 2000;
			}
			else
			{
				year += 1900;
			}
		}

		if (sscanf(s, "%d:%d:%d", &hour, &minute, &seconds) != 3)
		{
			seconds = 0;
			if (sscanf(s, "%d:%d", &hour, &minute) != 2)
				return false;
		}

		// skip hour & minute & seconds
		SKIP_NON_WHITESPACE();
		SKIP_WHITESPACE();
	}
	else
	{
		// Parse the From field
		// From 1583290388308606088@xxx Mon Nov 06 04:42:58 +0000 2017 

		if (sscanf(s, "%3s%d", monthStr, &day) != 2)
			return false;

		// skip month
		SKIP_NON_WHITESPACE();
		SKIP_WHITESPACE();
		// skip day
		SKIP_NON_WHITESPACE();
		SKIP_WHITESPACE();

		for (j = 0; j < 12; j++)
		{
			if (strncmp(monthStr, months[j], 3) == 0)
			{
				break;
			}
		}
		if (j == 12)
			return false;
		month = j;

		if (sscanf(s, "%d:%d:%d", &hour, &minute, &seconds) != 3)
		{
			seconds = 0;
			if (sscanf(s, "%d:%d", &hour, &minute) != 2)
				return false;
		}

		// skip hour & minute & seconds
		SKIP_NON_WHITESPACE();
		SKIP_WHITESPACE();

		char *s_save = 0;
		// TODO: fixme
		if ((*s == '+') || (*s == '-'))
		{
			s_save = s;  // pointer to zone

			// skip zone
			SKIP_NON_WHITESPACE();
			SKIP_WHITESPACE();
			CHECK_PREMATURE_END();
		}

		if (sscanf(s, "%d", &year) != 1)
			return false;

		if (year < 1900)
		{
			if (year < 50)
			{
				year += 2000;
			}
			else
			{
				year += 1900;
			}
		}

		s = s_save;
	}

	//CHECK_PREMATURE_END();
 
	if (*s == '+') 
		s++;

	//CHECK_PREMATURE_END();

	char zoneStr[300];
	if (sscanf(s,"%s",zoneStr) != 1)
	{
		strcpy(zoneStr,"GMT");
	}
 
	if (zoneStr[0] == '-' ||
		(zoneStr[0] >= '0' && zoneStr[0] <= '9'))
	{
		char *z = zoneStr;
		if (*z == '-')
			z++;
		char *zend = TextUtilsEx::SkipNumeric(z);
		if (*zend)
		{
			if (!_istdigit(*zend))
				*zend = NULL;
		}
		; // Do nothing.
	}
	else
	{
		char *zend = TextUtilsEx::findOneOf(zoneStr, &zoneStr[strlen(zoneStr)], " ,)}]>");
		if (zend)
			*zend = NULL;

		if (strcmp(zoneStr, "ut") == 0)
		{
			strcpy(zoneStr, "0000");
		}
		else if (strcmp(zoneStr, "gmt") == 0)
		{
			strcpy(zoneStr, "0000");
		}
		else if (strcmp(zoneStr, "est") == 0)
		{
			strcpy(zoneStr, "-0500");
		}
		else if (strcmp(zoneStr, "edt") == 0)
		{
			strcpy(zoneStr, "-0400");
		}
		else if (strcmp(zoneStr, "cst") == 0)
		{
			strcpy(zoneStr, "-0600");
		}
		else if (strcmp(zoneStr, "cdt") == 0)
		{
			strcpy(zoneStr, "-0500");
		}
		else if (strcmp(zoneStr, "mst") == 0)
		{
			strcpy(zoneStr, "-0700");
		}
		else if (strcmp(zoneStr, "mdt") == 0)
		{
			strcpy(zoneStr, "-0600");
		}
		else if (strcmp(zoneStr, "pst") == 0)
		{
			strcpy(zoneStr, "-0800");
		}
		else if (strcmp(zoneStr, "pdt") == 0)
		{
			strcpy(zoneStr, "-0700");
		}
		else if (strcmp(zoneStr, "a") == 0)
		{
			strcpy(zoneStr, "-0100");
		}
		else if (strcmp(zoneStr, "z") == 0)
		{
			strcpy(zoneStr, "0000");
		}
		else if (strcmp(zoneStr, "m") == 0)
		{
			strcpy(zoneStr, "-1200");
		}
		else if (strcmp(zoneStr, "n") == 0)
		{
			strcpy(zoneStr, "0100");
		}
		else if (strcmp(zoneStr, "y") == 0)
		{
			strcpy(zoneStr, "1200");
		}
		else
		{
			strcpy(zoneStr, "0000");
		}
	}
 
	// Convert zoneStr from (-)hhmm to minutes.

	int hh,mm;
	int sign = 1;

	s = zoneStr;
	CHECK_PREMATURE_END();  // overkill should/will never be true
	if (*s == '-')
	{
		sign = -1;
		s++;
	}
	CHECK_PREMATURE_END(); // overkill should/will never be true

	if (sscanf(s,"%02d%02d",&hh,&mm) != 2)
	{
		return false;
	}
	int adjustment = sign * ((hh*60) + mm);
 
	// SystemTimeToVariantTime
	sysTime->wYear = year;
	sysTime->wMonth = month+1;
	sysTime->wDayOfWeek = dayOfWeek;
	sysTime->wDay = day;
	sysTime->wHour = hour;
	sysTime->wMinute = minute;
	sysTime->wSecond = seconds;
	sysTime->wMilliseconds = 0;
 
 
	int temp = hour*60 + minute;
	temp -= adjustment;
	if (temp < 0)
	{
		// Need to go back 1 day.
		temp += 1440;
		double d;
		SystemTimeToVariantTime(sysTime,&d);
		d -= 1.0;
		VariantTimeToSystemTime(d,sysTime);
	}
	else if (temp >= 1440)
	{
		// Need to go forward 1 day.
		temp -= 1440;
		double d;
		SystemTimeToVariantTime(sysTime,&d);
		d += 1.0;
		VariantTimeToSystemTime(d,sysTime);
	}
 
	sysTime->wHour = temp/60;
	sysTime->wMinute = temp%60;
 
	return true;
}

// bool DateParser::parseRFC822Date(const char *str1, SYSTEMTIME *sysTime) doesn't validate 
// result and causes CTime() to fail.
bool DateParser::validateSystemtime(const SYSTEMTIME *sysTime)
{
	// if (sysTime->wYear ??
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
 
static int _afxMonthDays[13] =
		{0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365};
 
#define MIN_DATE					 (-657434L)  // about year 100
#define MAX_DATE					 2958465L	 // about year 9999
 
// Half a second, expressed in days
#define HALF_SECOND  (1.0/172800.0)
 
int DateParser::OleDateFromTm(WORD wYear, WORD wMonth, WORD wDay,
		WORD wHour, WORD wMinute, WORD wSecond, DATE& dtDest)
{
	// Validate year and month (ignore day of week and milliseconds)
	if (wYear > 9999 || wMonth < 1 || wMonth > 12)
		return FALSE;
 
	//  Check for leap year and set the number of days in the month
	BOOL bLeapYear = ((wYear & 3) == 0) &&
		((wYear % 100) != 0 || (wYear % 400) == 0);
 
	int nDaysInMonth =
		_afxMonthDays[wMonth] - _afxMonthDays[wMonth-1] +
		((bLeapYear && wDay == 29 && wMonth == 2) ? 1 : 0);
 
	// Finish validating the date
	if (wDay < 1 || wDay > nDaysInMonth ||
		wHour > 23 || wMinute > 59 ||
		wSecond > 59)
	{
		return FALSE;
	}
 
	// Cache the date in days and time in fractional days
	long nDate;
	double dblTime;
 
	//It is a valid date; make Jan 1, 1AD be 1
	nDate = wYear*365L + wYear/4 - wYear/100 + wYear/400 +
		_afxMonthDays[wMonth-1] + wDay;
 
	//  If leap year and it's before March, subtract 1:
	if (wMonth <= 2 && bLeapYear)
		--nDate;
 
	//  Offset so that 12/30/1899 is 0
	nDate -= 693959L;
 
	dblTime = ((double)((long)wHour * 3600L) +  // hrs in seconds
		((long)wMinute * 60L) +  // mins in seconds
		((long)wSecond)) / 86400.;
 
	dtDest = (double) nDate + ((nDate >= 0) ? dblTime : -dblTime);
 
	return TRUE;
}
 
DATE DateParser::SystemTimeToDate(SYSTEMTIME &systime)
{
	DATE pVal;
	OleDateFromTm(systime.wYear, systime.wMonth,
		systime.wDay, systime.wHour, systime.wMinute,
		systime.wSecond, pVal);
	return pVal;
}
 
static bool TmFromOleDate(DATE dtSrc, struct tm& tmDest)
{
	// The legal range does not actually span year 0 to 9999.
	if (dtSrc > MAX_DATE || dtSrc < MIN_DATE) // about year 100 to about 9999
		return FALSE;
 
	long nDays;				 // Number of days since Dec. 30, 1899
	long nDaysAbsolute;	  // Number of days since 1/1/0
	long nSecsInDay;		  // Time in seconds since midnight
	long nMinutesInDay;	  // Minutes in day
 
	long n400Years;			// Number of 400 year increments since 1/1/0
	long n400Century;		 // Century within 400 year block (0,1,2 or 3)
	long n4Years;			  // Number of 4 year increments since 1/1/0
	long n4Day;				 // Day within 4 year block
	//  (0 is 1/1/yr1, 1460 is 12/31/yr4)
	long n4Yr;				  // Year within 4 year block (0,1,2 or 3)
	BOOL bLeap4 = TRUE;	  // TRUE if 4 year block includes leap year
 
	double dblDate = dtSrc; // tempory serial date
 
	// If a valid date, then this conversion should not overflow
	nDays = (long)dblDate;
 
	// Round to the second
	dblDate += ((dtSrc > 0.0) ? HALF_SECOND : -HALF_SECOND);
 
	nDaysAbsolute = (long)dblDate + 693959L; // Add days from 1/1/0 to 12/30/1899
 
	dblDate = fabs(dblDate);
	nSecsInDay = (long)((dblDate - floor(dblDate)) * 86400.);
 
	// Calculate the day of week (sun=1, mon=2...)
	//	-1 because 1/1/0 is Sat.  +1 because we want 1-based
	tmDest.tm_wday = (int)((nDaysAbsolute - 1) % 7L) + 1;
 
	// Leap years every 4 yrs except centuries not multiples of 400.
	n400Years = (long)(nDaysAbsolute / 146097L);
 
	// Set nDaysAbsolute to day within 400-year block
	nDaysAbsolute %= 146097L;
 
	// -1 because first century has extra day
	n400Century = (long)((nDaysAbsolute - 1) / 36524L);
 
	// Non-leap century
	if (n400Century != 0)
	{
		// Set nDaysAbsolute to day within century
		nDaysAbsolute = (nDaysAbsolute - 1) % 36524L;
 
		// +1 because 1st 4 year increment has 1460 days
		n4Years = (long)((nDaysAbsolute + 1) / 1461L);
 
		if (n4Years != 0)
			n4Day = (long)((nDaysAbsolute + 1) % 1461L);
		else
		{
			bLeap4 = FALSE;
			n4Day = (long)nDaysAbsolute;
		}
	}
	else
	{
		// Leap century - not special case!
		n4Years = (long)(nDaysAbsolute / 1461L);
		n4Day = (long)(nDaysAbsolute % 1461L);
	}
 
	if (bLeap4)
	{
		// -1 because first year has 366 days
		n4Yr = (n4Day - 1) / 365;
 
		if (n4Yr != 0)
			n4Day = (n4Day - 1) % 365;
	}
	else
	{
		n4Yr = n4Day / 365;
		n4Day %= 365;
	}
 
	// n4Day is now 0-based day of year. Save 1-based day of year, year number
	tmDest.tm_yday = (int)n4Day + 1;
	tmDest.tm_year = n400Years * 400 + n400Century * 100 + n4Years * 4 + n4Yr;
 
	// Handle leap year: before, on, and after Feb. 29.
	if (n4Yr == 0 && bLeap4)
	{
		// Leap Year
		if (n4Day == 59)
		{
			/* Feb. 29 */
			tmDest.tm_mon = 2;
			tmDest.tm_mday = 29;
			goto DoTime;
		}
 
		// Pretend it's not a leap year for month/day comp.
		if (n4Day >= 60)
			--n4Day;
	}
 
	// Make n4DaY a 1-based day of non-leap year and compute
	//  month/day for everything but Feb. 29.
	++n4Day;
 
	// Month number always >= n/32, so save some loop time */
	for (tmDest.tm_mon = (n4Day >> 5) + 1;
		n4Day > _afxMonthDays[tmDest.tm_mon]; tmDest.tm_mon++);
 
		tmDest.tm_mday = (int)(n4Day - _afxMonthDays[tmDest.tm_mon-1]);
 
DoTime:
	if (nSecsInDay == 0)
		tmDest.tm_hour = tmDest.tm_min = tmDest.tm_sec = 0;
	else
	{
		tmDest.tm_sec = (int)nSecsInDay % 60L;
		nMinutesInDay = nSecsInDay / 60L;
		tmDest.tm_min = (int)nMinutesInDay % 60;
		tmDest.tm_hour = (int)nMinutesInDay / 60;
	}
 
	return TRUE;
}
 
void DateParser::SystemTimeFromOleDate(DATE &date, SYSTEMTIME &sysTime)
{
 
	struct tm tmTemp;
	if (TmFromOleDate(date, tmTemp))
	{
		sysTime.wYear = (WORD) tmTemp.tm_year;
		sysTime.wMonth = (WORD) tmTemp.tm_mon;
		sysTime.wDayOfWeek = (WORD) (tmTemp.tm_wday - 1);
		sysTime.wDay = (WORD) tmTemp.tm_mday;
		sysTime.wHour = (WORD) tmTemp.tm_hour;
		sysTime.wMinute = (WORD) tmTemp.tm_min;
		sysTime.wSecond = (WORD) tmTemp.tm_sec;
		sysTime.wMilliseconds = 0;
	}
	else
	{
		GetSystemTime(&sysTime);
	}
 
	return;
}
 
void DateParser::generateCurrentDateRFC822(char *formattedDate, int bufSize)
{
	TIME_ZONE_INFORMATION tz;
	int bias;
 
	if (GetTimeZoneInformation(&tz) == TIME_ZONE_ID_DAYLIGHT)
	{
		bias = tz.Bias + tz.DaylightBias;
	}
	else
	{
		bias = tz.Bias;
	}
 
	char biasStr[60];
	sprintf(biasStr,"%+.2d%.2d",-bias/60, bias%60);
 
	time_t currentTime = time(0);
	struct tm *tmPtr = localtime(&currentTime);
 
	strftime(formattedDate,bufSize,"%a, %d %b %Y %H:%M:%S ",tmPtr);
	strcat(formattedDate,biasStr);
 
	return;
}

DateParser::DateParser()
{
}

DateParser::~DateParser()
{
}