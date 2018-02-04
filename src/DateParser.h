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
// DateParser.h: interface for the DateParser class.
//
//////////////////////////////////////////////////////////////////////

/****************************************************************************/
/**																								**/
/** Modified 2005 by Diodia Software				(http://www.diodia.com/)	**/
/**																								**/
/****************************************************************************/
#pragma once

class DateParser 
{
 public:
	DateParser();
	virtual ~DateParser();

	static bool parseRFC822Date(const char *str, SYSTEMTIME *sysTime);

	int OleDateFromTm(WORD wYear, WORD wMonth, WORD wDay,
		WORD wHour, WORD wMinute, WORD wSecond, DATE& dtDest);

	DATE SystemTimeToDate(SYSTEMTIME &sysTime);

	void SystemTimeFromOleDate(DATE &date, SYSTEMTIME &systime);

	void generateCurrentDateRFC822(char *formattedDate, int bufSize);
};