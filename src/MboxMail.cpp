
#include "stdafx.h"
#include <afxtempl.h>
#include "mboxview.h"
#include "MboxMail.h"

_int64 MboxMail::s_curmap = 0;
_int64 MboxMail::s_step = 0;
const CUPDUPDATA* MboxMail::pCUPDUPData = NULL;

bool MboxMail::b_mails_sorted = false;
int MboxMail::b_mails_which_sorted = 0;  // order by file offset
CArray<MboxMail*, MboxMail*> MboxMail::s_mails_ref;
CArray<MboxMail*, MboxMail*> MboxMail::s_mails;
_int64 MboxMail::s_fSize = 0;
_int64 MboxMail::s_oSize = 0;
CString MboxMail::s_path;

#define EATNL() while( p < e && *p++ != '\n' );
#define EATNLh() while( p < eh && *p++ != '\n' );
inline void GetDetail(char *&p, char *e, CString &detail)
{
	char *s = p;
	EATNL();
	char *p1 = p - 1;
	while (s <= p1 && g_tu.IsSpace(*s)) //(*s == ' ' || *s == '\t' || *s == '\r' || *s == '\n') )
		s++;
	while (p1 >= s && g_tu.IsSpace(*p1)) //(*(p1-1) == ' ' || *(p1-1) == '\t' || *(p1-1) == '\r' || *(p1-1) == '\n') )
		p1--;
	p1++;
	char *ss = detail.GetBufferSetLength(p1 - s);
	strncpy(ss, s, p1 - s);
}

#include "DateParser.h"

/*void MboxMail::Parse(LPCSTR path)
{
Destroy();
MboxMail::s_path = path;
HANDLE hFile = CreateFile(path, GENERIC_READ, 0, NULL,
OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
if( hFile == INVALID_HANDLE_VALUE )
return;

DWORD fSize = GetFileSize(hFile, NULL);
HANDLE hFileMap = CreateFileMapping(hFile, NULL, PAGE_READONLY, 0, 0, NULL);
if( hFileMap == NULL ) {
CloseHandle( hFile );
return;
}
register char *p = (char *)MapViewOfFile( hFileMap, FILE_MAP_READ, 0, 0, 0 );
char *orig = p;
char *e = p + fSize;
if( p == NULL || fSize < 10) {
CloseHandle( hFileMap );
CloseHandle( hFile );
return;
}
#ifdef _DEBUG
DWORD tc = GetTickCount();
#endif
CString to, from, subject, date;
bool	bTo = true, bFrom = true, bSubject = true, bDate = true;
char *msgStart = NULL;
int recv = TRUE;
while( p < e - 4 ) {
DWORD *dp = (DWORD *)p;
//		if( strncmp(p, "From ", 4) == 0 ) { // "From "
if( *dp == 0x6d6f7246 && p[4] == ' ') { // "From "
if( msgStart == NULL )
msgStart = p;
else {
if( !(bTo && bFrom && bSubject) ) {
MboxMail *m = new MboxMail();
m->m_startOff = msgStart - orig;
m->m_length = p - msgStart - 1;
m->m_to = to;
m->m_from = from;
m->m_subj = subject;
m->m_date= date;
m->m_recv = recv;
s_mails.Add(m);
}
msgStart = p;
to.Empty();
from.Empty();
subject.Empty();
date.Empty();
bTo = bFrom = bSubject = bDate = true;
}
p += 5;
EATNL();
} else
//		if( bSubject && strncmp(p, "From: ", 6) == 0 ) {
if( bFrom && *dp == 0x6d6f7246 && *(WORD*)(p+4) == 0x203a ) { // "From: "
bFrom = false;
p+=6;
GetDetail(p, e, from);
} else
//		if( bSubject && strncmp(p, "To: ", 4) == 0 ) {
if( bTo && *dp == 0x203a6f54 ) { //&& strncmp(p, "To: ", 4) == 0 ) {
bTo = false;
bTo = false;
p += 4;
GetDetail(p, e, to);
} else
//		if( bSubject && strncmp(p, "Subject: ", 5) == 0 ) {
if( bSubject && *dp == 0x6a627553 && dp[1] == 0x3a746365 && p[8] == ' ' ) { // "Subject: "
bSubject = false;
p += 9;
GetDetail(p, e, subject);
} else
//		if( bSubject && strncmp(p, "Received: ", 10) == 0 ) {
if( bDate && *dp == 0x65636552 && dp[1] == 0x64657669 && *(WORD*)(p+8) == 0x203a ) { // "Received: "
bDate = false;
p += 10;
char *p1 = p+strcspn(p, ";\n");
while( *p1 == '\n' && (*(p1+1) == ' ' || *(p1+1) == '\t') ) {
p1++;
p1 += strcspn(p1, ";\n");
}
if( *p1 == ';' ) {
p1++;
while( g_tu.IsSpace(*p1) )
p1++;
char *el = strchr(p1, '\n');
while( g_tu.IsSpace(*el) )
el--;
el++;
char *d = date.GetBufferSetLength(el-p1);
strncpy(d, p1, el-p1);
SYSTEMTIME tm;
DateParser::parseRFC822Date(d, &tm);
date = CTime(tm).Format("%d/%m/%y %H:%I");
recv = TRUE;
EATNL();
} else {
EATNL();
}
} else
if( bDate && *dp == 0x65746144 && *(WORD*)(p+4) == 0x203a ) { // "Date: "
bDate = false;
p += 6;
GetDetail(p, e, date);
SYSTEMTIME tm;
DateParser::parseRFC822Date(date, &tm);
date = CTime(tm).Format("%d/%m/%y %H:%I");
recv = FALSE;
} else
EATNL();
}
if( msgStart != NULL && msgStart != p ) {
MboxMail *m = new MboxMail();
//		TRACE("start: %d length: %d\n", msgStart - orig, p - msgStart);
m->m_startOff = msgStart - orig;
m->m_length = p - msgStart - 1;
m->m_to = to;
m->m_from = from;
m->m_subj = subject;
m->m_date = date;
m->m_recv = recv;
s_mails.Add(m);
}
#ifdef _DEBUG
tc = (GetTickCount() - tc);
TRACE("Parse Took %d:%d %d\n", (tc/1000)/60, (tc/1000)%60, tc);
#endif
UnmapViewOfFile( orig );
CloseHandle( hFileMap );
CloseHandle( hFile );
}
*/


CString GetDateFormat(int i)
{
	CString format;
	switch (i) {
	case 1:
		format = "%m/%d/%Y %H:%M";
		break;
	case 2:
		format = "%Y/%m/%d %H:%M";
		break;
	default:
		format = "%d/%m/%Y %H:%M";
		break;
	}
	return format;
}

UINT MboxMail::Str2PageCode(const  char* PageCodeStr)
{
	//UINT CodePage = 0;
	UINT CodePage = CP_UTF8;

	if (_stricmp(PageCodeStr, "utf-8") == 0) {
		CodePage = CP_UTF8; // 65001
	}
	else if (_stricmp(PageCodeStr, "ISO-8859-1") == 0) {  
		CodePage = 28591;  // Western European 
	}
	else if (_stricmp(PageCodeStr, "ISO-8859-2") == 0) {
		CodePage = 28592;  // Central  European 
	}
	else if (_stricmp(PageCodeStr, "ISO-8859-3") == 0) {
		CodePage = 28593;  // Latin 3
	}
	else if (_stricmp(PageCodeStr, "ISO-8859-4") == 0) {
		CodePage = 28594;  // Baltic
	}
	else if (_stricmp(PageCodeStr, "ISO-8859-5") == 0) {
		CodePage = 28595;  // Cyrillic
	}
	else if (_stricmp(PageCodeStr, "ISO-8859-6") == 0) {
		CodePage = 28596;  // Arabic
	}
	else if (_stricmp(PageCodeStr, "ISO-8859-7") == 0) {
		CodePage = 28597;  // Greek
	}
	else if (_stricmp(PageCodeStr, "ISO-8859-8") == 0) {
		CodePage = 28598;  // Hebrew
	}
	else if (_stricmp(PageCodeStr, "ISO-8859-9") == 0) {
		CodePage = 28599;  // Turkish
	}
	else if (_stricmp(PageCodeStr, "ISO-8859-13") == 0) {
		CodePage = 28603;  // Estonian
	}
	else if (_stricmp(PageCodeStr, "ISO-8859-15") == 0) {
		CodePage = 28605;  // Latin 9
	}
	else if (_stricmp(PageCodeStr, "US-ASCII") == 0) {
		CodePage = 20127;  // US-ASCII 7
	}
	else if (_stricmp(PageCodeStr, "ISO-2022-JP") == 0) {
		CodePage = 50220;  // Japanese with no halfwidth Katakana; Japanese (JIS) , maybe ??
	}
	else if (_stricmp(PageCodeStr, "csISO2022JP") == 0) {
		CodePage = 50221;  // Japanese with halfwidth Katakana; Japanese (JIS-Allow 1 byte Kana)
	}
	else if (_stricmp(PageCodeStr, "ISO-2022-KR") == 0) {
		CodePage = 50225;  // Korean
	}
	return CodePage;
}

void MboxMail::Str2Ansi(CString &res, UINT CodePage)
{
	int len = res.GetLength() * 2 + 1;
	LPWSTR buff = (LPWSTR)malloc(len);  // or  we could call MultiByteToWideChar first to get the required length
	int len1 = MultiByteToWideChar(CodePage, 0, res, res.GetLength(), buff, len);
	if (len1 == 0) {
		free(buff);
		// error - implement error log file
		const DWORD error = ::GetLastError();
		return;
	}
	char * buff1 = (char *)malloc(len1 + 1); // or could  call WideCharToMultiByte first to get the required length
	int len2 = WideCharToMultiByte(CP_ACP, 0, buff, len1, buff1, len1 + 1, NULL, NULL);
	if (len2 == 0) {
		free(buff);
		free(buff1);
		// error - implement error log file
		const DWORD error = ::GetLastError();
		/*ERROR_INSUFFICIENT_BUFFER.A supplied buffer size was not large enough, or it was incorrectly set to NULL.
		ERROR_INVALID_FLAGS.The values supplied for flags were not valid.
		ERROR_INVALID_PARAMETER.Any of the parameter values was invalid.
		ERROR_NO_UNICODE_TRANSLATION.Invalid Unicode was found in a string.*/
		return;
	}
	buff1[len2] = 0;
	res = buff1;
	free(buff);
	free(buff1);

}

#include "MimeCode.h"

CString DecodeString(CString &subj)
{
	CFieldCodeText tfc;

	// Not critical. Should optimize this function and Str2Ansi to reduce memory allocations/deallocations
	tfc.SetInput(subj.GetBuffer(), subj.GetLength(), false);
	int outputLen = tfc.GetOutputLength();
	if (outputLen > 0) {

		unsigned char *outBuf = (unsigned char*)malloc(outputLen + 2);
		int decodeLen = tfc.GetOutput(outBuf, outputLen);
		outBuf[decodeLen] = 0;
		CString str(outBuf);
		free(outBuf);

		UINT CodePage = MboxMail::Str2PageCode(tfc.GetCharset());
		if (CodePage > 0) {
			MboxMail::Str2Ansi(str, CodePage);
			return str;
		}
		else
			return subj;
	}
	else
		return subj;
}

char szFrom5[] = "From ";
char szFrom6[] = "\nFrom ";
char	*g_szFrom;
int		g_szFromLen;

bool MboxMail::Process(register char *p, DWORD bufSize, _int64 startOffset, bool bEml)
{
	char *orig = p;
	register char *e = p + bufSize - 1;
	if (p == NULL || bufSize < 10) {
		return false;
	}
	CString contentDisposition = "Content-Disposition: attachment";
	int iFormat = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "format");
	CString format = GetDateFormat(iFormat);
	CString to, from, subject, date;
	time_t tdate = 0;
	bool	bTo = true, bFrom = true, bSubject = true, bDate = true;
	char *msgStart = NULL;
	int recv = TRUE;
	int curstep = (int)(startOffset / (s_step?s_step:1));
	if (bEml) { // duplicated code and taken if out of loop for speed
		while (p < e - 4) {
			_int64 curOff = startOffset + (p - orig);
			if ((int)(curOff / s_step) > curstep) {
				if (pCUPDUPData->ShouldTerminate())
					break;
				int step = (int)(curOff / s_step);
				pCUPDUPData->SetProgress(step);
			}
			//		if( strncmp(p, "From ", 4) == 0 ) { // "From "
			if (bEml || (*(DWORD*)p == 0x6d6f7246 && p[4] == ' ')) { // "From "
				if (msgStart == NULL)
					msgStart = p;
				else {
					if (!(bTo && bFrom && bSubject)) {
						if (!(to.IsEmpty() && subject.IsEmpty() && date.IsEmpty())) {
							MboxMail *m = new MboxMail();
							m->m_startOff = startOffset + (_int64)(msgStart - orig);
							m->m_length = p - msgStart - 1;
							m->m_to = DecodeString(to);
							m->m_from = DecodeString(from);
							m->m_subj = DecodeString(subject);
							m->m_timeDate = tdate;
							m->m_recv = recv;
							int pos = g_tu.BMHSearchW((unsigned char *)msgStart, p - msgStart - 1, (unsigned char *)(LPCSTR)contentDisposition, contentDisposition.GetLength(), false);
							if (pos >= 0)
								m->m_hasAttachments = 1;
							s_mails.Add(m);
						}
					}
					msgStart = p;
					to.Empty();
					from.Empty();
					subject.Empty();
					date.Empty();
					tdate = 0;
					bTo = bFrom = bSubject = bDate = true;
				}
				p += 5;
				if (!bEml) {
					EATNL();
				}
				else
					bEml = false;
			}
			else
				//		if( bSubject && strncmp(p, "From: ", 6) == 0 ) {
				if (bFrom && *(DWORD*)p == 0x6d6f7246 && *(WORD*)(p + 4) == 0x203a) { // "From: "
					bFrom = false;
					p += 6;
					GetDetail(p, e, from);
				}
				else
					//		if( bSubject && strncmp(p, "To: ", 4) == 0 ) {
					if (bTo && *(DWORD*)p == 0x203a6f54) { //&& strncmp(p, "To: ", 4) == 0 ) {
						bTo = false;
						bTo = false;
						p += 4;
						GetDetail(p, e, to);
					}
					else
						//		if( bSubject && strncmp(p, "Subject: ", 5) == 0 ) {
						if (bSubject && *(DWORD*)p == 0x6a627553 && *(DWORD*)(p + 4) == 0x3a746365 && p[8] == ' ') { // "Subject: "
							bSubject = false;
							p += 9;
							GetDetail(p, e, subject);
						}
						else
							//		if( bSubject && strncmp(p, "Received: ", 10) == 0 ) {
							if (bDate && *(DWORD*)p == 0x65636552 && *(DWORD*)(p + 4) == 0x64657669 && *(WORD*)(p + 8) == 0x203a) { // "Received: "
								bDate = false;
								p += 10;
								char *p1 = p + strcspn(p, ";\n");
								while (*p1 == '\n' && (*(p1 + 1) == ' ' || *(p1 + 1) == '\t')) {
									p1++;
									p1 += strcspn(p1, ";\n");
								}
								if (*p1 == ';') {
									p1++;
									while (g_tu.IsSpace(*p1))
										p1++;
									char *el = strchr(p1, '\n');
									if (el != NULL) { // test for git change
										char *el1 = el;
										char *el2 = el;
										while (g_tu.IsSpace(*el))
											el--;
										el++;
										int l = el - p1, l1 = 0;
										if (*(el1 + 1) == ' ') { // for multiline dates
											el1 = el1 + 1;
											el2 = strchr(el1, '\n');
											while (el1 < el2 && g_tu.IsSpace(*el1)) //strip leading spaces
												el1++;
											while (el2 < el1 && g_tu.IsSpace(*el2)) // strip trailing spaces
												el2--;
											l1 = el2 - el1 + 1;
										}
										char *d = date.GetBufferSetLength(l + l1);
										strncpy(d, p1, l);
										if (el1 != el2) {
											d[l] = ' '; // add a space for parser
											strncpy(d + l + 1, el1, el2 - el1);
										}
										d[l + l1] = 0;
										SYSTEMTIME tm;
										if (DateParser::parseRFC822Date(d, &tm)) {
											CTime tt(tm);
											date = tt.Format(format);
											tdate = tt.GetTime();
										}
										recv = TRUE;
										EATNL();
									}
								}
								else {
									EATNL();
								}
							}
							else
								if (bDate && *(DWORD*)p == 0x65746144 && *(WORD*)(p + 4) == 0x203a) { // "Date: "
									bDate = false;
									p += 6;
									GetDetail(p, e, date);
									SYSTEMTIME tm;
									if (DateParser::parseRFC822Date(date, &tm)) {
										CTime tt(tm);
										date = tt.Format(format);
										tdate = tt.GetTime();
									}
									recv = FALSE;
								}
								else
									EATNL();
		}
	}
	while (p < e - 4) {
		try {
			if (*(DWORD*)p == 0x6d6f7246 && p[4] == ' ') { // "From "
				if (msgStart == NULL)
					msgStart = p;
				else {
					if (!(bTo && bFrom && bSubject)) {
						if (!(to.IsEmpty() && subject.IsEmpty() && date.IsEmpty())) {
							MboxMail *m = new MboxMail();
							m->m_startOff = startOffset + (_int64)(msgStart - orig);
							m->m_length = p - msgStart;
							m->m_to = DecodeString(to);
							m->m_from = DecodeString(from);
							m->m_subj = DecodeString(subject);
							m->m_timeDate = tdate;
							m->m_recv = recv;
							int pos = g_tu.BMHSearch((unsigned char *)msgStart, m->m_length, (unsigned char *)(LPCSTR)contentDisposition, contentDisposition.GetLength(), 1);
							if (pos >= 0)
								m->m_hasAttachments = 1;
							s_mails.Add(m);
							_int64 curOff = startOffset + (p - orig);
							if ((int)(curOff / s_step) > curstep) {
								if (pCUPDUPData->ShouldTerminate())
									break;
								int step = (int)(curOff / s_step);
								pCUPDUPData->SetProgress(step);
							}
						}
					}
					msgStart = p;
					to.Empty();
					from.Empty();
					subject.Empty();
					date.Empty();
					tdate = 0;
					bTo = bFrom = bSubject = bDate = true;
				}
				p += 5;
				EATNL();
			}
			else
				//		if( bSubject && strncmp(p, "From: ", 6) == 0 ) {
				if (bFrom && *(DWORD*)p == 0x6d6f7246 && *(WORD*)(p + 4) == 0x203a) { // "From: "
					bFrom = false;
					p += 6;
					GetDetail(p, e, from);
				}
				else
					//		if( bSubject && strncmp(p, "To: ", 4) == 0 ) {
					if (bTo && *(DWORD*)p == 0x203a6f54) { //&& strncmp(p, "To: ", 4) == 0 ) {
						bTo = false;
						bTo = false;
						p += 4;
						GetDetail(p, e, to);
					}
					else
						//		if( bSubject && strncmp(p, "Subject: ", 5) == 0 ) {
						if (bSubject && *(DWORD*)p == 0x6a627553 && *(DWORD*)(p + 4) == 0x3a746365 && p[8] == ' ') { // "Subject: "
							bSubject = false;
							p += 9;
							GetDetail(p, e, subject);
						}
						else
							//		if( bSubject && strncmp(p, "Received: ", 10) == 0 ) {
							if (bDate && *(DWORD*)p == 0x65636552 && *(DWORD*)(p + 4) == 0x64657669 && *(WORD*)(p + 8) == 0x203a) { // "Received: "
								bDate = false;
								p += 10;
								char *p1 = p + strcspn(p, ";\n");
								while (*p1 == '\n' && (*(p1 + 1) == ' ' || *(p1 + 1) == '\t')) {
									p1++;
									p1 += strcspn(p1, ";\n");
								}
								if (*p1 == ';') {
									p1++;
									while (g_tu.IsSpace(*p1))
										p1++;
									char *el = strchr(p1, '\n');
									if (el != NULL) {
										char *el1 = el;
										char *el2 = el;
										while (g_tu.IsSpace(*el))
											el--;
										el++;
										int l = el - p1, l1 = 0;
										if (*(el1 + 1) == ' ') { // for multiline dates
											el1 = el1 + 1;
											el2 = strchr(el1, '\n');
											while (el1 < el2 && g_tu.IsSpace(*el1)) //strip leading spaces
												el1++;
											while (el2 < el1 && g_tu.IsSpace(*el2)) //strip trailing spaces
												el2--;
											l1 = el2 - el1 + 1;
										}
										char *d = date.GetBufferSetLength(l + l1);
										strncpy(d, p1, l);
										if (el1 != el2) {
											d[l] = ' '; // add a space for parser
											strncpy(d + l + 1, el1, el2 - el1);
										}
										d[l + l1] = 0;
										SYSTEMTIME tm;
										if (DateParser::parseRFC822Date(d, &tm)) {
											CTime tt(tm);
											date = tt.Format(format);
											tdate = tt.GetTime();
										}
										recv = TRUE;
										EATNL();
									}
								}
								else {
									EATNL();
								}
							}
							else
								if (bDate && *(DWORD*)p == 0x65746144 && *(WORD*)(p + 4) == 0x203a) { // "Date: "
									bDate = false;
									p += 6;
									GetDetail(p, e, date);
									SYSTEMTIME tm;
									if (DateParser::parseRFC822Date(date, &tm)) {
										CTime tt(tm);
										date = tt.Format(format);
										tdate = tt.GetTime();
									}
									recv = FALSE;
								}
								else
								{
									EATNL();
								}
		}
		catch (CMemoryException *) {
			// catch when buffer ends
		}
	}

	// should save state for cross boundary emails
	if (msgStart != NULL && msgStart != p) {
		MboxMail *m = new MboxMail();
		//		TRACE("start: %d length: %d\n", msgStart - orig, p - msgStart);
		m->m_startOff = startOffset + (_int64)(msgStart - orig);
		m->m_length = p - msgStart;
		m->m_to = DecodeString(to);
		m->m_from = DecodeString(from);
		m->m_subj = DecodeString(subject);
		m->m_recv = recv;
		m->m_timeDate = tdate;
		s_mails.Add(m);
	}
	return true;
}

#include <algorithm>
using namespace std;

bool cmpAllMbox(bool desc) {
	int outofOrderCnt = 0;
	for (int m = 0; m < (MboxMail::s_mails.GetSize() - 1); m++) {
		MboxMail *mb = MboxMail::s_mails[m];
		MboxMail *mb_next = MboxMail::s_mails[m + 1];
		if (desc) {
			if (mb->m_startOff < mb_next->m_startOff) {
				outofOrderCnt++;
			}
		}
		else
		{
			if (mb->m_startOff >= mb_next->m_startOff) {
				outofOrderCnt++;
			}
		}
	}
	if (outofOrderCnt > 0) {
		int deb = 1;  // setup breakpoint
		return false;
	}
	return true;
}

bool cmpMbox(MboxMail* a, MboxMail *b) {
	return a->m_startOff < b->m_startOff;
}
void MboxMail::Parse(LPCSTR path)
{
	Destroy();
	MboxMail::s_path = path;
	bool	bEml = false;
	int l = strlen(path);
	if (l > 4 && _strnicmp(path + l - 4, ".eml", 4) == 0)
		bEml = true;
	HANDLE hFile = CreateFile(path, GENERIC_READ, 0, NULL,
		OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL);
	if (hFile == INVALID_HANDLE_VALUE)
		return;

	unsigned _int64 fSize = 0;
	MboxMail::s_oSize = MboxMail::s_fSize;
	MboxMail::s_fSize = fSize;
	LARGE_INTEGER li;
	GetFileSizeEx(hFile, &li);
	fSize = li.QuadPart;
	s_step = fSize / 100;
	int mappingsInFile = (int)(fSize / (_int64)MAPPING_SIZE);
	HANDLE hFileMap = CreateFileMapping(hFile, NULL, PAGE_READONLY, (DWORD)(fSize >> 32), (DWORD)fSize, NULL);
	if (hFileMap == NULL) {
		CloseHandle(hFile);
		return;
	}
#ifdef _DEBUG
	DWORD tc = GetTickCount();
#endif
	g_szFrom = szFrom5;
	g_szFromLen = 5;
	TRACE("fSize = %lld\n", fSize);
	pCUPDUPData->SetProgress(_T("Parsing file..."), 0);
	for (int i = 0; i <= mappingsInFile && !pCUPDUPData->ShouldTerminate(); i++) {
		s_curmap = (_int64)i * (_int64)MAPPING_SIZE;
		DWORD bufSize = ((fSize - s_curmap < MAPPING_SIZE) ? (int)(fSize - s_curmap) : MAPPING_SIZE);
		TRACE("offset=%lld, bufsize=%ld, fSize-curmap=%lld, end=%lld\n", s_curmap, bufSize, fSize - s_curmap, s_curmap + bufSize);
		char * p = (char *)MapViewOfFileEx(hFileMap, FILE_MAP_READ, (DWORD)(s_curmap >> 32), (DWORD)s_curmap, bufSize, NULL);
		MboxMail::Process(p, bufSize, s_curmap, bEml);
		UnmapViewOfFile(p);
	}
#ifdef _DEBUG
	tc = (GetTickCount() - tc);
	CString out;
	out.Format("Parse Took %d:%d %d\n", tc / 60000, (tc / 1000) % 60, tc);
	OutputDebugString(out);
	out.Format("Found %d mails\n", MboxMail::s_mails.GetSize());
	OutputDebugString(out);
//	TRACE("Parse Took %d:%d %d\n", tc / 60000, (tc / 1000) % 60, tc);
//	TRACE("Found %d mails\n", MboxMail::s_mails.GetSize());
#endif
	CloseHandle(hFileMap);
	CloseHandle(hFile);

	// Set the breakpoint under debugger to detect out of order emails
	//cmpAllMbox(false);

	// sort the mails by offset, they should already be ordered by file offset anyway
	sort(MboxMail::s_mails.GetData(), MboxMail::s_mails.GetData() + MboxMail::s_mails.GetSize(), [](MboxMail* o1, MboxMail* o2)
	{
		return cmpMbox(o1, o2);
	});

	//cmpAllMbox(false);

#if 0
	CStdioFile fp("mboxmails.txt", CFile::modeCreate | CFile::modeWrite | CFile::typeText);
	for (int m = 0; m < MboxMail::s_mails.GetSize(); m++) {
		CString s;
		MboxMail *mb = MboxMail::s_mails[m];
		CTime t(mb->m_timeDate);
		s.Format("From: %s\nTo: %s\nDate: %s\nSubject: %s\nSize:%d\n", 
			(LPCSTR)mb->m_from, (LPCSTR)mb->m_to, (LPCSTR)t.Format("%Y%m%d"), (LPCSTR)mb->m_subj, mb->m_length);
		fp.WriteString(s);
	}
	fp.Close();
#endif
}

#include <algorithm>

bool sortByDate(MboxMail *cr1, MboxMail *cr2) { return (cr1->m_timeDate < cr2->m_timeDate); }
bool sortByDateDesc(MboxMail *cr1, MboxMail *cr2) { return (cr1->m_timeDate > cr2->m_timeDate); }

bool sortByFrom(MboxMail *cr1, MboxMail *cr2) {
	int cmp = cr1->m_from.Compare(cr2->m_from);
	if (cmp == 0)
		return (cr1->m_timeDate > cr2->m_timeDate); // make stable sort
	return (cmp < 0);
}
bool sortByFromDesc(MboxMail *cr1, MboxMail *cr2) {
	int cmp = cr1->m_from.Compare(cr2->m_from);
	if (cmp == 0)
		return (cr1->m_timeDate > cr2->m_timeDate); // make stable sort
	return (cmp > 0);
}

bool sortByTo(MboxMail *cr1, MboxMail *cr2) {
	int cmp = cr1->m_to.Compare(cr2->m_to);
	if (cmp == 0)
		return (cr1->m_timeDate < cr2->m_timeDate); // make stable sort
	return (cmp < 0);
}
bool sortByToDesc(MboxMail *cr1, MboxMail *cr2) {
	int cmp = cr1->m_to.Compare(cr2->m_to);
	if (cmp == 0)
		return (cr1->m_timeDate < cr2->m_timeDate); // make stable sort
	return (cmp > 0);
}

bool sortBySubject(MboxMail *cr1, MboxMail *cr2) {
	unsigned char *subj1 = (unsigned char *)(LPCSTR)cr1->m_subj;
	unsigned char *subj2 = (unsigned char *)(LPCSTR)cr2->m_subj;
	int subjlen1 = cr1->m_subj.GetLength();
	int subjlen2 = cr2->m_subj.GetLength();

	// Not 100% correct without using from and to fields
	if (subj1[0] == 'R') {
		if (subjlen1 >= 4) {
			if ((strncmp((char*)subj1, "Re: ", 4) == 0) || (strncmp((char*)subj1, "RE: ", 4) == 0)) {
				subj1 += 4; subjlen1 -= 4;
			}
		}
	}
	else if (subj1[0] == 'F') {
		if (subjlen1 >= 5) {
			if ((strncmp((char*)subj1, "Fwd: ", 5) == 0) || (strncmp((char*)subj1, "FWD: ", 5) == 0)) {
				subj1 += 5; subjlen1 -= 5;
			}
		}
	}
	if (subj2[0] == 'R') {
		if (subjlen2 >= 4) {
			if ((strncmp((char*)subj2, "Re: ", 4) == 0) || (strncmp((char*)subj2, "RE: ", 4) == 0)) {
				subj2 += 4; subjlen2 -= 4;
			}
		}
	}
	else if (subj2[0] == 'F') {
		if (subjlen2 >= 5) {
			if ((strncmp((char*)subj2, "Fwd: ", 5) == 0) || (strncmp((char*)subj2, "FWD: ", 5) == 0)) {
				subj2 += 5; subjlen2 -= 5;
			}
		}
	}

	if ((subjlen1 == subjlen2) && (strncmp((char*)subj1, (char*)subj2, subjlen1) == 0))
	{
		return (cr1->m_timeDate < cr2->m_timeDate); // make stable sort
	}
	else
	{
		if (strcmp((char*)subj1, (char*)subj2) < 0)
			return true;
		else
			return false;
	}
}
bool sortBySubjectDesc(MboxMail *cr1, MboxMail *cr2) {
	unsigned char *subj1 = (unsigned char *)(LPCSTR)cr1->m_subj;
	unsigned char *subj2 = (unsigned char *)(LPCSTR)cr2->m_subj;
	int subjlen1 = cr1->m_subj.GetLength();
	int subjlen2 = cr2->m_subj.GetLength();

	// Not 100% correct  without using from and to fields
	if (subj1[0] == 'R') {
		if (subjlen1 >= 4) {
			if ((strncmp((char*)subj1, "Re: ", 4) == 0) || (strncmp((char*)subj1, "RE: ", 4) == 0)) {
				subj1 += 4; subjlen1 -= 4;
			}
		}
	}
	else if (subj1[0] == 'F') {
		if (subjlen1 >= 5) {
			if ((strncmp((char*)subj1, "Fwd: ", 5) == 0) || (strncmp((char*)subj1, "FWD: ", 5) == 0)) {
				subj1 += 5; subjlen1 -= 5;
			}
		}
	}
	if (subj2[0] == 'R') {
		if (subjlen2 >= 4) {
			if ((strncmp((char*)subj2, "Re: ", 4) == 0) || (strncmp((char*)subj2, "RE: ", 4) == 0)) {
				subj2 += 4; subjlen2 -= 4;
			}
		}
	}
	else if (subj2[0] == 'F') {
		if (subjlen2 >= 5) {
			if ((strncmp((char*)subj2, "Fwd: ", 5) == 0) || (strncmp((char*)subj2, "FWD: ", 5) == 0)) {
				subj2 += 5; subjlen2 -= 5;
			}
		}
	}

	if ((subjlen1 == subjlen2) && (strncmp((char*)subj1, (char*)subj2, subjlen1) == 0))
	{
		return (cr1->m_timeDate < cr2->m_timeDate); // make stable sort
	}
	else
	{
		if (strcmp((char*)subj1, (char*)subj2) > 0)
			return true;
		else
			return false;
	}
}
bool sortBySize(MboxMail *cr1, MboxMail *cr2) {
	if (cr1->m_length == cr2->m_length)
		return (cr1->m_timeDate < cr2->m_timeDate); // make stable sort
	return (cr1->m_length < cr2->m_length);
}
bool sortBySizeDesc(MboxMail *cr1, MboxMail *cr2) {
	if (cr1->m_length == cr2->m_length)
		return (cr1->m_timeDate < cr2->m_timeDate); // make stable sort
	return (cr1->m_length > cr2->m_length);
}

void MboxMail::SortByDate(CArray<MboxMail*, MboxMail*> *s_m, bool bDesc)
{
	if (s_m == 0) s_m = &MboxMail::s_mails;
	std::stable_sort(s_m->GetData(), s_m->GetData() + s_m->GetSize(), bDesc ? sortByDateDesc : sortByDate);
}
void MboxMail::SortByFrom(CArray<MboxMail*, MboxMail*> *s_m, bool bDesc)
{
	if (s_m == 0) s_m = &MboxMail::s_mails;
	std::sort(s_m->GetData(), s_m->GetData() + s_m->GetSize(), bDesc ? sortByFromDesc : sortByFrom);
}
void MboxMail::SortByTo(CArray<MboxMail*, MboxMail*> *s_m, bool bDesc)
{
	if (s_m == 0) s_m = &MboxMail::s_mails;
	std::sort(s_m->GetData(), s_m->GetData() + s_m->GetSize(), bDesc ? sortByToDesc : sortByTo);
}
void MboxMail::SortBySubject(CArray<MboxMail*, MboxMail*> *s_m, bool bDesc)
{
	if (s_m == 0) s_m = &MboxMail::s_mails;
	std::sort(s_m->GetData(), s_m->GetData() + s_m->GetSize(), bDesc ? sortBySubjectDesc : sortBySubject);
}
void MboxMail::SortBySize(CArray<MboxMail*, MboxMail*> *s_m, bool bDesc)
{
	if (s_m == 0) s_m = &MboxMail::s_mails;
	std::sort(s_m->GetData(), s_m->GetData() + s_m->GetSize(), bDesc ? sortBySizeDesc : sortBySize);
}





