
#include "stdafx.h"
#include <afxtempl.h>
#include "mboxview.h"
#include "MboxMail.h"

UINT charset2Id(const char *char_set);
BOOL id2charset(UINT id, std::string &charset);
char * ParseContent(MboxMail *mail, char *startPos, char *endPos);

unsigned char *MboxMail::m_pbOutput = 0;
int MboxMail::m_maxOutput = 0;

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

inline void GetLine(char *p, char *e, CString &line)
{
	char *s = p;
	char *pnl = p;
	while (pnl < e && *pnl++ != '\n');
	char *p1 = pnl - 1;
	while (s <= p1 && ((*s == ' ') || (*s == '\t'))) 
		s++;
	while (p1 >= s && ((*p1 == ' ') || (*p1 == '\t')))
		p1--;
	p1++;
	char *ss = line.GetBufferSetLength(p1 - s);
	::strncpy(ss, s, p1 - s);
}

#include "DateParser.h"

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
	UINT CodePage = 0;

	CodePage = charset2Id(PageCodeStr);

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

CString DecodeString(CString &subj, CString &charset, UINT &charsetId)
{
	CFieldCodeText tfc;
	tfc.SetInput(subj.GetBuffer(), subj.GetLength(), false);
	int outputLen = tfc.GetOutputLength();
	if (outputLen > 0) {
		int maxOutputLen = 2 * outputLen + 2;
		unsigned char *outBuf = (unsigned char*)malloc(maxOutputLen);
		int decodeLen = tfc.GetOutput(outBuf, maxOutputLen);
		outBuf[decodeLen] = 0;
		CString str(outBuf);
		free(outBuf);
		charset = tfc.GetCharset();
		UINT CodePage = MboxMail::Str2PageCode(tfc.GetCharset());
		charsetId = CodePage;
		// No remapping for now. Header of message windows should show text properly
		return str;

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

CString MboxMail::GetBody()
{
	CString res;
	CFile fp;
	if (fp.Open(s_path, CFile::modeRead | CFile::shareDenyWrite)) {
		char *p = res.GetBufferSetLength(m_length);
		TRACE("offset = %lld\n", m_startOff);
		fp.Seek(m_startOff, SEEK_SET);
		fp.Read(p, m_length);
		char *ms = strchr(p, '\n'); //"-Version: 1.0");
		if (ms) {
			BOOL bAddCR = FALSE;
			if (*(ms - 1) != '\r')
				bAddCR = TRUE;
			res = res.Mid(ms - p + 1); // - 4);
			if (bAddCR) // for correct mime parsing
				res.Replace("\n", "\r\n");
		}
	}
	else
	{
		DWORD err = GetLastError();
		TRACE("Open Mail File failed err=%ld\n", err);
	}

	return res;
};

char szFrom5[] = "From ";
char szFrom6[] = "\nFrom ";
char	*g_szFrom;
int		g_szFromLen;

bool MboxMail::Process(register char *p, DWORD bufSize, _int64 startOffset,  bool bFirstView, bool bLastView, _int64 &lastStartOffset, bool bEml)
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
	bool	bTo = true, bFrom = true, bSubject = true, bDate = true;  // indicates not found, false means found 
	char *msgStart = NULL;
	int recv = TRUE;
	int curstep = (int)(startOffset / (s_step?s_step:1));

	if (bEml && bFirstView) {
		msgStart = p;
		p += 5;
		bEml = false;  // not used below and not needed :)
	}
	while (p < e - 4) 
	{
		try {
			if (*(DWORD*)p == 0x6d6f7246 && p[4] == ' ') // "From "  marks beginning of the next mail
			{
				if (msgStart == NULL)
					msgStart = p;
				else {
					if ((bTo == false) || (bFrom == false) ||  (bSubject == false)) 
					{
						if (!from.IsEmpty() || !to.IsEmpty() ||  !subject.IsEmpty() ||  !date.IsEmpty())
						{
							MboxMail *m = new MboxMail();
							m->m_startOff = startOffset + (_int64)(msgStart - orig);
							m->m_length = p - msgStart;
							m->m_to = DecodeString(to, m->m_to_charset, m->m_to_charsetId);
							m->m_from = DecodeString(from, m->m_from_charset, m->m_from_charsetId);
							m->m_subj = DecodeString(subject, m->m_subj_charset, m->m_subj_charsetId);
							m->m_timeDate = tdate;
							m->m_recv = recv;
							m->m_headLength = 0;  // set in ParseContent()

							MailBodyContent body;
							char *bodyStart = msgStart;
							char *bodyEnd = msgStart + m->m_length;
							char *nextStart = ParseContent(m, bodyStart, bodyEnd);
							if (m->m_ContentDetailsArray.size() == 0) {
								int deb = 1;
							}

							s_mails.Add(m);

							int index = s_mails.GetCount() - 1;
							//m->DumpMailBox(m, index);
							//NListView::DumpItemDetails(index);

							_int64 curOff = startOffset + (p - orig);
							if ((int)(curOff / s_step) > curstep) {
								if (pCUPDUPData->ShouldTerminate())
									break;
								int step = (int)(curOff / s_step);
								pCUPDUPData->SetProgress(step);
							}

							if (pCUPDUPData->ShouldTerminate())
								break;
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
			{
				//		if( bSubject && strncmp(p, "From: ", 6) == 0 ) {
				if (bFrom && *(DWORD*)p == 0x6d6f7246 && *(WORD*)(p + 4) == 0x203a) { // "From: "
					bFrom = false;
					p += 6;
					GetDetail(p, e, from);
				}
				else
				{
					//		if( bSubject && strncmp(p, "To: ", 4) == 0 ) {
					if (bTo && *(DWORD*)p == 0x203a6f54) { //&& strncmp(p, "To: ", 4) == 0 ) {
						bTo = false;
						bTo = false;
						p += 4;
						GetDetail(p, e, to);
					}
					else
					{
						//		if( bSubject && strncmp(p, "Subject: ", 5) == 0 ) {
						if (bSubject && *(DWORD*)p == 0x6a627553 && *(DWORD*)(p + 4) == 0x3a746365 && p[8] == ' ') // "Subject: "
						{
							bSubject = false;
							p += 9;
							GetDetail(p, e, subject);
							// handle multiline subject
							// Not the most efficient but this should rare
							int cnt = 0;
							while ((p < e - 4) && ((*p == ' ') || (*p == '\t')) && (cnt++ < 10)) 
							{
								CString line;
								CString part;
								GetLine(p, e, line);
								GetDetail(p, e, part);
								subject += part;
							}
						}
						else
						{
							//		if( bSubject && strncmp(p, "Received: ", 10) == 0 ) {
							if (bDate && *(DWORD*)p == 0x65636552 && *(DWORD*)(p + 4) == 0x64657669 && *(WORD*)(p + 8) == 0x203a) // "Received: "
							{
								bDate = false;
								p += 10;
								char *p1 = p + strcspn(p, ";\n");
								while (*p1 == '\n' && (*(p1 + 1) == ' ' || *(p1 + 1) == '\t')) {
									p1++;
									p1 += strcspn(p1, ";\n");
								}
								if (*p1 == ';') 
								{
									p1++;
									while (g_tu.IsSpace(*p1))
										p1++;
									char *el = strchr(p1, '\n');
									if (el != NULL) 
									{
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
							{
								if (bDate && *(DWORD*)p == 0x65746144 && *(WORD*)(p + 4) == 0x203a) // "Date: "
								{
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
						}
					}
				}
			}
		}
		catch (CMemoryException *) {
			// catch when buffer ends
		}
	}

	// should save state for cross boundary emails - resolved 2018
	if (msgStart != NULL && msgStart != p) {
		if (!bLastView) {
			lastStartOffset = startOffset + (_int64)(msgStart - orig);
		}
		else if ((bTo == false) || (bFrom == false) || (bSubject == false))
		{
			if (!from.IsEmpty() || !to.IsEmpty() || !subject.IsEmpty() || !date.IsEmpty())
			{
				MboxMail *m = new MboxMail();
				//		TRACE("start: %d length: %d\n", msgStart - orig, p - msgStart);
				m->m_startOff = startOffset + (_int64)(msgStart - orig);
				m->m_length = p - msgStart;
				m->m_to = DecodeString(to, m->m_to_charset, m->m_to_charsetId);
				m->m_from = DecodeString(from, m->m_from_charset, m->m_from_charsetId);
				m->m_subj = DecodeString(subject, m->m_subj_charset, m->m_subj_charsetId);
				m->m_recv = recv;
				m->m_timeDate = tdate;
				m->m_headLength = 0; // set in ParseContent()

				MailBodyContent body;
				char *bodyStart = msgStart;
				char *bodyEnd = msgStart + m->m_length;
				char *nextStart = ParseContent(m, bodyStart, bodyEnd);

				s_mails.Add(m);

				int index = s_mails.GetCount() - 1;
				//m->DumpMailBox(m, index);
				//NListView::DumpItemDetails(index);
			}
		}
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
	HANDLE hFile = CreateFile(path, GENERIC_READ, FILE_SHARE_READ, NULL,
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
	int mappingsInFile = (int)((fSize - 1) / (_int64)MAPPING_SIZE) + 1;
	HANDLE hFileMap = CreateFileMapping(hFile, NULL, PAGE_READONLY, (DWORD)(fSize >> 32), (DWORD)fSize, NULL);
	if (hFileMap == NULL) {
		CloseHandle(hFile);
		return;
	}

	SYSTEM_INFO systemInfo;
	GetSystemInfo(&systemInfo);
	DWORD dwPageSize = systemInfo.dwPageSize;
	DWORD dwAllocationGranularity = systemInfo.dwAllocationGranularity;

#ifdef _DEBUG
	DWORD tc = GetTickCount();
#endif
	g_szFrom = szFrom5;
	g_szFromLen = 5;
	TRACE("fSize = %lld\n", fSize);
	pCUPDUPData->SetProgress(_T("Parsing file..."), 0);
	// TODO: due to breaking the file into multiple chunks, it looks some emails can be lost
	bool firstView = true;;
	bool lastView = false;
	DWORD bufSize = 0;
	_int64 aligned_offset = 0;
	_int64 delta = 0;

	_int64 lastStartOffset = 0;
	while  ((lastView == false) && !pCUPDUPData->ShouldTerminate()) 
	{
		s_curmap = lastStartOffset;
		aligned_offset = (s_curmap / dwAllocationGranularity) * dwAllocationGranularity;
		delta = s_curmap - aligned_offset;
		bufSize = ((fSize - aligned_offset) < MAPPING_SIZE) ? (DWORD)(fSize - aligned_offset) : (DWORD)MAPPING_SIZE;

		TRACE("offset=%lld, bufsize=%ld, fSize-curmap=%lld, end=%lld\n", s_curmap, bufSize, fSize - s_curmap, s_curmap + bufSize);
		char * pview = (char *)MapViewOfFileEx(hFileMap, FILE_MAP_READ, (DWORD)(aligned_offset >> 32), (DWORD)aligned_offset, bufSize, NULL);

		if ((aligned_offset + bufSize) < fSize)
			lastView = false;
		else 
			lastView = true;

		lastStartOffset = aligned_offset + bufSize;

		char *p = pview + delta;
		int viewBufSize = bufSize - (DWORD)delta;
		_int64 viewOffset = s_curmap;
		MboxMail::Process(p, viewBufSize, viewOffset, firstView, lastView, lastStartOffset, bEml);
		firstView = false;

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


typedef struct {
	UINT m_charsetId;
	char *m_charset;
	char *m_info;
} CP2NM;

CP2NM cp2name[] = {
	//{ Info.CodePage , "Info.Name" , "Info.DisplayName" },
	{ 37 , "IBM037" , "IBM EBCDIC (US-Canada)" },
	{ 437 , "IBM437" , "OEM United States" },
	{ 500 , "IBM500" , "IBM EBCDIC (International)" },
	{ 708 , "ASMO-708" , "Arabic (ASMO 708)" },
	{ 720 , "DOS-720" , "Arabic (DOS)" },
	{ 737 , "ibm737" , "Greek (DOS)" },
	{ 775 , "ibm775" , "Baltic (DOS)" },
	{ 850 , "ibm850" , "Western European (DOS)" },
	{ 852 , "ibm852" , "Central European (DOS)" },
	{ 855 , "IBM855" , "OEM Cyrillic" },
	{ 857 , "ibm857" , "Turkish (DOS)" },
	{ 858 , "IBM00858" , "OEM Multilingual Latin I" },
	{ 860 , "IBM860" , "Portuguese (DOS)" },
	{ 861 , "ibm861" , "Icelandic (DOS)" },
	{ 862 , "DOS-862" , "Hebrew (DOS)" },
	{ 863 , "IBM863" , "French Canadian (DOS)" },
	{ 864 , "IBM864" , "Arabic (864)" },
	{ 865 , "IBM865" , "Nordic (DOS)" },
	{ 866 , "cp866" , "Cyrillic (DOS)" },
	{ 869 , "ibm869" , "Greek, Modern (DOS)" },
	{ 870 , "IBM870" , "IBM EBCDIC (Multilingual Latin-2)" },
	{ 874 , "windows-874" , "Thai (Windows)" },
	{ 875 , "cp875" , "IBM EBCDIC (Greek Modern)" },
	{ 932 , "shift_jis" , "Japanese (Shift-JIS)" },
	{ 936 , "gb2312" , "Chinese Simplified (GB2312)" },
	{ 949 , "ks_c_5601-1987" , "Korean" },
	{ 950 , "big5" , "Chinese Traditional (Big5)" },
	{ 1026 , "IBM1026" , "IBM EBCDIC (Turkish Latin-5)" },
	{ 1047 , "IBM01047" , "IBM Latin-1" },
	{ 1140 , "IBM01140" , "IBM EBCDIC (US-Canada-Euro)" },
	{ 1141 , "IBM01141" , "IBM EBCDIC (Germany-Euro)" },
	{ 1142 , "IBM01142" , "IBM EBCDIC (Denmark-Norway-Euro)" },
	{ 1143 , "IBM01143" , "IBM EBCDIC (Finland-Sweden-Euro)" },
	{ 1144 , "IBM01144" , "IBM EBCDIC (Italy-Euro)" },
	{ 1145 , "IBM01145" , "IBM EBCDIC (Spain-Euro)" },
	{ 1146 , "IBM01146" , "IBM EBCDIC (UK-Euro)" },
	{ 1147 , "IBM01147" , "IBM EBCDIC (France-Euro)" },
	{ 1148 , "IBM01148" , "IBM EBCDIC (International-Euro)" },
	{ 1149 , "IBM01149" , "IBM EBCDIC (Icelandic-Euro)" },
	{ 1200 , "utf-16" , "Unicode" },
	{ 1201 , "unicodeFFFE" , "Unicode (Big-Endian)" },
	{ 1250 , "windows-1250" , "Central European (Windows)" },
	{ 1251 , "windows-1251" , "Cyrillic (Windows)" },
	{ 1252 , "Windows-1252" , "Western European (Windows)" },
	{ 1253 , "windows-1253" , "Greek (Windows)" },
	{ 1254 , "windows-1254" , "Turkish (Windows)" },
	{ 1255 , "windows-1255" , "Hebrew (Windows)" },
	{ 1256 , "windows-1256" , "Arabic (Windows)" },
	{ 1257 , "windows-1257" , "Baltic (Windows)" },
	{ 1258 , "windows-1258" , "Vietnamese (Windows)" },
	{ 1361 , "Johab" , "Korean (Johab)" },
	{ 10000 , "macintosh" , "Western European (Mac)" },
	{ 10001 , "x-mac-japanese" , "Japanese (Mac)" },
	{ 10002 , "x-mac-chinesetrad" , "Chinese Traditional (Mac)" },
	{ 10003 , "x-mac-korean" , "Korean (Mac)" },
	{ 10004 , "x-mac-arabic" , "Arabic (Mac)" },
	{ 10005 , "x-mac-hebrew" , "Hebrew (Mac)" },
	{ 10006 , "x-mac-greek" , "Greek (Mac)" },
	{ 10007 , "x-mac-cyrillic" , "Cyrillic (Mac)" },
	{ 10008 , "x-mac-chinesesimp" , "Chinese Simplified (Mac)" },
	{ 10010 , "x-mac-romanian" , "Romanian (Mac)" },
	{ 10017 , "x-mac-ukrainian" , "Ukrainian (Mac)" },
	{ 10021 , "x-mac-thai" , "Thai (Mac)" },
	{ 10029 , "x-mac-ce" , "Central European (Mac)" },
	{ 10079 , "x-mac-icelandic" , "Icelandic (Mac)" },
	{ 10081 , "x-mac-turkish" , "Turkish (Mac)" },
	{ 10082 , "x-mac-croatian" , "Croatian (Mac)" },
	{ 12000 , "utf-32" , "Unicode (UTF-32)" },
	{ 12001 , "utf-32BE" , "Unicode (UTF-32 Big-Endian)" },
	{ 20000 , "x-Chinese-CNS" , "Chinese Traditional (CNS)" },
	{ 20001 , "x-cp20001" , "TCA Taiwan" },
	{ 20002 , "x-Chinese-Eten" , "Chinese Traditional (Eten)" },
	{ 20003 , "x-cp20003" , "IBM5550 Taiwan" },
	{ 20004 , "x-cp20004" , "TeleText Taiwan" },
	{ 20005 , "x-cp20005" , "Wang Taiwan" },
	{ 20105 , "x-IA5" , "Western European (IA5)" },
	{ 20106 , "x-IA5-German" , "German (IA5)" },
	{ 20107 , "x-IA5-Swedish" , "Swedish (IA5)" },
	{ 20108 , "x-IA5-Norwegian" , "Norwegian (IA5)" },
	{ 20127 , "us-ascii" , "US-ASCII" },
	{ 20261 , "x-cp20261" , "T.61" },
	{ 20269 , "x-cp20269" , "ISO-6937" },
	{ 20273 , "IBM273" , "IBM EBCDIC (Germany)" },
	{ 20277 , "IBM277" , "IBM EBCDIC (Denmark-Norway)" },
	{ 20278 , "IBM278" , "IBM EBCDIC (Finland-Sweden)" },
	{ 20280 , "IBM280" , "IBM EBCDIC (Italy)" },
	{ 20284 , "IBM284" , "IBM EBCDIC (Spain)" },
	{ 20285 , "IBM285" , "IBM EBCDIC (UK)" },
	{ 20290 , "IBM290" , "IBM EBCDIC (Japanese katakana)" },
	{ 20297 , "IBM297" , "IBM EBCDIC (France)" },
	{ 20420 , "IBM420" , "IBM EBCDIC (Arabic)" },
	{ 20423 , "IBM423" , "IBM EBCDIC (Greek)" },
	{ 20424 , "IBM424" , "IBM EBCDIC (Hebrew)" },
	{ 20833 , "x-EBCDIC-KoreanExtended" , "IBM EBCDIC (Korean Extended)" },
	{ 20838 , "IBM-Thai" , "IBM EBCDIC (Thai)" },
	{ 20866 , "koi8-r" , "Cyrillic (KOI8-R)" },
	{ 20871 , "IBM871" , "IBM EBCDIC (Icelandic)" },
	{ 20880 , "IBM880" , "IBM EBCDIC (Cyrillic Russian)" },
	{ 20905 , "IBM905" , "IBM EBCDIC (Turkish)" },
	{ 20924 , "IBM00924" , "IBM Latin-1" },
	{ 20932 , "EUC-JP" , "Japanese (JIS 0208-1990 and 0212-1990)" },
	{ 20936 , "x-cp20936" , "Chinese Simplified (GB2312-80)" },
	{ 20949 , "x-cp20949" , "Korean Wansung" },
	{ 21025 , "cp1025" , "IBM EBCDIC (Cyrillic Serbian-Bulgarian)" },
	{ 21866 , "koi8-u" , "Cyrillic (KOI8-U)" },
	{ 28591 , "iso-8859-1" , "Western European (ISO)" },
	{ 28592 , "iso-8859-2" , "Central European (ISO)" },
	{ 28593 , "iso-8859-3" , "Latin 3 (ISO)" },
	{ 28594 , "iso-8859-4" , "Baltic (ISO)" },
	{ 28595 , "iso-8859-5" , "Cyrillic (ISO)" },
	{ 28596 , "iso-8859-6" , "Arabic (ISO)" },
	{ 28597 , "iso-8859-7" , "Greek (ISO)" },
	{ 28598 , "iso-8859-8" , "Hebrew (ISO-Visual)" },
	{ 28599 , "iso-8859-9" , "Turkish (ISO)" },
	{ 28603 , "iso-8859-13" , "Estonian (ISO)" },
	{ 28605 , "iso-8859-15" , "Latin 9 (ISO)" },
	{ 29001 , "x-Europa" , "Europa" },
	{ 38598 , "iso-8859-8-i" , "Hebrew (ISO-Logical)" },
	{ 50220 , "iso-2022-jp" , "Japanese (JIS)" },
	{ 50221 , "csISO2022JP" , "Japanese (JIS-Allow 1 byte Kana)" },
	{ 50222 , "iso-2022-jp" , "Japanese (JIS-Allow 1 byte Kana - SO/SI)" },
	{ 50225 , "iso-2022-kr" , "Korean (ISO)" },
	{ 50227 , "x-cp50227" , "Chinese Simplified (ISO-2022)" },
	{ 51932 , "euc-jp" , "Japanese (EUC)" },
	{ 51936 , "EUC-CN" , "Chinese Simplified (EUC)" },
	{ 51949 , "euc-kr" , "Korean (EUC)" },
	{ 52936 , "hz-gb-2312" , "Chinese Simplified (HZ)" },
	{ 54936 , "GB18030" , "Chinese Simplified (GB18030)" },
	{ 57002 , "x-iscii-de" , "ISCII Devanagari" },
	{ 57003 , "x-iscii-be" , "ISCII Bengali" },
	{ 57004 , "x-iscii-ta" , "ISCII Tamil" },
	{ 57005 , "x-iscii-te" , "ISCII Telugu" },
	{ 57006 , "x-iscii-as" , "ISCII Assamese" },
	{ 57007 , "x-iscii-or" , "ISCII Oriya" },
	{ 57008 , "x-iscii-ka" , "ISCII Kannada" },
	{ 57009 , "x-iscii-ma" , "ISCII Malayalam" },
	{ 57010 , "x-iscii-gu" , "ISCII Gujarati" },
	{ 57011 , "x-iscii-pa" , "ISCII Punjabi" },
	{ 65000 , "utf-7" , "Unicode (UTF-7)" },
	{ 65001 , "utf-8" , "Unicode (UTF-8)" },
};

#include <unordered_map>

typedef unordered_map<std::string, unsigned int> myMap;
static myMap *cids = 0;

void delete_charset2Id() {
	delete cids;
}

UINT charset2Id(const char *char_set)
{
	UINT id = 0;
	CP2NM *item;
	myMap::iterator it;

	if (!cids) {
		cids = new myMap;

		int cp2name_size = sizeof(cp2name) / sizeof(CP2NM);
		for (int i = 0; i < cp2name_size; i++)
		{
			item = &cp2name[i];
			std::string charset = item->m_charset;
			transform(charset.begin(), charset.end(), charset.begin(), ::tolower);
			if (cids->find(charset) == cids->end()) {  // not found, invalid iterator returned
				cids->insert(myMap::value_type(charset, item->m_charsetId));
			}
		}
	}
#if 0
	for (it = cids->begin(); it != cids->end(); it++) {
		TRACE("%d %s\n", it->second, it->first);
	}
#endif
	std::string charset = char_set;
	transform(charset.begin(), charset.end(), charset.begin(), ::tolower);
	if ((it = cids->find(charset)) != cids->end()) {
		id = it->second;
	}
	return id;
}

typedef unordered_map<unsigned int, std::string> myIdMap;
static myIdMap *ids = 0;

void delete_id2charset() {
	delete ids;
}

BOOL id2charset(UINT id, std::string &charset)
{
	CP2NM *item;
	myIdMap::iterator it;

	if (!ids) {
		ids = new myIdMap;

		int cp2name_size = sizeof(cp2name) / sizeof(CP2NM);
		for (int i = 0; i < cp2name_size; i++)
		{
			item = &cp2name[i];
			if (ids->find(item->m_charsetId) == ids->end()) {  // not found, invalid iterator returned
				ids->insert(myIdMap::value_type(item->m_charsetId, item->m_charset));
			}
		}
	}
#if 0
	for (it = ids->begin(); it != ids->end(); it++) {
		TRACE("%d %s\n", it->first, it->second);
	}
#endif
	if ((it = ids->find(id)) != ids->end()) {
		charset = it->second;
		return TRUE;
	}
	return FALSE;
}

int Write2File(HANDLE hFile, LPCVOID lpBuffer, DWORD nNumberOfBytesToWrite, LPDWORD lpNumberOfBytesWritten)
{
	DWORD bytesToWrite = nNumberOfBytesToWrite;
	DWORD nwritten = 0;
	while (bytesToWrite > 0)
	{
		nwritten = 0;
		if (!WriteFile(hFile, lpBuffer, bytesToWrite, &nwritten, NULL)) {
			DWORD retval = GetLastError();
			break;
		}
		bytesToWrite -= nwritten;
	}
	*lpNumberOfBytesWritten = nNumberOfBytesToWrite - bytesToWrite;
	if (*lpNumberOfBytesWritten != nNumberOfBytesToWrite)
		return FALSE;
	else
		return TRUE;
}

int MboxMail::DumpMailBox(MboxMail *mailBox, int which)
{
	//return 1;
	static BOOL restarted = TRUE;
	int deb; // set breakpoint
	char *CRLF = "\n\n";
	char buff[32768+2];
	int buffsize = 32768;
	//char *buff = &buffer[0];

	char datebuff[256];
	DWORD offset = 0;
	DWORD nwritten = 0;
	DWORD count = 0;
	BOOL retval;

	CString cacheTextFile = "DumpMailBox.txt";
	HANDLE hFile = CreateFile(cacheTextFile, GENERIC_WRITE, 0, NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE) {
		DWORD err = GetLastError();
		TRACE(_T("(dumpSelectedItem)INVALID_HANDLE_VALUE error= %ld\n"), err);
		return -1;
	}

	HANDLE mbox_hFile = CreateFile(MboxMail::s_path, GENERIC_READ, FILE_SHARE_READ, NULL,
		OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL);
	if (mbox_hFile == INVALID_HANDLE_VALUE) {
		DWORD err = GetLastError();
		TRACE(_T("(dumpSelectedItem)INVALID_HANDLE_VALUE error= %ld\n"), err);
		CloseHandle(hFile);
		return -1;
	}

	_int64 pos = 0;
	if (restarted == TRUE) {
		SetEndOfFile(hFile);
		restarted = FALSE;
	}
	else
	{
		LARGE_INTEGER li;
		
		retval = GetFileSizeEx(hFile, &li);
		long long fsize = li.QuadPart;
		pos = FileSeek(hFile, fsize, FILE_BEGIN);
		if (pos < 0) {
			DWORD err = GetLastError();
			return -1;
		}
	}

	MboxMail *m = mailBox;

	CString format;
	datebuff[0] = 0;
	if (m->m_timeDate > 0) {
		CTime tt(m->m_timeDate);
		strcpy(datebuff, (LPCSTR)tt.Format(format));
	}

	count = sprintf_s(buff, "INDX=%d first=%lld len=%d last=%lld att=%d hlen=%d rcv=%d date=\"%s\" from=\"%s\" to=\"%s\" subj=\"%s\"\n",
		which, m->m_startOff, m->m_length, (m->m_startOff + m->m_length - 1), m->m_hasAttachments, m->m_headLength, m->m_recv,
		datebuff, (LPCSTR)m->m_from, (LPCSTR)m->m_to, (LPCSTR)m->m_subj);

	nwritten = 0;
	if (!Write2File(hFile, buff, count, &nwritten)) {
		DWORD retval = GetLastError();
	}

	long long start_offset = m->m_startOff;
	pos = FileSeek(mbox_hFile, start_offset, FILE_BEGIN);

	DWORD bytestoRead = m->m_length;
	DWORD nNumberOfBytesToRead = buffsize;
	DWORD lpNumberOfBytesRead = 0;
	retval = TRUE;

#if 0
	while (bytestoRead > 0)
	{
		if (bytestoRead > buffsize)
			nNumberOfBytesToRead = buffsize;
		else
			nNumberOfBytesToRead = bytestoRead;

		retval = ReadFile(mbox_hFile, buff, nNumberOfBytesToRead, &lpNumberOfBytesRead, 0);
		if (retval != TRUE) {
			DWORD retval = GetLastError();
			break;
		}
		bytestoRead -= lpNumberOfBytesRead;

		nwritten = 0;
		if (!Write2File(hFile, buff, lpNumberOfBytesRead, &nwritten)) {
			DWORD retval = GetLastError();
			break;
		}
	}
#endif

	std::vector <MailBodyContent*>::iterator it;

	if (m_ContentDetailsArray.size() > 1)
		int deb = 1;

	MailBodyContent *body;
	for (it = m_ContentDetailsArray.begin(); it != m_ContentDetailsArray.end(); it++)
	{
		body = *it;
		CString txt;
		txt.Format(_T("\n\ntype=%s encoding=%s distribution=%s attachment=%s offset=%lld length=%ld\n"),
			body->m_contentType, body->m_contentTransferEncoding, body->m_contentDisposition, body->m_attachmentName, body->m_contentOffset, body->m_contentLength);

		nwritten = 0;
		if (!Write2File(hFile, txt.GetBuffer(), txt.GetLength(), &nwritten)) {
			DWORD retval = GetLastError();
		}

		start_offset = m->m_startOff + body->m_contentOffset;
		pos = FileSeek(mbox_hFile, start_offset, FILE_BEGIN);

		bytestoRead = body->m_contentLength;
		nNumberOfBytesToRead = bytestoRead;
		lpNumberOfBytesRead = 0;
		retval = TRUE;

		if (body->m_contentTransferEncoding.CompareNoCase("base64") == 0) 
		{
			char *bigbuff = new char[body->m_contentLength+2];
			while (bytestoRead > 0)
			{
				retval = ReadFile(mbox_hFile, &bigbuff[lpNumberOfBytesRead], bytestoRead, &lpNumberOfBytesRead, 0);
				if (retval != TRUE) {
					DWORD retval = GetLastError();
					break;
				}
				bytestoRead -= lpNumberOfBytesRead;
			}

			MboxCMimeCodeBase64 *d64 = new MboxCMimeCodeBase64(bigbuff, body->m_contentLength);
			int dlength = d64->GetOutputLength();
			if (dlength > MboxMail::m_maxOutput) {
				if (MboxMail::m_pbOutput) delete[] MboxMail::m_pbOutput;
				MboxMail::m_pbOutput = new unsigned char[dlength];
				MboxMail::m_maxOutput = dlength;
			}
			int retlen = d64->GetOutput(MboxMail::m_pbOutput, dlength);
#if 0
			nwritten = 0;
			if (!Write2File(hFile, bigbuff, body->m_contentLength, &nwritten)) {
				DWORD retval = GetLastError();
			}

			if (!Write2File(hFile, CRLF, strlen(CRLF), &nwritten)) {
				DWORD retval = GetLastError();
			}
#endif

			nwritten = 0;
			if (!Write2File(hFile, MboxMail::m_pbOutput, retlen, &nwritten)) {
				DWORD retval = GetLastError();
			}

			if (!Write2File(hFile, CRLF, strlen(CRLF), &nwritten)) {
				DWORD retval = GetLastError();
			}
			deb = 1;
		}
		else if (body->m_contentTransferEncoding.CompareNoCase("quoted-printable") == 0)
		{
			char *bigbuff = new char[body->m_contentLength + 2];
			while (bytestoRead > 0)
			{
				retval = ReadFile(mbox_hFile, &bigbuff[lpNumberOfBytesRead], bytestoRead, &lpNumberOfBytesRead, 0);
				if (retval != TRUE) {
					DWORD retval = GetLastError();
					break;
				}
				bytestoRead -= lpNumberOfBytesRead;
			}

			MboxCMimeCodeQP *dGP = new MboxCMimeCodeQP(bigbuff, body->m_contentLength);
			int dlength = dGP->GetOutputLength();
			if (dlength > MboxMail::m_maxOutput) {
				if (MboxMail::m_pbOutput) delete[] MboxMail::m_pbOutput;
				MboxMail::m_pbOutput = new unsigned char[dlength];
				MboxMail::m_maxOutput = dlength;
			}
			int retlen = dGP->GetOutput(MboxMail::m_pbOutput, dlength);
#if 0
			nwritten = 0;
			if (!Write2File(hFile, bigbuff, body->m_contentLength, &nwritten)) {
				DWORD retval = GetLastError();
			}

			if (!Write2File(hFile, CRLF, strlen(CRLF), &nwritten)) {
				DWORD retval = GetLastError();
			}
#endif
			nwritten = 0;
			if (!Write2File(hFile, MboxMail::m_pbOutput, retlen, &nwritten)) {
				DWORD retval = GetLastError();
			}

			if (!Write2File(hFile, CRLF, strlen(CRLF), &nwritten)) {
				DWORD retval = GetLastError();
			}
			deb = 1;
		}
		else
		{
			bytestoRead = body->m_contentLength;
			while (bytestoRead > 0)
			{
				if (bytestoRead > buffsize)
					nNumberOfBytesToRead = buffsize;
				else
					nNumberOfBytesToRead = bytestoRead;

				retval = ReadFile(mbox_hFile, buff, nNumberOfBytesToRead, &lpNumberOfBytesRead, 0);
				if (retval != TRUE) {
					DWORD retval = GetLastError();
					break;
				}
				bytestoRead -= lpNumberOfBytesRead;

				nwritten = 0;
				if (!Write2File(hFile, buff, lpNumberOfBytesRead, &nwritten)) {
					DWORD retval = GetLastError();
					break;
				}
			}
		}
	}
	nwritten = 0;
	if (!Write2File(hFile, CRLF, strlen(CRLF), &nwritten)) {
		DWORD retval = GetLastError();
	}

	CloseHandle(hFile);
	CloseHandle(mbox_hFile);

	return 1;
}

char * GetMultiLine(char *p, char *e, CString &line)
{
	char *p_beg = p;
	while (p < e && *p++ != '\n');

	char *ss = line.GetBufferSetLength(p - p_beg);
	::memcpy(ss, p_beg, p - p_beg);
	line.Trim();

	if (line.IsEmpty())
		return p;

	while ((p < e) && ((*p == ' ') || (*p == '\t')))
	{
		char *p_next_beg = p;
		EATNL();

		CString nextLine;
		char *ss = nextLine.GetBufferSetLength(p - p_next_beg);
		::memcpy(ss, p_next_beg, p - p_next_beg);
		nextLine.Trim();
		line += nextLine;
	}
	return p;
};

int GetFieldValue(CString &fieldLine, int startPos, CString &value)
{
	int posEnd = fieldLine.FindOneOf(";\n\r");
	if (posEnd < 0)
		value = fieldLine.Mid(startPos);
	else
		value = fieldLine.Mid(startPos, posEnd - startPos);
	value.Trim();
	return 1;
}

int strncmpUpper2Lower(char *any, char *end, const char *lower, int lowerlength) {
	// any can be multi line
	// lower is fixed length
	while ((lowerlength > 0) && (any < end) && (*lower++ == tolower(*any++))) { lowerlength--; }
	return lowerlength;
}

int strncmpUpper2Lower(char *any, int anyLength, const char *lower, int lowerlength) {
	// any can be multi line
	// lower is fixed length
	if (anyLength >= lowerlength) {
		while ((lowerlength > 0) && (*lower++ == tolower(*any++))) { lowerlength--; }
	}
	return lowerlength;
}

int strncmpExact(char *any, char *end, const char *lower, int lowerlength) {
	while ((lowerlength > 0) && (any < end) && (*lower++ == *any++)) { lowerlength--; }
	return lowerlength;
}

/* ParseContent Parser is not a full blown parser but rather simplified to support searching of content blocks.
   MIME is defined as a nested set of data containers similar to the tree structure. 
   Boundary tags are used to indicate begining and end of the data block.
   Example:
   header
   multipart/mixed ; boundary=mixTag
   mixTag
     multipart/alernative; boundary=altTag
	 --altTag
	 text/plain
	 --altTag
	 text/html
	 --altTag--
   -mixTag
   attachment
   --mixTag--
*/


char * ParseContent(MboxMail *mail, char *startPos, char *endPos)
{
	static const char *cType = "content-type:";
	static const int cTypeLen = strlen(cType);
	static const char *cTransferEncoding = "content-transfer-encoding:";
	static const int cTransferEncodingLen = strlen(cTransferEncoding);
	static const char *cDisposition = "content-disposition:";
	static const int cDispositionLen = strlen(cDisposition);
	static const char *cAttachment = "attachment";
	static const int cAttachmentLen = strlen(cAttachment);
	static const char *cName = "name=";
	static const int cNameLen = strlen(cName);
	static const char *cFileName = "filename=";
	static const int cFileNameLen = strlen(cFileName);

	char *p = startPos;
	char *e = endPos;
	char *tag = "";
	char *potential_tag = "";
	int contentIndex = 0;
	int contentLength = 0;
	int deb; // set breakpoint

	BOOL contentTypeFound = FALSE;
	while ((p < e) && ((*p == '\r') || (*p == '\n')))  // eat empty lines
		p++;

	// make sure p is incremeneted within a loop or break
	while (p < e)
	{
		if (((*p == '\r') && (*(p + 1) == '\n')) || (*p == '\n')) {
			mail->m_headLength = p - startPos;
			EATNL();
			break;  // end of header
		}
		else if (strncmpUpper2Lower(p, e, cType, cTypeLen) == 0) {
			tag = potential_tag;
			contentTypeFound = TRUE;
			break;
		}
		else
			potential_tag = p;
		EATNL();
	}

	char *contentBegin = p;
	char *contentEnd = p;
	if (contentTypeFound == FALSE)
	{
		// We are at empty line; Find Content-Type line if any or the end of implicit content block
		while (p < e)
		{
			if (strncmpUpper2Lower(p, e, cType, cTypeLen) == 0)
			{
				tag = potential_tag;
				contentTypeFound = TRUE;
				break;
			}
			potential_tag = p;
			EATNL();

		}
		contentEnd = p;
		contentLength = contentEnd - contentBegin;
		if (contentLength > 0) {
			MailBodyContent *contentDetails = new MailBodyContent;
			contentDetails->m_contentOffset = contentBegin - startPos;;
			contentDetails->m_contentLength = contentLength;

			mail->m_ContentDetailsArray.push_back(contentDetails);
		}
		else
			deb = 1;
		if (contentTypeFound == FALSE)
			return p;
	}

	CString line;
	CString transferEncoding;
	CString disposition;
	CString contentType;
	CString fileName;
	CString plainTextTag;

	// make sure p is incremeneted within a loop or break
	while (p < e)
	{
		if (strncmpUpper2Lower(p, e, cType, cTypeLen) == 0)
		{
			transferEncoding.SetString("");
			disposition.SetString("");
			contentType.SetString("");
			fileName.SetString("");

			tag = potential_tag;
			p = GetMultiLine(p, e, line);
			GetFieldValue(line, cTypeLen, contentType);
			int pc = line.Find(cName);
			if (pc != -1) {
				CString name = line.Mid(pc);
				int nameLength = name.FindOneOf(";\n\r");
				if (nameLength < 0)
					fileName = name.Mid(cNameLen);
				else
					fileName = name.Mid(cNameLen, nameLength);
				fileName.Trim("\"\\");
			}

			// make sure p is incremeneted within a loop or break
			while ((p < e) && !(((*p == '\r') && (*(p + 1) == '\n')) || (*p == '\n')))
			{
				potential_tag = p;  // may need to update if we decide to handle unexpected cType
				if (strncmpUpper2Lower(p, e, cTransferEncoding, cTransferEncodingLen) == 0)
				{
					p = GetMultiLine(p, e, line);
					GetFieldValue(line, cTransferEncodingLen, transferEncoding);
				}
				else if (strncmpUpper2Lower(p, e, cDisposition, cDispositionLen) == 0)
				{
					p = GetMultiLine(p, e, line);
					GetFieldValue(line, cDispositionLen, disposition);
					int pc = line.Find(cFileName);
					if (pc != -1) {
						CString name = line.Mid(pc);
						int nameLength = name.FindOneOf(";\n\r");
						if (nameLength < 0)
							fileName = name.Mid(cFileNameLen);
						else
							fileName = name.Mid(cFileNameLen, nameLength);
						fileName.Trim("\"\\");
					}
				}
				else if (strncmpUpper2Lower(p, e, cType, cTypeLen) == 0)
				{
					// TODO: file might not be well formatted or be corrupted. 
					// Ignore for now to keep it simple ?
					//p = GetMultiLine(p, e, line);
					//GetFieldValue(line, cTypeLen, contentType);
					contentTypeFound = TRUE;
					// break; 
					EATNL();  // delete if p = GetMultiLine(p, e, line); is enabled
				}
				else {
					EATNL();
				}
			}

			contentBegin = p;
			contentEnd = p;
			potential_tag = p;

			if (disposition.CompareNoCase(cAttachment) == 0) {
					mail->m_hasAttachments = 1;
			}

			contentType.MakeLower();
			if (contentType.Compare("text") == 0)
				contentType += "/plain";
			if (contentType.Find("text/") >= 0)
			{
				if (strncmp(tag, "--", 2) == 0)
				{
					char *tagEnd = strpbrk(tag, ";\r\n");
					int tagLen = tagEnd - tag;

					if (contentType.Compare("text/plain") == 0) {
						LPTSTR buf = plainTextTag.GetBufferSetLength(tagLen);
						memcpy(buf, tag, tagLen);
					}

					// make sure p is incremeneted within a loop or break
					while (p < e)
					{
						potential_tag = p;  // may need to update if we decide to handle unexpected cType
						if (strncmpExact(p, e, tag, tagLen) == 0) {
							break;
						}
						else if (strncmpUpper2Lower(p, e, cType, cTypeLen) == 0)
						{
							// TODO: file might not be well formatted or be corrupted. 
							// Ignore for now to keep it simple ?
							//p = GetMultiLine(p, e, line);
							//GetFieldValue(line, cDispositionLen, disposition);
							contentTypeFound = TRUE;
							// break; 
							EATNL();  // delete if p = GetMultiLine(p, e, line); is enabled
						}
						else {
							EATNL();
							contentEnd = p;
						}
					}
					if (transferEncoding.CompareNoCase("base64") == 0)
					{
#if 0
						MboxCMimeCodeBase64 *d64 = new MboxCMimeCodeBase64(contentBegin, contentLength);
						int dlength = d64->GetOutputLength();
						if (dlength > MboxMail::m_maxOutput) {
							if (MboxMail::m_pbOutput) delete[] MboxMail::m_pbOutput;
							MboxMail::m_pbOutput = new unsigned char[dlength];
							MboxMail::m_maxOutput = dlength;
						}
						int ret = d64->GetOutput(MboxMail::m_pbOutput, dlength);
#endif
						deb = 1;
					}
				}
				else  // no tag; find next Content-Type of end of file view
				{
					// make sure p is incremeneted within a loop or break
					while (p < e)
					{
						if (strncmpUpper2Lower(p, e, cType, cTypeLen) == 0)
						{
							tag = potential_tag;
							break;
						}

						potential_tag = p;
						EATNL();
						contentEnd = p;
					}
					deb = 1;
				}

				BOOL ignoreDuplicateHtml = FALSE;
				if (contentType.Compare("text/html") == 0) {
					if (strncmp(plainTextTag, tag, plainTextTag.GetLength()) == 0) {
						ignoreDuplicateHtml = TRUE;
						plainTextTag.SetString("");
					}
				}

				contentLength = contentEnd - contentBegin;
				if ((ignoreDuplicateHtml == FALSE) && (contentLength > 0)) {
					MailBodyContent *contentDetails = new MailBodyContent;
					contentDetails->m_contentOffset = contentBegin - startPos;;
					contentDetails->m_contentLength = contentLength;
					contentDetails->m_contentType = contentType;
					contentDetails->m_contentDisposition = disposition;
					contentDetails->m_contentTransferEncoding = transferEncoding;
					contentDetails->m_attachmentName = fileName;

					mail->m_ContentDetailsArray.push_back(contentDetails);
				}
				else {
					deb = 1;
				}
			}
			potential_tag = p;
		}
		else {
			potential_tag = p;
			EATNL();
		}
	}

	if (e != endPos)
		deb = 1;
	if (e > endPos)
		deb = 1;
	return e;
}











