
#include "stdafx.h"
#include <afxtempl.h>
#include "mboxview.h"
#include "MboxMail.h"

MailBodyPool *MailBody::m_mpool = new MailBodyPool;

UINT charset2Id(const char *char_set);
BOOL id2charset(UINT id, std::string &charset);

MessageIdTableType *MboxMail::m_pMessageIdTable = 0;
int MboxMail::m_nextGroupId = 0;

SimpleString* MboxMail::m_outbuf = new SimpleString(10000);
SimpleString* MboxMail::m_inbuf = new SimpleString(10000);
SimpleString* MboxMail::m_workbuf = new SimpleString(10000);

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

inline char *EatNewLine(char* p, char*e) {
	while ((p < e) && (*p++ != '\n'));
	return p;
}

struct MsgIdHash {
public:
	size_t operator()(const MboxMail *key) const
	{
		size_t hashsum = StrHash((const char*)(LPCSTR)key->m_messageId, key->m_messageId.GetLength());
		return hashsum;
	}
};

struct MsgIdEqual {
public:
	bool operator()(const MboxMail *key1, const MboxMail *key2) const
	{
		if (key1->m_messageId == key2->m_messageId)
			return true;
		else
			return false;
	}
};

struct MessageIdHash {
public:
	size_t operator()(const CString *key) const
	{
		size_t hashsum = StrHash((const char*)(LPCSTR)*key, key->GetLength());
		return hashsum;
	}
};

struct MessageIdEqual {
public:
	bool operator()(const CString *key1, const CString *key2) const
	{
		if (*key1 == *key2)
			return true;
		else
			return false;
	}
};

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

BOOL Str2CodePage(SimpleString *str, UINT inCodePage, UINT outCodePage, SimpleString *result, SimpleString *workBuff)
{
	int buffLen = str->Count() * 4 + 2;
	workBuff->ClearAndResize(buffLen);
	LPWSTR buff = (LPWSTR)workBuff->Data();  // or  we could call MultiByteToWideChar first to get the required length
	int wlen = MultiByteToWideChar(inCodePage, 0, str->Data(), str->Count(), buff, buffLen);
	if (wlen == 0) {
		result->Clear();
		// error - implement error log file
		const DWORD error = ::GetLastError();
		return FALSE;
	}

	int outLen = wlen *4 + 2;
	result->ClearAndResize(outLen); // or could  call WideCharToMultiByte first to get the required length
	int utf8Len = WideCharToMultiByte(CP_UTF8, 0, buff, wlen, result->Data(), outLen, NULL, NULL);
	if (utf8Len == 0) {
		result->Clear();
		// error - implement error log file
		const DWORD error = ::GetLastError();
		/*ERROR_INSUFFICIENT_BUFFER.A supplied buffer size was not large enough, or it was incorrectly set to NULL.
		ERROR_INVALID_FLAGS.The values supplied for flags were not valid.
		ERROR_INVALID_PARAMETER.Any of the parameter values was invalid.
		ERROR_NO_UNICODE_TRANSLATION.Invalid Unicode was found in a string.*/
		return FALSE;
	}
	result->SetCount(utf8Len);
	return TRUE;
}

BOOL Str2UTF8(SimpleString *str, UINT CodePage, SimpleString *result, SimpleString *workBuff)
{
	BOOL ret = Str2CodePage(str, CodePage, CP_UTF8, result, workBuff);
	return ret;
}

BOOL Str2CurrentCodepage(SimpleString *str, UINT CodePage, SimpleString *result, SimpleString *workBuff)
{
	UINT currentCodePage = GetACP();
	BOOL ret = Str2CodePage(str, CodePage, currentCodePage, result, workBuff);
	return ret;
}

#include "MimeCode.h"

CString MboxMail::DecodeString(CString &subj, CString &charset, UINT &charsetId, UINT toCharacterId)
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

		if ((CodePage > 0) && (toCharacterId > 0)) {
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

int strncmpUpper2Lower(char *any, char *end, const char *lower, int lowerlength);
char * GetMultiLine(char *p, char *e, CString &line);
int GetFieldValue(CString &fieldLine, int startPos, CString &value);

bool IsFromValidDelimiter(char *p, char *e)
{
	// make sure it is not user typed "^From" in the message text
	// We expect certain number of digit and colon charcaters if valid "From " mail delimiter
	// There is no inexpensive and sinple method to determine that "^From  points to the begining of the mail
	// Would likely have to attempt to see if the proper mail header follow
	// From 1572544789079124110@xxx Mon Jul 10 14:06:16 +0000 2017
	int digitCount = 0;
	int colonCount = 0;
	char c;
	while ( (p < e) && ((c = (*p++)) != '\n'))
	{
		if (c == ':')
			colonCount++;
		else if (isdigit(c))
			digitCount++;
	}

	if ((colonCount >= 2) && (digitCount >= 6))
		return true;
	else
		return false;
}

char szFrom5[] = "From ";
char szFrom6[] = "\nFrom ";
char	*g_szFrom;
int		g_szFromLen;

bool MboxMail::Process(register char *p, DWORD bufSize, _int64 startOffset,  bool bFirstView, bool bLastView, _int64 &lastStartOffset, bool bEml)
{
	static const char *cFrom = "from:";
	static const int cFromLen = strlen(cFrom);
	static const char *cTo = "to:";
	static const int cToLen = strlen(cTo);
	static const char *cCC = "cc:";
	static const int cCCLen = strlen(cCC);
	static const char *cBCC = "bcc:";
	static const int cBCCLen = strlen(cBCC);
	static const char *cSubject = "subject:";
	static const int cSubjectLen = strlen(cSubject);
	static const char *cDate = "date:";
	static const int cDateLen = strlen(cDate);
	static const char *cReceived = "received:";
	static const int cReceivedLen = strlen(cReceived);

	char *orig = p;
	register char *e = p + bufSize - 1;
	if (p == NULL || bufSize < 10) {
		return false;
	}
	CString contentDisposition = "Content-Disposition: attachment";
	int iFormat = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "format");
	CString format = GetDateFormat(iFormat);
	CString to, from, subject, date;
	CString cc, bcc;
	time_t tdate = 0;
	bool	bTo = true, bFrom = true, bSubject = true, bDate = true, bRcvDate = true; // indicates not found, false means found 
	bool bCC = true, bBCC = true;
	char *msgStart = NULL;
	int recv = TRUE;
	int curstep = (int)(startOffset / (s_step?s_step:1));
	CString line;
	CString rcved;

	if (bEml && bFirstView) {
		msgStart = p;
		p += 5;
		bEml = false;  // not used below and not needed :)
	}
	while (p < e - 4)   // TODO: why (e - 4) ??
	{
		try {
			if ((*(DWORD*)p == 0x6d6f7246 && p[4] == ' ') && IsFromValidDelimiter(p,e)) // "From "  marks beginning of the next mail
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
							m->m_cc = DecodeString(cc, m->m_cc_charset, m->m_cc_charsetId);
							m->m_bcc = DecodeString(bcc, m->m_bcc_charset, m->m_bcc_charsetId);
							if (!bcc.IsEmpty())
								int deb = 1;
							if (!cc.IsEmpty())
								int deb = 1;
							m->m_subj.Replace("\r", "");
							m->m_subj.Replace("\n", "");
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

							if (m->m_messageId == m->m_replyId)
								int deb = 1;


							int index = s_mails.GetCount() - 1;
							//m->DumpMailBox(m, index);
							//NListView::DumpItemDetails(index);

							if (s_mails.GetCount() == 42)
								int deb = 1;

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
					cc.Empty();
					bcc.Empty();
					tdate = 0;
					bTo = bFrom = bSubject = bDate = bRcvDate = true;
				}
				p += 5;
				p = EatNewLine(p, e);
			}
			else
			{
				if (bFrom && strncmpUpper2Lower(p, e, cFrom, cFromLen) == 0)
				{
					bFrom = false;
					p = GetMultiLine(p, e, line);
					from = line.Mid(cFromLen);
					from.Trim();
				}
				else if (bTo && strncmpUpper2Lower(p, e, cTo, cToLen) == 0)
				{
					bTo = false;
					p = GetMultiLine(p, e, line);
					to = line.Mid(cToLen);
					to.Trim();
				}
				else if (bCC && strncmpUpper2Lower(p, e, cCC, cCCLen) == 0)
				{
					bCC = false;
					p = GetMultiLine(p, e, line);
					cc = line.Mid(cCCLen);
					cc.Trim();
				}
				else if (bBCC && strncmpUpper2Lower(p, e, cBCC, cBCCLen) == 0)
				{
					bBCC = false;
					p = GetMultiLine(p, e, line);
					bcc = line.Mid(cBCCLen);
					bcc.Trim();
				}
				else if (bSubject && strncmpUpper2Lower(p, e, cSubject, cSubjectLen) == 0)
				{
					bSubject = false;
					p = GetMultiLine(p, e, line);
					subject = line.Mid(cSubjectLen);
					subject.Trim();
				}
				else if (bDate && bRcvDate && strncmpUpper2Lower(p, e, cReceived, cReceivedLen) == 0)
				{
					//bDate = false; 
					p = GetMultiLine(p, e, line);
					rcved = line.Mid(cReceivedLen);
					rcved.Trim();
					// find last ';' separator
					int pos = rcved.ReverseFind(';');
					date = rcved.Mid(pos + 1);
					SYSTEMTIME tm;
					if (DateParser::parseRFC822Date(date, &tm)) {
						if (DateParser::validateSystemtime(&tm)) {
							CTime tt(tm);
							date = tt.Format(format);
							tdate = tt.GetTime();
							bRcvDate = false;
							recv = TRUE;
						}
					}
				}
				else if (bDate && strncmpUpper2Lower(p, e, cDate, cDateLen) == 0)
				{
					p = GetMultiLine(p, e, line);
					date = line.Mid(cDateLen);
					date.Trim();
					SYSTEMTIME tm;
					if (DateParser::parseRFC822Date(date, &tm)) {
						if (DateParser::validateSystemtime(&tm)) {
							CTime tt(tm);
							date = tt.Format(format);
							tdate = tt.GetTime();
							bDate = false;
							recv = FALSE;
						}
					}
				}
				else
				{
					p = EatNewLine(p, e);
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
				m->m_cc = DecodeString(cc, m->m_cc_charset, m->m_cc_charsetId);
				m->m_bcc = DecodeString(bcc, m->m_bcc_charset, m->m_bcc_charsetId);
				if (!bcc.IsEmpty())
					int deb = 1;
				if (!cc.IsEmpty())
					int deb = 1;
				m->m_subj.Replace("\r", "");
				m->m_subj.Replace("\n", "");
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
	MboxMail *mb;
	MboxMail *mb_next;
	MboxMail *mb_ref;

	for (int m = 0; m < (MboxMail::s_mails.GetSize() - 1); m++) {
		mb = MboxMail::s_mails[m];
		mb_next = MboxMail::s_mails[m + 1];
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
		mb_ref = MboxMail::s_mails_ref[m];
		if (mb != mb_ref)
			outofOrderCnt++;
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

	TCHAR ext[_MAX_EXT + 1]; ext[0] = 0;
	TCHAR fname[_MAX_FNAME + 1]; fname[0] = 0;
	_tsplitpath(path, 0, 0, fname, ext);
	CString mailFile = CString(fname) + CString(ext);
	CString parsingFileText = _T("Parsing \"") + mailFile + _T("\" ...");

	//pCUPDUPData->SetProgress(parsingFileText, 0);  // works but doesn't always fit into progress bar
	pCUPDUPData->SetProgress(_T("Parsing file ..."), 0);
	// TODO: due to breaking the file into multiple chunks, it looks some emails can be lost : Fixed
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
	sort(MboxMail::s_mails.GetData(), MboxMail::s_mails.GetData() + MboxMail::s_mails.GetSize(), 
		[](MboxMail* o1, MboxMail* o2) {
			return cmpMbox(o1, o2);
		}
	);

	//cmpAllMbox(false);

#if 0
	// Just test to see if final sort wil restore order according to mail file
	MboxMail::s_mails_ref.SetSize(MboxMail::s_mails.GetSize());
	MboxMail::s_mails_ref.Copy(MboxMail::s_mails);

	MboxMail::SortByDate(0, 1);
	MboxMail::SortByFrom(0, 2);
	MboxMail::SortByTo(0, 3);
	MboxMail::SortBySubject(0, 4);
	MboxMail::SortBySize(0, 5);

	// sort the mails by offset
	sort(MboxMail::s_mails.GetData(), MboxMail::s_mails.GetData() + MboxMail::s_mails.GetSize(), 
		[](MboxMail* o1, MboxMail* o2) {
			return cmpMbox(o1, o2);
		}
	);

	cmpAllMbox(false);

#endif

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

bool sortByFileOffset(MboxMail *cr1, MboxMail *cr2) {
	return (cr1->m_startOff < cr2->m_startOff);
}
bool sortByFileOffsetDesc(MboxMail *cr1, MboxMail *cr2) {
	return (cr1->m_startOff > cr2->m_startOff);
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

void MboxMail::SortByFileOffset(CArray<MboxMail*, MboxMail*> *s_m, bool bDesc)
{
	if (s_m == 0) s_m = &MboxMail::s_mails;
	std::sort(s_m->GetData(), s_m->GetData() + s_m->GetSize(), bDesc ? sortByFileOffsetDesc : sortByFileOffset);
}

void MboxMail::SortByConverstionGroups(CArray<MboxMail*, MboxMail*> *s_m, bool bDesc) 
{
	// Assumption here is that s_mails_ref array is already ordered by groups
	int arraySize = s_mails_ref.GetSize();
	MboxMail::s_mails.SetSize(arraySize);
	if (bDesc) // descending order
	{
		for (int i = 0; i < arraySize; i++)
		{
			s_mails[arraySize - i - 1] = s_mails_ref[i];
		}
	}
	else
		MboxMail::s_mails.Copy(MboxMail::s_mails_ref);
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

typedef struct {
	char *m_charset;
	char *m_alt_charset;
} Name2AlternateName;

Name2AlternateName name2altname[] = {
	//{ "Name" , Info.Name },
	//{ "gbk" , "gb2312" },       // { 936 , "gb2312" , "Chinese Simplified (GB2312)" },
	//{ "gbk" , "x-cp20936" },  // { 20936 , "x-cp20936" , "Chinese Simplified (GB2312-80)" },
	//{ "gbk" , "hz-gb-2312" }, // { 52936 , "hz-gb-2312" , "Chinese Simplified (HZ)" },
	{ "gbk" , "GB18030" },    // { 54936 , "GB18030" , "Chinese Simplified (GB18030)" },
	{ "cp1252" , "Windows-1252" },    // { 1252 , "Windows-1252" , "Western European (Windows)" },
	{ "ascii" , "us-ascii" },    // { 20127 , "us-ascii" , "US-ASCII" },
	{ "cp819" , "iso-8859-1" },    // { 28591 , "iso-8859-1" , "Western European (ISO)" },
	{ "latin1" , "iso-8859-1" },    // { 28591 , "iso-8859-1" , "Western European (ISO)" },
	{ "latin2" , "iso-8859-2" },    // { 28592 , "iso-8859-2" , "Central European (ISO)" },
	{ "big-5" , "big5" },    // { 950 , "big5" , "Chinese Traditional (Big5)" },
	{ "SJIS" , "shift_jis" },    // { 932 , "shift_jis" , "Japanese (Shift-JIS)" },
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
	else
	{
		int i;
		int nsize = sizeof(name2altname) / sizeof(Name2AlternateName);
		for (i = 0; i < nsize; i++)
		{
			if (strcmp(charset.c_str(), name2altname[i].m_charset) == 0)
			{
				std::string alt_charset = name2altname[i].m_alt_charset;
				transform(alt_charset.begin(), alt_charset.end(), alt_charset.begin(), ::tolower);
				if ((it = cids->find(alt_charset)) != cids->end())
					id = it->second;
			}
		}
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


// Function to help investigate implementation issues
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

	SimpleString *outbuf = MboxMail::m_outbuf;
	outbuf->ClearAndResize(10000);
	SimpleString *inbuf = MboxMail::m_inbuf;
	inbuf->ClearAndResize(10000);

#if 0
	// log raw data
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
		txt.Format(_T("\n\ntype=%s encoding=%s distribution=%s attachment=%s offset=%d length=%d\n"),
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

		inbuf->ClearAndResize(bytestoRead);

		lpNumberOfBytesRead = 0;
		while (bytestoRead > 0)
		{
			retval = ReadFile(mbox_hFile, inbuf->Data(inbuf->Count()), bytestoRead, &lpNumberOfBytesRead, 0);
			if (retval != TRUE) {
				DWORD retval = GetLastError();
				break;
			}
			inbuf->SetCount(lpNumberOfBytesRead);
			bytestoRead -= lpNumberOfBytesRead;
		}

		if (body->m_contentTransferEncoding.CompareNoCase("base64") == 0) 
		{
			MboxCMimeCodeBase64 d64 = MboxCMimeCodeBase64(inbuf->Data(), inbuf->Count());

			int dlength = d64.GetOutputLength();
			int needLength = dlength;
			outbuf->ClearAndResize(needLength);

			int retlen = d64.GetOutput((unsigned char*)outbuf->Data(), needLength);
			if (retlen > 0)
				outbuf->SetCount(retlen);
			else
				outbuf->Clear();

			nwritten = 0;
			if (!Write2File(hFile, outbuf->Data(), outbuf->Count(), &nwritten)) {
				DWORD retval = GetLastError();
			}

			if (!Write2File(hFile, CRLF, strlen(CRLF), &nwritten)) {
				DWORD retval = GetLastError();
			}
			deb = 1;
		}
		else if (body->m_contentTransferEncoding.CompareNoCase("quoted-printable") == 0)
		{
			MboxCMimeCodeQP dGP = MboxCMimeCodeQP(inbuf->Data(), inbuf->Count());

			int dlength = dGP.GetOutputLength();
			int needLength = dlength;
			outbuf->ClearAndResize(needLength);

			int retlen = dGP.GetOutput((unsigned char*)outbuf->Data(), needLength);
			if (retlen > 0)
				outbuf->SetCount(retlen);
			else
				outbuf->Clear();

			nwritten = 0;
			if (!Write2File(hFile, outbuf->Data(), outbuf->Count(), &nwritten)) {
				DWORD retval = GetLastError();
			}

			if (!Write2File(hFile, CRLF, strlen(CRLF), &nwritten)) {
				DWORD retval = GetLastError();
			}
			deb = 1;
		}
		else
		{
			nwritten = 0;
			if (!Write2File(hFile, inbuf->Data(), inbuf->Count(), &nwritten)) {
				DWORD retval = GetLastError();
				break;
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
	p = EatNewLine(p, e);

	char *ss = line.GetBufferSetLength(p - p_beg);
	::memcpy(ss, p_beg, p - p_beg);
	line.Trim();

	if (line.IsEmpty())
		return p;

	while ((p < e) && ((*p == ' ') || (*p == '\t')))
	{
		char *p_next_beg = p;
		p = EatNewLine(p, e);

		CString nextLine;
		char *ss = nextLine.GetBufferSetLength(p - p_next_beg);
		::memcpy(ss, p_next_beg, p - p_next_beg);
		nextLine.Trim();
		line += " " + nextLine;
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

int GetMessageId(CString &fieldLine, int startPos, CString &value)
{
	int posEnd = fieldLine.FindOneOf(">;\n\r");
	if (posEnd < 0)  // indicates problem, '>' expected
		value = fieldLine.Mid(startPos);
	else if (fieldLine.GetAt(posEnd) == '>')
		value = fieldLine.Mid(startPos, posEnd - startPos + 1);
	else
		value = fieldLine.Mid(startPos, posEnd - startPos);
	value.Trim();
	return 1;
}

int GetParamValue(CString &fieldLine, int startPos, const char *param, int paramLen, CString &value)
{
	// TODO: it will break if ';' is part of the value. Need to handle quoted values containing ';' ?
	value.Empty();

	char *pbegin_sv = (char*)(LPCSTR)fieldLine;
	char *pend_sv = pbegin_sv + fieldLine.GetLength();
	char *p = pbegin_sv + startPos;

	// or keep it simple and require that *p == '=' at this point ?
	while (p < pend_sv)
	{
		//p = strstr(p, param);
		p = strnstrUpper2Lower(p, pend_sv, param, paramLen);
		if (p == 0)
			return 0;

		p = p + paramLen;
		char c = *p;
		if (c == '=')
			break;
		else if ((c != ' ') && (c != '\t'))
			continue;
	}

	char * pend = strchr(p, ';');
	if (pend == 0)
		pend = pend_sv;

	while (p < pend)
	{
		if (*p == '=')
			break;
		else
			p++;
	}
	if (p >= pend)
		return 0;

	char *posBegin = ++p;
	char *posEnd = pend;

	value = fieldLine.Mid(posBegin - pbegin_sv, posEnd - posBegin);

	value.Trim("\"\t ");

	return 1;
}

char *strnstrUpper2Lower(char *any, char *end, const char *lower, int lowerlength)
{
	// TODO: not very efficient; optimize
	char *p;
	for (p = any; p < (end - lowerlength) ; p++)
	{
		if (strncmpUpper2Lower(p, end, lower, lowerlength) == 0)
			return p;
	}
	return 0;
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

/* ParseContent Parser parses mail file to determine the file offset and length of content blocks.
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


char * MboxMail::ParseContent(MboxMail *mail, char *startPos, char *endPos)
{
	int bodyCnt = 0;

	MboxMail *m = mail;

	int mailLength = m->m_length;
	int mLength = endPos - startPos; // mailLength and mLength should be equal

	char *mailBegin = startPos;
	char *bodyDataOrigin = mailBegin;

	MailBody *mBody = MailBody::CreateMailBody();
	mBody->Load(bodyDataOrigin, mailBegin, mailLength);

	m->m_messageId = mBody->m_MessageId;
	m->m_replyId = mBody->m_ReplyId;

	// Iterate all the descendant body parts
	MailBody::MailBodyList bodies;

	int nCount = mBody->GetBodyPartList(bodies);
	MailBody::MailBodyList::const_iterator it;

	for (it = bodies.begin(); it != bodies.end(); it++)
	{
		MailBody* pBP = *it;

		if(pBP->m_IsAttachment)
			m->m_hasAttachments = 1;

		// For now ignore non "text/plain" bodies
		//if (!pBP->m_IsTextPlain)
		if (!pBP->m_IsText)
			continue;

		bodyCnt++;

		int bodyLength = pBP->m_bodyDataLength;
		char *bodyBegin = mailBegin + pBP->m_bodyDataOffset;

		if (pBP->m_bodyDataLength > 0)
		{
			MailBodyContent *contentDetails = new MailBodyContent;
			contentDetails->m_contentOffset = pBP->m_bodyDataOffset;
			contentDetails->m_contentLength = pBP->m_bodyDataLength;
			contentDetails->m_contentType = pBP->m_ContentType;
			contentDetails->m_contentDisposition = pBP->m_Disposition;
			contentDetails->m_contentTransferEncoding = pBP->m_TransferEncoding;
			contentDetails->m_attachmentName = pBP->m_AttachmentName;
			contentDetails->m_pageCode = pBP->m_PageCode;

			if ((pBP->m_PageCode == 0) && pBP->m_IsTextHtml)
				int deb = 1;

			m->m_ContentDetailsArray.push_back(contentDetails);
		}
	}
	// now it is save to free mBody
	MailBody::FreeMailBody(mBody);
	return mailBegin;  // not used, what should we return ?
}

void MboxMail::Destroy() 
{
	if (m_pMessageIdTable)
		m_pMessageIdTable->clear();

	for (int i = 0; i < s_mails.GetSize(); i++)
	{
		for (int j = 0; j < s_mails[i]->m_ContentDetailsArray.size(); j++) {
			delete s_mails[i]->m_ContentDetailsArray[j];
			s_mails[i]->m_ContentDetailsArray[j] = 0;
		}
		delete s_mails[i];
	}
	s_mails.RemoveAll();
	s_mails_ref.RemoveAll();
	b_mails_sorted = false;
	MboxMail::s_path = "";

};

unsigned long StrHash(const char* buf, const UINT length)
{
	register unsigned long hash = 0;
	register UINT len = length;
	while (len-- > 0)
	{
		hash = (hash << 4) + (unsigned long)*buf++;
		hash ^= (hash >> 12);
	}
	return hash;
}

//
int MboxMail::getMessageId(CString *key)
{
	MboxMail *m = 0;
	MessageIdTableType::iterator it;
	if (m_pMessageIdTable == 0)
		createMessageIdTable(10000);

	it = m_pMessageIdTable->find(key);
	if (it != m_pMessageIdTable->cend())
		return it->second;
	else
		return -1;
}

bool MboxMail::insertMessageId(CString *key, int val)
{
	CString *mapKey = new CString;
	*mapKey = *key;
	std::pair<MessageIdTableType::iterator, bool> result = m_pMessageIdTable->insert(MessageIdTableType::value_type(mapKey, val));
	return result.second;
}

int MboxMail::getReplyId(CString *key)
{
	MboxMail *m = 0;
	MessageIdTableType::iterator it;
	if (m_pMessageIdTable == 0)
		createMessageIdTable(10000);

	it = m_pMessageIdTable->find(key);
	if (it != m_pMessageIdTable->cend())
		return it->second;
	else
		return -1;
}

UINT MboxMail::createMessageIdTable(UINT count)
{
	m_pMessageIdTable = new MessageIdTableType;
	m_pMessageIdTable->reserve(count);
	return count;
}

int MboxMail::add2ConversationGroup(int mid, MboxMail *m)
{
	// TODO: ??
	return 1;
}

bool MboxMail::preprocessConversations()
{
	MboxMail *m;
	MboxMail *m_next;
	MboxMail *m_reply_to;

	MboxMail::SortByDate();

	// !! s_mails is already populated from mboxview file
	// Create links between all related messages based on messageId and replyId (from mail header)
	// Ignore duplicate messageId
	for (int i = 0; i < s_mails.GetSize(); i++)
	{
		m = s_mails[i];

		int id = i;
		if (m->m_messageId.GetLength())
		{
			int mId = getMessageId(&m->m_messageId);
			if (mId < 0) {
				insertMessageId(&m->m_messageId, id);
				m->m_duplicateId = false;
			}
			else {
				m->m_duplicateId = true;
			}
		}
		else
			int deb = 1;
	}

	int orphanCount = 0;
	// Link all mails with related maessageIds and replyIds
	for (int i = 0; i < s_mails.GetSize(); i++)
	{
		m = s_mails[i];

		if (m->m_messageId == m->m_replyId) // just in case
			continue;

		if (m->m_replyId.GetLength())
		{
			int rId = getReplyId(&m->m_replyId);
			if (rId < 0) {
				orphanCount++;
			}
			else 
			{
				m_reply_to = s_mails[rId];
				// find the last mail linked to reply mail
				while (m_reply_to->m_nextMail >= 0)
				{
					rId = m_reply_to->m_nextMail;
					m_reply_to = s_mails[m_reply_to->m_nextMail];
				}
				//link this mail to reply mail (directly or indirectly)
				int next = m_reply_to->m_nextMail;
				m_reply_to->m_nextMail = i;
				m->m_prevMail = rId;
				if (next >= 0) {
					m->m_nextMail = next;
					m_next = s_mails[next];
					m_next->m_prevMail = i;
				}
				int deb = 1;
			}
		}
	}
	return true;
}

// traverse s_mails by conversations and copy mail groups into s_mails_ref.
// Finally, copy result from s_mails_ref to s_mails. Upon return, both s_mails and s_mails_ref are equal.
bool MboxMail::sortConversations()
{
	MboxMail *m;
	MboxMail *m_sav;
	int i_ref = 0;
	int i_ref_begin, i_ref_end;

	s_mails_ref.SetSize(s_mails.GetSize());

	if (!validateSortConversations()) {
		//s_mails is already left sorted by date
		return false;
	}

	int currentGroupColor = 0; //white 

	// init s_mails_ref by conversation groups
	for (int i = 0; i < s_mails.GetSize(); i++)
	{
		m = s_mails[i];
		if (m->m_done)
			continue;

		i_ref_begin = i_ref;
		m->m_groupId = m_nextGroupId;
		s_mails_ref[i_ref++] = m;
		m->m_done = true;

		if (m->m_messageId == m->m_replyId) {
			m_nextGroupId++;
			m->m_groupColor = 0;
			continue;
		}
		currentGroupColor++;
		if (currentGroupColor > 2)
			currentGroupColor = 1;

		m->m_groupColor = currentGroupColor;

		m_sav = m;
		while (m->m_nextMail >= 0)
		{
			m = s_mails[m->m_nextMail];
			if (m->m_done)
				continue;  // should never be here ??
			m->m_groupId = m_nextGroupId;
			m->m_groupColor = currentGroupColor;
			s_mails_ref[i_ref++] = m;
			m->m_done = true;
		}
		m = m_sav;
		while (m->m_prevMail >= 0)
		{
			m = s_mails[m->m_prevMail];
			if (m->m_done)
				continue; // should never be here ??
			m->m_groupId = m_nextGroupId;
			m->m_groupColor = currentGroupColor;
			s_mails_ref[i_ref++] = m;
			m->m_done = true;  
		}
		i_ref_end = i_ref;
		m_nextGroupId++;

		int j;

		// check sorting order
		bool outOfOrder = false;

		for (j = i_ref_begin; j < (i_ref_end - 1); j++)
		{
			MboxMail *s_mails_ref_begin = s_mails_ref[j];
			MboxMail *s_mails_ref_end = s_mails_ref[j + 1];
			time_t timeBegin = s_mails_ref_begin->m_timeDate;
			time_t timeEnd = s_mails_ref_end->m_timeDate;

			if (timeEnd < timeBegin) {
				outOfOrder = true;
				break;
			}
		}

		if (outOfOrder)
		{
			for (j = i_ref_begin; j < (i_ref_end - 1); j++)
			{
				MboxMail *s_mails_ref_begin = s_mails_ref[j];
				MboxMail *s_mails_ref_end = s_mails_ref[j + 1];
				time_t timeBegin = s_mails_ref_begin->m_timeDate;
				time_t timeEnd = s_mails_ref_end->m_timeDate;

				if (timeEnd < timeBegin) {
					outOfOrder = true;
					break;
				}
			}

			; // reorder mails i_ref_begin < i_ref_end
			CArray<MboxMail*, MboxMail*> s_group_mails;

			for (j = i_ref_begin; j < i_ref_end; j++)
			{
				s_group_mails.Add(s_mails_ref[j]);
			}

			MboxMail::SortByDate(&s_group_mails);

			j = i_ref_begin;
			for (int i = 0; i < s_group_mails.GetSize(); i++)
			{
				time_t mtime = s_group_mails[i]->m_timeDate;

				s_mails_ref[j++] = s_group_mails[i];
			}

			// Recheck mail order
			for (j = i_ref_begin; j < (i_ref_end - 1); j++)
			{
				MboxMail *s_mails_ref_begin = s_mails_ref[j];
				MboxMail *s_mails_ref_end = s_mails_ref[j + 1];
				time_t timeBegin = s_mails_ref_begin->m_timeDate;
				time_t timeEnd = s_mails_ref_end->m_timeDate;

				if (timeEnd < timeBegin) {
					outOfOrder = true;
					break;
				}
			}
			s_group_mails.RemoveAll();
		}
	}
	if (s_mails.GetSize() != i_ref) {
		return false;
	}
	else {
		s_mails.Copy(s_mails_ref);
		return true;
	}
}


// verify that after al mails are accounted after sorting by group
// i.e. all mails will end up in s_mails_ref
bool MboxMail::validateSortConversations()
{
	MboxMail *m;
	MboxMail *m_sav;
	int i_ref = 0;

	for (int i = 0; i < s_mails.GetSize(); i++)
	{
		m = s_mails[i];
		if (m->m_done)
			continue;

		i_ref++;
		m->m_done = true;

		if (m->m_messageId == m->m_replyId)
			continue;

		m_sav = m;
		while (m->m_nextMail >= 0)
		{
			m = s_mails[m->m_nextMail];
			if (m->m_done)
				continue; // should never be here ??
			i_ref++;
			m->m_done = true;
		}
		m = m_sav;
		while (m->m_prevMail >= 0)
		{
			m = s_mails[m->m_prevMail];
			if (m->m_done)
				continue;  // should never be here ??
			i_ref++;
			m->m_done = true;
		}
	}

	// restore m->m_done
	for (int i = 0; i < s_mails.GetSize(); i++)
	{
		m = s_mails[i];
		m->m_done = false;
	}

	if (s_mails.GetSize() != i_ref)
		return false; // failure
	else
		return true;
}

void MboxMail::assignColor2ConvesationGroups()
{
	MboxMail *m;
	MboxMail *m_next;
	int i_ref = 0;

	int groupColor = 0;
	int i;
	for (i = 0; i < s_mails_ref.GetSize() -1; )
	{
		m = s_mails_ref[i];
		m_next = s_mails_ref[i+1];
		if (m->m_groupId == m_next->m_groupId) {
			groupColor++;
			if (groupColor > 2)
				groupColor = 1;
			for (++i; i < s_mails_ref.GetSize(); i++)
			{
				m = s_mails_ref[i];
				m->m_groupColor = groupColor;
			}
		}
		else
		{
			m->m_groupColor = 0;
		}
	}
	if (i < s_mails_ref.GetSize()) {
		i = s_mails_ref.GetSize() - 1;
		m = s_mails_ref[i];
		m->m_groupColor = groupColor;
	}
}

int MboxMail::exportHeaderToCSVFile(CSVFILE_CONFIG &csvConfig, CFile &fp)
{
	CString colLabels;
	bool separatorNeeded;

	separatorNeeded = false;

	if (csvConfig.m_bDate) {
		if (separatorNeeded)
			colLabels += csvConfig.m_separator;
		colLabels += "Date";
		separatorNeeded = true;
	}

	if (csvConfig.m_bFrom) {
		if (separatorNeeded)
			colLabels += csvConfig.m_separator;
		colLabels += "Name";
		colLabels += csvConfig.m_separator;
		colLabels += "From";
		separatorNeeded = true;
	}
	if (csvConfig.m_bTo) {
		if (separatorNeeded)
			colLabels += csvConfig.m_separator;
		colLabels += "To";
		separatorNeeded = true;
	}

	if (csvConfig.m_bSubject) {
		if (separatorNeeded)
			colLabels += csvConfig.m_separator;
		colLabels += "Subject";
		separatorNeeded = true;
	}

	if (csvConfig.m_bContent) {
		if (separatorNeeded)
			colLabels += csvConfig.m_separator;
		colLabels += "Message";
		separatorNeeded = true;
	}

	colLabels += "\r\n";

	fp.Write(colLabels, colLabels.GetLength());

	return 1;
}

// export mails to CSV file by leveraging index/mboxview data.
// Header fields from the index/mboview files are used.
int MboxMail::exportToCSVFile(CSVFILE_CONFIG &csvConfig, CString &csvFileName)
{
	MboxMail *m;
	CString colLabels;
	CString separator;
	bool separatorNeeded;

	CString mailFile = MboxMail::s_path;

	if (s_path.IsEmpty()) {
		CString txt = _T("Please open mail file first.");
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	if (path.IsEmpty())
		return -1;  // Hopefully s_path wil fail first

	TCHAR ext[_MAX_EXT + 1]; ext[0] = 0;
	TCHAR drive[_MAX_DRIVE + 1]; drive[0] = 0;
	TCHAR dir[_MAX_DIR + 1]; dir[0] = 0;
	TCHAR fname[_MAX_FNAME + 1];

	_tsplitpath(mailFile, drive, dir, fname, ext);

	CFile fp;
	CString csvFile = path + "\\" + CString(fname) + ".csv";

	if (!fp.Open(csvFile, CFile::modeWrite | CFile::modeCreate)) {
		CString txt = _T("Could not create \"") + csvFile;
		txt += _T("\" file.\nMake sure file is not open on other applications.");
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CFile fpm;
	if (csvConfig.m_bContent)
	{
		if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite)) {
			return -1;
		}
	}

	exportHeaderToCSVFile(csvConfig, fp);

	char buff[10000];
	char *token = 0;
	char *seppos;
	char *atpos;
	char *last;
	int tokenlen = 0;

	char sepchar = csvConfig.m_separator.GetAt(0);
	char *sepstr = (char*)(LPCSTR)csvConfig.m_separator;

	SimpleString outbuf(1024,256);
	SimpleString tmpbuf(1024, 256);
	SimpleString name(256);
	SimpleString addr(256);

	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;
	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);

	for (int i = 0; i < s_mails.GetSize(); i++)
	{
		m = s_mails[i];

		separatorNeeded = false;

		outbuf.Clear();

		if (csvConfig.m_bDate)
		{
			SYSTEMTIME st;
			SYSTEMTIME lst;
			char datebuff[32];

			CString format = GetDateFormat(csvConfig.m_dateFormat);

			datebuff[0] = 0;
			if (m->m_timeDate > 0)
			{
				CTime tt(m->m_timeDate);
				if (!csvConfig.m_bGMTTime) {
					bool ret = tt.GetAsSystemTime(st);
					SystemTimeToTzSpecificLocalTime(0, &st, &lst);
					CTime ltt(lst);
					strcpy(datebuff, (LPCSTR)ltt.Format(format));
				}
				else {
					strcpy(datebuff, (LPCSTR)tt.Format(format));
				}
			}

			if (separatorNeeded)
				outbuf.Append(sepchar);

			outbuf.Append('"');

			int datelen = strlen(datebuff);
			if (charCount(datebuff, '"') > 0)
			{
				tmpbuf.ClearAndResize(2 * datelen);
				int ret_datelen = escapeSeparators(tmpbuf.Data(), datebuff, datelen, '"');
				outbuf.Append(tmpbuf.Data(), ret_datelen);
			}
			else
				outbuf.Append(datebuff, datelen);

			outbuf.Append('"');

			separatorNeeded = true;
		}

		if (csvConfig.m_bFrom) 
		{
			int fromlen = m->m_from.GetLength();
			name.ClearAndResize(fromlen);
			addr.ClearAndResize(fromlen);
			splitMailAddress(m->m_from, fromlen, &name, &addr);

			if (separatorNeeded)
				outbuf.Append(sepchar);

			outbuf.Append('"');

			int namelen = name.Count();

			tmpbuf.ClearAndResize(2 * namelen);
			int ret_namelen = escapeSeparators(tmpbuf.Data(), name.Data(), namelen, '"');
			tmpbuf.SetCount(ret_namelen);

			UINT pageCode = m->m_from_charsetId;
			if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
			{
				BOOL ret = Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
				if (ret)
					outbuf.Append(*inbuf);
				else
					outbuf.Append(tmpbuf);
			}
			else {
				outbuf.Append(tmpbuf);
			}


			outbuf.Append('"');
			outbuf.Append(sepchar);

			outbuf.Append('"');
			int addrlen = addr.Count();

			tmpbuf.ClearAndResize(2 * addrlen);
			int ret_addrlen = escapeSeparators(tmpbuf.Data(), addr.Data(), addrlen, '"');
			tmpbuf.SetCount(ret_addrlen);

			pageCode = m->m_from_charsetId;
			if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
			{
				BOOL ret = Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
				if (ret)
					outbuf.Append(*inbuf);
				else
					outbuf.Append(tmpbuf);
			}
			else
				outbuf.Append(tmpbuf);

			outbuf.Append('"');

			separatorNeeded = true;
		}

		if (csvConfig.m_bTo)
		{
			int tolen = m->m_to.GetLength();
			memcpy(buff, m->m_to, tolen);
			buff[tolen] = 0;

			last = &buff[tolen];

			if (separatorNeeded)
				outbuf.Append(sepchar);

			outbuf.Append('"');

			token = buff;
			atpos = _tcschr(token, '@');
			while (atpos != 0)
			{
				seppos = _tcschr(atpos, sepchar);
				if (seppos == 0)
				{
					tokenlen = last - token;
					name.ClearAndResize(tokenlen);
					addr.ClearAndResize(tokenlen);
					splitMailAddress(token, tokenlen, &name, &addr);
					atpos = 0;
				}
				else
				{
					*seppos = 0;
					tokenlen = seppos - token;

					name.ClearAndResize(tokenlen);
					addr.ClearAndResize(tokenlen);
					splitMailAddress(token, tokenlen, &name, &addr);

					token = seppos + 1;
					atpos = _tcschr(token, '@');
				}
				
				int addrlen = addr.Count();

				tmpbuf.ClearAndResize(2 * addrlen);
				int ret_addrlen = escapeSeparators(tmpbuf.Data(), addr.Data(), addrlen, '"');
				tmpbuf.SetCount(ret_addrlen);

				UINT pageCode = m->m_to_charsetId;
				if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
				{
					BOOL ret = Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
					if (ret)
						outbuf.Append(*inbuf);
					else
						outbuf.Append(tmpbuf);
				}
				else
					outbuf.Append(tmpbuf);

				if (atpos)
					outbuf.Append(sepchar);
			}
			outbuf.Append('"');

			separatorNeeded = true;
		}

		if (csvConfig.m_bSubject) 
		{
			if (separatorNeeded)
				outbuf.Append(sepchar);

			outbuf.Append('"');

			int subjlen = m->m_subj.GetLength();
			char *subjstr = (char *)(LPCSTR)m->m_subj;

			tmpbuf.ClearAndResize(2 * subjlen);
			int ret_subjlen = escapeSeparators(tmpbuf.Data(), subjstr, subjlen, '"');
			tmpbuf.SetCount(ret_subjlen);

			UINT pageCode = m->m_subj_charsetId;
			if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
			{
				BOOL ret = Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
				if (ret)
					outbuf.Append(*inbuf);
				else
					outbuf.Append(tmpbuf);
			}
			else
				outbuf.Append(tmpbuf);

			outbuf.Append('"');

			separatorNeeded = true;
		}

		if (csvConfig.m_bContent)
		{
			if (separatorNeeded)
				outbuf.Append(sepchar);

			outbuf.Append('"');

			fp.Write(outbuf.Data(), outbuf.Count());

			MboxMail::m_outbuf->Clear();
			UINT pageCode = 0;
			//int textlen = GetMailBody_MailBody(fpm, i, MboxMail::m_outbuf, pageCode);  // slow, to compare slow and fast
			int textlen = GetMailBody_mboxview(fpm, i, MboxMail::m_outbuf, pageCode);  // fast
			if (textlen != MboxMail::m_outbuf->Count())
				int deb = 1;

			if (MboxMail::m_outbuf->Count())
			{
				int needLength = MboxMail::m_outbuf->Count() * 2 + 1; // worst case scenario or get '"' count first
				inbuf->ClearAndResize(needLength);  // escapeSeparators terminates result with null

				if (MboxMail::m_outbuf->Data()[0] != '\n')
					inbuf->Append("\n");
				int cnt_sv = inbuf->Count();
				int retcnt = escapeSeparators(inbuf->Data(cnt_sv), MboxMail::m_outbuf->Data(), MboxMail::m_outbuf->Count(), '"');
				inbuf->SetCount(retcnt + cnt_sv);

				if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
				{
					BOOL ret = Str2CodePage(inbuf, pageCode, csvConfig.m_nCodePageId, MboxMail::m_outbuf, workbuf);
					if (ret) {
						fp.Write(MboxMail::m_outbuf->Data(), MboxMail::m_outbuf->Count());
					}
					else {
						fp.Write(inbuf->Data(), inbuf->Count());
					}
				}
				else
					fp.Write(inbuf->Data(), inbuf->Count());
			}

			outbuf.Clear();
			outbuf.Append('"');

			separatorNeeded = true;
		}

		outbuf.Append("\r\n");

		fp.Write(outbuf.Data(), outbuf.Count());
	}
	fp.Close();
	if (csvConfig.m_bContent)
		fpm.Close();

	csvFileName = csvFile;
	return 1;
}

int MboxMail::printSingleMailToTextFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig)
{
	static char *border = "--------------------------------------------------------------------------------\r\n";
	char *token = 0;
	int tokenlen = 0;

	char sepchar = textConfig.m_separator.GetAt(0);
	char *sepstr = (char*)(LPCSTR)textConfig.m_separator;

	SimpleString outbuf(1024, 256);
	SimpleString tmpbuf(1024, 256);

	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;
	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);

	MboxMail *m = s_mails[mailPosition];


	outbuf.Clear();
	outbuf.Append(border);

	/////
	int subjlen = m->m_subj.GetLength();
	char *subjstr = (char *)(LPCSTR)m->m_subj;

	char *cSubject = "SUBJECT: ";
	int cSubjectLen = strlen(cSubject);
	outbuf.Append(cSubject, cSubjectLen);

	UINT pageCode = m->m_subj_charsetId;
	if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
	{
		tmpbuf.Copy(subjstr, subjlen);
		BOOL ret = Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
		if (ret)
			outbuf.Append(*inbuf);
		else
			outbuf.Append(tmpbuf);
	}
	else {
		outbuf.Append(subjstr, subjlen);
	}
	outbuf.Append("\r\n");
	//////
	int fromlen = m->m_from.GetLength();
	char *fromstr = (char *)(LPCSTR)m->m_from;

	char *cFrom = "FROM: ";
	int cFromLen = strlen(cFrom);
	outbuf.Append(cFrom, cFromLen);

	pageCode = m->m_from_charsetId;
	if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
	{
		tmpbuf.Copy(fromstr, fromlen);
		BOOL ret = Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
		if (ret)
			outbuf.Append(*inbuf);
		else
			outbuf.Append(tmpbuf);
	}
	else {
		outbuf.Append(fromstr, fromlen);
	}
	outbuf.Append("\r\n");
	//////
	int tolen = m->m_to.GetLength();
	char *tostr = (char *)(LPCSTR)m->m_to;

	char *cTo = "TO: ";
	int cToLen = strlen(cTo);
	outbuf.Append(cTo, cToLen);

	pageCode = m->m_to_charsetId;
	if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
	{
		tmpbuf.Copy(tostr, tolen);
		BOOL ret = Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
		if (ret)
			outbuf.Append(*inbuf);
		else
			outbuf.Append(tmpbuf);
	}
	else {
		outbuf.Append(tostr, tolen);
	}
	outbuf.Append("\r\n");
	////
	SYSTEMTIME st;
	SYSTEMTIME lst;
	char datebuff[32];

	CString format = textConfig.m_dateFormat;

	datebuff[0] = 0;
	if (m->m_timeDate > 0)
	{
		CTime tt(m->m_timeDate);
		if (!textConfig.m_bGMTTime) {
			bool ret = tt.GetAsSystemTime(st);
			SystemTimeToTzSpecificLocalTime(0, &st, &lst);
			CTime ltt(lst);
			strcpy(datebuff, (LPCSTR)ltt.Format(format));
		}
		else {
			strcpy(datebuff, (LPCSTR)tt.Format(format));
		}
	}

	char *cDate = "DATE: ";
	int cDateLen = strlen(cDate);
	outbuf.Append(cDate, cDateLen);

	int datelen = strlen(datebuff);
	outbuf.Append(datebuff, datelen);

	outbuf.Append("\r\n");

	outbuf.Append(border);

	outbuf.Append("\r\n");
	////

	fp.Write(outbuf.Data(), outbuf.Count());

	//
	MboxMail::m_outbuf->Clear();
	pageCode = 0;
	int textlen = GetMailBody_mboxview(fpm, mailPosition, MboxMail::m_outbuf, pageCode);  // returns pageCode
	if (textlen != MboxMail::m_outbuf->Count())
		int deb = 1;

	if (MboxMail::m_outbuf->Count())
	{
		if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
		{
			int needLength = MboxMail::m_outbuf->Count() * 2 + 1;
			inbuf->ClearAndResize(needLength);

			BOOL ret = Str2CodePage(MboxMail::m_outbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
			if (ret) {
				fp.Write(inbuf->Data(), inbuf->Count());
			}
			else {
				fp.Write(MboxMail::m_outbuf->Data(), MboxMail::m_outbuf->Count());
			}
		}
		else
			fp.Write(MboxMail::m_outbuf->Data(), MboxMail::m_outbuf->Count());
	}

	outbuf.Clear();
	outbuf.Append("\r\n");

	fp.Write(outbuf.Data(), outbuf.Count());

	return 1;
}

UINT getCodePageFrmHtmlBody(SimpleString *buffer)
{
	char *pat = "charset=";
	std::string charset;
	UINT codePage = 0;
	int patlen = strlen(pat);
	int posEnd = -1;
	int posBegin = buffer->FindNoCase(0, pat, patlen);
	if (posBegin >= 0)
		posBegin = buffer->FindAny(posBegin, "=");
	if (posBegin >= 0) {
		posEnd = buffer->FindAny(posBegin, "\";\r\n");
	}

	if (posEnd >= 0) {

		int valLen = posEnd - posBegin;
		charset.assign(buffer->Data(posBegin+1), valLen-1);
	}
	if (!charset.empty()) {
		codePage = charset2Id(charset.c_str());
	}
	int deb = 1;
	return codePage;
}


int MboxMail::printSingleMailToHtmlFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig)
{
	char *token = 0;
	int tokenlen = 0;

	SimpleString outbuf(1024, 256);
	SimpleString tmpbuf(1024, 256);

	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;
	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);

	MboxMail *m = s_mails[mailPosition];

	outbuf.Clear();

	SimpleString *outbuflarge = MboxMail::m_outbuf;

	outbuflarge->Clear();
	UINT pageCode = 0;
	int textType = 1; // try first Html
	int textlen = GetMailBody_mboxview(fpm, mailPosition, outbuflarge, pageCode, textType);  // returns pageCode
	if (textlen != MboxMail::m_outbuf->Count())
		int deb = 1;

	if (outbuflarge->Count() != 0)
	{
		CString bdycharset = "UTF-8";
		CString bdy;

		bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head>";
		bdy += "<body><pre style=\"background-color:#eee9e9;font-weight:normal;\"><font size=\"+1\">";
		fp.Write(bdy, bdy.GetLength());

		int ret = printMailHeaderToHtmlFile(fp, mailPosition, fpm, textConfig);

		bdy = "</font></pre></body></html>\r\n";
		fp.Write(bdy, bdy.GetLength());

		bool extraHtmlHdr = false;
		if (outbuflarge->FindNoCase(0, "<body", 5) < 0) // didn't find if true
		{
			extraHtmlHdr = true;
			CString bdycharset = "UTF-8";
			bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body>";
			fp.Write(bdy, bdy.GetLength());
		}

		if (pageCode == 0) {
			pageCode = getCodePageFrmHtmlBody(outbuflarge);
		}

		if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
		{
			int needLength = outbuflarge->Count() * 2 + 1;
			inbuf->ClearAndResize(needLength);

			BOOL ret = Str2CodePage(outbuflarge, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
			if (ret) {
				fp.Write(inbuf->Data(), inbuf->Count());
			}
			else {
				fp.Write(outbuflarge->Data(), outbuflarge->Count());
			}
		}
		else
			fp.Write(outbuflarge->Data(), outbuflarge->Count());


		if (extraHtmlHdr) {
			CString bdy = "</body></html>\r\n";
			fp.Write(bdy, bdy.GetLength());
		}
	}
	else
	{
		outbuflarge->Clear();
		pageCode = 0;
		textType = 0; // no Html, try to get Plain
		textlen = GetMailBody_mboxview(fpm, mailPosition, outbuflarge, pageCode, textType);  // returns pageCode
		if (textlen != outbuflarge->Count())
			int deb = 1;

		CString bdycharset = "UTF-8";
		CString bdy;
		bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head>";
		bdy += "<body><pre style=\"background-color:#eee9e9;font-weight:normal;\"><font size=\"+1\">";
		fp.Write(bdy, bdy.GetLength());

		int ret = printMailHeaderToHtmlFile(fp, mailPosition, fpm, textConfig);

		bdy = "</font></pre><pre>\r\n";
		fp.Write(bdy, bdy.GetLength());

		if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
		{
			int needLength = outbuflarge->Count() * 2 + 1;
			inbuf->ClearAndResize(needLength);

			BOOL ret = Str2CodePage(outbuflarge, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
			if (ret) {
				fp.Write(inbuf->Data(), inbuf->Count());
			}
			else {
				fp.Write(outbuflarge->Data(), outbuflarge->Count());
			}
		}
		else
			fp.Write(outbuflarge->Data(), outbuflarge->Count());

		bdy = "</pre></body></html>\r\n";
		fp.Write(bdy, bdy.GetLength());
	}

	outbuf.Clear();
	outbuf.Append("\r\n");

	fp.Write(outbuf.Data(), outbuf.Count());

	return 1;
}

void encodeTextAsHtml(SimpleString &txt) {
	SimpleString buffer(256);
	for (size_t pos = 0; pos != txt.Count(); ++pos) {
		char c = txt.GetAt(pos);
		switch (c) {
		case '&':  buffer.Append("&amp;");       break;
		case '\"': buffer.Append("&quot;");      break;
		case '\'': buffer.Append("&apos;");      break;
		case '<':  buffer.Append("&lt;");        break;
		case '>':  buffer.Append("&gt;");        break;
		default:   buffer.Append(c); break;
		}
	}
	txt.Copy(buffer);
}


int MboxMail::printMailHeaderToHtmlFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig)
{
	char *token = 0;
	int tokenlen = 0;

	SimpleString outbuf(1024, 256);
	SimpleString tmpbuf(1024, 256);

	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;
	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);

	MboxMail *m = s_mails[mailPosition];

	//
	int subjlen = m->m_subj.GetLength();
	char *subjstr = (char *)(LPCSTR)m->m_subj;

	char *cSubject = "SUBJECT: ";
	int cSubjectLen = strlen(cSubject);
	outbuf.Append(cSubject, cSubjectLen);
	tmpbuf.Copy(subjstr, subjlen);
	encodeTextAsHtml(tmpbuf);

	UINT pageCode = m->m_subj_charsetId;
	if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
	{
		BOOL ret = Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
		if (ret)
			outbuf.Append(*inbuf);
		else
			outbuf.Append(tmpbuf);
	}
	else {
		outbuf.Append(tmpbuf);
	}
	outbuf.Append("\r\n");

	//
	int fromlen = m->m_from.GetLength();
	char *fromstr = (char *)(LPCSTR)m->m_from;

	char *cFrom = "FROM: ";
	int cFromLen = strlen(cFrom);
	outbuf.Append(cFrom, cFromLen);
	tmpbuf.Copy(fromstr, fromlen);
	encodeTextAsHtml(tmpbuf);

	pageCode = m->m_from_charsetId;
	if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
	{
		BOOL ret = Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
		if (ret)
			outbuf.Append(*inbuf);
		else
			outbuf.Append(tmpbuf);
	}
	else {
		outbuf.Append(tmpbuf);
	}
	outbuf.Append("\r\n");

	//
	int tolen = m->m_to.GetLength();
	char *tostr = (char *)(LPCSTR)m->m_to;

	char *cTo = "TO: ";
	int cToLen = strlen(cTo);
	outbuf.Append(cTo, cToLen);
	tmpbuf.Copy(tostr, tolen);
	encodeTextAsHtml(tmpbuf);

	pageCode = m->m_to_charsetId;
	if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
	{
		BOOL ret = Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
		if (ret)
			outbuf.Append(*inbuf);
		else
			outbuf.Append(tmpbuf);
	}
	else {
		outbuf.Append(tmpbuf);
	}
	outbuf.Append("\r\n");

	//
	SYSTEMTIME st;
	SYSTEMTIME lst;
	char datebuff[32];

	CString format = textConfig.m_dateFormat;

	datebuff[0] = 0;
	if (m->m_timeDate > 0)
	{
		CTime tt(m->m_timeDate);
		if (!textConfig.m_bGMTTime) {
			bool ret = tt.GetAsSystemTime(st);
			SystemTimeToTzSpecificLocalTime(0, &st, &lst);
			CTime ltt(lst);
			strcpy(datebuff, (LPCSTR)ltt.Format(format));
		}
		else {
			strcpy(datebuff, (LPCSTR)tt.Format(format));
		}
	}

	char *cDate = "DATE: ";
	int cDateLen = strlen(cDate);
	outbuf.Append(cDate, cDateLen);

	int datelen = strlen(datebuff);
	outbuf.Append(datebuff, datelen);

	//outbuf.Append("\r\n\r\n");
	fp.Write(outbuf.Data(), outbuf.Count());

	outbuf.Clear();
	//outbuf.Append("\r\n");

	fp.Write(outbuf.Data(), outbuf.Count());

	return 1;
}

int MboxMail::exportToTextFile(TEXTFILE_CONFIG &textConfig, CString &textFileName, int firstMail, int lastMail, int textType)
{
	CString colLabels;
	CString separator;

	CString mailFile = MboxMail::s_path;

	if (s_path.IsEmpty()) {
		CString txt = _T("Please open mail file first.");
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	if (path.IsEmpty())
		return -1;  // Hopefully s_path wil fail first

	TCHAR ext[_MAX_EXT + 1]; ext[0] = 0;
	TCHAR drive[_MAX_DRIVE + 1]; drive[0] = 0;
	TCHAR dir[_MAX_DIR + 1]; dir[0] = 0;
	TCHAR fname[_MAX_FNAME + 1];

	_tsplitpath(mailFile, drive, dir, fname, ext);

	CFile fp;
	CString textFile;
	if (textType == 0)
		textFile = path + "\\" + CString(fname) + ".txt";
	else
		textFile = path + "\\" + CString(fname) + ".htm";

	if (!fp.Open(textFile, CFile::modeWrite | CFile::modeCreate)) {
		CString txt = _T("Could not create \"") + textFile;
		txt += _T("\" file.\nMake sure file is not open on other applications.");
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CFile fpm;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite)) {
		return -1;
	}

	if (textType == 0) {
		if (textConfig.m_nCodePageId == CP_UTF8) {
			const char *BOM_UTF8 = "\xEF\xBB\xBF";
			fp.Write(BOM_UTF8, 3);
		}
	}

	for (int i = firstMail; i <= lastMail; i++)
	{
		if (textType == 0)
			printSingleMailToTextFile(fp, i, fpm, textConfig);
		else
			printSingleMailToHtmlFile(fp, i, fpm, textConfig);
	}

	fp.Close();
	fpm.Close();

	textFileName = textFile;
	return 1;
}

int MboxMail::splitMailAddress(const char *buff, int bufflen, SimpleString *name, SimpleString *addr)
{
	register char *token_begin;
	register char *token_end;
	char *token_begin_sv;
	char *token_end_sv;

	register char *last = (char *)&buff[bufflen-1];
	char *first = (char*)&buff[0];

	token_begin = first;
	token_end = last;
	while ((token_begin <= token_end) && (
		(*token_end == ' ') || (*token_end == '\t') ||
		(*token_end == '>') || (*token_end == '<') || 
		(*token_end == '"') || (*token_end == '\''))
		)
		token_end--;
	token_end_sv = token_end;

	while ((token_begin <= token_end) && !(
		(*token_end == ' ') || (*token_end == '\t') ||
		(*token_end == '>') || (*token_end == '<') ||
		(*token_end == '"') || (*token_end == '\''))
		)
		token_end--;

	token_begin = token_end + 1;
	token_end = token_end_sv;

	int addrlen = token_end - token_begin + 1;
	int n = 0;
	char *to_addr = addr->Data();
	for (char *pos = token_begin; pos <= token_end; pos++, n++) 
		to_addr[n] = tolower(*pos);

	addr->SetCount(addrlen);

	token_begin_sv = token_begin;
	token_begin = first;
	token_end = token_begin_sv - 1;

	while ((token_begin <= token_end) && (
		(*token_end == ' ') || (*token_end == '\t') ||
		(*token_end == '>') || (*token_end == '<') ||
		(*token_end == '"') || (*token_end == '\''))
		)
		token_end--;

	while ((token_begin <= token_end) && (
		(*token_begin == ' ') || (*token_begin == '\t') ||
		(*token_begin == '>') || (*token_begin == '<') ||
		(*token_begin == '"') || (*token_begin == '\''))
		)
		token_begin++;

	int namelen = token_end - token_begin + 1;
	name->Copy(token_begin, namelen);

	return 1;
}

int MboxMail::nstrcpy(register char *dst, register char *src)
{
	char *src_sv = src;
	char c;
	while ((c = *src++)) { *dst++ = c; }
	*dst = 0;
	int len = src - src_sv - 1;
	return len;
}

int MboxMail::charCount(char *fld, char c)
{
	int cnt = 0;
	register char ch;
	while (ch = *fld++)
		if (ch == c) cnt++;
	return cnt;
}

int MboxMail::escapeSeparators(char *workbuff, char *fldstr, int fldlen, char sepchar)
{
	int outcnt = 0;
	for (int i = 0; i < fldlen; i++)
	{
		
		if (fldstr[i] == sepchar) {
			workbuff[outcnt++] = sepchar;
			workbuff[outcnt++] = sepchar;
		}
		else
			workbuff[outcnt++] = fldstr[i];

	}
	workbuff[outcnt] = 0;
	return outcnt;
}

// Get mail text by leveraging the file offset to the mail text body in index/mboxview file. 
// Full parsing of the mail is not done or rquired.
int MboxMail::GetMailBody_mboxview(CFile &fpm, int mailPosition, SimpleString *outbuf, UINT &pageCode, int textType)
{
	_int64 fileOffset;
	int bodyCnt = 0;

	MboxMail *m = MboxMail::s_mails[mailPosition];

	// We are using global buffers so check and assert if we collide. 
	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;
	ASSERT(outbuf != inbuf);
	ASSERT(outbuf != workbuf);

	outbuf->ClearAndResize(10000);
	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);

	MailBodyContent *body;
	pageCode = 0;
	bool pageCodeConflict = false;

	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];

		if (!body->m_attachmentName.IsEmpty()) {
				continue;
		}
		if (textType == 0) {
			if (body->m_contentType.CompareNoCase("text/plain") != 0)
				continue;
		} 
		else if (textType == 1) {
			if (body->m_contentType.CompareNoCase("text/html") != 0)
				continue;
		}

		bodyCnt++;
		if (pageCodeConflict == false) {
			if (pageCode == 0)
				pageCode = body->m_pageCode;
			else if (pageCode != body->m_pageCode)
				pageCodeConflict = true;
		}

		int bodyLength = body->m_contentLength;
		if ((body->m_contentOffset + body->m_contentLength) > m->m_length) {
			// something is not consistent
			bodyLength = m->m_length - body->m_contentOffset;
		}
		
		fileOffset = m->m_startOff + body->m_contentOffset;
		fpm.Seek(fileOffset, SEEK_SET);

		inbuf->ClearAndResize(bodyLength);

		char *bodyBegin = inbuf->Data();
		int retlen = fpm.Read(bodyBegin, bodyLength);
		if (retlen != bodyLength) {
			bodyLength = retlen;
		}

		if (body->m_contentTransferEncoding.CompareNoCase("base64") == 0)
		{
			MboxCMimeCodeBase64 d64(bodyBegin, bodyLength);
			int dlength = d64.GetOutputLength();
			int needLength = dlength + outbuf->Count();
			outbuf->Resize(needLength);

			char *outptr = outbuf->Data(outbuf->Count());
			int retlen = d64.GetOutput((unsigned char*)outptr, dlength);
			if (retlen > 0)
				outbuf->SetCount(outbuf->Count() + retlen);
		}
		else if (body->m_contentTransferEncoding.CompareNoCase("quoted-printable") == 0)
		{
			MboxCMimeCodeQP dGP(bodyBegin, bodyLength);
			int dlength = dGP.GetOutputLength();
			int needLength = dlength + outbuf->Count();
			outbuf->Resize(needLength);

			char *outptr = outbuf->Data(outbuf->Count());
			int retlen = dGP.GetOutput((unsigned char*)outptr, dlength);
			if (retlen > 0)
				outbuf->SetCount(outbuf->Count() + retlen);
		}
		else
		{
			// in case we have multiple bodies of the same type ?? not sure it is valid case/concern
			// asking for trouble ??
			outbuf->Append(bodyBegin, bodyLength);
		}

		if (bodyCnt > 1)
			int deb = 1;
	}

	if (pageCodeConflict)
		pageCode = 0;

	return outbuf->Count();
}

int SimpleString::FindNoCase(int offset, void const* Src, int  Size)
{
	int i;
	char *p;
	int count = m_count - offset;
	for (i = offset; i < (count - Size); i++)
	{
		p = &m_data[i];
		if (strncmpUpper2Lower(p, (m_count - i), (char*)Src, Size) == 0)
			return i;
	}
	return -1;
}

int SimpleString::FindAny(int offset, void const* Src)
{
	int i;
	const char *p = m_data;
	int c;
	for (i = offset; i < m_count; i++)
	{
		c = m_data[i];
		p = ::strchr((const char*)Src, c);
		if (p)
			return i;
	}
	return -1;
}

int SimpleString::Resize(int size)
{
	if (size > m_capacity)
	{
		int new_capacity = size + m_grow_size; // fudge factor ?
		char *new_data = new char[new_capacity + 1]; // one extra to set NULL
		if (new_data) {
			if (m_count > 0)
				::memcpy(new_data, m_data, m_count);
			delete[] m_data;
			m_data = new_data;
			m_capacity = new_capacity;
}
		else
			; // trouble :) caller needs to handle this ?
	}
	return m_capacity;
}

void SimpleString::append_internal(void const* Src, size_t  Size) {
	int spaceNeeded = m_count + Size;
	if (spaceNeeded > m_capacity)
		Resize(spaceNeeded);
	::memcpy(&m_data[m_count], Src, Size);
	SetCount(spaceNeeded);
}

class CMBodyHdr
{
public:
	CMimeMessage * Mail;
	bool IsText;
	bool IsMessage;
	bool IsAttachement;
	bool IsMultiPart;
	CString Content;
	int ContentLength;
	CString Charset;
	CString Description;
	CString Disposition;
	CString TransferEncoding;
	CString SubType;
	CString MainType;
	CString Boundary;
	CString ContentType;
	CMimeHeader::MediaType MediaType;
	CString Name;
	int BodiesCount;
	int TextBodiesCount;
	int HtmlTextBodiesCount;
};


// Debug helper
void MboxMail::getCMimeBodyHeader(CMimeMessage *mail, CMimeBody* pBP, CMBodyHdr *pHdr)
{
	pHdr->Mail = mail;
	pHdr->IsText = pBP->IsText();
	pHdr->IsMessage = pBP->IsMessage();
	pHdr->IsAttachement = pBP->IsAttachment();
	pHdr->IsMultiPart = pBP->IsMultiPart();
	pHdr->Content = (unsigned char*)pBP->GetContent();
	pHdr->ContentLength = pBP->GetContentLength();
	pHdr->Charset = pBP->GetCharset().c_str();
	pHdr->Description = pBP->GetDescription();
	pHdr->Disposition = pBP->GetDisposition();
	pHdr->TransferEncoding = pBP->GetTransferEncoding();
	pHdr->SubType = pBP->GetSubType().c_str();
	pHdr->MainType = pBP->GetMainType().c_str();
	pHdr->Boundary = pBP->GetBoundary().c_str();
	pHdr->ContentType = pBP->GetContentType();
	pHdr->MediaType = pBP->GetMediaType();
	pHdr->Name = pBP->GetName().c_str();
	pHdr->BodiesCount = 0;
	pHdr->TextBodiesCount = 0;
	pHdr->HtmlTextBodiesCount = 0;
	if (mail) 
	{
		// Iterate all the descendant body parts
		CMimeBody::CBodyList bodies;
		int nCount = mail->GetBodyPartList(bodies);
		CMimeBody::CBodyList::const_iterator it;
		for (it = bodies.begin(); it != bodies.end(); it++) {
			CMimeBody* pB = *it;
			pHdr->BodiesCount++;
			if (pB) {
				if (pB->IsText()) {
					pHdr->TextBodiesCount++;
					CString SubType = pB->GetSubType().c_str();
					if (SubType.CompareNoCase("html") == 0)
						pHdr->HtmlTextBodiesCount++;
				}
			}
		}

		if (pHdr->BodiesCount != nCount)
			int deb = 1;
		if (pHdr->TextBodiesCount != pHdr->HtmlTextBodiesCount)
			int deb = 1;
		if ((pHdr->TextBodiesCount > 0) && (pHdr->HtmlTextBodiesCount == 0))
			int deb = 1;

		int deb = 1;
	}
	if (pHdr->Mail) {
		if ((pHdr->TextBodiesCount == 0) && (pHdr->HtmlTextBodiesCount == 0))
			int deb = 1;
	}
	if (pHdr->Mail)
		int deb = 1;
	int deb = 1;
}


// export mails to CSV file by performing full parse of the mail.
// Header fields from the index/mboview files are not used.
int MboxMail::exportToCSVFileFullMailParse(CSVFILE_CONFIG &csvConfig)
{
	MboxMail *ml;
	MboxMail ml_tmp;
	MboxMail *m = &ml_tmp;
	CString colLabels;
	CString separator;
	bool separatorNeeded;

	CString mailFile = MboxMail::s_path;

	if (s_path.IsEmpty()) {
		CString txt = _T("Please open mail file first.");
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	if (path.IsEmpty())
		return -1;  // Hopefully s_path wil fail first

	TCHAR ext[_MAX_EXT + 1]; ext[0] = 0;
	TCHAR drive[_MAX_DRIVE + 1]; drive[0] = 0;
	TCHAR dir[_MAX_DIR + 1]; dir[0] = 0;
	TCHAR fname[_MAX_FNAME + 1];

	_tsplitpath(mailFile, drive, dir, fname, ext);

	CFile fp;
	CString csvFile = path + "\\" + CString(fname) + ".csv";

	if (!fp.Open(csvFile, CFile::modeWrite | CFile::modeCreate)) {
		CString txt = _T("Could not create \"") + csvFile;
		txt += _T("\" file.\nMake sure file is not open on other applications.");
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CFile fpm;

	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite)) {
		return -1;
	}

	exportHeaderToCSVFile(csvConfig, fp);

	char buff[10000];
	char *token = 0;
	char *seppos;
	char *atpos;
	char *last;
	int tokenlen = 0;

	char sepchar = csvConfig.m_separator.GetAt(0);
	char *sepstr = (char*)(LPCSTR)csvConfig.m_separator;

	SimpleString outbuf(1024, 256);
	SimpleString tmpbuf(1024);
	SimpleString name(256);
	SimpleString addr(256);

	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;
	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);


	CMBodyHdr cmh;
	for (int i = 0; i < s_mails.GetSize(); i++)
	{
		ml = s_mails[i];

		_int64 fileOffset = ml->m_startOff;
		fpm.Seek(fileOffset, SEEK_SET);

		int bodyLength = ml->m_length;
		inbuf->ClearAndResize(bodyLength);

		char *bodyBegin = inbuf->Data();
		int retlen = fpm.Read(bodyBegin, bodyLength);
		if (retlen != bodyLength) {
			bodyLength = retlen;
		}
		inbuf->SetCount(bodyLength);

		// Decode MIME message
		CMimeMessage mail;
		const char* bodyData = bodyBegin;
		int nLoadedSize = mail.Load(bodyData, bodyBegin, bodyLength);

		CMimeMessage::CMimeBody &body = mail;
		CMimeBody* pBP = &mail;

		//getCMimeBodyHeader(&mail, pBP, &cmh); // debug helper

		// Iterate all the header fields of mail body
		CMimeHeader::CFieldList& fds = body.Fields();
		CMimeHeader::CFieldList::const_iterator itfd;

		CString cstr;
		m->m_from.Empty(); m->m_to.Empty(); m->m_subj.Empty();
		m->m_from_charset.Empty(); m->m_to_charset.Empty(); m->m_subj_charset.Empty();
		m->m_from_charsetId = 0; m->m_to_charsetId = 0; m->m_subj_charsetId;

		
		for (itfd = fds.begin(); itfd != fds.end(); itfd++)
		{
			const CMimeField& fd = *itfd;
			const char *fname = fd.GetName();
			const char *fval = fd.GetValue();
			const char *charset = fd.GetCharset();
			UINT page_code = 0;
			if (charset && (charset[0] != 0)) {
				page_code = MboxMail::Str2PageCode(charset);
			}
			int deb = 1;
			cstr = fname;
			if (cstr.CompareNoCase("from") == 0) {
				m->m_from = fval;
				m->m_from.Replace("\n", "");
				m->m_from.Replace("\r", "");
				m->m_from_charsetId = page_code;
			}
			else if (cstr.CompareNoCase("to") == 0) {
				m->m_to = fval;
				m->m_to.Replace("\n", "");
				m->m_to.Replace("\r", "");
				m->m_to_charsetId = page_code;
			}
			else if (cstr.CompareNoCase("subject") == 0) {
				m->m_subj = fval;
				if (m->m_subj.Find("Microsoft account") >= 0)
					int deb = 1;
				m->m_subj.Replace("\n", "");
				m->m_subj.Replace("\r", "");
				//m->m_subj.TrimLeft(' ');
				//m->m_subj.TrimRight(' ');
				m->m_subj.Trim();
				m->m_subj_charsetId = page_code;
			}
			else if (cstr.CompareNoCase("date") == 0) 
			{
				SYSTEMTIME tm;
				if (DateParser::parseRFC822Date(fval, &tm)) {
					if (DateParser::validateSystemtime(&tm)) {
						CTime tt(tm);
						m->m_timeDate = tt.GetTime();
					}
				}
			}
		}
		if (m->m_timeDate == 0)
			int deb = 1;

		separatorNeeded = false;

		outbuf.Clear();

		if (csvConfig.m_bDate)
		{
			SYSTEMTIME st;
			SYSTEMTIME lst;
			char datebuff[32];

			CString format = GetDateFormat(csvConfig.m_dateFormat);

			datebuff[0] = 0;
			if (m->m_timeDate > 0)
			{
				CTime tt(m->m_timeDate);
				if (!csvConfig.m_bGMTTime) {
					bool ret = tt.GetAsSystemTime(st);
					SystemTimeToTzSpecificLocalTime(0, &st, &lst);
					CTime ltt(lst);
					strcpy(datebuff, (LPCSTR)ltt.Format(format));
				}
				else {
					strcpy(datebuff, (LPCSTR)tt.Format(format));
				}
			}

			if (separatorNeeded)
				outbuf.Append(sepchar);

			outbuf.Append('"');

			int datelen = strlen(datebuff);
			if (charCount(datebuff, '"') > 0)
			{
				tmpbuf.ClearAndResize(2 * datelen);
				int ret_datelen = escapeSeparators(tmpbuf.Data(), datebuff, datelen, '"');
				outbuf.Append(tmpbuf.Data(), ret_datelen);
			}
			else
				outbuf.Append(datebuff, datelen);

			outbuf.Append('"');

			separatorNeeded = true;
		}

		if (csvConfig.m_bFrom)
		{
			int fromlen = m->m_from.GetLength();
			name.ClearAndResize(fromlen);
			addr.ClearAndResize(fromlen);
			splitMailAddress(m->m_from, fromlen, &name, &addr);

			if (separatorNeeded)
				outbuf.Append(sepchar);

			outbuf.Append('"');

			int namelen = name.Count();

			tmpbuf.ClearAndResize(2 * namelen);
			int ret_namelen = escapeSeparators(tmpbuf.Data(), name.Data(), namelen, '"');
			tmpbuf.SetCount(ret_namelen);

			UINT pageCode = m->m_from_charsetId;
			if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
			{
				BOOL ret = Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
				if (ret)
					outbuf.Append(*inbuf);
				else
					outbuf.Append(tmpbuf);
			}
			else {
				outbuf.Append(tmpbuf);
			}

			outbuf.Append('"');
			outbuf.Append(sepchar);

			outbuf.Append('"');
			int addrlen = addr.Count();

			tmpbuf.ClearAndResize(2 * addrlen);
			int ret_addrlen = escapeSeparators(tmpbuf.Data(), addr.Data(), addrlen, '"');
			tmpbuf.SetCount(ret_addrlen);

			pageCode = m->m_from_charsetId;
			if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
			{
				BOOL ret = Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
				if (ret)
					outbuf.Append(*inbuf);
				else
					outbuf.Append(tmpbuf);
			}
			else
				outbuf.Append(tmpbuf);

			outbuf.Append('"');

			separatorNeeded = true;
		}

		if (csvConfig.m_bTo)
		{
			int tolen = m->m_to.GetLength();
			memcpy(buff, m->m_to, tolen);
			buff[tolen] = 0;

			last = &buff[tolen];

			if (separatorNeeded)
				outbuf.Append(sepchar);

			outbuf.Append('"');

			token = buff;
			atpos = _tcschr(token, '@');
			while (atpos != 0)
			{
				seppos = _tcschr(atpos, sepchar);
				if (seppos == 0)
				{
					tokenlen = last - token;
					name.ClearAndResize(tokenlen);
					addr.ClearAndResize(tokenlen);
					splitMailAddress(token, tokenlen, &name, &addr);
					atpos = 0;
				}
				else
				{
					*seppos = 0;
					tokenlen = seppos - token;

					name.ClearAndResize(tokenlen);
					addr.ClearAndResize(tokenlen);
					splitMailAddress(token, tokenlen, &name, &addr);

					token = seppos + 1;
					atpos = _tcschr(token, '@');
				}

				int addrlen = addr.Count();

				tmpbuf.ClearAndResize(2 * addrlen);
				int ret_addrlen = escapeSeparators(tmpbuf.Data(), addr.Data(), addrlen, '"');
				tmpbuf.SetCount(ret_addrlen);

				UINT pageCode = m->m_to_charsetId;
				if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
				{
					BOOL ret = Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
					if (ret)
						outbuf.Append(*inbuf);
					else
						outbuf.Append(tmpbuf);
				}
				else
					outbuf.Append(tmpbuf);


				if (atpos)
					outbuf.Append(sepchar);
			}
			outbuf.Append('"');

			separatorNeeded = true;
		}

		if (csvConfig.m_bSubject)
		{
			if (separatorNeeded)
				outbuf.Append(sepchar);

			outbuf.Append('"');

			int subjlen = m->m_subj.GetLength();
			char *subjstr = (char *)(LPCSTR)m->m_subj;

			tmpbuf.ClearAndResize(2 * subjlen);
			int ret_subjlen = escapeSeparators(tmpbuf.Data(), subjstr, subjlen, '"');
			tmpbuf.SetCount(ret_subjlen);

			UINT pageCode = m->m_subj_charsetId;
			if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
			{
				BOOL ret = Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
				if (ret)
					outbuf.Append(*inbuf);
				else
					outbuf.Append(tmpbuf);
			}
			else
				outbuf.Append(tmpbuf);

			outbuf.Append('"');

			separatorNeeded = true;
		}

		if (csvConfig.m_bContent)
		{
			if (separatorNeeded)
				outbuf.Append(sepchar);

			outbuf.Append('"');

			fp.Write(outbuf.Data(), outbuf.Count());

			MboxMail::m_outbuf->Clear();
			UINT pageCode = 0;
			int textlen = GetMailBody_CMimeMessage(mail, i, MboxMail::m_outbuf, pageCode);
			if (textlen != MboxMail::m_outbuf->Count())
				int deb = 1;

			if (MboxMail::m_outbuf->Count())
			{
				int needLength = MboxMail::m_outbuf->Count() * 2 + 1; // worst case scenario or get '"' count first
				inbuf->ClearAndResize(needLength);  // escapeSeparators terminates result with null

				if (MboxMail::m_outbuf->Data()[0] != '\n')
					inbuf->Append("\n");
				int cnt_sv = inbuf->Count();
				int retcnt = escapeSeparators(inbuf->Data(cnt_sv), MboxMail::m_outbuf->Data(), MboxMail::m_outbuf->Count(), '"');
				inbuf->SetCount(retcnt + cnt_sv);

				if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
				{
					BOOL ret = Str2CodePage(inbuf, pageCode, csvConfig.m_nCodePageId, MboxMail::m_outbuf, workbuf);
					if (ret) {
						fp.Write(MboxMail::m_outbuf->Data(), MboxMail::m_outbuf->Count());
					}
					else {
						fp.Write(inbuf->Data(), inbuf->Count());
					}
				}
				else
					fp.Write(inbuf->Data(), inbuf->Count());
			}

			outbuf.Clear();
			outbuf.Append('"');

			separatorNeeded = true;
		}

		outbuf.Append('\n');

		fp.Write(outbuf.Data(), outbuf.Count());

	}
	fp.Close();
	if (csvConfig.m_bContent)
		fpm.Close();

	CString infotxt = _T("Created \"") + csvFile + _T("\" file.");
	HWND h = NULL; // we don't have any window yet
	int answer = ::MessageBox(h, infotxt, _T("Success"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);

	return 1;
}

int MboxMail::GetMailBody_CMimeMessage(CMimeMessage &mail, int mailPosition, SimpleString *outbuf, UINT &pageCode)
{
	int bodyCnt = 0;

	if (mailPosition == 448)
		int deb = 1;

	MboxMail *m = MboxMail::s_mails[mailPosition];

	outbuf->ClearAndResize(10000);
	//SimpleString *inbuf = MboxMail::m_inbuf;
	//inbuf->ClearAndResize(10000);

	// Iterate all the descendant body parts
	CMimeBody::CBodyList bodies;

	int nCount = mail.GetBodyPartList(bodies);
	CMimeBody::CBodyList::const_iterator it;

	CMBodyHdr cmh;

	pageCode = 0;
	bool pageCodeConflict = false;

	if (bodies.begin() == bodies.end()) {
		// Should ASSERT here, there should one at least
		int debug = 1;
	}
	else
	{
		int indx = 0;
		for (it = bodies.begin(); it != bodies.end(); it++)
		{
			CMimeBody* pBP = *it;

			//getCMimeBodyHeader(0, pBP, &cmh);

			if (pBP->IsText() && !pBP->IsAttachment()) 
			{
				if (strcmp(pBP->GetSubType().c_str(), "plain") == 0) {

					char* content = (char*)pBP->GetContent();
					int contentLength = pBP->GetContentLength();
					outbuf->Append(content, contentLength);

					std::string charset = pBP->GetCharset();
					UINT page_code = MboxMail::Str2PageCode(charset.c_str());
					if (pageCodeConflict == false) {
						if (pageCode == 0)
							pageCode = page_code;
						else if (pageCode != page_code)
							pageCodeConflict = true;
					}
				}
			}
		}
		if (pageCodeConflict)
			pageCode = 0;
	}
	return 1;
}

// Get mail text by leveraging the file offset to the entire mail in the index/mboxview file. 
// Full parsing of the mail is performed.
int MboxMail::GetMailBody_MailBody(CFile &fpm, int mailPosition, SimpleString *outbuf, UINT &pageCode)
{
	_int64 fileOffset;
	int bodyCnt = 0;

	MboxMail *m = MboxMail::s_mails[mailPosition];

	outbuf->ClearAndResize(10000);
	SimpleString *inbuf = MboxMail::m_inbuf;
	inbuf->ClearAndResize(10000);

	fileOffset = m->m_startOff;
	fpm.Seek(fileOffset, SEEK_SET);

	int mailLength = m->m_length;
	inbuf->ClearAndResize(mailLength);

	char *mailBegin = inbuf->Data();
	int retlen = fpm.Read(mailBegin, mailLength);
	if (retlen != mailLength) {
		mailLength = retlen;
		int deb = 1;
	}

	char *bodyDataOrigin = mailBegin;

	MailBody *mBody = MailBody::CreateMailBody();
	mBody->Load(bodyDataOrigin, mailBegin, mailLength);

	// Iterate all the descendant body parts
	MailBody::MailBodyList bodies;

	int nCount = mBody->GetBodyPartList(bodies);
	MailBody::MailBodyList::const_iterator it;

	pageCode = 0;
	bool pageCodeConflict = false;

	for (it = bodies.begin(); it != bodies.end(); it++)
	{
		MailBody* pBP = *it;

		if (pBP->m_IsAttachment)
			continue;

		if (!pBP->m_IsTextPlain)
			continue;

		bodyCnt++;
		if (pageCodeConflict == false) {
			if (pageCode == 0)
				pageCode = pBP->m_PageCode;
			else if (pageCode != pBP->m_PageCode)
				pageCodeConflict = true;
		}

		int bodyLength = pBP->m_bodyDataLength;
		if ((pBP->m_bodyDataOffset + pBP->m_bodyDataLength) > m->m_length) {
			// TODO: something is not consistent
			; //bodyLength = m->m_length - body->m_contentOffset;
		}

		char *bodyBegin = mailBegin + pBP->m_bodyDataOffset;

		if (pBP->m_TransferEncoding.CompareNoCase("base64") == 0)
		{
			MboxCMimeCodeBase64 d64(bodyBegin, bodyLength);
			int dlength = d64.GetOutputLength();
			int needLength = dlength + outbuf->Count();
			outbuf->Resize(needLength);

			char *outptr = outbuf->Data(outbuf->Count());
			int retlen = d64.GetOutput((unsigned char*)outptr, dlength);
			if (retlen > 0)
				outbuf->SetCount(outbuf->Count() + retlen);
		}
		else if (pBP->m_TransferEncoding.CompareNoCase("quoted-printable") == 0)
		{
			MboxCMimeCodeQP dGP(bodyBegin, bodyLength);
			int dlength = dGP.GetOutputLength();
			int needLength = dlength + outbuf->Count();
			outbuf->Resize(needLength);

			char *outptr = outbuf->Data(outbuf->Count());
			int retlen = dGP.GetOutput((unsigned char*)outptr, dlength);
			if (retlen > 0)
				outbuf->SetCount(outbuf->Count() + retlen);
		}
		else
		{
			outbuf->Append(bodyBegin, bodyLength);
		}

		if (bodyCnt > 1)
			int deb = 1;
	}

	if (pageCodeConflict)
		pageCode = 0;

	// now it is save to free mBody
	MailBody::FreeMailBody(mBody);

	return outbuf->Count();
}

int MboxMail::LoadMail(const char* pszData, int nDataSize)
{

	return 1;
}

void MailHeader::Clear()
{
	m_IsText = false;
	m_IsTextPlain = false;
	m_IsTextHtml = false;
	m_IsMessage = false;
	m_IsAttachment = false;
	m_IsMultiPart = false;
	m_Charset.Empty();
	m_PageCode = 0;
	m_Description.Empty();
	m_Disposition.Empty();
	m_TransferEncoding.Empty();
	m_SubType.Empty();
	m_MainType.Empty();
	m_Boundary.Empty();
	m_ContentType.Empty();
	m_MediaType = CMimeHeader::MediaType::MEDIA_UNKNOWN;
	m_AttachmentName.Empty();
	m_MessageId.Empty();
	m_ReplyId.Empty();
}

int MailHeader::Load(const char* pszData, int nDataSize)
{
	static const char *cType = "content-type:";
	static const int cTypeLen = strlen(cType);
	static const char *cTransferEncoding = "content-transfer-encoding:";
	static const int cTransferEncodingLen = strlen(cTransferEncoding);
	static const char *cDisposition = "content-disposition:";
	static const int cDispositionLen = strlen(cDisposition);
	static const char *cAttachment = "attachment";
	static const int cAttachmentLen = strlen(cAttachment);
	static const char *cName = "name";
	static const int cNameLen = strlen(cName);
	static const char *cFileName = "filename";
	static const int cFileNameLen = strlen(cFileName);

	static const char *cMsgId = "message-id:";
	static const int cMsgIdLen = strlen(cMsgId);
	static const char *cReplyId = "in-reply-to:";
	static const int cReplyIdLen = strlen(cReplyId);

	static const char *cAlternative = "alternative";
	static const int cAlternativeLen = strlen(cAlternative);

	static const char *cMultipart = "multipart";
	static const int cMultipartLen = strlen(cMultipart);

	static const char *cBoundary = "boundary";
	static const int cBoundaryLen = strlen(cBoundary);

	static const char *cText = "text";
	static const int cTextLen = strlen(cText);

	static const char *cCharset = "charset";
	static const int cCharsetLen = strlen(cCharset);

	int contentIndex = 0;
	int contentLength = 0;

	CString line;
	char *p = (char*)pszData;
	char *e = (char*)pszData + nDataSize;

	while ((p < e) && ((*p == ' ') || (*p == '\t'))) // eat lines starting with ' ' or '\t'
		p = EatNLine(p, e);


	while ((p < e) && ((*p == '\r') || (*p == '\n')))  // eat empty lines
		p++;

	// make sure p is incremented within a loop or break
	while (p < e)
	{
		if (((*p == '\r') && (*(p + 1) == '\n')) || (*p == '\n')) {
			
			p = EatNLine(p, e);
			int headLength = p - pszData;
			return headLength;
			break;  // end of header
		}
		else if (strncmpUpper2Lower(p, e, cType, cTypeLen) == 0) 
		{
			p = GetMultiLine(p, e, line);
			GetFieldValue(line, cTypeLen, m_ContentType);

			if (strncmpUpper2Lower((char*)(LPCSTR)m_ContentType, m_ContentType.GetLength(), cMultipart, cMultipartLen) == 0)
			{
				m_IsMultiPart = true;
				int ret = GetParamValue(line, cTypeLen, cBoundary, cBoundaryLen, m_Boundary);

			}
			else if (strncmpUpper2Lower((char*)(LPCSTR)m_ContentType, m_ContentType.GetLength(), cText, cTextLen) == 0)
			{
				m_IsText = true;
				if (m_ContentType.Compare("text") == 0)
					m_ContentType += "/plain";
				if (m_ContentType.Compare("text/plain") == 0) {
					m_IsTextPlain = true;
					m_IsTextHtml = false;
				}
				else if (m_ContentType.Compare("text/html") == 0) {
					m_IsTextHtml = true;
					m_IsTextPlain = false;
				}
				int ret = GetParamValue(line, cTypeLen, cCharset, cCharsetLen, m_Charset);
				if (!m_Charset.IsEmpty()) {
					m_PageCode = MboxMail::Str2PageCode(m_Charset);
					if (m_PageCode > CP_UTF8)
						int deb = 1;
					if ((m_PageCode == 0) && m_IsTextHtml)
						int deb = 1;
				}
				int deb = 1;
			}
			else {
				int ret = GetParamValue(line, cTypeLen, cName, cNameLen, m_AttachmentName);
				if (!m_AttachmentName.IsEmpty())
					m_IsAttachment = true;
			}
		}
		else if (strncmpUpper2Lower(p, e, cTransferEncoding, cTransferEncodingLen) == 0)
		{
			p = GetMultiLine(p, e, line);
			GetFieldValue(line, cTransferEncodingLen, m_TransferEncoding);
		}
		else if (strncmpUpper2Lower(p, e, cDisposition, cDispositionLen) == 0)
		{
			p = GetMultiLine(p, e, line);
			GetFieldValue(line, cDispositionLen, m_Disposition);

			if (strncmpUpper2Lower((char*)(LPCSTR)m_Disposition, m_Disposition.GetLength(), cAttachment, cAttachmentLen) == 0) {
				m_IsAttachment = true;
				int ret = GetParamValue(line, cTypeLen, cFileName, cFileNameLen, m_AttachmentName);
				if (m_AttachmentName.IsEmpty())
					ret = GetParamValue(line, cTypeLen, cName, cNameLen, m_AttachmentName);
			}
		}
		else if (strncmpUpper2Lower(p, e, cMsgId, cMsgIdLen) == 0) {
			p = GetMultiLine(p, e, line);
			GetMessageId(line, cMsgIdLen, m_MessageId);
		}
		else if (strncmpUpper2Lower(p, e, cReplyId, cReplyIdLen) == 0) {
			p =  GetMultiLine(p, e, line);
			GetMessageId(line, cReplyIdLen, m_ReplyId);
		}
		else {
			p = EatNLine(p, e);
		}
	}
	if ((m_PageCode == 0) && m_IsTextHtml)
		int deb = 1;
	int headLength = p - pszData;
	return headLength;
}

// Followed the design of CMimeBody::Load in mime.cpp, so ite should be reliable :)
// This parser doesn't decode body parts. It only determines length and offsets to the body parts.
int MailBody::Load(char *& pszDataBase, const char* pszData, int nDataSize)
{
	//
	int nSize = MailHeader::Load(pszData, nDataSize);
	if (nSize <= 0)
		return nSize;

	const char* pszDataBegin = pszData;	// preserve start position
	pszData += nSize;
	nDataSize -= nSize;

	// determine the length of the content
	const char* pszEnd = pszData + nDataSize;

	if (IsMultiPart())
	{
		// find the begin boundary
		if (!m_Boundary.IsEmpty())
		{
			// Should match against "\--" or just "--"
			pszEnd = FindBoundary(pszData, pszEnd, m_Boundary, m_Boundary.GetLength());
			if (!pszEnd)
				pszEnd = pszData + nDataSize;
		}
	}

	// load content
	nSize = (int)(pszEnd - pszData);
	if (nSize > 0)
	{
		// save offset and length of body data
		m_bodyDataOffset = pszData - pszDataBase;
		m_bodyDataLength = nSize;
		pszData += nSize;
		nDataSize -= nSize;

		if (m_ContentType.IsEmpty()) {
			m_ContentType = "text/plain";
			m_IsText = true;
			m_IsTextPlain = true;
			m_IsTextHtml = false;
		}
	}
	if (nDataSize <= 0)
		return (int)(pszData - pszDataBegin);

	// load child body parts
	ASSERT(m_Boundary.GetLength() > 0);

	pszEnd = pszData + nDataSize;
	const char* pszBound1 = FindBoundary(pszData, pszEnd, m_Boundary, m_Boundary.GetLength());

	while (pszBound1 != NULL && pszBound1 < pszEnd)
	{
		const char* pszStart = EatNLine((char*)pszBound1, (char*)pszEnd);
		if (!pszStart)
			break;
		if (pszBound1[m_Boundary.GetLength()+2] == '-' && pszBound1[m_Boundary.GetLength() +2 + 1] == '-')
			return (int)(pszStart - pszDataBegin);	// reach the closing boundary

													// look for the next boundary
		const char* pszBound2 = FindBoundary(pszStart, pszEnd, m_Boundary, m_Boundary.GetLength());
		if (!pszBound2)				// overflow, boundary may be truncated
			pszBound2 = pszEnd;
		int nEntitySize = (int)(pszBound2 - pszStart);

		MailBody* pBP = CreatePart();
		m_listBodies.insert_tail(pBP);

		int nInputSize = pBP->Load(pszDataBase, pszStart, nEntitySize);
		if (nInputSize < 0)
		{
			ErasePart(pBP);
			return nInputSize;
		}
		pszBound1 = pszBound2;
	}
	return (int)(pszEnd - pszDataBegin);
}

// return a list of all child body parts belong to this body part
int MailBody::GetBodyPartList(MailBodyList& rList)
{
	int nCount = 0;

	if (!m_IsMultiPart)
	{
		rList.push_back(this);
		nCount++;
	}
	else
	{
		MailBody *iter;
		for (iter = m_listBodies.head(); iter != 0; iter = m_listBodies.next(iter))
		{
			nCount += iter->GetBodyPartList(rList);
		}
	}
	return nCount;
}

char* MailBody::FindBoundary(const char* pszBegin, const char* pszEnd, const char* boundary, int boundaryLength)
{
	static char *bPrefix = "--";
	register char *p = (char*)pszBegin;
	register char *e = (char*)pszEnd;
	while (p < e)
	{
		if (strncmpExact(p, e, bPrefix, 2) == 0)
		{
			if (strncmpExact(&p[2], e, boundary, boundaryLength) == 0)
				return  p;
		}
		p = EatNLine(p, e);
	}
	return 0;
}

MailBody* MailBody::CreatePart()
{
	MailBody* mbody = m_mpool->AllocPart();
	return new MailBody;
}


void MailBody::ErasePart(MailBody* mbody)
{
	mbody->DeleteAll();
	m_mpool->FreePart(mbody);
}

// public static
MailBody* MailBody::CreateMailBody()
{
	MailBody* mbody = m_mpool->AllocPart();
	return new MailBody;
}

// public static
void MailBody::FreeMailBody(MailBody* mbody)
{
	mbody->DeleteAll();
	m_mpool->FreePart(mbody);
}

MailBodyPool::~MailBodyPool()
{
	MailBody *body;
	while ((body = m_freeMailBodyList.remove_head()) != 0)
		delete body;
}

MailBody* MailBodyPool::AllocPart()
{
	MailBody* mbody;
	mbody = m_freeMailBodyList.remove_head();
	if (mbody)
		return mbody;
	else
		return new MailBody;
}


void MailBodyPool::FreePart(MailBody* mbody)
{
	m_freeMailBodyList.insert_head(mbody);
	int deb = 1;
}

// delete all child body parts
void MailBody::DeleteAllParts()
{
	MailBody *body;
	while ((body = m_listBodies.remove_head()) != 0)
	{
		body->DeleteAll();
		m_mpool->FreePart(body);
	}
}

void MailBody::DeleteAll()
{
	DeleteAllParts();
	MailHeader::Clear();
	m_bodyDataOffset = 0;
	m_bodyDataLength = 0;
}

// search for string2 in string1 (strstr)
const char* MailBody::FindString(const char* pszStr1, const char* pszStr2, const char* pszEnd)
{
	pszEnd -= ::strlen(pszStr2);
	const char *s1, *s2;
	while (pszStr1 <= pszEnd)
	{
		s1 = pszStr1;
		s2 = pszStr2;
		while (*s1 == *s2 && *s2)
			s1++, s2++;
		if (!*s2)
			return pszStr1;
		pszStr1++;
	}
	return NULL;
}

void SerializerHelper::close() {
	if (m_buff != NULL) {
		if (m_writing) {
			m_hFile = CreateFile(m_path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
			if (m_hFile != INVALID_HANDLE_VALUE) {
				DWORD nwritten = 0;
				WriteFile(m_hFile, m_buff, m_offset, &nwritten, NULL);
				CloseHandle(m_hFile);
			}
		}
		free(m_buff);
	}
	m_buff = NULL;
}

BOOL SerializerHelper::open(BOOL bWrite) {
	m_writing = bWrite;
	if (bWrite) {
		m_buff = (char *)malloc(m_buffSize = SZBUFFSIZE);
		if (m_buff == NULL)
			return false;
		return TRUE;
		//m_hFile = CreateFile(m_path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	}
	else {
		m_buffSize = (int)FileSize(m_path);
		m_buff = (char *)malloc(m_buffSize);
		if (m_buff == NULL)
			return false;
		m_hFile = CreateFile(m_path, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL);
		DWORD nread;
		BOOL res = ReadFile(m_hFile, m_buff, m_buffSize, &nread, NULL);
		CloseHandle(m_hFile);
		m_hFile = INVALID_HANDLE_VALUE;
		return res;
	}
	return m_hFile != INVALID_HANDLE_VALUE;
}

BOOL SerializerHelper::readN(void *v, int sz) {
	if (m_buff == 0)
		return FALSE;
	if (m_offset + sz > m_buffSize)
		return FALSE;
	memcpy(v, m_buff + m_offset, sz);
	m_offset += sz;
	return TRUE;
}

BOOL SerializerHelper::writeN(void *v, int sz) {
	if (m_buff == 0)
		return FALSE;
	if (m_offset + sz > m_buffSize) {
		m_buff = (char *)realloc(m_buff, m_buffSize += SZBUFFSIZE);
		if (m_buff == NULL)
			return FALSE;
	}
	memcpy(m_buff + m_offset, v, sz);
	m_offset += sz;
	return TRUE;
}

BOOL PathFileExist(LPCSTR path);

int showCodePageTable()
{
	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	if (!path.IsEmpty())
	{
		if (!PathFileExist(path)) {
			return -1;
		}
	}
	else
		return -1;

	CString codePageIdsFile = "WindowsCodePageIds.htm";
	CString fullPath = path + "\\" + codePageIdsFile;

	CFile fp;
	if (!fp.Open(fullPath, CFile::modeWrite | CFile::modeCreate)) {
		CString txt = _T("Could not create \"") + fullPath;
		txt += _T("\" file.\nMake sure file is not open on other applications.");
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CString htmlHdr;
	
	htmlHdr += "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=US-ASCII\"></head><body><pre>";
	htmlHdr += "\r\n\r\n<font size=\"+2\"><b>WINDOWS CODE PAGE IDENTIFIERS TABLE\r\n\r\n</b></font></pre>";
	htmlHdr += "For additional information check the Microsoft document <a href=\"https://docs.microsoft.com/en-us/windows/desktop/intl/code-page-identifiers\">Code Page Identifiers</a>";
	htmlHdr += "<div style=\'text-decoration:underline\'><pre>\r\n";

	fp.Write((LPCSTR)htmlHdr, htmlHdr.GetLength());

	CP2NM *item;

	CString txt;
	txt.Format(" %22s  %26s  %42s  \n",
		" ", " ", " ");
	fp.Write((LPCSTR)txt, txt.GetLength());

	txt.Empty();
	txt.Format("|%22s |%26s |%42s |\n",
		"Code Page Identifiers", "Name", "Additional Infornmation");
	fp.Write((LPCSTR)txt, txt.GetLength());

	int cp2name_size = sizeof(cp2name) / sizeof(CP2NM);
	for (int i = 0; i < cp2name_size; i++)
	{
		item = &cp2name[i];
		txt.Empty();
		txt.Format("|%22d |%26s |%42s |\n", 
			item->m_charsetId, item->m_charset, item->m_info);
		fp.Write((LPCSTR)txt, txt.GetLength());
	}

	CString htmlEnd = "\r\n</pre></body></html>";
	fp.Write((LPCSTR)htmlEnd, htmlEnd.GetLength());

	fp.Close();

	ShellExecute(NULL, _T("open"), fullPath, NULL, NULL, SW_SHOWNORMAL);

	int deb = 1;
	return 1;
}

void ShellExecuteError2Text(UINT errorCode, CString &errorText) {
	errorText.Format("Error code: %u", errorCode);
	if ((errorCode == SE_ERR_FNF) || (errorCode == ERROR_FILE_NOT_FOUND))
		errorText = "File not found";
	else if ((errorCode == SE_ERR_PNF) || (errorCode == ERROR_PATH_NOT_FOUND))
		errorText = "Path not found";
	else if ((errorCode == SE_ERR_NOASSOC))
		errorText = "No associated application";
	else if (errorCode == SE_ERR_SHARE)
		errorText = "Sharing violation";
	else if (errorCode == SE_ERR_ACCESSDENIED)
		errorText = "Access denied";
	else if (errorCode == ERROR_BAD_FORMAT)
		errorText = "Bad executable image";
}