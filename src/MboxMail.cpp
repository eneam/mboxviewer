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
#include "FileUtils.h"
#include "TextUtilsEx.h"
#include "HtmlUtils.h"
#include "mboxview.h"
#include "MboxMail.h"
#include "AttachmentsConfig.h"
#include "MyCTime.h"
#include "ExceptionUtil.h"
#include "MimeParser.h"
#include "SerializationHelper.h"
#include "MimeHelper.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

BOOL MboxMail::m_seExceptionMsgBox = TRUE;
BOOL MboxMail::m_cppExceptionMsgBox = FALSE;

HintConfig MboxMail::m_HintConfig;

MailBodyPool *MailBody::m_mpool = new MailBodyPool;

MessageIdTableType *MboxMail::m_pMessageIdTable = 0;
MboxMailTableType *MboxMail::m_pMboxMailTable = 0;
int MboxMail::m_nextGroupId = 0;

SimpleString* MboxMail::m_outbuf = new SimpleString(10000);
SimpleString* MboxMail::m_inbuf = new SimpleString(10000);
SimpleString* MboxMail::m_outdata = new SimpleString(10000);
SimpleString* MboxMail::m_indata = new SimpleString(10000);
SimpleString* MboxMail::m_workbuf = new SimpleString(10000);
SimpleString* MboxMail::m_tmpbuf = new SimpleString(10000);
SimpleString* MboxMail::m_largebuf = new SimpleString(10000);

int MboxMail::m_Html2TextCount = 0;

_int64 MboxMail::s_curmap = 0;
_int64 MboxMail::s_step = 0;
const CUPDUPDATA* MboxMail::pCUPDUPData = NULL;

bool MboxMail::b_mails_sorted = false;
int MboxMail::b_mails_which_sorted = 0;  // order by file offset

MailArray MboxMail::s_mails_ref;
MailArray MboxMail::s_mails;
MailArray MboxMail::s_mails_all;
MailArray MboxMail::s_mails_find;
MailArray MboxMail::s_mails_edit;
MailArray MboxMail::s_mails_folder;

MailList *MboxMail::m_mailList = &MboxMail::m_allMails;

DLLIST(MailList, m_listLink) MboxMail::m_folderList;

MailList MboxMail::m_allMails(IDC_ARCHIVE_LIST);
MailList MboxMail::m_findMails(IDC_FIND_LIST);
MailList MboxMail::m_editMails(IDC_EDIT_LIST);
MailList MboxMail::m_folderMails(IDC_FOLDER_LIST);

MailArray MboxMail::s_mails_selected;
MailArray MboxMail::s_mails_merged;

_int64 MboxMail::s_fSize = 0;
_int64 MboxMail::s_oSize = 0;
CString MboxMail::s_path;
int MboxMail::nWhichMailList = -1;

int MboxMail::m_EmbededImagesNoMatch = 0;

int MboxMail::m_EmbededImagesFoundMHtml = 0;
int MboxMail::m_EmbededImagesFoundMHtmlHtml = 0;
int MboxMail::m_EmbededImagesFoundUnexpectedMHtml = 0;

int MboxMail::m_EmbededImagesFound = 0;
int MboxMail::m_EmbededImagesFoundCid = 0;
int MboxMail::m_EmbededImagesFoundHttp = 0;
int MboxMail::m_EmbededImagesFoundHttps = 0;
int MboxMail::m_EmbededImagesFoundMHtmlHttp = 0;
int MboxMail::m_EmbededImagesFoundMHtmlHttps = 0;
int MboxMail::m_EmbededImagesFoundData = 0;
int MboxMail::m_EmbededImagesFoundLocalFile = 0;
//
int MboxMail::m_EmbededImagesNotFound = 0;
int MboxMail::m_EmbededImagesNotFoundCid = 0;
int MboxMail::m_EmbededImagesNotFoundHttp = 0;
int MboxMail::m_EmbededImagesNotFoundHttps = 0;
int MboxMail::m_EmbededImagesNotFoundMHtmlHttp = 0;
int MboxMail::m_EmbededImagesNotFoundMHtmlHttps = 0;
int MboxMail::m_EmbededImagesNotFoundData = 0;
int MboxMail::m_EmbededImagesNotFoundLocalFile = 0;

//int fixInlineSrcImgPath(char *inData, int indDataLen, SimpleString *outbuf, CListCtrl *attachments, int mailPosition, bool useMailPosition);
UINT getCodePageFromHtmlBody(SimpleString *buffer, std::string &charset);

///////
// Kept adding and adding Print to functions but now cleanup is needed, better reusability, possible abstractions, error handling, etc
// Postponed to the next relase 1.0.3.3 since larger effort is needed
///////


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

CString MboxMail::GetDateFormat(int i)
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

BOOL MboxMail::GetBody(CString &res)
{
	BOOL ret = TRUE;
	CFile fp;
	if (fp.Open(s_path, CFile::modeRead | CFile::shareDenyWrite)) 
	{
		char *p = res.GetBufferSetLength(m_length);
		//TRACE("offset = %lld\n", m_startOff);
		fp.Seek(m_startOff, SEEK_SET);
		fp.Read(p, m_length);
		char *ms = strchr(p, '\n'); //"-Version: 1.0");
		if (ms) {
			BOOL bAddCR = FALSE;
			if (*(ms - 1) != '\r')
				bAddCR = TRUE;
			int pos = ms - p + 1;
			//res = res.Mid(pos); // - 4);
			if (bAddCR) // for correct mime parsing
			{
				TextUtilsEx::ReplaceNL2CRNL((LPCSTR)res, res.GetLength(), MboxMail::m_tmpbuf);
				res.Empty();
				res.Append(MboxMail::m_tmpbuf->Data(), MboxMail::m_tmpbuf->Count());

				//res.Replace("\n", "\r\n");
			}
			else
				int deb = 1;
		}
		fp.Close();  // TODO: verify why Close() was not called
	}
	else
	{
		DWORD err = GetLastError();
		TRACE("Open Mail File failed err=%ld\n", err);
		ret = FALSE;
	}

	return ret;
};

BOOL MboxMail::GetBody(SimpleString *res)
{
	BOOL ret = TRUE;
	CFile fp;
	res->Clear();
	if (fp.Open(s_path, CFile::modeRead | CFile::shareDenyWrite)) 
	{
		res->Resize(m_length);
		//TRACE("offset = %lld\n", m_startOff);
		fp.Seek(m_startOff, SEEK_SET);
		fp.Read(res->Data(), m_length);
		res->SetCount(m_length);
		fp.Close();
	}
	else
	{
		DWORD err = GetLastError();
		TRACE("Open Mail File failed err=%ld\n", err);
		ret = FALSE;
	}

	return ret;
};

//char *GetMultiLine(char *p, char *e, CString &line);
//int GetFieldValue(CString &fieldLine, int startPos, CString &value);

bool IsFromValidDelimiter(char *p, char *e)
{
	static int cnt = 0;

	return true;

	// EX_TEST
#if 0
	volatile char *badPtr = 0;
	if (cnt == 200)
	{
		*badPtr = 'a';
	}
	cnt++;
#endif

	// make sure it is not user typed "^From" in the message text
	// We expect certain number of digit and colon charcaters if valid "From " mail delimiter
	// There is no inexpensive and sinple method to determine that "^From  points to the begining of the mail
	// Would likely have to attempt to see if the proper mail header follow
	// From 1572544789079124110@xxx Mon Jul 10 14:06:16 +0000 2017
	int digitCount = 0;
	int colonCount = 0;
	int atCount = 0;
	char c;
	while ( (p < e) && ((c = (*p++)) != '\n'))
	{
		if (c == ':')
			colonCount++;
		else if ((c > 47) && (c < 58))
			//else if (isdigit(c))
			digitCount++;
		else  if (c == '@')
			atCount++;
		else if (c == ' ')
		{
			if (TextUtilsEx::strncmpUpper2Lower(p, e, "at ", 3) == 0)
				atCount++;
		}
	}

	if ((colonCount >= 2) && (digitCount >= 6) && (atCount > 0))
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
	static const char *cFromMailBegin = "From ";
	static const int cFromMailBeginLen = strlen(cFromMailBegin);
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
	static const char *cContentType = "content-type:";  // some mails such as from Apple don't seem to have blank line to separate header from body
	static const int cContentTypeLen = strlen(cContentType);

	char *orig = p;
	register char *e = p + bufSize - 1;
	if (p == NULL || bufSize < 10) {
		return false;
	}
	CString contentDisposition = "Content-Disposition: attachment";
	int iFormat = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "format");
	CString format = MboxMail::GetDateFormat(iFormat);
	CString to, from, subject, date;
	CString date_fromField;
	CString cc, bcc;
	time_t tdate = -1;
	time_t tdate_fromField = 1;
	bool	bTo = true, bFrom = true, bSubject = true, bDate = true, bRcvDate = true; // indicates not found, false means found 
	bool bCC = true, bBCC = true;
	char *msgStart = NULL;
	int recv = TRUE;
	int curstep = (int)(startOffset / (s_step?s_step:1));
	CString line;
	CString rcved;

	char *badPtr = 0;

	
#if 1
	if (bFirstView) 
	{
		p = MimeParser::SkipEmptyLines(p, e);
		if (!((TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0) && IsFromValidDelimiter(p, e))) // "From "  marks beginning of the next mail
		{
			msgStart = p;
		}
		bEml = false;  // not used below and not needed :)
	}
#else
	if (bEml && bFirstView) {
		p = MimeParser::SkipEmptyLines(p, e);
		msgStart = p;
		p += 5;
		p = MimeParser::SkipEmptyLines(p, e);
		bEml = false;  // not used below and not needed :)
	}
#endif

	CString progDir;
	BOOL retDir = GetProgramDir(progDir);

	BOOL mustAbort = FALSE;
	TCHAR szCauseBuffer[256];

	char *stackDumpFileName = "Exception_StackTrace.txt";
	char *mailDumpFileName = "Exception_MailDump.txt";
	char *exceptionName = "UnknownException";
	UINT seNumb = 0;
	char *szCause = "Unknown";
	BOOL exceptionThrown = FALSE;
	BOOL exceptionDone = FALSE;

	char *pat = "From: Anu Garg";
	int patLen = strlen(pat);

	BOOL headerDone = FALSE;
	while (p < (e - 4))   // TODO: why (e - 4) ?? chaneg 4 -> 5
	{
		try 
		{
			// EX_TEST
#if 0
			if ((s_mails.GetCount() == 200) && (exceptionDone == FALSE))
			{
				//int i = 0;  intEx = rand() / i;   // This will throw a SE (divide by zero).
				char* szTemp = (char*)1;
				//strcpy_s(szTemp, 1000, "A");
				*badPtr = 'a';
				SCODE sc = 99;
				//AfxThrowOleException(sc);
			}
#endif

			mustAbort = FALSE;

			exceptionThrown = FALSE;
			seNumb = 0;
			szCause = "Unknown";
			stackDumpFileName = "Exception_StackTrace.txt";
			mailDumpFileName = "Exception_MailDump.txt";
			exceptionName = "UnknownException";

#if 0
			if (TextUtilsEx::strncmpExact(p, e, pat, patLen) == 0)
			{
				char *pmsg = p - 10000;
				int deb = 1;
			}
#endif

			// IsFromValidDelimiter() may need to be stronger to avoid false positive detection of the mail beginning
			//if ((*(DWORD*)p == 0x6d6f7246 && p[4] == ' ') && IsFromValidDelimiter(p, e)) // "From "  marks beginning of the next mail
			if ((TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0) && IsFromValidDelimiter(p,e)) // "From "  marks beginning of the next mail
			{
				headerDone = FALSE;
				if (msgStart == NULL)  // keep parsing header until next "From" or the end file
				{
					time_t tdate_fromField = 1;
					date_fromField.Empty();
					SYSTEMTIME tm;

					if (ParseDateInFromField(p, e, &tm))
					{
						MyCTime::fixSystemtime(&tm);
						if (DateParser::validateSystemtime(&tm))
						{
							MyCTime tt(tm);
							date_fromField = tt.FormatGmtTm(format);
							if (date_fromField.IsEmpty())
								date_fromField = tt.FormatLocalTm(format);
							if (date_fromField.IsEmpty())
								int deb = 1;
							tdate_fromField = tt.GetTime();
						}
					}

					msgStart = p;
				}
				else 
				{
					if ((bTo == false) || (bFrom == false) ||  (bSubject == false)) 
					{
						if ((tdate < 0) && (tdate_fromField >= 0))
						{
							date = date_fromField;
							tdate = tdate_fromField;
						}

						time_t tdate_fromField = 1;
						date_fromField.Empty();

						SYSTEMTIME tm;
						if (ParseDateInFromField(p, e, &tm))
						{
							MyCTime::fixSystemtime(&tm);
							if (DateParser::validateSystemtime(&tm))
							{
								MyCTime tt(tm);
								date_fromField = tt.FormatGmtTm(format);
								if (date_fromField.IsEmpty())
									date_fromField = tt.FormatLocalTm(format);
								if (date_fromField.IsEmpty())
									int deb = 1;
								tdate_fromField = tt.GetTime();
							}
						}

						if (!from.IsEmpty() || !to.IsEmpty() ||  !subject.IsEmpty() ||  !date.IsEmpty())
						{
							MboxMail *m = new MboxMail();
							m->m_startOff = startOffset + (_int64)(msgStart - orig);
							m->m_length = p - msgStart;
							m->m_to = TextUtilsEx::DecodeString(to, m->m_to_charset, m->m_to_charsetId);
							m->m_from = TextUtilsEx::DecodeString(from, m->m_from_charset, m->m_from_charsetId);
							m->m_subj = TextUtilsEx::DecodeString(subject, m->m_subj_charset, m->m_subj_charsetId);
							m->m_cc = TextUtilsEx::DecodeString(cc, m->m_cc_charset, m->m_cc_charsetId);
							m->m_bcc = TextUtilsEx::DecodeString(bcc, m->m_bcc_charset, m->m_bcc_charsetId);
							if (!bcc.IsEmpty())
								int deb = 1;
							if (!cc.IsEmpty())
								int deb = 1;
							m->m_subj.Replace("\r", "");
							m->m_subj.Replace("\n", "");
							m->m_timeDate = tdate;
							m->m_recv = recv;
							m->m_headLength = 0;  // set in ParseContent()

							//m->m_crc32 = TextUtilities::CalcCRC32(msgStart, m->m_length);

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

							if (s_mails.GetCount() == 1098)
								int deb = 1;

							if ((s_mails.GetCount() % 100) == 0)
							{
								CString mailNum;
								mailNum.Format(_T("Parsing archive file to create index file ... %d"), s_mails.GetCount());

								_int64 curOff = startOffset + (p - orig);
								int step = (int)(curOff / s_step);

								if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(mailNum, (UINT_PTR)(step));
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
					tdate = -1;
					bTo = bFrom = bSubject = bDate = bRcvDate = bCC = bBCC = true;
				}
				p += 5;
				p = MimeParser::EatNewLine(p, e);
			}
			else if (msgStart && !headerDone)
			{
				if (bFrom && TextUtilsEx::strncmpUpper2Lower(p, e, cFrom, cFromLen) == 0)
				{
					bFrom = false;
					p = MimeParser::GetMultiLine(p, e, line);
					from = line.Mid(cFromLen);
					from.Trim();
				}
				else if (bTo && TextUtilsEx::strncmpUpper2Lower(p, e, cTo, cToLen) == 0)
				{
					bTo = false;
					p = MimeParser::GetMultiLine(p, e, line);
					to = line.Mid(cToLen);
					to.Trim();
				}
				else if (bCC && TextUtilsEx::strncmpUpper2Lower(p, e, cCC, cCCLen) == 0)
				{
					bCC = false;
					p = MimeParser::GetMultiLine(p, e, line);
					cc = line.Mid(cCCLen);
					cc.Trim();
				}
				else if (bBCC && TextUtilsEx::strncmpUpper2Lower(p, e, cBCC, cBCCLen) == 0)
				{
					bBCC = false;
					p = MimeParser::GetMultiLine(p, e, line);
					bcc = line.Mid(cBCCLen);
					bcc.Trim();
				}
				else if (bSubject && TextUtilsEx::strncmpUpper2Lower(p, e, cSubject, cSubjectLen) == 0)
				{
					bSubject = false;
					p = MimeParser::GetMultiLine(p, e, line);
					subject = line.Mid(cSubjectLen);
					subject.Trim();
				}
				else if (bDate && bRcvDate && TextUtilsEx::strncmpUpper2Lower(p, e, cReceived, cReceivedLen) == 0)
				{
					//bDate = false; 
					p = MimeParser::GetMultiLine(p, e, line);
					rcved = line.Mid(cReceivedLen);
					rcved.Trim();
					// find last ';' separator
					int pos = rcved.ReverseFind(';');
					if (pos >= 0)
						date = rcved.Mid(pos + 1);
					else
						date = rcved;
					SYSTEMTIME tm;
					if (DateParser::parseRFC822Date(date, &tm)) 
					{
						MyCTime::fixSystemtime(&tm);
						if (DateParser::validateSystemtime(&tm)) 
						{
							MyCTime tt(tm);
							date = tt.FormatGmtTm(format);
							if (date.IsEmpty())
								date = tt.FormatLocalTm(format);
							if (date.IsEmpty())
								int deb = 1;
							tdate = tt.GetTime();
							bRcvDate = false;
							recv = TRUE;
						}
					}
					else
						int deb = 1;
				}
				else if (bDate && TextUtilsEx::strncmpUpper2Lower(p, e, cDate, cDateLen) == 0)
				{
					p = MimeParser::GetMultiLine(p, e, line);
					date = line.Mid(cDateLen);
					date.Trim();
					SYSTEMTIME tm;
					if (DateParser::parseRFC822Date(date, &tm)) 
					{
						MyCTime::fixSystemtime(&tm);
						if (DateParser::validateSystemtime(&tm)) 
						{
							MyCTime tt(tm);
							date = tt.FormatGmtTm(format);
							if (date.IsEmpty())
								date = tt.FormatLocalTm(format);
							if (date.IsEmpty())
								int deb = 1;
							tdate = tt.GetTime();
							bDate = false;
							recv = FALSE;
						}
					}
					else
						int deb = 1;
				}
#if 0
				// doesn't quite work; looking for empty lines seem to work best
				else if (TextUtilsEx::strncmpUpper2Lower(p, e, cContentType, cContentTypeLen) == 0)
				{
					// 
					p = MimeParser::GetMultiLine(p, e, line);
					CString contTypeVal = line.Mid(cContentTypeLen);
					contTypeVal.Trim();
					// find first ';' separator
					int pos = contTypeVal.Find(';');
					CString contType;
					if (pos >= 0)
						contType = contTypeVal.Mid(0, pos);
					else
						contType = contTypeVal;
					contType.Trim();

					CString contentTypeMain = contType.Left(4);  // we are interested in "text" type
					if (contentTypeMain.CompareNoCase("text") == 0)
						headerDone = TRUE;   // skip remaining lines until next "From" or the end file 
				}
#endif
				else
				{
					BOOL isEmpty = FALSE;
					p = MimeParser::EatNewLine(p, e, isEmpty);
					// TODO:  This check may not completely reliable and may need better check and likley more expensive check
					// Without the end of header check, we could pickup CC and BCC fiekds from mail text
					// but it could be worst if we stop looking for headear fields prematuraly
					if (isEmpty)
					{
						if (from.IsEmpty() && to.IsEmpty() && date.IsEmpty())
							int deb = 1;

						headerDone = TRUE;   // skip remaining lines until next "From" or the end file 
					}
				}
			}
			else
			{
				p = MimeParser::EatNewLine(p, e);
			}
		}

		//   Other supported exceptions. At the moment we are file using 
		//catch (CMemoryException * e)
		//catch (CFileException * e)
		//catch (COleException * e)
		//catch (CUserException* e)
		//catch (CResourceException * e)
		//catch CSimpleException
		//catch CInvalidArgException
		//catch CNotSupportedException
		//catch CArchiveException
		//catch CDBException
		//catch COleDispatchException
		//catch CDaoException
		//catch CInternetException

		catch (const SE_Exception& ex)
		{
			// TODO: this is very long catch  !!!
			//exceptionThrown = TRUE;
			exceptionDone = TRUE;
			seNumb = ex.getSeNumber();
			szCause = (char*)seDescription(seNumb);

			int mailPosition = s_mails.GetCount();
			char *se_stackDumpFileName = "SystemE_StackTrace.txt";
			BOOL ret = DumpStack(se_stackDumpFileName, szCause, seNumb, 0, mailPosition);

			int datalen = 0;
			char *data = p;
			char *msgEnd = p;
			char *p_save = p;

			while (p < e)
			{
				//if ((TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0) && IsFromValidDelimiter(p, e))
				if (TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0) {
					msgEnd = p;
					break;
				}
				p = MimeParser::EatNewLine(p, e);
				msgEnd = p;
			}
			// or update above -:)
			p = MimeParser::EatNewLine(p, e);
			msgEnd = p;
			while (p < e)
			{
				//if ((TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0) && IsFromValidDelimiter(p, e))
				if (TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0) {
					msgEnd = p;
					break;
				}
				p = MimeParser::EatNewLine(p, e);
				msgEnd = p;
			}

			if (msgStart)
			{
				data = msgStart;
				datalen = msgEnd - msgStart;
			}
			else
			{
				data = orig;
				datalen = msgEnd - orig;
			}
			if (datalen < 0)
				datalen = 0;

			char *se_mailDumpFileName = "SystemE_MailDump.txt";
			BOOL retW = DumpMailData(se_mailDumpFileName, szCause, seNumb, mailPosition, data, datalen);

			char *se_translation_stackDumpFileName = "TranslationSystemE_StackTrace.txt";

			CString errorTxt;
#ifdef USE_STACK_WALKER
			errorTxt.Format(_T("SE_Exception: Code=%8.8x Description=%s\n\n"
				"To help to diagnose the problem, created files\n\n%s , %s and %s\n\nin\n\n%s directory.\n\n"
				"Please provide the files to the development team.\n\n"
				"%s file contains mail data so please review the content before forwarding to the development."),
				seNumb, szCause, se_stackDumpFileName, se_mailDumpFileName, se_translation_stackDumpFileName, progDir, se_mailDumpFileName);
#else
			errorTxt.Format(_T("SE_Exception: Code=%8.8x Description=%s\n\n"
				"To help to diagnose the problem, created file\n\n%s\n\nin\n\n%s directory.\n\n"
				"Please provide the files to the development team.\n\n"
				"%s file contains mail data so please review the content before forwarding to the development."
				"Please run mboxview residing in DEBUG directory to generate additional information."),
				seNumb, szCause, se_mailDumpFileName, progDir, se_mailDumpFileName);
#endif

			if (m_seExceptionMsgBox)
			{
				// Documents say it is not safe to call AfxMessageBox from separate thread but it seem to kind of work
				// May need to return info to the main thread and create MessageBox
				AfxMessageBox((LPCTSTR)errorTxt, MB_OK | MB_ICONHAND);
				AfxAbort();
			}
		}
		catch (CException* e)
		{
			exceptionName = "Exception";

			seNumb = 0;
			e->GetErrorMessage(szCauseBuffer, 255, &seNumb);
			szCause = &szCauseBuffer[0];

			stackDumpFileName = "E_StackTrace.txt";
			mailDumpFileName = "E_MailDump.txt";

			e->Delete();
			exceptionThrown = TRUE;
			exceptionDone = TRUE;
		}
		catch (...)  // memory leak ??
		{
			exceptionName = "AnyException";
			seNumb = 0;
			szCause = "Unknown";

			stackDumpFileName = "AnyE_StackTrace.txt";
			mailDumpFileName = "AnyE_MailDump.txt";

			exceptionThrown = TRUE;
			mustAbort = TRUE;
			exceptionDone = TRUE;
		}

		if (exceptionThrown)
		{
			// TODO: this is very long catch  !!!
			int mailPosition = s_mails.GetCount();
			BOOL ret = DumpStack(stackDumpFileName, szCause, seNumb, 0, mailPosition);

			int datalen = 0;
			char *data = p;
			char *msgEnd = p;
			char *p_save = p;

			while (p < e)
			{
				//if ((TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0) && IsFromValidDelimiter(p, e))
				if (TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0) {
					msgEnd = p;
					break;
				}
				p = MimeParser::EatNewLine(p, e);
				msgEnd = p;
			}
			// or update above -:)
			p = MimeParser::EatNewLine(p, e);
			msgEnd = p;
			while (p < e)
			{
				//if ((TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0) && IsFromValidDelimiter(p, e))
				if (TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0) {
					msgEnd = p;
					break;
				}
				p = MimeParser::EatNewLine(p, e);
				msgEnd = p;
			}

			if (msgStart)
			{
				data = msgStart;
				datalen = msgEnd - msgStart;
			}
			else
			{
				data = orig;
				datalen = msgEnd - orig;
			}
			if (datalen < 0)
				datalen = 0;

			BOOL retW = DumpMailData(mailDumpFileName, szCause, seNumb, mailPosition, data, datalen);

			CString errorTxt;
#ifdef USE_STACK_WALKER
			errorTxt.Format(_T("%s: Code=%8.8x Description=%s\n\n"
				"To help to diagnose the problem, created files\n\n%s and %s\n\nin\n\n%s directory.\n\n"
				"Please provide the files to the development team.\n\n"
				"%s file contains mail data so please review the content before forwarding to the development."),
				exceptionName, seNumb, szCause, stackDumpFileName, mailDumpFileName, progDir, mailDumpFileName);
#else
			errorTxt.Format(_T("%s: Code=%8.8x Description=%s\n\n"
				"To help to diagnose the problem, created file\n\n%s\n\nin\n\n%s directory.\n\n"
				"Please provide the files to the development team.\n\n"
				"%s file contains mail data so please review the content before forwarding to the development.\n\n"
				"Please run mboxview residing in DEBUG directory to generate additional information."),
				exceptionName, seNumb, szCause, mailDumpFileName, progDir, mailDumpFileName);
#endif

			if (m_cppExceptionMsgBox || mustAbort)
			{
				// Documents say it is not safe to call AfxMessageBox from separate thread but it seem to kind of work
				// May need to return info to the main thread and create MessageBox
				AfxMessageBox((LPCTSTR)errorTxt, MB_OK | MB_ICONHAND);
				AfxAbort();
			}
			// To attempt to continue, reset the below
			p = p_save;
			p = MimeParser::EatNewLine(p, e);

			msgStart = NULL; tdate = -1; headerDone = FALSE;
			to.Empty(); from.Empty(); subject.Empty(); date.Empty(); cc.Empty(); bcc.Empty();
			bTo = bFrom = bSubject = bDate = bRcvDate = bCC = bBCC = true;
		}
	}
	
	// should save state for cross boundary emails - resolved 2018
	if (msgStart != NULL && msgStart != p)
	{
		if (!bLastView) {
			lastStartOffset = startOffset + (_int64)(msgStart - orig);
		}
		else if ((bTo == false) || (bFrom == false) || (bSubject == false))
		{
			mustAbort = FALSE;

			exceptionThrown = FALSE;
			seNumb = 0;
			szCause = "Unknown";
			stackDumpFileName = "Exception_StackTrace.txt";
			mailDumpFileName = "Exception_MailDump.txt";
			exceptionName = "UnknownException";

			//  TODO: Handling of exception - lots of duplications FIXME
			try
			{
				// EX_TEST
#if 0
				//if (s_mails.GetCount() == 200)
				{
					//int i = 0;  intEx = rand() / i;   // This will throw a SE (divide by zero).
					char* szTemp = (char*)1;
					//strcpy_s(szTemp, 1000, "A");
					//*badPtr = 'a';
					SCODE sc = 99;
					//AfxThrowOleException(sc);
				}
#endif
				if (!from.IsEmpty() || !to.IsEmpty() || !subject.IsEmpty() || !date.IsEmpty())
				{
					MboxMail *m = new MboxMail();
					//		TRACE("start: %d length: %d\n", msgStart - orig, p - msgStart);
					m->m_startOff = startOffset + (_int64)(msgStart - orig);
					m->m_length = p - msgStart;
					m->m_to = TextUtilsEx::DecodeString(to, m->m_to_charset, m->m_to_charsetId);
					m->m_from = TextUtilsEx::DecodeString(from, m->m_from_charset, m->m_from_charsetId);
					m->m_subj = TextUtilsEx::DecodeString(subject, m->m_subj_charset, m->m_subj_charsetId);
					m->m_cc = TextUtilsEx::DecodeString(cc, m->m_cc_charset, m->m_cc_charsetId);
					m->m_bcc = TextUtilsEx::DecodeString(bcc, m->m_bcc_charset, m->m_bcc_charsetId);
					if (!bcc.IsEmpty())
						int deb = 1;
					if (!cc.IsEmpty())
						int deb = 1;
					m->m_subj.Replace("\r", "");
					m->m_subj.Replace("\n", "");
					m->m_recv = recv;
					m->m_timeDate = tdate;
					m->m_headLength = 0; // set in ParseContent()

					//m->m_crc32 = TextUtilities::CalcCRC32(msgStart, m->m_length);

					MailBodyContent body;
					char *bodyStart = msgStart;
					char *bodyEnd = msgStart + m->m_length;
					char *nextStart = ParseContent(m, bodyStart, bodyEnd);

					s_mails.Add(m);

					int index = s_mails.GetCount() - 1;
					//m->DumpMailBox(m, index);
					//NListView::DumpItemDetails(index);

					if (s_mails.GetCount() == 1098)
						int deb = 1;
				}
			}
			catch (const SE_Exception& ex)
			{
				//exceptionThrown = TRUE;
				exceptionDone = TRUE;
				seNumb = ex.getSeNumber();
				szCause = (char*)seDescription(seNumb);

				int mailPosition = s_mails.GetCount();
				char *se_stackDumpFileName = "SystemE_StackTrace.txt";
				BOOL ret = DumpStack(stackDumpFileName, szCause, seNumb, 0, mailPosition);

				
				int datalen = p - msgStart;
				char *msgEnd = p + datalen;

				char *se_mailDumpFileName = "SystemE_MailDump.txt";
				BOOL retW = DumpMailData(mailDumpFileName, szCause, seNumb, mailPosition, msgStart, datalen);

				char *se_translation_stackDumpFileName = "TranslationSystemE_StackTrace.txt";

				CString errorTxt;
				errorTxt.Format(_T("SE_Exception: Code=%8.8x Description=%s\n\n"
					"To help to diagnose the problem, files:\n\n%s and %s and %s were created in:\n\n%s directory.\n\n"
					"Please provide the files to the development team.\n\n"
					"%s file contains mail data so please review the content before forwarding to the development."),
					seNumb, szCause, se_stackDumpFileName, se_mailDumpFileName, se_translation_stackDumpFileName, progDir, mailDumpFileName);

				if (m_seExceptionMsgBox)
				{
					AfxMessageBox((LPCTSTR)errorTxt, MB_OK | MB_ICONHAND);
					AfxAbort();
				}

				// To attempt to continue, reset the below
				msgStart = NULL; tdate = -1; headerDone = FALSE;
				to.Empty(); from.Empty(); subject.Empty(); date.Empty(); cc.Empty(); bcc.Empty();
				bTo = bFrom = bSubject = bDate = bRcvDate = bCC = bBCC = true;
			}
			catch (CException* e)
			{
				exceptionThrown = TRUE;
				exceptionDone = TRUE;

				exceptionName = "Exception";

				seNumb = 0;
				e->GetErrorMessage(szCauseBuffer, 255, &seNumb);
				szCause = &szCauseBuffer[0];

				stackDumpFileName = "E_StackTrace.txt";
				mailDumpFileName = "E_MailDump.txt";

				e->Delete();
			}
			catch (...)  // memory leak ??
			{
				exceptionThrown = TRUE;
				exceptionDone = TRUE;

				exceptionName = "AnyException";
				seNumb = 0;
				szCause = "Unknown";

				stackDumpFileName = "AnyE_StackTrace.txt";
				mailDumpFileName = "AnyE_MailDump.txt";

				mustAbort = TRUE;
			}

			if (exceptionThrown)
			{
				int mailPosition = s_mails.GetCount();
				BOOL ret = DumpStack(stackDumpFileName, szCause, seNumb, 0, mailPosition);

				int datalen = p - msgStart;
				char *msgEnd = p + datalen;
				char *p_save = p;

				BOOL retW = DumpMailData(mailDumpFileName, szCause, seNumb, mailPosition, msgStart, datalen);

				CString errorTxt;
				errorTxt.Format(_T("%s: Code=%8.8x Description=%s\n\n"
					"To help to diagnose the problem, files:\n\n%s and %s were created in:\n\n%s directory.\n\n"
					"Please provide the files to the development team.\n\n"
					"%s file contains mail data so please review the content before forwarding to the development."),
					exceptionName, seNumb, szCause, stackDumpFileName, mailDumpFileName, progDir, mailDumpFileName);

				if (m_cppExceptionMsgBox || mustAbort)
				{
					AfxMessageBox((LPCTSTR)errorTxt, MB_OK | MB_ICONHAND);
					AfxAbort();
				}
				// To attempt to continue, reset the below
				p = p_save;
				p = MimeParser::EatNewLine(p, e);
				msgStart = NULL; tdate = -1; headerDone = FALSE;
				to.Empty(); from.Empty(); subject.Empty(); date.Empty(); cc.Empty(); bcc.Empty();
				bTo = bFrom = bSubject = bDate = bRcvDate = bCC = bBCC = true;
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

	// Delete all files in print and image cache directories
	CString cpath = path;
	CString rootPrintSubFolder = "PrintCache";;
	CString targetPrintSubFolder;
	CString prtCachePath;
	CString errorText;
#if 0
	// Removing Printcache not critical I think
	bool ret1 = MboxMail::GetArchiveSpecificCachePath(cpath, rootPrintSubFolder, targetPrintSubFolder, prtCachePath, errorText);
	if (errorText.IsEmpty() && FileUtils::PathDirExists(prtCachePath)) {
		pCUPDUPData->SetProgress(_T("Deleting all files in the PrintCache directory ..."), 0);
		FileUtils::RemoveDir(prtCachePath, true);
	}
#endif

	rootPrintSubFolder = "ImageCache";
	errorText.Empty();
	bool ret2 = MboxMail::GetArchiveSpecificCachePath(cpath, rootPrintSubFolder, targetPrintSubFolder, prtCachePath, errorText);
	if (errorText.IsEmpty() && FileUtils::PathDirExists(prtCachePath)) {
		pCUPDUPData->SetProgress(_T("Deleting all related files in the ImageCache directory ..."), 0);
		FileUtils::RemoveDir(prtCachePath, true);
	}

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
	_int64 mappingSize = (_int64)MAPPING_SIZE;

	int mappingsInFile = (int)((fSize - 1)/mappingSize) + 1;
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

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	CString mailFilePath = path;
	FileUtils::SplitFilePath(mailFilePath, driveName, directory, fileNameBase, fileNameExtention);

	CString mailFile = fileNameBase + fileNameExtention;
	CString parsingFileText = _T("Parsing \"") + mailFile + _T("\" ...");

	//pCUPDUPData->SetProgress(parsingFileText, 0);  // works but doesn't always fit into progress bar
	pCUPDUPData->SetProgress(_T("Parsing archive file to create index file ..."), 0);
	// TODO: due to breaking the file into multiple chunks, it looks some emails can be lost : Fixed
	bool firstView = true;;
	bool lastView = false;
	DWORD bufSize = 0;
	_int64 aligned_offset = 0;
	_int64 delta = 0;


#if 0
	//set_terminate(term_func);
	char *dllList[] = { "msvcrt.dll" , "msvcr80.dll", "msvcr90.dll", "msvcr100.dll" };
	HMODULE hDLL = 0;
	for (int ii = 0; ii < sizeof(dllList); ii++)
	{
		hDLL = LoadLibrary(dllList[ii]);
		if (hDLL != NULL)
			break;
	}
	if (hDLL)
		Scoped_SE_Translator scoped_se_translator(trans_func);
#endif

#ifdef USE_STACK_WALKER
	Scoped_SE_Translator scoped_se_translator(trans_func);
#endif

	// TODO:  Need to consider to redo reading and create  Character Stream for reading
	_int64 lastStartOffset = 0;
	while  ((lastView == false) && !pCUPDUPData->ShouldTerminate()) 
	{
		s_curmap = lastStartOffset;
		aligned_offset = (s_curmap / dwAllocationGranularity) * dwAllocationGranularity;
		delta = s_curmap - aligned_offset;
		bufSize = ((fSize - aligned_offset) < mappingSize) ? (DWORD)(fSize - aligned_offset) : (DWORD)mappingSize;

		TRACE("offset=%lld, bufsize=%ld, fSize-curmap=%lld, end=%lld\n", s_curmap, bufSize, fSize - s_curmap, s_curmap + bufSize);
		char * pview = (char *)MapViewOfFileEx(hFileMap, FILE_MAP_READ, (DWORD)(aligned_offset >> 32), (DWORD)aligned_offset, bufSize, NULL);

		if (pview == 0)
		{
			DWORD err = GetLastError();
			CString txt = _T("Could not finish parsing due to memory fragmentaion. Please restart the mbox viewer to resolve.");
			HWND h = NULL; // we don't have any window yet ??
			int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			int deb = 1;
			break;
		}

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
	MboxMail::s_mails_ref.SetSizeKeepData(MboxMail::s_mails.GetSize());
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
	else
		return (cmp < 0);
}
bool sortByFromDesc(MboxMail *cr1, MboxMail *cr2) {
	int cmp = cr1->m_from.Compare(cr2->m_from);
	if (cmp == 0)
		return (cr1->m_timeDate > cr2->m_timeDate); // make stable sort
	else
		return (cmp > 0);
}

bool sortByTo(MboxMail *cr1, MboxMail *cr2) {
	int cmp = cr1->m_to.Compare(cr2->m_to);
	if (cmp == 0)
		return (cr1->m_timeDate < cr2->m_timeDate); // make stable sort
	else
		return (cmp < 0);
}
bool sortByToDesc(MboxMail *cr1, MboxMail *cr2) {
	int cmp = cr1->m_to.Compare(cr2->m_to);
	if (cmp == 0)
		return (cr1->m_timeDate < cr2->m_timeDate); // make stable sort
	else
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

bool sortByGroupId(MboxMail *cr1, MboxMail *cr2) {
	if (cr1->m_groupId == cr2->m_groupId)
		return (cr1->m_timeDate < cr2->m_timeDate); // make stable sort
	else
		return (cr1->m_groupId < cr2->m_groupId);
}
bool sortByGroupIdDesc(MboxMail *cr1, MboxMail *cr2) {
	if (cr1->m_groupId == cr2->m_groupId)
		return (cr1->m_timeDate < cr2->m_timeDate); // make stable sort
	else
		return (cr1->m_groupId > cr2->m_groupId);
}

bool sortBySize(MboxMail *cr1, MboxMail *cr2) {
	if (cr1->m_length == cr2->m_length)
		return (cr1->m_timeDate < cr2->m_timeDate); // make stable sort
	else
		return (cr1->m_length < cr2->m_length);
}
bool sortBySizeDesc(MboxMail *cr1, MboxMail *cr2) {
	if (cr1->m_length == cr2->m_length)
		return (cr1->m_timeDate < cr2->m_timeDate); // make stable sort
	else
		return (cr1->m_length > cr2->m_length);
}

bool sortByFileOffset(MboxMail *cr1, MboxMail *cr2) {
	return (cr1->m_startOff < cr2->m_startOff);
}
bool sortByFileOffsetDesc(MboxMail *cr1, MboxMail *cr2) {
	return (cr1->m_startOff > cr2->m_startOff);
}

bool sortByArrayIndex(MboxMail *cr1, MboxMail *cr2) {
	return (cr1->m_index < cr2->m_index);
}
bool sortByArrayIndexDesc(MboxMail *cr1, MboxMail *cr2) {
	return (cr1->m_index > cr2->m_index);
}

void MboxMail::SortByDate(MailArray *s_m, bool bDesc)
{
	if (s_m == 0) s_m = &MboxMail::s_mails;
	std::stable_sort(s_m->GetData(), s_m->GetData() + s_m->GetSize(), bDesc ? sortByDateDesc : sortByDate);
}
void MboxMail::SortByFrom(MailArray *s_m, bool bDesc)
{
	if (s_m == 0) s_m = &MboxMail::s_mails;
	std::sort(s_m->GetData(), s_m->GetData() + s_m->GetSize(), bDesc ? sortByFromDesc : sortByFrom);
}
void MboxMail::SortByTo(MailArray *s_m, bool bDesc)
{
	if (s_m == 0) s_m = &MboxMail::s_mails;
	std::sort(s_m->GetData(), s_m->GetData() + s_m->GetSize(), bDesc ? sortByToDesc : sortByTo);
}
void MboxMail::SortBySubject(MailArray *s_m, bool bDesc)
{
	if (s_m == 0) s_m = &MboxMail::s_mails;
	std::sort(s_m->GetData(), s_m->GetData() + s_m->GetSize(), bDesc ? sortBySubjectDesc : sortBySubject);
}
void MboxMail::SortBySize(MailArray *s_m, bool bDesc)
{
	if (s_m == 0) s_m = &MboxMail::s_mails;
	std::sort(s_m->GetData(), s_m->GetData() + s_m->GetSize(), bDesc ? sortBySizeDesc : sortBySize);
}
void MboxMail::SortByGroupId(MailArray *s_m, bool bDesc)
{
	if (s_m == 0) s_m = &MboxMail::s_mails;
	std::sort(s_m->GetData(), s_m->GetData() + s_m->GetSize(), bDesc ? sortByGroupId : sortByGroupIdDesc);
}

void MboxMail::SortByFileOffset(MailArray *s_m, bool bDesc)
{
	if (s_m == 0) s_m = &MboxMail::s_mails;
	std::sort(s_m->GetData(), s_m->GetData() + s_m->GetSize(), bDesc ? sortByFileOffsetDesc : sortByFileOffset);
}

void MboxMail::SortByConverstionGroups(MailArray *s_m, bool bDesc)
{
	// Assumption here is that s_mails_ref array is already ordered by groups
	int arraySize = s_mails_ref.GetSize();
	MboxMail::s_mails.SetSizeKeepData(arraySize);
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

void MboxMail::SortByIndex(MailArray *s_m, bool bDesc)
{
	if (s_m == 0) s_m = &MboxMail::s_mails;
	std::sort(s_m->GetData(), s_m->GetData() + s_m->GetSize(), bDesc ? sortByArrayIndexDesc : sortByArrayIndex);
}

// Function to help investigate implementation issues
int MboxMail::DumpMailBox(MboxMail *mailBox, int which)
{
	//return 1;
	static BOOL restarted = TRUE;
	int deb; // set breakpoint
	char *CRLF = "\n\n";
	char buff[16382 +2]; // TODO: allocate on heap
	int buffsize = 16382;
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
		pos = FileUtils::FileSeek(hFile, fsize, FILE_BEGIN);
		if (pos < 0) {
			DWORD err = GetLastError();
			return -1;
		}
	}

	MboxMail *m = mailBox;

	CString format;
	datebuff[0] = 0;
	if (m->m_timeDate >= 0) {
		MyCTime tt(m->m_timeDate);
		CString lDateTime = tt.FormatGmtTm(format);
		strcpy(datebuff, (LPCSTR)lDateTime);
	}

	count = sprintf_s(buff, "INDX=%d first=%lld len=%d last=%lld att=%d hlen=%d rcv=%d date=\"%s\" from=\"%s\" to=\"%s\" subj=\"%s\"\n",
		which, m->m_startOff, m->m_length, (m->m_startOff + m->m_length - 1), m->m_hasAttachments, m->m_headLength, m->m_recv,
		datebuff, (LPCSTR)m->m_from, (LPCSTR)m->m_to, (LPCSTR)m->m_subj);

	nwritten = 0;
	if (!FileUtils::Write2File(hFile, buff, count, &nwritten)) {
		DWORD retval = GetLastError();
	}

	long long start_offset = m->m_startOff;
	pos = FileUtils::FileSeek(mbox_hFile, start_offset, FILE_BEGIN);

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
		if (!FileUtils::Write2File(hFile, buff, lpNumberOfBytesRead, &nwritten)) {
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
		if (!FileUtils::Write2File(hFile, txt.GetBuffer(), txt.GetLength(), &nwritten)) {
			DWORD retval = GetLastError();
		}

		start_offset = m->m_startOff + body->m_contentOffset;
		pos = FileUtils::FileSeek(mbox_hFile, start_offset, FILE_BEGIN);

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
			if (!FileUtils::Write2File(hFile, outbuf->Data(), outbuf->Count(), &nwritten)) {
				DWORD retval = GetLastError();
			}

			if (!FileUtils::Write2File(hFile, CRLF, strlen(CRLF), &nwritten)) {
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
			if (!FileUtils::Write2File(hFile, outbuf->Data(), outbuf->Count(), &nwritten)) {
				DWORD retval = GetLastError();
			}

			if (!FileUtils::Write2File(hFile, CRLF, strlen(CRLF), &nwritten)) {
				DWORD retval = GetLastError();
			}
			deb = 1;
		}
		else
		{
			nwritten = 0;
			if (!FileUtils::Write2File(hFile, inbuf->Data(), inbuf->Count(), &nwritten)) {
				DWORD retval = GetLastError();
				break;
			}
		}
	}
	nwritten = 0;
	if (!FileUtils::Write2File(hFile, CRLF, strlen(CRLF), &nwritten)) {
		DWORD retval = GetLastError();
	}

	CloseHandle(hFile);
	CloseHandle(mbox_hFile);

	return 1;
}

/* ParseContent Parser parses mail file to determine the file offset and length of content blocks.
   MIME is defined as a nested set of data containers similar to the tree structure. 
   Boundary tags are used to indicate begining and end of the data block. Structure can be recursive.
   Example:

   optional-header fields such as List-Unsubscribe: seem to appear in the optional header also, see Expedia

Header
| From: email
| To : email
| MIME - Version : 1.0
| Content - Type : multipart/mixed; boundary = "boundaryMix1";
Message body
|    emails from Expedia insert optional-header fields such as List-Unsubscribe: after multipart/mixed; doesn't seem to be correct ZM
|    also, emails from Expedia insert optional-header fields such as List-Unsubscribe: as part of the header; that seem to be correct ZM
| multipart/mixed --boundaryMix1
| --boundaryMix1
| multipart/related --boundaryRel1
|   |--boundaryReleated1
|   |   multipart/alternative --boundaryAlt1
|   |   |--boundaryAlt1
|   |   |text / plain
|   |   |--boundaryAlt1
|   |   |text / html
|   |   |--boundaryAlt1--
|   |--boundaryRel1
|   |Inline image
|   |--boundaryRel1
|   |Inline image
|   |--boundaryRel1--
| --boundaryMix1
| Attachment1
| --boundaryMix1
| Attachment2
| --boundaryMix1
| Attachment3
| --boundaryMix1--
|
.
*/

/*
You can set by selecting "Optional Header" as the search Target and set the following conditions as the Optional Header :

Optional Header Fields:
: X-TO-ADDR …Extracts a unique email address included in TO
: X-CC-ADDR …Extracts a unique email address included in CC
: X-TO-CC-ADDR …Extracts a unique email address included in TO and CC
: X-TO-DOMAIN …Extracts a unique domain included in TO
: X-CC-DOMAIN …Extracts a unique domain included in CC
: X-TO-CC-DOMAIN …Extracts a unique domain included in TO and CC
*/

/*
http://www.ietf.org/rfc/rfc2369.txt
Optional Header Fields:
3.2.List - Unsubscribe

The List - Unsubscribe field describes the command(preferably using
	mail) to directly unsubscribe the user(removing them from the list).

	Examples:

List - Unsubscribe : <mailto : list@host.com ? subject = unsubscribe>
List - Unsubscribe : (Use this command to get off the list)
< mailto : list - manager@host.com ? body = unsubscribe % 20list >
List - Unsubscribe : <mailto : list - off@host.com>



Neufeld & Baer              Standards Track[Page 4]

RFC 2369                  URLs as Meta - Syntax                  July 1998


List - Unsubscribe: < http ://www.host.com/list.cgi?cmd=unsub&lst=list>,
	<mailto : list - request@host.com ? subject = unsubscribe>

	3.3.List - Subscribe

	The List - Subscribe field describes the command(preferably using
		mail) to directly subscribe the user(request addition to the list).

	Examples:

List - Subscribe : <mailto : list@host.com ? subject = subscribe>
List - Subscribe : <mailto : list - request@host.com ? subject = subscribe>
List - Subscribe : (Use this command to join the list)
< mailto : list - manager@host.com ? body = subscribe % 20list >
List - Subscribe : <mailto : list - on@host.com>
List - Subscribe : < http ://www.host.com/list.cgi?cmd=sub&lst=list>,
	<mailto : list - manager@host.com ? body = subscribe % 20list>

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
			contentDetails->m_contentId = pBP->m_ContentId;
			contentDetails->m_contentTransferEncoding = pBP->m_TransferEncoding;

			if (!pBP->m_Name.IsEmpty())
			{
				contentDetails->m_attachmentName = pBP->m_Name;
				contentDetails->m_attachmentNamePageCode = pBP->m_NamePageCode;
			}
			else if (!pBP->m_AttachmentName.IsEmpty())
			{
				contentDetails->m_attachmentName = pBP->m_AttachmentName;
				contentDetails->m_attachmentNamePageCode = pBP->m_AttachmentNamePageCode;
			}

			contentDetails->m_contentLocation = pBP->m_ContentLocation;
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

	clearMboxMailTable();
	clearMessageIdTable();

	for (int i = 0; i < s_mails_ref.GetSize(); i++)
	{
		for (int j = 0; j < s_mails_ref[i]->m_ContentDetailsArray.size(); j++) {
			delete s_mails_ref[i]->m_ContentDetailsArray[j];
			s_mails_ref[i]->m_ContentDetailsArray[j] = 0;
		}
		delete s_mails_ref[i];
		s_mails_ref[i] = 0;
	}

	s_mails.SetSizeKeepData(0);
	s_mails_ref.SetSizeKeepData(0);
	s_mails_find.SetSizeKeepData(0);
	s_mails_edit.SetSizeKeepData(0);
	s_mails_selected.SetSizeKeepData(0);
	s_mails_merged.SetSizeKeepData(0);
	s_mails_folder.SetSizeKeepData(0);

	b_mails_sorted = false;
	MboxMail::s_path = "";

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
		pFrame->EnableAllMailLists(FALSE);

	MboxMail::nWhichMailList = -1;
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
	MessageIdTableType::iterator it;
	if (m_pMessageIdTable == 0)
		createMessageIdTable(50000);

	it = m_pMessageIdTable->find(key);
	if (it != m_pMessageIdTable->cend()) {
		return it->second;
	}
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
		createMessageIdTable(50000);

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

void MboxMail::clearMessageIdTable()
{
	MessageIdTableType::iterator it;

	if (m_pMessageIdTable) {
		for (it = m_pMessageIdTable->begin(); it != m_pMessageIdTable->end(); it++) {
			delete it->first;
			//it->first = 0;
		}
		m_pMessageIdTable->clear();
	}
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

		if (m->m_duplicateId) // TODO: what can be done :)
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
	int i_ref = s_mails.GetSize() - 1;
	int i_ref_begin, i_ref_end;

	s_mails_ref.SetSizeKeepData(s_mails.GetSize());

	if (!validateSortConversations()) {
		//s_mails is already left sorted by date
		return false;
	}

	int currentGroupColor = 0; // TODO; white , not used currently, remove ?
	m_nextGroupId = 0;

							   // init s_mails_ref by conversation groups
	for (int i = s_mails.GetSize() - 1; i >= 0; i--)
	{
		m = s_mails[i];
		if (m->m_done)
			continue;

		i_ref_end = i_ref;
		m->m_groupId = m_nextGroupId;
		s_mails_ref[i_ref--] = m;
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
		while (m->m_prevMail >= 0)
		{
			m = s_mails[m->m_prevMail];
			if (m->m_done)
				continue;  // should never be here ??
			m->m_groupId = m_nextGroupId;
			m->m_groupColor = currentGroupColor;
			s_mails_ref[i_ref--] = m;
			m->m_done = true;
		}
		m = m_sav;
		while (m->m_nextMail >= 0)
		{
			m = s_mails[m->m_nextMail];
			if (m->m_done)
				continue; // should never be here ??
			m->m_groupId = m_nextGroupId;
			m->m_groupColor = currentGroupColor;
			s_mails_ref[i_ref--] = m;
			m->m_done = true;
		}
		i_ref_begin = i_ref;
		m_nextGroupId++;

		int j;

		// check sorting order
		bool outOfOrder = false;

		for (j = (i_ref_begin+1); j < i_ref_end; j++)
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
			for (j = (i_ref_begin+1); j < i_ref_end; j++)
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
			MailArray s_group_mails;

			for (j = (i_ref_begin+1); j <= i_ref_end; j++)
			{
				s_group_mails.Add(s_mails_ref[j]);
			}

			MboxMail::SortByDate(&s_group_mails);

			j = i_ref_begin;
			for (int k = 0; i < s_group_mails.GetSize(); i++)
			{
				time_t mtime = s_group_mails[k]->m_timeDate;

				s_mails_ref[j++] = s_group_mails[k];
			}

			// Recheck mail order
			for (j = (i_ref_begin+1); j < i_ref_end; j++)
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
#if 1
	MboxMail::SortByFileOffset(&MboxMail::s_mails_ref);
	for (int i = 0; i < s_mails_ref.GetSize(); i++)
	{
		m = s_mails_ref[i];
		m->m_index = i;
	}
	s_mails.Copy(s_mails_ref);
	return true;

#else
	// make content of both s_mails_ref and s_mails the same 
	if (i_ref != -1) {  // TODO: implementation error
		s_mails_ref.Copy(s_mails);
		for (int i = 0; i < s_mails.GetSize(); i++)
		{
			m = s_mails[i];
			m->m_index = i;
		}
		return false;
	}
	else {
		s_mails.Copy(s_mails_ref);
		for (int i = 0; i < s_mails.GetSize(); i++)
		{
			m = s_mails[i];
			m->m_index = i;
		}
		return true;
	}
#endif
}

// verify that after all mails are accounted after sorting by group
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


// Not used currently/anymore delete ?? !!!
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


// This works on the entire list or subset
void MboxMail::assignColor2ConvesationGroups(MailArray *mails)
{
	MboxMail *m;
	MboxMail *m_prev;
	int i_ref = 0;

	int groupColor = 0;

	if (mails->GetSize() <= 0)
		return;

	m = mails->operator[](0);
	m->m_groupColor = groupColor;

	int i;
	for (i = 1; i < mails->GetSize(); i++)
	{
		m = mails->operator[](i);
		m_prev = mails->operator[](i-1);
		if (m->m_groupId != m_prev->m_groupId) 
		{
			if (groupColor == 0)
				groupColor = 1;
			else
				groupColor = 0;

		}
		m->m_groupColor = groupColor;
	}
}

bool ALongRightProcessProcPrintMailArchiveToCSV(const CUPDUPDATA* pCUPDUPData)
{
	PRINT_MAIL_ARCHIVE_TO_CSV_ARGS *args = (PRINT_MAIL_ARCHIVE_TO_CSV_ARGS*)pCUPDUPData->GetAppData();
	MboxMail::pCUPDUPData = pCUPDUPData;

	HANDLE h = GetCurrentThread();
	BOOL prio = SetThreadPriority(h, THREAD_PRIORITY_ABOVE_NORMAL);
	DWORD myThreadId = GetCurrentThreadId();
	DWORD myThreadPri = GetThreadPriority(h);
	TRACE(_T("(ALongRightProcessProcPrintMailArchiveToCSV) threadId=%ld threadPriority=%ld\n"), myThreadId, myThreadPri);

	BOOL progressBar = TRUE;
	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as param to printMailArchiveToTextFile
	CString errorText;
	int retpos = MboxMail::printMailArchiveToCSVFile(args->csvConfig, args->csvFile, args->firstMail, args->lastMail, args->selectedMailIndexList, progressBar, errorText);
	args->errorText = errorText;
	args->exitted = TRUE;
	return true;
}

int MboxMail::exportHeaderFieldLabelsToCSVFile(CSVFILE_CONFIG &csvConfig, CFile &fp)
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
		colLabels += "From Name";
		colLabels += csvConfig.m_separator;
		colLabels += "From Address";
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

	if (csvConfig.m_bCC) {
		if (separatorNeeded)
			colLabels += csvConfig.m_separator;
		colLabels += "CC";
		separatorNeeded = true;
	}

	if (csvConfig.m_bBCC) {
		if (separatorNeeded)
			colLabels += csvConfig.m_separator;
		colLabels += "BCC";
		separatorNeeded = true;
	}

	if (csvConfig.m_bAttachmentNames) {
		if (separatorNeeded)
			colLabels += csvConfig.m_separator;
		colLabels += "AttachmentNames";
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

int MboxMail::DetermineLimitedLength(SimpleString *str, int maxLinesTextLimit)
{
	char *p = str->Data();
	char *e = p + str->Count();

	while ((p < e) && (maxLinesTextLimit > 0)) {
		p = MimeParser::EatNewLine(p, e);
		maxLinesTextLimit--;
	}
	int limitedLength = p - str->Data();
	if (limitedLength > str->Count())
		limitedLength = str->Count();

	return limitedLength;
}

int MboxMail::printMailHeaderToCSVFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, CSVFILE_CONFIG &csvConfig)
{
	MboxMail *m;
	CString colLabels;
	CString separator;
	bool separatorNeeded;
#define BUFFLEN 10000
	char *buff = 0;
	char *token = 0;
	char *seppos;
	char *atpos;
	char *last;
	int tokenlen = 0;

	char sepchar = csvConfig.m_separator.GetAt(0);
	char *sepstr = (char*)(LPCSTR)csvConfig.m_separator;

	SimpleString outbuf(1024, 256);
	SimpleString tmpbuf(1024, 256);
	SimpleString name(256);
	SimpleString addr(256);

	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;
	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);

	if (mailPosition == 1749)
		int deb = 1;

	m = s_mails[mailPosition];

	separatorNeeded = false;

	outbuf.Clear();

	if (csvConfig.m_bDate)
	{
		char datebuff[32];
		CString format = MboxMail::GetDateFormat(csvConfig.m_dateFormat);

		datebuff[0] = 0;
		if (m->m_timeDate >= 0)
		{
			MyCTime tt(m->m_timeDate);
			if (!csvConfig.m_bGMTTime) 
			{
				CString lDateTime = tt.FormatLocalTm(format);
				strcpy(datebuff, (LPCSTR)lDateTime);
			}
			else {
				CString lDateTime = tt.FormatGmtTm(format);
				strcpy(datebuff, (LPCSTR)lDateTime);
			}
		}

		if (separatorNeeded)
			outbuf.Append(sepchar);

		outbuf.Append('"');

		int begCount = outbuf.Count();

		int datelen = strlen(datebuff);
		if (charCount(datebuff, '"') > 0)
		{
			tmpbuf.ClearAndResize(2 * datelen);
			int ret_datelen = escapeSeparators(tmpbuf.Data(), datebuff, datelen, '"');
			outbuf.Append(tmpbuf.Data(), ret_datelen);
		}
		else
			outbuf.Append(datebuff, datelen);

		char *data = outbuf.Data(begCount);
		int dataLength = outbuf.Count() - begCount;
		int retLength = EnforceFieldTextCharacterLimit(data, dataLength, csvConfig.m_MessageLimitCharsString);
		outbuf.SetCount(begCount + retLength);

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

		int begCount = outbuf.Count();

		// disable to support investigation of group Ids
#if 1
		int namelen = name.Count();

		tmpbuf.ClearAndResize(2 * namelen);
		int ret_namelen = escapeSeparators(tmpbuf.Data(), name.Data(), namelen, '"');
		tmpbuf.SetCount(ret_namelen);

		UINT pageCode = m->m_from_charsetId;
		if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
		{
			BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
			if (ret)
				outbuf.Append(*inbuf);
			else
				outbuf.Append(tmpbuf);
		}
		else {
			outbuf.Append(tmpbuf);
		}
#else
		UINT pageCode = m->m_from_charsetId;
		// print GroupId instead of name for testing
		CString strGroupId;
		strGroupId.Format("%d", m->m_groupId);
		outbuf.Append((LPCSTR)strGroupId, strGroupId.GetLength());
#endif
		char *data = outbuf.Data(begCount);
		int dataLength = outbuf.Count() - begCount;
		int retLength = EnforceFieldTextCharacterLimit(data, dataLength, csvConfig.m_MessageLimitCharsString);
		outbuf.SetCount(begCount + retLength);

		outbuf.Append('"');
		outbuf.Append(sepchar);

		outbuf.Append('"');

		begCount = outbuf.Count();

		int addrlen = addr.Count();

		tmpbuf.ClearAndResize(2 * addrlen);
		int ret_addrlen = escapeSeparators(tmpbuf.Data(), addr.Data(), addrlen, '"');
		tmpbuf.SetCount(ret_addrlen);

		pageCode = m->m_from_charsetId;
		if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
		{
			BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
			if (ret)
				outbuf.Append(*inbuf);
			else
				outbuf.Append(tmpbuf);
		}
		else
			outbuf.Append(tmpbuf);

		data = outbuf.Data(begCount);
		dataLength = outbuf.Count() - begCount;
		retLength = EnforceFieldTextCharacterLimit(data, dataLength, csvConfig.m_MessageLimitCharsString);
		outbuf.SetCount(begCount + retLength);

		outbuf.Append('"');

		separatorNeeded = true;
	}

	if (csvConfig.m_bTo)
	{
		int tolen = m->m_to.GetLength();
		if ((tolen*1.5) > MboxMail::m_largebuf->Capacity()) 
		{
			MboxMail::m_largebuf->ClearAndResize(2 * tolen);
		}
		buff = MboxMail::m_largebuf->Data();
		memcpy(buff, m->m_to, tolen);
		buff[tolen] = 0;

		last = &buff[tolen];

		if (separatorNeeded)
			outbuf.Append(sepchar);

		outbuf.Append('"');

		int begCount = outbuf.Count();

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
				BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
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

		char *data = outbuf.Data(begCount);
		int dataLength = outbuf.Count() - begCount;
		int retLength = EnforceFieldTextCharacterLimit(data, dataLength, csvConfig.m_MessageLimitCharsString);
		outbuf.SetCount(begCount + retLength);

		outbuf.Append('"');

		separatorNeeded = true;
	}

	if (csvConfig.m_bSubject)
	{
		if (separatorNeeded)
			outbuf.Append(sepchar);

		outbuf.Append('"');

		int begCount = outbuf.Count();

		int subjlen = m->m_subj.GetLength();
		char *subjstr = (char *)(LPCSTR)m->m_subj;

		tmpbuf.ClearAndResize(2 * subjlen);
		int ret_subjlen = escapeSeparators(tmpbuf.Data(), subjstr, subjlen, '"');
		tmpbuf.SetCount(ret_subjlen);

		UINT pageCode = m->m_subj_charsetId;
		if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
		{
			BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
			if (ret)
				outbuf.Append(*inbuf);
			else
				outbuf.Append(tmpbuf);
		}
		else
			outbuf.Append(tmpbuf);

		char *data = outbuf.Data(begCount);
		int dataLength = outbuf.Count() - begCount;
		int retLength = EnforceFieldTextCharacterLimit(data, dataLength, csvConfig.m_MessageLimitCharsString);
		outbuf.SetCount(begCount + retLength);

		outbuf.Append('"');

		separatorNeeded = true;
	}

	if (csvConfig.m_bCC)
	{
		int tolen = m->m_cc.GetLength();
		if ((tolen*1.5) > MboxMail::m_largebuf->Capacity())
		{
			MboxMail::m_largebuf->ClearAndResize(2 * tolen);
		}
		buff = MboxMail::m_largebuf->Data();
		memcpy(buff, m->m_cc, tolen);
		buff[tolen] = 0;

		last = &buff[tolen];

		if (separatorNeeded)
			outbuf.Append(sepchar);

		outbuf.Append('"');

		int begCount = outbuf.Count();

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
				BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
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

		char *data = outbuf.Data(begCount);
		int dataLength = outbuf.Count() - begCount;
		int retLength = EnforceFieldTextCharacterLimit(data, dataLength, csvConfig.m_MessageLimitCharsString);
		outbuf.SetCount(begCount + retLength);

		outbuf.Append('"');

		separatorNeeded = true;
	}

	if (csvConfig.m_bBCC)
	{
		int tolen = m->m_bcc.GetLength();
		if ((tolen*1.5) > MboxMail::m_largebuf->Capacity())
		{
			MboxMail::m_largebuf->ClearAndResize(2 * tolen);
		}
		buff = MboxMail::m_largebuf->Data();
		memcpy(buff, m->m_bcc, tolen);
		buff[tolen] = 0;

		last = &buff[tolen];

		if (separatorNeeded)
			outbuf.Append(sepchar);

		outbuf.Append('"');

		int begCount = outbuf.Count();

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
				BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
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

		char *data = outbuf.Data(begCount);
		int dataLength = outbuf.Count() - begCount;
		int retLength = EnforceFieldTextCharacterLimit(data, dataLength, csvConfig.m_MessageLimitCharsString);
		outbuf.SetCount(begCount + retLength);

		outbuf.Append('"');

		separatorNeeded = true;
	}

	if (csvConfig.m_bAttachmentNames)
	{
		if (separatorNeeded)
			outbuf.Append(sepchar);

		NListView::PrintAttachmentNames(m, &outbuf, csvConfig.m_MessageLimitCharsString);

		separatorNeeded = true;
	}

	if (csvConfig.m_bContent)
	{
		if (separatorNeeded)
			outbuf.Append(sepchar);
	}

	fp.Write(outbuf.Data(), outbuf.Count());

	return 1;
}

// export mails to CSV file by leveraging index/mboxview data.
// Header fields from the index/mboview files are used.
int MboxMail::exportToCSVFile(CSVFILE_CONFIG &csvConfig, CString &csvFileName, int firstMail, int lastMail, MailIndexList *selectedMailsIndexList, BOOL progressBar)
{
	CString mailFile = MboxMail::s_path;

	if (s_path.IsEmpty()) {
		CString txt = _T("Please open mail file first.");
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	if (path.IsEmpty()) {
		CString txt = _T("No path to archive file folder.");
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;  // Hopefully s_path wil fail first
	}

	//
	CFile fp;
	int ret = 1;
	BOOL selectedMailIndexList = FALSE;
	bool fileExists;
	CString errorText;
	CString targetPrintSubFolder;
	CString csvFile;
	int textType = 3;  // CSV
	if (!progressBar) // no worker thread  TODO: currently never TRUE
	{
		CString targetPrintSubFolder;
		if (selectedMailIndexList)
		{
			ret = MakeFileNameFromMailArchiveName(textType, csvFile, targetPrintSubFolder, fileExists, errorText);
		}
		else if (firstMail == lastMail)
		{
			ret = MakeFileNameFromMailHeader(firstMail, textType, csvFile, targetPrintSubFolder, fileExists, errorText);
		}
		else
		{
			ret = MakeFileNameFromMailArchiveName(textType, csvFile, targetPrintSubFolder, fileExists, errorText);
		}
	}
	else  // worker thread
	{
		CString targetPrintSubFolder;
		ret = MakeFileNameFromMailArchiveName(textType, csvFile, targetPrintSubFolder, fileExists, errorText);
	}

	if (ret < 0) {
		HWND h = NULL;
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}
	csvFileName = csvFile;

	if (!progressBar)
	{
		if (!fp.Open(csvFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone)) {
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
				CString txt = _T("Could not open mail archive \"") + s_path;
				HWND h = NULL; // we don't have any window yet ??
				int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
				fp.Close();
				return -1;
			}
		}

		int mailPosition;
		bool first = true;
		if (selectedMailIndexList == 0)
		{
			for (int i = firstMail; i < lastMail; i++)
			{
				mailPosition = i;
				first = (i == firstMail);
				printSingleMailToCSVFile(fp, mailPosition, fpm, csvConfig, first);
			}
		}
		else
		{
			int i;
			firstMail = (*selectedMailsIndexList)[0];
			int cnt = selectedMailsIndexList->GetCount();
			for (int j = 0; j < cnt; j++)
			{
				//i = selectedMailIndexList->ElementAt(j);
				i = (*selectedMailsIndexList)[j];
				mailPosition = j;
				printSingleMailToCSVFile(fp, mailPosition, fpm, csvConfig, first);
				first = false;
			}
		}

		fp.Close();
		if (csvConfig.m_bContent)
			fpm.Close();
	}
	else
	{
		PRINT_MAIL_ARCHIVE_TO_CSV_ARGS args;
		args.csvConfig = csvConfig;
		args.csvFile = csvFile;
		args.firstMail = firstMail;
		args.lastMail = lastMail;
		args.selectedMailIndexList = selectedMailsIndexList;
		args.terminated = FALSE; // flag set by MboxMail::printMailArchiveToTextFile to ack termination request
		args.exitted = FALSE; // flag set by ALongRightProcessProcPrintMailArchive  upon return

		MboxMail::m_Html2TextCount = 0;

		CUPDialog	Dlg(AfxGetMainWnd()->GetSafeHwnd(), ALongRightProcessProcPrintMailArchiveToCSV, (LPVOID)(PRINT_MAIL_ARCHIVE_TO_CSV_ARGS*)&args);

		INT_PTR nResult = Dlg.DoModal();

		if (!nResult) // should never be true ?
		{
			MboxMail::assert_unexpected();
			return -1;
		}

		TRACE(_T("m_Html2TextCount=%d MailCount=%d\n"), m_Html2TextCount, MboxMail::s_mails.GetCount());

		int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
		int retResult = LOWORD(nResult);

		if (retResult != IDOK)
		{  // IDOK==1, IDCANCEL==2
			// We should be here when user selects Cancel button
			ASSERT(cancelledbyUser == TRUE);
			int loopCnt = 20;
			DWORD tc_start = GetTickCount();
			while ((loopCnt-- > 0) && (args.exitted == FALSE))
			{
				Sleep(25);
			}
			DWORD tc_end = GetTickCount();
			DWORD delta = tc_end - tc_start;
			TRACE("(exportToCSVFile)Waited %ld milliseconds for thread to exist.\n", delta);
		}

		MboxMail::pCUPDUPData = NULL;

		if (!args.errorText.IsEmpty()) 
		{
			HWND h = NULL; 
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return - 1;
		}

	}
	return 1;
}

int MboxMail::printSingleMailToCSVFile(/*out*/ CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, CSVFILE_CONFIG &csvConfig, bool firstMail)
{
	if (firstMail)
		exportHeaderFieldLabelsToCSVFile(csvConfig, fp);

	printMailHeaderToCSVFile(fp, mailPosition, fpm, csvConfig);

	char sepchar = csvConfig.m_separator.GetAt(0);
	char *sepstr = (char*)(LPCSTR)csvConfig.m_separator;

	SimpleString outbuf(1024, 256);
	SimpleString tmpbuf(1024, 256);
	SimpleString name(256);
	SimpleString addr(256);

	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;
	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);

	outbuf.Clear();

	if (csvConfig.m_bContent)
	{
		outbuf.Append('"');

		fp.Write(outbuf.Data(), outbuf.Count());

		MboxMail::m_outbuf->Clear();
		UINT pageCode = 0;
		int textlen = GetMailBody_mboxview(fpm, mailPosition, MboxMail::m_outbuf, pageCode);  // fast
		if (textlen != MboxMail::m_outbuf->Count())
			int deb = 1;

		if (MboxMail::m_outbuf->Count())
		{
			int lineLimit = -1;
			if (TextUtilsEx::isNumeric(csvConfig.m_MessageLimitString))
				lineLimit = _ttoi(csvConfig.m_MessageLimitString);

			if (lineLimit >= 0) {
				HtmlUtils::MergeWhiteLines(MboxMail::m_outbuf, lineLimit);
			}

			if (MboxMail::m_outbuf->Count())
			{
				int needLength = MboxMail::m_outbuf->Count() * 2 + 10; // worst case scenario or get '"' count first
				inbuf->ClearAndResize(needLength);  // escapeSeparators terminates result with null

				BOOL insertEmptyLine = FALSE;   // make this configurable one day -:)
				if ((MboxMail::m_outbuf->Data()[0] == '\n') || ((MboxMail::m_outbuf->Data()[0] == '\r') && (MboxMail::m_outbuf->Data()[1] == '\n')))
					; // do noting
				else if (insertEmptyLine)
					inbuf->Append("\r\n");  // insert empty line at the beginning

				int cnt_sv = inbuf->Count();
				int retcnt = escapeSeparators(inbuf->Data(cnt_sv), MboxMail::m_outbuf->Data(), MboxMail::m_outbuf->Count(), '"');
				inbuf->SetCount(retcnt + cnt_sv);

				if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
				{
					BOOL ret = TextUtilsEx::Str2CodePage(inbuf, pageCode, csvConfig.m_nCodePageId, MboxMail::m_outbuf, workbuf);
					if (ret) {
						EnforceCharacterLimit(MboxMail::m_outbuf, csvConfig.m_MessageLimitCharsString);
						fp.Write(MboxMail::m_outbuf->Data(), MboxMail::m_outbuf->Count());
					}
					else {
						EnforceCharacterLimit(inbuf, csvConfig.m_MessageLimitCharsString);
						fp.Write(inbuf->Data(), inbuf->Count());
					}
				}
				else {
					EnforceCharacterLimit(inbuf, csvConfig.m_MessageLimitCharsString);
					fp.Write(inbuf->Data(), inbuf->Count());
				}
			}
		}
		else
		{
			SimpleString *inbuf = MboxMail::m_inbuf;
			SimpleString *outbuf = MboxMail::m_outbuf;
			SimpleString *workbuf = MboxMail::m_workbuf;


			inbuf->ClearAndResize(10000);
			outbuf->ClearAndResize(10000);
			workbuf->ClearAndResize(10000);

			pageCode = 0;
			int textType = 1; // HTML

			int textlen = GetMailBody_mboxview(fpm, mailPosition, /*out*/outbuf, pageCode, textType);  // returns pageCode
			if (textlen != outbuf->Count())
				int deb = 1;

			if (outbuf->Count())
			{
				MboxMail::m_Html2TextCount++;

				CString bdycharset;
				std::string charSet;
				if (pageCode > 0)
					BOOL ret = TextUtilsEx::id2charset(pageCode, charSet);
				else
				{
					pageCode = getCodePageFromHtmlBody(outbuf, charSet);

					if (pageCode == 0) {
						pageCode = CP_UTF8;
						BOOL ret = TextUtilsEx::id2charset(pageCode, charSet);
					}
				}

				bdycharset.Append(charSet.c_str());

				inbuf->Clear();

				CString bdy;
				bool extraHtmlHdr = false;
				if (outbuf->FindNoCase(0, "<body", 5) < 0) // didn't find if true
				{
					extraHtmlHdr = true;
					bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body>";
					inbuf->Append(bdy, bdy.GetLength());
				}

				inbuf->Append(*outbuf);
				if (extraHtmlHdr)
					inbuf->Append("</body></html>\r\n");

				outbuf->Clear();

				UINT outPageCode = CP_UTF8;
				HtmlUtils::GetTextFromIHTMLDocument(/*in*/inbuf, /*out*/outbuf, pageCode, outPageCode);

				inbuf->Clear();
				int needLength = outbuf->Count() * 2 + 10; // worst case scenario or get '"' count first
				inbuf->ClearAndResize(needLength);  // escapeSeparators terminates result with null

				if ((outbuf->Data()[0] == '\n') || ((outbuf->Data()[0] == '\r') && (outbuf->Data()[1] == '\n')))
					; // do nothing
				else
					inbuf->Append("\r\n");

				int cnt_sv = inbuf->Count();
				int retcnt = escapeSeparators(/*out*/inbuf->Data(cnt_sv), /*in*/outbuf->Data(), /*in*/outbuf->Count(), '"');
				inbuf->SetCount(retcnt + cnt_sv);

				int lineLimit = -1;
				if (TextUtilsEx::isNumeric(csvConfig.m_MessageLimitString))
					lineLimit = _ttoi(csvConfig.m_MessageLimitString);

				if (lineLimit > 0) {
					HtmlUtils::MergeWhiteLines(inbuf, lineLimit);
				}

				EnforceCharacterLimit(inbuf, csvConfig.m_MessageLimitCharsString);
				fp.Write(inbuf->Data(), inbuf->Count());
			}
		}

		outbuf.Clear();
		outbuf.Append('"');
	}

	outbuf.Append("\r\n");

	fp.Write(outbuf.Data(), outbuf.Count());

	return 1;
}


int MboxMail::printMailHeaderToTextFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig)
{
	static char *border = "--------------------------------------------------------------------------------\r\n";

	SimpleString tmpbuf(1024, 256);

	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;
	SimpleString *outbuf = MboxMail::m_tmpbuf;

	inbuf->Clear();
	workbuf->Clear();
	outbuf->Clear();

	MboxMail *m = s_mails[mailPosition];

	outbuf->Append(border);

	/////
	int subjlen = m->m_subj.GetLength();
	char *subjstr = (char *)(LPCSTR)m->m_subj;

	char *cSubject = "SUBJECT: ";
	int cSubjectLen = strlen(cSubject);
	outbuf->Append(cSubject, cSubjectLen);

	UINT pageCode = m->m_subj_charsetId;
	if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
	{
		tmpbuf.Copy(subjstr, subjlen);
		BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
		if (ret)
			outbuf->Append(*inbuf);
		else
			outbuf->Append(tmpbuf);
	}
	else {
		outbuf->Append(subjstr, subjlen);
	}
	outbuf->Append("\r\n");
	//////
	int fromlen = m->m_from.GetLength();
	char *fromstr = (char *)(LPCSTR)m->m_from;

	char *cFrom = "FROM: ";
	int cFromLen = strlen(cFrom);
	outbuf->Append(cFrom, cFromLen);

	pageCode = m->m_from_charsetId;
	if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
	{
		tmpbuf.Copy(fromstr, fromlen);
		BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
		if (ret)
			outbuf->Append(*inbuf);
		else
			outbuf->Append(tmpbuf);
	}
	else {
		outbuf->Append(fromstr, fromlen);
	}
	outbuf->Append("\r\n");
	//////
	int tolen = m->m_to.GetLength();
	char *tostr = (char *)(LPCSTR)m->m_to;

	char *cTo = "TO: ";
	int cToLen = strlen(cTo);
	outbuf->Append(cTo, cToLen);

	pageCode = m->m_to_charsetId;
	if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
	{
		tmpbuf.Copy(tostr, tolen);
		BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
		if (ret)
			outbuf->Append(*inbuf);
		else
			outbuf->Append(tmpbuf);
	}
	else {
		outbuf->Append(tostr, tolen);
	}
	outbuf->Append("\r\n");
	////
	if (!m->m_cc.IsEmpty())
	{
		int cclen = m->m_cc.GetLength();
		char *ccstr = (char *)(LPCSTR)m->m_cc;

		char *cCC = "CC: ";
		int cCCLen = strlen(cCC);
		outbuf->Append(cCC, cCCLen);

		pageCode = m->m_cc_charsetId;
		if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
		{
			tmpbuf.Copy(ccstr, cclen);
			BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
			if (ret)
				outbuf->Append(*inbuf);
			else
				outbuf->Append(tmpbuf);
		}
		else {
			outbuf->Append(ccstr, cclen);
		}
		outbuf->Append("\r\n");
		////
	}
	if (!m->m_bcc.IsEmpty())
	{
		int bcclen = m->m_bcc.GetLength();
		char *bccstr = (char *)(LPCSTR)m->m_bcc;

		char *cBCC = "BCC: ";
		int cBCCLen = strlen(cBCC);
		outbuf->Append(cBCC, cBCCLen);

		pageCode = m->m_bcc_charsetId;
		if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
		{
			tmpbuf.Copy(bccstr, bcclen);
			BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
			if (ret)
				outbuf->Append(*inbuf);
			else
				outbuf->Append(tmpbuf);
		}
		else {
			outbuf->Append(bccstr, bcclen);
		}
		outbuf->Append("\r\n");
		////
	}
	char datebuff[32];

	CString format = textConfig.m_dateFormat;

	datebuff[0] = 0;
	if (m->m_timeDate >= 0)
	{
		MyCTime tt(m->m_timeDate);
		if (!textConfig.m_bGMTTime) 
		{
			CString lDateTime = tt.FormatLocalTm(format);
			strcpy(datebuff, (LPCSTR)lDateTime);
		}
		else {
			CString lDateTime = tt.FormatGmtTm(format);
			strcpy(datebuff, (LPCSTR)lDateTime);
		}
	}

	char *cDate = "DATE: ";
	int cDateLen = strlen(cDate);
	outbuf->Append(cDate, cDateLen);

	int datelen = strlen(datebuff);
	outbuf->Append(datebuff, datelen);

	outbuf->Append("\r\n");

	outbuf->Append(border);

	outbuf->Append("\r\n");
	////

	fp.Write(outbuf->Data(), outbuf->Count());

	return 1;
}

int MboxMail::printSingleMailToTextFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig)
{
	UINT pageCode;

	MboxMail *m = MboxMail::s_mails[mailPosition];

	SimpleString *outbuf = MboxMail::m_outbuf;
	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;

	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);
	outbuf->ClearAndResize(10000);

	printMailHeaderToTextFile(fp, mailPosition, fpm, textConfig);

	//
	pageCode = 0;
	int textlen = GetMailBody_mboxview(fpm, mailPosition, outbuf, pageCode);  // returns pageCode
	if (textlen != outbuf->Count())
		int deb = 1;

	if (outbuf->Count())
	{
		if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
		{
			int needLength = outbuf->Count() * 2 + 1;
			inbuf->ClearAndResize(needLength);

			BOOL ret = TextUtilsEx::Str2CodePage(outbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
			if (ret) {
				fp.Write(inbuf->Data(), inbuf->Count());
			}
			else {
				fp.Write(outbuf->Data(), outbuf->Count());
			}
		}
		else
			fp.Write(outbuf->Data(), outbuf->Count());
	}
	else
	{
		outbuf->Clear();
		pageCode = 0;
		int textType = 1; // HTML

		int textlen = GetMailBody_mboxview(fpm, mailPosition, outbuf, pageCode, textType);  // returns pageCode
		if (textlen != outbuf->Count())
			int deb = 1;

		if (outbuf->Count())
		{
			MboxMail::m_Html2TextCount++;

			CString bdycharset;
			std::string charSet;
			if (pageCode > 0) {
				BOOL ret = TextUtilsEx::id2charset(pageCode, charSet);
#if 0
				// TODO: charset in the mail Content-Type an in mail body differs; current approach seems to work 
				UINT pageCodeFromBody = getCodePageFromHtmlBody(outbuf, charSet);
				if ((pageCodeFromBody) && (pageCodeFromBody != pageCode))
					int deb = 1;
#endif
			}
			else
			{
				pageCode = getCodePageFromHtmlBody(outbuf, charSet);

				if (pageCode == 0) {
					pageCode = CP_UTF8;
					BOOL ret = TextUtilsEx::id2charset(pageCode, charSet);
				}
			}
				
			bdycharset.Append(charSet.c_str());

			inbuf->Clear();

			CString bdy;
			bool extraHtmlHdr = false;
			if (outbuf->FindNoCase(0, "<body", 5) < 0) // didn't find if true
			{
				extraHtmlHdr = true;
				bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body>";
				inbuf->Append(bdy, bdy.GetLength());
			}

			inbuf->Append(*outbuf);
			if (extraHtmlHdr)
				inbuf->Append("</body></html>\r\n");

			outbuf->Clear();

			UINT outPageCode = CP_UTF8;
			// Below Relies on IHTMLDocument2
			HtmlUtils::GetTextFromIHTMLDocument(inbuf, outbuf, pageCode, outPageCode);
			fp.Write(outbuf->Data(), outbuf->Count());
		}
	}

	outbuf->Clear();
	outbuf->Append("\r\n");
	fp.Write(outbuf->Data(), outbuf->Count());

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		if (pFrame->m_NamePatternParams.m_bAddBreakPageAfterEachMailInPDF)
		{
			char formFeed = 12;
			outbuf->Clear();
			outbuf->Append(formFeed);
			fp.Write(outbuf->Data(), outbuf->Count());
		}
	}

	return 1;
}

UINT getCodePageFromHtmlBody(SimpleString *buffer, std::string &charset)
{
	char *pat = "charset=";
	charset.assign("");
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
		codePage = TextUtilsEx::charset2Id(charset.c_str());
	}
	int deb = 1;
	return codePage;
}


int MboxMail::printSingleMailToHtmlFile(/*out*/CFile &fp, int mailPosition, /*in - mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig, bool firstMail)
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
	if (textlen != outbuflarge->Count())
		int deb = 1;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	//CString marginLeft = "margin-left:5px;";
	//CString backgroundColor = "background-color:transparent;";
	//CString marginLeft = "margin-left:0px;";
	//CString backgroundColor = "background-color:white;";
#if 0
	if (!pFrame->m_NamePatternParams.m_bKeepMailBodyBackgroundColor)
	{
		CString body = "\r\n\r\n<html><head></head><body style=\'background-color:white\'></body></html>\r\n";
		fp.Write(body, body.GetLength());
		body = "\r\n\r\n<html><head></head><body bgColor=white></body></html>\r\n";
		fp.Write(body, body.GetLength());
	}
#endif

	CString htmlHdrFldNameStyle;
	CString fldNameFontStyle;
	CString fldTextFontStyle;
	if (pFrame)
	{
		CreateFldFontStyle(pFrame->m_HdrFldConfig, fldNameFontStyle, fldTextFontStyle);
		htmlHdrFldNameStyle.Format("<style>\r\n.hdrfldname{%s}\r\n.hdrfldtext{%s}\r\n</style>", fldNameFontStyle, fldTextFontStyle);
	}

	if (outbuflarge->Count() != 0)
	{
		//int retBG = NListView::RemoveBackgroundColor(outbuflarge->Data(), outbuflarge->Count(), &tmpbuf, mailPosition);
		//outbuflarge->Copy(tmpbuf);

		CString bdycharset = "UTF-8";
		CString bdy;

		SimpleString *workbuf = MboxMail::m_workbuf;
		workbuf->ClearAndResize(outbuflarge->Count() * 2);
		bool useMailPosition = true;
		NListView::UpdateInlineSrcImgPath(outbuflarge->Data(), outbuflarge->Count(), workbuf, 0, mailPosition, useMailPosition);
		outbuflarge->Copy(*workbuf);

		bdy = "\r\n\r\n<div style=\'background-color:transparent;margin-left:5px;text-align:left\'>\r\n";
		fp.Write(bdy, bdy.GetLength());

		CString font = "\">\r\n";
		
		if (pFrame)
		{
			if (pFrame->m_NamePatternParams.m_bAddBackgroundColorToMailHeader)
			{
				bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset +
					"\">" + htmlHdrFldNameStyle + "</head>\r\n<body>\r\n<div style=\"background-color:#eee9e9" + font;
			}
			else
			{
				bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + 
					"\">" + htmlHdrFldNameStyle + "</head>\r\n<body>\r\n<div style=\"font-weight:normal" + font;
			}
			fp.Write(bdy, bdy.GetLength());
		}
		else
		{
			bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset +
				"\">" + htmlHdrFldNameStyle + "</head>\r\n<body>\r\n<div style=\"background-color:#eee9e9" + font;
			fp.Write(bdy, bdy.GetLength());
		}

		int ret;
		if (pFrame)
		{
			ret = printMailHeaderToHtmlFile(fp, mailPosition, fpm, textConfig, pFrame->m_HdrFldConfig);
		}
		else
			ret = printMailHeaderToHtmlFile(fp, mailPosition, fpm, textConfig);

		bdy = "</div></body></html>";
		fp.Write(bdy, bdy.GetLength());

		bdy = "\r\n</div>";
		fp.Write(bdy, bdy.GetLength());


		bdy = "\r\n<div style=\'background-color:transparent;margin-left:5px;text-align:left\'><br>\r\n";
		fp.Write(bdy, bdy.GetLength());

		bool extraHtmlHdr = false;
		if (outbuflarge->FindNoCase(0, "<html", 5) < 0) // didn't find if true
		{
			if (outbuflarge->FindNoCase(0, "<body", 5) < 0) // didn't find if true
			{
				extraHtmlHdr = true;
				bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body>\r\n";
				fp.Write(bdy, bdy.GetLength());
			}
			else
			{
				bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head></html>\r\n";
				fp.Write(bdy, bdy.GetLength());
			}
		}
		else
		{
			bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head></html>\r\n";
			fp.Write(bdy, bdy.GetLength());
		}

		std::string charSet;
		if (pageCode == 0) {
			pageCode = getCodePageFromHtmlBody(outbuflarge, charSet);
		}

		if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
		{
			int needLength = outbuflarge->Count() * 2 + 1;
			inbuf->ClearAndResize(needLength);

			BOOL ret = TextUtilsEx::Str2CodePage(outbuflarge, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
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
			bdy = "\r\n</body></html>";
			fp.Write(bdy, bdy.GetLength());
		}

		bdy = "\r\n</div>";
		fp.Write(bdy, bdy.GetLength());

		if (pFrame)
		{
			if (pFrame->m_NamePatternParams.m_bAddBreakPageAfterEachMailInPDF && !pFrame->m_NamePatternParams.m_bPrintToSeparatePDFFiles)
			{
				bdy = "\r\n<div style=\"page-break-before:always\"></div>";
				fp.Write(bdy, bdy.GetLength());
			}
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

		CString bdy;

		CString bdycharset = "UTF-8";

		bdy = "\r\n\r\n<div style=\'background-color:transparent;margin-left:5px;text-align:left\'>\r\n";
		fp.Write(bdy, bdy.GetLength());

		CString font = "\">\r\n";

		if (pFrame)
		{
			if (pFrame->m_NamePatternParams.m_bAddBackgroundColorToMailHeader)
			{
				//bdy = "<html><head></head><body>\r\n<div style=\"background-color:#eee9e9" + font;

				bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset +
					"\">" + htmlHdrFldNameStyle + "</head>\r\n<body>\r\n<div style=\"background-color:#eee9e9" + font;
			}
			else
			{
				//bdy = "<html><head></head><body>\r\n" + font;

				bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset +
					"\">" + htmlHdrFldNameStyle + "</head>\r\n<body>\r\n<div style=\"font-weight:normal" + font;
			}
			fp.Write(bdy, bdy.GetLength());
		}
		else
		{
			//bdy = "<html><head></head><body>\r\n<div style=\"background-color:#eee9e9" + font;

			bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset +
				"\">" + htmlHdrFldNameStyle + "</head>\r\n<body>\r\n<div style=\"background-color:#eee9e9" + font;

			fp.Write(bdy, bdy.GetLength());
		}

		int ret;
		if (pFrame)
			ret = printMailHeaderToHtmlFile(fp, mailPosition, fpm, textConfig, pFrame->m_HdrFldConfig);
		else
			ret = printMailHeaderToHtmlFile(fp, mailPosition, fpm, textConfig);

		bdy = "</div></body></html>";
		fp.Write(bdy, bdy.GetLength());

		bdy = "\r\n</div>";
		fp.Write(bdy, bdy.GetLength());

		bdy = "\r\n<div style=\'background-color:transparent;margin-left:5px;text-align:left\'>\r\n";
		fp.Write(bdy, bdy.GetLength());

		bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body><br>\r\n";
		fp.Write(bdy, bdy.GetLength());

		char *inData = outbuflarge->Data();
		int inDataLen = outbuflarge->Count();
		workbuf->ClearAndResize(2 * outbuflarge->Count());
		TextUtilsEx::EncodeAsHtml(inData, inDataLen, workbuf);
		outbuflarge->Copy(*workbuf);

		if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
		{
#if 0
			char *inData = outbuflarge->Data();
			int inDataLen = outbuflarge->Count();
			TextUtilsEx::EncodeAsHtml(inData, inDataLen, workbuf);
			outbuflarge->Copy(*workbuf);
#endif

			int needLength = outbuflarge->Count() * 2 + 1;
			inbuf->ClearAndResize(needLength);

			BOOL ret = TextUtilsEx::Str2CodePage(outbuflarge, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
			// TODO: replace CR LF with <br> to enable line wrapping
			if (ret) {
				fp.Write(inbuf->Data(), inbuf->Count());
			}
			else {
				fp.Write(outbuflarge->Data(), outbuflarge->Count());
			}
		}
		else
			fp.Write(outbuflarge->Data(), outbuflarge->Count());

		bdy = "</body></html>";
		fp.Write(bdy, bdy.GetLength());

		bdy = "\r\n</div>";
		fp.Write(bdy, bdy.GetLength());


		if (pFrame)
		{
			if (pFrame->m_NamePatternParams.m_bAddBreakPageAfterEachMailInPDF && !pFrame->m_NamePatternParams.m_bPrintToSeparatePDFFiles)
			{
				bdy = "<div style=\"page-break-before:always\"></div>";
				fp.Write(bdy, bdy.GetLength());
			}
		}
	}

	CString bdy = "\r\n<div><br></div>";
	fp.Write(bdy, bdy.GetLength());

	return 1;
}

void encodeTextAsHtml(SimpleString &txt) 
{
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
	// encodeTextAsHtml(tmpbuf);

	UINT pageCode = m->m_subj_charsetId;
	if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
	{
		BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
		if (ret)
		{
			encodeTextAsHtml(*inbuf);
			outbuf.Append(*inbuf);
		}
		else
		{
			encodeTextAsHtml(tmpbuf);
			outbuf.Append(tmpbuf);
		}
	}
	else {
		encodeTextAsHtml(tmpbuf);
		outbuf.Append(tmpbuf);
	}
	outbuf.Append("<br>\r\n");

	//
	int fromlen = m->m_from.GetLength();
	char *fromstr = (char *)(LPCSTR)m->m_from;

	char *cFrom = "FROM: ";
	int cFromLen = strlen(cFrom);
	outbuf.Append(cFrom, cFromLen);
	tmpbuf.Copy(fromstr, fromlen);
	// encodeTextAsHtml(tmpbuf);

	pageCode = m->m_from_charsetId;
	if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
	{
		BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
		if (ret)
		{
			encodeTextAsHtml(*inbuf);
			outbuf.Append(*inbuf);
		}
		else
		{
			encodeTextAsHtml(tmpbuf);
			outbuf.Append(tmpbuf);
		}
	}
	else {
		encodeTextAsHtml(tmpbuf);
		outbuf.Append(tmpbuf);
	}
	outbuf.Append("<br>\r\n");

	//
	int tolen = m->m_to.GetLength();
	char *tostr = (char *)(LPCSTR)m->m_to;

	char *cTo = "TO: ";
	int cToLen = strlen(cTo);
	outbuf.Append(cTo, cToLen);
	tmpbuf.Copy(tostr, tolen);
	// encodeTextAsHtml(tmpbuf);

	pageCode = m->m_to_charsetId;
	if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
	{
		BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
		if (ret)
		{
			encodeTextAsHtml(*inbuf);
			outbuf.Append(*inbuf);
		}
		else
		{
			encodeTextAsHtml(tmpbuf);
			outbuf.Append(tmpbuf);
		}
	}
	else {
		encodeTextAsHtml(tmpbuf);
		outbuf.Append(tmpbuf);
	}
	outbuf.Append("<br>\r\n");

	if (!m->m_cc.IsEmpty())
	{
		//
		int cclen = m->m_cc.GetLength();
		char *ccstr = (char *)(LPCSTR)m->m_cc;

		char *cCC = "CC: ";
		int cCCLen = strlen(cCC);
		outbuf.Append(cCC, cCCLen);
		tmpbuf.Copy(ccstr, cclen);
		// encodeTextAsHtml(tmpbuf);

		pageCode = m->m_cc_charsetId;
		if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
		{
			BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
			if (ret)
			{
				encodeTextAsHtml(*inbuf);
				outbuf.Append(*inbuf);
			}
			else
			{
				encodeTextAsHtml(tmpbuf);
				outbuf.Append(tmpbuf);
			}
		}
		else {
			encodeTextAsHtml(tmpbuf);
			outbuf.Append(tmpbuf);
		}
		outbuf.Append("<br>\r\n");
	}

	if (!m->m_bcc.IsEmpty())
	{
		//
		int bcclen = m->m_bcc.GetLength();
		char *bccstr = (char *)(LPCSTR)m->m_bcc;

		char *cBCC = "BCC: ";
		int cBCCLen = strlen(cBCC);
		outbuf.Append(cBCC, cBCCLen);
		tmpbuf.Copy(bccstr, bcclen);
		// encodeTextAsHtml(tmpbuf);

		pageCode = m->m_bcc_charsetId;
		if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
		{
			BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
			if (ret)
			{
				encodeTextAsHtml(*inbuf);
				outbuf.Append(*inbuf);
			}
			else
			{
				encodeTextAsHtml(tmpbuf);
				outbuf.Append(tmpbuf);
			}
		}
		else {
			encodeTextAsHtml(tmpbuf);
			outbuf.Append(tmpbuf);
		}
		outbuf.Append("<br>\r\n");
	}

	//
	char datebuff[32];

	CString format = textConfig.m_dateFormat;

	datebuff[0] = 0;
	if (m->m_timeDate >= 0)
	{
		MyCTime tt(m->m_timeDate);
		if (!textConfig.m_bGMTTime) 
		{
			CString lDateTime = tt.FormatLocalTm(format);
			strcpy(datebuff, (LPCSTR)lDateTime);
		}
		else {
			CString lDateTime = tt.FormatGmtTm(format);
			strcpy(datebuff, (LPCSTR)lDateTime);
		}
	}

	char *cDate = "DATE: ";
	int cDateLen = strlen(cDate);
	outbuf.Append(cDate, cDateLen);

	int datelen = strlen(datebuff);
	outbuf.Append(datebuff, datelen);

	outbuf.Append("<br>\r\n");
	fp.Write(outbuf.Data(), outbuf.Count());

	return 1;
}

int MboxMail::CreateFldFontStyle(HdrFldConfig &hdrFieldConfig, CString &fldNameFontStyle, CString &fldTextFontStyle)
{
	if (hdrFieldConfig.m_bHdrFontDflt == 0)
	{
		int nameFontSize = hdrFieldConfig.m_nHdrFontSize;
		int textFontSize = hdrFieldConfig.m_nHdrFontSize;

		fldTextFontStyle.Format("font-size:%dpx;", textFontSize);

		if (hdrFieldConfig.m_bHdrBoldFldName)
		{
			fldNameFontStyle.Format("font-weight:bold;font-size:%dpx;",  nameFontSize);
		}
		else
			fldNameFontStyle.Format("font-size:%dpx;", nameFontSize);
	}
	else
	{

		int nameFontSize = hdrFieldConfig.m_HdrFldFontName.m_nFontSize;
		int textFontSize = hdrFieldConfig.m_HdrFldFontText.m_nFontSize;
		//
		CString nameFont = hdrFieldConfig.m_HdrFldFontName.m_fontName;
		CString textFont = hdrFieldConfig.m_HdrFldFontText.m_fontName;

		//
		CString nameFontFamily = hdrFieldConfig.m_HdrFldFontName.m_genericFontName;
		CString textFontFamily = hdrFieldConfig.m_HdrFldFontText.m_genericFontName;
		//

		CString nameFontWeight;
		if (hdrFieldConfig.m_HdrFldFontName.m_nFontStyle == 400)
			nameFontWeight = "normal";
		else if (hdrFieldConfig.m_HdrFldFontName.m_bIsBold)
			nameFontWeight = "bold";
		else
			nameFontWeight.Format("%d", hdrFieldConfig.m_HdrFldFontName.m_nFontStyle);

		CString textFontWeight;
		if (hdrFieldConfig.m_HdrFldFontText.m_nFontStyle == 400)
			textFontWeight = "normal";
		else if (hdrFieldConfig.m_HdrFldFontText.m_bIsBold)
			textFontWeight = "bold";
		else
			textFontWeight.Format("%d", hdrFieldConfig.m_HdrFldFontText.m_nFontStyle);

		CString nameFontStyle = "normal";
		if (hdrFieldConfig.m_HdrFldFontName.m_bIsItalic)
			nameFontStyle = "italic";

		CString textFontStyle = "normal";
		if (hdrFieldConfig.m_HdrFldFontText.m_bIsItalic)
			textFontWeight = "italic";


		fldNameFontStyle.Format("font-style:%s;font-weight:%s;font-size:%dpx;font-family:\"%s\",%s;",
			nameFontStyle, nameFontWeight, nameFontSize, nameFont, nameFontFamily);

		fldTextFontStyle.Format("font-style:%s;font-weight:%s;font-size:%dpx;font-family:\"%s\",%s;",
			textFontStyle, textFontWeight, textFontSize, textFont, textFontFamily);
	}

	return 1;
}

int MboxMail::printMailHeaderToHtmlFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig, HdrFldConfig &hdrFieldConfig)
{
	int fldNumb;
	int fldNameLength;
	int fldTxtLength;
	char *fldText;
	char *fldName = "";
	int ii;

	char *token = 0;
	int tokenlen = 0;

	SimpleString outbuf(1024, 256);
	SimpleString tmpbuf(1024, 256);

	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;
	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);

	MboxMail *m = s_mails[mailPosition];

	UINT pageCode = 0;
	CString txt;
	for (ii = 0; ii < HdrFldList::HDR_FLD_MAX; ii++)
	{
		if (hdrFieldConfig.m_HdrFldList.IsFldSet(ii))
		{
			fldNumb = ii+1;
			if (fldNumb == HdrFldList::HDR_FLD_SUBJECT)
			{
				fldName = "SUBJECT: ";
				fldNameLength = strlen(fldName);
				fldTxtLength = m->m_subj.GetLength();
				fldText = (char *)(LPCSTR)m->m_subj;
				pageCode = m->m_subj_charsetId;
			}
			else if (fldNumb == HdrFldList::HDR_FLD_FROM)
			{
				fldName = "FROM: ";
				fldNameLength = strlen(fldName);
				fldTxtLength = m->m_from.GetLength();
				fldText = (char *)(LPCSTR)m->m_from;
				pageCode = m->m_from_charsetId;
			}
			else if (fldNumb == HdrFldList::HDR_FLD_TO)
			{
				fldName = "TO: ";
				fldNameLength = strlen(fldName);
				fldTxtLength = m->m_to.GetLength();
				fldText = (char *)(LPCSTR)m->m_to;
				pageCode = m->m_to_charsetId;
			}
			else if (fldNumb == HdrFldList::HDR_FLD_CC)
			{
				fldName = "CC: ";
				fldNameLength = strlen(fldName);
				fldTxtLength = m->m_cc.GetLength();
				fldText = (char *)(LPCSTR)m->m_cc;
				pageCode = m->m_cc_charsetId;
				if (fldTxtLength <= 0)
					continue;
			}
			else if (fldNumb == HdrFldList::HDR_FLD_BCC)
			{
				fldName = "BCC: ";
				fldNameLength = strlen(fldName);
				fldTxtLength = m->m_bcc.GetLength();
				fldText = (char *)(LPCSTR)m->m_bcc;
				pageCode = m->m_bcc_charsetId;
				if (fldTxtLength <= 0)
					continue;
			}
			else if (fldNumb == HdrFldList::HDR_FLD_DATE)
			{
				fldName = "DATE: ";
				fldNameLength = strlen(fldName);
				fldTxtLength = 0;
				fldText = "";
				continue;
			}
			else
			{
				continue;
			}

			//txt.Format("<span style=\'font-style:%s;font-weight:%s;font-size:%dpx;font-family:\"%s\",%s;\'>", 
				//nameFontStyle, nameFontWeight, nameFontSize, nameFont, nameFontFamily);
			//outbuf.Append(txt, txt.GetLength());
			//outbuf.Append(fldNameFontStyle, fldNameFontStyle.GetLength());

			txt.Format("<span class=\'hdrfldname\'>");
			outbuf.Append(txt, txt.GetLength());
			outbuf.Append(fldName, fldNameLength);
			outbuf.Append("</span>");

			tmpbuf.Copy(fldText, fldTxtLength);
			// encodeTextAsHtml(tmpbuf);

			//txt.Format("<span style=\'font-style:%s;font-weight:%s;font-size:%dpx;font-family:\"%s\",%s;\'>", 
				//textFontStyle, textFontWeight, textFontSize, textFont, textFontFamily);
			//outbuf.Append(txt, txt.GetLength());
			//outbuf.Append(fldTextFontStyle, fldTextFontStyle.GetLength());

			txt.Format("<span class=\'hdrfldtext\'>");
			outbuf.Append(txt, txt.GetLength());
			if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
			{
				BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
				if (ret)
				{
					encodeTextAsHtml(*inbuf);
					outbuf.Append(*inbuf);
				}
				else
				{
					encodeTextAsHtml(tmpbuf);
					outbuf.Append(tmpbuf);
				}
			}
			else {
				encodeTextAsHtml(tmpbuf);
				outbuf.Append(tmpbuf);
			}
			outbuf.Append("</span>");
			outbuf.Append("<br>\r\n");
		}
	}


	//
	fldNumb = HdrFldList::HDR_FLD_DATE -1;
	if (hdrFieldConfig.m_HdrFldList.IsFldSet(fldNumb))
	{
		char datebuff[32];

		CString format = textConfig.m_dateFormat;

		datebuff[0] = 0;
		if (m->m_timeDate >= 0)
		{
			MyCTime tt(m->m_timeDate);
			if (!textConfig.m_bGMTTime)
			{
				CString lDateTime = tt.FormatLocalTm(format);
				strcpy(datebuff, (LPCSTR)lDateTime);
			}
			else {
				CString lDateTime = tt.FormatGmtTm(format);
				strcpy(datebuff, (LPCSTR)lDateTime);
			}
		}

		fldName = "DATE: ";
		fldNameLength = strlen(fldName);
		fldTxtLength = m->m_subj.GetLength();
		fldText = (char *)datebuff;

		//txt.Format("<span style=\'font-style:%s;font-weight:%s;font-size:%dpx;font-family:\"%s\",%s;\'>",
			//nameFontStyle, nameFontWeight, nameFontSize, nameFont, nameFontFamily);
		//outbuf.Append(txt, txt.GetLength());
		//outbuf.Append(fldNameFontStyle, fldNameFontStyle.GetLength());

		txt.Format("<span class=\'hdrfldname\'>");
		outbuf.Append(txt, txt.GetLength());
		outbuf.Append(fldName, fldNameLength);
		outbuf.Append("</span>");

		//txt.Format("<span style=\'font-style:%s;font-weight:%s;font-size:%dpx;font-family:\"%s\",%s;\'>",
			//textFontStyle, textFontWeight, textFontSize, textFont, textFontFamily);
		//outbuf.Append(txt, txt.GetLength());
		//outbuf.Append(fldTextFontStyle, fldTextFontStyle.GetLength());

		txt.Format("<span class=\'hdrfldtext\'>");
		outbuf.Append(txt, txt.GetLength());

		int datelen = strlen(datebuff);
		outbuf.Append(datebuff, datelen);

		outbuf.Append("</span>");
		outbuf.Append("<br>\r\n");
	}

	fp.Write(outbuf.Data(), outbuf.Count());

	return 1;
}

bool ALongRightProcessProcPrintMailArchive(const CUPDUPDATA* pCUPDUPData)
{
	PRINT_MAIL_ARCHIVE_ARGS *args = (PRINT_MAIL_ARCHIVE_ARGS*)pCUPDUPData->GetAppData();
	MboxMail::pCUPDUPData = pCUPDUPData;

	HANDLE h = GetCurrentThread();
	BOOL prio = SetThreadPriority(h, THREAD_PRIORITY_ABOVE_NORMAL);
	DWORD myThreadId = GetCurrentThreadId();
	DWORD myThreadPri = GetThreadPriority(h);
	TRACE(_T("(ALongRightProcessProcPrintMailArchive) threadId=%ld threadPriority=%ld\n"), myThreadId, myThreadPri);

	BOOL progressBar = TRUE;
	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as param to printMailArchiveToTextFile
	CString errorText;
	int retpos = MboxMail::printMailArchiveToTextFile(args->textConfig, args->textFile, args->firstMail, args->lastMail, args->textType, progressBar, errorText);
	args->errorText = errorText;
	args->exitted = TRUE;
	return true;
}

BOOL MboxMail::GetPrintCachePath(CString &prtCachePath)
{
	prtCachePath.Empty();

	CString mailFile = MboxMail::s_path;

	if (s_path.IsEmpty()) {
		CString txt = _T("Please open mail file first.");
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return FALSE;
	}

	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	if (path.IsEmpty()) {
		CString txt = _T("No path to archive file folder.");
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return FALSE;  // Hopefully s_path wil fail first
	}

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(mailFile, driveName, directory, fileNameBase, fileNameExtention);

	CString mailArchiveFileName;
	FileUtils::CPathStripPath((char*)(LPCSTR)MboxMail::s_path, mailArchiveFileName);

	int position = mailArchiveFileName.ReverseFind('.');
	CString baseFileArchiveName;
	if (position >= 0)
		baseFileArchiveName = mailArchiveFileName.Mid(0, position);
	else
		baseFileArchiveName = mailArchiveFileName;

	CString printCachePath;
	BOOL ret = FileUtils::CPathGetPath(MboxMail::s_path, printCachePath);
	printCachePath.Append("\\");
	printCachePath.Append("PrintCache");

	BOOL createDirOk = TRUE;
	if (!FileUtils::PathDirExists(printCachePath)) {
		createDirOk = CreateDirectory(printCachePath, NULL);
	}

	if (!createDirOk) {
		CString txt = _T("Could not create \"") + printCachePath;
		txt += _T("\" folder for print destination.\nResolve the problem and try again.");
		HWND h = NULL;
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return FALSE;
	}

	printCachePath.Append("\\");
	printCachePath.Append(baseFileArchiveName);

	createDirOk = TRUE;
	if (!FileUtils::PathDirExists(printCachePath))
		createDirOk = CreateDirectory(printCachePath, NULL);

	if (!createDirOk) {
		CString txt = _T("Could not create \"") + printCachePath;
		txt += _T("\" folder for print destination.\nResolve the problem and try again.");
		HWND h = NULL;
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return FALSE;
	}
	prtCachePath.Append(printCachePath);
	return TRUE;
}


int MboxMail::exportToTextFile(TEXTFILE_CONFIG &textConfig, CString &textFileName, int firstMail, int lastMail, MailIndexList *selectedMailIndexList, int textType, BOOL progressBar)
{
	CString colLabels;
	CString separator;

	CFile fp;
	CString textFile;
	CString errorText;
	bool fileExists = false;
	int ret = 1;

	if (!progressBar) // no worker thread
	{
		CString targetPrintSubFolder;
		if (selectedMailIndexList)
		{
			ret = MakeFileNameFromMailArchiveName(textType, textFile, targetPrintSubFolder, fileExists, errorText);
		}
		else if (firstMail == lastMail)
		{
			ret = MakeFileNameFromMailHeader(firstMail, textType, textFile, targetPrintSubFolder, fileExists, errorText);
		}
		else
		{
			ret = MakeFileNameFromMailArchiveName(textType, textFile, targetPrintSubFolder, fileExists, errorText);
		}
	}
	else  // worker thread
	{
		CString targetPrintSubFolder;
		ret = MakeFileNameFromMailArchiveName(textType, textFile, targetPrintSubFolder, fileExists, errorText);
	}

	if (ret < 0) {
		HWND h = NULL; 
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	textFileName = textFile;
	int textFileLength = textFile.GetLength();
	int maxPath = _MAX_PATH;

	if (!progressBar) // no worker thread
	{
		if (!fp.Open(textFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone)) {
			CString txt = _T("Could not create \"") + textFile;
			txt += _T("\" file.\nMake sure file is not open on other applications.");
			HWND h = NULL; // we don't have any window yet
			int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}

		CFile fpm;
		if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite)) {
			CString txt = _T("Could not open mail archive \"") + s_path;
			HWND h = NULL; // we don't have any window yet
			int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			fp.Close();
			return -1;
		}

		if (textType == 0) {
			if (textConfig.m_nCodePageId == CP_UTF8) {
				const char *BOM_UTF8 = "\xEF\xBB\xBF";
				fp.Write(BOM_UTF8, 3);
			}
		}

		if (selectedMailIndexList == 0) 
		{
			for (int i = firstMail; i <= lastMail; i++)
			{
				if (textType == 0)
					printSingleMailToTextFile(fp, i, fpm, textConfig);
				else {
					printSingleMailToHtmlFile(fp, i, fpm, textConfig, i == firstMail);
				}
			}
		}
		else 
		{
			int i;
			firstMail = (*selectedMailIndexList)[0];
			int cnt = selectedMailIndexList->GetCount();
			for (int j = 0; j < cnt; j++)
			{
				//i = selectedMailIndexList->ElementAt(j);
				i = (*selectedMailIndexList)[j];
				if (textType == 0)
					printSingleMailToTextFile(fp, i, fpm, textConfig);
				else {
					printSingleMailToHtmlFile(fp, i, fpm, textConfig, i == firstMail);
				}
			}
		}

		fp.Close();
		fpm.Close();
	}
	else  // worker thread
	{
		PRINT_MAIL_ARCHIVE_ARGS args;
		args.textConfig = textConfig;
		args.textFile = textFile;
		args.firstMail = firstMail;
		args.lastMail = lastMail;
		args.textType = textType;
		args.terminated = FALSE; // flag set by MboxMail::printMailArchiveToTextFile to ack termination request
		args.exitted = FALSE; // flag set by ALongRightProcessProcPrintMailArchive  upon return

		MboxMail::m_Html2TextCount = 0;

		CUPDialog	Dlg(AfxGetMainWnd()->GetSafeHwnd(), ALongRightProcessProcPrintMailArchive, (LPVOID)(PRINT_MAIL_ARCHIVE_ARGS*)&args);

		INT_PTR nResult = Dlg.DoModal();

		if (!nResult) // should never be true ?
		{
			MboxMail::assert_unexpected();
			return -1;
		}

		TRACE(_T("m_Html2TextCount=%d MailCount=%d\n"), m_Html2TextCount, MboxMail::s_mails.GetCount());

		int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
		int retResult = LOWORD(nResult);

		if (retResult != IDOK)
		{  // IDOK==1, IDCANCEL==2
			// We should be here when user selects Cancel button
			ASSERT(cancelledbyUser == TRUE);
			int loopCnt = 20;
			DWORD tc_start = GetTickCount();
			while ((loopCnt-- > 0) && (args.exitted == FALSE))
			{
				Sleep(25);
			}
			DWORD tc_end = GetTickCount();
			DWORD delta = tc_end - tc_start;
			TRACE("(exportToTextFile)Waited %ld milliseconds for thread to exist.\n", delta);
		}

		MboxMail::pCUPDUPData = NULL;

		if (!args.errorText.IsEmpty()) {
			HWND h = NULL; // we don't have any window yet
			int answer = ::MessageBox(h, args.errorText, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
			return -1;
		}
	}
	return 1;
}

int MboxMail::printMailArchiveToTextFile(TEXTFILE_CONFIG &textConfig, CString &textFile, int firstMail, int lastMail, int textType, BOOL progressBar, CString &errorText)
{
	CString colLabels;
	CString separator;

	CString mailFile = MboxMail::s_path;

	// Simplify and start MessageBox here if error inset of errorText ?
	// Cons is that MessageBox and progress bar are visible at the same time

	if (s_path.IsEmpty()) {
		errorText = _T("Please open mail file first.");
		return -1;
	}

	CFileException ex;
	CFile fp;
	if (!fp.Open(textFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &ex)) {
		TCHAR szError[1024];
		ex.GetErrorMessage(szError, 1024);
		CFileStatus rStatus;
		BOOL ret = fp.GetStatus(rStatus);
		errorText = _T("Could not create \"") + textFile;
		errorText += _T("\" file.\nMake sure file is not open on other applications.");
		return -1;
	}

	CFile fpm;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite)) {
		errorText = _T("Could not open mail archive \"") + s_path + _T("\"");
		fp.Close();
		return -1;
	}

	if (textType == 0) {
		if (textConfig.m_nCodePageId == CP_UTF8) {
			const char *BOM_UTF8 = "\xEF\xBB\xBF";
			fp.Write(BOM_UTF8, 3);
		}
	}

	if (progressBar && MboxMail::pCUPDUPData)
	{
		if (textType == 0)
			MboxMail::pCUPDUPData->SetProgress(_T("Printing mails to TEXT file ..."), 0);
		else
			MboxMail::pCUPDUPData->SetProgress(_T("Printing mails to HTML file ..."), 0);
	}

	double delta = (double)(lastMail - firstMail);
	if (delta <= 0) delta = 1;
	double step = delta / 100;
	double curstep = 1;


	if (progressBar && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	CString fileNum;

	int cnt = lastMail - firstMail;
	if (cnt <= 0)
		cnt = 1;
	for (int i = firstMail; i <= lastMail; i++)
	{
		if (textType == 0) {
			printSingleMailToTextFile(fp, i, fpm, textConfig);
		}
		else {
			printSingleMailToHtmlFile(fp, i, fpm, textConfig, i == firstMail);
		}

		if (progressBar && MboxMail::pCUPDUPData)
		{
			double newstep = (double)(i - firstMail) / step + 1;
			if (newstep > curstep)
			{
				if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
					int deb = 1;
					break;
				}

				if (MboxMail::pCUPDUPData)
					MboxMail::pCUPDUPData->SetProgress((UINT_PTR)newstep);
				curstep = newstep;
			}

			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				int deb = 1;
				break;
			}

			int nFileNum = (i - firstMail  + 1);
			int stepSize = 100;
			if (textType == 0)
				stepSize = 100;  // 1000 ??

			if ((nFileNum % stepSize) == 0) 
			{
				if (textType == 0)
					fileNum.Format(_T("Printing mails to single TEXT file ... %d of %d"), nFileNum, cnt);
				else
					fileNum.Format(_T("Printing mails to single HTML file ... %d of %d"), nFileNum, cnt);

				if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(newstep));
			}
		}
	}

	fp.Close();
	fpm.Close();

	return 1;
}

int MboxMail::printMailArchiveToCSVFile(CSVFILE_CONFIG &csvConfig, CString &csvFile, int firstMail, int lastMail, MailIndexList *selectedMailIndexList, BOOL progressBar, CString &errorText)
{
	CString colLabels;
	CString separator;
	int retval;

	CString mailFile = MboxMail::s_path;

	// Simplify and start MessageBox here if error inset of errorText ?
	// Cons is tehat MessageBox and progress bar are visible at the same time

	if (s_path.IsEmpty()) {
		errorText = _T("Please open mail file first.");
		return -1;
	}

	CFileException ex;
	CFile fp;
	if (!fp.Open(csvFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &ex)) 
	{
		TCHAR szError[1024];
		ex.GetErrorMessage(szError, 1024);
		CFileStatus rStatus;
		BOOL ret = fp.GetStatus(rStatus);
		errorText = _T("Could not create \"") + csvFile;
		errorText += _T("\" file.\nMake sure file is not open on other applications.");
		return -1;
	}

	CFile fpm;
	if (csvConfig.m_bContent) 
	{
		if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite)) 
		{
			errorText = _T("Could not open mail archive \"") + s_path + _T("\"");
			fp.Close();
			return -1;
		}
	}

	if (progressBar && MboxMail::pCUPDUPData)
	{
			MboxMail::pCUPDUPData->SetProgress(_T("Printing mails to CSV file ..."), 0);
	}

	double delta = (double)(lastMail - firstMail);
	if (delta <= 0) delta = 1;
	double step = delta / 100;
	double curstep = 1;

	if (progressBar && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	if (1) 
	{
		if (csvConfig.m_nCodePageId == CP_UTF8) {
			const char *BOM_UTF8 = "\xEF\xBB\xBF";
			fp.Write(BOM_UTF8, 3);
		}
	}

	CString fileNum;
	bool first = true;
	if (selectedMailIndexList == 0)
	{
		int cnt = lastMail - firstMail;
		if (cnt <= 0)
			cnt = 1;
		for (int i = firstMail; i <= lastMail; i++)
		{
			retval = printSingleMailToCSVFile(fp, i, fpm, csvConfig, first);
			first = false;

			if (progressBar && MboxMail::pCUPDUPData)
			{
				double newstep = (double)(i - firstMail) / step + 1;
				if (newstep > curstep)
				{
					if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
						int deb = 1;
						break;
					}

					if (MboxMail::pCUPDUPData)
						MboxMail::pCUPDUPData->SetProgress((UINT_PTR)newstep);
					curstep = newstep;
				}

				if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
					int deb = 1;
					break;
				}

				int nFileNum = (i - firstMail + 1);
				int stepSize = 100;

				if ((nFileNum % stepSize) == 0)
				{
					fileNum.Format(_T("Printing mails to CSV file ... %d of %d"), nFileNum, cnt);

					if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(newstep));
				}
			}
		}
	}
	else
	{
		int i;
		int cnt = selectedMailIndexList->GetCount();

		double delta = (double)(cnt);
		if (delta <= 0) delta = 1;
		double step = delta / 100;
		double curstep = 1;

		for (int j = 0; j < cnt; j++)
		{
			i = (*selectedMailIndexList)[j];
			retval = printSingleMailToCSVFile(fp, i, fpm, csvConfig, first);
			first = false;

			if (progressBar && MboxMail::pCUPDUPData)
			{
				double newstep = (double)(j - 0 + 1) / step + 1;
				if (newstep > curstep)
				{
					if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
						int deb = 1;
						break;
					}

					if (MboxMail::pCUPDUPData)
						MboxMail::pCUPDUPData->SetProgress((UINT_PTR)newstep);
					curstep = newstep;
				}

				if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
					int deb = 1;
					break;
				}

				int nFileNum = (j + 1);
				int stepSize = 100;

				if ((nFileNum % stepSize) == 0)
				{
					fileNum.Format(_T("Printing mails to CSV file ... %d of %d"), nFileNum, cnt);

					if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(newstep));
				}
			}
		}
	}

	fp.Close();
	if (csvConfig.m_bContent)
		fpm.Close();

	if (!csvConfig.m_bContent)
		Sleep(500);  // helps to deliver all messages to Dialog and display 100% progress bar

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
	char c;
	for (int i = 0; i < fldlen; i++)
	{
		c = fldstr[i];
		if (c == sepchar) {
			workbuff[outcnt++] = sepchar;
			workbuff[outcnt++] = sepchar;
		}
		else
			workbuff[outcnt++] = fldstr[i];

	}
	workbuff[outcnt] = 0;
	return outcnt;
}

int MboxMail::DecodeBody(CFile &fpm, MailBodyContent *body, int mailPosition, SimpleString *outbuf)
{
	_int64 fileOffset;
	int bodyCnt = 0;

	MboxMail *m = MboxMail::s_mails[mailPosition];

	// We are using global buffers so check and assert if we collide. 
	SimpleString *inbuf = MboxMail::m_indata;
	SimpleString *workbuf = MboxMail::m_outdata;
	ASSERT(outbuf != inbuf);
	ASSERT(outbuf != workbuf);

	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);

	int bodyLength = body->m_contentLength;
	if ((body->m_contentOffset + body->m_contentLength) > m->m_length) {
		// something is not consistent
		bodyLength = m->m_length - body->m_contentOffset;
		if (bodyLength < 0)
			bodyLength = 0;
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

	return outbuf->Count();
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

		if ((!body->m_attachmentName.IsEmpty()) || 
			(body->m_contentDisposition.CompareNoCase("attachment") == 0)) 
		{
				continue;
		}
		if (textType == 0) {
			if (body->m_contentType.CompareNoCase("text/plain") != 0)
			{
				continue;
			}
		} 
		else if (textType == 1) {
			if (body->m_contentType.CompareNoCase("text/html") != 0)
			{
				continue;
			}
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
			if (bodyLength < 0)
				bodyLength = 0;
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


// TODO: Not used currently and code incomplete. might be used to dynamically create image files
int MboxMail::CreateImgAttachmentFiles(CFile &fpm, int mailPosition, SimpleString *outbuf)
{
	_int64 fileOffset;
	int bodyCnt = 0;

	// TODO" validate mailPosition
	MboxMail *m = MboxMail::s_mails[mailPosition];

	// We are using global buffers so check and assert if we collide. 
	SimpleString *inbuf = MboxMail::m_inbuf;
	//SimpleString *workbuf = MboxMail::m_workbuf;
	ASSERT(outbuf != inbuf);
	//ASSERT(outbuf != workbuf);

	outbuf->ClearAndResize(10000);
	inbuf->ClearAndResize(10000);
	//workbuf->ClearAndResize(10000);

	MailBodyContent *body;
	char *data;
	int dataLen;

	FileUtils::RemoveDirW(FileUtils::GetmboxviewTempPathW());

	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];

		if (body->m_attachmentName.IsEmpty()) {
			continue;
		}

		if (body->m_contentType.CompareNoCase("text/html") != 0) {
			; //continue;
		}

		bodyCnt++;

		int bodyLength = body->m_contentLength;
		if ((body->m_contentOffset + body->m_contentLength) > m->m_length) {
			// TODO: something is not consistent
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

			data = outbuf->Data();
			dataLen = outbuf->Count();
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

			data = outbuf->Data();
			dataLen = outbuf->Count();
		}
		else
		{
			// in case we have multiple bodies of the same type ?? not sure it is valid case/concern
			// asking for trouble ??
			data = bodyBegin;
			dataLen = bodyLength;
		}

		const char *fileName = (LPCSTR)body->m_attachmentName;
		CFile fp(fileName, CFile::modeWrite | CFile::modeCreate);
		fp.Write(data, dataLen);
		fp.Close();
	}

	return outbuf->Count();
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
	//pHdr->IsAttachement = pBP->IsAttachment();
	pHdr->IsAttachement = MboxCMimeHelper::IsAttachment(pBP);  // this differs from the attachment attribute as in Content-Disposition
	pHdr->IsMultiPart = pBP->IsMultiPart();
	pHdr->Content = FixIfNull((char*)pBP->GetContent());
	pHdr->ContentLength = pBP->GetContentLength();
	pHdr->Charset = pBP->GetCharset().c_str();

	pHdr->Description = FixIfNull(pBP->GetDescription());
	pHdr->Disposition = FixIfNull(pBP->GetDisposition());
	pHdr->TransferEncoding = FixIfNull(pBP->GetTransferEncoding());

	pHdr->SubType = pBP->GetSubType().c_str();
	pHdr->MainType = pBP->GetMainType().c_str();
	pHdr->Boundary = pBP->GetBoundary().c_str();

	pHdr->ContentType = FixIfNull(pBP->GetContentType());

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
// This function is NOT  USED  at the moment  !!!!!
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

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(mailFile, driveName, directory, fileNameBase, fileNameExtention);

	CFile fp;
	CString csvFile = path + "\\" + fileNameBase + ".csv";

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

	exportHeaderFieldLabelsToCSVFile(csvConfig, fp);

	char buff[10000];  // TODO:
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
				page_code = TextUtilsEx::Str2PageCode(charset);
			}

			cstr = fname ? fname : "";
			if (cstr.CompareNoCase("from") == 0) {
				m->m_from = fval ? fval : "";
				m->m_from.Replace("\n", "");
				m->m_from.Replace("\r", "");
				m->m_from_charsetId = page_code;
			}
			else if (cstr.CompareNoCase("to") == 0) {
				m->m_to = fval ? fval : "";
				m->m_to.Replace("\n", "");
				m->m_to.Replace("\r", "");
				m->m_to_charsetId = page_code;
			}
			else if (cstr.CompareNoCase("subject") == 0) {
				m->m_subj = fval ? fval : "";
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
				if (DateParser::parseRFC822Date(fval, &tm)) 
				{
					if (DateParser::validateSystemtime(&tm)) 
					{
						MyCTime tt(tm);
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
			char datebuff[32];
			CString format = MboxMail::GetDateFormat(csvConfig.m_dateFormat);

			datebuff[0] = 0;
			if (m->m_timeDate >= 0)
			{
				MyCTime tt(m->m_timeDate);
				if (!csvConfig.m_bGMTTime) 
				{
					CString lDateTime = tt.FormatLocalTm(format);
					strcpy(datebuff, (LPCSTR)lDateTime);
				}
				else {
					CString lDateTime = tt.FormatGmtTm(format);
					strcpy(datebuff, (LPCSTR)lDateTime);
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
				BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
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
				BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
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
					BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
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
				BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf);
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

			// some emails don't have pageCode; unfortunately there is no way to know dflt code installed on the sender's PC
			//if (pageCode == 0)  pageCode = CP_UTF8;  // don't massage the pageCode to keep visibility

			if (MboxMail::m_outbuf->Count())
			{
				int needLength = MboxMail::m_outbuf->Count() * 2 + 1; // worst case scenario or get '"' count first
				inbuf->ClearAndResize(needLength);  // escapeSeparators terminates result with null

				if (MboxMail::m_outbuf->Data()[0] != '\n')
					inbuf->Append("\n");
				int cnt_sv = inbuf->Count();
				int retcnt = escapeSeparators(inbuf->Data(cnt_sv), MboxMail::m_outbuf->Data(), MboxMail::m_outbuf->Count(), '"');
				inbuf->SetCount(retcnt + cnt_sv);

				// some emails don't have pageCode; unfortunately there is no way to know dflt code installed on the sender's PC
				// therefore can't do conversion, it will fail
				if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
				{
					BOOL ret = TextUtilsEx::Str2CodePage(inbuf, pageCode, csvConfig.m_nCodePageId, MboxMail::m_outbuf, workbuf);
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
			// else TODO: get HTML block and do html2text, see print single mail to text function

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

			bool isAttachment = MboxCMimeHelper::IsAttachment(pBP);
			if (pBP->IsText() && !isAttachment)
			{
				if (strcmp(pBP->GetSubType().c_str(), "plain") == 0) {

					char* content = (char*)pBP->GetContent();
					int contentLength = pBP->GetContentLength();
					if (content)
						outbuf->Append(content, contentLength);
					else
						outbuf->Append("");

					std::string charset = pBP->GetCharset();
					UINT page_code = TextUtilsEx::Str2PageCode(charset.c_str());
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

void MboxMail::ReleaseResources()
{
	TextUtilsEx::delete_charset2Id();
	TextUtilsEx::delete_id2charset();
	delete MboxMail::m_inbuf;
	MboxMail::m_inbuf = 0;
	delete MboxMail::m_outbuf;
	MboxMail::m_outbuf = 0;
	delete MboxMail::m_workbuf;
	MboxMail::m_workbuf = 0;
	delete MboxMail::m_tmpbuf;
	MboxMail::m_tmpbuf = 0;
	delete MboxMail::m_largebuf;
	MboxMail::m_largebuf = 0;
	delete MboxMail::m_indata;
	MboxMail::m_indata = 0;
	delete MboxMail::m_outdata;
	MboxMail::m_outdata = 0;
	delete MailBody::m_mpool;
	MailBody::m_mpool = 0;
	delete m_pMessageIdTable;
	m_pMessageIdTable = 0;
	delete m_pMboxMailTable;
	m_pMboxMailTable = 0;

	MboxMail::m_HintConfig.SaveToRegistry();

	CString processpath = "";
	CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("processPath"), processpath);
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

// Debug support functions

int MboxMail::DumpMailStatsToFile(MailArray *mailsArray, int mailsArrayCount)
{
	MboxMail *m;

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

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(mailFile, driveName, directory, fileNameBase, fileNameExtention);

	CFile fp;
	CString statsFile = path + "\\" + fileNameBase + "_stats.txt";

	if (!fp.Open(statsFile, CFile::modeWrite | CFile::modeCreate)) {
		CString txt = _T("Could not create \"") + statsFile;
		txt += _T("\" file.\nMake sure file is not open on other applications.");
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CFile fpm;  
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite)) {
		return -1;
	}

	SimpleString *outbuf = MboxMail::m_outbuf;
	outbuf->ClearAndResize(10000);

	BOOL hasPlain;
	BOOL hasHtml;
	BOOL hasBody;
	BOOL hasCharset;


	CString text;

	for (int i = 0; i < mailsArray->GetSize(); i++)
	{
		m = (*mailsArray)[i];

		outbuf->Clear();

		hasPlain = FALSE;
		hasHtml = FALSE;
		hasBody = FALSE;
		hasCharset = FALSE;

		UINT pageCode = 0;
		int textType = 0; //PLAIN
		int textlen = GetMailBody_mboxview(fpm, i, outbuf, pageCode, textType);  // fast
		if (textlen != outbuf->Count())
			int deb = 1;

		if (outbuf->Count())
		{
			hasPlain = TRUE;
		}

		outbuf->Clear();

		pageCode = 0;
		textType = 1; //HTML
		textlen = GetMailBody_mboxview(fpm, i, MboxMail::m_outbuf, pageCode, textType);  // fast
		if (textlen != outbuf->Count())
			int deb = 1;

		if (outbuf->Count())
		{
			hasHtml = TRUE;

			if (TextUtilsEx::findNoCase(outbuf->Data(), outbuf->Count(), "<body", 5) >= 0)
				hasBody = TRUE;

			if (TextUtilsEx::findNoCase(outbuf->Data(), outbuf->Count(), "charset=", 5) >= 0)
				hasCharset = TRUE;
		}

		text.Empty();
		text.Format("INDX=%d PLAIN=%d HTML=%d BODY=%d CHARSET=%d\n",
			i, hasPlain, hasHtml, hasBody, hasCharset);

		fp.Write(text, text.GetLength());
	}
	fp.Close();
	return 1;
}

MailList::MailList(int nId)
{
	m_nId = nId;
	//s_mails;
	b_mails_which_sorted = 0;
	m_lastSel = -1;
	//m_bIsDirty;
}

MailList::MailList()
{
	m_nId = IDC_ARCHIVE_LIST;
	//s_mails;
	b_mails_which_sorted = 0;
	m_lastSel = -1;
	//m_bIsDirty;
}

MailList::~MailList()
{
	;
}


int MboxMail::MergeTwoMailLists(MailArray *mails1, MailArray *mails2, MailArray *merged_mails)
{
	MboxMail::SortByIndex(mails1, false);
	MboxMail::SortByIndex(mails2, false);

	int mails1Cnt = mails1->GetCount();
	int mails2Cnt = mails2->GetCount();

	merged_mails->SetSizeKeepData(mails1Cnt + mails2Cnt);

	int mails1Index = 0;
	int mails2Index = 0;
	int merged_mailsIndex = 0;

	MboxMail *m1, *m2;

	for (;;)
	{
		while ((mails1Index < mails1Cnt) && (mails2Index < mails2Cnt)) 
		{
			m1 = (*mails1)[mails1Index];
			m2 = (*mails2)[mails2Index];
			if (m1->m_index == m2->m_index) {
				(*merged_mails)[merged_mailsIndex++] = m1;
				mails1Index++;
				mails2Index++;
			}
			else
				break;

		}
		if ((mails1Index >= mails1Cnt) || (mails2Index >= mails2Cnt))
			break;

		m2 = (*mails2)[mails2Index];
		while (mails1Index < mails1Cnt) 
		{
			m1 = (*mails1)[mails1Index];
			if ((m1->m_index < m2->m_index)) {
				(*merged_mails)[merged_mailsIndex++] = m1;
				mails1Index++;
			}
			else
				break;
		}

		if (mails1Index >= mails1Cnt)
			break;

		m1 = (*mails1)[mails1Index];
		while (mails2Index < mails2Cnt)
		{
			m2 = (*mails2)[mails2Index];
			if ((m2->m_index < m1->m_index)) {
				(*merged_mails)[merged_mailsIndex++] = m2;
				mails2Index++;
			}
			else
				break;
		}

		if (mails2Index >= mails2Cnt)
			break;

	}

	if ((mails1Index < mails1Cnt) && (mails2Index < mails2Cnt))
		int deb = 1; // it should never be true

	while (mails1Index < mails1Cnt)
	{
		m1 = (*mails1)[mails1Index++];
		(*merged_mails)[merged_mailsIndex++] = m1;
	}

	while (mails2Index < mails2Cnt)
	{
		m2 = (*mails2)[mails2Index++];
		(*merged_mails)[merged_mailsIndex++] = m2;
	}

	merged_mails->SetSizeKeepData(merged_mailsIndex);

	return merged_mailsIndex;
}

BOOL MboxMail::VerifyMergeOfTwoMailLists(MailArray *mails1, MailArray *mails2, MailArray *merged_mails)
{
	typedef unordered_map<int, int> ArrayIndexMap;
	ArrayIndexMap indexMap;

	ArrayIndexMap::iterator it;

	BOOL ret = TRUE;
	int index;
	int ii = 0;
	for (ii = 0; ii < merged_mails->GetCount(); ii++)
	{
		index = (*merged_mails)[ii]->m_index;

		if (indexMap.find(index) == indexMap.end())
			indexMap.insert(ArrayIndexMap::value_type(index, index));
		else
		{
			ret = FALSE;
			int deb = 1; // duplicates
		}
	}

	for (ii = 0; ii < mails1->GetCount(); ii++)
	{
		index = (*mails1)[ii]->m_index;

		if (indexMap.find(index) == indexMap.end()) {
			ret = FALSE;
			int deb = 1; // missing
		}
	}

	for (ii = 0; ii < mails2->GetCount(); ii++)
	{
		index = (*mails2)[ii]->m_index;

		if (indexMap.find(index) == indexMap.end()) {
			ret = FALSE;
			int deb = 1; // missing
		}
	}
	if (ret == FALSE)
		int deb = 1;

	return ret;
}

BOOL MboxMail::Test_MergeTwoMailLists()
{
	MailArray mails1;
	MailArray mails2;
	MailArray merged_mails;

	int carray1[] = { 1,2,4,5,11,26,30,42,50 };
	int carray1Size = sizeof(carray1) / sizeof(int);
	int carray2[] = { 2,4,5,9,10,11,25,26,30,40,42,48,50 };
	int carray2Size = sizeof(carray2) / sizeof(int);

	int expectedCarray[] = { 1,2,4,5,9,10,11,25,26,30,40,42,48,50 };
	int expectedCarraySize = sizeof(expectedCarray) / sizeof(int);

	PopulateCArray(&mails1, carray1, carray1Size);
	PopulateCArray(&mails2, carray2, carray2Size);

	MergeTwoMailLists(&mails1, &mails2, &merged_mails);
	BOOL ret = VerifyMerge(&merged_mails, expectedCarray, expectedCarraySize);
	if (ret == FALSE)
		int deb = 1;

	return TRUE;
}

void MboxMail::PopulateCArray(MailArray *mails, int *carray, int carrayCnt)
{
	MboxMail *m;
	for (int i = 0; i < carrayCnt; i++) {
		m = new MboxMail;
		int cval = carray[i];
		m->m_index = cval;
		mails ->Add(m);
	}
}

// mails is merged array, carray constains list of indexes into root/full mail list
BOOL MboxMail::VerifyMerge(MailArray *mails, int *carray, int carrayCnt)
{
	MboxMail *m;
	if (carrayCnt != mails->GetCount())
		return FALSE;

	for (int i = 0; i < carrayCnt; i++) {
		m = (*mails)[i];
		int cval = carray[i];
		if (m->m_index != cval)
			return FALSE;
	}
	return TRUE;
}

BOOL MboxMail::IsAllMailsSelected() {
	return (nWhichMailList == IDC_ARCHIVE_LIST);
}
BOOL MboxMail::IsFindMailsSelected() {
	return (nWhichMailList == IDC_FIND_LIST);
}
BOOL MboxMail::IsUserMailsSelected() {
	return (nWhichMailList == IDC_EDIT_LIST);
}
BOOL MboxMail::IsFolderMailsSelected() {
	return (nWhichMailList == IDC_FOLDER_LIST);
}

int MboxMail::AllMailsSelectedId() {
	return (IDC_ARCHIVE_LIST);
}
int MboxMail::FindMailsSelectedId() {
	return (IDC_FIND_LIST);
}
int MboxMail::UserMailsSelectedId() {
	return (IDC_EDIT_LIST);
}
int MboxMail::FolderMailsSelectedId() {
	return (IDC_FOLDER_LIST);
}

//void *bsearch(const void *key, const void *base, size_t num, size_t width, int(__cdecl *compare) (const void *key, const void *datum));

// template for search for X in sorted array
int Binary_search(int arr[], int X, int low, int high)
{

	while (low <= high) // till low is less than high i.e. there is atleast one integer in the considered part of array
	{

		int mid = low + (high - low) / 2; //compute the middle index

		if (arr[mid] == X) //if equal then return
			return mid;

		else if (arr[mid] < X) //if smaller then increase the lower limit
			low = mid + 1;

		else //if larger then decrease the upper limit
			high = mid - 1;
	}

	return -1;
}

// class MyMailArray
void MyMailArray::SetSizeKeepData(INT_PTR nNewSize, INT_PTR nGrowBy)
{
	if (nNewSize == 0)
		m_nSize = 0;
	else
		SetSize(nNewSize, nGrowBy);
}
void MyMailArray::CopyKeepData(const MyMailArray& src)
{
	if (src.GetSize() == 0) 
		m_nSize = 0;
	else
		Copy(src);
}

int MboxMail::MakeFileName(MboxMail *m, struct NamePatternParams *namePatternParams, CString &fileName)
{
	SimpleString name;
	SimpleString addr;
	SimpleString addrTo;
	//CString tmpAddr;
	CString fromAddr;
	CString toAddr;
	CString subjAddr;
	CString subj;
	CString tmp;
	int pos;
	BOOL allowUnderscore = FALSE;
	BOOL separatorNeeded = FALSE;
	char sepchar = '-';

	if (m->m_timeDate < 0)
	{
		CString m_strDate;
		if (namePatternParams->m_bDate)
		{
			if (namePatternParams->m_bTime)
				//m_strDate = ltt.Format("%Y%m%d-%H%M%S");
				m_strDate = "000000000-000000";
			else
				//m_strDate = ltt.Format("%H%M%S");
				m_strDate = "000000";

			fileName.Append(m_strDate);
		}
		else if (namePatternParams->m_bTime)
		{
			m_strDate = "000000";
			fileName.Append(m_strDate);
		}
		separatorNeeded = TRUE;
	}
	else if (namePatternParams->m_bDate) 
	{
		MyCTime tt(m->m_timeDate);

		CString m_strDate;

		CString format;
		if (namePatternParams->m_bTime)
		{
			format = "%Y%m%d-%H%M%S";
		}
		else
		{
			format = "%Y%m%d";
		}

		m_strDate = tt.FormatLocalTm(format);
		fileName.Append(m_strDate);

		separatorNeeded = TRUE;
	}
	else if (namePatternParams->m_bTime) 
	{
		MyCTime tt(m->m_timeDate);
		CString format = "%H%M%S";
		CString m_strDate;
		m_strDate = tt.FormatLocalTm(format);

		fileName.Append(m_strDate);

		separatorNeeded = TRUE;
	}

	if ((fileName.GetLength() < namePatternParams->m_nFileNameFormatSizeLimit) && namePatternParams->m_bFrom)
	{
		if (separatorNeeded) {
			fileName.AppendChar(sepchar);
			int deb = 1;
		}

		int fromlen = m->m_from.GetLength();
		name.ClearAndResize(fromlen);
		addr.ClearAndResize(fromlen);
		MboxMail::splitMailAddress(m->m_from, fromlen, &name, &addr);
		int pos = addr.Find(0, '@');
		if (pos >= 0)
			fromAddr.Append(addr.Data(), pos);
		else
			fromAddr.Append(addr.Data(), addr.Count());
		fromAddr.Trim(" \t\"<>:;,");  // may need to remove more not aphanumerics
		fileName.Append(fromAddr);

		separatorNeeded = TRUE;
	}
	if ((fileName.GetLength() < namePatternParams->m_nFileNameFormatSizeLimit) &&  namePatternParams->m_bTo)
	{
		if (separatorNeeded) {
			fileName.AppendChar(sepchar);
			int deb = 1;
		}

		int maxNumbOfAddr = 1;
		NListView::TrimToAddr(&m->m_to, toAddr, maxNumbOfAddr);

		pos = toAddr.Find("@");

		if (pos >= 0)
			addrTo.Append(toAddr, pos);
		else
			addrTo.Append(toAddr, toAddr.GetLength());

		toAddr.Empty();
		toAddr.Append(addrTo.Data(), addrTo.Count());

		toAddr.Trim(" \t\"<>:;,"); // may need to remove more not aphanumerics

		//fileName.Append(addrTo.Data(), addrTo.Count());
		fileName.Append(toAddr);

		separatorNeeded = TRUE;

		int deb = 1;

	}
	
	int subjLengthAllowed = namePatternParams->m_nFileNameFormatSizeLimit - fileName.GetLength();

	CString crc32;
	//crc32.Format("%X", m->m_crc32);
	crc32.Format("%d", m->m_index);

	subjLengthAllowed = subjLengthAllowed - (crc32.GetLength() + 3 + 6);  // 3 is for 2 * dash plus at least one character for subject plus 6 for file extension
	if (subjLengthAllowed < 0)
		subjLengthAllowed = 0;

	int subjLength = m->m_subj.GetLength();
	if (subjLengthAllowed > subjLength)
		subjLengthAllowed = subjLength;

	if ((subjLengthAllowed > 0) && namePatternParams->m_bSubject)
	{
		if (separatorNeeded) {
			fileName.AppendChar(sepchar);
			int deb = 1;
		}
		// TODO: create function
		int outCnt = 0;
		int ignoreCnt = 0;
		int i;
		for (i = 0; i < subjLength; i++)
		{
			if (outCnt >= subjLengthAllowed)
				break;

			char c = m->m_subj.GetAt(i);
			unsigned char cc = c;

			if ((cc < 127) && isalnum(c)) {
				fileName.AppendChar(c);
				allowUnderscore = TRUE;
				outCnt++;
			}
			else if (allowUnderscore)
			{
				//if (outCnt < (subjLengthAllowed - 1)) {
					fileName.AppendChar('_');
					allowUnderscore = FALSE;
					outCnt++;
				//}
				//else break;
			}
			else
				ignoreCnt++;
		}
		separatorNeeded = TRUE;
	}

	fileName.TrimRight("_");
	if (separatorNeeded) {
		fileName.AppendChar(sepchar);
		int deb = 1;
	}

	int fileLength = fileName.GetLength();
	if ((fileLength + crc32.GetLength()) > namePatternParams->m_nFileNameFormatSizeLimit) 
	{
		int deb = 1;
	}

	fileName.Append((LPCSTR)crc32, crc32.GetLength());

	if (fileName.GetLength() > namePatternParams->m_nFileNameFormatSizeLimit)
	{
		fileName.Truncate(namePatternParams->m_nFileNameFormatSizeLimit);
	}

	int fileNameLen = fileName.GetLength();
	return 1;

}

int MboxMail::MakeFileName(MboxMail *m, struct NameTemplateCnf &nameTemplateCnf, CString &fileName, int maxFileNameLength)
{
	SimpleString name;
	SimpleString addr;
	CString cname;
	CString caddr;

	CString username;
	CString domain;

	CString fromAddr;
	CString fromName;
	SimpleString addrTo;

	CString toAddr;
	CString toName;
	SimpleString nameTo;

	CString subjAddr;
	CString subj;
	CString tmp;
	CString FieldText;
	BOOL allowUnderscore = FALSE;
	char sepchar = '-';
	int needLength;

	// make room for extension - actually caller should do this
	maxFileNameLength -= 6; 
	int maxNonAdjustedFileNameLength = maxFileNameLength;

	CArray<CString> labelArray;
	BOOL ret = MboxMail::ParseTemplateFormat(nameTemplateCnf.m_TemplateFormat, labelArray);

	// make room for UNIQUE_ID
	CString idLabel = "%UNIQUE_ID%";
	CString uID;
	uID.Format("%d", m->m_index);
	if (TemplateFormatHasLabel(idLabel, labelArray) == TRUE)
	{
		maxFileNameLength -= uID.GetLength() + 6;
	}
	if (maxFileNameLength < 0)
		maxFileNameLength = 0;

	int i = 0;
	for (i = 0; i < labelArray.GetCount(); i++)
	{
		CString label = labelArray[i];
		if (labelArray[i].CompareNoCase("%DATE_TIME%") == 0)
		{
			CString strDate;
			if (m->m_timeDate < 0)
			{
				strDate = "0000-00-00";
			}
			else
			{
				MyCTime tm(m->m_timeDate);
				strDate = tm.FormatLocalTm(nameTemplateCnf.m_DateFormat);

			}
			needLength = fileName.GetLength() + strDate.GetLength();
			if (needLength < maxFileNameLength)
				fileName.Append(strDate);
			int deb = 1;
		}
		else if (labelArray[i].CompareNoCase("%FROM_NAME%") == 0)
		{
			CString addrLabel = "%FROM_ADDR%";

			int fromlen = m->m_from.GetLength();
			name.ClearAndResize(fromlen);
			addr.ClearAndResize(fromlen);
			username.Empty();
			domain.Empty();

			MboxMail::splitMailAddress(m->m_from, fromlen, &name, &addr);
			FileUtils::MakeValidFileName(name, nameTemplateCnf.m_bReplaceWhiteWithUnderscore);
			if ((name.Count() > 0) && !((name.Count() == 1) && (name.GetAt(0) == '_')))
			{
				fromName.Empty();
				fromName.Append(name.Data(), name.Count());
				fromName.Trim(" \t\"<>:;,()");  // may need to remove more not only aphanumerics
				needLength = fileName.GetLength() + fromName.GetLength();
				if (needLength < maxFileNameLength)
					fileName.Append(fromName);
				;
			}
			else if (TemplateFormatHasLabel(addrLabel, labelArray) == FALSE)
			{
				int pos = addr.Find(0, '@');
				if (pos >= 0)
				{
					username.Append(addr.Data(), pos);
					domain.Append(addr.Data(pos + 1), addr.Count() - (pos + 1));
				}
				else
					username.Append(addr.Data(), addr.Count());

				fromAddr.Empty();
				if (nameTemplateCnf.m_bFromUsername && nameTemplateCnf.m_bFromDomain)
					fromAddr.Append(addr.Data(), addr.Count());
				else if (nameTemplateCnf.m_bFromUsername)
					fromAddr.Append(username);
				else
					fromAddr.Append(domain);

				fromAddr.Trim(" \t\"<>:;,()");  // may need to remove more not only aphanumerics
				needLength = fileName.GetLength() + fromAddr.GetLength();
				if (needLength < maxFileNameLength)
					fileName.Append(fromAddr);
			}
		}
		else if (labelArray[i].CompareNoCase("%FROM_ADDR%") == 0)
		{
			int fromlen = m->m_from.GetLength();
			name.ClearAndResize(fromlen);
			addr.ClearAndResize(fromlen);
			username.Empty();
			domain.Empty();

			MboxMail::splitMailAddress(m->m_from, fromlen, &name, &addr);
			int pos = addr.Find(0, '@');
			if (pos >= 0)
			{
				username.Append(addr.Data(), pos);
				domain.Append(addr.Data(pos+1), addr.Count() - (pos+1));
			}
			else
				username.Append(addr.Data(), addr.Count());

			fromAddr.Empty();
			if (nameTemplateCnf.m_bFromUsername && nameTemplateCnf.m_bFromDomain)
				fromAddr.Append(addr.Data(), addr.Count());
			else if (nameTemplateCnf.m_bFromUsername)
				fromAddr.Append(username);
			else
				fromAddr.Append(domain);

			fromAddr.Trim(" \t\"<>:;,");  // may need to remove more not only aphanumerics
			needLength = fileName.GetLength() + fromAddr.GetLength();
			if (needLength < maxFileNameLength)
				fileName.Append(fromAddr);

			int deb = 1;
		}

		//
		else if (labelArray[i].CompareNoCase("%TO_NAME%") == 0)
		{
			CString addrLabel = "%TO_ADDR%";

			int tolen = m->m_to.GetLength();
			name.ClearAndResize(tolen);
			addr.ClearAndResize(tolen);
			username.Empty();
			domain.Empty();

			MboxMail::splitMailAddress(m->m_to, tolen, &name, &addr);
			FileUtils::MakeValidFileName(name, nameTemplateCnf.m_bReplaceWhiteWithUnderscore);
			if ((name.Count() > 0) && !((name.Count() == 1) && (name.GetAt(0) == '_')))
			{
				toName.Empty();
				toName.Append(name.Data(), name.Count());
				toName.Trim(" \t\"<>:;,()");  // may need to remove more not only aphanumerics
				needLength = fileName.GetLength() + toName.GetLength();
				if (needLength < maxFileNameLength)
					fileName.Append(toName);
				;
			}
			else if (TemplateFormatHasLabel(addrLabel, labelArray) == FALSE)
			{
				int pos = addr.Find(0, '@');
				if (pos >= 0)
				{
					username.Append(addr.Data(), pos);
					domain.Append(addr.Data(pos + 1), addr.Count() - (pos + 1));
				}
				else
					username.Append(addr.Data(), addr.Count());

				toAddr.Empty();
				if (nameTemplateCnf.m_bToUsername && nameTemplateCnf.m_bToDomain)
					toAddr.Append(addr.Data(), addr.Count());
				else if (nameTemplateCnf.m_bToUsername)
					toAddr.Append(username);
				else
					toAddr.Append(domain);

				toAddr.Trim(" \t\"<>:;,()");  // may need to remove more not only aphanumerics
				needLength = fileName.GetLength() + toAddr.GetLength();
				if (needLength < maxFileNameLength)
					fileName.Append(toAddr);
			}
		}
		else if (labelArray[i].CompareNoCase("%TO_ADDR%") == 0)
		{
			int tolen = m->m_to.GetLength();
			name.ClearAndResize(tolen);
			addr.ClearAndResize(tolen);
			username.Empty();
			domain.Empty();

			MboxMail::splitMailAddress(m->m_to, tolen, &name, &addr);
			int pos = addr.Find(0, '@');
			if (pos >= 0)
			{
				username.Append(addr.Data(), pos);
				domain.Append(addr.Data(pos + 1), addr.Count() - (pos + 1));
			}
			else
				username.Append(addr.Data(), addr.Count());

			toAddr.Empty();
			if (nameTemplateCnf.m_bToUsername && nameTemplateCnf.m_bToDomain)
				toAddr.Append(addr.Data(), addr.Count());
			else if (nameTemplateCnf.m_bToUsername)
				toAddr.Append(username);
			else
				toAddr.Append(domain);

			toAddr.Trim(" \t\"<>:;,");  // may need to remove more not only aphanumerics
			needLength = fileName.GetLength() + toAddr.GetLength();
			if (needLength < maxFileNameLength)
				fileName.Append(toAddr);

			int deb = 1;
		}
		else if (labelArray[i].CompareNoCase("%SUBJECT%") == 0)
		{
			int lengthAvailable = maxFileNameLength - fileName.GetLength();
			if (lengthAvailable < 0)
				lengthAvailable = 0;
			int subjectLength = m->m_subj.GetLength();
			if (subjectLength > lengthAvailable)
				subjectLength = lengthAvailable;

			fileName.Append(m->m_subj, subjectLength);
		}
		else if (labelArray[i].CompareNoCase("%UNIQUE_ID%") == 0)
		{
			needLength = fileName.GetLength() + uID.GetLength();
			if (needLength < maxNonAdjustedFileNameLength)
				fileName.Append(uID);
		}
		else
		{
			needLength = fileName.GetLength() + labelArray[i].GetLength();
			if (needLength < maxFileNameLength)
			{
				fileName.Append(labelArray[i]);
			}
		}
	}
	fileName.TrimLeft();

	if (fileName.GetLength() > maxFileNameLength)
	{
		fileName.Truncate(maxFileNameLength);
	}

	FileUtils::MakeValidFileName(fileName, nameTemplateCnf.m_bReplaceWhiteWithUnderscore);

	int fileNameLen = fileName.GetLength();
	return 1;

}

static CArray<CString> FldLabelTable;
static BOOL fldLabelsTableSet = FALSE;

CString MboxMail::MatchFldLabel(char *p, char *e)
{
	CString empty;
	if (fldLabelsTableSet == FALSE)
	{
		FldLabelTable.Add(_T("%DATE_TIME%"));
		FldLabelTable.Add(_T("%FROM_NAME%"));
		FldLabelTable.Add(_T("%FROM_ADDR%"));
		FldLabelTable.Add(_T("%TO_NAME%"));
		FldLabelTable.Add(_T("%TO_ADDR%"));
		FldLabelTable.Add(_T("%SUBJECT%"));
		FldLabelTable.Add(_T("%UNIQUE_ID%"));

		fldLabelsTableSet = TRUE;
	}


	int left = e - p;
	int fldCnt = FldLabelTable.GetCount();
	int labelLength;
	char *labelData;
	int i;
	for (i = 0; i < fldCnt; i++)
	{
		labelLength = FldLabelTable[i].GetLength();
		labelData = (char*)((LPCSTR)FldLabelTable[i]);
		if ((FldLabelTable[i].GetLength() <= left) && (strncmp(p, labelData, labelLength) == 0))
			return FldLabelTable[i];
	}
	return empty;
}

BOOL MboxMail::ParseTemplateFormat(CString &templateFormat, CArray<CString> &labelArray)
{
	char *p = (char*)((LPCSTR)templateFormat);
	char *e = p + templateFormat.GetLength();
	CString label;
	CString plainText;
	char c;
	while (p < e)
	{
		c = *p;
		if (c == '%')
		{
			label = MatchFldLabel(p, e);
		}
		if (!label.IsEmpty())
		{
			if (!plainText.IsEmpty())
			{
				labelArray.Add(plainText);
				plainText.Empty();
			}

			labelArray.Add(label);
			p += label.GetLength();
			label.Empty();
		}
		else
		{
			plainText.AppendChar(c);
			p++;
		}
	}
	if (!plainText.IsEmpty())
	{
		labelArray.Add(plainText);
		plainText.Empty();
	}
	if (p != e)
		int deb = 1;

#if 0
	TRACE(_T("LABEL ARRAY:\n\n"));
	for (int ii = 0; ii < labelArray.GetCount(); ii++)
	{
		TRACE(_T("\"%s\"\n"), labelArray[ii]);
	}

#endif

	return TRUE;
}

BOOL MboxMail::TemplateFormatHasLabel(CString &label, CArray<CString> &labelArray)
{
	for (int ii = 0; ii < labelArray.GetCount(); ii++)
	{
		if (labelArray[ii].CompareNoCase(label) == 0)
			return TRUE;
	}
	return FALSE;
}

int MboxMail::RemoveDuplicateMails()
{
	MboxMail *m;
	MboxMail *m_found;

	// !! s_mails is already populated from mboxview file

	// Remove duplicate mails
	if (s_mails.GetCount() > s_mails_edit.GetCount())
		s_mails_edit.SetSizeKeepData(s_mails.GetCount());

	if ((s_mails.GetCount()/2+1) > s_mails_find.GetCount())
		s_mails_find.SetSizeKeepData(s_mails.GetCount()/2 + 1);

	int to_dup_i = 0;
	int to_i = 0;
	for (int i = 0; i < s_mails.GetSize(); i++)
	{
		m = s_mails[i];

		if (m->m_messageId.GetLength())
		{
			m_found = getMboxMail(&m->m_messageId);
			if (m_found == 0)
			{
				insertMboxMail(&m->m_messageId, m);

				s_mails_edit[to_i] = s_mails[i];
				to_i++;
				m->m_duplicateId = false;
			}
			else 
			{
				if ((m_found->m_timeDate != m->m_timeDate) ||
					(m_found->m_from != m->m_from) ||
					(m_found->m_to != m->m_to) )
				{
					s_mails_edit[to_i] = s_mails[i];
					to_i++;
					m->m_duplicateId = false;
				}
				else
				{
					s_mails_find[to_dup_i] = s_mails[i];
					to_dup_i++;
					m->m_duplicateId = true;
				}
			}
		}
		else 
		{
			s_mails_edit[to_i] = s_mails[i];
			to_i++;
			int deb = 1;
		}
	}
	s_mails_edit.SetSizeKeepData(to_i);

	s_mails_find.SetSizeKeepData(to_dup_i);
	MboxMail::m_findMails.m_lastSel = -1;
	MboxMail::m_findMails.b_mails_which_sorted = MboxMail::b_mails_which_sorted;

	int dupCnt = s_mails.GetSize() - s_mails_edit.GetSize();

	s_mails.CopyKeepData(s_mails_edit);

	clearMboxMailTable();

	return dupCnt;
}


//
MboxMail* MboxMail::getMboxMail(CString *key)
{
	MboxMailTableType::iterator it;
	if (m_pMboxMailTable == 0)
		createMboxMailTable(50000);

	it = m_pMboxMailTable->find(key);
	if (it != m_pMboxMailTable->cend()) {
		return it->second;
	}
	else
		return 0;
}

bool MboxMail::insertMboxMail(CString *key, MboxMail *mbox)
{
	CString *mapKey = new CString;
	*mapKey = *key;
	std::pair<MboxMailTableType::iterator, bool> result = m_pMboxMailTable->insert(MboxMailTableType::value_type(mapKey, mbox));
	return result.second;
}

UINT MboxMail::createMboxMailTable(UINT count)
{
	m_pMboxMailTable = new MboxMailTableType;
	m_pMboxMailTable->reserve(count);
	return count;
}

void MboxMail::clearMboxMailTable()
{
	MboxMailTableType::iterator it;

	if (m_pMboxMailTable) {
		for (it = m_pMboxMailTable->begin(); it != m_pMboxMailTable->end(); it++) {
			delete it->first;
			//it->first = 0;
		}
		m_pMboxMailTable->clear();
	}
}

BOOL MboxMail::CreatePrintCachePath(CString &rootPrintSubFolder, CString &targetPrintSubFolder, CString &prtCachePath, CString &errorText)
{
	prtCachePath.Empty();

	CString mailArchiveFilePath = MboxMail::s_path;

	if (s_path.IsEmpty()) {
		errorText = _T("Please open mail archive file first.");
		return FALSE;
	}

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;
	FileUtils::SplitFilePath(mailArchiveFilePath, driveName, directory, fileNameBase, fileNameExtention);

	CString mailArchiveFileName;
	FileUtils::CPathStripPath((char*)(LPCSTR)MboxMail::s_path, mailArchiveFileName);

	int position = mailArchiveFileName.ReverseFind('.');
	CString baseFileArchiveName;
	if (position >= 0)
		baseFileArchiveName = mailArchiveFileName.Mid(0, position);
	else
		baseFileArchiveName = mailArchiveFileName;

	BOOL createDirOk = TRUE;
	CString printCachePath;
	BOOL ret = FileUtils::CPathGetPath(MboxMail::s_path, printCachePath);
	if (!rootPrintSubFolder.IsEmpty())
	{
		printCachePath.Append("\\");
		printCachePath.Append(rootPrintSubFolder);

		createDirOk = TRUE;
		if (!FileUtils::PathDirExists(printCachePath)) {
			createDirOk = CreateDirectory(printCachePath, NULL);
		}

		if (!createDirOk) {
			errorText = _T("Could not create \"") + printCachePath;
			errorText += _T("\" folder for print destination.\nResolve the problem and try again.");
			return FALSE;
		}
	}

	printCachePath.Append("\\");
	printCachePath.Append(baseFileArchiveName);

	createDirOk = TRUE;
	if (!FileUtils::PathDirExists(printCachePath))
		createDirOk = CreateDirectory(printCachePath, NULL);

	if (!createDirOk) {
		errorText = _T("Could not create \"") + printCachePath;
		errorText += _T("\" folder for print destination.\nResolve the problem and try again.");
		return FALSE;
	}
	if (!targetPrintSubFolder.IsEmpty())
	{
		printCachePath.Append("\\");
		printCachePath.Append(targetPrintSubFolder);

		createDirOk = TRUE;
		if (!FileUtils::PathDirExists(printCachePath)) {
			createDirOk = CreateDirectory(printCachePath, NULL);
		}

		if (!createDirOk) {
			errorText = _T("Could not create \"") + printCachePath;
			errorText += _T("\" folder for print destination.\nResolve the problem and try again.");
			return FALSE;
		}
	}

	prtCachePath.Append(printCachePath);

	return TRUE;
}

bool MboxMail::GetPrintCachePath(CString &rootPrintSubFolder, CString &targetPrintSubFolder, CString &prtCachePath, CString &errorText)
{
	prtCachePath.Empty();

	CString mailArchiveFilePath = MboxMail::s_path;

	if (s_path.IsEmpty()) {
		errorText = _T("Please open mail archive file first.");
		return FALSE;
	}

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;
	FileUtils::SplitFilePath(mailArchiveFilePath, driveName, directory, fileNameBase, fileNameExtention);

	CString mailArchiveFileName;
	FileUtils::CPathStripPath((char*)(LPCSTR)MboxMail::s_path, mailArchiveFileName);
	int position = mailArchiveFileName.ReverseFind('.');
	CString baseFileArchiveName;
	if (position >= 0)
		baseFileArchiveName = mailArchiveFileName.Mid(0, position);
	else
		baseFileArchiveName = mailArchiveFileName;

	CString printCachePath;
	BOOL ret = FileUtils::CPathGetPath(MboxMail::s_path, printCachePath);
	if (!rootPrintSubFolder.IsEmpty())
	{
		printCachePath.Append("\\");
		printCachePath.Append(rootPrintSubFolder);
	}

	printCachePath.Append("\\");
	printCachePath.Append(baseFileArchiveName);

	prtCachePath.Append(printCachePath);

	return TRUE;
}

bool MboxMail::GetArchiveSpecificCachePath(CString &path, CString &rootPrintSubFolder, CString &targetPrintSubFolder, CString &prtCachePath, CString &errorText)
{
	prtCachePath.Empty();

	CString mailArchiveFilePath = path;

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;
	FileUtils::SplitFilePath(mailArchiveFilePath, driveName, directory, fileNameBase, fileNameExtention);

	CString mailArchiveFileName;
	FileUtils::CPathStripPath((char*)(LPCSTR)path, mailArchiveFileName);
	int position = mailArchiveFileName.ReverseFind('.');
	CString baseFileArchiveName;
	if (position >= 0)
		baseFileArchiveName = mailArchiveFileName.Mid(0, position);
	else
		baseFileArchiveName = mailArchiveFileName;

	CString printCachePath;
	BOOL ret = FileUtils::CPathGetPath(path, printCachePath);
	if (!rootPrintSubFolder.IsEmpty())
	{
		printCachePath.Append("\\");
		printCachePath.Append(rootPrintSubFolder);
	}

	printCachePath.Append("\\");
	printCachePath.Append(baseFileArchiveName);

	prtCachePath.Append(printCachePath);

	return TRUE;
}

int MboxMail::MakeFileNameFromMailArchiveName(int fileType, CString &fileName, CString &targetPrintSubFolder, bool &fileExists, CString &errorText)
{
	CString mailArchiveFileName = MboxMail::s_path;

	if (mailArchiveFileName.IsEmpty()) {
		errorText = _T("Please open mail archive file first.");
		return -1;
	}

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;
	FileUtils::SplitFilePath(mailArchiveFileName, driveName, directory, fileNameBase, fileNameExtention);

	CString printCachePath;
	CString rootPrintSubFolder = "PrintCache";

	BOOL retval = CreatePrintCachePath(rootPrintSubFolder, targetPrintSubFolder, printCachePath, errorText);
	if (retval == FALSE) {
		return -1;
	}

	if (fileType == 0)
		fileName = printCachePath + "\\" + fileNameBase + ".txt";
	else if (fileType == 1)
		fileName = printCachePath + "\\" + fileNameBase + ".htm";
	else if (fileType == 2)
		fileName = printCachePath + "\\" + fileNameBase + ".pdf";
	else if (fileType == 3)
		fileName = printCachePath + "\\" + fileNameBase + ".csv";

	fileExists = FileUtils::PathFileExist(fileName);

	return 1;
}

int MboxMail::MakeFileNameFromMailHeader(int mailIndex, int fileType, CString &fileName, CString &targetPrintSubFolder, bool &fileExists, CString &errorText)
{
	CString mailArchiveFileName = MboxMail::s_path;

	if (mailArchiveFileName.IsEmpty()) {
		errorText = _T("Please open mail archive file first.");
		return -1;
	}

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (!pFrame) {
		errorText = _T("Internal error. Try again.");
		return -1;
	}

	CString printCachePath;
	CString rootPrintSubFolder = "PrintCache";
	BOOL retval = CreatePrintCachePath(rootPrintSubFolder, targetPrintSubFolder, printCachePath, errorText);
	if (retval == FALSE) {
		return -1;
	}

#if 0
	// Added fudge factor in print Config
	int nFileNameFormatSizeLimit = 260 - (printCachePath.GetLength() + 2);
	if (nFileNameFormatSizeLimit < pFrame->m_NamePatternParams.m_nFileNameFormatSizeLimit)
		pFrame->m_NamePatternParams.m_nFileNameFormatSizeLimit = nFileNameFormatSizeLimit;  // TODO: need better fix in Print Config
#endif

	CString mailFileNameBase;
#if 0
	// Just for investigation
	mailFileNameBase.Preallocate(500);
	int allocLen = mailFileNameBase.GetAllocLength();  // returns 500
	mailFileNameBase.Append("zzz");
	int strLen = mailFileNameBase.GetLength();  // returns 3
	// Empty() seems to release the buffer !!!
	// We could keep one character in the buffer but the next operation must be copy/replace. Not really friendly.
	// mailFileNameBase.Empty();  
	allocLen = mailFileNameBase.GetAllocLength(); // returns 0 if Empty() is called
	//
#endif


	MboxMail *m = MboxMail::s_mails[mailIndex];

	if (!pFrame->m_NamePatternParams.m_bCustomFormat)
		MakeFileName(m, &pFrame->m_NamePatternParams, mailFileNameBase);
	else
		MakeFileName(m, pFrame->m_NamePatternParams.m_nameTemplateCnf, mailFileNameBase, pFrame->m_NamePatternParams.m_nFileNameFormatSizeLimit);

	if (fileType == 0)
		fileName = printCachePath + "\\" + mailFileNameBase + ".txt";
	else if (fileType == 1)
		fileName = printCachePath + "\\" + mailFileNameBase + ".htm";
	else if (fileType == 2)
		fileName = printCachePath + "\\" + mailFileNameBase + ".pdf";
	else if (fileType == 3)
		fileName = printCachePath + "\\" + mailFileNameBase + ".csv";

	fileExists = FileUtils::PathFileExist(fileName);

	return 1;
}

int MboxMail::PrintMailRangeToSingleTextFile(TEXTFILE_CONFIG &textConfig, CString &textFileName, int firstMail, int lastMail, int textType, CString &targetPrintSubFolderName, CString errorText)
{
	CFile fp;
	CString textFile;
	bool fileExists = false;
	int ret = 1;

	if (firstMail == lastMail)
	{
		ret = MakeFileNameFromMailHeader(firstMail, textType, textFile, targetPrintSubFolderName, fileExists, errorText);
		if (fileExists)
			int deb = 1;
	}
	else
	{
		ret = MakeFileNameFromMailArchiveName(textType, textFile, targetPrintSubFolderName, fileExists, errorText);
		if (fileExists)
			int deb = 1;
	}

	if (ret < 0) 
		return -1;

	textFileName = textFile;
	// Useful durimg debug
	int textFileLength = textFile.GetLength();
	int maxPath = _MAX_PATH;

	CFileException ex;
	//if (!fp.Open(textFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone), &ex)
	if (!fp.Open(textFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone))
	{
		DWORD er = GetLastError();
		//TCHAR szError[1024];
		//ex.GetErrorMessage(szError, 1024);
		//CFileStatus rStatus;
		//BOOL ret = fp.GetStatus(rStatus);
		errorText = _T("Could not create \"") + textFile;
		errorText += _T("\" file.\nMake sure file is not open on other applications.");
		//return -1;
	}

	CFile fpm;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite)) {
		errorText = _T("Could not open mail archive \"") + s_path;
		fp.Close();
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
		else {
			printSingleMailToHtmlFile(fp, i, fpm, textConfig, i == firstMail);
		}
	}

	fp.Close();
	fpm.Close();

	return 1;
}

int MboxMail::PrintMailRangeToSingleTextFile_WorkerThread(TEXTFILE_CONFIG &textConfig, CString &textFileName, int firstMail, int lastMail, int textType, CString errorText)
{
	BOOL progressBar = TRUE;
	CFile fp;
	CString textFile;
	bool fileExists = false;
	int ret = 1;

	CString targetPrintSubFolder;
	if (firstMail == lastMail)
		ret = MakeFileNameFromMailHeader(firstMail, textType, textFile, targetPrintSubFolder, fileExists, errorText);
	else
		ret = MakeFileNameFromMailArchiveName(textType, textFile, targetPrintSubFolder, fileExists, errorText);

	if (ret < 0)
		return -1;

	textFileName = textFile;
	// Useful durimg debug
	int textFileLength = textFile.GetLength();
	int maxPath = _MAX_PATH;

	if (!fp.Open(textFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone)) {
		errorText = _T("Could not create \"") + textFile;
		errorText += _T("\" file.\nMake sure file is not open on other applications.");
		return -1;
	}

	CFile fpm;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite)) {
		errorText = _T("Could not open mail archive \"") + s_path;
		fp.Close();
		return -1;
	}

	if (progressBar && MboxMail::pCUPDUPData)
	{
		if (textType == 0)
			MboxMail::pCUPDUPData->SetProgress(_T("Printing mails to TEXT file ..."), 0);
		else
			MboxMail::pCUPDUPData->SetProgress(_T("Printing mails to HTML file ..."), 0);
	}

	double delta = (double)(lastMail - firstMail);
	if (delta <= 0) delta = 1;
	double step = delta / 100;
	double curstep = 1;
	double newstep = 0;

	if (progressBar && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	if (textType == 0) {
		if (textConfig.m_nCodePageId == CP_UTF8) {
			const char *BOM_UTF8 = "\xEF\xBB\xBF";
			fp.Write(BOM_UTF8, 3);
		}
	}

	CString fileNum;

	int cnt = lastMail - firstMail;
	if (cnt <= 0)
		cnt = 0;

	for (int i = firstMail; i <= lastMail; i++)
	{
		if (textType == 0)
			printSingleMailToTextFile(fp, i, fpm, textConfig);
		else {
			printSingleMailToHtmlFile(fp, i, fpm, textConfig, i == firstMail);
		}

		if (progressBar && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				int deb = 1;
				break;
			}

			newstep = ((double)(i + 1 - 0)) / step + 1;
			if (newstep >= curstep)
			{
				if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
					int deb = 1;
					break;
				}

				if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress((UINT_PTR)(newstep));

				curstep = newstep;
			}

			int nFileNum = (i + 1);
			int stepSize = 100;
			if (textType == 0)
				stepSize = 1000;

			if ((nFileNum % stepSize) == 0) 
			{
				if (textType == 0)
					fileNum.Format(_T("Printing mails to single TEXT file ... %d of %d"), nFileNum, cnt);
				else
					fileNum.Format(_T("Printing mails to single HTML file ... %d of %d"), nFileNum, cnt);
				if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(newstep));
			}
		}
	}

	newstep = 100;
	if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress((UINT_PTR)(newstep));

	fp.Close();
	fpm.Close();

	return 1;
}

int MboxMail::PrintMailSelectedToSingleTextFile_WorkerThread(TEXTFILE_CONFIG &textConfig, CString &textFileName, MailIndexList *selectedMailIndexList, int textType, CString errorText)
{
	BOOL progressBar = TRUE;
	CFile fp;
	CString textFile;
	bool fileExists = false;
	int ret = 1;

	CString targetPrintSubFolder;

	if (selectedMailIndexList->GetCount() <= 0)
		return 1;

	ret = MakeFileNameFromMailArchiveName(textType, textFile, targetPrintSubFolder, fileExists, errorText);

	if (ret < 0)
		return -1;

	textFileName = textFile;
	// Useful durimg debug
	int textFileLength = textFile.GetLength();
	int maxPath = _MAX_PATH;

	if (!fp.Open(textFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone)) {
		errorText = _T("Could not create \"") + textFile;
		errorText += _T("\" file.\nMake sure file is not open on other applications.");
		return -1;
	}

	CFile fpm;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite)) {
		errorText = _T("Could not open mail archive \"") + s_path;
		fp.Close();
		return -1;
	}

	if (progressBar && MboxMail::pCUPDUPData)
	{
		if (textType == 0)
			MboxMail::pCUPDUPData->SetProgress(_T("Printing mails to TEXT file ..."), 0);
		else
			MboxMail::pCUPDUPData->SetProgress(_T("Printing mails to HTML file ..."), 0);
	}

	if (textType == 0) {
		if (textConfig.m_nCodePageId == CP_UTF8) {
			const char *BOM_UTF8 = "\xEF\xBB\xBF";
			fp.Write(BOM_UTF8, 3);
		}
	}

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	double delta = (double)(selectedMailIndexList->GetCount());
	if (delta <= 0) delta = 1;
	double step = delta / 100;
	double curstep = 1;
	double newstep = 0;
	CString fileNum;

	if (progressBar && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	int i;
	int cnt = selectedMailIndexList->GetCount();
	for (int j = 0; j < cnt; j++)
	{
		i = (*selectedMailIndexList)[j];

		if (textType == 0)
			printSingleMailToTextFile(fp, i, fpm, textConfig);
		else {
			printSingleMailToHtmlFile(fp, i, fpm, textConfig, TRUE);
		}

		if (progressBar && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				int deb = 1;
				break;
			}

			newstep = ((double)(j - 0 + 1)) / step;
			if (newstep >= curstep)
			{
				if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress((UINT_PTR)(newstep));
				curstep = newstep;
			}
#if 1
			int nFileNum = (j + 1);
			if ((nFileNum % 100) == 0) {
				if (textType == 0)
					fileNum.Format(_T("Printing mails to single TEXT file ... %d of %d"), nFileNum, cnt);
				else
					fileNum.Format(_T("Printing mails to single HTML file ... %d of %d"), nFileNum, cnt);
				if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(newstep));
			}
#endif
		}
	}
	
	newstep = 100;
	if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress((UINT_PTR)(newstep));
#if 1
	int nFileNum = cnt;
	if (textType == 0)
		fileNum.Format(_T("Printing mails to single TEXT file ... %d of %d"), nFileNum, cnt);
	else
		fileNum.Format(_T("Printing mails to single HTML file ... %d of %d"), nFileNum, cnt);
	if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(newstep));
#endif
	int deb = 1;

	fp.Close();
	fpm.Close();

	return 1;
}


void MboxMail::assert_unexpected()
{
#ifdef _DEBUG
	ASSERT(0);
#endif
	int deb = 1;
}

int MboxMail::EnforceCharacterLimit(SimpleString *buffer, CString &characterLimit)
{
	if (characterLimit.IsEmpty())
		return 1;

	int limit = _ttoi(characterLimit);
	int adjustedLimit = limit;
	if (limit > 0)
	{
		int bufferLength = buffer->Count();
		adjustedLimit = limit - 10;  // headroom
		if (adjustedLimit < 0)
			adjustedLimit = 11;
		if (bufferLength > adjustedLimit)
		{
			// Below doesn't quite work in rare cases if adjustedLimit index splits text between double quotation
			// such as abc "de fg hi"  , "de fg hi" represents single token, -> 'abc "de fg '  ; now de and fg are separate token
			// May need to implement better algo, however none is perfect
			int i;
			for (i = (adjustedLimit -1); i >= 0; i--)
			{
				// if (buffer->GetAt(i) == '\n') preferable ? but will not work if no '\n' in text
				if (buffer->GetAt(i) == ' ')  // likely works better
				{
					break;
				}
			}
			int newBufferLength = i + 1;
			buffer->SetCount(newBufferLength);
		}
		else // no change
		{
			int deb = 1;
		}
	}
	else
		buffer->SetCount(0);

	if ((buffer->Count() > limit)  && (buffer->Count() > adjustedLimit))
		int deb = 1;

	return 1;
}

int MboxMail::EnforceFieldTextCharacterLimit(char *buffer, int bufferLength, CString &characterLimit)
{
#if 0
	// characterLimit is not configurable yet
	if (characterLimit.IsEmpty())
		return bufferLength;
#endif

	int newBufferLength = bufferLength;
	int limit = _ttoi(characterLimit);
	limit = 32500;  // hardcode for now  (2^^15 -1) is (32768 - 1) is supported max by Excel

	int adjustedLimit = limit;
	if (limit > 0)
	{
		adjustedLimit = limit - 10;  // headroom
		if (adjustedLimit < 0)
			adjustedLimit = 11;
		if (bufferLength > adjustedLimit)
		{
			// find last non '"' character and return Count
			int i;
			for (i = (adjustedLimit - 1); i >= 0; i--)
			{
				// Below doesn't quite work if adjustedLimit index splits text between double quotation
				// such as abc "de fg hi"  , "de fg hi" represents single token, -> 'abc "de fg ' ; now de and fg are separate token
				// May need to better algo, however none is perfect
				if (buffer[i] != '"')
				{
					break;
				}
			}
			newBufferLength = i + 1;
			if (i < 0)
			{
				int deb = 1;
			}
		}
		else  // no change
		{
			int deb = 1;
		}

	}
	else
		newBufferLength = 0;

	if ((newBufferLength > limit) && (newBufferLength > adjustedLimit))
		int deb = 1;

	return newBufferLength;
}

const _int64 maxLargeLen = 100000;

BOOL MailBody::AssertData(MboxMail *mail)
{
	MboxMail *m = mail;

	AssertHdr();

	// Iterate all the descendant body parts
	MailBody::MailBodyList bodies;

	int nCount = GetBodyPartList(bodies);
	MailBody::MailBodyList::const_iterator it;

	for (it = bodies.begin(); it != bodies.end(); it++)
	{
		MailBody* pBP = *it;

		pBP->AssertHdr();

		int bodyLength = pBP->m_bodyDataLength;
		//char *bodyBegin = m->m_startOff + pBP->m_bodyDataOffset;

		if (pBP->m_bodyDataLength < 0)
			MboxMail::assert_unexpected();

		if (pBP->m_bodyDataLength > 0)
		{
			if (pBP->m_bodyDataOffset < 0)
				MboxMail::assert_unexpected();

			if (pBP->m_ContentType.GetLength() > maxLargeLen)
				MboxMail::assert_unexpected();

			if (pBP->m_Disposition.GetLength() > maxLargeLen)
				MboxMail::assert_unexpected();

			if (pBP->m_ContentId.GetLength() > maxLargeLen)
				MboxMail::assert_unexpected();

			if (pBP->m_TransferEncoding.GetLength() > maxLargeLen)
				MboxMail::assert_unexpected();

			if (pBP->m_AttachmentName.GetLength() > maxLargeLen)
				MboxMail::assert_unexpected();
		}
	}

	return TRUE;
}

bool MailHeader::AssertHdr()
{
	//bool m_IsText;
	//bool m_IsTextPlain;
	//bool m_IsTextHtml;
	//bool m_IsMessage;
	//bool m_IsAttachment;
	//bool m_IsMultiPart;


	if (m_Charset.GetLength() > maxLargeLen)
		MboxMail::assert_unexpected();

	//UINT m_PageCode;

	if (m_Description.GetLength() > maxLargeLen)
		MboxMail::assert_unexpected();

	if (m_Disposition.GetLength() > maxLargeLen)
		MboxMail::assert_unexpected();

	if (m_TransferEncoding.GetLength() > maxLargeLen)
		MboxMail::assert_unexpected();

	if (m_SubType.GetLength() > maxLargeLen)
		MboxMail::assert_unexpected();

	if (m_MainType.GetLength() > maxLargeLen)
		MboxMail::assert_unexpected();

	if (m_Boundary.GetLength() > maxLargeLen)
		MboxMail::assert_unexpected();

	if (m_ContentType.GetLength() > maxLargeLen)
		MboxMail::assert_unexpected();

	if (m_ContentId.GetLength() > maxLargeLen)
		MboxMail::assert_unexpected();

	//CMimeHeader::MediaType m_MediaType;

	if (m_AttachmentName.GetLength() > maxLargeLen)
		MboxMail::assert_unexpected();

	if (m_MessageId.GetLength() > maxLargeLen)
		MboxMail::assert_unexpected();

	if (m_ReplyId.GetLength() > maxLargeLen)
		MboxMail::assert_unexpected();

	return TRUE;
}


#include <windows.h>
#include <stdio.h>
#include <tchar.h>

// Use to convert bytes to KB
#define DIV 1024

// Specify the width of the field in which to print the numbers. 
// The asterisk in the format specifier "%*I64d" takes an integer 
// argument and uses it to pad and right justify the number.
#define WIDTH 7

void ShowMemStatus()
{
	MEMORYSTATUSEX statex;

	statex.dwLength = sizeof(statex);

	GlobalMemoryStatusEx(&statex);

	TRACE(TEXT("There is  %*ld percent of memory in use.\n"),
		WIDTH, statex.dwMemoryLoad);
	TRACE(TEXT("There are %*I64d total KB of physical memory.\n"),
		WIDTH, statex.ullTotalPhys / DIV);
	TRACE(TEXT("There are %*I64d free  KB of physical memory.\n"),
		WIDTH, statex.ullAvailPhys / DIV);
	TRACE(TEXT("There are %*I64d total KB of paging file.\n"),
		WIDTH, statex.ullTotalPageFile / DIV);
	TRACE(TEXT("There are %*I64d free  KB of paging file.\n"),
		WIDTH, statex.ullAvailPageFile / DIV);
	TRACE(TEXT("There are %*I64d total KB of virtual memory.\n"),
		WIDTH, statex.ullTotalVirtual / DIV);
	TRACE(TEXT("There are %*I64d free  KB of virtual memory.\n"),
		WIDTH, statex.ullAvailVirtual / DIV);

	// Show the amount of extended memory available.

	TRACE(TEXT("There are %*I64d free  KB of extended memory.\n"),
		WIDTH, statex.ullAvailExtendedVirtual / DIV);
}

HintConfig::HintConfig()
{
	m_nHintBitmap = 0;
	BOOL retval;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("hintBitmap"), m_nHintBitmap))
	{
		; // all done
	}
	else
	{
		m_nHintBitmap = 0xFFFFFFFF;

		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "hintBitmap", m_nHintBitmap);
	}
}

HintConfig::~HintConfig()
{
	int deb = 1;
}

BOOL HintConfig::IsHintSet(int hintNumber)
{
	DWORD position = 1 << hintNumber;
	if (m_nHintBitmap & position)
		return TRUE;
	else
		return FALSE;
}

void HintConfig::ClearHint(int hintNumber)
{
	DWORD position = 1 << hintNumber;
	m_nHintBitmap &= ~position;
}


void HintConfig::LoadHintBitmap()
{
	BOOL retval;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("hintBitmap"), m_nHintBitmap))
	{
		; //all done
	}
	else
	{
		m_nHintBitmap = 0xFFFFFFFF;
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "hintBitmap", m_nHintBitmap);
	}
}


void HintConfig::SaveToRegistry()
{
	CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "hintBitmap", m_nHintBitmap);
}


void MboxMail::LoadHintBitmap()
{
	MboxMail::m_HintConfig.LoadHintBitmap();
}

BOOL MboxMail::ParseDateInFromField(char *p, char *end, SYSTEMTIME *sysTime)
{
	static char *tm = "Thu Oct 27 09:02:59 +0000 2011";
	static int tmlen = strlen(tm);

	CString line;
	char *p_next = MimeParser::GetMultiLine(p, end, line);

	p = (char*)((LPCSTR)line);
	char *e = (p + line.GetLength()) - tmlen;

	time_t timeDT = -1;
	
	char c;
	while (p < e)
	{
		c = tolower(*p);
		if ((c == '\n') || (c == '\r'))
			return FALSE;

		if (c == 'm')
		{
			if (TextUtilsEx::strncmpUpper2Lower(p, end, "mon", 3) == 0)
				break;
		}
		else if (c == 't')
		{
			if (TextUtilsEx::strncmpUpper2Lower(p, end, "tue", 3) == 0)
				break;
			else if (TextUtilsEx::strncmpUpper2Lower(p, end, "thu", 3) == 0)
				break;
		}
		else if (c == 'w')
		{
			if (TextUtilsEx::strncmpUpper2Lower(p, end, "wed", 3) == 0)
				break;
		}
		else if (c == 'f')
		{
			if (TextUtilsEx::strncmpUpper2Lower(p, end, "fri", 3) == 0)
				break;
		}
		else if (c == 's')
		{
			if (TextUtilsEx::strncmpUpper2Lower(p, end, "sat", 3) == 0)
				break;
			else if (TextUtilsEx::strncmpUpper2Lower(p, end, "sun", 3) == 0)
				break;
		}
		p++;
	}
	if (p < e)
	{
		char *date = p;
		if (DateParser::parseRFC822Date(date, sysTime, 2))
			return TRUE;
		else
			return FALSE;
	}
	else
		return FALSE;
}

void MboxMail::ShowHint(int hintNumber, HWND h)
{
	// Or should I create table with all Hints to support potentially large number of hints ?
	CString hintText;
	BOOL isHintSet = MboxMail::m_HintConfig.IsHintSet(hintNumber);
	if (isHintSet)
	{
		if (hintNumber == HintConfig::GeneralUsageHint)
		{
			hintText.Append(
				"To get started, please install one or more mbox files in\n"
				"the local folder and then select \"File->Select Folder...\"\n"
				"menu option to open that folder.\n"
				"\n"
				"Left click on one of the loaded mbox file to view all associated mails.\n"
				"\n"
				"Please review the User Guide provided with the package\n"
				"and/or right/left single and double click on any item\n"
				"within the Mbox Viewer window and try all presented options."
			);
		}
		else if (hintNumber == HintConfig::MsgWindowPlacementHint)
		{
			hintText.Append(
				"You can place  Message Window at Bottom (default), Left or Right.\n"
				"\n"
				"Select \"View->Message Window...\" to configure."
			);
		}
		else if (hintNumber == HintConfig::PrintToPDFHint)
		{
			hintText.Append(
				"Chrome Browser is used by deafult to print to PDF file.\n"
				"However, the header and footer are always printed.\n "
				"\n"
				"Select \"File->Print Config->Path to User Defined Script\" to \n"
				"use free wkhtmltopdf tool to remove header and footer\n"
				"and evalute if wkhtmltopdf works for you."
			);
		}
		else if (hintNumber == HintConfig::PrintToPrinterHint)
		{
			hintText.Append(
				"Select \"File->Print Config->Page Setup\" to configure\n"
				"page header, footer, etc via Windows standard setup.\n"
			);
		}
		else if (hintNumber == HintConfig::MailSelectionHint)
		{
			hintText.Append(
				"You can select multiple mails using standard Windows methods: "
				"\"Shift-Left Click\", \"Ctrl-Left Click\" and \"Ctrl-A\".\n\n"
				"Right click on a single or multiple selected mails to see all available options.\n\n"
				"Right click on a mail attachment to see all available options.\n\n"
				"Left double click on a mail to open folder with all associated files for that mail."
			);
		}
		else if (hintNumber == HintConfig::FindDialogHint)
		{
			hintText.Append(
				"You can specify single ‘*’ character as the search string to FIND dialog to find or traverse subset of mails:\n"
				"\n"
				"1. Find mails that have CC header field by checking out CC check box only.\n\n"
				"2. Find mails that have BCC header field by checking out BCC check box only.\n\n"
				"3. Find mails that have at least one attachment by checking out Attachment Name check box only.\n\n"
				"4. Match all mails by checking out any of others check boxes only.\n\n"
			);
		}
		else if (hintNumber == HintConfig::AdvancedFindDialogHint)
		{
			hintText.Append(
				"You can specify single ‘*’ character as the search string in any of the Filter fields in Advanced Find dialog to find subset of mails:\n"
				"\n"
				"1. Find all mails that have CC header field by checking out CC check box only.\n\n"
				"2. Find all mails that have BCC header field by checking out BCC check box only.\n\n"
				"3. Find all mails that have at least one attachment by checking out Attachment Name check box only.\n\n"
				"4. Match all mails by checking out any of others check boxes only.\n\n"
			);
		}
		else if (hintNumber == HintConfig::PrintToPDFScriptHint)
		{
			hintText.Append(
				"If you need to remove the background color when printing to\n"
				"PDF file directly you need to configure wkhtmltopdf for\n"
				"printing. By default the Chrome browser is configured to\n"
				"print to PDF file. However, Chrome browser doesn't support\n"
				"removing the background color via the command line option\n\n"
				"Select \"File->Print Config->Page Setup\" to configure\n"
				"HTML2PDF-single-wkhtmltopdf.cmd script.\n\n"
				"Note that the background color can be removed via\n"
				"Page Setup when printing to PDF file\n"
				"via Print to Printer option and by opening the mail\n"
				"within a browser for printing.\n"
			);
		}
		else if (hintNumber == HintConfig::AttachmentConfigHint)
		{
			hintText.Append(
				"By default attachments other than those embeded into mail\n"
				"are shown in the attachment window.\n\n"
				"You can configure to shown all attachments, both inline and\n"
				"non-inline, by selecting\n"
				"\"File->Attachments Config->Attachment Window\" dialog\n"
			);
		}


		if (!hintText.IsEmpty())
		{
			::MessageBox(h, hintText, _T("Did you know?"), MB_OK | MB_ICONEXCLAMATION);
#if 0
			// Right now wil not use this
			MSGBOXPARAMSA mbp;
			mbp.cbSize = 0;
			mbp.hwndOwner = h;
			mbp.hInstance = 0;
			mbp.lpszText =  hintText;
			mbp.lpszCaption = _T("Did you know?");
			mbp.dwStyle = 0;
			mbp.lpszIcon = "";
			mbp.dwContextHelpId = 0;
			mbp.dwLanguageId = 0;
			int ret = MessageBoxIndirect(&mbp);
#endif
			MboxMail::m_HintConfig.ClearHint(hintNumber);
		}
	}
}

// TODO: Need to add menu option to display Copyright/License information
const char *LicensText =
"  Windows Mbox Viewer is a free tool to view, search and print mbox mail archives.\n"
"\n"
"  Copyright(C) 2019  Enea Mansutti, Zbigniew Minciel\n"
"\n"
"  This program is free software; you can redistribute it and/or modify\n"
"  it under the terms of the version 3 of GNU Affero General Public License\n"
"  as published by the Free Software Foundation.\n"
"\n"
"  This program is distributed in the hope that it will be useful,\n"
"  but WITHOUT ANY WARRANTY; without even the implied warranty of\n"
"  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU\n"
"  Library General Public License for more details.\n"
"\n"
"  You should have received a copy of the GNU Library General Public\n"
"  License along with this program; if not, write to the\n"
"  Free Software Foundation, Inc., 51 Franklin St, Fifth Floor,\n"
"  Boston, MA  02110 - 1301, USA.\n";
