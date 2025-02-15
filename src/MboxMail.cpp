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
#include "ResHelper.h"

#ifdef USE_STACK_WALKER
#include "StackWalker.h"
#endif

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

MyStackWalker *MboxMail::glStackWalker = 0;
BOOL MboxMail::ignoreException = FALSE;
int MboxMail::runningWorkerThreadType = 0;

BOOL MboxMail::m_seExceptionMsgBox = TRUE;
BOOL MboxMail::m_cppExceptionMsgBox = FALSE;

BOOL MboxMail::developerMode = FALSE;

HintConfig MboxMail::m_HintConfig;

MailBodyPool *MailBody::m_mpool = new MailBodyPool;

ThreadIdTableType *MboxMail::m_pThreadIdTable = 0;
MessageIdTableType *MboxMail::m_pMessageIdTable = 0;
MboxMailTableType *MboxMail::m_pMboxMailTable = 0;
MboxMail::MboxMailMapType *MboxMail::m_pMboxMailMap = 0;
int MboxMail::m_nextGroupId = 0;

SimpleString* MboxMail::m_outbuf = new SimpleString(1*1024*1024);
SimpleString* MboxMail::m_inbuf = new SimpleString(1*1024*1024);
SimpleString* MboxMail::m_outdata = new SimpleString(1*1024*1024);
SimpleString* MboxMail::m_indata = new SimpleString(1*1024*1024);
SimpleString* MboxMail::m_workbuf = new SimpleString(1*1024*1024);
SimpleString* MboxMail::m_tmpbuf = new SimpleString(1*1024*1024);
SimpleString* MboxMail::m_largebuf = new SimpleString(1*1024*1024);
//
SimpleString* MboxMail::m_largelocal1 = new SimpleString(10*1024*1024);
SimpleString* MboxMail::m_largelocal2 = new SimpleString(10*1024*1024);
SimpleString* MboxMail::m_largelocal3 = new SimpleString(10*1024*1024);
SimpleString* MboxMail::m_smalllocal1 = new SimpleString(1*1024*1024);
SimpleString* MboxMail::m_smalllocal2 = new SimpleString(1*1024*1024);


BOOL MboxMail::m_outbufBusy = FALSE;
BOOL MboxMail::m_inbufBusy = FALSE;
BOOL MboxMail::m_outdataBusy = FALSE;
BOOL MboxMail::m_indataBusy = FALSE;
BOOL MboxMail::m_workbufBusy = FALSE;
BOOL MboxMail::m_tmpbufBusy = FALSE;
BOOL MboxMail::m_largebufBusy = FALSE;

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
MailArray MboxMail::s_mails_label;

MailList *MboxMail::m_mailList = &MboxMail::m_allMails;

DLLIST(MailList, m_listLink) MboxMail::m_folderList;

MailList MboxMail::m_allMails(IDC_ARCHIVE_LIST);
MailList MboxMail::m_findMails(IDC_FIND_LIST);
MailList MboxMail::m_editMails(IDC_EDIT_LIST);
MailList MboxMail::m_folderMails(IDC_FOLDER_LIST);
MailList MboxMail::m_labelMails(IDC_LABEL_LIST);

MailArray MboxMail::s_mails_selected;
MailArray MboxMail::s_mails_merged;

_int64 MboxMail::s_fSize = 0;
_int64 MboxMail::s_oSize = 0;
CString MboxMail::s_path;
CString MboxMail::s_path_label;
CString MboxMail::s_datapath;  // root directory for files created by mbox viewer such as index file, PrintCache
int MboxMail::nWhichMailList = -1;
FolderContext MboxMail::s_folderContext;

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

MailEncodingStats MboxMail::m_textEncodingStats;

UINT getCodePageFromHtmlBody(SimpleString *buffer, std::string &charset);


BOOL MailBodyContent::AttachmentName2WChar(CString& attachmentName)
{
	attachmentName.Empty();
	UINT nameCharsetId = m_attachmentNamePageCode;
	if (m_attachmentNamePageCode == 0)
		nameCharsetId = CP_UTF8;

	DWORD dwFlags = 0;
	DWORD error = 0;
	if (TextUtilsEx::Str2WStr(m_attachmentName, nameCharsetId, attachmentName, error, dwFlags))
	{
		return TRUE;
	}
	else
		return FALSE;
}

#if 0
bool MailBodyContent::IsAttachment()
{
	// text/plain and text/html can be attachments or  main mail message
	// check if attachment and not the mail main body
	if ((m_contentType.CompareNoCase("text/plain") == 0) ||
		(m_contentType.CompareNoCase("text/html") == 0))
	{
		if (!m_attachmentName.IsEmpty() ||
			!m_contentId.IsEmpty() ||
			!m_contentLocation.IsEmpty() ||
			(m_contentDisposition.CompareNoCase("attachment") == 0)
			)
			return true;
		else
			return false;
	}
	else
		return true;
}
#else
bool MailBodyContent::IsAttachment()
{
	// Temp fix in 1.0.3.39 for incorrect determination of attachment type. Review in v1.0.3.40 UNICODE

	BOOL isContentTypeText;
	BOOL isTextPlain;
	BOOL isTextHtml;
	CStringA contentSubType;
	CStringA contentTypeMain;

	MailBodyContent* body = this;

	int pos = body->m_contentType.ReverseFind('/');
	if (pos > 0)
	{
		contentSubType = body->m_contentType.Mid(pos + 1);
		contentTypeMain = body->m_contentType.Left(pos);
	}

	if (contentTypeMain.CompareNoCase("text") != 0) {
		isContentTypeText = TRUE;
	}
	if (contentSubType.CompareNoCase("plain") != 0) {
		isTextPlain = TRUE;
	}
	if (contentSubType.CompareNoCase("html") != 0) {
		isTextHtml = TRUE;
	}

	// Consider all text type blocks; not only plain and html
	if (((m_contentType.CompareNoCase("text/html") == 0) || (m_contentType.CompareNoCase("text/plain") == 0)) &&
		(m_contentDisposition.CompareNoCase("attachment") != 0))
	{
		return false;
	}
	// end fix
	else if (!m_attachmentName.IsEmpty() ||
		!m_contentId.IsEmpty() ||
		!m_contentLocation.IsEmpty() ||
		(m_contentDisposition.CompareNoCase("attachment") == 0) ||
		((m_contentType.CompareNoCase("text/html") != 0) && (m_contentType.CompareNoCase("text/plain") != 0))
		)
	{
		return true;
	}
	else
		return false;
}

bool MailBodyContent::IsInlineAttachment()
{
	// Temp fix in 1.0.3.39 for incorrect determination of attachment type. Review in v1.0.3.40 UNICODE

	BOOL isContentTypeText;
	BOOL isTextPlain;
	BOOL isTextHtml;
	CStringA contentSubType;
	CStringA contentTypeMain;

	MailBodyContent* body = this;

	int pos = body->m_contentType.ReverseFind('/');
	if (pos > 0)
	{
		contentSubType = body->m_contentType.Mid(pos + 1);
		contentTypeMain = body->m_contentType.Left(pos);
	}

	if (contentTypeMain.CompareNoCase("text") != 0) {
		isContentTypeText = TRUE;
	}
	if (contentSubType.CompareNoCase("plain") != 0) {
		isTextPlain = TRUE;
	}
	if (contentSubType.CompareNoCase("html") != 0) {
		isTextHtml = TRUE;
	}

	// Consider all text type blocks; not only plain and html
	if (((m_contentType.CompareNoCase("text/html") == 0) || (m_contentType.CompareNoCase("text/plain") == 0)) &&
		(m_contentDisposition.CompareNoCase("inline") == 0))
	{
		if (!m_attachmentName.IsEmpty() ||
			!m_contentId.IsEmpty() ||
			!m_contentLocation.IsEmpty())
			return true;
		else
			return false;
	}
	else
		return false;
}

#endif

struct MsgIdHash {
public:
	hashsum_t operator()(const MboxMail *key) const
	{
		hashsum_t hashsum = StrHash((const char*)(LPCSTR)key->m_messageId, key->m_messageId.GetLength());
		return hashsum;
	}
};

bool MailBodyContent::DecodeBodyData(char* msgData, int msgLength, SimpleString* outbuf)
{
	MailBodyContent* body = this;

	int bodyLength = body->m_contentLength;
	_ASSERTE((body->m_contentOffset + body->m_contentLength) <= msgLength);
	if ((body->m_contentOffset + body->m_contentLength) > msgLength) {
		// something is not consistent
		bodyLength = msgLength - body->m_contentOffset;
	}
	char* bodyBegin = msgData + body->m_contentOffset;

	if (body->m_contentTransferEncoding.CompareNoCase("base64") == 0)
	{
		MboxCMimeCodeBase64 d64(bodyBegin, bodyLength);
		int dlength = d64.GetOutputLength();
		outbuf->ClearAndResize(dlength + 1);

		int retlen = d64.GetOutput((unsigned char*)outbuf->Data(), dlength);
		if (retlen > 0)
		{
			outbuf->SetCount(retlen);
		}
		else
		{
			outbuf->Clear();
		}
	}
	else if (body->m_contentTransferEncoding.CompareNoCase("quoted-printable") == 0)
	{
		MboxCMimeCodeQP dGP(bodyBegin, bodyLength);
		int dlength = dGP.GetOutputLength();
		outbuf->ClearAndResize(dlength + 1);

		int retlen = dGP.GetOutput((unsigned char*)outbuf->Data(), dlength);
		if (retlen > 0)
		{
			outbuf->SetCount(retlen);
		}
		else
		{
			outbuf->Clear();
		}
	}
	else
	{
		outbuf->ClearAndResize(bodyLength + 1);
		outbuf->Append(bodyBegin, bodyLength);
	}
	return true;
}

struct MboxHash {
public:
	hashsum_t operator()(const MboxMail *key) const
	{
		hashsum_t hashsum;
		if (key->m_messageId.GetLength() > 0)
			hashsum = StrHash((const char*)(LPCSTR)key->m_messageId, key->m_messageId.GetLength());
		else
		{
			hashsum_t hashsum1 = StrHash((const char*)(LPCSTR)key->m_from, key->m_from.GetLength());
			hashsum_t hashsum2 = StrHash((const char*)(LPCSTR)key->m_to, key->m_to.GetLength());
			hashsum = hashsum1 + hashsum2 + (0xffff & key->m_timeDate);
		}
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

struct MboxEqual {
public:
	bool operator()(const MboxMail *key1, const MboxMail *key2) const
	{
		if ((key1->m_timeDate == key2->m_timeDate) &&
			(key1->m_messageId == key2->m_messageId) &&
			(key1->m_from == key2->m_from) &&
			(key1->m_threadId == key2->m_threadId) &&  // works for Gmails; works for non-gmail since m_threadId are empty
			(key1->m_to == key2->m_to) )  // last; might be the longest string
			return true;
		else
			return false;
	}
};

struct MessageIdHash {
public:
	hashsum_t operator()(const CStringA *key) const
	{
		hashsum_t hashsum = StrHash((const char*)(LPCSTR)*key, key->GetLength());
		return hashsum;
	}
};

struct MessageIdEqual {
public:
	bool operator()(const CStringA *key1, const CStringA *key2) const
	{
		if (*key1 == *key2)
			return true;
		else
			return false;
	}
};

struct ThrdIdHash {
public:
	hashsum_t operator()(const MboxMail *key) const
	{
		hashsum_t hashsum = StrHash((const char*)(LPCSTR)key->m_threadId, key->m_threadId.GetLength());
		return hashsum;
	}
};

struct ThrdIdEqual {
public:
	bool operator()(const MboxMail *key1, const MboxMail *key2) const
	{
		if (key1->m_threadId == key2->m_threadId)
			return true;
		else
			return false;
	}
};

struct ThreadIdHash {
public:
	hashsum_t operator()(const CStringA *key) const
	{
		hashsum_t hashsum = StrHash((const char*)(LPCSTR)*key, key->GetLength());
		return hashsum;
	}
};

struct ThreadIdEqual {
public:
	bool operator()(const CStringA *key1, const CStringA *key2) const
	{
		if (*key1 == *key2)
			return true;
		else
			return false;
	}
};

struct AddressHash {
public:
	hashsum_t operator()(const CStringA* key) const
	{
		hashsum_t hashsum = StrHash((const char*)(LPCSTR)*key, key->GetLength());
		return hashsum;
	}
};

struct MailFromToAddressInfoEqual {
public:
	bool operator()(const CStringA* key1, const MailFromToAddressInfo* key2) const
	{
		if (*key1 == key2->m_fromAddress)
			return true;
		else
			return false;
	}
};

void FolderContext::SetFolderPath(CString &folderPath)
{
	CString section_lastSelection = CString(sz_Software_mboxview) + L"\\LastSelection";

	CString path = folderPath;
	path.TrimRight(L"\\");
	if (path.IsEmpty())
	{
		MboxMail::assert_unexpected();
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_lastSelection, L"lastPath", path);
		m_folderPath = "";
		m_dataFolderPath = "";
		return;
	}

	path.Append(L"\\");

	if (m_folderPath.Compare(path) == 0)
	{
		m_setPathCount++;  // redundant calls count
		return;
	}

	if (m_setPathCount > 1)
		TRACE(L"SetFolderPath: SetPathCount=%d\n", m_setPathCount);

	m_setPathCount = 0;

	CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_lastSelection, L"lastPath", path);

	m_folderPath = path;

	if (m_rootDataFolderPathConfig.IsEmpty())
		m_rootDataFolderPath.Empty();
	else
		m_rootDataFolderPath = m_rootDataFolderPathConfig;

	if (m_rootDataFolderPath.IsEmpty())
	{
		m_dataFolderPath = path + "UMBoxViewer\\";
		BOOL isReadOnly = FileUtils::IsReadonlyFolder(m_dataFolderPath);
		if (!isReadOnly)
		{
			return;
		}
		else
		{
			m_rootDataFolderPath = FileUtils::CreateMboxviewLocalAppPath();
			m_dataFolderPath = m_rootDataFolderPath + "UMBoxViewer\\";
			BOOL isReadOnly = FileUtils::IsReadonlyFolder(m_rootDataFolderPath);
			if (isReadOnly)
			{
				MboxMail::assert_unexpected();
				// start dialog ??
			}

		}
	}
	else
	{
		m_dataFolderPath = m_rootDataFolderPath + "UMBoxViewer\\";
		if (!m_rootDataFolderPathSubFolderConfig.IsEmpty())
			m_dataFolderPath.Append(m_rootDataFolderPathSubFolderConfig);
	}

	CString fileName;
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(path, driveName, directory, fileNameBase, fileNameExtention);

	driveName.TrimRight(L":");
	m_dataFolderPath = m_dataFolderPath + driveName + directory;

	int deb = 1;
}
void FolderContext::GetFolderPath(CString &folderpath)
{
	folderpath = m_folderPath;
}

// FIXME not used yet
void FolderContext::GetDataPath(CString &dataFolderPath)
{
	dataFolderPath = m_dataFolderPath;
}
// 
BOOL FolderContext::LoadFromRegistry()
{
	return TRUE;
}
BOOL FolderContext::SaveToRegistry()
{
	return TRUE;
}

void MboxMail::SetMboxFilePath(CString &filePath, BOOL ignoreCheck)
{
	if (!ignoreCheck && filePath.IsEmpty())
	{
		assert_unexpected();
	}
	s_path = filePath;

};

void MboxMail::SetLastPath(CString &lastPath)
{
	CString path = lastPath;
	path.TrimRight(L"\\");
	path.Append(L"\\");
	MboxMail::s_datapath = "";

	CString section_lastSelection = CString(sz_Software_mboxview) + L"\\LastSelection";

	CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_lastSelection, L"lastPath", path);
	s_folderContext.SetFolderPath(lastPath);

	if (path.IsEmpty())
	{
		return;
	}

	if (!s_folderContext.m_rootDataFolderPath.IsEmpty())
	{
		if (_tcsncmp(s_folderContext.m_dataFolderPath, s_folderContext.m_rootDataFolderPath, s_folderContext.m_rootDataFolderPath.GetLength()))
			MboxMail::assert_unexpected();
	}
	s_datapath = s_folderContext.m_dataFolderPath;

	int deb = 1;
}

CString MboxMail::GetLastPath()
{
	CString section_lastSelection = CString(sz_Software_mboxview) + L"\\LastSelection";

	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_lastSelection, L"lastPath");

	if (!path.IsEmpty() && s_datapath.IsEmpty())
		int deb = 1;
	if (path.Compare(s_datapath))
		int deb = 1;

	return path;
}

CString MboxMail::GetLastDataPath()
{
	if (s_datapath.IsEmpty())
		int deb = 1;
	return s_datapath;
}

CString MboxMail::SetLastDataPath(CString *lastMboxDirPath)
{
	CString path = MboxMail::GetLastPath();
	if (lastMboxDirPath)
		path = *lastMboxDirPath;

	path.TrimRight(L"\\");
	path.Append(L"\\");

	CString fileName;
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(path, driveName, directory, fileNameBase, fileNameExtention);

	driveName.TrimRight(L":");
	CString lastPath = driveName + directory;

	CString datapath = FileUtils::CreateMboxviewLocalAppDataPath();
	if (!datapath.IsEmpty())
		datapath += lastPath;

	return datapath;
}

#include "DateParser.h"

CString MboxMail::GetDateFormat(int i)
{
	CString format;
	switch (i) {
	case 1:
		format = L"%m/%d/%Y %H:%M";
		break;
	case 2:
		format = L"%Y/%m/%d %H:%M";
		break;
	default:
		format = L"%d/%m/%Y %H:%M";
		break;
	}
	return format;
}
CString MboxMail::GetPickerDateFormat(int i)
{
	CString format;
	switch (i) {
	case 1:
		format = L"MM/dd/yyyy";
		break;
	case 2:
		format = L"yyyy/MM/dd";
		break;
	default:
		format = L"dd/MM/yyyy";
		break;
	}
	return format;
}

BOOL MboxMail::GetBody(CStringA &res)
{
	BOOL ret = TRUE;
	CFile fp;
	CFileException ExError;
	if (fp.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
	{
		char *p = res.GetBufferSetLength(m_length);
		//TRACE(L"offset = %lld\n", m_startOff);
		_int64 filePos = fp.Seek(m_startOff, CFile::begin);  // same as SEEK_SET == 0
		if (filePos != m_startOff)
		{
			_ASSERTE(FALSE);
		}

		UINT readbytes = fp.Read(p, m_length);
		if (readbytes != m_length)
		{
			_ASSERTE(FALSE);
		}
		res.ReleaseBufferSetLength(readbytes);
		char *ms = strchr(p, '\n'); //"-Version: 1.0"); // FIXME need test to see if needed
		// TODO: Should we always make sure lines are terminated with CR LF
		// Sometimes mail files have different  line ending at the end  of a file, etc.
		// Or enhance mime parser to handle line ending with CR or NL only
		if (ms) 
		{
			BOOL bAddCR = FALSE;
			if (*(ms - 1) != '\r')
				bAddCR = TRUE;
			int pos = IntPtr2Int(ms - p + 1);
			//res = res.Mid(pos); // - 4);
			if (bAddCR) // for correct mime parsing
			{
				SimpleString *tmpbuf = MboxMail::get_tmpbuf();
				TextUtilsEx::ReplaceNL2CRNL((LPCSTR)res, res.GetLength(), tmpbuf);
				res.Empty();
				res.Append(tmpbuf->Data(), tmpbuf->Count());
				MboxMail::rel_tmpbuf();

				//res.Replace("\n", "\r\n");
			}
			else
				int deb = 1;
		}
		fp.Close();  // TODO: verify why Close() was not called
	}
	else
	{
		DWORD lastErr = ::GetLastError();
#if 1
		//HWND h = GetSafeHwnd();
		HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		CString fmt = L"Could not open mail file:\n\n\"%s\"\n\n%s";  // new format
		CString errorText = FileUtils::ProcessCFileFailure(fmt, MboxMail::s_path, ExError, lastErr, h); 
#else
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not open \"" + MboxMail::s_path;
		txt += L"\" mail file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);
		//errorText = txt;
#endif

		ret = FALSE;
	}
	return ret;
};

BOOL MboxMail::GetBodySS(CFile &fp, SimpleString *res, int maxLength)
{
	BOOL ret = TRUE;
	int bytes2Read = m_length;
	if (maxLength > 0)
	{
		if (maxLength < m_length)
			bytes2Read = maxLength;
	}

	res->Resize(bytes2Read);
	//TRACE(L"offset = %lld\n", m_startOff);

	_int64 filePos = fp.Seek(m_startOff, CFile::begin);  // SEEK_SET == CFile::begin == 0
	UINT readLength = fp.Read(res->Data(), bytes2Read);
	if (readLength != bytes2Read)
	{
		_ASSERTE(FALSE);
	}
	res->SetCount(bytes2Read);
	return ret;
};

BOOL MboxMail::GetBodySS(SimpleString *res, int maxLength)
{
	BOOL ret = TRUE;
	CFile fp;
	res->Clear();
	int bytes2Read = m_length;
	if (maxLength > 0)
	{
		if (maxLength < m_length)
			bytes2Read = maxLength;
	}
	CFileException ExError;
	if (fp.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
	{
		res->Resize(bytes2Read);
		//TRACE(L"offset = %lld\n", m_startOff);

		_int64 filePos = fp.Seek(m_startOff, CFile::begin);  // SEEK_SET == CFile::begin == 0
		if (filePos != m_startOff)
		{
			_ASSERTE(FALSE);
		}

		UINT readLength = fp.Read(res->Data(), bytes2Read);
		if (readLength != bytes2Read)
		{
			_ASSERTE(FALSE);
		}
		res->SetCount(bytes2Read);

		char* p = res->Data();
		char* ms = strchr(p, '\n'); //"-Version: 1.0");  // FIXME need test to see if needed
		// TODO: Should we always make sure lines are terminated with CR LF
		// Sometimes mail files have different  line ending at the end  of a file, etc.
		// Or enhance mime parser to handle line ending with CR NL or NL only
		if (ms)
		{
			BOOL bAddCR = FALSE;
			if (*(ms - 1) != '\r')
				bAddCR = TRUE;
			int pos = IntPtr2Int(ms - p + 1);
			//res = res.Mid(pos); // - 4);
			if (bAddCR) // for correct mime parsing
			{
				SimpleString* tmpbuf = MboxMail::get_tmpbuf();
				TextUtilsEx::ReplaceNL2CRNL(res->Data(), res->Count(), tmpbuf);
				res->ClearAndResize(tmpbuf->Count()+1);
				res->Append(tmpbuf->Data(), tmpbuf->Count());
				MboxMail::rel_tmpbuf();
			}
			else
				int deb = 1;
		}
		fp.Close();
	}
	else
	{
		DWORD lastErr = ::GetLastError();
#if 1
		//HWND h = GetSafeHwnd();
		HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		CString fmt = L"Could not open mail file:\n\n\"%s\"\n\n%s";  // new format
		CString errorText = FileUtils::ProcessCFileFailure(fmt, MboxMail::s_path, ExError, lastErr, h);
#else
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not open \"" + MboxMail::s_path;
		txt += L"\" mail file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);
		//errorText = txt;
#endif

		ret = FALSE;
	}

	return ret;
};

int MboxMail::SingleMailSizeof()
{
	int cstringSz = sizeof(CString);

	int thisSize = 0;
	{
		thisSize = sizeof(m_startOff);
		thisSize += sizeof(m_length);
		thisSize += sizeof(m_hasAttachments);
		thisSize += sizeof(m_headLength);

		thisSize += sizeof(m_from_charsetId);
		thisSize += sizeof(m_to_charsetId);
		thisSize = sizeof(m_subj_charsetId);

		thisSize += sizeof(m_cc_charsetId);
		thisSize += sizeof(m_bcc_charsetId);


		thisSize += sizeof(m_timeDate);
		thisSize += sizeof(m_recv);
		thisSize += sizeof(m_groupId);
		thisSize += sizeof(m_nextMail);
		thisSize += sizeof(m_prevMail);

		thisSize += sizeof(m_duplicateId);
		thisSize += sizeof(m_done);
		thisSize += sizeof(m_groupColor);
		thisSize += sizeof(m_index);
		thisSize += sizeof(m_headLength);

		//m_crc32 = 0xffffffff;

		thisSize += sizeof(m_isOnUserSelectedMailList);
		thisSize += sizeof(m_DetermineEmbeddedImagesDone);


		//thisSize += sizeof(fp);
		thisSize += sizeof(m_ContentDetailsArray);

		thisSize += sizeof(m_hashMapLink);

		thisSize += sizeof(m_from);
		thisSize += sizeof(m_to);
		thisSize += sizeof(m_subj);
		thisSize += sizeof(m_cc);
		thisSize += sizeof(m_bcc);

		thisSize += sizeof(m_from_charset);
		thisSize += sizeof(m_to_charset);
		thisSize += sizeof(m_subj_charset);
		thisSize += sizeof(m_cc_charset);
		thisSize += sizeof(m_bcc_charset);

		thisSize += sizeof(m_messageId);
		thisSize += sizeof(m_replyId);
		thisSize += sizeof(m_threadId);

		int thisSizeof = sizeof(*this);
		int delta = thisSizeof - thisSize;
		// delta = 33 we are missing some overhead, class overhead ?
		int deb = 1;
	}

	int allocExtra = 28;  // TODO: try to do better than just a guess
	int fieldsSz = sizeof(*this) + allocExtra;
	fieldsSz += m_from.GetAllocLength() + allocExtra;
	fieldsSz += m_to.GetAllocLength() + allocExtra;
	fieldsSz += m_subj.GetAllocLength() + allocExtra;
	fieldsSz += m_cc.GetAllocLength() + allocExtra;
	fieldsSz += m_bcc.GetAllocLength() + allocExtra;

	//
	int charsetSz = 0;
	charsetSz += m_from_charset.GetAllocLength() + allocExtra;
	charsetSz += m_to_charset.GetAllocLength() + allocExtra;
	charsetSz += m_subj_charset.GetAllocLength() + allocExtra;
	charsetSz += m_cc_charset.GetAllocLength() + allocExtra;
	charsetSz += m_bcc_charset.GetAllocLength() + allocExtra;
	//
	int msgidsSz = 0;
	msgidsSz += m_messageId.GetAllocLength() + allocExtra;
	msgidsSz += m_replyId.GetAllocLength() + allocExtra;
	msgidsSz += m_threadId.GetAllocLength() + allocExtra;
	//

	int bodySz = 0;
	MailBodyContent *body;
	for (int j = 0; j < m_ContentDetailsArray.size(); j++)
	{
		body = m_ContentDetailsArray[j];
		bodySz = sizeof(*body) + allocExtra;
		bodySz += body->m_contentType.GetAllocLength() + allocExtra;
		bodySz += body->m_contentTransferEncoding.GetAllocLength() + allocExtra;
		bodySz += body->m_contentDisposition.GetAllocLength() + allocExtra;
		bodySz += body->m_contentId.GetAllocLength() + allocExtra;
		bodySz += body->m_attachmentName.GetAllocLength() + allocExtra;
		bodySz += body->m_contentLocation.GetAllocLength() + allocExtra;
	}

	int sz = sizeof(*this) + fieldsSz + charsetSz + msgidsSz + bodySz;
	return  sz;
};

size_t MboxMail::AllMailsSizeof(int mailCount)
{
	MboxMail *m;
	int singleMailsMemorySz = 0;
	size_t allMailsMemorySz = 0;

	int maxCnt = s_mails.GetCount();
	if (mailCount > maxCnt)
		mailCount = maxCnt;

	for (int i = 0; i < mailCount; i++)
	{
		m = s_mails[i];
		if (m == 0)
			break;

		singleMailsMemorySz = m->SingleMailSizeof();
		allMailsMemorySz += singleMailsMemorySz;
	}

	int bufSz = 0;
	if (m_outbuf)
		bufSz += m_outbuf->Capacity();
	if (m_inbuf)
		bufSz += m_inbuf->Capacity();
	if (m_outdata)
		bufSz += m_outdata->Capacity();
	if (m_indata)
		bufSz += m_indata->Capacity();
	if (m_workbuf)
		bufSz += m_workbuf->Capacity();
	if (m_tmpbuf)
		bufSz += m_tmpbuf->Capacity();
	if (m_largebuf)
		bufSz += m_largebuf->Capacity();

	size_t mailArraysSz = 0;
	mailArraysSz += sizeof(s_mails_ref);
	mailArraysSz += sizeof(s_mails);
	mailArraysSz += sizeof(s_mails_all);
	mailArraysSz += sizeof(s_mails_find);
	mailArraysSz += sizeof(s_mails_edit);
	mailArraysSz += sizeof(s_mails_folder);
	mailArraysSz += sizeof(s_mails_selected);
	mailArraysSz += sizeof(s_mails_merged);

	char *p;
	mailArraysSz += s_mails_ref.GetUpperBound()*(1 * sizeof(p));
	mailArraysSz += s_mails.GetUpperBound()*(1 * sizeof(p));
	mailArraysSz += s_mails_all.GetUpperBound()*(1 * sizeof(p));
	mailArraysSz += s_mails_find.GetUpperBound()*(1 * sizeof(p));
	mailArraysSz += s_mails_edit.GetUpperBound()*(1 * sizeof(p));
	mailArraysSz += s_mails_folder.GetUpperBound()*(1 * sizeof(p));
	mailArraysSz += s_mails_selected.GetUpperBound()*(1 * sizeof(p));
	mailArraysSz += s_mails_merged.GetUpperBound()*(1 * sizeof(p));


	allMailsMemorySz += bufSz + mailArraysSz;

	return allMailsMemorySz;
}

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

BOOL FoundMonth()
{
	return TRUE;
}

void MboxMail::MonthToString(int monthInt, CStringA &monthStr)
{
	static char *months[] = { "jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec" };
	if ((monthInt >= 1) && (monthInt <= 12))
	{
		monthStr = months[monthInt-1];
	}
	return;
}

void MboxMail::FindDateInHeader(char *data, int datalen, CStringA &dateStr)
{
	static char *days[] = { "sun","mon","tue","wed","thu","fri","sat" };
	static char *months[] = { "jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec" };

	int headerlen = NMsgView::FindMailHeader(data, datalen);


	char *p = data;
	char *e = data + headerlen;

	// probbaly should parse and create dictionary of words
	char c;
	BOOL found = FALSE;
	while (p < e)
	{
		found = FALSE;
		c = *p;
		c = tolower(c);
		switch (c)
		{
		case 'a':
		{
			if (TextUtilsEx::strncmpUpper2Lower(p - 1, e, " apr ", 5) == 0)
			{
				FoundMonth();
				found = TRUE;
			}
			else if (TextUtilsEx::strncmpUpper2Lower(p - 1, e, " aug ", 5) == 0)
			{
				FoundMonth();
				found = TRUE;
			}
			int deb = 1; break;
		}
		case 'd':
		{
			if (TextUtilsEx::strncmpUpper2Lower(p - 1, e, " dec ", 5) == 0)
			{
				FoundMonth();
				found = TRUE;
			}
			int deb = 1; break;
		}
		case 'f':
		{
			if (TextUtilsEx::strncmpUpper2Lower(p - 1, e, " feb ", 5) == 0)
			{
				FoundMonth();
				found = TRUE;
			}
			int deb = 1; break;
		}
		case 'j':
		{
			if (TextUtilsEx::strncmpUpper2Lower(p - 1, e, " jan ", 5) == 0)
			{
				FoundMonth();
				found = TRUE;
			}
			else if (TextUtilsEx::strncmpUpper2Lower(p - 1, e, " jun ", 5) == 0)
			{
				FoundMonth();
				found = TRUE;
			}
			else if (TextUtilsEx::strncmpUpper2Lower(p - 1, e, " jul ", 5) == 0)
			{
				FoundMonth();
				found = TRUE;
			}
			int deb = 1; break;
		}
		case 'm':
		{
			if (TextUtilsEx::strncmpUpper2Lower(p - 1, e, " mar ", 5) == 0)
			{
				FoundMonth();
				found = TRUE;
			}
			else if (TextUtilsEx::strncmpUpper2Lower(p - 1, e, " may ", 5) == 0)
			{
				FoundMonth();
				found = TRUE;
			}
			int deb = 1; break;
		}
		case 'n':
		{
			if (TextUtilsEx::strncmpUpper2Lower(p - 1, e, " nov ", 5) == 0)
			{
				FoundMonth();
				found = TRUE;
			}
			int deb = 1; break;
		}
		case 'o':
		{
			if (TextUtilsEx::strncmpUpper2Lower(p - 1, e, " oct ", 5) == 0)
			{
				FoundMonth();
				found = TRUE;
			}
			int deb = 1; break;
		}
		case 's':
		{
			if (TextUtilsEx::strncmpUpper2Lower(p - 1, e, " sep ", 5) == 0)
			{
				FoundMonth();
				found = TRUE;
			}
			int deb = 1; break;
		}
		default:
		{
			int deb = 1; break;
		}
		}
		if (found == TRUE)
		{
			char *d = TextUtilsEx::SkipWhiteReverse(p-1);
			char *day = TextUtilsEx::SkipNumericReverse(d);
			day++;
			if (_istdigit(*d))
			{
				TextUtilsEx::CopyLine(day, e, dateStr);
				int deb = 1;
			}
		}
		p++;
	}
	int deb = 1;
}

time_t MboxMail::parseRFC822Date(CStringA &date, CStringA &format)
{
	time_t tdate = -1;
	SYSTEMTIME tm;
	if (DateParser::parseRFC822Date(date, &tm))
	{
		MyCTime::fixSystemtime(&tm);
		if (DateParser::validateSystemtime(&tm))
		{
			MyCTime tt(tm);
			date = tt.FormatGmtTmA(format);
			if (date.IsEmpty())
				date = tt.FormatLocalTmA(format);
			if (date.IsEmpty())
				int deb = 1;
			tdate = tt.GetTime();
		}
	}
	return tdate;
}

BOOL MboxMail::CreateRFC822Date(CStringA &date, CStringA &rfcDate)
{
	rfcDate.Empty();

#if 0
	// Test
	//date = "2000/09/02";
	//date = "2000/13/02";
	date = "09/02/2000";
	date = "09/13/2000";
	date = "02/09/2000";
#endif

	CStringA delim = "/:- ";
	CStringArrayA va;
	TextUtilsEx::SplitStringA2A(date, delim, va);

	int val1_Int;
	int val2_Int;
	int val3_Int;

	if (va.GetCount() >= 3)
	{
		int fcnt1 = sscanf((LPCSTR)va[0], "%d", &val1_Int);
		int fcnt2 = sscanf((LPCSTR)va[1], "%d", &val2_Int);
		int fcnt3 = sscanf((LPCSTR)va[2], "%d", &val3_Int);

		if ((fcnt1 == 1) && (fcnt2 == 1) && (fcnt3 == 1))
		{
			if (val1_Int > 1000)
			{
				CStringA monthStr;
				if (val2_Int < 13)
				{
					if (val3_Int < 32)
					{
						MonthToString(val2_Int, monthStr);
						if (!monthStr.IsEmpty())
							rfcDate = va[2] + " " + monthStr + " " + va[0] + " 00:00:00 +0000 (UTC)";
					}
				}
				else // (val2_Int >= 13)
				{
					// very unlikely case
					if (val3_Int < 13)
					{
						MonthToString(val3_Int, monthStr);
						if (!monthStr.IsEmpty())
							rfcDate = va[1] + " " + monthStr + " " + va[0] + " 00:00:00 +0000 (UTC)";
					}
				}
			}
			else if (val3_Int > 1000)
			{
				CStringA monthStr;
				if (val2_Int < 13)
				{
					if (val1_Int < 32)
					{
						MonthToString(val2_Int, monthStr);
						if (!monthStr.IsEmpty())
							rfcDate = va[0] + " " + monthStr + " " + va[2] + " 00:00:00 +0000 (UTC)";
					}
				}
				else // (val2_Int >= 13)
				{
					// very unlikely case
					if (val1_Int < 13)
					{
						MonthToString(val1_Int, monthStr);
						if (!monthStr.IsEmpty())
							rfcDate = va[1] + " " + monthStr + " " + va[2] + " 00:00:00 +0000 (UTC)";
					}
				}
			}
		}
	}

	if (!rfcDate.IsEmpty())
		return TRUE;
	else
		return FALSE;
}

int MboxMail::NormalizeText(MboxMail* m)
{
	UINT CP_ASCII = 20127;

	DWORD error;
	CStringA textUTF8;
	CStringW text1W;
	CStringW text2W;
	BOOL retStr2WStr = TRUE;
	BOOL retStr2UTF8 = TRUE;
	BOOL retUTF82WStr = TRUE;


	//if ((m->m_subj_charsetId != 0) && (m->m_subj_charsetId != CP_UTF8) && (m->m_subj_charsetId != CP_ASCII) && m->m_subj.GetLength())
	if ((m->m_subj_charsetId != CP_UTF8) && m->m_subj.GetLength())
	{
		retStr2UTF8 = TextUtilsEx::Str2UTF8((LPCSTR)m->m_subj, m->m_subj.GetLength(), m->m_subj_charsetId, textUTF8, error);
#ifdef _DEBUG
		retStr2WStr = TextUtilsEx::Str2WStr(m->m_subj, m->m_subj_charsetId, text1W, error);
		retUTF82WStr = TextUtilsEx::UTF82WStr(&textUTF8, &text2W, error);

		_ASSERTE(text1W.Compare(text2W) == 0);
#endif

		m->m_subj = textUTF8;
		m->m_subj_charsetId = CP_UTF8;
	}

	//if ((m->m_from_charsetId != 0) && (m->m_from_charsetId != CP_UTF8) && (m->m_from_charsetId != CP_ASCII) && m->m_from.GetLength())
	if ((m->m_from_charsetId != CP_UTF8) && m->m_from.GetLength())
	{
		retStr2UTF8 = TextUtilsEx::Str2UTF8((LPCSTR)m->m_from, m->m_from.GetLength(), m->m_from_charsetId, textUTF8, error);
#ifdef _DEBUG
		retStr2WStr = TextUtilsEx::Str2WStr(m->m_from, m->m_from_charsetId, text1W, error);
		retUTF82WStr = TextUtilsEx::UTF82WStr(&textUTF8, &text2W, error);

		_ASSERTE(text1W.Compare(text2W) == 0);
#endif

		m->m_from = textUTF8;
		m->m_from_charsetId = CP_UTF8; // FIXME
	}

	//if ((m->m_to_charsetId != 0) && (m->m_to_charsetId != CP_UTF8) && (m->m_to_charsetId != CP_ASCII) && m->m_to.GetLength())
	if ((m->m_to_charsetId != CP_UTF8) && m->m_to.GetLength())
	{
		retStr2UTF8 = TextUtilsEx::Str2UTF8((LPCSTR)m->m_to, m->m_to.GetLength(), m->m_to_charsetId, textUTF8, error);
#ifdef _DEBUG
		retStr2WStr = TextUtilsEx::Str2WStr(m->m_to, m->m_to_charsetId, text1W, error);
		retUTF82WStr = TextUtilsEx::UTF82WStr(&textUTF8, &text2W, error);

		_ASSERTE(text1W.Compare(text2W) == 0);
#endif

		m->m_to = textUTF8;
		m->m_to_charsetId = CP_UTF8;
	}

	//if ((m->m_cc_charsetId != 0) && (m->m_cc_charsetId != CP_UTF8) && (m->m_cc_charsetId != CP_ASCII) && m->m_cc.GetLength())
	if ((m->m_cc_charsetId != CP_UTF8) && m->m_cc.GetLength())
	{
		retStr2UTF8 = TextUtilsEx::Str2UTF8((LPCSTR)m->m_cc, m->m_cc.GetLength(), m->m_cc_charsetId, textUTF8, error);
#ifdef _DEBUG
		retStr2WStr = TextUtilsEx::Str2WStr(m->m_cc, m->m_cc_charsetId, text1W, error);
		retUTF82WStr = TextUtilsEx::UTF82WStr(&textUTF8, &text2W, error);

		_ASSERTE(text1W.Compare(text2W) == 0);
#endif

		m->m_cc = textUTF8;
		m->m_cc_charsetId = CP_UTF8;
	}

	//if ((m->m_bcc_charsetId != 0) && (m->m_bcc_charsetId != CP_UTF8) && (m->m_bcc_charsetId != CP_ASCII) && m->m_bcc.GetLength())
	if ((m->m_bcc_charsetId != CP_UTF8) && m->m_bcc.GetLength())
	{
		retStr2UTF8 = TextUtilsEx::Str2UTF8((LPCSTR)m->m_bcc, m->m_bcc.GetLength(), m->m_bcc_charsetId, textUTF8, error);
#ifdef _DEBUG
		retStr2WStr = TextUtilsEx::Str2WStr(m->m_bcc, m->m_bcc_charsetId, text1W, error);
		retUTF82WStr = TextUtilsEx::UTF82WStr(&textUTF8, &text2W, error);

		_ASSERTE(text1W.Compare(text2W) == 0);
#endif

		m->m_bcc = textUTF8;
		m->m_bcc_charsetId = CP_UTF8;
	}
	return 1;
}

char szFrom5[] = "From ";
char szFrom6[] = "\nFrom ";
char	*g_szFrom;
int		g_szFromLen;

bool MboxMail::Process(ProgressTimer& progressTimer, register char *p, DWORD bufSize, _int64 startOffset,  bool bFirstView, bool bLastView, _int64 &lastStartOffset, bool bEml, _int64 &msgOffset, CString &statusText, BOOL parseContent)
{
	static const char *cFromMailBegin = "From ";
	static const int cFromMailBeginLen = istrlen(cFromMailBegin);
	static const char *cFrom = "from:";
	static const int cFromLen = istrlen(cFrom);
	static const char *cTo = "to:";
	static const int cToLen = istrlen(cTo);
	static const char *cCC = "cc:";
	static const int cCCLen = istrlen(cCC);
	static const char *cBCC = "bcc:";
	static const int cBCCLen = istrlen(cBCC);
	static const char *cSubject = "subject:";
	static const int cSubjectLen = istrlen(cSubject);
	static const char *cDate = "date:";
	static const int cDateLen = istrlen(cDate);
	static const char *cReceived = "received:";
	static const int cReceivedLen = istrlen(cReceived);
	static const char *cContentType = "content-type:";  // some mails such as from Apple don't seem to have blank line to separate header from body
	static const int cContentTypeLen = istrlen(cContentType);
	static const char *cThreadId = "x-gm-thrid:";
	static const int cThreadIdLen = istrlen(cThreadId);
	static const char *cMsgId = "message-id:";
	static const int cMsgIdLen = istrlen(cMsgId);
	static const char *cReferenceId = "references:";
	static const int cReferenceIdLen = istrlen(cReferenceId);

	char *orig = p;
	register char *e = p + bufSize - 1;
	if (p == NULL || bufSize < 10) {
		return false;
	}

	CString section_options = CString(sz_Software_mboxview) + L"\\Options";

	CStringA contentDisposition = "Content-Disposition: attachment";

	int iFormat = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"format");

	CStringA format = MboxMail::GetDateFormat(iFormat);
	CStringA to, from, subject, date, date_orig;
	CStringA date_fromField;
	CStringA cc, bcc;
	CStringA msgId, thrdId, referenceIds;
	time_t tdate = -1;
	time_t tdate_fromField = 1;
	bool	bTo = true, bFrom = true, bSubject = true, bDate = true, bRcvDate = true; // indicates not found, false means found
	bool bMsgId = true, bThrdId = true, bReferenceIds = true;
	bool bCC = true, bBCC = true;
	char *msgStart = NULL;
	int recv = TRUE;
	int curstep = (int)(startOffset / (s_step?s_step:1));
	CStringA line;
	CStringA rcved;
	CStringA dateStr;
	DWORD error;

	int mailCnt = s_mails.GetCount();
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
	//wchar_t szCauseBuffer[256];

	char *stackDumpFileName = "Exception_StackTrace.txt";
	char *mailDumpFileName = "Exception_MailDump.txt";
	char *exceptionName = "UnknownException";
	UINT seNumb = 0;
	char *szCause = "Unknown";
	BOOL exceptionThrown = FALSE;
	BOOL exceptionDone = FALSE;

	char *pat = "From:  XXXX";
	int patLen = istrlen(pat);

	ULONGLONG tc_start = GetTickCount64();

	BOOL headerDone = FALSE;
	while (p < (e - 4))   // TODO: why (e - 4) ?? chaneg 4 -> 5
	{
		{
			// EX_TEST
#if 0
			if ((s_mails.GetCount() == 50) && (exceptionDone == FALSE))
			{
				int i = 0; 
				//int Ex = rand() / i;   // This will throw a SE (divide by zero).
				//char* szTemp = (char*)1;
				//strcpy_s(szTemp, 1000, "A");
				*badPtr = 'a';
				SCODE sc = 99;
				//AfxThrowOleException(sc);
				MboxMail *m = s_mails[100000];
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
							date_fromField = tt.FormatGmtTmA(format);
							if (date_fromField.IsEmpty())
								date_fromField = tt.FormatLocalTmA(format);
							if (date_fromField.IsEmpty())
								int deb = 1;
							tdate_fromField = tt.GetTime();
						}
					}

					msgStart = p;
					msgOffset = startOffset + (_int64)(msgStart - orig);
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
								date_fromField = tt.FormatGmtTmA(format);
								if (date_fromField.IsEmpty())
									date_fromField = tt.FormatLocalTmA(format);
								if (date_fromField.IsEmpty())
									int deb = 1;
								tdate_fromField = tt.GetTime();
							}
						}

						if (!from.IsEmpty() || !to.IsEmpty() ||  !subject.IsEmpty() ||  !date.IsEmpty())
						{

							if (subject.Find("Dostosuj Gmaila") >= 0)
								int deb = 1;

							UINT toCharacterId = 0;  // not used anyway
							MboxMail *m = new MboxMail();
							m->m_startOff = startOffset + (_int64)(msgStart - orig);
							m->m_length = IntPtr2Int(p - msgStart);
							m->m_to = TextUtilsEx::DecodeString(to, m->m_to_charset, m->m_to_charsetId, toCharacterId, error);
							m->m_from = TextUtilsEx::DecodeString(from, m->m_from_charset, m->m_from_charsetId, toCharacterId, error);
							m->m_subj = TextUtilsEx::DecodeString(subject, m->m_subj_charset, m->m_subj_charsetId, toCharacterId, error);
							m->m_subj.Trim();
							m->m_cc = TextUtilsEx::DecodeString(cc, m->m_cc_charset, m->m_cc_charsetId, toCharacterId, error);
							m->m_bcc = TextUtilsEx::DecodeString(bcc, m->m_bcc_charset, m->m_bcc_charsetId, toCharacterId, error);
							if (!bcc.IsEmpty())
								int deb = 1;
							if (!cc.IsEmpty())
								int deb = 1;
							m->m_subj.Replace("\r", "");
							m->m_subj.Replace("\n", "");
							m->m_timeDate = tdate;
							m->m_recv = recv;
							m->m_headLength = 0;  // set in ParseContent()
							m->m_messageId = msgId;
							m->m_threadId = thrdId;

							//m->m_crc32 = TextUtilities::CalcCRC32(msgStart, m->m_length);

							char *bodyStart = msgStart;
							char *bodyEnd = msgStart + m->m_length;
							if (parseContent)
								char *nextStart = ParseContent(m, bodyStart, bodyEnd);
							if (m->m_ContentDetailsArray.size() == 0) {
								int deb = 1;
							}

							s_mails.Add(m);

							int retN = MboxMail::NormalizeText(m);

							if (m->m_messageId == m->m_replyId)
								int deb = 1;


							int index = s_mails.GetCount() - 1;

							if (s_mails.GetCount() == 1098)
								int deb = 1;

							UINT_PTR dwProgressbarPos = 0;
							ULONGLONG workRangePos = startOffset + (p - orig);
							BOOL needToUpdateStatusBar = progressTimer.UpdateWorkPos(workRangePos, dwProgressbarPos);
							if (needToUpdateStatusBar)
							{
								CString mailNum;
								if (statusText.IsEmpty())
								{
									CString fmt = L"Parsing archive file to create index file ... %d";
									ResHelper::TranslateString(fmt);
									mailNum.Format(fmt, s_mails.GetCount());
								}
								else
								{
									CString fmt = L"%s ... %d";
									ResHelper::TranslateString(fmt);
									mailNum.Format(fmt, statusText, s_mails.GetCount());
								}

								if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(mailNum, (UINT_PTR)(dwProgressbarPos));

								tc_start = GetTickCount64();

								int debug = 1;
							}

							if ((s_mails.GetCount() % 100000) == 1024)
							{
								int deb = 1;
							}

							if (pCUPDUPData && pCUPDUPData->ShouldTerminate())
							{
								break;
							}
						}
					}
					msgStart = p;
					msgOffset = startOffset + (_int64)(msgStart - orig);
					to.Empty();
					from.Empty();
					subject.Empty();
					date.Empty();
					cc.Empty();
					bcc.Empty();
					msgId.Empty();
					thrdId.Empty();
					tdate = -1;
					bTo = bFrom = bSubject = bDate = bRcvDate = bCC = bBCC = bMsgId = bThrdId = bReferenceIds = true;
					//Sleep(10);
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

					if ((tdate = MboxMail::parseRFC822Date(date, format)) > 0)
					{
						bRcvDate = false;
						recv = TRUE;
					}
					else
					{
						date_orig = date;

						date.Empty();

						_int64 lsize = bufSize - (_int64)(msgStart - orig);
						int datalen;
						if (lsize > 0x1fffffff)
							datalen = 0x1fffffff;
						else
							datalen = (int)lsize;

						dateStr.Empty();
						FindDateInHeader(msgStart, datalen, dateStr);

						if (dateStr.IsEmpty())
						{
							CStringA rfcDateStr;
							if (CreateRFC822Date(date_orig, rfcDateStr))
							{
								date = rfcDateStr;
								if ((tdate = MboxMail::parseRFC822Date(date, format)) > 0)
								{
									bDate = false;
									recv = FALSE;
								}
								else
									date.Empty();
							}
						}
						else   //  (!dateStr.IsEmpty())
						{
							date = dateStr;
							if ((tdate = MboxMail::parseRFC822Date(date, format)) > 0)
							{
								bDate = false;
								recv = FALSE;
							}
							else // if (dateStr.IsEmpty())
							{
								date.Empty();

								CStringA rfcDateStr;
								if (CreateRFC822Date(date_orig, rfcDateStr))
								{
									date = rfcDateStr;
									if ((tdate = MboxMail::parseRFC822Date(date, format)) > 0)
									{
										bDate = false;
										recv = FALSE;
									}
									else
										date.Empty();
								}
								int deb = 1;
							}
						}
					}
				}
				else if (bDate && TextUtilsEx::strncmpUpper2Lower(p, e, cDate, cDateLen) == 0)
				{
					p = MimeParser::GetMultiLine(p, e, line);
					date = line.Mid(cDateLen);
					date.Trim();

					if ((tdate = MboxMail::parseRFC822Date(date, format)) > 0)
					{
						bDate = false;
						recv = FALSE;
					}
					else
					{
						date_orig = date;

						date.Empty();

						_int64 lsize = bufSize - (_int64)(msgStart - orig);
						int datalen;
						if (lsize > 0x1fffffff)
							datalen = 0x1fffffff;
						else
							datalen = (int)lsize;

						dateStr.Empty();
						FindDateInHeader(msgStart, datalen, dateStr);

						if (dateStr.IsEmpty())
						{
							CStringA rfcDateStr;
							if (CreateRFC822Date(date_orig, rfcDateStr))
							{
								date = rfcDateStr;
								if ((tdate = MboxMail::parseRFC822Date(date, format)) > 0)
								{
									bDate = false;
									recv = FALSE;
								}
								else
									date.Empty();
							}
						}
						else   //  (!dateStr.IsEmpty())
						{
							date = dateStr;
							if ((tdate = MboxMail::parseRFC822Date(date, format)) > 0)
							{
								bDate = false;
								recv = FALSE;
							}
							else // if (dateStr.IsEmpty())
							{
								date.Empty();

								CStringA rfcDateStr;
								if (CreateRFC822Date(date_orig, rfcDateStr))
								{
									date = rfcDateStr;
									if ((tdate = MboxMail::parseRFC822Date(date, format)) > 0)
									{
										bDate = false;
										recv = FALSE;
									}
									else
										date.Empty();
								}
								int deb = 1;
							}
						}
						int deb = 1;
					}
				}
				else if (bMsgId && TextUtilsEx::strncmpUpper2Lower(p, e, cMsgId, cMsgIdLen) == 0)
				{
					bMsgId = false;
					p = MimeParser::GetMultiLine(p, e, line);
					MimeParser::GetMessageId(line, cMsgIdLen, msgId);
#if 0
					// Support code for testing
					int startPos = cMsgIdLen;
					MessageIdList msgIdList;
					int msgIdCnt = MimeParser::GetMessageIdList(line, startPos, msgIdList);
					if (msgIdList.size())
					{
						CString& messageId = msgIdList.front();
						if (msgId.Compare(messageId) != 0)
						{
							int deb = 1;
						}
					}
#endif
				}
				else if (bThrdId && TextUtilsEx::strncmpUpper2Lower(p, e, cThreadId, cThreadIdLen) == 0)
				{
					bThrdId = false;
					p = MimeParser::GetMultiLine(p, e, line);
					thrdId = line.Mid(cThreadIdLen);
					thrdId.Trim();
				}
#if 0
				// TODO: For future support of Refernces header field
				else if (bReferenceIds && TextUtilsEx::strncmpUpper2Lower(p, e, cReferenceId, cReferenceIdLen) == 0)
				{
					bReferenceIds = false;
					p = MimeParser::GetMultiLine(p, e, line);
					referenceIds = line.Mid(cReferenceIdLen);
					referenceIds.Trim();
					//
					int startPos = cReferenceIdLen;
					MessageIdList msgIdList;
					int msgIdCnt = MimeParser::GetMessageIdList(line, startPos, msgIdList);
					if (msgIdList.size())
					{
						CString & messageId = msgIdList.front();
						CString message_Id;
						MimeParser::GetMessageId(line, cReferenceIdLen, message_Id);
						if (message_Id.Compare(messageId) != 0)
						{
							int deb = 1;
						}
					}
					int deb = 1;
				}
#endif
#if 1
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
				}
#endif
				else
				{
					bool isEmpty = false;
					char *psave = p;
					p = MimeParser::EatNewLine(p, e, isEmpty);
					// TODO:  This check may not completely reliable and may need better check and likley more expensive check
					// Without the end of header check, we could pickup CC and BCC fiekds from mail text
					// but it could be worst if we stop looking for headear fields prematuraly
					if (isEmpty)
					{
						if (from.IsEmpty() && to.IsEmpty() && date.IsEmpty())
							int deb = 1;

						if (from.IsEmpty() || to.IsEmpty() || date.IsEmpty())
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

			{
				// EX_TEST
#if 0
				//if (s_mails.GetCount() == 200)
				{
					//int i = 0;  int Ex = rand() / i;   // This will throw a SE (divide by zero).
					char* szTemp = (char*)1;
					//strcpy_s(szTemp, 1000, "A");
					//*badPtr = 'a';
					SCODE sc = 99;
					//AfxThrowOleException(sc);
				}
#endif
				if (!from.IsEmpty() || !to.IsEmpty() || !subject.IsEmpty() || !date.IsEmpty())
				{
					UINT toCharacterId = 0;  // not used anyway
					MboxMail *m = new MboxMail();
					//		TRACE(L"start: %d length: %d\n", msgStart - orig, p - msgStart);
					m->m_startOff = startOffset + (_int64)(msgStart - orig);
					m->m_length = IntPtr2Int(p - msgStart);
					m->m_to = TextUtilsEx::DecodeString(to, m->m_to_charset, m->m_to_charsetId, toCharacterId, error);
					m->m_from = TextUtilsEx::DecodeString(from, m->m_from_charset, m->m_from_charsetId, toCharacterId, error);
					m->m_subj = TextUtilsEx::DecodeString(subject, m->m_subj_charset, m->m_subj_charsetId, toCharacterId, error);
					m->m_subj.Trim();
					m->m_cc = TextUtilsEx::DecodeString(cc, m->m_cc_charset, m->m_cc_charsetId, toCharacterId, error);
					m->m_bcc = TextUtilsEx::DecodeString(bcc, m->m_bcc_charset, m->m_bcc_charsetId, toCharacterId, error);
					if (!bcc.IsEmpty())
						int deb = 1;
					if (!cc.IsEmpty())
						int deb = 1;
					m->m_subj.Replace("\r", "");
					m->m_subj.Replace("\n", "");
					m->m_recv = recv;
					m->m_timeDate = tdate;
					m->m_headLength = 0; // set in ParseContent()
					m->m_messageId = msgId;
					m->m_threadId = thrdId;

					//m->m_crc32 = TextUtilities::CalcCRC32(msgStart, m->m_length);

					MailBodyContent body;
					char *bodyStart = msgStart;
					char *bodyEnd = msgStart + m->m_length;
					if (parseContent)
						char *nextStart = ParseContent(m, bodyStart, bodyEnd);

					s_mails.Add(m);

					int retN = MboxMail::NormalizeText(m);

					int index = s_mails.GetCount() - 1;
					//m->DumpMailBox(m, index);
					//NListView::DumpItemDetails(index);

					if (s_mails.GetCount() == 1098)
						int deb = 1;
				}
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

void MboxMail::Parse(LPCWSTR path)
{
	Destroy();

	s_mails.SetSize(100000, 100000);
	s_mails.SetCountKeepData(0);

	// Delete all files in print and image cache directories
	CString cpath = path;
	CString rootPrintSubFolder = L"PrintCache";;
	CString targetPrintSubFolder;
	CString prtCachePath;
	CString errorText;
#if 0
	// Removing Printcache not critical I think
	bool ret1 = MboxMail::GetArchiveSpecificCachePath(cpath, rootPrintSubFolder, targetPrintSubFolder, prtCachePath, errorText);
	if (errorText.IsEmpty() && FileUtils::PathDirExists(prtCachePath)) {
		pCUPDUPData->SetProgress(L"Deleting all files in the PrintCache directory ...", 0);
		FileUtils::RemoveDir(prtCachePath, true);
	}
#endif

	rootPrintSubFolder = L"ImageCache";
	targetPrintSubFolder.Empty();
	CString imageCachePath;
	errorText.Empty();
	bool ret2 = MboxMail::GetCachePath(rootPrintSubFolder, targetPrintSubFolder, imageCachePath, errorText, &cpath);
	if (errorText.IsEmpty() && FileUtils::PathDirExists(imageCachePath))
	{
		CString txt = L"Deleting all related files in the ImageCache directory ...";
		ResHelper::TranslateString(txt);
		pCUPDUPData->SetProgress(txt, 0);


		FileUtils::RemoveDir(imageCachePath, true);
	}

	//MboxMail::s_path = path;
	MboxMail::SetMboxFilePath(CString(path));
	CString lastPath = MboxMail::GetLastPath();
	CString lastDataPath = MboxMail::GetLastDataPath();

	bool	bEml = false;
	int l = iwstrlen(path);
	if (l > 4 && _wcsnicmp(path + l - 4, L".eml", 4) == 0)
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
	ULONGLONG tc = GetTickCount64();
#endif
	g_szFrom = szFrom5;
	g_szFromLen = 5;
	TRACE(L"fSize = %lld\n", fSize);

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	CString mailFilePath = path;
	FileUtils::SplitFilePath(mailFilePath, driveName, directory, fileNameBase, fileNameExtention);

	CString mailFile = fileNameBase + fileNameExtention;
	CString parsingFileText = L"Parsing \"" + mailFile + L"\" ...";

	//pCUPDUPData->SetProgress(parsingFileText, 0);  // works but doesn't always fit into progress bar

	CString txt = L"Parsing archive file to create index file ...";
	ResHelper::TranslateString(txt);
	pCUPDUPData->SetProgress(txt, 0);
	// TODO: due to breaking the file into multiple chunks, it looks some emails can be lost : Fixed
	bool firstView = true;;
	bool lastView = false;
	DWORD bufSize = 0;
	_int64 aligned_offset = 0;
	_int64 delta = 0;

	// TODO:  Need to consider to redo reading and create  Character Stream for reading
	ULONGLONG workRangeFirstPos = 0;
	ULONGLONG workRangeLastPos = fSize - 1;
	ProgressTimer progressTimer(workRangeFirstPos, workRangeLastPos);

	CString statusText;
	_int64 lastStartOffset = 0;
	while  ((lastView == false) && !pCUPDUPData->ShouldTerminate()) 
	{
		s_curmap = lastStartOffset;
		aligned_offset = (s_curmap / dwAllocationGranularity) * dwAllocationGranularity;
		delta = s_curmap - aligned_offset;
		bufSize = ((fSize - aligned_offset) < mappingSize) ? (DWORD)(fSize - aligned_offset) : (DWORD)mappingSize;

		TRACE(L"offset=%lld, bufsize=%ld, fSize-curmap=%lld, end=%lld\n", s_curmap, bufSize, fSize - s_curmap, s_curmap + bufSize);
		char * pview = (char *)MapViewOfFileEx(hFileMap, FILE_MAP_READ, (DWORD)(aligned_offset >> 32), (DWORD)aligned_offset, bufSize, NULL);

		if (pview == 0)
		{
			DWORD err = GetLastError();
			CString txt = L"Could not finish parsing due to memory fragmentaion. Please restart the mbox viewer to resolve.";
			ResHelper::TranslateString(txt);

			HWND h = NULL; // we don't have any window yet ??
			int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
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
		_int64 msgOffset = viewOffset;
		BOOL processException = FALSE;
#ifdef USE_STACK_WALKER
		try
		{
#endif
			MboxMail::Process(progressTimer, p, viewBufSize, viewOffset, firstView, lastView, lastStartOffset, bEml, msgOffset, statusText);
#ifdef USE_STACK_WALKER
		}
		catch (...)
		{
			processException = TRUE;
		}
#endif
		firstView = false;
		UnmapViewOfFile(p);

#ifdef USE_STACK_WALKER
		if (processException)
		{
			DumpMailParseException(msgOffset);
			exit(0);
			//AfxAbort();
			//break;
		}
#endif
	}
#ifdef _DEBUG
	tc = (GetTickCount64() - tc);
	CString out;
	out.Format(L"Parse Took %d:%d %d\n", tc / 60000, (tc / 1000) % 60, tc);
	ResHelper::MyTrace(out);
	out.Format(L"Found %d mails\n", MboxMail::s_mails.GetSize());
	ResHelper::MyTrace(out);
//	TRACE(L"Parse Took %d:%d %d\n", tc / 60000, (tc / 1000) % 60, tc);
//	TRACE(L"Found %d mails\n", MboxMail::s_mails.GetSize());
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


// This is basically the same as MboxMail::Parse().
// It parses and find all messages wothout parsing mail message content
// Used during  merging of all mbox files in root and subfolders
void MboxMail::Parse_LabelView(LPCWSTR path)
{
	MboxMail::SetMboxFilePath(CString(path));
	CString lastPath = MboxMail::GetLastPath();
	CString lastDataPath = MboxMail::GetLastDataPath();

	bool	bEml = false;

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

	int mappingsInFile = (int)((fSize - 1) / mappingSize) + 1;
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
	ULONGLONG tc = GetTickCount64();
#endif
	g_szFrom = szFrom5;
	g_szFromLen = 5;
	TRACE(L"fSize = %lld\n", fSize);

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	CString mailFilePath = path;
	FileUtils::SplitFilePath(mailFilePath, driveName, directory, fileNameBase, fileNameExtention);

	CString mailFile = fileNameBase + fileNameExtention;
#if 0
	CString parsingFileText = L"Parsing \"" + mailFile + L"\"";
#endif

	CString fmt = L"Parsing \"%s\"";
	ResHelper::TranslateString(fmt);
	CString parsingFileText;
	parsingFileText.Format(fmt, mailFile);

	if (pCUPDUPData)
		pCUPDUPData->SetProgress(parsingFileText, 0);  // works but doesn't always fit into progress bar
	//pCUPDUPData->SetProgress(L"Parsing archive file to create index file ...", 0);

	// TODO: due to breaking the file into multiple chunks, it looks some emails can be lost : Fixed

	bool firstView = true;;
	bool lastView = false;
	DWORD bufSize = 0;
	_int64 aligned_offset = 0;
	_int64 delta = 0;

	// TODO:  Need to consider to redo reading and create  Character Stream for reading
	ULONGLONG workRangeFirstPos = 0;
	ULONGLONG workRangeLastPos = fSize - 1;
	ProgressTimer progressTimer(workRangeFirstPos, workRangeLastPos);

	_int64 lastStartOffset = 0;
	while ((lastView == false) && !(pCUPDUPData && pCUPDUPData->ShouldTerminate()))
	{
		s_curmap = lastStartOffset;
		aligned_offset = (s_curmap / dwAllocationGranularity) * dwAllocationGranularity;
		delta = s_curmap - aligned_offset;
		bufSize = ((fSize - aligned_offset) < mappingSize) ? (DWORD)(fSize - aligned_offset) : (DWORD)mappingSize;

		TRACE(L"offset=%lld, bufsize=%ld, fSize-curmap=%lld, end=%lld\n", s_curmap, bufSize, fSize - s_curmap, s_curmap + bufSize);
		char * pview = (char *)MapViewOfFileEx(hFileMap, FILE_MAP_READ, (DWORD)(aligned_offset >> 32), (DWORD)aligned_offset, bufSize, NULL);

		if (pview == 0)
		{
			DWORD err = GetLastError();
			CString txt = L"Could not finish parsing due to memory fragmentaion. Please restart the mbox viewer to resolve.";
			ResHelper::TranslateString(txt);

			HWND h = NULL; // we don't have any window yet ??
			int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
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
		_int64 msgOffset = viewOffset;
		BOOL processException = FALSE;
#ifdef USE_STACK_WALKER
		try
		{
#endif
			MboxMail::Process(progressTimer, p, viewBufSize, viewOffset, firstView, lastView, lastStartOffset, bEml, msgOffset, parsingFileText, FALSE);
#ifdef USE_STACK_WALKER
		}
		catch (...)
		{
			processException = TRUE;
		}
#endif
		firstView = false;
		UnmapViewOfFile(p);

#ifdef USE_STACK_WALKER
		if (processException)
		{
			DumpMailParseException(msgOffset);
			exit(0);
			//AfxAbort();
			//break;
		}
#endif
	}

#ifdef _DEBUG
	tc = (GetTickCount64() - tc);
	CString out;
	out.Format(L"Parse Took %d:%d %d\n", tc / 60000, (tc / 1000) % 60, tc);
	ResHelper::MyTrace(out);
	out.Format(L"Found %d mails\n", MboxMail::s_mails.GetSize());
	ResHelper::MyTrace(out);
	//	TRACE(L"Parse Took %d:%d %d\n", tc / 60000, (tc / 1000) % 60, tc);
	//	TRACE(L"Found %d mails\n", MboxMail::s_mails.GetSize());
#endif

	CloseHandle(hFileMap);
	CloseHandle(hFile);
}


#include <algorithm>

bool sortByDate(MboxMail *cr1, MboxMail *cr2)
{
	return (cr1->m_timeDate < cr2->m_timeDate);
}
bool sortByDateDesc(MboxMail *cr1, MboxMail *cr2)
{
	return (cr1->m_timeDate > cr2->m_timeDate);
}

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

char* ExtractSubjectEx(char *subj, int subjlen)
{
	// Find Re and/or Fwd and return pointer to subj left
	return 0;
}

int ExtractSubject(CStringA &subject, char *&subjectText) 
{
	// TODO: Enhance to handle more cases: Re[2] Re: [hello] Re: etc
	char *subj = (char *)(LPCSTR)subject;
	int subjlen = subject.GetLength();

	//TODO: should I skip white spaces 
	// should compare
	while (subjlen >= 4)
	{
		if (subj[0] == 'R')
		{
			if ((strncmp((char*)subj, "Re: ", 4) == 0) || (strncmp((char*)subj, "RE: ", 4) == 0))
			{
				subj += 4; subjlen -= 4;
				if (subj[0] == '[')
				{
					char* pch;
					if (pch = strchr(subj, ']'))
					{
						int r1 = strncmp((char*)pch, "] Re: ", 6);
						int r2 = strncmp((char*)pch, "] RE: ", 6);
						if ((r1 == 0) || (r2 == 0))
						{
							subjlen -= (int)((pch - subj) + 6);
							subj = pch + 6;
						}
						else
						{
							int r1 = strncmp((char*)pch, "] Fwd: ", 7);
							int r2 = strncmp((char*)pch, "] FWD: ", 7);
							if ((r1 == 0) || (r2 == 0))
							{
								subjlen -= (int)((pch - subj) + 7);
								subj = pch + 7;
							}
						}
					}
				}
			}
			else
				break;
		}
		else if (subj[0] == 'F')
		{
			if (subjlen >= 5)
			{
				if ((strncmp((char*)subj, "Fwd: ", 5) == 0) || (strncmp((char*)subj, "FWD: ", 5) == 0)) {
					subj += 5; subjlen -= 5;
					if (subj[0] == '[')
					{
						char* pch;
						if (pch = strchr(subj, ']'))
						{
							int r1 = strncmp((char*)pch, "] Fwd: ", 7);
							int r2 = strncmp((char*)pch, "] FWD: ", 7);
							if ((r1 == 0) || (r2 == 0))
							{
								subjlen -= (int)((pch - subj) + 7);
								subj = pch + 7;
							}
							else
							{
								int r1 = strncmp((char*)pch, "] Re: ", 6);
								int r2 = strncmp((char*)pch, "] RE: ", 6);
								if ((r1 == 0) || (r2 == 0))
								{
									subjlen -= (int)((pch - subj) + 6);
									subj = pch + 6;
								}
							}
						}
					}

				}
				else
					break;
			}
			else
				break;
		}
		else
			break;
	}

	subjectText = subj;
	return subjlen;
}


bool sortBySubjectAsc(MboxMail *cr1, MboxMail *cr2)
{
	char *subj1;
	char *subj2;
	int subjlen1 = ExtractSubject(cr1->m_subj, subj1);
	int subjlen2 = ExtractSubject(cr2->m_subj, subj2);

	if ((subjlen1 == subjlen2) && (strncmp((char*)subj1, (char*)subj2, subjlen1) == 0))
	{
		return (cr1->m_timeDate < cr2->m_timeDate);
	}
	else
	{
		return (strcmp((char*)subj1, (char*)subj2) < 0);
	}
}

bool sortBySubjectDesc(MboxMail *cr1, MboxMail *cr2)
{
	char *subj1;
	char *subj2;
	int subjlen1 = ExtractSubject(cr1->m_subj, subj1);
	int subjlen2 = ExtractSubject(cr2->m_subj, subj2);

	if ((subjlen1 == subjlen2) && (strncmp((char*)subj1, (char*)subj2, subjlen1) == 0))
	{
		return (cr1->m_timeDate < cr2->m_timeDate);
	}
	else
	{
		return  (strcmp((char*)subj1, (char*)subj2) > 0);
	}
}

bool CompBySubject(MboxMail *cr1, MboxMail *cr2)
{
	char *subj1;
	char *subj2;
	int subjlen1 = ExtractSubject(cr1->m_subj, subj1);
	int subjlen2 = ExtractSubject(cr2->m_subj, subj2);

	return (strcmp((char*)subj1, (char*)subj2) == 0);
}

int CompareBySubject(MboxMail *cr1, MboxMail *cr2)
{
	char *subj1;
	char *subj2;
	int subjlen1 = ExtractSubject(cr1->m_subj, subj1);
	int subjlen2 = ExtractSubject(cr2->m_subj, subj2);

	return strcmp((char*)subj1, (char*)subj2);
}

bool sortByGroupId(MboxMail *cr1, MboxMail *cr2) {
	if (cr1->m_groupId == cr2->m_groupId)
		return (cr1->m_timeDate < cr2->m_timeDate); 
	else
		return (cr1->m_groupId < cr2->m_groupId);
}
bool sortByGroupIdDesc(MboxMail *cr1, MboxMail *cr2) {
	if (cr1->m_groupId == cr2->m_groupId)
		return (cr1->m_timeDate < cr2->m_timeDate); 
	else
		return (cr1->m_groupId > cr2->m_groupId);
}

bool sortBySize(MboxMail *cr1, MboxMail *cr2) {
	if (cr1->m_length == cr2->m_length)
		return (cr1->m_timeDate < cr2->m_timeDate);
	else
		return (cr1->m_length < cr2->m_length);
}
bool sortBySizeDesc(MboxMail *cr1, MboxMail *cr2) {
	if (cr1->m_length == cr2->m_length)
		return (cr1->m_timeDate < cr2->m_timeDate);
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
	std::sort(s_m->GetData(), s_m->GetData() + s_m->GetSize(), bDesc ? sortBySubjectDesc : sortBySubjectAsc);
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
	int arraySize = (int)s_mails_ref.GetSize();
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

int CheckIf8Bit( const char * data, int datalen)
{
	unsigned char *p = (unsigned char*)data;
	unsigned char c;
	int i;
	for (i = 0; i < datalen; i++)
	{
		c = *p++;;
		if (c > 0x7f)
			return 1;
	}
	return 0;
}

char * MboxMail::ParseContent(MboxMail *mail, char *startPos, char *endPos)
{
	int bodyCnt = 0;

	MboxMail *m = mail;

	int mailLength = m->m_length;
	int mLength = (int)(endPos - startPos); // mailLength and mLength should be equal

	char *mailBegin = startPos;
	char *bodyDataOrigin = mailBegin;

#if 0
	// helper when testing root merging
	MailBody *mBody = MailBody::CreateMailBody();
	int nSize = mBody->MailHeader::Load(mailBegin, mailLength);
	if (nSize <= 0)
		int deb = 0;

	m->m_messageId = mBody->m_MessageId;
	m->m_replyId = mBody->m_InReplyId;
	m->m_threadId = mBody->m_ThreadId;

#else
	MailBody *mBody = MailBody::CreateMailBody();
	mBody->Load(bodyDataOrigin, mailBegin, mailLength);

	m->m_messageId = mBody->m_MessageId;
	m->m_replyId = mBody->m_InReplyId;
	m->m_threadId = mBody->m_ThreadId;

#if 1
	// Expensive Best effort guess. Fix lack of charset for mail header fields
	// Just Check first if Content-Transfer-Encoding: 8bit to optimize/reduce cpu ?
	if ((mBody->m_ContentType.IsEmpty() == FALSE) && mBody->m_IsTextPlain && 
		mBody->m_Charset.CompareNoCase("us-ascii") && mBody->m_Charset.CompareNoCase("utf-8") &&
		!mBody->m_Charset.IsEmpty())
	{
		CStringA charset = mBody->m_Charset;
		UINT charsetId = TextUtilsEx::StrPageCodeName2PageCode(charset);

		if (m->m_from_charset.IsEmpty())
		{
			if (CheckIf8Bit((LPCSTR)m->m_from, m->m_from.GetLength()))
			{
				m->m_from_charset = charset;
				m->m_from_charsetId = charsetId;
				int deb = 1;
			}
		}
		if (m->m_to_charset.IsEmpty())
		{
			if (CheckIf8Bit((LPCSTR)m->m_to, m->m_to.GetLength()))
			{
				m->m_to_charset = charset;
				m->m_to_charsetId = charsetId;
				int deb = 1;
			}
		}
		if (m->m_subj_charset.IsEmpty())
		{
			if (CheckIf8Bit((LPCSTR)m->m_subj, m->m_subj.GetLength()))
			{
				m->m_subj_charset = charset;
				m->m_subj_charsetId = charsetId;
				int deb = 1;
			}
		}
		if (m->m_cc_charset.IsEmpty())
		{
			if (CheckIf8Bit((LPCSTR)m->m_cc, m->m_cc.GetLength()))
			{
				m->m_cc_charset = charset;
				m->m_cc_charsetId = charsetId;
				int deb = 1;
			}
		}
		if (m->m_bcc_charset.IsEmpty())
		{
			if (CheckIf8Bit((LPCSTR)m->m_bcc, m->m_bcc.GetLength()))
			{
				m->m_bcc_charset = charset;
				m->m_bcc_charsetId = charsetId;
				int deb = 1;
			}
		}
	}
#endif

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

			if (!pBP->m_AttachmentName.IsEmpty())
			{
				contentDetails->m_attachmentName = pBP->m_AttachmentName;
				contentDetails->m_attachmentNamePageCode = pBP->m_AttachmentNamePageCode;
			}
			else if (!pBP->m_Name.IsEmpty())
			{
				contentDetails->m_attachmentName = pBP->m_Name;
				contentDetails->m_attachmentNamePageCode = pBP->m_NamePageCode;
			}

			contentDetails->m_contentLocation = pBP->m_ContentLocation;
			contentDetails->m_pageCode = pBP->m_PageCode;

			if ((pBP->m_PageCode == 0) && pBP->m_IsTextHtml)
				int deb = 1;

			m->m_ContentDetailsArray.push_back(contentDetails);
		}
	}
#endif
	// now it is save to free mBody
	MailBody::FreeMailBody(mBody);
	return mailBegin;  // not used, what should we return ?
}


// Or enhance delete MboxMail
void MboxMail::DestroyMboxMail(MboxMail *m)
{
	int j;
	for (j = 0; j < m->m_ContentDetailsArray.size(); j++) {
		delete m->m_ContentDetailsArray[j];
		m->m_ContentDetailsArray[j] = 0;
	}
	delete m;
}

void MboxMail::Destroy(MailArray *array)
{
	clearMboxMailTable();
	clearMessageIdTable();

	if (array == 0)
		array = &s_mails_ref;

	MailArray &mails = *array;

	int cnt = (int)mails.GetSize();
	if (cnt < mails.MaxSize())
	{
		MboxMail **e = mails.GetData();
		if (e[cnt])
			int deb = 1;
	}
	int i;
	for (i = 0; i < cnt; i++)
	{
		DestroyMboxMail(mails[i]);
		mails[i] = 0;
	}

	s_mails.SetSizeKeepData(0);
	s_mails_ref.SetSizeKeepData(0);
	s_mails_find.SetSizeKeepData(0);
	s_mails_edit.SetSizeKeepData(0);
	s_mails_selected.SetSizeKeepData(0);
	s_mails_merged.SetSizeKeepData(0);
	s_mails_folder.SetSizeKeepData(0);

	b_mails_sorted = false;
	//MboxMail::s_path = "";
	MboxMail::SetMboxFilePath(CString(""), TRUE);

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
		pFrame->EnableAllMailLists(FALSE);

	MboxMail::nWhichMailList = -1;
};

#define HASH_ARRAY_SIZE 50013
//
int MboxMail::getMessageId(CStringA *key)
{
	MessageIdTableType::iterator it;
	if (m_pMessageIdTable == 0)
		createMessageIdTable(HASH_ARRAY_SIZE);

	it = m_pMessageIdTable->find(key);
	if (it != m_pMessageIdTable->cend()) {
		return it->second;
	}
	else
		return -1;
}

bool MboxMail::insertMessageId(CStringA *key, int val)
{
	if (m_pMessageIdTable == 0)
		createMessageIdTable(HASH_ARRAY_SIZE);

	CStringA *mapKey = new CStringA;
	*mapKey = *key;
	std::pair<MessageIdTableType::iterator, bool> result = m_pMessageIdTable->insert(MessageIdTableType::value_type(mapKey, val));
	return result.second;
}

UINT MboxMail::createMessageIdTable(UINT count)
{
	if (m_pMessageIdTable == 0)
	{
		m_pMessageIdTable = new MessageIdTableType;
		m_pMessageIdTable->reserve(count);
	}
	return (UINT)m_pMessageIdTable->size();
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

//
int MboxMail::getThreadId(CStringA *key)
{
	ThreadIdTableType::iterator it;
	if (m_pThreadIdTable == 0)
		createThreadIdTable(HASH_ARRAY_SIZE);

	it = m_pThreadIdTable->find(key);
	if (it != m_pThreadIdTable->cend()) {
		return it->second;
	}
	else
		return -1;
}

bool MboxMail::insertThreadId(CStringA *key, int val)
{
	if (m_pThreadIdTable == 0)
		createThreadIdTable(HASH_ARRAY_SIZE);

	CStringA *mapKey = new CStringA;
	*mapKey = *key;
	std::pair<ThreadIdTableType::iterator, bool> result = m_pThreadIdTable->insert(ThreadIdTableType::value_type(mapKey, val));
	return result.second;
}

UINT MboxMail::createThreadIdTable(UINT count)
{
	if (m_pThreadIdTable == 0)
	{
		m_pThreadIdTable = new ThreadIdTableType;
		m_pThreadIdTable->reserve(count);
	}
	return (UINT)m_pThreadIdTable->size();
}

void MboxMail::clearThreadIdTable()
{
	ThreadIdTableType::iterator it;

	if (m_pThreadIdTable) {
		for (it = m_pThreadIdTable->begin(); it != m_pThreadIdTable->end(); it++) {
			delete it->first;
			//it->first = 0;
		}
		m_pThreadIdTable->clear();
	}
}

int MboxMail::add2ConversationGroup(int mid, MboxMail *m)
{
	// TODO: ??
	return 1;
}

bool MboxMail::HasGMThreadId()
{
	MboxMail *m;

	int gmThreadIdCnt = 0;
	int maxMails = 1000;
	if ((int)s_mails.GetSize() < maxMails)
		maxMails = s_mails.GetSize();

	for (int i = 0; i < maxMails; i++)
	{
		m = s_mails[i];

		int id = i;
		if (m->m_threadId.GetLength())
			gmThreadIdCnt++;
	}
	if (gmThreadIdCnt > (maxMails*0.9))
		return true;
	else
		return false;
}

bool MboxMail::preprocessConversationsByThreadId()
{
	MboxMail *m;
	MboxMail *m_reply_to;

	MboxMail::SortByDate();

	// !! s_mails is already populated from mboxview file
	// Create links between all related messages based on X-GM-THRID (from mail header)

	createThreadIdTable(s_mails.GetSize());

	for (int i = 0; i < s_mails.GetSize(); i++)
	{
		m = s_mails[i];
		m->m_duplicateId = false;  // first mail message within list of conversation list

		int id = i;
		if (m->m_threadId.GetLength())
		{
			int mId = getThreadId(&m->m_threadId);
			if (mId < 0) 
				insertThreadId(&m->m_threadId, id);
			else 
				m->m_duplicateId = true;  // not the first mail message within list of conversation list
		}
		else
		{
			// What to do to maintain order of mails?
			int deb = 1;
		}
	}

	int orphanCount = 0;
	// Link all mails with related X-GM-THRID 
	for (int i = 0; i < s_mails.GetSize(); i++)
	{
		m = s_mails[i];

		if (m->m_duplicateId == false)  // first mail message within conversation list
			continue;

		if (m->m_threadId.GetLength())
		{
			int rId = getThreadId(&m->m_threadId);
			if (rId < 0) 
			{
				// should never be here
				orphanCount++;
			}
			else
			{
				m_reply_to = s_mails[rId];  // find root mail of the conversation thread
				if (m_reply_to == m)
					continue;

				// find the last mail linked to convesation mail
				while (m_reply_to->m_nextMail >= 0)
				{
					rId = m_reply_to->m_nextMail;
					m_reply_to = s_mails[rId];
				}
				//link this mail to conversation list
				// we use indexes to array to link mails, therfore order of the array elements must not change
				// Using m_nextMail and m_prevMail we can always find all related mails

				m_reply_to->m_nextMail = i;  
				m->m_prevMail = rId;

				int deb = 1;
			}
		}
		else
		{
			int deb = 1;
		}
		_ASSERTE(orphanCount == 0);
	}
	return true;
}

bool MboxMail::preprocessConversations()
{
	MboxMail *m;
	MboxMail *m_reply_to;

	MboxMail::SortByDate();

	// !! s_mails is already populated from mboxview file
	// Create links between all related messages based on messageId and replyId (from mail header)
	// Ignore duplicate messageId

	createMessageIdTable(s_mails.GetSize());

	for (int i = 0; i < s_mails.GetSize(); i++)
	{
		m = s_mails[i];

		m->m_duplicateId = false;

		int id = i;
		if (m->m_messageId.GetLength())
		{
			int mId = MboxMail::getMessageId(&m->m_messageId);
			if (mId < 0) 
				insertMessageId(&m->m_messageId, id);
			else 
				m->m_duplicateId = true;
		}
		else
			int deb = 1;
	}

	for (int i = 0; i < s_mails.GetSize(); i++)
	{
		m = s_mails[i];

		int id = i;

		if (m->m_replyId.GetLength())
		{
			int mId = MboxMail::getMessageId(&m->m_replyId);
			if ((mId < 0) && (m->m_duplicateId == false))
			{
				insertMessageId(&m->m_replyId, id);
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
			int rId = MboxMail::getMessageId(&m->m_replyId);
			if (rId < 0) 
			{
				orphanCount++;
			}
			else 
			{
				m_reply_to = s_mails[rId];
				if (m == m_reply_to)
					continue;

				// find the last mail linked to reply mail
				while (m_reply_to->m_nextMail >= 0)
				{
					rId = m_reply_to->m_nextMail;
					m_reply_to = s_mails[rId];
				}
				//link this mail to reply mail (directly or indirectly)
				m_reply_to->m_nextMail = i;
				m->m_prevMail = rId;

				int deb = 1;
			}
		}
		else
		{
			int deb = 1;
		}
	}
	return true;
}

// traverse s_mails by conversations and copy mail groups into s_mails_ref.
// Finally, copy result from s_mails_ref to s_mails. Upon return, both s_mails and s_mails_ref are equal.
		//
		// s_mails are sorted by data here; 
		// sortConversations assigns related mails to unique group ID, 
		// mails are sorted by group ID into s_mails_ref; s_mails is not touched, i.e re mails sorted by date
		// finally unique index is assigned to each mail according to its position in the master  array s_mails_ref
		// TODO: fix the function name ??

// Two versions; different sort order, from last  to first or from first to last
// Deleted from first to last

bool MboxMail::sortConversations()
{
	MboxMail *m;
	int i_ref = s_mails.GetSize() - 1;

	s_mails_ref.SetSizeKeepData(s_mails.GetSize());

	int refSize = s_mails_ref.GetSize();

#if 0
	if (!validateSortConversations()) {
		//s_mails is already left sorted by date
		return false;
	}
#endif

	int currentGroupColor = 0; // TODO; white , not used currently, remove ?
	m_nextGroupId = 0;

	int refCnt = 0;
					   // init s_mails_ref by conversation groups
	for (int i = s_mails.GetSize() - 1; i >= 0; i--)
	{
		m = s_mails[i];
		if (m->m_done)
			continue;

		if (m->m_prevMail >= 0)
			continue;  // root mail of the list

		if (m->m_nextMail < 0)
		{
			m->m_groupId = m_nextGroupId;
			m->m_groupColor = currentGroupColor;
			m->m_done = true;

			s_mails_ref[i_ref--] = m;
			refCnt++;

			m_nextGroupId++;
			currentGroupColor++;
			if (currentGroupColor > 2)
				currentGroupColor = 1;
		}
		else if (m->m_nextMail >= 0)
		{
			m->m_groupId = m_nextGroupId;
			m->m_groupColor = currentGroupColor;
			m->m_done = true;

			s_mails_ref[i_ref--] = m;
			refCnt++;

			while (m->m_nextMail >= 0)
			{
				m = s_mails[m->m_nextMail];
				if (m->m_done)
					continue; // should never be here ??

				m->m_groupId = m_nextGroupId;
				m->m_groupColor = currentGroupColor;
				m->m_done = true;

				s_mails_ref[i_ref--] = m;
				refCnt++;
			}

			m_nextGroupId++;
			currentGroupColor++;
			if (currentGroupColor > 2)
				currentGroupColor = 1;
		}
	}

	if (refCnt != refSize)
	{
		int deb = 1;
		s_mails_ref.Copy(s_mails);
		_ASSERTE(FALSE);
	}

	MboxMail::SortByFileOffset(&MboxMail::s_mails_ref);

	for (int i = 0; i < refSize; i++)
	{
		m = s_mails_ref[i];
		m->m_index = i;
	}
	s_mails.Copy(s_mails_ref);
	return true;
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

void MboxMail::SortBySubjectBasedConversasions(MailArray *s_mails, bool bDesc)
{
	MboxMail::SortBySubject(0, bDesc);

	MboxMail *m_thread_oldest;
	MboxMail *m_thread;
	MboxMail *m;
	MboxMail *m_prev = 0;

	int mailCount = MboxMail::s_mails.GetCount();
	s_mails_selected.SetSizeKeepData(mailCount);
	s_mails_selected.SetSizeKeepData(0);

	int i = 0;
	for (i = 0; i < mailCount; i++)
	{
		m = MboxMail::s_mails[i];
		if (m_prev != 0)
		{
			if (CompBySubject(m, m_prev))
			{
				m_thread_oldest->m_mail_cnt++;
			}
			else  // m != m_prev first mail of subject thread
			{
				m_thread_oldest = m;
				m->m_mail_cnt = 1;
				m->m_mail_index = i;
				MboxMail::s_mails_selected.Add(m);
			}
		}
		else  // m_prev == 0 , first mail of subject thread
		{
			m_thread_oldest = m;
			m->m_mail_cnt = 1;
			m->m_mail_index = i;
			MboxMail::s_mails_selected.Add(m);
		}
		m_prev = m;
	}

	SortByDate(&s_mails_selected, bDesc);

	s_mails_merged.SetSizeKeepData(mailCount);
	s_mails_merged.SetSizeKeepData(0);

	int j = 0;
	int groupId = 0;
	int groupColor = 0;
	int mCount = MboxMail::s_mails_selected.GetCount();
	for (i = 0; i < mCount; i++)
	{
		m = MboxMail::s_mails_selected[i];
		int k = m->m_mail_index;
		int cnt = m->m_mail_cnt;
		for (j = 0; j < cnt; j++)
		{
			m_thread = MboxMail::s_mails[k++];
			m_thread->m_mail_index = groupId;
			m_thread->m_groupColor = groupColor;
			MboxMail::s_mails_merged.Add(m_thread);
		}
		groupId++;
		groupColor++;
		if (groupColor > 1)
			groupColor = 0;
	}

	_ASSERTE(mailCount == MboxMail::s_mails_merged.GetCount());

	MboxMail::s_mails.CopyKeepData(MboxMail::s_mails_merged);

	int deb = 1;
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
	TRACE(L"(ALongRightProcessProcPrintMailArchiveToCSV) threadId=%ld threadPriority=%ld\n", myThreadId, myThreadPri);

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
	CStringA colLabels;
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
	int limitedLength = (int)(p - str->Data());
	if (limitedLength > str->Count())
		limitedLength = str->Count();

	return limitedLength;
}

int MboxMail::printMailHeaderToCSVFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, CSVFILE_CONFIG &csvConfig)
{
	DWORD error;
	MboxMail *m;
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
		CStringA format = MboxMail::GetDateFormat(csvConfig.m_dateFormat);

		datebuff[0] = 0;
		if (m->m_timeDate >= 0)
		{
			MyCTime tt(m->m_timeDate);
			if (!csvConfig.m_bGMTTime) 
			{
				CStringA lDateTime = tt.FormatLocalTmA(format);
				strcpy(datebuff, (LPCSTR)lDateTime);
			}
			else {
				CStringA lDateTime = tt.FormatGmtTmA(format);
				strcpy(datebuff, (LPCSTR)lDateTime);
			}
		}

		if (separatorNeeded)
			outbuf.Append(sepchar);

		outbuf.Append('"');

		int begCount = outbuf.Count();

		int datelen = istrlen(datebuff);
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
			BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf, error);
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
			BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf, error);
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
		atpos = strchr(token, '@');
		while (atpos != 0)
		{
			seppos = strchr(atpos, sepchar);
			if (seppos == 0)
			{
				tokenlen = IntPtr2Int(last - token);
				name.ClearAndResize(tokenlen);
				addr.ClearAndResize(tokenlen);
				splitMailAddress(token, tokenlen, &name, &addr);
				atpos = 0;
			}
			else
			{
				*seppos = 0;
				tokenlen = IntPtr2Int(seppos - token);

				name.ClearAndResize(tokenlen);
				addr.ClearAndResize(tokenlen);
				splitMailAddress(token, tokenlen, &name, &addr);

				token = seppos + 1;
				atpos = strchr(token, '@');
			}

			int addrlen = addr.Count();

			tmpbuf.ClearAndResize(2 * addrlen);
			int ret_addrlen = escapeSeparators(tmpbuf.Data(), addr.Data(), addrlen, '"');
			tmpbuf.SetCount(ret_addrlen);

			UINT pageCode = m->m_to_charsetId;
			if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
			{
				BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf, error);
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
			BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf, error);
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
		atpos = strchr(token, '@');
		while (atpos != 0)
		{
			seppos = strchr(atpos, sepchar);
			if (seppos == 0)
			{
				tokenlen = IntPtr2Int(last - token);
				name.ClearAndResize(tokenlen);
				addr.ClearAndResize(tokenlen);
				splitMailAddress(token, tokenlen, &name, &addr);
				atpos = 0;
			}
			else
			{
				*seppos = 0;
				tokenlen = IntPtr2Int(seppos - token);

				name.ClearAndResize(tokenlen);
				addr.ClearAndResize(tokenlen);
				splitMailAddress(token, tokenlen, &name, &addr);

				token = seppos + 1;
				atpos = strchr(token, '@');
			}

			int addrlen = addr.Count();

			tmpbuf.ClearAndResize(2 * addrlen);
			int ret_addrlen = escapeSeparators(tmpbuf.Data(), addr.Data(), addrlen, '"');
			tmpbuf.SetCount(ret_addrlen);

			UINT pageCode = m->m_to_charsetId;
			if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
			{
				BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf, error);
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
		atpos = strchr(token, '@');
		while (atpos != 0)
		{
			seppos = strchr(atpos, sepchar);
			if (seppos == 0)
			{
				tokenlen = IntPtr2Int(last - token);
				name.ClearAndResize(tokenlen);
				addr.ClearAndResize(tokenlen);
				splitMailAddress(token, tokenlen, &name, &addr);
				atpos = 0;
			}
			else
			{
				*seppos = 0;
				tokenlen = IntPtr2Int(seppos - token);

				name.ClearAndResize(tokenlen);
				addr.ClearAndResize(tokenlen);
				splitMailAddress(token, tokenlen, &name, &addr);

				token = seppos + 1;
				atpos = strchr(token, '@');
			}

			int addrlen = addr.Count();

			tmpbuf.ClearAndResize(2 * addrlen);
			int ret_addrlen = escapeSeparators(tmpbuf.Data(), addr.Data(), addrlen, '"');
			tmpbuf.SetCount(ret_addrlen);

			UINT pageCode = m->m_to_charsetId;
			if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
			{
				BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, csvConfig.m_nCodePageId, inbuf, workbuf, error);
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

		if (!m->m_DetermineEmbeddedImagesDone)
		{
			int ret = MboxMail::DetermineEmbeddedImages(mailPosition, fpm);
		}

		BOOL targetCSV = TRUE;
		NListView::PrintAttachmentNamesAsText2CSV(mailPosition, &outbuf, csvConfig.m_MessageLimitCharsString, csvConfig.m_AttachmentNamesSeparatorString);

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
		CString txt = L"Please open mail file first.";
		ResHelper::TranslateString(txt);

		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	CString datapath = MboxMail::GetLastDataPath();
	CString path = MboxMail::GetLastPath();
	if (path.IsEmpty()) {
		CString txt = L"No path to archive file folder.";
		ResHelper::TranslateString(txt);

		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}
	csvFileName = csvFile;

	if (!progressBar)
	{
		CFileException ExError;
		if (!fp.Open(csvFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &ExError))
		{
			DWORD lastErr = ::GetLastError();
#if 1
			//HWND h = GetSafeHwnd();
			HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
			CString fmt = L"Could not create file:\n\n\"%s\"\n\n%s";  // new format
			CString errorText = FileUtils::ProcessCFileFailure(fmt, csvFile, ExError, lastErr, h);
#else

			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt;
			CString fmt = L"Could not create \"%s\" file.\n%s";
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, csvFile, exErrorStr);

			//TRACE(L"%s\n", txt);

			CFileStatus rStatus;
			BOOL ret = fp.GetStatus(rStatus);

			//errorText = txt;

			HWND h = NULL; // we don't have any window yet
			int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
#endif
			return -1;
		}

		CFile fpm;
		if (csvConfig.m_bContent)
		{
			CFileException ExError;
			if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
			{
				DWORD lastErr = ::GetLastError();
#if 1
				//HWND h = GetSafeHwnd();
				HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
				CString fmt = L"Could not open mail file:\n\n\"%s\"\n\n%s";  // new format
				CString errorText = FileUtils::ProcessCFileFailure(fmt, MboxMail::s_path, ExError, lastErr, h);
#else
				CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

				CString txt;
				CString fmt = L"Could not open \"%s\" mail file.\n%s";
				ResHelper::TranslateString(fmt);
				txt.Format(fmt, MboxMail::s_path, exErrorStr);

				//TRACE(L"%s\n", txt);
				//errorText = txt;

				HWND h = NULL; // we don't have any window yet
				int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
#endif

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
			int cnt = (int)selectedMailsIndexList->GetCount();
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
		Dlg.SetDialogTemplate(AfxGetApp()->m_hInstance, MAKEINTRESOURCE(IDD_PROGRESS_DLG), IDC_STATIC, IDC_PROGRESS_BAR, IDCANCEL);

		INT_PTR nResult = Dlg.DoModal();

		if (!nResult) // should never be true ?
		{
			MboxMail::assert_unexpected();
			return -1;
		}

		TRACE(L"m_Html2TextCount=%d MailCount=%d\n", m_Html2TextCount, MboxMail::s_mails.GetCount());

		int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
		int retResult = LOWORD(nResult);

		if (retResult != IDOK)
		{  // IDOK==1, IDCANCEL==2
			// We should be here when user selects Cancel button
			_ASSERTE(cancelledbyUser == TRUE);

			DWORD terminationDelay = Dlg.GetTerminationDelay();
			int loopCnt = (terminationDelay+100)/25;

			ULONGLONG tc_start = GetTickCount64();
			while ((loopCnt-- > 0) && (args.exitted == FALSE))
			{
				Sleep(25);
			}
			ULONGLONG tc_end = GetTickCount64();
			DWORD delta = (DWORD)(tc_end - tc_start);
			TRACE(L"(exportToCSVFile)Waited %ld milliseconds for thread to exist.\n", delta);
		}

		MboxMail::pCUPDUPData = NULL;

		if (!args.errorText.IsEmpty()) 
		{
			HWND h = NULL; 
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return - 1;
		}

	}
	return 1;
}

int MboxMail::printSingleMailToCSVFile(/*out*/ CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, CSVFILE_CONFIG &csvConfig, bool firstMail)
{
	DWORD error;

	if (firstMail)
		exportHeaderFieldLabelsToCSVFile(csvConfig, fp);

	printMailHeaderToCSVFile(fp, mailPosition, fpm, csvConfig);

	char sepchar = csvConfig.m_separator.GetAt(0);
	char *sepstr = (char*)(LPCSTR)csvConfig.m_separator;

	SimpleString outb(1024, 256);
	SimpleString tmpbuf(1024, 256);
	SimpleString name(256);
	SimpleString addr(256);

	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;
	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);

	outb.Clear();

	if (csvConfig.m_bContent)
	{
		outb.Append('"');

		fp.Write(outb.Data(), outb.Count());

		SimpleString *outbuf = MboxMail::m_outbuf;

		outbuf->Clear();
		UINT pageCode = 0;
		int textType = 0; // plain text
		int plainTextMode = 0;
		int textlen = GetMailBody_mboxview(fpm, mailPosition, outbuf, pageCode, textType, plainTextMode);  // fast  FIXME name of plainTextMode ??
		if (textlen != outbuf->Count())
			int deb = 1;

		if (outbuf->Count())
		{
			int lineLimit = -1;
			if (TextUtilsEx::isNumeric(csvConfig.m_MessageLimitString))
				lineLimit = _ttoi(csvConfig.m_MessageLimitString);

			if (lineLimit >= 0) {
				HtmlUtils::MergeWhiteLines(outbuf, lineLimit);
			}

			if (outbuf->Count())
			{
				int needLength = outbuf->Count() * 2 + 10; // worst case scenario or get '"' count first
				inbuf->ClearAndResize(needLength);  // escapeSeparators terminates result with null

				BOOL insertEmptyLine = FALSE;   // make this configurable one day -:)
				if ((outbuf->Data()[0] == '\n') || ((outbuf->Data()[0] == '\r') && (outbuf->Data()[1] == '\n')))
					; // do noting
				else if (insertEmptyLine)
					inbuf->Append("\r\n");  // insert empty line at the beginning

				int cnt_sv = inbuf->Count();
				int retcnt = escapeSeparators(inbuf->Data(cnt_sv), outbuf->Data(), outbuf->Count(), '"');
				inbuf->SetCount(retcnt + cnt_sv);

				if (csvConfig.m_nCodePageId && (pageCode != 0) && (pageCode != csvConfig.m_nCodePageId))
				{
					BOOL ret = TextUtilsEx::Str2CodePage(inbuf, pageCode, csvConfig.m_nCodePageId, outbuf, workbuf, error);
					if (ret) {
						EnforceCharacterLimit(outbuf, csvConfig.m_MessageLimitCharsString);
						fp.Write(outbuf->Data(), outbuf->Count());
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

			int plainTextMode = 0;  // no extra img tags; html text has img tags already
			int textlen = GetMailBody_mboxview(fpm, mailPosition, /*out*/outbuf, pageCode, textType, plainTextMode);  // returns pageCode
			if (textlen != outbuf->Count())
				int deb = 1;

			if (outbuf->Count())
			{
				MboxMail::m_Html2TextCount++;

				CStringA bdycharset;
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

				CStringA bdy;
				bool extraHtmlHdr = false;
				if (outbuf->FindNoCase(0, "<body", 5) < 0) // didn't find if true
				{
					extraHtmlHdr = true;
					bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body>";
					inbuf->Append((LPCSTR)bdy, bdy.GetLength());
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

		outb.Clear();
		outb.Append('"');
	}

	outb.Append("\r\n");

	fp.Write(outb.Data(), outb.Count());

	return 1;
}

int MboxMail::printMailHeaderToTextFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig, HdrFldConfig &hdrFieldConfig)
{
	static char *border = "--------------------------------------------------------------------------------\r\n";

	int fldNumb;
	int fldNameLength;
	int fldTxtLength;
	char *fldText;
	char *fldName = "";
	int ii;
	DWORD error;

	SimpleString outbuf(1024, 256);
	SimpleString tmpbuf(1024, 256);

	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;
	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);

	MboxMail *m = s_mails[mailPosition];

	outbuf.Append(border);

	UINT pageCode = 0;
	int fldCnt = 0;
	for (ii = 0; ii < HdrFldList::HDR_FLD_MAX; ii++)
	{
		if (hdrFieldConfig.m_HdrFldList.IsFldSet(ii))
		{
			fldNumb = ii + 1;
			if (fldNumb == HdrFldList::HDR_FLD_SUBJECT)
			{
				fldName = "SUBJECT: ";
				fldNameLength = istrlen(fldName);
				fldTxtLength = m->m_subj.GetLength();
				fldText = (char *)(LPCSTR)m->m_subj;
				pageCode = m->m_subj_charsetId;
			}
			else if (fldNumb == HdrFldList::HDR_FLD_FROM)
			{
				fldName = "FROM: ";
				fldNameLength = istrlen(fldName);
				fldTxtLength = m->m_from.GetLength();
				fldText = (char *)(LPCSTR)m->m_from;
				pageCode = m->m_from_charsetId;
			}
			else if (fldNumb == HdrFldList::HDR_FLD_TO)
			{
				fldName = "TO: ";
				fldNameLength = istrlen(fldName);
				fldTxtLength = m->m_to.GetLength();
				fldText = (char *)(LPCSTR)m->m_to;
				pageCode = m->m_to_charsetId;
			}
			else if (fldNumb == HdrFldList::HDR_FLD_CC)
			{
				fldName = "CC: ";
				fldNameLength = istrlen(fldName);
				fldTxtLength = m->m_cc.GetLength();
				fldText = (char *)(LPCSTR)m->m_cc;
				pageCode = m->m_cc_charsetId;
				if (fldTxtLength <= 0)
					continue;
			}
			else if (fldNumb == HdrFldList::HDR_FLD_BCC)
			{
				fldName = "BCC: ";
				fldNameLength = istrlen(fldName);
				fldTxtLength = m->m_bcc.GetLength();
				fldText = (char *)(LPCSTR)m->m_bcc;
				pageCode = m->m_bcc_charsetId;
				if (fldTxtLength <= 0)
					continue;
			}
			else if (fldNumb == HdrFldList::HDR_FLD_DATE)
			{
				fldName = "DATE: ";
				fldNameLength = istrlen(fldName);
				fldTxtLength = 0;
				fldText = "";
				continue;
			}
			else if (fldNumb == HdrFldList::HDR_FLD_ATTACHMENTS)
			{
				fldName = "DATE: ";
				fldNameLength = istrlen(fldName);
				fldTxtLength = 0;
				fldText = "";
				continue;
			}
			else
			{
				continue;
			}

			if (fldCnt)
				outbuf.Append("\r\n");

			outbuf.Append(fldName, fldNameLength);
			tmpbuf.Copy(fldText, fldTxtLength);

			if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
			{
				BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf, error);
				if (ret)
				{
					outbuf.Append(*inbuf);
				}
				else
				{
					outbuf.Append(tmpbuf);
				}
			}
			else
			{
				outbuf.Append(tmpbuf);
			}
			fldCnt++;
		}
	}
	//
	fldNumb = HdrFldList::HDR_FLD_DATE - 1;
	if (hdrFieldConfig.m_HdrFldList.IsFldSet(fldNumb))
	{
		char datebuff[32];

		CStringA format = textConfig.m_dateFormat;

		datebuff[0] = 0;
		if (m->m_timeDate >= 0)
		{
			MyCTime tt(m->m_timeDate);
			if (!textConfig.m_bGMTTime)
			{
				CStringA lDateTime = tt.FormatLocalTmA(format);
				strcpy(datebuff, (LPCSTR)lDateTime);
			}
			else
			{
				CStringA lDateTime = tt.FormatGmtTmA(format);
				strcpy(datebuff, (LPCSTR)lDateTime);
			}
		}

		if (fldCnt)
			outbuf.Append("\r\n");

		fldName = "DATE: ";
		fldNameLength = istrlen(fldName);
		fldTxtLength = m->m_subj.GetLength();
		fldText = (char *)datebuff;

		outbuf.Append(fldName, fldNameLength);

		int datelen = istrlen(datebuff);
		outbuf.Append(datebuff, datelen);

		fldCnt++;
	}

	fldNumb = HdrFldList::HDR_FLD_ATTACHMENTS - 1;
	if (hdrFieldConfig.m_HdrFldList.IsFldSet(fldNumb))
	{
		tmpbuf.Clear();

		CString attachmentFileNamePrefix;
		int ret = MboxMail::printAttachmentNamesAsText(&fpm, mailPosition, &tmpbuf, attachmentFileNamePrefix);

		if (tmpbuf.Count() > 0)
		{
			char datebuff[32];

			CStringA format = textConfig.m_dateFormat;

			datebuff[0] = 0;
			if (m->m_timeDate >= 0)
			{
				MyCTime tt(m->m_timeDate);
				if (!textConfig.m_bGMTTime)
				{
					CStringA lDateTime = tt.FormatLocalTmA(format);
					strcpy(datebuff, (LPCSTR)lDateTime);
				}
				else
				{
					CStringA lDateTime = tt.FormatGmtTmA(format);
					strcpy(datebuff, (LPCSTR)lDateTime);
				}
			}

			if (fldCnt)
				outbuf.Append("\r\n");


			CStringA attachmentFileNamePrefixA = attachmentFileNamePrefix;
			CStringA label;
			label.Format("ATTACHMENTS (%s):", (LPCSTR)attachmentFileNamePrefixA);
			fldName = (char*)(LPCSTR)label;
			fldNameLength = label.GetLength();

			outbuf.Append(fldName, fldNameLength);
			outbuf.Append(tmpbuf.Data(), tmpbuf.Count());

			fldCnt++;
		}
	}

	outbuf.Append("\r\n");

	outbuf.Append(border);

	outbuf.Append("\r\n");

	fp.Write(outbuf.Data(), outbuf.Count());

	return 1;
}

int MboxMail::printSingleMailToTextFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig, bool singleMail, BOOL pageBreakNeeded)
{
	UINT pageCode;
	DWORD error;

	MboxMail *m = MboxMail::s_mails[mailPosition];

	SimpleString *outbuf = MboxMail::m_outbuf;
	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;

	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);
	outbuf->ClearAndResize(10000);

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	int ret;
	if (pFrame)
	{
		ret = printMailHeaderToTextFile(fp, mailPosition, fpm, textConfig, pFrame->m_HdrFldConfig);
	}

	//
	int textType = 0; // Text
	pageCode = 0;
	int plainTextMode = 1;  // [] img tag
	int textlen = GetMailBody_mboxview(fpm, mailPosition, outbuf, pageCode, textType, plainTextMode);  // returns pageCode
	if (textlen != outbuf->Count())
		int deb = 1;

	if (outbuf->Count())
	{
		if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
		{
			int needLength = outbuf->Count() * 2 + 1;
			inbuf->ClearAndResize(needLength);

			BOOL ret = TextUtilsEx::Str2CodePage(outbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf, error);
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
		// Extract text from Html text body
		outbuf->Clear();
		pageCode = 0;
		int textType = 1; // HTML

		int plainTextMode = 0;  // no extra img tags; html text has img tags already
		int textlen = GetMailBody_mboxview(fpm, mailPosition, outbuf, pageCode, textType, plainTextMode);  // returns pageCode
		if (textlen != outbuf->Count())
			int deb = 1;

		if (outbuf->Count())
		{
			MboxMail::m_Html2TextCount++;

			CStringA bdycharset;
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

			CStringA bdy;
			bool extraHtmlHdr = false;
			if (outbuf->FindNoCase(0, "<body", 5) < 0) // didn't find if true
			{
				extraHtmlHdr = true;
				bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body>";
				inbuf->Append((LPCSTR)bdy, bdy.GetLength());
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

	if (pageBreakNeeded)
	{
		char formFeed = 12;
		outbuf->Clear();
		outbuf->Append(formFeed);
		fp.Write(outbuf->Data(), outbuf->Count());
	}

	return 1;
}

UINT getCodePageFromHtmlBody(SimpleString *buffer, std::string &charset)
{
	char *pat = "charset=";
	charset.assign("");
	UINT codePage = 0;
	int patlen = istrlen(pat);
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


// Determines which image attachment will be embeded into message and 
// sets m_isEmbeddedImage to TRUE in MailBodyContent of MboxMail
int MboxMail::DetermineEmbeddedImages(int mailPosition, /*in - mail body*/ CFile &fpm)
{
	if ((mailPosition < 0) || (mailPosition >= MboxMail::s_mails.GetCount()))
		return -1;

	MboxMail *m = s_mails[mailPosition];

	if (m->m_DetermineEmbeddedImagesDone)
		return 1;

	BOOL createEmbeddedImageFiles = FALSE;
	CString imageCachePath = L".\\";
	int ret = NListView::CreateEmbeddedImageFilesEx(fpm, mailPosition, imageCachePath, createEmbeddedImageFiles);

	m->m_DetermineEmbeddedImagesDone = 1;

	return ret;
}

// FIXME FIXME - must review 
int MboxMail::printAttachmentNamesAsHtml(CFile *fpm, int mailPosition, SimpleString *outbuf, CString &attachmentFileNamePrefix, CString& attachmentFilesFolderPath)
{
	if ((mailPosition < 0) || (mailPosition >= MboxMail::s_mails.GetCount()))
		return -1;

	MboxMail *m = s_mails[mailPosition];

	if (m == 0)
	{
		// TODO: Assert ??
		return 1;
	}

	CFile mboxFp;
	CFile *fpm_save = fpm;
	if (fpm == 0)
	{
		CFileException ExError;
		if (!mboxFp.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
		{
			DWORD lastErr = ::GetLastError();
#if 1
			//HWND h = GetSafeHwnd();
			HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
			CString fmt = L"Could not open mail file:\n\n\"%s\"\n\n%s";  // new format
			CString errorText = FileUtils::ProcessCFileFailure(fmt, MboxMail::s_path, ExError, lastErr, h);
#else
			// TODO: critical failure
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt = L"Could not open \"" + MboxMail::s_path;
			txt += L"\" mail file.\n";
			txt += exErrorStr;

			TRACE(L"%s\n", txt);
			//errorText = txt;
#endif

			return FALSE;
		}
		fpm = &mboxFp;
	}

	CString errorText;
	CString attachmentCachePath;
	CStringW printCachePathW;
	CString rootPrintSubFolder = "AttachmentCache";
	CString targetPrintSubFolder;
	if (NListView::m_exportMailsMode)
	{
		rootPrintSubFolder = L"ExportCache";
		targetPrintSubFolder = L"Attachments";
	}

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, attachmentCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		// TODO: what to do ?
		//int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		//return -1;
	}

	CString mboxFileNamePath = MboxMail::s_path;
	CString driveName;
	CString mboxFileDirectory;
	CString mboxFileNameBase;
	CString mboxFileNameExtention;

	FileUtils::SplitFilePath(mboxFileNamePath, driveName, mboxFileDirectory, mboxFileNameBase, mboxFileNameExtention);

	CString mboxFileName = mboxFileNameBase + mboxFileNameExtention;

	if (!m->m_DetermineEmbeddedImagesDone)
		retval = MboxMail::DetermineEmbeddedImages(mailPosition, *fpm);

	SimpleString tmpbuf(1024, 256);
	SimpleString validNameUTF8(1024, 256);
	CStringW nameW;
	CStringW validNameW;
	BOOL ret;
	BOOL bReplaceWhiteWithUnderscore = FALSE;

	MailBodyContent *body;

	if (!m->m_DetermineEmbeddedImagesDone)
	{
		_ASSERTE(FALSE);
		ret = MboxMail::DetermineEmbeddedImages(mailPosition, *fpm);
	}

	NListView::GetExtendedMailId(m, attachmentFileNamePrefix);

	DWORD error;
	CStringA attachmentCachePathA;
	UINT CP_US_ASCII = 20127;
	UINT outCodePage = CP_UTF8;
	BOOL retW2A = TextUtilsEx::WStr2CodePage((LPCWSTR)attachmentCachePath, attachmentCachePath.GetLength(), outCodePage, &attachmentCachePathA, error);

	bool showAllAttachments = false;
	AttachmentConfigParams *attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
	if (attachmentConfigParams)
	{
		showAllAttachments = attachmentConfigParams->m_bShowAllAttachments_Window;
	}

	CString attachementFilePath;
	CStringA attachmentFilePathA;
	SimpleString attachmentFilePathA_SS_Root;

	if (NListView::m_fullImgFilePath)
	{
		attachementFilePath = attachmentCachePath + L"\\";
		attachementFilePath.Append(attachmentFileNamePrefix, attachmentFileNamePrefix.GetLength());

		BOOL retW2A2 = TextUtilsEx::WStr2CodePage((LPCWSTR)attachementFilePath, attachementFilePath.GetLength(), outCodePage, &attachmentFilePathA, error);

		attachmentFilePathA_SS_Root.Append((LPCSTR)attachmentFilePathA, attachmentFilePathA.GetLength());
		encodeTextAsHtmlLink(attachmentFilePathA_SS_Root);
	}
	else
	{
		attachementFilePath = attachmentFilesFolderPath + L"\\";
		attachementFilePath.Append(attachmentFileNamePrefix, attachmentFileNamePrefix.GetLength());

		BOOL retW2A2 = TextUtilsEx::WStr2CodePage((LPCWSTR)attachementFilePath, attachementFilePath.GetLength(), outCodePage, &attachmentFilePathA, error);

		attachmentFilePathA_SS_Root.Append((LPCSTR)attachmentFilePathA, attachmentFilePathA.GetLength());
		encodeTextAsHtmlLink(attachmentFilePathA_SS_Root);
	}

	CMainFrame* pFrame = 0;
	pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	//if (!pFrame)
		//return -1;

	AttachmentMgr attachmentDB;
	int attachmentCnt = 0;
	SimpleString fileName(1024);
	SimpleString href(1024);
	TRACE("printAttachmentNamesAsHtml List BEGIN:\n");
	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];

		TRACE("printAttachmentNamesAsHtml: RAW block name=%s\n", body->m_attachmentName);

		BOOL showAttachment = FALSE;
		if (body->IsAttachment() || body->IsInlineAttachment())
		{
			if (!body->m_isEmbeddedImage)
				showAttachment = TRUE;
		}

		if (showAttachment)
		{
			UINT inCodePage = body->m_attachmentNamePageCode;

			SimpleString*buf = 0;
			// FIXME remapDuplicateNames = FALSE implies that duplicate names will not be added into attachmentDB; why ??
			// why not remapp duplicate names and add to attachmentDB always ??
			// I guess attachmentDB is not used later ??
			BOOL remapDuplicateNames = FALSE; 
			NListView::DetermineAttachmentName(fpm, mailPosition, body, buf, nameW, attachmentDB, remapDuplicateNames);

			TRACE(L"printAttachmentNamesAsHtml: DETERMINED block name=%s\n", nameW);

			validNameW.Empty();
			validNameW.Append(nameW);  // FIXME

			//FileUtils::MakeValidFileName(nameW, validNameW, bReplaceWhiteWithUnderscore);  // FIXMEFIXME

			UINT outCodePage = CP_UTF8;
			BOOL retW2A = TextUtilsEx::WStr2CodePage((LPCWSTR)validNameW, validNameW.GetLength(), outCodePage, &validNameUTF8, error);

			TRACE("printAttachmentNamesAsHtml: UTF8 block name=%s\n", validNameUTF8.Data());

			// Create hyperlink
			if (attachmentCnt)
				outbuf->Append(" , ");

			SimpleString validNameUTF8Link;
			
			validNameUTF8Link.Append(validNameUTF8);

			SimpleString attachmentFilePathA_SS;
			encodeTextAsHtmlLink(validNameUTF8Link);
			attachmentFilePathA_SS.Append(attachmentFilePathA_SS_Root);
			attachmentFilePathA_SS.Append(validNameUTF8Link);

			//validNameUTF8.Append("<>;@&");  // for testing
			encodeTextAsHtmlLinkLabel(validNameUTF8);


			if (NListView::m_fullImgFilePath)
				outbuf->Append("\r\n<a href=\"file:\\\\\\");
			else
				outbuf->Append("\r\n<a href=\"");

			href.Clear();
			
			href.Append(attachmentFilePathA_SS);
			//href.Append('\\');
			//href.Append(fileName.Data(), fileName.Count());

			if (pFrame && (pFrame->m_HdrFldConfig.m_bHdrAttachmentLinkOpenMode == 1))
				href.Append("\"target=\"_blank\">");
			else
				href.Append("\">");
			href.Append(validNameUTF8.Data(), validNameUTF8.Count());
			href.Append("</a>");

			outbuf->Append(href);
			attachmentCnt++;
		}
	}
	TRACE("printAttachmentNamesAsHtml List END:\n");

	if (fpm_save == 0)
		mboxFp.Close();

	return 1;
}

int MboxMail::printAttachmentNamesAsText(CFile *fpm, int mailPosition, SimpleString *outbuf, CString &attachmentFileNamePrefix)
{
	if ((mailPosition < 0) || (mailPosition >= MboxMail::s_mails.GetCount()))
		return -1;

	MboxMail *m = s_mails[mailPosition];

	if (m == 0)
	{
		// TODO: Assert ??
		return 1;
	}

	CFile mboxFp;
	CFile *fpm_save = fpm;
	if (fpm == 0)
	{
		CFileException ExError;
		if (!mboxFp.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
		{
			DWORD lastErr = ::GetLastError();

			// TODO: critical failure
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt = L"Could not open \"" + MboxMail::s_path;
			txt += L"\" mail file.\n";
			txt += exErrorStr;

			TRACE(L"%s\n", txt);
			//errorText = txt;

			return FALSE;
		}
		fpm = &mboxFp;
	}

	SimpleString tmpbuf(1024, 256);
	SimpleString resultUTF8(1024, 256);
	CStringW nameW;
	CStringW validNameW;
	DWORD error;
	BOOL ret;
	BOOL bReplaceWhiteWithUnderscore = FALSE;

	MailBodyContent *body;

	CString errorText;
	if (!m->m_DetermineEmbeddedImagesDone)
		ret = MboxMail::DetermineEmbeddedImages(mailPosition, *fpm);

	NListView::GetExtendedMailId(m, attachmentFileNamePrefix);

	bool showAllAttachments = false;
	AttachmentConfigParams *attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
	if (attachmentConfigParams)
	{
		showAllAttachments = attachmentConfigParams->m_bShowAllAttachments_Window;
	}

	AttachmentMgr attachmentDB;
	int attachmentCnt = 0;
	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];

		BOOL showAttachment = FALSE;
		if (body->IsAttachment() || body->IsInlineAttachment())
		{
			if (showAllAttachments || !body->m_isEmbeddedImage)
				showAttachment = TRUE;
		}

		if (showAttachment)
		{
			UINT inCodePage = body->m_attachmentNamePageCode;

			SimpleString*buf = 0;
			// FIXME remapDuplicateNames = FALSE implies that duplicate names wil not be added into attachmentDB; why ??
			// why not remapp duplicate names and add to attachmentDB always ??
			// I guess attachmentDB is not used later ??
			BOOL remapDuplicateNames = FALSE;
			BOOL extraValidation = TRUE;  // FIXMME
			NListView::DetermineAttachmentName(fpm, mailPosition, body, buf, nameW, attachmentDB, remapDuplicateNames);
			// NListView::DetermineAttachmentName(fpm, mailPosition, body, buf, nameW, attachmentDB, remapDuplicateNames, extraValidation);

			validNameW.Empty();
			validNameW.Append(nameW);  // FIXMEFIXME
			//FileUtils::MakeValidFileName(nameW, validNameW, bReplaceWhiteWithUnderscore);

			UINT outCodePage = CP_UTF8;
			BOOL ret2 = TextUtilsEx::WStr2CodePage((wchar_t*)(LPCWSTR)validNameW, validNameW.GetLength(), outCodePage, &resultUTF8, error);

			if (attachmentCnt)
				outbuf->Append(" , ");
			else
				outbuf->Append(" ");


			outbuf->Append('"');
			outbuf->Append(resultUTF8.Data(), resultUTF8.Count());
			outbuf->Append('"');

			attachmentCnt++;
		}
	}

	if (fpm_save == 0)
		mboxFp.Close();

	return 1;
}

BOOL MboxMail::PageBreakNeeded(MailIndexList *selectedMailIndexList, int index, bool singleMail)
{
	int cnt = (int)selectedMailIndexList->GetCount();
	if ((index + 1) >= cnt)
		return FALSE;

	int mailPosition = (*selectedMailIndexList)[index];
	int nextMailPosition = (*selectedMailIndexList)[index+1];

	BOOL breakNeeded = PageBreakNeeded(mailPosition, nextMailPosition, singleMail);
	return breakNeeded;
}

BOOL MboxMail::PageBreakNeeded(int mailPosition, int nextMailPosition, bool singleMail)
{
	if (singleMail)
		return FALSE;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame == 0)
		return FALSE;

	NListView *pListView = pFrame->GetListView();
	if (pListView == 0)
		return FALSE;

	if (pFrame->m_NamePatternParams.m_bPrintToSeparatePDFFiles == TRUE)
		return FALSE;
	else
	{
		if ((pFrame->m_NamePatternParams.m_bAddPageBreakAfterEachMailInPDF == FALSE) && 
			(pFrame->m_NamePatternParams.m_bAddPageBreakAfterEachMailConversationThreadInPDF == FALSE))
			return FALSE;

		if (pFrame->m_NamePatternParams.m_bAddPageBreakAfterEachMailInPDF)
			return TRUE;
	}

	if (nextMailPosition >= s_mails.GetCount())
		return FALSE;


	BOOL addPageBreak = FALSE;
	if ((abs(MboxMail::b_mails_which_sorted) == 99) || ((abs(MboxMail::b_mails_which_sorted) == 4) && (pListView->m_subjectSortType == 1)))
	{
		if (nextMailPosition < s_mails.GetCount())
		{
			MboxMail *m = s_mails[mailPosition];
			MboxMail *m_next = s_mails[nextMailPosition];

			if (abs(MboxMail::b_mails_which_sorted) == 99)
			{
				if (m_next->m_groupId != m->m_groupId)
				{
					addPageBreak = TRUE;
				}
			}
			else if (abs(MboxMail::b_mails_which_sorted) == 4)
			{
				if (m_next->m_mail_index != m->m_mail_index)
				{
					addPageBreak = TRUE;
				}
			}
		}
	}
	return addPageBreak;
}

int RemoveString(char *data, int datalen, CStringA &str, SimpleString *outbuf)
{
	BOOL m_bCaseSens = FALSE;
	char *p = data;
	int len = datalen;
	UINT inCodePage = CP_UTF8;

	int pos = 0;
	while (pos >= 0)
	{
		int pos = g_tu.StrSearch((unsigned char *)p, len, inCodePage, (unsigned char *)(LPCSTR)str, str.GetLength(), m_bCaseSens, 0);
		if (pos >= 0)
		{
			outbuf->Append(p, pos);
			p += pos + str.GetLength();
			len -= pos + str.GetLength();
		}
		else
			break;
	}
	return  outbuf->Count();
}

int MboxMail::printDummyMailToHtmlFile(/*out*/CFile &fp)
{
#if 0
	// Commented out. Didn't fix small font issue
	CString bdyy;

	bdyy.Append("\r\n\r\n<html><body style=\"background-color:#FFFFFF;\"><div></div></body></html>");
	bdyy.Append("<article style = \"float:left;position:left;width:100%;background-color:#FFFFFF;margin: 0mm 0mm 0mm 0mm;\">");
	bdyy.Append("<style>@page {size: auto; margin: 12mm 4mm 12mm 12mm;}; * {font-size:100% !important;};</style>");

	fp.Write(bdyy, bdyy.GetLength());
#endif

	return 1;
}

int MboxMail::printSingleMailToHtmlFile(/*out*/CFile &fp, int mailPosition, /*in - mail body*/ CFile &fpm,
	TEXTFILE_CONFIG &textConfig, bool singleMail, BOOL addPageBreak, BOOL fullImgFilePath)
{
	DWORD error;
	CStringA scale = "98%;";
	char *token = 0;
	int tokenlen = 0;

	SimpleString outbuf(1024, 256);
	SimpleString tmpbuf(1024, 256);

	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;
	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);

	MboxMail *m = s_mails[mailPosition];

	CString attachmentFilesFolderPath = L"..\\AttachmentCache";
	if (NListView::m_exportMailsMode)
		attachmentFilesFolderPath = L"..\\ExportCache";



	CString errorText;
	CString aboluteImageCachePath;
	CString rootSubFolder = L"ImageCache";
	CString targetSubFolder;
	if (NListView::m_exportMailsMode)
	{
		rootSubFolder = L"ExportCache";
		targetSubFolder = L"Attachments";
	}

	BOOL retval = MboxMail::CreateCachePath(rootSubFolder, targetSubFolder, aboluteImageCachePath, errorText);
	aboluteImageCachePath.Append(L"\\");

	//////////////////////

	CString absoluteAttachmentCachePath;
	rootSubFolder = L"AttachmentCache";

	CString attachmentRootSubFolder = L"AttachmentCache";
	if (NListView::m_exportMailsMode)
	{
		rootSubFolder = L"ExportCache";
		attachmentRootSubFolder = L"Attachments";
	}

	retval = MboxMail::CreateCachePath(rootSubFolder, targetSubFolder, absoluteAttachmentCachePath, errorText);
	absoluteAttachmentCachePath.Append(L"\\");

	//////////////////////

	CString fpath = fp.GetFilePath();
	fpath.TrimRight(L"\\");
#if 0
	CStringArray a;
	FileUtils::SplitFileFolder(fpath, a);
	TextUtilsEx::TraceStringArrayW(a);
#endif
	// Sort of hack to avoid changes to interfaces // FIXME
	// Determine relative folder path for links to inline pictures

	CString cacheName = L"PrintCache";
	if (NListView::m_exportMailsMode)
		cacheName = L"ExportCache";
	CString pref;
	int pos = fpath.ReverseFind('\\');
	int itercationCnt = 0;
	while (pos >= 0)
	{
		CString tok = fpath.Mid(pos);
		tok.TrimLeft(L"\\");
		if (tok.Compare(cacheName) == 0)
			break;

		if (NListView::m_exportMailsMode)
		{
			if (itercationCnt++ > 0)
			pref += L"..\\";
		}
		else
			pref += L"..\\";

		if (pos > 0)
		{
			fpath = fpath.Mid(0, pos);
			fpath.TrimRight(L"\\");
		}

		pos = fpath.ReverseFind('\\');
	}

	if (pos >= 0)
	{
		attachmentFilesFolderPath = pref + attachmentRootSubFolder;
	}
	else
	{
		pref = L"..\\";
		_ASSERTE(FALSE);
	}

	outbuf.Clear();

	SimpleString *outbuflarge = MboxMail::m_outbuf;

	outbuflarge->Clear();
	UINT pageCode = 0;
	int textType = 1; // try first Html
	int plainTextMode = 0;  // no extra img tags; html text has img tags already
	int textlen = GetMailBody_mboxview(fpm, mailPosition, outbuflarge, pageCode, textType, plainTextMode);  // returns pageCode
	if (textlen != outbuflarge->Count())
		int deb = 1;

	workbuf->ClearAndResize(outbuflarge->Count() * 2);

#if 0
	// Doesn't seem to help
	CString searchString = "!important";
	int retcnt = RemoveString(outbuflarge->Data(), outbuflarge->Count(), searchString, workbuf);
	if (retcnt > 0)
	{
		outbuflarge->Copy(*workbuf);
		//return -1;
	}
#endif

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	int fontSizePDF = 0;
	NListView* pListView = 0;
	if (pFrame)
	{
		pListView = pFrame->GetListView();
		if (pListView)
			fontSizePDF = pListView->m_fontSizePDF;
	}


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

	CStringA htmlHdrFldNameStyle;
	CStringA fldNameFontStyle;
	CStringA fldTextFontStyle;
	if (pFrame)
	{
		CreateFldFontStyle(pFrame->m_HdrFldConfig, fldNameFontStyle, fldTextFontStyle);
		htmlHdrFldNameStyle.Format("<style>\r\n.hdrfldname{%s}\r\n.hdrfldtext{%s}\r\n</style>", fldNameFontStyle, fldTextFontStyle);
	}

	//CString preStyle = "pre { overflow-x:auto;  white-space:-moz-pre-wrap; white-space:-o-pre-wrap;  white-space:-pre-wrap; white-space:pre-wrap; word-wrap:break-word;}";
	CStringA divAvoidPageBreak = "divAvoidBreak {break-inside:avoid;}";
	CStringA preStyle = "pre { overflow-x:break-word; white-space:pre; white-space:hp-pre-wrap; white-space:-moz-pre-wrap; white-space:-o-pre-wrap;  white-space:-pre-wrap; white-space:pre-wrap; word-wrap:break-word;}";

	CStringA bdyy;
	bdyy.Append("\r\n\r\n<!DOCTYPE html>\r\n");
	bdyy.Append("<html><body style=\"background-color:#FFFFFF;\"><div></div></body></html>");
	bdyy.Append("<article style=\"width:");
	bdyy.Append("100%;");
	bdyy.Append("float:left; position:left;background-color:#FFFFFF; margin: 0mm 0mm 0mm 0mm; \">");
	bdyy.Append("<style>\r\n");
	bdyy.Append("@media print {\r\n");
	//bdyy.Append(divAvoidPageBreak);
	//bdyy.Append("\r\n");
	bdyy.Append(preStyle);
	bdyy.Append("\r\n}");
	bdyy.Append(preStyle);
	bdyy.Append("\r\n@page {size: auto; margin: 12mm 4mm 12mm 6mm; }");

	bdyy.Append("\r\n</style>");

	CStringA margin = "margin:0mm 8mm 0mm 12mm;";
	margin = "margin:0mm 0mm 0mm 0mm;";
	margin = "margin-left:5px;";

	fp.Write(bdyy, bdyy.GetLength());

	CString* srcImgFilePath = 0; // UpdateInlineSrcImgPathEx will generate srcImgFilePath
	AttachmentMgr attachmentDB;
	attachmentDB.Clear();


	bool extraHtmlHdr = false;
	if (outbuflarge->Count() != 0)
	{
		// FIXME drop or enhance
		//int retBG = NListView::RemoveBackgroundColor(outbuflarge->Data(), outbuflarge->Count(), &tmpbuf, mailPosition);
		//outbuflarge->Copy(tmpbuf);

		CStringA bdycharset = "UTF-8";
		CStringA bdy;

		BOOL createEmbeddedImageFiles = TRUE;  // FIXME ???

		SimpleString* outbuf = MboxMail::m_workbuf;
		outbuf->ClearAndResize(outbuflarge->Count() * 2);

		char *inData = outbuflarge->Data();
		int inDataLen = outbuflarge->Count();

		AttachmentMgr embededAttachmentDB;

		CString absoluteSrcImgFilePath = aboluteImageCachePath;

		CString cacheSubFolder = L"ImageCache";
		CString relativeSrcImgFilePath = pref + cacheSubFolder + L"\\";

		if (NListView::m_exportMailsMode)
		{
			cacheSubFolder = L"ExportCache";
			relativeSrcImgFilePath = pref + L"Attachments\\";  // TODOE
		}

		BOOL verifyAttachmentDataAsImageType = FALSE;
		BOOL insertMaxWidthForImgTag = TRUE;
		CStringA maxWidth = "100%";
		CStringA maxHeight = "";
		BOOL makeFileNameUnique = TRUE;
		BOOL makeAbsoluteImageFilePath = NListView::m_fullImgFilePath;
		NListView::UpdateInlineSrcImgPathEx(&fpm, inData, inDataLen, outbuf, makeFileNameUnique, makeAbsoluteImageFilePath,
			relativeSrcImgFilePath, absoluteSrcImgFilePath, embededAttachmentDB,
			pListView->m_EmbededImagesStats, mailPosition, createEmbeddedImageFiles, verifyAttachmentDataAsImageType, insertMaxWidthForImgTag,
			maxWidth, maxHeight);

		if (outbuflarge->Count() == workbuf->Count())
			int deb = 1;

		if (workbuf->Count() > 0)
			outbuflarge->Copy(*workbuf);
		else
			int deb = 1;

#if 0
#if 0
		// consider adding global style  <style> https://www.w3docs.com/snippets/css/how-to-wrap-text-in-a-pre-tag-with-css.html
		pre{
			white-space: pre-wrap;       /* Since CSS 2.1 */
			white-space: -moz-pre-wrap;  /* Mozilla, since 1999 */
			white-space: -pre-wrap;      /* Opera 4-6 */
			white-space: -o-pre-wrap;    /* Opera 7 */
			word-wrap: break-word;       /* Internet Explorer 5.5+ */
		}
#endif
		workbuf->ClearAndResize(outbuflarge->Count() * 2);
		int retPrep = NListView::ReplacePreTagWitPTag((char*)outbuflarge->Data(), outbuflarge->Count(), workbuf, TRUE);

		if (workbuf->Count())
		{
			outbuflarge->Copy(*workbuf);
		}
		else
			int deb = 1;
#endif

		workbuf->ClearAndResize(outbuflarge->Count() * 2);
		int retWidth = NListView::AddMaxWidthToHref((char*)outbuflarge->Data(), outbuflarge->Count(), workbuf, TRUE);

		if (workbuf->Count())
			outbuflarge->Copy(*workbuf);
		else
			int deb = 1;

#if 0
		workbuf->ClearAndResize(outbuflarge->Count() * 2);
		// Drop or proof it helps
		retWidth = NListView::AddMaxWidthToDIV((char*)outbuflarge->Data(), outbuflarge->Count(), workbuf, TRUE);

		if (workbuf->Count())
			outbuflarge->Copy(*workbuf);
		else
			int deb = 1;
#endif

#if 1
		// Fix Blo0cquote tag to avoid line truncation
		workbuf->ClearAndResize(outbuflarge->Count() * 2);
		BOOL ReplaceAllWhiteBackgrounTags = TRUE;
		retWidth = NListView::ReplaceBlockqouteTagWithDivTag((char*)outbuflarge->Data(), outbuflarge->Count(), workbuf, ReplaceAllWhiteBackgrounTags);

		if (workbuf->Count())
			outbuflarge->Copy(*workbuf);
		else
			int deb = 1;
#endif

#if 0
		workbuf->ClearAndResize(outbuflarge->Count() * 2);
		// Fix Blo0cquote tag to avoid line truncation
		// ReplaceBlockqouteTagWithDivTag seem to work better than AddMaxWidthToBlockquote
		// May need to enhance AddMaxWidthToBlockquote to update existing style= instead of adding new style=  FIXME
		//
		retWidth = NListView::AddMaxWidthToBlockquote((char*)outbuflarge->Data(), outbuflarge->Count(), workbuf, TRUE);

		if (workbuf->Count())
			outbuflarge->Copy(*workbuf);
		else
			int deb = 1;
#endif
		// Remove color and width from body tag, add later
		CStringA bodyBackgroundColor;
		CStringA bodyWidth;
		// if (!singleMail)    // TODO: do it for single mail also
		{
			BOOL removeBackgroundColor = TRUE;
			BOOL removeWidth = TRUE;

			// workbuf is set inside RemoveBodyBackgroundColorAndWidth
			int retval = NListView::RemoveBodyBackgroundColorAndWidth((char*)outbuflarge->Data(), outbuflarge->Count(), workbuf,
				bodyBackgroundColor, bodyWidth, removeBackgroundColor, removeWidth);

			if (!bodyWidth.IsEmpty() || !bodyBackgroundColor.IsEmpty())
			{
				outbuflarge->Copy(*workbuf);
			}
		}

		//bdy = "\r\n<div style=\'width:100%;position:initial;float:left;background-color:transparent;" + margin + "text-align:left\'>\r\n";
		CStringA newBodyWidth = "width:";
		newBodyWidth.Append("100%;");
		CStringA newBodyMargin = margin;
		bdy = "\r\n<div style=\"page-break-inside:avoid;position:initial;float:left;background-color:transparent;text-align:left;"
			+ newBodyWidth + newBodyMargin
			+ "\">\r\n";   // div #1

		fp.Write(bdy, bdy.GetLength());

		CStringA font = "\">\r\n";

		// Setp mail header
		CStringA hdrBackgroundColor = "background-color:#eee9e9;";
		if (pFrame && (pFrame->m_NamePatternParams.m_bAddBackgroundColorToMailHeader == FALSE))
			hdrBackgroundColor = "";

		CStringA hdrWidth = "width:";
		hdrWidth.Append("100%;");
		CStringA hdrMargin = "";
		CStringA base;
		if (NListView::m_exportMailsMode)
		{
			base = "<base href=\"\"/>";  // TODOE define empty base ??
		}
		if (pFrame)
		{
			bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + ";\">"
				+ base
				+ htmlHdrFldNameStyle + "</head>\r\n"
				+ "<body bgColor=#ffffff>\r\n<div style=\"position:initial;float:left;text-align:left;font-weight:normal;"   // div #2
				+ hdrWidth + hdrMargin + hdrBackgroundColor
				+ font;
			fp.Write(bdy, bdy.GetLength());
		}
		CString mboxFileName = fpm.GetFileName();
		CString htmlFileName = fp.GetFileName();

		BOOL printAttachments = TRUE;
		if (pFrame)
		{
			printAttachments = pFrame->m_HdrFldConfig.m_HdrFldList.IsFldSet(HdrFldList::HDR_FLD_ATTACHMENTS - 1);
		}

		if (printAttachments || NListView::m_appendAttachmentPictures)
		{
			//AttachmentMgr attachmentDB;
			//attachmentDB.Clear();
			CString* attachmentFolderPath = 0;
			BOOL prependMailId = TRUE;
			
			NListView::CreateMailAttachments(&fpm, mailPosition, attachmentFolderPath, prependMailId, attachmentDB);
		}

		int ret;
		if (pFrame)
		{
			ret = printMailHeaderToHtmlFile(fp, mailPosition, fpm, textConfig, pFrame->m_HdrFldConfig, attachmentFilesFolderPath);
		}

		bdy = "</div></body></html>";   // --> end of div #2
		fp.Write(bdy, bdy.GetLength());

		bdy = "\r\n</div>";   // --> end of div #1
		fp.Write(bdy, bdy.GetLength());

		// ---> Setup mail header done

		// General body setup, is this needed ??
		// if (!singleMail)  // do it for single and multiple mails
		bodyWidth = "width:";
		bodyWidth.Append(scale);
		if (1)
		{
			CStringA fontColor = "black";  // TODO: probably no harm done setting to black always
			CStringA bdyBackgroundColor;
			if (bodyBackgroundColor.IsEmpty())
			{
				bdyBackgroundColor = "background-color:#FFFFFF;color:black;";
			}
			else // if (!bodyBackgroundColor.IsEmpty())
			{
				bdyBackgroundColor.Append("background-color:");
				bdyBackgroundColor.Append(bodyBackgroundColor);
				bdyBackgroundColor.Append(";color:");
				bdyBackgroundColor.Append(fontColor);
				bdyBackgroundColor.Append(";");
			}

			bdy = "\r\n<div style=\"position:initial;float:left;text-align:left;overflow-wrap:break-word !important;"
				+ bodyWidth + margin + bdyBackgroundColor
				+ "\"><br>\r\n";

			fp.Write(bdy, bdy.GetLength());
		}

		// --> Done with general Body Setup

		// Setup if no html and body tags; moved up, global
		//bool extraHtmlHdr = false;
		int htmlTagPos = outbuflarge->FindNoCase(0, "<html", 5);
		if (htmlTagPos  < 0) // didn't find if true
		{
			if (outbuflarge->FindNoCase(0, "<body", 5) < 0) // didn't find if true
			{
				extraHtmlHdr = true;
			}
		}
		else
		{
			// remove cClassWordSection1 & ClassSection1 from html text otherwise Chrome and Edge will generate page break in PDF
			static char* cClassWordSection1 = "class=wordsection1";
			static int cClassWordSection1Len = istrlen(cClassWordSection1);
			static char* cClassSection1 = "class=section1";
			static int cClassSection1Len = istrlen(cClassSection1);

			static char* cOfficeWord = "office:word";
			static int cOfficeWordLen = istrlen(cOfficeWord);

			int headTagPos = outbuflarge->FindNoCase(htmlTagPos, "<head", 5);
			if (headTagPos > 0)
			{
				const int wordTagPos = outbuflarge->FindNoCase(htmlTagPos, cOfficeWord, cOfficeWordLen, headTagPos);
				if (wordTagPos > 0)
				{
					char* beg = outbuflarge->Data(headTagPos);  // helper to view
					int ret = NListView::RemoveMicrosoftSection1Class(outbuflarge, headTagPos);

					int deb = 1;
				}
			
			}
			const int deb = 1;
		}

		bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + ";\">";
		bdy.Append("<style>\r\n");
		bdy.Append(preStyle);
		bdy.Append("\r\n</style>");
		bdy.Append("</head>");
		if (extraHtmlHdr == false)
			bdy.Append("</html>\r\n");
		else
			bdy.Append("<body bgColor=#ffffff>\r\n");  // append </html> after message text

		fp.Write(bdy, bdy.GetLength());

		std::string charSet;
		if (pageCode == 0) {
			pageCode = getCodePageFromHtmlBody(outbuflarge, charSet);
		}

		if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
		{
			int needLength = outbuflarge->Count() * 2 + 1;
			inbuf->ClearAndResize(needLength);

			BOOL ret = TextUtilsEx::Str2CodePage(outbuflarge, pageCode, textConfig.m_nCodePageId, inbuf, workbuf, error);
			if (ret) {
				fp.Write(inbuf->Data(), inbuf->Count());
			}
			else {
				fp.Write(outbuflarge->Data(), outbuflarge->Count());
			}
		}
		else
			fp.Write(outbuflarge->Data(), outbuflarge->Count());

#if 0

		if (extraHtmlHdr) {
			bdy = "\r\n</body></html>";
			fp.Write(bdy, bdy.GetLength());
		}

		bdy = "\r\n</div>";
		fp.Write(bdy, bdy.GetLength());
#endif
	}
	else
	{
		outbuflarge->Clear();
		pageCode = 0;
		textType = 0; // !!!!!!!!  no Html, try to get Plain  !!!!!!
		int plainTextMode = 2;  // insert <img src=attachment name> image tags
		textlen = GetMailBody_mboxview(fpm, mailPosition, outbuflarge, pageCode, textType, plainTextMode);  // returns pageCode
		if (textlen != outbuflarge->Count())
			int deb = 1;

		BOOL createEmbeddedImageFiles = TRUE;  // FIXME ???

		SimpleString* outbuf = MboxMail::m_workbuf;
		outbuf->ClearAndResize(outbuflarge->Count() * 2);

		char* inData = outbuflarge->Data();
		int inDataLen = outbuflarge->Count();

		AttachmentMgr embededAttachmentDB;
		//CString* srcImgFilePath = 0; // UpdateInlineSrcImgPathEx will generate srcImgFilePath

		CString absoluteSrcImgFilePath = aboluteImageCachePath;
		CString cacheSubFolder = L"ImageCache";
		CString relativeSrcImgFilePath = pref + cacheSubFolder + L"\\";
		if (NListView::m_exportMailsMode)
		{
			cacheSubFolder = L"ExportCache";
			relativeSrcImgFilePath = pref + L"Attachments\\";  // TODOE
		}
		
		BOOL verifyAttachmentDataAsImageType = FALSE;
		BOOL insertMaxWidthForImgTag = TRUE;
		CStringA maxWidth = "40%";
		CStringA maxHeight = "";
		BOOL makeFileNameUnique = TRUE;
		BOOL makeAbsoluteImageFilePath = NListView::m_fullImgFilePath;
		NListView::UpdateInlineSrcImgPathEx(&fpm, inData, inDataLen, outbuf, makeFileNameUnique, makeAbsoluteImageFilePath,
			relativeSrcImgFilePath, absoluteSrcImgFilePath, embededAttachmentDB,
			pListView->m_EmbededImagesStats, mailPosition, createEmbeddedImageFiles, verifyAttachmentDataAsImageType, insertMaxWidthForImgTag,
			maxWidth, maxHeight);

		if (outbuflarge->Count() == outbuf->Count())
			int deb = 1;

		if (outbuf->Count() > 0)
			outbuflarge->Copy(*outbuf);
		else
			int deb = 1;

		CStringA bdy;
		CStringA bdycharset = "UTF-8";

		CStringA newBodyWidth = "width:";
		newBodyWidth.Append("100%;");
		CStringA newBodyMargin = margin;
		bdy = "\r\n<div style=\"position:initial;float:left;background-color:transparent;text-align:left;"
			+ newBodyWidth + newBodyMargin
			+ "\">\r\n";   // div #1

		fp.Write(bdy, bdy.GetLength());

		CStringA font = "\">\r\n";

		// Setp mail header
		CStringA hdrBackgroundColor = "background-color:#eee9e9;";
		if (pFrame && (pFrame->m_NamePatternParams.m_bAddBackgroundColorToMailHeader == FALSE))
			hdrBackgroundColor = "";

		CStringA hdrWidth = "width:";
		hdrWidth.Append(scale);
		CStringA hdrMargin = "";
		if (pFrame)
		{
			bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + ";\">"
				+ htmlHdrFldNameStyle + "</head>\r\n"
				+ "<body bgColor=#ffffff>\r\n<div style=\"position:initial;float:left;text-align:left;font-weight:normal;"   // div #2
				+ hdrWidth + hdrMargin + hdrBackgroundColor
				+ font;
			fp.Write(bdy, bdy.GetLength());
		}

		BOOL printAttachments = TRUE;
		if (pFrame)
		{
			printAttachments = pFrame->m_HdrFldConfig.m_HdrFldList.IsFldSet(HdrFldList::HDR_FLD_ATTACHMENTS - 1);
		}

		if (printAttachments || NListView::m_appendAttachmentPictures)
		{
			//AttachmentMgr attachmentDB;
			//attachmentDB.Clear();
			CString* attachmentFolderPath = 0;
			BOOL prependMailId = TRUE;

			NListView::CreateMailAttachments(&fpm, mailPosition, attachmentFolderPath, prependMailId, attachmentDB);
		}

		int ret;
		if (pFrame)
			ret = printMailHeaderToHtmlFile(fp, mailPosition, fpm, textConfig, pFrame->m_HdrFldConfig, attachmentFilesFolderPath);

		bdy = "</div></body></html>";
		fp.Write(bdy, bdy.GetLength());

		bdy = "\r\n</div>";
		fp.Write(bdy, bdy.GetLength());

		CStringA bodyWidth = "width:";
		bodyWidth.Append(scale);
		CStringA bodyBackgroundColor = "background-color:#FFFFFF;";
		//bdy = "\r\n<div style=\'width:auto;position:initial;float:left;color:black;background-color:#FFFFFF;" + margin + "text-align:left\'>\r\n";
		bdy = "\r\n<div style=\"position:initial;float:left;text-align:left;color:black;overflow-wrap:break-word !important;"
			+ bodyWidth + margin + bodyBackgroundColor
			+ "\">\r\n";
		fp.Write(bdy, bdy.GetLength());

		bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\">";
		bdy.Append("<style>\r\n");
		bdy.Append(preStyle);
		bdy.Append("\r\n</style></head><body bgColor=#ffffff><br>\r\n");
		fp.Write(bdy, bdy.GetLength());

		if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
		{
			int needLength = outbuflarge->Count() * 2 + 1;
			inbuf->ClearAndResize(needLength);

			BOOL ret = TextUtilsEx::Str2CodePage(outbuflarge, pageCode, textConfig.m_nCodePageId, inbuf, workbuf, error);
			// TODO: replace CR LF with <br> to enable line wrapping
			if (ret)
			{
				char *inData = inbuf->Data();
				int inDataLen = inbuf->Count();
				workbuf->ClearAndResize(3 * inDataLen);
				TextUtilsEx::EncodeAsHtmlText(inData, inDataLen, workbuf);

				fp.Write(workbuf->Data(), workbuf->Count());
			}
			else
			{
				char *inData = outbuflarge->Data();
				int inDataLen = outbuflarge->Count();
				workbuf->ClearAndResize(3 * inDataLen);
				TextUtilsEx::EncodeAsHtmlText(inData, inDataLen, workbuf);

				fp.Write(workbuf->Data(), workbuf->Count());
			}
		}
		else
		{
			char *inData = outbuflarge->Data();
			int inDataLen = outbuflarge->Count();
			workbuf->ClearAndResize(3 * inDataLen);
			TextUtilsEx::EncodeAsHtmlText(inData, inDataLen, workbuf);

			fp.Write(workbuf->Data(), workbuf->Count());
		}
		extraHtmlHdr = TRUE;
#if 0
		bdy = "</body></html>";
		fp.Write(bdy, bdy.GetLength());

		bdy = "\r\n</div>";
		fp.Write(bdy, bdy.GetLength());
#endif
	}

#if 0

	/// List attachment names at the end of Mailk instead in the beginning ogf the mail
	SimpleString htmlbuf(1024);
	CString attachmentFimenamePrefix;
	printAttachmentNamesAsHtml(&fpm, mailPosition, &htmlbuf, attachmentFimenamePrefix);

	fp.Write(htmlbuf.Data(), htmlbuf.Count());
#endif

	attachmentDB.Sort();

	if (NListView::m_appendAttachmentPictures)
	{
		CString relativeFolderPath = attachmentFilesFolderPath;

		// AppendPictureAttachments prefers srcImgFilePath if != 0
		CString* pRelativeFolderPath = &relativeFolderPath;
		CString* pAbsoluteFolderPath = &absoluteAttachmentCachePath;
		if (!NListView::m_fullImgFilePath)
			pAbsoluteFolderPath = 0;
		NListView::AppendPictureAttachments(m, attachmentDB, pAbsoluteFolderPath, pRelativeFolderPath, &fp);
	}

	CStringA bdy;

	if (extraHtmlHdr == true)
	{
		bdy = "</body></html>";
		fp.Write(bdy, bdy.GetLength());
	}

	bdy = "\r\n</div>";
	fp.Write(bdy, bdy.GetLength());


#if 0
	bdy = "\r\n<footer style=\"position:fixed; bottom:-8mm; left : 0;\"><p>";
	bdy.Append(fpm.GetFileName());
	bdy.Append("\r\n</p></footer>");
	fp.Write(bdy, bdy.GetLength());

	bdy = "\r\n<header style=\"position:fixed; top:-8mm; left:0;\"><p>";
	bdy.Append(fpm.GetFileName());
	bdy.Append("\r\n</p></header>");
	fp.Write(bdy, bdy.GetLength());
#endif

	bdy = "\r\n</article>";
	fp.Write(bdy, bdy.GetLength());

	bdy = "\r\n<div>&nbsp;<br></div>";
	fp.Write(bdy, bdy.GetLength());

	if (addPageBreak)
	{
		bdy = "\r\n<div style=\"page-break-before:always\"></div>\r\n";
		fp.Write(bdy, bdy.GetLength());
	}
	else
	{
		; //CString bdy = "\r\n<div>&nbsp;</div>";
		//fp.Write(bdy, bdy.GetLength());
	}

	return 1;
}

void MboxMail::encodeTextAsHtml(SimpleString &txt)
{
	SimpleString buffer(256);
	for (int pos = 0; pos != txt.Count(); ++pos) {
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

void MboxMail::encodeTextAsHtmlLink(SimpleString &txt)
{
	SimpleString buffer(256);
	for (int pos = 0; pos != txt.Count(); ++pos) {
		char c = txt.GetAt(pos);
		switch (c) {
		case ' ':  buffer.Append("%20");         break;
		default:   buffer.Append(c); break;
		}
	}
	txt.Copy(buffer);
}

void MboxMail::encodeTextAsHtmlLinkLabel(SimpleString &txt)
{
	SimpleString buffer(256);
	for (int pos = 0; pos != txt.Count(); ++pos)
	{
		const char c = txt.GetAt(pos);
		switch (c) {
		case '&':  buffer.Append("&amp;");       break;
		case '\"': buffer.Append("&quot;");      break;
		case '\'': buffer.Append("&apos;");      break;
		case '<':  buffer.Append("&lt;");        break;
		case '>':  buffer.Append("&gt;");        break;
		case ' ':  buffer.Append("&#95;");       break;
		case '-':  buffer.Append("&minus;");     break;
		case '_':  buffer.Append("&#95;");       break;
		default:   buffer.Append(c); break;
		}
	}
	txt.Copy(buffer);
}

void MboxMail::encodeTextAsHtmlLinkLabel(CStringA link, CStringA& htmlLink)
{
#if 0
	SimpleString txt;
	txt.Append((LPCSTR)link, link.GetLength());
	MboxMail::encodeTextAsHtmlLinkLabel(txt);
	htmlLink.Append(txt.Data(), txt.Count());
#else
	SimpleString buffer(256);
	for (int pos = 0; pos != link.GetLength(); ++pos) {
		char c = link.GetAt(pos);
		switch (c) {
		case '&':  htmlLink.Append("&amp;");       break;
		case '\"': htmlLink.Append("&quot;");      break;
		case '\'': htmlLink.Append("&apos;");      break;
		case '<':  htmlLink.Append("&lt;");        break;
		case '>':  htmlLink.Append("&gt;");        break;
		default:   htmlLink.AppendChar(c); break;
		}
	}
#endif
}

int MboxMail::CreateFldFontStyle(HdrFldConfig &hdrFieldConfig, CStringA &fldNameFontStyle, CStringA &fldTextFontStyle)
{
	const int lineHeight = 120;
	if (hdrFieldConfig.m_bHdrFontDflt == 0)
	{
		int nameFontSize = hdrFieldConfig.m_nHdrFontSize;
		int textFontSize = hdrFieldConfig.m_nHdrFontSize;

		fldTextFontStyle.Format("overflow-wrap:break-word;color:black;font-size:%dpx;line-height:%d%%;", textFontSize, lineHeight);

		if (hdrFieldConfig.m_bHdrBoldFldName)
		{
			fldNameFontStyle.Format("color:black;font-weight:bold;font-size:%dpx;line-height:%d%%;", nameFontSize, lineHeight);
		}
		else
		{
			fldNameFontStyle.Format("color:black;font-size:%dpx; line-height:%d%%;", nameFontSize, lineHeight);
		}
	}
	else
	{
		// hdrFieldConfig.m_HdrFldFontName.m_fontStyleName not used due to problems trying to use

		int nameFontSize = hdrFieldConfig.m_HdrFldFontName.m_nFontSize;
		int textFontSize = hdrFieldConfig.m_HdrFldFontText.m_nFontSize;
		//
		CStringA nameFont = hdrFieldConfig.m_HdrFldFontName.m_fontName;
		CStringA textFont = hdrFieldConfig.m_HdrFldFontText.m_fontName;

		//
		CStringA nameFontFamily = hdrFieldConfig.m_HdrFldFontName.m_genericFontName;
		CStringA textFontFamily = hdrFieldConfig.m_HdrFldFontText.m_genericFontName;
		//

		CStringA nameFontWeight;
		if (hdrFieldConfig.m_HdrFldFontName.m_nFontStyle == 400)
			nameFontWeight = "normal";
		else if (hdrFieldConfig.m_HdrFldFontName.m_bIsBold)
			nameFontWeight = "bold";
		else
			nameFontWeight.Format("%d", hdrFieldConfig.m_HdrFldFontName.m_nFontStyle);

		CStringA textFontWeight;
		if (hdrFieldConfig.m_HdrFldFontText.m_nFontStyle == 400)
			textFontWeight = "normal";
		else if (hdrFieldConfig.m_HdrFldFontText.m_bIsBold)
			textFontWeight = "bold";
		else
			textFontWeight.Format("%d", hdrFieldConfig.m_HdrFldFontText.m_nFontStyle);

		CStringA nameFontStyle = "normal";
		if (hdrFieldConfig.m_HdrFldFontName.m_bIsItalic)
			nameFontStyle = "italic";

		CStringA textFontStyle = "normal";
		if (hdrFieldConfig.m_HdrFldFontText.m_bIsItalic)
			textFontStyle = "italic";

		fldNameFontStyle.Format("color:black;font-style:%s;font-weight:%s;font-size:%dpx;line-height:%d%%;font-family:\"%s\",%s;",
			nameFontStyle, nameFontWeight, nameFontSize, lineHeight, nameFont, nameFontFamily);

		fldTextFontStyle.Format("overflow-wrap:break-word;color:black;font-style:%s;font-weight:%s;font-size:%dpx;line-height:%d%%;font-family:\"%s\",%s;",
			textFontStyle, textFontWeight, textFontSize, lineHeight, textFont, textFontFamily);
	}
	return 1;
}

int MboxMail::printMailHeaderToHtmlFile(/*out*/CFile &fp, int mailPosition, /*in*/ CFile &fpm, TEXTFILE_CONFIG &textConfig, HdrFldConfig &hdrFieldConfig, CString &attachmentFilesFolderPath)
{
	int fldNumb;
	int fldNameLength;
	int fldTxtLength;
	char *fldText;
	char *fldName = "";
	int ii;
	DWORD error;

	SimpleString outbuf(1024, 256);
	SimpleString tmpbuf(1024, 256);

	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;
	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);

	MboxMail *m = s_mails[mailPosition];

	UINT pageCode = 0;
	CStringA txt;
	for (ii = 0; ii < HdrFldList::HDR_FLD_MAX; ii++)
	{
		if (hdrFieldConfig.m_HdrFldList.IsFldSet(ii))
		{
			fldNumb = ii+1;
			if (fldNumb == HdrFldList::HDR_FLD_SUBJECT)
			{
				fldName = "SUBJECT: ";
				fldNameLength = istrlen(fldName);
				fldTxtLength = m->m_subj.GetLength();
				fldText = (char *)(LPCSTR)m->m_subj;
				pageCode = m->m_subj_charsetId;
			}
			else if (fldNumb == HdrFldList::HDR_FLD_FROM)
			{
				fldName = "FROM: ";
				fldNameLength = istrlen(fldName);
				fldTxtLength = m->m_from.GetLength();
				fldText = (char *)(LPCSTR)m->m_from;
				pageCode = m->m_from_charsetId;
			}
			else if (fldNumb == HdrFldList::HDR_FLD_TO)
			{
				fldName = "TO: ";
				fldNameLength = istrlen(fldName);
				fldTxtLength = m->m_to.GetLength();
				fldText = (char *)(LPCSTR)m->m_to;
				pageCode = m->m_to_charsetId;
			}
			else if (fldNumb == HdrFldList::HDR_FLD_CC)
			{
				fldName = "CC: ";
				fldNameLength = istrlen(fldName);
				fldTxtLength = m->m_cc.GetLength();
				fldText = (char *)(LPCSTR)m->m_cc;
				pageCode = m->m_cc_charsetId;
				if (fldTxtLength <= 0)
					continue;
			}
			else if (fldNumb == HdrFldList::HDR_FLD_BCC)
			{
				fldName = "BCC: ";
				fldNameLength = istrlen(fldName);
				fldTxtLength = m->m_bcc.GetLength();
				fldText = (char *)(LPCSTR)m->m_bcc;
				pageCode = m->m_bcc_charsetId;
				if (fldTxtLength <= 0)
					continue;
			}
			else if (fldNumb == HdrFldList::HDR_FLD_DATE)
			{
				fldName = "DATE: ";
				fldNameLength = istrlen(fldName);
				fldTxtLength = 0;
				fldText = "";
				continue;
			}
			else if (fldNumb == HdrFldList::HDR_FLD_ATTACHMENTS)
			{
				fldName = "DATE: ";
				fldNameLength = istrlen(fldName);
				fldTxtLength = 0;
				fldText = "";
				continue;
			}
			else
			{
				continue;
			}

			txt.Format("<span class=\'hdrfldname\'>");
			outbuf.Append(txt, txt.GetLength());
			outbuf.Append(fldName, fldNameLength);
			outbuf.Append("</span>");

			tmpbuf.Copy(fldText, fldTxtLength);

			txt.Format("<span class=\'hdrfldtext\'>");
			outbuf.Append(txt, txt.GetLength());

			if (textConfig.m_nCodePageId && (pageCode != 0) && (pageCode != textConfig.m_nCodePageId))
			{
				BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf, error);
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
			else 
			{
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

		CStringA format = textConfig.m_dateFormat;

		datebuff[0] = 0;
		if (m->m_timeDate >= 0)
		{
			MyCTime tt(m->m_timeDate);
			if (!textConfig.m_bGMTTime)
			{
				CStringA lDateTime = tt.FormatLocalTmA(format);
				strcpy(datebuff, (LPCSTR)lDateTime);
			}
			else 
			{
				CStringA lDateTime = tt.FormatGmtTmA(format);
				strcpy(datebuff, (LPCSTR)lDateTime);
			}
		}

		fldName = "DATE: ";
		fldNameLength = istrlen(fldName);
		fldTxtLength = m->m_subj.GetLength();
		fldText = (char *)datebuff;

		txt.Format("<span class=\'hdrfldname\'>");
		outbuf.Append(txt, txt.GetLength());
		outbuf.Append(fldName, fldNameLength);
		outbuf.Append("</span>");

		txt.Format("<span class=\'hdrfldtext\'>");
		outbuf.Append(txt, txt.GetLength());

		int datelen = istrlen(datebuff);
		outbuf.Append(datebuff, datelen);

		outbuf.Append("</span>");
		outbuf.Append("<br>\r\n");
	}

	fldNumb = HdrFldList::HDR_FLD_ATTACHMENTS -1;
	if (hdrFieldConfig.m_HdrFldList.IsFldSet(fldNumb))
	{
		tmpbuf.Clear();

		CString attachmentFileNamePrefix;
		//int ret = MboxMail::printAttachmentNamesAsText(&fpm, mailPosition, &tmpbuf, attachmentFileNamePrefix);
		int ret = MboxMail::printAttachmentNamesAsHtml(&fpm, mailPosition, &tmpbuf, attachmentFileNamePrefix, attachmentFilesFolderPath);

		if (tmpbuf.Count() > 0)
		{
			//encodeTextAsHtml(tmpbuf);


			CStringA attachmentFileNamePrefixA = attachmentFileNamePrefix;
			CStringA label;
			label.Format("ATTACHMENTS (%s):", (LPCSTR)attachmentFileNamePrefixA);
			fldName = (char*)(LPCSTR)label;
			fldNameLength = label.GetLength();

			txt.Format("<span class=\'hdrfldname\'>");
			outbuf.Append(txt, txt.GetLength());
			outbuf.Append(fldName, fldNameLength);
			outbuf.Append("</span>");

			txt.Format("<span class=\'hdrfldtext\'>");
			outbuf.Append(txt, txt.GetLength());

			outbuf.Append(tmpbuf.Data(), tmpbuf.Count());

			outbuf.Append("</span>");
			outbuf.Append("<br>\r\n");
		}
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
	TRACE(L"(ALongRightProcessProcPrintMailArchive) threadId=%ld threadPriority=%ld\n", myThreadId, myThreadPri);

	BOOL progressBar = TRUE;
	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as param to printMailArchiveToTextFile
	CString errorText;
	int retpos = MboxMail::printMailArchiveToTextFile(args->textConfig, args->textFile, args->firstMail, args->lastMail, args->textType, progressBar, errorText);
	args->errorText = errorText;
	args->exitted = TRUE;
	return true;
}

BOOL MboxMail::GetLongestCachePath(CString &longestCachePath) // FIXMEFIXME
{
	CString rootCacheSubFolder = L"AttachmentCache";
	rootCacheSubFolder = L"PrintCache";

	CString datapath = MboxMail::GetLastDataPath();
	if (datapath.IsEmpty())
		return FALSE;

	CString fileDataPath = MboxMail::s_datapath;
	CString filePath = MboxMail::s_path;
	if (s_path.GetLength() == 0)
		return FALSE;

	longestCachePath = datapath + rootCacheSubFolder;

	CString errorText;
	CString targetPrintSubFolder;

	longestCachePath.Empty();

	BOOL ret = CreateCachePath(rootCacheSubFolder, targetPrintSubFolder, longestCachePath, errorText);
	if (ret)
		return TRUE;
	else
		return FALSE;

	return TRUE;
}

BOOL MboxMail::GetPrintCachePath(CString &prtCachePath)
{
	CString errorText;
	CString rootPrintSubFolder = L"PrintCache";
	CString targetPrintSubFolder;

	BOOL ret = CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, prtCachePath, errorText);
	if (ret)
		return TRUE;
	else
		return FALSE;
}

BOOL MboxMail::GetMboxviewFilePath(CString &mboxFilePath, CString &mboxviewFilePath)
{
	CString mailFileName;
	FileUtils::CPathStripPath(mboxFilePath, mailFileName);
	CString datapath = MboxMail::GetLastDataPath();
	mboxviewFilePath = datapath + mailFileName + L".mboxview";
	return TRUE;
}

BOOL MboxMail::GetFileCachePathWithoutExtension(CString &mboxFilePath, CString &mboxBaseFileCachePath)
{
	CString errorText;
	CString printCachePath;
	CString rootPrintSubFolder = L"PrintCache";
	CString targetPrintSubFolder;

	BOOL ret = CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, printCachePath, errorText);
	if (ret)
		return TRUE;
	else
		return FALSE;
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
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	textFileName = textFile;
	int textFileLength = textFile.GetLength();
	int maxPath = _MAX_PATH;

	if (!progressBar) // no worker thread
	{
		CFileException ExError;
		if (!fp.Open(textFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &ExError))
		{
			DWORD lastErr = ::GetLastError();
#if 1
			//HWND h = GetSafeHwnd();
			HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
			CString fmt = L"Could not create file:\n\n\"%s\"\n\n%s";  // new format
			CString errorText = FileUtils::ProcessCFileFailure(fmt, textFile, ExError, lastErr, h);
#else
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt;
			CString fmt = L"Could not create \"%s\" file.\n%s";
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, textFile, exErrorStr);

			//TRACE(L"%s\n", txt);

			CFileStatus rStatus;
			BOOL ret = fp.GetStatus(rStatus);

			//errorText = txt;

			HWND h = NULL; // we don't have any window yet
			int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
#endif
			return -1;
		}

		CFile fpm;
		CFileException ExError2;
		if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError2))
		{
			DWORD lastErr = ::GetLastError();
#if 1
			//HWND h = GetSafeHwnd();
			HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
			CString fmt = L"Could not open mail file:\n\n\"%s\"\n\n%s";  // new format
			CString errorText = FileUtils::ProcessCFileFailure(fmt, MboxMail::s_path, ExError2, lastErr, h);
#else
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError2);

			CString txt;
			CString fmt = L"Could not open \"%s\" mail file.\n%s";
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, MboxMail::s_path, exErrorStr);

			//TRACE(L"%s\n", txt);
			//errorText = txt;

			HWND h = NULL; // we don't have any window yet
			int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
#endif
			fp.Close();
			return -1;
		}

		if (textType == 0) {
			if (textConfig.m_nCodePageId == CP_UTF8) {
				const char *BOM_UTF8 = "\xEF\xBB\xBF";
				fp.Write(BOM_UTF8, 3);
			}
		}

		BOOL fullImgFilePath = FALSE;  // This is duplication; FIXMEFIXME
		if (NListView::m_fullImgFilePath)
			fullImgFilePath = TRUE;

		AttachmentMgr attachmentDB;
		if (selectedMailIndexList == 0) 
		{
			BOOL singleMail = (firstMail == lastMail) ? TRUE : FALSE;
			for (int i = firstMail; i <= lastMail; i++)
			{
				//BOOL singleMail = (firstMail == lastMail) ? TRUE : FALSE;
				BOOL addPageBreak = MboxMail::PageBreakNeeded(i, i + 1, singleMail);
				if (textType == 0)
				{
					printSingleMailToTextFile(fp, i, fpm, textConfig, singleMail, addPageBreak);
				}
				else
				{
					int pos = printSingleMailToHtmlFile(fp, i, fpm, textConfig, singleMail, addPageBreak, fullImgFilePath);
					if (pos < 0)
						continue;
				}
			}
			if ((textType == 1) && (singleMail == FALSE))
				printDummyMailToHtmlFile(fp);
		}
		else 
		{
			int i;
			firstMail = (*selectedMailIndexList)[0];
			int cnt = (int)selectedMailIndexList->GetCount();

			BOOL singleMail = (cnt == 1) ? TRUE : FALSE;
			for (int j = 0; j < cnt; j++)
			{
				i = (*selectedMailIndexList)[j];

				//BOOL singleMail = (cnt == 1) ? TRUE : FALSE;
				BOOL addPageBreak = MboxMail::PageBreakNeeded(selectedMailIndexList, j, singleMail);

				if (textType == 0)
				{
					printSingleMailToTextFile(fp, i, fpm, textConfig, singleMail, addPageBreak);
				}
				else 
				{
					int pos = printSingleMailToHtmlFile(fp, i, fpm, textConfig, singleMail, addPageBreak, fullImgFilePath);
					if (pos < 0)
						continue;
				}
			}
			if ((textType == 1) && (singleMail == FALSE))
				printDummyMailToHtmlFile(fp);
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
		Dlg.SetDialogTemplate(AfxGetApp()->m_hInstance, MAKEINTRESOURCE(IDD_PROGRESS_DLG), IDC_STATIC, IDC_PROGRESS_BAR, IDCANCEL);

		INT_PTR nResult = Dlg.DoModal();

		if (!nResult) // should never be true ?
		{
			MboxMail::assert_unexpected();
			return -1;
		}

		TRACE(L"m_Html2TextCount=%d MailCount=%d\n", m_Html2TextCount, MboxMail::s_mails.GetCount());

		int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
		int retResult = LOWORD(nResult);

		if (retResult != IDOK)
		{  // IDOK==1, IDCANCEL==2
			// We should be here when user selects Cancel button
			_ASSERTE(cancelledbyUser == TRUE);

			DWORD terminationDelay = Dlg.GetTerminationDelay();
			int loopCnt = (terminationDelay+100)/25;

			ULONGLONG tc_start = GetTickCount64();
			while ((loopCnt-- > 0) && (args.exitted == FALSE))
			{
				Sleep(25);
			}
			ULONGLONG tc_end = GetTickCount64();
			DWORD delta = (DWORD)(tc_end - tc_start);
			TRACE(L"(exportToTextFile)Waited %ld milliseconds for thread to exist.\n", delta);
		}

		MboxMail::pCUPDUPData = NULL;

		if (!args.errorText.IsEmpty()) {
			HWND h = NULL; // we don't have any window yet
			int answer = ::MessageBox(h, args.errorText, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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
		errorText = L"Please open mail file first.";
		return -1;
	}

	CFileException ExError;;
	CFile fp;
	if (!fp.Open(textFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &ExError))
	{
		DWORD lastErr = ::GetLastError();
#if 1
		//HWND h = GetSafeHwnd();
		HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		CString fmt = L"Could not create file:\n\n\"%s\"\n\n%s";  // new format
		errorText = FileUtils::ProcessCFileFailure(fmt, textFile, ExError, lastErr, h);  // it looks like it may  result in duplicate MessageBox ??
#else

		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not create \"" + textFile;
		txt += L"\" file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);

		CFileStatus rStatus;
		BOOL ret = fp.GetStatus(rStatus);


		errorText = txt;
#endif
		return -1;
	}

	CFile fpm;
	CFileException ExError2;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError2))
	{
		DWORD lastErr = ::GetLastError();
#if 1
		//HWND h = GetSafeHwnd();
		HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		CString fmt = L"Could not open mail file:\n\n\"%s\"\n\n%s";  // new format
		errorText = FileUtils::ProcessCFileFailure(fmt, MboxMail::s_path, ExError2, lastErr, h);
#else
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError2);

		CString txt = L"Could not open \"" + MboxMail::s_path;
		txt += L"\" mail file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);
		errorText = txt;
#endif
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
		{
			CString txt = L"Printing mails to TEXT file ...";
			ResHelper::TranslateString(txt);
			MboxMail::pCUPDUPData->SetProgress(txt, 0);
		}
		else
		{
			CString txt = L"Printing mails to HTML file ...";
			ResHelper::TranslateString(txt);
			MboxMail::pCUPDUPData->SetProgress(txt, 0);
		}
	}

	UINT curstep = 1;

	if (progressBar && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	CString fileNum;

	int cnt = lastMail - firstMail;
	if (cnt <= 0)
		cnt = 1;

	AttachmentMgr attachmentDB;

	ULONGLONG workRangeFirstPos = firstMail;
	ULONGLONG workRangeLastPos = lastMail;
	ProgressTimer progressTimer(workRangeFirstPos, workRangeLastPos);

	BOOL singleMail = (firstMail == lastMail) ? TRUE : FALSE;
	for (int i = firstMail; i <= lastMail; i++)
	{
		//BOOL singleMail = (firstMail == lastMail) ? TRUE : FALSE;
		BOOL addPageBreak = MboxMail::PageBreakNeeded(i, i + 1, singleMail);

		if (textType == 0)
		{
			printSingleMailToTextFile(fp, i, fpm, textConfig, singleMail, addPageBreak);
		}
		else 
		{
			BOOL fullImgFilePath = FALSE;
			int pos = printSingleMailToHtmlFile(fp, i, fpm, textConfig, singleMail, addPageBreak, fullImgFilePath);
			if (pos < 0)
				continue;
		}

		if (progressBar && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				int deb = 1;
				break;
			}

			UINT_PTR dwProgressbarPos = 0;
			ULONGLONG workRangePos = i;
			BOOL needToUpdateStatusBar = progressTimer.UpdateWorkPos(workRangePos, dwProgressbarPos);
			if (needToUpdateStatusBar)
			{
				int nFileNum = i - firstMail + 1;
				if (textType == 0)
				{
					CString fmt = L"Printing mails to single TEXT file ... %d of %d";
					ResHelper::TranslateString(fmt);
					fileNum.Format(fmt, nFileNum, cnt);
				}
				else
				{
					CString fmt = L"Printing mails to single HTML file ... %d of %d";
					ResHelper::TranslateString(fmt);
					fileNum.Format(fmt, nFileNum, cnt);
				}

				if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(dwProgressbarPos));

				int debug = 1;
			}
		}
	}

	if ((textType == 1) && (singleMail == FALSE))
		printDummyMailToHtmlFile(fp);

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
		errorText = L"Please open mail file first.";
		return -1;
	}

	CFileException ExError;
	CFile fp;
	if (!fp.Open(csvFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &ExError))
	{
		DWORD lastErr = ::GetLastError();
#if 1
		//HWND h = GetSafeHwnd();
		HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		CString fmt = L"Could not create file:\n\n\"%s\"\n\n%s";  // new format
		errorText = FileUtils::ProcessCFileFailure(fmt, csvFile, ExError, lastErr, h);  // it looks like it may  result in duplicate MessageBox ??
#else
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not create \"" + csvFile;
		txt += L"\" file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);

		CFileStatus rStatus;
		BOOL ret = fp.GetStatus(rStatus);

		errorText = txt;
#endif
		return -1;
	}

	CFile fpm;
	BOOL isFpmOpen = FALSE;
	//if (csvConfig.m_bContent) // open regardless; required by 
	{
		CFileException ExError;
		if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
		{
			DWORD lastErr = ::GetLastError();
#if 1
			//HWND h = GetSafeHwnd();
			HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
			CString fmt = L"Could not open mail file:\n\n\"%s\"\n\n%s";  // new format
			errorText = FileUtils::ProcessCFileFailure(fmt, MboxMail::s_path, ExError, lastErr, h);
#else
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt = L"Could not open \"" + MboxMail::s_path;
			txt += L"\" mail file.\n";
			txt += exErrorStr;

			TRACE(L"%s\n", txt);

			errorText = txt;
#endif
			fp.Close();
			return -1;
		}
	}

	if (progressBar && MboxMail::pCUPDUPData)
	{
		CString txt = L"Printing mails to CSV file ...";
		ResHelper::TranslateString(txt);
		MboxMail::pCUPDUPData->SetProgress(txt, 0);
	}

	UINT curstep = 1;
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

		ULONGLONG workRangeFirstPos = firstMail;
		ULONGLONG workRangeLastPos = lastMail;
		ProgressTimer progressTimer(workRangeFirstPos, workRangeLastPos);

		for (int i = firstMail; i <= lastMail; i++)
		{
			retval = printSingleMailToCSVFile(fp, i, fpm, csvConfig, first);
			first = false;

			if (progressBar && MboxMail::pCUPDUPData)
			{
				if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
					int deb = 1;
					break;
				}

				UINT_PTR dwProgressbarPos = 0;
				ULONGLONG workRangePos = i;
				BOOL needToUpdateStatusBar = progressTimer.UpdateWorkPos(workRangePos, dwProgressbarPos);
				if (needToUpdateStatusBar)
				{
					int nFileNum = i - firstMail + 1;
					//fileNum.Format(L"Printing mails to CSV file ... %d of %d", nFileNum, cnt);

					CString fmt = L"Printing mails to CSV file ... %d of %d";
					ResHelper::TranslateString(fmt);
					fileNum.Format(fmt, nFileNum, cnt);

					if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(dwProgressbarPos));

					int debug = 1;
				}
			}
		}
	}
	else
	{
		int i;
		int cnt = (int)selectedMailIndexList->GetCount();

		ULONGLONG workRangeFirstPos = 0;
		ULONGLONG workRangeLastPos = cnt - 1;
		ProgressTimer progressTimer(workRangeFirstPos, workRangeLastPos);

		for (int j = 0; j < cnt; j++)
		{
			i = (*selectedMailIndexList)[j];
			retval = printSingleMailToCSVFile(fp, i, fpm, csvConfig, first);
			first = false;

			if (progressBar && MboxMail::pCUPDUPData)
			{
				if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
					int deb = 1;
					break;
				}

				UINT_PTR dwProgressbarPos = 0;
				ULONGLONG workRangePos = j;
				BOOL needToUpdateStatusBar = progressTimer.UpdateWorkPos(workRangePos, dwProgressbarPos);
				if (needToUpdateStatusBar)
				{
					int nFileNum = j + 1;
					//fileNum.Format(L"Printing mails to CSV file ... %d of %d", nFileNum, cnt);

					CString fmt = L"Printing mails to CSV file ... %d of %d";
					ResHelper::TranslateString(fmt);
					fileNum.Format(fmt, nFileNum, cnt);

					if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(dwProgressbarPos));

					int debug = 1;
				}
			}
		}
	}

	fp.Close();
	// if (csvConfig.m_bContent)  // was open if we are here
		fpm.Close();

	if (!csvConfig.m_bContent)
		Sleep(500);  // helps to deliver all messages to Dialog and display 100% progress bar

	return 1;
}

int MboxMail::splitMailAddress(const char* buff, int bufflen, CStringA& name, CStringA& addr)
{
	SimpleString nameA;
	SimpleString addrA;

	int ret = MboxMail::splitMailAddress(buff, bufflen, &nameA, &addrA);

	name.Append(nameA.Data(), nameA.Count());
	addr.Append(addrA.Data(), addrA.Count());

	return ret;
}

int MboxMail::splitMailAddress(const char* buff, int bufflen, CString& name, CString& addr)
{
	SimpleString nameA;
	SimpleString addrA;

	int ret = MboxMail::splitMailAddress(buff, bufflen, &nameA, &addrA);

	//name.Append(nameA.Data(), nameA.Count());
	//addr.Append(addrA.Data(), addrA.Count());

	return ret;
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

	int addrlen = IntPtr2Int(token_end - token_begin + 1);
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

	int namelen = IntPtr2Int(token_end - token_begin + 1);
	name->Copy(token_begin, namelen);

	return 1;
}

int MboxMail::nstrcpy(register char *dst, register char *src)
{
	char *src_sv = src;
	char c;
	while ((c = *src++)) { *dst++ = c; }
	*dst = 0;
	int len = IntPtr2Int(src - src_sv - 1);
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
	_ASSERTE(outbuf != inbuf);;

	inbuf->ClearAndResize(10000);

	int bodyLength = body->m_contentLength;
	if ((body->m_contentOffset + body->m_contentLength) > m->m_length) {
		// something is not consistent
		bodyLength = m->m_length - body->m_contentOffset;
		if (bodyLength < 0)
			bodyLength = 0;
	}

	fileOffset = m->m_startOff + body->m_contentOffset;
	_int64 filePos = fpm.Seek(fileOffset, SEEK_SET);

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
int MboxMail::GetMailBody_mboxview(CFile &fpm, int mailPosition, SimpleString *outbuf, UINT &pageCode, int textType, int plainTextMode)
{
	MboxMail *m = MboxMail::s_mails[mailPosition];
	int ret = MboxMail::GetMailBody_mboxview(fpm, m, outbuf, pageCode, textType, plainTextMode);
	return ret;
}


int MboxMail::AppendInlineAttachmentNameSeparatorLine(MailBodyContent* body, int bodyCnt, SimpleString* outbuf, int textType)
{
	if (bodyCnt > 0)
	{
		if (body->m_contentDisposition.CompareNoCase("inline") == 0)
		{
			if (!body->m_attachmentName.IsEmpty())
			{
				if (textType == 0)
				{
					outbuf->Append("\n\n\n----- ");
					outbuf->Append((LPCSTR)body->m_attachmentName, body->m_attachmentName.GetLength());
					outbuf->Append(" ---------------------\n\n");
				}
				else
				{
					CStringA bdycharset = "UTF-8"; // FIXMEFIXME
					//outbuf->Append("<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\">");
					outbuf->Append("\r\n<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=");
					outbuf->Append((LPCSTR)bdycharset, bdycharset.GetLength());
					outbuf->Append("\"><body><span><br><br><br>----- ");
					outbuf->Append((LPCSTR)body->m_attachmentName, body->m_attachmentName.GetLength());
					outbuf->Append(" ---------------------<br><br>");
					outbuf->Append("</span></body></html>\r\n");
				}
			}
		}
	}
	return 0;
}


// Quite messy !!! Need cleaner solution. Find time. FIXMEFIXME
int MboxMail::GetMailBody_mboxview(CFile &fpm, MboxMail *m, SimpleString *outbuf, UINT &pageCode, int textType, int plainTextMode)
{
	// Needs some work to properly process multiple blocks of the same type plain or html // FIXMEFIXME
	//   and with different text encoding; rencode al as UTF8
	// Need to handle Content-Type: multipart/alternative; to properly display multiple html block with embeded images

	_int64 fileOffset;
	int bodyCnt = 0;

	// We are using global buffers so check and assert if we collide. 
	SimpleString *inbuf = MboxMail::m_inbuf;
	_ASSERTE(outbuf != inbuf);

	outbuf->ClearAndResize(10000);
	inbuf->ClearAndResize(10000);

	SimpleString* tmpbuf = MboxMail::m_largelocal3;
	tmpbuf->Clear();

	MailBodyContent *body;
	pageCode = 0;  // return body final page code
	UINT bdyPageCode = 0;  // current page code for accumulated data
	bool pageCodeConflict = false;

	BOOL reencodeCurrent = FALSE;  // need to reencode accumalated body data
	BOOL reencodeNew = FALSE;  // need to reencode new  body data
	BOOL setAsUTF8 = FALSE;
	UINT CP_US_ASCII = 20127;
	DWORD error;

	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		reencodeCurrent = FALSE;
		reencodeNew = FALSE;
		setAsUTF8 = FALSE;

		tmpbuf->Clear();

		body = m->m_ContentDetailsArray[j];

#if 0

		if ((!body->m_attachmentName.IsEmpty()) || 
			(body->m_contentDisposition.CompareNoCase("attachment") == 0)) 
		{
			if ((textType == 0) && (plainTextMode > 0))
			{
#if 1
				// Proper solution might be forward list of text blocks to the caller to decide whether to
				// embedd given text block into merged mail text
				int pos = body->m_contentType.ReverseFind('/');
				if (pos > 0)
				{
					CStringA contentTypeMain;
					CStringA contentTypeExtension;
					contentTypeExtension = body->m_contentType.Mid(pos + 1);
					contentTypeMain = body->m_contentType.Left(pos);
					if (contentTypeMain.CompareNoCase("image") == 0)
					{
						//isValidContentTypeExtension = IsSupportedPictureFileExtension(contentTypeExt);
						if (plainTextMode == 2)
						{
							//CStringA img = "\r\n<img  style=\"max-width:40%;\" src=\"";
							CStringA img = "\r\n<img src=\"";
							img.Append(body->m_attachmentName);
							img.Append("\">\r\n");
							outbuf->Append(img, img.GetLength());
						}
						else
						{
							CStringA img = "\r\n[";
							img.Append(body->m_attachmentName);
							img.Append("]\r\n");
							outbuf->Append(img, img.GetLength());
						}
					}
				}
#endif
			}
			continue;
		}
#else
		//if ((!body->m_attachmentName.IsEmpty()) ||
			//(body->m_contentDisposition.CompareNoCase("attachment") == 0))
		//{
		// textType = 0 plain text, textType = 1 html text
		// plainTextMode = 1 create plain text (print mail as text)
		// plainTextMode = 2 encapsulate plain text as html text (print mail as html)
		if ((textType == 0) && (plainTextMode > 0) && !body->m_attachmentName.IsEmpty())
		{
			// Support emails from iphone. Plain text with embeded images
#if 1
				// Proper solution might be forward list of text blocks to the caller to decide whether to
				// embedd given text block into merged mail text
			int pos = body->m_contentType.ReverseFind('/');
			if (pos > 0)
			{
				CStringA contentTypeMain;
				CStringA contentTypeExtension;
				contentTypeExtension = body->m_contentType.Mid(pos + 1);
				contentTypeMain = body->m_contentType.Left(pos);
				if (contentTypeMain.CompareNoCase("image") == 0)
				{
					//isValidContentTypeExtension = IsSupportedPictureFileExtension(contentTypeExt);
					if (plainTextMode == 2)
					{
						//CStringA img = "\r\n<img  style=\"max-width:40%;\" src=\"";
						CStringA img = "\r\n<img src=\"";
						img.Append(body->m_attachmentName);
						img.Append("\">\r\n");
						outbuf->Append(img, img.GetLength());
					}
					else
					{
						CStringA img = "\r\n[";
						img.Append(body->m_attachmentName);
						img.Append("]\r\n");
						outbuf->Append(img, img.GetLength());
					}
				}
			}
#endif
		}
			//continue;
		//}
#endif
		if (textType == 0)
		{
			if (body->m_contentType.CompareNoCase("text/plain") != 0)
			{
				continue;
			}
		} 
		else if (textType == 1)
		{
			if (body->m_contentType.CompareNoCase("text/html") != 0)
			{
				continue;
			}
		}

#if 1
		//  Fix for content-type of type text 
		if (body->m_contentDisposition.CompareNoCase("attachment") == 0)
		{
			continue;
		}
#endif

		if (bodyCnt == 0)
		{
			bdyPageCode = body->m_pageCode;  // current page code
			pageCode = body->m_pageCode;  // return page code
		}
		else
		{
			UINT newBdyPageCode = body->m_pageCode;

			if (bdyPageCode != newBdyPageCode)
			{
				if (((bdyPageCode == CP_UTF8) || (bdyPageCode == CP_US_ASCII)) && ((newBdyPageCode == CP_UTF8) || (newBdyPageCode == CP_US_ASCII)))
				{
					setAsUTF8 = TRUE;  
				}
				else if ((bdyPageCode == CP_UTF8) || (bdyPageCode == CP_US_ASCII))
				{
					// _ASSERTE((newBdyPageCode != CP_UTF8) && (newBdyPageCode != CP_US_ASCII));
					reencodeNew = TRUE;
				}
				else if ((newBdyPageCode == CP_UTF8) || (newBdyPageCode == CP_US_ASCII))
				{
					// _ASSERTE((bdyPageCode != CP_UTF8) && (bdyPageCode != CP_US_ASCII));
					reencodeCurrent = TRUE;
				}
				else
				{
					reencodeCurrent = TRUE;
					reencodeNew = TRUE;
				}
				if (textType == 1)  // HTML case
				{
					if (newBdyPageCode == CP_US_ASCII)
						pageCode = bdyPageCode;
					else if (bdyPageCode == CP_US_ASCII)
						pageCode = newBdyPageCode;
				}
				else
				{
					pageCode = newBdyPageCode;
				}
			}
		}
		if (textType == 1)  // HTML
		{
			// May need to reencode or if part(s) doean't have charset defined

			reencodeCurrent = FALSE;
			reencodeNew = FALSE;
		}

		int bodyLength = body->m_contentLength;
		if ((body->m_contentOffset + body->m_contentLength) > m->m_length)
		{
			// something is not consistent
			bodyLength = m->m_length - body->m_contentOffset;
			if (bodyLength < 0)
				bodyLength = 0;
		}
		
		fileOffset = m->m_startOff + body->m_contentOffset;
		_int64 filePos = fpm.Seek(fileOffset, SEEK_SET);

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
			//
			if (!reencodeCurrent && !reencodeNew)
			{
				int needLength = dlength + outbuf->Count() + 512;
				outbuf->Resize(needLength);

				AppendInlineAttachmentNameSeparatorLine(body, bodyCnt, outbuf, textType);
				if ((outbuf->Capacity() - outbuf->Count()) < dlength)
				{
					_ASSERTE(1);
					int needMoreBytes = dlength - (outbuf->Capacity() - outbuf->Count());
					needLength = outbuf->Capacity() + needMoreBytes;
					outbuf->Resize(needLength);
				}
				char* outptr = outbuf->Data(outbuf->Count());

				int retlen = d64.GetOutput((unsigned char*)outptr, dlength);
				if (retlen > 0)
					outbuf->SetCount(outbuf->Count() + retlen);

				int deb2 = 1;
			}
			else
			{
				// Need to reencode to UTF-8 and append to output buffer
				int needLength = dlength;
				tmpbuf->ClearAndResize(needLength);

				int retlen = d64.GetOutput((unsigned char*)tmpbuf->Data(), dlength);
				if (retlen > 0)
				{
					tmpbuf->SetCount(retlen);
				}
			}
		}
		else if (body->m_contentTransferEncoding.CompareNoCase("quoted-printable") == 0)
		{
			MboxCMimeCodeQP dGP(bodyBegin, bodyLength);
			int dlength = dGP.GetOutputLength();
			//
			if (!reencodeCurrent && !reencodeNew)
			{
				int needLength = dlength + outbuf->Count() + 512;
				outbuf->Resize(needLength);

				AppendInlineAttachmentNameSeparatorLine(body, bodyCnt, outbuf, textType);
				if ((outbuf->Capacity() - outbuf->Count()) < dlength)
				{
					_ASSERTE(1);
					int needMoreBytes = dlength - (outbuf->Capacity() - outbuf->Count());
					needLength = outbuf->Capacity() + needMoreBytes;
					outbuf->Resize(needLength);
				}
				char* outptr = outbuf->Data(outbuf->Count());

				int retlen = dGP.GetOutput((unsigned char*)outptr, dlength);
				if (retlen > 0)
				{
					outbuf->SetCount(outbuf->Count() + retlen);
				}
			}
			else
			{
				int needLength = dlength;
				tmpbuf->ClearAndResize(needLength);

				int retlen = dGP.GetOutput((unsigned char*)tmpbuf->Data(), dlength);
				if (retlen > 0)
				{
					tmpbuf->SetCount(retlen);
				}
			}
		}
		else
		{
			// in case we have multiple bodies of the same type ?? not sure it is valid case/concern
			// asking for trouble ??
			_ASSERTE((body->m_contentTransferEncoding.CompareNoCase("base64") != 0) &&
				(body->m_contentTransferEncoding.CompareNoCase("quoted-printable") != 0));
			if (!reencodeCurrent && !reencodeNew)
			{
				AppendInlineAttachmentNameSeparatorLine(body, bodyCnt, outbuf, textType);
				outbuf->Append(bodyBegin, bodyLength);
			}
			else
			{
				tmpbuf->Append(bodyBegin, bodyLength);
			}
		}

		// reencodeCurrent and reencodeNew are always set to FALSE for HTML text block
		if ((reencodeCurrent || reencodeNew) && (textType == 1))
			_ASSERTE(FALSE);

		if (reencodeCurrent || reencodeNew)
		{
			SimpleString *result;
			if (reencodeCurrent)
			{
				SimpleString *wBuff = MboxMail::m_largelocal1;
				wBuff->ClearAndResize(4*outbuf->Count());

				result = MboxMail::m_largelocal2;
				result->ClearAndResize(4*outbuf->Count());

				BOOL retResult = TextUtilsEx::Str2UTF8(outbuf, bdyPageCode, result, wBuff, error);
				if (retResult == TRUE)
				{
					outbuf->Clear();
					outbuf->Append(result->Data(), result->Count());
					pageCode = CP_UTF8;
				}
				else
					; // don't touch outbuf
			}
			if (reencodeNew)
			{
				SimpleString *wBuff = MboxMail::m_largelocal1;
				wBuff->ClearAndResize(4*tmpbuf->Count());

				result = MboxMail::m_largelocal2;
				result->ClearAndResize(4*tmpbuf->Count());

				// TODO: make wBuff local to Str2UTF8
				BOOL retResult = TextUtilsEx::Str2UTF8(tmpbuf, pageCode, result, wBuff, error);

				AppendInlineAttachmentNameSeparatorLine(body, bodyCnt, outbuf, textType);

				if (retResult == TRUE)
				{
					outbuf->Append(result->Data(), result->Count());
					pageCode = CP_UTF8;
				}
				else
				{
					outbuf->Append(tmpbuf->Data(), tmpbuf->Count());
				}
			}
			//charset = "utf-8";
		}
		else
		{
			; // all done
		}

		if (setAsUTF8)
		{
			pageCode = CP_UTF8;
			bdyPageCode = CP_UTF8;
		}
		else if ((bdyPageCode == CP_UTF8) || (pageCode == CP_UTF8))
		{
			pageCode = CP_UTF8;
			bdyPageCode = CP_UTF8;
			if (bodyCnt > 0)
				int deb = 1;
		}
	
		bodyCnt++;

		if (bodyCnt > 1)
			int deb = 1;
	}

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
	_ASSERTE(outbuf != inbuf);
	//_ASSERTE(outbuf != workbuf);

	outbuf->ClearAndResize(10000);
	inbuf->ClearAndResize(10000);
	//workbuf->ClearAndResize(10000);

	MailBodyContent *body;
	char *data;
	int dataLen;

	FileUtils::RemoveDir(CMainFrame::GetMboxviewTempPath());

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
		_int64 filePos = fpm.Seek(fileOffset, SEEK_SET);

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

		const char *fileNameA = (LPCSTR)body->m_attachmentName;
		CString fileName = (LPCSTR)body->m_attachmentName;;
		CFile fp;
		CFileException ExError;
		if (fp.Open(fileName, CFile::modeWrite | CFile::modeCreate, &ExError))
		{
			fp.Write(data, dataLen);
			fp.Close();
		}
		else
		{
			DWORD lastErr = ::GetLastError();
#if 1
			//HWND h = GetSafeHwnd();
			HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
			CString fmt = L"Could not create file:\n\n\"%s\"\n\n%s";  // new format
			CString attachmentName;
			body->AttachmentName2WChar(attachmentName);
			CString errorText = FileUtils::ProcessCFileFailure(fmt, attachmentName, ExError, lastErr, h);
#else
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt = L"Could not create \"" + body->m_attachmentName;
			txt += L"\" file.\n";
			txt += exErrorStr;

			TRACE(L"%s\n", txt);
#endif
		}
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
					CStringA SubType = pB->GetSubType().c_str();
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
int MboxMail::exportToCSVFileFullMailParse(CSVFILE_CONFIG &csvConfig)  // FIXMEFIXME
{
#if 0
	MboxMail *ml;
	MboxMail ml_tmp;
	MboxMail *m = &ml_tmp;
	CString colLabels;
	CString separator;
	bool separatorNeeded;

	CString mailFile = MboxMail::s_path;

	if (s_path.IsEmpty()) {
		CString txt = L"Please open mail file first.";
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	CString path = MboxMail::GetLastPath();
	CString datapath = MboxMail::GetLastDataPath();
	if (path.IsEmpty())
		return -1;  // Hopefully s_path wil fail first

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(mailFile, driveName, directory, fileNameBase, fileNameExtention);

	CFile fp;
	CString csvFile = datapath + fileNameBase + ".csv";
	CFileException ExError;
	if (!fp.Open(csvFile, CFile::modeWrite | CFile::modeCreate, &ExError))
	{
		DWORD lastErr = ::GetLastError();

		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not create \"" + csvFile;
		txt += L"\" file.\n";
		txt += exErrorStr;

		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CFile fpm;
	CFileException ExError2;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError2))
	{
		DWORD lastErr = ::GetLastError();

		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError2);

		CString txt = L"Could not open \"" + MboxMail::s_path;
		txt += L"\" mail file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);

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
		_int64 filePos = fpm.Seek(fileOffset, SEEK_SET);

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

			int datelen = istrlen(datebuff);
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
					tokenlen = IntPtr2Int(last - token);
					name.ClearAndResize(tokenlen);
					addr.ClearAndResize(tokenlen);
					splitMailAddress(token, tokenlen, &name, &addr);
					atpos = 0;
				}
				else
				{
					*seppos = 0;
					tokenlen = IntPtr2Int(seppos - token);

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

	CString infotxt = L"Created \"" + csvFile + L"\" file.";
	HWND h = NULL; // we don't have any window yet
	int answer = ::MessageBox(h, infotxt, L"Success", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
#endif
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
		// Should _ASSERTE here, there should one at least
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
					UINT page_code = TextUtilsEx::StrPageCodeName2PageCode(charset.c_str());
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
	_int64 filePos = fpm.Seek(fileOffset, SEEK_SET);

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

void MboxMail::ReleaseResources(BOOL updateRegistry)
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
	delete MboxMail::m_largelocal1;
	MboxMail::m_largelocal1 = 0;
	delete MboxMail::m_largelocal2;
	MboxMail::m_largelocal2 = 0;
	delete MboxMail::m_largelocal3;
	MboxMail::m_largelocal3 = 0;
	delete MboxMail::m_smalllocal1;
	MboxMail::m_smalllocal1 = 0;
	delete MboxMail::m_smalllocal2;
	MboxMail::m_smalllocal2 = 0;
	//
	delete MailBody::m_mpool;
	MailBody::m_mpool = 0;
	delete m_pMessageIdTable;
	m_pMessageIdTable = 0;
	delete m_pMboxMailMap;
	m_pMboxMailMap = 0;
	delete m_pMboxMailTable;
	m_pMboxMailTable = 0;
	delete m_pThreadIdTable;
	m_pThreadIdTable = 0;

	MyStackWalker *sw = MboxMail::glStackWalker;
	MboxMail::glStackWalker = 0;
#ifdef USE_STACK_WALKER
	delete sw;
#endif

	CString section_general = CString(sz_Software_mboxview) + L"\\General";
	if (updateRegistry)
	{
		MboxMail::m_HintConfig.SaveToRegistry();

		CString processpath = "";
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_general, L"processPath", processpath);
	}
}

void ShellExecuteError2Text(UINT_PTR errorCode, CString &errorText) {
	errorText.Format(L"Error code: %u", errorCode);
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
#if 0
	MboxMail *m;

	CString mailFile = MboxMail::s_path;

	if (s_path.IsEmpty()) {
		CString txt = L"Please open mail file first.";
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	CString path = MboxMail::GetLastPath();
	CString datapath = MboxMail::GetLastDataPath();
	if (path.IsEmpty())
		return -1;  // Hopefully s_path wil fail first

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(mailFile, driveName, directory, fileNameBase, fileNameExtention);

	CFile fp;
	CString statsFile = datapath + fileNameBase + "_stats.txt";

	CFileException ExError;
	if (!fp.Open(statsFile, CFile::modeWrite | CFile::modeCreate, &ExError))
	{
		DWORD lastErr = ::GetLastError();

		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not create \"" + statsFile;
		txt += L"\" file.\n";
		txt += exErrorStr;

		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CFile fpm;
	CFileException ExError2;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError2))
	{
		DWORD lastErr = ::GetLastError();

		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError2);

		CString txt = L"Could not open \"" + MboxMail::s_path;
		txt += L"\" mail file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);

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

			if (TextUtilsEx::findNoCase(outbuf->Data(), outbuf->Count(), "charset=", 8) >= 0)
				hasCharset = TRUE;
		}

		text.Empty();
		text.Format("INDX=%d PLAIN=%d HTML=%d BODY=%d CHARSET=%d\n",
			i, hasPlain, hasHtml, hasBody, hasCharset);

		fp.Write(text, text.GetLength());
	}
	fp.Close();
#endif
	return 1;
}

int MboxMail::DumpMailSummaryToFile(MailArray* mailsArray, int mailsArrayCount)
{
	MboxMail* m;

	CString mailFile = MboxMail::s_path;

	if (s_path.IsEmpty()) {
		CString txt = L"Please open mail file first.";
		ResHelper::TranslateString(txt);
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	CString path = MboxMail::GetLastPath();
	CString datapath = MboxMail::GetLastDataPath();
	if (path.IsEmpty())
		return -1;  // Hopefully s_path wil fail first

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(mailFile, driveName, directory, fileNameBase, fileNameExtention);

	CFile fp;
	CString statsFile = datapath + fileNameBase + "_summary.txt";

	CFileException ExError;
	if (!fp.Open(statsFile, CFile::modeWrite | CFile::modeCreate, &ExError))
	{
		DWORD lastErr = ::GetLastError();
#if 1
		//HWND h = GetSafeHwnd();
		HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		CString fmt = L"Could not create file:\n\n\"%s\"\n\n%s";  // new format
		CString errorText = FileUtils::ProcessCFileFailure(fmt, statsFile, ExError, lastErr, h);
#else
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt;
		CString fmt = L"Could not create \"%s\" file.\n%s";
		txt.Format(fmt, statsFile, exErrorStr);

		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
#endif
		return -1;
	}

	CFile fpm;
	CFileException ExError2;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError2))
	{
		DWORD lastErr = ::GetLastError();
#if 1
		//HWND h = GetSafeHwnd();
		HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		CString fmt = L"Could not open mail file:\n\n\"%s\"\n\n%s";  // new format
		CString errorText = FileUtils::ProcessCFileFailure(fmt, MboxMail::s_path, ExError, lastErr, h); 
#else
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError2);

		CString txt = L"Could not open \"" + MboxMail::s_path;
		txt += L"\" mail file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);
#endif

		return -1;
	}

	CStringA text;
	for (int i = 0; i < mailsArray->GetSize(); i++)
	{
		m = (*mailsArray)[i];

		text.Empty();
		text.Format("%s\n", m->m_subj);
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
	m_bIsDirty = FALSE;
}

MailList::MailList()
{
	m_nId = IDC_ARCHIVE_LIST;
	//s_mails;
	b_mails_which_sorted = 0;
	m_lastSel = -1;
	m_bIsDirty = FALSE;
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
BOOL MboxMail::IsLabelMailsSelected() {
	return (nWhichMailList == IDC_LABEL_LIST);
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
int MboxMail::LabelMailsSelectedId() {
	return (IDC_LABEL_LIST);
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
	{
		if (nNewSize > this->GetUpperBound())
			SetSize(nNewSize, nGrowBy);
		else
			m_nSize = nNewSize;
	}
}

void MyMailArray::SetCountKeepData(INT_PTR nNewSize, INT_PTR nGrowBy)
{
	if (nNewSize == 0)
		m_nSize = 0;
	else
	{
		if (nNewSize > this->GetUpperBound())
			SetSize(nNewSize, nGrowBy);
		else
			m_nSize = nNewSize;
	}
}

void MyMailArray::CopyKeepData(const MyMailArray& src)
{
	if (((MyMailArray&)src).GetSize() == 0)
		m_nSize = 0;
	else
		Copy(src);
}

INT_PTR MyMailArray::MaxSize()
{
	return m_nMaxSize;
}

int MboxMail::MakeFileName(MboxMail *m, struct NamePatternParams *namePatternParams, CString &fileName, CStringA& fileNameA)
{
	SimpleString name;
	SimpleString addr;
	SimpleString addrFrom;
	SimpleString addrTo;
	CStringA tempAddr;
	CStringA fromAddr;
	CStringA toAddr;
	CStringA subjAddr;
	CStringA subj;
	CString tmp;
	int pos;
	BOOL allowUnderscore = FALSE;
	BOOL separatorNeeded = FALSE;
	wchar_t sepchar = '-';
	DWORD error;
	SimpleString workBuff;

	if (m->m_timeDate < 0)
	{
		CString m_strDate;
		if (namePatternParams->m_bDate)
		{
			if (namePatternParams->m_bTime)
				//m_strDate = ltt.Format("%Y%m%d-%H%M%S");
				m_strDate = L"000000000-000000";
			else
				//m_strDate = ltt.Format("%H%M%S");
				m_strDate = L"000000";

			fileName.Append(m_strDate);
		}
		else if (namePatternParams->m_bTime)
		{
			m_strDate = L"000000";
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
			format = L"%Y%m%d-%H%M%S";
		}
		else
		{
			format = L"%Y%m%d";
		}

		m_strDate = tt.FormatLocalTm(format);
		fileName.Append(m_strDate);

		separatorNeeded = TRUE;
	}
	else if (namePatternParams->m_bTime) 
	{
		MyCTime tt(m->m_timeDate);
		CString format = L"%H%M%S";
		CString m_strDate;
		m_strDate = tt.FormatLocalTm(format);

		CString strDate = m_strDate;
		fileName.Append(strDate);

		separatorNeeded = TRUE;
	}

	CString crc32;
	//crc32.Format("%X", m->m_crc32);
	crc32.Format(L"%07d", m->m_index);

	int fileLength = fileName.GetLength();
	if ((fileLength + crc32.GetLength()) > namePatternParams->m_nFileNameFormatSizeLimit)
	{
		int deb = 1;
	}

	if (separatorNeeded) {
		fileName.AppendChar(sepchar);
		int deb = 1;
	}

	fileName.Append((LPCWSTR)crc32, crc32.GetLength());

	if ((fileName.GetLength() < namePatternParams->m_nFileNameFormatSizeLimit) && namePatternParams->m_bFrom)
	{
		if (separatorNeeded) {
			fileName.AppendChar(sepchar);
			int deb = 1;
		}

		int fromlen = m->m_from.GetLength();
		name.ClearAndResize(fromlen);
		addr.ClearAndResize(fromlen);
		addrFrom.ClearAndResize(fromlen);

		const char* data = (LPCSTR)m->m_from;
		int datalen = fromlen;;

		if (m->m_from_charsetId != CP_UTF8)
		{
			DWORD error;
			BOOL retStr2CP = TextUtilsEx::Str2CodePage((LPCSTR)m->m_from, fromlen, m->m_from_charsetId, CP_UTF8,
				&addrFrom, &workBuff, error);

			const char* data = addrFrom.Data();
			int datalen = addrFrom.Count();
		}

		MboxMail::splitMailAddress(data, datalen, &name, &addr);
		fromAddr.Empty();
		int pos = addr.Find(0, '@');
		if (pos >= 0)
			fromAddr.Append(addr.Data(), pos);
		else
			fromAddr.Append(addr.Data(), addr.Count());
		fromAddr.Trim(" \t\"<>:;,");  // may need to remove more not aphanumerics

		CString fromAddrW;
		int retUTF82Wstr = TextUtilsEx::UTF82WStr(&fromAddr, &fromAddrW, error);
		fileName.Append(fromAddrW);

		separatorNeeded = TRUE;
	}
	if ((fileName.GetLength() < namePatternParams->m_nFileNameFormatSizeLimit) &&  namePatternParams->m_bTo)
	{
		if (separatorNeeded) {
			fileName.AppendChar(sepchar);
			int deb = 1;
		}

		CStringA* pToAddr = &m->m_to;

		int tolen = m->m_to.GetLength();

		if (m->m_from_charsetId != CP_UTF8)
		{
			tempAddr.Preallocate(tolen*2);

			DWORD error;
			BOOL retStr2Utf8 = TextUtilsEx::Str2UTF8((LPCSTR)m->m_to, tolen, m->m_to_charsetId, tempAddr, error);

			pToAddr = &tempAddr;
		}

		int maxNumbOfAddr = 1;
		NListView::TrimToAddr(pToAddr, toAddr, maxNumbOfAddr);

		pos = toAddr.Find("@");

		addrTo.Clear();
		if (pos >= 0)
			addrTo.Append(toAddr, pos);
		else
			addrTo.Append(toAddr, toAddr.GetLength());

		toAddr.Empty();
		toAddr.Append(addrTo.Data(), addrTo.Count());

		toAddr.Trim(" \t\"<>:;,"); // may need to remove more not aphanumerics

		CString toAddrW;
		int retUTF82Wstr = TextUtilsEx::UTF82WStr(&toAddr, &toAddrW, error);
		fileName.Append(toAddrW);

		separatorNeeded = TRUE;
	}
	
	int subjLengthAllowed = namePatternParams->m_nFileNameFormatSizeLimit - fileName.GetLength();

#if 0
	CString crc32;
	//crc32.Format("%X", m->m_crc32);
	crc32.Format("%07d", m->m_index);

	subjLengthAllowed = subjLengthAllowed - (crc32.GetLength() + 3 + 6);  // 3 is for 2 * dash plus at least one character for subject plus 6 for file extension
#endif
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

		UINT strCodePage = m->m_subj_charsetId;
		DWORD error;
		CString subjW;
		BOOL retC2W = TextUtilsEx::CodePage2WStr(&m->m_subj, strCodePage, &subjW, error);
		fileName.Append(subjW);
#if 0
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
					fileName.AppendChar(L'_');
					allowUnderscore = FALSE;
					outCnt++;
				//}
				//else break;
			}
			else
				ignoreCnt++;
		}
#endif
		separatorNeeded = TRUE;
	}


	fileName.Trim(L" \t_");
#if 0
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
#endif

	// Convert to Ascii otherwise pdf merge cmd script may not work
	if (namePatternParams->m_bConvert2Ansi)
	{
		TextUtilsEx::WStr2Ascii(fileName, fileNameA, error);
		TextUtilsEx::Ascii2Wstr(fileNameA, fileName, error);
	}

	BOOL bReplaceWhiteWithUnderscore = FALSE;

	CString result;
	BOOL extraValidation = TRUE;
	FileUtils::MakeValidFileName(fileName, result, bReplaceWhiteWithUnderscore, extraValidation);
	fileName = result;

	fileName.Trim(L" \t_");

	int fileNameLength = fileName.GetLength();
	if (fileNameLength > namePatternParams->m_nFileNameFormatSizeLimit)
	{
		fileName.Truncate(namePatternParams->m_nFileNameFormatSizeLimit);
		fileName.Trim(L" \t_");
	}
	return 1;
}

int MboxMail::MakeFileName(MboxMail *m, struct NamePatternParams* namePatternParams, struct NameTemplateCnf &nameTemplateCnf, CString &fileName, int maxFileNameLength, CStringA& fileNameA)
{
	SimpleString name;
	SimpleString addr;
	CStringA cname;
	CStringA caddr;

	CStringA username;
	CStringA domain;

	CStringA fromAddr;
	CStringA fromName;
	
	CStringA toAddr;
	CStringA toName;

	CStringA tempAddr;

	SimpleString workBuff;
	SimpleString addrFrom;
	SimpleString addrTo;
	SimpleString nameTo;

	CStringA addrStr;
	CStringA nameStr;

	CString subjAddr;
	CString subj;
	CString tmp;
	CString FieldText;
	BOOL allowUnderscore = FALSE;
	char sepchar = '-';
	int needLength;
	DWORD error;

	// make room for extension - actually caller should do this
	maxFileNameLength -= 6; 
	int maxNonAdjustedFileNameLength = maxFileNameLength;

	CArray<CString> labelArray;
	BOOL ret = MboxMail::ParseTemplateFormat(nameTemplateCnf.m_TemplateFormat, labelArray);

	// make room for UNIQUE_ID
	CString idLabel = L"%UNIQUE_ID%";
	CString uID;
	uID.Format(L"%07d", m->m_index);
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
		if (labelArray[i].CompareNoCase(L"%DATE_TIME%") == 0)
		{
			CString strDate;
			if (m->m_timeDate < 0)
			{
				strDate = L"0000-00-00";
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
		else if (labelArray[i].CompareNoCase(L"%FROM_NAME%") == 0)
		{
			CString addrLabel = L"%FROM_ADDR%";

			const int fromlen = m->m_from.GetLength();
			name.ClearAndResize(fromlen);
			addr.ClearAndResize(fromlen);
			username.Empty();
			domain.Empty();
			cname.Empty();
			fromAddr.Empty();
			fromName.Empty();

			addrFrom.ClearAndResize(fromlen);

			const char* data = (LPCSTR)m->m_from;
			int datalen = fromlen;;

			if (m->m_from_charsetId != CP_UTF8)
			{
				DWORD error;
				const BOOL retStr2CP = TextUtilsEx::Str2CodePage((LPCSTR)m->m_from, fromlen, m->m_from_charsetId, CP_UTF8,
					&addrFrom, &workBuff, error);

				data = addrFrom.Data();
				datalen = addrFrom.Count();
			}

			MboxMail::splitMailAddress(data, datalen, &name, &addr);

			// Call MakeValidFileNameA to make sure validated name/cname has at least one valid character other than '_'
			// FIXME Calling MakeValidFileNameA might be just overhead now since MakeValidFileNameA doesn't convert to Ascii anymore
			BOOL extraValidation = TRUE;
			FileUtils::MakeValidFileNameA(name.Data(), name.Count(), cname, nameTemplateCnf.m_bReplaceWhiteWithUnderscore, extraValidation);
			if ((cname.GetLength() > 0) && !((cname.GetLength() == 1) && (cname.GetAt(0) == '_')))
			{
				fromName.Append((LPCSTR)cname, cname.GetLength());
				fromName.Trim(" \t\"<>:;,()");  // may need to remove more not only aphanumerics
				needLength = fileName.GetLength() + fromName.GetLength();

				if (needLength < maxFileNameLength)
				{
					CString fromNameW;
					const int retUTF82Wstr = TextUtilsEx::UTF82WStr(&fromName, &fromNameW, error);
					fileName.Append(fromNameW);
				}
			}
			else if (TemplateFormatHasLabel(addrLabel, labelArray) == FALSE)  
			{
				// Ok, getting Name didn't work, try Address instead
				int pos = addr.Find(0, '@');
				if (pos >= 0)
				{
					username.Append(addr.Data(), pos);
					domain.Append(addr.Data(pos + 1), addr.Count() - (pos + 1));
				}
				else
					username.Append(addr.Data(), addr.Count());

				if (nameTemplateCnf.m_bFromUsername && nameTemplateCnf.m_bFromDomain)
					fromAddr.Append(addr.Data(), addr.Count());
				else if (nameTemplateCnf.m_bFromUsername)
					fromAddr.Append(username);
				else
					fromAddr.Append(domain);

				fromAddr.Trim(" \t\"<>:;,()");  // may need to remove more not only aphanumerics
				needLength = fileName.GetLength() + fromAddr.GetLength();

				if (needLength < maxFileNameLength)
				{
					CString fromAddrW;
					const int retUTF82Wstr = TextUtilsEx::UTF82WStr(&fromAddr, &fromAddrW, error);
					fileName.Append(fromAddrW);
				}
			}
		}
		else if (labelArray[i].CompareNoCase(L"%FROM_ADDR%") == 0)
		{
			const int fromlen = m->m_from.GetLength();
			name.ClearAndResize(fromlen);
			addr.ClearAndResize(fromlen);
			username.Empty();
			domain.Empty();
			fromAddr.Empty();

			addrFrom.ClearAndResize(fromlen);

			const char* data = (LPCSTR)m->m_from;
			int datalen = fromlen;;

			if (m->m_from_charsetId != CP_UTF8)
			{
				DWORD error;
				BOOL retStr2CP = TextUtilsEx::Str2CodePage((LPCSTR)m->m_from, fromlen, m->m_from_charsetId, CP_UTF8,
					&addrFrom, &workBuff, error);

				data = addrFrom.Data();
				datalen = addrFrom.Count();
			}

			MboxMail::splitMailAddress(data, datalen, &name, &addr);

			int pos = addr.Find(0, '@');
			if (pos >= 0)
			{
				username.Append(addr.Data(), pos);
				domain.Append(addr.Data(pos+1), addr.Count() - (pos+1));
			}
			else
				username.Append(addr.Data(), addr.Count());

			if (nameTemplateCnf.m_bFromUsername && nameTemplateCnf.m_bFromDomain)
				fromAddr.Append(addr.Data(), addr.Count());
			else if (nameTemplateCnf.m_bFromUsername)
				fromAddr.Append(username);
			else
				fromAddr.Append(domain);

			fromAddr.Trim(" \t\"<>:;,");  // may need to remove more not only aphanumerics
			needLength = fileName.GetLength() + fromAddr.GetLength();

			if (needLength < maxFileNameLength)
			{
				CString fromAddrW;
				const int retUTF82Wstr = TextUtilsEx::UTF82WStr(&fromAddr, &fromAddrW, error);
				fileName.Append(fromAddrW);
			}

			int deb = 1;
		}

		//
		else if (labelArray[i].CompareNoCase(L"%TO_NAME%") == 0)
		{
			CString addrLabel = L"%TO_ADDR%";

			int tolen = m->m_to.GetLength();
			name.ClearAndResize(tolen);
			addr.ClearAndResize(tolen);
			username.Empty();
			domain.Empty();
			cname.Empty();
			toName.Empty();
			toAddr.Empty();

			int maxNumbOfAddr = 1;
			NListView::TrimToAddr(&m->m_to, addrStr, maxNumbOfAddr);
			NListView::TrimToName(&m->m_to, nameStr, maxNumbOfAddr);

			name.Append(nameStr, nameStr.GetLength());
			addr.Append(addrStr, addrStr.GetLength());

			// Call MakeValidFileNameA to make sure validated name/cname has at least one valid character other than '_'
			// FIXME Calling MakeValidFileNameA might be just overhead now since MakeValidFileNameA doesn't convert to Ascii anymore

			BOOL extraValidation = TRUE;
			FileUtils::MakeValidFileNameA(name.Data(), name.Count(), cname, nameTemplateCnf.m_bReplaceWhiteWithUnderscore, extraValidation);
			if ((cname.GetLength() > 0) && !((cname.GetLength() == 1) && (cname.GetAt(0) == '_')))
			{
				toName.Append((LPCSTR)cname, cname.GetLength());
				toName.Trim(" \t\"<>:;,()");  // may need to remove more not only aphanumerics
				needLength = fileName.GetLength() + toName.GetLength();

				if (needLength < maxFileNameLength)
				{
					CString toNameW;
					const int retUTF82Wstr = TextUtilsEx::UTF82WStr(&toName, &toNameW, error);
					fileName.Append(toNameW);
				}
			}
			else if (TemplateFormatHasLabel(addrLabel, labelArray) == FALSE)
			{
				// Ok, getting Name didn't work, try Address instead
				int pos = addr.Find(0, '@');
				if (pos >= 0)
				{
					username.Append(addr.Data(), pos);
					domain.Append(addr.Data(pos + 1), addr.Count() - (pos + 1));
				}
				else
					username.Append(addr.Data(), addr.Count());

				if (nameTemplateCnf.m_bToUsername && nameTemplateCnf.m_bToDomain)
					toAddr.Append(addr.Data(), addr.Count());
				else if (nameTemplateCnf.m_bToUsername)
					toAddr.Append(username);
				else
					toAddr.Append(domain);

				toAddr.Trim(" \t\"<>:;,()");  // may need to remove more not only aphanumerics
				needLength = fileName.GetLength() + toAddr.GetLength();

				if (needLength < maxFileNameLength)
				{
					CString toAddrW;
					int retUTF82Wstr = TextUtilsEx::UTF82WStr(&toAddr, &toAddrW, error);
					fileName.Append(toAddrW);
				}
			}
		}
		else if (labelArray[i].CompareNoCase(L"%TO_ADDR%") == 0)
		{
			int tolen = m->m_to.GetLength();
			name.ClearAndResize(tolen);
			addr.ClearAndResize(tolen);
			username.Empty();
			domain.Empty();
			toAddr.Empty();

			int maxNumbOfAddr = 1;
			NListView::TrimToAddr(&m->m_to, addrStr, maxNumbOfAddr);
			NListView::TrimToName(&m->m_to, nameStr, maxNumbOfAddr);

			name.Append(nameStr, nameStr.GetLength());
			addr.Append(addrStr, addrStr.GetLength());

			int pos = addr.Find(0, '@');
			if (pos >= 0)
			{
				username.Append(addr.Data(), pos);
				domain.Append(addr.Data(pos + 1), addr.Count() - (pos + 1));
			}
			else
				username.Append(addr.Data(), addr.Count());

			if (nameTemplateCnf.m_bToUsername && nameTemplateCnf.m_bToDomain)
				toAddr.Append(addr.Data(), addr.Count());
			else if (nameTemplateCnf.m_bToUsername)
				toAddr.Append(username);
			else
				toAddr.Append(domain);

			toAddr.Trim(" \t\"<>:;,");  // may need to remove more not only aphanumerics
			needLength = fileName.GetLength() + toAddr.GetLength();

			if (needLength < maxFileNameLength)
			{
				CString toAddrW;
				const int retUTF82Wstr = TextUtilsEx::UTF82WStr(&toAddr, &toAddrW, error);
				fileName.Append(toAddrW);
			}

			int deb = 1;
		}
		else if (labelArray[i].CompareNoCase(L"%SUBJECT%") == 0)
		{
			int lengthAvailable = maxFileNameLength - fileName.GetLength();
			if (lengthAvailable < 0)
				lengthAvailable = 0;
			int subjectLength = m->m_subj.GetLength();
			if (subjectLength > lengthAvailable)
				subjectLength = lengthAvailable;

			UINT strCodePage = m->m_subj_charsetId;
			CString subjW;
			BOOL retC2W = TextUtilsEx::CodePage2WStr(&m->m_subj, strCodePage, &subjW, error);
			fileName.Append(subjW);
		}
		else if (labelArray[i].CompareNoCase(L"%UNIQUE_ID%") == 0)
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
	fileName.Trim(L" \t_");


	// Convert to Ascii otherwise pdf merge cmd script may not work
	if (namePatternParams->m_bConvert2Ansi)
	{
		TextUtilsEx::WStr2Ascii(fileName, fileNameA, error);
		TextUtilsEx::Ascii2Wstr(fileNameA, fileName, error);
	}

	BOOL extraValidation = TRUE;
	FileUtils::MakeValidFileName(fileName, nameTemplateCnf.m_bReplaceWhiteWithUnderscore, extraValidation);

	fileName.Trim(L" \t_");

	const int fileNameLength = fileName.GetLength();
	if (fileNameLength > maxFileNameLength)
	{
		fileName.Truncate(maxFileNameLength);
	}
	return 1;
}

static CArray<CString> FldLabelTable;
static BOOL fldLabelsTableSet = FALSE;

CString MboxMail::MatchFldLabel(LPCWSTR p, LPCWSTR e)
{
	CString empty;
	if (fldLabelsTableSet == FALSE)
	{
		FldLabelTable.Add(L"%DATE_TIME%");
		FldLabelTable.Add(L"%FROM_NAME%");
		FldLabelTable.Add(L"%FROM_ADDR%");
		FldLabelTable.Add(L"%TO_NAME%");
		FldLabelTable.Add(L"%TO_ADDR%");
		FldLabelTable.Add(L"%SUBJECT%");
		FldLabelTable.Add(L"%UNIQUE_ID%");

		fldLabelsTableSet = TRUE;
	}

	int left = IntPtr2Int(e - p);
	int fldCnt = (int)FldLabelTable.GetCount();
	int labelLength;
	LPCWSTR labelData;
	int i;
	for (i = 0; i < fldCnt; i++)
	{
		labelLength = FldLabelTable[i].GetLength();
		labelData = (LPCWSTR)FldLabelTable[i];
		if ((FldLabelTable[i].GetLength() <= left) && (_tcsncmp(p, labelData, labelLength) == 0))
			return FldLabelTable[i];
	}
	return empty;
}

BOOL MboxMail::ParseTemplateFormat(CString &templateFormat, CArray<CString> &labelArray)
{
	LPCWSTR p = (LPCWSTR)templateFormat;
	LPCWSTR e = p + templateFormat.GetLength();
	CString label;
	CString plainText;
	wchar_t c;
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
	TRACE(L"LABEL ARRAY:\n\n");
	for (int ii = 0; ii < labelArray.GetCount(); ii++)
	{
		TRACE(L"\"%s\"\n", labelArray[ii]);
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

int MboxMail::RemoveDuplicateMails(MailArray &s_mails_array, BOOL putDuplicatesOnFindArray)
{
	MboxMail *m;
	MboxMail *m_found;

	// !! s_mails is already populated from mboxview file

	ULONGLONG tc_start = GetTickCount64();

	if (s_mails.GetCount() > s_mails_array.GetCount())
	{
		s_mails_array.SetSizeKeepData(s_mails.GetCount());
	}

	if (putDuplicatesOnFindArray)
	{
		if (s_mails.GetCount() > s_mails_find.GetCount())
		{
			s_mails_find.SetSizeKeepData(s_mails.GetCount());
		}
	}

	int i = 0;
	int to_dup_i = 0;
	int to_i = 0;

	// Test exceptions
	// CTime ct(1, 1, 1, 1, 1, 1,1);
	// m = s_mails[s_mails.GetSize() + 1000];

	for (i; i < s_mails.GetSize(); i++)
	{
		m = s_mails[i];

		m_found = getMboxMail(m);
		if (m_found == 0)
		{
			insertMboxMail(m, m);

			s_mails_array[to_i] = s_mails[i];
			to_i++;
			m->m_duplicateId = false;
		}
		else
		{
			if (putDuplicatesOnFindArray)
			{
				s_mails_find[to_dup_i] = s_mails[i];
				to_dup_i++;
				m->m_duplicateId = true;
			}
			else
			{
				// delete m now or add to table to be deleted later
				MboxMail::DestroyMboxMail(m);
				to_dup_i++;
			}
		}
	}

	s_mails_array.SetSizeKeepData(to_i);

	if (putDuplicatesOnFindArray)
	{
		s_mails_find.SetSizeKeepData(to_dup_i);

		MboxMail::m_findMails.m_lastSel = -1;
		MboxMail::m_findMails.b_mails_which_sorted = MboxMail::b_mails_which_sorted;
	}

	int dupCnt = s_mails.GetSize() - s_mails_array.GetSize();

	s_mails.CopyKeepData(s_mails_array);

	// moved to NListView::OnRClickSingleSelect
	//MboxMail::m_editMails.m_bIsDirty = TRUE;

	ULONGLONG tc_clear_start = GetTickCount64();

	clearMboxMailTable();
	ULONGLONG tc_clear_end = GetTickCount64();
	DWORD delta_clear = (DWORD)(tc_clear_end - tc_clear_start);

	ULONGLONG tc_end = GetTickCount64();
	DWORD delta = (DWORD)(tc_end - tc_start);

	TRACE(L"RemoveDuplicateMails: total time=%d !!!!!!!!!!!!.\n", delta);
	TRACE(L"RemoveDuplicateMails: clear time=%d !!!!!!!!!!!!.\n", delta_clear);

	return dupCnt;
}

int MboxMail::LinkDuplicateMails(MailArray &s_mails_array)
{
	MboxMail *m;
	MboxMail *m_found;

	// !! s_mails is already populated from mboxview file

	ULONGLONG tc_start = GetTickCount64();

	// Remove duplicate mails
	if (s_mails.GetCount() > s_mails_array.GetCount())
	{
		s_mails_array.SetSizeKeepData(s_mails.GetCount());
	}

	if (s_mails.GetCount() > s_mails_find.GetCount())
	{
		s_mails_find.SetSizeKeepData(s_mails.GetCount());
	}


	int i = 0;
	int to_dup_i = 0;
	int to_i = 0;

	for (i = 0; i < s_mails.GetSize(); i++)
	{
		m = s_mails[i];

		//if (m->m_subj.Find("[Customer Support - Query #537] (In Progress) RE: log issue related to AdvanceSubscriber") >= 0)
			//int deb = 1;

		m_found = getMboxMail(m);

		if (m_found == 0)
		{
			to_i++;
			m->m_duplicateId = false;
		}
		else
		{
			to_dup_i++;
			m->m_duplicateId = true;
		}
		insertMboxMail(m, m);
	}

	//s_mails_array.SetSizeKeepData(to_i);

	//s_mails_find.SetSizeKeepData(to_dup_i);

	int dupCnt = s_mails.GetSize() - s_mails_array.GetSize();

	//s_mails.CopyKeepData(s_mails_array);

	ULONGLONG tc_clear_start = GetTickCount64();

	//clearMboxMailTable();
	ULONGLONG tc_clear_end = GetTickCount64();
	DWORD delta_clear = (DWORD)(tc_clear_end - tc_clear_start);

	ULONGLONG tc_end = GetTickCount64();
	DWORD delta = (DWORD)(tc_end - tc_start);

	TRACE(L"LinkDuplicateMails: total time=%d !!!!!!!!!!!!.\n", delta);
	TRACE(L"LinkDuplicateMails: clear time=%d !!!!!!!!!!!!.\n", delta_clear);

	return dupCnt;
}

//

// Will increate MboxMail by 8 bytes by using intrusive IHashTable
// Performance will be better, no allocation and deallocation required
// For 200,000 email it will use 1.6 Mbytes plus extra overhead for collisiomn lists
//#define HASH_ARRAY_SIZE 50013
MboxMail* MboxMail::getMboxMail(MboxMail *key)
{
	MboxMailTableType::iterator it;
	if (m_pMboxMailMap == 0)
		createMboxMailTable(HASH_ARRAY_SIZE);

	MboxMail *elem = m_pMboxMailMap->find(key);
	return elem;
}

bool MboxMail::insertMboxMail(MboxMail *key, MboxMail *mbox)
{
	if (m_pMboxMailMap == 0)
		createMboxMailTable(HASH_ARRAY_SIZE);

	m_pMboxMailMap->insert(key, mbox);

	return true;
}

UINT MboxMail::createMboxMailTable(UINT count)
{
	m_pMboxMailMap = new MboxMailMapType(count);
	return count;
}

void MboxMail::clearMboxMailTable()
{
	if (m_pMboxMailMap)
		m_pMboxMailMap->clear();

	int deb = 1;
}

dllist<MboxMail, &MboxMail::m_hashMapLink>* MboxMail::getCollisionList(MboxMail *key)
{
	if (m_pMboxMailMap == 0)
		createMboxMailTable(HASH_ARRAY_SIZE);

	dllist<MboxMail, &MboxMail::m_hashMapLink> *list = m_pMboxMailMap->get_collision_list(key);
	return list;
}

dllist<MboxMail, &MboxMail::m_hashMapLink>* MboxMail::getCollisionList(hashsum_t hashsum)
{
	if (m_pMboxMailMap == 0)
		createMboxMailTable(HASH_ARRAY_SIZE);

	dllist<MboxMail, &MboxMail::m_hashMapLink> *list = m_pMboxMailMap->get_collision_list(hashsum);
	return list;
}

bool MboxMail::equal(MboxMail *elem1, MboxMail *elem2)
{
	if (m_pMboxMailMap == 0)
		createMboxMailTable(HASH_ARRAY_SIZE);

	return m_pMboxMailMap->isequal(elem1, elem2);
}

BOOL MboxMail::CreateCachePath(CString &rootPrintSubFolder, CString &targetPrintSubFolder, CString &prtCachePath, CString &errorText)
{
	BOOL ret = MboxMail::GetCachePath(rootPrintSubFolder, targetPrintSubFolder, prtCachePath, errorText);
	if (ret == FALSE)
		return FALSE;

	BOOL createDirOk = TRUE;
	if (!FileUtils::PathDirExists(prtCachePath)) {
		createDirOk = FileUtils::CreateDir(prtCachePath);
	}

	if (!createDirOk)
	{
		errorText = L"Could not create \n\n\"" + prtCachePath;
		errorText += L"\"\n\n folder for print destination.";

		int prtCachePathLen = prtCachePath.GetLength();
		if (prtCachePathLen > MAX_PATH)
		{
			CString lstr;
			lstr.Format(L"\n\nFolder path of %d longer than MAX_PATH (260).", prtCachePathLen);
			errorText += lstr;
		}

		errorText += L"\n\nResolve the problem and try again.";
		return FALSE;
	}

	return TRUE;
}

BOOL MboxMail::GetCachePath(CString &rootPrintSubFolder, CString &targetPrintSubFolder, CString &prtCachePath, CString &errorText, CString *mboxFilePath)
{
	prtCachePath.Empty();

	CString mailArchiveFilePath = MboxMail::s_path;
	if (mboxFilePath)
		mailArchiveFilePath = *mboxFilePath;

	if (mailArchiveFilePath.IsEmpty()) {
		errorText = L"Please open mail archive file first.";
		return FALSE;
	}

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;
	FileUtils::SplitFilePath(mailArchiveFilePath, driveName, directory, fileNameBase, fileNameExtention);

	CString printCachePath = MboxMail::GetLastDataPath();
	//printCachePath.Append(L"\\");
	//printCachePath.Append(rootPrintSubFolder);

	//printCachePath.Append(L"\\");
	fileNameBase.Replace(L'`', L'_');
	fileNameBase.Replace(L'\'', L'_');
	fileNameBase.Replace(L'!', L'_');
	fileNameBase.Replace(L'#', L'_');   // ! and # are valid file cname characters but print to PDF seems to have problems
	printCachePath.Append(fileNameBase);
	if (!fileNameExtention.IsEmpty())
	{
		//printCachePath.Append(mailArchiveFileName);
		printCachePath.TrimRight(L"\\");
		fileNameExtention.TrimLeft(L".");
		printCachePath.Append(L"-");
		printCachePath.Append(fileNameExtention);
	}

	if (!rootPrintSubFolder.IsEmpty())
	{
		printCachePath.Append(L"\\");
		printCachePath.Append(rootPrintSubFolder);
	}

	if (!targetPrintSubFolder.IsEmpty())
	{
		printCachePath.Append(L"\\");
		printCachePath.Append(targetPrintSubFolder);
	}

	prtCachePath.Append(printCachePath);

	return TRUE;
}

bool MboxMail::GetPrintCachePath(CString &rootPrintSubFolder, CString &targetPrintSubFolder, CString &prtCachePath, CString &errorText)
{
	BOOL ret = MboxMail::GetCachePath(rootPrintSubFolder, targetPrintSubFolder, prtCachePath, errorText);
	if (ret == FALSE)
		return FALSE;
	else
		return TRUE;
}

int MboxMail::MakeFileNameFromMailArchiveName(int fileType, CString &fileName, CString &targetPrintSubFolder, bool &fileExists, CString &errorText)
{
	CString mailArchiveFileName = MboxMail::s_path;

	if (mailArchiveFileName.IsEmpty()) {
		errorText = L"Please open mail archive file first.";
		return -1;
	}

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;
	FileUtils::SplitFilePath(mailArchiveFileName, driveName, directory, fileNameBase, fileNameExtention);

	CString printCachePath;
	CString rootPrintSubFolder = "PrintCache";
	if (NListView::m_exportMailsMode)
		rootPrintSubFolder = L"ExportCache";

	BOOL retval = CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, printCachePath, errorText);
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
	CString mailArchiveFileNamePath = MboxMail::s_path;

	if (mailArchiveFileNamePath.IsEmpty()) {
		errorText = L"Please open mail archive file first.";
		return -1;
	}

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (!pFrame) {
		errorText = L"Internal error. Try again.";
		return -1;
	}

	CString printCachePath;
	CString rootPrintSubFolder = "PrintCache";
	if (NListView::m_exportMailsMode)
		rootPrintSubFolder = L"ExportCache";
	BOOL retval = CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, printCachePath, errorText);
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

	CStringA mailFileNameBaseA;
	int retMakeFileName;
	if (pFrame->m_NamePatternParams.m_bCustomFormat == FALSE)
		retMakeFileName = MboxMail::MakeFileName(m, &pFrame->m_NamePatternParams, mailFileNameBase, mailFileNameBaseA);
	else
		retMakeFileName = MboxMail::MakeFileName(m, &pFrame->m_NamePatternParams, pFrame->m_NamePatternParams.m_nameTemplateCnf, mailFileNameBase, pFrame->m_NamePatternParams.m_nFileNameFormatSizeLimit, mailFileNameBaseA);

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

int MboxMail::PrintMailRangeToSingleTextFile(TEXTFILE_CONFIG &textConfig, CString &textFileName, int firstMail, int lastMail, int textType, CString &targetPrintSubFolderName, CString &errorText)
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

	CFileException ExError;
	if (!fp.Open(textFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &ExError))
	//if (!fp.Open(textFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone))
	{
		DWORD lastErr = ::GetLastError();
#if 1
		//HWND h = GetSafeHwnd();
		HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		CString fmt = L"Could not create file:\n\n\"%s\"\n\n%s";  // new format
		errorText = FileUtils::ProcessCFileFailure(fmt, textFile, ExError, lastErr, h);  // it looks like it may  result in duplicate MessageBox ??
#else
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not create \"" + textFile;
		txt += L"\" file.\n";
		txt += exErrorStr;

		errorText = txt;
#endif
		// TODO:
		//return -1;
	}

	CFile fpm;
	CFileException ExError2;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError2))
	{
		DWORD lastErr = ::GetLastError();
#if 1
		//HWND h = GetSafeHwnd();
		HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		CString fmt = L"Could not open mail file:\n\n\"%s\"\n\n%s";  // new format
		errorText = FileUtils::ProcessCFileFailure(fmt, MboxMail::s_path, ExError2, lastErr, h);
#else
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError2);

		CString txt = L"Could not open \"" + MboxMail::s_path;
		txt += L"\" mail file.";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);

		errorText = txt;
#endif

		fp.Close();
		return -1;
	}

	if (textType == 0) {
		if (textConfig.m_nCodePageId == CP_UTF8) {
			const char *BOM_UTF8 = "\xEF\xBB\xBF";
			fp.Write(BOM_UTF8, 3);
		}
	}

	AttachmentMgr attachmentDB;
	BOOL singleMail = (firstMail == lastMail) ? TRUE : FALSE;
	for (int i = firstMail; i <= lastMail; i++)
	{
		//BOOL singleMail = (firstMail == lastMail) ? TRUE : FALSE;
		BOOL addPageBreak = MboxMail::PageBreakNeeded(i, i + 1, singleMail);

		if (textType == 0)
		{
			printSingleMailToTextFile(fp, i, fpm, textConfig, singleMail, addPageBreak);
		}
		else 
		{
			BOOL fullImgFilePath = FALSE;
			int pos = printSingleMailToHtmlFile(fp, i, fpm, textConfig, singleMail, addPageBreak, fullImgFilePath);
			if (pos < 0)
				continue;
		}
	}

	if ((textType == 1) && (singleMail == FALSE))
		printDummyMailToHtmlFile(fp);

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

	CFileException ExError;
	if (!fp.Open(textFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &ExError))
	{
		DWORD lastErr = ::GetLastError();
#if 1
		//HWND h = GetSafeHwnd();
		HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		CString fmt = L"Could not create file:\n\n\"%s\"\n\n%s";  // new format
		errorText = FileUtils::ProcessCFileFailure(fmt, textFile, ExError, lastErr, h);  // it looks like it may  result in duplicate MessageBox ??
#else
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not create file \"" + textFile;
		txt += L"\" file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);

		errorText = txt;
#endif
		return -1;
	}

	CFile fpm;
	CFileException ExError2;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError2))
	{
		DWORD lastErr = ::GetLastError();
#if 1
		//HWND h = GetSafeHwnd();
		HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		CString fmt = L"Could not open mail file:\n\n\"%s\"\n\n%s";  // new format
		errorText = FileUtils::ProcessCFileFailure(fmt, MboxMail::s_path, ExError2, lastErr, h);
#else
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not open \"" + MboxMail::s_path;
		txt += L"\" mail file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);

		errorText = txt;
#endif
		fp.Close();
		return -1;
	}

	if (progressBar && MboxMail::pCUPDUPData)
	{
		if (textType == 0)
		{
			CString txt = L"Printing mails to TEXT file ...";
			ResHelper::TranslateString(txt);
			MboxMail::pCUPDUPData->SetProgress(txt, 0);
		}
		else
		{
			CString txt = L"Printing mails to HTML file ...";
			ResHelper::TranslateString(txt);
			MboxMail::pCUPDUPData->SetProgress(txt, 0);
		}
	}

	UINT curstep = 1;
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

	AttachmentMgr attachmentDB;

	ULONGLONG workRangeFirstPos = firstMail;
	ULONGLONG workRangeLastPos = lastMail;
	ProgressTimer progressTimer(workRangeFirstPos, workRangeLastPos);

	BOOL singleMail = (firstMail == lastMail) ? TRUE : FALSE;
	for (int i = firstMail; i <= lastMail; i++)
	{
		//BOOL singleMail = (firstMail == lastMail) ? TRUE : FALSE;
		BOOL addPageBreak = MboxMail::PageBreakNeeded(i, i + 1, singleMail);

		if (textType == 0)
		{
			printSingleMailToTextFile(fp, i, fpm, textConfig, singleMail, addPageBreak);
		}
		else 
		{
			BOOL fullImgFilePath = FALSE;
			int pos = printSingleMailToHtmlFile(fp, i, fpm, textConfig, singleMail, addPageBreak, fullImgFilePath);
			if (pos < 0)
				continue;
		}

		if (progressBar && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				int deb = 1;
				break;
			}

			UINT_PTR dwProgressbarPos = 0;
			ULONGLONG workRangePos = i;
			BOOL needToUpdateStatusBar = progressTimer.UpdateWorkPos(workRangePos, dwProgressbarPos);
			if (needToUpdateStatusBar)
			{
				int nFileNum = i + 1;
				if (textType == 0)
				{
					CString fmt = L"Printing mails to single TEXT file ... %d of %d";
					ResHelper::TranslateString(fmt);
					fileNum.Format(fmt, nFileNum, cnt);
				}
				else
				{
					CString fmt = L"Printing mails to single HTML file ... %d of %d";
					ResHelper::TranslateString(fmt);
					fileNum.Format(fmt, nFileNum, cnt);
				}

				if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(dwProgressbarPos));

				int debug = 1;
			}
		}
	}

	if ((textType == 1) && (singleMail == FALSE))
		printDummyMailToHtmlFile(fp);

	UINT newstep = 100;
	if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress((UINT_PTR)(newstep));

	fp.Close();
	fpm.Close();

	return 1;
}

int MboxMail::PrintMailSelectedToSingleTextFile_WorkerThread(TEXTFILE_CONFIG &textConfig, CString &textFileName, MailIndexList *selectedMailsIndexList, int textType, CString errorText)
{
	BOOL progressBar = TRUE;
	CFile fp;
	CString textFile;
	bool fileExists = false;
	int ret = 1;

	CString targetPrintSubFolder;

	if (selectedMailsIndexList->GetCount() <= 0)
		return 1;

	//
	//  int textType = xx;  // 0=.txt 1=.htm 2=.pdf 3=.csv

	// Prone to inconsistency, fileName is also determine in ALongRightProcessProcPrintMailGroupToSingleHTML above
	if (selectedMailsIndexList->GetCount() > 1)
		ret = MboxMail::MakeFileNameFromMailArchiveName(textType, textFile, targetPrintSubFolder, fileExists, errorText);
	else
		ret = MboxMail::MakeFileNameFromMailHeader(selectedMailsIndexList->GetAt(0), textType, textFile, targetPrintSubFolder, fileExists, errorText);
	//

	if (ret < 0)
		return -1;

	textFileName = textFile;
	// Useful durimg debug
	int textFileLength = textFile.GetLength();
	int maxPath = _MAX_PATH;

	CFileException ExError;
	if (!fp.Open(textFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &ExError))
	{
		DWORD lastErr = ::GetLastError();
#if 1
		//HWND h = GetSafeHwnd();
		HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		CString fmt = L"Could not create file:\n\n\"%s\"\n\n%s";  // new format
		errorText = FileUtils::ProcessCFileFailure(fmt, textFile, ExError, lastErr, h);  // it looks like it may  result in duplicate MessageBox ??
#else
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not create \"" + textFile;
		txt += L"\" file.\n";
		txt += exErrorStr;

		errorText = txt;
#endif
		return -1;
	}

	CFile fpm;
	CFileException ExError2;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError2))
	{
		DWORD lastErr = ::GetLastError();
#if 1
		//HWND h = GetSafeHwnd();
		HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		CString fmt = L"Could not open mail file:\n\n\"%s\"\n\n%s";  // new format
		errorText = FileUtils::ProcessCFileFailure(fmt, MboxMail::s_path, ExError2, lastErr, h);
#else
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError2);

		CString txt = L"Could not open \"" + MboxMail::s_path;
		txt += L"\" mail file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);

		errorText = txt;
#endif

		fp.Close();
		return -1;
	}

	if (progressBar && MboxMail::pCUPDUPData)
	{
		if (textType == 0)
		{
			CString txt = L"Printing mails to TEXT file ...";
			ResHelper::TranslateString(txt);
			MboxMail::pCUPDUPData->SetProgress(txt, 0);
		}
		else
		{
			CString txt = L"Printing mails to HTML file ...";
			ResHelper::TranslateString(txt);
			MboxMail::pCUPDUPData->SetProgress(txt, 0);
		}
	}

	if (textType == 0) {
		if (textConfig.m_nCodePageId == CP_UTF8) {
			const char *BOM_UTF8 = "\xEF\xBB\xBF";
			fp.Write(BOM_UTF8, 3);
		}
	}

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	UINT curstep = 1;
	CString fileNum;

	if (progressBar && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	AttachmentMgr attachmentDB;
	int i;
	int cnt = (int)selectedMailsIndexList->GetCount();

	ULONGLONG workRangeFirstPos = 0;
	ULONGLONG workRangeLastPos = cnt - 1;
	ProgressTimer progressTimer(workRangeFirstPos, workRangeLastPos);

	BOOL singleMail = (cnt == 1) ? TRUE : FALSE;
	for (int j = 0; j < cnt; j++)
	{
		i = (*selectedMailsIndexList)[j];

		//BOOL singleMail = (cnt == 1) ? TRUE : FALSE;
		BOOL addPageBreak = MboxMail::PageBreakNeeded(selectedMailsIndexList, j, singleMail);

		if (textType == 0)
		{
			printSingleMailToTextFile(fp, i, fpm, textConfig, singleMail, addPageBreak);
		}
		else 
		{
			BOOL fullImgFilePath = FALSE;
			int pos = printSingleMailToHtmlFile(fp, i, fpm, textConfig, singleMail, addPageBreak, fullImgFilePath);
			if (pos < 0)
				continue;
		}

		if (progressBar && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				int deb = 1;
				break;
			}

			UINT_PTR dwProgressbarPos = 0;
			ULONGLONG workRangePos = j;
			BOOL needToUpdateStatusBar = progressTimer.UpdateWorkPos(workRangePos, dwProgressbarPos);
			if (needToUpdateStatusBar)
			{
				int nFileNum = j + 1;
				if (textType == 0)
				{
					CString fmt = L"Printing mails to single TEXT file ... %d of %d";
					ResHelper::TranslateString(fmt);
					fileNum.Format(fmt, nFileNum, cnt);
				}
				else
				{
					CString fmt = L"Printing mails to single HTML file ... %d of %d";
					ResHelper::TranslateString(fmt);
					fileNum.Format(fmt, nFileNum, cnt);
				}

				if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(dwProgressbarPos));

				int debug = 1;
			}
		}
	}

	if ((textType == 1) && (singleMail == FALSE))
		printDummyMailToHtmlFile(fp);
	
	UINT newstep = 100;
	int nFileNum = cnt;

	if (textType == 0)
	{
		CString fmt = L"Printing mails to single TEXT file ... %d of %d";
		ResHelper::TranslateString(fmt);
		fileNum.Format(fmt, nFileNum, cnt);
	}
	else
	{
		CString fmt = L"Printing mails to single HTML file ... %d of %d";
		ResHelper::TranslateString(fmt);
		fileNum.Format(fmt, nFileNum, cnt);
	}

	if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(newstep));

	int deb = 1;

	fp.Close();
	fpm.Close();

	return 1;
}


void MboxMail::assert_unexpected()
{
#ifdef _DEBUG
	int deb = 1;
	//_ASSERTE(0);
#else
	int deb = 1;
#endif
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

	if (m_InReplyId.GetLength() > maxLargeLen)
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

	TRACE(L"There is  %*ld percent of memory in use.\n",
		WIDTH, statex.dwMemoryLoad);
	TRACE(L"There are %*I64d total KB of physical memory.\n",
		WIDTH, statex.ullTotalPhys / DIV);
	TRACE(L"There are %*I64d free  KB of physical memory.\n",
		WIDTH, statex.ullAvailPhys / DIV);
	TRACE(L"There are %*I64d total KB of paging file.\n",
		WIDTH, statex.ullTotalPageFile / DIV);
	TRACE(L"There are %*I64d free  KB of paging file.\n",
		WIDTH, statex.ullAvailPageFile / DIV);
	TRACE(L"There are %*I64d total KB of virtual memory.\n",
		WIDTH, statex.ullTotalVirtual / DIV);
	TRACE(L"There are %*I64d free  KB of virtual memory.\n",
		WIDTH, statex.ullAvailVirtual / DIV);

	// Show the amount of extended memory available.

	TRACE(L"There are %*I64d free  KB of extended memory.\n",
		WIDTH, statex.ullAvailExtendedVirtual / DIV);
}

HintConfig::HintConfig()
{
	m_nHintBitmap = 0;
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

void HintConfig::ClearAllHints()
{
	m_nHintBitmap = 0xFFFFFFFF;
}


void HintConfig::LoadHintBitmap()
{
	BOOL retval;
	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_general, L"hintBitmap", m_nHintBitmap))
	{
		; //all done
	}
	else
	{
		m_nHintBitmap = 0xFFFFFFFF;
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_general, L"hintBitmap", m_nHintBitmap);
	}
}


void HintConfig::SaveToRegistry()
{
	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_general, L"hintBitmap", m_nHintBitmap);
}


void MboxMail::LoadHintBitmap()
{
	MboxMail::m_HintConfig.LoadHintBitmap();
}

BOOL MboxMail::ParseDateInFromField(char *p, char *end, SYSTEMTIME *sysTime)
{
	static char *tm = "Thu Oct 27 09:02:59 +0000 2011";
	static int tmlen = istrlen(tm);

	BOOL r = p < end;

	CStringA line;
	char *p_next = MimeParser::GetMultiLine(p, end, line);

	if (tmlen > line.GetLength())
		return FALSE;

	p = (char*)((LPCSTR)line);
	char *e = p + line.GetLength();
	p = e - tmlen;

	time_t timeDT = -1;
	
	char c;
	while (p < e)
	{
		c = tolower(*p);
		if ((c == '\n') || (c == '\r'))
			return FALSE;

		if (c == 'm')
		{
			if (TextUtilsEx::strncmpUpper2Lower(p, e, "mon", 3) == 0)
				break;
		}
		else if (c == 't')
		{
			if (TextUtilsEx::strncmpUpper2Lower(p, e, "tue", 3) == 0)
				break;
			else if (TextUtilsEx::strncmpUpper2Lower(p, e, "thu", 3) == 0)
				break;
		}
		else if (c == 'w')
		{
			if (TextUtilsEx::strncmpUpper2Lower(p, e, "wed", 3) == 0)
				break;
		}
		else if (c == 'f')
		{
			if (TextUtilsEx::strncmpUpper2Lower(p, e, "fri", 3) == 0)
				break;
		}
		else if (c == 's')
		{
			if (TextUtilsEx::strncmpUpper2Lower(p, e, "sat", 3) == 0)
				break;
			else if (TextUtilsEx::strncmpUpper2Lower(p, e, "sun", 3) == 0)
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

void MboxMail::LoadHintTextResource()
{
#ifdef _DEBUG
	if (ResHelper::g_LoadMenuItemsInfo == FALSE)
		return;

	CString strClassName;
	int hintNumber;
	CString hintText;
	for (hintNumber = 1; hintNumber < HintConfig::MaxHintNumber; hintNumber++)
	{
		hintText.Empty();
		CreateHintText(hintNumber, hintText);
		const wchar_t* p = (LPCWSTR)hintText;
		int contributorID = ResHelper::CONTRIB_HINTS;
		ResHelper::MergeAddItemInfo(hintText, strClassName);
	}
#endif
}

void MboxMail::ShowHint(int hintNumber, HWND h)
{
	CString hintText;
	//CStringA hintText;
	BOOL isHintSet = MboxMail::m_HintConfig.IsHintSet(hintNumber);
	if (isHintSet)
	{
		CreateHintText(hintNumber, hintText);
	}

	if (!hintText.IsEmpty())
	{
		ResHelper::TranslateString(hintText);
		CString title = L"Did you know?";
		ResHelper::TranslateString(title);
		::MessageBox(h, hintText, title, MB_OK | MB_ICONEXCLAMATION);
#if 0
		// Right now wil not use this
		MSGBOXPARAMSA mbp;
		mbp.cbSize = 0;
		mbp.hwndOwner = h;
		mbp.hInstance = 0;
		mbp.lpszText = hintText;
		mbp.lpszCaption = L"Did you know?";
		mbp.dwStyle = 0;
		mbp.lpszIcon = "";
		mbp.dwContextHelpId = 0;
		mbp.dwLanguageId = 0;
		int ret = MessageBoxIndirect(&mbp);
#endif
	}
	MboxMail::m_HintConfig.ClearHint(hintNumber);
}

void MboxMail::CreateHintText(int hintNumber, CString& hintText)
{
	// Or should I create table with all Hints to support potentially large number of hints ?
	if (hintNumber == HintConfig::GeneralUsageHint)
	{
		hintText.Append(
			L"To get started, please place one or more mbox files in\n"
			"the local folder and then select \"File->Select Folder...\"\n"
			"menu option to open that folder.\n"
			"\n"
			"Left-click on one of the loaded mbox files to view all associated mails.\n"
			"\n"
			"Please review the User Guide provided with the package\n"
			"and/or single and double right-click/left-click on any item\n"
			"within the Mbox Viewer window and try all presented options."
		);
	}
	else if (hintNumber == HintConfig::MsgWindowPlacementHint)
	{
		hintText.Append(
			L"You can place  Message Window at Bottom (default), Left or Right.\n"
			"\n"
			"Select \"View->Message Window...\" to configure."
		);
	}
	else if (hintNumber == HintConfig::PrintToPDFHint)
	{
		hintText.Append(
			L"Microsoft Edge Browser is used by default to print to PDF file.\n"
			"If Edge Browser is not installed on your computer\n"
			"Chrome Browser can be installed to print to PDF file."
			"\n\n"
			"You can configure Edge and Chrome browsers to optionally print default page header and footer"
			"\n\n"
			"You can also select \"File->Print Config->User Defined Script\" to "
			"use free wkhtmltopdf tool to print custom header and footer"
			" and evalute whether wkhtmltopdf works for you."
		);
	}
	else if (hintNumber == HintConfig::PrintToPrinterHint)
	{
		hintText.Append(
			L"Select \"File->Print Config->Page Setup\" to configure\n"
			"page header, footer, etc via Windows standard setup.\n"
		);
	}
	else if (hintNumber == HintConfig::MailSelectionHint)
	{
		hintText.Append(
			L"You can select multiple mails using standard Windows methods: "
			"\"Shift+Left-Click\", \"Ctrl+Left-Click\" and \"Ctrl+A\".\n\n"
			"Right-click on a single or multiple selected mails to see all available options.\n\n"
			"Right-click on a mail attachment to see all available options.\n\n"
			"Double Left-click on a mail to open folder with all associated files for that mail.\n\n"
			"Right-click on the Header Pane at the top of the Message Window to see additional context menu options.\n\n"
			"You can expand list of fields in the message header pane by setting \"View -> Message Header Pane Layout\" to Expanded.\n\n"
		);
	}
	else if (hintNumber == HintConfig::FindDialogHint)
	{
		hintText.Append(
			L"You can specify single ‘*’ character as the search string to FIND dialog to find subset of mails:\n"
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
			L"You can specify single ‘*’ character as the search string in any of the Filter fields in Advanced Find dialog to find subset of mails:\n"
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
			L"If you need to remove the background color when printing to\n"
			"PDF file directly you need to configure wkhtmltopdf for\n"
			"printing. Both MS Edge and Chrome browsers don't support\n"
			"removing the background color via the command line option\n\n"
			"Select \"File->Print Config->Page Setup\" to configure\n"
			"HTML2PDF-single-wkhtmltopdf.cmd script.\n\n"
			"Note that the background color can be removed via\n"
			"Page Setup when printing to PDF file\n"
			"via Print to Printer option and/or by opening the mail\n"
			"within a browser for printing.\n\n"
			"Another approach is to open the selected mail in web browser and configure to remove the background color in the Print Dialog\n\n"
		);
	}
	else if (hintNumber == HintConfig::AttachmentConfigHint)
	{
		hintText.Append(
			L"By default attachments other than those embeded into mail\n"
			"are shown in the attachment window.\n\n"
			"You can configure to shown all attachments, both inline and\n"
			"non-inline, by selecting\n"
			"\"File->Attachments Config->Attachment Window\" dialog\n\n"
			"Attachment Config dialog enables users to append image attachments to mails when viewing and/or ptinting\n"
		);
	}
	else if (hintNumber == HintConfig::MessageHeaderConfigHint)
	{
		hintText.Append(
			L"By default the mail text is shown in Message Window.\n"
			"Enable global \"View -> View Message Headers\" option to show the message header instead of the text.\n\n"
			"To switch between text and headers of the selected mail\n"
			"right-click on the header pane in Message Window and check/uncheck the \"View Message Header\" option.\n"
			"\n"
		);
	}
	else if (hintNumber == HintConfig::MessageRemoveFolderHint)
	{
		hintText.Append(
			L"The selected folder will be removed from the current view.\n"
			"The physical folder and content on the disk will not be deleted.\n"
			"\n"
		);
	}
	else if (hintNumber == HintConfig::MessageRemoveFileHint)
	{
		hintText.Append(
			L"The selected file will be removed from the current view.\n"
			"The physical file and content on the disk will not be deleted.\n"
			"\n"
		);
	}
	else if (hintNumber == HintConfig::MergeFilesHint)
	{
		hintText.Append(
			L"Mbox and eml mail files can also be merged via two command line options:\n\n"
			"-MBOX_MERGE_LIST_FILE=Path to File containing list of mbox files to merge\n\n"
			"-MBOX_MERGE_TO_FILE=Path to File to save merge results\n\n"
			"List of files to merge supports widcard file names\n\n\n"
			"Use \"Select root folder for merging..\" folder option to merge root folder and subfolders\n"
			"\n"
		);
	}
	else if (hintNumber == HintConfig::LabelsHint)
	{
		hintText.Append(
			L"Gmail Labels are supported by performing a separate step on a active mail folder.\n\n"
			"Right-click on the folder and select \"File->Gmail Labels->Create\" option.\n\n"
			"\n"
		);
	}
	else if (hintNumber == HintConfig::SubjectSortingHint)
	{
		hintText.Append(
			L"Sorting by subject creates subject threads. Emails within each subject thread are sorted by time.\n\n"
			"By default subject threads are sorted alphanumerically. "
			"Subject threads can be sorted by time by selecting  \"File->Options->time ordered threads\" option.\n"
			"\n"
		);
	}
}

// Fix static buffer allocation, it is bad !!!

static void LargeBufferAllocFailed()
{
	_ASSERTE(FALSE);
}

SimpleString *MboxMail::get_outbuf()
{
	if (m_outbufBusy)
		LargeBufferAllocFailed();
	m_outbufBusy = TRUE;
	return m_outbuf;
}
SimpleString *MboxMail::get_inbuf()
{
	if (m_inbufBusy)
		LargeBufferAllocFailed();
	m_inbufBusy = TRUE;
	return m_inbuf;
}
SimpleString *MboxMail::get_outdata()
{
	if (m_outdataBusy)
		LargeBufferAllocFailed();
	m_outdataBusy = TRUE;
	return m_outdata;
}
SimpleString *MboxMail::get_indata()
{
	if (m_indataBusy)
		LargeBufferAllocFailed();
	m_indataBusy = TRUE;
	return m_indata;
}
SimpleString *MboxMail::get_workbuf()
{
	if (m_workbufBusy)
		LargeBufferAllocFailed();
	m_workbufBusy = TRUE;
	return m_workbuf;
}
SimpleString *MboxMail::get_tmpbuf()
{
	if (m_tmpbufBusy)
		LargeBufferAllocFailed();
	m_tmpbufBusy = TRUE;
	return m_tmpbuf;
}
SimpleString *MboxMail::get_largebuf()
{
	if (m_largebufBusy)
		LargeBufferAllocFailed();
	m_largebufBusy = TRUE;
	return m_largebuf;
}

///
void MboxMail::rel_outbuf()
{
	if (!m_outbufBusy)
		LargeBufferAllocFailed();
	m_outbufBusy = FALSE;
}
void MboxMail::rel_inbuf()
{
	if (!m_inbufBusy)
		LargeBufferAllocFailed();
	m_inbufBusy = FALSE;
}
void MboxMail::rel_outdata()
{
	if (!m_outdataBusy)
		LargeBufferAllocFailed();
	m_outdataBusy = FALSE;
}
void MboxMail::rel_indata()
{
	if (!m_indataBusy)
		LargeBufferAllocFailed();
	m_indataBusy = FALSE;
}
void MboxMail::rel_workbuf()
{
	if (!m_workbufBusy)
		LargeBufferAllocFailed();
	m_workbufBusy = FALSE;
}
void MboxMail::rel_tmpbuf()
{
	if (!m_tmpbufBusy)
		LargeBufferAllocFailed();
	m_tmpbufBusy = FALSE;
}
void MboxMail::rel_largebuf()
{
	if (!m_largebufBusy)
		LargeBufferAllocFailed();
	m_largebufBusy = FALSE;
}

void MboxMail::DumpMailParseException(_int64 msgOffset)
{
#if 1
	static const char *cFromMailBegin = "From ";
	static const int cFromMailBeginLen = istrlen(cFromMailBegin);

	CString progDir;
	BOOL retDir = GetProgramDir(progDir);

	int seNumb = 0;
	CStringA szCause = seDescription(seNumb);

	int mailPosition = s_mails.GetCount();

	BOOL ret = TRUE;
	CFile fp;

	int bytes2Read = MAPPING_SIZE;

	SimpleString *res = MboxMail::m_outdata;
	res->Clear();

	CFileException ExError;
	if (fp.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
	{
		res->Resize(bytes2Read);
		//TRACE(L"offset = %lld\n", m_startOff);
		_int64 filePos = fp.Seek(msgOffset, SEEK_SET);
		UINT readLength = fp.Read(res->Data(), bytes2Read);
		if (readLength != bytes2Read)
			int deb = 1;
		res->SetCount(bytes2Read);
		fp.Close();
	}
	else
	{
		DWORD lastErr = ::GetLastError();

		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);
#if 1
		//HWND h = GetSafeHwnd();
		HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		CString fmt = L"Could not open mail file:\n\n\"%s\"\n\n%s";  // new format
		CString errorText = FileUtils::ProcessCFileFailure(fmt, MboxMail::s_path, ExError, lastErr, h);
#else
		CString txt = L"Could not open \"" + MboxMail::s_path;
		txt += L"\" mail file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);
#endif
		return;
	}

	char *p = res->Data();
	char *e = p + res->Count();

	int datalen = 0;
	char *data = p;
	char *msgEnd = p;
	char *msgStart = p;

	p = MimeParser::EatNewLine(p, e);

	while (p < e)
	{
		if (TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0) {
			msgEnd = p;
			break;
		}
		p = MimeParser::EatNewLine(p, e);
		msgEnd = p;
	}
	p = MimeParser::EatNewLine(p, e);
	msgEnd = p;
	while (p < e)
	{
		if (TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0) {
			msgEnd = p;
			break;
		}
		p = MimeParser::EatNewLine(p, e);
		msgEnd = p;
	}

	data = msgStart;
	datalen = IntPtr2Int(msgEnd - msgStart);

	if (datalen < 0)
		datalen = 0;

	wchar_t *se_mailDumpFileName = L"UnhandledExeption_MailDump.txt";
	BOOL retW = DumpMailData(se_mailDumpFileName, szCause, seNumb, mailPosition, data, datalen);
#if 0
	CString errorTxt;
	CString szCauseW(szCause);
	errorTxt.Format(L"Unhandled_Exception: Code=%8.8x Description=%s\n\n"
		"To help to diagnose the problem, created files\n\n%s\n\nin\n\n%s directory.\n\n"
		"Please provide the files to the development team.\n\n"
		"%s file contains mail data so please review the content before forwarding to the development.",
		seNumb, szCauseW, se_mailDumpFileName, progDir, se_mailDumpFileName);
#endif

	CString errorTxt;
	CString szCauseW(szCause);
	CString fmt = L"Unhandled_Exception: Code=%8.8x Description=%s\n\n"
		L"To help to diagnose the problem, created files\n\n%s\n\nin\n\n%s directory.\n\n"
		L"Please provide the files to the development team.\n\n"
		L"%s file contains mail data so please review the content before forwarding to the development."
		;
	ResHelper::TranslateString(fmt);
	errorTxt.Format(fmt, seNumb, szCauseW, se_mailDumpFileName, progDir, se_mailDumpFileName);

	AfxMessageBox((LPCWSTR)errorTxt, MB_OK | MB_ICONHAND);
#endif
}

void MboxMail::mbassert()
{
	if (!s_path.IsEmpty() && s_datapath.IsEmpty())
		int deb = 1;
	if (s_path.Compare(s_datapath))
		int deb = 1;
}

ProgressTimer::ProgressTimer(ULONGLONG workRangeFirstPos, ULONGLONG workRangeLastPos)
{

	// intsafe.h  supports safe LongLongToDWord, etc.
	// ULONGLONG ULLONG_MAX   0xffffffffffffffff or 18446744073709551615

	m_workLastPos = 0;
	m_workCurrentPos = 0;
	m_workUpdateCnt = 0;

	m_maxTimeValue = 0;
	m_startTime = 0;
	m_CurrentTime = 0;

	m_progressBarSize = 100;
	m_maxTimeValue = 0xffffffff;  // or 4294967295

	m_workRangeFirstPos = workRangeFirstPos;
	m_workRangeLastPos = workRangeLastPos;

	m_workSize = m_workRangeLastPos - m_workRangeFirstPos;

	m_workStepSize = (double)m_workSize / m_progressBarSize;
	if (m_workStepSize < 0)
		m_workStepSize = 0;

	m_checkNewWorkPosRate = 100;
	// Force to call GetTickCount after each task done to simplify progress reporting
	// GetTickCount overhead acceptable
	m_checkNewWorkPosRate = 1;  
}

BOOL ProgressTimer::UpdateWorkPos(ULONGLONG workRangePos, UINT_PTR& dwProgressbarPos, BOOL force)
{
	// GetTickCount64() lowest overhead, 10-16 msec tick
	// timeGetTime()
	// timeGetSystemTime()
	// QueryPerformanceCounter()

	if (m_workStepSize == 0)
	{
		dwProgressbarPos = 0;
		return FALSE;
	}

	if (force)
	{
		dwProgressbarPos = (UINT_PTR)((double)workRangePos / m_workStepSize);
		if (dwProgressbarPos > m_progressBarSize)
			dwProgressbarPos = m_progressBarSize;
		m_startTime = m_CurrentTime;
		return TRUE;
	}

	m_workUpdateCnt++;
	ULONGLONG deltaElapsedTime = 0;
	// 
	if ((m_workUpdateCnt % m_checkNewWorkPosRate) == 0)
	{
		m_CurrentTime = GetTickCount64();
		
		// Handle Overflow
		if (m_CurrentTime >= m_startTime)
		{
			deltaElapsedTime = m_CurrentTime - m_startTime;
		}
		else  // overflow or time goes backwards ?? no harm done if backwards result in huge deltaElapsedTime
		{
			deltaElapsedTime = (m_maxTimeValue - m_startTime) + m_CurrentTime;
			_ASSERTE(m_CurrentTime < m_startTime);
		}

		ULONGLONG elapsed_seconds = deltaElapsedTime / 1000;
		ULONGLONG elapsed_milliseconds = deltaElapsedTime % 1000;

		if (deltaElapsedTime >= 100)
		{
			dwProgressbarPos = (UINT_PTR)((double)workRangePos / m_workStepSize);
			if (dwProgressbarPos > m_progressBarSize)
				dwProgressbarPos = m_progressBarSize;
			m_startTime = m_CurrentTime;
			return TRUE;
		}
	}
	return FALSE;
}

void MailEncodingStats::UpdateTextEncodingStats(INT stats[], CStringA & fld, UINT code)
{
	if (fld.IsEmpty())
		return;
	UINT CP_ASCII = 20127;
	if (code == 0)
	{
		int i;
		for (i = 0; i < fld.GetLength(); i++)
		{
			char c = fld.GetAt(i);
			if (__isascii(c) == 0)
				break;
		}
		if (i == fld.GetLength())
			stats[TEXT_ENCODING_ASCII]++;
		else
			stats[TEXT_ENCODING_0]++;
	}
	else if (code == CP_ASCII)
	{
		int i;
		for (i = 0; i < fld.GetLength(); i++)
		{
			char c = fld.GetAt(i);
			if (__isascii(c) == 0)
				break;
		}
		if (i == fld.GetLength())
			stats[TEXT_ENCODING_ASCII]++;
		else
			stats[TEXT_ENCODING_0]++;
	}
	else if (code == CP_UTF8)
		stats[TEXT_ENCODING_UTF8]++;
	else
		stats[TEXT_ENCODING_OTHER]++;
}

void MailEncodingStats::PrintTextEncodingStats(CString &label, INT stats[], CString& stats2Text)
{
	CString text;
	text.Format(L"%8s  %06d %06d %06d %06d\n", label, 
		stats[TEXT_ENCODING_0], stats[TEXT_ENCODING_UTF8], stats[TEXT_ENCODING_UTF8], stats[TEXT_ENCODING_OTHER]);
	stats2Text.Append(text);
}

void MailEncodingStats::PrintTextEncodingStats(MailEncodingStats& textEncodingStats, CString& stats2Text)
{
	stats2Text.Format(L"\n%8s  %6s %6s %6s %6s\n", L"", L"ANSI", L"ASCII", L"UTF8", L"OTHER");
	MailEncodingStats::PrintTextEncodingStats(CString(L"From"), m_from, stats2Text);
	MailEncodingStats::PrintTextEncodingStats(CString(L"To"), m_to, stats2Text);
	MailEncodingStats::PrintTextEncodingStats(CString(L"CC"), m_cc, stats2Text);
	MailEncodingStats::PrintTextEncodingStats(CString(L"BCC"), m_bcc, stats2Text);
	MailEncodingStats::PrintTextEncodingStats(CString(L"SUBJECT"), m_subj, stats2Text);
	MailEncodingStats::PrintTextEncodingStats(CString(L"TEXT"), m_bodyTEXT, stats2Text);
	MailEncodingStats::PrintTextEncodingStats(CString(L"HTML"), m_bodyHTML, stats2Text);
}

void MailEncodingStats::Clear()
{
	INT *ar = &m_from[0];
	ar[0] = 0; ar[1] = 0; ar[2] = 0;
	ar = &m_to[0];
	ar[0] = 0; ar[1] = 0; ar[2] = 0;
	ar = &m_cc[0];
	ar[0] = 0; ar[1] = 0; ar[2] = 0;
	ar = &m_bcc[0];
	ar[0] = 0; ar[1] = 0; ar[2] = 0;
	ar = &m_subj[0];
	ar[0] = 0; ar[1] = 0; ar[2] = 0;
	ar = &m_bodyTEXT[0];
	ar[0] = 0; ar[1] = 0; ar[2] = 0;
	ar = &m_bodyHTML[0];
	ar[0] = 0; ar[1] = 0; ar[2] = 0;
}

void MboxMail::CollectTextEncodingStats(MailEncodingStats& textEncodingStats)
{
	UINT CP_ASCII = 20127;
	MailBodyContent* body;
	CStringA bodyDataA("body");

#if 0
	// Test specific case
	MboxMail* m;
	for (int i = 0; i < MboxMail::s_mails_ref.GetSize(); i++)
	{
		m = s_mails_ref[i];
		//MailEncodingStats::UpdateTextEncodingStats(m_textEncodingStats.m_from, m->m_from, m->m_from_charsetId);
		//MailEncodingStats::UpdateTextEncodingStats(m_textEncodingStats.m_to, m->m_to, m->m_to_charsetId);
		//MailEncodingStats::UpdateTextEncodingStats(m_textEncodingStats.m_cc, m->m_cc, m->m_cc_charsetId);
		//MailEncodingStats::UpdateTextEncodingStats(m_textEncodingStats.m_bcc, m->m_bcc, m->m_bcc_charsetId);
		MailEncodingStats::UpdateTextEncodingStats(m_textEncodingStats.m_subj, m->m_subj, m->m_subj_charsetId);
	}
#else
	MboxMail* m;
	for (int i = 0; i < MboxMail::s_mails_ref.GetSize(); i++)
	{
		m = s_mails_ref[i];
		MailEncodingStats::UpdateTextEncodingStats(m_textEncodingStats.m_from, m->m_from, m->m_from_charsetId);
		MailEncodingStats::UpdateTextEncodingStats(m_textEncodingStats.m_to, m->m_to, m->m_to_charsetId);
		MailEncodingStats::UpdateTextEncodingStats(m_textEncodingStats.m_cc, m->m_cc, m->m_cc_charsetId);
		MailEncodingStats::UpdateTextEncodingStats(m_textEncodingStats.m_bcc, m->m_bcc, m->m_bcc_charsetId);
		MailEncodingStats::UpdateTextEncodingStats(m_textEncodingStats.m_subj, m->m_subj, m->m_subj_charsetId);

		for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
		{
			body = m->m_ContentDetailsArray[j];
			if (body->m_contentType.CompareNoCase("text/plain") == 0)
			{
				MailEncodingStats::UpdateTextEncodingStats(m_textEncodingStats.m_bodyTEXT, bodyDataA, body->m_pageCode);
			}
			else if (body->m_contentType.CompareNoCase("text/html") == 0)
			{
				MailEncodingStats::UpdateTextEncodingStats(m_textEncodingStats.m_bodyHTML, bodyDataA, body->m_pageCode);
				if (body->m_pageCode == CP_ASCII)
					int deb = 1;

				if ((body->m_pageCode != 0) && (body->m_pageCode != CP_ASCII) && (body->m_pageCode != CP_UTF8))
					int deb = 1;
			}
		}
	}
#endif
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


