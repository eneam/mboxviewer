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

HintConfig MboxMail::m_HintConfig;

MailBodyPool *MailBody::m_mpool = new MailBodyPool;

ThreadIdTableType *MboxMail::m_pThreadIdTable = 0;
MessageIdTableType *MboxMail::m_pMessageIdTable = 0;
MboxMailTableType *MboxMail::m_pMboxMailTable = 0;
MboxMail::MboxMailMapType *MboxMail::m_pMboxMailMap = 0;
int MboxMail::m_nextGroupId = 0;

SimpleString* MboxMail::m_outbuf = new SimpleString(10000);
SimpleString* MboxMail::m_inbuf = new SimpleString(10000);
SimpleString* MboxMail::m_outdata = new SimpleString(10000);
SimpleString* MboxMail::m_indata = new SimpleString(10000);
SimpleString* MboxMail::m_workbuf = new SimpleString(10000);
SimpleString* MboxMail::m_tmpbuf = new SimpleString(10000);
SimpleString* MboxMail::m_largebuf = new SimpleString(10000);

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

UINT getCodePageFromHtmlBody(SimpleString *buffer, std::string &charset);

///////
// Kept adding and adding Print to functions but now cleanup is needed, better reusability, possible abstractions, error handling, etc
// Postponed to the future relases since the large effort is needed
///////

bool MailBodyContent::IsAttachment()
{
	if (!m_attachmentName.IsEmpty() ||
		!m_contentId.IsEmpty() ||
		!m_contentLocation.IsEmpty() ||
		(m_contentDisposition.CollateNoCase("attachment") == 0)
		)
		return true;
	else
		return false;
}

struct MsgIdHash {
public:
	hashsum_t operator()(const MboxMail *key) const
	{
		hashsum_t hashsum = StrHash((const char*)(LPCSTR)key->m_messageId, key->m_messageId.GetLength());
		return hashsum;
	}
};

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
			(key1->m_to == key2->m_to) )
			return true;
		else
			return false;
	}
};

struct MessageIdHash {
public:
	hashsum_t operator()(const CString *key) const
	{
		hashsum_t hashsum = StrHash((const char*)(LPCSTR)*key, key->GetLength());
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
	hashsum_t operator()(const CString *key) const
	{
		hashsum_t hashsum = StrHash((const char*)(LPCSTR)*key, key->GetLength());
		return hashsum;
	}
};

struct ThreadIdEqual {
public:
	bool operator()(const CString *key1, const CString *key2) const
	{
		if (*key1 == *key2)
			return true;
		else
			return false;
	}
};

void FolderContext::SetFolderPath(CString &folderPath)
{
	CString path = folderPath;
	path.TrimRight("\\");
	if (path.IsEmpty())
	{
		MboxMail::assert_unexpected();
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath", path);
		m_folderPath = "";
		m_dataFolderPath = "";
		return;
	}

	path.Append("\\");

	if (m_folderPath.Compare(path) == 0)
	{
		m_setPathCount++;  // redundant calls count
		return;
	}

	if (m_setPathCount > 1)
		TRACE("SetFolderPath: SetPathCount=%d\n", m_setPathCount);

	m_setPathCount = 0;

	CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath", path);

	m_folderPath = path;

	if (m_rootDataFolderPathConfig.IsEmpty())
		m_rootDataFolderPath.Empty();
	else
		m_rootDataFolderPath = m_rootDataFolderPathConfig;

	if (m_rootDataFolderPath.IsEmpty())
	{
		m_dataFolderPath = path + "MBoxViewer\\";
		BOOL isReadOnly = FileUtils::IsReadonlyFolder(m_dataFolderPath);
		if (!isReadOnly)
		{
			return;
		}
		else
		{
			m_rootDataFolderPath = FileUtils::CreateMboxviewLocalAppPath();
			m_dataFolderPath = m_rootDataFolderPath + "MBoxViewer\\";
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
		m_dataFolderPath = m_rootDataFolderPath + "MBoxViewer\\";
		if (!m_rootDataFolderPathSubFolderConfig.IsEmpty())
			m_dataFolderPath.Append(m_rootDataFolderPathSubFolderConfig);
	}

	CString fileName;
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(path, driveName, directory, fileNameBase, fileNameExtention);

	driveName.TrimRight(":");
	m_dataFolderPath = m_dataFolderPath + driveName + directory;

	int deb = 1;
}
void FolderContext::GetFolderPath(CString &folderpath)
{
	folderpath = m_folderPath;
}
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
	path.TrimRight("\\");
	path.Append("\\");
	MboxMail::s_datapath = "";
	CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath", path);
	s_folderContext.SetFolderPath(lastPath);

	if (path.IsEmpty())
	{
		return;
	}

	if (!s_folderContext.m_rootDataFolderPath.IsEmpty())
	{
		if (strncmp(s_folderContext.m_dataFolderPath, s_folderContext.m_rootDataFolderPath, s_folderContext.m_rootDataFolderPath.GetLength()))
			MboxMail::assert_unexpected();
	}
	s_datapath = s_folderContext.m_dataFolderPath;

	int deb = 1;
}


// This is very expensive; Need to determine once per current selection
CString MboxMail::GetDataPath(CString &path)
{
	if (path.IsEmpty())
	{
		return path;
	}

	// Set database path; Verify that the default mbox folder is not Read-Only
	CString TestFile = path + "\\" + ".WriteSupported";
	if (FileUtils::CreateDirectory(TestFile))
	{
			BOOL retD = FileUtils::RemoveDirW(TestFile, true, true);
			return path;
	}

	CString dataPath = MboxMail::SetLastDataPath(&path);
	return dataPath;
}

CString MboxMail::GetLastPath()
{
	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");

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

	path.TrimRight("\\");
	path.Append("\\");

	CString fileName;
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(path, driveName, directory, fileNameBase, fileNameExtention);

	driveName.TrimRight(":");
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

BOOL MboxMail::GetBody(CFile fp, CString &res)
{
	BOOL ret = TRUE;
	//CFile fp;
	//if (fp.Open(s_path, CFile::modeRead | CFile::shareDenyWrite))
	{
		char *p = res.GetBufferSetLength(m_length);
		//TRACE("offset = %lld\n", m_startOff);
		ULONGLONG pos = fp.Seek(m_startOff, SEEK_SET);  // TODO: handle errors
		int retR = fp.Read(p, m_length);   // TODO: handle errors
		char *ms = strchr(p, '\n'); //"-Version: 1.0");  // TODO: optimize for some cases
		if (ms) {
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
		//fp.Close();  // TODO: verify why Close() was not called
	}
#if 0
	else
	{
		DWORD err = GetLastError();
		TRACE("Open Mail File failed err=%ld\n", err);
		ret = FALSE;
	}
#endif
	return ret;
};

BOOL MboxMail::GetBody(CString &res)
{
	BOOL ret = TRUE;
	CFile fp;
	CFileException ExError;
	if (fp.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
	{
		char *p = res.GetBufferSetLength(m_length);
		//TRACE("offset = %lld\n", m_startOff);
		fp.Seek(m_startOff, SEEK_SET);
		fp.Read(p, m_length);
		char *ms = strchr(p, '\n'); //"-Version: 1.0");
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
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = _T("Could not open \"") + MboxMail::s_path;
		txt += _T("\" mail file.\n");
		txt += exErrorStr;

		TRACE(_T("%s\n"), txt);
		//errorText = txt;

		ret = FALSE;
	}
	return ret;
};

BOOL MboxMail::GetBody(CFile &fp, SimpleString *res, int maxLength)
{
	BOOL ret = TRUE;
	//CFile fp;
	//res->Clear();
	int bytes2Read = m_length;
	if (maxLength > 0)
	{
		if (maxLength < m_length)
			bytes2Read = maxLength;
	}
	//if (fp.Open(s_path, CFile::modeRead | CFile::shareDenyWrite))
	{
		res->Resize(bytes2Read);
		//TRACE("offset = %lld\n", m_startOff);
		fp.Seek(m_startOff, SEEK_SET);
		UINT readLength = fp.Read(res->Data(), bytes2Read);
		if (readLength != bytes2Read)
			int deb = 1;
		res->SetCount(bytes2Read);
		//fp.Close();
	}
#if 0
	else
	{
		DWORD err = GetLastError();
		TRACE("Open Mail File failed err=%ld\n", err);
		ret = FALSE;
	}
#endif
	return ret;
};

BOOL MboxMail::GetBody(SimpleString *res, int maxLength)
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
		//TRACE("offset = %lld\n", m_startOff);
		fp.Seek(m_startOff, SEEK_SET);
		UINT readLength = fp.Read(res->Data(), bytes2Read);
		if (readLength != bytes2Read)
			int deb = 1;
		res->SetCount(bytes2Read);
		fp.Close();
	}
	else
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = _T("Could not open \"") + MboxMail::s_path;
		txt += _T("\" mail file.\n");
		txt += exErrorStr;

		TRACE(_T("%s\n"), txt);
		//errorText = txt;

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

char szFrom5[] = "From ";
char szFrom6[] = "\nFrom ";
char	*g_szFrom;
int		g_szFromLen;

bool MboxMail::Process(register char *p, DWORD bufSize, _int64 startOffset,  bool bFirstView, bool bLastView, _int64 &lastStartOffset, bool bEml, _int64 &msgOffset, CString &statusText, BOOL parseContent)
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
	CString msgId, thrdId;
	time_t tdate = -1;
	time_t tdate_fromField = 1;
	bool	bTo = true, bFrom = true, bSubject = true, bDate = true, bRcvDate = true; // indicates not found, false means found
	bool bMsgId = true, bThrdId = true;
	bool bCC = true, bBCC = true;
	char *msgStart = NULL;
	int recv = TRUE;
	int curstep = (int)(startOffset / (s_step?s_step:1));
	CString line;
	CString rcved;

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
	//TCHAR szCauseBuffer[256];

	char *stackDumpFileName = "Exception_StackTrace.txt";
	char *mailDumpFileName = "Exception_MailDump.txt";
	char *exceptionName = "UnknownException";
	UINT seNumb = 0;
	char *szCause = "Unknown";
	BOOL exceptionThrown = FALSE;
	BOOL exceptionDone = FALSE;

	char *pat = "From:  XXXX";
	int patLen = istrlen(pat);

	BOOL headerDone = FALSE;
	while (p < (e - 4))   // TODO: why (e - 4) ?? chaneg 4 -> 5
	{
		{
			// EX_TEST
#if 0
			if ((s_mails.GetCount() == 3000) && (exceptionDone == FALSE))
			{
				//int i = 0;  intEx = rand() / i;   // This will throw a SE (divide by zero).
				//char* szTemp = (char*)1;
				//strcpy_s(szTemp, 1000, "A");
				//*badPtr = 'a';
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
							date_fromField = tt.FormatGmtTm(format);
							if (date_fromField.IsEmpty())
								date_fromField = tt.FormatLocalTm(format);
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
							m->m_length = IntPtr2Int(p - msgStart);
							m->m_to = TextUtilsEx::DecodeString(to, m->m_to_charset, m->m_to_charsetId);
							m->m_from = TextUtilsEx::DecodeString(from, m->m_from_charset, m->m_from_charsetId);
							m->m_subj = TextUtilsEx::DecodeString(subject, m->m_subj_charset, m->m_subj_charsetId);
							m->m_subj.Trim();
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
								if (statusText.IsEmpty())
									mailNum.Format(_T("Parsing archive file to create index file ... %d"), s_mails.GetCount());
								else
									mailNum.Format(_T("%s ... %d"), statusText, s_mails.GetCount());

								_int64 curOff = startOffset + (p - orig);
								int step = (int)(curOff / s_step);

								if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(mailNum, (UINT_PTR)(step));
							}

							if ((s_mails.GetCount() % 100000) == 1024)
							{
								int deb = 1;
							}

							if (pCUPDUPData && pCUPDUPData->ShouldTerminate())
								break;
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
					bTo = bFrom = bSubject = bDate = bRcvDate = bCC = bBCC = bMsgId  = bThrdId  = true;
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
				else if (bMsgId && TextUtilsEx::strncmpUpper2Lower(p, e, cMsgId, cMsgIdLen) == 0)
				{
					bMsgId = false;
					p = MimeParser::GetMultiLine(p, e, line);
					msgId = line.Mid(cMsgIdLen);
					msgId.Trim();
				}
				else if (bThrdId && TextUtilsEx::strncmpUpper2Lower(p, e, cThreadId, cThreadIdLen) == 0)
				{
					bThrdId = false;
					p = MimeParser::GetMultiLine(p, e, line);
					thrdId = line.Mid(cThreadIdLen);
					thrdId.Trim();
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
					m->m_length = IntPtr2Int(p - msgStart);
					m->m_to = TextUtilsEx::DecodeString(to, m->m_to_charset, m->m_to_charsetId);
					m->m_from = TextUtilsEx::DecodeString(from, m->m_from_charset, m->m_from_charsetId);
					m->m_subj = TextUtilsEx::DecodeString(subject, m->m_subj_charset, m->m_subj_charsetId);
					m->m_subj.Trim();
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
					m->m_messageId = msgId;
					m->m_threadId = thrdId;

					//m->m_crc32 = TextUtilities::CalcCRC32(msgStart, m->m_length);

					MailBodyContent body;
					char *bodyStart = msgStart;
					char *bodyEnd = msgStart + m->m_length;
					if (parseContent)
						char *nextStart = ParseContent(m, bodyStart, bodyEnd);

					s_mails.Add(m);

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

void MboxMail::Parse(LPCSTR path)
{
	Destroy();

	s_mails.SetSize(100000, 100000);
	s_mails.SetCountKeepData(0);

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
	targetPrintSubFolder.Empty();
	CString imageCachePath;
	errorText.Empty();
	bool ret2 = MboxMail::GetCachePath(rootPrintSubFolder, targetPrintSubFolder, imageCachePath, errorText, &cpath);
	if (errorText.IsEmpty() && FileUtils::PathDirExists(imageCachePath)) {
		pCUPDUPData->SetProgress(_T("Deleting all related files in the ImageCache directory ..."), 0);
		FileUtils::RemoveDirW(imageCachePath, true);
	}

	//MboxMail::s_path = path;
	MboxMail::SetMboxFilePath(CString(path));
	CString lastPath = MboxMail::GetLastPath();
	CString lastDataPath = MboxMail::GetLastDataPath();

	bool	bEml = false;
	int l = istrlen(path);
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

	// TODO:  Need to consider to redo reading and create  Character Stream for reading
	CString statusText;
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
		_int64 msgOffset = viewOffset;
		BOOL processException = FALSE;
#ifdef USE_STACK_WALKER
		try
		{
#endif
			MboxMail::Process(p, viewBufSize, viewOffset, firstView, lastView, lastStartOffset, bEml, msgOffset, statusText);
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


void MboxMail::Parse_LabelView(LPCSTR path)
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
	CString parsingFileText = _T("Parsing \"") + mailFile + _T("\"");

	if (pCUPDUPData)
		pCUPDUPData->SetProgress(parsingFileText, 0);  // works but doesn't always fit into progress bar
	//pCUPDUPData->SetProgress(_T("Parsing archive file to create index file ..."), 0);

	// TODO: due to breaking the file into multiple chunks, it looks some emails can be lost : Fixed

	bool firstView = true;;
	bool lastView = false;
	DWORD bufSize = 0;
	_int64 aligned_offset = 0;
	_int64 delta = 0;

	// TODO:  Need to consider to redo reading and create  Character Stream for reading
	_int64 lastStartOffset = 0;
	while ((lastView == false) && !(pCUPDUPData && pCUPDUPData->ShouldTerminate()))
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
		_int64 msgOffset = viewOffset;
		BOOL processException = FALSE;
#ifdef USE_STACK_WALKER
		try
		{
#endif
			MboxMail::Process(p, viewBufSize, viewOffset, firstView, lastView, lastStartOffset, bEml, msgOffset, parsingFileText, FALSE);
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

int ExtractSubject(CString &subject, char *&subjectText) 
{
	char *subj = (char *)(LPCSTR)subject;
	int subjlen = subject.GetLength();

	//TODO: should I skip white spaces 
	// should compare

	while (subjlen >= 4)
	{
		if (subj[0] == 'R')
		{
			if ((strncmp((char*)subj, "Re: ", 4) == 0) || (strncmp((char*)subj, "RE: ", 4) == 0)) {
				subj += 4; subjlen -= 4;
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

	HANDLE mbox_hFile = CreateFile(MboxMail::s_datapath, GENERIC_READ, FILE_SHARE_READ, NULL,
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

			if (!FileUtils::Write2File(hFile, CRLF, istrlen(CRLF), &nwritten)) {
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

			if (!FileUtils::Write2File(hFile, CRLF, istrlen(CRLF), &nwritten)) {
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
	if (!FileUtils::Write2File(hFile, CRLF, istrlen(CRLF), &nwritten)) {
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
	for (int i = 0; i < cnt; i++)
	{
		for (int j = 0; j < mails[i]->m_ContentDetailsArray.size(); j++) {
			delete mails[i]->m_ContentDetailsArray[j];
			mails[i]->m_ContentDetailsArray[j] = 0;
		}
		delete mails[i];
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
#define HASH_ARRAY_SIZE 50013
//
int MboxMail::getMessageId(CString *key)
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

bool MboxMail::insertMessageId(CString *key, int val)
{
	if (m_pMessageIdTable == 0)
		createMessageIdTable(HASH_ARRAY_SIZE);

	CString *mapKey = new CString;
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
int MboxMail::getThreadId(CString *key)
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

bool MboxMail::insertThreadId(CString *key, int val)
{
	if (m_pThreadIdTable == 0)
		createThreadIdTable(HASH_ARRAY_SIZE);

	CString *mapKey = new CString;
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
		_ASSERT(orphanCount == 0);
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
			int mId = getMessageId(&m->m_messageId);
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
			int mId = getMessageId(&m->m_replyId);
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
			int rId = getMessageId(&m->m_replyId);
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
#if 0
// Two versions; different sort order
bool MboxMail::sortConversations()
{
	MboxMail *m;
	int i_ref = 0;

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
	int i;
	for (i = 0; i < s_mails.GetSize(); i++)
	{
		m = s_mails[i];
		if (m->m_done)
			continue;

		if (m->m_prevMail >= 0)
		{
			MboxMail::assert_unexpected();
			continue;  // root mail of the list
		}

		if (m->m_nextMail < 0)
		{
			m->m_groupId = m_nextGroupId;
			m->m_groupColor = currentGroupColor;
			m->m_done = true;

			s_mails_ref[i_ref++] = m;
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

			s_mails_ref[i_ref++] = m;
			refCnt++;

			while (m->m_nextMail >= 0)
			{
				m = s_mails[m->m_nextMail];
				if (m->m_done)
					continue; // should never be here ??

				m->m_groupId = m_nextGroupId;
				m->m_groupColor = currentGroupColor;
				m->m_done = true;

				s_mails_ref[i_ref++] = m;
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
		_ASSERT(refCnt == refSize);
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
#else
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
		_ASSERT(refCnt == refSize);
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
#endif

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
			else
			{
				m_thread_oldest = m;
				m->m_mail_cnt = 1;
				m->m_mail_index = i;
				MboxMail::s_mails_selected.Add(m);
			}
		}
		else
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

	_ASSERT(mailCount == MboxMail::s_mails_merged.GetCount());

	MboxMail::s_mails.CopyKeepData(MboxMail::s_mails_merged);

	int deb = 1;
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
	int limitedLength = (int)(p - str->Data());
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
			int ret = MboxMail::DetermineEmbeddedImages(m, fpm);

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
		CString txt = _T("Please open mail file first.");
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	CString datapath = MboxMail::GetLastDataPath();
	CString path = MboxMail::GetLastPath();
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
		CFileException ExError;
		if (!fp.Open(csvFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &ExError))
		{
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt = _T("Could not create \"") + csvFile;
			txt += _T("\" file.\n");
			txt += exErrorStr;

			//TRACE(_T("%s\n"), txt);

			CFileStatus rStatus;
			BOOL ret = fp.GetStatus(rStatus);

			//errorText = txt;

			HWND h = NULL; // we don't have any window yet
			int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}

		CFile fpm;
		if (csvConfig.m_bContent)
		{
			CFileException ExError;
			if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
			{
				CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

				CString txt = _T("Could not open \"") + MboxMail::s_path;
				txt += _T("\" mail file.\n");
				txt += exErrorStr;

				//TRACE(_T("%s\n"), txt);
				//errorText = txt;

				HWND h = NULL; // we don't have any window yet
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
		int textlen = GetMailBody_mboxview(fpm, mailPosition, outbuf, pageCode);  // fast
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
					BOOL ret = TextUtilsEx::Str2CodePage(inbuf, pageCode, csvConfig.m_nCodePageId, outbuf, workbuf);
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

	SimpleString outbuf(1024, 256);
	SimpleString tmpbuf(1024, 256);

	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;
	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);

	MboxMail *m = s_mails[mailPosition];

	outbuf.Append(border);

	UINT pageCode = 0;
	CString txt;
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
				BOOL ret = TextUtilsEx::Str2CodePage(&tmpbuf, pageCode, textConfig.m_nCodePageId, inbuf, workbuf);
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
			else
			{
				CString lDateTime = tt.FormatGmtTm(format);
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
				else
				{
					CString lDateTime = tt.FormatGmtTm(format);
					strcpy(datebuff, (LPCSTR)lDateTime);
				}
			}

			if (fldCnt)
				outbuf.Append("\r\n");

			CString label;
			label.Format("ATTACHMENTS (%s):", attachmentFileNamePrefix);
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
	int textlen = GetMailBody_mboxview(fpm, mailPosition, outbuf, pageCode, textType);  // returns pageCode
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
	int ret = DetermineEmbeddedImages(m, fpm);
	return ret;
}

int MboxMail::DetermineEmbeddedImages(MboxMail *m, /*in - mail body*/ CFile &fpm)
{
	SimpleString *outbuflarge = MboxMail::m_outdata;

	BOOL htmlBodyFound = FALSE;
	MailBodyContent *body;
	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];

		if (body->m_contentType.CompareNoCase("text/html") == 0)
		{
			htmlBodyFound = TRUE;
		}
	}
	if (htmlBodyFound == FALSE)
	{
		m->m_DetermineEmbeddedImagesDone = 2;
		return 1;
	}

	outbuflarge->Clear();
	UINT pageCode = 0;
	int textType = 1; // try Html
	int textlen = GetMailBody_mboxview(fpm, m, outbuflarge, pageCode, textType);  // returns pageCode
	if (textlen != outbuflarge->Count())
		int deb = 1;

	if (outbuflarge->Count() != 0)
	{
		if (NListView::DetermineEmbeddedImages(outbuflarge->Data(), outbuflarge->Count(), m) > 0)
			m->m_DetermineEmbeddedImagesDone = 1;
	}

	m->m_DetermineEmbeddedImagesDone = 1;

	return 1;
}

int MboxMail::printAttachmentNamesAsHtml(CFile *fpm, int mailPosition, SimpleString *outbuf, CString &attachmentFileNamePrefix)
{
#if 1
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
			// TODO: critical failure
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt = _T("Could not open \"") + MboxMail::s_path;
			txt += _T("\" mail file.\n");
			txt += exErrorStr;

			TRACE(_T("%s\n"), txt);
			//errorText = txt;

			return FALSE;
		}
		fpm = &mboxFp;
	}

	CString errorText;
	CString attachmentCachePath;
	CStringW printCachePathW;
	CString rootPrintSubFolder = "AttachmentCache";
	CString targetPrintSubFolder;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, attachmentCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		// TODO: what to do ?
		//int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		//return -1;
	}

	DWORD error;
	retval = TextUtilsEx::Ansi2Wide(attachmentCachePath, printCachePathW, error);

	CString mboxFileNamePath = MboxMail::s_path;
	CString driveName;
	CString mboxFileDirectory;
	CString mboxFileNameBase;
	CString mboxFileNameExtention;

	FileUtils::SplitFilePath(mboxFileNamePath, driveName, mboxFileDirectory, mboxFileNameBase, mboxFileNameExtention);

	CString mboxFileName = mboxFileNameBase + mboxFileNameExtention;

	CStringW mboxFileNameBaseW;
	if (TextUtilsEx::Ansi2Wide(mboxFileNameBase, mboxFileNameBaseW, error))
	{
		int deb = 1;
	}

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
		ret = MboxMail::DetermineEmbeddedImages(mailPosition, *fpm);

	CString strDate;
	CString format;
	format = "%Y%m%d-%H%M%S-";

	MyCTime tt(m->m_timeDate);
	strDate = tt.FormatLocalTm(format);

	CString uID;
	uID.Format("%07d", m->m_index);

	attachmentFileNamePrefix = strDate;
	attachmentFileNamePrefix.Append(uID);

	bool showAllAttachments = false;
	AttachmentConfigParams *attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
	if (attachmentConfigParams)
	{
		showAllAttachments = attachmentConfigParams->m_bShowAllAttachments_Window;
	}

	AttachmentMgr attachmentDB;
	int attachmentCnt = 0;
	SimpleString fileName(1024);
	SimpleString href(1024);
	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];

		BOOL showAttachment = FALSE;
		if (body->IsAttachment())
		{
			if (showAllAttachments || !body->m_isEmbeddedImage)
				showAttachment = TRUE;
		}

		if (showAttachment)
		{
			UINT inCodePage = body->m_attachmentNamePageCode;

			SimpleString*buf = 0;
			NListView::DetermineAttachmentName(fpm, mailPosition, body, buf, nameW, attachmentDB);

			validNameW.Empty();
			FileUtils::MakeValidFileNameW(nameW, validNameW, bReplaceWhiteWithUnderscore);

			UINT outCodePage = CP_UTF8;
			BOOL ret2 = TextUtilsEx::WStr2CodePage((wchar_t*)(LPCWSTR)validNameW, validNameW.GetLength(), outCodePage, &validNameUTF8, error);

			// Create hyperlink

			if (attachmentCnt)
				outbuf->Append(" , ");
			else
				outbuf->Append(" ");

			CStringW filePathW = printCachePathW + L"\\" + validNameW;

			fileName.Clear();
			fileName.Append(attachmentFileNamePrefix, attachmentFileNamePrefix.GetLength());
			fileName.Append(' ');
			fileName.Append(validNameUTF8);
			encodeTextAsHtmlLink(fileName);

			//validNameUTF8.Append("<>;@&");  // for testing
			encodeTextAsHtmlLinkLabel(validNameUTF8);

			//outbuf->Append("\r\n<a href=\"file:../../AttachmentCache/");
			//outbuf->Append((char*)(LPCSTR)mboxFileNameBase, mboxFileNameBase.GetLength());
			outbuf->Append("\r\n<a href=\"file:///");

			href.Clear();
			href.Append(attachmentCachePath, attachmentCachePath.GetLength());
			href.Append('\\');
			href.Append(fileName.Data(), fileName.Count());
			href.Append("\"target=\"_blank\">");
			href.Append(validNameUTF8.Data(), validNameUTF8.Count());
			href.Append("</a>");

			outbuf->Append(href);
			attachmentCnt++;
		}
	}

	if (fpm_save == 0)
		mboxFp.Close();

	return 1;
#else
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
		if (!mboxFp.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite), &ExError)
		{
			// TODO: critical failure
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError); // TODO
			return FALSE;
		}
		fpm = &mboxFp;
	}


	BOOL bReplaceWhiteWithUnderscore = FALSE;
	DWORD error;
	BOOL ret;

	CStringW nameW;
	CStringW validNameW;
	CStringW fileNameW;

	MailBodyContent *body;
	CString fileName;

	CString errorText;
	CString printCachePath;
	CStringW printCachePathW;
	CString rootPrintSubFolder = "AttachmentCache";
	CString targetPrintSubFolder;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, printCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		// TODO: what to do ?
		//int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		//return -1;
	}

	ret = TextUtilsEx::Ansi2Wide(printCachePath, printCachePathW, error);

	CString mboxFileNamePath = MboxMail::s_path;
	CString driveName;
	CString mboxFileDirectory;
	CString mboxFileNameBase;
	CString mboxFileNameExtention;

	FileUtils::SplitFilePath(mboxFileNamePath, driveName, mboxFileDirectory, mboxFileNameBase, mboxFileNameExtention);

	CString mboxFileName = mboxFileNameBase + mboxFileNameExtention;

	CStringW mboxFileNameBaseW;
	if (TextUtilsEx::Ansi2Wide(mboxFileNameBase, mboxFileNameBaseW, error))
	{
		int deb = 1;
	}


	if (!m->m_DetermineEmbeddedImagesDone)
		ret = MboxMail::DetermineEmbeddedImages(mailPosition, *fpm);

	CString strDate;
	CString format;
	format = "%Y%m%d-%H%M%S-";

	MyCTime tt(m->m_timeDate);
	strDate = tt.FormatLocalTm(format);

	CString uID;
	uID.Format("%07d", m->m_index);

	CString attachmentFileNamePrefix = strDate;
	attachmentFileNamePrefix.Append(uID);
	CStringW attachmentFileNamePrefixW;
	if (TextUtilsEx::Ansi2Wide(attachmentFileNamePrefix, attachmentFileNamePrefixW, error))
	{
		int deb = 1;
	}

	htmlbuf.Append(L"\r\n<br><div><p>ATTACHMENTS:</p><br>");

	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];

		if (strncmp((LPCSTR)body->m_attachmentName, "ESTADISTICA", 11) == 0)
			int deb = 1;

		if (!body->m_attachmentName.IsEmpty() && !body->m_isEmbeddedImage)
		{
			if ((body->m_contentDisposition.CompareNoCase("attachment") == 0))
			{
				UINT inCodePage = body->m_attachmentNamePageCode;

				SimpleString *buf = 0;
				NListView::DetermineAttachmentName(fpm, mailPosition, body, buf, nameW);

				fileNameW = attachmentFileNamePrefixW + L" " + nameW;

				validNameW.Empty();
				FileUtils::MakeValidFileNameW(fileNameW, validNameW, bReplaceWhiteWithUnderscore);

				CStringW filePathW = printCachePathW + L"\\" + validNameW;

				htmlbuf.Append(L"<a href=\"file:../../AttachmentCache/");
				htmlbuf.Append(mboxFileNameBaseW);
				htmlbuf.Append(L"/");
				htmlbuf.Append(validNameW);
				htmlbuf.Append(L"\"target=\"_blank\">");
				htmlbuf.Append(nameW);
				htmlbuf.Append(L"</a>&nbsp&nbsp&nbsp&nbsp");
			}
		}
	}
	htmlbuf.Append(L"</div>");

	if (fpm_save == 0)
		mboxFp.Close();

	return 1;
#endif
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
			// TODO: critical failure
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt = _T("Could not open \"") + MboxMail::s_path;
			txt += _T("\" mail file.\n");
			txt += exErrorStr;

			TRACE(_T("%s\n"), txt);
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

	CString strDate;
	CString format;
	format = "%Y%m%d-%H%M%S-";

	MyCTime tt(m->m_timeDate);
	strDate = tt.FormatLocalTm(format);

	CString uID;
	uID.Format("%07d", m->m_index);

	attachmentFileNamePrefix = strDate;
	attachmentFileNamePrefix.Append(uID);

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
		if (body->IsAttachment())
		{
			if (showAllAttachments || !body->m_isEmbeddedImage)
				showAttachment = TRUE;
		}

		if (showAttachment)
		{
			UINT inCodePage = body->m_attachmentNamePageCode;

			SimpleString*buf = 0;
			NListView::DetermineAttachmentName(fpm, mailPosition, body, buf, nameW, attachmentDB);

			validNameW.Empty();
			FileUtils::MakeValidFileNameW(nameW, validNameW, bReplaceWhiteWithUnderscore);

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

	if ((abs(MboxMail::b_mails_which_sorted) != 99) && ((abs(MboxMail::b_mails_which_sorted) != 5) && (pListView->m_subjectSortType != 1)))
		return FALSE;

	if (pFrame->m_NamePatternParams.m_bPrintToSeparatePDFFiles == TRUE)
		return FALSE;
	else
	{
		if ((pFrame->m_NamePatternParams.m_bAddBreakPageAfterEachMailInPDF == FALSE) && 
			(pFrame->m_NamePatternParams.m_bAddBreakPageAfterEachMailConversationThreadInPDF == FALSE))
			return FALSE;

		if (pFrame->m_NamePatternParams.m_bAddBreakPageAfterEachMailInPDF)
			return TRUE;
	}

	if (nextMailPosition >= s_mails.GetCount())
		return FALSE;

	BOOL addPageBreak = FALSE;
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
	return addPageBreak;
}

int RemoveString(char *data, int datalen, CString &str, SimpleString *outbuf)
{
	BOOL m_bCaseSens = FALSE;
	char *p = data;
	int len = datalen;

	int pos = 0;
	while (pos >= 0)
	{
		int pos = g_tu.BMHSearch((unsigned char *)p, len, (unsigned char *)(LPCSTR)str, str.GetLength(), m_bCaseSens);
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

int MboxMail::printSingleMailToHtmlFile(/*out*/CFile &fp, int mailPosition, /*in - mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig, bool singleMail, BOOL addPageBreak)
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

	workbuf->ClearAndResize(outbuflarge->Count() * 2);
#if 0
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
	if (pFrame)
	{
		NListView *pListView = pFrame->GetListView();
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

	CString htmlHdrFldNameStyle;
	CString fldNameFontStyle;
	CString fldTextFontStyle;
	if (pFrame)
	{
		CreateFldFontStyle(pFrame->m_HdrFldConfig, fldNameFontStyle, fldTextFontStyle);
		htmlHdrFldNameStyle.Format("<style>\r\n.hdrfldname{%s}\r\n.hdrfldtext{%s}\r\n</style>", fldNameFontStyle, fldTextFontStyle);
	}

	CString bdyy;
	bdyy.Append("\r\n\r\n<html><body style=\"background-color:#FFFFFF;\"><div></div></body></html>");
	bdyy.Append("<article style=\"float:left;position:left;width:100%;background-color:#FFFFFF;margin: 0mm 0mm 0mm 0mm;\">");
	bdyy.Append("<style>@page {size: auto; margin: 12mm 4mm 12mm 12mm;}; * {font-size:100% !important;};</style>");
	

	if (fontSizePDF == 0)
	{
		; // bdyy.Append("</style>");
	}

	else if (singleMail == false)
	{
		if (fontSizePDF == 16)
			bdyy.Append("<style> * {font-size:16px !important;}</style>");
		else if (fontSizePDF == 20)
			bdyy.Append("<style> * {font-size:20px !important;}</style>");
		else if (fontSizePDF == 24)
			bdyy.Append("<style> * {font-size:24px !important;}</style>");

	}


	CString margin = "margin:0mm 8mm 0mm 12mm;";
	margin = "margin-left:5px;";

	fp.Write(bdyy, bdyy.GetLength());

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


		// Remove color and width from body tag, add later
		CString bodyBackgroundColor;
		CString bodyWidth;
		// if (!singleMail)    // TODO: do it for single mail also
		{
			BOOL removeBackgroundColor = TRUE;
			BOOL removeWidth = TRUE;

			int retval = NListView::RemoveBodyBackgroundColorAndWidth((char*)outbuflarge->Data(), outbuflarge->Count(), workbuf,
				bodyBackgroundColor, bodyWidth, removeBackgroundColor, removeWidth);

			if (!bodyWidth.IsEmpty() || !bodyBackgroundColor.IsEmpty())
			{
				outbuflarge->Copy(*workbuf);
			}
		}
		//bdy = "\r\n<div style=\'width:100%;position:initial;float:left;background-color:transparent;" + margin + "text-align:left\'>\r\n";
		CString newBodyWidth = "width:100%;";
		CString newBodyMargin = margin;
		bdy = "\r\n<div style=\"position:initial;float:left;background-color:transparent;text-align:left;"
			+ newBodyWidth + newBodyMargin
			+ "\">\r\n";   // div #1

		fp.Write(bdy, bdy.GetLength());

		CString font = "\">\r\n";

		// Setp mail header
		CString hdrBackgroundColor = "background-color:#eee9e9;";
		if (pFrame && (pFrame->m_NamePatternParams.m_bAddBackgroundColorToMailHeader == FALSE))
			hdrBackgroundColor == "";

		CString hdrWidth = "width:100%;";
		CString hdrMargin = "";
		if (pFrame)
		{
			bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + ";\">"
				+ htmlHdrFldNameStyle + "</head>\r\n"
				+ "<body bgColor=#ffffff>\r\n<div style=\"position:initial;float:left;text-align:left;font-weight:normal;"   // div #2
				+ hdrWidth + hdrMargin + hdrBackgroundColor
				+ font;
			fp.Write(bdy, bdy.GetLength());
		}

		int ret;
		if (pFrame)
		{
			ret = printMailHeaderToHtmlFile(fp, mailPosition, fpm, textConfig, pFrame->m_HdrFldConfig);
		}

		bdy = "</div></body></html>";   // --> end of div #2
		fp.Write(bdy, bdy.GetLength());

		bdy = "\r\n</div>";   // --> end of div #1
		fp.Write(bdy, bdy.GetLength());

		// ---> Setup mail header done

		// General body setup, is this needed ??
		// if (!singleMail)  // do it for single and multiple mails
		bodyWidth = "width:99%;";
		if (1)
		{
			CString fontColor = "black";  // TODO: probably no harm done setting to black always
			CString bdyBackgroundColor;
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

			bdy = "\r\n<div style=\"position:initial;float:left;text-align:left;"
				+ bodyWidth + margin + bdyBackgroundColor
				+ "\"><br>\r\n";

			fp.Write(bdy, bdy.GetLength());
		}

		// --> Done with general Body Setup


		// Setup if no html and body tags
		bool extraHtmlHdr = false;
		if (outbuflarge->FindNoCase(0, "<html", 5) < 0) // didn't find if true
		{
			if (outbuflarge->FindNoCase(0, "<body", 5) < 0) // didn't find if true
			{
				extraHtmlHdr = true;
			}
		}

		bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + ";\"></head>";
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
	}
	else
	{
		outbuflarge->Clear();
		pageCode = 0;
		textType = 0; // !!!!!!!!  no Html, try to get Plain  !!!!!!
		textlen = GetMailBody_mboxview(fpm, mailPosition, outbuflarge, pageCode, textType);  // returns pageCode
		if (textlen != outbuflarge->Count())
			int deb = 1;

		CString bdy;

		CString bdycharset = "UTF-8";

		CString newBodyWidth = "width:100%;";
		CString newBodyMargin = margin;
		bdy = "\r\n<div style=\"position:initial;float:left;background-color:transparent;text-align:left;"
			+ newBodyWidth + newBodyMargin
			+ "\">\r\n";   // div #1

		fp.Write(bdy, bdy.GetLength());

		CString font = "\">\r\n";


		// Setp mail header
		CString hdrBackgroundColor = "background-color:#eee9e9;";
		if (pFrame && (pFrame->m_NamePatternParams.m_bAddBackgroundColorToMailHeader == FALSE))
			hdrBackgroundColor == "";

		CString hdrWidth = "width:100%;";
		CString hdrMargin = "";
		if (pFrame)
		{
			bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + ";\">"
				+ htmlHdrFldNameStyle + "</head>\r\n"
				+ "<body bgColor=#ffffff>\r\n<div style=\"position:initial;float:left;text-align:left;font-weight:normal;"   // div #2
				+ hdrWidth + hdrMargin + hdrBackgroundColor
				+ font;
			fp.Write(bdy, bdy.GetLength());
		}

		int ret;
		if (pFrame)
			ret = printMailHeaderToHtmlFile(fp, mailPosition, fpm, textConfig, pFrame->m_HdrFldConfig);

		bdy = "</div></body></html>";
		fp.Write(bdy, bdy.GetLength());

		bdy = "\r\n</div>";
		fp.Write(bdy, bdy.GetLength());

		CString bodyWidth = "width:99%;";
		CString bodyBackgroundColor = "background-color:#FFFFFF;";
		//bdy = "\r\n<div style=\'width:auto;position:initial;float:left;color:black;background-color:#FFFFFF;" + margin + "text-align:left\'>\r\n";
		bdy = "\r\n<div style=\"position:initial;float:left;text-align:left;color:black;"
			+ bodyWidth + margin + bodyBackgroundColor
			+ "\">\r\n";
		fp.Write(bdy, bdy.GetLength());

		bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body bgColor=#ffffff><br>\r\n";
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
	}

#if 0
	SimpleString htmlbuf(1024);
	CString attachmentFimenamePrefix;
	printAttachmentNamesAsHtml(&fpm, mailPosition, &htmlbuf, attachmentFimenamePrefix);

	fp.Write(htmlbuf.Data(), htmlbuf.Count());
#endif

	CString bdy;

	bdy = "\r\n</article>";
	fp.Write(bdy, bdy.GetLength());

	bdy = "\r\n<div>&nbsp;<br></div>";
	fp.Write(bdy, bdy.GetLength());

	if (addPageBreak)
	{
		bdy = "\r\n<div style=\"page-break-before:always\"></div>";
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
	for (int pos = 0; pos != txt.Count(); ++pos) {
		char c = txt.GetAt(pos);
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

int MboxMail::CreateFldFontStyle(HdrFldConfig &hdrFieldConfig, CString &fldNameFontStyle, CString &fldTextFontStyle)
{
	const int lineHeight = 120;
	if (hdrFieldConfig.m_bHdrFontDflt == 0)
	{
		int nameFontSize = hdrFieldConfig.m_nHdrFontSize;
		int textFontSize = hdrFieldConfig.m_nHdrFontSize;

		fldTextFontStyle.Format("color:black;font-size:%dpx;line-height:%d%%;", textFontSize, lineHeight);

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
			textFontStyle = "italic";

		fldNameFontStyle.Format("color:black;font-style:%s;font-weight:%s;font-size:%dpx;line-height:%d%%;font-family:\"%s\",%s;",
			nameFontStyle, nameFontWeight, nameFontSize, lineHeight, nameFont, nameFontFamily);

		fldTextFontStyle.Format("color:black;font-style:%s;font-weight:%s;font-size:%dpx;line-height:%d%%;font-family:\"%s\",%s;",
			textFontStyle, textFontWeight, textFontSize, lineHeight, textFont, textFontFamily);
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
			else 
			{
				CString lDateTime = tt.FormatGmtTm(format);
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
		int ret = MboxMail::printAttachmentNamesAsHtml(&fpm, mailPosition, &tmpbuf, attachmentFileNamePrefix);

		if (tmpbuf.Count() > 0)
		{
			//encodeTextAsHtml(tmpbuf);

			CString label;
			label.Format("ATTACHMENTS (%s):", attachmentFileNamePrefix);
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

BOOL MboxMail::GetLongestCachePath(CString &longestCachePath)
{
	CString rootCacheSubFolder = "AttachmentCache";
	rootCacheSubFolder = "PrintCache";

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
	CString rootPrintSubFolder = "PrintCache";
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
	mboxviewFilePath = datapath + mailFileName + ".mboxview";
	return TRUE;
}

BOOL MboxMail::GetFileCachePathWithoutExtension(CString &mboxFilePath, CString &mboxBaseFileCachePath)
{
	CString errorText;
	CString printCachePath;
	CString rootPrintSubFolder = "PrintCache";
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
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
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
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt = _T("Could not create \"") + textFile;
			txt += _T("\" file.\n");
			txt += exErrorStr;

			//TRACE(_T("%s\n"), txt);

			CFileStatus rStatus;
			BOOL ret = fp.GetStatus(rStatus);

			//errorText = txt;

			HWND h = NULL; // we don't have any window yet
			int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}

		CFile fpm;
		CFileException ExError2;
		if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError2))
		{
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError2);

			CString txt = _T("Could not open \"") + MboxMail::s_path;
			txt += _T("\" mail file.\n");
			txt += exErrorStr;

			//TRACE(_T("%s\n"), txt);
			//errorText = txt;

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

		CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
		BOOL printAttachments = TRUE;
		if (pFrame)
		{
			printAttachments = pFrame->m_HdrFldConfig.m_HdrFldList.IsFldSet(HdrFldList::HDR_FLD_ATTACHMENTS-1);
		}

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
					if (printAttachments)
					{
						attachmentDB.Clear();
						NListView::PrintMailAttachments(&fpm, i, attachmentDB);
					}
					int pos = printSingleMailToHtmlFile(fp, i, fpm, textConfig, singleMail, addPageBreak);
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
					if (printAttachments)
					{
						attachmentDB.Clear();
						NListView::PrintMailAttachments(&fpm, i, attachmentDB);
					}
					int pos = printSingleMailToHtmlFile(fp, i, fpm, textConfig, singleMail, addPageBreak);
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

	CFileException ExError;;
	CFile fp;
	if (!fp.Open(textFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &ExError))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = _T("Could not create \"") + textFile;
		txt += _T("\" file.\n");
		txt += exErrorStr;

		TRACE(_T("%s\n"), txt);

		CFileStatus rStatus;
		BOOL ret = fp.GetStatus(rStatus);

		errorText = txt;
		return -1;
	}

	CFile fpm;
	CFileException ExError2;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError2))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError2);

		CString txt = _T("Could not open \"") + MboxMail::s_path;
		txt += _T("\" mail file.\n");
		txt += exErrorStr;

		TRACE(_T("%s\n"), txt);

		errorText = txt;
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

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	BOOL printAttachments = TRUE;
	if (pFrame)
	{
		printAttachments = pFrame->m_HdrFldConfig.m_HdrFldList.IsFldSet(HdrFldList::HDR_FLD_ATTACHMENTS-1);
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
			if (printAttachments)
			{
				attachmentDB.Clear();
				NListView::PrintMailAttachments(&fpm, i, attachmentDB);
			}
			int pos = printSingleMailToHtmlFile(fp, i, fpm, textConfig, singleMail, addPageBreak);
			if (pos < 0)
				continue;
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
		errorText = _T("Please open mail file first.");
		return -1;
	}

	CFileException ExError;
	CFile fp;
	if (!fp.Open(csvFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &ExError))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = _T("Could not create \"") + csvFile;
		txt += _T("\" file.\n");
		txt += exErrorStr;

		TRACE(_T("%s\n"), txt);

		CFileStatus rStatus;
		BOOL ret = fp.GetStatus(rStatus);

		errorText = txt;
		return -1;
	}

	CFile fpm;
	BOOL isFpmOpen = FALSE;
	//if (csvConfig.m_bContent) // open regardless; required by 
	{
		CFileException ExError;
		if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
		{
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt = _T("Could not open \"") + MboxMail::s_path;
			txt += _T("\" mail file.\n");
			txt += exErrorStr;

			TRACE(_T("%s\n"), txt);

			errorText = txt;
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
		int cnt = (int)selectedMailIndexList->GetCount();

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
	// if (csvConfig.m_bContent)  // was open if we are here
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
	ASSERT(outbuf != inbuf);;

	inbuf->ClearAndResize(10000);

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
	MboxMail *m = MboxMail::s_mails[mailPosition];
	int ret = MboxMail::GetMailBody_mboxview(fpm, m, outbuf, pageCode, textType);
	return ret;
}

int MboxMail::GetMailBody_mboxview(CFile &fpm, MboxMail *m, SimpleString *outbuf, UINT &pageCode, int textType)
{
	_int64 fileOffset;
	int bodyCnt = 0;

	// We are using global buffers so check and assert if we collide. 
	SimpleString *inbuf = MboxMail::m_inbuf;
	ASSERT(outbuf != inbuf);

	outbuf->ClearAndResize(10000);
	inbuf->ClearAndResize(10000);

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

	FileUtils::RemoveDirW(FileUtils::GetMboxviewTempPathW());

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
		CFile fp;
		CFileException ExError;
		if (fp.Open(fileName, CFile::modeWrite | CFile::modeCreate, &ExError))
		{
			fp.Write(data, dataLen);
			fp.Close();
		}
		else
		{
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt = _T("Could not create \"") + body->m_attachmentName;
			txt += _T("\" file.\n");
			txt += exErrorStr;

			TRACE(_T("%s\n"), txt);
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
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = _T("Could not create \"") + csvFile;
		txt += _T("\" file.\n");
		txt += exErrorStr;

		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CFile fpm;
	CFileException ExError2;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError2))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError2);

		CString txt = _T("Could not open \"") + MboxMail::s_path;
		txt += _T("\" mail file.\n");
		txt += exErrorStr;

		TRACE(_T("%s\n"), txt);

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

	if (updateRegistry)
	{
		MboxMail::m_HintConfig.SaveToRegistry();

		CString processpath = "";
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("processPath"), processpath);
	}
}

void ShellExecuteError2Text(UINT_PTR errorCode, CString &errorText) {
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
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = _T("Could not create \"") + statsFile;
		txt += _T("\" file.\n");
		txt += exErrorStr;

		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CFile fpm;
	CFileException ExError2;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError2))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError2);

		CString txt = _T("Could not open \"") + MboxMail::s_path;
		txt += _T("\" mail file.\n");
		txt += exErrorStr;

		TRACE(_T("%s\n"), txt);

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

	CString crc32;
	//crc32.Format("%X", m->m_crc32);
	crc32.Format("%07d", m->m_index);

	int fileLength = fileName.GetLength();
	if ((fileLength + crc32.GetLength()) > namePatternParams->m_nFileNameFormatSizeLimit)
	{
		int deb = 1;
	}

	if (separatorNeeded) {
		fileName.AppendChar(sepchar);
		int deb = 1;
	}

	fileName.Append((LPCSTR)crc32, crc32.GetLength());

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


	fileName.Trim(" \t_");
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

	int fileNameLength = fileName.GetLength();
	if (fileNameLength > namePatternParams->m_nFileNameFormatSizeLimit)
	{
		fileName.Truncate(namePatternParams->m_nFileNameFormatSizeLimit);
	}

	BOOL bReplaceWhiteWithUnderscore = FALSE;
	FileUtils::MakeValidFileName(fileName, bReplaceWhiteWithUnderscore);

	fileName.Trim(" \t_");

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
	
	CString toAddr;
	CString toName;

	SimpleString addrTo;
	SimpleString nameTo;

	CString addrStr;
	CString nameStr;

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
	uID.Format("%07d", m->m_index);
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

			int maxNumbOfAddr = 1;
			NListView::TrimToAddr(&m->m_to, addrStr, maxNumbOfAddr);
			NListView::TrimToName(&m->m_to, nameStr, maxNumbOfAddr);

			name.Append(nameStr, nameStr.GetLength());
			addr.Append(addrStr, addrStr.GetLength());

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
	fileName.Trim(" \t_");

	if (fileName.GetLength() > maxFileNameLength)
	{
		fileName.Truncate(maxFileNameLength);
	}

	FileUtils::MakeValidFileName(fileName, nameTemplateCnf.m_bReplaceWhiteWithUnderscore);

	fileName.Trim(" \t_");

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


	int left = IntPtr2Int(e - p);
	int fldCnt = (int)FldLabelTable.GetCount();
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


// global s_mails is the source array, s_mails_array is the destination work array
int MboxMail::RemoveDuplicateMails_Generic(MailArray &s_mails_array)
{
	MboxMail *m;
	MboxMail *m_found;

	// !! s_mails is already populated from mboxview file

	DWORD tc_start = GetTickCount();

	// Remove duplicate mails
	if (s_mails.GetCount() > s_mails_array.GetCount())
	{
		s_mails_array.SetSizeKeepData(s_mails.GetCount());
	}

	int i = 0;
	int to_dup_i = 0;
	int to_i = 0;

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
	}

	s_mails_array.SetSizeKeepData(to_i);

	int dupCnt = s_mails.GetSize() - s_mails_array.GetSize();

	s_mails.CopyKeepData(s_mails_array);

	DWORD tc_clear_start = GetTickCount();

	clearMboxMailTable();

	DWORD tc_clear_end = GetTickCount();
	DWORD delta_clear = tc_clear_end - tc_clear_start;

	DWORD tc_end = GetTickCount();
	DWORD delta = tc_end - tc_start;

	TRACE("RemoveDuplicateMails_Generic: total time=%d !!!!!!!!!!!!.\n", delta);
	TRACE("RemoveDuplicateMails_Generic: clear time=%d !!!!!!!!!!!!.\n", delta_clear);

	return dupCnt;
}

int MboxMail::RemoveDuplicateMails(MailArray &s_mails_array)
{
	MboxMail *m;
	MboxMail *m_found;

	// !! s_mails is already populated from mboxview file

	DWORD tc_start = GetTickCount();

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
			s_mails_find[to_dup_i] = s_mails[i];
			to_dup_i++;
			m->m_duplicateId = true;
		}
	}

	s_mails_array.SetSizeKeepData(to_i);

	s_mails_find.SetSizeKeepData(to_dup_i);
	MboxMail::m_findMails.m_lastSel = -1;
	MboxMail::m_findMails.b_mails_which_sorted = MboxMail::b_mails_which_sorted;

	int dupCnt = s_mails.GetSize() - s_mails_array.GetSize();

	s_mails.CopyKeepData(s_mails_array);

	// moved to NListView::OnRClickSingleSelect
	//MboxMail::m_editMails.m_bIsDirty = TRUE;

	DWORD tc_clear_start = GetTickCount();

	clearMboxMailTable();
	DWORD tc_clear_end = GetTickCount();
	DWORD delta_clear = tc_clear_end - tc_clear_start;

	DWORD tc_end = GetTickCount();
	DWORD delta = tc_end - tc_start;

	TRACE("RemoveDuplicateMails: total time=%d !!!!!!!!!!!!.\n", delta);
	TRACE("RemoveDuplicateMails: clear time=%d !!!!!!!!!!!!.\n", delta_clear);

	return dupCnt;
}

int MboxMail::LinkDuplicateMails(MailArray &s_mails_array)
{
	MboxMail *m;
	MboxMail *m_found;

	// !! s_mails is already populated from mboxview file

	DWORD tc_start = GetTickCount();

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

	DWORD tc_clear_start = GetTickCount();

	//clearMboxMailTable();
	DWORD tc_clear_end = GetTickCount();
	DWORD delta_clear = tc_clear_end - tc_clear_start;

	DWORD tc_end = GetTickCount();
	DWORD delta = tc_end - tc_start;

	TRACE("LinkDuplicateMails: total time=%d !!!!!!!!!!!!.\n", delta);
	TRACE("LinkDuplicateMails: clear time=%d !!!!!!!!!!!!.\n", delta_clear);

	return dupCnt;
}

//

#if 1

// Will increate MboxMail by 8 bytes
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

#else

MboxMail* MboxMail::getMboxMail(MboxMail *key)
{
	MboxMailTableType::iterator it;
	if (m_pMboxMailTable == 0)
		createMboxMailTable(HASH_ARRAY_SIZE);

	it = m_pMboxMailTable->find(key);
	if (it != m_pMboxMailTable->cend()) {
		return it->second;
	}
	else
		return 0;
}

bool MboxMail::insertMboxMail(MboxMail *key, MboxMail *mbox)
{
	if (m_pMboxMailTable == 0)
		createMboxMailTable(HASH_ARRAY_SIZE);

	std::pair<MboxMailTableType::iterator, bool> result = m_pMboxMailTable->insert(MboxMailTableType::value_type(key, mbox));
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
	if (m_pMboxMailTable)
		m_pMboxMailTable->clear();

	int deb = 1;
}

#endif

BOOL MboxMail::CreateCachePath(CString &rootPrintSubFolder, CString &targetPrintSubFolder, CString &prtCachePath, CString &errorText)
{
	BOOL ret = MboxMail::GetCachePath(rootPrintSubFolder, targetPrintSubFolder, prtCachePath, errorText);
	if (ret == FALSE)
		return FALSE;

	BOOL createDirOk = TRUE;
	if (!FileUtils::PathDirExists(prtCachePath)) {
		createDirOk = FileUtils::CreateDirectory(prtCachePath);
	}

	if (!createDirOk) {
		errorText = _T("Could not create \"") + prtCachePath;
		errorText += _T("\" folder for print destination.\nResolve the problem and try again.");
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
		errorText = _T("Please open mail archive file first.");
		return FALSE;
	}

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;
	FileUtils::SplitFilePath(mailArchiveFilePath, driveName, directory, fileNameBase, fileNameExtention);

	CString mailArchiveFileName;
	FileUtils::CPathStripPath((char*)(LPCSTR)mailArchiveFilePath, mailArchiveFileName);

	CString baseFileArchiveName;
	// TODO: verify and enable below
	//FileUtils::GetFileBaseName(mailArchiveFileName, baseFileArchiveName);

	int position = mailArchiveFileName.ReverseFind('.');
	if (position >= 0)
		baseFileArchiveName = mailArchiveFileName.Mid(0, position);
	else
		baseFileArchiveName = mailArchiveFileName;

	CString printCachePath = MboxMail::GetLastDataPath();
	//printCachePath.Append("\\");
	printCachePath.Append(rootPrintSubFolder);

	printCachePath.Append("\\");
	printCachePath.Append(baseFileArchiveName);
	//printCachePath.Append(mailArchiveFileName);
	printCachePath.TrimRight("\\");
	fileNameExtention.TrimLeft(".");
	printCachePath.Append("-");
	printCachePath.Append(fileNameExtention);

	if (!targetPrintSubFolder.IsEmpty())
	{
		printCachePath.Append("\\");
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
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = _T("Could not create \"") + textFile;
		txt += _T("\" file.\n");
		txt += exErrorStr;

		errorText = txt;
		// TODO:
		//return -1;
	}

	CFile fpm;
	CFileException ExError2;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError2))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError2);

		CString txt = _T("Could not open \"") + MboxMail::s_path;
		txt += _T("\" mail file.");
		txt += exErrorStr;

		TRACE(_T("%s\n"), txt);

		errorText = txt;

		fp.Close();
		return -1;
	}

	if (textType == 0) {
		if (textConfig.m_nCodePageId == CP_UTF8) {
			const char *BOM_UTF8 = "\xEF\xBB\xBF";
			fp.Write(BOM_UTF8, 3);
		}
	}

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	BOOL printAttachments = TRUE;
	if (pFrame)
	{
		printAttachments = pFrame->m_HdrFldConfig.m_HdrFldList.IsFldSet(HdrFldList::HDR_FLD_ATTACHMENTS-1);
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
			if (printAttachments)
			{
				attachmentDB.Clear();
				NListView::PrintMailAttachments(&fpm, i, attachmentDB);
			}
			int pos = printSingleMailToHtmlFile(fp, i, fpm, textConfig, singleMail, addPageBreak);
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
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = _T("Could not create file \"") + textFile;
		txt += _T("\" file.\n");
		txt += exErrorStr;

		TRACE(_T("%s\n"), txt);

		errorText = txt;
		return -1;
	}

	CFile fpm;
	CFileException ExError2;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError2))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = _T("Could not open \"") + MboxMail::s_path;
		txt += _T("\" mail file.\n");
		txt += exErrorStr;

		TRACE(_T("%s\n"), txt);

		errorText = txt;
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

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	BOOL printAttachments = TRUE;
	if (pFrame)
	{
		printAttachments = pFrame->m_HdrFldConfig.m_HdrFldList.IsFldSet(HdrFldList::HDR_FLD_ATTACHMENTS-1);
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
			if (printAttachments)
			{
				attachmentDB.Clear();
				NListView::PrintMailAttachments(&fpm, i, attachmentDB);
			}
			int pos = printSingleMailToHtmlFile(fp, i, fpm, textConfig, singleMail, addPageBreak);
			if (pos < 0)
				continue;
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

	if ((textType == 1) && (singleMail == FALSE))
		printDummyMailToHtmlFile(fp);

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

	CFileException ExError;
	if (!fp.Open(textFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &ExError))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = _T("Could not create \"") + textFile;
		txt += _T("\" file.\n");
		txt += exErrorStr;

		errorText = txt;
		return -1;
	}

	CFile fpm;
	CFileException ExError2;
	if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError2))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError2);

		CString txt = _T("Could not open \"") + MboxMail::s_path;
		txt += _T("\" mail file.\n");
		txt += exErrorStr;

		TRACE(_T("%s\n"), txt);

		errorText = txt;

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

	BOOL printAttachments = TRUE;
	if (pFrame)
	{
		printAttachments = pFrame->m_HdrFldConfig.m_HdrFldList.IsFldSet(HdrFldList::HDR_FLD_ATTACHMENTS-1);
	}

	AttachmentMgr attachmentDB;
	int i;
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
			if (printAttachments)
			{
				attachmentDB.Clear();
				NListView::PrintMailAttachments(&fpm, i, attachmentDB);
			}
			int pos = printSingleMailToHtmlFile(fp, i, fpm, textConfig, singleMail, addPageBreak);
			if (pos < 0)
				continue;
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

	if ((textType == 1) && (singleMail == FALSE))
		printDummyMailToHtmlFile(fp);
	
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
	int deb = 1;
	//ASSERT(0);
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

void HintConfig::ClearAllHints()
{
	m_nHintBitmap = 0xFFFFFFFF;
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
	static int tmlen = istrlen(tm);

	BOOL r = p < end;

	CString line;
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
				"Left-click on one of the loaded mbox file to view all associated mails.\n"
				"\n"
				"Please review the User Guide provided with the package\n"
				"and/or single and double right-click/left-click on any item\n"
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
				"Microsoft Edge Browser is used by default to print to PDF file.\n"
				"You can configure Edge browser to optionally print page header and footer\n\n"
				"If Edge Browser is not installed on your computer\n"
				"Chrome Browser can be installed to print to PDF file.\n"
				"However, the header and footer are always printed.\n "
				"\n"
				"You can also select \"File->Print Config->User Defined Script\" to "
				"use free wkhtmltopdf tool to remove header and footer\n"
				"and evalute whether wkhtmltopdf works for you."
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
				"printing. Both MS Edge and Chrome browsers don't support\n"
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
		else if (hintNumber == HintConfig::MessageHeaderConfigHint)
		{
			hintText.Append(
				"By default the mail text is shown in Message Window.\n"
				"Enable global \"View -> View Message Headers\" option to show the message header instead of the text.\n\n"
				"To switch between text and headers of the selected mail\n"
				"right-click on the header pane in Message Window and check/uncheck the \"View Message Header\" option.\n"
				"\n"
			);
		}
		else if (hintNumber == HintConfig::MessageRemoveFolderHint)
		{
			hintText.Append(
				"The selected folder will be removed from the current view.\n"
				"The physical folder and content on the disk will not be deleted.\n"
				"\n"
			);
		}
		else if (hintNumber == HintConfig::MessageRemoveFileHint)
		{
			hintText.Append(
				"The selected file will be removed from the current view.\n"
				"The physical file and content on the disk will not be deleted.\n"
				"\n"
			);
		}
		else if (hintNumber == HintConfig::MergeFilesHint)
		{
			hintText.Append(
				"Mbox and eml mail files can also be merged via two command line options:\n\n"
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
				"Gmail Labels are supported by performing a separate step on a active mail folder.\n\n"
				"Right-click on the folder and select \"File->Gmail Labels->Create\" option.\n\n"
				"\n"
			);
		}
		else if (hintNumber == HintConfig::SubjectSortingHint)
		{
			hintText.Append(
				"Sorting by subject creates subject threads. Emails within each subject thread are sorted by time.\n\n"
				"By default subject threads are sorted alphanumerically. "
				"Subject threads can be sorted by time by selecting  \"File->Options->time ordered threads\" option.\n"
				"\n"
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
		}
		MboxMail::m_HintConfig.ClearHint(hintNumber);
	}
}

// Fix static buffer allocation, it is bad !!!

static void LargeBufferAllocFailed()
{
	_ASSERT(0);
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
	static const char *cFromMailBegin = "From ";
	static const int cFromMailBeginLen = istrlen(cFromMailBegin);

	CString progDir;
	BOOL retDir = GetProgramDir(progDir);

	int seNumb = 0;
	CString szCause = seDescription(seNumb);

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
		//TRACE("offset = %lld\n", m_startOff);
		fp.Seek(msgOffset, SEEK_SET);
		UINT readLength = fp.Read(res->Data(), bytes2Read);
		if (readLength != bytes2Read)
			int deb = 1;
		res->SetCount(bytes2Read);
		fp.Close();
	}
	else
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = _T("Could not open \"") + MboxMail::s_path;
		txt += _T("\" mail file.\n");
		txt += exErrorStr;

		TRACE(_T("%s\n"), txt);
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

	char *se_mailDumpFileName = "UnhandledExeption_MailDump.txt";
	BOOL retW = DumpMailData(se_mailDumpFileName, szCause, seNumb, mailPosition, data, datalen);

	CString errorTxt;

	errorTxt.Format(_T("Unhandled_Exception: Code=%8.8x Description=%s\n\n"
		"To help to diagnose the problem, created files\n\n%s\n\nin\n\n%s directory.\n\n"
		"Please provide the files to the development team.\n\n"
		"%s file contains mail data so please review the content before forwarding to the development."),
		seNumb, szCause, se_mailDumpFileName, progDir, se_mailDumpFileName);

	AfxMessageBox((LPCTSTR)errorTxt, MB_OK | MB_ICONHAND);
}

void MboxMail::mbassert()
{
	if (!s_path.IsEmpty() && s_datapath.IsEmpty())
		int deb = 1;
	if (s_path.Compare(s_datapath))
		int deb = 1;
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


