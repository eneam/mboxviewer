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


#if !defined(_MBOX_MAIL_)
#define _MBOX_MAIL_

#pragma once

#include <vector>
#include "UPDialog.h"
#include "Mime.h"
#include "MimeCode.h"
#include <unordered_map>
#include <algorithm>
#include "dllist.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

_int64 FileSize(LPCSTR fileName);
void ShellExecuteError2Text(UINT errorCode, CString &errorText);
BOOL CodePage2WStr(SimpleString *str, UINT inCodePage, SimpleString *wstr);
BOOL WStr2CodePage(wchar_t *wbuff, int wlen, UINT outCodePage, SimpleString *result);
int findNoCase(const char *input, int count, void const* Str, int  Size);
void ShowMemStatus();

class MboxMail;
struct NamePatternParams;

typedef CArray<int, int> MailIndexList;

class MyMailArray : public CArray<MboxMail*, MboxMail*>
{
public:
	// Kludge to prevent CArray from deleting data buffer
	// Otherwise will need to reimplement array
	void SetSizeKeepData(INT_PTR nNewSize, INT_PTR nGrowBy = -1);
	void CopyKeepData(const MyMailArray& src);
};

typedef MyMailArray MailArray;

class HintConfig
{
public:
	HintConfig();
	~HintConfig();

	DWORD m_nHintBitmap;

	void LoadHintBitmap();
	void SaveToRegistry();

	BOOL IsHintSet(int nintNumber);
	void ClearHint(int nintNumber);

	enum {
		GeneralUsageHint = 1,
		MsgWindowPlacementHint = 2,
		PrintToPDFHint = 3,
		PrintToPrinterHint = 4,
		MailSelectionHint = 5,
		FindDialogHint,
		AdvancedFindDialogHint
	};
};

class MailList
{
public:
	MailList(int nId);
	MailList();
	~MailList();

	int m_nId;
	//MailArray s_mails;
	int b_mails_which_sorted;  // column to sort
	int m_lastSel;  // position of last selected mail
	//BOOL m_bIsDirty;  // content from s_mails was touched and needs to be copied to this mail list
};

struct TEXTFILE_CONFIG
{
public:
	BOOL m_bFrom;
	BOOL m_bTo;
	BOOL m_bSubject;
	BOOL m_bDate;
	BOOL m_bContent;
	CString m_dateFormat;
	int m_bGMTTime;
	int m_nCodePageId;
	CString m_separator;
	CString errorText;
};

struct PRINT_MAIL_ARCHIVE_ARGS
{
	TEXTFILE_CONFIG textConfig;
	CString textFile;
	int firstMail;
	int lastMail;
	int textType;
	BOOL terminated;
	BOOL exitted;
	CString errorText;
};

struct PRINT_MAIL_ARCHIVE_TO_CSV_ARGS
{
	CSVFILE_CONFIG csvConfig;
	CString csvFile;
	int firstMail;
	int lastMail;
	MailIndexList *selectedMailIndexList;
	BOOL terminated;
	BOOL exitted;
	CString errorText;
};

class MboxCMimeHelper
{
public:
	static void GetContentType(CMimeBody* pBP, CString &value);
	static void GetTransferEncoding(CMimeBody* pBP, CString &value);
	static void GetContentID(CMimeBody* pBP, CString &value);
	static void GetContentDescription(CMimeBody* pBP, CString &value);
	static void GetContentDisposition(CMimeBody* pBP, CString &value);
	static void GetCharset(CMimeBody* pBP, CString &value);
	static void Name(CMimeBody* pBP, CString &value);
	static void Filename(CMimeBody* pBP, CString &value);
protected:
	static void GetValue(CMimeBody* pBP, const char* fieldName, CString &value);
};


class MboxCMimeCodeBase64 : public CMimeCodeBase64
{
public:
	MboxCMimeCodeBase64(const char* pbInput, int nInputSize) {
		SetInput(pbInput, nInputSize, false);
	}
};

class MboxCMimeCodeQP : public CMimeCodeQP
{
public:
	MboxCMimeCodeQP(const char* pbInput, int nInputSize) {
		SetInput(pbInput, nInputSize, false);
	}
};

// Investigate and use standard CString instead  ?
class SimpleString
{
public:
#define DFLT_GROW_SIZE 16
	SimpleString() { Initialize(0, DFLT_GROW_SIZE);  }
	SimpleString(int capacity) { Initialize(capacity, DFLT_GROW_SIZE); }
	SimpleString(int capacity, int grow_size) { Initialize(capacity, grow_size); }

	void Initialize(int capacity, int grow_size) {
		m_data = 0;  m_count = 0;  m_grow_size = grow_size;
		m_capacity = capacity;
		if (m_capacity < DFLT_GROW_SIZE)
			m_capacity = DFLT_GROW_SIZE;
		m_data = new char[m_capacity + 1];  // extra byte for NULL
		if (m_data)
			m_data[0] = 0;
		else
			m_capacity = 0;  // this would be a problem. introduce m_error ?
	};
	~SimpleString() { delete[] m_data; m_data = 0; m_count = 0; m_capacity = 0; };
	void Release() { delete[] m_data; m_data = 0; m_count = 0;  m_capacity = 0; };

	char *m_data;
	int m_capacity;
	int m_count;
	int m_grow_size;

	char *Data() { return m_data;  }
	char *Data(int pos) { return &m_data[pos]; }  // zero based pos
	int Capacity() { return m_capacity; }
	int Count() { return m_count; }
	void SetCount(int count) {
		ASSERT(count <= m_capacity); ASSERT(m_data);
		m_count = count; m_data[count] = 0;
	}
	void Clear() {
		SetCount(0);
	}
	bool CanAdd(int characters) {
		if ((m_count + characters) <= m_capacity) 
			return true; else  return false;
	}
	int Resize(int size);

	int ClearAndResize(int size) {
		int new_size = Resize(size);
		SetCount(0);
		return new_size;
	}

	void Copy(char c) {
		if (1 > m_capacity) // should never be true
			Resize(1);
		m_data[0] = c;
		SetCount(1);
	}

	void Append(char c) {
		if (!CanAdd(1)) Resize(m_capacity + 1);
		m_data[m_count++] = c;
		m_data[m_count] = 0;
	}

	void Copy(void const* Src, size_t  Size) {
		if (Size > m_capacity) Resize(Size);
		::memcpy(m_data, Src, Size);
		SetCount(Size);
	}

	void append_internal(void const* Src, size_t  Size);

	void Append(void const* Src, size_t  Size) {
		if (!CanAdd(Size))
			Resize(m_count + Size);
		append_internal(Src, Size);
	}

	void Copy(register char *src) {
		int slen = strlen(src);
		SimpleString::Copy(src, slen);
	}

	void Append(register char *src) {
		int slen = strlen(src);
		SimpleString::Append(src, slen);
	}

	void Copy(SimpleString &str) { 
		Copy(str.Data(), str.Count());
	}

	void Append(SimpleString &str) {
		Append(str.Data(), str.Count());
	}

	int FindNoCase(int offset, void const* Src, int  Size);
	int FindAny(int offset, void const * Src);
	int Find(int offset, char const c);

	char GetAt(int pos) {
		return m_data[pos];
	}
};

class MailBodyContent
{
public:
	MailBodyContent() { m_pageCode = 0; m_contentOffset = 0; m_contentLength = 0;  };
	~MailBodyContent() {};
	CString m_contentType;
	CString m_contentTransferEncoding;
	CString m_contentDisposition;
	CString m_contentId;
	CString m_attachmentName;
	UINT m_pageCode;
	int  m_contentOffset;
	int m_contentLength;
};

unsigned long StrHash(const char* buf, const UINT length);

struct MessageIdHash;
struct MessageIdEqual;

typedef std::unordered_map<CString*, int, MessageIdHash, MessageIdEqual> MessageIdTableType;
typedef std::unordered_map<CString*, MboxMail*, MessageIdHash, MessageIdEqual> MboxMailTableType;


class CMBodyHdr;

#define White  RGB(255,255,255)
#define Black  RGB(0,0,0)

#define PeachPuff1 	RGB(255,218,185)
#define PeachPuff2 	RGB(238,203,173)
#define PeachPuff3 	RGB(205,175,149)

#define SandyBrown RGB(244,164,96)
#define Burlywood  RGB(222,184,135)

#define AntiqueWhite2 RGB(238,223,204)
#define AntiqueWhite3 RGB(205,192,176)


class MboxMail
{
public:
	MboxMail() 
	{
		m_startOff = m_length = m_hasAttachments = m_headLength = 0;
		m_from_charsetId = m_to_charsetId = m_subj_charsetId = 0;
		m_cc_charsetId = m_bcc_charsetId = 0;
		m_timeDate = 0;
		m_recv = 0;
		m_groupId = -1;
		m_nextMail = -1;
		m_prevMail = -1;
		m_duplicateId = false;
		m_done = false;
		m_groupColor = 0;
		m_index = -1;
		m_headLength = 0;
		m_isOnUserSelectedMailList = false;
		//m_crc32 = 0xffffffff;
	}

	std::vector <MailBodyContent*> m_ContentDetailsArray;

	//UINT32 m_crc32;
	_int64 m_startOff;
	int m_hasAttachments;
	int m_length, m_headLength, m_recv;
	time_t m_timeDate;
	CString m_from, m_to, m_subj;
	//CString m_from_name, m_to_name;  // TODO: should we precalculate ? or calculate run time ?
	CString m_cc, m_bcc;
	CString m_from_charset, m_to_charset, m_subj_charset;
	CString m_cc_charset, m_bcc_charset;
	UINT m_from_charsetId, m_to_charsetId, m_subj_charsetId;
	UINT m_cc_charsetId, m_bcc_charsetId;
	CString m_messageId;
	CString m_replyId;
	int m_groupId;
	int m_groupColor;
	int m_nextMail;
	int m_prevMail;
	bool m_duplicateId;
	bool m_done;
	int m_index; // servers as unique id == index
	bool m_isOnUserSelectedMailList;

	BOOL GetBody(CString &str);
	BOOL GetBody(SimpleString *str);
	int DumpMailBox(MboxMail *mailBox, int which);

	static MessageIdTableType *m_pMessageIdTable;
	static MboxMailTableType *m_pMboxMailTable;
	//
	static UINT MboxMail::createMessageIdTable(UINT count);
	static void MboxMail::clearMessageIdTable();
	static int getMessageId(CString *key);
	static bool insertMessageId(CString *key, int val);
	static int getReplyId(CString *key);
	static int add2ConversationGroup(int mid, MboxMail *m);

	static MboxMail* getMboxMail(CString *key);
	static bool insertMboxMail(CString *key, MboxMail *mbox);
	static UINT createMboxMailTable(UINT count);
	static void clearMboxMailTable();
	//
	static int m_nextGroupId;
	// Tricky to use to avoid ownership conflict when two function on stack use the same buffer; Asking for trouble :)
	static SimpleString *m_outbuf;
	static SimpleString *m_inbuf;
	static SimpleString *m_outdata;
	static SimpleString *m_indata;
	static SimpleString *m_workbuf;
	static SimpleString *m_tmpbuf;
	static SimpleString *m_largebuf;

	static int m_Html2TextCount;

	static _int64 s_fSize; // current File size
	static _int64 s_oSize; // old file size
	static CString s_path;
	static _int64 s_curmap, s_step;
	static const CUPDUPDATA* pCUPDUPData;
	static void Str2Ansi(CString &res, UINT CodePage);
	static UINT Str2PageCode(const  char* PageCodeStr);
	static void Parse(LPCSTR path);
	static bool Process(char *p, DWORD size, _int64 startOffset, bool bFirstView, bool bLastView, _int64 &lastStartOffset, bool bEml = false);
	

	// TODO:  need to rewrite, encapsulate in objects, etc
	static MailArray s_mails_ref;  // original cache
	static MailArray s_mails;  // other lists need to be copied here since that is how initial implementation works, would need to change in too many places
	static MailArray s_mails_all;
	static MailArray s_mails_find;
	static MailArray s_mails_edit;  // TODO: rename to User Selected List
	//
	static MailArray s_mails_selected;
	static MailArray s_mails_merged;
	static int nWhichMailList;

	static MailList *m_mailList;  // must match nWhichMailList
	static MailList m_allMails;
	static MailList m_findMails;
	static MailList m_editMails;

	static BOOL IsAllMailsSelected();
	static BOOL IsFindMailsSelected();
	static BOOL IsUserMailsSelected();
	static int AllMailsSelectedId();
	static int FindMailsSelectedId();
	static int UserMailsSelectedId();

	static bool b_mails_sorted;
	static int b_mails_which_sorted;

	static void SortByDate(MailArray *s_mails = 0, bool bDesc = false);
	static void SortByFrom(MailArray *s_mails = 0, bool bDesc = false);
	static void SortByTo(MailArray *s_mails = 0, bool bDesc = false);
	static void SortBySubject(MailArray *s_mails = 0, bool bDesc = false);
	static void SortBySize(MailArray *s_mails = 0, bool bDesc = false);
	static void SortByGroupId(MailArray *s_mails = 0, bool bDesc = false);
	static void SortByFileOffset(MailArray *s_mails = 0, bool bDesc = false);
	static void SortByConverstionGroups(MailArray *s_mails = 0, bool bDesc = false);
	static void SortByIndex(MailArray *s_mails = 0, bool bDesc = false);
	static void assignColor2ConvesationGroups();
	static void assignColor2ConvesationGroups(MailArray *mails);
	static void Destroy();
	static bool preprocessConversations();
	static bool sortConversations();
	static bool validateSortConversations();
	static int charCount(char *fld, char c);
	static int nstrcpy(char *dst, char *src);
	static int escapeSeparators(char *workbuff, char *fldstr, int fldlen, char sepchar);
	static int splitMailAddress(const char *buff, int bufflen, SimpleString *name, SimpleString *addr);
	static CString DecodeString(CString &subj, CString &charset, UINT &charsetId, UINT toCharacterId = 0);
	static void EncodeAsHtml(const char *in, int inLength, SimpleString *out);
	static int EnforceCharacterLimit(SimpleString *buffer, CString &characterLimit);
	static int EnforceFieldTextCharacterLimit(char *buffer, int bufferLength, CString &characterLimit);
	//
	static int printMailHeaderToHtmlFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig);
	static int printSingleMailToHtmlFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig, bool firstMail);
	static int printSingleMailToTextFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig);
	static int printMailHeaderToTextFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig);
	static int exportToTextFile(TEXTFILE_CONFIG &textConfig, CString &textFileName, int firstMail, int lastMail, MailIndexList *selectedMailIndexList, int textType, BOOL progressBar);
	//
	static int exportHeaderFieldLabelsToCSVFile(CSVFILE_CONFIG &csvConfig, CFile &fp);
	static int printMailHeaderToCSVFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, CSVFILE_CONFIG &csvConfig);
	static int exportToCSVFile(CSVFILE_CONFIG &csvConfig, CString &csvFileName, int firstMail, int lastMail, MailIndexList *selectedMailsIndexList, BOOL progressBar);
	static int printSingleMailToCSVFile(/*out*/ CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, CSVFILE_CONFIG &csvConfig, bool firstMail);
	static int exportToCSVFileFullMailParse(CSVFILE_CONFIG &csvConfig);  // NOT used currently
	//
	static int printMailArchiveToTextFile(TEXTFILE_CONFIG &textConfig, CString &textFileName, int firstMail, int lastMail, int textType, BOOL progessBar, CString &errorText);
	static int printMailArchiveToCSVFile(CSVFILE_CONFIG &csvConfig, CString &csvFile, int firstMail, int lastMail, MailIndexList *selectedMailIndexList, BOOL progressBar, CString &errorText);
	//
	static int PrintMailRangeToSingleTextFile(TEXTFILE_CONFIG &textConfig, CString &textFileName, int firstMail, int lastMail, int textType, CString &targetPrintSubFolderName, CString errorText);
	static int PrintMailRangeToSingleTextFile_WorkerThread(TEXTFILE_CONFIG &textConfig, CString &textFileName, int firstMail, int lastMail, int textType, CString errorText);
	//
	static int PrintMailSelectedToSingleTextFile_WorkerThread(TEXTFILE_CONFIG &textConfig, CString &textFileName, MailIndexList *selectedMailIndexList, int textType, CString errorText);
	//
	static int GetMailBody_mboxview(CFile &fpm, int mailPosition, SimpleString *outbuf, UINT &pageCode, int textMinorType = 0);  // 0 if text/plain, 1 if text/html
	static int GetMailBody_MailBody(CFile &fpm, int mailPosition, SimpleString *outbuf, UINT &pageCode);
	static int GetMailBody_CMimeMessage(CMimeMessage &mail, int mailPosition, SimpleString *outbuf, UINT &pageCode);
	static void getCMimeBodyHeader(CMimeMessage *mail, CMimeBody* pBP, CMBodyHdr *pHdr);
	static int LoadMail(const char* pszData, int nDataSize);
	static char *ParseContent(MboxMail *mail, char *startPos, char *endPos);
	static int CreateImgAttachmentFiles(CFile &fpm, int mailPosition, SimpleString *outbuf);
	static int DecodeBody(CFile &fpm, MailBodyContent *body, int mailPosition, SimpleString *outbuf);
	static int DumpMailStatsToFile(MailArray *mailsArray, int mailArrayCount);
	static int DetermineLimitedLength(SimpleString *str, int maxLinesTextLimit);
	static int MergeTwoMailLists(MailArray *mails1, MailArray *mails2, MailArray *merged_mails);
	static BOOL VerifyMergeOfTwoMailLists(MailArray *mails1, MailArray *mails2, MailArray *merged_mails);
	static BOOL Test_MergeTwoMailLists();
	static void PopulateCArray(MailArray *mails, int *carray, int carrayCnt);
	static BOOL VerifyMerge(MailArray *mails, int *carray, int carrayCnt);
	//static void ShellExecuteError2Text(UINT errorCode, CString errorText);
	static int MakeFileName(MboxMail *m, struct NamePatternParams *namePatternParams, CString &fileName);
	static CString GetDateFormat(int i);
	static BOOL GetPrintCachePath(CString &printCachePath);
	static int RemoveDuplicateMails();
	static void SplitFilePath(CString &fileName, CString &driveName, CString &directory, CString &fileNameBase, CString &fileNameExtention);
	static int MakeFileNameFromMailArchiveName(int fileType, CString &fileName, CString &targetPrintSubFolder, bool &fileExists, CString &errorText);
	static int MakeFileNameFromMailHeader(int mailIndex, int fileType, CString &fileName, CString &targetPrintSubFolder, bool &fileExists, CString &errorText);
	static BOOL CreatePrintCachePath(CString &rootPrintSubFolder, CString &targetPrintSubFolder, CString &prtCachePath, CString &errorText);
	static void UpdateFileExtension(CString &fileName, CString &newSuffix);
	//
	static bool GetPrintCachePath(CString &rootPrintSubFolder, CString &targetPrintSubFolder, CString &prtCachePath, CString &errorText);
	static void MakeValidFileName(CString &name);
	static void MakeValidFileName(SimpleString &name);

	//
	
	static void ReleaseResources();
	static void assert_unexpected();


	static HintConfig m_HintConfig;
	static void ShowHint(BOOL hintNumber);
	static void LoadHintBitmap();
};

#define SZBUFFSIZE 1024*1024*100

class SerializerHelper
{
private:
	BOOL m_writing;
	CString m_path;
	HANDLE m_hFile;
	int m_offset;
	char * m_buff;
	int m_buffSize;
public:
	SerializerHelper(LPCSTR fn) {
		m_path = fn;
		m_offset = 0;
		m_buff = NULL;
	}
	~SerializerHelper() {
		close();
	}
	void close();
	BOOL open(BOOL bWrite);
	int GetReadPointer();
	BOOL SetReadPointer(int pos);
	BOOL readN(void *v, int sz);
	BOOL writeN(void *v, int sz);
	BOOL writeString(LPCSTR val);
	BOOL readString(CString &val);

	BOOL writeInt(int val) {
		return writeN(&val, sizeof(int));
	}
	BOOL readInt(int *val) {
		return readN(val, sizeof(int));
	}
	BOOL readUInt(unsigned int *val) {
		return readN(val, sizeof(int));
	}
	BOOL writeInt64(_int64 value) {
		return writeN(&value, sizeof(_int64));
	}
	BOOL readInt64(_int64 *val) {
		return readN(val, sizeof(_int64));
	}
};


class MailHeader
{
public:
	MailHeader() { Clear(); }
	~MailHeader() { }
public:
	void Clear();
	char *EatNLine(char* p, char* e) { while (p < e && *p++ != '\n'); return p;  }
	
	int Load(const char* pszData, int nDataSize);
	bool IsMultiPart() { return m_IsMultiPart; }
	bool AssertHdr();

	bool m_IsText;
	bool m_IsTextPlain;
	bool m_IsTextHtml;
	bool m_IsMessage;
	bool m_IsAttachment;
	bool m_IsMultiPart;
	CString m_Charset;
	UINT m_PageCode;
	CString m_Description;
	CString m_Disposition;
	CString m_TransferEncoding;
	CString m_SubType;
	CString m_MainType;
	CString m_Boundary;
	CString m_ContentType;
	CString m_ContentId;
	CMimeHeader::MediaType m_MediaType;
	CString m_AttachmentName;
	CString m_MessageId;
	CString m_ReplyId;
};

class MailBodyPool;

class MailBody:public MailHeader
{
	friend class MailBodyPool;
public:
	MailBody() { m_bodyDataOffset = 0; };
	~MailBody() { DeleteAll(); };

	int Load(char *& pszDatabase, const char* pszData, int nDataSize);
	BOOL AssertData(MboxMail *mail);

	typedef list<MailBody*> MailBodyList;
	int GetBodyPartList(MailBodyList& rList);
	int m_bodyDataOffset;
	int m_bodyDataLength;

protected:
	char* FindBoundary(const char* pszData, const char* pszDataEnd, const char* boundary, int boundaryLength);
	const char* FindString(const char* pszStr1, const char* pszStr2, const char* pszEnd);

	DLLIST_NODE(MailBody) m_freeList_link;
	DLLIST(MailBody, m_freeList_link) m_listBodies;

	void DeleteAllParts();
	void DeleteAll();
	
	static MailBody *CreatePart();
	static void ErasePart(MailBody* body);

public:
	static MailBody *CreateMailBody();
	static void FreeMailBody(MailBody* body);
	static MailBodyPool *m_mpool;
};

class MailBodyPool
{
public:
	MailBodyPool() {};
	~MailBodyPool();

	MailBody *AllocPart();
	void FreePart(MailBody* body);

	DLLIST(MailBody, m_freeList_link) m_freeMailBodyList;
};

inline char* FixIfNull(const char* ptr)
{
	return ptr ? ptr : "null";
};

#endif // _MBOX_MAIL_