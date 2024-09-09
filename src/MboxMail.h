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
#include "SimpleString.h"
#include "HtmlPdfHdrConfigDlg.h"
#include "IHashTable.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

class SimpleString;
class StackWalker;
class MyStackWalker;

void ShellExecuteError2Text(UINT_PTR errorCode, CString &errorText);
void ShowMemStatus();

class Tracer
{
public:
	Tracer(const char *functionName) : functionName_(functionName)
	{
		TRACE("Entering function %s\n", functionName);
	}

	~Tracer()
	{
		TRACE("Exitting function %s\n", functionName_);
	}

	const char *functionName_;
};

#ifdef _DEBUG
#define DEBUG_TRACER
#ifdef DEBUG_TRACER
#define DBGT(f) Tracer(f);
#endif
#else
#define DBGT(f)
#endif

class MboxMail;
struct NamePatternParams;

typedef CArray<int, int> MailIndexList;

class ProgressTimer
{
public:
	// LONGLONG == _int64
	ProgressTimer(ULONGLONG workRangeFirstPos, ULONGLONG workRangeLastPos);
	BOOL UpdateWorkPos(ULONGLONG workRangeFirstPos, UINT_PTR& dwProgressbarPos, BOOL force = FALSE);

	int m_progressBarSize;
	double m_workStepSize;
	ULONGLONG m_checkNewWorkPosRate;

	ULONGLONG m_maxTimeValue;
	ULONGLONG m_startTime;
	ULONGLONG m_CurrentTime;
	//
	LONGLONG m_workUpdateCnt;
	ULONGLONG m_workRangeFirstPos;
	ULONGLONG m_workRangeLastPos;
	ULONGLONG m_workSize;
	ULONGLONG m_workLastPos;
	ULONGLONG m_workCurrentPos;
};

class MyMailArray : public CArray<MboxMail*, MboxMail*>
{
public:
	MyMailArray()
	{
		SetSize(1024);
		SetCountKeepData(0);
	}
	// Kludge to prevent CArray from deleting data buffer
	// Otherwise will need to reimplement array
	// This will set count of items. if nNewSize == 0, array data segment is not changed.
	// If nNewSize >  0, SetSize() is called and both count and count of items is set
	void SetSizeKeepData(INT_PTR nNewSize, INT_PTR nGrowBy = 1024);
	// This will set count of items and will not touch array data segment
	void SetCountKeepData(INT_PTR nNewSize, INT_PTR nGrowBy = 1024);
	void CopyKeepData(const MyMailArray& src);
	INT_PTR MaxSize();
	int GetCount() { return (int)CArray<MboxMail*, MboxMail*>::GetCount(); }
	int GetSize() { return (int)CArray<MboxMail*, MboxMail*>::GetSize(); }
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
	void ClearAllHints();

	enum {
		GeneralUsageHint = 1,
		MsgWindowPlacementHint = 2,
		PrintToPDFHint = 3,
		PrintToPrinterHint = 4,
		MailSelectionHint = 5,
		FindDialogHint = 6,
		AdvancedFindDialogHint = 7,
		PrintToPDFScriptHint = 8,
		AttachmentConfigHint = 9,
		MessageHeaderConfigHint = 10,
		MessageRemoveFolderHint = 11,
		MessageRemoveFileHint = 12,
		MergeFilesHint = 13,
		LabelsHint = 14,
		SubjectSortingHint = 15
	};
};

class MailList
{
public:
	MailList(int nId);
	MailList();
	~MailList();

	DLLIST_NODE(MailList) m_listLink;
	//DLLIST(MailBody, m_freeList_link) m_listBodies;

	CString m_path;
	CString m_folderName;
	int m_nId;
	//MailArray s_mails;
	int b_mails_which_sorted;  // column to sort
	int m_lastSel;  // position of last selected mail
	BOOL m_bIsDirty;  // content from s_mails was touched and needs to be copied to this mail list
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

class MailBodyContent
{
public:
	MailBodyContent() { m_pageCode = 0; m_contentOffset = 0; m_contentLength = 0; m_attachmentNamePageCode = 0; m_isEmbeddedImage = false; };
	~MailBodyContent() {};
	CStringA m_contentType;
	CStringA m_contentTransferEncoding;
	CStringA m_contentDisposition;
	CStringA m_contentId;
	CStringA m_attachmentName;
	UINT m_attachmentNamePageCode;
	//CString m_Name;  // TODO: do we need both name and filename/attachment name
	CStringA m_contentLocation;
	UINT m_pageCode;
	int  m_contentOffset;
	int m_contentLength;
	bool m_isEmbeddedImage;

	bool IsAttachment();
	bool IsInlineAttachment();
	// DecodeBodyData() doesn't merge multiple blocks of the same type
	// MboxMail::GetMailBody_mboxview() will merge multiple blocks
	bool DecodeBodyData(char* viewMapBegin, int msgLength, SimpleString* outbuf);
};

struct MboxHash;
struct MboxEqual;

struct MessageIdHash;
struct MessageIdEqual;

struct ThreadIdHash;
struct ThreadIdEqual;

class MboxMail;

typedef std::unordered_map<CStringA*, int, ThreadIdHash, ThreadIdEqual> ThreadIdTableType;
typedef std::unordered_map<CStringA*, int, MessageIdHash, MessageIdEqual> MessageIdTableType;
typedef std::unordered_map<MboxMail*, MboxMail*, MboxHash, MboxEqual> MboxMailTableType;


// Candidate for v1.0.3.42
class MailFromToAddressInfo
{
public:
	MailFromToAddressInfo() {};
	~MailFromToAddressInfo() {};

	CStringA m_fromAddress;
	long m_cnt;
};

struct AddressHash;
struct MailFromToAddressInfoEqual;

typedef std::unordered_map<CStringA*, MailFromToAddressInfo*, AddressHash, MailFromToAddressInfoEqual> MailsFromAnyToAddressTable;

class MailsFromAnyToAddressStats
{
	MailsFromAnyToAddressStats() {};
	~MailsFromAnyToAddressStats() {};

	CStringA m_toAddress;
	MailsFromAnyToAddressTable* m_receiverStats;
	long m_cnt;

	bool sortsSenders() {};
};

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


// TODO: review redundant structures
class FolderContext
{
public:
	FolderContext() {}
	~FolderContext() {}

	void SetFolderPath(CString &folderPath);
	void GetFolderPath(CString &folderpath);
	void GetDataPath(CString &dataFolderPath);
	// 
	BOOL LoadFromRegistry();
	BOOL SaveToRegistry();

	int m_setPathCount;
	CString m_folderPath;
	CString m_dataFolderPath;
	CString m_rootDataFolderPath;
	CString m_rootDataFolderPathConfig;
	CString m_rootDataFolderPathSubFolderConfig;
	CString m_mboxFileName;
	int m_emailPosition;
	int m_emailSortType;
};

#if 0
class SubjectThreadInfo
{
public:
	dlink_node<GmailLabel> m_hashMapLink;
	MboxMail *m_mail;
	int m_index;
	int m_count;
};

struct SubjectThreadHelper
{
	hashsum_t operator()(const CString *key) const
	{
		hashsum_t hashsum = StrHash((const char*)(LPCSTR)*key, key->GetLength());
		return hashsum;
	}
	bool operator()(CString *key1, CString *key2) const
	{
		if (*key1 == *key2)
			return true;
		else
			return false;
	}
	bool operator()(CString *key1, SubjectThreadInfo *key2) const
	{
		if (*key1 == key2->m_mail->m_subj)
			return true;
		else
			return false;
	}
};
#endif

enum MailTextEncodingTypes
{
	TEXT_ENCODING_0 = 0,
	TEXT_ENCODING_ASCII = 1,
	TEXT_ENCODING_UTF8 = 2,
	TEXT_ENCODING_OTHER = 3,
	TEXT_ENCODING_TYPE_MAX = 4
};

struct MailEncodingStats
{
	MailEncodingStats() { Clear(); }

	INT m_from[TEXT_ENCODING_TYPE_MAX];
	INT m_to[TEXT_ENCODING_TYPE_MAX];
	INT m_cc[TEXT_ENCODING_TYPE_MAX];
	INT m_bcc[TEXT_ENCODING_TYPE_MAX];
	INT m_subj[TEXT_ENCODING_TYPE_MAX];
	INT m_bodyTEXT[TEXT_ENCODING_TYPE_MAX];
	INT m_bodyHTML[TEXT_ENCODING_TYPE_MAX];

	static void UpdateTextEncodingStats(INT* stats, CStringA &fld, UINT code);
	static void MailEncodingStats::PrintTextEncodingStats(CString& label, INT stats[], CString& stats2Text);
	void PrintTextEncodingStats(MailEncodingStats& textEncodingStats, CString& Text);
	void Clear();
};

struct MboxMailBoolFields
{
	unsigned int m_duplicateId : 1;
	unsigned int m_DetermineEmbeddedImagesDone : 2;
	unsigned int m_isOnUserSelectedMailList : 1;
	unsigned int m_done : 1;
	unsigned int m_hasAttachments : 1;
	unsigned int m_recv : 1;
} ;

class MboxMail
{
public:

	MboxMail()
	{
		m_startOff = m_length = m_hasAttachments = m_headLength = 0;
		m_from_charsetId = m_to_charsetId = m_subj_charsetId = 0;
		m_cc_charsetId = m_bcc_charsetId = 0;
		m_timeDate = -1;
		m_recv = 0;
		m_groupId = -1;
		m_nextMail = -1;
		m_prevMail = -1;
		m_duplicateId = false;
		m_done = false;
		m_groupColor = 0;
		m_index = -1;      // index to s_mails_ref, can't be reused
		//m_mail_index = -1; // index to s_mails_*, can be reused with care
		//m_mail_cnt = 0;    // can be reused with care
		m_headLength = 0;
		m_isOnUserSelectedMailList = false;
		m_DetermineEmbeddedImagesDone = 0;
		//m_crc32 = 0xffffffff;
#if 0
		// TODO: Could save 20 bytes. May not be worth for now given CString memory inefficiency
		m_boolFields.m_done = FALSE;
		m_boolFields.m_hasAttachments = FALSE;
		m_boolFields.m_duplicateId = FALSE;
		m_boolFields.m_isOnUserSelectedMailList = FALSE;
		m_boolFields.m_DetermineEmbeddedImagesDone = 0;
		m_boolFields.m_recv = FALSE;
#endif

	}
	static MyStackWalker *glStackWalker;
	static BOOL ignoreException;
	static int runningWorkerThreadType;
	static void CollectTextEncodingStats(MailEncodingStats &textEncodingStats);
	static MailEncodingStats m_textEncodingStats;
	//

	static BOOL developerMode; 

	dlink_node<MboxMail> m_hashMapLink;
	// Need to handle Content-Type: multipart/mixed;  ZMFIXME
	std::vector <MailBodyContent*> m_ContentDetailsArray;

	//UINT32 m_crc32;
	//
	_int64 m_startOff;
	int m_hasAttachments;
	int m_length, m_headLength;
	int m_recv;         // TODO: check if it can be reused
	time_t m_timeDate;
	CStringA m_from, m_to, m_subj;
	//CString m_from_name, m_to_name;  // TODO: should we precalculate ? or calculate run time ?
	CStringA m_cc, m_bcc;
	CStringA m_from_charset, m_to_charset, m_subj_charset;
	CStringA m_cc_charset, m_bcc_charset;
	UINT m_from_charsetId, m_to_charsetId, m_subj_charsetId;
	UINT m_cc_charsetId, m_bcc_charsetId;
	CStringA m_messageId;
	CStringA m_replyId;
	CStringA m_threadId;
	CStringA m_contentType;
	// Commented out m_referenceId for now
	//CStringA m_referenceId;
	//unsigned _int64 m_threadId; // conversation thread id
	int m_groupId;       // can't be reused
	int m_groupColor;    // TODO: can't be reused practicaly
#if 1
	union {
		int m_nextMail;
		int m_mail_index;
	};
	union {
		int m_prevMail;
		int m_mail_cnt;
	};
#else
	int m_nextMail;      // can be reused after parsing og mbox and sorting
	int m_prevMail;      // can be reused after parsing og mbox and sorting
	int m_mail_index;    // index to MailArray, to s_mails_*, can be reused with care
	int m_mail_cnt;      // count of mails with same subject, can be reused with care
#endif
	bool m_duplicateId;  // Used druring processing conversations only, can be reused !!!
	bool m_done;         // Used during processing conversations only, can be reused  !!!
	int m_index;         // serves as unique id == index
	bool m_isOnUserSelectedMailList;    // TODO: check if and/or when it can be reused
	int m_DetermineEmbeddedImagesDone;  // can't be reused  !!!
	//
	// struct MboxMailBoolFields m_boolFields;
//

	//CFile fp;
	BOOL GetBody(CStringA &str);
	BOOL GetBodySS(CFile &fp, SimpleString *str, int maxLength = -1);
	BOOL GetBodySS(SimpleString *str, int maxLength = -1);
	//
	int SingleMailSizeof();
	//
	static size_t AllMailsSizeof(int count);

	static ThreadIdTableType *m_pThreadIdTable;
	static MessageIdTableType *m_pMessageIdTable;
	static MboxMailTableType *m_pMboxMailTable;
	using MboxMailMapType = IHashMap<MboxMail, MboxMail, MboxHash, MboxEqual, &MboxMail::m_hashMapLink>;
	using MboxMailMapCollionList = dllist<MboxMail, &MboxMail::m_hashMapLink>;
	static MboxMailMapType *m_pMboxMailMap;
	//
	static UINT createMessageIdTable(UINT count);
	static void clearMessageIdTable();
	static int getMessageId(CStringA*key);
	static bool insertMessageId(CStringA*key, int val);
	static int add2ConversationGroup(int mid, MboxMail *m);

	static UINT createThreadIdTable(UINT count);
	static void clearThreadIdTable();
	static int getThreadId(CStringA*key);
	static bool insertThreadId(CStringA*key, int val);

	//static MboxMail* getMboxMail(CString *key);
	//static bool insertMboxMail(CString *key, MboxMail *mbox);
	static MboxMail* getMboxMail(MboxMail *key);
	static bool insertMboxMail(MboxMail *key, MboxMail *mbox);
	static UINT createMboxMailTable(UINT count);
	static void clearMboxMailTable();
	static MboxMailMapCollionList* MboxMail::getCollisionList(MboxMail *key);
	static MboxMailMapCollionList* MboxMail::getCollisionList(hashsum_t hashsum);
	static bool equal(MboxMail *elem1, MboxMail *elem2);
	//
	static int m_nextGroupId;
	// Tricky to use to avoid ownership conflict when two function on stack use the same buffer; Asking for trouble :) Big Yes
	static SimpleString *m_outbuf;
	static SimpleString *m_inbuf;
	static SimpleString *m_outdata;
	static SimpleString *m_indata;
	static SimpleString *m_workbuf;
	static SimpleString *m_tmpbuf;
	static SimpleString *m_largebuf;
	// define buffer for local use only within single function
	// TODO: find time and fix this mess
	static SimpleString *m_largelocal1;
	static SimpleString *m_largelocal2;
	static SimpleString *m_largelocal3;
	static SimpleString *m_smalllocal1;
	static SimpleString *m_smalllocal2;

	// added below later; must implement large buffer manager ? just created mess to resolve later
	static BOOL m_outbufBusy;
	static BOOL m_inbufBusy;
	static BOOL m_outdataBusy;
	static BOOL m_indataBusy;
	static BOOL m_workbufBusy;
	static BOOL m_tmpbufBusy;
	static BOOL m_largebufBusy;

	static SimpleString *get_outbuf();
	static SimpleString *get_inbuf();
	static SimpleString *get_outdata();
	static SimpleString *get_indata();
	static SimpleString *get_workbuf();
	static SimpleString *get_tmpbuf();
	static SimpleString *get_largebuf();

	static void rel_outbuf();
	static void rel_inbuf();
	static void rel_outdata();
	static void rel_indata();
	static void rel_workbuf();
	static void rel_tmpbuf();
	static void rel_largebuf();

	static int m_Html2TextCount;

	static _int64 s_fSize; // current File size
	static _int64 s_oSize; // old file size
	static CString s_path; // current archive file path
	static CString s_path_label; // current label list
	static CString s_datapath;  //current root directory for .mboxview and index files and many temp files
	static void SetLastPath(CString &path);
	static CString GetLastPath();
	//
	static CString GetLastDataPath();
	static CString SetLastDataPath(CString *lastMboxDirPath = 0);

	static FolderContext s_folderContext;  // current folder context, data folder, etc

	static void SetMboxFilePath(CString &filepath, BOOL ignoreCheck = FALSE);
	
	static _int64 s_curmap, s_step;
	static const CUPDUPDATA* pCUPDUPData;
	static void Parse(LPCWSTR path);
	static bool Process(ProgressTimer &progressTimer, char *p, DWORD size, _int64 startOffset, bool bFirstView, bool bLastView, _int64 &lastStartOffset, bool bEml, _int64 &msgOffset, CString &statusText, BOOL parseContent = TRUE);
	static void FindDateInHeader(char *data, int datelen, CStringA& dateStr);
	static void MonthToString(int month, CStringA &monthStr);
	static time_t parseRFC822Date(CStringA &date, CStringA &format);
	static BOOL CreateRFC822Date(CStringA &date, CStringA &rfcDateString);
	static int NormalizeText(MboxMail* m);
	//
	static void Parse_LabelView(LPCWSTR path);

	static BOOL m_seExceptionMsgBox;
	static BOOL m_cppExceptionMsgBox;
	

	// TODO:  need to rewrite, encapsulate in objects, etc
	static MailArray s_mails_ref;  // original master cache. elemenyt order must not change !!!!!
	static MailArray s_mails;  // other lists need to be copied here since that is how initial implementation works, would need to change in too many places
	static MailArray s_mails_all;
	static MailArray s_mails_find;
	static MailArray s_mails_edit;  // TODO: rename to User Selected List
	static MailArray s_mails_folder;
	static MailArray s_mails_label;
	//
	static MailArray s_mails_selected;
	static MailArray s_mails_merged;
	static int nWhichMailList;

	static MailList *m_mailList;  // must match nWhichMailList
	static MailList m_allMails;
	static MailList m_findMails;
	static MailList m_editMails;
	static MailList m_folderMails;
	static MailList m_labelMails;

	static DLLIST(MailList, m_listLink) m_folderList;

	static BOOL IsAllMailsSelected();
	static BOOL IsFindMailsSelected();
	static BOOL IsUserMailsSelected();
	static BOOL IsFolderMailsSelected();
	static BOOL IsLabelMailsSelected();

	static int AllMailsSelectedId();
	static int FindMailsSelectedId();
	static int UserMailsSelectedId();
	static int FolderMailsSelectedId();
	static int LabelMailsSelectedId();

	static bool b_mails_sorted;
	static int b_mails_which_sorted;

	static int m_EmbededImagesNoMatch;

	static int m_EmbededImagesFoundMHtml;
	static int m_EmbededImagesFoundMHtmlHtml;
	static int m_EmbededImagesFoundUnexpectedMHtml;

	static int m_EmbededImagesFound;
	static int m_EmbededImagesFoundCid;
	static int m_EmbededImagesFoundHttp;
	static int m_EmbededImagesFoundHttps;
	static int m_EmbededImagesFoundMHtmlHttp;
	static int m_EmbededImagesFoundMHtmlHttps;
	static int m_EmbededImagesFoundData;
	static int m_EmbededImagesFoundLocalFile;
	//
	static int m_EmbededImagesNotFound;
	static int m_EmbededImagesNotFoundCid;
	static int m_EmbededImagesNotFoundHttp;
	static int m_EmbededImagesNotFoundHttps;
	static int m_EmbededImagesNotFoundMHtmlHttp;
	static int m_EmbededImagesNotFoundMHtmlHttps;
	static int m_EmbededImagesNotFoundData;
	static int m_EmbededImagesNotFoundLocalFile;

	static void SortBySubjectBasedConversasions(MailArray *s_mails = 0, bool bDesc = false);
	static void SortByDate(MailArray *s_mails = 0, bool bDesc = false);
	static void SortByFrom(MailArray *s_mails = 0, bool bDesc = false);
	static void SortByTo(MailArray *s_mails = 0, bool bDesc = false);
	static void SortBySubject(MailArray *s_mails = 0, bool bDesc = false);
	static void SortBySize(MailArray *s_mails = 0, bool bDesc = false);
	static void SortByGroupId(MailArray *s_mails = 0, bool bDesc = false);
	static void SortByFileOffset(MailArray *s_mails = 0, bool bDesc = false);
	static void SortByConverstionGroups(MailArray *s_mails = 0, bool bDesc = false);
	static void SortByIndex(MailArray *s_mails = 0, bool bDesc = false);
	static void assignColor2ConvesationGroups(MailArray *mails);
	static void Destroy(MailArray *marray = 0);
	static void DestroyMboxMail(MboxMail *m);
	static bool preprocessConversationsByThreadId();
	static bool HasGMThreadId();
	static bool preprocessConversations();
	static bool sortConversations();
	static bool validateSortConversations();
	static int charCount(char *fld, char c);
	static int nstrcpy(char *dst, char *src);
	static int escapeSeparators(char *workbuff, char *fldstr, int fldlen, char sepchar);
	static int splitMailAddress(const char *buff, int bufflen, SimpleString *name, SimpleString *addr);
	static int splitMailAddress(const char* buff, int bufflen, CStringA &name, CStringA& addr);
	static int splitMailAddress(const char* buff, int bufflen, CString& name, CString& addr);
	static int EnforceCharacterLimit(SimpleString *buffer, CString &characterLimit);
	static int EnforceFieldTextCharacterLimit(char *buffer, int bufferLength, CString &characterLimit);
	//
	static int CreateFldFontStyle(HdrFldConfig &hdrFieldConfig, CStringA &fldNameFontStyle, CStringA &fldTextFontStyle);
	static int printMailHeaderToHtmlFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig, HdrFldConfig &hdrFieldConfig, CString &attachmentFilesFolderPath);
	static int printSingleMailToHtmlFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig, bool singleMail, BOOL pageBreakNeeded, BOOL fullImgFilePath);
	static int printSingleMailToTextFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig, bool singleMail, BOOL pageBreakNeeded);
	static int printMailHeaderToTextFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm, TEXTFILE_CONFIG &textConfig, HdrFldConfig &hdrFieldConfig);
	static int exportToTextFile(TEXTFILE_CONFIG &textConfig, CString &textFileName, int firstMail, int lastMail, MailIndexList *selectedMailIndexList, int textType, BOOL progressBar);
	static BOOL PageBreakNeeded(int mailPosition, int nextMailPosition, bool singleMail);
	static BOOL PageBreakNeeded(MailIndexList *selectedMailIndexList, int mailPosition, bool singleMail);
	static int printDummyMailToHtmlFile(/*out*/CFile &fp);
	
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
	static int PrintMailRangeToSingleTextFile(TEXTFILE_CONFIG &textConfig, CString &textFileName, int firstMail, int lastMail, int textType, CString &targetPrintSubFolderName, CString &errorText);
	static int PrintMailRangeToSingleTextFile_WorkerThread(TEXTFILE_CONFIG &textConfig, CString &textFileName, int firstMail, int lastMail, int textType, CString errorText);
	//
	static int PrintMailSelectedToSingleTextFile_WorkerThread(TEXTFILE_CONFIG &textConfig, CString &textFileName, MailIndexList *selectedMailIndexList, int textType, CString errorText);
	//
	// textMinorType = 0 if text/plain, 1 if text/html;  plainTextMode = 0 for old mode no image info, 1 for plain text, 2 for html text
	static int GetMailBody_mboxview(CFile &fpm, int mailPosition, SimpleString *outbuf, UINT &pageCode, int textMinorType, int plainTextMode);  
	static int GetMailBody_mboxview(CFile &fpm, MboxMail *m, SimpleString *outbuf, UINT &pageCode, int textMinorType, int plainTextMode);
	static int AppendInlineAttachmentNameSeparatorLine(MailBodyContent* body, int bodyCnt, SimpleString* outbuf, int textType);
	//
	static int GetMailBody_MailBody(CFile &fpm, int mailPosition, SimpleString *outbuf, UINT &pageCode);
	static int GetMailBody_CMimeMessage(CMimeMessage &mail, int mailPosition, SimpleString *outbuf, UINT &pageCode);
	static void getCMimeBodyHeader(CMimeMessage *mail, CMimeBody* pBP, CMBodyHdr *pHdr);
	static int LoadMail(const char* pszData, int nDataSize);
	static char *ParseContent(MboxMail *mail, char *startPos, char *endPos);
	static int CreateImgAttachmentFiles(CFile &fpm, int mailPosition, SimpleString *outbuf);
	static int DecodeBody(CFile &fpm, MailBodyContent *body, int mailPosition, SimpleString *outbuf);
	//
	static int DumpMailSummaryToFile(MailArray* mailsArray, int mailArrayCount);
	static int DumpMailStatsToFile(MailArray *mailsArray, int mailArrayCount);
	//
	static int DetermineLimitedLength(SimpleString *str, int maxLinesTextLimit);
	static int MergeTwoMailLists(MailArray *mails1, MailArray *mails2, MailArray *merged_mails);
	static BOOL VerifyMergeOfTwoMailLists(MailArray *mails1, MailArray *mails2, MailArray *merged_mails);
	static BOOL Test_MergeTwoMailLists();
	static void PopulateCArray(MailArray *mails, int *carray, int carrayCnt);
	static BOOL VerifyMerge(MailArray *mails, int *carray, int carrayCnt);
	//static void ShellExecuteError2Text(UINT errorCode, CString errorText);
	static int MakeFileName(MboxMail *m, struct NamePatternParams *namePatternParams, CString &fileName, CStringA &fileNameA);
	static int MakeFileName(MboxMail *m, struct NamePatternParams* namePatternParams, struct NameTemplateCnf &nameTemplateCnf, CString &fileName, int maxFileNameLength, CStringA& fileNameA);
	static CString MatchFldLabel(LPCWSTR p, LPCWSTR e);
	static BOOL ParseTemplateFormat(CString &templateFormat, CArray<CString> &labelArray);
	static BOOL TemplateFormatHasLabel(CString &label, CArray<CString> &labelArray);
	//
	static CString GetPickerDateFormat(int i);
	static CString GetDateFormat(int i);
	static int RemoveDuplicateMails(MailArray &s_mails_array, BOOL putDuplicatesOnFindArray);
	static int LinkDuplicateMails(MailArray &s_mails_array);

	static int MakeFileNameFromMailArchiveName(int fileType, CString &fileName, CString &targetPrintSubFolder, bool &fileExists, CString &errorText);
	static int MakeFileNameFromMailHeader(int mailIndex, int fileType, CString &fileName, CString &targetPrintSubFolder, bool &fileExists, CString &errorText);
	//
	static BOOL CreateCachePath(CString &rootPrintSubFolder, CString &targetPrintSubFolder, CString &prtCachePath, CString &errorText);
	static BOOL GetCachePath(CString &rootPrintSubFolder, CString &targetPrintSubFolder, CString &prtCachePath, CString &errorText, CString *mboxFilePath=NULL);


	static BOOL GetLongestCachePath(CString &longestCachePath);
	static BOOL GetPrintCachePath(CString &printCachePath);
	static bool GetPrintCachePath(CString &rootPrintSubFolder, CString &targetPrintSubFolder, CString &prtCachePath, CString &errorText);
	static BOOL GetMboxviewFilePath(CString &mboxFilePath, CString &mboxviewFilePath);
	static BOOL GetFileCachePathWithoutExtension(CString &mboxFilePath, CString &mboxBaseFileCachePath);
	//
	static void ReleaseResources(BOOL upadteRegistry = TRUE);
	static void assert_unexpected();

	static HintConfig m_HintConfig;
	static void ShowHint(int hintNumber, HWND h = NULL);
	static void LoadHintBitmap();

	static BOOL ParseDateInFromField(char *p, char *end, SYSTEMTIME *sysTime);

	static int DetermineEmbeddedImages(int mailPosition, /*in - mail body*/ CFile &fpm);
	//static int printAttachmentNamesAsHtml(CFile *fpm, int mailPosition, CStringW &htmlbuf);
	static int printAttachmentNamesAsHtml(CFile *fpm, int mailPosition, SimpleString *outbuf, CString &attachmentFileNamePrefix, CString& attachmentFilesFolderPath);
	static int printAttachmentNamesAsText(CFile *fpm, int mailPosition, SimpleString *outbuf, CString &attachmentFileNamePrefix);

	static void DumpMailParseException(_int64 msgOffset);

	static void encodeTextAsHtml(SimpleString &txt);
	static void encodeTextAsHtmlLink(SimpleString &txt);
	static void encodeTextAsHtmlLinkLabel(SimpleString &txt);
	static void encodeTextAsHtmlLinkLabel(CStringA link, CStringA& htmlLink);

	static void mbassert();
};

inline char* FixIfNull(const char* ptr)
{
	return ptr ? ptr : "null";
};

#endif // _MBOX_MAIL_