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

#if !defined(AFX_NLISTVIEW_H__DAE62ED6_932A_4145_B641_8CFD7B72EB2D__INCLUDED_)
#define AFX_NLISTVIEW_H__DAE62ED6_932A_4145_B641_8CFD7B72EB2D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// NListView.h : header file
//

#include <vector>
#include <map>
#include "Mime.h"
#include "MimeCode.h"
#include "FileUtils.h"
#include "WheelListCtrl.h"
#include "UPDialog.h"
#include "FindAdvancedDlg.h"
#include "FindDlg.h"
#include "TextUtilities.h"
#include "AttachmentsConfig.h"
#include "MyCTime.h"
#include "ForwardMailDlg.h"


class CMimeMessage;
class MboxMail;
class MyMailArray;
class SerializerHelper;
class SimpleString;
class NMsgView;

// BOOL SaveMails(LPCWSTR cache, BOOL mainThread, CString &errorText);  // FIXMEFIXME

typedef CArray<int, int> MailIndexList;
class MailBodyContent;
typedef MyMailArray MailArray;

/////////////////////////////////////////////////////////////////////////////
// NListView window

class HtmlAttribInfo
{
public:
	HtmlAttribInfo(CStringA &name)
	{
		m_name = name; m_quoted = FALSE; m_separator = 0; m_terminator = ' ';
	}
	HtmlAttribInfo()
	{
		m_quoted = FALSE; m_separator = 0; m_terminator = ' ';
	}
	~HtmlAttribInfo() {}

	void Reset()
	{
		m_name.Empty(); m_value.Empty(); m_quoted = FALSE;  m_separator = 0; m_terminator = ' '; m_attribString.Empty();
	}

	CStringA m_name;
	CStringA m_value;
	BOOL m_quoted;
	char m_separator;
	char m_terminator;
	CStringA m_attribString;
};

class MailBodyInfo
{
public:
	MailBodyInfo() { m_index = 0; };
	~MailBodyInfo() {};
	CStringA m_CID;
	CStringA m_imgFileName;
	int m_index;
};

typedef MyCArray<MailBodyInfo*> MailBodyInfoArray;

class AttachmentData
{
public:
	AttachmentData() {
		m_nextId = 0; m_isEmbedded = FALSE; m_width = 0; m_height = 0; 
		m_length = 0; m_crc = 0;
	}
	~AttachmentData() {}

	CStringW m_nameW;
	BOOL m_isEmbedded;
	int m_width;
	int m_height;
	int m_nextId;
	_int64 m_length;
	_int64 m_crc;
};

typedef CMap<CStringW, LPCWSTR, AttachmentData, AttachmentData> AttachmentDB;
typedef CArray<AttachmentData> AttachmentArray;


// Allows to detect duplicate names. When duplication is detected, unique ID is returned to the caller
// to enable caller to decide to create unique name or not
class AttachmentMgr
{
public:
	AttachmentMgr() {}
	~AttachmentMgr() {}

	void Clear();

	// Use std::map to reduce overhead // FIXME
	// std::map<CStringW, AttachmentData> m_attachmentMap;

	AttachmentDB m_attachmentMap;
	AttachmentArray m_attachmentArray;

	// Insert into DB if not present and return -1
	// otherwise retrun >= 0

	AttachmentData* GetAttachmentData(const CStringW& inName);
	BOOL HasName(CStringW &inName);
	void AddName(CStringW& inName, BOOL isEmbedded);
	int InsertName(CStringW& inName, BOOL isEmbedded);
	void Sort();
	void PrintMap(CString& title);
	void PrintArray(CString& title);
};

class EmbededImagesStats
{
public:
	EmbededImagesStats() { Clear(); }

	void Clear();

	int m_EmbededImagesNoMatch;

	int m_FoundMHtml;
	int m_FoundMHtmlHtml;
	int m_FoundUnexpectedMHtml = 0;

	int m_Found = 0;
	int m_FoundCid = 0;
	int m_FoundHttp = 0;
	int m_FoundHttps = 0;
	int m_FoundMHtmlHttp = 0;
	int m_FoundMHtmlHttps = 0;
	int m_FoundData = 0;
	int m_FoundLocalFile = 0;
	//
	int m_NotFound = 0;
	int m_NotFoundCid = 0;
	int m_NotFoundHttp = 0;
	int m_NotFoundHttps = 0;
	int m_NotFoundMHtmlHttp = 0;
	int m_NotFoundMHtmlHttps = 0;
	int m_NotFoundData = 0;
	int m_NotFoundLocalFile = 0;
};

struct MailArchiveFileInfo
{
	int m_version;
	_int64 m_fileSize;
	_int64 m_oldestMailTime;
	_int64 m_latestMailTime;
	CStringA m_reservedString1;
	CStringA m_reservedString2;
	_int64 m_reservedInt1;
	_int64 m_reservedInt2;
	int m_mailCount;
};

class NListView : public CWnd
{
	CWheelListCtrl	m_list;
	CFont		m_font;
	CFont m_boldFont;
	UINT m_acp; // ANSI code page for this system

// Construction
public:
	NListView();
	DECLARE_DYNCREATE(NListView)
// Attributes
public:
	CString m_format;
	void ResetSize();
	void SetSetFocus();
// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(NListView)
	protected:
	virtual void PostNcDestroy();
	//}}AFX_VIRTUAL

// Implementation
public:

	// Global Vars -- FIXME
	// avoid changes to multiple interfaces -- find better solution
	// set/reset  in NListView::OnRClickSingleSelect: case S_PRINTER_Id and  case S_PRINTER_GROUP_Id:
	// and in NListView::OnRClickMultipleSelect: case S_PRINTER_GROUP_Id: 
	static BOOL m_fullImgFilePath;   // examined in NListView::UpdateInlineSrcImgPathEx
	static BOOL m_fullImgFilePath_Config;
	static BOOL m_appendAttachmentPictures;;
	//
	// Global Vars  -- FIXME

	EmbededImagesStats m_EmbededImagesStats;

	int m_frameCx_TreeNotInHide;
	int m_frameCy_TreeNotInHide;
	int m_frameCx_TreeInHide;
	int m_frameCy_TreeInHide;
	//
	BOOL m_bApplyColorStyle;
	BOOL m_bLastApplyColorStyle;
	BOOL m_showAsPaperClip; // Show attachment indicataor as paper Clip, otherwise as star character
	CBitmap m_paperClip;
	CImageList m_imgList;

	// Forward Mail
	ForwardMailData m_ForwardMailData;
	//CImageList m_ImageList;
	MailIndexList m_selectedMailsList;

	// vars to handle mapped mail file
	BOOL m_bMappingError;
	HANDLE m_hMailFile;  // handle to mail file
	_int64 m_MailFileSize;
		//
	HANDLE m_hMailFileMap;
	_int64 m_mappingSize;
	int m_mappingsInFile;

		// offset and pointers for map view
	_int64 m_curMapBegin;
	_int64 m_curMapEnd;
	char *m_pMapViewBegin;
	char *m_pMapViewEnd;
	DWORD m_dwPageSize;
	DWORD m_dwAllocationGranularity;

		// offset & bytes requsted and returned pointers 
	_int64 m_OffsetRequested;
	DWORD m_BytesRequested;
	char *m_pViewBegin;
	char *m_pViewEnd;
	DWORD m_dwViewSize;
		//
	_int64 m_MapViewOfFileExCount;
		//
	int m_fontSizePDF;
		//
	BOOL SetupFileMapView(_int64 offset, DWORD length, BOOL findNext);

	// TESTING
	BOOL MatchIfFieldFolded(int mailPosition, char *fld);
	// 
	int CheckMatch(int which, CStringA &searchString);
	
	int MatchHeaderFldSingleAddress(int fldIndx, CStringA &fld, UINT hdrFldCodePage, CFindAdvancedParams &params, int pos = 1);
	int MatchHeaderFld(int fldIndx, CStringA &fld, UINT hdrFldCodePage, CFindAdvancedParams &params, int pos = 1);
	int CheckMatchAdvanced(int i, CFindAdvancedParams &params);
	void DetermineKeywordsForProgresBar(CString *m_string, CString &keyword1, CString &keyword2);  // for Advanced Find
	BOOL FindInMailContent(int mailPosition, BOOL bContent, BOOL bAttachment, CStringA &searchString);
	BOOL AdvancedFindInMailContent(int mailPosition, BOOL bContent, BOOL bAttachment, CFindAdvancedParams &params);
	void CloseMailFile();
	void ResetFileMapView();
	void SortByColumn(int colNumber, BOOL sortByPosition = FALSE);
	void SortBySubjectBasedConversasions(BOOL justRefresh = FALSE);
	void RefreshSortByColumn();
	BOOL HasAnyAttachment(MboxMail *m);
	// end of vars

	BOOL m_bExportEml;
	int m_subjectSortType;

	//#####################################
	// *** Search related vars block
	
	int m_SearchType;  // 0 = none, 1=basic, 2=advanced

	// For single Mail.
	CString m_searchStringInMail;
	CStringA m_searchStringInMailA;
	BOOL m_bCaseSensInMail;
	BOOL m_bWholeWordInMail;

	// BOOL m_bFindNext;  // FIXMEFIXME
	BOOL m_bInFind;
	HTREEITEM m_which;

	int SelectItem(int which, BOOL ignoreViewMessageHeader = FALSE);
	int AppendInlineAttachmentNameSeparatorLine(CMimeBody* pBP, int bodyCnt, CStringA& bdy, int textType);
	int DoFastFind(int searchstart, BOOL mainThreadContext, int maxSearchDuration, BOOL findAll);
	int DoFastFindLegacy(int searchstart, BOOL mainThreadContext, int maxSearchDuration, BOOL findAll);
	int DoFastFindAdvanced(int searchstart, BOOL mainThreadContext, int maxSearchDuration, BOOL findAll);
	static BOOL MyCTimeToOleTime(MyCTime& ctimeDateTime, COleDateTime& coleDateTime);
	static BOOL OleTime2MyCTime(COleDateTime& coleDateTime, MyCTime& ctimeDateTime, BOOL roundUP);

	CString m_searchString;
	int m_lastFindPos;
	int m_maxSearchDuration;
	BOOL m_bEditFindFirst;
	int m_findAllCount;

	// m_bHighlightAllSet is set when Highlight option is selected in Find to highlight all search string occurences 
	// when page is loaded , see CBrowser::OnDocumentCompleteExplorer
	BOOL m_bHighlightAllSet;    
	// Used by Search Text option
	BOOL	m_bCaseSens;
	BOOL	m_bWholeWord;

	struct CFindDlgParams m_findParams;
	struct CFindAdvancedParams m_advancedParams;
	CStringA m_stringWithCase[FILTER_FIELDS_NUMB];
	BOOL m_advancedFind;

	// *** End of Search related vars block
	//#####################################

	void ClearDescView();
	CString m_curFile;
	int m_lastSel;
	int m_bStartSearchAtSelectedItem;
	int m_gmtTime;
	CString m_path;
	CString m_path_label;

	// Used in Custom Draw
	SimpleString *m_name;
	SimpleString *m_addr;
	BOOL m_bLongMailAddress;
	// timer
	UINT_PTR m_nIDEvent;
	UINT m_nElapse;
	//
	std::vector <MailBodyInfo*> m_BodyInfoArray;
	//
	int SaveAsEmlFile(CStringA &bdy);
	int SaveAsEmlFile(char *bdy, int bdylen);
	void FindImageFileName(CString &cid);
	//
	int LoadMails(LPCWSTR cache, MailArchiveFileInfo& maileFileInfo, MailArray *mails, CString& errorText);
	int LoadMailsInfo(SerializerHelper& sz, MailArchiveFileInfo& maileFileInfo, CString &errorText);

	int LoadSingleMail(MboxMail* m, SerializerHelper &sz);
	int Cache2Text(LPCWSTR cache, CString format);
	void FillCtrl();
	int FillCtrl_ParseMbox(CString &mboxPath);
	int  MailFileFillCtrl(CString &errorText);
	virtual ~NListView();
	void SelectItemFound(int iItem);
	void ResetFont();
	void RedrawMails();
	void ResizeColumns();
	time_t OleToTime_t(COleDateTime *ot);
	int MailsWhichColumnSorted() const;
	void SetLabelOwnership();
	void ItemState2Str(UINT uState, CString &strState);
	void ItemChange2Str(UINT uChange, CString &strState);
	void PrintSelected();  // debug support
	void OnRClickSingleSelect(NMHDR* pNMHDR, LRESULT* pResult);
	void OnRClickMultipleSelect(NMHDR* pNMHDR, LRESULT* pResult);
	void PrintToPDFMultipleSelect(int fontSize);
	int RemoveSelectedMails();
	int RemoveAllMails();
	int CopySelectedMails();
	int CopyAllMails();
	int FindInHTML(int iItem);
	void SwitchToMailList(int nID, BOOL force = FALSE);
	//void EditFindAdvanced(CStringA *from = 0, CStringA *to = 0, CStringA *subject = 0);
	void EditFindAdvanced(MboxMail *m = 0);
	void RunFindAdvancedOnSelectedMail(int iItem);
	int PopulateUserMailArray(SerializerHelper &sz, int mailListCnt, BOOL verifyOnly);
	int PopulateMailArray(SerializerHelper &sz, MailArray &mArray, int mailListCnt, BOOL verifyOnly);
	int OpenArchiveFileLocation();
	int OpenMailListFileLocation();
	int RemoveDuplicateMails(MailArray &s_mails_array);
	BOOL IsUserSelectedMailListEmpty();
	int ReloadMboxListFile_v2(CString *mboxListFile=0);
	int SaveAsMboxListFile_v2();
	int SaveAsMboxArchiveFile_v2();
	int FindMailListFileWithHighestNumber(CString &folder, CString &extension);
	// Folder related
	int CreateEmptyFolder(CString &drivename, CString &mboxDirectory, CString &mboxFolderName, CString &parentSubFolderPath, CString &newFolderName);
	int CreateEmptyFolderListFile(CString &path, CString &folderListFile);
	int LoadFolderListFile_v2(CString &folderPath, CString &folderName);
	int CopyMailsToFolders();
	//
	int SaveAsLabelFile(MailArray *marray, CString &targetDir, CStringA &labelName, DWORD labelCodePage, CStringA& mappedLabelName, CString &errorText);
	int LoadLabelListFile_v2(CString &folderPath, CString &folderName, CString &mboxFilePath);
	int WriteMboxListFile_v2(MailArray *mailsArray, CString &listFilePath, _int64 mboxFileSize, CString &errorText);
	int WriteMboxLabelListFile(MailArray* mailsArray, CStringA &gLabel, DWORD gLabelCodePage, CString& listFilePath, _int64 mboxFileSize, CString& errorText);
	int GetLabelFromLabelListFile_v2(CString& listFilePath, CStringA& gLabel, DWORD& gLabelCodePage, CString &mboxFilePath);

	// #########################################
	// Forward emails vars block

	int m_tcpPort;
	BOOL m_enableForwardMailsLog;
	BOOL m_enableSMTPProtocolLog;

	int VerifyPathToForwardEmlFileExecutable(CString& ForwardEmlFileExePath, CString& errorText);
	int DeleteForwardEmlFiles(CString &folderPath);
	int ForwardSingleMail(int iItem, BOOL progressBar, CString &progressText, CString &errorText);
	int ForwardMailDialog(int iItem);
	int ForwardSelectedMails(int iItem);
	int ForwardMailRange(int iItem);
	int ForwardMails_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName);
	int ForwardSelectedMails_Thread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName);
	int ForwardMails_WorkerThread(ForwardMailData &mailData, MailIndexList *selectedMailIndexList, CString &errorText);
	INT64  ExecCommand_WorkerThread(int tcpPort, CString instanceId, CString &password, ForwardMailData &mailData, 
		CString &emlFile, CString &errorText, BOOL progressBar, CString &progressText, int timeout = -1);

	// End of Forward emails vars block
	// #########################################

	CString FixCommandLineArgument(CString &in);
	CString FixCommandLineArgument(int in);
	INT64 ExecCommand_KillProcess(CString processName, CString &errorText, BOOL progressBar, CString &progressText);

	MailIndexList *PopulateSelectedMailsListFromRange(int firstIndex, int lstIndex);
	MailIndexList *PopulateSelectedMailsList();
	void FindFirstAndLastMailOfConversation(int iItem, int &firstMail, int &lastMail);
	void FindFirstAndLastMailOfMailThreadConversation(int iItem, int &firstMail, int &lastMail);
	void FindFirstAndLastMailOfSubjectConversation(int iItem, int &firstMail, int &lastMail);
	int RefreshMailsOnUserSelectsMailListMark();
	int VerifyMailsOnUserSelectsMailListMarkCounts();
	//
	void PrintMailGroupToText(BOOL multipleSelectedMails, int iItem, int textType, BOOL forceOpen = FALSE, BOOL printToPrinter = FALSE, BOOL createFileOnly = FALSE);
	//void PrintMailSubjectThreadToText(BOOL multipleSelectedMails, int iItem, int textType, BOOL forceOpen = FALSE, BOOL printToPrinter = FALSE, BOOL createFileOnly = FALSE);
	
	int PrintMailRangeToSingleCSV_Thread(int iItem);
	// Debug Helpers
	int ScanAllMailsInMbox();
	int ScanAllMailsInMbox_NewParser();
	//

	//BOOL PrepopulateAdvancedSearchParams(CStringA *from, CStringA *to, CStringA *subject);
	BOOL PrepopulateAdvancedSearchParams(MboxMail *m);
	void FindMinMaxTime(MyCTime &minTime, MyCTime &maxTime);
	void ResetFilterDates();

	void PostMsgCmdParamAttachmentHint();
	//
	int ExportTextTextToTextFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm);
	//
	static int ExportAsEmlFile(CFile *fpm, int mailPosition, CString &targetDirectory, CString &emlFile, CString &errorText);
	static int PrintAsEmlFile(CFile *fpm, int mailPosition, CString &emlFile);
	//
	//  Attachment related functions
	// FIXME Both PrintMailAttachments & CreateMailAttachments create attachments; compare and consolidate if possible
	static int PrintMailAttachments(CFile *fpm, int mailPosition, AttachmentMgr &attachmentDB);
	static int CreateMailAttachments(CFile* fpm, int mailPosition, CString* attachmentFolderPath, BOOL prependMailId, AttachmentMgr& attachmentDB, 
		BOOL extraFileNameValidation = TRUE);
	//
	// PrintAttachmentNamesAsText2CSV & MboxMail::printAttachmentNamesAsHtml & printAttachmentNamesAsText print names but don't create files
	static int PrintAttachmentNamesAsText2CSV(int mailPosition, SimpleString *outbuf, CString &characterLimit, CStringA &attachmentSeparator);
	// int MboxMail::printAttachmentNamesAsHtml(CFile* fpm, int mailPosition, SimpleString* outbuf, CString& attachmentFileNamePrefix, CStringA& attachmentFilesFolderPath);
	// int MboxMail::printAttachmentNamesAsText(CFile* fpm, int mailPosition, SimpleString* outbuf, CString& attachmentFileNamePrefix);
	//
	static int DetermineAttachmentName(CFile* fpm, int mailPosition, MailBodyContent* body, SimpleString* bodyData,
		CStringW& nameW, AttachmentMgr& attachmentDB, BOOL remapDuplicateNames, BOOL extraValidation = TRUE);
	// Called in CheckMatchAdvanced and CheckMatch
	static int FindAttachmentName(MboxMail *m, CStringA &searchString, BOOL bWholeWord, BOOL bCaseSensitive);
	//
	
	static void GetExtendedMailId(MboxMail* m, CString& extendedId);
	// //
	//////////////////////////////////////////////////////
	////////////  PDF
	//////////////////////////////////////////////////////
	int PrintMailConversationToSeparatePDF_Thread(int mailIndex, BOOL mergePDFs, CString &errorText);
	int PrintMailConversationToSinglePDF_Thread(int mailIndex, CString &errorText);
	//
	// Range to Separate PDF
	int PrintMailRangeToSeparatePDF_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName, CString& targetPrintFolderPath, BOOL mergePDFs);
	//
	// Range to Single PDF
	int PrintMailRangeToSinglePDF_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName);
	//
	// Selected to Separate PDF
	int PrintMailSelectedToSeparatePDF_Thread(MailIndexList* selectedMailsIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, BOOL mergePDFs);
	int PrintMailSelectedToSeparatePDF_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText, 
		BOOL mergeFiles, CString &mergedFileName);
	//
	// Selected to Single PDF
	int PrintMailSelectedToSinglePDF_Thread(MailIndexList* selectedMailsIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath);
	int PrintMailSelectedToSinglePDF_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);
	//
	int NListView::MergePDfFileList(CFile &fp, CStringArray &in_array, CStringArray &out_array, CString &filePrefix, CString &targetPrintFolderPath, 
		CString &pdfboxJarFileName, CString &errorText);
	//
	//////////////////////////////////////////////////////
	////////////  PDF  END
	//////////////////////////////////////////////////////
	//
	//////////////////////////////////////////////////////
	////////////  HTML
	//////////////////////////////////////////////////////
	//
	int PrintMailArchiveToSeparateHTML_Thread(CString &errorText);
		//
	int PrintMailConversationToSeparateHTML_Thread(int mailIndex, CString &errorText);
	int PrintMailConversationToSingleHTML_Thread(int mailIndex, CString &errorText);
	//
	// Range to Separate HTML
	int PrintMailRangeToSeparateHTML_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName);
	//
	// Range to Single HTML
	int PrintMailRangeToSingleHTML_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName);
	//
	// Selected to Separate HTML
	int PrintMailSelectedToSeparateHTML_Thread(MailIndexList* selectedMailsIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath);
	int PrintMailSelectedToSeparateHTML_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);
	//
	// Selected to Single HTML
	int PrintMailSelectedToSingleHTML_Thread(MailIndexList* selectedMailsIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath);
	int PrintMailSelectedToSingleHTML_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);
	//
	//////////////////////////////////////////////////////
	////////////  HTML  END
	//////////////////////////////////////////////////////
	//
	//////////////////////////////////////////////////////
	////////////  TEXT
	//////////////////////////////////////////////////////
	//
		// Selected to Separate TEXT
	//int PrintMailSelectedToSeparateTEXT_Thread(CString &targetPrintSubFolderName, CString &targetPrintFolderPath);
	//int PrintMailSelectedToSeparateTEXT_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);

	// Selected to Single TEXT
	int PrintMailSelectedToSingleTEXT_Thread(CString &targetPrintSubFolderName, CString &targetPrintFolderPath);
	int PrintMailSelectedToSingleTEXT_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);
	///

	int CreateInlineImageCache_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName);
	//int CreateInlineImageCache_WorkerThread(int firstMail, int lastMail, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);
	//
	int CreateAttachmentCache_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName);
	int CreateAttachmentCache_WorkerThread(int firstMail, int lastMail, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);
	//
	int CreateEmlCache_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName);
	int CreateEmlCache_WorkerThread(int firstMail, int lastMail, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);

	static BOOL IsSingleAddress(CStringA *to);
	static void TrimToAddr(CStringA *to, CStringA &toAddr, int maxNumbOfAddr);
	static void TrimToName(CStringA *to, CStringA &toName, int maxNumbOfAddr);
	static int DeleteAllHtmAndPDFFiles(CString &targetFolder);
	//static int DeleteAllHtmFiles(CString &targetFolder);
	//
	static int CreateEmbeddedImageFilesEx(CFile& fpm, int mailPosition, CString& imageCachePath, BOOL createEmbeddedImageFiles);
	static int DetermineImageFileName(CFile* fpm, BOOL verifyAttachmentDataAsImageType, MboxMail *m, CStringA &cidName, CString &imageFilePath, MailBodyContent **foundBody, int mailPosition);
	//
	static int UpdateInlineSrcImgPathEx(CFile* fpm, char* inData, int indDataLen, SimpleString* outbuf,
		BOOL makeFileNameUnique,  // add unique prefix to image file name
		BOOL makeAbsoluteImageFilePath,  
		CString& relativeSrcImgFilePath, CString& absoluteSrcImgFilePath, AttachmentMgr& attachmentDB,
		EmbededImagesStats& embededImgStats, int mailPosition, BOOL createEmbeddedImageFiles ,
		BOOL verifyAttachmentDataAsImageType, BOOL insertMaxWidthForImgTag,
		CStringA& maxWidth, CStringA& maxHeight);

	static void AppendPictureAttachments(MboxMail *m, AttachmentMgr &attachmentDB, CString *absoluteFolderPath, CString *relativeFolderPath,  CFile* fpm);
	//
	static int FindFilenameCount(std::vector <MailBodyContent*> &contentDetailsArray, CStringA &fileName);
	//
	static int FindFilenameCount(CMimeBody::CBodyList &bodies, CStringA &fileName);

	static int RemoveBackgroundColor(char *inData, int indDataLen, SimpleString *outbuf, int mailPosition);
	static int SetBackgroundColor(char *inData, int indDataLen, SimpleString *outbuf, BOOL ReplaceAllWhiteBackgrounTags);
	static int ReplacePreTagWitPTag(char *inData, int indDataLen, SimpleString *outbuf, BOOL ReplaceAllWhiteBackgrounTags);
	static int ReplaceBlockqouteTagWithDivTag(char* inData, int indDataLen, SimpleString* outbuf, BOOL ReplaceAllWhiteBackgrounTags);
	static int AddMaxWidthToHref(char *inData, int indDataLen, SimpleString *outbuf, BOOL ReplaceAllWhiteBackgrounTags);
	static int AddMaxWidthToBlockquote(char *inData, int indDataLen, SimpleString *outbuf, BOOL ReplaceAllWhiteBackgrounTags);
	static int AddMaxWidthToDIV(char *inData, int indDataLen, SimpleString *outbuf, BOOL ReplaceAllWhiteBackgrounTags);
	static int MakeSpacesAsNBSP(char *inData, int indDataLen, SimpleString *outbuf, BOOL ReplaceAllWhiteBackgrounTags);
	static int RemoveMicrosoftSection1Class(SimpleString* inbuf, int offset);
	
	static int FindBodyTag(char *inData, int indDataLen, char *&tagBeg, int &tagDataLen);
	static int FindBodyBackgroundColor(char *inData, int indDataLen, char *&attribTag, int &attribTagLen, CStringA &bodyBackgroundColor, HtmlAttribInfo &bodyBackgroundColorAttrib);
	static int FindBodyTagAttrib(char *inData, int indDataLen, char *tag, int tagLen, /*out*/ char *&attribTag, int &attribTagLen, CStringA &attribVal, HtmlAttribInfo &bodyAttrib);
	static int RemoveBodyBackgroundColor(char *inData, int indDataLen, SimpleString *outbuf, CStringA &bodyBackgroundColor);
	static int RemoveBodyBackgroundColorAndWidth(char *inData, int indDataLen, SimpleString *outbuf, CStringA &bodyBackgroundColor, 
		CStringA &bodyWidth, BOOL removeBgColor, BOOL removeWidth);
	static int SetBodyWidth(char *inData, int indDataLen, SimpleString *outbuf, CString &bodyBackgroundColor);

	static int Color2Str(DWORD color, CStringA &colorStr);
	static int Color2Str(DWORD color, CString& colorStr);

	//static BOOL loadImage(BYTE* pData, size_t nSize, CStringW &extensionW, CStringA &extension);
	static int DetermineListFileName(CString &fileName, CString &listFileName);
	void SetListFocus();

	static BOOL SaveMails(LPCWSTR cache, BOOL mainThread, CString& errorText);

	BOOL m_developmentMode;

	// Generated message map functions
protected:
	//{{AFX_MSG(NListView)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	//}}AFX_MSG
	afx_msg void OnActivating(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnKeydown(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnColumnClick(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnUpdateEditFind(CCmdUI* pCmdUI);
	afx_msg void OnEditFind();
	afx_msg void OnUpdateEditFindAgain(CCmdUI* pCmdUI);
	afx_msg void OnEditFindAgain();
	afx_msg void OnCustomDraw(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnDoubleClick(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnDividerdblclick(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnRClick(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnItemchangedListCtrl(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnODStateChangedListCtrl(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnReleaseCaptureListCtrl(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnODFindItemListCtrl(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnODCacheHintListCtrl(NMHDR* pNMHDR, LRESULT* pResult);
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnEditVieweml();
	afx_msg void OnUpdateEditVieweml(CCmdUI *pCmdUI);
	afx_msg void OnEditFindadvanced();
	afx_msg void OnUpdateEditFindadvanced(CCmdUI *pCmdUI);
	//virtual BOOL PreTranslateMessage(MSG* pMsg);
	//afx_msg void OnSetFocus(CWnd* pOldWnd);
	afx_msg void OnMouseHover(UINT nFlags, CPoint point);
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	afx_msg LRESULT OnCmdParam_AttachmentHint(WPARAM wParam, LPARAM lParam);
	afx_msg void OnClose();
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnTvnGetInfoTip(NMHDR* pNMHDR, LRESULT* pResult);

	afx_msg BOOL OnToolNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult);

	virtual void PreSubclassWindow();

	void CellHitTest(const CPoint& pt, int& nRow, int& nCol) const;
	bool ShowToolTip(const CPoint& pt) const;
	CString GetToolTipText(int nRow, int nCol);
};

struct FORWARD_MAILS_ARGS
{
	ForwardMailData forwardMailsData;
	CString password;
	BOOL separatePDFs;
	CString errorText;
	CString targetPrintFolderPath;
	CString targetPrintSubFolderName;
	int firstMail;
	int lastMail;
	MailIndexList *selectedMailIndexList;
	int nItem;
	BOOL exitted;
	int ret;
	NListView *lview;
};

struct PARSE_ARGS
{
	CString path;
	BOOL exitted;
};

struct FIND_ARGS
{
	BOOL findAll;
	int searchstart;
	int retpos;
	BOOL exitted;
	NListView *lview;
};

struct PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS
{
	BOOL separatePDFs;
	BOOL mergePDFs;
	CString mergedPDFPath;
	CString errorText;
	CString targetPrintFolderPath;
	CString targetPrintSubFolderName;
	int firstMail;
	int lastMail;
	MailIndexList *selectedMailIndexList;
	int nItem;
	BOOL exitted;
	int ret;
	NListView *lview;
};

typedef PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS PRINT_MAIL_GROUP_TO_SEPARATE_HTML_ARGS;

struct WRITE_INDEX_FILE_ARGS
{
	CString cache;
	CString errorText;
	BOOL exitted;
	int ret;
	NListView* lview;
};

struct WRITE_IMAGE_FILE_ARGS
{
	CString cache;
	CString errorText;
	BOOL exitted;
	int ret;
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_NLISTVIEW_H__DAE62ED6_932A_4145_B641_8CFD7B72EB2D__INCLUDED_)
