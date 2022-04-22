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

BOOL SaveMails(LPCSTR cache, BOOL mainThread, CString &errorText);

typedef CArray<int, int> MailIndexList;
class MailBodyContent;
typedef MyMailArray MailArray;

/////////////////////////////////////////////////////////////////////////////
// NListView window

class HtmlAttribInfo
{
public:
	HtmlAttribInfo(CString &name)
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

	CString m_name;
	CString m_value;
	BOOL m_quoted;
	char m_separator;
	char m_terminator;
	CString m_attribString;
};

class MailBodyInfo
{
public:
	MailBodyInfo() { m_index = 0; };
	~MailBodyInfo() {};
	CString m_CID;
	CString m_imgFileName;
	int m_index;
};

typedef MyCArray<MailBodyInfo*> MailBodyInfoArray;

class AttachmentData
{
public:
	AttachmentData() { m_nextId = 0;}
	~AttachmentData() {}

	CStringW m_nameW;
	int m_nextId;
};

class AttachmentMgr
{
public:
	AttachmentMgr() {}
	~AttachmentMgr() {}

	void Clear();

	CMap<CStringW, LPCWSTR, AttachmentData, AttachmentData> m_attachmentMap;
	// Insert into DB if not present; append unique Id if name present
	int GetValidName(CStringW &inName);
};

typedef CMap<CStringW, LPCWSTR, AttachmentData, AttachmentData> AttachmentDB;


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
	BOOL SetupFileMapView(_int64 offset, DWORD length, BOOL findNext);

	// TESTING
	BOOL MatchIfFieldFolded(int mailPosition, char *fld);
	// 
	int CheckMatch(int which, CString &searchString);
	
	int MatchHeaderFldSingleAddress(int fldIndx, CString &fld, CFindAdvancedParams &params, int pos = 1);
	int MatchHeaderFld(int fldIndx, CString &fld, CFindAdvancedParams &params, int pos = 1);
	int CheckMatchAdvanced(int i, CFindAdvancedParams &params);
	void DetermineKeywordsForProgresBar(CString *stringWithCase, CString &keyword1, CString &keyword2);  // for Advanced Find
	BOOL FindInMailContent(int mailPosition, BOOL bContent, BOOL bAttachment);
	BOOL AdvancedFindInMailContent(int mailPosition, BOOL bContent, BOOL bAttachment, CFindAdvancedParams &params);
	void CloseMailFile();
	void ResetFileMapView();
	void SortByColumn(int colNumber, BOOL sortByPosition = FALSE);
	void SortBySubjectBasedConversasions(BOOL justRefresh = FALSE);
	void RefreshSortByColumn();
	BOOL HasAnyAttachment(MboxMail *m);
	// end of vars

	// For single Mail. TODO: Should consider to define structure to avoid confusing var names
	CString m_searchStringInMail;
	BOOL m_bCaseSensInMail;
	BOOL m_bWholeWordInMail;

	int m_subjectSortType;
	BOOL m_bExportEml;
	BOOL m_bFindNext;
	BOOL m_bInFind;
	HTREEITEM m_which;
	void SelectItem(int which, BOOL ignoreViewMessageHeader = FALSE);
	int DoFastFind(int searchstart, BOOL mainThreadContext, int maxSearchDuration, BOOL findAll);
	CString m_searchString;
	int m_lastFindPos;
	int m_maxSearchDuration;
	BOOL m_bEditFindFirst;

	BOOL m_bNeedToFindMailMinMaxTime;
	MyCTime m_mboxMailStartDate;
	MyCTime m_mboxMailEndDate;
	BOOL m_needToRestoreArchiveListDateTime;

	BOOL	m_filterDates;
	BOOL	m_bCaseSens;
	BOOL	m_bWholeWord;
	MyCTime m_lastStartDate;
	MyCTime m_lastEndDate;
	BOOL m_bFrom;
	BOOL m_bTo;
	BOOL m_bSubject;
	BOOL m_bContent;
	BOOL m_bCC;
	BOOL m_bBCC;
	BOOL m_bAttachments;
	BOOL m_bAttachmentName;
	BOOL m_bHighlightAll;
	BOOL m_bFindAll;
	BOOL m_bFindAllMailsThatDontMatch;
	BOOL m_bHighlightAllSet;

	//struct CFindDlgParams m_findParams;  // TODO: review later, requires too many chnages for now
	struct CFindAdvancedParams m_advancedParams;
	CString m_stringWithCase[FILTER_FIELDS_NUMB];
	BOOL m_advancedFind;

	void ClearDescView();
	CString m_curFile;
	int m_lastSel;
	int m_bStartSearchAtSelectedItem;
	int m_gmtTime;
	CString m_path;
	CString m_path_label;
	int m_findAllCount;
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
	int SaveAsEmlFile(CString &bdy);
	int SaveAsEmlFile(char *bdy, int bdylen);
	void FindImageFileName(CString &cid);
	//
	int LoadMails(LPCSTR cache, MailArray *mails = 0);
	void FillCtrl();
	int FillCtrl_ParseMbox(CString &mboxPath);
	int  MailFileFillCtrl(CString &errorText);
	virtual ~NListView();
	void SelectItemFound(int iItem);
	int DumpSelectedItem(int which);
	static int DumpItemDetails(int which);
	static int DumpItemDetails(MboxMail *m);
	static int DumpItemDetails(int which, MboxMail *m, CMimeMessage &mail);
	static int DumpCMimeMessage(CMimeMessage &mail, HANDLE hFile);
	void ResetFont();
	void RedrawMails();
	void ResizeColumns();
	time_t OleToTime_t(COleDateTime *ot);
	void MarkColumns();
	int MailsWhichColumnSorted() const;
	void SetLabelOwnership();
	void ItemState2Str(UINT uState, CString &strState);
	void ItemChange2Str(UINT uChange, CString &strState);
	void PrintSelected();
	void OnRClickSingleSelect(NMHDR* pNMHDR, LRESULT* pResult);
	void OnRClickMultipleSelect(NMHDR* pNMHDR, LRESULT* pResult);
	int RemoveSelectedMails();
	int RemoveAllMails();
	int CopySelectedMails();
	int CopyAllMails();
	int FindInHTML(int iItem);
	void SwitchToMailList(int nID, BOOL force = FALSE);
	void EditFindAdvanced(CString *from = 0, CString *to = 0, CString *subject = 0);
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
	int SaveAsLabelFile(MailArray *marray, CString &targetDir, CString &labelName, CString &errorText);
	int LoadLabelListFile_v2(CString &folderPath, CString &folderName);
	int WriteMboxListFile_v2(MailArray *mailsArray, CString &listFilePath, _int64 mboxFileSize, CString &errorText);

	int VerifyPathToForwardEmlFileExecutable(CString &ForwardEmlFileExePath, CString &errorText);

	BOOL m_developmentMode;
	int m_tcpPort;
	BOOL m_enbaleForwardMailsLog;
	BOOL m_enbaleSMTPProtocolLog;
	int DeleteForwardEmlFiles(CString &folderPath);
	int ForwardSingleMail(int iItem, BOOL progressBar, CString &progressText, CString &errorText);
	int ForwardMailDialog(int iItem);
	int ForwardSelectedMails(int iItem);
	int ForwardMailRange(int iItem);
	int ForwardMails_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName);
	int ForwardSelectedMails_Thread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName);
	int ForwardMails_WorkerThread(ForwardMailData &mailData, MailIndexList *selectedMailIndexList, CString &errorText);
	INT64  ExecCommand_WorkerThread(int tcpPort, CString instanceId, CString &password, ForwardMailData &mailData, CString &emlFile, CString &errorText, BOOL progressBar, CString &progressText);
	CString FixCommandLineArgument(CString &in);
	CString FixCommandLineArgument(int in);
	INT64 ExecCommand_KillProcess(CString processName, CString &errorText, BOOL progressBar, CString &progressText);

	MailIndexList * PopulateSelectedMailsList();
	void FindFirstAndLastMailOfConversation(int iItem, int &firstMail, int &lastMail);
	void FindFirstAndLastMailOfMailThreadConversation(int iItem, int &firstMail, int &lastMail);
	void FindFirstAndLastMailOfSubjectConversation(int iItem, int &firstMail, int &lastMail);
	int RefreshMailsOnUserSelectsMailListMark();
	int VerifyMailsOnUserSelectsMailListMarkCounts();
	//
	void PrintMailGroupToText(BOOL multipleSelectedMails, int iItem, int textType, BOOL forceOpen = FALSE, BOOL printToPrinter = FALSE, BOOL createFileOnly = FALSE);
	int PrintMailRangeToSingleCSV_Thread(int iItem);
	// Debug Helpers
	int ScanAllMailsInMbox();
	int ScanAllMailsInMbox_NewParser();
	//
	static BOOL MyCTimeToOleTime(MyCTime &ctimeDateTime, COleDateTime &coleDateTime);
	static BOOL OleTime2MyCTime(COleDateTime &coleDateTime, MyCTime &ctimeDateTime, BOOL roundUP);

	BOOL PrepopulateAdvancedSearchParams(CString *from, CString *to, CString *subject);
	void FindMinMaxTime(MyCTime &minTime, MyCTime &maxTime);
	void ResetFilterDates();

	void PostMsgCmdParamAttachmentHint();
	//
	
	static int ExportAsEmlFile(CFile *fpm, int mailPosition, CString &targetDirectory, CString &emlFile, CString &errorText);
	static int PrintAsEmlFile(CFile *fpm, int mailPosition, CString &emlFile);
	static int PrintMailAttachments(CFile *fpm, int mailPosition, AttachmentMgr &attachmentDB);
	static int DetermineAttachmentName(CFile *fpm, int mailPosition, MailBodyContent *body, SimpleString *bodyData, CStringW &nameW, AttachmentMgr &attachmentDB);
	static int PrintAttachmentNamesAsText2CSV(int mailPosition, SimpleString *outbuf, CString &characterLimit, CString &attachmentSeparator);
	// Called in CheckMatchAdvanced and CheckMatch
	static int FindAttachmentName(MboxMail *m, CString &searchString, BOOL bWholeWord, BOOL bCaseSensitive);
	//
	//////////////////////////////////////////////////////
	////////////  PDF
	//////////////////////////////////////////////////////
	int PrintMailConversationToSeparatePDF_Thread(int mailIndex, CString &errorText);
	int PrintMailConversationToSinglePDF_Thread(int mailIndex, CString &errorText);
	//
	// Range to Separate PDF
	int PrintMailRangeToSeparatePDF_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName);
	int PrintMailRangeToSeparatePDF_WorkerThread(int firstMail, int lastMail, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);
	//
	// Range to Single PDF
	int PrintMailRangeToSinglePDF_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName);
	int PrintMailRangeToSinglePDF_WorkerThread(int firstMail, int lastMail, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);
	//
	// Selected to Separate PDF
	int PrintMailSelectedToSeparatePDF_Thread(CString &targetPrintSubFolderName, CString &targetPrintFolderPath);
	int PrintMailSelectedToSeparatePDF_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);
	//
	// Selected to Single PDF
	int PrintMailSelectedToSinglePDF_Thread(CString &targetPrintSubFolderName, CString &targetPrintFolderPath);
	int PrintMailSelectedToSinglePDF_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);
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
	int PrintMailRangeToSeparateHTML_WorkerThread(int firstMail, int lastMail, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);
	//
	// Range to Single HTML
	int PrintMailRangeToSingleHTML_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName);
	int PrintMailRangeToSingleHTML_WorkerThread(int firstMail, int lastMail, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);
	//
	// Selected to Separate HTML
	int PrintMailSelectedToSeparateHTML_Thread(CString &targetPrintSubFolderName, CString &targetPrintFolderPath);
	int PrintMailSelectedToSeparateHTML_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);
	//
	// Selected to Single HTML
	int PrintMailSelectedToSingleHTML_Thread(CString &targetPrintSubFolderName, CString &targetPrintFolderPath);
	int PrintMailSelectedToSingleHTML_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);
	//
	//////////////////////////////////////////////////////
	////////////  HTML  END
	//////////////////////////////////////////////////////
	//
		//
	//////////////////////////////////////////////////////
	////////////  TEXT
	//////////////////////////////////////////////////////
	//
		// Selected to Separate TEXT
	//int PrintMailSelectedToSeparateTEXT_Thread(CString &targetPrintSubFolderName, CString &targetPrintFolderPath);
	//int PrintMailSelectedToSeparateTEXT_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);

	// Selected to Single HTML
	int PrintMailSelectedToSingleTEXT_Thread(CString &targetPrintSubFolderName, CString &targetPrintFolderPath);
	int PrintMailSelectedToSingleTEXT_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);
	///

	int CreateInlineImageCache_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName);
	//int CreateInlineImageCache_WorkerThread(int firstMail, int lastMail, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);
	//
	int CreateAttachmentCache_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName);
	//int CreateAttachmentCache_WorkerThread(int firstMail, int lastMail, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);
	//
	int CreateEmlCache_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName);
	//int CreateEmlCache_WorkerThread(int firstMail, int lastMail, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText);

	static BOOL IsSingleAddress(CString *to);
	static void TrimToAddr(CString *to, CString &toAddr, int maxNumbOfAddr);
	static void TrimToName(CString *to, CString &toName, int maxNumbOfAddr);
	static int DeleteAllHtmAndPDFFiles(CString &targetFolder);
	//static int DeleteAllHtmFiles(CString &targetFolder);
	//
	static int CreateInlineImageFiles(CFile &fpm, int mailPosition, CString &imageCachePath, bool runInvestigation = false);
	static int UpdateInlineSrcImgPath(char *inData, int indDataLen, SimpleString *outbuf, CListCtrl *attachments, int mailPosition, bool useMailPosition, bool runInvestigation = false);
	static int DetermineImageFileName(MboxMail *m, CString &cidName, CString &imageFilePath, MailBodyContent **foundBody, int mailPosition);
	//
	static int DetermineEmbeddedImages(char *inData, int indDataLen, int mailPosition);
	static int DetermineEmbeddedImages(char *inData, int indDataLen, MboxMail *m);
	//
	static int FindFilenameCount(std::vector <MailBodyContent*> &contentDetailsArray, CString &fileName);
	//
	static int CreateInlineImageFiles_SelectedItem(CMimeBody::CBodyList &bodies, NMsgView *pMsgView, int mailPosition, MailBodyInfoArray &cidArray, MyCArray<bool> &fileImgAlreadyCreatedArray, bool runInvestigation = false);
	static int UpdateInlineSrcImgPath_SelectedItem(char *inData, int indDataLen, SimpleString *outbuf, int mailPosition, bool useMailPosition, MailBodyInfoArray &cidArray);
	static int DetermineImageFileName_SelectedItem(CMimeBody::CBodyList &bodies, MboxMail *m, CString &cidName, CString &imageFilePath, CMimeBody **foundBody, MyCArray<bool> &fileImgAlreadyCreatedArray, int mailPosition);
	static int GetMailBody_SelectedItem(CMimeBody::CBodyList &bodies, CMimeBody** pBP);  // return body text type or 0-plain, 1-html, -1-not found
	static int FindFilenameCount(CMimeBody::CBodyList &bodies, CString &fileName);

	static int RemoveBackgroundColor(char *inData, int indDataLen, SimpleString *outbuf, int mailPosition);
	static int SetBackgroundColor(char *inData, int indDataLen, SimpleString *outbuf, int mailPosition);
	

	static int FindBodyTag(char *inData, int indDataLen, char *&tagBeg, int &tagDataLen);
	static int FindBodyBackgroundColor(char *inData, int indDataLen, char *&attribTag, int &attribTagLen, CString &bodyBackgroundColor, HtmlAttribInfo &bodyBackgroundColorAttrib);
	static int FindBodyTagAttrib(char *inData, int indDataLen, char *tag, int tagLen, /*out*/ char *&attribTag, int &attribTagLen, CString &attribVal, HtmlAttribInfo &bodyAttrib);
	static int RemoveBodyBackgroundColor(char *inData, int indDataLen, SimpleString *outbuf, CString &bodyBackgroundColor);
	static int RemoveBodyBackgroundColorAndWidth(char *inData, int indDataLen, SimpleString *outbuf, CString &bodyBackgroundColor, 
		CString &bodyWidth, BOOL removeBgColor, BOOL removeWidth);
	static int SetBodyWidth(char *inData, int indDataLen, SimpleString *outbuf, CString &bodyBackgroundColor);

	static int Color2Str(DWORD color, CString &colorStr);

	static BOOL loadImage(BYTE* pData, size_t nSize, CStringW &extensionW, CString &extension);
	static int DetermineListFileName(CString &fileName, CString &listFileName);
	void SetListFocus();

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
	afx_msg void OnGetDispInfo(NMHDR* pNMHDR, LRESULT* pResult);
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
	//afx_msg void OnMouseHover(UINT nFlags, CPoint point);
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	afx_msg LRESULT OnCmdParam_AttachmentHint(WPARAM wParam, LPARAM lParam);
	afx_msg void OnClose();
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
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
