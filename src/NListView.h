#if !defined(AFX_NLISTVIEW_H__DAE62ED6_932A_4145_B641_8CFD7B72EB2D__INCLUDED_)
#define AFX_NLISTVIEW_H__DAE62ED6_932A_4145_B641_8CFD7B72EB2D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// NListView.h : header file
//

#include "WheelListCtrl.h"
#include "UPDialog.h"
#include "FindAdvancedDlg.h"
#include "FindDlg.h"

__int64 FileSeek(HANDLE hf, __int64 distance, DWORD MoveMethod);
void CPathStripPath(const char *path, CString &fileName);
BOOL CPathGetPath(const char *path, CString &filePath);

class CMimeMessage;
class MboxMail;
class MyMailArray;
class SerializerHelper;
class SimpleString;


typedef CArray<int, int> MailIndexList;

/////////////////////////////////////////////////////////////////////////////
// NListView window

class NListView : public CWnd
{
	CWheelListCtrl	m_list;
	CFont		m_font;
	CFont m_boldFont;

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
	int CheckMatch(int which, CString &searchString);
	int CheckMatchAdvanced(int i, CFindAdvancedParams &params);
	BOOL FindInMailContent(int mailPosition, BOOL bContent, BOOL bAttachment);
	BOOL AdvancedFindInMailContent(int mailPosition, BOOL bContent, BOOL bAttachment, CFindAdvancedParams &params);
	void CloseMailFile();
	void ResetFileMapView();
	void SortByColumn(int colNumber, BOOL sortByPosition = FALSE);
	void RefreshSortByColumn();
	// end of vars

	// For single Mail. TODO: Should consider to define structure to avoid confusing var names
	CString m_searchStringInMail;
	BOOL m_bCaseSensInMail;
	BOOL m_bWholeWordInMail;

	BOOL m_bExportEml;
	BOOL m_bFindNext;
	BOOL m_bInFind;
	HTREEITEM m_which;
	void SelectItem(int which);
	int DoFastFind(int searchstart, BOOL mainThreadContext, int maxSearchDuration, BOOL findAll);
	CString m_searchString;
	int m_lastFindPos;
	int m_maxSearchDuration;
	BOOL m_bEditFindFirst;

	BOOL	m_filterDates;
	BOOL	m_bCaseSens;
	BOOL	m_bWholeWord;
	CTime m_lastStartDate;
	CTime m_lastEndDate;
	BOOL m_bFrom;
	BOOL m_bTo;
	BOOL m_bSubject;
	BOOL m_bContent;
	BOOL m_bAttachments;
	BOOL m_bHighlightAll;
	BOOL m_bFindAll;
	BOOL m_bHighlightAllSet;

	//struct CFindDlgParams m_findParams;  // TODO: review later, requires too many chnages for now
	struct CFindAdvancedParams m_advancedParams;
	CString m_stringWithCase[7];
	BOOL m_advancedFind;

	void ClearDescView();
	CString m_curFile;
	int m_lastSel;
	int m_bStartSearchAtSelectedItem;
	int m_gmtTime;
	CString m_path;
	int m_findAllCount;
	// Used in Custom Draw
	SimpleString *m_name;
	SimpleString *m_addr;
	BOOL m_bLongMailAddress;
	//
	void FillCtrl();
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
	int SaveAsMboxFile();
	int ReloadMboxFile();
	int PopulateUserMailArray(SerializerHelper &sz, int mailListCnt, BOOL verifyOnly);
	int OpenArchiveFileLocation();
	int RemoveDuplicateMails();
	BOOL IsUserSelectedMailListEmpty();

	MailIndexList * PopulateSelectedMailsList();
	void FindFirstAndLastMailOfConversation(int iItem, int &firstMail, int &lastMail);
	int RefreshMailsOnUserSelectsMailListMark();
	int VerifyMailsOnUserSelectsMailListMarkCounts();
	//
	int ExportMailGroupToSeparatePDF(int firstMail, int lastMail, BOOL multipleSelectedMails, int nItem);
	int ExportMailGroupToSeparateHTML(int firstMail, int lastMail, BOOL multipleSelectedMails, int nItem);
	void PrintMailGroupToText(BOOL multipleSelectedMails, int iItem, int textType, BOOL forceOpen = FALSE, BOOL printToPrinter = FALSE, BOOL createFileOnly = FALSE);
	int PrintMailRangeToSingleCSV_Thread(int iItem);
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

	static void TrimToAddr(CString *to, CString &toAddr, int maxNumbOfAddr);
	static int DeleteAllHtmAndPDFFiles(CString &targetFolder);

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


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_NLISTVIEW_H__DAE62ED6_932A_4145_B641_8CFD7B72EB2D__INCLUDED_)
