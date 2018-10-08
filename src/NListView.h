#if !defined(AFX_NLISTVIEW_H__DAE62ED6_932A_4145_B641_8CFD7B72EB2D__INCLUDED_)
#define AFX_NLISTVIEW_H__DAE62ED6_932A_4145_B641_8CFD7B72EB2D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// NListView.h : header file
//

#include "WheelListCtrl.h"
#include "UPDialog.h"

__int64 FileSeek(HANDLE hf, __int64 distance, DWORD MoveMethod);

class CMimeMessage;
class MboxMail;

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
	static const CUPDUPDATA * pCUPDUPData; 

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
	BOOL SetupFileMapView(_int64 offset, DWORD length);
	int CheckMatch(int which, CString &searchString);
	BOOL FindInMailContent(int mailPosition, BOOL bContent, BOOL bAttachment);
	void CloseMailFile();
	void ResetFileMapView();
	void SortByColumn(int colNumber);
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
	int DoFastFind(int searchstart, BOOL mainThreadContext, int maxSearchDuration);
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
	void ClearDescView();
	CString m_curFile;
	int m_lastSel;
	int m_bStartSearchAtSelectedItem;
	int m_gmtTime;
	CString m_path;
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
	void PrintMailGroupToText(int iItem, int textType, BOOL forceOpen = FALSE);
	int MailsWhichColumnSorted() const;

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
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnEditVieweml();
	afx_msg void OnUpdateEditVieweml(CCmdUI *pCmdUI);
	//afx_msg void OnViewConversations();
	//fx_msg void OnUpdateViewConversations(CCmdUI *pCmdUI);
};

typedef struct _ParseArgs {
	CString path;
	BOOL exitted;
} PARSE_ARGS;

typedef struct _FindArgs {
	int searchstart;
	int retpos;
	BOOL exitted;
	NListView *lview;
} FIND_ARGS;

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_NLISTVIEW_H__DAE62ED6_932A_4145_B641_8CFD7B72EB2D__INCLUDED_)
