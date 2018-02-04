#if !defined(AFX_NLISTVIEW_H__DAE62ED6_932A_4145_B641_8CFD7B72EB2D__INCLUDED_)
#define AFX_NLISTVIEW_H__DAE62ED6_932A_4145_B641_8CFD7B72EB2D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// NListView.h : header file
//

#include "WheelListCtrl.h"

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
	BOOL m_bInFind;
	HTREEITEM m_which;
	void SelectPos(_int64 offset);
	void SelectItem(int which);
	_int64 DoFind(_int64 searchstart, _int64 searchend);
	int DoFastFind(int searchstart);
	CString m_searchString;
	_int64		m_searchPos;
	_int64 m_startoff, m_endoff;
	BOOL	m_filterDates;
	BOOL	m_bCaseSens;
	BOOL	m_bWholeWord;
	CTime m_lastStartDate;
	CTime m_lastEndDate;
	int m_lastScope;
	void ClearDescView();
	CString m_curFile;
	int m_lastSel;
	void FillCtrl();
	CString m_path;
	virtual ~NListView();
	void SelectItemFound(int iItem);

	int WhichOne( _int64 offset, int hint = -1, int lowhint = 0, int highhint = -1 );
	void ResetFont();
	void RedrawMails();

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
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnEditVieweml();
	afx_msg void OnUpdateEditVieweml(CCmdUI *pCmdUI);
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_NLISTVIEW_H__DAE62ED6_932A_4145_B641_8CFD7B72EB2D__INCLUDED_)
