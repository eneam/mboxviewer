#if !defined(AFX_NMSGVIEW_H__A0D2564F_34F6_4D24_9054_C7FD35135480__INCLUDED_)
#define AFX_NMSGVIEW_H__A0D2564F_34F6_4D24_9054_C7FD35135480__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// NMsgView.h : header file
//

#include "Browser.h"
/////////////////////////////////////////////////////////////////////////////
// NMsgView window

class NMsgView : public CWnd
{
	CFont m_BoldFont, m_NormFont, m_BigFont;
// Construction
public:
	NMsgView();
	DECLARE_DYNCREATE(NMsgView)
	CBrowser m_browser;
	CListCtrl	m_attachments;
	CFont m_font;
	CString m_strTitle1, m_strTitle2, m_strTitle3, m_strDescp1, m_strDescp2, m_strDescp3;
// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(NMsgView)
	protected:
	virtual void PostNcDestroy();
	//}}AFX_VIRTUAL

// Implementation
public:
	BOOL m_bMax;
	CRect m_rcCaption;
	void UpdateLayout();
	int m_nAttachSize;
	BOOL m_bAttach;
	virtual ~NMsgView();

	// Generated message map functions
protected:
	//{{AFX_MSG(NMsgView)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnPaint();
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
	//}}AFX_MSG
	afx_msg void OnActivating(NMHDR* pNMHDR, LRESULT* pResult);
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_NMSGVIEW_H__A0D2564F_34F6_4D24_9054_C7FD35135480__INCLUDED_)
