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

class SimpleString;

class NMsgView : public CWnd
{
	CFont m_BoldFont, m_NormFont, m_BigFont, m_BigBoldFont;
// Construction
public:
	NMsgView();
	DECLARE_DYNCREATE(NMsgView)
	CBrowser m_browser;
	CListCtrl	m_attachments;
	CFont m_font;
	CString m_strTitleSubject, m_strTitleFrom, m_strTitleDate, m_strTitleTo, m_strTitleBody;
	CString m_strSubject, m_strFrom, m_strDate, m_strTo, m_strBody;
	UINT m_subj_charsetId, m_from_charsetId, m_date_charsetId, m_to_charsetId, m_body_charsetId;
	UINT m_cnf_subj_charsetId, m_cnf_from_charsetId, m_cnf_date_charsetId, m_cnf_to_charsetId;
	CString m_subj_charset, m_from_charset, m_date_charset, m_to_charset, m_body_charset;
	int m_show_charsets;
	int m_bImageViewer;
// Attributes
public:

// Operations
public:
protected:
	int PaintHdrField(CPaintDC &dc, CRect	&r, int x_pos, int y_pos, BOOL bigFont, CString &FieldTitle, CString &FieldText, CString &Charset, UINT CharsetId, UINT CnfCharsetId);

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(NMsgView)
	protected:
	virtual void PostNcDestroy();
	//}}AFX_VIRTUAL

// Implementation
public:
	void FindStringInIHTMLDocument(CString &searchText, BOOL matchWord, BOOL matchCase);
	void ClearSearchResultsInIHTMLDocument(CString searchID);
	static void GetTextFromIHTMLDocument(SimpleString *inbuf, SimpleString *workbuf, UINT inCodePage, UINT outCodePage);
	static BOOL CreateHTMLDocument(struct IHTMLDocument2 **lpDocument, SimpleString *inbuf, SimpleString *workbuf, UINT inCodepage);
	BOOL FindElementByTagInIHTMLDocument(struct IHTMLDocument2 *lpDocument, struct IHTMLElement **ppvEl, CString &tag);
	void PrintIHTMLDocument(struct IHTMLDocument2 *lpDocument);
	void PrintIHTMLElement(struct IHTMLElement *lpElm, CStringW &text);
	static void RemoveStyleTagFromIHTMLDocument(struct IHTMLElement *lpElm);
	static void MergeWhiteLines(SimpleString *inbuf, int maxoutLines);
	BOOL m_bMax;
	CRect m_rcCaption;
	void UpdateLayout();
	int m_nAttachSize;
	BOOL m_bAttach;
	CString m_searchID;
	CString m_matchStyle;
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
	afx_msg void OnDoubleClick(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnRClick(NMHDR* pNMHDR, LRESULT* pResult);
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_NMSGVIEW_H__A0D2564F_34F6_4D24_9054_C7FD35135480__INCLUDED_)
