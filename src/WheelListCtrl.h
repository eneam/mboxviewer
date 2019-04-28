#if !defined(AFX_WHEELLISTCTRL_H__CEACCFA6_7FB8_495E_867C_B81C5112F5A6__INCLUDED_)
#define AFX_WHEELLISTCTRL_H__CEACCFA6_7FB8_495E_867C_B81C5112F5A6__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// WheelListCtrl.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CWheelListCtrl window

class NListView;

class CWheelListCtrl : public CListCtrl
{
// Construction
public:
	CWheelListCtrl(const NListView *cListCtrl);

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CWheelListCtrl)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CWheelListCtrl();

	virtual COLORREF OnGetCellBkColor(int /*nRow*/, int /*nColum*/);

	// Generated message map functions
protected:
	const NListView *m_list;
	//{{AFX_MSG(CWheelListCtrl)
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnDrawItem(int nIDCtl, LPDRAWITEMSTRUCT lpDrawItemStruct);
	//afx_msg void OnMouseHover(UINT nFlags, CPoint point);
	//afx_msg void OnSetFocus(CWnd* pOldWnd);
	//afx_msg int OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message);
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_WHEELLISTCTRL_H__CEACCFA6_7FB8_495E_867C_B81C5112F5A6__INCLUDED_)
