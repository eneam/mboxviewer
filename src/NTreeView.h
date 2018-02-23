#if !defined(AFX_NTREEVIEW_H__28F2C044_02DA_4048_8F8F_3D17ADEE1421__INCLUDED_)
#define AFX_NTREEVIEW_H__28F2C044_02DA_4048_8F8F_3D17ADEE1421__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// NTreeView.h : header file
//

#include "WheelTreeCtrl.h"

/////////////////////////////////////////////////////////////////////////////
// NTreeView window

class NTreeView : public CWnd
{
	CFont		m_font;
	CImageList	m_il;
// Construction
public:
	NTreeView();
	DECLARE_DYNCREATE(NTreeView)
	CWheelTreeCtrl	m_tree;
	CMap<CString, LPCSTR, _int64, _int64>	fileSizes;

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(NTreeView)
	protected:
	virtual void PostNcDestroy();
	//}}AFX_VIRTUAL

// Implementation
public:
	void Traverse( HTREEITEM hItem, CFile &fp );
	void SaveData();
	void FillCtrl();
	void SelectMailFile(); // based on -MBOX= command line argument
	virtual ~NTreeView();

	// Generated message map functions
protected:
	//{{AFX_MSG(NTreeView)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg void OnUpdateFileRefresh(CCmdUI* pCmdUI);
	afx_msg void OnFileRefresh();
	//}}AFX_MSG
	afx_msg void OnSelchanged(NMHDR* pNMHDR, LRESULT* pResult);
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_NTREEVIEW_H__28F2C044_02DA_4048_8F8F_3D17ADEE1421__INCLUDED_)
