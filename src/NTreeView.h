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
	BOOL m_bSelectMailFileDone;
	int m_timerTickCnt;
	// timer
	UINT_PTR m_nIDEvent;
	UINT m_nElapse;

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
	void ForceParseMailFile(HTREEITEM hItem);
	void UpdateFileSizesTable(CString &path, _int64 fSize);
	virtual ~NTreeView();
	HTREEITEM FindItem(HTREEITEM hItem, CString &mailFileName);
	void StartTimer();

	static void FindAllDirs(LPCTSTR pstr);

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
	afx_msg void OnSelchanging(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnRClick(NMHDR* pNMHDR, LRESULT* pResult);
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnTimer(UINT_PTR nIDEvent);
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_NTREEVIEW_H__28F2C044_02DA_4048_8F8F_3D17ADEE1421__INCLUDED_)
