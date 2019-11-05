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
