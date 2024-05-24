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

// WheelTreeCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "WheelTreeCtrl.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

/////////////////////////////////////////////////////////////////////////////
// CWheelTreeCtrl

CWheelTreeCtrl::CWheelTreeCtrl()
{
	
}

CWheelTreeCtrl::~CWheelTreeCtrl()
{
}


BEGIN_MESSAGE_MAP(CWheelTreeCtrl, CTreeCtrl)
	//{{AFX_MSG_MAP(CWheelTreeCtrl)
	ON_WM_MOUSEWHEEL()
	//}}AFX_MSG_MAP
	//ON_NOTIFY_REFLECT(NM_RCLICK, &CWheelTreeCtrl::OnNMRClick)
	ON_WM_ERASEBKGND()
	ON_WM_SETFOCUS()
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CWheelTreeCtrl message handlers

BOOL CWheelTreeCtrl::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt)
{
	CWnd *wnd = WindowFromPoint(pt);
	if (wnd == NULL)
		return FALSE;
	if (wnd == this) {
		return CTreeCtrl::OnMouseWheel(nFlags, zDelta, pt);
	}
	if ((GetKeyState(VK_CONTROL) & 0x80) == 0) { // if CTRL key not Down; Do we need to post msg further anyway
		// Commented out, it freezes mbox viewer and and even IE for few seconds when CTRL/SHIFT/etc key are kept down
		; //return wnd->PostMessage(WM_MOUSEWHEEL, MAKELPARAM(nFlags, zDelta), MAKELPARAM(pt.x, pt.y));
	}
	return TRUE;
}

#if 0
void CWheelTreeCtrl::OnNMRClick(NMHDR *pNMHDR, LRESULT *pResult)
{
	// TODO: Add your control notification handler code here
	// Added handle void NTreeView::OnRClick(NMHDR* pNMHDR, LRESULT* pResult)
	*pResult = 0;
}
#endif


BOOL CWheelTreeCtrl::OnEraseBkgnd(CDC* pDC)
{
	//BOOL ret = CWnd::OnEraseBkgnd(pDC);

	return FALSE;
}

#if 1

void CWheelTreeCtrl::OnSetFocus(CWnd* pOldWnd)
{
	TRACE(L"CWheelTreeCtrl::OnSetFocus\n");

	CTreeCtrl::OnSetFocus(pOldWnd);
	CmboxviewApp::wndFocus = this;

	int deb = 1;
	// TODO: Add your message handler code here
}
#endif
