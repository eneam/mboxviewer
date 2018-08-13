// WheelTreeCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "WheelTreeCtrl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
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
	return wnd->PostMessage(WM_MOUSEWHEEL, MAKELPARAM(nFlags, zDelta), MAKELPARAM(pt.x, pt.y));
}

#if 0
void CWheelTreeCtrl::OnNMRClick(NMHDR *pNMHDR, LRESULT *pResult)
{
	// TODO: Add your control notification handler code here
	// Added handle void NTreeView::OnRClick(NMHDR* pNMHDR, LRESULT* pResult)
	*pResult = 0;
}
#endif

