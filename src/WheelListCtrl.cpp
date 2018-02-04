// WheelListCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "WheelListCtrl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CWheelListCtrl

CWheelListCtrl::CWheelListCtrl()
{
}

CWheelListCtrl::~CWheelListCtrl()
{
}


BEGIN_MESSAGE_MAP(CWheelListCtrl, CListCtrl)
	//{{AFX_MSG_MAP(CWheelListCtrl)
	ON_WM_MOUSEWHEEL()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CWheelListCtrl message handlers

BOOL CWheelListCtrl::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt) 
{
	CWnd *wnd = WindowFromPoint(pt);
	if( wnd == NULL )
		return FALSE;
	if( wnd == this ) {
		return CListCtrl::OnMouseWheel(nFlags, zDelta, pt);
	}
	return wnd->PostMessage(WM_MOUSEWHEEL, MAKELPARAM(nFlags,zDelta), MAKELPARAM(pt.x, pt.y));
}
