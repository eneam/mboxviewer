// WheelListCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "WheelListCtrl.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
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
	//ON_NOTIFY_REFLECT(NM_RCLICK, &CWheelListCtrl::OnNMRClick)
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


 COLORREF CWheelListCtrl::OnGetCellBkColor(int /*nRow*/, int /*nColum*/) 
 { 
	 return GetBkColor(); 
 }

void CWheelListCtrl::OnNMRClick(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: Add your control notification handler code here
#define ID_POPUPMENU 0



	CMenu menu;
	CPoint pt;
	::GetCursorPos(&pt);
	CWnd *wnd = WindowFromPoint(pt);

	menu.CreatePopupMenu();
	menu.AppendMenu(MF_STRING, ID_POPUPMENU, _T("Not Current"));
	menu.AppendMenu(MF_STRING, ID_POPUPMENU + 1, _T("Current"));
	menu.AppendMenu(MF_SEPARATOR, ID_POPUPMENU + 2);
	menu.AppendMenu(MF_STRING, ID_POPUPMENU + 3, _T("Delete"));
	int res = menu.TrackPopupMenu(TPM_LEFTALIGN|TPM_RIGHTBUTTON|TPM_RETURNCMD, pt.x, pt.y, this);

	UINT nFlags = MF_BYPOSITION;
	CString menuString;
	BOOL ret = menu.GetMenuString(res, menuString, nFlags);

	*pResult = 0;
}
