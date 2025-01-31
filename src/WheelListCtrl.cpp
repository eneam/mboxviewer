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

// WheelListCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "WheelListCtrl.h"
#include "NListView.h"
#include "ResHelper.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

/////////////////////////////////////////////////////////////////////////////
// CWheelListCtrl

CWheelListCtrl::CWheelListCtrl(const NListView *cListCtrl)
{
	m_list = cListCtrl;
	int deb = 1;
}

CWheelListCtrl::~CWheelListCtrl()
{
	int deb = 1;
}


BEGIN_MESSAGE_MAP(CWheelListCtrl, CListCtrl)
	//{{AFX_MSG_MAP(CWheelListCtrl)
	ON_WM_MOUSEWHEEL()
	//}}AFX_MSG_MAP
	ON_WM_DRAWITEM()
	//ON_WM_MOUSEHOVER()
	ON_WM_SETFOCUS()
	//ON_WM_MOUSEACTIVATE()
	ON_MESSAGE(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, &CWheelListCtrl::OnCmdParam_OnSwitchWindow)
	ON_WM_CREATE()
	//ON_NOTIFY_EX(TTN_NEEDTEXT, 0, &CWheelListCtrl::OnTtnNeedText)
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
	if ((GetKeyState(VK_CONTROL) & 0x80) == 0) { // if CTRL key not Down; Do we need to post msg further anyway
		// Commented out, it freezes mbox viewer and and even IE for few seconds when CTRL/SHIFT/etc key are kept down
		; // return wnd->PostMessage(WM_MOUSEWHEEL, MAKELPARAM(nFlags, zDelta), MAKELPARAM(pt.x, pt.y));
	}
	return TRUE;
}


COLORREF CWheelListCtrl::OnGetCellBkColor(int /*nRow*/, int /*nColum*/) 
{ 
	return GetBkColor(); 
}

void CWheelListCtrl::OnDrawItem(int nIDCtl, LPDRAWITEMSTRUCT lpDrawItemStruct)
{
	// TODO: Add your message handler code here and/or call default

	//CListCtrl::OnDrawItem(nIDCtl, lpDrawItemStruct);

	CDC dc;
	BOOL ret = dc.Attach(lpDrawItemStruct->hDC);

	CRect rect = lpDrawItemStruct->rcItem;

	DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailSummaryTitles);
	dc.FillRect(&rect, &CBrush(color));
	dc.SetBkMode(TRANSPARENT);

	dc.Detach();

	if ((lpDrawItemStruct->itemID >= 0) && (lpDrawItemStruct->itemID <= 5))
	{
		CStringW txtW;
		if (lpDrawItemStruct->itemID == 0)
			txtW = L"!";
		else if (lpDrawItemStruct->itemID == 1) {
			txtW = L"date";
			if (m_list->m_gmtTime == 1)
				txtW += L" (GMT)";
			else
				txtW += L" (Local)";
		}
		if (lpDrawItemStruct->itemID == 2)
			txtW = L"from";
		if (lpDrawItemStruct->itemID == 3)
			txtW = L"to";
		if (lpDrawItemStruct->itemID == 4)
			txtW = L"subject";
		if (lpDrawItemStruct->itemID == 5)
			txtW = L"size(KB)";

		CString newText;
		BOOL retFind = ResHelper::DetermineString(txtW, newText);
		if (!newText.IsEmpty())
			txtW = newText;

		txtW.Append(L" ");

		int x_offset = 6;
		int xpos = rect.left + x_offset;
		int ypos = rect.top + 6;

		HDC hDC = lpDrawItemStruct->hDC;

		int which_sorted = m_list->MailsWhichColumnSorted();
		if ((lpDrawItemStruct->itemID == abs(which_sorted)) || ((lpDrawItemStruct->itemID == 0) && (abs(which_sorted) == 99)))
		{
			if (abs(which_sorted) == 5)
			{
				if (which_sorted > 0)
					txtW += L'\x2191';
				else
					txtW += L'\x2193';
			}
			else
			{
				if (which_sorted < 0)
					txtW += L'\x2191';
				else
					txtW += L'\x2193';
			}
		}

		::ExtTextOutW(hDC, xpos, ypos, ETO_CLIPPED, &rect, (LPCWSTR)txtW, txtW.GetLength(), NULL);

		BOOL ret = dc.Attach(lpDrawItemStruct->hDC);
		int nsave = dc.SaveDC();

		CRect rect = lpDrawItemStruct->rcItem;

		xpos = rect.left + rect.Width();
		ypos = rect.top;

		CPen penRed(PS_SOLID, 1, RGB(212, 208, 200));  // gray
		dc.SelectObject(&penRed);

		dc.MoveTo(xpos - 1, ypos);
		dc.LineTo(xpos - 1, rect.bottom);

		xpos = rect.left;
		ypos = rect.top + rect.Height() - 1;

		dc.MoveTo(xpos, ypos);
		//dc.LineTo(rect.right, ypos);  // comment out to draw horizontal line


		dc.RestoreDC(nsave);
		dc.Detach();
	}

	CListCtrl::OnDrawItem(nIDCtl, lpDrawItemStruct);

	int deb = 1;
}

int CWheelListCtrl::PreTranslateMessage(MSG* pMsg)
{
	return CWnd::PreTranslateMessage(pMsg);
}

void CWheelListCtrl::PreSubclassWindow()
{
	// Doesn't work. Probably since we do customdraw
	_ASSERT(1);
	//SetExtendedStyle(LVS_EX_INFOTIP | LVS_EX_LABELTIP | LVS_EX_DOUBLEBUFFER | GetExtendedStyle());
	CListCtrl::PreSubclassWindow();
	SetExtendedStyle(LVS_EX_INFOTIP | LVS_EX_LABELTIP | LVS_EX_DOUBLEBUFFER | GetExtendedStyle());
}


#if 0
// Didn't work., more to investiagte
void CWheelListCtrl::OnMouseHover(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default

	CWnd *wnd = WindowFromPoint(point);
	if (wnd == NULL)
		return;
	if (wnd == this) {
		return CListCtrl::OnMouseHover(nFlags, point);
	}
	if ((GetKeyState(VK_CONTROL) & 0x80) == 0) { // if CTRL key not Down; Do we need to post msg further anyway
		// Commented out, it freezes mbox viewer and and even IE for few seconds when CTRL/SHIFT/etc key are kept down
		; // return wnd->PostMessage(WM_MOUSEWHEEL, MAKELPARAM(nFlags, zDelta), MAKELPARAM(pt.x, pt.y));
	}

	CListCtrl::OnMouseHover(nFlags, point);
}

#endif

#if 1

void CWheelListCtrl::OnSetFocus(CWnd* pOldWnd)
{
	static int id = 0;

	TRACE(L"CWheelListCtrl::OnSetFocus id=%d\n", id);

	CListCtrl::OnSetFocus(pOldWnd);
	CmboxviewApp::wndFocus = this;
	id++;

	// TODO: Add your message handler code here
}
#endif

#if 0
int CWheelListCtrl::OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message)
{
	// TODO: Add your message handler code here and/or call default

	return CListCtrl::OnMouseActivate(pDesktopWnd, nHitTest, message);
}

#endif

LRESULT CWheelListCtrl::OnCmdParam_OnSwitchWindow(WPARAM wParam, LPARAM lParam)
{
	TRACE(L"CWheelListCtrl::OnCmdParam_OnSwitchWindow\n");
	CWnd* wnd = CMainFrame::SetWindowFocus(this);

	return 0;
}


int CWheelListCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CListCtrl::OnCreate(lpCreateStruct) == -1)
		return -1;

	// TODO:  Add your specialized creation code here

	// Doesn't work. Likely due to Custom Draw
	// Need to investigate when time permits
	//BOOL retA = ResHelper::ActivateToolTips(this, m_toolTip);
	//EnableToolTips();

	return 0;
}

// Doesn't work. Likely due to Custom Draw
// Need to investigate when time permits
BOOL CWheelListCtrl::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	NListView* list = (NListView*)m_list;
	BOOL bRet = list->OnToolNeedText(id, pNMHDR, pResult);
	//BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}

BOOL CWheelListCtrl::OnToolNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	CPoint pt(GetMessagePos());
	ScreenToClient(&pt);

	int nRow = -1, nCol = -1;
	//CellHitTest(pt, nRow, nCol);

	CString tooltip = GetToolTipText(nRow, nCol);
	if (tooltip.IsEmpty())
		return FALSE;

	// Non-unicode applications can receive requests for tooltip-text in unicode
	TOOLTIPTEXTA* pTTTA = (TOOLTIPTEXTA*)pNMHDR;
	TOOLTIPTEXTW* pTTTW = (TOOLTIPTEXTW*)pNMHDR;
#ifndef _UNICODE
	if (pNMHDR->code == TTN_NEEDTEXTA)
		lstrcpyn(pTTTA->szText, static_cast<LPCWSTR>(tooltip), sizeof(pTTTA->szText));
	else
		mbstowcs(pTTTW->szText, static_cast<LPCWSTR>(tooltip), sizeof(pTTTW->szText) / sizeof(WCHAR));
#else
	if (pNMHDR->code == TTN_NEEDTEXTA)
		_wcstombsz(pTTTA->szText, static_cast<LPCWSTR>(tooltip), sizeof(pTTTA->szText));
	else
		lstrcpyn(pTTTW->szText, static_cast<LPCWSTR>(tooltip), sizeof(pTTTW->szText) / sizeof(WCHAR));
#endif
	// If wanting to display a tooltip which is longer than 80 characters,
	// then one must allocate the needed text-buffer instead of using szText,
	// and point the TOOLTIPTEXT::lpszText to this text-buffer.
	// When doing this, then one is required to release this text-buffer again
	return TRUE;
}

bool CWheelListCtrl::ShowToolTip(const CPoint& pt) const
{
	// Lookup up the cell
	int nRow = -1, nCol = -1;
	//CellHitTest(pt, nRow, nCol);

	if (nRow != -1 && nCol != -1)
		return true;
	else
		return false;
}

CString CWheelListCtrl::GetToolTipText(int nRow, int nCol)
{
	if (nRow != -1 && nCol != -1)
		return CString("");
		//return m_list->m_list.GetItemText(nRow, nCol);	// Cell-ToolTip
	else
		return CString("");
}
#if 0
void CWheelListCtrl::OnTvnGetInfoTip(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMTVGETINFOTIP pGetInfoTip = (LPNMTVGETINFOTIP)pNMHDR;

	HTREEITEM hItem = pGetInfoTip->hItem;
	if (hItem)
	{
#if 0
		DWORD nId = (DWORD)m_tree.GetItemData(hItem);
		LabelInfo* linfo = m_labelInfoStore.Find(nId);
		if (linfo && ((linfo->m_nodeType == LabelInfo::MailFolder) || (linfo->m_nodeType == LabelInfo::MailSubFolder)))
		{
			CString csItemTxt = m_tree.GetItemText(hItem);
			CString folderPath = linfo->m_mailFolderPath;
			int len = folderPath.GetLength();
			if (len > (pGetInfoTip->cchTextMax - 2))
				len = pGetInfoTip->cchTextMax - 2;

			_tcsncpy(pGetInfoTip->pszText, folderPath, len);
			pGetInfoTip->pszText[len] = 0;
		}
#endif
		int deb = 1;
	}
	*pResult = 0;
}
#endif

