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

// ChildView.cpp : implementation of the CChildView class
//

#include "stdafx.h"
#include "mboxview.h"
#include "mainfrm.h"
#include "Profile.h"
#include "ChildView.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

/////////////////////////////////////////////////////////////////////////////
// CChildView

CChildView::CChildView(int msgViewPosition)
{
	m_msgViewPosition = msgViewPosition;
}

CChildView::~CChildView()
{
}


BEGIN_MESSAGE_MAP(CChildView,CWnd )
	//{{AFX_MSG_MAP(CChildView)
	ON_WM_PAINT()
	//}}AFX_MSG_MAP
	ON_WM_SIZE()
	ON_WM_CREATE()
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CChildView message handlers

BOOL CChildView::PreCreateWindow(CREATESTRUCT& cs) 
{
	if (!CWnd::PreCreateWindow(cs))
		return FALSE;

	cs.dwExStyle &= ~WS_EX_CLIENTEDGE;
	cs.style &= ~WS_BORDER;	
	cs.style |= WS_CHILD;	
//	cs.dwExStyle |= WS_EX_CLIENTEDGE;
//	cs.style &= ~WS_BORDER;
//	cs.lpszClass = AfxRegisterWndClass(CS_HREDRAW|CS_VREDRAW|CS_DBLCLKS, 
//		::LoadCursor(NULL, IDC_ARROW), HBRUSH(COLOR_WINDOW+1), NULL);

	return TRUE;
}

void CChildView::OnPaint() 
{
	CPaintDC dc(this); // device context for painting
	
	// TODO: Add your message handler code here
	
	// Do not call CWnd::OnPaint() for painting messages
}

int CChildView::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CWnd::OnCreate(lpCreateStruct) == -1)
		return -1;

	BOOL ret;

	if (!m_verSplitter.CreateStatic(this, 1, 2, WS_CHILD | WS_VISIBLE, AFX_IDW_PANE_FIRST)) {
		TRACE("Failed to create first splitter\n");
		return -1;
	}

	if (m_msgViewPosition == 1)
		ret = m_horSplitter.CreateStatic(&m_verSplitter, 2, 1, WS_CHILD | WS_VISIBLE, m_verSplitter.IdFromRowCol(0, 1));
	else 
		ret = m_horSplitter.CreateStatic(&m_verSplitter, 1, 2, WS_CHILD | WS_VISIBLE, m_verSplitter.IdFromRowCol(0, 1));

	if (ret == FALSE) {
		TRACE("Failed to create second splitter\n");
		return -1;
	}

	CRect frameRect;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame) 
	{
		pFrame->GetClientRect(frameRect);
		int deb = 1;
	}
	// frameSize is not correct yet. 
	// SetWindowPlacement(&wpr); is called in CMainFrame::OnCreate but executed after CChildView::OnCreate
	// Need to find better solution to windows placement
	CSize frameSize = frameRect.Size();

	CString m_section = CString(sz_Software_mboxview) + "\\" + "Window Placement";
#if 1
	WINDOWPLACEMENT wpr;
	DWORD cb = sizeof(wpr);
	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
		ret = CProfile::_GetProfileBinary(HKEY_CURRENT_USER, m_section, "MainFrame_EmlPreviewMode", (LPBYTE)&wpr, cb);
	else
		ret = CProfile::_GetProfileBinary(HKEY_CURRENT_USER, m_section, "MainFrame", (LPBYTE)&wpr, cb);
	int cx = -1;
	int cy = -1;
	if (ret && (cb == sizeof(wpr)))
	{
		cx = wpr.rcNormalPosition.right - wpr.rcNormalPosition.left;
		cy = wpr.rcNormalPosition.bottom - wpr.rcNormalPosition.top;
		int deb = 1;
	}
	//frameSize.cx = cx;
	//frameSize.cy = cy;

#endif

	CRect rect;
	GetClientRect(rect);
	CSize treeSize = rect.Size();

	treeSize.cx = 177;
	treeSize.cy = 200;

	int tree_frameCx = 177;
	int tree_frameCy = 200;
	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "TreeFrameWidth", tree_frameCx);
	
	if (ret)
		treeSize.cx = tree_frameCx;

	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "TreeFrameHeight", tree_frameCy);
	if (ret)
		treeSize.cy = tree_frameCy;


	BOOL bEmlPrewviewMode = CMainFrame::m_commandLineParms.m_bEmlPreviewMode;

	if (bEmlPrewviewMode)
		treeSize.cx = 0;

	if (!m_verSplitter.CreateView(0, 0, RUNTIME_CLASS(NTreeView), treeSize, NULL)) {
		TRACE("Failed to create Tree left view\n");
		return -1;
	}

	//
	CSize msgSize;

	int msg_frameCx_TreeNotInHide = 700;
	int msg_frameCy_TreeNotInHide = 200;

	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "MsgFrameTreeNotHiddenWidth", msg_frameCx_TreeNotInHide);
	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "MsgFrameTreeNotHiddenHeight", msg_frameCy_TreeNotInHide);

	if (m_msgViewPosition == 1)   // windows on top
	{
		msgSize.cx = frameSize.cx - treeSize.cx;
		msgSize.cy = msg_frameCy_TreeNotInHide;
	}
	else if (m_msgViewPosition == 2)   // windows on right
	{
		msgSize.cx = msg_frameCx_TreeNotInHide;
		msgSize.cy = 200;
	}
	else if (m_msgViewPosition == 3)   // windows on left
	{
		msgSize.cx = msg_frameCx_TreeNotInHide;
		if (bEmlPrewviewMode)
			msgSize.cx = frameSize.cx*2;
		msgSize.cy = 200;
	}

	if (msgSize.cx < 0)
		int deb = 1;
	if (msgSize.cy < 0)
		int deb = 1;
	//

	CSize listSize;

	int list_frameCx_TreeNotInHide = 700;
	int list_frameCy_TreeNotInHide = 200;

	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "ListFrameTreeNotHiddenWidth", list_frameCx_TreeNotInHide);;
	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "ListFrameTreeNotHiddenHeight", list_frameCy_TreeNotInHide);;

	if (m_msgViewPosition == 1)   // windows on top
	{
		listSize.cx = frameSize.cx - treeSize.cx;
		listSize.cy = list_frameCy_TreeNotInHide;
		if (bEmlPrewviewMode)
			listSize.cy = 0;
	}
	else if (m_msgViewPosition == 2)   // windows on right
	{
		listSize.cx = list_frameCx_TreeNotInHide;
		if (bEmlPrewviewMode)
			listSize.cx = 0;
		listSize.cy = 200;
	}
	else if (m_msgViewPosition == 3)   // windows on left
	{
		listSize.cx = list_frameCx_TreeNotInHide;
		if (bEmlPrewviewMode)
			listSize.cx = 0;
		listSize.cy = 200;
	}

	if (listSize.cx < 0)
		int deb = 1;
	if (listSize.cy < 0)
		int deb = 1;

	if (bEmlPrewviewMode == FALSE)
	{
		if (m_msgViewPosition > 1)
		{
			if (msgSize.cx < 64)
			{
				msgSize.cx = 64;
				listSize.cx -= 60;
			}
			else if (listSize.cx < 64)
			{
				listSize.cx = 20;
				msgSize.cx -= 60;
			}
		}
		else
		{
			if (msgSize.cy < 64)
			{
				msgSize.cy = 64;
				listSize.cy -= 60;
			}
			else if (listSize.cx < 64)
			{
				listSize.cy = 20;
				msgSize.cy -= 60;
			}
		}
	}

	// m_msgViewPosition:: 1=bottom 2=right  3=left
	if (m_msgViewPosition == 1)
		ret = m_horSplitter.CreateView(0, 0, RUNTIME_CLASS(NListView), listSize, NULL);
	else if (m_msgViewPosition == 2)
		ret = m_horSplitter.CreateView(0, 0, RUNTIME_CLASS(NListView), listSize, NULL);
	else if (m_msgViewPosition == 3)
	{
		
		ret = m_horSplitter.CreateView(0, 0, RUNTIME_CLASS(NMsgView), msgSize, NULL);
		ret = m_horSplitter.CreateView(0, 1, RUNTIME_CLASS(NListView), listSize, NULL);
	}
	else
		;
	if (ret == FALSE) {
		TRACE("Failed to create top view\n");
		return -1;
	}

	// m_msgViewPosition:: 1=bottom 2=right  3=left
	if (m_msgViewPosition == 1)
		ret = m_horSplitter.CreateView(1, 0, RUNTIME_CLASS(NMsgView), msgSize, NULL);
	else if (m_msgViewPosition == 2)
		ret = m_horSplitter.CreateView(0, 1, RUNTIME_CLASS(NMsgView), msgSize, NULL);
	else if (m_msgViewPosition == 3)
		; // ret = m_horSplitter.CreateView(0, 0, RUNTIME_CLASS(NMsgView), msgSize, NULL);
	else
		;
	if (ret == FALSE) {
		TRACE("Failed to create bottom view\n");
		return -1;
	}

	return 0;
}

void CChildView::GetTreePosition(int &row, int &col)
{
	row = 0;
	col = 0;
}

void CChildView::GetListPosition(int &row, int &col)
{
	// m_msgViewPosition:: 1=bottom 2=right  3=left
	row = 0; col = 0;
	if (m_msgViewPosition == 1)
	{
		row = 0;  col = 0;
	}
	else if (m_msgViewPosition == 2)
	{
		row = 0;  col = 0;
	}
	else if (m_msgViewPosition == 3)
	{
		row = 0;  col = 1;
	}
}
void CChildView::GetMsgPosition(int &row, int &col)
{
	// m_msgViewPosition:: 1=bottom 2=right  3=left
	row = 0; col = 0;
	if (m_msgViewPosition == 1)
	{
		row = 1;  col = 0;
	}
	else if (m_msgViewPosition == 2)
	{
		row = 0;  col = 1;
	}
	else if (m_msgViewPosition == 3)
	{
		row = 0;  col = 0;
	}
}

void CChildView::OnSize(UINT nType, int cx, int cy) 
{
	CWnd::OnSize(nType, cx, cy);
	
	m_verSplitter.MoveWindow(0, 0, cx, cy);	
}


BOOL CChildView::OnCmdMsg(UINT nID, int nCode, void* pExtra, AFX_CMDHANDLERINFO* pHandlerInfo) 
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if( pFrame ) {
		CWnd *pTreeView = pFrame->GetTreeView();
		if( pTreeView && pTreeView->OnCmdMsg(nID, nCode, pExtra, pHandlerInfo) )
			return TRUE;
		CWnd *pListView = pFrame->GetListView();
		if( pListView && pListView->OnCmdMsg(nID, nCode, pExtra, pHandlerInfo) )
			return TRUE;
		CWnd *pMsgView = pFrame->GetMsgView();
		if( pMsgView && pMsgView->OnCmdMsg(nID, nCode, pExtra, pHandlerInfo) )
			return TRUE;
	}
	
	return CWnd ::OnCmdMsg(nID, nCode, pExtra, pHandlerInfo);
}
