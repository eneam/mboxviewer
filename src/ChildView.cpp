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

	CRect frameRect;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame) {
		pFrame->GetClientRect(frameRect);
		
		int deb = 1;
	}
	CSize frameSize = frameRect.Size();
	
	CRect rect;
	GetClientRect(rect);
	CSize size = rect.Size();
	size.cx = 170;
	size.cy = 200;

	if (!m_verSplitter.CreateStatic(this, 1, 2, WS_CHILD | WS_VISIBLE, AFX_IDW_PANE_FIRST)) {
		TRACE("Failed to create first splitter\n");
		return -1;
	}

	BOOL ret;
	if (m_msgViewPosition == 1)
		ret = m_horSplitter.CreateStatic(&m_verSplitter, 2, 1, WS_CHILD | WS_VISIBLE, m_verSplitter.IdFromRowCol(0, 1));
	else 
		ret = m_horSplitter.CreateStatic(&m_verSplitter, 1, 2, WS_CHILD | WS_VISIBLE, m_verSplitter.IdFromRowCol(0, 1));

	if (ret == FALSE) {
		TRACE("Failed to create second splitter\n");
		return -1;
	}

	CSize listSize = size;
	CSize msgSize = size;
	msgSize.cx = 700;
	msgSize.cy = 200;
	if (m_msgViewPosition == 2)   // windows on right
	{
		listSize.cx = frameSize.cx - size.cx - msgSize.cx - 30;
		listSize.cy = 200;
	}
	else if (m_msgViewPosition == 3)   // windows on left
	{
		listSize.cx = frameSize.cx - size.cx - msgSize.cx;
		listSize.cy = 200;
	}
	

	if (m_msgViewPosition == 1)
		ret = m_horSplitter.CreateView(0, 0, RUNTIME_CLASS(NListView), size, NULL);
	else if (m_msgViewPosition == 2)
		ret = m_horSplitter.CreateView(0, 0, RUNTIME_CLASS(NListView), listSize, NULL);
	else if (m_msgViewPosition == 3)
		ret = m_horSplitter.CreateView(0, 1, RUNTIME_CLASS(NListView), listSize, NULL);
	else
		;
	if (ret == FALSE) {
		TRACE("Failed to create top view\n");
		return -1;
	}


	if (m_msgViewPosition == 1)
		ret = m_horSplitter.CreateView(1, 0, RUNTIME_CLASS(NMsgView), size, NULL);
	else if (m_msgViewPosition == 2)
		ret = m_horSplitter.CreateView(0, 1, RUNTIME_CLASS(NMsgView), msgSize, NULL);
	else if (m_msgViewPosition == 3)
		ret = m_horSplitter.CreateView(0, 0, RUNTIME_CLASS(NMsgView), msgSize, NULL);
	else
		;
	if (ret == FALSE) {
		TRACE("Failed to create bottom view\n");
		return -1;
	}

	if( ! m_verSplitter.CreateView(0, 0, RUNTIME_CLASS(NTreeView), size, NULL) ) {
		TRACE("Failed to create left view\n");
		return -1;
	}
	
	return 0;
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
