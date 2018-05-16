// MainFrm.cpp : implementation of the CMainFrame class
//

#include "stdafx.h"
#include "mboxview.h"

#include "Resource.h"       // main symbols
#include "MainFrm.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CMainFrame

IMPLEMENT_DYNAMIC(CMainFrame, CFrameWnd)


BEGIN_MESSAGE_MAP(CMainFrame, CFrameWnd)
	//{{AFX_MSG_MAP(CMainFrame)
	ON_WM_CREATE()
	ON_WM_SETFOCUS()
	ON_UPDATE_COMMAND_UI(ID_FILE_OPEN, OnUpdateFileOpen)
	ON_UPDATE_COMMAND_UI(ID_FILE_OPTIONS, OnUpdateFileOpen)
	ON_COMMAND(ID_FILE_OPEN, OnFileOpen)
	ON_COMMAND(ID_FILE_OPTIONS, OnFileOptions)
	ON_WM_ERASEBKGND()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

static UINT indicators[] =
{
	ID_SEPARATOR,           // status line indicator
	ID_INDICATOR_CAPS,
	ID_INDICATOR_NUM,
	ID_INDICATOR_SCRL,
};

/////////////////////////////////////////////////////////////////////////////
// CMainFrame construction/destruction

CMainFrame::CMainFrame()
{
	// TODO: add member initialization code here
	m_bSelectMailFileDone = FALSE;
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

CMainFrame::~CMainFrame()
{
	// To stop memory leaks reports by debugger
	void delete_charset2Id();
	void delete_id2charset();
	delete_charset2Id();
	delete_id2charset();
	int deb = 1;
}

int CMainFrame::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CFrameWnd::OnCreate(lpCreateStruct) == -1)
		return -1;
	// create a view to occupy the client area of the frame
	if (!m_wndView.Create(NULL, NULL, AFX_WS_DEFAULT_VIEW,
		CRect(0, 0, 0, 0), this, AFX_IDW_PANE_FIRST, NULL))
	{
		TRACE0("Failed to create view window\n");
		return -1;
	}
	
	if (!m_wndToolBar.CreateEx(this, TBSTYLE_FLAT, WS_CHILD | WS_VISIBLE | CBRS_TOP
		| CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_SIZE_DYNAMIC) ||
		!m_wndToolBar.LoadToolBar(IDR_MAINFRAME))
	{
		TRACE0("Failed to create toolbar\n");
		return -1;      // fail to create
	}

	if (!m_wndStatusBar.Create(this) ||
		!m_wndStatusBar.SetIndicators(indicators,
		  sizeof(indicators)/sizeof(UINT)))
	{
		TRACE0("Failed to create status bar\n");
		return -1;      // fail to create
	}

	// TODO: Delete these three lines if you don't want the toolbar to
	//  be dockable
	m_wndToolBar.EnableDocking(CBRS_ALIGN_ANY);
	EnableDocking(CBRS_ALIGN_ANY);
	DockControlBar(&m_wndToolBar);

	SetIcon(m_hIcon, TRUE);			// Großes Symbol verwenden // use big icon
	SetIcon(m_hIcon, FALSE);		// Kleines Symbol verwenden // Use a small icon

	return 0;
}

BOOL CMainFrame::PreCreateWindow(CREATESTRUCT& cs)
{
	if( !CFrameWnd::PreCreateWindow(cs) )
		return FALSE;
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

//	cs.style = WS_OVERLAPPED | WS_CAPTION | FWS_ADDTOTITLE
//		| WS_SYSMENU | WS_MINIMIZEBOX | WS_MAXIMIZEBOX;

	cs.dwExStyle &= ~(WS_EX_MDICHILD|WS_EX_CLIENTEDGE);
	cs.lpszClass = AfxRegisterWndClass(0);
	return TRUE;
}

/////////////////////////////////////////////////////////////////////////////
// CMainFrame diagnostics

#ifdef _DEBUG
void CMainFrame::AssertValid() const
{
	CFrameWnd::AssertValid();
}

void CMainFrame::Dump(CDumpContext& dc) const
{
	CFrameWnd::Dump(dc);
}

#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CMainFrame message handlers
void CMainFrame::OnSetFocus(CWnd* pOldWnd)
{
	// forward focus to the view window
	m_wndView.SetFocus();
}

BOOL CMainFrame::OnCmdMsg(UINT nID, int nCode, void* pExtra, AFX_CMDHANDLERINFO* pHandlerInfo)
{
	// let the view have first crack at the command
	if (m_wndView.OnCmdMsg(nID, nCode, pExtra, pHandlerInfo)) {
		// Hope all objects should be initialized by now
		if (m_bSelectMailFileDone == FALSE) {
			m_bSelectMailFileDone = TRUE;
			NTreeView *treeView = GetTreeView();
			treeView->SelectMailFile();
		}
		return TRUE;
	}

	// otherwise, do default handling
	return CFrameWnd::OnCmdMsg(nID, nCode, pExtra, pHandlerInfo);
}

void CMainFrame::OnUpdateFileOpen(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable();
}

#include "OptionsDlg.h"

CString GetDateFormat(int i);

void CMainFrame::OnFileOptions()
{
	COptionsDlg d;
	if (d.DoModal() == IDOK) {
		CString format = GetDateFormat(d.m_format);
		if (GetListView()->m_format.Compare(format) != 0) {
			GetListView()->m_format = GetDateFormat(d.m_format);
			GetListView()->Invalidate();
			GetMsgView()->m_browser.m_ie.Invalidate(); // .RedrawWindow();
		}

		if (GetListView()->m_maxSearchDuration != d.m_barDelay) {
			GetListView()->m_maxSearchDuration = d.m_barDelay;
			DWORD barDelay = d.m_barDelay;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("progressBarDelay"), barDelay);
		}

		BOOL exportEml = (d.m_exportEML > 0) ? TRUE : FALSE;
		if (GetListView()->m_bExportEml != exportEml) {
			GetListView()->m_bExportEml = exportEml;
			CString str_exportEML = _T("n");
			if (exportEml == TRUE)
				str_exportEML = _T("y");
			CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("exportEML"), str_exportEML);
		}

		if (GetMsgView()->m_cnf_subj_charsetId != d.m_subj_charsetId) {
			GetMsgView()->m_cnf_subj_charsetId = d.m_subj_charsetId;
			DWORD charsetId = d.m_subj_charsetId;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("subjCharsetId"), charsetId);
		}
		if (GetMsgView()->m_cnf_from_charsetId != d.m_from_charsetId) {
			GetMsgView()->m_cnf_from_charsetId = d.m_from_charsetId;
			DWORD charsetId = d.m_from_charsetId;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("fromCharsetId"), charsetId);
		}
		if (GetMsgView()->m_cnf_to_charsetId != d.m_to_charsetId) {
			GetMsgView()->m_cnf_to_charsetId = d.m_to_charsetId;
			DWORD charsetId = d.m_to_charsetId;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("toCharsetId"), charsetId);
		}
		if (GetMsgView()->m_show_charsets != d.m_show_charsets) {
			GetMsgView()->m_show_charsets = d.m_show_charsets;
			DWORD show_charsets = d.m_show_charsets;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("showCharsets"), show_charsets);
		}
		if (GetMsgView()->m_bImageViewer != d.m_bImageViewer) {
			GetMsgView()->m_bImageViewer = d.m_bImageViewer;
			DWORD bImageViewer = d.m_bImageViewer;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("showCharsets"), bImageViewer);
		}
	}
}
void CMainFrame::OnFileOpen() 
{
	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	CBrowseForFolder bff(GetSafeHwnd(), CSIDL_DESKTOP, IDS_SELECT_FOLDER);
	if( ! path.IsEmpty() )
		bff.SetDefaultFolder(path);
	bff.SetFlags(BIF_RETURNONLYFSDIRS);
	if( bff.SelectFolder() ) {
		path = bff.GetSelectedFolder();
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath", path);
		GetTreeView()->FillCtrl();
		AfxGetApp()-> AddToRecentFileList(path);
	}
}

void CMainFrame::DoOpen(CString& path) 
{
	CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath", path);
	GetTreeView()->FillCtrl();
}

NTreeView * CMainFrame::GetTreeView()
{
	return (NTreeView *)m_wndView.m_verSplitter.GetPane(0, 0);
}

NListView * CMainFrame::GetListView()
{
	return (NListView *)m_wndView.m_horSplitter.GetPane(0, 0);
}

NMsgView * CMainFrame::GetMsgView()
{
	return (NMsgView *)m_wndView.m_horSplitter.GetPane(1, 0);
}

BOOL CMainFrame::OnEraseBkgnd(CDC* pDC) 
{
	// active view will paint itself
	return TRUE;
}