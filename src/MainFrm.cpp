// MainFrm.cpp : implementation of the CMainFrame class
//

#include "stdafx.h"
#include "mboxview.h"

#include "Resource.h"       // main symbols
#include "MainFrm.h"
#include "MboxMail.h"
#include "ExportToCSVDlg.h"
#include "OpenContainingFolderDlg.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

BOOL BrowseToFile(LPCTSTR filename)
{
	BOOL retval = TRUE;

	// It appears m_ie.Create  in CBrowser::OnCreate calls CoInitilize
	//Com_Initialize();

	ITEMIDLIST *pidl = ILCreateFromPath(filename);
	if (pidl) {
		HRESULT  ret = SHOpenFolderAndSelectItems(pidl, 0, 0, 0);
		if (ret != S_OK)
			retval = FALSE;
		ILFree(pidl);
	}
	else
		retval = FALSE;
	return retval;
}

#define MaxShellExecuteErrorCode 32
void CheckShellExecuteResult(HINSTANCE  result, HWND h)
{
	if ((UINT)result <= MaxShellExecuteErrorCode) {
		CString errorText;
		ShellExecuteError2Text((UINT)result, errorText);
		int answer = ::MessageBox(h, errorText, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
	}
}

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
	//ON_COMMAND(ID_FILE_EXPORT_TO_CSV, &CMainFrame::OnFileExportToCsv)
	ON_COMMAND(ID_VIEW_CODEPAGEIDS, &CMainFrame::OnViewCodepageids)
	ON_COMMAND(ID_PRINTTO_CSV, &CMainFrame::OnPrinttoCsv)
	ON_COMMAND(ID_PRINTTO_TEXT, &CMainFrame::OnPrinttoText)
	ON_COMMAND(ID_PRINTTO_HTML, &CMainFrame::OnPrinttoHtml)
	ON_COMMAND(ID_SORTBY_DATE, &CMainFrame::OnBydate)
	ON_UPDATE_COMMAND_UI(ID_SORTBY_DATE, &CMainFrame::OnUpdateBydate)
	ON_COMMAND(ID_SORTBY_FROM, &CMainFrame::OnByfrom)
	ON_UPDATE_COMMAND_UI(ID_SORTBY_FROM, &CMainFrame::OnUpdateByfrom)
	ON_COMMAND(ID_SORTBY_TO, &CMainFrame::OnByto)
	ON_UPDATE_COMMAND_UI(ID_SORTBY_TO, &CMainFrame::OnUpdateByto)
	ON_COMMAND(ID_SORTBY_SUBJECT, &CMainFrame::OnBysubject)
	ON_UPDATE_COMMAND_UI(ID_SORTBY_SUBJECT, &CMainFrame::OnUpdateBysubject)
	ON_COMMAND(ID_SORTBY_SIZE, &CMainFrame::OnBysize)
	ON_UPDATE_COMMAND_UI(ID_SORTBY_SIZE, &CMainFrame::OnUpdateBysize)
	ON_COMMAND(ID_SORTBY_CONVERSATION, &CMainFrame::OnByconversation)
	ON_UPDATE_COMMAND_UI(ID_SORTBY_CONVERSATION, &CMainFrame::OnUpdateByconversation)
	ON_UPDATE_COMMAND_UI(ID_INDICATOR_MAIL, &CMainFrame::OnUpdateMailDownloadStatus)
	ON_UPDATE_COMMAND_UI(ID_INDICATOR_NUM, &CMainFrame::OnUpdateMailIndex)
END_MESSAGE_MAP()

static UINT indicators[] =
{
	ID_SEPARATOR,           // status line indicator
	ID_INDICATOR_MAIL,
	ID_INDICATOR_CAPS,
	ID_INDICATOR_NUM,
	ID_INDICATOR_SCRL,
};

/////////////////////////////////////////////////////////////////////////////
// CMainFrame construction/destruction

CMainFrame::CMainFrame()
{
	// TODO: add member initialization code here
	m_bMailDownloadComplete = FALSE;
	m_bSelectMailFileDone = FALSE;
	m_MailIndex = -1;
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

CMainFrame::~CMainFrame()
{
	// To stop memory leaks reports by debugger
	MboxMail::ReleaseResources();
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

	// TODO: Create seprate Mail File download status bar ?
	//m_wndStatusBar.SetWindowText("Mail Download Complete: ");

	// TODO: Delete these three lines if you don't want the toolbar to
	//  be dockable
	m_wndToolBar.EnableDocking(CBRS_ALIGN_ANY);
	EnableDocking(CBRS_ALIGN_ANY);
	DockControlBar(&m_wndToolBar);

	SetIcon(m_hIcon, TRUE);			// use big icon
	SetIcon(m_hIcon, FALSE);		// Use a small icon

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
			if (treeView)
				treeView->SelectMailFile();
		}
		return TRUE;
	}

	// otherwise, do default handling
	return CFrameWnd::OnCmdMsg(nID, nCode, pExtra, pHandlerInfo);
}

void CMainFrame::OnUpdateFileOpen(CCmdUI* pCmdUI) 
{
	if (pCmdUI)
		pCmdUI->Enable();
}

#include "OptionsDlg.h"

CString GetDateFormat(int i);

void CMainFrame::OnFileOptions()
{
	bool needRedraw = false;
	COptionsDlg d;
	if (d.DoModal() == IDOK) {
		CString format = GetDateFormat(d.m_format);
		if (GetListView()->m_format.Compare(format) != 0) {
			GetListView()->m_format = GetDateFormat(d.m_format);
			needRedraw = true;
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
			needRedraw = true;
		}
		if (GetMsgView()->m_cnf_from_charsetId != d.m_from_charsetId) {
			GetMsgView()->m_cnf_from_charsetId = d.m_from_charsetId;
			DWORD charsetId = d.m_from_charsetId;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("fromCharsetId"), charsetId);
			needRedraw = true;
		}
		if (GetMsgView()->m_cnf_to_charsetId != d.m_to_charsetId) {
			GetMsgView()->m_cnf_to_charsetId = d.m_to_charsetId;
			DWORD charsetId = d.m_to_charsetId;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("toCharsetId"), charsetId);
			needRedraw = true;
		}
		if (GetMsgView()->m_show_charsets != d.m_show_charsets) {
			GetMsgView()->m_show_charsets = d.m_show_charsets;
			DWORD show_charsets = d.m_show_charsets;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("showCharsets"), show_charsets);
			needRedraw = true;
		}
		if (GetMsgView()->m_bImageViewer != d.m_bImageViewer) {
			GetMsgView()->m_bImageViewer = d.m_bImageViewer;
			DWORD bImageViewer = d.m_bImageViewer;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("showCharsets"), bImageViewer);
		}
		if (GetListView()->m_gmtTime != d.m_bTimeType)
		{
			GetListView()->m_gmtTime = d.m_bTimeType;
			DWORD bTimeType = d.m_bTimeType;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("timeType"), bTimeType);
			needRedraw = true;
		}
		if (needRedraw) {
			// TODO: Invalidate() below doesn't updates column labels, so call MarkColumns() directly
			// mboxview is basically single threaded so it should be ok to do so
			GetListView()->MarkColumns();  // Invalidate() below doesn't updates column labels
			GetListView()->ClearDescView();
			GetListView()->Invalidate();
			GetMsgView()->m_browser.m_ie.Invalidate(); // .RedrawWindow(); 
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

BOOL PathFileExist(LPCSTR path);

void CMainFrame::OnFileExportToCsv()
{
	// TODO: Add your command handler code here
	ExportToCSVDlg d;
	if (d.DoModal() == IDOK) 
	{
		CSVFILE_CONFIG csvConfig;
		csvConfig.m_bFrom = d.m_bFrom;
		csvConfig.m_bTo = d.m_bTo;
		csvConfig.m_bSubject = d.m_bSubject;
		csvConfig.m_bDate = d.m_bDate;
		csvConfig.m_bContent = d.m_bContent;
		csvConfig.m_dateFormat = d.m_dateFormat;
		csvConfig.m_bGMTTime = d.m_bGMTTime;
	
		//csvConfig.m_separator = d.m_separator;
		// Hardcoded for now.
		csvConfig.m_separator = ",";

		csvConfig.m_nCodePageId = 0;
		if (d.m_bEncodingType == 1)
			csvConfig.m_nCodePageId = CP_UTF8;
		else if (d.m_bEncodingType == 2) {
			csvConfig.m_nCodePageId = d.m_nCodePageId;
		}

		CString csvFileName;

		//int ret = MboxMail::exportToCSVFileFullMailParse(csvConfig);
		int ret = MboxMail::exportToCSVFile(csvConfig, csvFileName);
		if (ret > 0) {
			CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
			if (!path.IsEmpty())  // not likely since the path was valid in MboxMail::exportToCSVFile(csvConfig);
			{
				if (PathFileExist(path)) { // likely :) 
					CString text = "Created file\n\n" + csvFileName;
					OpenContainingFolderDlg dlg(text, 0);
					INT_PTR nResponse = dlg.DoModal();
					if (nResponse == IDOK)
					{
						if (BrowseToFile(csvFileName) == FALSE) {
							HWND h = GetSafeHwnd();
							HINSTANCE result = ShellExecute(h, _T("open"), path, NULL, NULL, SW_SHOWNORMAL);
							CheckShellExecuteResult(result, h);
						}
						int deb = 1;
					}
					else if (nResponse == IDYES)
					{
						HWND h = GetSafeHwnd();
						HINSTANCE result = ShellExecute(h, _T("open"), csvFileName, NULL, NULL, SW_SHOWNORMAL);
						CheckShellExecuteResult(result, h);
						int deb = 1;
					}
					else if (nResponse == IDCANCEL)
					{
						int deb = 1;
					}
				}
				else
					;
			}
			else
				; // TODO: 
		}
	}
}

void CMainFrame::OnViewCodepageids()
{
	// TODO: Add your command handler code here
	int showCodePageTable();
	int ret = showCodePageTable();
}

// Called when File->"Print To"->CSV
void CMainFrame::OnPrinttoCsv()
{
	// TODO: Add your command handler code here
	if (MboxMail::s_mails.GetSize() == 0) {
		CString txt = _T("Please open mail file first.");
		HWND h = GetSafeHwnd(); // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;
	}
	OnFileExportToCsv();
}

// Called when File->"Print To"->Text
void CMainFrame::OnPrinttoText()
{
	OnPrinttoTextFile(0);
}

void CMainFrame::OnPrinttoTextFile(int textType)
{
	// TODO: Add your command handler code here
#if 0
	//CPrintDialog prtDlg(FALSE);
	CPageSetupDialog prtDlg(FALSE);
	prtDlg.DoModal();
#else
	if (MboxMail::s_mails.GetSize() == 0) {
		CString txt = _T("Please open mail file first.");
		HWND h = GetSafeHwnd(); // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;
	}

	//ExportToCSVDlg d;
	//if (d.DoModal() == IDOK)
	{
		TEXTFILE_CONFIG textConfig;
#if 1
		NListView * pListView = CMainFrame::GetListView();
		if (pListView) {
			textConfig.m_dateFormat = pListView->m_format;
			textConfig.m_bGMTTime = pListView->m_gmtTime;
		}
		else
		{
			textConfig.m_dateFormat = GetDateFormat(0);
			textConfig.m_bGMTTime = 0;
		}
		textConfig.m_nCodePageId = CP_UTF8;
#else
		textConfig.m_dateFormat = GetDateFormat(d.m_dateFormat);
		textConfig.m_bGMTTime = d.m_bGMTTime;

		textConfig.m_nCodePageId = 0;
		if (d.m_bEncodingType == 1)
			textConfig.m_nCodePageId = CP_UTF8;
		else if (d.m_bEncodingType == 2) {
			textConfig.m_nCodePageId = d.m_nCodePageId;
		}
#endif

		CString textFileName;
		int firstMail = 0;
		int lastMail = MboxMail::s_mails.GetCount() - 1;
		int ret = MboxMail::exportToTextFile(textConfig, textFileName, firstMail, lastMail, textType);
		if (ret > 0) {
			CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
			if (!path.IsEmpty())  // not likely since the path was valid in MboxMail::exportToCSVFile(csvConfig);
			{
				if (PathFileExist(path)) { // likely :) 
					CString txt = "Created file\n\n" + textFileName;
					OpenContainingFolderDlg dlg(txt, 0);
					INT_PTR nResponse = dlg.DoModal();
					if (nResponse == IDOK)
					{
						if (BrowseToFile(textFileName) == FALSE) {
							HWND h = GetSafeHwnd();
							HINSTANCE result = ShellExecute(h, _T("open"), path, NULL, NULL, SW_SHOWNORMAL);
							CheckShellExecuteResult(result, h);
						}
						int deb = 1;
					}
					else if (nResponse == IDYES)
					{
						HWND h = GetSafeHwnd();
						HINSTANCE result = ShellExecute(h, _T("open"), textFileName, NULL, NULL, SW_SHOWNORMAL);
						CheckShellExecuteResult(result, h);
						int deb = 1;
					}
					else if (nResponse == IDCANCEL)
					{
						int deb = 1;
					}
				}
				else
					;
			}
			else
				; // TODO: 
		}
	}
#endif
}

// called by NListView::OnRClick
void CMainFrame::OnPrintSingleMailtoText(int mailPosition, int textType, BOOL forceOpen)  // textType 0==Plain, 1==Html
{
	// TODO: Add your command handler code here
#if 0
	//CPrintDialog prtDlg(FALSE);
	CPageSetupDialog prtDlg(FALSE);
	prtDlg.DoModal();
#else
	if (MboxMail::s_mails.GetSize() == 0) {
		CString txt = _T("Please open mail file first.");
		HWND h = GetSafeHwnd(); // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;
	}

	// Do we need extra dialog? Bypass for now
	//ExportToCSVDlg d;
	//if (d.DoModal() == IDOK)
	{
		TEXTFILE_CONFIG textConfig;
#if 1
		NListView * pListView = CMainFrame::GetListView();
		if (pListView) {
			textConfig.m_dateFormat = pListView->m_format;
			textConfig.m_bGMTTime = pListView->m_gmtTime;
		}
		else
		{
			textConfig.m_dateFormat = GetDateFormat(0);
			textConfig.m_bGMTTime = 0;
		}
		textConfig.m_nCodePageId = CP_UTF8;
#else
		textConfig.m_dateFormat = d.m_dateFormat;
		textConfig.m_bGMTTime = d.m_bGMTTime;

		textConfig.m_nCodePageId = 0;
		if (d.m_bEncodingType == 1)
			textConfig.m_nCodePageId = CP_UTF8;
		else if (d.m_bEncodingType == 2) {
			textConfig.m_nCodePageId = d.m_nCodePageId;
		}
#endif

		int ret = 0;
		CString textFileName;
		ret = MboxMail::exportToTextFile(textConfig, textFileName, mailPosition, mailPosition, textType);
		if (ret > 0) {
			CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
			if (!path.IsEmpty())  // not likely since the path was valid in MboxMail::exportToCSVFile(csvConfig);
			{
				if (PathFileExist(path)) { // likely :) 
					CString txt = "Created file\n\n" + textFileName;
					if (forceOpen == FALSE)
					{
						OpenContainingFolderDlg dlg(txt, 0);
						INT_PTR nResponse = dlg.DoModal();
						if (nResponse == IDOK)
						{
							if (BrowseToFile(textFileName) == FALSE) {
								HWND h = GetSafeHwnd();
								HINSTANCE result = ShellExecute(h, _T("open"), path, NULL, NULL, SW_SHOWNORMAL);
								CheckShellExecuteResult(result, h);
							}
							int deb = 1;
						}
						else if (nResponse == IDYES)
						{
							HWND h = GetSafeHwnd();
							HINSTANCE result = ShellExecute(h, _T("open"), textFileName, NULL, NULL, SW_SHOWNORMAL);
							CheckShellExecuteResult(result, h);
							int deb = 1;
						}
						else if (nResponse == IDCANCEL)
						{
							int deb = 1;
						}
					}
					else
					{
						HWND h = GetSafeHwnd();
						HINSTANCE result = ShellExecute(h, _T("open"), textFileName, NULL, NULL, SW_SHOWNORMAL);
						CheckShellExecuteResult(result, h);
						int deb = 1;
					}
				}
				else
					;
			}
			else
				; // TODO: 
		}
	}
#endif
}

void CMainFrame::OnPrinttoHtml()
{
	// TODO: Add your command handler code here
	OnPrinttoTextFile(1);
}

// Called when View->"Sort by"->Date
void CMainFrame::OnBydate()
{
	// TODO: Add your command handler code here
	NListView *pListView = 0;
	pListView = GetListView();
	if (pListView)
		pListView->SortByColumn(1);
}


void CMainFrame::OnUpdateBydate(CCmdUI *pCmdUI)
{
	// TODO: Add your command update UI handler code here
	pCmdUI->Enable(MboxMail::s_mails.GetSize() > 0);
}


void CMainFrame::OnByfrom()
{
	// TODO: Add your command handler code here
	NListView *pListView = 0;
	pListView = GetListView();
	if (pListView)
		pListView->SortByColumn(2);
}


void CMainFrame::OnUpdateByfrom(CCmdUI *pCmdUI)
{
	// TODO: Add your command update UI handler code here
	pCmdUI->Enable(MboxMail::s_mails.GetSize() > 0);
}


void CMainFrame::OnByto()
{
	// TODO: Add your command handler code here
	NListView *pListView = 0;
	pListView = GetListView();
	if (pListView)
		pListView->SortByColumn(3);
}


void CMainFrame::OnUpdateByto(CCmdUI *pCmdUI)
{
	// TODO: Add your command update UI handler code here
	pCmdUI->Enable(MboxMail::s_mails.GetSize() > 0);
}


void CMainFrame::OnBysubject()
{
	// TODO: Add your command handler code here
	NListView *pListView = 0;
	pListView = GetListView();
	if (pListView)
		pListView->SortByColumn(4);
}


void CMainFrame::OnUpdateBysubject(CCmdUI *pCmdUI)
{
	// TODO: Add your command update UI handler code here
	pCmdUI->Enable(MboxMail::s_mails.GetSize() > 0);
}


void CMainFrame::OnBysize()
{
	// TODO: Add your command handler code here
	NListView *pListView = 0;
	pListView = GetListView();
	if (pListView)
		pListView->SortByColumn(5);
}


void CMainFrame::OnUpdateBysize(CCmdUI *pCmdUI)
{
	// TODO: Add your command update UI handler code here
	pCmdUI->Enable(MboxMail::s_mails.GetSize() > 0);
}


void CMainFrame::OnByconversation()
{
	// TODO: Add your command handler code here
	NListView *pListView = 0;
	pListView = GetListView();
	if (pListView)
		pListView->SortByColumn(0);
}


void CMainFrame::OnUpdateByconversation(CCmdUI *pCmdUI)
{
	// TODO: Add your command update UI handler code here
	pCmdUI->Enable(MboxMail::s_mails.GetSize() > 0);
}

void CMainFrame::OnUpdateMailDownloadStatus(CCmdUI *pCmdUI)
{
	// TODO: Add your command update UI handler code here

	NMsgView * msgView = GetMsgView();
	NListView * listView = GetListView();
	if ((listView == 0) || (msgView == 0))
	{
		pCmdUI->Enable(0);
		return;
	}

	if (listView->m_lastSel < 0)
	{
		pCmdUI->Enable(0);
		return;
	}

	pCmdUI->Enable();
	CString strPage;
	if (msgView->m_browser.m_bNavigateComplete)
		m_bMailDownloadComplete = TRUE;
	else
		m_bMailDownloadComplete = FALSE;

	if (m_bMailDownloadComplete)
		strPage.Format("%s", _T("Mail Retrieval Complete"));
	else
		strPage.Format("%s", _T("Mail Retrieval In Progress ..."));

	pCmdUI->SetText(strPage);
}

void CMainFrame::OnUpdateMailIndex(CCmdUI *pCmdUI)
{
	// TODO: Add your command update UI handler code here

	NMsgView * msgView = GetMsgView();
	NListView * listView = GetListView();
	if ((listView == 0) || (msgView == 0))
	{
		pCmdUI->Enable(0);
		return;
	}

	if (listView->m_lastSel < 0)
	{
		pCmdUI->Enable(0);
		return;
	}

	pCmdUI->Enable();
	CString strMailIndex;
	if (listView->m_lastSel >= 0)
		strMailIndex.Format("%d", listView->m_lastSel);

	pCmdUI->SetText(strMailIndex);
}
