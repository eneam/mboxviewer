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
	ON_UPDATE_COMMAND_UI(ID_INDICATOR_MAIL_INDEX, &CMainFrame::OnUpdateMailIndex)
	ON_BN_CLICKED(IDC_ARCHIVE_LIST, &CMainFrame::OnBnClickedArchiveList)
	ON_BN_CLICKED(IDC_FIND_LIST, &CMainFrame::OnBnClickedFindList)
	ON_BN_CLICKED(IDC_EDIT_LIST, &CMainFrame::OnBnClickedEditList)
	ON_NOTIFY(NM_CUSTOMDRAW, IDC_ARCHIVE_LIST, &CMainFrame::OnNMCustomdrawArchiveList)
	ON_NOTIFY(NM_CUSTOMDRAW, IDC_FIND_LIST, &CMainFrame::OnNMCustomdrawFindList)
	ON_NOTIFY(NM_CUSTOMDRAW, IDC_EDIT_LIST, &CMainFrame::OnNMCustomdrawEditList)
	ON_BN_CLICKED(IDC_BUTTON2, &CMainFrame::OnBnClickedButton2)
	ON_NOTIFY(NM_CUSTOMDRAW, IDC_BUTTON2, &CMainFrame::OnNMCustomdrawButton2)
	ON_COMMAND(ID_VIEW_USERSELECTEDMAILS, &CMainFrame::OnViewUserselectedmails)
	ON_UPDATE_COMMAND_UI(ID_VIEW_USERSELECTEDMAILS, &CMainFrame::OnUpdateViewUserselectedmails)
	ON_COMMAND(ID_HELP_MBOXVIEWHELP, &CMainFrame::OnHelpMboxviewhelp)
END_MESSAGE_MAP()

static UINT indicators[] =
{
	ID_SEPARATOR,           // status line indicator
	ID_INDICATOR_MAIL,
	ID_INDICATOR_MAIL_INDEX,
	ID_INDICATOR_CAPS,
	ID_INDICATOR_NUM,
	ID_INDICATOR_SCRL,
};

/////////////////////////////////////////////////////////////////////////////
// CMainFrame construction/destruction

CMainFrame::CMainFrame()
{
	// TODO: add member initialization code here
	m_bMailDownloadComplete = FALSE;  // Page download complete in CBrowser
	m_bSelectMailFileDone = FALSE;
	m_MailIndex = -1;  // Not used ??
	m_bUserSelectedMailsCheckSet = FALSE;  // User Selected List checked/unched state
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

#if 0
void CMainFrame::DoDataExchange(CDataExchange* pDX)
{
	//m_wndDlgBar.DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CMainFrame);
	DDX_Check(pDX, IDC_CASE, m_bMailListType);
	//}}AFX_DATA_MAP
}
#endif

CMainFrame::~CMainFrame()
{
	// To stop memory leaks reports by debugger
	MboxMail::ReleaseResources();
	int deb = 1;
}

int CMainFrame::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	// Just a test of not related functionality :)
	//BOOL ret = MboxMail::Test_MergeTwoMailLists();

	if (CFrameWnd::OnCreate(lpCreateStruct) == -1)
		return -1;
	// create a view to occupy the client area of the frame
	if (!m_wndView.Create(NULL, NULL, AFX_WS_DEFAULT_VIEW,
		CRect(0, 0, 0, 0), this, AFX_IDW_PANE_FIRST, NULL))
	{
		TRACE0("Failed to create view window\n");
		return -1;
	}

#if 0
	if (!m_wndToolBar.CreateEx(this, TBSTYLE_FLAT, WS_CHILD | WS_VISIBLE | CBRS_TOP
		| CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_SIZE_DYNAMIC) ||
		!m_wndToolBar.LoadToolBar(IDR_MAINFRAME))
	{
		TRACE0("Failed to create toolbar\n");
		return -1;      // fail to create
	}
#else
	if (!m_wndToolBar.CreateEx(this) ||
		!m_wndToolBar.LoadToolBar(IDR_MAINFRAME))
	{
		TRACE0("Failed to create toolbar\n");
		return -1;      // fail to create
	}

	// TODO: Remove this if you don't want tool tips
	m_wndToolBar.SetBarStyle(m_wndToolBar.GetBarStyle() | CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_TOP | CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_SIZE_DYNAMIC);
#endif

	CToolBarCtrl &wndToolBarCtrl = m_wndToolBar.GetToolBarCtrl();

	if (!m_wndStatusBar.Create(this) ||
		!m_wndStatusBar.SetIndicators(indicators,
		  sizeof(indicators)/sizeof(UINT)))
	{
		TRACE0("Failed to create status bar\n");
		return -1;      // fail to create
	}

	if (!m_wndDlgBar.Create(this, IDD_EDITOR,
		CBRS_ALIGN_TOP | CBRS_ALIGN_LEFT, AFX_IDW_DIALOGBAR))
	{
		TRACE0("Failed to create dialogbar\n");
		return -1;		// fail to create
	}

	DWORD dwCtrlStyle = RBS_BANDBORDERS;
	DWORD dwStyle = WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | CBRS_TOP;

	if (!m_wndReBar.Create(this, dwCtrlStyle, dwStyle) ||
		!m_wndReBar.AddBar(&m_wndToolBar) ||
		!m_wndReBar.AddBar(&m_wndDlgBar))
	{
		TRACE0("Failed to create rebar\n");
		return -1;      // fail to create
	}

	CReBarCtrl &rebarCtl = m_wndReBar.GetReBarCtrl();
	int iIndex = 0;
	rebarCtl.MinimizeBand(iIndex);

	// TODO: Create seprate Mail File download status bar ?
	//m_wndStatusBar.SetWindowText("Mail Download Complete: ");

	// TODO: Delete these three lines if you don't want the toolbar to be dockable
	//m_wndToolBar.EnableDocking(CBRS_ALIGN_ANY);
	//EnableDocking(CBRS_ALIGN_ANY);
	//DockControlBar(&m_wndToolBar);

	EnableAllMailLists(FALSE);

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
			NTreeView *pTreeView = GetTreeView();
			if (pTreeView)
				pTreeView->SelectMailFile();
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
			GetMsgView()->Invalidate(); 
			//GetMsgView()->m_browser.m_ie.Invalidate(); // TODO: changed to GetMsgView()->Invalidate();
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
		csvConfig.m_MessageLimitString = d.m_MessageLimitString;
	
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
		int firstMail = 0;
		int lastMail = MboxMail::s_mails.GetCount() - 1;
		BOOL progressBar = TRUE;

		int ret = MboxMail::exportToCSVFile(csvConfig, csvFileName, firstMail, lastMail, progressBar);
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
// File->"Print To"->CVS -> OnPrinttoCsv() -> OnFileExportToCsv() -> exportToCSVFile() for firstMail;lastMail -> NO separate printSingleMailToTextFile(index) yet
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
// File->"Print To"->Text -> OnPrinttoText() -> OnPrinttoTextFile() -> exportToTextFile() for firstMail;lastMail -> printSingleMailToTextFile(index)
void CMainFrame::OnPrinttoText()
{
	OnPrinttoTextFile(0);
}

void CMainFrame::OnPrinttoTextFile(int textType)
{
	// TODO: Add your command handler code here
	if (MboxMail::s_mails.GetSize() == 0) {
		CString txt = _T("Please open mail file first.");
		HWND h = GetSafeHwnd(); // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;
	}

	// For now Kept in case we will add extra dialog similar to print to CSV
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
#if 1
		CString textFileName;
		int firstMail = 0;
		int lastMail = MboxMail::s_mails.GetCount() - 1;
		BOOL progressBar = TRUE;
		MailIndexList *selectedMailsIndexList = 0;
		int ret = MboxMail::exportToTextFile(textConfig, textFileName, firstMail, lastMail, selectedMailsIndexList, textType, progressBar);

		if (ret > 0) {
			CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
			if (!path.IsEmpty())  // not likely since the path was valid in MboxMail::exportToTextFile(...);
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
void CMainFrame::OnPrintSingleMailtoText(int mailPosition, int textType, BOOL forceOpen, BOOL printToPrinter)  // textType 0==Plain, 1==Html
{
	// TODO: Add your command handler code here
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
#if 1

		int ret = 0;
		BOOL progressBar = FALSE;
		CString textFileName;
		MailIndexList *selectedMailsIndexList = 0;
		ret = MboxMail::exportToTextFile(textConfig, textFileName, mailPosition, mailPosition, selectedMailsIndexList, textType, progressBar);
		if (ret > 0) {
			CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
			if (!path.IsEmpty())  // not likely since the path was valid in MboxMail::exportToTextFile(...);
			{
				if (PathFileExist(path)) { // likely :) 
					CString txt = "Created file\n\n" + textFileName;
					if (printToPrinter)
					{
						CFile fp;
						if (fp.Open(textFileName, CFile::modeRead | CFile::shareDenyWrite)) {
							ULONGLONG ll = fp.GetLength();
							SimpleString *inbuf = MboxMail::m_inbuf;
							SimpleString *workbuf = MboxMail::m_workbuf;
							inbuf->ClearAndResize((int)ll);

							UINT l = fp.Read(inbuf->Data(), (UINT)ll);
							inbuf->SetCount(l);

							UINT inCodePage = CP_UTF8;
							NMsgView::PrintHTMLDocumentToPrinter(inbuf, workbuf, inCodePage);

							int deb = 1;
						}
						else {
							// MessageBox ??
							int deb = 1;
						}
					}
					else if (forceOpen == FALSE)
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


// File->"Print To"->Text->OnPrinttoHtml()->OnPrinttoTextFile()->exportToTextFile() for firstMail; lastMail->printSingleMailToTextFile(index)
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


// Show on the status bar mailNumber of totalMailCnt
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

	//if (listView->m_lastSel < 0)
	int mailCnt = MboxMail::s_mails.GetCount();
	if (mailCnt <= 0)
	{
		pCmdUI->Enable(0);
		return;
	}

	pCmdUI->Enable();
	CString strMailIndex;
	if (listView->m_lastSel >= 0)
		strMailIndex.Format("Mail %d of %d", listView->m_lastSel + 1, mailCnt);
	else
		strMailIndex.Format("Mail xx of %d", mailCnt);

	pCmdUI->SetText(strMailIndex);

}

void CMainFrame::SetupMailListsToInitialState()
{
	// called when index archive is loaded into MboxMail::s_mails
	MboxMail::s_mails_all.SetSizeKeepData(0);
	MboxMail::s_mails_find.SetSizeKeepData(0);
	MboxMail::s_mails_edit.SetSizeKeepData(0);


	if (MboxMail::b_mails_which_sorted != 1)
		int deb = 1; // TODO: ASSERT ?


	MboxMail::m_allMails.m_lastSel = -1;
	MboxMail::m_allMails.b_mails_which_sorted = 1; // by date
	MboxMail::m_findMails.m_lastSel = -1;
	MboxMail::m_findMails.b_mails_which_sorted = 1; // not really, shpuld be unknown yet
	MboxMail::m_editMails.m_lastSel = -1;
	MboxMail::m_editMails.b_mails_which_sorted = 1; // not really, shpuld be unknown yet

	CMenu *menu = this->GetMenu();
	menu->CheckMenuItem(ID_VIEW_USERSELECTEDMAILS, MF_UNCHECKED);
	m_bUserSelectedMailsCheckSet = FALSE;

	// User Selected Mails will not be enabled :) fix it/make it clear ?
	EnableAllMailLists(TRUE);

	int nID = IDC_ARCHIVE_LIST;
	SetMailList(nID);
}

void CMainFrame::SetMailList(int nID)
{
	// clear check marks from all buttons and set nID button
	int nIDFirstButton = IDC_ARCHIVE_LIST;
	int nIDLastButton = IDC_EDIT_LIST;
	m_wndDlgBar.CheckRadioButton(nIDFirstButton, nIDLastButton, nID);
}

void CMainFrame::EnableMailList(int nId, BOOL enable)
{
	if (m_wndDlgBar.GetSafeHwnd())
	{
		CWnd *p = m_wndDlgBar.GetDlgItem(nId);
		if (p) {
			int nFlags = ((CButton*)p)->GetCheck();
			((CButton*)p)->SetCheck(0);
			p->EnableWindow(enable);
		}
	}
}

void CMainFrame::EnableAllMailLists(BOOL enable)  // enable/disable
{
	//int idList[3] = { IDC_ARCHIVE_LIST , IDC_FIND_LIST , IDC_EDIT_LIST };
	if (m_wndDlgBar.GetSafeHwnd()) 
	{
		CWnd *p = m_wndDlgBar.GetDlgItem(IDC_ARCHIVE_LIST);
		if (p) {
			int nFlags = ((CButton*)p)->GetCheck();
			((CButton*)p)->SetCheck(0);
			p->EnableWindow(enable);
		}
		p = m_wndDlgBar.GetDlgItem(IDC_FIND_LIST);
		if (p) {
			((CButton*)p)->SetCheck(0);
			p->EnableWindow(enable);
		}

		p = m_wndDlgBar.GetDlgItem(IDC_EDIT_LIST);  // disable only; enable in other places
		if (p) {
			((CButton*)p)->SetCheck(0);
			if (enable == FALSE)
				p->EnableWindow(enable);
		}
	}
}


void CMainFrame::OnBnClickedArchiveList()
{
	// TODO: Add your control notification handler code here
	int nID = IDC_ARCHIVE_LIST;
	if (MboxMail::nWhichMailList == nID)
		return; // do nothing

	NListView * pListView = GetListView();
	if (pListView) {
		pListView->SwitchToMailList(nID);
	}
	int deb = 1;
}

void CMainFrame::OnBnClickedFindList()
{
	// TODO: Add your control notification handler code here
	int nID = IDC_FIND_LIST;
	if (MboxMail::nWhichMailList == nID)
		return; // do nothing

	NListView * pListView = GetListView();
	if (pListView) {
		pListView->SwitchToMailList(nID);
	}
	int deb = 1;
}

void CMainFrame::OnBnClickedEditList()
{
	// TODO: Add your control notification handler code here
	int nID = IDC_EDIT_LIST;
	if (MboxMail::nWhichMailList == nID)
		return; // do nothing

	NListView * pListView = GetListView();
	if (pListView) {
		pListView->SwitchToMailList(nID);
	}
	int deb = 1;
}

void CMainFrame::OnNMCustomdrawArchiveList(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMCUSTOMDRAW pNMCD = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;
	NMCustomdrawEditList(pNMHDR, pResult);
}


void CMainFrame::OnNMCustomdrawFindList(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMCUSTOMDRAW pNMCD = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;
	NMCustomdrawEditList(pNMHDR, pResult);
}


void CMainFrame::OnNMCustomdrawEditList(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMCUSTOMDRAW pNMCD = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;
	NMCustomdrawEditList(pNMHDR, pResult);
}

void CMainFrame::NMCustomdrawEditList(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMCUSTOMDRAW pNMCD = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;

	CWnd *p = m_wndDlgBar.GetDlgItem(pNMCD->dwItemSpec);
	if (p) {
		CButton *button = (CButton*)p;
		int nFlags = button->GetCheck();

		if (nFlags == 0)
			return;
		int deb = 1;
	}

	switch (pNMCD->dwDrawStage)
	{
	case CDDS_PREERASE:
	{
		CDC dc;
		CRect rect = pNMCD->rc;
		HDC hDC = pNMCD->hdc;

#define LightGreen 	RGB(193,255,193) // green

		if (dc.Attach(hDC))
		{
			dc.FillRect(&rect, &CBrush(LightGreen));
			dc.SetBkMode(TRANSPARENT);
			dc.SetTextColor(RGB(0, 0, 0));

			dc.Detach();
		}
		int deb = 1;
	}
	break;
	default:
	{
		*pResult = CDDS_PREERASE;
		int deb = 1;  // should not be here
	}
	}
	int deb = 1;
}


void CMainFrame::OnBnClickedButton2()  // help button on tool bar
{
	// TODO: Add your control notification handler code here

	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	if (!path.IsEmpty())
	{
		if (!PathFileExist(path)) {
			return;
		}
	}
	else
		return;

	CString codePageIdsFile = "MailListsInfo.htm";
	CString fullPath = path + "\\" + codePageIdsFile;

	CFile fp;
	if (!fp.Open(fullPath, CFile::modeWrite | CFile::modeCreate)) {
		CString txt = _T("Could not create \"") + fullPath;
		txt += _T("\" file.\nMake sure file is not open on other applications.");
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return;
	}

	CreateMailListsInfoText(fp);

	fp.Close();

	ShellExecute(NULL, _T("open"), fullPath, NULL, NULL, SW_SHOWNORMAL);

	int deb = 1;
	return;
}


void CMainFrame::OnNMCustomdrawButton2(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMCUSTOMDRAW pNMCD = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0x1ff;

	switch (pNMCD->dwDrawStage)
	{
	case CDDS_PREERASE:
		break;
	case CDDS_PREPAINT:
	{
		CDC dc;
		CRect rect = pNMCD->rc;
		HDC hDC = pNMCD->hdc;

#define LightGray 	RGB(240,240,240) // gray

		if (dc.Attach(hDC))
		{
			dc.FillRect(&rect, &CBrush(LightGray));
			dc.SetBkMode(TRANSPARENT);
			dc.SetTextColor(RGB(0, 0, 0));

			dc.Detach();
		}
		int deb = 1;
	}
	break;
	default:
	{
		int deb = 1;  // should not be here
	}
	}
	int deb = 1;
}

void CMainFrame::CreateMailListsInfoText(CFile &fp)
{
	CString htmlHdr;

	htmlHdr += "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=US-ASCII\"></head><body>";
	htmlHdr += "<br><font size=\"+2\"><b>Mail List Dialog Bar</b></font><br><br>";

	fp.Write((LPCSTR)htmlHdr, htmlHdr.GetLength());

	CString text;
	text.Append(
		"Mbox Viewer maintains 3 internal mail lists:<br>"
		"<ul>"
		"<li>All Mails list is populated from the selected archive file under Mail Tree.</li>"
		"<li>Find Mails list is populated by the search results. User can run Find Advanced dialog or set the Find All option in the Find dialog.</li>"
		"<li>User Selected Mails list is composed by a user from the mails in the All Mails and Find Mails lists.</li>"
		"</ul>"
		"Each internal mail list has associated button in the dialog bar located next to the tool bar. When a particular mail list is shown in the Mail Summary Window, associated button is highlighted.<br>"
		"<br>"
		"Access to the User Selected Mails list is disabled upon startup and the associated Button gray out. It can be enabled by the user to perform simple list auditing if desired by selecting View->User Selected Mail List to enable/disable.<br>"
		"<br>"
		"When User Selected Mails list is enabled, additional mail menu options will also be enabled such as Copy Selected Mails to User Selected Mails.<br>"
		"<br>"
		"Content of the User Selected Mails list is controlled by the user. User can merge search results with the content of the User Selected List.<br>"
		"<br>"
		"User can run the search multiples time and merge results multiple times with the User Selected List. Search results can be pruned before merging.<br>"
		"<br>"
		"In addition, user has an option to select/highlight one or more mails in the Summary Mail Window and copy to the User Selected Mails list.<br>"
		"<br>"
		"Merging/copying process will not create duplicate mails in the User Selected List.<br>"
		"<br>"
		"User can select and highlight one or more mails in the Summary Mail Window and remove from the active list. Mails can't be removed from the All Mails list.<br>"
		"<br>"
		"All Mails list content persists until new mail archive is selected.<br>"
		"<br>"
		"Find Mails list content persists until new search or when All Mails list is invalidated.<br>"
		"<br>"
		"User Selected Mails list content persists until cleared by the user or when All Mails list is invalidated.<br>"
		"<br>"
		"<font size=\"+1\"><b>Mail List Archiving</b></font><br>"
		"<br>"
		"Find Mails list and User Selected Mails list content can be saved to new mbox archive files.<br>"
		"<br>"
		"Archive file will be created in the the same folder as the All Mails main archive.<br>"
		"<br>"
		"Archive file created from the Find Mails list will be created by appending _FIND suffix to the base name of the main archive file name.<br>"
		"<br>"
		"The _USER suffix will be appended when creating archive file from the User Selected Mails list.<br>"
		"<br>"
		"For example, if the root mbox archive file from gmail is called \"All mail Including Spam and Trash.11.09.2018.mbox\", then created archive file will be named \"All mail Including Spam and Trash.11.09.2018_USER.mbox.\"<br>"
		"<br>"
		"When creating archive file for User Selected Mails list, the mbox viewer will also create file with the list of all mails for that archive, for example \"All mail Including Spam and Trash.11.09.2018_USER.mbox.mboxlist.\"<br>"
		"<br>"
		"The mail list file allows user to reload the last archived mail list into the User Select Mails. Reload can be requested at any time including after restart of the mbox viewer.<br>"
		"<br>"
		"After mbox viewer is restarted or when User Selected Mails list is empty, User Selected Mails list can be restored when All Mails list is active.<br>"
		"<br>"
		"When User Selected Mails list is active, it can be reloaded when List is not empty.<br>"
		"<br>"
	);

	fp.Write((LPCSTR)text, text.GetLength());

	CString htmlEnd = "\r\n</body></html>";
	fp.Write((LPCSTR)htmlEnd, htmlEnd.GetLength());
}


void CMainFrame::OnViewUserselectedmails()
{
	// TODO: Add your command handler code here
	CMenu *menu = this->GetMenu();
	if (m_bUserSelectedMailsCheckSet) {
		menu->CheckMenuItem(ID_VIEW_USERSELECTEDMAILS, MF_UNCHECKED);
		m_bUserSelectedMailsCheckSet = FALSE;
		EnableMailList(IDC_EDIT_LIST, FALSE);
		//TRACE("OnViewUserselectedmails: MF_CHECKED->MF_UNCHECKED\n");

		// we disabled User Selected List, switch to archive list
		if (MboxMail::IsUserMailsSelected()) {
			NListView * pListView = GetListView();
			if (pListView) {
				pListView->SwitchToMailList(IDC_ARCHIVE_LIST);
				CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
				if (pFrame)
					pFrame->SetMailList(IDC_ARCHIVE_LIST);

				MboxMail::nWhichMailList = IDC_ARCHIVE_LIST;
			}
		}
	}
	else
	{
		menu->CheckMenuItem(ID_VIEW_USERSELECTEDMAILS, MF_CHECKED);
		m_bUserSelectedMailsCheckSet = TRUE;
		EnableMailList(IDC_EDIT_LIST, TRUE);
		//TRACE("OnViewUserselectedmails: MF_UNCHECKED->MF_CHECKED\n");
	}
}


void CMainFrame::OnUpdateViewUserselectedmails(CCmdUI *pCmdUI)
{
	pCmdUI->Enable(MboxMail::s_mails.GetSize() > 0);
}

BOOL CMainFrame::IsUserMailsListEnabled()
{
	return m_bUserSelectedMailsCheckSet;
}

void CMainFrame::OnHelpMboxviewhelp()
{
	// TODO: Add your command handler code here

	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	if (!path.IsEmpty())
	{
		if (!PathFileExist(path)) {
			return;
		}
	}
	else
		return;

	CString codePageIdsFile = "MboxviewerHelp.htm";
	CString fullPath = path + "\\" + codePageIdsFile;

	CFile fp;
	if (!fp.Open(fullPath, CFile::modeWrite | CFile::modeCreate)) {
		CString txt = _T("Could not create \"") + fullPath;
		txt += _T("\" file.\nMake sure file is not open on other applications.");
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return;
	}

	CString htmlHdr;

	htmlHdr += "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=US-ASCII\"></head><body>";
	htmlHdr += "<br><font size=\"+2\"><b>Mbox Viewer Help</b></font><br><br>";

	fp.Write((LPCSTR)htmlHdr, htmlHdr.GetLength());

	CString text;
	text.Append(
		"Please review the User Guide provided with the package and/or<br>"
		"<br>"
		"right/left single and double click on any item within the Mbox Viewer window and try all presented options.<br>"
		"<br>"
		"To get started, please install the mbox mail archive in the local folder and then select File->Select Folder menu option to open that folder.<br>"
		"<br>"
	);

	fp.Write((LPCSTR)text, text.GetLength());

	CString htmlEnd = "\r\n</body></html>";
	fp.Write((LPCSTR)htmlEnd, htmlEnd.GetLength());

	fp.Close();

	ShellExecute(NULL, _T("open"), fullPath, NULL, NULL, SW_SHOWNORMAL);
	int deb = 1;
}
