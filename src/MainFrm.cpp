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

// MainFrm.cpp : implementation of the CMainFrame class
//

#include "stdafx.h"
#include "FileUtils.h"
#include "TextUtilsEx.h"
#include "HtmlUtils.h"
#include "Profile.h"
#include "mboxview.h"

#include "Resource.h"       // main symbols
#include "MainFrm.h"
#include "MboxMail.h"
#include "ExportToCSVDlg.h"
#include "PrintConfigDlg.h"
#include "OpenContainingFolderDlg.h"
#include "ColorStyleConfigDlg.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

///////
// Kept adding and adding Print to functions but now cleanup is needed, better reusability, possible abstractions, error handling, etc
// Postponed to the next relase 1.0.3.3 since larger effort is needed
///////

ColorStylesDB CMainFrame::m_ColorStylesDB;

#define MaxShellExecuteErrorCode 32

int CMainFrame::CheckShellExecuteResult(HINSTANCE  result, HWND h, CString *filePath)
{
	int ret = 0;
	CStringW filePathW;
	if (filePath)
	{
		DWORD error;
		TextUtilsEx::Ansi2Wide(*filePath, filePathW, error);
		ret = CheckShellExecuteResult(result, h, &filePathW);
	}
	else
		ret = CheckShellExecuteResult(result, h, (CStringW*)0);

	return ret;
}

int CMainFrame::CheckShellExecuteResult(HINSTANCE  result, HWND h, CStringW *filename)
{
	if ((UINT)result <= MaxShellExecuteErrorCode) 
	{
		CString errText;
		ShellExecuteError2Text((UINT)result, errText);
		CStringW errorTextW;
		DWORD error;
		TextUtilsEx::Ansi2Wide(errText, errorTextW, error);
		if (filename) 
		{
			if (FileUtils::PathFileExistW(*filename))
			{
				errorTextW += L"\n\nFile exists. Open the file location and check file properties.\n";
				errorTextW += L"Make sure the default application is configured to open the file.\n";
			}
			else
			{
				errorTextW += L"\n\nFile doesn't exist.\n";
			}
		}
		int answer = ::MessageBoxW(h, errorTextW, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}
	return 1;
}

int CMainFrame::CheckShellExecuteResult(HINSTANCE  result, CString &errorText)
{
	if ((UINT)result <= MaxShellExecuteErrorCode) {
		ShellExecuteError2Text((UINT)result, errorText);
		return -1;
	}
	return 1;
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
	ON_COMMAND(ID_TREE_HIDE, OnTreeHide)
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
	ON_COMMAND(ID_MESSAGEWINDOW_BOTTOM, &CMainFrame::OnMessagewindowBottom)
	ON_COMMAND(ID_MESSAGEWINDOW_RIGHT, &CMainFrame::OnMessagewindowRight)
	ON_COMMAND(ID_MESSAGEWINDOW_LEFT, &CMainFrame::OnMessagewindowLeft)
	ON_COMMAND(ID_FILE_PRINTCONFIG, &CMainFrame::OnFilePrintconfig)
	ON_WM_SIZE()
	ON_UPDATE_COMMAND_UI(ID_FILE_PRINTCONFIG, &CMainFrame::OnUpdateFilePrintconfig)
	ON_COMMAND(ID_FILE_MERGEARCHIVEFILES, &CMainFrame::OnFileMergearchivefiles)
	ON_COMMAND(ID_PRINTTO_PDF, &CMainFrame::OnPrinttoPdf)
	ON_WM_CLOSE()
	//ON_BN_CLICKED(IDC_FOLDER_LIST, &CMainFrame::OnBnClickedFolderList)
	ON_COMMAND(ID_FILE_ATTACHMENTSCONFIG, &CMainFrame::OnFileAttachmentsconfig)
	ON_COMMAND(ID_FILE_COLORCONFIG, &CMainFrame::OnFileColorconfig)
	ON_MESSAGE(WM_CMD_PARAM_NEW_COLOR_MESSAGE, &CMainFrame::OnCmdParam_ColorChanged)
	ON_MESSAGE(WM_CMD_PARAM_LOAD_FOLDERS_MESSAGE, &CMainFrame::OnCmdParam_LoadFolders)
	ON_COMMAND(ID_FORDEVELOPERS_SORTBYID, &CMainFrame::OnFordevelopersSortbyid)
	ON_COMMAND(ID_FORDEVELOPERS_MEMORY, &CMainFrame::OnFordevelopersMemory)
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

CMainFrame::CMainFrame(int msgViewPosition):m_wndView(msgViewPosition)
{
	// TODO: add member initialization code here

	m_colorStyleDlg = 0;
	m_NamePatternParams.SetDflts();
	//m_csvConfig.SetDflts();
	m_csvConfig.LoadFromRegistry();
	

	m_msgViewPosition = msgViewPosition; // bottom=1
	m_newMsgViewPosition = m_msgViewPosition;
	m_bMailDownloadComplete = FALSE;  // Page download complete in CBrowser
	m_bSelectMailFileDone = FALSE;
	m_MailIndex = -1;  // Not used ??
	m_bUserSelectedMailsCheckSet = FALSE;  // User Selected List checked/unched state
	m_bEnhancedSelectFolderDlg = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "enhancedSelectFolderDialog");

	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);

	treeColWidth = 177;
	m_PlusIcon = 0;
	m_MinusIcon = 0;
	m_bIsTreeHidden = FALSE;

	m_pListView = 0;
	m_pTreeView = 0;
	m_pMsgView = 0;
}

CMainFrame::~CMainFrame()
{
	// To stop memory leaks reports by debugger
	MboxMail::ReleaseResources();
	if (m_colorStyleDlg)
		delete m_colorStyleDlg;
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

	m_pTreeView = DetTreeView();
	m_pListView = DetListView();
	m_pMsgView = DetMsgView();

#if 0
	if (!m_wndToolBar.CreateEx(this, TBSTYLE_FLAT, WS_CHILD | WS_VISIBLE | CBRS_TOP
		| CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_SIZE_DYNAMIC) ||
		!m_wndToolBar.LoadToolBar(IDR_MAINFRAME))
	{
		TRACE0("Failed to create toolbar\n");
		return -1;      // fail to create
	}
#else
	BOOL createRet = m_wndToolBar.CreateEx(this);
	BOOL lbarRet = m_wndToolBar.LoadToolBar(IDR_MAINFRAME);
	if (!createRet || !lbarRet)
	{
		TRACE0("Failed to create toolbar\n");
		return -1;      // fail to create
	}

	// TODO: Remove this if you don't want tool tips
	m_wndToolBar.SetBarStyle(m_wndToolBar.GetBarStyle() | CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_TOP | CBRS_SIZE_DYNAMIC);
#endif

	CToolBarCtrl &wndToolBarCtrl = m_wndToolBar.GetToolBarCtrl();
	CImageList* imgList = wndToolBarCtrl.GetImageList();

	DWORD bkcolor = RGB(0, 0, 0);
	m_imgListBag.Create(16, 15, ILC_COLOR16, 0, 0);
	CBitmap bmp;
	bmp.LoadBitmap(IDB_BITMAP1);
	m_imgListBag.Add(&bmp, bkcolor);

	m_PlusIcon = imgList->ExtractIcon(3);
	m_MinusIcon = m_imgListBag.ExtractIcon(4);

	m_HideIcon = imgList->ExtractIcon(4);
	m_UnHideIcon = m_imgListBag.ExtractIcon(5);

	m_bTreeExpanded = FALSE;

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
	if (m_msgViewPosition != 3)
		rebarCtl.MinimizeBand(iIndex);

	// TODO: Create seprate Mail File download status bar ?
	//m_wndStatusBar.SetWindowText("Mail Download Complete: ");

	// TODO: Delete these three lines if you don't want the toolbar to be dockable
	//m_wndToolBar.EnableDocking(CBRS_ALIGN_ANY);
	//EnableDocking(CBRS_ALIGN_ANY);
	//DockControlBar(&m_wndToolBar);

	EnableAllMailLists(FALSE);

	NListView *pListView = GetListView();
	if (pListView) {
		if (m_msgViewPosition == 1)
			pListView->m_bLongMailAddress = TRUE;
		else
			pListView->m_bLongMailAddress = FALSE;
	}

#if 0
	DWORD dwFileNameFormatSizeLimit = m_NamePatternParams.m_nFileNameFormatSizeLimit;
	BOOL retval;
	retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printFileNameSizeLimit"), dwFileNameFormatSizeLimit);
	if (retval == TRUE) {
		m_NamePatternParams.m_nFileNameFormatSizeLimit = dwFileNameFormatSizeLimit;
	}
#endif

	m_NamePatternParams.LoadFromRegistry();
	m_attachmentConfigParams.LoadFromRegistry();
	m_HdrFldConfig.LoadFromRegistry();

	ColorStyleConfigDlg *dlg = new ColorStyleConfigDlg(this);
	//dlg->Create(IDD_COLOR_STYLE_DLG, GetDesktopWindow());
	if (dlg->Create(IDD_COLOR_STYLE_DLG, this) == FALSE)
	{
		m_colorStyleDlg = 0;
	}
	else
	{
		m_colorStyleDlg = dlg;
		m_colorStyleDlg->ShowWindow(SW_HIDE);
	}

	SetIcon(m_hIcon, TRUE);			// use big icon
	SetIcon(m_hIcon, FALSE);		// Use a small icon

	PostMessageA(WM_CMD_PARAM_LOAD_FOLDERS_MESSAGE, 0, 0);

	CString m_section = CString(sz_Software_mboxview) + "\\" + "Window Placement";

	WINDOWPLACEMENT wpr;
	DWORD cb = sizeof(wpr);
	BOOL ret = CProfile::_GetProfileBinary(HKEY_CURRENT_USER, m_section, "MainFrame", (LPBYTE)&wpr, cb);
	if (ret && (cb == sizeof(wpr)))
	{
		HWND hWnd = GetSafeHwnd();
		if (::IsWindow(m_hWnd))
		{
			ret = SetWindowPlacement(&wpr);
		}
		else
			int deb = 1;
	}

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
	if (m_wndView.OnCmdMsg(nID, nCode, pExtra, pHandlerInfo)) 
	{
		// Hope all objects should be initialized by now
		if (m_bSelectMailFileDone == FALSE) {
			m_bSelectMailFileDone = TRUE;
			NTreeView *pTreeView = GetTreeView();
			if (pTreeView) {
				//pTreeView->StartTimer();
				// Added application custom message to play with application level messages
				// Sending message to append to msg queue works as well and it is cleaner I think
				pTreeView->PostMsgCmdParamFileName();
			}
			else
				int deb = 1;
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

//CString GetDateFormat(int i);

void CMainFrame::OnFileOptions()
{
	bool needRedraw = false;
	COptionsDlg d;
	d.m_bEnhancedSelectFolderDlg = m_bEnhancedSelectFolderDlg;
	if (d.DoModal() == IDOK) 
	{
		m_bEnhancedSelectFolderDlg = d.m_bEnhancedSelectFolderDlg;

		NListView *pListView = GetListView();
		NMsgView *pMsgView = GetMsgView();

		CString format = MboxMail::GetDateFormat(d.m_format);
		if (pListView->m_format.Compare(format) != 0) {
			pListView->m_format = MboxMail::GetDateFormat(d.m_format);
			needRedraw = true;
		}

		if (pListView->m_maxSearchDuration != d.m_barDelay) {
			pListView->m_maxSearchDuration = d.m_barDelay;
			DWORD barDelay = d.m_barDelay;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("progressBarDelay"), barDelay);
		}

		BOOL exportEml = (d.m_exportEML > 0) ? TRUE : FALSE;
		if (pListView->m_bExportEml != exportEml) {
			pListView->m_bExportEml = exportEml;
			CString str_exportEML = _T("n");
			if (exportEml == TRUE)
				str_exportEML = _T("y");
			CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("exportEML"), str_exportEML);
		}

		if (pMsgView->m_cnf_subj_charsetId != d.m_subj_charsetId) {
			pMsgView->m_cnf_subj_charsetId = d.m_subj_charsetId;
			DWORD charsetId = d.m_subj_charsetId;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("subjCharsetId"), charsetId);
			needRedraw = true;
		}
		if (pMsgView->m_cnf_from_charsetId != d.m_from_charsetId) {
			pMsgView->m_cnf_from_charsetId = d.m_from_charsetId;
			DWORD charsetId = d.m_from_charsetId;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("fromCharsetId"), charsetId);
			needRedraw = true;
		}
		if (pMsgView->m_cnf_to_charsetId != d.m_to_charsetId) {
			pMsgView->m_cnf_to_charsetId = d.m_to_charsetId;
			DWORD charsetId = d.m_to_charsetId;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("toCharsetId"), charsetId);
			needRedraw = true;
		}
		if (pMsgView->m_show_charsets != d.m_show_charsets) {
			pMsgView->m_show_charsets = d.m_show_charsets;
			DWORD show_charsets = d.m_show_charsets;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("showCharsets"), show_charsets);
			needRedraw = true;
		}
		if (pMsgView->m_bImageViewer != d.m_bImageViewer) {
			pMsgView->m_bImageViewer = d.m_bImageViewer;
			DWORD bImageViewer = d.m_bImageViewer;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("showCharsets"), bImageViewer);
		}
		if (pListView->m_gmtTime != d.m_bTimeType)
		{
			pListView->m_gmtTime = d.m_bTimeType;
			DWORD bTimeType = d.m_bTimeType;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("timeType"), bTimeType);
			needRedraw = true;
		}
		if (needRedraw) {
#if 0
			// TODO: Invalidate() below doesn't updates column labels, so call MarkColumns() directly
			// mboxview is basically single threaded so it should be ok to do so
			// MarkColumns() i no longr used
			//pListView->MarkColumns();  // Invalidate() below doesn't updates column labels
			//pListView->ClearDescView();
			pListView->Invalidate();
			pMsgView->Invalidate();
			pMsgView->m_browser.m_ie.Invalidate();;
			//GetMsgView()->m_browser.m_ie.Invalidate(); // TODO: changed to GetMsgView()->Invalidate();
#else
			// Will refresh summary, text and attachments
			// Hope it is safe assuming no more code below this line
			pListView->m_bApplyColorStyle = TRUE;
			pListView->Invalidate();
			pListView->SelectItem(pListView->m_lastSel);
#endif
		}
	}
}

void CMainFrame::OnFileOpen() 
{
	CString fileName;
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	CString path;

	if (m_bEnhancedSelectFolderDlg == FALSE)
	{
		path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
		CBrowseForFolder bff(GetSafeHwnd(), CSIDL_DESKTOP, IDS_SELECT_FOLDER);
		if (!path.IsEmpty())
			bff.SetDefaultFolder(path);

		bff.SetFlags(BIF_RETURNONLYFSDIRS);
		if (bff.SelectFolder()) 
		{
			path = bff.GetSelectedFolder();
			path.TrimRight("\\");
		}
		else
		{
			int deb = 1;
			return;
		}
	}
	else
	{
		INT_PTR ret = SelectFolder(path);
		if (ret == IDOK)
		{
			path.TrimRight("\\");
		}
		else
		{
			int deb = 1;
			return;
		}
	}

	path.Append("\\");

	FileUtils::SplitFilePath(path, driveName, directory, fileNameBase, fileNameExtention);
	if (directory.GetLength() <= 1)
	{
		CString txt = _T("The mbox files must be installed under a named folder\n."
			"Please create folder, move the mbox files to that folder and try again.");
		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
		return;
	}

	path.TrimRight("\\");
	AfxGetApp()->AddToRecentFileList(path);

	path.Append("\\");
	CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath", path);

	int paneId = 0;
	CString sText;
	sText.Format("Opening new mail folder ...");
	SetStatusBarPaneText(paneId, sText, TRUE);

	GetTreeView()->FillCtrl();

	GetTreeView()->m_tree.SortChildren(0);

#if 0
	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");

	DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
	CFolderPickerDialog fdlg(path);

	INT_PTR ret = fdlg.DoModal();
	if (ret == IDOK)
	{
		CString path = fdlg.GetPathName();
		if (path.GetLength() > 0)
		{
			if (path.GetAt(path.GetLength() - 1) != _T('\\'))
				path.AppendChar(_T('\\'));
		}
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath", path);
		GetTreeView()->FillCtrl();
		AfxGetApp()->AddToRecentFileList(path);
		int deb = 1;
	}
	else
	{
		int deb = 1;
	}
#endif

	sText.Format("Ready");
	SetStatusBarPaneText(paneId, sText, FALSE);
}

void CMainFrame::OnTreeHide()
{
	BOOL ret;
	int col = 0;
	int cxCur;
	int cxMin;

	NTreeView * pTreeView = GetTreeView();
	NListView * pListView = GetListView();
	NMsgView * pMsgView = GetMsgView();

	// m_msgViewPosition:: 1=bottom 2=right  3=left
	int rowTree, colTree;
	int rowList, colList;
	int rowMsg, colMsg;

	m_wndView.GetTreePosition(rowTree, colTree);
	m_wndView.GetListPosition(rowList, colList);
	m_wndView.GetMsgPosition(rowMsg, colMsg);


	CToolBarCtrl &wndToolBarCtrl = m_wndToolBar.GetToolBarCtrl();
	CImageList* imgList = wndToolBarCtrl.GetImageList();

	m_wndView.m_verSplitter.GetColumnInfo(col, cxCur, cxMin);

	if (cxCur != 0)
	{
		treeColWidth = cxCur;
	}

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	BOOL isTreeHidden = FALSE;

	pListView = pFrame->GetListView();
	if (m_pListView)
	{
		if ((m_pListView->m_which == 0) && !m_bIsTreeHidden)
			return;
	}

	isTreeHidden = IsTreeHidden();

	HICON icHide = 0;

	BOOL bIsTreeHidden = m_bIsTreeHidden;
	m_bIsTreeHidden = !m_bIsTreeHidden;

	TRACE("TreeState %d --> %d m_frameCy_TreeInHide=%d m_frameCx_TreeInHide=%d m_frameCy_TreeNotInHide=%d m_frameCx_TreeNotInHide=%d\n", 
		bIsTreeHidden, m_bIsTreeHidden, 
		pListView->m_frameCy_TreeInHide, pListView->m_frameCx_TreeInHide, 
		pListView->m_frameCy_TreeNotInHide, pListView->m_frameCx_TreeNotInHide);



	if (!bIsTreeHidden)
	{
		m_wndView.m_verSplitter.SetColumnInfo(rowTree, 0, 0);

		if (m_pListView)
		{
			if (m_msgViewPosition == 1)
				m_wndView.m_horSplitter.SetRowInfo(rowList, pListView->m_frameCy_TreeInHide, 0);
			else if (m_msgViewPosition == 2)
				m_wndView.m_horSplitter.SetColumnInfo(colList, pListView->m_frameCx_TreeInHide, 0);
			else if (m_msgViewPosition == 3)
			{
				CRect frameRect;
				pFrame->GetClientRect(frameRect);
				CSize frameSize = frameRect.Size();
				int msgWidth = frameSize.cx - 0 - pListView->m_frameCx_TreeInHide;
				if (msgWidth < 0)
					int deb = 1;
				m_wndView.m_horSplitter.SetColumnInfo(colMsg, pMsgView->m_frameCx_TreeInHide, 0);
			}
		}

		// Disable/enable expand tree button
		// Install blank icon since EnableButton() doesn't gray out
#if 0
		// Install gray empty icon
		HICON ic = imgList->ExtractIcon(5);
		if (ic)
			ret = imgList->Replace(3, ic);
#endif
		BOOL r = wndToolBarCtrl.EnableButton(ID_TREE_EXPAND, FALSE);

		icHide = m_UnHideIcon;
	}
	else
	{
		m_wndView.m_verSplitter.SetColumnInfo(rowTree, m_pTreeView->m_frameCx, 0);

		if (m_pListView)
		{
			if (m_msgViewPosition == 1)
				m_wndView.m_horSplitter.SetRowInfo(rowList, pListView->m_frameCy_TreeNotInHide, 0);
			else if (m_msgViewPosition == 2)
				m_wndView.m_horSplitter.SetColumnInfo(colList, pListView->m_frameCx_TreeNotInHide, 0);
			else if (m_msgViewPosition == 3)
			{
				CRect frameRect;
				pFrame->GetClientRect(frameRect);
				CSize frameSize = frameRect.Size();
				int msgWidth = frameSize.cx - m_pTreeView->m_frameCx - pListView->m_frameCx_TreeNotInHide;
				if (msgWidth < 0)
					int deb = 1;
				m_wndView.m_horSplitter.SetColumnInfo(colMsg, pMsgView->m_frameCx_TreeNotInHide, 0);
			}
		}

		// See above comments
#if 0
		if (m_bTreeExpanded)
			ret = imgList->Replace(3, m_PlusIcon);
		else
			ret = imgList->Replace(3, m_MinusIcon);
#endif
		BOOL r = wndToolBarCtrl.EnableButton(ID_TREE_EXPAND, TRUE);
		icHide = m_HideIcon;
	}
	if (icHide)
		ret = imgList->Replace(4, icHide);

	//m_wndView.m_horSplitter.RecalcLayout();
	m_wndView.m_verSplitter.RecalcLayout();
	//m_wndView.m_verSplitter.Invalidate();
	//m_wndView.m_horSplitter.Invalidate();

	//m_wndToolBar.Invalidate();

	//m_bIsTreeHidden = !m_bIsTreeHidden;

	int deb = 1;
}

void CMainFrame::DoOpen(CString& fpath) 
{
	CString fileName;
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;


	CString path = fpath;
	path.TrimRight("\\");
	path.Append("\\");

	FileUtils::SplitFilePath(path, driveName, directory, fileNameBase, fileNameExtention);
	if (directory.GetLength() <= 1)
	{
		CString txt = _T("The mbox files must be installed under a named folder\n."
			"Please create folder, move the mbox files to that folder and try again.");
		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
		return;
	}

	path.TrimRight("\\");
	AfxGetApp()->AddToRecentFileList(path);

	int paneId = 0;
	CString sText;
	sText.Format("Opening new mail folder ...");
	SetStatusBarPaneText(paneId, sText, TRUE);

	path.Append("\\");
	CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath", path);
	GetTreeView()->FillCtrl();

	GetTreeView()->m_tree.SortChildren(0);

	sText.Format("Ready");
	SetStatusBarPaneText(paneId, sText, FALSE);
}

NTreeView * CMainFrame::DetTreeView()
{
	return (NTreeView *)m_wndView.m_verSplitter.GetPane(0, 0);
}

NTreeView * CMainFrame::GetTreeView()
{
	return m_pTreeView;
}

NListView * CMainFrame::DetListView()
{
	if (m_msgViewPosition == 1)
		return (NListView *)m_wndView.m_horSplitter.GetPane(0, 0);
	else if (m_msgViewPosition == 2)
		return (NListView *)m_wndView.m_horSplitter.GetPane(0, 0);
	else if (m_msgViewPosition == 3)
		return (NListView *)m_wndView.m_horSplitter.GetPane(0, 1);
	else
		return (NListView *)m_wndView.m_horSplitter.GetPane(0, 0);  // should never be here ASSERT ?
}

NListView * CMainFrame::GetListView()
{
	return m_pListView;
}

NMsgView * CMainFrame::DetMsgView()
{
	if (m_msgViewPosition == 1)
		return (NMsgView *)m_wndView.m_horSplitter.GetPane(1, 0);
	else if (m_msgViewPosition == 2)
		return (NMsgView *)m_wndView.m_horSplitter.GetPane(0, 1);
	else if (m_msgViewPosition == 3)
		return (NMsgView *)m_wndView.m_horSplitter.GetPane(0, 0);
	else
		return (NMsgView *)m_wndView.m_horSplitter.GetPane(1, 0);  // should never be here ASSERT ?
}

NMsgView * CMainFrame::GetMsgView()
{
	return m_pMsgView;
}

BOOL CMainFrame::OnEraseBkgnd(CDC* pDC) 
{
	// active view will paint itself

	return TRUE;
}

void CMainFrame::OnFileExportToCsv()
{
	// TODO: Add your command handler code here

	int firstMail = 0;
	int lastMail = MboxMail::s_mails.GetCount() - 1;
	BOOL selectedMails = FALSE;
	PrintMailsToCSV(firstMail, lastMail, selectedMails);
}

void CMainFrame::PrintMailsToCSV(int firstMail, int lastMail, BOOL selectedMails)
{
	// TODO: Add your command handler code here
	ExportToCSVDlg d;

	d.m_bFrom = m_csvConfig.m_bFrom;
	d.m_bTo = m_csvConfig.m_bTo;
	d.m_bSubject = m_csvConfig.m_bSubject;
	d.m_bDate = m_csvConfig.m_bDate;
	d.m_bCC = m_csvConfig.m_bCC;
	d.m_bBCC = m_csvConfig.m_bBCC;
	d.m_bContent = m_csvConfig.m_bContent;
	d.m_bAttachmentNames = m_csvConfig.m_bAttachmentNames;
	d.m_AttachmentNamesSeparatorString = m_csvConfig.m_AttachmentNamesSeparatorString;
	d.m_dateFormat = m_csvConfig.m_dateFormat;
	d.m_bGMTTime = m_csvConfig.m_bGMTTime;
	d.m_MessageLimitString = m_csvConfig.m_MessageLimitString;
	d.m_MessageLimitCharsString = m_csvConfig.m_MessageLimitCharsString;
	d.m_nCodePageId = m_csvConfig.m_nCodePageId;

	if (d.DoModal() == IDOK) 
	{
		//CSVFILE_CONFIG csvConfig;

		m_csvConfig.m_bFrom = d.m_bFrom;
		m_csvConfig.m_bTo = d.m_bTo;
		m_csvConfig.m_bSubject = d.m_bSubject;
		m_csvConfig.m_bDate = d.m_bDate;
		m_csvConfig.m_bCC = d.m_bCC;
		m_csvConfig.m_bBCC = d.m_bBCC;
		m_csvConfig.m_bContent = d.m_bContent;
		m_csvConfig.m_bAttachmentNames = d.m_bAttachmentNames;
		m_csvConfig.m_AttachmentNamesSeparatorString = d.m_AttachmentNamesSeparatorString;
		m_csvConfig.m_dateFormat = d.m_dateFormat;
		m_csvConfig.m_bGMTTime = d.m_bGMTTime;
		m_csvConfig.m_MessageLimitString = d.m_MessageLimitString;
		m_csvConfig.m_MessageLimitCharsString = d.m_MessageLimitCharsString;

		m_csvConfig.SaveToRegistry();
	
		//csvConfig.m_separator = d.m_separator;
		// Hardcoded for now.
		m_csvConfig.m_separator = ",";

		m_csvConfig.m_nCodePageId = 0;
		if (d.m_bEncodingType == 1)
			m_csvConfig.m_nCodePageId = CP_UTF8;
		else if (d.m_bEncodingType == 2) {
			m_csvConfig.m_nCodePageId = d.m_nCodePageId;
		}

		CString csvFileName;
		//int firstMail = 0;
		//int lastMail = MboxMail::s_mails.GetCount() - 1;
		BOOL progressBar = TRUE;

		NListView * pListView = GetListView();

		MailIndexList *selectedMailsIndexList = 0;
		if (selectedMails && pListView)
		{
			if ((selectedMailsIndexList = pListView->PopulateSelectedMailsList()) > 0) {
				//selectedMailsIndexList = &m_selectedMailsList;
				firstMail = 0;
				lastMail = 0;
			}
		}

		int ret = MboxMail::exportToCSVFile(m_csvConfig, csvFileName, firstMail, lastMail, selectedMailsIndexList, progressBar);
		if (ret > 0) 
		{
			CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
			if (!path.IsEmpty())  // not likely since the path was valid in MboxMail::exportToCSVFile(csvConfig);
			{
				if (FileUtils::PathDirExists(path)) { // likely :) 
					CString text = "Created file\n\n" + csvFileName;
					OpenContainingFolderDlg dlg(text);
					INT_PTR nResponse = dlg.DoModal();
					if (nResponse == IDOK)
					{
						if (FileUtils::BrowseToFile(csvFileName) == FALSE) {
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
						CheckShellExecuteResult(result, h, &csvFileName);
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

	CString HelpPath = FileUtils::GetmboxviewTempPath("MboxHelp");

	BOOL createDirOk = TRUE;
	if (!FileUtils::PathDirExists(HelpPath))
		createDirOk = CreateDirectory(HelpPath, NULL);

	int ret = TextUtilsEx::showCodePageTable(HelpPath);
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
	PrintMailArchiveToTEXT();
}

void CMainFrame::OnPrinttoTextFile(int textType)
{
	// TODO: Add your command handler code here
	if (MboxMail::s_mails.GetSize() == 0) 
	{
		CString txt = _T("Please open mail file first.");
		HWND h = GetSafeHwnd(); 
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
			textConfig.m_dateFormat = MboxMail::GetDateFormat(0);
			textConfig.m_bGMTTime = 0;
		}
		textConfig.m_nCodePageId = CP_UTF8;
#else
		textConfig.m_dateFormat = MboxMail::GetDateFormat(d.m_dateFormat);
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
		BOOL progressBar = TRUE;  // run in worker thread 
		MailIndexList *selectedMailsIndexList = 0;
		int ret = MboxMail::exportToTextFile(textConfig, textFileName, firstMail, lastMail, selectedMailsIndexList, textType, progressBar);

		if (ret > 0) 
		{
			CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
			if (!path.IsEmpty())  // not likely since the path was valid in MboxMail::exportToTextFile(...);
			{
				if (FileUtils::PathDirExists(path)) { // likely :) 
					CString txt = "Created file\n\n" + textFileName;
					OpenContainingFolderDlg dlg(txt);
					INT_PTR nResponse = dlg.DoModal();
					if (nResponse == IDOK)
					{
						if (FileUtils::BrowseToFile(textFileName) == FALSE) {
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
						CheckShellExecuteResult(result, h, &textFileName);
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
int CMainFrame::OnPrintSingleMailtoText(int mailPosition, int textType, CString &createdTextFileName, BOOL forceOpen, BOOL printToPrinter, BOOL createFileOnly)  // textType 0==Plain, 1==Html
{
	// TODO: Add your command handler code here
	if (MboxMail::s_mails.GetSize() == 0) {
		CString txt = _T("Please open mail file first.");
		HWND h = GetSafeHwnd(); 
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
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
			textConfig.m_dateFormat = MboxMail::GetDateFormat(0);
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
		BOOL progressBar = FALSE;  // no worker thread
		CString textFileName;
		MailIndexList *selectedMailsIndexList = 0;
		ret = MboxMail::exportToTextFile(textConfig, textFileName, mailPosition, mailPosition, selectedMailsIndexList, textType, progressBar);
		if (ret > 0)
		{
			CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
			if (!path.IsEmpty())  // not likely since the path was valid in MboxMail::exportToTextFile(...);
			{
				if (FileUtils::PathDirExists(path))
				{ // likely :) 
					CString txt = "Created file\n\n" + textFileName;
					if (createFileOnly) {
						int deb = 1;
					}
					else if (printToPrinter)
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
						OpenContainingFolderDlg dlg(txt);
						INT_PTR nResponse = dlg.DoModal();
						if (nResponse == IDOK)
						{
							if (FileUtils::BrowseToFile(textFileName) == FALSE) {
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
							CheckShellExecuteResult(result, h, &textFileName);
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
						CheckShellExecuteResult(result, h, &textFileName);
						int deb = 1;
					}
				}
				else
					;
			}
			else
				; // TODO:

			createdTextFileName.Append(textFileName);
			return 1;
		}
		else
		{
			createdTextFileName.Empty();
			return -1;
		}
	}
#endif
}

int  CMainFrame::PrintSingleMailtoHTML(int iItem, CString &targetPrintSubFolderName, CString &errorText)
{
	CString textFileName;

	TEXTFILE_CONFIG textConfig;
	NListView * pListView = CMainFrame::GetListView();
	if (pListView)
	{
		textConfig.m_dateFormat = pListView->m_format;
		textConfig.m_bGMTTime = pListView->m_gmtTime;
	}
	else
	{
		textConfig.m_dateFormat = MboxMail::GetDateFormat(0);
		textConfig.m_bGMTTime = 0;
	}
	textConfig.m_nCodePageId = CP_UTF8;

	int textType = 1;
	int ret = MboxMail::PrintMailRangeToSingleTextFile(textConfig, textFileName, iItem, iItem, textType, targetPrintSubFolderName, errorText);

	return 1;
}

int  CMainFrame::PrintSingleMailtoPDF(int iItem, CString &targetPrintSubFolderName, BOOL progressBar, CString &errorText)
{
	CString textFileName;

	TEXTFILE_CONFIG textConfig;
	NListView * pListView = CMainFrame::GetListView();
	if (pListView) 
	{
		textConfig.m_dateFormat = pListView->m_format;
		textConfig.m_bGMTTime = pListView->m_gmtTime;
	}
	else
	{
		textConfig.m_dateFormat = MboxMail::GetDateFormat(0);
		textConfig.m_bGMTTime = 0;
	}
	textConfig.m_nCodePageId = CP_UTF8;

	int textType = 1;
	int ret = MboxMail::PrintMailRangeToSingleTextFile(textConfig, textFileName, iItem, iItem, textType, targetPrintSubFolderName, errorText);

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	CString fileName;
	
	fileName.Append(textFileName);
	CString newSuffix;
	newSuffix.Append(".pdf");
	FileUtils::UpdateFileExtension(fileName, newSuffix);

	CString htmFileName;
	htmFileName.Append(textFileName);

	CString progressText;
	ret = ExecCommand_WorkerThread(htmFileName, errorText, progressBar, progressText);
	return ret;
}

// File->"Print To"->HTML->OnPrinttoHtml()->OnPrinttoTextFile()->exportToTextFile() for firstMail; lastMail->printSingleMailToTextFile(index)
void CMainFrame::OnPrinttoHtml()
{
	// TODO: Add your command handler code here
	PrintMailArchiveToHTML();
}

void CMainFrame::PrintMailArchiveToHTML()
{
	CString errorText;
	if (m_NamePatternParams.m_bPrintToSeparateHTMLFiles)
	{
		int firstMail = 0;
		int lastMail = MboxMail::s_mails.GetCount() - 1;

		NListView * pListView = GetListView();

		if (pListView)
		{
			int ret = pListView->PrintMailArchiveToSeparateHTML_Thread(errorText);
		}
		int deb = 1;
	}
	else
	{
		OnPrinttoTextFile(1);
	}
}

void CMainFrame::PrintMailArchiveToTEXT()
{
	if (m_NamePatternParams.m_bPrintToSeparateTEXTFiles)
	{
		OnPrinttoTextFile(0);
	}
	else
	{
		OnPrinttoTextFile(0);
	}
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

void CMainFrame::SortByColumn(int column)
{
	NListView *pListView = 0;
	pListView = GetListView();
	if (pListView)
	{
		pListView->SortByColumn(column);
	}
}


void CMainFrame::OnByfrom()
{
	// TODO: Add your command handler code here
	SortByColumn(2);
}


void CMainFrame::OnUpdateByfrom(CCmdUI *pCmdUI)
{
	// TODO: Add your command update UI handler code here
	pCmdUI->Enable(MboxMail::s_mails.GetSize() > 0);
}


void CMainFrame::OnByto()
{
	// TODO: Add your command handler code here
	SortByColumn(3);
}


void CMainFrame::OnUpdateByto(CCmdUI *pCmdUI)
{
	// TODO: Add your command update UI handler code here
	pCmdUI->Enable(MboxMail::s_mails.GetSize() > 0);
}


void CMainFrame::OnBysubject()
{
	// TODO: Add your command handler code here
	SortByColumn(4);
}


void CMainFrame::OnUpdateBysubject(CCmdUI *pCmdUI)
{
	// TODO: Add your command update UI handler code here
	pCmdUI->Enable(MboxMail::s_mails.GetSize() > 0);
}


void CMainFrame::OnBysize()
{
	// TODO: Add your command handler code here
	SortByColumn(5);
}


void CMainFrame::OnUpdateBysize(CCmdUI *pCmdUI)
{
	// TODO: Add your command update UI handler code here
	pCmdUI->Enable(MboxMail::s_mails.GetSize() > 0);
}


void CMainFrame::OnByconversation()
{
	// TODO: Add your command handler code here
	SortByColumn(0);
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
#if 0
	// Doesn't always work
	pCmdUI->SetText(strPage);
#else
	CMainFrame *pFrame = (CMainFrame*)AfxGetMainWnd();
	if (pFrame)
	{
		if (pFrame) {
			int paneId = 1;
			pFrame->SetStatusBarPaneText(paneId, strPage, FALSE);
		}
	}
#endif
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
	// Just test of large mailCount
	//mailCnt = 173245;

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
	MboxMail::s_mails_folder.SetSizeKeepData(0);


	if (MboxMail::b_mails_which_sorted != 1)
		int deb = 1; // TODO: ASSERT ?


	MboxMail::m_allMails.m_lastSel = -1;
	MboxMail::m_allMails.b_mails_which_sorted = 1; // by date
	MboxMail::m_findMails.m_lastSel = -1;
	MboxMail::m_findMails.b_mails_which_sorted = 1; // not really, shuld be unknown yet
	MboxMail::m_editMails.m_lastSel = -1;
	MboxMail::m_editMails.b_mails_which_sorted = 1; // not really, shuld be unknown yet
	MboxMail::m_folderMails.m_lastSel = -1;
	MboxMail::m_folderMails.b_mails_which_sorted = 1; // not really, shuld be unknown yet

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
	//int nIDLastButton = IDC_FOLDER_LIST;
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

		p = m_wndDlgBar.GetDlgItem(IDC_FOLDER_LIST);
		if (p) {
			((CButton*)p)->SetCheck(0);
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
		*pResult = CDRF_DODEFAULT;
		int deb = 1;
	}
	break;
	default:
	{
		*pResult = CDRF_DODEFAULT;
		int deb = 1;  // should not be here
	}
	}
	int deb = 1;
}


void CMainFrame::OnBnClickedButton2()  // help button on tool bar
{
	// TODO: Add your control notification handler code here

	CString HelpPath = FileUtils::GetmboxviewTempPath("MboxHelp");

	BOOL createDirOk = TRUE;
	if (!FileUtils::PathDirExists(HelpPath))
		createDirOk = CreateDirectory(HelpPath, NULL);

	CString codePageIdsFile = "MailListsInfo.htm";
	CString fullPath = HelpPath + "\\" + codePageIdsFile;

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
	*pResult = 0x1ff;  // ?? Explain

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

	htmlHdr += "<u>NOTE: Please review also important help information under \"View -> Mboxview Help\" option</u></font><br><br>";

	fp.Write((LPCSTR)htmlHdr, htmlHdr.GetLength());

	CString text;
	text.Append(
		"Mbox Viewer maintains 3 internal mail lists:<br>"
		"<ul>"
		"<li>All Mails list is populated from the selected archive file under Mail Tree.</li>"
		"<li>Found Mails list is populated by the search results. User can run Find Advanced dialog or set the Find All option in the Find dialog.</li>"
		"<li>User Selected Mails list is composed by a user from the mails in the All Mails and Found Mails lists. Mails on All Mails and Found Mails lists are marked by the red vertical bar in the first column if they are also on User Selected Mails list.</li>"
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
		"Found Mails list content persists until new search or when new mail archive is selected.<br>"
		"<br>"
		"User Selected Mails list content persists until cleared by the user or when new mail archive is selected.<br>"
		"<br>"
		"<u>NOTE</u> that mails in the All Mails and Found Mails list will be marked with the vertical bar in the first column to indicate that the same mail is present in the User Selected Mails list.<br>"
		"<br>"
	);
	fp.Write((LPCSTR)text, text.GetLength());

	text.Empty();
	text.Append(
		"<font size=\"+1\"><b>Mail List Archiving</b></font><br>"
		"<br>"
		"Found Mails list and User Selected Mails list content can be saved to mbox archive files and to mboxlist mail list files.<br>"
		"<br>"
		"Archive file will be created in the ROOT_MAIL_DIRECTORY\\ArchiveCache\\MailArchiveFileX directory.<br>"
		"<br>"
		"Archive file created from the Found Mails list will be created by appending _FIND suffix to the base name of the main archive file name.<br>"
		"<br>"
		"The _USER suffix will be appended when creating archive file from the User Selected Mails list.<br>"
		"<br>"
		"For example, if the root mbox archive file from gmail is called \"All mail Including Spam and Trash.11.09.2018.mbox\", then created archive file will be named \"All mail Including Spam and Trash.11.09.2018_USER.mbox.\"<br>"
		"<br>"
		"Separately, user can save mails in User Select Mails list to .mboxlist file, for example \"All mail Including Spam and Trash.11.09.2018_USER.mbox.mboxlist\".<br>"
		"<br>"
		"The .mboxlist files are much smaller than the mail archive files. The final .mboxlist file can be reloaded at any time and then the archive file can be created.\"<br>"
		"<br>"
		"WARNING: The .mboxlist file will no longer be valid if the master mail archive is changed.<br>"
		"<br>"
		"The mail list file allows user to reload the last archived mail list into the User Select Mails. Reload can be requested at any time including after restart of the mbox viewer.<br>"
		"<br>"
		"After mbox viewer is restarted or when User Selected Mails list is empty, User Selected Mails list can be restored when All Mails list is active.<br>"
		"<br>"
		"When User Selected Mails list is active, it can be reloaded when List is not empty.<br>"
		"<br>"
	);

	fp.Write((LPCSTR)text, text.GetLength());

	text.Empty();
	text.Append(
		"<font size=\"+1\"><b>Managing Multiple Archives</b></font><br>"
		"<br>"
		"Mutiple archives can be created from the the same single main (or master) mail archive. However, before you create new content and new archive from Found Mails or User Selected List, the last archive with hard-coded name must be renamed first.<br>"
		"<br>"
		"Click on Open Mail Archive Location and rename the last hard-coded mail archive, for example:<br>"
		"<br>"
		"\"All mail Including Spam and Trash.11.09.2018_USER.mbox.\" --> \"All mail from Amazon.11.09.2018_USER.mbox.\"<br>"
		"<br>"
		"Not likely but if desired, multiple mboxlist files can be maintained by following similar steps.<br>"
		"<br>"
		"You must restore mboxlist file to original name if you plan to continue to update the mail list at later date. For example:<br>"
		"<br>"
		"\"All mail from Amazon.11.09.2018_USER.mbox.mboxlist\"  ---copy or --> \"All mail Including Spam and Trash.11.09.2018_USER.mbox.mboxlist\"<br>"
		"<br>"
		"Future releases may address the issue better.<br>"
		"<br>"
	);
	fp.Write((LPCSTR)text, text.GetLength());

	text.Empty();
	text.Append(
		"<font size=\"+1\"><b>Merging Multiple Mail Archives</b></font><br>"
		"<br>"
		"Multiple archives can be concatenated into a single archive using \"File -> Merge Archive Files\" option.<br>"
		"<br>"
		"This is useful if you need to analyze multiple mail views, provided as separate archive files, derived from the same archive file.<br>"
		"<br>"
		"After merging is completed, select \"File -> Refresh\" to discover new archive by the mbox viewer.<br>"
		"<br>"
		"In order to remove duplicate mails from the concatenated file:"
		"<br>"
		"<ol>"
		"<li>Select the created archive.</li>"
		"<br>"
		"<li>Select \"View -> User Selected Mails\" to enable User Select Mails list.</li>"
		"<br>"
		"<li>Right click on any mail to select \"Copy All into User Selected Mails\" option to copy all mails into User Select Mails list.</li>"
		"<br>"
		"<li>Click on the  \"User Selected Mails\" button to select User Select Mails list.</li>"
		"<br>"
		"<li>Right click on any mail to select \"Remove Duplicate Mails\" option to remove duplicate  mails from User Select Mails list.</li>"
		"<br>"
		"<li>Mails are considered duplicate if the Data, From, To header fields and unique message ID generated by the mail provider match.</li>"
		"<br>"
		"<li>Select \"Save as Mail Archive file\" option to save mails in the User Selecected Mails list.</li>"
		"<br>"
		"<li>Select \"File -> Refresh\" to discover new no duplicate mails archive file by the mbox viewer.</li>"
		"<br>"
		"<li>You may want to rename the default name assigned to the archive file before the above step.</li>"
		"</ol>"
		"<br>"
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

	CString HelpPath = FileUtils::GetmboxviewTempPath("MboxHelp");

	BOOL createDirOk = TRUE;
	if (!FileUtils::PathDirExists(HelpPath))
		createDirOk = CreateDirectory(HelpPath, NULL);

	CString codePageIdsFile = "MboxviewerHelp.htm";
	CString fullPath = HelpPath + "\\" + codePageIdsFile;

	CFile fp;
	if (!fp.Open(fullPath, CFile::modeWrite | CFile::modeCreate)) {
		CString txt = _T("Could not create \"") + fullPath;
		txt += _T("\" file.\nMake sure file is not open on other applications.");
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return;
	}

	CString htmlHdr;

	htmlHdr += "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=US-ASCII\"></head><body style=\'margin-left:0.5cm\'>";
	htmlHdr += "<br><font size=\"+2\"><b>Mbox Viewer Help</b></font><br><br>";

	fp.Write((LPCSTR)htmlHdr, htmlHdr.GetLength());

	CString text;

	text.Empty();
	text.Append(
		"Please review the User Guide provided with the package and/or "
		"right/left single and double click on any item within the Mbox Viewer window and try all presented options.<br>"
		"<br>"
		"To get started, please install the mbox mail archive in the local folder,"
		" install and left double click on the mboxview executable and then select File->Select Folder menu option to open that folder.<br>"
		"<br>"
	);
	fp.Write((LPCSTR)text, text.GetLength());

	text.Empty();
	text.Append(
		"<font size=\"+2\"><u><b>Mail Printing Overview</b></u></font><br>"
		"<br>"
		"Mbox Viewer supports direct printing of all, single or multiple mails to CSV, Text, HTML, PDF files and to PDF printer.<br>"
		"<br>"
		"Mails can also be printed to PDF from any Web Browser by opening mails printed to HTML files.<br>"
		"<br>"
		"By default all, single or multiple mails are printed to single CSV, Text, HTML and PDF files without any additional configuration or manual steps.<br>"
		"<br>"
		"However there is a limit how many mails can effectively be printed to a single file.<br>"
		"<br>"
		"Option is provided to print mails to separate files to support larger number of mails.<br>"
		"<br>"
		"The next sections below will cover these limitations and alternate solutions.<br>"
		"<br><br>"
	);
	fp.Write((LPCSTR)text, text.GetLength());

	text.Empty();
	text.Append(
		"<font size=\"+1\"><b>Mail Printing to CSV Spreasheet file</b></font><br>"
		"<br>"
		"Mbox Viewer supports printing of all mails or selected group of mails to single CSV file only.<br>"
		"<br>"
		"All mails can be printed to single CSV file but a particular Spreadsheet tool may limit the maximum supported size of the spreadsheet file.<br>"
		"<br>"
		"User can select and print groups of mails to separate CVS files as a work around.<br>"
		"<br>"
		"There is no option to automatically split all mails into configurable number of mail groups for printing to separate CSV files<br>"
		"<br>"
	);
	fp.Write((LPCSTR)text, text.GetLength());

	text.Empty();
	text.Append(
		"<font size=\"+1\"><b>Mail Printing to TEXT file</b></font><br>"
		"<br>"
		"Mbox Viewer supports printing of all mail, single or selected group of mails to single Text file only.<br>"
		"<br>"
		"All mails can be printed to single TEXT file but a particular Text viewer tool may limit the maximum supported size of the text file.<br>"
		"<br>"
		"User can select and print groups of mails to separate Text files as a work around.<br>"
		"<br>"
		"There is no option to automatically split all mails into configurable number of mail groups for printing to separate Text files<br>"
		"<br>"
	);
	fp.Write((LPCSTR)text, text.GetLength());

	text.Empty();
	text.Append(
		"<font size=\"+1\"><b>Mail Printing to HTML file</b></font><br>"
		"<br>"
		"Mbox Viewer supports printing of all mails, single or selected group of mails to single HTML file or to separate per mail HTML file.<br>"
		"<br>"
		"All archive mails can be printed to single HTML file but HTML viewer tools/Web Browsers may not be able to process large HTML files, become very slow or completely overloaded.<br>"
		"<br>"
		"User can select and print groups of mails to separate HTML files as a work around.<br>"
		"<br>"
		"There is no option to automatically split all mails into configurable number of mail groups for printing to separate HTML files<br>"
		"<br>"
		"There is no hard rule how many mails can be printed to a single HTML file and viewed by a Web Browser.<br>"
		"<br>"
		"It depends on the size and content (such as heavy graphics) of mails. It should be doable to print up to a couple of thousands small to medium size text mails to a single HTML file.<br>"
		"<br>"
		"Mails can be printed to a separate per mail HTML file for further processing as described in Mail Printing to PDF section.<br>"
		"<br>"
	);
	fp.Write((LPCSTR)text, text.GetLength());

	text.Empty();
	text.Append(
		"<font size=\"+1\"><b>Mail Printing to PDF printer</b></font><br>"
		"<br>"
		"Mbox Viewer supports printing of all mails, single or selected group of mails to PDF printer.<br>"
		"<br>"
		"Printing to PDF printer relies on Microsoft HTML Document object to print its content.<br>"
		"<br>"
		"First, Mbox Viewer will print mails to a single HTML file, then it will load the file into HTML Document object and request the document object to print its content.<br>"
		"<br>"
		"The limitations as far as the maximum number of mails that can be printed are described in the \"Printing Mails to HTML file\" section above.<br>"
		"<br>"
		"The \"File -> Print Config -> Page Setup\" dialog option allows users to control <u>the page title, header, footer and background color</u>.<br>"
		"<br>"
		"By default user is prompted to select PDF printer for printing mails.<br>"
		"<br>"
		"If the PDF printer is configured as the default printer, user can set \"File -> Print Config -> Do Not Prompt\" dialog option to skip the printer prompt.<br>"
		"<br>"
	);
	fp.Write((LPCSTR)text, text.GetLength());

	text.Empty();
	text.Append(
		"<font size=\"+1\"><b>Mail Printing to PDF files</b></font><br>"
		"<br>"
		"Mbox Viewer supports multiple ways to print all, single or multiple mails to PDF files.<br>"
		"<br>"
		"All methods convert  HTML files, created first from mails by the Mbox Viewer, to PDF format.<br>"
		"<br>"
		"By default all, single or selected mails are printed to a single HTML file and then converted to PDF format.<br>"
		"<br>"
		"The default method limits the number of mails that can be effectively printed to single HTML file as noted in previous sections.<br>"
		"<br>"
		"More scalable method is described in in \"Mail Printing to PDF files by Power Users\" section.<br>"
		"<br>"
		"Mbox Viewer supports an option to print mails directly to PDF file without any user interaction.<br>"
		"<br>"
		"Direct printing to PDF file invokes external application to convert HTML files to PDF.<br>"
		"<br>"
		"By default the standard Google Chrome browser in so called headless mode is invoked to perform conversion.<br>"
		"<br>"
		"The two options are provided to leverage developer version of Google Chrome Canary and free wkhtmltopdf application to covert HTML to PDF.<br>"
		"<br>"
		"Options to leverage the wkhtmltopdf and Chrome Canary are provided since the standard Chrome browser in the headless mode doesn't support options to control the page title, header and footer and the background color.<br>"
		"<br>"
		"User can set \"File -> Print Config -> Path To User Defined Script\" to invoke HTML2PDF-single-wkhtmltopdf.cmd script included in the release package.<br>"
		"<br>"
		"In order to leverage Chrome Canary browser, \"Path To User Defined Script\" needs to updated by the user to point to HTML2PDF-single-chrome-canary.cmd script.<br>"
		"<br>"
		"The Chrome Canary browser can be downloaded from <a href=\"https://www.google.com/chrome/canary\">Chrome Canary downloads</a>.<br>"
		"<br>"
		"HTML2PDF-single-wkhtmltopdf.cmd script creates PDF files with <u>the right footer \"Page Number of Total Pages\" and no page Title and Header</u>.<br>"
		"<br>"
		"The wkhtmltopdf can be downloaded from <a href=\"https://wkhtmltopdf.org/downloads.html\">wkhtmltopdf.org downloads</a>.<br>"
		"<br>"
		"The wkhtmltopdf command line options are documented in <a href=\"https://wkhtmltopdf.org/usage/wkhtmltopdf.txt\">wkhtmltopdf.org usage</a>.<br>"
		"<br>"
		"User can replicate HTML2PDF-single-wkhtmltopdf.cmd to new file and customize to leverage different converter possibly commercial.<br>"
		"<br>"
		"User will need to update \"Path To User Defined Script\" in \"File -> Print Config\" dialog to point to new script path.<br>"
		"<br>"
		"When printing multiple mails to a single PDF file, user can set \"File -> Print Config -> Page Break After Each Mail\" option <u>to start each mail at new page</u>.<br>"
		"<br>"
		"User can unset \"File -> Print Config -> Mail Hdr Background Color\" option to remove mail header background color and save on the ink when printing.<br>"
		"<br>"
		"Lastly, HTML file can be open within Firefox, Chrome, IE, etc. Browser and printed to PDF file.<br>"
		"<br>"
	);
	fp.Write((LPCSTR)text, text.GetLength());

	text.Empty();
	text.Append(
		"<font size=\"+1\"><b>Mail Printing to PDF files by Power Users</b></font><br>"
		"<br>"
		"Some users may need to print larger number of mails to a single PDF file.<br>"
		"<br>"
		"This is supported by Mbox Viewer but one or more manual steps are required.<br>"
		"<br>"
		"User must first print all mails to separate PDF files and then merge them into a single PDF file by running the provided script.<br>"
		"<br>"
		"In order to print mails to separate files, user must first set \"File -> Print Config -> Print mails to separate files -> PDF\" option.<br>"
		"<br>"
		"After that, using standard mail print to PDF options, user can print all or large subset of mails to separate PDF files into sub-folder PDF_GROUP created by Mbox Viewer.<br>"
		"<br>"
		"Printing large number of mails to separate PDF files is time consuming.<br>"
		"<br>"
		"Depending on the size and content of a mail, it may take fraction of a second to create single PDF file or a tens of seconds.<br>"
		"<br>"
		"PDFMerge-pdfbox.cmd script included in the release package can be used to merge PDF files.<br>"
		"<br>"
		"PDFMerge-pdfbox.cmd script invokes free PDFBox java tool to merge PDF files.<br>"
		"<br>"
		"Free PDFBox java tool can be downloaded from <a href=\"https://pdfbox.apache.org/\">pdfbox.apache.org</a>.<br>"
		"<br>"
		"PDFBox command line tools usage example <a href=\"https://pdfbox.apache.org/2.0/commandline.html\">pdfbox.apache.org commandline tools</a>.<br>"
		"<br>"
		"Java 8 can be dowwnloaded from  <a href=\"https://www.oracle.com/technetwork/java/javase/downloads/jdk8-downloads-2133151.html\">jdk8-downloads</a>.<br>"
		"<br>"
		"PDFMerge-pdfbox.cmd script must be copied into the directory housing all per mail PDF files and executed.<br>"
		"<br>"
		"PDFMerge-pdfbox.cmd script will create one or more merged PDF files in the PDF_MERGE sub-folder.<br>"
		"<br>"
		"If more than one merged PDF file is created in the PDF_MERGE sub-folder, PDFMerge-pdfbox.cmd script needs to be copied to that sub-folder and script executed again.<br>"
		"<br>"
		"The above steps need to be repeated until PDFMerge-pdfbox.cmd script creates single PDF file only.<br>"
		"<br>"
		"Printing mails to separate PDF files can also be done offline. User must first set \"File -> Print Config -> Print mails to separate files -> HTML\" option.<br>"
		"<br>"
		"After that, using standard mail print to HTML options, user can print all or large subset of mails to separate HTML files into sub-folder HTML_GROUP created by Mbox Viewer.<br>"
		"<br>"
		"HTML2PDF-all-chrome.cmd or HTML2PDF-all-wkhtmltopdf.cmd script included in the release package can then be used to convert HTML files to PDF files.<br>"
		"<br>"
		"One of these scripts must be copied into the directory housing all per mail HTML files, i.e. HTML_GROUP, and executed.<br>"
		"<br>"
		"Created PDF files can then be merged into a single PDF file using PDFMerge-pdfbox.cmd script as described above.<br>"
		"<br>"
	);
	fp.Write((LPCSTR)text, text.GetLength());

	text.Empty();
	text.Append(
		"<font size=\"+2\"><u><b>Advanced Find Overview</b></u></font><br>"
		"<br>"
		"Advanced Find dialog allows users to use multiple fields to compose more complex search criteria.<br>"
		"<br>"
		"The search logic is hard-coded for the most common case and is as follow:<br>"
		"<br>"
		"(unidirectional or bidirectional and operator of From and To) and Subject and CC and BCC and (Message or Attachments)<br>"
		"<br>"
		"The search criteria is basically ANDing all checked fields or pairs of fields such as (From and To) and (Message and Attachment).<br>"
		"<br>"
		"Find Advanced dialog allows users to specify relation between From and To as biderectional or unidirectional and the result is ANDed with other checked fields.<br>"
		"<br>"
		"The Message and Attachment are handled as OR expression and the result is ANDed with other checked fields.<br>"
		"<br>"
		"More complex search criteria can be accomplished by leveraging User Selected Mails list.<br>"
		"<br>"
		"User can run Advanced Find multiple times and merge results into User Selected Mails list.<br>"
		"<br><br>"
	);
	fp.Write((LPCSTR)text, text.GetLength());

	text.Empty();
	text.Append(
		"<font size=\"+2\"><u><b>Directory Structure</b></u></font><br>"
		"<br>"
		"Mbox Viewer creates and manages several sub-directories in the directory that houses one or more mail archive files.<br>"
		"<br>"
		"ROOT_MAIL_DIRECTORY - <font color=red>target directory for mbox mail files, mbox mail index files, help Html files and various sub-directories</font><br>"
		"ROOT_MAIL_DIRECTORY\\MailArchiveFile1.mbox<br>"
		"ROOT_MAIL_DIRECTORY\\MailArchiveFile1.mbox.mboxview<br>"
		"ROOT_MAIL_DIRECTORY\\MailArchiveFile2.mbox<br>"
		"ROOT_MAIL_DIRECTORY\\MailArchiveFile2.mbox.mboxview<br>"
		"ROOT_MAIL_DIRECTORY\\ImageCache<br>"
		"ROOT_MAIL_DIRECTORY\\ImageCache\\MailArchiveFile1 - <font color=red>target directory for image files, such as png,jpg,etc, embeded into mails</font><br>"
		"ROOT_MAIL_DIRECTORY\\ImageCache\\MailArchiveFile2<br>"
		"ROOT_MAIL_DIRECTORY\\ArchiveCache<br>"
		"ROOT_MAIL_DIRECTORY\\ArchiveCache\\MailArchiveFile1 - <font color=red>target directory for saving Found Mails and User Selected Mails to mbox and mboxlist files</font><br>"
		"ROOT_MAIL_DIRECTORY\\ArchiveCache\\MailArchiveFile2<br>"
		"ROOT_MAIL_DIRECTORY\\PrintCache<br>"
		"ROOT_MAIL_DIRECTORY\\PrintCache\\MailArchiveFile1  -  <font color=red>target directory for printing to single CSV,TEXT,HTML and PDF files</font><br>"
		"ROOT_MAIL_DIRECTORY\\PrintCache\\MailArchiveFile1\\PDF_GROUP  -  <font color=red>target directory for printing to separate PDF files</font><br>"
		"ROOT_MAIL_DIRECTORY\\PrintCache\\MailArchiveFile1\\PDF_GROUP\\PDF_MERGE  -  <font color=red>target directory for merged PDF files</font><br>"
		"ROOT_MAIL_DIRECTORY\\PrintCache\\MailArchiveFile1\\PDF_GROUP\\PDF_MERGE\\PDF_MERGE  -  <font color=red>target directory for merged PDF files if multiple merge steps are needed</font><br>"
		"ROOT_MAIL_DIRECTORY\\PrintCache\\MailArchiveFile1\\HTML_GROUP  -  <font color=red>target directory for printing to separate HTML files</font><br>"
		"ROOT_MAIL_DIRECTORY\\PrintCache\\MailArchiveFile1\\HTML_GROUP\\PDF_MERGE  -  <font color=red>target directory for merged PDF files</font><br>"
		"ROOT_MAIL_DIRECTORY\\PrintCache\\MailArchiveFile1\\HTML_GROUP\\PDF_MERGE\\PDF_MERGE  -  <font color=red>target directory for merged PDF files if multiple merge steps are needed</font><br>"
		"ROOT_MAIL_DIRECTORY\\PrintCache\\MailArchiveFile2<br>"
		"<br>"
		"<br>"
		"Mbox Viewer also creates and manages temporary directory C:\\Users\\UserName\\AppData\\Local\\Temp\\mboxview directory to store temporary files (attachments, eml and htm) created when a single mail is selected by the user.<br>"
		"Mbox Viewer created help files such as MailListsInfo.htm and MboxviewerHelp.htm are stored in the temporary directory C:\\Users\\UserName\\AppData\\Local\\Temp\\mboxview\\MboxHelp directory.<br>"
		"<br><br>"
	);
	fp.Write((LPCSTR)text, text.GetLength());

	text.Empty();
	text.Append(
		"<br><br>"
		"<br><br>"
	);
	fp.Write((LPCSTR)text, text.GetLength());

	CString htmlEnd = "\r\n</body></html>";
	fp.Write((LPCSTR)htmlEnd, htmlEnd.GetLength());

	fp.Close();

	ShellExecute(NULL, _T("open"), fullPath, NULL, NULL, SW_SHOWNORMAL);
	int deb = 1;
}


void CMainFrame::OnMessagewindowBottom()
{
	// TODO: Add your command handler code here
	ConfigMessagewindowPosition(1);
}


void CMainFrame::OnMessagewindowRight()
{
	// TODO: Add your command handler code here
	ConfigMessagewindowPosition(2);
}


void CMainFrame::OnMessagewindowLeft()
{
	// TODO: Add your command handler code here
	ConfigMessagewindowPosition(3);
}

void CMainFrame::ConfigMessagewindowPosition(int msgViewPosition)
{
	if ((msgViewPosition < 1) && (msgViewPosition > 3)) {
		int deb = 1;
		return;  // should never be here
	}

	int curMsgViewPosition = m_msgViewPosition;

	CString curPos = "Bottom";
	if (curMsgViewPosition == 1)
		curPos = "Bottom";
	else if (curMsgViewPosition == 2)
		curPos = "Right";
	else if (curMsgViewPosition == 3)
		curPos = "Left";
	else
		; // shpould never be here

	CString newPos = "Bottom";
	if (msgViewPosition == 1)
		newPos = "Bottom";
	else if (msgViewPosition == 2)
		newPos = "Right";
	else if (msgViewPosition == 3)
		newPos = "Left";
	else
		; // shpould never be here


	HWND h = GetSafeHwnd();

	if (msgViewPosition == curMsgViewPosition) {
		CString txt = _T("New and current Message Window position are the same: ") + newPos + _T(" and ") + curPos + _T(".\nNo action will be taken.");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONEXCLAMATION | MB_OK);
		m_newMsgViewPosition = msgViewPosition;
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("messageWindowPosition"), m_newMsgViewPosition);
		return;
	}

	CString txt = _T("Do you want to configure Message Window position from " + curPos + " to " + newPos 
		+ " ?\nIf you say Yes, please exit and restart mbox viewer for new position to take effect.");

	int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
	if (answer == IDYES) 
	{
		m_newMsgViewPosition = msgViewPosition;
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("messageWindowPosition"), m_newMsgViewPosition);
	}
}


void CMainFrame::OnFilePrintconfig()
{
	// TODO: Add your command handler code here

	PrintConfigDlg dlg;

	dlg.m_NamePatternParams.Copy(m_NamePatternParams);

	if (dlg.DoModal() == IDOK)
	{
		m_NamePatternParams.UpdateRegistry(m_NamePatternParams, dlg.m_NamePatternParams);
		m_NamePatternParams.Copy(dlg.m_NamePatternParams);
	}

	MboxMail::ShowHint(HintConfig::PrintToPDFScriptHint, GetSafeHwnd());
}


void CMainFrame::OnSize(UINT nType, int cx, int cy)
{
	CFrameWnd::OnSize(nType, cx, cy);

	// TODO: Add your message handler code here
	int deb1, deb2, deb3, deb4;

	if (nType == SIZE_RESTORED)
		int deb = 1;

	switch (nType)
	{
	case SIZE_MAXIMIZED:
		// window was maximized
		deb1 = 1;
		break;

	case SIZE_MINIMIZED:
		// window was minimized
		deb2 = 1;
		break;

	case SIZE_RESTORED:
		// misleading - this occurs when restored from minimized/maximized AND
		// for normal size operations when already restored
		deb3 = 1;
		break;

	default:
		// you could also deal with SIZE_MAXHIDE and SIZE_MAXSHOW
		// but rarely need to
		deb4 = 1;
		break;
}

	// TODO: resize panes dynamically; more investigation and changes needed
#if 0
	if (nType == 2) 
	{
		if (m_msgViewPosition == 2) 
		{
			NListView *pListView = GetListView();
			NTreeView *pTreeView = GetTreeView();

			CRect rect;
			m_wndView.GetWindowRect(&rect);
			m_wndView.m_horSplitter.GetWindowRect(&rect);

			m_wndView.m_horSplitter.SetColumnInfo(0, rect.Width() - 700, 20);
			//m_wndView.m_horSplitter.SetColumnInfo(1, 700, 20);

			m_wndView.m_horSplitter.RecalcLayout();
		}
	}
#endif
}


void CMainFrame::OnUpdateFilePrintconfig(CCmdUI *pCmdUI)
{
	// TODO: Add your command update UI handler code here
	pCmdUI->Enable(MboxMail::s_mails.GetSize() > 0);
	// TODO: Update code to always enable Print Config
	//pCmdUI->Enable(TRUE);
}

void CMainFrame::UpdateFilePrintconfig()
{
	NamePatternParams::UpdateFilePrintconfig(m_NamePatternParams);
}

int CMainFrame::MergeArchiveFiles()
{
	// TODO: customize CFileDialog to avoid potential buffer overflow and corruption
#define MAX_CFileDialog_FILE_COUNT 1000
	int FILE_LIST_BUFFER_SIZE = (MAX_CFileDialog_FILE_COUNT * (MAX_PATH + 1)) + 1;

	MboxMail::m_outbuf->ClearAndResize(FILE_LIST_BUFFER_SIZE);
	TCHAR *fileNameBuffer = MboxMail::m_outbuf->Data();

	CString  sFilters  = "Mail Files (*.mbox)|*.mbox|All Files (*.*)|*.*||";

	DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
	CFileDialog dlgFile(TRUE, NULL, NULL, dwFlags, sFilters);

	CString inFolderPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
#if 0
	// No need to enforce the below
	if (!inFolderPath.IsEmpty())
	{
		if (!FileUtils::PathDirExists(inFolderPath)) {
			return -1;
		}
	}
	else {
		return -1;
	}
#endif

	OPENFILENAME& ofn = dlgFile.GetOFN();
	ofn.Flags |= OFN_ALLOWMULTISELECT;
	ofn.lpstrFile = fileNameBuffer;
	ofn.nMaxFile = FILE_LIST_BUFFER_SIZE;
	ofn.lpstrInitialDir = inFolderPath;
	ofn.lpstrTitle = "Select Mail Archives For Merging";

restart:
	INT_PTR ret = dlgFile.DoModal();
	if (ret == IDOK)
	{
		CString folder = dlgFile.GetFolderPath();
		TRACE("FOLDER=%s\n", folder);

		int fileCount = 0;
		CString strFilePath;
		POSITION pos = dlgFile.GetStartPosition();
		while (NULL != pos) 
		{
			fileCount++;
			strFilePath = dlgFile.GetNextPathName(pos);
			if (!FileUtils::PathFileExist(strFilePath))
			{
				CString txt = _T("File path invalid.\n");
				txt.Append(strFilePath);
				txt.Append("\nRetry or Cancel?");

				int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_RETRYCANCEL);
				if (answer == IDRETRY)
				{
					goto restart;
				}
			}
			TRACE("FILE=%s\n", strFilePath);
		}

		// All archive files are valid; merge
		CString title = "Enter Name for New Archive File";
		CString  fileNameFilter = "Mail Files (*.mbox)|*.mbox|All Files (*.*)|*.*||";
		CString dfltExtension = "mbox";

		CString outFolderPath;
		CString fileName;
		BOOL ret = SaveFileDialog(fileName, fileNameFilter, dfltExtension, inFolderPath, outFolderPath, title);
		CString filePath = outFolderPath + "\\" + fileName;
		CString fileExtension = ::PathFindExtension(fileName);
		CString fileName2 = ::PathFindFileName(filePath);
		//FileUtils::SplitFilePath(CString &fileName, CString &driveName, CString &directory, CString &fileNameBase, CString &fileNameExtention);
		if (ret == TRUE) 
		{
			CFile fp;
			if (!fp.Open(filePath, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone)) {
				CString txt = _T("Could not create \"") + filePath;
				txt += _T("\" file.\nMake sure file is not open on other applications.");
				HWND h = NULL; // we don't have any window yet
				int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
				return -1;
			}

			pos = dlgFile.GetStartPosition();
			while (NULL != pos)
			{
				strFilePath = dlgFile.GetNextPathName(pos);
				// Check again anyway
				if (FileUtils::PathFileExist(strFilePath))
				{
					CFile fp_input;
					if (!fp_input.Open(strFilePath, CFile::modeRead)) {
						CString txt = _T("Could not create \"") + strFilePath;
						txt += _T("\" file.\nMake sure file is not open on other applications.");
						HWND h = NULL; // we don't have any window yet
						int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);

						fp.Close();
						return -1;
					}
					// Append file
					UINT wantBytes = 16 * 1024;
					// don't use  MboxMail::m_oubuf is still used by dlgFile above !!!!! 
					MboxMail::m_inbuf->ClearAndResize(wantBytes+2);
					TCHAR *inBuffer = MboxMail::m_inbuf->Data();

					UINT readBytes = 0;
					do {
						readBytes = fp_input.Read(inBuffer, wantBytes);
						fp.Write(inBuffer, readBytes);
					} while (readBytes > 0);

#if 0
					// not sure whether below is neeeded in case no CR NL or just NL at the file end
					// It would change total length on new archive which should not be an issue 
					// unless someone would try to add all lengths manually to verify
					MboxMail::m_inbuf->Clear();
					MboxMail::m_inbuf->Append("\r\n"); 
					fp.Write(inBuffer, readBytes);
#endif

					fp_input.Close();

				}
				TRACE("FILE=%s\n", strFilePath);
			}
			fp.Close();
			int deb = 1;
		}
	}
	else
	{
		int deb = 1;
	}
	return 1;
}

BOOL CMainFrame::SaveFileDialog(CString &fileName, CString &fileNameFilter, CString &dfltExtention, CString &inFolderPath, CString &outFolderPath, CString &title)
{
	// TODO: customize CFileDialog to avoid potential buffer overflow and corruption
	int FILE_LIST_BUFFER_SIZE = (2 * (MAX_PATH + 1)) + 1;

	MboxMail::m_inbuf->ClearAndResize(FILE_LIST_BUFFER_SIZE);
	TCHAR *fileNameBuffer = MboxMail::m_inbuf->Data();

	DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
	CFileDialog dlgFile(FALSE, dfltExtention, NULL, dwFlags, fileNameFilter);

	OPENFILENAME& ofn = dlgFile.GetOFN();
	//ofn.Flags |= ??;
	ofn.lpstrFile = fileNameBuffer;
	ofn.nMaxFile = FILE_LIST_BUFFER_SIZE;
	ofn.lpstrInitialDir = inFolderPath;
	ofn.lpstrTitle = title;

	INT_PTR ret = dlgFile.DoModal();
	if (ret == IDOK)
	{
		outFolderPath = dlgFile.GetFolderPath();
		fileName = dlgFile.GetFileName();
		TRACE("FOLDER=%s\n", outFolderPath);
		TRACE("FILE=%s\n", fileName);
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}


void CMainFrame::OnFileMergearchivefiles()
{
	// TODO: Add your command handler code here
	MergeArchiveFiles();
}


void CMainFrame::OnPrinttoPdf()
{
	// TODO: Add your command handler code here
	if (MboxMail::s_mails.GetSize() == 0) {
		CString txt = _T("Please open mail file first.");
		HWND h = GetSafeHwnd(); // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;
	}

	NListView *pListView = GetListView();
	int firstMail = 0;
	int lastMail = MboxMail::s_mails.GetCount() - 1;
	if (m_NamePatternParams.m_bPrintToSeparatePDFFiles)
	{
		if (pListView)
		{
			CString targetPrintSubFolderName = "PDF_GROUP";
			pListView->PrintMailRangeToSeparatePDF_Thread(firstMail, lastMail, targetPrintSubFolderName);
		}
	}
	else
	{
		CString targetPrintSubFolderName;
		pListView->PrintMailRangeToSinglePDF_Thread(firstMail, lastMail, targetPrintSubFolderName);
	}

	int deb = 1;
}

void CMainFrame::PrintMailArchiveToPDF()
{
	OnPrinttoPdf();
}

int CMainFrame::VerifyPathToHTML2PDFExecutable(CString &errorText)
{
	CString path;
	if (m_NamePatternParams.m_bScriptType == 0)
	{
		path = m_NamePatternParams.m_ChromeBrowserPath;
		if (!FileUtils::PathFileExist(path))
		{
			errorText = _T("Path to Chrome Browser not valid.\nPlease make sure Chrome is installed and path is correct.\nSelect File->Print Config to update the setup.");
			return -1;
		}
	}
	else
	{
		path = m_NamePatternParams.m_UserDefinedScriptPath;
		if (!FileUtils::PathFileExist(path))
		{
			errorText = _T("Path to user defined HTML2PDF script not valid.\nPlease make sure script is installed.\nSelect File->Print Config to update the setup.");
			return -1;
		}
	}
	return 1;
}

int CMainFrame::ExecCommand_WorkerThread(CString &htmFileName, CString &errorText, BOOL progressBar, CString &progressText)
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	// TODO: Duplicate check, done already in _Thread function calls 
	if (pFrame)
	{
		int ret = pFrame->VerifyPathToHTML2PDFExecutable(errorText);
		if (ret < 0)
		{
			return -1;
		}
	}
	else
	{
		errorText.Append("Internal error. Try again.");
		return -1;
	}

	CString pdfFileName;
	pdfFileName.Append(htmFileName);

	CString newSuffix;
	newSuffix.Append(".pdf");
	FileUtils::UpdateFileExtension(pdfFileName, newSuffix);

	BOOL delStatus = DeleteFile(pdfFileName);
	if (delStatus == FALSE) {
		DWORD error = GetLastError();
	}

	int bScriptType = 0;
	CString path = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe";
	CString args = "--headless --disable-gpu --print-to-pdf=\"" + pdfFileName + "\" \"" + htmFileName + "\"";

	if (pFrame->m_NamePatternParams.m_bScriptType == 0)
	{
		// args no change
		path = pFrame->m_NamePatternParams.m_ChromeBrowserPath;
		if (!FileUtils::PathFileExist(path)) {
			errorText = _T("Path to Chrome Browser not valid.\nPlease make sure Chrome is installed and path is correct.\nSelect File->Print Config to update setup.");
			int deb = 1;
			return -1;
		}
	}
	else
	{
		if (pFrame->m_NamePatternParams.m_bKeepMailBodyBackgroundColor)
			args = "\"" + htmFileName + "\"";
		else
			args = "\"" + htmFileName + "\"" + " " + "--no-background";
		path = pFrame->m_NamePatternParams.m_UserDefinedScriptPath;
		if (!FileUtils::PathFileExist(path)) {
			errorText = _T("Path to user defined HTML2PDF script not valid.\nPlease make sure script is installed.\nSelect File->Print Config to update setup.");
			int deb = 1;
			return -1;
		}
	}

	HINSTANCE result = S_OK;
	SHELLEXECUTEINFO ShExecInfo = { 0 };
	ShExecInfo.cbSize = sizeof(SHELLEXECUTEINFO);
	ShExecInfo.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NO_CONSOLE;
	ShExecInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
	ShExecInfo.hwnd = NULL;
	ShExecInfo.lpVerb = NULL;
	ShExecInfo.lpFile = path;
	ShExecInfo.lpParameters = args;
	ShExecInfo.lpDirectory = NULL;
	ShExecInfo.nShow = SW_HIDE;
	ShExecInfo.hInstApp = result;
	BOOL retval = ShellExecuteEx(&ShExecInfo);
	if (retval == FALSE)
	{
		DWORD err = GetLastError();
		int ret = CMainFrame::CheckShellExecuteResult(ShExecInfo.hInstApp, errorText);
		return -1;
	}

	int step = 10;
	int stepCnt = 0;
	int nSeconds = 0;
	CString secondsBar;

	DWORD msec = 100;
	BOOL signaled = FALSE;
	BOOL failed = FALSE;
	for (;;)
	{
		msec = 100;
		DWORD ret = WaitForSingleObject(ShExecInfo.hProcess, msec);
		switch (ret)
		{
		case WAIT_ABANDONED: {
			failed = TRUE;
			break;
		}
		case WAIT_OBJECT_0: {
			signaled = TRUE;
			break;
		}
		case WAIT_FAILED: {
			failed = TRUE;
			break;
		}
		case WAIT_TIMEOUT: {
			if (MboxMail::pCUPDUPData)
			{

				if (MboxMail::pCUPDUPData->ShouldTerminate())
				{
					if (ShExecInfo.hProcess)
					{
						TerminateProcess(ShExecInfo.hProcess, IDCANCEL);
						CloseHandle(ShExecInfo.hProcess);
					}
					return 2;
				}
				if (progressBar)
				{
					if (stepCnt % 10 == 0)
					{
						nSeconds++;
						if (progressText.IsEmpty())
							MboxMail::pCUPDUPData->SetProgress(step);
						else {
							secondsBar.Format(_T("%s  %d seconds"), progressText, nSeconds);
							MboxMail::pCUPDUPData->SetProgress(secondsBar, step);
						}
						step += 10;
						if (step > 100)
							step = 10;
					}
					stepCnt++;
				}
			}
			break;
		}
		default: {
			failed = TRUE;
			break;
		}
		}
		if (signaled || failed)
		{
			// ???
			break;
		}
	}
	if (ShExecInfo.hProcess)
		CloseHandle(ShExecInfo.hProcess);

	return 1;
}

void CSVFILE_CONFIG::SetDflts()
{
	// This duplicate alsoe set in ExportToCSVDlg
	m_bFrom = TRUE;
	m_bTo = TRUE;
	m_bSubject = TRUE;
	m_bDate = TRUE;
	m_bCC = FALSE;
	m_bBCC = FALSE;
	m_bContent = FALSE;
	m_bAttachmentNames = FALSE;
	//m_bAttachmentNames = TRUE;
	//m_MessageLimitString;
	m_MessageLimitCharsString = "32500";
	m_AttachmentNamesSeparatorString = " CRLF";
	m_dateFormat = 0;
	m_bGMTTime = 0;
	m_nCodePageId = CP_UTF8;
	//CString m_separator;

	//m_bEncodingType = 1;  // UTF8  ??

}

void CSVFILE_CONFIG::Copy(CSVFILE_CONFIG &src)
{
	if (this == &src)
		return;

	m_bFrom = src.m_bFrom;
	m_bTo = src.m_bTo;
	m_bSubject = src.m_bSubject;
	m_bDate = src.m_bDate;
	m_bCC = src.m_bCC;
	m_bBCC = src.m_bBCC;
	m_bContent = src.m_bContent;
	m_bAttachmentNames = src.m_bAttachmentNames;
	m_AttachmentNamesSeparatorString = src.m_AttachmentNamesSeparatorString;
	m_dateFormat = src.m_dateFormat;
	m_bGMTTime = src.m_bGMTTime;
	m_MessageLimitString = src.m_MessageLimitString;
	//
	m_nCodePageId = src.m_nCodePageId;
	m_separator = src.m_separator;

	int deb = 1;
}

void CSVFILE_CONFIG::SaveToRegistry()
{
	CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "CSV_attachmentNamesSeparator", m_AttachmentNamesSeparatorString);
}

void CSVFILE_CONFIG::LoadFromRegistry()
{
	SetDflts();

	BOOL retval;
	if (retval = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "CSV_attachmentNamesSeparator", m_AttachmentNamesSeparatorString))
	{
		;
	}
}

INT_PTR CMainFrame::SelectFolder(CString &folder)
{
	// TODO: customize CFileDialog to avoid potential buffer overflow and corruption
#define MAX_CFileDialog_FCOUNT 10
	int FILE_LIST_BUFFER_SIZE = (MAX_CFileDialog_FCOUNT * (MAX_PATH + 1)) + 1;

	MboxMail::m_outbuf->ClearAndResize(FILE_LIST_BUFFER_SIZE);
	TCHAR *fileNameBuffer = MboxMail::m_outbuf->Data();

	CString  sFilters = "Mail Files (*.mbox;*eml)|*.mbox;*.eml||";

	DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
	MySelectFolder dlgFile(TRUE, NULL, NULL, dwFlags, sFilters);

	CString inFolderPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");

	OPENFILENAME& ofn = dlgFile.GetOFN();
	//ofn.Flags |= OFN_ALLOWMULTISELECT;
	ofn.lpstrFile = fileNameBuffer;
	ofn.nMaxFile = FILE_LIST_BUFFER_SIZE;
	ofn.lpstrInitialDir = inFolderPath;
	ofn.lpstrTitle = "Select Folder with Mail Archives";

	INT_PTR ret = dlgFile.DoModal();
	if (ret == IDOK)
	{
		folder = dlgFile.GetFolderPath();
		TRACE("FOLDER=%s\n", folder);

		int fileCount = 0;
		CString strFilePath;
		POSITION pos = dlgFile.GetStartPosition();
		while (NULL != pos)
		{
			fileCount++;
			strFilePath = dlgFile.GetNextPathName(pos);

			TRACE("FILE=%s\n", strFilePath);

			BOOL retval = FileUtils::CPathGetPath(strFilePath, folder);
			break;
		}
	}
	return ret;
}

MySelectFolder::MySelectFolder(
		BOOL bOpenFileDialog, LPCTSTR lpszDefExt, LPCTSTR lpszFileName,
		DWORD dwFlags, LPCTSTR lpszFilter, CWnd* pParentWnd, DWORD dwSize, BOOL bVistaStyle
		) : CFileDialog(bOpenFileDialog, lpszDefExt, lpszFileName, dwFlags, lpszFilter, pParentWnd, dwSize, bVistaStyle)
{
	int deb = 1;

}

int CMainFrame::CountMailFilesInFolder(CString &folder, CString &extension)
{
	// Break id at least one file exists
	CString filePath;
	int fileCnt = 0;
	CString fn;

	CString fw = folder + extension;
	WIN32_FIND_DATA	wf;
	HANDLE f = FindFirstFile(fw, &wf);
	if (f != INVALID_HANDLE_VALUE)
	{
		do
		{
			if ((wf.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == FILE_ATTRIBUTE_DIRECTORY)
				continue;
			fileCnt++;
			break;
		} while (FindNextFile(f, &wf));
		FindClose(f);
	}
	return fileCnt;
}

BOOL MySelectFolder::OnFileNameOK()
{
	CString folder = GetFolderPath();
	TRACE("FOLDER=%s\n", folder);

	int fileCount = 0;
	CString strFilePath;
	POSITION pos = GetStartPosition();
	while (NULL != pos)
	{
		fileCount++;
		strFilePath = GetNextPathName(pos);

		TRACE("FILE=%s\n", strFilePath);

		BOOL retval = FileUtils::CPathGetPath(strFilePath, folder);
		break;
	}

	CString extension = "\\*.mbox";
	int fcnt = CMainFrame::CountMailFilesInFolder(folder, extension);
	if (fcnt == 0)
	{
		extension = "\\*.eml";
		fcnt = CMainFrame::CountMailFilesInFolder(folder, extension);
	}
	if (fcnt > 0)
		return FALSE;
	else
		return TRUE;
}

void CMainFrame::OnClose()
{
	// TODO: Add your message handler code here and/or call default
	if ((MboxMail::IsUserMailsSelected() && (MboxMail::s_mails.GetCount() > 0)) || (MboxMail::s_mails_edit.GetCount() > 0))
	{
		CString txt = _T("User Selelected Mails List not empty. Terminate program?");
		int answer = MessageBox(txt, _T("Warning"), MB_APPLMODAL | MB_ICONWARNING | MB_YESNO);
		if (answer == IDNO)
			return;
	}

	CString m_section = CString(sz_Software_mboxview) + "\\" + "Window Placement";
	WINDOWPLACEMENT wp;
	GetWindowPlacement(&wp);

	if (wp.showCmd == SW_MAXIMIZE)
		wp.showCmd = SW_SHOWMAXIMIZED; 
	if ((wp.showCmd != SW_SHOWMAXIMIZED) && (wp.showCmd != SW_SHOW))
		wp.showCmd = SW_SHOW;


	BOOL ret = CProfile::_WriteProfileBinary(HKEY_CURRENT_USER, m_section, "MainFrame", (LPBYTE)&wp, sizeof(wp));


	NTreeView * pTreeView = GetTreeView();
	NListView * pListView = GetListView();
	NMsgView * pMsgView = GetMsgView();

	if (m_msgViewPosition == m_newMsgViewPosition)
	{
		if (pMsgView)
		{
			ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section, "MsgFrameTreeNotHiddenWidth", pMsgView->m_frameCx_TreeNotInHide);
			ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section, "MsgFrameTreeNotHiddenHeight", pMsgView->m_frameCy_TreeNotInHide);

			ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section, "MsgFrameTreeHiddenWidth", pMsgView->m_frameCx_TreeInHide);
			ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section, "MsgFrameTreeHiddenHeight", pMsgView->m_frameCy_TreeInHide);
		}

		if (pListView)
		{
			ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section, "ListFrameTreeNotHiddenWidth", pListView->m_frameCx_TreeNotInHide);
			ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section, "ListFrameTreeNotHiddenHeight", pListView->m_frameCy_TreeNotInHide);

			ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section, "ListFrameTreeHiddenWidth", pListView->m_frameCx_TreeInHide);
			ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section, "ListFrameTreeHiddenHeight", pListView->m_frameCy_TreeInHide);
		}

		if (pTreeView)
		{
			ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section, "TreeFrameWidth", m_pTreeView->m_frameCx);
			ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section, "TreeFrameHeight", m_pTreeView->m_frameCy);
		}
	}
	else
		DeleteAllPlacementKeys();

	MboxMail::ReleaseResources();

	CFrameWnd::OnClose();
}

void CMainFrame::SetStatusBarPaneText(int paneId, CString &sText, BOOL setColor)
{
	// Implements synchrous Set Text capability to make sure Text is set and 
	// visible before long function is called (such as sort column)
	// Command UI may not work in this case. TODO: verify
	if (this->GetSafeHwnd())
	{
		if (m_wndStatusBar.GetSafeHwnd())
		{
			if (setColor == FALSE) 
			{
				m_wndStatusBar.GetStatusBarCtrl().SetText(sText, paneId, 0);
			}
			else
			{
				// Is below a hack ?, not sure -:) seems to work
				CClientDC dc(&m_wndStatusBar);

				CRect r;
				m_wndStatusBar.GetItemRect(paneId, &r);

				COLORREF bgRet;
				bgRet = dc.SetBkColor(::GetSysColor(COLOR_3DFACE));

				COLORREF newBgCol = dc.GetBkColor();
				dc.FillSolidRect(r, newBgCol);

				r.DeflateRect(4, 1);

				COLORREF bgColo = dc.GetBkColor();
				COLORREF txtColo = dc.GetTextColor();
				CFont *curFont = dc.GetCurrentFont();
				CFont newFont;
				newFont.CreatePointFont(85, _T("Tahoma"));

				COLORREF crRet = 0;
				if (setColor)
					crRet = dc.SetTextColor(RGB(0, 0, 200));

				// Set new font
				CFont  *pOldFont = dc.SelectObject(&newFont);

				UINT nFormat = DT_EXTERNALLEADING | DT_SINGLELINE | DT_VCENTER;
				int retVal = dc.DrawText(sText, &r, nFormat);

				// Restore oldFont
				if (pOldFont) CFont *pRetFornt = dc.SelectObject(pOldFont);
			}
			int deb = 1;
		}
	}
	return;
}

#if 0
void CMainFrame::OnBnClickedFolderList()
{
	// TODO: Add your control notification handler code here
	int nID = IDC_FOLDER_LIST;
	if (MboxMail::nWhichMailList == nID)
		return; // do nothing

	NListView * pListView = GetListView();
	if (pListView) {
		pListView->SwitchToMailList(nID);
	}

	CWnd *p = m_wndDlgBar.GetDlgItem(IDC_FOLDER_LIST);
	if (p) {
		((CButton*)p)->SetCheck(1);
		((CButton*)p)->SetWindowText("Folder XXX Mails");
	}

	int deb = 1;
}
#endif


void CMainFrame::OnFileAttachmentsconfig()
{
	// TODO: Add your command handler code here

	AttachmentsConfig dlg;

	dlg.m_attachmentConfigParams.Copy(m_attachmentConfigParams);

	if (dlg.DoModal() == IDOK)
	{
		m_attachmentConfigParams.UpdateRegistry(m_attachmentConfigParams, dlg.m_attachmentConfigParams);
		m_attachmentConfigParams.Copy(dlg.m_attachmentConfigParams);

		// TODO: Invalidate() below doesn't updates column labels, so call MarkColumns() directly
		CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
		if (pFrame)
		{
			NListView *pListView = pFrame->GetListView();
			NMsgView *pMsgView = pFrame->GetMsgView();
			if (pListView && pMsgView)
			{
#if 0
				pListView->MarkColumns();  // Invalidate() below doesn't updates column labels - not used anymore
				pListView->ClearDescView();
				pListView->Invalidate();
				pMsgView->Invalidate();
#else
				// Will refresh summary, text and attachments
				// Hope it is safe assuming no more code below this line
				pListView->m_bApplyColorStyle = TRUE;
				pListView->Invalidate();
				pListView->SelectItem(pListView->m_lastSel);
#endif
			}
		}
	}
}

AttachmentConfigParams *CMainFrame::GetAttachmentConfigParams()
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
		return &pFrame->m_attachmentConfigParams;
	else
		return 0;
}


void CMainFrame::OnFileColorconfig()
{
	// TODO: Add your command handler code here

	if (m_colorStyleDlg)
	{
#if 1
		m_colorStyleDlg->ShowWindow(SW_SHOW);
		m_colorStyleDlg->ShowWindow(SW_RESTORE);
#else
		BOOL isVisible = m_colorStyleDlg->IsWindowVisible();
		if (!isVisible)
		{
			m_colorStyleDlg->ShowWindow(SW_SHOW);
			int deb = 1;
		}
		BOOL isIconic = m_colorStyleDlg->IsIconic();
		if (isIconic)
		{
			m_colorStyleDlg->ShowWindow(SW_RESTORE);
			int deb = 1;
		}
#endif
	}
	else
	{
		ColorStyleConfigDlg *dlg = new ColorStyleConfigDlg(this);
		//dlg->Create(IDD_COLOR_STYLE_DLG, GetDesktopWindow());
		if (dlg->Create(IDD_COLOR_STYLE_DLG, this) == FALSE)
		{
			m_colorStyleDlg = 0;
		}
		else
		{
			m_colorStyleDlg = dlg;
			m_colorStyleDlg->ShowWindow(SW_SHOW);
		}
	}
}

LRESULT CMainFrame::OnCmdParam_ColorChanged(WPARAM wParam, LPARAM lParam)
{
	static int flag = 0;

	int selectedColorStyle = (int)wParam;

	//m_ColorStylesDB.m_colorStyles.SetColorStyle(selectedColorStyle);


	NTreeView * pTreeView = GetTreeView();
	if (pTreeView)
		pTreeView->Invalidate();

	//NListView * 
	NListView *pListView = GetListView();
	NMsgView * pMsgView = GetMsgView();
	if (pListView)
	{
		pListView->Invalidate();
		if (pListView->m_lastSel >= 0)
		{
			pListView->m_bApplyColorStyle = TRUE;
			pListView->SelectItem(pListView->m_lastSel);
		}
		else if (pMsgView)
		{
			CString url = "about:blank";
			DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);

			{
				CString colorStr;
				int retC2A = NListView::Color2Str(color, colorStr);

				url = "about:<html><head><style>body{background-color: #";
				url.Append(colorStr);
				url.Append(";}</style></head><body></body><br>");
			}

			pMsgView->m_browser.Navigate(url, NULL);
			pMsgView->UpdateLayout();
			pMsgView->PostMessage(WM_PAINT);
		}
	}
	return 0;
}

BOOL CMainFrame::IsTreeHidden()
{
	int col = 0;
	int cxCur;
	int cxMin;

	m_wndView.m_verSplitter.GetColumnInfo(col, cxCur, cxMin);

	if (cxCur == 0)
		return TRUE;
	else
		return FALSE;
}

void CMainFrame::UpdateToolsBar()
{
	BOOL ret;
	CToolBarCtrl &wndToolBarCtrl = m_wndToolBar.GetToolBarCtrl();
	CImageList* imgList = wndToolBarCtrl.GetImageList();

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	HICON ic = 0;

#if 0
	BOOL isTreeHidden = FALSE;
	if (pFrame)
		isTreeHidden = pFrame->IsTreeHidden();

	if (isTreeHidden)
	{
		ic = imgList->ExtractIcon(5);
		if (ic)
			ret = imgList->Replace(3, ic);

		m_wndToolBar.Invalidate();
		return;
	}
#endif

	// must be a better way of changing icon, need to investigate
	if (m_bTreeExpanded)
	{
		ic = m_MinusIcon;
	}
	else
	{
		ic = m_PlusIcon;
	}

	if (ic)
		ret = imgList->Replace(3, ic);

	NTreeView * pTreeView = GetTreeView();
	if (pTreeView)
		pTreeView->ExpandOrCollapseTree(m_bTreeExpanded);

	m_bTreeExpanded = !m_bTreeExpanded;

	m_wndToolBar.Invalidate();
	int deb = 1;
}

LRESULT CMainFrame::OnCmdParam_LoadFolders(WPARAM wParam, LPARAM lParam)
{
	if (GetTreeView())
		GetTreeView()->LoadFolders();
	return 0;
}


BOOL CMainFrame::DeleteAllPlacementKeys()
{
	LSTATUS ret;

	CString m_section = CString(sz_Software_mboxview) + "\\" + "Window Placement";

	ret = CProfile::_DeleteKey(HKEY_CURRENT_USER, m_section, "MsgFrameTreeNotHiddenWidth");
	ret = CProfile::_DeleteKey(HKEY_CURRENT_USER, m_section, "MsgFrameTreeNotHiddenHeight");

	ret = CProfile::_DeleteKey(HKEY_CURRENT_USER, m_section, "MsgFrameTreeHiddenWidth");
	ret = CProfile::_DeleteKey(HKEY_CURRENT_USER, m_section, "MsgFrameTreeHiddenHeight");

	ret = CProfile::_DeleteKey(HKEY_CURRENT_USER, m_section, "ListFrameTreeNotHiddenWidth");
	ret = CProfile::_DeleteKey(HKEY_CURRENT_USER, m_section, "ListFrameTreeNotHiddenHeight");

	ret = CProfile::_DeleteKey(HKEY_CURRENT_USER, m_section, "ListFrameTreeHiddenWidth");
	ret = CProfile::_DeleteKey(HKEY_CURRENT_USER, m_section, "ListFrameTreeHiddenHeight");

	ret = CProfile::_DeleteKey(HKEY_CURRENT_USER, m_section, "TreeFrameWidth");
	ret = CProfile::_DeleteKey(HKEY_CURRENT_USER, m_section, "TreeFrameHeight");

	return TRUE;
}

void CMainFrame::OnFordevelopersSortbyid()
{
	// TODO: Add your command handler code here
#if 0

	NListView *pListView = 0;
	pListView = this->GetListView();

	if (pListView)
	{
		pListView->SortByColumn(1, TRUE);
	}
#endif
	int deb = 1;
}

void CMainFrame::OnFordevelopersMemory()
{
	// TODO: Add your command handler code here

	int mailCount = MboxMail::s_mails_ref.GetCount();
	int allMailsSize = MboxMail::AllMailsSizeof(mailCount);
	CString txt;
	txt.Format("Allocated Memory: %d\n", allMailsSize);

	HWND h = GetSafeHwnd();
	int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
	int deb = 1;
}
