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
#include "MimeParser.h"

#include "Resource.h"       // main symbols
#include "MainFrm.h"
#include "MboxMail.h"
#include "ExportToCSVDlg.h"
#include "PrintConfigDlg.h"
#include "OpenContainingFolderDlg.h"
#include "ColorStyleConfigDlg.h"
#include "SMTPMailServerConfigDlg.h"
#include "DataFolderConfigDlg.h"
#include "MergeRootFolderAndSubfolders.h"

#include "NTreeView.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

///////
// Kept adding and adding Print to functions but now cleanup is needed, better reusability, possible abstractions, error handling, etc
// Postponed to the next relase 1.0.3.3 since larger effort is needed
///////


MBoxViewerDB CMainFrame::m_ViewerGlobalDB;

CommandLineParms CMainFrame::m_commandLineParms;

ColorStylesDB CMainFrame::m_ColorStylesDB;

#define MaxShellExecuteErrorCode 32

MBoxViewerDB::MBoxViewerDB()
{
	m_rootDBFolder = "MBOXV";
	m_rootLocalDBFolder = "MBoxViewer";
	m_rootPrintSubFolder = "PrintCache";
	m_rootImageSubFolder = "ImageCache";
	m_rootAttachmentSubFolder = "AttachmentCache";
	m_rootArchiveSubFolder = "ArchiveCache";
	m_rootLabelSubFolder = "LabelCache";
	m_rootEmlSubFolder = "EmlCache";
	m_rootMergedSubFolder = "MergeCache";
	m_isReadOnlyMboxDataMedia = 2;  // Unknown, not set
};

CString MBoxViewerDB::GetDBFolder() {
	return m_rootDBFolder;
}
CString MBoxViewerDB::GetLocalDBFolder() {
	return m_rootLocalDBFolder;
}
CString MBoxViewerDB::GetPrintSubFolder() {
	return m_rootPrintSubFolder;
	//return m_rootDBFolder + "\\" + m_rootPrintSubFolder;
}
CString MBoxViewerDB::GetImageSubFolder() {
	return m_rootImageSubFolder;
	//return m_rootDBFolder + "\\" + m_rootImageSubFolder;
}

CString MBoxViewerDB::GetAttachmentSubFolder() {
	return m_rootAttachmentSubFolder;
	//return m_rootDBFolder + "\\" + m_rootAttachmentSubFolder;
}
CString MBoxViewerDB::GetArchiveSubFolder() {
	return m_rootArchiveSubFolder;
	//return m_rootDBFolder + "\\" + m_rootArchiveSubFolder;
}
CString MBoxViewerDB::GetLabelSubFolder() {
	return m_rootLabelSubFolder;
	//return m_rootDBFolder + "\\" + m_rootLabelSubFolder;
}
CString MBoxViewerDB::GetEmlSubFolder() {
	return  m_rootEmlSubFolder;
	//return m_rootDBFolder + "\\" + m_rootEmlSubFolder;
}
CString MBoxViewerDB::GetMergedSubFolder() {
	return  m_rootMergedSubFolder;
	//return m_rootDBFolder + "\\" + m_rootEmlSubFolder;
}

// TODO: Not used
BOOL MBoxViewerDB::IsReadOnlyMboxDataMedia(CString *path)
{
	// This is madness, so expensive
	CString rootAppDataPath = FileUtils::GetMboxviewLocalAppDataPath();

	CString dataPath;
	if (path)
		dataPath = *path;
	else
		dataPath = MboxMail::GetLastDataPath();

	if (dataPath.GetLength() <= rootAppDataPath.GetLength())
		return FALSE;

	CString leftPath = dataPath.Mid(0, rootAppDataPath.GetLength());

	if (leftPath.CompareNoCase(rootAppDataPath))
		return FALSE;
	else
		return TRUE;
}

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
	//ON_WM_KEYDOWN()
	ON_WM_ACTIVATEAPP()
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
	//ON_COMMAND(ID_PRINTTO_CSV, &CMainFrame::OnPrinttoCsv)
	//ON_COMMAND(ID_PRINTTO_TEXT, &CMainFrame::OnPrinttoText)
	//ON_COMMAND(ID_PRINTTO_HTML, &CMainFrame::OnPrinttoHtml)
	//ON_COMMAND(ID_SORTBY_DATE, &CMainFrame::OnBydate)
	//ON_UPDATE_COMMAND_UI(ID_SORTBY_DATE, &CMainFrame::OnUpdateBydate)
	//ON_COMMAND(ID_SORTBY_FROM, &CMainFrame::OnByfrom)
	//ON_UPDATE_COMMAND_UI(ID_SORTBY_FROM, &CMainFrame::OnUpdateByfrom)
	//ON_COMMAND(ID_SORTBY_TO, &CMainFrame::OnByto)
	//ON_UPDATE_COMMAND_UI(ID_SORTBY_TO, &CMainFrame::OnUpdateByto)
	//ON_COMMAND(ID_SORTBY_SUBJECT, &CMainFrame::OnBysubject)
	//ON_UPDATE_COMMAND_UI(ID_SORTBY_SUBJECT, &CMainFrame::OnUpdateBysubject)
	//ON_COMMAND(ID_SORTBY_SIZE, &CMainFrame::OnBysize)
	//ON_UPDATE_COMMAND_UI(ID_SORTBY_SIZE, &CMainFrame::OnUpdateBysize)
	//ON_COMMAND(ID_SORTBY_CONVERSATION, &CMainFrame::OnByconversation)
	//ON_UPDATE_COMMAND_UI(ID_SORTBY_CONVERSATION, &CMainFrame::OnUpdateByconversation)
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
	//ON_COMMAND(ID_FILE_MERGEARCHIVEFILES, &CMainFrame::OnFileMergearchivefiles)
	//ON_COMMAND(ID_PRINTTO_PDF, &CMainFrame::OnPrinttoPdf)
	ON_WM_CLOSE()
	//ON_BN_CLICKED(IDC_FOLDER_LIST, &CMainFrame::OnBnClickedFolderList)
	ON_COMMAND(ID_FILE_ATTACHMENTSCONFIG, &CMainFrame::OnFileAttachmentsconfig)
	ON_COMMAND(ID_FILE_COLORCONFIG, &CMainFrame::OnFileColorconfig)
	ON_MESSAGE(WM_CMD_PARAM_NEW_COLOR_MESSAGE, &CMainFrame::OnCmdParam_ColorChanged)
	ON_MESSAGE(WM_CMD_PARAM_LOAD_FOLDERS_MESSAGE, &CMainFrame::OnCmdParam_LoadFolders)
	ON_COMMAND(ID_FORDEVELOPERS_SORTBYID, &CMainFrame::OnFordevelopersSortbyid)
	ON_COMMAND(ID_FORDEVELOPERS_MEMORY, &CMainFrame::OnFordevelopersMemory)
	ON_COMMAND(ID_VIEW_MESSAGEHEADERS, &CMainFrame::OnViewMessageheaders)
	ON_COMMAND(ID_FILE_RESTOREHINTMESSAGES, &CMainFrame::OnFileRestorehintmessages)
	ON_COMMAND(ID_MESSAGEHEADERPANELAYOUT_DEFAULT, &CMainFrame::OnMessageheaderpanelayoutDefault)
	ON_COMMAND(ID_MESSAGEHEADERPANELAYOUT_EXPANDED, &CMainFrame::OnMessageheaderpanelayoutExpanded)
	ON_COMMAND(ID_FILE_SMTPMAILSERVERCONFIG, &CMainFrame::OnFileSmtpmailserverconfig)
	ON_COMMAND(ID_HELP_USERGUIDE, &CMainFrame::OnHelpUserguide)
	ON_COMMAND(ID_HELP_README, &CMainFrame::OnHelpReadme)
	ON_COMMAND(ID_HELP_LICENSE, &CMainFrame::OnHelpLicense)
	ON_COMMAND(ID_FILE_DATAFOLDERCONFIG, &CMainFrame::OnFileDatafolderconfig)
	ON_COMMAND(ID_FILE_MERGEROOTFOLDERSUB, &CMainFrame::OnFileMergerootfoldersub)
	ON_COMMAND(ID_FILE_SELECTASROOTFOLDER, &CMainFrame::OnFileSelectasrootfolder)
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
	m_bViewMessageHeaders = FALSE;
	m_bUserSelectedMailsCheckSet = FALSE;  // User Selected List checked/unched state
	m_bEnhancedSelectFolderDlg = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "enhancedSelectFolderDialog");
	m_lastPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("lastPath"));

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
	BOOL  ret;
	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
		ret = CProfile::_GetProfileBinary(HKEY_CURRENT_USER, m_section, "MainFrame_EmlPreviewMode", (LPBYTE)&wpr, cb);
	else
		ret = CProfile::_GetProfileBinary(HKEY_CURRENT_USER, m_section, "MainFrame", (LPBYTE)&wpr, cb);
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

	m_mailDB.Initialize();

	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
	{
		// TODO: ESC not accepted until window is selected
		// How to fix that so app is ready to accept ESC from keyboard upon startup
		; // SetFocus();
		//OnSetFocus(0);
	}

	CToolTipCtrl* toolTpis = m_pTreeView->m_tree.GetToolTips();
	if (toolTpis)
	{
		int w = toolTpis->GetMaxTipWidth();
		DWORD dwDuration = TTDT_INITIAL;
		int d = toolTpis->GetDelayTime(dwDuration);
		dwDuration = TTDT_RESHOW;
		d = toolTpis->GetDelayTime(dwDuration);
		dwDuration = TTDT_AUTOPOP;
		d = toolTpis->GetDelayTime(dwDuration);
		int deb = 1;
	}

	SetFocus();
	SetActiveWindow();
	SetForegroundWindow();

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

void CMainFrame::OnActivateApp(BOOL bActive, DWORD dwThreadID)
{
	if (bActive)
	{
		TRACE("OnActivateApp: ACTIVE\n");
		int deb = 1;
	}
	else
	{
		TRACE("OnActivateApp: INACTIVE\n");
		// Hack to prevent selected text from being erased when switching to differnt application
		// Doesn't work when you minimize window via minimize button
		m_wndView.SetFocus();
		int deb = 1;
	}
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

		if (pListView->m_subjectSortType != d.m_bSubjectSortType) {
			pListView->m_subjectSortType = d.m_bSubjectSortType;
			DWORD subjectSortType = d.m_bSubjectSortType;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("subjectSortType"), subjectSortType);
		}


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
#if 0
		if (pMsgView->m_cnf_cc_charsetId != d.m_cc_charsetId) {
			pMsgView->m_cnf_cc_charsetId = d.m_cc_charsetId;
			DWORD charsetId = d.m_cc_charsetId;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("ccCharsetId"), charsetId);
			needRedraw = true;
		}
		if (pMsgView->m_cnf_bcc_charsetId != d.m_bcc_charsetId) {
			pMsgView->m_cnf_bcc_charsetId = d.m_bcc_charsetId;
			DWORD charsetId = d.m_bcc_charsetId;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("bccCharsetId"), charsetId);
			needRedraw = true;
		}
#endif
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
		if (needRedraw)
		{
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
		path = MboxMail::GetLastPath();
		CBrowseForFolder bff(GetSafeHwnd(), CSIDL_DESKTOP, IDS_SELECT_FOLDER);
		if (!path.IsEmpty())
			bff.SetDefaultFolder(path);

		bff.SetFlags(BIF_RETURNONLYFSDIRS);
		if (bff.SelectFolder())
		{
			path = bff.GetSelectedFolder();
			path.TrimRight(_T("\\"));
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
			path.TrimRight(_T("\\"));
		}
		else
		{
			int deb = 1;
			return;
		}
	}

	DoOpen(path);
}

void CMainFrame::OpenFolderAndSubfolders(CString &path)
{
	NTreeView *pTreeView = GetTreeView();

	if (pTreeView)
	{
		HTREEITEM hParent = TVI_ROOT;
		pTreeView->OpenFolderAndSubfolders(hParent, path);
	}
}

void CMainFrame::OpenRootFolderAndSubfolders(CString &path)
{
	NTreeView *pTreeView = GetTreeView();

	if (pTreeView)
	{
		HTREEITEM hParent = TVI_ROOT;
		BOOL selectFolder = TRUE;
		pTreeView->DoOpenRootFolderAndSubfolders(hParent, path, selectFolder);
	}
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
		// TODO: m_pListView->m_which needs clarification; there is potential for big trouble
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

	m_wndView.m_verSplitter.RecalcLayout();

	int deb = 1;
}

void CMainFrame::DoOpen(CString& path) 
{
	NTreeView *pTreeView = GetTreeView();
	if (pTreeView)
	{
		BOOL expand = TRUE;
		pTreeView->DoOpen(path, expand);
	}
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
			CString datapath = MboxMail::GetLastDataPath();
			CString path = MboxMail::GetLastPath();
			if (!datapath.IsEmpty())  // not likely since the path was valid in MboxMail::exportToCSVFile(csvConfig);
			{
				if (FileUtils::PathDirExists(datapath)) { // likely :) 
					CString text = "Created file \n\n" + csvFileName;
					OpenContainingFolderDlg dlg(text);
					INT_PTR nResponse = dlg.DoModal();
					if (nResponse == IDOK)
					{
						if (FileUtils::BrowseToFile(csvFileName) == FALSE) {
							HWND h = GetSafeHwnd();
							HINSTANCE result = ShellExecute(h, _T("open"), datapath, NULL, NULL, SW_SHOWNORMAL);
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

	CString HelpPath = FileUtils::GetMboxviewTempPath("MboxHelp");

	BOOL createDirOk = TRUE;
	if (!FileUtils::PathDirExists(HelpPath))
		createDirOk = FileUtils::CreateDirectory(HelpPath);

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
			CString datapath = MboxMail::GetLastDataPath();
			CString path = MboxMail::GetLastPath();
			if (!datapath.IsEmpty())  // not likely since the path was valid in MboxMail::exportToTextFile(...);
			{
				if (FileUtils::PathDirExists(datapath)) { // likely :) 
					CString txt = "Created file \n\n" + textFileName;
					OpenContainingFolderDlg dlg(txt);
					INT_PTR nResponse = dlg.DoModal();
					if (nResponse == IDOK)
					{
						if (FileUtils::BrowseToFile(textFileName) == FALSE) {
							HWND h = GetSafeHwnd();
							HINSTANCE result = ShellExecute(h, _T("open"), datapath, NULL, NULL, SW_SHOWNORMAL);
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
			CString datapath = MboxMail::GetLastDataPath();
			CString path = MboxMail::GetLastPath();
			if (!datapath.IsEmpty())  // not likely since the path was valid in MboxMail::exportToTextFile(...);
			{
				if (FileUtils::PathDirExists(datapath))
				{ // likely :) 
					CString txt = "Created file \n\n" + textFileName;
					if (createFileOnly) {
						int deb = 1;
					}
					else if (printToPrinter)
					{
						CFile fp;
						CFileException ExError;
						if (fp.Open(textFileName, CFile::modeRead | CFile::shareDenyWrite, &ExError))
						{
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
						else
						{
							CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

							CString txt = _T("Could not open \"") + textFileName;
							txt += _T("\" mail file.\n");
							txt += exErrorStr;

							TRACE(_T("%s\n"), txt);
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
								HINSTANCE result = ShellExecute(h, _T("open"), datapath, NULL, NULL, SW_SHOWNORMAL);
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
	MboxMail::s_mails_label.SetSizeKeepData(0);


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
	MboxMail::m_labelMails.m_lastSel = -1;
	MboxMail::m_labelMails.b_mails_which_sorted = 1; // not really, shuld be unknown yet

	CMenu *menu = this->GetMenu();

	menu->CheckMenuItem(ID_VIEW_USERSELECTEDMAILS, MF_CHECKED);
	m_bUserSelectedMailsCheckSet = TRUE;
	EnableMailList(IDC_EDIT_LIST, TRUE);

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
	NTreeView * pTreeView = GetTreeView();
	if (pTreeView && pTreeView->m_labelView)
		nID = IDC_LABEL_LIST;
	if (MboxMail::nWhichMailList == nID)
		return; // do nothing

	NListView * pListView = GetListView();
	if (pListView) {
		pListView->SwitchToMailList(nID);
		if (pTreeView && pTreeView->m_labelView) {
			SetMailList(IDC_ARCHIVE_LIST);
		}
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

	CString HelpPath = FileUtils::GetMboxviewTempPath("MboxHelp");

	BOOL createDirOk = TRUE;
	if (!FileUtils::PathDirExists(HelpPath))
		createDirOk = FileUtils::CreateDirectory(HelpPath);

	CString codePageIdsFile = "MailListsInfo.htm";
	CString fullPath = HelpPath + "\\" + codePageIdsFile;

	CFile fp;
	CFileException ExError;
	if (!fp.Open(fullPath, CFile::modeWrite | CFile::modeCreate, &ExError))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = _T("Could not create \"") + fullPath;
		txt += _T("\" file.\n");
		txt += exErrorStr;

		//TRACE(_T("%s\n"), txt);

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
		"Each internal mail list has associated button in the dialog bar located next to the tool bar. When a particular mail list is selected and shown in the Mail Summary Window, associated button is highlighted.<br>"
		"<br>"
		"Access to the User Selected Mails list is enabled upon startup and allows users to perform simple list auditing. It can be disabled and the associated Button gray out by the user if desired by selecting View->User Selected Mail List to enable/disable.<br>"
		"<br>"
		"When User Selected Mails list is enabled, additional mail menu options are enabled such as Copy Selected Mails to User Selected Mails.<br>"
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
		"Archive file will be created in the DIRECTORY\\ArchiveCache\\MailArchiveFileX directory.<br>"
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
		"<li>Select \"Save as Mbox Mail Archive file\" option to save mails in the User Selecected Mails list.</li>"
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

	CString HelpPath = FileUtils::GetMboxviewTempPath("MboxHelp");

	BOOL createDirOk = TRUE;
	if (!FileUtils::PathDirExists(HelpPath))
		createDirOk = FileUtils::CreateDirectory(HelpPath);

	CString codePageIdsFile = "MboxviewerHelp.htm";
	CString fullPath = HelpPath + "\\" + codePageIdsFile;

	CFile fp;
	CFileException ExError;
	if (!fp.Open(fullPath, CFile::modeWrite | CFile::modeCreate, &ExError))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = _T("Could not create \"") + fullPath;
		txt += _T("\" file.\n");
		txt += exErrorStr;

		//TRACE(_T("%s\n"), txt);

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
		"MBox Viewer creates data directory and uses as target directory for files and sub-folders created by MboxViewer such as \"Print...\" export all mails as eml files, etc<br><br>"
		"The DIRECTORY is constructed as follow depending on the selected data folder when running \"File-->Data Folder Config\" dialog.<br>"
		"If the data folder is selected per each folder with mbox files, then DIRECTORY is constructed as follow (assuming the mbox folder is not read-only).<br>"
		"DIRECTORY = PathToFolderHosingMboxFiles\\MBoxViewer<br>"
		"Otherwise, that is, for data folder selected by user or created by Windows, DIRECTORY is constructed as follow :<br>"
		"    DIRECTORY = UserSelectedDataFolder\\MBoxViewer\\MappedMboxFilePath<br>"
		"Mapping example. Assuming<br>"
		"MboxFilePath = F:\\Account\\Inbox.mbox, then it will be mapped to<br>"
		"MappedMboxFilePath = F\\Account\\Inbox-mbox<br>"
		"and DIRECTORY to<br>"
		"DIRECTORY=UserSelectedDataFolder\\F\\Account\\Inbox-mbox<br>"
		"The ‘:’ character will be removed and “.mbox” extension, if present, will be mapped to \"-mbox\"<br>"
		"<br><br>"
		"DIRECTORY - <font color=red>target directory for mbox mail index files, help Html files and various sub-directories</font><br>"
		"DIRECTORY\\MailArchiveFile1.mbox<br>"
		"DIRECTORY\\MailArchiveFile1.mbox.mboxview<br>"
		"DIRECTORY\\MailArchiveFile2.mbox<br>"
		"DIRECTORY\\MailArchiveFile2.mbox.mboxview<br>"
		"DIRECTORY\\ImageCache<br>"
		"DIRECTORY\\ImageCache\\MailArchiveFile1 - <font color=red>target directory for image files, such as png,jpg,etc, embeded into mails</font><br>"
		"DIRECTORY\\ImageCache\\MailArchiveFile2<br>"
		"DIRECTORY\\AttachmentCache<br>"
		"DIRECTORY\\AttachmentCache\\MailArchiveFile1 - <font color=red>target directory for attachment files</font><br>"
		"DIRECTORY\\AttachmentCache\\MailArchiveFile2<br>"
		"DIRECTORY\\EmlCache<br>"
		"DIRECTORY\\EmlCache\\MailArchiveFile1 - <font color=red>target directory for Eml files</font><br>"
		"DIRECTORY\\EmlCache\\MailArchiveFile2<br>"
		"DIRECTORY\\ArchiveCache<br>"
		"DIRECTORY\\ArchiveCache\\MailArchiveFile1 - <font color=red>target directory for saving Found Mails and User Selected Mails to mbox and mboxlist files</font><br>"
		"DIRECTORY\\ArchiveCache\\MailArchiveFile2<br>"
		"DIRECTORY\\PrintCache<br>"
		"DIRECTORY\\PrintCache\\MailArchiveFile1  -  <font color=red>target directory for printing to single CSV,TEXT,HTML and PDF files</font><br>"
		"DIRECTORY\\PrintCache\\MailArchiveFile1\\PDF_GROUP  -  <font color=red>target directory for printing to separate PDF files</font><br>"
		"DIRECTORY\\PrintCache\\MailArchiveFile1\\PDF_GROUP\\PDF_MERGE  -  <font color=red>target directory for merged PDF files</font><br>"
		"DIRECTORY\\PrintCache\\MailArchiveFile1\\PDF_GROUP\\PDF_MERGE\\PDF_MERGE  -  <font color=red>target directory for merged PDF files if multiple merge steps are needed</font><br>"
		"DIRECTORY\\PrintCache\\MailArchiveFile1\\HTML_GROUP  -  <font color=red>target directory for printing to separate HTML files</font><br>"
		"DIRECTORY\\PrintCache\\MailArchiveFile1\\HTML_GROUP\\PDF_MERGE  -  <font color=red>target directory for merged PDF files</font><br>"
		"DIRECTORY\\PrintCache\\MailArchiveFile1\\HTML_GROUP\\PDF_MERGE\\PDF_MERGE  -  <font color=red>target directory for merged PDF files if multiple merge steps are needed</font><br>"
		"DIRECTORY\\PrintCache\\MailArchiveFile2<br>"
		"DIRECTORY\\LabelCache - <font color=red>target directory for Label files and sub-folders</font><br>"
		"DIRECTORY\\MergeCache - <font color=red>target directory for intermediate file created when “File-->Select root folder for merging” is selected</font><br>"
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
	CheckMessagewindowPositionMenuOption(1);
}


void CMainFrame::OnMessagewindowRight()
{
	// TODO: Add your command handler code here
	ConfigMessagewindowPosition(2);
	CheckMessagewindowPositionMenuOption(2);
}


void CMainFrame::OnMessagewindowLeft()
{
	// TODO: Add your command handler code here
	ConfigMessagewindowPosition(3);
	CheckMessagewindowPositionMenuOption(3);
}

void CMainFrame::CheckMessagewindowPositionMenuOption(int msgViewPosition)
{
	CMenu *menu = this->GetMenu();
	if (msgViewPosition == 1)
	{
		menu->CheckMenuItem(ID_MESSAGEWINDOW_BOTTOM, MF_CHECKED);
		menu->CheckMenuItem(ID_MESSAGEWINDOW_RIGHT, MF_UNCHECKED);
		menu->CheckMenuItem(ID_MESSAGEWINDOW_LEFT, MF_UNCHECKED);
	}
	else if (msgViewPosition == 2)
	{
		menu->CheckMenuItem(ID_MESSAGEWINDOW_BOTTOM, MF_UNCHECKED);
		menu->CheckMenuItem(ID_MESSAGEWINDOW_RIGHT, MF_CHECKED);
		menu->CheckMenuItem(ID_MESSAGEWINDOW_LEFT, MF_UNCHECKED);
	}
	else if (msgViewPosition == 3)
	{
		menu->CheckMenuItem(ID_MESSAGEWINDOW_BOTTOM, MF_UNCHECKED);
		menu->CheckMenuItem(ID_MESSAGEWINDOW_RIGHT, MF_UNCHECKED);
		menu->CheckMenuItem(ID_MESSAGEWINDOW_LEFT, MF_CHECKED);
	}
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


// TODO: Dead code; to be removed
int CMainFrame::MergeArchiveFiles()
{
	// TODO: customize CFileDialog to avoid potential buffer overflow and corruption
#define MAX_CFileDialog_FILE_COUNT 1000
	int FILE_LIST_BUFFER_SIZE = (MAX_CFileDialog_FILE_COUNT * (MAX_PATH + 1)) + 1;

	MboxMail::m_outbuf->ClearAndResize(FILE_LIST_BUFFER_SIZE);
	TCHAR *fileNameBuffer = MboxMail::m_outbuf->Data();

	//CString  sFilters  = "Mail Files (*.mbox)|*.mbox|All Files (*.*)|*.*||";
	CString  sFilters = "Mail Files All Files (*.*)|*.*|(*.mbox;*.eml)|*.mbox;*eml|(*.mbox)|*.mbox|(*.eml)|*.eml||";

	DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
	CFileDialog dlgFile(TRUE, NULL, NULL, dwFlags, sFilters);

	CString inFolderPath = MboxMail::GetLastPath();
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
		CString  fileNameFilter = "Mail Files (*.mbox)|*.mbox||";
		CString dfltExtension = ".mbox";

		CString outFolderPath;
		CString fileName;
		CString datapath = MboxMail::GetLastDataPath();
		if (inFolderPath.Compare(datapath))
		{
			CString appFolderPath = FileUtils::GetMboxviewLocalAppDataPath();
			inFolderPath = appFolderPath + "MboxRepo";
			FileUtils::CreateDirectory(inFolderPath);
		}
		BOOL ret = SaveFileDialog(fileName, fileNameFilter, dfltExtension, inFolderPath, outFolderPath, title);
		if (ret == TRUE)
		{
			CString filePath = outFolderPath + "\\" + fileName;
			CString fileExtension = ::PathFindExtension(fileName);
			CString fileName2 = ::PathFindFileName(filePath);
			//FileUtils::SplitFilePath(CString &fileName, CString &driveName, CString &directory, CString &fileNameBase, CString &fileNameExtention);

#if 1
			//CString dataFolder = outFolderPath + "\\" + CMainFrame::m_ViewerGlobalDB.GetDBFolder();
			CString dataFolder = outFolderPath;
			CString viewFile = dataFolder + "\\" + fileName + ".mboxview";
			if (FileUtils::PathFileExist(viewFile))
			{
				if (FileUtils::DeleteFile(viewFile, FALSE))
				{
					int deb = 1;
				}
				int deb = 1;
			}
			int deb = 1;
#endif

			CFile fp;
			CFileException ExError;
			if (!fp.Open(filePath, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &ExError))
			{
				CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

				CString txt = _T("Could not create \"") + filePath;
				txt += _T("\" file.\n");
				txt += exErrorStr;

				//TRACE(_T("%s\n"), txt);

				HWND h = NULL; // we don't have any window yet
				int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);

				MboxMail::m_outbuf->Clear();
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
					CFileException ExError;
					if (!fp_input.Open(strFilePath, CFile::modeRead, &ExError))
					{
						CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);  // TODO

						CString txt = _T("Could not open \"") + strFilePath;
						txt += _T("\" file.\nMake sure file is not open on other applications.");
						HWND h = NULL; // we don't have any window yet
						int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);

						fp.Close();

						MboxMail::m_outbuf->Clear();
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
			fp.Flush();
			fp.Close();

			/// TODO:  Encapsulate as function, refersh single file only instead of entire folder
			CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

			NTreeView *pTreeView = pFrame->GetTreeView();
			HTREEITEM hFolder = pTreeView->HasFolder(outFolderPath);
			HTREEITEM hItem = 0;
			if (hFolder)
				hItem = pTreeView->FindItem(hFolder, fileName);

			if (pFrame && pTreeView && hFolder && hItem)
			{
				MboxMail::SetLastPath(outFolderPath);
				pTreeView->DeleteMBoxAllWorkFolders(fileName);

				BOOL ret = pTreeView->RefreshFolder(hFolder);

				//hFolder = pTreeView->HasFolder(outFolderPath);
				//hItem = pTreeView->FindItem(hFolder, fileName);

				BOOL retval = pTreeView->SelectTreeItem(hItem);

				NListView *pListView = pFrame->GetListView();
				if (pListView)
				{
					if (MboxMail::s_mails.GetCount() > 0)
					{
						pListView->SelectItemFound(0);
					}
					pListView->SetListFocus();
				}
			}
			else
			{
				if (FileUtils::BrowseToFile(filePath) == FALSE) {
					HWND h = GetSafeHwnd();
					HINSTANCE result = ShellExecute(h, _T("open"), outFolderPath, NULL, NULL, SW_SHOWNORMAL);
					CheckShellExecuteResult(result, h);
				}
			}
		}
	}
	else
	{
		int deb = 1;
	}
	MboxMail::m_outbuf->Clear();
	return 1;
}

int CMainFrame::MergeMboxArchiveFiles(CArray<MergeFileInfo> &fileList, CString &mergedMboxFilePath)
{
	CString filePath;

	CFile fpMergeTo;
	CFileException exMergeTo;
	if (!fpMergeTo.Open(mergedMboxFilePath, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &exMergeTo))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exMergeTo);

		CString txt = _T("Could not create Merge To File \"") + mergedMboxFilePath;
		txt += _T("\" file.\n");
		txt += exErrorStr;

		//TRACE(_T("%s\n"), txt);

		CFileStatus rStatus;
		BOOL ret = fpMergeTo.GetStatus(rStatus);;

		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);

		return -1;
	}

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	CString sText;
	int paneId = 0;

	CString fullPath;
	CString fileName;

	for (int i = 0; i < fileList.GetCount(); i++)
	{
		MergeFileInfo &finfo = fileList.GetAt(i);
		fullPath = finfo.m_filepath;

		TRACE("FilePath=%s\n", filePath);

		_int64 fileSize = FileUtils::FileSize(fullPath);
		FileUtils::GetFileName(fullPath, fileName);
		sText.Format("Merging %s (%lld bytes) ...", fileName, fileSize);
		TRACE("%s\n", sText);
		if (pFrame)
			pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

		int ret = MergeMboxArchiveFile(fpMergeTo, fullPath, i == 0);

	}

	fpMergeTo.Close();

	_int64 fileSize = FileUtils::FileSize(mergedMboxFilePath);
	TRACE("Created merge file %s (%lld bytes) ...", mergedMboxFilePath, fileSize);

	return 1;
}

int CMainFrame::MergeMboxArchiveFiles(CString &mboxListFilePath, CString &mergedMboxFilePath)
{
	NTreeView *pTreeView = GetTreeView();
	NListView *pListView = GetListView();

	CString path;
	FileUtils::GetFolderPath(mergedMboxFilePath, path);

	MboxMail::SetLastPath(path);

	// Determine whether the folder path of mergedMboxFilePath is already open
	// or is a subfolder of any folder already open
	// This will help to determine whether we need to open the folder or refresh
	// standard foldr, subfolder or root folder and its subfolder
	FolderInfo *finfoBeg = 0;
	BOOL folderIsAlreadyOpen = FALSE;
	int iterCnt = 0;
	if (pTreeView)
	{
		CString folderPath;
		CString filePath = mergedMboxFilePath;

		while (!filePath.IsEmpty())
		{
			filePath.TrimRight("\\");

			FileUtils::GetFolderPath(filePath, folderPath);
			int len = folderPath.GetLength();
			if (len <= 0)
				break;

			if (folderPath.GetAt(len - 1) != '\\')
				break;

			finfoBeg = pTreeView->m_globalFolderInfoDB.Find(&folderPath);
			if (finfoBeg)
			{
				if (iterCnt == 0)
					folderIsAlreadyOpen = TRUE;
				break;
			}
			filePath = folderPath;
		}
		int deb = 1;
	}

	CStdioFile fpList;
	CFileException exList;
	if (!fpList.Open(mboxListFilePath, CFile::typeText | CFile::modeRead | CFile::shareDenyNone, &exList))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exList);

		CString txt = _T("Could not open list file \"") + mboxListFilePath;
		txt += _T("\" file.\n");
		txt += exErrorStr;

		TRACE(_T("MergeMboxArchiveFiles: %s\n"), txt);
		//errorText = txt;

		CFileStatus rStatus;
		BOOL ret = fpList.GetStatus(rStatus);

		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;
	CString folderPath;
	CString fileNameExtentionRoot;

	CString filePath;
	// TODO: To optimize, save all results in an Array for further processing
	CArray<MergeFileInfo> fileList;
	while (fpList.ReadString(filePath))
	{
		if (TextUtilsEx::isWhiteLine(filePath))
			continue;

		if (filePath.CompareNoCase(mergedMboxFilePath) == 0)
		{
			CString txt = _T("Invalid File Path:\n\n");
			txt.Append(filePath);
			txt.Append("\n\nin List File:\n\n");
			txt.Append(mboxListFilePath);
			txt.Append("\n\nMbox File Path and Merge To File Path can't be the same\nCan't continue merging files\n");

			int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);

			fpList.Close();
			return -1;
		}

		FileUtils::SplitFilePath(filePath, driveName, directory, fileNameBase, fileNameExtention);
		folderPath = driveName + directory;
		// check if wildcard name 
		// should we search for *.* rather than *; probably not if we want to examine file without extension
		if (fileNameBase.Find("*") >= 0)  
		{
			// examine all files in folderPath

			if (!FileUtils::PathDirExists(folderPath))  // should never fail =:)  ???
			{
				CString txt;
				txt.Format("Invalied Folder path \"%s\".\n\nCan't continue merging files\n", filePath);

				int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);

				fpList.Close();
				return -1;
			}

			CString searchFilter = filePath;
			WIN32_FIND_DATA	wf;
			HANDLE f = FindFirstFile(searchFilter, &wf);
			if (f != INVALID_HANDLE_VALUE)
			{
				do
				{
					if ((wf.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != FILE_ATTRIBUTE_DIRECTORY && wf.cFileName[0] != '.')
					{
						CString fn = wf.cFileName;
						CString fullPath = folderPath + fn;
						if (fullPath.CompareNoCase(mergedMboxFilePath) == 0)
						{
							int deb = 1;
							continue;  // don't merge target to merge file
						}
						int mboxFileType = NTreeView::ImboxviewFile(fullPath);
						if (mboxFileType < 1)
						{
							int deb = 1;
							continue;  // ignore non mbox type files
						}

						MergeFileInfo mergeFileInfo(fullPath, mboxFileType);
						fileList.Add(mergeFileInfo);
					}
				} while (FindNextFile(f, &wf));
				FindClose(f);
			}
		}
		else
		{
			TRACE("FilePath=%s\n", filePath);
			if (!FileUtils::PathFileExist(filePath))
			{
				CString txt = _T("Invalid File Path:\n\n");
				txt.Append(filePath);
				txt.Append("\n\nin List File:\n\n");
				txt.Append(mboxListFilePath);
				txt.Append("\n\nCan't continue merging files\n");

				int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);

				fpList.Close();
				return -1;
			}
			int mboxFileType = NTreeView::ImboxviewFile(filePath);
			if (mboxFileType < 1)
			{
				int deb = 1;
			}
			else
			{
				MergeFileInfo mergeFileInfo(filePath, mboxFileType);
				fileList.Add(mergeFileInfo);
			}
		}
	}
	fpList.Close();

	if (MergeMboxArchiveFiles(fileList, mergedMboxFilePath) < 0)
		return -1;

	CString mboxFolderPath;
	CString mboxFileName;
	FileUtils::GetFolderPathAndFileName(mergedMboxFilePath, mboxFolderPath, mboxFileName);
	//mboxFolderPath.TrimRight("\\");
	size_t hashsum = GlobalFolderInfoDB::GetHashsum(&mboxFolderPath);
	FolderInfo* finfo = pTreeView->m_globalFolderInfoDB.Find(&mboxFolderPath, hashsum);

	if (finfoBeg == 0)
	{
		DoOpen(mboxFolderPath);
	}
	else
	{
		if (pTreeView)
		{
			if (HTREEITEM hRootFolder = pTreeView->IsFolderOpen(finfoBeg->m_RootParent->m_folderPath))
			{
				// TODPO: optimize later if RootSubfolder
				pTreeView->RefreshFolder(hRootFolder);
			}
			else
				MboxMail::assert_unexpected();
		}
	}

	if (pTreeView)
	{

		MailSelectionInfo mailSelection;
		mailSelection.m_lastPath = mboxFolderPath;
		mailSelection.m_lastMailFilePath = mergedMboxFilePath;

		MboxMail::SetLastPath(mboxFolderPath);
		pTreeView->OpenLastSelection(&mailSelection);
	}

	return 1;
}


// Append single file content to the merge file
int CMainFrame::MergeMboxArchiveFile(CFile &fpMergeTo, CString &mboxFilePath, BOOL firstFile)
{
	// All archive files are assumed valid; merge
	// Caller should check the file content to validate. Not done yet

	CString filePath;
	CString fileName;
	FileUtils::GetFolderPathAndFileName(mboxFilePath, filePath, fileName);

	if (FileUtils::PathFileExist(mboxFilePath))
	{
		CFile fp_input;
		CFileException ExError;
		if (!fp_input.Open(mboxFilePath, CFile::modeRead, &ExError))
		{
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt = _T("Could not open \"") + mboxFilePath;
			txt += _T("\" mail file.\n");
			txt += exErrorStr;

			//TRACE(_T("%s\n"), txt);
			//errorText = txt;

			HWND h = NULL; // we don't have any window yet
			int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}

		// Append file
		UINT wantBytes = 16 * 1024;
		// don't use  MboxMail::m_oubuf is still used by dlgFile above !!!!! 
		MboxMail::m_inbuf->ClearAndResize(wantBytes + 2);
		TCHAR *inBuffer = MboxMail::m_inbuf->Data();

		UINT readBytes = 0;

		readBytes = fp_input.Read(inBuffer, wantBytes);
		if (readBytes <= 0)
		{
			fp_input.Close();
			return -1;
		}

		static const char *cFromMailBegin = "From ";
		static const int cFromMailBeginLen = strlen(cFromMailBegin);

		char *p = inBuffer;
		char *e = p + readBytes;
		char *fldName = p;
		char *pend = e;
		char *p_newline = e;

		p = MimeParser::SkipEmptyLines(p, e);

		char *ch_end_line = "\r\n";
		int ch_end_line_len = 2;

		if (TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) != 0)
		{
			// From 1513218656940664977@xxx Thu Sep 24 18:02:48 +0000 2015
			// More parsing would be needed to find first line with Date and Time
			// The below is safe for parsing the created mbox file
			CString FromLine = "From 1513218656940664977@xxx ";

			p_newline = MimeParser::EatNewLine(p, e);
			char ch_last = *(p_newline - 1);
			char ch_before_last = *(p_newline - 2);
			if (ch_last == '\n')
			{
				if (ch_before_last != '\r')
				{
					ch_end_line = "\n";
					ch_end_line_len = 1;
				}
			}

			fpMergeTo.Write(FromLine, FromLine.GetLength());
			fpMergeTo.Write(ch_end_line, ch_end_line_len);
		}
		inBuffer = p;
		readBytes = e - p;

		if (readBytes > 0)
		{
			fpMergeTo.Write(inBuffer, readBytes);
		}

		do
		{
			readBytes = fp_input.Read(inBuffer, wantBytes);
			if (readBytes > 0)
			{
				fpMergeTo.Write(inBuffer, readBytes);
			}
		} while (readBytes > 0);

		fpMergeTo.Write(ch_end_line, ch_end_line_len);

#if 0
		// not sure whether below is neeeded in case no CR NL or just NL at the file end
		// It would change total length on new archive which should not be an issue 
		// unless someone would try to add all lengths manually to verify
		MboxMail::m_inbuf->Clear();
		MboxMail::m_inbuf->Append("\r\n");
		fpMergeTo.Write(inBuffer, readBytes);
#endif

		fp_input.Close();
	}
	return 1;
}

BOOL CMainFrame::SaveFileDialog(CString &fileName, CString &fileNameFilter, CString &dfltExtention, CString &inFolderPath, CString &outFolderPath, CString &title)
{
	// TODO: customize CFileDialog to avoid potential buffer overflow and corruption
	int FILE_LIST_BUFFER_SIZE = (2 * (MAX_PATH + 1)) + 1;

	CString path = MboxMail::GetLastPath();
	CString datapath = MboxMail::GetLastDataPath();
	CString mergeCachePath = datapath + "MergeCache";
	BOOL ret = FileUtils::CreateDirectory(mergeCachePath);

	for (;;)
	{
		MboxMail::m_inbuf->ClearAndResize(FILE_LIST_BUFFER_SIZE);
		TCHAR *fileNameBuffer = MboxMail::m_inbuf->Data();

		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
		dwFlags &= ~OFN_NOCHANGEDIR;
		CFileDialog dlgFile(FALSE, dfltExtention, NULL, dwFlags, fileNameFilter);
		//MySaveFileDialog dlgFile(FALSE, dfltExtention, NULL, dwFlags, fileNameFilter);

		OPENFILENAME& ofn = dlgFile.GetOFN();
		//ofn.Flags |= ??;
		//ofn.lpstrFile = fileNameBuffer;
		//ofn.nMaxFile = FILE_LIST_BUFFER_SIZE;
		ofn.lpstrInitialDir = mergeCachePath;
		ofn.lpstrTitle = title;

		INT_PTR ret = dlgFile.DoModal();
		if (ret == IDOK)
		{
			fileName = dlgFile.GetFileName();
			CString fileFilePath = dlgFile.GetPathName();
			BOOL ret = FileUtils::CPathGetPath(fileFilePath, outFolderPath);

			TRACE("FOLDER=%s\n", outFolderPath);
			TRACE("FILE=%s\n", fileName);

			MboxMail::m_inbuf->Clear();

			if (outFolderPath.Compare(mergeCachePath))
			{
				CString txt = _T("Changing target folder for merged file is not allowed.\n\n");
				txt.Append(_T("Please try again."));
				HWND h = GetSafeHwnd(); // we don't have any window yet
				int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
				continue;
			}
			return TRUE;
		}
		else
		{
			MboxMail::m_inbuf->Clear();
			return FALSE;
		}
	}
	return FALSE;
}

// Added to NTreeView as Folder option; Noy used anymore, to be removed later
void CMainFrame::OnFileMergearchivefiles()
{
	// TODO: Add your command handler code here
	HWND h = GetSafeHwnd();
	MboxMail::ShowHint(HintConfig::MergeFilesHint, h);
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
			errorText = _T("Path to Chrome Browser not valid.\nPlease make sure Chrome is installed and path is correct."
				"\nOr configure Microsoft Edge Browser for printing."
				"\nSelect File->Print Config to update the setup.");
			return -1;
		}
	}
	else if (m_NamePatternParams.m_bScriptType == 1)
	{
		path = m_NamePatternParams.m_MSEdgeBrowserPath;
		if (!FileUtils::PathFileExist(path))
		{
			errorText = _T("Path to MSEdge Browser not valid.\nPlease make sure Microsoft Edge is installed and path is correct."
				"\nOr configure Google Chrome Browser for printing."
				"\nSelect File->Print Config to update the setup.");
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

	BOOL delStatus = FileUtils::DeleteFile(pdfFileName);
	if (delStatus == FALSE) {
		DWORD error = GetLastError();
		CString errorText = FileUtils::GetLastErrorAsString();
		TRACE("ExecCommand_WorkerThread: DeleteFile error=%s\n", errorText);
	}

	int bScriptType = 0;
	CString path = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe";
	CString args = "--headless --disable-gpu --print-to-pdf=\"" + pdfFileName + "\" \"" + htmFileName + "\"";

	if (pFrame->m_NamePatternParams.m_bScriptType == 0)
	{
		// args no change
		path = pFrame->m_NamePatternParams.m_ChromeBrowserPath;
		if (!FileUtils::PathFileExist(path)) {
			errorText = _T("Path to Chrome Browser not valid.\nPlease make sure Chrome is installed and path is correct."
				"\nOr configure Microsoft Edge Browser for printing."
				"\nSelect File->Print Config to update setup.");
			int deb = 1;
			return -1;
		}
	}
	else if (pFrame->m_NamePatternParams.m_bScriptType == 1)
	{
		if (pFrame->m_NamePatternParams.m_bHeaderAndFooter)
			args = "--headless --disable-gpu";
		else
			args = "--headless --disable-gpu --print-to-pdf-no-header";

		args += " --print-to-pdf=\"" + pdfFileName + "\" \"" + htmFileName + "\"";
		path = pFrame->m_NamePatternParams.m_MSEdgeBrowserPath;
		if (!FileUtils::PathFileExist(path)) {
			errorText = _T("Path to MSEdge Browser not valid.\nPlease make sure MSEdge is installed and path is correct."
				"\nOr configure Google  Chrome Browser for printing."
				"\nSelect File->Print Config to update the setup.");
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

	CString inFolderPath = MboxMail::GetLastPath();

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

	NTreeView * pTreeView = GetTreeView();
	NListView * pListView = GetListView();
	NMsgView * pMsgView = GetMsgView();

	CString lastMailFilePath;
	CString lastLabelFilePath;
	int lastMailIndex = -1;
	int lastWhichSort = 1;
	int selMailIndex = -1;

	if (pListView)
	{
		lastMailFilePath = pListView->m_path;
		lastLabelFilePath = pListView->m_path_label;
		selMailIndex = pListView->m_lastSel;
		if ((selMailIndex >= 0) && (selMailIndex < MboxMail::s_mails.GetCount()))
		{
			lastMailIndex = MboxMail::s_mails[selMailIndex]->m_index;
		}

		lastWhichSort = MboxMail::b_mails_which_sorted;
	}

	BOOL ret;
	CString m_section = CString(sz_Software_mboxview);

	ret = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section, _T("lastMailFilePath"), lastMailFilePath);
	ret = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section, _T("lastLabelFilePath"), lastLabelFilePath);
	ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section, _T("lastMailIndex"), lastMailIndex);
	ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section, _T("lastWhichSort"), lastWhichSort);

	BOOL isDirty = MboxMail::m_editMails.m_bIsDirty;
	if ((MboxMail::IsUserMailsSelected() && (MboxMail::s_mails.GetCount() > 0)) || (MboxMail::s_mails_edit.GetCount() > 0))
	{
		if (isDirty)
		{
			CString txt = _T("User Selected Mails List not empty. Terminate program?");
			int answer = MessageBox(txt, _T("Warning"), MB_APPLMODAL | MB_ICONWARNING | MB_YESNO);
			if (answer == IDNO)
				return;
		}
	}
	else if (isDirty)
		int deb = 1;


	//if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode && !m_commandLineParms.m_bEmlPreviewFolderExisted)
	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
	{
		NTreeView *pTreeView = GetTreeView();
		if (pTreeView)
		{
			CString mboxFolderPath;
			CString mboxFileName;
			FileUtils::GetFolderPathAndFileName(CMainFrame::m_commandLineParms.m_mboxFileNameOrPath, mboxFolderPath, mboxFileName);

			HTREEITEM hFolder = pTreeView->HasFolder(mboxFolderPath);
			if (hFolder)
			{
				if (!m_commandLineParms.m_bEmlPreviewFolderExisted)
				{
					pTreeView->DeleteFolder(hFolder);
				}
#if 0
				else
				{
					BOOL recursive = FALSE;
					int childrenCount = 0;
					childrenCount = pTreeView->GetChildrenCount(hFolder, recursive);
					if (childrenCount == 1)
					{
						HTREEITEM hItem = pTreeView->FindItem(hFolder, mboxFileName);
						if (hItem)
						{
							pTreeView->DeleteFolder(hFolder);
						}
					}
				}
#endif
				int deb = 1;
			}
		}
	}

	m_section = CString(sz_Software_mboxview) + "\\" + "Window Placement";
	WINDOWPLACEMENT wp;
	GetWindowPlacement(&wp);

	if (wp.showCmd == SW_MAXIMIZE)
		wp.showCmd = SW_SHOWMAXIMIZED; 
	if ((wp.showCmd != SW_SHOWMAXIMIZED) && (wp.showCmd != SW_SHOW))
		wp.showCmd = SW_SHOW;

	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
	{
		BOOL ret = CProfile::_WriteProfileBinary(HKEY_CURRENT_USER, m_section, "MainFrame_EmlPreviewMode", (LPBYTE)&wp, sizeof(wp));
		if (m_msgViewPosition != m_newMsgViewPosition)
			DeleteAllPlacementKeys();		
	}
	else
	{
		BOOL ret = CProfile::_WriteProfileBinary(HKEY_CURRENT_USER, m_section, "MainFrame", (LPBYTE)&wp, sizeof(wp));

		if (m_msgViewPosition == m_newMsgViewPosition)
		{
			BOOL ret;
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
	}

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
	NTreeView *pTreeView = GetTreeView();
	NListView *pListView = GetListView();

	if (pTreeView)
	{
		//pTreeView->SimplyTest();  // just a test

			// Move legacy folders created by older versions under a single folder
		BOOL anwer = pTreeView->MoveMBoxViewerCreatedLegacyFolders();
		if (anwer == FALSE)
			return 0;
	}

	CString saveDataFolder = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("dataFolder"));

	int selectedDataFolderConfigMethod = -1;
	int retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("dataFolderConfigMethod"), selectedDataFolderConfigMethod);
	//retval = FALSE;
	if (retval == FALSE)
	{
		if (!saveDataFolder.IsEmpty())
			MboxMail::assert_unexpected();

		DataFolderConfigDlg dlg;

		INT_PTR retCode = dlg.DoModal();
		if (retCode == IDOK)
		{
			int deb = 1;
		}
		else if (retCode == IDCANCEL)
		{
			int deb = 1;
		}
	}

	CString dataFolder = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("dataFolder"));
	MboxMail::s_folderContext.m_rootDataFolderPathConfig = dataFolder;

	if (pTreeView)
	{
		if (!m_commandLineParms.m_bEmlPreviewMode)
			pTreeView->LoadFolders();
	}

	CString listFilePath = CMainFrame::m_commandLineParms.m_mboxListFilePath;
	CString mergeToFilePath = CMainFrame::m_commandLineParms.m_mergeToFilePath;

	if (!listFilePath.IsEmpty() && !mergeToFilePath.IsEmpty())
	{
		CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
		CString sText;
		int paneId = 0;

		sText.Format("Merging files listed in %s file ...", listFilePath);
		TRACE("%s\n", sText);
		if (pFrame)
			pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

		int ret = MergeMboxArchiveFiles(listFilePath, mergeToFilePath);
		if (ret < 0)
			int deb = 1;

		sText.Format("Ready");
		TRACE("%s\n", sText);
		if (pFrame)
			pFrame->SetStatusBarPaneText(paneId, sText, TRUE);
	}
	else if (m_commandLineParms.m_bEmlPreviewMode && !m_commandLineParms.m_mboxFileNameOrPath.IsEmpty())
	{
		pTreeView->DoOpenFilePreview(m_commandLineParms.m_mboxFileNameOrPath);
	}
	else if (!m_commandLineParms.m_mboxFileNameOrPath.IsEmpty())
	{
		CString mboxFilePath;
		CString mboxFileName;
		FileUtils::GetFolderPathAndFileName(m_commandLineParms.m_mboxFileNameOrPath, mboxFilePath, mboxFileName);

		mboxFilePath.TrimRight("\\");
		mboxFilePath.Append("\\");

		CString mboxFileNameOrPath = m_commandLineParms.m_mboxFileNameOrPath;
		mboxFileNameOrPath.TrimRight("\\");
		DoOpen(mboxFilePath);

		// Don't delete any of ccreated files or folder i OnClose()
		// This option is like normal Open
		//if (pTreeView->m_folderArray.Find(mboxFilePath) >= 0)
			//m_commandLineParms.m_bEmlPreviewFolderExisted = TRUE;

		if (pTreeView->m_folderArray.Find(mboxFilePath) >= 0)
			HTREEITEM hItem = pTreeView->SelectFolderItem(mboxFileNameOrPath);
	}
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

	NListView *pListView = 0;
	pListView = this->GetListView();

	if (pListView)
	{
		pListView->SortByColumn(1, TRUE);
	}
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


void CMainFrame::OnViewMessageheaders()
{
	// TODO: Add your command handler code here

	MboxMail::ShowHint(HintConfig::MessageHeaderConfigHint, GetSafeHwnd());

	m_bViewMessageHeaders = !m_bViewMessageHeaders;

	NListView *pListView = GetListView();
	NMsgView * pMsgView = GetMsgView();

	if (pListView && pMsgView)
	{
		int iItem = pListView->m_lastSel;
		if (m_bViewMessageHeaders)
		{
			pMsgView->ShowMailHeader(iItem);
		}
		else
		{
			pListView->Invalidate();
			pListView->SelectItem(pListView->m_lastSel, TRUE);
		}
	}

	CMenu *menu = this->GetMenu();
	if (m_bViewMessageHeaders)
		menu->CheckMenuItem(ID_VIEW_MESSAGEHEADERS, MF_CHECKED);
	else
		menu->CheckMenuItem(ID_VIEW_MESSAGEHEADERS, MF_UNCHECKED);

	int deb = 1;
}


void CMainFrame::OnFileRestorehintmessages()
{
	// TODO: Add your command handler code here

	MboxMail::m_HintConfig.ClearAllHints();
}


void CMainFrame::OnMessageheaderpanelayoutDefault()
{
	// TODO: Add your command handler code here

	NMsgView * pMsgView = GetMsgView();
	if (pMsgView)
	{
		pMsgView->OnMessageheaderpanelayoutDefault();

		CMenu *menu = this->GetMenu();
		menu->CheckMenuItem(ID_MESSAGEHEADERPANELAYOUT_DEFAULT, MF_CHECKED);
		menu->CheckMenuItem(ID_MESSAGEHEADERPANELAYOUT_EXPANDED, MF_UNCHECKED);
	}
}

void CMainFrame::OnMessageheaderpanelayoutExpanded()
{
	// TODO: Add your command handler code here

	NMsgView * pMsgView = GetMsgView();
	if (pMsgView)
	{
		pMsgView->OnMessageheaderpanelayoutExpanded();

		CMenu *menu = this->GetMenu();
		menu->CheckMenuItem(ID_MESSAGEHEADERPANELAYOUT_DEFAULT, MF_UNCHECKED);
		menu->CheckMenuItem(ID_MESSAGEHEADERPANELAYOUT_EXPANDED, MF_CHECKED);
	}
}

int CommandLineParms::VerifyParameters()
{
	BOOL bError = FALSE;

	if ((!m_mboxListFilePath.IsEmpty() || !m_mergeToFilePath.IsEmpty()) &&
		(!m_mboxFileNameOrPath.IsEmpty() || !m_mboxFolderPath.IsEmpty() || (m_bEmlPreviewMode == TRUE)))
	{
		bError = TRUE;
	}
	if ((bError == FALSE) && (!m_mboxFolderPath.IsEmpty() &&
		(!m_mboxListFilePath.IsEmpty() || !m_mergeToFilePath.IsEmpty() ||
			!m_mboxFileNameOrPath.IsEmpty() || (m_bEmlPreviewMode == TRUE))))
	{
		bError = TRUE;
	}
	if ((bError == FALSE) && !m_mboxFileNameOrPath.IsEmpty() &&
		(!m_mboxListFilePath.IsEmpty() || !m_mergeToFilePath.IsEmpty() || !m_mboxFolderPath.IsEmpty()))
	{
		bError = TRUE;
	}
	if ((bError == FALSE) && (m_bEmlPreviewMode == TRUE) && m_mboxFileNameOrPath.IsEmpty())
	{
		bError = TRUE;
	}
	if (bError)
	{
		CString txt = _T("Invalid Command Line \n\n\"");
		CString opt = m_allCommanLineOptions;
		txt += opt
			+ _T("\"")
			+ _T("\n\nValid option combinations are:\n\n")
			+ _T("  -FOLDER=Folder Path to open\n")
			+ _T("  \nor\n\n")
			+ _T("  -MAIL_FILE=Mbox File Path to open\n")
			+ _T("  \nor\n\n")
			+ _T("  -MAIL_FILE=Mbox File Path to open\n")
			+ _T("  -EML_PREVIEW_MODE\n")
			+ _T("  \nor\n\n")
			+ _T("  -MBOX_MERGE_LIST_FILE=Path to File containing list of mbox files to merge\n")
			+ _T("  -MBOX_MERGE_TO_FILE=Path to File to save merge results\n")
			//+ _T("\nDo you want to continue?")
			;

		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
		return -1;
#if 0
		if (answer == IDNO)
			return -1;
		else
			return -2;
#endif
	}

	if (!m_mboxFileNameOrPath.IsEmpty() && !m_bEmlPreviewMode)
	{
		if (!FileUtils::PathFileExist(m_mboxFileNameOrPath))
		{
			CString txt;
			txt.Format(_T("Invalid -MAIL_FILE=\"%s\" option.\n"
				"No such File or Directory\"%s\"\n"),
				m_mboxFileNameOrPath, m_mboxFileNameOrPath);
			txt += _T("Do you want to continue?");
			HWND h = NULL; // we don't have any window yet  
			int answer = MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
			return -1;
#if 0
			if (answer == IDNO)
				return -1;
			else
				return -2;
#endif
		}

		CString mboxFilePath;
		CString mboxFileName;
		FileUtils::GetFolderPathAndFileName(m_mboxFileNameOrPath, mboxFilePath, mboxFileName);

		CString folderPath = mboxFilePath;

		CString driveName;
		CString directory;
		CString fileNameBase;
		CString fileNameExtention;

		folderPath.TrimRight("\\");
		folderPath.Append("\\");

		FileUtils::SplitFilePath(folderPath, driveName, directory, fileNameBase, fileNameExtention);
		if (directory.GetLength() <= 1)
		{
			CString txt;
			txt.Format(_T("Invalid -MAIL_FILE=\"%s\" option.\nInvalid folder \"%s\" .\n"
				"The mbox files must be installed under a named folder.\n"
				"Please create folder, move the mbox files to that folder and try again."),
				m_mboxFileNameOrPath, mboxFilePath);
			HWND h = NULL; // we don't have any window yet  
			int answer = MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONWARNING | MB_OK);
			return -1;
		}


		// LoadFolders() needs "lastPath" to open folder housing "mailFile"
		MboxMail::SetLastPath(folderPath);
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("mailFile"), m_mboxFileNameOrPath);
	}

	if (!m_mboxFolderPath.IsEmpty())
	{
		CString driveName;
		CString directory;
		CString fileNameBase;
		CString fileNameExtention;

		CString folderPath = m_mboxFolderPath;
		folderPath.TrimRight("\\");
		folderPath.Append("\\");

		FileUtils::SplitFilePath(folderPath, driveName, directory, fileNameBase, fileNameExtention);
		if (directory.GetLength() <= 1)
		{
			CString txt;
			folderPath.TrimRight("\\");

			txt.Format(_T("Invalid -FOLDER=\"%s\" option.\n\nInvalid folder \"%s\" .\n"
				"The mbox files must be installed under a named folder\n."
				"Please create folder, move the mbox files to that folder and try again."),
				folderPath, folderPath);
			HWND h = NULL; // we don't have any window yet  
			int answer = MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
			return -1;
		}

		BOOL ret = FileUtils::PathDirExists(folderPath);
		if (ret == FALSE)
		{
			CString txt;
			txt.Format(_T("Invalid -FOLDER=\"%s\" option.\n\n"
				"No such Folder \"%s\" .\n"),
				folderPath, folderPath);
			//txt += _T("Do you want to continue?");
			HWND h = NULL; // we don't have any window yet  
			int answer = MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
			return -1;
#if 0
			if (answer == IDNO)
				return -1;
			else
				return -2;
#endif
		}
		MboxMail::SetLastPath(folderPath);
	}

	if (!m_mboxListFilePath.IsEmpty() && !m_mergeToFilePath.IsEmpty())
	{
		if (!FileUtils::PathFileExist(m_mboxListFilePath))
		{
			CString txt;
			txt.Format(_T("Invalid -MBOX_MERGE_LIST_FILE=\"%s\" option.\n"
				"No such File:  \"%s\"\n\n"),
				m_mboxListFilePath, m_mboxListFilePath);
			//txt += _T("Do you want to continue?");
			HWND h = NULL; // we don't have any window yet  
			int answer = MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
			return -1;
#if 0
			if (answer == IDNO)
				return -1;
			else
				return -2;
#endif
		}

		CString mboxFolderPath;
		CString mboxFileName;
		FileUtils::GetFolderPathAndFileName(m_mergeToFilePath, mboxFolderPath, mboxFileName);

		CString driveName;
		CString directory;
		CString fileNameBase;
		CString fileNameExtention;

		FileUtils::SplitFilePath(m_mergeToFilePath, driveName, directory, fileNameBase, fileNameExtention);
		if (fileNameExtention.CompareNoCase(_T(".mbox")))
		{
			CString txt;
				txt.Format(_T("Invalid -MBOX_MERGE_TO_FILE=\"%s\" option.\nInvalid file extension \"%s\" .\n"
					"File extension must be set to .mbox.\n"),
					m_mergeToFilePath, fileNameExtention);
				HWND h = NULL; // we don't have any window yet  
				int answer = MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONWARNING | MB_OK);
				return -1;
		}

		if (directory.GetLength() <= 1)
		{
			CString txt;
			txt.Format(_T("Invalid -MBOX_MERGE_TO_FILE=\"%s\" option.\nInvalid folder \"%s\" .\n"
				"The mbox files must be installed under a named folder.\n"
				"Please create folder, move the mbox files to that folder and try again."),
				m_mergeToFilePath, mboxFolderPath);
			HWND h = NULL; // we don't have any window yet  
			int answer = MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONWARNING | MB_OK);
			return -1;
		}

		if (!CMainFrame::CanMboxBeSavedInFolder(mboxFolderPath))
		{
#if 0
			CString txt = _T("The \"MBoxViewer\" name is reserved and it can't appear in the file path.\n"
				"Please create different name and try again.");
			HWND h = NULL;
			int answer = MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
#endif
			return -1;
		}

		BOOL ret = FileUtils::PathDirExists(mboxFolderPath);
		if (ret == FALSE)
		{
			CString txt;
			txt.Format(_T("Invalid -MBOX_MERGE_TO_FILE=\"%s\" option.\n\nNo such folder: \"%s\".\n\nCommand line:\n\n\"%s\"\n\n"),
				m_mergeToFilePath, mboxFolderPath, m_allCommanLineOptions);
			//txt += _T("Do you want to continue?");
			HWND h = NULL; // we don't have any window yet  
			int answer = MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
			return -1;
#if 0
			if (answer == IDNO)
				return -1;
			else
				return -2;
#endif
		}

		if (FileUtils::PathFileExist(m_mergeToFilePath))
		{
			CString txt;
			txt.Format(_T("Possibly Invalid -MBOX_MERGE_TO_FILE=\"%s\" option.\n\n"
				"File Exists\"%s\"\n\n"
			"!!!  Make sure you are not overwritting original mbox file from mail service such as Gmail !!!\n\n"
			"It is always good idea to have backup of original mbox mail files\n\n"),
				m_mergeToFilePath, m_mergeToFilePath);
			txt += _T("Override and continue?");
			HWND h = NULL; // we don't have any window yet  
			int answer = MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
			if (answer == IDNO)
				return -1;
			else
			{
				//FileUtils::DeleteFile(m_mergeToFilePath);
				return 1;
			}
		}
	}
	else if (!m_mboxListFilePath.IsEmpty() && m_mergeToFilePath.IsEmpty())
	{
		CString txt;
		txt.Format(_T("Missing or Invalid -MBOX_MERGE_TO_FILE option.\nCommand line:\n\n\"%s\"\n\n"),
			m_allCommanLineOptions);
		//txt += _T("\nDo you want to continue?");
		HWND h = NULL; // we don't have any window yet  
		int answer = MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
		return -1;
#if 0
		if (answer == IDNO)
			return -1;
		else
			return -2;
#endif
	}
	else if (m_mboxListFilePath.IsEmpty() && !m_mergeToFilePath.IsEmpty())
	{
		CString txt;
		txt.Format(_T("Missing or Invalid -MBOX_MERGE_LIST_FILE option.\nCommand line:\n\n\"%s\"\n\n"),
			m_allCommanLineOptions);
		//txt += _T("\nDo you want to continue?");
		HWND h = NULL; // we don't have any window yet  
		int answer = MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
		return -1;
#if 0
		if (answer == IDNO)
			return -1;
		else
			return -2;
#endif
	}
	return(1);
}


// TODO: 
BOOL CMainFrame::CanMboxBeSavedInFolder(CString &destinationFolder)
{
	// Imposed more restrictions to simplify
	if ((destinationFolder.Find("\\MBoxViewer\\") >= 0) || (destinationFolder.Find("\\MBoxViewer") >= 0))
	{
		CString txt;
		txt.Format(_T("Invalid -MBOX_MERGE_TO_FILE=\"%s\" option.\n\nInvalid folder \"%s\" .\n"),
			destinationFolder, destinationFolder);
		txt.Append(_T("\nDestination folder can't contain \"MBoxViewer\" folder  in the path \n\n"));
			HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return FALSE;
	}
	else
		return TRUE;
}

void CMainFrame::OnFileSmtpmailserverconfig()
{
	// TODO: Add your command handler code here

	SMTPMailServerConfigDlg dlg;

	dlg.m_mailDB.Copy(m_mailDB);

	if (dlg.DoModal() == IDOK)
	{
		int deb = 1;
	}
	else
	{
		int deb = 1;
	}
	m_mailDB.Copy(dlg.m_mailDB);

	CString appDataPath = FileUtils::GetMboxviewLocalAppDataPath("MailService");
	CString smtpConfigFilePath = appDataPath + "SMTP.ini";

	CreateMailDbFile(m_mailDB, smtpConfigFilePath);

	int deb = 1;
}


void CMainFrame::OnHelpUserguide()
{
	// TODO: Add your command handler code here
	CString processPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("processPath"));

	CString processDir;
	FileUtils::CPathGetPath(processPath, processDir);
	CString filePath = processDir + "\\UserGuide.pdf";

	if (FileUtils::PathFileExist(filePath))
	{
		ShellExecute(NULL, _T("open"), filePath, NULL, NULL, SW_SHOWNORMAL);
	}
	else
	{
		CString txt;
		txt.Format("User Guide file \"%s\" doesn't exist", filePath);
		HWND h = GetSafeHwnd();
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
	}
}


void CMainFrame::OnHelpReadme()
{
	// TODO: Add your command handler code here
	CString processPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("processPath"));

	CString processDir;
	FileUtils::CPathGetPath(processPath, processDir);
	CString filePath = processDir + "\\README.txt";

	if (FileUtils::PathFileExist(filePath))
	{
		ShellExecute(NULL, _T("open"), filePath, NULL, NULL, SW_SHOWNORMAL);
	}
	else
	{
		CString txt;
		txt.Format("Change History file \"%s\" doesn't exist", filePath);
		HWND h = GetSafeHwnd();
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
	}
}


void CMainFrame::OnHelpLicense()
{
	// TODO: Add your command handler code here
	CString processPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("processPath"));

	CString processDir;
	FileUtils::CPathGetPath(processPath, processDir);
	CString filePath = processDir + "\\LICENSE.txt";

	if (FileUtils::PathFileExist(filePath))
	{
		ShellExecute(NULL, _T("open"), filePath, NULL, NULL, SW_SHOWNORMAL);
	}
	else
	{
		CString txt;
		txt.Format("License file \"%s\" doesn't exist", filePath);
		HWND h = GetSafeHwnd();
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
	}
}

void CMainFrame::OpenHelpFile(CString &helpFileName, HWND h)
{
	CString processPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("processPath"));

	CString processDir;
	FileUtils::CPathGetPath(processPath, processDir);
	CString filePath = processDir + "\\HelpFiles\\" + helpFileName;

	if (FileUtils::PathFileExist(filePath))
	{
		ShellExecute(NULL, _T("open"), filePath, NULL, NULL, SW_SHOWNORMAL);
	}
	else
	{
		CString txt;
		txt.Format("Help file \"%s\" doesn't exist", filePath);
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
	}
	int deb = 1;
}

BOOL CMainFrame::CreateMailDbFile(MailDB &m_mailDB, CString &fileName)
{
	CFile fp;
	CFileException ExError;
	if (!fp.Open(fileName, CFile::modeWrite | CFile::modeCreate, &ExError))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = _T("Could not create \"") + fileName;
		txt += _T("\" file.\n");
		txt += exErrorStr;

		TRACE(_T("%s\n"), txt);

		return FALSE;
	}

	CString section = "[MailService]\r\n";
	fp.Write(section, section.GetLength());
	CString fld = "ActiveMailService=" + m_mailDB.ActiveMailService + "\r\n";
	fp.Write(fld, fld.GetLength());

	WriteMTPServerConfig(m_mailDB.GmailSMTPConfig, fp);
	WriteMTPServerConfig(m_mailDB.YahooSMTPConfig, fp);
	WriteMTPServerConfig(m_mailDB.OutlookSMTPConfig, fp);
	WriteMTPServerConfig(m_mailDB.CustomSMTPConfig, fp);

	return TRUE;
}

BOOL CMainFrame::WriteMTPServerConfig(MailConfig &serverConfig, CFile &fp)
{
	CString CR = "\r\n";
	CString section = "[" + serverConfig.MailServiceName + "]" + CR;
	fp.Write(section, section.GetLength());
	CString fld = "MailServiceName=" + serverConfig.MailServiceName + CR;
	fp.Write(fld, fld.GetLength());
	fld = "SmtpServerAddress=" + serverConfig.SmtpServerAddress + CR;
	fp.Write(fld, fld.GetLength());
	fld.Format("SmtpServerPort=%d\r\n", serverConfig.SmtpServerPort);
	fp.Write(fld, fld.GetLength());
	fld = "UserAccount=" + serverConfig.UserAccount + CR;
	fp.Write(fld, fld.GetLength());
	//fld = "UserPassword=" + serverConfig.UserPassword + CR;
	fld = "UserPassword=" + CR;
	fp.Write(fld, fld.GetLength());
	fld.Format("EncryptionType=%d\r\n", serverConfig.EncryptionType);
	fp.Write(fld, fld.GetLength());

	return TRUE;
}

#if 0
void CMainFrame::OnKeydown(NMHDR* pNMHDR, LRESULT* pResult)
{
	NMLVKEYDOWN* pLVKeyDow = (NMLVKEYDOWN*)pNMHDR;
	*pResult = 0;

	switch (pLVKeyDow->wVKey) {
	case VK_ESCAPE:
	{
		if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
		{
			AfxGetMainWnd()->PostMessage(WM_CLOSE);
		}
	}
	}
}

BOOL CMainFrame::PreTranslateMessage(MSG* pMsg)
{
	if ((pMsg->message & 0xffff) == WM_MOUSEMOVE)
	{
		SetFocus();
		int deb = 1;
	}
	else if ((pMsg->message & 0xffff) == WM_KEYDOWN)
	{
		if (pMsg->wParam == VK_ESCAPE)
		{
			if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
			{
				AfxGetMainWnd()->PostMessage(WM_CLOSE);
			}
		}
	}

	return CWnd::PreTranslateMessage(pMsg);
}
#endif

int CMainFrame::FileSelectrootfolder(int treeType)
{
	// TODO: Add your command handler code here
	CString fileName;
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	CString path;

	if (m_bEnhancedSelectFolderDlg == FALSE)
	{
		path = MboxMail::GetLastPath();
		CBrowseForFolder bff(GetSafeHwnd(), CSIDL_DESKTOP, IDS_SELECT_FOLDER);
		if (!path.IsEmpty())
			bff.SetDefaultFolder(path);

		bff.SetFlags(BIF_RETURNONLYFSDIRS);
		if (bff.SelectFolder())
		{
			path = bff.GetSelectedFolder();
			path.TrimRight(_T("\\"));
		}
		else
		{
			int deb = 1;
			return -1;
		}
	}
	else
	{
		INT_PTR ret = SelectFolder(path);
		if (ret == IDOK)
		{
			path.TrimRight(_T("\\"));
		}
		else
		{
			int deb = 1;
			return -1;
		}
	}

	if (!FileUtils::PathDirExistsW(path))
	{
		CString errorText = FileUtils::GetLastErrorAsString();

		CString txt;
		
		txt.Format(_T("Selected Folder\n\n%s\n\ndoesn't exist. Error: \n\n%s\n\n"
			"Only ANSI folder names are curently supported. Please verify and rename the folder if that is the case."),
			path, errorText);

		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	if (treeType == 2)
	{
		if ((m_mergeRootFolderStyle == 1) && (m_labelAssignmentStyle == 1))  // labels per mbox file
		{
			BOOL recursive = TRUE;

			int fileCnt = FileUtils::GetFolderFileCount(path, recursive);
			if (fileCnt > 1000)
			{
				CString txt;
				txt.Format(_T("Found [%d] files under the root folder and sub-folders.\n\n"
					"Content of files was not examined to determine if all files are of the mbox type.\n\n"
					"In case all/most of files found are of the mbox type,"
					" this will create very large number of labels under Mail Tree which could make viewing of mails very challenging.\n\n"
					"Do you want to continue, cancel or retry Merge Configuration?"
				), fileCnt);

				int answer = MessageBox(txt, _T("Warning"), MB_APPLMODAL | MB_ICONQUESTION | MB_CANCELTRYCONTINUE);
				if (answer == IDCANCEL)
					return -1;
				else if (answer == IDTRYAGAIN)
					return -2;
				// else IDCONTINUE
			}
		}
	}
	else if (treeType == 1)
	{
		BOOL recursive = TRUE;

		int fileCnt = FileUtils::GetFolderFileCount(path, recursive);
		if (fileCnt > 1000)
		{
			CString txt;
			txt.Format(_T("Found [%d] files under the root folder and sub-folders.\n\n"
				"Content of files was not examined to determine if all files are of the mbox type.\n\n"
				"In case all/most of files found are of the mbox type,"
				" this will create very large number of items under Mail Tree which could make viewing of mails very challenging.\n\n"
				"Do you want to continue, cancel or retry Select root folder dialog?"
			), fileCnt);

			int answer = MessageBox(txt, _T("Warning"), MB_APPLMODAL | MB_ICONQUESTION | MB_CANCELTRYCONTINUE);
			if (answer == IDCANCEL)
				return -1;
			else if (answer == IDTRYAGAIN)
				return -2;
			// else IDCONTINUE
		}
	}

	path.Append(_T("\\"));

	FileUtils::SplitFilePath(path, driveName, directory, fileNameBase, fileNameExtention);
	if (directory.GetLength() <= 1)
	{
		CString txt = _T("The mbox files must be installed under a named folder\n."
			"Please create folder, move the mbox files to that folder and try again.");
		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
		return -1;
	}

	path.TrimRight(_T("\\"));
	//AfxGetApp()->AddToRecentFileList(path);

	path.Append(_T("\\"));
	MboxMail::SetLastPath(path);

	if (treeType == 1)
		OpenRootFolderAndSubfolders(path);
	else if (treeType == 2)
		OpenRootFolderAndSubfolders_LabelView(path);
	else
		return -1;

	return 1;
}


void CMainFrame::OnFileDatafolderconfig()
{
	// TODO: Add your command handler code here

	CString saveDataFolder = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("dataFolder"));

	int selectedDataFolderConfigMethod = -1;
	int retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("folderConfigMethod"), selectedDataFolderConfigMethod);
	retval = FALSE;
	if (retval == FALSE)
	{
		if (!saveDataFolder.IsEmpty())
			MboxMail::assert_unexpected();

		DataFolderConfigDlg dlg(TRUE);

		INT_PTR retCode = dlg.DoModal();
		if (retCode == IDOK)
		{
			int deb = 1;
		}
		else if (retCode == IDCANCEL)
		{
			int deb = 1;
		}
	}

	CString dataFolder = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("dataFolder"));
	MboxMail::s_folderContext.m_rootDataFolderPathConfig = dataFolder;

	if (saveDataFolder.Compare(dataFolder))
	{
		NTreeView *pTreeView = GetTreeView();

		if (pTreeView)
		{
			if (dataFolder.IsEmpty())
			{
				pTreeView->DeleteOldDataFolder(saveDataFolder, dataFolder);
			}
			else
				pTreeView->DeleteOldDataFolder(saveDataFolder, dataFolder);
			int deb = 1;
		}
	}

	MboxMail::s_folderContext.m_rootDataFolderPathConfig = dataFolder;
}

void CMainFrame::OpenRootFolderAndSubfolders_LabelView(CString &path)
{
	NTreeView *pTreeView = GetTreeView();

	if (pTreeView)
	{
		BOOL selectFolder = TRUE;
		pTreeView->DoOpenRootFolderAndSubfolders_LabelView(path, selectFolder);
	}
}

void CMainFrame::OnFileMergerootfoldersub()
{
	// TODO: Add your command handler code here
#if 0
	CString txt = _T("MBox Viewer will traverse selected root folder and all sub-folders and merge all mbox files found into a single mbox file.\n\n"
		"MBox Viewer will assign a label to each folder and each mbox file and create Tree of Labels View\n\n"
		"Do you want to continue?\n");
	int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_YESNO);
	if (answer == IDYES)
	{
		FileSelectrootfolder(2);
	}
#else
	int wantsToRetry = -2;
	MergeRootFolderAndSubfolders dlg;

	while (wantsToRetry == -2)
	{
		INT_PTR retCode = dlg.DoModal();
		if (retCode == IDOK)
		{
			m_mergeRootFolderStyle = dlg.m_mergeRootFolderStyle;
			m_labelAssignmentStyle = dlg.m_labelAssignmentStyle;

			wantsToRetry = FileSelectrootfolder(2); // retry == -2, cancel == -1, continue == 1
			int deb = 1;
		}
		else if (retCode == IDCANCEL)
		{
			break;
		}
		else
			break;
	}
#endif
}


void CMainFrame::OnFileSelectasrootfolder()
{
	// TODO: Add your command handler code here
	CString txt = _T("MBox Viewer will traverse selected root folder and all sub-folders and create Tree of Folders View.\n\n"
		"Do you want to continue?\n");
#if 0
	int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_YESNO);
	if (answer == IDYES)
	{
		FileSelectrootfolder(1);
	}
#else
	int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_YESNO);
	if (answer == IDYES)
	{
		int wantsToRetry = -2;
		while (wantsToRetry == -2)
		{
			wantsToRetry = FileSelectrootfolder(1); // retry == -2, cancel == -1, continue == 1
			int deb = 1;
		}
	}
#endif
}
