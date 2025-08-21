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
#include "Windows.h"
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
#include "DevelopmentCreateArchive.h"  // this is dialog
#include "GeneralOptionsDlg.h"
#include "ResHelper.h"
#include "SelectLanguageDlg.h"
#include "FileConfigurationDB.h"
#include "FontConfig.h"

#include "NTreeView.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

CString CMainFrame::m_processPath = L"";
CString CMainFrame::m_startupPath = L"";
CString CMainFrame::m_currentPath = L"";

CString CMainFrame::m_mboxviewTempPath = L"";


MBoxViewerDB CMainFrame::m_ViewerGlobalDB;

CommandLineParms CMainFrame::m_commandLineParms;

ColorStylesDB CMainFrame::m_ColorStylesDB;

BOOL CMainFrame::m_relaxedMboxFileValidation = FALSE;
BOOL CMainFrame::m_relativeInlineImageFilePath = TRUE;
BOOL CMainFrame::m_enableUserAgent = TRUE;
CString CMainFrame::m_userAgentString = L"--user-agent/132.0.0.0";
CString CMainFrame::m_numberOfHTML2ToMerge = L"1";

QWORD CMainFrame::m_lastMailDataPreviewDeletTime = 0;

// Global Font Config
int CMainFrame::m_dfltPointFontSize = 110;  //  set by Enea to 85
CString CMainFrame::m_dfltFontName = L"Tahoma";
int CMainFrame::m_dfltFontSize = 8;
int CMainFrame::m_cnfFontSize = 8;
CFont CMainFrame::m_dfltFont;
CFont CMainFrame::m_cnfFont;
CFont CMainFrame::m_dfltStatusBarFont;
// Browser 
int CMainFrame::m_dfltBrowserZoom = 100;
int CMainFrame::m_currentBrowserZoom = 0;
int CMainFrame::m_cnfBrowserZoom = 0;
// valid Zoom range
int CMainFrame::m_minBrowserZoom = 0;
int CMainFrame::m_maxBrowserZoom = 0;


#define MaxShellExecuteErrorCode 32

MBoxViewerDB::MBoxViewerDB()
{
	m_rootDBFolder = L"MBOXV";
	m_rootLocalDBFolder = L"UMBoxViewer";
	m_rootPrintSubFolder = L"PrintCache";
	m_rootImageSubFolder = L"ImageCache";
	m_rootAttachmentSubFolder = L"AttachmentCache";
	m_rootArchiveSubFolder = L"ArchiveCache";
	m_rootListSubFolder = L"ListCache";
	m_rootLabelSubFolder = L"LabelCache";
	m_rootEmlSubFolder = L"EmlCache";
	m_rootMergedSubFolder = L"MergeCache";
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
	CString rootAppDataPath = CMainFrame::GetMboxviewLocalAppDataPath();

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

int CMainFrame::CheckShellExecuteResult(HINSTANCE  result, HWND h, CStringA *filePath)
{
	int ret = 0;
	CStringW filePathW;
	if (filePath)
	{
		DWORD error;
		TextUtilsEx::Ansi2WStr(*filePath, filePathW, error);
		ret = CheckShellExecuteResult(result, h, &filePathW);
	}
	else
		ret = CheckShellExecuteResult(result, h, (CStringW*)0);

	return ret;
}

int CMainFrame::CheckShellExecuteResult(HINSTANCE  result, HWND h, CStringW *filename)
{
	if ((UINT_PTR)result <= MaxShellExecuteErrorCode) 
	{
		CString errText;
		CStringW errorTextW;
		ShellExecuteError2Text((UINT_PTR)result, errText);
		//CStringW errorTextW;
		if (filename) 
		{
			if (FileUtils::PathFileExist(*filename))
			{
#if 0
				errorTextW += L"\n\nFile exists. Open the file location and check file properties.\n";
				errorTextW += L"Make sure the default application is configured to open the file.\n";
#endif

				CString fmt = L"%s\n\nFile exists. Open the file location and check file properties.\nMake sure the default application is configured to open the file.\n";
				ResHelper::TranslateString(fmt);
				errorTextW.Format(fmt, errText);
				errorTextW.Append(L"\n\n\n");
				errorTextW.Append(*filename);
				errorTextW.Append(L"\n");
			}
			else
			{
				//errorTextW += L"\n\nFile doesn't exist.\n";

				CString fmt = L"%s\n\nFile doesn't exist.\n";
				ResHelper::TranslateString(fmt);
				errorTextW.Format(fmt, errText);
				errorTextW.Append(L"\n\n\n");
				errorTextW.Append(*filename);
				errorTextW.Append(L"\n");
			}
		}
		int answer = ::MessageBoxW(h, errorTextW, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}
	return 1;
}

int CMainFrame::CheckShellExecuteResult(HINSTANCE  result, CString &errorText)
{
	if ((UINT_PTR)result <= MaxShellExecuteErrorCode) {
		ShellExecuteError2Text((UINT_PTR)result, errorText);
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
	ON_UPDATE_COMMAND_UI(ID_FILE_SELECT_FOLDER, OnUpdateFileOpen)
	ON_UPDATE_COMMAND_UI(ID_FILE_OPTIONS, OnUpdateFileOpen)
	ON_COMMAND(ID_FILE_SELECT_FOLDER, OnFileOpen)
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
	ON_COMMAND(ID_DEVELOPMENTOPTIONS_DUMPRAWDATA, &CMainFrame::OnDevelopmentoptionsDumprawdata)
	ON_COMMAND(ID_DEVELOPMENTOPTIONS_DEVELO, &CMainFrame::OnDevelopmentoptionsDevelo)
	ON_COMMAND(ID_DEVELOPMENTOPTIONS_ABOUTSYSTEM, &CMainFrame::OnDeveloperOptionsAboutSystem)
	ON_COMMAND(ID_FILE_GENERALOPTIONS, &CMainFrame::OnFileGeneraloptions)
	ON_COMMAND(ID_HELP_SHORTCUTS, &CMainFrame::OnHelpShortcuts)
	ON_COMMAND(ID_HELP_CHANGE_LOG, &CMainFrame::OnHelpChangeLog)
	ON_COMMAND(ID_DEVELOPMENTOPTIONS_CODEPAGEINSTALLED, &CMainFrame::OnDevelopmentoptionsCodepageinstalled)
	ON_COMMAND(ID_DEVELOPMENTOPTIONS_SELECTLANGUAGE, &CMainFrame::OnDevelopmentoptionsSelectlanguage)
	//ON_UPDATE_COMMAND_UI(ID_FILE_PRINTCONFIG, &CMainFrame::OnUpdateFilePrintconfig)
	ON_COMMAND(ID_LANGUAGETOOLS_CREATERESOURCEFILE, &CMainFrame::OnLanguagetoolsCreateresourcefile)
	ON_COMMAND(ID_LANGUAGETOOLS_CREATETRANSLATIONFILE, &CMainFrame::OnLanguagetoolsCreatetranslationfile)
	ON_COMMAND(ID_LANGUAGETOOLS_RESORTTRANSLATIONFILE, &CMainFrame::OnLanguagetoolsResorttranslationfile)
	ON_COMMAND(ID_LANGUAGETOOLS_RESOURCEFILEPROPERTY, &CMainFrame::OnLanguagetoolsResourcefileproperty)
	ON_COMMAND(ID_HELP_FILEBASEDCONFIG, &CMainFrame::OnHelpFilebasedconfig)
	ON_COMMAND(ID_LANGUAGETOOLS_SPLITTRANSLATIONFILE, &CMainFrame::OnLanguagetoolsSplittranslationfile)
	ON_COMMAND(ID_LANGUAGE_SELECTLANGUANGE, &CMainFrame::OnLanguageSelectlanguange)
	ON_COMMAND(ID_LANGUAGE_HELP, &CMainFrame::OnLanguageHelp)
	ON_COMMAND(ID_TEST_CFILEOPENFAILURE, &CMainFrame::OnTestCfileopenfailure)
	ON_COMMAND(ID_FILE_FONTCONFIG, &CMainFrame::OnFileFontconfig)
	ON_WM_MEASUREITEM()
	ON_WM_DRAWITEM()
	ON_COMMAND(ID_HELP_CRASHHELP, &CMainFrame::OnHelpCrashhelp)
	ON_COMMAND(ID_DEVELOPMENTOPTIONS_TOGGLERTLFORDIALOGS, &CMainFrame::OnDevelopmentoptionsTogglertlfordialogs)
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

	m_createCompleted = FALSE;

	Com_Initialize();  // FIXME

	m_mergeSelectTargetFolder = FALSE;
	m_mergeRootFolderStyle = FALSE;
	m_labelAssignmentStyle = 0;

	CString section_font_config = CString(sz_Software_mboxview) + L"\\FontConfig";
	CString param = L"FontSize";
	int fontSize = 8;
	BOOL retval;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_font_config, param, fontSize))
	{
		if ((fontSize >= 8) && (fontSize <= 72))
			CMainFrame::m_cnfFontSize = fontSize;
		else
		{
			retval = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_font_config, param, CMainFrame::m_dfltFontSize);
			CMainFrame::m_cnfFontSize = CMainFrame::m_dfltFontSize;
		}
	}
	else
	{
		retval = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_font_config, param, CMainFrame::m_dfltFontSize);
		CMainFrame::m_cnfFontSize = CMainFrame::m_dfltFontSize;
	}

	m_ColorStylesDB.LoadFromRegistry();

	CString section_general = CString(sz_Software_mboxview) + L"\\General";

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
	m_bEnhancedSelectFolderDlg = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_general, L"enhancedSelectFolderDialog");

	DWORD relativeImgFilePath_Config;
	BOOL found = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_general, L"relativeInlineImageFilePath", relativeImgFilePath_Config);
	if (found)
		m_relativeInlineImageFilePath = relativeImgFilePath_Config;

	CString section_lastSelection = CString(sz_Software_mboxview) + L"\\LastSelection";
	m_lastPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_lastSelection, L"lastPath");

	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);

	treeColWidth = 177;
	m_PlusIcon = 0;
	m_MinusIcon = 0;
	m_bIsTreeHidden = FALSE;

	m_pListView = 0;
	m_pTreeView = 0;
	m_pMsgView = 0;

#if 0
	SimpleString* pStaticSS = 0;;
	pStaticSS = MboxMail::m_outbuf;
	pStaticSS = MboxMail::m_inbuf;
	pStaticSS = MboxMail::m_outdata;
	pStaticSS = MboxMail::m_indata;
	pStaticSS = MboxMail::m_workbuf;
	pStaticSS = MboxMail::m_tmpbuf;
	pStaticSS = MboxMail::m_largebuf; // = new SimpleString(1 * 1024 * 1024);
	//
	pStaticSS = MboxMail::m_largelocal1;
	pStaticSS = MboxMail::m_largelocal2;
	pStaticSS = MboxMail::m_largelocal3; // = new SimpleString(10 * 1024 * 1024);
	pStaticSS = MboxMail::m_smalllocal1; // = new SimpleString(1 * 1024 * 1024);
	pStaticSS = MboxMail::m_smalllocal2; // = new SimpleString(1 * 1024 * 1024);
#endif
}

void CMainFrame::CreateTooltipFont(CFont& font, CString& fontName, LOGFONT& logFont, HDC hdc)
{
	// TODO: Font setting is inconsistent currently !!!
	// To many different set font functions exist. Find time abd fix it
#if 0
	font.DeleteObject();
	int fontHeight = CMainFrame::m_dfltFontSize;
	if (CMainFrame::m_cnfFontSize != CMainFrame::m_dfltFontSize)
	{
		fontHeight = CMainFrame::m_cnfFontSize;
	}

	if (hdc == 0) hdc = ::GetWindowDC(NULL);
	int pointFontHeight = -MulDiv(fontHeight, GetDeviceCaps(hdc, LOGPIXELSY), 72);
	::ReleaseDC(NULL, hdc);

	if (!font.CreatePointFont(pointFontHeight, fontName))
		font.CreatePointFont(pointFontHeight, L"Arial");

	font.GetLogFont(&logFont);
#else
	int pointFontHeight = CMainFrame::m_dfltPointFontSize;
	if (CMainFrame::m_cnfFontSize != CMainFrame::m_dfltFontSize)
	{
		pointFontHeight = CMainFrame::m_cnfFontSize * 10;
	}

	font.DeleteObject();
	if (!font.CreatePointFont(pointFontHeight, fontName))
		font.CreatePointFont(pointFontHeight, L"Arial");
	font.GetLogFont(&logFont);
#endif
}


CMainFrame::~CMainFrame()
{
	// To stop memory leaks reports by debugger
	//MboxMail::ReleaseResources();
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

	MyPopupMenu::SetupFonts();

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

	CToolBarCtrl& wndToolBarCtrl = m_wndToolBar.GetToolBarCtrl();
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

	DWORD dwCtrlStatusStyle = WS_CHILD | WS_VISIBLE | CBRS_BOTTOM | CBRS_TOOLTIPS;  // dflt from source code except CBRS_TOOLTIPS
	//if (!m_wndStatusBar.Create(this, dwCtrlStatusStyle) ||
	if (!m_wndStatusBar.Create(this) ||
		!m_wndStatusBar.SetIndicators(indicators,
			sizeof(indicators) / sizeof(UINT)))
	{
		TRACE0("Failed to create status bar\n");
		return -1;      // fail to create
	}

	m_wndStatusBar.ModifyStyle(0, CBRS_TOOLTIPS);   // DOESN'T SEEM TO WORK for dynamic text

	HWND hwnd = GetSafeHwnd();
#if (_WIN32_WINNT >= _WIN32_WINNT_WIN10)
	float dpi = ::GetDpiForWindow(hwnd); // (WINVER >= 0x0605)  _WIN32_WINNT_WIN10
#endif
	//m_wndStatusBar.GetStatusBarCtrl().SetMinHeight(28);  // WORKS BUT MORE NEEDS TO BE DONE

	if (!m_wndDlgBar.Create(this, IDD_EDITOR,
		CBRS_ALIGN_TOP | CBRS_ALIGN_LEFT | CBRS_TOOLTIPS, AFX_IDW_DIALOGBAR))
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

	m_NamePatternParams.LoadFromRegistry();
	m_attachmentConfigParams.LoadFromRegistry();
	m_HdrFldConfig.LoadFromRegistry();

#if 0
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
#endif

	SetIcon(m_hIcon, TRUE);			// use big icon
	SetIcon(m_hIcon, FALSE);		// Use a small icon

	PostMessage(WM_CMD_PARAM_LOAD_FOLDERS_MESSAGE, 0, 0);

	CString section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacement";
#if 0
	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
		section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacementPreview";
	else if (CMainFrame::m_commandLineParms.m_bDirectFileOpenMode)
		section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacementDirect";
#endif

	WINDOWPLACEMENT wpr;
	DWORD cb = sizeof(wpr);
	BOOL  ret;
#if 0
	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
		ret = CProfile::_GetProfileBinary(HKEY_CURRENT_USER, section_wnd, L"MainFrame_EmlPreviewMode", (LPBYTE)&wpr, cb);
	else
		ret = CProfile::_GetProfileBinary(HKEY_CURRENT_USER, section_wnd, L"MainFrame", (LPBYTE)&wpr, cb);
#else
	ret = CProfile::_GetProfileBinary(HKEY_CURRENT_USER, section_wnd, L"MainFrame", (LPBYTE)&wpr, cb);
#endif
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
		; // this->SetSetFocus();
		//OnSetFocus(0);  // ???
	}

	CString section_general = CString(sz_Software_mboxview) + L"\\General";
	m_relaxedMboxFileValidation = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_general, L"relaxedMboxFileValidation");

	CString section_merge = CString(sz_Software_mboxview) + L"\\PrintConfig\\Merge";

	m_enableUserAgent = TRUE;
	int enableUserAgent = TRUE;;
	if (CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_merge, L"enableUserAgent", enableUserAgent))
		m_enableUserAgent = enableUserAgent;

	m_numberOfHTML2ToMerge = L"1";
	CString numberOfHTML2ToMerge;
	if (CProfile::_GetProfileString(HKEY_CURRENT_USER, section_merge, L"numberOfHTML2Merge", numberOfHTML2ToMerge))
		m_numberOfHTML2ToMerge = numberOfHTML2ToMerge;

#if 0
	CMenu* menu = this->GetMenu();
	if (CMainFrame::m_relaxedMboxFileValidation)
		menu->CheckMenuItem(ID_DEVELOPMENTOPTIONS_RELAXMAILFILEVALIDATION, MF_CHECKED);
	else
		menu->CheckMenuItem(ID_DEVELOPMENTOPTIONS_RELAXMAILFILEVALIDATION, MF_UNCHECKED);
#endif

	CString chromeVer = CProfile::_GetProfileString(HKEY_CURRENT_USER, L"SOFTWARE\\Google\\Chrome\\BLBeacon", L"version");
	CString edgeVer = CProfile::_GetProfileString(HKEY_CURRENT_USER, L"SOFTWARE\\Microsoft\\Edge\\BLBeacon", L"version");

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

	// Good luck with Locale functions. Big mess. Ignore for now
	LCID myDfltLocale = GetUserDefaultLCID();

	LCID   Locale = LOCALE_USER_DEFAULT;  // [in 
	LCTYPE LCType = LOCALE_FONTSIGNATURE;  // [in
	LOCALESIGNATURE LocSig;
	LPWSTR lpLCData = (LPWSTR)&LocSig;              // [out, optional] 
	int cchData = sizeof(LocSig) / sizeof(WCHAR);   // [in

	int retLocale = GetLocaleInfoW(Locale, LCType, lpLCData, cchData);

	int deb = 1;


#if (NTDDI_VERSION >= NTDDI_VISTA)

	// GetLocaleInfoEx requires _WIN32_WINNT_VISTA
	// Current #define _WIN32_WINNT _WIN32_WINNT_WINXP
	// Requried #define _WIN32_WINNT _WIN32_WINNT_VISTA

	CALID calid;
	DWORD value;

	int retLocaleExId = GetLocaleInfoEx(LOCALE_NAME_USER_DEFAULT,
		LOCALE_ICALENDARTYPE | LOCALE_RETURN_NUMBER,
		(LPWSTR)&value,
		sizeof(value) / sizeof(WCHAR));
	calid = value;

	LOCALESIGNATURE LocSigEx;

	int retLocaleExStr = GetLocaleInfoEx(LOCALE_NAME_USER_DEFAULT,
		LOCALE_FONTSIGNATURE,
		(LPWSTR)&LocSigEx,
		sizeof(LocSigEx) / sizeof(WCHAR));
#endif

	CString version;
	BOOL retGetVersion = CmboxviewApp::GetMboxviewLongVersion(version);
	if (retGetVersion)
	{
		CString section_general = CString(sz_Software_mboxview) + L"\\General";

		CString language = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"language");

		if (language.IsEmpty())
		{
			CString languageFolderPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"languageFolderPath");
			FileUtils::CPathStripPath((LPCWSTR)languageFolderPath, language);
		}

		if (language.IsEmpty())
			language = L"english";
		else if (language.CompareNoCase(L"portuguese-brazil") == 0)
		{
			language = L"Portuguese-Brazil";
		}
		else if (language.CompareNoCase(L"chinese-simplified") == 0)
		{
			language = L"Chinese-Simplified";
		}

		CString languageName(language.GetAt(0));
		languageName.MakeUpper();
		languageName.Append(language.Right(language.GetLength() - 1));

		version += L" " + languageName;

		SetWindowText(version);
	}

	//MyPopupMenu::SetupFonts();

	CString fontName = CMainFrame::m_dfltFontName;
	int fontHeight = CMainFrame::m_dfltFontSize;
	LOGFONT logFont;
	CMainFrame::CreateTooltipFont(CMainFrame::m_dfltFont, fontName, logFont);

	LOGFONT barLogFont;
	//CMainFrame::CreateTooltipFont(CMainFrame::m_dfltStatusBarFont, fontName, barLogFont);
	CMainFrame::m_dfltStatusBarFont.DeleteObject();
	if (!CMainFrame::m_dfltStatusBarFont.CreatePointFont(CMainFrame::m_dfltPointFontSize,fontName))
		CMainFrame::m_dfltStatusBarFont.CreatePointFont(CMainFrame::m_dfltPointFontSize, L"Arial");
	CMainFrame::m_dfltStatusBarFont.GetLogFont(&barLogFont);
	
	m_wndStatusBar.SetFont(&m_dfltStatusBarFont); 
	CWnd* wnd = CMainFrame::SetWindowFocus(this);
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

	if (CmboxviewApp::m_isRTL == FALSE)
	{
		cs.dwExStyle &= ~(WS_EX_MDICHILD | WS_EX_CLIENTEDGE);
	}
	else
	{
		cs.dwExStyle &= ~(WS_EX_MDICHILD | WS_EX_CLIENTEDGE | WS_EX_LTRREADING | WS_EX_RTLREADING | WS_EX_LEFT);
		cs.dwExStyle |= WS_EX_LAYOUTRTL| WS_EX_RIGHT;
	}
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
		TRACE(L"OnActivateApp: ACTIVE\n");
		CMainFrame::RestoreWindowFocus();
		int deb = 1;
	}
	else
	{
		TRACE(L"OnActivateApp: INACTIVE\n");
		// Hack to prevent selected text from being erased when switching to differnt application
		// Doesn't work when you minimize window via minimize button
		//m_wndView.SetFocus();

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
	CString section_options = CString(sz_Software_mboxview) + L"\\Options";
	COptionsDlg d;
	d.m_bEnhancedSelectFolderDlg = m_bEnhancedSelectFolderDlg;
	if (d.DoModal() == IDOK) 
	{
		NListView *pListView = GetListView();
		NMsgView *pMsgView = GetMsgView();
		NTreeView *pTreeView = GetTreeView();

		pTreeView->m_filesToValidateAsMboxType = d.m_filesToValidateAsMboxType;

		m_bEnhancedSelectFolderDlg = d.m_bEnhancedSelectFolderDlg;

		if (pListView->m_subjectSortType != d.m_bSubjectSortType) {
			pListView->m_subjectSortType = d.m_bSubjectSortType;
			DWORD subjectSortType = d.m_bSubjectSortType;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"subjectSortType", subjectSortType);
		}


		CString format = MboxMail::GetDateFormat(d.m_format);
		if (pListView->m_format.Compare(format) != 0) {
			pListView->m_format = MboxMail::GetDateFormat(d.m_format);
			pListView->m_advancedParams.m_dateTimeFormat = MboxMail::GetPickerDateFormat(d.m_format);
			pListView->m_findParams.m_dateTimeFormat = MboxMail::GetPickerDateFormat(d.m_format);
			needRedraw = true;
		}

		if (pListView->m_maxSearchDuration != d.m_barDelay) {
			pListView->m_maxSearchDuration = d.m_barDelay;
			DWORD barDelay = d.m_barDelay;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"progressBarDelay", barDelay);
		}

		BOOL exportEml = (d.m_exportEML > 0) ? TRUE : FALSE;
		if (pListView->m_bExportEml != exportEml) {
			pListView->m_bExportEml = exportEml;
			CString str_exportEML = L"n";
			if (exportEml == TRUE)
				str_exportEML = L"y";
			CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_options, L"exportEML", str_exportEML);
		}

		if (pMsgView->m_cnf_subj_charsetId != d.m_subj_charsetId) {
			pMsgView->m_cnf_subj_charsetId = d.m_subj_charsetId;
			DWORD charsetId = d.m_subj_charsetId;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"subjCharsetId", charsetId);
			needRedraw = true;
		}
		if (pMsgView->m_cnf_from_charsetId != d.m_from_charsetId) {
			pMsgView->m_cnf_from_charsetId = d.m_from_charsetId;
			DWORD charsetId = d.m_from_charsetId;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"fromCharsetId", charsetId);
			needRedraw = true;
		}
		if (pMsgView->m_cnf_to_charsetId != d.m_to_charsetId) {
			pMsgView->m_cnf_to_charsetId = d.m_to_charsetId;
			DWORD charsetId = d.m_to_charsetId;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"toCharsetId", charsetId);
			needRedraw = true;
		}
#if 0
		if (pMsgView->m_cnf_cc_charsetId != d.m_cc_charsetId) {
			pMsgView->m_cnf_cc_charsetId = d.m_cc_charsetId;
			DWORD charsetId = d.m_cc_charsetId;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"ccCharsetId", charsetId);
			needRedraw = true;
		}
		if (pMsgView->m_cnf_bcc_charsetId != d.m_bcc_charsetId) {
			pMsgView->m_cnf_bcc_charsetId = d.m_bcc_charsetId;
			DWORD charsetId = d.m_bcc_charsetId;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"bccCharsetId", charsetId);
			needRedraw = true;
		}
#endif
		if (pMsgView->m_show_charsets != d.m_show_charsets) {
			pMsgView->m_show_charsets = d.m_show_charsets;
			DWORD show_charsets = d.m_show_charsets;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"showCharsets", show_charsets);
			needRedraw = true;
		}
		if (pMsgView->m_bImageViewer != d.m_bImageViewer) {
			pMsgView->m_bImageViewer = d.m_bImageViewer;
			DWORD bImageViewer = d.m_bImageViewer;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"showCharsets", bImageViewer);
		}
		if (pListView->m_gmtTime != d.m_bTimeType)
		{
			pListView->m_gmtTime = d.m_bTimeType;
			DWORD bTimeType = d.m_bTimeType;
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"timeType", bTimeType);
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
			path.TrimRight(L"\\");
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
			path.TrimRight(L"\\");
		}
		else
		{
			int deb = 1;
			return;
		}
	}

	if (!FileUtils::PathDirExists(path))
	{
		CString txt;
		CString fmt = L"Trying to open folder \"%s\" that doesn't exist.\n\n\n";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, path);

		HWND h = GetSafeHwnd(); // we don't have any window yet
		int answer = MyMessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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

	TRACE(L"TreeState %d --> %d m_frameCy_TreeInHide=%d m_frameCx_TreeInHide=%d m_frameCy_TreeNotInHide=%d m_frameCx_TreeNotInHide=%d\n", 
		bIsTreeHidden, m_bIsTreeHidden, 
		pListView->m_frameCy_TreeInHide, pListView->m_frameCx_TreeInHide, 
		pListView->m_frameCy_TreeNotInHide, pListView->m_frameCx_TreeNotInHide);

	CRect rectMsgView;
	CRect rectListView;
	CRect rectTreeView;
	CRect rectFrameView;


	NMsgView* msgView = GetMsgView();
	NListView* listView = GetListView();
	NTreeView* treeView = GetTreeView();

	msgView->GetClientRect(&rectMsgView);
	listView->GetClientRect(&rectListView);
	treeView->GetClientRect(&rectTreeView);

	CRect rectMsgViewWindow;
	CRect rectListViewWindow;
	CRect rectTreeViewWindow;

	msgView->GetWindowRect(rectMsgViewWindow);
	listView->GetWindowRect(rectListViewWindow);
	treeView->GetWindowRect(rectTreeViewWindow);

	int listPlusMsg = rectListViewWindow.Width() + rectMsgViewWindow.Width();
	int listPlusMsgPlusTree = rectListViewWindow.Width() + rectMsgViewWindow.Width() + rectTreeViewWindow.Width();;

	this->GetClientRect(rectFrameView);

	if (bIsTreeHidden)
	{
		pListView->m_frameCx_TreeInHide = rectListViewWindow.Width();
		pListView->m_frameCy_TreeInHide = rectListViewWindow.Height();

		pMsgView->m_frameCx_TreeInHide = rectMsgViewWindow.Width();
		pMsgView->m_frameCy_TreeInHide = rectMsgViewWindow.Height();

		if (rectTreeViewWindow.Width() == 0)
			pTreeView->m_frameCx = 200;
		//pTreeView->m_frameCy = rectTreeViewWindow.Height();
	}
	else
	{
		pListView->m_frameCx_TreeNotInHide = rectListViewWindow.Width();
		pListView->m_frameCy_TreeNotInHide = rectListViewWindow.Height();

		pMsgView->m_frameCx_TreeNotInHide = rectMsgViewWindow.Width();
		pMsgView->m_frameCy_TreeNotInHide = rectMsgViewWindow.Height();

		pTreeView->m_frameCx = rectTreeViewWindow.Width();
		pTreeView->m_frameCy = rectTreeViewWindow.Height();
	}

	if (!bIsTreeHidden)
	{
		m_wndView.m_verSplitter.SetColumnInfo(rowTree, 0, 0);

		if (m_pListView)
		{
			if (m_msgViewPosition == 1)
				//m_wndView.m_horSplitter.SetRowInfo(rowList, pListView->m_frameCx_TreeInHide, 0);
				//m_wndView.m_horSplitter.SetColumnInfo(colList, pListView->m_frameCx_TreeInHide, 0);
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

	//m_bIsTreeHidden = !m_bIsTreeHidden;

	int deb = 1;
}

BOOL CMainFrame::DoOpen(CString& path) 
{
	//if (path.Find(MboxMail::s_folderContext.m_rootDataFolderPathConfig) == 0)
	if (path.Find(L"UMBoxViewer") >= 0)  // changed to look for UMBoxViewer subfolder created under root data folder
	{
		CString fpath = MboxMail::s_folderContext.m_rootDataFolderPathConfig + L"\\UMBoxViewer";
		CString txt;
#if 0
		txt.Format(L"Can't open folder \n\n\"%s\"\n\nFolder must not be located under \"UMBoxViewer\" subfolder created under the configured root data folder \n\n\"%s\"\n\n"
			L"Please manually move the respective mbox files to valid folder. ",
				path, fpath);
#endif

		CString fmt = L"Can't open folder \n\n\"%s\"\n\nFolder must not be located under \"UMBoxViewer\" subfolder created under the configured root data folder \n\n\"%s\"\n\n"
			L"Please manually move the respective mbox files to valid folder.";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, path, fpath);

		int answer = MyMessageBox(txt, L"Warning", MB_APPLMODAL | MB_ICONWARNING | MB_OK);
		return FALSE;
	}

	NTreeView *pTreeView = GetTreeView();
	if (pTreeView)
	{
		BOOL expand = TRUE;
		pTreeView->DoOpen(path, expand);

		if (CMainFrame::m_relaxedMboxFileValidation == FALSE)
		{
			HTREEITEM hRoot = 0;
			BOOL recursive = TRUE;
			CString folderPath;

			folderPath = path;
			folderPath.TrimRight(L"\\");
			HTREEITEM hItem;
			BOOL retFind = pTreeView->FindFolder(hRoot, folderPath, hItem, recursive);

			FolderInfo *finfoBeg = pTreeView->m_globalFolderInfoDB.Find(&path);
			int mailFolderCount = 0;
			if (hItem)
			{
				mailFolderCount = pTreeView->GetChildrenCount(hItem, recursive);
				if (mailFolderCount == 0)
				{
					CString txt = L"No mail archive files were discovered using strong mail file validation rules\n\n"
						L"Do you wish to rerun discovery using relaxed mail file validation rules ???\n";
					ResHelper::TranslateString(txt);

					int answer = MyMessageBox(txt, L"Warning", MB_APPLMODAL | MB_ICONWARNING | MB_YESNO);

					if (answer == IDYES)
					{
						CMainFrame::m_relaxedMboxFileValidation = TRUE;
						BOOL ret = pTreeView->RefreshFolder(hItem);

						HTREEITEM hItem;
						BOOL retFind = pTreeView->FindFolder(hRoot, folderPath, hItem, recursive);
						if (hItem)
						{
							mailFolderCount = pTreeView->GetChildrenCount(hItem, recursive);
							if (mailFolderCount == 0)
							{
								int deb1 = 0;
							}
						}
					}
					int deb = 1;
				}
			}
			else
			{
				CString txt = L"No mail archive files were discovered using strong mail file validation rules\n\n"
					L"Do you wish to rerun discovery using relaxed mail file validation rules ???\n";
				ResHelper::TranslateString(txt);

				int answer = MyMessageBox(txt, L"Warning", MB_APPLMODAL | MB_ICONWARNING | MB_YESNO);

				if (answer == IDYES)
				{
					CMainFrame::m_relaxedMboxFileValidation = TRUE;
					pTreeView->DoOpen(path, expand);

					folderPath = path;
					folderPath.TrimRight(L"\\");
					HTREEITEM hItem;
					BOOL retFind = pTreeView->FindFolder(hRoot, folderPath, hItem, recursive);
					if (hItem)
					{
						mailFolderCount = pTreeView->GetChildrenCount(hItem, recursive);
						if (mailFolderCount == 0)
						{
							int deb1 = 0;
						}
						int deb = 1;
					}
				}
			}
			if (CMainFrame::m_relaxedMboxFileValidation == TRUE)  //eser requested discovery using Relaxed Validation Rules
			{
				if (mailFolderCount == 0)
				{
					CString txt = L"No mail archive files were discovered using relaxed mail file validation rules\n\n"
						L"Do you wish to keep relaxed mail file validation rules for the future discovery ?\n\n"
						L"The relaxed mail file validation rules can be later enabled/disabled "
						L"by selecting \"File->Development->Relaxed Mail File Validation\" option\n\n";
					ResHelper::TranslateString(txt);

					int answer = MyMessageBox(txt, L"Warning", MB_APPLMODAL | MB_ICONWARNING | MB_YESNO);
					if (answer == IDNO)
					{
						CMainFrame::m_relaxedMboxFileValidation = FALSE;
					}
					int deb1 = 0;
				}
				else
				{
					CString txt = L"Several mail archive files were discovered using relaxed mail file validation rules\n\n"
						L"Do you wish to keep relaxed mail file validation rules for the future discovery ?\n\n"
						L"The relaxed mail file validation rules can be later enabled/disabled "
						L"by selecting \"File->Development->Relexed Mail File Validation\" option\n\n";
					ResHelper::TranslateString(txt);

					int answer = MyMessageBox(txt, L"Warning", MB_APPLMODAL | MB_ICONWARNING | MB_YESNO);
					if (answer == IDNO)
					{
						CMainFrame::m_relaxedMboxFileValidation = FALSE;
					}
				}

				CMenu* menu = this->GetMenu();
				UINT retval = 0;
				if (CMainFrame::m_relaxedMboxFileValidation)
					retval = menu->CheckMenuItem(ID_DEVELOPMENTOPTIONS_RELAXMAILFILEVALIDATION, MF_CHECKED);
				else
					retval = menu->CheckMenuItem(ID_DEVELOPMENTOPTIONS_RELAXMAILFILEVALIDATION, MF_UNCHECKED);

				CString section_general = CString(sz_Software_mboxview) + L"\\General";
				CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_general, L"relaxedMboxFileValidation", (WORD)m_relaxedMboxFileValidation);
			}
		}
		int deb = 1;
	}
	return TRUE;
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
		return (NListView *)m_wndView.m_horSplitter.GetPane(0, 0);  // should never be here _ASSERTE ?
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
		return (NMsgView *)m_wndView.m_horSplitter.GetPane(1, 0);  // should never be here _ASSERTE ?
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
					//CString text = L"Created file \n\n" + csvFileName;
					CString text;
					CString fmt = L"Created file \n\n%s";
					ResHelper::TranslateString(fmt);
					text.Format(fmt, csvFileName);

					OpenContainingFolderDlg dlg(text, csvFileName);
					INT_PTR nResponse = dlg.DoModal();
					if (nResponse == IDOK)
					{
						if (FileUtils::BrowseToFile(csvFileName) == FALSE) {
							HWND h = GetSafeHwnd();
							HINSTANCE result = ShellExecute(h, L"open", datapath, NULL, NULL, SW_SHOWNORMAL);
							CheckShellExecuteResult(result, h);
						}
						int deb = 1;
					}
					else if (nResponse == IDYES)
					{
						HWND h = GetSafeHwnd();
						HINSTANCE result = ShellExecute(h, L"open", csvFileName, NULL, NULL, SW_SHOWNORMAL);
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

	CString HelpPath = CMainFrame::GetMboxviewTempPath(L"MboxHelp");

	BOOL createDirOk = TRUE;
	if (!FileUtils::PathDirExists(HelpPath))
		createDirOk = FileUtils::CreateDir(HelpPath);

	int ret = TextUtilsEx::showCodePageTable(HelpPath);
}

// Called when File->"Print To"->CSV
// File->"Print To"->CVS -> OnPrinttoCsv() -> OnFileExportToCsv() -> exportToCSVFile() for firstMail;lastMail -> NO separate printSingleMailToTextFile(index) yet
void CMainFrame::OnPrinttoCsv()
{
	// TODO: Add your command handler code here
	if (MboxMail::s_mails.GetSize() == 0) {
		CString txt = L"Please open mail file first.";
		ResHelper::TranslateString(txt);
		HWND h = GetSafeHwnd(); // we don't have any window yet
		int answer = MyMessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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
		CString txt = L"Please open mail file first.";
		ResHelper::TranslateString(txt);
		HWND h = GetSafeHwnd(); 
		int answer = MyMessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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
					//CString text = L"Created file \n\n" + textFileName;
					CString text = L"Created file";
					ResHelper::TranslateString(text);   // RESTODO

					OpenContainingFolderDlg dlg(text, textFileName);
					INT_PTR nResponse = dlg.DoModal();
					if (nResponse == IDOK)
					{
						if (FileUtils::BrowseToFile(textFileName) == FALSE) {
							HWND h = GetSafeHwnd();
							HINSTANCE result = ShellExecute(h, L"open", datapath, NULL, NULL, SW_SHOWNORMAL);
							CheckShellExecuteResult(result, h);
						}
						int deb = 1;
					}
					else if (nResponse == IDYES)
					{
						HWND h = GetSafeHwnd();
						HINSTANCE result = ShellExecute(h, L"open", textFileName, NULL, NULL, SW_SHOWNORMAL);
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
		CString txt = L"Please open mail file first.";
		ResHelper::TranslateString(txt);
		HWND h = GetSafeHwnd(); 
		int answer = MyMessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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
					//CString text = L"Created file \n\n" + textFileName;
					CString text = L"Created file";
					ResHelper::TranslateString(text);
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

							CString folder;
							FileUtils::CPathGetPath(textFileName, folder);


							// FIXMEFIXME
							DWORD BUFSIZE = 1024;
							wchar_t Buffer[1024];
							wchar_t Buffer2[1024];

							DWORD dwRet = ::GetCurrentDirectory(BUFSIZE, Buffer);

							// works with or without setting directory
							if (!::SetCurrentDirectory((LPCWSTR)folder))
							{
								;
							}

							dwRet = ::GetCurrentDirectory(BUFSIZE, Buffer2);

							UINT inCodePage = CP_UTF8;
							NMsgView::PrintHTMLDocumentToPrinter(inbuf, workbuf, inCodePage);

							if (!::SetCurrentDirectory(Buffer))
							{
								;
							}

							int deb = 1;
						}
						else
						{
#if 1
							DWORD lastErr = ::GetLastError();
							HWND h = GetSafeHwnd();

							//CString fmt = L"Could not open \"%s\" mail file.\n%s";
							CString fmt = L"Could not open mail file:\n\n\"%s\"\n\n%s";
							CString errorText = FileUtils::ProcessCFileFailure(fmt, textFileName, ExError, lastErr, h);
#else

							DWORD lastErr = ::GetLastError();
							CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError, lastErr);

							CString fmt = L"Could not open \"%s\" mail file.\n%s";
							ResHelper::TranslateString(fmt);
							CString txt;
							txt.Format(fmt, textFileName, exErrorStr);

							TRACE(L"%s\n", txt);

							HWND h = GetSafeHwnd();
							int answer1 = MyMessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
							int answer2 = FileUtils::CheckIfFileLocked(textFileName, lastErr, h);
#endif
							int deb = 1;
						}
					}
					else if (forceOpen == FALSE)
					{
						OpenContainingFolderDlg dlg(text, textFileName);
						INT_PTR nResponse = dlg.DoModal();
						if (nResponse == IDOK)
						{
							if (FileUtils::BrowseToFile(textFileName) == FALSE) {
								HWND h = GetSafeHwnd();
								HINSTANCE result = ShellExecute(h, L"open", datapath, NULL, NULL, SW_SHOWNORMAL);
								CheckShellExecuteResult(result, h);
							}
							int deb = 1;
						}
						else if (nResponse == IDYES)
						{
							HWND h = GetSafeHwnd();
							HINSTANCE result = ShellExecute(h, L"open", textFileName, NULL, NULL, SW_SHOWNORMAL);
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
						HINSTANCE result = ShellExecute(h, L"open", textFileName, NULL, NULL, SW_SHOWNORMAL);
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

int  CMainFrame::PrintSingleMailtoPDF(int iItem, CString& targetPrintSubFolderName, BOOL progressBar, CString& progressText, CString& errorText)
{
	CString textFileName;

	TEXTFILE_CONFIG textConfig;
	NListView* pListView = CMainFrame::GetListView();
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

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	CString fileName;

	fileName.Append(textFileName);
	CString newSuffix;
	newSuffix.Append(L".pdf");
	FileUtils::UpdateFileExtension(fileName, newSuffix);

	CString htmFileName;
	htmFileName.Append(textFileName);

	// CString progressText;  // FIXME
	TRACE(L"Printing %s to PDF\n", htmFileName);

	const int timeout = 1200; // seconds FIXME. this is MBox Viewer internal hardcode print to PDF timeout
	// I am aware of cases it takes 10 or minutes to finish 
	// or print will never finish sometimes. Had to pick something
	const int headlessTimout = -1;  // Don't configure --timeout >= 0 option to Edge/Chrome, doesn't work as documented
	// reported --timeout issue to Chromium team July 5, 2023
	// --timeout it is fixed in Canary version but it was not integrated into official release yet
	ret = ExecCommand_WorkerThread(htmFileName, errorText, progressBar, progressText, timeout, headlessTimout);
	if (ret == -2)
	{
#if 0
		// internal hard timer fired which means print failed, retry
		// --timeout and --virtual-time-budget options don't seem to work. Will not retry
		// 
		// set headlessTimeout and retry
		errorText.Empty();
		const int headlessTiemout = 6;
		ret = ExecCommand_WorkerThread(htmFileName, errorText, progressBar, progressText, timeout, headlessTiemout);
#endif
		int deb = 1;
	}		
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
	{
		strPage = L"Mail Retrieval Complete";
	}
	else
	{
		strPage = L"Mail Retrieval In Progress ...";
	}
	ResHelper::TranslateString(strPage);
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
		strMailIndex.Format(L"Mail %d of %d", listView->m_lastSel + 1, mailCnt);
	else
		strMailIndex.Format(L"Mail xx of %d", mailCnt);

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
		int deb = 1; // TODO: _ASSERTE ?


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

	UINT retval = 0;
	retval = menu->CheckMenuItem(ID_VIEW_USERSELECTEDMAILS, MF_CHECKED);
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

#ifdef _DEBUG
	CMenu* menu = (MyPopupMenu*)this->GetMenu();
	CString title = L"SetNewMenu ";
	int index = 0;
	//UINT mask = MyPopupMenu::MenuItemInfoMaskTypeAllSet;
	UINT mask = MyPopupMenu::MenuItemInfoMaskFTypeAllSet;

	BOOL retval = MyPopupMenu::TraceMenu(title, menu, index, mask);
#endif

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

	if (pListView)
	{
		LRESULT lres = pListView->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
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
	if (pListView)
	{
		LRESULT lres = pListView->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
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
	if (pListView)
	{
		LRESULT lres = pListView->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
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

	CWnd *p = m_wndDlgBar.GetDlgItem((int)pNMCD->dwItemSpec);
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

	CString helpFileName = L"MailListsInfo.pdf";
	HWND h = GetSafeHwnd();
	CMainFrame::OpenHelpFile(helpFileName, h);

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

void CMainFrame::OnViewUserselectedmails()
{
	// TODO: Add your command handler code here
	CMenu *menu = this->GetMenu();
	UINT retval = 0;
	if (m_bUserSelectedMailsCheckSet)
	{
		retval = menu->CheckMenuItem(ID_VIEW_USERSELECTEDMAILS, MF_UNCHECKED);
		m_bUserSelectedMailsCheckSet = FALSE;
		EnableMailList(IDC_EDIT_LIST, FALSE);
		//TRACE(L"OnViewUserselectedmails: MF_CHECKED->MF_UNCHECKED\n");

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
		retval = menu->CheckMenuItem(ID_VIEW_USERSELECTEDMAILS, MF_CHECKED);
		m_bUserSelectedMailsCheckSet = TRUE;
		EnableMailList(IDC_EDIT_LIST, TRUE);
		//TRACE(L"OnViewUserselectedmails: MF_UNCHECKED->MF_CHECKED\n");
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

	CString helpFileName = L"MboxviewerHelp.pdf";
	HWND h = GetSafeHwnd();
	CMainFrame::OpenHelpFile(helpFileName, h);

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
	UINT retval = 0;
	if (msgViewPosition == 1)
	{
		retval = menu->CheckMenuItem(ID_MESSAGEWINDOW_BOTTOM, MF_CHECKED);
		retval = menu->CheckMenuItem(ID_MESSAGEWINDOW_RIGHT, MF_UNCHECKED);
		retval = menu->CheckMenuItem(ID_MESSAGEWINDOW_LEFT, MF_UNCHECKED);
	}
	else if (msgViewPosition == 2)
	{
		retval = menu->CheckMenuItem(ID_MESSAGEWINDOW_BOTTOM, MF_UNCHECKED);
		retval = menu->CheckMenuItem(ID_MESSAGEWINDOW_RIGHT, MF_CHECKED);
		retval = menu->CheckMenuItem(ID_MESSAGEWINDOW_LEFT, MF_UNCHECKED);
	}
	else if (msgViewPosition == 3)
	{
		retval = menu->CheckMenuItem(ID_MESSAGEWINDOW_BOTTOM, MF_UNCHECKED);
		retval = menu->CheckMenuItem(ID_MESSAGEWINDOW_RIGHT, MF_UNCHECKED);
		retval = menu->CheckMenuItem(ID_MESSAGEWINDOW_LEFT, MF_CHECKED);
	}
}

void CMainFrame::ConfigMessagewindowPosition(int msgViewPosition)
{
	if ((msgViewPosition < 1) && (msgViewPosition > 3)) {
		int deb = 1;
		return;  // should never be here
	}

	int curMsgViewPosition = m_msgViewPosition;

	CString curPos = L"Bottom";
	if (curMsgViewPosition == 1)
		curPos = L"Bottom";
	else if (curMsgViewPosition == 2)
		curPos = L"Right";
	else if (curMsgViewPosition == 3)
		curPos = L"Left";
	else
		; // shpould never be here

	CString newPos = L"Bottom";
	if (msgViewPosition == 1)
		newPos = L"Bottom";
	else if (msgViewPosition == 2)
		newPos = L"Right";
	else if (msgViewPosition == 3)
		newPos = L"Left";
	else
		; // shpould never be here


	HWND h = GetSafeHwnd();

	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	if (msgViewPosition == curMsgViewPosition)
	{
		//CString txt = L"New and current Message Window position are the same: " + newPos + L" and " + curPos + L".\nNo action will be taken.";


		CString txt;
		CString fmt = L"New and current Message Window position are the same: %s and %s.\nNo action will be taken.";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, newPos, curPos);

		int answer = MyMessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONEXCLAMATION | MB_OK);
		m_newMsgViewPosition = msgViewPosition;
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_general, L"messageWindowPosition", m_newMsgViewPosition);

		CString section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacement";
#if 0
		if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
			section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacementPreview";
		else if (CMainFrame::m_commandLineParms.m_bDirectFileOpenMode)
			section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacementDirect";
#endif

		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"messageWindowPosition", m_newMsgViewPosition);
		return;
	}

#if 0
	CString txt = L"Do you want to configure Message Window position from " + curPos + " to " + newPos 
		+ " ?\nIf you say Yes, please exit and restart mbox viewer for new position to take effect.";
#endif

	CString txt;
	CString fmt = L"Do you want to configure Message Window position from %s to %s ?\nIf you say Yes, please exit and restart mbox viewer for new position to take effect.";
	ResHelper::TranslateString(fmt);
	txt.Format(fmt, curPos, newPos);

	int answer = MyMessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
	if (answer == IDYES) 
	{
		m_newMsgViewPosition = msgViewPosition;
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_general, L"messageWindowPosition", m_newMsgViewPosition);

		CString section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacement";
#if 0
		if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
			section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacementPreview";
		else if (CMainFrame::m_commandLineParms.m_bDirectFileOpenMode)
			section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacementDirect";
#endif

		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"messageWindowPosition", m_newMsgViewPosition);
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
	{
		CMainFrame::RestoreWindowFocus();
		int deb = 1;
	}


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
	wchar_t *fileNameBuffer = (LPWSTR)MboxMail::m_outbuf->Data();

	//CString  sFilters  = "Mail Files (*.mbox)|*.mbox|All Files (*.*)|*.*||";
	CString  sFilters = L"Mail Files All Files (*.*)|*.*|(*.mbox;*.eml)|*.mbox;*eml|(*.mbox)|*.mbox|(*.eml)|*.eml||";

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
	ofn.lpstrTitle = L"Select Mail Archives For Merging";

restart:
	INT_PTR ret = dlgFile.DoModal();
	if (ret == IDOK)
	{
		CString folder = dlgFile.GetFolderPath();
		TRACE(L"FOLDER=%s\n", folder);

		int fileCount = 0;
		CString strFilePath;
		POSITION pos = dlgFile.GetStartPosition();
		while (NULL != pos) 
		{
			fileCount++;
			strFilePath = dlgFile.GetNextPathName(pos);
			if (!FileUtils::PathFileExist(strFilePath))
			{
#if 0
				CString txt = L"File path invalid.\n";
				txt.Append(strFilePath);
				txt.Append(L"\nRetry or Cancel?");
#endif
				CString txt;
				CString fmt = L"File path invalid.\n%s\nRetry or Cancel?";
				ResHelper::TranslateString(fmt);
				txt.Format(fmt, strFilePath);

				int answer = MyMessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_RETRYCANCEL);
				if (answer == IDRETRY)
				{
					goto restart;
				}
			}
			TRACE(L"FILE=%s\n", strFilePath);
		}

		// All archive files are valid; merge
		CString title = L"Enter Name for New Archive File";
		CString  fileNameFilter = L"Mail Files (*.mbox)|*.mbox||";
		CString dfltExtension = L".mbox";

		CString outFolderPath;
		CString fileName;
		CString datapath = MboxMail::GetLastDataPath();
		if (inFolderPath.Compare(datapath))
		{
			CString appFolderPath = CMainFrame::GetMboxviewLocalAppDataPath();
			inFolderPath = appFolderPath + L"MboxRepo";
			FileUtils::CreateDir(inFolderPath);
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
				if (FileUtils::DelFile(viewFile, FALSE))
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
				DWORD lastErr = ::GetLastError();
#if 1
				HWND h = GetSafeHwnd();
				CString fmt = L"Could not create file:\n\n\"%s\"\n\n%s";
				CString errorText = FileUtils::ProcessCFileFailure(fmt, filePath, ExError, lastErr, h);

#else
				CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError, lastErr);

				CString txt;
				CString fmt = L"Could not create \"%s\" file.\n%s";
				ResHelper::TranslateString(fmt);
				txt.Format(fmt, filePath, exErrorStr);

				TRACE(L"%s\n", txt);

				int answer = MyMessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
				HWND h = GetSafeHwnd();
				int answer2 = FileUtils::CheckIfFileLocked(filePath, lastErr, h);
#endif

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
						DWORD lastErr = ::GetLastError();
#if 1
						fp.Close();
						MboxMail::m_outbuf->Clear();

						HWND h = GetSafeHwnd();
						CString fmt = L"Could not open file:\n\n\"%s\"\n\n%s";  // new format
						CString errorText = FileUtils::ProcessCFileFailure(fmt, strFilePath, ExError, lastErr, h); 
#else
						CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError, lastErr);  // TODO

						fp.Close();
						MboxMail::m_outbuf->Clear();

						CString txt;
						CString fmt = L"Could not open \"%s\" file.\nMake sure file is not open on other applications.\n%s";
						ResHelper::TranslateString(fmt);
						txt.Format(fmt, filePath, exErrorStr);

						TRACE(L"%s\n", txt);

						//HWND h = NULL; // we don't have any window yet
						int answer = MyMessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
						HWND h = GetSafeHwnd();
						int answer2 = FileUtils::CheckIfFileLocked(filePath, lastErr, h);
#endif

						return -1;
					}
					// Append file
					UINT wantBytes = 16 * 1024;
					// don't use  MboxMail::m_oubuf is still used by dlgFile above !!!!! 
					MboxMail::m_inbuf->ClearAndResize(wantBytes+2);
					char *inBuffer = MboxMail::m_inbuf->Data();

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
				TRACE(L"FILE=%s\n", strFilePath);
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
					HINSTANCE result = ShellExecute(h, L"open", outFolderPath, NULL, NULL, SW_SHOWNORMAL);
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
		DWORD lastErr = ::GetLastError();
#if 1
		HWND h = GetSafeHwnd();
		//CString fmt = L"Could not create file:\n\n\"%s\"\n\n%s";  // new format
		CString fmt = L"Could not create Merge To File \"%s\" file.\n%s";  // old format
		CString errorText = FileUtils::ProcessCFileFailure(fmt, mergedMboxFilePath, exMergeTo, lastErr, h);

#else
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exMergeTo, lastErr);

		CString txt;
		CString fmt = L"Could not create Merge To File \"%s\" file.\n%s";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, mergedMboxFilePath, exErrorStr);

		TRACE(L"%s\n", txt);

		CFileStatus rStatus;
		BOOL ret = fpMergeTo.GetStatus(rStatus);

		//HWND h = NULL; // we don't have any window yet
		int answer = MyMessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		HWND h = GetSafeHwnd();
		int answer2 = FileUtils::CheckIfFileLocked(mergedMboxFilePath, lastErr, h);
#endif

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

		TRACE(L"FilePath=%s\n", filePath);

		_int64 fileSize = FileUtils::FileSize(fullPath);
		FileUtils::GetFileName(fullPath, fileName);

		CString fmt = L"Merging %s (%lld bytes) ...";
		ResHelper::TranslateString(fmt);

		sText.Format(fmt, fileName, fileSize);
		TRACE(L"%s\n", sText);
		if (pFrame)
			pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

		int ret = MergeMboxArchiveFile(fpMergeTo, fullPath, i == 0);

	}

	fpMergeTo.Close();

	_int64 fileSize = FileUtils::FileSize(mergedMboxFilePath);
	TRACE(L"Created merge file %s (%lld bytes) ...", mergedMboxFilePath, fileSize);

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
			filePath.TrimRight(L"\\");

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
	//if (!fpList.Open(mboxListFilePath, CFile::typeText | CFile::modeRead | CFile::typeUnicode | CFile::shareDenyNone, &exList))
	//if (!fpList.Open(mboxListFilePath, CFile::modeRead | CFile::typeUnicode | CFile::shareDenyNone, &exList))
	if (!fpList.Open(mboxListFilePath, CFile::modeRead | CFile::shareDenyNone, &exList))
	{
		DWORD lastErr =:: GetLastError();
#if 1
		HWND h = GetSafeHwnd();
		CString fmt = L"Could not open file:\n\n\"%s\"\n\n%s";  // new format
		CString errorText = FileUtils::ProcessCFileFailure(fmt, mboxListFilePath, exList, lastErr, h);
#else
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exList);

		CString txt;
		CString fmt = L"Could not open list file \"%s\" file.\n%s";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, mboxListFilePath, exErrorStr);

		TRACE(L"MergeMboxArchiveFiles: %s\n", txt);
		//errorText = txt;

		CFileStatus rStatus;
		BOOL ret = fpList.GetStatus(rStatus);

		HWND h = NULL; // we don't have any window yet
		int answer = MyMessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
#endif
		return -1;
	}

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;
	CString folderPath;
	CString fileNameExtentionRoot;

	CString filePath;

	unsigned char BOM[2];
	UINT retlen = fpList.Read(BOM, 2);
	unsigned char BOM1 = BOM[0];
	unsigned char BOM2 = BOM[1];
	if ((BOM1 == 0xFF) && (BOM2 == 0xFE))
	{
		fpList.Close();
		if (!fpList.Open(mboxListFilePath, CFile::modeRead | CFile::typeUnicode | CFile::shareDenyNone, &exList))
		{
			DWORD lastErr = ::GetLastError();
#if 1
			HWND h = GetSafeHwnd();
			//HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
			CString fmt = L"Could not open file:\n\n\"%s\"\n\n%s";  // new format
			CString errorText = FileUtils::ProcessCFileFailure(fmt, mboxListFilePath, exList, lastErr, h); 
#else
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exList);

			CString txt;
			CString fmt = L"Could not open list file \"%s\" file.\n%s";
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, mboxListFilePath, exErrorStr);

			TRACE(L"MergeMboxArchiveFiles: %s\n", txt);
			//errorText = txt;

			CFileStatus rStatus;
			BOOL ret = fpList.GetStatus(rStatus);

			HWND h = NULL; // we don't have any window yet
			int answer = MyMessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
#endif
			return -1;
		}
		UINT retlen = fpList.Read(BOM, 2);
		int deb = 1;
	}
	else
	{

		fpList.SeekToBegin();
	}


	// TODO: To optimize, save all results in an Array for further processing
	CArray<MergeFileInfo> fileList;
	while (fpList.ReadString(filePath))
	{
		const wchar_t* str = (LPCWSTR)filePath;

		if (TextUtilsEx::isWhiteLine(filePath))
			continue;

		if (filePath.CompareNoCase(mergedMboxFilePath) == 0)
		{
#if 0
			CString txt = L"Invalid File Path:\n\n";
			txt.Append(filePath);
			txt.Append(L"\n\nin List File:\n\n");
			txt.Append(mboxListFilePath);
			txt.Append(L"\n\nMbox File Path and Merge To File Path can't be the same\nCan't continue merging files\n");
#endif
			CString txt;
			CString fmt = L"Invalid File Path:\n\n%s\n\nin List File:\n\n%s\n\nMbox File Path and Merge To File Path can't be the same\nCan't continue merging files\n";
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, filePath, mboxListFilePath);

			int answer = MyMessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);

			fpList.Close();
			return -1;
		}

		filePath.TrimRight(L"\r\n"); // FIXMEFIXME I should not have to do that  to remeove \r
		FileUtils::SplitFilePath(filePath, driveName, directory, fileNameBase, fileNameExtention);
		folderPath = driveName + directory;
		
		// check if wildcard name 
		// should we search for *.* rather than *; probably not if we want to examine file without extension
		if (fileNameBase.Find(L"*") >= 0)  
		{
			// examine all files in folderPath

			if (!FileUtils::PathDirExists(folderPath))  // should never fail =:)  ???
			{
#if 0
				CString txt;
				txt.Format(L"Invalied Folder path \"%s\".\n\nCan't continue merging files\n", filePath);
#endif


				CString txt;
				CString fmt = L"Invalied Folder path \"%s\".\n\nCan't continue merging files\n";
				ResHelper::TranslateString(fmt);
				txt.Format(fmt, filePath);

				int answer = MyMessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);

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
			TRACE(L"FilePath=%s\n", filePath);
			if (!FileUtils::PathFileExist(filePath))
			{
#if 0
				CString txt = L"Invalid File Path:\n\n";
				txt.Append(filePath);
				txt.Append(L"\n\nin List File:\n\n");
				txt.Append(mboxListFilePath);
				txt.Append(L"\n\nCan't continue merging files\n");
#endif
				CString txt;
				CString fmt = L"Invalid File Path:\n\n%s\n\nin List File:\n\n%s\n\nCan't continue merging files\n";
				ResHelper::TranslateString(fmt);
				txt.Format(fmt, filePath, mboxListFilePath);

				int answer = MyMessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);

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
	hashsum_t hashsum = GlobalFolderInfoDB::GetHashsum(&mboxFolderPath);
	FolderInfo* finfo = pTreeView->m_globalFolderInfoDB.Find(&mboxFolderPath, hashsum);

	BOOL openret = TRUE;
	if (finfoBeg == 0)
	{
		openret = DoOpen(mboxFolderPath);
		if (openret == FALSE)
			return -1;
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
			DWORD lastErr =::GetLastError();
#if 1
			//HWND h = GetSafeHwnd();
			HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
			CString fmt = L"Could not open mail file:\n\n\"%s\"\n\n%s";  // new format
			CString errorText = FileUtils::ProcessCFileFailure(fmt, mboxFilePath, ExError, lastErr, h); 
#else
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt;
			CString fmt = L"Could not open \"%s\" mail file.\n%s";
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, mboxFilePath, exErrorStr);

			//TRACE(L"%s\n", txt);
			//errorText = txt;

			HWND h = NULL; // we don't have any window yet
			int answer = MyMessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
#endif
			return -1;
		}

		// Append file
		UINT wantBytes = 16 * 1024;
		// don't use  MboxMail::m_oubuf is still used by dlgFile above !!!!! 
		MboxMail::m_inbuf->ClearAndResize(wantBytes + 2);
		char *inBuffer = MboxMail::m_inbuf->Data();

		UINT readBytes = 0;

		readBytes = fp_input.Read(inBuffer, wantBytes);
		if (readBytes <= 0)
		{
			fp_input.Close();
			return -1;
		}

		static const char *cFromMailBegin = "From ";
		static const int cFromMailBeginLen = istrlen(cFromMailBegin);

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
			CStringA FromLine = "From 1513218656940664977@xxx ";

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
		readBytes = IntPtr2Int(e - p);

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
	BOOL ret = FileUtils::CreateDir(mergeCachePath);

	for (;;)
	{
		MboxMail::m_inbuf->ClearAndResize(FILE_LIST_BUFFER_SIZE);
		wchar_t *fileNameBuffer = (LPWSTR)MboxMail::m_inbuf->Data();

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

			TRACE(L"FOLDER=%s\n", outFolderPath);
			TRACE(L"FILE=%s\n", fileName);

			MboxMail::m_inbuf->Clear();

			if (outFolderPath.Compare(mergeCachePath))
			{
				CString txt = L"Changing target folder for merged file is not allowed.\n\nPlease try again.";
				ResHelper::TranslateString(txt);

				HWND h = GetSafeHwnd(); // we don't have any window yet
				int answer = MyMessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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

void CMainFrame::OnPrinttoPdf()
{
	// TODO: Add your command handler code here
	if (MboxMail::s_mails.GetSize() == 0) {
		CString txt = L"Please open mail file first.";
		ResHelper::TranslateString(txt);

		HWND h = GetSafeHwnd(); // we don't have any window yet
		int answer = MyMessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;
	}

	NListView *pListView = GetListView();
	int firstMail = 0;
	int lastMail = MboxMail::s_mails.GetCount() - 1;
	if (m_NamePatternParams.m_bPrintToSeparatePDFFiles)
	{
		if (pListView)
		{
			CString targetPrintFolderPath;
			CString targetPrintSubFolderName = L"PDF_GROUP";
			BOOL mergePDFs = FALSE;
			pListView->PrintMailRangeToSeparatePDF_Thread(firstMail, lastMail, targetPrintSubFolderName, targetPrintFolderPath, mergePDFs);
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
			errorText = L"Path to Chrome Browser not valid.\nPlease make sure Chrome is installed and path is correct."
				"\nOr configure Microsoft Edge Browser for printing."
				"\nSelect File->Print Config to update the setup.";
			return -1;
		}
	}
	else if (m_NamePatternParams.m_bScriptType == 1)
	{
		path = m_NamePatternParams.m_MSEdgeBrowserPath;
		if (!FileUtils::PathFileExist(path))
		{
			errorText = L"Path to MSEdge Browser not valid.\nPlease make sure Microsoft Edge is installed and path is correct."
				"\nOr configure Google Chrome Browser for printing."
				"\nSelect File->Print Config to update the setup.";
			return -1;
		}
	}
	else
	{
		path = m_NamePatternParams.m_UserDefinedScriptPath;
		if (!FileUtils::PathFileExist(path))
		{
			errorText = L"Path to user defined HTML2PDF script not valid.\nPlease make sure script is installed.\nSelect File->Print Config to update the setup.";
			return -1;
		}
	}
	return 1;
}

int CMainFrame::ExecCommand_WorkerThread(CString &htmFileName, CString &errorText, BOOL progressBar, CString &progressText, 
	int forcedTimeout, int headlessTimout)
{
	_ASSERTE(headlessTimout == -1);  // headlessTimout is broken in official Chrome and Edge

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	int timeout = forcedTimeout;  // set internal timout

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
		errorText.Append(L"Internal error. Try again.");
		return -1;
	}

	CString htmlFilePath = htmFileName;
	CString folderPath;
	FileUtils::GetFolderPath(htmFileName, folderPath);
	//FileUtils::GetFolderPathAndFileName(htmlFilePath, folderPath, htmFileName);

	CString pdfFileName;
	pdfFileName.Append(htmFileName);

	CString newSuffix;
	newSuffix.Append(L".pdf");
	FileUtils::UpdateFileExtension(pdfFileName, newSuffix);

	BOOL delStatus = FileUtils::DelFile(pdfFileName);
	if (delStatus == FALSE) {
		DWORD error = GetLastError();
		CString errorText = FileUtils::GetLastErrorAsString();
		TRACE(L"ExecCommand_WorkerThread: DeleteFile %s\nerror=%s\n", pdfFileName, errorText);
	}

	CString args;
	CString path;
	int bScriptType = pFrame->m_NamePatternParams.m_bScriptType;

	CString m_userAgentString;
	if (CMainFrame::m_enableUserAgent)
		m_userAgentString = CMainFrame::m_userAgentString;


	// Edge and Chrome configuration is basically the same for the latest versions
	// except for --headless=new  option for Chrome. --headless=new doesn't work in official Chrome, works in Canary version
	if (bScriptType == 0)   // Chrome
	{
		if (pFrame->m_NamePatternParams.m_bHeaderAndFooter)
		{
			args = m_userAgentString + L" --headless=new --disable-gpu";
		}
		else
		{
			//args = L" --headless --disable-gpu --print-to-pdf-no-header";  // --print-to-pdf-no-header was deprecated
			//args = L" --headless --disable-gpu --no-pdf-header-footer";  //  not sure if both --print-to-pdf-no-header and --no-pdf-header-footer can be set just in case
			args = m_userAgentString + L" --headless=new --disable-gpu --print-to-pdf-no-header --no-pdf-header-footer";
		}
#if 0
		// Ignore headlessTimout for now. The --timeout and --virtual-time-budget don't seem to work as expected
#else
		if (headlessTimout > 0)
		{
			int browserTimeout = headlessTimout;
			timeout = headlessTimout + 3;  // ignore forcedTimeout, set new value larger than headlessTimout

			CString tm;
			tm.Format(L" --timeout=%d", browserTimeout * 1000);
			args += tm;
		}
#endif

		args += L" --print-to-pdf=\"" + pdfFileName + L"\" \"" + htmFileName + L"\"";

		path = pFrame->m_NamePatternParams.m_ChromeBrowserPath;
		if (!FileUtils::PathFileExist(path))
		{
			errorText = L"Path to Chrome Browser not valid.\nPlease make sure Chrome is installed and path is correct."
				"\nOr configure Microsoft Edge Browser for printing."
				"\nSelect File->Print Config to update setup.";
			int deb = 1;
			return -1;
		}
	}
	else if (bScriptType == 1)   // Edge
	{
		if (pFrame->m_NamePatternParams.m_bHeaderAndFooter)
		{
			args = m_userAgentString + L" --headless=new --disable-gpu";  // header & footer is printer by default
		}
		else
		{
			//args = L" --headless --disable-gpu --print-to-pdf-no-header";  // --print-to-pdf-no-header was deprecated
			//args = L" --headless --disable-gpu --no-pdf-header-footer"; //  not sure if both --print-to-pdf-no-header and --no-pdf-header-footer can be set just in case
			args = m_userAgentString + L" --headless=new --disable-gpu --print-to-pdf-no-header --no-pdf-header-footer";
		}

#if 0
		// Ignore headlessTimout for now. The --timeout and --virtual-time-budget don't seem to work as expected
#else
		if (headlessTimout > 0)
		{
			int browserTimeout = headlessTimout;
			timeout = headlessTimout + 3;   // ignore forcedTimeout, set new value larger than headlessTimout

			CString tm;
			tm.Format(L" --timeout=%d", browserTimeout * 1000);
			args += tm;
		}
#endif

		args += L" --print-to-pdf=\"" + pdfFileName + L"\" \"" + htmFileName + L"\"";

		path = pFrame->m_NamePatternParams.m_MSEdgeBrowserPath;
		if (!FileUtils::PathFileExist(path))
		{
			errorText = L"Path to MSEdge Browser not valid.\nPlease make sure MSEdge is installed and path is correct."
				"\nOr configure Google  Chrome Browser for printing."
				"\nSelect File->Print Config to update the setup.";
			int deb = 1;
			return -1;
		}
	}
	else
	{
		if (pFrame->m_NamePatternParams.m_bKeepMailBodyBackgroundColor)
			args = L"\"" + htmFileName + L"\"";
		else
			args = L"\"" + htmFileName + L"\"" + L" " + L"--no-background";
		path = pFrame->m_NamePatternParams.m_UserDefinedScriptPath;
		if (!FileUtils::PathFileExist(path)) {
			errorText = L"Path to user defined HTML2PDF script not valid.\nPlease make sure script is installed.\nSelect File->Print Config to update setup.";
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
	ShExecInfo.lpDirectory = folderPath;   // FIXME
	ShExecInfo.nShow = SW_HIDE;
	ShExecInfo.hInstApp = result;
	BOOL retval = ShellExecuteEx(&ShExecInfo);
	if (retval == FALSE)
	{
		DWORD err = GetLastError();
		int ret = CMainFrame::CheckShellExecuteResult(ShExecInfo.hInstApp, errorText);
		return -1;
	}

	_ASSERTE(ShExecInfo.hProcess);
	if (ShExecInfo.hProcess == 0)
	{
		CString errText(L"Invalid ShExecInfo.hProcess == 0");
		HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		int answer = MyMessageBox(h, errText, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
	}

	int step = 10;
	int stepCnt = 0;
	int nSeconds = 0;
	CString secondsBar;

	DWORD msec = 100;
	BOOL signaled = FALSE;
	BOOL failed = FALSE;
	int returnCode = 1;

	HANDLE ProcHandle[1];
	ProcHandle[0] = ShExecInfo.hProcess;
	DWORD nCount = 1;
	DWORD dwWakeMask = 0;
	DWORD dwFlags = 0;

	for (;;)
	{
		msec = 100;
		//DWORD ret = MsgWaitForMultipleObjectsEx(nCount, &ProcHandle[0], msec, dwWakeMask, dwFlags);  // 06/15/2025 Just for eval; seems to work also
		DWORD ret = WaitForSingleObject(ShExecInfo.hProcess, msec);
		switch (ret)
		{
		case WAIT_ABANDONED: {
			errorText = L"Conversion of HTML to PDF failed. WaitForSingleObject returns WAIT_ABANDONED";
			failed = TRUE;
			break;
		}
		case WAIT_OBJECT_0: {
			signaled = TRUE;
			break;
		}
		case WAIT_FAILED: {
			errorText = L"Conversion of HTML to PDF failed. WaitForSingleObject returns WAIT_FAILED";
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
					returnCode = 2;
					return returnCode;
				}
				if (progressBar)
				{
					if (stepCnt % 10 == 0)
					{
						nSeconds++;
						if (progressText.IsEmpty()) {
							MboxMail::pCUPDUPData->SetProgress(step);
						}
						else
						{
#if 0
							secondsBar.Format(L"%s  %d seconds", progressText, nSeconds);
#endif

							CString fmt = L"%s  %d seconds";
							ResHelper::TranslateString(fmt);
							secondsBar.Format(fmt, progressText, nSeconds);

							MboxMail::pCUPDUPData->SetProgress(secondsBar, step);
						}
						step += 10;
						if (step > 100)
							step = 10;
					}
					stepCnt++;
					if ((timeout >= 0) && (nSeconds > timeout))
					{
						errorText.Format(L"Timeout. Conversion of HTML to PDF didn't finish within %d seconds", timeout);
						TRACE(L"Internal %s\n\nArguments: %s\n", errorText, args);
						failed = TRUE;
						if (ShExecInfo.hProcess)
						{
							TerminateProcess(ShExecInfo.hProcess, IDCANCEL);
							CloseHandle(ShExecInfo.hProcess);
						}
						Sleep(1000);
						return -2;
					}
				}
				else
				{
					if (stepCnt % 10 == 0)
					{
						nSeconds++;
						if ((timeout >= 0) && (nSeconds > timeout))
						{
							errorText.Format(L"Timeout. Conversion of HTML to PDF didn't finish within %d seconds", timeout);
							TRACE(L"Internal %s\n\nArguments: %s\n", errorText, args);
							failed = TRUE;
							if (ShExecInfo.hProcess)
							{
								TerminateProcess(ShExecInfo.hProcess, IDCANCEL);
								CloseHandle(ShExecInfo.hProcess);
							}
							Sleep(1000);
							return -2;
						}
					}
					stepCnt++;
				}
			}
			break;
		}
		default: {
			errorText.Format(L"Conversion of HTML to PDF failed. WaitForSingleObject returns unknown code=%s", ret);
			failed = TRUE;
			break;
		}
		}
		if (signaled || failed)
		{
			// ??
			if (failed)
				returnCode = -1;
			break;
		}
	}
	if (ShExecInfo.hProcess)
		CloseHandle(ShExecInfo.hProcess);

	if (returnCode < 0)
	{
		// 08/15/2025 Delay return if pdf file doesnt exist
		// For some not understood reason ShellExecuteEx return immediately but pdf is created eventually
		// This happens very rarely
		// Total delay of 2000 msec may  work for smaller mostly text emails, maybe
		// In case of true problem, the delay is just waste of time, hopefully not noticed by users
		if (!FileUtils::PathFileExist(pdfFileName))
		{
			for (int i = 0; i <= 3; i++)
			{
				if (!FileUtils::PathFileExist(pdfFileName))
				{
					Sleep(500);
				}
			}
		}
	}

	return returnCode;
}

int CMainFrame::ExecCommand_WorkerThread(CString &directory, CString &cmd, CString &args, CString &errorText, BOOL progressBar, CString &progressText, int timeout)
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	// TODO: Duplicate check, done already in _Thread function calls 

	HINSTANCE result = S_OK;
	SHELLEXECUTEINFO ShExecInfo = { 0 };
	ShExecInfo.cbSize = sizeof(SHELLEXECUTEINFO);
	ShExecInfo.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NOASYNC | SEE_MASK_NO_CONSOLE;
	//ShExecInfo.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NOASYNC;
	ShExecInfo.hwnd = NULL;
	ShExecInfo.lpVerb = NULL;
	ShExecInfo.lpFile = cmd;
	ShExecInfo.lpParameters = args;
	ShExecInfo.lpDirectory = (LPCWSTR)directory;
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
#if 0
			DWORD exitCode;
			BOOL ret = GetExitCodeProcess(ShExecInfo.hProcess, &exitCode);
			if ((ret == TRUE) && (exitCode == STILL_ACTIVE))
				;
			else
			{
				signaled = TRUE;
					break;
			}
#else
			signaled = TRUE;
			break;
#endif
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
						else
						{
#if 0
							secondsBar.Format(L"%s  %d seconds", progressText, nSeconds);
#endif

							CString fmt = L"%s  %d seconds";
							ResHelper::TranslateString(fmt);
							secondsBar.Format(fmt, progressText, nSeconds);

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
	m_MessageLimitCharsString = L"32500";
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
	MY_USES_CONVERSION;
	CString AttachmentNamesSeparatorString = MYA2W(m_AttachmentNamesSeparatorString);

	CString section_csv = CString(sz_Software_mboxview) + L"\\PrintConfig\\CSV";

	CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_csv, L"CSV_attachmentNamesSeparator", AttachmentNamesSeparatorString);
}

void CSVFILE_CONFIG::LoadFromRegistry()
{
	SetDflts();

	BOOL retval;
	CString section_csv = CString(sz_Software_mboxview) + L"\\PrintConfig\\CSV";

	CString AttachmentNamesSeparatorString;
	if (retval = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_csv, L"CSV_attachmentNamesSeparator", AttachmentNamesSeparatorString))
	{
		m_AttachmentNamesSeparatorString = AttachmentNamesSeparatorString;
		;
	}
}

INT_PTR CMainFrame::SelectFolder(CString &folder, CString *path)
{
	// TODO: customize CFileDialog to avoid potential buffer overflow and corruption
#define MAX_CFileDialog_FCOUNT 10
	int FILE_LIST_BUFFER_SIZE = (MAX_CFileDialog_FCOUNT * (MAX_PATH + 1)) + 1;

	MboxMail::m_outbuf->ClearAndResize(FILE_LIST_BUFFER_SIZE);
	wchar_t *fileNameBuffer = (LPWSTR)MboxMail::m_outbuf->Data();
	fileNameBuffer[0] = 0;

	//CString  sFilters = "Mail Files (*.txt)|*.txt||";
	CString  sFilters = L"Mail Files (*.mbox;*eml)|*.mbox;*.eml||";

	DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
	MySelectFolder dlgFile(TRUE, NULL, NULL, dwFlags, sFilters);

	CString inFolderPath = MboxMail::GetLastPath();
	if (path)
		inFolderPath = *path;

	OPENFILENAME& ofn = dlgFile.GetOFN();
	//ofn.Flags |= OFN_ALLOWMULTISELECT;
	ofn.lpstrFile = fileNameBuffer;
	ofn.nMaxFile = FILE_LIST_BUFFER_SIZE;
	ofn.lpstrInitialDir = inFolderPath;
	ofn.lpstrTitle = L"Select Folder with Mail Archives";

	INT_PTR ret = dlgFile.DoModal();
	if (ret == IDOK)
	{
		folder = dlgFile.GetFolderPath();
		TRACE(L"FOLDER=%s\n", folder);

		int fileCount = 0;
		CString strFilePath;
		POSITION pos = dlgFile.GetStartPosition();
		while (NULL != pos)
		{
			fileCount++;
			strFilePath = dlgFile.GetNextPathName(pos);

			TRACE(L"FILE=%s\n", strFilePath);

			BOOL retval = FileUtils::CPathGetPath(strFilePath, folder);
			break;
		}
	}
	return ret;
}

MySelectFolder::MySelectFolder(
		BOOL bOpenFileDialog, LPCWSTR lpszDefExt, LPCWSTR lpszFileName,
		DWORD dwFlags, LPCWSTR lpszFilter, CWnd* pParentWnd, DWORD dwSize, BOOL bVistaStyle
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
	TRACE(L"FOLDER=%s\n", folder);

	int fileCount = 0;
	CString strFilePath;
	POSITION pos = GetStartPosition();
	while (NULL != pos)
	{
		fileCount++;
		strFilePath = GetNextPathName(pos);

		TRACE(L"FILE=%s\n", strFilePath);

		BOOL retval = FileUtils::CPathGetPath(strFilePath, folder);
		break;
	}

	CString extension = L"\\*.mbox";
	int fcnt = CMainFrame::CountMailFilesInFolder(folder, extension);
	if (fcnt == 0)
	{
		extension = L"\\*.eml";
		fcnt = CMainFrame::CountMailFilesInFolder(folder, extension);
	}
	if (fcnt > 0)
		return FALSE;
	else
		return TRUE;
}

void CMainFrame::SaveLastSelection()
{
	// TODO: Add your message handler code here and/or call default

	NTreeView* pTreeView = GetTreeView();
	NListView* pListView = GetListView();
	NMsgView* pMsgView = GetMsgView();

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
	CString section_lastSelection = CString(sz_Software_mboxview) + L"\\LastSelection";

	ret = CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_lastSelection, L"lastMailFilePath", lastMailFilePath);
	ret = CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_lastSelection, L"lastLabelFilePath", lastLabelFilePath);
	ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_lastSelection, L"lastMailIndex", lastMailIndex);
	ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_lastSelection, L"lastWhichSort", lastWhichSort);
}

void CMainFrame::OnClose()
{
	NTreeView* pTreeView = GetTreeView();
	NListView* pListView = GetListView();
	NMsgView* pMsgView = GetMsgView();

	SaveLastSelection();

	BOOL isDirty = MboxMail::m_editMails.m_bIsDirty;
	if ((MboxMail::IsUserMailsSelected() && (MboxMail::s_mails.GetCount() > 0)) || (MboxMail::s_mails_edit.GetCount() > 0))
	{
		if (isDirty)
		{
			CString txt = L"User Selected Mails List not empty. Terminate program?";
			ResHelper::TranslateString(txt);
			int answer = MyMessageBox(txt, L"Warning", MB_APPLMODAL | MB_ICONWARNING | MB_YESNO);
			if (answer == IDNO)
				return;
		}
	}
	else if (isDirty)
		int deb = 1;


	//if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode && !m_commandLineParms.m_bEmlPreviewFolderExisted)
	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
	{
		NTreeView* pTreeView = GetTreeView();
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

	CString section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacement";
#if 0
	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
		section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacementPreview";
	else if (CMainFrame::m_commandLineParms.m_bDirectFileOpenMode)
		section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacementDirect";
#endif

	WINDOWPLACEMENT wp;
	GetWindowPlacement(&wp);

	if (wp.showCmd == SW_MAXIMIZE)
		wp.showCmd = SW_SHOWMAXIMIZED;
	if ((wp.showCmd != SW_SHOWMAXIMIZED) && (wp.showCmd != SW_SHOW))
		wp.showCmd = SW_SHOW;

	CRect rectMsgView;
	CRect rectListView;
	CRect rectTreeView;
	CRect rectFrameView;


	NMsgView* msgView = GetMsgView();
	NListView* listView = GetListView();
	NTreeView* treeView = GetTreeView();

	msgView->GetClientRect(&rectMsgView);
	listView->GetClientRect(&rectListView);
	treeView->GetClientRect(&rectTreeView);

	CRect rectMsgViewWindow;
	CRect rectListViewWindow;
	CRect rectTreeViewWindow;

	msgView->GetWindowRect(rectMsgViewWindow);
	listView->GetWindowRect(rectListViewWindow);
	treeView->GetWindowRect(rectTreeViewWindow);

	this->GetClientRect(rectFrameView);

	{
		BOOL ret = CProfile::_WriteProfileBinary(HKEY_CURRENT_USER, section_wnd, L"MainFrame", (LPBYTE)&wp, sizeof(wp));

		int new_msgViewPosition = m_msgViewPosition;
		ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_wnd, L"messageWindowPosition", new_msgViewPosition);

		BOOL isHidden = IsTreeHidden();
		isHidden = FALSE;  // FIXMEFIXME

		if (m_msgViewPosition != new_msgViewPosition)
		{
			DeleteAllPlacementKeys(section_wnd);
		}
		else
		{
			BOOL ret;
			if (pMsgView)
			{
				if (!isHidden)
				{
					ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"MsgFrameTreeNotHiddenWidth", rectMsgView.Width());
					ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"MsgFrameTreeNotHiddenHeight", rectMsgView.Height());
				}
				else
				{

					ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"MsgFrameTreeNotHiddenWidth", pMsgView->m_frameCx_TreeNotInHide);
					ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"MsgFrameTreeNotHiddenHeight", pMsgView->m_frameCy_TreeNotInHide);
				}

				ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"MsgFrameTreeHiddenWidth", pMsgView->m_frameCx_TreeInHide);
				ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"MsgFrameTreeHiddenHeight", pMsgView->m_frameCy_TreeInHide);
			}

			if (pListView)
			{
				if (!isHidden)
				{
					ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"ListFrameTreeNotHiddenWidth", rectListView.Width());
					ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"ListFrameTreeNotHiddenHeight", rectListView.Height());
				}
				else
				{
					ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"ListFrameTreeNotHiddenWidth", pListView->m_frameCx_TreeNotInHide);
					ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"ListFrameTreeNotHiddenHeight", pListView->m_frameCy_TreeNotInHide);
				}

				ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"ListFrameTreeHiddenWidth", pListView->m_frameCx_TreeInHide);
				ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"ListFrameTreeHiddenHeight", pListView->m_frameCy_TreeInHide);
			}

			if (pTreeView)
			{
				if (!isHidden)
				{
					ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"TreeFrameWidth", rectTreeView.Width());
					ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"TreeFrameHeight", rectTreeView.Height());
				}
				else
				{
					ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"TreeFrameWidth", m_pTreeView->m_frameCx);
					ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_wnd, L"TreeFrameHeight", m_pTreeView->m_frameCy);
				}
			}
		}
	}

	CString ActiveMailService = m_mailDB.SMTPConfig.MailServiceName;;
	m_mailDB.Write2Registry(ActiveMailService);

	if (!CProfile::IsRegistryConfig())
	{
		ConfigTree* confTree = CProfile::GetConfigTree();
#ifdef _DEBUG
		CString label = L"CMainFrame::OnClose";
		confTree->DumpTree(label);
#endif
	}

	CMenu* menu = GetMenu();
	MyPopupMenu::RestoreCMenu(menu, 0);

	//MboxMail::ReleaseResources();
	//ResHelper::ReleaseResources();

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
			// TODO: review fint handling, looks suspect, seems to work
			if (setColor == FALSE) 
			{
				CClientDC dc(&m_wndStatusBar);
				CFont* pOldFont = dc.SelectObject(&CMainFrame::m_dfltStatusBarFont);
				m_wndStatusBar.GetStatusBarCtrl().SetText(sText, paneId, 0);
				if (pOldFont) CFont* pRetFornt = dc.SelectObject(pOldFont);
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

				COLORREF crRet = 0;
				if (setColor)
					crRet = dc.SetTextColor(RGB(0, 0, 200));

				// Set font
				CFont  *pOldFont = dc.SelectObject(&CMainFrame::m_dfltStatusBarFont);

				UINT nFormat = DT_EXTERNALLEADING | DT_SINGLELINE | DT_VCENTER;
				int retVal = dc.DrawText(sText, &r, nFormat);

				// Restore oldFont
				if (pOldFont) CFont *pRetFornt = dc.SelectObject(pOldFont);
				if (setColor)
					crRet = dc.SetTextColor(txtColo);

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
		//if (dlg->Create(IDD_COLOR_STYLE_DLG, this) == FALSE)
		if (dlg->CreateIndirect(IDD_COLOR_STYLE_DLG, this) == FALSE)
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
	MY_USES_CONVERSION;
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
			CStringA url = "about:blank";
			DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);

			{
				CStringA colorStr;
				int retC2A = NListView::Color2Str(color, colorStr);

				url = "about:<html><head><style>body{background-color: #";
				url.Append(colorStr);
				url.Append(";}</style></head><body></body></html><br>");
			}
			CString wurl = MYA2W(url);
			pMsgView->m_browser.Navigate(wurl, NULL);
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

	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	CString dataFolder = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"dataFolder");



	int actionCode = 0;
	if (dataFolder.IsEmpty())
	{
		if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode == FALSE)
		{
			DataFolderConfigDlg dlg;
			dlg.m_fontSize = CMainFrame::m_cnfFontSize;
			INT_PTR retCode = dlg.DoModal();
			if (retCode == IDOK)
			{
				actionCode = dlg.m_returnCode;
				int deb = 1;
			}
			else if (retCode == IDCANCEL)
			{
				int deb = 1;
			}
		}
		else
		{
			wchar_t* folder = L"UMBoxViewer";
			folder = 0;
			wchar_t* subfolder = L"MailPreview";
			subfolder = 0;

			CString configFolderPath = FileUtils::GetMboxviewLocalAppDataPath(folder, subfolder);
			if (!FileUtils::PathDirExists(configFolderPath))
			{
				int deb = 1;
			}

			CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_general, L"dataFolder", configFolderPath);
		}
	}

	dataFolder = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"dataFolder");
	if (dataFolder.IsEmpty())
	{
		CString txt = L"Data Folder is not configured. Exiting.\n\n"
			L"Please restart MBox Viewer and configure Data Folder by selecting \"File -> Data Folder Config\" option.";
		ResHelper::TranslateString(txt);

		HWND h = GetSafeHwnd();
		int answer = MyMessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		
		AfxGetMainWnd()->PostMessage(WM_CLOSE);
		return 0;
	}

	// Check if preview data folder needs to be deleted
	QWORD lastMailDataPreviewDeletTime = 0;
	BOOL retGet64 = CProfile::_GetProfileInt64(HKEY_CURRENT_USER, section_general, L"mailPreviewDataFolderDeleteTime", lastMailDataPreviewDeletTime);
	if (m_lastMailDataPreviewDeletTime == 0)
		m_lastMailDataPreviewDeletTime = lastMailDataPreviewDeletTime;

	QWORD nowTime = _time64(0);
	if ((m_lastMailDataPreviewDeletTime > lastMailDataPreviewDeletTime) || (nowTime > (m_lastMailDataPreviewDeletTime + (48 * 3600))))
	{
		CString previewDataFolder = dataFolder + L"UMBoxViewer\\PView";

		bool recursive = true;
		bool removeFolders = true;
		CString errorText;
		if (FileUtils::PathDirExists(previewDataFolder))
		{
			BOOL retRemove = FileUtils::RemoveDir(previewDataFolder, recursive, removeFolders, &errorText);
		}

		m_lastMailDataPreviewDeletTime = nowTime;
		BOOL retWrite64 = CProfile::_WriteProfileInt64(HKEY_CURRENT_USER, section_general, L"mailPreviewDataFolderDeleteTime", m_lastMailDataPreviewDeletTime);
	}

	if (!dataFolder.IsEmpty())
	{
		if (!::SetCurrentDirectory((LPCWSTR)dataFolder))
		{
			int deb = 1;
		}
	}

	if (actionCode == IDTRYAGAIN)
	{
		// dlete UMBoxViewer Folder under Data Folder
		CString folder = dataFolder + L"UMBoxViewer";
		bool recursive = true;
		bool removeFolders = false;
		BOOL retCode1 = FileUtils::RemoveDir(folder, recursive, removeFolders);
		removeFolders = true;
		BOOL retCode2 = FileUtils::RemoveDir(folder, recursive, removeFolders);
		const int deb = 1;
	}

	MboxMail::s_folderContext.m_rootDataFolderPathConfig = dataFolder;

	if (pTreeView)
	{
		if (!m_commandLineParms.m_bEmlPreviewMode && !m_commandLineParms.m_bDirectFileOpenMode)
			pTreeView->LoadFolders();
	}

	CString listFilePath = CMainFrame::m_commandLineParms.m_mboxListFilePath;
	CString mergeToFilePath = CMainFrame::m_commandLineParms.m_mergeToFilePath;

	if (!listFilePath.IsEmpty() && !mergeToFilePath.IsEmpty())
	{
		CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
		CString sText;
		int paneId = 0;

		CString fmt = L"Merging files listed in %s file ...";
		ResHelper::TranslateString(fmt);

		sText.Format(fmt, listFilePath);
		TRACE(L"%s\n", sText);
		if (pFrame)
			pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

		int ret = MergeMboxArchiveFiles(listFilePath, mergeToFilePath);
		if (ret < 0)
			int deb = 1;

		sText.Format(L"Ready");
		ResHelper::TranslateString(sText);
		TRACE(L"%s\n", sText);
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

		mboxFilePath.TrimRight(L"\\");
		mboxFilePath.Append(L"\\");

		CString mboxFileNameOrPath = m_commandLineParms.m_mboxFileNameOrPath;
		mboxFileNameOrPath.TrimRight(L"\\");
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

BOOL CMainFrame::DeleteAllPlacementKeys(CString & section_wnd)
{
	LSTATUS ret;

	// Delete each value or delete key recursively
	// ret = CProfile::_DeleteKey(HKEY_CURRENT_USER, section_wnd, L"WindowPlacement", TRUE);

	ret = CProfile::_DeleteValue(HKEY_CURRENT_USER, section_wnd, L"MainFrame");

	ret = CProfile::_DeleteValue(HKEY_CURRENT_USER, section_wnd, L"MsgFrameTreeNotHiddenWidth");
	ret = CProfile::_DeleteValue(HKEY_CURRENT_USER, section_wnd, L"MsgFrameTreeNotHiddenHeight");

	ret = CProfile::_DeleteValue(HKEY_CURRENT_USER, section_wnd, L"MsgFrameTreeHiddenWidth");
	ret = CProfile::_DeleteValue(HKEY_CURRENT_USER, section_wnd, L"MsgFrameTreeHiddenHeight");

	ret = CProfile::_DeleteValue(HKEY_CURRENT_USER, section_wnd, L"ListFrameTreeNotHiddenWidth");
	ret = CProfile::_DeleteValue(HKEY_CURRENT_USER, section_wnd, L"ListFrameTreeNotHiddenHeight");

	ret = CProfile::_DeleteValue(HKEY_CURRENT_USER, section_wnd, L"ListFrameTreeHiddenWidth");
	ret = CProfile::_DeleteValue(HKEY_CURRENT_USER, section_wnd, L"ListFrameTreeHiddenHeight");

	ret = CProfile::_DeleteValue(HKEY_CURRENT_USER, section_wnd, L"TreeFrameWidth");
	ret = CProfile::_DeleteValue(HKEY_CURRENT_USER, section_wnd, L"TreeFrameHeight");

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
	size_t allMailsSize = MboxMail::AllMailsSizeof(mailCount);

	wchar_t sizeStr_inKB[256];
	wchar_t sizeStr_inBytes[256];
	int sizeStrSize = 256;
	CString txt;
	CString tmp;

	//_int64 fileSize = allMailsSize;
	LPCWSTR fileSizeStr_inKB = StrFormatKBSize(allMailsSize, &sizeStr_inKB[0], sizeStrSize);
	if (!fileSizeStr_inKB)
		sizeStr_inKB[0] = 0;

	LPCWSTR fileSizeStr_inBytes = StrFormatByteSize64(allMailsSize, &sizeStr_inBytes[0], sizeStrSize);
	if (!fileSizeStr_inBytes)
		sizeStr_inBytes[0] = 0;

	txt.Empty();
	CString cstr;
	INT64 numb = allMailsSize;
	TextUtilsEx::Int2WstrWithCommas(numb, cstr);

	tmp.Format(L"%s  (%s) (%s)\n", cstr, sizeStr_inKB, sizeStr_inBytes);
	txt += tmp;

	CString text;
	text.Format(L"Allocated Memory: %s\n", txt);

	HWND h = GetSafeHwnd();
	int answer = MyMessageBox(h, text, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
	int deb = 1;

	CString bufferSizes;
	bufferSizes.Format(L"Static Buffer Sizes:\n"
		L"\tm_outbuf=%s\n\tm_inbuf=%s\n\tm_outdata=%s\n\tm_indata=%s\n\tm_workbuf=%s\n\tm_tmpbuf=%s\n"
		L"\tm_largebuf=%s\n\tm_largelocal1=%s\n\tm_largelocal2=%s\n\tm_largelocal3=%s\n\tm_smalllocal1=%s\n\tm_smalllocal2=%s\n\n",
		FileUtils::SizeToString(MboxMail::m_outbuf->Capacity()),
		FileUtils::SizeToString(MboxMail::m_inbuf->Capacity()),
		FileUtils::SizeToString(MboxMail::m_outdata->Capacity()),
		FileUtils::SizeToString(MboxMail::m_indata->Capacity()),
		FileUtils::SizeToString(MboxMail::m_workbuf->Capacity()),
		FileUtils::SizeToString(MboxMail::m_tmpbuf->Capacity()),
		FileUtils::SizeToString(MboxMail::m_largebuf->Capacity()),
		FileUtils::SizeToString(MboxMail::m_largelocal1->Capacity()),
		FileUtils::SizeToString(MboxMail::m_largelocal2->Capacity()),
		FileUtils::SizeToString(MboxMail::m_largelocal3->Capacity()),
		FileUtils::SizeToString(MboxMail::m_smalllocal1->Capacity()),
		FileUtils::SizeToString(MboxMail::m_smalllocal2->Capacity())
	);
	TRACE(L"%s", bufferSizes);
	const int deb2 = 1;

#if 1
#ifdef _DEBUG
	MboxMail::m_textEncodingStats.Clear();
	MboxMail::CollectTextEncodingStats(MboxMail::m_textEncodingStats);

	CString stats2Text;
	MboxMail::m_textEncodingStats.PrintTextEncodingStats(MboxMail::m_textEncodingStats, stats2Text);
	TRACE(L"%s", stats2Text);

	const int deb3 = 1;
#endif
#endif

}

// View Raw message headers in Mail Message Pane
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
			// Show Raw Message Header , not a full mail message
			pMsgView->ShowMailHeader(iItem);
		}
		else
		{
			// Show full Mail Message
			pListView->Invalidate();
			pListView->SelectItem(pListView->m_lastSel, TRUE);
		}
	}

	CMenu *menu = this->GetMenu();
	UINT retval = 0;
	if (m_bViewMessageHeaders)
		retval = menu->CheckMenuItem(ID_VIEW_MESSAGEHEADERS, MF_CHECKED);
	else
		retval = menu->CheckMenuItem(ID_VIEW_MESSAGEHEADERS, MF_UNCHECKED);

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
		UINT retval = 0;
		retval = menu->CheckMenuItem(ID_MESSAGEHEADERPANELAYOUT_DEFAULT, MF_CHECKED);
		retval = menu->CheckMenuItem(ID_MESSAGEHEADERPANELAYOUT_EXPANDED, MF_UNCHECKED);
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
		UINT retval = 0;
		retval = menu->CheckMenuItem(ID_MESSAGEHEADERPANELAYOUT_DEFAULT, MF_UNCHECKED);
		retval = menu->CheckMenuItem(ID_MESSAGEHEADERPANELAYOUT_EXPANDED, MF_CHECKED);
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
		CString txt = L"Invalid Command Line \n\n\"";
		CString opt = m_allCommanLineOptions;
		txt += opt
			+ L"\""
			+ L"\n\nValid option combinations are:\n\n"
			+ L"  -FOLDER=Folder Path to open\n"
			+ L"  \nor\n\n"
			+ L"  -MAIL_FILE=Mbox File Path to open\n"
			+ L"  \nor\n\n"
			+ L"  -MAIL_FILE=Mbox File Path to open\n"
			+ L"  -EML_PREVIEW_MODE\n"
			+ L"  \nor\n\n"
			+ L"  -MBOX_MERGE_LIST_FILE=Path to File containing list of mbox files to merge\n"
			+ L"  -MBOX_MERGE_TO_FILE=Path to File to save merge results\n"
			//+ L"\nDo you want to continue?"
			;

		HWND h = NULL; // we don't have any window yet
		int answer = MyMessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
		return -1;
#if 0
		if (answer == IDNO)
			return -1;
		else
			return -2;
#endif
	}

	// cmd: mboxview -MAIL_FILE=m_mboxFileNameOrPath
	if (!m_mboxFileNameOrPath.IsEmpty() && !m_bEmlPreviewMode)
	{
		if (!FileUtils::PathFileExist(m_mboxFileNameOrPath))
		{
#if 1  // FIXMEFIXME was commented  out why ??
#if 0
			CString txt;
			txt.Format(L"Invalid -MAIL_FILE=\"%s\" option.\n"
				"No such File or Directory\n\n\"%s\"\n\n",
				m_mboxFileNameOrPath, m_mboxFileNameOrPath);
			txt += L"Do you want to continue?";
#endif

			CString txt;
			CString fmt = L"Invalid -MAIL_FILE=\"%s\" option.\nNo such File or Directory\n\n\"%s\"\n\nDo you want to continue ?";
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, m_mboxFileNameOrPath, m_mboxFileNameOrPath);

			HWND h = NULL; // we don't have any window yet  
			const int answer = MyMessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
			if (answer == IDNO)
				return -1;
			else
				return -2;

#else
			CString txt;
			txt.Format(L"Invalid -MAIL_FILE=\"%s\" option.\n"
				"No such File or Directory\n\n\"%s\"\n\n",
				m_mboxFileNameOrPath, m_mboxFileNameOrPath);
			HWND h = NULL; // we don't have any window yet  
			const int answer = MyMessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);

			return -1;
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

		folderPath.TrimRight(L"\\");
		folderPath.Append(L"\\");

		FileUtils::SplitFilePath(folderPath, driveName, directory, fileNameBase, fileNameExtention);
		if (directory.GetLength() <= 1)
		{
#if 0
			CString txt;
			txt.Format(L"Invalid -MAIL_FILE=\"%s\" option.\nInvalid folder \"%s\" .\n"
				"The mbox files must be installed under a named folder.\n"
				"Please create folder, move the mbox files to that folder and try again.",
				m_mboxFileNameOrPath, mboxFilePath);
#endif
			CString txt;
			CString fmt = L"Invalid -MAIL_FILE=\"%s\" option.\nInvalid folder \"%s\" .\n"
				L"The mbox files must be installed under a named folder.\n"
				L"Please create folder, move the mbox files to that folder and try again.";

			ResHelper::TranslateString(fmt);
			txt.Format(fmt, m_mboxFileNameOrPath, mboxFilePath);

			HWND h = NULL; // we don't have any window yet  
			int answer = MyMessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONWARNING | MB_OK);
			return -1;
		}

		// LoadFolders() needs "lastPath" to open folder housing "mailFile"
		// MboxMail::SetLastPath(folderPath);  // FIXMEFIXME added explicit _WriteProfileString

		CString section_general = CString(sz_Software_mboxview) + L"\\General";
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_general, L"mailFile", m_mboxFileNameOrPath);

		CString section_lastSelection = CString(sz_Software_mboxview) + L"\\LastSelection";
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_lastSelection, L"lastPath", folderPath);
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_lastSelection, L"lastMailFilePath", m_mboxFileNameOrPath);
	}

	// Cmd: mboxview -FOLDER=m_mboxFolderPath
	if (!m_mboxFolderPath.IsEmpty())
	{
		CString driveName;
		CString directory;
		CString fileNameBase;
		CString fileNameExtention;

		CString folderPath = m_mboxFolderPath;
		folderPath.TrimRight(L"\\");
		folderPath.Append(L"\\");

		FileUtils::SplitFilePath(folderPath, driveName, directory, fileNameBase, fileNameExtention);
		if (directory.GetLength() <= 1)
		{
			CString txt;
			folderPath.TrimRight(L"\\");
#if 0
			txt.Format(L"Invalid -FOLDER=\"%s\" option.\n\nInvalid folder \"%s\" .\n"
				"The mbox files must be installed under a named folder\n."
				"Please create folder, move the mbox files to that folder and try again.",
				folderPath, folderPath);
#endif
			CString fmt = L"Invalid -FOLDER=\"%s\" option.\n\nInvalid folder \"%s\" .\n"
				L"The mbox files must be installed under a named folder\n."
				L"Please create folder, move the mbox files to that folder and try again.";
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, folderPath, folderPath);

			HWND h = NULL; // we don't have any window yet  
			int answer = MyMessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
			return -1;
		}

		BOOL ret = FileUtils::PathDirExists(folderPath);
		if (ret == FALSE)
		{
#if 0
			CString txt;
			txt.Format(L"Invalid -FOLDER=\"%s\" option.\n\n"
				"No such Folder \"%s\" .\n",
				folderPath, folderPath);
#endif

			CString txt;
			CString fmt = L"Invalid -FOLDER=\"%s\" option.\n\nNo such Folder \"%s\" .\n";
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, folderPath, folderPath);

			//txt += L"Do you want to continue?";
			HWND h = NULL; // we don't have any window yet  
			int answer = MyMessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
			return -1;
#if 0
			if (answer == IDNO)
				return -1;
			else
				return -2;
#endif
		}
		//MboxMail::SetLastPath(folderPath);
		CString section_lastSelection = CString(sz_Software_mboxview) + L"\\LastSelection";
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_lastSelection, L"lastPath", folderPath);
	}

	if (!m_mboxListFilePath.IsEmpty() && !m_mergeToFilePath.IsEmpty())
	{
		if (!FileUtils::PathFileExist(m_mboxListFilePath))
		{
#if 0
			CString txt;
			txt.Format(L"Invalid -MBOX_MERGE_LIST_FILE=\"%s\" option.\n"
				"No such File:  \"%s\"\n\n",
				m_mboxListFilePath, m_mboxListFilePath);
#endif

			CString txt;
			CString fmt = L"Invalid -MBOX_MERGE_LIST_FILE=\"%s\" option.\nNo such File:  \"%s\"\n\n";
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, m_mboxListFilePath, m_mboxListFilePath);

			//txt += L"Do you want to continue?";
			HWND h = NULL; // we don't have any window yet  
			int answer = MyMessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
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
		if (fileNameExtention.CompareNoCase(L".mbox"))
		{
#if 0
			CString txt;
			txt.Format(L"Invalid -MBOX_MERGE_TO_FILE=\"%s\" option.\nInvalid file extension \"%s\" .\n"
				"File extension must be set to .mbox.\n",
				m_mergeToFilePath, fileNameExtention);
#endif

			CString txt;
			CString fmt = L"Invalid -MBOX_MERGE_TO_FILE=\"%s\" option.\nInvalid file extension \"%s\" .\n"
				"File extension must be set to .mbox.\n";
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, m_mergeToFilePath, fileNameExtention);


			HWND h = NULL; // we don't have any window yet  
			int answer = MyMessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONWARNING | MB_OK);
			return -1;
		}

		if (directory.GetLength() <= 1)
		{
#if 0
			CString txt;
			txt.Format(L"Invalid -MBOX_MERGE_TO_FILE=\"%s\" option.\nInvalid folder \"%s\" .\n"
				"The mbox files must be installed under a named folder.\n"
				"Please create folder, move the mbox files to that folder and try again.",
				m_mergeToFilePath, mboxFolderPath);
#endif

			CString txt;
			CString fmt = L"Invalid -MBOX_MERGE_TO_FILE=\"%s\" option.\nInvalid folder \"%s\" .\n"
				"The mbox files must be installed under a named folder.\n"
				"Please create folder, move the mbox files to that folder and try again.";
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, m_mergeToFilePath, mboxFolderPath);

			HWND h = NULL; // we don't have any window yet  
			int answer = MyMessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONWARNING | MB_OK);
			return -1;
		}

		if (!CMainFrame::CanMboxBeSavedInFolder(mboxFolderPath))   // FIXMEFIXME
		{
#if 0
			CString txt = L"The \"MBoxViewer\" name is reserved and it can't appear in the file path.\n"
				"Please create different name and try again.";
			HWND h = NULL;
			int answer = MyMessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
#endif
			return -1;
		}

		BOOL ret = FileUtils::PathDirExists(mboxFolderPath);
		if (ret == FALSE)
		{
			CString txt;
			CString fmt = L"Invalid -MBOX_MERGE_TO_FILE=\"%s\" option.\n\nNo such folder: \"%s\".\n\nCommand line:\n\n\"%s\"\n\n";
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, m_mergeToFilePath, mboxFolderPath, m_allCommanLineOptions);

			//txt += L"Do you want to continue?";
			HWND h = NULL; // we don't have any window yet  
			int answer = MyMessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
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
#if 0
			CString txt;
			txt.Format(L"Possibly Invalid -MBOX_MERGE_TO_FILE=\"%s\" option.\n\n"
				"File Exists\"%s\"\n\n"
			"!!!  Make sure you are not overwritting original mbox file from mail service such as Gmail !!!\n\n"
			"It is always good idea to have backup of original mbox mail files\n\n",
				m_mergeToFilePath, m_mergeToFilePath);
			txt += L"Override and continue?";
#endif

			CString txt;
			CString fmt = L"Possibly Invalid -MBOX_MERGE_TO_FILE=\"%s\" option.\n\n"
				L"File Exists\"%s\"\n\n"
				L"!!!  Make sure you are not overwritting original mbox file from mail service such as Gmail !!!\n\n"
				L"It is always good idea to have backup of original mbox mail files\n\n"
				L"Override and continue?";

			ResHelper::TranslateString(fmt);
			txt.Format(fmt, m_mergeToFilePath, m_mergeToFilePath);

			HWND h = NULL; // we don't have any window yet  
			int answer = MyMessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
			if (answer == IDNO)
				return -1;
			else
			{
				//FileUtils::DelFile(m_mergeToFilePath);
				return 1;
			}
		}
	}
	else if (!m_mboxListFilePath.IsEmpty() && m_mergeToFilePath.IsEmpty())
	{
		CString txt;
		CString fmt = L"Missing or Invalid -MBOX_MERGE_TO_FILE option.\nCommand line:\n\n\"%s\"\n\n";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, m_allCommanLineOptions);

		//txt += L"\nDo you want to continue?";
		HWND h = NULL; // we don't have any window yet  
		int answer = MyMessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
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
		CString fmt = L"Missing or Invalid -MBOX_MERGE_LIST_FILE option.\nCommand line:\n\n\"%s\"\n\n";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, m_allCommanLineOptions);

		//txt += L"\nDo you want to continue?";
		HWND h = NULL; // we don't have any window yet  
		int answer = MyMessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
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
	if ((destinationFolder.Find(L"\\UMBoxViewer\\") >= 0) || (destinationFolder.Find(L"\\UMBoxViewer") >= 0))
	{
#if 0
		CString txt;
		txt.Format(L"Invalid -MBOX_MERGE_TO_FILE=\"%s\" option.\n\nInvalid folder \"%s\" .\n",
			destinationFolder, destinationFolder);
		txt.Append(L"\nDestination folder can't contain \"UMBoxViewer\" folder  in the path \n\n");
#endif
		CString txt;
		CString fmt = L"Invalid -MBOX_MERGE_TO_FILE=\"%s\" option.\n\nInvalid folder \"%s\" .\n"
			L"\nDestination folder can't contain \"UMBoxViewer\" folder  in the path \n\n";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, destinationFolder, destinationFolder);

		HWND h = NULL; // we don't have any window yet
		int answer = MyMessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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

	CString appDataPath = CMainFrame::GetMboxviewLocalAppDataPath(L"MailService");
	CString smtpConfigFilePath = appDataPath + "SMTP.ini";

	CreateMailDbFile(m_mailDB, smtpConfigFilePath);

	int deb = 1;
}


void CMainFrame::OnHelpUserguide()
{
	// TODO: Add your command handler code here


	CString helpFileName = L"UserGuide.pdf";
	HWND h = GetSafeHwnd();
	CString languageName;
	BOOL ignoreLanguage = ResHelper::IsEnglishConfigured(languageName);
	OpenHelpFile(helpFileName, h, ignoreLanguage);
}


void CMainFrame::OnHelpReadme()
{
	// TODO: Add your command handler code here

	CString helpFileName = L"README.md.txt";
	HWND h = GetSafeHwnd();
	CString languageName;
	BOOL ignoreLanguage = TRUE;
	OpenHelpFile(helpFileName, h, ignoreLanguage);

	OpenTranslatedHelpFile(helpFileName, h);
}


void CMainFrame::OnHelpLicense()
{
	// TODO: Add your command handler code here

	CString helpFileName = L"LICENSE.txt";
	HWND h = GetSafeHwnd();
	CString languageName;
	BOOL ignoreLanguage = TRUE;
	OpenHelpFile(helpFileName, h, ignoreLanguage);

	OpenTranslatedHelpFile(helpFileName, h);
}

void CMainFrame::OpenHelpFile(CString &helpFileName, HWND h, BOOL ignoreLanguage)
{
	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	CString processPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"processPath");
	CString language = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"language");

	CString processDir;
	FileUtils::CPathGetPath(processPath, processDir);

	CString filePath;
	if (ignoreLanguage)
	{
		filePath = processDir + L"\\" + helpFileName;
	}
	else
	{
		if (language.IsEmpty())
			language = L"english"; //  handling of english language too complicated

		filePath = processDir + "\\HelpFiles\\" + language + L"\\" + helpFileName;
	}
#ifdef USE_STACK_WALKER
	if (!FileUtils::PathFileExist(filePath))
	{
		if (ignoreLanguage)
		{
			filePath = processDir + L"\\..\\" + helpFileName;
		}
		else
		{
			if (language.IsEmpty())
				language = L"english"; //  handling of english language too complicated

			filePath = processDir + "\\..\\HelpFiles\\" + language + L"\\" + helpFileName;
		}
	}
#endif

	if (FileUtils::PathFileExist(filePath))
	{
		ShellExecute(NULL, L"open", filePath, NULL, NULL, SW_SHOWNORMAL);
	}
	else
	{
		CString txt;
		CString fmt = L"Help file \n\n\"%s\"\n\ndoesn't exist";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, filePath);

		int answer = MyMessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
	}
	int deb = 1;
}

void CMainFrame::OpenTranslatedHelpFile(CString& helpFileName, HWND h)
{
	CString languageName;
	if (!ResHelper::IsEnglishConfigured(languageName))
	{
		CString section_general = CString(sz_Software_mboxview) + L"\\General";

		CString processPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"processPath");
		CString language = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"language");

		CString processDir;
		FileUtils::CPathGetPath(processPath, processDir);

		CString inputFile = processDir + L"\\" + helpFileName;
#ifdef USE_STACK_WALKER
		if (!FileUtils::PathFileExist(inputFile))
			inputFile = processDir + L"\\..\\" + helpFileName;
#endif

		CString HelpPath = CMainFrame::GetMboxviewTempPath(L"MboxHelp");

		BOOL createDirOk = TRUE;
		if (!FileUtils::PathDirExists(HelpPath))
			createDirOk = FileUtils::CreateDir(HelpPath);
		if (!createDirOk)
			return;  // Handle errors

		CString outputHtmlFile = HelpPath + L"\\" + helpFileName + L".html";
		CString targetLanguageCode = ResHelper::GetLanguageCode(languageName);

		int retval = HtmlUtils::CreateTranslationHtml(inputFile, targetLanguageCode, outputHtmlFile);
		if (retval >= 0)
			ShellExecute(NULL, L"open", outputHtmlFile, NULL, NULL, SW_SHOWNORMAL);
	}
}


// Convert to Unicode wchar ???  // FIXMEFIXME
BOOL CMainFrame::CreateMailDbFile(MailDB &m_mailDB, CString &fileName)
{
	MY_USES_CONVERSION;
	CFile fp;
	CFileException ExError;
	if (!fp.Open(fileName, CFile::modeWrite | CFile::modeCreate, &ExError))
	{
		DWORD lastErr = ::GetLastError();
#if 1
		HWND h = GetSafeHwnd();
		CString fmt = L"Could not create file:\n\n\"%s\"\n\n%s";  // new format
		CString errorText = FileUtils::ProcessCFileFailure(fmt, fileName, ExError, lastErr, h);

#else
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError)
		CString txt = L"Could not create \"" + fileName;
		txt += L"\" file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);
#endif

		return FALSE;
	}

	// Convert to wchar ???  // FIXMEFIXME
	CStringA section = "[MailService]\r\n";
	fp.Write(section, section.GetLength());

	CStringA ActiveMailServiceA = MYW2A(m_mailDB.ActiveMailService);
	CStringA fld = "ActiveMailService=" + ActiveMailServiceA + "\r\n";
	fp.Write(fld, fld.GetLength());

	WriteMTPServerConfig(m_mailDB.GmailSMTPConfig, fp);
	WriteMTPServerConfig(m_mailDB.YahooSMTPConfig, fp);
	WriteMTPServerConfig(m_mailDB.OutlookSMTPConfig, fp);
	WriteMTPServerConfig(m_mailDB.CustomSMTPConfig, fp);

	return TRUE;
}

// Convert to Unicode wchar ???  // FIXMEFIXME
BOOL CMainFrame::WriteMTPServerConfig(MailConfig &serverConfigW, CFile &fp)
{
	MailConfigA serverConfig;
	serverConfigW.WStr2Ansi(serverConfig);

	CStringA CR = "\r\n";
	CStringA section = "[" + serverConfig.MailServiceName + "]" + CR;
	fp.Write(section, section.GetLength());
	CStringA fld = "MailServiceName=" + serverConfig.MailServiceName + CR;
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


// Top level handler to process root and subfolder
// if treeType == 2, all mbox files discoverd will be merged.
// Optionally, label will be inserted to each mail corresponding to mbox file name
// If label is inserted, label based tree view will be created, otherwise flat mail view is created
// If treeType == 1, the folder based tree is created and mbox files are not merged.
// Merge option creates a single file so global search acress all mbox files is supported.
//

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
		//CBrowseForFolder bff(GetSafeHwnd(), CSIDL_DESKTOP, IDS_SELECT_FOLDER);
		CBrowseForFolder bff(GetSafeHwnd(), CSIDL_DESKTOP, NULL);
		if (!path.IsEmpty())
			bff.SetDefaultFolder(path);

		CString title(L"Select Root Folder");
		ResHelper::TranslateString(title);
		bff.SetTitle(title);

		bff.SetFlags(BIF_RETURNONLYFSDIRS);
		if (bff.SelectFolder())
		{
			path = bff.GetSelectedFolder();
			path.TrimRight(L"\\");
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
			path.TrimRight(L"\\");
		}
		else
		{
			int deb = 1;
			return -1;
		}
	}

	if (!FileUtils::PathDirExists(path))
	{
		CString errorText = FileUtils::GetLastErrorAsString();
#if 0
		CString txt;
		txt.Format(L"Selected Folder\n\n%s\n\ndoesn't exist. Error: \n\n%s\n\n" ,
			path, errorText);
#endif
		CString txt;
		CString fmt = L"Selected Folder\n\n%s\n\ndoesn't exist. Error: \n\n%s\n\n";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, path, errorText);

		int answer = MyMessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	int maxFilesToMergeLimitWarnings = 300;
	if (treeType == 2)  // merge all mbox files and optinally create label based tree view
	{
		// m_mergeRootFolderStyle == 0 merge all mbox files under root folder only
		// m_mergeRootFolderStyle == 1 merge all mbox files under root folder and all subfolders
		//
		// m_labelAssignmentStyle == 0 no labels
		// m_labelAssignmentStyle == 1 assign mbox file name as label to each mail
		// m_labelAssignmentStyle == 2 assign folder name as label to each mail in that folder
		//

		if ((m_mergeRootFolderStyle == 1) && (m_labelAssignmentStyle == 1))  // labels per mbox file
		{
			BOOL recursive = TRUE;

			int fileCnt = FileUtils::GetFolderFileCount(path, recursive);
			if (fileCnt > maxFilesToMergeLimitWarnings)
			{
#if 0
				CString txt;
				txt.Format(L"Found [%d] files under the root folder and sub-folders.\n\n"
					"Content of files was not pre-examined to determine if all files are of the mbox type.\n\n"
					"In case all/most of files found are of the mbox type,"
					" this will create very large number of labels under Mail Tree which could make viewing of mails very challenging.\n\n"
					"Do you want to continue, cancel or retry Merge Configuration?"
				, fileCnt);
#endif

				CString txt;
				CString fmt = L"Found [%d] files under the root folder and sub-folders.\n\n"
					L"Content of files was not pre-examined to determine if all files are of the mbox type.\n\n"
					L"In case all/most of files found are of the mbox type,"
					L" this will create very large number of labels under Mail Tree which could make viewing of mails very challenging.\n\n"
					L"Do you want to continue, cancel or retry Merge Configuration?"
					;

				ResHelper::TranslateString(fmt);
				txt.Format(fmt, fileCnt);

				int answer = MyMessageBox(txt, L"Warning", MB_APPLMODAL | MB_ICONQUESTION | MB_CANCELTRYCONTINUE);
				if (answer == IDCANCEL)
					return -1;
				else if (answer == IDTRYAGAIN)
					return -2;
				// else IDCONTINUE
			}
		}
		// else if ((m_mergeRootFolderStyle == 1) &&	// merge all mbox files under root folder and all subfolders
		//    (    (m_labelAssignmentStyle == 0)        // no labels
		//      || (m_labelAssignmentStyle == 2))       // assign folder name as label for each mail in that folder
	}
	else if (treeType == 1)  // create folder based tree view
	{
		BOOL recursive = TRUE;

		int fileCnt = FileUtils::GetFolderFileCount(path, recursive);
		if (fileCnt > maxFilesToMergeLimitWarnings)
		{
#if 0
			CString txt;
			txt.Format(L"Found [%d] files under the root folder and sub-folders.\n\n"
				"Content of files was not pre-examined to determine if all files are of the mbox type.\n\n"
				"In case all/most of files found are of the mbox type,"
				" this will create very large number of folders under Mail Tree which could make viewing of mails challenging.\n\n"
				"Do you want to continue, cancel or retry Select root folder dialog?"
			, fileCnt);
#endif

			CString txt;
			CString fmt = L"Found [%d] files under the root folder and sub-folders.\n\n"
				"Content of files was not pre-examined to determine if all files are of the mbox type.\n\n"
				"In case all/most of files found are of the mbox type,"
				" this will create very large number of folders under Mail Tree which could make viewing of mails challenging.\n\n"
				"Do you want to continue, cancel or retry Select root folder dialog?";
				;
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, fileCnt);

			int answer = MyMessageBox(txt, L"Warning", MB_APPLMODAL | MB_ICONQUESTION | MB_CANCELTRYCONTINUE);
			if (answer == IDCANCEL)
				return -1;
			else if (answer == IDTRYAGAIN)
				return -2;
			// else IDCONTINUE
		}
	}

	path.Append(L"\\");

	FileUtils::SplitFilePath(path, driveName, directory, fileNameBase, fileNameExtention);
	if (directory.GetLength() <= 1)
	{
		CString txt = L"The mbox files must be installed under a named folder\n."
			"Please create folder, move the mbox files to that folder and try again.";
		ResHelper::TranslateString(txt);

		int answer = MyMessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
		return -1;
	}

	path.TrimRight(L"\\");
	//AfxGetApp()->AddToRecentFileList(path);

	path.Append(L"\\");
	MboxMail::SetLastPath(path);

	if (treeType == 1)
		OpenRootFolderAndSubfolders(path);   // create folder based tree view
	else if (treeType == 2)
		OpenRootFolderAndSubfolders_LabelView(path);  //  create label or no label based tree view
	else
		return -1;

	return 1;
}

// Configure data folder for files created by MBox Viewer (such as PDF)
void CMainFrame::OnFileDatafolderconfig()
{
	// TODO: Add your command handler code here

	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	CString currentDataFolder = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"dataFolder");
	_ASSERTE(!currentDataFolder.IsEmpty());

	DataFolderConfigDlg dlg(TRUE);
	dlg.m_fontSize = CMainFrame::m_cnfFontSize;
	int actionCode = 0;
	INT_PTR retCode = dlg.DoModal();

	if (retCode == IDOK)
	{
		actionCode = dlg.m_returnCode;
		int deb = 1;
	}
	else if (retCode == IDCANCEL)
	{
		int deb = 1;
	}

	CString newDataFolder = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"dataFolder");

	if (currentDataFolder.Compare(newDataFolder))
	{
		NTreeView *pTreeView = GetTreeView();

		if (pTreeView)
		{
			if (actionCode == IDTRYAGAIN)
			{
				// dlete UMBoxViewer Folder under Data Folder
				CString folder = newDataFolder + L"UMBoxViewer";
				bool recursive = true;
				bool removeFolders = false;
				BOOL retCode1 = FileUtils::RemoveDir(folder, recursive, removeFolders);
				removeFolders = true;
				BOOL retCode2 = FileUtils::RemoveDir(folder, recursive, removeFolders);
				const int deb = 1;
			}

			_ASSERTE(!currentDataFolder.IsEmpty() && !newDataFolder.IsEmpty());
			// Keep it simple for now. Later inspect data under old Root Data Folder and decide what to to
			//pTreeView->DeleteOldDataFolder(currentDataFolder, newDataFolder);
			int deb = 1;
		}
		MboxMail::s_folderContext.m_rootDataFolderPathConfig = newDataFolder;
	}
}

//  create label or no label based tree view
void CMainFrame::OpenRootFolderAndSubfolders_LabelView(CString &path)
{
	NTreeView *pTreeView = GetTreeView();

	if (pTreeView)
	{
		BOOL selectFolder = TRUE;
		pTreeView->DoOpenRootFolderAndSubfolders_LabelView(path, selectFolder);
	}
}


// Root function for merging all mbox files under root folder MergeFIXME
void CMainFrame::OnFileMergerootfoldersub()
{
	// TODO: Add your command handler code here

	SaveLastSelection();
#if 0
	CString txt = L"MBox Viewer will traverse selected root folder and all sub-folders and merge all mbox files found into a single mbox file.\n\n"
		"MBox Viewer will assign a label to each folder and each mbox file and create Tree of Labels View\n\n"
		"Do you want to continue?\n";
	int answer = MyMessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_YESNO);
	if (answer == IDYES)
	{
		FileSelectrootfolder(2);
	}
#else
	int wantsToRetry = -2;
	MergeRootFolderAndSubfolders dlg;
	dlg.m_fontSize = CMainFrame::m_cnfFontSize;

	while (wantsToRetry == -2)
	{
		m_mergeSelectTargetFolder = FALSE;
		INT_PTR retCode = dlg.DoModal();
		if (retCode == IDOK)
		{
			// 0==merge mbox files discovered under root folder only;1==root and subfolders
			m_mergeRootFolderStyle = dlg.m_mergeRootFolderStyle;  
			// 0==no labels;1=labels per each mail same as mbox file name;2==same as folder name
			m_labelAssignmentStyle = dlg.m_labelAssignmentStyle; 

			m_mergeSelectTargetFolder = dlg.m_selectTargetFolder;

			// user may decide to return and try different options
			// wantsToRetry == -2, cancel == -1, continue == 1
			int treeType = 2;
			// create single folder based tree view
			// create label or no label based tree view
			wantsToRetry = FileSelectrootfolder(treeType);
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
	CString txt = L"MBox Viewer will traverse selected root folder and all sub-folders and create Tree of Folders View.\n\n"
		L"Do you want to continue?\n\n\n"
		L"Note that search arcoss all mails is not supported by this option. "
		L"Select \"File->Select root folder for merging\" option to enable search across all mails\n"
		;
#if 0
	int answer = MyMessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_YESNO);
	if (answer == IDYES)
	{
		FileSelectrootfolder(1);
	}
#else

	BOOL ret = ResHelper::TranslateString(txt);
	int answer = MyMessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_YESNO);
	if (answer == IDYES)
	{
		int wantsToRetry = -2;
		while (wantsToRetry == -2)
		{
			// user may decide to return and try different options
			// wantsToRetry == -2, cancel == -1, continue == 1
			int treeType = 1;   // create multiple sub-folder based tree view
			wantsToRetry = FileSelectrootfolder(treeType);
			int deb = 1;
		}
	}
#endif
}


void CMainFrame::OnDevelopmentoptionsDumprawdata()
{
	// TODO: Add your command handler code here
#define RWBUFLEN 10240

	NListView * pListView = GetListView();
	NMsgView * pMsgView = GetMsgView();

	if (pListView == 0)
	{
		return;
	}

	CString lastMailFilePath;
	int lastMailIndex = -1;
	int lastWhichSort = 1;
	int selMailIndex = -1;

	lastMailFilePath = pListView->m_path;
	selMailIndex = pListView->m_lastSel;

	int maxIndex = MboxMail::s_mails.GetCount() - 1;
	if ((selMailIndex < 0) || (selMailIndex > maxIndex))
	{
		CString txt = L"Please select mail archive and one of mails.";
		ResHelper::TranslateString(txt);
		int answer = MyMessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;
	}

	lastMailIndex = MboxMail::s_mails[selMailIndex]->m_index;

	int maxRefIndex = MboxMail::s_mails_ref.GetCount() - 1;
	if ((lastMailIndex < 0) || (lastMailIndex > maxRefIndex))
	{
		CString txt = L"Internal Error. Please re-select mail archive and one of mails.";
		ResHelper::TranslateString(txt);
		int answer = MyMessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;
	}

	MboxMail *m = MboxMail::s_mails_ref[lastMailIndex];

	DevelopmentCreateArchive dlg;
	dlg.m_sourceArchiveFile = lastMailFilePath;
	dlg.m_selectedMailSubject = m->m_subj;

	INT_PTR ret = dlg.DoModal();
	if (ret != IDOK)
	{
		return;
	}

	CString mboxFilePath = dlg.m_sourceArchiveFolder + dlg.m_createdArchiveName;

	BOOL delcode = FileUtils::DelFile(mboxFilePath);
	CString dataPath = MboxMail::GetLastDataPath();
	CString indexFilePath = dataPath + L"\\" + dlg.m_createdArchiveName + L".mboxview";
	BOOL retdel = FileUtils::DelFile(indexFilePath);

	int firstIndex = lastMailIndex;
	if (dlg.m_leadingMailCount < 0)
		firstIndex = 0;
	else
	{
		firstIndex = lastMailIndex - dlg.m_leadingMailCount;
		if (firstIndex < 0)
			firstIndex = 0;
	}

	int lastIndex = lastMailIndex;
	if (dlg.m_trailingMailCount < 0)
		lastIndex = maxRefIndex;
	else
	{
		lastIndex = lastMailIndex + dlg.m_trailingMailCount;
		if (lastIndex > maxRefIndex)
			lastIndex = maxRefIndex;
	}

	MboxMail *mFirst = MboxMail::s_mails_ref[firstIndex];
	MboxMail *mLast = MboxMail::s_mails_ref[lastIndex];

	_int64 firstStartOff = mFirst->m_startOff;
	_int64 lastStartOff = mLast->m_startOff;

	_int64 lastEndOff = mLast->m_startOff + mLast->m_length;

	CFile fpr;
	CString &fname = MboxMail::s_path;
	CFileException rExError;
	if (!fpr.Open(lastMailFilePath, CFile::modeRead | CFile::shareDenyWrite, &rExError))
	{
		DWORD lastErr = ::GetLastError();
#if 1
		HWND h = GetSafeHwnd();
		//HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		CString fmt = L"Could not open mail archive file:\n\n\"%s\"\n\n%s";  // new format
		CString errorText = FileUtils::ProcessCFileFailure(fmt, lastMailFilePath, rExError, lastErr, h); 
#else
		// TODO: critical failure
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(rExError);

		CString txt;
		CString fmt = L"Could not open \"%s\" mbox mail archive file.\n%s";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, lastMailFilePath, exErrorStr);

		TRACE(L"%s\n", txt);
		//errorText = txt;

		HWND h = NULL; // we don't have any window yet
		int answer = MyMessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
#endif

		return;
	}

	CFile fpw;
	CFileException wExError;
	if (!fpw.Open(mboxFilePath, CFile::modeWrite | CFile::modeCreate, &wExError))
	{
		DWORD lastErr = ::GetLastError();
#if 1
		HWND h = GetSafeHwnd();
		CString fmt = L"Could not create mail archive file:\n\n\"%s\"\n\n%s";  // new format
		CString errorText = FileUtils::ProcessCFileFailure(fmt, mboxFilePath, wExError, lastErr, h);
#else

		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(wExError);

		CString txt;
		CString fmt = L"Could not create \"%s\" mail archive file.\n%s";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, mboxFilePath, exErrorStr);

		//TRACE(L"%s\n", txt);

		HWND h = NULL; // we don't have any window yet
		int answer = MyMessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
#endif

		fpr.Close();
		return;
	}

	_int64 totalLength = lastEndOff - firstStartOff;

	char buff[RWBUFLEN];
	_int64 bytesLeft = totalLength;

	UINT bytes2Read;
	UINT bytes2Write;
	_int64 filePos = fpr.Seek(firstStartOff, SEEK_SET);
	while (bytesLeft > 0)
	{
		if (bytesLeft > RWBUFLEN)
			bytes2Read = RWBUFLEN;
		else
			bytes2Read = (UINT)bytesLeft;

		bytes2Write = fpr.Read(buff, bytes2Read);
		if (bytes2Write != bytes2Read)
			int deb = 1;

		// TODO:  Write() can create exception !!!!! 
		fpw.Write(buff, bytes2Write);
		fpw.Flush();

		bytesLeft -= bytes2Write;
	}

	fpr.Close();
	fpw.Close();
	int deb = 1;
}


void CMainFrame::OnDevelopmentoptionsDevelo()
{
	// TODO: Add your command handler code here

	MboxMail::developerMode = !MboxMail::developerMode;

	CMenu *menu = this->GetMenu();
	UINT retval = 0;
	if (MboxMail::developerMode)
		retval = menu->CheckMenuItem(ID_DEVELOPMENTOPTIONS_DEVELO, MF_CHECKED);
	else
		retval = menu->CheckMenuItem(ID_DEVELOPMENTOPTIONS_DEVELO, MF_UNCHECKED);

}

#include "PageCodeListDlg.h"

void CMainFrame::OnDeveloperOptionsAboutSystem()
{
	// TODO: Add your command handler code here
	CString aboutSystem;
	CString info;

	UINT codePage = GetACP();
	CString codePageInfo;
	CString codePageOEMInfo;

	DWORD dwFlags = 0;
	CPINFOEXA CPInfoEx;

	BOOL retval = GetCPInfoExA(codePage, dwFlags, &CPInfoEx);

	BOOL ret = TextUtilsEx::Id2LongInfo(codePage, codePageInfo);

	UINT codePageOEM = GetOEMCP();

	retval = GetCPInfoExA(codePageOEM, dwFlags, &CPInfoEx);

	ret = TextUtilsEx::Id2LongInfo(codePageOEM, codePageOEMInfo);

	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	CString dataFolder = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"dataFolder");


	info.Format(L"Windows Code Page:\n%d \"%s\"\n", codePage, codePageInfo);
	aboutSystem.Append(info);

	info.Format(L"\nOEM Code Page:\n%d \"%s\"\n", codePageOEM, codePageOEMInfo);
	aboutSystem.Append(info);

	info.Format(L"\nmboxview Process  Path:\n\"%s\"\n", m_processPath);
	aboutSystem.Append(info);

	if (CProfile::IsRegistryConfig())
	{
		CString regEntry(L"HKEY_CURRENT_USER\\SOFTWARE\\UMBoxViewer");
		info.Format(L"\nmboxview Configuration Store: Windows Registry\n", regEntry);
	}
	else
	{
		info.Format(L"\nmboxview Configuration Store: Configuration File:\n\"%s\"\n", CmboxviewApp::m_configFilePath);
	}
	aboutSystem.Append(info);

	dataFolder.Append(L"UMBoxViewer");
	info.Format(L"\nmboxview Data Folder:\n\"%s\"\n", dataFolder);
	aboutSystem.Append(info);

	info.Format(L"\nmboxview Startup  Directory:\n\"%s\"\n", m_startupPath);
	aboutSystem.Append(info);

	codePage = 20949;
	retval = GetCPInfoExA(codePage, dwFlags, &CPInfoEx);

	codePage = 1147;
	retval = GetCPInfoExA(codePage, dwFlags, &CPInfoEx);

	DWORD BUFSIZE = 1024;
	wchar_t Buffer[1024];

	//if (!::SetCurrentDirectory((LPCWSTR)m_processPath))
	{
		int deb = 1;
	}

	DWORD dwRet = ::GetCurrentDirectory(BUFSIZE, Buffer);
	if (dwRet > 0)
	{
		info.Format(L"\nmboxview Current  Directory:\n\"%s\"\n", Buffer);
		aboutSystem.Append(info);
	}
	aboutSystem.Append(L"\n");
	
	HWND h = GetSafeHwnd();
	int answer = MyMessageBox(aboutSystem, L"Info", MB_APPLMODAL | MB_OK);

	int deb = 1;
}

void CMainFrame::OnFileGeneraloptions()
{
	// TODO: Add your command handler code here

	GeneralOptionsDlg dlg;

	dlg.m_relaxedMboxFileValidation = m_relaxedMboxFileValidation;
	dlg.m_relativeInlineImageFilePath = m_relativeInlineImageFilePath;
	dlg.m_enableUserAgent = m_enableUserAgent;
	dlg.m_numberOfHTML2ToMerge = m_numberOfHTML2ToMerge;

	if (dlg.DoModal() == IDOK)
	{
		m_relaxedMboxFileValidation = dlg.m_relaxedMboxFileValidation;
		m_relativeInlineImageFilePath = dlg.m_relativeInlineImageFilePath;
		NListView::m_fullImgFilePath_Config = !m_relativeInlineImageFilePath;
		m_enableUserAgent = dlg.m_enableUserAgent;
		m_numberOfHTML2ToMerge = dlg.m_numberOfHTML2ToMerge;

		CString section_general = CString(sz_Software_mboxview) + L"\\General";
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_general, L"relaxedMboxFileValidation", (WORD)m_relaxedMboxFileValidation);

		// FIXME relativeInlineImageFilePath is get in NListWiew and set here -:(((
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_general, L"relativeInlineImageFilePath", (WORD)m_relativeInlineImageFilePath);

		CString section_merge = CString(sz_Software_mboxview) + L"\\PrintConfig\\Merge";

		CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_merge, L"numberOfHTML2Merge", m_numberOfHTML2ToMerge);
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_merge, L"enableUserAgent", m_enableUserAgent);

		const int deb = 1;
	}
}


CString CMainFrame::GetMboxviewTempPath(const wchar_t* name)
{
	CString tempFolder = L"UMBoxViewer";
	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode || CMainFrame::m_commandLineParms.m_bDirectFileOpenMode)
		tempFolder = L"UMBoxViewerPreview";
	if (name == 0)
	{
		if (m_mboxviewTempPath.IsEmpty())
		{
			m_mboxviewTempPath = FileUtils::GetMboxviewTempPath(tempFolder, name);
		}
		return m_mboxviewTempPath;
	}
	else
	{
		CString path = FileUtils::GetMboxviewTempPath(tempFolder, name);
		return path;
	}
}


CString CMainFrame::GetMboxviewLocalAppDataPath(const wchar_t* name)
{
	CString tempFolder = L"UMBoxViewer";
	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode || CMainFrame::m_commandLineParms.m_bDirectFileOpenMode)
		tempFolder = L"UMBoxViewerPreview";

	CString mboxviewLocalAppDataPath = FileUtils::GetMboxviewLocalAppDataPath(tempFolder, name);
	return mboxviewLocalAppDataPath;
}


CString CMainFrame::CreateTempFileName(const wchar_t *ext)
{
	CString tempFolder = L"UMBoxViewer";
	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode || CMainFrame::m_commandLineParms.m_bDirectFileOpenMode)
		tempFolder = L"UMBoxViewerPreview";
	CString fileName = FileUtils::CreateTempFileName(tempFolder, ext);
	return fileName;
}

CWnd * CMainFrame::SetWindowFocus(CWnd* wnd)
{
	_ASSERTE(wnd);
	if (wnd == 0)
		return 0;
	//CWnd* wnd1 = wnd->SetActiveWindow();
	BOOL ret = wnd->SetForegroundWindow();
	CWnd* wnd1 = wnd->SetActiveWindow();
	CWnd *wnd2 = wnd->SetFocus();

#ifdef _DEBUG
	if (wnd2 == 0)
	{
		CString errorText = FileUtils::GetLastErrorAsString();
		TRACE(L"SetWindowFocus: \"%s\"\n", errorText);
	}
#endif

	return wnd2;
}

// TODO: there must simpler way to accomplish focus restoration
CWnd* CMainFrame::RestoreWindowFocus()
{
	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (pFrame && !pFrame->m_createCompleted)
	{
		return 0;
	}

	if (pFrame)
	{
		NTreeView* m_pTreeView = pFrame->DetTreeView();
		NListView* m_pListView = pFrame->DetListView();
		NMsgView* m_pMsgView = pFrame->DetMsgView();

		CBrowser* m_browser = &m_pMsgView->m_browser;
		CWebBrowser2* m_browser2 = &m_pMsgView->m_browser.m_ie;

		CWheelTreeCtrl* m_tree = &m_pTreeView->m_tree;
		CWheelListCtrl* m_wheelList = m_pListView->GetListCtrl();

		//TRACE(L"OnActivateApp: ACTIVE\n");

		CWnd* old = GetFocus();

		if (CmboxviewApp::wndFocus == m_wheelList)
		{
			LRESULT lres = m_wheelList->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
			//CWnd *w = SetWindowFocus(m_wheelList);
			TRACE(L"PostMessage to CWheelListCtrl\n");
			int deb = 0;
		}
		else if (CmboxviewApp::wndFocus == m_pListView)
		{
			LRESULT lres = m_pListView->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
			TRACE(L"PostMessage to ListView\n");
			int deb = 0;
		}
		else if (CmboxviewApp::wndFocus == m_tree)
		{
			LRESULT lres = m_pTreeView->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
			TRACE(L"PostMessage to TreeView\n");
			int deb = 0;
		}
		else if (CmboxviewApp::wndFocus == m_browser)
		{
			LRESULT lres = m_browser->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
			TRACE(L"PostMessage to Browser\n");
			int deb = 0;
		}
		else if (CmboxviewApp::wndFocus == m_browser2)
		{
			LRESULT lres = m_browser->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
			TRACE(L"PostMessage to Browser2\n");
			int deb = 0;
		}
		else
		{
			CWnd* old = GetFocus();
			LRESULT lres = pFrame->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
			TRACE(L"PostMessage to Frame\n");
		}
	}
	return CmboxviewApp::wndFocus;
}

void CMainFrame::OnHelpShortcuts()
{
	// TODO: Add your command handler code here

	CString helpFileName = L"ShortcutsSummary.pdf";
	HWND h = GetSafeHwnd();
	CMainFrame::OpenHelpFile(helpFileName, h);
}


void CMainFrame::OnHelpChangeLog()
{
	// TODO: Add your command handler code here

	CString helpFileName = L"CHANGE_LOG.md.txt";
	HWND h = GetSafeHwnd();

	BOOL ignoreLanguage = TRUE;
	CMainFrame::OpenHelpFile(helpFileName, h, ignoreLanguage);

	OpenTranslatedHelpFile(helpFileName, h);
}

void CMainFrame::OnDevelopmentoptionsCodepageinstalled()
{
	// TODO: Add your command handler code here
	PageCodeListDlg dlg;
	INT_PTR retcode = dlg.DoModal();
}

void CMainFrame::OnDevelopmentoptionsSelectlanguage()
{
	// TODO: Add your command handler code here

	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	CString lastFolderName;
	CString lastLanguageFolderPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"languageFolderPath");
	if (!lastLanguageFolderPath.IsEmpty())
	{
#if 0
		lastFolderName = lastLanguageFolderPath;
		lastFolderName.TrimRight(L"\\");

		FileUtils::CPathStripPath((LPCWSTR)lastFolderName, lastFolderName);
#else
		lastLanguageFolderPath.TrimRight(L"\\");
		FileUtils::CPathStripPath((LPCWSTR)lastLanguageFolderPath, lastFolderName);
#endif
	}
	else
		lastFolderName = L"english";

	// TODO: make dlg memeber of CMainFrame
	CString languageFolder;
	SelectLanguageDlg dlg;
	dlg.m_lastFolderName = lastFolderName;

	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK) {
		languageFolder = dlg.m_LanguagetName;
		int deb = 1;
	}
	else if (nResponse == IDCANCEL)
	{
		int deb = 1;
	}

	if (languageFolder.IsEmpty())
		return;

	CString processFolderPath;
	ResHelper::GetProcessFolderPath(processFolderPath);
	
	CString languageFolderPath = processFolderPath + L"Language" + L"\\" + languageFolder;
	if (languageFolder.CompareNoCase(L"english") == 0)
	{
		languageFolderPath.Empty();
	}

	if (languageFolder.Compare(lastFolderName) != 0)
	{
#if 0
		CString txt;
		txt.Format(L"MBox Viewer will exit. Please restart MBox Viewer for new language \"%s\" to take effect\n", languageFolder);
#endif
		CString txt;
		CString fmt = L"MBox Viewer will exit. Please restart MBox Viewer for new language \"%s\" to take effect\n";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, languageFolder);

		HWND h = GetSafeHwnd();
		int answer = MyMessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONEXCLAMATION | MB_OK);

		if (languageFolderPath.IsEmpty())
			languageFolder.Empty();

		CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_general, L"language", languageFolder);
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_general, L"languageFolderPath", languageFolderPath);

		CString translationFileName = languageFolder + L".txt";

		CString languageTranslationFilePath = processFolderPath + L"Language\\" + languageFolder + L"\\" + translationFileName;
		//ResHelper::LoadLanguageMap(languageTranslationFilePath);

		AfxGetMainWnd()->PostMessage(WM_CLOSE);

		CMainFrame::OnFileRestorehintmessages();
	}
	else
	{
		int deb = 1;
	}
	int deb = 1;
}


void CMainFrame::OnLanguagetoolsCreateresourcefile()
{
	// TODO: Add your command handler code here
	int deb1 = 1;
	
#ifdef _DEBUG
	ResHelper::PrintResInfo();
	int deb = 1;
#endif
}


void CMainFrame::OnLanguagetoolsCreatetranslationfile()
{
	// TODO: Add your command handler code here
		// TODO: Add your command handler code here
#ifdef _DEBUG
	TRACE(L"BEGIN Create Language Translation File\n");
	ResHelper::CreateLanguageFile();
	int deb = 1;
	TRACE(L"END Create Language Translation File\n");
#endif
}


void CMainFrame::OnLanguagetoolsResorttranslationfile()
{
	// TODO: Add your command handler code here

	//ResHelper::ResortLanguageFile();
	ResHelper::RenumberLanguageFile();
	int deb = 1;
}


void CMainFrame::OnLanguagetoolsResourcefileproperty()
{
	// TODO: Add your command handler code here


	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	CString lastFolderName;
	CString folderName;
	CString fileName;
	CString filePath;
	CString lastLanguageFolderPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"languageFolderPath");
	if (!lastLanguageFolderPath.IsEmpty())
	{
		lastFolderName = lastLanguageFolderPath;
		lastFolderName.TrimRight(L"\\");


		FileUtils::CPathStripPath((LPCWSTR)lastFolderName, lastFolderName);
		filePath = lastLanguageFolderPath + L"\\" + lastFolderName + L".txt";
	}

	int  isFile_UTF8 = ResHelper::IsFileUTF8(filePath);

	SimpleMemoryFile memFile;
	BOOL retOpen = memFile.Open(filePath);

	BOOL isASCII = TRUE;
	BOOL isText_UTF8 = ResHelper::IsTextUTF8((const char*)memFile.m_buffer.Data(), memFile.m_buffer.Count(), isASCII);

	int slen = memFile.m_buffer.Count();
	ResHelper::TextEncoding BOM;
	memFile.m_buffer.Append('\0');
	memFile.m_buffer.Append('\0');
	memFile.m_buffer.Append('\0');
	memFile.m_buffer.Append('\0');
	memFile.m_buffer.Append('\0');
	ResHelper::TextEncoding nEncoding = ResHelper::GetTextEncoding(memFile.m_buffer.Data(), slen, BOM);
	if (nEncoding == ResHelper::TextEncoding::ASCII)
		int deb = 1;
	else 	if (nEncoding == ResHelper::TextEncoding::ANSI)
		int deb = 1;
	else 	if (nEncoding == ResHelper::TextEncoding::UTF8)
		int deb = 1;
	else 	if (nEncoding == ResHelper::TextEncoding::UTF16LE)
		int deb = 1;
	else 	if (nEncoding == ResHelper::TextEncoding::UTF16BE)
		int deb = 1;
	else 	if (nEncoding == ResHelper::TextEncoding::UTF32LE)
		int deb = 1;
	else 	if (nEncoding == ResHelper::TextEncoding::UTF32BE)
		int deb = 1;

	int nCP = ResHelper::FindCodePageFromFile(filePath);


	int deb = 1;
}


void CMainFrame::OnHelpFilebasedconfig()
{
	// TODO: Add your command handler code here

	CString helpFileName = L"FileConfigurationHelp.pdf";
	HWND h = GetSafeHwnd();
	CMainFrame::OpenHelpFile(helpFileName, h);
}


void CMainFrame::OnLanguagetoolsSplittranslationfile()
{
	// TODO: Add your command handler code here
	ResHelper::SplitTranslationFile();
}

void CMainFrame::OnLanguageSelectlanguange()
{
	// TODO: Add your command handler code here
	CMainFrame::OnDevelopmentoptionsSelectlanguage();
}


void CMainFrame::OnLanguageHelp()
{
	// TODO: Add your command handler code here
	CString helpFileName = L"LanguagesHelp.pdf";
	HWND h = GetSafeHwnd();
	CMainFrame::OpenHelpFile(helpFileName, h);
}


void CMainFrame::OnTestCfileopenfailure()
{
	// TODO: Add your command handler code here

	CString HelpPath = CMainFrame::GetMboxviewTempPath(L"MboxHelp");

	BOOL createDirOk = TRUE;
	if (!FileUtils::PathDirExists(HelpPath))
		createDirOk = FileUtils::CreateDir(HelpPath);

	CStdioFile file1;
	CFileException exList1;
	CString noShareFilePath = HelpPath + L"noShareTestFile";

	UINT nOpenFlags1 = CFile::modeCreate | CFile::modeRead | CFile::typeText | CFile::shareExclusive | CFile::typeUnicode;
	if (!file1.Open(noShareFilePath, nOpenFlags1, &exList1))
	{
		int deb = 1;
	}

	CStdioFile file2;
	CFileException exList2;
	//CString noShareFilePath = createDirOk + L"\\noShareTestFile";

	UINT nOpenFlags2 = CFile::modeCreate | CFile::modeRead | CFile::typeText | CFile::shareExclusive | CFile::typeUnicode;
	if (!file2.Open(noShareFilePath, nOpenFlags2, &exList2))
	{
		DWORD lastErr = ::GetLastError();
		HWND h = GetSafeHwnd();
		CString fmt = L"Could not create file:\n\n\"%s\"\n\n%s";
		CString errorText = FileUtils::ProcessCFileFailure(fmt, noShareFilePath, exList2, lastErr, h);

		file1.Close();
		BOOL retdel = FileUtils::DelFile(noShareFilePath);

		int deb = 1;
		return;
	}
	file2.Close();
	file1.Close();
	int deb = 1;
}

INT_PTR CMainFrame::SetTemplate(CDialog* dlg, UINT Id, CWnd* parent)
{
	CDialogTemplate dlt;

	// load dialog template
	if (!dlt.Load(MAKEINTRESOURCE(Id)))
		return 0;  // Or  IDCANCEL ???

	if (CMainFrame::m_cnfFontSize != CMainFrame::m_dfltFontSize)
		dlt.SetSystemFont(CMainFrame::m_cnfFontSize);

	// dlt.m_hTemplate will be release by CDialogTemplate destructor
#if 0
	// display dialog box
	BOOL ret = dlg->InitModalIndirect(dlt.m_hTemplate, parent);

	INT_PTR retModal = dlg->CDialog::DoModal();
#else
	// Examples om Web suggest to use GlobalLock
	// But it doesn't seem to be needed as per
	// https://devblogs.microsoft.com/oldnewthing/20240604-00/?p=109847
	// However, no harm done using GlobalLock occasionally
	// Hope GlobalLock is released even when crash
	// 
	// get pointer to the modified dialog template 

	LPDLGTEMPLATE pData = (LPDLGTEMPLATE)GlobalLock(dlt.m_hTemplate);
	if (!pData)
		return 0; // Or  IDCANCEL ???

	// display dialog box
	BOOL ret = dlg->InitModalIndirect(dlt.m_hTemplate, parent);

	INT_PTR retModal = dlg->CDialog::DoModal();

	// unlock memory object
	GlobalUnlock(dlt.m_hTemplate);
#endif

#if 0
	// Below works but overkill
	DWORD dwTemplateSize = dlt.m_dwTemplateSize;
	BYTE* templateData = new BYTE[dwTemplateSize];

	memcpy_s(templateData, (size_t)dwTemplateSize, dlt.m_hTemplate, (size_t)dwTemplateSize);

	// display dialog box
	BOOL ret = dlg->InitModalIndirect((HGLOBAL)templateData, parent);

	INT_PTR retModal = dlg->CDialog::DoModal();

	delete[] templateData;
#endif

	return retModal;

}


void CMainFrame::OnFileFontconfig()
{
	FontConfig dlg;

	CString section_font_config = CString(sz_Software_mboxview) + L"\\FontConfig";
	CString param = L"FontSize";
	int fontSize = 8;
	BOOL retval;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_font_config, param, fontSize))
	{
		;
	}

	dlg.m_fontSize = fontSize;

	// TODO: Add your command handler code here

	if (dlg.DoModal() == IDOK)
	{
		retval = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_font_config, param, dlg.m_fontSize);
		CMainFrame::m_cnfFontSize = dlg.m_fontSize;

		NListView* listView = GetListView();
		NTreeView* treeView = GetTreeView();
		NMsgView* msgView = GetMsgView();

		treeView->ResetFont();
		listView->ResetFont();

		listView->ResizeColumns();

		//treeView->m_tree.Invalidate();
		//listView->GetListCtrl()->Invalidate();

		MyPopupMenu::SetupFonts();
		CString fontName = CMainFrame::m_dfltFontName;
		LOGFONT logFont;
		CMainFrame::CreateTooltipFont(CMainFrame::m_dfltFont, fontName, logFont);

		msgView->RecreateMailHdrMenu();

		CMenu* menu = GetMenu();

		CString title = L"SetNewMenu ";
		int index = 0;
		//UINT mask = MyPopupMenu::MenuItemInfoMaskTypeAllSet;
		UINT mask = MyPopupMenu::MenuItemInfoMaskFTypeAllSet;
		BOOL retval = MyPopupMenu::TraceMenu(title, menu, index, mask);

		MyPopupMenu::RestoreCMenu(menu, 0);
		MyPopupMenu::SetCMenuAsCustom(menu);

		this->DrawMenuBar();

	}

	int deb = 1;
}

void CMainFrame::SetNewMenu()
{
#if 1
	MyPopupMenu* menu = (MyPopupMenu*)this->GetMenu();

	CString title = L"SetNewMenu ";
	int index = 0;
	//UINT mask = MyPopupMenu::MenuItemInfoMaskTypeAllSet;
	UINT mask = MyPopupMenu::MenuItemInfoMaskFTypeAllSet;
	BOOL retval = MyPopupMenu::TraceMenu(title, menu, index, mask);

	MyPopupMenu::SetCMenuAsCustom(menu);

#else
	MyPopupMenu* pPopupMenu = &m_myMainMenu;
	CMenu* pMenu = (CMenu*)pPopupMenu;
	//CMenu* pMenu = &m_myMainCMenu;
	// 
	//HMENU oldMenu = this->GetMenu();
	CMenu* oldMenu = this->GetMenu();
	BOOL retLoad = pMenu->LoadMenu(IDR_MAINFRAME);
	pPopupMenu->SetCMenuAsCustom();
	BOOL retSet = SetMenu(pMenu);
	//delete oldMenu;
#endif
}


// Measure item implementation relies on unique control/menu IDs
void CMainFrame::OnMeasureItem(int nIDCtl, LPMEASUREITEMSTRUCT lpMeasureItemStruct)
{
#if 0
	CWnd::OnMeasureItem(nIDCtl, lpMeasureItemStruct);
#else
	// Modified version of CWnd::OnMeasureItem(..) to support submenues
	if (lpMeasureItemStruct->CtlType == ODT_MENU)
	{
		_ASSERTE(lpMeasureItemStruct->CtlID == 0);
		CMenu* pMenu = NULL;

		_AFX_THREAD_STATE* pThreadState = _afxThreadState.GetData();
		if (pThreadState->m_hTrackingWindow == m_hWnd)
		{
			// start from popup
			pMenu = CMenu::FromHandle(pThreadState->m_hTrackingMenu);
		}
		else
		{
			// start from menubar
			pMenu = GetMenu();
		}

		_ASSERTE(pMenu);

		// Commented out: _AfxFindPopupMenuFromID ignores submenus 
		// pMenu = _AfxFindPopupMenuFromID(pMenu, lpMeasureItemStruct->itemID);
		// 
		// Is this overkill ??
		if (lpMeasureItemStruct->itemID != -1)
		{
			if (!MyPopupMenu::HasID(pMenu, lpMeasureItemStruct->itemID))
				pMenu = 0;
		}

		if (pMenu != NULL)
		{
			HWND hwnd = GetSafeHwnd();
			MyPopupMenu::OnMeasureItem(hwnd, lpMeasureItemStruct);
			//MyPopupMenu* pPopupMenu = (MyPopupMenu*)pMenu;
			//pPopupMenu->MeasureItem(lpMeasureItemStruct);
			//((MyPopupMenu*)pMenu)->MeasureItem(lpMeasureItemStruct);
		}
		else
		{
			TRACE(traceAppMsg, 0, "Warning: unknown WM_MEASUREITEM for menu item 0x%04X.\n",
				lpMeasureItemStruct->itemID);
		}
	}
	else
	{
		CWnd* pChild = CWnd::GetDescendantWindow(lpMeasureItemStruct->CtlID, TRUE);
		if (pChild != NULL && pChild->SendChildNotifyLastMsg())
			return;     // eaten by child
	}
	// not handled - do default
	CWnd::Default();
#endif
}

void CMainFrame::OnDrawItem(int nIDCtl, LPDRAWITEMSTRUCT lpDrawItemStruct)
{
	MyPopupMenu::OnDrawItem(nIDCtl, lpDrawItemStruct);
	int deb = 1;
}


void CMainFrame::OnHelpCrashhelp()
{
	// TODO: Add your command handler code here

	CString helpFileName = L"CrashHelp.pdf";
	HWND h = GetSafeHwnd();
	CMainFrame::OpenHelpFile(helpFileName, h);
}

void CMainFrame::OnDevelopmentoptionsTogglertlfordialogs()
{
	// TODO: Add your command handler code here
	if (CmboxviewApp::m_isRTL)
	{
		CmboxviewApp::m_isRTLForDialogs = !CmboxviewApp::m_isRTLForDialogs;
	}
}
