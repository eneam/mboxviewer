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

// NListView.cpp : implementation file
//

#include "stdafx.h"
#include "FileUtils.h"
#include "TextUtilsEx.h"
#include "HtmlUtils.h"
#include "mboxview.h"
#include "NListView.h"
#pragma warning (disable : 4786)
#include "Mime.h"
#include "MimeCode.h"
#include "MboxMail.h"
#include "OpenContainingFolderDlg.h"
#include "OpenArchiveFileDlg.h"
#include "FindInMailDlg.h"
#include "FindAdvancedDlg.h"
#include "AttachmentsConfig.h"
#include "CheckListBoxDlg.h"
#include "ExceptionUtil.h"
#include "SerializationHelper.h"
#include "MimeHelper.h"
#include "ColorStyleConfigDlg.h"
#include "MimeParser.h"
#include "ForwardMailDlg.h"
#include "MyTcpClient.h"
#include "GdiUtils.h"
#include "ResHelper.h"


#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

#ifdef _DEBUG
inline void BreakNlistView()
{
	int deb = 1;
}
#else
inline void BreakNlistView() {}
#endif

BOOL CreateInlineImageCache_WorkerThread(LPCWSTR cache, BOOL mainThread, CString &errorText);
BOOL CreateAttachmentCache_WorkerThread(LPCWSTR cache, BOOL mainThread, CString &errorText);
BOOL CreateEmlCache_WorkerThread(LPCWSTR cache, BOOL mainThread, CString &errorText);

BOOL NListView::m_fullImgFilePath = TRUE;
BOOL NListView::m_fullImgFilePath_Config = TRUE;
BOOL NListView::m_appendAttachmentPictures = FALSE;
BOOL NListView::m_exportMailsMode = FALSE;

bool ALongRightProcessProcForwardMails(const CUPDUPDATA* pCUPDUPData)
{
	FORWARD_MAILS_ARGS *args = (FORWARD_MAILS_ARGS*)pCUPDUPData->GetAppData();
	MboxMail::pCUPDUPData = pCUPDUPData;

	HANDLE h = GetCurrentThread();
	BOOL prio = SetThreadPriority(h, THREAD_PRIORITY_ABOVE_NORMAL);
	DWORD myThreadId = GetCurrentThreadId();
	DWORD myThreadPri = GetThreadPriority(h);
	TRACE(L"(ALongRightProcessProcForwardMails) threadId=%ld threadPriority=%ld\n", myThreadId, myThreadPri);

	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as parameter

	if (args->selectedMailIndexList)  // should always be true
	{
		args->ret = args->lview->ForwardMails_WorkerThread(args->forwardMailsData, args->selectedMailIndexList, args->errorText);
	}
	else
	{
		;// args->ret = args->lview->ForwardMailRange_WorkerThread(args->firstMail, args->lastMail, args->targetPrintSubFolderName, args->targetPrintFolderPath, args->errorText);
	}

	args->exitted = TRUE;
	return true;
}

bool ALongRightProcessProcWriteIndexFile(const CUPDUPDATA* pCUPDUPData)
{
	WRITE_INDEX_FILE_ARGS *args = (WRITE_INDEX_FILE_ARGS*)pCUPDUPData->GetAppData();
	MboxMail::pCUPDUPData = pCUPDUPData;

	HANDLE h = GetCurrentThread();
	BOOL prio = SetThreadPriority(h, THREAD_PRIORITY_ABOVE_NORMAL);
	DWORD myThreadId = GetCurrentThreadId();
	DWORD myThreadPri = GetThreadPriority(h);
	TRACE(L"(ALongRightProcessProcWriteIndexFile) threadId=%ld threadPriority=%ld\n", myThreadId, myThreadPri);

	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as param to printMailArchiveToTextFile

	BOOL mainThread = FALSE;
	args->ret = args->lview->SaveMails(args->cache, mainThread, args->errorText);

	args->exitted = TRUE;
	return true;
}


bool ALongRightProcessProcWriteAttachmentFile(const CUPDUPDATA* pCUPDUPData)
{
	WRITE_IMAGE_FILE_ARGS *args = (WRITE_IMAGE_FILE_ARGS*)pCUPDUPData->GetAppData();
	MboxMail::pCUPDUPData = pCUPDUPData;

	HANDLE h = GetCurrentThread();
	BOOL prio = SetThreadPriority(h, THREAD_PRIORITY_ABOVE_NORMAL);
	DWORD myThreadId = GetCurrentThreadId();
	DWORD myThreadPri = GetThreadPriority(h);
	TRACE(L"(ALongRightProcessProcWriteAttachmentFile) threadId=%ld threadPriority=%ld\n", myThreadId, myThreadPri);

	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as param to printMailArchiveToTextFile

	BOOL mainThread = FALSE;
	args->ret = CreateAttachmentCache_WorkerThread(args->cache, mainThread, args->errorText);

	args->exitted = TRUE;
	return true;
}

bool ALongRightProcessProcWriteEmlFile(const CUPDUPDATA* pCUPDUPData)
{
	WRITE_IMAGE_FILE_ARGS *args = (WRITE_IMAGE_FILE_ARGS*)pCUPDUPData->GetAppData();
	MboxMail::pCUPDUPData = pCUPDUPData;

	HANDLE h = GetCurrentThread();
	BOOL prio = SetThreadPriority(h, THREAD_PRIORITY_ABOVE_NORMAL);
	DWORD myThreadId = GetCurrentThreadId();
	DWORD myThreadPri = GetThreadPriority(h);
	TRACE(L"(ALongRightProcessProcWriteEmlFile) threadId=%ld threadPriority=%ld\n", myThreadId, myThreadPri);

	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as param to printMailArchiveToTextFile

	BOOL mainThread = FALSE;
	args->ret = CreateEmlCache_WorkerThread(args->cache, mainThread, args->errorText);

	args->exitted = TRUE;
	return true;
}

bool ALongRightProcessProcWriteInlineImageFile(const CUPDUPDATA* pCUPDUPData)
{
	WRITE_IMAGE_FILE_ARGS *args = (WRITE_IMAGE_FILE_ARGS*)pCUPDUPData->GetAppData();
	MboxMail::pCUPDUPData = pCUPDUPData;

	HANDLE h = GetCurrentThread();
	BOOL prio = SetThreadPriority(h, THREAD_PRIORITY_ABOVE_NORMAL);
	DWORD myThreadId = GetCurrentThreadId();
	DWORD myThreadPri = GetThreadPriority(h);
	TRACE(L"(ALongRightProcessProcWriteInlineImageFile) threadId=%ld threadPriority=%ld\n", myThreadId, myThreadPri);

	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as param to printMailArchiveToTextFile

	BOOL mainThread = FALSE;
	args->ret = CreateInlineImageCache_WorkerThread(args->cache, mainThread, args->errorText);

	args->exitted = TRUE;
	return true;
}

bool ALongRightProcessProcPrintMailGroupToSeparatePDF(const CUPDUPDATA* pCUPDUPData)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS *args = (PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS*)pCUPDUPData->GetAppData();
	MboxMail::pCUPDUPData = pCUPDUPData;

	HANDLE h = GetCurrentThread();
	BOOL prio = SetThreadPriority(h, THREAD_PRIORITY_ABOVE_NORMAL);
	DWORD myThreadId = GetCurrentThreadId();
	DWORD myThreadPri = GetThreadPriority(h);
	TRACE(L"(ALongRightProcessProcPrintMailGroupToSeparatePDF) threadId=%ld threadPriority=%ld\n", myThreadId, myThreadPri);

	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as param to printMailArchiveToTextFile

	// FIXME Cosider converting Range type functions to Select type function to reduce source code and maintenace cost
	// Select type expects array of mail indexes. Range type function could populate array from the range defined by indexes
	// Array can be large but users can select all mails anyway and create large array for Select type function 
	if (args->selectedMailIndexList)
	{
		args->ret = args->lview->PrintMailSelectedToSeparatePDF_WorkerThread(args->selectedMailIndexList, args->targetPrintSubFolderName, args->targetPrintFolderPath, args->errorText, 
			args->mergePDFs, args->mergedPDFPath);
	}
	else
	{
		_ASSERTE(FALSE);
	}

	args->exitted = TRUE;
	return true;
}

bool ALongRightProcessProcPrintMailGroupToSinglePDF(const CUPDUPDATA* pCUPDUPData)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS *args = (PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS*)pCUPDUPData->GetAppData();
	MboxMail::pCUPDUPData = pCUPDUPData;

	HANDLE h = GetCurrentThread();
	BOOL prio = SetThreadPriority(h, THREAD_PRIORITY_ABOVE_NORMAL);
	DWORD myThreadId = GetCurrentThreadId();
	DWORD myThreadPri = GetThreadPriority(h);
	TRACE(L"(ALongRightProcessProcPrintMailGroupToSinglePDF) threadId=%ld threadPriority=%ld\n", myThreadId, myThreadPri);

	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as parameter

	// FIXME Cosider converting Range type functions to Select type function to reduce source code and maintenace cost
	// Select type expects array of mail indexes. Range type function could populate array from the range defined by indexes
	// Array can be large but users can select all mails anyway and create large array for Select type function 
	if (args->selectedMailIndexList)
	{
		args->ret = args->lview->PrintMailSelectedToSinglePDF_WorkerThread(args->selectedMailIndexList, args->targetPrintSubFolderName, args->targetPrintFolderPath, args->errorText);
	}
	else
	{
		_ASSERTE(FALSE);
	}

	args->exitted = TRUE;
	return true;
}

bool ALongRightProcessProcPrintMailGroupToSeparateHTML(const CUPDUPDATA* pCUPDUPData)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_HTML_ARGS *args = (PRINT_MAIL_GROUP_TO_SEPARATE_HTML_ARGS*)pCUPDUPData->GetAppData();
	MboxMail::pCUPDUPData = pCUPDUPData;

	HANDLE h = GetCurrentThread();
	BOOL prio = SetThreadPriority(h, THREAD_PRIORITY_ABOVE_NORMAL);
	DWORD myThreadId = GetCurrentThreadId();
	DWORD myThreadPri = GetThreadPriority(h);
	TRACE(L"(ALongRightProcessProcPrintMailGroupToSeparateHTML) threadId=%ld threadPriority=%ld\n", myThreadId, myThreadPri);

	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as param to printMailArchiveToTextFile

	// FIXME Cosider converting Range type functions to Select type function to reduce source code and maintenace cost
	// Select type expects array of mail indexes. Range type function could populate array from the range defined by indexes
	// Array can be large but users can select all mails anyway and create large array for Select type function 
	if (args->selectedMailIndexList)
	{
		args->ret = args->lview->PrintMailSelectedToSeparateHTML_WorkerThread(args->selectedMailIndexList, args->targetPrintSubFolderName, args->targetPrintFolderPath, args->errorText);
	}
	else
	{
		_ASSERTE(FALSE);
	}

	args->exitted = TRUE;
	return true;
}

bool ALongRightProcessProcPrintMailGroupToSingleHTML(const CUPDUPDATA* pCUPDUPData)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_HTML_ARGS *args = (PRINT_MAIL_GROUP_TO_SEPARATE_HTML_ARGS*)pCUPDUPData->GetAppData();
	MboxMail::pCUPDUPData = pCUPDUPData;

	HANDLE h = GetCurrentThread();
	BOOL prio = SetThreadPriority(h, THREAD_PRIORITY_ABOVE_NORMAL);
	DWORD myThreadId = GetCurrentThreadId();
	DWORD myThreadPri = GetThreadPriority(h);

	TRACE(L"(ALongRightProcessProcPrintMailGroupToSingleHTML) threadId=%ld threadPriority=%ld\n", myThreadId, myThreadPri);

	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as param to printMailArchiveToTextFile

	// FIXME Cosider converting Range type functions to Select type function to reduce source code and maintenace cost
	// Select type expects array of mail indexes. Range type function could populate array from the range defined by indexes
	// Array can be large but users can select all mails anyway and create large array for Select type function 
	if (args->selectedMailIndexList)
	{
		args->ret = args->lview->PrintMailSelectedToSingleHTML_WorkerThread(args->selectedMailIndexList, args->targetPrintSubFolderName, args->targetPrintFolderPath, args->errorText);
	}
	else
	{
		_ASSERTE(FALSE);
	}

	args->exitted = TRUE;
	return true;
}

bool ALongRightProcessProcPrintMailGroupToSingleTEXT(const CUPDUPDATA* pCUPDUPData)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_HTML_ARGS *args = (PRINT_MAIL_GROUP_TO_SEPARATE_HTML_ARGS*)pCUPDUPData->GetAppData();
	MboxMail::pCUPDUPData = pCUPDUPData;

	HANDLE h = GetCurrentThread();
	BOOL prio = SetThreadPriority(h, THREAD_PRIORITY_ABOVE_NORMAL);
	DWORD myThreadId = GetCurrentThreadId();
	DWORD myThreadPri = GetThreadPriority(h);
	TRACE(L"(ALongRightProcessProcPrintMailGroupToSingleTEXT) threadId=%ld threadPriority=%ld\n", myThreadId, myThreadPri);

	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as param to printMailArchiveToTextFile

	// FIXME Cosider converting Range type functions to Select type function to reduce source code and maintenace cost
	// Select type expects array of mail indexes. Range type function could populate array from the range defined by indexes
	// Array can be large but users can select all mails anyway and create large array for Select type function 
	if (args->selectedMailIndexList)
	{
		args->ret = args->lview->PrintMailSelectedToSingleTEXT_WorkerThread(args->selectedMailIndexList, args->targetPrintSubFolderName, args->targetPrintFolderPath, args->errorText);
	}
	else
	{
		;// args->ret = args->lview->PrintMailRangeToSingleTEXT_WorkerThread(args->firstMail, args->lastMail, args->targetPrintSubFolderName, args->targetPrintFolderPath, args->errorText);
	}

	args->exitted = TRUE;
	return true;
}

IMPLEMENT_DYNCREATE(NListView, CWnd)

/////////////////////////////////////////////////////////////////////////////
// NListView

//CString GetDateFormat(int i);

NListView::NListView() : m_list(this)
{
	ResetFileMapView();

	m_which = NULL;

	m_bApplyColorStyle = TRUE;
	m_bLongMailAddress = TRUE;
	m_name = new SimpleString(64);
	m_addr = new SimpleString(64);

	m_SearchType = 0;

	m_searchStringInMail.Empty();
	m_bCaseSensInMail = FALSE;
	m_bWholeWordInMail = FALSE;

	m_advancedParams.SetDflts();  // FIXME why not set in constructor
	m_advancedParams.ResetFilterDates();

	m_advancedFind = FALSE;
	m_gmtTime = 0;
	m_bStartSearchAtSelectedItem = 0; // FALSE; TODO: is this desired feature ?

	m_findAllCount = 0;
	m_bEditFindFirst = FALSE;  // must call OnEditFind() first and not OnEditFindAgain()
	m_lastFindPos = -1;  // last find position, start for the next find again
	m_searchString.Empty();

	// Find Text related
	m_bHighlightAllSet = FALSE;
	m_bCaseSens = FALSE;
	m_bWholeWord = FALSE;

	m_lastSel = -1;  // last/currently selected item
	m_bInFind = FALSE;
	m_subjectSortType = 0;

	CString section_options = CString(sz_Software_mboxview) + L"\\Options";

	m_subjectSortType = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"subjectSortType");

	int iFormat = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"dayMonthYearFormat");
	m_format = MboxMail::GetDateFormat(iFormat);
	m_advancedParams.m_dateTimeFormat = MboxMail::GetPickerDateFormat(iFormat);
	m_findParams.m_dateTimeFormat = MboxMail::GetPickerDateFormat(iFormat);

	m_gmtTime = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"timeType");

	CString exportEML;
	BOOL retval = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_options, L"exportEML", exportEML);
	if (retval == TRUE) {
		if (exportEML.Compare(L"y") == 0)
			m_bExportEml = TRUE;
		else
			m_bExportEml = FALSE;
	}
	else {
		exportEML = L"n";
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_options, L"exportEML", exportEML);
		m_bExportEml = FALSE;
	}

	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	DWORD relativeImgFilePath_Config = 0;
	BOOL found  = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_general, L"relativeInlineImageFilePath", relativeImgFilePath_Config);
	if (found)
		m_fullImgFilePath_Config = !relativeImgFilePath_Config;
	else
		; // m_fullImgFilePath_Config == default

	DWORD barDelay;
	retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"progressBarDelay", barDelay);
	if (retval == TRUE) {
		m_maxSearchDuration = barDelay;
	}
	else {
		barDelay = 1;
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"progressBarDelay", barDelay);
		m_maxSearchDuration = barDelay;
	}

	m_frameCx_TreeNotInHide = 700;
	m_frameCy_TreeNotInHide = 200;
	m_frameCx_TreeInHide = 700;
	m_frameCy_TreeInHide = 200;

	CString section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacement";
	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
		section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacementPreview";
	else if (CMainFrame::m_commandLineParms.m_bDirectFileOpenMode)
		section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacementDirect";

	BOOL ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_wnd, L"ListFrameTreeNotHiddenWidth", m_frameCx_TreeNotInHide);
	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_wnd, L"ListFrameTreeNotHiddenHeight", m_frameCy_TreeNotInHide);
	//
	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_wnd, L"ListFrameTreeHiddenWidth", m_frameCx_TreeInHide);
	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_wnd, L"ListFrameTreeHiddenHeight", m_frameCy_TreeInHide);

	m_acp = GetACP();
	m_tcpPort = 61333;
	m_enableForwardMailsLog = FALSE;  // FIXME set both to TRUE to always log
	m_enableSMTPProtocolLog = FALSE;
	m_developmentMode = FALSE;
	m_fontSizePDF = 0;
}

NListView::~NListView()
{
	delete m_name;
	delete m_addr;
}

BEGIN_MESSAGE_MAP(NListView, CWnd)
	//{{AFX_MSG_MAP(NListView)
	ON_WM_CREATE()
	ON_WM_SIZE()
	ON_WM_MOUSEWHEEL()
	ON_WM_ERASEBKGND()
	//}}AFX_MSG_MAP
	//ON_NOTIFY(LVN_ITEMACTIVATE, IDC_LIST, OnActivating)

	//ON_NOTIFY(HDN_DIVIDERDBLCLICK, IDC_LIST, OnDividerdblclick)
	//ON_NOTIFY(LVN_ITEMCHANGED, IDC_LIST, OnItemchangedListCtrl)
	ON_NOTIFY_REFLECT(HDN_DIVIDERDBLCLICK, OnDividerdblclick)
	ON_NOTIFY(NM_CLICK, IDC_LIST, OnActivating)  // Left Click
	ON_NOTIFY(NM_RCLICK, IDC_LIST, OnRClick)  // Right Click Menu
	ON_NOTIFY(NM_DBLCLK, IDC_LIST, OnDoubleClick)
	ON_NOTIFY(LVN_COLUMNCLICK, IDC_LIST, OnColumnClick)
	ON_NOTIFY(LVN_KEYDOWN, IDC_LIST, OnKeydown)
	ON_NOTIFY(NM_CUSTOMDRAW, IDC_LIST, OnCustomDraw)
	//ON_NOTIFY(LVN_GETDISPINFO, IDC_LIST, OnGetDispInfo)  // no longer needed, using Custom Draw
	ON_UPDATE_COMMAND_UI(ID_EDIT_FIND, OnUpdateEditFind)
	ON_COMMAND(ID_EDIT_FIND, OnEditFind)
	ON_UPDATE_COMMAND_UI(ID_EDIT_FINDAGAIN, OnUpdateEditFindAgain)
	ON_COMMAND(ID_EDIT_FINDAGAIN, OnEditFindAgain)
	ON_COMMAND(ID_EDIT_VIEWEML, &NListView::OnEditVieweml)
	ON_UPDATE_COMMAND_UI(ID_EDIT_VIEWEML, &NListView::OnUpdateEditVieweml)
	//
#if 0  // was trying to learn how this works
	ON_NOTIFY(LVN_ITEMCHANGED, IDC_LIST, OnItemchangedListCtrl)
	ON_NOTIFY(LVN_ODSTATECHANGED, IDC_LIST, OnODStateChangedListCtrl)
	ON_NOTIFY(LVN_ODFINDITEM, IDC_LIST, OnODFindItemListCtrl)
	ON_NOTIFY(LVN_ODCACHEHINT, IDC_LIST, OnODCacheHintListCtrl)
	ON_NOTIFY(NM_RELEASEDCAPTURE, IDC_LIST, OnReleaseCaptureListCtrl)
#endif
	ON_COMMAND(ID_EDIT_FINDADVANCED, &NListView::OnEditFindadvanced)
	ON_UPDATE_COMMAND_UI(ID_EDIT_FINDADVANCED, &NListView::OnUpdateEditFindadvanced)
	//ON_WM_SETFOCUS()
	//ON_WM_MOUSEHOVER()
	ON_WM_TIMER()
	ON_WM_CLOSE()
	ON_MESSAGE(WM_CMD_PARAM_ATTACHMENT_HINT_MESSAGE, &NListView::OnCmdParam_AttachmentHint)
	//ON_NOTIFY(LVN_GETINFOTIP, IDC_LIST, OnTvnGetInfoTip)
	ON_MESSAGE(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, &NListView::OnCmdParam_OnSwitchWindow)
	//ON_NOTIFY_EX(TTN_NEEDTEXT, 0, &NListView::OnTtnNeedText)
END_MESSAGE_MAP()

int g_pointSize = 85;
CString g_fontName = "Tahoma";

void NListView::ResetFont()
{
	m_boldFont.DeleteObject();
	m_font.DeleteObject();
	if( ! m_font.CreatePointFont (g_pointSize, g_fontName) )
		m_font.CreatePointFont (85, L"Arial");
	LOGFONT	lf;
	m_font.GetLogFont(&lf);
	lf.lfWeight = FW_BOLD;
	m_boldFont.CreateFontIndirect(&lf);
    m_list.SetFont(&m_font);
}

void NListView::OnDividerdblclick(NMHDR* pNMHDR, LRESULT* pResult)
{
	*pResult = 0;
	return;
}

void NListView::OnDoubleClick(NMHDR* pNMHDR, LRESULT* pResult) 
{
	*pResult = 0;

	SHORT ctrlState = GetAsyncKeyState(VK_CONTROL);
	SHORT shiftState = GetAsyncKeyState(VK_SHIFT);
	// TODO: this is work around. Not sure why OnDoubleClick() is called when SHIFT+MOUSEWHEEL_UP
	if (((ctrlState & 0x8000) != 0) || ((shiftState & 0x8000) != 0)) return;

	NMITEMACTIVATE *pnm = (NMITEMACTIVATE *)pNMHDR;
	if( m_lastSel == pnm->iItem ) {
		CString path = CMainFrame::GetMboxviewTempPath();

		if(FileUtils::PathDirExists(path) )
			ShellExecute(NULL, L"open", path, NULL,NULL, SW_SHOWNORMAL );
	}
}

void NListView::OnActivating(NMHDR* pNMHDR, LRESULT* pResult) 
{
	NMITEMACTIVATE *pnm = (NMITEMACTIVATE *)pNMHDR;
	if( m_lastSel != pnm->iItem ) {
		TRACE(L"Selecting %d\n", pnm->iItem);
		m_bApplyColorStyle = TRUE;
		SelectItem(pnm->iItem);
	}

	if (m_lastSel != pnm->iItem)
		int deb = 1;
	else
		int deb = 1;

	*pResult = 0;
}

void NListView::OnRClick(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMITEMACTIVATE pnm = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		NMsgView *pMsgView = pFrame->GetMsgView();

		if (pFrame->m_bViewMessageHeaders || (pMsgView && pMsgView->m_hdrWindowLen))
		{
			CString InfoTxt =
				L"Mail Context Menu is not available while View Message Headers are enabled.\n\n"
				"Please disable viewing of raw mail headers to reenable the context menu.\n\n"
				"UnCheck \"View -> View Message Headers\"\n"
				"and Right click on Mail Header Pane at the top of the Message Window and\n"
				"UnCheck \"View Message Header\"\n";

			ResHelper::TranslateString(InfoTxt);

			int answer = MessageBox(InfoTxt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
			return;
		}
	}


	int iItem = pnm->iItem;
	if (iItem < 0)
	{
		int cnt = MboxMail::s_mails.GetCount();
		int refcnt = MboxMail::s_mails_ref.GetCount();
		if ((cnt <= 0) && (refcnt > 0))
		{
			CPoint pt;
			::GetCursorPos(&pt);
			CWnd *wnd = WindowFromPoint(pt);

			CMenu menu;
			menu.CreatePopupMenu();
			menu.AppendMenu(MF_SEPARATOR);

			const UINT M_Reload_Id = 1;
			MyAppendMenu(&menu, M_Reload_Id, L"Restore User Selected Mails from saved Mail List file");

			CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

			int index = 0;
			ResHelper::LoadMenuItemsInfo(&menu, index);
			ResHelper::UpdateMenuItemsInfo(&menu, index);

			UINT command = menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, this);

			UINT nFlags = TPM_RETURNCMD;
			CString menuString;
			int chrCnt = menu.GetMenuString(command, menuString, nFlags);

			switch (command)
			{
			case M_Reload_Id:
			{
				ReloadMboxListFile_v2();
				RefreshMailsOnUserSelectsMailListMark();
			}
			break;
			default:
			{
				;
			}
			}
		}
		return;
	}

	POSITION pos = m_list.GetFirstSelectedItemPosition();
	int nItem;
	int nItemCnt = 0;
	while (pos)
	{
		nItem = m_list.GetNextSelectedItem(pos);
		nItemCnt++;
		if (nItemCnt > 1)
			break;
	}

	NListView::m_fullImgFilePath = NListView::m_fullImgFilePath_Config;

	if (nItemCnt == 1)
		OnRClickSingleSelect(pNMHDR, pResult);
	else
		OnRClickMultipleSelect(pNMHDR, pResult);
}

void NListView::OnRClickSingleSelect(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMITEMACTIVATE pnm = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	int iItem = pnm->iItem;

	if (iItem < 0)
		return;

	CPoint pt;
	::GetCursorPos(&pt);
	CWnd *wnd = WindowFromPoint(pt);

	RECT clrect;
	GetClientRect(&clrect);
	CPoint ptpos = pnm->ptAction;
	ClientToScreen(&ptpos);

	CMenu menu;
	menu.CreatePopupMenu();
	menu.AppendMenu(MF_SEPARATOR);

	CMenu printToSubMenu;
	printToSubMenu.CreatePopupMenu();
	printToSubMenu.AppendMenu(MF_SEPARATOR);

	CMenu printGroupToSubMenu;
	printGroupToSubMenu.CreatePopupMenu();
	printGroupToSubMenu.AppendMenu(MF_SEPARATOR);

	CMenu printPDFGroupToSubMenu;
	printPDFGroupToSubMenu.CreatePopupMenu();

	CMenu exportMailsToSubMenu;
	exportMailsToSubMenu.CreatePopupMenu();

	// Create enums or replace switch statment with if else ..
	// XXX_GROUP_Id represents group of related emails, related == convesation
	const UINT S_TEXT_Id = 1;
	MyAppendMenu(&printToSubMenu, S_TEXT_Id, L"Text..");

	const UINT S_HTML_Id = 2;
	MyAppendMenu(&printToSubMenu, S_HTML_Id, L"HTML...");

	const UINT S_PRINTER_Id = 3;
	MyAppendMenu(&printToSubMenu, S_PRINTER_Id, L"Printer...");

	const UINT S_PDF_DIRECT_Id = 4;
	MyAppendMenu(&printToSubMenu, S_PDF_DIRECT_Id, L"PDF...");

	menu.AppendMenu(MF_POPUP | MF_STRING, (UINT_PTR)printToSubMenu.GetSafeHmenu(), L"Print To");
	menu.AppendMenu(MF_SEPARATOR);

	const UINT S_TEXT_GROUP_Id = 5;
	MyAppendMenu(&printGroupToSubMenu, S_TEXT_GROUP_Id, L"Text..");

	const UINT S_HTML_GROUP_Id = 6;
	MyAppendMenu(&printGroupToSubMenu, S_HTML_GROUP_Id, L"HTML...");

	const UINT S_PRINTER_GROUP_Id = 7;
	MyAppendMenu(&printGroupToSubMenu, S_PRINTER_GROUP_Id, L"Printer...");

	const UINT S_PDF_DIRECT_GROUP_Id = 8;
	//MyAppendMenu(&printGroupToSubMenu, S_PDF_DIRECT_GROUP_Id, L"PDF...");

	const UINT S_CSV_GROUP_Id = 9;
	MyAppendMenu(&printGroupToSubMenu, S_CSV_GROUP_Id, L"CSV...");

	const UINT S_PDF_DIRECT_DEFAULT_Id = 30;
	MyAppendMenu(&printPDFGroupToSubMenu, S_PDF_DIRECT_DEFAULT_Id, L"Default");

	// IDs 31,32, 33 available now
	const UINT S_PDF_DIRECT_FONT_16_Id = 31;
	const UINT S_PDF_DIRECT_FONT_20_Id = 32;
	const UINT S_PDF_DIRECT_FONT_24_Id = 33;

	const UINT S_PDF_DIRECT_MERGE_Id = 34;
	MyAppendMenu(&printPDFGroupToSubMenu, S_PDF_DIRECT_MERGE_Id, L"Merge");

	const UINT S_PDF_DIRECT_HELP_Id = 35;
	MyAppendMenu(&printPDFGroupToSubMenu, S_PDF_DIRECT_HELP_Id, L"Help");

	if (abs(MboxMail::b_mails_which_sorted) == 99)
	{
		menu.AppendMenu(MF_POPUP | MF_STRING, (UINT_PTR)printGroupToSubMenu.GetSafeHmenu(), L"Print Conversation Thread Mails To");
		menu.AppendMenu(MF_SEPARATOR);
		printGroupToSubMenu.AppendMenu(MF_POPUP | MF_STRING, (UINT_PTR)printPDFGroupToSubMenu.GetSafeHmenu(), L"PDF..");
		printGroupToSubMenu.AppendMenu(MF_SEPARATOR);
	}

	if ((abs(MboxMail::b_mails_which_sorted) == 4) && (m_subjectSortType == 1))
	{
		menu.AppendMenu(MF_POPUP | MF_STRING, (UINT_PTR)printGroupToSubMenu.GetSafeHmenu(), L"Print Subject Thread Mails To");
		menu.AppendMenu(MF_SEPARATOR);
		printGroupToSubMenu.AppendMenu(MF_POPUP | MF_STRING, (UINT_PTR)printPDFGroupToSubMenu.GetSafeHmenu(), L"PDF..");
		printGroupToSubMenu.AppendMenu(MF_SEPARATOR);
	}

	const UINT S_HTML_OPEN_Id = 10;
	MyAppendMenu(&menu, S_HTML_OPEN_Id, L"Open in Browser");

	const UINT S_HTML_OPEN_RELATED_Id = 11;
	if (abs(MboxMail::b_mails_which_sorted) == 99)
	{
		MyAppendMenu(&menu, S_HTML_OPEN_RELATED_Id, L"Open Conversation Thread Mails in Browser");
	}

	if ((abs(MboxMail::b_mails_which_sorted) == 4) && (m_subjectSortType == 1))
	{
		MyAppendMenu(&menu, S_HTML_OPEN_RELATED_Id, L"Open Subject Thread Mails in Browser");
	}

	const UINT S_HTML_OPEN_RELATED_FILES_Id = 12;
	MyAppendMenu(&menu, S_HTML_OPEN_RELATED_FILES_Id, L"Open Related Files Location");

	const UINT S_HTML_FIND_Id = 13;
	MyAppendMenu(&menu, S_HTML_FIND_Id, L"Find Text");

	const UINT S_ADVANCED_FIND_Id = 14;
	MyAppendMenu(&menu, S_ADVANCED_FIND_Id, L"Find Advanced");

	const UINT S_REMOVE_Id = 15;
	if (pFrame && (MboxMail::IsFindMailsSelected() || MboxMail::IsUserMailsSelected())) {
		menu.AppendMenu(MF_SEPARATOR);
		MyAppendMenu(&menu, S_REMOVE_Id, L"Remove Selected Mail");
	}

	const UINT S_REMOVE_ALL_Id = 16;
	if (pFrame && (MboxMail::IsFindMailsSelected() || (MboxMail::IsUserMailsSelected()) && (pFrame->IsUserMailsListEnabled()))) {
		MyAppendMenu(&menu, S_REMOVE_ALL_Id, L"Remove All Mails");
	}

	const UINT S_COPY_ALL_Id = 17;
	const UINT S_COPY_SELECTED_Id = 18;
	if (pFrame && (MboxMail::IsAllMailsSelected() || MboxMail::IsLabelMailsSelected() || MboxMail::IsFindMailsSelected()) && pFrame->IsUserMailsListEnabled()) {
		menu.AppendMenu(MF_SEPARATOR);
		MyAppendMenu(&menu, S_COPY_ALL_Id, L"Copy All into User Selected Mails");
		MyAppendMenu(&menu, S_COPY_SELECTED_Id, L"Copy Selected into User Selected Mails");
	}

	const UINT S_SAVE_AS_ARCHIVE_Id = 19;
	const UINT M_ARCHIVE_LOCATION_Id = 20;
	if (pFrame && (MboxMail::IsFindMailsSelected() || MboxMail::IsUserMailsSelected())) {
		menu.AppendMenu(MF_SEPARATOR);
		MyAppendMenu(&menu, S_SAVE_AS_ARCHIVE_Id, L"Save All as Mbox Mail Archive file");
		MyAppendMenu(&menu, M_ARCHIVE_LOCATION_Id, L"Open Mbox Mail Archive Location");
	}

	const UINT S_SAVE_AS_LIST_Id = 21;
	const UINT S_RELOAD_LIST_Id = 22;
	const UINT M_LIST_LOCATION_Id = 23;
	if (pFrame && (MboxMail::IsUserMailsSelected())) {
		menu.AppendMenu(MF_SEPARATOR);
		MyAppendMenu(&menu, S_SAVE_AS_LIST_Id, L"Save All as Mail List file");
		MyAppendMenu(&menu, S_RELOAD_LIST_Id, L"Reload from Mail List file");
		MyAppendMenu(&menu, M_LIST_LOCATION_Id, L"Open Mail List Location");
	}

	const UINT S_RESTORE_LIST_Id = 24;
	if (pFrame && MboxMail::IsAllMailsSelected() && pFrame->IsUserMailsListEnabled() && (MboxMail::s_mails_edit.GetCount() == 0)) {
		//menu.AppendMenu(MF_SEPARATOR);
		MyAppendMenu(&menu, S_RESTORE_LIST_Id, L"Restore User Selected Mails from saved Mail List file");
	}

	const UINT M_REMOVE_DUPLICATE_MAILS_Id = 25;
	//if (pFrame && (MboxMail::IsUserMailsSelected() || MboxMail::IsLabelMailsSelected())) {
	if (pFrame && (MboxMail::IsUserMailsSelected())) {
		menu.AppendMenu(MF_SEPARATOR);
		MyAppendMenu(&menu, M_REMOVE_DUPLICATE_MAILS_Id, L"Remove Duplicate Mails");
	}

	const UINT M_COPY_MAILS_TO_FOLDERS_Id = 26;
	if (pFrame && !MboxMail::IsFolderMailsSelected()) {
		//menu.AppendMenu(MF_SEPARATOR);
		//MyAppendMenu(&menu, M_COPY_MAILS_TO_FOLDERS_Id, L"Copy to Folders");
		;
	}

	const UINT M_FORWARD_MAIL_Id = 27;
	MyAppendMenu(&menu, M_FORWARD_MAIL_Id, L"Forward Mail");

	const UINT M_FORWARD_RELATED_MAILS_Id = 28;
	MyAppendMenu(&menu, M_FORWARD_RELATED_MAILS_Id, L"Forward Related Mails");

	const UINT M_EXPORT_SELECTED_MAILS_TO_SINGLE_FILE_Id = 48;
	MyAppendMenu(&exportMailsToSubMenu, M_EXPORT_SELECTED_MAILS_TO_SINGLE_FILE_Id, L"To File");

	const UINT M_EXPORT_SELECTED_MAILS_CONFIG_Id = 49;
	MyAppendMenu(&exportMailsToSubMenu, M_EXPORT_SELECTED_MAILS_CONFIG_Id, L"Config");

	const UINT M_EXPORT_SELECTED_MAILS_HELP_Id = 50;
	MyAppendMenu(&exportMailsToSubMenu, M_EXPORT_SELECTED_MAILS_HELP_Id, L"Help");

	menu.AppendMenu(MF_POPUP | MF_STRING, (UINT_PTR)exportMailsToSubMenu.GetSafeHmenu(), L"Export Selected Mail");
	menu.AppendMenu(MF_SEPARATOR);

	int index = 1;
	ResHelper::LoadMenuItemsInfo(&menu, index);
	ResHelper::UpdateMenuItemsInfo(&menu, index);

	UINT command = menu.TrackPopupMenuEx(TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, this, 0);

	DWORD error = GetLastError();

	UINT nFlags = TPM_RETURNCMD;
	CString menuString;
	int chrCnt = menu.GetMenuString(command, menuString, nFlags);

	BOOL multipleSelectedMails = FALSE;
	BOOL itemSelected = FALSE;
	switch (command)
	{
	case S_HTML_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bSingleMailView || attachmentConfigParams->m_bMultipleMailsView)
				NListView::m_appendAttachmentPictures = TRUE;
		}
		
		CString fileName;
		if (pFrame) {
			pFrame->OnPrintSingleMailtoText(iItem, 1, fileName);
		}

		NListView::m_appendAttachmentPictures = FALSE;
		int deb = 1;
	}
	break;
	case S_TEXT_Id:
	{
		CString fileName;
		if (pFrame) {
			pFrame->OnPrintSingleMailtoText(iItem, 0, fileName);
		}
		int deb = 1;
	}
	break;
	case S_PRINTER_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bSingleMailPrint || attachmentConfigParams->m_bMultipleMailsPrint)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		// IHTMLDocument require full path to embedded images
		m_fullImgFilePath = TRUE;

		CString fileName;
		if (pFrame) {
			MboxMail::ShowHint(HintConfig::PrintToPrinterHint, GetSafeHwnd());
			pFrame->OnPrintSingleMailtoText(iItem, 1, fileName, FALSE, TRUE);
		}

		m_fullImgFilePath = m_fullImgFilePath_Config;
		NListView::m_appendAttachmentPictures = FALSE;
	}
	break;
	case S_PDF_DIRECT_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bSingleMailPrint || attachmentConfigParams->m_bMultipleMailsPrint)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		MboxMail::ShowHint(HintConfig::PrintToPDFHint, GetSafeHwnd());
		CString errorText;
		CString targetPrintSubFolderName;
		CString targetPrintFolderPath;
		int ret = PrintMailRangeToSinglePDF_Thread(iItem, iItem, targetPrintSubFolderName);

		NListView::m_appendAttachmentPictures = FALSE;
		int deb = 1;
	}
	break;
	case S_CSV_GROUP_Id:
	{
		int ret = PrintMailRangeToSingleCSV_Thread(iItem);
		int deb = 1;
	}
	break;
	case S_HTML_GROUP_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bMultipleMailsView)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		CString fileName;
		if (pFrame)
		{
			if (pFrame->m_NamePatternParams.m_bPrintToSeparateHTMLFiles)
			{
				int firstMail = 0;
				int lastMail = 0;
				FindFirstAndLastMailOfConversation(iItem, firstMail, lastMail);
				if (firstMail == lastMail)
				{
					pFrame->OnPrintSingleMailtoText(iItem, 1, fileName);
				}
				else
				{
					CString errorText;
					int ret = PrintMailConversationToSeparateHTML_Thread(iItem, errorText);
				}
				int deb = 1;
			}
			else
			{
				PrintMailGroupToText(multipleSelectedMails, iItem, 1);
				//CString errorText;
				//int ret = PrintMailConversationToSingleHTML_Thread(iItem, errorText);
				int deb = 1;
			}
		}

		NListView::m_appendAttachmentPictures = FALSE;
		int deb = 1;
	}
	break;
	case S_TEXT_GROUP_Id:
	{
		//PrintMailGroupToText(multipleSelectedMails, iItem, 0);
		// TODO: we alway make the same call; fix it
		if (pFrame)
		{
			if (pFrame->m_NamePatternParams.m_bPrintToSeparateTEXTFiles)
			{
				int firstMail = 0;
				int lastMail = 0;
				FindFirstAndLastMailOfConversation(iItem, firstMail, lastMail);
				if (firstMail == lastMail)
				{
					//pFrame->OnPrintSingleMailtoText(iItem, 1, fileName);
					PrintMailGroupToText(multipleSelectedMails, iItem, 0);
				}
				else
				{
					CString errorText;
					//int ret = PrintMailConversationToSeparateHTML_Thread(iItem, errorText);
					PrintMailGroupToText(multipleSelectedMails, iItem, 0);
				}
				int deb = 1;
			}
			else
			{
				//PrintMailGroupToText(multipleSelectedMails, iItem, 1);
				PrintMailGroupToText(multipleSelectedMails, iItem, 0);
				//CString errorText;
				//int ret = PrintMailConversationToSingleHTML_Thread(iItem, errorText);
				int deb = 1;
			}
		}
		int deb = 1;
	}
	break;
	case S_PRINTER_GROUP_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bMultipleMailsPrint)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		// IHTMLDocument require full path to embedded images
		m_fullImgFilePath = TRUE;

		MboxMail::ShowHint(HintConfig::PrintToPrinterHint, GetSafeHwnd());
		//PrintMailGroupToText(multipleSelectedMails, iItem, 1, FALSE, TRUE);
		// TODO: we alway make the same call; fix it
		if (pFrame)
		{
			if (pFrame->m_NamePatternParams.m_bPrintToSeparateTEXTFiles)
			{
				int firstMail = 0;
				int lastMail = 0;
				FindFirstAndLastMailOfConversation(iItem, firstMail, lastMail);
				if (firstMail == lastMail)
				{
					//pFrame->OnPrintSingleMailtoText(iItem, 1, fileName);
					PrintMailGroupToText(multipleSelectedMails, iItem, 1, FALSE, TRUE);
				}
				else
				{
					CString errorText;
					//int ret = PrintMailConversationToSeparateHTML_Thread(iItem, errorText);
					PrintMailGroupToText(multipleSelectedMails, iItem, 1, FALSE, TRUE);
				}
				int deb = 1;
			}
			else
			{
				//PrintMailGroupToText(multipleSelectedMails, iItem, 1);
				PrintMailGroupToText(multipleSelectedMails, iItem, 1, FALSE, TRUE);
				//CString errorText;
				//int ret = PrintMailConversationToSingleHTML_Thread(iItem, errorText);
				int deb = 1;
			}
		}

		m_fullImgFilePath = m_fullImgFilePath_Config;
		NListView::m_appendAttachmentPictures = FALSE;
	}
	break;
	case S_PDF_DIRECT_GROUP_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bMultipleMailsPrint)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		if (pFrame)
		{
			MboxMail::ShowHint(HintConfig::PrintToPDFHint, GetSafeHwnd());
			if (pFrame->m_NamePatternParams.m_bPrintToSeparatePDFFiles)
			{
				CString errorText;
				BOOL mergePDFs = FALSE;
				int ret = PrintMailConversationToSeparatePDF_Thread(iItem, mergePDFs, errorText);
				int deb = 1;
			}
			else
			{
				CString errorText;
				int ret = PrintMailConversationToSinglePDF_Thread(iItem, errorText);
				int deb = 1;
			}
		}

		NListView::m_appendAttachmentPictures = FALSE;
		int deb = 1;
	}
	break;
	case S_PDF_DIRECT_DEFAULT_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bSingleMailPrint || attachmentConfigParams->m_bMultipleMailsPrint)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		//PrintToPDFMultipleSelect(0);
		if (pFrame)
		{
			MboxMail::ShowHint(HintConfig::PrintToPDFHint, GetSafeHwnd());
			if (pFrame->m_NamePatternParams.m_bPrintToSeparatePDFFiles)
			{
				CString errorText;
				BOOL mergePDFs = FALSE;
				int ret = PrintMailConversationToSeparatePDF_Thread(iItem, mergePDFs, errorText);
				int deb = 1;
			}
			else
			{
				CString errorText;
				int ret = PrintMailConversationToSinglePDF_Thread(iItem, errorText);
				int deb = 1;
			}
		}
		NListView::m_appendAttachmentPictures = TRUE;
		int deb = 1;
	}
	break;
	case S_PDF_DIRECT_MERGE_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bSingleMailPrint || attachmentConfigParams->m_bMultipleMailsPrint)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		CString targetPrintSubFolderName = L"PDF_GROUP";
		CString targetPrintFolderPath;
		//int ret = PrintMailSelectedToSeparatePDF_Thread(targetPrintSubFolderName, targetPrintFolderPath, TRUE);

		CString errorText;
		BOOL mergePDFs = TRUE;
		int ret = PrintMailConversationToSeparatePDF_Thread(iItem, mergePDFs, errorText);

		NListView::m_appendAttachmentPictures = FALSE;
		int deb = 1;
	}
	break;
	case S_PDF_DIRECT_HELP_Id:
	{
		//CString helpFileName = "PrintMultipleMailsToSinglePDF.htm";
		CString helpFileName = L"PrintMultipleMailsToPDF.pdf";
		HWND h = GetSafeHwnd();
		CMainFrame::OpenHelpFile(helpFileName, h);
		int deb = 1;
	}
	break;
	case S_HTML_OPEN_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bSingleMailView || attachmentConfigParams->m_bMultipleMailsView)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		CString fileName;
		if (pFrame)
		{
			BOOL forceOpen = TRUE;
			pFrame->OnPrintSingleMailtoText(iItem, 1, fileName, forceOpen);
		}

		NListView::m_appendAttachmentPictures = FALSE;
	}
	break;
	case S_HTML_OPEN_RELATED_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bMultipleMailsView)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		BOOL forceOpen = TRUE;
		PrintMailGroupToText(multipleSelectedMails, iItem, 1, forceOpen);

		NListView::m_appendAttachmentPictures = FALSE;
	}
	break;

	case S_HTML_OPEN_RELATED_FILES_Id:
	{
		CString path = CMainFrame::GetMboxviewTempPath();

		if (FileUtils::PathDirExists(path)) {
			HWND h = GetSafeHwnd();
			HINSTANCE result = ShellExecute(h, L"open", path, NULL, NULL, SW_SHOWNORMAL);
			CMainFrame::CheckShellExecuteResult(result, h);
		}
		int deb = 1;
	}
	break;

	case S_HTML_FIND_Id: 
	{
		itemSelected = FindInHTML(iItem);
	}
	break;
	case S_REMOVE_Id:
	{
		int firstItemRemoved = RemoveSelectedMails();
		if (MboxMail::s_mails.GetCount() > 0)
		{
			if (firstItemRemoved > 0)
				SelectItemFound(firstItemRemoved - 1);
			else if (firstItemRemoved == 0)
				SelectItemFound(0);
		}
		else
			m_lastSel = -1;

		// TODO: SelectItemFound sets  m_lastFindPos. Need to reset to 0 ?
		m_lastFindPos = -1;
		m_bEditFindFirst = TRUE;

		VerifyMailsOnUserSelectsMailListMarkCounts();
	}
	break;
	case S_COPY_ALL_Id:
	{
		CopyAllMails();
		VerifyMailsOnUserSelectsMailListMarkCounts();
	}
	break;
	case S_REMOVE_ALL_Id:
	{
		CString txt = L"Do you want to remove all mails?";
		ResHelper::TranslateString(txt);

		HWND h = GetSafeHwnd();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDYES) {
			RemoveAllMails();
		}
		VerifyMailsOnUserSelectsMailListMarkCounts();
	}
	break;
	case S_SAVE_AS_ARCHIVE_Id:
	{
		//SaveAsMboxAndAsMboxlistFile();
		SaveAsMboxArchiveFile_v2();
	}
	break;
	case S_SAVE_AS_LIST_Id:
	{
		SaveAsMboxListFile_v2();
	}
	break;
	case S_COPY_SELECTED_Id:
	{
		CopySelectedMails();
		VerifyMailsOnUserSelectsMailListMarkCounts();
	}
	break;
	case S_ADVANCED_FIND_Id:
	{
		RunFindAdvancedOnSelectedMail(iItem);
		MboxMail::ShowHint(HintConfig::AdvancedFindDialogHint, GetSafeHwnd());
	}
	break;
	case S_RELOAD_LIST_Id:
	{
		ReloadMboxListFile_v2();
		RefreshMailsOnUserSelectsMailListMark();
	}
	break;
	case S_RESTORE_LIST_Id:
	{
		ReloadMboxListFile_v2();
		RefreshMailsOnUserSelectsMailListMark();
	}
	break;
	case M_ARCHIVE_LOCATION_Id:
	{
		OpenArchiveFileLocation();
	}
	break;
	case M_LIST_LOCATION_Id:
	{
		OpenMailListFileLocation();
	}
	break;
	case M_REMOVE_DUPLICATE_MAILS_Id:
	{
		if (MboxMail::IsLabelMailsSelected())  // TODO: review please
		{
			RemoveDuplicateMails(MboxMail::s_mails_label);
			if (pFrame) {
				NTreeView *pTreeView = pFrame->GetTreeView();
				int retval = pTreeView->UpdateLabelMailListFile(0);
			}
			int deb = 1;
		}
		else
		{
			RemoveDuplicateMails(MboxMail::s_mails_edit);
			MboxMail::m_editMails.m_bIsDirty = TRUE;
			RefreshMailsOnUserSelectsMailListMark();
		}
	}
	break;
	case M_COPY_MAILS_TO_FOLDERS_Id:
	{
		int retval = CopyMailsToFolders();
	}
	break;
	case M_FORWARD_MAIL_Id:
	{
		//MboxMail::ShowHint(HintConfig::ForwardMailHint, GetSafeHwnd());
		ForwardSelectedMails(iItem);
		int deb = 1;
	}
	break;
	case M_FORWARD_RELATED_MAILS_Id:
	{
		//MboxMail::ShowHint(HintConfig::ForwardMailHint, GetSafeHwnd());
		int ret = ForwardMailRange(iItem);
		int deb = 1;
	}
	break;
	case M_EXPORT_SELECTED_MAILS_TO_SINGLE_FILE_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bMultipleMailsView)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		if (pFrame)
		{
			NListView::m_exportMailsMode = TRUE;

			BOOL bPrintToSeparateHTMLFiles = pFrame->m_NamePatternParams.m_bPrintToSeparateHTMLFiles;
			pFrame->m_NamePatternParams.m_bPrintToSeparateHTMLFiles = TRUE;
			//
			BOOL fullImgFilePath = NListView::m_fullImgFilePath;
			NListView::m_fullImgFilePath = FALSE;
			
			CString errorText;
			CString targetPrintSubFolderName;
			CString targetPrintFolderPath;
			MailIndexList* selectedMailsIndexList = 0;
			int ret = PrintMailSelectedToSingleHTML_Thread(selectedMailsIndexList, targetPrintSubFolderName, targetPrintFolderPath);

			pFrame->m_NamePatternParams.m_bPrintToSeparateHTMLFiles = bPrintToSeparateHTMLFiles;
			NListView::m_fullImgFilePath = fullImgFilePath;
			NListView::m_exportMailsMode = FALSE;

			int deb = 1;
		}

		NListView::m_appendAttachmentPictures = FALSE;
		int deb = 1;
	}
	break;
	case M_EXPORT_SELECTED_MAILS_CONFIG_Id:
	{
		BOOL sts = ConfigureExportOfMails();
	}
	break;
	case M_EXPORT_SELECTED_MAILS_HELP_Id:
	{
		CString helpFileName = L"MailExportHelp.pdf";
		HWND h = GetSafeHwnd();
		CMainFrame::OpenHelpFile(helpFileName, h);
	}
	break;
	default: {
		int deb = 1;
	}
	break;
	}

	// TODO: review below check; set itemSelected to TRUE ??  // FIXME
	if ((command == S_REMOVE_Id) || 
		(command == M_REMOVE_DUPLICATE_MAILS_Id) || 
		(command == S_REMOVE_ALL_Id) || 
		(command == S_RELOAD_LIST_Id) ||
		(command == S_RESTORE_LIST_Id) ||
		(command == S_ADVANCED_FIND_Id) ||
		(command == S_SAVE_AS_ARCHIVE_Id) ||
		(command == S_SAVE_AS_LIST_Id) ||
		(command == S_COPY_SELECTED_Id) ||
		(command == S_COPY_ALL_Id)
		)
		; // done
	else if ((itemSelected == FALSE) && (m_lastSel != iItem))
	{
		TRACE(L"Selecting %d\n", iItem);
		m_bApplyColorStyle = TRUE;
		SelectItem(iItem);
	}

	Invalidate();
	UpdateWindow();

	*pResult = 0;
}


void NListView::PrintToPDFMultipleSelect(int fontSize)
{
	m_fontSizePDF = fontSize;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		MboxMail::ShowHint(HintConfig::PrintToPDFHint, GetSafeHwnd());
		if (pFrame->m_NamePatternParams.m_bPrintToSeparatePDFFiles)
		{
			CString targetPrintSubFolderName = L"PDF_GROUP";
			CString targetPrintFolderPath;
			MailIndexList* selectedMailsIndexList = 0;
			int ret = PrintMailSelectedToSeparatePDF_Thread(selectedMailsIndexList, targetPrintSubFolderName, targetPrintFolderPath, FALSE);
		}
		else
		{
			CString targetPrintSubFolderName;
			CString targetPrintFolderPath;
			MailIndexList* selectedMailsIndexList = 0;
			int ret = PrintMailSelectedToSinglePDF_Thread(selectedMailsIndexList,targetPrintSubFolderName, targetPrintFolderPath);
			int deb = 1;
		}
		int deb = 1;
	}
	int deb = 1;
}

void NListView::OnRClickMultipleSelect(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMITEMACTIVATE pnm = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;

	int iItem = pnm->iItem;

	if (iItem < 0)
		return;

	CPoint pt;
	::GetCursorPos(&pt);
	CWnd *wnd = WindowFromPoint(pt);

	CMenu menu;
	menu.CreatePopupMenu();
	menu.AppendMenu(MF_SEPARATOR);

	CMenu printToSubMenu;
	printToSubMenu.CreatePopupMenu();
	printToSubMenu.AppendMenu(MF_SEPARATOR);

	CMenu printGroupToSubMenu;
	printGroupToSubMenu.CreatePopupMenu();
	printGroupToSubMenu.AppendMenu(MF_SEPARATOR);

	CMenu printPDFGroupToSubMenu;
	printPDFGroupToSubMenu.CreatePopupMenu();
	printPDFGroupToSubMenu.AppendMenu(MF_SEPARATOR);

	CMenu exportMailsToSubMenu;
	exportMailsToSubMenu.CreatePopupMenu();

	// Create enums or replace switch statment with if else ..

	const UINT S_TEXT_GROUP_Id = 4;
	MyAppendMenu(&printGroupToSubMenu, S_TEXT_GROUP_Id, L"Text..");

	const UINT S_HTML_GROUP_Id = 5;
	MyAppendMenu(&printGroupToSubMenu, S_HTML_GROUP_Id, L"HTML...");

	const UINT S_PRINTER_GROUP_Id = 6;
	MyAppendMenu(&printGroupToSubMenu, S_PRINTER_GROUP_Id, L"Printer...");

	const UINT S_PDF_DIRECT_Id = 28;
	//AppendMenu(&printGroupToSubMenu, S_PDF_DIRECT_Id, L"PDF...");  // FIXME

	const UINT S_PDF_DIRECT_DEFAULT_Id = 30;
	MyAppendMenu(&printPDFGroupToSubMenu, S_PDF_DIRECT_DEFAULT_Id, L"Default");

	// IDs 31,32, 33 available now
	const UINT S_PDF_DIRECT_FONT_16_Id = 31;
	const UINT S_PDF_DIRECT_FONT_20_Id = 32;
	const UINT S_PDF_DIRECT_FONT_24_Id = 33;

	const UINT S_PDF_DIRECT_MERGE_Id = 34;
	MyAppendMenu(&printPDFGroupToSubMenu, S_PDF_DIRECT_MERGE_Id, L"Merge");

	const UINT S_PDF_DIRECT_HELP_Id = 35;
	MyAppendMenu(&printPDFGroupToSubMenu, S_PDF_DIRECT_HELP_Id, L"Help");

	const UINT S_CSV_GROUP_Id = 29;
	MyAppendMenu(&printGroupToSubMenu, S_CSV_GROUP_Id, L"CSV...");

	menu.AppendMenu(MF_POPUP | MF_STRING, (UINT_PTR)printGroupToSubMenu.GetSafeHmenu(), L"Print Selected Mails To");
	menu.AppendMenu(MF_SEPARATOR);

	printGroupToSubMenu.AppendMenu(MF_POPUP | MF_STRING, (UINT_PTR)printPDFGroupToSubMenu.GetSafeHmenu(), L"PDF..");
	printGroupToSubMenu.AppendMenu(MF_SEPARATOR);

	const UINT S_HTML_OPEN_RELATED_Id = 7;
	MyAppendMenu(&menu, S_HTML_OPEN_RELATED_Id, L"Open Selected Mails in Browser");

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	const UINT S_REMOVE_SELECTED_Id = 8;
	if (pFrame && (MboxMail::IsFindMailsSelected() || MboxMail::IsUserMailsSelected())) {
		MyAppendMenu(&menu, S_REMOVE_SELECTED_Id, L"Remove Selected Mails");
	}

	const UINT S_COPY_SELECTED_Id = 22;
	if (pFrame && (MboxMail::IsAllMailsSelected() || MboxMail::IsFindMailsSelected() || MboxMail::IsLabelMailsSelected())) {
		if (pFrame->IsUserMailsListEnabled()) {
			MyAppendMenu(&menu, S_COPY_SELECTED_Id, L"Copy Selected into User Selected Mails");
		}
	}

	const UINT S_FORWARD_SELECTED_MAILS_Id = 23;
	MyAppendMenu(&menu, S_FORWARD_SELECTED_MAILS_Id, L"Forward Selected Mails");

	// Used above for printing to PDF and CSV
	//const UINT S_PDF_DIRECT_Id = 28;
	//const UINT S_CSV_GROUP_Id = 29;
	// 30-35 used see above

	/////////////
	const UINT S_REMOVE_ALL_Id = 40;
	const UINT S_SAVE_AS_ARCHIVE_Id = 41;
	const UINT M_ARCHIVE_LOCATION_Id = 42;

	if (pFrame && (MboxMail::IsFindMailsSelected() || (MboxMail::IsUserMailsSelected() && pFrame->IsUserMailsListEnabled())))
	{
		MyAppendMenu(&menu, S_REMOVE_ALL_Id, L"Remove All Mails");

		menu.AppendMenu(MF_SEPARATOR);
		MyAppendMenu(&menu, S_SAVE_AS_ARCHIVE_Id, L"Save All as Mbox Mail Archive file");
		MyAppendMenu(&menu, M_ARCHIVE_LOCATION_Id, L"Open Mbox Mail Archive Location");
	}

	const UINT S_SAVE_AS_LIST_Id = 43;
	const UINT S_RELOAD_LIST_Id = 44;
	const UINT M_LIST_LOCATION_Id = 45;
	if (pFrame && (MboxMail::IsUserMailsSelected() && pFrame->IsUserMailsListEnabled()))
	{
		menu.AppendMenu(MF_SEPARATOR);
		MyAppendMenu(&menu, S_SAVE_AS_LIST_Id, L"Save All as Mail List file");
		MyAppendMenu(&menu, S_RELOAD_LIST_Id, L"Reload from Mail List file");
		MyAppendMenu(&menu, M_LIST_LOCATION_Id, L"Open Mail List Location");
	}

	const UINT S_RESTORE_LIST_Id = 46;
	if (pFrame && MboxMail::IsAllMailsSelected() && pFrame->IsUserMailsListEnabled() && (MboxMail::s_mails_edit.GetCount() == 0))
	{
		//menu.AppendMenu(MF_SEPARATOR);
		MyAppendMenu(&menu, S_RESTORE_LIST_Id, L"Restore User Selected Mails from saved Mail List file");
	}

	const UINT S_EXPORT_SELECTED_MAILS_TO_SEPARATE_FILES_Id = 47;
	MyAppendMenu(&exportMailsToSubMenu, S_EXPORT_SELECTED_MAILS_TO_SEPARATE_FILES_Id, L"To Separate Files");

	const UINT S_EXPORT_SELECTED_MAILS_TO_SINGLE_FILE_Id = 48;
	MyAppendMenu(&exportMailsToSubMenu, S_EXPORT_SELECTED_MAILS_TO_SINGLE_FILE_Id, L"To Single File");

	const UINT S_EXPORT_SELECTED_MAILS_CONFIG_Id = 49;
	MyAppendMenu(&exportMailsToSubMenu, S_EXPORT_SELECTED_MAILS_CONFIG_Id, L"Config");

	const UINT S_EXPORT_SELECTED_MAILS_HELP_Id = 50;
	MyAppendMenu(&exportMailsToSubMenu, S_EXPORT_SELECTED_MAILS_HELP_Id, L"Help");

	menu.AppendMenu(MF_POPUP | MF_STRING, (UINT_PTR)exportMailsToSubMenu.GetSafeHmenu(), L"Export Selected Mails");
	menu.AppendMenu(MF_SEPARATOR);
	
	//////////////

	int index = 0;
	ResHelper::UpdateMenuItemsInfo(&menu, index);
	ResHelper::LoadMenuItemsInfo(&menu, index);
	
	UINT command = menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, this);

	UINT nFlags = TPM_RETURNCMD;
	CString menuString;
	int chrCnt = menu.GetMenuString(command, menuString, nFlags);

	BOOL multipleSelectedMails = TRUE;
	BOOL itemSelected = FALSE;
	m_fontSizePDF = 0;
	switch (command)
	{
	case S_HTML_GROUP_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bMultipleMailsView)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		if (pFrame)
		{
			if (pFrame->m_NamePatternParams.m_bPrintToSeparateHTMLFiles)
			{
				CString errorText;
				CString targetPrintSubFolderName = L"HTML_GROUP";
				CString targetPrintFolderPath;
				MailIndexList* selectedMailsIndexList = 0;
				int ret = PrintMailSelectedToSeparateHTML_Thread(selectedMailsIndexList, targetPrintSubFolderName, targetPrintFolderPath);
				int deb = 1;
			}
			else
			{
				CString errorText;
				CString targetPrintSubFolderName;
				CString targetPrintFolderPath;
				MailIndexList* selectedMailsIndexList = 0;
				int ret = PrintMailSelectedToSingleHTML_Thread(selectedMailsIndexList, targetPrintSubFolderName, targetPrintFolderPath);
				int deb = 1;
			}
		}

		NListView::m_appendAttachmentPictures = FALSE;
		int deb = 1;
	}
	break;
	case S_TEXT_GROUP_Id:
	{
		//PrintMailGroupToText(multipleSelectedMails, iItem, 0);
		if (pFrame)
		{
			if (pFrame->m_NamePatternParams.m_bPrintToSeparateTEXTFiles)
			{
				CString errorText;
				CString targetPrintSubFolderName = L"TEXT_GROUP";
				CString targetPrintFolderPath;
				//int ret = PrintMailSelectedToSeparateTEXT_Thread(targetPrintSubFolderName, targetPrintFolderPath);
				int deb = 1;
			}
			else
			{
				CString errorText;
				CString targetPrintSubFolderName;
				CString targetPrintFolderPath;
				int ret = PrintMailSelectedToSingleTEXT_Thread(targetPrintSubFolderName, targetPrintFolderPath);
				int deb = 1;
			}
		}
		int deb = 1;
	}
	break;
	case S_PRINTER_GROUP_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bMultipleMailsPrint)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		// IHTMLDocument require full path to embedded images
		// It is global for now FIXME
		m_fullImgFilePath = TRUE;

		MboxMail::ShowHint(HintConfig::PrintToPrinterHint, GetSafeHwnd());
		PrintMailGroupToText(multipleSelectedMails, iItem, 1, FALSE, TRUE);

		m_fullImgFilePath = m_fullImgFilePath_Config;
		NListView::m_appendAttachmentPictures = FALSE;
	}
	break;
	case S_PDF_DIRECT_DEFAULT_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bMultipleMailsPrint)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		PrintToPDFMultipleSelect(0);

		NListView::m_appendAttachmentPictures = FALSE;
	}
	break;
	case S_PDF_DIRECT_MERGE_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bMultipleMailsPrint)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		CString targetPrintSubFolderName = L"PDF_GROUP";
		CString targetPrintFolderPath;
		int ret = PrintMailSelectedToSeparatePDF_Thread(0, targetPrintSubFolderName, targetPrintFolderPath, TRUE);

		NListView::m_appendAttachmentPictures = FALSE;
	}
	break;
	case S_PDF_DIRECT_HELP_Id:
	{
		//CString helpFileName = "PrintMultipleMailsToSinglePDF.htm";
		CString helpFileName = L"PrintMultipleMailsToPDF.pdf";
		HWND h = GetSafeHwnd();
		CMainFrame::OpenHelpFile(helpFileName, h);
		int deb = 1;
	}
	break;
	case S_PDF_DIRECT_Id:  // dead code, remove
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bMultipleMailsPrint)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		if (pFrame)
		{
			MboxMail::ShowHint(HintConfig::PrintToPDFHint, GetSafeHwnd());
			if (pFrame->m_NamePatternParams.m_bPrintToSeparatePDFFiles)
			{
				CString targetPrintSubFolderName = L"PDF_GROUP";
				CString targetPrintFolderPath;
				int ret = PrintMailSelectedToSeparatePDF_Thread(0, targetPrintSubFolderName, targetPrintFolderPath, FALSE);
			}
			else 
			{
				CString targetPrintSubFolderName;
				CString targetPrintFolderPath;
				int ret = PrintMailSelectedToSinglePDF_Thread(0, targetPrintSubFolderName, targetPrintFolderPath);
				int deb = 1;
			}
			int deb = 1;
		}

		NListView::m_appendAttachmentPictures = FALSE;
		int deb = 1;
	}
	break;
	case S_HTML_OPEN_RELATED_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bMultipleMailsView)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		BOOL forceOpen = TRUE;
		PrintMailGroupToText(multipleSelectedMails, iItem, 1, forceOpen);

		NListView::m_appendAttachmentPictures = FALSE;
	}
	break;

	case S_REMOVE_SELECTED_Id: {
		int firstItemRemoved = RemoveSelectedMails();
		if ((firstItemRemoved > 0) && MboxMail::s_mails.GetCount())
			SelectItemFound(firstItemRemoved - 1);
		else if ((firstItemRemoved == 0) && MboxMail::s_mails.GetCount())
			SelectItemFound(0);

		// TODO: SelectItemFound sets m_lastFindPos. Need to reset to 0 ?
		m_lastFindPos = -1;
		m_bEditFindFirst = TRUE;
		VerifyMailsOnUserSelectsMailListMarkCounts();
	}
	break;
	case S_COPY_SELECTED_Id:
	{
		CopySelectedMails();
		VerifyMailsOnUserSelectsMailListMarkCounts();
	}
	break;
	case S_CSV_GROUP_Id:
	{
		if (pFrame) {
			BOOL selectedMails = TRUE;
			int firstMail = 0;
			int lastMail = MboxMail::s_mails.GetCount() - 1;
			pFrame->PrintMailsToCSV(firstMail, lastMail, selectedMails);
		}
		int deb = 1;
	}
	break;
	case S_FORWARD_SELECTED_MAILS_Id:
	{
		//MboxMail::ShowHint(HintConfig::ForwardMailHint, GetSafeHwnd());
		ForwardSelectedMails(iItem);
		int deb = 1;
	}
	break;
	////// 
	case S_REMOVE_ALL_Id:
	{
		CString txt = L"Do you want to remove all mails?";
		ResHelper::TranslateString(txt);

		HWND h = GetSafeHwnd();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDYES) {
			RemoveAllMails();
		}
		VerifyMailsOnUserSelectsMailListMarkCounts();
	}
	break;
	case S_SAVE_AS_ARCHIVE_Id:
	{
		//SaveAsMboxAndAsMboxlistFile();
		SaveAsMboxArchiveFile_v2();
	}
	break;
	case S_SAVE_AS_LIST_Id:
	{
		SaveAsMboxListFile_v2();
	}
	break;
#if 0
	case S_ADVANCED_FIND_Id: {
		RunFindAdvancedOnSelectedMail(iItem);
		MboxMail::ShowHint(HintConfig::AdvancedFindDialogHint, GetSafeHwnd());
	}
	break;
#endif
	case S_RELOAD_LIST_Id:
	{
		ReloadMboxListFile_v2();
		RefreshMailsOnUserSelectsMailListMark();
	}
	break;
	case S_RESTORE_LIST_Id:
	{
		ReloadMboxListFile_v2();
		RefreshMailsOnUserSelectsMailListMark();
	}
	break;
	case M_ARCHIVE_LOCATION_Id:
	{
		OpenArchiveFileLocation();
	}
	break;
	case M_LIST_LOCATION_Id:
	{
		OpenMailListFileLocation();
	}
	break;

	case S_EXPORT_SELECTED_MAILS_TO_SEPARATE_FILES_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bMultipleMailsView)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		if (pFrame)
		{
			NListView::m_exportMailsMode = TRUE;

			BOOL bPrintToSeparateHTMLFiles = pFrame->m_NamePatternParams.m_bPrintToSeparateHTMLFiles;
			pFrame->m_NamePatternParams.m_bPrintToSeparateHTMLFiles = TRUE;
			//
			BOOL fullImgFilePath = NListView::m_fullImgFilePath;
			NListView::m_fullImgFilePath = FALSE;

			CString errorText;
			CString targetPrintSubFolderName = L"Mails";
			CString targetPrintFolderPath;
			MailIndexList* selectedMailsIndexList = 0;
			int ret = PrintMailSelectedToSeparateHTML_Thread(selectedMailsIndexList, targetPrintSubFolderName, targetPrintFolderPath);

			pFrame->m_NamePatternParams.m_bPrintToSeparateHTMLFiles = bPrintToSeparateHTMLFiles;
			NListView::m_fullImgFilePath = fullImgFilePath;
			NListView::m_exportMailsMode = FALSE;

			int deb = 1;
		}

		NListView::m_appendAttachmentPictures = FALSE;
		int deb = 1;
	}
	break;
	case S_EXPORT_SELECTED_MAILS_TO_SINGLE_FILE_Id:
	{
		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams)
		{
			if (attachmentConfigParams->m_bMultipleMailsView)
				NListView::m_appendAttachmentPictures = TRUE;
		}

		if (pFrame)
		{
			NListView::m_exportMailsMode = TRUE;

			BOOL bPrintToSeparateHTMLFiles = pFrame->m_NamePatternParams.m_bPrintToSeparateHTMLFiles;
			pFrame->m_NamePatternParams.m_bPrintToSeparateHTMLFiles = TRUE;
			//
			BOOL fullImgFilePath = NListView::m_fullImgFilePath;
			NListView::m_fullImgFilePath = FALSE;

			CString errorText;
			CString targetPrintSubFolderName;
			CString targetPrintFolderPath;
			MailIndexList* selectedMailsIndexList = 0;
			int ret = PrintMailSelectedToSingleHTML_Thread(selectedMailsIndexList, targetPrintSubFolderName, targetPrintFolderPath);

			pFrame->m_NamePatternParams.m_bPrintToSeparateHTMLFiles = bPrintToSeparateHTMLFiles;
			NListView::m_fullImgFilePath = fullImgFilePath;
			NListView::m_exportMailsMode = FALSE;

			int deb = 1;
		}

		NListView::m_appendAttachmentPictures = FALSE;
		int deb = 1;
	}
	break;
	case S_EXPORT_SELECTED_MAILS_CONFIG_Id:
	{
		BOOL sts = ConfigureExportOfMails();
	}
	break;
	case S_EXPORT_SELECTED_MAILS_HELP_Id:
	{
		CString helpFileName = L"MailExportHelp.pdf";
		HWND h = GetSafeHwnd();
		CMainFrame::OpenHelpFile(helpFileName, h);
	}
	break;
	default:
	{
		int deb = 1;
	}
	break;
	}

	Invalidate();
	UpdateWindow();

	*pResult = 0;
}

void NListView::OnColumnClick(NMHDR* pNMHDR, LRESULT* pResult)
{
	static BOOL enlarge = TRUE;
	NMLISTVIEW * pLV = (NMLISTVIEW*)pNMHDR;

	int iSubItem = pLV->iSubItem;

	*pResult = 0;
	
	SortByColumn(pLV->iSubItem);
}

void NListView::SortBySubjectBasedConversasions(BOOL justRefresh)
{
	if (justRefresh == FALSE)
	{
		if (abs(MboxMail::b_mails_which_sorted) == 4) {
			MboxMail::b_mails_which_sorted = -MboxMail::b_mails_which_sorted;
		}
		else {
			MboxMail::b_mails_which_sorted = 4;
		}
	}

	MboxMail::SortBySubjectBasedConversasions(0, MboxMail::b_mails_which_sorted < 0);

	MboxMail::b_mails_sorted = true;

	m_findParams.m_bContent = FALSE;
	m_findParams.m_bAttachments = FALSE;

	m_lastFindPos = -1;
	m_bEditFindFirst = TRUE;  // must call EditFind()

	RedrawMails();

	return;
}

void NListView::SortByColumn(int colNumber, BOOL sortByPosition) // use sortByPosition for testing
{
	bool mustSort = false;

	// to follow OnEditFindAgain approach
	if (m_bInFind)  
		return;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		int paneId = 0;
		CString colStr;
		CString sText;
		if (colNumber == 0)
			colStr = L"Conversations";
		else if (colNumber == 1)
			colStr = L"Date";
		else if (colNumber == 2)
			colStr = L"From";
		else if (colNumber == 3)
			colStr = L"To";
		else if (colNumber == 4)
			colStr = L"Subject";
		else if (colNumber == 5)
			colStr = L"Size";

		if ((colNumber == 1) && sortByPosition)  // Bad hack to sort by Id-:)  FIX IT
		{
			colStr = L"Id";
		}
		CString fmt = L"Sorting by %s ...";
		ResHelper::TranslateString(fmt);

		sText.Format(fmt, colStr);
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);
	}

	m_bInFind = true;

	if (colNumber == 0) // special case. Can't make negative 0
	{
		if (abs(MboxMail::b_mails_which_sorted) == 99) {
			MboxMail::b_mails_which_sorted = -MboxMail::b_mails_which_sorted;
		}
		else {
			MboxMail::b_mails_which_sorted = 99;
		}
	}
	else
	{
		if (abs(MboxMail::b_mails_which_sorted) == colNumber) {
			MboxMail::b_mails_which_sorted = -MboxMail::b_mails_which_sorted;
		}
		else {
			MboxMail::b_mails_which_sorted = colNumber;
		}
	}

	switch (colNumber) {
	case 0: // !
	{
		MboxMail::SortByGroupId(0, MboxMail::b_mails_which_sorted < 0);
		MboxMail::assignColor2ConvesationGroups(&MboxMail::s_mails);
		//MboxMail::SortByConverstionGroups(0, MboxMail::b_mails_which_sorted < 0);  // TODO: review
		mustSort = true;
	}
		break;
	case 1: // date
	{
		if (sortByPosition == FALSE)
			MboxMail::SortByDate(0, MboxMail::b_mails_which_sorted < 0);
		else
			MboxMail::SortByFileOffset(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
	}
		break;
	case 2: // from
	{
		MboxMail::SortByFrom(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
	}
		break;
	case 3: // to
	{
		MboxMail::SortByTo(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
	}
		break;
	case 4: // subj
	{
		HWND h = GetSafeHwnd();
		MboxMail::ShowHint(HintConfig::SubjectSortingHint, h);

		if (m_subjectSortType == 0)
			MboxMail::SortBySubject(0, MboxMail::b_mails_which_sorted < 0);
		else
			MboxMail::SortBySubjectBasedConversasions(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
	}
		break;
	case 5: // size
	{
		MboxMail::SortBySize(0, MboxMail::b_mails_which_sorted > 0);
		mustSort = true;
	}
		break;
	}

	if (mustSort)
	{
		MboxMail::b_mails_sorted = true;

		m_findParams.m_bContent = FALSE;
		m_findParams.m_bAttachments = FALSE;

		m_lastFindPos = -1;
		m_bEditFindFirst = TRUE;  // must call EditFind()

		// Don't reset below vars. It helps user to keep these while searching and sorting
		//m_lastStartDate = 0;
		//m_lastEndDate = 0;
		//m_searchString.Empty();


		//  sets m_lastSel to -1 via ClearDescView
		RedrawMails();
		// MessageBeep(MB_OK); // too much ??
	}

	m_bInFind = false;

	if (pFrame)
	{
		int paneId = 0;
		CString sText;
		sText.Format(L"Ready");
		ResHelper::TranslateString(sText);
		pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
	}
}


// Refresh mail list, don't update MboxMail::b_mails_which_sorted
// Currently not called. FIXME 
void NListView::RefreshSortByColumn()
{
	bool mustSort = false;

	// to follow OnEditFindAgain approach
	if (m_bInFind)
		return;

	m_bInFind = true;

	int colNumber;
	if (abs(MboxMail::b_mails_which_sorted) == 99)
		colNumber = 0;
	else
		colNumber = abs(MboxMail::b_mails_which_sorted);

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		int paneId = 0;
		CString colStr;
		CString sText;
		if (colNumber == 0)
			colStr = L"Conversations";
		else if (colNumber == 1)
			colStr = L"Date";
		else if (colNumber == 2)
			colStr = L"From";
		else if (colNumber == 3)
			colStr = L"To";
		else if (colNumber == 4)
			colStr = L"Subject";
		else if (colNumber == 5)
			colStr = L"Size";

		CString fmt = L"Sorting by %s ...";
		ResHelper::TranslateString(fmt);

		sText.Format(fmt, colStr);
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);
	}

	switch (colNumber) {
	case 0: // !
		MboxMail::SortByGroupId(0, MboxMail::b_mails_which_sorted < 0);
		MboxMail::assignColor2ConvesationGroups(&MboxMail::s_mails);
		//MboxMail::SortByConverstionGroups(0, MboxMail::b_mails_which_sorted < 0); // TODO: review
		mustSort = true;
		break;
	case 1: // date
		MboxMail::SortByDate(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
		break;
	case 2: // from
		MboxMail::SortByFrom(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
		break;
	case 3: // to
		MboxMail::SortByTo(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
		break;
	case 4: // subj
		if (m_subjectSortType == 0)
			MboxMail::SortBySubject(0, MboxMail::b_mails_which_sorted < 0);
		else
			MboxMail::SortBySubjectBasedConversasions(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
		break;
	case 5: // size
		MboxMail::SortBySize(0, MboxMail::b_mails_which_sorted > 0);
		mustSort = true;
		break;
	}
	if (mustSort)
	{
		MboxMail::b_mails_sorted = true;

		m_findParams.m_bContent = FALSE;
		m_findParams.m_bAttachments = FALSE;

		m_lastFindPos = -1;
		m_bEditFindFirst = TRUE;  // must call EditFind()

		// Don't reset below vars. It helps user to keep these while searching and sorting
		//m_lastStartDate = 0;
		//m_lastEndDate = 0;
		//m_searchString.Empty();

		//  sets m_lastSel to -1 via ClearDescView
		RedrawMails();
		// MessageBeep(MB_OK); // too much ??
	}

	m_bInFind = false;

	if (pFrame)
	{
		int paneId = 0;
		CString sText;
		sText.Format(L"Ready");
		ResHelper::TranslateString(sText);
		pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
	}
}

void NListView::OnKeydown(NMHDR* pNMHDR, LRESULT* pResult) 
{
//	m_headerCtrl.PostMessage(WM_CANCELMODE);
	NMLVKEYDOWN* pLVKeyDow = (NMLVKEYDOWN*)pNMHDR;
	*pResult = 0;

	int whichSel;
	m_bApplyColorStyle = TRUE;
	switch( pLVKeyDow->wVKey ) {
	case VK_HOME:
		SelectItem(0);
		break;
	case VK_END:
		SelectItem(m_list.GetItemCount()-1);
		break;
	case VK_UP:
		if( m_lastSel )
			SelectItem(m_lastSel-1);
		break;
	case VK_DOWN:
		if( m_lastSel < m_list.GetItemCount()-1 )
			SelectItem(m_lastSel+1);
		break;
	case VK_NEXT:
		{
			int curBottom = m_list.GetTopIndex() + m_list.GetCountPerPage() - 1;
			if( m_lastSel >= curBottom )
				whichSel = m_lastSel + m_list.GetCountPerPage() - 1;
			else
				whichSel = curBottom;
		}
		if( whichSel >= m_list.GetItemCount() )
			whichSel = m_list.GetItemCount()-1;
		SelectItem(whichSel);
		break;
	case VK_PRIOR:
		{
			int curTop = m_list.GetTopIndex();
			if( m_lastSel <= curTop ) {
				if( m_lastSel > m_list.GetCountPerPage() )
					whichSel = m_lastSel - m_list.GetCountPerPage();
				else
					whichSel = 0;
			} else
				whichSel = curTop;
		}
		SelectItem(whichSel);
		break;
	case VK_CONTROL:
	{
		//if (m_lastSel) SelectItem(m_lastSel);
		int deb = 1;
	}
	break;
	case VK_ESCAPE:
	{
		if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
		{
			AfxGetMainWnd()->PostMessage(WM_CLOSE);
		}
	}
	break;

	case 'A':
	case 'a':
	{
		SHORT ctrlState = GetAsyncKeyState(VK_CONTROL);
		SHORT AState = GetAsyncKeyState('A');
		BOOL CtrlKeyDown = FALSE;
		BOOL AKeyDown = FALSE;
		if ((ctrlState & 0x8000) != 0) {
			CtrlKeyDown = TRUE;
		}
			
		if ((AState & 0x8000) != 0) {
			AKeyDown = TRUE;
		}

		if (AKeyDown && CtrlKeyDown) {
			m_list.SetItemState(-1, LVIS_SELECTED, LVIS_SELECTED);
		}

		int deb = 1;
	}
	break;

	case 'N':
	case 'n':
	{
		SHORT ctrlState = GetAsyncKeyState(VK_CONTROL);
		SHORT AState = GetAsyncKeyState('N');
		BOOL CtrlKeyDown = FALSE;
		BOOL AKeyDown = FALSE;
		if ((ctrlState & 0x8000) != 0) {
			CtrlKeyDown = TRUE;
		}

		if ((AState & 0x8000) != 0) {
			AKeyDown = TRUE;
		}

		if (AKeyDown && CtrlKeyDown)
		{
			CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
			if (pFrame)
			{
				NMsgView* m_pMsgView = pFrame->DetMsgView();
				if (m_pMsgView)
				{
					LRESULT lres = m_pMsgView->m_browser.PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
					int deb = 1;
				}
			}
		}
		int deb = 1;
	}
	break;

	default:
	{
		SHORT ctrlState = GetAsyncKeyState(VK_CONTROL);
		BOOL CtrlKeyDown = FALSE;
		if ((ctrlState & 0x8000) != 0) {
			CtrlKeyDown = TRUE;
		}
		int deb = 1;
	}

/*	case VK_BACK:
	case VK_DELETE:
		if( theApp.ShowHidingDlg(GetConfirmDelete(m_listCtrl.GetSelectedCount()), g_hideDeletePopup, "hideDeletePopup", TRUE) ) {
			CWaitCursor cursor;
			DeleteSelectedArticle();
		}
		break;
	*/
	}
}

/////////////////////////////////////////////////////////////////////////////
// NListView message handlers

int NListView::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CWnd::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	//if( !m_list.Create(WS_CHILD|WS_VISIBLE|LVS_SINGLESEL|LVS_REPORT|LVS_SHOWSELALWAYS|LVS_OWNERDATA, CRect(), this, IDC_LIST) )
	if (!m_list.Create(WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_OWNERDATA , CRect(), this, IDC_LIST))
		return -1;


	DWORD exStyle = m_list.GetExtendedStyle();
	m_list.SetExtendedStyle(exStyle|LVS_EX_FULLROWSELECT);
#if 0
	// // TVS_INFOTIP doesn't seem to work with LS_OWNERDATA I believe
	exStyle = m_list.GetExtendedStyle();
	m_list.SetExtendedStyle(exStyle | LVS_EX_LABELTIP );
	//m_list.SetExtendedStyle(exStyle | LVS_EX_LABELTIP | LVS_EX_INFOTIP);

	//SetWindowTheme(m_list.m_hWnd, L"Explorer", NULL);
	//exStyle = m_list.GetExtendedStyle();
	//m_list.SetExtendedStyle(exStyle | LVS_EX_DOUBLEBUFFER);
#endif

	m_list.SendMessage((CCM_FIRST + 0x7), 5, 0); // #define CCM_SETVERSION          (CCM_FIRST + 0x7)
	m_list.SetTextColor (::GetSysColor(COLOR_WINDOWTEXT));
	ResetFont();

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if ((pFrame == 0) || (pFrame->GetMessageWindowPosition() == 1)) // m_msgViewPosition:: 1 = bottom 2 = right  3 = left
	{
		m_list.InsertColumn(0, L"!", LVCFMT_LEFT, 22, 0);
		m_list.InsertColumn(1, L"date", LVCFMT_LEFT, 100, 0);
		m_list.InsertColumn(2, L"from", LVCFMT_LEFT, 150, 0);
		m_list.InsertColumn(3, L"to", LVCFMT_LEFT, 150, 0);
		m_list.InsertColumn(4, L"subject", LVCFMT_LEFT, 400, 0);
		m_list.InsertColumn(5, L"size(KB)", LVCFMT_LEFT, 120, 0);
	}
	else
	{
		m_list.InsertColumn(0, L"!", LVCFMT_LEFT, 22, 0);
		m_list.InsertColumn(1, L"date", LVCFMT_LEFT, 100, 0);
		m_list.InsertColumn(2, L"from", LVCFMT_LEFT, 100, 0);
		m_list.InsertColumn(3, L"to", LVCFMT_LEFT, 100, 0);
		m_list.InsertColumn(4, L"subject", LVCFMT_LEFT, 400, 0);
		m_list.InsertColumn(5, L"size(KB)", LVCFMT_LEFT, 120, 0);
	}

#if 0
	// informational
	CRect rc_total_window;
	GetWindowRect(&rc_total_window);

	int nTextWidth = m_list.GetColumnWidth(0);
	nTextWidth = m_list.GetColumnWidth(1);
	nTextWidth = m_list.GetColumnWidth(2);
	nTextWidth = m_list.GetColumnWidth(3);
	nTextWidth = m_list.GetColumnWidth(4);
	nTextWidth = m_list.GetColumnWidth(5);
#endif

	// Take ownership of header draw
	SetLabelOwnership();

	m_nIDEvent = 1;
	m_nElapse = 1000;
	//this->SetTimer(m_nIDEvent, m_nElapse, NULL);

#if 0
	// Create and attach image list
	// TODO: finish or remove
	m_ImageList.Create(16, 16, ILC_COLOR16 | ILC_MASK, 1, 0);
	m_ImageList.Add(AfxGetApp()->LoadIcon(IDI_ICON_CHECK_ON));
	m_ImageList.Add(AfxGetApp()->LoadIcon(IDI_ICON_CHECK_OFF));
	m_list.SetImageList(&m_ImageList, LVSIL_SMALL);
#endif

	m_showAsPaperClip = TRUE;
	BOOL aRt = m_paperClip.LoadBitmap(IDB_PAPER_CLIP);
#if 0
	BOOL cr = m_imgList.Create(6, 10, ILC_COLOR4, 1, 1);
	int ar = m_imgList.Add(&m_paperClip, RGB(0, 0, 0));
#endif

#if 0
	// Doen't work, incomplete, investigate
	//EnableToolTips(TRUE);

	TRACKMOUSEEVENT tr;
	tr.cbSize = sizeof(tr);
	tr.dwFlags = TME_HOVER;
	tr.dwHoverTime = 400;
	tr.hwndTrack = this->m_hWnd;

	BOOL ret = TrackMouseEvent(&tr);
#endif
	//BOOL retA = ResHelper::ActivateToolTips(this, m_toolTip);  // doesn't work due to CustomDraw ? investigate when 
	//EnableToolTips();
	return 0;
}

void NListView::SetLabelOwnership()
{
	CHeaderCtrl* lhdr = m_list.GetHeaderCtrl();
	if (lhdr) {
		int nClmCnt = lhdr->GetItemCount();
		HDITEM rHdr;
		for (int j = 0; j < nClmCnt; j++) {
			rHdr.mask = HDI_FORMAT;
			lhdr->GetItem(j, &rHdr);
			rHdr.fmt |= HDF_OWNERDRAW;  // TODO: seem to also work without merging fmt options
			lhdr->SetItem(j, &rHdr);
		}
	}
	else
		int deb = 1;
}

void NListView::OnSize(UINT nType, int cx, int cy)
{
	CWnd::OnSize(nType, cx, cy);

	ResizeColumns();

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		int msgViewPosition = pFrame->MsgViewPosition();
		BOOL bTreeHideVal = pFrame->TreeHideValue();
		BOOL isTreeHidden = pFrame->IsTreeHidden();

		TRACE(L"NListView::OnSize cx=%d cy=%d viewPos=%d IsTreeHideVal=%d IsTreeHidden=%d\n",
			cx, cy, msgViewPosition, bTreeHideVal, isTreeHidden);

#if 0
		// No lnger needed, delete.
		if (!pFrame->m_bIsTreeHidden)
		{
			m_frameCx_TreeNotInHide = cx;
			m_frameCy_TreeNotInHide = cy;
		}
		else
		{
			m_frameCx_TreeInHide = cx;
			m_frameCy_TreeInHide = cy;
		}
#endif
	}

	m_list.MoveWindow(0, 0, cx, cy);
}

void NListView::ResizeColumns()
{
	CRect rc;
	GetWindowRect(&rc);
	int w = rc.Width();
	int sb_width = GetSystemMetrics(SM_CXVSCROLL);
	//w -= sb_width + 6;
	w -= sb_width;

	int col_zero_len = 22;
	int date_len = 100;
	int min_from_len = 150;
	int max_from_len = 400;
	int min_to_len = 150;
	int max_to_len = 400;
	int min_subj_len = 200;
	int dflt_subj_len = 400;
	int size_len = 80;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if ((pFrame != 0) || (pFrame->GetMessageWindowPosition() != 1))
	{
		min_from_len = 100;
		max_from_len = 400;
		min_to_len = 100;
		max_to_len = 400;
		min_subj_len = 150;
		dflt_subj_len = 400;
		size_len = 80;
	}

	int from_len = min_from_len;
	int to_len = min_to_len;
	int subj_len = min_subj_len;

	int min_len = col_zero_len + date_len + from_len + to_len + subj_len + size_len;
	int extra_space = w - min_len;


	if (extra_space > 0) 
	{
		if (extra_space < (dflt_subj_len - min_subj_len)) {
			subj_len += extra_space;
		}
		else 
		{
			subj_len = dflt_subj_len;
			int total_len = col_zero_len + date_len + from_len + to_len + subj_len + size_len;
			int space_left = w - total_len; // divide across from, to and subject

			from_len += (int)(space_left * 0.2);
			to_len += (int)(space_left * 0.2);
			subj_len = (int)(space_left * 0.6);
			if (from_len >= max_from_len)
				from_len = max_from_len;
			if (to_len >= max_to_len)
				to_len = max_to_len;

			total_len = col_zero_len + date_len + from_len + to_len + subj_len + size_len;
			space_left = w - total_len;
			subj_len += space_left;
		}
	}

	m_list.SetColumnWidth(1, date_len);
	m_list.SetColumnWidth(2, from_len);
	m_list.SetColumnWidth(3, to_len);
	m_list.SetColumnWidth(4, subj_len);
	m_list.SetColumnWidth(5, size_len);

	// no redraw seem to be needed
}

void NListView::ResetSize()
{
	MboxMail::s_fSize = 0;
	m_list.SetItemCount(0);
	m_list.Invalidate();
}

void NListView::OnCustomDraw(NMHDR* pNMHDR, LRESULT* pResult)
{
#define MarkColor	RGB(220, 20, 60)

	LPNMLVCUSTOMDRAW lplvcd = (LPNMLVCUSTOMDRAW)pNMHDR;
	NMCUSTOMDRAW   &nmcd = lplvcd->nmcd;
	MboxMail *m;
	DWORD error;

	DWORD_PTR dwItemSpec = nmcd.dwItemSpec;
	int iSubItem = lplvcd->iSubItem;
	DWORD dwItemType = lplvcd->dwItemType;
	DWORD dwDrawStage = nmcd.dwDrawStage;
	UINT uItemState = nmcd.uItemState;

	NMsgView *pMsgView = 0;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame) {
		pMsgView = pFrame->GetMsgView();
	}

	*pResult = 0;

	DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailSummaryList);

	switch (lplvcd->nmcd.dwDrawStage)
	{
	case CDDS_PREPAINT:	// Request prepaint notifications for each item.
	{
		*pResult = CDRF_NOTIFYITEMDRAW;
	}
	break;
	case CDDS_ITEMPREPAINT: //Before an item is drawn
	{
		*pResult = CDRF_NOTIFYSUBITEMDRAW;
	}
	break;
	case CDDS_ITEMPOSTPAINT | CDDS_SUBITEM:
	{
		int deb = 1;  // should not be here
	}
	break;
	case CDDS_ITEMPOSTPAINT:
	{
		int deb = 1; // should not be here
	}
	break;
	case CDDS_ITEMPREPAINT | CDDS_SUBITEM:
	{
		// Need this fix because when item is selected system overrides any user defined drawing

		//strW = L'\u2713';  // check mark
		//strW = L'\u2714';  // check mark
		//strW = L'\u1F4CE';  // paperclip but doesn't seem to be supported as wide character
		// UTF - 16 Encoding:	0xD83D 0xDCCE
		// UTF-8 Encoding:	0xF0 0x9F 0x93 0x8E
		// UTF-32 Encoding:	0x0001F4CE

		const char * ast = "*";
		const char * nul = "";

		LVITEM rItem;
		ZeroMemory(&rItem, sizeof(LVITEM));
		rItem.mask = LVIF_IMAGE | LVIF_STATE;
		rItem.iItem = (int)dwItemSpec;
		rItem.stateMask = LVIS_SELECTED;
		m_list.GetItem(&rItem);

		if ((lplvcd->nmcd.dwItemSpec < 0) || (lplvcd->nmcd.dwItemSpec > MboxMail::s_mails.GetCount()))
			int deb = 1;

		m = MboxMail::s_mails[lplvcd->nmcd.dwItemSpec];

		CDC dc;
		CRect rect = nmcd.rc;
		HDC hDC = lplvcd->nmcd.hdc;
		CStringW strW;
		CStringW txtW;
		int xpos;
		int ypos;
		CStringA FieldText;
		CStringA Charset;
		UINT charsetId;

		BOOL hasAttachents = FALSE;
		if (m->m_hasAttachments || HasAnyAttachment(m))
			hasAttachents = TRUE;

		DWORD bkcolor = ::GetSysColor(COLOR_HIGHLIGHT);
		DWORD txcolor = ::GetSysColor(COLOR_HIGHLIGHTTEXT);

		//if ((iSubItem == 0) && (abs(MboxMail::b_mails_which_sorted) == 99))
		if (iSubItem == 0)
		{
			CRect rect0 = rect;
			rect0.left -= 6;  // Why ?? Please alwas add comment
			if (dc.Attach(hDC))
			{
				if ((abs(MboxMail::b_mails_which_sorted) == 99) || ((abs(MboxMail::b_mails_which_sorted) == 4) && (m_subjectSortType == 1)))
				{
					if (m->m_groupColor == 0)
					{
						DWORD color1 = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailConversion1);
						dc.FillRect(&rect0, &CBrush(color1));
					}
					else
					{
						DWORD color2 = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailConversion2);
						dc.FillRect(&rect0, &CBrush(color2));
					}

					dc.SetBkMode(TRANSPARENT);
					dc.SetTextColor(RGB(0, 0, 0));
				}
				else
				{
					if (rItem.state & LVIS_SELECTED)
					{
						dc.FillRect(&rect0, &CBrush(bkcolor));
						dc.SetBkMode(TRANSPARENT);
						dc.SetTextColor(txcolor);
					}
					else
					{
						dc.FillRect(&rect0, &CBrush(color));
						dc.SetBkMode(TRANSPARENT);
						dc.SetTextColor(RGB(0, 0, 0));
					}
				}

				if ((m->m_isOnUserSelectedMailList) && !MboxMail::IsUserMailsSelected())
				{
					CPen penRed(PS_SOLID, 2, MarkColor);  // 
					CPen *oldPen = dc.SelectObject(&penRed);

					int xfrom = rect.left + 2;
					int yfrom = rect.top + 4;

					int xto = rect.left + 2;
					int yto = rect.bottom - 6;

					dc.MoveTo(xfrom, yfrom);
					dc.LineTo(xto, yto);

					CPen *pen = dc.SelectObject(oldPen);
				}

				if (hasAttachents && m_showAsPaperClip)
				{
					BITMAP  bi;
					BOOL gbr = m_paperClip.GetBitmap(&bi);

					int x = rect.left + 8;
					int y = rect.top + 4;
					// both below solutions work fine
#if 0
					CPoint pt(x, y);
					CSize sz(bi.bmWidth, bi.bmHeight);
					UINT fStyle = ILD_NORMAL;
					DWORD dwRop = SRCAND;
					BOOL r = m_imgList.DrawIndirect(&dc, 0, pt, sz, ptOrigin, fStyle, dwRop);
#else
					CDC bmDC;
					BOOL ccr = bmDC.CreateCompatibleDC(&dc);
					CBitmap *pOldbmp = bmDC.SelectObject(&m_paperClip);
					BOOL bbr = dc.BitBlt(x, y, bi.bmWidth, bi.bmHeight, &bmDC, 0, 0, SRCAND);
					CBitmap *bmap = bmDC.SelectObject(pOldbmp);
#endif
				}
				dc.Detach();
			}

			if (hasAttachents && !m_showAsPaperClip)
			{
				int x_offset = 4;
				xpos = rect.left + x_offset;
				ypos = rect.top + 3;

				FieldText = ast;

				Charset = "UTF-8";
				charsetId = CP_UTF8;

				strW.Empty();
				if (TextUtilsEx::Str2WStr(FieldText, charsetId, strW, error))
				{
					rect.InflateRect(-4, -1);
					int r = ::DrawTextExW(hDC, (LPWSTR)((LPCWSTR)strW), strW.GetLength(), &rect, DT_END_ELLIPSIS | DT_SINGLELINE, NULL);
				}
			}
		}
		else if ((iSubItem >= 1) && (iSubItem <= 5))
		{
			if (iSubItem == 1)
			{
				char datebuff[32];
				datebuff[0] = 0;
				if (m->m_timeDate >= 0)
				{
					MyCTime tt(m->m_timeDate);
					if (!m_gmtTime)
					{
						CStringA lDateTime = tt.FormatLocalTm(m_format);
						strcpy(datebuff, (LPCSTR)lDateTime);
					}
					else {
						CStringA lDateTime = tt.FormatGmtTm(m_format);
						strcpy(datebuff, (LPCSTR)lDateTime);
					}
				}
				FieldText = datebuff;
				Charset = "UTF-8";
				charsetId = CP_UTF8;
			}
			else if (iSubItem == 2)
			{
				Charset = m->m_from_charset;
				charsetId = m->m_from_charsetId;

				if ((charsetId == 0) && pMsgView) {
					if (pMsgView->m_cnf_from_charsetId)
						charsetId = pMsgView->m_cnf_from_charsetId;
					else
						charsetId = CP_UTF8;

					//Charset = Charset;
				}

				if (m_bLongMailAddress) {
					FieldText = m->m_from;
				}
				else
				{
					int fromlen = m->m_from.GetLength();
					m_name->ClearAndResize(fromlen);
					m_addr->ClearAndResize(fromlen);
					MboxMail::splitMailAddress(m->m_from, fromlen, m_name, m_addr);
					FieldText.Empty();
					if (m_name->Count()) {
						FieldText.Append(m_name->Data(), m_name->Count());
					}
					else {
						int pos = m_addr->Find(0, '@');
						if (pos >= 0)
							FieldText.Append(m_addr->Data(), pos);
						else
							FieldText.Append(m_addr->Data(), m_addr->Count());
					}
				}
			}
			else if (iSubItem == 3)
			{

				Charset = m->m_to_charset;
				charsetId = m->m_to_charsetId;

				if ((charsetId == 0) && pMsgView) {
					if (pMsgView->m_cnf_to_charsetId)
						charsetId = pMsgView->m_cnf_to_charsetId;
					else
						charsetId = CP_UTF8;

					//Charset = Charset;
				}

				if (m_bLongMailAddress) {
					FieldText = m->m_to;
				}
				else
				{
					CStringA from;

					int posBeg = 0;
					int posEnd = 0;
					posEnd = m->m_to.Find("@", posBeg);
					if ((posEnd >= 0) && ((posEnd + 1) < m->m_to.GetLength()))
						posEnd = m->m_to.Find(",", posEnd + 1);

					if (posEnd >= 0)
						from = m->m_to.Mid(posBeg, posEnd - posBeg);
					else
						from = m->m_to.Mid(posBeg, m->m_to.GetLength());

					int fromlen = from.GetLength();
					m_name->ClearAndResize(fromlen);
					m_addr->ClearAndResize(fromlen);
					MboxMail::splitMailAddress(from, fromlen, m_name, m_addr);
					FieldText.Empty();
					if (m_name->Count()) {
						FieldText.Append(m_name->Data(), m_name->Count());
					}
					else {
						int pos = m_addr->Find(0, '@');
						if (pos >= 0)
							FieldText.Append(m_addr->Data(), pos);
						else
							FieldText.Append(m_addr->Data(), m_addr->Count());
					}
				}
			}
			else if (iSubItem == 4)
			{
				FieldText = m->m_subj;
				Charset = m->m_subj_charset;
				charsetId = m->m_subj_charsetId;

				if ((charsetId == 0) && pMsgView) {
					if (pMsgView->m_cnf_subj_charsetId)
						charsetId = pMsgView->m_cnf_subj_charsetId;
					else
						charsetId = CP_UTF8;

					//Charset = Charset;
				}
			}
			else if (iSubItem == 5)
			{
				char sizebuff[32];

				int length = m->m_length;
				int kb = length / 1000;
				if (length % 1000)
					kb++;
				sizebuff[0] = 0;
				_itoa(kb, sizebuff, 10);

				FieldText = sizebuff;
				Charset = "UTF-8";
				charsetId = CP_UTF8;
			}

			if (dc.Attach(hDC))
			{
				if (rItem.state & LVIS_SELECTED)
				{
					dc.FillRect(&rect, &CBrush(bkcolor));
					dc.SetBkMode(TRANSPARENT);
					dc.SetTextColor(txcolor);
				}
				else
				{
					dc.FillRect(&rect, &CBrush(color));
					dc.SetBkMode(TRANSPARENT);
					dc.SetTextColor(RGB(0, 0, 0));
				}

				dc.Detach();
			}

			int x_offset = 4;
			xpos = rect.left + x_offset;
			ypos = rect.top + 3;

			int h = rect.Height();
			int w = rect.Width();
			w -= x_offset;

			strW.Empty();
			if (TextUtilsEx::Str2WStr(FieldText, charsetId, strW, error))
			{
				strW.Replace(L"&", L"&&");
				rect.InflateRect(-4, -1);
				int r = ::DrawTextExW(hDC, (LPWSTR)((LPCWSTR)strW), strW.GetLength(), &rect, DT_END_ELLIPSIS | DT_SINGLELINE, NULL);
			}
		}

		*pResult = CDRF_SKIPDEFAULT;  // we are done, tell clist to stop futher processing
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

void NListView::PostNcDestroy() 
{
	m_font.DeleteObject();
	m_boldFont.DeleteObject();

	FileUtils::RemoveDir(CMainFrame::GetMboxviewTempPath());
	MboxMail::Destroy();
	DestroyWindow();
	delete this;
}

bool ALongRightProcessProc(const CUPDUPDATA* pCUPDUPData)
{
	PARSE_ARGS *args = (PARSE_ARGS*)pCUPDUPData->GetAppData();

	CString path = args->path;
	MboxMail::pCUPDUPData = pCUPDUPData;

	HANDLE h = GetCurrentThread();
	BOOL prio = SetThreadPriority(h, THREAD_PRIORITY_ABOVE_NORMAL);
	DWORD myThreadId = GetCurrentThreadId();
	DWORD myThreadPri = GetThreadPriority(h);
	TRACE(L"threadId=%ld threadPriority=%ld\n", myThreadId, myThreadPri);

	MboxMail::Parse(path);
	args->exitted = TRUE;
	return true;
}

bool ALongRightProcessProc_LabelView(const CUPDUPDATA* pCUPDUPData)
{
	PARSE_ARGS *args = (PARSE_ARGS*)pCUPDUPData->GetAppData();

	CString path = args->path;
	MboxMail::pCUPDUPData = pCUPDUPData;

	HANDLE h = GetCurrentThread();
	BOOL prio = SetThreadPriority(h, THREAD_PRIORITY_ABOVE_NORMAL);
	DWORD myThreadId = GetCurrentThreadId();
	DWORD myThreadPri = GetThreadPriority(h);
	TRACE(L"threadId=%ld threadPriority=%ld\n", myThreadId, myThreadPri);

	MboxMail::Parse_LabelView(path);
	args->exitted = TRUE;
	return true;
}

bool ALongRightProcessProcFastSearch(const CUPDUPDATA* pCUPDUPData)
{
	FIND_ARGS *args = (FIND_ARGS*)pCUPDUPData->GetAppData();
	MboxMail::pCUPDUPData = pCUPDUPData;

	HANDLE h = GetCurrentThread();
	BOOL prio = SetThreadPriority(h, THREAD_PRIORITY_ABOVE_NORMAL);
	DWORD myThreadId = GetCurrentThreadId();
	DWORD myThreadPri = GetThreadPriority(h);
	TRACE(L"threadId=%ld threadPriority=%ld\n", myThreadId, myThreadPri);

	ULONGLONG tc_start = GetTickCount64();

	BOOL mainThreadContext = FALSE;
	int maxSearchDuration = -1;
	args->retpos = args->lview->DoFastFind(args->searchstart, mainThreadContext, maxSearchDuration, args->findAll);
	args->exitted = TRUE;

	ULONGLONG tc_end = GetTickCount64();
	DWORD delta = (DWORD)(tc_end - tc_start);
	TRACE(L"(DoFastFind) Total search time %ld milliseconds !!!!!!!!!!!.\n", delta);
	return true;
}

#define CACHE_VERSION	19

BOOL NListView::SaveMails(LPCWSTR cache, BOOL mainThread, CString &errorText)
{
	int ni = MboxMail::s_mails.GetSize();
	SerializerHelper sz(cache);
	//if (!sz.open(TRUE, 64))
	if (!sz.open(TRUE)) 
	{
		return FALSE;
	}

	if (!mainThread)
	{
		if (MboxMail::pCUPDUPData)
		{
			CString txt = L"Creating index file ...";
			ResHelper::TranslateString(txt);
			MboxMail::pCUPDUPData->SetProgress(txt, 0);
		}
	}

	UINT curstep = 1;

	if (!mainThread && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	sz.writeInt(CACHE_VERSION);			// version
	sz.writeInt64(MboxMail::s_fSize = FileUtils::FileSize(MboxMail::s_path));	// file size

	_int64 oldestMailTime = 0;
	_int64 latestMailTime = 0;

	MboxMail* m;
	if (ni > 0)
	{
		m = MboxMail::s_mails[0];
		oldestMailTime = m->m_timeDate;
		latestMailTime = m->m_timeDate;;
	}

	int i;
	for (i = 0; i < ni; i++)
	{
		m = MboxMail::s_mails[i];
		if (m->m_timeDate < oldestMailTime)
			oldestMailTime = m->m_timeDate;
		else if (m->m_timeDate > latestMailTime)
			latestMailTime = m->m_timeDate;
	}
	sz.writeInt64(oldestMailTime);
	sz.writeInt64(latestMailTime);
	//
	CStringA reservedString;
	_int64 reservedInt = 0;
	sz.writeString(reservedString);
	sz.writeString(reservedString);
	sz.writeInt64(reservedInt);
	sz.writeInt64(reservedInt);

	sz.writeInt(ni);					// number of elements

	_int64 lastoff = 0;
	BOOL terminated = FALSE;
	CString itemNumberStr;
	std::vector <MailBodyContent*>::iterator it;

	ULONGLONG workRangeFirstPos = 0;
	ULONGLONG workRangeLastPos = ni - 1;
	ProgressTimer progressTimer(workRangeFirstPos, workRangeLastPos);

	for (i = 0; i < ni; i++)
	{
		m = MboxMail::s_mails[i];

		sz.writeInt64(m->m_startOff);
		sz.writeInt(m->m_length);
		sz.writeInt(m->m_index);
		sz.writeInt(m->m_hasAttachments);
		sz.writeInt(m->m_headLength);
		sz.writeInt(m->m_recv);
		sz.writeString(m->m_from);
		sz.writeString(m->m_from_charset);
		sz.writeInt(m->m_from_charsetId);
		sz.writeString(m->m_to);
		sz.writeString(m->m_to_charset);
		sz.writeInt(m->m_to_charsetId);
		sz.writeString(m->m_subj);
		sz.writeString(m->m_subj_charset);
		sz.writeInt(m->m_subj_charsetId);

		sz.writeString(m->m_cc);
		sz.writeString(m->m_cc_charset);
		sz.writeInt(m->m_cc_charsetId);

		sz.writeString(m->m_bcc);
		sz.writeString(m->m_bcc_charset);
		sz.writeInt(m->m_bcc_charsetId);

		sz.writeInt64(m->m_timeDate);
		sz.writeInt(m->m_groupId);
		sz.writeInt(m->m_groupColor);
		sz.writeString(m->m_messageId);
		sz.writeString(m->m_contentType);
		//sz.writeInt(m->m_crc32);

		int count = (int)m->m_ContentDetailsArray.size();
		sz.writeInt(count);

		MailBodyContent *body;
		for (it = m->m_ContentDetailsArray.begin(); it != m->m_ContentDetailsArray.end(); it++)
		{
			body = *it;

			sz.writeInt(body->m_contentOffset);
			sz.writeInt(body->m_contentLength);
			sz.writeString(body->m_contentTransferEncoding);
			sz.writeString(body->m_contentType);
			sz.writeString(body->m_contentId);
			sz.writeString(body->m_attachmentName);
			sz.writeString(body->m_contentDisposition);
			sz.writeString(body->m_contentLocation);
			sz.writeInt(body->m_pageCode);
			sz.writeInt(body->m_attachmentNamePageCode);
		}

		if (!mainThread && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				terminated = TRUE;
				break;
			}

			UINT_PTR dwProgressbarPos = 0;
			ULONGLONG workRangePos = i;
			BOOL needToUpdateStatusBar = progressTimer.UpdateWorkPos(workRangePos, dwProgressbarPos);
			if (needToUpdateStatusBar)
			{
				int nItem = i + 1;
				CString fmt = L"Creating index file ... %d of %d";
				ResHelper::TranslateString(fmt);
				itemNumberStr.Format(fmt, nItem, ni);
				if (MboxMail::pCUPDUPData)  MboxMail::pCUPDUPData->SetProgress(itemNumberStr, (UINT_PTR)(dwProgressbarPos));

				int debug = 1;
			}
		}

		if (lastoff < m->m_startOff)
			lastoff = m->m_startOff;
	}
	TRACE(L"lastoff=%lld\n", lastoff);

	sz.close();

	if (terminated)
	{
		errorText = L"Partial Index file will be deleted.";
		BOOL ret = FileUtils::DelFile(cache);
		if (ret == FALSE)
		{
			errorText.Format(L"Failed to delete index file %s. Please remove manually.");
			int deb = 1;
		}
	}

	return TRUE;
}

int NListView::LoadSingleMail(MboxMail* m, SerializerHelper& sz)
{
	if (!sz.readInt64(&m->m_startOff))
		return 0;
	if (!sz.readInt(&m->m_length))
		return 0;
	if (!sz.readInt(&m->m_index))
		return 0;
	if (!sz.readInt(&m->m_hasAttachments))
		return 0;
	if (!sz.readInt(&m->m_headLength))
		return 0;
	if (!sz.readInt(&m->m_recv))
		return 0;
	if (!sz.readString(m->m_from))
		return 0;
	if (!sz.readString(m->m_from_charset))
		return 0;
	if (!sz.readUInt(&m->m_from_charsetId))
		return 0;
	if (!sz.readString(m->m_to))
		return 0;
	if (!sz.readString(m->m_to_charset))
		return 0;
	if (!sz.readUInt(&m->m_to_charsetId))
		return 0;
	if (!sz.readString(m->m_subj))
		return 0;
	if (!sz.readString(m->m_subj_charset))
		return 0;
	if (!sz.readUInt(&m->m_subj_charsetId))
		return 0;

	if (!sz.readString(m->m_cc))
		return 0;
	if (!sz.readString(m->m_cc_charset))
		return 0;
	if (!sz.readUInt(&m->m_cc_charsetId))
		return 0;

	if (!sz.readString(m->m_bcc))
		return 0;
	if (!sz.readString(m->m_bcc_charset))
		return 0;
	if (!sz.readUInt(&m->m_bcc_charsetId))
		return 0;

	if (!sz.readInt64(&m->m_timeDate))
		return 0;
	if (!sz.readInt(&m->m_groupId))
		return 0;
	if (!sz.readInt(&m->m_groupColor))
		return 0;
	if (!sz.readString(m->m_messageId))
		return 0;
	if (!sz.readString(m->m_contentType))
		return 0;
	//if (!sz.readUInt(&m->m_crc32))
		//break;

	int count = 0;
	if (!sz.readInt(&count))
		return 0;

	MailBodyContent* body;
	int ii;
	for (ii = 0; ii < count; ii++)
	{
		body = new MailBodyContent;
		if (!sz.readInt(&body->m_contentOffset))
			break;
		if (!sz.readInt(&body->m_contentLength))
			break;
		if (!sz.readString(body->m_contentTransferEncoding))
			break;
		if (!sz.readString(body->m_contentType))
			break;
		if (!sz.readString(body->m_contentId))
			break;
		if (!sz.readString(body->m_attachmentName))
			break;
		if (!sz.readString(body->m_contentDisposition))
			break;
		if (!sz.readString(body->m_contentLocation))
			break;
		if (!sz.readUInt(&body->m_pageCode))
			break;
		if (!sz.readUInt(&body->m_attachmentNamePageCode))
			break;

		m->m_ContentDetailsArray.push_back(body);
	}
	return 1;
}

int NListView::LoadMailsInfo(SerializerHelper &sz, MailArchiveFileInfo& maileFileInfo, CString &errorText)
{
	int ret = -1;
#if 0
	//_int64 fileSize = FileUtils::FileSize(MboxMail::s_path);

	SerializerHelper sz(cache);
	//if (!sz.open(FALSE, 64))
	if (!sz.open(FALSE))
	{
		goto ERR;
	}
#endif

	int version;
	_int64 fSize;
	int ni = 0;

	CStringA reservedString;
	_int64 reservedInt = 0;

	if (!sz.readInt(&version))
	{
		goto ERR;
	}
	maileFileInfo.m_version = version;

	if (version != CACHE_VERSION)
	{
		CString cacheName;
		FileUtils::CPathStripPath(sz.GetFileName(), cacheName);

#if 0
		errorText = L"Index file\n\"" + cacheName;
		CString strVersion;
		strVersion.Format(L"%d", version);
		errorText += L"\".\nhas incompatible version\"" + strVersion + L"\". Expected version \"";
		strVersion.Format(L"%d", CACHE_VERSION);
		errorText += strVersion + L"\".\nWill remove\n\"" + cacheName + L"\"\nand recreate from the mail archive file.";
#endif

		CString strVersion;
		strVersion.Format(L"%d", version);
		CString strCacheVersion;
		strCacheVersion.Format(L"%d", CACHE_VERSION);
		CString fmt = L"Index file\n\"%s\".\nhas incompatible version\"%s\". Expected version \"%s\".\n"
			"Will remove\n\"%s\"\nand recreate from the mail archive file."
			;
		ResHelper::TranslateString(fmt);
		errorText.Format(fmt, cacheName, strVersion, strCacheVersion, cacheName);

		HWND h = GetSafeHwnd();
		TRACE(L"%s\n", errorText);
		//int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);

		ret = -2;
		goto ERR;
	}
	if (!sz.readInt64(&fSize))
	{
		goto ERR;
	}

	maileFileInfo.m_fileSize = fSize;

	// TODO: Verify file length of both mbox and mboxview files
	if (fSize != FileUtils::FileSize(MboxMail::s_path))
	{
		goto ERR;
	}

	_int64 oldestMailTime;
	_int64 latestMailTime;

	if (!sz.readInt64(&oldestMailTime))
	{
		goto ERR;
	}
	maileFileInfo.m_oldestMailTime = oldestMailTime;

	if (!sz.readInt64(&latestMailTime))
	{
		goto ERR;;
	}
	maileFileInfo.m_latestMailTime = latestMailTime;

	// Implement goto ERR
	if (!sz.readString(reservedString))
	{
		goto ERR;
	}
	maileFileInfo.m_reservedString1 = reservedString;
	if (!sz.readString(reservedString))
	{
		goto ERR;
	}
	maileFileInfo.m_reservedString2 = reservedString;
	if (!sz.readInt64(&reservedInt))
	{
		goto ERR;
	}
	maileFileInfo.m_reservedInt1 = reservedInt;
	if (!sz.readInt64(&reservedInt))
	{
		sz.close();
		return -1;
	}
	maileFileInfo.m_reservedInt2 = reservedInt;

	_int64 lastoff = 0;
	int i = 0;
	if (sz.readInt(&ni))
	{
		maileFileInfo.m_mailCount = ni;

		if (ni >= 0)
			ret = ni;
	}

ERR:
	return ret;
}

int NListView::LoadMails(LPCWSTR cache, MailArchiveFileInfo& maileFileInfo, MailArray* mails, CString& errorText)
{
	//_int64 fileSize = FileUtils::FileSize(MboxMail::s_path);

	SerializerHelper sz(cache);
	//if (!sz.open(FALSE, 64))
	if (!sz.open(FALSE))
	{
		return -1;
	}

	int ret = NListView::LoadMailsInfo(sz, maileFileInfo, errorText);
	if (ret < 0)
	{
		sz.close();

		HWND h = GetSafeHwnd();
		int answer = ::MessageBox(h, errorText, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return ret;
	}

	ret = -1;
	_int64 lastoff = 0;
	int i = 0;
	int ni = maileFileInfo.m_mailCount;

	if (ni >= 0)
	{
		if (mails == 0)
		{
			MboxMail::s_mails.SetSize(ni);
			MboxMail::s_mails.SetSizeKeepData(0);
		}
		else
		{
			mails->SetSize(ni);
			mails->SetSizeKeepData(0);
		}

		MboxMail* m;
		for (i = 0; i < ni; i++)
		{
			m = new MboxMail();

			if (LoadSingleMail(m, sz) == 0)
				break;

			if (lastoff < m->m_startOff)
				lastoff = m->m_startOff;

			if (mails == 0)
				MboxMail::s_mails.Add(m);
			else
				mails->Add(m);
#if 0
			if (i % 10000 == 0)
			{
				int sz = MboxMail::AllMailsSizeof(MboxMail::s_mails.GetSize());
				int deb = 1;
			}
#endif
		}
	}
	else
	{
		int deb = 1;
	}

	TRACE(L"lastoff=%lld\n", lastoff);

	sz.close();

	if (mails == 0)
		MboxMail::s_fSize = MboxMail::s_oSize = maileFileInfo.m_fileSize;

	if (i < ni)
	{
		if (mails == 0)
			MboxMail::Destroy(&MboxMail::s_mails);
		else
			MboxMail::Destroy(mails);
	}
	else
		ret = ni;

	return ret;
}

int NListView::Cache2Text(LPCWSTR cache, CString format)
{
	MboxMail mm;
	char buff[2048];
	char datebuff[256];
	DWORD offset = 0;
	DWORD nwritten = 0;
	DWORD count = 0;

	CString cacheTextFile = CString(cache) + ".txt";
	HANDLE hFile = CreateFile(cacheTextFile, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE) {
		DWORD retval = GetLastError();
		CloseHandle(hFile);
	}
	SerializerHelper sz(cache);
	if (!sz.open(FALSE)) {
		CloseHandle(hFile);
		return -1;
	}

	_int64 oldestMailTime;
	_int64 latestMailTime;
	CStringA reservedString;
	_int64 reservedInt = 0;

	int version;
	_int64 fSize;
	int ni = -1;
	MboxMail *m = &mm;

	if (!sz.readInt(&version))
		goto ERR;
	if (version != CACHE_VERSION)
		goto ERR;
	if (!sz.readInt64(&fSize))
		goto ERR;
	if (fSize != FileUtils::FileSize(MboxMail::s_path))
		goto ERR;

	if (!sz.readInt64(&oldestMailTime))
		goto ERR;

	if (!sz.readInt64(&latestMailTime))
		goto ERR;

	if (!sz.readString(reservedString))
		goto ERR;
	if (!sz.readString(reservedString))
		goto ERR;
	if (!sz.readInt64(&reservedInt))
		goto ERR;
	if (!sz.readInt64(&reservedInt))
		goto ERR;

	ni = 0;
	_int64 lastoff = 0;
	if (sz.readInt(&ni))
	{
		for (int i = 0; i < ni; i++)
		{
			if (NListView::LoadSingleMail(m, sz) == 0)
				break;

			if (lastoff < m->m_startOff)
				lastoff = m->m_startOff;

			datebuff[0] = 0;
			if (m->m_timeDate >= 0) {
				MyCTime tt(m->m_timeDate);
				CStringA lDateTime = tt.FormatGmtTmA(format);
				strcpy(datebuff, (LPCSTR)lDateTime);
			}

			// TODO: read attachments

			count = sprintf_s(buff, "indx=%d first=%lld len=%d last=%lld att=%d hlen=%d rcv=%d date=\"%s\" from=\"%s\" to=\"%s\" subj=\"%s\"\n",
				i, m->m_startOff, m->m_length, (m->m_startOff+ m->m_length-1), m->m_hasAttachments, m->m_headLength, m->m_recv,
				datebuff, (LPCSTR)m->m_from, (LPCSTR)m->m_to, (LPCSTR)m->m_subj);
			//count++;
			nwritten = 0;
			if (!WriteFile(hFile, buff, count, &nwritten, NULL)) {
				DWORD retval = GetLastError();
			}
			offset += nwritten;
		}
	}
ERR:
	TRACE(L"lastoff=%lld\n", lastoff);
	sz.close();
	CloseHandle(hFile);

	return ni;
}


// Similar to NListView::FillCtrl() used during merge of all mbox files in root and subfolders
// It will parse given mbox file and create MboxMail::s_mails mail array
//
int NListView::FillCtrl_ParseMbox(CString &mboxPath)
{
	MboxMail::Destroy(&MboxMail::s_mails);

	if (mboxPath.IsEmpty())
	{
		//MboxMail::assert_unexpected();
		return -1;
	}
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (!pFrame)
		return -1;
	NTreeView *pTreeView = pFrame->GetTreeView();
	if (!pTreeView)
		return -1;

	// TODO: MessageBox ??

	if (FileUtils::PathFileExist(mboxPath) == FALSE)
	{
		MboxMail::assert_unexpected();
		return -1;
	}

#ifdef _DEBUG
	ULONGLONG tc = GetTickCount64(); // montano@
#endif

#if 0
		MboxMail::pCUPDUPData = 0;
		MboxMail::Parse_LabelView(mboxPath);
#else
	{
		PARSE_ARGS args;
		args.path = mboxPath;
		args.exitted = FALSE;

		//  ALongRightProcessProc will set MboxMail::s_path = m_path;
		MboxMail::runningWorkerThreadType = 1;
		CUPDialog	Dlg(AfxGetMainWnd()->GetSafeHwnd(), ALongRightProcessProc_LabelView, (LPVOID)(PARSE_ARGS*)&args);
		Dlg.SetDialogTemplate(AfxGetApp()->m_hInstance, MAKEINTRESOURCE(IDD_PROGRESS_DLG), IDC_STATIC, IDC_PROGRESS_BAR, IDCANCEL);

		INT_PTR nResult = Dlg.DoModal();

		if (!nResult)  // should never be true ?
		{
			MboxMail::runningWorkerThreadType = 0;
			MboxMail::assert_unexpected();
			return -1;
		}

		int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
		int retResult = LOWORD(nResult);

		if (retResult != IDOK)
		{  // IDOK==1, IDCANCEL==2
			// We should be here when user selects Cancel button
			//_ASSERTE(cancelledbyUser == TRUE);

			DWORD terminationDelay = Dlg.GetTerminationDelay();
			int loopCnt = (terminationDelay+100)/25;

			ULONGLONG tc_start = GetTickCount64();
			while ((loopCnt-- > 0) && (args.exitted == FALSE))
			{
				Sleep(25);
			}
			ULONGLONG tc_end = GetTickCount64();
			DWORD delta = (DWORD)(tc_end - tc_start);
			TRACE(L"(FillCtrl_ParseMbox)Waited %ld milliseconds for thread to exist.\n", delta);

			CString cache;
			BOOL ret = MboxMail::GetMboxviewFilePath(m_path, cache);

			FileUtils::DelFile(cache);
		}
		MboxMail::runningWorkerThreadType = 0;

		MboxMail::pCUPDUPData = NULL;

		int mailCnt = MboxMail::s_mails.GetCount();

	}
#endif

#ifdef _DEBUG
	tc = (GetTickCount64() - tc);
	TRACE(L"FillCtrl Took %d:%d %d\n", (tc / 1000) / 60, (tc / 1000) % 60, tc);
#endif

	return 1;
}

void NListView::FillCtrl()
{
	ClearDescView();
	m_lastFindPos = -1;

	// Don't reset below vars. It helps user to keep thse while searching amd sorting
	//m_lastStartDate = 0;
	//m_lastEndDate = 0;
	//m_searchString.Empty();

	MboxMail::Destroy();
	m_list.SetRedraw(FALSE);
	m_list.DeleteAllItems();
	if (m_path.IsEmpty())
	{
		//MboxMail::assert_unexpected();
		return;
	}
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if( ! pFrame )
		return;
	NTreeView *pTreeView = pFrame->GetTreeView();
	if( ! pTreeView )
		return;

	NMsgView *pMsgView = pFrame->GetMsgView();
	if (pMsgView)
		pMsgView->DisableMailHeader();


	// TODO: MessageBox ??
	// m_path is set to mboxFilePath not just to mboxFileName
	if (FileUtils::PathFileExist(m_path) == FALSE)
	{
		MboxMail::assert_unexpected();
		return;
	}

	CString cache;
	BOOL ret = MboxMail::GetMboxviewFilePath(m_path, cache);

	int ni = 0;
#ifdef _DEBUG
	ULONGLONG tc = GetTickCount64(); // montano@
#endif

#if 0
	// for testing of changes to mboxview file
	{
		int ni_old = 0;
		int ni_new = 0;
		CString cache_new = cache + ".new";
		MailArray s_mails_new;
		if (FileUtils::fileExists(cache_new)) {
			//MboxMail::s_path = m_path;
			MboxMail::SetMboxFilePath(m_path);
			ni_new = LoadMails(cache_new, &s_mails_new);
			if (ni_new < 0) {
				ni_new = 0;
			}
		}
		CString cache_old = cache + ".old";
		MailArray s_mails_old;
		if (FileUtils::fileExists(cache_old)) {
			//MboxMail::s_path = m_path;
			MboxMail::SetMboxFilePath(m_path);
			ni_old = LoadMails(cache_old, &s_mails_old);
			if (ni_old < 0) {
				ni_old = 0;
			}
		}
		if (ni_new == ni_old)
		{
			MboxMail *m_new;
			MboxMail *m_old;
			int i;
			for (i = 0; i < s_mails_old.GetSize(); i++) {
				m_new = s_mails_new[i];
				m_old = s_mails_old[i];
				if (m_new->m_startOff != m_old->m_startOff)
					int deb = 1;
				if (m_new->m_length != m_old->m_length) {
					int delta = m_new->m_length - m_old->m_length;
					TRACE(L"indx=%d off=%lld newlen=%d oldlen=%d deelta=%d\n", i, m_new->m_startOff, m_new->m_length, m_old->m_length, delta);
					int deb = 1;
				}
			}
		}
	}
#endif

	if( FileUtils::PathFileExist(cache) )
	{
		//MboxMail::s_path = m_path;
		MboxMail::SetMboxFilePath(m_path);
		// it populates s_mails from mail index/mboxview file
		CString errorText;
		MailArray* marray = 0;
		m_maileFileInfo.Clear();
		ni = LoadMails(cache, m_maileFileInfo, marray, errorText);
		if ( ni < 0 )
		{
			if (ni != -2)
			{
#if 0
				CString txt = L"Index file\n\n\"" + cache;
				txt += L"\n\nappears to be corrupted. The index file will be recreated.";
#endif
				CString txt;
				CString fmt = L"Index file\n\n\"%s\"\n\nappears to be corrupted. The index file will be recreated.";

				ResHelper::TranslateString(fmt);
				errorText.Format(fmt, cache);

				HWND h = GetSafeHwnd();
				int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
			}

			ni = 0;
			FileUtils::DelFile(cache);
		} else
			m_list.SetItemCount(ni);

		int deb = 1;
	}

	BOOL parseCanceledByUser = FALSE;
	if( !FileUtils::PathFileExist(cache) )
	{
		PARSE_ARGS args;
		args.path = m_path;
		args.exitted = FALSE;

		//  ALongRightProcessProc will set MboxMail::s_path = m_path;
		MboxMail::runningWorkerThreadType = 1;
		CUPDialog	Dlg(AfxGetMainWnd()->GetSafeHwnd(), ALongRightProcessProc, (LPVOID)(PARSE_ARGS*)&args);
		Dlg.SetDialogTemplate(AfxGetApp()->m_hInstance, MAKEINTRESOURCE(IDD_PROGRESS_DLG), IDC_STATIC, IDC_PROGRESS_BAR, IDCANCEL);

		INT_PTR nResult = Dlg.DoModal();

		if (!nResult)  // should never be true ?
		{
			MboxMail::runningWorkerThreadType = 0;
			MboxMail::assert_unexpected();
			return;
		}

		int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
		int retResult = LOWORD(nResult);

		if (retResult != IDOK)
		{  // IDOK==1, IDCANCEL==2
			// We should be here when user selects Cancel button
			//_ASSERTE(cancelledbyUser == TRUE);

			DWORD terminationDelay = Dlg.GetTerminationDelay();
			int loopCnt = (terminationDelay+100)/25;

			ULONGLONG tc_start = GetTickCount64();
			while ((loopCnt-- > 0) && (args.exitted == FALSE))
			{
				Sleep(25);
			}
			ULONGLONG tc_end = GetTickCount64();
			DWORD delta = (DWORD)(tc_end - tc_start);
			TRACE(L"(FillCtrl)Waited %ld milliseconds for thread to exist.\n", delta);

			parseCanceledByUser = TRUE;

			CString txt = L"Parsing of mail archive file was cancelled by user.\n\n"
				L"Only partial index file was created. You can view all mails in the partial index file as long as you stay within the selected mail file\n\n"
				L"\nWhile viewing the selected mail file, in order to view all mails you need to refresh the index file. In order to refresh index file, right-click on the mail file and select \"Refresh Index File\".\n\n"
				L"Or, Select parent folder or any folder or different mail file under Tree View and select again the same mail file.\n";
			ResHelper::TranslateString(txt);
			int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
			int deb = 1;
		}
		MboxMail::runningWorkerThreadType = 0;

		MboxMail::pCUPDUPData = NULL;

		bool ret;
		// s_mails will be sorted by date, and all related mails are lined by indexing
		MboxMail::clearMessageIdTable();
		MboxMail::clearThreadIdTable();
		if (MboxMail::HasGMThreadId())
			//ret = MboxMail::preprocessConversations();
			ret = MboxMail::preprocessConversationsByThreadId();
		else
			ret = MboxMail::preprocessConversations();

		//
		// s_mails are sorted by data here; 
		// sortConversations assigns related mails to unique group ID, 
		// mails are sorted by group ID into s_mails_ref; s_mails is not touched, i.e re mails sorted by date
		// finally unique index is assigned to each mail according to its position in the master  array s_mails_ref
		// TODO: fix the function name ??
		ret = MboxMail::sortConversations();

#if 0
		// creates to index/mboxview file from s_mails
		if( SaveMails(cache) ) {
			ni = MboxMail::s_mails.GetSize();
			m_list.SetItemCount(ni);
			//Cache2Text(cache, m_format);
		}

#else
		WRITE_INDEX_FILE_ARGS wargs;
		wargs.lview = this;
		wargs.cache = cache;
		wargs.exitted = FALSE;
		wargs.ret = TRUE;

		if (parseCanceledByUser == FALSE)
		{
			FileUtils::DelFile(cache);

			CUPDialog	WDlg(AfxGetMainWnd()->GetSafeHwnd(), ALongRightProcessProcWriteIndexFile, (LPVOID)(PARSE_ARGS*)&wargs);
			Dlg.SetDialogTemplate(AfxGetApp()->m_hInstance, MAKEINTRESOURCE(IDD_PROGRESS_DLG), IDC_STATIC, IDC_PROGRESS_BAR, IDCANCEL);

			nResult = WDlg.DoModal();
			if (!nResult)  // should never be true ?
			{
				MboxMail::assert_unexpected();
				return;
			}

			cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
			retResult = LOWORD(nResult);

			if (retResult != IDOK)
			{  // IDOK==1, IDCANCEL==2
				// We should be here when user selects Cancel button
				//_ASSERTE(cancelledbyUser == TRUE);

				DWORD terminationDelay = Dlg.GetTerminationDelay();
				int loopCnt = (terminationDelay + 100) / 25;

				ULONGLONG tc_start = GetTickCount64();
				while ((loopCnt-- > 0) && (args.exitted == FALSE))
				{
					Sleep(25);
				}
				ULONGLONG tc_end = GetTickCount64();
				DWORD delta = (DWORD)(tc_end - tc_start);
				TRACE(L"(FillCtrl)Waited %ld milliseconds for thread to exist.\n", delta);
			}

			if (!wargs.errorText.IsEmpty())
			{

				int answer = MessageBox(wargs.errorText, L"Warning", MB_APPLMODAL | MB_ICONWARNING | MB_OK);
			}
		}

		if (wargs.ret)
		{
			ni = MboxMail::s_mails.GetSize();
			m_list.SetItemCount(ni);
			//Cache2Text(cache, m_format);
		}

		MboxMail::pCUPDUPData = NULL;
#endif

		MboxMail::clearMessageIdTable();
		MboxMail::clearThreadIdTable();

		int mailCnt = MboxMail::s_mails.GetCount();

		if (parseCanceledByUser == FALSE)
		{
			if (mailCnt > 1000)
			{
				CString txt = L"Inline image files can be pre-created now or created on the fly later. "
					"Image cache improves performance when printing large number of mails to PDF or viewing in Browser.\n"
					"\nCreate cache of inline images (may take several minutes depending on the mail count and content) ?";
				ResHelper::TranslateString(txt);
				int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
				if (answer == IDYES)
				{
					int firstMail = 0;
					int lastMail = MboxMail::s_mails.GetCount() - 1;

					CString targetPrintSubFolderName = L"";
					{
						targetPrintSubFolderName = L"Attachments";
					}
					int retval = CreateInlineImageCache_Thread(firstMail, lastMail, targetPrintSubFolderName);
				}
			}
			else if (mailCnt > 0)
			{
				int firstMail = 0;
				int lastMail = MboxMail::s_mails.GetCount() - 1;

				CString targetPrintSubFolderName = L"";
				if (NListView::m_exportMailsMode)
				{
					targetPrintSubFolderName = L"Attachments";
				}
				int retval = CreateInlineImageCache_Thread(firstMail, lastMail, targetPrintSubFolderName);
			}
		}
	}

	MboxMail::s_mails_ref.SetSizeKeepData(MboxMail::s_mails.GetSize());
	MboxMail::s_mails_ref.Copy(MboxMail::s_mails);
	MboxMail::b_mails_sorted = true;
	MboxMail::b_mails_which_sorted = 1;
	MboxMail::SortByDate();

	m_bEditFindFirst = TRUE;  // must call OnEditFind() first and not OnEditFindAgain()

#ifdef _DEBUG
	tc = (GetTickCount64() - tc);
	TRACE(L"FillCtrl Took %d:%d %d\n", (tc/1000)/60, (tc/1000)%60, tc);
#endif
	m_list.EnsureVisible(ni, FALSE);
	m_list.SetRedraw(TRUE);
	BOOL retval;
	if (m_which == 0)
		int deb = 1;
	// m_which == 0 is legal, no exceptiomn will be thrown
	// retval = pTreeView->m_tree.SelectItem(m_which);
	// TODO: Try to move SelectItem after all done
	retval = pTreeView->SelectTreeItem(m_which);
	CString mboxFileName = pTreeView->m_tree.GetItemText(m_which);

	_int64 fSize = FileUtils::FileSize(MboxMail::s_path);
	CString folderPath;
	CString fileName;
	FileUtils::GetFolderPathAndFileName(m_path, folderPath, fileName);

	// This doesn't have any sense. What I am trying to do here  FIXME
	if (fileName.Compare(mboxFileName) != 0)
		_ASSERTE(FALSE);

	//pTreeView->UpdateFileSizesTable(txt, fSize);
	//pTreeView->SaveData(m_which);

	if (pFrame)
		pFrame->SetupMailListsToInitialState();
}

int NListView::MailFileFillCtrl(CString &errorText)
{
	ClearDescView();
	m_lastFindPos = -1;

	// Don't reset below vars. It helps user to keep thse while searching amd sorting
	//m_lastStartDate = 0;
	//m_lastEndDate = 0;
	//m_searchString.Empty();

	MboxMail::Destroy();
	m_list.SetRedraw(FALSE);
	m_list.DeleteAllItems();
	if (m_path.IsEmpty())
	{
		MboxMail::assert_unexpected();
		return -1;
	}
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (!pFrame)
		return -1;
	NTreeView *pTreeView = pFrame->GetTreeView();
	if (!pTreeView)
		return -1;

	NMsgView *pMsgView = pFrame->GetMsgView();
	if (pMsgView)
		pMsgView->DisableMailHeader();


	// TODO: MessageBox ??
	// m_path is set to mboxFilePath not just to mboxFileName
	if (FileUtils::PathFileExist(m_path) == FALSE)
		return -1;

	CString cache;
	BOOL ret = MboxMail::GetMboxviewFilePath(m_path, cache);

	int ni = 0;
	if (FileUtils::PathFileExist(cache))
	{
		//MboxMail::s_path = m_path;
		MboxMail::SetMboxFilePath(m_path);
		// it populates s_mails from mail index/mboxview file
		CString errorText;
		MailArray* marray = 0;
		m_maileFileInfo.Clear();
		ni = LoadMails(cache, m_maileFileInfo, marray, errorText);
		if (ni < 0)
		{
			if (ni != -2)
			{
#if 0
				CString txt = L"Index file\n\n\"" + cache;
				txt += L"\n\nappears to be corrupted. The index file will be recreated.";
#endif

				CString txt;
				CString fmt = L"Index file\n\n\"%s\"\n\nappears to be corrupted. The index file will be recreated.";
				ResHelper::TranslateString(fmt);
				errorText.Format(fmt, cache);

				HWND h = GetSafeHwnd();
				int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
			}

			ni = 0;
			FileUtils::DelFile(cache);
		}
		else
			m_list.SetItemCount(ni);

		int deb = 1;
	}
	if (!FileUtils::PathFileExist(cache))
	{
		MboxMail::assert_unexpected();
		return -1;
	}

	MboxMail::s_mails_ref.SetSizeKeepData(MboxMail::s_mails.GetSize());
	MboxMail::s_mails_ref.Copy(MboxMail::s_mails);
	MboxMail::b_mails_sorted = true;
	MboxMail::b_mails_which_sorted = 1;
	MboxMail::SortByDate();

	m_bEditFindFirst = TRUE;  // must call OnEditFind() first and not OnEditFindAgain()

	if (m_which == 0)
		int deb = 1;

	//BOOL retval;
	// m_which == 0 is legal, no exceptiomn will be thrown
	//retval = pTreeView->m_tree.SelectItem(m_which);
	CString mboxFileName = pTreeView->m_tree.GetItemText(m_which);

	_int64 fSize = FileUtils::FileSize(MboxMail::s_path);
	CString folderPath;
	CString fileName;
	FileUtils::GetFolderPathAndFileName(m_path, folderPath, fileName);

	// This doesn't have any sense. What I am trying to do here
	if (fileName.Compare(mboxFileName) != 0)
		_ASSERTE(TRUE);  // FIXME ignore condition for now
	//_ASSERTE(FALSE);

	if (pFrame)
		pFrame->SetupMailListsToInitialState();

	return 1;
}

static int CALLBACK
MyCompareProc(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort)
{
	// lParamSort contains a pointer to the list view control.
	CListCtrl* pListCtrl = (CListCtrl*)lParamSort;
	int index1 = static_cast <int> (lParam1);
	int index2 = static_cast <int> (lParam2);
	CString    strItem1 = pListCtrl->GetItemText(index1, 0);
	CString    strItem2 = pListCtrl->GetItemText(index2, 0);

	return _tcscmp(strItem1, strItem2);
}

int NListView::AppendInlineAttachmentNameSeparatorLine(CMimeBody* pBP, int bodyCnt, CStringA& bdy, int textType)
{
	if (bodyCnt <= 0)
		return 1;

	bool isInlineAttachment = MboxCMimeHelper::IsInlineAttachment(pBP);

	if (isInlineAttachment)
	{
		CStringA name;
		CStringA fileName;
		CStringA disposition;

		MboxCMimeHelper::Name(pBP, name);
		MboxCMimeHelper::Filename(pBP, fileName);
		CStringA attachmentName = name;
		if (fileName)
			attachmentName = fileName;

		if (!attachmentName.IsEmpty())
		{
			if (textType == 0)
			{
				bdy.Append("\n\n\n----- ");
				bdy.Append(attachmentName);
				bdy.Append(" ---------------------\n\n");
			}
			else
			{
				CStringA bdycharset = "UTF-8"; // FIXMEFIXME
				//bdy.Append(    "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\">");
				bdy.Append("\r\n<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=");
				bdy.Append(bdycharset);
				bdy.Append("\"><body><span><br><br><br>----- ");
				bdy.Append(attachmentName);
				bdy.Append(" ---------------------<br><br>");
				bdy.Append("</span></body></html>\r\n");
			}
		}
	}
	return 1;
}

int NListView::SelectItem(int iItem, BOOL ignoreViewMessageHeader)
{
	CMainFrame *pFrame = 0;
	char* bdy;
	int bdylen;
	char* data;
	int datalen;

	//  EX_TEST 
#if 0
	*(char*)pFrame = 'a';
#endif

	// Sanity check
	if (iItem < 0 || iItem >= MboxMail::s_mails.GetSize()) {
		MboxMail::assert_unexpected();
		//ClearDescView();
		return -1;
	}

	pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (!pFrame)
		return -1;

	NMsgView *pMsgView = pFrame->GetMsgView();
	if (!pMsgView)
		return -1;

	NListView* pListView = pFrame->GetListView();
	if (!pListView)
		return -1;

	HWND h = GetSafeHwnd();
	MboxMail::ShowHint(HintConfig::LabelsHint, h);

	m_lastSel = iItem;

	if (pFrame->m_bViewMessageHeaders == TRUE)
	{
		if (ignoreViewMessageHeader == FALSE)
		{
			pMsgView->SetMsgHeader(iItem, m_gmtTime, m_format);
			pMsgView->ShowMailHeader(iItem);
			return -1;
		}
		else
			pMsgView->DisableMailHeader();
	}
	else
		pMsgView->DisableMailHeader();

	pMsgView->m_bAttach = FALSE;

	// Already set for ID_INDICATOR_MAIL
	if (pFrame) {
		CString sText = L"Mail Retrieval In Progress ...";
		ResHelper::TranslateString(sText);
		int paneId = 1;
		pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
	}

	// Erase any files previously saved
	FileUtils::RemoveDir(CMainFrame::GetMboxviewTempPath());

	// Get cached mail
	MboxMail *m = MboxMail::s_mails[iItem];

	TRACE(L"SelectItem: m_selLast=%d refIndex=%d\n", m_lastSel, m->m_index);

	// Set header data
	pMsgView->SetMsgHeader(iItem, m_gmtTime, m_format);

	// Save raw message
	if (m_bExportEml)
	{
		// Get raw mail body
		CStringA mailbody;
		BOOL ret = m->GetBody(mailbody);
		SaveAsEmlFile(mailbody);
	}

	CFile fpm;
	CFileException ExError2;
	if (!fpm.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError2))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError2);

		CString txt;
		CString fmt = L"Could not open \"%s\" mail file.\n%s";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, MboxMail::s_path, exErrorStr);

		//TRACE(L"%s\n", txt);
		//errorText = txt;

		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		//fp.Close();
		return -1;
	}

	pMsgView->m_attachments.Reset();

	AttachmentMgr attachmentDB;
	AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();

	SimpleString* outbuflarge = MboxMail::m_outdata;

	outbuflarge->Clear();
	UINT pageCode = 0;
	int textType = 1; // try first Html
	int plainTextMode = 0;  // no extra img tags; html text has img tags already
	int textlen = MboxMail::GetMailBody_mboxview(fpm, iItem, outbuflarge, pageCode, textType, plainTextMode);  // returns pageCode
	if (textlen != outbuflarge->Count())
		int deb = 1;

	if (outbuflarge->Count() > 0)
	{
		// mail has one or more HTML blocks
		char* inData = outbuflarge->Data();
		int inDataLen = outbuflarge->Count();  // SimpleString* outbuflarge = MboxMail::m_outdata;

		CStringA bdycharset = "UTF-8";
		if (pageCode == 0)
		{
			_ASSERTE(pageCode > 0);
			const int deb = 1; // get local ANSI code ??
		}
		std::string pageCode2Str;
		BOOL ret = TextUtilsEx::id2charset(pageCode, pageCode2Str);
		if (ret)
			bdycharset = pageCode2Str.c_str();

		SimpleString* result = 0;
		BOOL retResult = FALSE;
		UINT strCodePage = pageCode;
#if 0
		if ((pageCode == 50220)
			|| (pageCode == 50221)
			|| (pageCode == 50222)
			|| (pageCode == 50225)
			|| (pageCode == 50227)
			|| (pageCode == 50229)
			)
#else
		UINT CP_ASCII = 20127;
		if (strCodePage == CP_ASCII)
			strCodePage = CP_UTF8;

		if (strCodePage != CP_UTF8)
#endif
		{
			SimpleString* wBuff = MboxMail::m_workbuf;
			wBuff->ClearAndResize(4 * inDataLen + 1000);

			result = MboxMail::m_largebuf;
			result->ClearAndResize(4 * inDataLen + 1000);

			DWORD error;
			retResult = TextUtilsEx::Str2UTF8(inData, inDataLen, pageCode, result, wBuff, error);
			if (retResult == TRUE)
			{
				inData = (char*)result->Data();
				inDataLen = result->Count();
				bdycharset = "UTF-8";
			}
			else
				result = 0;
		}

		if (bdycharset.IsEmpty())
			bdycharset = "UTF-8";

		CListCtrl* attachments = 0;
		int mailPosition = iItem;
		BOOL useMailPosition = TRUE;
		BOOL createEmbeddedImageFiles = TRUE;

		SimpleString* outbuf = MboxMail::m_workbuf;
		outbuf->ClearAndResize(outbuflarge->Count() * 2);

		CString srcImgFilePath;
		CString embededSrcImgFilePath;
		CString absoluteSrcImgFilePath = CMainFrame::GetMboxviewTempPath();
		CString relativeSrcImgFilePath = L".\\";

		BOOL verifyAttachmentDataAsImageType = TRUE;
		BOOL insertMaxWidthForImgTag = FALSE;
		CStringA maxWidth = "100%";
		CStringA maxHeight = "";
		BOOL makeFileNameUnique = FALSE;
		if (m_fullImgFilePath)
			embededSrcImgFilePath = absoluteSrcImgFilePath;
		else
			embededSrcImgFilePath = relativeSrcImgFilePath;

		BOOL makeAbsoluteImageFilePath = m_fullImgFilePath;
		NListView::UpdateInlineSrcImgPathEx(&fpm, inData, inDataLen, outbuf, makeFileNameUnique, makeAbsoluteImageFilePath, 
			relativeSrcImgFilePath, absoluteSrcImgFilePath, attachmentDB,
			pListView->m_EmbededImagesStats, mailPosition, createEmbeddedImageFiles, verifyAttachmentDataAsImageType, insertMaxWidthForImgTag,
			maxWidth, maxHeight);

		if (outbuf->Count())
		{
			bdy = outbuf->Data();
			bdylen = outbuf->Count();
		}
		else
		{
			bdy = inData;
			bdylen = inDataLen;
		}

		////////////////////////////////////////

		BOOL bodyTagMissing = FALSE;

		if (TextUtilsEx::findNoCase(bdy, bdylen, "<body", 5) < 0)
			bodyTagMissing = TRUE;

		SimpleString* tmpbuf = MboxMail::get_tmpbuf();

		tmpbuf->ClearAndResize(bdylen + 256);
		//mailPosition = 0;
		BOOL m_ReplaceAllWhiteBackgrounTags = TRUE;
		if (m_bApplyColorStyle)
		{
			int retSet = NListView::SetBackgroundColor(bdy, bdylen, tmpbuf, m_ReplaceAllWhiteBackgrounTags);
		}

		if (tmpbuf->Count())
		{
			bdy = tmpbuf->Data();
			bdylen = tmpbuf->Count();
		}
		else
		{
			; // no change
		}

		// MboxMail::m_outdata; MboxMail::m_workbuf; MboxMail::get_tmpbuf(); might be used here  as bdy, bdylen
		// MboxMail::m_outbuf; use to create html messge

		if (bodyTagMissing)
		{
			// File has htm extension but body tag is missing. 
			// TODO: Below assume html TAG is present.  Do we need to check for html TAG ?
			SimpleString* outbuf = MboxMail::m_outbuf;
			outbuf->ClearAndResize(bdylen + 1024);

			outbuf->Append("<!DOCTYPE html>\r\n");
			CStringA hdr = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body><br>\r\n";
			DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);
			if (m_bApplyColorStyle && (color != COLOR_WHITE))
			{
				CStringA colorStr;
				int retC2S = Color2Str(color, colorStr);
				hdr = "<html><head><style>body{background-color: #";
				hdr.Append(colorStr);
				hdr.Append(";}</style><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body><br>\r\n");
			}
			outbuf->Append((LPCSTR)hdr, hdr.GetLength());

			outbuf->Append(bdy, bdylen);
			outbuf->Append("</body></html>");

			// data and datalen point to ready html message
			data = outbuf->Data();
			datalen = outbuf->Count();
			int deb = 1;
		}
		else // body TAG found
		{
			// Assume HTML tag is present :)
			BOOL charsetMissing = FALSE;

			if (TextUtilsEx::findNoCase(bdy, bdylen, "charset=", 8) < 0)
				charsetMissing = TRUE;

			SimpleString* outbuf = MboxMail::m_outbuf;
			outbuf->ClearAndResize(bdylen + 1000);

			outbuf->Append("<!DOCTYPE html>\r\n");
			if (charsetMissing)
			{
				// Old approach would create extra html block depending on m_bApplyColorStyle and hasInlineAttachments in an inconsistent way.
				// Don't think any harm done if we always create extra block
				// Unfortunately don't have respective test eml files to prove

				CStringA hdr = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head></html><br>\r\n";
				DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);
				if (m_bApplyColorStyle && (color != COLOR_WHITE))
				{
					CStringA colorStr;
					int retC2S = Color2Str(color, colorStr);
					hdr = "<html><head><style>body{background-color: #";
					hdr.Append(colorStr);
					hdr.Append(";}</style><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head></html><br>\r\n");
				}

				outbuf->Append((LPCSTR)hdr, hdr.GetLength());
				int deb = 1;
			}
			else
			{
				DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);
				if (m_bApplyColorStyle && (color != COLOR_WHITE))
				{
					CStringA colorStr;
					int retC2S = Color2Str(color, colorStr);
					CStringA hdr = "<html><head><style>body{background-color: #";
					hdr.Append(colorStr);
					hdr.Append(";}</style><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head></html><br>\r\n");
					outbuf->Append((LPCSTR)hdr, hdr.GetLength());
				}
				else
				{
					CStringA hdr = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head></html><br>\r\n";
					outbuf->Append((LPCSTR)hdr, hdr.GetLength());
				}
			}
			outbuf->Append(bdy, bdylen);

			// data and datalen point to ready html message
			data = outbuf->Data();
			datalen = outbuf->Count();
			int deb = 1;
		}

		MboxMail::rel_tmpbuf();

		// MboxMail::m_outdata; MboxMail::m_workbuf; MboxMail::get_tmpbuf(); no longer used

		//////////////////////////////////////

		pMsgView->m_body_charsetId = pageCode;
		//pMsgView->m_body_charset = bdycharset;
		pMsgView->m_body_charset = pageCode2Str.c_str();

		int deb = 1;
	}
	else
	{
		outbuflarge->Clear();  // SimpleString* outbuflarge = MboxMail::m_outdata;
		pageCode = 0;
		textType = 0; // try Text next
		int plainTextMode = 2;  // insert <img src=attachment name> image tags

		int textlen = MboxMail::GetMailBody_mboxview(fpm, iItem, outbuflarge, pageCode, textType, plainTextMode);  // returns pageCode
		if (textlen != outbuflarge->Count())
			int deb = 1;

		bdy = outbuflarge->Data();
		bdylen = outbuflarge->Count();

		////////////////////////////////////////


		CStringA bdycharset = "UTF-8";
		std::string pageCode2Str;
		BOOL ret = TextUtilsEx::id2charset(pageCode, pageCode2Str);
		if (ret)
			bdycharset = pageCode2Str.c_str();

		CStringA divAvoidPageBreak = "divAvoidBreak {break-inside:avoid;}";
		CStringA preStyle = "pre { overflow-x:break-word; white-space:pre; white-space:hp-pre-wrap; white-space:-moz-pre-wrap; white-space:-o-pre-wrap;  white-space:-pre-wrap; white-space:pre-wrap; word-wrap:break-word;}";

		SimpleString* result = 0;
		BOOL retResult = FALSE;
		UINT strCodePage = pageCode;
#if 0
		if ((strCodePage == 50220)
			|| (strCodePage == 50221)
			|| (strCodePage == 50222)
			|| (strCodePage == 50225)
			|| (strCodePage == 50227)
			|| (strCodePage == 50229)
			)
#else
		UINT CP_ASCII = 20127;
		if (strCodePage == CP_ASCII)
			strCodePage = CP_UTF8;

		if (strCodePage != CP_UTF8)
#endif
		{
			SimpleString* wBuff = MboxMail::m_workbuf;
			wBuff->ClearAndResize(4 * bdylen + 1000);

			result = MboxMail::m_largebuf;
			result->ClearAndResize(4 * bdylen + 1000);

			DWORD error;
			retResult = TextUtilsEx::Str2UTF8(bdy, bdylen, strCodePage, result, wBuff, error);
			if (retResult == TRUE)
			{
				bdy = (char*)result->Data();
				bdylen = result->Count();
				bdycharset = "UTF-8";
			}
			else
				result = 0;
		}

		if (bdycharset.IsEmpty())
			bdycharset = "UTF-8";

		SimpleString* outbuf = MboxMail::m_outbuf;
		outbuf->ClearAndResize(bdylen*4 + 1000);

		outbuf->Append("<!DOCTYPE html>\r\n");
		CStringA hdr = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\">";
		hdr.Append("<style>\r\n");
		hdr.Append("@media print {\r\n");
		//hdr.Append(divAvoidPageBreak);
		//hdr.Append("\r\n");
		hdr.Append(preStyle);
		hdr.Append("\r\n}\r\n");
		hdr.Append(preStyle);

		DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);
		if (m_bApplyColorStyle && (color != COLOR_WHITE))
		{
			CStringA colorStr;
			int retC2S = Color2Str(color, colorStr);
			hdr.Append("\r\nbody {background-color: #");
			hdr.Append(colorStr);
			hdr.Append(";}");
		}
		hdr.Append("\r\n</style></head><body><br>\r\n");

		outbuf->Append((LPCSTR)hdr, hdr.GetLength());

		char* inData = bdy;
		int inDataLen = bdylen;
		SimpleString* tmpbuf = MboxMail::get_tmpbuf();

		TextUtilsEx::EncodeAsHtmlText(inData, inDataLen, tmpbuf);

		inData = tmpbuf->Data();
		inDataLen = tmpbuf->Count();

		outbuf->Append(inData, inDataLen);

		hdr = "</body></html>";
		outbuf->Append((LPCSTR)hdr, hdr.GetLength());

		data = outbuf->Data();
		datalen = outbuf->Count();

		MboxMail::rel_tmpbuf();

		int mailPosition = iItem;
		BOOL useMailPosition = TRUE;
		BOOL createEmbeddedImageFiles = TRUE;

		SimpleString* workbuff = MboxMail::m_workbuf;
		workbuff->ClearAndResize(datalen * 2);

		AttachmentMgr embededAttachmentDB;
		CString absoluteSrcImgFilePath = CMainFrame::GetMboxviewTempPath();
		CString relativeSrcImgFilePath = L".\\";

		CString srcImgFilePath = CMainFrame::GetMboxviewTempPath();
		CStringA embededSrcImgFilePath = ".\\";
		BOOL verifyAttachmentDataAsImageType = FALSE;
		BOOL insertMaxWidthForImgTag = FALSE;
		CStringA maxWidth = "100%";
		CStringA maxHeight = "";
		BOOL makeFileNameUnique = FALSE;
		BOOL makeAbsoluteImageFilePath = m_fullImgFilePath;
		NListView::UpdateInlineSrcImgPathEx(&fpm, data, datalen, workbuff, makeFileNameUnique, makeAbsoluteImageFilePath, 
			relativeSrcImgFilePath, absoluteSrcImgFilePath, attachmentDB,
			pListView->m_EmbededImagesStats, mailPosition, createEmbeddedImageFiles, verifyAttachmentDataAsImageType, insertMaxWidthForImgTag,
			maxWidth, maxHeight);

		if (workbuff->Count())
		{
			data = workbuff->Data();
			datalen = workbuff->Count();
		}
		else
		{
			; // no change
		}

		pMsgView->m_body_charsetId = pageCode;
		//pMsgView->m_body_charset = bdycharset;
		pMsgView->m_body_charset = pageCode2Str.c_str();

		int deb = 1;
	}

	attachmentDB.Sort();

	int attachmentCnt = (int)pMsgView->m_attachments.m_attachmentTbl.GetCount();

	attachmentDB.PrintMap(CString(L"Attachment Map with Embedded Images\n"));
	attachmentDB.PrintArray(CString(L"Attachment Map with Embedded Images\n"));

	CString attachmentFolderPath = CMainFrame::GetMboxviewTempPath();
	BOOL prependMailId = FALSE;

	if (attachmentConfigParams && !attachmentConfigParams->m_bShowAllAttachments_Window)
		attachmentDB.Clear();

	BOOL extraFileNameValidation = FALSE;
	CreateMailAttachments(&fpm, iItem, &attachmentFolderPath, prependMailId, attachmentDB, extraFileNameValidation);

	attachmentDB.PrintMap(CString(L"Attachment Map with Embedded Images and Attachments\n"));
	attachmentDB.PrintArray(CString(L"Attachment Map with Embedded Images and Attachments\n"));

	attachmentDB.Sort();

	if (MboxMail::s_mails.GetCount())
		int deb = 1;

	int clistCnt = pMsgView->m_attachments.GetItemCount();

	attachmentCnt = (int)pMsgView->m_attachments.m_attachmentTbl.GetCount();

	int i;
	for (i = 0; i < attachmentDB.m_attachmentArray.GetCount(); i++)
	{
		AttachmentData* item = &attachmentDB.m_attachmentArray[i];

		CString attachmenName = item->m_nameW;
		CString attachementPath = attachmentFolderPath + "\\" + attachmenName;

		TRACE(L"attachmentDB:All: %d value \"%s\" %d\n",
			i, item->m_nameW, item->m_isEmbedded);

		if (item->m_isEmbedded)
		{
			if (attachmentConfigParams && attachmentConfigParams->m_bShowAllAttachments_Window)
			{
				pMsgView->m_attachments.InsertItemW(attachmenName, i, 0);
			}
		}
		else
		{
			pMsgView->m_attachments.InsertItemW(attachmenName, i, 0);
		}
	}

	pMsgView->m_attachments.Complete();

	// if (pMsgView->m_attachments.m_attachmentTbl.GetCount()) or
	if (pMsgView->m_attachments.GetItemCount())
		pMsgView->m_bAttach = TRUE;

	fpm.Close();

	// Get temporary file name with correct extension for IE to display
	CString extW = L"htm";
	m_curFile = CMainFrame::CreateTempFileName(extW);
	CFileException ExError;
	CFile fp;
	if (fp.Open(m_curFile, CFile::modeWrite | CFile::modeCreate, &ExError))
	{
		fp.Write(data, datalen);

		AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
		if (attachmentConfigParams && ((attachmentConfigParams->m_bSingleMailView || attachmentConfigParams->m_bMultipleMailsView)))
		{
			CString relativeFolderPath = L".\\";
			CString* pRelativeFolderPath = &relativeFolderPath;
			CString* pAbsoluteFolderPath = &attachmentFolderPath;
			if (!m_fullImgFilePath)
				pAbsoluteFolderPath = 0;

			AppendPictureAttachments(0, attachmentDB, pAbsoluteFolderPath, pRelativeFolderPath, &fp);
		}

		fp.Close();
	}
	else
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);  // TODO
		;
	}

	// Display mail in internal IE
	pMsgView->m_browser.Navigate(m_curFile, NULL);
	// Update layou to show/hide attachments
	pMsgView->UpdateLayout();

	MboxMail::ShowHint(HintConfig::MailSelectionHint, GetSafeHwnd());

	if (pMsgView->m_attachments.GetItemCount())
		PostMsgCmdParamAttachmentHint();

	if (MboxMail::nWhichMailList == IDC_FIND_LIST)
		m_bHighlightAllSet = m_bHighlightAll;

	_ASSERTE(!MboxMail::m_tmpbufBusy);
	return 1;
}

BOOL NListView::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt) 
{
	CWnd *wnd = WindowFromPoint(pt);
	if( wnd == NULL )
		return FALSE;
	if( wnd == this ) {
		BOOL ret = CWnd::OnMouseWheel(nFlags, zDelta, pt);
		return TRUE;
	}
	if ((GetKeyState(VK_CONTROL) & 0x80) == 0) {  // if CTRL key not Down; Do we need to post msg further anyway
		// Commented out, it freezes mbox viewer and and even IE for few seconds when CTRL/SHIFT/etc key are kept down
		; // return wnd->PostMessage(WM_MOUSEWHEEL, MAKELPARAM(nFlags, zDelta), MAKELPARAM(pt.x, pt.y));
	}
	return TRUE;
}

void NListView::ClearDescView()
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if( ! pFrame )
		return;
	NMsgView *pMsgView = pFrame->GetMsgView();
	if( ! pMsgView )
		return;
	pMsgView->m_strSubject.LoadString(IDS_DESC_NONE);
	pMsgView->m_strFrom.LoadString(IDS_DESC_NONE);
	pMsgView->m_strDate.LoadString(IDS_DESC_NONE);
	pMsgView->m_strTo.LoadString(IDS_DESC_NONE);
	pMsgView->m_strCC.LoadString(IDS_DESC_NONE);
	pMsgView->m_strBCC.LoadString(IDS_DESC_NONE);


	pMsgView->m_subj_charset.SetString("");
	pMsgView->m_from_charset.SetString("");
	pMsgView->m_date_charset.SetString("");
	pMsgView->m_to_charset.SetString("");
	pMsgView->m_cc_charset.SetString("");
	pMsgView->m_bcc_charset.SetString("");

	pMsgView->m_body_charset.SetString("");

	pMsgView->m_body_charsetId = 0;
	pMsgView->m_attachments.DeleteAllItems();
	FileUtils::RemoveDir(CMainFrame::GetMboxviewTempPath());
	m_curFile.Empty();

	CString url = "about:blank";
	DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);

	{
		CString colorStr;
		int retC2A = NListView::Color2Str(color, colorStr);

		url = L"about:<html><head><style>body{background-color: #";
		url.Append(colorStr);
		url.Append(L";}</style></head><body></body></html><br>");
	}

	pMsgView->m_browser.Navigate(url, NULL);
	pMsgView->UpdateLayout();
	pMsgView->PostMessage(WM_PAINT);
	m_lastSel = -1;
}

void NListView::OnUpdateEditFind(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(! m_bInFind && m_list.GetItemCount()>0 );
	return;
}

#include "FindDlg.h"

time_t NListView::OleToTime_t(COleDateTime *ot) {
	SYSTEMTIME st;
	ot->GetAsSystemTime(st);
	st.wHour = 0;
	st.wMinute = 0;
	st.wSecond = 0;
	MyCTime ct(st);
	return ct.GetTime();
}

void NListView::OnEditFind()
{
	if (m_bInFind)
		return;

	m_advancedFind = FALSE;

	m_bInFind = true;
	CFindDlg dlg;

	// This code is too comples and error prone. Find time to fix this
	// same in OnEdistAdvancedFind. Create function to initialize param structure ?

	int nWhichMailList = MboxMail::nWhichMailList;

	if ((nWhichMailList != m_findParams.m_last_nWhichMailList) && (nWhichMailList != IDC_LABEL_LIST))
	{
		m_findParams.m_filterDates = FALSE;
		m_findParams.m_last_nWhichMailList = nWhichMailList;
	}

	if (m_findParams.m_mboxMailStartDate.GetTime() < 0)
	{
		_ASSERTE(m_findParams.m_bNeedToFindMailMinMaxTime == TRUE);
		if (m_maileFileInfo.m_oldestMailTime == m_maileFileInfo.m_latestMailTime)  // old version possibly
			FindMinMaxTime(m_findParams.m_mboxMailStartDate, m_findParams.m_mboxMailEndDate);
		else
		{
			m_findParams.m_mboxMailStartDate = m_maileFileInfo.m_oldestMailTime;
			m_findParams.m_mboxMailEndDate = m_maileFileInfo.m_latestMailTime;
		}
	}

	if ((MboxMail::nWhichMailList == IDC_ARCHIVE_LIST) || (MboxMail::nWhichMailList == IDC_LABEL_LIST))
	{
		if (m_findParams.m_bNeedToFindMailMinMaxTime)
		{
			m_findParams.m_lastStartDate = m_findParams.m_mboxMailStartDate;
			m_findParams.m_lastEndDate = m_findParams.m_mboxMailEndDate;

			m_findParams.m_archiveList_lastStartDate = m_findParams.m_lastStartDate;
			m_findParams.m_archiveList_lastEndDate = m_findParams.m_lastEndDate;
			m_findParams.m_archiveList_filterDates = m_findParams.m_filterDates;

			m_findParams.m_bNeedToFindMailMinMaxTime = FALSE;
		}
	}

	MailArchiveFileInfo* maileFileInfo = &m_maileFileInfo;

	_ASSERTE(m_findParams.m_mboxMailStartDate.GetTime() == m_maileFileInfo.m_oldestMailTime);
	_ASSERTE(m_findParams.m_mboxMailEndDate.GetTime() == m_maileFileInfo.m_latestMailTime);

	if ((MboxMail::nWhichMailList != IDC_ARCHIVE_LIST) && (MboxMail::nWhichMailList != IDC_LABEL_LIST))
	{
		FindMinMaxTime(m_findParams.m_lastStartDate, m_findParams.m_lastEndDate);
	}
	else
	{
		m_findParams.m_lastStartDate = m_findParams.m_archiveList_lastStartDate;
		m_findParams.m_lastEndDate = m_findParams.m_archiveList_lastEndDate;
		m_findParams.m_filterDates = m_findParams.m_archiveList_filterDates;
	}

	BOOL retDT = MyCTimeToOleTime(m_findParams.m_lastStartDate, m_findParams.m_startDate, 1);  // DOWN
	retDT = MyCTimeToOleTime(m_findParams.m_lastEndDate, m_findParams.m_endDate, 2);  // UP


	m_findParams.m_string = m_searchString;

	dlg.m_params.Copy(m_findParams);

	if (dlg.DoModal() == IDOK) 
	{
		m_findParams.Copy(dlg.m_params);

		if (m_findParams.m_bFindAll && !m_findParams.m_bFindNext)  // not implemented and it makes no difference if we set m_bFindNext to TRUE
			m_findParams.m_bFindNext = TRUE;

		if (m_findParams.m_filterDates == TRUE)
		{
			// mails betwenn m_findParams.m_lastStartDate and m_findParams.m_lastEndDate will be processed
			retDT = NListView::OleTime2MyCTime(m_findParams.m_startDate, m_findParams.m_lastStartDate, FALSE);
			retDT = NListView::OleTime2MyCTime(m_findParams.m_endDate, m_findParams.m_lastEndDate, TRUE);

			if ((MboxMail::nWhichMailList == IDC_ARCHIVE_LIST) || (MboxMail::nWhichMailList == IDC_LABEL_LIST))
			{
				m_findParams.m_archiveList_lastStartDate = m_findParams.m_lastStartDate;
				m_findParams.m_archiveList_lastEndDate = m_findParams.m_lastEndDate;
				m_findParams.m_archiveList_filterDates = TRUE;

			}
		}
		else  // all mails will be processed
		{
			if ((MboxMail::nWhichMailList == IDC_ARCHIVE_LIST) || (MboxMail::nWhichMailList == IDC_LABEL_LIST))
			{
				m_findParams.m_archiveList_filterDates = FALSE;
			}
		}

		m_lastFindPos = -1;

		// To optimize content/message and/or attachments search performance
		if ((m_findParams.m_bContent || m_findParams.m_bAttachments) && (MboxMail::b_mails_sorted == true))
		{
			int whichSorted = abs(MboxMail::b_mails_which_sorted);
			if ((whichSorted != 1) && (whichSorted != 99)) { // related mails should be in close proximity in the mail file
				MboxMail::SortByDate();
				MboxMail::b_mails_sorted = true;
				MboxMail::b_mails_which_sorted = 1;

				RedrawMails();
			}
		}

		m_bHighlightAll = m_findParams.m_bHighlightAll;

		BOOL isTextUnicode = IsWindowUnicode(this->GetSafeHwnd());
		UINT localCP = GetACP();
		std::string str;
		BOOL ret = TextUtilsEx::id2charset(localCP, str);
		//  http://html-codes.info/ansi/html/Windows-1252-superscript%20one_185

		m_bEditFindFirst = FALSE;  // done with OnEditFind() , now we allow to call OnEditFindAgain()

		int sz = MboxMail::s_mails.GetSize();
		if (sz > 0) 
		{
			int which = 0, w = -1;
			if (m_bStartSearchAtSelectedItem == 0)
			{
				if (m_findParams.m_bFindNext == TRUE)
					m_lastFindPos = 0;
				else
					m_lastFindPos = sz - 1;
			}
			else
			{
				if (m_findParams.m_bFindNext == TRUE) {
					if ((m_lastSel >= 0) && (m_lastSel < sz))
						m_lastFindPos = m_lastSel;
					else
						m_lastFindPos = 0;
				}
				else
				{
					if ((m_lastSel >= 0) && (m_lastSel < sz))
						m_lastFindPos = m_lastSel;
					else
						m_lastFindPos = sz - 1;
				}
			}

			which = m_lastFindPos;

			//
			int maxSearchDuration = m_maxSearchDuration;
			if (m_findParams.m_bFindAll)
				maxSearchDuration = 0;
			if (maxSearchDuration > 0) {
				BOOL findAll = FALSE;
				w = DoFastFind(which, TRUE, maxSearchDuration, findAll);
			}
			//_ASSERTE(m_lastFindPos >= 0);
			FIND_ARGS args;
			/*IN*/ args.lview = this; args.searchstart = m_lastFindPos;
			/*OUT*/ args.exitted = FALSE; args.retpos = -1; args.findAll = m_findParams.m_bFindAll;
			if ((w == -2) || (maxSearchDuration == 0)) 
			{
				CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcFastSearch, (LPVOID)(FIND_ARGS*)&args);
				Dlg.SetDialogTemplate(AfxGetApp()->m_hInstance, MAKEINTRESOURCE(IDD_PROGRESS_DLG), IDC_STATIC, IDC_PROGRESS_BAR, IDCANCEL);

				ULONGLONG tc_start = GetTickCount64();

				INT_PTR nResult = Dlg.DoModal();

				if (!nResult) { // should never be true ?
					m_bInFind = false;
					MboxMail::assert_unexpected();
					return;
				}

				ULONGLONG tc_end = GetTickCount64();
				DWORD delta = (DWORD)(tc_end - tc_start);
				CString total;
				total.Format(L"(DoFastFind) Total search time %d milliseconds !!!!!!!!!!!.\n", delta);
				TRACE(L"%s", (LPCWSTR)total);
				//MessageBox(L"info", total, MB_OK | MB_ICONHAND);

				int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
				int retResult = LOWORD(nResult);

				if (retResult != IDOK)
				{  // IDOK==1, IDCANCEL==2
					// We should be here when user selects Cancel button
					//_ASSERTE(cancelledbyUser == TRUE);

					DWORD terminationDelay = Dlg.GetTerminationDelay();
					int loopCnt = (terminationDelay+100)/25;

					ULONGLONG tc_start = GetTickCount64();
					while ((loopCnt-- > 0) && (args.exitted == FALSE))
					{
						Sleep(25);
					}
					ULONGLONG tc_end = GetTickCount64();
					DWORD delta = (DWORD)(tc_end - tc_start);
					TRACE(L"(OnEditFind)Waited %ld milliseconds for thread to exist.\n", delta);
				}

				if (args.exitted == FALSE)
					MboxMail::s_mails_find.SetCountKeepData(0);

				w = args.retpos;
				MboxMail::pCUPDUPData = NULL;


			}
			if (m_findParams.m_bFindAll)
			{
				if (MboxMail::s_mails_find.GetCount() > 0)
					MboxMail::m_findMails.m_lastSel = 0;
				else
					MboxMail::m_findMails.m_lastSel = -1;
				MboxMail::m_findMails.b_mails_which_sorted = MboxMail::b_mails_which_sorted;

				if (MboxMail::IsFindMailsSelected())
				{
					// TODO: how about additional function; what would be the name ?
					SwitchToMailList(IDC_FIND_LIST, TRUE);
				}
				else
				{
					SwitchToMailList(IDC_FIND_LIST, FALSE);
				}
				m_bInFind = false;
				MboxMail::ShowHint(HintConfig::FindDialogHint, GetSafeHwnd());

				LRESULT lres = this->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);

				// FIXMEFIXME
				// To compare search results with non-unicode MBox Viewer
				// int d = MboxMail::DumpMailSummaryToFile(&MboxMail::s_mails, MboxMail::s_mails.GetCount());
				return;
			}
			else if (w >= 0)
			{
				m_bHighlightAllSet = m_findParams.m_bHighlightAll;
				SelectItemFound(w);  // sets m_lastFindPos
			}
			else {
				MessageBeep(MB_OK);
				m_lastFindPos = -1;
			}
		}
		else {
			MessageBeep(MB_OK);
			m_lastFindPos = -1;
		}
	}
	m_bInFind = false;

	MboxMail::ShowHint(HintConfig::FindDialogHint, GetSafeHwnd());

	LRESULT lres = this->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
}

int NListView::CheckMatch(int i, CStringA &searchString)
{
	int pos = -1;
	int w = -1;
	MboxMail *m = MboxMail::s_mails[i];

	pos = -1;
	if (m_findParams.m_bSubject)
	{
		if (searchString.GetLength() == 1 && searchString[0] == '*') {
			w = i;
			return w;
		}
		if (m_findParams.m_bWholeWord)
			pos = g_tu.StrSearchW((unsigned char *)(LPCSTR)m->m_subj, m->m_subj.GetLength(), m->m_subj_charsetId,
				(unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_findParams.m_bCaseSensitive, 0);
		else
			pos = g_tu.StrSearch((unsigned char *)(LPCSTR)m->m_subj, m->m_subj.GetLength(), m->m_subj_charsetId,
				(unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_findParams.m_bCaseSensitive, 0);
		if (pos >= 0) {
			w = i;
			return w;
		}
	}
	pos = -1;
	if (m_findParams.m_bFrom)
	{
		if (searchString.GetLength() == 1 && searchString[0] == '*') {
			w = i;
			return w;
		}
		if (m_findParams.m_bWholeWord)
			pos = g_tu.StrSearchW((unsigned char *)(LPCSTR)m->m_from, m->m_from.GetLength(), m->m_from_charsetId,
				(unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_findParams.m_bCaseSensitive, 0);
		else
			pos = g_tu.StrSearch((unsigned char *)(LPCSTR)m->m_from, m->m_from.GetLength(), m->m_from_charsetId,
				(unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_findParams.m_bCaseSensitive, 0);
		if (pos >= 0) {
			w = i;
			return w;
		}
	}
	if (m_findParams.m_bTo)
	{
		if (searchString.GetLength() == 1 && searchString[0] == '*') {
			w = i;
			return w;
		}
		if (m_findParams.m_bWholeWord)
			pos = g_tu.StrSearchW((unsigned char *)(LPCSTR)m->m_to, m->m_to.GetLength(), m->m_to_charsetId,
				(unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_findParams.m_bCaseSensitive, 0);
		else
			pos = g_tu.StrSearch((unsigned char *)(LPCSTR)m->m_to, m->m_to.GetLength(), m->m_to_charsetId,
				(unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_findParams.m_bCaseSensitive, 0);
		if (pos >= 0) {
			w = i;
			return w;
		}
	}
	if (m_findParams.m_bCC)
	{
		if (!m->m_cc.IsEmpty()) {
			if (searchString.GetLength() == 1 && searchString[0] == '*') {
				w = i;
				return w;
			}
		}
		if (m_findParams.m_bWholeWord)
			pos = g_tu.StrSearchW((unsigned char *)(LPCSTR)m->m_cc, m->m_cc.GetLength(), m->m_cc_charsetId,
				(unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_findParams.m_bCaseSensitive, 0);
		else
			pos = g_tu.StrSearch((unsigned char *)(LPCSTR)m->m_cc, m->m_cc.GetLength(), m->m_cc_charsetId,
				(unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_findParams.m_bCaseSensitive, 0);
		if (pos >= 0) {
			w = i;
			return w;
		}
	}
	if (m_findParams.m_bBCC)
	{
		if (!m->m_bcc.IsEmpty()) {
			if (searchString.GetLength() == 1 && searchString[0] == '*') {
				w = i;
				return w;
			}
		}
		if (m_findParams.m_bWholeWord)
			pos = g_tu.StrSearchW((unsigned char *)(LPCSTR)m->m_bcc, m->m_bcc.GetLength(), m->m_bcc_charsetId,
				(unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_findParams.m_bCaseSensitive, 0);
		else
			pos = g_tu.StrSearch((unsigned char *)(LPCSTR)m->m_bcc, m->m_bcc.GetLength(), m->m_bcc_charsetId,
				(unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_findParams.m_bCaseSensitive, 0);
		if (pos >= 0) {
			w = i;
			return w;
		}
	}

	if (m_findParams.m_bAttachmentName)
	{
		if (FindAttachmentName(m, searchString, m_findParams.m_bWholeWord, m_findParams.m_bCaseSensitive) == TRUE) {
			w = i;
			return w;
		}
	}

	if (m_findParams.m_bContent || m_findParams.m_bAttachments)
	{
		if (searchString.GetLength() == 1 && searchString[0] == '*') {
			w = i;
			return w;
		}
		if (FindInMailContent(i, m_findParams.m_bContent, m_findParams.m_bAttachments, searchString)) {
			w = i;
			return w;
		}
	}
	return w;
}


//typedef int(__cdecl* _CoreCrtNonSecureSearchSortCompareFunction)(void const*, void const*);
int __cdecl CompPriority(void const* p1, void const* p2)
{
	int pr1 = *((int*)p1);
	int pr2 = *((int*)p2);
	if (pr1 < pr2)
		return -1;
	else if (pr1 > pr2)
		return 1;
	else
		return 0;
}

void NListView::DetermineKeywordsForProgresBar(CString *m_string, CString &keyword1, CString &keyword2)
{
	// Assign priority to each field from 0 to FILTER_FIELDS_NUMB, 0 is highest
	// FindFields[] = { "From" (0->0), "To" (1->5), "Subject" (2->1), "CC" (3->6), "BCC" (4->7), "Message" (5->2), "Attachments" (6->3), "Attachment Name" (7->4) };
	int FldIndx2PriorityMap[FILTER_FIELDS_NUMB] = { 0, 5, 1, 6, 7, 2, 3, 4 };
	int FldPriority2Indx2PriorityMap[FILTER_FIELDS_NUMB] = { 0 };

	int CheckedList_FieldsPriority[FILTER_FIELDS_NUMB] = { -1 };

	int checkedFieldCnt = 0;
	int i;
	for (i = 0; i < FILTER_FIELDS_NUMB; i++)
	{
		if (m_advancedParams.m_bEditChecked[i])
			CheckedList_FieldsPriority[checkedFieldCnt++] = FldIndx2PriorityMap[i];
	}

	qsort(CheckedList_FieldsPriority, checkedFieldCnt, sizeof(int), CompPriority);

	for (i = 0; i < FILTER_FIELDS_NUMB; i++)
	{
		FldPriority2Indx2PriorityMap[FldIndx2PriorityMap[i]] = i;
	}

	CString key1;
	CString key2;

	if (checkedFieldCnt >= 2)
	{
		int pr1 = CheckedList_FieldsPriority[0];
		int indx1 = FldPriority2Indx2PriorityMap[pr1];
		key1.Append(m_string[indx1]);

		int pr2 = CheckedList_FieldsPriority[1];
		int indx2 = FldPriority2Indx2PriorityMap[pr2];
		key2.Append(m_string[indx2]);
	}
	else if (checkedFieldCnt == 1)
	{
		int pr1 = CheckedList_FieldsPriority[0];
		int indx1 = FldPriority2Indx2PriorityMap[pr1];
		key1.Append(m_string[indx1]);

		key2.Empty();
	}

	int availableSpace = 56;
	if (!key1.IsEmpty() && !key2.IsEmpty())
	{
		int addSpace =  18 - key2.GetLength();
		if (addSpace < 0)
			addSpace = 0;
		if (key1.GetLength() > (10 + addSpace))
			keyword1 = key1.Left(10 + addSpace);
		else
			keyword1 = key1;

		if (key2.GetLength() > 18)
			keyword2 = key2.Left(18);
		else
			keyword2 = key2;
	}
	else if (!key1.IsEmpty())
	{
		if (key1.GetLength() > availableSpace)
			keyword1 = key1.Left(availableSpace);
		else
			keyword1 = key1;
	}
	else if (!key2.IsEmpty())
	{
		if (key2.GetLength() > availableSpace)
			keyword1 = key2.Left(availableSpace);
		else
			keyword1 = key2;
	}
}

int NListView::DoFastFind(int which, BOOL mainThreadContext, int maxSearchDuration, BOOL findAll)
{
	int w = -1;
	if (m_advancedFind)
		w = NListView::DoFastFindAdvanced(which, mainThreadContext, maxSearchDuration, findAll);
	else
		w = NListView::DoFastFindLegacy(which, mainThreadContext, maxSearchDuration, findAll);
	return w;
}

int NListView::DoFastFindLegacy(int which, BOOL mainThreadContext, int maxSearchDuration, BOOL findAll)
{
	DWORD error;
	int w = -1;
	int sz = MboxMail::s_mails.GetSize();
	time_t sd = m_findParams.m_lastStartDate.GetTime(), ed = m_findParams.m_lastEndDate.GetTime();

	CStringA searchStringA;
	m_searchString = m_findParams.m_string;

	int availableSpace = 56;
	CString keyword = m_searchString.Left(availableSpace);
	CString searchText = L"Searching for \"" + keyword + L"\" ...";

	if (MboxMail::pCUPDUPData && !mainThreadContext)
	{
		MboxMail::pCUPDUPData->SetProgress(searchText, 0);
	}

	if (m_findParams.m_bCaseSensitive == FALSE)
		m_searchString.MakeLower();

	BOOL W2U = TextUtilsEx::WStr2UTF8(&m_searchString, &searchStringA, error);

	BOOL bFindAllMailsThatDontMatch = FALSE;
	BOOL bFilterDates = FALSE;

	bFindAllMailsThatDontMatch = m_findParams.m_bFindAllMailsThatDontMatch;
	bFilterDates = m_findParams.m_filterDates;

	DWORD myThreadId = GetCurrentThreadId();
	ULONGLONG tc_start = GetTickCount64();

	if (m_findParams.m_bFindNext || findAll)
	{
		m_findAllCount = 0;
		if (findAll)
		{
			MboxMail::s_mails_find.SetSizeKeepData(MboxMail::s_mails.GetCount());   // FIXMEFIXME
			// MboxMail::s_mails_find.SetCountKeepData(0);  // must use  s_mails_find.Add() otherwise exception
		}
		ULONGLONG workRangeFirstPos = which;
		ULONGLONG workRangeLastPos = sz - 1;
		ProgressTimer progressTimer(workRangeFirstPos, workRangeLastPos);

		int mailCnt = sz - which;
		for (int i = which; i < sz; i++)
		{
			MboxMail* m = MboxMail::s_mails[i];
			bool process = false;
			if (bFilterDates) {
				if (m->m_timeDate >= sd && m->m_timeDate <= ed)
					process = true;
			}
			else
				process = true;

			if (process)
			{
				if (CheckMatch(i, searchStringA) >= 0)
				{
					if (findAll)
					{
						if (!bFindAllMailsThatDontMatch)
						{
							MboxMail::s_mails_find[m_findAllCount] = MboxMail::s_mails[i];
							m_findAllCount++;
						}
					}
					else
					{
						w = i;
						break;
					}
				}
				else if (bFindAllMailsThatDontMatch)
				{
					// found mail matching negative search criteria
					if (findAll)
					{
						MboxMail::s_mails_find[m_findAllCount] = MboxMail::s_mails[i];
						m_findAllCount++;
					}
					else
					{
						w = i;
						break;
					}
				}

				ULONGLONG tc_curr = GetTickCount64();
				DWORD tc_elapsed_milliseconds = (DWORD)(tc_curr - tc_start);
				DWORD tc_elapsed_seconds = (DWORD)((tc_curr - tc_start) / 1000);
				//TRACE(L"index=%d pos=%lld threadId=0x%lx tc_elapsed_milliseconds=%ld tc_elapsed_seconds=%ld\n",
					//i, pos, myThreadId, tc_elapsed_milliseconds, tc_elapsed_seconds);

				if (mainThreadContext)
				{
					if ((maxSearchDuration > 0) && (tc_elapsed_seconds > maxSearchDuration))
					{
						m_lastFindPos = i + 1;
						w = -2;
						break;
					}
				}
				else
				{
					UINT_PTR dwProgressbarPos = 0;
					ULONGLONG workRangePos = i - which + 1;
					BOOL needToUpdateStatusBar = progressTimer.UpdateWorkPos(workRangePos, dwProgressbarPos);
					if (needToUpdateStatusBar)
					{
						CString progressText;
						CString fmt = L"%s    %llu of %d";
						ResHelper::TranslateString(fmt);
						progressText.Format(fmt, searchText, workRangePos, mailCnt);

						if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(progressText, (UINT_PTR)(dwProgressbarPos));
						int debug = 1;
					}

					if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate())
					{
						w = -1;
						break;
					}
				}
			}
		}
		if (findAll)
		{
			MboxMail::s_mails_find.SetSizeKeepData(m_findAllCount);
			TRACE(L"DoFastFindLegacy: s_mails_find count=%d\n", MboxMail::s_mails_find.GetCount());
		}
	}
	else
	{
		ULONGLONG workRangeFirstPos = 0;
		ULONGLONG workRangeLastPos = which;
		ProgressTimer progressTimer(workRangeFirstPos, workRangeLastPos);

		int mailCnt = which;
		for (int i = which; i >= 0; i--)
		{
			MboxMail* m = MboxMail::s_mails[i];
			bool process = false;
			if (bFilterDates)
			{
				if (m->m_timeDate >= sd && m->m_timeDate <= ed)
					process = true;
			}
			else
				process = true;

			if (process)
			{
				if (CheckMatch(i, searchStringA) >= 0)
				{
					if (!bFindAllMailsThatDontMatch)
					{
						w = i;
						break;
					}
				}
				else if (bFindAllMailsThatDontMatch)
				{
					// found mail matching negative search criteria
					w = i;
					break;
				}

				ULONGLONG tc_curr = GetTickCount64();
				DWORD tc_elapsed_milliseconds = (DWORD)(tc_curr - tc_start);
				DWORD tc_elapsed_seconds = (DWORD)((tc_curr - tc_start) / 1000);
				//TRACE(L"index=%d pos=%lld threadId=0x%lx tc_elapsed_milliseconds=%ld tc_elapsed_seconds=%ld\n",
				//i, pos, myThreadId, tc_elapsed_milliseconds, tc_elapsed_seconds);

				if (mainThreadContext)
				{
					if ((maxSearchDuration > 0) && (tc_elapsed_seconds > maxSearchDuration))
					{
						m_lastFindPos = i - 1;
						w = -2;
						break;
					}
				}
				else
				{
					UINT_PTR dwProgressbarPos = 0;
					ULONGLONG workRangePos = mailCnt - i + 1;
					BOOL needToUpdateStatusBar = progressTimer.UpdateWorkPos(workRangePos, dwProgressbarPos);
					if (needToUpdateStatusBar)
					{
						CString progressText;
						CString fmt = L"%s    %llu of %d";
						ResHelper::TranslateString(fmt);
						progressText.Format(fmt, searchText, workRangePos, mailCnt);

						if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(progressText, (UINT_PTR)(dwProgressbarPos));
						int debug = 1;
					}

					if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate())
					{
						w = -1;
						break;
					}
				}
			}
		}
	}

	CloseMailFile();
	//if (w >= 0)  CloseMailFile();

	return w;
}

//###########################################

int NListView::DoFastFindAdvanced(int which, BOOL mainThreadContext, int maxSearchDuration, BOOL findAll)
{
	DWORD error;
	int w = -1;
	int sz = MboxMail::s_mails.GetSize();
	time_t sd = m_advancedParams.m_lastStartDate.GetTime(), ed = m_advancedParams.m_lastEndDate.GetTime();

	for (int i = 0; i < FILTER_FIELDS_NUMB; i++)
	{
		m_stringWithCase[i] = m_advancedParams.m_stringA[i];
		if (m_advancedParams.m_bEditChecked[i])
		{
			if (m_advancedParams.m_bCaseSensitive[i] == FALSE)
			{
				CString strW = m_advancedParams.m_string[i];
				strW.MakeLower();
				BOOL ret = TextUtilsEx::WStr2UTF8(&strW, &m_stringWithCase[i], error);
			}
		}
	}

	int messageIndx = 5;
	if (m_advancedParams.m_bEditChecked[messageIndx])
	{
		m_searchString = m_advancedParams.m_string[messageIndx];
		if (m_advancedParams.m_bCaseSensitive[messageIndx] == FALSE)
		{
			m_searchString.MakeLower();
		}
	}

	CString searchText;
	if (MboxMail::pCUPDUPData && !mainThreadContext)
	{
		CString keyword1;
		CString keyword2;

		DetermineKeywordsForProgresBar(m_advancedParams.m_string, keyword1, keyword2);

		if (!keyword1.IsEmpty() && !keyword2.IsEmpty())
		{
			searchText.Append(L"Searching for \"" + keyword1 + L"\" && \"" + keyword2 + L"\" ...");
		}
		else if (!keyword1.IsEmpty())
		{
			searchText.Append(L"Searching for \"" + keyword1 + L"\" ...");
		}
		else if (!keyword2.IsEmpty())
		{
			searchText.Append(L"Searching for \"" + keyword2 + L"\" ...");
		}
		else
			searchText.Append(L"Searching for \"\"");

		MboxMail::pCUPDUPData->SetProgress(searchText, 0);
	}

	BOOL bFindAllMailsThatDontMatch = m_advancedParams.m_bFindAllMailsThatDontMatch;
	BOOL bFilterDates = m_advancedParams.m_filterDates;

	DWORD myThreadId = GetCurrentThreadId();
	ULONGLONG tc_start = GetTickCount64();

	m_findAllCount = 0;
	_ASSERTE(findAll == TRUE);
	_ASSERTE(m_advancedParams.m_bFindAll == TRUE);

	if (findAll)
	{
		MboxMail::s_mails_find.SetSizeKeepData(MboxMail::s_mails.GetCount());
		// MboxMail::s_mails_find.SetCountKeepData(0);  // must use  s_mails_find.Add() otherwise exception
	}
	ULONGLONG workRangeFirstPos = which;
	ULONGLONG workRangeLastPos = sz - 1;
	ProgressTimer progressTimer(workRangeFirstPos, workRangeLastPos);

	int mailCnt = sz - which;

	for (int i = which; i < sz; i++)
	{
		MboxMail* m = MboxMail::s_mails[i];
		bool process = false;
		if (bFilterDates)
		{
			if (m->m_timeDate >= sd && m->m_timeDate <= ed)
				process = true;
			else
				int deb = 1;
		}
		else
			process = true;

		if (process)
		{
			if (CheckMatchAdvanced(i, m_advancedParams) >= 0)
			{
				if (findAll)
				{
					if (!bFindAllMailsThatDontMatch)
					{
						MboxMail::s_mails_find[m_findAllCount] = MboxMail::s_mails[i];
						m_findAllCount++;
					}
				}
			}
			else if (bFindAllMailsThatDontMatch)
			{
				if (findAll)
				{
					MboxMail::s_mails_find[m_findAllCount] = MboxMail::s_mails[i];
					m_findAllCount++;
				}
			}

			ULONGLONG tc_curr = GetTickCount64();
			DWORD tc_elapsed_milliseconds = (DWORD)(tc_curr - tc_start);
			DWORD tc_elapsed_seconds = (DWORD)((tc_curr - tc_start) / 1000);
			//TRACE(L"index=%d pos=%lld threadId=0x%lx tc_elapsed_milliseconds=%ld tc_elapsed_seconds=%ld\n",
				//i, pos, myThreadId, tc_elapsed_milliseconds, tc_elapsed_seconds);

			if (mainThreadContext)
			{
				if ((maxSearchDuration > 0) && (tc_elapsed_seconds > maxSearchDuration)) {
					m_lastFindPos = i + 1;
					w = -2;
					break;
				}
			}
			else
			{
				UINT_PTR dwProgressbarPos = 0;
				ULONGLONG workRangePos = i - which + 1;
				BOOL needToUpdateStatusBar = progressTimer.UpdateWorkPos(workRangePos, dwProgressbarPos);
				if (needToUpdateStatusBar)
				{
					CString progressText;
					CString fmt = L"%s    %llu of %d";
					ResHelper::TranslateString(fmt);
					progressText.Format(fmt, searchText, workRangePos, mailCnt);

					if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(progressText, (UINT_PTR)(dwProgressbarPos));
					int debug = 1;
				}

				if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
					w = -1;
					break;
				}
			}
		}
	}

	if (findAll)
	{
		MboxMail::s_mails_find.SetSizeKeepData(m_findAllCount);
		TRACE(L"DoFastFindAdvanced: s_mails_find count=%d\n", MboxMail::s_mails_find.GetCount());
	}

	CloseMailFile();

	return w;
}

void NListView::OnUpdateEditFindAgain(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(! m_bInFind && m_list.GetItemCount()>0);
	return;
}

void NListView::OnEditFindAgain() 
{
	if ( m_bInFind )
		return;

	if (m_findParams.m_bFindAll)
	{
		OnEditFind();
		return;
	}

	if( m_searchString.IsEmpty() )
	{
		OnEditFind();
		return;
	}

	if (m_bEditFindFirst == TRUE)
	{
		OnEditFind();
		return;
	}

	m_advancedFind = FALSE;

	m_bInFind = true;

	int which = 0, w = -1;
	int sz = MboxMail::s_mails.GetSize();

	if (m_lastFindPos != m_lastSel)
		int deb = 1;

	if (m_bStartSearchAtSelectedItem == 0)
	{
		if (m_findParams.m_bFindNext == TRUE)
		{
			if (m_lastFindPos < 0)
				m_lastFindPos = 0;
			else
				m_lastFindPos++;
			if (m_lastFindPos >= sz)
				m_lastFindPos = 0;
		}
		else
		{
			if (m_lastFindPos < 0)
				m_lastFindPos = sz - 1;
			else
				m_lastFindPos--;
			if (m_lastFindPos < 0)
				m_lastFindPos = sz - 1;
		}
	}
	else
	{
		if (m_findParams.m_bFindNext == TRUE)
		{
			if (m_lastSel < 0)
				m_lastFindPos = 0;
			else
			{
				m_lastFindPos = m_lastSel;
				m_lastFindPos++;
			}
			if (m_lastFindPos >= sz)
				m_lastFindPos = 0;
		}
		else
		{
			if (m_lastSel < 0)
				m_lastFindPos = sz - 1;
			else
			{
				m_lastFindPos = m_lastSel;
				m_lastFindPos--;
			}
			if (m_lastFindPos < 0)
				m_lastFindPos = sz - 1;
		}
	}
	which = m_lastFindPos;

	//
	if (m_maxSearchDuration > 0)
	{
		BOOL findAll = FALSE;  // m_findAll should be FALSE
		w = DoFastFind(which, TRUE, m_maxSearchDuration, findAll);
	}


	FIND_ARGS args;
	/*IN*/ args.lview = this; args.searchstart = m_lastFindPos;
	/*OUT*/ args.exitted = FALSE; args.retpos = -1; args.findAll = FALSE; // m_findAll should be FALSE
	if ((w == -2) || (m_maxSearchDuration == 0)) 
	{
		_ASSERTE(m_lastFindPos >= 0);

		CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcFastSearch, (LPVOID)(FIND_ARGS*)&args);
		Dlg.SetDialogTemplate(AfxGetApp()->m_hInstance, MAKEINTRESOURCE(IDD_PROGRESS_DLG), IDC_STATIC, IDC_PROGRESS_BAR, IDCANCEL);

		ULONGLONG tc_start = GetTickCount64();

		INT_PTR nResult = Dlg.DoModal();

		if (!nResult) { // should never be true ?
			m_bInFind = false;
			MboxMail::assert_unexpected();
			return;
		}

		ULONGLONG tc_end = GetTickCount64();
		DWORD delta = (DWORD)(tc_end - tc_start);
		CString total;
		total.Format(L"(DoFastFind) Total search time %d milliseconds !!!!!!!!!!!.\n", delta);
		TRACE(L"%s", (LPCWSTR)total);
		//MessageBox(L"info", total, MB_OK | MB_ICONHAND);

		int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
		int retResult = LOWORD(nResult);

		if (retResult != IDOK)
		{  // IDOK==1, IDCANCEL==2
			// We should be here when user selects Cancel button
			//_ASSERTE(cancelledbyUser == TRUE);

			DWORD terminationDelay = Dlg.GetTerminationDelay();
			int loopCnt = (terminationDelay+100)/25;

			ULONGLONG tc_start = GetTickCount64();
			while ((loopCnt-- > 0) && (args.exitted == FALSE))
			{
				Sleep(25);
			}
			ULONGLONG tc_end = GetTickCount64();
			DWORD delta = (DWORD)(tc_end - tc_start);
			TRACE(L"(OnEditFindAgain)Waited %ld milliseconds for thread to exist.\n", delta);
		}
		//if (args.exitted == FALSE) w = -1;

		w = args.retpos;
		MboxMail::pCUPDUPData = NULL;
	}
	if (w >= 0)
	{
		m_bHighlightAllSet = m_findParams.m_bHighlightAll;
		SelectItemFound(w); // sets m_lastFindPos
	}
	else
	{
		MessageBeep(MB_OK);
		m_lastFindPos = -1;
	}

	m_bInFind = false;
}

void NListView::SelectItemFound(int which)
{
	if (which < 0)
	{
		MessageBeep(MB_OK);
		return;
	}
	int mailCnt = MboxMail::s_mails.GetCount();
	if (which >= mailCnt)
	{
		MessageBeep(MB_OK);
		return;
	}

	POSITION pos = m_list.GetFirstSelectedItemPosition();
	int nItem;
	while (pos)
	{
		nItem = m_list.GetNextSelectedItem(pos);
		m_list.SetItemState(nItem, 0, LVIS_SELECTED | LVIS_FOCUSED);
	}
	m_list.SetItemState(which, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
	m_list.EnsureVisible(which, FALSE);
	m_bApplyColorStyle = TRUE;
	//if (which < MboxMail::s_mails.GetCount())
	{
		SelectItem(which);
		MboxMail *m = MboxMail::s_mails[which];
		m_lastFindPos = which;
	}
}


void NListView::OnEditVieweml()
{	// Save raw message
	if (m_bExportEml == FALSE)
	{
		POSITION pos = m_list.GetFirstSelectedItemPosition();
		int nItem = m_list.GetNextSelectedItem(pos);
		MboxMail *m = 0;
		if (nItem >= 0) 
		{
			m = MboxMail::s_mails[nItem];
			// Get raw mail body
			CStringA bdy;
			BOOL ret = m->GetBody(bdy);
			// Save mail
			CString emlFile = CMainFrame::GetMboxviewTempPath() + L"mime-message.eml";
			CFileException ExError;
			CFile fp;
			if (fp.Open(emlFile, CFile::modeWrite | CFile::modeCreate, &ExError))
			{
				fp.Write(bdy, bdy.GetLength());
				fp.Close();
			}
			else
			{
				CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError); // TODO
				;
			}
		}
	}

	CString path = CMainFrame::GetMboxviewTempPath();
	HINSTANCE result = ShellExecute(NULL, L"open", path, NULL, NULL, SW_SHOWNORMAL);
}


void NListView::OnUpdateEditVieweml(CCmdUI *pCmdUI)
{
	POSITION pos = m_list.GetFirstSelectedItemPosition();
	int nItem = m_list.GetNextSelectedItem(pos);
	if (nItem >= 0)
		MboxMail *m = MboxMail::s_mails[nItem];
	pCmdUI->Enable(m_list.GetFirstSelectedItemPosition() > 0);
}

void NListView::RedrawMails()
{
	// Based on NListView::FillCtrl(). 
	ClearDescView();
	m_list.SetRedraw(FALSE);
	m_list.DeleteAllItems();
	if (m_path.IsEmpty())
		return;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (!pFrame)
		return;
	NTreeView *pTreeView = pFrame->GetTreeView();
	if (!pTreeView)
		return;

	int ni = MboxMail::s_mails.GetSize();
	m_list.SetItemCount(ni);

	m_list.EnsureVisible(ni, FALSE);
	m_list.SetRedraw(TRUE);
#if 0
	// TODO: Commneted out in 1.0.3.0;  doesn't seem to be needed
	//pTreeView->m_tree.SetItemData(m_which, (DWORD)FileUtils::FileSize(m_path));
	//BOOL retval = pTreeView->m_tree.SetItemState(m_which, 0, TVIS_BOLD);
	BOOL retval = pTreeView->m_tree.SetItemState(m_which, LVIS_SELECTED, TVIS_BOLD);
	retval = pTreeView->m_tree.SelectItem(m_which);
	CString txt = pTreeView->m_tree.GetItemText(m_which);
	pTreeView->SaveData(); // TODO: commented out
#endif
}

void NListView::ResetFileMapView()
{
	m_bMappingError = FALSE;
	m_hMailFile = INVALID_HANDLE_VALUE;
	m_MailFileSize = 0;
	//
	m_hMailFileMap = INVALID_HANDLE_VALUE;
	m_mappingSize = (_int64)MAPPING_SIZE;
	m_mappingsInFile = 0;
	//
	m_curMapBegin = 0;
	m_curMapEnd = 0;
	m_pMapViewBegin = 0;
	m_pMapViewEnd = 0;
	//
	m_OffsetRequested = 0;
	m_BytesRequested = 0;
	m_pViewBegin = 0;
	m_pViewEnd = 0;
	//
}

BOOL NListView::SetupFileMapView(_int64 offset, DWORD length, BOOL findNext)
{
	if (m_bMappingError) {
		return FALSE;
	}

	if (m_hMailFile == INVALID_HANDLE_VALUE)
	{
		m_hMailFile = CreateFile(MboxMail::s_path, GENERIC_READ, FILE_SHARE_READ, NULL,
			OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL);
		if (m_hMailFile == INVALID_HANDLE_VALUE)
		{
			DWORD err = GetLastError();
			TRACE(L"(SetFileMapView)CreateFile failed: INVALID_HANDLE_VALUE error= %ld\n", err);
			CloseMailFile();
			m_bMappingError = TRUE;
			return FALSE;
		}
		LARGE_INTEGER li;
		GetFileSizeEx(m_hMailFile, &li);
		m_MailFileSize = li.QuadPart;
		if (m_MailFileSize < 10)
		{
			CloseHandle(m_hMailFile);
			m_hMailFile = INVALID_HANDLE_VALUE;
			TRACE(L"(SetFileMapView)GetFileSizeEx: Mail file size too small= %lld\n", m_MailFileSize);
			m_MailFileSize = 0;
			CloseMailFile();
			m_bMappingError = TRUE;
			return FALSE;
		}
		DWORD MailFileSizeLow = (DWORD)m_MailFileSize; // low four bytes bytes m_MailFileSize
		DWORD MailFileSizeHigh = (DWORD)(m_MailFileSize >> 32); // high four bytes of m_MailFileSize
		m_hMailFileMap = CreateFileMapping(m_hMailFile, NULL, PAGE_READONLY, MailFileSizeHigh, MailFileSizeLow, NULL);
		if (m_hMailFileMap == NULL)
		{
			CloseHandle(m_hMailFile);
			m_hMailFile = INVALID_HANDLE_VALUE;
			DWORD err = GetLastError();
			TRACE(L"(SetFileMapView)CreateFileMapping failed: INVALID_HANDLE_VALUE error= %ld\n", err);
			CloseMailFile();
			m_bMappingError = TRUE;
			return FALSE;
		}
		m_mappingsInFile = (int)(m_MailFileSize / m_mappingSize);  // plus 1 possibly, not used currently

		SYSTEM_INFO systemInfo;
		GetSystemInfo(&systemInfo);
		m_dwPageSize = systemInfo.dwPageSize;
		m_dwAllocationGranularity = systemInfo.dwAllocationGranularity;
		m_MapViewOfFileExCount = 0;
	}
	// Check if curent map can satisfy request for the per mail map 
	if ((m_curMapEnd > 0) && (offset >= m_curMapBegin) && ((offset + length) <= m_curMapEnd)) {
		m_pViewBegin = m_pMapViewBegin + offset - m_curMapBegin;
		m_pViewEnd = m_pMapViewBegin + (offset + length) - m_curMapBegin;
		return TRUE;
	}
	else if (m_pMapViewBegin)
	{
		UnmapViewOfFile(m_pMapViewBegin);
		m_pMapViewBegin = 0;
		m_pMapViewEnd = 0;
		m_curMapBegin = 0;
		m_curMapEnd = 0;
	}

	if (length > (m_mappingSize - m_dwAllocationGranularity))
	{
		int deb = 1; // TODO
		return FALSE;
	}

	// Create new map
	_int64 aligned_offset;
	DWORD bufSize;
	_int64 dataEnd;
	if (findNext == TRUE)
	{
		aligned_offset = (offset / m_dwAllocationGranularity) * m_dwAllocationGranularity;
		bufSize = ((m_MailFileSize - aligned_offset) < m_mappingSize) ? (DWORD)(m_MailFileSize - aligned_offset) : (DWORD)m_mappingSize;
	}
	else
	{
		dataEnd = offset + length;
		aligned_offset = 0;
		if (dataEnd <= m_mappingSize)
		{
			aligned_offset = 0;
		}
		else
		{
			_int64 viewOffset = dataEnd - m_mappingSize;
			aligned_offset = (viewOffset / m_dwAllocationGranularity) * m_dwAllocationGranularity;
		}
		if ((dataEnd - aligned_offset) > m_mappingSize)
			aligned_offset += m_dwAllocationGranularity;
		bufSize = ((dataEnd - aligned_offset) < m_mappingSize) ? (DWORD)(dataEnd - aligned_offset) : (DWORD)m_mappingSize;
	}
	_int64 delta = offset - aligned_offset;

	DWORD MapSizeLow = (DWORD)aligned_offset; // low four bytes bytes of requested offset
	DWORD MapSizeHigh = (DWORD)(aligned_offset >> 32); // high four bytes of requested offset

	m_MapViewOfFileExCount++;
	char * p = (char *)MapViewOfFileEx(m_hMailFileMap, FILE_MAP_READ, MapSizeHigh, MapSizeLow, bufSize, NULL);
	if (p != 0)
	{
		m_pMapViewBegin = p;
		m_pMapViewEnd = p + bufSize;
		m_curMapBegin = aligned_offset;
		m_curMapEnd = aligned_offset + bufSize;
		//
		m_dwViewSize = length;
		m_pViewBegin = p + delta;
		m_pViewEnd = m_pMapViewBegin + (_int64)length;

		if (bufSize < length) {
			; // TODO ??
		}
	}
	else
	{
		DWORD err = GetLastError();
#if 0
		CString errTxt;
		errTxt.Format(L"MapViewOfFileEx failed err=%ld fsize=%lld offset=%lld align_offset=%lld bufSize=%ld mapBegin=%lld mapEnd=%lld\n",
			err, m_MailFileSize, offset, aligned_offset, bufSize, m_pViewBegin, m_pViewEnd);
#endif
		CString errTxt;
		CString fmt = L"MapViewOfFileEx failed err=%ld fsize=%lld offset=%lld align_offset=%lld bufSize=%ld mapBegin=%lld mapEnd=%lld\n";
		ResHelper::TranslateString(fmt);
		errTxt.Format(fmt, err, m_MailFileSize, offset, aligned_offset, bufSize, m_pViewBegin, m_pViewEnd);

		TRACE(L"(SetFileMapView)MapViewOfFileEx failed: error= %ld\n", err);

		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errTxt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
		CloseMailFile();
		m_bMappingError = TRUE;
		return FALSE;
	}
	return TRUE;
}

void NListView::CloseMailFile()
{
	BOOL unmapRet = TRUE;
	BOOL closeMailFileMapHandle = TRUE;
	BOOL closeMailFileHandle = TRUE;
	if (m_pMapViewBegin) 
		unmapRet = UnmapViewOfFile(m_pMapViewBegin);
	if (m_hMailFileMap != INVALID_HANDLE_VALUE)
		closeMailFileMapHandle = CloseHandle(m_hMailFileMap);
	if (m_hMailFile != INVALID_HANDLE_VALUE)
		closeMailFileHandle = CloseHandle(m_hMailFile);

	ResetFileMapView();
}

BOOL NListView::FindInMailContent(int mailPosition, BOOL bContent, BOOL bAttachment, CStringA &searchString)
{
	char  *pData = 0;
	int datalen = 0;

	SimpleString *outbuf = MboxMail::m_outbuf;
	outbuf->ClearAndResize(10000);

#if 0   // FIXMEFIXME
	CStringA searchString(m_searchStringA);
	if (m_bCaseSens == 0)
		searchString.MakeLower();
#endif

	MboxMail *m = MboxMail::s_mails[mailPosition];
	if (SetupFileMapView(m->m_startOff, m->m_length, m_findParams.m_bFindNext) == FALSE)
		return FALSE;

#if 0
	// old code
	MailBodyContent *body;
	BOOL textPlainFound = FALSE;
	BOOL searchHTML = FALSE;
	int i;
	for (i = 0; i < 2; i++)  // search plain text blocks first or html text blocks if no text blocks
	{
		for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
		{
			body = m->m_ContentDetailsArray[j];

			if (!body->m_attachmentName.IsEmpty()) {  // this is attachment
				if (bAttachment == FALSE)
					continue;
			}
			else if (bContent == FALSE)  // not an attachment but we are not interested in the content/message
				continue;

			if (searchHTML == FALSE) // first iteration we are not looking html blocks
			{
				if (body->m_contentType.CompareNoCase("text/plain") != 0)
					continue;
				else if (bContent && body->m_attachmentName.IsEmpty())
					textPlainFound = TRUE;
			}
			else 
			{
				if (body->m_contentType.CompareNoCase("text/html") != 0)
					continue;
				else
					int deb = 1;
			}

			pData = 0;
			datalen = 0;

			body->DecodeBodyData(m_pViewBegin, m->m_length, outbuf);
			if (outbuf->Count())
			{
				pData = outbuf->Data();
				datalen = outbuf->Count();
			}

			if (pData)
			{
				int pos = -1;
				if (m_findParams.m_bWholeWord)
					pos = g_tu.StrSearchW((unsigned char *)pData, datalen, body->m_pageCode,
						(unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_findParams.m_bCaseSensitive, 0);
				else
					pos = g_tu.StrSearch((unsigned char *)pData, datalen, body->m_pageCode,
						(unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_findParams.m_bCaseSensitive, 0);
				if (pos >= 0) {
					return TRUE;
				}
			}
		}
		if (textPlainFound == FALSE)  // first iteration done; didn't find any non attachment plain text blocks 
			searchHTML = TRUE;
	}
	return FALSE;
#else
	// new code
	// START this should be the same as in AdvancedFindInMailContent
	BOOL cnf_plainText = TRUE;
	BOOL cnf_htmlText = FALSE;
	BOOL cnf_htmlTextOnlyIfNoPlainText = TRUE;

	//BOOL bContent, BOOL bAttachment;
	BOOL cnf_bWholeWord_Content = m_findParams.m_bWholeWord;
	BOOL cnf_bCaseSensitive_Content = m_findParams.m_bCaseSensitive;

	BOOL cnf_bWholeWord_Attachment = m_findParams.m_bWholeWord;
	BOOL cnf_bCaseSensitive_Attachment = m_findParams.m_bCaseSensitive;

#if 0
	BOOL cnf_plainText = params.m_plainText;
	BOOL cnf_htmlText = params.m_htmlText;
	BOOL cnf_htmlTextOnlyIfNoPlainText = params.m_htmlTextOnlyIfNoPlainText;

	int fldIndx = 5;
	//BOOL bContent = params.m_bEditChecked[fldIndx];
	BOOL cnf_bWholeWord_Content = params.m_bWholeWord[fldIndx];
	BOOL cnf_bCaseSensitive_Content = params.m_bCaseSensitive[fldIndx];

	int fldIndx = 6;
	//BOOL bAttachment = params.m_bEditChecked[fldIndx];
	BOOL cnf_bWholeWord_Attachment = params.m_bWholeWord[fldIndx];
	BOOL cnf_bCaseSensitive_Attachment = params.m_bCaseSensitive[fldIndx];

#endif

	MailBodyContent* body;
	BOOL isContentTypeText;
	BOOL isTextPlain;
	BOOL isTextHtml;
	CStringA contentSubType;
	CStringA contentTypeMain;

	BOOL isInline;

	BOOL textPlainFound = FALSE;
	BOOL isAttachment = FALSE;
	BOOL isInlineAttachment = FALSE;
	int i;
	// search plain text blocks first or html text blocks if no text blocks
	// Handle inline attachments as plain text blocks or html text blocks
	// in addition as attachments
	// what about other text subtypes: css, csv, calendar, javascript, 
	for (i = 0; i < 2; i++)
	{
		for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
		{
			body = m->m_ContentDetailsArray[j];

			contentSubType.Empty();
			contentTypeMain.Empty();

			int pos = body->m_contentType.ReverseFind('/');
			if (pos > 0)
			{
				contentSubType = body->m_contentType.Mid(pos + 1);
				contentTypeMain = body->m_contentType.Left(pos);
			}

			isContentTypeText = FALSE;
			isTextPlain = FALSE;
			isTextHtml = FALSE;

			if (contentTypeMain.CompareNoCase("text") == 0) {
				isContentTypeText = TRUE;
			}

			if (!isContentTypeText)
				continue;

			if (contentSubType.CompareNoCase("plain") == 0) {
				isTextPlain = TRUE;
			}
			else if (contentSubType.CompareNoCase("html") == 0) {
				isTextHtml = TRUE;
			}

			isAttachment = FALSE;
			isInline = FALSE;
			isInlineAttachment = FALSE;

			if (body->m_contentDisposition.CompareNoCase("inline") == 0)
			{
				isInline = TRUE;
				isAttachment = FALSE;
			}
			else if (body->m_contentDisposition.CompareNoCase("attachment") == 0)
			{
				isInline = FALSE;
				isAttachment = TRUE;
			}
			// if (isInline)  // body->m_attachmentName can be empty to indicate message content

			if (!body->m_attachmentName.IsEmpty())
			{
				if (isInline)
					isInlineAttachment = TRUE;
				else
					isAttachment = TRUE;
			}

			_ASSERTE((isAttachment && isInlineAttachment) == FALSE);

			if (isAttachment && !bAttachment && !isInlineAttachment)
				continue;

			if (i == 0)
			{
				if (!isTextPlain)
					continue;

				if (cnf_plainText == 0)
					continue;
				if (!isAttachment)
					textPlainFound = TRUE;

			}
			else  // i == 1
			{
				if (!isTextHtml)
					continue;

				; // if (cnf_htmlText == 0)
					//break;
			}
			///////////////

			//////////////
			pData = 0;
			datalen = 0;

			body->DecodeBodyData(m_pViewBegin, m->m_length, outbuf);
			if (outbuf->Count())
			{
				pData = outbuf->Data();
				datalen = outbuf->Count();
			}


			if (pData)
			{
				// main diff comparing with FindInMailContent
				// What is this, ignored anyway
				int fldIndx = 5; // Message
				if (isAttachment || isInlineAttachment)
					fldIndx = 6;

				int pos = -1;
				if (i == 1)
				{
					_ASSERTE(isTextHtml);
#ifdef _DEBUG
					ULONGLONG tc_start = GetTickCount64();
#endif
					SimpleString* inbufPtr;
					SimpleStringWrapper  inbuf(pData, datalen);
					inbufPtr = inbuf.getBasePtr();

					UINT pageCode = CP_UTF8;
					if (body->m_pageCode)
						pageCode = body->m_pageCode;
					UINT outPageCode = pageCode;

					SimpleString* retbuf = MboxMail::m_workbuf;
					retbuf->ClearAndResize(2 * inbuf.Count());

					BOOL bestEffortTextExtract = TRUE;
					if (bestEffortTextExtract)
					{
						HtmlUtils::ExtractTextFromHTML_BestEffort(inbufPtr, retbuf, pageCode, outPageCode);
					}
					else
					{
						// Below Relies on IHTMLDocument2 // This is very very slow.
						HtmlUtils::GetTextFromIHTMLDocument(inbufPtr, retbuf, pageCode, outPageCode);
					}

					pData = retbuf->Data();
					datalen = retbuf->Count();

#ifdef _DEBUG
					ULONGLONG tc_end = GetTickCount64();
					DWORD delta = (DWORD)(tc_end - tc_start);


					if (bestEffortTextExtract)
						; // TRACE(L"AdvancedFindInMailContent:ExtractTextFromHTML_BestEffort extracted text in %ld milliseconds.\n", delta);
					else
						; // TRACE(L"AdvancedFindInMailContent:GetTextFromIHTMLDocument extracted text in %ld milliseconds.\n", delta);
#endif
				}
				// END of this should be the same as in AdvancedFindInMailContent

				pos = -1;
				if (cnf_bWholeWord_Content)
					pos = g_tu.StrSearchW((unsigned char*)pData, datalen, body->m_pageCode,
						(unsigned char*)(LPCSTR)searchString, searchString.GetLength(), cnf_bCaseSensitive_Content, 0);
				else
					pos = g_tu.StrSearch((unsigned char*)pData, datalen, body->m_pageCode,
						(unsigned char*)(LPCSTR)searchString, searchString.GetLength(), cnf_bCaseSensitive_Content, 0);
				if (pos >= 0) {
					return TRUE;
				}
			}
		}
		// no match yet
		// i == 1

		if (cnf_htmlText == 1)
		{
			; //  continue;
		}
		else if (cnf_htmlTextOnlyIfNoPlainText == 1)
		{
			if (textPlainFound == TRUE)
				break;
			else
				int deb = 1;
		}
		else
			break;

	}
	return FALSE;
#endif
}

int  NListView::AdvancedFindInMailContent_DevTest(int mailPosition, BOOL bContent, BOOL bAttachment, CFindAdvancedParams& params)
{
	if (MboxMail::developerMode)
	{
		MboxMail* m = MboxMail::s_mails[mailPosition];
		MailBodyContent* body;

		// Match mails based on composition of plain and html text blocks
		if (params.m_bEditChecked[5])  // message
		{
			BOOL plainTextOnly = FALSE;
			BOOL htmlTextOnly = FALSE;
			BOOL bothTextOnly = FALSE;
			int textBlockCnt = 1;

			// FIXME Below is for testing; Review and Comment out ???
			if (m_stringWithCase[5].Compare("1") == 0)
				plainTextOnly = TRUE;
			else if (m_stringWithCase[5].Compare("1.2") == 0)
			{
				plainTextOnly = TRUE;
				textBlockCnt = 2;
			}
			else if (m_stringWithCase[5].Compare("2") == 0)
				htmlTextOnly = TRUE;
			else if (m_stringWithCase[5].Compare("2.2") == 0)
			{
				htmlTextOnly = TRUE;
				textBlockCnt = 2;
			}
			else if (m_stringWithCase[5].Compare("3") == 0)
				bothTextOnly = TRUE;

			if (plainTextOnly || htmlTextOnly || bothTextOnly)
			{
				int attachmentCnt = 0;
				int plainTextCnt = 0;
				int htmlTextCnt = 0;
				for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
				{
					body = m->m_ContentDetailsArray[j];

					if (!body->m_attachmentName.IsEmpty())
					{
						attachmentCnt++;
					}
					else if (body->m_contentType.CompareNoCase("text/html") == 0)
					{
						htmlTextCnt++;
					}
					else if (body->m_contentType.CompareNoCase("text/plain") == 0)
					{
						plainTextCnt++;
					}
				}

				if (bothTextOnly)
				{
					if (plainTextCnt && htmlTextCnt)
						return 1;
				}
				else if (plainTextOnly)
				{
					if (textBlockCnt == 1)
					{
						if ((plainTextCnt == textBlockCnt) && (htmlTextCnt == 0))
							return 1;
					}
					else
					{
						if ((plainTextCnt >= textBlockCnt) && (htmlTextCnt == 0))
							return 1;
					}
				}
				else if (htmlTextOnly)
				{
					if (textBlockCnt == 1)
					{
						if ((htmlTextCnt == textBlockCnt) && (plainTextCnt == 0))
							return 1;
					}
					else
					{
						if ((htmlTextCnt >= textBlockCnt) && (plainTextCnt == 0))
							return 1;
					}
				}
				return 0;
			}
		}
	}
	return -1;
}


// TODO: Below is similar to FindInMailContent above; Merge?, yes when time permits, please
BOOL NListView::AdvancedFindInMailContent(int mailPosition, BOOL bContent, BOOL bAttachment, CFindAdvancedParams &params)
{
	/*
	
	bContent // search message content
	bAttachment // search attachments

 	params.m_plainText;  // search plain text
	paramsm._htmlText;   // always search html text if no match in plain text
	params.m_htmlTextOnlyIfNoPlainText;    // search html txt only if no match in plain text

	inline attachments are consider as message content blocks
	inline attachments are not considered as attachments for searching
	message content blocks are not attachments

	*/
	const char* cText = "text";
	const int cTextLen = istrlen(cText);
	char  *pData = 0;
	int datalen = 0;

	SimpleString *outbuf = MboxMail::m_outbuf;
	outbuf->ClearAndResize(10000);

	MboxMail *m = MboxMail::s_mails[mailPosition];
	if (SetupFileMapView(m->m_startOff, m->m_length, TRUE) == FALSE)
		return FALSE;

	if (MboxMail::developerMode)
	{
		// Match mails based on composition of plain and html text blocks
		int retval = NListView::AdvancedFindInMailContent_DevTest(mailPosition, bContent, bAttachment, params);
		if (retval >= 0)
			return retval;
	}

	BOOL cnf_plainText = params.m_plainText;
	BOOL cnf_htmlText = params.m_htmlText;
	BOOL cnf_htmlTextOnlyIfNoPlainText = params.m_htmlTextOnlyIfNoPlainText;

	MailBodyContent* body;
	BOOL isContentTypeText;
	BOOL isTextPlain;
	BOOL isTextHtml;
	CStringA contentSubType;
	CStringA contentTypeMain;

	BOOL isInline;

	BOOL textPlainFound = FALSE;
	BOOL isAttachment = FALSE;
	BOOL isInlineAttachment = FALSE;
	int i;
	// search plain text blocks first or html text blocks if no text blocks
	// Handle inline attachments as plain text blocks or html text blocks
	// in addition as attachments
	// what about other text subtypes: css, csv, calendar, javascript, 
	for (i = 0; i < 2; i++) 
	{
		for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
		{
			body = m->m_ContentDetailsArray[j];

			contentSubType.Empty();
			contentTypeMain.Empty();

			int pos = body->m_contentType.ReverseFind('/');
			if (pos > 0)
			{
				contentSubType = body->m_contentType.Mid(pos + 1);
				contentTypeMain = body->m_contentType.Left(pos);
			}

			isContentTypeText = FALSE;
			isTextPlain = FALSE;
			isTextHtml = FALSE;

			if (contentTypeMain.CompareNoCase("text") == 0) {
				isContentTypeText = TRUE;
			}

			if (!isContentTypeText)
				continue;

			if (contentSubType.CompareNoCase("plain") == 0) {
				isTextPlain = TRUE;
			}
			else if (contentSubType.CompareNoCase("html") == 0) {
				isTextHtml = TRUE;
			}

			isAttachment = FALSE;
			isInline = FALSE;
			isInlineAttachment = FALSE;

			if (body->m_contentDisposition.CompareNoCase("inline") == 0)
			{
				isInline = TRUE;
				isAttachment = FALSE;
			}
			else if (body->m_contentDisposition.CompareNoCase("attachment") == 0)
			{
				isInline = FALSE;
				isAttachment = TRUE;
			}
			// if (isInline)  // body->m_attachmentName can be empty to indicate message content

			if (!body->m_attachmentName.IsEmpty())
			{
				if (isInline)
					isInlineAttachment = TRUE;
				else
					isAttachment = TRUE;
			}

			_ASSERTE((isAttachment && isInlineAttachment) == FALSE);

			if (isAttachment && !bAttachment && !isInlineAttachment)
				continue;

			if (i == 0)
			{
				if (!isTextPlain)
					continue;

				if (cnf_plainText == 0)
					continue;
				if (!isAttachment)
					textPlainFound = TRUE;

			}
			else  // i == 1
			{
				if (!isTextHtml)
					continue;

				; // if (cnf_htmlText == 0)
					//break;
			}

			pData = 0;
			datalen = 0;

			body->DecodeBodyData(m_pViewBegin, m->m_length, outbuf);
			if (outbuf->Count())
			{
				pData = outbuf->Data();
				datalen = outbuf->Count();
			}

			if (pData)
			{  
				// main diff comparing with FindInMailContent
				// What is this, ignored anyway
				int fldIndx = 5; // Message
				if (isAttachment || isInlineAttachment)
					fldIndx = 6;

				int pos = -1;
				if (i == 1)
				{
					_ASSERTE(isTextHtml);
#ifdef _DEBUG
					ULONGLONG tc_start = GetTickCount64();
#endif
					SimpleString *inbufPtr;
					SimpleStringWrapper  inbuf(pData, datalen);
					inbufPtr = inbuf.getBasePtr();

					UINT pageCode = CP_UTF8;
					if (body->m_pageCode)
						pageCode = body->m_pageCode;
					UINT outPageCode = pageCode;

					SimpleString *retbuf = MboxMail::m_workbuf;
					retbuf->ClearAndResize(2*inbuf.Count());

					BOOL bestEffortTextExtract = TRUE;
					if (bestEffortTextExtract)
					{
						HtmlUtils::ExtractTextFromHTML_BestEffort(inbufPtr, retbuf, pageCode, outPageCode);
					}
					else
					{
						// Below Relies on IHTMLDocument2 // This is very very slow.
						HtmlUtils::GetTextFromIHTMLDocument(inbufPtr, retbuf, pageCode, outPageCode);
					}

					pData = retbuf->Data();
					datalen = retbuf->Count();

#ifdef _DEBUG
					ULONGLONG tc_end = GetTickCount64();
					DWORD delta = (DWORD)(tc_end - tc_start);


					if (bestEffortTextExtract)
						; // TRACE(L"AdvancedFindInMailContent:ExtractTextFromHTML_BestEffort extracted text in %ld milliseconds.\n", delta);
					else
						; // TRACE(L"AdvancedFindInMailContent:GetTextFromIHTMLDocument extracted text in %ld milliseconds.\n", delta);
#endif
				}

				fldIndx = 5;
				if (params.m_bEditChecked[fldIndx] && !isAttachment)
				{
					if (params.m_bWholeWord[fldIndx])
					{
						pos = g_tu.StrSearchW((unsigned char *)pData, datalen, body->m_pageCode,
							(unsigned char *)(LPCSTR)m_stringWithCase[fldIndx], m_stringWithCase[fldIndx].GetLength(), params.m_bCaseSensitive[fldIndx], 0);
					}
					else
					{
						pos = g_tu.StrSearch((unsigned char *)pData, datalen, body->m_pageCode,
							(unsigned char *)(LPCSTR)m_stringWithCase[fldIndx], m_stringWithCase[fldIndx].GetLength(), params.m_bCaseSensitive[fldIndx], 0);
					}
					if (pos >= 0) {
						return TRUE;
					}
				}

				fldIndx = 6;
				if (params.m_bEditChecked[fldIndx] && isAttachment)
				{
					if (params.m_bWholeWord[fldIndx])
					{
						pos = g_tu.StrSearchW((unsigned char*)pData, datalen, body->m_pageCode,
							(unsigned char*)(LPCSTR)m_stringWithCase[fldIndx], m_stringWithCase[fldIndx].GetLength(), params.m_bCaseSensitive[fldIndx], 0);
					}
					else
					{
						pos = g_tu.StrSearch((unsigned char*)pData, datalen, body->m_pageCode,
							(unsigned char*)(LPCSTR)m_stringWithCase[fldIndx], m_stringWithCase[fldIndx].GetLength(), params.m_bCaseSensitive[fldIndx], 0);
					}
					if (pos >= 0) {
						return TRUE;
					}
				}

			}
		}
		// no match yet
		// i == 1

		if (cnf_htmlText == 1)
		{
			; //  continue;
		}
		else if (cnf_htmlTextOnlyIfNoPlainText == 1)
		{
			if (textPlainFound == TRUE)
				break;
			else
			{
				MboxMail* mm = m;
				int deb = 1;
			}
		}
		else
		{
			MboxMail* mm = m;
			break;
		}

	}
	return FALSE;
}

void NListView::PrintMailGroupToText(BOOL multipleSelectedMails, int iItem, int textType, BOOL forceOpen, BOOL printToPrinter, BOOL createFileOnly)
{
	int firstMail = 0;
	int lastMail = 0;
	MailIndexList *selectedMailsIndexList = 0;

	int mailCnt = 0;
	if (multipleSelectedMails == FALSE)
	{
		if (!((abs(MboxMail::b_mails_which_sorted) == 99) || ((abs(MboxMail::b_mails_which_sorted) == 4) && (m_subjectSortType == 1))))
		{
			//if (abs(MboxMail::b_mails_which_sorted) != 99)
			{

				CString txt = L"Please sort all mails by conversation threads or subject threads first.\n";
				ResHelper::TranslateString(txt);

				//txt += "Select \"View\"->\"Sort By\" ->\"Conversation\" or left click on the first column.";
				HWND h = GetSafeHwnd(); // we don't have any window yet
				int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
				//return;
			}
			return;
		}
		FindFirstAndLastMailOfConversation(iItem, firstMail, lastMail);
		mailCnt = lastMail - firstMail + 1;
	}
	else //  (multipleSelectedMails == TRUE)
	{
		selectedMailsIndexList = PopulateSelectedMailsList();
		mailCnt = (int)selectedMailsIndexList->GetCount();
	}

	if (printToPrinter && (mailCnt > 100))
	{
		CString txt = L"Printing large number of mails using \"Print To -> Printer\""
			L" is not the prefered option, use \"Print to ->PDF\" option instead.\n\n"
			L"\"Print to -> Printer\" runs in the main thread and it can't be cancelled"
			L" easly if desired except by killing MBox Viewer process\n\n"
			L"Depending on the content of mails it may take long time for \"Print to -> Printer\" to complete."
			L" No feedback is provided to users until user is asked to select folder and PDF document name to save\n\n"
			L"Run \"Print to -> Printer\" option if other print options didn't create well formatted PDF.\n\n"
			L"Do you want to continue?";

		ResHelper::TranslateString(txt);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
		{
			return;
		}
		else if (answer == IDYES)
			int deb = 1;
	}

	TEXTFILE_CONFIG textConfig;
	textConfig.m_dateFormat = m_format;
	textConfig.m_bGMTTime = m_gmtTime;
	textConfig.m_nCodePageId = CP_UTF8;

	CString textFileName;
	int ret = 0;
	BOOL progressBar = FALSE;  // No worker thread

	ret = MboxMail::exportToTextFile(textConfig, textFileName, firstMail, lastMail, selectedMailsIndexList, textType, progressBar);
	if (ret > 0) 
	{
		CString path = MboxMail::GetLastPath();
		CString datapath = MboxMail::GetLastDataPath();
		if (!datapath.IsEmpty())  // not likely since the path was valid in MboxMail::exportToTextFile(....);
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

						UINT inCodePage = CP_UTF8;
						NMsgView::PrintHTMLDocumentToPrinter(inbuf, workbuf, inCodePage);

						int deb = 1;
					}
					else
					{
						// MessageBox ??
						CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);  //TODO
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
							CMainFrame::CheckShellExecuteResult(result, h);
						}
						int deb = 1;
					}
					else if (nResponse == IDYES)
					{
						HWND h = GetSafeHwnd();
						HINSTANCE result = ShellExecute(h, L"open", textFileName, NULL, NULL, SW_SHOWNORMAL);
						CMainFrame::CheckShellExecuteResult(result, h, &textFileName);
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
					CMainFrame::CheckShellExecuteResult(result, h, &textFileName);
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

int NListView::RemoveSelectedMails()
{
	int i = 0;
	int nIndx;
	MboxMail *m;

	int removeFirst;
	int removeLast;
	int keepFirst;
	int keepLast;
	int position;
	int nextPositionToMoveMails = -1;
	int removedTotal = 0;
	int removeFirstPosition = -1;

	int selectedCnt = m_list.GetSelectedCount();

	m_list.SetRedraw(FALSE);

	// Assume indexes are sorted from lowest to highest
	// TODO: assert assumption , sort if needed
	POSITION p = m_list.GetFirstSelectedItemPosition();

	if (MboxMail::IsUserMailsSelected())
	{
		int mailsCnt = MboxMail::s_mails.GetCount();
		while (p)
		{
			nIndx = m_list.GetNextSelectedItem(p);
			if (nIndx >= mailsCnt)
				return -1; // TODO: error _ASSERTE ?
			m = MboxMail::s_mails[nIndx];
			m->m_isOnUserSelectedMailList = false;
		}
	}

	p = m_list.GetFirstSelectedItemPosition();

	int sv_mailsCnt = MboxMail::s_mails.GetCount();
	while (p)
	{
		// find block of mails to remove
		if (nextPositionToMoveMails < 0) 
		{
			removeFirst = m_list.GetNextSelectedItem(p);
			removeLast = removeFirst;
			nextPositionToMoveMails = removeFirst;
			removeFirstPosition = removeFirst;
		}

		BOOL moreToRemove = FALSE;
		while (p)
		{
			position = m_list.GetNextSelectedItem(p);
			if (position == (removeLast + 1))
				removeLast = position;
			else
			{
				moreToRemove = TRUE;
				break;
			}
		}

		removedTotal += removeLast - removeFirst + 1;

		// if there is more mails to remove
		if (moreToRemove)
		{
			// move block of mails to keep
			keepFirst = removeLast + 1;
			keepLast = position;

			// move mails to keep to removeFirst position
			int keepCnt = keepLast - keepFirst;
			for (i = keepFirst; i < keepLast; i++)
			{
				MboxMail::s_mails[nextPositionToMoveMails++] = MboxMail::s_mails[i];
			}
			// position points already to the next mail block
			removeFirst = position;
			removeLast = removeFirst;
			if (p == 0)
			{
				// leveragee CArray to shift remaining mails
				int removeCnt = removeLast - nextPositionToMoveMails + 1;
				if (removeCnt)
					MboxMail::s_mails.RemoveAt(nextPositionToMoveMails, removeCnt);

				removedTotal += removeLast - removeFirst + 1;
			}
		}
		else
		{
			// leveragee CArray to shift remaining mails
			int removeCnt = removeLast - nextPositionToMoveMails + 1;
			if (removeCnt)
				MboxMail::s_mails.RemoveAt(nextPositionToMoveMails, removeCnt);
		}
	}

	if (selectedCnt != removedTotal)
		int deb = 1;

	if (sv_mailsCnt == MboxMail::s_mails.GetCount())
		int deb = 1;
	else
		int deb = 1;

	if (abs(MboxMail::b_mails_which_sorted) == 99) {
		MboxMail::SortByGroupId(0, MboxMail::b_mails_which_sorted < 0);  // TODO: this is probably not needed; keep it simple for now
		MboxMail::assignColor2ConvesationGroups(&MboxMail::s_mails);
	}

	if ((abs(MboxMail::b_mails_which_sorted) == 4) && (m_subjectSortType == 1))
	{
		SortBySubjectBasedConversasions(TRUE);
	}

	RedrawMails();

	return removeFirstPosition;
}

int NListView::RemoveAllMails()
{
	m_list.SetRedraw(FALSE);

	if (MboxMail::IsUserMailsSelected())
	{

		MboxMail *m;
		int i;
		for (i = 0; i < MboxMail::s_mails.GetCount(); i++)
		{
			m = MboxMail::s_mails[i];
			m->m_isOnUserSelectedMailList = false;
		}
	}

	MboxMail::s_mails.SetSize(0);
	if (MboxMail::IsFindMailsSelected())
		MboxMail::s_mails_find.SetSize(0);
	else if (MboxMail::IsUserMailsSelected())
		MboxMail::s_mails_edit.SetSize(0);
	else
		int deb = 1; // _ASSERTE ?
	

	// RedrawMails sets m_list.SetItemCount();
	RedrawMails();

	m_lastFindPos = -1;
	m_lastSel = -1;
	m_bEditFindFirst = TRUE;

	return 1;
}

int NListView::CopySelectedMails()
{
	int i = 0;
	int nIndx;
	MailArray &selected_mails = MboxMail::s_mails_selected;
	MboxMail *m;

	int selectedCnt = m_list.GetSelectedCount();
	selected_mails.SetSizeKeepData(selectedCnt);

	POSITION p = m_list.GetFirstSelectedItemPosition();
	int mailsCnt = MboxMail::s_mails.GetCount();
	while (p)
	{
		nIndx = m_list.GetNextSelectedItem(p);
		if (nIndx >= mailsCnt)
			return -1; // TODO: error
		m = MboxMail::s_mails[nIndx];
		selected_mails.SetAtGrow(i++, m);

		m->m_isOnUserSelectedMailList = true;
	}

	// _ASSERTE i <= selectedCnt
	selected_mails.SetSizeKeepData(i);

	MailArray &merged_mails = MboxMail::s_mails_merged;
	MboxMail::MergeTwoMailLists(&selected_mails, &MboxMail::s_mails_edit, &merged_mails);
	//MboxMail::VerifyMergeOfTwoMailLists(&selected_mails, &MboxMail::s_mails_edit, &merged_mails);
	MboxMail::s_mails_edit.CopyKeepData(merged_mails);

	// TODO: consider marking MboxMail::s_mails_edit as dirty, i.e. to allow sorting when selected
	MboxMail::SortByDate(&MboxMail::s_mails_edit);
	MboxMail::m_editMails.m_lastSel = -1;
	if (MboxMail::s_mails_edit.GetCount() > 0)
		MboxMail::m_editMails.m_lastSel = 0;
	MboxMail::m_editMails.b_mails_which_sorted = 1;

	MboxMail::m_editMails.m_bIsDirty = TRUE;

	return 1;
}

int NListView::CopyAllMails()
{
	CString txt = L"Do you want to copy all mails to Users Selected Mails?";
	ResHelper::TranslateString(txt);
	int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
	if (answer == IDYES) 
	{
		MboxMail *m;
		int i;
		for (i = 0; i < MboxMail::s_mails.GetCount(); i++)
		{
			m = MboxMail::s_mails[i];
			m->m_isOnUserSelectedMailList = true;
		}

		// TODO: evaluate other merge solutions
		MailArray &merged_mails = MboxMail::s_mails_merged;
		MboxMail::s_mails_selected.CopyKeepData(MboxMail::s_mails);
		MboxMail::MergeTwoMailLists(&MboxMail::s_mails_selected, &MboxMail::s_mails_edit, &merged_mails);
		//MboxMail::VerifyMergeOfTwoMailLists(&MboxMail::s_mails_selected, &MboxMail::s_mails_edit, &merged_mails);
		MboxMail::s_mails_edit.CopyKeepData(merged_mails);

		// TODO: consider marking MboxMail::s_mails_edit as dirty, i.e. to allow sorting when selected
		MboxMail::SortByDate(&MboxMail::s_mails_edit);
		MboxMail::m_editMails.m_lastSel = -1;
		if (MboxMail::s_mails_edit.GetCount() > 0)
			MboxMail::m_editMails.m_lastSel = 0;
		MboxMail::m_editMails.b_mails_which_sorted = 1;

		MboxMail::m_editMails.m_bIsDirty = TRUE;
	}

	return 1;
}

int NListView::FindInHTML(int iItem)
{
	BOOL itemSelected = FALSE;
	if (m_lastSel != iItem)
	{
		TRACE(L"Selecting %d\n", iItem);
		m_bApplyColorStyle = TRUE;
		SelectItem(iItem);
		itemSelected = TRUE;
	}

	CFindInMailDlg dlg;
	dlg.m_bWholeWord = m_bWholeWordInMail;
	dlg.m_bCaseSensitive = m_bCaseSensInMail;
	dlg.m_string = m_searchStringInMail;
	if (dlg.DoModal() == IDOK) 
	{
		m_searchStringInMail = dlg.m_string;
		m_bWholeWordInMail = dlg.m_bWholeWord;
		m_bCaseSensInMail = dlg.m_bCaseSensitive;

		CMainFrame *pFrame = (CMainFrame*)AfxGetMainWnd();
		if (!pFrame)
			return itemSelected;
		if (!::IsWindow(pFrame->m_hWnd) || !pFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
			return itemSelected;;
		NMsgView *pMsgView = pFrame->GetMsgView();
		if (!pMsgView || !::IsWindow(pMsgView->m_hWnd))
			return itemSelected;;

		if (m_searchStringInMail.IsEmpty())
			pMsgView->ClearSearchResultsInIHTMLDocument(pMsgView->m_searchID);
		else
		{
			pMsgView->FindStringInIHTMLDocument(m_searchStringInMail, m_bWholeWordInMail, m_bCaseSensInMail);
		}

		int dbg = 1;
	}
	return itemSelected;
}

int NListView::MailsWhichColumnSorted() const
{
	return MboxMail::b_mails_which_sorted;
}

void NListView::OnItemchangedListCtrl(NMHDR* pNMHDR, LRESULT* pResult)
{
	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;

	CString strNewState;
	CString strOldState;
	CString strChanged;
	ItemState2Str(pNMListView->uNewState, strNewState);
	ItemState2Str(pNMListView->uOldState, strOldState);
	ItemChange2Str(pNMListView->uChanged, strChanged);

	TRACE(L"\n\tOnItemchangedListCtrl.  Item index: %d SubItem index: %d NewState: %s OldState: %s strChanged: %s\n",
		pNMListView->iItem, pNMListView->iSubItem, (LPCWSTR)strNewState, (LPCWSTR)strOldState, (LPCWSTR)strChanged);

	PrintSelected();

	*pResult = 0;
}

void NListView::OnODStateChangedListCtrl(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMLVODSTATECHANGE lpStateChange = (LPNMLVODSTATECHANGE)pNMHDR;


	CString strNewState;
	CString strOldState;
	ItemState2Str(lpStateChange->uNewState, strNewState);
	ItemState2Str(lpStateChange->uOldState, strOldState);

	TRACE(L"\n\tOnODStateChangedListCtrl.  From Item index: %d To Item index: %d NewState: %s OldState: %s\n", 
		lpStateChange->iFrom, lpStateChange->iTo, (LPCWSTR)strNewState, (LPCWSTR)strOldState);

	PrintSelected();

	*pResult = 0;
}

void NListView::OnODFindItemListCtrl(NMHDR* pNMHDR, LRESULT* pResult)
{
	NMLVFINDITEM *lpFindItem = (NMLVFINDITEM*)pNMHDR;

	CString str;
	//str.Format(L"OnODFindItemListCtrl.  From Item index: %d To Item index: %d", lpStateChange->iFrom, lpStateChange->iTo);

	*pResult = 0;
}


void NListView::OnReleaseCaptureListCtrl(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMHDR lpReleaseCapture = (LPNMHDR)pNMHDR;

	CString str;
	//str.Format(L"OnReleaseCaptureListCtrl.  From Item index: %d To Item index: %d", lpReleaseCapture->iFrom, lpStateChange->iTo);

	*pResult = 0;
}

void NListView::OnODCacheHintListCtrl(NMHDR* pNMHDR, LRESULT* pResult)
{
	NMLVCACHEHINT  *lpCacheHint = (NMLVCACHEHINT*)pNMHDR;

	CString str;
	//str.Format(L"OnODCacheHintListCtrl.  From Item index: %d To Item index: %d", lpReleaseCapture->iFrom, lpStateChange->iTo);

	*pResult = 0;
}


void NListView::ItemState2Str(UINT uState, CString &strState)
{
	if (uState & LVIS_ACTIVATING) {
		strState.Append(L"LVIS_ACTIVATING");
	}
	if (uState & LVIS_CUT) {
		if (!strState.IsEmpty()) { strState.Append(L" | "); }
		strState.Append(L"LVIS_CUT");
	}
	if (uState & LVIS_DROPHILITED) {
		if (!strState.IsEmpty()) { strState.Append(L" | "); }
		strState.Append(L"LVIS_DROPHILITED");
	}
	if (uState & LVIS_FOCUSED) {
		if (!strState.IsEmpty()) { strState.Append(L" | "); }
		strState.Append(L"LVIS_FOCUSED");
	}
	if (uState & LVIS_OVERLAYMASK) {
		if (!strState.IsEmpty()) { strState.Append(L" | "); }
		strState.Append(L"LVIS_OVERLAYMASK");
	}
	if (uState & LVIS_SELECTED) {
		if (!strState.IsEmpty()) { strState.Append(L" | "); }
		strState.Append(L"LVIS_SELECTED");
	}
	if (uState & LVIS_STATEIMAGEMASK) {
		if (!strState.IsEmpty()) { strState.Append(L" | "); }
		strState.Append(L"LVIS_STATEIMAGEMASK");
	}
}


void NListView::ItemChange2Str(UINT uChange, CString &strState)
{
	strState.Empty();
#if (NTDDI_VERSION >= NTDDI_VISTA)
	if (uChange & LVIF_COLFMT) strState.Append(L"LVIF_COLFMT ");   //  require (NTDDI_VERSION >= NTDDI_VISTA)
#endif
	if (uChange & LVIF_COLUMNS) strState.Append(L"LVIF_COLUMNS ");
	if (uChange & LVIF_DI_SETITEM) strState.Append(L"LVIF_DI_SETITEM ");
	if (uChange & LVIF_GROUPID) strState.Append(L"LVIF_GROUPID ");
	if (uChange & LVIF_IMAGE) strState.Append(L"LVIF_IMAGE ");

	if (uChange & LVIF_INDENT) strState.Append(L"LVIF_INDENT ");
	if (uChange & LVIF_NORECOMPUTE) strState.Append(L"LVIF_NORECOMPUTE ");
	if (uChange & LVIF_PARAM) strState.Append(L"LVIF_PARAM ");
	if (uChange & LVIF_STATE) strState.Append(L"LVIF_STATE ");
	if (uChange & LVIF_TEXT) strState.Append(L"LVIF_TEXT ");

	if (strState.IsEmpty()) strState.Append(L"LVIS_UNKNOWN");
}

// Debug support
void NListView::PrintSelected()
{
	int deb = 1;
	// Get the selected items in the control

	POSITION p = m_list.GetFirstSelectedItemPosition();
	CString msg = L"";
	CString str;

	if (p)
		msg = L"The following items are selected: \n\t";

	while (p)
	{
		int nSelected = m_list.GetNextSelectedItem(p);

		str.Empty();
		str.Format(L"%d", nSelected);

		if (p)
			str += L", ";

		msg += str;
		if (msg.GetLength() > 100)
		{
			TRACE(L"%s\n", (LPCWSTR)msg);
			msg.Empty();
			msg.Append(L"cont:\n\t");
		}
	}
	if (msg.IsEmpty()) {
		msg = L"There are no selected items";
	}
	TRACE(L"%s\n", (LPCWSTR)msg);
}

BOOL NListView::IsUserSelectedMailListEmpty()
{
	int nWhichMailList = MboxMail::nWhichMailList;
	if (nWhichMailList == IDC_EDIT_LIST) 
	{
		return(MboxMail::s_mails.IsEmpty());
	}
	else
	{
		return(MboxMail::s_mails_edit.IsEmpty());
	}
}

void NListView::SwitchToMailList(int nID, BOOL force)
{
	// TODO: there is more optimization/simplication/cleanup when force is set to TRUE
	// force is set to TRUE when running FIND option while FIND list is active
	//
	// IDC_ARCHIVE_LIST                1056
	// IDC_FIND_LIST                   1057
	// IDC_EDIT_LIST                   1152
	// IDC_LABEL_LIST                  1197
	// IDC_FOLDER_LIST                 1191


	int nWhichMailList = MboxMail::nWhichMailList;
	if ((nWhichMailList == nID) && (force == FALSE))
		return;

	CMainFrame *pFrame = (CMainFrame*)AfxGetMainWnd();
	if (!pFrame)
		return;

	if (!::IsWindow(pFrame->m_hWnd) || !pFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
		return;

	if (nWhichMailList != nID)
	{
		// save current environment before switching to new mail list
		if (nWhichMailList == IDC_ARCHIVE_LIST)
		{
			MboxMail::s_mails_all.Copy(MboxMail::s_mails);
			MboxMail::m_allMails.m_lastSel = m_lastSel;
			MboxMail::m_allMails.b_mails_which_sorted = MboxMail::b_mails_which_sorted;
		}
		else if (nWhichMailList == IDC_FIND_LIST)
		{
			MboxMail::s_mails_find.CopyKeepData(MboxMail::s_mails);
			MboxMail::m_findMails.m_lastSel = m_lastSel;
			MboxMail::m_findMails.b_mails_which_sorted = MboxMail::b_mails_which_sorted;
		}
		else if (nWhichMailList == IDC_EDIT_LIST)
		{
			MboxMail::s_mails_edit.CopyKeepData(MboxMail::s_mails);
			MboxMail::m_editMails.m_lastSel = m_lastSel;
			MboxMail::m_editMails.b_mails_which_sorted = MboxMail::b_mails_which_sorted;
		}
		else if (nWhichMailList == IDC_LABEL_LIST)
		{
			MboxMail::s_mails_label.CopyKeepData(MboxMail::s_mails);
			MboxMail::m_labelMails.m_lastSel = m_lastSel;
			MboxMail::m_labelMails.b_mails_which_sorted = MboxMail::b_mails_which_sorted;
		}
		else if (nWhichMailList == IDC_FOLDER_LIST)
		{
			MboxMail::s_mails_folder.CopyKeepData(MboxMail::s_mails);
			MboxMail::m_folderMails.m_lastSel = m_lastSel;
			MboxMail::m_folderMails.b_mails_which_sorted = MboxMail::b_mails_which_sorted;
		}
		else
			; // TODO: _ASSERTE ?
	}

	// setup environment for new mail list
	if (nID == IDC_ARCHIVE_LIST)
	{
		MboxMail::s_mails.CopyKeepData(MboxMail::s_mails_all);
		m_lastSel = MboxMail::m_allMails.m_lastSel;
		MboxMail::m_mailList = &MboxMail::m_allMails;
		MboxMail::b_mails_which_sorted = MboxMail::m_allMails.b_mails_which_sorted;
	}
	else if (nID == IDC_FIND_LIST)
	{
		MboxMail::s_mails.CopyKeepData(MboxMail::s_mails_find);
		m_lastSel = MboxMail::m_findMails.m_lastSel;
		MboxMail::m_mailList = &MboxMail::m_findMails;
		MboxMail::b_mails_which_sorted = MboxMail::m_findMails.b_mails_which_sorted;
	}
	else if (nID == IDC_EDIT_LIST)
	{
		MboxMail::s_mails.CopyKeepData(MboxMail::s_mails_edit);
		m_lastSel = MboxMail::m_editMails.m_lastSel;
		MboxMail::m_mailList = &MboxMail::m_editMails;
		MboxMail::b_mails_which_sorted = MboxMail::m_editMails.b_mails_which_sorted;
	}
	else if (nID == IDC_LABEL_LIST)
	{
		MboxMail::s_mails.CopyKeepData(MboxMail::s_mails_label);
		m_lastSel = MboxMail::m_labelMails.m_lastSel;
		MboxMail::m_mailList = &MboxMail::m_labelMails;
		MboxMail::b_mails_which_sorted = MboxMail::m_labelMails.b_mails_which_sorted;
	}
	else if (nID == IDC_FOLDER_LIST)
	{
		MboxMail::s_mails.CopyKeepData(MboxMail::s_mails_folder);
		m_lastSel = MboxMail::m_folderMails.m_lastSel;
		MboxMail::m_mailList = &MboxMail::m_folderMails;
		MboxMail::b_mails_which_sorted = MboxMail::m_folderMails.b_mails_which_sorted;
	}
	else
		; // TODO: _ASSERTE ?

	if (abs(MboxMail::b_mails_which_sorted) == 99)
	{
		MboxMail::SortByGroupId(0, MboxMail::b_mails_which_sorted < 0);  // TODO: this is probably not needed; keep it simple for now
		MboxMail::assignColor2ConvesationGroups(&MboxMail::s_mails);
	}

	if ((abs(MboxMail::b_mails_which_sorted) == 4) && (m_subjectSortType == 1))
	{
		SortBySubjectBasedConversasions(TRUE);
	}

	m_list.SetItemCount(MboxMail::s_mails.GetCount());
	MboxMail::nWhichMailList = nID;
	pFrame->SetMailList(nID);


#if 0
	if (abs(MboxMail::b_mails_which_sorted) == 99)
		RefreshSortByColumn();
#endif

	int sv_lastSel = m_lastSel;
	ClearDescView();  // set m_lastSel to -1 ; m_lastFindPos is not updated
	m_lastSel = sv_lastSel;

	Invalidate();

	NMsgView *pMsgView = pFrame->GetMsgView();
	if (pMsgView)
		pMsgView->Invalidate(); 
		//pMsgView->m_browser.m_ie.Invalidate(); // TODO: changed to GetMsgView()->Invalidate(); and it seem to work :)

	int mailCnt = MboxMail::s_mails.GetCount();
	if (mailCnt > 0)
	{
		if ((m_lastSel >= mailCnt)  || (m_lastSel < 0))
			m_lastSel = 0;
		SelectItemFound(m_lastSel);  // m_lastFindPos to  m_lastSel
	}
	else
		m_lastSel = -1;

	m_lastFindPos = -1;  // TODO: can we preserve Find in process while mswitching ?
}

void NListView::OnEditFindadvanced()
{
	EditFindAdvanced();
}

// FIXMEFIXME How about wchar version ?
BOOL NListView::IsSingleAddress(CStringA *to)  
{
	int posBeg = 0;
	int posEnd = 0;

	posEnd = to->Find("@", posBeg);
	if ((posEnd >= 0) && ((posEnd + 1) < to->GetLength()))
	{
		posEnd = to->Find(",", posEnd + 1);
		if (posEnd < 0)
			return TRUE;
	}
	else
		return TRUE;

	// find out if the second addr is on the list
	if ((posEnd >= 0) && ((posEnd + 1) < to->GetLength()))
	{
		posEnd = to->Find("@", posEnd + 1);
		if (posEnd < 0)
			return TRUE;
		else
			return FALSE;
	}
	else
		return TRUE;
}

void NListView::TrimToAddr(CStringA *to, CStringA &toAddr, int maxNumbOfAddr)
{
	SimpleString name;
	SimpleString addr;
	CStringA tmpAddr;
	CStringA tmp;

	int posBeg = 0;
	int posEnd = 0;
	for (int i = 0; i < maxNumbOfAddr; )
	{
		posEnd = to->Find("@", posBeg);
		if ((posEnd >= 0) && ((posEnd + 1) < to->GetLength()))
			posEnd = to->Find(",", posEnd + 1);

		if (posEnd >= 0)
			tmp = to->Mid(posBeg, posEnd - posBeg);
		else
			tmp = to->Mid(posBeg, to->GetLength());

		int tolen = tmp.GetLength();
		name.ClearAndResize(tolen);
		addr.ClearAndResize(tolen);
		MboxMail::splitMailAddress((LPCSTR)tmp, tolen, &name, &addr);
		tmpAddr = addr.Data();
		tmpAddr.Trim(" \t\"<>()");
		toAddr.Append(tmpAddr);

		if (posEnd < 0)
			break;
		posBeg = posEnd + 1;
		i++;
		if (i < maxNumbOfAddr)
			toAddr.Append(",");
	}
}

void NListView::TrimToName(CStringA *to, CStringA &toName, int maxNumbOfAddr)
{
	SimpleString name;
	SimpleString addr;
	CStringA tmpName;
	CStringA tmp;

	int posBeg = 0;
	int posEnd = 0;
	for (int i = 0; i < maxNumbOfAddr; )
	{
		posEnd = to->Find("@", posBeg);
		if ((posEnd >= 0) && ((posEnd + 1) < to->GetLength()))
			posEnd = to->Find(",", posEnd + 1);

		if (posEnd >= 0)
			tmp = to->Mid(posBeg, posEnd - posBeg);
		else
			tmp = to->Mid(posBeg, to->GetLength());

		int tolen = tmp.GetLength();
		name.ClearAndResize(tolen);
		addr.ClearAndResize(tolen);
		MboxMail::splitMailAddress((LPCSTR)tmp, tolen, &name, &addr);

		tmpName = name.Data();
		tmpName.Trim(" \t\"<>()");
		toName.Append(tmpName);

		if (posEnd < 0)
			break;
		posBeg = posEnd + 1;
		i++;
		if (i < maxNumbOfAddr)
			toName.Append(",");
	}
}

BOOL NListView::PrepopulateAdvancedSearchParams(MboxMail *m)
{
	_ASSERTE(m != 0);
	if (m == 0)
		return FALSE;

	SimpleString name;
	SimpleString addr;
	CStringA tmpAddr;
	CStringA fromAddr;
	CStringA toAddr;
	CStringA subjAddr;
	CStringA subj;
	CStringA tmp;

	for (int i = 0; i < FILTER_FIELDS_NUMB; i++)   // FIXME i = 3 ??
	{
		m_advancedParams.m_string[i].Empty();
		m_advancedParams.m_stringA[i].Empty();
		m_advancedParams.m_charsetId[i] = 0;
		m_advancedParams.m_bEditChecked[i] = FALSE;
		m_advancedParams.m_bCaseSensitive[i] = FALSE;
		m_advancedParams.m_bWholeWord[i] = FALSE;
	}

	m_advancedParams.m_charsetId[0] = m->m_from_charsetId;
	m_advancedParams.m_charsetId[1] = m->m_to_charsetId;
	m_advancedParams.m_charsetId[2] = m->m_subj_charsetId;
	m_advancedParams.m_charsetId[3] = m->m_cc_charsetId;
	m_advancedParams.m_charsetId[4] = m->m_bcc_charsetId;

	int fromlen = m->m_from.GetLength();
	name.ClearAndResize(fromlen);
	addr.ClearAndResize(fromlen);
	MboxMail::splitMailAddress(m->m_from.operator LPCSTR(), fromlen, &name, &addr);
	fromAddr = addr.Data();
	fromAddr.Trim(" \t\"<>");

	DWORD error;
	CString fromAddrW;
	TextUtilsEx::Ansi2WStr(fromAddr, fromAddrW, error);

	m_advancedParams.m_string[0].Empty();
	m_advancedParams.m_string[0].Append(fromAddrW);
	if (!fromAddr.IsEmpty())
	{
		m_advancedParams.m_bEditChecked[0] = TRUE;
		m_advancedParams.m_bCaseSensitive[0] = FALSE;
	}
	else
		m_advancedParams.m_bEditChecked[0] = FALSE;

	int maxNumbOfAddr = 1;
	TrimToAddr(&m->m_to, toAddr, maxNumbOfAddr);

	CString toAddrW;
	TextUtilsEx::Ansi2WStr(toAddr, toAddrW, error);

	m_advancedParams.m_string[1].Empty();
	m_advancedParams.m_string[1].Append(toAddrW);
	if (!toAddr.IsEmpty())
	{
		m_advancedParams.m_bEditChecked[1] = TRUE;
		m_advancedParams.m_bCaseSensitive[1] = FALSE;
	}
	else
		m_advancedParams.m_bEditChecked[1] = FALSE;

	subj.Append(m->m_subj);

	int length = subj.GetLength();
	for (int i = 0; i < 10; i++) // to be safe, limit number of iterations  // FIXME
	{
		subj.TrimLeft();

		if (subj.GetLength() >= 5)
		{
			if ((strncmp((LPCSTR)subj, "Fwd: ", 5) == 0) || (strncmp((LPCSTR)subj, "FWD: ", 5) == 0))
				subj.Delete(0, 5);
			else if ((strncmp((LPCSTR)subj, "Re: ", 4) == 0) || (strncmp((LPCSTR)subj, "RE: ", 4) == 0))
				subj.Delete(0, 4);
		}
		else if (subj.GetLength() >= 4)
		{
			if ((strncmp((LPCSTR)subj, "Re: ", 4) == 0) || (strncmp((LPCSTR)subj, "RE: ", 4) == 0))
				subj.Delete(0, 4);
		}
		if (length == subj.GetLength())
			break;

		length = subj.GetLength();
	}
	subj.Trim();

	CString subjW;
	//TextUtilsEx::Ansi2WStr(subj, subjW, error);
	TextUtilsEx::CodePage2WStr(&subj, m_advancedParams.m_charsetId[2], &subjW, error);

	m_advancedParams.m_string[2].Empty();
	m_advancedParams.m_string[2].Append(subjW);
	if (!subj.IsEmpty())
	{
		m_advancedParams.m_bEditChecked[2] = TRUE;
		m_advancedParams.m_bCaseSensitive[2] = FALSE;
	}
	else
		m_advancedParams.m_bEditChecked[2] = FALSE;

	return TRUE;
}

void NListView::FindMinMaxTime(MyCTime &minTime, MyCTime &maxTime)
{
	minTime = -1;
	maxTime = -1;

	time_t min = -1;
	time_t max = 0;
	int sz = MboxMail::s_mails.GetSize();
	for (int i = 0; i < sz; i++)
	{
		time_t t = MboxMail::s_mails[i]->m_timeDate;
		if (min < 0)
		{
			if (t >= 0)
				min = t;
		}
		else if (t < min && t >= 0)
			min = t;
		if (t > max)
			max = t;
	}
	if (min <= 0) // TODO: need to handle < 0 better
		min = 0;
	minTime = MyCTime(min);

	if (max <= 0)
		maxTime = 0; // TODO: need to handle < 0 better
	else
	{
		maxTime = MyCTime(max);
	}

	if ((minTime.GetTime() < 0) || (maxTime.GetTime() < 0))
		int deb = 1;
}

void NListView::EditFindAdvanced(MboxMail *m)
{
	// TODO: Add your command handler code here

	if (m_bInFind)
		return;

	m_advancedFind = TRUE;

	m_bInFind = true;

	CFindAdvancedDlg dlg;

	// from && to && subject are set when Advanced Find is invoked as menu option on the selected item
	// when invoked via Edit -> Find Advanced menu option, they all will be set to NULL
	if (m)
	{
		PrepopulateAdvancedSearchParams(m);
	}

	// This code is too comples and error prone. Find time to fix this

	int nWhichMailList = MboxMail::nWhichMailList;

	if ((nWhichMailList != m_advancedParams.m_last_nWhichMailList) && (nWhichMailList != IDC_LABEL_LIST))
	{
		m_advancedParams.m_filterDates = FALSE;
		m_advancedParams.m_last_nWhichMailList = nWhichMailList;
	}

	if (m_advancedParams.m_mboxMailStartDate.GetTime() < 0)
	{
		_ASSERTE(m_advancedParams.m_bNeedToFindMailMinMaxTime == TRUE);
		if (m_maileFileInfo.m_oldestMailTime == m_maileFileInfo.m_latestMailTime)  // old version possibly
			FindMinMaxTime(m_advancedParams.m_mboxMailStartDate, m_advancedParams.m_mboxMailEndDate);
		else
		{
			m_advancedParams.m_mboxMailStartDate = m_maileFileInfo.m_oldestMailTime;
			m_advancedParams.m_mboxMailEndDate = m_maileFileInfo.m_latestMailTime;
		}
	}


	if ((MboxMail::nWhichMailList == IDC_ARCHIVE_LIST) || (MboxMail::nWhichMailList == IDC_LABEL_LIST))
	{
		if (m_advancedParams.m_bNeedToFindMailMinMaxTime)
		{
			m_advancedParams.m_lastStartDate = m_advancedParams.m_mboxMailStartDate;
			m_advancedParams.m_lastEndDate = m_advancedParams.m_mboxMailEndDate;

			m_advancedParams.m_archiveList_lastStartDate = m_advancedParams.m_lastStartDate;
			m_advancedParams.m_archiveList_lastEndDate = m_advancedParams.m_lastEndDate;
			m_advancedParams.m_archiveList_filterDates = m_advancedParams.m_filterDates;

			m_advancedParams.m_bNeedToFindMailMinMaxTime = FALSE;
		}
	}

	MailArchiveFileInfo* maileFileInfo = &m_maileFileInfo;

	_ASSERTE(m_advancedParams.m_mboxMailStartDate.GetTime() == m_maileFileInfo.m_oldestMailTime);
	_ASSERTE(m_advancedParams.m_mboxMailEndDate.GetTime() == m_maileFileInfo.m_latestMailTime);

	if ((MboxMail::nWhichMailList != IDC_ARCHIVE_LIST) && (MboxMail::nWhichMailList != IDC_LABEL_LIST))
	{
		FindMinMaxTime(m_advancedParams.m_lastStartDate, m_advancedParams.m_lastEndDate);
	}
	else
	{
		m_advancedParams.m_lastStartDate = m_advancedParams.m_archiveList_lastStartDate;
		m_advancedParams.m_lastEndDate = m_advancedParams.m_archiveList_lastEndDate;
		m_advancedParams.m_filterDates = m_advancedParams.m_archiveList_filterDates;
	}
	 
	BOOL retDT = MyCTimeToOleTime(m_advancedParams.m_lastStartDate, m_advancedParams.m_startDate, 1);  // DOWN
	retDT = MyCTimeToOleTime(m_advancedParams.m_lastEndDate, m_advancedParams.m_endDate, 2);  // UP

	dlg.m_params.Copy(m_advancedParams);

	if (dlg.DoModal() == IDOK) 
	{
		m_SearchType = 1;  // Advanced

		m_advancedParams.Copy(dlg.m_params);

		if (m_advancedParams.m_filterDates == TRUE)
		{
			// mails betwenn m_advancedParams.m_lastStartDate and m_advancedParams.m_lastEndDate will be processed
			retDT = NListView::OleTime2MyCTime(m_advancedParams.m_startDate, m_advancedParams.m_lastStartDate, FALSE);
			retDT = NListView::OleTime2MyCTime(m_advancedParams.m_endDate, m_advancedParams.m_lastEndDate, TRUE);

			if ((MboxMail::nWhichMailList == IDC_ARCHIVE_LIST) || (MboxMail::nWhichMailList == IDC_LABEL_LIST))
			{
				m_advancedParams.m_archiveList_lastStartDate = m_advancedParams.m_lastStartDate;
				m_advancedParams.m_archiveList_lastEndDate = m_advancedParams.m_lastEndDate;
				m_advancedParams.m_archiveList_filterDates = TRUE;
			}
		}
		else  // all mails will be processed
		{
			if ((MboxMail::nWhichMailList == IDC_ARCHIVE_LIST) || (MboxMail::nWhichMailList == IDC_LABEL_LIST))
			{
				m_advancedParams.m_archiveList_filterDates = FALSE;
			}
		}

		m_lastFindPos = -1;

		// To optimize content/message and/or attachments search performance
		if (((m_advancedParams.m_bEditChecked[3] && m_advancedParams.m_string[3])  ||
			(m_advancedParams.m_bEditChecked[4] && m_advancedParams.m_string[4])) && 
			(MboxMail::b_mails_sorted == true))
		{
			int whichSorted = abs(MboxMail::b_mails_which_sorted);
			if ((whichSorted != 1) && (whichSorted != 99)) { // related mails should be in close proximity in the mail file
				MboxMail::SortByDate();
				MboxMail::b_mails_sorted = true;
				MboxMail::b_mails_which_sorted = 1;

				RedrawMails();
			}
		}

		m_bHighlightAll = m_advancedParams.m_bHighlightAll;

		BOOL isTextUnicode = IsWindowUnicode(this->GetSafeHwnd());
		UINT localCP = GetACP();
		std::string str;
		BOOL ret = TextUtilsEx::id2charset(localCP, str);
		//  http://html-codes.info/ansi/html/Windows-1252-superscript%20one_185

		// m_bEditFindFirst = FALSE;  // this is applicable only to OnEditFind/OnEditFindAgain

		int sz = MboxMail::s_mails.GetSize();
		if (sz > 0)
		{
			int which = 0, w = -1;

			m_lastFindPos = 0;
			which = m_lastFindPos;

			//
			int maxSearchDuration = 0;
			maxSearchDuration = 0;

			_ASSERTE(m_lastFindPos >= 0);

			FIND_ARGS args;
			/*IN*/ args.lview = this; args.searchstart = m_lastFindPos;
			/*OUT*/ args.exitted = FALSE; args.retpos = -1; args.findAll = TRUE;
	
			{
				CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcFastSearch, (LPVOID)(FIND_ARGS*)&args);
				Dlg.SetDialogTemplate(AfxGetApp()->m_hInstance, MAKEINTRESOURCE(IDD_PROGRESS_DLG), IDC_STATIC, IDC_PROGRESS_BAR, IDCANCEL);

				ULONGLONG tc_start = GetTickCount64();

				INT_PTR nResult = Dlg.DoModal();

				if (!nResult) { // should never be true ? 
					m_bInFind = false;
					LRESULT lres = this->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
					MboxMail::assert_unexpected();
					// Or goto ??
					return;
				}

				ULONGLONG tc_end = GetTickCount64();
				DWORD delta = (DWORD)(tc_end - tc_start);
				CString total;
				total.Format(L"(EditFindAdvanced) Total search time %ld milliseconds !!!!!!!!!!!.\n", delta);
				TRACE(L"%s", (LPCWSTR)total);
				//MessageBox(L"info", total, MB_OK | MB_ICONHAND);

				int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
				int retResult = LOWORD(nResult);

				if (retResult != IDOK)
				{  // IDOK==1, IDCANCEL==2
					// We should be here when user selects Cancel button
					//_ASSERTE(cancelledbyUser == TRUE);

					DWORD terminationDelay = Dlg.GetTerminationDelay();
					int loopCnt = (terminationDelay+100)/25;

					ULONGLONG tc_start = GetTickCount64();
					while ((loopCnt-- > 0) && (args.exitted == FALSE))
					{
						Sleep(25);
					}
					ULONGLONG tc_end = GetTickCount64();
					DWORD delta = (DWORD)(tc_end - tc_start);
					TRACE(L"(EditFindAdvanced)Waited %ld milliseconds for thread to exist.\n", delta);
				}
				if (args.exitted == FALSE)
					MboxMail::s_mails_find.SetCountKeepData(0);

				w = args.retpos;
				MboxMail::pCUPDUPData = NULL;
			}
			if (MboxMail::s_mails_find.GetCount() > 0)
				MboxMail::m_findMails.m_lastSel = 0;
			else
				MboxMail::m_findMails.m_lastSel = -1;
			MboxMail::m_findMails.b_mails_which_sorted = MboxMail::b_mails_which_sorted;

			if (MboxMail::IsFindMailsSelected()) 
			{
				SwitchToMailList(IDC_FIND_LIST, TRUE);
			}
			else 
			{
				SwitchToMailList(IDC_FIND_LIST, FALSE);
			}
			m_bInFind = false;

			// FIXMEFIXME
			// To compare search results with non-unicode MBox Viewer
			//  int d = MboxMail::DumpMailSummaryToFile(&MboxMail::s_mails, MboxMail::s_mails.GetCount());
			// return;
		}
		else
		{
			MessageBeep(MB_OK);
			m_lastFindPos = -1;
		}
	}

	m_bInFind = false;
	LRESULT lres = this->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
}

int NListView::MatchHeaderFldSingleAddress(int fldIndx, CStringA &hdrFld, UINT hdrFldCodePage, CFindAdvancedParams &params, int pos)
{
	if (params.m_bSingleTo)
	{
		BOOL isSingle = IsSingleAddress(&hdrFld);
		if (isSingle)
			pos = MatchHeaderFld(fldIndx, hdrFld, hdrFldCodePage, params, pos);
		else
			pos = -2;
	}
	else
	{
		pos = MatchHeaderFld(fldIndx, hdrFld, hdrFldCodePage, params, pos);
	}
	return pos;
}

int NListView::MatchHeaderFld(int fldIndx, CStringA &hdrFld, UINT hdrFldCodePage, CFindAdvancedParams &params, int pos)
{
	if (params.m_bEditChecked[fldIndx])
	{
		int fldLength = m_stringWithCase[fldIndx].GetLength();
		CStringA &fld = m_stringWithCase[fldIndx];

		if (!hdrFld.IsEmpty() && (fldLength == 1 && (fld.GetAt(0) == '*')))
		{
			pos = 1;
		}
		else if (params.m_bWholeWord[fldIndx]) 
		{
			pos = g_tu.StrSearchW((unsigned char *)(LPCSTR)hdrFld, hdrFld.GetLength(), hdrFldCodePage,
				(unsigned char *)(LPCSTR)fld, fldLength, params.m_bCaseSensitive[fldIndx], 0);
		}
		else 
		{
			pos = g_tu.StrSearch((unsigned char *)(LPCSTR)hdrFld, hdrFld.GetLength(), hdrFldCodePage, 
				(unsigned char *)(LPCSTR)fld, fldLength, params.m_bCaseSensitive[fldIndx], 0);
		}
	}
	if (pos < 0) pos = -1;
	return pos;
}

#if 0
const char  *ruleText[] = {
	"((From <--> To) and CC and BCC and Subject and (Message Text or Attachment Text) and Attachment Name",
	"(From->To) and CC and BCC and Subject and (Message Text or Attachment Text) and Attachment Name",
	"((From -> (To or CC or BCC)) and Subject and (Message Text or Attachment Text) and Attachment Name",
	"((From <--> To) and no CC and no BCC and Subject and (Message Text or Attachment Text) and Attachment Name",
	"(From->To) and no CC and no BCC and Subject and (Message Text or Attachment Text) and Attachment Name",
	"(From or To or CC or BCC) and Subject and (Message Text or Attachment Text) and Attachment Name"
};
#endif


// TODO: needs some work to simplify and clarify
int NListView::CheckMatchAdvanced(int i, CFindAdvancedParams &params)
{
	int pos = -1;
	MboxMail *m = MboxMail::s_mails[i];

	if (i == 16445)
		int deb = 1;

#if 0
	// Just testing
	char * fldName = "Subject:";
	fldName = "To:";
	BOOL match = MatchIfFieldFolded(i, fldName);
	if (match)
		return i;
	else
		return - 1;
#endif

	int fromPos = 1;
	int toPos = 1;
	int fromRPos = 1;
	int toRPos = 1;
	int subjectPos = 1;
	int ccPos = 1;
	int bccPos = 1;

	int fldIndx;
	if (params.m_filterNumb == 0)
	{
		// "((From <--> To) and CC and BCC and Subject and (Message Text or Attachment Text) and Attachment Name"

		if (params.m_bEditChecked[0] || params.m_bEditChecked[1])
		{
			fldIndx = 0; // From
			fromPos = MatchHeaderFld(fldIndx, m->m_from, m->m_from_charsetId, params);

			fldIndx = 1; // To
			toPos = MatchHeaderFldSingleAddress(fldIndx, m->m_to, m->m_to_charsetId, params);

			if ((fromPos < 0) || (toPos < 0))
			{
				fldIndx = 1; // To search string is now From  // Reverse order of matching
				fromPos = MatchHeaderFld(fldIndx, m->m_from, m->m_from_charsetId, params);
				if (fromPos < 0)
					return -1;

				fldIndx = 0; // From search string is now To
				toPos = MatchHeaderFldSingleAddress(fldIndx, m->m_to, m->m_to_charsetId, params);
				if (toPos < 0)
					return -1;
			}
		}

		fldIndx = 3; // CC
		ccPos = MatchHeaderFld(fldIndx, m->m_cc, m->m_cc_charsetId, params);
		if (ccPos < 0)
			return -1;

		fldIndx = 4; // BCC
		bccPos = MatchHeaderFld(fldIndx, m->m_bcc, m->m_bcc_charsetId, params);
		if (bccPos < 0)
			return -1;
	}
	else if (params.m_filterNumb == 1)
	{
		// "(From->To) and CC and BCC and Subject and (Message Text or Attachment Text) and Attachment Name"

		fldIndx = 0; // From
		fromPos = MatchHeaderFld(fldIndx, m->m_from, m->m_from_charsetId, params);
		if (fromPos < 0)
			return -1;

		fldIndx = 1; // To
		toPos = MatchHeaderFldSingleAddress(fldIndx, m->m_to, m->m_to_charsetId, params);
		if (toPos < 0)
			return -1;

		if (params.m_bEditChecked[3] || params.m_bEditChecked[4])
		{
			fldIndx = 3; // CC
			ccPos = MatchHeaderFld(fldIndx, m->m_cc, m->m_cc_charsetId, params);
			if (ccPos < 0)
				return -1;

			fldIndx = 4; // BCC
			bccPos = MatchHeaderFld(fldIndx, m->m_bcc, m->m_bcc_charsetId, params);
			if (bccPos < 0)
				return -1;
		}
	}
	else if (params.m_filterNumb == 2)
	{
		// "((From -> (To or CC or BCC)) and Subject and (Message Text or Attachment Text) and Attachment Name"

		fldIndx = 0; // From
		fromPos = MatchHeaderFld(fldIndx, m->m_from, m->m_from_charsetId, params);
		if (fromPos < 0)
			return -1;

		// need special hack/check first because of "To or CC or BCC) and not "To and CC and BCC"
		// if none of the fields set, assume it is a match
		if (params.m_bEditChecked[1] || params.m_bEditChecked[3] || params.m_bEditChecked[4])
		{
			fldIndx = 1; // To
			toPos = -1;
			toPos = MatchHeaderFldSingleAddress(fldIndx, m->m_to, m->m_to_charsetId, params, toPos);

			if (toPos < 0)
			{
				fldIndx = 3; // CC
				ccPos = -1;
				ccPos = MatchHeaderFld(fldIndx, m->m_cc, m->m_cc_charsetId, params, ccPos);

				if (ccPos < 0)
				{
					fldIndx = 4; // BCC
					bccPos = -1;
					bccPos = MatchHeaderFld(fldIndx, m->m_bcc, m->m_bcc_charsetId, params, bccPos);
					if (bccPos < 0)
						return -1;
				}
			}
		}
	}
	else if (params.m_filterNumb == 3)
	{
		// "((From <--> To) and no CC and no BCC and Subject and (Message Text or Attachment Text) and Attachment Name"

		if (!m->m_cc.IsEmpty() || !m->m_bcc.IsEmpty())
			return -1;

		fldIndx = 0; // From
		fromPos = MatchHeaderFld(fldIndx, m->m_from, m->m_from_charsetId, params);

		fldIndx = 1; // To
		toPos = MatchHeaderFldSingleAddress(fldIndx, m->m_to, m->m_to_charsetId, params);

		if ((fromPos < 0) || (toPos < 0))
		{
			fldIndx = 1; // To search string is now From
			fromPos = MatchHeaderFld(fldIndx, m->m_from, m->m_from_charsetId, params);
			if (fromPos < 0)
				return -1;

			fldIndx = 0; // From search string is now To
			toPos = MatchHeaderFldSingleAddress(fldIndx, m->m_to, m->m_to_charsetId, params);
			if (toPos < 0)
				return -1;
		}
	}
	else if (params.m_filterNumb == 4)
	{
	// "(From->To) and no CC and no BCC and Subject and (Message Text or Attachment Text) and Attachment Name"

		if (!m->m_cc.IsEmpty() || !m->m_bcc.IsEmpty())
			return -1;

		fldIndx = 0; // From
		fromPos = MatchHeaderFld(fldIndx, m->m_from, m->m_from_charsetId, params);
		if (fromPos < 0)
			return -1;

		fldIndx = 1; // To
		toPos = MatchHeaderFldSingleAddress(fldIndx, m->m_to, m->m_to_charsetId, params);
		if (toPos < 0)
			return -1;
	}
	else if (params.m_filterNumb == 5)
	{
		// "(From or To or CC or BCC) and Subject and (Message Text or Attachment Text) and Attachment Name"

		// need special hack/check first because of "From or To or CC or BCC) and not "From and To and CC and BCC"
		// if none of the fields set, assume it is a match
		if (params.m_bEditChecked[0] || params.m_bEditChecked[1] || params.m_bEditChecked[3] || params.m_bEditChecked[4])
		{
			fldIndx = 0; // From
			fromPos = -1;
			fromPos = MatchHeaderFld(fldIndx, m->m_from, m->m_from_charsetId, params, fromPos);

			if (fromPos < 0)
			{
				fldIndx = 1; // To
				toPos = -1;
				toPos = MatchHeaderFldSingleAddress(fldIndx, m->m_to, m->m_to_charsetId, params, toPos);

				if (toPos < 0)
				{
					fldIndx = 3; // CC
					ccPos = -1;
					ccPos = MatchHeaderFld(fldIndx, m->m_cc, m->m_cc_charsetId, params, ccPos);

					if (ccPos < 0)
					{
						fldIndx = 4; // BCC
						bccPos = -1;
						bccPos = MatchHeaderFld(fldIndx, m->m_bcc, m->m_bcc_charsetId, params, bccPos);
						if (bccPos < 0)
							return -1;
					}
				}
			}
		}
	}

	fldIndx = 2; // subject
	if (params.m_bEditChecked[fldIndx])
	{
		int fldLength = m_stringWithCase[fldIndx].GetLength();
		CStringA &fld = m_stringWithCase[fldIndx];

		UINT inCodePage = m->m_subj_charsetId;

		if (!m->m_subj.IsEmpty() && ((fldLength == 1) && (fld.GetAt(0) == '*'))) {
			subjectPos = 1;
		}
		else if (params.m_bWholeWord[fldIndx]) {
			subjectPos = g_tu.StrSearchW((unsigned char *)(LPCSTR)m->m_subj, m->m_subj.GetLength(), inCodePage,
				(unsigned char *)(LPCSTR)fld, fldLength, params.m_bCaseSensitive[fldIndx], 0);
		}
		else {
			subjectPos = g_tu.StrSearch((unsigned char *)(LPCSTR)m->m_subj, m->m_subj.GetLength(), inCodePage,
				(unsigned char *)(LPCSTR)fld, fldLength, params.m_bCaseSensitive[fldIndx], 0);
		}
		if (subjectPos < 0) {
			return -1;
		}
	}

	fldIndx = 7; // attachment name
	if (params.m_bEditChecked[fldIndx])
	{
		if (FindAttachmentName(m, m_stringWithCase[fldIndx], params.m_bWholeWord[fldIndx], params.m_bCaseSensitive[fldIndx]) == FALSE) {
			return -1;
		}
	}

	if (params.m_bEditChecked[5] || params.m_bEditChecked[6])   // 5==message 6=content
	{
		if ((m_stringWithCase[5].GetLength() == 1 && m_stringWithCase[5][0] == '*') ||
			(m_stringWithCase[6].GetLength() == 1 && m_stringWithCase[6][0] == '*'))
		{
			;
		}
		else if (AdvancedFindInMailContent(i, params.m_bEditChecked[5], params.m_bEditChecked[6], params) == FALSE) {
			return -1;
		}
	}

	return i;
}


void NListView::OnUpdateEditFindadvanced(CCmdUI *pCmdUI)
{
	// TODO: Add your command update UI handler code here
	pCmdUI->Enable(!m_bInFind && (m_list.GetItemCount() > 0));
	return;
}

void NListView::RunFindAdvancedOnSelectedMail(int iItem)
{
	if (iItem >= MboxMail::s_mails.GetCount())
		return;

	MboxMail *m = MboxMail::s_mails[iItem];
	EditFindAdvanced(m);

	int deb = 1;
}

#define MAIL_LIST_VERSION_BASE  0x73215500
#define MAIL_LIST_VERSION  (MAIL_LIST_VERSION_BASE+3)
#define MAIL_LABEL_LIST_VERSION  (MAIL_LIST_VERSION_BASE+4)

int NListView::PopulateUserMailArray(SerializerHelper &sz, int mailListCnt, BOOL verifyOnly)
{
	int ret = 1;

	MboxMail *m;

	_int64 startOff;
	int length;
	int index;
	int contentDetailsArrayCnt;

	if (!verifyOnly)
		MboxMail::s_mails_edit.SetSize(mailListCnt);

	int i = 0;
	for (i = 0; i < mailListCnt; i++)
	{
		if (!sz.readInt64(&startOff))
			break;
		if (!sz.readInt(&length))
			break;
		if (!sz.readInt(&index))
			break;
		if (!sz.readInt(&contentDetailsArrayCnt))
			break;

		if ((index < 0) || (index >= MboxMail::s_mails_ref.GetCount()))
		{
			ret = -1;
			break;
		}

		m = MboxMail::s_mails_ref[index];
		if ((m->m_startOff != startOff) ||
			(m->m_length != length) ||
			(m->m_index != index) ||
			(m->m_ContentDetailsArray.size() != contentDetailsArrayCnt))
		{
			ret = -1;
			break;
		}

		if (!verifyOnly)
			MboxMail::s_mails_edit[i] = m;

		int deb = 1;
	}

	if (!verifyOnly) 
	{
		if (ret > 0)
			MboxMail::s_mails_edit.SetSizeKeepData(i);
		else
			MboxMail::s_mails_edit.SetSizeKeepData(0);
	}


	return ret;
}

int NListView::OpenArchiveFileLocation()
{
	CString mboxFileSuffix;
	if (MboxMail::IsUserMailsSelected())
		mboxFileSuffix = L"_USER.mbox";
	else if (MboxMail::IsFindMailsSelected())
		mboxFileSuffix = L"_FIND.mbox";
	else
	{
		// We should never be here
		return -1;
	}

	CString mailFile = MboxMail::s_path;
	CString datapath = MboxMail::GetLastDataPath();

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(mailFile, driveName, directory, fileNameBase, fileNameExtention);

	CString errorText;
	CString archiveCachePath;
	CString rootPrintSubFolder = L"ArchiveCache";
	CString targetPrintSubFolder;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, archiveCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CString mboxFile = fileNameBase + fileNameExtention;
	CString mboxFilePath = archiveCachePath + L"\\" + fileNameBase + mboxFileSuffix;

	//CString path = MboxMail::GetLastPath();
	//CString mboxFile = path + L"\\" + fileNameBase + mboxFileSuffix;

	if (!FileUtils::PathFileExist(mboxFilePath))
	{
#if 0
		CString txt = L"File\n\n\"" + mboxFilePath;
		txt += L"\"\n\ndoesn't exist. Can't open location.";
#endif
		CString txt;
		CString fmt = L"File\n\n\"%s\"\n\ndoesn't exist. Can't open location.";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, mboxFilePath);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		//return -1;  // Go ahead and open folder without highlighting merge file
	}
	if (FileUtils::BrowseToFile(mboxFilePath) == FALSE) {  // TODO: s_path error checking ??
		HWND h = GetSafeHwnd();
		HINSTANCE result = ShellExecute(h, L"open", archiveCachePath, NULL, NULL, SW_SHOWNORMAL);
		CMainFrame::CheckShellExecuteResult(result, h);
	}
	return 1;
}

int NListView::OpenMailListFileLocation()
{
	CString mboxFileSuffix;
	if (MboxMail::IsUserMailsSelected())
		mboxFileSuffix = L"_USER.mbox.mboxlist";
	else if (MboxMail::IsFindMailsSelected())
		mboxFileSuffix = L"_FIND.mbox.mboxlist";
	else
	{
		// We should never be here
		return -1;
	}

	CString mailFile = MboxMail::s_path;
	CString datapath = MboxMail::GetLastDataPath();

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(mailFile, driveName, directory, fileNameBase, fileNameExtention);

	CString errorText;
	CString archiveCachePath;
	CString rootPrintSubFolder = L"ListCache";
	CString targetPrintSubFolder;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, archiveCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CString mboxFile = fileNameBase + fileNameExtention;
	CString mboxFilePath = archiveCachePath + L"\\" + fileNameBase + mboxFileSuffix;

	//CString path = MboxMail::GetLastPath();
	//CString mboxFile = path + L"\\" + fileNameBase + mboxFileSuffix;

	if (!FileUtils::PathFileExist(mboxFilePath))
	{
#if 0
		CString txt = L"File \"" + mboxFilePath;
		txt += L"\" doesn't exist. Can't open location.";
#endif
		CString txt;
		CString fmt = L"File\n\n\"%s\"\n\ndoesn't exist. Can't open location.";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, mboxFilePath);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}
	if (FileUtils::BrowseToFile(mboxFilePath) == FALSE) {  // TODO: s_path error checking ??
		HWND h = GetSafeHwnd();
		HINSTANCE result = ShellExecute(h, L"open", archiveCachePath, NULL, NULL, SW_SHOWNORMAL);
		CMainFrame::CheckShellExecuteResult(result, h);
	}
	return 1;
}

// Below ADDED in 1.0.3.1
int NListView::RemoveDuplicateMails(MailArray &s_mails_array)
{
	m_list.SetRedraw(FALSE);

	BOOL putDuplicatesOnFindArray = TRUE;
	MboxMail::RemoveDuplicateMails(s_mails_array, putDuplicatesOnFindArray);

	RedrawMails();

	if (MboxMail::s_mails.GetCount() > 0)
	{
		SelectItemFound(0);
	}
	else
		m_lastSel = -1;

	// TODO: SelectItemFound sets  m_lastFindPos. Need to reset to 0 ?
	m_lastFindPos = -1;
	m_bEditFindFirst = TRUE;

	return 1;
}

int NListView::RefreshMailsOnUserSelectsMailListMark()
{
	MboxMail *m;
	int markCnt = 0;
	int errorCnt = 0;
	int i;
	for (i = 0; i < MboxMail::s_mails_ref.GetCount(); i++)
	{
		m = MboxMail::s_mails_ref[i];
		if (m->m_isOnUserSelectedMailList == true)
			markCnt++;
		m->m_isOnUserSelectedMailList = false;
	}

	if (MboxMail::IsUserMailsSelected())
	{
		int i;
		for (i = 0; i < MboxMail::s_mails.GetCount(); i++)
		{
			m = MboxMail::s_mails[i];
			m->m_isOnUserSelectedMailList = true;
		}
		errorCnt = MboxMail::s_mails.GetCount() - markCnt;
	}
	else
	{
		int i;
		for (i = 0; i < MboxMail::s_mails_edit.GetCount(); i++)
		{
			m = MboxMail::s_mails_edit[i];
			m->m_isOnUserSelectedMailList = true;
		}
		errorCnt = MboxMail::s_mails_edit.GetCount() - markCnt;
	}
	errorCnt = VerifyMailsOnUserSelectsMailListMarkCounts();
	if (errorCnt != 0)
		int deb = 1;

	return errorCnt;
}

int NListView::VerifyMailsOnUserSelectsMailListMarkCounts()
{
#ifndef _DEBUG
	return 1;
#endif

	MboxMail *m;
	int markCnt = 0;
	int errorCnt = 0;
	int i;
	for (i = 0; i < MboxMail::s_mails_ref.GetCount(); i++)
	{
		m = MboxMail::s_mails_ref[i];
		if (m->m_isOnUserSelectedMailList == true)
			markCnt++;
	}

	if (MboxMail::IsUserMailsSelected())
	{
		errorCnt = MboxMail::s_mails.GetCount() - markCnt;
	}
	else
	{
		errorCnt = MboxMail::s_mails_edit.GetCount() - markCnt;
	}
	if (errorCnt != 0)
	{
		_ASSERTE(FALSE);
		int deb = 1;
	}
	return errorCnt;
}

int NListView::PrintMailConversationToSeparatePDF_Thread(int iItem, BOOL mergePDFs, CString &errorText)
{
	int firstMail = 0;
	int lastMail = 0;

	// Convert range to list to avoid enhancancing PrintMailRangeToSeparatePDF_Thread
	FindFirstAndLastMailOfConversation(iItem, firstMail, lastMail);

	MailIndexList* selectedMailsIndexList = PopulateSelectedMailsListFromRange(firstMail, lastMail);

	CString targetPrintFolderPath;
	CString targetPrintSubFolderName = L"PDF_GROUP";
	int ret = PrintMailSelectedToSeparatePDF_Thread(selectedMailsIndexList, targetPrintSubFolderName, targetPrintFolderPath, mergePDFs);

	return ret;
}

int NListView::PrintMailConversationToSinglePDF_Thread(int iItem, CString &errorText)
{
	int firstMail = 0;
	int lastMail = 0;

	FindFirstAndLastMailOfConversation(iItem, firstMail, lastMail);

	CString targetPrintSubFolderName;
	int ret = PrintMailRangeToSinglePDF_Thread(firstMail, lastMail, targetPrintSubFolderName);

	return ret;
}

int NListView::PrintMailRangeToSeparatePDF_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName, CString& targetPrintFolderPath, BOOL mergePDFs)
{
	MailIndexList* selectedMailsIndexList = PopulateSelectedMailsListFromRange(firstMail, lastMail);

	int ret = PrintMailSelectedToSeparatePDF_Thread(selectedMailsIndexList, targetPrintSubFolderName, targetPrintFolderPath, mergePDFs);
	return ret;
}

int NListView::PrintMailRangeToSinglePDF_Thread(int firstMail, int lastMail, CString& targetPrintSubFolderName)
{
	MailIndexList* selectedMailsIndexList = PopulateSelectedMailsListFromRange(firstMail, lastMail);

	CString targetPrintFolderPath;
	int ret = PrintMailSelectedToSinglePDF_Thread(selectedMailsIndexList, targetPrintSubFolderName, targetPrintFolderPath);

	return ret;
}

int NListView::PrintMailSelectedToSeparatePDF_Thread(MailIndexList* selectedMailsIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, BOOL mergePDFs)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS args;

	CString rootPrintSubFolder = L"PrintCache";
	CString targetPrintSubFolder = targetPrintSubFolderName;
	CString printCachePath;
	CString errorText;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (pFrame)
	{
		int ret = pFrame->VerifyPathToHTML2PDFExecutable(errorText);
		if (ret < 0)
		{
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else
	{
		HWND h = GetSafeHwnd();
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		errorText.Append(L"Internal error. Try again.");
		return -1;
	}

	if (selectedMailsIndexList == 0)
	{
			selectedMailsIndexList = PopulateSelectedMailsList();
	}

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, printCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = GetSafeHwnd(); // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	/*IN*/
	args.lview = this;
	args.targetPrintFolderPath = printCachePath;
	args.targetPrintSubFolderName = targetPrintSubFolder;
	args.firstMail = -1;
	args.lastMail = -1;
	args.selectedMailIndexList = selectedMailsIndexList;
	args.nItem = -1;
	args.separatePDFs = TRUE;
	args.mergePDFs = mergePDFs;
	args.mergedPDFPath.Empty();

	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcPrintMailGroupToSeparatePDF, (LPVOID)(PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS*)&args);
	Dlg.SetDialogTemplate(AfxGetApp()->m_hInstance, MAKEINTRESOURCE(IDD_PROGRESS_DLG), IDC_STATIC, IDC_PROGRESS_BAR, IDCANCEL);

	INT_PTR nResult = Dlg.DoModal();

	if (!nResult) { // should never be true ?
		MboxMail::assert_unexpected();
		return -1;
	}

	int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
	int retResult = LOWORD(nResult);

	if (retResult != IDOK)
	{  // IDOK==1, IDCANCEL==2
		// We should be here when user selects Cancel button
		//_ASSERTE(cancelledbyUser == TRUE);

		DWORD terminationDelay = Dlg.GetTerminationDelay();
		int loopCnt = (terminationDelay+100)/25;

		ULONGLONG tc_start = GetTickCount64();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		ULONGLONG tc_end = GetTickCount64();
		DWORD delta = (DWORD)(tc_end - tc_start);
		TRACE(L"(PrintMailSelectedToSeparatePDF_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret > 0)
	{
		if (args.errorText.IsEmpty())
		{
			if (mergePDFs == FALSE)
			{
				//CString txt = L"Created separate PDF files in \n\n" + printCachePath + L" \n\ndirectory.";
				CString txt = L"Created separate PDF files in the following directory";
				ResHelper::TranslateString(txt); // RESTODO

				OpenContainingFolderDlg dlg(txt, printCachePath, TRUE);
				INT_PTR nResponse = dlg.DoModal();
				if (nResponse == IDOK)
				{
					HWND h = GetSafeHwnd();
					HINSTANCE result = ShellExecute(h, L"open", printCachePath, NULL, NULL, SW_SHOWNORMAL);
					CMainFrame::CheckShellExecuteResult(result, h);
					int deb = 1;
				}
				else if (nResponse == IDCANCEL)
				{
					int deb = 1;
				}
			}
			else
			{
				CString mergeErrorFilepath = printCachePath + L"\\MergePDFsError.log";
				_int64 mergeErrorFileSize = FileUtils::FileSize(mergeErrorFilepath);

				HWND h = GetSafeHwnd();
				if (mergeErrorFileSize > 0)
				{
#if 0
					CString text = L"Merge File was created.\n"
						L"However, one or more mails didn't finish printing to PDF in 5 minutes and were not merged\n"
						L"Select OK to open the file with the list of failed mails\n";
#endif

					CString txt = L"Merge File was created.\n"
						L"However, one or more mails didn't finish printing to PDF in 5 minutes and were not merged\n"
						L"Select OK to open the file with the list of failed mails\n";
					ResHelper::TranslateString(txt);

					int answer = ::MessageBox(h, txt, L"Warning", MB_APPLMODAL | MB_ICONWARNING | MB_OK);

					HINSTANCE result = ShellExecute(h, L"open", mergeErrorFilepath, NULL, NULL, SW_SHOWNORMAL);
					CMainFrame::CheckShellExecuteResult(result, h, &mergeErrorFilepath);
				}

				CString pdfFileName = args.mergedPDFPath;

				//CString txt = L"Created PDF file \n\n" + pdfFileName;
				CString txt = L"Created PDF file";
				ResHelper::TranslateString(txt);  // RESTODO
				OpenContainingFolderDlg dlg(txt, pdfFileName, FALSE);
				INT_PTR nResponse = dlg.DoModal();
				////////////
				if (nResponse == IDOK)
				{
					if (FileUtils::BrowseToFile(pdfFileName) == FALSE) {
						HWND h = GetSafeHwnd();
						HINSTANCE result = ShellExecute(h, L"open", printCachePath, NULL, NULL, SW_SHOWNORMAL);
						CMainFrame::CheckShellExecuteResult(result, h);
					}
					int deb = 1;
				}
				else if (nResponse == IDYES)
				{
					HWND h = GetSafeHwnd();
					HINSTANCE result = ShellExecute(h, L"open", pdfFileName, NULL, NULL, SW_SHOWNORMAL);
					CMainFrame::CheckShellExecuteResult(result, h, &pdfFileName);
					int deb = 1;
				}
				else if (nResponse == IDCANCEL)
				{
					int deb = 1;
				}
			}
		}
		else
		{
			MboxMail::assert_unexpected();
			HWND h = GetSafeHwnd();

			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}
	return 1;
}

int NListView::PrintMailSelectedToSinglePDF_Thread(MailIndexList* selectedMailsIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS args;

	CString rootPrintSubFolder = L"PrintCache";
	CString targetPrintSubFolder = targetPrintSubFolderName;
	CString printCachePath;
	CString errorText;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (pFrame)
	{
		int ret = pFrame->VerifyPathToHTML2PDFExecutable(errorText);
		if (ret < 0)
		{
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else
	{
		HWND h = GetSafeHwnd();
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		errorText.Append(L"Internal error. Try again.");
		return -1;
	}

	if (selectedMailsIndexList == 0)
	{
		selectedMailsIndexList = PopulateSelectedMailsList();
	}

	if (selectedMailsIndexList->GetCount() <= 0)
		return 1;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, printCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = GetSafeHwnd(); // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	/*IN*/
	args.lview = this;
	args.targetPrintFolderPath = printCachePath;
	args.targetPrintSubFolderName = targetPrintSubFolder;
	args.firstMail = -1;
	args.lastMail = -1;
	args.selectedMailIndexList = selectedMailsIndexList;
	args.nItem = -1;
	args.separatePDFs = FALSE;

	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcPrintMailGroupToSinglePDF, (LPVOID)(PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS*)&args);
	Dlg.SetDialogTemplate(AfxGetApp()->m_hInstance, MAKEINTRESOURCE(IDD_PROGRESS_DLG), IDC_STATIC, IDC_PROGRESS_BAR, IDCANCEL);

	INT_PTR nResult = Dlg.DoModal();

	if (!nResult) { // should never be true ?
		MboxMail::assert_unexpected();
		return -1;
	}

	int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
	int retResult = LOWORD(nResult);

	if (retResult != IDOK)
	{  // IDOK==1, IDCANCEL==2
		// We should be here when user selects Cancel button
		//_ASSERTE(cancelledbyUser == TRUE);

		DWORD terminationDelay = Dlg.GetTerminationDelay();
		int loopCnt = (terminationDelay+100)/25;

		ULONGLONG tc_start = GetTickCount64();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		ULONGLONG tc_end = GetTickCount64();
		DWORD delta = (DWORD)(tc_end - tc_start);
		TRACE(L"(PrintMailSelectedToSinglePDF_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret > 0)
	{
		if (args.errorText.IsEmpty())
		{
			int fileType = 2;  // 0=.txt 1=.htm 2=.pdf 3=.csv
			bool fileExists;
			CString pdfFileName;

			// Prone to inconsistency, fileName is also determine in ALongRightProcessProcPrintMailGroupToSingleHTML above
			if (selectedMailsIndexList->GetCount() > 1)
				int ret = MboxMail::MakeFileNameFromMailArchiveName(fileType, pdfFileName, targetPrintSubFolderName, fileExists, errorText);
			else
				int ret = MboxMail::MakeFileNameFromMailHeader(selectedMailsIndexList->GetAt(0), fileType, pdfFileName, targetPrintSubFolderName, fileExists, errorText);

			//CString txt = L"Created PDF file \n\n" + pdfFileName;
			CString txt = L"Created PDF file";
			ResHelper::TranslateString(txt);

			OpenContainingFolderDlg dlg(txt, pdfFileName, FALSE);
			INT_PTR nResponse = dlg.DoModal();
			////////////
			if (nResponse == IDOK)
			{
				if (FileUtils::BrowseToFile(pdfFileName) == FALSE) {
					HWND h = GetSafeHwnd();
					HINSTANCE result = ShellExecute(h, L"open", printCachePath, NULL, NULL, SW_SHOWNORMAL);
					CMainFrame::CheckShellExecuteResult(result, h);
				}
				int deb = 1;
			}
			else if (nResponse == IDYES)
			{
				HWND h = GetSafeHwnd();
				HINSTANCE result = ShellExecute(h, L"open", pdfFileName, NULL, NULL, SW_SHOWNORMAL);
				CMainFrame::CheckShellExecuteResult(result, h, &pdfFileName);
				int deb = 1;
			}
			else if (nResponse == IDCANCEL)
			{
				int deb = 1;
			}

		}
		else
		{
			MboxMail::assert_unexpected();
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}
	return 1;
}

// Print selected set of emails to a seprate PDF files.
// Optionally, merge created PDF files into a single PDF file. Merging is controlled by BOOL mergePDFs and  CString mergedFileName
//
int NListView::PrintMailSelectedToSeparatePDF_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText,
	BOOL mergePDFs, CString &mergedFileName)
{
	BOOL progressBar = TRUE;

	if (selectedMailIndexList->GetCount() <= 0)
		return 1;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	// TODO: Just delete folder instead ?
	int retdel = NListView::DeleteAllHtmAndPDFFiles(targetPrintFolderPath);

	// delete files created during last merge
	CString fpath = targetPrintFolderPath + L"\\MergePDFsError.log";
	BOOL r = FileUtils::DelFile(fpath);
	fpath = targetPrintFolderPath + L"\\MergePDFs.log";
	r = FileUtils::DelFile(fpath);
	fpath = targetPrintFolderPath + L"\\SelectedMailsList.txt";
	r = FileUtils::DelFile(fpath);
	// Delete cmd files MergePDFs.cmd and MergeAllPDFs.cmd ?   // FIXME

	///////////////////////////////////////////////
	CFile fp; CFile fpm; CFile fptxt; CFile fpMergeError;

	CString pdfboxJarFileName;
	CString mergeCmdFilePath;
	CString mergeAllCmdFilePath;

	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	if (mergePDFs)
	{
		CString processExePath = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"processPath");

		CString processFolderPath;
		FileUtils::CPathGetPath(processExePath, processFolderPath);

		// build a pdfbox-app tool wildcard string
		CString strWildcard(processFolderPath);
		strWildcard += L"\\pdfbox-app-*.jar";

		CString appNamePrefix = L"pdfbox-app-";

		CString fPath;
		CString fName;

		// Make sure pdfbox-app-*.jar file exists in MBox Viewer process directory
		CFileFind finder;
		BOOL bWorking = finder.FindFile(strWildcard);
		while (bWorking)
		{
			bWorking = finder.FindNextFile();
			if (finder.IsDots())
				continue;

			CString fPath = finder.GetFilePath();
			CString fName = finder.GetFileName();
			CString pref = fName.Left(appNamePrefix.GetAllocLength());
			if (pref.CompareNoCase(appNamePrefix) == 0)
			{
				wchar_t v = fName.GetAt(appNamePrefix.GetAllocLength());
				if ((v =='4') || (v == '3'))
				{
					pdfboxJarFileName = fName;
					break;
				}
			}
		}

		if (pdfboxJarFileName.IsEmpty())
		{
			CString fmt = L"Did not find  \"pdfbox-app-3.*.jar\" file in\n\n\"%s\"\n\ndirectory.";
			fmt.Append(L"\n\nPlease review the help information by selecting \"Print to->PDF->Help\" option.");
			ResHelper::TranslateString(fmt);
			errorText.Format(fmt, processFolderPath);

			TRACE(L"%s\n", errorText);
			return -1;
		}

		// Create empty MergePDFsError.log file for appending. Close after creation.
		CFileException ExMergeError;
		CString mergeErrorFilePath = targetPrintFolderPath + L"\\MergePDFsError.log";
		if (!fpMergeError.Open(mergeErrorFilePath, CFile::modeReadWrite | CFile::modeCreate | CFile::shareDenyNone,
			&ExMergeError))
		{
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExMergeError);

			errorText = L"Could not create \"" + mergeErrorFilePath;
			errorText += L"\" file.\n";
			errorText += exErrorStr;

			TRACE(L"%s\n", errorText);
			return -1;
		}
		fpMergeError.Close();  // Write Append in 

		// Make sure MergePDFs.cmd can be created and then close
		CFileException ExError;
		// We can't merge all PDF files in a single step due to limit on the max length of command line arguments
		// MergePDFs.cmd file will merge subset of PDF files
		// Merging will continue until there is no more files to merge, see NListView::MergePDfFileList
		//
		mergeCmdFilePath = targetPrintFolderPath + L"\\MergePDFs.cmd";
		if (!fp.Open(mergeCmdFilePath, CFile::modeReadWrite | CFile::modeCreate | CFile::shareDenyNone,
			&ExError)) 
		{
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			errorText = L"Merge PDF files failed. Could not create \"" + mergeCmdFilePath;
			errorText += L"\" file.\n";
			errorText += exErrorStr;

			TRACE(L"%s\n", errorText);
			return -1;
		}
		fp.Close();

		// Make sure MergeAllPDFs.cmd can be created and keep open
		// The MergeAllPDFs.cmd will contain all steps to create merged PDF file.
		// The MergeAllPDFs.cmd is not excuted by MBox Viewer. User can execute manually if needed.
		//
		mergeAllCmdFilePath = targetPrintFolderPath + L"\\MergeAllPDFs.cmd";
		if (!fpm.Open(mergeAllCmdFilePath, CFile::modeReadWrite | CFile::modeCreate | CFile::shareDenyNone,
			&ExError))
		{
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			errorText = L"Merge PDF files failed. Could not create \"" + mergeAllCmdFilePath;
			errorText += L"\" file.\n";
			errorText += exErrorStr;

			TRACE(L"%s\n", errorText);
			return -1;
		}
	}  // end of if (mergePDFs)

	if (progressBar && MboxMail::pCUPDUPData)
	{
		CString txt = L"Printing mails to PDF files ...";
		ResHelper::TranslateString(txt);
		MboxMail::pCUPDUPData->SetProgress(txt, 1);
	}

	UINT curstep = 1;
	CString fileNum;
	int nFileNum;

	if (progressBar && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	int i;
	int cnt = (int)selectedMailIndexList->GetCount();

	ULONGLONG workRangeFirstPos = 0;
	ULONGLONG workRangeLastPos = cnt - 1;
	ProgressTimer progressTimer(workRangeFirstPos, workRangeLastPos);

	BOOL progress_Bar = TRUE;  // FIXME
	for (int j = 0; j < cnt; j++)
	{
		if (progressBar && MboxMail::pCUPDUPData)
		{
			UINT_PTR dwProgressbarPos = 0;
			ULONGLONG workRangePos = j;

			nFileNum = j + 1;

			CString fmt = L"Printing mails to PDF files ... %d of %d";
			ResHelper::TranslateString(fmt);
			fileNum.Format(fmt, nFileNum, cnt);

			//if (MboxMail::pCUPDUPData)  MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(dwProgressbarPos));  // FIXME
			if (MboxMail::pCUPDUPData)  MboxMail::pCUPDUPData->SetProgress(fileNum);

			int debug = 1;
		}

		i = (*selectedMailIndexList)[j];
		if (pFrame)
		{
			pFrame->m_NamePatternParams.m_bConvert2Ansi = FALSE;
			if (mergePDFs)
				pFrame->m_NamePatternParams.m_bConvert2Ansi = TRUE;
			// progressText mac length = 43 characters
			int len = fileNum.GetLength();
			int lenAvail = 43 - len;  // space available for seconds counter
			CString progressText = fileNum + L"    ";
			pFrame->PrintSingleMailtoPDF(i, targetPrintSubFolderName, progress_Bar, progressText, errorText);
			pFrame->m_NamePatternParams.m_bConvert2Ansi = FALSE;
		}

		if (progressBar && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				if (mergePDFs)
					fpm.Close();
				return -1;
			}

			UINT_PTR dwProgressbarPos = 0;
			ULONGLONG workRangePos = j+1;
			BOOL needToUpdateStatusBar = progressTimer.UpdateWorkPos(workRangePos, dwProgressbarPos, TRUE);
			if (needToUpdateStatusBar)
			{
				nFileNum = j + 1;

				CString fmt = L"Printing mails to PDF files ... %d of %d";
				ResHelper::TranslateString(fmt);
				fileNum.Format(fmt, nFileNum, cnt);

				if (MboxMail::pCUPDUPData)  MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(dwProgressbarPos));

				int debug = 1;
			}
		}
	}
	nFileNum = cnt;
	UINT newstep = 100;

	CString fmt = L"Printing mails to PDF files ... %d of %d";
	ResHelper::TranslateString(fmt);
	fileNum.Format(fmt, nFileNum, cnt);


	if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(newstep));

	// Printing all mails to separate PDF file done.
	if (mergePDFs)
	{
		CString processExePath = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"processPath");

		CString processFolderPath;
		FileUtils::CPathGetPath(processExePath, processFolderPath);

		CString pdfboxJarFilePath = processFolderPath + L"\\" + pdfboxJarFileName;
		CString targetPdfboxJarFilePath = targetPrintFolderPath + L"\\" + pdfboxJarFileName;

		BOOL bFailIfExists = TRUE;
		if (!::CopyFileW(pdfboxJarFilePath, targetPdfboxJarFilePath, bFailIfExists))
		{
			CString errText = FileUtils::GetLastErrorAsString();
			TRACE(L"CopyFile: %s\n%s", pdfboxJarFilePath, errText);
		}
		//
		CString pdfboxConfigFilePath = processFolderPath + L"\\scripts\\pdfbox-config.txt";
		CString targetPdfboxConfigFilePath = targetPrintFolderPath + L"\\pdfbox-config.txt";

		if (!::CopyFileW(pdfboxConfigFilePath, targetPdfboxConfigFilePath, bFailIfExists))
		{
			CString errText = FileUtils::GetLastErrorAsString();
			TRACE(L"CopyFile: %s\n%s", pdfboxConfigFilePath, errText);
		}
		//
		CString mergePDFsLogFilePath = targetPrintFolderPath + L"\\MergePDFs.log";
		CString pdfboxLogFilePath = targetPrintFolderPath + L"\\pdfbox.log";

		FileUtils::DelFile(mergePDFsLogFilePath);
		FileUtils::DelFile(pdfboxLogFilePath);

		CString selectedMailsList = targetPrintFolderPath + L"\\SelectedMailsList.txt";
		BOOL selectedMailsListFileOpen = FALSE;
		CFileException ExError;
		if (!fptxt.Open(selectedMailsList, CFile::modeReadWrite | CFile::modeCreate | CFile::shareDenyNone,
			&ExError))
		{
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			errorText = L"Merge PDF files failed. Could not create \"" + mergeCmdFilePath;
			errorText += L"\" file.\n";
			errorText += exErrorStr;

			TRACE(L"%s\n", errorText);
			fpm.Close();
			return -1;
		}
		selectedMailsListFileOpen = TRUE;

		DWORD error;
		CString filePath;
		CString fileName;
		CStringA fileNameA;
		CString fPath;
		CString textFile;
		bool fileExists = false;
		CString newSuffix = L".pdf";
		CStringArray in_array;

		CString line;
		CStringA lineA;

		int textType = 1;     // FIXMEFIXME
		int indx = 0;
		for (int j = 0; j < cnt; j++)
		{
			i = (*selectedMailIndexList)[j];

			fileExists = false;
			int ret = 1;

			// MakeFileNameFromMailHeader must create the same PDF file name as PrintSingleMailtoPDF
			// FIXME: consider enhancing PrintSingleMailtoPDF to return name of created PDF file to avoid
			// potential inconsistency

			pFrame->m_NamePatternParams.m_bConvert2Ansi = TRUE;
			ret = MboxMail::MakeFileNameFromMailHeader(i, textType, textFile, targetPrintSubFolderName, fileExists, errorText);
			pFrame->m_NamePatternParams.m_bConvert2Ansi = FALSE;
			if (fileExists)
				int deb = 1;

			fPath = textFile;
			FileUtils::UpdateFileExtension(fPath, newSuffix);
			FileUtils::GetFolderPathAndFileName(fPath, filePath, fileName);

			in_array.Add(fileName);

			TextUtilsEx::WStr2Ascii(fileName, fileNameA, error);

			lineA.Format("%s\n", fileNameA);
			fptxt.Write(lineA, lineA.GetLength());   // FIXMEFIXME  SelectedMailsList.txt just info
		}

		if (selectedMailsListFileOpen)
			fptxt.Close();

		// write to MergeAllPDFs.cmd, i.e to fpm
		CStringA mergeAllCmdScriptA = "@echo on\n\n";
		fpm.Write(mergeAllCmdScriptA, mergeAllCmdScriptA.GetLength());

		CString mergedFilePrefix = L"m";
		CStringArray out_array;
		// convert to while ()
		errorText.Empty();
		int j;
		for (j = 0; j < 100; j++)
		{
			// MergePDfFileList creates batch cmd files to keep merging subsets of PDF files
			// Drawback is that cmd files don't seem to support UNICODE so PDF files must be encoded as Ansi/Ascii
			// which means that original files names are mapped to strange looking names sometimes
			// ShellExecuteEx could be used directly (was verified) to keep UNICODE names but stderr and stout can't redirected 
			// when running application (i.e. java and pdfbox jar file) directly by ShellExecuteEx
			//
			int ret = NListView::MergePDfFileList(fpm, in_array, out_array, mergedFilePrefix, targetPrintFolderPath, 
				pdfboxJarFileName, errorText);
			if (ret >= 0)
			{
				errorText.Empty();
				if (out_array.GetCount() > 1)
				{
					mergedFilePrefix.Append(L"m");
					in_array.Copy(out_array);
					out_array.SetSize(0);
				}
				else
					break;
			}
			else
			{
				fpm.Close();
				return -1;
			}
		}

		FolderContext *fc = &MboxMail::s_folderContext;

		CString mailArchiveFileName = MboxMail::s_path;

		if (mailArchiveFileName.IsEmpty())
		{
			mergedFileName = L"MergedMails.pdf";
		}
		else
		{
			CString driveName;
			CString directory;
			CString fileNameBase;
			CString fileNameExtention;
			FileUtils::SplitFilePath(mailArchiveFileName, driveName, directory, fileNameBase, fileNameExtention);

			mergedFileName = fileNameBase + L".pdf";
		}

		if (out_array.GetCount() == 0)
		{
			fpm.Close();
			return -1;
		}

		// Finish writing to MergeAllPDFs.cmd, i.e to fpm
		CString mergeAllCmdScript;
		mergeAllCmdScript.Format(L"\n\n@echo on\n\ndel /f /q \"%s\"\n\nren \"%s\" \"%s\"\n\npause\n\n", mergedFileName, out_array.GetAt(0), mergedFileName);

		BOOL retW2A = TextUtilsEx::WStr2Ansi(mergeAllCmdScript, mergeAllCmdScriptA, error);

		fpm.Write(mergeAllCmdScriptA, mergeAllCmdScriptA.GetLength());

		fpm.Close();  // i.e close MergeAllPDFs.cmd file

		CString mergedFilePath = targetPrintFolderPath + L"\\" + mergedFileName;
		mergedFileName = mergedFilePath;

		CString oldFilePath = targetPrintFolderPath + L"\\" + out_array.GetAt(0);

		DWORD nFlags = MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH | MOVEFILE_COPY_ALLOWED;
		BOOL retMove = MoveFileEx(oldFilePath, mergedFilePath, nFlags);
		if (retMove == FALSE)
		{
			CString errText = FileUtils::GetLastErrorAsString();
			TRACE(L"PrintMailSelectedToSeparatePDF_WorkerThread: MoveFileEx failed \"%s\"\n", errorText);

			CString txt = L"Failed to move \"" + oldFilePath;
			txt += L"\" file.\n\n";
			txt.Append(errorText);

			TRACE(L"%s\n", txt);

			mergedFileName = oldFilePath;

			return 1;
		}
	}  // end of if (mergePDFs)

	int deb = 1;
	return 1;
}

int NListView::MergePDfFileList(CFile &fpm, CStringArray &in_array, CStringArray &out_array, CString &mergedFilePrefix, CString &targetPrintFolderPath, 
	CString& pdfboxJarFileName, CString &errortext)
{
	CString mergeScript;
	CStringA mergeScriptA;
	CStringA echoCmd = "@echo off\n\n";
	CString errorText;

	int maxCmdLineLength = 8191 - 700;
	//maxCmdLineLength = 700;

	CFile fpMergeError;
	CFileException ExMergeError;
	CString mergeErrorFilePath = targetPrintFolderPath + L"\\MergePDFsError.log";
	if (!fpMergeError.Open(mergeErrorFilePath, CFile::modeReadWrite | CFile::modeCreate | CFile::modeNoTruncate | CFile::shareDenyNone,
		&ExMergeError))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExMergeError);

		errorText = L"Could not create \"" + mergeErrorFilePath;
		errorText += L"\" file.\n";
		errorText += exErrorStr;

		TRACE(L"%s\n", errorText);
		return -1;
	}

	fpMergeError.SeekToEnd(); // Append Mode

	CFile fp;
	CFileException ExError;
	CString mergeCmdFilePath = targetPrintFolderPath + L"\\MergePDFs.cmd";
	if (!fp.Open(mergeCmdFilePath, CFile::modeReadWrite | CFile::modeCreate | CFile::shareDenyNone,
		&ExError))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		errorText = L"Could not create \"" + mergeCmdFilePath;
		errorText += L"\" file.\n";
		errorText += exErrorStr;

		TRACE(L"%s\n", errorText);

		fpMergeError.Close();

		return -1;
	}

	CString mergedFileName;
	CString mergedFile;
	CString fileName;
	CString filePath;

	BOOL progressBar = TRUE;
	CString progressText = L"Merging PDFs";
	CString args;

	CString pdfboxtoolVMArgs;
	pdfboxtoolVMArgs.Append(L" -Dorg.apache.commons.logging.Log=org.apache.commons.logging.impl.Jdk14Logger");
	CString configFilePath = L"pdfbox-config.txt";
	pdfboxtoolVMArgs.Append(L" -Djava.util.logging.config.file=");
	pdfboxtoolVMArgs.Append(configFilePath);

	CString pdftoolCmd = L"java" + pdfboxtoolVMArgs + L" -jar " + pdfboxJarFileName + L" merge";

	mergeScript.Append(pdftoolCmd);  // FIXME review above lines

	DWORD error = 0;
	int cnt = 0;
	int mergedFileCnt = 0;
	int i;
	for (i = 0; i < in_array.GetCount(); i++)
	{
		if (mergeScript.GetLength() > maxCmdLineLength)
		{
			mergedFileName.Format(L"%s-%d.pdf", mergedFilePrefix, cnt++);

			mergedFile.Format(L" -o %s >> MergePDFs.log 2>&1\n\n", mergedFileName);
			mergeScript.Append(mergedFile);

			BOOL retW2A = TextUtilsEx::WStr2Ansi(mergeScript, mergeScriptA, error);

			fp.Write(echoCmd, echoCmd.GetLength());
			fp.Write(mergeScriptA, mergeScriptA.GetLength());
			fpm.Write(mergeScriptA, mergeScriptA.GetLength());

			CString mergeCmdFilePath = fp.GetFilePath();

			fp.Close();

			out_array.Add(mergedFileName);

			progressText.Format(L"Merging %d PDFs", mergedFileCnt);

			const int forced_timeout = -1;
			int retexec = CMainFrame::ExecCommand_WorkerThread(targetPrintFolderPath, mergeCmdFilePath, args, errorText, progressBar, progressText, forced_timeout);

			CFileException ExError;
			if (!fp.Open(mergeCmdFilePath, CFile::modeReadWrite | CFile::modeCreate | CFile::shareDenyNone,
				&ExError))
			{
				CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

				errorText = L"Could not create \"" + mergeCmdFilePath;
				errorText += L"\" file.\n";
				errorText += exErrorStr;

				TRACE(L"%s\n", errorText);

				CStringA errorTextA = errorText;

				fpMergeError.Write(errorTextA, errorTextA.GetLength());

				fpMergeError.Close();
				return -1;
			}

			mergeScript.Empty();
			mergeScript.Append(pdftoolCmd);
			mergedFileCnt = 0;
		}

		fileName = in_array.GetAt(i);
		filePath = targetPrintFolderPath + L"\\" + fileName;
		// Make sure Browser was able to create the file, oterwise merge will fail !!!!
		if (FileUtils::PathFileExist(filePath))
		{
			mergeScript.Append(L" -i \"");
			mergeScript.Append(fileName);
			mergeScript.Append(L"\"");

			mergedFileCnt++;
		}
		else
		{
			CString line;
			line.Format(L"File \"%s\" not found and it will not be merged\n", fileName);
			CStringA lineA = line;
			fpMergeError.Write(lineA, lineA.GetLength());
		}
	}

	if (mergedFileCnt)
	{
		mergedFileName.Format(L"%s-%d.pdf", mergedFilePrefix, cnt++);

		mergedFile.Format(L" -o %s >> MergePDFs.log 2>&1\n\n", mergedFileName);
		mergeScript.Append(mergedFile);

		BOOL retW2A = TextUtilsEx::WStr2Ansi(mergeScript, mergeScriptA, error);

		fp.Write(echoCmd, echoCmd.GetLength());
		fp.Write(mergeScriptA, mergeScriptA.GetLength());
		fpm.Write(mergeScriptA, mergeScriptA.GetLength());

		CString mergeCmdFilePath = fp.GetFilePath();

		fp.Close();

		out_array.Add(mergedFileName);

		progressText.Format(L"Merging %d PDFs", mergedFileCnt);

		const int forced_timeout = -1;
		int retexec = CMainFrame::ExecCommand_WorkerThread(targetPrintFolderPath, mergeCmdFilePath, args, errorText, progressBar, progressText, forced_timeout);
	}
	else
	{
		fp.Close();
	}

	fpMergeError.Close();
	return 1;
}

int NListView::PrintMailSelectedToSinglePDF_WorkerThread(MailIndexList *selectedMailsIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText)
{
	BOOL progressBar = TRUE;

	//int retdel = NListView::DeleteAllHtmAndPDFFiles(targetPrintFolderPath);

	if (selectedMailsIndexList->GetCount() <= 0)
		return 1;

	PrintMailSelectedToSingleHTML_WorkerThread(selectedMailsIndexList, targetPrintSubFolderName, targetPrintFolderPath, errorText);

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	int cnt = (int)selectedMailsIndexList->GetCount();

	CString progressText;
	if (progressBar && MboxMail::pCUPDUPData)
	{
		CString fmt = L"Printing %d mails to single PDF file ...";
		ResHelper::TranslateString(fmt);
		progressText.Format(fmt, (int)cnt);

		MboxMail::pCUPDUPData->SetProgress(progressText, 1);
	}

	double curstep = 1;

	if (progressBar && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

#if 0  // FIXME
	int textType = 1;
	CString targetPrintSubFolder;
	bool fileExists;
	CString htmFileName;
	int ret = MboxMail::MakeFileNameFromMailArchiveName(textType, htmFileName, targetPrintSubFolder, fileExists, errorText);
#endif

	//
	int textType = 1;  // 0=.txt 1=.htm 2=.pdf 3=.csv
	bool fileExists = false;
	CString htmFileName;
	CString targetPrintSubFolder;
	int ret;


	// FIXME determine the file name in the root caller to avoid inconsistency
	// It is ridicules now
	if (selectedMailsIndexList->GetCount() > 1)
		ret = MboxMail::MakeFileNameFromMailArchiveName(textType, htmFileName, targetPrintSubFolder, fileExists, errorText);
	else
		ret = MboxMail::MakeFileNameFromMailHeader(selectedMailsIndexList->GetAt(0), textType, htmFileName, targetPrintSubFolder, fileExists, errorText);

	//

	const int forced_timeout = -1;
	const int headless_timeout = -1;
	// Number of emails are concataneted, can't set reasonable timeout.
	// In a few case, User may have to decide how long to wait amd cancel print.
	ret = CMainFrame::ExecCommand_WorkerThread(htmFileName, errorText, TRUE, progressText, forced_timeout, headless_timeout);
	return 1;
}

void NListView::FindFirstAndLastMailOfConversation(int iItem, int &firstMail, int &lastMail)
{
	if (abs(MboxMail::b_mails_which_sorted) == 99)
		FindFirstAndLastMailOfMailThreadConversation(iItem, firstMail, lastMail);
	else if ((abs(MboxMail::b_mails_which_sorted) == 4) && (m_subjectSortType == 1))
		FindFirstAndLastMailOfSubjectConversation(iItem, firstMail, lastMail);
}

void NListView::FindFirstAndLastMailOfMailThreadConversation(int iItem, int &firstMail, int &lastMail)
{
	firstMail = 0;
	lastMail = 0;

	MboxMail *m = 0;
	m = MboxMail::s_mails[iItem];
	int groupId = m->m_groupId;
	firstMail = iItem;
	lastMail = iItem;

	int i;
	for (i = iItem; i >= 0; i--)
	{
		m = MboxMail::s_mails[i];
		if (m->m_groupId != groupId)
			break;
	}
	if (i < 0)
		firstMail = 0;
	else
		firstMail = i + 1;

	int mailCount = MboxMail::s_mails.GetCount();
	for (i = iItem; i < mailCount; i++)
	{
		m = MboxMail::s_mails[i];
		if (m->m_groupId != groupId)
			break;
	}
	if (i >= mailCount)
		lastMail = mailCount - 1;
	else
		lastMail = i - 1;
}


// TODO: can I reuse m->m_groupId and 
void NListView::FindFirstAndLastMailOfSubjectConversation(int iItem, int &firstMail, int &lastMail)
{
	firstMail = 0;
	lastMail = 0;

	MboxMail *m = 0;
	m = MboxMail::s_mails[iItem];
	int groupId = m->m_mail_index;
	firstMail = iItem;
	lastMail = iItem;

	int i;
	for (i = iItem; i >= 0; i--)
	{
		m = MboxMail::s_mails[i];
		if (m->m_mail_index != groupId)
			break;
	}
	if (i < 0)
		firstMail = 0;
	else
		firstMail = i + 1;

	int mailCount = MboxMail::s_mails.GetCount();
	for (i = iItem; i < mailCount; i++)
	{
		m = MboxMail::s_mails[i];
		if (m->m_mail_index != groupId)
			break;
	}
	if (i >= mailCount)
		lastMail = mailCount - 1;
	else
		lastMail = i - 1;
}

MailIndexList* NListView::PopulateSelectedMailsListFromRange(int firstIndex, int lastIndex)
{
	int selectedCnt = lastIndex - firstIndex + 1;
	m_selectedMailsList.SetSize(selectedCnt);

	int i = 0;
	int j;
	for (j = firstIndex; j <= lastIndex; j++)
	{
		if (i >= selectedCnt) // just a test 
			int deb = 1;

		m_selectedMailsList.SetAtGrow(i, j); // to be safe grow just in case
		i++;
	}
	return &m_selectedMailsList;
}

MailIndexList * NListView::PopulateSelectedMailsList()
{
	int selectedCnt = m_list.GetSelectedCount();
	m_selectedMailsList.SetSize(selectedCnt);

	POSITION p = m_list.GetFirstSelectedItemPosition();
	int nSelected;
	int i = 0;
	while (p)
	{
		if (i >= selectedCnt) // just a test 
			int deb = 1;
		nSelected = m_list.GetNextSelectedItem(p);
		m_selectedMailsList.SetAtGrow(i, nSelected); // to be safe grow just in case
		i++;
	}
	return &m_selectedMailsList;
}

#if 0
BOOL NListView::PreTranslateMessage(MSG* pMsg)
{
	// TODO: Add your specialized code here and/or call the base class
	//if (pMsg->message == WM_KEYDOWN)
	{
		//::SetFocus(GetSafeHwnd());
	}

	return CWnd::PreTranslateMessage(pMsg);
}
#endif

#if 1

void NListView::OnSetFocus(CWnd* pOldWnd)
{
	CWnd::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here

	//SetFocus();
}

#endif

void NListView::OnMouseHover(UINT nFlags, CPoint point)   // FIXME review why it was enabled
{
	// TODO: Add your message handler code here and/or call default

	CWnd::OnMouseHover(nFlags, point);
}


CWnd * NListView::SetSetFocus()
{
	SetFocus();
	m_list.SetFocus();

	return &m_list;
}

int NListView::DeleteAllHtmAndPDFFiles(CString &targetPath)
{
	BOOL progressBar = TRUE;
	CString filePath;
	BOOL delStatus;
	int fileCnt = 0;
	CString fn;

	if (progressBar && MboxMail::pCUPDUPData)
	{
		CString txt = L"Deleting HTM and PDF files in target print location ...";
		ResHelper::TranslateString(txt);
		MboxMail::pCUPDUPData->SetProgress(txt, 0);
	}

	CString fw = targetPath + L"\\*.htm";
	WIN32_FIND_DATA	wf;
	HANDLE f = FindFirstFile(fw, &wf);
	if (f != INVALID_HANDLE_VALUE) 
	{
		do 
		{
			if ((wf.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == FILE_ATTRIBUTE_DIRECTORY)
				continue;
			fileCnt++;
			fn = wf.cFileName;
			//TRACE(L"HTM file=%s\n", fn);
			filePath = targetPath + L"\\" + fn;
			delStatus = FileUtils::DelFile(filePath);
			if (delStatus == FALSE) {
				DWORD error = GetLastError();
			}

		} while (FindNextFile(f, &wf));
		FindClose(f);
	}

	fw = targetPath + L"\\*.pdf";
	f = FindFirstFile(fw, &wf);
	if (f != INVALID_HANDLE_VALUE)
	{
		do
		{
			if ((wf.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == FILE_ATTRIBUTE_DIRECTORY)
				continue;
			fileCnt++;
			fn = wf.cFileName;
			//TRACE(L"PDF file=%s\n", fn);
			filePath = targetPath + L"\\" + fn;
			delStatus = FileUtils::DelFile(filePath);
			if (delStatus == FALSE) {
				DWORD error = GetLastError();
			}

		} while (FindNextFile(f, &wf));
		FindClose(f);
	}
	return fileCnt;
}
/////////////////////////////////////////////////
//////////////     Print to HTML   //////////////
/////////////////////////////////////////////////

int NListView::PrintMailArchiveToSeparateHTML_Thread(CString &errorText)
{
	int firstMail = 0;
	int lastMail = MboxMail::s_mails.GetCount() - 1;

	CString targetPrintSubFolderName = L"HTML_GROUP";
	int ret = PrintMailRangeToSeparateHTML_Thread(firstMail, lastMail, targetPrintSubFolderName);

	return ret;
}

int NListView::PrintMailConversationToSeparateHTML_Thread(int iItem, CString &errorText)
{
	int firstMail = 0;
	int lastMail = 0;

	// Convert range to list to avoid enhancancing PrintMailRangeToSeparatePDF_Thread
	FindFirstAndLastMailOfConversation(iItem, firstMail, lastMail);

	MailIndexList* selectedMailsIndexList = PopulateSelectedMailsListFromRange(firstMail, lastMail);

	CString targetPrintFolderPath;
	CString targetPrintSubFolderName = L"HTML_GROUP";
	int ret = PrintMailSelectedToSeparateHTML_Thread(selectedMailsIndexList, targetPrintSubFolderName, targetPrintFolderPath);

	return ret;
}

int NListView::PrintMailConversationToSingleHTML_Thread(int iItem, CString &errorText)
{
	int firstMail = 0;
	int lastMail = 0;

	FindFirstAndLastMailOfConversation(iItem, firstMail, lastMail);

	CString targetPrintSubFolderName;
	int ret = PrintMailRangeToSingleHTML_Thread(firstMail, lastMail, targetPrintSubFolderName);

	return ret;
}

int NListView::PrintMailRangeToSeparateHTML_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName)
{
	MailIndexList* selectedMailsIndexList = PopulateSelectedMailsListFromRange(firstMail, lastMail);

	CString targetPrintFolderPath;
	//CString targetPrintSubFolderName = L"HTML_GROUP";
	int ret = PrintMailSelectedToSeparateHTML_Thread(selectedMailsIndexList, targetPrintSubFolderName, targetPrintFolderPath);

	return ret;
}

int NListView::PrintMailRangeToSingleHTML_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName)
{
		MailIndexList* selectedMailsIndexList = PopulateSelectedMailsListFromRange(firstMail, lastMail);

		CString targetPrintFolderPath;
		int ret = PrintMailSelectedToSingleHTML_Thread(selectedMailsIndexList, targetPrintSubFolderName, targetPrintFolderPath);

		return ret;
}

#include <afxdlgs.h>

int NListView::ExportMails_CopyExportMails2PDF()
{
	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	CString processExePath = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"processPath");

	if (processExePath.IsEmpty())
	{
		;
	}

	CString processFolderPath;
	FileUtils::CPathGetPath(processExePath, processFolderPath);
	
	CString rootExportSubFolder = L"ExportCache";
	CString targetExportSubFolder;
	CString exportCachePath;
	CString errorText;

	BOOL retval = MboxMail::CreateCachePath(rootExportSubFolder, targetExportSubFolder, exportCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = GetSafeHwnd(); // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	BOOL bFailIfExists = TRUE;
	//
	CString exportMails2PdfFilePath = processFolderPath + L"\\scripts\\exportHTML2PDF.cmd";
	CString targetExportMails2PdfFilePath = exportCachePath + L"\\exportHTML2PDF.cmd";

	if (!::CopyFileW(exportMails2PdfFilePath, targetExportMails2PdfFilePath, bFailIfExists))
	{
		CString errText = FileUtils::GetLastErrorAsString();
		TRACE(L"CopyFile: %s\n%s", exportMails2PdfFilePath, errText);
	}
	return 1;
}
//

BOOL  NListView::CanGoAheadWithExport()
{
	if (!NListView::m_exportMailsMode)
		return TRUE;

	BOOL can = FALSE;

	CString rootPrintSubFolder = L"ExportCache";
	CString errorText;

	CString exportCachePath;
	CString exportSubFolder;
	if (NListView::m_exportMailsMode)
	{
		BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, exportSubFolder, exportCachePath, errorText);
		if (retval == FALSE) {
			HWND h = NULL; // we don't have any window yet  
			int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}

		BOOL firstTry = TRUE;
		for (;;)
		{
			if (FileUtils::PathDirExists(exportCachePath))
			{
				CString errorText;
				CString fmt;
				if (firstTry)
				{
					fmt = L"Directory\n\n\"%s\"\n\nexists already. Only single export per mail archive file is supported.\n\n"
						"Select Yes to override the existing export directory\n"
						"Select No to copy directory to another location\n"
						"Select Cancel\n";
				}
				else
				{
					fmt = L"Directory\n\n\"%s\"\n\nstill exists, was not able to delete the directory.\n\n"
						"Select Yes to try override again the existing export directory\n"
						"Select No to copy directory to another location\n"
						"Select Cancel\n";
				}

				ResHelper::TranslateString(fmt);
				errorText.Format(fmt, exportCachePath);

				HWND h = GetSafeHwnd();
				int answer = ::MessageBox(h, errorText, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNOCANCEL);
				if (answer == IDYES)
				{
					bool recursive = true;
					bool removeFolders = true;
					CString errorText;
					BOOL delstatus = FileUtils::RemoveDir(exportCachePath, recursive, removeFolders, &errorText);
					if (!errorText.IsEmpty())
					{
						CString txt1;
						CString fmt = L"Failed to remove and recreate folder for Export:\n\n\"%s\"\n\n%s";
						ResHelper::TranslateString(fmt);
						txt1.Format(fmt, exportCachePath, errorText);

						CString txt2 = L"\n\nFolder or subfolders or files in the folder are likely used by other programs.\n\n"
							L"\nNOTE: This issue will happen if the folder or folder's items are open in Command Prompt.\n\n"
							L"\nNOTE: This issue will NOT happen if the folder or folder's items are open in File Explorer or Windows Terminal or Windows PowerShell.\n\n"
							L"\nNOTE: Other than above programs may cause this issue as well. You can use Handle.exe or Process Explorer.exe programs to discover programs creating the issue. "
							L"Handle.exe and Process Explorer.exe programs are part of Microsoft Sysinternal package. They can also be downloaded as individual programs\n\n"
							;
						ResHelper::TranslateString(txt2);
						CString txt = txt1 + txt2;


						HWND h = GetSafeHwnd();
						int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
						int deb = 1;
					}
					int deb = 1;
				}
				else if (answer == IDNO)
				{
					rootPrintSubFolder.Empty();
					CString exportCacheParentPath;
					BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, exportSubFolder, exportCacheParentPath, errorText);
					if (retval == FALSE) {
						HWND h = NULL; // we don't have any window yet  
						int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
						return -1;
					}
#if 1
					// 
					DWORD dwFlags = OFN_EXPLORER;
					CFileDialog dlgFile(TRUE, NULL, exportCachePath, dwFlags, NULL, NULL, 0, TRUE);

					//dlgFile.SetControlText(IDCANCEL, L"Quit"); // may need to add Hook processing; InitDialog doesn't seem to be supported

					dlgFile.DoModal();  // Blocking call
#else

					if (FileUtils::BrowseToFile(exportCachePath) == FALSE)  // Not blocking call
					{
						DWORD dwFlags = OFN_EXPLORER;
						CFileDialog dlgFile(TRUE, NULL, exportCachePath, dwFlags, NULL, NULL, 0, TRUE);
						dlgFile.DoModal();
					}
#endif
					int deb = 1;
				}
				else
				{
					return FALSE;
				}
			}
			else
				return TRUE;
		}
	}
	return can;
}

int NListView::PrintMailSelectedToSeparateHTML_Thread(MailIndexList* selectedMailsIndexList, CString &targetPrintSubFolderName, 
	CString &targetPrintFolderPath)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS args;

	if (NListView::m_exportMailsMode && !CanGoAheadWithExport())
		return -1;

	CString rootPrintSubFolder = L"PrintCache";
	CString targetPrintSubFolder = targetPrintSubFolderName;
	if (NListView::m_exportMailsMode)
	{
		rootPrintSubFolder = L"ExportCache";
		_ASSERTE(targetPrintSubFolderName.Compare(L"Mails") == 0);
	}

	CString printCachePath;
	CString errorText;

	if (selectedMailsIndexList == 0)
	{
		selectedMailsIndexList = PopulateSelectedMailsList();
	}

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, printCachePath, errorText);
	if (retval == FALSE) {
		HWND h = GetSafeHwnd(); // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	/*IN*/
	args.lview = this;
	args.targetPrintFolderPath = printCachePath;
	args.targetPrintSubFolderName = targetPrintSubFolder;
	args.firstMail = -1;
	args.lastMail = -1;
	args.selectedMailIndexList = selectedMailsIndexList;
	args.nItem = -1;
	args.separatePDFs = TRUE;
	args.mergePDFs = FALSE; // not applicable in this case anyway

	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcPrintMailGroupToSeparateHTML, (LPVOID)(PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS*)&args);
	Dlg.SetDialogTemplate(AfxGetApp()->m_hInstance, MAKEINTRESOURCE(IDD_PROGRESS_DLG), IDC_STATIC, IDC_PROGRESS_BAR, IDCANCEL);

	INT_PTR nResult = Dlg.DoModal();

	if (!nResult) { // should never be true ?
		MboxMail::assert_unexpected();
		return -1;
	}

	int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
	int retResult = LOWORD(nResult);

	if (retResult != IDOK)
	{  // IDOK==1, IDCANCEL==2
		// We should be here when user selects Cancel button
		//_ASSERTE(cancelledbyUser == TRUE);

		DWORD terminationDelay = Dlg.GetTerminationDelay();
		int loopCnt = (terminationDelay+100)/25;

		ULONGLONG tc_start = GetTickCount64();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		ULONGLONG tc_end = GetTickCount64();
		DWORD delta = (DWORD)(tc_end - tc_start);
		TRACE(L"(PrintMailSelectedToSeparateHTML_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret > 0)
	{
		if (args.errorText.IsEmpty())
		{
			if (!NListView::m_exportMailsMode)
			{
				//CString txt = L"Created separate HTML files in \n\n" + printCachePath + L" \n\ndirectory.";
				CString txt = L"Created separate HTML files in the following directory";
				ResHelper::TranslateString(txt);

				OpenContainingFolderDlg dlg(txt, printCachePath, TRUE);
				INT_PTR nResponse = dlg.DoModal();
				if (nResponse == IDOK)
				{
					HWND h = GetSafeHwnd();
					HINSTANCE result = ShellExecute(h, L"open", printCachePath, NULL, NULL, SW_SHOWNORMAL);
					CMainFrame::CheckShellExecuteResult(result, h);
					int deb = 1;
				}
				else if (nResponse == IDCANCEL)
				{
					int deb = 1;
				}
			}
			else
			{
				CString exportCachePath;
				CString exportSubFolder;
				if (NListView::m_exportMailsMode)
				{
					// exportCachePath == xxx/ExportCache
					BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, exportSubFolder, exportCachePath, errorText);
					if (retval == FALSE) {
						HWND h = GetSafeHwnd(); // we don't have any window yet  
						int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
						return -1;
					}
				}

				BOOL retindex = CreateIndexFileForExportedMails_Thread(selectedMailsIndexList, targetPrintSubFolderName, targetPrintFolderPath);

				CString fileName = exportCachePath + L"\\index.html";

				int exret = ExportMails_CopyExportMails2PDF();

				CString txt = L"Created separate mails files as HTML files and index.html file in the following directory";
				ResHelper::TranslateString(txt);

				OpenContainingFolderDlg dlg(txt, exportCachePath, FALSE);
				INT_PTR nResponse = dlg.DoModal();
				////////////
				if (nResponse == IDOK)
				{
					if (FileUtils::BrowseToFile(fileName) == FALSE) {
						HWND h = GetSafeHwnd();
						HINSTANCE result = ShellExecute(h, L"open", exportCachePath, NULL, NULL, SW_SHOWNORMAL);
						CMainFrame::CheckShellExecuteResult(result, h);
					}
					int deb = 1;
				}
				else if (nResponse == IDYES)
				{
					HWND h = GetSafeHwnd();

					HINSTANCE result = ShellExecute(h, L"open", fileName, NULL, NULL, SW_SHOWNORMAL);
					CMainFrame::CheckShellExecuteResult(result, h, &fileName);
					int deb = 1;
				}
				else if (nResponse == IDCANCEL)
				{
					int deb = 1;
				}

			}
		}
		else
		{
			MboxMail::assert_unexpected();
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}
	return 1;
}

int NListView::PrintMailSelectedToSingleHTML_Thread(MailIndexList* selectedMailsIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS args;

	if (NListView::m_exportMailsMode && !CanGoAheadWithExport())
		return -1;

	CString rootPrintSubFolder = L"PrintCache";
	CString targetPrintSubFolder = targetPrintSubFolderName;
	if (NListView::m_exportMailsMode)
	{
		rootPrintSubFolder = L"ExportCache";
		_ASSERTE(targetPrintSubFolder.IsEmpty());
	}
	CString printCachePath;
	CString errorText;

	if (selectedMailsIndexList == 0)
	{
		selectedMailsIndexList = PopulateSelectedMailsList();
	}
	if (selectedMailsIndexList->GetCount() <= 0)
		return 1;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, printCachePath, errorText);
	if (retval == FALSE) {
		HWND h = GetSafeHwnd(); // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	/*IN*/
	args.lview = this;
	args.targetPrintFolderPath = printCachePath;
	args.targetPrintSubFolderName = targetPrintSubFolder;
	args.firstMail = -1;
	args.lastMail = -1;
	args.selectedMailIndexList = selectedMailsIndexList;
	args.nItem = -1;
	args.separatePDFs = FALSE;

	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcPrintMailGroupToSingleHTML, (LPVOID)(PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS*)&args);
	Dlg.SetDialogTemplate(AfxGetApp()->m_hInstance, MAKEINTRESOURCE(IDD_PROGRESS_DLG), IDC_STATIC, IDC_PROGRESS_BAR, IDCANCEL);

	INT_PTR nResult = Dlg.DoModal();

	if (!nResult) { // should never be true ?
		MboxMail::assert_unexpected();
		return -1;
	}

	int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
	int retResult = LOWORD(nResult);

	if (retResult != IDOK)
	{  // IDOK==1, IDCANCEL==2
		// We should be here when user selects Cancel button
		//_ASSERTE(cancelledbyUser == TRUE);

		DWORD terminationDelay = Dlg.GetTerminationDelay();
		int loopCnt = (terminationDelay+100)/25;

		ULONGLONG tc_start = GetTickCount64();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		ULONGLONG tc_end = GetTickCount64();
		DWORD delta = (DWORD)(tc_end - tc_start);
		TRACE(L"(PrintMailSelectedToSingleHTML_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret > 0)
	{
		if (args.errorText.IsEmpty())
		{
			int fileType = 1;  // 0=.txt 1=.htm 2=.pdf 3=.csv
			bool fileExists;
			CString fileName;

			// Prone to inconsistency, fileName is also determine in ALongRightProcessProcPrintMailGroupToSingleHTML above
			if (selectedMailsIndexList->GetCount() > 1)
				int ret = MboxMail::MakeFileNameFromMailArchiveName(fileType, fileName, targetPrintSubFolderName, fileExists, errorText);
			else
				int ret = MboxMail::MakeFileNameFromMailHeader(selectedMailsIndexList->GetAt(0), fileType, fileName, targetPrintSubFolderName, fileExists, errorText);

			if (NListView::m_exportMailsMode)
				int exret = ExportMails_CopyExportMails2PDF();

			//CString txt = L"Created HTML file \n\n" + fileName;
			CString txt = L"Created HTML file";
			ResHelper::TranslateString(txt);

			OpenContainingFolderDlg dlg(txt, fileName, FALSE);
			INT_PTR nResponse = dlg.DoModal();
			////////////
			if (nResponse == IDOK)
			{
				if (FileUtils::BrowseToFile(fileName) == FALSE) {
					HWND h = GetSafeHwnd();
					HINSTANCE result = ShellExecute(h, L"open", printCachePath, NULL, NULL, SW_SHOWNORMAL);
					CMainFrame::CheckShellExecuteResult(result, h);
				}
				int deb = 1;
			}
			else if (nResponse == IDYES)
			{
				HWND h = GetSafeHwnd();
				HINSTANCE result = ShellExecute(h, L"open", fileName, NULL, NULL, SW_SHOWNORMAL);
				CMainFrame::CheckShellExecuteResult(result, h, &fileName);
				int deb = 1;
			}
			else if (nResponse == IDCANCEL)
			{
				int deb = 1;
			}

		}
		else
		{
			MboxMail::assert_unexpected();
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}
	return 1;
}

int NListView::PrintMailSelectedToSeparateHTML_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText)
{
	BOOL progressBar = TRUE;

	int retdel = NListView::DeleteAllHtmAndPDFFiles(targetPrintFolderPath);

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (selectedMailIndexList->GetCount() <= 0)
		return 1;

	if (progressBar && MboxMail::pCUPDUPData)
	{
		CString txt = L"Printing mails to HTML files ...";
		ResHelper::TranslateString(txt);
		MboxMail::pCUPDUPData->SetProgress(txt, 1);
	}

	UINT curstep = 1;
	UINT nFileNum;
	CString fileNum;

	if (progressBar && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	int i;
	int retval = 1;
	int cnt = (int)selectedMailIndexList->GetCount();

	ULONGLONG workRangeFirstPos = 0;
	ULONGLONG workRangeLastPos = cnt - 1;
	ProgressTimer progressTimer(workRangeFirstPos, workRangeLastPos);

	for (int j = 0; j < cnt; j++)
	{
		i = (*selectedMailIndexList)[j];
		if (pFrame) {
			retval = pFrame->PrintSingleMailtoHTML(i, targetPrintSubFolderName, errorText);
			if (retval < 0)
				int deb = 1;
		}

		if (progressBar && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				int deb = 1;
				break;
			}

			UINT_PTR dwProgressbarPos = 0;
			ULONGLONG workRangePos = j;
			BOOL needToUpdateStatusBar = progressTimer.UpdateWorkPos(workRangePos, dwProgressbarPos);
			if (needToUpdateStatusBar)
			{
				nFileNum = j + 1;

				CString fmt = L"Printing mails to HTML files ... %d of %d";
				ResHelper::TranslateString(fmt);
				fileNum.Format(fmt, nFileNum, cnt);

				if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(dwProgressbarPos));

				int debug = 1;
			}
		}
	}
	
	if (progressBar && MboxMail::pCUPDUPData)
	{
		UINT newstep = 100;
		nFileNum = cnt;

		CString fmt = L"Printing mails to HTML files ... %d of %d";
		ResHelper::TranslateString(fmt);
		fileNum.Format(fmt, nFileNum, cnt);

		if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(newstep));
	}

	int deb = 1;
	return 1;
}

int NListView::PrintMailSelectedToSingleHTML_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText)
{
	//int retdel = NListView::DeleteAllHtmAndPDFFiles(targetPrintFolderPath);

	if (selectedMailIndexList->GetCount() <= 0)
		return 1;

	TEXTFILE_CONFIG textConfig;
	CString textFileName;
	int textType = 1;

	textConfig.m_dateFormat = m_format;
	textConfig.m_bGMTTime = m_gmtTime;
	textConfig.m_nCodePageId = CP_UTF8;

	int ret = MboxMail::PrintMailSelectedToSingleTextFile_WorkerThread(textConfig, textFileName, selectedMailIndexList, textType, errorText);

	return 1;
}

int NListView::PrintMailRangeToSingleCSV_Thread(int iItem)
{
	if (!((abs(MboxMail::b_mails_which_sorted) == 99) || ((abs(MboxMail::b_mails_which_sorted) == 4) && (m_subjectSortType == 1))))
	{
		//if (abs(MboxMail::b_mails_which_sorted) != 99)
		{
			CString txt = L"Please sort all mails by conversation threads or subject threads first.\n";
			ResHelper::TranslateString(txt);

			//txt += "Select \"View\"->\"Sort By\" ->\"Conversation\" or left click on the first column.";
			HWND h = GetSafeHwnd(); // we don't have any window yet
			int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
			//return -1;
		}
		return -1;
	}

	int firstMail = 0;
	int lastMail = 0;
	FindFirstAndLastMailOfConversation(iItem, firstMail, lastMail);

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame) {
		BOOL selectedMails = FALSE;
		pFrame->PrintMailsToCSV(firstMail, lastMail, selectedMails);
	}

	return 1;
}

// ZMINDEX
int NListView::CreateIndexFileForExportedMails_Thread(MailIndexList* selectedMailsIndexList, CString& targetPrintSubFolderName,
	CString& targetPrintFolderPath)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS args;

	if (!NListView::m_exportMailsMode)
		return -1;

	CString rootPrintSubFolder = L"ExportCache";
	
	//CString targetPrintSubFolder = targetPrintSubFolderName;
	CString targetPrintSubFolder = targetPrintSubFolderName;
	CString exportCachePath;
	CString exportHtmlCachePath;
	CString errorText;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (selectedMailsIndexList == 0)
	{
		selectedMailsIndexList = PopulateSelectedMailsList();
	}

	if (selectedMailsIndexList->GetCount() <= 0)
		return 1;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, exportHtmlCachePath, errorText);
	if (retval == FALSE) {
		HWND h = GetSafeHwnd(); // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	targetPrintSubFolder.Empty();
	retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, exportCachePath, errorText);
	if (retval == FALSE) {
		HWND h = GetSafeHwnd(); // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	TEXTFILE_CONFIG textConfig;
	textConfig.m_dateFormat = m_format;
	textConfig.m_bGMTTime = m_gmtTime;
	textConfig.m_nCodePageId = CP_UTF8;

	CString mailArchiveFilePath = MboxMail::s_path;

	CString mailArchiveName;

	if (mailArchiveFilePath.IsEmpty()) {
		CString errorText = L"Please open mail archive file first.";
		return -1;
	}

	FileUtils::GetFileName(mailArchiveFilePath, mailArchiveName);

	DWORD error;
	CStringA mailArchiveNameA;
	BOOL w2u = TextUtilsEx::WStr2UTF8(&mailArchiveName, &mailArchiveNameA, error);
	
	// Vertical used in Summary pane
	CStringA paperClipData = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAYAAAAKCAYAAACXD"
		"i8zAAAAP0lEQVQIW2P8DwSMjIwMMADkgpkgEZAcXAKkCMTHKgHXAWLAjYAaC9ZBugTYM"
		"qgRMDaG5WCLgYpw24HuQZhDAJx+M/hD/nwvAAAAAElFTkSuQmCC";
	
	// Vertical PaperClip
#if 1
	paperClipData = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAABU0lEQVRIS+3VsUtCURTH"
		"cYWgyakhiXIRanAIIodAiHAIJFoKrD+gmqVcTNDAGqJotv6AimgqGoTAIGgwooYGi5aIoMGlTRD0e+JZj+ez89DanvDhPS/3nN99V"
		"716Pf/88jroP8ScbUwac6+4JvHmoNajBQzS5AE9ODYaxrnWMOokRAs4pEkMY3gxAoJc73CBRe0ptIB3Gpxj2dJon/czGOg2oE6DDWQ"
		"tjeR9BtoC1QnmAPO9OeDXMG0FboD7GXztgPst+v4lV9mOHaxjE2vo/cstKtBsHKeYQwnT3QbIqneRQgAHmMANlvCKLawaT2M5svTD"
		"6pYKOVFnWyp/Bs649SNsN0c7i/YoWoH8uTzbNBhm7B55JDoJ6KPoER+Yt4RI8xP0I4RKJwFSM4Uj+HCJJ0jzKD6xgKJdcxnTtqhZJ"
		"0+SRgQjKOMauXYrbxY6DWi3QHW8AZAubhlQEVdiAAAAAElFTkSuQmCC";
#else
	// Horizontal PaperClip
	paperClipData = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAABHklEQVRIS+2Uuw4BQRSG"
		"iURHqaBUu0QUOiqdUsQD8ASoeQyXQukVJCToEApUHoCERKFB5zvJjiwrsRu70Zjky87Mnpx/zz9nx+1yeLgdzu/6C3x0+CcWefisGCQ"
		"h+PETnwNuLDcwgIu8eq0gwl4XEhYTv4Zv2cjDSi8gyedwgBqMYW9RyEt8Gjog1cSUgNgyg4Bmz0lL7OdZgJAJoRExgows9CGnBMSSBR"
		"Shp0u+ZB42kVxCGlDXYn08z1BRAmUWTZBDVbaUmLcgBVOTIirMUIFdAnIGGWjDFeJ2WqQv0tBFdhyyCEjnrGEIhv9AtemRl1WYwM6i9"
		"4ZwO380fRc9hN7dRWJXFOSqMNP/KtmIifA0fnLZfWv7vwJrDjp+yHf1eTY2To4IIAAAAABJRU5ErkJggg==";
#endif

	CStringA imgstyle = "img.PaperClip {\n";
	imgstyle.Append("content: url(");
	imgstyle.Append(paperClipData);
	imgstyle.Append(");\nhight:16px;width:16px;\n}");

	// Create index file
	CStringA index = " \
<!DOCTYPE html>\n\
<html>\n\
<head>\n\
<style>\n\
table {\n\
	font-family: arial, sans-serif;\n\
	border-collapse: collapse;\n\
	width: 99%; \n\
}\n\
td, th {\n\
	border: 1px solid #dddddd;\n\
	text-align: left;\n\
	padding: 8px;\n\
}\n\
tr:nth-child(even) {\n\
	background-color: #dddddd;\n\
}\n\
td:nth-child(1) { text-align: center; }\n\
td:nth-child(2) { text-align: center; white-space:nowrap;}\n\
";
	index.Append(imgstyle);
	index.Append("\
\n</style>\n\
<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n\
<meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\">\n\
");

		index.Append("<title>");
		index.Append(mailArchiveNameA);

		index.Append("</title>\n</head>\n<body>\n");

		index.Append("<h2>");
		index.Append(mailArchiveNameA);
		index.Append("</h2>\n");

		index.Append("\
<table width=\"99%\" border=\"1\">\n\
<tr>\n\
	<th style=\"width:2%;text-align: center;\"><img class=\"PaperClip\"></th>\n\
	<th style=\"width:7%\">Date</th>\n\
	<th style=\"width:40%\">Subject</th>\n\
	<th style=\"width:25%\">From</th>\n\
	<th style=\"width:25%\">To</th>\n\
</tr>\n\
");

		int i;
		int cnt = (int)selectedMailsIndexList->GetCount();

		char datebuff[32];
		CStringA format = textConfig.m_dateFormat;

		CString textFile;
		CString fileName;
		CStringA fileNameA;
		CStringA subj;
		CStringA from;
		CStringA to;
		CStringA lDateTime;

		for (int j = 0; j < cnt; j++)
		{
			i = (*selectedMailsIndexList)[j];

			if ((i < 0) || (i >= MboxMail::s_mails.GetCount()))
				return -1;

			MboxMail* m = MboxMail::s_mails[i];

			BOOL hasAttachents = FALSE;
			if (m->m_hasAttachments || HasAnyAttachment(m))
				hasAttachents = TRUE;

			if (pFrame)
			{
				index.Append("<tr>\n");

				index.Append("<td>");
				//index.Append("&#x1F4CE;");  // but need to reverse backround color and text color somehow
				if (hasAttachents)
					index.Append("<img class=\"PaperClip\">");
				else
					index.Append("");
				index.Append("</td>\n");

				datebuff[0] = 0;
				if (m->m_timeDate >= 0)
				{
					MyCTime tt(m->m_timeDate);
					if (!textConfig.m_bGMTTime)
					{
						lDateTime = tt.FormatLocalTmA(format);
						strcpy(datebuff, (LPCSTR)lDateTime);
					}
					else
					{
						lDateTime = tt.FormatGmtTmA(format);
						strcpy(datebuff, (LPCSTR)lDateTime);
					}
				}

				index.Append("<td>");
				index.Append(datebuff);
				index.Append("</td>\n");

				index.Append("<td>");
				index.Append("<a href=\"Mails\\");

				int textType = 1;
				bool fileExists = false;
				errorText.Empty();
				textFile.Empty();
				BOOL nmret = MboxMail::MakeFileNameFromMailHeader(i, textType, textFile, targetPrintSubFolderName, fileExists, errorText);
				if (fileExists)
					int deb = 1;

				fileName.Empty();
				FileUtils::GetFileName(textFile, fileName);
				fileNameA.Empty();
				BOOL w2u = TextUtilsEx::WStr2UTF8(&fileName, &fileNameA, error);

				index.Append(fileNameA);

				CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
				if (pFrame && (pFrame->m_HdrFldConfig.m_bHdrAttachmentLinkOpenMode == 1))
					index.Append("\"target=\"_blank\">");
				else
					index.Append("\">");

				subj.Empty();
				MboxMail::encodeTextAsHtmlLinkLabel(m->m_subj, subj);
				if (subj.IsEmpty())
					subj.Append("&nbsp;&nbsp;&nbsp;&nbsp;");
				index.Append(subj);
				index.Append("</a>");
				index.Append("</td>\n");

				index.Append("<td>");
				from.Empty();
				MboxMail::encodeTextAsHtmlLinkLabel(m->m_from, from);
				index.Append(from);
				index.Append("</td>\n");

				index.Append("<td>");
				to.Empty();
				MboxMail::encodeTextAsHtmlLinkLabel(m->m_to, to);
				index.Append(to);
				index.Append("</td>\n");

				index.Append("</tr>\n");
			}
		}

		index.Append("</table></body></html>\n");
		CString indexFile = exportCachePath + L"\\index.html";
		const unsigned char* data = (unsigned char*)(LPCSTR)index;
		int dataLength = index.GetLength();
		BOOL wret = FileUtils::Write2File(indexFile, data, dataLength);

	return 1;
}

/////////////////////////////////////////////////
//////////////     Print to HTML END //////////////
/////////////////////////////////////////////////

// End of  ADDED in 1.0.3.1
void NListView::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: Add your message handler code here and/or call default
	MboxMail::ignoreException = FALSE;
	MessageBeep(MB_OK);
	MessageBeep(MB_OK);
	KillTimer(nIDEvent);
	//Sleep(10000);
	//CWnd::OnTimer(nIDEvent);
}

// TODO: Please define the naming standard to end this name nightmare  !!!!!!
int NListView::ReloadMboxListFile_v2(CString *mbxListFile)
{
	int ret = 1;  //OK

	MailArray *mailsArray = &MboxMail::s_mails;

	if (!(MboxMail::IsUserMailsSelected() || MboxMail::IsAllMailsSelected()))
	{
		// should never be here
		return -1;
	}

	BOOL isDirty = MboxMail::m_editMails.m_bIsDirty;
	if (MboxMail::IsUserMailsSelected())
	{
		if (isDirty)
		{
			CString txt = L"User Selected Mails List is not empty. Overwrite?";
			ResHelper::TranslateString(txt);
			int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
			if (answer == IDNO)
				return -1;
			else if (answer == IDYES)
				MboxMail::m_editMails.m_bIsDirty = FALSE;
		}
	}

	CString mboxFileSuffix = L"_USER.mbox";

	CString mailFile = MboxMail::s_path;

	if (MboxMail::s_path.IsEmpty()) {
		CString txt = L"Please open mail file first.";
		ResHelper::TranslateString(txt);
		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	CString path = MboxMail::GetLastPath();
	CString datapath = MboxMail::GetLastDataPath();
	if (path.IsEmpty())
		return -1;  // Hopefully s_path will fail first

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(mailFile, driveName, directory, fileNameBase, fileNameExtention);

	CString errorText;
	CString archiveCachePath;
	CString rootPrintSubFolder = L"ListCache";
	CString targetPrintSubFolder;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, archiveCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CString mboxFile = fileNameBase + fileNameExtention;  // main mbox file name
	CString mboxFilePath = archiveCachePath + L"\\" + fileNameBase + mboxFileSuffix;  // derived _USER.mbox mbox file name

	CString mboxFileListSuffix = L".mboxlist";
	CString mboxListFile;
	if (mbxListFile)
		mboxListFile.Append(*mbxListFile);
	else
		mboxListFile = mboxFilePath + mboxFileListSuffix;

	if (!FileUtils::PathFileExist(mailFile))
	{
#if 0
		CString txt = L"Mail Archive File \"" + mailFile;
		txt += L"\" doesn't exist.\nCan't reload.";
#endif

		CString txt;
		CString fmt = L"Mail Archive File \"%s\" doesn't exist.\nCan't reload.";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, mailFile);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	if (!FileUtils::PathFileExist(mboxListFile))
	{
#if 0
		CString txt = L"Mail List File \"" + mboxListFile;
		txt += L"\" doesn't exist.\nCan't reload.";
#endif

		CString txt;
		CString fmt = L"Mail List File \"%s\" doesn't exist.\nCan't reload.";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, mboxListFile);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	SerializerHelper sz(mboxListFile);
	if (!sz.open(FALSE))
	{
#if 0
		CString txt = L"Could not open \"" + mboxListFile;
		txt += L"\" file.\nMake sure file is not open on other applications.";
#endif
		CString txt;
		CString fmt = L"Could not open \"%s\" file.\nMake sure file is not open on other applications.";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, mboxListFile);

		int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	int version;
	_int64 mailFileSize;
	_int64 mboxFileSize;
	int mailListCnt;
#if 0
	CString txt = L"Mail list file\n\"" + mboxListFile;
	txt += L"\"\nhas incompatible version or file is corrupted.\nCan't reload.\nPlease remove the file.";
#endif
	CString txt;
	CString fmt = L"Mail list file\n\"%s\"\nhas incompatible version or file is corrupted.\nCan't reload.\nPlease remove the file.";
	ResHelper::TranslateString(fmt);
	txt.Format(fmt, mboxListFile);

	__int64 pos = 0;
	retval = sz.GetReadPointer(pos);

	if (!sz.readInt(&version)) {
		sz.close();
		return -1;
	}

	pos = 0;
	retval = sz.GetReadPointer(pos);

	if (version != MAIL_LIST_VERSION)
	{
		sz.close();
#if 0
		CString text = L"Mail list file\n\"" + mboxListFile;
		CString strVersion;
		strVersion.Format(L"%d", (version - MAIL_LIST_VERSION_BASE));
		text += L"\".\nhas incompatible version\"" + strVersion + L"\". Expected version \"";
		strVersion.Format(L"%d", (MAIL_LIST_VERSION - MAIL_LIST_VERSION_BASE));
		text += strVersion + L"\".\nCan't reload.\nPlease remove the file.";
#endif

		CString txt;
		CString strVersion;
		strVersion.Format(L"%d", (version - MAIL_LIST_VERSION_BASE));
		CString strCachedVersion;
		strCachedVersion.Format(L"%d", (MAIL_LIST_VERSION - MAIL_LIST_VERSION_BASE));

		CString fmt = L"Mail list file\n\"%s\".\nhas incompatible version\"%s\". Expected version \"%s\".\nCan't reload.\nPlease remove the file.";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, mboxListFile, strVersion, strCachedVersion);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	// from SaveAsMboxListFile_v2()
	//
	// sz.writeInt64(MboxMail::s_fSize);
	// _int64 mailFileSize = -1;  
	// sz.writeInt64(mailFileSize);
	// sz.writeInt(mailsArray->GetSize());

	if (!sz.readInt64(&mailFileSize)) { // main mbox mail
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	if (!sz.readInt64(&mboxFileSize)) {  // derived mbox mail file
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	if (!sz.readInt(&mailListCnt)) {
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	_int64 nMailFileSize = FileUtils::FileSize(mailFile);  // MboxMail::s_path; main mbox file
	_int64 nMboxFileSize = -1;

	if (FileUtils::PathFileExist(mboxFilePath))
		nMboxFileSize = FileUtils::FileSize(mboxFilePath);  // derived _USER.mbox mbox file name

	// force nMboxFileSize = nMboxFileSize = -1;
	nMboxFileSize = -1;
	nMboxFileSize = -1;

	if ((mailListCnt < 0) || (mailListCnt > MboxMail::s_mails_ref.GetCount()) ||
		(mailFileSize != nMailFileSize) || (mboxFileSize != nMboxFileSize))
	{
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	BOOL errorDoNothing = FALSE;
	BOOL verifyOnly = TRUE;
	pos = 0;
	retval = sz.GetReadPointer(pos);
	if (retval == FALSE)
	{
		int deb = 1;
	}

	ret = PopulateUserMailArray(sz, mailListCnt, verifyOnly);
	if (ret > 0) 
	{
		verifyOnly = FALSE;
		BOOL retval = sz.SetReadPointer(pos);
		if (retval == FALSE)
		{
			int deb = 1;
		}
		ret = PopulateUserMailArray(sz, mailListCnt, verifyOnly);
	}
	else
	{
		sz.close();
#if 0
		CString txt = L"Mail list file\n\"" + mboxListFile;
		txt += L"\"\n is invalid or corrupted. Can't reload.\nPlease remove the file.";
#endif
		CString txt;
		CString fmt = L"Mail list file\n\"%s\"\n is invalid or corrupted. Can't reload.\nPlease remove the file.";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, mboxListFile);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);

		return -1; // do nothing
	}

	sz.close();

	if (ret < 0)
	{
		MboxMail::s_mails_edit.SetSizeKeepData(0);
#if 0
		CString txt = L"Mail list file\n\"" + mboxListFile;
		txt += L"\"\n is invalid or corrupted. Can't reload.\nPlease remove the file.";
#endif
		CString txt;
		CString fmt = L"Mail list file\n\"%s\"\n is invalid or corrupted. Can't reload.\nPlease remove the file.";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, mboxListFile);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
	}
	else
	{
		MboxMail::SortByDate(&MboxMail::s_mails_edit);
		if (MboxMail::s_mails_edit.GetCount() > 0)
			MboxMail::m_editMails.m_lastSel = 0;
		else
			MboxMail::m_editMails.m_lastSel = -1;
		MboxMail::m_editMails.b_mails_which_sorted = 1;

		if (MboxMail::IsUserMailsSelected())
		{
			SwitchToMailList(IDC_EDIT_LIST, TRUE);
		}
		else if (MboxMail::IsAllMailsSelected())
		{
			SwitchToMailList(IDC_EDIT_LIST, FALSE);
		}
	}

	MboxMail::m_editMails.m_bIsDirty = FALSE;

	return ret;
}

int NListView::SaveAsMboxListFile_v2()
{
	MailArray *mailsArray = &MboxMail::s_mails;

	CString mboxFileSuffix;
	if (MboxMail::IsUserMailsSelected())
		mboxFileSuffix = L"_USER.mbox";
	else if (MboxMail::IsFindMailsSelected())
		mboxFileSuffix = L"_FIND.mbox";
	else
	{
		// We should never be here
		return -1;
	}

	MboxMail *m;

	CString mailFile = MboxMail::s_path;

	if (MboxMail::s_path.IsEmpty()) {
		CString txt = L"Please open mail file first.";
		ResHelper::TranslateString(txt);
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	CString path = MboxMail::GetLastPath();
	CString datapath = MboxMail::GetLastDataPath();
	if (path.IsEmpty())
		return -1;  // Hopefully s_path wil fail first

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(mailFile, driveName, directory, fileNameBase, fileNameExtention);

	CString errorText;
	CString archiveCachePath;
	CString rootPrintSubFolder = L"ListCache";
	CString targetPrintSubFolder;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, archiveCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CFile fp;
	CString mboxFile = fileNameBase + fileNameExtention;
	CString mboxFilePath = archiveCachePath + L"\\" + fileNameBase + mboxFileSuffix;

	CString mboxFileListSuffix = L".mboxlist";
	CString mboxListFile = mboxFilePath + mboxFileListSuffix;

	BOOL ret2 = FALSE;
	int ret1 = 0;
	int ret0 = 0;

	if (FileUtils::PathFileExist(mboxListFile))
	{
		CString mboxListFileBak1 = mboxListFile + L".bak1";
		CString mboxListFileBak2 = mboxListFile + L".bak2";

		if (FileUtils::PathFileExist(mboxListFileBak2))
			ret2 = FileUtils::DelFile(mboxListFileBak2);

		if (FileUtils::PathFileExist(mboxListFileBak1))
			ret1 = _wrename(mboxListFileBak1, mboxListFileBak2);

		ret0 = _wrename(mboxListFile, mboxListFileBak1);
	}

	SerializerHelper sz(mboxListFile);
	if (!sz.open(TRUE)) 
	{
#if 0
		CString txt = L"Could not create \"" + mboxListFile;
		txt += L"\" file.\nMake sure file is not open on other applications.";
#endif

		CString txt;
		CString fmt = L"Could not create \"%s\" file.\nMake sure file is not open on other applications.";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, mboxListFile);

		int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	// Create mboxlist file to allow reload of archive file list

	sz.writeInt(MAIL_LIST_VERSION);			// version
	//sz.writeString(mboxFile);  // TODO: ??
	// main mbox mail file size other mbox files are derived/contain subset
	sz.writeInt64(MboxMail::s_fSize);	
	// we don't always create new mail archive corresponding to mboxListFile/mail list
	// mailFileSize is mboxMailFileSize corresponding to mboxListFile
	_int64 mailFileSize = -1;  //  FileSize(mboxFilePath);  
	sz.writeInt64(mailFileSize);	// file size
	sz.writeInt(mailsArray->GetSize());

	for (int i = 0; i < mailsArray->GetSize(); i++)
	{
		m = (*mailsArray)[i];
		sz.writeInt64(m->m_startOff);
		sz.writeInt(m->m_length);
		sz.writeInt(m->m_index);
		sz.writeInt((int)m->m_ContentDetailsArray.size());
	}
	sz.close();

	MboxMail::m_editMails.m_bIsDirty = FALSE;
	return 1;
}

int NListView::SaveAsMboxArchiveFile_v2()
{
	MailArray *mailsArray = &MboxMail::s_mails;

	CString mboxFileSuffix;
	if (MboxMail::IsUserMailsSelected())
		mboxFileSuffix = L"_USER.mbox";
	else if (MboxMail::IsFindMailsSelected())
		mboxFileSuffix = L"_FIND.mbox";
	else
	{
		// We should never be here
		return -1;
	}

	MboxMail *m;

	CString mailFile = MboxMail::s_path;

	if (MboxMail::s_path.IsEmpty()) {
		CString txt = L"Please open mail file first.";
		ResHelper::TranslateString(txt);
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	CString path = MboxMail::GetLastPath();
	CString datapath = MboxMail::GetLastDataPath();
	if (path.IsEmpty())
		return -1;  // Hopefully s_path wil fail first

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(mailFile, driveName, directory, fileNameBase, fileNameExtention);

	CString errorText;
	CString archiveCachePath;
	CString rootPrintSubFolder = L"ArchiveCache";
	CString targetPrintSubFolder;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, archiveCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CFile fp;

	CString mboxFile = fileNameBase + fileNameExtention;
	CString archiveFile = fileNameBase + mboxFileSuffix;;
	CString mboxFilePath = archiveCachePath + L"\\" + fileNameBase + mboxFileSuffix;

	if (FileUtils::PathFileExist(mboxFilePath))
	{
		CString txt;
		CString fmt = L"File\n\n\"%s\"\n\nexists.\nOverwrite?";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, mboxFilePath);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
			return -1;
	}

	CFileException ExError;
	if (!fp.Open(mboxFilePath, CFile::modeWrite | CFile::modeCreate, &ExError))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);
#if 0
		CString txt = L"Could not create \"" + mboxFilePath;
		txt += L"\" file.\n";
		txt += exErrorStr;
#endif
		CString txt;
		CString fmt = L"Could not create \"%s\" file.\n%s";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, mboxFilePath, exErrorStr);

		int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	SimpleString *outbuf = MboxMail::m_outbuf;
	outbuf->ClearAndResize(10000);

	BOOL ret;
	for (int i = 0; i < mailsArray->GetSize(); i++)
	{
		m = (*mailsArray)[i];
		outbuf->Clear();
		ret = m->GetBodySS(outbuf);
		// TODO: may need to validate last mail, i.e check for From presence
		fp.Write(outbuf->Data(), outbuf->Count());
	}
	fp.Close();

	MboxMail::m_editMails.m_bIsDirty = FALSE;

	//CString txt = L"Created Mbox Mail Archive file \n\n" + mboxFilePath;
	CString txt = L"Created Mbox Mail Archive file";
	ResHelper::TranslateString(txt);

	BOOL supressOpenFileOption = FALSE;
	if (path.Compare(datapath) == 0)
		supressOpenFileOption = TRUE;
	OpenContainingFolderDlg dlg(txt, mboxFilePath, supressOpenFileOption);
	INT_PTR nResponse = dlg.DoModal();
	////////////
	if (nResponse == IDOK)
	{
		if (FileUtils::BrowseToFile(mboxFilePath) == FALSE) {
			HWND h = GetSafeHwnd();
			HINSTANCE result = ShellExecute(h, L"open", archiveCachePath, NULL, NULL, SW_SHOWNORMAL);
			CMainFrame::CheckShellExecuteResult(result, h);
		}
		int deb = 1;
	}
	else if (nResponse == IDYES)
	{
		//CString txt = L"Open Created Archive File \n\n" + mboxFilePath;
		CString txt;
		CString fmt = L"Open Created Archive File \n\n%s";
		txt.Format(fmt, mboxFilePath);
		OpenArchiveFileDlg dlg;
		dlg.m_sourceFolder = archiveCachePath;
		dlg.m_targetFolder = path;
		dlg.m_archiveFileName = archiveFile;
		INT_PTR nResponse = dlg.DoModal();
		////////////
		if (nResponse == IDOK)
		{
			CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
			if (pFrame)
			{
				CString archiveFilePath = dlg.m_targetFolder + dlg.m_archiveFileName;

				DWORD nFlags = MOVEFILE_REPLACE_EXISTING|MOVEFILE_WRITE_THROUGH|MOVEFILE_COPY_ALLOWED;
				BOOL retMove = MoveFileEx(mboxFilePath, archiveFilePath, nFlags);
				if (retMove == FALSE)
				{
					CString errorText = FileUtils::GetLastErrorAsString();
					TRACE(L"SaveAsMboxArchiveFile_v2: MoveFileEx failed \"%s\"\n", errorText);
				}

				//pFrame->DoOpen(archiveFilePath);

				NTreeView *pTreeView = pFrame->GetTreeView();
				if (pTreeView)
				{
					// delete index file to make sure it is not used in case old and new length of new mbox file are the same
					// InsertMailFile will delete index file
					//CString indexFile = archiveFilePath + ".mboxview";
					//DeleteFile(indexFile);

					pTreeView->InsertMailFile(archiveFilePath);
				}
			}
		}
		else if (nResponse == IDCANCEL)
		{
			int deb = 1;
		}
		int deb = 1;
	}
	else if (nResponse == IDCANCEL)
	{
		int deb = 1;
	}
	return 1;
}

int NListView::FindMailListFileWithHighestNumber(CString &folder, CString &extension)
{
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

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

			filePath.Empty();
			filePath.Append(wf.cFileName);
			FileUtils::SplitFilePath(filePath, driveName, directory, fileNameBase, fileNameExtention);

			fileCnt++;
			break;
		} while (FindNextFile(f, &wf));
		FindClose(f);
	}
	return fileCnt;
}


void NListView::OnClose()
{
	// TODO: Add your message handler code here and/or call default

	CWnd::OnClose();
}

void NListView::FindImageFileName(CString &cid)
{
	std::vector <MailBodyInfo*>::iterator it;

	if (m_BodyInfoArray.size() > 1)
		int deb = 1;

	MailBodyInfo *body;
	for (it = m_BodyInfoArray.begin(); it != m_BodyInfoArray.end(); it++)
	{
		body = *it;
	}
}

int NListView::FindFilenameCount(std::vector <MailBodyContent*> &contentDetailsArray, CStringA &fileName)
{
	int fileCnt = 0;
	MailBodyContent *body;
	for (int j = 0; j < contentDetailsArray.size(); j++)
	{
		body = contentDetailsArray[j];

		if (fileName.Compare(body->m_attachmentName) == 0)
		{
			fileCnt++;
		}
	}
	return fileCnt;
}

int NListView::DetermineImageFileName(CFile* fpm, BOOL verifyAttachmentDataAsImageType, MboxMail *m, CStringA &cidName, CString &imageFilePath, MailBodyContent **foundBody, int mailPosition)
{
	if (m == 0)
	{
		// TODO: Assert ??
		return 1;
	}

	SimpleString cid;
	//CString imageCachePath;
	CStringA imageFileName;
	CStringA nameExtension;
	CStringA idExtension;
	CStringA locationExtension;
	MailBodyContent *body;

	CStringA contentTypeExtension;
	CStringA contentTypeMain;
	CStringA contentSubType;

	BOOL isOctetStream;

	*foundBody = 0;

	CStringA mailIndex;
	//mailIndex.Format("%07d_", m->m_index);   // FIIXME mailIndex is empty; do we need index anymore ?

	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];

		contentTypeExtension.Empty();
		contentTypeMain.Empty();
		contentSubType.Empty();
		nameExtension.Empty();
		idExtension.Empty();
		locationExtension.Empty();


		BOOL isValidContentTypeExtension = FALSE;
		int pos = body->m_contentType.ReverseFind('/');
		if (pos > 0)
		{
			contentSubType = body->m_contentType.Mid(pos + 1);
			contentTypeMain = body->m_contentType.Left(pos);
			CString contentTypeExt = "." + contentSubType;
			isValidContentTypeExtension = GdiUtils::IsSupportedPictureFileExtension(contentTypeExt);
			contentTypeExtension = contentTypeExt;

		}

		if (contentTypeMain.CompareNoCase("message") == 0)
		{
			TRACE(L"Content-Type=message index=%d\n", mailPosition);
			int deb = 1;
		}

		isOctetStream = FALSE;
		if (contentSubType.CompareNoCase("octet-stream") == 0)
			isOctetStream = TRUE;
		
		if ((contentTypeMain.CompareNoCase("image") == 0) || 
			(body->m_contentDisposition.CompareNoCase("inline") == 0) || isOctetStream)
		{
			BOOL isValidAttachmentNameExtension = FALSE;
			int pos = body->m_attachmentName.ReverseFind('.');
			if (pos >= 0)
			{
				nameExtension = body->m_attachmentName.Mid(pos);
				CString nameExtensionW = nameExtension;
				isValidAttachmentNameExtension = GdiUtils::IsSupportedPictureFileExtension(nameExtensionW);
			}

			BOOL isValidContentIdExtension = FALSE;
			pos = body->m_contentId.ReverseFind('.');
			if (pos >= 0)
			{
				idExtension = body->m_contentId.Mid(pos);
				CString idExtensionW = idExtension;
				isValidContentIdExtension = GdiUtils::IsSupportedPictureFileExtension(idExtensionW);
			}

			BOOL isValidContentLocationExtension = FALSE;
			pos = body->m_contentLocation.ReverseFind('.');
			if (pos >= 0)
			{
				locationExtension = body->m_contentLocation.Mid(pos);
				CString idExtensionW = idExtension;
				isValidContentLocationExtension = GdiUtils::IsSupportedPictureFileExtension(idExtensionW);
			}

			CStringA attachmentName = body->m_attachmentName;

			if (!body->m_contentId.IsEmpty())
			{
				CStringA contentId = body->m_contentId;
				if (contentId.Compare(cidName) == 0)
				{
					// TODO: we assume all contentIds are unique which may not be always true
					int fileCnt = 0;
					if (!attachmentName.IsEmpty())
					{
						fileCnt = FindFilenameCount(m->m_ContentDetailsArray, attachmentName);
					}

					if (!attachmentName.IsEmpty() && (fileCnt < 2))
					{
						if (isValidAttachmentNameExtension)
							imageFileName = mailIndex + attachmentName;
						else if (isValidContentTypeExtension)
							imageFileName = mailIndex + attachmentName + contentTypeExtension;
						else if (isValidContentIdExtension)
							imageFileName = mailIndex + attachmentName + idExtension;
						else if (isValidContentLocationExtension)
							imageFileName = mailIndex + attachmentName + locationExtension;
						else
						{
							// Only Legacy Internet Explorer cares about proper file extension
							// Firefox, Edge and Chrome don't care about image file extension
							// They take action based on the file meta data
							imageFileName = mailIndex + attachmentName + ".jpg";
						}
					}
					else if (isValidContentIdExtension)
					{
						imageFileName = mailIndex + contentId;
					}
					else if (isValidAttachmentNameExtension)
					{
						imageFileName = mailIndex + contentId + nameExtension;
					}
					else if (isValidContentTypeExtension)
					{
						imageFileName = mailIndex + contentId + contentTypeExtension;
					}
					else
					{
						imageFileName = mailIndex + contentId + ".jpg";
					}
					*foundBody = body;
					// break;  // FIXMEFIXME
				}
				else
				{
					int deb = 1;
				}
			}

			if ((*foundBody == 0) && !body->m_contentLocation.IsEmpty())
			{
				CStringA contentLocation = body->m_contentLocation;
				if (contentLocation.Compare(cidName) == 0)
				{
					// TODO: we assume all contentLocations are unique which may not be always true
					int fileCnt = 0;
					if (!attachmentName.IsEmpty())
						fileCnt = FindFilenameCount(m->m_ContentDetailsArray, attachmentName);

					if (!attachmentName.IsEmpty() && (fileCnt < 2))
					{
						if (isValidAttachmentNameExtension)
							imageFileName = mailIndex + attachmentName;
						else if (isValidContentTypeExtension)
							imageFileName = mailIndex + attachmentName + contentTypeExtension;
						else if (isValidContentIdExtension)
							imageFileName = mailIndex + attachmentName + idExtension;
						else if (isValidContentLocationExtension)
							imageFileName = mailIndex + attachmentName + locationExtension;
						else
						{
							// Only Legacy Internet Explorer cares about proper file extension
							// Firefox, Edge and Chrome don't care about image file extension
							// They take action based on the file meta data
							imageFileName = mailIndex + attachmentName + ".jpg";
						}
					}
					else if (isValidContentLocationExtension)
					{
						imageFileName = mailIndex + contentLocation;
					}
					else if (isValidAttachmentNameExtension)
					{
						imageFileName = mailIndex + contentLocation + nameExtension;
					}
					else if (isValidContentTypeExtension)
					{
						imageFileName = mailIndex + contentLocation + contentTypeExtension;
					}
					else
					{
						// Only Legacy Internet Explorer cares about proper file extension
						// Firefox, Edge and Chrome don't care about image file extension
						// They take action based on the file meta data
						imageFileName = mailIndex + contentLocation + ".jpg";
					}
					*foundBody = body;
					// break // FIXME
				}
				else
				{
					int deb = 1;
				}
			}

			if ((*foundBody == 0) && !body->m_attachmentName.IsEmpty())
			{
				if (attachmentName.Compare(cidName) == 0)
				{
					if (isValidAttachmentNameExtension)
						imageFileName = mailIndex + attachmentName;
					else if (isValidContentTypeExtension)
						imageFileName = mailIndex + attachmentName + contentTypeExtension;
					else if (isValidContentIdExtension)
						imageFileName = mailIndex + attachmentName + idExtension;
					else if (isValidContentLocationExtension)
						imageFileName = mailIndex + attachmentName + locationExtension;
					else
					{
						// Only Legacy Internet Explorer cares about proper file extension
						// Firefox, Edge and Chrome don't care about image file extension
						// They take action based on the file meta data

						imageFileName = mailIndex + attachmentName + ".jpg";
					}
					*foundBody = body;
					//break;  // FIXME
				}
				else
				{
					int deb = 1;
				}

			}
			if (*foundBody == 0) {
				int deb = 1;
			}
			int deb = 1;
		}

		// Below is only needed for Interent Explorer
		// Firefox, Edge and Chrome don't really care about image file extension
		// They take action based on file meta data

		if (*foundBody)
		{
			if (verifyAttachmentDataAsImageType == FALSE)
				break;

			if (!imageFileName.IsEmpty())
			{
				int pos = imageFileName.ReverseFind('.');
				if (pos > 0)  // extension found
				{
					nameExtension = imageFileName.Mid(pos);
					CStringA baseName = imageFileName.Mid(0, pos);


					// In few cases the file name extension is incorrect and needs to be corercted in order for Internet Explorer to display the image instead of the blank image.
					// Othet Browser such as Chrome, Firefox don't need this fix.
					//if ((nameExtension.CompareNoCase(".gif") == 0) || (nameExtension.CompareNoCase(".png") == 0) || (nameExtension.CompareNoCase(".bmp") == 0))
					{
						SimpleString* outbuf = MboxMail::get_tmpbuf();
						outbuf->ClearAndResize(body->m_contentLength * 2);

						int retLen = MboxMail::DecodeBody(*fpm, body, mailPosition, outbuf);

						// FIXME: fpm unknown here
						const unsigned char* data = (const unsigned char*)outbuf->Data();
						int dataLength = outbuf->Count();

						CStringW extensionW;
						CStringA extension;
						BOOL ret = GdiUtils::loadImage((BYTE*)data, (size_t)dataLength, extensionW, extension);
						if (ret)
						{
							if ((nameExtension.CompareNoCase(".jpg") == 0) && (extension.CompareNoCase(".jpeg") == 0))
								;
							else if ((nameExtension.CompareNoCase(".jpeg") == 0) && (extension.CompareNoCase(".jpg") == 0))
								;
							else if (nameExtension.CompareNoCase(extension) != 0)
							{
								imageFileName = baseName + extension;
							}
							int deb = 1;
						}

						MboxMail::rel_tmpbuf();
					}
				}
				else
				{
					int deb = 1;
				}
			}
			else  // imageFileName.IsEmpty()
			{
				imageFileName = "XXX.jpg"; // FIXME
			}
			break;
		}
	}
	BOOL bReplaceWhiteWithUnderscore = TRUE;
	BOOL extraValidation = TRUE;
	FileUtils::MakeValidFileNameA(imageFileName, bReplaceWhiteWithUnderscore, extraValidation);
	imageFilePath = imageFileName;

#if 0
	TRACE(L"--->>> SUBJ=%s\n", m->m_subj);
	if (imageFileName.IsEmpty())
	{
		TRACE(L"\tCID=%s FNAME=%s\n", cidName, imageFilePath);
		int deb = 1;
	}
	else
	{
		TRACE(L"\tCID=%s FNAME=%s\n", cidName, imageFilePath);
		int deb = 1;
	}
#endif
	return 1;
}

int NListView::FindFilenameCount(CMimeBody::CBodyList &bodies,  CStringA &fileName)
{
	CMimeBody* pBP;
	string strName;
	strName.resize(128);
	int fileCnt = 0;
	CMimeBody::CBodyList::const_iterator it;
	for (it = bodies.begin(); it != bodies.end(); it++)
	{
		pBP = *it;

		string strName = pBP->GetFilename();
		if (strName.empty())
			strName = pBP->GetName();

		if (fileName.Compare(strName.c_str()) == 0)
		{
			fileCnt++;
		}
	}
	return fileCnt;
}

//
int NListView::CreateAttachmentCache_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName)
{
	WRITE_IMAGE_FILE_ARGS args;

	CString rootPrintSubFolder = L"AttachmentCache";
	CString attachmentCachePath;
	CString errorText;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolderName, attachmentCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = GetSafeHwnd(); // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	/*IN*/
	args.cache = "";


	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcWriteAttachmentFile, (LPVOID)(WRITE_IMAGE_FILE_ARGS*)&args);
	Dlg.SetDialogTemplate(AfxGetApp()->m_hInstance, MAKEINTRESOURCE(IDD_PROGRESS_DLG), IDC_STATIC, IDC_PROGRESS_BAR, IDCANCEL);

	INT_PTR nResult = Dlg.DoModal();

	if (!nResult) { // should never be true ?
		// TODO: errorText ??
		MboxMail::assert_unexpected();
		return -1;
	}

	int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
	int retResult = LOWORD(nResult);

	if (retResult != IDOK)
	{  // IDOK==1, IDCANCEL==2
		// We should be here when user selects Cancel button
		//_ASSERTE(cancelledbyUser == TRUE);

		DWORD terminationDelay = Dlg.GetTerminationDelay();
		int loopCnt = (terminationDelay+100)/25;

		ULONGLONG tc_start = GetTickCount64();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		ULONGLONG tc_end = GetTickCount64();
		DWORD delta = (DWORD)(tc_end - tc_start);
		TRACE(L"(CreateAttachmentCache_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret > 0)
	{
		if (!args.errorText.IsEmpty())  // TODO: should not  be true
		{
			MboxMail::assert_unexpected();
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}

	CString txt = L"Do you want to open folder with all attachments?";
	ResHelper::TranslateString(txt);

	OpenContainingFolderDlg dlg(txt, attachmentCachePath, TRUE);
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		HWND h = GetSafeHwnd();
		HINSTANCE result = ShellExecute(h, L"open", attachmentCachePath, NULL, NULL, SW_SHOWNORMAL);
		CMainFrame::CheckShellExecuteResult(result, h);
		int deb = 1;
	}
	else if (nResponse == IDYES)
	{
		HWND h = GetSafeHwnd();
		HINSTANCE result = ShellExecute(h, L"open", attachmentCachePath, NULL, NULL, SW_SHOWNORMAL);
		CMainFrame::CheckShellExecuteResult(result, h);
		int deb = 1;
	}
	else if (nResponse == IDCANCEL)
	{
		int deb = 1;
	}
	return 1;
}

BOOL CreateAttachmentCache_WorkerThread(LPCWSTR cache, BOOL mainThread, CString &errorText)
{
	CFile fpm;
	CFileException ExError;
	if (!fpm.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
	{
		// TODO: critical failure
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not open \"" + MboxMail::s_path;
		txt += L"\" mail file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);

		errorText = txt;

		return FALSE;
	}

	int ni = MboxMail::s_mails.GetSize();

	//CString errorText;
	CString attachmentCachePath;
	CString rootPrintSubFolder = L"AttachmentCache";
	CString targetPrintSubFolder;

	BOOL createDirOk = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, attachmentCachePath, errorText);

	//
	if (FileUtils::PathDirExists(attachmentCachePath))
	{
		CString txt = L"Deleting all attachment files in the Attachment Cache directory ...";
		ResHelper::TranslateString(txt);
		MboxMail::pCUPDUPData->SetProgress(txt, 0);
		FileUtils::RemoveDir(attachmentCachePath, false);
	}
	//
	// FileUtils::RemoveDirW(attachmentCachePath, false); removes files only; no need to recreate attachmentCachePath directory
#if 0
	attachmentCachePath.Empty();
	rootPrintSubFolder = "AttachmentCache";
	targetPrintSubFolder.Empty();

	createDirOk = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, attachmentCachePath, errorText);
#endif

	if ((fpm.m_hFile == CFile::hFileNull) || !createDirOk)
	{
		return TRUE;
	}

	if (!mainThread)
	{
		if (MboxMail::pCUPDUPData)
		{
			CString txt = L"Creating attachment files ...";
			ResHelper::TranslateString(txt);
			MboxMail::pCUPDUPData->SetProgress(txt, 0);
		}
	}

	BOOL terminated = FALSE;
	CString itemNumberStr;
	std::vector <MailBodyContent*>::iterator it;


	if (!mainThread)
	{
		if (MboxMail::pCUPDUPData)
		{
			CString txt = L"Creating attachment files ...";
			ResHelper::TranslateString(txt);
			MboxMail::pCUPDUPData->SetProgress(txt, 0);
		}
	}

	AttachmentMgr m_attachmentDB;

	ULONGLONG workRangeFirstPos = 0;
	ULONGLONG workRangeLastPos = ni - 1;
	ProgressTimer progressTimer(workRangeFirstPos, workRangeLastPos);

	MboxMail *m;
	for (int i = 0; i < ni; i++)
	{
		m_attachmentDB.Clear();

		m = MboxMail::s_mails[i];

		int mailPosition = i;
		CString* attachmentFolderPath = 0;
		BOOL prependMailId = TRUE;
		NListView::CreateMailAttachments(&fpm, mailPosition, attachmentFolderPath, prependMailId, m_attachmentDB);

		if (!mainThread && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				terminated = TRUE;
				break;
			}

			UINT_PTR dwProgressbarPos = 0;
			ULONGLONG workRangePos = i;
			BOOL needToUpdateStatusBar = progressTimer.UpdateWorkPos(workRangePos, dwProgressbarPos);
			if (needToUpdateStatusBar)
			{
				int nItem = i + 1;
				//itemNumberStr.Format(L"Creating attachment files ... %d of %d", nItem, ni);

				CString fmt = L"Creating attachment files ... %d of %d";
				ResHelper::TranslateString(fmt);
				itemNumberStr.Format(fmt, nItem, ni);

				if (MboxMail::pCUPDUPData)  MboxMail::pCUPDUPData->SetProgress(itemNumberStr, (UINT_PTR)(dwProgressbarPos));
				int debug = 1;
			}
		}

		if (terminated)
		{
			errorText = L"Already created attachment files will not be deleted.";
		}
	}

	fpm.Close();
	return TRUE;
}
//
//
int NListView::CreateEmlCache_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName)
{
	WRITE_IMAGE_FILE_ARGS args;

	CString rootPrintSubFolder = L"EmlCache";
	CString emlCachePath;
	CString errorText;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolderName, emlCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = GetSafeHwnd(); // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	/*IN*/
	args.cache = "";


	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcWriteEmlFile, (LPVOID)(WRITE_IMAGE_FILE_ARGS*)&args);
	Dlg.SetDialogTemplate(AfxGetApp()->m_hInstance, MAKEINTRESOURCE(IDD_PROGRESS_DLG), IDC_STATIC, IDC_PROGRESS_BAR, IDCANCEL);

	INT_PTR nResult = Dlg.DoModal();

	if (!nResult) { // should never be true ?
		// TODO: errorText ??
		MboxMail::assert_unexpected();
		return -1;
	}

	int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
	int retResult = LOWORD(nResult);

	if (retResult != IDOK)
	{  // IDOK==1, IDCANCEL==2
		// We should be here when user selects Cancel button
		//_ASSERTE(cancelledbyUser == TRUE);

		DWORD terminationDelay = Dlg.GetTerminationDelay();
		int loopCnt = (terminationDelay+100)/25;

		ULONGLONG tc_start = GetTickCount64();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		ULONGLONG tc_end = GetTickCount64();
		DWORD delta = (DWORD)(tc_end - tc_start);
		TRACE(L"(CreateEmlCache_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret > 0)
	{
		if (!args.errorText.IsEmpty())  // TODO: should not  be true
		{
			MboxMail::assert_unexpected();
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}

	CString txt = L"Do you want to open folder with all Eml files?";
	ResHelper::TranslateString(txt);

	OpenContainingFolderDlg dlg(txt, emlCachePath, TRUE);
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		HWND h = GetSafeHwnd();
		HINSTANCE result = ShellExecute(h, L"open", emlCachePath, NULL, NULL, SW_SHOWNORMAL);
		CMainFrame::CheckShellExecuteResult(result, h);
		int deb = 1;
	}
	else if (nResponse == IDYES)
	{
		HWND h = GetSafeHwnd();
		HINSTANCE result = ShellExecute(h, L"open", emlCachePath, NULL, NULL, SW_SHOWNORMAL);
		CMainFrame::CheckShellExecuteResult(result, h);
		int deb = 1;
	}
	else if (nResponse == IDCANCEL)
	{
		int deb = 1;
	}
	return 1;
}

//
BOOL CreateEmlCache_WorkerThread(LPCWSTR cache, BOOL mainThread, CString &errorText)
{
	CFile fpm;
	CFileException ExError;
	if (!fpm.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
	{
		// TODO: critical failure
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not open \"" + MboxMail::s_path;
		txt += L"\" mail file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);

		errorText = txt;
		return FALSE;
	}

#if 0
	CString mailFileName;
	FileUtils::CPathStripPath((char*)(LPCSTR)MboxMail::s_path, mailFileName);
	int pos = mailFileName.ReverseFind('.');
	CString baseFailFileName;
	if (pos >= 0)
		baseFailFileName = mailFileName.Mid(0, pos);
	else
		baseFailFileName = mailFileName;

	int ni = MboxMail::s_mails.GetSize();

	CString emlPath;
	BOOL ret = FileUtils::CPathGetPath(MboxMail::s_path, emlPath);
	emlPath.Append("\\");
	emlPath.Append("EmlCache");
	emlPath.Append("\\");

	CString emlCachePath = emlPath;
	emlCachePath.Append(baseFailFileName);

	//
	if (FileUtils::PathDirExists(emlCachePath))
	{
		MboxMail::pCUPDUPData->SetProgress(L"Deleting all Eml files in the Eml Cache directory ...", 0);
		FileUtils::RemoveDirW(emlCachePath, false);
	}
	//

	emlCachePath.Append("\\");

	BOOL createDirOk = TRUE;
	if (!FileUtils::PathDirExists(emlPath))
		createDirOk = CreateDirectory(emlPath, NULL);;

	if (!FileUtils::PathDirExists(emlCachePath))
		createDirOk = CreateDirectory(emlCachePath, NULL);
#endif

	int ni = MboxMail::s_mails.GetSize();

	//CString errorText;
	CString emlCachePath;
	CString rootPrintSubFolder = L"EmlCache";
	CString targetPrintSubFolder;

	BOOL createDirOk = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, emlCachePath, errorText);

	//
	if (FileUtils::PathDirExists(emlCachePath))
	{
		CString txt = L"Deleting all Eml files in the Eml Cache directory ...";
		ResHelper::TranslateString(txt);
		MboxMail::pCUPDUPData->SetProgress(txt, 0);


		FileUtils::RemoveDir(emlCachePath, false);
	}
	//

#if 0
	emlCachePath.Empty();
	rootPrintSubFolder = "EmlCache";
	targetPrintSubFolder.Empty();

	createDirOk = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, emlCachePath, errorText);
#endif

	emlCachePath.Append(L"\\");


	if ((fpm.m_hFile == CFile::hFileNull) || !createDirOk)
	{
		return TRUE;
	}

	if (!mainThread)
	{
		if (MboxMail::pCUPDUPData)
		{
			CString txt = L"Creating Eml files ...";
			ResHelper::TranslateString(txt);
			MboxMail::pCUPDUPData->SetProgress(txt, 0);
		}
	}

	double delta = (double)ni;

	BOOL terminated = FALSE;
	CString itemNumberStr;
	std::vector <MailBodyContent*>::iterator it;


	if (!mainThread)
	{
		if (MboxMail::pCUPDUPData)
		{
			CString txt = L"Creating Eml files ...";
			ResHelper::TranslateString(txt);
			MboxMail::pCUPDUPData->SetProgress(txt, 0);
		}
	}

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	ULONGLONG workRangeFirstPos = 0;
	ULONGLONG workRangeLastPos = ni - 1;
	ProgressTimer progressTimer(workRangeFirstPos, workRangeLastPos);

	MboxMail *m;
	CString emlFile;
	for (int i = 0; i < ni; i++)
	{
		emlFile.Empty();
		m = MboxMail::s_mails[i];

		int mailPosition = i;
		NamePatternParams *pNamePP = &pFrame->m_NamePatternParams;
		NListView::PrintAsEmlFile(&fpm, mailPosition, emlFile);

		if (!mainThread && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				terminated = TRUE;
				break;
			}

			UINT_PTR dwProgressbarPos = 0;
			ULONGLONG workRangePos = i;
			BOOL needToUpdateStatusBar = progressTimer.UpdateWorkPos(workRangePos, dwProgressbarPos);
			if (needToUpdateStatusBar)
			{
				int nItem = i + 1;

				CString fmt = L"Creating Eml files ... %d of %d";
				ResHelper::TranslateString(fmt);
				itemNumberStr.Format(fmt, nItem, ni);

				if (MboxMail::pCUPDUPData)  MboxMail::pCUPDUPData->SetProgress(itemNumberStr, (UINT_PTR)(dwProgressbarPos));

				int debug = 1;
			}
		}

		if (terminated)
		{
			errorText = L"Already created Eml files will not be deleted.";
		}
	}

	fpm.Close();
	return TRUE;
}

//

int NListView::CreateInlineImageCache_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName)
{
	WRITE_IMAGE_FILE_ARGS args;

	CString rootPrintSubFolder = L"ImageCache";
	_ASSERTE(NListView::m_exportMailsMode == FALSE);
	if (NListView::m_exportMailsMode)
	{
		rootPrintSubFolder = L"ExportCache";
		targetPrintSubFolderName = L"Attachments";
	}
	CString imageCachePath;
	CString errorText;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolderName, imageCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = GetSafeHwnd(); // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	/*IN*/
	args.cache = imageCachePath;

	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcWriteInlineImageFile, (LPVOID)(WRITE_IMAGE_FILE_ARGS*)&args);
	Dlg.SetDialogTemplate(AfxGetApp()->m_hInstance, MAKEINTRESOURCE(IDD_PROGRESS_DLG), IDC_STATIC, IDC_PROGRESS_BAR, IDCANCEL);

	INT_PTR nResult = Dlg.DoModal();

	if (!nResult) { // should never be true ?
		// TODO: errorText ??
		MboxMail::assert_unexpected();
		return -1;
	}

	int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
	int retResult = LOWORD(nResult);

	if (retResult != IDOK)
	{  // IDOK==1, IDCANCEL==2
		// We should be here when user selects Cancel button
		//_ASSERTE(cancelledbyUser == TRUE);

		DWORD terminationDelay = Dlg.GetTerminationDelay();
		int loopCnt = (terminationDelay+100)/25;

		ULONGLONG tc_start = GetTickCount64();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		ULONGLONG tc_end = GetTickCount64();
		DWORD delta = (DWORD)(tc_end - tc_start);
		TRACE(L"(CreateInlineImageCache_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret > 0)
	{
		if (!args.errorText.IsEmpty())  // TODO: should not  be true
		{
			MboxMail::assert_unexpected();
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}
	return 1;
}

BOOL CreateInlineImageCache_WorkerThread(LPCWSTR cache, BOOL mainThread, CString &errorText)
{
	CFile fpm;
	CFileException ExError;
	if (!fpm.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
	{
		// TODO: critical failure
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not open \"" + MboxMail::s_path;
		txt += L"\" mail file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);

		errorText = txt;
		return FALSE;
	}

	CString mailFileName;
	FileUtils::CPathStripPath((LPCWSTR)MboxMail::s_path, mailFileName);
	int pos = mailFileName.ReverseFind('.');
	CString baseFailFileName;
	if (pos >= 0)
		baseFailFileName = mailFileName.Mid(0, pos);
	else
		baseFailFileName = mailFileName;

#if 0
	int ni = MboxMail::s_mails.GetSize();

	CString imageCachePath;
	BOOL ret = FileUtils::CPathGetPath(MboxMail::s_path, imageCachePath);
	imageCachePath.Append("\\");
	imageCachePath.Append("ImageCache");
	imageCachePath.Append("\\");

	BOOL createDirOk = TRUE;
	if (!FileUtils::PathDirExists(imageCachePath))
		createDirOk = CreateDirectory(imageCachePath, NULL);

	imageCachePath.Append(baseFailFileName);
	imageCachePath.Append("\\");

	if (!FileUtils::PathDirExists(imageCachePath))
		createDirOk = CreateDirectory(imageCachePath, NULL);

#endif
	//
	CString imageCachePath;
	CString rootPrintSubFolder = L"ImageCache";
	CString targetPrintSubFolder;

	BOOL createDirOk = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, imageCachePath, errorText);
	if ((fpm.m_hFile == CFile::hFileNull) || !createDirOk)
	{
		return TRUE;
	}
	imageCachePath.Append(L"\\");

	if (!mainThread)
	{
		if (MboxMail::pCUPDUPData)
		{
			CString txt = L"Creating inline image files ...";
			ResHelper::TranslateString(txt);
			MboxMail::pCUPDUPData->SetProgress(txt, 0);
		}
	}
	int ni = MboxMail::s_mails.GetSize();

	BOOL terminated = FALSE;
	CString itemNumberStr;
	std::vector <MailBodyContent*>::iterator it;


	if (!mainThread)
	{
		if (MboxMail::pCUPDUPData)
		{
			CString txt = L"Index file created. Creating inline image files ...";
			ResHelper::TranslateString(txt);
			MboxMail::pCUPDUPData->SetProgress(txt, 0);
		}
	}

	ULONGLONG workRangeFirstPos = 0;
	ULONGLONG workRangeLastPos = ni - 1;
	ProgressTimer progressTimer(workRangeFirstPos, workRangeLastPos);

	MboxMail *m;
	for (int i = 0; i < ni; i++)
	{
		m = MboxMail::s_mails[i];

		int mailPosition = i;
		BOOL createEmbeddedImageFiles = TRUE;
		NListView::CreateEmbeddedImageFilesEx(fpm, mailPosition, imageCachePath, createEmbeddedImageFiles);  // FIXME maybe ??

		if (!mainThread && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				terminated = TRUE;
				break;
			}

			UINT_PTR dwProgressbarPos = 0;
			ULONGLONG workRangePos = i;
			BOOL needToUpdateStatusBar = progressTimer.UpdateWorkPos(workRangePos, dwProgressbarPos);
			if (needToUpdateStatusBar)
			{
				int nItem = i + 1;

				CString fmt = L"Creating inline image files ... %d of %d";
				ResHelper::TranslateString(fmt);
				itemNumberStr.Format(fmt, nItem, ni);

				if (MboxMail::pCUPDUPData)  MboxMail::pCUPDUPData->SetProgress(itemNumberStr, (UINT_PTR)(dwProgressbarPos));

				int debug = 1;
			}
		}

		if (terminated)
		{
			errorText = L"Already created image files will not be deleted.";
		}
	}

	fpm.Close();
	return TRUE;
}

int NListView::PrintMailSelectedToSingleTEXT_Thread(CString &targetPrintSubFolderName, CString &targetPrintFolderPath)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS args;

	CString rootPrintSubFolder = L"PrintCache";
	CString targetPrintSubFolder = targetPrintSubFolderName;
	CString printCachePath;
	CString errorText;

	MailIndexList *selectedMailsIndexList = 0;
	if (PopulateSelectedMailsList() > 0)
		selectedMailsIndexList = &m_selectedMailsList;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, printCachePath, errorText);
	if (retval == FALSE) {
		HWND h = GetSafeHwnd(); // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	/*IN*/
	args.lview = this;
	args.targetPrintFolderPath = printCachePath;
	args.targetPrintSubFolderName = targetPrintSubFolder;
	args.firstMail = -1;
	args.lastMail = -1;
	args.selectedMailIndexList = selectedMailsIndexList;
	args.nItem = -1;
	args.separatePDFs = FALSE;

	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcPrintMailGroupToSingleTEXT, (LPVOID)(PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS*)&args);
	Dlg.SetDialogTemplate(AfxGetApp()->m_hInstance, MAKEINTRESOURCE(IDD_PROGRESS_DLG), IDC_STATIC, IDC_PROGRESS_BAR, IDCANCEL);

	INT_PTR nResult = Dlg.DoModal();

	if (!nResult) { // should never be true ?
		MboxMail::assert_unexpected();
		return -1;
	}

	int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
	int retResult = LOWORD(nResult);

	if (retResult != IDOK)
	{  // IDOK==1, IDCANCEL==2
		// We should be here when user selects Cancel button
		//_ASSERTE(cancelledbyUser == TRUE);

		DWORD terminationDelay = Dlg.GetTerminationDelay();
		int loopCnt = (terminationDelay+100)/25;

		ULONGLONG tc_start = GetTickCount64();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		ULONGLONG tc_end = GetTickCount64();
		DWORD delta = (DWORD)(tc_end - tc_start);
		TRACE(L"(PrintMailSelectedToSingleTEXT_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret > 0)
	{
		if (args.errorText.IsEmpty())
		{
			int textType = 0;
			bool fileExists = false;
			CString fileName;

			//int ret = MboxMail::MakeFileNameFromMailArchiveName(textType, fileName, targetPrintSubFolder, fileExists, errorText);

			// FIXME determine the file name in the root caller to avoid inconsistency
			// It is ridicules now
			if (selectedMailsIndexList->GetCount() > 1)
				ret = MboxMail::MakeFileNameFromMailArchiveName(textType, fileName, targetPrintSubFolder, fileExists, errorText);
			else
				ret = MboxMail::MakeFileNameFromMailHeader(selectedMailsIndexList->GetAt(0), textType, fileName, targetPrintSubFolder, fileExists, errorText);

			//CString txt = L"Created TEXT file \n\n" + fileName;
			CString txt = L"Created TEXT file";
			ResHelper::TranslateString(txt);

			OpenContainingFolderDlg dlg(txt, fileName, FALSE);
			INT_PTR nResponse = dlg.DoModal();
			////////////
			if (nResponse == IDOK)
			{
				if (FileUtils::BrowseToFile(fileName) == FALSE) {
					HWND h = GetSafeHwnd();
					HINSTANCE result = ShellExecute(h, L"open", printCachePath, NULL, NULL, SW_SHOWNORMAL);
					CMainFrame::CheckShellExecuteResult(result, h);
				}
				int deb = 1;
			}
			else if (nResponse == IDYES)
			{
				HWND h = GetSafeHwnd();
				HINSTANCE result = ShellExecute(h, L"open", fileName, NULL, NULL, SW_SHOWNORMAL);
				CMainFrame::CheckShellExecuteResult(result, h, &fileName);
				int deb = 1;
			}
			else if (nResponse == IDCANCEL)
			{
				int deb = 1;
			}

		}
		else
		{
			MboxMail::assert_unexpected();
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}
	return 1;
}

int NListView::PrintMailSelectedToSingleTEXT_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, 
	CString &targetPrintFolderPath, CString &errorText)
{
	//int retdel = NListView::DeleteAllHtmAndPDFFiles(targetPrintFolderPath);

	if (selectedMailIndexList->GetCount() <= 0)
		return 1;

	TEXTFILE_CONFIG textConfig;
	CString textFileName;
	int textType = 0;

	textConfig.m_dateFormat = m_format;
	textConfig.m_bGMTTime = m_gmtTime;
	textConfig.m_nCodePageId = CP_UTF8;

	int ret = MboxMail::PrintMailSelectedToSingleTextFile_WorkerThread(textConfig, textFileName, selectedMailIndexList, textType, errorText);

	return 1;
}
//  FIXME what is remapDuplicateNames; why not remapp duplicate names and add to attachmentDB always ??
int NListView::DetermineAttachmentName(CFile *fpm, int mailPosition, MailBodyContent *body, SimpleString *bodyData, 
	CStringW &nameW, AttachmentMgr &attachmentDB, BOOL remapDuplicateNames, BOOL extraValidation)
{
	DBGT(__FUNCTION__);

	// This is based on NListView::SelectItem()

	CStringW contentSubtypeW;
	CStringA contentSubtypeA;
	CStringW contentTypeW;
	CStringA contentTypeA;
	//
	CString standardDocumentExtensionA;
	CStringW contentTypeMainW;
	//
	CString cStrNameW;
	CStringA nameExtensionA;
	CString nameExtensionW;
	CStringA nameBaseA;
	CStringA nameBaseW;
	//
	BOOL alloc_tmpbuf = FALSE;
	DWORD error;

	MboxMail* m = MboxMail::s_mails[mailPosition];

	CStringA cStrNameA = body->m_attachmentName;
	UINT nameCharsetId = body->m_attachmentNamePageCode;
	if (nameCharsetId == 0)
		nameCharsetId = CP_UTF8;

	if (TextUtilsEx::Str2WStr(cStrNameA, nameCharsetId, cStrNameW, error))
	{
		int deb = 1;
	}

#if 0
	if (cStrNameW.Find(L"XXX") >= 0)
		const int deb = 1;
#endif

	TRACE(L"DetermineAttachmentName: body->m_attachmentName=\"%s\"\n", cStrNameW);

	int pos = body->m_contentType.ReverseFind('/');
	if (pos > 0)
	{
		contentSubtypeA = body->m_contentType.Mid(pos + 1);
		contentTypeA = body->m_contentType.Left(pos);
		contentSubtypeW = contentSubtypeA;
		contentTypeW = contentTypeA;
	}

	// Check first if SubType is "octet-stream" and data is one of the image types
	CStringW imageExtensionW;
	CStringA imageExtensionA;
	BOOL isOctetStream = FALSE;
	if (contentSubtypeA.CompareNoCase("octet-stream") == 0)
		isOctetStream = TRUE;

	if (isOctetStream && !cStrNameW.IsEmpty())
	{
		isOctetStream = TRUE;
		if ((bodyData == 0) && (fpm != 0))
		{
			bodyData = MboxMail::get_tmpbuf();
			alloc_tmpbuf = TRUE;
			bodyData->ClearAndResize(body->m_contentLength * 2);

			int retLen = MboxMail::DecodeBody(*fpm, body, mailPosition, bodyData);
			if (bodyData->Count() > 1500000)
				int deb = 1;
		}

		if ((bodyData != 0) && (fpm != 0))
		{
			const unsigned char* data = (unsigned char*)bodyData->Data();
			int dataLength = bodyData->Count();

			BOOL ret = GdiUtils::loadImage((BYTE*)data, (size_t)dataLength, imageExtensionW, imageExtensionA);
			if (ret == FALSE)
			{
				imageExtensionW.Empty();
				imageExtensionA.Empty();
			}
		}
	}

	// Validate assumption.Should be true most of the time
	if (isOctetStream && !cStrNameW.IsEmpty())
		_ASSERTE(TRUE);

	if (!cStrNameW.IsEmpty())
	{
		//_ASSERTE(isOctetStream == FALSE);  // FIXME do we need to handle isOctetStream here

		int pos = cStrNameW.ReverseFind('.');
		if (pos > 0)
		{
			nameExtensionW = cStrNameW.Mid(pos+1);
			nameBaseW = cStrNameW.Left(pos);
		}

		if (nameExtensionW.IsEmpty())
		{
			if (!imageExtensionW.IsEmpty())  // Not likely scenario but must check
			{
				cStrNameW += imageExtensionW;
			}
			else
			{
				CString standardDocumentExtensionA;
				HtmlUtils::CommonMimeType2DocumentTypes(body->m_contentType, standardDocumentExtensionA);
				CString standardDocumentExtensionW = standardDocumentExtensionA;

				if (!standardDocumentExtensionW.IsEmpty())
				{
					cStrNameW += standardDocumentExtensionW;
				}
				else
					cStrNameW += L"." + contentSubtypeW;
			}
		}
		else
		{
			; // just return cStrNameW
		}
	}
	else  // cStrNameW.IsEmpty()
	{
		CString standardDocumentExtensionA;
		HtmlUtils::CommonMimeType2DocumentTypes(body->m_contentType, standardDocumentExtensionA);
		CString standardDocumentExtensionW = standardDocumentExtensionA;
		if ((isOctetStream == TRUE) && !imageExtensionW.IsEmpty())
		{
			cStrNameW = contentTypeW + L"." + contentSubtypeW + imageExtensionW;
		}
		else
		{
			cStrNameW = contentTypeW + L"." + contentSubtypeW + standardDocumentExtensionW;
		}
	}

	if (alloc_tmpbuf)  // this is a big mess
		MboxMail::rel_tmpbuf();

	if (cStrNameW.IsEmpty())
		_ASSERTE(FALSE);

	CString validNameW;
	BOOL bReplaceWhiteWithUnderscore = FALSE;
	//BOOL extraValidation = TRUE;  // FIXME
	FileUtils::MakeValidFileName(cStrNameW, validNameW, bReplaceWhiteWithUnderscore, extraValidation);
	if (validNameW.Compare(cStrNameW))
		int deb = 1;

	cStrNameW = validNameW;

	// FIXME: below should be moved to the caller to simpify
	// attachmentDB needs to be enhanced to compare lengtr and CRC ? of attachments
	BOOL isEmbedded = FALSE;
	int prefix = attachmentDB.InsertName(cStrNameW, isEmbedded);   // returns unique id/prefix to make name unique
	if (prefix >= 0)
	{
		CStringW nameW;
		//nameW.Format(L"%d_%s", prefix, cStrNameW);

		CString fileBaseName;
		CString fileNameExtention;
		FileUtils::GetFileBaseNameAndExtension(cStrNameW, fileBaseName, fileNameExtention);

		nameW.Format(L"%s%s%03d%s", fileBaseName, L"_", prefix, fileNameExtention);

		cStrNameW = nameW;
		if (remapDuplicateNames)  // FIXME why we would not add always but returned unique name always
		{
			BOOL isEmbedded = FALSE;
			attachmentDB.AddName(cStrNameW, isEmbedded);
		}
	}

	nameW = cStrNameW;

	TRACE(L"DetermineAttachmentName: returned attachmentName=\"%s\"  \n", nameW);

	return 1;
}
//////////////

// Creates mail attachments files except embeded image attachments
//
// Not used currently FIXMEFIXME
int NListView::PrintMailAttachments(CFile *fpm, int mailPosition, AttachmentMgr &attachmentDB)
{
	_ASSERTE(FALSE);

	if ((mailPosition >= MboxMail::s_mails.GetCount()) || (mailPosition < 0))
		return -1;


	MboxMail *m = MboxMail::s_mails[mailPosition];
	if (m == 0)
	{
		// TODO: Assert ??
		return 1;
	}

	CFile mboxFp;
	CFile *fpm_save = fpm;
	if (fpm == 0)
	{
		CFileException ExError;
		if (!mboxFp.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
		{
			// TODO: critical failure
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt = L"Could not open \"" + MboxMail::s_path;
			txt += L"\" mail file.\n";
			txt += exErrorStr;

			TRACE(L"%s\n", txt);

			return FALSE;
		}
		fpm = &mboxFp;
	}


	BOOL bReplaceWhiteWithUnderscore = FALSE;
	BOOL ret;

	CStringW resultW;
	CStringW validNameW;
	CStringW fileNameW;

	MailBodyContent *body;

	CString errorText;
	CString attachmentCachePath;
	CStringW printCachePathW;
	CString rootPrintSubFolder = L"AttachmentCache";
	CString targetPrintSubFolder;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, attachmentCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		// TODO: what to do ?
		//int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		//return -1;
	}

	//ret = TextUtilsEx::Ansi2WStr(attachmentCachePath, printCachePathW, error);   // FIXME delete or document

	printCachePathW = attachmentCachePath;

	if (!m->m_DetermineEmbeddedImagesDone)
	{
		// Sets body->m_isEmbeddedImage
		ret = MboxMail::DetermineEmbeddedImages(mailPosition, *fpm);
	}

	bool showAllAttachments = false;
	AttachmentConfigParams *attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
	if (attachmentConfigParams)
	{
		showAllAttachments = attachmentConfigParams->m_bShowAllAttachments_Window;
	}

	attachmentDB.Clear();

	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];

		BOOL showAttachment = FALSE;
		if (body->IsAttachment() || body->IsInlineAttachment())
		{
			if (showAllAttachments || !body->m_isEmbeddedImage)
				showAttachment = TRUE;
		}

		if (showAttachment)
		{
			UINT inCodePage = body->m_attachmentNamePageCode;
			resultW.Empty();

			CString fileName;
			NListView::GetExtendedMailId(m, fileName);

			CStringW fileNameW;

			//if (TextUtilsEx::Ansi2WStr(fileName, fileNameW, error))  // FIXME delete or document
			{
				int deb = 1;
			}

			fileNameW = fileName;
			// make attachment name

			SimpleString*outbuf = MboxMail::get_tmpbuf();
			outbuf->ClearAndResize(body->m_contentLength * 2);

			int retLen = MboxMail::DecodeBody(*fpm, body, mailPosition, outbuf);
			if (outbuf->Count() > 1500000)
				int deb = 1;

			// FIXME remapDuplicateNames = TRUE implies that duplicate names are remapped and added into attachmentDB
			BOOL remapDuplicateNames = TRUE;
			BOOL extraValidation = TRUE;
			DetermineAttachmentName(fpm, mailPosition, body, outbuf, resultW, attachmentDB, remapDuplicateNames, extraValidation);

			fileNameW.Append(resultW);

			validNameW.Empty();
			validNameW.Append(resultW);  // FIXME

			//FileUtils::MakeValidFileName(fileNameW, validNameW, bReplaceWhiteWithUnderscore);

			CStringW filePathW = printCachePathW + L"\\" + validNameW;

			const char *data = outbuf->Data();
			int dataLength = outbuf->Count();

			ret = FileUtils::Write2File(filePathW, (unsigned char*)data, dataLength);

			MboxMail::rel_tmpbuf();
		}
	}

	if (fpm_save == 0)
		mboxFp.Close();

	return 1;
}

//
// Creates mail attachments files except embeded image attachments
//
int NListView::PrintAsEmlFile(CFile *fpm, int mailPosition, CString &emlFile)
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	NamePatternParams *pNamePP = &pFrame->m_NamePatternParams;

	if ((mailPosition >= MboxMail::s_mails.GetCount()) || (mailPosition < 0))
		return -1;

	MboxMail *m = MboxMail::s_mails[mailPosition];
	if (m == 0)
	{
		// TODO: Assert ??
		return 1;
	}

	CFile mboxFp;
	CFile *fpm_save = fpm;
	if (fpm == 0)
	{
		CFileException ExError;
		if (!mboxFp.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
		{
			// TODO: critical failure
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt = L"Could not open \"" + MboxMail::s_path;
			txt += L"\" mail file.\n";
			txt += exErrorStr;

			TRACE(L"%s\n", txt);
			return FALSE;
		}
		fpm = &mboxFp;
	}

	CStringA mailFileNameBaseA;
	CString mailFileNameBase;
	int retMakeFileName;
	if (!pFrame->m_NamePatternParams.m_bCustomFormat)
		retMakeFileName = MboxMail::MakeFileName(m, &pFrame->m_NamePatternParams, mailFileNameBase, mailFileNameBaseA);
	else
		retMakeFileName = MboxMail::MakeFileName(m, &pFrame->m_NamePatternParams, pFrame->m_NamePatternParams.m_nameTemplateCnf, mailFileNameBase, pFrame->m_NamePatternParams.m_nFileNameFormatSizeLimit, mailFileNameBaseA);

	CString emlCachePath;
	CString rootPrintSubFolder = L"EmlCache";
	CString errorText;
	CString targetPrintSubFolder;
	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, emlCachePath, errorText);
	if (retval == FALSE) {
		return -1;
	}

	int emlCachePathLength = emlCachePath.GetLength();
	int mailFileNameBaseLength = mailFileNameBase.GetLength();
	int total = emlCachePathLength + mailFileNameBaseLength;
	CString fileName = emlCachePath + L"\\" + mailFileNameBase + L".eml";
	int fileNamePathLength = fileName.GetLength();
	if (fileNamePathLength > _MAX_FNAME)
		int deb = 1;

	SimpleString *outbuf = MboxMail::m_outbuf;
	outbuf->ClearAndResize(10000);

	CFile fp;
	CFileException ExError;
	if (!fp.Open(fileName, CFile::modeWrite | CFile::modeCreate, &ExError))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);
#if 0
		CString txt = L"Could not create \"" + fileName;
		txt += L"\" eml file.\n";
		txt += exErrorStr;
#endif
		CString txt;
		CString fmt = L"Could not create \"%s\" eml file.\n%s";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, fileName, exErrorStr);

		// PrintAsEmlFile is the static function. Must use global MessageBox
		// Not ideal because program is not blocked. TODO: invetsigate and change 
		HWND h = NULL;
		int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		// continue for now ??
		return -1;
	}

	BOOL ret;
	outbuf->Clear();
	ret = m->GetBodySS(outbuf);
	// Remove From line at the begining of mail body

	static const char *cFromMailBegin = "From ";
	static const int cFromMailBeginLen = istrlen(cFromMailBegin);

	char *p = outbuf->Data();
	int bodyCnt = outbuf->Count();
	char *e = p + bodyCnt;
	p = MimeParser::SkipEmptyLines(p, e);
	if (TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0)
	{
		// Can From line be splitted ?
		//p = MimeParser::GetMultiLine(p, e, line);
		p = MimeParser::EatNewLine(p, e);
		bodyCnt -= IntPtr2Int(p - outbuf->Data());
		if (bodyCnt < 0)
			bodyCnt = 0;
	}

	fp.Write(p, bodyCnt);
	fp.Close();

	if (fpm_save == 0)
		mboxFp.Close();

	emlFile = fileName;

	return 1;
}

int NListView::ExportAsEmlFile(CFile *fpm, int mailPosition, CString &targetDirectory, CString &emlFile, CString &errorText)
{
	if ((mailPosition >= MboxMail::s_mails.GetCount()) || (mailPosition < 0))
		return -1;

	MboxMail *m = MboxMail::s_mails[mailPosition];
	if (m == 0)
	{
		errorText = L"Internal error, try again and see if that help, otherwise restart application";
		return -1;
	}

	CFile mboxFp;
	CFile *fpm_save = fpm;
	if (fpm == 0)
	{
		CFileException ExError;
		if (!mboxFp.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
		{
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt = L"Could not open \"" + MboxMail::s_path;
			txt += L"\" mail file.\n";
			txt += exErrorStr;

			TRACE(L"%s\n", txt);

			errorText = txt;
			return -1;
		}
		fpm = &mboxFp;
	}

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	CStringA mailFileNameBaseA;
	CString mailFileNameBase;
	int retMakeFileName;
	if (!pFrame->m_NamePatternParams.m_bCustomFormat)
		retMakeFileName = MboxMail::MakeFileName(m, &pFrame->m_NamePatternParams, mailFileNameBase, mailFileNameBaseA);
	else
		retMakeFileName = MboxMail::MakeFileName(m, &pFrame->m_NamePatternParams, pFrame->m_NamePatternParams.m_nameTemplateCnf, mailFileNameBase, pFrame->m_NamePatternParams.m_nFileNameFormatSizeLimit, mailFileNameBaseA);

	CString fileName = targetDirectory + mailFileNameBase + L".eml";

	SimpleString *outbuf = MboxMail::m_outbuf;
	outbuf->ClearAndResize(10000);

	CFile fp;
	CFileException ExError;
	if (!fp.Open(fileName, CFile::modeWrite | CFile::modeCreate, &ExError))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not create \"" + fileName;
		txt += L"\" file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);

		errorText = txt;
#if 0
		// ExportAsEmlFile is the static function. Must use global MessageBox
		// It can be called from non main thread and need to figure whether calling
		// :: MessageBox is valid. Note: it seems to work
		// Not ideal because program is not blocked. TODO: invetsigate and change 

		HWND h = NULL;
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
#endif
		// continue for now.
		return -1;
	}

	BOOL ret;
	outbuf->Clear();
	ret = m->GetBodySS(outbuf);
	// Remove From line at the begining of mail body

	static const char *cFromMailBegin = "From ";
	static const int cFromMailBeginLen = istrlen(cFromMailBegin);

	char *p = outbuf->Data();
	int bodyCnt = outbuf->Count();
	char *e = p + bodyCnt;
	p = MimeParser::SkipEmptyLines(p, e);
	if (TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0)
	{
		// Can From line be splitted ?
		//p = MimeParser::GetMultiLine(p, e, line);
		p = MimeParser::EatNewLine(p, e);
		bodyCnt -= IntPtr2Int(p - outbuf->Data());
		if (bodyCnt < 0)
			bodyCnt = 0;
	}

	fp.Write(p, bodyCnt);
	fp.Close();

	if (fpm_save == 0)
		mboxFp.Close();

	emlFile = fileName;

	return 1;
}

int NListView::PrintAttachmentNamesAsText2CSV(int mailPosition, SimpleString *outbuf, CString &characterLimit, CStringA &attachmentSeparator)
{
	if ((mailPosition >= MboxMail::s_mails.GetCount()) || (mailPosition < 0))
		return -1;

	MboxMail *m = MboxMail::s_mails[mailPosition];
	if (m == 0)
	{
		// TODO: Assert ??
		return 1;
	}

	int limit = _ttoi(characterLimit);
	limit = 32500;  // hardcode for now  (2^^15 - 1)  is (32768 - 1) is supported max by Excel

	SimpleString tmpbuf(1024, 256);
	SimpleString tmpbuf2(1024, 256);
	SimpleString resultUTF8(1024, 256);
	CStringW nameW;
	CStringW validNameW;
	DWORD error;
	BOOL bReplaceWhiteWithUnderscore = FALSE;

	MailBodyContent *body;

	int begCount = outbuf->Count();

	outbuf->Append('"');

	limit -= 2; // two " enclosing chracters

	bool showAllAttachments = false;
	AttachmentConfigParams *attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
	if (attachmentConfigParams)
	{
		showAllAttachments = attachmentConfigParams->m_bShowAllAttachments_Window;
	}

	AttachmentMgr attachmentDB;
	int attachmentCnt = 0;
	CFile *fpm = 0;
	SimpleString *bodyData = 0;


	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];

		BOOL showAttachment = FALSE;
		if (body->IsAttachment() || body->IsInlineAttachment())
		{
			if (showAllAttachments || !body->m_isEmbeddedImage)
				showAttachment = TRUE;
		}

		if (showAttachment)
		{
			UINT inCodePage = body->m_attachmentNamePageCode;

			SimpleString*buf = 0;
			// FIXME remapDuplicateNames = FALSE implies that duplicate names will not be added into attachmentDB
			// attachmentDB is local and is not used layer in this function,  required by DetermineAttachmentName
			BOOL remapDuplicateNames = FALSE;
			BOOL extraValidation = TRUE;  // FIXME
			NListView::DetermineAttachmentName(fpm, mailPosition, body, buf, nameW, attachmentDB, remapDuplicateNames, extraValidation);

			validNameW.Empty();
			FileUtils::MakeValidFileName(nameW, validNameW, bReplaceWhiteWithUnderscore, extraValidation);
			validNameW = nameW; // FIXMEFIXME

			UINT outCodePage = CP_UTF8;
			BOOL ret2 = TextUtilsEx::WStr2CodePage((wchar_t*)(LPCWSTR)validNameW, validNameW.GetLength(), outCodePage, &resultUTF8, error);

			int namelen = resultUTF8.Count();
			tmpbuf2.ClearAndResize(namelen + 10);

			begCount = outbuf->Count();

			if (attachmentCnt)
			{
				CStringA fmt = attachmentSeparator;
				fmt.Replace("\\r\\n", "\r\n");
				fmt.Replace("CRLF", "\r\n");
				//fmt.Replace("CR", "\r");
				//fmt.Replace("LF", "\n");
				tmpbuf2.Append(fmt, fmt.GetLength());
			}

			tmpbuf2.Append('"');
			tmpbuf2.Append(resultUTF8.Data(), resultUTF8.Count());
			tmpbuf2.Append('"');

			tmpbuf.ClearAndResize(3 * tmpbuf2.Count());
			int ret_addrlen = MboxMail::escapeSeparators(tmpbuf.Data(), tmpbuf2.Data(), tmpbuf2.Count(), '"');
			tmpbuf.SetCount(ret_addrlen);

			attachmentCnt++;

			outbuf->Append(tmpbuf);
			if (outbuf->Count() > limit)
			{
				outbuf->SetCount(begCount); // reverse count and return
				break;
			}
		}
	}
	outbuf->Append('"');

	return 1;
}

int NListView::FindAttachmentName(MboxMail *m, CStringA &searchString, BOOL bWholeWord, BOOL bCaseSensitive)
{
	if (m == 0)
	{
		_ASSERTE(FALSE);
		return 0;
	}

	MailBodyContent *body;

	CString contentTypeExtension;
	CString contentTypeMain;

	int searchStringPos = 0;

	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];

		contentTypeExtension.Empty();
		contentTypeMain.Empty();

		int pos = body->m_contentType.ReverseFind('/');
		if (pos > 0)
		{
			contentTypeExtension = body->m_contentType.Mid(pos + 1);
			contentTypeMain = body->m_contentType.Left(pos);
		}

		if (!body->m_attachmentName.IsEmpty())
		{
			UINT inCodePage = body->m_attachmentNamePageCode;
			if ((body->m_contentDisposition.CompareNoCase("attachment") == 0))
			{
				if (searchString.GetLength() == 1 && searchString[0] == '*')
					return 1;

				if (bWholeWord) {
					searchStringPos = g_tu.StrSearchW((unsigned char *)(LPCSTR)body->m_attachmentName, body->m_attachmentName.GetLength(), inCodePage,
						(unsigned char *)(LPCSTR)searchString, searchString.GetLength(), bCaseSensitive, 0);
				}
				else {
					searchStringPos = g_tu.StrSearch((unsigned char *)(LPCSTR)body->m_attachmentName, body->m_attachmentName.GetLength(), inCodePage,
						(unsigned char *)(LPCSTR)searchString, searchString.GetLength(), bCaseSensitive, 0);
				}
				if (searchStringPos >= 0) {
					return 1;
				}
			}
		}
	}
	return 0;
}

// Below doesn't work in some cases when the text is white on non-withe background.
// Fix is more complex and requires to build DOM, make chnages and get the HTML text back.\
// Can't do this today -:), will reexamine later.
// Removing backgroound color works  when printing to PDF printer and when using wkhtmltopdf tool to print to PDF file.
// Headless Chrome doesn't support removing background color and below was an attempt to support Chrome.
int NListView::RemoveBackgroundColor(char *inData, int indDataLen, SimpleString *outbuf, int mailPosition)
{
	static char *bgColorPat1 = "bgcolor=";
	static int bgColorPat1_len = istrlen(bgColorPat1);
	static char *bgColorPat2 = "background-color:";
	static int bgColorPat2_len = istrlen(bgColorPat2);


	char *input = inData;
	int inputLength = indDataLen;
	char *inputEnd = input + inputLength;

	char *pos = input;
	char c;
	char *c_pos_last = pos;
	while ((pos != 0) && (pos < inputEnd))
	{
		c = *pos;
		if ((c == 'b') || (c == 'B'))
		{
			if (TextUtilsEx::strncmpUpper2Lower(pos, inputEnd, bgColorPat1, bgColorPat1_len) == 0)
			{
				// find ending space or " character
				char *c_pos_color = pos + bgColorPat1_len;
				c_pos_color = TextUtilsEx::SkipWhitePlus(c_pos_color);
				char *c_pos_end = TextUtilsEx::findOneOf(c_pos_color, inputEnd, " \"");
				if (c_pos_end)
				{
					int lenData = IntPtr2Int(c_pos_color - c_pos_last);
					outbuf->Append(c_pos_last, lenData);

					CStringA color;
					int lenColor = IntPtr2Int(c_pos_end - c_pos_color);
					color.Append(c_pos_color, lenColor);
					color.Trim();
					color = "#FFFFFF";
					outbuf->Append(color, color.GetLength());

					c_pos_last = c_pos_end;
					pos = c_pos_end;

					int deb = 1;
				}
			}
			else
			{
				if (TextUtilsEx::strncmpUpper2Lower(pos, inputEnd, bgColorPat2, bgColorPat2_len) == 0)
				{
					// find ending ; or " character
					char *c_pos_color = pos + bgColorPat2_len;
					c_pos_color = TextUtilsEx::SkipWhitePlus(c_pos_color);
					char *c_pos_end = TextUtilsEx::findOneOf(c_pos_color, inputEnd, ";\"");
					if (c_pos_end)
					{
						int lenData = IntPtr2Int(c_pos_color - c_pos_last);
						outbuf->Append(c_pos_last, lenData);

						CStringA color;
						int lenColor = IntPtr2Int(c_pos_end - c_pos_color);
						color.Append(c_pos_color, lenColor);
						color.Trim();
						color = "#FFFFFF";
						outbuf->Append(color, color.GetLength());

						c_pos_last = c_pos_end;
						pos = c_pos_end;

						int deb = 1;
					}
				}
			}
			int deb = 1;
		}
		pos++;
	}
	if (c_pos_last != inputEnd)
	{
		int lenData = IntPtr2Int(inputEnd - c_pos_last);
		outbuf->Append(c_pos_last, lenData);
	}
	return 1;
}

int NListView::CreateEmptyFolder(CString &driveName, CString &mboxDirectory, CString &mboxFolderName, CString &parentSubFolderPath, CString &newFolderName)
{
#if 0   // FIXME
	CString path;
	path.Append(driveName);
	path.Append(mboxDirectory);
	path.Append("Folders\\");
	//path.Append(mboxFolderName);
	if (!parentSubFolderPath.IsEmpty())
	{
		path.Append("\\");
		path.Append(parentSubFolderPath);
	}
	path.Append("\\");
	path.Append(newFolderName);

#if 1
	CStringW wpath(path);
	int ret = SHCreateDirectory(0, wpath);
	
#else
	int ret = SHCreateDirectoryEx(0, path, 0);  // requires that files are visible
#endif
	if (ret != ERROR_SUCCESS)
	{
#if 0
		if (ret == ERROR_FILENAME_EXCED_RANGE) {
			return ret;
		}
		else
#endif
			if ((ret == ERROR_ALREADY_EXISTS) || (ret == ERROR_FILE_EXISTS))
				int deb = 1;
			else
				return ret;
	}

	int ret2 = CreateEmptyFolderListFile(path, newFolderName);
#endif
	return ERROR_SUCCESS;
}

int NListView::CreateEmptyFolderListFile(CString &path, CString &folderNameFile)
{
#if 0
	MailArray *mailsArray = &MboxMail::s_mails;

	CString mboxFileSuffix;
	if (MboxMail::IsUserMailsSelected())
		mboxFileSuffix = "_USER.mbox";
	else if (MboxMail::IsFindMailsSelected())
		mboxFileSuffix = "_FIND.mbox";
	else
	{
		// We should never be here
		return -1;
	}

	MboxMail *m;

	CString mailFile = MboxMail::s_path;

	if (MboxMail::s_path.IsEmpty()) {
		CString txt = L"Please open mail file first.";
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	CString path = MboxMail::GetLastPath();
	if (path.IsEmpty())
		return -1;  // Hopefully s_path wil fail first

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(mailFile, driveName, directory, fileNameBase, fileNameExtention);

	CString errorText;
	CString printCachePath;
	CString rootPrintSubFolder = "ArchiveCache";
	CString targetPrintSubFolder;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, printCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CFile fp;
	CString mboxFile = fileNameBase + fileNameExtention;
	CString mboxFilePath = printCachePath + "\\" + fileNameBase + mboxFileSuffix;

	CString mboxFileListSuffix = ".mboxlist";
	CString mboxListFile = mboxFilePath + mboxFileListSuffix;

	BOOL ret2 = FALSE;
	int ret1 = 0;
	int ret0 = 0;

	if (FileUtils::PathFileExist(mboxListFile))
	{
		CString mboxListFileBak1 = mboxListFile + ".bak1";
		CString mboxListFileBak2 = mboxListFile + ".bak2";

		if (FileUtils::PathFileExist(mboxListFileBak2))
			ret2 = DeleteFile(mboxListFileBak2);

		if (FileUtils::PathFileExist(mboxListFileBak1))
			ret1 = rename(mboxListFileBak1, mboxListFileBak2);

		ret0 = rename(mboxListFile, mboxListFileBak1);
	}
#endif

	CString mboxListFile = path + L"\\" + folderNameFile + L".mbox.mboxlist";

	SerializerHelper sz(mboxListFile);
	if (!sz.open(TRUE))
	{
#if 0
		CString txt = L"Could not create \"" + mboxListFile;
		txt += L"\" file.\nMake sure file is not open on other applications.";
#endif

		CString txt;
		CString fmt = L"Could not create \"%s\" file.\nMake sure file is not open on other applications.";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, mboxListFile);

		int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	// Create mboxlist file to allow reload of archive file list

	sz.writeInt(MAIL_LIST_VERSION);			// version
	//sz.writeString(mboxFile);  // TODO: ??
	sz.writeInt64(MboxMail::s_fSize);	// root mail file size
	_int64 mailFileSize = 0; //  -1;  //  FileSize(mboxFilePath);  // we don't always create new mail archive and mail list
	sz.writeInt64(mailFileSize);	// file size
	sz.writeInt(0);

	sz.close();

	MailList *mailFolder = new MailList;
	mailFolder->m_nId = IDC_FOLDER_LIST;
	mailFolder->m_path = path;
	mailFolder->m_folderName = folderNameFile + L".mbox.mboxlist";

	MboxMail::m_folderList.insert_tail(mailFolder);

	return 1;
}

// FIXME is this used ?? This is using old parser, no longer used. Delete ??
// Update  NListView::ScanAllMailsInMbox_NewParser() instead
// Actually NListView::UpdateInlineSrcImgPathEx updates these stats so review alol and decide what to delete
int NListView::ScanAllMailsInMbox()
{
	ULONGLONG tc_start = GetTickCount64();

	MboxMail::m_EmbededImagesNoMatch = 0;

	MboxMail::m_EmbededImagesFoundMHtml = 0;
	MboxMail::m_EmbededImagesFoundMHtmlHtml = 0;
	MboxMail::m_EmbededImagesFoundUnexpectedMHtml = 0;

	MboxMail::m_EmbededImagesFound = 0;
	MboxMail::m_EmbededImagesFoundCid = 0;
	MboxMail::m_EmbededImagesFoundHttp = 0;
	MboxMail::m_EmbededImagesFoundHttps = 0;
	MboxMail::m_EmbededImagesFoundMHtmlHttp = 0;
	MboxMail::m_EmbededImagesFoundMHtmlHttps = 0;
	MboxMail::m_EmbededImagesFoundData = 0;
	MboxMail::m_EmbededImagesFoundLocalFile = 0;
	//
	MboxMail::m_EmbededImagesNotFound = 0;
	MboxMail::m_EmbededImagesNotFoundCid = 0;
	MboxMail::m_EmbededImagesNotFoundHttp = 0;
	MboxMail::m_EmbededImagesNotFoundHttps = 0;
	MboxMail::m_EmbededImagesNotFoundMHtmlHttp = 0;
	MboxMail::m_EmbededImagesNotFoundMHtmlHttps = 0;
	MboxMail::m_EmbededImagesNotFoundData = 0;
	MboxMail::m_EmbededImagesNotFoundLocalFile = 0;

	TRACE(L"Mbox=%s\n\n", MboxMail::s_path);

	bool runInvestigation = true;
	int iItem;
	BOOL ret;
	for (iItem = 0; iItem < MboxMail::s_mails.GetCount(); iItem++)
	{
		// Get cached mail
		MboxMail *m = MboxMail::s_mails[iItem];

		// Get raw mail body
		CStringA bdy;
		ret = m->GetBody(bdy);

		// Decode MIME message
		CMimeMessage mail;
		const char *bodyData = bdy;
		int nLoadedSize = mail.Load(bodyData, bdy, bdy.GetLength());

		bool hasInlineAttachments = false;

		// Iterate all the descendant body parts
		CMimeBody::CBodyList bodies;
		int nCount = mail.GetBodyPartList(bodies);

		int mailPosition = iItem;
		MailBodyInfoArray cidArray;
		MyCArray<bool> fileImgAlreadyCreatedArray;

		// CreateInlineImageFiles_SelectedItem will create files for embeded images and mark each body as embeded or not in cidArray
		NMsgView *pMsgView = 0;

		// FIXME int ret = NListView::CreateInlineImageFiles_SelectedItem(bodies, pMsgView, mailPosition, cidArray, fileImgAlreadyCreatedArray, runInvestigation);
		if (cidArray.GetCount())
			hasInlineAttachments = true;

		MailBodyInfo* info;
		for (int i = 0; i < cidArray.GetCount(); i++)
		{
			info = cidArray[i];
			delete info;
		}
	}

	ULONGLONG tc_curr = GetTickCount64();
	DWORD tc_elapsed_seconds = (DWORD)((tc_curr - tc_start) / 1000);
	DWORD tc_elapsed_milliseconds = (DWORD)(tc_curr - tc_start);

	TRACE(
		L"Mbox=%s\n"
		"\tEmbededImagesNoMatch=%d\n"
		"\tEmbededImagesFound=%d EmbededImagesNotFound=%d\n"
		"\tEmbededImagesFoundMHtml=%d EmbededImagesFoundMHtmlHtml=%d EmbededImagesFoundUnexpectedMHtml=%d\n"
		"\tEmbededImagesFoundCid=%d EmbededImagesNotFoundCid=%d\n"
		"\tEmbededImagesFoundHttp=%d EmbededImagesNotFoundHttp=%d\n"
		"\tEmbededImagesFoundHttps=%d EmbededImagesNotFoundHttps=%d\n"
		"\tEmbededImagesFoundMHtmlHttp=%d EmbededImagesNotFoundMHtmlHttp=%d\n"
		"\tEmbededImagesFoundMHtmlHttps=%d EmbededImagesNotFoundMHtmlHttps=%d\n"
		"\tEmbededImagesFoundData=%d EmbededImagesNotFoundData=%d\n"
		"\tEmbededImagesFoundLocalFile=%d EmbededImagesNotFoundLocalFile=%d\n"
		"\tElapsedTime=%ld.%ld\n\n",
			MboxMail::s_path,
			MboxMail::m_EmbededImagesNoMatch,
			MboxMail::m_EmbededImagesFound, MboxMail::m_EmbededImagesNotFound,
			MboxMail::m_EmbededImagesFoundMHtml, MboxMail::m_EmbededImagesFoundMHtmlHtml, MboxMail::m_EmbededImagesFoundUnexpectedMHtml,
			MboxMail::m_EmbededImagesFoundCid, MboxMail::m_EmbededImagesNotFoundCid,
			MboxMail::m_EmbededImagesFoundHttp, MboxMail::m_EmbededImagesNotFoundHttp,
			MboxMail::m_EmbededImagesFoundHttps, MboxMail::m_EmbededImagesNotFoundHttps,
			MboxMail::m_EmbededImagesFoundMHtmlHttp, MboxMail::m_EmbededImagesNotFoundMHtmlHttp,
			MboxMail::m_EmbededImagesFoundMHtmlHttps, MboxMail::m_EmbededImagesNotFoundMHtmlHttps,
			MboxMail::m_EmbededImagesFoundData, MboxMail::m_EmbededImagesNotFoundData,
			MboxMail::m_EmbededImagesFoundLocalFile, MboxMail::m_EmbededImagesNotFoundLocalFile,
			tc_elapsed_seconds, tc_elapsed_milliseconds
	);
	MessageBeep(MB_OK);
	MessageBeep(MB_OK);
	MessageBeep(MB_OK);

	return 1;
}

int NListView::ScanAllMailsInMbox_NewParser()
{
	ULONGLONG tc_start = GetTickCount64();

	MboxMail::m_EmbededImagesNoMatch = 0;

	MboxMail::m_EmbededImagesFoundMHtml = 0;
	MboxMail::m_EmbededImagesFoundMHtmlHtml = 0;
	MboxMail::m_EmbededImagesFoundUnexpectedMHtml = 0;

	MboxMail::m_EmbededImagesFound = 0;
	MboxMail::m_EmbededImagesFoundCid = 0;
	MboxMail::m_EmbededImagesFoundHttp = 0;
	MboxMail::m_EmbededImagesFoundHttps = 0;
	MboxMail::m_EmbededImagesFoundMHtmlHttp = 0;
	MboxMail::m_EmbededImagesFoundMHtmlHttps = 0;
	MboxMail::m_EmbededImagesFoundData = 0;
	MboxMail::m_EmbededImagesFoundLocalFile = 0;
	//
	MboxMail::m_EmbededImagesNotFound = 0;
	MboxMail::m_EmbededImagesNotFoundCid = 0;
	MboxMail::m_EmbededImagesNotFoundHttp = 0;
	MboxMail::m_EmbededImagesNotFoundHttps = 0;
	MboxMail::m_EmbededImagesNotFoundMHtmlHttp = 0;
	MboxMail::m_EmbededImagesNotFoundMHtmlHttps = 0;
	MboxMail::m_EmbededImagesNotFoundData = 0;
	MboxMail::m_EmbededImagesNotFoundLocalFile = 0;

	CFile fpm;
	CFileException ExError;
	if (!fpm.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);
#if 0
		CString txt = L"Could not open mail archive \"" + MboxMail::s_path;
		txt += L"\" file.\n";
		txt += exErrorStr;
#endif
		CString txt;
		CString fmt = L"Could not open mail archive \"%s\" file.\n%s";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, MboxMail::s_path, exErrorStr);

		//TRACE(L"%s\n", txt);

		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	char *token = 0;
	int tokenlen = 0;

	SimpleString outbuf(1024, 256);
	SimpleString tmpbuf(1024, 256);

	SimpleString *outbuflarge = MboxMail::m_outbuf;
	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;
	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);

	bool useMailPosition = true;

	CString bdycharset = L"UTF-8";
	CString bdy;

	UINT pageCode = 0;
	int textType = 1; //  Html
	int plainTextMode = 0;  // no extra img tags; html text has img tags already
	bool runInvestigation = true;

	CString imageCachePath;
	bool createCacheInvestigate = false;

	TRACE(L"Mbox=%s\n\n", MboxMail::s_path);

	int iItem;
	if (createCacheInvestigate == false)
	{
		for (iItem = 0; iItem < MboxMail::s_mails.GetCount(); iItem++)
		{
			if (iItem == 134049)
				int deb = 1;

			MboxMail *m = MboxMail::s_mails[iItem];

			outbuf.Clear();
			outbuflarge->Clear();

			pageCode = 0;

			int textlen = MboxMail::GetMailBody_mboxview(fpm, iItem, outbuflarge, pageCode, textType, plainTextMode);  // returns pageCode
			if (textlen != outbuflarge->Count())
				int deb = 1;

			if (outbuflarge->Count() != 0)
			{
				workbuf->ClearAndResize(outbuflarge->Count() * 2);
				// FIXME NListView::UpdateInlineSrcImgPath(outbuflarge->Data(), outbuflarge->Count(), workbuf, 0, iItem, useMailPosition, runInvestigation);
			}
		}
	}
	else
	{
		for (iItem = 0; iItem < MboxMail::s_mails.GetCount(); iItem++)
		{
			if (iItem == 134049)
				int deb = 1;

			MboxMail *m = MboxMail::s_mails[iItem];

			//CreateInlineImageFiles(fpm, iItem, imageCachePath, runInvestigation);   // FIXME

		}
	}
	fpm.Close();

	ULONGLONG tc_curr = GetTickCount64();
	DWORD tc_elapsed_seconds = (DWORD)((tc_curr - tc_start)/1000);
	DWORD tc_elapsed_milliseconds = (DWORD)(tc_curr - tc_start);

	TRACE(
		L"Mbox=%s\n"
		"\tEmbededImagesNoMatch=%d\n"
		"\tEmbededImagesFound=%d EmbededImagesNotFound=%d\n"
		"\tEmbededImagesFoundMHtml=%d EmbededImagesFoundMHtmlHtml=%d EmbededImagesFoundUnexpectedMHtml=%d\n"
		"\tEmbededImagesFoundCid=%d EmbededImagesNotFoundCid=%d\n"
		"\tEmbededImagesFoundHttp=%d EmbededImagesNotFoundHttp=%d\n"
		"\tEmbededImagesFoundHttps=%d EmbededImagesNotFoundHttps=%d\n"
		"\tEmbededImagesFoundMHtmlHttp=%d EmbededImagesNotFoundMHtmlHttp=%d\n"
		"\tEmbededImagesFoundMHtmlHttps=%d EmbededImagesNotFoundMHtmlHttps=%d\n"
		"\tEmbededImagesFoundData=%d EmbededImagesNotFoundData=%d\n"
		"\tEmbededImagesFoundLocalFile=%d EmbededImagesNotFoundLocalFile=%d\n"
		"\tElapsedTime=%ld.%ld\n\n",
		MboxMail::s_path,
		MboxMail::m_EmbededImagesNoMatch,
		MboxMail::m_EmbededImagesFound, MboxMail::m_EmbededImagesNotFound,
		MboxMail::m_EmbededImagesFoundMHtml, MboxMail::m_EmbededImagesFoundMHtmlHtml, MboxMail::m_EmbededImagesFoundUnexpectedMHtml,
		MboxMail::m_EmbededImagesFoundCid, MboxMail::m_EmbededImagesNotFoundCid,
		MboxMail::m_EmbededImagesFoundHttp, MboxMail::m_EmbededImagesNotFoundHttp,
		MboxMail::m_EmbededImagesFoundHttps, MboxMail::m_EmbededImagesNotFoundHttps,
		MboxMail::m_EmbededImagesFoundMHtmlHttp, MboxMail::m_EmbededImagesNotFoundMHtmlHttp,
		MboxMail::m_EmbededImagesFoundMHtmlHttps, MboxMail::m_EmbededImagesNotFoundMHtmlHttps,
		MboxMail::m_EmbededImagesFoundData, MboxMail::m_EmbededImagesNotFoundData,
		MboxMail::m_EmbededImagesFoundLocalFile, MboxMail::m_EmbededImagesNotFoundLocalFile,
		tc_elapsed_seconds, tc_elapsed_milliseconds
	);
	MessageBeep(MB_OK);
	MessageBeep(MB_OK);
	MessageBeep(MB_OK);

	return 1;
}

BOOL NListView::HasAnyAttachment(MboxMail *m)
{
	AttachmentConfigParams *attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
	if (attachmentConfigParams)
	{
		if (!attachmentConfigParams->m_bAnyAttachment_Indicator)
			return FALSE;
	}

	MailBodyContent *body;
	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];
		
		if (!body->m_attachmentName.IsEmpty())
			return TRUE;

		char *p = (char*)(LPCSTR)body->m_contentType;
		char *p_end = p + body->m_contentType.GetLength();
		int retLen1 = TextUtilsEx::strncmpUpper2Lower(p, p_end, "text/plain", 10);
		int retLen2 = TextUtilsEx::strncmpUpper2Lower(p, p_end, "text/htm", 8);

		if ((retLen1 != 0) &&  (retLen2 != 0))
			return true;
	}
	return FALSE;
}


BOOL NListView::MyCTimeToOleTime(MyCTime &ctimeDateTime, COleDateTime &coleDateTime, int rounding)
{
	// FIXMEFIXME: need to find time to investigate proper solution
	SYSTEMTIME sysTime;
	SYSTEMTIME sysTimeGMT;
	SYSTEMTIME sysTimeLocal;
	ctimeDateTime.GetAsSystemTime(sysTimeGMT, TRUE);
	ctimeDateTime.GetAsSystemTime(sysTimeLocal, FALSE);
	if (rounding == 2)  // DOWN
	{
		if (sysTimeGMT.wYear < sysTimeLocal.wYear)
			sysTime = sysTimeGMT;
		else if (sysTimeGMT.wMonth < sysTimeLocal.wMonth)
			sysTime = sysTimeGMT;
		else if (sysTimeGMT.wDay < sysTimeLocal.wDay)
			sysTime = sysTimeGMT;
		else
			sysTime = sysTimeLocal;
	}
	else if (rounding == 1)  // UP
	{
		if (sysTimeGMT.wYear < sysTimeLocal.wYear)
			sysTime = sysTimeLocal;
		else if (sysTimeGMT.wMonth < sysTimeLocal.wMonth)
			sysTime = sysTimeLocal;
		else if (sysTimeGMT.wDay < sysTimeLocal.wDay)
			sysTime = sysTimeLocal;
		else
			sysTime = sysTimeGMT;
	}
	else
	{
		sysTime = sysTimeGMT;
	}

	MyCTime::fixSystemtime(&sysTime);

	coleDateTime.SetDate(sysTime.wYear, sysTime.wMonth, sysTime.wDay);

	return TRUE;
}

BOOL NListView::OleTime2MyCTime(COleDateTime &coleDateTime, MyCTime &ctimeDateTime, BOOL roundUP)
{
	SYSTEMTIME sysTime;

	coleDateTime.GetAsSystemTime(sysTime);
	if (roundUP)
	{
		sysTime.wMinute = 59;
		sysTime.wHour = 23;
		sysTime.wSecond = 59;
		sysTime.wMilliseconds = 999;
		MyCTime::fixSystemtime(&sysTime);
	}
	ctimeDateTime.SetDateTime(sysTime);
	return TRUE;
}

void NListView::ResetFilterDates()
{
	m_advancedParams.ResetFilterDates();
	m_findParams.ResetFilterDates();
}

void NListView::PostMsgCmdParamAttachmentHint()
{
	BOOL isHintSet = MboxMail::m_HintConfig.IsHintSet(HintConfig::AttachmentConfigHint);
	if (isHintSet)
	{
		if (GetSafeHwnd())
		{
			BOOL res = PostMessage(WM_CMD_PARAM_ATTACHMENT_HINT_MESSAGE, 0, 0);
		}
	}
}

LRESULT NListView::OnCmdParam_AttachmentHint(WPARAM wParam, LPARAM lParam)
{
	MboxMail::ShowHint(HintConfig::AttachmentConfigHint, GetSafeHwnd());
	return 0;
}

int NListView::LoadFolderListFile_v2(CString &folderPath, CString &folderName)
{
	int ret = 1;  //OK

	MailArray *mailsArray = &MboxMail::s_mails;

#if 0
	if (MboxMail::IsFolderMailsSelected())
	{
		// should never be here
		return -1;
	}
#endif

	if (MboxMail::IsFolderMailsSelected())
	{
		CString txt = L"Folder Selected Mails List not empty. Overwrite?";
		ResHelper::TranslateString(txt);
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
			return -1;
	}

	CString mboxFileNamePath = MboxMail::s_path;

	if (mboxFileNamePath.IsEmpty()) {
		CString txt = L"Please open mail file first.";
		ResHelper::TranslateString(txt);
		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	CString path = MboxMail::GetLastPath();
	if (path.IsEmpty())
		return -1;  // Hopefully s_path will fail first


	CString driveName;
	CString mboxFileDirectory;
	CString mboxFileNameBase;
	CString mboxFileNameExtention;

	FileUtils::SplitFilePath(mboxFileNamePath, driveName, mboxFileDirectory, mboxFileNameBase, mboxFileNameExtention);

	CString mboxFileName = mboxFileNameBase + mboxFileNameExtention;

	CString folderCompletePath = driveName + mboxFileDirectory + L"\\" + L"Folders" + L"\\" + folderPath;

	CString folderNameCompletePath = folderCompletePath + L"\\" + folderName + L".mbox.mboxlist";

	if (!FileUtils::PathFileExist(folderNameCompletePath))
	{
#if 0
		CString txt = L"Mail List File \"" + folderNameCompletePath;
		txt += L"\" doesn't exist.\nCan't reload.";
#endif
		CString txt;
		CString fmt = L"Mail List File \"%s\" doesn't exist.\nCan't reload.";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, folderNameCompletePath);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	SerializerHelper sz(folderNameCompletePath);
	if (!sz.open(FALSE))
	{
#if 0
		CString txt = L"Could not open \"" + folderNameCompletePath;
		txt += L"\" file.\nMake sure file is not open on other applications.";
#endif
		CString txt;
		CString fmt = L"Could not open \"%s\" file.\nMake sure file is not open on other applications.";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, folderNameCompletePath);

		int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	int version;
	_int64 mailFileSize;
	_int64 mboxFileSize;
	int mailListCnt;

	CString txt = L"Mail list file\n\"" + folderNameCompletePath;
	txt += L"\"\nhas incompatible version or file is corrupted.\nCan't reload.\nPlease remove the file.";

	if (!sz.readInt(&version)) {
		sz.close();
		return -1;
	}

	if (version != MAIL_LIST_VERSION)
	{
		sz.close();

#if 0
		CString text = L"Mail list file\n\"" + folderNameCompletePath;
		CString strVersion;
		strVersion.Format(L"%d", (version - MAIL_LIST_VERSION_BASE));
		text += L"\".\nhas incompatible version\"" + strVersion + L"\". Expected version \"";
		strVersion.Format(L"%d", (MAIL_LIST_VERSION - MAIL_LIST_VERSION_BASE));
		text += strVersion + L"\".\nCan't reload.\nPlease remove the file.";
#endif
		CString txt;
		CString strVersion;
		strVersion.Format(L"%d", (version - MAIL_LIST_VERSION_BASE));
		CString strCachedVersion;
		strCachedVersion.Format(L"%d", (MAIL_LIST_VERSION - MAIL_LIST_VERSION_BASE));

		CString fmt = L"Mail list file\n\"%s\".\nhas incompatible version\"%s\". Expected version \"%s\".\nCan't reload.\nPlease remove the file.";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, folderNameCompletePath, strVersion, strCachedVersion);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	if (!sz.readInt64(&mailFileSize)) {
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	if (!sz.readInt64(&mboxFileSize)) {
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	if (!sz.readInt(&mailListCnt)) {
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	_int64 nMailFileSize = FileUtils::FileSize(mboxFileNamePath);


	if ((mailListCnt < 0) || (mailListCnt > MboxMail::s_mails_ref.GetCount()) ||
		(mailFileSize != nMailFileSize))
	{
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	BOOL errorDoNothing = FALSE;
	BOOL verifyOnly = TRUE;
	__int64 pos = 0;
	BOOL retval = sz.GetReadPointer(pos);
	if (retval == FALSE)
	{
		int deb = 1;
	}
	ret = PopulateMailArray(sz, MboxMail::s_mails_folder, mailListCnt, verifyOnly);
	if (ret > 0) 
	{
		verifyOnly = FALSE;
		BOOL retval = sz.SetReadPointer(pos);
		if (retval == FALSE)
		{
			int deb = 1;
		}
		ret = PopulateMailArray(sz, MboxMail::s_mails_folder, mailListCnt, verifyOnly);
	}
	else
	{
		sz.close();
#if 0
		CString txt = L"Mail list file\n\"" + folderNameCompletePath;
		txt += L"\"\n is invalid or corrupted. Can't reload.\nPlease remove the file.";
#endif

		CString fmt = L"Mail list file\n\"%s\"\n is invalid or corrupted. Can't reload.\nPlease remove the file.";
		ResHelper::TranslateString(fmt);
		CString txt;
		txt.Format(fmt, folderNameCompletePath);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);

		return -1; // do nothing
	}

	sz.close();

	if (ret < 0)
	{
		MboxMail::s_mails_folder.SetSizeKeepData(0);
#if 0
		CString txt = L"Mail list file\n\"" + folderNameCompletePath;
		txt += L"\"\n is invalid or corrupted. Can't reload.\nPlease remove the file.";
#endif

		CString fmt = L"Mail list file\n\"%s\"\n is invalid or corrupted. Can't reload.\nPlease remove the file.";
		ResHelper::TranslateString(fmt);
		CString txt;
		txt.Format(fmt, folderNameCompletePath);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
	}
	else
	{
		MboxMail::SortByDate(&MboxMail::s_mails_folder);
		if (MboxMail::s_mails_folder.GetCount() > 0)
			MboxMail::m_folderMails.m_lastSel = 0;
		else
			MboxMail::m_folderMails.m_lastSel = -1;
		MboxMail::m_folderMails.b_mails_which_sorted = 1;

		if (MboxMail::IsFolderMailsSelected())
		{
			SwitchToMailList(IDC_FOLDER_LIST, TRUE);
		}
		else if (MboxMail::IsAllMailsSelected())
		{
			SwitchToMailList(IDC_FOLDER_LIST, FALSE);
		}
	}

	return ret;
}


// This is used tgo load mail indexes into arra from list file such as Label list file
int NListView::PopulateMailArray(SerializerHelper &sz, MailArray &mArray, int mailListCnt, BOOL verifyOnly)
{
	int ret = 1;

	MboxMail *m;

	_int64 startOff;
	int length;
	int index;
	int contentDetailsArrayCnt;

	if (!verifyOnly)
		mArray.SetSize(mailListCnt);

	int i = 0;
	for (i = 0; i < mailListCnt; i++)
	{
		if (!sz.readInt64(&startOff))
			break;
		if (!sz.readInt(&length))
			break;
		if (!sz.readInt(&index))
			break;
		if (!sz.readInt(&contentDetailsArrayCnt))
			break;

		if ((index < 0) || (index >= MboxMail::s_mails_ref.GetCount()))
		{
			ret = -1;
			break;
		}

		m = MboxMail::s_mails_ref[index];
		if ((m->m_startOff != startOff) ||
			(m->m_length != length) ||
			(m->m_index != index) ||
			(m->m_ContentDetailsArray.size() != contentDetailsArrayCnt))
		{
			ret = -1;
			break;
		}

		if (!verifyOnly)
			mArray[i] = m;

		int deb = 1;
	}

	if (!verifyOnly)
	{
		if (ret > 0)
			mArray.SetSizeKeepData(i);
		else
			mArray.SetSizeKeepData(0);
	}


	return ret;
}

int NListView::CopyMailsToFolders()
{
	CCheckListBoxDlg dlg;
	dlg.m_title = L"Select Folders";
	for (int ii = 0; ii < 100; ii++)
	{
		dlg.m_InList.Add(L"11111111111xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyy");
	}

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	NTreeView *pTreeView = 0;
	if (pFrame)
		pTreeView = pFrame->GetTreeView();

	CArray<CString> folderList;
	if (pTreeView)
	{
		CString mboxFileName;

		FileUtils::CPathStripPath((LPCWSTR)MboxMail::s_path, mboxFileName);

		pTreeView->CreateFlatFolderList(mboxFileName, folderList);
	}

	dlg.m_InList.Copy(folderList);

	int nResponse = (int)dlg.DoModal();
	if (nResponse == IDOK)
	{
		int i;
		for (i = 0; i < dlg.m_OutList.GetCount(); i++)
		{
			CString &s = dlg.m_OutList[i];
			int deb = 1;
		}
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with Cancel
	}
	return 1;
}

BOOL NListView::OnEraseBkgnd(CDC* pDC)
{
	//BOOL ret = CWnd::OnEraseBkgnd(pDC);

	CRect rect;
	GetClientRect(&rect);

	DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailSummaryList);
	//if (!m_bApplyColorStyle) color = RGB(255, 255, 255);

	m_list.SetBkColor(color);
	pDC->FillRect(&rect, &CBrush(color));

	return TRUE;
	//return FALSE;
}


// Expensive. Optimistic and needs stronger parsing
// Finds first <div class=*Section1* >  ns removes class=*Section1*
// Should I simplify and remove *Section1*  only ???
int NListView::RemoveMicrosoftSection1Class(SimpleString* inbuf, int offset)
{
	static char* cDiv = "div";
	static int cDivLen = istrlen(cDiv);

	static char* pat = "section1";
	static int patLen = istrlen(pat);

	static char* cClass = "class";
	static int cClassLen = istrlen(cClass);

	int bodyTagPos = inbuf->FindNoCase(offset, "<body", 5);
	if (bodyTagPos < 0)
		return -1;

	// Find first <div class=*Section1* >.  Optimistic and simplistic
	int divTagPos = inbuf->FindNoCase(bodyTagPos, "<div", 4);
	if (divTagPos < 0)
		return -1;

	offset = divTagPos;

	char* divTagBeg;
	int divTagTotalLen;

	char* data = inbuf->Data(offset);
	int datalen = inbuf->Count() - offset;
	int ret = HtmlUtils::FindHtmlTag(data, datalen, cDiv,cDivLen, divTagBeg, divTagTotalLen);
	if (ret <= 0)
		return -1;

	int max_offset = offset + divTagTotalLen;

	int patPos = inbuf->FindNoCase(offset, pat, patLen, max_offset);
	if (patPos < 0)
		return -1;

	int i;
	char c;
	for (i = patPos; i >= offset; i--)
	{
		c = inbuf->GetAt(i);
		if (c == '=')
			break;
	}
	if (i == patPos)
		return -1;

	// Assumes class=*Sections1 or class="*Sections1" , no white spaces. Potential problem
	int classBegin = i - 5;
	char* classAttrib = inbuf->Data(classBegin);

	int   retClass = TextUtilsEx::strncmpUpper2Lower(classAttrib, cClassLen, cClass, cClassLen);
	if (retClass == 0)
	{
		int classEnd = patPos + patLen;
		c = inbuf->GetAt(classEnd);
		if (c == '"')  // FIXMEFIXME should skip white characters first and check for '"'
			classEnd++;

		char* next = inbuf->Data(classEnd);

		int cnt = inbuf->Remove(classBegin, classEnd);

		return inbuf->Count();
	}
	else
	{
		; // remove *Section1* or do the proper job
	}

	return -1;
}

int NListView::ReplaceBlockqouteTagWithDivTag(char* inData, int indDataLen, SimpleString* outbuf, BOOL ReplaceAllWhiteBackgrounTags)
{
	static char* pOldTag = "blockquote";
	static int oldTagLen = istrlen(pOldTag);

	static char* pNewTag = "div style=\"margin:0px 0px 0px 0.8ex; border-left:1px solid rgb(204, 204, 204); padding-left:1ex;\" ";
	static int newTagLen = istrlen(pNewTag);

	char* p_mark;
	char* p = inData;
	char* e = p + indDataLen;
	char* pTag;
	char c;
	int count;

	p_mark = p;

	BOOL blockQuoteFound = FALSE;  // helps to detect valid BLOCKQUOTE tags
	while (p < e)
	{
		c = *p;
		if ((c != 'b') && (c != 'B'))
		{
			p++;
			continue;
		}
		count = IntPtr2Int(e - p);

		pTag = TextUtilsEx::findNoCaseP(p, count, pOldTag, oldTagLen);
		if (pTag == 0)
		{
			break;
		}

		c = *(pTag - 1);
		if ((c != '<') && (c != '/'))
		{
			p = pTag + oldTagLen;
			continue;
		}

		if (c == '/')
		{
			c = *(pTag - 2);
			if (c != '<')
			{
				p = pTag + oldTagLen;
				continue;
			}
			else
				blockQuoteFound = TRUE;  // not bullet prove but we need to discover at least one instance
		}

		outbuf->Append(p_mark, IntPtr2Int(pTag - p_mark));
		outbuf->Append(pNewTag, newTagLen);

		p = pTag + oldTagLen;
		p_mark = p;
	}

	if (blockQuoteFound == FALSE)
	{
		outbuf->Clear();
	}
	else if (outbuf->Count())
	{
		int cnt = IntPtr2Int(e - p_mark);
		outbuf->Append(p_mark, cnt);
	}

	return 1;
}

// Expensive:  Find and Convert <pre ... >  ... CRLF ... CRLF .. </pre>  to <p ... > ... <br>CRLF ... <br>CRLF ... </p>
// to force wrap of long text lines 
int NListView::ReplacePreTagWitPTag(char *inData, int indDataLen, SimpleString *outbuf, BOOL ReplaceAllWhiteBackgrounTags)
{
	static char *preTag = "<pre";
	static int preTagLen = istrlen(preTag);

	static char *preEndTag = "</pre";
	static int preEndTagLen = istrlen(preEndTag);

	static char *pTag = "<p";
	static int pTagLen = istrlen(pTag);

	static char *pEndTag = "</p";
	static int pEndTagLen = istrlen(pEndTag);

	static char *nbsp1 = "&nbsp;";
	static int nbsp1Len = istrlen(nbsp1);

	//outbuf->ClearAndResize(indDataLen + 128);

	char *p_mark;
	char *p = inData;
	int count = indDataLen;
	char *e = p + indDataLen;
	char *e_data = e;

	char *pPreTag = 0;
	char *pPreEndTag = 0;

	char c;
	char c_next;

	p_mark = p;

	while (p < e)  // always true
	{
		pPreTag = 0;
		pPreEndTag = 0;
		count = IntPtr2Int(e - p);

		pPreTag = TextUtilsEx::findNoCaseP(p, count, preTag, preTagLen);
		if (pPreTag == 0)
		{
			break;
		}

		outbuf->Append(p, IntPtr2Int(pPreTag - p));
		outbuf->Append(pTag, pTagLen); // append <p

		p += (pPreTag - p) + preTagLen;
		p_mark = p;

		while (p < e)
		{
			count = IntPtr2Int(e - p);
			p = TextUtilsEx::findNoCaseP(p, count, "<", 1);
			if (p == 0)
			{
				outbuf->Clear();
				return 0;
			}

			if (TextUtilsEx::strncmpUpper2Lower(p, e, preEndTag, preEndTagLen) != 0)
			{
				//  <p.. doesn't match </pre
				// check if <p.. matches </pre
				if (TextUtilsEx::strncmpUpper2Lower(p, e, preTag, preTagLen) != 0)
				{
					// ignore and continue
					p++;
					continue;
				}
				else
				{
					//  nested pre tag, give up
					outbuf->Clear();
					return 0;
				}
			}
			else
				break;  // found ending </pre, i.e p_mark points to <pre and p points to </pre
		}

		pPreEndTag = p;  // found ending </pre, i.e p_mark points to <pre and p points to </pre i.e pPreEndTag

		p = p_mark;
		while (p < pPreEndTag)
		{
			c = *p;
			if (c == ' ')
			{
				c_next = *(p+1);
				if (c_next == ' ')
				{
					outbuf->Append(p_mark, IntPtr2Int(p - p_mark));
					outbuf->Append(nbsp1, nbsp1Len);
					p++;
					while (p < pPreEndTag)
					{
						outbuf->Append(nbsp1, nbsp1Len);
						p++;
						c_next = *p;
						if (c_next != ' ')
							break;
					}
					p_mark = p;
				}
				else
					p++;
			}
			else if ((c == '\r') || (c == '\n'))
			{
				outbuf->Append(p_mark, IntPtr2Int(p - p_mark));
				outbuf->Append("<br>", 4);
				p++;
				c_next = *p;
				if (c == '\r')
				{
					if (c_next == '\n')
					{
						p++;
					}
				}
				p_mark = p;
			}
			else
				p++;
		}
		outbuf->Append(p_mark, IntPtr2Int(p - p_mark));
		outbuf->Append(pEndTag, pEndTagLen);  // append </p
		p += preEndTagLen;
		p_mark = p;
	}
	if (outbuf->Count())
	{
		int cnt = IntPtr2Int(e - p_mark);
		outbuf->Append(p_mark, cnt);
	}

	return 1;
}

int NListView::MakeSpacesAsNBSP(char *inData, int indDataLen, SimpleString *outbuf, BOOL ReplaceAllWhiteBackgrounTags)
{
	static char *minMultiSpaces = "  ";
	static int minMultiSpacesLen = istrlen(minMultiSpaces);

	static char *nbsp2 = "&nbsp;&nbsp;";
	static int nbsp2Len = istrlen(nbsp2);

	char *p_mark;
	char *p = inData;
	int count = indDataLen;
	char *e = p + indDataLen;
	char *e_data = e;

	char *pMSpace = 0;

	p_mark = p;

	while (p < e)  // always true
	{
		count = IntPtr2Int(e - p);

		pMSpace = TextUtilsEx::findNoCaseP(p, count, minMultiSpaces, minMultiSpacesLen);
		if (pMSpace == 0)
		{
			break;
		}

		outbuf->Append(p, IntPtr2Int(pMSpace - p));
		outbuf->Append(nbsp2, nbsp2Len);

		p += (pMSpace - p) + minMultiSpacesLen;
		p_mark = p;
	}

	int cnt = IntPtr2Int(e - p_mark);
	if (cnt)
	{
		outbuf->Append(p_mark, cnt);
	}

	return 1;
}

int NListView::AddMaxWidthToHref(char *inData, int indDataLen, SimpleString *outbuf, BOOL ReplaceAllWhiteBackgrounTags)
{
	static char *hrefAttr = "href=";
	static int hrefAttrLen = istrlen(hrefAttr);

	char *p_mark;
	char *p = inData;
	int count = indDataLen;
	char *e = p + indDataLen;
	char *e_data = e;

	char c;
	int cnt1;
	int cnt2;

	p_mark = p;
	while (p < e)  // always true
	{
		count = IntPtr2Int(e - p);
		p = TextUtilsEx::findNoCaseP(p, count, hrefAttr, hrefAttrLen);
		if (p == 0)
		{
			break;
		}
		c = *(p-1);
		if ((c != ' ') && (c != '\r') && (c != '\n'))
		{
			p++;
			continue;
		}
		
		cnt1 = IntPtr2Int(p - p_mark);
		outbuf->Append(p_mark, cnt1);
		outbuf->Append(" style=\"max-width:100%;\" ");

		p_mark = p;
		p++;
	}
	if (outbuf->Count())
	{
		cnt2 = IntPtr2Int(e - p_mark);
		outbuf->Append(p_mark, cnt2);
	}

	return 1;
}


// FIXME not used anymore ?? didn't work, conidder to delete this. Found different "solution", see NListView::ReplaceBlockqouteTagWithDivTag
int NListView::AddMaxWidthToBlockquote(char *inData, int indDataLen, SimpleString *outbuf, BOOL ReplaceAllWhiteBackgrounTags)
{
	static char *blockquoteAttr = "<blockquote";
	static int blockquoteAttrLen = istrlen(blockquoteAttr);

	char *p_mark;
	char *p = inData;
	int count = indDataLen;
	char *e = p + indDataLen;
	char *e_data = e;

	//char c;
	int cnt1;
	int cnt2;

	p_mark = p;
	while (p < e)  // always true
	{
		count = IntPtr2Int(e - p);
		p = TextUtilsEx::findNoCaseP(p, count, blockquoteAttr, blockquoteAttrLen);
		if (p == 0)
		{
			break;
		}
#if 0
		c = *(p-1);
		if ((c != ' ') && (c != '\r') && (c != '\n'))
		{
			p++;
			continue;
		}
#endif
		p += blockquoteAttrLen;

		cnt1 = IntPtr2Int(p - p_mark);
		outbuf->Append(p_mark, cnt1);
		outbuf->Append(" style=\"max-width:100%; overflow-x:break-word; word-wrap:break-word; overflow-wrap:break-word !important;width:99%;\" ");

		p_mark = p;
		p++;
	}
	if (outbuf->Count())
	{
		cnt2 = IntPtr2Int(e - p_mark);
		outbuf->Append(p_mark, cnt2);
	}

	return 1;
}

int NListView::AddMaxWidthToDIV(char *inData, int indDataLen, SimpleString *outbuf, BOOL ReplaceAllWhiteBackgrounTags)
{
	static char *divAttr = "<div";
	static int divAttrLen = istrlen(divAttr);

	char *p_mark;
	char *p = inData;
	int count = indDataLen;
	char *e = p + indDataLen;
	char *e_data = e;

	//char c;
	int cnt1;
	int cnt2;

	p_mark = p;
	while (p < e)  // always true
	{
		count = IntPtr2Int(e - p);
		p = TextUtilsEx::findNoCaseP(p, count, divAttr, divAttrLen);
		if (p == 0)
		{
			break;
		}
#if 0
		c = *(p-1);
		if ((c != ' ') && (c != '\r') && (c != '\n'))
		{
			p++;
			continue;
		}
#endif
		p += divAttrLen;

		cnt1 = IntPtr2Int(p - p_mark);
		outbuf->Append(p_mark, cnt1);
		outbuf->Append(" style=\"max-width:80%;\" ");

		p_mark = p;
		p++;
	}
	if (outbuf->Count())
	{
		cnt2 = IntPtr2Int(e - p_mark);
		outbuf->Append(p_mark, cnt2);
	}

	return 1;
}

int NListView::SetBackgroundColor(char *inData, int indDataLen, SimpleString *outbuf, BOOL ReplaceAllWhiteBackgrounTags)
{
	static char *bodyTag = "<body ";
	static int bodyTagLen = istrlen(bodyTag);

	static char *styleTag = " style";
	static int styleTagLen = istrlen(styleTag);

	static char *bgColorTag = "bgcolor";
	static int bgColorTagLen = istrlen(bgColorTag);

	static char *backgroundColorTag = "background-color";
	static int backgroundColorTagLen = istrlen(backgroundColorTag);

	static char *backgroundTag = "background";
	static int backgroundTagLen = istrlen(backgroundTag);

	static char *bodyTagEnd = ">";
	static int bodyTagEndLen = istrlen(bodyTagEnd);

	outbuf->ClearAndResize(indDataLen + 128);

	if (indDataLen < bodyTagLen)
		return 1;

	DWORD clr = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);
	if (clr == COLOR_WHITE)
		return 1;

	CStringA colorStr;

	int retC2S = NListView::Color2Str(clr, colorStr);

	char *p_mark;
	char *p = inData;
	int count = indDataLen;
	char *e = p + indDataLen;
	char *e_data = e;

	char * pBodyBeg = 0;
	char * pBodyEnd = 0;

	char *pStyleTag = 0;
	char *pBgColorTag = 0;
	char *pBackgroundColorTag = 0;
	char *pBackgroundTag = 0;
	char *pBgTag;

	char *pBackgroundColorTagBeg = 0;
	char *pBackgroundColorTagEnd = 0;

	// Replace background def in body only.  Limit search to body begin and body end
	if (ReplaceAllWhiteBackgrounTags == FALSE)
	{
		pBodyBeg = TextUtilsEx::findNoCaseP(p, count, bodyTag, bodyTagLen);
		if (pBodyBeg == 0)
			return 0;

		p = pBodyBeg + bodyTagLen;
		count = IntPtr2Int(e - p);
		if (count <= 0)
			return 0;

		pBodyEnd = TextUtilsEx::findNoCaseP(p, count, bodyTagEnd, bodyTagEndLen);
		if (pBodyEnd == 0)
			return 0;

		e = pBodyEnd+1;

		count = IntPtr2Int(e - p);
		if (count <= 0)
			return 0;
	}

	outbuf->Append(inData, IntPtr2Int(p - inData));

	p_mark = p;

	while (p < e)  // always true
	{
		pStyleTag = 0;
		pBgColorTag = 0;
		pBackgroundColorTag = 0;
		pBackgroundTag = 0;
		pBgTag = 0;

		pBackgroundColorTagBeg = 0;
		pBackgroundColorTagEnd = 0;

		count = IntPtr2Int(e - p);
		p = TextUtilsEx::findNoCaseP(p, count, "b", 1);
		if (p == 0)
			break;

		if (TextUtilsEx::strncmpUpper2Lower(p, e, backgroundColorTag, backgroundColorTagLen) == 0)
		{
			pBgTag = p;
			pBackgroundColorTag = p;
			p += backgroundColorTagLen;
		}
		else if (TextUtilsEx::strncmpUpper2Lower(p, e, bgColorTag, bgColorTagLen) == 0)
		{
			pBgTag = p;
			pBgColorTag = p;
			p += bgColorTagLen;
		}
		else if (TextUtilsEx::strncmpUpper2Lower(p, e, backgroundTag, backgroundTagLen) == 0)
		{
			pBgTag = p;
			pBackgroundTag = p;
			p += backgroundTagLen;
		}

		if ((pBackgroundColorTag == 0) && (pBgColorTag == 0) && (pBackgroundTag == 0))
		{
			p++;
			continue;
		}

		char c = *(pBgTag - 1);
		if ((c != ' ') && (c != ';'))
			continue;  // TODO: ??

		p = TextUtilsEx::SkipWhite(p);
		if ((*p != '=') && (*p != ':'))
		{
			if (pBackgroundTag)
			{
				p++;
				continue;
			}
		}

		p++;
		if ((p + 2) > e)
			break;

		p = TextUtilsEx::SkipWhite(p);
		BOOL valQuoted = FALSE;
		if (*p == '"')
		{
			valQuoted = TRUE;
			p++;
		}

		if (*p == '#')
			p++;

		pBackgroundColorTagBeg = p;

		if (valQuoted)
		{
			p = TextUtilsEx::findOneOf(p, e, "\" ;><");
			if (p == 0)
			{
				outbuf->Clear();
				return 0;
			}
			if (*p != '\"')
			{
				outbuf->Clear();
				return 0;
			}
		}
		else
		{
			p = TextUtilsEx::findOneOf(p, e, " ;><");
			if (p == 0)
			{
				outbuf->Clear();
				return 0;
			}
			if (*p == '<')
			{
				outbuf->Clear();
				return 0;
			}
		}

		pBackgroundColorTagEnd = p;

		CStringA pBackgroundColorTagVal;
		if (pBackgroundColorTagBeg && pBackgroundColorTagEnd)
		{
			int tagLength = IntPtr2Int(pBackgroundColorTagEnd - pBackgroundColorTagBeg);
			char *b = pBackgroundColorTagVal.GetBufferSetLength(tagLength);
			::memcpy(b, pBackgroundColorTagBeg, tagLength);
			pBackgroundColorTagVal.ReleaseBuffer();

			if ((pBackgroundColorTagVal.CompareNoCase("fff") == 0) 
				|| (pBackgroundColorTagVal.CompareNoCase("ffffff") == 0)
				)
			{
				outbuf->Append(p_mark, IntPtr2Int(pBackgroundColorTagBeg - p_mark));
				outbuf->Append(colorStr, colorStr.GetLength());
				p_mark = pBackgroundColorTagEnd;
			}
			else if ((pBackgroundColorTagVal.CompareNoCase("white") == 0)
				|| (pBackgroundColorTagVal.CompareNoCase("rgb(255,255,255)") == 0)
				|| (pBackgroundColorTagVal.CompareNoCase("hsl(0,0%,100%)") == 0)
				)
			{
				outbuf->Append('#');
				outbuf->Append(p_mark, IntPtr2Int(pBackgroundColorTagBeg - p_mark));
				outbuf->Append(colorStr, colorStr.GetLength());
				p_mark = pBackgroundColorTagEnd;
			}
			else
			{
				int deb = 1;
			}
		}
		p++;
	}
	outbuf->Append(p_mark, IntPtr2Int(e_data - p_mark));

	return 1;
}

int NListView::FindBodyTag(char *inData, int indDataLen, char *&tagBeg, int &tagDataLen)
{
	static char *bodyTag = "<body ";
	static int bodyTagLen = istrlen(bodyTag);

	static char *bodyTagEnd = ">";
	static int bodyTagEndLen = istrlen(bodyTagEnd);

	tagBeg = 0;
	tagDataLen = 0;

	char *p = inData;
	int count = indDataLen;
	char *e = p + indDataLen;
	char *e_data = p + indDataLen;

	char *pBodyBeg = 0;
	char *pBodyEnd = 0;

	pBodyBeg = TextUtilsEx::findNoCaseP(p, count, bodyTag, bodyTagLen);
	if (pBodyBeg == 0)
		return 0;

	p = pBodyBeg + bodyTagLen;
	count = IntPtr2Int(e - p);
	if (count <= 0)
		return 0;

	pBodyEnd = TextUtilsEx::findNoCaseP(p, count, bodyTagEnd, bodyTagEndLen);
	if (pBodyEnd == 0)
		return 0;

	tagDataLen = IntPtr2Int(pBodyEnd + 1 - pBodyBeg);
	if (tagDataLen <= 0)
		tagDataLen = 0;
	else
		tagBeg = pBodyBeg;

	return tagDataLen;
}

int NListView::RemoveBodyBackgroundColorAndWidth(char *inData, int inDataLen, SimpleString *outbuf, CStringA &bodyBackgroundColor, CStringA &bodyWidth, 
	BOOL removeBgColor, BOOL removeWidth)
{
	static char *widthTag = "width";
	static int widthTagLen = istrlen(widthTag);

	static char *bgColorTag = "bgcolor";
	static int bgColorTagLen = istrlen(bgColorTag);

#if 0
	DWORD clr = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);
	if (clr == COLOR_WHITE)
		int deb = 1; // return 1;

	CString colorStr;

	int retC2S = NListView::Color2Str(clr, colorStr);
#endif

	HtmlAttribInfo bodyBackgroundColorAttrib;
	CStringA widthAttribName = widthTag;
	HtmlAttribInfo bodyWidthAttrib(widthAttribName);

	char *bodyTagBeg = 0;
	int bodyTagDataLen = 0;
	int bodyLen = FindBodyTag(inData, inDataLen, bodyTagBeg, bodyTagDataLen);
	if (bodyLen == 0)
		return 0;

	char *bgTagBeg = 0;
	int bgTagDataLen = 0;
	int bgLen = 0;

	if (removeBgColor)
	{
		bgLen = NListView::FindBodyTagAttrib(bodyTagBeg, bodyTagDataLen, bgColorTag, bgColorTagLen, bgTagBeg, bgTagDataLen, bodyBackgroundColor, bodyBackgroundColorAttrib);

		if (bgTagDataLen <= 0)
		{
			bgTagBeg = 0;
			bgTagDataLen = 0;
			bgLen = 0;

			bgLen = NListView::FindBodyBackgroundColor(bodyTagBeg, bodyTagDataLen, bgTagBeg, bgTagDataLen, bodyBackgroundColor, bodyBackgroundColorAttrib);
		}
	}

	char *widthTagBeg = 0;
	int widthTagDataLen = 0;
	int widthLen = 0;
	if (removeWidth)
		widthLen = NListView::FindBodyTagAttrib(bodyTagBeg, bodyTagDataLen, widthTag, widthTagLen, widthTagBeg, widthTagDataLen, bodyWidth, bodyWidthAttrib);

	char *inDataEnd = inData + inDataLen;
	char *bgTagEnd = bgTagBeg + bgTagDataLen;
	char *widthTagEnd = widthTagBeg + widthTagDataLen;

	outbuf->ClearAndResize(inDataLen + 128);

	if (bgTagDataLen && widthTagDataLen)
	{
		if (bgTagBeg < widthTagBeg)
		{
			outbuf->Append(inData, IntPtr2Int(bgTagBeg - inData));
			outbuf->Append(bgTagEnd, IntPtr2Int(widthTagBeg - bgTagEnd));
			outbuf->Append(widthTagEnd, IntPtr2Int(inDataEnd - widthTagEnd));
		}
		else
		{
			outbuf->Append(inData, IntPtr2Int(widthTagBeg - inData));
			outbuf->Append(widthTagEnd, IntPtr2Int(bgTagBeg - widthTagEnd));
			outbuf->Append(bgTagEnd, IntPtr2Int(inDataEnd - bgTagEnd));
		}
	}
	else if (bgTagDataLen)
	{
		outbuf->Append(inData, IntPtr2Int(bgTagBeg - inData));
		outbuf->Append(bgTagEnd, IntPtr2Int(inDataEnd - bgTagEnd));
	}
	else if (widthTagDataLen)
	{
		outbuf->Append(inData, IntPtr2Int(widthTagBeg - inData));
		outbuf->Append(widthTagEnd, IntPtr2Int(inDataEnd - widthTagEnd));
	}
	if (outbuf->Count())
		int deb = 1;

	return outbuf->Count();
}

int NListView::FindBodyTagAttrib(char *inData, int indDataLen, char *tag, int tagLen, char *&tagBeg, int &tagDataLen, CStringA &attribVal, HtmlAttribInfo &bodyAttrib)
{
	int widthTagPos = TextUtilsEx::findNoCase(inData, indDataLen, tag, tagLen);
	if (widthTagPos >= 0)
		int deb = 1;

	char *pTagBeg= inData + widthTagPos;

	char c = *(pTagBeg - 1);
	if ((c != ' ') && (c != ';'))
		return 0;

	// get attrib len and value


	char *pAttribValBeg = pTagBeg;
	char *pAtribValEnd = pTagBeg;
	//
	char *v = pTagBeg + tagLen;
	char *p = v;
	char *e = inData + indDataLen;

	p = TextUtilsEx::SkipWhite(p);
	if ((*p != '=') && (*p != ':'))
	{
		return 0;
	}

	bodyAttrib.m_separator = *p;

	p++;
	if ((p + 2) > e)
		return 0;

	BOOL quotedVal = FALSE;
	p = TextUtilsEx::SkipWhite(p);
	if (*p == '"')
	{
		quotedVal = TRUE;
		bodyAttrib.m_quoted = TRUE;
		p++;
	}

	pAttribValBeg = p;

	if (quotedVal)
	{
		p = TextUtilsEx::findOneOf(p, e, "\"><");
		if (p == 0)
			return 0;
		if (*p != '\"')
			return 0;
		pAtribValEnd = p;
		p++;
	}
	else
	{
		p = TextUtilsEx::findOneOf(p, e, "\" ;><");
		if (p == 0)
			return 0;
		if (*p == '<')
			return 0;
		pAtribValEnd = p;
	}


	char *pTagEnd = p;
	p = TextUtilsEx::SkipWhite(p);
	if (*p == ';')
	{
		bodyAttrib.m_terminator = ';';
		p++;
		pTagEnd = p;
	}

	CStringA val;
	int tagLength = IntPtr2Int(pAtribValEnd - pAttribValBeg);
	char *b = val.GetBufferSetLength(tagLength);
	::memcpy(b, pAttribValBeg, tagLength);
	val.ReleaseBuffer();

	attribVal = val;
	bodyAttrib.m_value = val;

	tagBeg = pTagBeg;
	tagDataLen = IntPtr2Int(pTagEnd - pTagBeg);

	if (tagBeg && (tagDataLen > 0))
		bodyAttrib.m_attribString.Append(tagBeg, tagDataLen);

	return tagDataLen;
}

int NListView::FindBodyBackgroundColor(char *inData, int inDataLen, char *&tagBeg, int &tagDataLen, CStringA &bodyBackgroundColor, HtmlAttribInfo &bodyBackgroundColorAttrib)
{
	static char *bodyTag = "<body ";
	static int bodyTagLen = istrlen(bodyTag);

	static char *styleTag = " style";
	static int styleTagLen = istrlen(styleTag);

	static char *bgColorTag = "bgcolor";
	static int bgColorTagLen = istrlen(bgColorTag);

	static char *backgroundColorTag = "background-color";
	static int backgroundColorTagLen = istrlen(backgroundColorTag);

	static char *backgroundTag = "background";
	static int backgroundTagLen = istrlen(backgroundTag);

	static char *bodyTagEnd = ">";
	static int bodyTagEndLen = istrlen(bodyTagEnd);

	char *p_mark;
	char *p = inData;
	int count = inDataLen;
	char *e = p + inDataLen;
	//char *e_data = p + inDataLen;

	char * pBodyBeg = inData;
	char * pBodyEnd = inData + inDataLen;

	char *pStyleTag = 0;
	char *pBgColorTag = 0;
	char *pBackgroundColorTag = 0;
	char *pBackgroundTag = 0;
	char *pBgTag = 0;
	char *pBgTagEnd = 0;
	int bgTagLen = 0;

	char *pBackgroundColorTagValBeg = 0;
	char *pBackgroundColorTagValEnd = 0;

	p = pBodyBeg;
	e = pBodyEnd;

	p_mark = p;

	tagBeg = 0;
	tagDataLen = 0;

	CStringA pBackgroundColorTagVal;
	while (p < e)  // always true
	{
		pStyleTag = 0;
		pBgColorTag = 0;
		pBackgroundColorTag = 0;
		pBackgroundTag = 0;
		pBgTag = 0;
		bgTagLen = 0;
		pBgTagEnd = 0;

		pBackgroundColorTagValBeg = 0;
		pBackgroundColorTagValEnd = 0;

		bodyBackgroundColorAttrib.Reset();

		count = IntPtr2Int(e - p);
		p = TextUtilsEx::findNoCaseP(p, count, "b", 1);
		if (p == 0)
			break;

		if (TextUtilsEx::strncmpUpper2Lower(p, e, backgroundColorTag, backgroundColorTagLen) == 0)
		{
			pBgTag = p;
			bgTagLen = backgroundColorTagLen;
			pBackgroundColorTag = p;
			p += backgroundColorTagLen;
		}
		else if (TextUtilsEx::strncmpUpper2Lower(p, e, bgColorTag, bgColorTagLen) == 0)
		{
			pBgTag = p;
			bgTagLen = bgColorTagLen;
			pBgColorTag = p;
			p += bgColorTagLen;
		}
		else if (TextUtilsEx::strncmpUpper2Lower(p, e, backgroundTag, backgroundTagLen) == 0)
		{
			pBgTag = p;
			bgTagLen = backgroundTagLen;
			pBackgroundTag = p;
			p += backgroundTagLen;
		}

		if ((pBackgroundColorTag == 0) && (pBgColorTag == 0) && (pBackgroundTag == 0))
		{
			p++;
			continue;
		}

		char c = *(pBgTag - 1);
		if ((c != ' ') && (c != ';'))
			continue;  // TODO: ??

		p = TextUtilsEx::SkipWhite(p);
		if ((*p != '=') && (*p != ':'))
		{
			if (pBackgroundTag)  // TODO: why don't we  return
			{
				p++;
				continue;
			}
		}

		bodyBackgroundColorAttrib.m_name.Append(pBgTag, bgTagLen);
		bodyBackgroundColorAttrib.m_separator = *p;

		p++;
		if ((p + 2) > e)
			break;

		p = TextUtilsEx::SkipWhite(p);
		BOOL valQuoted = FALSE;
		if (*p == '"')
		{
			bodyBackgroundColorAttrib.m_quoted = TRUE;
			valQuoted = TRUE;
			p++;
		}

		pBackgroundColorTagValBeg = p;

		if (valQuoted)
		{
			p = TextUtilsEx::findOneOf(p, e, "\"><");
			if (p == 0) 
				return 0;
			if (*p != '\"')
				return  0;
			pBackgroundColorTagValEnd = p;
			p++;
		}
		else
		{
			p = TextUtilsEx::findOneOf(p, e, "\" ;><");
			if (p == 0)
				return 0;
			if (*p == '<')
				return  0;
			pBackgroundColorTagValEnd = p;
		}

		pBgTagEnd = p;
		p = TextUtilsEx::SkipWhite(p);
		if (*p == ';')
		{
			bodyBackgroundColorAttrib.m_terminator = ';';
			p++;
			pBgTagEnd = p;
		}

		pBackgroundColorTagVal.Empty();
		if (pBackgroundColorTagValBeg && pBackgroundColorTagValEnd)
		{
			int tagLength = IntPtr2Int(pBackgroundColorTagValEnd - pBackgroundColorTagValBeg);
			char *b = pBackgroundColorTagVal.GetBufferSetLength(tagLength);
			::memcpy(b, pBackgroundColorTagValBeg, tagLength);
			pBackgroundColorTagVal.ReleaseBuffer();

			tagBeg = pBgTag;
			tagDataLen = IntPtr2Int(pBgTagEnd - pBgTag);

			if (pBackgroundTag)
			{
				if (pBackgroundColorTagVal.GetAt(0) != '#') // enhance in the future
				{
					return 0;
				}
				else
				{
					int deb = 1; // enhance in the future
				}
			}
			break;
		}
		p++;
	}
	if (tagBeg && (tagDataLen > 0))
	{
		bodyBackgroundColor = pBackgroundColorTagVal;
		bodyBackgroundColor.TrimLeft();
		bodyBackgroundColor.TrimRight();

		bodyBackgroundColorAttrib.m_value = bodyBackgroundColor;
		if (tagBeg)
			bodyBackgroundColorAttrib.m_attribString.Append(pBgTag, tagDataLen);
	}
	else
		bodyBackgroundColorAttrib.Reset();

	return tagDataLen;
}

// No longer used,  review and delete possibly.
// Replaced with NListView::RemoveBodyBackgroundColorAndWidth
int NListView::RemoveBodyBackgroundColor(char *inData, int indDataLen, SimpleString *outbuf, CStringA &bodyBackgroundColor)
{
	_ASSERTE(FALSE);
	return 1;
#if 0
	static char *bodyTag = "<body ";
	static int bodyTagLen = istrlen(bodyTag);

	static char *styleTag = " style";
	static int styleTagLen = istrlen(styleTag);

	static char *bgColorTag = "bgcolor";
	static int bgColorTagLen = istrlen(bgColorTag);

	static char *backgroundColorTag = "background-color";
	static int backgroundColorTagLen = istrlen(backgroundColorTag);

	static char *backgroundTag = "background";
	static int backgroundTagLen = istrlen(backgroundTag);

	static char *bodyTagEnd = ">";
	static int bodyTagEndLen = istrlen(bodyTagEnd);

	outbuf->ClearAndResize(indDataLen + 128);

	if (indDataLen < bodyTagLen)
		return 1;

	DWORD clr = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);
	if (clr == COLOR_WHITE)
		int deb = 1; // return 1;

	CString colorStr;

	int retC2S = NListView::Color2Str(clr, colorStr);

	char *p_mark;
	char *p = inData;
	int count = indDataLen;
	char *e = p + indDataLen;
	char *e_data = p + indDataLen;

	char * pBodyBeg = 0;
	char * pBodyEnd = 0;

	char *pStyleTag = 0;
	char *pBgColorTag = 0;
	char *pBackgroundColorTag = 0;
	char *pBackgroundTag = 0;
	char *pBgTag;

	char *pBackgroundColorTagBeg = 0;
	char *pBackgroundColorTagEnd = 0;

	// Replace background def in body only.  Limit search to body begin and body end tag
	{
		pBodyBeg = TextUtilsEx::findNoCaseP(p, count, bodyTag, bodyTagLen);
		if (pBodyBeg == 0)
			return 1;

		p = pBodyBeg + bodyTagLen;
		count = e - p;
		if (count <= 0)
			return 1;

		pBodyEnd = TextUtilsEx::findNoCaseP(p, count, bodyTagEnd, bodyTagEndLen);
		if (pBodyEnd == 0)
			return 1;

		e = pBodyEnd + 1;

		count = e - p;
		if (count <= 0)
			return 1;
	}

	p_mark = p;

	CString pBackgroundColorTagVal;
	while (p < e)  // always true
	{
		pStyleTag = 0;
		pBgColorTag = 0;
		pBackgroundColorTag = 0;
		pBackgroundTag = 0;
		pBgTag = 0;

		pBackgroundColorTagBeg = 0;
		pBackgroundColorTagEnd = 0;

		count = e - p;
		p = TextUtilsEx::findNoCaseP(p, count, "b", 1);
		if (p == 0)
			break;

		if (TextUtilsEx::strncmpUpper2Lower(p, e, backgroundColorTag, backgroundColorTagLen) == 0)
		{
			pBgTag = p;
			pBackgroundColorTag = p;
			p += backgroundColorTagLen;
		}
		else if (TextUtilsEx::strncmpUpper2Lower(p, e, bgColorTag, bgColorTagLen) == 0)
		{
			pBgTag = p;
			pBgColorTag = p;
			p += bgColorTagLen;
		}
#if 1
		else if (TextUtilsEx::strncmpUpper2Lower(p, e, backgroundTag, backgroundTagLen) == 0)
		{
			pBgTag = p;
			pBackgroundTag = p;
			p += backgroundTagLen;
		}

		if ((pBackgroundColorTag == 0) && (pBgColorTag == 0) && (pBackgroundTag == 0))
		{
			p++;
			continue;
		}
#else
		if ((pBackgroundColorTag == 0) && (pBgColorTag == 0))
		{
			p++;
			continue;
		}
#endif

		p = TextUtilsEx::SkipWhite(p);
		if ((*p != '=') && (*p != ':'))
		{
			if (pBackgroundTag)
			{
				p++;
				continue;
			}
		}

		p++;
		if ((p + 2) > e)
			break;

		p = TextUtilsEx::SkipWhite(p);
		if (*p == '"')
			p++;

		if (*p == '#')
			int deb = 1;// p++;

		pBackgroundColorTagBeg = p;

		p = TextUtilsEx::findOneOf(p, e, "\" ;>");
		if (p == 0)
		{
			outbuf->Clear();
			return 1;
		}

		pBackgroundColorTagEnd = p;

		pBackgroundColorTagVal.Empty();
		if (pBackgroundColorTagBeg && pBackgroundColorTagEnd)
		{
			int tagLength = pBackgroundColorTagEnd - pBackgroundColorTagBeg;
			char *b = pBackgroundColorTagVal.GetBufferSetLength(tagLength);
			::memcpy(b, pBackgroundColorTagBeg, tagLength);
			pBackgroundColorTagVal.ReleaseBuffer();

#if 1
			if (pBackgroundTag)
			{
				if (pBackgroundColorTagVal.GetAt(0) != '#')

				{
					return 0;
				}
				else
				{
					int deb = 1; // enhance in the future
				}
			}
#endif
			outbuf->Append(inData, pBgTag - inData);
			p_mark = pBackgroundColorTagEnd;
			outbuf->Append(p_mark, e_data - p_mark);
			break;
		}
		p++;
	}
	//outbuf->Append(p_mark, e_data - p_mark);

	bodyBackgroundColor = pBackgroundColorTagVal;

	return 1;
#endif
}

int NListView::SetBodyWidth(char *inData, int indDataLen, SimpleString *outbuf, CString &bodyWidth)
{
	static char *bodyTag = "<body ";
	static int bodyTagLen = istrlen(bodyTag);

	static char *widthTag = " width";
	static int widthTagLen = istrlen(widthTag);

	static char *widthTag2 = ";width";
	static int widthTag2Len = istrlen(widthTag2);

	static char *styleTag = " style";
	static int styleTagLen = istrlen(styleTag);

	static char *bgColorTag = "bgcolor";
	static int bgColorTagLen = istrlen(bgColorTag);

	static char *backgroundColorTag = "background-color";
	static int backgroundColorTagLen = istrlen(backgroundColorTag);

	static char *backgroundTag = "background";
	static int backgroundTagLen = istrlen(backgroundTag);

	static char *bodyTagEnd = ">";
	static int bodyTagEndLen = istrlen(bodyTagEnd);

	outbuf->ClearAndResize(indDataLen + 128);

	if (indDataLen < (bodyTagLen+24))
		return 1;

	char *p_mark;
	char *p = inData;
	int count = indDataLen;
	char *e = p + indDataLen;
	char *e_data = p + indDataLen;

	char * pBodyBeg = 0;
	char * pBodyEnd = 0;

	char *pStyleTag = 0;
	char *pWidthTag = 0;
	char *pBgTag;

	char *pWidthTagBeg = 0;
	char *pWidthTagEnd = 0;

	// Replace width def in body ony.  Limit search to body begin and body end
	{
		pBodyBeg = TextUtilsEx::findNoCaseP(p, count, bodyTag, bodyTagLen);
		if (pBodyBeg == 0)
			return 1;

		p = pBodyBeg + bodyTagLen;
		count = IntPtr2Int(e - p);
		if (count <= 0)
			return 1;

		pBodyEnd = TextUtilsEx::findNoCaseP(p, count, bodyTagEnd, bodyTagEndLen);
		if (pBodyEnd == 0)
			return 1;

		e = pBodyEnd + 1;

		count = IntPtr2Int(e - p);
		if (count <= 0)
			return 1;
	}

	//outbuf->Append(inData, p - inData);

	p_mark = p;

	CStringA pWidthTagVal;
	while (p < e)  // always true
	{
		pStyleTag = 0;
		pWidthTag = 0;

		pWidthTagBeg = 0;
		pWidthTagEnd = 0;

		count = IntPtr2Int(e - p);
		p = TextUtilsEx::findNoCaseP(p, count, "w", 1);
		if (p == 0)
			break;

		if ((TextUtilsEx::strncmpUpper2Lower(p-1, e, widthTag, widthTagLen) == 0) ||
			(TextUtilsEx::strncmpUpper2Lower(p-1, e, widthTag2, widthTag2Len) == 0))
		{
			pBgTag = p;
			pWidthTag = p;
			p += widthTagLen -1;
		}
		if (pWidthTag == 0)
		{
			p++;
			continue;
		}

		p = TextUtilsEx::SkipWhite(p);
		if ((*p != '=') && (*p != ':'))
		{
			if (pWidthTag)
			{
				p++;
				continue;
			}
		}

		p++;
		if ((p + 2) > e)
			break;

		p = TextUtilsEx::SkipWhite(p);
		if (*p == '"')
			p++;

		if (*p == '#')
			int deb = 1;// p++;

		pWidthTagBeg = p;

		p = TextUtilsEx::findOneOf(p, e, "\" ;>");
		if (p == 0)
		{
			outbuf->Clear();
			return 1;
		}

		pWidthTagEnd = p;

		pWidthTagVal.Empty();
		if (pWidthTagBeg && pWidthTagEnd)
		{
			int tagLength = IntPtr2Int(pWidthTagEnd - pWidthTagBeg);
			char *b = pWidthTagVal.GetBufferSetLength(tagLength);
			::memcpy(b, pWidthTagBeg, tagLength);
			pWidthTagVal.ReleaseBuffer();

#if 0
			if (pWidthTag)
			{
				if (pWidthTagVal.GetAt(0) != '#')

				{
					return 1;
				}
				else
				{
					int deb = 1; // enhance in the future
				}
			}
#endif
			outbuf->Append(inData, IntPtr2Int(pWidthTag - inData));
			p_mark = pWidthTagEnd;
			outbuf->Append(p_mark, IntPtr2Int(e_data - p_mark));
			break;
		}
		p++;
	}
	//outbuf->Append(p_mark, e_data - p_mark);

	bodyWidth = pWidthTagVal;
	return 1;
}

int NListView::Color2Str(DWORD color, CStringA &colorStr)
{
	COLORREF cr = color;
	RGBTRIPLE rgb;

	rgb.rgbtRed = GetRValue(cr);
	rgb.rgbtGreen = GetGValue(cr);
	rgb.rgbtBlue = GetBValue(cr);

	colorStr.Format("%02x%02x%02x", rgb.rgbtRed, rgb.rgbtGreen, rgb.rgbtBlue);
	//colorStr.Format("0x%06x", color);
	return 1;
}

int NListView::Color2Str(DWORD color, CString& colorStr)
{
	COLORREF cr = color;
	RGBTRIPLE rgb;

	rgb.rgbtRed = GetRValue(cr);
	rgb.rgbtGreen = GetGValue(cr);
	rgb.rgbtBlue = GetBValue(cr);

	colorStr.Format(L"%02x%02x%02x", rgb.rgbtRed, rgb.rgbtGreen, rgb.rgbtBlue);
	//colorStr.Format("0x%06x", color);
	return 1;
}

int NListView::SaveAsEmlFile(CStringA &bdy)
{
	int retval = SaveAsEmlFile((char*)(LPCSTR)bdy, bdy.GetLength());
	return retval;
}

int NListView::SaveAsEmlFile(char *bdy, int bdylen)
{
	static const char *cFromMailBegin = "From ";
	static const int cFromMailBeginLen = istrlen(cFromMailBegin);
	// Save mail
	CFile fp;
	CFileException ExError;
	CString emlFile = CMainFrame::GetMboxviewTempPath() + L"mime-message.eml";
	if (fp.Open(emlFile, CFile::modeWrite | CFile::modeCreate, &ExError))
	{
		char *pb = bdy;
		int len = bdylen;
		char *e = pb + len;
		char *p = MimeParser::SkipEmptyLines(pb, e);
		if (TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0)
		{
			p = MimeParser::EatNewLine(p, e);
		}
		int eml_len = len - IntPtr2Int(p - pb);
		fp.Write(p, eml_len);
		fp.Close();
	}
	else
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not create \"" + emlFile;
		txt += L"\" file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);
	}
	return  1;
}

BOOL NListView::MatchIfFieldFolded(int mailPosition, char *fldName) // FIXME document
{
	const int maxBodyLength = 128 * 1024;
	static SimpleString mBody(maxBodyLength);

	if ((mailPosition >= MboxMail::s_mails.GetCount()) || (mailPosition < 0))
		return -1;

	MboxMail *m = MboxMail::s_mails[mailPosition];
	if (m == 0)
	{
		// TODO: Assert ??
		return 1;
	}

	mBody.Clear();
	BOOL ret = m->GetBodySS(&mBody, maxBodyLength);

	char *data = mBody.Data();
	int datalen = mBody.Count();
	char *p = data;
	char *e = data + datalen;

	int fldNameLength = istrlen(fldName);

	while (p < e)
	{
		if ((*p == '\n') || ((*p == '\r') && (*(p + 1) == '\n')))
		{
			break;
		}
		else if (TextUtilsEx::strncmpExact(p, e, fldName, fldNameLength) == 0)
		{
			p = MimeParser::EatNewLine(p, e);
			if ((*p == ' ') || (*p == '\t'))
				return TRUE;
			else
				p = MimeParser::EatNewLine(p, e);
		}
		else
		{
			p = MimeParser::EatNewLine(p, e);
		}
	}
	return FALSE;
}

int NListView::DetermineListFileName(CString &fileName, CString &listFileName)
{
	CString mailFile = MboxMail::s_path;

	CString datapath = MboxMail::GetLastDataPath();
	if (datapath.IsEmpty())
		return -1;

	CString fileNameBase;
	CString fileNameExtention;
	FileUtils::GetFileBaseNameAndExtension(fileName, fileNameBase, fileNameExtention);

	//CString mboxListFile = path + "ArchiveCache" + "\\"  + fileNameBase + "\\" + fileNameBase + "_USER" + fileNameExtention + ".mboxlist";
	CString mboxListFile = datapath + L"ListCache" + L"\\" + fileNameBase + L"\\" + fileNameBase + L"_USER.mbox.mboxlist";
	if (FileUtils::PathFileExist(mboxListFile))
	{
		listFileName.Append(mboxListFile);
	}
	return 1;
}

CWnd * NListView::SetListFocus()
{
	m_list.SetFocus();
	return &m_list;
}

void AttachmentMgr::Clear()
{
	m_attachmentMap.RemoveAll();
}

AttachmentData* AttachmentMgr::GetAttachmentData(const CStringW& inName)
{
	AttachmentDB::CPair* infopair = m_attachmentMap.PLookup(inName);
	if (infopair != nullptr)
		return &infopair->value;
	else
		return 0;
}

int AttachmentMgr::HasName(CStringW& inNameW)
{
	AttachmentDB::CPair* infopair = m_attachmentMap.PLookup(inNameW);
	if (infopair != nullptr)
		return TRUE;
	else
		return FALSE;
}

void AttachmentMgr::AddName(CStringW& inNameW, BOOL isEmbedded)
{
	AttachmentData item;
	item.m_nameW.Append(inNameW);
	item.m_isEmbedded = isEmbedded;
	m_attachmentMap[inNameW] = item;
}

// Add name if not there and return -1
// otherwise return >= 0
int AttachmentMgr::InsertName(CStringW &inNameW, BOOL isEmbedded)
{
	int nextId = -1;

	AttachmentDB::CPair *infopair = m_attachmentMap.PLookup(inNameW);
	if (infopair != nullptr)
	{
		AttachmentData &item = infopair->value;
		return item.m_nextId++;
	}
	else
	{
		AttachmentData item;
		item.m_nameW.Append(inNameW);
		item.m_isEmbedded = isEmbedded;
		m_attachmentMap[inNameW] = item;
		return -1;
	}
}

int __cdecl MailAttachmentPred(void const* first, void const* second)
{
	AttachmentData* f = ((AttachmentData*)first);
	AttachmentData* s = ((AttachmentData*)second);
	int ret = _wcsicmp(f->m_nameW.operator LPCWSTR(), s->m_nameW.operator LPCWSTR());
	return ret;
}

void AttachmentMgr::Sort()
{
	int i = 0;
	m_attachmentArray.SetSize(m_attachmentMap.GetCount());
	AttachmentDB::CPair* pCurVal;
	pCurVal = m_attachmentMap.PGetFirstAssoc();
	while (pCurVal != NULL)
	{
		// m_attachmentArray[i++] = pCurVal->value; //  FIXME
		m_attachmentArray.SetAtGrow(i++, pCurVal->value);  // extra check not really needed

		pCurVal = m_attachmentMap.PGetNextAssoc(pCurVal);
	}

	std::qsort(m_attachmentArray.GetData(), m_attachmentArray.GetCount(), sizeof(AttachmentData), MailAttachmentPred);
}

void AttachmentMgr::PrintMap(CString &title)
{
#ifdef _DEBUG
	TRACE(L"%s", title);

	AttachmentDB::CPair* pCurVal = m_attachmentMap.PGetFirstAssoc();
	int j = 0;
	while (pCurVal != NULL)
	{
		TRACE(L"attachmentDB:PrintMap:  %d value \"%s\" %d\n",
			j, pCurVal->value.m_nameW, pCurVal->value.m_isEmbedded);

		pCurVal = m_attachmentMap.PGetNextAssoc(pCurVal);
		j++;
	}
#endif
}

void AttachmentMgr::PrintArray(CString& title)
{
#ifdef _DEBUG
	TRACE(L"%s", title);
	int i;
	for (i = 0; i < m_attachmentArray.GetCount(); i++)
	{
		AttachmentData* item = &m_attachmentArray[i];

		TRACE(L"attachmentDB:PrintArray: %d value \"%s\" %d\n",
			i, item->m_nameW, item->m_isEmbedded);
	}
#endif
}

int NListView::ForwardSingleMail(int iItem, BOOL progressBar, CString &progressText, CString &errorText)
{
	CString htmFileName;
	//CString errorText;
	//BOOL progressBar;
	//CString progressText;

	CFile fpm;
	CFileException ExError;
	if (!fpm.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not open \"" + MboxMail::s_path;
		txt += L"\" mail file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);

		errorText = txt;
		return -1;
	}

	int mailPosition = iItem;
	CString emlFile;
	CString targetDirectory = CMainFrame::GetMboxviewLocalAppDataPath(L"MailService");

	int rexport = NListView::ExportAsEmlFile(&fpm, mailPosition, targetDirectory, emlFile, errorText);
	if (rexport < 0) {
		if (errorText.IsEmpty())
			MboxMail::assert_unexpected();
		return -1;
	}

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	
	CString password;
	DWORD64 id = (DWORD64)(time(0));
	CString instanceId;
	instanceId.Format(L"%lld", id);
	
	if (pFrame)
	{
		password.Format(L"%s:%s", instanceId, pFrame->m_mailDB.SMTPConfig.UserPassword);
	}

	// Try max 10 times
	int i;
	for (i = 0; i < 10; i++)
	{
		INT64 retval = NListView::ExecCommand_WorkerThread(m_tcpPort, instanceId, password, m_ForwardMailData, emlFile, errorText, progressBar, progressText);

		if (retval == 0)
			break;

		if (retval == -13) // PortAlreadyInUse
		{
			srand((unsigned int)time(0));
			
			int range = 65533 - 49152;
			m_tcpPort = 49152 + rand()%range;  // next port to try
		}
		else
			return -1;
	}
	if (i == 10) {
		//ForwardEmlFile.exe should create error file by now
		return -1;
	}
	else
		return 1;
}

int NListView::ForwardMailDialog(int iItem)
{
	DWORD error;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame == 0) {
		//internal error;
		return -1;
	}

	if (pFrame->m_mailDB.SMTPConfig.UserAccount.IsEmpty() || pFrame->m_mailDB.SMTPConfig.UserPassword.IsEmpty())
	{
#if 0
		CString txt = L"User name and/or password are not provided.\n";
		txt += L"Select \"File\"->\"SMTP Mail Server Config\" and provide user name and password.\n";
#endif

		CString txt = L"User name and/or password are not provided.\n"
			"Select \"File\"->\"SMTP Mail Server Config\" and provide user name and password.";

		ResHelper::TranslateString(txt);

		HWND h = GetSafeHwnd(); // we don't have any window yet
		int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	m_ForwardMailData.m_MailService = pFrame->m_mailDB.SMTPConfig.MailServiceName;
	if (pFrame && (m_ForwardMailData.m_MailService.CompareNoCase(L"Custom") == 0))
	{
		if (m_ForwardMailData.m_From.IsEmpty())
		{
			if (!pFrame->m_mailDB.SMTPConfig.UserMailAddress.IsEmpty())
				m_ForwardMailData.m_From = pFrame->m_mailDB.SMTPConfig.UserMailAddress;
			else
				m_ForwardMailData.m_From = pFrame->m_mailDB.SMTPConfig.UserAccount;
		}
	}
	else
		m_ForwardMailData.m_From = pFrame->m_mailDB.SMTPConfig.UserAccount;

	ForwardMailDlg dlg;

	CStringW subjectW;
	if (iItem >= 0)
	{
		MboxMail *m = MboxMail::s_mails[iItem];
		BOOL ret = TextUtilsEx::Str2WStr(m->m_subj, m->m_subj_charsetId, subjectW, error);
	}

	m_ForwardMailData.m_subjectW = subjectW;

	dlg.m_Data.Copy(m_ForwardMailData);

	if (dlg.DoModal() == IDOK)
	{
		m_ForwardMailData.Copy(dlg.m_Data);
		if (pFrame && (m_ForwardMailData.m_MailService.CompareNoCase(L"Custom") == 0))
		{
			if (pFrame->m_mailDB.SMTPConfig.UserMailAddress.Compare(m_ForwardMailData.m_From) != 0)
			{
				pFrame->m_mailDB.SMTPConfig.UserMailAddress = m_ForwardMailData.m_From;
				pFrame->m_mailDB.CustomSMTPConfig.Copy(pFrame->m_mailDB.SMTPConfig);
			}
		}
		return 1;
	}
	else
	{
		return -1;
	}
}

// Forward related mails aprt of the same conversation
int NListView::ForwardMailRange(int iSelectedItem)
{
	CString errorText;
	CString targetPrintSubFolderName;
	CString targetPrintFolderPath;

	if (!((abs(MboxMail::b_mails_which_sorted) == 99) || ((abs(MboxMail::b_mails_which_sorted) == 4) && (m_subjectSortType == 1))))
	{
		//if (abs(MboxMail::b_mails_which_sorted) != 99)
		{

			CString txt = L"Please sort all mails by conversation threads or subject threads first.\n";
			//txt += "Select \"View\"->\"Sort By\" ->\"Conversation\" or left click on the first column.";
			ResHelper::TranslateString(txt);


			HWND h = GetSafeHwnd(); // we don't have any window yet
			int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
			//return -1;
		}
		return -1;
	}

	int firstMail = iSelectedItem;
	int lastMail = iSelectedItem;
	FindFirstAndLastMailOfConversation(iSelectedItem, firstMail, lastMail);

	int selectedCnt = lastMail - firstMail + 1;
	if (selectedCnt > 100)
	{
		CString txt = L"Mail sending limit imposed by Mail Service is not well defined.\n"
			L"Number of mails you selected to forward is greater than 100\n"
			L"Mail Service imposed sending limit may or not be exceeded and result in your mail account being suspended for 24 hurs.\n\n"
			L"Do you want to forward selected mails anyway?\n";

		ResHelper::TranslateString(txt);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer != IDYES)
		{
			return -1;
		}
	}
	m_selectedMailsList.SetSize(0);
	m_selectedMailsList.SetSize(selectedCnt);

	int i = 0;
	int j = 0;
	for (i = firstMail; i < lastMail; i++, j++)
	{
		m_selectedMailsList.SetAtGrow(j, i); // to be safe grow just in case
	}

	// Function please
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		int MaxMailSize = pFrame->m_mailDB.SMTPConfig.MaxMailSize * 1024;
		int largeMailsCnt = 0;
		int i;
		int cnt = (int)m_selectedMailsList.GetCount();
		for (int j = 0; j < cnt; j++)
		{
			i = m_selectedMailsList[j];
			MboxMail *m = MboxMail::s_mails[i];
			if (m->m_length > MaxMailSize)
				largeMailsCnt++;
		}

		if (largeMailsCnt)
		{
			CString fmt = L"Found %d mails in the selected list that are larger than the max mail size of %d KB imposed by the active service %s. Sending will fail.\n"
				L"Size of each mail is shown in the size column in the Summary Pane.\n\n"
				L"Do you want to send anyway?\n";
			
			ResHelper::TranslateString(fmt);
			CString txt;
			txt.Format(fmt, largeMailsCnt, MaxMailSize, pFrame->m_mailDB.SMTPConfig.MailServiceName);


			int answer = MessageBox(txt, L"Warning", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
			if (answer != IDYES)
			{
				return -1;
			}
		}
	}

	MailIndexList *selectedMailsIndexList = &m_selectedMailsList;
	int iItem = (*selectedMailsIndexList)[0];

	int rval = ForwardMailDialog(iItem);
	if (rval <= 0)
		return -1;

	int ret = ForwardSelectedMails_Thread(&m_selectedMailsList, targetPrintSubFolderName);
	if (ret < 0)
		return -1;
	else
		return 1;
}

int NListView::ForwardSelectedMails(int iSelectedItem)
{
	CString errorText;
	CString targetPrintSubFolderName;
	CString targetPrintFolderPath;

	if (PopulateSelectedMailsList() <= 0)
		return 0;

	int selectedCnt = (int)m_selectedMailsList.GetCount();
	if (selectedCnt > 100)
	{
		CString txt = L"Mail sending limit imposed by Mail Service is not well defined.\n"
			L"Number of mails you selected to forward is greater than 100\n"
			L"Mail Service imposed sending limit may or not be exceeded and result in your mail account being suspended for 24 hurs.\n"
			L"You make take manual action to unblock the accout without delay\n\n"
			L"Do you want to forward selected emails anyway?\n";

		ResHelper::TranslateString(txt);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer != IDYES)
		{
			return -1;
		}
	}

	// Need function ??
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		int MaxMailSize = pFrame->m_mailDB.SMTPConfig.MaxMailSize * 1024;
		int largeMailsCnt = 0;
		int i;
		int cnt = (int)m_selectedMailsList.GetCount();
		for (int j = 0; j < cnt; j++)
		{
			i = m_selectedMailsList[j];
			MboxMail *m = MboxMail::s_mails[i];
			if (m->m_length > MaxMailSize)
				largeMailsCnt++;
		}

		if (largeMailsCnt)
		{
			CString fmt = L"Found %d mails in the selected list that are larger than the max mail size of %d KB imposed by the active service %s. Sending will fail.\n"
				L"Size of each mail is shown in the size column in the Summary Pane.\n\n"
				L"Do you want to send anyway?\n";

			ResHelper::TranslateString(fmt);
			CString txt;
			txt.Format(fmt, largeMailsCnt, MaxMailSize, pFrame->m_mailDB.SMTPConfig.MailServiceName);

			int answer = MessageBox(txt, L"Warning", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
			if (answer != IDYES)
			{
				return -1;
			}
		}
	}

	MailIndexList *selectedMailsIndexList = &m_selectedMailsList;
	int iItem = (*selectedMailsIndexList)[0];

	int rval = ForwardMailDialog(iItem);
	if (rval <= 0)
		return -1;  // user cancelled the sending

	int ret = ForwardSelectedMails_Thread(selectedMailsIndexList, targetPrintSubFolderName);

	return 1;
}

int NListView::ForwardMails_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName)
{
	m_selectedMailsList.SetSize(0);
	int selectedCnt = lastMail - firstMail;
	m_selectedMailsList.SetSize(selectedCnt);

	int i = 0;
	int j = 0;
	for (i = firstMail; i < lastMail; i++, j++)
	{
		m_selectedMailsList.SetAtGrow(j, i); // to be safe grow just in case
	}
	if (m_selectedMailsList.GetCount() > 100)
		int deb = 1;

	int ret = NListView::ForwardSelectedMails_Thread(&m_selectedMailsList, targetPrintSubFolderName);
	return ret;
}

// FIXME review  forward feature
int NListView::ForwardSelectedMails_Thread(MailIndexList *selectedMailsIndexList, CString &targetPrintSubFolderName)
{
	FORWARD_MAILS_ARGS args;

	//CString rootPrintSubFolder = "PrintCache";
	//CString targetPrintSubFolder = targetPrintSubFolderName;
	//CString printCachePath;
	CString errorText;
	CString forwardEmlFileExePath;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		int ret = VerifyPathToForwardEmlFileExecutable(forwardEmlFileExePath, errorText);
		if (ret < 0)
		{
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else
	{
		HWND h = GetSafeHwnd();
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		errorText.Append(L"Internal error. Try again.");
		return -1;
	}

	/*IN*/
	args.forwardMailsData.Copy(m_ForwardMailData);

	DWORD error;
	CStringA result;
	BOOL rval = TextUtilsEx::WStr2UTF8(&args.forwardMailsData.m_Text, &result, error);

	CString appDataPath = CMainFrame::GetMboxviewLocalAppDataPath(L"MailService");
	CString cStrNamePath = appDataPath + L"MailText.txt";
	BOOL wret = FileUtils::Write2File(cStrNamePath, (unsigned char*)((LPCSTR)result), result.GetLength());

	args.forwardMailsData.m_Text.Empty();

	args.password = pFrame->m_mailDB.SMTPConfig.UserPassword;
	args.lview = this;
	args.targetPrintFolderPath = L""; // printCachePath;
	args.targetPrintSubFolderName = L""; // targetPrintSubFolder;
	args.firstMail = -1;
	args.lastMail = -1;
	args.selectedMailIndexList = selectedMailsIndexList;
	args.nItem = -1;
	args.separatePDFs = TRUE;

	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcForwardMails, (LPVOID)(FORWARD_MAILS_ARGS*)&args);
	Dlg.SetDialogTemplate(AfxGetApp()->m_hInstance, MAKEINTRESOURCE(IDD_PROGRESS_DLG), IDC_STATIC, IDC_PROGRESS_BAR, IDCANCEL);

	INT_PTR nResult = Dlg.DoModal();

	if (!nResult) { // should never be true ?
		MboxMail::assert_unexpected();
		return -1;
	}

	int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
	int retResult = LOWORD(nResult);

	if (retResult != IDOK)
	{  
		// IDOK==1, IDCANCEL==2
		// We should be here when user selects Cancel button
		//_ASSERTE(cancelledbyUser == TRUE);

		DWORD terminationDelay = Dlg.GetTerminationDelay();
		int loopCnt = (terminationDelay+100)/25;

		ULONGLONG tc_start = GetTickCount64();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		ULONGLONG tc_end = GetTickCount64();
		DWORD delta = (DWORD)(tc_end - tc_start);
		TRACE(L"(ForwardSelectedMails_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret == -2)
		return 1; // IDCANCEL

	CString ForwardMailOSuccessFilePath = appDataPath + L"ForwardMailSuccess.txt";
	CString ForwardMailErrorFilePath = appDataPath + L"ForwardMailError.txt";
	BOOL ForwardMailOSuccess = FALSE;
	BOOL ForwardMailError = FALSE;

	if (FileUtils::PathFileExist(ForwardMailOSuccessFilePath))
		ForwardMailOSuccess = TRUE;

	SimpleString errorTxt;
	if (FileUtils::PathFileExist(ForwardMailErrorFilePath))
	{
		BOOL rval = FileUtils::ReadEntireFile(ForwardMailErrorFilePath, errorTxt);
		ForwardMailError = TRUE;
	}

	if ((ForwardMailOSuccess == TRUE) && (ForwardMailError == TRUE))
	{
		// Should nver happen -:)
		MboxMail::assert_unexpected();
	}

	if ((ret == 0) && (ForwardMailError == TRUE))
		MboxMail::assert_unexpected();

	if (ForwardMailError == TRUE)
	{
		UINT strCodePage = CP_UTF8;
		SimpleString errorTxtW;
		DWORD error;

		BOOL ret = TextUtilsEx::CodePage2WStr(&errorTxt, strCodePage, &errorTxtW, error);


		HWND h = GetSafeHwnd();
		// Both statements work
		int answer = ::MessageBox(h, (LPCWSTR)errorTxtW.Data(), L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		//int answer = ::MessageBoxA(h, (LPCSTR)errorTxt.Data(), "Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	// We rely on errorTxt; we are not consistent
	if (ret > 0)
	{
		if (args.errorText.IsEmpty())
		{
			int deb = 1;
		}
		else
		{
			MboxMail::assert_unexpected();
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			MboxMail::assert_unexpected();
			return -1;
		}
	}
	else if (ret < 0) 
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, args.errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}
	return 1;
}

int NListView::DeleteForwardEmlFiles(CString &targetPath)
{
	CString filePath;
	BOOL delStatus;
	int fileCnt = 0;
	CString fn;

	CString fw = targetPath + L"\\*.eml";
	WIN32_FIND_DATA	wf;
	HANDLE f = FindFirstFile(fw, &wf);
	if (f != INVALID_HANDLE_VALUE)
	{
		do
		{
			if ((wf.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == FILE_ATTRIBUTE_DIRECTORY)
				continue;
			fileCnt++;
			fn = wf.cFileName;
			filePath = targetPath + L"\\" + fn;
			delStatus = FileUtils::DelFile(filePath);
			if (delStatus == FALSE) {
				DWORD error = GetLastError();
			}

		} while (FindNextFile(f, &wf));
		FindClose(f);
	}

	return fileCnt;
}

int NListView::ForwardMails_WorkerThread(ForwardMailData &mailData, MailIndexList *selectedMailIndexList, CString &errorText)
{
	BOOL progressBar = TRUE;

	CString appDataPath = CMainFrame::GetMboxviewLocalAppDataPath(L"MailService");
	int retdel = NListView::DeleteForwardEmlFiles(appDataPath);

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (selectedMailIndexList->GetCount() <= 0)
		return 1;

	UINT curstep = 1;
	CString fileNum;
	int nFileNum;

	if (progressBar && MboxMail::pCUPDUPData)
	{
		//fileNum.Format(L"Forwarding mails to %s  ... ", mailData.m_MailService);
		;// MboxMail::pCUPDUPData->SetProgress(fileNum, 1);
	}

	if (progressBar && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	int i;
	int cnt = (int)selectedMailIndexList->GetCount();

	ULONGLONG workRangeFirstPos = 0;
	ULONGLONG workRangeLastPos = cnt - 1;
	ProgressTimer progressTimer(workRangeFirstPos, workRangeLastPos);

	BOOL progress_Bar = TRUE;
	for (int j = 0; j < cnt; j++)
	{
		i = (*selectedMailIndexList)[j];
		if (progressBar && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				break;
			}

			UINT_PTR dwProgressbarPos = 0;
			ULONGLONG workRangePos = j;
			BOOL needToUpdateStatusBar = progressTimer.UpdateWorkPos(workRangePos, dwProgressbarPos);
			if (needToUpdateStatusBar)
			{
				nFileNum = j + 1;

				CString fmt = L"Forwarding mails to %s  ... %d of %d";
				ResHelper::TranslateString(fmt);
				fileNum.Format(fmt, mailData.m_MailService, nFileNum, cnt);

				if (MboxMail::pCUPDUPData)  MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(dwProgressbarPos));
				int debug = 1;
			}
		}
		int retval = ForwardSingleMail(i, progress_Bar, fileNum, errorText);
		if (retval < 0)
			return -1;
	}
	nFileNum = cnt;
	UINT newstep = 100;

	CString fmt = L"Forwarding mails to %s  ... %d of %d";
	ResHelper::TranslateString(fmt);
	fileNum.Format(fmt, mailData.m_MailService, nFileNum, cnt);

	if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(newstep));
	int deb = 1;
	return 1;
}

CString NListView::FixCommandLineArgument(CString &in)
{
	CString out;
	CString arg(in);
	arg.Replace(L"\"", L"\"\"");
	out = L"\"" + arg + L"\"";
	return out;
}

CString NListView::FixCommandLineArgument(int in)
{
	CString out;
	out.Format(L"\"%d\"", in);
	return out;
}

INT64 NListView::ExecCommand_WorkerThread(int tcpPort, CString instanceId, CString &password, ForwardMailData &mailData, CString &emlFile, 
	CString &errorText, BOOL progressBar, CString &progressText, int timeout)
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (AfxSocketInit() == FALSE)
	{
		errorText.Append(L"Sockets Could Not Be Initialized");
		return -1;
	}

	//CString instanceIdW, CString& password, ForwardMailData& mailData, CString& emlFile,

	// TODO: Duplicate check, done already in _Thread function calls 
	CString forwardEmlFileExePath;
	if (pFrame)
	{
		int ret = VerifyPathToForwardEmlFileExecutable(forwardEmlFileExePath, errorText);
		if (ret < 0)
		{
			if (errorText.IsEmpty())
				MboxMail::assert_unexpected();
			return -1;
		}
	}
	else
	{
		errorText.Append(L"Internal error. Try again.");
		return -1;
	}

	CString appDataPath = CMainFrame::GetMboxviewLocalAppDataPath(L"MailService");
	CString smtpConfigFilePath = appDataPath + L"SMTP.ini";
	CString smtProtocolFilePath = appDataPath + L"SMTPLog.txt";
	CString loggerFilePath = appDataPath +  L"ForwardMailLog.txt";
	CString textFilePath = appDataPath + L"MailText.txt";

	CString args;
	args = args + L" --smtp-cnf " + FixCommandLineArgument(smtpConfigFilePath);

	if (mailData.m_MailService.CompareNoCase(L"Custom") == 0)
	{
		args = args + L" --from " + FixCommandLineArgument(mailData.m_From);
	}

	args = args + " --to " + FixCommandLineArgument(mailData.m_To);
	if (!mailData.m_CC.IsEmpty())
		args = args + L" --cc " + FixCommandLineArgument(mailData.m_CC);
	if (!mailData.m_BCC.IsEmpty())
		args = args + L" --bcc " + FixCommandLineArgument(mailData.m_BCC);

	// ForwardEmlFile.exe  program will extract subject from eml file
	args = args + L" --eml-file " + FixCommandLineArgument(emlFile);
	if (!textFilePath.IsEmpty())
		args = args + L" --mail-text-file " + FixCommandLineArgument(textFilePath);
	args = args + L" --tcp-port " + FixCommandLineArgument(tcpPort);
	if (!instanceId.IsEmpty())
		args = args + L" --instance-id " + FixCommandLineArgument(instanceId);
	//
	if (m_enableForwardMailsLog)
	{
		if (!loggerFilePath.IsEmpty())
			args = args + L" --logger-file " + FixCommandLineArgument(loggerFilePath);
	}
	if (m_enableSMTPProtocolLog)
	{
		if (!smtProtocolFilePath.IsEmpty())
			args = args + L" --smtp-protocol-logger " + FixCommandLineArgument(smtProtocolFilePath);
	}

	CStringA argsA;
	DWORD error;   // FIXME handle errors
	BOOL retW2A;

	retW2A = TextUtilsEx::WStr2Ansi(args, argsA, error);

	// FIXMEFIXME  enable below for debugging ForwardEmlFile.exe
	// forwardEmlFileExePath = "F:\\Documents\\GIT1.0.3.38\\mboxviewer\\ForwardEmlFile\\bin\\Debug\\netcoreapp3.1\\ForwardEmlFile.exe";

	HINSTANCE result = S_OK;
	SHELLEXECUTEINFO ShExecInfo = { 0 };
	ShExecInfo.cbSize = sizeof(SHELLEXECUTEINFO);
	ShExecInfo.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NO_CONSOLE;
	ShExecInfo.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_FLAG_NO_UI;
	ShExecInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
	ShExecInfo.hwnd = NULL;
	ShExecInfo.lpVerb = NULL;
	ShExecInfo.lpFile = forwardEmlFileExePath;
	ShExecInfo.lpParameters = args;
	ShExecInfo.lpDirectory = NULL;
	ShExecInfo.nShow = SW_HIDE;
	ShExecInfo.hInstApp = result;
	BOOL retval = ShellExecuteEx(&ShExecInfo);
	if (retval == FALSE)
	{
		DWORD err = GetLastError();
		int ret = CMainFrame::CheckShellExecuteResult(ShExecInfo.hInstApp, errorText);
		if (errorText.IsEmpty())
			MboxMail::assert_unexpected();
		return -1;
	}

	DWORD procId = GetProcessId(ShExecInfo.hProcess);
	CString processName(L"ForwardEmlFile.exe");
	CString errText;

	int step = 10;
	int stepCnt = 0;
	int nSeconds = 0;
	CString secondsBar;

	DWORD msec = 100;
	BOOL failed = FALSE;
	BOOL passwordAlreadySend = FALSE;
	int delaySendingPasswod = 20;
	int timeout_cnt = 0;
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
		case WAIT_OBJECT_0:
		{
			INT64 exitcode = -1;
			DWORD dwExitStatus = 0;
			if (ShExecInfo.hProcess)
			{
				BOOL retval = GetExitCodeProcess(ShExecInfo.hProcess, &dwExitStatus);
				exitcode = (int)dwExitStatus;
				CloseHandle(ShExecInfo.hProcess);
				//  ??
				dwExitStatus = GetLastError();
				CString errorText = FileUtils::GetLastErrorAsString();
			}
			return exitcode;
		}
		case WAIT_FAILED: {
			failed = TRUE;
			break;
		}
		case WAIT_TIMEOUT: 
		{
			if (delaySendingPasswod-- > 0)
			{
				timeout_cnt++;
			}
			else if (!passwordAlreadySend)
			{
				// FIXME Do we wait 2 seconds before sending password to throttle sending ?? Max 30 mails per minute possibly less ??
				// 
				// FIX: Make client non blocking
				MyTcpClient tcpClient(tcpPort);


				//CString errorText;
				CStringA passwordA;
				DWORD error;   // FIXME handle errors
				BOOL retW2A;
				retW2A = TextUtilsEx::WStr2Ansi(password, passwordA, error);

				int retConnect = tcpClient.ConnectSendClose(passwordA, errorText);
				if (retConnect == 0)
					passwordAlreadySend = TRUE;
			}
			if (MboxMail::pCUPDUPData)
			{
				if (MboxMail::pCUPDUPData->ShouldTerminate())
				{
					if (ShExecInfo.hProcess)
					{
						TerminateProcess(ShExecInfo.hProcess, IDCANCEL);
						CloseHandle(ShExecInfo.hProcess);
					}
					else
						INT64 retstatus = NListView::ExecCommand_KillProcess(processName, errText, FALSE, progressText);
					if (errorText.IsEmpty())
						MboxMail::assert_unexpected();
					return -1;
				}
				if (progressBar)
				{
					if (stepCnt % 10 == 0)
					{
						nSeconds++;
						if (progressText.IsEmpty())
							MboxMail::pCUPDUPData->SetProgress(step);
						else {
							CString fmt = L"%s     %d seconds";
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
		if (failed)
		{
			break;
		}
	}

	DWORD dwExitStatus = GetLastError();
	errorText = FileUtils::GetLastErrorAsString();

	INT64 retstatus = NListView::ExecCommand_KillProcess(processName, errText, FALSE, progressText);
	return -1;
}

INT64 NListView::ExecCommand_KillProcess(CString processName, CString &errorText, BOOL progressBar, CString &progressText)
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	HINSTANCE result = S_OK;
	SHELLEXECUTEINFO ShExecInfo = { 0 };
	ShExecInfo.cbSize = sizeof(SHELLEXECUTEINFO);
	ShExecInfo.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NO_CONSOLE;
	ShExecInfo.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_FLAG_NO_UI;
	ShExecInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
	ShExecInfo.hwnd = NULL;
	ShExecInfo.lpVerb = NULL;
	ShExecInfo.lpFile = L"C:\\Windows\\System32\\taskkill.exe";
	ShExecInfo.lpParameters = L"/F /IM" + processName + L"/T";
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

	DWORD procId = GetProcessId(ShExecInfo.hProcess);

	int step = 10;
	int stepCnt = 0;
	int nSeconds = 0;
	CString secondsBar;

	DWORD msec = 100;
	BOOL failed = FALSE;
	int waitTimeout = 30; // don't wait forever
	for (;;)
	{
		msec = 100;
		DWORD ret = WaitForSingleObject(ShExecInfo.hProcess, msec);
		switch (ret)
		{
		case WAIT_OBJECT_0: {
			INT64 exitcode = -1;
			DWORD dwExitStatus = 0;
			if (ShExecInfo.hProcess)
			{
				BOOL retval = GetExitCodeProcess(ShExecInfo.hProcess, &dwExitStatus);
				exitcode = (int)dwExitStatus;
				CloseHandle(ShExecInfo.hProcess);
				//  ??
				dwExitStatus = GetLastError();
				CString errorText = FileUtils::GetLastErrorAsString();
			}
			return exitcode;
		}
		case WAIT_TIMEOUT: 
		{
			if (waitTimeout-- < 0)
			{
				if (ShExecInfo.hProcess)
				{
					CloseHandle(ShExecInfo.hProcess);
				}
				failed = TRUE;
				break;
			}
			if (MboxMail::pCUPDUPData)
			{
				if (MboxMail::pCUPDUPData->ShouldTerminate())
				{
					if (ShExecInfo.hProcess)
					{
						TerminateProcess(ShExecInfo.hProcess, IDCANCEL);
						CloseHandle(ShExecInfo.hProcess);
					}
					failed = TRUE;
					break;
				}
				if (progressBar)
				{
					if (stepCnt % 10 == 0)
					{
						nSeconds++;
						if (progressText.IsEmpty())
							MboxMail::pCUPDUPData->SetProgress(step);
						else {
							CString fmt = L"%s     %d seconds";
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
		if (failed)
		{
			break;
		}
	}

	DWORD dwExitStatus = GetLastError();
	errorText = FileUtils::GetLastErrorAsString();
	return -1;
}

int NListView::VerifyPathToForwardEmlFileExecutable(CString &ForwardEmlFileExePath , CString &errorText)
{
	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	CString processPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"processPath");

	CString processDir;
	FileUtils::CPathGetPath(processPath, processDir);

	if (m_developmentMode)
	{
#ifdef _DEBUG
		processDir = L"G:\\MailKit\\ForwardEmlFile\\bin\\Debug\\netcoreapp3.1";
		processDir = L"G:\\Documents\\GIT1.0.3.21 - Copy\\mboxviewer\\ForwardEmlFile\\bin\\Debug\\netcoreapp3.1";
#else
		processDir = L"G:\\MailKit\\ForwardEmlFile\\bin\\Rekease\\netcoreapp3.1";
		processDir = L"G:\\Documents\\GIT1.0.3.21 - Copy\\mboxviewer\\ForwardEmlFile\\bin\\Rekease\\netcoreapp3.1";
#endif
	}
	CString filePath = processDir + "\\ForwardMails\\ForwardEmlFile.exe";

	ForwardEmlFileExePath = filePath;

	if (!FileUtils::PathFileExist(filePath))
	{
		CString dir = processDir + L"\\ForwardMails";
		errorText = L"ForwardEmlFile.exe not found in ";
		errorText.Append(dir);
		errorText.Append(L" directory. ");
		errorText.Append(L"MBox Viewer will not be able to forward emails until ForwardEmlFile.exe is reinstalled.\n");
		MboxMail::assert_unexpected();
		return -1;
	}
	return 1;
}

int NListView::WriteMboxListFile_v2(MailArray *mailsArray, CString &filePath, _int64 mboxFileSize, CString &errorText)
{
	SerializerHelper sz(filePath);
	if (!sz.open(TRUE))
	{
		CString lastError = FileUtils::GetLastErrorAsString();
#if 0
		CString txt = L"Could not create \"" + filePath + L"\"\n";
		txt += L"Error: " + lastError + L"\n";
#endif

		CString fmt = L"Could not create \"%s\"\nError: %s\n";
		ResHelper::TranslateString(fmt);
		CString txt;
		txt.Format(fmt, filePath, lastError);


		int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	// Create mboxlist file to allow reload of archive file list

	sz.writeInt(MAIL_LIST_VERSION);			// version
	//sz.writeString(mboxFile);  // TODO: ??
	// main mbox mail file size other mbox files are derived/contain subset
	sz.writeInt64(mboxFileSize);
	// we don't always create new mail archive corresponding to mboxListFile/mail list
	// mailFileSize is mboxMailFileSize corresponding to mboxListFile
	_int64 mailFileSize = -1;  //  FileSize(mboxFilePath);  
	sz.writeInt64(mailFileSize);	// file size
	sz.writeInt(mailsArray->GetSize());

	MboxMail *m;
	for (int i = 0; i < mailsArray->GetSize(); i++)
	{
		m = (*mailsArray)[i];
		sz.writeInt64(m->m_startOff);
		sz.writeInt(m->m_length);
		sz.writeInt(m->m_index);
		sz.writeInt((int)m->m_ContentDetailsArray.size());
	}
	sz.close();

	return 1;
}

int NListView::WriteMboxLabelListFile(MailArray* mailsArray, CStringA &gLabel, DWORD gLabelCodePage, CString& filePath, _int64 mboxFileSize, CString& errorText)
{
	SerializerHelper sz(filePath);
	if (!sz.open(TRUE))
	{
		CString lastError = FileUtils::GetLastErrorAsString();
#if 0
		CString txt = L"Could not create \"" + filePath + L"\"\n";
		txt += L"Error: " + lastError + L"\n";
#endif

		CString fmt = L"Could not create \"%s\"\nError: %s\n";
		ResHelper::TranslateString(fmt);
		CString txt;
		txt.Format(fmt, filePath, lastError);

		int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	// Create mboxlist file to allow reload of archive file list

	sz.writeInt(MAIL_LABEL_LIST_VERSION);			// version
	//sz.writeString(mboxFile);  // TODO: ??
	// main mbox mail file size other mbox files are derived/contain subset
	sz.writeInt64(mboxFileSize);
	// we don't always create new mail archive corresponding to mboxListFile/mail list
	// mailFileSize is mboxMailFileSize corresponding to mboxListFile
	_int64 mailFileSize = -1;  //  FileSize(mboxFilePath);  
	sz.writeInt64(mailFileSize);	// file size
	sz.writeInt(mailsArray->GetSize());
	sz.writeString((LPCSTR)gLabel);
	sz.writeInt(gLabelCodePage);

	MboxMail* m;
	for (int i = 0; i < mailsArray->GetSize(); i++)
	{
		m = (*mailsArray)[i];
		sz.writeInt64(m->m_startOff);
		sz.writeInt(m->m_length);
		sz.writeInt(m->m_index);
		sz.writeInt((int)m->m_ContentDetailsArray.size());
	}
	sz.close();

	return 1;
}

int NListView::SaveAsLabelFile(MailArray *marray, CString &targetDir, CStringA &labelName, DWORD labelCodePage, CStringA& mappedLabelName, CString &errorText)
{
	static const char* cCategory = "category ";
	static const int cCategoryLength = istrlen(cCategory);

	MailArray *mailsArray = marray;

	TRACE("SaveAsLabelFile: ArrayCnt=%d targetDir=%s labelName=%s labelCodePage=%d mappedLabelName=%s\n", 
		mailsArray->GetCount(), targetDir, labelName, labelCodePage, mappedLabelName);  // FIXME string types diff

#if 1
	CString mailFile = MboxMail::s_path;

	if (MboxMail::s_path.IsEmpty()) {
		CString txt = L"Please open mail file first.";
		ResHelper::TranslateString(txt);
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}


	CString path = MboxMail::GetLastPath();
	CString datapath = MboxMail::GetLastDataPath();
	if (path.IsEmpty()) {
		MboxMail::assert_unexpected();
		TRACE(L"SaveAsLabelFile: GetLastPath empty\n");
		return -1;  // Hopefully s_path wil fail first
	}
#endif

	CString labelsCachePath;
	CString rootPrintSubFolder = L"LabelCache";
	//CString errorText;
	CString targetPrintSubFolder;
	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, labelsCachePath, errorText);
	if (retval == FALSE) {
		MboxMail::assert_unexpected();
		TRACE(L"SaveAsLabelFile: CreateCachePath  LabelCache failed\n");
		return -1;
	}

	CString rootLabelsCachePath = labelsCachePath;

	CStringA delim = "/";

	CStringArray gLabelArray;
	TextUtilsEx::SplitStringA(labelName, delim, gLabelArray);
	int glarCnt = (int)gLabelArray.GetSize();
	CStringA gLabel;
	if (glarCnt > 0)
		gLabel = gLabelArray.ElementAt(glarCnt-1);

	CStringArray l;
	TextUtilsEx::SplitStringA(mappedLabelName, delim, l);

	int lcnt = (int)l.GetSize();
	if (glarCnt != lcnt)
		int deb = 1;

	CStringA label;
	if (lcnt > 0)
		label = l.GetAt(lcnt - 1);
	if (lcnt > 1)
	{
		for (int i = 0; i < lcnt-1; i++)
		{
			CString ll = l.ElementAt(i);
			labelsCachePath = labelsCachePath + L"\\" + ll;
			//TRACE(L"    |%s|\n", ll);
		}
	}

	BOOL cret = FileUtils::CreateDir(labelsCachePath);

	CFile fp;
	CString labelW = label;
	CString labelListFilePath = labelsCachePath + L"\\" + labelW + L".mboxlist";

	_int64 mailFileSize = MboxMail::s_fSize;
	int wret = NListView::WriteMboxLabelListFile(marray, gLabel, labelCodePage, labelListFilePath, mailFileSize, errorText);
	if (wret < 0)
	{
		TRACE(L"SaveAsLabelFile: WriteMboxLabelListFile  failed\n");
		return -1;
	}

	return 1;
}

int NListView::LoadLabelListFile_v2(CString &listFilePath, CString &folderName, CString &mboxFilePath)
{
	int ret = 1;  //OK

	MailArray *mailsArray = &MboxMail::s_mails;

#if 0
	if (MboxMail::IsFolderMailsSelected())
	{
		// should never be here
		return -1;
	}
#endif

#if 0
	if (MboxMail::IsFolderMailsSelected())
	{
		CString txt = L"Folder Selected Mails List not empty. Overwrite?";
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
			return -1;
	}
#endif

	CString mboxFileNamePath = MboxMail::s_path;

	if (mboxFileNamePath.IsEmpty()) {
		CString txt = L"Please open mail file first.";
		ResHelper::TranslateString(txt);

		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	CString path = MboxMail::GetLastPath();
	if (path.IsEmpty())
	{
		MboxMail::assert_unexpected();
		return -1;  // Hopefully s_path will fail first
	}


	if (!FileUtils::PathFileExist(listFilePath))
	{
#if 0
		CString txt = L"Mail Label List File \"" + listFilePath;
		txt += L"\" doesn't exist.\nCan't reload.";
#endif

		CString fmt = L"Mail Label List File \"%s\" doesn't exist.\nCan't reload.";
		ResHelper::TranslateString(fmt);
		CString txt;
		txt.Format(fmt, listFilePath);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	SerializerHelper sz(listFilePath);
	if (!sz.open(FALSE)) {

		CString lastError = FileUtils::GetLastErrorAsString();
#if 0
		CString txt = L"Could not open \"" + listFilePath + L"\"\n";
		txt += L"Error: " + lastError + L"\n";
#endif

		CString fmt = L"Could not open \"%s\"\nError: %s\n";
		ResHelper::TranslateString(fmt);
		CString txt;
		txt.Format(fmt, listFilePath, lastError);

		int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	int version;
	_int64 mailFileSize;
	_int64 mboxFileSize;
	int mailListCnt;
#if 0
	CString txt = L"Mail label list file\n\"" + listFilePath;
	txt += L"\"\nhas incompatible version or file is corrupted.\nCan't reload.\nPlease select Refresh Gmail Labels.";
#endif
#if 0
	CString fmt = L"Mail label list file\n\"%s\"\nhas incompatible version or file is corrupted.\nCan't reload.\nPlease select Refresh Gmail Labels.";
	ResHelper::TranslateString(fmt);
	CString txt;
	txt.Format(fmt, listFilePath);
#endif

	if (!sz.readInt(&version)) {
		sz.close();
		MboxMail::assert_unexpected();
		return -1;
	}

	int mailLabelListVersion = MAIL_LABEL_LIST_VERSION;
	int versionBase = MAIL_LIST_VERSION_BASE;
	int expectedVersion = mailLabelListVersion - versionBase;
	int listFileVersion = version - versionBase;
	if (version != mailLabelListVersion)
	{
		sz.close();
#if 0
		CString text = L"Mail label list file\n\n\"" + listFilePath;
		CString strVersion;

		strVersion.Format(L"%d", listFileVersion);
		text += L"\".\n\nhas incompatible version\"" + strVersion + L"\". Expected version \"";
		strVersion.Format(L"%d", expectedVersion);
		text += strVersion + L"\".\nCan't reload.\nPlease select Refresh Gmail Labels";
		text += L" for mbox mail file\n\n" + mboxFilePath;
#endif
		CString text;
		CString strVersion;
		strVersion.Format(L"%d", listFileVersion);
		CString strExpectedVersion;
		strExpectedVersion.Format(L"%d", expectedVersion);

		CString fmt = L"Mail label list file\n\n\"%s\".\n\nhas incompatible version\"%s\". Expected version \"%s\".\nCan't reload.\nPlease select Refresh Gmail Labels"
			L" for mbox mail file\n\n%s"
			;

		ResHelper::TranslateString(fmt);
		text.Format(fmt, listFilePath, strVersion, strExpectedVersion, mboxFilePath);

		int answer = MessageBox(text, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	CString fmt = L"Mail label list file\n\"%s\"\nhas incompatible version or file is corrupted.\nCan't reload.\nPlease select Refresh Gmail Labels.";
	ResHelper::TranslateString(fmt);
	CString txt;
	txt.Format(fmt, listFilePath);

	if (!sz.readInt64(&mailFileSize)) {
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	if (!sz.readInt64(&mboxFileSize)) {
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	if (!sz.readInt(&mailListCnt)) {
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	_int64 nMailFileSize = FileUtils::FileSize(mboxFileNamePath);


	if ((mailListCnt < 0) || (mailListCnt > MboxMail::s_mails_ref.GetCount()) ||
		(mailFileSize != nMailFileSize))
	{
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	CStringA gLabel;
	if (!sz.readString(gLabel)) {
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	DWORD gLabelCodePage = 0;
	int codePage = 0;
	if (!sz.readInt(&codePage)) {
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	gLabelCodePage = codePage;

	BOOL errorDoNothing = FALSE;
	BOOL verifyOnly = TRUE;
	__int64 pos = 0;
	BOOL retval = sz.GetReadPointer(pos);
	if (retval == FALSE)
	{
		int deb = 1;
	}
	ret = PopulateMailArray(sz, MboxMail::s_mails_label, mailListCnt, verifyOnly);
	if (ret > 0)
	{
		verifyOnly = FALSE;
		BOOL retval = sz.SetReadPointer(pos);
		if (retval == FALSE)
		{
			int deb = 1;
		}
		ret = PopulateMailArray(sz, MboxMail::s_mails_label, mailListCnt, verifyOnly);
	}
	else
	{
		sz.close();
#if 0
		CString txt = L"Mail list file\n\"" + listFilePath;
		txt += L"\"\n is invalid or corrupted. Can't reload.\nPlease refresh Gmail Labels.";
#endif
		CString fmt = L"Mail list file\n\"%s\"\n is invalid or corrupted. Can't reload.\nPlease refresh Gmail Labels.";
		ResHelper::TranslateString(fmt);
		CString txt;
		txt.Format(fmt, listFilePath);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1; // do nothing
	}

	sz.close();

	if (ret < 0)
	{
		MboxMail::s_mails_folder.SetSizeKeepData(0);
#if 0
		CString txt = L"Mail list file\n\"" + listFilePath;
		txt += L"\"\n is invalid or corrupted. Can't reload.\nPlease refresh Gmail Labels.";
#endif
		CString fmt = L"Mail list file\n\"%s\"\n is invalid or corrupted. Can't reload.\nPlease refresh Gmail Labels.";
		ResHelper::TranslateString(fmt);
		CString txt;
		txt.Format(fmt, listFilePath);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
	}
	else
	{
		MboxMail::SortByDate(&MboxMail::s_mails_label);
		if (MboxMail::s_mails_label.GetCount() > 0)
			MboxMail::m_labelMails.m_lastSel = 0;
		else
			MboxMail::m_labelMails.m_lastSel = -1;
		MboxMail::m_labelMails.b_mails_which_sorted = 1;

#if 1
		SwitchToMailList(IDC_LABEL_LIST, TRUE);
#else

		if (MboxMail::IsLabelMailsSelected())
		{
			SwitchToMailList(IDC_LABEL_LIST, TRUE);
		}
		else if (MboxMail::IsAllMailsSelected())
		{
			SwitchToMailList(IDC_LABEL_LIST, FALSE);
		}
#endif
	}

	int ni = MboxMail::s_mails.GetCount();

	m_list.EnsureVisible(ni, FALSE);
	m_list.SetRedraw(TRUE);

	return ret;
}

int NListView::GetLabelFromLabelListFile_v2(CString& listFilePath, CStringA& gLabel, DWORD& gLabelCodepage, CString& mboxFilePath)
{
	int ret = 1;  //OK

	MailArray* mailsArray = &MboxMail::s_mails;

#if 0
	if (MboxMail::IsFolderMailsSelected())
	{
		// should never be here
		return -1;
	}
#endif

#if 0
	if (MboxMail::IsFolderMailsSelected())
	{
		CString txt = L"Folder Selected Mails List not empty. Overwrite?";
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
			return -1;
	}
#endif

	CString mboxFileNamePath = MboxMail::s_path;

	if (mboxFileNamePath.IsEmpty())
	{
		CString txt = L"Please open mail file first.";
		ResHelper::TranslateString(txt);

		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	CString path = MboxMail::GetLastPath();
	if (path.IsEmpty())
	{
		MboxMail::assert_unexpected();
		return -1;  // Hopefully s_path will fail first
	}


	if (!FileUtils::PathFileExist(listFilePath))
	{
#if 0
		CString txt = L"Mail Label List File \"" + listFilePath;
		txt += L"\" doesn't exist.\nCan't reload.";
#endif

		CString fmt = L"Mail Label List File \"%s\" doesn't exist.\nCan't reload.";
		ResHelper::TranslateString(fmt);
		CString txt;
		txt.Format(fmt, listFilePath);

		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	SerializerHelper sz(listFilePath);
	if (!sz.open(FALSE))
	{

		CString lastError = FileUtils::GetLastErrorAsString();
#if 0
		CString txt = L"Could not open \"" + listFilePath + L"\"\n";
		txt += L"Error: " + lastError + L"\n";
#endif

		CString fmt = L"Could not open \"%s\"\nError: %s\n";
		ResHelper::TranslateString(fmt);
		CString txt;
		txt.Format(fmt, listFilePath, lastError);

		int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	int version;
	_int64 mailFileSize;
	_int64 mboxFileSize;
	int mailListCnt;
#if 0
	CString txt = L"Mail label list file\n\"" + listFilePath;
	txt += L"\"\nhas incompatible version or file is corrupted.\nCan't reload.\nPlease select Refresh Gmail Labels.";
	txt += L"\nfor mbox mail file\n\n" + mboxFilePath;
#endif
#if 0
	CString fmt = L"Mail label list file\n\"%s\"\nhas incompatible version or file is corrupted.\n"
		"Can't reload.\nPlease select Refresh Gmail Labels.\nfor mbox mail file\n\n%s";
	ResHelper::TranslateString(fmt);
	CString txt;
	txt.Format(fmt, listFilePath, mboxFilePath);
#endif

	if (!sz.readInt(&version))
	{
		sz.close();
		MboxMail::assert_unexpected();
		return -1;
	}

	int mailLabelListVersion = MAIL_LABEL_LIST_VERSION;
	int versionBase = MAIL_LIST_VERSION_BASE;
	int expectedVersion = mailLabelListVersion - versionBase;
	int listFileVersion = version - versionBase;
	if (version != mailLabelListVersion)
	{
		sz.close();
#if 0
		CString text = L"Mail label list file\n\n\"" + listFilePath;
		CString strVersion;

		strVersion.Format(L"%d", listFileVersion);
		text += L"\".\n\nhas incompatible version\"" + strVersion + L"\". Expected version \"";
		strVersion.Format(L"%d", expectedVersion);
		text += strVersion + L"\".\nCan't reload.\nPlease select Refresh Gmail Labels";
		text += L" for mbox mail file\n\n" + mboxFilePath;
#endif

		CString text;
		CString strVersion;
		strVersion.Format(L"%d", listFileVersion);
		CString strExpectedVersion;
		strExpectedVersion.Format(L"%d", expectedVersion);

		CString fmt = L"Mail label list file\n\n\"%s\".\n\nhas incompatible version\"%s\". Expected version \"%s\".\nCan't reload.\nPlease select Refresh Gmail Labels"
			L" for mbox mail file\n\n%s"
			;

		ResHelper::TranslateString(fmt);
		text.Format(fmt, listFilePath, strVersion, strExpectedVersion, mboxFilePath);

		int answer = MessageBox(text, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	CString fmt = L"Mail label list file\n\"%s\"\nhas incompatible version or file is corrupted.\n"
		"Can't reload.\nPlease select Refresh Gmail Labels.\nfor mbox mail file\n\n%s";
	ResHelper::TranslateString(fmt);
	CString txt;
	txt.Format(fmt, listFilePath, mboxFilePath);

	if (!sz.readInt64(&mailFileSize))
	{
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	if (!sz.readInt64(&mboxFileSize))
	{
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	if (!sz.readInt(&mailListCnt))
	{
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	_int64 nMailFileSize = FileUtils::FileSize(mboxFileNamePath);

#if 0
	if ((mailListCnt < 0) || (mailListCnt > MboxMail::s_mails_ref.GetCount()) ||
		(mailFileSize != nMailFileSize))
	{
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}
#endif

	//CString gLabel;
	if (!sz.readString(gLabel)) {
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	int codePage = 0;
	if (!sz.readInt(&codePage))
	{
		sz.close();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	gLabelCodepage = codePage;
	
	sz.close();

	return 1;
}

#if 0  // FIXME check if need to keep
// Get Text from Html and create file to be able later to compare with text plain block
int NListView::ExportTextToTextFile(CString &textFileName, MailIndexList *selectedMailIndexList)
{
	CString colLabels;
	CString separator;

	CFile fp;
	CString textFile;
	CString errorText;
	bool fileExists = false;
	int ret = 1;


	ret = MboxMail::MakeFileNameFromMailArchiveName(textType, textFile, targetPrintSubFolder, fileExists, errorText);


	if (ret < 0) {
		HWND h = NULL;
		int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	textFileName = textFile;
	int textFileLength = textFile.GetLength();
	int maxPath = _MAX_PATH;

	{
		CFileException ExError;
		if (!fp.Open(textFile, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &ExError))
		{
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt = L"Could not create \"" + textFile;
			txt += L"\" file.\n";
			txt += exErrorStr;

			//TRACE(L"%s\n", txt);

			CFileStatus rStatus;
			BOOL ret = fp.GetStatus(rStatus);

			//errorText = txt;

			HWND h = NULL; // we don't have any window yet
			int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}

		CFile fpm;
		CFileException ExError2;
		if (!fpm.Open(s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError2))
		{
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError2);

			CString txt = L"Could not open \"" + MboxMail::s_path;
			txt += L"\" mail file.\n";
			txt += exErrorStr;

			//TRACE(L"%s\n", txt);
			//errorText = txt;

			HWND h = NULL; // we don't have any window yet
			int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			fp.Close();
			return -1;
		}

		if (textType == 0) {
			if (textConfig.m_nCodePageId == CP_UTF8) {
				const char *BOM_UTF8 = "\xEF\xBB\xBF";
				fp.Write(BOM_UTF8, 3);
			}
		}

		CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
		BOOL printAttachments = TRUE;
		if (pFrame)
		{
			printAttachments = pFrame->m_HdrFldConfig.m_HdrFldList.IsFldSet(HdrFldList::HDR_FLD_ATTACHMENTS-1);
		}

		AttachmentMgr attachmentDB;

		{
			int i;
			firstMail = (*selectedMailIndexList)[0];
			int cnt = (int)selectedMailIndexList->GetCount();

			BOOL singleMail = (cnt == 1) ? TRUE : FALSE;
			for (int j = 0; j < cnt; j++)
			{
				i = (*selectedMailIndexList)[j];

					printSingleMailToTextFile(fp, i, fpm, textConfig, singleMail, addPageBreak);
		}

		fp.Close();
		fpm.Close();
	}
}
#endif

// Get Text from Html and create file to be able later to compare with text plain block
// Not finished and not used yet
int NListView::ExportTextTextToTextFile(/*out*/CFile &fp, int mailPosition, /*in mail body*/ CFile &fpm)
{
	UINT pageCode;

	MboxMail *m = MboxMail::s_mails[mailPosition];

	SimpleString *outbuf = MboxMail::m_outbuf;
	SimpleString *inbuf = MboxMail::m_inbuf;
	SimpleString *workbuf = MboxMail::m_workbuf;

	inbuf->ClearAndResize(10000);
	workbuf->ClearAndResize(10000);
	outbuf->ClearAndResize(10000);

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	//
	int textType = 0; // Text
	pageCode = 0;
	int plainTextMode = 1;  // insert [attachment name] image info
	int textlen = MboxMail::GetMailBody_mboxview(fpm, mailPosition, outbuf, pageCode, textType, plainTextMode);  // returns pageCode
	if (textlen != outbuf->Count())
		int deb = 1;

	if (outbuf->Count())
	{
		fp.Write(outbuf->Data(), outbuf->Count());
	}
#if 0
	else
	{
		outbuf->Clear();
		pageCode = 0;
		int textType = 1; // HTML

		int textlen = GetMailBody_mboxview(fpm, mailPosition, outbuf, pageCode, textType);  // returns pageCode
		if (textlen != outbuf->Count())
			int deb = 1;

		if (outbuf->Count())
		{
			MboxMail::m_Html2TextCount++;

			CString bdycharset;
			std::string charSet;
			if (pageCode > 0)
			{
				BOOL ret = TextUtilsEx::id2charset(pageCode, charSet);
#if 0
				// TODO: charset in the mail Content-Type an in mail body differs; current approach seems to work 
				UINT pageCodeFromBody = getCodePageFromHtmlBody(outbuf, charSet);
				if ((pageCodeFromBody) && (pageCodeFromBody != pageCode))
					int deb = 1;
#endif
			}
			else
			{
				pageCode = getCodePageFromHtmlBody(outbuf, charSet);

				if (pageCode == 0) {
					pageCode = CP_UTF8;
					BOOL ret = TextUtilsEx::id2charset(pageCode, charSet);
				}
			}

			bdycharset.Append(charSet.c_str());

			inbuf->Clear();

			CString bdy;
			bool extraHtmlHdr = false;
			if (outbuf->FindNoCase(0, "<body", 5) < 0) // didn't find if true
			{
				extraHtmlHdr = true;
				bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body>";
				inbuf->Append(bdy, bdy.GetLength());
			}

			inbuf->Append(*outbuf);
			if (extraHtmlHdr)
				inbuf->Append("</body></html>\r\n");

			outbuf->Clear();

			UINT outPageCode = CP_UTF8;
			// Below Relies on IHTMLDocument2
			HtmlUtils::GetTextFromIHTMLDocument(inbuf, outbuf, pageCode, outPageCode);
			fp.Write(outbuf->Data(), outbuf->Count());
		}
	}
#endif

	outbuf->Clear();
	outbuf->Append("\r\n");
	fp.Write(outbuf->Data(), outbuf->Count());

	return 1;
}

/// ToolTips
// Doesn't work. Likely du to Custom Draw. Neeed to investigate when timepermits
// move below to CWheelListCtrl
//
void NListView::CellHitTest(const CPoint& pt, int& nRow, int& nCol) const
{
	nRow = -1;
	nCol = -1;

	LVHITTESTINFO lvhti = { 0 };
	lvhti.pt = pt;
	nRow = ListView_SubItemHitTest(m_hWnd, &lvhti);	// SubItemHitTest is non-const
	nCol = lvhti.iSubItem;
	if (!(lvhti.flags & LVHT_ONITEMLABEL))
		nRow = -1;
}

BOOL NListView::OnToolNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	CPoint pt(GetMessagePos());
	ScreenToClient(&pt);

	int nRow, nCol;
	CellHitTest(pt, nRow, nCol);

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

bool NListView::ShowToolTip(const CPoint& pt) const
{
	// Lookup up the cell
	int nRow, nCol;
	CellHitTest(pt, nRow, nCol);

	if (nRow != -1 && nCol != -1)
		return true;
	else
		return false;
}

CString NListView::GetToolTipText(int nRow, int nCol)
{
	if (nRow != -1 && nCol != -1)
		return m_list.GetItemText(nRow, nCol);	// Cell-ToolTip
	else
		return CString("");
}

namespace {
	LRESULT EnableWindowTheme(HWND hwnd, LPCWSTR classList, LPCWSTR subApp, LPCWSTR idlist)
	{
		LRESULT lResult = S_FALSE;
		HMODULE hinstDll;
		BOOL(WINAPI * pIsThemeActive)();
		HRESULT(WINAPI * pSetWindowTheme)(HWND hwnd, LPCWSTR pszSubAppName, LPCWSTR pszSubIdList);
		HANDLE(WINAPI * pOpenThemeData)(HWND hwnd, LPCWSTR pszClassList);
		HRESULT(WINAPI * pCloseThemeData)(HANDLE hTheme);

		// Check if running on Windows XP or newer
		hinstDll = ::LoadLibrary(L"UxTheme.dll");
		if (hinstDll)
		{
			// Check if theme service is running
			(FARPROC&)pIsThemeActive = ::GetProcAddress(hinstDll, "IsThemeActive");
			if (pIsThemeActive && pIsThemeActive())
			{
				(FARPROC&)pOpenThemeData = ::GetProcAddress(hinstDll, "OpenThemeData");
				(FARPROC&)pCloseThemeData = ::GetProcAddress(hinstDll, "CloseThemeData");
				(FARPROC&)pSetWindowTheme = ::GetProcAddress(hinstDll, "SetWindowTheme");
				if (pSetWindowTheme && pOpenThemeData && pCloseThemeData)
				{
					// Check is themes is available for the application
					HANDLE hTheme = pOpenThemeData(hwnd, classList);
					if (hTheme != NULL)
					{
						VERIFY(pCloseThemeData(hTheme) == S_OK);
						// Enable Windows Theme Style
						lResult = pSetWindowTheme(hwnd, subApp, idlist);
					}
				}
			}
			::FreeLibrary(hinstDll);
		}
		return lResult;
	}
}

void NListView::PreSubclassWindow()
{
	CWnd::PreSubclassWindow();

#if 0
	// Focus retangle is not painted properly without double-buffering
#if (_WIN32_WINNT >= _WIN32_WINNT_WINXP)
	m_list.SetExtendedStyle(LVS_EX_DOUBLEBUFFER | m_list.GetExtendedStyle());
#endif
	m_list.SetExtendedStyle(m_list.GetExtendedStyle() | LVS_EX_FULLROWSELECT);
	m_list.SetExtendedStyle(m_list.GetExtendedStyle() | LVS_EX_HEADERDRAGDROP);
	m_list.SetExtendedStyle(m_list.GetExtendedStyle() | LVS_EX_GRIDLINES);

	// Enable Vista-look if possible
	EnableWindowTheme(m_hWnd, L"ListView", L"Explorer", NULL);
#endif
}

///

void NListView::OnTvnGetInfoTip(NMHDR* pNMHDR, LRESULT* pResult)
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

void NListView::AppendPictureAttachments(MboxMail *m, AttachmentMgr& attachmentDB, CString* absoluteFolderPath, CString* relativeFolderPath, CFile* fpm)
{
	//attachmentDB.Sort();

	_ASSERTE(absoluteFolderPath || relativeFolderPath);

	CStringA srcPrefix;
	CString folderPath;
	if (absoluteFolderPath)
	{
		srcPrefix = "file:\\\\\\";
		folderPath = *absoluteFolderPath;
	}
	else if (relativeFolderPath)
	{
		srcPrefix = "";
		folderPath = *relativeFolderPath;
		folderPath.TrimRight(L"\\");
		folderPath.Append(L"\\");
	}

	int attachmentCnt = (int)attachmentDB.m_attachmentArray.GetCount();

	int pictureCnt = 0;
	CStringArray arr;
	int i;
	for (i = 0; i < attachmentDB.m_attachmentArray.GetCount(); i++)
	{
		AttachmentData* item = &attachmentDB.m_attachmentArray[i];
		if (!item->m_isEmbedded)
		{
			// More reliable approach would be to check metadata
			CString fileNameExtention;
			FileUtils::GetFileExtension(item->m_nameW, fileNameExtention);
			if (GdiUtils::IsSupportedPictureFileExtension(fileNameExtention))
			{
				arr.Add(item->m_nameW);
			}
		}
	}

	if (arr.GetCount())
	{
		CString attachmentFileNamePrefix;

		if (m)
		{
			NListView::GetExtendedMailId(m, attachmentFileNamePrefix);
		}

		SimpleString pics;
		CStringA bdycharset = "UTF-8";
		CStringA hdr;
		CStringA hdrExtra;
		CStringA base;
		if (NListView::m_exportMailsMode)
		{
			base = "<base href=\"\"/>";  // TODOE define empty base ??
		}
		//hdr = "\r\n<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\">" + base + "</head><body>\r\n";
		hdr = "\r\n<div>";

		pics.Append((LPCSTR)hdr, hdr.GetLength());

		//pics.Append("<span><br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br><br></span>");
		pics.Append("<div></div><span><br>--------&nbsp;Picture Attachments&nbsp;--------</span><div></div><br>");

		for (i = 0; i < arr.GetCount(); i++)
		{
			const CString& attachmenName = arr.GetAt(i);
			CString attachementPath = folderPath + attachmentFileNamePrefix + attachmenName;
			CStringA attachementPathA;
			CStringA attachmenNameA;

			AttachmentData *attachmentinfo = attachmentDB.GetAttachmentData(attachmenName);

			DWORD error;

			UINT CP_US_ASCII = 20127;
			UINT outCodePage = CP_UTF8;
			BOOL retW2A = TextUtilsEx::WStr2CodePage((LPCWSTR)attachementPath, attachementPath.GetLength(), outCodePage, &attachementPathA, error);

			SimpleString attachmentFilePathA_SS;
			attachmentFilePathA_SS.Append((LPCSTR)attachementPathA, attachementPathA.GetLength());

			retW2A = TextUtilsEx::WStr2CodePage((LPCWSTR)attachmenName, attachmenName.GetLength(), outCodePage, &attachmenNameA, error);

			SimpleString validNameUTF8;
			validNameUTF8.Append((LPCSTR)attachmenNameA, attachmenNameA.GetLength());

			attachmenNameA.Empty();
			attachmenNameA.Append(validNameUTF8.Data(), validNameUTF8.Count());

			TRACE(L"AppendPictureAttachments:All: %d value \"%s\"\n", i, attachmenName);

#if 0
			pics.Append("<br><br><span>---&nbsp;");
			pics.Append((LPCSTR)attachmenNameA, attachmenNameA.GetLength());
			pics.Append("&nbsp;---------------------</span><br><br>");
#endif

			int imageWidth = 0;
			int imageHeight = 0;
			int imageW = 90;

			if (attachmentinfo)
			{
				imageWidth = attachmentinfo->m_width;
				imageHeight = attachmentinfo->m_height;
			}

			if ((imageHeight > imageWidth) && (imageWidth > 0))
			{
				float h2wPage;
				h2wPage = (float)(297 - 24) / (float)(210 - 10);  // A4
				h2wPage = (float)(279 - 24) / (float)(216 - 10);  // Letter
				float h2wImage = (float)imageHeight / (float)imageWidth;
				float w2hImage = (float)imageWidth / (float)imageHeight;

				imageW = (int)(imageW * w2hImage * h2wPage);
			}

			CStringA widthStr;
			widthStr.Format("\r\n<img style=\"max-width:%d%%;\" src=\"", imageW);
			pics.Append((LPCSTR)widthStr, widthStr.GetLength());

			//pics.Append("\r\n<img style=\"max-width:60%;\" src=\"");
			//pics.Append("\r\n<img style=\"max-height:80%;\" src=\"");
			pics.Append((LPCSTR)srcPrefix, srcPrefix.GetLength());
			pics.Append((LPCSTR)attachementPathA, attachementPathA.GetLength());
			pics.Append("\">\r\n");

#if 1
			pics.Append("<br><br><span>^^^---------------------&nbsp;");
			pics.Append((LPCSTR)attachmenNameA, attachmenNameA.GetLength());
			pics.Append("&nbsp;</span><br><br>");
			//pics.Append("&nbsp;---------------------</span><br><br>");
#endif

		}
		//pics.Append("\r\n</body></html>\r\n");
		pics.Append("\r\n</div>\r\n");

		fpm->Write(pics.Data(), pics.Count());

		const int deb = 1;
	}
}

//#################################################################################################
//# inData  - points to encoded HTML block retrived frim raw mail
//# inDataLen - length of inData
//# output - (out) buffer with updated HTML referncing created inline/embedded images
//# relativeSrcImgFilePath - relative path to folder containing created image files
//# absoluteSrcImgFilePath - absolute folder path to created image files.
//#

//#################################################################################################

int NListView::UpdateInlineSrcImgPathEx(CFile *fpm, char* inData, int indDataLen, SimpleString* outbuf,
	BOOL makeFileNameUnique,  // add unique prefix to image file name
	BOOL makeAbsoluteImageFilePath,
	CString& relativeSrcImgFilePath, CString& absoluteSrcImgFilePath, AttachmentMgr& attachmentDB,
	EmbededImagesStats &embededImgStats, int mailPosition, BOOL createEmbeddedImageFiles,
	BOOL verifyAttachmentDataAsImageType, BOOL insertMaxWidthForImgTag,
	CStringA &maxWidth, CStringA &maxHeight)
{
	static char* img_pattern = "<img";
	static int img_patternLen = istrlen(img_pattern);

	static char* src_pattern = "src=";
	static int src_patternLen = istrlen(src_pattern);

	static char* cid_pattern = "cid:";
	static int cid_patternLen = istrlen(cid_pattern);

	static char* http_pattern = "http:";
	static int http_patternLen = istrlen(http_pattern);

	static char* https_pattern = "https:";
	static int https_patternLen = istrlen(https_pattern);

	static char* mhtml_pattern = "mhtml:";
	static int mhtml_patternLen = istrlen(mhtml_pattern);

	static char* data_pattern = "data:";
	static int data_patternLen = istrlen(data_pattern);

	char* input = inData;
	int inputLength = indDataLen;
	char* inputEnd = input + inputLength;
	char* srcImgEnd = inputEnd;
	char* cidBegin;
	char* cidEnd;
	SimpleString cid;
	CString imageFileName;
	CString imageFilePath;

	char* fromBegin = input;
	char* alt_pos;
	int patternLen = cid_patternLen;
	int mhtmlIntroLen = 0;
	char* srcBegin;
	char* srcEnd;

	MboxMail* m = 0;
	if ((mailPosition < MboxMail::s_mails.GetCount()) && (mailPosition >= 0))
		m = MboxMail::s_mails[mailPosition];
	else
		return -1;

	MailBodyContent* body;
	bool foundImage = false;
	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];

		CStringA contentTypeExtension;
		CStringA contentTypeMain;
		int pos = body->m_contentType.ReverseFind('/');
		if (pos > 0)
		{
			contentTypeExtension = body->m_contentType.Mid(pos + 1);
			contentTypeMain = body->m_contentType.Left(pos);
		}
		if (contentTypeMain.CompareNoCase("text") != 0) {
			foundImage = true;   // Maybe !!! Go ahead and try 
			break;
		}
	}

	if (createEmbeddedImageFiles)
	{
		if (!foundImage)
		{
			if (insertMaxWidthForImgTag)
			{
				;
				// We can't call return here because we need to insert 
				// style="max-width:100%;" if we find <img tag (no attachment image found)
			}
			else
			{
				m->m_DetermineEmbeddedImagesDone = 1;
				return -1;
			}
		}
	}

	if (!foundImage)
	{
		const int deb = 1;
	}

	CString imgFile;
	CStringA cidName;

	CFile fpmbox;
	CFile *fpm_save = fpm;
	if (fpm == 0)
	{
		CFileException ExError;
		if (!fpmbox.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
		{
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt = L"Could not open \"" + MboxMail::s_path;
			txt += L"\" mail file.\n";
			txt += exErrorStr;

			TRACE(L"%s\n", txt);

			//errorText = txt;

			return -1;
		}
		fpm = &fpmbox;
	}

	bool srcCid = false;
	bool srcHttp = false;
	bool srcHttps = false;
	bool srcMHtml = false;  // found mhtml
	bool srcMHtmlHtml = false;  // found mhtml and http or https
	bool srcMHtmlNoHtml = false; // did not found mhtml and http or https
	bool srcMHtmlHttp = false;  // found mhtml and http 
	bool srcMHtmlHttps = false;  // found mhtml and https
	bool srcData = false;
	bool srcLocalFile = false;

	char* pos = input;
	while ((pos != 0) && (pos < inputEnd))
	{
		pos = TextUtilsEx::strnstrUpper2Lower(pos, inputEnd, img_pattern, img_patternLen);
		if (pos == 0) {
			if (outbuf->Count() > 0)
				outbuf->Append(fromBegin, IntPtr2Int(inputEnd - fromBegin));
			break;
		}
		pos += img_patternLen;


		if ((*pos == ' ') || (*pos == '\n') || (*pos == '\r'))
		{
			outbuf->Append(fromBegin, IntPtr2Int(pos - fromBegin));
			//outbuf->Append(" style=\"max-width:100%;\" ");   // FIXME
			// FIXME best to determine max_width and max-heigth vased on photo metadata
			if (!maxWidth.IsEmpty() || !maxHeight.IsEmpty())
			{
				outbuf->Append(" style=");
			}
			if (!maxWidth.IsEmpty())
			{
				CStringA width;
				width.Format("\"max-width:%s;", maxWidth);
				outbuf->Append((LPCSTR)width, width.GetLength());
			}
			if (!maxHeight.IsEmpty())
			{
				CStringA height;
				height.Format("\"max-height:%s;", maxHeight);
				outbuf->Append((LPCSTR)height, height.GetLength());
			}
			if (!maxWidth.IsEmpty() || !maxHeight.IsEmpty())
			{
				outbuf->Append("\" ");
			}
			fromBegin = pos;
			pos += 1;
		}
		else
			continue;

		// find end of <img src ...>
		srcImgEnd = TextUtilsEx::strnstrUpper2Lower(pos, inputEnd, ">", 1);
		if (srcImgEnd == 0) { // TODO: corrupted file ?

			outbuf->Append(fromBegin, IntPtr2Int(inputEnd - fromBegin));
			break;
		}
		srcImgEnd++; // jump over

		alt_pos = TextUtilsEx::strnstrUpper2Lower(pos, srcImgEnd, "alt=\"", 5);

		pos = TextUtilsEx::strnstrUpper2Lower(pos, srcImgEnd, src_pattern, src_patternLen);
		if (pos == 0) {
			pos = srcImgEnd;  // jump over this <img .. > since no cid:
			continue;
		}

		srcCid = false;
		srcHttp = false;
		srcHttps = false;
		srcMHtml = false;  // foung mhtml
		srcMHtmlHtml = false;  // found mhtml and http or https
		srcMHtmlNoHtml = false; // did not found mhtml and http or https
		srcMHtmlHttp = false;  // found mhtml and http 
		srcMHtmlHttps = false;  // found mhtml and https
		srcData = false;
		srcLocalFile = false;

		patternLen = cid_patternLen;
		srcBegin = pos;
		srcEnd = srcBegin + src_patternLen;
		pos += src_patternLen;

		BOOL inQuotes = FALSE;
		if (*pos == '"')
		{
			inQuotes = TRUE;
			pos++;
		}

		BOOL foundHTTP = FALSE;
		cidBegin = pos;
		pos = TextUtilsEx::strnstrUpper2Lower(pos, pos + cid_patternLen, cid_pattern, cid_patternLen);
		if (pos != 0)
		{
			srcCid = true;
			patternLen = cid_patternLen;
			pos += cid_patternLen;
			cidBegin = pos;
		}

		mhtmlIntroLen = 0;
		if (pos == 0)
		{
			pos = cidBegin;
			pos = TextUtilsEx::strnstrUpper2Lower(pos, pos + mhtml_patternLen, mhtml_pattern, mhtml_patternLen);
			if (pos) {
				srcMHtml = true;
				MboxMail::m_EmbededImagesFoundMHtml++;
				pos = TextUtilsEx::strchar(pos, srcImgEnd, '!');
				if (pos) {
					mhtmlIntroLen = IntPtr2Int(pos - cidBegin + 1);
					cidBegin = pos + 1;
					pos = 0;
					foundHTTP = TRUE;
					srcMHtmlHtml = true;
					embededImgStats.m_FoundMHtmlHtml++;
					// TODOEX: why I didn't  jump over ??
				}
				else
				{
					embededImgStats.m_FoundUnexpectedMHtml++;
					pos = srcImgEnd;  // jump over this <img .. > sonce no cid:
					continue;
				}
				int deb = 1;
			}
		}

		if (pos == 0)
		{
			pos = cidBegin;
			pos = TextUtilsEx::strnstrUpper2Lower(pos, pos + http_patternLen, http_pattern, http_patternLen);
			if (pos) {
				patternLen = 0;
				pos += patternLen;  // include http:
				cidBegin = pos;
				foundHTTP = true;
				srcHttp = true;
				// TODOEX: why I didn't  jump over ??
			}
		}
		if (pos == 0)
		{
			pos = cidBegin;
			pos = TextUtilsEx::strnstrUpper2Lower(pos, pos + https_patternLen, https_pattern, https_patternLen);
			if (pos) {
				patternLen = 0;
				pos += patternLen;  // include https:
				cidBegin = pos;
				foundHTTP = true;
				srcHttps = true;
				// TODOEX: why I didn't  jump over ??
			}
		}

		if (pos == 0)
		{
			pos = cidBegin;
			pos = TextUtilsEx::strnstrUpper2Lower(pos, pos + data_patternLen, data_pattern, data_patternLen);
			if (pos) { // not implemented ; implemented by browser, etrc
				// set pos = 0 since src=data: case is not handled
				srcData = true;
				//pos = 0;
				pos = srcImgEnd;  // jump over this <img .. > sonce no cid:
				continue;
			}
		}

		if (pos == 0)
		{
			// find end of cid token or local file path
			//
			if (inQuotes)
			{
				pos = TextUtilsEx::findOneOf(cidBegin, srcImgEnd, "\"");
			}
			else
			{
				pos = TextUtilsEx::findOneOf(cidBegin, srcImgEnd, " >");
			}
			if (pos == 0)
			{
				outbuf->Append(fromBegin, IntPtr2Int(srcImgEnd - fromBegin));
				pos = srcImgEnd;  // jump over this <img .. > since no cid:
				fromBegin = pos;
				continue;
			}
			cidEnd = pos;
			srcLocalFile = true;
			if ((srcImgEnd - cidEnd) > 10)
				int deb = 1;
		}
		else
		{
			// find end of cid token
			//
			if (inQuotes)
			{
				pos = TextUtilsEx::findOneOf(pos, srcImgEnd, "\"");
			}
			else
			{
				pos = TextUtilsEx::findOneOf(pos, srcImgEnd, " >");
			}
			if (pos == 0)
			{
				pos = srcImgEnd;  // jump over this <img .. > since no cid:
				continue;
			}
			cidEnd = pos;
			if ((srcImgEnd - cidEnd) > 10)
				int deb = 1;
		}

		cid.Copy(cidBegin, IntPtr2Int(cidEnd - cidBegin));
		if (foundHTTP) {
			int retLen = TextUtilsEx::DecodeURL(cid.Data(), cid.Count());
			cid.SetCount(retLen);
		}

		cidName.Empty();
		cidName.Append(cid.Data(), cid.Count());

		MailBodyContent* pBodyFound;
		//BOOL verifyAttachmentDataAsImage = TRUE;  // FIXME
		int retval = NListView::DetermineImageFileName(fpm, verifyAttachmentDataAsImageType, m, cidName, imageFileName, &pBodyFound, mailPosition);

		if (makeFileNameUnique)
		{
			CString filePrefix;
			GetExtendedMailId(m, filePrefix);
			imageFileName = filePrefix + imageFileName;
		}

		CString relativeImgPath;
		if (makeAbsoluteImageFilePath)
		{
			relativeImgPath = absoluteSrcImgFilePath;
		}
		else
		{
			relativeImgPath = relativeSrcImgFilePath;
		}

		imgFile = relativeImgPath + imageFileName;
		CString absoluteImgFilePath = absoluteSrcImgFilePath + imageFileName;

		relativeImgPath.Append(imageFileName);   // relative or full path to image file

		CStringA imgFileA;
		DWORD error;
		TextUtilsEx::WStr2CodePage((LPCWSTR)relativeImgPath, relativeImgPath.GetLength(), (UINT)CP_UTF8, &imgFileA, error);

		SimpleString tmpstr(imgFileA.GetLength() * 3);
		tmpstr.Append((LPCSTR)imgFileA, imgFileA.GetLength());

		MboxMail::encodeTextAsHtmlLink(tmpstr);

		imgFileA.Empty();
		imgFileA.Append(tmpstr.Data(), tmpstr.Count());

		if (pBodyFound && (retval > 0))
		{
			pBodyFound->m_isEmbeddedImage = true;

#ifdef _DEBUG
			embededImgStats.m_Found++;

			if (srcCid)
			{
				embededImgStats.m_FoundCid++;
				TRACE("CidFound indx=%d cid=%s \n", mailPosition, cidName);
			}
			else if (srcHttp)
			{
				if (srcMHtmlHtml)
					embededImgStats.m_FoundMHtmlHttp++;
				else
					embededImgStats.m_FoundHttp++;
			}
			else if (srcHttps)
			{
				if (srcMHtmlHtml)
					embededImgStats.m_FoundMHtmlHttps++;
				else
					embededImgStats.m_FoundHttps++;
			}
			else if (srcData)
			{
				embededImgStats.m_FoundData++;
			}
			else if (srcLocalFile)
			{
				embededImgStats.m_FoundLocalFile++;
			}
			else
				int deb = 1;
#endif

			int length = IntPtr2Int(srcEnd - fromBegin);
			outbuf->Append(fromBegin, length);

			if (makeAbsoluteImageFilePath)
				outbuf->Append("\"file:\\\\\\");
			else
				outbuf->Append("\"");

			outbuf->Append(imgFileA, imgFileA.GetLength());
			outbuf->Append('\"');
			//outbuf->Append(" style=\"max-width:100%;\" ");

			fromBegin = cidEnd;
			if (*cidEnd == '\"')
				fromBegin++;

			// Mime message can have multiple attachments with the same name
			// However, only the first instance will be embedded into html text
			// and the remaining instances will be handled as non embedded attachments
			// Attachments with the same name may have different contetnt and thus data length
			if (attachmentDB.HasName(imageFileName))  // name
			{ 
				;// _ASSERTE(FileUtils::PathFileExist(imgFile));   // path  ZMFIXME
			}
			else if (createEmbeddedImageFiles)
			{
				BOOL ret = FileUtils::PathFileExist(absoluteImgFilePath);
				//BOOL ret = TRUE;
				if (ret)
				{
					int deb = 1; // TRACE(L"UpdateInlineSrcImgPath: FOUND image %s\n", imgFile);
				}
				else
				{
					;// TRACE(L"UpdateInlineSrcImgPath: NOT FOUND image %s\n", imgFile);
					imageFilePath = absoluteImgFilePath;

					SimpleString* outb = MboxMail::get_tmpbuf();
					outb->ClearAndResize(pBodyFound->m_contentLength * 2);
					int retLen = MboxMail::DecodeBody(*fpm, pBodyFound, mailPosition, outb);
					if (outb->Count() > 1500000)
						int deb = 1;

					int imageWidth = 0;
					int imageHeight = 0;

					int retRotate = GdiUtils::AutoRotateImageFromMemory(outb->Data(), outb->Count(), imageFilePath, imageWidth, imageHeight);

					_ASSERTE((imageWidth > 0) && (imageHeight > 0));

					MboxMail::rel_tmpbuf();

					if (retRotate == 0)
					{
						BOOL isEmbedded = TRUE;
						int index = attachmentDB.InsertName(imageFileName, isEmbedded);
						_ASSERTE(index < 0);
					}

					else
					{
						CFileException ExError;
						CFile fp;

						BOOL retryOpen = FALSE;
						if (!fp.Open(imageFilePath, CFile::modeWrite | CFile::modeCreate, &ExError))
						{
							CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);  // TODO

							CString errorText;
							CString imageCachePath;
							CString rootPrintSubFolder = "ImageCache";
							//CString targetPrintSubFolder = baseFileArchiveName;
							CString targetPrintSubFolder;
							if (NListView::m_exportMailsMode)
							{
								rootPrintSubFolder = L"ExportCache";
								targetPrintSubFolder = L"Attachments";
							}

							BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, imageCachePath, errorText);
							if (retval == FALSE)
							{
								HWND h = NULL; // we don't have any window yet  
								// TODO: what to do ?
								//int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
								//return -1;
							}
							retryOpen = TRUE;
						}

						if (retryOpen && !fp.Open(imageFilePath, CFile::modeWrite | CFile::modeCreate, &ExError))
						{
							CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);  // TODO

							CFileStatus rStatus;
							BOOL ret = fp.GetStatus(rStatus);

							CString errorText(exErrorStr);

							HWND h = NULL;
							// Ignore for now
							//int answer = ::MessageBox(h, errorText, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
						}
						else
						{
							SimpleString* outb = MboxMail::get_tmpbuf();
							outb->ClearAndResize(pBodyFound->m_contentLength * 2);
							int retLen = MboxMail::DecodeBody(*fpm, pBodyFound, mailPosition, outb);
							if (outb->Count() > 1500000)
								int deb = 1;

							fp.Write(outb->Data(), outb->Count());
							fp.Close();

							BOOL isEmbedded = TRUE;
							int index = attachmentDB.InsertName(imageFileName, isEmbedded);
							_ASSERTE(index < 0);

							MboxMail::rel_tmpbuf();
						}
					}

					AttachmentData* attachmentInfo = attachmentDB.GetAttachmentData(imageFileName);
					if (attachmentInfo)
					{
						attachmentInfo->m_width = imageWidth;
						attachmentInfo->m_height = imageHeight;
					}
				}
			}
		}
		else
		{
#ifdef _DEBUG
			embededImgStats.m_NotFound++;

			if (srcCid)
			{
				embededImgStats.m_NotFoundCid++;
				TRACE("CidNotFound indx=%d cid=%s \n", mailPosition, cidName);
			}
			else if (srcHttp)
			{
				if (srcMHtmlHtml)
					embededImgStats.m_NotFoundMHtmlHttp++;
				else
					embededImgStats.m_NotFoundHttp++;
			}
			else if (srcHttps)
			{
				if (srcMHtmlHtml)
					embededImgStats.m_NotFoundMHtmlHttps++;
				else
					embededImgStats.m_NotFoundHttps++;
			}
			else if (srcData)
			{
				embededImgStats.m_NotFoundData++;
			}
			else if (srcLocalFile)
			{
				embededImgStats.m_NotFoundLocalFile++;
				TRACE("LocalFileNotFound indx=%d cid=%s \n", mailPosition, cidName);
			}
			else
				int deb = 1;
#endif

			int deb = 1;
		}
#if 0
		SimpleString alt;
		if (alt_pos == 0) {
			alt.Append("alt=\"image");
			alt.Append("\" ");
		}

		if (alt.Count()) {
			outbuf->Append(" ");
			outbuf->Append(alt);
		}
#endif
		pos = srcImgEnd;
		int deb = 1;
	}

	if (fpm_save == 0)
		fpm->Close();

	m->m_DetermineEmbeddedImagesDone = 1;

	return 1;
}
/*
// It creates files for non-embeded attachments. Unique id can be prepended to the file name.
// On return, attachmentDB contains the list of created file names without folder path specified by attachmentFolderPath
*/
int NListView::CreateMailAttachments(CFile* fpm, int mailPosition, CString* attachmentFolderPath, BOOL prependMailId, AttachmentMgr& attachmentDB, BOOL extraFileNameValidation)
{
	if ((mailPosition >= MboxMail::s_mails.GetCount()) || (mailPosition < 0))
		return -1;

	MboxMail* m = MboxMail::s_mails[mailPosition];
	if (m == 0)
	{
		// TODO: Assert ??
		return -1;
	}

	CFile mboxFp;
	CFile* fpm_save = fpm;
	if (fpm == 0)
	{
		CFileException ExError;
		if (!mboxFp.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError))
		{
			// TODO: critical failure
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

			CString txt = L"Could not open \"" + MboxMail::s_path;
			txt += L"\" mail file.\n";
			txt += exErrorStr;

			TRACE(L"%s\n", txt);

			return FALSE;
		}
		fpm = &mboxFp;
	}

	BOOL bReplaceWhiteWithUnderscore = FALSE;
	BOOL ret;

	CStringW resultW;
	CStringW validNameW;
	CStringW fileNameW;

	MailBodyContent* body;

	CString errorText;
	CString attachmentCachePath;
	CStringW printCachePathW;
	CString rootPrintSubFolder = "AttachmentCache";
	CString targetPrintSubFolder;
	if (NListView::m_exportMailsMode)
	{
		rootPrintSubFolder = "ExportCache";
		targetPrintSubFolder = "Attachments";
	}

	if (attachmentFolderPath == 0)
	{

		BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, attachmentCachePath, errorText);
		if (retval == FALSE)
		{
			HWND h = NULL; // we don't have any window yet  
			// TODO: what to do ?
			//int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			_ASSERTE(FALSE);
			//return -1;  // FIXME
		}

		//ret = TextUtilsEx::Ansi2WStr(attachmentCachePath, printCachePathW, error);

		printCachePathW = attachmentCachePath;
	}
	else
	{
		printCachePathW = *attachmentFolderPath;
		printCachePathW.TrimRight(L"\\");
	}

	if (!m->m_DetermineEmbeddedImagesDone)
	{
		// Sets body->m_isEmbeddedImage
		ret = MboxMail::DetermineEmbeddedImages(mailPosition, *fpm);
	}
	_ASSERTE(m->m_DetermineEmbeddedImagesDone == TRUE);

	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];

		if (body->IsAttachment() || body->IsInlineAttachment())
		{
			if (body->m_isEmbeddedImage)
				continue;
		}
		else
			continue;

		UINT inCodePage = body->m_attachmentNamePageCode;
		resultW.Empty();

		CString fileName;
		if (prependMailId)
		{
			GetExtendedMailId(m, fileName);
		}
		CStringW fileNameW;
		fileNameW = fileName;

		SimpleString* outbuf = MboxMail::get_tmpbuf();
		outbuf->ClearAndResize(body->m_contentLength * 2);

		int retLen = MboxMail::DecodeBody(*fpm, body, mailPosition, outbuf);
		if (outbuf->Count() > 1500000)
			int deb = 1;

		// FIXME remapDuplicateNames = TRUE implies that duplicate names are remapped and added into attachmentDB
		BOOL remapDuplicateNames = TRUE;
		DetermineAttachmentName(fpm, mailPosition, body, outbuf, resultW, attachmentDB, remapDuplicateNames, extraFileNameValidation);

		if (resultW.IsEmpty())
			_ASSERTE(FALSE);

#if 0
		// Moved to DetermineAttachmentName to address inconsistent file name mappings
		fileNameW.Append(resultW);
		validNameW.Empty();
		FileUtils::MakeValidFileName(fileNameW, validNameW, bReplaceWhiteWithUnderscore);
		if (validNameW.Compare(resultW))
			int deb = 1;
#endif

		validNameW = fileNameW;
		validNameW.Append(resultW);  // FIXME

		CStringW filePathW = printCachePathW + L"\\" + validNameW;
		//CStringW filePathW = printCachePathW + validNameW;  // FIXME

		const char* data = outbuf->Data();
		int dataLength = outbuf->Count();


		PWSTR ext = PathFindExtension(filePathW);
		if (GdiUtils::IsSupportedPictureFileNameExtension(filePathW) || (ext == 0))
		{
			int imageWidth = 0;
			int imageHeight = 0;

			int retRotate = GdiUtils::AutoRotateImageFromMemory(data, dataLength, filePathW, imageWidth, imageHeight);

			_ASSERTE((imageWidth > 0) && (imageHeight > 0));

			AttachmentData* attachmentInfo = attachmentDB.GetAttachmentData(resultW);
			if (attachmentInfo)
			{
				attachmentInfo->m_width = imageWidth;
				attachmentInfo->m_height = imageHeight;
			}

			if (retRotate != 0)
				ret = FileUtils::Write2File(filePathW, (unsigned char*)data, dataLength);  // FIXMEFIXME
		}
		else
		{
			ret = FileUtils::Write2File(filePathW, (unsigned char*)data, dataLength);  // FIXMEFIXME
		}

		MboxMail::rel_tmpbuf();
	}

	if (fpm_save == 0)
		mboxFp.Close();
	return 1;
}

void NListView::GetExtendedMailId(MboxMail* m, CString &extendedId)
{
	MyCTime tt(m->m_timeDate);

	CString strDate;

	CString format;
	format = L"%Y%m%d-%H%M%S-";
	strDate = tt.FormatLocalTm(format);

	CString uID;
	uID.Format(L"%07d", m->m_index);

	extendedId.Append(strDate);
	extendedId.Append(uID);
	extendedId.Append(L" ");

	return;
}

void EmbededImagesStats::Clear()
{
	m_EmbededImagesNoMatch;

	m_FoundMHtml;
	m_FoundMHtmlHtml;
	m_FoundUnexpectedMHtml = 0;

	m_Found = 0;
	m_FoundCid = 0;
	m_FoundHttp = 0;
	m_FoundHttps = 0;
	m_FoundMHtmlHttp = 0;
	m_FoundMHtmlHttps = 0;
	m_FoundData = 0;
	m_FoundLocalFile = 0;
	//
	m_NotFound = 0;
	m_NotFoundCid = 0;
	m_NotFoundHttp = 0;
	m_NotFoundHttps = 0;
	m_FoundMHtmlHttp = 0;
	m_NotFoundMHtmlHttps = 0;
	m_NotFoundData = 0;
	m_NotFoundLocalFile = 0;
}

int NListView::CreateEmbeddedImageFilesEx(CFile& fpm, int mailPosition, CString& imageCachePath, BOOL createEmbeddedImageFiles)
{
	CMainFrame* pFrame = 0;
	char* bdy;
	int bdylen;

	pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (!pFrame)
		return -1;

	NMsgView* pMsgView = pFrame->GetMsgView();
	if (!pMsgView)
		return -1;

	NListView* pListView = pFrame->GetListView();
	if (!pListView)
		return -1;

	MboxMail* m = 0;
	if ((mailPosition < MboxMail::s_mails.GetCount()) && (mailPosition >= 0))
		m = MboxMail::s_mails[mailPosition];
	else
		return -1;

	MailBodyContent* body;
	bool foundImage = false;
	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];

		CStringA contentTypeExtension;
		CStringA contentTypeMain;
		int pos = body->m_contentType.ReverseFind('/');
		if (pos > 0)
		{
			contentTypeExtension = body->m_contentType.Mid(pos + 1);
			contentTypeMain = body->m_contentType.Left(pos);
		}
		// Should just compare against "image" and ..
		if (contentTypeMain.CompareNoCase("text") != 0) {
			foundImage = true;
			break;
		}
	}
	if (!foundImage)
		return 1;

	// Uses global buffers:
	//    MboxMail::m_outdata
	//    MboxMail::m_workbuf;
	// make sure they are not in use already  !!!!
	//
	SimpleString* outbuflarge = MboxMail::m_outdata;

	outbuflarge->Clear();
	UINT pageCode = 0;
	int textType = 1; // try first Html
	int plainTextMode = 0;  // no extra img tags; html text has img tags already
	int textlen = MboxMail::GetMailBody_mboxview(fpm, mailPosition, outbuflarge, pageCode, textType, plainTextMode);  // returns pageCode
	if (textlen != outbuflarge->Count())
		int deb = 1;

	if (outbuflarge->Count() > 0)
	{
		// mail has one or more HTML blocks
		char* inData = outbuflarge->Data();
		int inDataLen = outbuflarge->Count();

		SimpleString* outbuf = MboxMail::m_workbuf;
		outbuf->ClearAndResize(outbuflarge->Count() * 2);

		AttachmentMgr embededAttachmentDB;

		CString absoluteSrcImgFilePath = CMainFrame::GetMboxviewTempPath();
		CString relativeSrcImgFilePath = L".\\ImageCache";

		CString* srcImgFilePath = 0;
		CStringA embededSrcImgFilePath = ".\\";
		BOOL verifyAttachmentDataAsImageType = FALSE;  // FIXME
		BOOL insertMaxWidthForImgTag = TRUE;
		CStringA maxWidth = "100%";
		CStringA maxHeight = "";
		BOOL makeFileNameUnique = TRUE;
		BOOL makeAbsoluteImageFilePath = m_fullImgFilePath;
		NListView::UpdateInlineSrcImgPathEx(&fpm, inData, inDataLen, outbuf, makeFileNameUnique, makeAbsoluteImageFilePath,
			relativeSrcImgFilePath, absoluteSrcImgFilePath, embededAttachmentDB,
			pListView->m_EmbededImagesStats, mailPosition, createEmbeddedImageFiles, verifyAttachmentDataAsImageType, insertMaxWidthForImgTag,
			maxWidth, maxHeight);

		if (outbuf->Count())
		{
			bdy = outbuf->Data();
			bdylen = outbuf->Count();
		}
		else
		{
			bdy = inData;
			bdylen = inDataLen;
		}
	}
	return 1;
}


BOOL NListView::ConfigureExportOfMails()
{
	CMainFrame* pFrame = 0;

	pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (!pFrame)
		return -1;

	CString helpText = L"Configure how to open linked documents\n\n\n";

	if (pFrame->m_HdrFldConfig.m_bHdrAttachmentLinkOpenMode == 0)
		helpText.Append(L"Current configuration is to open linked document in\n\n    \"the same Tab\"\n\n");
	else
		helpText.Append(L"Current configuration is to open linked document in\n\n    \"a new Tab\"\n\n");

	helpText.Append(L"Select  \"Yes\"  to open the linked document in the same Tab as it was clicked\n\n"
		L"Select  \"No\"  to open the linked document in a new Tab\n\n"
		L"Selection is persistent\n\n"
		L"Open Help option for additional details\n\n"
		);

	ResHelper::TranslateString(helpText);

	HWND h = GetSafeHwnd();
	int answer = ::MessageBox(h, helpText, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_YESNOCANCEL);
	BOOL bHdrAttachmentLinkOpenMode = -1;
	if (answer == IDYES)
	{
		bHdrAttachmentLinkOpenMode = 0;  // open in current Tab
	}
	else if (answer == IDNO)
	{
		bHdrAttachmentLinkOpenMode = 1;  // Open un new Tab
	}
	if (bHdrAttachmentLinkOpenMode >= 0)
	{
		CString section_hdr = CString(sz_Software_mboxview) + L"\\PrintConfig\\HeaderFields";

		pFrame->m_HdrFldConfig.m_bHdrAttachmentLinkOpenMode = bHdrAttachmentLinkOpenMode;

		CString param = L"AttachmentLinkOpenMode";
		BOOL ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_hdr, param, bHdrAttachmentLinkOpenMode);

		return TRUE;
	}
	return FALSE;
}

LRESULT NListView::OnCmdParam_OnSwitchWindow(WPARAM wParam, LPARAM lParam)
{
	TRACE(L"NListView::OnCmdParam_OnSwitchWindow\n");
	CWnd *wnd = CMainFrame::SetWindowFocus(&m_list);
	
	return 0;
}

BOOL NListView::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}
