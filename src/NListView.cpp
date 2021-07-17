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


#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

///////
// Kept adding and adding Print to functions but now cleanup is needed, better reusability, possible abstractions, error handling, etc
// Postponed to the next relase 1.0.3.3 since larger effort is needed
///////

#ifdef _DEBUG
inline void BreakNlistView()
{
	int deb = 1;
}
#else
inline void BreakNlistView() {}
#endif

BOOL CreateInlineImageCache_WorkerThread(LPCSTR cache, BOOL mainThread, CString &errorText);
BOOL CreateAttachmentCache_WorkerThread(LPCSTR cache, BOOL mainThread, CString &errorText);
BOOL CreateEmlCache_WorkerThread(LPCSTR cache, BOOL mainThread, CString &errorText);

bool ALongRightProcessProcForwardMails(const CUPDUPDATA* pCUPDUPData)
{
	FORWARD_MAILS_ARGS *args = (FORWARD_MAILS_ARGS*)pCUPDUPData->GetAppData();
	MboxMail::pCUPDUPData = pCUPDUPData;

	HANDLE h = GetCurrentThread();
	BOOL prio = SetThreadPriority(h, THREAD_PRIORITY_ABOVE_NORMAL);
	DWORD myThreadId = GetCurrentThreadId();
	DWORD myThreadPri = GetThreadPriority(h);
	TRACE(_T("(ALongRightProcessProcForwardMails) threadId=%ld threadPriority=%ld\n"), myThreadId, myThreadPri);

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
	TRACE(_T("(ALongRightProcessProcWriteIndexFile) threadId=%ld threadPriority=%ld\n"), myThreadId, myThreadPri);

	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as param to printMailArchiveToTextFile

	BOOL mainThread = FALSE;
	args->ret = SaveMails(args->cache, mainThread, args->errorText);

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
	TRACE(_T("(ALongRightProcessProcWriteAttachmentFile) threadId=%ld threadPriority=%ld\n"), myThreadId, myThreadPri);

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
	TRACE(_T("(ALongRightProcessProcWriteEmlFile) threadId=%ld threadPriority=%ld\n"), myThreadId, myThreadPri);

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
	TRACE(_T("(ALongRightProcessProcWriteInlineImageFile) threadId=%ld threadPriority=%ld\n"), myThreadId, myThreadPri);

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
	TRACE(_T("(ALongRightProcessProcPrintMailGroupToSeparatePDF) threadId=%ld threadPriority=%ld\n"), myThreadId, myThreadPri);

	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as param to printMailArchiveToTextFile

	if (args->selectedMailIndexList)
	{
		args->ret = args->lview->PrintMailSelectedToSeparatePDF_WorkerThread(args->selectedMailIndexList, args->targetPrintSubFolderName, args->targetPrintFolderPath, args->errorText);
	}
	else
	{
		args->ret = args->lview->PrintMailRangeToSeparatePDF_WorkerThread(args->firstMail, args->lastMail, args->targetPrintSubFolderName, args->targetPrintFolderPath, args->errorText);
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
	TRACE(_T("(ALongRightProcessProcPrintMailGroupToSinglePDF) threadId=%ld threadPriority=%ld\n"), myThreadId, myThreadPri);

	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as parameter

	if (args->selectedMailIndexList)
	{
		args->ret = args->lview->PrintMailSelectedToSinglePDF_WorkerThread(args->selectedMailIndexList, args->targetPrintSubFolderName, args->targetPrintFolderPath, args->errorText);
	}
	else
	{
		args->ret = args->lview->PrintMailRangeToSinglePDF_WorkerThread(args->firstMail, args->lastMail, args->targetPrintSubFolderName, args->targetPrintFolderPath, args->errorText);
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
	TRACE(_T("(ALongRightProcessProcPrintMailGroupToSeparateHTML) threadId=%ld threadPriority=%ld\n"), myThreadId, myThreadPri);

	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as param to printMailArchiveToTextFile
	if (args->selectedMailIndexList)
	{
		args->ret = args->lview->PrintMailSelectedToSeparateHTML_WorkerThread(args->selectedMailIndexList, args->targetPrintSubFolderName, args->targetPrintFolderPath, args->errorText);
	}
	else
	{
		args->ret = args->lview->PrintMailRangeToSeparateHTML_WorkerThread(args->firstMail, args->lastMail, args->targetPrintSubFolderName, args->targetPrintFolderPath, args->errorText);
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

	TRACE(_T("(ALongRightProcessProcPrintMailGroupToSingleHTML) threadId=%ld threadPriority=%ld\n"), myThreadId, myThreadPri);

	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as param to printMailArchiveToTextFile
	if (args->selectedMailIndexList)
	{
		args->ret = args->lview->PrintMailSelectedToSingleHTML_WorkerThread(args->selectedMailIndexList, args->targetPrintSubFolderName, args->targetPrintFolderPath, args->errorText);
	}
	else
	{
		args->ret = args->lview->PrintMailRangeToSingleHTML_WorkerThread(args->firstMail, args->lastMail, args->targetPrintSubFolderName, args->targetPrintFolderPath, args->errorText);
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
	TRACE(_T("(ALongRightProcessProcPrintMailGroupToSingleTEXT) threadId=%ld threadPriority=%ld\n"), myThreadId, myThreadPri);

	Com_Initialize();

	// TODO: CUPDUPDATA* pCUPDUPData is global set as  MboxMail::pCUPDUPData = pCUPDUPData; Should conider to pass as param to printMailArchiveToTextFile
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

NListView::NListView() : m_list(this), m_lastStartDate((time_t)-1), m_lastEndDate((time_t)-1),
		m_mboxMailStartDate((time_t)-1), m_mboxMailEndDate((time_t)-1)
{
	ResetFileMapView();

	m_which = NULL;

	m_bApplyColorStyle = TRUE;
	m_bLongMailAddress = TRUE;
	m_name = new SimpleString(64);
	m_addr = new SimpleString(64);

	m_searchStringInMail.Empty();
	m_bCaseSensInMail = FALSE;
	m_bWholeWordInMail = FALSE;

	m_bNeedToFindMailMinMaxTime = TRUE;
	m_needToRestoreArchiveListDateTime = FALSE;

	m_advancedParams.SetDflts();
	m_advancedFind = FALSE;
	m_gmtTime = 0;
	m_bStartSearchAtSelectedItem = 0; // FALSE; TODO: is this desired feature ?
	m_bFindNext = TRUE;
	m_bFindAll = FALSE;
	m_bFindAllMailsThatDontMatch = FALSE;
	m_findAllCount = 0;
	m_bEditFindFirst = FALSE;  // must call OnEditFind() first and not OnEditFindAgain()
	m_lastFindPos = -1;  // last find position, start for the next find again
	m_searchString.Empty();
	m_bCaseSens = FALSE;
	m_bWholeWord = FALSE;
	m_bFrom = TRUE;
	m_bTo = TRUE;
	m_bCC = FALSE;
	m_bBCC = FALSE;
	m_bSubject = TRUE;
	m_bContent = FALSE;
	m_bAttachments = FALSE;
	m_bAttachmentName = FALSE;
	m_bHighlightAll = FALSE;
	m_bHighlightAllSet = FALSE;
	m_filterDates = FALSE;
	m_lastSel = -1;  // last/currently selected item
	m_bInFind = FALSE;
	int iFormat = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "format");
	m_format = MboxMail::GetDateFormat(iFormat);
	m_gmtTime = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "timeType");

	CString exportEML;
	BOOL retval = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("exportEML"), exportEML);
	if (retval == TRUE) {
		if (exportEML.Compare(_T("y")) == 0)
			m_bExportEml = TRUE;
		else
			m_bExportEml = FALSE;
	}
	else {
		exportEML = _T("n");
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("exportEML"), exportEML);
		m_bExportEml = FALSE;
	}

	DWORD barDelay;
	retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("progressBarDelay"), barDelay);
	if (retval == TRUE) {
		m_maxSearchDuration = barDelay;
	}
	else {
		barDelay = 1;
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("progressBarDelay"), barDelay);
		m_maxSearchDuration = barDelay;
	}

	m_frameCx_TreeNotInHide = 700;
	m_frameCy_TreeNotInHide = 200;
	m_frameCx_TreeInHide = 700;
	m_frameCy_TreeInHide = 200;


	CString m_section = CString(sz_Software_mboxview) + "\\" + "Window Placement";

	BOOL ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "ListFrameTreeNotHiddenWidth", m_frameCx_TreeNotInHide);
	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "ListFrameTreeNotHiddenHeight", m_frameCy_TreeNotInHide);
	//
	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "ListFrameTreeHiddenWidth", m_frameCx_TreeInHide);
	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "ListFrameTreeHiddenHeight", m_frameCy_TreeInHide);

	m_acp = GetACP();
	m_tcpPort = 61333;
	m_enbaleForwardMailsLog = FALSE;
	m_enbaleSMTPProtocolLog = FALSE;
	m_developmentMode = FALSE;
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
END_MESSAGE_MAP()

int g_pointSize = 85;
CString g_fontName = "Tahoma";

void NListView::ResetFont()
{
	m_boldFont.DeleteObject();
	m_font.DeleteObject();
	if( ! m_font.CreatePointFont (g_pointSize, g_fontName) )
		m_font.CreatePointFont (85, _T("Arial"));
	LOGFONT	lf;
	m_font.GetLogFont(&lf);
	lf.lfWeight = FW_BOLD;
	m_boldFont.CreateFontIndirect(&lf);
    m_list.SetFont(&m_font);
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
		CString path = FileUtils::GetmboxviewTempPath();

		if(FileUtils::PathDirExists(path) )
			ShellExecute(NULL, _T("open"), path, NULL,NULL, SW_SHOWNORMAL );
	}
}

void NListView::OnActivating(NMHDR* pNMHDR, LRESULT* pResult) 
{
	NMITEMACTIVATE *pnm = (NMITEMACTIVATE *)pNMHDR;
	if( m_lastSel != pnm->iItem ) {
		TRACE("Selecting %d\n", pnm->iItem);
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
				"Mail Context Menu is not available while View Message Headers are enabled.\n\n"
				"Please disable viewing of raw mail headers to reenable the context menu.\n\n"
				"UnCheck \"View -> View Message Headers\"\n"
				"and Right click on Mail Header Pane at the top of the Message Window and\n"
				"UnCheck \"View Message Header\"\n";
			int answer = MessageBox(InfoTxt, "Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
			return;
		}
	}


	int iItem = pnm->iItem;

	if (iItem < 0)
		return;

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

	// Create enums or replace switch statment with if else ..
	// XXX_GROUP_Id represents group of related emails, related == convesation
	const UINT S_TEXT_Id = 1;
	AppendMenu(&printToSubMenu, S_TEXT_Id, _T("Text.."));

	const UINT S_HTML_Id = 2;
	AppendMenu(&printToSubMenu, S_HTML_Id, _T("HTML..."));

	const UINT S_PRINTER_Id = 3;
	AppendMenu(&printToSubMenu, S_PRINTER_Id, _T("Printer..."));

	const UINT S_PDF_DIRECT_Id = 4;
	AppendMenu(&printToSubMenu, S_PDF_DIRECT_Id, _T("PDF..."));

	menu.AppendMenu(MF_POPUP | MF_STRING, (UINT)printToSubMenu.GetSafeHmenu(), _T("Print To"));
	menu.AppendMenu(MF_SEPARATOR);

	const UINT S_TEXT_GROUP_Id = 5;
	AppendMenu(&printGroupToSubMenu, S_TEXT_GROUP_Id, _T("Text.."));

	const UINT S_HTML_GROUP_Id = 6;
	AppendMenu(&printGroupToSubMenu, S_HTML_GROUP_Id, _T("HTML..."));

	const UINT S_PRINTER_GROUP_Id = 7;
	AppendMenu(&printGroupToSubMenu, S_PRINTER_GROUP_Id, _T("Printer..."));

	const UINT S_PDF_DIRECT_GROUP_Id = 8;
	AppendMenu(&printGroupToSubMenu, S_PDF_DIRECT_GROUP_Id, _T("PDF..."));

	const UINT S_CSV_GROUP_Id = 9;
	AppendMenu(&printGroupToSubMenu, S_CSV_GROUP_Id, _T("CSV..."));

	menu.AppendMenu(MF_POPUP | MF_STRING, (UINT)printGroupToSubMenu.GetSafeHmenu(), _T("Print Related Mails To"));
	menu.AppendMenu(MF_SEPARATOR);

	const UINT S_HTML_OPEN_Id = 10;
	AppendMenu(&menu, S_HTML_OPEN_Id, _T("Open in Browser"));

	const UINT S_HTML_OPEN_RELATED_Id = 11;
	AppendMenu(&menu, S_HTML_OPEN_RELATED_Id, _T("Open Related Mails in Browser"));

	const UINT S_HTML_OPEN_RELATED_FILES_Id = 12;
	AppendMenu(&menu, S_HTML_OPEN_RELATED_FILES_Id, _T("Open Related Files Location"));

	const UINT S_HTML_FIND_Id = 13;
	AppendMenu(&menu, S_HTML_FIND_Id, _T("Find Text"));

	const UINT S_ADVANCED_FIND_Id = 14;
	AppendMenu(&menu, S_ADVANCED_FIND_Id, _T("Find Advanced"));

	const UINT S_REMOVE_Id = 15;
	if (pFrame && (MboxMail::IsFindMailsSelected() || MboxMail::IsUserMailsSelected())) {
		menu.AppendMenu(MF_SEPARATOR);
		AppendMenu(&menu, S_REMOVE_Id, _T("Remove Selected Mail"));
	}

	const UINT S_REMOVE_ALL_Id = 16;
	if (pFrame && (MboxMail::IsFindMailsSelected() || (MboxMail::IsUserMailsSelected()) && (pFrame->IsUserMailsListEnabled()))) {
		AppendMenu(&menu, S_REMOVE_ALL_Id, _T("Remove All Mails"));
	}

	const UINT S_COPY_ALL_Id = 17;
	const UINT S_COPY_SELECTED_Id = 18;
	if (pFrame && (MboxMail::IsAllMailsSelected() || MboxMail::IsLabelMailsSelected() || MboxMail::IsFindMailsSelected()) && pFrame->IsUserMailsListEnabled()) {
		menu.AppendMenu(MF_SEPARATOR);
		AppendMenu(&menu, S_COPY_ALL_Id, _T("Copy All into User Selected Mails"));
		AppendMenu(&menu, S_COPY_SELECTED_Id, _T("Copy Selected into User Selected Mails"));
	}

	const UINT S_SAVE_AS_ARCHIVE_Id = 19;
	const UINT M_ARCHIVE_LOCATION_Id = 20;
	if (pFrame && (MboxMail::IsFindMailsSelected() || MboxMail::IsUserMailsSelected())) {
		menu.AppendMenu(MF_SEPARATOR);
		AppendMenu(&menu, S_SAVE_AS_ARCHIVE_Id, _T("Save as Mbox Mail Archive file"));
		AppendMenu(&menu, M_ARCHIVE_LOCATION_Id, _T("Open Mbox Mail Archive Location"));
	}

	const UINT S_SAVE_AS_LIST_Id = 21;
	const UINT S_RELOAD_LIST_Id = 22;
	const UINT M_LIST_LOCATION_Id = 23;
	if (pFrame && (MboxMail::IsUserMailsSelected())) {
		menu.AppendMenu(MF_SEPARATOR);
		AppendMenu(&menu, S_SAVE_AS_LIST_Id, _T("Save as Mail List file"));
		AppendMenu(&menu, S_RELOAD_LIST_Id, _T("Reload from Mail List file"));
		AppendMenu(&menu, M_LIST_LOCATION_Id, _T("Open Mail List Location"));
	}

	const UINT S_RESTORE_LIST_Id = 24;
	if (pFrame && MboxMail::IsAllMailsSelected() && pFrame->IsUserMailsListEnabled() && (MboxMail::s_mails_edit.GetCount() == 0)) {
		//menu.AppendMenu(MF_SEPARATOR);
		AppendMenu(&menu, S_RESTORE_LIST_Id, _T("Restore User Selected List from Mail List file"));
	}

	const UINT M_REMOVE_DUPLICATE_MAILS_Id = 25;
	if (pFrame && (MboxMail::IsUserMailsSelected() || MboxMail::IsLabelMailsSelected())) {
		menu.AppendMenu(MF_SEPARATOR);
		AppendMenu(&menu, M_REMOVE_DUPLICATE_MAILS_Id, _T("Remove Duplicate Mails"));
	}

	const UINT M_COPY_MAILS_TO_FOLDERS_Id = 26;
	if (pFrame && !MboxMail::IsFolderMailsSelected()) {
		//menu.AppendMenu(MF_SEPARATOR);
		//AppendMenu(&menu, M_COPY_MAILS_TO_FOLDERS_Id, _T("Copy to Folders"));
		;
	}

	const UINT M_FORWARD_MAIL_Id = 27;
	AppendMenu(&menu, M_FORWARD_MAIL_Id, _T("Forward Mail"));

	const UINT M_FORWARD_RELATED_MAILS_Id = 28;
	AppendMenu(&menu, M_FORWARD_RELATED_MAILS_Id, _T("Forward Related Mails"));

	UINT command = menu.TrackPopupMenuEx(TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, this, 0);

	DWORD error = GetLastError();

	UINT nFlags = TPM_RETURNCMD;
	CString menuString;
	int chrCnt = menu.GetMenuString(command, menuString, nFlags);

	BOOL multipleSelectedMails = FALSE;
	BOOL itemSelected = FALSE;
	switch (command)
	{
	case S_HTML_Id: {
		CString fileName;
		if (pFrame) {
			pFrame->OnPrintSingleMailtoText(iItem, 1, fileName);
		}
		int deb = 1;
	}
	break;
	case S_TEXT_Id: {
		CString fileName;
		if (pFrame) {
			pFrame->OnPrintSingleMailtoText(iItem, 0, fileName);
		}
		int deb = 1;
	}
	break;
	case S_PRINTER_Id: {
		CString fileName;
		if (pFrame) {
			MboxMail::ShowHint(HintConfig::PrintToPrinterHint, GetSafeHwnd());
			pFrame->OnPrintSingleMailtoText(iItem, 1, fileName, FALSE, TRUE);
		}
		int deb = 1;
	}
	break;
	case S_PDF_DIRECT_Id: {
		MboxMail::ShowHint(HintConfig::PrintToPDFHint, GetSafeHwnd());
		CString errorText;
		CString targetPrintSubFolderName;
		CString targetPrintFolderPath;
		int ret = PrintMailRangeToSinglePDF_Thread(iItem, iItem, targetPrintSubFolderName);
		int deb = 1;
	}
	break;
	case S_CSV_GROUP_Id: {
		int ret = PrintMailRangeToSingleCSV_Thread(iItem);
		int deb = 1;
	}
	break;
	case S_HTML_GROUP_Id: {
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
		int deb = 1;
	}
	break;
	case S_TEXT_GROUP_Id: {
		PrintMailGroupToText(multipleSelectedMails, iItem, 0);
	}
	break;
	case S_PRINTER_GROUP_Id: {
		MboxMail::ShowHint(HintConfig::PrintToPrinterHint, GetSafeHwnd());
		PrintMailGroupToText(multipleSelectedMails, iItem, 1, FALSE, TRUE);
	}
	break;
	case S_PDF_DIRECT_GROUP_Id: {
		if (pFrame)
		{
			MboxMail::ShowHint(HintConfig::PrintToPDFHint, GetSafeHwnd());
			if (pFrame->m_NamePatternParams.m_bPrintToSeparatePDFFiles)
			{
				CString errorText;
				int ret = PrintMailConversationToSeparatePDF_Thread(iItem, errorText);
				int deb = 1;
			}
			else
			{
				CString errorText;
				int ret = PrintMailConversationToSinglePDF_Thread(iItem, errorText);
				int deb = 1;
			}
		}
		int deb = 1;
	}
	break;
	case S_HTML_OPEN_Id: {
		CString fileName;
		if (pFrame) {
			BOOL forceOpen = TRUE;
			pFrame->OnPrintSingleMailtoText(iItem, 1, fileName, forceOpen);
		}
	}
	break;
	case S_HTML_OPEN_RELATED_Id: {
			BOOL forceOpen = TRUE;
			PrintMailGroupToText(multipleSelectedMails, iItem, 1, forceOpen);
	}
	break;

	case S_HTML_OPEN_RELATED_FILES_Id: {
		CString path = FileUtils::GetmboxviewTempPath();

		if (FileUtils::PathDirExists(path)) {
			HWND h = GetSafeHwnd();
			HINSTANCE result = ShellExecute(h, _T("open"), path, NULL, NULL, SW_SHOWNORMAL);
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
	case S_REMOVE_Id: {
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
	case S_COPY_ALL_Id: {
		CopyAllMails();
		VerifyMailsOnUserSelectsMailListMarkCounts();
	}
	break;
	case S_REMOVE_ALL_Id: {
		CString txt = _T("Do you want to remove all mails?");
		HWND h = GetSafeHwnd();
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDYES) {
			RemoveAllMails();
		}
		VerifyMailsOnUserSelectsMailListMarkCounts();
	}
	break;
	case S_SAVE_AS_ARCHIVE_Id: {
		//SaveAsMboxAndAsMboxlistFile();
		SaveAsMboxArchiveFile_v2();
	}
	break;
	case S_SAVE_AS_LIST_Id: {
		SaveAsMboxListFile_v2();
	}
	break;
	case S_COPY_SELECTED_Id: {
		CopySelectedMails();
		VerifyMailsOnUserSelectsMailListMarkCounts();
	}
	break;
	case S_ADVANCED_FIND_Id: {
		RunFindAdvancedOnSelectedMail(iItem);
		MboxMail::ShowHint(HintConfig::AdvancedFindDialogHint, GetSafeHwnd());
	}
	break;
	case S_RELOAD_LIST_Id: {
		ReloadMboxListFile_v2();
		RefreshMailsOnUserSelectsMailListMark();
	}
	break;
	case S_RESTORE_LIST_Id: {
		ReloadMboxListFile_v2();
		RefreshMailsOnUserSelectsMailListMark();
	}
	break;
	case M_ARCHIVE_LOCATION_Id: {
		OpenArchiveFileLocation();
	}
	break;
	case M_LIST_LOCATION_Id: {
		OpenMailListFileLocation();
	}
	break;
	case M_REMOVE_DUPLICATE_MAILS_Id: {
		if (MboxMail::IsLabelMailsSelected())
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
	case M_COPY_MAILS_TO_FOLDERS_Id: {
		int retval = CopyMailsToFolders();
	}
	break;
	case M_FORWARD_MAIL_Id: {
		//MboxMail::ShowHint(HintConfig::ForwardMailHint, GetSafeHwnd());
		ForwardSelectedMails(iItem);
		int deb = 1;
	}
	break;
	case M_FORWARD_RELATED_MAILS_Id: {
		//MboxMail::ShowHint(HintConfig::ForwardMailHint, GetSafeHwnd());
		int ret = ForwardMailRange(iItem);
		int deb = 1;
	}
	break;
	default: {
		int deb = 1;
	}
	break;
	}

	// TODO: review below check; set itemSelected to TRUE ??
	if ((command == S_REMOVE_Id) || 
		(command == M_REMOVE_DUPLICATE_MAILS_Id) || 
		(command == S_REMOVE_ALL_Id) || 
		(command == S_RELOAD_LIST_Id) ||
		(command == S_RESTORE_LIST_Id) ||
		(command == S_ADVANCED_FIND_Id) 
		)
		; // done
	else if ((itemSelected == FALSE) && (m_lastSel != iItem)) {
		TRACE("Selecting %d\n", iItem);
		m_bApplyColorStyle = TRUE;
		SelectItem(iItem);
	}

	Invalidate();
	UpdateWindow();

	*pResult = 0;
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

	// Create enums or replace switch statment with if else ..

	const UINT S_TEXT_GROUP_Id = 4;
	AppendMenu(&printGroupToSubMenu, S_TEXT_GROUP_Id, _T("Text.."));

	const UINT S_HTML_GROUP_Id = 5;
	AppendMenu(&printGroupToSubMenu, S_HTML_GROUP_Id, _T("HTML..."));

	const UINT S_PRINTER_GROUP_Id = 6;
	AppendMenu(&printGroupToSubMenu, S_PRINTER_GROUP_Id, _T("Printer..."));

	const UINT S_PDF_DIRECT_Id = 28;
	AppendMenu(&printGroupToSubMenu, S_PDF_DIRECT_Id, _T("PDF..."));

	const UINT S_CSV_GROUP_Id = 29;
	AppendMenu(&printGroupToSubMenu, S_CSV_GROUP_Id, _T("CSV..."));

	menu.AppendMenu(MF_POPUP | MF_STRING, (UINT)printGroupToSubMenu.GetSafeHmenu(), _T("Print Selected Mails To"));
	menu.AppendMenu(MF_SEPARATOR);

	const UINT S_HTML_OPEN_RELATED_Id = 7;
	AppendMenu(&menu, S_HTML_OPEN_RELATED_Id, _T("Open Selected Mails in Browser"));

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	const UINT S_REMOVE_SELECTED_Id = 8;
	if (pFrame && (MboxMail::IsFindMailsSelected() || MboxMail::IsUserMailsSelected())) {
		AppendMenu(&menu, S_REMOVE_SELECTED_Id, _T("Remove Selected Mails"));
	}

	const UINT S_COPY_SELECTED_Id = 22;
	if (pFrame && (MboxMail::IsAllMailsSelected() || MboxMail::IsFindMailsSelected() || MboxMail::IsLabelMailsSelected())) {
		if (pFrame->IsUserMailsListEnabled()) {
			AppendMenu(&menu, S_COPY_SELECTED_Id, _T("Copy Selected into User Selected Mails"));
		}
	}

	const UINT S_FORWARD_SELECTED_MAILS_Id = 23;
	AppendMenu(&menu, S_FORWARD_SELECTED_MAILS_Id, _T("Forward Selected Mails"));
	

	// Used above for printing to PDF and CSV
	//const UINT S_PDF_DIRECT_Id = 28;
	//const UINT S_CSV_GROUP_Id = 29;

	UINT command = menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, this);

	UINT nFlags = TPM_RETURNCMD;
	CString menuString;
	int chrCnt = menu.GetMenuString(command, menuString, nFlags);

	BOOL multipleSelectedMails = TRUE;
	BOOL itemSelected = FALSE;
	switch (command)
	{
	case S_HTML_GROUP_Id: {
		if (pFrame)
		{
			if (pFrame->m_NamePatternParams.m_bPrintToSeparateHTMLFiles)
			{
				CString errorText;
				CString targetPrintSubFolderName = "HTML_GROUP";
				CString targetPrintFolderPath;
				int ret = PrintMailSelectedToSeparateHTML_Thread(targetPrintSubFolderName, targetPrintFolderPath);
				int deb = 1;
			}
			else
			{
				CString errorText;
				CString targetPrintSubFolderName;
				CString targetPrintFolderPath;
				int ret = PrintMailSelectedToSingleHTML_Thread(targetPrintSubFolderName, targetPrintFolderPath);
				int deb = 1;
			}
		}
		int deb = 1;
	}
	break;
	case S_TEXT_GROUP_Id: {
		//PrintMailGroupToText(multipleSelectedMails, iItem, 0);
		if (pFrame)
		{
			if (pFrame->m_NamePatternParams.m_bPrintToSeparateTEXTFiles)
			{
				CString errorText;
				CString targetPrintSubFolderName = "TEXT_GROUP";
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
	case S_PRINTER_GROUP_Id: {
		MboxMail::ShowHint(HintConfig::PrintToPrinterHint, GetSafeHwnd());
		PrintMailGroupToText(multipleSelectedMails, iItem, 1, FALSE, TRUE);
	}
	break;
	case S_PDF_DIRECT_Id: {
		if (pFrame)
		{
			MboxMail::ShowHint(HintConfig::PrintToPDFHint, GetSafeHwnd());
			if (pFrame->m_NamePatternParams.m_bPrintToSeparatePDFFiles)
			{
				CString targetPrintSubFolderName = "PDF_GROUP";
				CString targetPrintFolderPath;
				int ret = PrintMailSelectedToSeparatePDF_Thread(targetPrintSubFolderName, targetPrintFolderPath);
			}
			else 
			{
				CString targetPrintSubFolderName;
				CString targetPrintFolderPath;
				int ret = PrintMailSelectedToSinglePDF_Thread(targetPrintSubFolderName, targetPrintFolderPath);
				int deb = 1;
			}
			int deb = 1;
		}
		int deb = 1;
	}
	break;
	case S_HTML_OPEN_RELATED_Id: {
		BOOL forceOpen = TRUE;
		PrintMailGroupToText(multipleSelectedMails, iItem, 1, forceOpen);
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
	case S_COPY_SELECTED_Id: {
		CopySelectedMails();
		VerifyMailsOnUserSelectsMailListMarkCounts();
	}
	break;
	case S_CSV_GROUP_Id: {
		if (pFrame) {
			BOOL selectedMails = TRUE;
			int firstMail = 0;
			int lastMail = MboxMail::s_mails.GetCount() - 1;
			pFrame->PrintMailsToCSV(firstMail, lastMail, selectedMails);
		}
		int deb = 1;
	}
	break;
	case S_FORWARD_SELECTED_MAILS_Id: {
		//MboxMail::ShowHint(HintConfig::ForwardMailHint, GetSafeHwnd());
		ForwardSelectedMails(iItem);
		int deb = 1;
	}
	break;

	default: {
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
	NMLISTVIEW * pLV = (NMLISTVIEW*)pNMHDR;

	*pResult = 0;

	SortByColumn(pLV->iSubItem);
}


void NListView::SortByColumn(int colNumber, BOOL sortByPosition) // use sortByPosition fro testing
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
			colStr = "Conversations";
		else if (colNumber == 1)
			colStr = "Date";
		else if (colNumber == 2)
			colStr = "From";
		else if (colNumber == 3)
			colStr = "To";
		else if (colNumber == 4)
			colStr = "Subject";
		else if (colNumber == 5)
			colStr = "Size";

		if ((colNumber == 1) && sortByPosition)  // Bad hack to sort by Id-:)  FIX IT
		{
			colStr = "Id";
		}
		sText.Format("Sorting by %s ...", colStr);
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);
	}

	m_bInFind = true;

	switch (colNumber) {
	case 0: // !
		if (abs(MboxMail::b_mails_which_sorted) == 99) {
			MboxMail::b_mails_which_sorted = -MboxMail::b_mails_which_sorted;
		}
		else {
			MboxMail::b_mails_which_sorted = 99;
		}

		MboxMail::SortByGroupId(0, MboxMail::b_mails_which_sorted < 0);
		MboxMail::assignColor2ConvesationGroups(&MboxMail::s_mails);
		//MboxMail::SortByConverstionGroups(0, MboxMail::b_mails_which_sorted < 0);  // TODO: review
		mustSort = true;
		break;
	case 1: // date
		if (abs(MboxMail::b_mails_which_sorted) == 1) {
			MboxMail::b_mails_which_sorted = -MboxMail::b_mails_which_sorted;
		} else {
			MboxMail::b_mails_which_sorted = 1;
		}
		if (sortByPosition == FALSE)
			MboxMail::SortByDate(0, MboxMail::b_mails_which_sorted < 0);
		else
			MboxMail::SortByFileOffset(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
		break;
	case 2: // from
		if (abs(MboxMail::b_mails_which_sorted) == 2) {
			MboxMail::b_mails_which_sorted = -MboxMail::b_mails_which_sorted;
		} else {
			MboxMail::b_mails_which_sorted = 2;
		}
		MboxMail::SortByFrom(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
		break;
	case 3: // to
		if (abs(MboxMail::b_mails_which_sorted) == 3) {
			MboxMail::b_mails_which_sorted = -MboxMail::b_mails_which_sorted;
		} else {
			MboxMail::b_mails_which_sorted = 3;
		}
		MboxMail::SortByTo(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
		break;
	case 4: // subj
		if (abs(MboxMail::b_mails_which_sorted) == 4) {
			MboxMail::b_mails_which_sorted = -MboxMail::b_mails_which_sorted;
		}
		else {
			MboxMail::b_mails_which_sorted = 4;
		}
		MboxMail::SortBySubject(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
		break;
	case 5: // size
		if (abs(MboxMail::b_mails_which_sorted) == 5) {
			MboxMail::b_mails_which_sorted = -MboxMail::b_mails_which_sorted;
		}
		else {
			MboxMail::b_mails_which_sorted = 5;
		}
		MboxMail::SortBySize(0, MboxMail::b_mails_which_sorted > 0);
		mustSort = true;
		break;
	}
	if (mustSort)
	{
		MboxMail::b_mails_sorted = true;

		m_bContent = FALSE;
		m_bAttachments = FALSE;
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
	else
		MarkColumns();  // not used anymore
	m_bInFind = false;

	if (pFrame)
	{
		int paneId = 0;
		CString sText;
		sText.Format("Ready");
		pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
	}
}


// Refresh mail list, don't update MboxMail::b_mails_which_sorted
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
			colStr = "Conversations";
		else if (colNumber == 1)
			colStr = "Date";
		else if (colNumber == 2)
			colStr = "From";
		else if (colNumber == 3)
			colStr = "To";
		else if (colNumber == 4)
			colStr = "Subject";
		else if (colNumber == 5)
			colStr = "Size";

		sText.Format("Sorting by %s ...", colStr);
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
		MboxMail::SortBySubject(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
		break;
	case 5: // size
		MboxMail::SortBySize(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
		break;
	}
	if (mustSort)
	{
		MboxMail::b_mails_sorted = true;

		m_bContent = FALSE;
		m_bAttachments = FALSE;
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
	else
		MarkColumns();
	m_bInFind = false;

	if (pFrame)
	{
		int paneId = 0;
		CString sText;
		sText.Format("Ready");
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

void NListView::MarkColumns()
{
	// not uysed naymore, replaced by CWheelListCtrl::OnDrawItem
	return;

	CString sDate = "date";
	CString sFrom = "from";
	CString sTo = "to";
	CString sSubj = "subject";
	CString sSize = "size(KB)";
	CString sAppendix = "!";

	if (m_gmtTime == 1)
		sDate += " (GMT)";
	else
		sDate += " (Local)";

	LV_COLUMN c;
	c.mask = LVCF_TEXT;
	if (MboxMail::b_mails_sorted)
	{
		char *indicator = " *";
		int which_sorted = abs(MboxMail::b_mails_which_sorted);
		if (which_sorted == 1)
			sDate += indicator;
		else if (which_sorted == 2)
			sFrom += indicator;
		else if (which_sorted == 3)
			sTo += indicator;
		else if (which_sorted == 4)
			sSubj += indicator;
		else if (which_sorted == 5)
			sSize += indicator;
		else if (which_sorted == 99)
			sAppendix += "*";
	}

	c.pszText = (LPSTR)((LPCSTR)sDate);
	m_list.SetColumn(1, &c);
	c.pszText = (LPSTR)((LPCSTR)sFrom);
	m_list.SetColumn(2, &c);
	c.pszText = (LPSTR)((LPCSTR)sTo);
	m_list.SetColumn(3, &c);
	c.pszText = (LPSTR)((LPCSTR)sSubj);
	m_list.SetColumn(4, &c);
	c.pszText = (LPSTR)((LPCSTR)sSize);
	m_list.SetColumn(5, &c);
	c.pszText = (LPSTR)((LPCSTR)sAppendix);
	m_list.SetColumn(0, &c);
}

/////////////////////////////////////////////////////////////////////////////
// NListView message handlers

int NListView::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CWnd::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	//if( !m_list.Create(WS_CHILD|WS_VISIBLE|LVS_SINGLESEL|LVS_REPORT|LVS_SHOWSELALWAYS|LVS_OWNERDATA, CRect(), this, IDC_LIST) )
	if (!m_list.Create(WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_OWNERDATA, CRect(), this, IDC_LIST))
		return -1;

	m_list.SetExtendedStyle(LVS_EX_FULLROWSELECT);
	m_list.SendMessage((CCM_FIRST + 0x7), 5, 0); // #define CCM_SETVERSION          (CCM_FIRST + 0x7)
	m_list.SetTextColor (::GetSysColor(COLOR_WINDOWTEXT));
	ResetFont();

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if ((pFrame == 0) || (pFrame->GetMessageWindowPosition() == 1))
	{
		m_list.InsertColumn(0, "!", LVCFMT_LEFT, 22, 0);
		m_list.InsertColumn(1, "date", LVCFMT_LEFT, 100, 0);
		m_list.InsertColumn(2, "from", LVCFMT_LEFT, 150, 0);
		m_list.InsertColumn(3, "to", LVCFMT_LEFT, 150, 0);
		m_list.InsertColumn(4, "subject", LVCFMT_LEFT, 400, 0);
		m_list.InsertColumn(5, _T("size(KB)"), LVCFMT_LEFT, 120, 0);
	}
	else
	{
		m_list.InsertColumn(0, "!", LVCFMT_LEFT, 22, 0);
		m_list.InsertColumn(1, "date", LVCFMT_LEFT, 100, 0);
		m_list.InsertColumn(2, "from", LVCFMT_LEFT, 100, 0);
		m_list.InsertColumn(3, "to", LVCFMT_LEFT, 100, 0);
		m_list.InsertColumn(4, "subject", LVCFMT_LEFT, 400, 0);
		m_list.InsertColumn(5, _T("size(KB)"), LVCFMT_LEFT, 120, 0);
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

		TRACE("NListView::OnSize cx=%d cy=%d viewPos=%d IsTreeHideVal=%d IsTreeHidden=%d\n",
			cx, cy, msgViewPosition, bTreeHideVal, isTreeHidden);

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
	int size_len = 60;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if ((pFrame != 0) || (pFrame->GetMessageWindowPosition() != 1))
	{
		min_from_len = 100;
		max_from_len = 400;
		min_to_len = 100;
		max_to_len = 400;
		min_subj_len = 150;
		dflt_subj_len = 400;
		size_len = 60;
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

void NListView::OnGetDispInfo(NMHDR* pNMHDR, LRESULT* pResult) 
{
	LV_DISPINFO* pDispInfo = (LV_DISPINFO*)pNMHDR;
	static const char * ast = "*";
	static const char * nul = "";
	char datebuff[32];
	char sizebuff[32];
	if (pDispInfo->item.mask & LVIF_TEXT)
	{
		MboxMail *m = MboxMail::s_mails[pDispInfo->item.iItem];
//		MboxMail *m = (MboxMail*)pDispInfo->item.lParam;
		const char *t;
		switch (pDispInfo->item.iSubItem) {
		case 0: // date
			/*
			if( m->m_recv == FALSE )
				t = ast;
			else
				t = nul;*/
			if (m->m_hasAttachments)
				t = ast;
			else
				t = nul;
			break;
		case 1: // date
			datebuff[0] = 0;
			if (m->m_timeDate >= 0) 
			{
				MyCTime tt(m->m_timeDate);
				if (!m_gmtTime) 
				{
					CString lDateTime = tt.FormatLocalTm(m_format);
					strcpy(datebuff, (LPCSTR)lDateTime);
				}
				else {
					CString lDateTime = tt.FormatGmtTm(m_format);
					strcpy(datebuff, (LPCSTR)lDateTime);
				}
			}
			t = datebuff;
			break;
		case 2: // from
			t = m->m_from.GetBufferSetLength(m->m_from.GetLength());
			break;
		case 3: // to
			t = m->m_to.GetBufferSetLength(m->m_to.GetLength());
			break;
		case 4: // subj
			t = m->m_subj.GetBufferSetLength(m->m_subj.GetLength());
			break;
		case 5: // size
			{
				int length = m->m_length;
				int kb = length / 1000;
				if (length % 1000)
					kb++;
				sizebuff[0] = 0;
				_itoa(kb, sizebuff, 10);
				t = sizebuff;
			}
			break;
		default:
			t = nul;
			break;
		}
		pDispInfo->item.pszText = (char*)t;
	}
/*	if( pDispInfo->item.mask & LVIF_PARAM ) {
		pDispInfo->item.lParam = (LPARAM)m;
	}*/
	*pResult = 0;
}

void NListView::OnCustomDraw(NMHDR* pNMHDR, LRESULT* pResult)
{
#define MarkColor	RGB(220, 20, 60)

	LPNMLVCUSTOMDRAW lplvcd = (LPNMLVCUSTOMDRAW)pNMHDR;
	NMCUSTOMDRAW   &nmcd = lplvcd->nmcd;
	MboxMail *m;

	DWORD dwItemSpec = nmcd.dwItemSpec;
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
		rItem.iItem = dwItemSpec;
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
		CString FieldText;
		CString Charset;
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
				if (abs(MboxMail::b_mails_which_sorted) == 99)  // mails sorted by comnversations
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
				if (TextUtilsEx::Str2Wide(FieldText, charsetId, strW)) {
					::ExtTextOutW(hDC, xpos, ypos, ETO_CLIPPED, rect, (LPCWSTR)strW, strW.GetLength(), NULL);
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
						CString lDateTime = tt.FormatLocalTm(m_format);
						strcpy(datebuff, (LPCSTR)lDateTime);
					}
					else {
						CString lDateTime = tt.FormatGmtTm(m_format);
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
					CString from;

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
			else if (iSubItem == 4) {
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
			if (TextUtilsEx::Str2Wide(FieldText, charsetId, strW))
			{
				SIZE size;
				int wlen = strW.GetLength();
				int n = wlen;

				txtW.Empty();
				BOOL ret = GetTextExtentPoint32W(hDC, strW, n, &size);
				if ((size.cx + 0) > w) {
					n = (int)((double)(wlen * w) / size.cx);
					if (n > 3)
						n -= 3;
					else
						n = 1;
					txtW = strW.Mid(0, n);
					txtW.Append(L"...");
				}
				else
					txtW = strW;

				::ExtTextOutW(hDC, xpos, ypos, ETO_CLIPPED, rect, (LPCWSTR)txtW, txtW.GetLength(), NULL);
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

	FileUtils::RemoveDirW(FileUtils::GetmboxviewTempPathW());
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
	TRACE(_T("threadId=%ld threadPriority=%ld\n"), myThreadId, myThreadPri);

	MboxMail::Parse(path);
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
	TRACE(_T("threadId=%ld threadPriority=%ld\n"), myThreadId, myThreadPri);

	BOOL mainThreadContext = FALSE;
	int maxSearchDuration = -1;
	args->retpos = args->lview->DoFastFind(args->searchstart, mainThreadContext, maxSearchDuration, args->findAll);
	args->exitted = TRUE;
	return true;
}

#define CACHE_VERSION	18

BOOL SaveMails(LPCSTR cache, BOOL mainThread, CString &errorText)
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
			//MboxMail::pCUPDUPData->SetProgress(_T("Creating index file and image files ..."), 0);
			MboxMail::pCUPDUPData->SetProgress(_T("Creating index file ..."), 0);
		}
	}

	double delta = (double)ni;
	if (delta <= 0) delta = 1;
	double step = delta / 10000;
	double curstep = 1;
	double newstep = 0;
	CString fileNum;
	int nFileNum = 0;

	if (!mainThread && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	sz.writeInt(CACHE_VERSION);			// version
	sz.writeInt64(MboxMail::s_fSize = FileUtils::FileSize(MboxMail::s_path));	// file size
	sz.writeInt(ni);					// number of elements
	_int64 lastoff = 0;
	BOOL terminated = FALSE;
	CString itemNumberStr;
	std::vector <MailBodyContent*>::iterator it;

	int i = 0;
	for (i = 0; i < ni; i++)
	{
		MboxMail *m = MboxMail::s_mails[i];
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
		//sz.writeInt(m->m_crc32);

		int count = m->m_ContentDetailsArray.size();
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

			newstep = ((double)(i - 0 + 1)) / step;
			if (newstep >= curstep)
			{
				curstep = newstep;
			}

			int nItem = (i + 1);
			if ((nItem % 10000) == 0) {
				itemNumberStr.Format(_T("Creating index file ... %d of %d"), nItem, ni);
				if (MboxMail::pCUPDUPData)  MboxMail::pCUPDUPData->SetProgress(itemNumberStr, (UINT_PTR)(curstep));
			}
		}

		if (lastoff < m->m_startOff)
			lastoff = m->m_startOff;
	}
	TRACE("lastoff=%lld\n", lastoff);

	sz.close();

	if (terminated)
	{
		errorText = "Partial Index file will be deleted.";
		BOOL ret = DeleteFile(cache);
		if (ret == FALSE)
		{
			errorText.Format("Failed to delete index file %s. Please remove manually.");
			int deb = 1;
		}
	}

	return TRUE;
}


int NListView::LoadMails(LPCSTR cache, MailArray *mails)
{
	SerializerHelper sz(cache);
	//if (!sz.open(FALSE, 64))
	if (!sz.open(FALSE))
	{
		return -1;
	}
	int version;
	_int64 fSize;
	int ni = 0;

	if (!sz.readInt(&version))
	{
		sz.close();
		return -1;
	}

	if (version != CACHE_VERSION) 
	{
		CString cacheName;
		FileUtils::CPathStripPath(cache, cacheName);

		CString txt = _T("Index file\n\"") + cacheName;
		CString strVersion;
		strVersion.Format(_T("%d"), version);
		txt += _T("\".\nhas incompatible version\"") + strVersion + "\". Expected version \"";
		strVersion.Format(_T("%d"), CACHE_VERSION);
		txt += strVersion + "\".\nWill remove\n\"" + cacheName + "\"\nand recreate from the mail archive file.";
		HWND h = GetSafeHwnd();
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);

		sz.close();
		return -1;
	}
	if (!sz.readInt64(&fSize))
	{
		sz.close();
		return -1;
	}

	// TODO: Verify file length of both mbox and mboxview files
	if (fSize != FileUtils::FileSize(MboxMail::s_path))
	{
		sz.close();
		return -1;
	}

	_int64 lastoff = 0;
	if (sz.readInt(&ni)) 
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

		for (int i = 0; i < ni; i++)
		{
			MboxMail *m = new MboxMail();
			if (!sz.readInt64(&m->m_startOff))
				break;
			if (!sz.readInt(&m->m_length))
				break;
			if (!sz.readInt(&m->m_index))
				break;
			if (!sz.readInt(&m->m_hasAttachments))
				break;
			if (!sz.readInt(&m->m_headLength))
				break;
			if (!sz.readInt(&m->m_recv))
				break;
			if (!sz.readString(m->m_from))
				break;
			if (!sz.readString(m->m_from_charset))
				break;
			if (!sz.readUInt(&m->m_from_charsetId))
				break;
			if (!sz.readString(m->m_to))
				break;
			if (!sz.readString(m->m_to_charset))
				break;
			if (!sz.readUInt(&m->m_to_charsetId))
				break;
			if (!sz.readString(m->m_subj))
				break;
			if (!sz.readString(m->m_subj_charset))
				break;
			if (!sz.readUInt(&m->m_subj_charsetId))
				break;

			if (!sz.readString(m->m_cc))
				break;
			if (!sz.readString(m->m_cc_charset))
				break;
			if (!sz.readUInt(&m->m_cc_charsetId))
				break;

			if (!sz.readString(m->m_bcc))
				break;
			if (!sz.readString(m->m_bcc_charset))
				break;
			if (!sz.readUInt(&m->m_bcc_charsetId))
				break;

			if (!sz.readInt64(&m->m_timeDate))
				break;
			if (!sz.readInt(&m->m_groupId))
				break;
			if (!sz.readInt(&m->m_groupColor))
				break;
			if (!sz.readString(m->m_messageId))
				break;
			//if (!sz.readUInt(&m->m_crc32))
				//break;

			int count = 0;
			if (!sz.readInt(&count))
				break;

			MailBodyContent *body;
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
	TRACE("lastoff=%lld\n", lastoff);
	sz.close();
	if (mails == 0)
		MboxMail::s_fSize = MboxMail::s_oSize = fSize;

	return ni;
}

int Cache2Text(LPCSTR cache, CString format)
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

	ni = 0;
	_int64 lastoff = 0;
	if (sz.readInt(&ni)) {
		for (int i = 0; i < ni; i++) {
			if (!sz.readInt64(&m->m_startOff))
				break;
			if (!sz.readInt(&m->m_length))
				break;
			if (!sz.readInt(&m->m_index))
				break;
			if (!sz.readInt(&m->m_hasAttachments))
				break;
			if (!sz.readInt(&m->m_headLength))
				break;
			if (!sz.readInt(&m->m_recv))
				break;
			if (!sz.readString(m->m_from))
				break;
			if (!sz.readString(m->m_from_charset))
				break;
			if (!sz.readUInt(&m->m_from_charsetId))
				break;
			if (!sz.readString(m->m_to))
				break;
			if (!sz.readString(m->m_to_charset))
				break;
			if (!sz.readUInt(&m->m_to_charsetId))
				break;
			if (!sz.readString(m->m_subj))
				break;
			if (!sz.readString(m->m_subj_charset))
				break;
			if (!sz.readUInt(&m->m_subj_charsetId))
				break;

			if (!sz.readString(m->m_cc))
				break;
			if (!sz.readString(m->m_cc_charset))
				break;
			if (!sz.readUInt(&m->m_cc_charsetId))
				break;

			if (!sz.readString(m->m_bcc))
				break;
			if (!sz.readString(m->m_bcc_charset))
				break;
			if (!sz.readUInt(&m->m_bcc_charsetId))
				break;

			if (!sz.readInt64(&m->m_timeDate))
				break;
			if (!sz.readInt(&m->m_groupId))
				break;
			if (!sz.readInt(&m->m_groupColor))
				break;
			if (!sz.readString(m->m_messageId))
				break;
			//if (!sz.readUInt(&m->m_crc32))
				//break;

			if (lastoff < m->m_startOff)
				lastoff = m->m_startOff;

			datebuff[0] = 0;
			if (m->m_timeDate >= 0) {
				MyCTime tt(m->m_timeDate);
				CString lDateTime = tt.FormatGmtTm(format);
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
	TRACE("lastoff=%lld\n", lastoff);
	sz.close();
	CloseHandle(hFile);

	return ni;
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
	if( m_path.IsEmpty() )
		return;
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
		return;

	CString cache;
	BOOL ret = MboxMail::GetMboxviewFilePath(m_path, cache);

	int ni = 0;
#ifdef _DEBUG
	DWORD tc = GetTickCount(); // montano@
#endif

#if 0
	// for testing of changes to mboxview file
	{
		int ni_old = 0;
		int ni_new = 0;
		CString cache_new = cache + ".new";
		MailArray s_mails_new;
		if (FileUtils::fileExists(cache_new)) {
			MboxMail::s_path = m_path;
			ni_new = LoadMails(cache_new, &s_mails_new);
			if (ni_new < 0) {
				ni_new = 0;
			}
		}
		CString cache_old = cache + ".old";
		MailArray s_mails_old;
		if (FileUtils::fileExists(cache_old)) {
			MboxMail::s_path = m_path;
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
					TRACE("indx=%d off=%lld newlen=%d oldlen=%d deelta=%d\n", i, m_new->m_startOff, m_new->m_length, m_old->m_length, delta);
					int deb = 1;
				}
			}
		}
	}
#endif

	if( FileUtils::fileExists(cache) ) 
	{
		MboxMail::s_path = m_path;
		// it populates s_mails from mail index/mboxview file
		ni = LoadMails(cache);
		if( ni < 0 ) {
			ni = 0;
			DeleteFile(cache);
		} else
			m_list.SetItemCount(ni);

		int deb = 1;
	}
	if( !FileUtils::fileExists(cache) )
	{
		PARSE_ARGS args;
		args.path = m_path;
		args.exitted = FALSE;

		//  ALongRightProcessProc will set MboxMail::s_path = m_path;
		MboxMail::runningWorkerThreadType = 1;
		CUPDialog	Dlg(AfxGetMainWnd()->GetSafeHwnd(), ALongRightProcessProc, (LPVOID)(PARSE_ARGS*)&args);
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
			//ASSERT(cancelledbyUser == TRUE);
			int loopCnt = 20;
			DWORD tc_start = GetTickCount();
			while ((loopCnt-- > 0) && (args.exitted == FALSE))
			{
				Sleep(25);
			}
			DWORD tc_end = GetTickCount();
			DWORD delta = tc_end - tc_start;
			TRACE("(FillCtrl)Waited %ld milliseconds for thread to exist.\n", delta);
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
		// finally unique index is assigned to each mail according to its position in the array
		// TODO: fix the function name ??
		ret = MboxMail::sortConversations();
		//MboxMail::assignColor2ConvesationGroups();

#if 0
		// creates to index/mboxview file from s_mails
		if( SaveMails(cache) ) {
			ni = MboxMail::s_mails.GetSize();
			m_list.SetItemCount(ni);
			//Cache2Text(cache, m_format);
		}

#else
		WRITE_INDEX_FILE_ARGS wargs;
		wargs.cache = cache;
		wargs.exitted = FALSE;
		wargs.ret = TRUE;
		CUPDialog	WDlg(AfxGetMainWnd()->GetSafeHwnd(), ALongRightProcessProcWriteIndexFile, (LPVOID)(PARSE_ARGS*)&wargs);
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
			//ASSERT(cancelledbyUser == TRUE);
			int loopCnt = 20;
			DWORD tc_start = GetTickCount();
			while ((loopCnt-- > 0) && (args.exitted == FALSE))
			{
				Sleep(25);
			}
			DWORD tc_end = GetTickCount();
			DWORD delta = tc_end - tc_start;
			TRACE("(FillCtrl)Waited %ld milliseconds for thread to exist.\n", delta);
		}

		if (!wargs.errorText.IsEmpty())
		{

			int answer = MessageBox(wargs.errorText, _T("Warning"), MB_APPLMODAL | MB_ICONWARNING | MB_OK);
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

		if (mailCnt > 1000)
		{
			CString txt = _T("Inline image files can be pre-created now or created on the fly later. "
				"Image cache improves performance when printing large number of mails to PDF or viewing in Browser.\n"
				"\nCreate cache of inline images (may take several minutes depending on the mail count and content) ?");
			int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
			if (answer == IDYES)
			{
				int firstMail = 0;
				int lastMail = MboxMail::s_mails.GetCount() - 1;

				CString targetPrintSubFolderName = "";
				int retval = CreateInlineImageCache_Thread(firstMail, lastMail, targetPrintSubFolderName);
			}
		}
		else if (mailCnt > 0)
		{
			int firstMail = 0;
			int lastMail = MboxMail::s_mails.GetCount() - 1;

			CString targetPrintSubFolderName = "";
			int retval = CreateInlineImageCache_Thread(firstMail, lastMail, targetPrintSubFolderName);
		}
	}

	MboxMail::s_mails_ref.SetSizeKeepData(MboxMail::s_mails.GetSize());
	MboxMail::s_mails_ref.Copy(MboxMail::s_mails);
	MboxMail::b_mails_sorted = true;
	MboxMail::b_mails_which_sorted = 1;
	MboxMail::SortByDate();

	MarkColumns();
	m_bEditFindFirst = TRUE;  // must call OnEditFind() first and not OnEditFindAgain()

#ifdef _DEBUG
	tc = (GetTickCount() - tc);
	TRACE("FillCtrl Took %d:%d %d\n", (tc/1000)/60, (tc/1000)%60, tc);
#endif
	m_list.EnsureVisible(ni, FALSE);
	m_list.SetRedraw(TRUE);
	BOOL retval;
	if (m_which == 0)
		int deb = 1;
	// m_which == 0 is legal, no exceptiomn will be thrown
	retval = pTreeView->m_tree.SelectItem(m_which);
	CString mboxFileName = pTreeView->m_tree.GetItemText(m_which);

	_int64 fSize = FileUtils::FileSize(MboxMail::s_path);
	CString folderPath;
	CString fileName;
	FileUtils::GetFolderPathAndFileName(m_path, folderPath, fileName);

	// This doesn't have any sense. What I am trying to do here
	if (fileName.Compare(mboxFileName) != 0)
		_ASSERT(fileName.Compare(mboxFileName) != 0);

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
		return -1;
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
	if (FileUtils::fileExists(cache))
	{
		MboxMail::s_path = m_path;
		// it populates s_mails from mail index/mboxview file
		ni = LoadMails(cache);
		if (ni < 0) {
			ni = 0;
			DeleteFile(cache);
		}
		else
			m_list.SetItemCount(ni);

		int deb = 1;
	}
	if (!FileUtils::fileExists(cache))
	{
		return -1;
	}

	MboxMail::s_mails_ref.SetSizeKeepData(MboxMail::s_mails.GetSize());
	MboxMail::s_mails_ref.Copy(MboxMail::s_mails);
	MboxMail::b_mails_sorted = true;
	MboxMail::b_mails_which_sorted = 1;
	MboxMail::SortByDate();

	MarkColumns();
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
		_ASSERT(fileName.Compare(mboxFileName) != 0);

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

	return strcmp(strItem1, strItem2);
}

void NListView::SelectItem(int iItem, BOOL ignoreViewMessageHeader)
{
	CMainFrame *pFrame = 0;

	//  EX_TEST 
#if 0
	*(char*)pFrame = 'a';
#endif

	// Sanity check
	if (iItem < 0 || iItem >= MboxMail::s_mails.GetSize()) {
		return;
		//ClearDescView();
		return;
	}

	pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (!pFrame)
		return;
	NMsgView *pMsgView = pFrame->GetMsgView();
	if (!pMsgView)
		return;

	HWND h = GetSafeHwnd();
	MboxMail::ShowHint(HintConfig::LabelsHint, h);

	m_lastSel = iItem;

	if (pFrame->m_bViewMessageHeaders == TRUE)
	{
		if (ignoreViewMessageHeader == FALSE)
		{
			pMsgView->SetMsgHeader(iItem, m_gmtTime, m_format);
			pMsgView->ShowMailHeader(iItem);
			return;
		}
		else
			pMsgView->DisableMailHeader();

	}
	else
		pMsgView->DisableMailHeader();

	pMsgView->m_bAttach = FALSE;

#if 1
	// Already set for ID_INDICATOR_MAIL
	if (pFrame) {
		CString sText = _T("Mail Retrieval In Progress ...");
		int paneId = 1;
		pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
	}
#endif

	// Erase any files previously saved
	FileUtils::RemoveDirW(FileUtils::GetmboxviewTempPathW());

	pMsgView->m_attachments.Reset();

	// Get cached mail
	MboxMail *m = MboxMail::s_mails[iItem];

	// Set header data
	pMsgView->SetMsgHeader(iItem, m_gmtTime, m_format);

	// Get raw mail body
	CString bdy;   m->GetBody(bdy);

	// Save raw message
	if (m_bExportEml)
	{
		SaveAsEmlFile(bdy);
	}
	CString bdycharset = "UTF-8";

	char *data = (char*)(LPCSTR)bdy;
	int datalen = 0;

	// Decode MIME message
	CMimeMessage mail;
	const char *bodyData = bdy;
	int nLoadedSize = mail.Load(bodyData, bdy, bdy.GetLength());

	//DumpItemDetails(iItem, m, mail);
	bool hasInlineAttachments = false;

	// Iterate all the descendant body parts
	CMimeBody::CBodyList bodies;
	int nCount = mail.GetBodyPartList(bodies);

	int mailPosition = iItem;
	MailBodyInfoArray cidArray;
	MyCArray<bool> fileImgAlreadyCreatedArray;

	// CreateInlineImageFiles_SelectedItem will create files for embeded images and mark each body as embeded or not in cidArray
	// 
	int ret = NListView::CreateInlineImageFiles_SelectedItem(bodies, pMsgView, mailPosition, cidArray, fileImgAlreadyCreatedArray);
	if (cidArray.GetCount())
		hasInlineAttachments = true;

	int bodyIndex = 0;
	CMimeBody::CBodyList::const_iterator it;
	bdy = "";
	bool alreadyFoundPlain = false; 
	bool alreadyFoundHTML = false;  // prefer HTML Text over Plain Text
	int partsCnt = 0;


	// Please please find time and stop making small changes and rewrite.
	CString ext = "";
	if (bodies.begin() == bodies.end())
	{  // should never be true
		char *contentData = (char*)mail.GetContent();
		int contentDataLength = mail.GetContentLength();
		LPCSTR buff = bdy.GetBufferSetLength(contentDataLength);
		memcpy((char*)buff, contentData, contentDataLength);
		bdy.ReleaseBuffer();
		ext = "txt";
	}
	else  for (it = bodies.begin(); it != bodies.end(); it++, bodyIndex++)
	{
		CMimeBody* pBP = *it;
		CString curExt = "txt";
		bool isPlainText = true;
		bool handleAsAttachment = false;

		bool imgFileDone = false;
		int count = fileImgAlreadyCreatedArray.GetCount();
		if (bodyIndex < count)
		{
			imgFileDone = fileImgAlreadyCreatedArray.GetAt(bodyIndex);
		}

		if (imgFileDone) {
			// embedded image, already processed in CreateInlineImageFiles_SelectedItem above so continue
			continue;
		}

		// process bodies that are not embedded images

		// Iterate all the header fields of this body part:
		CMimeHeader::CFieldList& fds = pBP->Fields();
		CMimeHeader::CFieldList::const_iterator itfd;
		CString charset;

		for (itfd = fds.begin(); itfd != fds.end(); itfd++)
		{
			const CMimeField& fd = *itfd;
			const char *fname = fd.GetName();
			const char *fval = fd.GetValue();
			// Check content type to get mail extension TODO: encapsulate in function
			if (_stricmp(fname, "Content-Type") == 0 && _strnicmp(fval, "text/", 5) == 0) 
			{
				const char *p = fd.GetValue() + 5;
				if (_strnicmp(p, "plain", 5) == 0)
					curExt = "txt";
				else  if (_strnicmp(p, "htm", 3) == 0) {
					curExt = "htm";
					isPlainText = false;
				}
				else
				{
					TRACE(_T("Content-Type=text %s index=%d"), fname, iItem);
					handleAsAttachment = true;
					curExt = "txt";
				}

				int pc = CString(fval).Find("charset=");
				if (pc != -1) {
					int charsetLength = CString(fval + pc + 8).FindOneOf(";\n\r");
					if (charsetLength < 0)
						charsetLength = strlen(fval + pc + 8);
					strncpy(charset.GetBufferSetLength(charsetLength), fval + pc + 8, charsetLength);
					charset.Trim("\"\\");
				}
				break;
			}
		}

		// TODO: inefficient when we have both text and html parts. 
		// We read and initialize bdy as text and later override with html.

		CString contentType;
		MboxCMimeHelper::GetContentType(pBP, contentType);

		CString disposition;
		MboxCMimeHelper::GetContentDisposition(pBP, disposition);

		CString contentId;
		MboxCMimeHelper::GetContentID(pBP, contentId);
		contentId.Trim();
		contentId.Trim("<>");

		const char* cDisposition = pBP->GetDisposition();
		if (cDisposition)
			int deb = 1;

		CString contentTypeExtension;
		CString contentTypeMain;
		int pos = contentType.ReverseFind('/');
		if (pos > 0)
		{
			contentTypeExtension = contentType.Mid(pos + 1);
			contentTypeMain = contentType.Left(pos);
		}

		BOOL  isOctetStream = FALSE;
		if (contentTypeExtension.CompareNoCase("octet-stream") == 0)
			isOctetStream = TRUE;

		CString contentIdExtension;;
		CString contentIdMain = contentType.Left(pos);
		pos = contentType.ReverseFind('/');
		if (pos > 0)
		{
			contentIdExtension = contentId.Mid(pos + 1);
			contentIdMain = contentId.Left(pos);
		}

		bool isText = pBP->IsText();
		bool isMessage = pBP->IsMessage();
		bool isAttachment = MboxCMimeHelper::IsAttachment(pBP);
		bool isInlineAttachment = MboxCMimeHelper::IsInlineAttachment(pBP);
		if (isInlineAttachment)
			int deb = 1;
		bool IsBodyEmpty = bdy.IsEmpty();

		if (contentTypeMain.CompareNoCase("message") == 0)
		{
			TRACE(_T("Content-Type=message %s index=%d"), contentType, iItem);
			int deb = 1;
		}

		if (isText && isPlainText && !IsBodyEmpty)
			int deb = 1;
		if (isMessage)
			int deb = 1;

		if (isMessage && !isAttachment)
			int deb = 1;


		if (isText && !(isAttachment || handleAsAttachment))
		{
			// concatanate text blocks of the same type, prefer Html
			// Apple mail may have multiple Plain and Html text blocks;
			// No checks are done to verify Html blocks are formatted correctly
			char *contentData = (char*)pBP->GetContent();
			int contentDataLength = pBP->GetContentLength();

			if (isPlainText)
			{
				if (!alreadyFoundHTML)
				{
					if (alreadyFoundPlain == false)
					{
						bdy.Empty();
						IsBodyEmpty = TRUE;
						partsCnt = 0;
						bdycharset.Empty();
					}

					alreadyFoundPlain = true;

					if (IsBodyEmpty)
					{
						LPCSTR buff = bdy.GetBufferSetLength(contentDataLength);
						memcpy((char*)buff, contentData, contentDataLength);
						bdy.ReleaseBuffer();
					}
					else
						bdy.Append(contentData, contentDataLength);

					partsCnt++;
					if (partsCnt > 1)
						int deb = 1;

					ext = curExt;
					if (!charset.IsEmpty())
						bdycharset = charset;
				}
			}
			else
			{
				if (alreadyFoundHTML == FALSE) {
					bdy.Empty();
					IsBodyEmpty = TRUE;
					partsCnt = 0;
					bdycharset.Empty();
				}

				alreadyFoundHTML = true;

				if (IsBodyEmpty)
				{
					LPCSTR buff = bdy.GetBufferSetLength(contentDataLength);
					memcpy((char*)buff, contentData, contentDataLength);
					bdy.ReleaseBuffer();
				}
				else
					bdy.Append(contentData, contentDataLength);

				partsCnt++;
				if (partsCnt > 1)
					int deb = 1;

				ext = curExt;
				if (!charset.IsEmpty())
					bdycharset = charset;
			}
			TRACE("ext=%s charset=%s\n", (LPCSTR)ext, (LPCSTR)charset);
		}
		else
		{
			// Attachment file names mess
			// https://blog.nodemailer.com/2017/01/27/the-mess-that-is-attachment-filenames/
			// https://tools.ietf.org/html/rfc2047
			// https://tools.ietf.org/html/rfc2231
			//

			string nameCharset = pBP->GetFilenameCharset();
			string strName = pBP->GetFilename();
			if (strName.empty())
			{
				strName = pBP->GetName();
				nameCharset = pBP->GetNameCharset();
			}

			DWORD error;
			CString documentExtension;
			HtmlUtils::CommonMimeType2DocumentTypes(contentType, documentExtension);

			if (!documentExtension.IsEmpty())
			{
				contentTypeExtension = documentExtension.Mid(1);
			}

			UINT nameCharsetId = TextUtilsEx::Str2PageCode(nameCharset.c_str());
			if (nameCharsetId == 0)
				nameCharsetId = CP_UTF8;

			CString cStrName = strName.c_str();

			CStringW contentTypeExtensionW;
			if (TextUtilsEx::Ansi2Wide(contentTypeExtension, contentTypeExtensionW, error))
			{
				int deb = 1;
			}

			CStringW contentTypeMainW;
			if (TextUtilsEx::Ansi2Wide(contentTypeMain, contentTypeMainW, error))
			{
				int deb = 1;
			}
			// 
			CStringW cStrNameW;
			if (TextUtilsEx::Str2Wide(cStrName, nameCharsetId, cStrNameW))
			{
				if (!cStrNameW.IsEmpty())
				{
					pos = cStrNameW.ReverseFind(L'.');
					if (pos < 0)  // no extension found
					{
						if (isOctetStream)
						{
							const unsigned char* data = pBP->GetContent();
							int dataLength = pBP->GetContentLength();

							CStringW extensionW;
							CString extension;
							BOOL ret = loadImage((BYTE*)data, (size_t)dataLength, extensionW, extension);

							if (ret)
								cStrNameW += extensionW; 
						}
						else
							cStrNameW += L"." + contentTypeExtensionW;
					}  // else extension found, all done
				}
				else // cStrNameW is empty
				{
					if (isOctetStream)
					{
						const unsigned char* data = pBP->GetContent();
						int dataLength = pBP->GetContentLength();

						CStringW extensionW;
						CString extension;
						BOOL ret = loadImage((BYTE*)data, (size_t)dataLength, extensionW, extension);
						if (ret)
							cStrNameW = contentTypeMainW + extensionW;
					}
					else
						cStrNameW = contentTypeMainW + L"." + contentTypeExtensionW;

					int deb = 1;
				}
				int deb = 1;
			}
			//
			if (cStrNameW.IsEmpty())
			{
				if (!cStrName.IsEmpty())
				{
					pos = cStrName.ReverseFind('.');
					if (pos < 0)
					{
						if (isOctetStream)
						{
							const unsigned char* data = pBP->GetContent();
							int dataLength = pBP->GetContentLength();

							CStringW extensionW;
							CString extension;
							BOOL ret = loadImage((BYTE*)data, (size_t)dataLength, extensionW, extension);
							if (ret)
							{
								cStrName = contentTypeMain + extension;
							}
						}
						else
							cStrName += "." + contentTypeExtension;
					}  // else extension found, all done
				}
				else // cStrName is empty
				{
					if (isOctetStream)
					{
						const unsigned char* data = pBP->GetContent();
						int dataLength = pBP->GetContentLength();

						CStringW extensionW;
						CString extension;
						BOOL ret = loadImage((BYTE*)data, (size_t)dataLength, extensionW, extension);
						if (ret)
						{
							cStrName = contentTypeMain + extension;
						}
					}
					else
						cStrName = contentTypeMain + "." + contentTypeExtension;

					int deb = 1;
				}

				if (TextUtilsEx::Str2Wide(cStrName, nameCharsetId, cStrNameW))
				{
					int deb = 1;
				}
				// else TODO: errror, indicates bigger problem
			}
			if (cStrNameW.IsEmpty())
				cStrNameW = contentTypeMainW + L"." + contentTypeExtensionW;

			pMsgView->m_attachments.InsertItemW(cStrNameW, bodyIndex, pBP);
			pMsgView->m_bAttach = TRUE;
		}
	}

	pMsgView->m_attachments.Complete();

	if (bdy.IsEmpty())
		ext = "txt";

	// Save mail
	if (ext.Compare("txt") == 0)
	{
		// Wrap as HTML to display text with different charsets correctly
		ext = "htm";

		SimpleString *outbuf = MboxMail::m_outbuf;
		outbuf->ClearAndResize(bdy.GetLength() + 1000);

		CString hdr = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body><br>";
		DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);
		if (m_bApplyColorStyle && (color != COLOR_WHITE))
		{
			CString colorStr;
			int retC2S = Color2Str(color, colorStr);
			hdr = "<html><head><style>body{background-color: #";
			hdr.Append(colorStr);
			hdr.Append(";}</style><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body><br>");
		}
		outbuf->Append((LPCSTR)hdr, hdr.GetLength());

		char *inData = (char*)(LPCSTR)bdy;
		int inDataLen = bdy.GetLength();
		SimpleString *tmpbuf = MboxMail::get_tmpbuf();
		TextUtilsEx::EncodeAsHtml(inData, inDataLen, tmpbuf);
		inData = tmpbuf->Data();
		inDataLen = tmpbuf->Count();

		outbuf->Append(inData, inDataLen);

		hdr = "</body></html>";
		outbuf->Append((LPCSTR)hdr, hdr.GetLength());

		data = outbuf->Data();
		datalen = outbuf->Count();

		MboxMail::rel_tmpbuf();
		int deb = 1;
	}
	else if (ext.Compare("htm") == 0)
	{
		BOOL bodyTagMissing = FALSE;

		const char *pBdy = bdy;
		int bodyLength = bdy.GetLength();

		if (TextUtilsEx::findNoCase(pBdy, bodyLength, "<body", 5) < 0)
			bodyTagMissing = TRUE;

		SimpleString *tmpbuf = MboxMail::get_tmpbuf();

		tmpbuf->ClearAndResize(bodyLength+256);
		int mailPosition = 0;
		BOOL m_ReplaceAllWhiteBackgrounTags = TRUE;
		if (m_bApplyColorStyle)
		{
			int retSet = NListView::SetBackgroundColor((char*)pBdy, bodyLength, tmpbuf, m_ReplaceAllWhiteBackgrounTags);
		}
#if 0
		if (tmpbuf->Count())
		{
			CStringW path = FileUtils::GetmboxviewTempPathW();
			CStringW filePath = path + L"\\" + L"Custom.html";
			FileUtils::Write2File(filePath, (unsigned char*)outb.Data(), outb.Count());
		}
#endif
		if (tmpbuf->Count())
		{
			pBdy = tmpbuf->Data();
			bodyLength = tmpbuf->Count();
		}
		else
		{
			pBdy = bdy;
			bodyLength = bdy.GetLength();
		}

		if (bodyTagMissing)
		{
			// File has htm extension but body tag is missing. 
			// TODO: Below assume html TAG is present.  Do we need to check for html TAG ?
			SimpleString *outbuf = MboxMail::m_outbuf;
			outbuf->ClearAndResize(bodyLength + 1024);

			CString hdr = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body><br>";
			DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);
			if (m_bApplyColorStyle && (color != COLOR_WHITE))
			{
				CString colorStr;
				int retC2S = Color2Str(color, colorStr);
				hdr = "<html><head><style>body{background-color: #";
				hdr.Append(colorStr);
				       hdr.Append(";}</style><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body><br>");
				//CString hdr = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body><br>";
			}
			outbuf->Append((LPCSTR)hdr, hdr.GetLength());

			char *inData = (char*)pBdy;
			int inDataLen = bodyLength;

			if (hasInlineAttachments)
			{
				int mailPosition = 0; // not used anyway here
				bool useMailPosition = false;
				// append to outbuf
				UpdateInlineSrcImgPath_SelectedItem(inData, inDataLen, outbuf, mailPosition, useMailPosition, cidArray);
			}
			else
				outbuf->Append(inData, inDataLen);

			hdr = "</body></html>";
			outbuf->Append((LPCSTR)hdr, hdr.GetLength());

			data = outbuf->Data();
			datalen = outbuf->Count();
			int deb = 1;
		}
		else // body TAG found
		{
			// Assume HTML tag is present :)

			// TODO: need to optimize use of buffers
			BOOL charsetMissing = FALSE;

			if (TextUtilsEx::findNoCase(pBdy, bodyLength, "charset=", 5) < 0)
				charsetMissing = TRUE;

			// For now always append missing charset
#if 0
			if (charsetMissing)
			{
				// assume that dflt for browsers is UTF8 so no need to do anything ?
				if (bdycharset.CompareNoCase("utf-8") == 0)  // 
					charsetMissing = FALSE;
				int deb = 1;
			}
#endif

			if ((charsetMissing) || (hasInlineAttachments))
			{
				SimpleString *outbuf = MboxMail::m_outbuf;
				outbuf->ClearAndResize(bodyLength + 1000);

				if (charsetMissing)
				{
					CString hdr = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head></html><br>";
					DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);
					if (m_bApplyColorStyle && (color != COLOR_WHITE))
					{
						CString colorStr;
						int retC2S = Color2Str(color, colorStr);
						hdr = "<html><head><style>body{background-color: #";
						hdr.Append(colorStr);
						       hdr.Append(";}</style><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head></html><br>");
						//CString hdr = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head></html><br>";
					}

					outbuf->Append((LPCSTR)hdr, hdr.GetLength());
					int deb = 1;
				}
				else
				{
					DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);
					if (m_bApplyColorStyle && (color != COLOR_WHITE))
					{
						CString colorStr;
						int retC2S = Color2Str(color, colorStr);
						CString hdr = "<html><head><style>body{background-color: ";
						hdr.Append(colorStr);
						       hdr.Append(";}</style><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head></html><br>");
						//CString hdr = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head></html><br>";
						outbuf->Append((LPCSTR)hdr, hdr.GetLength());
					}
				}

				char *inData = (char*)pBdy;
				int inDataLen = bodyLength;

				if (hasInlineAttachments)
				{
					int mailPosition = 0; // not used anyway here
					bool useMailPosition = false;
					UpdateInlineSrcImgPath_SelectedItem(inData, inDataLen, outbuf, mailPosition, useMailPosition, cidArray);
				}
				else
					outbuf->Append(inData, inDataLen);

				data = outbuf->Data();
				datalen = outbuf->Count();
				int deb = 1;
			}
			else
			{
				DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);
				if (m_bApplyColorStyle && (color != COLOR_WHITE))
				{
					SimpleString *outbuf = MboxMail::m_outbuf;
					outbuf->ClearAndResize(bodyLength + 1000);

					CString colorStr;
					int retC2S = Color2Str(color, colorStr);
					CString hdr = "<html><head><style>body{background-color: #";
					hdr.Append(colorStr);
					hdr.Append(";}</style></head></html><br>");
					//CString hdr = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head></html><br>";
					outbuf->Append((LPCSTR)hdr, hdr.GetLength());

					outbuf->Append(pBdy, bodyLength);

					data = outbuf->Data();
					datalen = outbuf->Count();
				}
				else
				{
					data = (char*)pBdy;
					datalen = bodyLength;
				}

				int deb = 1;
			}
			int deb = 1;
		}
		MboxMail::rel_tmpbuf();
	}
	else 
	{
		int deb = 1; 
	}

	if (ext.Compare("unk") == 0)
		int deb = 1;

	pMsgView->m_body_charsetId = TextUtilsEx::charset2Id(bdycharset);
	pMsgView->m_body_charset = bdycharset;

	// Get temporary file name with correct extension for IE to display
	m_curFile = FileUtils::CreateTempFileName(ext);
	CFile fp(m_curFile, CFile::modeWrite | CFile::modeCreate);
	fp.Write(data, datalen);
	fp.Close();

	if (datalen == 0)
		int deb = 1;

	// Display mail in internal IE
	pMsgView->m_browser.Navigate(m_curFile, NULL);
	// Update layou to show/hide attachments
	pMsgView->UpdateLayout();

	MboxMail::ShowHint(HintConfig::MailSelectionHint, GetSafeHwnd());

	if (cidArray.GetCount() || pMsgView->m_attachments.GetItemCount())
		PostMsgCmdParamAttachmentHint();

	MailBodyInfo* info;
	for (int i = 0; i < cidArray.GetCount(); i++)
	{
		info = cidArray[i];
		delete info;
	}

	_ASSERT(!MboxMail::m_tmpbufBusy);
	return;
}

// Dump header fields based on MboxMail *m and corresponding data from mime decoder
int NListView::DumpItemDetails(MboxMail *m)
{
	int iItem = 1;

	// Get raw mail body
	CString bdy;   m->GetBody(bdy);

	// Decode MIME message
	CMimeMessage mail;
	const char* bodyData = bdy;
	int nLoadedSize = mail.Load(bodyData, bdy, bdy.GetLength());

	int retval = DumpItemDetails(iItem, m, mail);
	return retval;
}

// Dump header fields based on MboxMail::s_mails[which] and corresponding data from mime decoder
int NListView::DumpItemDetails(int iItem)
{
	MboxMail *m = MboxMail::s_mails[iItem];

	// Get raw mail body
	CString bdy;   m->GetBody(bdy);

	// Decode MIME message
	CMimeMessage mail;
	const char* bodyData = bdy;
	int nLoadedSize = mail.Load(bodyData, bdy, bdy.GetLength());

	int retval = DumpItemDetails(iItem, m, mail);
	return retval;
}

// Dump header fields based on MboxMail *m and corresponding data from mime decoder
int NListView::DumpItemDetails(int iItem, MboxMail *m, CMimeMessage &mail)
{
	static BOOL restarted = TRUE;
	char buff[2048];
	char datebuff[256];
	DWORD offset = 0;
	DWORD nwritten = 0;
	DWORD count = 0;

	// Sanity check
	if (iItem < 0 || iItem >= MboxMail::s_mails.GetSize()) {
		return -1;
	}

	CString messageTextFile = "DumpItemDetails.txt";
	HANDLE hFile = CreateFile(messageTextFile, GENERIC_WRITE, 0, NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE) {
		DWORD err = GetLastError();
		TRACE(_T("(DumpItemDetails)INVALID_HANDLE_VALUE error= %ld\n"), err);
		return -1;
	}

	_int64 pos = 0;
	if (restarted == TRUE) {
		SetEndOfFile(hFile);
		restarted = FALSE;
	}
	else
	{
		LARGE_INTEGER li;

		BOOL retval = GetFileSizeEx(hFile, &li);
		long long fsize = li.QuadPart;
		pos = FileUtils::FileSeek(hFile, fsize, FILE_BEGIN);
		if (pos < 0) {
			DWORD err = GetLastError();
			CloseHandle(hFile);
			return -1;
		}
	}

	datebuff[0] = 0;
	if (m->m_timeDate >= 0) {
		MyCTime tt(m->m_timeDate);
		CString format = ""; // FIXME
		CString lDateTime = tt.FormatGmtTm(format);
		strcpy(datebuff, (LPCSTR)lDateTime);
	}

	count = sprintf_s(buff, "INDX=%d first=%lld len=%d last=%lld att=%d hlen=%d rcv=%d date=\"%s\" from=\"%s\" to=\"%s\" subj=\"%s\"\n\n",
		iItem, m->m_startOff, m->m_length, (m->m_startOff + m->m_length - 1), m->m_hasAttachments, m->m_headLength, m->m_recv,
		datebuff, (LPCSTR)m->m_from, (LPCSTR)m->m_to, (LPCSTR)m->m_subj);

	nwritten = 0;
	if (!WriteFile(hFile, buff, count, &nwritten, NULL)) {
		DWORD retval = GetLastError();
	}

	count = sprintf_s(buff, "INDX=%d date=\"%s\" from=\"%s\" from_charset=\"%s\" to=\"%s\" to_charset=\"%s\" cc=\"%s\" bcc=\"%s\" subj=\"%s\" to_charset=\"%s\"\n\n",
		iItem, FixIfNull(mail.GetDate()), FixIfNull(mail.GetFrom()), FixIfNull(mail.GetFieldCharset("From")), 
		FixIfNull(mail.GetTo()), FixIfNull(mail.GetFieldCharset("To")),
		FixIfNull(mail.GetCc()), FixIfNull(mail.GetBcc()), FixIfNull(mail.GetSubject()), 
		FixIfNull(mail.GetFieldCharset("Subject")));

	nwritten = 0;
	if (!WriteFile(hFile, buff, count, &nwritten, NULL)) {
		DWORD retval = GetLastError();
	}

	DumpCMimeMessage(mail, hFile);
	SetEndOfFile(hFile);
	CloseHandle(hFile);

	return 1;
}


int NListView::DumpCMimeMessage(CMimeMessage &mail, HANDLE hFile)
{
	char buff[32768];
	DWORD offset = 0;
	DWORD nwritten = 0;
	DWORD count = 0;


	CMimeMessage::CMimeBody &body = mail;
	CMimeBody* pBP = &body;
	int indx = 0;

	string strText;
	mail.GetText(strText);
	CString bdy = strText.c_str();

	count = sprintf_s(buff, "Message Main Body:\n\n");
	if (!WriteFile(hFile, buff, count, &nwritten, NULL)) {
		DWORD retval = GetLastError();
	}

	const unsigned char* content = pBP->GetContent();
	int contentLength = pBP->GetContentLength();

	char* cContentType = FixIfNull(pBP->GetContentType());
	char* cDescription = FixIfNull(pBP->GetDescription());
	char* cDisposition = FixIfNull(pBP->GetDisposition());
	char* cTransferEncoding = FixIfNull(pBP->GetTransferEncoding());

	bool isAttachment = MboxCMimeHelper::IsAttachment(pBP);

	count = sprintf_s(buff, "BodyIndx=%d IsText=%d IsMessage=%d HasBody=%d IsAttachement=%d IsMultiPart=%d ContentLength=%d "
		"Charset=%s Description=%s Disposition=%s TransferEncoding=%s SubType=%s MainType=%s "
		"Boundary=%s ContentType=%s MediaType=%d Name=%s\n",
		indx, pBP->IsText(), pBP->IsMessage(), !bdy.IsEmpty(), isAttachment, pBP->IsMultiPart(), contentLength,
		pBP->GetCharset().c_str(), cDescription, cDisposition, cTransferEncoding, pBP->GetSubType().c_str(), pBP->GetMainType().c_str(),
		pBP->GetBoundary().c_str(), cContentType, pBP->GetMediaType(), pBP->GetName().c_str());

	if (!WriteFile(hFile, buff, count, &nwritten, NULL)) {
		DWORD retval = GetLastError();
	}

	// Iterate all the header fields of mail body
	CMimeHeader::CFieldList& fds = body.Fields();
	CMimeHeader::CFieldList::const_iterator itfd;
	//CString charset;
	for (itfd = fds.begin(); itfd != fds.end(); itfd++)
	{
		const CMimeField& fd = *itfd;
		const char *fname = FixIfNull(fd.GetName());
		const char *fval = FixIfNull(fd.GetValue());
		const char *charset = FixIfNull(fd.GetCharset());
		count = sprintf_s(buff, "fname=%s fval=%s charset=%s\n", fname, fval, charset);

		if (!WriteFile(hFile, buff, count, &nwritten, NULL)) {
			DWORD retval = GetLastError();
		}
	}

	// Iterate all the descendant body parts
	CMimeBody::CBodyList bodies;

	int nCount = mail.GetBodyPartList(bodies);
	CMimeBody::CBodyList::const_iterator it;

	bdy = "";
	CString ext = "";
	if (bodies.begin() == bodies.end()) {
		//bdy = mail.GetContent();
		string strText;
		mail.GetText(strText);
		bdy = strText.c_str();
		ext = "txt";
	}
	else {
		count = sprintf_s(buff, "Message Body List:\n\n");
		if (!WriteFile(hFile, buff, count, &nwritten, NULL)) {
			DWORD retval = GetLastError();
		}
		int indx = 0;
		for (it = bodies.begin(); it != bodies.end(); it++)
		{
			CMimeBody* pBP = *it;
			CString curExt;

			const unsigned char* content = pBP->GetContent();
			int contentLength = pBP->GetContentLength();

			char* cContentType = FixIfNull(pBP->GetContentType());
			char* cDescription = FixIfNull(pBP->GetDescription());
			char* cDisposition = FixIfNull(pBP->GetDisposition());
			char* cTransferEncoding = FixIfNull(pBP->GetTransferEncoding());

			bool isAttachment = MboxCMimeHelper::IsAttachment(pBP);

			count = sprintf_s(buff, "BodyIndx=%d IsText=%d IsMessage=%d HasBody=%d IsAttachement=%d IsMultiPart=%d ContentLength=%d "
				"Charset=%s Description=%s Disposition=%s TransferEncoding=%s SubType=%s MainType=%s "
				"Boundary=%s ContentType=%s MediaType=%d Name=%s\n",
				indx, pBP->IsText(), pBP->IsMessage(), !bdy.IsEmpty(), isAttachment, pBP->IsMultiPart(), contentLength,
				pBP->GetCharset().c_str(), cDescription, cDisposition, cTransferEncoding, pBP->GetSubType().c_str(), pBP->GetMainType().c_str(),
				pBP->GetBoundary().c_str(), cContentType, pBP->GetMediaType(), pBP->GetName().c_str());

			if (!WriteFile(hFile, buff, count, &nwritten, NULL)) {
				DWORD retval = GetLastError();
			}

			indx++;

			// Iterate all the header fields of this body part:
			CMimeHeader::CFieldList& fds = pBP->Fields();
			CMimeHeader::CFieldList::const_iterator itfd;
			//CString charset;
			for (itfd = fds.begin(); itfd != fds.end(); itfd++)
			{
				const CMimeField& fd = *itfd;
				const char *fname = FixIfNull(fd.GetName());
				const char *fval = FixIfNull(fd.GetValue());
				const char *charset = FixIfNull(fd.GetCharset());
				count = sprintf_s(buff, "fname=%s fval=%s charset=%s\n", fname, fval, charset);

				if (!WriteFile(hFile, buff, count, &nwritten, NULL)) {
					DWORD retval = GetLastError();
				}
			}
			if (content && (pBP->GetMainType() == "text")) {
				if (!WriteFile(hFile, content, contentLength, &nwritten, NULL)) {
					DWORD retval = GetLastError();
				}
				char *CRLF = "\n\n";
				if (!WriteFile(hFile, CRLF, strlen(CRLF), &nwritten, NULL)) {
					DWORD retval = GetLastError();
				}
			}
		}
	}
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


	pMsgView->m_subj_charset.SetString(_T(""));
	pMsgView->m_from_charset.SetString(_T(""));
	pMsgView->m_date_charset.SetString(_T(""));
	pMsgView->m_to_charset.SetString(_T(""));
	pMsgView->m_cc_charset.SetString(_T(""));
	pMsgView->m_bcc_charset.SetString(_T(""));

	pMsgView->m_body_charset.SetString(_T(""));

	pMsgView->m_body_charsetId = 0;
	pMsgView->m_attachments.DeleteAllItems();
	FileUtils::RemoveDirW(FileUtils::GetmboxviewTempPathW());
	m_curFile.Empty();

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

	dlg.m_params.m_bWholeWord = m_bWholeWord;
	dlg.m_params.m_bCaseSensitive = m_bCaseSens;
	dlg.m_params.m_string = m_searchString;

	// The m_lastStartDate and m_lastEndDate are shared across Find and Advanced Find
	// Should eventually make them independed
	// MyCTime m_lastStartDate;
	// MyCTime m_lastEndDate;

	if (m_bNeedToFindMailMinMaxTime)
	{
		FindMinMaxTime(m_mboxMailStartDate, m_mboxMailEndDate);
		m_lastStartDate = m_mboxMailStartDate;
		m_lastEndDate = m_mboxMailEndDate;
		m_bNeedToFindMailMinMaxTime = FALSE;
	}

	if (MboxMail::nWhichMailList != IDC_ARCHIVE_LIST)
	{
		FindMinMaxTime(m_lastStartDate, m_lastEndDate);
		m_needToRestoreArchiveListDateTime = TRUE;
	}
	else if ((m_filterDates == FALSE) || m_needToRestoreArchiveListDateTime)
	{
		m_lastStartDate = m_mboxMailStartDate;
		m_lastEndDate = m_mboxMailEndDate;
		m_needToRestoreArchiveListDateTime = FALSE;
	}

	BOOL retDT = MyCTimeToOleTime(m_lastStartDate, dlg.m_params.m_startDate);
	retDT = MyCTimeToOleTime(m_lastEndDate, dlg.m_params.m_endDate);

	dlg.m_params.m_bFindNext = m_bFindNext;
	dlg.m_params.m_bFrom = m_bFrom;
	dlg.m_params.m_bTo = m_bTo;
	dlg.m_params.m_bSubject = m_bSubject;
	dlg.m_params.m_bContent = m_bContent;
	dlg.m_params.m_bCC = m_bCC;
	dlg.m_params.m_bBCC = m_bBCC;
	dlg.m_params.m_bAttachments = m_bAttachments;
	dlg.m_params.m_bAttachmentName = m_bAttachmentName;
	dlg.m_params.m_bHighlightAll = m_bHighlightAll;
	dlg.m_params.m_bFindAll = m_bFindAll;
	dlg.m_params.m_bFindAllMailsThatDontMatch = m_bFindAllMailsThatDontMatch;
	dlg.m_params.m_filterDates = m_filterDates;


	if (dlg.DoModal() == IDOK) 
	{
		m_filterDates = dlg.m_params.m_filterDates;
		m_searchString = dlg.m_params.m_string;
		m_bWholeWord = dlg.m_params.m_bWholeWord;
		m_bCaseSens = dlg.m_params.m_bCaseSensitive;

		retDT = NListView::OleTime2MyCTime(dlg.m_params.m_startDate, m_lastStartDate, FALSE);
		retDT = NListView::OleTime2MyCTime(dlg.m_params.m_endDate, m_lastEndDate, TRUE);

		m_bFindNext = dlg.m_params.m_bFindNext;
		m_bFrom = dlg.m_params.m_bFrom;
		m_bTo = dlg.m_params.m_bTo;
		m_bSubject = dlg.m_params.m_bSubject;
		m_bContent = dlg.m_params.m_bContent;
		m_bCC = dlg.m_params.m_bCC;
		m_bBCC = dlg.m_params.m_bBCC;
		m_bAttachments = dlg.m_params.m_bAttachments;
		m_bAttachmentName = dlg.m_params.m_bAttachmentName;
		m_bHighlightAll = dlg.m_params.m_bHighlightAll;
		m_bFindAllMailsThatDontMatch = dlg.m_params.m_bFindAllMailsThatDontMatch;
		m_bFindAll = dlg.m_params.m_bFindAll;

		if (m_bFindAll && !m_bFindNext)  // not implemented and it makes no difference if we set m_bFindNext to TRUE
			m_bFindNext = TRUE;

		m_lastFindPos = -1;

		// To optimize content/message and/or attachments search performance
		if ((m_bContent || m_bAttachments) && (MboxMail::b_mails_sorted == true))
		{
			int whichSorted = abs(MboxMail::b_mails_which_sorted);
			if ((whichSorted != 1) && (whichSorted != 99)) { // related mails should be in close proximity in the mail file
				MboxMail::SortByDate();
				MboxMail::b_mails_sorted = true;
				MboxMail::b_mails_which_sorted = 1;

				RedrawMails();
			}
		}

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
				if (m_bFindNext == TRUE)
					m_lastFindPos = 0;
				else
					m_lastFindPos = sz - 1;
			}
			else
			{
				if (m_bFindNext == TRUE) {
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
			if (m_bFindAll)
				maxSearchDuration = 0;
			if (maxSearchDuration > 0) {
				BOOL findAll = FALSE;
				w = DoFastFind(which, TRUE, maxSearchDuration, findAll);
			}
			FIND_ARGS args;
			/*IN*/ args.lview = this; args.searchstart = m_lastFindPos;
			/*OUT*/ args.exitted = FALSE; args.retpos = -1; args.findAll = m_bFindAll;
			if ((w == -2) || (maxSearchDuration == 0)) 
			{
				CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcFastSearch, (LPVOID)(FIND_ARGS*)&args);

				INT_PTR nResult = Dlg.DoModal();

				if (!nResult) { // should never be true ?
					m_bInFind = false;
					MboxMail::assert_unexpected();
					return;
				}

				int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
				int retResult = LOWORD(nResult);

				if (retResult != IDOK)
				{  // IDOK==1, IDCANCEL==2
					// We should be here when user selects Cancel button
					//ASSERT(cancelledbyUser == TRUE);
					int loopCnt = 20;
					DWORD tc_start = GetTickCount();
					while ((loopCnt-- > 0) && (args.exitted == FALSE))
					{
						Sleep(25);
					}
					DWORD tc_end = GetTickCount();
					DWORD delta = tc_end - tc_start;
					TRACE("(OnEditFind)Waited %ld milliseconds for thread to exist.\n", delta);
				}
				w = args.retpos;
				MboxMail::pCUPDUPData = NULL;
			}
			if (m_bFindAll) 
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
				return;
			}
			else if (w >= 0) {
				m_bHighlightAllSet = m_bHighlightAll;
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
}

int NListView::CheckMatch(int i, CString &searchString)
{
	int pos = -1;
	int w = -1;
	MboxMail *m = MboxMail::s_mails[i];

	pos = -1;
	if (m_bSubject) {
		if (searchString.GetLength() == 1 && searchString[0] == '*') {
			w = i;
			return w;
		}
		if (m_bWholeWord)
			pos = g_tu.BMHSearchW((unsigned char *)(LPCSTR)m->m_subj, m->m_subj.GetLength(), (unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_bCaseSens);
		else
			pos = g_tu.BMHSearch((unsigned char *)(LPCSTR)m->m_subj, m->m_subj.GetLength(), (unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_bCaseSens);
		if (pos >= 0) {
			w = i;
			return w;
		}
	}
	pos = -1;
	if (m_bFrom) {
		if (searchString.GetLength() == 1 && searchString[0] == '*') {
			w = i;
			return w;
		}
		if (m_bWholeWord)
			pos = g_tu.BMHSearchW((unsigned char *)(LPCSTR)m->m_from, m->m_from.GetLength(), (unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_bCaseSens);
		else
			pos = g_tu.BMHSearch((unsigned char *)(LPCSTR)m->m_from, m->m_from.GetLength(), (unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_bCaseSens);
		if (pos >= 0) {
			w = i;
			return w;
		}
	}
	if (m_bTo) {
		if (searchString.GetLength() == 1 && searchString[0] == '*') {
			w = i;
			return w;
		}
		if (m_bWholeWord)
			pos = g_tu.BMHSearchW((unsigned char *)(LPCSTR)m->m_to, m->m_to.GetLength(), (unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_bCaseSens);
		else
			pos = g_tu.BMHSearch((unsigned char *)(LPCSTR)m->m_to, m->m_to.GetLength(), (unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_bCaseSens);
		if (pos >= 0) {
			w = i;
			return w;
		}
	}
	if (m_bCC) {
		if (!m->m_cc.IsEmpty()) {
			if (searchString.GetLength() == 1 && searchString[0] == '*') {
				w = i;
				return w;
			}
		}
		if (m_bWholeWord)
			pos = g_tu.BMHSearchW((unsigned char *)(LPCSTR)m->m_cc, m->m_cc.GetLength(), (unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_bCaseSens);
		else
			pos = g_tu.BMHSearch((unsigned char *)(LPCSTR)m->m_cc, m->m_cc.GetLength(), (unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_bCaseSens);
		if (pos >= 0) {
			w = i;
			return w;
		}
	}
	if (m_bBCC) {
		if (!m->m_bcc.IsEmpty()) {
			if (searchString.GetLength() == 1 && searchString[0] == '*') {
				w = i;
				return w;
			}
		}
		if (m_bWholeWord)
			pos = g_tu.BMHSearchW((unsigned char *)(LPCSTR)m->m_bcc, m->m_bcc.GetLength(), (unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_bCaseSens);
		else
			pos = g_tu.BMHSearch((unsigned char *)(LPCSTR)m->m_bcc, m->m_bcc.GetLength(), (unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_bCaseSens);
		if (pos >= 0) {
			w = i;
			return w;
		}
	}

	if (m_bAttachmentName) {
		if (FindAttachmentName(m, searchString, m_bWholeWord, m_bCaseSens) == TRUE) {
			w = i;
			return w;
		}
	}

	if (m_bContent || m_bAttachments) {
		if (searchString.GetLength() == 1 && searchString[0] == '*') {
			w = i;
			return w;
		}
		if (FindInMailContent(i, m_bContent, m_bAttachments)) {
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

void NListView::DetermineKeywordsForProgresBar(CString *stringWithCase, CString &keyword1, CString &keyword2)
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
		key1.Append(stringWithCase[indx1]);

		int pr2 = CheckedList_FieldsPriority[1];
		int indx2 = FldPriority2Indx2PriorityMap[pr2];
		key2.Append(stringWithCase[indx2]);
	}
	else if (checkedFieldCnt == 1)
	{
		int pr1 = CheckedList_FieldsPriority[0];
		int indx1 = FldPriority2Indx2PriorityMap[pr1];
		key1.Append(stringWithCase[indx1]);

		key2.Empty();
	}

	int availableSpace = 28;
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
	int sz = MboxMail::s_mails.GetSize();
	time_t sd = m_lastStartDate.GetTime(), ed = m_lastEndDate.GetTime();
	CString searchString(m_searchString);
	if (m_advancedFind) 
	{
		for (int i = 0; i < FILTER_FIELDS_NUMB; i++)
		{
			m_stringWithCase[i] = m_advancedParams.m_string[i];
			if (m_advancedParams.m_bEditChecked[i])
			{
				if (m_advancedParams.m_bCaseSensitive[i] == FALSE) 
				{
					m_stringWithCase[i].MakeLower();
				}
			}
		}

		if (MboxMail::pCUPDUPData && !mainThreadContext)
		{
			CString keyword1;
			CString keyword2;
			DetermineKeywordsForProgresBar(m_stringWithCase, keyword1, keyword2);

			CString searchText;
			if (!keyword1.IsEmpty() && !keyword2.IsEmpty())
			{
				searchText.Append(_T("Searching for \"") + keyword1 + _T("\" && \"") + keyword2 + _T("\" ..."));
			}
			else if (!keyword1.IsEmpty())
			{
				searchText.Append(_T("Searching for \"") + keyword1 + _T("\" ..."));
			}
			else if (!keyword2.IsEmpty())
			{
				searchText.Append(_T("Searching for \"") + keyword2 + _T("\" ..."));
			}
			else
				searchText.Append(_T("Searching for \"\""));

			MboxMail::pCUPDUPData->SetProgress(searchText, 0);
		}
	}
	else 
	{
		if (m_bCaseSens == 0)
			searchString.MakeLower();

		if (MboxMail::pCUPDUPData && !mainThreadContext) {
			int availableSpace = 30;
			CString keyword = searchString.Left(availableSpace);
			CString searchText = _T("Searching for \"") + keyword + _T("\" ...");
			MboxMail::pCUPDUPData->SetProgress(searchText, 0);
		}
	}

	BOOL bFindAllMailsThatDontMatch = FALSE;
	BOOL bFilterDates = FALSE;
	if (m_advancedFind)
	{
		bFindAllMailsThatDontMatch = m_advancedParams.m_bFindAllMailsThatDontMatch;
		bFilterDates = m_advancedParams.m_filterDates;
	}
	else
	{
		bFindAllMailsThatDontMatch = m_bFindAllMailsThatDontMatch;
		bFilterDates = m_filterDates;
	}

	DWORD myThreadId = GetCurrentThreadId();
	DWORD tc_start = GetTickCount();

	if (m_bFindNext || findAll)
	{
		m_findAllCount = 0;
		if (findAll)
			MboxMail::s_mails_find.SetSizeKeepData(MboxMail::s_mails.GetCount());
		for (int i = which; i < sz; i++)
		{
			MboxMail *m = MboxMail::s_mails[i];
			bool process = false;
			if (bFilterDates) {
				if (m->m_timeDate >= sd && m->m_timeDate <= ed)
					process = true;
			}
			else
				process = true;
			if (process)
			{
				if (m_advancedFind) {
					if (CheckMatchAdvanced(i, m_advancedParams) >= 0) {
						if (findAll) {
							if (!bFindAllMailsThatDontMatch) {
								MboxMail::s_mails_find[m_findAllCount] = MboxMail::s_mails[i];
								m_findAllCount++;
							}
						}
					}
					else if (bFindAllMailsThatDontMatch)
					{
						if (findAll) {
							MboxMail::s_mails_find[m_findAllCount] = MboxMail::s_mails[i];
							m_findAllCount++;
						}
					}
				}
				else
				{
					if (CheckMatch(i, searchString) >= 0) {
						if (findAll) {
							if (!bFindAllMailsThatDontMatch)
							{
								MboxMail::s_mails_find[m_findAllCount] = MboxMail::s_mails[i];
								m_findAllCount++;
							}
						}
						else {
							w = i;
							break;
						}
					}
					else if (bFindAllMailsThatDontMatch)
					{
						// found mail matching negative search criteria
						if (findAll) {
							MboxMail::s_mails_find[m_findAllCount] = MboxMail::s_mails[i];
							m_findAllCount++;
						}
						else {
							w = i;
							break;
						}
					}
				}

				DWORD tc_curr = GetTickCount();
				DWORD tc_elapsed_milliseconds = (tc_curr - tc_start);
				DWORD tc_elapsed_seconds = (tc_curr - tc_start) / 1000;
				//TRACE(_T("index=%d pos=%lld threadId=0x%lx tc_elapsed_milliseconds=%ld tc_elapsed_seconds=%ld\n"),
					//i, pos, myThreadId, tc_elapsed_milliseconds, tc_elapsed_seconds);

				if (mainThreadContext) {
					if ((maxSearchDuration > 0) && (tc_elapsed_seconds > maxSearchDuration)) {
						m_lastFindPos = i + 1;
						w = -2;
						break;
					}
				}
				else {
					if (i % 10 == 0) {
						int step = (int)(100.0*((float)(i + 1) / sz));
						if (step < 0) step = 0;
						if (step > 100) step = 100;
						if (MboxMail::pCUPDUPData)
							MboxMail::pCUPDUPData->SetProgress(step);
					}

					if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
						w = -1;
						break;
					}
				}
			}
		}
		if (findAll) {
			MboxMail::s_mails_find.SetSizeKeepData(m_findAllCount);
			TRACE("DoFastFind: s_mails_find count=%d\n", MboxMail::s_mails_find.GetCount());
		}
	}
	else
	{
		for (int i = which; i >= 0; i--)
		{
			MboxMail *m = MboxMail::s_mails[i];
			bool process = false;
			if (bFilterDates) {
				if (m->m_timeDate >= sd && m->m_timeDate <= ed)
					process = true;
			}
			else
				process = true;
			if (process) 
			{
				if (CheckMatch(i, searchString) >= 0) {
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

				DWORD tc_curr = GetTickCount();
				DWORD tc_elapsed_milliseconds = (tc_curr - tc_start);
				DWORD tc_elapsed_seconds = (tc_curr - tc_start) / 1000;
				//TRACE(_T("index=%d pos=%lld threadId=0x%lx tc_elapsed_milliseconds=%ld tc_elapsed_seconds=%ld\n"),
				//i, pos, myThreadId, tc_elapsed_milliseconds, tc_elapsed_seconds);

				if (mainThreadContext) {
					if ((maxSearchDuration > 0) && (tc_elapsed_seconds > maxSearchDuration)) {
						m_lastFindPos = i - 1;
						w = -2;
						break;
					}
				}
				else {
					if (i % 10 == 0) {
						int mappingsToSearch = which;
						int step = (int)(100.0*((float)(sz - i) / sz));
						if (step < 0) step = 0;
						if (step > 100) step = 100;
						if (MboxMail::pCUPDUPData)
							MboxMail::pCUPDUPData->SetProgress(step);
					}

					if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
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
void NListView::OnUpdateEditFindAgain(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(! m_bInFind && m_list.GetItemCount()>0);
	return;
}

void NListView::OnEditFindAgain() 
{
	if( m_bInFind )
		return;

	if (m_bFindAll) {
		OnEditFind();
		return;
	}

	if( m_searchString.IsEmpty() ) {
		OnEditFind();
		return;
	}
	if (m_bEditFindFirst == TRUE) {
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
		if (m_bFindNext == TRUE)
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
		if (m_bFindNext == TRUE)
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
	if (m_maxSearchDuration > 0) {
		BOOL findAll = FALSE;  // m_findAll should be FALSE
		w = DoFastFind(which, TRUE, m_maxSearchDuration, findAll);
	}

	FIND_ARGS args;
	/*IN*/ args.lview = this; args.searchstart = m_lastFindPos;
	/*OUT*/ args.exitted = FALSE; args.retpos = -1; args.findAll = FALSE; // m_findAll should be FALSE
	if ((w == -2) || (m_maxSearchDuration == 0)) 
	{
		CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcFastSearch, (LPVOID)(FIND_ARGS*)&args);

		INT_PTR nResult = Dlg.DoModal();

		if (!nResult) { // should never be true ?
			m_bInFind = false;
			MboxMail::assert_unexpected();
			return;
		}

		int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
		int retResult = LOWORD(nResult);

		if (retResult != IDOK)
		{  // IDOK==1, IDCANCEL==2
			// We should be here when user selects Cancel button
			//ASSERT(cancelledbyUser == TRUE);
			int loopCnt = 20;
			DWORD tc_start = GetTickCount();
			while ((loopCnt-- > 0) && (args.exitted == FALSE))
			{
				Sleep(25);
			}
			DWORD tc_end = GetTickCount();
			DWORD delta = tc_end - tc_start;
			TRACE("(OnEditFindAgain)Waited %ld milliseconds for thread to exist.\n", delta);
		}
		w = args.retpos;
		MboxMail::pCUPDUPData = NULL;
	}
	if (w >= 0) {
		m_bHighlightAllSet = m_bHighlightAll;
		SelectItemFound(w); // sets m_lastFindPos
	}
	else {
		MessageBeep(MB_OK);
		m_lastFindPos = -1;
	}

	m_bInFind = false;
}

// Dump header fields based on MboxMail::s_mails[which] and corresponding raw data from mbox file
int NListView::DumpSelectedItem(int which)
{
	static BOOL restarted = TRUE;
	char buff[2048];
	char datebuff[256];
	DWORD offset = 0;
	DWORD nwritten = 0;
	DWORD count = 0;
	BOOL retval;

	CString cacheTextFile = "DumpSelectedItem.txt";
	HANDLE hFile = CreateFile(cacheTextFile, GENERIC_WRITE, 0, NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE) {
		DWORD err = GetLastError();
		TRACE(_T("(dumpSelectedItem)INVALID_HANDLE_VALUE error= %ld\n"), err);
		return -1;
	}

	HANDLE mbox_hFile = CreateFile(MboxMail::s_path, GENERIC_READ, 0, NULL,
		OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL);
	if (mbox_hFile == INVALID_HANDLE_VALUE) {
		DWORD err = GetLastError();
		TRACE(_T("(dumpSelectedItem)INVALID_HANDLE_VALUE error= %ld\n"), err);
		CloseHandle(hFile);
		return -1;
	}

	_int64 pos = 0;
	if (restarted == TRUE) {
		SetEndOfFile(hFile);
		restarted = FALSE;
	}
	else
	{
	LARGE_INTEGER li;

		retval = GetFileSizeEx(hFile, &li);
		long long fsize = li.QuadPart;
		pos = FileUtils::FileSeek(hFile, fsize, FILE_BEGIN);
		if (pos < 0) {
			DWORD err = GetLastError();
			CloseHandle(hFile);
			CloseHandle(mbox_hFile);
			return -1;
		}
	}

	MboxMail *m = MboxMail::s_mails[which];

	datebuff[0] = 0;
	if (m->m_timeDate >= 0) {
		MyCTime tt(m->m_timeDate);
		CString lDateTime = tt.FormatGmtTm(m_format);
		strcpy(datebuff, (LPCSTR)lDateTime);
	}

	count = sprintf_s(buff, "INDX=%d first=%lld len=%d last=%lld att=%d hlen=%d rcv=%d date=\"%s\" from=\"%s\" to=\"%s\" subj=\"%s\"\n\n",
		which, m->m_startOff, m->m_length, (m->m_startOff + m->m_length - 1), m->m_hasAttachments, m->m_headLength, m->m_recv,
		datebuff, (LPCSTR)m->m_from, (LPCSTR)m->m_to, (LPCSTR)m->m_subj);

	nwritten = 0;
	if (!WriteFile(hFile, buff, count, &nwritten, NULL)) {
		DWORD retval = GetLastError();
	}
	long long start_offset = m->m_startOff;
	pos = FileUtils::FileSeek(mbox_hFile, start_offset, FILE_BEGIN);

	DWORD bytestoRead = m->m_length;
	DWORD nNumberOfBytesToRead = 1024;
	DWORD lpNumberOfBytesRead = 0;
	retval = TRUE;
	while (bytestoRead > 0)
	{
		if (bytestoRead > 1024)
			nNumberOfBytesToRead = 1024;
		else
			nNumberOfBytesToRead = bytestoRead;

		retval = ReadFile(mbox_hFile, buff, nNumberOfBytesToRead, &lpNumberOfBytesRead, 0);
		if (retval != TRUE) {
			DWORD retval = GetLastError();
			break;
		}
		bytestoRead -= lpNumberOfBytesRead;

		nwritten = 0;
		if (!WriteFile(hFile, buff, lpNumberOfBytesRead, &nwritten, NULL)) {
			DWORD retval = GetLastError();
			break;
		}
	}

	CloseHandle(hFile);
	CloseHandle(mbox_hFile);

	return 1;
}

void NListView::SelectItemFound(int which)
{
	if (which < 0) {
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
	SelectItem(which);
	MboxMail *m = MboxMail::s_mails[which];
	m_lastFindPos = which;
	//dumpSelectedItem(which);
	//DumpItemDetails(which);
}


void NListView::OnEditVieweml()
{	// Save raw message
	if (m_bExportEml == FALSE)
	{
		POSITION pos = m_list.GetFirstSelectedItemPosition();
		int nItem = m_list.GetNextSelectedItem(pos);
		MboxMail *m = 0;
		if (nItem >= 0) {
			m = MboxMail::s_mails[nItem];
			// Get raw mail body
			CString bdy;   m->GetBody(bdy);
			// Save mail
			CFile fp(FileUtils::GetmboxviewTempPath() + "message.eml", CFile::modeWrite | CFile::modeCreate);
			fp.Write(bdy, bdy.GetLength());
			fp.Close();
		}
	}

	CString path = FileUtils::GetmboxviewTempPath();
	HINSTANCE result = ShellExecute(NULL, _T("open"), path, NULL, NULL, SW_SHOWNORMAL);
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
	MarkColumns();
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
		if (m_hMailFile == INVALID_HANDLE_VALUE) {
			DWORD err = GetLastError();
			TRACE(_T("(SetFileMapView)CreateFile failed: INVALID_HANDLE_VALUE error= %ld\n"), err);
			CloseMailFile();
			m_bMappingError = TRUE;
			return FALSE;
		}
		LARGE_INTEGER li;
		GetFileSizeEx(m_hMailFile, &li);
		m_MailFileSize = li.QuadPart;
		if (m_MailFileSize < 10) {
			CloseHandle(m_hMailFile);
			m_hMailFile = INVALID_HANDLE_VALUE;
			TRACE(_T("(SetFileMapView)GetFileSizeEx: Mail file size too small= %lld\n"), m_MailFileSize);
			m_MailFileSize = 0;
			CloseMailFile();
			m_bMappingError = TRUE;
			return FALSE;
		}
		DWORD MailFileSizeLow = (DWORD)m_MailFileSize; // low four bytes bytes m_MailFileSize
		DWORD MailFileSizeHigh = (DWORD)(m_MailFileSize >> 32); // high four bytes of m_MailFileSize
		m_hMailFileMap = CreateFileMapping(m_hMailFile, NULL, PAGE_READONLY, MailFileSizeHigh, MailFileSizeLow, NULL);
		if (m_hMailFileMap == NULL) {
			CloseHandle(m_hMailFile);
			m_hMailFile = INVALID_HANDLE_VALUE;
			DWORD err = GetLastError();
			TRACE(_T("(SetFileMapView)CreateFileMapping failed: INVALID_HANDLE_VALUE error= %ld\n"), err);
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
	else if (m_pMapViewBegin) {
		UnmapViewOfFile(m_pMapViewBegin);
		m_pMapViewBegin = 0;
		m_pMapViewEnd = 0;
		m_curMapBegin = 0;
		m_curMapEnd = 0;
	}

	if (length > (m_mappingSize - m_dwAllocationGranularity)) {
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
		if (dataEnd <= m_mappingSize) {
			aligned_offset = 0;
		}
		else {
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
	if (p != 0) {
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
	else {
		DWORD err = GetLastError();
		CString errTxt;
		errTxt.Format("MapViewOfFileEx failed err=%ld fsize=%lld offset=%lld align_offset=%lld bufSize=%ld mapBegin=%lld mapEnd=%lld\n",
			err, m_MailFileSize, offset, aligned_offset, bufSize, m_pViewBegin, m_pViewEnd);
		TRACE(_T("(SetFileMapView)MapViewOfFileEx failed: error= %ld\n"), err);
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errTxt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
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

BOOL NListView::FindInMailContent(int mailPosition, BOOL bContent, BOOL bAttachment)
{
	char  *pData = 0;
	int datalen = 0;

	SimpleString *outbuf = MboxMail::m_outbuf;
	outbuf->ClearAndResize(10000);

	CString searchString(m_searchString);
	if (m_bCaseSens == 0)
		searchString.MakeLower();

	MboxMail *m = MboxMail::s_mails[mailPosition];
	if (SetupFileMapView(m->m_startOff, m->m_length, m_bFindNext) == FALSE)
		return FALSE;

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
				else
					textPlainFound = TRUE;
			}
			else 
			{
				if (body->m_contentType.CompareNoCase("text/html") != 0)
					continue;
				else
					int deb = 1;
			}

			int bodyLength = body->m_contentLength;
			if ((body->m_contentOffset + body->m_contentLength) > m->m_length) {
				// something is not consistent
				bodyLength = m->m_length - body->m_contentOffset;
			}
			char *bodyBegin = m_pViewBegin + body->m_contentOffset;

			if (body->m_contentTransferEncoding.CompareNoCase("base64") == 0)
			{
				MboxCMimeCodeBase64 d64(bodyBegin, bodyLength);
				int dlength = d64.GetOutputLength();
				outbuf->ClearAndResize(dlength);

				int retlen = d64.GetOutput((unsigned char*)outbuf->Data(), dlength);
				if (retlen > 0) {
					outbuf->SetCount(retlen);
					pData = outbuf->Data();
					datalen = outbuf->Count();
				}
				else {
					outbuf->Clear();
					pData = 0;
					datalen = 0;
				}
			}
			else if (body->m_contentTransferEncoding.CompareNoCase("quoted-printable") == 0)
			{
				MboxCMimeCodeQP dGP(bodyBegin, bodyLength);
				int dlength = dGP.GetOutputLength();
				outbuf->ClearAndResize(dlength);

				int retlen = dGP.GetOutput((unsigned char*)outbuf->Data(), dlength);
				if (retlen > 0) {
					outbuf->SetCount(retlen);
					pData = outbuf->Data();
					datalen = outbuf->Count();
				}
				else {
					outbuf->Clear();
					pData = 0;
					datalen = 0;
				}
			}
			else
			{
				pData = bodyBegin;
				datalen = bodyLength;

			}
			if (pData)
			{
				int pos = -1;
				if (m_bWholeWord)
					pos = g_tu.BMHSearchW((unsigned char *)pData, datalen, (unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_bCaseSens);
				else
					pos = g_tu.BMHSearch((unsigned char *)pData, datalen, (unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_bCaseSens);
				if (pos >= 0) {
					return TRUE;
				}
			}
		}
		if (textPlainFound == FALSE)  // first iteration done; didn't find any plain text blocks
			searchHTML = TRUE;
	}
	return FALSE;
}


// TODO: Below is similar to FindInMailContent above; Merge?, yes when time permits ?
BOOL NListView::AdvancedFindInMailContent(int mailPosition, BOOL bContent, BOOL bAttachment, CFindAdvancedParams &params)
{
	char  *pData = 0;
	int datalen = 0;

	SimpleString *outbuf = MboxMail::m_outbuf;
	outbuf->ClearAndResize(10000);

	MboxMail *m = MboxMail::s_mails[mailPosition];
	if (SetupFileMapView(m->m_startOff, m->m_length, TRUE) == FALSE)
		return FALSE;

	MailBodyContent *body;
	BOOL textPlainFound = FALSE;
	BOOL searchHTML = FALSE;
	BOOL isAttachment = FALSE;
	int i;
	for (i = 0; i < 2; i++) // search plain text blocks first or html text blocks if no text blocks
	{
		for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
		{
			body = m->m_ContentDetailsArray[j];

			isAttachment = FALSE;
			if (!body->m_attachmentName.IsEmpty()) {
				if (bAttachment == FALSE)
					continue;
				else
					isAttachment = TRUE;
			}
			else if (bContent == FALSE)
				continue;

			if (searchHTML == FALSE)
			{
				if (body->m_contentType.CompareNoCase("text/plain") != 0)
					continue;
				else
					textPlainFound = TRUE;
			}
			else
			{
				if (body->m_contentType.CompareNoCase("text/html") != 0)
					continue;
				else
					int deb = 1;
			}

			int bodyLength = body->m_contentLength;
			if ((body->m_contentOffset + body->m_contentLength) > m->m_length) {
				// something is not consistent
				bodyLength = m->m_length - body->m_contentOffset;
			}
			char *bodyBegin = m_pViewBegin + body->m_contentOffset;

			if (body->m_contentTransferEncoding.CompareNoCase("base64") == 0)
			{
				MboxCMimeCodeBase64 d64(bodyBegin, bodyLength);
				int dlength = d64.GetOutputLength();
				outbuf->ClearAndResize(dlength);

				int retlen = d64.GetOutput((unsigned char*)outbuf->Data(), dlength);
				if (retlen > 0) {
					outbuf->SetCount(retlen);
					pData = outbuf->Data();
					datalen = outbuf->Count();
				}
				else {
					outbuf->Clear();
					pData = 0;
					datalen = 0;
				}
			}
			else if (body->m_contentTransferEncoding.CompareNoCase("quoted-printable") == 0)
			{
				MboxCMimeCodeQP dGP(bodyBegin, bodyLength);
				int dlength = dGP.GetOutputLength();
				outbuf->ClearAndResize(dlength);

				int retlen = dGP.GetOutput((unsigned char*)outbuf->Data(), dlength);
				if (retlen > 0) {
					outbuf->SetCount(retlen);
					pData = outbuf->Data();
					datalen = outbuf->Count();
				}
				else {
					outbuf->Clear();
					pData = 0;
					datalen = 0;
				}
			}
			else
			{
				pData = bodyBegin;
				datalen = bodyLength;

			}
			if (pData)
			{  
				// main diff comparing with FindInMailContent
				int pos = -1;
				int fldIndx = 5; // Message
				if (isAttachment)
					fldIndx = 6;

				if (params.m_bEditChecked[fldIndx])
				{
					if (params.m_bWholeWord[fldIndx]) {
						pos = g_tu.BMHSearchW((unsigned char *)pData, datalen,
							(unsigned char *)(LPCSTR)m_stringWithCase[fldIndx], m_stringWithCase[fldIndx].GetLength(), params.m_bCaseSensitive[fldIndx]);
					}
					else {
						pos = g_tu.BMHSearch((unsigned char *)pData, datalen,
							(unsigned char *)(LPCSTR)m_stringWithCase[fldIndx], m_stringWithCase[fldIndx].GetLength(), params.m_bCaseSensitive[fldIndx]);
					}
					if (pos >= 0) {
						return TRUE;
					}
				}

			}
		}
		if (textPlainFound == FALSE)
			searchHTML = TRUE;
	}
	return FALSE;
}

void NListView::PrintMailGroupToText(BOOL multipleSelectedMails, int iItem, int textType, BOOL forceOpen, BOOL printToPrinter, BOOL createFileOnly)
{
	int firstMail = 0;
	int lastMail = 0;
	MailIndexList *selectedMailsIndexList = 0;

	if (multipleSelectedMails == FALSE)
	{
		if (abs(MboxMail::b_mails_which_sorted) != 99) {

			CString txt = _T("Please sort all mails by conversation first.\n");
			txt += "Select \"View\"->\"Sort By\" ->\"Conversation\" or left click on the first column.";
			HWND h = GetSafeHwnd(); // we don't have any window yet
			int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
			return;
		}

		FindFirstAndLastMailOfConversation(iItem, firstMail, lastMail);
	}
	else //  (multipleSelectedMails == TRUE)
	{
		if (PopulateSelectedMailsList() > 0)
			selectedMailsIndexList = &m_selectedMailsList;
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
				CString txt = "Created file \n\n" + textFileName;
				if (createFileOnly) {
					int deb = 1;
				}
				else if (printToPrinter)
				{
					CFile fp;
					if (fp.Open(textFileName, CFile::modeRead | CFile::shareDenyWrite)) 
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
							HINSTANCE result = ShellExecute(h, _T("open"), datapath, NULL, NULL, SW_SHOWNORMAL);
							CMainFrame::CheckShellExecuteResult(result, h);
						}
						int deb = 1;
					}
					else if (nResponse == IDYES)
					{
						HWND h = GetSafeHwnd();
						HINSTANCE result = ShellExecute(h, _T("open"), textFileName, NULL, NULL, SW_SHOWNORMAL);
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
					HINSTANCE result = ShellExecute(h, _T("open"), textFileName, NULL, NULL, SW_SHOWNORMAL);
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
				return -1; // TODO: error ASSERT ?
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
		int deb = 1; // ASSERT ?
	

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

	// ASSERT i <= selectedCnt
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
	CString txt = _T("Do you want to copy all mails to Users Selected Mails?");
	int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
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
	if (m_lastSel != iItem) {
		TRACE("Selecting %d\n", iItem);
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
			pMsgView->FindStringInIHTMLDocument(m_searchStringInMail, m_bWholeWordInMail, m_bCaseSensInMail);

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

	TRACE(_T("\n\tOnItemchangedListCtrl.  Item index: %d SubItem index: %d NewState: %s OldState: %s strChanged: %s\n"),
		pNMListView->iItem, pNMListView->iSubItem, (LPCSTR)strNewState, (LPCSTR)strOldState, (LPCSTR)strChanged);

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

	TRACE(_T("\n\tOnODStateChangedListCtrl.  From Item index: %d To Item index: %d NewState: %s OldState: %s\n"), 
		lpStateChange->iFrom, lpStateChange->iTo, (LPCSTR)strNewState, (LPCSTR)strOldState);

	PrintSelected();

	*pResult = 0;
}

void NListView::OnODFindItemListCtrl(NMHDR* pNMHDR, LRESULT* pResult)
{
	NMLVFINDITEM *lpFindItem = (NMLVFINDITEM*)pNMHDR;

	CString str;
	//str.Format(_T("OnODFindItemListCtrl.  From Item index: %d To Item index: %d"), lpStateChange->iFrom, lpStateChange->iTo);

	*pResult = 0;
}


void NListView::OnReleaseCaptureListCtrl(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMHDR lpReleaseCapture = (LPNMHDR)pNMHDR;

	CString str;
	//str.Format(_T("OnReleaseCaptureListCtrl.  From Item index: %d To Item index: %d"), lpReleaseCapture->iFrom, lpStateChange->iTo);

	*pResult = 0;
}

void NListView::OnODCacheHintListCtrl(NMHDR* pNMHDR, LRESULT* pResult)
{
	NMLVCACHEHINT  *lpCacheHint = (NMLVCACHEHINT*)pNMHDR;

	CString str;
	//str.Format(_T("OnODCacheHintListCtrl.  From Item index: %d To Item index: %d"), lpReleaseCapture->iFrom, lpStateChange->iTo);

	*pResult = 0;
}


void NListView::ItemState2Str(UINT uState, CString &strState)
{
	strState.Empty();
	switch (uState)
	{
	case LVIS_ACTIVATING: strState.Append("LVIS_ACTIVATING"); break;
	case LVIS_CUT: strState.Append("LVIS_CUT"); break;
	case LVIS_DROPHILITED: strState.Append("LVIS_DROPHILITED"); break;
	case LVIS_FOCUSED: strState.Append("LVIS_FOCUSED"); break;
	case LVIS_OVERLAYMASK: strState.Append("LVIS_OVERLAYMASK"); break;
	case LVIS_SELECTED: strState.Append("LVIS_SELECTED"); break;
	case LVIS_STATEIMAGEMASK: strState.Append("LVIS_STATEIMAGEMASK"); break;
	default: strState.Append("LVIS_UNKNOWN"); break;
	}
}


void NListView::ItemChange2Str(UINT uChange, CString &strState)
{
	strState.Empty();
	// if (uChange & LVIF_COLFMT) strState.Append("LVIF_COLFMT ");   //  require (NTDDI_VERSION >= NTDDI_VISTA)
	if (uChange & LVIF_COLUMNS) strState.Append("LVIF_COLUMNS ");
	if (uChange & LVIF_DI_SETITEM) strState.Append("LVIF_DI_SETITEM ");
	if (uChange & LVIF_GROUPID) strState.Append("LVIF_GROUPID ");
	if (uChange & LVIF_IMAGE) strState.Append("LVIF_IMAGE ");

	if (uChange & LVIF_INDENT) strState.Append("LVIF_INDENT ");
	if (uChange & LVIF_NORECOMPUTE) strState.Append("LVIF_NORECOMPUTE ");
	if (uChange & LVIF_PARAM) strState.Append("LVIF_PARAM ");
	if (uChange & LVIF_STATE) strState.Append("LVIF_STATE ");
	if (uChange & LVIF_TEXT) strState.Append("LVIF_TEXT ");

	if (strState.IsEmpty()) strState.Append("LVIS_UNKNOWN");
}

// Debug support
void NListView::PrintSelected()
{
	// Get the selected items in the control
	POSITION p = m_list.GetFirstSelectedItemPosition();
	CString msg = _T("");
	CString str;

	if (p)
		msg = _T("The following items are selected: \n\t");

	while (p)
	{
		int nSelected = m_list.GetNextSelectedItem(p);

		str.Empty();
		str.Format(_T("%d"), nSelected);

		if (p)
			str += _T(", ");

		msg += str;
		if (msg.GetLength() > 100) {
			TRACE(_T("%s\n"), (LPCSTR)msg);
			msg.Empty();
			msg.Append("cont:\n\t");
		}
	}
	if (msg.IsEmpty()) {
		msg = _T("There are no selected items");
	}
	TRACE(_T("%s\n"), (LPCSTR)msg);
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
		if (nWhichMailList == IDC_ARCHIVE_LIST) {
			MboxMail::s_mails_all.Copy(MboxMail::s_mails);
			MboxMail::m_allMails.m_lastSel = m_lastSel;
			MboxMail::m_allMails.b_mails_which_sorted = MboxMail::b_mails_which_sorted;
		}
		else if (nWhichMailList == IDC_FIND_LIST) {
			MboxMail::s_mails_find.CopyKeepData(MboxMail::s_mails);
			MboxMail::m_findMails.m_lastSel = m_lastSel;
			MboxMail::m_findMails.b_mails_which_sorted = MboxMail::b_mails_which_sorted;
		}
		else if (nWhichMailList == IDC_EDIT_LIST) {
			MboxMail::s_mails_edit.CopyKeepData(MboxMail::s_mails);
			MboxMail::m_editMails.m_lastSel = m_lastSel;
			MboxMail::m_editMails.b_mails_which_sorted = MboxMail::b_mails_which_sorted;
		}
		else if (nWhichMailList == IDC_LABEL_LIST) {
			MboxMail::s_mails_label.CopyKeepData(MboxMail::s_mails);
			MboxMail::m_labelMails.m_lastSel = m_lastSel;
			MboxMail::m_labelMails.b_mails_which_sorted = MboxMail::b_mails_which_sorted;
		}
		else if (nWhichMailList == IDC_FOLDER_LIST) {
			MboxMail::s_mails_folder.CopyKeepData(MboxMail::s_mails);
			MboxMail::m_folderMails.m_lastSel = m_lastSel;
			MboxMail::m_folderMails.b_mails_which_sorted = MboxMail::b_mails_which_sorted;
		}
		else
			; // TODO: ASSERT ?
	}

	// setup environment for new mail list
	if (nID == IDC_ARCHIVE_LIST) {
		MboxMail::s_mails.CopyKeepData(MboxMail::s_mails_all);
		m_lastSel = MboxMail::m_allMails.m_lastSel;
		MboxMail::m_mailList = &MboxMail::m_allMails;
		MboxMail::b_mails_which_sorted = MboxMail::m_allMails.b_mails_which_sorted;
	}
	else if (nID == IDC_FIND_LIST) {
		MboxMail::s_mails.CopyKeepData(MboxMail::s_mails_find);
		m_lastSel = MboxMail::m_findMails.m_lastSel;
		MboxMail::m_mailList = &MboxMail::m_findMails;
		MboxMail::b_mails_which_sorted = MboxMail::m_findMails.b_mails_which_sorted;
	}
	else if (nID == IDC_EDIT_LIST) {
		MboxMail::s_mails.CopyKeepData(MboxMail::s_mails_edit);
		m_lastSel = MboxMail::m_editMails.m_lastSel;
		MboxMail::m_mailList = &MboxMail::m_editMails;
		MboxMail::b_mails_which_sorted = MboxMail::m_editMails.b_mails_which_sorted;
	}
	else if (nID == IDC_LABEL_LIST) {
		MboxMail::s_mails.CopyKeepData(MboxMail::s_mails_label);
		m_lastSel = MboxMail::m_labelMails.m_lastSel;
		MboxMail::m_mailList = &MboxMail::m_labelMails;
		MboxMail::b_mails_which_sorted = MboxMail::m_labelMails.b_mails_which_sorted;
	}
	else if (nID == IDC_FOLDER_LIST) {
		MboxMail::s_mails.CopyKeepData(MboxMail::s_mails_folder);
		m_lastSel = MboxMail::m_folderMails.m_lastSel;
		MboxMail::m_mailList = &MboxMail::m_folderMails;
		MboxMail::b_mails_which_sorted = MboxMail::m_folderMails.b_mails_which_sorted;
	}
	else
		; // TODO: ASSERT ?

	if (abs(MboxMail::b_mails_which_sorted) == 99) {
		MboxMail::SortByGroupId(0, MboxMail::b_mails_which_sorted < 0);  // TODO: this is probably not needed; keep it simple for now
		MboxMail::assignColor2ConvesationGroups(&MboxMail::s_mails);
	}

	m_list.SetItemCount(MboxMail::s_mails.GetCount());
	MboxMail::nWhichMailList = nID;
	pFrame->SetMailList(nID);


#if 0
	if (abs(MboxMail::b_mails_which_sorted) == 99)
		RefreshSortByColumn();
#endif

	MarkColumns();  // not used anymore -Invalidate() below doesn't updates column labels

	int sv_lastSel = m_lastSel;
	ClearDescView();  // set m_lastSel to -1 ; m_lastFindPos is not updated
	m_lastSel = sv_lastSel;

	Invalidate();

	NMsgView *pMsgView = pFrame->GetMsgView();
	if (pMsgView)
		pMsgView->Invalidate(); 
		//pMsgView->m_browser.m_ie.Invalidate(); // TODO: changed to GetMsgView()->Invalidate(); and it seem to work :)

	int mailCnt = MboxMail::s_mails.GetCount();
	if (mailCnt > 0) {
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

BOOL NListView::IsSingleAddress(CString *to)
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

void NListView::TrimToAddr(CString *to, CString &toAddr, int maxNumbOfAddr)
{
	SimpleString name;
	SimpleString addr;
	CString tmpAddr;
	CString tmp;

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

void NListView::TrimToName(CString *to, CString &toName, int maxNumbOfAddr)
{
	SimpleString name;
	SimpleString addr;
	CString tmpName;
	CString tmp;

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

BOOL NListView::PrepopulateAdvancedSearchParams(CString *from, CString *to, CString *subject)
{
	if (!(from && to && subject))
		return FALSE;

	SimpleString name;
	SimpleString addr;
	CString tmpAddr;
	CString fromAddr;
	CString toAddr;
	CString subjAddr;
	CString subj;
	CString tmp;

	for (int i = 3; i < FILTER_FIELDS_NUMB; i++)
	{
		m_advancedParams.m_string[i].Empty();
		m_advancedParams.m_bEditChecked[i] = FALSE;
		m_advancedParams.m_bCaseSensitive[i] = FALSE;
		m_advancedParams.m_bWholeWord[i] = FALSE;
	}

	int fromlen = from->GetLength();
	name.ClearAndResize(fromlen);
	addr.ClearAndResize(fromlen);
	MboxMail::splitMailAddress(from->operator LPCSTR(), fromlen, &name, &addr);
	fromAddr = addr.Data();
	fromAddr.Trim(" \t\"<>");

	m_advancedParams.m_string[0].Empty();
	m_advancedParams.m_string[0].Append(fromAddr);
	if (!fromAddr.IsEmpty()) {
		m_advancedParams.m_bEditChecked[0] = TRUE;
		m_advancedParams.m_bCaseSensitive[0] = FALSE;
	}
	else
		m_advancedParams.m_bEditChecked[0] = FALSE;

	int maxNumbOfAddr = 1;
	TrimToAddr(to, toAddr, maxNumbOfAddr);

	m_advancedParams.m_string[1].Empty();
	m_advancedParams.m_string[1].Append(toAddr);
	if (!toAddr.IsEmpty()) {
		m_advancedParams.m_bEditChecked[1] = TRUE;
		m_advancedParams.m_bCaseSensitive[1] = FALSE;
	}
	else
		m_advancedParams.m_bEditChecked[1] = FALSE;

	subj.Append(*subject);

	int length = subj.GetLength();
	for (int i = 0; i < 10; i++) // to be safe, limit number of iterations
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

	m_advancedParams.m_string[2].Empty();
	m_advancedParams.m_string[2].Append(subj);
	if (!subj.IsEmpty()) {
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
		MyCTime maxT(max);
		SYSTEMTIME endDateSystemTime;
		maxT.GetAsSystemTime(endDateSystemTime);
		endDateSystemTime.wMinute = 59;
		endDateSystemTime.wHour = 23;
		endDateSystemTime.wSecond = 59;
		endDateSystemTime.wMilliseconds = 999;
		MyCTime::fixSystemtime(&endDateSystemTime);
		maxT.SetDateTime(endDateSystemTime);

		maxTime = maxT;
	}

	if ((minTime.GetTime() < 0) || (maxTime.GetTime() < 0))
		int deb = 1;
}

void NListView::EditFindAdvanced(CString *from, CString *to, CString *subject)
{
	// TODO: Add your command handler code here

	if (m_bInFind)
		return;

	m_advancedFind = TRUE;

	m_bInFind = true;

	CFindAdvancedDlg dlg;

	// from && to && subject are set when Advanced Find is invoked as menu option on the selected item
	// when invoked via Edit -> Find Advanced menu option, they all will be set to NULL
	if (from && to && subject)
	{
		PrepopulateAdvancedSearchParams(from, to, subject);
	}

	dlg.m_params.Copy(m_advancedParams);

	// The m_lastStartDate and m_lastEndDate are shared across Find and Advanced Find
	// Should eventually make them independed
	// MyCTime m_lastStartDate;
	// MyCTime m_lastEndDate;

	if (m_bNeedToFindMailMinMaxTime)
	{
		FindMinMaxTime(m_mboxMailStartDate, m_mboxMailEndDate);
		m_lastStartDate = m_mboxMailStartDate;
		m_lastEndDate = m_mboxMailEndDate;
		m_bNeedToFindMailMinMaxTime = FALSE;
	}

	if (MboxMail::nWhichMailList != IDC_ARCHIVE_LIST)
	{
		FindMinMaxTime(m_lastStartDate, m_lastEndDate);
		m_needToRestoreArchiveListDateTime = TRUE;
	}
	else if ((m_advancedParams.m_filterDates == FALSE) || m_needToRestoreArchiveListDateTime)
	{
		m_lastStartDate = m_mboxMailStartDate;
		m_lastEndDate = m_mboxMailEndDate;
		m_needToRestoreArchiveListDateTime = FALSE;
	}

	BOOL retDT = MyCTimeToOleTime(m_lastStartDate, dlg.m_params.m_startDate);
	retDT = MyCTimeToOleTime(m_lastEndDate, dlg.m_params.m_endDate);

	if (dlg.DoModal() == IDOK) 
	{
		retDT = NListView::OleTime2MyCTime(dlg.m_params.m_startDate, m_lastStartDate, FALSE);
		retDT = NListView::OleTime2MyCTime(dlg.m_params.m_endDate, m_lastEndDate, TRUE);

		m_advancedParams.Copy(dlg.m_params);
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
			FIND_ARGS args;
			/*IN*/ args.lview = this; args.searchstart = m_lastFindPos;
			/*OUT*/ args.exitted = FALSE; args.retpos = -1; args.findAll = TRUE;
	
			{
				CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcFastSearch, (LPVOID)(FIND_ARGS*)&args);

				INT_PTR nResult = Dlg.DoModal();

				if (!nResult) { // should never be true ? 
					m_bInFind = false;
					MboxMail::assert_unexpected();
					return;
				}

				int cancelledbyUser = HIWORD(nResult); // when Cancel button is selected
				int retResult = LOWORD(nResult);

				if (retResult != IDOK)
				{  // IDOK==1, IDCANCEL==2
					// We should be here when user selects Cancel button
					//ASSERT(cancelledbyUser == TRUE);
					int loopCnt = 20;
					DWORD tc_start = GetTickCount();
					while ((loopCnt-- > 0) && (args.exitted == FALSE))
					{
						Sleep(25);
					}
					DWORD tc_end = GetTickCount();
					DWORD delta = tc_end - tc_start;
					TRACE("(EditFindAdvanced)Waited %ld milliseconds for thread to exist.\n", delta);
				}
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
			return;
		}
		else {
			MessageBeep(MB_OK);
			m_lastFindPos = -1;
		}
	}
	m_bInFind = false;
}

int NListView::MatchHeaderFldSingleAddress(int fldIndx, CString &hdrFld, CFindAdvancedParams &params, int pos)
{
	if (params.m_bSingleTo)
	{
		BOOL isSingle = IsSingleAddress(&hdrFld);
		if (isSingle)
			pos = MatchHeaderFld(fldIndx, hdrFld, params, pos);
		else
			pos = -2;
	}
	else
	{
		pos = MatchHeaderFld(fldIndx, hdrFld, params, pos);
	}
	return pos;
}

int NListView::MatchHeaderFld(int fldIndx, CString &hdrFld, CFindAdvancedParams &params, int pos)
{
	if (params.m_bEditChecked[fldIndx])
	{
		int fldLength = m_stringWithCase[fldIndx].GetLength();
		CString &fld = m_stringWithCase[fldIndx];

		if (!hdrFld.IsEmpty() && (fldLength == 1 && (fld.GetAt(0) == '*')))
		{
			pos = 1;
		}
		else if (params.m_bWholeWord[fldIndx]) 
		{
			pos = g_tu.BMHSearchW((unsigned char *)(LPCSTR)hdrFld, hdrFld.GetLength(),
				(unsigned char *)(LPCSTR)fld, fldLength, params.m_bCaseSensitive[fldIndx]);
		}
		else 
		{
			pos = g_tu.BMHSearch((unsigned char *)(LPCSTR)hdrFld, hdrFld.GetLength(),
				(unsigned char *)(LPCSTR)fld, fldLength, params.m_bCaseSensitive[fldIndx]);
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

		fldIndx = 0; // From
		fromPos = MatchHeaderFld(fldIndx, m->m_from, params);

		fldIndx = 1; // To
		toPos = MatchHeaderFldSingleAddress(fldIndx, m->m_to, params);

		if ((fromPos < 0) || (toPos < 0))
		{
			fldIndx = 1; // To search string is now From
			fromPos = MatchHeaderFld(fldIndx, m->m_from, params);
			if (fromPos < 0)
				return -1;

			fldIndx = 0; // From search string is now To
			toPos = MatchHeaderFldSingleAddress(fldIndx, m->m_to, params);
			if (toPos < 0)
				return -1;
		}

		fldIndx = 3; // CC
		ccPos = MatchHeaderFld(fldIndx, m->m_cc, params);
		if (ccPos < 0)
			return -1;

		fldIndx = 4; // BCC
		bccPos = MatchHeaderFld(fldIndx, m->m_bcc, params);
		if (bccPos < 0)
			return -1;
	}
	else if (params.m_filterNumb == 1)
	{
		// "(From->To) and CC and BCC and Subject and (Message Text or Attachment Text) and Attachment Name"

		fldIndx = 0; // From
		fromPos = MatchHeaderFld(fldIndx, m->m_from, params);
		if (fromPos < 0)
			return -1;

		fldIndx = 1; // To
		toPos = MatchHeaderFldSingleAddress(fldIndx, m->m_to, params);
		if (toPos < 0)
			return -1;

		fldIndx = 3; // CC
		ccPos = MatchHeaderFld(fldIndx, m->m_cc, params);
		if (ccPos < 0)
			return -1;

		fldIndx = 4; // BCC
		bccPos = MatchHeaderFld(fldIndx, m->m_bcc, params);
		if (bccPos < 0)
			return -1;
	}
	else if (params.m_filterNumb == 2)
	{
		// "((From -> (To or CC or BCC)) and Subject and (Message Text or Attachment Text) and Attachment Name"

		fldIndx = 0; // From
		fromPos = MatchHeaderFld(fldIndx, m->m_from, params);
		if (fromPos < 0)
			return -1;

		// need special hack/check first because of "To or CC or BCC) and not "To and CC and BCC"
		// if none of the fields set, assume it is a match
		if (params.m_bEditChecked[1] || params.m_bEditChecked[3] || params.m_bEditChecked[4])
		{
			fldIndx = 1; // To
			toPos = -1;
			toPos = MatchHeaderFldSingleAddress(fldIndx, m->m_to, params, toPos);

			if (toPos < 0)
			{
				fldIndx = 3; // CC
				ccPos = -1;
				ccPos = MatchHeaderFld(fldIndx, m->m_cc, params, ccPos);

				if (ccPos < 0)
				{
					fldIndx = 4; // BCC
					bccPos = -1;
					bccPos = MatchHeaderFld(fldIndx, m->m_bcc, params, bccPos);
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
		fromPos = MatchHeaderFld(fldIndx, m->m_from, params);

		fldIndx = 1; // To
		toPos = MatchHeaderFldSingleAddress(fldIndx, m->m_to, params);

		if ((fromPos < 0) || (toPos < 0))
		{
			fldIndx = 1; // To search string is now From
			fromPos = MatchHeaderFld(fldIndx, m->m_from, params);
			if (fromPos < 0)
				return -1;

			fldIndx = 0; // From search string is now To
			toPos = MatchHeaderFldSingleAddress(fldIndx, m->m_to, params);
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
		fromPos = MatchHeaderFld(fldIndx, m->m_from, params);
		if (fromPos < 0)
			return -1;

		fldIndx = 1; // To
		toPos = MatchHeaderFldSingleAddress(fldIndx, m->m_to, params);
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
			fromPos = MatchHeaderFld(fldIndx, m->m_from, params, fromPos);

			if (fromPos < 0)
			{
				fldIndx = 1; // To
				toPos = -1;
				toPos = MatchHeaderFldSingleAddress(fldIndx, m->m_to, params, toPos);

				if (toPos < 0)
				{
					fldIndx = 3; // CC
					ccPos = -1;
					ccPos = MatchHeaderFld(fldIndx, m->m_cc, params, ccPos);

					if (ccPos < 0)
					{
						fldIndx = 4; // BCC
						bccPos = -1;
						bccPos = MatchHeaderFld(fldIndx, m->m_bcc, params, bccPos);
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
		CString &fld = m_stringWithCase[fldIndx];

		if (!m->m_subj.IsEmpty() && ((fldLength == 1) && (fld.GetAt(0) == '*'))) {
			subjectPos = 1;
		}
		else if (params.m_bWholeWord[fldIndx]) {
			subjectPos = g_tu.BMHSearchW((unsigned char *)(LPCSTR)m->m_subj, m->m_subj.GetLength(),
				(unsigned char *)(LPCSTR)fld, fldLength, params.m_bCaseSensitive[fldIndx]);
		}
		else {
			subjectPos = g_tu.BMHSearch((unsigned char *)(LPCSTR)m->m_subj, m->m_subj.GetLength(),
				(unsigned char *)(LPCSTR)fld, fldLength, params.m_bCaseSensitive[fldIndx]);
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
	EditFindAdvanced(&m->m_from, &m->m_to, &m->m_subj);

	int deb = 1;
}

#define MAIL_LIST_VERSION_BASE  0x73215500
#define MAIL_LIST_VERSION  (MAIL_LIST_VERSION_BASE+1)

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
		mboxFileSuffix = "_USER.mbox";
	else if (MboxMail::IsFindMailsSelected())
		mboxFileSuffix = "_FIND.mbox";
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
	CString rootPrintSubFolder = "ArchiveCache";
	CString targetPrintSubFolder;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, archiveCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CString mboxFile = fileNameBase + fileNameExtention;
	CString mboxFilePath = archiveCachePath + "\\" + fileNameBase + mboxFileSuffix;

	//CString path = MboxMail::GetLastPath();
	//CString mboxFile = path + "\\" + fileNameBase + mboxFileSuffix;

	if (!FileUtils::PathFileExist(mboxFilePath))
	{
		CString txt = _T("File \"") + mboxFilePath;
		txt += _T("\" doesn't exist. Can't open location.");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}
	if (FileUtils::BrowseToFile(mboxFilePath) == FALSE) {  // TODO: s_path error checking ??
		HWND h = GetSafeHwnd();
		HINSTANCE result = ShellExecute(h, _T("open"), archiveCachePath, NULL, NULL, SW_SHOWNORMAL);
		CMainFrame::CheckShellExecuteResult(result, h);
	}
	return 1;
}

int NListView::OpenMailListFileLocation()
{
	CString mboxFileSuffix;
	if (MboxMail::IsUserMailsSelected())
		mboxFileSuffix = "_USER.mbox.mboxlist";
	else if (MboxMail::IsFindMailsSelected())
		mboxFileSuffix = "_FIND.mbox.mboxlist";
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
	CString rootPrintSubFolder = "ArchiveCache";
	CString targetPrintSubFolder;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, archiveCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CString mboxFile = fileNameBase + fileNameExtention;
	CString mboxFilePath = archiveCachePath + "\\" + fileNameBase + mboxFileSuffix;

	//CString path = MboxMail::GetLastPath();
	//CString mboxFile = path + "\\" + fileNameBase + mboxFileSuffix;

	if (!FileUtils::PathFileExist(mboxFilePath))
	{
		CString txt = _T("File \"") + mboxFilePath;
		txt += _T("\" doesn't exist. Can't open location.");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}
	if (FileUtils::BrowseToFile(mboxFilePath) == FALSE) {  // TODO: s_path error checking ??
		HWND h = GetSafeHwnd();
		HINSTANCE result = ShellExecute(h, _T("open"), archiveCachePath, NULL, NULL, SW_SHOWNORMAL);
		CMainFrame::CheckShellExecuteResult(result, h);
	}
	return 1;
}

// Below ADDED in 1.0.3.1
int NListView::RemoveDuplicateMails(MailArray &s_mails_array)
{
	m_list.SetRedraw(FALSE);

	MboxMail::RemoveDuplicateMails(s_mails_array);

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
	if (errorCnt != 0) {
#ifdef _DEBUG
		ASSERT(0);
#endif
		int deb = 1;
	}
	return errorCnt;
}

int NListView::PrintMailConversationToSeparatePDF_Thread(int iItem, CString &errorText)
{
	int firstMail = 0;
	int lastMail = 0;

	FindFirstAndLastMailOfConversation(iItem, firstMail, lastMail);

	CString targetPrintSubFolderName = "PDF_GROUP";
	int ret = PrintMailRangeToSeparatePDF_Thread(firstMail, lastMail, targetPrintSubFolderName);

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

int NListView::PrintMailRangeToSeparatePDF_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS args;

	CString rootPrintSubFolder = "PrintCache";
	CString printCachePath;
	CString errorText;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (pFrame)
	{
		int ret = pFrame->VerifyPathToHTML2PDFExecutable(errorText);
		if (ret < 0)
		{
			HWND h = NULL; 
			int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else
	{
		HWND h = NULL;
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		errorText.Append("Internal error. Try again.");
		return - 1;
	}

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolderName, printCachePath, errorText);
	if (retval == FALSE) 
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	/*IN*/
	args.lview = this;
	args.targetPrintFolderPath = printCachePath;
	args.targetPrintSubFolderName = targetPrintSubFolderName;
	args.firstMail = firstMail;
	args.lastMail = lastMail;
	args.selectedMailIndexList = 0;
	args.nItem = -1;
	args.separatePDFs = TRUE;

	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcPrintMailGroupToSeparatePDF, (LPVOID)(PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS*)&args);

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
		//ASSERT(cancelledbyUser == TRUE);
		int loopCnt = 20;
		DWORD tc_start = GetTickCount();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		DWORD tc_end = GetTickCount();
		DWORD delta = tc_end - tc_start;
		TRACE("(PrintMailRangeToSeparatePDF_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret > 0)
	{
		if (args.errorText.IsEmpty())  // TODO: should always be true
		{
			CString txt = "Created separate PDF files in \n\n" + printCachePath + " \n\ndirectory.";
			OpenContainingFolderDlg dlg(txt, TRUE);
			INT_PTR nResponse = dlg.DoModal();
			if (nResponse == IDOK)
			{
				HWND h = GetSafeHwnd();
				HINSTANCE result = ShellExecute(h, _T("open"), printCachePath, NULL, NULL, SW_SHOWNORMAL);
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
			MboxMail::assert_unexpected();
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}
	return 1;
}

int NListView::PrintMailRangeToSinglePDF_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS args;

	CString rootPrintSubFolder = "PrintCache";
	CString printCachePath;
	CString errorText;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (pFrame)
	{
		int ret = pFrame->VerifyPathToHTML2PDFExecutable(errorText);
		if (ret < 0)
		{
			HWND h = NULL;
			int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else
	{
		HWND h = NULL;
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		errorText.Append("Internal error. Try again.");
		return -1;
	}

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolderName, printCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	/*IN*/
	args.lview = this;
	args.targetPrintFolderPath = printCachePath;
	args.targetPrintSubFolderName = targetPrintSubFolderName;
	args.firstMail = firstMail;
	args.lastMail = lastMail;
	args.selectedMailIndexList = 0;
	args.nItem = -1;
	args.separatePDFs = FALSE;

	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcPrintMailGroupToSinglePDF, (LPVOID)(PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS*)&args);

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
		//ASSERT(cancelledbyUser == TRUE);
		int loopCnt = 20;
		DWORD tc_start = GetTickCount();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		DWORD tc_end = GetTickCount();
		DWORD delta = tc_end - tc_start;
		TRACE("(PrintMailRangeToSeparatePDF_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret > 0)
	{
		if (args.errorText.IsEmpty())
		{
			int fileType = 2;
			bool fileExists;
			CString pdfFileName;

			if (firstMail != lastMail)
				int ret = MboxMail::MakeFileNameFromMailArchiveName(fileType, pdfFileName, targetPrintSubFolderName, fileExists, errorText);
			else
				int ret = MboxMail::MakeFileNameFromMailHeader(firstMail, fileType, pdfFileName, targetPrintSubFolderName, fileExists, errorText);

			CString txt = "Created PDF file \n\n" + pdfFileName;
			OpenContainingFolderDlg dlg(txt, FALSE);
			INT_PTR nResponse = dlg.DoModal();
			////////////
			if (nResponse == IDOK)
			{
				if (FileUtils::BrowseToFile(pdfFileName) == FALSE) {
					HWND h = GetSafeHwnd();
					HINSTANCE result = ShellExecute(h, _T("open"), printCachePath, NULL, NULL, SW_SHOWNORMAL);
					CMainFrame::CheckShellExecuteResult(result, h);
				}
				int deb = 1;
			}
			else if (nResponse == IDYES)
			{
				HWND h = GetSafeHwnd();
				HINSTANCE result = ShellExecute(h, _T("open"), pdfFileName, NULL, NULL, SW_SHOWNORMAL);
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
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}
	return 1;
}

int NListView::PrintMailRangeToSeparatePDF_WorkerThread(int firstMail, int lastMail, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText)
{
	BOOL progressBar = TRUE;

	int retdel = NListView::DeleteAllHtmAndPDFFiles(targetPrintFolderPath);

	if (progressBar && MboxMail::pCUPDUPData)
	{
		MboxMail::pCUPDUPData->SetProgress(_T("Printing mails to PDF files ..."), 0);
	}

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	double delta = (double)(lastMail - firstMail + 1);
	if (delta <= 0) delta = 1;
	double step = delta / 100;
	double curstep = 1;
	double newstep = 0;
	CString fileNum;
	int nFileNum;

	if (progressBar && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	BOOL progress_Bar = FALSE;
	int mailCnt = lastMail - firstMail + 1;
	for (int i = firstMail; i <= lastMail; i++)
	{
		if (pFrame)
		{
			// PrintSingleMailtoPDF will make a unique name and print HTML to PDF
			pFrame->PrintSingleMailtoPDF(i, targetPrintSubFolderName, progress_Bar, errorText);
		}

		if (progressBar && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				break;
			}

			newstep = ((double)(i - firstMail + 1)) / step;
			if (newstep >= curstep)
			{
				curstep = newstep;
			}

			nFileNum = ((i - firstMail) + 1);
			fileNum.Format(_T("Printing mails to PDF files ... %d of %d"), nFileNum, mailCnt);
			if (MboxMail::pCUPDUPData)
				MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(curstep));
		}

	}
	nFileNum = mailCnt;
	newstep = 100;
	fileNum.Format(_T("Printing mails to PDF files ... %d of %d"), nFileNum, mailCnt);
	if (MboxMail::pCUPDUPData)  MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(newstep));

	return  1;
}

int NListView::PrintMailRangeToSinglePDF_WorkerThread(int firstMail, int lastMail, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText)
{
	BOOL progressBar = TRUE;

	TEXTFILE_CONFIG textConfig;
	CString htmFileName;
	int textType = 1;

	textConfig.m_dateFormat = m_format;
	textConfig.m_bGMTTime = m_gmtTime;
	textConfig.m_nCodePageId = CP_UTF8;

	int ret = MboxMail::PrintMailRangeToSingleTextFile_WorkerThread(textConfig, htmFileName, firstMail, lastMail, textType, errorText);

	if (progressBar && MboxMail::pCUPDUPData)
	{
		MboxMail::pCUPDUPData->SetProgress(_T("Printing mails to single PDF file ..."), 0);
	}

	double delta = (double)(lastMail - firstMail + 1);
	CString progressText;
	if (progressBar && MboxMail::pCUPDUPData)
	{
		progressText.Format(_T("Printing %d mails to single PDF file ..."), (int)delta);
		MboxMail::pCUPDUPData->SetProgress(progressText, 1);
	}

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (delta <= 0) delta = 1;
	double step = delta / 100;
	double curstep = 1;

	if (progressBar && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	ret = CMainFrame::ExecCommand_WorkerThread(htmFileName, errorText, TRUE, progressText);

	return  1;
}

int NListView::PrintMailSelectedToSeparatePDF_Thread(CString &targetPrintSubFolderName, CString &targetPrintFolderPath)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS args;

	CString rootPrintSubFolder = "PrintCache";
	CString targetPrintSubFolder = targetPrintSubFolderName;
	CString printCachePath;
	CString errorText;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (pFrame)
	{
		int ret = pFrame->VerifyPathToHTML2PDFExecutable(errorText);
		if (ret < 0)
		{
			HWND h = NULL;
			int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else
	{
		HWND h = NULL;
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		errorText.Append("Internal error. Try again.");
		return -1;
	}

	MailIndexList *selectedMailsIndexList = 0;
	if (PopulateSelectedMailsList() > 0)
		selectedMailsIndexList = &m_selectedMailsList;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, printCachePath, errorText);
	if (retval == FALSE) {
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
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

	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcPrintMailGroupToSeparatePDF, (LPVOID)(PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS*)&args);

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
		//ASSERT(cancelledbyUser == TRUE);
		int loopCnt = 20;
		DWORD tc_start = GetTickCount();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		DWORD tc_end = GetTickCount();
		DWORD delta = tc_end - tc_start;
		TRACE("(PrintMailSelectedToSeparatePDF_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret > 0)
	{
		if (args.errorText.IsEmpty())
		{

			CString txt = "Created separate PDF files in \n\n" + printCachePath + " \n\ndirectory.";
			OpenContainingFolderDlg dlg(txt, TRUE);
			INT_PTR nResponse = dlg.DoModal();
			if (nResponse == IDOK)
			{
				HWND h = GetSafeHwnd();
				HINSTANCE result = ShellExecute(h, _T("open"), printCachePath, NULL, NULL, SW_SHOWNORMAL);
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
			MboxMail::assert_unexpected();
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}
	return 1;
}

int NListView::PrintMailSelectedToSinglePDF_Thread(CString &targetPrintSubFolderName, CString &targetPrintFolderPath)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS args;

	CString rootPrintSubFolder = "PrintCache";
	CString targetPrintSubFolder = targetPrintSubFolderName;
	CString printCachePath;
	CString errorText;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (pFrame)
	{
		int ret = pFrame->VerifyPathToHTML2PDFExecutable(errorText);
		if (ret < 0)
		{
			HWND h = NULL;
			int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else
	{
		HWND h = NULL;
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		errorText.Append("Internal error. Try again.");
		return -1;
	}

	MailIndexList *selectedMailsIndexList = 0;
	if (PopulateSelectedMailsList() > 0)
		selectedMailsIndexList = &m_selectedMailsList;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, printCachePath, errorText);
	if (retval == FALSE) {
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
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
		//ASSERT(cancelledbyUser == TRUE);
		int loopCnt = 20;
		DWORD tc_start = GetTickCount();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		DWORD tc_end = GetTickCount();
		DWORD delta = tc_end - tc_start;
		TRACE("(PrintMailSelectedToSinglePDF_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret > 0)
	{
		if (args.errorText.IsEmpty())
		{
			int textType = 2;
			bool fileExists;
			CString pdfFileName;

			int ret = MboxMail::MakeFileNameFromMailArchiveName(textType, pdfFileName, targetPrintSubFolder, fileExists, errorText);

			CString txt = "Created PDF file \n\n" + pdfFileName;
			OpenContainingFolderDlg dlg(txt, FALSE);
			INT_PTR nResponse = dlg.DoModal();
			////////////
			if (nResponse == IDOK)
			{
				if (FileUtils::BrowseToFile(pdfFileName) == FALSE) {
					HWND h = GetSafeHwnd();
					HINSTANCE result = ShellExecute(h, _T("open"), printCachePath, NULL, NULL, SW_SHOWNORMAL);
					CMainFrame::CheckShellExecuteResult(result, h);
				}
				int deb = 1;
			}
			else if (nResponse == IDYES)
			{
				HWND h = GetSafeHwnd();
				HINSTANCE result = ShellExecute(h, _T("open"), pdfFileName, NULL, NULL, SW_SHOWNORMAL);
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
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}
	return 1;
}

int NListView::PrintMailSelectedToSeparatePDF_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText)
{
	BOOL progressBar = TRUE;

	int retdel = NListView::DeleteAllHtmAndPDFFiles(targetPrintFolderPath);

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (selectedMailIndexList->GetCount() <= 0)
		return 1;

	if (progressBar && MboxMail::pCUPDUPData)
	{
		MboxMail::pCUPDUPData->SetProgress(_T("Printing mails to PDF files ..."), 1);
	}

	double delta = (double)(selectedMailIndexList->GetCount());
	if (delta <= 0) delta = 1;
	double step = delta / 100;
	double curstep = 1;
	double newstep = 0;
	CString fileNum;
	int nFileNum;

	if (progressBar && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	int i;
	int cnt = selectedMailIndexList->GetCount();
	BOOL progress_Bar = FALSE;
	for (int j = 0; j < cnt; j++)
	{
		i = (*selectedMailIndexList)[j];
		if (pFrame) {
			pFrame->PrintSingleMailtoPDF(i, targetPrintSubFolderName, progress_Bar, errorText);
		}

		if (progressBar && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				break;
			}

			newstep = ((double)(j - 0 + 1)) / step;
			if (newstep >= curstep)
			{
				curstep = newstep;
			}

			nFileNum = (j + 1);
			fileNum.Format(_T("Printing mails to PDF files ... %d of %d"), nFileNum, cnt);
			if (MboxMail::pCUPDUPData)  MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(curstep));
		}
	}
	nFileNum = cnt;
	newstep = 100;
	fileNum.Format(_T("Printing mails to PDF files ... %d of %d"), nFileNum, cnt);
	if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(newstep));
	int deb = 1;
	return 1;
}

int NListView::PrintMailSelectedToSinglePDF_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText)
{
	BOOL progressBar = TRUE;

	//int retdel = NListView::DeleteAllHtmAndPDFFiles(targetPrintFolderPath);

	if (selectedMailIndexList->GetCount() <= 0)
		return 1;

	PrintMailSelectedToSingleHTML_WorkerThread(selectedMailIndexList, targetPrintSubFolderName, targetPrintFolderPath, errorText);

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	double delta = (double)(selectedMailIndexList->GetCount());
	CString progressText;
	if (progressBar && MboxMail::pCUPDUPData)
	{
		progressText.Format(_T("Printing %d mails to single PDF file ..."), (int)delta);
		MboxMail::pCUPDUPData->SetProgress(progressText, 1);
	}

	if (delta <= 0) delta = 1;
	double step = delta / 100;
	double curstep = 1;

	if (progressBar && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	int textType = 1;
	CString targetPrintSubFolder;
	bool fileExists;
	CString htmFileName;
	int ret = MboxMail::MakeFileNameFromMailArchiveName(textType, htmFileName, targetPrintSubFolder, fileExists, errorText);

	ret = CMainFrame::ExecCommand_WorkerThread(htmFileName, errorText, TRUE, progressText);
	return 1;
}

void NListView::FindFirstAndLastMailOfConversation(int iItem, int &firstMail, int &lastMail)

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
	if (m_selectedMailsList.GetCount() > 0)
		return &m_selectedMailsList;
	else
		return 0;
}

#if 0
BOOL NListView::PreTranslateMessage(MSG* pMsg)
{
	// TODO: Add your specialized code here and/or call the base class

	//if (pMsg->message == WM_KEYDOWN)
	{
		SetFocus();
	}

	return CWnd::PreTranslateMessage(pMsg);
}

void NListView::OnSetFocus(CWnd* pOldWnd)
{
	CWnd::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here

	SetFocus();
}


void NListView::OnMouseHover(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default

	CWnd::OnMouseHover(nFlags, point);
}
#endif

void NListView::SetSetFocus()
{
	SetFocus();
	m_list.SetFocus();
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
		MboxMail::pCUPDUPData->SetProgress(_T("Deleting HTM and PDF files in target print location ..."), 0);
	}

	CString fw = targetPath + "\\*.htm";
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
			//TRACE(_T("HTM file=%s\n"), fn);
			filePath = targetPath + "\\" + fn;
			delStatus = DeleteFile(filePath);
			if (delStatus == FALSE) {
				DWORD error = GetLastError();
			}

		} while (FindNextFile(f, &wf));
		FindClose(f);
	}

	fw = targetPath + "\\*.pdf";
	f = FindFirstFile(fw, &wf);
	if (f != INVALID_HANDLE_VALUE)
	{
		do
		{
			if ((wf.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == FILE_ATTRIBUTE_DIRECTORY)
				continue;
			fileCnt++;
			fn = wf.cFileName;
			//TRACE(_T("PDF file=%s\n"), fn);
			filePath = targetPath + "\\" + fn;
			delStatus = DeleteFile(filePath);
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


#if 1
int NListView::PrintMailArchiveToSeparateHTML_Thread(CString &errorText)
{
	int firstMail = 0;
	int lastMail = MboxMail::s_mails.GetCount() - 1;

	CString targetPrintSubFolderName = "HTML_GROUP";
	int ret = PrintMailRangeToSeparateHTML_Thread(firstMail, lastMail, targetPrintSubFolderName);

	return ret;
}

int NListView::PrintMailConversationToSeparateHTML_Thread(int iItem, CString &errorText)
{
	int firstMail = 0;
	int lastMail = 0;

	FindFirstAndLastMailOfConversation(iItem, firstMail, lastMail);

	CString targetPrintSubFolderName = "HTML_GROUP";
	int ret = PrintMailRangeToSeparateHTML_Thread(firstMail, lastMail, targetPrintSubFolderName);

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
	PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS args;

	CString rootPrintSubFolder = "PrintCache";
	CString printCachePath;
	CString errorText;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolderName, printCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	/*IN*/
	args.lview = this;
	args.targetPrintFolderPath = printCachePath;
	args.targetPrintSubFolderName = targetPrintSubFolderName;
	args.firstMail = firstMail;
	args.lastMail = lastMail;
	args.selectedMailIndexList = 0;
	args.nItem = -1;
	args.separatePDFs = TRUE;

	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcPrintMailGroupToSeparateHTML, (LPVOID)(PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS*)&args);

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
		//ASSERT(cancelledbyUser == TRUE);
		int loopCnt = 20;
		DWORD tc_start = GetTickCount();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		DWORD tc_end = GetTickCount();
		DWORD delta = tc_end - tc_start;
		TRACE("(PrintMailRangeToSeparateHTML_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret > 0)
	{
		if (args.errorText.IsEmpty())  // TODO: should always be true
		{
			CString txt = "Created separate HTML files in \n\n" + printCachePath + " \n\ndirectory.";
			OpenContainingFolderDlg dlg(txt, TRUE);
			INT_PTR nResponse = dlg.DoModal();
			if (nResponse == IDOK)
			{
				HWND h = GetSafeHwnd();
				HINSTANCE result = ShellExecute(h, _T("open"), printCachePath, NULL, NULL, SW_SHOWNORMAL);
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
			MboxMail::assert_unexpected();
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}
	return 1;
}

int NListView::PrintMailRangeToSingleHTML_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS args;

	CString rootPrintSubFolder = "PrintCache";
	CString printCachePath;
	CString errorText;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolderName, printCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	/*IN*/
	args.lview = this;
	args.targetPrintFolderPath = printCachePath;
	args.targetPrintSubFolderName = targetPrintSubFolderName;
	args.firstMail = firstMail;
	args.lastMail = lastMail;
	args.selectedMailIndexList = 0;
	args.nItem = -1;
	args.separatePDFs = FALSE;

	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcPrintMailGroupToSingleHTML, (LPVOID)(PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS*)&args);

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
		//ASSERT(cancelledbyUser == TRUE);
		int loopCnt = 20;
		DWORD tc_start = GetTickCount();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		DWORD tc_end = GetTickCount();
		DWORD delta = tc_end - tc_start;
		TRACE("(PrintMailRangeToSingleHTML_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret > 0)
	{
		if (args.errorText.IsEmpty())
		{
			int fileType = 1;
			bool fileExists;
			CString fileName;

			if (firstMail != lastMail)
				int ret = MboxMail::MakeFileNameFromMailArchiveName(fileType, fileName, targetPrintSubFolderName, fileExists, errorText);
			else
				int ret = MboxMail::MakeFileNameFromMailHeader(firstMail, fileType, fileName, targetPrintSubFolderName, fileExists, errorText);

			CString txt = "Created HTML file \n\n" + fileName;
			OpenContainingFolderDlg dlg(txt, FALSE);
			INT_PTR nResponse = dlg.DoModal();
			////////////
			if (nResponse == IDOK)
			{
				if (FileUtils::BrowseToFile(fileName) == FALSE) {
					HWND h = GetSafeHwnd();
					HINSTANCE result = ShellExecute(h, _T("open"), printCachePath, NULL, NULL, SW_SHOWNORMAL);
					CMainFrame::CheckShellExecuteResult(result, h);
				}
				int deb = 1;
			}
			else if (nResponse == IDYES)
			{
				HWND h = GetSafeHwnd();
				HINSTANCE result = ShellExecute(h, _T("open"), fileName, NULL, NULL, SW_SHOWNORMAL);
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
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}

	return 1;
}

int NListView::PrintMailRangeToSeparateHTML_WorkerThread(int firstMail, int lastMail, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText)
{
	BOOL progressBar = TRUE;

	int retdel = NListView::DeleteAllHtmAndPDFFiles(targetPrintFolderPath);

	if (progressBar && MboxMail::pCUPDUPData)
	{
		MboxMail::pCUPDUPData->SetProgress(_T("Printing mails to HTML files ..."), 0);
	}

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	double delta = (double)(lastMail - firstMail + 1);
	if (delta <= 0) delta = 1;
	double step = delta / 100;
	double curstep = 1;
	double newstep = 0;
	CString fileNum;
	int nFileNum = 0;

	if (progressBar && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	int mailCnt = lastMail - firstMail + 1;
	for (int i = firstMail; i <= lastMail; i++)
	{
		if (pFrame)
		{
			// PrintSingleMailtoPDF will make a unique name and print HTML to PDF
			pFrame->PrintSingleMailtoHTML(i, targetPrintSubFolderName, errorText);
		}

		if (progressBar && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				int deb = 1;
				break;
			}

			newstep = ((double)(i - firstMail + 1)) / step;
			if (newstep >= curstep)
			{
				if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress((UINT_PTR)(newstep));
				curstep = newstep;
			}

#if 1
			nFileNum = ((i - firstMail) + 1);
			if ((nFileNum % 100) == 0) {
				fileNum.Format(_T("Printing mails to HTML files ... %d of %d"), nFileNum, mailCnt);
				if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(curstep));
			}
#endif


		}
	}
	
	newstep = 100;
	if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress((UINT_PTR)(newstep));
#if 1
	nFileNum = mailCnt;
	fileNum.Format(_T("Printing mails to HTML files ... %d of %d"), nFileNum, mailCnt);
	if (MboxMail::pCUPDUPData)  MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(newstep));
#endif
	return  1;
}

int NListView::PrintMailRangeToSingleHTML_WorkerThread(int firstMail, int lastMail, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText)
{
	//BOOL progressBar = TRUE;

	TEXTFILE_CONFIG textConfig;
	CString htmFileName;
	int textType = 1;

	textConfig.m_dateFormat = m_format;
	textConfig.m_bGMTTime = m_gmtTime;
	textConfig.m_nCodePageId = CP_UTF8;

	int ret = MboxMail::PrintMailRangeToSingleTextFile_WorkerThread(textConfig, htmFileName, firstMail, lastMail, textType, errorText);

	return  1;
}

int NListView::PrintMailSelectedToSeparateHTML_Thread(CString &targetPrintSubFolderName, CString &targetPrintFolderPath)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS args;

	CString rootPrintSubFolder = "PrintCache";
	CString targetPrintSubFolder = targetPrintSubFolderName;
	CString printCachePath;
	CString errorText;

	MailIndexList *selectedMailsIndexList = 0;
	if (PopulateSelectedMailsList() > 0)
		selectedMailsIndexList = &m_selectedMailsList;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, printCachePath, errorText);
	if (retval == FALSE) {
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
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

	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcPrintMailGroupToSeparateHTML, (LPVOID)(PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS*)&args);

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
		//ASSERT(cancelledbyUser == TRUE);
		int loopCnt = 20;
		DWORD tc_start = GetTickCount();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		DWORD tc_end = GetTickCount();
		DWORD delta = tc_end - tc_start;
		TRACE("(PrintMailSelectedToSeparateHTML_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret > 0)
	{
		if (args.errorText.IsEmpty())
		{

			CString txt = "Created separate HTML files in \n\n" + printCachePath + " \n\ndirectory.";
			OpenContainingFolderDlg dlg(txt, TRUE);
			INT_PTR nResponse = dlg.DoModal();
			if (nResponse == IDOK)
			{
				HWND h = GetSafeHwnd();
				HINSTANCE result = ShellExecute(h, _T("open"), printCachePath, NULL, NULL, SW_SHOWNORMAL);
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
			MboxMail::assert_unexpected();
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}
	return 1;
}

int NListView::PrintMailSelectedToSingleHTML_Thread(CString &targetPrintSubFolderName, CString &targetPrintFolderPath)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS args;

	CString rootPrintSubFolder = "PrintCache";
	CString targetPrintSubFolder = targetPrintSubFolderName;
	CString printCachePath;
	CString errorText;

	MailIndexList *selectedMailsIndexList = 0;
	if (PopulateSelectedMailsList() > 0)
		selectedMailsIndexList = &m_selectedMailsList;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, printCachePath, errorText);
	if (retval == FALSE) {
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
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
		//ASSERT(cancelledbyUser == TRUE);
		int loopCnt = 20;
		DWORD tc_start = GetTickCount();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		DWORD tc_end = GetTickCount();
		DWORD delta = tc_end - tc_start;
		TRACE("(PrintMailSelectedToSingleHTML_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret > 0)
	{
		if (args.errorText.IsEmpty())
		{
			int textType = 1;
			bool fileExists;
			CString fileName;

			int ret = MboxMail::MakeFileNameFromMailArchiveName(textType, fileName, targetPrintSubFolder, fileExists, errorText);

			CString txt = "Created HTML file \n\n" + fileName;
			OpenContainingFolderDlg dlg(txt, FALSE);
			INT_PTR nResponse = dlg.DoModal();
			////////////
			if (nResponse == IDOK)
			{
				if (FileUtils::BrowseToFile(fileName) == FALSE) {
					HWND h = GetSafeHwnd();
					HINSTANCE result = ShellExecute(h, _T("open"), printCachePath, NULL, NULL, SW_SHOWNORMAL);
					CMainFrame::CheckShellExecuteResult(result, h);
				}
				int deb = 1;
			}
			else if (nResponse == IDYES)
			{
				HWND h = GetSafeHwnd();
				HINSTANCE result = ShellExecute(h, _T("open"), fileName, NULL, NULL, SW_SHOWNORMAL);
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
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
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
		MboxMail::pCUPDUPData->SetProgress(_T("Printing mails to HTML files ..."), 1);
	}

	double delta = (double)(selectedMailIndexList->GetCount());
	if (delta <= 0) delta = 1;
	double step = delta / 100;
	double curstep = 1;
	double newstep = 0;
	CString fileNum;
	int nFileNum = 0;

	if (progressBar && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	int i;
	int retval = 1;
	int cnt = selectedMailIndexList->GetCount();
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

			newstep = ((double)(j - 0 + 1)) / step;
			if (newstep >= curstep)
			{
				if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress((UINT_PTR)(newstep));
				curstep = newstep;
			}
#if 1
			nFileNum = j + 1;
			if ((nFileNum % 100) == 0) {
				fileNum.Format(_T("Printing mails to HTML files ... %d of %d"), nFileNum, cnt);
				if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(curstep));
			}
#endif
		}
	}
	
	newstep = 100;
	if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress((UINT_PTR)(newstep));
#if 1
	nFileNum = cnt;
	fileNum.Format(_T("Printing mails to HTML files ... %d of %d"), nFileNum, cnt);
	if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(newstep));
#endif

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
	if (abs(MboxMail::b_mails_which_sorted) != 99) {

		CString txt = _T("Please sort all mails by conversation first.\n");
		txt += "Select \"View\"->\"Sort By\" ->\"Conversation\" or left click on the first column.";
		HWND h = GetSafeHwnd(); // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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

#endif

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

#define MAIL_LIST_VERSION2  (MAIL_LIST_VERSION_BASE+2)

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
			CString txt = _T("User Selected Mails List is not empty. Overwrite?");
			int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
			if (answer == IDNO)
				return -1;
			else if (answer == IDYES)
				MboxMail::m_editMails.m_bIsDirty = FALSE;
		}
	}

	CString mboxFileSuffix = "_USER.mbox";

	CString mailFile = MboxMail::s_path;

	if (MboxMail::s_path.IsEmpty()) {
		CString txt = _T("Please open mail file first.");
		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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
	CString rootPrintSubFolder = "ArchiveCache";
	CString targetPrintSubFolder;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, archiveCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CString mboxFile = fileNameBase + fileNameExtention;  // main mbox file name
	CString mboxFilePath = archiveCachePath + "\\" + fileNameBase + mboxFileSuffix;  // derived _USER.mbox mbox file name

	CString mboxFileListSuffix = ".mboxlist";
	CString mboxListFile;
	if (mbxListFile)
		mboxListFile.Append(*mbxListFile);
	else
		mboxListFile = mboxFilePath + mboxFileListSuffix;

	if (!FileUtils::PathFileExist(mailFile))
	{
		CString txt = _T("Mail Archive File \"") + mailFile;
		txt += _T("\" doesn't exist.\nCan't reload.");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	if (!FileUtils::PathFileExist(mboxListFile))
	{
		CString txt = _T("Mail List File \"") + mboxListFile;
		txt += _T("\" doesn't exist.\nCan't reload.");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	SerializerHelper sz(mboxListFile);
	if (!sz.open(FALSE)) {
		CString txt = _T("Could not open \"") + mboxListFile;
		txt += _T("\" file.\nMake sure file is not open on other applications.");
		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	int version;
	_int64 mailFileSize;
	_int64 mboxFileSize;
	int mailListCnt;

	CString txt = _T("Mail list file\n\"") + mboxListFile;
	txt += _T("\"\nhas incompatible version or file is corrupted.\nCan't reload.\nPlease remove the file.");

	__int64 pos = 0;
	retval = sz.GetReadPointer(pos);

	if (!sz.readInt(&version)) {
		sz.close();
		return -1;
	}

	pos = 0;
	retval = sz.GetReadPointer(pos);

	if (version != MAIL_LIST_VERSION2)
	{
		sz.close();

		CString text = _T("Mail list file\n\"") + mboxListFile;
		CString strVersion;
		strVersion.Format(_T("%d"), (version - MAIL_LIST_VERSION_BASE));
		text += _T("\".\nhas incompatible version\"") + strVersion + "\". Expected version \"";
		strVersion.Format(_T("%d"), (MAIL_LIST_VERSION - MAIL_LIST_VERSION_BASE));
		text += strVersion + "\".\nCan't reload.\nPlease remove the file.";

		int answer = MessageBox(text, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	if (!sz.readInt64(&mboxFileSize)) {  // derived mbox mail file
		sz.close();
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	if (!sz.readInt(&mailListCnt)) {
		sz.close();
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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
		CString txt = _T("Mail list file\n\"") + mboxListFile;
		txt += _T("\"\n is invalid or corrupted. Can't reload.\nPlease remove the file.");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);

		return -1; // do nothing
	}

	sz.close();

	if (ret < 0)
	{
		MboxMail::s_mails_edit.SetSizeKeepData(0);

		CString txt = _T("Mail list file\n\"") + mboxListFile;
		txt += _T("\"\n is invalid or corrupted. Can't reload.\nPlease remove the file.");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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
		CString txt = _T("Please open mail file first.");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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
	CString rootPrintSubFolder = "ArchiveCache";
	CString targetPrintSubFolder;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, archiveCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CFile fp;
	CString mboxFile = fileNameBase + fileNameExtention;
	CString mboxFilePath = archiveCachePath + "\\" + fileNameBase + mboxFileSuffix;

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

	SerializerHelper sz(mboxListFile);
	if (!sz.open(TRUE)) 
	{
		CString txt = _T("Could not create \"") + mboxListFile;
		txt += _T("\" file.\nMake sure file is not open on other applications.");
		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	// Create mboxlist file to allow reload of archive file list

	sz.writeInt(MAIL_LIST_VERSION2);			// version
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
		sz.writeInt(m->m_ContentDetailsArray.size());
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
		CString txt = _T("Please open mail file first.");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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
	CString rootPrintSubFolder = "ArchiveCache";
	CString targetPrintSubFolder;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, archiveCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CFile fp;

	CString mboxFile = fileNameBase + fileNameExtention;
	CString archiveFile = fileNameBase + mboxFileSuffix;;
	CString mboxFilePath = archiveCachePath + "\\" + fileNameBase + mboxFileSuffix;

	if (FileUtils::PathFileExist(mboxFilePath))
	{
		CString txt = _T("File \"") + mboxFilePath;
		txt += _T("\" exists.\nOverwrite?");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
			return -1;
	}

	if (!fp.Open(mboxFilePath, CFile::modeWrite | CFile::modeCreate)) {
		CString txt = _T("Could not create \"") + mboxFilePath;
		txt += _T("\" file.\nMake sure file is not open on other applications.");
		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	SimpleString *outbuf = MboxMail::m_outbuf;
	outbuf->ClearAndResize(10000);

	BOOL ret;
	for (int i = 0; i < mailsArray->GetSize(); i++)
	{
		m = (*mailsArray)[i];
		outbuf->Clear();
		ret = m->GetBody(outbuf);
		fp.Write(outbuf->Data(), outbuf->Count());
	}
	fp.Close();

	MboxMail::m_editMails.m_bIsDirty = FALSE;

	CString txt = "Created Mbox Mail Archive file \n\n" + mboxFilePath;
	BOOL supressOpenFileOption = FALSE;
	if (path.Compare(datapath))
		supressOpenFileOption = TRUE;
	OpenContainingFolderDlg dlg(txt, supressOpenFileOption);
	INT_PTR nResponse = dlg.DoModal();
	////////////
	if (nResponse == IDOK)
	{
		if (FileUtils::BrowseToFile(mboxFilePath) == FALSE) {
			HWND h = GetSafeHwnd();
			HINSTANCE result = ShellExecute(h, _T("open"), archiveCachePath, NULL, NULL, SW_SHOWNORMAL);
			CMainFrame::CheckShellExecuteResult(result, h);
		}
		int deb = 1;
	}
	else if (nResponse == IDYES)
	{
		CString txt = "Open Created Archive File \n\n" + mboxFilePath;
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

				BOOL retMove = MoveFile(mboxFilePath, archiveFilePath);

				//pFrame->DoOpen(archiveFilePath);

				NTreeView *pTreeView = pFrame->GetTreeView();
				if (pTreeView)
				{
					// delete index file to make sure it is not used i ncase old and new length of new mbox file are the same
					// InsertMailFile will delete index file
					//CString indexFile = archiveFilePath + ".mboxview";
					//DeleteFile((LPCSTR)indexFile);

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

BOOL IsSupportedPictureFileExtension(CString &cext)
{
	if ((cext.CompareNoCase(".png") == 0) ||
		(cext.CompareNoCase(".jpg") == 0) ||
		(cext.CompareNoCase(".gif") == 0) ||
		(cext.CompareNoCase(".pjpg") == 0) ||
		(cext.CompareNoCase(".jpeg") == 0) ||
		(cext.CompareNoCase(".pjpeg") == 0) ||
		(cext.CompareNoCase(".jpe") == 0) ||
		(cext.CompareNoCase(".bmp") == 0) ||
		(cext.CompareNoCase(".tif") == 0) ||
		(cext.CompareNoCase(".tiff") == 0) ||
		(cext.CompareNoCase(".dib") == 0) ||
		(cext.CompareNoCase(".jfif") == 0) ||
		(cext.CompareNoCase(".emf") == 0) ||
		(cext.CompareNoCase(".wmf") == 0) ||
		(cext.CompareNoCase(".ico") == 0))
	{
		return TRUE;
	}
	else
		return FALSE;
}


int NListView::FindFilenameCount(std::vector <MailBodyContent*> &contentDetailsArray, CString &fileName)
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

int NListView::DetermineImageFileName(MboxMail *m, CString &cidName, CString &imageFilePath, MailBodyContent **foundBody, int mailPosition)
{
	if (m == 0)
	{
		// TODO: Assert ??
		return 1;
	}

	SimpleString cid;
	//CString imageCachePath;
	CString imageFileName;
	CString nameExtension;
	CString idExtension;
	CString locationExtension;
	MailBodyContent *body;

	CString contentTypeExtension;
	CString contentTypeMain;

	BOOL isOctetStream;

	*foundBody = 0;

	CString mailIndex;
	mailIndex.Format("%07d_", m->m_index);

	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];

		contentTypeExtension.Empty();
		contentTypeMain.Empty();
		nameExtension.Empty();
		idExtension.Empty();
		locationExtension.Empty();


		BOOL isValidContentTypeExtension = FALSE;
		int pos = body->m_contentType.ReverseFind('/');
		if (pos > 0)
		{
			contentTypeExtension = body->m_contentType.Mid(pos + 1);
			contentTypeMain = body->m_contentType.Left(pos);
			CString contentTypeExt = "." + contentTypeExtension;
			isValidContentTypeExtension = IsSupportedPictureFileExtension(contentTypeExt);
		}

		if (contentTypeMain.CompareNoCase("message") == 0)
		{
			TRACE(_T("Content-Type=message index=%d\n"), mailPosition);
			int deb = 1;
		}

		isOctetStream = FALSE;
		if (contentTypeExtension.CompareNoCase("octet-stream") == 0)
			isOctetStream = TRUE;
		
		if ((contentTypeMain.CompareNoCase("image") == 0) || 
			(body->m_contentDisposition.CompareNoCase("inline") == 0) || isOctetStream)
		{
			BOOL isValidAttachmentNameExtension = FALSE;
			int pos = body->m_attachmentName.ReverseFind('.');
			if (pos >= 0)
			{
				nameExtension = body->m_attachmentName.Mid(pos);
				isValidAttachmentNameExtension = IsSupportedPictureFileExtension(nameExtension);
			}

			BOOL isValidContentIdExtension = FALSE;
			pos = body->m_contentId.ReverseFind('.');
			if (pos >= 0)
			{
				idExtension = body->m_contentId.Mid(pos);
				isValidContentIdExtension = IsSupportedPictureFileExtension(idExtension);
			}

			BOOL isValidContentLocationExtension = FALSE;
			pos = body->m_contentLocation.ReverseFind('.');
			if (pos >= 0)
			{
				locationExtension = body->m_contentLocation.Mid(pos);
				isValidContentLocationExtension = IsSupportedPictureFileExtension(idExtension);
			}

			CString attachmentName = body->m_attachmentName;

			if (!body->m_contentId.IsEmpty())
			{
				CString contentId = body->m_contentId;
				if (contentId.Compare(cidName) == 0)
				{
					// TODO: we assume all contentIds are unique which may not be always true
					int fileCnt = 0;
					if (!attachmentName.IsEmpty())
						fileCnt = FindFilenameCount(m->m_ContentDetailsArray, attachmentName);

					if (!attachmentName.IsEmpty() && (fileCnt < 2))
					{
						if (isValidAttachmentNameExtension)
							imageFileName = mailIndex + attachmentName;
						else if (isValidContentTypeExtension)
							imageFileName = mailIndex + attachmentName + "." + contentTypeExtension;
						else if (isValidContentIdExtension)
							imageFileName = mailIndex + attachmentName + "." + idExtension;
						else if (isValidContentLocationExtension)
							imageFileName = mailIndex + attachmentName + "." + locationExtension;
						else
						{
							// Should try to determine if that is an image and what type
							// Would need to LoadImage from file and check
							// Add .jpg for now; don't tink browsers care about suffix
							imageFileName = mailIndex + attachmentName + ".jpg";
						}
					}
					else if (isValidContentIdExtension)
					{
						imageFileName = mailIndex + contentId;
					}
					else if (isValidAttachmentNameExtension)
					{
						imageFileName = mailIndex + contentId + "." + nameExtension;
					}
					else if (isValidContentTypeExtension)
					{
						imageFileName = mailIndex + contentId + "." + contentTypeExtension;
					}
					else
					{
						imageFileName = mailIndex + contentId + ".jpg";
					}
					*foundBody = body;
					break;
				}
				else
				{
					int deb = 1;
				}
			}

			if ((*foundBody == 0) && !body->m_contentLocation.IsEmpty())
			{
				CString contentLocation = body->m_contentLocation;
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
							imageFileName = mailIndex + attachmentName + "." + contentTypeExtension;
						else if (isValidContentIdExtension)
							imageFileName = mailIndex + attachmentName + "." + idExtension;
						else if (isValidContentLocationExtension)
							imageFileName = mailIndex + attachmentName + "." + locationExtension;
						else
						{
							// Should try to determine if that is an image and what type
							// Would need to LoadImage from file and check
							// Add .jpg for now; don't tink browsers care about suffix
							imageFileName = mailIndex + attachmentName + ".jpg";
						}
					}
					else if (isValidContentLocationExtension)
					{
						imageFileName = mailIndex + contentLocation;
					}
					else if (isValidAttachmentNameExtension)
					{
						imageFileName = mailIndex + contentLocation + "." + nameExtension;
					}
					else if (isValidContentTypeExtension)
					{
						imageFileName = mailIndex + contentLocation + "." + contentTypeExtension;
					}
					else
					{
						// Should try to detrmine if that is an image and what type
						// Would need to LoadImage from file and check
						// Add .jpg for now; don't tink browsers care about suffix
						imageFileName = mailIndex + contentLocation + ".jpg";
					}
					*foundBody = body;
					break;
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
						imageFileName = mailIndex + attachmentName + "." + contentTypeExtension;
					else if (isValidContentIdExtension)
						imageFileName = mailIndex + attachmentName + "." + idExtension;
					else if (isValidContentLocationExtension)
						imageFileName = mailIndex + attachmentName + "." + locationExtension;
					else
					{
						imageFileName = mailIndex + attachmentName + ".jpg";
					}
					*foundBody = body;
					break;
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
	}
	FileUtils::MakeValidFileName(imageFileName);
	imageFilePath = imageFileName;

#if 0
	TRACE(_T("--->>> SUBJ=%s\n"), m->m_subj);
	if (imageFileName.IsEmpty())
	{
		TRACE(_T("\tCID=%s FNAME=%s\n"), cidName, imageFilePath);
		int deb = 1;
	}
	else
	{
		TRACE(_T("\tCID=%s FNAME=%s\n"), cidName, imageFilePath);
		int deb = 1;
	}
#endif

	return 1;
}

int NListView::CreateInlineImageFiles(CFile &fpm, int mailPosition, CString &imageCachePath, bool runInvestigation)
{
	static char * img_pattern = "<img";
	static int img_patternLen = strlen(img_pattern);

	static char * src_pattern = "src=";
	static int src_patternLen = strlen(src_pattern);

	static char * cid_pattern = "cid:";
	static int cid_patternLen = strlen(cid_pattern);

	static char * http_pattern = "http:";
	static int http_patternLen = strlen(http_pattern);

	static char * https_pattern = "https:";
	static int https_patternLen = strlen(https_pattern);

	static char * mhtml_pattern = "mhtml:";
	static int mhtml_patternLen = strlen(mhtml_pattern);

	static char * data_pattern = "data:";
	static int data_patternLen = strlen(data_pattern);

	MboxMail *m = 0;
	if ((mailPosition < MboxMail::s_mails.GetCount()) && (mailPosition >= 0))
		m = MboxMail::s_mails[mailPosition];
	else
		return -1;

	CString contentTypeMain;
	MailBodyContent *body;
	bool foundImage = false;
	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];

		CString contentTypeExtension;
		CString contentTypeMain;
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

	// Enable if needed
	// CArray<CString*, CString*> cidArray;

	char *cidBegin;
	char *cidEnd;
	SimpleString cid;
	CString imageFileName;
	CString imageFilePath;

	// MboxMail::m_outbuf is a global buffer; make sure is not in use already  !!!!
	SimpleString *outbuflarge = MboxMail::m_outbuf;

	outbuflarge->Clear();
	UINT pageCode = 0;
	int textType = 1; // try first Html
	int textlen = MboxMail::GetMailBody_mboxview(fpm, mailPosition, outbuflarge, pageCode, textType);  // returns pageCode
	if (textlen != outbuflarge->Count())
		int deb = 1;

	char *input = outbuflarge->Data();
	int inputLength = outbuflarge->Count();
	char *inputEnd = input + inputLength;
	char *srcImgEnd = inputEnd;

	int mIndex = 0;

	CString mailIndex;
	mailIndex.Format("%07d_", m->m_index);

	int cidCnt = 0;
	char *pos = input;
	while ((pos != 0) && (pos < inputEnd))
	{
		imageFileName.Empty();

		// Best would be to join all lines between <img ... and > and then process

		pos = TextUtilsEx::strnstrUpper2Lower(pos, inputEnd, img_pattern, img_patternLen);
		if (pos == 0) {
			break;
		}

		pos += img_patternLen;

		if ((*pos == ' ') || (*pos == '\n') || (*pos == '\r'))
			pos += 1;
		else
			continue;

		// find end of <img src ...>
		srcImgEnd = TextUtilsEx::strnstrUpper2Lower(pos, inputEnd, ">", 1);
		if (srcImgEnd == 0) { // TODO: corrupted file ?
			break;
		}
		srcImgEnd++; // jump over

		pos = TextUtilsEx::strnstrUpper2Lower(pos, srcImgEnd, src_pattern, src_patternLen);
		if (pos == 0) {
			pos = srcImgEnd;  // jump over this <img .. > sonce no cid:
			continue;
		}

		bool srcCid = false;
		bool srcHttp = false;
		bool srcHttps = false;
		bool srcMHtml = false;  // foung mhtml
		bool srcMHtmlHtml = false;  // found mhtml and http or https
		bool srcMHtmlNoHtml = false; // did not found mhtml and http or https
		bool srcMHtmlHttp = false;  // found mhtml and http 
		bool srcMHtmlHttps = false;  // found mhtml and https
		bool srcData = false;
		bool srcLocalFile = false;

		pos += src_patternLen;
		BOOL foundHTTP = FALSE;
		cidBegin = pos;

		BOOL inQuotes = FALSE;
		if (*pos == '"')
		{
			inQuotes = TRUE;
			pos++;
		}

		pos = TextUtilsEx::strnstrUpper2Lower(pos, pos + cid_patternLen, cid_pattern, cid_patternLen);
		if (pos != 0)
		{
			srcCid = true;
			pos += cid_patternLen;
			cidBegin = pos;
		}

		if (pos == 0)
		{
			pos = cidBegin;
			pos = TextUtilsEx::strnstrUpper2Lower(pos, pos + mhtml_patternLen, mhtml_pattern, mhtml_patternLen);
			if (pos) {
				srcMHtml = true;
				MboxMail::m_EmbededImagesFoundMHtml++;
				pos = TextUtilsEx::strchar(pos, srcImgEnd, '!');
				if (pos) {
					cidBegin = pos + 1;
					pos = 0;
					foundHTTP = true;
					srcMHtmlHtml = true;
					MboxMail::m_EmbededImagesFoundMHtmlHtml++;
				}
				else
				{
					MboxMail::m_EmbededImagesFoundUnexpectedMHtml++;
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
				pos += 0;  // include http:
				cidBegin = pos;
				foundHTTP = true;
				srcHttp = true;
			}
		}
		if (pos == 0)
		{
			pos = cidBegin;
			pos = TextUtilsEx::strnstrUpper2Lower(pos, pos + https_patternLen, https_pattern, https_patternLen);
			if (pos) { 
				pos += 0;  // include https:
				cidBegin = pos;
				foundHTTP = true;
				srcHttps = true;
			}
		}

		if (pos == 0)
		{
			pos = cidBegin;
			pos = TextUtilsEx::strnstrUpper2Lower(pos, pos + data_patternLen, data_pattern, data_patternLen);
			if (pos) { // not implemented ; implemented by broesers, etc
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
				MboxMail::m_EmbededImagesNoMatch++;
				pos = srcImgEnd;  // jump over this <img .. > since no image
				continue;
			}
			cidEnd = pos;
			srcLocalFile = true;
		}
		else
		{
			// find end of cid token
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
		}

		cid.Copy(cidBegin, cidEnd - cidBegin);
		if (foundHTTP) {
			int retLen = TextUtilsEx::DecodeURL(cid.Data(), cid.Count());
			cid.SetCount(retLen);
		}

		CString cidName;
		cidName.Insert(0, cid.Data());

		MailBodyContent *pBodyFound;
		int retval = NListView::DetermineImageFileName(m, cidName, imageFileName, &pBodyFound, mailPosition);
		if (pBodyFound && (retval > 0))
		{
			MboxMail::m_EmbededImagesFound++;
			if (srcCid)
			{
				MboxMail::m_EmbededImagesFoundCid++;
			}
			else if (srcHttp)
			{
				if (srcMHtmlHtml)
					MboxMail::m_EmbededImagesFoundMHtmlHttp++;
				else
					MboxMail::m_EmbededImagesFoundHttp++;
			}
			else if (srcHttps)
			{
				if (srcMHtmlHtml)
					MboxMail::m_EmbededImagesFoundMHtmlHttps++;
				else
					MboxMail::m_EmbededImagesFoundHttps++;
			}
			else if (srcData)
			{
				MboxMail::m_EmbededImagesFoundData++;
			}
			else if (srcLocalFile)
			{
				MboxMail::m_EmbededImagesFoundLocalFile++;
			}

			if (!runInvestigation)
			{
				//imageFilePath = imageCachePath + mailIndex + imageFileName;
				imageFilePath = imageCachePath + imageFileName;
				CFileException ex;
				CFile fp;
				if (!fp.Open(imageFilePath, CFile::modeWrite | CFile::modeCreate, &ex))
				{
					TCHAR szError[1024];
					ex.GetErrorMessage(szError, 1024);
					CFileStatus rStatus;
					BOOL ret = fp.GetStatus(rStatus);
					CString errorText(szError);

					HWND h = NULL;
					// Ignore for now
					//int answer = ::MessageBox(h, errorText, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
				}
				else
				{
					SimpleString*outbuf = MboxMail::get_tmpbuf();
					outbuf->ClearAndResize(pBodyFound->m_contentLength * 2);
					int retLen = MboxMail::DecodeBody(fpm, pBodyFound, mailPosition, outbuf);
					if (outbuf->Count() > 1500000)
						int deb = 1;

					fp.Write(outbuf->Data(), outbuf->Count());
					fp.Close();
					MboxMail::rel_tmpbuf();
				}
			}
		}
		else
		{
			MboxMail::m_EmbededImagesNotFound++;

			if (srcCid)
			{
				MboxMail::m_EmbededImagesNotFoundCid++;
				TRACE(_T("CidNotFound indx=%d cid=%s \n"), mailPosition, cidName);
			}
			else if (srcHttp)
			{
				if (srcMHtmlHtml)
					MboxMail::m_EmbededImagesNotFoundMHtmlHttp++;
				else
					MboxMail::m_EmbededImagesNotFoundHttp++;
			}
			else if (srcHttps)
			{
				if (srcMHtmlHtml)
					MboxMail::m_EmbededImagesNotFoundMHtmlHttps++;
				else
					MboxMail::m_EmbededImagesNotFoundHttps++;
			}
			else if (srcData)
			{
				MboxMail::m_EmbededImagesNotFoundData++;
			}
			else if (srcLocalFile)
			{
				MboxMail::m_EmbededImagesNotFoundLocalFile++;
				TRACE(_T("LocalFileNotFound indx=%d cid=%s \n"), mailPosition, cidName);
			}
			else
				int deb = 1;

			int deb = 1;
		}

		pos++; // jump over \"
		pos = srcImgEnd;

		int deb = 1;
	}
#if 0
	TRACE(_T("Mail File=%s\n"), (LPCSTR)MboxMail::s_path);
	int i;
	for (i = 0; i < cidArray.GetCount(); i++)
	{
		TRACE(_T("CID=%s\n"), *cidArray[i]);
	}
#endif
	return 1;
}

int NListView::UpdateInlineSrcImgPath(char *inData, int indDataLen, SimpleString *outbuf, CListCtrl *attachments, int mailPosition, bool useMailPosition, bool runInvestigation)
{
	static char * img_pattern = "<img";
	static int img_patternLen = strlen(img_pattern);

	static char * src_pattern = "src=";
	static int src_patternLen = strlen(src_pattern);

	static char * cid_pattern = "cid:";
	static int cid_patternLen = strlen(cid_pattern);

	static char * http_pattern = "http:";
	static int http_patternLen = strlen(http_pattern);

	static char * https_pattern = "https:";
	static int https_patternLen = strlen(https_pattern);

	static char * mhtml_pattern = "mhtml:";
	static int mhtml_patternLen = strlen(mhtml_pattern);

	static char * data_pattern = "data:";
	static int data_patternLen = strlen(data_pattern);

	char *input = inData;
	int inputLength = indDataLen;
	char *inputEnd = input + inputLength;
	char *srcImgEnd = inputEnd;
	char *cidBegin;
	char *cidEnd;
	SimpleString cid;
	CString imageFileName;
	CString imageFilePath;

	char *fromBegin = input;
	char *alt_pos;
	int patternLen = cid_patternLen;
	int mhtmlIntroLen = 0;
	char *srcBegin;
	char *srcEnd;

	MboxMail *m = 0;
	if (useMailPosition)
	{
		if ((mailPosition < MboxMail::s_mails.GetCount()) && (mailPosition >= 0))
			m = MboxMail::s_mails[mailPosition];
		else
			return -1;
	}

	MailBodyContent *body;
	bool foundImage = false;
	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
	{
		body = m->m_ContentDetailsArray[j];

		CString contentTypeExtension;
		CString contentTypeMain;
		int pos = body->m_contentType.ReverseFind('/');
		if (pos > 0)
		{
			contentTypeExtension = body->m_contentType.Mid(pos + 1);
			contentTypeMain = body->m_contentType.Left(pos);
		}
		if (contentTypeMain.CompareNoCase("text") != 0) {
			foundImage = true;
			break;
		}
	}

	if (runInvestigation)
	{
		if (!foundImage)
			return 1;
	}


#if 0
	CString mailArchiveFileName;
	FileUtils::CPathStripPath((char*)(LPCSTR)MboxMail::s_path, mailArchiveFileName);
	int position = mailArchiveFileName.ReverseFind('.');
	CString baseFileArchiveName;
	if (position >= 0)
		baseFileArchiveName = mailArchiveFileName.Mid(0, position);
	else
		baseFileArchiveName = mailArchiveFileName;

	CString imageCachePath;
	BOOL ret = FileUtils::CPathGetPath(MboxMail::s_path, imageCachePath);
	imageCachePath.Append("\\");
	imageCachePath.Append("ImageCache");
	imageCachePath.Append("\\");
	imageCachePath.Append(baseFileArchiveName);
	imageCachePath.Append("\\");
#endif

	CString errorText;
	CString imageCachePath;
	CString rootPrintSubFolder = "ImageCache";
	CString targetPrintSubFolder;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, imageCachePath, errorText);
	imageCachePath.Append("\\");

	CString imgFile;
	CString cidName;

	CFile fpm;
	if (!fpm.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite)) {
		// TODO: ??
		//errorText = _T("Could not open mail archive \"") + MboxMail::s_path;
		fpm.Close();
		return -1;
	}

	if (!foundImage)
	{
		outbuf->Append(inData, indDataLen);
		fpm.Close();
		return 1;
	}

	char *pos = input;
	while ((pos != 0) && (pos < inputEnd))
	{
		pos = TextUtilsEx::strnstrUpper2Lower(pos, inputEnd, img_pattern, img_patternLen);
		if (pos == 0) {
			outbuf->Append(fromBegin, inputEnd - fromBegin);
			break;
		}
		pos += img_patternLen;

		if ((*pos == ' ') || (*pos == '\n') || (*pos == '\r'))
			pos += 1;
		else
			continue;

		// find end of <img src ...>
		srcImgEnd = TextUtilsEx::strnstrUpper2Lower(pos, inputEnd, ">", 1);
		if (srcImgEnd == 0) { // TODO: corrupted file ?
			outbuf->Append(fromBegin, inputEnd - fromBegin);
			break;
		}
		srcImgEnd++; // jump over

		alt_pos = TextUtilsEx::strnstrUpper2Lower(pos, srcImgEnd, "alt=\"", 5);

		pos = TextUtilsEx::strnstrUpper2Lower(pos, srcImgEnd, src_pattern, src_patternLen);
		if (pos == 0) {
			pos = srcImgEnd;  // jump over this <img .. > since no cid:
			continue;
		}

		bool srcCid = false;
		bool srcHttp = false;
		bool srcHttps = false;
		bool srcMHtml = false;  // foung mhtml
		bool srcMHtmlHtml = false;  // found mhtml and http or https
		bool srcMHtmlNoHtml = false; // did not found mhtml and http or https
		bool srcMHtmlHttp = false;  // found mhtml and http 
		bool srcMHtmlHttps = false;  // found mhtml and https
		bool srcData = false;
		bool srcLocalFile = false;

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
					mhtmlIntroLen = pos - cidBegin + 1;
					cidBegin = pos + 1;
					pos = 0;
					foundHTTP = TRUE;
					srcMHtmlHtml = true;
					MboxMail::m_EmbededImagesFoundMHtmlHtml++;
				}
				else
				{
					MboxMail::m_EmbededImagesFoundUnexpectedMHtml++;
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
				outbuf->Append(fromBegin, srcImgEnd - fromBegin);
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

		cid.Copy(cidBegin, cidEnd - cidBegin);
		if (foundHTTP) {
			int retLen = TextUtilsEx::DecodeURL(cid.Data(), cid.Count());
			cid.SetCount(retLen);
		}

		cidName.Empty();
		cidName.Append(cid.Data(), cid.Count());

		MailBodyContent *pBodyFound;
		int retval = NListView::DetermineImageFileName(m, cidName, imageFileName, &pBodyFound, mailPosition);
		imgFile = imageCachePath + imageFileName;

		if (pBodyFound && (retval > 0))
		{
			MboxMail::m_EmbededImagesFound++;

			if (srcCid)
			{
				MboxMail::m_EmbededImagesFoundCid++;
				TRACE(_T("CidFound indx=%d cid=%s \n"), mailPosition, cidName);
			}
			else if (srcHttp)
			{
				if (srcMHtmlHtml)
					MboxMail::m_EmbededImagesFoundMHtmlHttp++;
				else
					MboxMail::m_EmbededImagesFoundHttp++;
			}
			else if (srcHttps)
			{
				if (srcMHtmlHtml)
					MboxMail::m_EmbededImagesFoundMHtmlHttps++;
				else
					MboxMail::m_EmbededImagesFoundHttps++;
			}
			else if (srcData)
			{
				MboxMail::m_EmbededImagesFoundData++;
			}
			else if (srcLocalFile)
			{
				MboxMail::m_EmbededImagesFoundLocalFile++;
			}

			int length = srcEnd - fromBegin;
			outbuf->Append(fromBegin, length);

			outbuf->Append("\"file:\\\\\\");
			outbuf->Append(imgFile, imgFile.GetLength());
			outbuf->Append('\"');

			fromBegin = cidEnd;
			if (*cidEnd == '\"')
				fromBegin++;

			if (!runInvestigation)
			{
				BOOL ret = FileUtils::PathFileExist(imgFile);
				//BOOL ret = TRUE;
				if (ret)
				{
					; // TRACE(_T("UpdateInlineSrcImgPath: FOUND image %s\n"), imgFile);
				}
				else
				{
					;// TRACE(_T("UpdateInlineSrcImgPath: NOT FOUND image %s\n"), imgFile);
					imageFilePath = imgFile;
					CFileException ex;
					CFile fp;

					BOOL retryOpen = FALSE;
					if (!fp.Open(imageFilePath, CFile::modeWrite | CFile::modeCreate, &ex))
					{
						CString errorText;
						CString imageCachePath;
						CString rootPrintSubFolder = "ImageCache";
						//CString targetPrintSubFolder = baseFileArchiveName;
						CString targetPrintSubFolder;

						BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, imageCachePath, errorText);
						if (retval == FALSE)
						{
							HWND h = NULL; // we don't have any window yet  
							// TODO: what to do ?
							//int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
							//return -1;
						}
						retryOpen = TRUE;
					}

					if (retryOpen && !fp.Open(imageFilePath, CFile::modeWrite | CFile::modeCreate, &ex))
					{
						TCHAR szError[1024];
						ex.GetErrorMessage(szError, 1024);
						CFileStatus rStatus;
						BOOL ret = fp.GetStatus(rStatus);
						CString errorText(szError);

						HWND h = NULL;
						// Ignore for now
						//int answer = ::MessageBox(h, errorText, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
					}
					else
					{
						SimpleString*outb = MboxMail::get_tmpbuf();
						outb->ClearAndResize(pBodyFound->m_contentLength * 2);
						int retLen = MboxMail::DecodeBody(fpm, pBodyFound, mailPosition, outb);
						if (outb->Count() > 1500000)
							int deb = 1;

						fp.Write(outb->Data(), outb->Count());
						fp.Close();
						MboxMail::rel_tmpbuf();
					}
				}
			}
		}
		else
		{
			MboxMail::m_EmbededImagesNotFound++;

			if (srcCid)
			{
				MboxMail::m_EmbededImagesNotFoundCid++;
				TRACE(_T("CidNotFound indx=%d cid=%s \n"), mailPosition, cidName);
			}
			else if (srcHttp)
			{
				if (srcMHtmlHtml)
					MboxMail::m_EmbededImagesNotFoundMHtmlHttp++;
				else
					MboxMail::m_EmbededImagesNotFoundHttp++;
			}
			else if (srcHttps)
			{
				if (srcMHtmlHtml)
					MboxMail::m_EmbededImagesNotFoundMHtmlHttps++;
				else
					MboxMail::m_EmbededImagesNotFoundHttps++;
			}
			else if (srcData)
			{
				MboxMail::m_EmbededImagesNotFoundData++;
			}
			else if (srcLocalFile)
			{
				MboxMail::m_EmbededImagesNotFoundLocalFile++;
				TRACE(_T("LocalFileNotFound indx=%d cid=%s \n"), mailPosition, cidName);
			}
			else
				int deb = 1;

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

	fpm.Close();


	return 1;
}

int NListView::DetermineEmbeddedImages(char *inData, int indDataLen, int mailPosition)
{
	if ((mailPosition < 0) || (mailPosition >= MboxMail::s_mails.GetCount()))
		return -1;

	MboxMail *m = MboxMail::s_mails[mailPosition];

	int ret = DetermineEmbeddedImages(inData, indDataLen, m);
	return ret;
}

// Find all inline/embeded image attachment and mark as MboxMail->m->m_ContentDetailsArray->m_isEmbeddedImage
int NListView::DetermineEmbeddedImages(char *inData, int indDataLen, MboxMail *m)
{
	static char * img_pattern = "<img";
	static int img_patternLen = strlen(img_pattern);

	static char * src_pattern = "src=";
	static int src_patternLen = strlen(src_pattern);

	static char * cid_pattern = "cid:";
	static int cid_patternLen = strlen(cid_pattern);

	static char * http_pattern = "http:";
	static int http_patternLen = strlen(http_pattern);

	static char * https_pattern = "https:";
	static int https_patternLen = strlen(https_pattern);

	static char * mhtml_pattern = "mhtml:";
	static int mhtml_patternLen = strlen(mhtml_pattern);

	static char * data_pattern = "data:";
	static int data_patternLen = strlen(data_pattern);

	char *input = inData;
	int inputLength = indDataLen;
	char *inputEnd = input + inputLength;
	char *srcImgEnd = inputEnd;
	char *cidBegin;
	char *cidEnd;
	SimpleString cid;
	CString imageFileName;

	char *fromBegin = input;
	char *alt_pos;
	int patternLen = cid_patternLen;
	int mhtmlIntroLen = 0;
	char *srcBegin;
	char *srcEnd;

	CString contentTypeExtension;
	CString contentTypeMain;

	MailBodyContent *body;
	bool foundImage = false;
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
		if (contentTypeMain.CompareNoCase("text") != 0) {
			foundImage = true;
			break;
		}
	}

	if (!foundImage)
		return -1;

	CString cidName;

	char *pos = input;
	while ((pos != 0) && (pos < inputEnd))
	{
		pos = TextUtilsEx::strnstrUpper2Lower(pos, inputEnd, img_pattern, img_patternLen);
		if (pos == 0) {
			break;
		}
		pos += img_patternLen;

		if ((*pos == ' ') || (*pos == '\n') || (*pos == '\r'))
			pos += 1;
		else
			continue;

		// find end of <img src ...>
		srcImgEnd = TextUtilsEx::strnstrUpper2Lower(pos, inputEnd, ">", 1);
		if (srcImgEnd == 0) { // TODO: corrupted file ?
			break;
		}
		srcImgEnd++; // jump over

		alt_pos = TextUtilsEx::strnstrUpper2Lower(pos, srcImgEnd, "alt=\"", 5);

		pos = TextUtilsEx::strnstrUpper2Lower(pos, srcImgEnd, src_pattern, src_patternLen);
		if (pos == 0) {
			pos = srcImgEnd;  // jump over this <img .. > since no cid:
			continue;
		}

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
				MboxMail::m_EmbededImagesFoundMHtml++;
				pos = TextUtilsEx::strchar(pos, srcImgEnd, '!');
				if (pos) {
					mhtmlIntroLen = pos - cidBegin + 1;
					cidBegin = pos + 1;
					pos = 0;
					foundHTTP = TRUE;
					MboxMail::m_EmbededImagesFoundMHtmlHtml++;
				}
				else
				{
					MboxMail::m_EmbededImagesFoundUnexpectedMHtml++;
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
			}
		}

		if (pos == 0)
		{
			pos = cidBegin;
			pos = TextUtilsEx::strnstrUpper2Lower(pos, pos + data_patternLen, data_pattern, data_patternLen);
			if (pos) { // not implemented ; implemented by browser, etrc
				// set pos = 0 since src=data: case is not handled
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
				pos = srcImgEnd;  // jump over this <img .. > since no cid:
				fromBegin = pos;
				continue;
			}
			cidEnd = pos;
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

		cid.Copy(cidBegin, cidEnd - cidBegin);
		if (foundHTTP) {
			int retLen = TextUtilsEx::DecodeURL(cid.Data(), cid.Count());
			cid.SetCount(retLen);
		}

		cidName.Empty();
		cidName.Append(cid.Data(), cid.Count());

		MailBodyContent *pBodyFound;
		int mailPosition = 0; // not used anyway
		int retval = NListView::DetermineImageFileName(m, cidName, imageFileName, &pBodyFound, mailPosition);

		if (pBodyFound && (retval > 0))
		{
			pBodyFound->m_isEmbeddedImage = true;

			int length = srcEnd - fromBegin;

			fromBegin = cidEnd;
			if (*cidEnd == '\"')
				fromBegin++;

		}
		pos = srcImgEnd;
		int deb = 1;
	}

	return 1;
}

int NListView::FindFilenameCount(CMimeBody::CBodyList &bodies,  CString &fileName)
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

int NListView::DetermineImageFileName_SelectedItem(CMimeBody::CBodyList &bodies, MboxMail *m, CString &cidName, CString &imageFilePath, CMimeBody **foundBody, MyCArray<bool> &fileImgAlreadyCreatedArray, int mailPosition)
{
	if (m == 0)
	{
		// TODO: Assert ??
		return 1;
	}

	SimpleString cid;
	CString imageFileName;
	CString nameExtension;
	CString baseName;
	CString idExtension;
	CString locationExtension;
	CMimeBody *body;

	*foundBody = 0;

	CString mailIndex;  // not needed anymore here

	CMimeBody::CBodyList::const_iterator it;
	int bodyIndex = 0;

	for (it = bodies.begin(); it != bodies.end(); it++, bodyIndex++)
	{
		int count = fileImgAlreadyCreatedArray.GetCount();
		if (bodyIndex > (count - 1))
			fileImgAlreadyCreatedArray.Add(false);

		CMimeBody* pBP = *it;
		body = *it;

		// TODO: inefficient when we have bot text and html parts. 
		// We read and initialize bdy as text and later override with html.
		// Need to redo.

		string strName = pBP->GetFilename();
		if (strName.empty())
			strName = pBP->GetName();

		CString cStrName = strName.c_str();
		CString attachmentName = strName.c_str();

		CString disposition;
		MboxCMimeHelper::GetContentDisposition(pBP, disposition);

		if (disposition.CompareNoCase("inline") == 0)
		{
			const char* cDisposition = pBP->GetDisposition();
			int deb = 1;
		}

		bool isAttachmentInline = false;
		CString contentId;
		MboxCMimeHelper::GetContentID(pBP, contentId);

		contentId.Trim();
		contentId.Trim("<>");

		CString contentType;
		MboxCMimeHelper::GetContentType(pBP, contentType);

		CString contentLocation;
		MboxCMimeHelper::GetContentLocation(pBP, contentLocation);

		CString contentTypeExtension;
		CString contentTypeMain;
		BOOL isValidContentTypeExtension = FALSE;
		int pos = contentType.ReverseFind('/');
		if (pos > 0)
		{
			contentTypeExtension = contentType.Mid(pos + 1);
			contentTypeMain = contentType.Left(pos);
			CString contentTypeExt = "." + contentTypeExtension;
			isValidContentTypeExtension = IsSupportedPictureFileExtension(contentTypeExt);
		}

		bool isText = pBP->IsText();
		bool isMessage = pBP->IsMessage();
		bool isAttachment = MboxCMimeHelper::IsAttachment(pBP);

		if (contentTypeMain.CompareNoCase("message") == 0)
		{
			TRACE(_T("Content-Type=message index=%d\n"), mailPosition);
			int deb = 1;
		}

		if (isMessage)
			int deb = 1;


		if (isText && !isAttachment)
		{
			continue;
		}

		BOOL  isOctetStream = FALSE;
		if (contentTypeExtension.CompareNoCase("octet-stream") == 0)
			isOctetStream = TRUE;

		BOOL isValidAttachmentNameExtension = FALSE;
		if (!attachmentName.IsEmpty())
		{
			pos = attachmentName.ReverseFind('.');
			if (pos >= 0)
			{
				nameExtension = attachmentName.Mid(pos);
				isValidAttachmentNameExtension = IsSupportedPictureFileExtension(nameExtension);
			}
		}
		
		BOOL isValidContentIdExtension = FALSE;
		if (!contentId.IsEmpty())
		{
			pos = contentId.ReverseFind('.');
			if (pos >= 0)
			{
				idExtension = contentId.Mid(pos);
				isValidContentIdExtension = IsSupportedPictureFileExtension(idExtension);
			}
		}

		BOOL isValidContentLocationExtension = FALSE;
		if (!contentLocation.IsEmpty()) 
		{
			pos = contentLocation.ReverseFind('.');
			if (pos >= 0)
			{
				locationExtension = contentLocation.Mid(pos);
				isValidContentLocationExtension = IsSupportedPictureFileExtension(locationExtension);
			}
		}

		// Commented out to attempt to match regardless of Content-Type
		// Can this result in false matches ??
		if ((contentTypeMain.CompareNoCase("image") == 0) || (disposition.CompareNoCase("inline") == 0) || isOctetStream)
		{
			if (!contentId.IsEmpty())
			{
				if (contentId.Compare(cidName) == 0)
				{
					// TODO: we assume all contentIds are unique which may not be always true
					int fileCnt = 0;
					if (!attachmentName.IsEmpty())
						fileCnt = FindFilenameCount(bodies, attachmentName);

					if (!attachmentName.IsEmpty() && (fileCnt < 2))  // if no duplicate file names
					{
						if (isValidAttachmentNameExtension)
							imageFileName = mailIndex + attachmentName;
						else if (isValidContentTypeExtension)
							imageFileName = mailIndex + attachmentName + _T(".") + contentTypeExtension;
						else if (isValidContentIdExtension)
							imageFileName = mailIndex + attachmentName + idExtension;
						else if (isValidContentLocationExtension)
							imageFileName = mailIndex + attachmentName + locationExtension;
						else
						{
							const unsigned char* data = pBP->GetContent();
							int dataLength = pBP->GetContentLength();

							CStringW extensionW;
							CString extension;
							BOOL ret = loadImage((BYTE*)data, (size_t)dataLength, extensionW, extension);
							if (ret)
							{
								imageFileName = mailIndex + attachmentName + extension;
							}
							else
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
						imageFileName = mailIndex + contentId + "." + contentTypeExtension;
					}
					else
					{
						const unsigned char* data = pBP->GetContent();
						int dataLength = pBP->GetContentLength();

						CStringW extensionW;
						CString extension;
						BOOL ret = loadImage((BYTE*)data, (size_t)dataLength, extensionW, extension);
						if (ret)
						{
							imageFileName = mailIndex + contentId + extension;
						}
						else
							imageFileName = mailIndex + contentId + ".jpg";
					}
					*foundBody = body;
				}
				else
				{
					int deb = 1;
				}
			}

			if ((*foundBody == 0) && !contentLocation.IsEmpty())
			{
				if (contentLocation.Compare(cidName) == 0)
				{
					// TODO: we assume all contentLocations are unique which may not be always true
					int fileCnt = 0;
					if (!attachmentName.IsEmpty())
						fileCnt = FindFilenameCount(bodies, attachmentName);

					if (!attachmentName.IsEmpty() && (fileCnt < 2))
					{
						if (isValidAttachmentNameExtension)
							imageFileName = mailIndex + attachmentName;
						else if (isValidContentTypeExtension)
							imageFileName = mailIndex + attachmentName + "." + contentTypeExtension;
						else if (isValidContentIdExtension)
							imageFileName = mailIndex + attachmentName + idExtension;
						else if (isValidContentLocationExtension)
							imageFileName = mailIndex + attachmentName + locationExtension;
						else
						{
							const unsigned char* data = pBP->GetContent();
							int dataLength = pBP->GetContentLength();

							CStringW extensionW;
							CString extension;
							BOOL ret = loadImage((BYTE*)data, (size_t)dataLength, extensionW, extension);
							if (ret)
							{
								imageFileName = mailIndex + attachmentName + extension;
							}
							else
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
						imageFileName = mailIndex + contentLocation + "." + contentTypeExtension;
					}
					else
					{
						const unsigned char* data = pBP->GetContent();
						int dataLength = pBP->GetContentLength();

						CStringW extensionW;
						CString extension;
						BOOL ret = loadImage((BYTE*)data, (size_t)dataLength, extensionW, extension);
						if (ret)
						{
							imageFileName = mailIndex + contentLocation + extension;
						}
						else
							imageFileName = mailIndex + contentLocation + ".jpg";
					}
					*foundBody = body;
				}
				else
				{
					int deb = 1;
				}
			}

			if ((*foundBody == 0) && !attachmentName.IsEmpty())
			{
				if (attachmentName.Compare(cidName) == 0)
				{
					int fileCnt = FindFilenameCount(bodies, attachmentName);

					if (fileCnt > 1)
					{
						// There are duplicate attachment names in bodies but may not be in
						// attachments array. TODO: Better solution is still needed
						int deb = 1;
					}

					if (isValidAttachmentNameExtension)
						imageFileName = mailIndex + attachmentName;
					else if (isValidContentIdExtension)
						imageFileName = mailIndex + attachmentName + "." + contentTypeExtension;
					else if (isValidContentIdExtension)
						imageFileName = mailIndex + attachmentName + idExtension;
					else if (isValidContentLocationExtension)
						imageFileName = mailIndex + attachmentName + locationExtension;
					else
					{
						const unsigned char* data = pBP->GetContent();
						int dataLength = pBP->GetContentLength();

						CStringW extensionW;
						CString extension;
						BOOL ret = loadImage((BYTE*)data, (size_t)dataLength, extensionW, extension);
						if (ret)
						{
							imageFileName = mailIndex + attachmentName + extension;
						}
						else
							imageFileName = mailIndex + attachmentName + ".jpg";
					}
					*foundBody = body;
				}
				else
				{
					int deb = 1;
				}
			}
			int deb = 1;
		}

		if (*foundBody)
		{
			if (!imageFileName.IsEmpty())
			{
				int pos = imageFileName.ReverseFind('.');
				if (pos > 0)  // extension found
				{
					nameExtension = imageFileName.Mid(pos);
					baseName = imageFileName.Mid(0, pos);


					// In few cases the file name extension is incorrect and needs to be corercted in order for Internet Explorer to display the image instead of the blank image.
					// Othet Browser such as Chrome, Firefox don't need this fix.
					//if ((nameExtension.CompareNoCase(".gif") == 0) || (nameExtension.CompareNoCase(".png") == 0) || (nameExtension.CompareNoCase(".bmp") == 0))
					{
						const unsigned char* data = pBP->GetContent();
						int dataLength = pBP->GetContentLength();

						CStringW extensionW;
						CString extension;
						BOOL ret = loadImage((BYTE*)data, (size_t)dataLength, extensionW, extension);
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
					}
				}  // else extension found, all done
				else
				{
					int deb = 1;
				}
			}
			break;
		}

	}

	if (*foundBody)
	{
		int count = fileImgAlreadyCreatedArray.GetCount();
		if (bodyIndex > (count - 1))
			int deb = 1;
		else
			fileImgAlreadyCreatedArray.SetAt(bodyIndex, true);
	}
	else  // src=value didn't match contentId, contentLocation and attachmentName, i.e it is not embedded
	{
		int deb = 1;
	}

	FileUtils::MakeValidFileName(imageFileName);
	imageFilePath = imageFileName;

#if 0
	TRACE(_T("--->>> SUBJ=%s\n"), m->m_subj);
	if (imageFileName.IsEmpty())
	{
		TRACE(_T("\tI_CID=%s FNAME=%s\n"), cidName, imageFilePath);
		int deb = 1;
	}
	else
	{
		TRACE(_T("\tI_CID=%s FNAME=%s\n"), cidName, imageFilePath);
		int deb = 1;
	}
#endif

	return 1;
}

int NListView::GetMailBody_SelectedItem(CMimeBody::CBodyList &bodies, CMimeBody** pBody)
{
	CMimeBody::CBodyList::const_iterator it;
	CMimeBody *body;
	CString ext = "";
	CString curExt;

	if (pBody == 0)
		return -1;

	*pBody = 0;

	for (it = bodies.begin(); it != bodies.end(); it++)
	{
		CMimeBody* pBP = *it;
		body = *it;
		curExt = "txt";

		// Iterate all the header fields of this body part:
		CMimeHeader::CFieldList& fds = pBP->Fields();
		CMimeHeader::CFieldList::const_iterator itfd;

		for (itfd = fds.begin(); itfd != fds.end(); itfd++)
		{
			const CMimeField& fd = *itfd;
			const char *fname = fd.GetName();
			const char *fval = fd.GetValue();
			// Check content type to get mail extension TODO: encapsulate in function
			if (_stricmp(fname, "Content-Type") == 0 && _strnicmp(fval, "text/", 5) == 0) {
				const char *p = fd.GetValue() + 5;
				if (_strnicmp(p, "plain", 5) == 0)
					curExt = "txt";
				else  if (_strnicmp(p, "htm", 3) == 0)
					curExt = "htm";
				else 
				{
					//TRACE(_T("Content-Type=text %s index=%d"), mailPosition);
					if (_strnicmp(p, "xml", 3) == 0)
						curExt = "xml";
					else
						curExt = "txt";
				}
			}
		}

		bool isText = pBP->IsText();
		bool isMessage = pBP->IsMessage();
		bool isAttachment = MboxCMimeHelper::IsAttachment(pBP);
		if ((isText || isMessage) && !isAttachment)
		{
			if (curExt.CompareNoCase("htm") == 0)
			{
				*pBody = pBP;
				return 1;
			}
		}
	}
	return -1;
}


// NListView::CreateInlineImageFiles_SelectedItem and UpdateInlineSrcImgPath_SelectedItem have a lot of common code
// Abstract and support callback/client callback to process separate  functionality
int NListView::CreateInlineImageFiles_SelectedItem(CMimeBody::CBodyList &bodies, NMsgView *pMsgView, int mailPosition, MailBodyInfoArray &cidArray, MyCArray<bool> &fileImgAlreadyCreatedArray, bool runInvestigation)
{
	static char * img_pattern = "<img";
	static int img_patternLen = strlen(img_pattern);

	static char * src_pattern = "src=";
	static int src_patternLen = strlen(src_pattern);

	static char * cid_pattern = "cid:";
	static int cid_patternLen = strlen(cid_pattern);

	static char * http_pattern = "http:";
	static int http_patternLen = strlen(http_pattern);

	static char * https_pattern = "https:";
	static int https_patternLen = strlen(https_pattern);

	static char * mhtml_pattern = "mhtml:";
	static int mhtml_patternLen = strlen(mhtml_pattern);

	static char * data_pattern = "data:";
	static int data_patternLen = strlen(data_pattern);

	MboxMail *m = 0;
	if ((mailPosition < MboxMail::s_mails.GetCount()) && (mailPosition >= 0))
		m = MboxMail::s_mails[mailPosition];
	else
		return 1;

	char *cidBegin;
	char *cidEnd;
	SimpleString cid;
	CString imageFileName;
	CString imageFilePath;

	CMimeBody::CBodyList::const_iterator it;
	CMimeBody *pBP;
	bool foundImage = false;
	for (it = bodies.begin(); it != bodies.end(); it++)
	{
		pBP = *it;
		bool isText = pBP->IsText();
		if (!isText) {
			foundImage = true;
			break;
		}
	}
	if (!foundImage)
		return 1;

	CMimeBody* pBody = 0;
	int htmlTextCount = 0;
	for (it = bodies.begin(); it != bodies.end(); it++)
	{
		pBP = *it;

		CString contentType;
		MboxCMimeHelper::GetContentType(pBP, contentType);

		CString contentTypeExtension;
		CString contentTypeMain;
		int posSep = contentType.ReverseFind('/');
		if (posSep > 0)
		{
			contentTypeExtension = contentType.Mid(posSep + 1);
			contentTypeMain = contentType.Left(posSep);
		}

		if ((contentTypeExtension.CompareNoCase("html") != 0) && (contentTypeExtension.CompareNoCase("htm") != 0))
		{
			continue;
		}

		if (MboxCMimeHelper::IsAttachment(pBP))
			continue;

		pBody = pBP;

		char *contentData = (char*)pBody->GetContent();
		int contentDataLength = pBody->GetContentLength();

		char *input = contentData;
		int inputLength = contentDataLength;
		char *inputEnd = input + inputLength;
		char *srcImgEnd = inputEnd;

		htmlTextCount++;
		if (htmlTextCount > 1)
			int deb = 1;

		CString cidName;
		int cidCnt = 0;
		char *pos = input;
		while ((pos != 0) && (pos < inputEnd))
		{
			imageFileName.Empty();

			// Best would be to join all lines between <img ... and > and then process
			pos = TextUtilsEx::strnstrUpper2Lower(pos, inputEnd, img_pattern, img_patternLen);
			if (pos == 0) {
				break;
			}

			pos += img_patternLen;

			if ((*pos == ' ') || (*pos == '\n') || (*pos == '\r'))
				pos += 1;
			else
				continue;

			// find end of <img src ...>
			srcImgEnd = TextUtilsEx::strnstrUpper2Lower(pos, inputEnd, ">", 1);
			if (srcImgEnd == 0) { // TODO: corrupted file ?
				break;
			}
			srcImgEnd++; // jump over

			pos = TextUtilsEx::strnstrUpper2Lower(pos, srcImgEnd, src_pattern, src_patternLen);
			if (pos == 0) {
				pos = srcImgEnd;  // jump over this <img .. > sonce no cid:
				continue;
			}

			bool srcCid = false;
			bool srcHttp = false;
			bool srcHttps = false;
			bool srcMHtml = false;  // foung mhtml
			bool srcMHtmlHtml = false;  // found mhtml and http or https
			bool srcMHtmlNoHtml = false; // did not found mhtml and http or https
			bool srcMHtmlHttp = false;  // found mhtml and http 
			bool srcMHtmlHttps = false;  // found mhtml and https
			bool srcData = false;
			bool srcLocalFile = false;

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
				pos += cid_patternLen;
				cidBegin = pos;  // cid value
			}

			if (pos == 0)
			{
				pos = cidBegin;
				pos = TextUtilsEx::strnstrUpper2Lower(pos, pos + mhtml_patternLen, mhtml_pattern, mhtml_patternLen);
				if (pos) {
					srcMHtml = true;
					MboxMail::m_EmbededImagesFoundMHtml++;
					pos = TextUtilsEx::strchar(pos, srcImgEnd, '!');
					if (pos) {
						cidBegin = pos + 1;
						pos = 0;
						foundHTTP = true;
						srcMHtmlHtml = true;
						MboxMail::m_EmbededImagesFoundMHtmlHtml++;
					}
					else
					{
						MboxMail::m_EmbededImagesFoundUnexpectedMHtml++;
						pos = srcImgEnd;  // jump over this <img .. > since no cid:
						continue;
					}
				}
			}

			if (pos == 0)
			{
				pos = cidBegin;
				pos = TextUtilsEx::strnstrUpper2Lower(pos, pos + http_patternLen, http_pattern, http_patternLen);
				if (pos) {
					pos += 0;  // include http:
					cidBegin = pos;
					foundHTTP = true;
					srcHttp = true;
				}
			}
			if (pos == 0)
			{
				pos = cidBegin;
				pos = TextUtilsEx::strnstrUpper2Lower(pos, pos + https_patternLen, https_pattern, https_patternLen);
				if (pos) {
					pos += 0;  // include https:
					cidBegin = pos;
					foundHTTP = true;
					srcHttps = true;
				}
			}

			if (pos == 0)
			{
				pos = cidBegin;
				pos = TextUtilsEx::strnstrUpper2Lower(pos, pos + data_patternLen, data_pattern, data_patternLen);
				if (pos) { // not implemented ; implemented by browsers, etc
					//pos += 0;  // include data:
					//cidBegin = pos;
					// set pos = 0 since src=data: case is not handled
					srcData = true;
					//pos = 0;
					pos = srcImgEnd;  // jump over this <img .. > since no cid:
					continue;
				}
			}

			if (pos == 0)
			{
				// Local File
				// find end of cid token (or http or https) or local file path
				//pos = TextUtilsEx::strnstrUpper2Lower(cidBegin, srcImgEnd, "\"", 1);
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
					MboxMail::m_EmbededImagesNoMatch++;
					pos = srcImgEnd;  // jump over this <img .. > since no image
					continue;
				}
				srcLocalFile = true;
				cidEnd = pos;
			}
			else
			{
				if (pos != cidBegin)
					int deb = 1;

				// find end of cid token (or http or https)
				//pos = TextUtilsEx::strnstrUpper2Lower(pos, srcImgEnd, "\"", 1);
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
			}

			cid.Copy(cidBegin, cidEnd - cidBegin);
			if (foundHTTP) {
				int retLen = TextUtilsEx::DecodeURL(cid.Data(), cid.Count());
				cid.SetCount(retLen);
			}

			cidName = cid.Data();

			CMimeBody *pBodyFound = 0;
			int retval = NListView::DetermineImageFileName_SelectedItem(bodies, m, cidName, imageFileName, &pBodyFound, fileImgAlreadyCreatedArray, mailPosition);
			if (pBodyFound && (retval > 0))
			{
				CString imgFile = FileUtils::GetmboxviewTempPath() + imageFileName;
				MailBodyInfo* info = new MailBodyInfo;
				info->m_CID.Append(cidName);
				info->m_imgFileName.Append(imgFile);

				cidArray.Add(info);

				MboxMail::m_EmbededImagesFound++;

				if (srcCid)
				{
					MboxMail::m_EmbededImagesFoundCid++;
					TRACE(_T("CidFound indx=%d cid=%s \n"), mailPosition, cidName);
				}
				else if (srcHttp)
				{
					if (srcMHtmlHtml)
						MboxMail::m_EmbededImagesFoundMHtmlHttp++;
					else
						MboxMail::m_EmbededImagesFoundHttp++;
				}
				else if (srcHttps)
				{
					if (srcMHtmlHtml)
						MboxMail::m_EmbededImagesFoundMHtmlHttps++;
					else
						MboxMail::m_EmbededImagesFoundHttps++;
				}
				else if (srcData)
				{
					MboxMail::m_EmbededImagesFoundData++;
				}
				else if (srcLocalFile)
					MboxMail::m_EmbededImagesFoundLocalFile++;

				if (runInvestigation == FALSE)
				{
					int pos = pMsgView->m_attachments.FindAttachmentByName(imageFileName);
					if (pos < 0)
					{
						pBodyFound->WriteToFile(imgFile);

						CString disposition;
						MboxCMimeHelper::GetContentDisposition(pBodyFound, disposition);
						bool isAttachment = false;
						if (disposition.CompareNoCase("attachment") == 0)
							isAttachment = true;

						bool showAllAttachments = false;
						AttachmentConfigParams *attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
						if (attachmentConfigParams)
						{
							if (attachmentConfigParams->m_bShowAllAttachments_Window)
							{
								showAllAttachments = attachmentConfigParams->m_bShowAllAttachments_Window;
							}
						}

						if (MboxMail::s_mails.GetCount())
							int deb = 1;

						if (showAllAttachments)
						{
							pMsgView->m_bAttach = TRUE;
							pMsgView->m_attachments.AddInlineAttachment(imageFileName);
						}
					}
				}
			}
			else
			{
				MboxMail::m_EmbededImagesNotFound++;

				if (srcCid)
				{
					MboxMail::m_EmbededImagesNotFoundCid++;
					TRACE(_T("CidNotFound indx=%d cid=%s \n"), mailPosition, cidName);
				}
				else if (srcHttp)
				{
					if (srcMHtmlHtml)
						MboxMail::m_EmbededImagesNotFoundMHtmlHttp++;
					else
						MboxMail::m_EmbededImagesNotFoundHttp++;
				}
				else if (srcHttps)
				{
					if (srcMHtmlHtml)
						MboxMail::m_EmbededImagesNotFoundMHtmlHttps++;
					else
						MboxMail::m_EmbededImagesNotFoundHttps++;
				}
				else if (srcData)
				{
					MboxMail::m_EmbededImagesNotFoundData++;
				}
				else if (srcLocalFile)
				{
					MboxMail::m_EmbededImagesNotFoundLocalFile++;
					TRACE(_T("LocalFileNotFound indx=%d cid=%s \n"), mailPosition, cidName);
				}
				else
					int deb = 1;
			}

			pos++; // jump over \"
			pos = srcImgEnd;

			int deb = 1;
		}
#if 0
		TRACE(_T("Mail File=%s\n"), (LPCSTR)MboxMail::s_path);
		int i;
		MailBodyInfo* info;
		for (i = 0; i < cidArray.GetCount(); i++)
		{
			info = cidArray[i];
			TRACE(_T("CID=%s FILE=%s\n"), info->m_CID, info->m_imgFileName);
		}
		TRACE(_T("fileImgAlreadyCreatedArray\n"));
		for (i = 0; i < fileImgAlreadyCreatedArray.GetCount(); i++)
		{
			TRACE(_T("CID_BOOL=%d\n"), fileImgAlreadyCreatedArray[i]);
		}
#endif
	}

	return 1;
}

int NListView::UpdateInlineSrcImgPath_SelectedItem(char *inData, int indDataLen, SimpleString *outbuf, int mailPosition, bool useMailPosition, MailBodyInfoArray &cidArray)
{
	static char * img_pattern = "<img";
	static int img_patternLen = strlen(img_pattern);

	static char * src_pattern = "src=";
	static int src_patternLen = strlen(src_pattern);

	static char * cid_pattern = "cid:";
	static int cid_patternLen = strlen(cid_pattern);

	static char * http_pattern = "http:";
	static int http_patternLen = strlen(http_pattern);

	static char * https_pattern = "https:";
	static int https_patternLen = strlen(https_pattern);

	static char * mhtml_pattern = "mhtml:";
	static int mhtml_patternLen = strlen(mhtml_pattern);

	static char * data_pattern = "data:";
	static int data_patternLen = strlen(data_pattern);

	char *input = inData;
	int inputLength = indDataLen;
	char *inputEnd = input + inputLength;
	char *srcImgEnd = inputEnd;
	char *cidBegin;
	char *cidEnd;
	SimpleString cid;
	CString imageFileName;
	CString imageFilePath;

	char *fromBegin = input;
	//char *alt_pos;
	int patternLen = cid_patternLen;
	int mhtmlIntroLen = 0;
	char *srcBegin;
	char *srcEnd;

	MboxMail *m = 0;
	if (useMailPosition)
	{
		if ((mailPosition < MboxMail::s_mails.GetCount()) && (mailPosition >= 0))
			m = MboxMail::s_mails[mailPosition];
		else
			return 1;
	}

	CString imageCachePath = FileUtils::GetmboxviewTempPath();
	CString imgFile;
	CString cidName;

	char *pos = input;
	while ((pos != 0) && (pos < inputEnd))
	{
		pos = TextUtilsEx::strnstrUpper2Lower(pos, inputEnd, img_pattern, img_patternLen);
		if (pos == 0) {
			outbuf->Append(fromBegin, inputEnd - fromBegin);
			break;
		}
		pos += img_patternLen;

		if ((*pos == ' ') || (*pos == '\n') || (*pos == '\r'))
			pos += 1;
		else
			continue;

		// find end of <img src ...>
		srcImgEnd = TextUtilsEx::strnstrUpper2Lower(pos, inputEnd, ">", 1);
		if (srcImgEnd == 0) { // TODO: corrupted file ?
			outbuf->Append(fromBegin, inputEnd - fromBegin);
			break;
		}
		srcImgEnd++; // jump over

		//alt_pos = TextUtilsEx::strnstrUpper2Lower(pos, srcImgEnd, "alt=\"", 5);

		pos = TextUtilsEx::strnstrUpper2Lower(pos, srcImgEnd, src_pattern, src_patternLen);
		if (pos == 0) {
			pos = srcImgEnd;  // jump over this <img .. > since no cid:
			continue;
		}

		patternLen = cid_patternLen;
		srcBegin = pos;  // point to src=
		srcEnd = srcBegin + src_patternLen;
		pos += src_patternLen;
		
		BOOL inQuotes = FALSE;
		if (*pos == '"')
		{
			inQuotes = TRUE;
			pos++;
		}

		BOOL foundHTTP = FALSE;
		BOOL foundMHTML = FALSE;
		cidBegin = pos;
		pos = TextUtilsEx::strnstrUpper2Lower(pos, pos + cid_patternLen, cid_pattern, cid_patternLen);
		if (pos != 0)
		{
			patternLen = cid_patternLen;
			pos += cid_patternLen;
			cidBegin = pos;  // point to cid value
		}

		mhtmlIntroLen = 0;
		if (pos == 0)
		{
			pos = cidBegin;
			pos = TextUtilsEx::strnstrUpper2Lower(pos, pos + mhtml_patternLen, mhtml_pattern, mhtml_patternLen);
			if (pos) { 
				pos = TextUtilsEx::strchar(pos, srcImgEnd, '!');
				if (pos) {
					mhtmlIntroLen = pos - cidBegin + 1;
					cidBegin = pos + 1;
					pos = 0;
					foundHTTP = TRUE;
					foundMHTML = TRUE;
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
				foundHTTP = TRUE;
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
				foundHTTP = TRUE;
			}
		}

		if (pos == 0)
		{
			pos = cidBegin;
			pos = TextUtilsEx::strnstrUpper2Lower(pos, pos + data_patternLen, data_pattern, data_patternLen);
			if (pos) { // not implemented ; implemented by browsers, etc
				//patternLen = 0;
				//pos += patternLen;  // include data:
				//cidBegin = pos;
				// set pos = 0 since src=data: case is not handled
				pos = 0;
			}
		}

		// find cid value
		if (pos == 0)
		{
			// find end of cid token or local file path
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
				pos = srcImgEnd;  // jump over this <img .. > since no cid:
				continue;
			}
			cidEnd = pos;
			if ((srcImgEnd - cidEnd) > 10)
				int deb = 1;
		}
		else
		{
			// find end of cid token
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

		cid.Copy(cidBegin, cidEnd - cidBegin);
		if (foundHTTP) {
			int retLen = TextUtilsEx::DecodeURL(cid.Data(), cid.Count());
			cid.SetCount(retLen);
		}

		cidName.Empty();
		cidName.Append(cid.Data(), cid.Count());

		int i;
		MailBodyInfo* info = 0;
		bool cidFound = false;
		for (i = 0; i < cidArray.GetCount(); i++)
		{
			info = cidArray[i];
			if ((info->m_imgFileName.IsEmpty() == FALSE) && (info->m_CID.Compare(cidName) == 0))
			{
				cidFound = true;
				break;
			}
		}

		if (cidFound && info)
		{
			if (foundHTTP && !foundMHTML)
			{
				int deb = 1;
			}

			imgFile = info->m_imgFileName;

			int length = srcEnd - fromBegin;
			outbuf->Append(fromBegin, length);

			outbuf->Append("\"file:\\\\\\");
			outbuf->Append(imgFile, imgFile.GetLength());
			outbuf->Append('\"');

			fromBegin = cidEnd;
			if (*cidEnd == '\"')
				fromBegin++;

		}
		else
		{
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
	}
	return 1;
}

//
int NListView::CreateAttachmentCache_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName)
{
	WRITE_IMAGE_FILE_ARGS args;

	CString rootPrintSubFolder = "AttachmentCache";
	CString attachmentCachePath;
	CString errorText;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolderName, attachmentCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	/*IN*/
	args.cache = "";


	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcWriteAttachmentFile, (LPVOID)(WRITE_IMAGE_FILE_ARGS*)&args);

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
		//ASSERT(cancelledbyUser == TRUE);
		int loopCnt = 20;
		DWORD tc_start = GetTickCount();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		DWORD tc_end = GetTickCount();
		DWORD delta = tc_end - tc_start;
		TRACE("(CreateAttachmentCache_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

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
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}

	CString txt = _T("Do you want to open folder with all attachements?");
	OpenContainingFolderDlg dlg(txt, TRUE);
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		HWND h = GetSafeHwnd();
		HINSTANCE result = ShellExecute(h, _T("open"), attachmentCachePath, NULL, NULL, SW_SHOWNORMAL);
		CMainFrame::CheckShellExecuteResult(result, h);
		int deb = 1;
	}
	else if (nResponse == IDYES)
	{
		HWND h = GetSafeHwnd();
		HINSTANCE result = ShellExecute(h, _T("open"), attachmentCachePath, NULL, NULL, SW_SHOWNORMAL);
		CMainFrame::CheckShellExecuteResult(result, h);
		int deb = 1;
	}
	else if (nResponse == IDCANCEL)
	{
		int deb = 1;
	}
	return 1;
}

BOOL CreateAttachmentCache_WorkerThread(LPCSTR cache, BOOL mainThread, CString &errorText)
{
	CFile fpm;
	if (!fpm.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite))
	{
		// TODO: critical failure
		return FALSE;
	}

	int ni = MboxMail::s_mails.GetSize();

	//CString errorText;
	CString attachmentCachePath;
	CString rootPrintSubFolder = "AttachmentCache";
	CString targetPrintSubFolder;

	BOOL createDirOk = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, attachmentCachePath, errorText);

	//
	if (FileUtils::PathDirExists(attachmentCachePath))
	{
		MboxMail::pCUPDUPData->SetProgress(_T("Deleting all attachment files in the Attachment Cache directory ..."), 0);
		FileUtils::RemoveDirW(attachmentCachePath, false);
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
			MboxMail::pCUPDUPData->SetProgress(_T("Creating attachment files ..."), 0);
		}
	}

	double delta = (double)ni;
	if (delta <= 0) delta = 1;
	double step = delta / 100;
	double curstep = 1;
	double newstep = 0;
	CString fileNum;
	int nFileNum = 0;

	BOOL terminated = FALSE;
	CString itemNumberStr;
	std::vector <MailBodyContent*>::iterator it;


	if (!mainThread)
	{
		if (MboxMail::pCUPDUPData)
		{
			MboxMail::pCUPDUPData->SetProgress(_T("Creating attachment files ..."), 0);
		}
	}

	AttachmentMgr m_attachmentDB;

	MboxMail *m;
	for (int i = 0; i < ni; i++)
	{
		m_attachmentDB.Clear();

		m = MboxMail::s_mails[i];

		int mailPosition = i;
		NListView::PrintMailAttachments(&fpm, mailPosition, m_attachmentDB);

		if (!mainThread && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				terminated = TRUE;
				break;
			}

			newstep = ((double)(i - 0 + 1)) / step;
			if (newstep >= curstep)
			{
				curstep = newstep;
			}

			int nItem = (i + 1);
			if ((nItem % 100) == 0) {
				itemNumberStr.Format(_T("Creating attachment files ... %d of %d"), nItem, ni);
				if (MboxMail::pCUPDUPData)  MboxMail::pCUPDUPData->SetProgress(itemNumberStr, (UINT_PTR)(curstep));
			}
		}

		if (terminated)
		{
			errorText = "Already created attachment files will not be deleted.";
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

	CString rootPrintSubFolder = "EmlCache";
	CString emlCachePath;
	CString errorText;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolderName, emlCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	/*IN*/
	args.cache = "";


	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcWriteEmlFile, (LPVOID)(WRITE_IMAGE_FILE_ARGS*)&args);

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
		//ASSERT(cancelledbyUser == TRUE);
		int loopCnt = 20;
		DWORD tc_start = GetTickCount();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		DWORD tc_end = GetTickCount();
		DWORD delta = tc_end - tc_start;
		TRACE("(CreateEmlCache_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

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
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}

	CString txt = _T("Do you want to open folder with all Eml files?");
	OpenContainingFolderDlg dlg(txt, TRUE);
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		HWND h = GetSafeHwnd();
		HINSTANCE result = ShellExecute(h, _T("open"), emlCachePath, NULL, NULL, SW_SHOWNORMAL);
		CMainFrame::CheckShellExecuteResult(result, h);
		int deb = 1;
	}
	else if (nResponse == IDYES)
	{
		HWND h = GetSafeHwnd();
		HINSTANCE result = ShellExecute(h, _T("open"), emlCachePath, NULL, NULL, SW_SHOWNORMAL);
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
BOOL CreateEmlCache_WorkerThread(LPCSTR cache, BOOL mainThread, CString &errorText)
{
	CFile fpm;
	if (!fpm.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite))
	{
		// TODO: critical failure
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
		MboxMail::pCUPDUPData->SetProgress(_T("Deleting all Eml files in the Eml Cache directory ..."), 0);
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
	CString rootPrintSubFolder = "EmlCache";
	CString targetPrintSubFolder;

	BOOL createDirOk = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, emlCachePath, errorText);

	//
	if (FileUtils::PathDirExists(emlCachePath))
	{
		MboxMail::pCUPDUPData->SetProgress(_T("Deleting all Eml files in the Eml Cache directory ..."), 0);
		FileUtils::RemoveDirW(emlCachePath, false);
	}
	//

	// FileUtils::RemoveDirW(EmlCache, false); removes files only; no need to recreate EmlCache directory
#if 0
	emlCachePath.Empty();
	rootPrintSubFolder = "EmlCache";
	targetPrintSubFolder.Empty();

	createDirOk = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, emlCachePath, errorText);
#endif

	emlCachePath.Append("\\");


	if ((fpm.m_hFile == CFile::hFileNull) || !createDirOk)
	{
		return TRUE;
	}

	if (!mainThread)
	{
		if (MboxMail::pCUPDUPData)
		{
			MboxMail::pCUPDUPData->SetProgress(_T("Creating Eml files ..."), 0);
		}
	}

	double delta = (double)ni;
	if (delta <= 0) delta = 1;
	double step = delta / 100;
	double curstep = 1;
	double newstep = 0;
	CString fileNum;
	int nFileNum = 0;

	BOOL terminated = FALSE;
	CString itemNumberStr;
	std::vector <MailBodyContent*>::iterator it;


	if (!mainThread)
	{
		if (MboxMail::pCUPDUPData)
		{
			MboxMail::pCUPDUPData->SetProgress(_T("Creating Eml files ..."), 0);
		}
	}

	MboxMail *m;
	CString emlFile;
	for (int i = 0; i < ni; i++)
	{
		emlFile.Empty();
		m = MboxMail::s_mails[i];

		int mailPosition = i;
		NListView::PrintAsEmlFile(&fpm, mailPosition, emlFile);

		if (!mainThread && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				terminated = TRUE;
				break;
			}

			newstep = ((double)(i - 0 + 1)) / step;
			if (newstep >= curstep)
			{
				curstep = newstep;
			}

			int nItem = (i + 1);
			if ((nItem % 100) == 0) {
				itemNumberStr.Format(_T("Creating Eml files ... %d of %d"), nItem, ni);
				if (MboxMail::pCUPDUPData)  MboxMail::pCUPDUPData->SetProgress(itemNumberStr, (UINT_PTR)(curstep));
			}
		}

		if (terminated)
		{
			errorText = "Already created Eml files will not be deleted.";
		}
	}

	fpm.Close();
	return TRUE;
}

//

int NListView::CreateInlineImageCache_Thread(int firstMail, int lastMail, CString &targetPrintSubFolderName)
{
	WRITE_IMAGE_FILE_ARGS args;

	CString rootPrintSubFolder = "ImageCache";
	CString imageCachePath;
	CString errorText;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolderName, imageCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	/*IN*/
	args.cache = imageCachePath;


	/*OUT*/
	args.exitted = FALSE;
	args.ret = 1;

	int ret = 1;
	CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcWriteInlineImageFile, (LPVOID)(WRITE_IMAGE_FILE_ARGS*)&args);

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
		//ASSERT(cancelledbyUser == TRUE);
		int loopCnt = 20;
		DWORD tc_start = GetTickCount();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		DWORD tc_end = GetTickCount();
		DWORD delta = tc_end - tc_start;
		TRACE("(CreateInlineImageCache_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

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
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}
	return 1;
}

BOOL CreateInlineImageCache_WorkerThread(LPCSTR cache, BOOL mainThread, CString &errorText)
{
	CFile fpm;
	if (!fpm.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite))
	{
		// TODO: critical failure
		return FALSE;
	}


	CString mailFileName;
	FileUtils::CPathStripPath((char*)(LPCSTR)MboxMail::s_path, mailFileName);
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
	CString rootPrintSubFolder = "ImageCache";
	CString targetPrintSubFolder;

	BOOL createDirOk = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, imageCachePath, errorText);
	if ((fpm.m_hFile == CFile::hFileNull) || !createDirOk)
	{
		return TRUE;
	}
	imageCachePath.Append("\\");

	if (!mainThread)
	{
		if (MboxMail::pCUPDUPData)
		{
			MboxMail::pCUPDUPData->SetProgress(_T("Creating inline image files ..."), 0);
		}
	}
	int ni = MboxMail::s_mails.GetSize();
	double delta = (double)ni;
	if (delta <= 0) delta = 1;
	double step = delta / 100;
	double curstep = 1;
	double newstep = 0;
	CString fileNum;
	int nFileNum = 0;

	BOOL terminated = FALSE;
	CString itemNumberStr;
	std::vector <MailBodyContent*>::iterator it;


	if (!mainThread)
	{
		if (MboxMail::pCUPDUPData)
		{
			MboxMail::pCUPDUPData->SetProgress(_T("Index file created. Creating inline image files ..."), 0);
		}
	}

	MboxMail *m;
	for (int i = 0; i < ni; i++)
	{
		m = MboxMail::s_mails[i];

		int mailPosition = i;
		NListView::CreateInlineImageFiles(fpm, mailPosition, imageCachePath);

		if (!mainThread && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				terminated = TRUE;
				break;
			}

			newstep = ((double)(i - 0 + 1)) / step;
			if (newstep >= curstep)
			{
				curstep = newstep;
			}

			int nItem = (i + 1);
			if ((nItem % 100) == 0) {
				itemNumberStr.Format(_T("Creating inline image files ... %d of %d"), nItem, ni);
				if (MboxMail::pCUPDUPData)  MboxMail::pCUPDUPData->SetProgress(itemNumberStr, (UINT_PTR)(curstep));
			}
		}

		if (terminated)
		{
			errorText = "Already created image files will not be deleted.";
		}
	}

	fpm.Close();
	return TRUE;
}

int NListView::PrintMailSelectedToSingleTEXT_Thread(CString &targetPrintSubFolderName, CString &targetPrintFolderPath)
{
	PRINT_MAIL_GROUP_TO_SEPARATE_PDF_ARGS args;

	CString rootPrintSubFolder = "PrintCache";
	CString targetPrintSubFolder = targetPrintSubFolderName;
	CString printCachePath;
	CString errorText;

	MailIndexList *selectedMailsIndexList = 0;
	if (PopulateSelectedMailsList() > 0)
		selectedMailsIndexList = &m_selectedMailsList;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, printCachePath, errorText);
	if (retval == FALSE) {
		HWND h = NULL; // we don't have any window yet  
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
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
		//ASSERT(cancelledbyUser == TRUE);
		int loopCnt = 20;
		DWORD tc_start = GetTickCount();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		DWORD tc_end = GetTickCount();
		DWORD delta = tc_end - tc_start;
		TRACE("(PrintMailSelectedToSingleTEXT_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

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
			bool fileExists;
			CString fileName;

			int ret = MboxMail::MakeFileNameFromMailArchiveName(textType, fileName, targetPrintSubFolder, fileExists, errorText);

			CString txt = "Created TEXT file \n\n" + fileName;
			OpenContainingFolderDlg dlg(txt, FALSE);
			INT_PTR nResponse = dlg.DoModal();
			////////////
			if (nResponse == IDOK)
			{
				if (FileUtils::BrowseToFile(fileName) == FALSE) {
					HWND h = GetSafeHwnd();
					HINSTANCE result = ShellExecute(h, _T("open"), printCachePath, NULL, NULL, SW_SHOWNORMAL);
					CMainFrame::CheckShellExecuteResult(result, h);
				}
				int deb = 1;
			}
			else if (nResponse == IDYES)
			{
				HWND h = GetSafeHwnd();
				HINSTANCE result = ShellExecute(h, _T("open"), fileName, NULL, NULL, SW_SHOWNORMAL);
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
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else if (ret != -2)  // not IDCANCEL
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
		else
			MboxMail::assert_unexpected();
	}
	return 1;
}

int NListView::PrintMailSelectedToSingleTEXT_WorkerThread(MailIndexList *selectedMailIndexList, CString &targetPrintSubFolderName, CString &targetPrintFolderPath, CString &errorText)
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
//
int NListView::DetermineAttachmentName(CFile *fpm, int mailPosition, MailBodyContent *body, SimpleString *bodyData, CStringW &nameW, AttachmentMgr &attachmentDB)
{
	DBGT(__FUNCTION__);

	// This is based on NListView::SelectItem()

	CString cStrName = body->m_attachmentName;
	UINT nameCharsetId = body->m_attachmentNamePageCode;
	if (nameCharsetId == 0)
		nameCharsetId = CP_UTF8;

	CString contentTypeExtension;
	CString contentTypeMain;
	CString nameExtension;

	CString documentExtension;
	HtmlUtils::CommonMimeType2DocumentTypes(body->m_contentType, documentExtension);

	int pos = body->m_contentType.ReverseFind('/');
	if (pos > 0)
	{
		contentTypeExtension = body->m_contentType.Mid(pos + 1);
		contentTypeMain = body->m_contentType.Left(pos);
	}

	if (!documentExtension.IsEmpty())
	{
		contentTypeExtension = documentExtension.Mid(1);
		int deb = 1;
	}

	BOOL  isOctetStream = FALSE;
	BOOL alloc_tmpbuf = FALSE;
	if (contentTypeExtension.CompareNoCase("octet-stream") == 0)
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
	}

	DWORD error;

	CStringW contentTypeExtensionW;
	if (TextUtilsEx::Ansi2Wide(contentTypeExtension, contentTypeExtensionW, error))
	{
		int deb = 1;
	}

	CStringW contentTypeMainW;
	if (TextUtilsEx::Ansi2Wide(contentTypeMain, contentTypeMainW, error))
	{
		int deb = 1;
	}
	// 
	CStringW cStrNameW;
	if (TextUtilsEx::Str2Wide(cStrName, nameCharsetId, cStrNameW))
	{
		if (!cStrNameW.IsEmpty())
		{
			pos = cStrNameW.ReverseFind(L'.');
			if (pos < 0)  // no extension found
			{
				if (isOctetStream && bodyData)
				{
					const unsigned char* data = (unsigned char*)bodyData->Data();
					int dataLength = bodyData->Count();

					CStringW extensionW;
					CString extension;
					BOOL ret = loadImage((BYTE*)data, (size_t)dataLength, extensionW, extension);

					if (ret)
						cStrNameW += extensionW;
				}
				else
					cStrNameW += L"." + contentTypeExtensionW;
			}  
			// else // else extension found, all done
		}
		else // cStrNameW is empty
		{
			if (isOctetStream && bodyData)
			{
				const unsigned char* data = (unsigned char*)bodyData->Data();
				int dataLength = bodyData->Count();

				CStringW extensionW;
				CString extension;
				BOOL ret = loadImage((BYTE*)data, (size_t)dataLength, extensionW, extension);
				if (ret)
					cStrNameW = contentTypeMainW + extensionW;
			}
			else
				cStrNameW = contentTypeMainW + L"." + contentTypeExtensionW;

			int deb = 1;
		}
		int deb = 1;
	}
	//
	if (cStrNameW.IsEmpty())
	{
		if (!cStrName.IsEmpty())
		{
			pos = cStrName.ReverseFind('.');
			if (pos < 0)
			{
				if (isOctetStream && bodyData)
				{
					const unsigned char* data = (unsigned char*)bodyData->Data();
					int dataLength = bodyData->Count();

					CStringW extensionW;
					CString extension;
					BOOL ret = loadImage((BYTE*)data, (size_t)dataLength, extensionW, extension);
					if (ret)
					{
						cStrName = contentTypeMain + extension;
					}
				}
				else
					cStrName += "." + contentTypeExtension;
			}  // else extension found, all done
		}
		else // cStrName is empty
		{
			if (isOctetStream && bodyData)
			{
				const unsigned char* data = (unsigned char*)bodyData->Data();
				int dataLength = bodyData->Count();

				CStringW extensionW;
				CString extension;
				BOOL ret = loadImage((BYTE*)data, (size_t)dataLength, extensionW, extension);
				if (ret)
				{
					cStrName = contentTypeMain + extension;
				}
			}
			else
				cStrName = contentTypeMain + "." + contentTypeExtension;

			int deb = 1;
		}

		if (TextUtilsEx::Str2Wide(cStrName, nameCharsetId, cStrNameW))
		{
			int deb = 1;
		}
		// else TODO: errror, indicates bigger problem
	}
	if (cStrNameW.IsEmpty())
		cStrNameW = contentTypeMainW + L"." + contentTypeExtensionW;

	int prefix = attachmentDB.GetValidName(cStrNameW);
	if (prefix >= 0)
	{
		CStringW nameW;
		nameW.Format(L"%d_%s", prefix, cStrNameW);
		cStrNameW = nameW;
	}

	nameW = cStrNameW;

	if (alloc_tmpbuf)  // this is a big mess
		MboxMail::rel_tmpbuf();
	return 1;
}
//////////////

// Creates mail attachments files except embeded image attachments
//
//
int NListView::PrintMailAttachments(CFile *fpm, int mailPosition, AttachmentMgr &attachmentDB)
{
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
		if (!mboxFp.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite))
		{
			// TODO: critical failure
			return FALSE;
		}
		fpm = &mboxFp;
	}


	BOOL bReplaceWhiteWithUnderscore = FALSE;
	DWORD error;
	BOOL ret;

	CStringW resultW;
	CStringW validNameW;
	CStringW fileNameW;

	MailBodyContent *body;

	CString errorText;
	CString attachmentCachePath;
	CStringW printCachePathW;
	CString rootPrintSubFolder = "AttachmentCache";
	CString targetPrintSubFolder;

	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, attachmentCachePath, errorText);
	if (retval == FALSE)
	{
		HWND h = NULL; // we don't have any window yet  
		// TODO: what to do ?
		//int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		//return -1;
	}

	ret = TextUtilsEx::Ansi2Wide(attachmentCachePath, printCachePathW, error);

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
		if (body->IsAttachment())
		{
			if (showAllAttachments || !body->m_isEmbeddedImage)
				showAttachment = TRUE;
		}

		if (showAttachment)
		{
			UINT inCodePage = body->m_attachmentNamePageCode;
			resultW.Empty();

			MyCTime tt(m->m_timeDate);

			CString strDate;

			CString format;
			format = "%Y%m%d-%H%M%S-";
			strDate = tt.FormatLocalTm(format);

			CString uID;
			uID.Format("%07d", m->m_index);

			CString fileName;
			fileName.Append(strDate);
			fileName.Append(uID);
			fileName.Append(" ");
			CStringW fileNameW;
			if (TextUtilsEx::Ansi2Wide(fileName, fileNameW, error))
			{
				int deb = 1;
			}

			// make attachment name

			SimpleString*outbuf = MboxMail::get_tmpbuf();
			outbuf->ClearAndResize(body->m_contentLength * 2);

			int retLen = MboxMail::DecodeBody(*fpm, body, mailPosition, outbuf);
			if (outbuf->Count() > 1500000)
				int deb = 1;

			DetermineAttachmentName(fpm, mailPosition, body, outbuf, resultW, attachmentDB);

			fileNameW.Append(resultW);

			validNameW.Empty();

			FileUtils::MakeValidFileNameW(fileNameW, validNameW, bReplaceWhiteWithUnderscore);

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
		if (!mboxFp.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite))
		{
			// TODO: critical failure
			return FALSE;
		}
		fpm = &mboxFp;
	}

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	CString mailFileNameBase;
	if (!pFrame->m_NamePatternParams.m_bCustomFormat)
		MboxMail::MakeFileName(m, &pFrame->m_NamePatternParams, mailFileNameBase);
	else
		MboxMail::MakeFileName(m, pFrame->m_NamePatternParams.m_nameTemplateCnf, mailFileNameBase, pFrame->m_NamePatternParams.m_nFileNameFormatSizeLimit);

	CString emlCachePath;
	CString rootPrintSubFolder = "EmlCache";
	CString errorText;
	CString targetPrintSubFolder;
	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, emlCachePath, errorText);
	if (retval == FALSE) {
		return -1;
	}
	CString fileName = emlCachePath + "\\" + mailFileNameBase + ".eml";

	SimpleString *outbuf = MboxMail::m_outbuf;
	outbuf->ClearAndResize(10000);

	CFile fp;

	if (!fp.Open(fileName, CFile::modeWrite | CFile::modeCreate)) 
	{
		CString txt = _T("Could not create \"") + fileName;
		txt += _T("\" file.\nMake sure file is not open on other applications.");
		// PrintAsEmlFile is the static function. Must use global MessageBox
		// Not ideal because program is not blocked. TODO: invetsigate and change 

		HWND h = NULL;
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		// continue for now.
		return -1;
	}


	BOOL ret;
	outbuf->Clear();
	ret = m->GetBody(outbuf);
	// Remove From line at the begining of mail body

	static const char *cFromMailBegin = "From ";
	static const int cFromMailBeginLen = strlen(cFromMailBegin);

	char *p = outbuf->Data();
	int bodyCnt = outbuf->Count();
	char *e = p + bodyCnt;
	p = MimeParser::SkipEmptyLines(p, e);
	if (TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0)
	{
		// Can From line be splitted ?
		//p = MimeParser::GetMultiLine(p, e, line);
		p = MimeParser::EatNewLine(p, e);
		bodyCnt -= p - outbuf->Data();
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
		errorText = _T("Internal error, try again and see if that help, otherwise restart application");
		return -1;
	}

	CFile mboxFp;
	CFile *fpm_save = fpm;
	if (fpm == 0)
	{
		if (!mboxFp.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite))
		{
			errorText = _T("Could not open \"") + MboxMail::s_path;
			errorText += _T("\" file.\nMake sure file is not open on other applications.");
			return -1;
		}
		fpm = &mboxFp;
	}

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	CString mailFileNameBase;
	if (!pFrame->m_NamePatternParams.m_bCustomFormat)
		MboxMail::MakeFileName(m, &pFrame->m_NamePatternParams, mailFileNameBase);
	else
		MboxMail::MakeFileName(m, pFrame->m_NamePatternParams.m_nameTemplateCnf, mailFileNameBase, pFrame->m_NamePatternParams.m_nFileNameFormatSizeLimit);

	CString fileName = targetDirectory + mailFileNameBase + ".eml";

	SimpleString *outbuf = MboxMail::m_outbuf;
	outbuf->ClearAndResize(10000);

	CFile fp;
	if (!fp.Open(fileName, CFile::modeWrite | CFile::modeCreate))
	{
		errorText = _T("Could not create \"") + fileName;
		errorText += _T("\" file.\nMake sure file is not open on other applications.");
#if 0
		// ExportAsEmlFile is the static function. Must use global MessageBox
		// It can be called from non main thread and need to figure whether calling
		// :: MessageBox is valid. Note: it seems to work
		// Not ideal because program is not blocked. TODO: invetsigate and change 

		HWND h = NULL;
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
#endif
		// continue for now.
		return -1;
	}

	BOOL ret;
	outbuf->Clear();
	ret = m->GetBody(outbuf);
	// Remove From line at the begining of mail body

	static const char *cFromMailBegin = "From ";
	static const int cFromMailBeginLen = strlen(cFromMailBegin);

	char *p = outbuf->Data();
	int bodyCnt = outbuf->Count();
	char *e = p + bodyCnt;
	p = MimeParser::SkipEmptyLines(p, e);
	if (TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0)
	{
		// Can From line be splitted ?
		//p = MimeParser::GetMultiLine(p, e, line);
		p = MimeParser::EatNewLine(p, e);
		bodyCnt -= p - outbuf->Data();
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

int NListView::PrintAttachmentNamesAsText2CSV(int mailPosition, SimpleString *outbuf, CString &characterLimit, CString &attachmentSeparator)
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
		if (body->IsAttachment())
		{
			if (showAllAttachments || !body->m_isEmbeddedImage)
				showAttachment = TRUE;
		}

		if (showAttachment)
		{
			UINT inCodePage = body->m_attachmentNamePageCode;

			SimpleString*buf = 0;
			NListView::DetermineAttachmentName(fpm, mailPosition, body, buf, nameW, attachmentDB);

			validNameW.Empty();
			FileUtils::MakeValidFileNameW(nameW, validNameW, bReplaceWhiteWithUnderscore);

			UINT outCodePage = CP_UTF8;
			BOOL ret2 = TextUtilsEx::WStr2CodePage((wchar_t*)(LPCWSTR)validNameW, validNameW.GetLength(), outCodePage, &resultUTF8, error);

			int namelen = resultUTF8.Count();
			tmpbuf2.ClearAndResize(namelen + 10);

			begCount = outbuf->Count();

			if (attachmentCnt)
			{
				CString fmt = attachmentSeparator;
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

int NListView::FindAttachmentName(MboxMail *m, CString &searchString, BOOL bWholeWord, BOOL bCaseSensitive)
{
	if (m == 0)
	{
		// TODO: Assert ??
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
			if ((body->m_contentDisposition.CompareNoCase("attachment") == 0))
			{
				if (searchString.GetLength() == 1 && searchString[0] == '*')
					return 1;

				if (bWholeWord) {
					searchStringPos = g_tu.BMHSearchW((unsigned char *)(LPCSTR)body->m_attachmentName, body->m_attachmentName.GetLength(),
						(unsigned char *)(LPCSTR)searchString, searchString.GetLength(), bCaseSensitive);
				}
				else {
					searchStringPos = g_tu.BMHSearch((unsigned char *)(LPCSTR)body->m_attachmentName, body->m_attachmentName.GetLength(),
						(unsigned char *)(LPCSTR)searchString, searchString.GetLength(), bCaseSensitive);
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
	static int bgColorPat1_len = strlen(bgColorPat1);
	static char *bgColorPat2 = "background-color:";
	static int bgColorPat2_len = strlen(bgColorPat2);


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
					int lenData = c_pos_color - c_pos_last;
					outbuf->Append(c_pos_last, lenData);

					CString color;
					int lenColor = c_pos_end - c_pos_color;
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
						int lenData = c_pos_color - c_pos_last;
						outbuf->Append(c_pos_last, lenData);

						CString color;
						int lenColor = c_pos_end - c_pos_color;
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
		int lenData = inputEnd - c_pos_last;
		outbuf->Append(c_pos_last, lenData);
	}
	return 1;
}

int NListView::CreateEmptyFolder(CString &driveName, CString &mboxDirectory, CString &mboxFolderName, CString &parentSubFolderPath, CString &newFolderName)
{
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
		CString txt = _T("Please open mail file first.");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
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

	CString mboxListFile = path + "\\" + folderNameFile + ".mbox.mboxlist";

	SerializerHelper sz(mboxListFile);
	if (!sz.open(TRUE))
	{
		CString txt = _T("Could not create \"") + mboxListFile;
		txt += _T("\" file.\nMake sure file is not open on other applications.");
		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	// Create mboxlist file to allow reload of archive file list

	sz.writeInt(MAIL_LIST_VERSION2);			// version
	//sz.writeString(mboxFile);  // TODO: ??
	sz.writeInt64(MboxMail::s_fSize);	// root mail file size
	_int64 mailFileSize = 0; //  -1;  //  FileSize(mboxFilePath);  // we don't always create new mail archive and mail list
	sz.writeInt64(mailFileSize);	// file size
	sz.writeInt(0);

	sz.close();

	MailList *mailFolder = new MailList;
	mailFolder->m_nId = IDC_FOLDER_LIST;
	mailFolder->m_path = path;
	mailFolder->m_folderName = folderNameFile + ".mbox.mboxlist";

	MboxMail::m_folderList.insert_tail(mailFolder);

	return 1;
}


int NListView::ScanAllMailsInMbox()
{
	DWORD tc_start = GetTickCount();

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

	TRACE(_T("Mbox=%s\n\n"), MboxMail::s_path);

	bool runInvestigation = true;
	int iItem;
	for (iItem = 0; iItem < MboxMail::s_mails.GetCount(); iItem++)
	{
		// Get cached mail
		MboxMail *m = MboxMail::s_mails[iItem];

		// Get raw mail body
		CString bdy;
		m->GetBody(bdy);

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

		int ret = NListView::CreateInlineImageFiles_SelectedItem(bodies, pMsgView, mailPosition, cidArray, fileImgAlreadyCreatedArray, runInvestigation);
		if (cidArray.GetCount())
			hasInlineAttachments = true;

		MailBodyInfo* info;
		for (int i = 0; i < cidArray.GetCount(); i++)
		{
			info = cidArray[i];
			delete info;
		}
	}

	DWORD tc_curr = GetTickCount();
	DWORD tc_elapsed_seconds = (tc_curr - tc_start) / 1000;
	DWORD tc_elapsed_milliseconds = (tc_curr - tc_start) % 1000;

	TRACE(
		_T("Mbox=%s\n")
		_T("\tEmbededImagesNoMatch=%d\n")
		_T("\tEmbededImagesFound=%d EmbededImagesNotFound=%d\n")
		_T("\tEmbededImagesFoundMHtml=%d EmbededImagesFoundMHtmlHtml=%d EmbededImagesFoundUnexpectedMHtml=%d\n")
		_T("\tEmbededImagesFoundCid=%d EmbededImagesNotFoundCid=%d\n")
		_T("\tEmbededImagesFoundHttp=%d EmbededImagesNotFoundHttp=%d\n")
		_T("\tEmbededImagesFoundHttps=%d EmbededImagesNotFoundHttps=%d\n")
		_T("\tEmbededImagesFoundMHtmlHttp=%d EmbededImagesNotFoundMHtmlHttp=%d\n")
		_T("\tEmbededImagesFoundMHtmlHttps=%d EmbededImagesNotFoundMHtmlHttps=%d\n")
		_T("\tEmbededImagesFoundData=%d EmbededImagesNotFoundData=%d\n")
		_T("\tEmbededImagesFoundLocalFile=%d EmbededImagesNotFoundLocalFile=%d\n")
		_T("\tElapsedTime=%ld.%ld\n\n"),
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
	DWORD tc_start = GetTickCount();

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
	if (!fpm.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite)) {
		CString txt = _T("Could not open mail archive \"") + MboxMail::s_path;
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
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

	CString bdycharset = "UTF-8";
	CString bdy;

	UINT pageCode = 0;
	int textType = 1; //  Html
	bool runInvestigation = true;

	CString imageCachePath;
	bool createCacheInvestigate = false;

	TRACE(_T("Mbox=%s\n\n"), MboxMail::s_path);

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
			int textlen = MboxMail::GetMailBody_mboxview(fpm, iItem, outbuflarge, pageCode, textType);  // returns pageCode
			if (textlen != outbuflarge->Count())
				int deb = 1;

			if (outbuflarge->Count() != 0)
			{
				workbuf->ClearAndResize(outbuflarge->Count() * 2);
				NListView::UpdateInlineSrcImgPath(outbuflarge->Data(), outbuflarge->Count(), workbuf, 0, iItem, useMailPosition, runInvestigation);
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

			CreateInlineImageFiles(fpm, iItem, imageCachePath, runInvestigation);

		}
	}
	fpm.Close();

	DWORD tc_curr = GetTickCount();
	DWORD tc_elapsed_seconds = (tc_curr - tc_start)/1000;
	DWORD tc_elapsed_milliseconds = (tc_curr - tc_start)%1000;

	TRACE(
		_T("Mbox=%s\n")
		_T("\tEmbededImagesNoMatch=%d\n")
		_T("\tEmbededImagesFound=%d EmbededImagesNotFound=%d\n")
		_T("\tEmbededImagesFoundMHtml=%d EmbededImagesFoundMHtmlHtml=%d EmbededImagesFoundUnexpectedMHtml=%d\n")
		_T("\tEmbededImagesFoundCid=%d EmbededImagesNotFoundCid=%d\n")
		_T("\tEmbededImagesFoundHttp=%d EmbededImagesNotFoundHttp=%d\n")
		_T("\tEmbededImagesFoundHttps=%d EmbededImagesNotFoundHttps=%d\n")
		_T("\tEmbededImagesFoundMHtmlHttp=%d EmbededImagesNotFoundMHtmlHttp=%d\n")
		_T("\tEmbededImagesFoundMHtmlHttps=%d EmbededImagesNotFoundMHtmlHttps=%d\n")
		_T("\tEmbededImagesFoundData=%d EmbededImagesNotFoundData=%d\n")
		_T("\tEmbededImagesFoundLocalFile=%d EmbededImagesNotFoundLocalFile=%d\n")
		_T("\tElapsedTime=%ld.%ld\n\n"),
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


BOOL NListView::MyCTimeToOleTime(MyCTime &ctimeDateTime, COleDateTime &coleDateTime)
{

	SYSTEMTIME sysTime;
	ctimeDateTime.GetAsSystemTime(sysTime);
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
	m_lastStartDate = (time_t)-1;
	m_lastEndDate = (time_t)-1;
	m_mboxMailStartDate = (time_t)-1;
	m_mboxMailEndDate = (time_t)-1;
	m_needToRestoreArchiveListDateTime = FALSE;
	m_bNeedToFindMailMinMaxTime = TRUE;
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
		CString txt = _T("Folder Selected Mails List not empty. Overwrite?");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
			return -1;
	}

	CString mboxFileNamePath = MboxMail::s_path;

	if (mboxFileNamePath.IsEmpty()) {
		CString txt = _T("Please open mail file first.");
		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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

	CString folderCompletePath = driveName + mboxFileDirectory + "\\" + "Folders" + "\\" + folderPath;

	CString folderNameCompletePath = folderCompletePath + "\\" + folderName + ".mbox.mboxlist";

	if (!FileUtils::PathFileExist(folderNameCompletePath))
	{
		CString txt = _T("Mail List File \"") + folderNameCompletePath;
		txt += _T("\" doesn't exist.\nCan't reload.");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	SerializerHelper sz(folderNameCompletePath);
	if (!sz.open(FALSE)) {
		CString txt = _T("Could not open \"") + folderNameCompletePath;
		txt += _T("\" file.\nMake sure file is not open on other applications.");
		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	int version;
	_int64 mailFileSize;
	_int64 mboxFileSize;
	int mailListCnt;

	CString txt = _T("Mail list file\n\"") + folderNameCompletePath;
	txt += _T("\"\nhas incompatible version or file is corrupted.\nCan't reload.\nPlease remove the file.");

	if (!sz.readInt(&version)) {
		sz.close();
		return -1;
	}

	if (version != MAIL_LIST_VERSION2)
	{
		sz.close();

		CString text = _T("Mail list file\n\"") + folderNameCompletePath;
		CString strVersion;
		strVersion.Format(_T("%d"), (version - MAIL_LIST_VERSION_BASE));
		text += _T("\".\nhas incompatible version\"") + strVersion + "\". Expected version \"";
		strVersion.Format(_T("%d"), (MAIL_LIST_VERSION - MAIL_LIST_VERSION_BASE));
		text += strVersion + "\".\nCan't reload.\nPlease remove the file.";

		int answer = MessageBox(text, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	if (!sz.readInt64(&mailFileSize)) {
		sz.close();
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	if (!sz.readInt64(&mboxFileSize)) {
		sz.close();
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	if (!sz.readInt(&mailListCnt)) {
		sz.close();
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	_int64 nMailFileSize = FileUtils::FileSize(mboxFileNamePath);


	if ((mailListCnt < 0) || (mailListCnt > MboxMail::s_mails_ref.GetCount()) ||
		(mailFileSize != nMailFileSize))
	{
		sz.close();
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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
		CString txt = _T("Mail list file\n\"") + folderNameCompletePath;
		txt += _T("\"\n is invalid or corrupted. Can't reload.\nPlease remove the file.");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);

		return -1; // do nothing
	}

	sz.close();

	if (ret < 0)
	{
		MboxMail::s_mails_folder.SetSizeKeepData(0);

		CString txt = _T("Mail list file\n\"") + folderNameCompletePath;
		txt += _T("\"\n is invalid or corrupted. Can't reload.\nPlease remove the file.");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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
	dlg.m_title = "Select Folders";
	for (int ii = 0; ii < 100; ii++)
	{
		dlg.m_InList.Add("11111111111xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyy");
	}

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	NTreeView *pTreeView = 0;
	if (pFrame)
		pTreeView = pFrame->GetTreeView();

	CArray<CString> folderList;
	if (pTreeView)
	{
		CString mboxFileName;

		FileUtils::CPathStripPath((char*)(LPCSTR)MboxMail::s_path, mboxFileName);

		pTreeView->CreateFlatFolderList(mboxFileName, folderList);
	}

	dlg.m_InList.Copy(folderList);

	int nResponse = dlg.DoModal();
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


#include "gdiplus.h"
using namespace Gdiplus;

BOOL NListView::loadImage(BYTE* pData, size_t nSize, CStringW &extensionW, CString &extension)
{
	GdiplusStartupInput gdiplusStartupInput;
	ULONG_PTR m_gdiplusToken;
	Status sts = GdiplusStartup(&m_gdiplusToken, &gdiplusStartupInput, NULL);


	extensionW.Empty();
	extension = ".jpg";

	IStream *pStream = NULL;
	if (CreateStreamOnHGlobal(NULL, TRUE, &pStream) == S_OK)
	{
		if (pStream->Write(pData, (ULONG)nSize, NULL) == S_OK)
		{
			Bitmap *pBmp = Bitmap::FromStream(pStream);
			if (pBmp == 0)
			{
				pStream->Release();
				GdiplusShutdown(m_gdiplusToken);
				return FALSE;
			}
			else if (pBmp->GetLastStatus() != Status::Ok)
			{ 
				delete pBmp; 
				pStream->Release();
				GdiplusShutdown(m_gdiplusToken);
				return FALSE;
			}

			Gdiplus::Status gStatus = Gdiplus::UnknownImageFormat;
			GUID gFormat = Gdiplus::ImageFormatUndefined;
			gStatus = pBmp->GetRawFormat(&gFormat);
			if (gStatus == Gdiplus::Status::Ok)
			{
				if (gFormat == Gdiplus::ImageFormatEXIF) {
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatTIFF) {
					extension = ".tiff";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatGIF) {
					extension = ".gif";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatJPEG) {
					extension = ".jpeg";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatMemoryBMP) {
					extension = ".bmp";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatPNG) {
					extension = ".png";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatBMP) {
					extension = ".bmp";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatEMF) {
					extension = ".emf";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatWMF) {
					extension = ".wmf";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatIcon) {
					extension = ".ico";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatUndefined) {
					int deb = 1;
				}
			}
			int deb = 1;
		}
		pStream->Release();
		GdiplusShutdown(m_gdiplusToken);

		DWORD error;
		BOOL ret = TextUtilsEx::Ansi2Wide(extension, extensionW, error);
		return TRUE;
	}
	GdiplusShutdown(m_gdiplusToken);
	return FALSE;
}


int NListView::SetBackgroundColor(char *inData, int indDataLen, SimpleString *outbuf, BOOL ReplaceAllWhiteBackgrounTags)
{
	static char *bodyTag = "<body ";
	static int bodyTagLen = strlen(bodyTag);

	static char *styleTag = " style";
	static int styleTagLen = strlen(styleTag);

	static char *bgColorTag = "bgcolor";
	static int bgColorTagLen = strlen(bgColorTag);

	static char *backgroundColorTag = "background-color";
	static int backgroundColorTagLen = strlen(backgroundColorTag);

	static char *backgroundTag = "background";
	static int backgroundTagLen = strlen(backgroundTag);

	static char *bodyTagEnd = ">";
	static int bodyTagEndLen = strlen(bodyTagEnd);

	outbuf->ClearAndResize(indDataLen + 128);

	if (indDataLen < bodyTagLen)
		return 1;

	DWORD clr = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);
	if (clr == COLOR_WHITE)
		return 1;

	CString colorStr;

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

	char *pBackgroundColorTagBeg = 0;
	char *pBackgroundColorTagEnd = 0;

	// Replace background def in body ony.  Limit search to body begin and body end
	if (ReplaceAllWhiteBackgrounTags == FALSE)
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

		e = pBodyEnd+1;

		count = e - p;
		if (count <= 0)
			return 1;
	}

	outbuf->Append(inData, p - inData);

	p_mark = p;

	while (p < e)  // always true
	{
		pStyleTag = 0;
		pBgColorTag = 0;
		pBackgroundColorTag = 0;
		pBackgroundTag = 0;

		pBackgroundColorTagBeg = 0;
		pBackgroundColorTagEnd = 0;

		count = e - p;
		p = TextUtilsEx::findNoCaseP(p, count, "b", 1);
		if (p == 0)
			break;

		if (TextUtilsEx::strncmpUpper2Lower(p, e, backgroundColorTag, backgroundColorTagLen) == 0)
		{
			pBackgroundColorTag = p;
			p += backgroundColorTagLen;
		}
		else if (TextUtilsEx::strncmpUpper2Lower(p, e, bgColorTag, bgColorTagLen) == 0)
		{
			pBgColorTag = p;
			p += bgColorTagLen;
		}
		else if (TextUtilsEx::strncmpUpper2Lower(p, e, backgroundTag, backgroundTagLen) == 0)
		{
			pBackgroundTag = p;
			p += backgroundTagLen;
		}

		if ((pBackgroundColorTag == 0) && (pBgColorTag == 0) && (pBackgroundTag == 0))
		{
			p++;
			continue;
		}

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
			p++;

		pBackgroundColorTagBeg = p;

		p = TextUtilsEx::findOneOf(p, e, "\" ;>");
		if (p == 0)
		{
			outbuf->Clear();
			return 1;
		}

		pBackgroundColorTagEnd = p;

		CString pBackgroundColorTagVal;
		if (pBackgroundColorTagBeg && pBackgroundColorTagEnd)
		{
			int tagLength = pBackgroundColorTagEnd - pBackgroundColorTagBeg;
			char *b = pBackgroundColorTagVal.GetBufferSetLength(tagLength);
			::memcpy(b, pBackgroundColorTagBeg, tagLength);
			pBackgroundColorTagVal.ReleaseBuffer();

			if ((pBackgroundColorTagVal.CompareNoCase("fff") == 0) 
				|| (pBackgroundColorTagVal.CompareNoCase("ffffff") == 0)
				)
			{
				outbuf->Append(p_mark, pBackgroundColorTagBeg - p_mark);
				outbuf->Append(colorStr, colorStr.GetLength());
				p_mark = pBackgroundColorTagEnd;
			}
			else if ((pBackgroundColorTagVal.CompareNoCase("white") == 0)
				|| (pBackgroundColorTagVal.CompareNoCase("rgb(255,255,255)") == 0)
				|| (pBackgroundColorTagVal.CompareNoCase("hsl(0,0%,100%)") == 0)
				)
			{
				outbuf->Append('#');
				outbuf->Append(p_mark, pBackgroundColorTagBeg - p_mark);
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
	outbuf->Append(p_mark, e_data - p_mark);

	return 1;
}

int NListView::RemoveBodyBackgroundColor(char *inData, int indDataLen, SimpleString *outbuf, CString &bodyBackgroundColor)
{
	static char *bodyTag = "<body ";
	static int bodyTagLen = strlen(bodyTag);

	static char *styleTag = " style";
	static int styleTagLen = strlen(styleTag);

	static char *bgColorTag = "bgcolor";
	static int bgColorTagLen = strlen(bgColorTag);

	static char *backgroundColorTag = "background-color";
	static int backgroundColorTagLen = strlen(backgroundColorTag);

	static char *backgroundTag = "background";
	static int backgroundTagLen = strlen(backgroundTag);

	static char *bodyTagEnd = ">";
	static int bodyTagEndLen = strlen(bodyTagEnd);

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

	// Replace background def in body ony.  Limit search to body begin and body end
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

	//outbuf->Append(inData, p - inData);

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
					return 1;
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
}

int NListView::Color2Str(DWORD color, CString &colorStr)
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

int NListView::SaveAsEmlFile(CString &bdy)
{
	static const char *cFromMailBegin = "From ";
	static const int cFromMailBeginLen = strlen(cFromMailBegin);
	// Save mail
	CFile fp(FileUtils::GetmboxviewTempPath() + "message.eml", CFile::modeWrite | CFile::modeCreate);
	char *pb = (char*)((LPCSTR)bdy);
	int len = bdy.GetLength();
	char *e = pb + len;
	char *p = MimeParser::SkipEmptyLines(pb, e);
	if (TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0)
	{
		p = MimeParser::EatNewLine(p, e);
	}
	int eml_len = len - (p - pb);
	fp.Write(p, eml_len);
	//fp.Write(bdy, bdy.GetLength());
	fp.Close();
	return  1;
}


BOOL NListView::MatchIfFieldFolded(int mailPosition, char *fldName)
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
	m->GetBody(&mBody, maxBodyLength);

	char *data = mBody.Data();
	int datalen = mBody.Count();
	char *p = data;
	char *e = data + datalen;

	int fldNameLength = strlen(fldName);

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
	CString mboxListFile = datapath + "ArchiveCache" + "\\" + fileNameBase + "\\" + fileNameBase + "_USER.mbox.mboxlist";
	if (FileUtils::PathFileExist(mboxListFile))
	{
		listFileName.Append(mboxListFile);
	}
	return 1;
}

void NListView::SetListFocus()
{
	m_list.SetFocus();
}

void AttachmentMgr::Clear()
{
	m_attachmentMap.RemoveAll();
}

int AttachmentMgr::GetValidName(CStringW &inNameW)
{
	int nextId = -1;

	AttachmentDB::CPair *infopair = m_attachmentMap.PLookup(inNameW);

	// Check for duplicate names. Sometimes two or more names can represent different content
	if (infopair != nullptr)
	{
		AttachmentData &item = infopair->value;
		nextId = item.m_nextId;
		item.m_nextId++;
		return nextId;
	}
	else
	{
		AttachmentData item;
		m_attachmentMap[inNameW] = item;
	}
	return nextId;
}

int NListView::ForwardSingleMail(int iItem, BOOL progressBar, CString &progressText, CString &errorText)
{
	CString htmFileName;
	//CString errorText;
	//BOOL progressBar;
	//CString progressText;

	CFile fpm;
	if (!fpm.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite))
	{
		errorText = _T("Could not open \"") + MboxMail::s_path;
		errorText += _T("\" file.\nMake sure file is not open on other applications.");
		return -1;
	}

	int mailPosition = iItem;
	CString emlFile;
	CString targetDirectory = FileUtils::GetMboxviewLocalAppDataPath("MailService");

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
	instanceId.Format("%lld", id);
	
	if (pFrame)
	{
		password.Format("%s:%s", instanceId, pFrame->m_mailDB.SMTPConfig.UserPassword);
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
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame == 0) {
		//internal error;
		return -1;
	}

	if (pFrame->m_mailDB.SMTPConfig.UserAccount.IsEmpty() || pFrame->m_mailDB.SMTPConfig.UserPassword.IsEmpty())
	{
		CString txt = _T("User name and/or password are not provided.\n");
		txt += "Select \"File\"->\"SMTP Mail Server Config\" and provide user name and password.\n";
		HWND h = GetSafeHwnd(); // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	m_ForwardMailData.m_From = pFrame->m_mailDB.SMTPConfig.UserAccount;
	m_ForwardMailData.m_MailService = pFrame->m_mailDB.SMTPConfig.MailServiceName;

	ForwardMailDlg dlg;

	CStringW subjectW;
	if (iItem >= 0)
	{
		MboxMail *m = MboxMail::s_mails[iItem];
		TextUtilsEx::Str2Wide(m->m_subj, m->m_subj_charsetId, subjectW);
	}

	m_ForwardMailData.m_subjectW = subjectW;

	dlg.m_Data.Copy(m_ForwardMailData);

	if (dlg.DoModal() == IDOK)
	{
		m_ForwardMailData.Copy(dlg.m_Data);
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

	if (abs(MboxMail::b_mails_which_sorted) != 99) {

		CString txt = _T("Please sort all mails by conversation first.\n");
		txt += "Select \"View\"->\"Sort By\" ->\"Conversation\" or left click on the first column.";
		HWND h = GetSafeHwnd(); // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}

	int firstMail = iSelectedItem;
	int lastMail = iSelectedItem;
	FindFirstAndLastMailOfConversation(iSelectedItem, firstMail, lastMail);

	int selectedCnt = lastMail - firstMail + 1;
	if (selectedCnt > 100)
	{
		CString txt = _T("Mail sending limit imposed by Mail Service is not well defined.\n"
			"Number of mails you selected to forward is greater than 100\n"
			"Mail Service imposed sending limit may or not be exceeded and result in your mail account being suspended for 24 hurs.\n\n"
			"Do you want to forward selected emails anyway?\n");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
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
		int cnt = m_selectedMailsList.GetCount();
		for (int j = 0; j < cnt; j++)
		{
			i = m_selectedMailsList[j];
			MboxMail *m = MboxMail::s_mails[i];
			if (m->m_length > MaxMailSize)
				largeMailsCnt++;
		}

		if (largeMailsCnt)
		{
			CString txt;
			txt.Format("Found %d mails in the selected list that are larger than the max mail size of %d KB imposed by the active service %s. Sending will fail.\n"
				"Size of each mail is shown in the size column in the Summary Pane.\n\n"
				"Do you want to send anyway?\n", largeMailsCnt, MaxMailSize, pFrame->m_mailDB.SMTPConfig.MailServiceName);
			int answer = MessageBox(txt, _T("Warning"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
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

	int selectedCnt = m_selectedMailsList.GetCount();
	if (selectedCnt > 100)
	{
		CString txt = _T("Mail sending limit imposed by Mail Service is not well defined.\n"
			"Number of mails you selected to forward is greater than 100\n"
			"Mail Service imposed sending limit may or not be exceeded and result in your mail account being suspended for 24 hurs.\n"
			"You make take manual action to unblock the accout without delay\n\n"
			"Do you want to forward selected emails anyway?\n");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
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
		int cnt = m_selectedMailsList.GetCount();
		for (int j = 0; j < cnt; j++)
		{
			i = m_selectedMailsList[j];
			MboxMail *m = MboxMail::s_mails[i];
			if (m->m_length > MaxMailSize)
				largeMailsCnt++;
		}

		if (largeMailsCnt)
		{
			CString txt;
			txt.Format("Found %d mails in the selected list that are larger than the max mail size of %d KB imposed by the active service %s. Sending will fail.\n"
				"Size of each mail is shown in the size column in the Summary Pane.\n\n"
				"Do you want to send anyway?\n", largeMailsCnt, MaxMailSize, pFrame->m_mailDB.SMTPConfig.MailServiceName);
			int answer = MessageBox(txt, _T("Warning"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
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
			HWND h = NULL;
			int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
	}
	else
	{
		HWND h = NULL;
		int answer = ::MessageBox(h, errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		errorText.Append("Internal error. Try again.");
		return -1;
	}

	/*IN*/
	args.forwardMailsData.Copy(m_ForwardMailData);

	SimpleString result;
	SimpleString workBuff;
	const char *data = (LPCSTR)args.forwardMailsData.m_Text;
	int datalen = args.forwardMailsData.m_Text.GetLength();
	BOOL rval = TextUtilsEx::Str2CodePage((char*)data, datalen, m_acp, CP_UTF8, &result, &workBuff);

	CString appDataPath = FileUtils::GetMboxviewLocalAppDataPath("MailService");
	CString cStrNamePath = appDataPath + "MailText.txt";
	BOOL wret = FileUtils::Write2File(cStrNamePath, (unsigned char*)result.Data(), result.Count());

	args.forwardMailsData.m_Text.Empty();

	args.password = pFrame->m_mailDB.SMTPConfig.UserPassword;
	args.lview = this;
	args.targetPrintFolderPath = ""; // printCachePath;
	args.targetPrintSubFolderName = ""; // targetPrintSubFolder;
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
		//ASSERT(cancelledbyUser == TRUE);
		int loopCnt = 20;
		DWORD tc_start = GetTickCount();
		while ((loopCnt-- > 0) && (args.exitted == FALSE))
		{
			Sleep(25);
		}
		DWORD tc_end = GetTickCount();
		DWORD delta = tc_end - tc_start;
		TRACE("(ForwardSelectedMails_Thread)Waited %ld milliseconds for thread to exist.\n", delta);

		ret = -2;
	}
	else
		ret = args.ret;

	MboxMail::pCUPDUPData = NULL;

	if (ret == -2)
		return 1; // IDCANCEL

	CString ForwardMailOSuccessFilePath = appDataPath + "ForwardMailSuccess.txt";
	CString ForwardMailErrorFilePath = appDataPath + "ForwardMailError.txt";
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
		HWND h = GetSafeHwnd();
		int answer = ::MessageBox(h, errorTxt.Data(), _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	// We rely on errorTxt; wea are not consistent
	if (ret > 0)
	{
		if (args.errorText.IsEmpty())
		{
			int deb = 1;
		}
		else
		{
			MboxMail::assert_unexpected();
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
			MboxMail::assert_unexpected();
			return -1;
		}
	}
	else if (ret < 0) 
	{
		if (!args.errorText.IsEmpty())
		{
			HWND h = NULL;
			int answer = ::MessageBox(h, args.errorText, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
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

	CString fw = targetPath + "\\*.eml";
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
			filePath = targetPath + "\\" + fn;
			delStatus = DeleteFile(filePath);
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

	CString appDataPath = FileUtils::GetMboxviewLocalAppDataPath("MailService");
	int retdel = NListView::DeleteForwardEmlFiles(appDataPath);

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (selectedMailIndexList->GetCount() <= 0)
		return 1;

	double delta = (double)(selectedMailIndexList->GetCount());
	if (delta <= 0) delta = 1;
	double step = delta / 100;
	double curstep = 1;
	double newstep = 0;
	CString fileNum;
	int nFileNum;

	if (progressBar && MboxMail::pCUPDUPData)
	{
		//fileNum.Format(_T("Forwarding mails to %s  ... "), mailData.m_MailService);
		;// MboxMail::pCUPDUPData->SetProgress(fileNum, 1);
	}

	if (progressBar && MboxMail::pCUPDUPData)
		MboxMail::pCUPDUPData->SetProgress((UINT_PTR)curstep);

	int i;
	int cnt = selectedMailIndexList->GetCount();
	BOOL progress_Bar = TRUE;
	for (int j = 0; j < cnt; j++)
	{
		i = (*selectedMailIndexList)[j];
		if (progressBar && MboxMail::pCUPDUPData)
		{
			if (MboxMail::pCUPDUPData && MboxMail::pCUPDUPData->ShouldTerminate()) {
				break;
			}

			newstep = ((double)(j - 0 + 1)) / step;
			if (newstep >= curstep)
			{
				curstep = newstep;
			}
			curstep = 1;
			nFileNum = (j + 1);
			fileNum.Format(_T("Forwarding mails to %s  ... %d of %d"), mailData.m_MailService, nFileNum, cnt);
			if (MboxMail::pCUPDUPData)  MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(curstep));
		}
		int retval = ForwardSingleMail(i, progress_Bar, fileNum, errorText);
		if (retval < 0)
			return -1;
	}
	nFileNum = cnt;
	newstep = 100;
	fileNum.Format(_T("Forwarding mails to %s ... %d of %d"), mailData.m_MailService, nFileNum, cnt);
	if (MboxMail::pCUPDUPData) MboxMail::pCUPDUPData->SetProgress(fileNum, (UINT_PTR)(newstep));
	int deb = 1;
	return 1;
}

CString NListView::FixCommandLineArgument(CString &in)
{
	CString out;
	CString arg(in);
	arg.Replace("\"", "\"\"");
	out = "\"" + arg + "\"";
	return out;
}

CString NListView::FixCommandLineArgument(int in)
{
	CString out;
	out.Format("\"%d\"", in);
	return out;
}

INT64 NListView::ExecCommand_WorkerThread(int tcpPort, CString instanceId, CString &password, ForwardMailData &mailData, CString &emlFile, CString &errorText, BOOL progressBar, CString &progressText)
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (AfxSocketInit() == FALSE)
	{
		errorText.Append("Sockets Could Not Be Initialized");
		return -1;
	}

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
		errorText.Append("Internal error. Try again.");
		return -1;
	}

	CString appDataPath = FileUtils::GetMboxviewLocalAppDataPath("MailService");
	CString smtpConfigFilePath = appDataPath + "SMTP.ini";
	CString smtProtocolFilePath = appDataPath + "SMTPLog.txt";
	CString loggerFilePath = appDataPath +  "ForwardMailLog.txt";
	CString textFilePath = appDataPath + "MailText.txt";

	CString args;
	args = args + " --smtp-cnf " + FixCommandLineArgument(smtpConfigFilePath);
	args = args + " --to " + FixCommandLineArgument(mailData.m_To);
	if (!mailData.m_CC.IsEmpty())
		args = args + " --cc " + FixCommandLineArgument(mailData.m_CC);
	if (!mailData.m_BCC.IsEmpty())
		args = args + " --bcc " + FixCommandLineArgument(mailData.m_BCC);
	args = args + " --eml-file " + FixCommandLineArgument(emlFile);
	if (!textFilePath.IsEmpty())
		args = args + " --mail-text-file " + FixCommandLineArgument(textFilePath);
	args = args + " --tcp-port " + FixCommandLineArgument(tcpPort);
	if (!instanceId.IsEmpty())
		args = args + " --instance-id " + FixCommandLineArgument(instanceId);
	//
	if (m_enbaleForwardMailsLog)
	{
		if (!loggerFilePath.IsEmpty())
			args = args + " --logger-file " + FixCommandLineArgument(loggerFilePath);
	}
	if (m_enbaleSMTPProtocolLog)
	{
		if (!smtProtocolFilePath.IsEmpty())
			args = args + " --smtp-protocol-logger " + FixCommandLineArgument(smtProtocolFilePath);
	}

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
	CString processName("ForwardEmlFile.exe");
	CString errText;

	int step = 10;
	int stepCnt = 0;
	int nSeconds = 0;
	CString secondsBar;

	DWORD msec = 100;
	BOOL failed = FALSE;
	BOOL passwordAlreadySend = FALSE;
	int delaySendingPasswod = 20;
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
				CString errorText = GetLastErrorAsString();
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
				;
			}
			else if (!passwordAlreadySend)
			{
				// FIX: Make client non blocking
				MyTcpClient tcpClient(tcpPort);
				int error = tcpClient.ConnectSendClose(password, errorText);
				if (error == 0)
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
							secondsBar.Format(_T("%s     %d seconds"), progressText, nSeconds);
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
	errorText = GetLastErrorAsString();

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
	ShExecInfo.lpFile = "C:\\Windows\\System32\\taskkill.exe";
	ShExecInfo.lpParameters = "/F /IM" + processName + "/T";
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
				CString errorText = GetLastErrorAsString();
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
							secondsBar.Format(_T("%s     %d seconds"), progressText, nSeconds);
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
	errorText = GetLastErrorAsString();
	return -1;
}

int NListView::VerifyPathToForwardEmlFileExecutable(CString &ForwardEmlFileExePath , CString &errorText)
{
	CString processPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("processPath"));

	CString processDir;
	FileUtils::CPathGetPath(processPath, processDir);

	if (m_developmentMode)
	{
#if _DEBUG
		processDir = "G:\\MailKit\\ForwardEmlFile\\bin\\Debug\\netcoreapp3.1";
		processDir = "G:\\Documents\\GIT1.0.3.21 - Copy\\mboxviewer\\ForwardEmlFile\\bin\\Debug\\netcoreapp3.1";
#else
		processDir = "G:\\MailKit\\ForwardEmlFile\\bin\\Rekease\\netcoreapp3.1";
		processDir = "G:\\Documents\\GIT1.0.3.21 - Copy\\mboxviewer\\ForwardEmlFile\\bin\\Rekease\\netcoreapp3.1";
#endif
	}
	CString filePath = processDir + "\\ForwardMails\\ForwardEmlFile.exe";

	ForwardEmlFileExePath = filePath;

	if (!FileUtils::PathFileExist(filePath))
	{
		CString dir = processDir + "\\ForwardMails";
		errorText = _T("ForwardEmlFile.exe not found in ");
		errorText.Append(dir);
		errorText.Append(" directory. ");
		errorText.Append("MBox Viewer will not be able to forward emails until ForwardEmlFile.exe is reinstalled.\n");
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
		CString lastError = GetLastErrorAsString();
		CString txt = _T("Could not create \"") + filePath + _T("\"\n");
		txt += _T("Error: ") + lastError + "\n";
		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	// Create mboxlist file to allow reload of archive file list

	sz.writeInt(MAIL_LIST_VERSION2);			// version
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
		sz.writeInt(m->m_ContentDetailsArray.size());
	}
	sz.close();

	return 1;
}

int NListView::SaveAsLabelFile(MailArray *marray, CString &targetDir, CString &labelName, CString &errorText)
{
	static const char* cCategory = "category ";
	static const int cCategoryLength = strlen(cCategory);

	MailArray *mailsArray = marray;

#if 1
	CString mailFile = MboxMail::s_path;

	if (MboxMail::s_path.IsEmpty()) {
		CString txt = _T("Please open mail file first.");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}


	CString path = MboxMail::GetLastPath();
	CString datapath = MboxMail::GetLastDataPath();
	if (path.IsEmpty()) {
		MboxMail::assert_unexpected();
		return -1;  // Hopefully s_path wil fail first
	}
#endif

	CString labelsCachePath;
	CString rootPrintSubFolder = "LabelCache";
	//CString errorText;
	CString targetPrintSubFolder;
	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, labelsCachePath, errorText);
	if (retval == FALSE) {
		MboxMail::assert_unexpected();
		return -1;
	}

	if (labelName.GetLength() >= cCategoryLength)
	{
		char *p = (char*)(LPCSTR)labelName;
		char *e = p + labelName.GetLength();
		if (TextUtilsEx::strncmpUpper2Lower(p, e, cCategory, cCategoryLength) == 0)
			labelName.SetAt(cCategoryLength-1, '/');
	}

	CString delim = "/";
	CStringArray l;
	TextUtilsEx::SplitString(labelName, delim, l);

	int lcnt = l.GetSize();
	CString label = l.GetAt(lcnt-1);
	if (l.GetSize() > 1)
	{
		for (int i = 0; i < lcnt-1; i++)
		{
			CString ll = l.ElementAt(i);
			labelsCachePath = labelsCachePath + "\\" + ll;
			//TRACE("    |%s|\n", ll);
		}
	}

	BOOL cret = FileUtils::CreateDirectory(labelsCachePath);

	CFile fp;
	CString labelListFilePath = labelsCachePath + "\\" + label + ".mboxlist";

	_int64 mailFileSize = MboxMail::s_fSize;
	int wret = NListView::WriteMboxListFile_v2(marray, labelListFilePath, mailFileSize, errorText);
	if (wret < 0)
	{
		return -1;
	}

	return 1;
}

int NListView::LoadLabelListFile_v2(CString &listFilePath, CString &folderName)
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
		CString txt = _T("Folder Selected Mails List not empty. Overwrite?");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
			return -1;
	}
#endif

	CString mboxFileNamePath = MboxMail::s_path;

	if (mboxFileNamePath.IsEmpty()) {
		CString txt = _T("Please open mail file first.");
		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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
		CString txt = _T("Mail List File \"") + listFilePath;
		txt += _T("\" doesn't exist.\nCan't reload.");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	SerializerHelper sz(listFilePath);
	if (!sz.open(FALSE)) {
		CString lastError = GetLastErrorAsString();
		CString txt = _T("Could not create \"") + listFilePath + _T("\"\n");
		txt += _T("Error: ") + lastError + "\n";
		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	int version;
	_int64 mailFileSize;
	_int64 mboxFileSize;
	int mailListCnt;

	CString txt = _T("Mail list file\n\"") + listFilePath;
	txt += _T("\"\nhas incompatible version or file is corrupted.\nCan't reload.\nPlease remove the file.");

	if (!sz.readInt(&version)) {
		sz.close();
		MboxMail::assert_unexpected();
		return -1;
	}

	if (version != MAIL_LIST_VERSION2)
	{
		sz.close();

		CString text = _T("Mail list file\n\"") + listFilePath;
		CString strVersion;
		strVersion.Format(_T("%d"), (version - MAIL_LIST_VERSION_BASE));
		text += _T("\".\nhas incompatible version\"") + strVersion + "\". Expected version \"";
		strVersion.Format(_T("%d"), (MAIL_LIST_VERSION - MAIL_LIST_VERSION_BASE));
		text += strVersion + "\".\nCan't reload.\nPlease remove the file.";

		int answer = MessageBox(text, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	if (!sz.readInt64(&mailFileSize)) {
		sz.close();
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	if (!sz.readInt64(&mboxFileSize)) {
		sz.close();
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	if (!sz.readInt(&mailListCnt)) {
		sz.close();
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1;
	}

	_int64 nMailFileSize = FileUtils::FileSize(mboxFileNamePath);


	if ((mailListCnt < 0) || (mailListCnt > MboxMail::s_mails_ref.GetCount()) ||
		(mailFileSize != nMailFileSize))
	{
		sz.close();
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
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
		CString txt = _T("Mail list file\n\"") + listFilePath;
		txt += _T("\"\n is invalid or corrupted. Can't reload.\nPlease remove the file.");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		MboxMail::assert_unexpected();
		return -1; // do nothing
	}

	sz.close();

	if (ret < 0)
	{
		MboxMail::s_mails_folder.SetSizeKeepData(0);

		CString txt = _T("Mail list file\n\"") + listFilePath;
		txt += _T("\"\n is invalid or corrupted. Can't reload.\nPlease remove the file.");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
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



