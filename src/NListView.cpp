// NListView.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "NListView.h"
#pragma warning (disable : 4786)
#include "Mime.h"
#include "MimeCode.h"
#include "MboxMail.h"
#include "OpenContainingFolderDlg.h"
#include "FindInMailDlg.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

const CUPDUPDATA* NListView::pCUPDUPData = NULL;

__int64 FileSeek(HANDLE hf, __int64 distance, DWORD MoveMethod)
{
	LARGE_INTEGER li;

	li.QuadPart = distance;

	li.LowPart = SetFilePointer(hf,
		li.LowPart,
		&li.HighPart,
		MoveMethod);

	if (li.LowPart == INVALID_SET_FILE_POINTER && GetLastError()
		!= NO_ERROR)
	{
		li.QuadPart = -1;
	}

	return li.QuadPart;
}

BOOL PathFileExist(LPCSTR path)
{
	DWORD fa = GetFileAttributes(path);

	if( fa != (DWORD)0xFFFFFFFF)
		return TRUE;
	return FALSE;
}

CString GetmboxviewTempPath(char *name = 0)
{
	char	buf[_MAX_PATH];
	GetTempPath(_MAX_PATH, buf);
	if( ! PathFileExist(buf) )
		strcpy(buf,"\\");
	strcat(buf, "mboxview\\");
	if (name) {
		strcat(buf, name);
		strcat(buf, "\\");
	}
	CreateDirectory(buf,NULL);
	return buf;
}

BOOL RemoveDir(CString & dir, bool recursive = false)
{
	WIN32_FIND_DATA FileData;
	HANDLE hSearch;
	CString	appPath;
	BOOL	bFinished = FALSE;
	// Start searching for all files in the current directory.
	CString searchPath = dir + "\\*.*";
	hSearch = FindFirstFile(searchPath, &FileData); 
	if (hSearch == INVALID_HANDLE_VALUE) {
		TRACE("No files found.");
		return FALSE;
	}
	while ( ! bFinished ) {
		if( ! ( strcmp(FileData.cFileName, ".") == 0 || strcmp(FileData.cFileName, "..") == 0 ) ) {
			CString	fileFound = dir + "\\" + FileData.cFileName;
			if( FileData.dwFileAttributes != FILE_ATTRIBUTE_DIRECTORY )
				DeleteFile( (LPCSTR) fileFound );
			else if (recursive)
				RemoveDir( dir + "\\" + CString(FileData.cFileName) );
		}
		if( ! FindNextFile( hSearch, &FileData ) )
			bFinished = TRUE;
	}
	FindClose(hSearch);
	//RemoveDirectory( dir );
	return TRUE;
}

CString CreateTempFileName(CString ext = "htm")
{
	CString fmt = GetmboxviewTempPath() + "PTT%05d."+ext;
	CString fileName;
ripeti:
	fileName.Format(fmt, abs((int)(1+(int) (100000.0*rand()/(RAND_MAX+1.0)))));
	if( PathFileExist(fileName) )
		goto ripeti;

	return fileName; //+_T(".HTM");
}

IMPLEMENT_DYNCREATE(NListView, CWnd)

/////////////////////////////////////////////////////////////////////////////
// NListView

CString GetDateFormat(int i);

NListView::NListView() : m_list(this), m_lastStartDate((time_t)-1), m_lastEndDate((time_t)-1)
{
	ResetFileMapView();

	m_searchStringInMail.Empty();
	m_bCaseSensInMail = TRUE;
	m_bWholeWordInMail = FALSE;

	m_gmtTime = 0;
	m_bStartSearchAtSelectedItem = 0; // FALSE; TODO: is this desired feature ?
	m_bFindNext = TRUE;
	m_bEditFindFirst = FALSE;
	m_lastFindPos = -1;
	m_searchString.Empty();
	m_bCaseSens = TRUE;
	m_bWholeWord = FALSE;
	m_bFrom = TRUE;
	m_bTo = TRUE;
	m_bSubject = TRUE;
	m_bContent = FALSE;
	m_bAttachments = FALSE;
	m_bHighlightAll = FALSE;
	m_filterDates = FALSE;
	m_lastSel = -1;
	m_bInFind = FALSE;
	int iFormat = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "format");
	m_format = GetDateFormat(iFormat);
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
}

NListView::~NListView()
{
}

BEGIN_MESSAGE_MAP(NListView, CWnd)
	//{{AFX_MSG_MAP(NListView)
	ON_WM_CREATE()
	ON_WM_SIZE()
	ON_WM_MOUSEWHEEL()
	//}}AFX_MSG_MAP
	//ON_NOTIFY(LVN_ITEMACTIVATE, IDC_LIST, OnActivating)
	ON_NOTIFY(NM_CLICK, IDC_LIST, OnActivating)  // Left Click
	ON_NOTIFY(NM_RCLICK, IDC_LIST, OnRClick)  // Right Click Menu
	ON_NOTIFY(NM_DBLCLK, IDC_LIST, OnDoubleClick)
	ON_NOTIFY(LVN_COLUMNCLICK, IDC_LIST, OnColumnClick)
	ON_NOTIFY(LVN_KEYDOWN, IDC_LIST, OnKeydown)
	ON_NOTIFY(NM_CUSTOMDRAW, IDC_LIST, OnCustomDraw)
	ON_NOTIFY(LVN_GETDISPINFO, IDC_LIST, OnGetDispInfo)
	ON_UPDATE_COMMAND_UI(ID_EDIT_FIND, OnUpdateEditFind)
	ON_COMMAND(ID_EDIT_FIND, OnEditFind)
	ON_UPDATE_COMMAND_UI(ID_EDIT_FINDAGAIN, OnUpdateEditFindAgain)
	ON_COMMAND(ID_EDIT_FINDAGAIN, OnEditFindAgain)
	ON_COMMAND(ID_EDIT_VIEWEML, &NListView::OnEditVieweml)
	ON_UPDATE_COMMAND_UI(ID_EDIT_VIEWEML, &NListView::OnUpdateEditVieweml)
	//ON_COMMAND(ID_VIEW_CONVERSATIONS, &NListView::OnViewConversations)
	//ON_UPDATE_COMMAND_UI(ID_VIEW_CONVERSATIONS, &NListView::OnUpdateViewConversations)

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
	NMITEMACTIVATE *pnm = (NMITEMACTIVATE *)pNMHDR;
	if( m_lastSel == pnm->iItem ) {
		CString path = GetmboxviewTempPath();

		if( PathFileExist(path) )
			ShellExecute(NULL, _T("open"), path, NULL,NULL, SW_SHOWNORMAL );
	}
	*pResult = 0;
}


void NListView::OnActivating(NMHDR* pNMHDR, LRESULT* pResult) 
{
	NMITEMACTIVATE *pnm = (NMITEMACTIVATE *)pNMHDR;
	if( m_lastSel != pnm->iItem ) {
		TRACE("Selecting %d\n", pnm->iItem);
		SelectItem(pnm->iItem);
	}

	*pResult = 0;
}

void NListView::OnRClick(NMHDR* pNMHDR, LRESULT* pResult)
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

	const UINT S_TEXT_Id = 1;
	AppendMenu(&printToSubMenu, S_TEXT_Id, _T("Text.."));

	const UINT S_HTML_Id = 2;
	AppendMenu(&printToSubMenu, S_HTML_Id, _T("HTML..."));

	menu.AppendMenu(MF_POPUP | MF_STRING, (UINT)printToSubMenu.GetSafeHmenu(), _T("Print To"));
	menu.AppendMenu(MF_SEPARATOR);

	const UINT S_TEXT_GROUP_Id = 3;
	AppendMenu(&printGroupToSubMenu, S_TEXT_GROUP_Id, _T("Text.."));

	const UINT S_HTML_GROUP_Id = 4;
	AppendMenu(&printGroupToSubMenu, S_HTML_GROUP_Id, _T("HTML..."));

	menu.AppendMenu(MF_POPUP | MF_STRING, (UINT)printGroupToSubMenu.GetSafeHmenu(), _T("Print Related Mails To"));
	menu.AppendMenu(MF_SEPARATOR);

	const UINT S_HTML_OPEN_Id = 5;
	AppendMenu(&menu, S_HTML_OPEN_Id, _T("Open in Browser"));

	const UINT S_HTML_OPEN_RELATED_Id = 6;
	AppendMenu(&menu, S_HTML_OPEN_RELATED_Id, _T("Open Related Emails in Browser"));

	const UINT S_HTML_FIND_Id = 7;
	AppendMenu(&menu, S_HTML_FIND_Id, _T("Find"));

	UINT command = menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, this);

	UINT nFlags = TPM_RETURNCMD;
	CString menuString;
	int chrCnt = menu.GetMenuString(command, menuString, nFlags);

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	switch (command)
	{
	case S_HTML_Id: {
		if (pFrame) {
			pFrame->OnPrintSingleMailtoText(iItem, 1);
		}
		int deb = 1;
	}
	break;
	case S_TEXT_Id: {
		if (pFrame) {
			pFrame->OnPrintSingleMailtoText(iItem, 0);
		}
		int deb = 1;
	}
	break;
	case S_HTML_GROUP_Id: {
		{
			PrintMailGroupToText(iItem, 1);
		}
		int deb = 1;
	}
	break;
	case S_TEXT_GROUP_Id: {
		{
			PrintMailGroupToText(iItem, 0);
		}
	}
	break;
	case S_HTML_OPEN_Id: {
		if (pFrame) {
			BOOL forceOpen = TRUE;
			pFrame->OnPrintSingleMailtoText(iItem, 1, forceOpen);
		}
	}
	break;
	case S_HTML_OPEN_RELATED_Id: {
		{
			BOOL forceOpen = TRUE;
			PrintMailGroupToText(iItem, 1, forceOpen);
		}
	}
	break;
	case S_HTML_FIND_Id: {
		CFindInMailDlg dlg;
		dlg.m_bWholeWord = m_bWholeWordInMail;
		dlg.m_bCaseSensitive = m_bCaseSensInMail;
		dlg.m_string = m_searchStringInMail;
		if (dlg.DoModal() == IDOK) {
			m_searchStringInMail = dlg.m_string;
			m_bWholeWordInMail = dlg.m_bWholeWord;
			m_bCaseSensInMail = dlg.m_bCaseSensitive;

			CMainFrame *pFrame = (CMainFrame*)AfxGetMainWnd();
			if (!pFrame)
				break;
			if (!::IsWindow(pFrame->m_hWnd) || !pFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)))
				break;
			NMsgView *pView = pFrame->GetMsgView();
			if (!pView || !::IsWindow(pView->m_hWnd))
				break;

			pView->FindStringInIHTMLDocument(m_searchStringInMail, m_bWholeWordInMail, m_bCaseSensInMail);

			int dbg = 1;
		}
	}
	break;

	default: {
		int deb = 1;
	}
	break;
	}


	if (m_lastSel != iItem) {
		TRACE("Selecting %d\n", iItem);
		SelectItem(iItem);
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


void NListView::SortByColumn(int colNumber)
{
	bool mustSort = false;

	// to follow OnEditFindAgain approach
	if (m_bInFind)  
		return;

	m_bInFind = true;

	switch (colNumber) {
	case 0: // !
		if (abs(MboxMail::b_mails_which_sorted) == 99) {
			MboxMail::b_mails_which_sorted = -MboxMail::b_mails_which_sorted;
		}
		else {
			MboxMail::b_mails_which_sorted = 99;
		}
		MboxMail::SortByConverstionGroups(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
		break;
	case 1: // date
		if (abs(MboxMail::b_mails_which_sorted) == 1) {
			MboxMail::b_mails_which_sorted = -MboxMail::b_mails_which_sorted;
		} else {
			MboxMail::b_mails_which_sorted = 1;
		}
		MboxMail::SortByDate(0, MboxMail::b_mails_which_sorted < 0);
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

		RedrawMails();
		// MessageBeep(MB_OK); // too much ??
	}
	else
		MarkColumns();
	m_bInFind = false;
}

void NListView::OnKeydown(NMHDR* pNMHDR, LRESULT* pResult) 
{
//	m_headerCtrl.PostMessage(WM_CANCELMODE);
	NMLVKEYDOWN* pLVKeyDow = (NMLVKEYDOWN*)pNMHDR;
	*pResult = 0;
	int whichSel;
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
	
	if( !m_list.Create(WS_CHILD|WS_VISIBLE|LVS_SINGLESEL|LVS_REPORT|LVS_SHOWSELALWAYS|LVS_OWNERDATA, CRect(), this, IDC_LIST) )
		return -1;

	m_list.SetExtendedStyle(LVS_EX_FULLROWSELECT);	m_list.SendMessage((CCM_FIRST + 0x7), 5, 0);
	m_list.SetTextColor (::GetSysColor(COLOR_WINDOWTEXT));
	ResetFont();

	m_list.InsertColumn(0, "!", LVCFMT_LEFT, 22, 0);
	m_list.InsertColumn(1, "date", LVCFMT_LEFT, 100, 0);
	m_list.InsertColumn(2, "from", LVCFMT_LEFT, 150, 0);
	m_list.InsertColumn(3, "to", LVCFMT_LEFT, 150, 0);
	m_list.InsertColumn(4, "subject", LVCFMT_LEFT, 400, 0);
	m_list.InsertColumn(5, _T("size(KB)"), LVCFMT_LEFT, 120, 0);

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
	CHeaderCtrl* lhdr = m_list.GetHeaderCtrl();
	int nClmCnt = lhdr->GetItemCount();
	HDITEM rHdr;
	for (int j = 0; j < nClmCnt; j++) {
		lhdr->GetItem(j, &rHdr);
		rHdr.mask = HDI_FORMAT;
		rHdr.fmt = HDF_OWNERDRAW;
		lhdr->SetItem(j, &rHdr);
	}

	return 0;
}

void NListView::OnSize(UINT nType, int cx, int cy)
{
	CWnd::OnSize(nType, cx, cy);

	ResizeColumns();

	m_list.MoveWindow(0, 0, cx, cy);
}

void NListView::ResizeColumns()
{
	CRect rc;
	GetWindowRect(&rc);
	int w = rc.Width();
	int sb_width = GetSystemMetrics(SM_CXVSCROLL);
	w -= sb_width + 6;

	int col_zero_len = 22;
	int date_len = 100;
	int min_from_len = 150;
	int max_from_len = 400;
	int min_to_len = 150;
	int max_to_len = 400;
	int min_subj_len = 200;
	int dflt_subj_len = 400;
	int size_len = 60;

	int from_len = min_from_len;
	int to_len = min_to_len;
	int subj_len = min_subj_len;

	int min_len = col_zero_len + date_len + from_len + to_len + subj_len + size_len;
	int extra_space = w - min_len;


	if (extra_space > 0) {
		if (extra_space < (dflt_subj_len - min_subj_len)) {
			subj_len += extra_space;
		}
		else {
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
			SYSTEMTIME st;
			SYSTEMTIME lst;

			datebuff[0] = 0;
			if (m->m_timeDate > 0) 
			{
				CTime tt(m->m_timeDate);
				if (!m_gmtTime) {
					bool ret = tt.GetAsSystemTime(st);
					SystemTimeToTzSpecificLocalTime(0, &st, &lst);
					CTime ltt(lst);
					strcpy(datebuff, (LPCSTR)ltt.Format(m_format));
				}
				else {
					strcpy(datebuff, (LPCSTR)tt.Format(m_format));
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
	LPNMLVCUSTOMDRAW lplvcd = (LPNMLVCUSTOMDRAW)pNMHDR;
	NMCUSTOMDRAW   &nmcd = lplvcd->nmcd;
	MboxMail *m;

	DWORD dwItemSpec = nmcd.dwItemSpec;
	int iSubItem = lplvcd->iSubItem;
	DWORD dwItemType = lplvcd->dwItemType;
	DWORD dwDrawStage = nmcd.dwDrawStage;
	UINT uItemState = nmcd.uItemState;

	*pResult = 0;

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

		if ((iSubItem == 0) && (abs(MboxMail::b_mails_which_sorted) == 99))
		{
			if (dc.Attach(hDC))
			{
				if ((m->m_groupId % 2) == 0)
					dc.FillRect(&rect, &CBrush(PeachPuff1));
				else
					dc.FillRect(&rect, &CBrush(AntiqueWhite3));

				dc.SetBkMode(TRANSPARENT);
				dc.SetTextColor(RGB(0, 0, 0));

				dc.Detach();
			}

			int x_offset = 4;
			xpos = rect.left + x_offset;
			ypos = rect.top + 3;

			if (m->m_hasAttachments)
				FieldText = ast;
			else
				FieldText = nul;

			Charset = "UTF-8";
			charsetId = 65001;

			strW.Empty();
			if (Str2Wide(FieldText, charsetId, strW)) {
				::ExtTextOutW(hDC, xpos, ypos, ETO_CLIPPED, rect, (LPCWSTR)strW, strW.GetLength(), NULL);
			}
		}
		else if ((iSubItem >= 0) && (iSubItem <= 5))
		{
			if (iSubItem == 0)
			{
				if (m->m_hasAttachments)
					FieldText = ast;
				else
					FieldText = nul;

				Charset = "UTF-8";
				charsetId = 65001;

			}
			else if (iSubItem == 1)
			{
				char datebuff[32];
				SYSTEMTIME st;
				SYSTEMTIME lst;

				datebuff[0] = 0;
				if (m->m_timeDate > 0)
				{
					CTime tt(m->m_timeDate);
					if (!m_gmtTime) {
						bool ret = tt.GetAsSystemTime(st);
						SystemTimeToTzSpecificLocalTime(0, &st, &lst);
						CTime ltt(lst);
						strcpy(datebuff, (LPCSTR)ltt.Format(m_format));
					}
					else {
						strcpy(datebuff, (LPCSTR)tt.Format(m_format));
					}
				}
				FieldText = datebuff;
				Charset = "UTF-8";
				charsetId = 65001;
			}
			else if (iSubItem == 2) {
				FieldText = m->m_from;
				Charset = m->m_from_charset;
				charsetId = m->m_from_charsetId;
			}
			else if (iSubItem == 3) {
				FieldText = m->m_to;
				Charset = m->m_to_charset;
				charsetId = m->m_to_charsetId;
			}
			else if (iSubItem == 4) {
				FieldText = m->m_subj;
				Charset = m->m_subj_charset;
				charsetId = m->m_subj_charsetId;
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
				charsetId = 65001;
			}

			DWORD bkcolor = ::GetSysColor(COLOR_HIGHLIGHT);
			DWORD txcolor = ::GetSysColor(COLOR_HIGHLIGHTTEXT);

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
					dc.FillRect(&rect, &CBrush(RGB(255, 255, 255)));
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
			if (Str2Wide(FieldText, charsetId, strW))
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

	RemoveDir(GetmboxviewTempPath());
	MboxMail::Destroy();
	DestroyWindow();
	delete this;
}

int fileExists(LPCSTR file)
{
   WIN32_FIND_DATA FindFileData;
   HANDLE handle = FindFirstFile(file, &FindFileData) ;
   int found = handle != INVALID_HANDLE_VALUE;
   if(found) 
   {
       FindClose(handle);
   }
   return found;
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
	NListView::pCUPDUPData = pCUPDUPData;

	HANDLE h = GetCurrentThread();
	BOOL prio = SetThreadPriority(h, THREAD_PRIORITY_ABOVE_NORMAL);
	DWORD myThreadId = GetCurrentThreadId();
	DWORD myThreadPri = GetThreadPriority(h);
	TRACE(_T("threadId=%ld threadPriority=%ld\n"), myThreadId, myThreadPri);

	args->retpos = args->lview->DoFastFind(args->searchstart, FALSE, -1);
	args->exitted = TRUE;
	return true;
}

//#define CACHE_VERSION	10  go back to 10 once all done
#define CACHE_VERSION	11

void CPathStripPath(const char *path, CString &fileName)
{
	int pathlen = strlen(path);
	char *pathbuff = new char[pathlen + 1];
	strcpy(pathbuff, path);
	PathStripPath(pathbuff);
	fileName.Empty();
	fileName.Append(pathbuff);
	delete[] pathbuff;
}

BOOL SaveMails(LPCSTR cache) 
{
	CFile fpm;
	if (!fpm.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite)) {
		; // return -1;
	}

	CString mailFileName;
	CPathStripPath((char*)(LPCSTR)MboxMail::s_path, mailFileName);
	int pos = mailFileName.ReverseFind('.');
	CString baseFailFileName = mailFileName.Mid(0, pos);

	int ni = MboxMail::s_mails.GetSize();
	SerializerHelper sz(cache);
	if( ! sz.open(TRUE) )
		return FALSE;
	sz.writeInt(CACHE_VERSION);			// version
	sz.writeInt64(MboxMail::s_fSize = FileSize(MboxMail::s_path));	// file size
	sz.writeInt(ni);					// number of elements
	_int64 lastoff = 0;
	for (int i = 0; i < ni; i++) {
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

		int count = m->m_ContentDetailsArray.size();
		sz.writeInt(count);
		std::vector <MailBodyContent*>::iterator it;
		MailBodyContent *body;
		for (it = m->m_ContentDetailsArray.begin(); it != m->m_ContentDetailsArray.end(); it++)
		{
			body = *it;

			if (fpm.m_hFile != CFile::hFileNull)
			{
				if (!body->m_attachmentName.IsEmpty())
				{
					if ((body->m_contentDisposition.CompareNoCase("inline") == 0) && !body->m_contentId.IsEmpty())
					{
						SimpleString*outbuf = MboxMail::m_outbuf;
						outbuf->ClearAndResize(body->m_contentLength * 2);
						int retLen = MboxMail::DecodeBody(fpm, body, i, outbuf);

						if (outbuf->Count() > 1500000)
							int deb = 1;

						CString targerFolder = GetmboxviewTempPath((char*)(LPCSTR)baseFailFileName);
						CString mailIndex;
						mailIndex.Format("%d_", m->m_index);
						CString attachmentName = mailIndex + body->m_contentId;
						int pos = body->m_attachmentName.ReverseFind('.');
						CString imageExtension = body->m_attachmentName.Mid(pos);
						attachmentName += imageExtension;
						CString imgFile = targerFolder + attachmentName;

						body->m_attachmentName = attachmentName;

						if (outbuf->Count() > 1500000)
							int deb = 1;

						const char *fileName = (LPCSTR)imgFile;
						CFile fp(fileName, CFile::modeWrite | CFile::modeCreate);
						fp.Write(outbuf->Data(), outbuf->Count());
						fp.Close();

						int deb = 1;
					}
				}
			}
			sz.writeInt(body->m_contentOffset);
			sz.writeInt(body->m_contentLength);
			sz.writeString(body->m_contentTransferEncoding);
			sz.writeString(body->m_contentType);
			sz.writeString(body->m_contentId);
			sz.writeString(body->m_attachmentName);
			sz.writeString(body->m_contentDisposition);
			sz.writeInt(body->m_pageCode);
		}


		if (lastoff < m->m_startOff)
			lastoff = m->m_startOff;
	}
	TRACE("lastoff=%lld\n", lastoff);
	sz.close();
	return TRUE;
}

int LoadMails(LPCSTR cache, CArray<MboxMail*, MboxMail*> *mails = 0)
{
	SerializerHelper sz(cache);
	if (!sz.open(FALSE))
		return -1;
	int version;
	_int64 fSize;
	int ni = 0;
	if (!sz.readInt(&version))
		return -1;
	if (version != CACHE_VERSION) {
		CString txt = _T("File \"") + CString(cache);
		CString strVersion; 
		strVersion.Format(_T("%d"), version);
		txt += _T("\".\nhas incorrect version\"") + strVersion + "\". Expected version \"";
		strVersion.Format(_T("%d"), CACHE_VERSION);
		txt += strVersion + "\". Will remove\n\"" + CString(cache) + "\"\nand recreate from the mail file.";
		HWND h = NULL; // we don't have any window yet  
		int answer = MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return -1;
	}
	if( ! sz.readInt64(&fSize) )
		return -1;
	if( fSize != FileSize(MboxMail::s_path) )
		return -1;
	_int64 lastoff = 0;
	if( sz.readInt(&ni) ) {
		for( int i = 0; i < ni; i++ ) {
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
				if (!sz.readUInt(&body->m_pageCode))
					break;

				m->m_ContentDetailsArray.push_back(body);
			}

			if (lastoff < m->m_startOff)
				lastoff = m->m_startOff;

			if (mails == 0)
				MboxMail::s_mails.Add(m);
			else
				mails->Add(m);
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
	if (fSize != FileSize(MboxMail::s_path))
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

			if (lastoff < m->m_startOff)
				lastoff = m->m_startOff;

			datebuff[0] = 0;
			if (m->m_timeDate > 0) {
				CTime tt(m->m_timeDate);
				strcpy(datebuff, (LPCSTR)tt.Format(format));
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
	NTreeView *pView = pFrame->GetTreeView();
	if( ! pView )
		return;
	CString cache= m_path + ".mboxview";
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
		CArray<MboxMail*, MboxMail*> s_mails_new;
		if (fileExists(cache_new)) {
			MboxMail::s_path = m_path;
			ni_new = LoadMails(cache_new, &s_mails_new);
			if (ni_new < 0) {
				ni_new = 0;
			}
		}
		CString cache_old = cache + ".old";
		CArray<MboxMail*, MboxMail*> s_mails_old;
		if (fileExists(cache_old)) {
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

	if( fileExists(cache) ) {
		MboxMail::s_path = m_path;
		// it populates s_mails from mail index/mboxview file
		ni = LoadMails(cache);
		if( ni < 0 ) {
			ni = 0;
			DeleteFile(cache);
		} else
			m_list.SetItemCount(ni);
	}
	if( ! fileExists(cache) ) {
		PARSE_ARGS args;
		args.path = m_path;
		args.exitted = FALSE;
		CUPDialog	Dlg(AfxGetMainWnd()->GetSafeHwnd(), ALongRightProcessProc, (LPVOID)(PARSE_ARGS*)&args);
		if( !Dlg.DoModal() )
			return;
		if (!args.exitted)
		{
			DWORD tc_start = GetTickCount();
			while (args.exitted == FALSE) {
				Sleep(1);
			}
			DWORD tc_end = GetTickCount();
			DWORD delta = tc_end - tc_start;
			TRACE("(FillCtrl)Waited %ld milliseconds for thread to exist.\n", delta);
			Sleep(5);
		}

		bool ret = MboxMail::preprocessConversations();
		ret = MboxMail::sortConversations();
		//MboxMail::assignColor2ConvesationGroups();

		// creates to index/mboxview file from s_mails
		if( SaveMails(cache) ) {
			ni = MboxMail::s_mails.GetSize();
			m_list.SetItemCount(ni);
			//Cache2Text(cache, m_format);
		}
	}

	//MboxMail::assignColor2ConvesationGroups();
	MboxMail::s_mails_ref.SetSize(MboxMail::s_mails.GetSize());
	MboxMail::s_mails_ref.Copy(MboxMail::s_mails);
	MboxMail::b_mails_sorted = true;
	MboxMail::b_mails_which_sorted = 1;
	MboxMail::SortByDate();
	MarkColumns();
	m_bEditFindFirst = TRUE;

#ifdef _DEBUG
	tc = (GetTickCount() - tc);
	TRACE("FillCtrl Took %d:%d %d\n", (tc/1000)/60, (tc/1000)%60, tc);
#endif
	m_list.EnsureVisible(ni, FALSE);
	m_list.SetRedraw(TRUE);
	//pView->m_tree.SetItemData(m_which, (DWORD)FileSize(m_path));
	//BOOL retval = pView->m_tree.SetItemState(m_which, 0, TVIS_BOLD);
	BOOL retval = pView->m_tree.SetItemState(m_which, TVIS_SELECTED, TVIS_BOLD);
	retval = pView->m_tree.SelectItem(m_which);
	CString txt = pView->m_tree.GetItemText(m_which);
	pView->SaveData();
}

int fixInlineSrcImgPath(char *inData, int indDataLen, SimpleString *outbuf, CListCtrl *attachments, int mailPosition, bool useMailPosition)
{
	static char * img_pattern = "<img ";
	static int img_patternLen = strlen(img_pattern);

	static char * cid_pattern = "src=\"cid:";
	static int cid_patternLen = strlen(cid_pattern);

	char *input = inData;
	int inputLength = indDataLen;
	char *inputEnd = input + inputLength;
	char *srcImgEnd = inputEnd;
	char *cidBegin;
	char *cidEnd;
	SimpleString cid;

	char *fromBegin = input;
	char *alt_pos;

	MboxMail *m = 0;
	if (useMailPosition)
		m = MboxMail::s_mails[mailPosition];

	CString mailArchiveFileName;
	CPathStripPath((char*)(LPCSTR)MboxMail::s_path, mailArchiveFileName);
	int position = mailArchiveFileName.ReverseFind('.');
	CString baseFileArchiveName = mailArchiveFileName.Mid(0, position);

	CString imgFile;

	char *pos = input;
	while ((pos != 0) && (pos < inputEnd)) 
	{
		pos = strnstrUpper2Lower(pos, inputEnd, img_pattern, img_patternLen);
		if (pos == 0) {
			outbuf->Append(fromBegin, inputEnd - fromBegin);
			break;
		}

		pos += img_patternLen;

		// find end of <img src ...>
		srcImgEnd = strnstrUpper2Lower(pos, inputEnd, ">", 1);
		if (srcImgEnd == 0) { // TODO: corrupted file ?
			outbuf->Append(fromBegin, inputEnd - fromBegin);
			break;
		}
		srcImgEnd++; // jump over

		alt_pos = strnstrUpper2Lower(pos, srcImgEnd, "alt=\"", 5);

		pos = strnstrUpper2Lower(pos, srcImgEnd, cid_pattern, cid_patternLen);
		if (pos == 0) {
			outbuf->Append(fromBegin, srcImgEnd - fromBegin);
			pos = srcImgEnd;  // jump over this <img .. > sonce no cid:
			fromBegin = pos;
			continue;
		}

		pos += cid_patternLen;
		cidBegin = pos;

		outbuf->Append(fromBegin, cidBegin - 5 - fromBegin);

		// find end of cid token
		pos = strnstrUpper2Lower(pos, srcImgEnd, "\"", 1);
		cidEnd = pos;

		cid.Copy(cidBegin, cidEnd - cidBegin);

		pos++; // jump over \"
		cidEnd = pos;
		fromBegin = pos;

		CString nameExtension;
		if (m == 0) {
			for (int row = 0; row < attachments->GetItemCount(); ++row)
			{
				CString name = attachments->GetItemText(row, 0);
				int pos = name.ReverseFind('.');
				CString cidPart = name.Mid(0, pos);
				if (cidPart.Compare(cid.Data()) == 0) {
					nameExtension = name.Mid(pos);
					break;
				}
			}
			imgFile = GetmboxviewTempPath() + CString(cid.Data()) + nameExtension;
		}
		else
		{
			MailBodyContent *body;
			for (int j = 0; j < m->m_ContentDetailsArray.size(); j++)
			{
				body = m->m_ContentDetailsArray[j];
				if (!body->m_attachmentName.IsEmpty())
				{
					if ((body->m_contentDisposition.CompareNoCase("inline") == 0) && !body->m_contentId.IsEmpty())
					{
						if (outbuf->Count() > 1500000)
							int deb = 1;

						int pos = body->m_attachmentName.ReverseFind('.');
						CString cidPart = body->m_attachmentName.Mid(0, pos);

						CString extendedCid;
						extendedCid.Format("%d_%s", m->m_index, cid.Data());

						if (cidPart.Compare(extendedCid) == 0) {
							nameExtension = body->m_attachmentName.Mid(pos);
							break;
						}

						if (outbuf->Count() > 1500000)
							int deb = 1;

						int deb = 1;
					}
				}
			}
			CString targerFolder = GetmboxviewTempPath((char*)(LPCSTR)baseFileArchiveName);
			imgFile = targerFolder + body->m_attachmentName;
		}

		if (nameExtension.IsEmpty()) {
			;// TODO: any harm if no action ?
		}

		SimpleString alt;
		if (alt_pos == 0) {
			alt.Append("alt=\"image");
			alt.Append((char*)(LPCSTR)nameExtension);
			alt.Append("\"");
			int deb = 1;
		}

		outbuf->Append("\"file:\\\\\\");
		outbuf->Append(imgFile, imgFile.GetLength());
		outbuf->Append("\"");

		if (alt.Count()) {
			outbuf->Append(" ");
			outbuf->Append(alt);
		}

		pos = srcImgEnd;

		int deb = 1;
	}
	return -1;
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

void NListView::SelectItem(int iItem) 
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if( ! pFrame )
		return;
	NMsgView *pView = pFrame->GetMsgView();
	if( ! pView )
		return;
	pView->m_bAttach = FALSE;
	// Sanity check
	if( iItem < 0 || iItem >= MboxMail::s_mails.GetSize() ) {
		return;
		//ClearDescView();
		return;
	}
	m_lastSel = iItem;
	// Erase any files previously saved
	RemoveDir(GetmboxviewTempPath());
	pView->m_attachments.DeleteAllItems();

	// Get cached mail
	MboxMail *m = MboxMail::s_mails[iItem];

	// Set header data
	pView->m_strSubject = m->m_subj;
	pView->m_strFrom = m->m_from;
	//CTime tt(m->m_timeDate);
	//pView->m_strDate = tt.Format(m_format);
	if (m->m_timeDate > 0)
	{
		SYSTEMTIME st;
		SYSTEMTIME lst;
		CTime tt(m->m_timeDate);
		if (!m_gmtTime) {
			bool ret = tt.GetAsSystemTime(st);
			SystemTimeToTzSpecificLocalTime(0, &st, &lst);
			CTime ltt(lst);
			pView->m_strDate = ltt.Format(m_format);
		}
		else {
			pView->m_strDate = tt.Format(m_format);
		}
	}
	else
		pView->m_strDate = "";

	pView->m_strTo = m->m_to;
	pView->m_subj_charsetId = m->m_subj_charsetId;
	pView->m_subj_charset = m->m_subj_charset;
	pView->m_from_charsetId = m->m_from_charsetId;
	pView->m_from_charset = m->m_from_charset;
	pView->m_date_charsetId = 0;
	pView->m_to_charsetId = m->m_to_charsetId;
	pView->m_to_charset = m->m_to_charset;

	// Get raw mail body
	CString bdy = m->GetBody();

	// Save raw message
	if (m_bExportEml)
	{
	// Save mail
		CFile fp(GetmboxviewTempPath()+"message.eml", CFile::modeWrite|CFile::modeCreate);
		fp.Write(bdy, bdy.GetLength());
		fp.Close();
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
    CMimeBody::CBodyList::const_iterator it;
	bdy = "";
	CString ext = "";
	if( bodies.begin() == bodies.end() ) {  // should never be true
		//bdy = mail.GetContent();
		string strText;
		mail.GetText(strText);
		bdy = strText.c_str();
		ext = "txt";
	} else
    for (it=bodies.begin(); it!=bodies.end(); it++)
    {
        CMimeBody* pBP = *it;
		CString curExt = "txt";

        // Iterate all the header fields of this body part:
        CMimeHeader::CFieldList& fds = pBP->Fields();
        CMimeHeader::CFieldList::const_iterator itfd;
		CString charset;
        for (itfd=fds.begin(); itfd!=fds.end(); itfd++)
        {
            const CMimeField& fd = *itfd;
			const char *fname = fd.GetName();
			const char *fval = fd.GetValue();
			// Check content type to get mail extension
			if( _stricmp(fname, "Content-Type") == 0 && _strnicmp(fval, "text/", 5) == 0 ) {
				const char *p = fd.GetValue()+5;
				if( _strnicmp(p, "plain", 5) == 0 )
					curExt = "txt";
				else
				if( _strnicmp(p, "htm", 3) == 0 )
					curExt = "htm";
				else
				if( _strnicmp(p, "xml", 3) == 0 )
					curExt = "xml";
				else
					curExt = "txt";
				int pc = CString(fval).Find("charset=");
				if (pc != -1) {
					int charsetLength = CString(fval + pc + 8).FindOneOf(";\n\r");
					if (charsetLength < 0)
						charsetLength = strlen(fval + pc + 8);
					strncpy(charset.GetBufferSetLength(charsetLength), fval + pc + 8, charsetLength);
					charset.Trim("\"\\");
				}
			}
        }

        if ((pBP->IsText() || pBP->IsMessage()) && (bdy.IsEmpty() || !pBP->IsAttachment()) )
        {
			// if message contains alternate parts display last one
			string strText;
			pBP->GetText(strText);
			bdy = strText.c_str();
			ext = curExt;
			bdycharset = charset;
			TRACE("ext=%s charset=%s\n", (LPCSTR)ext, (LPCSTR)charset);
        }
        else if (pBP->IsAttachment())
        {
			// Save all attachments
            string strName = pBP->GetName();
			if (strName.empty())
				strName = pBP->GetFilename();

			CString disposition = pBP->GetDisposition();
			const CMimeField *pfldDisposition = pBP->CMimeHeader::GetField(CMimeConst::ContentDisposition());
			string strDisposition;
			if (pfldDisposition != NULL)
			{
				pfldDisposition->GetValue(strDisposition);
				transform(strDisposition.begin(), strDisposition.end(), strDisposition.begin(), ::tolower);
			}

			const CMimeField *pfldContentId = pBP->CMimeHeader::GetField(CMimeConst::ContentID());
			string strContentId;
			if (pfldContentId != NULL)
			{
				pfldContentId->GetValue(strContentId);
			}

			bool isAttachmentInline = false;
			if (strDisposition.compare("inline") == 0) {
				isAttachmentInline = true;
				hasInlineAttachments = true;
				CString contentId = strContentId.c_str();
				contentId.Trim();
				contentId.Trim("<>");

				if (!contentId.IsEmpty()) {
					CString cStrName = strName.c_str();
					int pos = cStrName.ReverseFind('.');
					CString nameExtension = cStrName.Mid(pos);
					CString ext = PathFindExtension(cStrName);
					CString newStrName = contentId + nameExtension;
					strName.assign(newStrName);

					int deb = 1;
				}
				int deb = 1;
			}

            //printf("File name: %s\r\n", strName.c_str());
            //printf("File size: %d\r\n", pBP->GetContentLength());
			string name = strName;
            strName = string((LPCSTR)GetmboxviewTempPath()) + strName;
            pBP->WriteToFile(strName.c_str());
			// Set item icon and insert in attachment list
			int iIcon = 0;
			SHFILEINFO shFinfo;
			if ( !SHGetFileInfo( strName.c_str(),
						0,
						&shFinfo,
						sizeof( shFinfo ),
						SHGFI_ICON | 
						SHGFI_SMALLICON ) )
			{
				TRACE("Error Gettting SystemFileInfo!\n");
			} else {
				iIcon = shFinfo.iIcon;
				// we only need the index from the system image list
				DestroyIcon( shFinfo.hIcon );
			}
			pView->m_bAttach = TRUE;
			pView->m_attachments.InsertItem(pView->m_attachments.GetItemCount(), name.c_str(), iIcon);
        }
    }

	pView->m_attachments.SortItemsEx(MyCompareProc, (LPARAM)&pView->m_attachments);

	// Save mail
	if (ext.Compare("txt") == 0) {
		bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body><pre>\r\n" + bdy + "</pre></body></html>";
		ext = "htm";
		data = (char*)(LPCSTR)bdy;
		datalen = bdy.GetLength();
	}
	else if (ext.Compare("htm") == 0 && CString(bdy).MakeLower().Find("<body") == -1) 
	{
		if (hasInlineAttachments) {
			SimpleString *outbuf = MboxMail::m_outbuf;
			outbuf->ClearAndResize(bdy.GetLength() + 1000);

			CString hdr = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body>";
			outbuf->Append((LPCSTR)hdr, hdr.GetLength());

			int mailPosition = 0; // not used anyway here
			bool useMailPosition = false;
			char *inData = (char*)(LPCSTR)bdy;
			int inDataLen = bdy.GetLength();
			fixInlineSrcImgPath(inData, inDataLen, outbuf, &pView->m_attachments, mailPosition, useMailPosition);

			hdr = "</body></html>";
			outbuf->Append((LPCSTR)hdr, hdr.GetLength());
			data = outbuf->Data();
			datalen = outbuf->Count();
			int deb = 1;
		}
		else {
			bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body>" + bdy +"</body></html>";
			data = (char*)(LPCSTR)bdy;
			datalen = bdy.GetLength();
		}
	}
	else {
		// TODO: need to optimize
		BOOL charsetMissing = FALSE;
		CString bdyLower(bdy);
		bdyLower.MakeLower();
		if ((ext.Compare("htm") == 0) && (bdyLower.Find("charset=") == -1))
		{
			//int pos = bdy.Find("<body>");
			if (bdycharset.MakeLower().CompareNoCase("utf-8") != 0)
				charsetMissing = TRUE;  // simplify for now, update html doc later ?
			int deb = 1;
		}
		if (hasInlineAttachments) {
			// TODO: handle charsetMissing
			SimpleString *outbuf = MboxMail::m_outbuf;
			outbuf->ClearAndResize(bdy.GetLength() + 1000);

			int mailPosition = 0; // not used anyway here
			bool useMailPosition = false;
			char *inData = (char*)(LPCSTR)bdy;
			int inDataLen = bdy.GetLength();
			fixInlineSrcImgPath(inData, inDataLen, outbuf, &pView->m_attachments, mailPosition, useMailPosition);

			data = outbuf->Data();
			datalen = outbuf->Count();
		}
		else if (charsetMissing) {
			SimpleString *outbuf = MboxMail::m_outbuf;
			outbuf->ClearAndResize(bdy.GetLength() + 1000);

			CString hdr = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=" + bdycharset + "\"></head><body>";
			outbuf->Append((LPCSTR)hdr, hdr.GetLength());

			data = (char*)(LPCSTR)bdy;
			datalen = bdy.GetLength();

			outbuf->Append(data, datalen);

			data = outbuf->Data();
			datalen = outbuf->Count();
		}
		else {
			// TODO: handle charsetMissing
			data = (char*)(LPCSTR)bdy;
			datalen = bdy.GetLength();
		}
		int deb = 1;
	}

	pView->m_body_charsetId = charset2Id(bdycharset);
	pView->m_body_charset = bdycharset;

	// Get temporary file name with correct extension for IE to display
	m_curFile = CreateTempFileName(ext);
	CFile fp(m_curFile, CFile::modeWrite|CFile::modeCreate);
	fp.Write(data, datalen);
	fp.Close();
	
	// Display mail in IE
	pView->m_browser.Navigate(m_curFile, NULL);
	// Update layou to show/hide attachments
	pView->UpdateLayout();
	return;
}

// Dump header fields based on MboxMail *m and corresponding data from mime decoder
int NListView::DumpItemDetails(MboxMail *m)
{
	int iItem = 1;

	// Get raw mail body
	CString bdy = m->GetBody();

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
	CString bdy = m->GetBody();

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
		pos = FileSeek(hFile, fsize, FILE_BEGIN);
		if (pos < 0) {
			DWORD err = GetLastError();
			CloseHandle(hFile);
			return -1;
		}
	}

	datebuff[0] = 0;
	if (m->m_timeDate > 0) {
		CTime tt(m->m_timeDate);
		CString format = "";
		strcpy(datebuff, (LPCSTR)tt.Format(format));
	}

	count = sprintf_s(buff, "INDX=%d first=%lld len=%d last=%lld att=%d hlen=%d rcv=%d date=\"%s\" from=\"%s\" to=\"%s\" subj=\"%s\"\n\n",
		iItem, m->m_startOff, m->m_length, (m->m_startOff + m->m_length - 1), m->m_hasAttachments, m->m_headLength, m->m_recv,
		datebuff, (LPCSTR)m->m_from, (LPCSTR)m->m_to, (LPCSTR)m->m_subj);

	nwritten = 0;
	if (!WriteFile(hFile, buff, count, &nwritten, NULL)) {
		DWORD retval = GetLastError();
	}

	const char *cc = mail.GetCc();
	const char *bcc = mail.GetBcc();
	count = sprintf_s(buff, "INDX=%d date=\"%s\" from=\"%s\" from_charset=\"%s\" to=\"%s\" to_charset=\"%s\" cc=\"%s\" bcc=\"%s\" subj=\"%s\" to_charset=\"%s\"\n\n",
		iItem, mail.GetDate(), (LPCSTR)mail.GetFrom(), mail.GetFieldCharset("From"), (LPCSTR)mail.GetTo(), mail.GetFieldCharset("To"),
		cc, bcc, (LPCSTR)mail.GetSubject(), mail.GetFieldCharset("Subject"));

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

	count = sprintf_s(buff, "BodyIndx=%d IsText=%d IsMessage=%d HasBody=%d IsAttachement=%d IsMultiPart=%d ContentLength=%d "
		"Charset=%s Description=%s Disposition=%s TransferEncoding=%s SubType=%s MainType=%s "
		"Boundary=%s ContentType=%s MediaType=%d Name=%s\n",
		indx, pBP->IsText(), pBP->IsMessage(), !bdy.IsEmpty(), pBP->IsAttachment(), pBP->IsMultiPart(), contentLength,
		pBP->GetCharset().c_str(), pBP->GetDescription(), pBP->GetDisposition(), pBP->GetTransferEncoding(), pBP->GetSubType().c_str(), pBP->GetMainType().c_str(),
		pBP->GetBoundary().c_str(), pBP->GetContentType(), pBP->GetMediaType(), pBP->GetName().c_str());
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
		const char *fname = fd.GetName();
		const char *fval = fd.GetValue();
		const char *charset = fd.GetCharset();
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

			count = sprintf_s(buff, "BodyIndx=%d IsText=%d IsMessage=%d HasBody=%d IsAttachement=%d IsMultiPart=%d ContentLength=%d "
				"Charset=%s Description=%s Disposition=%s TransferEncoding=%s SubType=%s MainType=%s "
				"Boundary=%s ContentType=%s MediaType=%d Name=%s\n",
				indx, pBP->IsText(), pBP->IsMessage(), !bdy.IsEmpty(), pBP->IsAttachment(), pBP->IsMultiPart(), contentLength, 
				pBP->GetCharset().c_str(), pBP->GetDescription(), pBP->GetDisposition(), pBP->GetTransferEncoding(), pBP->GetSubType().c_str(), pBP->GetMainType().c_str(),
				pBP->GetBoundary().c_str(), pBP->GetContentType(), pBP->GetMediaType(), pBP->GetName().c_str());
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
				const char *fname = fd.GetName();
				const char *fval = fd.GetValue();
				const char *charset = fd.GetCharset();
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
		return CWnd::OnMouseWheel(nFlags, zDelta, pt);
	}
	return wnd->PostMessage(WM_MOUSEWHEEL, MAKELPARAM(nFlags,zDelta), MAKELPARAM(pt.x, pt.y));
}

void NListView::ClearDescView()
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if( ! pFrame )
		return;
	NMsgView *pView = pFrame->GetMsgView();
	if( ! pView )
		return;
	pView->m_strSubject.LoadString(IDS_DESC_NONE);
	pView->m_strFrom.LoadString(IDS_DESC_NONE);
	pView->m_strDate.LoadString(IDS_DESC_NONE);
	pView->m_strTo.LoadString(IDS_DESC_NONE);
	pView->m_strTo.LoadString(IDS_DESC_NONE);
	pView->m_subj_charset.SetString(_T(""));
	pView->m_from_charset.SetString(_T(""));
	pView->m_date_charset.SetString(_T(""));
	pView->m_to_charset.SetString(_T(""));
	pView->m_body_charset.SetString(_T(""));
	pView->m_body_charsetId = 0;
	pView->m_attachments.DeleteAllItems();
	RemoveDir(GetmboxviewTempPath());
	m_curFile.Empty();
	pView->m_browser.Navigate("about:blank", NULL);
	pView->UpdateLayout();
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
	CTime ct(st);
	return ct.GetTime();
}

void NListView::OnEditFind()
{
	if (m_bInFind)
		return;
	m_bInFind = true;
	CFindDlg dlg;

	dlg.m_bWholeWord = m_bWholeWord;
	dlg.m_bCaseSensitive = m_bCaseSens;
	dlg.m_string = m_searchString;
	if (m_filterDates == FALSE)
		m_lastStartDate = 0;
	if (m_lastStartDate <= 0) {
		time_t min = time(NULL);
		time_t max = 0;
		int sz = MboxMail::s_mails.GetSize();
		for (int i = 0; i < sz; i++) {
			time_t t = MboxMail::s_mails[i]->m_timeDate;
			if (t < min && t > 0)
				min = t;
			if (t > max)
				max = t;
		}
		m_lastStartDate = CTime(min);
		if (max == 0)
			m_lastEndDate = CTime(CTime::GetCurrentTime());
		else
			m_lastEndDate = CTime(max);
	}
	dlg.m_startDate = COleDateTime(m_lastStartDate.GetTime());
	dlg.m_endDate = COleDateTime(m_lastEndDate.GetTime());
	dlg.m_bFindNext = m_bFindNext;
	dlg.m_bFrom = m_bFrom;
	dlg.m_bTo = m_bTo;
	dlg.m_bSubject = m_bSubject;
	dlg.m_bContent = m_bContent;
	dlg.m_bAttachments = m_bAttachments;
	dlg.m_bHighlightAll = m_bHighlightAll;
	dlg.m_filterDates = m_filterDates;

	if (dlg.DoModal() == IDOK) {
		m_filterDates = dlg.m_filterDates;
		m_searchString = dlg.m_string;
		m_bWholeWord = dlg.m_bWholeWord;
		m_bCaseSens = dlg.m_bCaseSensitive;
		m_lastStartDate = CTime(OleToTime_t(&dlg.m_startDate));
		m_lastEndDate = CTime(OleToTime_t(&dlg.m_endDate));
		m_bFindNext = dlg.m_bFindNext;
		m_bFrom = dlg.m_bFrom;
		m_bTo = dlg.m_bTo;
		m_bSubject = dlg.m_bSubject;
		m_bContent = dlg.m_bContent;
		m_bAttachments = dlg.m_bAttachments;
		m_bHighlightAll = dlg.m_bHighlightAll;

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

		m_bEditFindFirst = FALSE;

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
			if (m_maxSearchDuration > 0)
				w = DoFastFind(which, TRUE, m_maxSearchDuration);
			FIND_ARGS args;
			/*IN*/ args.lview = this; args.searchstart = m_lastFindPos;
			/*OUT*/ args.exitted = FALSE; args.retpos = -1;
			if ((w == -2) || (m_maxSearchDuration == 0)) {
				CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcFastSearch, (LPVOID)(FIND_ARGS*)&args);
				if (!Dlg.DoModal())
					return;
				if (!args.exitted)
				{
					DWORD tc_start = GetTickCount();
					while (args.exitted == FALSE) {
						Sleep(1);
					}
					DWORD tc_end = GetTickCount();
					DWORD delta = tc_end - tc_start;
					TRACE("(OnEditFind)Waited %ld milliseconds for thread to exist.\n", delta);
					Sleep(5);
				}
				w = args.retpos;
			}
			if (w >= 0) {
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
}

int NListView::CheckMatch(int i, CString &searchString)
{
	int pos = -1;
	int w = -1;
	MboxMail *m = MboxMail::s_mails[i];

	if (searchString.GetLength() == 1 && searchString[0] == '*') {
		w = i;
		return w;
	}
	pos = -1;
	if (m_bSubject) {
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
		if (m_bWholeWord)
			pos = g_tu.BMHSearchW((unsigned char *)(LPCSTR)m->m_to, m->m_to.GetLength(), (unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_bCaseSens);
		else
			pos = g_tu.BMHSearch((unsigned char *)(LPCSTR)m->m_to, m->m_to.GetLength(), (unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_bCaseSens);
		if (pos >= 0) {
			w = i;
			return w;
		}
	}
	if (m_bContent || m_bAttachments) {
		if (FindInMailContent(i, m_bContent, m_bAttachments)) {
			w = i;
			return w;
		}
	}
	return w;
}


int NListView::DoFastFind(int which, BOOL mainThreadContext, int maxSearchDuration)
{
	int w = -1;
	int sz = MboxMail::s_mails.GetSize();
	time_t sd = m_lastStartDate.GetTime(), ed = m_lastEndDate.GetTime();
	CString searchString(m_searchString);
	if (m_bCaseSens == 0)
		searchString.MakeLower();

	if (!mainThreadContext) {
		CString searchText = _T("Searching file for \"") + searchString + _T("\" ...");
		pCUPDUPData->SetProgress(searchText, 0);
	}

	DWORD myThreadId = GetCurrentThreadId();
	DWORD tc_start = GetTickCount();

	if (m_bFindNext)
	{
		for (int i = which; i < sz; i++)
		{
			MboxMail *m = MboxMail::s_mails[i];
			bool process = false;
			if (m_filterDates) {
				if (m->m_timeDate >= sd && m->m_timeDate <= ed)
					process = true;
			}
			else
				process = true;
			if (process)
			{
				if (CheckMatch(i, searchString) >= 0) {
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
						pCUPDUPData->SetProgress(step);
					}

					if (pCUPDUPData->ShouldTerminate()) {
						w = -1;
						break;
					}
				}
			}
		}
	}
	else
	{
		for (int i = which; i >= 0; i--)
		{
			MboxMail *m = MboxMail::s_mails[i];
			bool process = false;
			if (m_filterDates) {
				if (m->m_timeDate >= sd && m->m_timeDate <= ed)
					process = true;
			}
			else
				process = true;
			if (process) 
			{
				if (CheckMatch(i, searchString) >= 0) {
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
						pCUPDUPData->SetProgress(step);
					}

					if (pCUPDUPData->ShouldTerminate()) {
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
	if( m_searchString.IsEmpty() ) {
		OnEditFind();
		return;
	}
	if (m_bEditFindFirst == TRUE) {
		OnEditFind();
		return;
	}

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
	if (m_maxSearchDuration > 0)
		w = DoFastFind(which, TRUE, m_maxSearchDuration);

	FIND_ARGS args;
	/*IN*/ args.lview = this; args.searchstart = m_lastFindPos;
	/*OUT*/ args.exitted = FALSE; args.retpos = -1;
	if ((w == -2) || (m_maxSearchDuration == 0)) {
		CUPDialog	Dlg(GetSafeHwnd(), ALongRightProcessProcFastSearch, (LPVOID)(FIND_ARGS*)&args);
		if (!Dlg.DoModal())
			return;
		if (!args.exitted)
		{
			DWORD tc_start = GetTickCount();
			while (args.exitted == FALSE) {
				Sleep(1);
			}
			DWORD tc_end = GetTickCount();
			DWORD delta = tc_end - tc_start;
			TRACE("(OnEditFind)Waited %ld milliseconds for thread to exist.\n", delta);
			Sleep(5);
		}
		w = args.retpos;
	}
	if (w >= 0) {
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
		pos = FileSeek(hFile, fsize, FILE_BEGIN);
		if (pos < 0) {
			DWORD err = GetLastError();
			CloseHandle(hFile);
			CloseHandle(mbox_hFile);
			return -1;
		}
	}

	MboxMail *m = MboxMail::s_mails[which];

	datebuff[0] = 0;
	if (m->m_timeDate > 0) {
		CTime tt(m->m_timeDate);
		strcpy(datebuff, (LPCSTR)tt.Format(m_format));
	}

	count = sprintf_s(buff, "INDX=%d first=%lld len=%d last=%lld att=%d hlen=%d rcv=%d date=\"%s\" from=\"%s\" to=\"%s\" subj=\"%s\"\n\n",
		which, m->m_startOff, m->m_length, (m->m_startOff + m->m_length - 1), m->m_hasAttachments, m->m_headLength, m->m_recv,
		datebuff, (LPCSTR)m->m_from, (LPCSTR)m->m_to, (LPCSTR)m->m_subj);

	nwritten = 0;
	if (!WriteFile(hFile, buff, count, &nwritten, NULL)) {
		DWORD retval = GetLastError();
	}
	long long start_offset = m->m_startOff;
	pos = FileSeek(mbox_hFile, start_offset, FILE_BEGIN);

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
	while (pos)
	{
		int nItem = m_list.GetNextSelectedItem(pos);
		m_list.SetItemState(nItem, 0, LVIS_SELECTED | LVIS_FOCUSED);
	}
	m_list.SetItemState(which, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
	m_list.EnsureVisible(which, FALSE);
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
			CString bdy = m->GetBody();
			// Save mail
			CFile fp(GetmboxviewTempPath() + "message.eml", CFile::modeWrite | CFile::modeCreate);
			fp.Write(bdy, bdy.GetLength());
			fp.Close();
		}
	}

	CString path = GetmboxviewTempPath();
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
	NTreeView *pView = pFrame->GetTreeView();
	if (!pView)
		return;

	int ni = MboxMail::s_mails.GetSize();
	m_list.SetItemCount(ni);

	m_list.EnsureVisible(ni, FALSE);
	m_list.SetRedraw(TRUE);
	//pView->m_tree.SetItemData(m_which, (DWORD)FileSize(m_path));
	//BOOL retval = pView->m_tree.SetItemState(m_which, 0, TVIS_BOLD);
	BOOL retval = pView->m_tree.SetItemState(m_which, LVIS_SELECTED, TVIS_BOLD);
	retval = pView->m_tree.SelectItem(m_which);
	CString txt = pView->m_tree.GetItemText(m_which);
	pView->SaveData();
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

BOOL NListView::SetupFileMapView(_int64 offset, DWORD length)
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
	if (m_bFindNext == TRUE)
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
	if (m_pMapViewBegin) 
		UnmapViewOfFile(m_pMapViewBegin);
	if (m_hMailFileMap != INVALID_HANDLE_VALUE)
		CloseHandle(m_hMailFileMap);
	if (m_hMailFile != INVALID_HANDLE_VALUE)
		CloseHandle(m_hMailFile);

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
	if (SetupFileMapView(m->m_startOff, m->m_length) == FALSE)
		return FALSE;

	MailBodyContent *body;
	for (int j = 0; j < m->m_ContentDetailsArray.size(); j++) 
	{
		body = m->m_ContentDetailsArray[j];

		if (!body->m_attachmentName.IsEmpty()) {
			if (bAttachment == FALSE)
				continue;
		}
		else if (bContent == FALSE)
			continue;

		if (body->m_contentType.CompareNoCase("text/plain") != 0)
			continue;

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
				pos = (_int64)g_tu.BMHSearchW((unsigned char *)pData, datalen, (unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_bCaseSens);
			else
				pos = (_int64)g_tu.BMHSearch((unsigned char *)pData, datalen, (unsigned char *)(LPCSTR)searchString, searchString.GetLength(), m_bCaseSens);
			if (pos >= 0) {
				return TRUE;
			}
		}
	}
	return FALSE;
}


#if 0
void NListView::OnViewConversations()
{
	// SendMessage to simulate Column Click ??, need to investigate

	// to follow OnEditFindAgain approach
	if (m_bInFind)
		return;

	m_bInFind = true;

	// TODO: Add your command handler code here
	if (abs(MboxMail::b_mails_which_sorted) == 99) {
		MboxMail::b_mails_which_sorted = -MboxMail::b_mails_which_sorted;
	}
	else {
		MboxMail::b_mails_which_sorted = 99;
	}
	MboxMail::SortByConverstionGroups(0, MboxMail::b_mails_which_sorted < 0);

	MboxMail::b_mails_sorted = true;

	m_bContent = FALSE;
	m_bAttachments = FALSE;
	m_lastFindPos = -1;
	m_bEditFindFirst = TRUE; 
	RedrawMails();

	m_bInFind = false;
}



void NListView::OnUpdateViewConversations(CCmdUI *pCmdUI)
{
	pCmdUI->Enable(MboxMail::s_mails.GetSize() > 0);
	//pCmdUI->Enable(m_list.GetItemCount() > 0);
}
#endif

void NListView::PrintMailGroupToText(int iItem, int textType, BOOL forceOpen)
{
	if (abs(MboxMail::b_mails_which_sorted) != 99) {

		CString txt = _T("Please sort all mails by conversation first.\n");
		txt += "Select \"View\"->\"Sort By\" ->\"Conversation\" or left click on the first column.";
		HWND h = GetSafeHwnd(); // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;
	}
	MboxMail *m = 0;
	m = MboxMail::s_mails[iItem];
	int groupId = m->m_groupId;
	int firstMail = iItem;
	int lastMail = iItem;
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

	TEXTFILE_CONFIG textConfig;
	textConfig.m_dateFormat = m_format;
	textConfig.m_bGMTTime = m_gmtTime;
	textConfig.m_nCodePageId = CP_UTF8;

	CString textFileName;
	int ret = 0;
	ret = MboxMail::exportToTextFile(textConfig, textFileName, firstMail, lastMail, textType);
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

int NListView::MailsWhichColumnSorted() const
{
	return MboxMail::b_mails_which_sorted;
}
