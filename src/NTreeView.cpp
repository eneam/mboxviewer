// NTreeView.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "MboxMail.h"
#include "NTreeView.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(NTreeView, CWnd)

/////////////////////////////////////////////////////////////////////////////
// NTreeView

NTreeView::NTreeView()
{
}

NTreeView::~NTreeView()
{
}


BEGIN_MESSAGE_MAP(NTreeView, CWnd)
	//{{AFX_MSG_MAP(NTreeView)
	ON_WM_CREATE()
	ON_WM_SIZE()
	ON_WM_ERASEBKGND()
	ON_WM_MOUSEWHEEL()
	ON_UPDATE_COMMAND_UI(ID_FILE_REFRESH, OnUpdateFileRefresh)
	ON_COMMAND(ID_FILE_REFRESH, OnFileRefresh)
	//}}AFX_MSG_MAP
	ON_NOTIFY(TVN_SELCHANGED, IDC_TREE, OnSelchanged)
	ON_NOTIFY(NM_RCLICK, IDC_TREE, OnRClick)  // Right Click Menu
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// NTreeView message handlers

int NTreeView::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CWnd::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	if( !m_tree.Create(WS_CHILD|WS_VISIBLE|
			TVS_HASLINES | TVS_LINESATROOT | TVS_HASBUTTONS |
			TVS_SHOWSELALWAYS | TVS_EDITLABELS, // TVS_EDITLABELS | TVS_DISABLEDRAGDROP,
			CRect(0,0,0,0), this, IDC_TREE))
		return -1;
	m_il.Create(16,16, ILC_MASK | ILC_COLOR24, 0, 0);
	CBitmap bmp;
	bmp.LoadBitmap(IDB_IL_TREE);
	m_il.Add(&bmp, RGB(255,255,255));
	m_tree.SetImageList(&m_il, TVSIL_NORMAL);
	
	return 0;
}

void NTreeView::OnSize(UINT nType, int cx, int cy) 
{
	CWnd::OnSize(nType, cx, cy);
	
	m_tree.MoveWindow(0, 0, cx, cy);
}

void NTreeView::PostNcDestroy() 
{
	m_font.DeleteObject();
	m_il.DeleteImageList();
	DestroyWindow();
	delete this;

//	CWnd::PostNcDestroy();
}

#define BUFF_PREVIEW_SIZE	1024

// This function checks whether file archive is valid
// TODO: add MessageBox to report an error
BOOL ImboxviewFile(LPCSTR fileName)
{
	BOOL bRet = FALSE;
	CFile fp;
	int l = strlen(fileName);
	if( l > 4 && _strnicmp(fileName + l - 4, ".eml", 4) == 0 )
		return TRUE;
	char *buff = (char*)calloc(BUFF_PREVIEW_SIZE+1, 1);
	if( buff == NULL )
		return FALSE;
	if( ! fp.Open(fileName, CFile::modeRead) ) {
		free(buff);
		return FALSE;
	}
	if(  fp.Read(buff, BUFF_PREVIEW_SIZE) == 0 ) {
		free(buff);
		fp.Close();
		return FALSE;
	}
	if( strncmp(buff, "From ", 5) == 0 )
		bRet = TRUE;
	else {
		char * p = strstr(buff, "\nFrom ");
		if( p != NULL )
			bRet = TRUE;
	}
	free(buff);
	fp.Close();
	return bRet;
}

BOOL PathFileExist(LPCSTR path);

_int64 FileSize(LPCSTR fileName)
{
	LARGE_INTEGER li = { 0 };
	HANDLE hFile = CreateFile(fileName, GENERIC_READ, FILE_SHARE_READ|FILE_SHARE_WRITE, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE) {
		DWORD err = GetLastError();
		TRACE(_T("(FileSize)INVALID_HANDLE_VALUE error=%ld file=%s\n"), err, fileName);
		return li.QuadPart;
	}
	else {
		BOOL retval = GetFileSizeEx(hFile, &li);
		CloseHandle(hFile);
		if (retval != TRUE) {
			DWORD err = GetLastError();
			TRACE(_T("(GetFileSizeEx)Failed with error=%ld file=%s\n"), err, fileName);
		}
		else {
			long long fsize = li.QuadPart;
			return li.QuadPart;
		}
	}
	return li.QuadPart;
}


// Called on Startup, File open and All File Refresh
// Items are inserted into CTree only by this function
void NTreeView::FillCtrl()
{
	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	m_tree.DeleteAllItems();
	if (path.IsEmpty() || !PathFileExist(path))
		return;
	CString root;
	char *last_slash = (char*)strrchr(path, '\\');
	if (last_slash && *last_slash) 
	{
		// Read first file index file if it exists from previous runs
		// and add to the filesSize hash table
		// new archive files might be discovered and added ??
		CStdioFile fp;
		if (fp.Open(path + "\\.mboxview", CFile::modeRead | CFile::typeText)) {
			CString line;
			while (fp.ReadString(line)) {
				int w = line.Find('\t', 0);
				if (w == -1)
					break;
				CString p = line.Left(w);
				_int64	fSize = 0;
				if (w != -1)
					fSize = _atoi64(line.Mid(w + 1));
				fileSizes[p] = fSize;
			}
			fp.Close();
		}
		else
			fileSizes.RemoveAll();

		root = last_slash + 1;
		HTREEITEM hRoot = m_tree.InsertItem(root, 0, 0, TVI_ROOT);
		CString fw = path + "\\*.*";
		WIN32_FIND_DATA	wf;
		BOOL found;
		// Iterate all archive file mbox or eml in the lastPath folder
		// New archives file can be addedd to CTree but not to fileSizes hash table
		HANDLE f = FindFirstFile(fw, &wf);
		if (f != INVALID_HANDLE_VALUE) {
			do {
				if ((wf.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != FILE_ATTRIBUTE_DIRECTORY && wf.cFileName[0] != '.') {
					CString fn = wf.cFileName;
					CString fullPath = path + "\\" + fn;
					if (ImboxviewFile(fullPath)) {
						HTREEITEM hItem = m_tree.InsertItem(fn, 4, 5, hRoot);
						_int64 fSize = 0;
						_int64 realFSize = FileSize(fullPath);
						found = fileSizes.Lookup(fn, fSize);
						if (fSize != realFSize) 
						{
							//TRACE("File=%s FileSize=%lld StoredFileSize=%lld\n", fn, realFSize, fSize);
							CString cache = fullPath + ".mboxview";
							DeleteFile(cache);
							m_tree.SetItemState(hItem, TVIS_BOLD, TVIS_BOLD);
							CString txt = m_tree.GetItemText(hItem);
							int deb = 1;
						}
					}
				}
			} while (FindNextFile(f, &wf));
			FindClose(f);

#if 0
			CString line;
			CString path;
			_int64 fSize = 0;

			TRACE("FillCtrl: Hash Table fileSizes Count=%d\n", fileSizes.GetCount());
			POSITION pos = fileSizes.GetStartPosition();
			while (pos) {
				fileSizes.GetNextAssoc(pos, path, fSize);
				TRACE("File=%s FileSize=%lld\n", path, fSize);
			}
#endif
		}
		m_tree.Expand(hRoot, TVE_EXPAND);
	}
	
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame) {
		NListView *pListView = pFrame->GetListView();
		if (pListView) {
			pListView->m_path = "";
			pListView->m_which = NULL;
			// pListView->ResetSize();  // check if needed
			pListView->FillCtrl();
			pListView->CloseMailFile();
		}
	}
}

void NTreeView::OnSelchanged(NMHDR* pNMHDR, LRESULT* pResult) 
{
	NMTREEVIEW *pNm = (LPNMTREEVIEW)pNMHDR;
	if( pNm->itemOld.hItem == pNm->itemNew.hItem )
		return;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if( pFrame == NULL )
		return;
	NListView *pListView = pFrame->GetListView();
	if( ! pListView )
		return;

	pListView->CloseMailFile();

	if( ! pNm->itemNew.hItem ) {
		pListView->m_path = "";
		pListView->m_which = NULL;
		pListView->FillCtrl();
		return;
	}
	// This is called when tree is closed
	HTREEITEM hRoot = m_tree.GetRootItem();
	if( pNm->itemNew.hItem == hRoot ) {
		pListView->m_path = "";
		pListView->m_which = NULL;
		pListView->ResetSize();
		pListView->FillCtrl();
		return;
	}
	CString str = m_tree.GetItemText(pNm->itemNew.hItem);
	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	if( str.IsEmpty() || path.IsEmpty() )
		return;
	pListView->m_path = path + "\\" + str;
	pListView->m_which = pNm->itemNew.hItem;
	pListView->ResetSize();
	pListView->FillCtrl();
	MboxMail::nWhichMailList = IDC_ARCHIVE_LIST;
}

void NTreeView::ForceParseMailFile(HTREEITEM hItem)
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame == NULL)
		return;
	NListView *pListView = pFrame->GetListView();
	if (!pListView)
		return;

	pListView->CloseMailFile();

	if (!hItem) {
		pListView->m_path = "";
		pListView->m_which = NULL;
		pListView->FillCtrl();
		return;
	}
	HTREEITEM hRoot = m_tree.GetRootItem();
	if (hItem == hRoot) {
		pListView->m_path = "";
		pListView->m_which = NULL;
		pListView->ResetSize();
		pListView->FillCtrl();
		return;
	}
	CString str = m_tree.GetItemText(hItem);
	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	if (str.IsEmpty() || path.IsEmpty())
		return;
	pListView->m_path = path + "\\" + str;
	CString cache = pListView->m_path + ".mboxview";
	DeleteFile(cache);
	pListView->m_which = hItem;
	pListView->ResetSize();
	pListView->FillCtrl();
	MboxMail::nWhichMailList = IDC_ARCHIVE_LIST;
}


// Called when user specified command line option MAIL_FILE=
void NTreeView::SelectMailFile()
{
	CString str = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("mailFile"));
	if (str.IsEmpty())
		return;

	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("lastPath"));
	if (path.IsEmpty())
		return;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame == NULL)
		return;
	NListView *pListView = pFrame->GetListView();
	if (!pListView)
		return;

	pListView->CloseMailFile();

	pListView->m_path = path + _T("\\") + str;

	CString txt;
	if (!PathFileExist(path))
		txt = _T("Nonexistent Directory \"") + path;
	else if (!PathFileExist(pListView->m_path))
		txt = _T("Nonexistent File \"") + pListView->m_path;

	if (!txt.IsEmpty()) {
		txt += _T("\".\nDo you want to continue?");
		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
			AfxGetMainWnd()->PostMessage(WM_CLOSE);
	}
	else {
		pListView->m_which = NULL;
		pListView->ResetSize();
		pListView->FillCtrl();
	}
}



BOOL NTreeView::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt) 
{
	CWnd *wnd = WindowFromPoint(pt);
	if( wnd == NULL )
		return FALSE;
	if( wnd == this ) {
		BOOL ret =  CWnd::OnMouseWheel(nFlags, zDelta, pt);
		return TRUE;
	}
	if ((GetKeyState(VK_CONTROL) & 0x80) == 0) { // if CTRL key not Down; Do we need to post msg further anyway
		// Commented out, it freezes mbox viewer and and even IE for few seconds when CTRL/SHIFT/etc key are kept down
		; // return wnd->PostMessage(WM_MOUSEWHEEL, MAKELPARAM(nFlags, zDelta), MAKELPARAM(pt.x, pt.y));
	}
	return TRUE;
}

void NTreeView::Traverse( HTREEITEM hItem, CFile &fp )
{
	CString line;
	CString path;
	CString fpath = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	while( hItem != NULL ) {
		path = m_tree.GetItemText(hItem);
		if( ! path.IsEmpty() && ! fpath.IsEmpty() ) {
			CString fullPath = fpath + "\\" + path;
			_int64 fSize = 0;
			//_int64 realFSize = FileSize(fullPath);
			BOOL found = fileSizes.Lookup(path, fSize);
			//TRACE("File=%s FileSize=%lld StoredFileSize=%lld\n", path, realFSize, fSize);
			if (!found) {
				_int64 realFSize = FileSize(fullPath);
				fileSizes[path] = fSize = realFSize;
			}
			line.Format("%s\t%lld\n", path, fSize);
			fp.Write(line, line.GetLength());
		}
		if (m_tree.ItemHasChildren(hItem)){
			HTREEITEM hChild = m_tree.GetChildItem(hItem);
			Traverse(hChild, fp);
		}
		hItem = m_tree.GetNextSiblingItem(hItem);
	}
}


// Thhis function creates or updates per folder .mboxview index file
// It is called when selcting new mail archive file or by RedrawMails called after sorting
// or touching mail list
void NTreeView::SaveData()
{
	HTREEITEM hRoot = m_tree.GetRootItem();
	CFile fp;
	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	if( path.IsEmpty() )
		return;
	if( fp.Open(path+"\\.mboxview", CFile::modeWrite | CFile::modeCreate) ) {
		Traverse(hRoot, fp);
		fp.Close();
	}
#if 0
	TRACE("SaveData: Hash Table fileSizes Count=%d\n", fileSizes.GetCount());
	_int64 fSize = 0;
	POSITION pos = fileSizes.GetStartPosition();
	while (pos) {
		fileSizes.GetNextAssoc(pos, path, fSize);
		TRACE("File=%s FileSize=%lld\n", path, fSize);
	}
#endif
}

void NTreeView::UpdateFileSizesTable(CString &path, _int64 realFSize)
{
	_int64 fSize = 0;
	BOOL found = fileSizes.Lookup(path, fSize);
	if (found)
	{
		if (fSize != realFSize) {
			fileSizes.RemoveKey(path);
			fileSizes[path] = realFSize;
		}
	}
	else
		fileSizes[path] = realFSize;
}

void NTreeView::OnUpdateFileRefresh(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(m_tree.GetRootItem() != NULL);
}

void NTreeView::OnFileRefresh() 
{
	m_tree.DeleteAllItems();
	FillCtrl();
	return;

	// Below was moved to FillCtrl() called above;
	HTREEITEM hSel = m_tree.GetSelectedItem();
	HTREEITEM hRoot = m_tree.GetRootItem();
	if( hRoot == NULL || ! m_tree.ItemHasChildren(hRoot) )
		return;
	HTREEITEM hItem = m_tree.GetChildItem(hRoot);
	CString fpath = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	BOOL bFound = FALSE;
	while( hItem ) {
		CString str = m_tree.GetItemText(hItem);
		if( ! str.IsEmpty() && ! fpath.IsEmpty() ) {
			CString fullPath = fpath + "\\" + str;
			_int64 fSize = 0;
			fileSizes.Lookup(str, fSize);
			if( fSize != FileSize(fullPath) ) {
				CString cache= fullPath + ".mboxview";
				DeleteFile(cache);
				m_tree.SetItemState(hItem, TVIS_BOLD, TVIS_BOLD);
				CString txt = m_tree.GetItemText(hItem);
				bFound = TRUE;
			}
		}
		hItem = m_tree.GetNextSiblingItem(hItem);
	}
	if( bFound )
		m_tree.SelectItem(NULL);
}

void AppendMenu(CMenu *menu, int commandId, const char *commandName)
{
	menu->AppendMenu(MF_STRING, commandId, commandName);
	menu->AppendMenu(MF_SEPARATOR);
}

void NTreeView::OnRClick(NMHDR* pNMHDR, LRESULT* pResult)
{
	// TODO: Add your control notification handler code here
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);

	*pResult = 0;

	HTREEITEM hRoot = m_tree.GetRootItem();

	if (hRoot == 0)
		return;

	if (!(m_tree.GetItemState(hRoot, TVIS_EXPANDED) & TVIS_EXPANDED))
		return;

	CPoint pt;
	//SendMessage(WM_CONTEXTMENU, (WPARAM)m_hWnd, GetMessagePos());
	::GetCursorPos(&pt);
	CWnd *wnd = WindowFromPoint(pt);

	CPoint ptClient(pt);
	ScreenToClient(&ptClient);
	UINT flags;
	HTREEITEM hTreeItem = m_tree.HitTest(ptClient, &flags);

	HTREEITEM hItem = m_tree.GetSelectedItem();
	CString txt = m_tree.GetItemText(hItem);
	CString mailFile = m_tree.GetItemText(hTreeItem);
	CString roottxt = m_tree.GetItemText(hRoot);

	if (!(hItem != 0) && (flags & TVHT_ONITEM))
		return;

	if ((hItem == 0) || (hItem == hRoot) || (hItem != hTreeItem)) {
		CString errorText = "Right Click is supported on selected file only.\nUse left click to select mail file.";
		HWND h = wnd->GetSafeHwnd();
		int answer = ::MessageBox(h, errorText, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;

	}

	CMenu menu;
	menu.CreatePopupMenu();
	menu.AppendMenu(MF_SEPARATOR);

	//
	CMenu printToSubMenu;
	printToSubMenu.CreatePopupMenu();
	printToSubMenu.AppendMenu(MF_SEPARATOR);

	const UINT S_CSV_Id = 1;
	AppendMenu(&printToSubMenu, S_CSV_Id, _T("CSV..."));

	const UINT S_TEXT_Id = 2;
	AppendMenu(&printToSubMenu, S_TEXT_Id, _T("Text.."));

	const UINT S_HTML_Id = 3;
	AppendMenu(&printToSubMenu, S_HTML_Id, _T("HTML.."));

	menu.AppendMenu(MF_POPUP | MF_STRING, (UINT)printToSubMenu.GetSafeHmenu(), _T("Print To"));
	menu.AppendMenu(MF_SEPARATOR);

	//
	CMenu sortSubMenu;
	sortSubMenu.CreatePopupMenu();
	sortSubMenu.AppendMenu(MF_SEPARATOR);

	const UINT S_SORT_BY_DATE_Id = 4;
	AppendMenu(&sortSubMenu, S_SORT_BY_DATE_Id, _T("Date"));

	const UINT S_SORT_BY_FROM_Id = 5;
	AppendMenu(&sortSubMenu, S_SORT_BY_FROM_Id, _T("From"));

	const UINT S_SORT_BY_TO_Id = 6;
	AppendMenu(&sortSubMenu, S_SORT_BY_TO_Id, _T("To"));

	const UINT S_SORT_BY_SUBJ_Id = 7;
	AppendMenu(&sortSubMenu, S_SORT_BY_SUBJ_Id, _T("Subject"));

	const UINT S_SORT_BY_SIZE_Id = 8;
	AppendMenu(&sortSubMenu, S_SORT_BY_SIZE_Id, _T("Size"));

	const UINT S_SORT_BY_CONVERSATION_Id = 9;
	AppendMenu(&sortSubMenu, S_SORT_BY_CONVERSATION_Id, _T("Conversation"));

	const UINT S_SORT_BY_POSITION_Id = 10;
#if 0
	// Sort by position in the archive file. Enabled for debugging only
	AppendMenu(&sortSubMenu, S_SORT_BY_POSITION_Id, _T("Position"));
#endif

	menu.AppendMenu(MF_POPUP | MF_STRING, (UINT)sortSubMenu.GetSafeHmenu(), _T("Sort By"));
	menu.AppendMenu(MF_SEPARATOR);
	//

	const UINT M_FileLocation_Id = 11;
	AppendMenu(&menu, M_FileLocation_Id, _T("Open File Location"));

	const UINT M_Properties_Id = 12;
	AppendMenu(&menu, M_Properties_Id, _T("Properties"));

	const UINT M_Reload_Id = 13;
	AppendMenu(&menu, M_Reload_Id, _T("Refresh Index File"));

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	UINT command = menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, this);
	
	UINT nFlags = TPM_RETURNCMD;
	CString menuString;
	int chrCnt = menu.GetMenuString(command, menuString, nFlags);

	NListView *pListView = 0;
	if (pFrame) {
		pListView = pFrame->GetListView();
	}

	switch (command)
	{
	case M_FileLocation_Id: {
		CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
		if (BrowseToFile(MboxMail::s_path) == FALSE) {  // TODO: s_path error checking ??
			HWND h = GetSafeHwnd();
			HINSTANCE result = ShellExecute(h, _T("open"), path, NULL, NULL, SW_SHOWNORMAL);
			CheckShellExecuteResult(result, h);
		}
	}
	break;
	case M_Properties_Id: {
		CHAR sizeStr_inKB[256]; 
		CHAR sizeStr_inBytes[256];
		int sizeStrSize = 256;
		CString txt;
		CString tmp;

		_int64 fileSize = FileSize(MboxMail::s_path); // TODO: error checking ??
		LPCSTR fileSizeStr_inKB = StrFormatKBSize(fileSize, &sizeStr_inKB[0], sizeStrSize);
		if (!fileSizeStr_inKB)
			sizeStr_inKB[0] = 0;

		LPCSTR fileSizeStr_inBytes = StrFormatByteSize64A(fileSize, &sizeStr_inBytes[0], sizeStrSize);
		if (!fileSizeStr_inBytes)
			sizeStr_inBytes[0] = 0;

		int mailCount = MboxMail::s_mails_ref.GetCount();

		txt.Empty();
		txt.Format("File: %s\n", mailFile);
		tmp.Format("File size:  %s (%s)\n", sizeStr_inKB, sizeStr_inBytes);
		txt += tmp;
		tmp.Empty();
		tmp.Format("Mail Count: %d\n", mailCount);
		txt += tmp;

		HWND h = GetSafeHwnd();
		int answer = ::MessageBox(h, txt, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		int deb = 1;
	}
	break;
	case S_CSV_Id: {
		if (pFrame) {
			pFrame->OnFileExportToCsv();
		}
		int deb = 1;
	}
	break;
	case S_TEXT_Id: {
		if (pFrame) {
			pFrame->OnPrinttoTextFile(0);
		}

		int deb = 1;
	}
	break;
	case S_HTML_Id: {
		if (pFrame) {
			pFrame->OnPrinttoTextFile(1);
		}

		int deb = 1;
	}
	break;
	case S_SORT_BY_DATE_Id: {
		if (pListView)
			pListView->SortByColumn(1);
	}
	break;
	case S_SORT_BY_FROM_Id: {
		if (pListView)
			pListView->SortByColumn(2);
	}
	break;
	case S_SORT_BY_TO_Id: {
		if (pListView)
			pListView->SortByColumn(3);
	}
	break;
	case S_SORT_BY_SUBJ_Id: {
		if (pListView)
			pListView->SortByColumn(4);
	}
	break;
	case S_SORT_BY_SIZE_Id: {
		if (pListView)
			pListView->SortByColumn(5);
	}
	break;
	case S_SORT_BY_CONVERSATION_Id: {
		if (pListView)
			pListView->SortByColumn(0);
	}
	break;
	// Sort by position in the archive file. Enabled for debugging only
	case S_SORT_BY_POSITION_Id: {
		if (pListView)
			pListView->SortByColumn(1, TRUE);
	}
	break;

	case M_Reload_Id: {
		CString txt = _T("Do you want to refresh index file?");
		HWND h = GetSafeHwnd();
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDYES)
			ForceParseMailFile(hItem);
		// For internal testing
		//int ret = MboxMail::DumpMailStatsToFile(&MboxMail::s_mails, MboxMail::s_mails.GetCount());
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
