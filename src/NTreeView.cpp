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

// NTreeView.cpp : implementation file
//

#include "stdafx.h"
#include "FileUtils.h"
#include "mboxview.h"
#include "MboxMail.h"
#include "NTreeView.h"
#include "InputBox.h"
#include "CheckListBoxDlg.h"


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
	m_bSelectMailFileDone = FALSE;
	m_bSelectMailFilePostMsgDone = FALSE;
	m_bGeneralHintPostMsgDone = FALSE;
	m_timerTickCnt = 0;
	m_nIDEvent = 1;
	m_nElapse = 1;
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
	ON_NOTIFY(TVN_SELCHANGING, IDC_TREE, OnSelchanging)
	ON_NOTIFY(NM_RCLICK, IDC_TREE, OnRClick)  // Right Click Menu
	ON_WM_TIMER()
	ON_MESSAGE(WM_CMD_PARAM_FILE_NAME_MESSAGE, OnCmdParam_FileName)
	ON_MESSAGE(WM_CMD_PARAM_GENERAL_HINT_MESSAGE, OnCmdParam_GeneralHint)
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


// Called on Startup, File open and All File Refresh
// Items are inserted into CTree only by this function
void NTreeView::FillCtrl()
{
	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	m_tree.DeleteAllItems();
	if (path.IsEmpty() || !FileUtils::PathDirExists(path))
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
						_int64 realFSize = FileUtils::FileSize(fullPath);
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

	HTREEITEM hItem = pNm->itemNew.hItem;
	HTREEITEM hParent = m_tree.GetParentItem(hItem);

	// FOLDER??
	if (hParent != hRoot)
	{
		CString mboxFileNamePath = MboxMail::s_path;
		CString driveName;
		CString mboxFileDirectory;
		CString mboxFileNameBase;
		CString mboxFileNameExtention;

		FileUtils::SplitFilePath(mboxFileNamePath, driveName, mboxFileDirectory, mboxFileNameBase, mboxFileNameExtention);

		CString mboxFileName = mboxFileNameBase + mboxFileNameExtention;

		CString folderPath; // relative path up to mboxFolderName
		int retpath = GetFolderPath(hItem, mboxFileName, folderPath);
		if (!retpath)
		{
			return;
		}

		CString folderName = m_tree.GetItemText(hItem);

		int retval = pListView->LoadFolderListFile_v2(folderPath, folderName);
#if 0

		CString folderCompletePath = driveName + mboxFileDirectory + "\\" + "Folders" + "\\" + folderPath;

		if (!FileUtils::PathDirExists(folderCompletePath))
			int deb = 1;

		CString folderName = m_tree.GetItemText(hItem);
		CString folderNameCompletePath = folderCompletePath + "\\" + folderName + ".mbox.mboxlist";

		if (!PathFileExists(folderNameCompletePath))
			int deb = 1;

		if (MboxMail::IsFolderMailsSelected())
		{
			pListView->SwitchToMailList(IDC_FOLDER_LIST, TRUE);
		}
		else
		{
			pListView->SwitchToMailList(IDC_FOLDER_LIST, FALSE);
		}
#endif

		return;
	}


	CString str = m_tree.GetItemText(pNm->itemNew.hItem);
	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	if( str.IsEmpty() || path.IsEmpty() )
		return;
	if (path.GetLength() > 0)
	{
		if (path.GetAt(path.GetLength()-1) != _T('\\'))
			pListView->m_path = path + _T('\\') + str;
		else
			pListView->m_path = path + str;
	}
	else
		pListView->m_path = path + _T('\\') + str;

	pListView->m_which = pNm->itemNew.hItem;

	pListView->ResetSize();

	pListView->ResetFilterDates();

	int paneId = 0;
	CString sText;
	sText.Format("Opening %s ...", str);
	pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	pListView->FillCtrl();
	MboxMail::nWhichMailList = IDC_ARCHIVE_LIST;

	pFrame->UpdateFilePrintconfig();

	sText.Format("Ready");
	pFrame->SetStatusBarPaneText(paneId, sText, FALSE);

	MboxMail::ShowHint(HintConfig::MsgWindowPlacementHint, GetSafeHwnd());

	//pListView->ScanAllMailsInMbox();
	//pListView->ScanAllMailsInMbox_NewParser();

	ShowMemStatus();
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
	if (path.GetLength() > 0)
	{
		if (path.GetAt(path.GetLength() - 1) != _T('\\'))
			pListView->m_path = path + _T('\\') + str;
		else
			pListView->m_path = path + str;
	}
	else
		pListView->m_path = path + _T('\\') + str;
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

	if (path.GetLength() > 0)
	{
		if (path.GetAt(path.GetLength() - 1) != _T('\\'))
			pListView->m_path = path + _T('\\') + str;
		else
			pListView->m_path = path + str;
	}
	else
		pListView->m_path = path + _T('\\') + str;

	CString txt;
	if (!FileUtils::PathDirExists(path))
		txt = _T("Nonexistent Directory \"") + path;
	else if (!FileUtils::PathFileExist(pListView->m_path))
		txt = _T("Nonexistent File \"") + pListView->m_path;

	if (!txt.IsEmpty()) {
		txt += _T("\".\nDo you want to continue?");
		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
			AfxGetMainWnd()->PostMessage(WM_CLOSE);
	}
	else {
		HTREEITEM hItem = NTreeView::FindItem(0, str);
		UINT nCode = TVGN_CARET;
		BOOL retval = m_tree.Select(hItem, nCode);
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
			//_int64 realFSize = FileUtils::FileSize(fullPath);
			BOOL found = fileSizes.Lookup(path, fSize);
			//TRACE("File=%s FileSize=%lld StoredFileSize=%lld\n", path, realFSize, fSize);
			if (!found) {
				_int64 realFSize = FileUtils::FileSize(fullPath);
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
	// Test
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
			if( fSize != FileUtils::FileSize(fullPath) ) {
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

	const UINT S_PDF_Id = 33;
	AppendMenu(&printToSubMenu, S_PDF_Id, _T("PDF.."));

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
	// Sort by position in the archive file. Enabled for debugging only
	// AppendMenu(&sortSubMenu, S_SORT_BY_POSITION_Id, _T("Position"));

	menu.AppendMenu(MF_POPUP | MF_STRING, (UINT)sortSubMenu.GetSafeHmenu(), _T("Sort By"));
	menu.AppendMenu(MF_SEPARATOR);
	//

	const UINT M_FileLocation_Id = 11;
	AppendMenu(&menu, M_FileLocation_Id, _T("Open File Location"));

	const UINT M_Properties_Id = 12;
	AppendMenu(&menu, M_Properties_Id, _T("Properties"));

	const UINT M_Reload_Id = 13;
	AppendMenu(&menu, M_Reload_Id, _T("Refresh Index File"));

	const UINT M_CreateFolder_Id = 14;
	//AppendMenu(&menu, M_CreateFolder_Id, _T("Create Folder"));  // TODO: later

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
		UINT nFlags = MF_BYCOMMAND;
		CString Label;
		int retLabel = menu.GetMenuString(M_FileLocation_Id, Label, nFlags);

		CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
		if (FileUtils::BrowseToFile(MboxMail::s_path) == FALSE) {  // TODO: s_path error checking ??
			HWND h = GetSafeHwnd();
			HINSTANCE result = ShellExecute(h, _T("open"), path, NULL, NULL, SW_SHOWNORMAL);
			CMainFrame::CheckShellExecuteResult(result, h);
		}
	}
	break;
	case M_Properties_Id: {
		CHAR sizeStr_inKB[256]; 
		CHAR sizeStr_inBytes[256];
		int sizeStrSize = 256;
		CString txt;
		CString tmp;

		_int64 fileSize = FileUtils::FileSize(MboxMail::s_path); // TODO: error checking ??
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
		UINT nFlags = MF_BYCOMMAND;
		CString Label;
		int retLabel = menu.GetMenuString(S_TEXT_Id, Label, nFlags);
		if (pFrame) 
		{
			pFrame->PrintMailArchiveToTEXT();
		}

		int deb = 1;
	}
	break;
	case S_HTML_Id: {
		if (pFrame) 
		{
			pFrame->PrintMailArchiveToHTML();
		}

		int deb = 1;
	}
	break;
	case S_PDF_Id: {
		if (pFrame) 
		{
			MboxMail::ShowHint(HintConfig::PrintToPDFHint, GetSafeHwnd());
			pFrame->PrintMailArchiveToPDF();
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

	case M_CreateFolder_Id: 
	{
		CreateEmptyFolder(hItem);
		CMenu menu;
		CString mboxFilePath = MboxMail::s_path;
		CString mboxFileName = ::PathFindFileName(mboxFilePath);
		HTREEITEM hItem = NTreeView::FindItem(0, mboxFileName);
		TRACE(_T("CreateFlatFolderListMenu\n"));
		CArray<CString> folderList;
		CreateFlatFolderList(hItem, folderList);
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

void NTreeView::OnSelchanging(NMHDR* pNMHDR, LRESULT* pResult)
{
	NMTREEVIEW *pNm = (LPNMTREEVIEW)pNMHDR;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame == NULL)
		return;
	NListView *pListView = pFrame->GetListView();
	if (!pListView)
		return;
	if (!pListView->IsUserSelectedMailListEmpty())
	{
		CString txt = _T("User Select Mails list is not empty.\n"
			"Content will be lost if you switch to new mail archive.\n"
			"Do you want to continue?");
		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
		{
			*pResult = TRUE;
		}
	}
}

HTREEITEM NTreeView::FindItem(HTREEITEM hItem, CString &mailFileName)
{
	CString path;
	if (hItem == 0)
		hItem = m_tree.GetRootItem();

	while (hItem != NULL)
	{
		path = m_tree.GetItemText(hItem);
		if (path.Compare(mailFileName) == 0)
			return hItem;
		if (m_tree.ItemHasChildren(hItem))
		{
			HTREEITEM hChild = m_tree.GetChildItem(hItem);
			HTREEITEM hFound = FindItem(hChild, mailFileName);
			if (hFound)
				return hFound;
		}
		hItem = m_tree.GetNextSiblingItem(hItem);
	}
	return 0;
}

void NTreeView::StartTimer()
{
	SetTimer(m_nIDEvent, m_nElapse, NULL);
}


void NTreeView::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: Add your message handler code here and/or call default

	if (m_timerTickCnt == 0)
		SelectMailFile();  // will trigger NTreeView::OnSelchanged
	m_timerTickCnt++;

	// Need just one shot timer
	// Is killing timer from the timer callback ok ? Seem to work -:)
	// May need to move to FillCtrl() to be safe
	KillTimer(nIDEvent);

	// But dont send msg after KillTimer
	//CWnd::OnTimer(nIDEvent);
}


void NTreeView::FindAllDirs(LPCTSTR pstr)
{
	CFileFind finder;

	// build a string with wildcards
	CString strWildcard(pstr);
	strWildcard += _T("\\*.*");

	// start working for files
	BOOL bWorking = finder.FindFile(strWildcard);

	while (bWorking)
	{
		bWorking = finder.FindNextFile();

		// skip . and .. files; otherwise, we'd
		// recur infinitely!

		if (finder.IsDots())
			continue;

		// if it's a directory, recursively search it

		if (finder.IsDirectory())
		{
			CString str = finder.GetFilePath();
			TRACE("%s\n", (LPCTSTR)str);
			NTreeView::FindAllDirs(str);
		}
	}

	finder.Close();
}

void NTreeView::PostMsgCmdParamFileName()
{
	if (GetSafeHwnd() && !m_bGeneralHintPostMsgDone)
	{
		BOOL res = PostMessage(WM_CMD_PARAM_GENERAL_HINT_MESSAGE, 0, 0);
		m_bGeneralHintPostMsgDone = TRUE;
	}

	if (GetSafeHwnd() && !m_bSelectMailFilePostMsgDone)
	{
		BOOL res = PostMessage(WM_CMD_PARAM_FILE_NAME_MESSAGE, 0, 0);
		m_bSelectMailFilePostMsgDone = TRUE;
	}
}

LRESULT NTreeView::OnCmdParam_FileName(WPARAM wParam, LPARAM lParam)
{
	SelectMailFile();
	return 0;
}

LRESULT NTreeView::OnCmdParam_GeneralHint(WPARAM wParam, LPARAM lParam)
{
	MboxMail::ShowHint(HintConfig::GeneralUsageHint, GetSafeHwnd());
	return 0;
}

// Mail Folder Helpers

int NTreeView::GetFolderPath(HTREEITEM hItem, CString &mboxName, CString &parentPath)
{
	CString parentName = m_tree.GetItemText(hItem);
	//if (parentName.Compare(mboxName) == 0) return 1;
	parentPath.Append(parentName);
	HTREEITEM hParent = m_tree.GetParentItem(hItem);
	HTREEITEM hNextParent;
	while (hParent)
	{
		hNextParent = m_tree.GetParentItem(hParent);
		if (hNextParent == 0) // root
			break;
		parentName = m_tree.GetItemText(hParent);
		//if (parentName.Compare(mboxName) == 0) break;
		parentPath.Insert(0,"\\");
		parentPath.Insert(0, parentName);
		hParent = hNextParent;
	}
	return 1;
}

int NTreeView::CreateEmptyFolder(HTREEITEM hItem)
{
	CString mboxFileNamePath = MboxMail::s_path;
	CString driveName;
	CString mboxFileDirectory;
	CString mboxFileNameBase;
	CString mboxFileNameExtention;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	NListView *pListView = 0;
	if (pFrame == 0) {
		return 0;
	}

	pListView = pFrame->GetListView();

	FileUtils::SplitFilePath(mboxFileNamePath, driveName, mboxFileDirectory, mboxFileNameBase, mboxFileNameExtention);

	CString mboxFolderName = mboxFileNameBase + mboxFileNameExtention;
	InputBox dlg(mboxFolderName);

	INT_PTR answer = dlg.DoModal();
	if (answer != IDOK)
	{
		return 0;
	}

	CString newFolderName = dlg.m_input;
	CString subFolderPath;  // path up to mboxFolderName
	int retpath = GetFolderPath(hItem, newFolderName, subFolderPath);
	if (!retpath)
	{
		return 1;
	}

	int retfolder = pListView->CreateEmptyFolder(driveName, mboxFileDirectory, mboxFolderName, subFolderPath, newFolderName);
	if (retfolder != ERROR_SUCCESS)
	{
		return 0;
	}
		

	HTREEITEM newItem = m_tree.InsertItem(newFolderName, hItem, TVI_SORT);
	if (newItem == 0) 
	{
		return 0;
	}

	m_tree.Expand(hItem, TVE_EXPAND);

	CString fileDirPath;
	BOOL ret = FileUtils::CPathGetPath(MboxMail::s_path, fileDirPath);
	fileDirPath += "\\Folders";
	FindAllDirs(fileDirPath);

	return 1;
}

int NTreeView::GetFolderPath(HTREEITEM hItem, CString &folderPath)
{
	HTREEITEM hNextParent = 0;
	HTREEITEM hNextNextParent = 0;
	HTREEITEM hParent = 0;

	while (hItem)
	{
		hParent = m_tree.GetParentItem(hItem);
		CString parentName;
		if (hParent) {
			parentName = m_tree.GetItemText(hParent);
			hNextParent = m_tree.GetParentItem(hParent);
		}

		CString nextParentName;
		if (hNextParent) {
			nextParentName = m_tree.GetItemText(hNextParent);
			hNextNextParent = m_tree.GetParentItem(hNextParent);
		}

		CString nextNextParentName;
		if (hNextNextParent) {
			nextNextParentName = m_tree.GetItemText(hNextNextParent);
		}

		CString itemName = m_tree.GetItemText(hItem);

		if (!folderPath.IsEmpty())
			folderPath.Insert(0, "\\");
		folderPath.Insert(0, itemName);

		if (hNextNextParent == 0)
			break;

		hItem = hParent;
	}
	return 1;
}


int NTreeView::CreateFlatFolderList(HTREEITEM hItem, CArray<CString> &folderList)
{
	CString path;
	CString itemName;
	HTREEITEM hLeaveItem = 0;
	//TRACE(_T("ItemTree!!\n"));
	if (hItem == 0)
	{
		hItem = m_tree.GetRootItem();
	}

	CString rootName = m_tree.GetItemText(hItem);
	if (!m_tree.ItemHasChildren(hItem))
		return 1;

	hItem = m_tree.GetChildItem(hItem);
	CString firstChildName = m_tree.GetItemText(hItem);

	while (hItem != NULL)
	{
		CString folderPath;
		int ret = GetFolderPath(hItem, folderPath);
		folderList.Add(folderPath);
		TRACE(_T("TreeFolderPath=%s\n"), folderPath);

		if (m_tree.ItemHasChildren(hItem))
		{
			HTREEITEM hChild = m_tree.GetChildItem(hItem);
			CString childName = m_tree.GetItemText(hChild);
			CString parentName = m_tree.GetItemText(hItem);
			CreateFlatFolderList(hItem, folderList);
		}
		hItem = m_tree.GetNextSiblingItem(hItem);
		CString nextChildName = m_tree.GetItemText(hItem);
	}

	return 1;
}

int NTreeView::CreateFlatFolderList(CString &mboxFileName, CArray<CString> &folderList)
{
	int retval = -1;
	HTREEITEM hRoot = 0;
	HTREEITEM hItem = NTreeView::FindItem(hRoot, mboxFileName);
	if (hItem != 0)
	{
		retval =  CreateFlatFolderList(hItem, folderList);
	}
	return retval;
}

