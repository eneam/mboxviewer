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
	ON_UPDATE_COMMAND_UI(ID_TREE_EXPAND, OnUpdateTreeExpand)
	ON_COMMAND(ID_TREE_EXPAND, OnTreeExpand)
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

	m_folderArray.LoadFromRegistry();

	//m_folderArray.Dump();
	
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

#define BUFF_PREVIEW_SIZE	10240

// This function checks whether file archive is valid
// TODO: add MessageBox to report an error
BOOL ImboxviewFile(CString & fName)
{
	LPCSTR fileName = fName;

	// need better validation check; may need to parse the file
	BOOL bRet = FALSE;
	CFile fp;

	CString fileNameExtention;
	FileUtils::GetFileExtension(fName, fileNameExtention);
	if (fileNameExtention.CompareNoCase(".eml") == 0)
		return TRUE;

	CString fileNameExtentionRoot = fileNameExtention.Mid(0, 5);

	if (fileNameExtentionRoot.CompareNoCase(".mbox") != 0)
		return FALSE;

	char *buff = (char*)malloc(BUFF_PREVIEW_SIZE + 1);
	if( buff == NULL )
		return FALSE;

	if( ! fp.Open(fileName, CFile::modeRead) ) {
		free(buff);
		return FALSE;
	}
	// Could do better -:)
	for (int i = 0; i < 5000; i++)
	{
		UINT nCount = fp.Read(buff, BUFF_PREVIEW_SIZE);
		if (nCount == 0) {
			free(buff);
			fp.Close();
			return FALSE;
		}
		buff[nCount] = 0;
		if ((i == 0) && strncmp(buff, "From ", 5) == 0)
		{
			bRet = TRUE;
			break;
		}
		else 
		{
			char * p = strstr(buff, "\nFrom ");
			if (p != NULL)
			{
				bRet = TRUE;
				break;
			}
		}
	}
	free(buff);
	fp.Close();
	return bRet;
}

HTREEITEM NTreeView::HasFolder(CString &path)
{
	CString folderPath;
	CString normalizedPath = path;
	normalizedPath.TrimRight("\\");

	HTREEITEM hRoot = m_tree.GetRootItem();
	if (hRoot)
	{
		HTREEITEM hNext = hRoot;
		while (hNext)
		{
			int retIndex = m_tree.GetItemData(hNext);
			m_folderArray.GetAt(retIndex, folderPath);
			folderPath.TrimRight("\\");
			if (normalizedPath.CompareNoCase(folderPath) == 0)
				return hNext;
				
			hNext = m_tree.GetNextItem(hNext, TVGN_NEXT);
		}
		return 0;
	}
	else
		return 0;
}

void NTreeView::LoadFolders()
{
	CString pathLast = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");

	int paneId = 0;
	CString sText;
	CString str;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	BOOL lastPathLoaded = FALSE;
	BOOL expand = FALSE;
	int ii = 0;
	for (ii = 0; ii < m_folderArray.m_array.GetCount(); ii++)
	{
		CString path = m_folderArray.m_array.GetAt(ii);
		if (!path.IsEmpty())
		{
			CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath", path);
			if (path.CompareNoCase(pathLast) == 0)
				continue;

			FileUtils::CPathStripPath(path, str);
			sText.Format("Opening %s ...", str);

			if (pFrame)
				pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

			FillCtrl(expand);

			sText.Format("Ready");
			if (pFrame)
				pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
			int deb = 1;
		}
	}

	CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath", pathLast);

	FileUtils::CPathStripPath(pathLast, str);
	sText.Format("Opening %s ...", str);

	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	expand = TRUE;
	FillCtrl(expand);

	sText.Format("Ready");
	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, FALSE);

	m_tree.SortChildren(0);
	// or m_tree.SortChildren(TVI_ROOT);
}

void  NTreeView::ExpandOrCollapseTree(BOOL expand)
{
	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");

	HTREEITEM hFolder = 0;
	if (!path.IsEmpty() && FileUtils::PathDirExists(path))
	{
		hFolder = HasFolder(path);
	}

	HTREEITEM hCurrentSelectedItem = m_tree.GetSelectedItem();
	if (hCurrentSelectedItem)
	{
		CString txt = m_tree.GetItemText(hCurrentSelectedItem);
		m_tree.SetItemState(hCurrentSelectedItem, 0, TVIS_BOLD);

		CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
		if (pFrame) 
		{
			NListView *pListView = pFrame->GetListView();
			if (pListView) 
			{
				pListView->m_path = "";
				pListView->m_which = NULL;
				pListView->ResetSize();
				pListView->FillCtrl();
			}
		}
		int deb = 1;
	}

	UINT nCode = TVE_EXPAND;
	if (!expand)
		nCode = TVE_COLLAPSE;

	HTREEITEM hRoot = m_tree.GetRootItem();
	HTREEITEM hNext = hRoot;
	while (hNext)
	{
		if (hFolder != hNext)
		{
			m_tree.Expand(hNext, nCode);
		}
		hNext = m_tree.GetNextItem(hNext, TVGN_NEXT);
	}
	if (hFolder)
	{
		m_tree.Expand(hFolder, nCode);
		m_tree.EnsureVisible(hFolder);
	}
}

// Called on Startup, File open and All File Refresh
// Items are inserted into CTree only by this function
void NTreeView::FillCtrl(BOOL expand)
{
	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	//m_tree.DeleteAllItems();
	if (path.IsEmpty() || !FileUtils::PathDirExists(path))
		return;

	// traverse the Tree to see if the folder already exist
	HTREEITEM hFolder;
	if (hFolder = HasFolder(path))
	{
		HTREEITEM hCurrentSelectedItem = m_tree.GetSelectedItem();
		if (hCurrentSelectedItem != hFolder)
		{
			if (hCurrentSelectedItem >= 0)
				m_tree.SetItemState(hCurrentSelectedItem, 0, TVIS_BOLD);

			UINT nCode = TVGN_CARET;
			if (hFolder)
			{
				BOOL retval = m_tree.Select(hFolder, nCode);
			}
		}

		m_tree.Expand(hFolder, TVE_EXPAND);
		return;
	}

	CString root;
	path.TrimRight("\\");
	char *last_slash = (char*)strrchr(path, '\\');
	if (last_slash && *last_slash) 
	{
		// Read first file index file if it exists from previous runs
		// and add to the filesSize hash table
		// new archive files might be discovered and added ??
		CStdioFile fp;
		if (fp.Open(path + "\\.mboxview", CFile::modeRead | CFile::typeText)) 
		{
			CString line;
			while (fp.ReadString(line)) 
			{
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
		if (hRoot == 0)
			return;

		CString registry_lastPath = path;
		registry_lastPath.TrimRight("\\");
		registry_lastPath.Append("\\");

		int index = m_folderArray.Add(registry_lastPath);

		HTREEITEM hCurrentSelectedItem = m_tree.GetSelectedItem();
		if (hCurrentSelectedItem >= 0)
			m_tree.SetItemState(hCurrentSelectedItem, 0, TVIS_BOLD);

		UINT nCode = TVGN_CARET;
		BOOL retval = m_tree.Select(hRoot, nCode);
		m_tree.SetItemData(hRoot, index);
		DWORD_PTR retIndex = m_tree.GetItemData(hRoot);

		CString fw = path + "\\*.*";
		WIN32_FIND_DATA	wf;
		BOOL found;
		// Iterate all archive file mbox or eml in the lastPath folder
		// New archives file can be addedd to CTree but not to fileSizes hash table
		HANDLE f = FindFirstFile(fw, &wf);
		if (f != INVALID_HANDLE_VALUE) 
		{
			do 
			{
				if ((wf.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != FILE_ATTRIBUTE_DIRECTORY && wf.cFileName[0] != '.') 
				{
					CString fn = wf.cFileName;
					CString fullPath = path + "\\" + fn;
					if (ImboxviewFile(fullPath)) 
					{
						HTREEITEM hItem = m_tree.InsertItem(fn, 4, 5, hRoot);
						if (hItem)
						{
							_int64 fSize = 0;
							_int64 realFSize = FileUtils::FileSize(fullPath);
							found = fileSizes.Lookup(fn, fSize);
							if (fSize != realFSize)
							{
								//TRACE("File=%s FileSize=%lld StoredFileSize=%lld\n", fn, realFSize, fSize);
								CString cache = fullPath + ".mboxview";
								DeleteFile(cache);
								//m_tree.SetItemState(hItem, TVIS_BOLD, TVIS_BOLD);
								CString txt = m_tree.GetItemText(hItem);
								int deb = 1;
							}
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
		if (expand)
			m_tree.Expand(hRoot, TVE_EXPAND);
	}
	
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame) 
	{
		NListView *pListView = pFrame->GetListView();
		if (pListView) 
		{
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

	HTREEITEM hNewItem = pNm->itemNew.hItem;
	HTREEITEM hOldItem = pNm->itemOld.hItem;

	if (hOldItem == hNewItem)
		return;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if ( pFrame == NULL )
		return;
	NListView *pListView = pFrame->GetListView();
	if ( ! pListView )
		return;

	pListView->CloseMailFile();
	
	if ( !hNewItem) 
	{
		pListView->m_path = "";
		pListView->m_which = NULL;
		pListView->FillCtrl();
		return;
	}

	// This is called when tree is closed
	HTREEITEM hRoot = DetermineRootItem(hNewItem);
	CString rootName = m_tree.GetItemText(hRoot);

	HTREEITEM hParent = m_tree.GetParentItem(hNewItem);
	if (hParent == 0)
	{
		if (hOldItem >= 0)
			m_tree.SetItemState(hOldItem, 0, TVIS_BOLD);
		m_tree.SetItemState(hNewItem, TVIS_BOLD, TVIS_BOLD);

		CString folderPath;
		DetermineFolderPath(hNewItem, folderPath);

		BOOL retval = CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath", folderPath);

		pListView->m_path = "";
		pListView->m_which = NULL;
		pListView->ResetSize();
		pListView->FillCtrl();
		return;
	}

	// FOLDER??
	//if (hParent != hRoot)
	if (0)
	{
		CString mboxFileNamePath = MboxMail::s_path;
		CString driveName;
		CString mboxFileDirectory;
		CString mboxFileNameBase;
		CString mboxFileNameExtention;

		FileUtils::SplitFilePath(mboxFileNamePath, driveName, mboxFileDirectory, mboxFileNameBase, mboxFileNameExtention);

		CString mboxFileName = mboxFileNameBase + mboxFileNameExtention;

		CString folderPath; // relative path up to mboxFolderName
		int retpath = GetFolderPath(hNewItem, mboxFileName, folderPath);
		if (!retpath)
		{
			return;
		}

		CString folderName = m_tree.GetItemText(hNewItem);

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

	CString str = m_tree.GetItemText(hNewItem);

	CString folderPath;
	DetermineFolderPath(hNewItem, folderPath);

	BOOL retval = CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath", folderPath);
	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");

	if( str.IsEmpty() || path.IsEmpty() )
		return;

	path.TrimRight("\\");
	pListView->m_path = path + _T('\\') + str;

	if (hOldItem >= 0)
	{
		CString text = m_tree.GetItemText(hOldItem);
		m_tree.SetItemState(hOldItem, 0, TVIS_BOLD);
	}
	m_tree.SetItemState(hNewItem, TVIS_BOLD, TVIS_BOLD);

	pListView->m_which = hNewItem;

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
	HTREEITEM hRoot = DetermineRootItem(hItem);
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
	path.TrimRight("\\");
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
	// mailFile can be mailFilePath or mailFileName
	CString mailFile = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("mailFile"));
	if (mailFile.IsEmpty())
		return;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame == NULL)
		return;
	NListView *pListView = pFrame->GetListView();
	if (!pListView)
		return;

	pListView->CloseMailFile();

	CString fileName;
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;
	FileUtils::SplitFilePath(mailFile, driveName, directory, fileNameBase, fileNameExtention);

	CString mailFileName = fileNameBase + fileNameExtention;

	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("lastPath"));
	if (directory.GetLength() > 0)
	{
		path = driveName + directory;
	}

	if (path.IsEmpty())
		return;

	path.TrimRight("\\");
	pListView->m_path = path + _T('\\') + mailFileName;

	CString txt;
	if (!FileUtils::PathDirExists(path))
		txt = _T("Nonexistent Directory \"") + path;
	else if (!FileUtils::PathFileExist(pListView->m_path))
		txt = _T("Nonexistent File \"") + pListView->m_path;

	if (!txt.IsEmpty()) 
	{
		txt += _T("\".\nDo you want to continue?");
		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
			AfxGetMainWnd()->PostMessage(WM_CLOSE);
	}
	else 
	{
		HTREEITEM hFolder = HasFolder(path);

		HTREEITEM hItem = NTreeView::FindItem(hFolder, mailFileName);
		UINT nCode = TVGN_CARET;
		if (hItem)
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
	CString ipath;
	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	path.TrimRight("\\");
	while( hItem != NULL ) 
	{
		ipath = m_tree.GetItemText(hItem);
		if( ! ipath.IsEmpty() && ! path.IsEmpty() ) 
		{
			CString fullPath = path + "\\" + ipath;
			_int64 fSize = 0;
			//_int64 realFSize = FileUtils::FileSize(fullPath);
			BOOL found = fileSizes.Lookup(ipath, fSize);
			//TRACE("File=%s FileSize=%lld StoredFileSize=%lld\n", path, realFSize, fSize);
			if (!found) 
			{
				_int64 realFSize = FileUtils::FileSize(fullPath);
				fileSizes[ipath] = fSize = realFSize;
			}
			line.Format("%s\t%lld\n", ipath, fSize);
			fp.Write(line, line.GetLength());
		}
		if (m_tree.ItemHasChildren(hItem))
		{
			HTREEITEM hChild = m_tree.GetChildItem(hItem);
	
			Traverse(hChild, fp);
		}
		hItem = m_tree.GetNextSiblingItem(hItem);
		HTREEITEM hParent = m_tree.GetParentItem(hItem);
		if (hParent == 0)
			return;
	}
}


// Thhis function creates or updates per folder .mboxview index file
// It is called when selcting new mail archive file or by RedrawMails called after sorting
// or touching mail list
void NTreeView::SaveData(HTREEITEM m_which)
{
	CFile fp;
	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	if( path.IsEmpty() )
		return;
	path.TrimRight("\\");

	HTREEITEM hRoot = HasFolder(path);
	if (hRoot == 0)
		return;

	CString rootName = m_tree.GetItemText(hRoot);

	if( fp.Open(path+"\\.mboxview", CFile::modeWrite | CFile::modeCreate) ) 
	{
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

void NTreeView::OnTreeExpand()
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (pFrame)
		pFrame->UpdateToolsBar();
}

void NTreeView::OnUpdateTreeExpand(CCmdUI* pCmdUI)
{
	pCmdUI->Enable(m_tree.GetRootItem() != NULL);
}

void NTreeView::OnUpdateFileRefresh(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(m_tree.GetRootItem() != NULL);
}

void NTreeView::OnFileRefresh() 
{
	m_tree.DeleteAllItems();
	LoadFolders();
	return;
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

	CPoint pt;
	::GetCursorPos(&pt);
	CWnd *wnd = WindowFromPoint(pt);

	CPoint ptClient(pt);
	ScreenToClient(&ptClient);

	HTREEITEM hItem = m_tree.GetSelectedItem();
	CString itemTxt;
	if (hItem)
		itemTxt = m_tree.GetItemText(hItem);

	UINT flags = TVHT_TORIGHT|TVHT_ABOVE;  // 1280 = 0x400 + 0x100  - flags are populated by HitTest
	HTREEITEM hTreeItem = m_tree.HitTest(ptClient, &flags);

	if (hTreeItem == 0)
		return;

	CString mailFile = m_tree.GetItemText(hTreeItem);

	if (!(flags & TVHT_ONITEM))
		return;

	HTREEITEM hRoot = DetermineRootItem(hTreeItem);

	if (hRoot == 0)
		return;

	CString roottxt = m_tree.GetItemText(hRoot);

	//if (!(m_tree.GetItemState(hRoot, TVIS_EXPANDED) & TVIS_EXPANDED))
		//return;

	HTREEITEM hParent = m_tree.GetParentItem(hTreeItem);

	if ((hParent == 0) && (hItem != hTreeItem))
	{
		CString errorText = "Right Click is supported on selected items only.\nUse left click to select mail folder first.";
		HWND h = wnd->GetSafeHwnd();
		int answer = ::MessageBox(h, errorText, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;
	}

	if ((hParent != 0) && (hItem != hTreeItem))
	{
		CString errorText = "Right Click is supported on selected items only.\nUse left click to select mail file first.";
		HWND h = wnd->GetSafeHwnd();
		int answer = ::MessageBox(h, errorText, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;
	}

	if ((hParent == 0) && (hItem == hTreeItem))
	{
		CMenu menu;
		menu.CreatePopupMenu();
		//menu.AppendMenu(MF_SEPARATOR);

		//
		const UINT M_DeleteItem_Id = 1;
		AppendMenu(&menu, M_DeleteItem_Id, _T("Delete Folder"));
		const UINT M_FolderPath_Id = 2;
		AppendMenu(&menu, M_FolderPath_Id, _T("Show Folder Path"));

		CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

		UINT command = menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, this);

		UINT nFlags = TPM_RETURNCMD;
		CString menuString;
		int chrCnt = menu.GetMenuString(command, menuString, nFlags);

		switch (command)
		{
		case M_DeleteItem_Id: {
			UINT nFlags = MF_BYCOMMAND;
			CString Label;
			int retLabel = menu.GetMenuString(M_DeleteItem_Id, Label, nFlags);

			CString infoText = 
				"The selected folder will be deleted from the current view.\n"
				"The physical folder and content on the disk will not be deleted.\n\n"
				"Delete the selected folder ?";
			HWND h = wnd->GetSafeHwnd();
			int answer = ::MessageBox(h, infoText, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
			if (answer == IDYES)
			{
				if (hItem)
					BOOL ret = DeleteFolder(hItem);
			}
		}
		break;
		case M_FolderPath_Id: {
			CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
			path.TrimRight("\\");
			HWND h = wnd->GetSafeHwnd();
			int answer = ::MessageBox(h, path, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_USERICON);
		}
		 break;

		default: {
				int deb = 1;
		}
		break;
		}

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

BOOL NTreeView::DeleteFolder(HTREEITEM hItem)
{
	if (hItem == 0)
		return FALSE;

	CString itemTxt;
	UINT nCode = TVE_COLLAPSE;
	m_tree.Expand(hItem, nCode);

	DeleteItemChildren(hItem);

	CString folderPath;
	DetermineFolderPath(hItem, folderPath);
	BOOL ret = m_folderArray.Delete(folderPath);

	itemTxt = m_tree.GetItemText(hItem);
	TRACE("Deleting %s\n", itemTxt);
	m_tree.DeleteItem(hItem);

	HTREEITEM hRootItem = m_tree.GetRootItem();
	if (hRootItem)
	{
		CString folderPath;
		DetermineFolderPath(hRootItem, folderPath);
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath", folderPath);
		FillCtrl(FALSE);
	}
	else
	{
		CString empty = "";
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath", empty);
	}
	return ret;
}

BOOL NTreeView::DeleteItemChildren(HTREEITEM hItem)
{
	if (hItem == 0)
		return FALSE;

	CString itemTxt;
	itemTxt = m_tree.GetItemText(hItem);

	HTREEITEM hNextNextItem;
	HTREEITEM hNextItem = m_tree.GetChildItem(hItem);

	while (hNextItem != NULL)
	{
		hNextNextItem = m_tree.GetNextSiblingItem(hNextItem);
		if (m_tree.ItemHasChildren(hNextItem))
		{
			HTREEITEM hChild = m_tree.GetChildItem(hNextItem);
			if (hChild)
				NTreeView::DeleteItem(hChild);
		}
		else
		{
			itemTxt = m_tree.GetItemText(hNextItem);
			TRACE("\tDeleting %s\n", itemTxt);
			m_tree.DeleteItem(hNextItem);
		}
		hNextItem = hNextNextItem;
	}
	return TRUE;
}

BOOL NTreeView::DeleteItem(HTREEITEM hItem)
{
	if (hItem == 0)
		return FALSE;

	CString itemTxt;
	UINT nCode = TVE_COLLAPSE;
	m_tree.Expand(hItem, nCode);

	HTREEITEM hNextNextItem;
	HTREEITEM hNextItem = m_tree.GetChildItem(hItem);

	while (hNextItem != NULL)
	{
		hNextNextItem = m_tree.GetNextSiblingItem(hNextItem);
		if (m_tree.ItemHasChildren(hNextItem))
		{
			HTREEITEM hChild = m_tree.GetChildItem(hNextItem);
			if (hChild)
				NTreeView::DeleteItem(hChild);
		}
		else
		{
			itemTxt = m_tree.GetItemText(hNextItem);
			TRACE("\tDeleting %s\n", itemTxt);
			m_tree.DeleteItem(hNextItem);
		}
		hNextItem = hNextNextItem;
	}
	itemTxt = m_tree.GetItemText(hItem);
	TRACE("Deleting %s\n", itemTxt);
	m_tree.DeleteItem(hItem);
	return TRUE;
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

BOOL NTreeView::OnEraseBkgnd(CDC* pDC)
{
	//BOOL ret = CWnd::OnEraseBkgnd(pDC);

	DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailArchiveList);

	m_tree.SetBkColor(color);

	CRect rect;
	GetClientRect(&rect);
	pDC->FillRect(&rect, &CBrush(color));

	return TRUE;
}

//  class CRegArray

void NTreeView::DetermineFolderPath(HTREEITEM hItem, CString &folderPath)
{
	folderPath.Empty();
	HTREEITEM hParent = hItem;
	while (hParent)
	{
		hItem = hParent;
		hParent = m_tree.GetParentItem(hParent);
	}
	if (hItem)
	{
		int retIndex = m_tree.GetItemData(hItem);
		m_folderArray.GetAt(retIndex, folderPath);
		//folderPath.TrimRight("\\");
	}
}

HTREEITEM NTreeView::DetermineRootItem(HTREEITEM hItem)
{
	HTREEITEM hParent = hItem;
	while (hParent)
	{
		hItem = hParent;
		hParent = m_tree.GetParentItem(hParent);
	}
	return hItem;
}

CRegArray::CRegArray()
{
	m_section = CString(sz_Software_mboxview) + "\\" + "folders";
	m_nMaxSize = 128;
}

CRegArray::CRegArray(CString &section)
{
	m_section = section;
	m_nMaxSize = 128;
}

CRegArray::~CRegArray()
{
	m_array.RemoveAll();
}

void CRegArray::Dump()
{
	CString key;

	TRACE("FOLDER LIST:\n");
	int ii = 0;
	for (ii = 0; ii < m_array.GetCount(); ii++)
	{
		TRACE("\t%s\n", m_array.GetAt(ii));
	}
}

int CRegArray::Find(CString &str)
{
	int ii = 0;
	for (ii = 0; ii < m_array.GetCount(); ii++)
	{
		if (m_array.GetAt(ii).CompareNoCase(str) == 0)
		{
			return ii;
		}
	}
	return -1;
}

int CRegArray::Add(CString &str)
{
	CString key;
	int indexEmpty = -1;
	int ii = 0;
	for (ii = 0; ii < m_array.GetCount(); ii++)
	{
		if (m_array.GetAt(ii).CompareNoCase(str) == 0)
		{
			return ii;
		}
	}
	for (ii = 0; ii < m_array.GetCount(); ii++)
	{
		if (m_array.GetAt(ii).IsEmpty())
		{
			indexEmpty = ii;
			break;
		}
	}
	if (indexEmpty >= 0)
		ii = indexEmpty;
	key.Format("%d", ii);
	BOOL retval = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section, key, str);
	m_array.SetAtGrow(ii, str);

	return ii;
}

int CRegArray::Delete(CString &str)
{
	CString empty = "";
	CString key;
	int ii = 0;
	for (ii = 0; ii < m_array.GetCount(); ii++)
	{
		if (m_array.GetAt(ii).CompareNoCase(str) == 0)
		{
			m_array.SetAt(ii, empty);
			key.Format("%d", ii);
			BOOL retval = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section, key, empty);
			return ii;
		}
	}
	return -1;
}

int CRegArray::Delete(int index, CString &str)
{
	CString empty = "";
	CString key;
	if ((index >= 0) && (index < m_array.GetCount()))
	{
		if (m_array.GetAt(index).CompareNoCase(str) == 0)
		{
			m_array.SetAt(index, empty);
			key.Format("%d", index);
			BOOL retval = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section, key, empty);
			return index;
		}
		else
			return -1;
	}
	else
		return -1;
}

BOOL CRegArray::GetAt(int index, CString &str)
{
	str.Empty();
	if ((index >= 0) && (index < m_array.GetCount()))
	{
		str = m_array.GetAt(index);
		return TRUE;
	}
	else
		return FALSE;
}

int CRegArray::GetCount()
{
	return m_array.GetCount();
}

BOOL CRegArray::CreateKey(CString &section, HKEY &hKey)
{
	DWORD	dwDisposition;
	HKEY	myKey;

	if (ERROR_SUCCESS == RegCreateKeyEx(HKEY_CURRENT_USER,
		(LPCTSTR)m_section, 0, NULL,
		REG_OPTION_NON_VOLATILE,
		KEY_WRITE, NULL, &myKey,
		&dwDisposition))
	{
		return TRUE;
	}
	else
		return FALSE;
}

BOOL CRegArray::SaveToRegistry()
{
	int ii;
	CString key;
	CString value;
	DWORD	dwDisposition;
	HKEY	myKey;

	if (ERROR_SUCCESS == RegCreateKeyEx(HKEY_CURRENT_USER,
		(LPCTSTR)m_section, 0, NULL,
		REG_OPTION_NON_VOLATILE,
		KEY_WRITE, NULL, &myKey,
		&dwDisposition))
	{
		for (ii = 0; ii < m_array.GetCount(); ii++)
		{
			key.Format("%d", ii);
			CString value = m_array.GetAt(ii);
			LSTATUS sts = RegSetValueEx(myKey, key, 0, REG_SZ, (CONST BYTE*)(LPCTSTR)value, value.GetLength() + 1);
			if (sts == FALSE)
			{
				DWORD err = GetLastError();
			}
			RegCloseKey(myKey);
			return TRUE;
		}
	}
	return FALSE;
}

BOOL CRegArray::LoadFromRegistry()
{
	BOOL ret = LoadFromRegistry(m_array);
	return ret;
}

BOOL CRegArray::LoadFromRegistry(CSArray &ar)
{
	DWORD	size = 4096;
	unsigned char	data[4096];
	int ii;
	data[0] = 0;
	CString key;
	CString val;
	HKEY	myKey;
	CString	result = (char *)("");

	LSTATUS res;
	if (ERROR_SUCCESS == RegOpenKeyEx(HKEY_CURRENT_USER, (LPCTSTR)m_section, NULL, KEY_READ | KEY_QUERY_VALUE, &myKey))
	{
		for (ii = 0; ii < m_nMaxSize; ii++)
		{
			key.Format("%d", ii);
			size = 4096;
			res = RegQueryValueEx(myKey, (LPCTSTR)key, NULL, NULL, data, &size);
			if (res == ERROR_SUCCESS) 
			{
				result = (char *)data;
				val = result;
				ar.SetAtGrow(ii, val);
			}
			else
			{
				break;
			}
		}
		RegCloseKey(myKey);
		return TRUE;
	}
	return FALSE;
}

//

