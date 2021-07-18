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
#include "OpenContainingFolderDlg.h"
#include "TextUtilsEx.h"
#include "MimeParser.h"


#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(NTreeView, CWnd)

/////////////////////////////////////////////////////////////////////////////
// NTreeView

// TODO: Must find time, some code is messy and  needs cleanup !!!


NTreeView::NTreeView()
{
	m_labelHT = new GmailLableMapType(37);
	m_labelView = FALSE;
	m_labelSeqNumb = 0;

	m_bIsDataDirty = FALSE;
	m_fileSizes.InitHashTable(37);
	fileSizes = &m_fileSizes;

	m_bSelectMailFileDone = FALSE;
	m_bSelectMailFilePostMsgDone = FALSE;
	m_bGeneralHintPostMsgDone = FALSE;
	m_timerTickCnt = 0;
	m_nIDEvent = 1;
	m_nElapse = 1;
	m_bInFillControl = FALSE;

	m_frameCx = 177;
	m_frameCy = 200;

	CString m_section = CString(sz_Software_mboxview) + "\\" + "Window Placement";

	BOOL ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "TreeFrameWidth", m_frameCx);
	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "TreeFrameHeight", m_frameCy);
}

NTreeView::~NTreeView()
{
	ClearGlobalFileSizeMap();
	ClearLabelHT();

	delete m_labelHT;
}

void NTreeView::ClearLabelHT()
{
	GmailLableMapType::IHashMapIter iter = m_labelHT->first();
	int totalCnt = 0;
	for (; !m_labelHT->last(iter); )
	{
		//m_labelHT->next(iter);
		GmailLabel *l = iter.element;
		m_labelHT->remove(iter);
		delete l;

	}
	m_labelHT->clear();
}

void NTreeView::ClearGlobalFileSizeMap()
{
	GlobalFileSizeMap::iterator it;

	ArchiveFileInfoMap *item;
	for (it = m_gFileSizes.begin(); it != m_gFileSizes.end(); it++)
	{
		item = it->second;
		delete item;
		it->second = 0;
	}

	m_gFileSizes.clear();
}

BOOL NTreeView::RemoveFileSizeMap(CString path)
{
	GlobalFileSizeMap::iterator it;

	path.TrimRight("\\");
	std::string stdPath = path;
	it = m_gFileSizes.find(stdPath);
	ArchiveFileInfoMap *item;
	if (it != m_gFileSizes.end())
	{
		item = it->second;
		delete item;
		it->second = 0;
		m_gFileSizes.erase(stdPath);
		return TRUE;
	}
	else
	{
		return FALSE;
	}
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
	ON_COMMAND(ID_GMAILLABELS_CREATE, &NTreeView::OnGmaillabelsCreate)
	ON_COMMAND(ID_GMAILLABELS_DELETE, &NTreeView::OnGmaillabelsDelete)
	ON_COMMAND(ID_GMAILLABELS_REFRESH, &NTreeView::OnGmaillabelsRefresh)
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
	
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		int msgViewPosition = pFrame->MsgViewPosition();
		BOOL bTreeHideVal = pFrame->TreeHideValue();
		BOOL isTreeHidden = pFrame->IsTreeHidden();

		TRACE("NTreeView::OnSize cx=%d cy=%d viewPos=%d IsTreeHideVal=%d IsTreeHidden=%d\n",
			cx, cy, msgViewPosition, bTreeHideVal, isTreeHidden);

		if (cx > 0)
		{
			m_frameCx = cx;
			m_frameCy = cy;
		}
	}
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

	// fileNameExtention can contain : .mbox, .mboxo, etc. i.e .mbox....
	CString fileNameExtentionRoot = fileNameExtention.Mid(0, 5);

	if (fileNameExtentionRoot.CompareNoCase(".mboxview") == 0)
		return FALSE;

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
	CString pathLast = MboxMail::GetLastPath();
	// Make sure data path is also set
	MboxMail::SetLastPath(pathLast);
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
			if (path.CompareNoCase(pathLast) == 0)
				continue;

			MboxMail::SetLastPath(path);

			FileUtils::CPathStripPath(path, str);
			sText.Format("Opening %s ...", str);

			if (pFrame)
				pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

			FillCtrl(expand);
#if 0
			// Create per tree node object with all key info
			// Move below to FillCtrl 
			HTREEITEM hFolder;
			if (hFolder = HasFolder(path))
			{
				CString lastDataPath = MboxMail::GetLastDataPath();
				CString lastPath = MboxMail::GetLastPath();

				//int nId = m_labelInfoStore.GetNextId();
				// m_tree.SetItemData(hFolder, nId);
				int nId = m_tree.GetItemData(hFolder);

				LabelInfo *linfo = new LabelInfo(nId, path, lastDataPath);
				m_labelInfoStore.Add(linfo, nId);

				LabelInfo *linfo2 = m_labelInfoStore.Find(nId);
			}
#endif

			sText.Format("Ready");
			if (pFrame)
				pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
			int deb = 1;
		}
	}

	MboxMail::SetLastPath(pathLast);

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

	HTREEITEM hFolder;
	if (hFolder = HasFolder(pathLast))
	{
		m_tree.Expand(hFolder, TVE_EXPAND);
#if 0
		// Create per tree node object with all key info
		// Move below to FillCtrl 
		//HTREEITEM hFolder;
		//if (hFolder = HasFolder(pathLast))
		{
			CString lastDataPath = MboxMail::GetLastDataPath();
			CString lastPath = MboxMail::GetLastPath();

			//int nId = m_labelInfoStore.GetNextId();
			// m_tree.SetItemData(hFolder, nId);
			int nId = m_tree.GetItemData(hFolder);

			LabelInfo *linfo = new LabelInfo(nId, pathLast, lastDataPath);
			m_labelInfoStore.Add(linfo, nId);

			LabelInfo *linfo2 = m_labelInfoStore.Find(nId);
		}
#endif
	}
}

void  NTreeView::ExpandOrCollapseTree(BOOL expand)
{
	CString path = MboxMail::GetLastPath();

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

BOOL NTreeView::SetupFileSizeMap(CString &path)
{
	GlobalFileSizeMap::iterator it;

	path.TrimRight("\\");
	std::string stdPath = path;
	it = m_gFileSizes.find(stdPath);
	ArchiveFileInfoMap *item;
	if (it != m_gFileSizes.end())
	{
		item = it->second;
		fileSizes = &item->m_fileSizes;
		return TRUE;
	}
	else
	{
		_ASSERT(it != m_gFileSizes.end());
		return FALSE;
	}
}

void NTreeView::LoadFileSizes(CString &path, FileSizeMap &fileSizes, BOOL dontUpdateTree)
{
	CString mboxFilePath;
	CString registry_lastPath;
	CString root;
	path.TrimRight("\\");
	char *last_slash = (char*)strrchr(path, '\\');

	// Read first file index file if it exists from previous runs
	// and add to the filesSize hash table
	// new archive files might be discovered and added ??
	m_bIsDataDirty = FALSE;
	fileSizes.RemoveAll();
	CString datapath = MboxMail::GetLastDataPath();
	CStdioFile fp;
	if (fp.Open(datapath + "\\.mboxview", CFile::modeRead | CFile::typeText))
	{
		CString line;
		while (fp.ReadString(line))
		{
			int w = line.Find('\t', 0);
			if (w == -1)
				break;

			CString mboxFileName = line.Left(w);

			_int64	fSize = 0;
			_int64  bShow = 1;

			CString sizeAndShowVal = line.Mid(w + 1);  // get size and show values
			int wShowValPos = sizeAndShowVal.Find('\t', 0);
			if (wShowValPos == -1)  // old format; no mbox show value
			{
				fSize = _atoi64(sizeAndShowVal);
			}
			else
			{
				fSize = _atoi64(sizeAndShowVal.Left(wShowValPos));
				bShow = _atoi64(sizeAndShowVal.Mid(wShowValPos + 1));
				if ((bShow != 0) && (bShow != 1))
					bShow = 1;
			}
			mboxFilePath = path + "\\" + mboxFileName;
			if (FileUtils::PathFileExist(mboxFilePath))
			{
				fileSizes[mboxFileName].fSize = fSize;
				fileSizes[mboxFileName].bShow = bShow;
			}
			else
				m_bIsDataDirty = TRUE;

		}
		fp.Close();
	}
	else
	{
		// TODO: verify implications
		fileSizes.RemoveAll();
	}

	root = last_slash + 1;

	HTREEITEM hRoot = m_tree.InsertItem(root, 0, 0, TVI_ROOT);
	if (hRoot == 0)
		return;

	CString itemName = m_tree.GetItemText(hRoot);

	registry_lastPath = path;
	registry_lastPath.TrimRight("\\");
	registry_lastPath.Append("\\");

	int index = m_folderArray.Add(registry_lastPath);


	HTREEITEM hCurrentSelectedItem = m_tree.GetSelectedItem();
	if (hCurrentSelectedItem > 0)
	{
		itemName = m_tree.GetItemText(hCurrentSelectedItem);
		m_tree.SetItemState(hCurrentSelectedItem, 0, TVIS_BOLD);
	}

	m_tree.SetItemData(hRoot, index);
	DWORD_PTR retIndex = m_tree.GetItemData(hRoot);

	UINT nCode = TVGN_CARET;
	BOOL retval = m_tree.Select(hRoot, nCode);

	CString fw = path + "\\*.*";
	WIN32_FIND_DATA	wf;
	BOOL found;
	// Iterate all archive mbox or eml files in the lastPath folder
	// New archives file is addedd to CTree but not to fileSizes hash table
	// TODO: explain why the size is not added to fileSizes hash table
	// Index file is removed however to force parsing to create new index file
	//
	HANDLE f = FindFirstFile(fw, &wf);
	if (f != INVALID_HANDLE_VALUE)
	{
		do
		{
			if ((wf.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != FILE_ATTRIBUTE_DIRECTORY && wf.cFileName[0] != '.')
			{
				CString fn = wf.cFileName;
				CString mboxFilePath = path + "\\" + fn;

				if (ImboxviewFile(mboxFilePath))
				{
					_int64 fSize = 0;
					ArchiveFileInfo info;
					_int64 realFSize = FileUtils::FileSize(mboxFilePath);
					found = fileSizes.Lookup(fn, info);

					// if found ==  FALSE, the fn is a new file and therefore fSize != realFSize
					// if found ==  TRUE, the fn file exists and fSize == realFSize should match but the fn was changed somehow
					// Currently only the file size is used to make sure the file didn't change, do we need to add checksum ??
					// mboxview will likely crash if file was changed and we didn't detect the change
					if (info.fSize != realFSize)
					{
						//TRACE("File=%s FileSize=%lld StoredFileSize=%lld\n", fn, realFSize, fSize);
						CString cache;
						BOOL retval = MboxMail::GetMboxviewFilePath(mboxFilePath, cache);

						FileUtils::DeleteFile(cache);
						// Delete List Files
						CString listFileName;
						int ret = NListView::DetermineListFileName(fn, listFileName);
						if (!listFileName.IsEmpty())
							FileUtils::DeleteFile(listFileName);

						fileSizes[fn].fSize = realFSize;
						m_bIsDataDirty = TRUE;
						int deb = 1;
					}

					if (info.bShow)
					{
						HTREEITEM hItem = m_tree.InsertItem(fn, 4, 5, hRoot);
						if (hItem)
						{
							//m_tree.SetItemState(hItem, TVIS_BOLD, TVIS_BOLD);
							CString txt = m_tree.GetItemText(hItem);

							int nId = m_labelInfoStore.GetNextId();
							m_tree.SetItemData(hItem, nId);

							LabelInfo *linfo = new LabelInfo(nId, mboxFilePath);
							m_labelInfoStore.Add(linfo, nId);
						}
					}
				}
			}
		} while (FindNextFile(f, &wf));
		FindClose(f);

		CString datapath = MboxMail::GetLastDataPath();
		fw = datapath + "*.mboxview";

		// Delete *.mbox.mboxview and *.eml.mboxview files without associated  mbox or eml archive files
		HANDLE f = FindFirstFile(fw, &wf);
		if (f != INVALID_HANDLE_VALUE)
		{
			do
			{
				if ((wf.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != FILE_ATTRIBUTE_DIRECTORY && wf.cFileName[0] != '.')
				{
					CString fn = wf.cFileName;
					CString mboxviewPath = datapath + fn;

					CString driveName;
					CString directory;
					CString fileNameBase;
					CString fileNameExtention;

					FileUtils::SplitFilePath(fn, driveName, directory, fileNameBase, fileNameExtention);
					CString mboxPath = path + "\\" + fileNameBase;

					ArchiveFileInfo info;
					found = fileSizes.Lookup(fileNameBase, info);
					if (found == FALSE)
					{

						BOOL ret = FileUtils::PathFileExist(mboxPath);
						if (ret == FALSE)
						{
							FileUtils::DeleteFile(mboxviewPath);
						}
					}
				}
			} while (FindNextFile(f, &wf));
			FindClose(f);
		}
	}
	if (m_bIsDataDirty)
	{
		CString itemName = m_tree.GetItemText(hRoot);
		SaveData(hRoot);
		m_bIsDataDirty = FALSE;
	}
}
// Called on Startup, File open and All File Refresh
// Items are inserted into CTree only by this function
void NTreeView::FillCtrl(BOOL expand)
{
	// called from 
	// void CMainFrame::OnFileOpen()
	// void CMainFrame::DoOpen(CString& fpath)
	// void NTreeView::LoadFolders()
	// BOOL NTreeView::DeleteFolder(HTREEITEM hItem)
	//
	// 
	if (m_bInFillControl)
		return;

	m_bInFillControl = TRUE;

	CString path = MboxMail::GetLastPath();

	//m_tree.DeleteAllItems();
	if (path.IsEmpty() || !FileUtils::PathDirExists(path))
	{
		m_bInFillControl = FALSE;
		return;
	}

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	CString fpath = path;
	fpath.TrimRight("\\");
	fpath.Append("\\");

	FileUtils::SplitFilePath(fpath, driveName, directory, fileNameBase, fileNameExtention);
	if (directory.GetLength() <= 1)
	{
		m_bInFillControl = FALSE;
		return;
	}

	// traverse the Tree to see if the folder already exist
	CString itemName;
	HTREEITEM hFolder;
	if (hFolder = HasFolder(path))
	{
		HTREEITEM hCurrentSelectedItem = m_tree.GetSelectedItem();
		if (hCurrentSelectedItem)
			itemName = m_tree.GetItemText(hCurrentSelectedItem);
		if (hCurrentSelectedItem != hFolder)
		{
			if (hCurrentSelectedItem > 0)
				m_tree.SetItemState(hCurrentSelectedItem, 0, TVIS_BOLD);

			UINT nCode = TVGN_CARET;
			if (hFolder)
			{
				BOOL retval = m_tree.Select(hFolder, nCode);
			}
		}

		m_tree.Expand(hFolder, TVE_EXPAND);
		m_bInFillControl = FALSE;
		return;
	}

	CString tmppath = path;
	path.TrimRight("\\");
	std::string stdPath = path;
	if (m_gFileSizes.find(stdPath) == m_gFileSizes.end()) 
	{
		ArchiveFileInfoMap *infoMap = new ArchiveFileInfoMap(path);
		m_gFileSizes.insert(GlobalFileSizeMap::value_type(stdPath, infoMap));
	}

	// TODO: globals
	BOOL ret = SetupFileSizeMap(path);

	if (ret == FALSE)
		int deb = 1;

	LoadFileSizes(path, *fileSizes, FALSE);

	int count = fileSizes->GetCount();
	int deb = 1;
	
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

	m_tree.SetRedraw();

	m_bInFillControl = FALSE;
}

void NTreeView::OnSelchanged(NMHDR* pNMHDR, LRESULT* pResult) 
{
	CString oldItemName;
	CString newItemName;

	NMTREEVIEW *pNm = (LPNMTREEVIEW)pNMHDR;

	HTREEITEM hNewItem = pNm->itemNew.hItem;
	HTREEITEM hOldItem = pNm->itemOld.hItem;

	if (hOldItem)
		oldItemName = m_tree.GetItemText(hOldItem);

	if (hNewItem)
		newItemName = m_tree.GetItemText(hNewItem);

	if (hOldItem == hNewItem)
		return;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if ( pFrame == NULL )
		return;
	NListView *pListView = pFrame->GetListView();
	if ( ! pListView )
		return;

	NMsgView *pMsgView = pFrame->GetMsgView();
	if (pMsgView)
		pMsgView->DisableMailHeader();

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
		if (hOldItem > 0)
			m_tree.SetItemState(hOldItem, 0, TVIS_BOLD);
		m_tree.SetItemState(hNewItem, TVIS_BOLD, TVIS_BOLD);

		CString folderPath;
		DetermineFolderPath(hNewItem, folderPath);

		MboxMail::SetLastPath(folderPath);

		BOOL ret = SetupFileSizeMap(folderPath);

		int count = fileSizes->GetCount();

		if (ret == FALSE)
			int deb = 1;

		pListView->m_path = "";
		pListView->m_which = NULL;
		pListView->ResetSize();
		pListView->FillCtrl();

		m_labelView = FALSE;
		return;
	}


	if (hParent != hRoot)   // label
	{
		DWORD nId = m_tree.GetItemData(hNewItem);

		// Set global variables - nightmare -:)
		CString mboxFileNamePath = MboxMail::s_path;
		CString lastPath = MboxMail::GetLastPath();
		CString lastDataPath = MboxMail::GetLastDataPath();

		LabelInfo *linfo = m_labelInfoStore.Find(nId);
		if ((linfo == 0) || linfo->m_listFilePath.IsEmpty())
		{
			return;
		}

		pListView->m_path_label = linfo->m_listFilePath;
		MboxMail::s_path_label = linfo->m_listFilePath;

		int paneId = 0;
		CString sText;
		sText.Format("Opening Label %s ...", linfo->m_label);
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

		CString driveName;
		CString mboxFileDirectory;
		CString mboxFileNameBase;
		CString mboxFileNameExtention;

		FileUtils::SplitFilePath(linfo->m_filePath, driveName, mboxFileDirectory, mboxFileNameBase, mboxFileNameExtention);
		CString newLastPath = driveName + mboxFileDirectory;

		if (mboxFileNamePath.Compare(linfo->m_filePath))
		{
			MboxMail::SetLastPath(newLastPath);
			CString lastPath = MboxMail::GetLastPath();
			CString lastDataPath = MboxMail::GetLastDataPath();

			pListView->m_path = linfo->m_filePath;
			MboxMail::s_path = linfo->m_filePath;

			CString errorText;
			int rval = pListView->MailFileFillCtrl(errorText);
			MboxMail::nWhichMailList = IDC_ARCHIVE_LIST;
		}

		lastPath = MboxMail::GetLastPath();
		lastDataPath = MboxMail::GetLastDataPath();

		//MboxMail::nWhichMailList = IDC_FOLDER_LIST;
		int retval = pListView->LoadLabelListFile_v2(linfo->m_listFilePath, linfo->m_label);

		if (MboxMail::s_mails.GetCount() > 0)
		{
			pListView->SelectItemFound(0);
		}

		if (hOldItem > 0)
		{
			CString text = m_tree.GetItemText(hOldItem);
			m_tree.SetItemState(hOldItem, 0, TVIS_BOLD);
		}
		m_tree.SetItemState(hNewItem, TVIS_BOLD, TVIS_BOLD);

		m_labelView = TRUE;

		nId = IDC_ARCHIVE_LIST;
		pFrame->SetMailList(nId);

		MboxMail::s_mails_find.SetSizeKeepData(0);
		MboxMail::m_findMails.m_lastSel = -1;
		MboxMail::m_findMails.b_mails_which_sorted = 1;

		sText.Format("Ready");
		pFrame->SetStatusBarPaneText(paneId, sText, FALSE);

		int deb = 1;
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

#if 0
	int nId = IDC_ARCHIVE_LIST;
	BOOL enable = TRUE;
	pFrame->EnableMailList(nId, enable);
#endif

	LabelInfo *linfoOld = 0;
	LabelInfo *linfoNew = 0;
	DWORD nIdOld = -1;
	DWORD nIdNew = -1;


	if (hOldItem)
	{
		nIdOld = m_tree.GetItemData(hOldItem);
		linfoOld = m_labelInfoStore.Find(nIdOld);
	}

	nIdNew = m_tree.GetItemData(hNewItem);
	linfoNew = m_labelInfoStore.Find(nIdNew);

	if (linfoOld && linfoNew)
	{
		if (linfoNew->m_filePath.Compare(linfoOld->m_filePath) == 0)
		{
			m_labelView = FALSE;

			pFrame->OnBnClickedArchiveList();

			if (hOldItem > 0)
			{
				CString text = m_tree.GetItemText(hOldItem);
				m_tree.SetItemState(hOldItem, 0, TVIS_BOLD);
			}
			m_tree.SetItemState(hNewItem, TVIS_BOLD, TVIS_BOLD);

			MboxMail::s_mails_find.SetSizeKeepData(0);
			MboxMail::m_findMails.m_lastSel = -1;
			MboxMail::m_findMails.b_mails_which_sorted = 1; 
			return;
		}
	}

	CString str = m_tree.GetItemText(hNewItem);

	CString folderPath;
	DetermineFolderPath(hNewItem, folderPath);

	MboxMail::SetLastPath(folderPath);
	CString path = MboxMail::GetLastPath();

	if( str.IsEmpty() || path.IsEmpty() )
		return;

	BOOL ret = SetupFileSizeMap(path);

	int count = fileSizes->GetCount();

	if (ret == FALSE)
		int deb = 1;

	path.TrimRight("\\");
	pListView->m_path = path + _T('\\') + str;

	if (hOldItem > 0)
	{
		CString text = m_tree.GetItemText(hOldItem);
		m_tree.SetItemState(hOldItem, 0, TVIS_BOLD);
	}
	m_tree.SetItemState(hNewItem, TVIS_BOLD, TVIS_BOLD);


	// TODO: relying on m_which is risky. 
	// Consider to eliminate and rely on strings and searches, assuming minimum impact on performance
	pListView->m_which = hNewItem;

	pListView->ResetSize();

	pListView->ResetFilterDates();

	int paneId = 0;
	CString sText;
	sText.Format("Opening %s ...", str);
	pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	pListView->FillCtrl();
	MboxMail::nWhichMailList = IDC_ARCHIVE_LIST;

	if (MboxMail::s_mails.GetCount() > 0)
	{
		pListView->SelectItemFound(0);
	}

	pFrame->UpdateFilePrintconfig();

	sText.Format("Ready");
	pFrame->SetStatusBarPaneText(paneId, sText, FALSE);

	MboxMail::ShowHint(HintConfig::MsgWindowPlacementHint, GetSafeHwnd());

	//pListView->ScanAllMailsInMbox();
	//pListView->ScanAllMailsInMbox_NewParser();

	m_labelView = FALSE;

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
	CString path = MboxMail::GetLastPath();
	if (str.IsEmpty() || path.IsEmpty())
		return;
	path.TrimRight("\\");
	pListView->m_path = path + _T('\\') + str;

	CString cache;
	BOOL retval = MboxMail::GetMboxviewFilePath(pListView->m_path, cache);

	FileUtils::DeleteFile(cache);
	pListView->m_which = hItem;
	pListView->ResetSize();
	pListView->FillCtrl();
	MboxMail::nWhichMailList = IDC_ARCHIVE_LIST;
}


// Called when user specified command line option MAIL_FILE=
// TODO: 
void NTreeView::SelectMailFile(CString *fileNm)
{
	// mailFile can be mailFilePath or mailFileName
	CString mailFile;
	if (fileNm)
	{
		mailFile.Append(*fileNm);
		//delete fileNm;
	}
	else
		mailFile = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("mailFile"));

	CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("mailFile"), CString(""));

	if (mailFile.IsEmpty())
		return;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame == NULL)
		return;

	NListView *pListView = pFrame->GetListView();
	if (!pListView)
		return;

	CString mboxFilePath;
	CString mboxFileName;
	FileUtils::GetFolderPathAndFileName(mailFile, mboxFilePath, mboxFileName);

	mboxFilePath.TrimRight("\\");
	mboxFilePath.Append("\\");

	CString txt;
	if (!FileUtils::PathFileExist(mailFile))
	{
		if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
		{
			txt = _T("Nonexistent Mail File \"") + mailFile;
			txt += _T("\".\nTerminating");
			int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
			AfxGetMainWnd()->PostMessage(WM_CLOSE);
		}
		else
		{
			txt = _T("Nonexistent Mail File \"") + mailFile;
			txt += _T("\".\nDo you want to continue?");
			int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
			if (answer == IDNO)
				AfxGetMainWnd()->PostMessage(WM_CLOSE);
		}
	}
	else 
	{
		pListView->m_path = mailFile;

		pListView->CloseMailFile();

		MboxMail::SetLastPath(mboxFilePath);

		HTREEITEM hFolder = HasFolder(mboxFilePath);
		if (hFolder)
		{
			HTREEITEM hItem = NTreeView::FindItem(hFolder, mboxFileName);
			UINT nCode = TVGN_CARET;
			if (hItem)
			{
				BOOL retval = m_tree.Select(hItem, nCode);

				if (MboxMail::s_mails.GetCount() > 0)
				{
					pListView->SelectItemFound(0);
					pListView->SetListFocus();
				}
			}
		}
	}
}

// TODO: Works but need to rewrite, it is too complex. Need to review  all globals  to reduce complexity,
// make reliable and maintainable
void NTreeView::InsertMailFile(CString &mailFile)
{
	// mailFile can be mailFilePath or mailFileName
	if (mailFile.IsEmpty())
		return;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame == NULL)
		return;
	NListView *pListView = pFrame->GetListView();
	if (!pListView)
		return;

	CString fileName;
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;
	FileUtils::SplitFilePath(mailFile, driveName, directory, fileNameBase, fileNameExtention);

	CString mailFileName = fileNameBase + fileNameExtention;

	CString path = MboxMail::GetLastPath();
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
		// TODO:
		return;
	}
	else
	{
		HTREEITEM hFolder = HasFolder(path);
		if (hFolder)
		{
			HTREEITEM hItem = NTreeView::FindItem(hFolder, mailFileName);
			UINT nCode = TVGN_CARET;
			if (hItem == 0)
			{
				(*fileSizes)[mailFileName].bShow = 1;
				_int64 fSize = FileUtils::FileSize(pListView->m_path);
				(*fileSizes)[mailFileName].fSize = fSize;
				SaveData(hFolder);

				CString mboxIndexFilepath; // = pListView->m_path + ".mboxview";
				BOOL retval = MboxMail::GetMboxviewFilePath(pListView->m_path, mboxIndexFilepath);
				BOOL ret = FileUtils::DeleteFile(mboxIndexFilepath);

				HTREEITEM hItem = m_tree.InsertItem(mailFileName, 4, 5, hFolder);
				if (hItem)
					BOOL retval = m_tree.Select(hItem, nCode);
				int deb = 1;
			}
		}
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

void NTreeView::Traverse( HTREEITEM hItem, CFile &fp, FileSizeMap &fileSizes)
{
#if 1
	CString line;
	CString path;

	ArchiveFileInfo info;
	POSITION pos = fileSizes.GetStartPosition();
	while (pos) 
	{
		fileSizes.GetNextAssoc(pos, path, info);

		// or for better performance; not really crusial here
		// CMap<CString, LPCSTR, ArchiveFileInfo, ArchiveFileInfo>::CPair *infopair = fileSizes.PLookup(path);
		// _int64 fSize = infopair->value.fSize;
		// _int64 bShow = infopair->value.bShow;

		_int64 fSize = fileSizes[path].fSize;
		_int64 bShow = fileSizes[path].bShow;

		line.Format("%s\t%lld\t%lld\n", path, fSize, bShow);
		fp.Write(line, line.GetLength());
	}
#else
	CString line;
	CString ipath;
	CString path = MboxMail::GetLastPath();
	path.TrimRight("\\");
	while( hItem != NULL ) 
	{
		ipath = m_tree.GetItemText(hItem);
		if( ! ipath.IsEmpty() && ! path.IsEmpty() ) 
		{
			CString fullPath = path + "\\" + ipath;
			_int64 fSize = 0;
			ArchiveFileInfo info;
			//_int64 realFSize = FileUtils::FileSize(fullPath);
			BOOL found = fileSizes.Lookup(ipath, info);
			//TRACE("File=%s FileSize=%lld StoredFileSize=%lld\n", path, realFSize, fSize);
			if (!found) 
			{
				_int64 realFSize = FileUtils::FileSize(fullPath);
				fileSizes[ipath].fSize = fSize = realFSize;
			}
			fSize = fileSizes[ipath].fSize;
			_int64 bShow = fileSizes[ipath].bShow;
			line.Format("%s\t%lld\t%lld\n", ipath, fSize, bShow);
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
#endif
}


// Thhis function creates or updates per folder .mboxview index file
// It is called when selcting new mail archive file or by RedrawMails called after sorting
// or touching mail list
void NTreeView::SaveData(HTREEITEM m_which)
{
	CFile fp;
	CString path = MboxMail::GetLastPath();
	if( path.IsEmpty() )
		return;
	path.TrimRight("\\");

	HTREEITEM hRoot = HasFolder(path);
	if (hRoot == 0)
		return;

	CString rootName = m_tree.GetItemText(hRoot);

	CString datapath = MboxMail::GetLastDataPath();
	FileUtils::CreateDirectory(datapath);
	datapath.TrimRight("\\");

	if( fp.Open(datapath+"\\.mboxview", CFile::modeWrite | CFile::modeCreate) ) 
	{
		Traverse(hRoot, fp, *fileSizes);
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


// Not used anymore
void NTreeView::UpdateFileSizesTable(CString &path, _int64 realFSize, FileSizeMap &fileSizes)
{
	_int64 fSize = 0;
	ArchiveFileInfo info;
	BOOL found = fileSizes.Lookup(path, info);
	if (found)
	{
		if (info.fSize != realFSize) 
		{
			fileSizes.RemoveKey(path);
			fileSizes[path].fSize = realFSize;
			fileSizes[path].bShow = 1;
		}
	}
	else
	{
		fileSizes[path].fSize = realFSize;
		fileSizes[path].bShow = 1;
	}
}

void NTreeView::OnTreeExpand()
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (pFrame)
		pFrame->UpdateToolsBar();
}

void NTreeView::OnUpdateTreeExpand(CCmdUI* pCmdUI)
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	BOOL isTreeHidden = FALSE;
	if (pFrame)
	{
		isTreeHidden = pFrame->IsTreeHidden();
	}

	BOOL enable = (m_tree.GetRootItem() != NULL) && !isTreeHidden;
	pCmdUI->Enable(enable);
}

void NTreeView::OnUpdateFileRefresh(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(m_tree.GetRootItem() != NULL);
}

void NTreeView::OnFileRefresh() 
{
	if (!CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
	{
		m_tree.DeleteAllItems();
		LoadFolders();
	}
	return;
}

void AppendMenu(CMenu *menu, int commandId, const char *commandName, BOOL checkMark)
{
	UINT nFlags = MF_STRING;
	if (checkMark)
		nFlags = MF_STRING | MF_CHECKED;
	menu->AppendMenu(nFlags, commandId, commandName);
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
		const UINT M_FolderPath_Id = 1;
		AppendMenu(&menu, M_FolderPath_Id, _T("Show Folder Path"));
		const UINT M_DeleteItem_Id = 2;
		AppendMenu(&menu, M_DeleteItem_Id, _T("Remove Folder"));
		const UINT M_FolderRefresh_Id = 3;
		AppendMenu(&menu, M_FolderRefresh_Id, _T("Refresh Folder"));
		const UINT M_OpenHiddenFiles_Id = 4;
		AppendMenu(&menu, M_OpenHiddenFiles_Id, _T("Restore Removed Files"));

		CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

		UINT command = menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, this);

		UINT nFlags = TPM_RETURNCMD;
		CString menuString;
		int chrCnt = menu.GetMenuString(command, menuString, nFlags);

		switch (command)
		{
		case M_DeleteItem_Id: 
		{
			UINT nFlags = MF_BYCOMMAND;
			CString Label;
			int retLabel = menu.GetMenuString(M_DeleteItem_Id, Label, nFlags);

			MboxMail::ShowHint(HintConfig::MessageRemoveFolderHint, GetSafeHwnd());
			if (hItem)
				BOOL ret = DeleteFolder(hItem);
		}
		break;
		case M_FolderPath_Id: 
		{
			CString pathLast = MboxMail::GetLastPath();
			pathLast.TrimRight("\\");
			int nId = m_tree.GetItemData(hItem);
			CString path = m_folderArray.m_array.GetAt(nId);
			if (path.Compare(pathLast))
				int deb = 1;
			HWND h = wnd->GetSafeHwnd();
			int answer = ::MessageBox(h, path, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_USERICON);
		}
		 break;
		case M_FolderRefresh_Id: 
		{
			CString path = MboxMail::GetLastPath();
			path.TrimRight("\\");
			HWND h = wnd->GetSafeHwnd();
			//int answer = ::MessageBox(h, path, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_USERICON);
			if (hItem)
			{
				RefreshFolder(hItem);
			}
		}
		break;

		case M_OpenHiddenFiles_Id: 
		{
			if (hItem)
			{
				int ret = OpenHiddenFiles(hItem, *fileSizes);
			}
		}
		break;

		default: {
				int deb = 1;
		}
		break;
		}

		return;
	}

	BOOL isLabel = FALSE;
	if (hItem)
	{
		DWORD nId = m_tree.GetItemData(hItem);
		LabelInfo *linfo = m_labelInfoStore.Find(nId);
		if (linfo && !linfo->m_listFilePath.IsEmpty())
			isLabel = TRUE;
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
	//AppendMenu(&sortSubMenu, S_SORT_BY_POSITION_Id, _T("Mail ID"));

	menu.AppendMenu(MF_POPUP | MF_STRING, (UINT)sortSubMenu.GetSafeHmenu(), _T("Sort By"));
	menu.AppendMenu(MF_SEPARATOR);
	//

	const UINT M_FileLocation_Id = 11;
	AppendMenu(&menu, M_FileLocation_Id, _T("Open File Location"));

	const UINT M_DatabaseLocation_Id = 12;
	AppendMenu(&menu, M_DatabaseLocation_Id, _T("Open Database Folder Location"));

	const UINT M_Properties_Id = 13;
	AppendMenu(&menu, M_Properties_Id, _T("Properties"));

	const UINT M_AttachmentCache_Id = 14;
	AppendMenu(&menu, M_AttachmentCache_Id, _T("Export All Mail Attachments"));

	const UINT M_EmlCache_Id = 15;
	AppendMenu(&menu, M_EmlCache_Id, _T("Export All Mails as Eml"));

	const UINT M_Reload_Id = 16;
	const UINT M_Remove_Id = 17;
	const UINT S_Labels_Create_Id = 18;
	const UINT S_Labels_Delete_Id = 19;
	const UINT S_Labels_Refresh_Id = 20;
	const UINT M_CreateFolder_Id = 21;

	CMenu labelsSubMenu;
	labelsSubMenu.CreatePopupMenu();

	if (isLabel == FALSE)
	{
		AppendMenu(&menu, M_Reload_Id, _T("Refresh Index File"));

		AppendMenu(&menu, M_Remove_Id, _T("Remove File"));

		AppendMenu(&labelsSubMenu, S_Labels_Create_Id, _T("Create..."));
		AppendMenu(&labelsSubMenu, S_Labels_Delete_Id, _T("Delete..."));
		AppendMenu(&labelsSubMenu, S_Labels_Refresh_Id, _T("Refresh..."));

		menu.AppendMenu(MF_POPUP | MF_STRING, (UINT)labelsSubMenu.GetSafeHmenu(), _T("Gmail Labels"));
		menu.AppendMenu(MF_SEPARATOR);

		//AppendMenu(&menu, M_CreateFolder_Id, _T("Create Folder"));  // TODO: later
	}

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
	case M_FileLocation_Id: 
	{
		UINT nFlags = MF_BYCOMMAND;
		CString Label;
		int retLabel = menu.GetMenuString(M_FileLocation_Id, Label, nFlags);

		CString path = MboxMail::GetLastPath();
		if (FileUtils::BrowseToFile(MboxMail::s_path) == FALSE) {  // TODO: s_path error checking ??
			HWND h = GetSafeHwnd();
			HINSTANCE result = ShellExecute(h, _T("open"), path, NULL, NULL, SW_SHOWNORMAL);
			CMainFrame::CheckShellExecuteResult(result, h);
		}
	}
	break;
	case M_DatabaseLocation_Id:
	{
		UINT nFlags = MF_BYCOMMAND;
		CString Label;
		int retLabel = menu.GetMenuString(M_DatabaseLocation_Id, Label, nFlags);

		CString path = MboxMail::GetLastDataPath();
		//if (FileUtils::BrowseToFile(MboxMail::s_path) == FALSE) {  // TODO: s_path error checking ??
		{
			HWND h = GetSafeHwnd();
			HINSTANCE result = ShellExecute(h, _T("open"), path, NULL, NULL, SW_SHOWNORMAL);
			CMainFrame::CheckShellExecuteResult(result, h);
		}
	}
	break;
	case M_Properties_Id: 
	{
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
	case S_CSV_Id: 
	{
		if (pFrame) {
			pFrame->OnFileExportToCsv();
		}
		int deb = 1;
	}
	break;
	case S_TEXT_Id: 
	{
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
	case S_HTML_Id: 
	{
		if (pFrame) 
		{
			pFrame->PrintMailArchiveToHTML();
		}

		int deb = 1;
	}
	break;
	case S_PDF_Id: 
	{
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

	case M_Reload_Id: 
	{
		CString txt = _T("Do you want to refresh index file?");
		HWND h = GetSafeHwnd();
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDYES)
			ForceParseMailFile(hItem);
		// For internal testing
		//int ret = MboxMail::DumpMailStatsToFile(&MboxMail::s_mails, MboxMail::s_mails.GetCount());
	}
	break;

	case M_AttachmentCache_Id: 
	{
		CString txt = _T("Do you want to create cache with all attachements?");
		txt.Append("\n\nNote that \"File->Attachments Config\" controls whether image attachments embedded into messages will also be exported");
		HWND h = GetSafeHwnd();
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDYES)
		{
			int firstMail = 0;
			int lastMail = MboxMail::s_mails.GetCount() - 1;

			CString targetPrintSubFolderName = "";
			if (pListView)
			{ 
				int retval = pListView->CreateAttachmentCache_Thread(firstMail, lastMail, targetPrintSubFolderName);
			}
		}
	}
	break;

	case M_EmlCache_Id: 
	{
		CString txt = _T("Do you want to create cache with all Eml files?");
		HWND h = GetSafeHwnd();
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDYES)
		{
			int firstMail = 0;
			int lastMail = MboxMail::s_mails.GetCount() - 1;

			CString targetPrintSubFolderName = "";
			if (pListView)
			{
				int retval = pListView->CreateEmlCache_Thread(firstMail, lastMail, targetPrintSubFolderName);
			}
		}
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

	case M_Remove_Id:
	{
		CString mailFile = m_tree.GetItemText(hItem);
		RemoveFileFromTreeView(hItem, *fileSizes);
	}
	break;

	case S_Labels_Create_Id:
	{
		//int retval = CreateLabelFilesForSingleMailFile(hItem);
		int retval = RefreshLabelsForSingleMailFile(hItem);
	}
	break;
	case S_Labels_Delete_Id:
	{
		int retval = DeleteLabelsForSingleMailFile(hItem);
	}
	break;
	case S_Labels_Refresh_Id:
	{
		int retval = RefreshLabelsForSingleMailFile(hItem);
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

	BOOL res = *pResult;
	*pResult = FALSE;

	if (pFrame == NULL)
		return;
	NListView *pListView = pFrame->GetListView();
	if (!pListView)
		return;
	HTREEITEM hNewItem = pNm->itemNew.hItem;
	HTREEITEM hOldItem = pNm->itemOld.hItem;

	DWORD nIdNew = m_tree.GetItemData(hNewItem);
	LabelInfo *linfoNew = m_labelInfoStore.Find(nIdNew);

	if (hOldItem == 0)
		return;

	DWORD nIdOld = m_tree.GetItemData(hOldItem);
	LabelInfo *linfoOld = m_labelInfoStore.Find(nIdOld);

	if (linfoNew && linfoOld)
	{
		if (linfoNew->m_filePath.Compare(linfoOld->m_filePath) == 0)
			return;
	}

	BOOL isDirty = MboxMail::m_editMails.m_bIsDirty;
	if (!pListView->IsUserSelectedMailListEmpty())
	{
		if (isDirty)
		{
			CString txt = _T("User Selected Mails list is not empty.\n"
				"Content will be lost if you switch to new mail archive.\n"
				"Do you want to continue?");
			int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
			if (answer == IDNO)
			{
				*pResult = TRUE;
			}
			else if (answer == IDYES)
				MboxMail::m_editMails.m_bIsDirty = FALSE;
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

	RemoveFileSizeMap(folderPath);

	itemTxt = m_tree.GetItemText(hItem);
	TRACE("Deleting %s\n", itemTxt);
	m_tree.DeleteItem(hItem);

	HTREEITEM hRootItem = m_tree.GetRootItem();
	if (hRootItem)
	{
		CString folderPath;
		DetermineFolderPath(hRootItem, folderPath);
		MboxMail::SetLastPath(folderPath);
		FillCtrl(FALSE);
	}
	else
	{
		CString empty = "";
		MboxMail::SetLastPath(empty);
	}
	return ret;
}

int NTreeView::GetChildrenCount(HTREEITEM hItem, BOOL recursive)
{
	if (hItem == 0)
		return -1;

	int count = 0;
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
				count++;
		}
		else
		{
			itemTxt = m_tree.GetItemText(hNextItem);
			count++;
		}
		hNextItem = hNextNextItem;
	}
	return count;
}

BOOL NTreeView::DeleteItemChildren(HTREEITEM hItem)
{
	if (hItem == 0)
		return FALSE;

	UINT nCode = TVE_COLLAPSE;
	m_tree.Expand(hItem, nCode);

	CString itemTxt;
	itemTxt = m_tree.GetItemText(hItem);

	HTREEITEM hNextNextItem;
	HTREEITEM hNextItem = m_tree.GetChildItem(hItem);

	while (hNextItem != NULL)
	{
		hNextNextItem = m_tree.GetNextSiblingItem(hNextItem);
		if (m_tree.ItemHasChildren(hNextItem))
		{
			NTreeView::DeleteItem(hNextItem);
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
	//UINT nCode = TVE_COLLAPSE;
	//m_tree.Expand(hItem, nCode);

	itemTxt = m_tree.GetItemText(hItem);

	HTREEITEM hNextNextItem;
	HTREEITEM hNextItem = m_tree.GetChildItem(hItem);

	while (hNextItem != NULL)
	{
		hNextNextItem = m_tree.GetNextSiblingItem(hNextItem);
		itemTxt = m_tree.GetItemText(hItem);
		if (hNextItem)
			itemTxt = m_tree.GetItemText(hNextItem);

		if (m_tree.ItemHasChildren(hNextItem))
		{
			NTreeView::DeleteItem(hNextItem);
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

void NTreeView::PostMsgCmdParamFileName(CString *fileName)
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
	// This is just one time message;
	// TO: Fix it
	//m_bGeneralHintPostMsgDone = FALSE;
	return 0;
}

LRESULT NTreeView::OnCmdParam_GeneralHint(WPARAM wParam, LPARAM lParam)
{
	MboxMail::ShowHint(HintConfig::GeneralUsageHint, GetSafeHwnd());
	// This is just one time message;
	// TO: Fix it
	//m_bGeneralHintPostMsgDone = FALSE;
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
		

	//HTREEITEM newItem = m_tree.InsertItem(newFolderName, hItem, TVI_SORT);
	HTREEITEM newItem = m_tree.InsertItem(newFolderName, 6,6,hItem, TVI_SORT);
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

void NTreeView::RemoveFileFromTreeView(HTREEITEM hItem, FileSizeMap &fileSizes)
{
	HWND h = GetSafeHwnd();
	MboxMail::ShowHint(HintConfig::MessageRemoveFileHint, h);

	ArchiveFileInfo info;
	CString mailFile = m_tree.GetItemText(hItem);
	BOOL found = fileSizes.Lookup(mailFile, info);
	if (found)
	{
		fileSizes[mailFile].bShow = 0;
		BOOL ret = m_tree.DeleteItem(hItem);
		SaveData(hItem);
	}
	return;
}

BOOL NTreeView::RefreshFolder(HTREEITEM hItem)
{
	CString path = MboxMail::GetLastPath();

	BOOL ret = DeleteFolder(hItem);

	MboxMail::SetLastPath(path);

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	int paneId = 0;
	CString sText;
	sText.Format("Opening new mail folder ...");
	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	FillCtrl();

	m_tree.SortChildren(0);

	HTREEITEM hFolder;
	if (hFolder = HasFolder(path))
	{
		m_tree.SortChildren(hFolder);
		m_tree.Expand(hFolder, TVE_EXPAND);
	}

	sText.Format("Ready");
	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
	return TRUE;
}


int NTreeView::OpenHiddenFiles(HTREEITEM hItem, FileSizeMap &fileSizes)
{
	CCheckListBoxDlg dlg;
	dlg.m_title = "Select Files to Restore";

	FileSizeMap::CPair *pCurVal;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	CString pathLast = MboxMail::GetLastPath();

	int nId = m_tree.GetItemData(hItem);
	CString path = m_folderArray.m_array.GetAt(nId);
	if (path.Compare(pathLast))
		int deb = 1;

	path.TrimRight("\\");
	path.AppendChar('\\');

	CString mboxFileName;
	pCurVal = fileSizes.PGetFirstAssoc();
	while (pCurVal != NULL)
	{
		if (pCurVal->value.bShow == 0)
		{
			mboxFileName = path + pCurVal->key;
			if (FileUtils::PathFileExist(mboxFileName))
			{
				dlg.m_InList.Add(pCurVal->key);
			}
		}
		pCurVal = fileSizes.PGetNextAssoc(pCurVal);
	}
	//
	int nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		int i;
		for (i = 0; i < dlg.m_OutList.GetCount(); i++)
		{
			CString &s = dlg.m_OutList[i];
			CString filePath = path + s;
			fileSizes[s].bShow = 1;

			HTREEITEM hItemRet = m_tree.InsertItem(s, 4, 5, hItem);

			int nId = m_labelInfoStore.GetNextId();
			m_tree.SetItemData(hItemRet, nId);


			LabelInfo *linfo = new LabelInfo(nId, filePath);
			m_labelInfoStore.Add(linfo, nId);


			// TODO: resolve later; manual action is required for now
			//DisplayGmailLabels(hItem);
			int deb = 1;
		}
		SaveData(hItem);
		m_tree.SortChildren(hItem);
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with Cancel
	}
	return 1;
}

BOOL NTreeView::RecreateGmailLabels(HTREEITEM hItem)
{
	int ret = CreateGmailLabelFiles(hItem);
	ret = DisplayGmailLabels(hItem);
	return ret;
}

int  NTreeView::CreateGmailLabelFiles(HTREEITEM hItem)
{
	HWND h = GetSafeHwnd();
	//MboxMail::ShowHint(HintConfig::MessageRemoveFileHint, h);

	static const char *cLabels = "x-gmail-labels";
	static const int cLabelsLen = strlen(cLabels);
	static const char *cOsobiste = "osobiste";
	static const int cOsobisteLen = strlen(cOsobiste);

	CString comma = ",";

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	NListView *pListView = 0;
	if (pFrame)
		pListView = pFrame->GetListView();

	if (pListView == 0) {
		MboxMail::assert_unexpected();
		return -1;
	}

	ClearLabelHT();

	CString mailFile = m_tree.GetItemText(hItem);
	int mailCnt = MboxMail::s_mails.GetCount();

	const int buffSmallSize = 256;
	SimpleString buff(buffSmallSize + 1);
	CString validFileName;
	//MboxMail *m_s;
	MboxMail *m;
	CString labels;
	CString val;
	CString v;
	CString el;
	CString line;
	int i;

	int maxNoLabelCnt = 100;
	BOOL foundLabels = FALSE;

	char *p = buff.Data();
	char *e = p + buff.Count();

	CFile fp;
	if (!fp.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite))
	{
		DWORD err = GetLastError();
		TRACE("Open Mail File failed err=%ld\n", err);
		return -1;
	}

	if (MboxMail::s_mails.GetCount() != MboxMail::s_mails_ref.GetCount())
		int deb = 1;

	int noLabelCnt = 0;
	for (i = 0; i < mailCnt; i++)
	{
		//m_s = MboxMail::s_mails[i];
		m = MboxMail::s_mails_ref[i];

		m->GetBody(fp, &buff, buffSmallSize);

		p = buff.Data();
		e = p + buff.Count();

		val.Empty();
		while (p < e)
		{
			if (TextUtilsEx::strncmpUpper2Lower(p, e, cLabels, cLabelsLen) == 0)
			{
				foundLabels = TRUE;
				p = MimeParser::GetMultiLine(p, e, line);
				val = line.Mid(cLabelsLen);
				val.TrimLeft(": ");
				val.TrimRight(" ");
				//TRACE("Labels=%s\n", val);
				break;
			}
			p = MimeParser::EatNewLine(p, e);
		}

		if (foundLabels == FALSE)
		{
			noLabelCnt++;
			int deb = 1;
		}

		if ((foundLabels == FALSE) && val.IsEmpty() && (i > maxNoLabelCnt))  // it looks like it is not gmail  file
		{
			fp.Close();
			return 1;
		}

		CStringArray a;
		TextUtilsEx::Tokenize(val, a);

		// remove duplicate labels
		int first = 0;
		while (first < a.GetSize())
		{
			el = a.ElementAt(first);
			for (int i = first + 1; i < a.GetSize(); i++)
			{
				v = a.ElementAt(i);
				if (el.Compare(v) == 0)
				{
					a.RemoveAt(i);
					break;
				}
			}
			first++;
		}

		// iterate all labels and add mail ptr to respectfull list/array
		for (int i = 0; i < a.GetSize(); i++)
		{
			v = a.ElementAt(i);
			v.Trim();
			//TRACE("|%s|\n", v);

#if 1
			char *pp = (char*)(LPCSTR)v;
			char *ee = pp + v.GetLength();
			if (TextUtilsEx::strncmpUpper2Lower(pp, ee, cOsobiste, cOsobisteLen) == 0)
				int deb = 1;
#endif

			validFileName = v;

			FileUtils::MakeValidFilePath(v, FALSE);
			if (validFileName.Compare(v))
			{
				CString seqNumb;
				seqNumb.Format(" %d", m_labelSeqNumb++);
				v += seqNumb;
			}
			{

				CString delim = "/";
				CStringArray va;
				TextUtilsEx::SplitString(v, delim, va);

				v.Empty();
				for (int i = 0; i < va.GetSize(); i++)
				{
					el = va.ElementAt(i);
					el.Trim();
					if (i == 0)
						v += el;
					else
						v += "/" + el;
				}

				// va array no longer needed
				int deb = 1;
			}

			GmailLabel *glabel = m_labelHT->find(&v);

			//TRACE("%u\n", glabel);
			if (glabel == 0)
			{
				glabel = new GmailLabel(v);
				glabel->m_ptrList->Add(m);
				m_labelHT->insert(&glabel->m_label, glabel);
			}
			else
			{
				glabel->m_ptrList->Add(m);
			}
			int deb = 1;
		}
	}
	fp.Close();

	CString labelsCachePath;
	CString rootPrintSubFolder = "LabelCache";
	CString errorText;
	CString targetPrintSubFolder;
	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, labelsCachePath, errorText);
	if (retval == FALSE) {
		MboxMail::assert_unexpected();
		return -1;
	}

	GmailLableMapType::IHashMapIter iter = m_labelHT->first();

	// find the size of largest list of mails
	int maxCnt = 0;
	for (; !m_labelHT->last(iter); )
	{
		int cnt = iter.element->m_ptrList->Count();
		if (cnt > maxCnt)
			maxCnt = cnt;
		m_labelHT->next(iter);
	}

	MboxMail::s_mails_label.SetSizeKeepData(maxCnt);

	// iterate labels and create folder and .mboxlist files for each label
	iter = m_labelHT->first();
	int totalCnt = 0;
	for (; !m_labelHT->last(iter); )
	{
		iter.element->m_ptrList->Assert();
		TRACE("Name=%s  cnt=%d\n", iter.element->m_label, iter.element->m_ptrList->Count());

		int cnt = iter.element->m_ptrList->Count();
		totalCnt += cnt;

		MboxMail::s_mails_label.SetSizeKeepData(cnt);
		for (i = 0; i < cnt; i++)
		{
			MboxMail::s_mails_label[i] = iter.element->m_ptrList->Get(i);
		}

		ASSERT(MboxMail::s_mails_label.GetCount() <= mailCnt);

		int r = pListView->SaveAsLabelFile(&MboxMail::s_mails_label, labelsCachePath, iter.element->m_label, errorText);
		if (r < 0) {
			MboxMail::assert_unexpected();
			return -1;
		}

		m_labelHT->next(iter);
	}
	return 1;
}

int NTreeView::DisplayGmailLabels(HTREEITEM hItem)
{
	CString name;
	LabelInfo *linfo;

	// add labels to Tree
	DWORD nId = m_tree.GetItemData(hItem);

	linfo = m_labelInfoStore.Find(nId);
	if (linfo == 0)
		return -1;

	if (!linfo->m_filePath.IsEmpty())
	{
		name = m_tree.GetItemText(hItem);
		//nId = m_tree.GetItemData(hItem);

		//linfo = m_labelInfoStore.Find(nId);

		CString labelsCachePath;
		CString rootPrintSubFolder = "LabelCache";
		CString errorText;
		CString targetPrintSubFolder;

		MboxMail::s_path = linfo->m_filePath;

		CString driveName;
		CString directory;
		CString fileNameBase;
		CString fileNameExtention;
		FileUtils::SplitFilePath(linfo->m_filePath, driveName, directory, fileNameBase, fileNameExtention);

		CString lastPath = driveName + directory;

		MboxMail::SetLastPath(lastPath);

		BOOL retval = MboxMail::GetCachePath(rootPrintSubFolder, targetPrintSubFolder, labelsCachePath, errorText, &linfo->m_filePath);
		if (retval == FALSE) {
			return -1;
		}
		if (!labelsCachePath.IsEmpty())
		{
			ShowGmailLabels(hItem, labelsCachePath, linfo->m_filePath);
			CString inbox("Inbox");
			CString sent("Sent");
			MoveLabelItem(linfo->m_filePath, sent);
			MoveLabelItem(linfo->m_filePath, inbox);
		}
	}
	//ASSERT(totalCnt >= mailCnt);
	return 1;
}

GmailLabel::GmailLabel(CString &str)
{
	m_label = str;
	m_ptrList = new MySimpleDeque;
}

GmailLabel::~GmailLabel()
{
	delete m_ptrList;
}

MySimpleDeque::MySimpleDeque()
{
	m_ar = 0;
	m_arcnt = 0;
}

MySimpleDeque::~MySimpleDeque()
{
	std::vector <LabelArray*>::iterator it;
	LabelArray *l;
	for (it = m_arList.begin(); it != m_arList.end(); it++)
	{
		l = *it;
		delete l;
	}
	delete m_ar;
}

BOOL MySimpleDeque::Assert()
{
	int cnt = 0;
	if (m_ar == 0)
		return TRUE;

	std::vector <LabelArray*>::iterator it;
	LabelArray *l;
	for (it = m_arList.begin(); it != m_arList.end(); it++)
	{
		l = *it;
		//ASSERT(m_arcnt == LSIZE);
	}
	int lcnt = Count();
	int mailCnt = MboxMail::s_mails.GetCount();
	//ASSERT(lcnt <= mailCnt);
	return TRUE;
}

void  MySimpleDeque::Clear()
{
	m_arcnt = 0;
	std::vector <LabelArray*>::iterator it;
	LabelArray *l;
	for (it = m_arList.begin(); it != m_arList.end(); it++)
	{
		l = *it;
		delete l;
	}
	m_arList.clear();
}

int MySimpleDeque::Count()
{
	int cnt = 0;
	if (m_ar == 0)
		return 0;

	cnt = m_arList.size() * LSIZE + m_arcnt;

	//this->Assert();

	return cnt;
}

void MySimpleDeque::Add(MboxMail* m)
{
	if (m_ar == 0)
	{
		m_ar = new LabelArray;
		m_arcnt = 0;
	}

	if (m_arcnt >= LSIZE)
	{
		m_arList.push_back(m_ar);
		m_ar = new LabelArray;
		m_arcnt = 0;
	}
	(*m_ar)[m_arcnt++] = m;
}

MboxMail* MySimpleDeque::Get(int position)
{
	LabelArray *ar;
	int arpos;

	int arListCount = m_arList.size() *LSIZE;
	int arCount = arListCount + m_arcnt;
	if (position >= arCount)
		return 0; // error

	if (position < arListCount)
	{
		int index = position / LSIZE;
		ar = m_arList[index];
		arListCount = index * LSIZE;
		arpos = position - arListCount;
	}
	else
	{
		ar = m_ar;
		arpos = position - arListCount;
	}

	if (arpos >= LSIZE)
		int deb = 1;
		
	MboxMail *m = (*ar)[arpos];

	if (m == 0)
		int deb = 1;

	if (m->m_startOff)
		int deb = 1;

	return m;
}

int NTreeView::ShowGmailLabels(HTREEITEM hItem, CString &listFilePath, CString &dataFilePath)
{
	CString errorText;
	CFileFind finder;
	if (hItem == 0)
		return 1;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	NListView *pListView = 0;
	if (pFrame)
		pListView = pFrame->GetListView();

	if (pListView == 0)
	{
		MboxMail::assert_unexpected();
		return -1;
	}

	CString itemName = m_tree.GetItemText(hItem);

	// build a string with wildcards
	CString strWildcard(listFilePath);
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

		CString fPath = finder.GetFilePath();
		CString fName = finder.GetFileName();
		CString label;
		FileUtils::GetFileBaseName(fName, label);
		//TRACE("fPath=|%s| fName=|%s|\n", fPath, fName);

#if 0
		if (label.Compare("Osobiste") == 0)
		{
			if (finder.IsDirectory())
				int deb = 1;
			int deb = 1;
		}
#endif

#if 0
		// Make sure the directory  and file order doesn't matter
		if (finder.IsDirectory())
		{
			if (label.Compare("Osobiste") == 0)
			{
				CString filePath = finder.GetFilePath();
				CString fileName = finder.GetFileName();
				//TRACE("%s\n", (LPCTSTR)filePath);

				CString label;
				FileUtils::GetFileBaseName(fileName, label);

				HTREEITEM found_hItem = NTreeView::HasLabel(hItem, label);
				if (found_hItem == 0)
				{
					HTREEITEM newItem = m_tree.InsertItem(label, 6, 6, hItem, TVI_SORT);
					if (newItem == 0)
					{
						MboxMail::assert_unexpected();
						return -1;
					}
					int nId = m_labelInfoStore.GetNextId();
					m_tree.SetItemData(newItem, nId);

					CString listFilePath = filePath + ".mboxlist";
					LabelInfo *linfo = new LabelInfo(nId, dataFilePath, label, listFilePath);
					m_labelInfoStore.Add(linfo, nId);
				}
				else
				{
					DWORD nId = m_tree.GetItemData(found_hItem);
					LabelInfo *linfo = m_labelInfoStore.Find(nId);

					if (linfo->m_listFilePath.Compare(filePath) == 0) // should be true
						int deb = 1;
					if (linfo)
					{
						linfo->m_listFilePath = filePath;
					}
				}
			}
			int deb = 1;
		}
#endif

		// if it's a directory, recursively search it
		if (finder.IsDirectory())
		{
			CString folderPath = finder.GetFilePath();
			CString folderName = finder.GetFileName();
			CString label;
			FileUtils::GetFileBaseName(folderName, label);
			//TRACE("%s\n", (LPCTSTR)folderPath);

			HTREEITEM found_hItem = NTreeView::HasLabel(hItem, label);
			if (found_hItem == 0)
			{
				HTREEITEM newItem = m_tree.InsertItem(folderName, 6, 6, hItem, TVI_SORT);
				if (newItem == 0)
				{
					MboxMail::assert_unexpected();
					return -1;
				}

				int nId = m_labelInfoStore.GetNextId();
				m_tree.SetItemData(newItem, nId);

				CString emptyPath = folderPath + ".mboxlist";

				if (!FileUtils::PathFileExist(emptyPath))
				{
					MboxMail::s_mails_label.SetCountKeepData(0);
					_int64 mailFileSize = FileUtils::FileSize(dataFilePath);
					int r = pListView->WriteMboxListFile_v2(&MboxMail::s_mails_label, emptyPath, mailFileSize, errorText);
					if (r < 0) {
						MboxMail::assert_unexpected();
						return -1;
					}
				}

				LabelInfo *linfo = new LabelInfo(nId, dataFilePath, label, emptyPath);
				m_labelInfoStore.Add(linfo, nId);

				int r = NTreeView::ShowGmailLabels(newItem, folderPath, dataFilePath);
				if (r < 0) {
					MboxMail::assert_unexpected();
					return -1;
				}
			}
			else
			{
				int r = NTreeView::ShowGmailLabels(found_hItem, folderPath, dataFilePath);
				if (r < 0) {
					MboxMail::assert_unexpected();
					return -1;
				}
			}
		}
		else
		{
			CString filePath = finder.GetFilePath();
			CString fileName = finder.GetFileName();
			//TRACE("%s\n", (LPCTSTR)filePath);

			CString label;
			FileUtils::GetFileBaseName(fileName, label);

			HTREEITEM found_hItem = NTreeView::HasLabel(hItem, label);
			if (found_hItem == 0)
			{
				HTREEITEM newItem = m_tree.InsertItem(label, 6, 6, hItem, TVI_SORT);
				if (newItem == 0)
				{
					MboxMail::assert_unexpected();
					return -1;
				}
				int nId = m_labelInfoStore.GetNextId();
				m_tree.SetItemData(newItem, nId);

				LabelInfo *linfo = new LabelInfo(nId, dataFilePath, label, filePath);
				m_labelInfoStore.Add(linfo, nId);
			}
			else
			{
				DWORD nId = m_tree.GetItemData(found_hItem);
				LabelInfo *linfo = m_labelInfoStore.Find(nId);

				if (linfo->m_listFilePath.Compare(filePath) == 0) // should be true
					int deb = 1;
				if (linfo)
				{
					linfo->m_listFilePath = filePath;
				}
			}
		}
	}
	finder.Close();
	return 1;
}

HTREEITEM NTreeView::HasLabel(HTREEITEM hItem, CString &label)
{
	CString name;
	if (hItem != NULL)
	{
		HTREEITEM hChild = m_tree.GetChildItem(hItem);
		while (hChild)
		{
			name = m_tree.GetItemText(hChild);
			if (name.Compare(label) == 0)
				return hChild;
			hChild = m_tree.GetNextSiblingItem(hChild);
		}
	}
	return 0;
}

LabelInfo::LabelInfo(int nId, CString &mailFolderPath, CString &mailDataFolderPath)
{
	m_nodeType = MailFolder;

	m_nId = nId;
	m_mailFolderPath = mailFolderPath;
	m_mailDataFolderPath = mailDataFolderPath;
}

LabelInfo::LabelInfo(int nId, CString &mailFilePath)
{
	m_nodeType = MailFile;
	m_nId = nId;
	m_filePath = mailFilePath;
}

LabelInfo::LabelInfo(int nId, CString &mailFilePath, CString &label, CString &listFilePath)
{
	m_nodeType = MailLabel;
	m_nId = nId;
	m_label = label;
	m_listFilePath = listFilePath;
	m_filePath = mailFilePath;
}

LabelInfo::~LabelInfo() 
{
	;
}


GlobalLabelInfo::GlobalLabelInfo()
{
	m_nId = 100000;  // base id
	m_labelInfoHT = new GlobalLableInfoMapType(311);
}

GlobalLabelInfo::~GlobalLabelInfo()
{
	Clear();
	delete m_labelInfoHT;
}

LabelInfo* GlobalLabelInfo::Find(int key)
{
	LabelInfo *linfo = m_labelInfoHT->find(&key);
	return linfo;
}

int GlobalLabelInfo::GetNextId()
{
	m_nId++;
	if (m_nId == 0x7fffffff)  // that might a problem; try rand() ??
	{
		// ASSERT(0);
		m_nId = 100000;
	}

	return m_nId;
}

int GlobalLabelInfo::Add(LabelInfo *linfo, int key)
{
	DWORD_PTR n = 0;

	// No check. User must call Find() first
	m_labelInfoHT->insert(&key, linfo);
	return m_nId;
}

void GlobalLabelInfo::Clear()
{
	GlobalLableInfoMapType::IHashMapIter iter = m_labelInfoHT->first();
	int totalCnt = 0;
	for (; !m_labelInfoHT->last(iter); )
	{
		LabelInfo *l = iter.element;
		m_labelInfoHT->remove(iter);
		delete l;

	}
	m_labelInfoHT->clear();
}

int NTreeView::LoadLabels()
{
	CString sText;
	int paneId = 0;
	CString path;
	CString name;
	CString name2;
	DWORD nId;
	DWORD nId2;
	LabelInfo *linfo;
	LabelInfo *linfo2;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame) {
		sText.Format("Loading labels ...");
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);
	}

	HTREEITEM hItem = m_tree.GetRootItem();
	name = m_tree.GetItemText(hItem);

	hItem = 0;
	HTREEITEM hChild = m_tree.GetChildItem(hItem);
	while (hChild)
	{
		name = m_tree.GetItemText(hChild);
		nId = m_tree.GetItemData(hChild);

		linfo = m_labelInfoStore.Find(nId);

		HTREEITEM hChild2 = m_tree.GetChildItem(hChild);
		while (hChild2)
		{
			name2 = m_tree.GetItemText(hChild2);
			nId2 = m_tree.GetItemData(hChild2);

			linfo2 = m_labelInfoStore.Find(nId2);
			if (linfo2 == 0) {
				MboxMail::assert_unexpected();
				hChild2 = m_tree.GetNextSiblingItem(hChild2);
				continue;
			}

			CString labelsCachePath;
			CString rootPrintSubFolder = "LabelCache";
			CString errorText;
			CString targetPrintSubFolder;

			MboxMail::s_path = path;

			CString driveName;
			CString directory;
			CString fileNameBase;
			CString fileNameExtention;
			FileUtils::SplitFilePath(linfo2->m_filePath, driveName, directory, fileNameBase, fileNameExtention);

			CString lastPath = driveName + directory;

			MboxMail::SetLastPath(lastPath);

			BOOL retval = MboxMail::GetCachePath(rootPrintSubFolder, targetPrintSubFolder, labelsCachePath, errorText, &linfo2->m_filePath);
			if (retval == FALSE) {
				MboxMail::assert_unexpected();
				sText.Format("Ready"); if (pFrame) pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
				return -1;  // goto instead ?
			}

			if (!labelsCachePath.IsEmpty() && FileUtils::PathDirExists(labelsCachePath))
			{
				int ret = ShowGmailLabels(hChild2, labelsCachePath, linfo2->m_filePath);
				if (ret < 0) {
					MboxMail::assert_unexpected();
					sText.Format("Ready"); if (pFrame) pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
					return -1;  // goto instead ?
				}
				CString inbox("Inbox");
				CString sent("Sent");
				MoveLabelItem(linfo2->m_filePath, sent);
				MoveLabelItem(linfo2->m_filePath, inbox);
			}
			hChild2 = m_tree.GetNextSiblingItem(hChild2);
		}
		hChild = m_tree.GetNextSiblingItem(hChild);
	}
	sText.Format("Ready"); if (pFrame) pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
	return 1;
}

HTREEITEM NTreeView::HasMailFile(CString &mailFilePath)
{
	CString name;
	CString name2;
	DWORD nId;
	DWORD nId2;
	LabelInfo *linfo;
	LabelInfo *linfo2;

	HTREEITEM hItem = m_tree.GetRootItem();
	name = m_tree.GetItemText(hItem);

	hItem = 0;
	HTREEITEM hChild = m_tree.GetChildItem(hItem);
	while (hChild)
	{
		name = m_tree.GetItemText(hChild);
		nId = m_tree.GetItemData(hChild);

		linfo = m_labelInfoStore.Find(nId);
		HTREEITEM hChild2 = m_tree.GetChildItem(hChild);
		while (hChild2)
		{
			name2 = m_tree.GetItemText(hChild2);
			nId2 = m_tree.GetItemData(hChild2);

			linfo2 = m_labelInfoStore.Find(nId2);
			if (linfo2 == 0) {
				MboxMail::assert_unexpected();
				hChild2 = m_tree.GetNextSiblingItem(hChild2);
				continue;
			}

			if (linfo2->m_filePath.Compare(mailFilePath) == 0)
				return hChild2;
			hChild2 = m_tree.GetNextSiblingItem(hChild2);
		}
		hChild = m_tree.GetNextSiblingItem(hChild);
	}
	return 0;
}

void NTreeView::MoveLabelItem(CString &mailFilePath, CString &label)
{
	HTREEITEM hRoot = HasMailFile(mailFilePath);
	HTREEITEM hitemToBeMoved = HasLabel(hRoot, label);
	if (hitemToBeMoved == 0)
		return;
	CString name = m_tree.GetItemText(hitemToBeMoved);
	DWORD nId = m_tree.GetItemData(hitemToBeMoved);

	HTREEITEM parent = m_tree.GetParentItem(hitemToBeMoved);
	m_tree.DeleteItem(hitemToBeMoved);
	HTREEITEM hItem = m_tree.InsertItem(name, 6, 6, parent, TVI_FIRST);
	m_tree.SetItemData(hItem, nId);
}

void NTreeView::MoveItem(HTREEITEM hitemToBeMoved, HTREEITEM hitemInsertAfter)
{
#if 0
	TV_INSERTSTRUCT tvStruct;
	TCHAR cBuffer[50];
	HTREEITEM hNewItem;

	tvStruct.item.hItem = hitemToBeMoved;
	tvStruct.item.cchTextMax = 49;
	tvStruct.item.pszText = cBuffer;
	tvStruct.item.mask = TVIF_CHILDREN | TVIF_HANDLE | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_TEXT | TVIF_PARAM;
	GetItem(&tvStruct.item);
	tvStruct.hParent = GetParentItem(hitemInsertAfter);
	tvStruct.hInsertAfter = hitemInsertAfter;
	tvStruct.item.mask = TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_TEXT | TVIF_PARAM;

	hNewItem = InsertItem(&tvStruct);
	HTREEITEM hChild = GetChildItem(hitemToBeMoved);
	while (hChild != NULL)
	{
		TransferItem(hChild, hNewItem);
		DeleteItem(hChild);
		hChild = GetChildItem(hitemToBeMoved);
	}
	DeleteItem(hitemToBeMoved);
#endif
}

//
void NTreeView::OnGmaillabelsCreate()
{
	// TODO: Add your command handler code here

	CString name;
	CString name2;
	DWORD nId;
	DWORD nId2;
	LabelInfo *linfo;
	LabelInfo *linfo2;

	HTREEITEM hItem = m_tree.GetRootItem();
	name = m_tree.GetItemText(hItem);

	hItem = 0;
	HTREEITEM hChild = m_tree.GetChildItem(hItem);
	while (hChild)
	{
		name = m_tree.GetItemText(hChild);
		nId = m_tree.GetItemData(hChild);

		linfo = m_labelInfoStore.Find(nId);
		HTREEITEM hChild2 = m_tree.GetChildItem(hChild);
		while (hChild2)
		{
			name2 = m_tree.GetItemText(hChild2);
			nId2 = m_tree.GetItemData(hChild2);

			linfo2 = m_labelInfoStore.Find(nId2);
			if (linfo2 == 0) {
				MboxMail::assert_unexpected();
				hChild2 = m_tree.GetNextSiblingItem(hChild2);
				continue;
			}

			if (!linfo2->m_filePath.IsEmpty())
			{
				int ret = CreateGmailLabelFiles(hChild2);
				if (ret < 0)
				{
					MboxMail::assert_unexpected();
					return;
				}
			}

			hChild2 = m_tree.GetNextSiblingItem(hChild2);
		}
		hChild = m_tree.GetNextSiblingItem(hChild);
	}
	BOOL ret = LoadLabels();
	int deb = 1;
	return;
}


void NTreeView::OnGmaillabelsDelete()
{
	// TODO: Add your command handler code here
#if 1
	int paneId = 0;
	CString lastDataPath;

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	BOOL lastPathLoaded = FALSE;
	BOOL expand = FALSE;
	int ii = 0;
	for (ii = 0; ii < m_folderArray.m_array.GetCount(); ii++)
	{
		CString path = m_folderArray.m_array.GetAt(ii);
		if (!FileUtils::PathDirExists(path))
			continue;

		if (!path.IsEmpty())
		{
			lastDataPath = MboxMail::GetDataPath(path);
			CString labelsDir = lastDataPath + "LabelCache";
			BOOL retval = FileUtils::RemoveDir(labelsDir, true, true);

			int deb = 1;
		}
	}

#else

	DWORD errorCode;
	CString name;
	CString name2;
	DWORD nId;
	DWORD nId2;
	LabelInfo *linfo;
	LabelInfo *linfo2;

	HTREEITEM hItem = m_tree.GetRootItem();
	name = m_tree.GetItemText(hItem);

	hItem = 0;
	HTREEITEM hChild = m_tree.GetChildItem(hItem);
	while (hChild)
	{
		name = m_tree.GetItemText(hChild);
		nId = m_tree.GetItemData(hChild);

		linfo = m_labelInfoStore.Find(nId);

		if (linfo == 0) {
			MboxMail::assert_unexpected();
			hChild = m_tree.GetNextSiblingItem(hChild);
			continue;
		}

		CString labelsDir = linfo->m_mailDataFolderPath + "LabelCache";
		BOOL retval = FileUtils::RemoveDir(labelsDir, true, true);
		//BOOL retval = FileUtils::RemoveDirectory(labelsDir, errorCode);

		hChild = m_tree.GetNextSiblingItem(hChild);
	}
#endif
	return;
}


void NTreeView::OnGmaillabelsRefresh()
{
	// TODO: Add your command handler code here
	OnGmaillabelsDelete();
	OnGmaillabelsCreate();

	int deb = 1;
}

// It wil create lable files and update the Tree
// Label files should be remoced before calling thsi function
int NTreeView::CreateLabelsForSingleMailFile(HTREEITEM hItem)
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	int paneId = 0;
	CString sText;
	CString mailFile = m_tree.GetItemText(hItem);

	sText.Format("Creating Gmail Labels for %s ...", mailFile);
	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	RecreateGmailLabels(hItem);

	sText.Format("Ready");
	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, FALSE);

	int labelsEnabled = 1;
	BOOL ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "labelsEnabled", labelsEnabled);

	return 1;
}


// Deletes all label files
int NTreeView::DeleteLabelsForSingleMailFile(HTREEITEM hItem)
{
	if (!hItem)
		return -1;

	CString labelsCachePath;
	CString rootPrintSubFolder = "LabelCache";
	CString errorText;
	CString targetPrintSubFolder;

	CString name = m_tree.GetItemText(hItem);
	int nId = m_tree.GetItemData(hItem);

	LabelInfo *linfo = m_labelInfoStore.Find(nId);

	if (linfo)
	{
		BOOL retval = MboxMail::GetCachePath(rootPrintSubFolder, targetPrintSubFolder, labelsCachePath, errorText, &linfo->m_filePath);
		if (retval == FALSE) {
			MboxMail::assert_unexpected();
			return -1;
		}
	}
	else
	{
		BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, labelsCachePath, errorText);
		if (retval == FALSE) {
			MboxMail::assert_unexpected();
			return -1;
		}
	}

	CString lastDataPath = MboxMail::GetLastDataPath();
	if (FileUtils::PathDirExists(labelsCachePath))
	{
		// TODO: remove files, keep tree item
		if (hItem) {
			DeleteItemChildren(hItem);
		}
		BOOL retval = FileUtils::RemoveDir(labelsCachePath, true, true);
		int deb = 1;
	}
	return 1;
}

int NTreeView::RefreshLabelsForSingleMailFile(HTREEITEM hItem)
{
	int retD = DeleteLabelsForSingleMailFile(hItem);
	int retC = CreateLabelsForSingleMailFile(hItem);

	int labelsEnabled = 1;
	BOOL ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "labelsEnabled", labelsEnabled);
	
	return retC;
}

int NTreeView::UpdateLabelMailListFile(HTREEITEM hItem)
{
	CString errorText;
	CString mboxFileNamePath = MboxMail::s_path;
	CString labelListFileNamePath = MboxMail::s_path_label;
	CString lastPath = MboxMail::GetLastPath();
	CString lastDataPath = MboxMail::GetLastDataPath();

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	NListView *pListView = 0;
	if (pFrame)
		pListView = pFrame->GetListView();

	if (pListView == 0)
	{
		MboxMail::assert_unexpected();
		return -1;
	}
	if (!labelListFileNamePath.IsEmpty())
	{
		_int64 mailFileSize = MboxMail::s_fSize;
		int wret = pListView->WriteMboxListFile_v2(&MboxMail::s_mails, labelListFileNamePath, mailFileSize, errorText);
		if (wret < 0)
		{
			return -1;
		}
	}
	else
	{
		MboxMail::assert_unexpected();
		return -1;
	}
	return 1;
}

