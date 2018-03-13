// NTreeView.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "NTreeView.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
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

void NTreeView::FillCtrl()
{
	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	m_tree.DeleteAllItems();
	if (path.IsEmpty() || !PathFileExist(path))
		return;
	CString root;
	char *last_slash = (char*)strrchr(path, '\\');
	if (last_slash && *last_slash) {
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
						fileSizes.Lookup(fn, fSize);
						if (fSize < realFSize) {
							CString cache = fullPath + ".mboxview";
							DeleteFile(cache);
							m_tree.SetItemState(hItem, TVIS_BOLD, TVIS_BOLD);
						}
					}
				}
			} while (FindNextFile(f, &wf));
			FindClose(f);
		}
		m_tree.Expand(hRoot, TVE_EXPAND);
	}
	
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame) {
		NListView *pView = pFrame->GetListView();
		if (pView) {
			pView->m_path = "";
			pView->m_which = NULL;
			// pView->ResetSize();  // check if needed
			pView->FillCtrl();
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
	NListView *pView = pFrame->GetListView();
	if( ! pView )
		return;
	if( ! pNm->itemNew.hItem ) {
		pView->m_path = "";
		pView->m_which = NULL;
		pView->FillCtrl();
		return;
	}
	HTREEITEM hRoot = m_tree.GetRootItem();
	if( pNm->itemNew.hItem == hRoot ) {
		pView->m_path = "";
		pView->m_which = NULL;
		pView->ResetSize();
		pView->FillCtrl();
		return;
	}
	CString str = m_tree.GetItemText(pNm->itemNew.hItem);
	CString path = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath");
	if( str.IsEmpty() || path.IsEmpty() )
		return;
	pView->m_path = path + "\\" + str;
	pView->m_which = pNm->itemNew.hItem;
	pView->ResetSize();
	pView->FillCtrl();
}

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
	NListView *pView = pFrame->GetListView();
	if (!pView)
		return;

	pView->m_path = path + _T("\\") + str;

	CString txt;
	if (!PathFileExist(path))
		txt = _T("Nonexistent Directory \"") + path;
	else if (!PathFileExist(pView->m_path))
		txt = _T("Nonexistent File \"") + pView->m_path;

	if (!txt.IsEmpty()) {
		txt += _T("\".\nDo you want to continue?");
		int answer = MessageBox(txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
			AfxGetMainWnd()->PostMessage(WM_CLOSE);
	}
	else {
		pView->m_which = NULL;
		pView->ResetSize();
		pView->FillCtrl();
	}
}

BOOL NTreeView::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt) 
{
	CWnd *wnd = WindowFromPoint(pt);
	if( wnd == NULL )
		return FALSE;
	if( wnd == this ) {
		return CWnd::OnMouseWheel(nFlags, zDelta, pt);
	}
	return wnd->PostMessage(WM_MOUSEWHEEL, MAKELPARAM(nFlags,zDelta), MAKELPARAM(pt.x, pt.y));
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
			if( ! fileSizes.Lookup(path, fSize) )
				fileSizes[path] = fSize = FileSize(fullPath);
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
			if( fSize < FileSize(fullPath) ) {
				CString cache= fullPath + ".mboxview";
				DeleteFile(cache);
				m_tree.SetItemState(hItem, TVIS_BOLD, TVIS_BOLD);
				bFound = TRUE;
			}
		}
		hItem = m_tree.GetNextSiblingItem(hItem);
	}
	if( bFound )
		m_tree.SelectItem(NULL);
}
