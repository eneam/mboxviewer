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
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
//  Library General Public License for more details.
//
//  You should have received a copy of the GNU Library General Public
//  License along with this program; if not, write to the
//  Free Software Foundation, Inc., 51 Franklin St, Fifth Floor,
//  Boston, MA  02110 - 1301, USA.
//
//////////////////////////////////////////////////////////////////
//

#if !defined(AFX_NTREEVIEW_H__28F2C044_02DA_4048_8F8F_3D17ADEE1421__INCLUDED_)
#define AFX_NTREEVIEW_H__28F2C044_02DA_4048_8F8F_3D17ADEE1421__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// NTreeView.h : header file
//


#include <algorithm>
#include <unordered_map>
#include "WheelTreeCtrl.h"
//#include <afxtempl.h>

/////////////////////////////////////////////////////////////////////////////
// NTreeView window

typedef  CArray<CString> CSArray;

class CRegArray
{
public:
	CRegArray(CString &section);
	CRegArray();
	~CRegArray();

	int Find(CString &str);
	int Add(CString &str);
	int Delete(CString &str);
	int Delete(int index, CString &str);
	BOOL GetAt(int index, CString &str);
	int GetCount();
	BOOL SaveToRegistry();
	BOOL LoadFromRegistry();
	BOOL LoadFromRegistry(CSArray &ar);
	BOOL CreateKey(CString &section, HKEY &hKey);

	void Dump();

	CString m_ProcSoftwarePath;
	CString m_section;
	int m_nMaxSize;
	CSArray m_array;
};

class ArchiveFileInfo
{
public:
	ArchiveFileInfo() { fSize = 0; bShow = 1; };
	_int64 fSize;
	_int64 bShow;
};

typedef CMap<CString, LPCSTR, ArchiveFileInfo, ArchiveFileInfo> FileSizeMap;

class ArchiveFileInfoMap
{
public:
	ArchiveFileInfoMap(CString &folderPath) { m_folderPath = folderPath; };
	~ArchiveFileInfoMap() { m_fileSizes.RemoveAll();}
	CString m_folderPath;
	FileSizeMap m_fileSizes;
};

typedef unordered_map<std::string, ArchiveFileInfoMap*> GlobalFileSizeMap;


class NTreeView : public CWnd
{
	CFont		m_font;
	CImageList	m_il;
// Construction
public:
	NTreeView();
	virtual ~NTreeView();

	DECLARE_DYNCREATE(NTreeView)
	CWheelTreeCtrl	m_tree;
	//
	//  TODO: consider to create class
	FileSizeMap	*fileSizes; // plus other data
	FileSizeMap	m_fileSizes; // plus other data
	GlobalFileSizeMap m_gFileSizes;  // fileSizes of all folders
	//
	BOOL m_bIsDataDirty;  // if true, folder data/file sizes needs to be saved to .mboxview
	//
	BOOL m_bSelectMailFileDone;
	BOOL m_bSelectMailFilePostMsgDone;
	BOOL m_bGeneralHintPostMsgDone;
	int m_timerTickCnt;
	// timer
	UINT_PTR m_nIDEvent;
	UINT m_nElapse;
	CRegArray m_folderArray;
	int treeColWidth;

	BOOL  m_bInFillControl;
	int m_frameCx;
	int m_frameCy;

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(NTreeView)
	protected:
	virtual void PostNcDestroy();
	//}}AFX_VIRTUAL

// Implementation
public:
	void ClearGlobalFileSizeMap();
	BOOL RemoveFileSizeMap(CString path);
	BOOL SetupFileSizeMap(CString &path);
	void LoadFileSizes(CString &path, FileSizeMap &fileSizes, BOOL dontUpdateTree);
	int OpenHiddenFiles(HTREEITEM hItem, FileSizeMap &fileSizes);
	void RemoveFileFromTreeView(HTREEITEM hItem, FileSizeMap &fileSizes);
	void DetermineFolderPath(HTREEITEM hItem, CString &folderPath);
	HTREEITEM DetermineRootItem(HTREEITEM hItem);
	void Traverse( HTREEITEM hItem, CFile &fp, FileSizeMap &fileSizes);
	void SaveData(HTREEITEM hItem);
	void FillCtrl(BOOL expand = TRUE);
	void ExpandOrCollapseTree(BOOL expand);
	void LoadFolders();
	HTREEITEM HasFolder(CString &path);
	void SelectMailFile(CString *fileName = 0); // based on -MBOX= command line argument
	void InsertMailFile(CString &mailFile);
	void ForceParseMailFile(HTREEITEM hItem);
	void UpdateFileSizesTable(CString &path, _int64 fSize, FileSizeMap &fileSizes);

	int GetChildrenCount(HTREEITEM hItem, BOOL recursive);
	HTREEITEM FindItem(HTREEITEM hItem, CString &mailFileName);
	BOOL DeleteItem(HTREEITEM hItem);
	BOOL DeleteItemChildren(HTREEITEM hItem);
	BOOL DeleteFolder(HTREEITEM hItem);
	void StartTimer();
	void PostMsgCmdParamFileName(CString *fileName = 0);
	BOOL RefreshFolder(HTREEITEM hItem);

	// Folder related
	int CreateEmptyFolder(HTREEITEM hItem);
	int GetFolderPath(HTREEITEM hItem, CString &mboxName, CString &parentPath);
	int GetFolderPath(HTREEITEM hItem, CString &folderPath);

	int CreateFlatFolderList(HTREEITEM hItem, CArray<CString> &folderList);
	int CreateFlatFolderList(CString &mboxFileName, CArray<CString> &folderList);

	static void FindAllDirs(LPCTSTR pstr);

	// Generated message map functions
protected:
	//{{AFX_MSG(NTreeView)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg void OnUpdateFileRefresh(CCmdUI* pCmdUI);
	afx_msg void OnFileRefresh();
	afx_msg void OnUpdateTreeExpand(CCmdUI* pCmdUI);
	afx_msg void OnTreeExpand();
	//}}AFX_MSG
	afx_msg void OnSelchanged(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnSelchanging(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnRClick(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg LRESULT OnCmdParam_FileName(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmdParam_GeneralHint(WPARAM wParam, LPARAM lParam);
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnTimer(UINT_PTR nIDEvent);
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_NTREEVIEW_H__28F2C044_02DA_4048_8F8F_3D17ADEE1421__INCLUDED_)
