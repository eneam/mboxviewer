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



#pragma once
#include "afxdialogex.h"


// PageCodeListDlg dialog


struct CodePageInfo {
	UINT codePage;
	CString codePageName;
	UINT maxCharSize;
};

class PageCodeListDlg : public CDialogEx
{
	DECLARE_DYNAMIC(PageCodeListDlg)

public:
	PageCodeListDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~PageCodeListDlg();

	virtual INT_PTR DoModal();
	CWnd* m_pParent;
	CFont m_font;

	void SortCodePageList(int column, BOOL descendingSort);
	void PrintListCtrl(HTREEITEM hItem, BOOL recursive, int depth);
	void ResetFont();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_PAGE_CODE_LIST_DLG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnItemactivateCodePageEnumeration(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnGetdispinfoCodePageEnumeration(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnItemclickCodePageEnumeration(NMHDR* pNMHDR, LRESULT* pResult);
	//
	virtual BOOL OnInitDialog();
	virtual void OnOK();
	virtual void OnCancel();

	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnPaint();

	static BOOL CALLBACK EnumCodePagesProc(LPWSTR  pCP);
	static PageCodeListDlg* pPageCodeListDlgInstance;  // Unfortunatelly EnumCodePagesProc doesn't accept  pointer to private data

	CListCtrl m_list;
	HICON m_hIcon;

	CArray<CodePageInfo> m_codePageList;
	BOOL m_descendingSort = TRUE;
	CString m_tmpstr;
	UINT m_CodePageSelected;

	void AddCodePage(UINT codePage);
	int CodePage2Index(UINT codePage);
	void ItemState2Str(UINT uState, CString& strState);

	afx_msg void OnClickCodePageEnumeration(NMHDR* pNMHDR, LRESULT* pResult);
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual BOOL OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult);
	afx_msg void OnTimer(UINT_PTR nIDEvent);
};
