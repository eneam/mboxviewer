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

#if !defined(_OPEN_CONTAINING_FOLDER_DLG_)
#define _OPEN_CONTAINING_FOLDER_DLG_

#pragma once


// OpenContainingFolderDlg dialog

class OpenContainingFolderDlg : public CDialogEx
{
	DECLARE_DYNAMIC(OpenContainingFolderDlg)

public:
	OpenContainingFolderDlg(CString &text, CString &path, bool blockOpenFileOption=FALSE, CWnd* pParent=0);
	virtual ~OpenContainingFolderDlg();

	virtual INT_PTR DoModal();
	CWnd* m_pParent;

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_OPEN_FOLDER };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();

	CString m_text;
	CString m_path;
	BOOL m_blockOpenFileOption;
	//
	CToolTipCtrl m_toolTip;

	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnBnClickedYes();
	afx_msg BOOL OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult);
};

#endif // _OPEN_CONTAINING_FOLDER_DLG_
