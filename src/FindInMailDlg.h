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

#if !defined(_FIND_IN_MAIL_DLG_)
#define _FIND_IN_MAIL_DLG_


#pragma once


// CFindInMailDlg dialog

class CFindInMailDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CFindInMailDlg)

public:
	CFindInMailDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~CFindInMailDlg();

	virtual INT_PTR DoModal();
	CWnd* m_pParent;

// Dialog Data
	//{{AFX_DATA(CFindInMailDlg)
	enum { IDD = IDD_FIND_IN_MAIL };

	CString	m_string;
	BOOL	m_bWholeWord;
	BOOL	m_bCaseSensitive;
	//
	CToolTipCtrl m_toolTip;

	//}}AFX_DATA

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	virtual BOOL OnInitDialog();
	afx_msg BOOL OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult);
};

#endif // _FIND_IN_MAIL_DLG_
