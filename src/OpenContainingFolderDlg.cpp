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

// OpenContainingFolderDlg.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "OpenContainingFolderDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

// OpenContainingFolderDlg dialog

IMPLEMENT_DYNAMIC(OpenContainingFolderDlg, CDialogEx)

OpenContainingFolderDlg::OpenContainingFolderDlg(CString &text, bool blockOpenFileOption, CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_OPEN_FOLDER, pParent)
{
	m_text = text;
	m_blockOpenFileOption = blockOpenFileOption;
}

OpenContainingFolderDlg::~OpenContainingFolderDlg()
{
}

void OpenContainingFolderDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(OpenContainingFolderDlg, CDialogEx)
ON_BN_CLICKED(IDYES, &OpenContainingFolderDlg::OnBnClickedYes)
END_MESSAGE_MAP()


// OpenContainingFolderDlg message handlers

BOOL OpenContainingFolderDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here

	if ((m_blockOpenFileOption == TRUE) && GetSafeHwnd()) {
		CWnd *p = GetDlgItem(IDYES);
		if (p) {
			p->EnableWindow(FALSE);
		}
	}
	UpdateData(TRUE);

	//SetDlgItemText(IDC_STATIC, "Desired Text String")

	GetDlgItem(IDC_STATIC)->SetWindowText(m_text);

	;// SetWindowText(caption);

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}

void OpenContainingFolderDlg::OnBnClickedYes()
{
	// TODO: Add your control notification handler code here
	EndDialog(IDYES);
}
