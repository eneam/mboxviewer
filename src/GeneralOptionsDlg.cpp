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


// GeneralOptionsDlg.cpp : implementation file
//

#include "StdAfx.h"
#include "afxdialogex.h"
#include "GeneralOptionsDlg.h"

// GeneralOptionsDlg dialog

IMPLEMENT_DYNAMIC(GeneralOptionsDlg, CDialogEx)

GeneralOptionsDlg::GeneralOptionsDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_GENERAL_OPTIONS_DLG, pParent)
	, m_relaxedMboxFileValidation(FALSE)
	, m_relativeInlineImageFilePath(FALSE)
{

}

GeneralOptionsDlg::~GeneralOptionsDlg()
{
}

void GeneralOptionsDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Check(pDX, IDC_RELAX_MBOX_FILE_VALIDATION, m_relaxedMboxFileValidation);
	DDX_Check(pDX, IDC_RELATIVE_ATTACHMENT_FILE_PATH, m_relativeInlineImageFilePath);
}


BEGIN_MESSAGE_MAP(GeneralOptionsDlg, CDialogEx)
	ON_BN_CLICKED(IDC_RELAX_MBOX_FILE_VALIDATION, &GeneralOptionsDlg::OnBnClickedRelaxMboxFileValidation)
	ON_BN_CLICKED(IDC_RELATIVE_ATTACHMENT_FILE_PATH, &GeneralOptionsDlg::OnBnClickedRelativeAttachmentFilePath)
	ON_BN_CLICKED(IDOK, &GeneralOptionsDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// GeneralOptionsDlg message handlers


void GeneralOptionsDlg::OnBnClickedRelaxMboxFileValidation()
{
	// TODO: Add your control notification handler code here
}


void GeneralOptionsDlg::OnBnClickedRelativeAttachmentFilePath()
{
	// TODO: Add your control notification handler code here
}


BOOL GeneralOptionsDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}


void GeneralOptionsDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here

	CDialogEx::OnOK();
}
