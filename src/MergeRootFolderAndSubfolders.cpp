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


// MergeRootFolderAndSubfolders.cpp : implementation file
//

#include "StdAfx.h"
#include "MergeRootFolderAndSubfolders.h"
#include "afxdialogex.h"
#include "ResHelper.h"


// MergeRootFolderAndSubfolders dialog

IMPLEMENT_DYNAMIC(MergeRootFolderAndSubfolders, CDialogEx)

MergeRootFolderAndSubfolders::MergeRootFolderAndSubfolders(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_MERGE_FOLDER_AND_SUBFOLDERS, pParent)
{
	m_mergeRootFolderStyle = 1;  // merge all mbox files under root folder and all subfolders
	m_labelAssignmentStyle = 1;  // assign mbox file name as label to each mail

	// Get the log font.
	NONCLIENTMETRICS ncm;
	memset(&ncm, 0, sizeof(NONCLIENTMETRICS));
	ncm.cbSize = sizeof(NONCLIENTMETRICS);

	BOOL ver = ::SystemParametersInfo(SPI_GETNONCLIENTMETRICS,
		sizeof(NONCLIENTMETRICS), &ncm, 0);

	HDC hdc = ::GetWindowDC(NULL);
	ncm.lfMessageFont.lfWeight = 400;
	ncm.lfMessageFont.lfHeight = -MulDiv(12, GetDeviceCaps(hdc, LOGPIXELSY), 72);;
	m_TextFont.CreateFontIndirect(&ncm.lfMessageFont);
	::ReleaseDC(NULL, hdc);

	hdc = ::GetWindowDC(NULL);
	ncm.lfMessageFont.lfWeight = FW_BOLD;
	ncm.lfMessageFont.lfHeight = -MulDiv(12, GetDeviceCaps(hdc, LOGPIXELSY), 72);;
	m_BoldFont.CreateFontIndirect(&ncm.lfMessageFont);
	::ReleaseDC(NULL, hdc);

	m_ButtonBrush.CreateSolidBrush(RGB(255, 255, 0));

}

MergeRootFolderAndSubfolders::~MergeRootFolderAndSubfolders()
{
}

void MergeRootFolderAndSubfolders::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT_INTRO, m_introText);
	DDX_Radio(pDX, IDC_MERGE_ROOT_FOLDER, m_mergeRootFolderStyle);
	DDX_Radio(pDX, IDC_LABEL_NONE, m_labelAssignmentStyle);
}


BEGIN_MESSAGE_MAP(MergeRootFolderAndSubfolders, CDialogEx)
	ON_BN_CLICKED(IDC_LABEL_NONE, &MergeRootFolderAndSubfolders::OnBnClickedLabelPerMboxFile)
	ON_BN_CLICKED(IDOK, &MergeRootFolderAndSubfolders::OnBnClickedOk)
	ON_BN_CLICKED(IDC_MERGE_ROOT_FOLDER, &MergeRootFolderAndSubfolders::OnBnClickedConfigFolderStyle1)
	ON_BN_CLICKED(IDC_MERGE_ROOT_FOLDER_AND_SUBFOLDERS, &MergeRootFolderAndSubfolders::OnBnClickedMergeRootFolderAndSubfolders)
END_MESSAGE_MAP()


// MergeRootFolderAndSubfolders message handlers

BOOL MergeRootFolderAndSubfolders::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here

	m_introText.SetFont(&m_TextFont);
#if 0
	m_introText.SetWindowText(L"This dialog enables users to configure root mail folder and sub-folders for merging mbox files exported from the same mail account.\r\n\r\n"
		L"When merging thousands of mail files, typically eml files, make sure to select an option to assign the same label to all files within containing folder (or assign no label)."
		L" Folder name will be assigned as the label name."
		L" By default, MBox Viewer creates a separate label per each mail file."
		L" This would create thousands of labels under the Mail Tree and make mail viewing unmanageable.\r\n")
	;
#else
	m_introText.SetWindowText(L"This dialog enables users to configure root mail folder and sub-folders for merging mbox files exported from the same mail account.\r\n\r\n"
		L"Select \"Assign no Labels\" option when merging Gmail mail files.\r\n"
		L"Select \"Assign no Labels\" or \"Assign same label to all MBOX files per containing folder\" option when merging thousands of mail files.\r\n"
		L"\r\n")
		;
#endif

	// Set TABSTOPS to FALSE to stop highlighting text
	m_introText.EnableWindow(TRUE);

	HWND h = this->GetSafeHwnd();
	ResHelper::LoadDialogItemsInfo(h);

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}

HBRUSH MergeRootFolderAndSubfolders::OnCtlColor(CDC* pDC, CWnd *pWnd, UINT nCtlColor)
{
	// The below works. It is called for every object in the dialog. May not be the most efficient.

	HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);
#if 0
	int nID = pWnd->GetDlgCtrlID();
	if (nID == IDC_USER_SELECTED_FOLDER_PATH)
	{
		return (HBRUSH)m_folderPathBrush;
	}
	else
		return hbr;
#else
	return hbr;
#endif
}


void MergeRootFolderAndSubfolders::OnBnClickedLabelPerMboxFile()
{
	// TODO: Add your control notification handler code here
	int deb = 1;
}


void MergeRootFolderAndSubfolders::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here

	UpdateData(TRUE);

	CDialogEx::OnOK();
}


void MergeRootFolderAndSubfolders::OnBnClickedConfigFolderStyle1()
{
	// TODO: Add your control notification handler code here

	CWnd *p = GetDlgItem(IDC_LABEL_NONE);
	if (p)
	{
		((CButton*)p)->SetCheck(1);
		p->EnableWindow(TRUE);
	}
	p = GetDlgItem(IDC_LABEL_PER_MBOX_FILE);
	if (p)
	{
		((CButton*)p)->SetCheck(0);
		p->EnableWindow(TRUE);
	}
	p = GetDlgItem(IDC_LABEL_PER_FOLDER);
	if (p)
	{
		((CButton*)p)->SetCheck(0);
		p->EnableWindow(FALSE);
	}
}


void MergeRootFolderAndSubfolders::OnBnClickedMergeRootFolderAndSubfolders()
{
	// TODO: Add your control notification handler code here
	CWnd *p = GetDlgItem(IDC_LABEL_NONE);
	if (p)
	{
		((CButton*)p)->SetCheck(0);
		p->EnableWindow(TRUE);
	}
	p = GetDlgItem(IDC_LABEL_PER_MBOX_FILE);
	if (p)
	{
		((CButton*)p)->SetCheck(1);
		p->EnableWindow(TRUE);
	}
	p = GetDlgItem(IDC_LABEL_PER_FOLDER);
	if (p)
	{
		((CButton*)p)->SetCheck(0);
		p->EnableWindow(TRUE);
	}
}
