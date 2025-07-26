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
#include "MainFrm.h"


// MergeRootFolderAndSubfolders dialog

IMPLEMENT_DYNAMIC(MergeRootFolderAndSubfolders, CDialogEx)

MergeRootFolderAndSubfolders::MergeRootFolderAndSubfolders(CWnd* pParent /*=nullptr*/)
DIALOG_FROM_TEMPLATE( : CDialogEx(IDD_MERGE_FOLDER_AND_SUBFOLDERS, pParent))
{
	m_mergeRootFolderStyle = 1;  // merge all mbox files under root folder and all subfolders
	m_labelAssignmentStyle = 1;  // assign mbox file name as label to each mail

	m_fontSize = 12;

	m_pParent = pParent;
}

MergeRootFolderAndSubfolders::~MergeRootFolderAndSubfolders()
{
}

INT_PTR MergeRootFolderAndSubfolders::DoModal()
{
#ifdef _DIALOG_FROM_TEMPLATE
	INT_PTR ret = CMainFrame::SetTemplate(this, IDD_MERGE_FOLDER_AND_SUBFOLDERS, m_pParent);
#else
	INT_PTR ret = CDialogEx::DoModal();
#endif
	return ret;
}

void MergeRootFolderAndSubfolders::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_MERGE_FOLDERS_INTRO_1, m_introText);
	DDX_Radio(pDX, IDC_MERGE_ROOT_FOLDER, m_mergeRootFolderStyle);
	DDX_Radio(pDX, IDC_LABEL_NONE, m_labelAssignmentStyle);
}


BEGIN_MESSAGE_MAP(MergeRootFolderAndSubfolders, CDialogEx)
	ON_BN_CLICKED(IDC_LABEL_NONE, &MergeRootFolderAndSubfolders::OnBnClickedLabelPerMboxFile)
	ON_BN_CLICKED(IDOK, &MergeRootFolderAndSubfolders::OnBnClickedOk)
	ON_BN_CLICKED(IDC_MERGE_ROOT_FOLDER, &MergeRootFolderAndSubfolders::OnBnClickedConfigFolderStyle1)
	ON_BN_CLICKED(IDC_MERGE_ROOT_FOLDER_AND_SUBFOLDERS, &MergeRootFolderAndSubfolders::OnBnClickedMergeRootFolderAndSubfolders)
	ON_NOTIFY_EX(TTN_NEEDTEXT, 0, &MergeRootFolderAndSubfolders::OnTtnNeedText)
END_MESSAGE_MAP()


// MergeRootFolderAndSubfolders message handlers

BOOL MergeRootFolderAndSubfolders::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here

	m_fontSize += 2;
	if (m_fontSize < 12)
		m_fontSize = 12;


		// Get the log font.
	NONCLIENTMETRICS ncm;
	memset(&ncm, 0, sizeof(NONCLIENTMETRICS));
	ncm.cbSize = sizeof(NONCLIENTMETRICS);

	BOOL ver = ::SystemParametersInfo(SPI_GETNONCLIENTMETRICS,
		sizeof(NONCLIENTMETRICS), &ncm, 0);

	HDC hdc = ::GetWindowDC(NULL);
	ncm.lfMessageFont.lfWeight = 400;
	ncm.lfMessageFont.lfHeight = -MulDiv(m_fontSize, GetDeviceCaps(hdc, LOGPIXELSY), 72);;
	m_TextFont.CreateFontIndirect(&ncm.lfMessageFont);
	::ReleaseDC(NULL, hdc);

	hdc = ::GetWindowDC(NULL);
	ncm.lfMessageFont.lfWeight = FW_BOLD;
	ncm.lfMessageFont.lfHeight = -MulDiv(m_fontSize, GetDeviceCaps(hdc, LOGPIXELSY), 72);;
	m_BoldFont.CreateFontIndirect(&ncm.lfMessageFont);
	::ReleaseDC(NULL, hdc);

	m_ButtonBrush.CreateSolidBrush(RGB(255, 255, 0));

	m_introText.SetFont(&m_TextFont);
#if 0
	CString introText(L"This dialog enables users to configure root mail folder and sub-folders for merging mbox files exported from the same mail account.\r\n\r\n"
		L"When merging thousands of mail files, typically eml files, make sure to select an option to assign the same label to all files within containing folder (or assign no label)."
		L" Folder name will be assigned as the label name."
		L" By default, MBox Viewer creates a separate label per each mail file."
		L" This would create thousands of labels under the Mail Tree and make mail viewing unmanageable.\r\n")
	;
#else
	CString introText(L"This dialog enables users to configure root mail folder and sub-folders for merging mbox files exported from the same mail account. "
		"Find functions will search all mails.\n\n"
		L"Select \"Assign no Labels\" option when merging Gmail mail files. "
		L"Select \"Assign no Labels\" or \"Assign same label to all MBOX files per containing folder\" option when merging thousands of mail files.")
		;
#endif

	m_introText.SetWindowText(introText);

	// Set TABSTOPS to FALSE to stop highlighting text
	m_introText.EnableWindow(TRUE);

	ResHelper::LoadDialogItemsInfo(this);
	ResHelper::UpdateDialogItemsInfo(this);
	BOOL retA = ResHelper::ActivateToolTips(this, m_toolTip);

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}

HBRUSH MergeRootFolderAndSubfolders::OnCtlColor(CDC* pDC, CWnd *pWnd, UINT nCtlColor)
{
	// The below works. It is called for every object in the dialog. May not be the most efficient.

	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);
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

BOOL MergeRootFolderAndSubfolders::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}

