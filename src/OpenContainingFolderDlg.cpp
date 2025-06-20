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
#include "ResHelper.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

// OpenContainingFolderDlg dialog

IMPLEMENT_DYNAMIC(OpenContainingFolderDlg, CDialogEx)

OpenContainingFolderDlg::OpenContainingFolderDlg(CString &text, CString &path, bool blockOpenFileOption, CWnd* pParent /*=nullptr*/)
DIALOG_FROM_TEMPLATE(: CDialogEx(IDD_OPEN_FOLDER, pParent))
{
	m_text = text;
	m_path = path;
	m_blockOpenFileOption = blockOpenFileOption;
	m_pParent = pParent;
}

OpenContainingFolderDlg::~OpenContainingFolderDlg()
{
}

INT_PTR OpenContainingFolderDlg::DoModal()
{
#ifdef _DIALOG_FROM_TEMPLATE
	INT_PTR ret = CMainFrame::SetTemplate(this, IDD_OPEN_FOLDER, m_pParent);
#else
	INT_PTR ret = CDialogEx::DoModal();
#endif
	return ret;
}

void OpenContainingFolderDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_STATIC_OPEN_FILE_OR_FOLDER, m_text);
	DDX_Text(pDX, IDC_EDIT_LOCATION_PATH, m_path);
}


BEGIN_MESSAGE_MAP(OpenContainingFolderDlg, CDialogEx)
	ON_BN_CLICKED(IDYES, &OpenContainingFolderDlg::OnBnClickedYes)
	ON_NOTIFY_EX(TTN_NEEDTEXT, 0, &OpenContainingFolderDlg::OnTtnNeedText)
END_MESSAGE_MAP()


// OpenContainingFolderDlg message handlers

int OpenContainingFolderDlg::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	//SS_EDITCONTROL
	return 0;
}

BOOL OpenContainingFolderDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here

	if (GetSafeHwnd())
	{
		CWnd* p = GetDlgItem(IDC_EDIT_LOCATION_PATH);
		if (p) {
			if (m_path.IsEmpty())
			{
				p->EnableWindow(FALSE);
				p->ShowWindow(SW_HIDE);
			}
			else
				p->EnableWindow(TRUE);
		}
	}

	if ((m_blockOpenFileOption == TRUE) && GetSafeHwnd())
	{
		CWnd *p = GetDlgItem(IDYES);
		if (p) {
			p->EnableWindow(FALSE);
			p->ShowWindow(SW_HIDE);
		}
	}
	//UpdateData(TRUE);

	//SetDlgItemText(IDC_STATIC, "Desired Text String")

	ResHelper::LoadDialogItemsInfo(this);
	ResHelper::UpdateDialogItemsInfo(this);
	BOOL retA = ResHelper::ActivateToolTips(this, m_toolTip);

	//GetDlgItem(IDC_STATIC)->SetWindowText(m_text);

	UpdateData(TRUE);

	;// SetWindowText(caption);

	//ResHelper::LoadDialogItemsInfo(this);

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}

void OpenContainingFolderDlg::OnBnClickedYes()
{
	// TODO: Add your control notification handler code here
	EndDialog(IDYES);
}

BOOL OpenContainingFolderDlg::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}
