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

// FindInMailDlg.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "TextUtilsEx.h"
#include "FindInMailDlg.h"
#include "afxdialogex.h"
#include "ResHelper.h"


// CFindInMailDlg dialog

IMPLEMENT_DYNAMIC(CFindInMailDlg, CDialogEx)

CFindInMailDlg::CFindInMailDlg(CWnd* pParent /*=nullptr*/)
DIALOG_FROM_TEMPLATE( : CDialogEx(IDD_FIND_IN_MAIL, pParent))
{
	m_pParent = pParent;
}

CFindInMailDlg::~CFindInMailDlg()
{
}

INT_PTR CFindInMailDlg::DoModal()
{
#ifdef _DIALOG_FROM_TEMPLATE
	INT_PTR ret = CMainFrame::SetTemplate(this, IDD_FIND_IN_MAIL, m_pParent);
#else
	INT_PTR ret = CDialogEx::DoModal();
#endif
	return ret;
}

void CFindInMailDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CFindDlg)
	DDX_Text(pDX, IDC_FIND_IN_MAIL_STRING, m_string);
	DDX_Check(pDX, IDC_WHOLE, m_bWholeWord);
	DDX_Check(pDX, IDC_CASE, m_bCaseSensitive);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CFindInMailDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &CFindInMailDlg::OnBnClickedOk)
	ON_NOTIFY_EX(TTN_NEEDTEXT, 0, &CFindInMailDlg::OnTtnNeedText)
END_MESSAGE_MAP()


// CFindInMailDlg message handlers


void CFindInMailDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	UpdateData();

	m_string.TrimRight();  // FIXME should we highlight spaces or remove ??
	if (m_string.IsEmpty()) {
		// If empty, clear the highlights
		; // AfxMessageBox("Cannot search for an empty string!", MB_OK | MB_ICONHAND);
		//return;
	}

	CDialogEx::OnOK();
}


BOOL CFindInMailDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here

	ResHelper::LoadDialogItemsInfo(this);
	ResHelper::UpdateDialogItemsInfo(this);
	BOOL retA = ResHelper::ActivateToolTips(this, m_toolTip);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}


BOOL CFindInMailDlg::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}

