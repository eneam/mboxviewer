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

// FindFilterRuleDlg.cpp : implementation file
//

#include "StdAfx.h"
#include "FindFilterRuleDlg.h"
#include "afxdialogex.h"
#include "ResHelper.h"
#include "MainFrm.h"


// FindFilterRuleDlg dialog

IMPLEMENT_DYNAMIC(FindFilterRuleDlg, CDialogEx)

FindFilterRuleDlg::FindFilterRuleDlg(CWnd* pParent /*=nullptr*/)
DIALOG_FROM_TEMPLATE( : CDialogEx(IDD_FIND_FILTER_DLG, pParent))
{
	m_filterNumb = 0;
	m_pParent = pParent;
}

FindFilterRuleDlg::~FindFilterRuleDlg()
{
}

INT_PTR FindFilterRuleDlg::DoModal()
{
#ifdef _DIALOG_FROM_TEMPLATE
	INT_PTR ret = CMainFrame::SetTemplate(this, IDD_FIND_FILTER_DLG, m_pParent);
#else
	INT_PTR ret = CDialogEx::DoModal();
#endif
	return ret;
}

void FindFilterRuleDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);

	DDX_Radio(pDX, IDC_FILTER_NUMBER0, m_filterNumb);
}


BEGIN_MESSAGE_MAP(FindFilterRuleDlg, CDialogEx)
	ON_NOTIFY_EX(TTN_NEEDTEXT, 0, &FindFilterRuleDlg::OnTtnNeedText)
END_MESSAGE_MAP()


// FindFilterRuleDlg message handlers


BOOL FindFilterRuleDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here
	BOOL retA = ResHelper::ActivateToolTips(this, m_toolTip);
	ResHelper::LoadDialogItemsInfo(this);
	ResHelper::UpdateDialogItemsInfo(this);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

BOOL FindFilterRuleDlg::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}
