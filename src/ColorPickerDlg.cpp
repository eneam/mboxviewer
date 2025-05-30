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

// ColorPickerDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ColorPickerDlg.h"
#include "ResHelper.h"
#include "MainFrm.h"

#ifndef _WIN32_WCE // CColorDialog is not supported for Windows CE.

// ColorPickerDlg

IMPLEMENT_DYNAMIC(ColorPickerDlg, CColorDialog)

ColorPickerDlg::ColorPickerDlg(COLORREF clrInit, DWORD dwFlags, CWnd* pParentWnd)
// DIALOG_FROM_TEMPLATE(: CColorDialog(clrInit, dwFlags, pParentWnd))  // doesn't work
	: CColorDialog(clrInit, dwFlags, pParentWnd)
{
	m_pParent = pParentWnd;
}

ColorPickerDlg::~ColorPickerDlg()
{
}

// Doesn't seem to work
INT_PTR ColorPickerDlg::DoModal()
{
#if 0
	// This will not work. Can't init from template. Not sure what is the solution to increase font size and template layout
	UINT Idd = (UINT)m_lpszTemplateName;
	LOAD_TEMPLATE_FOR_DIALOG(INT_PTR ret = CMainFrame::SetTemplate(this, Idd, m_pParent));
#else
	INT_PTR ret = CColorDialog::DoModal();
#endif
	return ret;
}


BEGIN_MESSAGE_MAP(ColorPickerDlg, CColorDialog)
	ON_NOTIFY_EX(TTN_NEEDTEXT, 0, &ColorPickerDlg::OnTtnNeedText)
END_MESSAGE_MAP()



// ColorPickerDlg message handlers


#endif // !_WIN32_WCE


BOOL ColorPickerDlg::OnInitDialog()
{
	CColorDialog::OnInitDialog();

	// TODO:  Add extra initialization here

	CWnd *pwnd = GetParent();
	CRect cr;
	GetClientRect(cr);
	CRect pr;
	pwnd->GetWindowRect(pr);
	int xpos = pr.left - cr.Width() - 16;
	if (xpos < 0)
		xpos = 0;
	SetWindowPos(pwnd, pr.left, pr.bottom, 0, 0, SWP_SHOWWINDOW | SWP_NOSIZE);

	ResHelper::LoadDialogItemsInfo(this);
	ResHelper::UpdateDialogItemsInfo(this);
	BOOL retA = ResHelper::ActivateToolTips(this, m_toolTip);

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}


void ColorPickerDlg::OnOK()
{
	// TODO: Add your specialized code here and/or call the base class



	CColorDialog::OnOK();
}


BOOL ColorPickerDlg::OnColorOK()
{
	// TODO: Add your specialized code here and/or call the base class

	//return TRUE;

	return CColorDialog::OnColorOK();
}

BOOL ColorPickerDlg::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}
