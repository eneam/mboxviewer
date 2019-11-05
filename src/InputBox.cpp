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

// InputBox.cpp : implementation file
//

#include "stdafx.h"
#include "InputBox.h"
#include "afxdialogex.h"


// InputBox dialog

IMPLEMENT_DYNAMIC(InputBox, CDialogEx)

InputBox::InputBox(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_INPUT_BOX_DLG, pParent)
{

}

InputBox::~InputBox()
{
}

void InputBox::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_STRING, m_input);
}


BEGIN_MESSAGE_MAP(InputBox, CDialogEx)
END_MESSAGE_MAP()


// InputBox message handlers


BOOL InputBox::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here

	if (GetSafeHwnd()) {
		CWnd *p = GetDlgItem(IDYES);
		if (p) {
			GetDlgItem(IDC_STATIC)->SetWindowText(m_input);
		}
	}

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}
