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

// GenericFontFamilyDlg.cpp : implementation file
//

#include "stdafx.h"
#include "GenericFontFamilyDlg.h"
#include "afxdialogex.h"


// GenericFontFamilyDlg dialog

IMPLEMENT_DYNAMIC(GenericFontFamilyDlg, CDialogEx)

GenericFontFamilyDlg::GenericFontFamilyDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_GENERIC_FONT_DLG, pParent)
{
	m_genericFontName = "serif";

}

GenericFontFamilyDlg::~GenericFontFamilyDlg()
{
}

void GenericFontFamilyDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_listBox);
}


BEGIN_MESSAGE_MAP(GenericFontFamilyDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &GenericFontFamilyDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// GenericFontFamilyDlg message handlers

void GenericFontFamilyDlg::LoadData()
{
	m_listBox.AddString(_T("serif"));
	m_listBox.AddString(_T("sans-serif"));
	m_listBox.AddString(_T("monospace"));
	m_listBox.AddString(_T("cursive"));
	m_listBox.AddString(_T("fantasy"));
}

BOOL GenericFontFamilyDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here


	m_listBox.SetFont(GetFont());

	LoadData();

	//m_genericFontName = "none";
	//m_listBox.SetCurSel(0);

	CString fontFamilyName;

	for (int i = 0; i < m_listBox.GetCount(); i++)
	{
		m_listBox.GetText(i, fontFamilyName);
		if (fontFamilyName.CompareNoCase(m_genericFontName) == 0)
		{
			m_listBox.SetCurSel(i);
			break;
		}
	}

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}


void GenericFontFamilyDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here

	int nSel = m_listBox.GetCurSel();
#if 0
	if ((nSel < 0) || (nSel >= ColorStyleConfig::MailMaxPanes) || (nSel >= m_listBox.GetCount()))
	{
		CSCAssert();
		return;
	}
#endif

	m_listBox.GetText(nSel, m_genericFontName);

	int deb = 1;

	CDialogEx::OnOK();
}
