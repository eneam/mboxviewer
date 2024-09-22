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


// SelectLanguageDlg.cpp : implementation file
//

#include "stdafx.h"
#include "afxdialogex.h"
#include "SelectLanguageDlg.h"
#include "FileUtils.h"
#include "ResHelper.h"


// SelectLanguageDlg dialog

IMPLEMENT_DYNAMIC(SelectLanguageDlg, CDialogEx)

SelectLanguageDlg::SelectLanguageDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_LANG_LIST_DLG, pParent)
{

}

SelectLanguageDlg::~SelectLanguageDlg()
{
}

void SelectLanguageDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_SELECT_LANG, m_listBox);
}


BEGIN_MESSAGE_MAP(SelectLanguageDlg, CDialogEx)
	ON_BN_CLICKED(IDC_STATIC_SELECT_LANG_1, &SelectLanguageDlg::OnBnClickedStaticSelectLang1)
	ON_BN_CLICKED(IDOK, &SelectLanguageDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &SelectLanguageDlg::OnBnClickedCancel)
END_MESSAGE_MAP()


// SelectLanguageDlg message handlers


void SelectLanguageDlg::OnBnClickedStaticSelectLang1()
{
	// TODO: Add your control notification handler code here
	int deb = 1;
}


BOOL SelectLanguageDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here

	LoadLanguages();

	int lastIndex = m_listBox.FindStringExact(0, m_lastFolderName);
	int retIndex = m_listBox.SetCurSel(lastIndex);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

void SelectLanguageDlg::LoadLanguages()
{
	CFileFind finder;

	// build a string with wildcards
	CString folderPath;
	ResHelper::GetProcessFolderPath(folderPath);

	CString languagesPath = folderPath + L"Languages";

	CString strWildcard(languagesPath);
	strWildcard += L"\\*.*";

	// start working for files
	BOOL bWorking = finder.FindFile(strWildcard);

	m_listBox.AddString(L"english");
	while (bWorking)
	{
		bWorking = finder.FindNextFile();

		// skip . and .. files; otherwise, we'd
		// recur infinitely!

		if (finder.IsDots())
			continue;

		if (finder.IsDirectory())
		{
			CString str = finder.GetFilePath();
			CString FolderName;
			FileUtils::CPathStripPath((LPCWSTR)str, FolderName);
			TRACE(L"%s\n", (LPCWSTR)str);

			m_listBox.AddString(FolderName);
		}
	}
	finder.Close();
}


void SelectLanguageDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here

	int nSel = m_listBox.GetCurSel();

	if ((nSel >= 0) && (nSel < m_listBox.GetCount()))
	{
		CString text;
		m_listBox.GetText(nSel, text);

		m_LanguagetName = text;
	}

	CDialogEx::OnOK();
}


void SelectLanguageDlg::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here
	CDialogEx::OnCancel();
}
