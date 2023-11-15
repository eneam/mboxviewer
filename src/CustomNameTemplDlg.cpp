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

// CustomNameTemplDld.cpp : implementation file
//

#include "StdAfx.h"
#include "CustomNameTemplDlg.h"
#include "afxdialogex.h"

#include "mainFrm.h"
#include "MboxMail.h"

// CustomNameTemplDld dialog

IMPLEMENT_DYNAMIC(CustomNameTemplDlg, CDialogEx)

CustomNameTemplDlg::CustomNameTemplDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_NAME_TEMPL_DLG, pParent)
{
	m_nameTemplateCnf.SetDflts();
}

CustomNameTemplDlg::~CustomNameTemplDlg()
{
}

void CustomNameTemplDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);

	DDX_Text(pDX, IDC_EDIT_TEMPL_FMT, m_nameTemplateCnf.m_TemplateFormat);
	DDX_Text(pDX, IDC_EDIT_DATE_FMT, m_nameTemplateCnf.m_DateFormat);
	DDX_Check(pDX, IDC_TEMPL_FROM_USERNAME, m_nameTemplateCnf.m_bFromUsername);
	DDX_Check(pDX, IDC_TEMPL_FROM_DOMAIN_NAME, m_nameTemplateCnf.m_bFromDomain);
	DDX_Check(pDX, IDC_TEMPL_WHITE_TO_UNDERSCORE, m_nameTemplateCnf.m_bReplaceWhiteWithUnderscore);
	DDX_Check(pDX, IDC_TEMPL_TO_USERNAME, m_nameTemplateCnf.m_bToUsername);
	DDX_Check(pDX, IDC_TEMPL_TO_DOMAIN_NAME, m_nameTemplateCnf.m_bToDomain);
}


BEGIN_MESSAGE_MAP(CustomNameTemplDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &CustomNameTemplDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDC_SRCFTIME, &CustomNameTemplDlg::OnBnClickedSrcftime)
END_MESSAGE_MAP()


// CustomNameTemplDlg message handlers

void CustomNameTemplDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here

	UpdateData(TRUE);

	CArray<CString> labelArray;
	BOOL ret = MboxMail::ParseTemplateFormat(m_nameTemplateCnf.m_TemplateFormat, labelArray);
	if (labelArray.GetCount() == 0)
	{
		AfxMessageBox(L"Template format can't be empty!", MB_OK | MB_ICONHAND);
		return;
	}

	if ((m_nameTemplateCnf.m_bFromUsername == FALSE) && (m_nameTemplateCnf.m_bFromDomain == FALSE))
	{
		AfxMessageBox(L"Either From User Account Name or From Domain must be checked!", MB_OK | MB_ICONHAND);
		return;
	}

	if ((m_nameTemplateCnf.m_bToUsername == FALSE) && (m_nameTemplateCnf.m_bToDomain == FALSE))
	{
		AfxMessageBox(L"Either To User Account Name or To Domain must be checked!", MB_OK | MB_ICONHAND);
		return;
	}

	CDialogEx::OnOK();
}

void NameTemplateCnf::ClearParts()
{
	m_bReplaceWhiteWithUnderscore = FALSE;
	m_bFromUsername = FALSE;
	m_bFromDomain = FALSE;
	m_bToUsername = FALSE;
	m_bToDomain = FALSE;

	int deb = 1;
}


void NameTemplateCnf::SetDflts()
{

	m_TemplateFormat = L"%DATE_TIME% %UNIQUE_ID% %FROM_NAME% %TO_NAME% %SUBJECT%";
	m_DateFormat = L"%Y-%m-%d";

	m_bReplaceWhiteWithUnderscore = FALSE;
	m_bFromUsername = TRUE;
	m_bFromDomain = TRUE;
	m_bToUsername = TRUE;
	m_bToDomain = TRUE;

	int deb = 1;
}

void NameTemplateCnf::Copy(NameTemplateCnf &src)
{
	if (this == &src)
		return;

	m_TemplateFormat = src.m_TemplateFormat;
	m_DateFormat = src.m_DateFormat;

	m_bReplaceWhiteWithUnderscore = src.m_bReplaceWhiteWithUnderscore;
	m_bFromUsername = src.m_bFromUsername;
	m_bFromDomain = src.m_bFromDomain;
	m_bToUsername = src.m_bToUsername;
	m_bToDomain = src.m_bToDomain;
}

void CustomNameTemplDlg::OnBnClickedSrcftime()
{
	// TODO: Add your control notification handler code here

	HINSTANCE result = ShellExecute(NULL, L"open", 
		L"https://docs.microsoft.com/en-us/cpp/c-runtime-library/reference/strftime-wcsftime-strftime-l-wcsftime-l?view=vs-2019", NULL, NULL, SW_SHOWNORMAL);
	int deb = 1;
}
