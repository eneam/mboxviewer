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


// DevelopmentCreateArchive.cpp : implementation file
//

#include "StdAfx.h"
#include "DevelopmentCreateArchive.h"
#include "afxdialogex.h"
#include "FileUtils.h"
#include "ResHelper.h"


// DevelopmentCreateArchive dialog

IMPLEMENT_DYNAMIC(DevelopmentCreateArchive, CDialogEx)

DevelopmentCreateArchive::DevelopmentCreateArchive(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_DEV_CREATE_ARCHIVE, pParent)
{
	m_leadingMailCount = 2;
	m_trailingMailCount = 2;

}

DevelopmentCreateArchive::~DevelopmentCreateArchive()
{
}

BOOL DevelopmentCreateArchive::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	m_leadingMailCount = 2;
	m_trailingMailCount = 2;

	CString fileNameExtention;
	CString fileBaseName;
	CString fileName;

	FileUtils::GetFolderPathAndFileName(m_sourceArchiveFile, m_sourceArchiveFolder, fileName);
	FileUtils::GetFileBaseNameAndExtension(fileName, fileBaseName, fileNameExtention);

	m_createdArchiveName = fileBaseName + L"_subset" + fileNameExtention;

	UpdateData(FALSE);

	HWND h = this->GetSafeHwnd();
	ResHelper::LoadDialogItemsInfo(h);

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}

void DevelopmentCreateArchive::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_MAIL_LEADING_CNT, m_leadingMailCount);
	DDX_Text(pDX, IDC_MAIL_TRAILING_CNT, m_trailingMailCount);
	DDX_Text(pDX, IDC_SOURCE_ARCHIVE, m_sourceArchiveFile);
	DDX_Text(pDX, IDC_SELECTED_SUBJECT, m_selectedMailSubject);
	DDX_Text(pDX, IDC_CREATED_ARCHIVE, m_createdArchiveName);
}

BEGIN_MESSAGE_MAP(DevelopmentCreateArchive, CDialogEx)
	ON_BN_CLICKED(IDOK, &DevelopmentCreateArchive::OnBnClickedOk)
END_MESSAGE_MAP()


// DevelopmentCreateArchive message handlers


void DevelopmentCreateArchive::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here

	// Uncomment if you plan to check data
	UpdateData(TRUE);

	CString mboxFilePath = m_sourceArchiveFolder + m_createdArchiveName;

	if (FileUtils::PathFileExist(mboxFilePath))
	{
		CString txt;
		txt.Format(L"File \"%s\" already exist. Do you want to override?", mboxFilePath);
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer != IDYES)
			return;

		//FileUtils::DelFile(m_createdArchiveName);
	}

	CDialogEx::OnOK();
}
