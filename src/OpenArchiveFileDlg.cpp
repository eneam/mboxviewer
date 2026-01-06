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

// OpenArchiveFileDlg.cpp : implementation file
//

#include "stdafx.h"
#include "OpenArchiveFileDlg.h"
#include "FileUtils.h"
#include "afxdialogex.h"
#include "ResHelper.h"
#include "MainFrm.h"


// OpenArchiveFileDlg dialog

IMPLEMENT_DYNAMIC(OpenArchiveFileDlg, CDialogEx)

OpenArchiveFileDlg::OpenArchiveFileDlg(CWnd* pParent /*=nullptr*/)
DIALOG_FROM_TEMPLATE( : CDialogEx(IDD_OPEN_ARCHIVE_DLG, pParent))
{
	m_pParent = pParent;
}

OpenArchiveFileDlg::~OpenArchiveFileDlg()
{
}

INT_PTR OpenArchiveFileDlg::DoModal()
{
#ifdef _DIALOG_FROM_TEMPLATE
	INT_PTR ret = CMainFrame::SetTemplate(this, IDD_OPEN_ARCHIVE_DLG, m_pParent);
#else
	INT_PTR ret = CDialogEx::DoModal();
#endif
	return ret;
}

void OpenArchiveFileDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_ARCHIVE_FILE_NAME, m_archiveFileName);
}


BEGIN_MESSAGE_MAP(OpenArchiveFileDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &OpenArchiveFileDlg::OnBnClickedOk)
	ON_NOTIFY_EX(TTN_NEEDTEXT, 0, &OpenArchiveFileDlg::OnTtnNeedText)
END_MESSAGE_MAP()


// OpenArchiveFileDlg message handlers

BOOL OpenArchiveFileDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here

	if (GetSafeHwnd())
	{
		CWnd *p = GetDlgItem(IDC_STATIC);
#if 0
		if (p)
		{
			CString RuleNumberText;

			RuleNumberText.Format("The created archive file will be moved to the folder housing root archive");

			p->SetWindowText(RuleNumberText);
			p->EnableWindow(TRUE);
		}
#endif
		p = GetDlgItem(IDC_EDIT_DFLT_FOLDER);
		if (p)
		{
			p->SetWindowText(m_sourceFolder);
			p->EnableWindow(FALSE);
		}
		p = GetDlgItem(IDC_EDIT_TARGET_FOLDER);
		if (p)
		{
			CString targetFolder = m_targetFolder;
			targetFolder.TrimRight(L"\\");
			p->SetWindowText(targetFolder);
			p->EnableWindow(FALSE);
		}
		p = GetDlgItem(IDC_EDIT_ARCHIVE_FILE_NAME);
		if (p)
		{
			p->SetWindowText(m_archiveFileName);
			p->EnableWindow(TRUE);
		}
	}

	//UpdateData(TRUE);

	//SetDlgItemText(IDC_STATIC, "Desired Text String")

	//GetDlgItem(IDC_STATIC)->SetWindowText(m_text);

	;// SetWindowText(caption);

	ResHelper::LoadDialogItemsInfo(this);
	ResHelper::UpdateDialogItemsInfo(this);
	BOOL retA = ResHelper::ActivateToolTips(this, m_toolTip);

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}

#if 0
void OpenArchiveFileDlg::OnBnClickedYes()
{
	// TODO: Add your control notification handler code here

	CString mboxFilePath = m_targetFolder + L"\\" + m_archiveFileName;

	if (FileUtils::PathFileExist(mboxFilePath))
	{
		CString txt = L"File \"" + mboxFilePath;
		txt += L"\" exists.\nOverwrite?";
		int answer = MyMessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
			return;
	}

	EndDialog(IDYES);
}
#endif



void OpenArchiveFileDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here

	UpdateData(TRUE);

	CString mboxFilePath = m_targetFolder + L"\\" + m_archiveFileName;

	if (FileUtils::PathFileExist(mboxFilePath))
	{
		CString fmt = L"File\n\n\"%s\"\n\nexists.\nOverwrite?";
		ResHelper::TranslateString(fmt);

		CString txt;
		txt.Format(fmt, mboxFilePath);

		int answer = MyMessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
			return;
	}

	CDialogEx::OnOK();
}

BOOL OpenArchiveFileDlg::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}
