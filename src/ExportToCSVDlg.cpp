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

// ExportToCSVDlg.cpp : implementation file
//

#include "stdafx.h"
#include "TextUtilsEx.h"
#include "mboxview.h"
#include "ExportToCSVDlg.h"
#include "afxdialogex.h"
#include "ResHelper.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

// ExportToCSVDlg dialog

IMPLEMENT_DYNAMIC(ExportToCSVDlg, CDialogEx)

ExportToCSVDlg::ExportToCSVDlg(CWnd* pParent /*=nullptr*/)
DIALOG_FROM_TEMPLATE(  : CDialogEx(IDD_EXPORT_TO_CSV, pParent))
{
	//{{AFX_DATA_INIT(ExportToCSVDlg)
	m_bFrom = TRUE;
	m_bTo = TRUE;
	m_bSubject = TRUE;
	m_bDate = TRUE;
	m_bCC = FALSE;
	m_bBCC = FALSE;
	m_bContent = FALSE;
	m_bAttachmentNames = FALSE;
	m_AttachmentNamesSeparatorString.Append(L" CRLF");
	//m_bAttachmentNames = TRUE;
	m_bGMTTime = 0;
	m_dateFormat = 0;
	m_bEncodingType = 1;  // UTF8
	m_nCodePageId = CP_UTF8;
	m_MessageLimitCharsString.Append(L"32500");
	//}}AFX_DATA_INIT
	m_pParent = pParent;
}

ExportToCSVDlg::~ExportToCSVDlg()
{
}

INT_PTR ExportToCSVDlg::DoModal()
{
#ifdef _DIALOG_FROM_TEMPLATE
	INT_PTR ret = CMainFrame::SetTemplate(this, IDD_EXPORT_TO_CSV, m_pParent);
#else
	INT_PTR ret = CDialogEx::DoModal();
#endif
	return ret;
}

void ExportToCSVDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(ExportToCSVDlg)
	DDX_Check(pDX, IDC_EXPORT_FROM, m_bFrom);
	DDX_Check(pDX, IDC_EXPORT_TO, m_bTo);
	DDX_Check(pDX, IDC_EXPORT_SUBJECT, m_bSubject);
	DDX_Check(pDX, IDC_EXPORT_DATE, m_bDate);
	DDX_Check(pDX, IDC_EXPORT_CC, m_bCC);
	DDX_Check(pDX, IDC_EXPORT_BCC, m_bBCC);
	DDX_Check(pDX, IDC_EXPORT_MESSAGE, m_bContent);
	DDX_Check(pDX, IDC_EXPORT_ATTACHMENT_NAMES, m_bAttachmentNames);
	DDX_Text(pDX, IDC_EXPORT_ATTACHMENT_SEPARATOR, m_AttachmentNamesSeparatorString);
	DDX_Text(pDX, IDC_EXPORT_MESSAGE_LIMIT, m_MessageLimitString);
	DDX_Text(pDX, IDC_EXPORT_MESSAGE_LIMIT_CHARS, m_MessageLimitCharsString);
	DDX_Radio(pDX, IDC_TIME_LOCAL, m_bGMTTime);
	DDX_Radio(pDX, IDC_DMY, m_dateFormat);
	DDX_Radio(pDX, IDC_ENCODING_NONE, m_bEncodingType);
	DDX_Text(pDX, IDC_ENCODING_CODE_PAGE_ID, m_nCodePageId);
	int deb = 1;
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(ExportToCSVDlg, CDialogEx)
ON_BN_CLICKED(IDC_ENCODING_NONE, &ExportToCSVDlg::OnBnClickedEncodingNone)
ON_BN_CLICKED(IDC_ENCODING_UTF8, &ExportToCSVDlg::OnBnClickedEncodingUtf8)
ON_BN_CLICKED(IDC_ENCODING_CODE_PAGE, &ExportToCSVDlg::OnBnClickedEncodingCodePage)
ON_NOTIFY_EX(TTN_NEEDTEXT, 0, &ExportToCSVDlg::OnTtnNeedText)
END_MESSAGE_MAP()


// ExportToCSVDlg message handlers


BOOL ExportToCSVDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here
	GetDlgItem(IDC_ENCODING_CODE_PAGE_ID)->EnableWindow(FALSE);

	ResHelper::UpdateDialogItemsInfo(this);
	BOOL retA = ResHelper::ActivateToolTips(this, m_toolTip);

	UpdateData(TRUE);

	ResHelper::LoadDialogItemsInfo(this);

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}


void ExportToCSVDlg::OnOK()
{
	// TODO: Add your specialized code here and/or call the base class
	//UpdateData(TRUE);  // Control to Data Optional in OnOk

	if ((m_nCodePageId <= 0) || (m_nCodePageId > 65001))
	{
		CString txt = L"Code Page Id out of valid range 0-65001 !";
		ResHelper::TranslateString(txt);
		AfxMessageBox(txt, MB_OK | MB_ICONHAND);
		return;
	}

	BOOL validString = TRUE;

	m_MessageLimitString.Trim();
	if (!m_MessageLimitString.IsEmpty())
	{
		if (TextUtilsEx::isNumeric(m_MessageLimitString)) {
			int limit = _ttoi(m_MessageLimitString);
			if (limit < 0)
				validString = FALSE;
			else
				validString = TRUE;

		}
		else
			validString = FALSE;
	}

	if (validString == FALSE)
	{
		CString txt = L"Invalid line limit!";
		ResHelper::TranslateString(txt);
		AfxMessageBox(txt, MB_OK | MB_ICONHAND);
		return;
	}

	validString = TRUE;

	m_MessageLimitCharsString.Trim();
	if (!m_MessageLimitCharsString.IsEmpty())
	{
		if (TextUtilsEx::isNumeric(m_MessageLimitCharsString)) {
			int limit = _ttoi(m_MessageLimitCharsString);
			if (limit < 0)
				validString = FALSE;
			else
				validString = TRUE;

		}
		else
			validString = FALSE;
	}

	if (validString == FALSE)
	{
		CString txt = L"Invalid character limit!";
		ResHelper::TranslateString(txt);
		AfxMessageBox(L"Invalid character limit!", MB_OK | MB_ICONHAND);
		return;
	}

	const wchar_t* strData = (LPCWSTR)m_AttachmentNamesSeparatorString;

	if (m_AttachmentNamesSeparatorString.Find('\"') >= 0)
	{
		CString txt = L"Character \" not supported as separator!";
		ResHelper::TranslateString(txt);
		AfxMessageBox(txt, MB_OK | MB_ICONHAND);
		return;
	}

	if (m_AttachmentNamesSeparatorString.Find('\'') >= 0)
	{
		CString txt = L"Character ' not supported as separator!";
		ResHelper::TranslateString(txt);
		AfxMessageBox(txt, MB_OK | MB_ICONHAND);
		return;
	}

	CDialogEx::OnOK();
}


void ExportToCSVDlg::OnBnClickedEncodingNone()
{
	// TODO: Add your control notification handler code here
	GetDlgItem(IDC_ENCODING_CODE_PAGE_ID)->EnableWindow(FALSE);
	int deb = 1;
}


void ExportToCSVDlg::OnBnClickedEncodingUtf8()
{
	// TODO: Add your control notification handler code here
	GetDlgItem(IDC_ENCODING_CODE_PAGE_ID)->EnableWindow(FALSE);
	int deb = 1;
}


void ExportToCSVDlg::OnBnClickedEncodingCodePage()
{
	// TODO: Add your control notification handler code here
	GetDlgItem(IDC_ENCODING_CODE_PAGE_ID)->EnableWindow(TRUE);
	int dreb = 1;
}

BOOL ExportToCSVDlg::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}
