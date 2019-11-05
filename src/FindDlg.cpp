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

// FindDlg.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "FindDlg.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

/////////////////////////////////////////////////////////////////////////////
// CFindDlg dialog


CFindDlg::CFindDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CFindDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CFindDlg)
	m_params.SetDflts();
	//}}AFX_DATA_INIT
}


void CFindDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CFindDlg)
	DDX_Text(pDX, IDC_STRING, m_params.m_string);
	DDX_Check(pDX, IDC_WHOLE, m_params.m_bWholeWord);
	DDX_Check(pDX, IDC_CASE, m_params.m_bCaseSensitive);
	//}}AFX_DATA_MAP
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER1, m_params.m_startDate);
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER2, m_params.m_endDate);
	DDX_Radio(pDX, IDC_SEARCH_PREV, m_params.m_bFindNext);
	DDX_Check(pDX, IDC_CHECK_FROM, m_params.m_bFrom);
	DDX_Check(pDX, IDC_CHECK_TO, m_params.m_bTo);
	DDX_Check(pDX, IDC_CHECK_SUBJECT, m_params.m_bSubject);
	DDX_Check(pDX, IDC_CHECK_CONTENT, m_params.m_bContent);
	DDX_Check(pDX, IDC_CHECK_CC, m_params.m_bCC);
	DDX_Check(pDX, IDC_CHECK_BCC, m_params.m_bBCC);
	DDX_Check(pDX, IDC_CHECK_ATTACHMENT, m_params.m_bAttachments);
	DDX_Check(pDX, IDC_CHECK_ATTACHMENT_NAME, m_params.m_bAttachmentName);
	DDX_Check(pDX, IDC_CHECK_HIGHLIGHT_ALL, m_params.m_bHighlightAll);
	DDX_Check(pDX, IDC_CHECK_FIND_ALL, m_params.m_bFindAll);
	DDX_Check(pDX, IDC_CHECK_NEGATE_SEARCH_CRITERIA, m_params.m_bFindAllMailsThatDontMatch);
	DDX_Check(pDX, IDC_FILTERDATES, m_params.m_filterDates);
	if (pDX->m_bSaveAndValidate) {
		GetDlgItem(IDC_DATETIMEPICKER1)->EnableWindow(m_params.m_filterDates);
		GetDlgItem(IDC_DATETIMEPICKER2)->EnableWindow(m_params.m_filterDates);
	}
}


BEGIN_MESSAGE_MAP(CFindDlg, CDialog)
	//{{AFX_MSG_MAP(CFindDlg)
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_CHECK1, &CFindDlg::OnBnClickedFilterDates)
	ON_BN_CLICKED(IDC_CHECK_FIND_ALL, &CFindDlg::OnBnClickedCheckFindAll)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CFindDlg message handlers

void CFindDlg::OnOK()
{
	UpdateData();

	m_params.m_endDate.SetDate(m_params.m_endDate.GetYear(), m_params.m_endDate.GetMonth(), m_params.m_endDate.GetDay());
	m_params.m_startDate.SetDate(m_params.m_startDate.GetYear(), m_params.m_startDate.GetMonth(), m_params.m_startDate.GetDay());

	m_params.m_string.TrimRight();
	g_tu.MakeLower(m_params.m_string);
	if (m_params.m_string.IsEmpty()) {
		AfxMessageBox("Cannot search for an empty string!", MB_OK | MB_ICONHAND);
		return;
	}

	//TRACE(_T("Extended: %u %u\n"), m_string.GetAt(0), m_string.GetAt(1));	
	CDialog::OnOK();
}


BOOL CFindDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	UpdateData(TRUE);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}


void CFindDlg::OnBnClickedFilterDates()
{
	UpdateData(TRUE);
}


void CFindDlgParams::SetDflts()
{
	m_startDate = COleDateTime::GetCurrentTime();
	m_endDate = COleDateTime::GetCurrentTime();
	m_bFindNext = TRUE;
	m_filterDates = FALSE;

	m_string = _T("");
	m_bWholeWord = FALSE;
	m_bCaseSensitive = FALSE;

	m_bFrom = TRUE;
	m_bTo = TRUE;
	m_bSubject = TRUE;
	m_bContent = FALSE;
	m_bCC = FALSE;
	m_bBCC = FALSE;
	m_bAttachments = FALSE;
	m_bAttachmentName = FALSE;
	m_bHighlightAll = FALSE;
	m_bFindAll = FALSE;
	m_bFindAllMailsThatDontMatch = FALSE;
}

void CFindDlgParams::Copy(CFindDlgParams &src)
{
	if (this == &src)
		return;

	m_startDate = src.m_startDate;
	m_endDate = src.m_endDate;
	m_bFindNext = src.m_bFindNext;
	m_filterDates = src.m_filterDates;

	m_string = src.m_string;
	m_bWholeWord = src.m_bWholeWord;
	m_bCaseSensitive = src.m_bCaseSensitive;

	m_bFrom = src.m_bFrom;
	m_bTo = src.m_bTo;
	m_bSubject = src.m_bSubject;
	m_bContent = src.m_bContent;
	m_bCC = src.m_bCC;
	m_bBCC = src.m_bBCC;
	m_bAttachments = src.m_bAttachments;
	m_bAttachmentName = src.m_bAttachmentName;
	m_bHighlightAll = src.m_bHighlightAll;
	m_bFindAll = src.m_bFindAll;
	m_bFindAllMailsThatDontMatch = src.m_bFindAllMailsThatDontMatch;
}



void CFindDlg::OnBnClickedCheckFindAll()
{
	// TODO: Add your control notification handler code here

	if (!GetSafeHwnd())
		return;

	int nFlags = 0;
	CWnd *p = GetDlgItem(IDC_CHECK_FIND_ALL);
	if (p) {
		nFlags = ((CButton*)p)->GetCheck();
	}

	if (nFlags) {
		CWnd *p = GetDlgItem(IDC_SEARCH_PREV);
		if (p) {
			p->EnableWindow(TRUE);
			((CButton*)p)->SetCheck(0);
			p->EnableWindow(FALSE);
		}
		p = GetDlgItem(IDC_SEARCH_NEXT);
		if (p) {
			p->EnableWindow(TRUE);
			((CButton*)p)->SetCheck(1);
		}
	}
	else
	{
		CWnd *p = GetDlgItem(IDC_SEARCH_PREV);
		if (p) {
			p->EnableWindow(TRUE);
			((CButton*)p)->SetCheck(0);
		}
		p = GetDlgItem(IDC_SEARCH_NEXT);
		if (p) {
			p->EnableWindow(TRUE);
			((CButton*)p)->SetCheck(1);
		}
	}
}
