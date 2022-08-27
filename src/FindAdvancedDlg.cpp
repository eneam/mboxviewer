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
#include "FindAdvancedDlg.h"
#include "FindFilterRuleDlg.h"
#include "MboxMail.h"   // looking for MboxMail::developerMode

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

/////////////////////////////////////////////////////////////////////////////
// CFindAdvancedDlg dialog

const char *FindFields[] = { "From", "To", "Subject", "CC", "BCC", "Message", "Attachments", "Attachment Name" };

const TCHAR  *ruleText[] = {
	_T("((From <--> To) and CC and BCC and Subject and (Message Text or Attachment Text) and Attachment Name"),
	_T("(From->To) and CC and BCC and Subject and (Message Text or Attachment Text) and Attachment Name"),
	_T("((From -> (To or CC or BCC)) and Subject and (Message Text or Attachment Text) and Attachment Name"),
	_T("((From <--> To) and no CC and no BCC and Subject and (Message Text or Attachment Text) and Attachment Name"),
	_T("(From->To) and no CC and no BCC and Subject and (Message Text or Attachment Text) and Attachment Name"),
	_T("(From or To or CC or BCC) and Subject and (Message Text or Attachment Text) and Attachment Name")
};


CFindAdvancedDlg::CFindAdvancedDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CFindAdvancedDlg::IDD, pParent)

{
	//{{AFX_DATA_INIT(CFindAdvancedDlg)

	m_params.SetDflts();

	m_dflBkColor = ::GetSysColor(COLOR_3DFACE);
	m_checkedColor = RGB(255, 255, 0);

	m_brBkMailsDontMatch.CreateSolidBrush(m_dflBkColor);
	m_brBkDate.CreateSolidBrush(m_dflBkColor);

	//}}AFX_DATA_INIT
}


void CFindAdvancedDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CFindAdvancedDlg)
	DDX_Check(pDX, IDC_EDIT_FROM_CHECKED, m_params.m_bEditChecked[0]);
	DDX_Check(pDX, IDC_EDIT_TO_CHECKED, m_params.m_bEditChecked[1]);
	DDX_Check(pDX, IDC_EDIT_SUBJECT_CHECKED, m_params.m_bEditChecked[2]);
	DDX_Check(pDX, IDC_EDIT_CC_CHECKED, m_params.m_bEditChecked[3]);
	DDX_Check(pDX, IDC_EDIT_BCC_CHECKED, m_params.m_bEditChecked[4]);
	DDX_Check(pDX, IDC_EDIT_MESSAGE_CHECKED, m_params.m_bEditChecked[5]);
	DDX_Check(pDX, IDC_EDIT_ATTACHMENT_CHECKED, m_params.m_bEditChecked[6]);
	DDX_Check(pDX, IDC_EDIT_ATTACHMENT_NAME_CHECKED, m_params.m_bEditChecked[7]);

	DDX_Text(pDX, IDC_EDIT_FROM, m_params.m_string[0]);
	DDX_Text(pDX, IDC_EDIT_TO, m_params.m_string[1]);
	DDX_Text(pDX, IDC_EDIT_SUBJECT, m_params.m_string[2]);
	DDX_Text(pDX, IDC_EDIT_CC, m_params.m_string[3]);
	DDX_Text(pDX, IDC_EDIT_BCC, m_params.m_string[4]);
	DDX_Text(pDX, IDC_EDIT_MESSAGE, m_params.m_string[5]);
	DDX_Text(pDX, IDC_EDIT_ATTACHMENT, m_params.m_string[6]);
	DDX_Text(pDX, IDC_EDIT_ATTACHMENT_NAME, m_params.m_string[7]);

	DDX_Check(pDX, IDC_EDIT_FROM_WHOLE, m_params.m_bWholeWord[0]);
	DDX_Check(pDX, IDC_EDIT_TO_WHOLE, m_params.m_bWholeWord[1]);
	DDX_Check(pDX, IDC_EDIT_SUBJECT_WHOLE, m_params.m_bWholeWord[2]);
	DDX_Check(pDX, IDC_EDIT_CC_WHOLE, m_params.m_bWholeWord[3]);
	DDX_Check(pDX, IDC_EDIT_BCC_WHOLE, m_params.m_bWholeWord[4]);
	DDX_Check(pDX, IDC_EDIT_MESSAGE_WHOLE, m_params.m_bWholeWord[5]);
	DDX_Check(pDX, IDC_EDIT_ATTACHMENT_WHOLE, m_params.m_bWholeWord[6]);
	DDX_Check(pDX, IDC_EDIT_ATTACHMENT_NAME_WHOLE, m_params.m_bWholeWord[7]);

	DDX_Check(pDX, IDC_EDIT_FROM_CASE, m_params.m_bCaseSensitive[0]);
	DDX_Check(pDX, IDC_EDIT_TO_CASE, m_params.m_bCaseSensitive[1]);
	DDX_Check(pDX, IDC_EDIT_SUBJECT_CASE, m_params.m_bCaseSensitive[2]);
	DDX_Check(pDX, IDC_EDIT_CC_CASE, m_params.m_bCaseSensitive[3]);
	DDX_Check(pDX, IDC_EDIT_BCC_CASE, m_params.m_bCaseSensitive[4]);
	DDX_Check(pDX, IDC_EDIT_MESSAGE_CASE, m_params.m_bCaseSensitive[5]);
	DDX_Check(pDX, IDC_EDIT_ATTACHMENT_CASE, m_params.m_bCaseSensitive[6]);
	DDX_Check(pDX, IDC_EDIT_ATTACHMENT_NAME_CASE, m_params.m_bCaseSensitive[7]);
	//}}AFX_DATA_MAP

	DDX_Check(pDX, IDC_EDIT_SET_ALL_WHOLE, m_params.m_bSetAllWholeWords);
	DDX_Check(pDX, IDC_EDIT_SET_ALL_CASE, m_params.m_bSetAllCaseSensitive);

	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER1, m_params.m_startDate);
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER2, m_params.m_endDate);

	DDX_Check(pDX, IDC_SINGLE_TO, m_params.m_bSingleTo);
	//DDX_Check(pDX, IDC_CHECK_FIND_ALL, m_params.m_bFindAll);

	DDX_Check(pDX, IDC_PLAIN_TEXT, m_params.m_plainText);
	DDX_Check(pDX, IDC_HTML_TEXT, m_params.m_htmlText);

	DDX_Check(pDX, IDC_CHECK_NEGATE_FIND_CRITERIA, m_params.m_bFindAllMailsThatDontMatch);

	DDX_Check(pDX, IDC_FILTERDATES, m_params.m_filterDates);
	if (pDX->m_bSaveAndValidate) {
		GetDlgItem(IDC_DATETIMEPICKER1)->EnableWindow(m_params.m_filterDates);
		GetDlgItem(IDC_DATETIMEPICKER2)->EnableWindow(m_params.m_filterDates);
	}
}


BEGIN_MESSAGE_MAP(CFindAdvancedDlg, CDialog)
	//{{AFX_MSG_MAP(CFindAdvancedDlg)

	//}}AFX_MSG_MAP
	ON_WM_CTLCOLOR()
	ON_BN_CLICKED(IDC_CHECK1, &CFindAdvancedDlg::OnBnClickedFilterDates)
	ON_BN_CLICKED(IDC_EDIT_SET_ALL_WHOLE, &CFindAdvancedDlg::OnBnClickedEditSetAllWhole)
	ON_BN_CLICKED(IDC_EDIT_SET_ALL_CASE, &CFindAdvancedDlg::OnBnClickedEditSetAllCase)
	ON_BN_CLICKED(IDC_EDIT_RESET, &CFindAdvancedDlg::OnBnClickedEditReset)
	ON_NOTIFY(DTN_DATETIMECHANGE, IDC_DATETIMEPICKER1, &CFindAdvancedDlg::OnDtnDatetimechangeDatetimepicker1)
	ON_NOTIFY(DTN_DATETIMECHANGE, IDC_DATETIMEPICKER2, &CFindAdvancedDlg::OnDtnDatetimechangeDatetimepicker2)
	ON_BN_CLICKED(IDC_BUTTON1, &CFindAdvancedDlg::OnBnClickedButtonSelectFilterRule)
	ON_BN_CLICKED(IDC_CHECK_NEGATE_FIND_CRITERIA, &CFindAdvancedDlg::OnBnClickedCheckNegateFindCriteria)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CFindAdvancedDlg message handlers

void CFindAdvancedDlg::OnOK()
{
	UpdateData();

	if ((m_params.m_plainText == 0) && (m_params.m_htmlText == 0))
	{
		CString txt;
		txt.Format(_T("No text type plain and/or html are checked!"));
		AfxMessageBox(txt, MB_OK | MB_ICONHAND);
		return;
	}

	m_params.m_endDate.SetDate(m_params.m_endDate.GetYear(), m_params.m_endDate.GetMonth(), m_params.m_endDate.GetDay());
	m_params.m_startDate.SetDate(m_params.m_startDate.GetYear(), m_params.m_startDate.GetMonth(), m_params.m_startDate.GetDay());

	int i;
	for (i = 0; i < FILTER_FIELDS_NUMB; i++)
	{
		if (m_params.m_bEditChecked[i] == TRUE)
			break;
	}
	if (i == FILTER_FIELDS_NUMB)
	{
		CString txt;
		txt.Format(_T("No filter fields are checked!"));
		AfxMessageBox(txt, MB_OK | MB_ICONHAND);
		return;
	}


	for (int i = 0; i < FILTER_FIELDS_NUMB; i++)
	{
		m_params.m_string[i].TrimRight();
		g_tu.MakeLower(m_params.m_string[i]);

		if ((m_params.m_bEditChecked[i] == TRUE) && m_params.m_string[i].IsEmpty()) {
			CString txt;
			txt.Format(_T("%s field is checked but search string is empty!"), FindFields[i]);
			AfxMessageBox(txt, MB_OK | MB_ICONHAND);
			return;
		}
	}

	//TRACE(_T("Extended: %u %u\n"), m_string.GetAt(0), m_string.GetAt(1));	
	CDialog::OnOK();
}


BOOL CFindAdvancedDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	if (GetSafeHwnd()) 
	{
		CWnd *p = GetDlgItem(IDC_CHECK_FIND_ALL);
		if (p) 
		{
			((CButton*)p)->SetCheck(1);
			p->EnableWindow(FALSE);
		}

		p = GetDlgItem(IDC_PLAIN_TEXT);
		if (p)
		{
			if (MboxMail::developerMode)
			{
				((CButton*)p)->SetCheck(m_params.m_plainText);
				p->EnableWindow(TRUE);
			}
			else
			{
				m_params.m_plainText = 1;
				((CButton*)p)->SetCheck(m_params.m_plainText);
				p->EnableWindow(FALSE);

			}
		}

		p = GetDlgItem(IDC_HTML_TEXT);
		if (p)
		{
			if (MboxMail::developerMode)
			{
				((CButton*)p)->SetCheck(m_params.m_htmlText);
				p->EnableWindow(TRUE);
			}
			else
			{
				m_params.m_htmlText = 0;
				((CButton*)p)->SetCheck(m_params.m_htmlText);
				p->EnableWindow(FALSE);
			}
		}

		SetRuleInfoText();

		p = GetDlgItem(IDC_FILTERDATES);
		CButton *b = (CButton*)p;
		if (b)
		{
			int state = b->GetCheck();
			if (state == BST_CHECKED)
			{
				m_brBkDate.DeleteObject();
				m_brBkDate.CreateSolidBrush(m_checkedColor);
			}
		}
		p = GetDlgItem(IDC_CHECK_NEGATE_FIND_CRITERIA);
		b = (CButton*)p;
		if (b)
		{
			int state = b->GetCheck();
			if (state == BST_CHECKED)
			{
				m_brBkMailsDontMatch.DeleteObject();
				m_brBkMailsDontMatch.CreateSolidBrush(m_checkedColor);
			}
		}
	}

	UpdateData(TRUE);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}


void CFindAdvancedDlg::OnBnClickedFilterDates()
{
	CWnd *p = GetDlgItem(IDC_FILTERDATES);
	CButton *b = (CButton*)p;
	if (b)
	{
		int state = b->GetCheck();
		m_brBkDate.DeleteObject();
		if (state == BST_CHECKED)
			m_brBkDate.CreateSolidBrush(m_checkedColor);
		else
			m_brBkDate.CreateSolidBrush(m_dflBkColor);
	}

	UpdateData(TRUE);
}

void CFindAdvancedDlg::OnBnClickedEditSetAllWhole()
{
	// TODO: Add your control notification handler code here
	// Should make Ids consecutive and employ the for loop

	if (!GetSafeHwnd())
		return;

	int nFlags = 0;
	CWnd *p = GetDlgItem(IDC_EDIT_SET_ALL_WHOLE);
	if (p) {
		nFlags = ((CButton*)p)->GetCheck();
	}

	p = GetDlgItem(IDC_EDIT_FROM_WHOLE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
	p = GetDlgItem(IDC_EDIT_TO_WHOLE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
	p = GetDlgItem(IDC_EDIT_SUBJECT_WHOLE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
	p = GetDlgItem(IDC_EDIT_CC_WHOLE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
	p = GetDlgItem(IDC_EDIT_BCC_WHOLE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
	p = GetDlgItem(IDC_EDIT_MESSAGE_WHOLE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
	p = GetDlgItem(IDC_EDIT_ATTACHMENT_WHOLE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
	p = GetDlgItem(IDC_EDIT_ATTACHMENT_NAME_WHOLE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
}


void CFindAdvancedDlg::OnBnClickedEditSetAllCase()
{
	// TODO: Add your control notification handler code here
	// Should make Ids consecutive and employ the for loop

	if (!GetSafeHwnd())
		return;

	int nFlags = 0;
	CWnd *p = GetDlgItem(IDC_EDIT_SET_ALL_CASE);
	if (p) {
		nFlags = ((CButton*)p)->GetCheck();
	}

	p = GetDlgItem(IDC_EDIT_FROM_CASE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
	p = GetDlgItem(IDC_EDIT_TO_CASE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
	p = GetDlgItem(IDC_EDIT_SUBJECT_CASE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
	p = GetDlgItem(IDC_EDIT_CC_CASE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
	p = GetDlgItem(IDC_EDIT_BCC_CASE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
	p = GetDlgItem(IDC_EDIT_MESSAGE_CASE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
	p = GetDlgItem(IDC_EDIT_ATTACHMENT_CASE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
	p = GetDlgItem(IDC_EDIT_ATTACHMENT_NAME_CASE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
}

void CFindAdvancedParams::SetDflts()
{
	for (int i = 0; i < FILTER_FIELDS_NUMB; i++) 
	{
		m_string[i] = _T("");
		m_bWholeWord[i] = FALSE;
		m_bCaseSensitive[i] = FALSE;
		m_bEditChecked[i] = FALSE;
	}

	m_bSetAllWholeWords = FALSE;
	m_bSetAllCaseSensitive = FALSE;

	m_startDate.SetDate(1970, 1, 1);
	m_endDate.SetDate(1970, 1, 1);
	m_filterDates = FALSE;

	m_bFindAll = TRUE;
	m_filterNumb = 0;
	m_bSingleTo = 0;
	m_plainText = 1;
	m_htmlText = 1;

	m_bFindAllMailsThatDontMatch = FALSE;
}

void CFindAdvancedParams::Copy(CFindAdvancedParams &src)
{
	if (this == &src)
		return;

	for (int i = 0; i < FILTER_FIELDS_NUMB; i++)
	{
		m_string[i] = src.m_string[i];
		m_bWholeWord[i] = src.m_bWholeWord[i];
		m_bCaseSensitive[i] = src.m_bCaseSensitive[i];
		m_bEditChecked[i] = src.m_bEditChecked[i];
	}
	m_bSetAllWholeWords = src.m_bSetAllWholeWords;
	m_bSetAllCaseSensitive = src.m_bSetAllCaseSensitive;

	m_startDate = src.m_startDate;
	m_endDate = src.m_endDate;
	m_filterDates = src.m_filterDates;

	m_bFindAll = src.m_bFindAll;
	m_filterNumb = src.m_filterNumb;
	m_bSingleTo = src.m_bSingleTo;

	m_plainText = src.m_plainText;
	m_htmlText = src.m_htmlText;

	m_bFindAllMailsThatDontMatch = src.m_bFindAllMailsThatDontMatch;
};


void CFindAdvancedDlg::OnBnClickedEditReset()
{
	// TODO: Add your control notification handler code here

	int m_filterNumb = m_params.m_filterNumb;
	int m_bSingleTo = m_params.m_bSingleTo;
	// TODO: Need to reset to initial range
	COleDateTime startDate = m_params.m_startDate;
	COleDateTime endDate = m_params.m_endDate;

	m_params.SetDflts();

	m_params.m_filterNumb = m_filterNumb;
	//m_params.m_bSingleTo = m_bSingleTo;
	m_params.m_startDate = startDate;
	m_params.m_endDate = endDate;

	m_brBkMailsDontMatch.DeleteObject();
	m_brBkDate.DeleteObject();
	m_brBkMailsDontMatch.CreateSolidBrush(m_dflBkColor);
	m_brBkDate.CreateSolidBrush(m_dflBkColor);

	GetDlgItem(IDC_DATETIMEPICKER1)->EnableWindow(m_params.m_filterDates);
	GetDlgItem(IDC_DATETIMEPICKER2)->EnableWindow(m_params.m_filterDates);

	UpdateData(FALSE);
}


void CFindAdvancedDlg::OnDtnDatetimechangeDatetimepicker1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMDATETIMECHANGE pDTChange = reinterpret_cast<LPNMDATETIMECHANGE>(pNMHDR);
	// TODO: Add your control notification handler code here

	SYSTEMTIME sysTime = pDTChange->st;
	MyCTime::fixSystemtime(&sysTime);

	CWnd *p = GetDlgItem(IDC_DATETIMEPICKER1);
	if (p) 
	{
		BOOL ret = ((CDateTimeCtrl*)p)->SetTime(&sysTime);
		int deb = 1;
	}

	UpdateData(TRUE);
	*pResult = 0;
}


void CFindAdvancedDlg::OnDtnDatetimechangeDatetimepicker2(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMDATETIMECHANGE pDTChange = reinterpret_cast<LPNMDATETIMECHANGE>(pNMHDR);
	// TODO: Add your control notification handler code here

	SYSTEMTIME sysTime = pDTChange->st;
	MyCTime::fixSystemtime(&sysTime);

	CWnd *p = GetDlgItem(IDC_DATETIMEPICKER2);
	if (p) 
	{
		BOOL ret = ((CDateTimeCtrl*)p)->SetTime(&sysTime);
		int deb = 1;
	}

	UpdateData(TRUE);
	*pResult = 0;
}


void CFindAdvancedDlg::OnBnClickedButtonSelectFilterRule()
{
	// TODO: Add your control notification handler code here


	FindFilterRuleDlg m_FindFilterRuleDlg;

	m_FindFilterRuleDlg.m_filterNumb = m_params.m_filterNumb;

	INT_PTR ret = m_FindFilterRuleDlg.DoModal();
	if (ret == IDOK)
	{
		m_params.m_filterNumb = m_FindFilterRuleDlg.m_filterNumb;

		SetRuleInfoText();
	}
}

void CFindAdvancedDlg::SetRuleInfoText()
{
	if (GetSafeHwnd())
	{
		CWnd *p = GetDlgItem(IDC_RULE_NUMBER_TEXT);
		if (p)
		{
			CString RuleNumberText;

			RuleNumberText.Format(_T("Rule %d is selected !!!"), m_params.m_filterNumb + 1);

			p->SetWindowText(RuleNumberText);
			p->EnableWindow(TRUE);
		}

		if ((m_params.m_filterNumb >= 0) && (m_params.m_filterNumb <= 5))
		{
			CWnd *p = GetDlgItem(IDC_RULE_TEXT);
			if (p)
			{
				p->SetWindowText(ruleText[m_params.m_filterNumb]);
				p->EnableWindow(TRUE);
			}
		}
		if ((m_params.m_filterNumb == 3) || (m_params.m_filterNumb == 4))
		{
			CWnd *p = GetDlgItem(IDC_EDIT_CC_CHECKED);
			CButton *b = (CButton*)p;
			if (p && p->IsWindowEnabled())
			{
				b->SetCheck(BST_UNCHECKED);
				p->EnableWindow(FALSE);
			}
			p = GetDlgItem(IDC_EDIT_BCC_CHECKED);
			b = (CButton*)p;
			if (p && p->IsWindowEnabled())
			{
				b->SetCheck(BST_UNCHECKED);
				p->EnableWindow(FALSE);
			}
		}
		else
		{
			CWnd *p = GetDlgItem(IDC_EDIT_CC_CHECKED);
			CButton *b = (CButton*)p;
			if (p && !p->IsWindowEnabled())
			{
				b->SetCheck(BST_UNCHECKED);
				p->EnableWindow(TRUE);
			}
			p = GetDlgItem(IDC_EDIT_BCC_CHECKED);
			b = (CButton*)p;
			if (p && !p->IsWindowEnabled())
			{
				b->SetCheck(BST_UNCHECKED);
				p->EnableWindow(TRUE);
			}
		}
	}
}


void CFindAdvancedDlg::OnBnClickedCheckNegateFindCriteria()
{
	// TODO: Add your control notification handler code here

	CWnd *p = GetDlgItem(IDC_CHECK_NEGATE_FIND_CRITERIA);
	CButton *b = (CButton*)p;
	if (b)
	{
		int state = b->GetCheck();
		m_brBkMailsDontMatch.DeleteObject();
		if (state == BST_CHECKED)
			m_brBkMailsDontMatch.CreateSolidBrush(m_checkedColor);
		else
			m_brBkMailsDontMatch.CreateSolidBrush(m_dflBkColor);
	}
}

HBRUSH CFindAdvancedDlg::OnCtlColor(CDC* pDC, CWnd *pWnd, UINT nCtlColor)
{
	// The below works. It is called for every object in the dialog. May not be the most efficient.
	// TODO: Use MyCButton instead ??
	HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);

	if (pWnd->GetDlgCtrlID() == IDC_CHECK_NEGATE_FIND_CRITERIA)
	{
		return (HBRUSH)m_brBkMailsDontMatch;
	}
	else if (pWnd->GetDlgCtrlID() == IDC_FILTERDATES)
	{
		return (HBRUSH)m_brBkDate;
	}
	else
		return hbr;
}
