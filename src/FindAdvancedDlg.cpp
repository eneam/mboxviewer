// FindDlg.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "FindAdvancedDlg.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

/////////////////////////////////////////////////////////////////////////////
// CFindAdvancedDlg dialog

const char *FindFields[] = { "From", "To", "Subject", "Message", "Attachments" };


CFindAdvancedDlg::CFindAdvancedDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CFindAdvancedDlg::IDD, pParent)

{
	//{{AFX_DATA_INIT(CFindAdvancedDlg)
	m_params.SetDflts();
	//}}AFX_DATA_INIT
}


void CFindAdvancedDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CFindAdvancedDlg)
	DDX_Check(pDX, IDC_EDIT_FROM_CHECKED, m_params.m_bEditChecked[0]);
	DDX_Check(pDX, IDC_EDIT_TO_CHECKED, m_params.m_bEditChecked[1]);
	DDX_Check(pDX, IDC_EDIT_SUBJECT_CHECKED, m_params.m_bEditChecked[2]);
	DDX_Check(pDX, IDC_EDIT_MESSAGE_CHECKED, m_params.m_bEditChecked[3]);
	DDX_Check(pDX, IDC_EDIT_ATTACHMENT_CHECKED, m_params.m_bEditChecked[4]);

	DDX_Text(pDX, IDC_EDIT_FROM, m_params.m_string[0]);
	DDX_Text(pDX, IDC_EDIT_TO, m_params.m_string[1]);
	DDX_Text(pDX, IDC_EDIT_SUBJECT, m_params.m_string[2]);
	DDX_Text(pDX, IDC_EDIT_MESSAGE, m_params.m_string[3]);
	DDX_Text(pDX, IDC_EDIT_ATTACHMENT, m_params.m_string[4]);

	DDX_Check(pDX, IDC_EDIT_FROM_WHOLE, m_params.m_bWholeWord[0]);
	DDX_Check(pDX, IDC_EDIT_TO_WHOLE, m_params.m_bWholeWord[1]);
	DDX_Check(pDX, IDC_EDIT_SUBJECT_WHOLE, m_params.m_bWholeWord[2]);
	DDX_Check(pDX, IDC_EDIT_MESSAGE_WHOLE, m_params.m_bWholeWord[3]);
	DDX_Check(pDX, IDC_EDIT_ATTACHMENT_WHOLE, m_params.m_bWholeWord[4]);

	DDX_Check(pDX, IDC_EDIT_FROM_CASE, m_params.m_bCaseSensitive[0]);
	DDX_Check(pDX, IDC_EDIT_TO_CASE, m_params.m_bCaseSensitive[1]);
	DDX_Check(pDX, IDC_EDIT_SUBJECT_CASE, m_params.m_bCaseSensitive[2]);
	DDX_Check(pDX, IDC_EDIT_MESSAGE_CASE, m_params.m_bCaseSensitive[3]);
	DDX_Check(pDX, IDC_EDIT_ATTACHMENT_CASE, m_params.m_bCaseSensitive[4]);
	//}}AFX_DATA_MAP

	DDX_Check(pDX, IDC_EDIT_SET_ALL_WHOLE, m_params.m_bSetAllWholeWords);
	DDX_Check(pDX, IDC_EDIT_SET_ALL_CASE, m_params.m_bSetAllCaseSensitive);

	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER1, m_params.m_startDate);
	DDX_DateTimeCtrl(pDX, IDC_DATETIMEPICKER2, m_params.m_endDate);

	//DDX_Check(pDX, IDC_CHECK_FIND_ALL, m_params.m_bFindAll);
	DDX_Radio(pDX, IDC_FILTER1, m_params.m_filterNumb);

	DDX_Check(pDX, IDC_FILTERDATES, m_params.m_filterDates);
	if (pDX->m_bSaveAndValidate) {
		GetDlgItem(IDC_DATETIMEPICKER1)->EnableWindow(m_params.m_filterDates);
		GetDlgItem(IDC_DATETIMEPICKER2)->EnableWindow(m_params.m_filterDates);
	}
}


BEGIN_MESSAGE_MAP(CFindAdvancedDlg, CDialog)
	//{{AFX_MSG_MAP(CFindAdvancedDlg)
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_CHECK1, &CFindAdvancedDlg::OnBnClickedFilterDates)
	ON_BN_CLICKED(IDC_EDIT_SET_ALL_WHOLE, &CFindAdvancedDlg::OnBnClickedEditSetAllWhole)
	ON_BN_CLICKED(IDC_EDIT_SET_ALL_CASE, &CFindAdvancedDlg::OnBnClickedEditSetAllCase)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CFindAdvancedDlg message handlers

void CFindAdvancedDlg::OnOK()
{
	UpdateData();

	m_params.m_endDate.SetDate(m_params.m_endDate.GetYear(), m_params.m_endDate.GetMonth(), m_params.m_endDate.GetDay());
	m_params.m_startDate.SetDate(m_params.m_startDate.GetYear(), m_params.m_startDate.GetMonth(), m_params.m_startDate.GetDay());

	int i;
	for (i = 0; i < 5; i++)
	{
		if (m_params.m_bEditChecked[i] == TRUE)
			break;
	}
	if (i == 5)
	{
		CString txt;
		txt.Format("No filter fields are checked!");
		AfxMessageBox(txt, MB_OK | MB_ICONHAND);
		return;
	}

	for (int i = 0; i < 5; i++) 
	{
		
		m_params.m_string[i].TrimRight();
		g_tu.MakeLower(m_params.m_string[i]);
		if ((m_params.m_bEditChecked[i] == TRUE) && m_params.m_string[i].IsEmpty()) {
			CString txt;
			txt.Format("%s field is checked but search string is empty!", FindFields[i]);
			AfxMessageBox(txt, MB_OK | MB_ICONHAND);
			return;
		}
	}

	m_params.m_bBiderectionalMatch = TRUE;
	if (m_params.m_filterNumb != 0)
		m_params.m_bBiderectionalMatch = FALSE;

	//TRACE(_T("Extended: %u %u\n"), m_string.GetAt(0), m_string.GetAt(1));	
	CDialog::OnOK();
}


BOOL CFindAdvancedDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	if (GetSafeHwnd()) {
		CWnd *p = GetDlgItem(IDC_CHECK_FIND_ALL);
		if (p) {
			((CButton*)p)->SetCheck(1);
			p->EnableWindow(FALSE);
		}
	}

	UpdateData(TRUE);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}


void CFindAdvancedDlg::OnBnClickedFilterDates()
{
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
	p = GetDlgItem(IDC_EDIT_MESSAGE_WHOLE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
	p = GetDlgItem(IDC_EDIT_ATTACHMENT_WHOLE);
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
	p = GetDlgItem(IDC_EDIT_MESSAGE_CASE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
	p = GetDlgItem(IDC_EDIT_ATTACHMENT_CASE);
	if (p) {
		((CButton*)p)->SetCheck(nFlags);
	}
}

void CFindAdvancedParams::SetDflts()
{
	for (int i = 0; i < 5; i++) {
		m_string[i] = _T("");
		m_bWholeWord[i] = FALSE;
		m_bCaseSensitive[i] = TRUE;
		m_bEditChecked[i] = FALSE;
	}

	m_bSetAllWholeWords = FALSE;
	m_bSetAllCaseSensitive = TRUE;

	m_startDate = COleDateTime::GetCurrentTime();
	m_endDate = COleDateTime::GetCurrentTime();
	m_filterDates = FALSE;

	m_bFindAll = TRUE;
	m_filterNumb = 0;

	m_bBiderectionalMatch = FALSE;
}

void CFindAdvancedParams::Copy(CFindAdvancedParams &src)
{
	if (this == &src)
		return;

	for (int i = 0; i < 5; i++)
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

	m_bBiderectionalMatch = src.m_bBiderectionalMatch;

};
