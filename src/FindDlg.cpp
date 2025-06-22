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
#include "TextUtilsEx.h"
#include "FindDlg.h"
#include "ResHelper.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

/////////////////////////////////////////////////////////////////////////////
// CFindDlg dialog


CFindDlg::CFindDlg(CWnd* pParent /*=NULL*/)
DIALOG_FROM_TEMPLATE( : CDialogEx(CFindDlg::IDD, pParent))
{
	//{{AFX_DATA_INIT(CFindDlg)
	m_params.SetDflts();
	//m_params.ResetFilterDates();

	m_dflBkColor = ::GetSysColor(COLOR_3DFACE);
	m_checkedColor = RGB(255, 255, 0);

	m_brBkDate.CreateSolidBrush(m_dflBkColor);

	m_pParent = pParent;

	//}}AFX_DATA_INIT
}

INT_PTR CFindDlg::DoModal()
{
#ifdef _DIALOG_FROM_TEMPLATE
	INT_PTR ret = CMainFrame::SetTemplate(this, CFindDlg::IDD, m_pParent);
#else
	INT_PTR ret = CDialogEx::DoModal();
#endif
	return ret;
}

void CFindDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
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
#if 0
	if (pDX->m_bSaveAndValidate) {
		GetDlgItem(IDC_DATETIMEPICKER1)->EnableWindow(m_params.m_filterDates);
		GetDlgItem(IDC_DATETIMEPICKER2)->EnableWindow(m_params.m_filterDates);
	}
#endif
}

BEGIN_MESSAGE_MAP(CFindDlg, CDialogEx)
	//{{AFX_MSG_MAP(CFindDlg)
	//}}AFX_MSG_MAP
	ON_WM_CTLCOLOR()
	ON_BN_CLICKED(IDC_CHECK1, &CFindDlg::OnBnClickedFilterDates)
	ON_BN_CLICKED(IDC_CHECK_FIND_ALL, &CFindDlg::OnBnClickedCheckFindAll)
	ON_NOTIFY(DTN_DATETIMECHANGE, IDC_DATETIMEPICKER1, &CFindDlg::OnDtnDatetimechangeDatetimepicker1)
	ON_NOTIFY(DTN_DATETIMECHANGE, IDC_DATETIMEPICKER2, &CFindDlg::OnDtnDatetimechangeDatetimepicker2)
	ON_BN_CLICKED(IDC_BUTTON1, &CFindDlg::OnBnClickedButtonHelp)
	ON_NOTIFY_EX(TTN_NEEDTEXT, 0, &CFindDlg::OnTtnNeedText)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CFindDlg message handlers

void CFindDlg::OnOK()
{
	// set bSaveAndValidate to FALSE to transfer from class member variables --> UI controls
	// set bSaveAndValidate to TRUE to transfer from UI controls --> class member variables. 

	BOOL bSaveAndValidate = TRUE;  
	UpdateData(bSaveAndValidate);

	SYSTEMTIME endDateSysTime;
	bool res1 = m_params.m_endDate.GetAsSystemTime(endDateSysTime);
	SYSTEMTIME startDateSysTime;
	bool res2 = m_params.m_startDate.GetAsSystemTime(startDateSysTime);

	bool res3 = MyCTime::fixSystemtime(&startDateSysTime);
	bool res4 = MyCTime::fixSystemtime(&endDateSysTime);

	m_params.m_endDate.SetDate(m_params.m_endDate.GetYear(), m_params.m_endDate.GetMonth(), m_params.m_endDate.GetDay());
	m_params.m_startDate.SetDate(m_params.m_startDate.GetYear(), m_params.m_startDate.GetMonth(), m_params.m_startDate.GetDay());

	m_params.m_string.TrimRight();  // FIXMEFIXME do we need to enforce this
	if (m_params.m_string.IsEmpty())
	{
		CString txt = L"Cannot search for an empty string!";
		ResHelper::TranslateString(txt);
		MyAfxMessageBox(txt, MB_OK | MB_ICONHAND);
		return;
	}

	DWORD error;
	TextUtilsEx::WStr2UTF8(&m_params.m_string, &m_params.m_stringA, error);  // FIXMEFIXME

	//TRACE(L"Extended: %u %u\n", m_string.GetAt(0), m_string.GetAt(1));	
	CDialogEx::OnOK();
}

BOOL CFindDlg::OnInitDialog()
{
	// set bSaveAndValidate to FALSE to transfer from class member variables --> UI controls
	// set bSaveAndValidate to TRUE to transfer from UI controls --> class member variables. 

	CDialogEx::OnInitDialog();

	GetDlgItem(IDC_DATETIMEPICKER1)->EnableWindow(m_params.m_filterDates);
	GetDlgItem(IDC_DATETIMEPICKER2)->EnableWindow(m_params.m_filterDates);

	CWnd *p = GetDlgItem(IDC_FILTERDATES);
	CButton* b = (CButton*)p;
	if (b)
	{
		int state = b->GetCheck();
		if (state == BST_CHECKED)
		{
			m_brBkDate.DeleteObject();
			m_brBkDate.CreateSolidBrush(m_checkedColor);
		}
	}

	p = GetDlgItem(IDC_DATETIMEPICKER1);
	if (p)
	{
		BOOL ret = ((CDateTimeCtrl*)p)->SetFormat(m_params.m_dateTimeFormat);
		int deb = 1;
	}
	p = GetDlgItem(IDC_DATETIMEPICKER2);
	if (p)
	{
		BOOL ret = ((CDateTimeCtrl*)p)->SetFormat(m_params.m_dateTimeFormat);
		int deb = 1;
	}

	p = GetDlgItem(IDC_FIND_DATE_FMT);
	if (p)
	{
		CString dateFmt;
		CString wTxt;
		p->GetWindowText(wTxt);

		if (m_params.m_dateTimeFormat.Compare(L"MM/dd/yyyy") == 0)
		{
			dateFmt.LoadStringW(IDS_MONTH_DAY_YEAR);
			//dateFmt.Append(L"month/day/year");
		}
		else if (m_params.m_dateTimeFormat.Compare(L"dd/MM/yyyy") == 0)
		{
			dateFmt.LoadStringW(IDS_DAY_MONTH_YEAR);
			//dateFmt.Append(L"day/month/year");
		}
		else if (m_params.m_dateTimeFormat.Compare(L"yyyy/MM/dd") == 0)
		{
			dateFmt.LoadStringW(IDS_YEAR_MONTH_DAY);
			//dateFmt.Append(L"year/month/day");
		}

		wTxt.Append(dateFmt);
		(p)->SetWindowText(wTxt);
		int deb = 1;
	}

	ResHelper::UpdateDialogItemsInfo(this);
	BOOL retA = ResHelper::ActivateToolTips(this, m_toolTip);

	UpdateData(FALSE);  // FIXMEFIXME

	ResHelper::LoadDialogItemsInfo(this);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

HBRUSH CFindDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	// The below works. It is called for every object in the dialog. May not be the most efficient.
	// TODO: Use MyCButton instead ??
	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);

	if (pWnd->GetDlgCtrlID() == IDC_FILTERDATES)
	{
		return (HBRUSH)m_brBkDate;
	}
	else
		return hbr;
}

void CFindDlg::OnBnClickedFilterDates()
{
	// set bSaveAndValidate to FALSE to transfer from class member variables --> UI controls
	// set bSaveAndValidate to TRUE to transfer from UI controls --> class member variables. 

	// Any problem with updating all class member variables ? ?
	//UpdateData(TRUE);

	//GetDlgItem(IDC_DATETIMEPICKER1)->EnableWindow(m_params.m_filterDates);
	//GetDlgItem(IDC_DATETIMEPICKER2)->EnableWindow(m_params.m_filterDates);

	CWnd* p = GetDlgItem(IDC_FILTERDATES);
	CButton* b = (CButton*)p;
	if (b)
	{
		int state = b->GetCheck();
		m_params.m_filterDates = state;  // or uncomment UpdateData(TRUE); above

		GetDlgItem(IDC_DATETIMEPICKER1)->EnableWindow(state);
		GetDlgItem(IDC_DATETIMEPICKER2)->EnableWindow(state);

		m_brBkDate.DeleteObject();
		if (state == BST_CHECKED)
			m_brBkDate.CreateSolidBrush(m_checkedColor);
		else
			m_brBkDate.CreateSolidBrush(m_dflBkColor);
	}

	int deb = 1;
}

void CFindDlgParams::SetDflts()
{
	m_startDate.SetDate(1970,1,1);
	m_endDate.SetDate(1970,1,1);
	m_bFindNext = TRUE;
	m_filterDates = FALSE;

	m_string = L"";
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

	// Following vars are not set or updated by FindDlg dialog
	// They are to help to maintain proper state
	// see CFindDlg::CFindDlg()

	ResetFilterDates();
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

	// FIXMEFIXME  CopyFilterDates() ??
	m_lastStartDate = src.m_lastStartDate;
	m_lastEndDate = src.m_lastEndDate;
	m_mboxMailStartDate = src.m_mboxMailStartDate;
	m_mboxMailEndDate = src.m_mboxMailEndDate;
	m_needToRestoreArchiveListDateTime = src.m_needToRestoreArchiveListDateTime;
	m_bNeedToFindMailMinMaxTime = src.m_bNeedToFindMailMinMaxTime;
	m_dateTimeFormat = src.m_dateTimeFormat;
	m_last_nWhichMailList = src.m_last_nWhichMailList;
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

void CFindDlg::OnDtnDatetimechangeDatetimepicker1(NMHDR *pNMHDR, LRESULT *pResult)
{
	// set bSaveAndValidate to FALSE to transfer from class member variables --> UI controls
	// set bSaveAndValidate to TRUE to transfer from UI controls --> class member variables. 

	LPNMDATETIMECHANGE pDTChange = reinterpret_cast<LPNMDATETIMECHANGE>(pNMHDR);
	// TODO: Add your control notification handler code here

	SYSTEMTIME sysTime = pDTChange->st;
	MyCTime::fixSystemtime(&sysTime);

	CWnd *p = GetDlgItem(IDC_DATETIMEPICKER1);
	if (p) {
		BOOL ret = ((CDateTimeCtrl*)p)->SetTime(&sysTime);
		int deb = 1;
	}

	UpdateData(TRUE);
	*pResult = 0;
}

void CFindDlg::OnDtnDatetimechangeDatetimepicker2(NMHDR *pNMHDR, LRESULT *pResult)
{
	// set bSaveAndValidate to FALSE to transfer from class member variables --> UI controls
	// set bSaveAndValidate to TRUE to transfer from UI controls --> class member variables. 

	LPNMDATETIMECHANGE pDTChange = reinterpret_cast<LPNMDATETIMECHANGE>(pNMHDR);
	// TODO: Add your control notification handler code here

	SYSTEMTIME sysTime = pDTChange->st;
	MyCTime::fixSystemtime(&sysTime);

	CWnd *p = GetDlgItem(IDC_DATETIMEPICKER2);
	if (p) {
		BOOL ret = ((CDateTimeCtrl*)p)->SetTime(&sysTime);
		int deb = 1;
	}

	UpdateData(TRUE);
	*pResult = 0;
}

CFindDlgParams::CFindDlgParams()
{
	SetDflts();
	ResetFilterDates();
	int deb = 1;
}

CFindDlgParams::~CFindDlgParams()
{
	int deb = 1;
}

void CFindDlgParams::ResetFilterDates()
{
	m_lastStartDate = (time_t)-1;
	m_lastEndDate = (time_t)-1;
	m_mboxMailStartDate = (time_t)-1;
	m_mboxMailEndDate = (time_t)-1;
	m_needToRestoreArchiveListDateTime = FALSE;
	m_bNeedToFindMailMinMaxTime = TRUE;
	m_last_nWhichMailList = 0;
	m_filterDates = FALSE;
	//m_dateTimeFormat = L"MM/dd/yyyy";
}

void CFindDlg::OnBnClickedButtonHelp()
{
	// TODO: Add your control notification handler code here

	CString helpFileName = L"SearchHelp.pdf";
	HWND h = GetSafeHwnd();
	CMainFrame::OpenHelpFile(helpFileName, h);

	int deb = 1;
}

BOOL CFindDlg::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}
