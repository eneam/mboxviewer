// PrintConfigDlg.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "PrintConfigDlg.h"
#include "afxdialogex.h"


// PrintConfigDlg dialog

IMPLEMENT_DYNAMIC(PrintConfigDlg, CDialogEx)

PrintConfigDlg::PrintConfigDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_PRINT_DLG, pParent)
{
	m_NamePatternParams.SetDflts();
}

PrintConfigDlg::~PrintConfigDlg()
{
}

void PrintConfigDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Check(pDX, IDC_FILE_DATE, m_NamePatternParams.m_bDate);
	DDX_Check(pDX, IDC_FILE_TIME, m_NamePatternParams.m_bTime);
	DDX_Check(pDX, IDC_FILE_FROM, m_NamePatternParams.m_bFrom);
	DDX_Check(pDX, IDC_FILE_TO, m_NamePatternParams.m_bTo);
	DDX_Check(pDX, IDC_FILE_SUBJECT, m_NamePatternParams.m_bSubject);
	DDX_Check(pDX, IDC_FILE_UNIQUE_ID, m_NamePatternParams.m_bUniqueId);
	DDX_Text(pDX, IDC_FILE_NAME_MAX_SIZE, m_NamePatternParams.m_nFileNameFormatSizeLimit);
}


BEGIN_MESSAGE_MAP(PrintConfigDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &PrintConfigDlg::OnBnClickedOk)
	ON_EN_CHANGE(IDC_FILE_NAME_MAX_SIZE, &PrintConfigDlg::OnEnChangeFileNameMaxSize)
END_MESSAGE_MAP()


// PrintConfigDlg message handlers


void PrintConfigDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	UpdateData();
#if 0
	// Introduced UniqueId so no need for thos ??
	if ((m_NamePatternParams.m_bDate == 0) &&
		(m_NamePatternParams.m_bTime == 0) &&
		(m_NamePatternParams.m_bFrom == 0) &&
		(m_NamePatternParams.m_bTo == 0) &&
		(m_NamePatternParams.m_bSubject == 0)
		)
	{
		CString txt;
		txt.Format("No fields are checked!");
		AfxMessageBox(txt, MB_OK | MB_ICONHAND);
		return;
	}
#endif

	if (m_NamePatternParams.m_nFileNameFormatSizeLimit > 256)
	{
		AfxMessageBox("File Name Size not valid > 255 !", MB_OK | MB_ICONHAND);
		return;
	}

	CDialogEx::OnOK();
}

BOOL PrintConfigDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	if (GetSafeHwnd()) {
		CWnd *p = GetDlgItem(IDC_FILE_UNIQUE_ID);
		if (p) {
			((CButton*)p)->SetCheck(1);
			p->EnableWindow(FALSE);
		}
	}

	UpdateData(TRUE);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}


void PrintConfigDlg::OnEnChangeFileNameMaxSize()
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialogEx::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	// TODO:  Add your control notification handler code here
}

void NamePatternParams::SetDflts()
{
	m_bDate = TRUE;
	m_bTime = TRUE;
	m_bFrom = TRUE;
	m_bTo = TRUE;
	m_bSubject = TRUE;
	m_bUniqueId = TRUE;
	m_nFileNameFormatSizeLimit = 100;
}

void NamePatternParams::Copy(NamePatternParams &src)
{
	if (this == &src)
		return;

	m_bDate = src.m_bDate;
	m_bTime = src.m_bTime;
	m_bFrom = src.m_bFrom;
	m_bTo = src.m_bTo;
	m_bSubject = src.m_bSubject;
	m_bUniqueId = src.m_bUniqueId;
	m_nFileNameFormatSizeLimit = src.m_nFileNameFormatSizeLimit;
}

void NamePatternParams::UpdateRegistry(NamePatternParams &current, NamePatternParams &updated)
{
	if (&current == &updated)
		return;

	if (updated.m_bDate != current.m_bDate) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "printFileNameDate", updated.m_bDate);
	}
	if (updated.m_bTime != current.m_bTime) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "printFileNameTime", updated.m_bTime);
	}
	if (updated.m_bFrom != current.m_bFrom) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "printFileNameFrom", updated.m_bFrom);
	}
	if (updated.m_bTo != current.m_bTo) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "printFileNameTo", updated.m_bTo);
	}
	if (updated.m_bSubject != current.m_bSubject) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "printFileNameSubject", updated.m_bSubject);

	}
	if (updated.m_nFileNameFormatSizeLimit != current.m_nFileNameFormatSizeLimit) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "printFileNameMaxLength", updated.m_nFileNameFormatSizeLimit);
	}
}


void NamePatternParams::LoadFromRegistry()
{
	BOOL retval;
	DWORD bDate, bTime, bFrom, bTo, bSubject, nFileNameFormatSizeLimit;

	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printFileNameDate"), bDate))
		m_bDate = bDate;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printFileNameTime"), bTime))
		m_bTime = bTime;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printFileNameFrom"), bFrom))
		m_bFrom = bFrom;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printFileNameTo"), bTo))
		m_bTo = bTo;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printFileNameSubject"), bSubject))
		m_bSubject = bSubject;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printFileNameMaxLength"), nFileNameFormatSizeLimit))
		m_nFileNameFormatSizeLimit = nFileNameFormatSizeLimit;
}
