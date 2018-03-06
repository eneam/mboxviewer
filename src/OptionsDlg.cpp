// OptionsDlg.cpp : file di implementazione
//

#include "stdafx.h"
#include "mboxview.h"
#include "OptionsDlg.h"


// finestra di dialogo COptionsDlg

IMPLEMENT_DYNAMIC(COptionsDlg, CDialog)

COptionsDlg::COptionsDlg(CWnd* pParent /*=NULL*/)
	: CDialog(COptionsDlg::IDD, pParent)
	, m_format(0)
{

}

COptionsDlg::~COptionsDlg()
{
}

void COptionsDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Radio(pDX, IDC_DMY, m_format);
	DDX_Radio(pDX, IDC_EXPORT_EML, m_exportEML);
	DDX_Text(pDX, IDC_PROGRESS_BAR_DELAY, m_barDelay);
}


BEGIN_MESSAGE_MAP(COptionsDlg, CDialog)
	ON_BN_CLICKED(IDOK, &COptionsDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// gestori di messaggi COptionsDlg


void COptionsDlg::OnBnClickedOk()
{
	if (UpdateData(TRUE)) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "format", m_format);
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("progressBarDelay"), m_barDelay);
		if (m_exportEML == 1)
			CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("exportEML"), CString(_T("y")));
		else
			CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("exportEML"), CString(_T("n")));

		CDialog::OnOK();
	}
}


BOOL COptionsDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	m_format = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "format");
	DWORD barDelay = 0;
	BOOL retval;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("progressBarDelay"), barDelay))
		m_barDelay = barDelay;
	else
		m_barDelay = 1;

	CString exportEML;
	if (retval = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("exportEML"), exportEML)) {
		if (exportEML.Compare(_T("y")) == 0)
			m_exportEML = 1;
		else
			m_exportEML = 0;
	}
	else
		m_exportEML = 0;  // we should not be here; it should be initialized in NListView::NListView()

	UpdateData(FALSE);
	return TRUE;  // return TRUE unless you set the focus to a control
	// ECCEZIONE: le pagine delle proprietà OCX devono restituire FALSE
}
