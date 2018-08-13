// OptionsDlg.cpp : file di implementazione
//

#include "stdafx.h"
#include "mboxview.h"
#include "OptionsDlg.h"


// finestra di dialogo COptionsDlg

IMPLEMENT_DYNAMIC(COptionsDlg, CDialog)

COptionsDlg::COptionsDlg(CWnd* pParent /*=NULL*/)
	: CDialog(COptionsDlg::IDD, pParent)
	, m_format(0), m_barDelay(0)
	, m_from_charsetId(0), m_to_charsetId(0), m_subj_charsetId(0), m_show_charsets(0)
	, m_bImageViewer(1), m_bTimeType(0)
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
	DDX_Text(pDX, IDC_FROM_ID, m_from_charsetId);
	DDX_Text(pDX, IDC_TO_ID, m_to_charsetId);
	DDX_Text(pDX, IDC_SUBJ_ID, m_subj_charsetId);
	DDX_Radio(pDX, IDC_SHOW_CHARSETS, m_show_charsets);
	DDX_Check(pDX, IDC_IMAGE_VIEWER, m_bImageViewer);
	DDX_Radio(pDX, IDC_TIME_LOCAL, m_bTimeType);
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

		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("fromCharsetId"), m_from_charsetId);
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("toCharsetId"), m_to_charsetId);
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("subjCharsetId"), m_subj_charsetId);
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("showCharsets"), m_show_charsets);
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("imageViewer"), m_bImageViewer);
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "timeType", m_bTimeType);

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

	m_from_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "fromCharsetId");
	m_to_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "toCharsetId");
	m_subj_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "subjCharsetId");
	m_show_charsets = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "showCharsets");
	m_bImageViewer = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "imageViewer");
	m_bTimeType = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "timeType");

	UpdateData(FALSE);
	return TRUE;  // return TRUE unless you set the focus to a control
	// ECCEZIONE: le pagine delle proprietà OCX devono restituire FALSE
}
