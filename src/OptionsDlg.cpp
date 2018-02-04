// OptionsDlg.cpp : file di implementazione
//

#include "stdafx.h"
#include "mboxview.h"
#include "OptionsDlg.h"


// finestra di dialogo COptionsDlg

IMPLEMENT_DYNAMIC(COptionsDlg, CDialog)

COptionsDlg::COptionsDlg(CWnd* pParent /*=NULL*/)
	: CDialog(COptionsDlg::IDD, pParent)
	, m_format(FALSE)
{

}

COptionsDlg::~COptionsDlg()
{
}

void COptionsDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Radio(pDX, IDC_DMY, m_format);
}


BEGIN_MESSAGE_MAP(COptionsDlg, CDialog)
	ON_BN_CLICKED(IDOK, &COptionsDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// gestori di messaggi COptionsDlg


void COptionsDlg::OnBnClickedOk()
{
	UpdateData(TRUE);
	CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "format", m_format);
	CDialog::OnOK();
}


BOOL COptionsDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	m_format = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "format");
	UpdateData(FALSE);
	return TRUE;  // return TRUE unless you set the focus to a control
	// ECCEZIONE: le pagine delle proprietà OCX devono restituire FALSE
}
