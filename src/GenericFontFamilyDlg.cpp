// GenericFontFamilyDlg.cpp : implementation file
//

#include "stdafx.h"
#include "GenericFontFamilyDlg.h"
#include "afxdialogex.h"


// GenericFontFamilyDlg dialog

IMPLEMENT_DYNAMIC(GenericFontFamilyDlg, CDialogEx)

GenericFontFamilyDlg::GenericFontFamilyDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_GENERIC_FONT_DLG, pParent)
{
	m_genericFontName = "serif";

}

GenericFontFamilyDlg::~GenericFontFamilyDlg()
{
}

void GenericFontFamilyDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_listBox);
}


BEGIN_MESSAGE_MAP(GenericFontFamilyDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &GenericFontFamilyDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// GenericFontFamilyDlg message handlers

void GenericFontFamilyDlg::LoadData()
{
	m_listBox.AddString(_T("serif"));
	m_listBox.AddString(_T("sans-serif"));
	m_listBox.AddString(_T("monospace"));
	m_listBox.AddString(_T("cursive"));
	m_listBox.AddString(_T("fantasy"));
}

BOOL GenericFontFamilyDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here


	m_listBox.SetFont(GetFont());

	LoadData();

	//m_genericFontName = "none";
	//m_listBox.SetCurSel(0);

	CString fontFamilyName;

	for (int i = 0; i < m_listBox.GetCount(); i++)
	{
		m_listBox.GetText(i, fontFamilyName);
		if (fontFamilyName.CompareNoCase(m_genericFontName) == 0)
		{
			m_listBox.SetCurSel(i);
			break;
		}
	}

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}


void GenericFontFamilyDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here

	int nSel = m_listBox.GetCurSel();
#if 0
	if ((nSel < 0) || (nSel >= ColorStyleConfig::MailMaxPanes) || (nSel >= m_listBox.GetCount()))
	{
		CSCAssert();
		return;
	}
#endif

	m_listBox.GetText(nSel, m_genericFontName);

	int deb = 1;

	CDialogEx::OnOK();
}
