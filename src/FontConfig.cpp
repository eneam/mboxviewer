// FontConfig.cpp : implementation file
//

#include "stdafx.h"
#include "afxdialogex.h"
#include "FontConfig.h"
#include "MainFrm.h"
#include "ResHelper.h"


// FontConfig dialog

IMPLEMENT_DYNAMIC(FontConfig, CDialogEx)

FontConfig::FontConfig(CWnd* pParent /*=nullptr*/)
DIALOG_FROM_TEMPLATE( : CDialogEx(IDD_DIALOG_FONT_SIZE, pParent))
{
	m_pParent = pParent;
	m_dflFontSizeSelected = 0;
	m_fontSize = 8;
}

FontConfig::~FontConfig()
{
}

INT_PTR FontConfig::DoModal()
{
#ifdef _DIALOG_FROM_TEMPLATE
	INT_PTR ret = CMainFrame::SetTemplate(this, IDD_DIALOG_FONT_SIZE, m_pParent);
#else
	INT_PTR ret = CDialogEx::DoModal();
#endif
	return ret;
}

void FontConfig::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Radio(pDX, IDC_RADIO_FONT_DFLT, m_dflFontSizeSelected);
	DDX_Text(pDX, IDC_EDIT_CUSTOM_FONT_SIZE, m_fontSize);
}


BEGIN_MESSAGE_MAP(FontConfig, CDialogEx)
	ON_BN_CLICKED(IDOK, &FontConfig::OnBnClickedOk)
	ON_BN_CLICKED(IDC_RADIO_FONT_DFLT, &FontConfig::OnBnClickedRadioFontDflt)
	ON_BN_CLICKED(IDC_RADIO_FONT_CUSTOM, &FontConfig::OnBnClickedRadioFontCustom)
	ON_NOTIFY_EX(TTN_NEEDTEXT, 0, &FontConfig::OnTtnNeedText)
END_MESSAGE_MAP()


// FontConfig message handlers

void FontConfig::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here

	UpdateData(FromControlsToVariables);

	if ((m_fontSize < 8) || (m_fontSize > 72))
	{
		CString txt = L"Font Size out of valid range 8-72 !!!";
		ResHelper::TranslateString(txt);
		AfxMessageBox(txt, MB_OK | MB_ICONHAND);
		return;
	}

	CDialogEx::OnOK();
}

void FontConfig::OnBnClickedRadioFontDflt()
{
	// TODO: Add your control notification handler code here

	CWnd* p = GetDlgItem(IDC_EDIT_CUSTOM_FONT_SIZE);
	if (p)
	{
		p->EnableWindow(FALSE);
	}

	m_fontSize = CMainFrame::m_dfltFontSize;
	m_dflFontSizeSelected = 0;

	UpdateData(FromVariablesToControls);

	int deb = 1;
}

BOOL FontConfig::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	//InitDialogControls();

	if (m_fontSize != CMainFrame::m_dfltFontSize)
		m_dflFontSizeSelected = 1;

	if (m_dflFontSizeSelected == 0)
		OnBnClickedRadioFontDflt();
	else
		OnBnClickedRadioFontCustom();

	UpdateData(FromVariablesToControls);

	ResHelper::UpdateDialogItemsInfo(this);
	BOOL retA = ResHelper::ActivateToolTips(this, m_toolTip);

	UpdateData(FromControlsToVariables);

	ResHelper::LoadDialogItemsInfo(this);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

void FontConfig::OnBnClickedRadioFontCustom()
{
	// TODO: Add your control notification handler code here

	CWnd* p = GetDlgItem(IDC_EDIT_CUSTOM_FONT_SIZE);
	if (p)
	{
		p->EnableWindow(TRUE);
	}
}

BOOL FontConfig::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}
