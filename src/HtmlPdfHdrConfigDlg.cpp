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

// HtmlPdfHdrConfigDlg.cpp : implementation file
//

#include "stdafx.h"
#include "FileUtils.h"
#include "HtmlPdfHdrConfigDlg.h"
#include "GenericFontFamilyDlg.h"
#include "MainFrm.h"
#include "afxdialogex.h"


// HtmlPdfHdrConfigDlg dialog

IMPLEMENT_DYNAMIC(HtmlPdfHdrConfigDlg, CDialogEx)

extern const wchar_t *sz_Software_mboxview;

HtmlPdfHdrConfigDlg::HtmlPdfHdrConfigDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_HTML_PDF_HDR_DLG, pParent)
{
	m_bPickFamilyFont = 0;
	m_bPickConcreteFont = 0;

	m_HdrFldConfig.LoadFromRegistry();
}

HtmlPdfHdrConfigDlg::~HtmlPdfHdrConfigDlg()
{
}

void HtmlPdfHdrConfigDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_fldListBox);
	//DDX_Radio(pDX, IDC_PICK_FAMILY_FONT, m_bPickFamilyFont);
	DDX_Radio(pDX, IDC_PICK_CONCRETE_FONT, m_bPickConcreteFont);
	DDX_Radio(pDX, IDC_FONT_DFLT, m_HdrFldConfig.m_bHdrFontDflt);
	DDX_Radio(pDX, IDC_HDR_FLD_NAME, m_HdrFldConfig.m_bHdrFldCustomNameFont);
	DDX_Check(pDX, IDC_FLD_NAME_BOLD, m_HdrFldConfig.m_bHdrBoldFldName);
	DDX_Text(pDX, IDC_FLD_SIZE, m_HdrFldConfig.m_nHdrFontSize);
}


BEGIN_MESSAGE_MAP(HtmlPdfHdrConfigDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &HtmlPdfHdrConfigDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDC_PICK_CONCRETE_FONT, &HtmlPdfHdrConfigDlg::OnBnClickedPickConcreteFont)
	ON_BN_CLICKED(IDC_FLD_NAME_BOLD, &HtmlPdfHdrConfigDlg::OnBnClickedFldNameBold)
	ON_EN_CHANGE(IDC_FLD_SIZE, &HtmlPdfHdrConfigDlg::OnEnChangeFldSize)
	ON_BN_CLICKED(IDC_PICK_FAMILY_FONT, &HtmlPdfHdrConfigDlg::OnBnClickedPickFamilyFont)
	ON_BN_CLICKED(IDC_FONT_DFLT, &HtmlPdfHdrConfigDlg::OnBnClickedFontDflt)
	ON_BN_CLICKED(IDC_FONT_CUSTOM, &HtmlPdfHdrConfigDlg::OnBnClickedFontCustom)
	ON_BN_CLICKED(IDC_HDR_FLD_HELP, &HtmlPdfHdrConfigDlg::OnBnClickedHdrFldHelp)
END_MESSAGE_MAP()


// HtmlPdfHdrConfigDlg message handlers

int HtmlPdfHdrConfigDlg::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	return 0;
}

void HtmlPdfHdrConfigDlg::LoadData()
{
	m_fldListBox.AddString(L"SUBJECT");
	m_fldListBox.AddString(L"FROM");
	m_fldListBox.AddString(L"TO");
	m_fldListBox.AddString(L"CC");
	m_fldListBox.AddString(L"BCC");
	m_fldListBox.AddString(L"DATE");
	m_fldListBox.AddString(L"ATTACHMENTS");

	int i;
	for (i = 0; i < m_fldListBox.GetCount(); i++)
	{
		if (m_HdrFldConfig.m_HdrFldList.IsFldSet(i))
			m_fldListBox.SetCheck(i, BST_CHECKED);
		else
			m_fldListBox.SetCheck(i, BST_UNCHECKED);
	}
}

BOOL HtmlPdfHdrConfigDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here
	m_fldListBox.SetFont(GetFont());

	CString m_title = L"HTML/PDF Mail Header Configuration";
	SetWindowText(m_title);

	LoadData();

	if (m_HdrFldConfig.m_bHdrFontDflt == 0)
		OnBnClickedFontDflt();
	else
		OnBnClickedFontCustom();

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}


void HtmlPdfHdrConfigDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here

	UpdateData(TRUE);

	CString errorTxt;
	CString txt;
	if ((m_HdrFldConfig.m_nHdrFontSize <= 0) || (m_HdrFldConfig.m_nHdrFontSize > 72))
	{
		txt.Format(L"Default Font:  Invalid Font Size of %d! Valid Size is >0 and <72\n", m_HdrFldConfig.m_nHdrFontSize);
		errorTxt.Append(txt);
	}

	if (m_HdrFldConfig.m_HdrFldFontName.m_fontName.IsEmpty())
	{
		errorTxt.Append(L"Field Name Font:  Empty!\n");
	}

	if (m_HdrFldConfig.m_HdrFldFontName.m_fontStyleName.IsEmpty())
	{
		errorTxt.Append(L"Field Name Font Style:  Empty!\n");
	}

	if ((m_HdrFldConfig.m_HdrFldFontName.m_nFontSize <= 0) || (m_HdrFldConfig.m_HdrFldFontName.m_nFontSize > 72))
	{
		txt.Format(L"Field Name Font Size:  Invalid Font Size of %d! Valid Size is >0 and <=72\n", m_HdrFldConfig.m_HdrFldFontName.m_nFontSize);
		errorTxt.Append(txt);
	}

	if (m_HdrFldConfig.m_HdrFldFontText.m_fontName.IsEmpty())
	{
		errorTxt.Append(L"Field Text Font:  Empty!\n");
	}

	if (m_HdrFldConfig.m_HdrFldFontText.m_fontStyleName.IsEmpty())
	{
		errorTxt.Append(L"Field Text Font Style:  Empty!\n");
	}

	if ((m_HdrFldConfig.m_HdrFldFontText.m_nFontSize <= 0) || (m_HdrFldConfig.m_HdrFldFontText.m_nFontSize > 72))
	{
		txt.Format(L"Field Text Font Size:  Invalid Font Size of %d! Valid Size is >0 and <=72\n", m_HdrFldConfig.m_HdrFldFontText.m_nFontSize);
		errorTxt.Append(txt);
	}

	if (!errorTxt.IsEmpty())
	{
		int answer = MessageBox(errorTxt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return;
	}

	m_HdrFldConfig.m_HdrFldList.ClearFldBitmap();

	int i;
	for (i = 0; i < m_fldListBox.GetCount(); i++)
	{
		int chk = m_fldListBox.GetCheck(i);
		if (chk)
			m_HdrFldConfig.m_HdrFldList.SetFld(i);
	}

	m_HdrFldConfig.SaveToRegistry();

	CDialogEx::OnOK();
}

///

HdrFldConfig::HdrFldConfig() :
	m_HdrFldFontName(CString(L"hdrFldName_")), 
	m_HdrFldFontText(CString(L"hdrFldText_"))
{
	m_bHdrFldCustomNameFont = 0;
	m_bHdrFontDflt = 0;
	m_nHdrFontSize = 20;
	m_bHdrBoldFldName = FALSE;
}

////

HdrFldList::HdrFldList()
{
}
HdrFldList::~HdrFldList()
{
}

void HdrFldList::ClearFldBitmap()
{
	m_nFldBitmap = 0x0;
}

void HdrFldList::LoadFldBitmap()
{
	BOOL retval;
	CString section_hdr = CString(sz_Software_mboxview) + L"\\PrintConfig\\HeaderFields";
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_hdr, L"hdrFldListBitmap", m_nFldBitmap))
	{
		; //all done
	}
	else
	{
		m_nFldBitmap = 0xFFFFFFFF;
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_hdr, L"hdrFldListBitmap", m_nFldBitmap);
	}
}

void HdrFldList::LoadFromRegistry()
{
	LoadFldBitmap();
}

void HdrFldList::SaveToRegistry()
{
	CString section_hdr = CString(sz_Software_mboxview) + L"\\PrintConfig\\HeaderFields";
	CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_hdr, L"hdrFldListBitmap", m_nFldBitmap);
}

BOOL HdrFldList::IsFldSet(int fldNumber)
{
	DWORD position = 1 << fldNumber;
	if (m_nFldBitmap & position)
		return TRUE;
	else
		return FALSE;
}

void HdrFldList::ClearFld(int fldNumber)
{
	DWORD position = 1 << fldNumber;
	m_nFldBitmap &= ~position;
}

void HdrFldList::SetFld(int fldNumber)
{
	DWORD position = 1 << fldNumber;
	m_nFldBitmap |= position;
}

void HtmlPdfHdrConfigDlg::OnBnClickedPickConcreteFont()
{
	// TODO: Add your control notification handler code here

	UpdateData(TRUE);

	CWnd *p = GetDlgItem(IDC_PICK_CONCRETE_FONT);
	if (p)
	{
		BOOL chk = ((CButton*)p)->GetCheck();
		if (chk == FALSE)
		{
			return;
		}
	}

	LPLOGFONT lplfInitial = NULL;
	DWORD dwFlags = CF_PRINTERFONTS | CF_SCREENFONTS | CF_NOSCRIPTSEL | CF_USESTYLE;
	//DWORD dwFlags = CF_PRINTERFONTS | CF_SCREENFONTS | CF_NOSCRIPTSEL;
	CDC* pdcPrinter = NULL;
	CWnd* pParentWnd = NULL;

	LOGFONT lf;
	memset(&lf, 0, sizeof(LOGFONT));

	HdrFldFont *pFont = &m_HdrFldConfig.m_HdrFldFontName;
	if (m_HdrFldConfig.m_bHdrFldCustomNameFont)
		pFont = &m_HdrFldConfig.m_HdrFldFontText;

	CClientDC dc(this); // expects a CWnd that has already been initialized
	int fontSize = pFont->m_nFontSize;
	lf.lfHeight = -MulDiv(fontSize, dc.GetDeviceCaps(LOGPIXELSY), 72);

	_tcscpy_s(lf.lfFaceName, LF_FACESIZE, pFont->m_fontName);

	LOGFONT *plf = &lf;

	CFontDialog fdlg(plf, dwFlags, pdcPrinter, pParentWnd);

	_tcsncpy(fdlg.m_cf.lpszStyle, pFont->m_fontStyleName, 64);
	fdlg.m_cf.lpszStyle[63] = 0;

	INT_PTR ret = fdlg.DoModal();
	if (ret == IDOK)
	{
		CString fn = fdlg.GetFaceName();
		CString sn = fdlg.GetStyleName();

		// GetFaceName() is not always returning the font selected in the font dialog.
		// Frequently the font and the font style names are concatenated
		// Until I find the better solution, try below hack
		CString faceNameRoot;
		if (fn.GetLength() > sn.GetLength())
		{
			int len = fn.GetLength() - sn.GetLength();
			faceNameRoot = fn.Mid(len);
			if (faceNameRoot.CompareNoCase(sn) == 0)
			{
				CString rootFont = fn.Left(len);
				fn = rootFont;
				fn.Trim();
				int deb = 1;
			}
		}


		int fs = fdlg.GetSize();
		int fw = fdlg.GetWeight();
		BOOL isBold = fdlg.IsBold();
		BOOL isItalic = fdlg.IsItalic();

		LOGFONT lf;
		fdlg.GetCurrentFont(&lf);
		//
		CHARFORMAT cf;
		cf.cbSize = sizeof(cf);
		fdlg.GetCharFormat(cf);

		pFont->m_fontName = fn;
		pFont->m_fontStyleName = sn;
		pFont->m_nFontStyle = fw;
		pFont->m_bIsBold = isBold;
		pFont->m_bIsItalic = isItalic;
		pFont->m_nFontSize = fs/10;
		if ((pFont->m_nFontSize <= 0) || (pFont->m_nFontSize > 72))
		{
			CString txt;
			CString errorTxt;
			if (m_HdrFldConfig.m_bHdrFldCustomNameFont == 0)
			{
				txt.Format(L"Field Name Font Size:  Invalid Font Size of %d! Valid Size is >0 and <=72\n", pFont->m_nFontSize);
				errorTxt.Append(txt);
				errorTxt.Append(L"Setting Font Size to 16\n");
			}
			else
			{
				txt.Format(L"Field Text Font Size:  Invalid Font Size of %d! Valid Size is >0 and <=72\n", pFont->m_nFontSize);
				errorTxt.Append(txt);
				errorTxt.Append(L"Setting Font Size to 16\n");
			}
			int answer = MessageBox(errorTxt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			pFont->m_nFontSize = 16;
		}

		int deb = 1;
	}

	int deb = 1;
}


void HtmlPdfHdrConfigDlg::OnBnClickedFldNameBold()
{
	// TODO: Add your control notification handler code here
	int deb = 1;
}


void HtmlPdfHdrConfigDlg::OnEnChangeFldSize()
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialogEx::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	// TODO:  Add your control notification handler code here

	int deb = 1;
}


void HtmlPdfHdrConfigDlg::OnBnClickedPickFamilyFont()
{
	// TODO: Add your control notification handler code here

	UpdateData(TRUE);

	CWnd *p = GetDlgItem(IDC_PICK_FAMILY_FONT);
	if (p)
	{
		BOOL chk = ((CButton*)p)->GetCheck();
		if (chk == FALSE)
		{
			return;
		}
	}

	GenericFontFamilyDlg fdlg(this);

	HdrFldFont *pFont = &m_HdrFldConfig.m_HdrFldFontName;
	if (m_HdrFldConfig.m_bHdrFldCustomNameFont)
		pFont = &m_HdrFldConfig.m_HdrFldFontText;

	fdlg.m_genericFontName = pFont->m_genericFontName;

	INT_PTR ret = fdlg.DoModal();

	if (ret == IDOK)
	{

		pFont->m_genericFontName = fdlg.m_genericFontName;

		int deb = 1;
	}
	return;
}


void HtmlPdfHdrConfigDlg::OnBnClickedFontDflt()
{
	// TODO: Add your control notification handler code here

	CWnd *p = GetDlgItem(IDC_PICK_FAMILY_FONT);
	if (p)
	{
		p->EnableWindow(FALSE);
	}

	p = GetDlgItem(IDC_PICK_CONCRETE_FONT);
	if (p)
	{
		p->EnableWindow(FALSE);
	}

	p = GetDlgItem(IDC_HDR_FLD_NAME);
	if (p)
	{
		p->EnableWindow(FALSE);
	}

	p = GetDlgItem(IDC_HDR_FLD_TEXT);
	if (p)
	{
		p->EnableWindow(FALSE);
	}

	p = GetDlgItem(IDC_FLD_NAME_BOLD);
	if (p)
	{
		p->EnableWindow(TRUE);
	}

	p = GetDlgItem(IDC_FLD_SIZE);
	if (p)
	{
		p->EnableWindow(TRUE);
	}

	int deb = 1;
}


void HtmlPdfHdrConfigDlg::OnBnClickedFontCustom()
{
	// TODO: Add your control notification handler code here

	CWnd *p = GetDlgItem(IDC_PICK_FAMILY_FONT);
	if (p)
	{
		p->EnableWindow(TRUE);
	}

	p = GetDlgItem(IDC_PICK_CONCRETE_FONT);
	if (p)
	{
		p->EnableWindow(TRUE);
	}

	p = GetDlgItem(IDC_HDR_FLD_NAME);
	if (p)
	{
		p->EnableWindow(TRUE);
	}

	p = GetDlgItem(IDC_HDR_FLD_TEXT);
	if (p)
	{
		p->EnableWindow(TRUE);
	}

	p = GetDlgItem(IDC_FLD_NAME_BOLD);
	if (p)
	{
		p->EnableWindow(FALSE);
	}

	p = GetDlgItem(IDC_FLD_SIZE);
	if (p)
	{
		p->EnableWindow(FALSE);
	}

	int deb = 1;
}

void HdrFldConfig::SaveToRegistry()
{
	CString section_hdr = CString(sz_Software_mboxview) + L"\\PrintConfig\\HeaderFields";

	CString param = L"ApplyCustomFonts";
	BOOL ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_hdr, param, m_bHdrFontDflt);

	CString section_hdr_sub = CString(sz_Software_mboxview) + L"\\PrintConfig\\HeaderFields\\DefaultFont";

	param = L"IsBold";
	ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_hdr_sub, param, m_bHdrBoldFldName);

	param = L"Size";
	ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_hdr_sub, param, m_nHdrFontSize);

	m_HdrFldList.SaveToRegistry();
	m_HdrFldFontName.SaveToRegistry();
	m_HdrFldFontText.SaveToRegistry();
}

void HdrFldConfig::LoadFromRegistry()
{
	BOOL retval;

	CString section_hdr = CString(sz_Software_mboxview) + L"\\PrintConfig\\HeaderFields";

	CString param = L"ApplyCustomFonts";
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_hdr, param, m_bHdrFontDflt))
	{
		;
	}

	CString section_hdr_sub = CString(sz_Software_mboxview) + L"\\PrintConfig\\HeaderFields\\DefaultFont";

	param = L"Size";
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_hdr_sub, param, m_nHdrFontSize))
	{
		;
	}

	param = L"IsBold";
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_hdr_sub, param, m_bHdrBoldFldName))
	{
		;
	}

	m_HdrFldList.LoadFromRegistry();
	m_HdrFldFontName.LoadFromRegistry();
	m_HdrFldFontText.LoadFromRegistry();
}

HdrFldFont::HdrFldFont(CString &id)
{
	m_id = id;
	SetDflts();
}

void HdrFldFont::SetDflts()
{
	m_genericFontName = L"sans-serif";
	m_fontName = L"Arial";
	m_fontStyleName = L"Regular";
	m_nFontStyle = 400;
	m_bIsBold = FALSE;
	m_bIsItalic = FALSE;
	m_nFontSize = 16;
}

void HdrFldFont::SaveToRegistry()
{
	CString subSection = L"";
	if (m_id.Compare(L"hdrFldName_") == 0)
		subSection = L"NameFont";
	else
		subSection = L"TextFont";

	CString section_hdr_sub = CString(sz_Software_mboxview) + L"\\PrintConfig\\HeaderFields\\" + subSection;

	BOOL retval;

	CString param = L"GenericName";
	retval = CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_hdr_sub, param, m_genericFontName);

	param = L"Name";
	retval = CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_hdr_sub, param, m_fontName);

	param = L"StyleName";
	retval = CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_hdr_sub, param, m_fontStyleName);

	param = L"StyleNumber";
	retval = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_hdr_sub, param, m_nFontStyle);

	param = L"IsBold";
	retval = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_hdr_sub, param, m_bIsBold);

	param = L"IsItalic";
	retval = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_hdr_sub, param, m_bIsItalic);

	param = L"Size";
	retval = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_hdr_sub, param, m_nFontSize);
}

void HdrFldFont::LoadFromRegistry()
{
	BOOL retval;

	CString subSection = L"";
	if (m_id.Compare(L"hdrFldName_") == 0)
		subSection = L"NameFont";
	else
		subSection = L"TextFont";

	CString section_hdr_sub = CString(sz_Software_mboxview) + L"\\PrintConfig\\HeaderFields\\" + subSection;

	CString param = L"GenericName";
	if (retval = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_hdr_sub, param, m_genericFontName))
	{
		;
	}

	param = L"Name";
	if (retval = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_hdr_sub, param, m_fontName))
	{
		;
	}

	param = L"StyleName";
	if (retval = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_hdr_sub, param, m_fontStyleName))
	{
		;
	}

	param = L"StyleNumber";
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_hdr_sub, param, m_nFontStyle))
	{
		;
	}

	param = L"IsBold";
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_hdr_sub, param, m_bIsBold))
	{
		;
	}

	param = L"IsItalic";
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_hdr_sub, param, m_bIsItalic))
	{
		;
	}

	param = L"Size";
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_hdr_sub, param, m_nFontSize))
	{
		;
	}
	int deb = 1;
}


void HtmlPdfHdrConfigDlg::OnBnClickedHdrFldHelp()
{
	// TODO: Add your control notification handler code here

	CString helpFileName = "HTML_PDF_HdrConfigHelp.pdf";
	HWND h = GetSafeHwnd();
	CMainFrame::OpenHelpFile(helpFileName, h);

	return;
}
