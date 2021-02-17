//
//////////////////////////////////////////////////////////////////
//
//  Windows Mbox Viewer is a free tool to view, search and print mbox mail archives..
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

// PrintConfigDlg.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "PrintConfigDlg.h"
#include "afxdialogex.h"
#include "MboxMail.h"
#include "CustomNameTemplDlg.h"
#include "FileUtils.h"
#include "HtmlUtils.h"
#include "HtmlPdfHdrConfigDlg.h"


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
	DDX_Text(pDX, IDC_FILE_NAME_ADJUSTED_MAX_SIZE, m_NamePatternParams.m_nFileNameFormatSizeLimit);
	DDX_Text(pDX, IDC_FILE_NAME_MAX_SIZE, m_NamePatternParams.m_nWantedFileNameFormatSizeLimit);
	DDX_Radio(pDX, IDC_PRT_DO_NOT_PROMPT, m_NamePatternParams.m_bPrintDialog);
	DDX_Radio(pDX, IDC_HTML2PDF_SCRIPT_TYPE, m_NamePatternParams.m_bScriptType);
	DDX_Text(pDX, IDC_CHROME_EXE_PATH, m_NamePatternParams.m_BrowserPath);
	//DDX_Text(pDX, IDC_MSEDGE_EXE_PATH, m_NamePatternParams.m_MSEdgeBrowserPath);
	DDX_Text(pDX, IDC_HTML2PDF_SCRIPT_PATH, m_NamePatternParams.m_UserDefinedScriptPath);
	DDX_Check(pDX, IDC_SEPARATE_PDF, m_NamePatternParams.m_bPrintToSeparatePDFFiles);
	DDX_Check(pDX, IDC_SEPARATE_HTML, m_NamePatternParams.m_bPrintToSeparateHTMLFiles);
	DDX_Check(pDX, IDC_ADD_HDR_COLOR, m_NamePatternParams.m_bAddBackgroundColorToMailHeader);
	DDX_Check(pDX, IDC_PAGE_BREAK, m_NamePatternParams.m_bAddBreakPageAfterEachMailInPDF);
	DDX_Check(pDX, IDC_KEEP_BODY_BKGRND_COLOR, m_NamePatternParams.m_bKeepMailBodyBackgroundColor);
	DDX_Check(pDX, IDC_HEADER_AND_FOOTER, m_NamePatternParams.m_bHeaderAndFooter);
	DDX_Check(pDX, IDC_CUSTOM_TEMPLATE, m_NamePatternParams.m_bCustomFormat);
	int deb = 1;
}


BEGIN_MESSAGE_MAP(PrintConfigDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &PrintConfigDlg::OnBnClickedOk)
	ON_EN_CHANGE(IDC_FILE_NAME_MAX_SIZE, &PrintConfigDlg::OnEnChangeFileNameMaxSize)
	ON_BN_CLICKED(IDC_PRT_PAGE_SETP, &PrintConfigDlg::OnBnClickedPrtPageSetp)
	ON_BN_CLICKED(IDC_HTML2PDF_SCRIPT_TYPE, &PrintConfigDlg::OnBnClickedHtml2pdfScriptType)
	ON_BN_CLICKED(IDC_USER_SCRIPT, &PrintConfigDlg::OnBnClickedUserDefinedScript)
	ON_BN_CLICKED(IDC_MSEDGE, &PrintConfigDlg::OnBnClickedMSEdge)
	ON_BN_CLICKED(IDC_CUSTOM_TEMPLATE, &PrintConfigDlg::OnBnClickedCustomTemplate)
	ON_BN_CLICKED(IDC_SET_CUSTOM_TEMPLATE, &PrintConfigDlg::OnBnClickedSetCustomTemplate)
	ON_BN_CLICKED(IDC_HTML_PDF_CNF, &PrintConfigDlg::OnBnClickedHtmlPdfCnf)
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

	NamePatternParams::UpdateFilePrintconfig(m_NamePatternParams);

	CDialogEx::OnOK();
}

BOOL PrintConfigDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	if (GetSafeHwnd()) 
	{
		CWnd *p = GetDlgItem(IDC_FILE_UNIQUE_ID);
		if (p) {
			((CButton*)p)->SetCheck(1);
			p->EnableWindow(FALSE);
		}
		p = GetDlgItem(IDC_FILE_NAME_ADJUSTED_MAX_SIZE);
		if (p) {
			NamePatternParams::UpdateFilePrintconfig(m_NamePatternParams);
			CString adjustedMaxSize;
			adjustedMaxSize.Format("%d", m_NamePatternParams.m_nFileNameFormatSizeLimit);
			((CButton*)p)->SetWindowText(adjustedMaxSize);
			p->EnableWindow(FALSE);
		}

		p = GetDlgItem(IDC_CHROME_EXE_PATH);
		if (p)
		{
			if (m_NamePatternParams.m_bScriptType == 0)
			{
				m_NamePatternParams.m_BrowserPath = m_NamePatternParams.m_ChromeBrowserPath;
			}
			else if (m_NamePatternParams.m_bScriptType == 1)
			{
				m_NamePatternParams.m_BrowserPath = m_NamePatternParams.m_MSEdgeBrowserPath;
			}
			p->SetWindowText(m_NamePatternParams.m_BrowserPath);
		}

		p = GetDlgItem(IDC_KEEP_BODY_BKGRND_COLOR);
		if (p) {
			if ((m_NamePatternParams.m_bScriptType == 0) || (m_NamePatternParams.m_bScriptType == 1))
			{
				((CButton*)p)->SetCheck(1);
				p->EnableWindow(FALSE);
			}
			else
				p->EnableWindow(TRUE);
		}

		p = GetDlgItem(IDC_HEADER_AND_FOOTER);
		if (p) {
			if (m_NamePatternParams.m_bScriptType == 0)
			{
				((CButton*)p)->SetCheck(1);
				p->EnableWindow(FALSE);
			}
			else
			{
				((CButton*)p)->SetCheck(m_NamePatternParams.m_bHeaderAndFooter);
				p->EnableWindow(TRUE);
			}
		}

		p = 0;
		p = GetDlgItem(IDC_SET_CUSTOM_TEMPLATE);
		if (p)
		{
			if (m_NamePatternParams.m_bCustomFormat)
				p->EnableWindow(TRUE);
			else
				p->EnableWindow(FALSE);
		}

		if (m_NamePatternParams.m_bCustomFormat)
			EnableNonCustomWindows(FALSE);
	}

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
	NamePatternParams::UpdateFilePrintconfig(m_NamePatternParams);
	if (GetSafeHwnd())
	{
		CString wantedMaxSize;
		CWnd *p = GetDlgItem(IDC_FILE_NAME_MAX_SIZE);
		if (p) {
			((CButton*)p)->GetWindowText(wantedMaxSize);
			int nWantedMaxSize = atoi((LPCSTR)wantedMaxSize);
			if (nWantedMaxSize > 260)
			{
				wantedMaxSize = "260";
				((CButton*)p)->SetWindowText(wantedMaxSize);
			}
		}
		else
			return;

		m_NamePatternParams.m_nWantedFileNameFormatSizeLimit = atoi((LPCSTR)wantedMaxSize);
		NamePatternParams::UpdateFilePrintconfig(m_NamePatternParams);
		p = GetDlgItem(IDC_FILE_NAME_ADJUSTED_MAX_SIZE);
		if (p) {
			CString adjustedMaxSize;
			adjustedMaxSize.Format("%d", m_NamePatternParams.m_nFileNameFormatSizeLimit);
			((CButton*)p)->SetWindowText(adjustedMaxSize);
			p->EnableWindow(FALSE);
		}
	}
	int deb = 1;
}

void NamePatternParams::SetDflts()
{
	m_bDate = TRUE;
	m_bTime = TRUE;
	m_bFrom = TRUE;
	m_bTo = TRUE;
	m_bSubject = TRUE;
	m_bUniqueId = TRUE;
	m_nFileNameFormatSizeLimit = 260;
	m_nWantedFileNameFormatSizeLimit = m_nFileNameFormatSizeLimit;
	m_bPrintDialog = 1;
	m_bScriptType = 1;
	m_ChromeBrowserPath = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe";
	m_MSEdgeBrowserPath = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";
	m_BrowserPath = m_MSEdgeBrowserPath;
	m_bPrintToSeparatePDFFiles = FALSE;
	m_bPrintToSeparateHTMLFiles = FALSE;
	m_bPrintToSeparateTEXTFiles = FALSE;

	m_bAddBackgroundColorToMailHeader = TRUE;
	m_bAddBreakPageAfterEachMailInPDF = FALSE;
	m_bKeepMailBodyBackgroundColor = TRUE;
	m_bHeaderAndFooter = FALSE;
	m_bCustomFormat = FALSE;

	m_nameTemplateCnf.SetDflts();

	AddressParts2Bitmap();

	
	char *pValue;
	errno_t  er = _get_pgmptr(&pValue);
	CString procFullPath;
	if ((er == 0) && pValue)
		procFullPath.Append(pValue);

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(procFullPath, driveName, directory, fileNameBase, fileNameExtention);

	CString procPath;
	BOOL ret = FileUtils::CPathGetPath(procFullPath, procPath);

	m_UserDefinedScriptPath = procPath + "\\scripts\\HTML2PDF-single-wkhtmltopdf.cmd";

	int deb = 1;
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
	m_nWantedFileNameFormatSizeLimit = src.m_nWantedFileNameFormatSizeLimit;
	m_bPrintDialog = src.m_bPrintDialog;
	m_bScriptType = src.m_bScriptType;
	m_BrowserPath = src.m_BrowserPath;
	m_ChromeBrowserPath = src.m_ChromeBrowserPath;
	m_MSEdgeBrowserPath = src.m_MSEdgeBrowserPath;
	m_UserDefinedScriptPath = src.m_UserDefinedScriptPath;
	m_bPrintToSeparatePDFFiles = src.m_bPrintToSeparatePDFFiles;
	m_bPrintToSeparateHTMLFiles = src.m_bPrintToSeparateHTMLFiles;
	m_bPrintToSeparateTEXTFiles = src.m_bPrintToSeparateTEXTFiles;
	m_bAddBackgroundColorToMailHeader = src.m_bAddBackgroundColorToMailHeader;
	m_bAddBreakPageAfterEachMailInPDF = src.m_bAddBreakPageAfterEachMailInPDF;
	m_bKeepMailBodyBackgroundColor = src.m_bKeepMailBodyBackgroundColor;
	m_bHeaderAndFooter = src.m_bHeaderAndFooter;
	m_bCustomFormat = src.m_bCustomFormat;

	m_nameTemplateCnf.Copy(src.m_nameTemplateCnf);
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
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "printFileNameMaxLength", updated.m_nWantedFileNameFormatSizeLimit);
	}
	if (updated.m_bPrintDialog != current.m_bPrintDialog) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "printDialogType", updated.m_bPrintDialog);
	}
	if (updated.m_bScriptType != current.m_bScriptType) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "printHTML2PDFScriptType", updated.m_bScriptType);
	}
	if (updated.m_bScriptType == 0)
	{
		if (updated.m_BrowserPath != current.m_ChromeBrowserPath) {
			updated.m_ChromeBrowserPath = updated.m_BrowserPath;
			CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "printChromeBrowserPath", updated.m_ChromeBrowserPath);
		}
	}
	else if (updated.m_bScriptType == 1)
	{
		if (updated.m_BrowserPath != current.m_MSEdgeBrowserPath) {
			updated.m_MSEdgeBrowserPath = updated.m_BrowserPath;
			CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "printMSEdgeBrowserPath", updated.m_MSEdgeBrowserPath);
		}
	}
	if (updated.m_UserDefinedScriptPath != current.m_UserDefinedScriptPath) {
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "printUserDefinedScriptPath", updated.m_UserDefinedScriptPath);
	}
	if (updated.m_bAddBackgroundColorToMailHeader != current.m_bAddBackgroundColorToMailHeader) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "printMailHdrBackgroundColor", updated.m_bAddBackgroundColorToMailHeader);
	}
	if (updated.m_bAddBreakPageAfterEachMailInPDF != current.m_bAddBreakPageAfterEachMailInPDF) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "printPDFPageBreakAfterEachMail", updated.m_bAddBreakPageAfterEachMailInPDF);
	}
	if (updated.m_bKeepMailBodyBackgroundColor != current.m_bKeepMailBodyBackgroundColor) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "printMailBodyBackgroundColor", updated.m_bKeepMailBodyBackgroundColor);
	}
	if (updated.m_bHeaderAndFooter != current.m_bHeaderAndFooter) {
		if (updated.m_bScriptType == 1)
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "printPageHeaderAndFooter", updated.m_bHeaderAndFooter);
		else
			updated.m_bHeaderAndFooter = current.m_bHeaderAndFooter;
	}
	if (updated.m_bCustomFormat != current.m_bCustomFormat) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "printFileNameCustomTemplate", updated.m_bCustomFormat);
	}
	if (updated.m_nameTemplateCnf.m_TemplateFormat != current.m_nameTemplateCnf.m_TemplateFormat) {
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "printFileNameCustomTemplateFormat", updated.m_nameTemplateCnf.m_TemplateFormat);
	}
	if (updated.m_nameTemplateCnf.m_DateFormat != current.m_nameTemplateCnf.m_DateFormat) {
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "printFileNameCustomTemplateDateFormat", updated.m_nameTemplateCnf.m_DateFormat);
	}
	AddressParts2Bitmap();
	if (updated.m_nAddressPartsBitmap != current.m_nAddressPartsBitmap) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "printAddressPartsBitmap", updated.m_nAddressPartsBitmap);
	}
}

void NamePatternParams::LoadFromRegistry()
{
	BOOL retval;
	DWORD bDate, bTime, bFrom, bTo, bSubject, bPrintDialog, bScriptType, nWantedFileNameFormatSizeLimit;
	DWORD bAddBackgroundColorToMailHeader, bAddBreakPageAfterEachMailInPDF, bKeepMailBodyBackgroundColor, bHeaderAndFooter;
	DWORD nAddressPartsBitmap;
	CString chromeBrowserPath;
	CString msedgeBrowserPath;
	CString userDefinedScriptPath;
	DWORD bCustomFormat;
	CString templateFormat;
	CString dateFormat;

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
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printFileNameMaxLength"), nWantedFileNameFormatSizeLimit))
		m_nWantedFileNameFormatSizeLimit = nWantedFileNameFormatSizeLimit;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printDialogType"), bPrintDialog))
		m_bPrintDialog = bPrintDialog;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printHTML2PDFScriptType"), bScriptType))
		m_bScriptType = bScriptType;
	if (retval = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printChromeBrowserPath"), chromeBrowserPath))
		m_ChromeBrowserPath = chromeBrowserPath;
	if (retval = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printMSEdgeBrowserPath"), msedgeBrowserPath))
		m_MSEdgeBrowserPath = msedgeBrowserPath;

	if (retval = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printUserDefinedScriptPath"), userDefinedScriptPath))
		m_UserDefinedScriptPath = userDefinedScriptPath;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printMailHdrBackgroundColor"), bAddBackgroundColorToMailHeader))
		m_bAddBackgroundColorToMailHeader = bAddBackgroundColorToMailHeader;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printPDFPageBreakAfterEachMail"), bAddBreakPageAfterEachMailInPDF))
		m_bAddBreakPageAfterEachMailInPDF = bAddBreakPageAfterEachMailInPDF;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printMailBodyBackgroundColor"), bKeepMailBodyBackgroundColor))
		m_bKeepMailBodyBackgroundColor = bKeepMailBodyBackgroundColor;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printPageHeaderAndFooter"), bHeaderAndFooter))
		m_bHeaderAndFooter = bHeaderAndFooter;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printFileNameCustomTemplate"), bCustomFormat))
		m_bCustomFormat = bCustomFormat;
	if (retval = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printFileNameCustomTemplateFormat"), templateFormat))
		m_nameTemplateCnf.m_TemplateFormat = templateFormat;
	if (retval = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printFileNameCustomTemplateDateFormat"), dateFormat))
		m_nameTemplateCnf.m_DateFormat = dateFormat;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("printAddressPartsBitmap"), nAddressPartsBitmap))
	{
		m_nAddressPartsBitmap = nAddressPartsBitmap;
		Bitmap2AddressParts();
	}

	if (m_bScriptType == 0)
	{
		m_bHeaderAndFooter = TRUE;
		if (chromeBrowserPath.IsEmpty())
			m_BrowserPath = m_ChromeBrowserPath;
		else
			m_BrowserPath = chromeBrowserPath;
	}
	else if (m_bScriptType == 1)
	{
		if (chromeBrowserPath.IsEmpty())
			m_BrowserPath = m_MSEdgeBrowserPath;
		else
			m_BrowserPath = msedgeBrowserPath;
	}
	else
		m_bHeaderAndFooter = TRUE;
}

void NamePatternParams::UpdateFilePrintconfig(struct NamePatternParams &namePatternParams)
{
	CString printCacheName;
	MboxMail::GetPrintCachePath(printCacheName);
	int maxFileSize = _MAX_PATH - printCacheName.GetLength() - 1 - 4 - 14; // -1 for flder separatorr, -4 for suffix, -14 for fudge factor (HTML_GROUP and PDF_GROUP)
	if (maxFileSize < 0)
		maxFileSize = 0;

	if (namePatternParams.m_nWantedFileNameFormatSizeLimit > maxFileSize)
	{
		namePatternParams.m_nFileNameFormatSizeLimit = maxFileSize;
	}
	else
		namePatternParams.m_nFileNameFormatSizeLimit = namePatternParams.m_nWantedFileNameFormatSizeLimit;

}


void PrintConfigDlg::OnBnClickedPrtPageSetp()
{
	// TODO: Add your control notification handler code here
	HtmlUtils::PrintToPrinterPageSetup(this);
}


void PrintConfigDlg::OnBnClickedHtml2pdfScriptType()
{
	// TODO: Add your control notification handler code here

	if (GetSafeHwnd())
	{
		CWnd *p = GetDlgItem(IDC_KEEP_BODY_BKGRND_COLOR);
		if (p) {
			((CButton*)p)->SetCheck(1);
			p->EnableWindow(FALSE);
		}

		p = GetDlgItem(IDC_CHROME_EXE_PATH);
		if (p)
		{
			m_NamePatternParams.m_BrowserPath = m_NamePatternParams.m_ChromeBrowserPath;
			p->SetWindowText(m_NamePatternParams.m_BrowserPath);
			p->EnableWindow(TRUE);
		}

		p = GetDlgItem(IDC_HEADER_AND_FOOTER);
		if (p) {
			((CButton*)p)->SetCheck(1);
			p->EnableWindow(FALSE);
		}
	}
	int deb = 1;
}

void PrintConfigDlg::OnBnClickedMSEdge()
{
	// TODO: Add your control notification handler code here

	if (GetSafeHwnd())
	{
		CWnd *p = GetDlgItem(IDC_KEEP_BODY_BKGRND_COLOR);
		if (p) {
			((CButton*)p)->SetCheck(1);
			p->EnableWindow(FALSE);
		}

		p = GetDlgItem(IDC_CHROME_EXE_PATH);
		if (p)
		{
			m_NamePatternParams.m_BrowserPath = m_NamePatternParams.m_MSEdgeBrowserPath;
			p->SetWindowText(m_NamePatternParams.m_BrowserPath);
			p->EnableWindow(TRUE);
		}
		p = GetDlgItem(IDC_HEADER_AND_FOOTER);
		if (p) {
			((CButton*)p)->SetCheck(m_NamePatternParams.m_bHeaderAndFooter);
			p->EnableWindow(TRUE);
		}
	}
	int deb = 1;
}


void PrintConfigDlg::OnBnClickedUserDefinedScript()
{
	// TODO: Add your control notification handler code here
	if (GetSafeHwnd())
	{
		CWnd *p = GetDlgItem(IDC_KEEP_BODY_BKGRND_COLOR);
		if (p) {
			p->EnableWindow(TRUE);
		}
		p = GetDlgItem(IDC_CHROME_EXE_PATH);
		if (p)
			p->EnableWindow(FALSE);
	}
	int deb = 1;
}


void PrintConfigDlg::OnBnClickedCustomTemplate()
{
	// TODO: Add your control notification handler code here

	if (GetSafeHwnd())
	{
		int checked = BST_UNCHECKED;
		CWnd *pRB = GetDlgItem(IDC_CUSTOM_TEMPLATE);
		if (pRB)
		{
			checked = ((CButton*)pRB)->GetCheck();
		}
		else
			return;

		// checked = m_NamePatternParams.m_bCustomFormat;  // shows old state
		CWnd *pB = GetDlgItem(IDC_SET_CUSTOM_TEMPLATE);
		if (pB)
		{
			if (checked)
			{
				EnableNonCustomWindows(FALSE);
				pB->EnableWindow(TRUE);
				pB->SetFocus();

				// SetFocus() doesn't seem to work and below doesn't seem to work
#if 0
				HWND hButtonBoxWnd = NULL;
				HWND hDlg = GetSafeHwnd();
				if((hButtonBoxWnd = ::GetDlgItem(hDlg, IDC_SET_CUSTOM_TEMPLATE)))
					::PostMessage(hButtonBoxWnd, WM_SETFOCUS, 0, 0);
#endif
			}
			else
			{
				EnableNonCustomWindows(TRUE);
				pB->EnableWindow(FALSE);
				pRB->SetFocus();
			}
		}
	}
	UpdateData(TRUE);
}


void PrintConfigDlg::OnBnClickedSetCustomTemplate()
{
	// TODO: Add your control notification handler code here

	CustomNameTemplDlg m_NameTemplDlg;

	m_NameTemplDlg.m_nameTemplateCnf.Copy(m_NamePatternParams.m_nameTemplateCnf);

	INT_PTR ret = m_NameTemplDlg.DoModal();
	if (ret == IDOK)
	{
		m_NamePatternParams.m_nameTemplateCnf.Copy(m_NameTemplDlg.m_nameTemplateCnf);
		m_NamePatternParams.AddressParts2Bitmap();
	}

	int deb = 1;
}


void PrintConfigDlg::EnableNonCustomWindows(BOOL enable)
{
	CWnd *p = 0;
	p = GetDlgItem(IDC_FILE_DATE);
	if (p)
		p->EnableWindow(enable);

	p = GetDlgItem(IDC_FILE_TIME);
	if (p)
		p->EnableWindow(enable);

	p = GetDlgItem(IDC_FILE_FROM);
	if (p)
		p->EnableWindow(enable);

	p = GetDlgItem(IDC_FILE_TO);
	if (p)
		p->EnableWindow(enable);

	p = GetDlgItem(IDC_FILE_SUBJECT);
	if (p)
		p->EnableWindow(enable);

	p = GetDlgItem(IDC_FILE_UNIQUE_ID);
	if (p)
		p->EnableWindow(enable);
}

void NamePatternParams::AddressParts2Bitmap()
{
	m_nAddressPartsBitmap = 0;

	if (m_nameTemplateCnf.m_bFromUsername)  m_nAddressPartsBitmap |= 1 << TemplFromUsername;
	if (m_nameTemplateCnf.m_bFromDomain)  m_nAddressPartsBitmap |= 1 << TemplFromDomain;

	if (m_nameTemplateCnf.m_bToUsername)  m_nAddressPartsBitmap |= 1 << TemplToUsername;
	if (m_nameTemplateCnf.m_bToDomain)  m_nAddressPartsBitmap |= 1 << TemplToDomain;

	if (m_nameTemplateCnf.m_bReplaceWhiteWithUnderscore)  m_nAddressPartsBitmap |= 1 << TemplReplaceWhiteWithUnderscore;
}

void NamePatternParams::Bitmap2AddressParts()
{
	m_nameTemplateCnf.ClearParts();


	if (m_nAddressPartsBitmap & (1 << TemplFromUsername))  m_nameTemplateCnf.m_bFromUsername = TRUE;
	if (m_nAddressPartsBitmap & (1 << TemplFromDomain))  m_nameTemplateCnf.m_bFromDomain = TRUE;

	if (m_nAddressPartsBitmap & (1 << TemplToUsername))  m_nameTemplateCnf.m_bToUsername = TRUE;
	if (m_nAddressPartsBitmap & (1 << TemplToDomain))  m_nameTemplateCnf.m_bToDomain = TRUE;

	if (m_nAddressPartsBitmap & (1 << TemplReplaceWhiteWithUnderscore))  m_nameTemplateCnf.m_bReplaceWhiteWithUnderscore = TRUE;
}


void PrintConfigDlg::OnBnClickedHtmlPdfCnf()
{
	// TODO: Add your control notification handler code here

	HtmlPdfHdrConfigDlg hdlg;
	INT_PTR ret = hdlg.DoModal();

	if (ret == IDOK)
	{
		CMainFrame *pFrame = (CMainFrame*)AfxGetMainWnd();
		if (pFrame)
		{
			pFrame->m_HdrFldConfig.LoadFromRegistry();
		}
		int deb = 1;
	}
}
