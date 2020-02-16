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
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
//  Library General Public License for more details.
//
//  You should have received a copy of the GNU Library General Public
//  License along with this program; if not, write to the
//  Free Software Foundation, Inc., 51 Franklin St, Fifth Floor,
//  Boston, MA  02110 - 1301, USA.
//
//////////////////////////////////////////////////////////////////
//

#if !defined(_PRINT_CONFIG_DLG_)
#define _PRINT_CONFIG_DLG_

#pragma once

#include "CustomNameTemplDlg.h"


// PrintConfigDlg dialog

class PrintNameConfig
{
};

struct NamePatternParams
{
	int m_bDate;
	int m_bTime;
	int m_bFrom;
	int m_bTo;
	int m_bSubject;
	int m_bUniqueId;
	int m_nFileNameFormatSizeLimit;
	int m_nWantedFileNameFormatSizeLimit;
	int m_bPrintDialog;
	int m_bScriptType;
	CString	m_ChromeBrowserPath;
	CString	m_UserDefinedScriptPath;
	int m_bPrintToSeparatePDFFiles;
	int m_bPrintToSeparateHTMLFiles;
	int m_bPrintToSeparateTEXTFiles;
	int m_bAddBackgroundColorToMailHeader;
	int m_bAddBreakPageAfterEachMailInPDF;
	int m_bKeepMailBodyBackgroundColor;
	int m_bCustomFormat;

	DWORD m_nAddressPartsBitmap;

	struct NameTemplateCnf m_nameTemplateCnf;

	void SetDflts();
	void Copy(NamePatternParams &src);
	void UpdateRegistry(NamePatternParams &current, NamePatternParams &updated);
	void LoadFromRegistry();
	void AddressParts2Bitmap();
	void Bitmap2AddressParts();

	static void UpdateFilePrintconfig(struct NamePatternParams &namePatternParams);

	enum {
		TemplFromUsername = 1,
		TemplFromDomain = 2,
		TemplToUsername = 3,
		TemplToDomain = 4,
		TemplReplaceWhiteWithUnderscore = 5
	};
};

class PrintConfigDlg : public CDialogEx
{
	DECLARE_DYNAMIC(PrintConfigDlg)

public:
	PrintConfigDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~PrintConfigDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_PRINT_DLG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();

	void EnableNonCustomWindows(BOOL enable);
public:
	struct NamePatternParams m_NamePatternParams;

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	afx_msg void OnEnChangeFileNameMaxSize();
	afx_msg void OnBnClickedPrtPageSetp();
	afx_msg void OnBnClickedHtml2pdfScriptType();
	afx_msg void OnBnClickedRadio3();

	afx_msg void OnBnClickedCustomTemplate();
	afx_msg void OnBnClickedSetCustomTemplate();
};

#endif _PRINT_CONFIG_DLG_
