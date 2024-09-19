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

#if !defined(AFX_FIND_ADVANCED_DLG_H__139BBE54_571D_4C79_A9CD_B6035D054686__INCLUDED_)
#define AFX_FIND_ADVANCED_DLG_H__139BBE54_571D_4C79_A9CD_B6035D054686__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FindAdvancedDlg.h : header file
//

#include "stdafx.h"
#include "ATLComTime.h"
#include "MyCTime.h"


/////////////////////////////////////////////////////////////////////////////
// CFindAdvancedDlg dialog

#define FILTER_FIELDS_NUMB 8

struct CFindAdvancedParams
{
	CFindAdvancedParams::CFindAdvancedParams();
	CFindAdvancedParams::~CFindAdvancedParams();

	CString	m_string[FILTER_FIELDS_NUMB];
	CStringA m_stringA[FILTER_FIELDS_NUMB];
	UINT	m_charsetId[FILTER_FIELDS_NUMB];
	BOOL	m_bWholeWord[FILTER_FIELDS_NUMB];
	BOOL	m_bCaseSensitive[FILTER_FIELDS_NUMB];
	BOOL	m_bEditChecked[FILTER_FIELDS_NUMB];

	BOOL	m_bSetAllWholeWords;
	BOOL	m_bSetAllCaseSensitive;

	int m_plainText;
	int m_htmlText;
	int m_htmlTextOnlyIfNoPlainText;

	COleDateTime m_startDate;
	COleDateTime m_endDate;
	BOOL m_filterDates;

	// from FindFilterRuleDlg
	int m_filterNumb;
	int m_bSingleTo;
	BOOL m_bHighlightAll;  // higlight all occurences of search patern in Message Text
	BOOL m_bFindAll;
	BOOL m_bFindAllMailsThatDontMatch;  // find all mails that didn't match

	// Run time args
	// Following vars are not set or updated by FindDlg dialog
	// They are to help to maintain proper state
	BOOL m_bNeedToFindMailMinMaxTime;
	MyCTime m_mboxMailStartDate;
	MyCTime m_mboxMailEndDate;
	//
	BOOL m_needToRestoreArchiveListDateTime;
	MyCTime m_lastStartDate;
	MyCTime m_lastEndDate;
	CString m_dateTimeFormat;
	MyCTime m_archiveList_lastStartDate;
	MyCTime m_archiveList_lastEndDate;
	BOOL m_archiveList_filterDates;
	int m_last_nWhichMailList;
	// in case we decide to keep last range per list
	//MyCTime m_lastStartDateByMailList[3];
	//MyCTime m_lastEndDateByMailList[3];

	void SetDflts();
	void Copy(CFindAdvancedParams &src);
	void ResetFilterFields();

	void ResetFilterDates();
};

class CFindAdvancedDlg : public CDialog
{
// Construction
public:
	CFindAdvancedDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CFindAdvancedDlg)
	enum { IDD = IDD_FIND_ADVANCED };

	struct CFindAdvancedParams m_params;
	COLORREF m_dflBkColor;
	COLORREF m_checkedColor;
	CBrush m_brBkMailsDontMatch;
	CBrush m_brBkDate;
	//
	CToolTipCtrl m_toolTip;

	//}}AFX_DATA


	void SetRuleInfoText();
	BOOL InitDialogControls();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CFindAdvancedDlg)
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CFindAdvancedDlg)
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
public:
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd *pWnd, UINT nCtlColor);
	afx_msg void OnBnClickedFilterDates();
	afx_msg void OnBnClickedEditSetAllWhole();
	afx_msg void OnBnClickedEditSetAllCase();
	afx_msg void OnBnClickedEditReset();
	afx_msg void OnDtnDatetimechangeDatetimepicker1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnDtnDatetimechangeDatetimepicker2(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedButtonSelectFilterRule();
	afx_msg void OnBnClickedCheckNegateFindCriteria();
	afx_msg void OnBnClickedButtonHelp();
	afx_msg void OnBnClickedHtmlText();
	afx_msg void OnBnClickedHtmlIf();
	afx_msg BOOL OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FIND_ADVANCED_DLG_H__139BBE54_571D_4C79_A9CD_B6035D054686__INCLUDED_)
