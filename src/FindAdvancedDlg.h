
#if !defined(AFX_FIND_ADVANCED_DLG_H__139BBE54_571D_4C79_A9CD_B6035D054686__INCLUDED_)
#define AFX_FIND_ADVANCED_DLG_H__139BBE54_571D_4C79_A9CD_B6035D054686__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FindAdvancedDlg.h : header file
//

#include "stdafx.h"
#include "ATLComTime.h"

/////////////////////////////////////////////////////////////////////////////
// CFindAdvancedDlg dialog

#define FILTER_FIELDS_NUMB 8

struct CFindAdvancedParams
{
	CString	m_string[FILTER_FIELDS_NUMB];
	BOOL	m_bWholeWord[FILTER_FIELDS_NUMB];
	BOOL	m_bCaseSensitive[FILTER_FIELDS_NUMB];
	BOOL	m_bEditChecked[FILTER_FIELDS_NUMB];

	BOOL	m_bSetAllWholeWords;
	BOOL	m_bSetAllCaseSensitive;

	COleDateTime m_startDate;
	COleDateTime m_endDate;
	BOOL m_filterDates;

	int m_filterNumb;
	BOOL m_bFindAll;

	BOOL m_bBiderectionalMatch;
	BOOL m_bFindAllMailsThatDontMatch;  // find all mails that didn't match

	void SetDflts();
	void Copy(CFindAdvancedParams &src);
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

	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CFindAdvancedDlg)
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//virtual BOOL OnInitDialog();
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CFindAdvancedDlg)
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
public:

	afx_msg void OnBnClickedFilterDates();
	afx_msg BOOL OnInitDialog();
	afx_msg void OnBnClickedEditSetAllWhole();
	afx_msg void OnBnClickedEditSetAllCase();
	afx_msg void OnBnClickedEditReset();
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FIND_ADVANCED_DLG_H__139BBE54_571D_4C79_A9CD_B6035D054686__INCLUDED_)
