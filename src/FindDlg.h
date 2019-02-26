#if !defined(AFX_FINDDLG_H__139BBE54_571D_4C79_A9CD_B6035D054686__INCLUDED_)
#define AFX_FINDDLG_H__139BBE54_571D_4C79_A9CD_B6035D054686__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FindDlg.h : header file
//

#include "ATLComTime.h"

/////////////////////////////////////////////////////////////////////////////
// CFindDlg dialog

struct CFindDlgParams
{
	CString	m_string;
	BOOL	m_bWholeWord;
	BOOL	m_bCaseSensitive;

	COleDateTime m_startDate;
	COleDateTime m_endDate;
	BOOL m_filterDates;

	BOOL m_bFindNext;
	BOOL m_bFrom;
	BOOL m_bTo;
	BOOL m_bSubject;
	BOOL m_bContent;
	BOOL m_bAttachments;
	BOOL m_bHighlightAll;
	BOOL m_bFindAll;

	void SetDflts();
	void Copy(CFindDlgParams &src);
};

class CFindDlg : public CDialog
{
// Construction
public:
	CFindDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CFindDlg)
	enum { IDD = IDD_FIND };

	struct CFindDlgParams m_params;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CFindDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CFindDlg)
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
public:
	//virtual BOOL OnInitDialog();

	afx_msg void OnBnClickedFilterDates();
	afx_msg BOOL OnInitDialog();
	afx_msg void OnBnClickedCheckFindAll();
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FINDDLG_H__139BBE54_571D_4C79_A9CD_B6035D054686__INCLUDED_)
