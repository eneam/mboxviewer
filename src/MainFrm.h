// MainFrm.h : interface of the CMainFrame class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_MAINFRM_H__342D6194_C2A0_47B2_87E0_C7651BE1D8DF__INCLUDED_)
#define AFX_MAINFRM_H__342D6194_C2A0_47B2_87E0_C7651BE1D8DF__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "ChildView.h"

class CMainFrame : public CFrameWnd
{
	
public:
	CMainFrame();
protected: 
	DECLARE_DYNAMIC(CMainFrame)

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMainFrame)
	public:
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	virtual BOOL OnCmdMsg(UINT nID, int nCode, void* pExtra, AFX_CMDHANDLERINFO* pHandlerInfo);
	//virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
public:
	NMsgView * GetMsgView();
	NListView * GetListView();
	NTreeView * GetTreeView();
	virtual ~CMainFrame();
	void DoOpen(CString& path);
	void SetMailList(int nID);
	void EnableMailList(int nId, BOOL enable);
	void EnableAllMailLists(BOOL enable);
	void SetupMailListsToInitialState();
	BOOL IsUserMailsListEnabled();

	void CreateMailListsInfoText(CFile &fp);
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:  // control bar embedded members
	CStatusBar	m_wndStatusBar;
	CToolBar	m_wndToolBar;
	CReBar      m_wndReBar;
	CDialogBar    m_wndDlgBar;
	CChildView    m_wndView;
	CComboBox m_wndSearchCombo;
	BOOL m_bSelectMailFileDone;
	HICON m_hIcon;
	BOOL m_bMailDownloadComplete;
	int m_MailIndex;
	//BOOL m_bMailListType;
	BOOL m_bUserSelectedMailsCheckSet;

// Generated message map functions
protected:
	//{{AFX_MSG(CMainFrame)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSetFocus(CWnd *pOldWnd);
	afx_msg void OnUpdateFileOpen(CCmdUI* pCmdUI);
	afx_msg void OnFileOpen();
	afx_msg void OnFileOptions();
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	void NMCustomdrawEditList(NMHDR *pNMHDR, LRESULT *pResult);


public:
	void OnPrintSingleMailtoText(int mailPosition, int textType, BOOL forceOpen = FALSE, BOOL printToPrinter = FALSE);
	void OnPrinttoTextFile(int textType);

	afx_msg void OnFileExportToCsv();
	afx_msg void OnViewCodepageids();
	afx_msg void OnPrinttoCsv();
	afx_msg void OnPrinttoText();
	afx_msg void OnPrinttoHtml();

	afx_msg void OnBydate();
	afx_msg void OnUpdateBydate(CCmdUI *pCmdUI);
	afx_msg void OnByfrom();
	afx_msg void OnUpdateByfrom(CCmdUI *pCmdUI);
	afx_msg void OnByto();
	afx_msg void OnUpdateByto(CCmdUI *pCmdUI);
	afx_msg void OnBysubject();
	afx_msg void OnUpdateBysubject(CCmdUI *pCmdUI);
	afx_msg void OnBysize();
	afx_msg void OnUpdateBysize(CCmdUI *pCmdUI);
	afx_msg void OnByconversation();
	afx_msg void OnUpdateByconversation(CCmdUI *pCmdUI);
	afx_msg void OnUpdateMailDownloadStatus(CCmdUI *pCmdUI);
	afx_msg void OnUpdateMailIndex(CCmdUI *pCmdUI);
	afx_msg void OnBnClickedArchiveList();
	afx_msg void OnBnClickedFindList();
	afx_msg void OnBnClickedEditList();
	afx_msg void OnNMCustomdrawArchiveList(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMCustomdrawFindList(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMCustomdrawEditList(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedButton2();
	afx_msg void OnNMCustomdrawButton2(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnViewUserselectedmails();
	afx_msg void OnUpdateViewUserselectedmails(CCmdUI *pCmdUI);
	afx_msg void OnHelpMboxviewhelp();
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MAINFRM_H__342D6194_C2A0_47B2_87E0_C7651BE1D8DF__INCLUDED_)
