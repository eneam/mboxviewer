// MainFrm.h : interface of the CMainFrame class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_MAINFRM_H__342D6194_C2A0_47B2_87E0_C7651BE1D8DF__INCLUDED_)
#define AFX_MAINFRM_H__342D6194_C2A0_47B2_87E0_C7651BE1D8DF__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "ChildView.h"
#include "PrintConfigDlg.h"

struct CSVFILE_CONFIG
{
public:
	void Copy(CSVFILE_CONFIG &src);
	void SetDflts();

	BOOL m_bFrom;
	BOOL m_bTo;
	BOOL m_bSubject;
	BOOL m_bDate;
	BOOL m_bCC;
	BOOL m_bBCC;
	BOOL m_bContent;
	CString m_MessageLimitString;
	CString m_MessageLimitCharsString;
	int m_dateFormat;
	int m_bGMTTime;
	int m_nCodePageId;
	CString m_separator;
};

class MySelectFolder : public CFileDialog
{
public:
	MySelectFolder(
		BOOL bOpenFileDialog, // TRUE for FileOpen, FALSE for FileSaveAs
		LPCTSTR lpszDefExt = NULL,
		LPCTSTR lpszFileName = NULL,
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		LPCTSTR lpszFilter = NULL,
		CWnd* pParentWnd = NULL,
		DWORD dwSize = 0,
		BOOL bVistaStyle = TRUE
	);

	virtual BOOL OnFileNameOK();
};

class CMainFrame : public CFrameWnd
{
	
public:
	CMainFrame(int msgViewPosition = 1);
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
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CMainFrame();

	NMsgView * GetMsgView();
	NListView * GetListView();
	NTreeView * GetTreeView();
	NMsgView * DetMsgView();
	NListView * DetListView();
	NTreeView * DetTreeView();

	void DoOpen(CString& path);
	void SetMailList(int nID);
	void EnableMailList(int nId, BOOL enable);
	void EnableAllMailLists(BOOL enable);
	void SetupMailListsToInitialState();
	BOOL IsUserMailsListEnabled();
	void ConfigMessagewindowPosition(int msgViewPosition);
	int GetMessageWindowPosition() { return m_msgViewPosition; }
	void CreateMailListsInfoText(CFile &fp);
	void UpdateFilePrintconfig();
	int MergeArchiveFiles();
	BOOL SaveFileDialog(CString &fileName, CString &fileNameFilter, CString &dfltExtention, CString &inFolderPath, CString &outFolderPath, CString &title);
	void NMCustomdrawEditList(NMHDR *pNMHDR, LRESULT *pResult);
	//
	int OnPrintSingleMailtoText(int mailPosition, int textType, CString &createdTextFileName, BOOL forceOpen = FALSE, BOOL printToPrinter = FALSE, BOOL createFileOnly = FALSE);
	void OnPrinttoTextFile(int textType);
	int PrintSingleMailtoPDF(int iItem, CString &targetPrintSubFolderName, BOOL progressBar, CString &errorText);
	int PrintSingleMailtoHTML(int iItem, CString &targetPrintSubFolderName, CString &errorText);
	//
	static int ExecCommand_WorkerThread(CString &htmFileName, CString &errorText, BOOL progressBar, CString &progressText);
	int VerifyPathToHTML2PDFExecutable(CString &errorText);

	void PrintMailArchiveToHTML();
	void PrintMailArchiveToTEXT();
	void PrintMailArchiveToPDF();

	void PrintMailsToCSV(int firstMail, int lastMail, BOOL selecteMails);

	static int CheckShellExecuteResult(HINSTANCE  result, HWND h);
	static int CheckShellExecuteResult(HINSTANCE  result, CString &errorText);

	INT_PTR SelectFolder(CString &folder);
	static int CountMailFilesInFolder(CString &folder, CString &extension);

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
	int m_msgViewPosition;
	NListView *m_pListView;
	NTreeView *m_pTreeView;
	NMsgView *m_pMsgView;
	CSVFILE_CONFIG m_csvConfig;
	int m_bEnhancedSelectFolderDlg;

public:
	struct NamePatternParams m_NamePatternParams;

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


public:

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
	afx_msg void OnMessagewindowBottom();
	afx_msg void OnMessagewindowRight();
	afx_msg void OnMessagewindowLeft();
	afx_msg void OnFilePrintconfig();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnUpdateFilePrintconfig(CCmdUI *pCmdUI);
	afx_msg void OnFileMergearchivefiles();
	afx_msg void OnPrinttoPdf();
	afx_msg void OnClose();
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MAINFRM_H__342D6194_C2A0_47B2_87E0_C7651BE1D8DF__INCLUDED_)
