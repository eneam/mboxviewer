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
#include "AttachmentsConfig.h"
#include "ColorStyleConfigDlg.h"
#include "HtmlPdfHdrConfigDlg.h"
#include "SMTPMailServerConfigDlg.h"


class MBoxViewerDB
{
public:
	MBoxViewerDB();

	CString GetDBFolder();
	CString GetLocalDBFolder();
	CString GetPrintSubFolder();
	CString GetImageSubFolder();
	CString GetAttachmentSubFolder();
	CString GetArchiveSubFolder();
	CString GetLabelSubFolder();
	CString GetEmlSubFolder();
	CString GetMergedSubFolder();
	BOOL IsReadOnlyMboxDataMedia(CString *path = 0);

	CString m_rootDBFolder;
	CString m_rootLocalDBFolder;
	CString m_rootPrintSubFolder;
	CString m_rootImageSubFolder;
	CString m_rootAttachmentSubFolder;
	CString m_rootArchiveSubFolder;
	CString m_rootListSubFolder;
	CString m_rootLabelSubFolder;
	CString m_rootEmlSubFolder;
	CString m_rootMergedSubFolder;
	int m_isReadOnlyMboxDataMedia;

};

typedef CArray<CString, CString> MboxFileList;

struct CommandLineParms
{
	CommandLineParms() 
	{
		m_bEmlPreviewMode = FALSE;  m_progressBarDelay = -1; 
		m_exportEml = FALSE; m_traceCase = 0; m_bEmlPreviewFolderExisted = FALSE;
		m_hasOptions = FALSE; m_bDirectFileOpenMode = FALSE;
	}
	void Clear()
	{
		m_bEmlPreviewMode = FALSE;  m_bEmlPreviewFolderExisted = FALSE;
		m_progressBarDelay = -1; m_exportEml = FALSE;
		m_mboxListFilePath.Empty(); m_mergeToFilePath.Empty();
		m_mboxFolderPath.Empty(); m_mboxFileNameOrPath.Empty();

	}
	int VerifyParameters();
	//
	BOOL m_hasOptions;
	CString m_allCommanLineOptions;
	//
	CString m_mboxListFilePath;
	CString m_mergeToFilePath;
	// params to open single eml or mbox file
	BOOL m_bEmlPreviewMode;
	CString m_mboxFolderPath;
	CString m_mboxFileNameOrPath;
	BOOL m_bEmlPreviewFolderExisted;
	// file name is the only command line param to open mail file directly by double left click
	BOOL m_bDirectFileOpenMode;
	//
	int m_progressBarDelay;
	BOOL m_exportEml;
	//
	int m_traceCase;
};

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
	BOOL m_bAttachmentNames;
	CStringA m_AttachmentNamesSeparatorString;
	CString m_MessageLimitString;
	CString m_MessageLimitCharsString;
	int m_dateFormat;
	int m_bGMTTime;
	int m_nCodePageId;
	CStringA m_separator;

	void SaveToRegistry();
	void LoadFromRegistry();
};

class MySelectFolder : public CFileDialog
{
public:
	MySelectFolder(
		BOOL bOpenFileDialog, // TRUE for FileOpen, FALSE for FileSaveAs
		LPCWSTR lpszDefExt = NULL,
		LPCWSTR lpszFileName = NULL,
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		LPCWSTR lpszFilter = NULL,
		CWnd* pParentWnd = NULL,
		DWORD dwSize = 0,
		BOOL bVistaStyle = TRUE
	);

	virtual BOOL OnFileNameOK();
};

class MergeFileInfo
{
public:
	MergeFileInfo()
	{
		mboxFileType = -1;
	}
	MergeFileInfo(CString &filePath, int mboxFileType)
	{
		m_filepath = filePath;
		mboxFileType = mboxFileType;
	}
	CString m_filepath;
	int mboxFileType;
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

	static CString m_processPath;
	static CString m_startupPath;
	static CString m_currentPath;

	static CString GetMboxviewTempPath(const wchar_t* name = 0);
	static CString m_mboxviewTempPath;
	static CString GetMboxviewLocalAppDataPath(const wchar_t* name = 0);
	static CString CreateTempFileName(const wchar_t *ext = L"htm");

	void UpdateToolsBar();
	BOOL IsTreeHidden();

	int MsgViewPosition() {
		return m_msgViewPosition;
	}

	int TreeHideValue() {
		return m_bIsTreeHidden;
	}

	BOOL DeleteAllPlacementKeys(CString& section_wnd);

	BOOL DoOpen(CString& path);
	void SetMailList(int nID);
	void EnableMailList(int nId, BOOL enable);
	void EnableAllMailLists(BOOL enable);
	void SetupMailListsToInitialState();
	BOOL IsUserMailsListEnabled();
	void ConfigMessagewindowPosition(int msgViewPosition);
	int GetMessageWindowPosition() { return m_msgViewPosition; }
	void CheckMessagewindowPositionMenuOption(int msgViewPosition);
	void UpdateFilePrintconfig();
	int MergeArchiveFiles();
	BOOL SaveFileDialog(CString &fileName, CString &fileNameFilter, CString &dfltExtention, CString &inFolderPath, CString &outFolderPath, CString &title);
	void NMCustomdrawEditList(NMHDR *pNMHDR, LRESULT *pResult);
	//
	int OnPrintSingleMailtoText(int mailPosition, int textType, CString &createdTextFileName, BOOL forceOpen = FALSE, BOOL printToPrinter = FALSE, BOOL createFileOnly = FALSE);
	void OnPrinttoTextFile(int textType);
	int PrintSingleMailtoPDF(int iItem, CString &targetPrintSubFolderName, BOOL progressBar, CString& progressText, CString &errorText);
	int PrintSingleMailtoHTML(int iItem, CString &targetPrintSubFolderName, CString &errorText);
	//
	// Used by Print to
	static int ExecCommand_WorkerThread(CString& htmFileName, CString& errorText, BOOL progressBar, CString& progressText, int timeout, int headlessTimout);
	int VerifyPathToHTML2PDFExecutable(CString &errorText);


	// Used by Merge PDFs
	static int ExecCommand_WorkerThread(CString &directory, CString &cmd, CString &args, CString &errorText, BOOL progressBar, CString &progressText, int timeout);

	void PrintMailArchiveToHTML();
	void PrintMailArchiveToTEXT();
	void PrintMailArchiveToPDF();
	void OnPrinttoPdf();

	void PrintMailsToCSV(int firstMail, int lastMail, BOOL selecteMails);

	static int CheckShellExecuteResult(HINSTANCE  result, HWND h, CStringA *filePath = 0);
	static int CheckShellExecuteResult(HINSTANCE  result, CString &errorText);
	static int CheckShellExecuteResult(HINSTANCE  result, HWND h, CStringW *filename);

	static INT_PTR SelectFolder(CString &folder);
	static int CountMailFilesInFolder(CString &folder, CString &extension);

	void SetStatusBarPaneText(int paneId, CString &sText, BOOL setColor);
	void SortByColumn(int column);


	int MergeMboxArchiveFiles(CArray<MergeFileInfo> &mboxFilePathList, CString &mergedMboxFilePath);
	int MergeMboxArchiveFiles(CString &mboxListFilePath, CString &mergedMboxFilePath);
	static int MergeMboxArchiveFile(CFile &fpMergeTo, CString & mboxFilePath, BOOL firstFile);

	static BOOL CanMboxBeSavedInFolder(CString &destinationFolder);

	static MBoxViewerDB m_ViewerGlobalDB;
	static CommandLineParms m_commandLineParms;
	static AttachmentConfigParams *GetAttachmentConfigParams();

	static void OpenHelpFile(CString &filePath, HWND h = NULL);

	static BOOL m_relaxedMboxFileValidation;
	static BOOL m_relativeInlineImageFilePath;

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
	//CComboBox m_wndSearchCombo;
	BOOL m_bSelectMailFileDone;
	HICON m_hIcon;
	BOOL m_bMailDownloadComplete;
	int m_MailIndex;
	//BOOL m_bMailListType;
	BOOL m_bUserSelectedMailsCheckSet;
	int m_msgViewPosition;
	int  m_newMsgViewPosition;
	NListView *m_pListView;
	NTreeView *m_pTreeView;
	NMsgView *m_pMsgView;
	CSVFILE_CONFIG m_csvConfig;
	int m_bEnhancedSelectFolderDlg;
	CImageList m_imgListBag;
	BOOL m_bTreeExpanded;

	int treeColWidth;
	HICON m_PlusIcon;
	HICON m_MinusIcon;
	HICON m_HideIcon;
	HICON m_UnHideIcon;

public:
	// From MergeFolderAndSubfolders
	int m_mergeRootFolderStyle;
	int m_labelAssignmentStyle;  // 0 == no labels; 1 == mbox file names as labels; 2 == folder names as labels

	CString m_lastPath;
	BOOL m_bIsTreeHidden;
	HdrFldConfig m_HdrFldConfig;
	NamePatternParams m_NamePatternParams;
	AttachmentConfigParams m_attachmentConfigParams;
	ColorStyleConfigDlg *m_colorStyleDlg;
	BOOL m_bViewMessageHeaders;
	MailDB m_mailDB;

	BOOL CreateMailDbFile(MailDB &m_mailDB, CString &fileName);
	BOOL WriteMTPServerConfig(MailConfig &serverConfig, CFile &fp);

	static ColorStylesDB m_ColorStylesDB;

	void OpenFolderAndSubfolders(CString &path);
	void OpenRootFolderAndSubfolders(CString &path);

	void OpenRootFolderAndSubfolders_LabelView(CString &path);
	//
	int FileSelectrootfolder(int treeType);


	//BOOL PreTranslateMessage(MSG* pMsg);

// Generated message map functions
protected:
	//{{AFX_MSG(CMainFrame)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSetFocus(CWnd *pOldWnd);
	afx_msg void OnUpdateFileOpen(CCmdUI* pCmdUI);
	afx_msg void OnFileOpen();
	afx_msg void OnFileOptions();
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnTreeHide();
	afx_msg void OnActivateApp(BOOL bActive, DWORD dwThreadID);
	//afx_msg void OnKeydown(NMHDR* pNMHDR, LRESULT* pResult);
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
	afx_msg void OnClose();
	//afx_msg void OnBnClickedFolderList();
	afx_msg void OnFileAttachmentsconfig();
	afx_msg void OnFileColorconfig();
	afx_msg LRESULT OnCmdParam_ColorChanged(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmdParam_LoadFolders(WPARAM wParam, LPARAM lParam);
	afx_msg void OnFordevelopersSortbyid();
	afx_msg void OnFordevelopersMemory();
	afx_msg void OnViewMessageheaders();
	afx_msg void OnFileRestorehintmessages();
	afx_msg void OnMessageheaderpanelayoutDefault();
	afx_msg void OnMessageheaderpanelayoutExpanded();
	afx_msg void OnFileSmtpmailserverconfig();
	afx_msg void OnHelpUserguide();
	afx_msg void OnHelpReadme();
	afx_msg void OnHelpLicense();
	afx_msg void OnFileDatafolderconfig();
	afx_msg void OnFileMergerootfoldersub();
	afx_msg void OnFileSelectasrootfolder();
	afx_msg void OnDevelopmentoptionsDumprawdata();
	afx_msg void OnDevelopmentoptionsDevelo();
	afx_msg void OnDeveloperOptionsAboutSystem();
	afx_msg void OnFileGeneraloptions();
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MAINFRM_H__342D6194_C2A0_47B2_87E0_C7651BE1D8DF__INCLUDED_)
