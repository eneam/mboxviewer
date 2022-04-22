#pragma once


// DataFolderConfigDlg dialog

#define DFLT_DATA_FOLDER 1
#define WINDOWS_APP_DATA_FOLDER 2
#define USER_DEFINED_DATA_FOLDER 3

class DataFolderConfigDlg : public CDialogEx
{
	DECLARE_DYNAMIC(DataFolderConfigDlg)

public:
	DataFolderConfigDlg(BOOL restartRequired = FALSE, CWnd* pParent = nullptr);   // standard constructor
	virtual ~DataFolderConfigDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DATA_FOLDER_DLG };
#endif

	void SetDlgItemState(BOOL setDataFolderButton, BOOL setSelectDataFolder, BOOL setWinAppDataFolder);
	void SetButtonColor(UINT nID, COLORREF color);

	BOOL m_restartRequired;
	CEdit m_introText;
	CString m_dataFolder;
	CString m_strWinAppDataFolder;
	CString m_strUserConfiguredDataFolder;
	CEdit m_winAppDataFolder;
	CEdit m_userConfiguredDataFolder;
	CButton m_selectDataButton;
	//
	int m_selectedDataFolderConfigStyle;
	//
	int m_currentDataFolderConfigStyle;
	CString m_strCurrentUserConfiguredDataFolder;
	BOOL m_alreadyConfiguredByUser;
	//
	COLORREF m_folderPathColor;
	CBrush m_folderPathBrush;
	CFont m_BoldFont;
	CFont m_TextFont;
	CBrush m_ButtonBrush;

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()

public:

	virtual BOOL OnInitDialog();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd *pWnd, UINT nCtlColor);
	afx_msg void OnBnClickedSelectDataFolderButton();
	afx_msg void OnBnClickedConfigFolderStyle(UINT nID);
	afx_msg void OnBnClickedOk();
	afx_msg void OnEnChangeEditWinAppDataFolder();
	afx_msg void OnEnChangeUserSelectedFolderPath();
};
