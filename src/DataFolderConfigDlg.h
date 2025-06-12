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

#pragma once


// DataFolderConfigDlg dialog

class DataFolderConfigDlg : public CDialogEx
{
	DECLARE_DYNAMIC(DataFolderConfigDlg)

public:
	DataFolderConfigDlg(BOOL restartRequired = FALSE, CWnd* pParent = nullptr);   // standard constructor
	virtual ~DataFolderConfigDlg();

	virtual INT_PTR DoModal();
	CWnd* m_pParent;

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DATA_FOLDER_DLG };
#endif

	void SetButtonColor(UINT nID, COLORREF color);

	int m_returnCode;
	BOOL m_restartRequired;
	//CEdit m_introText;
	CStatic m_introText;
	int m_fontSize;
	//
	CStatic m_defaultDataFolderText;
	//
	CString m_strCurrentDataFolder;
	CEdit m_currentDataFolder;
	//
	CString m_strUserConfiguredDataFolder;
	CEdit m_userConfiguredDataFolder;
	//
	CButton m_selectDataButton;
	//
	COLORREF m_folderPathColor;
	CBrush m_folderPathBrush;
	CFont m_BoldFont;
	CFont m_TextFont;
	CBrush m_ButtonBrush;
	//
	CToolTipCtrl m_toolTip;

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()

public:

	virtual BOOL OnInitDialog();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd *pWnd, UINT nCtlColor);
	afx_msg void OnBnClickedSelectDataFolderButton();
	afx_msg void OnBnClickedOk();
	afx_msg void OnEnChangeUserSelectedFolderPath();
	afx_msg void OnEnChangeCurrentFolderPath();
	afx_msg BOOL OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult);
};
