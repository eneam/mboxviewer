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

#pragma once
#include "afxdialogex.h"


// CustomMsgBox dialog

class MB_BUTTON
{
public:
	MB_BUTTON()
	{
		m_id = 0;   
		m_name = L"";
		m_action = 0;
		m_hidden = TRUE;
	}
	void Set(WORD id, CString name, UINT icon, UINT action, BOOL hidden)
	{
		m_id = id;
		m_name = name;
		m_action = action;
		m_hidden = hidden;
	}

	void Set(CButton& button, WORD id, int actionType, CString name, BOOL hidden = FALSE);

	WORD m_id;  // // ID_MSG_BOX_BUTTON_(1-3)
	CString m_name;  // "OK", "NO", etc

	UINT m_action;  // MB_OK, MB_NO, MB_YES, MB_CANCEL, MB_RETRY, etc
	BOOL m_hidden;
	int m_actionType; // IDOK, IDNO, IDYES, IDCANCEL, IDRETRY, etc
};

class CustomMsgBox : public CDialogEx
{
	DECLARE_DYNAMIC(CustomMsgBox)

public:
	CustomMsgBox(CWnd* pParent = nullptr);   // standard constructor

	CustomMsgBox(LPCTSTR lpszText, LPCTSTR lpszCaption, UINT nType, int textFontHeight, CWnd* pParent = nullptr);

	virtual ~CustomMsgBox();

	CString m_textStr;
	CString captionStr;
	int m_textFontHeight;
	UINT m_nType;
	UINT m_iconType;  // MB_ICONERROR, MB_ICONQUESTION, MB_ICONWARNING, MB_ICONHAND

	void ProcessType(UINT nType);

	MB_BUTTON m_buttons[3];
	MB_BUTTON* FindButton(WORD id);
	CButton* FindDfltCButton();
	BOOL IsDfltCButton(CButton& button);
	void SetDefaultButton(CButton& newDflt, WORD id, CDialogEx *dlg);
	LPWSTR GetIconId(UINT nType);

	CButton m_button0;
	CButton m_button1;
	CButton m_button2;
	CEdit m_text;
	int m_sctionCode;

	CStatusBar	m_wndStatusBar;  // just to implement Gripper
	int m_StatusBarHeight;

	HICON	m_hIcon;  // standard MessageBox win system icon
	CStatic m_icon;
	//

	// timer
	UINT_PTR m_nIDEvent;
	UINT m_nElapse;

	void StartTimer();

	HICON LoadMsgBoxIcon(BOOL systemIcon, PCWSTR pszName, int lims, CString &errorText);
	int LongestWordLength(CString& str, CString* longetWord = 0);
	int LongestLineLength(CString& str, CString* longetLine = 0);

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_CUSTOM_MSG_BOX };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	BOOL SetClipboardText(CString& buffer);

	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg LRESULT OnCmdParam_OnSetRTLForButtons(WPARAM wParam, LPARAM lParam);
	afx_msg void OnTimer(UINT_PTR nIDEvent);

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedMsgBoxButton1();
	afx_msg void OnBnClickedMsgBoxButton2();
	afx_msg void OnBnClickedMsgBoxButton3();
};

void TestCustomMsgBox();
