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
#include "afxdialogex.h"


// SelectLanguageDlg dialog

class DerivedCListBox : public CListBox
{
public:
	DerivedCListBox();
	~DerivedCListBox() {; }

	DECLARE_MESSAGE_MAP()
public:
	void SetBkColor(COLORREF crColor); // This Function is to set the BackGround Color for the Text and the Edit Box.
	void SetTextColor(COLORREF crColor); // This Function is to set the Color for the Text. TODO: doesn't work

	CBrush m_brBkgnd; // Holds Brush Color for the Edit Box
	COLORREF m_crBkColor; // Holds the Background Color for the Text
	COLORREF m_crTextColor; // Holds the Color for the Text

	afx_msg HBRUSH CtlColor(CDC* pDC, UINT nCtlColor); // This Function Gets Called Every Time Your Window Gets Redrawn.
};

class SelectLanguageDlg : public CDialogEx
{
	DECLARE_DYNAMIC(SelectLanguageDlg)

public:
	SelectLanguageDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~SelectLanguageDlg();

	CString m_lastFolderName;
	DerivedCListBox m_listBox;
	//CListBox m_listBox;
	CString m_LanguagetName;  // font family such as 

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_LANG_LIST_DLG };
#endif

protected:
	void LoadLanguages();

	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedStaticSelectLang1();
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
};
