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


// GeneralOptionsDlg dialog

class GeneralOptionsDlg : public CDialogEx
{
	DECLARE_DYNAMIC(GeneralOptionsDlg)

public:
	GeneralOptionsDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~GeneralOptionsDlg();

	virtual INT_PTR DoModal();
	CWnd* m_pParent;

	BOOL m_relaxedMboxFileValidation;
	BOOL m_relativeInlineImageFilePath;
	// ODF Merge Config
	BOOL m_enableUserAgent;
	CString m_numberOfHTML2ToMerge;

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_GENERAL_OPTIONS_DLG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//
	CToolTipCtrl m_toolTip;

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedRelaxMboxFileValidation();
	afx_msg void OnBnClickedRelativeAttachmentFilePath();

	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
	afx_msg BOOL OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnEnChangeMergeHtmlCnt();
	afx_msg void OnBnClickedMergeUserAgent();
};
