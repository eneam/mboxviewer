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

#if !defined(_EXPORT_TO_CVS_DLG_)
#define _EXPORT_TO_CVS_DLG_

#pragma once


// ExportToCSVDlg dialog

class ExportToCSVDlg : public CDialogEx
{
	DECLARE_DYNAMIC(ExportToCSVDlg)

public:
	ExportToCSVDlg(CWnd* pParent = 0);   // standard constructor
	virtual ~ExportToCSVDlg();

	virtual INT_PTR DoModal();
	CWnd* m_pParent;

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_EXPORT_TO_CSV };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();

	DECLARE_MESSAGE_MAP()
public:
	BOOL m_bFrom;
	BOOL m_bTo;
	BOOL m_bSubject;
	BOOL m_bDate;
	BOOL m_bCC;
	BOOL m_bBCC;
	BOOL m_bContent;
	BOOL m_bAttachmentNames;
	CString m_AttachmentNamesSeparatorString;
	CString m_MessageLimitString;
	CString m_MessageLimitCharsString;
	int m_dateFormat;
	int m_bGMTTime;
	int m_bEncodingType;
	int m_nCodePageId;
	//
	CToolTipCtrl m_toolTip;

	afx_msg void OnOK();
	afx_msg void OnBnClickedEncodingNone();
	afx_msg void OnBnClickedEncodingUtf8();
	afx_msg void OnBnClickedEncodingCodePage();
	afx_msg BOOL OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult);

};

#endif // _EXPORT_TO_CVS_DLG_
