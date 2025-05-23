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

// CustomNameTemplDlg dialog


struct NameTemplateCnf
{
	CString	m_TemplateFormat;
	CString	m_DateFormat;
	int m_bFromUsername;
	int m_bFromDomain;
	int m_bToUsername;
	int m_bToDomain;
	int m_bReplaceWhiteWithUnderscore;

	void ClearParts();
	void SetDflts();
	void Copy(NameTemplateCnf &src);
};

class CustomNameTemplDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CustomNameTemplDlg)

public:
	CustomNameTemplDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~CustomNameTemplDlg();

	virtual INT_PTR DoModal();
	CWnd* m_pParent;

	struct NameTemplateCnf m_nameTemplateCnf;
	//
	CToolTipCtrl m_toolTip;

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_NAME_TEMPL_DLG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedSrcftime();
	virtual BOOL OnInitDialog();
	afx_msg BOOL OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult);
};
