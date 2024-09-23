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


// ForwardMailDlg dialog

class MyCustomEdit : public CEdit
{
public:
	MyCustomEdit();
	void SetSubject(CStringW &subjectW) {
		m_subjectW = subjectW;
	};
	CStringW m_subjectW;

protected:
	DECLARE_MESSAGE_MAP()
	afx_msg void OnPaint();
};

class ForwardMailData
{
public:
	CString m_MailService;
	CString m_From;
	CString m_To;
	CString m_CC;
	CString m_BCC;;
	CString m_Text;
	CStringW m_subjectW;

	void Copy(ForwardMailData &data);
};

class ForwardMailDlg : public CDialogEx
{
	DECLARE_DYNAMIC(ForwardMailDlg)

public:
	ForwardMailDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~ForwardMailDlg();

	ForwardMailData m_Data;
	CEdit m_EditText;
	MyCustomEdit m_Subject;
	//MyCustomEdit m_From;
		//
	CToolTipCtrl m_toolTip;

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_FORWARD_EMAIL };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnBnClickedForwardClear();
	afx_msg BOOL OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult);
};
