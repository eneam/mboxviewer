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

#if !defined(_OPTIONS_DLG_)
#define _OPTIONS_DLG_

#pragma once


// finestra di dialogo COptionsDlg

class COptionsDlg : public CDialog
{
	DECLARE_DYNAMIC(COptionsDlg)

public:
	COptionsDlg(CWnd* pParent = NULL);   // costruttore standard
	virtual ~COptionsDlg();

// Dati della finestra di dialogo
	enum { IDD = IDD_OPTIONS };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // Supporto DDX/DDV

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	virtual BOOL OnInitDialog();
	int m_format;
	int m_exportEML;
	int	m_barDelay;
	int m_from_charsetId;
	int m_to_charsetId;
	int m_subj_charsetId;
	int m_show_charsets;
	int m_bImageViewer;
	int m_bTimeType;
	int m_bEnhancedSelectFolderDlg;
	int m_bSubjectSortType;
	afx_msg void OnBnClickedRadio2();
};

#endif // _OPTIONS_DLG_
