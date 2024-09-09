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

// OptionsDlg.cpp : file di implementazione
//

#include "stdafx.h"
#include "mboxview.h"
#include "OptionsDlg.h"
#include "ResHelper.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif


// finestra di dialogo COptionsDlg

IMPLEMENT_DYNAMIC(COptionsDlg, CDialog)

COptionsDlg::COptionsDlg(CWnd* pParent /*=NULL*/)
	: CDialog(COptionsDlg::IDD, pParent)
	, m_format(0), m_barDelay(0)
	, m_from_charsetId(0), m_to_charsetId(0), m_subj_charsetId(0), m_show_charsets(0)
	, m_bImageViewer(1), m_bTimeType(0), m_bEnhancedSelectFolderDlg(0), m_bSubjectSortType(0)
	, m_filesToValidateAsMboxType(0)
{
	//m_from_charsetId = GetACP();
	//m_to_charsetId = GetACP();
	//m_subj_charsetId = GetACP();
}

COptionsDlg::~COptionsDlg()
{
}

void COptionsDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Radio(pDX, IDC_DMY, m_format);
	DDX_Radio(pDX, IDC_EXPORT_EML, m_exportEML);
	DDX_Text(pDX, IDC_PROGRESS_BAR_DELAY, m_barDelay);
	DDX_Text(pDX, IDC_FROM_ID, m_from_charsetId);
	DDX_Text(pDX, IDC_TO_ID, m_to_charsetId);
	DDX_Text(pDX, IDC_SUBJ_ID, m_subj_charsetId);
	DDX_Radio(pDX, IDC_SHOW_CHARSETS, m_show_charsets);
	DDX_Check(pDX, IDC_IMAGE_VIEWER, m_bImageViewer);
	DDX_Check(pDX, IDC_ENHANCED_SELECT_FOLDER, m_bEnhancedSelectFolderDlg);
	DDX_Radio(pDX, IDC_TIME_LOCAL, m_bTimeType);
	DDX_Radio(pDX, IDC_SUBJECT_SORT_TYPE, m_bSubjectSortType);
	DDX_Radio(pDX, IDC_VALIDATE_MBOX_FILES, m_filesToValidateAsMboxType);
}


BEGIN_MESSAGE_MAP(COptionsDlg, CDialog)
	ON_BN_CLICKED(IDOK, &COptionsDlg::OnBnClickedOk)
	//ON_BN_CLICKED(IDC_RADIO2, &COptionsDlg::OnBnClickedRadio2)
	//ON_BN_CLICKED(IDC_PICTURE_VIEWER, &COptionsDlg::OnBnClickedPictureViewer)
END_MESSAGE_MAP()


// gestori di messaggi COptionsDlg


void COptionsDlg::OnBnClickedOk()
{
	CString section_options = CString(sz_Software_mboxview) + L"\\Options";
	CString section_options_set_charset_ids = CString(sz_Software_mboxview) + L"\\Options\\SetCharacterSetIds";
	if (UpdateData(TRUE))
	{
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"dayMonthYearFormat", m_format);
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"progressBarDelay", m_barDelay);
		if (m_exportEML == 1)
			CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_options, L"exportEML", CString(L"y"));
		else
			CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_options, L"exportEML", CString(L"n"));

		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options_set_charset_ids, L"fromCharsetId", m_from_charsetId);
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options_set_charset_ids, L"toCharsetId", m_to_charsetId);
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options_set_charset_ids, L"subjCharsetId", m_subj_charsetId);
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options_set_charset_ids, L"showCharsets", m_show_charsets);

		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"imageViewer", m_bImageViewer);
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"timeType", m_bTimeType);
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"enhancedSelectFolderDialog", m_bEnhancedSelectFolderDlg);
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"subjectSortType", m_bSubjectSortType);
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"filesToValidateAsMboxType", m_filesToValidateAsMboxType);

		CDialog::OnOK();
	}
}


BOOL COptionsDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	CString section_options = CString(sz_Software_mboxview) + L"\\Options";

	m_format = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"dayMonthYearFormat");

	DWORD barDelay = 0;
	BOOL retval;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"progressBarDelay", barDelay))
		m_barDelay = barDelay;
	else
		m_barDelay = 1;

	CString exportEML;
	if (retval = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_options, L"exportEML", exportEML)) {
		if (exportEML.Compare(L"y") == 0)
			m_exportEML = 1;
		else
			m_exportEML = 0;
	}
	else
		m_exportEML = 0;  // we should not be here; it should be initialized in NListView::NListView()

	m_from_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"fromCharsetId");
	m_to_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"toCharsetId");
	m_subj_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"subjCharsetId");
	m_show_charsets = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"showCharsets");
	m_bImageViewer = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"imageViewer");
	m_bTimeType = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"timeType");
	m_bEnhancedSelectFolderDlg = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"enhancedSelectFolderDialog");
	m_bSubjectSortType = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"subjectSortType");
	m_filesToValidateAsMboxType = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"filesToValidateAsMboxType");

	UpdateData(FALSE);

	HWND h = this->GetSafeHwnd();
	ResHelper::LoadDialogItemsInfo(h);

	return TRUE;  // return TRUE unless you set the focus to a control
	// ECCEZIONE: le pagine delle proprietà OCX devono restituire FALSE
}


void COptionsDlg::OnBnClickedRadio2()
{
	// TODO: Add your control notification handler code here
	int deb = 1;
}


void COptionsDlg::OnBnClickedPictureViewer()
{
	// TODO: Add your control notification handler code here
	int deb = 1;
}
