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


// SelectLanguageDlg.cpp : implementation file
//

#include "stdafx.h"
#include "afxdialogex.h"
#include "SelectLanguageDlg.h"
#include "FileUtils.h"
#include "ResHelper.h"
#include "MainFrm.h"



// SelectLanguageDlg dialog

IMPLEMENT_DYNAMIC(SelectLanguageDlg, CDialogEx)

SelectLanguageDlg::SelectLanguageDlg(CWnd* pParent /*=nullptr*/)
DIALOG_FROM_TEMPLATE( : CDialogEx(IDD_LANG_LIST_DLG, pParent))
{
	m_pParent = pParent;
}

SelectLanguageDlg::~SelectLanguageDlg()
{
}

INT_PTR SelectLanguageDlg::DoModal()
{
#ifdef _DIALOG_FROM_TEMPLATE
	INT_PTR ret = CMainFrame::SetTemplate(this, IDD_LANG_LIST_DLG, m_pParent);
#else
	INT_PTR ret = CDialogEx::DoModal();
#endif
	return ret;
}

void SelectLanguageDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_SELECT_LANG, m_listBox);
}


BEGIN_MESSAGE_MAP(SelectLanguageDlg, CDialogEx)
	ON_BN_CLICKED(IDC_STATIC_SELECT_LANG_1, &SelectLanguageDlg::OnBnClickedStaticSelectLang1)
	ON_BN_CLICKED(IDOK, &SelectLanguageDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &SelectLanguageDlg::OnBnClickedCancel)
END_MESSAGE_MAP()


// SelectLanguageDlg message handlers


void SelectLanguageDlg::OnBnClickedStaticSelectLang1()
{
	// TODO: Add your control notification handler code here
	int deb = 1;
}

BOOL SelectLanguageDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here

	LoadLanguages();

	int lastIndex = m_listBox.FindStringExact(0, m_lastFolderName);
	if (lastIndex < 0)
		lastIndex = 0;
	int retIndex = m_listBox.SetCurSel(lastIndex);

	//this->SetBackgroundColor(::GetSysColor(COLOR_3DFACE));

	//ResHelper::LoadDialogItemsInfo(this);  // keep list of languages in English ??
	ResHelper::UpdateDialogItemsInfo(this);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

void SelectLanguageDlg::LoadLanguages()
{
	CFileFind finder;

	// build a string with wildcards
	CString folderPath;
	ResHelper::GetProcessFolderPath(folderPath);

	CString languagesPath = folderPath + L"Language";

	CString strWildcard(languagesPath);
	strWildcard += L"\\*.*";

	// start working for files
	BOOL bWorking = finder.FindFile(strWildcard);

	m_listBox.AddString(L"english");
	while (bWorking)
	{
		bWorking = finder.FindNextFile();

		// skip . and .. files; otherwise, we'd
		// recur infinitely!

		if (finder.IsDots())
			continue;

		if (finder.IsDirectory())
		{
			CString str = finder.GetFilePath();
			CString FolderName;
			FileUtils::CPathStripPath((LPCWSTR)str, FolderName);
			TRACE(L"%s\n", (LPCWSTR)str);

			if (FolderName.CompareNoCase(L"english") != 0)
			{
				m_listBox.AddString(FolderName);
			}
		}
	}
	finder.Close();
}

void SelectLanguageDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here

	int nSel = m_listBox.GetCurSel();

	if ((nSel >= 0) && (nSel < m_listBox.GetCount()))
	{
		CString text;
		m_listBox.GetText(nSel, text);

		m_LanguagetName = text;
	}

	CDialogEx::OnOK();
}


void SelectLanguageDlg::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here
	CDialogEx::OnCancel();
}

#define BLACK      RGB(0,0,0)
#define YELLOW     RGB(255, 255, 0)
#define RED        RGB(127,0,0)

BEGIN_MESSAGE_MAP(DerivedCListBox, CListBox)
	ON_WM_CTLCOLOR_REFLECT()
END_MESSAGE_MAP()

DerivedCListBox::DerivedCListBox()
{

	m_crBkColor = ::GetSysColor(COLOR_3DFACE); // Initializing background color to the system face color.
	//m_crBkColor = YELLOW;
	m_crTextColor = BLACK; // Initializing text color to black; not abl to change anyway
	m_brBkgnd.CreateSolidBrush(m_crBkColor); // Creating the Brush Color For the Edit Box Background
}


void DerivedCListBox::SetTextColor(COLORREF crColor)
{
	m_crTextColor = crColor; // Passing the value passed by the dialog to the member varaible for Text Color
	RedrawWindow();
}

void DerivedCListBox::SetBkColor(COLORREF crColor)
{
	m_crBkColor = crColor; // Passing the value passed by the dialog to the member varaible for Backgound Color
	m_brBkgnd.DeleteObject(); // Deleting any Previous Brush Colors if any existed.
	m_brBkgnd.CreateSolidBrush(crColor); // Creating the Brush Color For the Edit Box Background
	RedrawWindow();
}

HBRUSH DerivedCListBox::CtlColor(CDC* pDC, UINT nCtlColor)
{
	HBRUSH hbr;

	hbr = (HBRUSH)m_brBkgnd; // Passing a Handle to the Brush
	pDC->SetBkColor(m_crBkColor); // Setting the Color of the Text Background to the one passed by the Dialog
	pDC->SetTextColor(m_crTextColor); // TODO: doesn't work ; Setting the Text Color to the one Passed by the Dialog

	return hbr;
}
