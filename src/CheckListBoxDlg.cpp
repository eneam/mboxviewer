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

// CheckListBoxDlg.cpp : implementation file
//

#include "stdafx.h"
#include "CheckListBoxDlg.h"
#include "afxdialogex.h"
#include "ResHelper.h"
#include "MainFrm.h"


// CCheckListBoxDlg dialog

IMPLEMENT_DYNAMIC(CCheckListBoxDlg, CDialogEx)

CCheckListBoxDlg::CCheckListBoxDlg(CWnd* pParent /*=nullptr*/)
DIALOG_FROM_TEMPLATE( : CDialogEx(IDD_CHECK_LIST_BOX_DLG, pParent))
{
	m_title = L"Select Items";
	m_pParent = pParent;
}

CCheckListBoxDlg::~CCheckListBoxDlg()
{
}

INT_PTR CCheckListBoxDlg::DoModal()
{
#ifdef _DIALOG_FROM_TEMPLATE
	INT_PTR ret = CMainFrame::SetTemplate(this, IDD_CHECK_LIST_BOX_DLG, m_pParent);
#else
	INT_PTR ret = CDialogEx::DoModal();
#endif
	return ret;
}

int CCheckListBoxDlg::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	return 0;
}

void CCheckListBoxDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_CHECK_LIST_BOX, m_listBox);
}


BEGIN_MESSAGE_MAP(CCheckListBoxDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &CCheckListBoxDlg::OnBnClickedOk)
	ON_NOTIFY_EX(TTN_NEEDTEXT, 0, &CCheckListBoxDlg::OnTtnNeedText)
END_MESSAGE_MAP()


// CCheckListBoxDlg message handlers


BOOL CCheckListBoxDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here
	m_listBox.SetFont(GetFont());

	SetWindowText(m_title);

	LoadData();

	ResHelper::LoadDialogItemsInfo(this);
	ResHelper::UpdateDialogItemsInfo(this);

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}


void CCheckListBoxDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here

	CString checked;
	int iCnt = m_listBox.GetCount();
	int chk;
	int i;
	for (i = 0; i < iCnt; i++)
	{
		chk = m_listBox.GetCheck(i);
		if (chk)
		{
			m_listBox.GetText(i, checked);
			m_OutList.Add(checked);
		}
		int deb = 1;
	}

	if (m_OutList.GetCount() <= 0) {
		CString txt = L"No folders selected!";
		ResHelper::TranslateString(txt);
		AfxMessageBox(txt, MB_OK | MB_ICONHAND);
		return;
	}

	CDialogEx::OnOK();
}

void CCheckListBoxDlg::LoadData()
{
	int i;
	for (i = 0; i < m_InList.GetCount(); i++)
	{
		m_listBox.AddString(m_InList.GetAt(i));
	}

	ResizeDialogWindow();
}

void CCheckListBoxDlg::ResizeDialogWindow()
{
	CRect wndRect;
	GetWindowRect(wndRect);


	int nMaxWidth;
	m_listBox.updateWidth(nMaxWidth);

	int cx = wndRect.Width();
	if (nMaxWidth > cx)
		cx = nMaxWidth;

	CRect wRect;
	m_listBox.GetWindowRect(wRect);

	CRect iRect;
	m_listBox.GetItemRect(0, iRect);

	int iCnt = m_listBox.GetCount();

	int scrollVwidth = GetSystemMetrics(SM_CXVSCROLL);
	int scrollHwidth = GetSystemMetrics(SM_CYHSCROLL);

	if (nMaxWidth > wRect.Width())
	{
		cx = wndRect.Width() + (nMaxWidth - wRect.Width()) + scrollVwidth + 9;
	}
	else
	{
		cx = wndRect.Width() + scrollVwidth + 9;
	}

	CRect screenRect;
	CWnd::GetDesktopWindow()->GetWindowRect(screenRect);

	int cy;
	if (cx > screenRect.Width() / 2)
	{
		cx = screenRect.Width() / 2;
		cy = iCnt * iRect.Height() + wndRect.Height();
		cy += 3;
	}
	else
	{
		cy = iCnt * iRect.Height() + wndRect.Height();
	}

	if (cy > screenRect.Height() / 2)
	{
		cy = screenRect.Height() / 2;
	}

	SetWindowPos(0, 0, 0, cx, cy, SWP_NOMOVE | SWP_NOZORDER);
}

BOOL CCheckListBoxDlg::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}

#if 0

Dialog resource configuration.

IDD_CHECK_LIST_BOX_DLG DIALOGEX 0, 0, 117, 45
STYLE DS_SETFONT | DS_FIXEDSYS | WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME
CAPTION "Dialog"
FONT 8, "MS Shell Dlg", 400, 0, 0x1
BEGIN
DEFPUSHBUTTON   "OK", IDOK, 7, 23, 50, 14
PUSHBUTTON      "Cancel", IDCANCEL, 60, 23, 50, 14
LISTBOX         IDC_LIST_CHECK_LIST_BOX, 8, 6, 102, 13, LBS_OWNERDRAWFIXED | LBS_HASSTRINGS | LBS_NOINTEGRALHEIGHT | WS_VSCROLL | WS_HSCROLL | WS_TABSTOP
END


Dialog :

Layout:
-------------------------------------------- -
Title
CCheckListBox Control - just single row
OK Button  Cancel button in the same row
Width is set to just make Ok and Cancel visible
-------------------------------------------- -

Border->Resing
Maximize Box - > False
Minimize Box->False
Style->Popup

CCheckListBox :

Horizonat Scroll->True
Vertical Scrollbar->True
Has Strings->True
Owner Draw->Fixed
Sort->False
Moving Type->None
Sizing Type->Both

Ok Button:

Moving Type->Vertical
Sizing Type->None

Cancel Button:

Moving Type->Vertical
Sizing Type->None

#endif
