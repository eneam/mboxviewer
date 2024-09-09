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

// ForwardMailDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ForwardMailDlg.h"
#include "TextUtilsEx.h"
#include "afxdialogex.h"
#include "ResHelper.h"


// ForwardMailDlg dialog

IMPLEMENT_DYNAMIC(ForwardMailDlg, CDialogEx)

ForwardMailDlg::ForwardMailDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_FORWARD_EMAIL, pParent)
{

}

ForwardMailDlg::~ForwardMailDlg()
{
}

void ForwardMailDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_FORWARD_TO, m_Data.m_To);
	DDX_Text(pDX, IDC_EDIT_FORWARD_CC, m_Data.m_CC);
	DDX_Text(pDX, IDC_EDIT_FORWARD_BCC, m_Data.m_BCC);
	DDX_Control(pDX, IDC_EDIT_FORWARD_TEXT, m_EditText);
	DDX_Control(pDX, IDC_EDIT_FORWARD_SUBJECT, m_Subject);
	DDX_Text(pDX, IDC_EDIT_FORWARD_FROM, m_Data.m_From);
	DDX_Text(pDX, IDC_EDIT_FORWARD_SERVICE, m_Data.m_MailService);
}


BEGIN_MESSAGE_MAP(ForwardMailDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &ForwardMailDlg::OnBnClickedOk)
	ON_WM_CREATE()
	ON_BN_CLICKED(IDC_FORWARD_CLEAR, &ForwardMailDlg::OnBnClickedForwardClear)
END_MESSAGE_MAP()


// ForwardMailDlg message handlers

int ForwardMailDlg::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CWnd::OnCreate(lpCreateStruct) == -1)
		return -1;

	//DWORD dwStyle = ES_READONLY| ES_LEFT;
	//BOOL ret = m_Subject.Create(dwStyle, CRect(), this, IDC_EDIT_FORWARD_SUBJECT);

	m_Subject.SetSubject(m_Data.m_subjectW);

	return TRUE;
}

BOOL ForwardMailDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	m_EditText.SetWindowText(m_Data.m_Text);
	if (GetSafeHwnd())
	{
		CWnd* p = GetDlgItem(IDC_EDIT_FORWARD_FROM);
		if (p)
		{
			if (m_Data.m_MailService.CompareNoCase(L"Custom") == 0)
				p->EnableWindow(TRUE);
			else
				p->EnableWindow(FALSE);

		}
	}

	HWND h = this->GetSafeHwnd();
	ResHelper::LoadDialogItemsInfo(h);

	return TRUE;
}


void ForwardMailDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	UpdateData();

	m_EditText.GetWindowText(m_Data.m_Text);

	if (m_Data.m_To.GetAllocLength() == 0)
	{
		CString txt;
		txt.Format(L"Empty \"To:\" field not allowed!");
		AfxMessageBox(txt, MB_OK | MB_ICONHAND);
		return;
	}

	CDialogEx::OnOK();
}


void ForwardMailData::Copy(ForwardMailData &data)
{
	if (this == &data)
		return;

	m_MailService = data.m_MailService;
	m_From = data.m_From;
	m_To =  data.m_To;
	m_CC = data.m_CC;
	m_BCC = data.m_BCC;
	m_Text = data.m_Text;
	m_subjectW = data.m_subjectW;
}


MyCustomEdit::MyCustomEdit()
{
	int deb = 1;
}

void MyCustomEdit::OnPaint()
{
	CPaintDC dc(this);

	RECT rect;
	GetClientRect(&rect);

	HDC hDC = dc.GetSafeHdc();

	CFont newFont;
	newFont.CreatePointFont(85, L"Tahoma");

	// Set new font. Should reinstall old oldFont?? doesn't seem to matter
	CFont  *pOldFont = dc.SelectObject(&newFont);

	COLORREF color = ::GetSysColor(COLOR_3DFACE);
	dc.FillRect(&rect, &CBrush(color));
	dc.SetBkMode(TRANSPARENT);
	//dc.SetTextColor(txcolor);

	int xpos = 3;
	int ypos = 3;

	BOOL ret = ::ExtTextOutW(hDC, xpos, ypos, ETO_CLIPPED, &rect, (LPCWSTR)m_subjectW, m_subjectW.GetLength(), NULL);

	int deb = 1;
}

BEGIN_MESSAGE_MAP(MyCustomEdit, CEdit)
	ON_WM_PAINT()
END_MESSAGE_MAP()



void ForwardMailDlg::OnBnClickedForwardClear()
{
	// TODO: Add your control notification handler code here
	m_Data.m_To.Empty();
	m_Data.m_CC.Empty();
	m_Data.m_BCC.Empty();
	m_EditText.SetWindowText(L"");

	UpdateData(FALSE);
}
