// InputBox.cpp : implementation file
//

#include "stdafx.h"
#include "InputBox.h"
#include "afxdialogex.h"


// InputBox dialog

IMPLEMENT_DYNAMIC(InputBox, CDialogEx)

InputBox::InputBox(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_INPUT_BOX_DLG, pParent)
{

}

InputBox::~InputBox()
{
}

void InputBox::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_STRING, m_input);
}


BEGIN_MESSAGE_MAP(InputBox, CDialogEx)
END_MESSAGE_MAP()


// InputBox message handlers


BOOL InputBox::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here

	if (GetSafeHwnd()) {
		CWnd *p = GetDlgItem(IDYES);
		if (p) {
			GetDlgItem(IDC_STATIC)->SetWindowText(m_input);
		}
	}

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}
