// ColorPickerDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ColorPickerDlg.h"

#ifndef _WIN32_WCE // CColorDialog is not supported for Windows CE.

// ColorPickerDlg

IMPLEMENT_DYNAMIC(ColorPickerDlg, CColorDialog)

ColorPickerDlg::ColorPickerDlg(COLORREF clrInit, DWORD dwFlags, CWnd* pParentWnd) :
	CColorDialog(clrInit, dwFlags, pParentWnd)
{

}

ColorPickerDlg::~ColorPickerDlg()
{
}


BEGIN_MESSAGE_MAP(ColorPickerDlg, CColorDialog)
END_MESSAGE_MAP()



// ColorPickerDlg message handlers


#endif // !_WIN32_WCE


BOOL ColorPickerDlg::OnInitDialog()
{
	CColorDialog::OnInitDialog();

	// TODO:  Add extra initialization here

	CWnd *pwnd = GetParent();
	CRect cr;
	GetClientRect(cr);
	CRect pr;
	pwnd->GetWindowRect(pr);
	int xpos = pr.left - cr.Width() - 16;
	if (xpos < 0)
		xpos = 0;
	SetWindowPos(pwnd, pr.left, pr.bottom, 0, 0, SWP_SHOWWINDOW | SWP_NOSIZE);

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}


void ColorPickerDlg::OnOK()
{
	// TODO: Add your specialized code here and/or call the base class



	CColorDialog::OnOK();
}


BOOL ColorPickerDlg::OnColorOK()
{
	// TODO: Add your specialized code here and/or call the base class

	//return TRUE;

	return CColorDialog::OnColorOK();
}
