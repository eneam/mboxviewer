// MyCToolBarEx.cpp : implementation file
//

#include "stdafx.h"
#include "afxdialogex.h"
#include "ResHelper.h"
#include "MyCToolBarEx.h"


// MyCToolBarEx dialog

IMPLEMENT_DYNAMIC(MyCToolBarEx, CToolBar)

MyCToolBarEx::MyCToolBarEx(CWnd* pParent /*=nullptr*/)
	: CToolBar()
	//: CToolBar(IDD_MyCToolBarEx, pParent)
{

}

MyCToolBarEx::~MyCToolBarEx()
{
}

void MyCToolBarEx::DoDataExchange(CDataExchange* pDX)
{
	CToolBar::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(MyCToolBarEx, CToolBar)
	ON_MESSAGE(WM_INITDIALOG, OnInitDialog)
	ON_NOTIFY_EX(TTN_NEEDTEXT, 0, &MyCToolBarEx::OnTtnNeedText)
END_MESSAGE_MAP()


// MyCToolBarEx message handlers


LRESULT MyCToolBarEx::OnInitDialog(WPARAM wParam, LPARAM lParam)
{
	//CToolBar::OnInitDialog();

	// TODO:  Add extra initialization here

	return 0;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

BOOL MyCToolBarEx::CreateEx(CWnd* pParentWnd, DWORD dwCtrlStyle, DWORD dwStyle, CRect rcBorders, UINT nID)
{
	return CToolBar::CreateEx(pParentWnd, dwCtrlStyle, dwStyle, rcBorders, nID);
}

BOOL MyCToolBarEx::Create(CWnd* pParent, DWORD nStyle, UINT nID)
{
	// TODO: Add your specialized code here and/or call the base class

	BOOL bReturn = CToolBar::Create(pParent, nStyle, nID);

	return bReturn;
}

BOOL MyCToolBarEx::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}

