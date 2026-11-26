// MyCDialogBar.cpp : implementation file
//

#include "stdafx.h"
//#include "afxdialogex.h"
#include "ResHelper.h"
#include "MainFrm.h"
#include "MyCDialogBar.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif


// MyCDialogBar dialog

IMPLEMENT_DYNAMIC(MyCDialogBar, CDialogBar)

MyCDialogBar::MyCDialogBar()
{
	m_darkBackgroundBrush.CreateSolidBrush(RGB(30, 33, 38));
	int deb = 1;
}

MyCDialogBar::~MyCDialogBar()
{
	int deb = 1;
}

void MyCDialogBar::DoDataExchange(CDataExchange* pDX)
{
	CDialogBar::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(MyCDialogBar, CDialogBar)
	//ON_MESSAGE(WM_INITDIALOG, OnInitDialog)   // not needed for ToolTips
	ON_NOTIFY_EX(TTN_NEEDTEXT, 0, &MyCDialogBar::OnTtnNeedText)
	ON_WM_CTLCOLOR()
END_MESSAGE_MAP()


// MyCDialogBar message handlers
#if 0
LRESULT MyCDialogBar::OnInitDialog(WPARAM wParam, LPARAM lParam)
{

	LRESULT bRet = HandleInitDialog(wParam, lParam);

	//ResHelper::LoadDialogItemsInfo(this);
	//ResHelper::UpdateDialogItemsInfo(this);
	BOOL retA = ResHelper::ActivateToolTips(this, m_toolTip);

	if (!UpdateData(FALSE))
	{
		TRACE0("Warning: UpdateData failed during dialog init.\n");
	}

	return bRet;
}
#endif

HBRUSH MyCDialogBar::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	if (CMainFrame::m_bDarkTheme)
	{
		pDC->SetTextColor(COLOR_DARK_TEXT);
		pDC->SetBkColor(RGB(30, 33, 38));
		pDC->SetBkMode(TRANSPARENT);
		return (HBRUSH)m_darkBackgroundBrush.GetSafeHandle();
	}
	return CDialogBar::OnCtlColor(pDC, pWnd, nCtlColor);
}

BOOL MyCDialogBar::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}

