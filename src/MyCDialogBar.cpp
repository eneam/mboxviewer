// MyCDialogBar.cpp : implementation file
//

#include "stdafx.h"
//#include "afxdialogex.h"
#include "ResHelper.h"
#include "MyCDialogBar.h"


// MyCDialogBar dialog

IMPLEMENT_DYNAMIC(MyCDialogBar, CDialogBar)

MyCDialogBar::MyCDialogBar()
{
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

BOOL MyCDialogBar::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}

