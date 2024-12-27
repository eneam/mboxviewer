// MyCToolBar.cpp : implementation file
//

#include "stdafx.h"
#include "ResHelper.h"
#include "MyCToolBar.h"


// MyCToolBar

IMPLEMENT_DYNAMIC(MyCToolBar, CToolBar)

MyCToolBar::MyCToolBar()
{
	int deb = 1;
}

MyCToolBar::~MyCToolBar()
{
	int deb = 1;
}


BEGIN_MESSAGE_MAP(MyCToolBar, CToolBar)
	//ON_MESSAGE(WM_INITDIALOG, OnInitDialog)  //  never called anyway
	ON_NOTIFY_EX(TTN_NEEDTEXT, 0, &MyCToolBar::OnTtnNeedText)
END_MESSAGE_MAP()



// MyCToolBar message handlers

BOOL MyCToolBar::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}

#if 0
// ToolTips are enabled by default anyway
LRESULT MyCToolBar::OnInitDialog(WPARAM wParam, LPARAM lParam)
{
	LRESULT bRet = 0;

	//ResHelper::LoadDialogItemsInfo(this);
	//ResHelper::UpdateDialogItemsInfo(this);
	BOOL retA = ResHelper::ActivateToolTips(this, m_toolTip);

	if (!UpdateData(FALSE))
	{
		TRACE0("Warning: UpdateData failed during dialog init.\n");
	}

	return bRet;
}

BOOL MyCToolBar::Create(CWnd* pParentWnd, DWORD dwStyle, UINT nID)
{
	// TODO: Add your specialized code here and/or call the base class

	return CToolBar::Create(pParentWnd, dwStyle, nID);
}


BOOL MyCToolBar::OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult)
{
	// TODO: Add your specialized code here and/or call the base class

	return CToolBar::OnNotify(wParam, lParam, pResult);
}
#endif
