
#pragma once


// MyCToolBar

class MyCToolBar : public CToolBar
{
	DECLARE_DYNAMIC(MyCToolBar)

public:
	MyCToolBar();
	virtual ~MyCToolBar();

	//CToolTipCtrl m_toolTip;

protected:
	DECLARE_MESSAGE_MAP()

public:
	//afx_msg LRESULT OnInitDialog(WPARAM, LPARAM);
	afx_msg BOOL OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult);

	//virtual BOOL Create(CWnd* pParentWnd, DWORD dwStyle = WS_CHILD | WS_VISIBLE | CBRS_TOP, UINT nID = AFX_IDW_TOOLBAR);
	//virtual BOOL OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult);
};


