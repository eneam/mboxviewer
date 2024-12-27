#pragma once
#include "afxdialogex.h"


// MyCToolBarEx dialog

class MyCToolBarEx : public CToolBar
{
	DECLARE_DYNAMIC(MyCToolBarEx)

public:
	MyCToolBarEx(CWnd* pParent = nullptr);   // standard constructor
	virtual ~MyCToolBarEx();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_MyCToolBarEx };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg LRESULT OnInitDialog(WPARAM, LPARAM);
	afx_msg BOOL OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult);
	virtual BOOL Create(CWnd* pParentWnd, DWORD dwStyle = WS_CHILD | WS_VISIBLE | CBRS_TOP, UINT nID = AFX_IDW_TOOLBAR);
	virtual BOOL CreateEx(CWnd* pParentWnd, DWORD dwCtrlStyle = TBSTYLE_FLAT, DWORD dwStyle = WS_CHILD | WS_VISIBLE | CBRS_ALIGN_TOP,
		CRect rcBorders = CRect(0, 0, 0, 0),
		UINT nID = AFX_IDW_TOOLBAR);
};
