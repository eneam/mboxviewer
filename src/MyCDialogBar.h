#pragma once

// 
// MyCDialogBar dialog

class MyCDialogBar : public CDialogBar
{
	DECLARE_DYNAMIC(MyCDialogBar)

public:
	MyCDialogBar();   // standard constructor
	virtual ~MyCDialogBar();

	//CToolTipCtrl m_toolTip;

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_MyCDialogBar };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg LRESULT OnInitDialog(WPARAM, LPARAM);
	afx_msg BOOL OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult);
};

