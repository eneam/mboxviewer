#pragma once
#include "afxdialogex.h"


// FontConfig dialog

class FontConfig : public CDialogEx
{
	DECLARE_DYNAMIC(FontConfig)

public:
	FontConfig(CWnd* pParent = nullptr);   // standard constructor
	virtual ~FontConfig();

	virtual INT_PTR DoModal();
	CWnd* m_pParent;

	CToolTipCtrl m_toolTip;

	int m_dflFontSizeSelected;
	int m_fontSize;  // hight
	int m_dfltFontSize;
	CEdit m_customFonSize;

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG_FONT_SIZE };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();

	DECLARE_MESSAGE_MAP()
public:
	afx_msg BOOL OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedRadioFontDflt();
	afx_msg void OnBnClickedRadioFontCustom();
};
