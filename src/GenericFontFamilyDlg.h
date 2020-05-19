#pragma once


// GenericFontFamilyDlg dialog

class GenericFontFamilyDlg : public CDialogEx
{
	DECLARE_DYNAMIC(GenericFontFamilyDlg)

public:
	GenericFontFamilyDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~GenericFontFamilyDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_GENERIC_FONT_DLG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()

	void LoadData();

public:
	CListBox m_listBox;
	CString m_genericFontName;  // font family such as 

	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
};
