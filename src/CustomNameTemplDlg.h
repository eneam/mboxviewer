#pragma once

// CustomNameTemplDlg dialog


struct NameTemplateCnf
{
	CString	m_TemplateFormat;
	CString	m_DateFormat;
	int m_bFromUsername;
	int m_bFromDomain;
	int m_bToUsername;
	int m_bToDomain;
	int m_bReplaceWhiteWithUnderscore;

	void ClearParts();
	void SetDflts();
	void Copy(NameTemplateCnf &src);
};

class CustomNameTemplDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CustomNameTemplDlg)

public:
	CustomNameTemplDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~CustomNameTemplDlg();

	struct NameTemplateCnf m_nameTemplateCnf;

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_NAME_TEMPL_DLG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedSrcftime();
};
