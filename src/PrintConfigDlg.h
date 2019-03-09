
#if !defined(_PRINT_CONFIG_DLG_)
#define _PRINT_CONFIG_DLG_

#pragma once


// PrintConfigDlg dialog

class PrintNameConfig
{
};

struct NamePatternParams
{
	int m_bDate;
	int m_bTime;
	int m_bFrom;
	int m_bTo;
	int m_bSubject;
	int m_bUniqueId;
	int m_nFileNameFormatSizeLimit;

	void SetDflts();
	void Copy(NamePatternParams &src);
	void UpdateRegistry(NamePatternParams &current, NamePatternParams &updated);
	void LoadFromRegistry();
};

class PrintConfigDlg : public CDialogEx
{
	DECLARE_DYNAMIC(PrintConfigDlg)

public:
	PrintConfigDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~PrintConfigDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_PRINT_DLG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
public:
	struct NamePatternParams m_NamePatternParams;

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	afx_msg void OnEnChangeFileNameMaxSize();
};

#endif _PRINT_CONFIG_DLG_
