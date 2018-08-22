#pragma once


// ExportToCSVDlg dialog

class ExportToCSVDlg : public CDialogEx
{
	DECLARE_DYNAMIC(ExportToCSVDlg)

public:
	ExportToCSVDlg(CWnd* pParent = 0);   // standard constructor
	virtual ~ExportToCSVDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_EXPORT_TO_CSV };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	BOOL m_bFrom;
	BOOL m_bTo;
	BOOL m_bSubject;
	BOOL m_bDate;
	BOOL m_bContent;
	int m_dateFormat;
	int m_bGMTTime;
	int m_bEncodingType;

	afx_msg BOOL OnInitDialog();
	afx_msg void OnOK();
	afx_msg void OnBnClickedEncodingNone();
	afx_msg void OnBnClickedEncodingUtf8();
	afx_msg void OnBnClickedEncodingCodePage();
	int m_nCodePageId;
};
