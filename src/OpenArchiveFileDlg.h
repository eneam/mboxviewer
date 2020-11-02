#pragma once


// OpenArchiveFileDlg dialog

class OpenArchiveFileDlg : public CDialogEx
{
	DECLARE_DYNAMIC(OpenArchiveFileDlg)

public:
	OpenArchiveFileDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~OpenArchiveFileDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_OPEN_ARCHIVE_DLG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()

public:
	virtual BOOL OnInitDialog();

	CString m_sourceFolder;
	CString m_targetFolder;
	CString m_archiveFileName;

	//afx_msg void OnBnClickedYes();
	afx_msg void OnBnClickedOk();
};
