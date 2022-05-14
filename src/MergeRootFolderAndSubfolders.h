#pragma once


// MergeRootFolderAndSubfolders dialog

class MergeRootFolderAndSubfolders : public CDialogEx
{
	DECLARE_DYNAMIC(MergeRootFolderAndSubfolders)

public:
	MergeRootFolderAndSubfolders(CWnd* pParent = nullptr);   // standard constructor
	virtual ~MergeRootFolderAndSubfolders();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_MERGE_FOLDER_AND_SUBFOLDERS };
#endif

	CEdit m_introText;

	int m_mergeRootFolderStyle;
	int m_labelAssignmentStyle;

	COLORREF m_folderPathColor;
	CBrush m_folderPathBrush;
	CFont m_BoldFont;
	CFont m_TextFont;
	CBrush m_ButtonBrush;

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()

public:
	virtual BOOL OnInitDialog();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd *pWnd, UINT nCtlColor);
	afx_msg void OnBnClickedLabelPerMboxFile();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedConfigFolderStyle1();
	afx_msg void OnBnClickedMergeRootFolderAndSubfolders();
};
