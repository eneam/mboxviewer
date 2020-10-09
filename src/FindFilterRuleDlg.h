#pragma once


// FindFilterRuleDlg dialog

class FindFilterRuleDlg : public CDialogEx
{
	DECLARE_DYNAMIC(FindFilterRuleDlg)

public:
	FindFilterRuleDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~FindFilterRuleDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_FIND_FILTER_DLG };
#endif

	int m_filterNumb;

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
};
