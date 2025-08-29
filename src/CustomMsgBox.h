#pragma once
#include "afxdialogex.h"


// CustomMsgBox dialog

class CustomMsgBox : public CDialogEx
{
	DECLARE_DYNAMIC(CustomMsgBox)

public:
	CustomMsgBox(CWnd* pParent = nullptr);   // standard constructor

	CustomMsgBox(LPCTSTR lpszText, LPCTSTR lpszCaption, UINT nType, int textFontHeight, CWnd* pParent = nullptr);

	virtual ~CustomMsgBox();

	CString m_textStr;
	CString captionStr;
	int m_textFontHeight;
	UINT m_nStyle;

	CButton m_button1;
	CButton m_button2;
	CButton m_button3;
	CEdit m_text;
	CStatusBar	m_wndStatusBar;
	int m_StatusBarHeight;

	HICON	m_hIcon;  // standard MessageBox win system icon
	CStatic m_icon;
	//
	HICON	m_hGripper;  // didn't yet work. May need to cleanup later
	CStatic m_gripper;

	HICON LoadMsgBoxIcon(BOOL systemIcon, PCWSTR pszName, int lims, CString &errorText);
	int LongestWordLength(CString& str, CString* longetWord = 0);
	int LongestLineLength(CString& str, CString* longetLine = 0);

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_CUSTOM_MSG_BOX };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();

	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSize(UINT nType, int cx, int cy);

	DECLARE_MESSAGE_MAP()
};
