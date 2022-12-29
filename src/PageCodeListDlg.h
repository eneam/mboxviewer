#pragma once
#include "afxdialogex.h"


// PageCodeListDlg dialog


struct CodePageInfo {
	UINT codePage;
	CString codePageName;
	UINT maxCharSize;
};

class PageCodeListDlg : public CDialogEx
{
	DECLARE_DYNAMIC(PageCodeListDlg)

public:
	PageCodeListDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~PageCodeListDlg();

	
	void SortCodePageList(int column, BOOL descendingSort);
	void PrintListCtrl(HTREEITEM hItem, BOOL recursive, int depth);

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_PAGE_CODE_LIST_DLG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnItemactivateCodePageEnumeration(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnGetdispinfoCodePageEnumeration(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnItemclickCodePageEnumeration(NMHDR* pNMHDR, LRESULT* pResult);
	//
	virtual BOOL OnInitDialog();
	virtual void OnOK();
	virtual void OnCancel();

	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnPaint();

	static BOOL CALLBACK EnumCodePagesProc(LPTSTR  pCP);
	static PageCodeListDlg* pPageCodeListDlgInstance;  // Unfortunatelly EnumCodePagesProc doesn't accept  pointer to private data

	CListCtrl m_list;
	HICON m_hIcon;

	CArray<CodePageInfo> m_codePageList;
	BOOL m_descendingSort = TRUE;
	CString m_tmpstr;
	UINT m_CodePageSelected;

	void AddCodePage(UINT codePage);
	int CodePage2Index(UINT codePage);
	void ItemState2Str(UINT uState, CString& strState);

	afx_msg void OnClickCodePageEnumeration(NMHDR* pNMHDR, LRESULT* pResult);
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual BOOL OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult);
	afx_msg void OnTimer(UINT_PTR nIDEvent);
};
