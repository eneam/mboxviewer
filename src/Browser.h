#if !defined(AFX_BROWSER_H__9D0E1F01_1CD0_11D3_8188_00105A4CBBB3__INCLUDED_)
#define AFX_BROWSER_H__9D0E1F01_1CD0_11D3_8188_00105A4CBBB3__INCLUDED_

#include "webbrowser2.h"	// Added by ClassView
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Browser.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CBrowser window

class CBrowser : public CWnd
{
// Construction
public:
	CBrowser();

// Attributes
protected:
//	CWebBrowserPrint m_wbprint;
	bool m_bDidNavigate, m_bDops;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBrowser)
	public:
	virtual BOOL OnAmbientProperty(COleControlSite* pSite, DISPID dispid, VARIANT* pvar);
	protected:
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	//}}AFX_VIRTUAL

// Implementation
public:
protected:
	HWND m_hWndParent;
	BOOL m_bOffline;
	CString m_url;

	void PrintCurrentPage();
public:
	void PageSetup();
	void Navigate(LPCSTR url, DWORD flags);
	CWebBrowser2 m_ie;

	// Generated message map functions
protected:
	//{{AFX_MSG(CBrowser)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnDocumentCompleteExplorer(LPDISPATCH pDisp, VARIANT FAR* URL);
	afx_msg void BeforeNavigate(LPDISPATCH, VARIANT* URL,VARIANT* Flags, VARIANT* TargetFrameName,VARIANT* PostData, VARIANT* Headers, BOOL* Cancel);
	afx_msg void OnTimer(UINT nIDEvent);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
	DECLARE_EVENTSINK_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BROWSER_H__9D0E1F01_1CD0_11D3_8188_00105A4CBBB3__INCLUDED_)
