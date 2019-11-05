//
//////////////////////////////////////////////////////////////////
//
//  Windows Mbox Viewer is a free tool to view, search and print mbox mail archives.
//
// Source code and executable can be downloaded from
//  https://sourceforge.net/projects/mbox-viewer/  and
//  https://github.com/eneam/mboxviewer
//
//  Copyright(C) 2019  Enea Mansutti, Zbigniew Minciel
//
//  This program is free software; you can redistribute it and/or
//  modify it under the terms of the GNU Library General Public
//  License as published by the Free Software Foundation; 
//  as version 2 of the License.
//  either version 2 of the License, or (at your option) any later version.
//
//  This program is distributed in the hope that it will be useful,
//  but WITHOUT ANY WARRANTY; without even the implied warranty of
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU
//  Library General Public License for more details.
//
//  You should have received a copy of the GNU Library General Public
//  License along with this program; if not, write to the
//  Free Software Foundation, Inc., 51 Franklin St, Fifth Floor,
//  Boston, MA  02110 - 1301, USA.
//
//////////////////////////////////////////////////////////////////
//

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
	BOOL m_bNavigateComplete;

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
