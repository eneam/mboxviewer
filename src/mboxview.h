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
//  This program is free software; you can redistribute it and/or modify
//  it under the terms of the version 3 of GNU Affero General Public License
//  as published by the Free Software Foundation; 
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

// mboxview.h : main header file for the mboxview application
//

#if !defined(AFX_mboxview_H__01D75A85_01DB_4CC8_A34E_AC20E309168D__INCLUDED_)
#define AFX_mboxview_H__01D75A85_01DB_4CC8_A34E_AC20E309168D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"       // main symbols
#include <string>
#include "MainFrm.h"       // main symbols
#include "TextUtilities.h"

extern const wchar_t *sz_Software_mboxview;

// Define below in single location if more user messages are implemented
#define WM_CMD_PARAM_FILE_NAME_MESSAGE  (WM_APP + 1)
#define WM_CMD_PARAM_ATTACHMENT_HINT_MESSAGE  (WM_APP + 2)
#define WM_CMD_PARAM_GENERAL_HINT_MESSAGE  (WM_APP + 3)
#define WM_CMD_PARAM_NEW_COLOR_MESSAGE  (WM_APP + 4)
#define WM_CMD_PARAM_LOAD_FOLDERS_MESSAGE  (WM_APP + 5)
#define WM_CMD_PARAM_RESET_TREE_POS_MESSAGE  (WM_APP + 6)
#define WM_CMD_PARAM_ON_SIZE_MSGVIEW_MESSAGE  (WM_APP + 7)
#define WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE  (WM_APP + 8)

/////////////////////////////////////////////////////////////////////////////
// CmboxviewApp:
// See mboxview.cpp for the implementation of this class
//

#define MAPPING_SIZE 268435456


class MyCRecentFileList :public CRecentFileList
{
public:
	MyCRecentFileList(UINT nStart, LPCTSTR lpszSection,
		LPCTSTR lpszEntryFormat, int nSize,
		int nMaxDispLen = AFX_ABBREV_FILENAME_LEN);

	virtual void ReadList();
	virtual void WriteList();

	void DumpMRUsArray(CString& prefix);
};

class CmboxviewApp : public CWinApp
{
public:
	CmboxviewApp();
	~CmboxviewApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CmboxviewApp)
public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();
	virtual void AddToRecentFileList(LPCWSTR lpszPathName);
	//}}AFX_VIRTUAL

// Implementation
	static BOOL GetMboxviewLongVersion(CString &version);
	static BOOL GetFileVersionInfo(HMODULE hModule, DWORD &ms, DWORD &ls);
	static DWORD m_versionMS;
	static DWORD m_versionLS;
	static CString m_savedVer;
	//
	static CString m_processPath;
	static CString m_startupPath;
	static CString m_currentPath;

	// Moved from CProfile otherwise m_registry & m_configFilePath are destroyed
	// in CmboxviewApp::ExitInstance and cause crash
	// static vars and objects are desttred after CmboxviewApp is destroyed
	static ConfigTree* m_configTree;
	static BOOL m_registry;
	static CString m_configFilePath;
	static BOOL m_configFileLoaded;
	//static DebugCString m_configFilePath;  // for testing object life cycle

	static BOOL GetProcessPath(CString &processpath);
	static HWND GetActiveWndGetSafeHwnd();

	static CWnd* wndFocus;

public:
	//{{AFX_MSG(CmboxviewApp)
	afx_msg void OnAppAbout();
	afx_msg void OnHelpDonate();
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	afx_msg void MyMRUFileHandler(UINT i);
	DECLARE_MESSAGE_MAP()
};


extern CmboxviewApp theApp;

void ShellExecuteError2Text(UINT_PTR errorCode, CString &errorText);
void MyAppendMenu(CMenu *menu, int commandId, const wchar_t *commandName, BOOL checkMark = FALSE);
int AttachIcon(CMenu* Menu, LPCWSTR MenuName, int resourceId, CBitmap  &cmap);
void Com_Initialize();
void SetMyExceptionHandler();
void UnSetMyExceptionHandler();

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_mboxview_H__01D75A85_01DB_4CC8_A34E_AC20E309168D__INCLUDED_)
