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

// MyPopupMenu.cpp : implementation file
// 
// You can set/change the background color
//


#pragma once

#include "stdafx.h"
#include "ResHelper.h"

typedef struct tagMyPopupMenuItem
{
	HFONT m_hfont;
	HFONT m_hmenuBarFont;
	CString m_text;
	int m_maxTextLeftPartLengthInPoints;  // true text
	int m_maxTextRightPartLengthInPoints;  // accelerators
	int m_fType;
	int m_drawSeparator;
	BOOL m_isMenuBarItem;
} MyPopupMenuItem;

class MyPopupMenu:public CMenu
{
public:
	MyPopupMenu();
	virtual ~MyPopupMenu();

	static UINT  MenuItemInfoMaskAllSet;
	static UINT  MenuItemInfoMaskTypeAllSet;
	static UINT  MenuItemInfoMaskFTypeAllSet;

	static int m_fontSize;
	static CFont m_font;
	static LOGFONT m_menuLogFont;
	//
	static int m_menuBarFontSize;
	static CFont m_menuBarFont;
	static LOGFONT m_menuBarLogFont;
	static BOOL m_isCustomFont;

	
	void UpdateFontSize(int fontSize, int index = 0);
	//
	static int TraceMenu(CString& title, CMenu* menu, int index, UINT mask);
	static BOOL TraceMenuItem(CString& title, CMenu* menu, int index, UINT mask);
	static BOOL PrintMENUITEMINFO(CString& text, MENUITEMINFO& minfo, UINT mask);
	static BOOL IsCustomFont(int& fontSize);
	static BOOL HasID(CMenu* menu, UINT ID);
	static void SetupFonts();
	static void ReleaseGlobalResources();
	static int SetCMenuAsCustom(CMenu* menu, int index = 0);
	static int RestoreCMenu(CMenu* menu, int index);
	static void FindLengthOfLongestText(CMenu* menu, BOOL& hasTab, int& maxTextLeftPartLengthInPoints, int& maxTextRightPartLengthInPoints, int index);
	static void OnMeasureItem(HWND hWnd, MEASUREITEMSTRUCT* lpMeasureItem);
	static void OnDrawItem(int nIDCtl, LPDRAWITEMSTRUCT lpDrawItemStruct);
	static void GetMenuFont(HDC hdc, CFont& font, int hight, LOGFONT& menuLogFont, CFont& menuFont);
	static void GetLengthOfMenuLabelPartsInPoints(HWND hWnd, HFONT hf, CString& label, int& textLeftPartLengthInPoints, int& textRightPartLengthInPoints);
	static BOOL GetLabelParts(CString& label, CString& text, CString& accelerator);


	// Below functions are specific for popup context menus
	void SetMenuAsCustom(int index = 0);
	virtual void DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct);
	virtual void MeasureItem(LPMEASUREITEMSTRUCT lpMeasureItemStruct);
	static int ReleaseCustomResources(CMenu* menu, int index = 0);
};

