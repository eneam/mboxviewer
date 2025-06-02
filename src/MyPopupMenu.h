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

// MyCButton.cpp : implementation file
// 
// You can set/change the background color
//


#pragma once

#include "stdafx.h"
#include "ResHelper.h"


typedef struct tagMyPopupMenuItem
{
	HFONT m_hfont;
	CString m_text;
	CString m_menuName;
	int m_fType;
} MyPopupMenuItem;

class MyPopupMenu:public CMenu
{
public:
	MyPopupMenu(CString& menuName);
	virtual ~MyPopupMenu();

	static int m_fontSize;
	static CFont m_font;
	static LOGFONT m_menuLogFont;
	static BOOL m_isCustomFont;

	CString m_menuName;

	void SetMenuAsCustom(int index = 0);
	void UpdateFontSize(int fontSize, int index = 0);
	void ReleaseCustomResources(int index = 0);
	//
	static BOOL IsCustomFont(int& fontSize);
	static BOOL HasID(CMenu* menu, UINT ID);
	static void SetupFonts();
	static void ReleaseGlobalResources();

	virtual void DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct);
	virtual void MeasureItem(LPMEASUREITEMSTRUCT lpMeasureItemStruct);
};

