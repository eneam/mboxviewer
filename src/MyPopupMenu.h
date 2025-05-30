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
	~MyPopupMenu();

	int m_fontSize;
	CFont m_font;
	BOOL m_isCustomFont;
	CString m_menuName;

	BOOL IsCustomFont(int& fontSize);

	void SetMenuAsCustom(CMenu* menu, int index);
	void ReleaseCustomResources(CMenu* menu, int index);

	virtual void DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct);
	virtual void MeasureItem(LPMEASUREITEMSTRUCT lpMeasureItemStruct);

	virtual void DrawCheck(HDC hdc, SIZE size);
	virtual void DrawUnCheck(HDC hdc, SIZE size);

	static RECT m_rec;
	static CRect m_crec;

	virtual BOOL OnMeasureItem(LPMEASUREITEMSTRUCT lpms);
};

