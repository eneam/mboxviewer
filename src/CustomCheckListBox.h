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

#pragma once
#include <afxwin.h>

class CCustomCheckListBox :
	public CCheckListBox
{
public:
	CCustomCheckListBox();
	~CCustomCheckListBox();

	//virtual BOOL Create(DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID);
	//void DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct);
	int m_nWidth;

	void updateWidth(CString &s);
	void updateWidth(int &maxWidth);

	//DECLARE_MESSAGE_MAP()
public:
	//afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);

protected:
};

