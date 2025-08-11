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


#include "stdafx.h"
#include "CustomCheckListBox.h"
#include "MainFrm.h"

CCustomCheckListBox::CCustomCheckListBox()
{
	m_nWidth = 0;
}

CCustomCheckListBox::~CCustomCheckListBox()
{
}

#if 0
int CCustomCheckListBox::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	return 0;
}
#endif

#if 0
void CCustomCheckListBox::DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct)
{
	int itemID = lpDrawItemStruct->itemID;
	CString txt;
	GetText(itemID, txt);
	CCheckListBox::DrawItem(lpDrawItemStruct);
	return;
}
#endif

void CCustomCheckListBox::updateWidth(int &maxWidth)
{
	CString txt;
	maxWidth = 0;
	for (int i = 0; i < GetCount(); i++)
	{
		GetText(i, txt);
		updateWidth(txt);
		if (m_nWidth > maxWidth)
			maxWidth = m_nWidth;
	}
}

void CCustomCheckListBox::updateWidth(CString &s)
{
#if 1
	CClientDC dc(this);
	//CFont * f = CListBox::GetFont();
	CFont& f = CMainFrame::m_dfltFont;

	dc.SelectObject(&f);
	CSize sz = dc.GetTextExtent(s);
	sz.cx += 3 * ::GetSystemMetrics(SM_CXBORDER);
	if (sz.cx > m_nWidth)
	{
		m_nWidth = sz.cx;
		m_nWidth += 18;
		CListBox::SetHorizontalExtent(m_nWidth);
	}
#else
	// Per Microsoft document. Works but needs to update dialog ResizeDialogWindow()
	CSize sz;
	CDC *pDC = GetDC();
	sz = pDC->GetTextExtent(s);
	if (sz.cx > m_nWidth)
	{
		m_nWidth = sz.cx;
		CListBox::SetHorizontalExtent(m_nWidth);
	}
	ReleaseDC(pDC);
#endif
}