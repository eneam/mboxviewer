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

#include "stdafx.h"
#include "MyPopupMenu.h"
#include "MainFrm.h"

int MyPopupMenu::m_fontSize = 0;
CFont MyPopupMenu::m_font;
LOGFONT MyPopupMenu::m_menuLogFont;
BOOL MyPopupMenu::m_isCustomFont = FALSE;

MyPopupMenu::MyPopupMenu(CString &menuName)
{
	m_menuName = menuName;
#if 0
	m_fontSize = 8;

	HDC hdc = 0;
	CFont font;

	LOGFONT menuLogFont;
	CFont menuFont;
	m_isCustomFont = IsCustomFont(m_fontSize);
	if (m_isCustomFont)
		ResHelper::GetMenuFont(hdc, font, m_fontSize, menuLogFont, m_font);
#endif
}

MyPopupMenu::~MyPopupMenu()
{
	if (m_isCustomFont)
		ReleaseCustomResources(0);
	m_font.DeleteObject();
	int deb = 1;
}

void MyPopupMenu::SetupFonts()
{
	HDC hdc = 0;
	CFont font;
	if (m_fontSize == 0)
	{
		m_fontSize = CMainFrame::m_dfltFontSize;
		m_isCustomFont = IsCustomFont(m_fontSize);
		if (m_isCustomFont)
			m_fontSize = CMainFrame::CMainFrame::m_cnfFontSize;

		m_font.DeleteObject();
		ResHelper::GetMenuFont(hdc, font, m_fontSize, m_menuLogFont, m_font);
	}
	else
	{
		m_isCustomFont = IsCustomFont(m_fontSize);
		if (m_isCustomFont)
		{
			m_font.DeleteObject();
			ResHelper::GetMenuFont(hdc, font, m_fontSize, m_menuLogFont, m_font);
		}
	}
}

BOOL MyPopupMenu::IsCustomFont(int& fontSize)
{
	if (CMainFrame::m_cnfFontSize != CMainFrame::m_dfltFontSize)
	{
		fontSize = CMainFrame::m_cnfFontSize;
		return TRUE;
	}
	else
		return FALSE;
}

void MyPopupMenu::DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct)
{
	CRect rcItem = lpDrawItemStruct->rcItem;
	MyPopupMenuItem* pMyItem = (MyPopupMenuItem*)lpDrawItemStruct->itemData;
	_ASSERTE(pMyItem);
	if (!pMyItem)
		return;

	COLORREF clrPrevText, clrPrevBkgnd;
	HFONT hfntPrev;
	int x, y;

	// Set the appropriate foreground and background colors. 
	if (lpDrawItemStruct->itemState & ODS_SELECTED)
	{
		clrPrevText = SetTextColor(lpDrawItemStruct->hDC,
			GetSysColor(COLOR_HIGHLIGHTTEXT));
		clrPrevBkgnd = SetBkColor(lpDrawItemStruct->hDC,
			GetSysColor(COLOR_MENUHILIGHT));
	}
	else
	{
		clrPrevText = SetTextColor(lpDrawItemStruct->hDC,
			GetSysColor(COLOR_MENUTEXT));
		clrPrevBkgnd = SetBkColor(lpDrawItemStruct->hDC,
			GetSysColor(COLOR_MENU));
	}

	hfntPrev = (HFONT)SelectObject(lpDrawItemStruct->hDC, pMyItem->m_hfont);

	int textOffset = (int)((float)2 * (float)GetSystemMetrics(SM_CXMENUCHECK));

	if (pMyItem->m_fType & MFT_SEPARATOR)
	{
		CDC dc;
		dc.Attach(lpDrawItemStruct->hDC);
		CRect rc = rcItem;
		rc.top += rc.Height() >> 1;
		rc.left += textOffset;

		UINT nFlags = BF_TOP | BF_FLAT;
		dc.DrawEdge(&rc, EDGE_ETCHED, nFlags);
		dc.Detach();

		// Restore the original font and colors.
		SetTextColor(lpDrawItemStruct->hDC, clrPrevText);
		SetBkColor(lpDrawItemStruct->hDC, clrPrevBkgnd);
		SelectObject(lpDrawItemStruct->hDC, hfntPrev);
		return;
	}

	// Determine where to draw and leave space for a check mark. 
	x = rcItem.left;
	y = rcItem.top;

	// No longer needed.
	// Added afx_msg void OnMeasureItem(int nIDCtl, LPMEASUREITEMSTRUCT lpMeasureItemStruct);
	// to NListView and NTreeView and 
#if 0
	// MeasureItem() is received by MyPopupMenu for all menu items except Popup submenus
	// This is a big limitation, may need to explore other more complex solutions
	// This limitation doesn't allow Windows to correctly size Popup submenue
	// Calculate adjustment to y to compensate, best guess by try and error
	int yAdjustmen = 0;
	if (lpDrawItemStruct->itemID == -1)
	{
		if (m_fontSize == 10)
			yAdjustmen  = 2;
		else if (m_fontSize == 12)
			yAdjustmen  = 3;
		else if (m_fontSize == 14)
			yAdjustmen  = 4;
		else if (m_fontSize == 16)
			yAdjustmen  = 5;
		// else doesn't work for larger font

		//y -= yAdjustmen;
	}
#endif
	x += textOffset;

	// Select the font and draw the text. 
	ExtTextOut(lpDrawItemStruct->hDC, x, y, ETO_OPAQUE,
		&rcItem, pMyItem->m_text,
		pMyItem->m_text.GetLength(), NULL);

	// Select the font and draw the check mark character.
	CString checkMark;
	if (lpDrawItemStruct->itemState & ODS_CHECKED)
		checkMark = L"\u2713";

	RECT rc = rcItem;
	rc.right = GetSystemMetrics(SM_CXMENUCHECK);
	x = rc.left + 4;
	y = rc.top;
#if 0
	if (lpDrawItemStruct->itemID == -1)
	{
		;// y -= yAdjustmen;
	}
#endif

	if (!checkMark.IsEmpty())
	{
		//DWORD color = GetSysColor(COLOR_MENUHILIGHT);  // too dark ??
		DWORD color = RGB(102, 178,255); 

		clrPrevBkgnd = SetBkColor(lpDrawItemStruct->hDC, color);

		ExtTextOut(lpDrawItemStruct->hDC, x, y, ETO_OPAQUE,
			&rc, checkMark,
			checkMark.GetLength(), NULL);
	}

	// Restore the original font and colors. 
	SelectObject(lpDrawItemStruct->hDC, hfntPrev);
	SetTextColor(lpDrawItemStruct->hDC, clrPrevText);
	SetBkColor(lpDrawItemStruct->hDC, clrPrevBkgnd);

	int deb = 1;
}

void MyPopupMenu::MeasureItem(LPMEASUREITEMSTRUCT lpMeasureItemStruct)
{
	MyPopupMenuItem* pMyItem = (MyPopupMenuItem*)lpMeasureItemStruct->itemData;

	if (lpMeasureItemStruct->itemID == -1)
	{
		;// TRACE(L"MeasureItem: ID=-1\n");
	}

	if (pMyItem->m_fType & MFT_SEPARATOR)
	{
		lpMeasureItemStruct->itemHeight = GetSystemMetrics(SM_CYMENU) >> 1;
		return;
	}

	HDC hdc = ::GetWindowDC(NULL);
	HFONT hfntOld = (HFONT)SelectObject(hdc, pMyItem->m_hfont);
	SIZE size;

	GetTextExtentPoint32(hdc, pMyItem->m_text, pMyItem->m_text.GetLength(), &size);

	SelectObject(hdc, hfntOld);
	::ReleaseDC(NULL, hdc);

	int textOffset = (int)((float)4 * (float)GetSystemMetrics(SM_CXMENUCHECK));

	lpMeasureItemStruct->itemWidth = size.cx + textOffset;
	lpMeasureItemStruct->itemHeight = size.cy;

	int deb = 1;
}

void MyPopupMenu::SetMenuAsCustom(int index)
{
	MyPopupMenu* menu = this;

	if (!m_isCustomFont)
		return;

	HMENU hM = m_hMenu;
	const HMENU hMenu = GetSafeHmenu();
	BOOL isMenu = ::IsMenu(hMenu);

	_ASSERTE(isMenu);
	if (!isMenu)
		return;

	int count = menu->GetMenuItemCount();
	CString label;
	UINT itemID = 0;
	int retval;
	UINT nFlags = MF_BYPOSITION;
	int i;
	MENUITEMINFO menuItemInfo;
	menuItemInfo.cbSize = sizeof(menuItemInfo);

	for (i = 0; i < count; i++)
	{
		itemID = menu->GetMenuItemID(i);

		if (itemID == 0)
			int deb = 1; // continue;

		retval = menu->GetMenuString(i, label, nFlags);

		memset(&menuItemInfo, 0, sizeof(menuItemInfo));
		menuItemInfo.cbSize = sizeof(menuItemInfo);
		menuItemInfo.fMask = MIIM_ID | MIIM_TYPE | MIIM_DATA;
		BOOL retval = menu->GetMenuItemInfo(i, &menuItemInfo, TRUE);
		if (retval == FALSE)
			int deb = 1; // ignore/return;

		_ASSERTE(menuItemInfo.dwItemData == 0);
		if (menuItemInfo.dwItemData)
		{
			// Already set, return error
			MyPopupMenuItem* myItem = (MyPopupMenuItem*)menuItemInfo.dwItemData;
			continue;
		}

		MyPopupMenuItem* myItem = new MyPopupMenuItem;
		myItem->m_fType = menuItemInfo.fType;
		myItem->m_menuName = m_menuName;
		myItem->m_text = label;
		myItem->m_hfont = m_font.operator HFONT();

		menuItemInfo.fMask = MIIM_ID | MIIM_DATA | MIIM_FTYPE;
		menuItemInfo.fType = MFT_OWNERDRAW;
		menuItemInfo.dwItemData = (ULONG_PTR)myItem;

		BOOL retset = menu->SetMenuItemInfo(i, &menuItemInfo, TRUE);
		_ASSERT(retset);

		if (itemID == (UINT)-1)
		{
			MyPopupMenu* submenu = (MyPopupMenu*)menu->GetSubMenu(i);

			_ASSERTE(submenu);
			if (submenu)
				submenu->SetMenuAsCustom(index + 1);
			int deb = 1;
		}
		int deb = 1;
	}
	int debM = 1;
}

void MyPopupMenu::ReleaseCustomResources(int index)
{
	CMenu* menu = this;

	if (!m_isCustomFont)
		return;

	HMENU hM = m_hMenu;
	const HMENU hMenu = GetSafeHmenu();
	if (!hMenu)
		return;
	BOOL isMenu = ::IsMenu(hMenu);

	if (!isMenu)
		return;

	int count = menu->GetMenuItemCount();
	CString label;
	UINT itemID = 0;
	int retval;
	UINT nFlags = MF_BYPOSITION;
	int i;
	MENUITEMINFO menuItemInfo;

	for (i = 0; i < count; i++)
	{
		itemID = menu->GetMenuItemID(i);

		retval = menu->GetMenuString(i, label, nFlags);

		memset(&menuItemInfo, 0, sizeof(menuItemInfo));
		menuItemInfo.cbSize = sizeof(menuItemInfo);
		menuItemInfo.fMask = MIIM_ID | MIIM_DATA | MIIM_FTYPE | MIIM_STATE | MIIM_STRING;
		BOOL retval = menu->GetMenuItemInfo(i, &menuItemInfo, nFlags);
		if (retval == FALSE)
			int deb = 1; // ignore/return;

		if (menuItemInfo.dwItemData)
		{
			MyPopupMenuItem* myItem = (MyPopupMenuItem*)menuItemInfo.dwItemData;
			delete myItem;
		}
		else
			int deb = 1;

		if (itemID == (UINT)-1)
		{
			CMenu* submenu = (MyPopupMenu*)menu->GetSubMenu(i);
			//_ASSERTE(submenu);
			// // Let each submenue release its dwItemData resource otherwise we will crash
			//if (submenu) submenu->ReleaseCustomResources(index + 1);
		}
		int deb = 1;
	}
	int debM = 1;
}

void MyPopupMenu::UpdateFontSize(int fontSize, int index)
{
	MyPopupMenu* menu = this;

	if (!m_isCustomFont)
		return;

	HMENU hM = m_hMenu;
	const HMENU hMenu = GetSafeHmenu();
	BOOL isMenu = ::IsMenu(hMenu);

	if (!isMenu)
		return;

	int count = menu->GetMenuItemCount();
	CString label;
	UINT itemID = 0;
	int retval;
	UINT nFlags = MF_BYPOSITION;
	int i;
	MENUITEMINFO menuItemInfo;

	for (i = 0; i < count; i++)
	{
		itemID = menu->GetMenuItemID(i);

		retval = menu->GetMenuString(i, label, nFlags);;

		memset(&menuItemInfo, 0, sizeof(menuItemInfo));
		menuItemInfo.cbSize = sizeof(menuItemInfo);
		menuItemInfo.fMask = MIIM_ID | MIIM_DATA | MIIM_FTYPE | MIIM_STATE | MIIM_STRING;
		BOOL retval = menu->GetMenuItemInfo(i, &menuItemInfo, nFlags);
		if (retval == FALSE)
			int deb = 1; // ignore/return;

		if (menuItemInfo.dwItemData)
		{
			MyPopupMenuItem* myItem = (MyPopupMenuItem*)menuItemInfo.dwItemData;
			myItem->m_hfont = m_font.operator HFONT();
		}

		if (itemID == (UINT)-1)
		{
			MyPopupMenu* submenu = (MyPopupMenu*)menu->GetSubMenu(i);
			_ASSERTE(submenu);
			if (submenu)
				submenu->UpdateFontSize(fontSize, index + 1);
		}
		int deb = 1;
	}
	int debM = 1;
}

BOOL MyPopupMenu::HasID(CMenu* menu, UINT ID)
{
	_ASSERTE(menu);
	if (!menu)
		return FALSE;

	HMENU hM = menu->m_hMenu;
	const HMENU hMenu = menu->GetSafeHmenu();
	BOOL isMenu = ::IsMenu(hMenu);

	if (!isMenu)
		return FALSE;

	// Both versions seem to work. If ID == -1 check will fail
#if 1
	MENUITEMINFO menuItemInfo;
	memset(&menuItemInfo, 0, sizeof(menuItemInfo));
	menuItemInfo.cbSize = sizeof(menuItemInfo);
	menuItemInfo.fMask = MIIM_ID;

	BOOL fByPos = FALSE;
	BOOL retExist = menu->GetMenuItemInfo(ID, &menuItemInfo, fByPos);
	return retExist;
#else
	int count = menu->GetMenuItemCount();
	UINT itemID = 0;
	int i;
	for (i = 0; i < count; i++)
	{
		itemID = menu->GetMenuItemID(i);

		if (itemID == ID)
		{
			_ASSERTE(retExist);
			return TRUE;
		}

		if (itemID == (UINT)-1)
		{
			CMenu* submenu = (MyPopupMenu*)menu->GetSubMenu(i);
			_ASSERTE(submenu);

			if (submenu && MyPopupMenu::HasID(submenu, ID))
			{
				_ASSERTE(retExist);
				return TRUE;
			}
		}
		int deb = 1;
	}
	_ASSERTE(!retExist);
	return FALSE;
#endif
}
