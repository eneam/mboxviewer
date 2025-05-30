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



#include "stdafx.h"
#include "MyPopupMenu.h"
#include "MainFrm.h"

RECT MyPopupMenu::m_rec;
CRect MyPopupMenu::m_crec;


MyPopupMenu::MyPopupMenu(CString &menuName)
{
	m_fontSize = 8;
	m_menuName = menuName;
	HDC hdc = 0;
	CFont font;
	int hight = 24;
	LOGFONT menuLogFont;
	CFont menuFont;
	m_isCustomFont = IsCustomFont(m_fontSize);
	ResHelper::GetMenuFont(hdc, font, m_fontSize, menuLogFont, m_font);
}

MyPopupMenu::~MyPopupMenu()
{
	ReleaseCustomResources(this, 0);
	m_font.DeleteObject();
	int deb = 1;
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
	// CMenu::DrawItem(lpDrawItemStruct);

	if (lpDrawItemStruct->itemID == 0)
		; // return;

	CRect rcItem = lpDrawItemStruct->rcItem;
	if (lpDrawItemStruct->itemID != -1)
	{
		m_rec = rcItem;
		m_crec = rcItem;
		TRACE(L"DrawItem ID = -1\n");

	}
	else
	{
		if (m_crec.top)
		{
			; //  rcItem.top = m_crec.bottom;  // doesn't work
		}
	}
	MyPopupMenuItem* pMyItem = (MyPopupMenuItem*)lpDrawItemStruct->itemData;
	COLORREF clrPrevText, clrPrevBkgnd;
	HFONT hfntPrev;
	int x, y;

	TRACE(L"%d %d %d %d\n", rcItem.left, rcItem.right, rcItem.top, rcItem.bottom);

	// Set the appropriate foreground and background colors. 
	if (lpDrawItemStruct->itemState & ODS_SELECTED)
	{
		clrPrevText = SetTextColor(lpDrawItemStruct->hDC,
			GetSysColor(COLOR_HIGHLIGHTTEXT));
		clrPrevBkgnd = SetBkColor(lpDrawItemStruct->hDC,
			GetSysColor(COLOR_HIGHLIGHT));
	}
	else
	{
		clrPrevText = SetTextColor(lpDrawItemStruct->hDC,
			GetSysColor(COLOR_MENUTEXT));
		clrPrevBkgnd = SetBkColor(lpDrawItemStruct->hDC,
			GetSysColor(COLOR_MENU));
	}

	hfntPrev = (HFONT)SelectObject(lpDrawItemStruct->hDC, pMyItem->m_hfont);

	if (pMyItem->m_fType & MFT_SEPARATOR)
	{
		CDC dc;
		dc.Attach(lpDrawItemStruct->hDC);
		CRect rc = rcItem;
		rc.top += rc.Height() >> 1;
		rc.left += 2 * GetSystemMetrics(SM_CXMENUCHECK);
		dc.DrawEdge(&rc, EDGE_ETCHED, BF_TOP);
		dc.Detach();
		TRACE(L"DrawItem: SEPARATOR\n");
		return;
	}

	// Determine where to draw and leave space for a check mark. 
	x = rcItem.left;
	y = rcItem.top;
	if (lpDrawItemStruct->itemID == -1)
	{
		if (m_fontSize == 10)
			y -= 2;
		else if (m_fontSize == 12)
			y -= 3;
		else if (m_fontSize == 14)
			y -= 4;
		else if (m_fontSize == 16)
			y -= 5;
	}
	x += 2* GetSystemMetrics(SM_CXMENUCHECK);

	// Select the font and draw the text. 
	ExtTextOut(lpDrawItemStruct->hDC, x, y, ETO_OPAQUE,
		&rcItem, pMyItem->m_text,
		pMyItem->m_text.GetLength(), NULL);

	// Select the font and draw the check mark character.
	CString checkMark;
	if (!(lpDrawItemStruct->itemState & ODS_CHECKED))
		checkMark = L"";
	else
		checkMark = L"\u2713";

	RECT rc = rcItem;
	rc.right = GetSystemMetrics(SM_CXMENUCHECK);;
	x = rc.left + 4;
	y = rc.top; //  -4;
	if (lpDrawItemStruct->itemID == -1)
	{
		if (m_fontSize == 10)
			y -= 2;
		else if (m_fontSize == 12)
			y -= 3;
		else if (m_fontSize == 14)
			y -= 4;
		else if (m_fontSize == 16)
			y -= 5;
	}

	TRACE(L"%d %d %d %d\n", rc.left, rc.right, rc.top, rc.bottom);

	if (!checkMark.IsEmpty())
	{
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
		TRACE(L"MeasureItem: ID=-1\n");
	}

	if (pMyItem->m_fType & MFT_SEPARATOR)
	{
		// separator: use half system height and zero width
		lpMeasureItemStruct->itemHeight = GetSystemMetrics(SM_CYMENU) >> 1;
		//lpMeasureItemStruct->itemWidth = 0;
		TRACE(L"MeasureItem: SEPARATOR\n");
		return;
	}

	if (pMyItem->m_fType & MFT_SEPARATOR)
	{
		int deb = 1;
	}

	TRACE(L"MeasureItem: NOT SEPARATOR %d\n", pMyItem->m_fType);

	HDC hdc = ::GetWindowDC(NULL);
	HFONT hfntOld = (HFONT)SelectObject(hdc, pMyItem->m_hfont);
	SIZE size;

	GetTextExtentPoint32(hdc, pMyItem->m_text,
		pMyItem->m_text.GetLength(), &size);

	SelectObject(hdc, hfntOld);
	::ReleaseDC(NULL, hdc);

	lpMeasureItemStruct->itemWidth = size.cx + 36;
	lpMeasureItemStruct->itemHeight = size.cy;

	int deb = 1;
}


void MyPopupMenu::SetMenuAsCustom(CMenu *menu, int index)
{
	//CMenu* menu = this;

	if (!m_isCustomFont)
		return;

	HMENU hM = m_hMenu;
	const HMENU hMenu = GetSafeHmenu();
	BOOL isMenu = ::IsMenu(hMenu);

	int count = menu->GetMenuItemCount();
	CString label;
	UINT itemID = 0;
	int retval;
	UINT nFlags = MF_BYPOSITION;
	int i;
	MENUITEMINFO menuItemInfo;
	menuItemInfo.cbSize = sizeof(menuItemInfo);

	TRACE(L"SetMenuAsCustom %ldd %s %d !!!!!\n", menu, m_menuName, count);

	for (i = 0; i < count; i++)
	{
		itemID = menu->GetMenuItemID(i);

		if (itemID == 0)
			int deb = 1; // continue;

		retval = menu->GetMenuString(i, label, nFlags);
		//retval = GetMenuItemString(menu, i, label, nFlags);

		memset(&menuItemInfo, 0, sizeof(menuItemInfo));
		menuItemInfo.cbSize = sizeof(menuItemInfo);
		//menuItemInfo.fMask = MIIM_SUBMENU | MIIM_ID | MIIM_TYPE | MIIM_DATA;
		menuItemInfo.fMask = MIIM_ID | MIIM_TYPE;
		BOOL retval = menu->GetMenuItemInfo(i, &menuItemInfo, nFlags);
		if (retval == FALSE)
			int deb = 1; // ignore/return;


		MyPopupMenuItem* myItem = new MyPopupMenuItem;
		myItem->m_fType = menuItemInfo.fType;
		myItem->m_menuName = m_menuName;
		myItem->m_text = label;
		myItem->m_hfont = m_font.operator HFONT();

		TRACE(L"SetMenuAsCustom: %s %s\n", myItem->m_menuName, myItem->m_text);

		menuItemInfo.fMask = MIIM_ID | MIIM_DATA | MIIM_FTYPE;
		menuItemInfo.fType = MFT_OWNERDRAW;
		menuItemInfo.dwItemData = (ULONG_PTR)myItem;

		BOOL retset = menu->SetMenuItemInfo(i, &menuItemInfo, TRUE);
#if 0
		CString newLabel;
		BOOL retFind = DetermineString(label, newLabel);

		DWORD nFlags2 = MF_STRING | MF_BYPOSITION;
		UINT_PTR nIDNewItem = (UINT_PTR)itemID;

		if (retFind && (itemID != (UINT)0))
		{
			BOOL retMod = menu->ModifyMenu((UINT)i, nFlags2, nIDNewItem, (LPCWSTR)newLabel);
			WORD nItemID = menu->GetMenuItemID(i);
			int deb = 1;
		}
		else
		{
			int deb = 1; // keep old text
		}
#endif

		if (itemID == (UINT)-1)
		{
			CMenu* submenu = menu->GetSubMenu(i);

			//SetMenuAsCustom(submenu, index + 1);
			int deb = 1;
		}
		int deb = 1;
	}
	int debM = 1;
}

void MyPopupMenu::ReleaseCustomResources(CMenu* menu, int index)
{
	//CMenu* menu = this;

	HMENU hM = m_hMenu;
	const HMENU hMenu = GetSafeHmenu();
	BOOL isMenu = ::IsMenu(hMenu);

	TRACE(L"ReleaseCustomResources %s isMenu=%d\n", m_menuName, isMenu);

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

	TRACE(L"ReleaseCustomResources %s %d\n", m_menuName, count);

	for (i = 0; i < count; i++)
	{
		itemID = menu->GetMenuItemID(i);

		if (itemID == 0)
			;// continue;

		retval = menu->GetMenuString(i, label, nFlags);
		//retval = GetMenuItemString(menu, i, label, nFlags);

		memset(&menuItemInfo, 0, sizeof(menuItemInfo));
		menuItemInfo.cbSize = sizeof(menuItemInfo);
		menuItemInfo.fMask = MIIM_ID | MIIM_DATA| MIIM_FTYPE| MIIM_STATE| MIIM_STRING;
		//menuItemInfo.dwItemData = 0;
		menuItemInfo.cch = 0;
		BOOL retval = menu->GetMenuItemInfo(i, &menuItemInfo, nFlags);
		if (retval == FALSE)
			int deb = 1; // ignore/return;

		if (menuItemInfo.dwItemData)
		{
			MyPopupMenuItem* myItem = (MyPopupMenuItem*)menuItemInfo.dwItemData;
			TRACE(L"Deleting: %s %s\n", myItem->m_menuName, myItem->m_text);
			delete myItem;
		}

#if 0
		menuItemInfo.fMask = MIIM_ID | MIIM_DATA | MIIM_FTYPE;
		menuItemInfo.fType = MFT_OWNERDRAW;
		menuItemInfo.dwItemData = 0;

		menu->SetMenuItemInfo(i, &menuItemInfo, TRUE);
#endif

		if (itemID == (UINT)-1)
		{
			CMenu* submenu = menu->GetSubMenu(i);

			//ReleaseCustomResources(submenu, index + 1);
		}
		int deb = 1;
	}
	int debM = 1;
}

void MyPopupMenu::DrawCheck(HDC hdc, SIZE size)
{

	HBRUSH hbrOld;
#if 0
	hbrOld = SelectObject(hdc, GetStockObject(NULL_BRUSH));
	Rectangle(hdc, 0, 0, size.cx, size.cy);
	MoveToEx(hdc, 0, 0, NULL);
	LineTo(hdc, size.cx, size.cy);
	MoveToEx(hdc, 0, size.cy - 1, NULL);
	LineTo(hdc, size.cx - 1, 0);
	SelectObject(hdc, hbrOld);
#endif
}

void MyPopupMenu::DrawUnCheck(HDC hdc, SIZE size)
{
	HBRUSH hbrOld;
#if 0
	hbrOld = SelectObject(hdc, GetStockObject(NULL_BRUSH));
	Rectangle(hdc, 0, 0, size.cx, size.cy);
	SelectObject(hdc, hbrOld);
#endif
}


#if 0
void MyCButton::SetTextColor(COLORREF crColor)
{
	m_crTextColor = crColor; // Passing the value passed by the dialog to the member varaible for Text Color
	RedrawWindow();
}

void MyCButton::SetBkColor(COLORREF crColor)
{
	m_crBkColor = crColor; // Passing the value passed by the dialog to the member varaible for Backgound Color
	m_brBkgnd.DeleteObject(); // Deleting any Previous Brush Colors if any existed.
	m_brBkgnd.CreateSolidBrush(crColor); // Creating the Brush Color For the Edit Box Background
	RedrawWindow();
}

HBRUSH MyCButton::CtlColor(CDC* pDC, UINT nCtlColor)
{
	HBRUSH hbr;

	hbr = (HBRUSH)m_brBkgnd; // Passing a Handle to the Brush
	pDC->SetBkColor(m_crBkColor); // Setting the Color of the Text Background to the one passed by the Dialog
	pDC->SetTextColor(m_crTextColor); // TODO: doesn't work ; Setting the Text Color to the one Passed by the Dialog

	return hbr;
}
#endif

BOOL MyPopupMenu::OnMeasureItem(LPMEASUREITEMSTRUCT lpms)
{
	return TRUE;
}
