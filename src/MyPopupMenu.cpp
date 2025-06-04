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

MyPopupMenu::MyPopupMenu()
{
}

MyPopupMenu::~MyPopupMenu()
{
	if (m_isCustomFont)
		ReleaseCustomResources(0);
	int deb = 1;
}

void MyPopupMenu::ReleaseGlobalResources()
{
	m_font.DeleteObject();
}

void MyPopupMenu::SetupFonts()
{
	HDC hdc = 0;
	CFont font;
	if (m_fontSize == 0)
	{
		m_fontSize = CMainFrame::m_dfltFontSize;
		m_isCustomFont = IsCustomFont(MyPopupMenu::m_fontSize);
		if (m_isCustomFont)
			MyPopupMenu::m_fontSize = CMainFrame::CMainFrame::m_cnfFontSize;

		m_font.DeleteObject();
		MyPopupMenu::GetMenuFont(hdc, font, MyPopupMenu::m_fontSize, MyPopupMenu::m_menuLogFont, MyPopupMenu::m_font);
	}
	else
	{
		m_isCustomFont = IsCustomFont(MyPopupMenu::m_fontSize);
		if (m_isCustomFont)
		{
			m_font.DeleteObject();
			MyPopupMenu::GetMenuFont(hdc, font, MyPopupMenu::m_fontSize, MyPopupMenu::m_menuLogFont, MyPopupMenu::m_font);
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


// Draw My popup menus
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


// Draw main manubar
void MyPopupMenu::OnDrawItem(int nIDCtl, LPDRAWITEMSTRUCT lpDrawItemStruct)
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

	int menuCheckSize = GetSystemMetrics(SM_CXMENUCHECK);
	int textOffset = (int)((float)2 * (float)menuCheckSize);

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

	x += textOffset;

	// Select the font and draw the text.
	CString text;
	CString accelerator;
	BOOL hasTab = MyPopupMenu::GetLabelParts(pMyItem->m_text, text, accelerator);

	if (!hasTab)
	{
		ExtTextOut(lpDrawItemStruct->hDC, x, y, ETO_OPAQUE,
			&rcItem, pMyItem->m_text, pMyItem->m_text.GetLength(), NULL);
	}
	else
	{
		ExtTextOut(lpDrawItemStruct->hDC, x, y, ETO_OPAQUE,
			&rcItem, text, text.GetLength(), NULL);

		CRect rc = rcItem;
		rc.left += pMyItem->m_maxTextLeftPartLengthInPoints + menuCheckSize;
		x += pMyItem->m_maxTextLeftPartLengthInPoints + menuCheckSize;

		ExtTextOut(lpDrawItemStruct->hDC, x, y, ETO_OPAQUE,
				&rc, accelerator, accelerator.GetLength(), NULL);
	}

	// Select the font and draw the check mark character.
	CString checkMark;
	if (lpDrawItemStruct->itemState & ODS_CHECKED)
		checkMark = L"\u2713";

	RECT rc = rcItem;
	rc.right = GetSystemMetrics(SM_CXMENUCHECK);
	x = rc.left + 4;
	y = rc.top;

	if (!checkMark.IsEmpty())
	{
		//DWORD color = GetSysColor(COLOR_MENUHILIGHT);  // too dark ??
		DWORD color = RGB(102, 178, 255);

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

// Received by My popup menu
void MyPopupMenu::MeasureItem(LPMEASUREITEMSTRUCT lpMeasureItemStruct)
{
	MyPopupMenuItem* pMyItem = (MyPopupMenuItem*)lpMeasureItemStruct->itemData;

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


// Received by menubar parent window
void MyPopupMenu::OnMeasureItem(HWND hWnd, MEASUREITEMSTRUCT* lpMeasureItemStruct)
{
	MyPopupMenuItem* pMyItem = (MyPopupMenuItem*)lpMeasureItemStruct->itemData;

	if (pMyItem->m_fType & MFT_SEPARATOR)
	{
		lpMeasureItemStruct->itemHeight = GetSystemMetrics(SM_CYMENU) >> 1;
		return;
	}

	HDC hdc = GetDC(hWnd);
	HFONT hfntOld = (HFONT)SelectObject(hdc, pMyItem->m_hfont);
	SIZE size;

	GetTextExtentPoint32(hdc, pMyItem->m_text, pMyItem->m_text.GetLength(), &size);

	SelectObject(hdc, hfntOld);
	ReleaseDC(hWnd, hdc);

	int textOffset = (int)((float)4 * (float)GetSystemMetrics(SM_CXMENUCHECK));

	lpMeasureItemStruct->itemWidth = pMyItem->m_maxTextLeftPartLengthInPoints + pMyItem->m_maxTextRightPartLengthInPoints + textOffset;
	lpMeasureItemStruct->itemHeight = size.cy;

	int deb = 1;
}


// Setup my popup menu
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
	CString checkLabel;
	UINT itemID = 0;
	UINT nFlags = MF_BYPOSITION;
	int i;
	MENUITEMINFO menuItemInfo;
	menuItemInfo.cbSize = sizeof(menuItemInfo);

	for (i = 0; i < count; i++)
	{
		itemID = menu->GetMenuItemID(i);

		if (itemID == 0)
			int deb = 1; // continue;

		int retGetStr = menu->GetMenuString(i, label, nFlags);
#if 0
		memset(&menuItemInfo, 0, sizeof(menuItemInfo));
		menuItemInfo.cbSize = sizeof(menuItemInfo);
		menuItemInfo.fMask = MIIM_ID | MIIM_FTYPE;
		BOOL retGetFtype = menu->GetMenuItemInfo(i, &menuItemInfo, TRUE);
		if (retGetFtype == FALSE)
			int deb = 1; // ignore/return;

		if (menuItemInfo.fType & MFT_MENUBARBREAK)
			int deb = 1;
#endif

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
		myItem->m_text = label;
		myItem->m_hfont = MyPopupMenu::m_font.operator HFONT();

		menuItemInfo.fMask =  MIIM_DATA | MIIM_FTYPE;
		//menuItemInfo.fMask = MIIM_ID | MIIM_DATA | MIIM_FTYPE;
		menuItemInfo.fType = MFT_OWNERDRAW;
		menuItemInfo.dwItemData = (ULONG_PTR)myItem;

		BOOL retSet = ((CMenu*)menu)->SetMenuItemInfo(i, &menuItemInfo, TRUE);
		_ASSERT(retSet);

		int retCheck = menu->GetMenuString(i, checkLabel, nFlags);

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
	int deb = 1;
}


// Setup menubar
void MyPopupMenu::SetCMenuAsCustom(CMenu *menu, int index)
{
	if (!m_isCustomFont)
		return;

	HMENU hM = menu->m_hMenu;
	const HMENU hMenu = menu->GetSafeHmenu();
	BOOL isMenu = ::IsMenu(hMenu);

	_ASSERTE(isMenu);
	if (!isMenu)
		return;

	BOOL hasTab = FALSE;

	int maxTextLeftPartLengthInPoints = 0; int maxTextRightPartLengthInPoints = 0;
	MyPopupMenu::FindLengthOfLongestText(menu, hasTab, maxTextLeftPartLengthInPoints, maxTextRightPartLengthInPoints, index);
	if (hasTab)
		int deb = 1;

	int count = menu->GetMenuItemCount();
	CString label;
	CString checkLabel;
	UINT itemID = 0;
	UINT nFlags = MF_BYPOSITION;
	int i;
	MENUITEMINFO menuItemInfo;
	menuItemInfo.cbSize = sizeof(menuItemInfo);

	for (i = 0; i < count; i++)
	{
		itemID = menu->GetMenuItemID(i);

		if (itemID == 0)
			int deb = 1; // continue;

		int retGetStr = menu->GetMenuString(i, label, nFlags);

		// Currently labels are already translated, no need to translate again, it will fail anyway
		// Left just in case if order of initialization would change
		CString fileMenuItem = L"&File";
		ResHelper::TranslateString(fileMenuItem);
		CString editMenuItem = L"&Edit";
		ResHelper::TranslateString(editMenuItem);
		CString viewMenuItem = L"&View";
		ResHelper::TranslateString(viewMenuItem);
		CString helpMenuItem = L"&Help";
		ResHelper::TranslateString(helpMenuItem);
		CString languageMenuItem = L"Language";
		ResHelper::TranslateString(languageMenuItem);

		if ((label.Compare(fileMenuItem) == 0) ||
			(label.Compare(editMenuItem) == 0) ||
			(label.Compare(viewMenuItem) == 0) ||
			(label.Compare(helpMenuItem) == 0) ||
			(label.Compare(languageMenuItem) == 0)
				)
		{
			int deb = 1; //
		}
		else
		{
			memset(&menuItemInfo, 0, sizeof(menuItemInfo));
			menuItemInfo.cbSize = sizeof(menuItemInfo);
			menuItemInfo.fMask = MIIM_ID | MIIM_TYPE | MIIM_DATA;
			BOOL retGetItemData = menu->GetMenuItemInfo(i, &menuItemInfo, TRUE);
			_ASSERTE(retGetItemData);
			if (retGetItemData == FALSE)
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
			CString text = label;
			text.Replace(L"&", L"");

			ResHelper::TranslateString(text);
			myItem->m_text = text;
			myItem->m_hfont = MyPopupMenu::m_font.operator HFONT();
			myItem->m_maxTextLeftPartLengthInPoints = maxTextLeftPartLengthInPoints;
			myItem->m_maxTextRightPartLengthInPoints = maxTextRightPartLengthInPoints;

			//menuItemInfo.fMask = MIIM_DATA | MIIM_FTYPE;
			menuItemInfo.fMask = MIIM_ID | MIIM_DATA | MIIM_FTYPE;
			menuItemInfo.fType = MFT_OWNERDRAW;
			menuItemInfo.dwItemData = (ULONG_PTR)myItem;

			BOOL retSet = ((CMenu*)menu)->SetMenuItemInfo(i, &menuItemInfo, TRUE);
			_ASSERT(retSet);
		}

		if (itemID == (UINT)-1)
		{
			CMenu* submenu = (MyPopupMenu*)menu->GetSubMenu(i);

			_ASSERTE(submenu);
			if (submenu)
				SetCMenuAsCustom(submenu, index + 1);
			int deb = 1;
		}
		int deb = 1;

	}
	int deb = 1;
}

// No iteration.
void  MyPopupMenu::FindLengthOfLongestText(CMenu* menu, BOOL &hasTab, 
	int& maxTextLeftPartLengthInPoints, int& maxTextRightPartLengthInPoints, int index)
{
	int textLeftPartLengthInPoints = 0;  // true text
	int textRightPartLengthInPoints = 0;  // accelerators

	maxTextLeftPartLengthInPoints = 0;  // true text
	maxTextRightPartLengthInPoints = 0;  // accelerators

	if (!m_isCustomFont)
		return;

	HMENU hM = menu->m_hMenu;
	const HMENU hMenu = menu->GetSafeHmenu();
	BOOL isMenu = ::IsMenu(hMenu);

	_ASSERTE(isMenu);
	if (!isMenu)
		return;

	int count = menu->GetMenuItemCount();
	CString label;
	CString checkLabel;
	UINT itemID = 0;
	UINT nFlags = MF_BYPOSITION;
	int i;
	MENUITEMINFO menuItemInfo;
	menuItemInfo.cbSize = sizeof(menuItemInfo);

	for (i = 0; i < count; i++)
	{
		itemID = menu->GetMenuItemID(i);

		if (itemID == 0)
			int deb = 1; // continue;

		int retGetStr = menu->GetMenuString(i, label, nFlags);

		// Currently labels are already translated, no need to translate again, it will fail anyway
		// Left just in case if order of initialization would change
		CString fileMenuItem = L"&File";
		ResHelper::TranslateString(fileMenuItem);
		CString editMenuItem = L"&Edit";
		ResHelper::TranslateString(editMenuItem);
		CString viewMenuItem = L"&View";
		ResHelper::TranslateString(viewMenuItem);
		CString helpMenuItem = L"&Help";
		ResHelper::TranslateString(helpMenuItem);
		CString languageMenuItem = L"Language";
		ResHelper::TranslateString(languageMenuItem);

		if ((label.Compare(fileMenuItem) == 0) ||
			(label.Compare(editMenuItem) == 0) ||
			(label.Compare(viewMenuItem) == 0) ||
			(label.Compare(helpMenuItem) == 0) ||
			(label.Compare(languageMenuItem) == 0)
			)
		{
			int deb = 1; //
		}
		else
		{
			memset(&menuItemInfo, 0, sizeof(menuItemInfo));
			menuItemInfo.cbSize = sizeof(menuItemInfo);
			menuItemInfo.fMask = MIIM_ID | MIIM_TYPE | MIIM_DATA;
			BOOL retGetItemData = menu->GetMenuItemInfo(i, &menuItemInfo, TRUE);
			_ASSERTE(retGetItemData);
			if (retGetItemData == FALSE)
				int deb = 1; // ignore/return;

			HFONT hf = MyPopupMenu::m_font.operator HFONT();

			HWND hWnd = 0;
			int textLeftPartLengthInPoints = 0;  // true text
			int textRightPartLengthInPoints = 0;  // accelerators
			MyPopupMenu::GetLengthOfMenuLabelPartsInPoints(hWnd, hf, label, textLeftPartLengthInPoints, textRightPartLengthInPoints);

			if (textLeftPartLengthInPoints > maxTextLeftPartLengthInPoints)
				maxTextLeftPartLengthInPoints = textLeftPartLengthInPoints;

			if (textRightPartLengthInPoints > maxTextRightPartLengthInPoints)
				maxTextRightPartLengthInPoints = textRightPartLengthInPoints;
		}

		if (itemID == (UINT)-1)
		{
			CMenu* submenu = (MyPopupMenu*)menu->GetSubMenu(i);

			_ASSERTE(submenu);
			//if (submenu) SetCMenuAsCustom(submenu, index + 1);
			int deb = 1;
		}
		int deb = 1;

	}
	return;
}

// My popup menu
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
			// Let each submenue release its dwItemData resource otherwise we will crash
			//if (submenu) submenu->ReleaseCustomResources(index + 1);
		}
		int deb = 1;
	}
	int deb = 1;
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
			myItem->m_hfont = MyPopupMenu::m_font.operator HFONT();
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
	int deb = 1;
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

	// Both versions work. If ID == -1 check will fail
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

void MyPopupMenu::GetMenuFont(HDC hdc, CFont& font, int hight, LOGFONT& menuLogFont, CFont& menuFont)
{
	NONCLIENTMETRICS ncm;
	memset(&ncm, 0, sizeof(NONCLIENTMETRICS));
	ncm.cbSize = sizeof(NONCLIENTMETRICS);

	VERIFY(::SystemParametersInfo(SPI_GETNONCLIENTMETRICS,
		sizeof(NONCLIENTMETRICS), &ncm, 0));

	menuLogFont = ncm.lfMenuFont;

	if (hdc == 0)
		hdc = ::GetWindowDC(NULL);
	menuLogFont.lfHeight = -MulDiv(hight, GetDeviceCaps(hdc, LOGPIXELSY), 72);

	::ReleaseDC(NULL, hdc);

	menuFont.CreateFontIndirect(&menuLogFont);
}

void MyPopupMenu::GetLengthOfMenuLabelPartsInPoints(HWND hWnd, HFONT hf, CString& label, int& textLength, int& acceleratorLength)
{
	HDC hdc = 0;
	if (hWnd)
		hdc = ::GetDC(hWnd);
	else
		hdc = ::GetWindowDC(hWnd);

	textLength = 0;
	acceleratorLength = 0;

	HFONT hfntOld = (HFONT)SelectObject(hdc, hf);
	SIZE textSize;
	SIZE acceleratorSize;

	CString leftPart;
	CString rightPart;
	int tabPos = label.FindOneOf(L"\t");
	if (tabPos >= 0)
	{
		leftPart = label.Mid(0, tabPos);
		rightPart = label.Mid(tabPos);

		BOOL retText = GetTextExtentPoint32(hdc, leftPart, leftPart.GetLength(), &textSize);
		BOOL retAccel = GetTextExtentPoint32(hdc, rightPart, rightPart.GetLength(), &acceleratorSize);
		textLength = textSize.cx;
		acceleratorLength = acceleratorSize.cx;
	}
	else
	{
		BOOL retText = GetTextExtentPoint32(hdc, label, label.GetLength(), &textSize);
		textLength = textSize.cx;
	}

	SelectObject(hdc, hfntOld);
	::ReleaseDC(hWnd, hdc);
}

BOOL MyPopupMenu::GetLabelParts(CString& label, CString& text, CString& accelerator)
{
	BOOL hasTab = FALSE;
	text.Empty();
	accelerator.Empty();
	int tabPos = label.FindOneOf(L"\t");
	if (tabPos >= 0)
	{
		text = label.Mid(0, tabPos);
		accelerator = label.Mid(tabPos);
		hasTab = TRUE;
	}
	return hasTab;
}
