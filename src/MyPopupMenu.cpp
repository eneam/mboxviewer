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
//
int MyPopupMenu::m_menuBarFontSize = 11;
CFont MyPopupMenu::m_menuBarFont;
LOGFONT MyPopupMenu::m_menuBarLogFont;
//
BOOL MyPopupMenu::m_isCustomFont = FALSE;


UINT  MyPopupMenu::MenuItemInfoMaskAllSet = MIIM_STATE | MIIM_ID | MIIM_SUBMENU | MIIM_CHECKMARKS | MIIM_TYPE | MIIM_DATA | MIIM_STRING | MIIM_BITMAP | MIIM_FTYPE;
UINT  MyPopupMenu::MenuItemInfoMaskFTypeAllSet = MIIM_STATE | MIIM_ID | MIIM_SUBMENU | MIIM_CHECKMARKS /* | MIIM_TYPE */ | MIIM_DATA | MIIM_STRING | MIIM_BITMAP | MIIM_FTYPE;
//UINT  MyPopupMenu::MenuItemInfoMaskTypeAllSet = MIIM_STATE | MIIM_ID | MIIM_SUBMENU | MIIM_CHECKMARKS | MIIM_TYPE | MIIM_DATA | MIIM_STRING | MIIM_BITMAP /* | MIIM_FTYPE */;
UINT  MyPopupMenu::MenuItemInfoMaskTypeAllSet = MIIM_STATE | MIIM_ID | MIIM_SUBMENU | MIIM_CHECKMARKS | MIIM_TYPE | MIIM_DATA ;

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
	CFont font;   // Remember to remove since it is not used
	if (m_fontSize == 0)
	{
		m_fontSize = CMainFrame::m_dfltFontSize;
		m_isCustomFont = IsCustomFont(MyPopupMenu::m_fontSize);
		if (m_isCustomFont)
			MyPopupMenu::m_fontSize = CMainFrame::CMainFrame::m_cnfFontSize;

		m_font.DeleteObject();
		MyPopupMenu::GetMenuFont(hdc, font, MyPopupMenu::m_fontSize, MyPopupMenu::m_menuLogFont, MyPopupMenu::m_font);

		// Sould never change since I was not able to resize menu bar
		m_menuBarFont.DeleteObject();
		MyPopupMenu::GetMenuFont(hdc, font, MyPopupMenu::m_menuBarFontSize, MyPopupMenu::m_menuBarLogFont, MyPopupMenu::m_menuBarFont);
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
	_ASSERTE(lpDrawItemStruct);
	if (!lpDrawItemStruct)
		return;
	if (lpDrawItemStruct->CtlType != ODT_MENU)
		return;


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
	_ASSERTE(lpDrawItemStruct);
	if (!lpDrawItemStruct)
		return;

	_ASSERTE(nIDCtl == lpDrawItemStruct->CtlID);
	if (lpDrawItemStruct->CtlType != ODT_MENU)
		return;

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
	if (pMyItem->m_isMenuBarItem)
	{
		menuCheckSize = 0;
		textOffset = 16;
	}

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
	CString itemText = pMyItem->m_text;
	itemText.Replace(L"&", L"");
	CString text;
	CString accelerator;
	BOOL hasTab = MyPopupMenu::GetLabelParts(itemText, text, accelerator);

	if (!hasTab)
	{
		ExtTextOut(lpDrawItemStruct->hDC, x, y, ETO_OPAQUE,
			&rcItem, itemText, itemText.GetLength(), NULL);

		TRACE(L"OnDrawItem:Label = \"%s\" \"%s\" cx=%d cy=%d rec.width=%d rec.hight=%d\n", 
			pMyItem->m_text, itemText, x, y, rcItem.Width(), rcItem.Height());
	}
	else
	{
#if 0
		// consider using DC
		// Compare dc.DrawText with ExtTextOut for suuport multiple langauges
		dc.DrawText(left, &rc, DT_MYSTANDARD);
		if (!hasTab)
			dc.DrawText(right, &rc, DT_MYSTANDARD | DT_RIGHT);
#endif
		ExtTextOut(lpDrawItemStruct->hDC, x, y, ETO_OPAQUE,
			&rcItem, text, text.GetLength(), NULL);

		TRACE(L"OnDrawItem:Label Left = \"%s\" cx=%d cy=%d rec.width=%d rec.hight=%d\n", text, x, y, rcItem.Width(), rcItem.Height());

		CRect rc = rcItem;
		rc.left += pMyItem->m_maxTextLeftPartLengthInPoints + menuCheckSize;
		x += pMyItem->m_maxTextLeftPartLengthInPoints + menuCheckSize;

		ExtTextOut(lpDrawItemStruct->hDC, x, y, ETO_OPAQUE,
				&rc, accelerator, accelerator.GetLength(), NULL);

		TRACE(L"OnDrawItem:Label Right = \"%s\" cx=%d cy=%d rec.width=%d rec.hight=%d\n", accelerator, x, y, rc.Width(), rc.Height());
	}

	// Select the font and draw the check mark character.
	CString checkMark;
	if (lpDrawItemStruct->itemState & ODS_CHECKED)
		checkMark = L"\u2713";

	RECT rc = rcItem;
	rc.right = GetSystemMetrics(SM_CXMENUCHECK);
	rc.bottom -= 8;  // Find character hight for adjustment
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

	if (pMyItem->m_drawSeparator && !pMyItem->m_isMenuBarItem)
	{
		CDC dc;
		dc.Attach(lpDrawItemStruct->hDC);
		CRect rc = rcItem;
		rc.left += textOffset;

		UINT nFlags = BF_BOTTOM | BF_FLAT;
		dc.DrawEdge(&rc, EDGE_ETCHED, nFlags);
		dc.Detach();
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
	_ASSERTE(lpMeasureItemStruct);
	if (!lpMeasureItemStruct)
		return;
	if (lpMeasureItemStruct->CtlType != ODT_MENU)
		return;

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
	_ASSERTE(lpMeasureItemStruct);
	if (!lpMeasureItemStruct)
		return;
	if (lpMeasureItemStruct->CtlType != ODT_MENU)
		return;

	MyPopupMenuItem* pMyItem = (MyPopupMenuItem*)lpMeasureItemStruct->itemData;

	if (pMyItem->m_fType & MFT_SEPARATOR)
	{
		lpMeasureItemStruct->itemHeight = GetSystemMetrics(SM_CYMENU) >> 1;
		TRACE(L"OnMeasureItem:Separator = \"%s\" cx=%d cy=%d\n", pMyItem->m_text, lpMeasureItemStruct->itemWidth, lpMeasureItemStruct->itemHeight);
		return;
	}

	HDC hdc = GetDC(hWnd);
	HFONT hfntOld = (HFONT)SelectObject(hdc, pMyItem->m_hfont);
	SIZE size;

	GetTextExtentPoint32(hdc, pMyItem->m_text, pMyItem->m_text.GetLength(), &size);

	SelectObject(hdc, hfntOld);
	ReleaseDC(hWnd, hdc);

	int textOffset = (int)((float)3 * (float)GetSystemMetrics(SM_CXMENUCHECK));
	if (pMyItem->m_isMenuBarItem)
		textOffset = 4;

	lpMeasureItemStruct->itemWidth = pMyItem->m_maxTextLeftPartLengthInPoints + pMyItem->m_maxTextRightPartLengthInPoints + textOffset;
	lpMeasureItemStruct->itemHeight = size.cy + 8;

	TRACE(L"OnMeasureItem:Label = \"%s\" cx=%d cy=%d\n", pMyItem->m_text, lpMeasureItemStruct->itemWidth, lpMeasureItemStruct->itemHeight);

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
	// Must always execute
	if (!m_isCustomFont)
		return;

	HMENU hM = menu->m_hMenu;
	const HMENU hMenu = menu->GetSafeHmenu();
	BOOL isMenu = ::IsMenu(hMenu);

	_ASSERTE(isMenu);
	if (!isMenu)
		return;

	BOOL hasTab = FALSE;

	MyPopupMenuItem* myLastItem = 0;

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
		TRACE(L"SetCMenuAsCustom:Label = \"%s\" MaxLeft=%d maxRight=%d\n", label, maxTextLeftPartLengthInPoints, maxTextRightPartLengthInPoints);

		//BOOL retTraceFType = MyPopupMenu::TraceMenuItem(CString(L"SetCMenuAsCustom:FType: Original Item Info: "), menu, i, MyPopupMenu::MenuItemInfoMaskFTypeAllSet);
		//BOOL retTraceType = MyPopupMenu::TraceMenuItem(CString(L"SetCMenuAsCustom:Type: Original Item Info: "), menu, i, MyPopupMenu::MenuItemInfoMaskTypeAllSet);

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

		BOOL isMenuBarItem = FALSE;
		int textLeftPartLengthInPoints = 0;  // true text
		int textRightPartLengthInPoints = 0;  // accelerators
		if ((label.Compare(fileMenuItem) == 0) ||
			(label.Compare(editMenuItem) == 0) ||
			(label.Compare(viewMenuItem) == 0) ||
			(label.Compare(helpMenuItem) == 0) ||
			(label.Compare(languageMenuItem) == 0)
				)
		{
			HFONT hf = MyPopupMenu::m_menuBarFont.operator HFONT();

			HWND hWnd = 0;
			MyPopupMenu::GetLengthOfMenuLabelPartsInPoints(hWnd, hf, label, textLeftPartLengthInPoints, textRightPartLengthInPoints);
			isMenuBarItem = TRUE;
			int deb = 1; //
		}

		if ((itemID > ID_FILE_MRU_FILE1) && (itemID < (ID_FILE_MRU_FILE1 + 16)))
		{
			int deb = 1;
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
				//delete myItem;
				//continue;
			}

			MyPopupMenuItem* myItem = new MyPopupMenuItem;
			if (menuItemInfo.fType & MFT_SEPARATOR)
			{
				if (myLastItem)
					myLastItem->m_drawSeparator = FALSE;

				myLastItem = 0;
			}
			else
				myLastItem = myItem;

			//if ((menuItemInfo.fType & MFT_SEPARATOR) == FALSE)
			{
				myItem->m_fType = menuItemInfo.fType;
				CString text = label;

				//ResHelper::TranslateString(text);  // It is already translated
				myItem->m_text = text;
				
				if (!isMenuBarItem)
				{
					myItem->m_hfont = MyPopupMenu::m_font.operator HFONT();
					myItem->m_maxTextLeftPartLengthInPoints = maxTextLeftPartLengthInPoints;
					myItem->m_maxTextRightPartLengthInPoints = maxTextRightPartLengthInPoints;
					myItem->m_drawSeparator = TRUE;
				}
				else
				{
					myItem->m_hfont = MyPopupMenu::m_menuBarFont.operator HFONT();
					myItem->m_maxTextLeftPartLengthInPoints = textLeftPartLengthInPoints;
					myItem->m_maxTextRightPartLengthInPoints = textRightPartLengthInPoints;
					myItem->m_drawSeparator = FALSE;

				}
				myItem->m_isMenuBarItem = isMenuBarItem;

				//menuItemInfo.fMask = MIIM_DATA | MIIM_FTYPE;
				//menuItemInfo.fMask = MIIM_ID | MIIM_DATA | MIIM_TYPE;  // FTYPE -> TYPE
				menuItemInfo.fMask = MIIM_DATA | MIIM_FTYPE;  // FTYPE -> TYPE
				menuItemInfo.fType |= MFT_OWNERDRAW;
				menuItemInfo.dwItemData = (ULONG_PTR)myItem;

				BOOL retSet = ((CMenu*)menu)->SetMenuItemInfo(i, &menuItemInfo, TRUE);
				_ASSERT(retSet);
			}
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

// Applies to  menubar
void MyPopupMenu::RestoreCMenu(CMenu* menu, int index)
{
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

	wchar_t labelText[256];

	for (i = 0; i < count; i++)
	{
		itemID = menu->GetMenuItemID(i);

		if (itemID == 0)
			int deb = 1; // continue;

		int retGetStr = menu->GetMenuString(i, label, nFlags);
#if 0
		// Currently labels are already translated, no need to translate again
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
#endif
		{
			memset(&menuItemInfo, 0, sizeof(menuItemInfo));
			menuItemInfo.cbSize = sizeof(menuItemInfo);
			labelText[0] = 0;
			menuItemInfo.dwTypeData = &labelText[0];
			menuItemInfo.cch = 255;
			//menuItemInfo.fMask = MIIM_ID | MIIM_FTYPE | MIIM_DATA;
			menuItemInfo.fMask = MIIM_SUBMENU | MIIM_DATA | MIIM_ID | MIIM_TYPE;
			BOOL retGetItemData = menu->GetMenuItemInfo(i, &menuItemInfo, TRUE);
			_ASSERTE(retGetItemData);
			if (retGetItemData == FALSE)
				int deb = 1; // ignore/return;

			//_ASSERTE(menuItemInfo.fType & MFT_OWNERDRAW);
			if (menuItemInfo.fType & MFT_OWNERDRAW) 	// if ownerdraw:
			{
				MyPopupMenuItem* myItem = (MyPopupMenuItem*)menuItemInfo.dwItemData;
				menuItemInfo.fType &= ~MFT_OWNERDRAW;		//   turn it off

				menuItemInfo.fMask |= MIIM_TYPE;  // TYPE -> FSTYPE
				menuItemInfo.dwTypeData = (LPWSTR)((LPCWSTR)myItem->m_text);
				menuItemInfo.cch = myItem->m_text.GetLength();

				BOOL deleteMyItem = FALSE;
				if (menuItemInfo.dwItemData)
				{
					MyPopupMenuItem* myItem = (MyPopupMenuItem*)menuItemInfo.dwItemData;
					menuItemInfo.dwItemData = NULL;				// item data is NULL
					menuItemInfo.fMask |= MIIM_DATA;			// change it
					deleteMyItem = TRUE;
				}

				BOOL retSet = ((CMenu*)menu)->SetMenuItemInfo(i, &menuItemInfo, TRUE);
				_ASSERT(retSet);

				if (deleteMyItem)
					delete myItem;
			}
			else
			{
				//BOOL retTraceFType = MyPopupMenu::TraceMenuItem(CString(L"RestoreCMenu:FType: Original Item Info: "), menu, i, MyPopupMenu::MenuItemInfoMaskFTypeAllSet);
				//BOOL retTraceType = MyPopupMenu::TraceMenuItem(CString(L"RestoreCMenu:Type: Original Item Info: "), menu, i, MyPopupMenu::MenuItemInfoMaskTypeAllSet);
				//continue;
				int deb = 1;

			}
		}

		if (itemID == (UINT)-1)
		{
			CMenu* submenu = (MyPopupMenu*)menu->GetSubMenu(i);

			_ASSERTE(submenu);
			if (submenu)
				RestoreCMenu(submenu, index + 1);
			int deb = 1;
		}
		int deb = 1;

	}
	TRACE(L"END of RestoreCMenu ++++++++++++++++++++++\n");
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
#if 0
			memset(&menuItemInfo, 0, sizeof(menuItemInfo));
			menuItemInfo.cbSize = sizeof(menuItemInfo);
			menuItemInfo.fMask = MIIM_ID | MIIM_TYPE | MIIM_DATA;
			BOOL retGetItemData = menu->GetMenuItemInfo(i, &menuItemInfo, TRUE);
			_ASSERTE(retGetItemData);
			if (retGetItemData == FALSE)
				int deb = 1; // ignore/return;
#endif

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
			//if (submenu) FindLengthOfLongestText(submenu, index + 1);
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
			// Unless we set dwItemData to 0 after deleting myItem
			//if (submenu) submenu->ReleaseCustomResources(index + 1);
		}
		int deb = 1;
	}
	int deb = 1;
}


// Not used currently. Delete ??
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

void MyPopupMenu::PrintMENUITEMINFO(CString &infoText, MENUITEMINFO& minfo, UINT mask)
{
	CString text;
	//infoText.Append(L"MENUITEMINFO !!!!!!!!!!!!\n");

	text.Format(L" fMask=0%04x", minfo.fMask);
	infoText.Append(text);

	if (mask & MIIM_ID)
	{
		text.Format(L" wID=%u", minfo.wID);
		infoText.Append(text);
	}
	if (mask & MIIM_STATE)
	{
		text.Format(L" fState=%u", minfo.fState);
		infoText.Append(text);
	}
	if (mask & MIIM_SUBMENU)
	{
		text.Format(L" hSubMenu=0x%04x", minfo.hSubMenu);
		infoText.Append(text);
	}
	if (mask & MIIM_CHECKMARKS)
	{
		text.Format(L" Checkmarks=0x%04x 0x%04x", minfo.hbmpChecked, minfo.hbmpUnchecked);
		infoText.Append(text);
	}
	if (mask & MIIM_DATA)
	{
		text.Format(L" dwItemData=0x%08x", minfo.dwItemData);
		infoText.Append(text);
	}
	if (mask & MIIM_STRING)
	{
		CString str(L"Empty");
		if (minfo.dwTypeData)
			str = CString(minfo.dwTypeData, minfo.cch);
		text.Format(L" dwTypeData=0x%08x %u \"%s\"", minfo.dwTypeData, minfo.cch, str);
		infoText.Append(text);
	}
	if (mask & MIIM_BITMAP)
	{
		text.Format(L" hbmpItem=0x%08x", minfo.hbmpItem);
		infoText.Append(text);
	}
	if (mask & MIIM_TYPE)
	{
		//if (minfo.fType == MFT_STRING)
		{
			CString str(L"Empty");
			if (minfo.dwTypeData)
				str = CString(minfo.dwTypeData, minfo.cch);
			text.Format(L" dwTypeData=0x%08x %u \"%s\"", minfo.dwTypeData, minfo.cch, str);
			infoText.Append(text);
		}
		text.Format(L" hbmpItem=0x%08x", minfo.hbmpItem);
		infoText.Append(text);

		text.Format(L" Type=%u", minfo.fType);
		infoText.Append(text);
		if (minfo.fType & MFT_OWNERDRAW)
		{
			infoText.Append(L" MFT_OWNERDRAW(0x100)");
		}
		if (minfo.fType & MFT_SEPARATOR)
		{
			infoText.Append(L" MFT_SEPARATOR(800)");
		}
		if (minfo.fType == MFT_STRING)
		{
			infoText.Append(L" MFT_STRING(0)");
		}
	}
	if (mask & MIIM_FTYPE)
	{
		text.Format(L" fType=%u", minfo.fType);
		infoText.Append(text);
		if (minfo.fType & MFT_OWNERDRAW)
		{
			infoText.Append(L" MFT_OWNERDRAW(0x100)");
		}
		if (minfo.fType & MFT_SEPARATOR)
		{
			infoText.Append(L" MFT_SEPARATOR(800)");
		}
		if (minfo.fType == MFT_STRING)
		{
			infoText.Append(L" MFT_STRING(0)");
		}
	}
	infoText.Append(L"\n");
}

BOOL MyPopupMenu::TraceMenuItem(CString &infoText, CMenu *menu, int index, UINT mask)
{
	BOOL retval = TRUE;
#ifdef _DEBUG
	CString errorText;
	wchar_t mstr[256];

	MENUITEMINFO menuItemInfo;
	memset(&menuItemInfo, 0, sizeof(menuItemInfo));
	menuItemInfo.cbSize = sizeof(menuItemInfo);
	menuItemInfo.fMask = mask;
	menuItemInfo.dwTypeData = &mstr[0]; mstr[0] = 0;
	menuItemInfo.cch = 255;
	retval = menu->GetMenuItemInfo(index, &menuItemInfo, TRUE);
	if (retval)
	{
		MyPopupMenu::PrintMENUITEMINFO(infoText, menuItemInfo, mask);
	}
	else
	{
		errorText = FileUtils::GetLastErrorAsString();
		infoText.Append(errorText);
		retval = FALSE;
	}
	TRACE(L"%s\n", infoText);
#endif
	return retval;
}


// Main menubar
BOOL MyPopupMenu::TraceMenu(CString& title, CMenu* menu, int index, UINT mask)
{
#ifdef _DEBUG
	HMENU hM = menu->m_hMenu;
	const HMENU hMenu = menu->GetSafeHmenu();
	BOOL isMenu = ::IsMenu(hMenu);

	_ASSERTE(isMenu);
	if (!isMenu)
		return FALSE;

	TRACE(L"BEGIN of TraceMenu ++++++++++++++++++++++\n");

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

		int retGetStr = menu->GetMenuString(i, label, nFlags);

		BOOL retTraceFType = MyPopupMenu::TraceMenuItem(label, menu, i, mask);

		if (itemID == (UINT)-1)
		{
			CMenu* submenu = (MyPopupMenu*)menu->GetSubMenu(i);

			_ASSERTE(submenu);
			if (submenu)
				TraceMenu(title, submenu, index + 1, mask);

			int deb = 1;
		}
		int deb = 1;

	}
	TRACE(L"END of TraceMenu ++++++++++++++++++++++\n");
#endif
	return TRUE;
}

