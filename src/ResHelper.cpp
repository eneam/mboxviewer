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
#include "ResHelper.h"



// Iterates all child windows in the window defined by hwndParent window

void ResHelper::IterateWindowChilds(HWND hwndParent)
{
#ifdef _DEBUG
	int textlen = 1023;
	wchar_t text[1024];

	if (hwndParent == 0)
		return;

	::GetWindowText(hwndParent, text, textlen);
	TRACE(L"DIALOG list of windows: %s\n", text);

	int maxcnt = 512;
	int iter = 1;
	IterateWindowChilds(hwndParent, maxcnt, iter);
	TRACE(L"DIALOG list of windows end: %s\n", text);
#endif
}

void ResHelper::IterateWindowChilds(HWND hwndParent, int maxcnt, int iter)
{
	int textlen = 1023;
	wchar_t* text = new wchar_t[1024];

	int classNameLen = 1023;
	wchar_t* className = new wchar_t[1024];

	HWND hwndChild = GetWindow(hwndParent, GW_CHILD);
	while ((hwndChild != NULL) && maxcnt > 0)
	{
		int slen = ::GetWindowText(hwndChild, text, textlen);
		int sclasslen = GetClassName(hwndChild, className, classNameLen);


		UINT id = ::GetDlgCtrlID(hwndChild);

		//if (slen > 0)
		{
			TRACE(L"%3d %6d %s %s\n", iter, id, className, text);
		}
		//ProcessDialogItem(hwndChild);

		IterateWindowChilds(hwndChild, --maxcnt, ++iter);
		hwndChild = GetWindow(hwndChild, GW_HWNDNEXT);
	}
	delete[] text;
	delete[] className;
}

BOOL ResHelper::GetMenuItemString(CMenu* menu, UINT nIDItem, CString& rString, UINT nFlags)
{
	MENUITEMINFO menuItemInfo;
	memset(&menuItemInfo, 0, sizeof(menuItemInfo));
	rString.Empty();

	menuItemInfo.cbSize = sizeof(menuItemInfo);
	menuItemInfo.fMask = MIIM_STRING;
	menuItemInfo.fType = MFT_STRING;

	menuItemInfo.dwTypeData = 0;

	// Get the required string length
	menuItemInfo.cch = 0;
	BOOL retval = menu->GetMenuItemInfo(nIDItem, &menuItemInfo, nFlags);
	if (retval == FALSE)
		return FALSE;

	menuItemInfo.cch += 1;
	wchar_t* nametext = new wchar_t[menuItemInfo.cch];
	menuItemInfo.dwTypeData = nametext;

	retval = menu->GetMenuItemInfo(nIDItem, &menuItemInfo, nFlags);
	if (retval == FALSE)
	{
		delete[] nametext;
		return FALSE;
	}

	rString.Append(nametext);
	delete[] nametext;
	return TRUE;
}

void ResHelper::IterateMenuItems(CMenu* menu)
{

	//BOOL GetMenuInfo(LPMENUINFO lpcmi) const;
	//CMenu* GetSubMenu(int nPos) const;

	UINT uItem = 0;
	MENUITEMINFO menuItemInfo;
	LPMENUITEMINFO lpMenuItemInfo = &menuItemInfo;
	memset(&menuItemInfo, 0, sizeof(menuItemInfo));

	menuItemInfo.cbSize = sizeof(menuItemInfo);
	menuItemInfo.fMask = MIIM_STRING | MIIM_ID;
	menuItemInfo.fMask = MIIM_STRING | MIIM_TYPE;
	menuItemInfo.fType = MFT_STRING;

	menuItemInfo.dwTypeData = 0;
	menuItemInfo.dwItemData = 0;


	BOOL fByPos = TRUE;

	BOOL retM = menu->GetMenuItemInfo(uItem, lpMenuItemInfo, fByPos);

	menuItemInfo.fMask = MIIM_STRING | MIIM_ID;
	wchar_t s[1024];
	menuItemInfo.dwTypeData = &s[0];
	menuItemInfo.dwItemData = 0;
	retM = menu->GetMenuItemInfo(uItem, lpMenuItemInfo, fByPos);

	uItem = ID_FILE_OPEN;
	fByPos = FALSE;
	retM = menu->GetMenuItemInfo(uItem, lpMenuItemInfo, fByPos);
}

void ResHelper::IterateMenuItems(CMenu* menu, int index)
{
	int count = menu->GetMenuItemCount();
	CString str;
	UINT itemID = 0;
	int retval;
	UINT nFlags = MF_BYPOSITION;
	int i;
	//HWND h;
	MENUITEMINFO menuItemInfo;
	menuItemInfo.cbSize = sizeof(menuItemInfo);


	//TRACE(L"MENU  list\n");
	for (i = 0; i < count; i++)
	{
		itemID = menu->GetMenuItemID(i);
		retval = menu->GetMenuString(i, str, nFlags);
		retval = GetMenuItemString(menu, i, str, nFlags);

		memset(&menuItemInfo, 0, sizeof(menuItemInfo));
		menuItemInfo.cbSize = sizeof(menuItemInfo);
		menuItemInfo.fMask = MIIM_ID | MIIM_DATA;
		BOOL retval = menu->GetMenuItemInfo(i, &menuItemInfo, nFlags);
		if (retval == FALSE)
			int deb = 1; // return;

		if (itemID == (UINT)-1)
		{
			TRACE(L"MENU Label: %d %d %ld %s\n", index, itemID, menuItemInfo.dwItemData, str);
		}
		else
		{
			TRACE(L"%d %d %ld %s\n", index, itemID, menuItemInfo.dwItemData, str);
		}

		if (itemID == (UINT)-1)
		{
			CMenu* submenu = menu->GetSubMenu(i);

			IterateMenuItems(submenu, index + 1);
		}

		int deb = 1;
	}
	//TRACE(L"MENU END  list\n");

	int debM = 1;
}

void ResHelper::IterateMenuItemsSetPopMenuData(CMenu* menu, int index)
{
	int count = menu->GetMenuItemCount();
	CString str;
	UINT itemID = 0;
	int retval;
	UINT nFlags = MF_BYPOSITION;
	int i;
	//HWND h;
	ULONG_PTR dataVal = 1;

	MENUITEMINFO menuItemInfo;
	menuItemInfo.cbSize = sizeof(menuItemInfo);

	//TRACE(L"MENU  list\n");
	for (i = 0; i < count; i++)
	{
		itemID = menu->GetMenuItemID(i);
		retval = menu->GetMenuString(i, str, nFlags);
		retval = GetMenuItemString(menu, i, str, nFlags);

		memset(&menuItemInfo, 0, sizeof(menuItemInfo));
		menuItemInfo.cbSize = sizeof(menuItemInfo);
		menuItemInfo.fMask = MIIM_ID | MIIM_DATA;
		BOOL retval = menu->GetMenuItemInfo(i, &menuItemInfo, nFlags);
		if (retval == FALSE)
			int deb = 1; // return;

		//if (itemID == (UINT)-1)
		{
			memset(&menuItemInfo, 0, sizeof(menuItemInfo));
			menuItemInfo.cbSize = sizeof(menuItemInfo);
			menuItemInfo.dwItemData = dataVal++;

			menuItemInfo.fMask = MIIM_DATA;
			retval = menu->SetMenuItemInfo(i, &menuItemInfo, nFlags);
			if (retval == FALSE)
				int deb = 1; // return;

			menuItemInfo.dwItemData = 0;

			menuItemInfo.fMask = MIIM_DATA;
			retval = menu->GetMenuItemInfo(i, &menuItemInfo, nFlags);
			if (retval == FALSE)
				int deb = 1; // return;
		}


		if (itemID == (UINT)-1)
		{
			TRACE(L"Ex MENU Label: %d %d %ld %s\n", index, itemID, menuItemInfo.dwItemData, str);
		}
		else
		{
			TRACE(L"Ex %d %d %ld %s\n", index, itemID, menuItemInfo.dwItemData, str);
		}

		if (itemID == (UINT)-1)
		{
			CMenu* submenu = menu->GetSubMenu(i);

			IterateMenuItemsSetPopMenuData(submenu, index + 1);
		}

		int deb = 1;
	}
	//TRACE(L"MENU END  list\n");

	int debM = 1;
}

#if 0

{
	CMenu* menu = pFrame->GetMenu();

	//BOOL GetMenuInfo(LPMENUINFO lpcmi) const;
	//CMenu* GetSubMenu(int nPos) const;

	UINT uItem = 0;
	MENUITEMINFO menuItemInfo;
	LPMENUITEMINFO lpMenuItemInfo = &menuItemInfo;
	memset(&menuItemInfo, 0, sizeof(menuItemInfo));

	menuItemInfo.cbSize = sizeof(menuItemInfo);
	menuItemInfo.fMask = MIIM_STRING | MIIM_ID;
	menuItemInfo.fMask = MIIM_STRING | MIIM_TYPE;
	menuItemInfo.fType = MFT_STRING;

	menuItemInfo.dwTypeData = 0;
	menuItemInfo.dwItemData = 0;


	BOOL fByPos = TRUE;

	BOOL retM = menu->GetMenuItemInfo(uItem, lpMenuItemInfo, fByPos);

	menuItemInfo.fMask = MIIM_STRING | MIIM_ID;
	wchar_t s[1024];
	menuItemInfo.dwTypeData = &s[0];
	menuItemInfo.dwItemData = 0;
	retM = menu->GetMenuItemInfo(uItem, lpMenuItemInfo, fByPos);

	uItem = ID_FILE_OPEN;
	fByPos = FALSE;
	retM = menu->GetMenuItemInfo(uItem, lpMenuItemInfo, fByPos);

	int count = menu->GetMenuItemCount();
	CString str;
	UINT itemID = 0;

	for (int i = 0; i < count; i++)
	{
		itemID = menu->GetMenuItemID(i);
		retval = menu->GetMenuString(i, str, MF_BYPOSITION);

		UINT nFlags = MF_STRING | MF_POPUP;
		UINT nIDNewItem = itemID;
		CString newItem = L"Override";
		LPCWSTR lp = (LPCWSTR)newItem;
		UINT_PTR ptrItem = (UINT_PTR)&nIDNewItem;
		//BOOL retMod = menu->ModifyMenu((UINT)i, nFlags, ptrItem, lp);

		retval = menu->GetMenuString(i, str, MF_BYPOSITION);

		if (itemID == (UINT)-1)
		{
			nFlags = MF_STRING | MF_BYPOSITION;
			nIDNewItem = itemID;
			UINT_PTR ptrItem = (UINT_PTR)&nIDNewItem;
			BOOL retMod = menu->ModifyMenu((UINT)i, nFlags, ptrItem, lp);

			CMenu* submenu = menu->GetSubMenu(i);
			int count = submenu->GetMenuItemCount();

			for (int ii = 0; ii < count; ii++)
			{
				itemID = submenu->GetMenuItemID(ii);

				if (itemID == 0)
					continue;

				if (itemID == (UINT)-1)
					continue;


				retval = submenu->GetMenuString(ii, str, MF_BYPOSITION);
				//retval = submenu->GetMenuString(itemID, str, MF_BYCOMMAND);
				//
				retM = submenu->GetMenuItemInfo(ii, lpMenuItemInfo, TRUE);
				//retM = submenu->GetMenuItemInfo(itemID, lpMenuItemInfo, MF_BYCOMMAND);

				nFlags = MF_STRING | MF_BYPOSITION;
				nIDNewItem = itemID;
				UINT_PTR ptrItem = (UINT_PTR)&nIDNewItem;
				BOOL retMod = submenu->ModifyMenu((UINT)ii, nFlags, ptrItem, lp);
				if (retMod == FALSE)
				{
					CString errTxt = FileUtils::GetLastErrorAsString();
					int deb = 1;
				}
				else
				{
					retval = submenu->GetMenuString(ii, str, MF_BYPOSITION);
					int deb = 1;
				}
			}
		}
		else if (itemID != 0)
		{
			itemID = menu->GetMenuItemID(i);
			retval = menu->GetMenuString(i, str, MF_BYPOSITION);

			nFlags = MF_STRING | MF_BYPOSITION;
			nIDNewItem = itemID;
			UINT_PTR ptrItem = (UINT_PTR)&nIDNewItem;
			BOOL retMod = menu->ModifyMenu((UINT)i, nFlags, ptrItem, lp);
		}
		else
		{
			int deb = 1;
		}

		int deb = 1;
	}

	pFrame->UpdateWindow();
	pFrame->DrawMenuBar();

	int debM = 1;
}
#endif

