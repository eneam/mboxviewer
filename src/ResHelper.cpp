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
#include <algorithm>
#include "FileUtils.h"
#include "mboxview.h"
#include "ResHelper.h"
#include "TextUtilsEx.h"

ResInfoMapType ResHelper::resInfoMap(2000);
ResInfoArrayType ResHelper::resInfoArray;

int ResHelper::maxWindowTextLength = WINDOW_TEXT_LENGTH;
wchar_t ResHelper::windowText[WINDOW_TEXT_LENGTH + 1];
int ResHelper::maxClassNameLength = CLASS_NAME_LENGTH;
wchar_t ResHelper::className[CLASS_NAME_LENGTH + 1];

CString ResHelper::resourceFile = L"resource.txt";
HANDLE ResHelper::hResourceFile = INVALID_HANDLE_VALUE;

// Iterates all child windows in the window defined by hwndParent window

void ResHelper::LoadDialogItemsInfo(HWND hwndParent)
{
#ifdef _DEBUG
	if (hwndParent == 0)
		return;

	::GetWindowText(hwndParent, windowText, maxWindowTextLength);

	AddItemInfo(CString(windowText), CString(L"DIALOGEX"));

	int maxcnt = 512;
	int iter = 1;
	LoadDialogItemsInfo(hwndParent, maxcnt, iter);

#endif
}

void ResHelper::LoadDialogItemsInfo(HWND hwndParent, int maxcnt, int iter)
{
#ifdef _DEBUG
	CString strClassName;
	HWND hwndChild = GetWindow(hwndParent, GW_CHILD);
	while ((hwndChild != NULL) && maxcnt > 0)
	{
		int slen = ::GetWindowText(hwndChild, windowText, maxWindowTextLength);
		int sclasslen = GetClassName(hwndChild, className, maxClassNameLength);

		strClassName = className;
		if (strClassName.CompareNoCase(L"EDIT"))
			AddItemInfo(CString(windowText), CString(className));

		UINT id = ::GetDlgCtrlID(hwndChild);

		LoadDialogItemsInfo(hwndChild, --maxcnt, ++iter);
		hwndChild = GetWindow(hwndChild, GW_HWNDNEXT);
	}
#endif
}

void ResHelper::LoadMenuItemsInfo(CMenu* menu, int index)
{
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
		retval = menu->GetMenuString(i, label, nFlags);
		retval = GetMenuItemString(menu, i, label, nFlags);

		memset(&menuItemInfo, 0, sizeof(menuItemInfo));
		menuItemInfo.cbSize = sizeof(menuItemInfo);
		menuItemInfo.fMask = MIIM_ID | MIIM_DATA;
		BOOL retval = menu->GetMenuItemInfo(i, &menuItemInfo, nFlags);
		if (retval == FALSE)
			int deb = 1; // ignore/return;


		if (itemID == (UINT)0)
		{
			; // MENUITEM SEPARATOR
		}
		else if (itemID == (UINT)-1)
		{
			AddItemInfo(label, CString(L"POPUP"));
		}
		else
		{
			AddItemInfo(label, CString(L"MENUITEM"));
		}

		if (itemID == (UINT)-1)
		{
			CMenu* submenu = menu->GetSubMenu(i);

			LoadMenuItemsInfo(submenu, index + 1);
		}
		int deb = 1;
	}
	int debM = 1;
}


void ResHelper::LoadToolBarItemsInfo(CToolBar* tbar)
{

	CToolBarCtrl& cbar = tbar->GetToolBarCtrl();
	int count = cbar.GetButtonCount();

	CString label;
	ResourceInfo* rinfo;
	int i;

	for (i = 0; i < count; i++)
	{
		tbar->GetButtonText(i, label);
		rinfo = AddItemInfo(label, CString(L"TOOLBARITEM"));
		int deb = 1;
	}
	int debM = 1;
}


ResourceInfo* ResHelper::AddItemInfo(CString& label)
{
	CString controlName;

	ResourceInfo* rinfo = ResHelper::AddItemInfo(label, controlName);

	return rinfo;
}

ResourceInfo* ResHelper::AddItemInfo(CString & label, CString& controlName)
{
	CString key;
	key.Format(L"\"%s\"", label);
	hashsum_t hsum = resInfoMap.hash(&key);

	ResourceInfo *rinfo = resInfoMap.find(&key, hsum);
	if (rinfo == 0)
	{
		ResourceInfo* rinfo = new ResourceInfo(key, controlName);
		resInfoMap.insert(hsum, rinfo);
		resInfoArray.Add(rinfo);
	}

	return rinfo;
}

int ResHelper::PrintResInfoMap(ResInfoMapType& resInfoMap)
{
	ResInfoMapType::IHashMapIter iter = resInfoMap.first();
	for (; !resInfoMap.last(iter); resInfoMap.next(iter))
	{
		TRACE(L"%8s %s\n", iter.element->m_controlName, iter.element->m_label);
	}
	return resInfoMap.count();
}

int ResHelper::PrintResInfoArray(ResInfoArrayType& resInfoArray)
{
	CString errorTxt;
	CString resInfoTxt;
	CStringA resInfoTxtA;
	DWORD nNumberOfBytesToWrite;
	DWORD nNumberOfBytesWritten;
	int retval = 0;

	CString processPath;
	CmboxviewApp::GetProcessPath(processPath);
	CString folderPath;
	FileUtils::GetFolderPath(processPath, folderPath);
	CString resourceFilePath = folderPath + L"\\" + resourceFile;

	hResourceFile = FileUtils::FileOpen(resourceFilePath, errorTxt, TRUE);
#if 1
	for (int i = 0; i < resInfoArray.GetCount(); i++)
	{
		ResourceInfo* rinfo = resInfoArray[i];
		TRACE(L"%8s %s\n", rinfo->m_controlName, rinfo->m_label);
		resInfoTxt.Format(L"%s\r\n", rinfo->m_label);

		DWORD error;
		BOOL retW2A = TextUtilsEx::WStr2UTF8(&resInfoTxt, &resInfoTxtA, error);

		if (hResourceFile != INVALID_HANDLE_VALUE)
		{
			nNumberOfBytesToWrite = resInfoTxtA.GetLength();
			nNumberOfBytesWritten = 0;
			retval = FileUtils::Write2File(hResourceFile, (LPCSTR)resInfoTxtA, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
			int deb = 1;
		}
	}
#else
	const char* BOM_UTF16_LE = "\xFF\xFE";
	if (hResourceFile)
	{
		nNumberOfBytesToWrite = 2;
		nNumberOfBytesWritten = 0;

		retval = FileUtils::Write2File(hResourceFile, BOM_UTF16_LE, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
	}

	for (int i = 0; i < resInfoArray.GetCount(); i++)
	{
		ResourceInfo* rinfo = resInfoArray[i];
		TRACE(L"%8s %s\n", rinfo->m_controlName, rinfo->m_label);
		resInfoTxt.Format(L"\"%s\"\n", rinfo->m_label);

		if (hResourceFile != INVALID_HANDLE_VALUE)
		{
			nNumberOfBytesToWrite = resInfoTxt.GetLength()*2;
			nNumberOfBytesWritten = 0;
			retval = FileUtils::Write2File(hResourceFile, (LPCWSTR)resInfoTxt, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
			int deb = 1;
		}
	}
#endif
	FileUtils::FileClose(hResourceFile);

	return (int)resInfoArray.GetCount();
}

int __cdecl ResourceInfoPred(void const* first, void const* second)
{
	ResourceInfo* f = *((ResourceInfo**)first);
	ResourceInfo* s = *((ResourceInfo**)second);

	//int ret = _wcsicmp(f->m_label.operator LPCWSTR(), s->m_label.operator LPCWSTR());
	int ret = f->m_label.CompareNoCase(s->m_label);
	return ret;
}

int ResHelper::SortResInfoArray(ResInfoArrayType& resInfoArray)
{
	std::qsort(resInfoArray.GetData(), resInfoArray.GetCount(), sizeof(void*), ResourceInfoPred);
	return (int)resInfoArray.GetCount();
}

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
			TRACE(L"%3d %6d %s [[%s]]\n", iter, id, className, text);
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
			TRACE(L"MENU Label: %d %d %ld [[%s]]\n", index, itemID, menuItemInfo.dwItemData, str);
		}
		else
		{
			TRACE(L"%d %d %ld [[%s]]\n", index, itemID, menuItemInfo.dwItemData, str);
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
			TRACE(L"Ex MENU Label: %d %d %ld [[[%s]]]\n", index, itemID, menuItemInfo.dwItemData, str);
		}
		else
		{
			TRACE(L"Ex %d %d %ld [[[%s]]]\n", index, itemID, menuItemInfo.dwItemData, str);
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


void ResHelper::ReleaseResources()
{
	resInfoMap.clear();

	for (int i = 0; i < resInfoArray.GetCount(); i++)
	{
		ResourceInfo* rinfo = resInfoArray[i];
		delete rinfo;
		resInfoArray[i] = 0;
	}
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

