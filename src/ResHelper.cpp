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


//

BOOL ResHelper::g_LoadMenuItemsInfo = FALSE; // Traverse dialog & menu controles and create resource L"resource.txt" file. i.e. collect text from all items/sub-controls
BOOL ResHelper::g_UpdateMenuItemsInfo = TRUE; // Update text of al controls using selected translation file

ResInfoMapType ResHelper::g_LanguageMap(2000);

ResInfoMapType ResHelper::g_resInfoMap(2000);
ResInfoArrayType ResHelper::g_resInfoArray;

ResInfoArrayType ResHelper::resArray1;
ResInfoArrayType ResHelper::resArray2;

int ResHelper::maxWindowTextLength = WINDOW_TEXT_LENGTH;
wchar_t ResHelper::windowText[WINDOW_TEXT_LENGTH + 1];
int ResHelper::maxClassNameLength = CLASS_NAME_LENGTH;
wchar_t ResHelper::className[CLASS_NAME_LENGTH + 1];
CString ResHelper::wndText;

CString ResHelper::resourceFile = L"resource.txt";
HANDLE ResHelper::hResourceFile = INVALID_HANDLE_VALUE;

// Iterates all child windows in the window defined by hwndParent window


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

void ResHelper::DisableLanguageLoading()
{
	g_LoadMenuItemsInfo = FALSE;
	g_UpdateMenuItemsInfo = FALSE;
}

void ResHelper::EnableLanguageLoading()
{
	g_LoadMenuItemsInfo = FALSE;
	g_UpdateMenuItemsInfo = TRUE;
}


void ResHelper::EnableResInfoCollecting()
{
	g_LoadMenuItemsInfo = TRUE;
	g_UpdateMenuItemsInfo = FALSE;
}

void ResHelper::LoadDialogItemsInfo(CDialog* dlg)
{
#ifdef _DEBUG
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_LoadMenuItemsInfo == FALSE)
		return;

	HWND hwndParent = dlg->GetSafeHwnd();

	if (hwndParent == 0)
		return;

	CWnd* wnd = (CWnd*)dlg;

	LoadDialogItemsInfo(wnd);

#endif
}

void ResHelper::LoadDialogItemsInfo(CWnd *wnd)
{
#ifdef _DEBUG
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_LoadMenuItemsInfo == FALSE)
		return;

	HWND hwndParent = wnd->GetSafeHwnd();

	if (hwndParent == 0)
		return;

	::GetWindowText(hwndParent, windowText, maxWindowTextLength);
	wnd->GetWindowText(wndText);

	AddItemInfo(CString(windowText), CString(L"DIALOGEX"));

	int maxcnt = 512;
	int iter = 1;
	LoadDialogItemsInfo(wnd, hwndParent, maxcnt, iter);

#endif
}

void ResHelper::LoadDialogItemsInfo(CWnd *wnd, HWND hwndParent, int maxcnt, int iter)
{
#ifdef _DEBUG
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_LoadMenuItemsInfo == FALSE)
		return;

	CString strClassName;
	CString strWindowText;
	CString strWindowText2;
	HWND hwndChild = GetWindow(hwndParent, GW_CHILD);
	while ((hwndChild != NULL) && maxcnt > 0)
	{
		int slen = ::GetWindowText(hwndChild, windowText, maxWindowTextLength);
		int sclasslen = GetClassName(hwndChild, className, maxClassNameLength);

		strWindowText = windowText;
		strClassName = className;

		UINT nID = ::GetDlgCtrlID(hwndChild);

		CWnd* pw = wnd->GetDlgItem(nID);

		pw->GetWindowText(strWindowText2);

		if (strClassName.CompareNoCase(L"Edit") == 0)
		{
			AddItemInfo(strWindowText, strClassName);
		}
		else if (strClassName.CompareNoCase(L"ListBox") == 0)
		{
			int index = 0;
			CListBox* listBox = (CListBox*)pw;
			LoadListBoxItemsInfo(listBox, index);
			int deb = 1;
		}
		else if (strClassName.CompareNoCase(L"ComboBox") == 0)
		{
			int index = 0;
			CComboBox* comboBox = (CComboBox*)pw;
			LoadComboBoxItemsInfo(comboBox, index);
			int deb = 1;
		}
		else if (strClassName.CompareNoCase(L"Button") == 0)
		{
			AddItemInfo(strWindowText, strClassName);
		}
		else if (strClassName.CompareNoCase(L"Static") == 0)
		{
			AddItemInfo(strWindowText, strClassName);
		}
		else if (strClassName.CompareNoCase(L"SysListView32") == 0)
		{
			AddItemInfo(strWindowText, strClassName);
#if 0
			int index = 0;
			CListBox* listBox = (CListBox*)pw;
			LoadListBoxItemsInfo(listBox, index);
#endif
		}
		else
		{
			AddItemInfo(strWindowText, strClassName);
			LoadDialogItemsInfo(wnd, hwndChild, --maxcnt, ++iter);
		}
		hwndChild = GetWindow(hwndChild, GW_HWNDNEXT);
	}
#endif
}


void ResHelper::UpdateDialogItemsInfo(CDialog* dlg)
{
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_UpdateMenuItemsInfo == FALSE)
		return;

	HWND hwndParent = dlg->GetSafeHwnd();

	if (hwndParent == 0)
		return;

	CWnd* wnd = (CWnd*)dlg;

	UpdateDialogItemsInfo(wnd);
}

void ResHelper::UpdateDialogItemsInfo(CWnd* wnd)
{
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_UpdateMenuItemsInfo == FALSE)
		return;

	HWND hwndParent = wnd->GetSafeHwnd();

	if (hwndParent == 0)
		return;

	::GetWindowText(hwndParent, windowText, maxWindowTextLength);
	wnd->GetWindowText(wndText);

	CString newWndText;
	BOOL retFind = DetermineString(wndText, newWndText);

	if (newWndText.GetLength())
	{
		wnd->SetWindowText(newWndText);
	}
	else
	{
		int deb = 1; // keep old text
	}

	int maxcnt = 512;
	int iter = 1;
	UpdateDialogItemsInfo(wnd, hwndParent, maxcnt, iter);
}

void ResHelper::UpdateDialogItemsInfo(CWnd* wnd, HWND hwndParent, int maxcnt, int iter)
{
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_UpdateMenuItemsInfo == FALSE)
		return;

	CString strClassName;
	HWND hwndChild = GetWindow(hwndParent, GW_CHILD);
	while ((hwndChild != NULL) && maxcnt > 0)
	{
		int slen = ::GetWindowText(hwndChild, windowText, maxWindowTextLength);
		int sclasslen = GetClassName(hwndChild, className, maxClassNameLength);

		CString wndText = windowText;
		strClassName = className;

		CString newWndText;
		BOOL retFind = DetermineString(wndText, newWndText);

		UINT nID = ::GetDlgCtrlID(hwndChild);

		CWnd* pw = wnd->GetDlgItem(nID);
		if (strClassName.CompareNoCase(L"Edit") == 0)
		{
			;// pw->SetWindowText(newWndText);
		}
		else if (strClassName.CompareNoCase(L"ListBox") == 0)
		{
			int index = 0;
			CListBox* listBox = (CListBox*)pw;
			UpdateListBoxItemsInfo(listBox, index);
			int deb = 1;
		}
		else if (strClassName.CompareNoCase(L"ComboBox") == 0)
		{
			int index = 0;
			CComboBox* comboBox = (CComboBox*)pw;
			UpdateComboBoxItemsInfo(comboBox, index);
			int deb = 1;
		}
		else
		{
			if (newWndText.GetLength())
			{
				pw->SetWindowText(newWndText);
			}
			else
			{
				int deb = 0;; // keep old text
			}

			UpdateDialogItemsInfo(pw, hwndChild, --maxcnt, ++iter);
		}
		hwndChild = GetWindow(hwndChild, GW_HWNDNEXT);
	}
}


void ResHelper::LoadListBoxItemsInfo(CListBox* menu, int index)
{
#ifdef _DEBUG
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_LoadMenuItemsInfo == FALSE)
		return;

	ResourceInfo* rinfo;
	int count = menu->GetCount();
	CString label;
	UINT itemID = 0;
	UINT nFlags = MF_BYPOSITION;
	int i;
	MENUITEMINFO menuItemInfo;
	menuItemInfo.cbSize = sizeof(menuItemInfo);

	for (i = 0; i < count; i++)
	{
		menu->GetText(i, label);
		rinfo = AddItemInfo(label, CString(L"LISTBOXITEM"));
		int deb = 1;
	}
	int debM = 1;
#endif
}

void ResHelper::UpdateListBoxItemsInfo(CListBox* menu, int index)
{
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_UpdateMenuItemsInfo == FALSE)
		return;

	int count = menu->GetCount();
	CString label;
	UINT itemID = 0;
	int retval;
	UINT nFlags = MF_BYPOSITION;
	int i;

	// Threre is no SetString() support. Clear current list and add new translated strings
	CStringArray ar;
	for (i = 0; i < count; i++)
	{
		menu->GetText(i, label);
		ar.Add(label);
	}

	_ASSERTE(count == ar.GetCount());

	int currSel = menu->GetCurSel();

	menu->ResetContent();

	CString newLabel;
	for (i = 0; i < count; i++)
	{
		label = ar[i];
		newLabel.Empty();
		BOOL retFind = DetermineString(label, newLabel);
		if (newLabel.GetLength())
		{
			retval = menu->AddString(newLabel);
		}
		else
		{
			retval = menu->AddString(label);  // keep old label/text
		}

		int deb = 1;
	}
	_ASSERTE(count == menu->GetCount());
	
	count = menu->GetCount();

	if ((currSel >= 0) && (currSel < count))
	{
		retval = menu->SetCurSel(currSel);
	}

	int debM = 1;
}

void ResHelper::LoadComboBoxItemsInfo(CComboBox* menu, int index)
{
#ifdef _DEBUG
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_LoadMenuItemsInfo == FALSE)
		return;

	ResourceInfo* rinfo;
	int count = menu->GetCount();
	CString label;
	UINT itemID = 0;
	UINT nFlags = MF_BYPOSITION;
	int i;
	MENUITEMINFO menuItemInfo;
	menuItemInfo.cbSize = sizeof(menuItemInfo);

	for (i = 0; i < count; i++)
	{
		menu->GetLBText(i, label);
		rinfo = AddItemInfo(label, CString(L"COMBOBOXITEM"));
		int deb = 1;
	}
	int debM = 1;
#endif
}


void ResHelper::UpdateComboBoxItemsInfo(CComboBox* menu, int index)
{
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_UpdateMenuItemsInfo == FALSE)
		return;

	int count = menu->GetCount();
	CString label;
	UINT itemID = 0;
	int retval;
	UINT nFlags = MF_BYPOSITION;
	int i;

	// Threre is no SetString() support. Clear current list and add new translated strings
	CStringArray ar;
	for (i = 0; i < count; i++)
	{
		menu->GetLBText(i, label);
		ar.Add(label);
	}

	_ASSERTE(count == ar.GetCount());

	int currSel = menu->GetCurSel();

	menu->ResetContent();

	CString newLabel;
	for (i = 0; i < count; i++)
	{
		label = ar[i];
		newLabel.Empty();
		BOOL retFind = DetermineString(label, newLabel);
		if (newLabel.GetLength())
		{
			retval = menu->AddString(newLabel);
		}
		else
		{
			retval = menu->AddString(label);  // keep old label/text
		}

		int deb = 1;
	}
	_ASSERTE(count == menu->GetCount());

	count = menu->GetCount();

	if ((currSel >= 0) && (currSel < count))
	{
		retval = menu->SetCurSel(currSel);
	}
	int debM = 1;
}


void ResHelper::LoadMenuItemsInfo(CMenu* menu, int index)
{
#ifdef _DEBUG
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_LoadMenuItemsInfo == FALSE)
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
#endif
}

BOOL  ResHelper::DetermineString(CString &str, CString &newString)
{
	hashsum_t hsum = g_LanguageMap.hash(&str);

	ResourceInfo* rinfo = g_LanguageMap.find(&str, hsum);
	if (rinfo)
	{
		newString = rinfo->m_controlName;
		return TRUE;
	}
	return FALSE;
}


void ResHelper::UpdateMenuItemsInfo(CMenu* menu, int index)
{
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_UpdateMenuItemsInfo == FALSE)
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
		retval = menu->GetMenuString(i, label, nFlags);
		retval = GetMenuItemString(menu, i, label, nFlags);

		memset(&menuItemInfo, 0, sizeof(menuItemInfo));
		menuItemInfo.cbSize = sizeof(menuItemInfo);
		menuItemInfo.fMask = MIIM_ID | MIIM_DATA;
		BOOL retval = menu->GetMenuItemInfo(i, &menuItemInfo, nFlags);
		if (retval == FALSE)
			int deb = 1; // ignore/return;

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

		if (itemID == (UINT)-1)
		{
			CMenu* submenu = menu->GetSubMenu(i);

			UpdateMenuItemsInfo(submenu, index + 1);
		}
		int deb = 1;
	}
	int debM = 1;
}


void ResHelper::LoadToolBarItemsInfo(CToolBar* tbar)
{
#ifdef _DEBUG
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_LoadMenuItemsInfo == FALSE)
		return;

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
#endif
}

void ResHelper::UpdateToolBarItemsInfo(CToolBar* tbar)
{
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_UpdateMenuItemsInfo == FALSE)
		return;

	CToolBarCtrl& cbar = tbar->GetToolBarCtrl();
	int count = cbar.GetButtonCount();

	CString label;
	CString label2;
	int i;
	wchar_t biText[512];
	for (i = 0; i < count; i++)
	{
		TBBUTTONINFO bi;
		bi.cbSize = sizeof(bi);
		bi.dwMask = TBIF_BYINDEX | TBIF_TEXT;
		bi.pszText = &biText[0];
		bi.cchText = 511;

		cbar.GetButtonInfo(i, &bi);
		tbar->GetButtonText(i, label);
		label2 = tbar->GetButtonText(i);
		CString newLabel;
		BOOL retFind = DetermineString(label, newLabel);
		//newLabel = L"Hello";
		BOOL ret = FALSE;
		if (newLabel.GetLength())
		{
			ret = tbar->SetButtonText(i, newLabel);
		}
		else
		{
			int deb = 1; // keep old text
		}

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
	key.Format(L"%s", label);
	hashsum_t hsum = g_resInfoMap.hash(&key);

	ResourceInfo *rinfo = g_resInfoMap.find(&key, hsum);
	if (rinfo == 0)
	{
		ResourceInfo* rinfo = new ResourceInfo(key, controlName);
		g_resInfoMap.insert(hsum, rinfo);
		g_resInfoArray.Add(rinfo);
	}

	return rinfo;
}

int ResHelper::PrintResInfoMap(ResInfoMapType& resInfoMap)
{
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_LoadMenuItemsInfo == FALSE)
		return -1;

	ResInfoMapType::IHashMapIter iter = resInfoMap.first();
	for (; !resInfoMap.last(iter); resInfoMap.next(iter))
	{
		TRACE(L"\"%8s\" %s\n", iter.element->m_controlName, iter.element->m_label);
	}
	return resInfoMap.count();
}

int ResHelper::PrintResInfoArray(ResInfoArrayType& resInfoArray, BOOL mustSort, BOOL printAsUnicode)
{
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_LoadMenuItemsInfo == FALSE)
		return -1;

	CString errorTxt;
	CString resInfoTxt;
	CStringA resInfoTxtA;
	CStringA resInfoTxtAnsi;
	DWORD nNumberOfBytesToWrite;
	DWORD nNumberOfBytesWritten;
	int retval = 0;
	ResInfoArrayType* resourceInfoArray = 0;
	ResInfoArrayType sorted_resourcefoArray;

	CString folderPath;
	GetProcessFolderPath(folderPath);
	//CString resourceFilePath = folderPath + L"\\";
	CString resourceFilePath;

	if (mustSort)
	{
		sorted_resourcefoArray.Append(resInfoArray);
		resourceInfoArray = &sorted_resourcefoArray;

		int sortCnt = ResHelper::SortResInfoArray(*resourceInfoArray);

		if (printAsUnicode)
			resourceFilePath = folderPath + L"sorted_unicode_" + resourceFile;
		else
			resourceFilePath = folderPath + L"sorted_" + resourceFile;
	}
	else
	{
		resourceInfoArray = &resInfoArray;
		if (printAsUnicode)
			resourceFilePath = folderPath + L"unsorted_unicode_" + resourceFile;
		else
			resourceFilePath = folderPath + L"unsorted_" + resourceFile;
	}

	hResourceFile = FileUtils::FileOpen(resourceFilePath, errorTxt, TRUE);
	if (hResourceFile == INVALID_HANDLE_VALUE)
	{
		return -1;
	}

	if (printAsUnicode == FALSE)
	{
		for (int i = 0; i < resourceInfoArray->GetCount(); i++)
		{
			ResourceInfo* rinfo = (*resourceInfoArray)[i];
			TRACE(L"%8s %s\n", rinfo->m_controlName, rinfo->m_label);
			resInfoTxt.Format(L"\"%s\"\r\n", rinfo->m_label);

			DWORD error;
			BOOL retW2A = TextUtilsEx::WStr2UTF8(&resInfoTxt, &resInfoTxtA, error);
			BOOL retW2Ansi = TextUtilsEx::WStr2Ansi(resInfoTxt, resInfoTxtAnsi, error);

			int retComp = resInfoTxtA.CompareNoCase(resInfoTxtAnsi);
			if (retComp)
			{
				_ASSERTE(retComp);
			}

			if (hResourceFile != INVALID_HANDLE_VALUE)
			{
				nNumberOfBytesToWrite = resInfoTxtA.GetLength();
				nNumberOfBytesWritten = 0;
				retval = FileUtils::Write2File(hResourceFile, (LPCSTR)resInfoTxtA, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
				int deb = 1;
			}
		}
	}
	else
	{
		const char* BOM_UTF16_LE = "\xFF\xFE";
		if (hResourceFile)
		{
			nNumberOfBytesToWrite = 2;
			nNumberOfBytesWritten = 0;

			retval = FileUtils::Write2File(hResourceFile, BOM_UTF16_LE, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
		}

		for (int i = 0; i < resourceInfoArray->GetCount(); i++)
		{
			ResourceInfo* rinfo = (*resourceInfoArray)[i];
			TRACE(L"%8s %s\n", rinfo->m_controlName, rinfo->m_label);
			resInfoTxt.Format(L"\"%s\"\r\n", rinfo->m_label);

			if (hResourceFile != INVALID_HANDLE_VALUE)
			{
				nNumberOfBytesToWrite = resInfoTxt.GetLength() * 2;
				nNumberOfBytesWritten = 0;
				retval = FileUtils::Write2File(hResourceFile, (LPCWSTR)resInfoTxt, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
				int deb = 1;
			}
		}
	}

	FileUtils::FileClose(hResourceFile);

	return (int)resourceInfoArray->GetCount();
}

int __cdecl ResourceInfoPred(void const* first, void const* second)
{
	ResourceInfo* f = *((ResourceInfo**)first);
	ResourceInfo* s = *((ResourceInfo**)second);

	//int ret = _wcsicmp(f->m_label.operator LPCWSTR(), s->m_label.operator LPCWSTR());
	int ret = f->m_label.CompareNoCase(s->m_label);
	//int ret = f->m_label.Compare(s->m_label);
	return ret;
}

int ResHelper::SortResInfoArray(ResInfoArrayType& resInfoArray)
{
	std::qsort(resInfoArray.GetData(), resInfoArray.GetCount(), sizeof(void*), ResourceInfoPred);
	return (int)resInfoArray.GetCount();
}


void ResHelper::MergeAllResInfo()
{
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_LoadMenuItemsInfo == FALSE)
		return;

	CStringArray fileList;
	fileList.Add(L"sorted_resource.txt");
	fileList.Add(L"german.ini");

	CString folderPath;
	GetProcessFolderPath(folderPath);
	CString resourceFilePath = folderPath + L"\\";

	CString sortedResFile = resourceFilePath + L"sorted_resource.txt";
	//CString germanResFile = resourceFilePath + L"german.ini";
	CString germanResFile = L"F:\\Documents\\GIT1.0.3.41.bak\\DLL\\german.ini";


	MergeResInfoFile(sortedResFile);
	MergeResInfoFile(germanResFile);
}

void ResHelper::MergeResInfoFile(CString &resFile)
{
	if (!FileUtils::PathFileExist(resFile))
		return;

	CString strLine;
	CString str;
	CStdioFile file(resFile, CFile::modeRead);
	while (file.ReadString(strLine))
	{
		str = strLine.Trim(L"\"");
		AddItemInfo(str);
		int deb = 1;
	}
}

void ResHelper::GetProcessFolderPath(CString &folderPath)
{
	CString processPath;
	CmboxviewApp::GetProcessPath(processPath);
	FileUtils::GetFolderPath(processPath, folderPath);
}

int ResHelper::CreateLanguageFile()
{
#if 0
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_LoadMenuItemsInfo == FALSE)
		return -1;
#endif

	CString folderPath;
	GetProcessFolderPath(folderPath);

	CString resourceFilePath = folderPath;

	//CString resFile1 = resourceFilePath + L"sorted_unicode_resource.txt";
	CString resFile1 = resourceFilePath + L"sorted_resource.txt";
	CString resFile2 = resourceFilePath + L"german.txt";

	CString languageFilePath = resourceFilePath + L"lang-german.txt";

	CString fileUTF8 = resourceFilePath + L"sorted_resource-UTF8.txt";

	BOOL isUTF8 = IsFileUTF8(fileUTF8);

	fileUTF8 = resourceFilePath + L"german-UTF8.txt";

	BOOL isUTF8_2 = IsFileUTF8(fileUTF8);

	ResHelper::TextEncoding bom = GetFileBOM(resFile1);
	if (bom == ResHelper::TextEncoding::NONE)
		int deb = 1;
	else if (bom == ResHelper::TextEncoding::UTF8)
		int deb = 1;
	else if (bom == ResHelper::TextEncoding::UTF16LE)
		int deb = 1;
	else if (bom == ResHelper::TextEncoding::UTF16BE)
		int deb = 1;
	else
		_ASSERTE(bom >= 0 && bom <= ResHelper::TextEncoding::UTF16BE);

	ResHelper::TextEncoding bom1 = bom;

	bom = GetFileBOM(resFile2);
	if (bom == ResHelper::TextEncoding::NONE)
		int deb = 1;
	else if (bom == ResHelper::TextEncoding::UTF8)
		int deb = 1;
	else if (bom == ResHelper::TextEncoding::UTF16LE)
		int deb = 1;
	else if (bom == ResHelper::TextEncoding::UTF16BE)
		int deb = 1;
	else
		_ASSERTE(bom >= 0 && bom <= ResHelper::TextEncoding::UTF16BE);

	ResHelper::TextEncoding bom2 = bom;

	// Should I first convert files to UNICODE in memory first ?? to simplify
	// and work with SimpleMemoryFile
	int ret = -1;
	if ((bom == ResHelper::TextEncoding::NONE) && (bom2 == ResHelper::TextEncoding::NONE))
	{
		ret = ResHelper::CreateLanguageFileFromANSIEncodedFiles(resFile1, resFile2, languageFilePath);
	}
	else if ((bom == ResHelper::TextEncoding::UTF8) && (bom2 == ResHelper::TextEncoding::UTF8))
	{
		ret = ResHelper::CreateLanguageFileFromUTF8EncodedFiles(resFile1, resFile2, languageFilePath);
	}
	else if ((bom == ResHelper::TextEncoding::UTF16LE) && (bom2 == ResHelper::TextEncoding::UTF16LE))
	{
		ret = ResHelper::CreateLanguageFileFromUTF16LEEncodedFiles(resFile1, resFile2, languageFilePath);
	}
	return ret;

}

int ResHelper::CreateLanguageFileFromANSIEncodedFiles(CString &resFile1, CString &resFile2, CString & languageFilePath)
{
	resArray1.RemoveAll();
	resArray2.RemoveAll();
	ResHelper::LoadResInfoFromFile(resFile1, resArray1);
	ResHelper::LoadResInfoFromFile(resFile2, resArray2);

	int cnt = (INT)resArray1.GetCount();
	_ASSERTE(resArray1.GetCount() == resArray2.GetCount());

	CString resInfoText1;
	CString resInfoText2;

	CStringA resInfoText1A;
	CStringA resInfoText2A;

	DWORD nNumberOfBytesToWrite = 0;
	DWORD nNumberOfBytesWritten = 0;

	int retval;
	CString errorText;

	hResourceFile = FileUtils::FileOpen(languageFilePath, errorText, TRUE);
	if (hResourceFile == INVALID_HANDLE_VALUE)
	{
		return -1;
	}

	int i;
	for (i = 0; i < cnt; i++)
	{
		BOOL unicodeFile = FALSE;
		if (unicodeFile == FALSE)
		{
			ResourceInfo* rinfo1 = resArray1[i];
			ResourceInfo* rinfo2 = resArray2[i];

			//TRACE(L"%8s %s\n", rinfo->m_controlName, rinfo->m_label);
			//resInfoText1.Format(L"[\"%s\"]\n", rinfo1->m_label);
			//resInfoText2.Format(L"\"%s\"\n", rinfo2->m_label);

			resInfoText1.Format(L"[%s]\n", rinfo1->m_label);
			resInfoText2.Format(L"%s\n\n", rinfo2->m_label);

			DWORD error;
			//BOOL retW2A = TextUtilsEx::WStr2UTF8(&resInfoText1, &resInfoText1A, error);
			//retW2A = TextUtilsEx::WStr2UTF8(&resInfoText2, &resInfoText2A, error);
			BOOL retW2A = TextUtilsEx::WStr2Ansi(resInfoText1, resInfoText1A, error);
			retW2A = TextUtilsEx::WStr2Ansi(resInfoText2, resInfoText2A, error);

			if (hResourceFile != INVALID_HANDLE_VALUE)
			{
				nNumberOfBytesToWrite = resInfoText1A.GetLength();
				nNumberOfBytesWritten = 0;
				retval = FileUtils::Write2File(hResourceFile, (LPCSTR)resInfoText1A, nNumberOfBytesToWrite, &nNumberOfBytesWritten);

				nNumberOfBytesToWrite = resInfoText2A.GetLength();
				nNumberOfBytesWritten = 0;
				retval = FileUtils::Write2File(hResourceFile, (LPCSTR)resInfoText2A, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
				int deb = 1;
			}
		}
	}
	FileUtils::FileClose(hResourceFile);
	return 1;
}

int ResHelper::CreateLanguageFileFromUTF8EncodedFiles(CString& resFile1, CString& resFile2, CString& languageFilePath)
{
	return -1;
}


int ResHelper::CreateLanguageFileFromUTF16LEEncodedFiles(CString& resFile1, CString& resFile2, CString& languageFilePath)
{
	resArray1.RemoveAll();
	resArray2.RemoveAll();

	ResHelper::LoadResInfoFromFile(resFile1, resArray1);
	ResHelper::LoadResInfoFromFile(resFile2, resArray2);

	int cnt = (INT)resArray1.GetCount();
	_ASSERTE(resArray1.GetCount() == resArray2.GetCount());

	CString resInfoText1;
	CString resInfoText2;

	CStringA resInfoText1A;
	CStringA resInfoText2A;

	DWORD nNumberOfBytesToWrite = 0;
	DWORD nNumberOfBytesWritten = 0;

	int retval;
	CString errorText;

	hResourceFile = FileUtils::FileOpen(languageFilePath, errorText, TRUE);
	if (hResourceFile == INVALID_HANDLE_VALUE)
	{
		return -1;
	}

	BYTE BOMF16LE[2] = { 0xFF, 0xFE };

	nNumberOfBytesToWrite = 2;
	nNumberOfBytesWritten = 0;
	retval = FileUtils::Write2File(hResourceFile, BOMF16LE, nNumberOfBytesToWrite, &nNumberOfBytesWritten);

	int i;
	for (i = 0; i < cnt; i++)
	{
		BOOL unicodeFile = FALSE;
		if (unicodeFile == FALSE)
		{
			ResourceInfo* rinfo1 = resArray1[i];
			ResourceInfo* rinfo2 = resArray2[i];

			resInfoText1.Format(L"[%s]\n", rinfo1->m_label);
			resInfoText2.Format(L"%s\n\n", rinfo2->m_label);

			if (hResourceFile != INVALID_HANDLE_VALUE)
			{
				nNumberOfBytesToWrite = resInfoText1.GetLength()*2;
				nNumberOfBytesWritten = 0;
				retval = FileUtils::Write2File(hResourceFile, (LPCWSTR)resInfoText1, nNumberOfBytesToWrite, &nNumberOfBytesWritten);

				nNumberOfBytesToWrite = resInfoText2.GetLength()*2;
				nNumberOfBytesWritten = 0;
				retval = FileUtils::Write2File(hResourceFile, (LPCWSTR)resInfoText2, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
				int deb = 1;
			}
		}
	}
	FileUtils::FileClose(hResourceFile);
	return 1;
}


int ResHelper::ResortLanguageFile()
{
	CString label;
	CString translatedLabel;
	ResInfoMapType::IHashMapIter iter = g_LanguageMap.first();

	resArray1.RemoveAll();
	for (; !g_LanguageMap.last(iter); g_LanguageMap.next(iter))
	{
		ResourceInfo* rinfo = iter.element;

		//TRACE(L"\"%8s\" %s\n", rinfo->m_controlName, rinfo->m_label);

		label.Format(L"\"%s\"", rinfo->m_label);
		translatedLabel.Format(L"\"%s\"", rinfo->m_controlName);

		ResourceInfo* rinfo2 = new ResourceInfo(label, translatedLabel);

		resArray1.Add(rinfo2);
		//resArray1.Add(rinfo);
	}

	int sortCnt = ResHelper::SortResInfoArray(resArray1);

	int mapCnt = g_LanguageMap.count();
	int arrayCnt = (INT)resArray1.GetCount();

	CString resInfoText1;
	CString resInfoText2;

	CStringA resInfoText1A;
	CStringA resInfoText2A;

	DWORD nNumberOfBytesToWrite = 0;
	DWORD nNumberOfBytesWritten = 0;

	CString folderPath;
	GetProcessFolderPath(folderPath);

	CString resourceFilePath = folderPath;

	CString languageFilePath = resourceFilePath + L"resorted-lang-german.txt";

	int retval;
	CString errorText;

	SimpleMemoryFile file;
	BOOL retOpen = file.Open(languageFilePath);
	CString str;

	hResourceFile = FileUtils::FileOpen(languageFilePath, errorText, TRUE);
	if (hResourceFile == INVALID_HANDLE_VALUE)
	{
		return -1;
	}

	int i;
	for (i = 0; i < arrayCnt; i++)
	{
		BOOL unicodeFile = FALSE;
		if (unicodeFile == FALSE)
		{
			ResourceInfo* rinfo1 = resArray1[i];

			if (rinfo1->m_label.Find(L"Valid template format") >= 0)
				int deb = 1;

			resInfoText1.Format(L"[%s]\n", rinfo1->m_label);
			resInfoText2.Format(L"%s\n\n", rinfo1->m_controlName);

			DWORD error;
			BOOL retW2A = TextUtilsEx::WStr2Ansi(resInfoText1, resInfoText1A, error);
			retW2A = TextUtilsEx::WStr2Ansi(resInfoText2, resInfoText2A, error);

			if (hResourceFile != INVALID_HANDLE_VALUE)
			{
				nNumberOfBytesToWrite = resInfoText1A.GetLength();
				nNumberOfBytesWritten = 0;
				retval = FileUtils::Write2File(hResourceFile, (LPCSTR)resInfoText1A, nNumberOfBytesToWrite, &nNumberOfBytesWritten);

				nNumberOfBytesToWrite = resInfoText2A.GetLength();
				nNumberOfBytesWritten = 0;
				retval = FileUtils::Write2File(hResourceFile, (LPCSTR)resInfoText2A, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
				int deb = 1;
			}
		}
	}
	FileUtils::FileClose(hResourceFile);
	return 1;
}


void ResHelper::LoadResInfoFromFile(CString& resFile, ResInfoArrayType &resArray)
{
	if (!FileUtils::PathFileExist(resFile))
		return;

	CString controlName;
	CString strLine;
	CString str;

	CStdioFile file;

	UINT nOpenFlags = CFile::modeRead | CFile::typeText | CFile::shareDenyNone;
	ResHelper::TextEncoding  bom = GetFileBOM(resFile);
	if (bom == ResHelper::TextEncoding::NONE)
		int deb = 1;
	else if (bom == ResHelper::TextEncoding::UTF8)
		int deb = 1;
	else if (bom == ResHelper::TextEncoding::UTF16LE)
		int deb = 1;
	else if (bom == ResHelper::TextEncoding::UTF16BE)
		int deb = 1;
	else
		_ASSERTE(bom >= 0 && bom <= ResHelper::TextEncoding::UTF16BE);

	ResHelper::TextEncoding bom1 = bom;

	if (bom == ResHelper::TextEncoding::UTF16LE)
		nOpenFlags |= CFile::typeUnicode;


	CFileException exList;
	if (!file.Open(resFile, nOpenFlags, &exList))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exList);

		CString txt = L"Could not open list file \"" + resFile;
		txt += L"\" file.\n";
		txt += exErrorStr;

		TRACE(L"ResHelper: %s\n", txt);
		//errorText = txt;

		CFileStatus rStatus;
		BOOL ret = file.GetStatus(rStatus);

		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return;
	}
	while (file.ReadString(strLine))
	{
		str = strLine.TrimRight(L"\r\n");

		ResourceInfo* rinfo = new ResourceInfo(str, controlName);
		resArray.Add(rinfo);

		int deb = 1;
	}
}

void ResHelper::LoadLanguageMap(CString& languageTranslationFilePath)
{
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_UpdateMenuItemsInfo == FALSE)  // Do I need separate control var ?
		return;

	CString controlName;
	CString strLine;
	CString str;
	CString strW;

	CString languageFile = languageTranslationFilePath;

	ResHelper::TextEncoding  bom = GetFileBOM(languageFile);
	if (bom == ResHelper::TextEncoding::NONE)
		int deb = 1;
	else if (bom == ResHelper::TextEncoding::UTF8)
		int deb = 1;
	else if (bom == ResHelper::TextEncoding::UTF16LE)
		int deb = 1;
	else if (bom == ResHelper::TextEncoding::UTF16BE)
		int deb = 1;
	else
		_ASSERTE(bom >= 0 && bom <= ResHelper::TextEncoding::UTF16BE);

	ResHelper::TextEncoding bom1 = bom;

	// TODO: Convert to ResHelper::TextEncoding::UTF16LE first if necessary

	if (bom == ResHelper::TextEncoding::UTF16LE)
	{
		LoadLanguageMapFromFileF16LE(languageFile);
		int deb = 1;
	}
	else  // TODO: Ansi ???
	{
		LoadLanguageMapFromFileAnsi(languageFile);
		int deb = 1;
	}
	return;
}

void ResHelper::LoadLanguageMapFromFileAnsi(CString& languageTranslationFilePath)
{
	CString controlName;
	CString strLine;
	CString str;
	CString strW;

	CString languageFile = languageTranslationFilePath;

	SimpleMemoryFile memFile;
	BOOL retOpen = memFile.Open(languageFile);
	CStringA strA;

	CStdioFile file;
	CFileException exList;
	//if (!file.Open(languageFile, CFile::modeRead | CFile::shareDenyNone | CFile::typeUnicode, &exList))
	if (!file.Open(languageFile, CFile::modeRead | CFile::shareDenyNone, &exList))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exList);

		CString txt = L"Could not open list file \"" + languageFile;
		txt += L"\" file.\n";
		txt += exErrorStr;

		TRACE(L"ResHelper: %s\n", txt);
		//errorText = txt;

		CFileStatus rStatus;
		BOOL ret = file.GetStatus(rStatus);

		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return;
	}
	while (file.ReadString(strLine))
	{
		int retEof  = memFile.ReadString(strA);

		DWORD dwFlags = 0;
		DWORD error;
		BOOL a2W = TextUtilsEx::Ansi2WStr(strA, strW, error, dwFlags);

		if (strW.Compare(strLine))
			int deb = 1;

		CString line = strLine;

		if (line.Find(L"Valid template format") >= 0)
			int deb = 1;
		if (line.Compare(L"[\"from\"]") == 0)
			int deb = 1;

		if (line.Compare(L"[\"All Mails\"]") == 0)
			int deb = 1;

		if (strLine[0] == L'[')
		{
			int slen = strLine.GetLength();
			CString key = strLine.Mid(2, slen - 4);
			//CString key = strLine.Trim(L"[]\"\r\n");

			hashsum_t hsum = g_LanguageMap.hash(&key);

			ResourceInfo* rinfo = g_LanguageMap.find(&key, hsum);
			if (rinfo == 0)
			{
				if (!file.ReadString(strLine))
					break;

				int retEof = memFile.ReadString(strA);
				DWORD dwFlags = 0;
				DWORD error;
				BOOL a2W = TextUtilsEx::Ansi2WStr(strA, strW, error, dwFlags);

				if (strW.Compare(strLine))
					int deb = 1;

				int slen = strLine.GetLength();
				controlName = strLine.Mid(1, slen - 2);
				//controlName = strLine.Trim(L"[]\"\r\n");
				ResourceInfo* rinfo = new ResourceInfo(key, controlName);
				g_LanguageMap.insert(hsum, rinfo);
			}
			else
				int deb = 1;
		}
	}
	int deb = 1;
}


void ResHelper::LoadLanguageMapFromFileF16LE(CString& languageTranslationFilePath)
{
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_UpdateMenuItemsInfo == FALSE)  // Do I need separate control var ?
		return;

	CString controlName;
	CString strLine;
	CString str;
	CString strW;

	CString languageFile = languageTranslationFilePath;

	UINT nOpenFlags = CFile::modeRead | CFile::typeText | CFile::shareDenyNone;
	ResHelper::TextEncoding  bom = GetFileBOM(languageFile);
	if (bom == ResHelper::TextEncoding::NONE)
		int deb = 1;
	else if (bom == ResHelper::TextEncoding::UTF8)
		int deb = 1;
	else if (bom == ResHelper::TextEncoding::UTF16LE)
		int deb = 1;
	else if (bom == ResHelper::TextEncoding::UTF16BE)
		int deb = 1;
	else
		_ASSERTE(bom >= 0 && bom <= ResHelper::TextEncoding::UTF16BE);

	ResHelper::TextEncoding bom1 = bom;

	if (bom1 == ResHelper::TextEncoding::UTF16LE)
		nOpenFlags |= CFile::typeUnicode;

	CStdioFile file;
	CFileException exList;
	//if (!file.Open(languageFile, CFile::modeRead | CFile::shareDenyNone | CFile::typeUnicode, &exList))
	if (!file.Open(languageFile, nOpenFlags, &exList))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exList);

		CString txt = L"Could not open list file \"" + languageFile;
		txt += L"\" file.\n";
		txt += exErrorStr;

		TRACE(L"ResHelper: %s\n", txt);
		//errorText = txt;

		CFileStatus rStatus;
		BOOL ret = file.GetStatus(rStatus);

		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return;
	}
	while (file.ReadString(strLine))
	{
		if (strLine[0] == L'[')
		{
			int slen = strLine.GetLength();
			CString key = strLine.Mid(2, slen - 4);

			hashsum_t hsum = g_LanguageMap.hash(&key);

			ResourceInfo* rinfo = g_LanguageMap.find(&key, hsum);
			if (rinfo == 0)
			{
				if (!file.ReadString(strLine))
					break;

				int slen = strLine.GetLength();
				controlName = strLine.Mid(1, slen - 2);

				ResourceInfo* rinfo = new ResourceInfo(key, controlName);
				g_LanguageMap.insert(hsum, rinfo);
			}
			else
				int deb = 1;
		}
	}
	int deb = 1;
}

int ResHelper::PrintResInfo()
{
	// TODO: Add your command handler code here
#ifdef _DEBUG
	BOOL LoadMenuItemsInfo = g_LoadMenuItemsInfo;
	BOOL UpdateMenuItemsInfo = g_UpdateMenuItemsInfo;
	g_LoadMenuItemsInfo = TRUE;
	g_UpdateMenuItemsInfo = FALSE;
#endif
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_LoadMenuItemsInfo == FALSE)
		return -1;

#if 0
	TRACE(L"BEGIN PrintResInfoMap\n");
	int cnt3 = ResHelper::PrintResInfoMap(ResHelper::resInfoMap);
	TRACE(L"END PrintResInfoMap\n");
#endif
	TRACE(L"BEGIN PrintResInfoArray\n");
	//int sortCnt = ResHelper::SortResInfoArray(ResHelper::resInfoArray);
	ResHelper::MergeAllResInfo();

	BOOL sort = FALSE;
	BOOL printAsUnicode = TRUE;
	int cnt4 = ResHelper::PrintResInfoArray(ResHelper::g_resInfoArray, sort, printAsUnicode);
	printAsUnicode = FALSE;
	int cnt5 = ResHelper::PrintResInfoArray(ResHelper::g_resInfoArray, sort, printAsUnicode);


	sort = TRUE;
	printAsUnicode = TRUE;
	int cnt6 = ResHelper::PrintResInfoArray(ResHelper::g_resInfoArray, sort, printAsUnicode);
	printAsUnicode = FALSE;
	int cnt7 = ResHelper::PrintResInfoArray(ResHelper::g_resInfoArray, sort, printAsUnicode);

	TRACE(L"END PrintResInfo\n");

	int deb = 1;
#ifdef _DEBUG
	g_LoadMenuItemsInfo = LoadMenuItemsInfo;
	g_UpdateMenuItemsInfo = UpdateMenuItemsInfo;
#endif

	return 1;
}

BOOL ResHelper::OnTtnNeedText(CWnd* parentWnd, NMHDR* pNMHDR, CString& toolTipText)
{
	NMTTDISPINFO* pTTT = (NMTTDISPINFO*)pNMHDR;
	HWND hwndID = (HWND)pNMHDR->idFrom;
	BOOL bRet = FALSE;

	if (pTTT->uFlags & TTF_IDISHWND)
	{
		// idFrom is actually the HWND of the tool
		int nID = ::GetDlgCtrlID(hwndID);
		//if (nID && (nID != IDC_EDIT_INTRO))
		if (nID)
		{
			CWnd* p = parentWnd->GetDlgItem(nID);
			if (p)
			{
				p->GetWindowText(toolTipText);

				CPaintDC dc(p);
				HDC hDC = dc.GetSafeHdc();

				SIZE sizeItem;
				BOOL retA = GetTextExtentPoint32(hDC, toolTipText, toolTipText.GetLength(), &sizeItem);

				CRect rec;
				p->GetClientRect(&rec);
				//rec.InflateRect(-5, 0);

				if (sizeItem.cx > rec.Width())
				{
					//int w = m_toolTip.SetMaxTipWidth(200);

					pTTT->lpszText = (LPWSTR)((LPCWSTR)toolTipText);

					RECT rec;
					p->GetWindowRect(&rec);

					HWND hwndToolTip = pNMHDR->hwndFrom;

					RECT rc = rec;
					::SendMessage(hwndToolTip, TTM_ADJUSTRECT, TRUE, (LPARAM)&rc);
					bRet = TRUE;
				}
			}
		}
	}
	else
	{
		// This is big hack. Need to investigate and find better solution
		// // Add tipText to language translation files
		//
		QWORD nID = pTTT->hdr.idFrom;

		CString tipText;
		if (ID_TREE_HIDE == nID)
		{
			//tipText = L"Expand/Colapse tree containing folders\nExpand/Collapse mail folders";
			tipText = L"Hide/Restore mail tree";
		}
		else if (ID_FILE_REFRESH == nID)
		{
			//tipText = L"Refresh folder containing mails\nRefresh mail folders";
			tipText = L"Refresh folder containing mails";
		}
		else if (ID_FILE_SELECT_FOLDER == nID)
		{
			//tipText = L"Select folder containing mails\nSelect folder";
			tipText = L"Select folder containing mails";
		}
		else if (ID_APP_ABOUT == nID)
		{
			//tipText = L"Display program information, version number and copyright\nAbout";
			tipText = L"Display program information, version number and copyright";
		}
		else if (ID_TREE_EXPAND == nID)
		{
			//tipText = L"Expand/Colapse tree containing folders\nExpand/Collapse mail folders";
			tipText = L"Expand/Collapse mail folders";
		}

		if (!tipText.IsEmpty())
		{
			CString newTipText;
			BOOL retFind = DetermineString(tipText, newTipText);

			if (newTipText.GetLength())
			{
				CWnd* p = parentWnd;

				pTTT->lpszText = (LPWSTR)((LPCWSTR)newTipText);

				RECT rec;
				p->GetWindowRect(&rec);

				HWND hwndToolTip = pNMHDR->hwndFrom;

				RECT rc = rec;
				::SendMessage(hwndToolTip, TTM_ADJUSTRECT, TRUE, (LPARAM)&rc);

				CMainFrame* pFrame = (CMainFrame*)AfxGetMainWnd();
				if (pFrame)
				{
					if (pFrame) {
						int paneId = 0;
						pFrame->SetStatusBarPaneText(paneId, newTipText, FALSE);
					}
				}
				bRet = TRUE;
			}
		}
	}
	return bRet;
}

BOOL ResHelper::ActivateToolTips(CWnd* parentWnd, CToolTipCtrl &toolTipCtrl)
{
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_UpdateMenuItemsInfo == FALSE) // do I need separate control var ?
		return FALSE;

	parentWnd->EnableToolTips(TRUE);
	BOOL cret = toolTipCtrl.Create(parentWnd);
	if (cret)
	{
		BOOL ret = toolTipCtrl.AddTool(parentWnd, LPSTR_TEXTCALLBACK);
		//int w = toolTipCtrl.SetMaxTipWidth(600);
		toolTipCtrl.Activate(TRUE);
		return TRUE;
	}
	return FALSE;
}

//////////////////////////////////////////////////////////////////////


int ResHelper::ReleaseResInfoMap(ResInfoMapType& resInfoMap)
{
	ResInfoMapType::IHashMapIter iter = resInfoMap.first();
	for (; !resInfoMap.last(iter); )
	{
		ResourceInfo* l = iter.element;
		resInfoMap.remove(iter);
		delete l;
	}
	resInfoMap.clear();

	return resInfoMap.count();
}

void ResHelper::ReleaseResInfoArray(ResInfoArrayType &resInfoArray)
{
	for (int i = 0; i < resInfoArray.GetCount(); i++)
	{
		ResourceInfo* rinfo = resInfoArray[i];
		delete rinfo;
		resInfoArray[i] = 0;
	}
}

void ResHelper::ReleaseResources()
{
	g_resInfoMap.clear();
	ReleaseResInfoMap(g_LanguageMap);

	ReleaseResInfoArray(g_resInfoArray);
	ReleaseResInfoArray(resArray1);
	ReleaseResInfoArray(resArray2);
}

ResHelper::TextEncoding ResHelper::GetFileBOM(LPCWSTR lpszFileName)
{
	CStdioFile File;
	if (!File.Open(lpszFileName, CFile::modeRead | CFile::typeBinary))
		return TextEncoding::NONE;  /// TextEncoding::ERROR ??

	BYTE string[4] = { 0, 0 , 0 };
	int slen = File.Read(&string, 4);

	ResHelper::TextEncoding BOM = ResHelper::GetBOM((unsigned char*)string, slen);
	return BOM;
}

ResHelper::TextEncoding ResHelper::GetBOM(unsigned char* string, int slen)
{
	TextEncoding BOM = TextEncoding::NONE;
	if (slen >= 4)
	{
		unsigned char c0 = string[0];
		unsigned char c1 = string[1];
		unsigned char c2 = string[2];
		unsigned char c3 = string[3];
		if ((c0 == 0xFF) && (c1 == 0xFE) && (c2 == 0) && (c3 == 0))
			BOM = TextEncoding::UTF32LE;
		else if ((c0 == 0x0) && (c1 == 0x0) && (c2 == 0XFE) && (c3 == 0xFF))
			BOM = TextEncoding::UTF32BE;
	}
	if ((BOM == TextEncoding::NONE) && (slen >= 3))
	{
		unsigned char c0 = string[0];
		unsigned char c1 = string[1];
		unsigned char c2 = string[2];
		if ((c0 == 0xEF) && (c1 == 0xBB) && (c2 == 0xBF))
			BOM = TextEncoding::UTF8;
	}
	if ((BOM == TextEncoding::NONE) && (slen >= 2))
	{
		unsigned char c0 = string[0];
		unsigned char c1 = string[1];
		if ((c0 == 0xFF) && (c1 == 0xFE))
			BOM = TextEncoding::UTF16LE;
		else if ((c0 == 0XFE) && (c1 == 0xFF))
			BOM = TextEncoding::UTF16BE;
	}

	return BOM;
}


int ResHelper::GetCodePageFromFile(LPCWSTR filePath)
{
	CStdioFile file;
	CFileException exList;
	UINT nOpenFlags = CFile::modeRead | CFile::typeText | CFile::shareDenyNone;
	if (!file.Open(filePath, nOpenFlags, &exList))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exList);

		CString txt = L"Could not open list file \"" + *filePath;
		txt += L"\" file.\n";
		txt += exErrorStr;

		TRACE(L"ResHelper: %s\n", txt);
		//errorText = txt;

		CFileStatus rStatus;
		BOOL ret = file.GetStatus(rStatus);

		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return FALSE;
	}
		return -1;


	file.Close();
	return ResHelper::TextEncoding::NONE;
}




// SimpleMemoryFile for relatively small files since entire file data is loaded into memory

SimpleMemoryFile::SimpleMemoryFile()
{

	m_position = 0;
}

SimpleMemoryFile::~SimpleMemoryFile()
{

}

BOOL SimpleMemoryFile::Open(CString& filePath)
{
	BOOL ret = FileUtils::ReadEntireFile(filePath, m_buffer);
	return TRUE;
}

int SimpleMemoryFile::ReadString(CStringA &str)
{
	char* p = m_buffer.Data(m_position);
	char* plast = p + m_buffer.Count() - m_position;

	int len = -1;
	char* pStrBegin = p;
	char* pStrEnd = 0;
	while (p < plast)
	{
		char c = *p;
		if (c == '\r')
		{
			pStrEnd = p;
			char* pnext = p + 1;
			if ((pnext < plast) && (*pnext == '\n'))
			{
				pStrEnd = p++;
			}
			break;
		}
		else if (c == '\n')
		{
			pStrEnd = p++;
			break;
		}
		else
			p++;
	}
	if (pStrEnd == 0)
		pStrEnd = p;

	m_position += (int)(p - pStrBegin);
	char* pnext = m_buffer.Data(m_position);

	int slen = (int)(pStrEnd - pStrBegin);

	str.Empty();
	str.Append(pStrBegin, slen);

	return slen;
}

int SimpleMemoryFile::ReadString(CString& str)
{
	return 0;
}

#include <stdint.h>

// copy from https://github.com/vstinner/unicode_book
// Updated by MBox Viewer team
// TODO investigate reliability of algorithm

BOOL ResHelper::IsTextUTF8Done(BOOL retval)
{
	return retval;
}

BOOL ResHelper::IsTextUTF8(const char *data, size_t size, BOOL &isASCII)
{
    const unsigned char *str = (unsigned char*)data;
    unsigned int code_length, i;
    uint32_t ch;
    const unsigned char *end = str + size;

	isASCII = TRUE;
	while (str != end)
	{
		if (*str == 0)
		{
			int deb = 1;
		}
		if (*str <= 0x7F) {
			/* 1 byte character (ASCII): U+0000..U+007F */
			str += 1;
			continue;
		}
		else
		{
			isASCII = FALSE;
			break;
		}

	}
	if (str == end)
	{
		_ASSERT(isASCII == TRUE);
		return IsTextUTF8Done(1);
		//return TRUE;
	}

	while (str != end)
	{
		if (*str == 0)
		{
			int deb = 1;
		}

		if (*str <= 0x7F) {
			/* 1 byte character (ASCII): U+0000..U+007F */
			str += 1;
			continue;
		}

        if (0xC2 <= *str && *str <= 0xDF)
            code_length = 2;
        else if (0xE0 <= *str && *str <= 0xEF)
            code_length = 3;
        else if (0xF0 <= *str && *str <= 0xF4)
            code_length = 4;
        else {
            /* invalid first byte of a multibyte character */
			return IsTextUTF8Done(0);
           //return 0;
        }

        if (str + (code_length -1) >= end) {
            /* truncated string or invalid byte sequence */
			return IsTextUTF8Done(0);
            //return 0;
        }

        /* Check continuation bytes: bit 7 should be set, bit 6 should unset */
        for (i=1; i < code_length; i++) {
            if ((str[i] & 0xc0) != 0x80)
				return IsTextUTF8Done(0);
                //return 0;
        }

        if (code_length == 2) {
            ch = ((str[0] & 0x1f) << 6) + (str[1] & 0x3f);
            /* 2 bytes sequence: U+0080..U+07FF */
            if ((ch < 0x0080) || (0x07FF < ch))
				return IsTextUTF8Done(0);
                //return 0;

        } else if (code_length == 3) {
            ch = ((str[0] & 0x0f) << 12) + ((str[1] & 0x3f) << 6) +
                  (str[2] & 0x3f);
            /* 3 bytes sequence: U+0800..U+FFFF */
            if ((ch < 0x0800) || (0xFFFF < ch))
				return IsTextUTF8Done(0);
                //return 0;

            /* 3 bytes sequence: U+0800..U+FFFF... excluding U+D800..U+DFFF:
             * surrogates are invalid in UTF-8 */
            if ((0xD800 <= ch) && (ch <= 0xDFFF))
				return IsTextUTF8Done(0);
                //return 0;

        } else if (code_length == 4) {
            ch = ((str[0] & 0x7) << 18) + ((str[1] & 0x3f) << 12) +
                 ((str[2] & 0x3f) << 6) + (str[3] & 0x3f);
            /* 4 bytes sequence: U+10000..U+10FFFF */
            if ((ch < 0x10000) || (0x10FFFF < ch))
				return IsTextUTF8Done(0);
                //return 0;
        }
        str += code_length;
    }
	return IsTextUTF8Done(1);
	//return 1;
}


int ResHelper::IsFileUTF8(LPCWSTR filePath)
{
	CStdioFile file;
	CFileException exList;
	//if (!file.Open(resFile, CFile::modeRead | CFile::shareDenyNone | CFile::typeUnicode, &exList))
	if (!file.Open(filePath, CFile::modeRead | CFile::shareDenyNone | CFile::typeBinary, &exList))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exList);

		CString txt = L"Could not open list file \"" + *filePath;
		txt += L"\" file.\n";
		txt += exErrorStr;

		TRACE(L"ResHelper: %s\n", txt);
		//errorText = txt;

		CFileStatus rStatus;
		BOOL ret = file.GetStatus(rStatus);

		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return FALSE;
	}
	const int bufflen = 100000 - 1;
	char buff[bufflen];
	char* pbuff = &buff[0];
	int retlen = 0;
	BOOL isASCII = TRUE;
	BOOL isAsciiText = TRUE;
	// Read all into memory
	ULONGLONG dwLength = file.GetLength();
	// TODO: enhnace to process  less than full file size
	if (dwLength > bufflen)
	{
		pbuff = new char[(UINT)dwLength];
	}
	int retcode = -1;
	if (pbuff)
	{
		if ((retlen = file.Read(pbuff, (UINT)dwLength)) > 0)
		{
			if (IsTextUTF8(pbuff, retlen, isAsciiText))
			{
				isASCII = isAsciiText;
				retcode = 1;
			}
			else
				retcode = 0;
		}
	}
	if (pbuff != &buff[0])
		delete[] pbuff;
	return retcode;  // could not memory
}

// codeproject Simple Character Encoding Detection
// // https://www.codeproject.com/tips/672470/simple-character-encoding-detection
// requires to add null characters, that is unusual
// #define NT "\0\0\0" 
// char* exampleString = "This is UTF-8" NT;
// Output:
// 0 = UTF - 8
// 1 = UTF - 16BE
// 2 = UTF - 16LE
// 3 = UTF - 32BE
// 4 = UTF - 32LE


ResHelper::TextEncoding ResHelper::GetTextEncoding(char* string, int slen, TextEncoding& BOM)
{
	BOM = ResHelper::GetBOM((unsigned char*)string, slen);

	// Need to adjuast slen and string ??

	int j;
	for (j = 0; j < slen; j++)
	{
		if (string[j] == 0)
			break;
	}
	if (j == slen)
	{
		// Did not find null
		BOOL isASCII;
		BOOL isUTF8 = IsTextUTF8(string, slen, isASCII);
		if (isASCII)
			return TextEncoding::ASCII;
		else if (isUTF8)
			return TextEncoding::UTF8;
		else
			return TextEncoding::ANSI;
	}
	// strings has nulls. check for two consecutive nulls
	// Optimistic code below
	int oddNullsCnt = 0;
	int evenNullsCnt = 0;
	unsigned char c1;
	unsigned char c2;
	for (j = 0; j < (slen - 1); j++)
	{
		c1 = string[j];
		c2 = string[j + 1];

		if ((c1 <= 0x7f) && (c1 > 0))
		{
			if ((INT_PTR)(&string[j + 1]) % 2)
				evenNullsCnt++;
			else
				oddNullsCnt++;

		}
		if ((c1 == 0) && (c2 == 0))
		{
			break;
		}
	}

	if (j == (slen - 1))
	{
		if (oddNullsCnt > 0)
			return TextEncoding::UTF16BE;
		else
			return TextEncoding::UTF16LE;
	}
	else
	{
		if (oddNullsCnt > 0)
			return TextEncoding::UTF32BE;
		else
			return TextEncoding::UTF32LE;
	}

#if 0
	// posted code doesn't seem to work
	unsigned c, i = 0, flags = 0;
	while (string[i] | string[i + 1] | string[i + 2] | string[i + 3])
		flags = (c = string[i++]) ? flags | ((!(flags % 4) &&
			c > 0x7F) << 3) : flags | 1 | (!(i & 1) << 1)
		| ((string[i] == 0) << 2);
	return (flags & 1) + ((flags & 2) != 0) +
		((flags & 4) != 0) + ((flags & 8) != 0);
#endif
}

int ResHelper::FindCodePageFromFile(LPCWSTR lpszFileName)
{
	return 1;
}

