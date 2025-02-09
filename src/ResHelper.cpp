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
#include "MboxMail.h"

//

CString ResHelper::resourceRootPath = LR"(F:\Documents\GIT1.0.3.42\ResourceRoot\WORK\)";

// When g_LoadMenuItemsInfo is set to TRUE, MBox Viewer will traverse all controls selected by a user and collect control's window text in the hash table
// When "File->Dvelopmen->Create Resource File" option is selected, MBox Viewer will merge content of rcFileName, last_sourceFileName, extraFileName 
// with the content of the hash table and generate the new sourceFileName
//
CString ResHelper::rcFileName = L"rc.txt";  // text of controls from mboxview.rc file
CString ResHelper::last_sourceFileName = L"last_source.txt";  // text of all controls including MessgeBox controls from the last run
CString ResHelper::sourceFileName = L"source.txt";  // text of all controls including MessgeBox controls
CString ResHelper::translated_sourceFileName = L"translated_source.txt";  // translated text of all controls in source.txt
CString ResHelper::extraFileName = L"extra.txt";  // text of controls that had be defined manually
CString ResHelper::msgsFileName = L"msgs.txt";


BOOL ResHelper::g_LoadMenuItemsInfo = FALSE; // Traverse dialog & menu controles and create resource L"resource.txt" file. i.e. collect text from all items/sub-controls
BOOL ResHelper::g_UpdateMenuItemsInfo = FALSE; // Update text of al controls using selected translation file

ResInfoMapType ResHelper::g_LanguageMap(2000);

ResInfoMapType ResHelper::g_mergedResInfoMap(2000);
ResInfoArrayType ResHelper::g_mergedResInfoArray;

ResInfoMapType ResHelper::g_resInfoMap(2000);
ResInfoArrayType ResHelper::g_resInfoArray;

ResInfoArrayType ResHelper::resArray1;
ResInfoArrayType ResHelper::resArray2;

int ResHelper::maxWindowTextLength = WINDOW_TEXT_LENGTH;
wchar_t ResHelper::windowText[WINDOW_TEXT_LENGTH + 1];
int ResHelper::maxClassNameLength = CLASS_NAME_LENGTH;
wchar_t ResHelper::className[CLASS_NAME_LENGTH + 1];
CString ResHelper::wndText;

CString ResHelper::resourceFile = L"rc.txt";
HANDLE ResHelper::hResourceFile = INVALID_HANDLE_VALUE;


void ResHelper::FindConversionSpecifiers(CString& str, CString& conversionSpecifiersList)
{
	const wchar_t* p = (LPCWSTR)str;
	wchar_t c;
	const wchar_t* pbegin;
	const wchar_t* pend;
	CString spec;
	conversionSpecifiersList.Empty();
	while ((c = *p) != 0)
	{
		if (c != L'%')
		{
			p++;
			continue;
		}
		pbegin = p++;
		// find s or d charcater
		while ((c = *p) != 0)
		{
			if ((c == L'0') || (c == L'l'))
			{
				p++;
				continue;
			}
			else if ((c != L's') && (c != L'd'))
			{
				p++;
				break;
			}
			pend = p;
			if (!conversionSpecifiersList.IsEmpty())
				conversionSpecifiersList.AppendChar(L' ');
			conversionSpecifiersList.Append(pbegin, (int)(pend - pbegin + 1));
			break;
		}
	}
}

void ResHelper::FindConversionSpecifiers(CString& str, CStringArray& conversionSpecifiersArray)
{
	const wchar_t* p = (LPCWSTR)str;
	wchar_t c;
	const wchar_t *pbegin;
	const wchar_t* pend;
	CString spec;
	while ((c = *p) != 0)
	{
		if (c != L'%')
		{
			p++;
			continue;
		}
		pbegin = p++;
		// find s or d charcater
		while ((c = *p) != 0)
		{
			if ((c == L'0') || (c == L'l'))
			{
				p++;
				continue;
			}
			else if ((c != L's') && (c != L'd'))
			{
				p++;
				break;
			}
			pend = p;
			spec.Empty();
			spec.Append(pbegin, (int)(pend - pbegin + 1));
			conversionSpecifiersArray.Add(spec);
			break;
		}
	}
}

BOOL ResHelper::CompareCStringArrays(CStringArray &ar1, CStringArray& ar2)
{
	if (ar1.GetCount() != ar2.GetCount())
		return FALSE;
	int cnt = (int)ar1.GetCount();
	int i;
	for (i = 0; i < cnt; i++)
	{
		if (ar1[i].Compare(ar2[i]))
			return FALSE ;
	}
	return TRUE;
}

BOOL ResHelper::IsEnglishConfigured(CString &languageName)
{
	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	CString lastFolderName;
	CString language = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"language");
	
	if (language.IsEmpty())  // Special case;Nasty complication
	{
		CString lastFolderName;
		CString lastLanguageFolderPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"languageFolderPath");
		_ASSERTE(lastFolderName.IsEmpty());
		languageName = L"english";
		return TRUE;
	}
	else
	{
		languageName = language;
		return FALSE;
	}
}

CString ResHelper::GetLanguageCode(CString& languageName)
{
	CString langaugeCode;
	if (languageName.CompareNoCase(L"English") == 0)
		return L"en";
	else if (languageName.CompareNoCase(L"italian") == 0)
		return L"it";
	else if (languageName.CompareNoCase(L"spanish") == 0)
		return L"es";
	else  if (languageName.CompareNoCase(L"portuguese") == 0)
		return L"pt-PT";
	if (languageName.CompareNoCase(L"portuguese-brazil") == 0)
		return L"pt";
	else if (languageName.CompareNoCase(L"polish") == 0)
		return L"pl";
	else if (languageName.CompareNoCase(L"german") == 0)
		return L"de";
	else
		return langaugeCode;
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
#ifdef _DEBUG
	g_LoadMenuItemsInfo = TRUE;
	g_UpdateMenuItemsInfo = FALSE;
#endif
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

	_ASSERTE(wnd);
	if (wnd == 0)
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

// Iterates all child windows in the window defined by hwndParent window
void ResHelper::LoadDialogItemsInfo(CWnd *wnd, HWND hwndParent, int maxcnt, int iter)
{
#ifdef _DEBUG
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_LoadMenuItemsInfo == FALSE)
		return;

	// TODO: wnd param not used ???

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

		if (pw == 0)
			continue;

		pw->GetWindowText(strWindowText2);

		if (strClassName.CompareNoCase(L"Button") == 0)
		{
			AddItemInfo(strWindowText, strClassName);
		}
		else if (strClassName.CompareNoCase(L"Static") == 0)
		{
			AddItemInfo(strWindowText, strClassName);
		}
		else if (strClassName.CompareNoCase(L"Edit") == 0)
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
			// TODO: can we do more ???
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

	_ASSERTE(wnd);
	if (wnd == 0)
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

	// TODO: wnd param not used ??

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

		if (pw == 0)
			continue;

		CString wtext;
		pw->GetWindowText(wtext);

		if (strClassName.CompareNoCase(L"Button") == 0)
		{
			if (newWndText.GetLength())
				pw->SetWindowText(newWndText);
		}
		else if (strClassName.CompareNoCase(L"Static") == 0)
		{
			if (newWndText.GetLength())
				pw->SetWindowText(newWndText);
		}
		else if (strClassName.CompareNoCase(L"Edit") == 0)
		{
			;//if (nID == IDC_DATA_FOLDER_INTRO_1)
				//pw->SetWindowText(newWndText);
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
				int deb = 0; // keep old text
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

// May need to leverage  ::FormatMessage to allow any order of format specs in the format string independed 
// from the order of  args passed to formatting function
// May need to implement in addition to standard CString Format function
BOOL ResHelper::TranslateString(CString& str)
{
	if (ResHelper::g_UpdateMenuItemsInfo)
	{
		CString translatedText;
		BOOL retFind = ResHelper::DetermineString(str, translatedText);
		if (!translatedText.IsEmpty())
		{
			str = translatedText;
			return TRUE;
		}
		else
			return FALSE;
	}
	return TRUE;
}

// Format example fmt = L"arg1=%1 arg2=%2 arg3= %1"
// TranslateMsg(fmt, L"arg1", L"arg2");
// output "arg1=arg1 arg2=arg2 arg3=arg1
// FormatMessage is tricky if you try use spec modifiers suchas width, etc

CString ResHelper::TranslateMsg(const wchar_t* fmt, ...)
{
	CString mesg;
	CString tfmt(fmt);
	TranslateString(tfmt);

	const wchar_t* strfmt = (LPCWSTR)tfmt;

	// create va_list to pass ::FormatMessage
	va_list argList;
	va_start(argList, fmt);

	wchar_t* lpBuffer;  // pointer to buffer with generated message allocated by FormatMessage
	const DWORD dwMessageId = 0;
	const DWORD dwLanguageId = 0;
	DWORD nSize = 0; //  FormatMessage allocates message buffer

	DWORD retcnt = ::FormatMessage(FORMAT_MESSAGE_FROM_STRING | FORMAT_MESSAGE_ALLOCATE_BUFFER,
		fmt, dwMessageId, dwLanguageId, (LPWSTR)&lpBuffer, nSize, &argList);

	if (retcnt != 0)
	{
		_ASSERTE(lpBuffer != 0);
		mesg.Append(lpBuffer, retcnt);
	}

	LocalFree(lpBuffer);  // should we allocate buffer for message
	va_end(argList);

	return mesg;
}

CString ResHelper::TranslateMsg(CString* fmt, ...)
{
	CString mesg;
	CString tfmt(*fmt);
	TranslateString(tfmt);
	TranslateString(*fmt);

	const wchar_t* strfmt = (LPCWSTR)tfmt;

	// create va_list to pass ::FormatMessage
	va_list argList;
	va_start(argList, fmt);

	//wchar_t* arg1 = va_arg(argList, wchar_t*);
	//wchar_t* arg2 = va_arg(argList, wchar_t*);

	//for (int i = 0;  )

	wchar_t* lpBuffer = 0;  // pointer to buffer with generated message allocated by FormatMessage
	const DWORD dwMessageId = 0;
	const DWORD dwLanguageId = 0;
	DWORD nSize = 0; //  FormatMessage allocates message buffer

	DWORD retcnt = ::FormatMessage(FORMAT_MESSAGE_FROM_STRING | FORMAT_MESSAGE_ALLOCATE_BUFFER,
		(LPCWSTR)(*fmt), dwMessageId, dwLanguageId, (LPWSTR)&lpBuffer, nSize, &argList);

	if ((retcnt != 0) && lpBuffer)
	{
		_ASSERTE(lpBuffer != 0);
		mesg.Append(lpBuffer, retcnt);
	}

	LocalFree(lpBuffer);  // should we allocate buffer for message
	va_end(argList);

	return mesg;
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

ResourceInfo* ResHelper::AddItemInfo(CString & label, CString& controlName, int contributorID)
{
	CString key;
	key.Format(L"%s", label);
	hashsum_t hsum = g_resInfoMap.hash(&key);

	ResourceInfo *rinfo = g_resInfoMap.find(&key, hsum);
	if (rinfo == 0)
	{
		ResourceInfo* rinfo = new ResourceInfo(key, controlName, contributorID);
		g_resInfoMap.insert(hsum, rinfo);
		g_resInfoArray.Add(rinfo);
	}

	return rinfo;
}
ResourceInfo* ResHelper::MergeAddItemInfo(CString& label, CString& controlName, int contributorID)
{
	CString key;
	key.Format(L"%s", label);
	hashsum_t hsum = g_mergedResInfoMap.hash(&key);

	ResourceInfo* rinfo = g_mergedResInfoMap.find(&key, hsum);
	if (rinfo == 0)
	{
		ResourceInfo* rinfo = new ResourceInfo(key, controlName, contributorID);
		g_mergedResInfoMap.insert(hsum, rinfo);
		g_mergedResInfoArray.Add(rinfo);
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


	CString folderPath = ResHelper::resourceRootPath;
	CString resourceFilePath;

	resourceInfoArray = &resInfoArray;
	if (printAsUnicode)
		resourceFilePath = folderPath + L"source.txt";
	else
		resourceFilePath = folderPath + L"unsorted_" + resourceFile;

	hResourceFile = FileUtils::FileOpen(resourceFilePath, errorTxt, TRUE);
	if (hResourceFile == INVALID_HANDLE_VALUE)
	{
		return -1;
	}

	const char* BOM_UTF16_LE = "\xFF\xFE";
	if (hResourceFile)
	{
		nNumberOfBytesToWrite = 2;
		nNumberOfBytesWritten = 0;

		retval = FileUtils::Write2File(hResourceFile, BOM_UTF16_LE, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
	}

	int contributorID = NONE;
	for (int i = 0; i < resourceInfoArray->GetCount(); i++)
	{
		ResourceInfo* rinfo = (*resourceInfoArray)[i];

		if (rinfo->m_contributorID != contributorID)
		{
			CString contributorName;
			CString controlName = rinfo->m_controlName;
			if (rinfo->m_contributorID == CONTRIB_CONTROLS)
				controlName = L"Runtime Collection";

			contributorName.Format(L"\n#### %s\n", controlName);
			nNumberOfBytesToWrite = contributorName.GetLength() * 2;
			nNumberOfBytesWritten = 0;
			retval = FileUtils::Write2File(hResourceFile, (LPCWSTR)contributorName, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
			contributorID = rinfo->m_contributorID;
		}

		TRACE(L"%8s %s\n", rinfo->m_controlName, rinfo->m_label);
		//resInfoTxt.Format(L"\"%s\"\r\n", rinfo->m_label);

		resInfoTxt.Format(L"[SB:%04d][\"%s\"][]\n", i, rinfo->m_label);

		if (rinfo->m_label.Find(L"Date format:") >= 0)
			int deb = 1;

		if (hResourceFile != INVALID_HANDLE_VALUE)
		{
			nNumberOfBytesToWrite = resInfoTxt.GetLength() * 2;
			nNumberOfBytesWritten = 0;
			retval = FileUtils::Write2File(hResourceFile, (LPCWSTR)resInfoTxt, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
			int deb = 1;
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


void ResHelper::MergeControlsResInfo()
{
	ResourceInfo* info;
	int cnt = (int)g_resInfoArray.GetCount();
	int i;
	for (i = 0; i < cnt; i++)
	{
		info = g_resInfoArray[i];
		int contributorID = CONTRIB_CONTROLS;
		MergeAddItemInfo(info->m_label, info->m_controlName, contributorID);
	}
}

void ResHelper::MergeAllResInfo()
{
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_LoadMenuItemsInfo == FALSE)
		return;

	FileUtils::CreateDir(resourceRootPath);

	CString last_sourceFileFilePath = resourceRootPath + ResHelper::last_sourceFileName;
	if (!FileUtils::PathFileExist(last_sourceFileFilePath))
		int deb = 1;
	else if (ResHelper::IsFileUTF16LE_BOM(last_sourceFileFilePath))
	{
		int contributorID = CONTRIB_LAST_SOURCE;
		MergeResInfoFile(last_sourceFileFilePath, contributorID);
		int deb = 1;
	}

	CString rcFilePath = resourceRootPath + ResHelper::rcFileName;
	if (!FileUtils::PathFileExist(rcFilePath))
		int deb = 1;
	else if (ResHelper::IsFileUTF16LE_BOM(rcFilePath))
	{
		int contributorID = CONTRIB_RC;
		MergeResInfoFile(rcFilePath, contributorID);
		int deb = 1;
	}

	CString extraFilePath = resourceRootPath + ResHelper::extraFileName;
	if (!FileUtils::PathFileExist(extraFilePath))
		int deb = 1;
	else if (ResHelper::IsFileUTF16LE_BOM(extraFilePath))
	{
		int contributorID = CONTRIB_EXTRA;
		MergeResInfoFile(extraFilePath, contributorID);
		int deb = 1;
	}

	MergeControlsResInfo();

	MboxMail::LoadHintTextResource();

	CString msgsFilePath = resourceRootPath + ResHelper::msgsFileName;
	if (!FileUtils::PathFileExist(msgsFilePath))
		int deb = 1;
	else if (ResHelper::IsFileUTF16LE_BOM(msgsFilePath))
	{
		MergeResInfoFile(msgsFilePath);
		int deb = 1;
	}
}

// expensive
int ResHelper::CStdioFileReadLine(CStdioFile& resFile, CString& strLine)
{
	wchar_t wbuff[2048+1];
	int wbufflen = 2048;
	wchar_t* wp = wp = resFile.ReadString(wbuff, wbufflen);
	if (wp == 0)
		return -1;
	strLine.Empty();
	strLine.Append(wp);
	_ASSERTE(strLine.GetLength() <= 2048);
	return strLine.GetLength();
}

int ResHelper::GetText(CStdioFile &resFile, CString &strLine, BOOL translated, int& nTextId)
{
	static int idlen = 4;
	static wchar_t* prefix1Source = L"[SB:";
	static wchar_t* prefix1Translation = L"[TB:";
	static wchar_t* prefix2 = L"][\"";
	static wchar_t* suffix = L"\"][]";
	static int prefix1Len = (int)wcslen(prefix1Source);
	static int prefix2Len = (int)wcslen(prefix2);
	static int suffixLen = (int)wcslen(suffix);
	static int prefixLen = prefix1Len + idlen + prefix2Len;
	static int overhead = prefixLen + suffixLen;

	wchar_t *prefix1 = prefix1Source;
	if (translated)
		prefix1 = prefix1Translation;

	int retval;
	CString text = strLine;
	strLine.Empty();

	CString textId;

	const wchar_t* first = (LPCWSTR)text;
	const wchar_t* first2 = (LPCWSTR)text + prefix1Len + idlen;
	int retval1;
	int retval2;
	while ((retval1 = wcsncmp(first, prefix1, prefix1Len) != 0) || (retval2 = wcsncmp(first2, prefix2, prefix2Len) != 0))
	{
#if 0
		text.TrimRight(L"\r\n ");
		int slen = text.GetLength();
		const wchar_t* last = (LPCWSTR)text + slen - suffixLen;
		if ((retval = wcsncmp(last, suffix, suffixLen)) == 0)
		{
			int deb = 1;
			_ASSERTE(0);
		}
#endif
		if (ResHelper::CStdioFileReadLine(resFile, text) < 0)
		{
			return 0;
		}
		// Needed in case CString text allocates new memory buffer
		first = (LPCWSTR)text;
		first2 = (LPCWSTR)text + prefix1Len + idlen;
	}
	const wchar_t* pId = first + prefix1Len;
	textId.Append(pId, 4);
	nTextId = _wtoi(textId);

	CString strLineSave;
	int strLineLen;
	int j;
	for (j = 0; j < 10000; j++)
	{
		strLineSave = text;

		text.TrimRight(L"\r\n ");
		int slen = text.GetLength();

		const wchar_t* first = (LPCWSTR)text + prefixLen;
		const wchar_t* last = (LPCWSTR)text + slen - suffixLen;

		if ((retval = wcsncmp(last, suffix, suffixLen)) != 0)
		{
			if (j == 0)
			{
				const wchar_t* first = (LPCWSTR)strLineSave + prefixLen;
				strLineLen = strLineSave.GetLength();
				strLine.Append(first, strLineLen - prefixLen);
			}
			else
			{
				const wchar_t* first = (LPCWSTR)strLineSave;
				strLineLen = strLineSave.GetLength();
				strLine.Append(first, strLineLen);
			}

			int len = ResHelper::CStdioFileReadLine(resFile, text);
			if (len < 0)
				return 1;
#if 0
			const wchar_t* first_text = (LPCWSTR)text;
			const wchar_t* first2_text = (LPCWSTR)text + prefix1Len + idlen;
			if ((retval1 = wcsncmp(first_text, prefix1, prefix1Len) == 0) && (retval2 = wcsncmp(first2_text, prefix2, prefix2Len) == 0))
			{
				int deb = 1;
				_ASSERTE(0);
			}
#endif
		}
		else
		{
			if (j == 0)
			{
				strLine.Append(first, slen - overhead);
			}
			else
			{
				const wchar_t* first = (LPCWSTR)text;
				strLine.Append(first, slen - suffixLen);
			}
			break;
		}
	}

	return 1;
}

void ResHelper::MergeResInfoFile(CString &resFile, int contributorID)
{
	if (!FileUtils::PathFileExist(resFile))
		return;

	CString fileName;
	FileUtils::GetFileName(resFile, fileName);

	CString strLine;
	CString str;

	UINT nOpenFlags = CFile::modeRead | CFile::typeText | CFile::shareDenyNone | CFile::typeUnicode;
	//UINT nOpenFlags = CFile::modeRead | CFile::typeBinary | CFile::shareDenyNone | CFile::typeUnicode;

	CStdioFile file;
	CFileException exList;
	if (!file.Open(resFile, nOpenFlags, &exList))
	{
		return;
	}

	// CStdioFile is not BOM aware !!!! 
	// wchar_t BOM[1];
	// file.Read(&BOM[0], sizeof(wchar_t));  // or see below while loop

	CString text;
	int retlen;
	while ((retlen = ResHelper::CStdioFileReadLine(file, strLine)) >= 0)
	{
		// TODO: Implement RemoveBOM(), seek(2) ??
		const unsigned char* p = (unsigned char*)((LPCWSTR)strLine);
		if ((*p == 0xFF) && (*(p+1) == 0xFE))  // UTF16LE
		{
			CString str = strLine.Right(strLine.GetLength() - 1);
			strLine = str;
		}
		
		text = strLine;
		int nTextId;
		int retval = ResHelper::GetText(file, text, FALSE, nTextId);
		MergeAddItemInfo(text, fileName, contributorID);

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

	FileUtils::CreateDir(resourceRootPath);

	CString folderPath = ResHelper::resourceRootPath;

	CString resourceFilePath = folderPath;


	CString resFile1 = resourceFilePath + L"source.txt";
	CString resFile2 = resourceFilePath + L"translated_source.txt";
	CString languageFilePath = resourceFilePath + L"translationFile.txt";

#if 0
	CString fileUTF8 = resourceFilePath + L"sorted_resource-UTF8.txt";
	BOOL isUTF8 = IsFileUTF8(fileUTF8);

	fileUTF8 = resourceFilePath + L"german-UTF8.txt";
	BOOL isUTF8_2 = IsFileUTF8(fileUTF8);
#endif

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

	_ASSERTE((bom1 == TextEncoding::UTF16LE) && (bom2 == TextEncoding::UTF16LE));

	int ret = -1;
	if ((bom == ResHelper::TextEncoding::UTF16LE) && (bom2 == ResHelper::TextEncoding::UTF16LE))
	{
		ret = ResHelper::CreateLanguageFileFromUTF16LEEncodedFiles(resFile1, resFile2, languageFilePath);
	}
	return ret;

}

int ResHelper::CreateLanguageFileFromUTF16LEEncodedFiles(CString& resFile1, CString& resFile2, CString& languageFilePath)
{
	FileUtils::CreateDir(resourceRootPath);

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

			//resInfoText1.Format(L"[%s]\n", rinfo1->m_label);
			//resInfoText2.Format(L"%s\n\n", rinfo2->m_label);

			resInfoText1.Format(L"[SB:%04d][\"%s\"][]\n", i, rinfo1->m_label);
			resInfoText2.Format(L"[TB:%04d][\"%s\"][]\n\n", i, rinfo2->m_label);

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

// TODO: needs update, don't use until further notice
int ResHelper::ResortLanguageFile()
{
	CString label;
	CString translatedLabel;

	FileUtils::CreateDir(resourceRootPath);

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
		DWORD lastErr = ::GetLastError();
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exList);

		CString txt;
		CString fmt = L"Could not open list file \"%s\" file.\n%s";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, resFile, exErrorStr);

		TRACE(L"ResHelper: %s\n", txt);
		//errorText = txt;

		CFileStatus rStatus;
		BOOL ret = file.GetStatus(rStatus);

		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return;
	}

	int nLastTextId = -1;
	int nTextId;

	CString text;
	int retlen;
	while ((retlen = ResHelper::CStdioFileReadLine(file, strLine)) >= 0)
	{
		// TODO: Implement RemoveBOM()
		const unsigned char* p = (unsigned char*)((LPCWSTR)strLine);
		if ((*p == 0xFF) && (*(p + 1) == 0xFE))  // UTF16LE
		{
			CString str = strLine.Right(strLine.GetLength() - 1);
			strLine = str;
		}

		text = strLine;
		int retval = ResHelper::GetText(file, text, FALSE, nTextId);
		if (retval == 0)
			break;

		if (nTextId == 394)
			int deb = 1;

		if (nLastTextId >= 0)
		{
			if (nTextId != (nLastTextId + 1))
			{
				_ASSERTE(0);
			}
			else
				nLastTextId = nTextId;
		}
		else
			nLastTextId = nTextId;

		ResourceInfo* rinfo = new ResourceInfo(text, controlName);
		resArray.Add(rinfo);
		if (resArray.GetCount()%100 == 0)
			int deb = 1;

		if (resArray.GetCount() % 50 == 0)
			int deb = 1;

		if (resArray.GetCount() % 10 == 0)
			int deb = 1;

		int deb = 1;
	}
	file.Close();
}

void ResHelper::LoadLanguageMap(CString& languageTranslationFilePath)
{
	CString out;
	out.Format(L"%d", g_UpdateMenuItemsInfo);
	ResHelper::MyTrace(L"LoadLanguageMap: %s\n", out);

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
	_ASSERTE(bom == ResHelper::TextEncoding::UTF16LE);

	ResHelper::MyTrace(L"LoadLanguageMap: %s\n", languageFile);

	// TODO: Convert to ResHelper::TextEncoding::UTF16LE first if necessary

	if (bom == ResHelper::TextEncoding::UTF16LE)
	{
		LoadLanguageMapFromFileF16LE(languageFile);
		int deb = 1;
	}
	else  // TODO: Ansi ???
	{
		int deb = 1;
	}
	return;
}

// This is the only valid function
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
	if (!file.Open(languageFile, nOpenFlags, &exList))
	{
		DWORD lastErr = ::GetLastError();

		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exList);

		CString txt;
		CString fmt = L"Could not open list file \"%s\" file.\n%s";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, languageFile, exErrorStr);

		TRACE(L"ResHelper: %s\n", txt);
		//errorText = txt;

		CFileStatus rStatus;
		BOOL ret = file.GetStatus(rStatus);

		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return;
	}


	int nLastSourceTextId = -1;
	int nSourceTextId;

	int nLastTranslatedTextId = -1;
	int nTranslatedTextId;

	CString esourceText;
	CString sourceText;
	CString etranslatedText;
	CString translatedText;
	int retlen;
	int cnt = 0;
	while ((retlen = ResHelper::CStdioFileReadLine(file, strLine)) >= 0)
	{
		// TODO: Implement RemoveBOM()
		const unsigned char* p = (unsigned char*)((LPCWSTR)strLine);
		if ((*p == 0xFF) && (*(p + 1) == 0xFE))  // UTF16LE
		{
			CString str = strLine.Right(strLine.GetLength() - 1);
			strLine = str;
		}

		esourceText = strLine;
		int retval = ResHelper::GetText(file, esourceText, FALSE, nSourceTextId);
		if (retval <= 0)
			break;

		if (nSourceTextId == 407)
		{
			int deb = 1;
		}
		if (nLastSourceTextId >= 0)
		{
			if (nSourceTextId != (nLastSourceTextId + 1))
			{
				_ASSERTE(0);
			}
			else
				nLastSourceTextId = nSourceTextId;
		}
		else
			nLastSourceTextId = nSourceTextId;

		ResHelper::UnescapeString(esourceText, sourceText);

		CString dd = L"mails in the selected list";
		if (sourceText.Find(dd) >= 0)
			int deb = 1;

		cnt++;


		if (ResHelper::CStdioFileReadLine(file, strLine) < 0)
			break;

		etranslatedText = strLine;
		retval = ResHelper::GetText(file, etranslatedText, TRUE, nTranslatedTextId);
		if (retval <= 0)
			break;

		if (nTranslatedTextId == 397)
			int deb = 1;

		if (nTranslatedTextId != nTranslatedTextId)
		{
			_ASSERTE(0);
		}

		if (nTranslatedTextId >= 0)
		{
			if (nTranslatedTextId != (nLastTranslatedTextId + 1))
			{
				_ASSERTE(0);
			}
			else
				nLastTranslatedTextId = nTranslatedTextId;
		}
		else
			nLastTranslatedTextId = nTranslatedTextId;

		ResHelper::UnescapeString(etranslatedText, translatedText);

		if (nTranslatedTextId == 590)
		{
			int deb = 1;
		}

		hashsum_t hsum = g_LanguageMap.hash(&sourceText);
		ResourceInfo* rinfo = g_LanguageMap.find(&sourceText, hsum);
		if (rinfo == 0)
		{
			CString specSourceList;
			CString specTraslatedList;

			ResHelper::FindConversionSpecifiers(esourceText, specSourceList);
			ResHelper::FindConversionSpecifiers(etranslatedText, specTraslatedList);
			if (specSourceList.GetLength() > 0)
				int deb = 1;
			if (specSourceList.Compare(specTraslatedList))
				int deb = 1;
				//_ASSERTE(0);

			ResourceInfo* rinfo = new ResourceInfo(sourceText, translatedText);
			g_LanguageMap.insert(hsum, rinfo);
		}
		else
			int deb = 1;

	}
	int deb = 1;
}

void ResHelper::RenumberLanguageFile()
{
	FileUtils::CreateDir(resourceRootPath);

	CString folderPath = ResHelper::resourceRootPath;
	CString resourceFilePath = folderPath;
	CString languageTranslationFilePath = resourceFilePath + L"translationFile.txt";

	int retval = ResHelper::RenumberLanguageFileF16LE(languageTranslationFilePath);
}

int ResHelper::RenumberLanguageFileF16LE(CString& languageTranslationFilePath)
{
	if (g_UpdateMenuItemsInfo == FALSE)  // Do I need separate control var ?
		return -1;

	CString controlName;
	CString strLine;
	CString str;
	CString strW;

	ResInfoMapType& resInfoMap =  g_mergedResInfoMap;
	ResInfoArrayType& resInfoArray = g_mergedResInfoArray;

	_ASSERTE(resInfoMap.count() == resInfoArray.GetCount());

	ReleaseResInfoArray(resInfoArray, TRUE);  // it will delete items
	ReleaseResInfoMap(resInfoMap, FALSE); // it doesn't delete item

	_ASSERTE(resInfoMap.count() == 0);
	_ASSERTE(resInfoArray.GetCount() == 0);

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
	if (!file.Open(languageFile, nOpenFlags, &exList))
	{
		DWORD lastErr = ::GetLastError();

		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exList);

		CString txt;
		CString fmt = L"Could not open list file \"%s\" file.\n%s";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, languageFile, exErrorStr);

		TRACE(L"ResHelper: %s\n", txt);
		//errorText = txt;

		CFileStatus rStatus;
		BOOL ret = file.GetStatus(rStatus);

		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}


	int nLastSourceTextId = -1;
	int nSourceTextId;

	int nLastTranslatedTextId = -1;
	int nTranslatedTextId;

	CString esourceText;
	CString sourceText;
	CString etranslatedText;
	CString translatedText;
	int retlen;
	int cnt = 0;
	while ((retlen = ResHelper::CStdioFileReadLine(file, strLine)) >= 0)
	{
		// TODO: Implement RemoveBOM()
		const unsigned char* p = (unsigned char*)((LPCWSTR)strLine);
		if ((*p == 0xFF) && (*(p + 1) == 0xFE))  // UTF16LE
		{
			CString str = strLine.Right(strLine.GetLength() - 1);
			strLine = str;
		}

		esourceText = strLine;
		int retval = ResHelper::GetText(file, esourceText, FALSE, nSourceTextId);
		if (retval <= 0)
			break;

		if (nSourceTextId == 407)
		{
			int deb = 1;
		}
		if (nLastSourceTextId >= 0)
		{
			if (nSourceTextId != (nLastSourceTextId + 1))
			{
				_ASSERTE(0);
			}
		}
		nLastSourceTextId = nSourceTextId;

		//ResHelper::UnescapeString(esourceText, sourceText);

		CString dd = L"mails in the selected list";
		if (sourceText.Find(dd) >= 0)
			int deb = 1;

		cnt++;


		if (ResHelper::CStdioFileReadLine(file, strLine) < 0)
			break;

		etranslatedText = strLine;
		retval = ResHelper::GetText(file, etranslatedText, TRUE, nTranslatedTextId);
		if (retval <= 0)
			break;

		if (nTranslatedTextId == 397)
			int deb = 1;

		if (nTranslatedTextId != nTranslatedTextId)
		{
			_ASSERTE(0);
		}

		if (nTranslatedTextId >= 0)
		{
			if (nTranslatedTextId != (nLastTranslatedTextId + 1))
			{
				_ASSERTE(0);
			}
		}
		nLastTranslatedTextId = nTranslatedTextId;

		//ResHelper::UnescapeString(etranslatedText, translatedText);

		if (nTranslatedTextId == 407)
		{
			int deb = 1;
		}

		// RESTODO Or call MergeAddItemInfo(CString& label, CString& controlName, int contributorID = 0);   ????
		hashsum_t hsum = resInfoMap.hash(&esourceText);
		ResourceInfo* rinfo = resInfoMap.find(&esourceText, hsum);
		if (rinfo == 0)
		{
			CString specSourceList;
			CString specTraslatedList;

			ResHelper::FindConversionSpecifiers(esourceText, specSourceList);
			ResHelper::FindConversionSpecifiers(etranslatedText, specTraslatedList);
			if (specSourceList.GetLength() > 0)
				int deb = 1;
			if (specSourceList.Compare(specTraslatedList))
				_ASSERTE(0);

			ResourceInfo* rinfo = new ResourceInfo(esourceText, etranslatedText);
			resInfoMap.insert(hsum, rinfo);
			resInfoArray.Add(rinfo);
		}
		else
			int deb = 1;

	}
	file.Close();

	int deb = 1;

	CString folderPath;
	CString fileName;
	FileUtils::GetFolderPathAndFileName(languageFile, folderPath, fileName);
	CString new_languageFile = folderPath + L"new_" + fileName;
	
	DWORD nNumberOfBytesToWrite = 0;
	DWORD nNumberOfBytesWritten = 0;

	int retval;
	CString errorText;

	hResourceFile = FileUtils::FileOpen(new_languageFile, errorText, TRUE);
	if (hResourceFile == INVALID_HANDLE_VALUE)
	{
		return -1;
	}

	BYTE BOMF16LE[2] = { 0xFF, 0xFE };

	nNumberOfBytesToWrite = 2;
	nNumberOfBytesWritten = 0;
	retval = FileUtils::Write2File(hResourceFile, BOMF16LE, nNumberOfBytesToWrite, &nNumberOfBytesWritten);


	CString resInfoText1;
	CString resInfoText2;
	int i;
	cnt = (int)resInfoArray.GetCount();
	for (i = 0; i < cnt; i++)
	{
		BOOL unicodeFile = TRUE;
		if (unicodeFile == TRUE)
		{
			ResourceInfo* rinfo1 = resInfoArray[i];
			ResourceInfo* rinfo2 = rinfo1;

			resInfoText1.Format(L"[SB:%04d][\"%s\"][]\n", i, rinfo1->m_label);
			resInfoText2.Format(L"[TB:%04d][\"%s\"][]\n\n", i, rinfo2->m_controlName);

			if (hResourceFile != INVALID_HANDLE_VALUE)
			{
				nNumberOfBytesToWrite = resInfoText1.GetLength() * 2;
				nNumberOfBytesWritten = 0;
				retval = FileUtils::Write2File(hResourceFile, (LPCWSTR)resInfoText1, nNumberOfBytesToWrite, &nNumberOfBytesWritten);

				nNumberOfBytesToWrite = resInfoText2.GetLength() * 2;
				nNumberOfBytesWritten = 0;
				retval = FileUtils::Write2File(hResourceFile, (LPCWSTR)resInfoText2, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
				int deb = 1;
			}
		}
	}
	FileUtils::FileClose(hResourceFile);
	return 1;
}

void ResHelper::SplitTranslationFile()
{
#if 0
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_LoadMenuItemsInfo == FALSE)
		return;
#endif

#ifdef _DEBUG
	FileUtils::CreateDir(resourceRootPath);

	CString languageTranslationFilePath = resourceRootPath + L"split.txt";
	ResHelper::SplitTranslationFile(languageTranslationFilePath);
#endif
}

int ResHelper::SplitTranslationFile(CString& languageTranslationFilePath)
{
#if 0
	_ASSERTE((g_LoadMenuItemsInfo == FALSE) || (g_UpdateMenuItemsInfo == FALSE));
	if (g_LoadMenuItemsInfo == FALSE)  // Do I need separate control var ?
		return 1;
#endif

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
	if (!file.Open(languageFile, nOpenFlags, &exList))
	{
		DWORD lastErr = ::GetLastError();

		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exList);

		CString txt;
		CString fmt = L"Could not open list file \n\"%s\"\n\n%s";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, languageFile, exErrorStr);

		TRACE(L"ResHelper: %s\n", txt);
		//errorText = txt;

		CFileStatus rStatus;
		BOOL ret = file.GetStatus(rStatus);

		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CString translationFolderPath;
	CString translationFileName;
	FileUtils::GetFolderPathAndFileName(languageTranslationFilePath, translationFolderPath, translationFileName);


	CString sourceFilePath = translationFolderPath + L"splitS.txt";
	CString errorText;
	HANDLE hSourceFile = FileUtils::FileOpen(sourceFilePath, errorText, TRUE);
	if (hSourceFile == INVALID_HANDLE_VALUE)
	{
		file.Close();
		return -1;
	}

	CString translatedSourceFilePath = translationFolderPath + L"splitT.txt";
	HANDLE hTranslatedSourceFile = FileUtils::FileOpen(translatedSourceFilePath, errorText, TRUE);
	if (hTranslatedSourceFile == INVALID_HANDLE_VALUE)
	{
		file.Close();
		FileUtils::FileClose(hResourceFile);
		return -1;
	}

	const char* BOM_UTF16_LE = "\xFF\xFE";
	DWORD nNumberOfBytesToWrite = 2;
	DWORD nNumberOfBytesWritten = 0;
	int retval;

	retval = FileUtils::Write2File(hSourceFile, BOM_UTF16_LE, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
	if (retval < 0)
	{
		file.Close();
		FileUtils::FileClose(hSourceFile);
		FileUtils::FileClose(hTranslatedSourceFile);
		return -1;
	}

	nNumberOfBytesWritten = 0;
	retval = FileUtils::Write2File(hTranslatedSourceFile, BOM_UTF16_LE, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
	if (retval < 0)
	{
		file.Close();
		FileUtils::FileClose(hSourceFile);
		FileUtils::FileClose(hTranslatedSourceFile);
		return -1;
	}

	CString text;
	CString translatedText;
	CString outText;
	int retlen;
	int cnt = 0;
	int i = 411;
	i = 0;
	while ((retlen = ResHelper::CStdioFileReadLine(file, strLine)) >= 0)
	{
		// TODO: Implement RemoveBOM()
		const unsigned char* p = (unsigned char*)((LPCWSTR)strLine);
		if ((*p == 0xFF) && (*(p + 1) == 0xFE))  // UTF16LE
		{
			CString str = strLine.Right(strLine.GetLength() - 1);
			strLine = str;
		}

		text = strLine;
		int nTextId;
		int retval = ResHelper::GetText(file, text, FALSE, nTextId);
		if (text.IsEmpty())
		{
			continue;
		}

		outText.Format(L"[SB:%04d][\"%s\"][]\n", i, text);

		nNumberOfBytesToWrite = outText.GetLength() * 2;
		nNumberOfBytesWritten = 0;
		retval = FileUtils::Write2File(hSourceFile, (LPCWSTR)outText, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
		if (retval < 0)
			break;

		if (ResHelper::CStdioFileReadLine(file, strLine) < 0)
		{
			continue;
		}

		translatedText = strLine;
		int nTranslatedTextId;
		retval = ResHelper::GetText(file, translatedText, TRUE, nTranslatedTextId);
		if (translatedText.IsEmpty())
		{
			continue;
		}

		outText.Format(L"[TB:%04d][\"%s\"][]\n", i, translatedText);

		nNumberOfBytesToWrite = outText.GetLength() * 2;
		nNumberOfBytesWritten = 0;
		retval = FileUtils::Write2File(hTranslatedSourceFile, (LPCWSTR)outText, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
		if (retval < 0)
			break;

		int deb = 1;
		i++;

	}
	file.Close();
	FileUtils::FileClose(hSourceFile);
	FileUtils::FileClose(hTranslatedSourceFile);
	int deb = 1;
	return 1;
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

	FileUtils::CreateDir(resourceRootPath);

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
	int cnt4 = ResHelper::PrintResInfoArray(ResHelper::g_mergedResInfoArray, sort, printAsUnicode);


#if 0
	printAsUnicode = FALSE;
	int cnt5 = ResHelper::PrintResInfoArray(ResHelper::g_resInfoArray, sort, printAsUnicode);

	sort = TRUE;
	printAsUnicode = TRUE;
	int cnt6 = ResHelper::PrintResInfoArray(ResHelper::g_resInfoArray, sort, printAsUnicode);
	printAsUnicode = FALSE;
	int cnt7 = ResHelper::PrintResInfoArray(ResHelper::g_resInfoArray, sort, printAsUnicode);
#endif

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
		//if (nID && (nID != IDC_DATA_FOLDER_INTRO_1))
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


int ResHelper::ReleaseResInfoMap(ResInfoMapType& resInfoMap, BOOL deleteItems)
{
	// internal hash array is not dleted only items
	// internal hash array is deleted when IHashMap instance is deleted
	if (deleteItems == FALSE)
	{
		resInfoMap.clear();
	}
	else
	{
		ResInfoMapType::IHashMapIter iter = resInfoMap.first();
		for (; !resInfoMap.last(iter); )
		{
			ResourceInfo* l = iter.element;
			if (deleteItems)
				resInfoMap.remove(iter);  // delete or remove
			delete l;
		}
		resInfoMap.clear();  // not neeeded
	}

	return resInfoMap.count();
}

void ResHelper::ReleaseResInfoArray(ResInfoArrayType &resInfoArray, BOOL deleteItems)
{
	// internal  array is not dleted only items
	// internal hash array is deleted when CStringArray instance is deleted
	if (deleteItems == FALSE)
	{
		resInfoArray.RemoveAll();
	}
	else
	{
		for (int i = 0; i < resInfoArray.GetCount(); i++)
		{
			ResourceInfo* rinfo = resInfoArray[i];
			if (deleteItems)  // remove or delete
				delete rinfo;
			resInfoArray[i] = 0;
		}
		resInfoArray.RemoveAll();
	}
}

void ResHelper::ReleaseResources()
{
	ReleaseResInfoArray(g_resInfoArray, TRUE);
	ReleaseResInfoArray(resArray1, TRUE);
	ReleaseResInfoArray(resArray2, TRUE);
	ReleaseResInfoArray(g_mergedResInfoArray, TRUE);

	ReleaseResInfoMap(g_resInfoMap, FALSE);
	ReleaseResInfoMap(g_mergedResInfoMap, FALSE);
	ReleaseResInfoMap(g_LanguageMap, TRUE);
}

ResHelper::TextEncoding ResHelper::GetFileBOM(LPCWSTR lpszFileName)
{
	CStdioFile File;
	if (!File.Open(lpszFileName, CFile::modeRead | CFile::typeBinary))
		return TextEncoding::NONE;  /// TextEncoding::ERROR ??

	BYTE string[4] = { 0, 0, 0, 0 };
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

BOOL ResHelper::IsFileUTF16LE_BOM(LPCWSTR lpszFileName)
{
	ResHelper::TextEncoding BOM = ResHelper::GetFileBOM(lpszFileName);
	if (BOM == TextEncoding::UTF16LE)
		return TRUE;
	else
		return FALSE;
}

// RESTO: incoplete . delete ?
int ResHelper::GetCodePageFromFile(LPCWSTR filePath)
{
	CStdioFile file;
	CFileException exList;
	UINT nOpenFlags = CFile::modeRead | CFile::typeText | CFile::shareDenyNone;
	if (!file.Open(filePath, nOpenFlags, &exList))
	{
		DWORD lastErr = ::GetLastError();

		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exList);

		CString txt;
		CString fmt = L"Could not open list file \n\"%s\"\n\n%s";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, filePath, exErrorStr);

		TRACE(L"ResHelper: %s\n", txt);
		//errorText = txt;

		CFileStatus rStatus;
		BOOL ret = file.GetStatus(rStatus);

		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return FALSE;
	}

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
		DWORD lastErr = ::GetLastError();

		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exList);

		CString txt;
		CString fmt = L"Could not open list file \n\"%s\"\n\n%s";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, filePath, exErrorStr);

		TRACE(L"ResHelper: %s\n", txt);
		//errorText = txt;

		CFileStatus rStatus;
		BOOL ret = file.GetStatus(rStatus);


		//HWND h = NULL; // we don't have any window yet
		HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		
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

// enable/disable MyTrace
//#define MY_TRACE

void ResHelper::MyTrace(CString& fmt, CString &txt)
{
#ifdef MY_TRACE
	CString out;
	out.Format(fmt, txt);
	::OutputDebugString(out);
#endif
}

void ResHelper::MyTrace(CString& txt)
{
#ifdef MY_TRACE
	::OutputDebugString(txt);
#endif
}

void ResHelper::MyTrace(const wchar_t* fmt, const wchar_t* txt)
{
#ifdef MY_TRACE
	CString out;
	out.Format(CString(fmt), CString(txt));
	::OutputDebugString(out);
#endif
}

void ResHelper::MyTrace(const wchar_t* txt)
{
#ifdef MY_TRACE
	::OutputDebugString(txt);
#endif
}

#include <stdio.h>
#include <stdarg.h>
void ResHelper::MyVaTrace(const wchar_t* format, ...)
{
#ifdef MY_TRACE
	wchar_t buffer[1000];

	va_list argptr;
	va_start(argptr, format);
	wvsprintfW(buffer, format, argptr);
	va_end(argptr);

	OutputDebugString(buffer);
#endif
}

void ResHelper::UnescapeString(const CString& input, CString &output)
{
	output.Empty();
	int inputLength = input.GetLength();
	output.Preallocate(inputLength);  // to avoid realloc

	int index;
	for (index = 0; index < input.GetLength(); index++)
	{
		if (input[index] == '\\' && index + 1 < inputLength) // Check for escape character and ensure it's not at the end
		{
			switch (input[index + 1]) {
			case 'r': // Replace \r with carriage return
				output += '\r';
				index++;
				break;

			case 'n': // Replace \n with newline
				output += '\n';
				index++;
				break;

			case 't': // Replace \t with tab
				output += '\t';
				index++;
				break;

			case '\\': // Replace double backslash
				output += '\\';
				index++;
				break;

			case '"': // Replace \" with "
				output += '"';
				index++;
				break;

			default:
				output += input[index];
				break; // In case of an unknown escape sequence, keep the original
			}
		}
		else
		{
			output += input[index]; // Copy the character as is
		}
	}

	return;
}

