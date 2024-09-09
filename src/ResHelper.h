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
#include "IHashTable.h"

#pragma once

typedef CArray<CString, CString> StringArray;

#if 0
sort(MboxMail::s_mails.GetData(), MboxMail::s_mails.GetData() + MboxMail::s_mails.GetSize(),
	[](MboxMail* o1, MboxMail* o2) {
		return cmpMbox(o1, o2);
	}
);

StringArray strArray;

BOOL ascedingOrder = TRUE;
std::sort(strArray.GetData(), strArray.GetData() + strArray.GetSize(), ascedingOrder);

#endif

class ResourceInfo
{
public:
	ResourceInfo(CString& label, CString &controlName) { m_label = label; m_controlName = controlName;  }
	dlink_node<ResourceInfo> m_hashMapLink;
	CString m_label;
	CString m_controlName;
};

struct ResourceInfoHelper
{
	hashsum_t operator()(const CString* key) const
	{
		hashsum_t hashsum = StrHash((const char*)(LPCWSTR)*key, key->GetLength()*2);
		return hashsum;
	}
	hashsum_t operator()(const ResourceInfo* key) const
	{
		hashsum_t hashsum = StrHash((const char*)(LPCWSTR)key->m_label, key->m_label.GetLength()*2);
		return hashsum;
	}
	bool operator()(ResourceInfo* key1, ResourceInfo* key2) const
	{
		if (key1->m_label == key2->m_label)
			return true;
		else
			return false;
	}
	bool operator()(CString* key1, CString* key2) const
	{
		if (*key1 == *key2)
			return true;
		else
			return false;
	}
	bool operator()(CString* key1, ResourceInfo* key2) const
	{
		if (*key1 == key2->m_label)
			return true;
		else
			return false;
	}
};

using ResInfoMapType = IHashMap<CString, ResourceInfo, ResourceInfoHelper, ResourceInfoHelper, &ResourceInfo::m_hashMapLink>;
using ResInfoArrayType = CArray<ResourceInfo*, ResourceInfo*>;

#define WINDOW_TEXT_LENGTH  4092
#define CLASS_NAME_LENGTH  4092

class ResHelper
{
public:
	static CString resourceFile;
	static HANDLE hResourceFile;
	static int maxWindowTextLength;
	static wchar_t windowText[WINDOW_TEXT_LENGTH+1];
	static int maxClassNameLength;
	static wchar_t className[CLASS_NAME_LENGTH+1];

	static void ReleaseResources();

	static ResourceInfo* AddItemInfo(CString & label, CString &controlName);
	static ResourceInfo* AddItemInfo(CString& label);
	static int PrintResInfoMap(ResInfoMapType& resStringMap);
	static int PrintResInfoArray(ResInfoArrayType& resInfoArray);
	static int SortResInfoArray(ResInfoArrayType& resInfoArray);

	static ResInfoArrayType resInfoArray;
	static ResInfoMapType resInfoMap;
	static int PopulateResStringList() { return 0; }

	static void LoadDialogItemsInfo(HWND hwndParent);
	static void LoadDialogItemsInfo(HWND hwndParent, int maxcnt, int iter);
	static void LoadMenuItemsInfo(CMenu* menu);
	static void LoadMenuItemsInfo(CMenu* menu, int index);
	static void LoadToolBarItemsInfo(CToolBar* tbar);

	static void IterateWindowChilds(HWND hwndParent);
	static void IterateWindowChilds(HWND hwndParent, int maxcnt, int iter);
	static void IterateMenuItems(CMenu* menu);
	static void IterateMenuItems(CMenu* menu, int index);
	static void IterateMenuItemsSetPopMenuData(CMenu* menu, int index);
	static BOOL GetMenuItemString(CMenu* menu, UINT nIDItem, CString& rString, UINT nFlags);
protected:
};



