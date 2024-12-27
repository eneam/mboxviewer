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
#include "SimpleString.h"

#pragma once

class SimpleMemoryFile
{
public:
	SimpleMemoryFile();
	~SimpleMemoryFile();
	BOOL Open(CString& filePath);
	int ReadString(CStringA &str);
	int ReadString(CString &str);

	SimpleString m_buffer;
	int  m_position;  // _int64 ??
};

class ResourceInfo
{
public:
	ResourceInfo(CString& label, CString &controlName) { m_label = label; m_controlName = controlName;  }
	ResourceInfo(CStringA& labelA, CString& controlName) { m_labelA = labelA; m_controlName = controlName; }

	dlink_node<ResourceInfo> m_hashMapLink;
	CString m_label;
	CStringA m_labelA;
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

	enum TextEncoding
	{
		NONE,
		ASCII,
		ANSI,
		UTF8,
		UTF16LE,
		UTF16BE,
		UTF32LE,
		UTF32BE
	};

	static BOOL g_LoadMenuItemsInfo;  // Traverse dialog & menu controles and create resource filke. i.e. collect text from all items/sub-controls
	static BOOL g_UpdateMenuItemsInfo;  // Update all text using selected translation file

	static CString resourceFile;
	static HANDLE hResourceFile;
	static int maxWindowTextLength;
	static wchar_t windowText[WINDOW_TEXT_LENGTH+1];
	static int maxClassNameLength;
	static wchar_t className[CLASS_NAME_LENGTH+1];
	//
	static CString wndText;

	static ResHelper::TextEncoding GetFileBOM(LPCWSTR lpszFileName);  // byte order mark
	static ResHelper::TextEncoding ResHelper::GetBOM(unsigned char* string, int slen);
	static int GetCodePageFromFile(LPCWSTR lpszFileName);  // find PAGE_CODE NNN  in the file until line with first [ character

	// Guessing functions
	static BOOL IsTextUTF8Done(BOOL retcode);
	static BOOL IsTextUTF8(const char* data, size_t size, BOOL &isASCII);
	static int IsFileUTF8(LPCWSTR lpszFileName);
	static TextEncoding GetTextEncoding(char* string, int slen, TextEncoding &BOM /* byteOrderMark*/);
	static int FindCodePageFromFile(LPCWSTR lpszFileName);

	static void ReleaseResources();
	static void ReleaseResInfoArray(ResInfoArrayType& resInfoArray);
	static int ReleaseResInfoMap(ResInfoMapType& resInfoMap);

	static ResourceInfo* AddItemInfo(CString & label, CString &controlName);
	static ResourceInfo* AddItemInfo(CString& label);

	static int PrintResInfo();
	static int PrintResInfoMap(ResInfoMapType& resStringMap);
	static int PrintResInfoArray(ResInfoArrayType& resInfoArray, BOOL mustSort, BOOL printAsUnicode);
	static int SortResInfoArray(ResInfoArrayType& resInfoArray);

	static void MergeAllResInfo();
	static void MergeResInfoFile(CString& resFile);

	// Create language file by merigin info from .rc file and manually discovered info
	static int CreateLanguageFile();

	static int CreateLanguageFileFromANSIEncodedFiles(CString& resFile1, CString& resFile2, CString& languageFilePath);
	static int CreateLanguageFileFromUTF8EncodedFiles(CString& resFile1, CString& resFile2, CString& languageFilePath);
	static int CreateLanguageFileFromUTF16LEEncodedFiles(CString& resFile1, CString& resFile2, CString& languageFilePath);

	static int ResortLanguageFile();
	static void LoadResInfoFromFile(CString& resFile, ResInfoArrayType& resArray);

	static ResInfoArrayType resArray1;
	static ResInfoArrayType resArray2;

	static ResInfoMapType g_LanguageMap;

	static ResInfoArrayType g_resInfoArray;
	static ResInfoMapType g_resInfoMap;

	static int PopulateResStringList() { return 0; }
	//
	static void LoadDialogItemsInfo(CDialog *dlg);
	static void LoadDialogItemsInfo(CWnd* wnd);
	static void LoadDialogItemsInfo(CWnd *wnd, HWND hwndParent, int maxcnt, int iter);
	//
	static void UpdateDialogItemsInfo(CDialog *dlg);
	static void UpdateDialogItemsInfo(CWnd* wnd);
	static void UpdateDialogItemsInfo(CWnd* wnd, HWND hwndParent, int maxcnt, int iter);
	//
	static void LoadListBoxItemsInfo(CListBox* menu, int index);
	static void UpdateListBoxItemsInfo(CListBox* menu, int index);
	//
	static void LoadComboBoxItemsInfo(CComboBox* menu, int index);
	static void UpdateComboBoxItemsInfo(CComboBox* menu, int index);
	//
	static void LoadMenuItemsInfo(CMenu* menu);
	static void LoadMenuItemsInfo(CMenu* menu, int index);
	static void UpdateMenuItemsInfo(CMenu* menu, int index);

	static void LoadToolBarItemsInfo(CToolBar* tbar);
	static void UpdateToolBarItemsInfo(CToolBar* tbar);

	static void LoadLanguageMap(CString & languageFolder);
	static void LoadLanguageMapFromFileF16LE(CString& languageFolder);
	
	static void DisableLanguageLoading();
	static void EnableLanguageLoading();

	static void EnableResInfoCollecting();

	static BOOL DetermineString(CString& str, CString& newString);

	static BOOL OnTtnNeedText(CWnd* parentWindow, NMHDR* pNMHDR, CString &toolTipText);
	static BOOL ActivateToolTips(CWnd* parentWnd, CToolTipCtrl &toolTipCtrl);
	
	static BOOL GetMenuItemString(CMenu* menu, UINT nIDItem, CString& rString, UINT nFlags);

	static void GetProcessFolderPath(CString& folderPath);
protected:
};



