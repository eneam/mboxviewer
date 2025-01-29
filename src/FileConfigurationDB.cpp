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


#pragma once

#include "stdafx.h"
#include "FileConfigurationDB.h"
#include "TextUtilsEx.h"
#include "SimpleString.h"
#include "FileUtils.h"
#include "ResHelper.h"
#include "mboxview.h"


ConfigNode::ConfigNode()
{
	SetDflt();
}

void ConfigNode::SetDflt()
{
	m_parent = 0;
	m_isSectionNode = TRUE;
	m_type = 0;
	m_name = L"";
	m_value = L"";
	m_parent = 0;
	m_pConfigList = &m_configList;
};

ConfigNode::~ConfigNode()
{
}

ConfigTree::ConfigTree()
	: m_configNodeTable(1)
{
	int deb = 1;
}

ConfigTree::ConfigTree(CString& name)
	: m_configNodeTable(1)
{
	m_rootNode.m_name = name;
}

ConfigTree::~ConfigTree()
{
}

void ConfigTree::DumpTree(CString &title)
{
	// Reenable for testing
#if 0
	TRACE(L"%s %s", title, L"TREE DUMP\n");
	//SortTree();

	DumpNode(&m_rootNode, 0);
#endif
}

void ConfigTree::DumpNode(ConfigNode* node, int level)
{
	CString indent;
	int i;
	for (i = 0; i < (level); i++)
		indent.AppendChar(L' ');

	CString indent_sv = indent;

	ConfigNode* childNode;
	for (childNode = node->m_configList.first(); childNode != 0; childNode = node->m_configList.next(childNode))
	{
		if (!childNode->m_isSectionNode)
		{
			TRACE(L"%s\"%s\"=\"%s\"\n", indent, childNode->m_name, childNode->m_value);
		}

	}

	indent = indent_sv;
	for (childNode = node->m_configList.first(); childNode != 0; childNode = node->m_configList.next(childNode))
	{
		CString text;
		if (childNode->m_isSectionNode)
		{
			TRACE(L"%s[%s]\n", indent, childNode->m_name);
		}

		if (childNode->m_configList.count())
		{
			DumpNode(childNode, level + 2);
		}
	}
	return;
}

void ConfigTree::DumpTree()
{
#ifdef _DEBUG
#if 0
	CString out;
	out.Preallocate(10000);

	DumpNode(&m_rootNode, out);

	TRACE(L"%s\n", L"TREE DUMP CONFIG TREE");
	TRACE(L"%s\n", out);
#endif
#endif
}

void ConfigTree::DumpNode(ConfigNode* node, CString &out)
{
	ConfigNode* childNode;
	CString nodePath;

	// Dump first non Section nodes. Create separate list for section and param nodes??

	for (childNode = node->m_configList.first(); childNode != 0; childNode = node->m_configList.next(childNode))
	{
		if (!childNode->m_isSectionNode)
		{
			_ASSERTE(childNode->m_configList.count() == 0);
			out.AppendChar(L'"');
			out.Append(childNode->m_name);
			out.AppendChar(L'"');
			out.AppendChar(L'=');

			out.AppendChar(L'"');
			out.Append(childNode->m_value);
			out.AppendChar(L'"');
			out.AppendChar(L'\n');
		}
	}
	out.AppendChar(L'\n');

	///
	for (childNode = node->m_configList.first(); childNode != 0; childNode = node->m_configList.next(childNode))
	{
		if (childNode->m_isSectionNode)
		{
			GetConfigNodePath(childNode, nodePath);
			out.AppendChar(L'[');
			out.Append(nodePath);
			out.AppendChar(L']');
			out.AppendChar(L'\n');
		}

		if (childNode->m_configList.count())
		{
			DumpNode(childNode, out);
		}
	}
	return;
}


void ConfigTree::DeleteAllNodes()
{
	TRACE(L"DeleteAllNodes\n");
	DeleteNode(&m_rootNode);
}

// Delete node and all subnodes
void ConfigTree::DeleteNode(ConfigNode* node)
{
	TRACE(L"DeleteNode: %s\n", node->m_name);
	_ASSERTE(node);

	ConfigNode* childNode;
	ConfigNode* next_childNode;
	for (childNode = node->m_configList.first(); childNode != 0; )
	{
		next_childNode = node->m_configList.next(childNode);
		if (childNode->m_configList.count())
		{
			DeleteNode(childNode);
		}
		else
		{
			if (childNode->m_parent)
			{
				_ASSERTE(node == childNode->m_parent);
				node->m_configList.remove(childNode);
				TRACE(L"DeletingNode: %s !!!\n", childNode->m_name);

				delete childNode;
			}
			else
				int deb = 1;
		}
		childNode = next_childNode;
	}
	if (node->m_parent)
	{
		_ASSERTE(node->m_pConfigList->count() == 0);
		node->m_parent->m_configList.remove(node);
		TRACE(L"DeletingNode: %s #####\n", node->m_name);

		delete node;
	}
	else
		int deb = 1;
	return;
}

void ConfigTree::GetConfigNodePath(ConfigNode* node, CString& path)
{
	_ASSERTE(node != 0);
	ConfigNode* n = node;
	path.Empty();
	while (n != &m_rootNode)
	{
		path.Insert(0, n->m_name);
		if (n->m_parent != &m_rootNode)
			path.Insert(0, L"\\");

		n = n->m_parent;
	}
}

void ConfigTree::GetSection(CString& registrySection, CString& section)
{
	section.Empty();
	int pos = registrySection.Find(L"UMBoxViewer");
	//_ASSERTE(pos >= 0);
	if (pos >= 0)
	{
		section = registrySection.Mid(pos);
	}
}

ConfigNode* ConfigTree::FindNode(CString &configSection)
{
	CStringArray ar;
	CString delim = L"\\";

	TextUtilsEx::SplitStringW(configSection, delim, ar);

	ConfigNode *node = &m_rootNode;

	ConfigNode* childNode;
	CString* name;
	int i = 0;
	for (i = 0; i < ar.GetCount(); i++)
	{
		name = &ar[i];
		childNode = FindNode(node, name);
		if (childNode == 0)
			break;
		else
			node = childNode;
	}

	return childNode;
}

ConfigNode* ConfigTree::FindNode(ConfigNode* node, CString* name)
{
	ConfigNode* childNode;
	for (childNode = node->m_configList.first(); childNode != 0; childNode = node->m_configList.next(childNode))
	{
		if (childNode->m_name.Compare(*name) == 0)
		{
			break;
		}
	}
	return childNode;
}

ConfigNode* ConfigTree::InsertNode(ConfigNode* parent, ConfigNode* node)
{
	_ASSERTE(parent != 0);
	_ASSERTE(node != 0);

	node->m_parent = parent;
	parent->m_configList.insert_head(node);
	return node;
}

BOOL ConfigTree::_DeleteProfileString(HKEY hKey, LPCWSTR section, LPCWSTR key)
{
#ifdef _DEBUG
	CString label = L"_DeleteProfileString";
	//DumpTree(label);
#endif

	CString configSection;
	CString registrySection = section;
	ConfigTree::GetSection(registrySection, configSection);
	if (configSection.IsEmpty())  // not UMBoxView
		return TRUE;

	CString configPath = configSection + L"\\" + key;

	ConfigNode* node = ConfigTree::FindNode(configPath);
	if (node && node->m_parent)
	{
		node->m_parent->m_pConfigList->remove(node);
		delete node;
	}
	return TRUE;
}

BOOL ConfigTree::_WriteProfileInt(HKEY hKey, LPCWSTR section, LPCWSTR key, DWORD value)
{
	// should we enforce _ASSERTE(key != 0); 
	CString configSection;
	CString registrySection = section;
	ConfigTree::GetSection(registrySection, configSection);
	if (configSection.IsEmpty())  // not UMBoxView
		return FALSE;

#ifdef _DEBUG
	CString label = L"ENTER _WriteProfileString" + CString(L" ") + section + CString(L" ") + key;
	DumpTree(label);
#endif

	CStringArray ar;
	CString delim = L"\\";

	TextUtilsEx::SplitStringW(configSection, delim, ar);

	if (key)
		ar.SetAtGrow(ar.GetCount(), key);

	ConfigNode* node = &m_rootNode;

	ConfigNode* childNode;
	CString* name;
	int i = 0;

	int arCnt = (int)ar.GetCount();
	int keyIndx = arCnt - 1;
	for (i = 0; i < arCnt; i++)
	{
		name = &ar[i];
		childNode = FindNode(node, name);
		if (childNode == 0)
		{
			ConfigNode* newNode = new ConfigNode;
			newNode->m_name = ar[i];
			newNode->m_parent = node;
			if (key && (i == keyIndx))
			{
				newNode->m_isSectionNode = FALSE;
				newNode->m_value.Format(L"%d", value);
			}

			node->m_configList.insert_head(newNode);
			node = newNode;
		}
		else
		{
			node = childNode;

			if (key && (i == keyIndx))
			{
				_ASSERTE(node->m_isSectionNode == FALSE);
				node->m_isSectionNode = FALSE;
				node->m_value.Format(L"%d", value);
			}
		}
	}

#ifdef _DEBUG
	label = L"EXIT _WriteProfileString";
	DumpTree(label);
#endif

	return TRUE;
}

BOOL ConfigTree::_WriteProfileString(HKEY hKey, LPCWSTR section, LPCWSTR key, CString& value)
{
	// should we enforce _ASSERTE(key != 0); 
	CString configSection;
	CString registrySection = section;
	ConfigTree::GetSection(registrySection, configSection);
	if (configSection.IsEmpty())  // not UMBoxView
		return FALSE;

#ifdef _DEBUG
	CString label = L"ENTER _WriteProfileString" + CString(L" ") + section + CString(L" ") + key;
	DumpTree(label);
#endif

	CStringArray ar;
	CString delim = L"\\";

	TextUtilsEx::SplitStringW(configSection, delim, ar);

	if (key)
		ar.SetAtGrow(ar.GetCount(), key);

	ConfigNode* node = &m_rootNode;

	ConfigNode* childNode;
	CString* name;
	int i = 0;
	int arCnt = (int)ar.GetCount();
	int keyIndx = arCnt - 1;
	for (i = 0; i < arCnt; i++)
	{
		name = &ar[i];
		childNode = FindNode(node, name);
		if (childNode == 0)
		{
			ConfigNode* newNode = new ConfigNode;
			newNode->m_name = ar[i];
			newNode->m_parent = node;
			if (key && (i == keyIndx))
			{
				newNode->m_isSectionNode = FALSE;
				newNode->m_value = value;
			}

			node->m_configList.insert_head(newNode);
			node = newNode;
		}
		else
		{
			node = childNode;

			if (key && (i == keyIndx))
			{
				_ASSERTE(node->m_isSectionNode == FALSE);
				if (node->m_isSectionNode == FALSE)
					node->m_value = value;
			}
		}
	}

#ifdef _DEBUG
	label = L"EXIT _WriteProfileString";
	DumpTree(label);
#endif

	return TRUE;
}

BOOL ConfigTree::_WriteProfileBinary(HKEY hKey, LPCWSTR section, LPCWSTR key, const BYTE* lpData, DWORD cbData)
{
	// should we enforce _ASSERTE(key != 0); 
	CString configSection;
	CString registrySection = section;
	ConfigTree::GetSection(registrySection, configSection);
	if (configSection.IsEmpty())  // not UMBoxView
		return FALSE;

#ifdef _DEBUG
	CString label = L"ENTER _WriteProfileString" + CString(L" ") + section + CString(L" ") + key;
	DumpTree(label);
#endif

	CStringArray ar;
	CString delim = L"\\";

	TextUtilsEx::SplitStringW(configSection, delim, ar);

	if (key)
		ar.SetAtGrow(ar.GetCount(), key);

	ConfigNode* node = &m_rootNode;

	ConfigNode* childNode;
	CString* name;
	int i = 0;
	int arCnt = (int)ar.GetCount();
	int keyIndx = arCnt - 1;
	for (i = 0; i < arCnt; i++)
	{
		name = &ar[i];
		childNode = FindNode(node, name);
		if (childNode == 0)
		{
			ConfigNode* newNode = new ConfigNode;
			if (key && (i == keyIndx))
			{
				newNode->m_isSectionNode = FALSE;
				BOOL ret = ConfigTree::Binary2Hex((BYTE*)lpData, cbData, newNode->m_value);
			}
			newNode->m_name = ar[i];
			newNode->m_parent = node;

			node->m_configList.insert_head(newNode);
			node = newNode;
		}
		else
		{
			node = childNode;
			if (key && (i == keyIndx))
			{
				_ASSERTE(node->m_isSectionNode == FALSE);
				BOOL ret = ConfigTree::Binary2Hex((BYTE*)lpData, cbData, node->m_value);
			}
		}
	}

#ifdef _DEBUG
	label = L"EXIT _WriteProfileString";
	DumpTree(label);
#endif

	return TRUE;
}
//
int ConfigTree::_GetProfileInt(HKEY hKey, LPCWSTR section, LPCWSTR key)
{
#ifdef _DEBUG
	CString label = L"_GetProfileInt";
	//DumpTree(label);
#endif

	CString configSection;
	CString registrySection = section;
	ConfigTree::GetSection(registrySection, configSection);
	if (configSection.IsEmpty())  // not UMBoxView
		return 0;

	CString configPath = configSection + L"\\" + key;

	ConfigNode* node = ConfigTree::FindNode(configPath);
	if (node && !node->m_isSectionNode)
	{
		__int64 intval = _ttoi64(node->m_value);
		return (int)intval;
	}
	else
	{
		return 0;
	}
}

CString ConfigTree::ConfigTree::_GetProfileString(HKEY hKey, LPCWSTR section, LPCWSTR key)
{
#ifdef _DEBUG
	CString label = L"_GetProfileString";
	//DumpTree(label);
#endif

	CString strval;

	CString configSection;
	CString registrySection = section;
	ConfigTree::GetSection(registrySection, configSection);
	if (configSection.IsEmpty())  // not UMBoxView
		return strval;

	CString configPath = configSection + L"\\" + key;

	ConfigNode* node = ConfigTree::FindNode(configPath);
	if (node && !node->m_isSectionNode)
		return node->m_value;
	else
		return strval;
}

BOOL ConfigTree::_GetProfileInt(HKEY hKey, LPCWSTR section, LPCWSTR key, DWORD& intval)
{
#ifdef _DEBUG
	//DumpTree(CString(L"_GetProfileInt"));
#endif

	CString configSection;
	CString registrySection = section;
	ConfigTree::GetSection(registrySection, configSection);
	if (configSection.IsEmpty())  // not UMBoxView
		return 0;

		CString configPath = configSection + L"\\" + key;

		ConfigNode* node = ConfigTree::FindNode(configPath);
		if (node && !node->m_isSectionNode)
		{
			intval = (DWORD)_ttoi64(node->m_value);
			return TRUE;
		}
		else
		{
			return FALSE;
		}

		return TRUE;
}

BOOL ConfigTree::_GetProfileInt(HKEY hKey, LPCWSTR section, LPCWSTR key, int& intval)
{
#ifdef _DEBUG
	//DumpTree(CString(L"_GetProfileInt"));
#endif

	CString configSection;
	CString registrySection = section;
	ConfigTree::GetSection(registrySection, configSection);
	if (configSection.IsEmpty())  // not UMBoxView
		return 0;

	CString configPath = configSection + L"\\" + key;

	ConfigNode* node = ConfigTree::FindNode(configPath);
	if (node && !node->m_isSectionNode)
	{
		intval = (int)_ttoi64(node->m_value);
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

BOOL ConfigTree::_GetProfileString(HKEY hKey, LPCWSTR section, LPCWSTR key, CString& str)
{
#ifdef _DEBUG
	CString label = L"_GetProfileString";
	//DumpTree(label);
#endif

	CString strval;

	CString configSection;
	CString registrySection = section;
	ConfigTree::GetSection(registrySection, configSection);
	if (configSection.IsEmpty())  // not UMBoxView
		return FALSE;

	CString configPath = configSection + L"\\" + key;

	ConfigNode* node = ConfigTree::FindNode(configPath);
	if (node && !node->m_isSectionNode)
	{
		str = node->m_value;
		return TRUE;
	}
	else
		return FALSE;
}

BOOL ConfigTree::_GetProfileBinary(HKEY hKey, LPCWSTR section, LPCWSTR key, BYTE* lpData, DWORD& cbData)
{
	static BYTE* bdata = 0;
	static int bdatalen = 0;

#ifdef _DEBUG
	CString label = L"_GetProfileBinary";
	//DumpTree(label);
#endif

	CString configSection;
	CString registrySection = section;
	ConfigTree::GetSection(registrySection, configSection);
	if (configSection.IsEmpty())  // not UMBoxView
		return FALSE;

	CString configPath = configSection + L"\\" + key;

	ConfigNode* node = ConfigTree::FindNode(configPath);
	if (node && !node->m_isSectionNode)
	{
		BOOL ret = ConfigTree::Hex2Binary(node->m_value, lpData, cbData);
		return ret;
	}
	else
		return FALSE;
}

BOOL ConfigTree::_DeleteValue(HKEY hKey, LPCWSTR section, LPCWSTR key)
{
#ifdef _DEBUG
	CString label = L"_DeleteValue";
	//DumpTree(label);
#endif

	CString configSection;
	CString registrySection = section;
	ConfigTree::GetSection(registrySection, configSection);
	if (configSection.IsEmpty())  // not UMBoxView
		return TRUE;

	CString configPath = configSection + L"\\" + key;

	ConfigNode* node = ConfigTree::FindNode(configPath);
	if (node && node->m_parent)
	{
		node->m_parent->m_pConfigList->remove(node);
		delete node;
	}
	return TRUE;
}

BOOL ConfigTree::_DeleteKey(HKEY hKey, LPCWSTR section, LPCWSTR key, BOOL recursive)
{
#ifdef _DEBUG
	//DumpTree(CString(L"BEGIN: _DeleteKey"));
#endif

	CString configSection;
	CString registrySection = section;
	ConfigTree::GetSection(registrySection, configSection);
	if (configSection.IsEmpty())  // not UMBoxView
		return TRUE;

	CString configPath = configSection + L"\\" + key;

	ConfigNode* node = ConfigTree::FindNode(configPath);
	if (node)
	{
		DeleteNode(node);
	}
	return TRUE;
}

int ConfigTree::HexChar2Int(wchar_t c)
{
	if ((c >= '0') && (c <= '9'))
		return (c - '0');
	else if ((c >= 'a') && (c <= 'f'))
		return (c - 'a' + 10);
	else if ((c >= 'A') && (c <= 'F'))
		return (c - 'A' + 10);
	else
		return -1;
}

BOOL ConfigTree::Hex2Binary(CString& hex, BYTE* lpData, DWORD& cbData)
{
	int cnt = hex.GetLength();
	const wchar_t* p = (LPCWSTR)hex;

	int i;
	DWORD byteCnt = 0;
	BYTE b;
	int v1;
	int v2;

	for (i = 0; i < cnt; i=i+2)
	{
		if (byteCnt >= cbData)
			return FALSE;

		v1 = ConfigTree::HexChar2Int(*p);
		v2 = ConfigTree::HexChar2Int(*(p+1));
		if ((v1 < 0) || (v2 < 0))
			return FALSE;

		b = v1*16 + v2;

		lpData[byteCnt++] = b;

		p += 2;
	}

	cbData = byteCnt;

#if 0
	//  Test
	CString outhex;
	BOOL ret = ConfigTree::Binary2Hex(lpData, cbData, outhex);
	_ASSERTE(outhex.GetLength() == cbData);
	int retcomp = hex.Compare(outhex);
	_ASSERTE(retcomp == 0);
#endif

	return TRUE;
}

BOOL ConfigTree::Binary2Hex(BYTE* lpData, DWORD& cbData, CString& hex)
{
	CString tmp;
	DWORD  k;
	BYTE b;
	hex.Empty();
	hex.Preallocate(4 * cbData);
	for (k = 0; k < cbData; k++)
	{
		b = lpData[k];
		tmp.Format(L"%02x", b);
		hex += tmp;
	}
#if 0
	// Test
	DWORD datalen = 4 * cbData;
	BYTE *data = new BYTE[datalen];
	BOOL ret = ConfigTree::Hex2Binary(hex, data, datalen);
	_ASSERTE(datalen == cbData);
	int retcomp = memcmp(data, lpData, datalen);
	_ASSERTE(retcomp == 0);
	delete [] data;
#endif
	return TRUE;
}

int  ConfigTree::Dump2File()
{
	_ASSERTE(!CProfile::IsRegistryConfig());

#if 0
	CString configFileName = L"MBoxViewer.conf";
	CString configFilePath = CmboxviewApp::m_currentPath + L"\\" + configFileName;

	int  ret = ConfigTree::Dump2File(configFilePath);

#endif
	CString configFileName = L"MBoxViewer.config";
	CString configFilePath = CmboxviewApp::m_configFilePath;


	SortTree();
	int ret = ConfigTree::Dump2File(configFilePath);

	return ret;
}

int  ConfigTree::Dump2File(CString& filepath)
{
	int retval;
	CString errorText;

	HANDLE hConfigFile = FileUtils::FileOpen(filepath, errorText, TRUE);
	if (hConfigFile == INVALID_HANDLE_VALUE)
	{
		return -1;
	}

	BYTE BOMF16LE[2] = { 0xFF, 0xFE };
	

	DWORD nNumberOfBytesToWrite = 2;
	DWORD nNumberOfBytesWritten = 0;
	retval = FileUtils::Write2File(hConfigFile, BOMF16LE, nNumberOfBytesToWrite, &nNumberOfBytesWritten);

	CString out;
	out.Preallocate(10000);

	out = LR"(
#
# MBox Viewer supports Windows Registry based configuration and the file based configuration.
# By default, Windows Registry is used to store configuration data.
# During startup, the MBox Viewer will check whether the MBoxViewer.config file exists and is writeable in:
#
#   1) the Config subfolder under the MBox Viewer software installation folder  or
#   2) in the UMBoxViewer\Config subfolder under the  user specific folder created by Windows system 
#        example : C:\Users\UserName\AppData\Local\UMBoxViewer\Config
#
# The config file format is similar to the format of ".reg" registry file
# [UMBoxViewer\LastSelection]
# "parameter"="value"
#
# White spaces are not allowed in the front of each line and around the "=" character.
# All parameter values are encoded as strings and converted by MBox Viewer to numbers or other data types when needed.
#
# MBoxViewer.config file must be encoded as UTF16LE BOM file
#
# MBoxViewer.config.sample file is included in the software package under the Config folder.
# In order to enable MBox Viewer to use the file based configuration, 
# user needs to rename this file to MBoxViewer.config file or copy the sample file
# to C:\Users\UserName\AppData\Local\UMBoxViewer\Config folder and rename.
#

)";



	nNumberOfBytesToWrite = out.GetLength() * 2;
	nNumberOfBytesWritten = 0;
	retval = FileUtils::Write2File(hConfigFile, (BYTE*)(LPWSTR)(LPCWSTR)out, nNumberOfBytesToWrite, &nNumberOfBytesWritten);


	out.Empty();
	DumpNode(&m_rootNode, out);

	nNumberOfBytesToWrite = out.GetLength()*2;
	nNumberOfBytesWritten = 0;
	retval = FileUtils::Write2File(hConfigFile, (BYTE*)(LPWSTR)(LPCWSTR)out, nNumberOfBytesToWrite, &nNumberOfBytesWritten);

	FileUtils::FileClose(hConfigFile);

	return 1;
}

void ConfigTree::SortTree()
{
	ConfigNode* node = &m_rootNode;
	SortNode(node);
}

void ConfigTree::TraceConfigNodeList(ConfigNodeList & configList)
{
	ConfigNode* node;
	CString out;
	out.Append(L"\n");
	for (node = configList.first(); node != 0; node = configList.next(node))
	{
		out.Append(node->m_name);
		out.Append(L"\n");
	}
	out.Append(L"\n");
	TRACE(L"%s\n", out);
}

void ConfigTree::SortNode(ConfigNode *node)
{
	CString nodePath;
	// Dump first non Section nodes. Create seprate list ??

	ConfigNode* firstNode = 0;

	ConfigNode* childNode;
	ConfigNode*next_childNode;

	for (childNode = node->m_configList.first(); childNode != 0; )
	{
		next_childNode = node->m_configList.next(childNode);
		if (!childNode->m_isSectionNode)
		{
			_ASSERTE(childNode->m_configList.count() == 0);
			
			if (firstNode == 0)
				firstNode = childNode;
			else
			{
				ConfigNode* child;
				ConfigNode* child_found = 0;

				child = node->m_configList.prev(childNode);
				while (child)
				{
					int res = child->m_name.CompareNoCase(childNode->m_name);
					if (res > 0)
					{
						child_found = child;
					}

					if (child == firstNode)
						break;

					child = node->m_configList.prev(child);

				}
				if (child_found)
				{
					node->m_configList.remove(childNode);
					node->m_configList.insert_before(childNode, child_found);
					if (child_found == firstNode)
						firstNode = childNode;
					int deb = 1;
				}
			}
		}
		childNode = next_childNode;
	}


	///
	firstNode = 0;

	for (childNode = node->m_configList.first(); childNode != 0; )
	{
		next_childNode = node->m_configList.next(childNode);
		if (childNode->m_isSectionNode)
		{
			if (childNode->m_configList.count())
			{
				SortNode(childNode);
			}

			if (firstNode == 0)
				firstNode = childNode;
			else
			{
				ConfigNode* child;
				ConfigNode* child_found = 0;

				child = node->m_configList.prev(childNode);
				while (child)
				{
					int res = child->m_name.CompareNoCase(childNode->m_name);
					if (res > 0)
					{
						child_found = child;
					}

					if (child == firstNode)
						break;

					child = node->m_configList.prev(child);

				}
				if (child_found)
				{
					node->m_configList.remove(childNode);
					node->m_configList.insert_before(childNode, child_found);
					if (child_found == firstNode)
						firstNode = childNode;
					int deb = 1;
				}
			}
		}
		childNode = next_childNode;
	}
}

int ConfigTree::LoadConfigFromUMBoxViewerRegistry()
{
	return 1;
}


int ConfigTree::LoadConfigFromFile()
{
	_ASSERTE(!CmboxviewApp::m_configFilePath.IsEmpty());
	if (CmboxviewApp::m_configFilePath.IsEmpty())
	{
		return -1;
	}


	CString configFileName = L"MBoxViewer.config";
	CString configFileNamePath = CmboxviewApp::m_configFilePath;

#if 0
	CString processNamePath;

	CString currentPath = L"F:\\Documents\\GIT1.0.3.42\\mboxviewer\\x64\\Debug";

	CString configFileNamePath = currentPath + L"\\" + configFileName;
#endif

	int ret = ConfigTree::LoadConfigFromFile(configFileNamePath);

	// lastSelectionSection is referenced most
	// Move lastSelectionSection to front of the list
	// or implement hash access
	CString lastSelectionSection = L"UMBoxViewer\\LastSelection";

	ConfigNode* node = ConfigTree::FindNode(lastSelectionSection);
	if (node && node->m_isSectionNode)
	{
		if (node->m_parent)
		{
			node->m_parent->m_pConfigList->remove(node);
			node->m_parent->m_pConfigList->insert_head(node);
		}
	}

	return 1;
}

int ConfigTree::LoadConfigFromFile(CString& configFileNamePath)
{
	CString controlName;
	CString strLine;
	CString str;
	CString strW;

	ResHelper::TextEncoding bom = ResHelper::TextEncoding::NONE;
	if (bom == ResHelper::TextEncoding::NONE)
		bom = ResHelper::GetFileBOM(configFileNamePath);
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
	{
		LoadLConfigFromFileUTF16LE(configFileNamePath);

		ConfigTree* config = CProfile::GetConfigTree();
		config->DumpTree();
		return 1;
	}
	else
		return -1;
}

#if 1

void ConfigTree::LoadLConfigFromFileUTF16LE(CString& configFileNamePath)
{

	CString controlName;
	CString strLine;
	CString str;
	CString strW;


	UINT nOpenFlags = CFile::modeRead | CFile::typeText | CFile::shareDenyNone;

	ResHelper::TextEncoding  bom = ResHelper::GetFileBOM(configFileNamePath);
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
	if (!file.Open(configFileNamePath, nOpenFlags, &exList))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exList);
#if 0
		CString txt = L"Could not open list file \"" + configFileNamePath;
		txt += L"\" file.\n";
		txt += exErrorStr;
#endif

		CString fmt = L"Could not open list file \"%s\" file.\n%s";
		ResHelper::TranslateString(fmt);
		CString txt;
		txt.Format(fmt, configFileNamePath, exErrorStr);

		TRACE(L"ResHelper: %s\n", txt);
		//errorText = txt;

		CFileStatus rStatus;
		BOOL ret = file.GetStatus(rStatus);

		//HWND h = NULL; // we don't have any window yet
		HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();

		int answer = MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return;
	}
	// CStdioFile is not BOM aware !!!! 
	// wchar_t BOM[1];
	// file.Read(&BOM[0], sizeof(wchar_t));  // or see below while look

	while (file.ReadString(strLine))
	{
		// TODO: Implement RemoveBOM()
		const unsigned char* p = (unsigned char*)((LPCWSTR)strLine);
		if ((*p == 0xFF) && (*(p + 1) == 0xFE))  // UTF16LE
		{
			CString str = strLine.Right(strLine.GetLength() - 1);
			strLine = str;
		}

		if (strLine[0] == L'[')
		{
			int slen = strLine.GetLength();
			int pos = strLine.Find(L']');
			if (pos < 2)
				return;


			LPCWSTR key = 0;
			CString value;
			CString section = strLine.Mid(1, pos - 1);

			// Create section in a file if not already present
			CProfile::_WriteProfileString(HKEY_CURRENT_USER, section, key, value);

			// Process all parameters in the section
			while (file.ReadString(strLine))
			{
				if (TextUtilsEx::isWhiteLine(strLine))
					break;

				int pos = strLine.Find(L'=');
				if (pos < 3)
					return;

				CString param = strLine.Mid(1, pos - 2);
				int slen = strLine.GetLength();
				if (strLine.GetAt(slen-1) == L'\"')
				{
					value = strLine.Mid(pos + 2, slen - pos - 3);
				}
				else
				{
					value = strLine.Mid(pos + 2, slen - pos - 2);
					pos = value.ReverseFind(L'\"');
					if (pos >= 0)
					{
						value = value.Left(pos);
						int deb = 1;
					}
					else
						int deb = 1; // TODO: ignore error ???
				}

				CProfile::_WriteProfileString(HKEY_CURRENT_USER, section, param, value);

				int deb = 1;
			}
			int deb = 1;
		}
	}
	int deb = 1;
}
#endif
