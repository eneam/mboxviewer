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


#include <minwindef.h>
#include <afxstr.h>
#include <algorithm>
#include "dllist.h"
#include "IHashTable.h"


class CFile;

using namespace std;

enum ConfigNodeType
{
	STRING_NODE = 1,
	INT_NODE = 2,
	BINARY_NODE = 3
};


class ConfigNode
{
public:
	
	ConfigNode();
	~ConfigNode();

	BOOL m_isSectionNode;
	int m_type;
	CString m_name;
	CString m_value;
	ConfigNode* m_parent;

	dlink_node<ConfigNode> m_listLink;
	dlink_node<ConfigNode> m_hashMapLink;  // optional: use to speed access to ConfigNodes via hash table
	//

	using ConfigNodeList = dllist<ConfigNode, &ConfigNode::m_listLink>;
	ConfigNodeList *m_pConfigList;
	ConfigNodeList m_configList;

	void SetDflt();

friend class ConfigTree;
};

struct ConfigNodeHelper
{
	hashsum_t operator()(const CString* key) const
	{
		hashsum_t hashsum = StrHash((const char*)(LPCWSTR)*key, key->GetLength() * 2);
		return hashsum;
	}
	bool operator()(CString* key1, CString* key2) const
	{
		if (*key1 == *key2)
			return true;
		else
			return false;
	}
	bool operator()(CString* key1, ConfigNode* key2) const
	{
		if (*key1 == key2->m_value)
			return true;
		else
			return false;
	}
};

using ConfigNodeList = dllist<ConfigNode, &ConfigNode::m_listLink>;
using ConfifgNodeHashTable = IHashMap<CString, ConfigNode, ConfigNodeHelper, ConfigNodeHelper, &ConfigNode::m_hashMapLink>;

class ConfigTree
{
public:
	ConfigTree(CString& name);
	ConfigTree();
	~ConfigTree();

	CString m_name;

	ConfigNode m_rootNode;
	ConfifgNodeHashTable m_configNodeTable;

	void GetConfigNodePath(ConfigNode* node, CString& path);
	static void GetSection(CString& regSection, CString& section);
	ConfigNode* FindNode(CString &configSection);
	ConfigNode* FindNode(ConfigNode* node, CString *name);
	ConfigNode* InsertNode(ConfigNode* parent, ConfigNode* node);

	BOOL _DeleteProfileString(HKEY hKey, LPCWSTR section, LPCWSTR key);
	BOOL _WriteProfileInt(HKEY hKey, LPCWSTR section, LPCWSTR key, DWORD value);
	BOOL _WriteProfileString(HKEY hKey, LPCWSTR section, LPCWSTR key, CString& value);
	BOOL _WriteProfileBinary(HKEY hKey, LPCWSTR section, LPCWSTR key, const BYTE* lpData, DWORD cbData);
	//
	int _GetProfileInt(HKEY hKey, LPCWSTR section, LPCWSTR key);
	CString _GetProfileString(HKEY hKey, LPCWSTR section, LPCWSTR key);
	BOOL _GetProfileInt(HKEY hKey, LPCWSTR section, LPCWSTR key, DWORD& intval);
	BOOL _GetProfileInt(HKEY hKey, LPCWSTR section, LPCWSTR key, int& intval);
	BOOL _GetProfileString(HKEY hKey, LPCWSTR section, LPCWSTR key, CString& str);
	BOOL _GetProfileBinary(HKEY hKey, LPCWSTR section, LPCWSTR key, BYTE* lpData, DWORD& cbData);

	BOOL _DeleteValue(HKEY hKey, LPCWSTR section, LPCWSTR key);
	BOOL _DeleteKey(HKEY hKey, LPCWSTR section, LPCWSTR key, BOOL recursive);

	BOOL IsEmpty() { return ChildCount(&m_rootNode) == 0; }
	int ChildCount(ConfigNode *node) { return node->m_configList.count(); }
	int Count() { return ChildCount(&m_rootNode); }
	void DeleteAllNodes();
	void DeleteNode(ConfigNode* node);

	ConfigNode* GetRootNode() { return &m_rootNode; }

	void SortTree();
	void SortNode(ConfigNode *node);
	void TraceConfigNodeList(ConfigNodeList& configList);

	void DumpTree(CString &title);
	void DumpNode(ConfigNode *node, int level);

	void DumpTree();
	void DumpNode(ConfigNode* node, CString &out);

	int Dump2File();
	int Dump2File(CString& filepath);

	int LoadConfigFromFile();
	int LoadConfigFromRegistry();
	int LoadConfigFromFile(CString& filepath);
	void LoadLConfigFromFileUTF16LE(CString& configFileNamePath);

	int LoadConfigFromUMBoxViewerRegistry();

	static BOOL Hex2Binary(CString& hex, BYTE* lpData, DWORD& cbData);
	static BOOL Binary2Hex(BYTE* lpData, DWORD& cbData, CString& hex);
	static int HexChar2Int(wchar_t c);
};
