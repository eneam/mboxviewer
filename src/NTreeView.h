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
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
//  Library General Public License for more details.
//
//  You should have received a copy of the GNU Library General Public
//  License along with this program; if not, write to the
//  Free Software Foundation, Inc., 51 Franklin St, Fifth Floor,
//  Boston, MA  02110 - 1301, USA.
//
//////////////////////////////////////////////////////////////////
//

#if !defined(AFX_NTREEVIEW_H__28F2C044_02DA_4048_8F8F_3D17ADEE1421__INCLUDED_)
#define AFX_NTREEVIEW_H__28F2C044_02DA_4048_8F8F_3D17ADEE1421__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// NTreeView.h : header file
//


#include <algorithm>
#include <unordered_map>
#include <vector>
#include <array>
#include <deque>
#include "WheelTreeCtrl.h"
#include "dllist.h"
#include "IHashTable.h"
//#include <afxtempl.h>

/////////////////////////////////////////////////////////////////////////////
// NTreeView window

#define ASSERT_TREE_CTRL 1
#define ASSERT_FOLDERS_IN_REG_AND_MEM 2
#define ASSERT_FOLDERS_IN_MEM_AND_TREE 4


class MySaveFileDialog : public CFileDialog
{

public:
	MySaveFileDialog(
		BOOL bOpenFileDialog, // TRUE for FileOpen, FALSE for FileSaveAs
		LPCWSTR lpszDefExt = NULL,
		LPCWSTR lpszFileName = NULL,
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
		LPCWSTR lpszFilter = NULL,
		CWnd* pParentWnd = NULL,
		DWORD dwSize = 0,
		BOOL bVistaStyle = TRUE
	);

	virtual BOOL OnFileNameOK();
};

// MBoxFolderTree is the helper class to support "File->Select root folder" feature
// Easy to use incorrectly so be  careful
class MBoxFolderNode
{
public:
	MBoxFolderNode() { m_parent = 0; m_valid = TRUE; }
	CString m_folderName;
	MBoxFolderNode *m_parent;
	list<MBoxFolderNode> m_nodeList;
	BOOL m_valid;
};

class MBoxFolderTree
{
public:
	MBoxFolderTree(CString &name) { m_name = name; m_root = 0;  }
	CString m_name;
	MBoxFolderNode m_rootNode;
	MBoxFolderNode *m_root;
	list<MBoxFolderNode> m_rootList;


	BOOL IsEmpty() { return m_root == 0;}
	void EraseRoot() { m_root = 0; }  // assume m_root->m_nodeList.Empty() == true
	MBoxFolderNode *GetRootNode() { return m_root; }

	BOOL PopulateFolderTree(CString &rootFolder, MBoxFolderTree &tree, MBoxFolderNode *rnode, CString &errorText, int maxDepth);

	MBoxFolderNode *CreateNode(MBoxFolderNode *node);

	void Print(CString &filepath);
	void PrintNode(CFile *fp, MBoxFolderNode *node);

	int Count();
	int NodeCount(MBoxFolderNode *node);

	void PruneNonMBoxFolderNode(MBoxFolderNode *node);
	void PruneNonMBoxFolders();
	static void GetRelativeFolderPath(MBoxFolderNode *rnode, CString &folderPath);
};

typedef  CArray<CString> CSArray;

// Store for Folder with MBox file present
// Not used for sub-folders of the root folder
class CRegArray
{
public:
	CRegArray(CString &section);
	CRegArray();
	~CRegArray();

	int Find(CString &str);
	int Add(CString &str);
	int Delete(CString &str);
	int Delete(int index, CString &str);
	BOOL GetAt(int index, CString &str);
	int GetCount();
	BOOL SaveToRegistry();
	BOOL LoadFromRegistry();
	BOOL LoadFromRegistry(CSArray& ar);
	BOOL LoadFromRegistry_registry(CSArray &ar);
	BOOL LoadFromConfigFile(CSArray& ar);
	BOOL CreateKey(CString &section, HKEY &hKey);
	LSTATUS DeleteSectionKey();
	void Optimize();

	void Dump();

	CString m_ProcSoftwarePath;
	CString m_section;
	int m_nMaxSize;
	CSArray m_array;
};

class ArchiveFileInfo
{
public:
	ArchiveFileInfo() { fSize = 0; bShow = 1; bNeedsValidation = TRUE; };
	_int64 fSize;
	_int64 bShow;
	BOOL bNeedsValidation;
};

typedef CMap<CString, LPCWSTR, ArchiveFileInfo, ArchiveFileInfo> FileSizeMap;

// Folder name plus table of mail archive names and sizes
// Loaded from .mboxview plus new files discovered
class ArchiveFileInfoMap
{
public:
	ArchiveFileInfoMap(CString &folderPath) : m_fileSizes(31)
	{ m_folderPath = folderPath; };
	~ArchiveFileInfoMap() { m_fileSizes.RemoveAll();}
	CString m_folderPath;
	FileSizeMap m_fileSizes;
};

// table of tables. Table of all folders info, i.e table of FileSizeMap info
// Folder name is the key. The same as m_folderPath in ArchiveFileInfoMap
typedef unordered_map<std::wstring, ArchiveFileInfoMap*> GlobalFileSizeMap;

class MySimpleDeque;

// Used during parsing. It represent label and the list of associated mails
// i.e pointer to MboxMail loaded into MboxMail::s_mails array
// Used later to create Windows folders with .mboxlist file for each label
// Label database is file based and that creates complexity due to restriction on the files names
// There are a number of invalid or reserved characters that can't appear in the file name
class GmailLabel
{
public:
	GmailLabel(CStringA &label, DWORD codepage, CStringA &mappedToLabel, CStringA &mappedLabelPath);
	~GmailLabel();

	dlink_node<GmailLabel> m_hashMapLink;
	dlink_node<GmailLabel> m_hashMapLinkToMappedLabel;
	//
	CStringA m_mappedLabelPath;  // mapped all sublabels
	CStringA m_mappedToLabel;   // mapped label
	//
	MySimpleDeque *m_ptrList;  // list of MboxMail* items for a given label
	//
	CStringA& GetLabelA();  // return { m_label; }
	DWORD GetCodePage() { return m_codePage; }
	CString GetLabelW();   // m_label+m_codePage encoded as wide CString
protected:
	CStringA m_label;  // m_label+m_codePage (same as in gmail mail file) shown in TreeCtrl
	UINT m_codePage;
};

struct GmailLabelHelper 
{
	hashsum_t operator()(const CStringA *key) const
	{
		hashsum_t hashsum = StrHash((const char*)(LPCSTR)*key, key->GetLength());
		return hashsum;
	}
	bool operator()(CStringA *key1, CStringA *key2) const
	{
		if (*key1 == *key2)
			return true;
		else
			return false;
	}
	bool operator()(CStringA *key1, GmailLabel *key2) const
	{
		if (*key1 == key2->GetLabelA())
			return true;
		else
			return false;
	}
};

struct GmailMappedLabelHelper
{
	hashsum_t operator()(const CStringA *key) const
	{
		hashsum_t hashsum = StrHash((const char*)(LPCSTR)*key, key->GetLength());
		return hashsum;
	}
	bool operator()(CStringA *key1, CStringA *key2) const
	{
		if (*key1 == *key2)
			return true;
		else
			return false;
	}
	bool operator()(CStringA *key1, GmailLabel *key2) const
	{
		if (*key1 == key2->m_mappedToLabel)
			return true;
		else
			return false;
	}
};

#define LSIZE 512

typedef std::array<MboxMail*, LSIZE> LabelArray;


// TODO: home grown deque. Can we use something standard??
class MySimpleDeque
{
public:
	MySimpleDeque();
	~MySimpleDeque();
//
	std::vector<LabelArray*> m_arList;
	int Count();
	void Add(MboxMail* l);
	MboxMail* Get(int position);
	void Clear();
	BOOL Assert();
	LabelArray *m_ar;
	int m_arcnt;
};


// Table of GmailLabels. Use key/label name to find
using GmailLableMapType = IHashMap<CStringA, GmailLabel, GmailLabelHelper, GmailLabelHelper, &GmailLabel::m_hashMapLink>;
using GmailMappedLabelMapType = IHashMap<CStringA, GmailLabel, GmailMappedLabelHelper, GmailMappedLabelHelper, &GmailLabel::m_hashMapLinkToMappedLabel>;


// TODO:  Rename LabelInfo to TreeNodeInfo
// Used during runtime
// LabelInfo is attached to each label and mbox file in the treectrl
class LabelInfo
{
public:
	enum {
		MailFolder = 1,
		MailSubFolder = 2,
		MailFile = 4,
		MailLabel = 8
	};

	LabelInfo() {};
	// User: Watch order of CString params
	LabelInfo(int nId, CString &mailFilePath, CString &label, CString &listFilePath);  // m_nodeType = MailLabel
	LabelInfo(int nId, CString &mailFolderPath, CString &mailDataFolderPath, int nodeType); // m_nodeType = MailSubFolder
	LabelInfo(int nId, CString &mailFilePath);  // m_nodeType = MailFile;
	LabelInfo(int nodeType, int nId, CString &mailFilePath, CString &mailFolderPath);  // m_nodeType = MailFile;
	LabelInfo(int nId, CString &mailFolderPath, CString &mailDataFolderPath);  // m_nodeType = MailFolder
	~LabelInfo();

	dlink_node<LabelInfo> m_hashMapLink;
	int m_nId;
	int m_nodeType;
	CString m_label;
	CString m_listFilePath;
	CString m_filePath;
	CString m_mailFolderPath;
	CString m_mailDataFolderPath;
};

struct LabelInfoHelper
{
	hashsum_t operator()(const int *key) const
	{
		hashsum_t hashsum = *key;
		return hashsum;
	}
	bool operator()(int *key1, LabelInfo *key2) const
	{
		if (*key1 == key2->m_nId)
			return true;
		else
			return false;
	}
};


// Hash Table that stores all LabelInfo objects that can be found by int as the key
// Key is set/stored in each item in TreeCtrl
using GlobalLableInfoMapType = IHashMap<int, LabelInfo, LabelInfoHelper, LabelInfoHelper, &LabelInfo::m_hashMapLink>;


// Used during runtime
class GlobalLabelInfo
{
public:
	GlobalLabelInfo();
	~GlobalLabelInfo();

	GlobalLableInfoMapType *m_labelInfoHT;
	LabelInfo* Find(int key);
	int GetNextId();
	int Add(LabelInfo *linfo, int key);  // add and return key
	void Clear();
	void Print();
	int m_nId;
};

///

class FolderInfo
{
public:
	enum {
		MailFolderNode = 1,
		MailRootFolderNode = 2,
		MailSubFolderNode = 4
	};

	FolderInfo(CString &folderPath, int folderType) {
		m_folderPath = folderPath; 
		m_folderType = folderType;
		m_RootParent = 0;

	}
	~FolderInfo() {}

	CString FolderTypeToStr();

	dlink_node<FolderInfo> m_hashMapLink;
	dlink_node<FolderInfo> m_listLink;

	DLLIST(FolderInfo, m_listLink) m_folderList;
	CString m_folderPath;
	int m_folderType;
	FolderInfo *m_RootParent;
	HTREEITEM m_hFolder;  // TODO: enhance to avoid searching entire CtreeCtrl/m_tree
};

struct FolderInfoHelper
{
	hashsum_t operator()(const CString *key) const
	{
		hashsum_t hashsum = StrHash((const char*)(LPCWSTR)*key, key->GetLength()*2);
		return hashsum;
	}
	bool operator()(CString *key1, CString *key2) const
	{
		if (*key1 == *key2)
			return true;
		else
			return false;
	}
	bool operator()(CString *key1, FolderInfo *key2) const
	{
		if (*key1 == key2->m_folderPath)
			return true;
		else
			return false;
	}
};

using AllFoldersInfoHashTable = IHashMap<CString, FolderInfo, FolderInfoHelper, FolderInfoHelper, &FolderInfo::m_hashMapLink>;

class GlobalFolderInfoDB
{
public:
	GlobalFolderInfoDB();
	~GlobalFolderInfoDB();

	void Print();
	FolderInfo* Find(CString * key);
	FolderInfo* Find(CString * key, hashsum_t hashsum);
	void Add(CString * key, FolderInfo *info); 
	void Add(hashsum_t hashsum, FolderInfo *info);
	FolderInfo* Remove(CString * key);
	FolderInfo* Remove(CString * key, hashsum_t hashsum);
	//void Add(FolderInfo *info, hashsum_t hashsum);

	FolderInfo* FindRoot(CString * key);
	void AddRoot(CString * key, FolderInfo *info);

	void RemoveRootFolder(CString &folderPath);

	static hashsum_t GetHashsum(CString * key);

	void Clear();
	void DeleteAll();  // delete content and allocated hash table

	AllFoldersInfoHashTable *m_allFoldersTable;
	AllFoldersInfoHashTable *m_allRootFoldersTable;
	//
	FolderInfo *m_rootInfo;
	CString m_problemFolderPath;
	int m_problemFolderType;
	CString m_problemFolderRootPath;
	BOOL m_deleteConflictingFolders;
};

// Yet Another HashTable, shame on me
// Used  by ShoeGmailLabels only
class TreeCtrlInfoDB;

class TreeCtrlInfo
{
public:
	enum {
		MailFolder = 1,
		MailSubFolder = 2,
		MailFile = 4,
		MailLabel = 8
	};

	TreeCtrlInfo(CString &filePath, HTREEITEM hItem, int fileType = 0)
	{
		m_filePath = filePath;
		m_fileType = fileType;
		m_hItem = hItem;
	}
	~TreeCtrlInfo() {}

	CString TreeCtrlTypeToStr();

	dlink_node<TreeCtrlInfo> m_hashMapLink;

	// TODO: file path can be up 255 chars long, so this is not the most efficient solution. FIX later
	// Enhance to have two hash tables. Hash first by folder path and then hash by file name
	CString m_filePath;
	int m_fileType;
	HTREEITEM m_hItem;
};

struct TreeCtrlInfoHelper
{
	hashsum_t operator()(const CString *key) const
	{
		hashsum_t hashsum = StrHash((const char*)(LPCWSTR)*key, key->GetLength()*2);
		return hashsum;
	}
	bool operator()(CString *key1, CString *key2) const
	{
		if (*key1 == *key2)
			return true;
		else
			return false;
	}
	bool operator()(CString *key1, TreeCtrlInfo *key2) const
	{
		if (*key1 == key2->m_filePath)
			return true;
		else
			return false;
	}
};

using TreeCtrlInfoHashTable = IHashMap<CString, TreeCtrlInfo, TreeCtrlInfoHelper, TreeCtrlInfoHelper, &TreeCtrlInfo::m_hashMapLink>;

class TreeCtrlInfoDB
{
public:
	TreeCtrlInfoDB();
	~TreeCtrlInfoDB();

	void Print();
	TreeCtrlInfo* Find(CString * key);
	TreeCtrlInfo* Find(CString * key, hashsum_t hashsum);
	void Add(CString * key, TreeCtrlInfo *info);
	void Add(hashsum_t hashsum, TreeCtrlInfo *info);
	TreeCtrlInfo* Remove(CString * key);
	TreeCtrlInfo* Remove(CString * key, hashsum_t hashsum);


	static hashsum_t GetHashsum(CString * key);

	void Clear();
	void DeleteAll();  // delete content and allocated hash table

	TreeCtrlInfoHashTable *m_treeCtrlTable;
	//
};

class MailSelectionInfo
{
public:
	MailSelectionInfo()
	{
		m_lastMailIndex = 0;
		m_lastWhichSort = 1;

	};
	~MailSelectionInfo() {};

	CString m_lastPath;
	CString m_lastMailFilePath;
	CString m_lastLabelFilePath;
	int m_lastMailIndex;
	int m_lastWhichSort;
};

class NTreeView : public CWnd
{
	CFont		m_font;
	CImageList	m_il;
// Construction
public:
	NTreeView();
	virtual ~NTreeView();

	DECLARE_DYNCREATE(NTreeView)

	void ResetFont();

	BOOL ExpandItem(HTREEITEM hItem, UINT nCode);

	CWheelTreeCtrl	m_tree;
	// TODO: Consider consolidating m_globalFolderInfoDB and m_labelInfoStore in one table
	GlobalFolderInfoDB m_globalFolderInfoDB;
	TreeCtrlInfoDB m_treeCtrlInfoDB;
	//
	// Transient during parsing gmail files to create list of labels.
	GmailLableMapType *m_labelHT;
	GmailMappedLabelMapType *m_mappedToLabelHT;
	//
	// Live for entire duration of MBox Viewer
	// There is one LabelInfo for every node in TreeCtrl 
	GlobalLabelInfo m_labelInfoStore;
	BOOL m_labelView;
	int m_labelSeqNumb;

	//  TODO: consider to create class
	FileSizeMap	*fileSizes; // Folder name plus table of mail archive names and sizes
	ArchiveFileInfoMap *m_fileSizesMap;
	FileSizeMap	m_fileSizes; // Folder name plus table of mail archive names and sizes 
	GlobalFileSizeMap m_gFileSizes;  // table of all folders info, i.e table of FileSizeMap
	BOOL m_bIsDataDirty;  // if true, folder data/file sizes needs to be saved to .mboxview
	//
	BOOL m_bSelectMailFileDone;
	BOOL m_bSelectMailFilePostMsgDone;
	BOOL m_bGeneralHintPostMsgDone;
	int m_timerTickCnt;
	// timer
	UINT_PTR m_nIDEvent;
	UINT m_nElapse;
	CRegArray m_folderArray;
	int treeColWidth;

	BOOL  m_bInFillControl;
	int m_frameCx;
	int m_frameCy;

	DWORD m_treeViewAssertFlags;

	CString m_rootMboxFilePath;
	CFile m_rootMboxCfile;

	static int m_filesToValidateAsMboxType;

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(NTreeView)
	protected:
	virtual void PostNcDestroy();
	//}}AFX_VIRTUAL

// Implementation
public:
	void ClearLabelHT();
	// File Size Map
	void ClearGlobalFileSizeMap();
	BOOL RemoveFileSizeMap(CString path);
	ArchiveFileInfoMap *SetupFileSizeMap(CString &path);

	int RemoveOrphanIndexFilesFromFolderAndFileSizeMap(CString& path, FileSizeMap& fileSizes);
	HTREEITEM InsertAllDiscoveredMailFilesToTreeCtrl(HTREEITEM hParent, CString& path, FileSizeMap& fileSizes, BOOL unconditionalFolderInsert);
	HTREEITEM LoadFileSizes(HTREEITEM hParent, CString &path, FileSizeMap &fileSizes, BOOL unconditionalFolderInsert);
	int OpenHiddenFiles(HTREEITEM hItem, FileSizeMap &fileSizes, BOOL isSubFolder = FALSE);
	void RemoveFileFromTreeView(HTREEITEM hItem, FileSizeMap &fileSizes);
	void DetermineFolderPath(HTREEITEM hItem, CString &folderPath);
	//
	HTREEITEM DetermineRootItem(HTREEITEM hItem);
	HTREEITEM DetermineRootFolder(HTREEITEM hItem);
	HTREEITEM DetermineRootMboxFile(HTREEITEM hItem);
	void Traverse( HTREEITEM hItem, CFile &fp, FileSizeMap &fileSizes);
	void SaveData(HTREEITEM hItem);  // save data to per folder .mboxview file
	void ClearData();
	HTREEITEM FillCtrl(HTREEITEM hParent, BOOL selectFolder = FALSE, BOOL unconditionalFolderInsert = FALSE, BOOL expand = TRUE);
	HTREEITEM FillCtrl_Internal(HTREEITEM hParent, BOOL unconditionalFolderInsert = FALSE, BOOL expand = TRUE);
	void ExpandOrCollapseTree(BOOL expand);
	void LoadFolders();

	static BOOL ImportLegacyRegistryData();
	static BOOL ImportLegacyMboxviewRegistryData();
	static BOOL ImportLegacyUMBoxViewerRegistryData();
	static BOOL CopyKeyAndSubkey(const wchar_t* keyPath, wchar_t* keyName);

	void DeleteMBoxAllWorkFolders(CString &mboxFileName);
	HTREEITEM HasFolder(CString &path);
	HTREEITEM IsFolderOpen(CString &path);
	HTREEITEM IsFolderOpen(CString &rootFolder, CString &folderPath);
	HTREEITEM HasFolder(HTREEITEM hItem, CString &path);
	//
	HTREEITEM FindFolder(CString &path);
	HTREEITEM FindFolder(HTREEITEM hRoot, CString &path);
	BOOL FindFolder(HTREEITEM hRoot, CString &path, HTREEITEM &hFolder, BOOL recursive);
	BOOL FindFolder_internal(HTREEITEM hRoot, CString& path, HTREEITEM& hFolder, BOOL recursive);
	BOOL FolderHasChilder(CString& path);
	//
	//
	BOOL PrintFolderNames(HTREEITEM hRoot, BOOL recursive);
	BOOL PrintFolderNames_internal(HTREEITEM hRoot, BOOL recursive);
	//
	HTREEITEM FindMailFile(CString &path);
	//
	HTREEITEM FindLabelFile(CString &path, CString &lableFilePath);
	//
	void SelectMailFile(CString *fileName = 0); // based on -MBOX= command line argument
	void OpenLastSelection(MailSelectionInfo *mailSelection = 0);
	HTREEITEM FindLabelItemByLabelPath(HTREEITEM hRoot, CString &path);
	int FindSelectedItemIndex(MailArray &s_mails, int referenceIndex, BOOL force);
	void InsertMailFile(CString &mailFile);
	void ForceParseMailFile(HTREEITEM hItem);
	void UpdateFileSizesTable(CString &path, _int64 fSize, FileSizeMap &fileSizes);

	HTREEITEM SelectFolderIfAlreadyOpen(CString *path, int folderType, BOOL expand);
	HTREEITEM SelectTreeItem(CString *path);
	BOOL SelectTreeItem(HTREEITEM hItem);
	void SortChildren(HTREEITEM hItem, BOOL recursive = FALSE);
	int GetChildrenCount(HTREEITEM hItem, BOOL recursive);
	HTREEITEM FindItem(HTREEITEM hItem, CString &mailFileName, BOOL recursive = FALSE);
	HTREEITEM FindItemByItemPath(HTREEITEM hItem, CString &mailFileName, BOOL recursive = FALSE);
	BOOL DeleteItem(HTREEITEM hItem);
	BOOL DeleteItemChildren(HTREEITEM hItem);
	BOOL DeleteFolder(HTREEITEM hItem, BOOL setNewPath = TRUE);
	BOOL DeleteFolderIfEmpty(CString &path);
	void StartTimer();
	void PostMsgCmdParamFileName(CString *fileName = 0);
	BOOL RefreshFolder(HTREEITEM hItem);
	void PrintTreeCtrl(HTREEITEM hItem, BOOL recursive, int depth);

	int CreateLabelsForSingleMailFile(HTREEITEM hItem);
	int DeleteLabelsForSingleMailFile(HTREEITEM hItem);
	int RefreshLabelsForSingleMailFile(HTREEITEM hItem);
	int UpdateLabelMailListFile(HTREEITEM hItem);

	int  IsGmailLabelFile(CString& mboxFile);
	BOOL RecreateGmailLabels(HTREEITEM hItem);
	int CreateGmailLabelFiles(HTREEITEM hItem);
	BOOL DisplayGmailLabels(HTREEITEM hItem);
	int ShowGmailLabels(HTREEITEM hItem, CString &listFilePath, CString &dataFilePath);
	int ShowGmailLabels_internal(HTREEITEM hItem, CString &listFilePath, CString &dataFilePath);
	HTREEITEM HasLabel(HTREEITEM hItem, CString &label);
	BOOL LoadLabels();
	BOOL LoadLabels(HTREEITEM hItem);
	void MoveLabelItem(CString &mailFilePath, CString &label);
	void MoveLabelItem(HTREEITEM hRoot, CString &mailFilePath, CString &label);
	HTREEITEM HasMailFile(CString &mailFilePath);

	// Folder related
	int CreateEmptyFolder(HTREEITEM hItem);
	int GetFolderPath(HTREEITEM hItem, CString &mboxName, CString &parentPath);
	int GetFolderPath(HTREEITEM hItem, CString &folderPath);

	int CreateFlatFolderList(HTREEITEM hItem, CArray<CString> &folderList);
	int CreateFlatFolderList(CString &mboxFileName, CArray<CString> &folderList);

	static void FindAllDirs(LPCWSTR pstr);


	static BOOL ImboxviewFileFilter(CString & fName);
	static int ImboxviewFile(CString & fName);
	static int IsValidMailFile(char *data, int datalen);

	void DoOpen(CString &path, BOOL expand);
	void DoOpenFilePreview(CString &path);
	void DoOpenFolder(CString &path, BOOL selectFolder, BOOL unconditionalFolderInsert, BOOL expand);
	HTREEITEM OpenFolder(HTREEITEM hParent, CString &path, BOOL selectFolder, BOOL unconditionalFolderInsert, BOOL addToRecentFileList, BOOL expand);
	void OpenFolderAndSubfolders(HTREEITEM hParent, CString &path);
	//
	HTREEITEM OpenRootSubFolder(HTREEITEM hFolder, CString &path, BOOL addToRecentFileList = FALSE, BOOL expand = TRUE);
	void OpenRootFolderAndSubfolders(HTREEITEM hParent, CString &path, BOOL selectFolder = FALSE);
	void DoOpenRootFolderAndSubfolders(HTREEITEM hParent, CString &path, BOOL selectFolder = FALSE);
	//
	void PruneFolderTree(MBoxFolderTree &tree);
	void PruneNonMBoxFolders(MBoxFolderNode *node);
	//
	void OpenRootFolderAndSubfolders_LabelView(CString &path, BOOL selectFolder = FALSE);
	void DoOpenRootFolderAndSubfolders_LabelView(CString &path, BOOL selectFolder = FALSE);
	int MergeTreeNode(MBoxFolderNode *node);
	int MergeTreeFolders(MBoxFolderTree &tree, CString &errorText);
	int MergeRootSubFolder_NoLabels(CString &relativeFolderPath, CString &path, BOOL addToRecentFileList = FALSE, BOOL expand = TRUE);
	int MergeRootSubFolder(CString &relativeFolderPath, CString &path, BOOL addToRecentFileList = FALSE, BOOL expand = TRUE);
	int MergeMails(CStringA &labels);
	int MergeMailsRemoveDuplicates();
	int ArchiveMailsRemoveDuplicates(CFile &fp, CString &filePath);
	BOOL SaveMergedFileDialog(CString &fileName, CString &fileNameFilter, CString &dfltExtention, CString &inFolderPath, CString &outFolderPath, CString &title);

	BOOL FolderHasAtLeastOneMboxFile(CString &mboxFilePath);

	HTREEITEM OpenTreeNode(HTREEITEM hParent, MBoxFolderNode *node);
	void OpenTreeFolders(HTREEITEM hParent, MBoxFolderTree &tree);
	//
	BOOL ValidateFolderTreeNode(MBoxFolderNode *node, GlobalFolderInfoDB &globalFolderInfoDB);
	BOOL ValidateFolderTree(MBoxFolderTree &tree, GlobalFolderInfoDB &globalFolderInfoDB);

	BOOL SetFolderAsRoorFolder(CString &folderPath);
	BOOL RemoveFolderAsRootFolder(CString &folderPath);
	BOOL IsFolderARootFolder(CString &folderPath);
	BOOL IsFolderOpenAsRootFolder(CString &folderPath);

	BOOL CanFolderBeOpen(CString &folderPath);
	BOOL CanRootFolderBeOpen(CString &folderPath);

	BOOL AssertTreeCtrl(HTREEITEM hItem = 0);
	BOOL AssertTreeCtrl_Internal(HTREEITEM hItem = 0);
	BOOL AssertNTreeView();
	BOOL AssertNTreeView_Internal();

	//
	BOOL DeleteOldDataFolder(CString &currentRootDataFolderPath, CString &newRootDataFolderPath, HTREEITEM hItem = 0);

	HTREEITEM InsertTreeItem(LPCWSTR lpszItem, int nImage, int nSelectedImage, HTREEITEM hParent = TVI_ROOT, HTREEITEM hInsertAfter = TVI_LAST);
	HTREEITEM InsertTreeItem(LPCWSTR lpszItem, HTREEITEM hParent = TVI_ROOT, HTREEITEM hInsertAfter = TVI_LAST);

	int MergeMailArchiveFiles(HTREEITEM hFolder);

	int RefreshSubFolder(HTREEITEM hItem);
	int RefreshMboxFilesList(CString &folderPath, HTREEITEM hFolder);
	int PruneFolder(HTREEITEM hItem);  // remove folders/items if no mbox file found

	int RemoveMboxAssociatedWorkData(HTREEITEM hItem);  /// hItem for mbox file
	int RemoveMboxAssociatedWorkData(HTREEITEM hFolder, CString &filePath);

	HTREEITEM SelectFolderItem(CString &mboxFilePath);

	void TreeCtrlTest();

	static void SetupCacheFolderList(CArray<CString>& cacheFolderList);

	// Generated message map functions
protected:
	//{{AFX_MSG(NTreeView)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg void OnUpdateFileRefresh(CCmdUI* pCmdUI);
	afx_msg void OnFileRefresh();
	afx_msg void OnUpdateTreeExpand(CCmdUI* pCmdUI);
	afx_msg void OnTreeExpand();
	//}}AFX_MSG
	afx_msg void OnSelchanged(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnSelchanging(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnRClick(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg LRESULT OnCmdParam_FileName(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmdParam_GeneralHint(WPARAM wParam, LPARAM lParam);
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnTvnGetInfoTip(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg LRESULT OnCmdParam_ResetTreePos(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCmdParam_OnSwitchWindow(WPARAM wParam, LPARAM lParam);
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnTimer(UINT_PTR nIDEvent);
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_NTREEVIEW_H__28F2C044_02DA_4048_8F8F_3D17ADEE1421__INCLUDED_)
