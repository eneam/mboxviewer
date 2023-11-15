#pragma once


#include <minwindef.h>
#include <afxstr.h>
#include <algorithm>
#include <list>


class CFile;

using namespace std;

// MBoxFolderTree is the helper class to support "File->Select root folder" feature
// Easy to use incorrectly so be  careful
class SimpleTreeNode
{
public:
	SimpleTreeNode() { m_parent = 0; m_valid = TRUE; }
	CString m_folderName;
	SimpleTreeNode*m_parent;
	list<SimpleTreeNode> m_nodeList;
	BOOL m_valid;
};

class SimpleTree
{
public:
	SimpleTree(CString &name) { m_name = name; }
	CString m_name;
	list<SimpleTreeNode> m_rootList;


	BOOL IsEmpty() { return m_rootList.size() == 0;}
	void EraseRoot() { ; }  // assume m_root->m_nodeList.Empty() == true
	SimpleTreeNode*GetRootNode() { return 0; }

	//BOOL PopulateFolderTree(CString &rootFolder, MBoxFolderTree &tree, MBoxFolderNode *rnode, CString &errorText, int maxDepth);

	SimpleTreeNode*CreateNode(SimpleTreeNode*node);

	void Print(CString &filepath);
	void PrintNode(CFile *fp, SimpleTreeNode*node);

	int Count();
	int NodeCount(SimpleTreeNode*node);

	void PruneNonMBoxFolderNode(SimpleTreeNode*node);
	void PruneNonMBoxFolders();
	static void GetRelativeFolderPath(SimpleTreeNode*rnode, CString &folderPath);
};


template <typename T>
class SimpleTreeNodeEx
{
public:
	SimpleTreeNodeEx() { m_parent = 0; m_data = 0;  m_valid = TRUE; }
	SimpleTreeNodeEx<T>* m_parent;
	list<SimpleTreeNodeEx<T>*> m_nodeList;
	T* m_data;

	BOOL m_valid;
};



template <typename T>
class SimpleTreeEx
{
public:
	SimpleTreeEx(CString& name) { m_name = name; }

	CString m_name;
	list<SimpleTreeNodeEx<T> *> m_nodeList;


	BOOL IsEmpty() { return m_rootNodeList.size() == 0; }
	size_t NodeCount(SimpleTreeNodeEx<T>* node) { return node ? node->m_nodeList.size(): m_nodeList.size(); }
	
	SimpleTreeNodeEx<T>* InsertNode(SimpleTreeNodeEx<T>* parentNode, CString& name)
#if 1
	{
		SimpleTreeNodeEx<T>* newNode = new SimpleTreeNodeEx<T>;

		list<SimpleTreeNodeEx<T>*>* nodeList;

		if (parentNode == 0)
			nodeList = &m_nodeList;
		else
			nodeList = &parentNode->m_nodeList;

		newNode->m_parent = parentNode;
		nodeList->push_back(newNode);
		return nodeList->back();
	}
#endif
	//int NodeCount(SimpleTreeNodeEx* node, BOOL recursive);
#if 0


	//void EraseRoot() { ; }  // assume m_root->m_nodeList.Empty() == true
	N* GetRootNode() { return 0; }

	//BOOL PopulateFolderTree(CString &rootFolder, MBoxFolderTree &tree, MBoxFolderNode *rnode, CString &errorText, int maxDepth);

	SimpleTreeNodeEx* CreateNode(N* node);

	void Print(CString& filepath);
	//void PrintNode(CFile* fp, SimpleTreeNodeEx* node);

	int Count();
	int NodeCount(SimpleTreeNodeEx* node);

	//void PruneNonMBoxFolderNode(SimpleTreeNodeEx* node);
	//void PruneNonMBoxFolders();
	//static void GetRelativeFolderPath(N* rnode, CString& folderPath);
#endif
};