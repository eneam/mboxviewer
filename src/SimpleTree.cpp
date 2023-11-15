#pragma once

#include "stdafx.h"
#include "SimpleTree.h"

#if 0
template <typename T>
SimpleTreeNodeEx<T>* SimpleTreeEx<T>::InsertNode(SimpleTreeNodeEx<T>* parentNode, CString& name)
{
	SimpleTreeNodeEx<T> *newNode = new SimpleTreeNodeEx<T>;

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

#if 0


template <typename T>
class SimpleTreeEx
{
public:
	SimpleTreeEx(CString& name) { m_name = name; }

	CString m_name;
	list<SimpleTreeNodeEx> m_rootNodeList;


	BOOL IsEmpty() { return m_rootNodeList.size() == 0; }
	int NodeCount(SimpleTreeNodeEx* node, BOOL recursive);


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
#endif

