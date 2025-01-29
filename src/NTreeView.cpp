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

// NTreeView.cpp : implementation file
//

#include "stdafx.h"
#include "FileUtils.h"
#include "mboxview.h"
#include "MboxMail.h"
#include "OpenArchiveFileDlg.h"
#include "NTreeView.h"
#include "InputBox.h"
#include "CheckListBoxDlg.h"
#include "OpenContainingFolderDlg.h"
#include "TextUtilsEx.h"
#include "MimeParser.h"
#include "TextUtilities.h"
#include "ResHelper.h"
#include "FileConfigurationDB.h"


#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(NTreeView, CWnd)

/////////////////////////////////////////////////////////////////////////////
// NTreeView

int NTreeView::m_filesToValidateAsMboxType = 0;

NTreeView::NTreeView() :
	m_gFileSizes(113)
{
	m_treeViewAssertFlags = ASSERT_TREE_CTRL | ASSERT_FOLDERS_IN_REG_AND_MEM | ASSERT_FOLDERS_IN_MEM_AND_TREE;
	m_labelHT = new GmailLableMapType(113);
	m_mappedToLabelHT = new GmailMappedLabelMapType(113);
	m_labelView = FALSE;
	m_labelSeqNumb = 0;

	m_bIsDataDirty = FALSE;
	m_fileSizes.InitHashTable(113);
	fileSizes = &m_fileSizes;

	m_bSelectMailFileDone = FALSE;
	m_bSelectMailFilePostMsgDone = FALSE;
	m_bGeneralHintPostMsgDone = FALSE;
	m_timerTickCnt = 0;
	m_nIDEvent = 1;
	m_nElapse = 1;
	m_bInFillControl = FALSE;

	m_frameCx = 177;
	m_frameCy = 200;

	CString section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacement";
	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
		section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacementPreview";
	else if (CMainFrame::m_commandLineParms.m_bDirectFileOpenMode)
		section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacementDirect";

	BOOL ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_wnd, L"TreeFrameWidth", m_frameCx);
	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_wnd, L"TreeFrameHeight", m_frameCy);

	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	m_filesToValidateAsMboxType = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_general, L"filesToValidateAsMboxType");
}

NTreeView::~NTreeView()
{
	ClearGlobalFileSizeMap();

	m_mappedToLabelHT->clear();
	ClearLabelHT();

	delete m_labelHT;
	delete m_mappedToLabelHT;
}

BOOL NTreeView::ExpandItem(HTREEITEM hItem, UINT nCode)
{
	CString name = m_tree.GetItemText(hItem);
	BOOL ret = m_tree.Expand(hItem, nCode);
	return ret;
}

void NTreeView::ClearLabelHT()
{
	GmailLableMapType::IHashMapIter iter = m_labelHT->first();
	int totalCnt = 0;
	for (; !m_labelHT->last(iter); )
	{
		//m_labelHT->next(iter);
		GmailLabel* l = iter.element;
		m_labelHT->remove(iter);
		delete l;

	}
	m_labelHT->clear();
}

void NTreeView::ClearGlobalFileSizeMap()
{
	GlobalFileSizeMap::iterator it;

	ArchiveFileInfoMap* item;
	for (it = m_gFileSizes.begin(); it != m_gFileSizes.end(); it++)
	{
		item = it->second;
		delete item;
		it->second = 0;
	}

	m_gFileSizes.clear();
}

BOOL NTreeView::RemoveFileSizeMap(CString path)
{
	GlobalFileSizeMap::iterator it;

	path.TrimRight(L"\\");
	std::wstring stdPath = (LPCWSTR)path;
	it = m_gFileSizes.find(stdPath);
	ArchiveFileInfoMap* item;
	if (it != m_gFileSizes.end())
	{
		item = it->second;
		delete item;
		it->second = 0;
		m_gFileSizes.erase(stdPath);
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

BEGIN_MESSAGE_MAP(NTreeView, CWnd)
	//{{AFX_MSG_MAP(NTreeView)
	ON_WM_CREATE()
	ON_WM_SIZE()
	ON_WM_ERASEBKGND()
	ON_WM_MOUSEWHEEL()
	//ON_UPDATE_COMMAND_UI(ID_FILE_REFRESH, OnUpdateFileRefresh)
	//ON_COMMAND(ID_FILE_REFRESH, OnFileRefresh)
	ON_UPDATE_COMMAND_UI(ID_TREE_EXPAND, OnUpdateTreeExpand)
	ON_COMMAND(ID_TREE_EXPAND, OnTreeExpand)
	//}}AFX_MSG_MAP
	ON_NOTIFY(TVN_SELCHANGED, IDC_TREE, OnSelchanged)
	ON_NOTIFY(TVN_SELCHANGING, IDC_TREE, OnSelchanging)
	ON_NOTIFY(NM_RCLICK, IDC_TREE, OnRClick)  // Right Click Menu
	ON_WM_TIMER()
	ON_MESSAGE(WM_CMD_PARAM_FILE_NAME_MESSAGE, OnCmdParam_FileName)
	ON_MESSAGE(WM_CMD_PARAM_GENERAL_HINT_MESSAGE, OnCmdParam_GeneralHint)
	ON_NOTIFY(TVN_GETINFOTIP, IDC_TREE, OnTvnGetInfoTip)
	ON_MESSAGE(WM_CMD_PARAM_RESET_TREE_POS_MESSAGE, &NTreeView::OnCmdParam_ResetTreePos)
	ON_MESSAGE(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, &NTreeView::OnCmdParam_OnSwitchWindow)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// NTreeView message handlers

int CALLBACK MyCompareProc(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort)
{
	// lParamSort contains a pointer to the tree control.
	// The lParam of an item is just its handle, 
	// as specified with SetItemData . 
	// !!!!! Not very friendly. 
	// lParam1/lParam2 is what you set via SetItemData
	// I want to do sorting based on whther item has children

	NTreeView* pTreeView = (NTreeView*)lParamSort;
	DWORD nId1 = (DWORD)lParam1;
	DWORD nId2 = (DWORD)lParam2;

	LabelInfo* linfo1 = pTreeView->m_labelInfoStore.Find(nId1);
	LabelInfo* linfo2 = pTreeView->m_labelInfoStore.Find(nId2);
	CString strItem1;
	CString strItem2;
	BOOL isNode1Folder = FALSE;
	BOOL isNode2Folder = FALSE;
	if ((linfo1->m_nodeType == LabelInfo::MailFolder) || (linfo1->m_nodeType == LabelInfo::MailSubFolder))
		isNode1Folder = TRUE;
	if ((linfo2->m_nodeType == LabelInfo::MailFolder) || (linfo2->m_nodeType == LabelInfo::MailSubFolder))
		isNode2Folder = TRUE;

	if (linfo1->m_nodeType == LabelInfo::MailLabel)  // all nodes should better be MailLabels
	{
		return linfo1->m_label.Compare(linfo2->m_label);
	}
	else if ((isNode1Folder == FALSE) && (isNode2Folder == FALSE))
	{
		FileUtils::CPathStripPath(linfo1->m_filePath, strItem1);
		FileUtils::CPathStripPath(linfo2->m_filePath, strItem2);
		return strItem1.Compare(strItem2);
	}
	else if (isNode1Folder && isNode2Folder)
	{
		FileUtils::CPathStripPath(linfo1->m_mailFolderPath, strItem1);
		FileUtils::CPathStripPath(linfo2->m_mailFolderPath, strItem2);
		return strItem1.Compare(strItem2);
	}
	else if (isNode1Folder)
	{
		return 1;
	}
	else if (isNode2Folder)
	{
		return -1;
	}
	else
	{
		MboxMail::assert_unexpected();
		return -1;
	}

	return -1;
}

int NTreeView::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CWnd::OnCreate(lpCreateStruct) == -1)
		return -1;

	if (!m_tree.Create(WS_CHILD | WS_VISIBLE | TVS_INFOTIP |
		TVS_HASLINES | TVS_LINESATROOT | TVS_HASBUTTONS |
		TVS_SHOWSELALWAYS | TVS_EDITLABELS, // TVS_EDITLABELS | TVS_DISABLEDRAGDROP,
		CRect(0, 0, 0, 0), this, IDC_TREE))
		return -1;

	m_il.Create(16, 16, ILC_MASK | ILC_COLOR24, 0, 0);
	CBitmap bmp;
	bmp.LoadBitmap(IDB_IL_TREE);
	m_il.Add(&bmp, RGB(255, 255, 255));
	m_tree.SetImageList(&m_il, TVSIL_NORMAL);

	m_folderArray.LoadFromRegistry();

	m_folderArray.Dump();

	return 0;
}

void NTreeView::OnSize(UINT nType, int cx, int cy)
{
	CWnd::OnSize(nType, cx, cy);

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		int msgViewPosition = pFrame->MsgViewPosition();
		BOOL bTreeHideVal = pFrame->TreeHideValue();
		BOOL isTreeHidden = pFrame->IsTreeHidden();

		TRACE(L"NTreeView::OnSize cx=%d cy=%d viewPos=%d IsTreeHideVal=%d IsTreeHidden=%d\n",
			cx, cy, msgViewPosition, bTreeHideVal, isTreeHidden);

#if 0
		// No longer needed, delete.
		if (cx > 0)
		{
			m_frameCx = cx;
			m_frameCy = cy;
		}
#endif
	}
	m_tree.MoveWindow(0, 0, cx, cy);
}

void NTreeView::PostNcDestroy()
{
	m_font.DeleteObject();
	m_il.DeleteImageList();
	DestroyWindow();
	delete this;

	//	CWnd::PostNcDestroy();
}

#define BUFF_PREVIEW_SIZE	6144

BOOL NTreeView::ImboxviewFileFilter(CString& fName)
{
	if (m_filesToValidateAsMboxType == 1)
		return TRUE;

	CString fileNameExtension;
	FileUtils::GetFileExtension(fName, fileNameExtension);

	if (fileNameExtension.IsEmpty())
		return TRUE;

	if (fileNameExtension.CompareNoCase(L".mbox") &&
		fileNameExtension.CompareNoCase(L".eml") &&
		fileNameExtension.CompareNoCase(L".mboxo") &&
		fileNameExtension.CompareNoCase(L".mboxrd") &&
		fileNameExtension.CompareNoCase(L".mboxcl") &&
		fileNameExtension.CompareNoCase(L".mboxcl2")
		)
	{
		return FALSE;
	}
	else
		return TRUE;
}

// This function checks whether file archive is valid
// TODO: add MessageBox to report an error
int NTreeView::ImboxviewFile(CString& fName)
{
#if 0
	if (fName.Compare("F:\\MBOX\\MailTkout1\\message-inline2.eml") == 0)
		int deb = 1;

	if (fName.Compare("G:\\GMAIL\\takeout-20210705T072046Z-001.zip") == 0)
		int deb = 1;

	if (fName.Compare("G:\\GMAIL\\takeout-20210713T143614Z-001.zip") == 0)
		int deb = 1;
#endif

	if (!ImboxviewFileFilter(fName))
		return 0;

	LPCWSTR fileName = fName;

	// need better validation check; may need to parse the file
	  // 0=not valid, 1=valid, -1=end of header not found and not valid yet so allocate larger buffer and redo
	CFile fp;

	int retval = 0;
	CString fileNameExtention;
	FileUtils::GetFileExtension(fName, fileNameExtention);

	if (fileNameExtention.CompareNoCase(L".mboxview") == 0)
		return 0;

	else if (fileNameExtention.CompareNoCase(L".mboxlist") == 0)
		return 0;

	else if (fileNameExtention.CompareNoCase(L".urootfolder") == 0)
		return 0;

	CFileException ExError;
	if (!fp.Open(fileName, CFile::modeRead, &ExError))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError); // TODO
		return 0;
	}

	char buffer[BUFF_PREVIEW_SIZE + 1];
	char* buff = &buffer[0];
	char* mallocbuff = 0;
	char* newMallocbuff = 0;
	char* data;
	int datalen;
	UINT nCount;
	int i;
	// TODO: Not the most efficient
	for (i = 0; i < 20; i++)
	{
		nCount = fp.Read(buff, BUFF_PREVIEW_SIZE);
		if (nCount == 0) {
			break;
		}

		if (i == 0)
		{
			data = &buff[0];
			datalen = nCount;
		}
		else
		{
			datalen = i * BUFF_PREVIEW_SIZE;
			memcpy(&mallocbuff[datalen], buff, nCount);
			datalen += nCount;
			data = mallocbuff;
		}

		retval = NTreeView::IsValidMailFile(data, datalen);

		if (nCount < BUFF_PREVIEW_SIZE)
			break;

		if (retval == -1)
		{
			if (i > 0)
				int deb = 1;
			// Allocate extra space for next file read
			newMallocbuff = (char*)malloc((i + 1 + 1) * BUFF_PREVIEW_SIZE);
			if (newMallocbuff == 0)
				break;

			datalen = 0;
			if (mallocbuff) // i > 0
			{
				datalen = i * BUFF_PREVIEW_SIZE;
				memcpy(&newMallocbuff[0], &mallocbuff[0], datalen);
				free(mallocbuff);
			}

			memcpy(&newMallocbuff[datalen], buff, nCount);
			mallocbuff = newMallocbuff;
			datalen += nCount;
			continue;
		}
		else
			break;
	}
	if (i > 0)
		int deb = 1;

	fp.Close();
	if (mallocbuff)
		free(mallocbuff);

	return retval;
}

int NTreeView::IsValidMailFile(char* data, int datalen)
{
	static const char* cFromMailBegin = "from ";
	static const int cFromMailBeginLen = istrlen(cFromMailBegin);
	static const char* cFrom = "from:";
	static const int cFromLen = istrlen(cFrom);
	static const char* cTo = "to:";
	static const int cToLen = istrlen(cTo);
	static const char* cSubject = "subject:";
	static const int cSubjectLen = istrlen(cSubject);
	static const char* cDate = "date:";
	static const int cDateLen = istrlen(cDate);
	static const char* cMimeVersion = "mime-version:";
	static const int cMimeVersionLen = istrlen(cMimeVersion);
	static const char* cMessageID = "message-id:";
	static const int cMessageIDLen = istrlen(cMessageID);

	BOOL bMboxFrom = FALSE;
	BOOL bFrom = FALSE;
	BOOL bTo = FALSE;
	BOOL bSubject = FALSE;
	BOOL bMimeVersion = FALSE;
	BOOL bMessageID = FALSE;
	BOOL bDate = FALSE;

	char* p = data;
	char* e = data + datalen;
	char* psave = p;


	// Simple approach, not bullet prove but should work just fine I think
	//
	// Empty line indicates the end of mime mail header
	// Ignore lines that start with white spaces (space and tab) i.e. folded lines
	// Except first "From " line if present
	// lines must contain ':" character to indicate end of header field
	// Lines must not be longer then lets say 2000 characters
	//
	// Header must contain the following fields: 
	// Mime-Version: and From: fields
	// to be valid mime mail header
	// What about Subject:, To:, Date:, Received:, Nessage-ID fields ???
	// TODO: Should list of required fileds be configrable?


	int maxLineLenght = 2000;
	int maxLineLen = maxLineLenght;

	p = MimeParser::SkipEmptyLines(p, e);
	if (p >= e)
		return 0;

	if (TextUtilsEx::strncmpUpper2Lower(p, e, cFromMailBegin, cFromMailBeginLen) == 0)
	{
		bMboxFrom = TRUE;
		maxLineLen = maxLineLenght;
		p = MimeParser::EatNewLine(p, e, maxLineLen);
		if (maxLineLen <= 0)
			return 0;
		else if (p >= e)
			return -1;
	}
	else
	{
		int deb = 1;
	}

	int allFieldCnt = 0;
	BOOL endOfHeaderFound = FALSE;

	while (p < e)
	{
		if ((*p == '\n') || ((*p == '\r') && (*(p + 1) == '\n')))  // empty line
		{
			endOfHeaderFound = TRUE;
			break;
		}
		else
		{
			if ((*p == ' ') || (*p == '\t'))
			{
				maxLineLen = maxLineLenght;
				p = MimeParser::EatNewLine(p, e, maxLineLen);
				if (maxLineLen <= 0)
					return 0;
				continue;
			}

			psave = p;
			maxLineLen = 100;
			while ((p < e) && (*p != '\n') && (*p != ':') && (maxLineLen-- > 0))
			{
				p++;
			}
			if (maxLineLen <= 0)
				return 0;
			else if (p >= e)  // TODO: didn't find ':' character
				break;

			allFieldCnt++;
			if (TextUtilsEx::strncmpUpper2Lower(psave, e, cFrom, cFromLen) == 0)
			{
				bFrom = TRUE;
				// TODO: should I check here if valid header ??
				// or continue to scan the entire header
				// From, To, Subject are usually at the end  of header anyway
			}
			else if (TextUtilsEx::strncmpUpper2Lower(psave, e, cMimeVersion, cMimeVersionLen) == 0)
			{
				bMimeVersion = TRUE;
			}
			else if (TextUtilsEx::strncmpUpper2Lower(psave, e, cMessageID, cMessageIDLen) == 0)
			{
				bMessageID = TRUE;
			}
			else if (TextUtilsEx::strncmpUpper2Lower(psave, e, cTo, cToLen) == 0)
			{
				bTo = TRUE;
			}
			else if (TextUtilsEx::strncmpUpper2Lower(psave, e, cSubject, cSubjectLen) == 0)
			{
				bSubject = TRUE;
			}
			else if (TextUtilsEx::strncmpUpper2Lower(psave, e, cDate, cDateLen) == 0)
			{
				bDate = TRUE;
			}

			maxLineLen = maxLineLenght;
			p = MimeParser::EatNewLine(p, e, maxLineLen);
			if (maxLineLen <= 0)
				return 0;
		}
	}

	// messageID is always present in my test mail files
	// Not sure I can assume that
	if (!bMessageID)
		int deb = 1;

	if (CMainFrame::m_relaxedMboxFileValidation == FALSE)
	{
		if (bFrom)
		{
			if ((bMimeVersion && (bMboxFrom || bMessageID)) ||
				(bMboxFrom && bMessageID) ||
				((bFrom && bSubject) && (bMimeVersion || bMessageID)) ||
				((bTo && bSubject) && (bMimeVersion || bMessageID))
				)
				if (bMboxFrom)
					return 1;
				else
					return 2;
		}
	}
	else
	{
		if (bFrom || bTo || bSubject || bDate)
			return 1;
	}
	if (endOfHeaderFound == FALSE)
		return -1;
	else
		return 0;
}

// Updated: recursive. Relies on Label Info instead of m_folderArray. Any issue with the chnage ?? FIXMEFIXME
// Non recursive. Iterates siblings 
// Searches m_folderArray for path 
// NTreeView::IsFolderOpen(CString &path) is doing the same except it relies on Label Info
HTREEITEM NTreeView::HasFolder(CString& path)
{
	HTREEITEM hFolder = FindFolder(path);
	return hFolder;
}

HTREEITEM NTreeView::HasFolder(HTREEITEM hItem, CString& path)
{
	HTREEITEM hFolder = FindFolder(hItem, path);
	return hFolder;
}


// Updated: recursive now. !!! FIXMEFIXME
// Non recursive. Iterates siblings 
// Search label store
HTREEITEM NTreeView::IsFolderOpen(CString& path)
{
	HTREEITEM hFolder = FindFolder(path);
	return hFolder;
}

HTREEITEM NTreeView::FindMailFile(CString& mailFilePath)
{
	CString folderPath;
	CString fileName;
	FileUtils::GetFolderPathAndFileName(mailFilePath, folderPath, fileName);
	HTREEITEM hFolder = FindFolder(folderPath);
	HTREEITEM hFile = FindItem(hFolder, fileName);

	return hFile;
}

HTREEITEM NTreeView::FindLabelFile(CString& mailFilePath, CString& labelFilePath)
{
	HTREEITEM hMailFile = FindMailFile(mailFilePath);
	HTREEITEM hLabelFile = FindLabelItemByLabelPath(hMailFile, labelFilePath);

	return hLabelFile;
}

HTREEITEM NTreeView::FindFolder(CString& folderPath)
{
	HTREEITEM hRoot = 0;
	HTREEITEM hFolder = FindFolder(hRoot, folderPath);
	return hFolder;
}

HTREEITEM NTreeView::FindFolder(HTREEITEM hRoot, CString& path)
{
	BOOL recursive = TRUE;
	HTREEITEM hItem;

	BOOL retFind = FindFolder(hRoot, path, hItem, recursive);
	if (retFind && hItem)
		return hItem;
	else
		return 0;
}


// Rather Expensive
BOOL NTreeView::FindFolder(HTREEITEM hRoot, CString& path, HTREEITEM& hFolder, BOOL recursive)
{
	BOOL retval = FALSE;
	CString folderPath;
	CString itemNameChild;
	CString normalizedPath = path;
	normalizedPath.TrimRight(L"\\");

	hFolder = 0;

	HTREEITEM hChild = m_tree.GetChildItem(hRoot);
	while (hChild)
	{
		itemNameChild = m_tree.GetItemText(hChild);
		retval = FindFolder_internal(hChild, path, hFolder, recursive);
		if (retval)
		{
			_ASSERTE(hFolder != 0);
			if (hFolder == 0)
			{
				retval = FALSE;
			}
			break;
		}
		hChild = m_tree.GetNextSiblingItem(hChild);
	}
	return retval;
}

BOOL NTreeView::FindFolder_internal(HTREEITEM hRoot, CString& path, HTREEITEM& hFolder, BOOL recursive)
{
	CString itemName;

	BOOL retval = FALSE;
	CString folderPath;
	CString folderPathFromArray;
	CString normalizedPath = path;
	normalizedPath.TrimRight(L"\\");

	_ASSERTE(hRoot != 0);
	if (hRoot == 0)
		return FALSE;

	if (hFolder)
		return TRUE;

	int index = (DWORD)m_tree.GetItemData(hRoot);
	itemName = m_tree.GetItemText(hRoot);

	if (path.Find(itemName) < 0)
		return FALSE;

	LabelInfo* linfo = m_labelInfoStore.Find(index);
	if (linfo == 0)
		return FALSE;

	if ((linfo->m_nodeType == LabelInfo::MailFolder) || (linfo->m_nodeType == LabelInfo::MailSubFolder))
	{
		folderPath = linfo->m_mailFolderPath;

		TRACE(L"FindFolder_internal: Folder \"%s\"\n", folderPath);
		if (normalizedPath.CompareNoCase(folderPath) == 0)
		{
#if 1
			// This doesn't work when folder is open as sub-folder within the tree root
			// Are there ay cases we shoulkd rely on folderPathFromArray in Find/Has function as before ??  ?? FIXMEFIXME
			m_folderArray.GetAt(index, folderPathFromArray);
			folderPathFromArray.TrimRight(L"\\");

			if (folderPath.Compare(folderPathFromArray))
				_ASSERTE(FALSE);
#endif
			hFolder = hRoot;
			return TRUE;
		}
	}
	else
	{
		TRACE(L"FindFolder_internal: File \"%s\"\n", linfo->m_filePath);
		return FALSE;
	}

	BOOL hasChildren = m_tree.ItemHasChildren(hRoot);
	if (recursive && hasChildren)
	{
		HTREEITEM child = m_tree.GetChildItem(hRoot);
		while (child != NULL)
		{
			retval = FindFolder_internal(child, path, hFolder, recursive);
			if (retval)
			{
				_ASSERTE(hFolder != 0);
				if (hFolder == 0)
				{
					retval = FALSE;
				}
				return retval;
			}
			child = m_tree.GetNextSiblingItem(child);
		}
	}
	return retval;
}

BOOL NTreeView::FolderHasChilder(CString& path)
{
	HTREEITEM hRoot = 0;
	HTREEITEM hFolder = 0;
	BOOL retval = FindFolder(hRoot, path, hFolder, TRUE);
	if (retval && hFolder)
	{
		if (m_tree.ItemHasChildren(hFolder))
			return TRUE;
	}
	return FALSE;
}

BOOL NTreeView::PrintFolderNames(HTREEITEM hRoot, BOOL recursive)
{
	BOOL retval = FALSE;
	CString folderPath;
	CString itemNameChild;

	HTREEITEM hChild = m_tree.GetChildItem(hRoot);
	while (hChild)
	{
		itemNameChild = m_tree.GetItemText(hChild);
		retval = PrintFolderNames_internal(hChild, recursive);
		if (retval == FALSE)
			break;

		hChild = m_tree.GetNextSiblingItem(hChild);
	}
	return retval;
}

BOOL NTreeView::PrintFolderNames_internal(HTREEITEM hRoot, BOOL recursive)
{
	CString itemName;

	BOOL retval = TRUE;
	CString folderPath;

	_ASSERTE(hRoot != 0);
	if (hRoot == 0)
		return FALSE;

	int index = (DWORD)m_tree.GetItemData(hRoot);
	itemName = m_tree.GetItemText(hRoot);

	m_folderArray.GetAt(index, folderPath);

	LabelInfo* linfo = m_labelInfoStore.Find(index);
	if (linfo == 0)
		return FALSE;

	if ((linfo->m_nodeType == LabelInfo::MailFolder) || (linfo->m_nodeType == LabelInfo::MailSubFolder))
	{
		TRACE(L"PrintFolderNames_internal: Folder \"%s\"\n", linfo->m_mailFolderPath);
	}
	else
	{
		TRACE(L"PrintFolderNames_internal: File \"%s\"\n", linfo->m_filePath);
	}

	BOOL hasChildren = m_tree.ItemHasChildren(hRoot);
	if (recursive && hasChildren)
	{
		HTREEITEM child = m_tree.GetChildItem(hRoot);
		while (child != NULL)
		{
			retval = PrintFolderNames_internal(child, recursive);
			if (retval == FALSE)
				break;

			child = m_tree.GetNextSiblingItem(child);
		}
	}
	return retval;
}

// Read folder names from registry and load
void NTreeView::LoadFolders()
{
	CString pathLast = MboxMail::GetLastPath();
	// Make sure data path is also set
	if (!pathLast.IsEmpty())
		MboxMail::SetLastPath(pathLast);

	CString path;
	int paneId = 0;
	CString sText;
	CString str;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	BOOL lastPathLoaded = FALSE;
	BOOL lastPathValid = FALSE;
	BOOL expand = FALSE;
	HTREEITEM hParent = TVI_ROOT;
	BOOL selectFolder = FALSE;
	BOOL unconditionalFolderInsert = FALSE;

	m_treeViewAssertFlags = ASSERT_TREE_CTRL | ASSERT_FOLDERS_IN_REG_AND_MEM;
	int ii = 0;
	for (ii = 0; ii < m_folderArray.GetCount(); ii++)
	{
		BOOL retGet = m_folderArray.GetAt(ii, path);
		if (!path.IsEmpty())
		{
			if (path.Compare(L"F://NewMerge") == 0)   // FIXME
			{
				int deb = 1;
			}
			if (path.CompareNoCase(pathLast) == 0)
			{
				lastPathValid = TRUE;
			}
			pathLast = MboxMail::GetLastPath();

			MboxMail::SetLastPath(path);
			FileUtils::CPathStripPath(path, str);

			if (IsFolderARootFolder(path))
			{
				CString fmt = L"Opening root folder %s ...";
				ResHelper::TranslateString(fmt);

				sText.Format(fmt, str);

				if (pFrame)
					pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

				HTREEITEM hParent = TVI_ROOT;
				OpenRootFolderAndSubfolders(hParent, path);

				TRACE(L"============>> Load root folder %s done\n", path);
				m_globalFolderInfoDB.Print();
			}
			else
			{
				BOOL expand = FALSE;
				BOOL unconditionalFolderInsert = FALSE;
				BOOL selectFolder = FALSE;

				HTREEITEM hOpenFolder = IsFolderOpen(path);
				if (!hOpenFolder)
				{
					DoOpenFolder(path, selectFolder, unconditionalFolderInsert, expand);
					TRACE(L"============>> Load folder %s done\n", path);
					m_globalFolderInfoDB.Print();
				}
				else
				{
					int index1 = m_folderArray.Delete(ii, path);
					MboxMail::SetLastPath(pathLast);
					const int deb = 1;
				}
			}

			sText.Format(L"Ready");
			ResHelper::TranslateString(sText);
			if (pFrame)
				pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
		}

		int deb = 1;
	}

	OpenLastSelection();
	int deb = 1;

	this->SortChildren(TVI_ROOT);
	TRACE(L"============>> Load Folders done\n");
	m_globalFolderInfoDB.Print();

	TRACE(L"============>> Assert after Load Folders done\n");
	//BOOL retAssert = AssertTreeCtrl();
	m_treeViewAssertFlags = ASSERT_TREE_CTRL | ASSERT_FOLDERS_IN_REG_AND_MEM | ASSERT_FOLDERS_IN_MEM_AND_TREE;
	BOOL retAssert = AssertNTreeView();
}

BOOL NTreeView::ImportLegacyRegistryData()
{
	if (CProfile::IsRegistryConfig())
	{
		const wchar_t* sz_Software_mboxview_Legacy = L"SOFTWARE\\mboxview";
		BOOL retLegacyExists = CProfile::CheckIfRegistryKeyExists(HKEY_CURRENT_USER, CString(sz_Software_mboxview_Legacy));
		if (retLegacyExists == TRUE)
		{
			NTreeView::ImportLegacyMboxviewRegistryData();

			CString section = L"SOFTWARE\\UMBoxViewer\\General";
			BOOL ret = CProfile::_DeleteValue(HKEY_CURRENT_USER, section, L"dataFolder");
			int deb = 1;
		}
	}
	else
	{
		const wchar_t* sz_Software_mboxview_Legacy = L"SOFTWARE\\UMBoxViewer\\General";
		BOOL retLegacyExists = CProfile::CheckIfRegistryKeyExists(HKEY_CURRENT_USER, CString(sz_Software_mboxview_Legacy));
		if (retLegacyExists == TRUE)
		{
			NTreeView::ImportLegacyUMBoxViewerRegistryData();
			int deb = 1;
		}
		else
		{
			const wchar_t* sz_Software_mboxview_Legacy = L"SOFTWARE\\mboxview";
			BOOL retLegacyExists = CProfile::CheckIfRegistryKeyExists(HKEY_CURRENT_USER, CString(sz_Software_mboxview_Legacy));
			if (retLegacyExists == TRUE)
			{
				NTreeView::ImportLegacyMboxviewRegistryData();

				CString section = L"SOFTWARE\\UMBoxViewer\\General";
				BOOL ret = CProfile::_DeleteValue(HKEY_CURRENT_USER, section, L"dataFolder");
				int deb = 1;
			}
		}
	}

	ConfigTree* confTree = CProfile::GetConfigTree();
#ifdef _DEBUG
	CString label = L"NTreeView::ImportLegacyRegistryData";
	if (confTree)
		confTree->DumpTree(label);
#endif

	return TRUE;
}


BOOL NTreeView::ImportLegacyMboxviewRegistryData()
{
	LSTATUS retCode;
	HKEY hKey = HKEY_CURRENT_USER;

	const wchar_t* sz_Software_mboxview_Legacy = L"SOFTWARE\\mboxview";

	CString fromSection = CString(sz_Software_mboxview_Legacy) + L"\\folders";
	CString toSection = CString(sz_Software_mboxview) + L"\\MailFolders";

	TRACE(L"ImportLegacyRegistryData: folders\n");
	LSTATUS retCode0 = CProfile::CopyKey(hKey, (LPCWSTR)fromSection, (LPCWSTR)toSection);

	fromSection = CString(sz_Software_mboxview_Legacy) + L"\\MailService";
	toSection = CString(sz_Software_mboxview) + L"\\MailService";

	TRACE(L"ImportLegacyRegistryData: MailService\n");
	LSTATUS retCode1 = CProfile::CopyKey(hKey, (LPCWSTR)fromSection, (LPCWSTR)toSection);

	TRACE(L"ImportLegacyRegistryData: MailService SubKeys\n");
	LSTATUS retCode2 = CProfile::CopySubKeys(hKey, (LPCWSTR)fromSection, (LPCWSTR)toSection);

	fromSection = CString(sz_Software_mboxview_Legacy) + L"\\WindowPlacement";
	toSection = CString(sz_Software_mboxview) + L"\\WindowPlacement";

	TRACE(L"ImportLegacyRegistryData: Window Placement\n");
	LSTATUS retCode3 = CProfile::CopyKey(hKey, (LPCWSTR)fromSection, (LPCWSTR)toSection);

	fromSection = CString(sz_Software_mboxview_Legacy) + L"\\mboxview\\MRUs";
	toSection = CString(sz_Software_mboxview) + L"\\mboxview\\MRUs";

	TRACE(L"ImportLegacyRegistryData: MRUs\n");
	LSTATUS retCode4 = CProfile::CopyKey(hKey, (LPCWSTR)fromSection, (LPCWSTR)toSection);

	fromSection = CString(sz_Software_mboxview_Legacy);

	//##############
	//toSection = CString(sz_Software_mboxview) + L"\\General";
	toSection = CString(sz_Software_mboxview) + L"\\WindowPlacement";

	KeyFromToTable arr;
	arr.SetSize(1);
	arr.RemoveAll();

	arr.SetSize(10);  // Just a test
	arr[0] = { REG_DWORD, L"messageWindowPosition", L"" };
	arr.SetSize(1);

	retCode = CProfile::CopyKeyValueList(hKey, fromSection, toSection, arr);

	//##############
	toSection = CString(sz_Software_mboxview) + L"\\LastSelection";

	arr.RemoveAll();
	arr.SetSize(5);
	arr[0] = { REG_SZ, L"lastLabelFilePath", L"" };
	arr[1] = { REG_SZ, L"lastMailFilePath", L"" };
	arr[2] = { REG_DWORD, L"lastMailIndex", L"" };
	arr[3] = { REG_SZ, L"lastPath", L"" };
	arr[4] = { REG_DWORD, L"lastWhichSort", L"" };

	retCode = CProfile::CopyKeyValueList(hKey, fromSection, toSection, arr);

	//##############
	toSection = CString(sz_Software_mboxview) + L"\\Options";

	arr.RemoveAll();
	arr.SetSize(8);
	arr[0] = { REG_DWORD, L"format", L"dayMonthYearFormat" };
	arr[1] = { REG_DWORD, L"progressBarDelay", L"" };
	arr[2] = { REG_SZ, L"exportEML", L"" };
	arr[3] = { REG_DWORD, L"imageViewer", L"" };
	arr[4] = { REG_DWORD, L"timeType", L"" };
	arr[5] = { REG_DWORD, L"enhancedSelectFolderDialog", L"" };
	arr[6] = { REG_DWORD, L"subjectSortType", L"" };
	arr[7] = { REG_DWORD, L"filesToValidateAsMboxType", L"" };

	retCode = CProfile::CopyKeyValueList(hKey, fromSection, toSection, arr);

	//##############
	toSection = CString(sz_Software_mboxview) + L"\\Options\\SetCharacterSetIds";

	arr.RemoveAll();
	arr.SetSize(4);
	arr[0] = { REG_DWORD, L"fromCharsetId", L"" };
	arr[1] = { REG_DWORD, L"toCharsetId", L"" };
	arr[2] = { REG_DWORD, L"subjCharsetId", L"" };
	arr[3] = { REG_DWORD, L"showCharsets", L"" };

	retCode = CProfile::CopyKeyValueList(hKey, fromSection, toSection, arr);

	//##############
	toSection = CString(sz_Software_mboxview) + L"\\PaneColors";

	arr.RemoveAll();
	arr.SetSize(2);
	arr[0] = { REG_DWORD, L"colorStyle", L"" };
	arr[1] = { REG_SZ, L"customStyleColorList", L"" };

	retCode = CProfile::CopyKeyValueList(hKey, fromSection, toSection, arr);

	//##############
	toSection = CString(sz_Software_mboxview) + L"\\PrintConfig\\HeaderFields";

	arr.RemoveAll();
	arr.SetSize(2);
	arr[0] = { REG_DWORD, L"hdrFld_FontStyleDflt", L"ApplyCustomFonts" };
	arr[1] = { REG_DWORD, L"hdrFldBitmap", L"hdrFldListBitmap" };

	retCode = CProfile::CopyKeyValueList(hKey, fromSection, toSection, arr);

	//##############
	toSection = CString(sz_Software_mboxview) + L"\\PrintConfig\\HeaderFields\\DefaultFont";

	arr.RemoveAll();
	arr.SetSize(2);
	arr[0] = { REG_DWORD, L"hdrFld_BoldFontName", L"IsBold" };
	arr[1] = { REG_DWORD, L"hdrFld_FontSize", L"Size" };

	retCode = CProfile::CopyKeyValueList(hKey, fromSection, toSection, arr);

	//##############
	toSection = CString(sz_Software_mboxview) + L"\\PrintConfig\\HeaderFields\\NameFont";

	arr.RemoveAll();
	arr.SetSize(7);
	arr[0] = { REG_DWORD, L"hdrFldName_FontIsBold", L"IsBold" };
	arr[1] = { REG_DWORD, L"hdrFldName_FontIsItalic", L"IsItalic" };
	arr[2] = { REG_SZ, L"hdrFldName_FontName", L"Name" };
	arr[3] = { REG_DWORD, L"hdrFldName_FontSize", L"Size" };
	arr[4] = { REG_DWORD, L"hdrFldName_FontStyle", L"StyleNumber" };
	arr[5] = { REG_SZ, L"hdrFldName_FontStyleName", L"StyleName" };
	arr[6] = { REG_SZ, L"hdrFldName_GenericFontName", L"GenericName" };

	retCode = CProfile::CopyKeyValueList(hKey, fromSection, toSection, arr);

	//##############
	toSection = CString(sz_Software_mboxview) + L"\\PrintConfig\\HeaderFields\\TextFont";

	arr.RemoveAll();
	arr.SetSize(7);
	arr[0] = { REG_DWORD, L"hdrFldText_FontIsBold", L"IsBold" };
	arr[1] = { REG_DWORD, L"hdrFldText_FontIsItalic", L"IsItalic" };
	arr[2] = { REG_SZ, L"hdrFldText_FontName", L"Name" };
	arr[3] = { REG_DWORD, L"hdrFldText_FontSize", L"Size" };
	arr[4] = { REG_DWORD, L"hdrFldText_FontStyle", L"StyleNumber" };
	arr[5] = { REG_SZ, L"hdrFldText_FontStyleName", L"StyleName" };
	arr[6] = { REG_SZ, L"hdrFldText_GenericFontName", L"genericFontName" };

	retCode = CProfile::CopyKeyValueList(hKey, fromSection, toSection, arr);

	//##############
	toSection = CString(sz_Software_mboxview) + L"\\PrintConfig";

	arr.RemoveAll();
	arr.SetSize(12);
	arr[0] = { REG_DWORD, L"printFileNameCustomTemplate", L"useFileNameCustomTemplate" };
	arr[1] = { REG_DWORD, L"printPageHeaderAndFooter", L"" };
	arr[2] = { REG_DWORD, L"printMailBodyBackgroundColor", L"" };
	arr[3] = { REG_DWORD, L"printPDFPageBreakAfterEachMailConversationThread", L"printPageBreakAfterEachMailConversationThread" };
	arr[4] = { REG_DWORD, L"printPDFPageBreakAfterEachMail", L"printPageBreakAfterEachMail" };
	arr[5] = { REG_DWORD, L"printMailHdrBackgroundColor", L"" };
	arr[6] = { REG_SZ, L"printUserDefinedScriptPath", L"userDefinedScriptPath" };
	arr[7] = { REG_SZ, L"printMSEdgeBrowserPath", L"MSEdgeBrowserPath" };
	arr[8] = { REG_SZ, L"printChromeBrowserPath", L"ChromeBrowserPath" };
	arr[9] = { REG_DWORD, L"printHTML2PDFScriptType", L"" };
	arr[10] = { REG_DWORD, L"printDialogType", L"" };
	arr[11] = { REG_DWORD, L"printFileNameMaxLength", L"fileNameMaxLength" };

	retCode = CProfile::CopyKeyValueList(hKey, fromSection, toSection, arr);
	
	//##############
	toSection = CString(sz_Software_mboxview) + L"\\PrintConfig\\FileNameCustomTemplate";

	arr.RemoveAll();
	arr.SetSize(3);
	arr[0] = { REG_SZ, L"printFileNameCustomTemplateFormat", L"templateFormat" };
	arr[1] = { REG_SZ, L"printFileNameCustomTemplateDateFormat", L"dateFormat" };
	arr[2] = { REG_DWORD, L"printAddressPartsBitmap", L"addressPartsBitmap" };

	retCode = CProfile::CopyKeyValueList(hKey, fromSection, toSection, arr);


	//##############
	toSection = CString(sz_Software_mboxview) + L"\\Attachments";

	arr.RemoveAll();
	arr.SetSize(3);
	arr[0] = { REG_DWORD, L"attachmentWindowMaxSize", L"" };
	arr[1] = { REG_DWORD, L"showAllAttachments_Window", L"" };
	arr[2] = { REG_DWORD, L"anyAttachment_Indicator", L"" };

	retCode = CProfile::CopyKeyValueList(hKey, fromSection, toSection, arr);


	// ################
	CString section = CString(sz_Software_mboxview);

	TRACE(L"EnumerateAllSubKeys\n");
	LSTATUS retCode20 = CProfile::EnumerateAllSubKeys(hKey, (LPCWSTR)section);

	TRACE(L"EnumerateAllSubKeyValues\n");
	LSTATUS retCode21 = CProfile::EnumerateAllSubKeyValues(hKey, (LPCWSTR)section);

	int deb = 1;
	return TRUE;
}

BOOL NTreeView::CopyKeyAndSubkey(const wchar_t* keyPath, wchar_t* keyName)
{
	HKEY hKey = HKEY_CURRENT_USER;

	CString fromSection = CString(keyPath) + L"\\" + keyName;
	CString toSection = fromSection;

	TRACE(L"ImportLegacyRegistryData: %s\n", keyName);
	LSTATUS retCode1 = CProfile::CopyKey(hKey, (LPCWSTR)fromSection, (LPCWSTR)toSection);

	TRACE(L"ImportLegacyRegistryData: %s SubKeys\n", keyName);
	LSTATUS retCode2 = CProfile::CopySubKeys(hKey, (LPCWSTR)fromSection, (LPCWSTR)toSection);

	return TRUE;
}


BOOL NTreeView::ImportLegacyUMBoxViewerRegistryData()
{
	// TODO: connsider to implement recursive registry walk
	const wchar_t* sz_Software_mboxview_Legacy = L"SOFTWARE\\UMBoxViewer";
	BOOL ret = TRUE;
	ret = CopyKeyAndSubkey(sz_Software_mboxview_Legacy, L"\\Attachments");
	ret = CopyKeyAndSubkey(sz_Software_mboxview_Legacy, L"\\General");
	ret = CopyKeyAndSubkey(sz_Software_mboxview_Legacy, L"\\LastSelection");
	ret = CopyKeyAndSubkey(sz_Software_mboxview_Legacy, L"\\MailFolders");
	ret = CopyKeyAndSubkey(sz_Software_mboxview_Legacy, L"\\MailService");
	ret = CopyKeyAndSubkey(sz_Software_mboxview_Legacy, L"\\mboxview");
	ret = CopyKeyAndSubkey(sz_Software_mboxview_Legacy, L"\\Options");
	ret = CopyKeyAndSubkey(sz_Software_mboxview_Legacy, L"\\PaneColors");
	ret = CopyKeyAndSubkey(sz_Software_mboxview_Legacy, L"\\PrintConfig");
	ret = CopyKeyAndSubkey(sz_Software_mboxview_Legacy, L"\\PrintConfig\\HeaderFields");
	ret = CopyKeyAndSubkey(sz_Software_mboxview_Legacy, L"\\WindowPlacement");
	ret = CopyKeyAndSubkey(sz_Software_mboxview_Legacy, L"\\WindowPlacementDirect");

	ConfigTree* confTree = CProfile::GetConfigTree();
	_ASSERTE(confTree);
	//confTree->Dump2File(currentPath);

	return TRUE;
}

void NTreeView::DeleteMBoxAllWorkFolders(CString& mboxFileName)
{
	// Make sure the last pah and last data path are already set
	//MboxMail::SetLastPath(pathLast);

	CString path = MboxMail::GetLastPath();
	CString dataPath = MboxMail::GetLastDataPath();

	CString fileNameExtention;
	CString fileBaseName;
	FileUtils::GetFileBaseNameAndExtension(mboxFileName, fileBaseName, fileNameExtention);

	CString cacheFolderPath = dataPath + fileBaseName;
	bool recursive = true;
	bool removeFolders = true;
	BOOL retRD = FileUtils::RemoveDir(cacheFolderPath, recursive, removeFolders);
	int deb = 1;
}

void  NTreeView::ExpandOrCollapseTree(BOOL expand)
{
	HTREEITEM hFolder = 0;
#if 1
	// TODO: Need to better here and keep cuurent mail view
	// Keep below for now
	CString path = MboxMail::GetLastPath();

	if (!path.IsEmpty() && FileUtils::PathDirExists(path))
	{
		hFolder = HasFolder(path);
	}

	HTREEITEM hCurrentSelectedItem = m_tree.GetSelectedItem();
	if (hCurrentSelectedItem)
	{
		CString txt = m_tree.GetItemText(hCurrentSelectedItem);
		m_tree.SetItemState(hCurrentSelectedItem, 0, TVIS_BOLD);

		CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
		if (pFrame)
		{
			NListView* pListView = pFrame->GetListView();
			if (pListView)
			{
				pListView->m_path = L"";
				pListView->m_which = NULL;
				pListView->ResetSize();
				pListView->FillCtrl();
			}
		}
		int deb = 1;
	}
#endif

	UINT nCode = TVE_EXPAND;
	if (!expand)
		nCode = TVE_COLLAPSE;

	HTREEITEM hRoot = m_tree.GetRootItem();
	HTREEITEM hNext = hRoot;
	while (hNext)
	{
		if (hFolder != hNext)
		{
			ExpandItem(hNext, nCode);
		}
		hNext = m_tree.GetNextItem(hNext, TVGN_NEXT);
	}
	if (hFolder)
	{
		ExpandItem(hFolder, nCode);
		m_tree.EnsureVisible(hFolder);
	}
}

ArchiveFileInfoMap* NTreeView::SetupFileSizeMap(CString& path)
{
	GlobalFileSizeMap::iterator it;

	path.TrimRight(L"\\");
	std::wstring stdPath = (LPCWSTR)path;
	it = m_gFileSizes.find(stdPath);
	ArchiveFileInfoMap* item;
	if (it != m_gFileSizes.end())
	{
		item = it->second;
		fileSizes = &item->m_fileSizes;
		return item;
	}
	else
	{
		_ASSERTE(it == m_gFileSizes.end());
		MboxMail::assert_unexpected();
		return 0;
	}
}

int NTreeView::RemoveOrphanIndexFilesFromFolderAndFileSizeMap(CString& path, FileSizeMap& fileSizes)
{
	WIN32_FIND_DATA	wf;
	BOOL found;
	int count = 0;
	CString mboxPath;
	CString fn;
	CString mboxviewPath;

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	path.TrimRight(L"\\");

	CString datapath = MboxMail::GetLastDataPath();
	CString fw = datapath + L"\\*.mboxview";

	// Delete *.mboxview files without associated  valid mail archive files
	HANDLE f = FindFirstFile(fw, &wf);
	if (f != INVALID_HANDLE_VALUE)
	{
		do
		{
			if ((wf.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != FILE_ATTRIBUTE_DIRECTORY && wf.cFileName[0] != L'.')
			{
				fn = wf.cFileName;
				mboxviewPath = datapath + fn;

				FileUtils::SplitFilePath(fn, driveName, directory, fileNameBase, fileNameExtention);
				mboxPath = path + L"\\" + fileNameBase;

				ArchiveFileInfo info;
				found = fileSizes.Lookup(fileNameBase, info);
				if (found == FALSE)
				{
					BOOL ret = FileUtils::PathFileExist(mboxPath);
					if (ret == FALSE)
					{
						FileUtils::DelFile(mboxviewPath);
						count++;
					}
				}
			}
		} while (FindNextFile(f, &wf));
		FindClose(f);
	}
	return count;
}

HTREEITEM NTreeView::InsertAllDiscoveredMailFilesToTreeCtrl(HTREEITEM hParent, CString& path, FileSizeMap& fileSizes, BOOL unconditionalFolderInsert)
{
	CString root;
	path.TrimRight(L"\\");
	FileUtils::CPathStripPath(path, root);

	HTREEITEM hRoot = 0;
	if ((fileSizes.GetCount() > 0) || unconditionalFolderInsert)
	{
		CString parentFolderName;
		if ((hParent != TVI_ROOT) && (hParent != 0))
			parentFolderName = m_tree.GetItemText(hParent);

		hRoot = InsertTreeItem(root, 4, 4, hParent);

		if (hRoot == 0)
		{
			//ClearViewFile();
			fileSizes.RemoveAll();
			return hRoot;
		}

		CString itemName = m_tree.GetItemText(hRoot);

		if (hParent == TVI_ROOT)
		{
			CString registry_lastPath = path;
			registry_lastPath.TrimRight(L"\\");
			registry_lastPath.Append(L"\\");

			int index = m_folderArray.Add(registry_lastPath);
			m_tree.SetItemData(hRoot, index);
			DWORD_PTR retIndex = m_tree.GetItemData(hRoot);

			int nId = index;
			LabelInfo* linfo = new LabelInfo(nId, path, path); // m_nodeType = MailFolder
			m_labelInfoStore.Add(linfo, nId);
		}
		else
		{
			int nId = m_labelInfoStore.GetNextId();
			m_tree.SetItemData(hRoot, nId);
			DWORD_PTR retIndex = m_tree.GetItemData(hRoot);
			LabelInfo* linfo = new LabelInfo(nId, path, path, LabelInfo::MailSubFolder); // m_nodeType = MailSubFolder
			m_labelInfoStore.Add(linfo, nId);
		}

		ArchiveFileInfo info;
		POSITION pos = fileSizes.GetStartPosition();
		CString fn;
		while (pos)
		{
			fileSizes.GetNextAssoc(pos, fn, info);
			if (info.bShow)
			{
				HTREEITEM hItem = InsertTreeItem(fn, 8, 9, hRoot);
				if (hItem)
				{
					//m_tree.SetItemState(hItem, TVIS_BOLD, TVIS_BOLD);
					CString txt = m_tree.GetItemText(hItem);

					int nId = m_labelInfoStore.GetNextId();
					m_tree.SetItemData(hItem, nId);

					CString mboxFilePath = path + L"\\" + fn;
					LabelInfo* linfo = new LabelInfo(nId, mboxFilePath);  // MailFile
					m_labelInfoStore.Add(linfo, nId);
				}
			}
		}
		this->SortChildren(hRoot);
	}
	return hRoot;
}

// Iterate all files in the path folder, validate and insert into Mail Tree
// unless unconditionalFolderInsert==FALSE and folder empty
HTREEITEM NTreeView::LoadFileSizes(HTREEITEM hParent, CString& path, FileSizeMap& fileSizes, BOOL unconditionalFolderInsert)
{
	CString mboxFilePath;
	//CString registry_lastPath;
	CString root;
	path.TrimRight(L"\\");
	FileUtils::CPathStripPath(path, root);

	// Read first file index file if it exists from previous runs
	// and add to the filesSize hash table
	// In the next step, new archive files might be discovered and added 
	m_bIsDataDirty = FALSE;
	fileSizes.RemoveAll();
	CString datapath = MboxMail::GetLastDataPath();
	CString val;
	CString leftVal;
	CString rightVal;
	CStdioFile fp;
	CFileException ExError;
	CString viewFile = datapath + L".mboxview";
	//if (fp.Open(viewFile, CFile::modeRead | CFile::typeUnicode | CFile::typeText, &ExError))  // FIXMEFIXME
	if (fp.Open(viewFile, CFile::modeRead | CFile::typeUnicode, &ExError))  // FIXMEFIXME
	{
		CString line;
		wchar_t buff[3];
		fp.Read(buff, 2);
		while (fp.ReadString(line))
		{
			// TODO: Implement RemoveBOM()
			const unsigned char* p = (unsigned char*)((LPCWSTR)line);
			if ((*p == 0xFF) && (*(p + 1) == 0xFE))  // UTF16LE
			{
				CString str = line.Right(line.GetLength() - 1);
				line = str;
			}

			line.TrimRight(L"\r\n"); // FIXMEFIXME not critical to remove 

			int w = line.Find(L'\t', 0);
			if (w == -1)
				break;

			const wchar_t* l1 = (LPCWSTR)line;
			CString mboxFileName = line.Left(w);
			const wchar_t* wp0 = (LPCWSTR)mboxFileName;

			if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
			{
				CString fileName;
				FileUtils::GetFileName(CMainFrame::m_commandLineParms.m_mboxFileNameOrPath, fileName);
				if (mboxFileName.Compare(fileName))
					continue;
			}

			_int64	fSize = 0;
			_int64  bShow = 1;

			val = line.Mid(w + 1);  // get size, show and root values
			int pos = val.Find(L'\t', 0);
			if (pos < 0)  // old format; no mbox show value
			{
				m_bIsDataDirty = TRUE;
				fSize = _tstoi64(val);
			}
			else
			{
				leftVal = val.Left(pos);
				fSize = _tstoi64(leftVal);

				val = val.Mid(pos + 1);
				pos = val.Find(L'\t', 0);
				if (pos < 0)
				{
					bShow = _tstoi64(val);
					if ((bShow != 0) && (bShow != 1))
						bShow = 1;
				}
				else
				{
					m_bIsDataDirty = TRUE;
					leftVal = val.Left(pos);
					bShow = _tstoi64(leftVal);
					if ((bShow != 0) && (bShow != 1))
						bShow = 1;
				}
			}
			mboxFilePath = path + L"\\" + mboxFileName;
			const wchar_t* wp1 = (LPCWSTR)mboxFileName;
			const wchar_t* wp2 = (LPCWSTR)mboxFilePath;
			if (FileUtils::PathFileExist(mboxFilePath))
			{
				fileSizes[mboxFileName].fSize = fSize;
				fileSizes[mboxFileName].bShow = bShow;
				fileSizes[mboxFileName].bNeedsValidation = FALSE;
			}
			else
			{
				// fileSizes is empty initially; no need to do this
				if (fileSizes.PLookup(mboxFileName))
					fileSizes.RemoveKey(mboxFileName);
				m_bIsDataDirty = TRUE;
			}

		}
		fp.Close();
	}
	else
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not open \"" + viewFile;
		txt += L"\" file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);

		// TODO: verify implications
		fileSizes.RemoveAll();
	}

	//root = last_slash + 1;

	HTREEITEM hFolder = HasFolder(path);
	if (hFolder)  // TODO: should never be true, _ASSERTE ??
		int deb = 1;

	HTREEITEM hRoot = 0;

	CString fw = path + L"\\*.*";
	WIN32_FIND_DATA	wf;
	BOOL found;
	// Iterate all files in the lastPath folder
	// Valid new archives file are addedd  to fileSizes hash table
	// Orphan index file.mboxview files are removed  to force parsing to create new index file
	//
	HANDLE f = FindFirstFile(fw, &wf);
	if (f != INVALID_HANDLE_VALUE)
	{
		do
		{
			if ((wf.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != FILE_ATTRIBUTE_DIRECTORY && wf.cFileName[0] != '.')
			{
				CString fn = wf.cFileName;
				CString mboxFilePath = path + L"\\" + fn;
				if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
				{
					if (mboxFilePath.Compare(CMainFrame::m_commandLineParms.m_mboxFileNameOrPath))
						continue;
				}

				// Filter files based on file extension configuration
				if (!ImboxviewFileFilter(fn))
					continue;

				_int64 fSize = 0;
				ArchiveFileInfo info;
				CString errText;
				_int64 realFSize = FileUtils::FileSize(mboxFilePath, &errText);
				if (!errText.IsEmpty())
				{
					CString txt;
#if 0
					CString txt = errText;
					txt.Append(L"\n\"");
					txt.Append(mboxFilePath);
#endif
					if (mboxFilePath.GetLength() > _MAX_PATH)
					{
						CString l;
						l.Format(L"%d", mboxFilePath.GetLength());

						CString fmt = L"%s\n\"%s\"\n\nThe file path length (%s) is greater than maximum length (260) allowed by Windows API.";
						fmt += L" Please rename the file manually or moved to different folder before the next restart of MBox Viewer.\n\n";
						fmt += L"Note that MBox Viewer is always in discovery mode upon startup and this message will be shown until issue is resolved.\n\n";
						ResHelper::TranslateString(fmt);
						txt.Format(fmt, errText, mboxFilePath, l);
#if 0
						CString l;
						l.Format(L"%d", mboxFilePath.GetLength());
						txt.Append(L"\"\n\nThe file path length (");
						txt.Append(l);
						txt.Append(L") is greater than maximum length (260) allowed by Windows API.");
						txt += L" Please rename the file manually or moved to different folder before the next restart of MBox Viewer.\n\n";
						txt += L"Note that MBox Viewer is always in discovery mode upon startup and this message will be shown until issue is resolved.\n\n";
#endif
					}
					else
					{
						_ASSERTE(FALSE);
						// We should not be here
#if 0
						CString txt = errText;
						txt.Append(L"\n\"");
						txt.Append(mboxFilePath);

						txt.Append(L"\"\n\nUnexpected error detected while trying to determine the size of file.\n\n");
						txt += L"You may need to rename the file manually or moved to different folder before the next restart of MBox Viewer.\n\n";
						txt += L"Note that MBox Viewer is always in discovery mode upon startup and this message will be shown until issue is resolved.\n\n";
#endif
						CString fmt = L"%s\n\"%s\"\n\nUnexpected error detected while trying to determine the size of file.\n\n";
						fmt += L"You may need to rename the file manually or moved to different folder before the next restart of MBox Viewer.\n\n";
						fmt += L"Note that MBox Viewer is always in discovery mode upon startup and this message will be shown until issue is resolved.\n\n";
						ResHelper::TranslateString(fmt);
						txt.Format(fmt, errText, mboxFilePath);
					}

					int answer = MessageBox(txt, L"Warning", MB_APPLMODAL | MB_ICONWARNING | MB_OK);
				}

				found = fileSizes.Lookup(fn, info);
				if (found)
				{
					if (info.fSize != realFSize)
					{
						//TRACE("File=%s FileSize=%lld StoredFileSize=%lld\n", fn, realFSize, fSize);
						CString cache;
						BOOL retval = MboxMail::GetMboxviewFilePath(mboxFilePath, cache);

						FileUtils::DelFile(cache);
						// Delete List Files
						CString listFileName;
						int ret = NListView::DetermineListFileName(fn, listFileName);
						if (!listFileName.IsEmpty())
							FileUtils::DelFile(listFileName);

						fileSizes[fn].fSize = realFSize;
						fileSizes[fn].bNeedsValidation = TRUE;
						m_bIsDataDirty = TRUE;
						int deb = 1;
					}
				}

				else if (NTreeView::ImboxviewFile(mboxFilePath))
				{
#if 0
					// No longer needed. Folder names are now created by keeping extension. Howver. dot is replaced with - char
					// That may still create an issue but it is very unlikely. Didn't want to create folder names that looks like mbox file type
					// May pay price for this -:)
					CString fileNameExtention;
					CString fileNameBase;
					FileUtils::GetFileBaseNameAndExtension(fn, fileNameBase, fileNameExtention);

					wchar_t c = fileNameBase.GetAt(fileNameBase.GetLength() - 1);
					if ((c == ' ') || (c == '\t'))
					{
						CString txt = L"Invalid mail file name \n\n\"" + fn;
						txt += L"\".\n\nBase name of the file (name without the extension) can't have trailing white spaces. Ignoring.\n";
						int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
						continue;
					}
#endif

					fileSizes[fn].fSize = realFSize;
					fileSizes[fn].bNeedsValidation = TRUE;
					m_bIsDataDirty = TRUE;
					int deb = 1;
				}
				else
				{
					int deb = 1; // ignore non mbox files
				}

				if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
				{
					break;
				}
			}

		} while (FindNextFile(f, &wf));
		FindClose(f);

		hRoot = InsertAllDiscoveredMailFilesToTreeCtrl(hParent, path, fileSizes, unconditionalFolderInsert);

		int deletedIndexFileCnt = NTreeView::RemoveOrphanIndexFilesFromFolderAndFileSizeMap(path, fileSizes);
	}
	if (m_bIsDataDirty)
	{
		if (hRoot)
		{
			CString itemName = m_tree.GetItemText(hRoot);
			SaveData(hRoot);
		}
		m_bIsDataDirty = FALSE;
	}
	return hRoot;
}

// Called on Startup, File open and All File Refresh
// Items are inserted into CTree only by this function
HTREEITEM NTreeView::FillCtrl(HTREEITEM hParent, BOOL selectFolder, BOOL unconditionalFolderInsert, BOOL expand)
{
	HTREEITEM hRoot = NTreeView::FillCtrl_Internal(hParent, unconditionalFolderInsert, expand);
	if (hRoot == 0)
	{
		int deb = 1;
	}

	if (hRoot && selectFolder)
	{
		HTREEITEM hCurrentSelectedItem = m_tree.GetSelectedItem();
		if (hCurrentSelectedItem > 0)
		{
			CString itemName = m_tree.GetItemText(hCurrentSelectedItem);
			m_tree.SetItemState(hCurrentSelectedItem, 0, TVIS_BOLD);
		}

		UINT nCode = TVGN_CARET;
		//BOOL retval = m_tree.Select(hRoot, nCode);
		BOOL retval = SelectTreeItem(hRoot);
}
	return hRoot;
	}

HTREEITEM NTreeView::FillCtrl_Internal(HTREEITEM hParent, BOOL unconditionalFolderInsert, BOOL expand)
{
	// called from 
	// void CMainFrame::OnFileOpen()
	// void CMainFrame::DoOpen(CString& fpath)
	// void NTreeView::LoadFolders()
	// BOOL NTreeView::DeleteFolder(HTREEITEM hItem)
	//
	//

	HTREEITEM hItem = 0;

	if (m_bInFillControl)
		return hItem;

	m_bInFillControl = TRUE;

	CString path = MboxMail::GetLastPath();

	//m_tree.DeleteAllItems();
	if (path.IsEmpty() || !FileUtils::PathDirExists(path))
	{
		m_bInFillControl = FALSE;
		return hItem;
	}

	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	CString fpath = path;
	fpath.TrimRight(L"\\");
	fpath.Append(L"\\");

	FileUtils::SplitFilePath(fpath, driveName, directory, fileNameBase, fileNameExtention);
	if (directory.GetLength() <= 1)
	{
		m_bInFillControl = FALSE;
		return hItem;
	}

	// traverse the Tree to see if the folder already exist
	CString itemName;
	HTREEITEM hFolder;
	//if (hFolder = HasFolder(path))
	if (hFolder = IsFolderOpen(path))
	{
		// This will happen when user tries to select folder which already is inserted into Tree
		HTREEITEM hCurrentSelectedItem = m_tree.GetSelectedItem();
		if (hCurrentSelectedItem)
			itemName = m_tree.GetItemText(hCurrentSelectedItem);
		if (hCurrentSelectedItem != hFolder)
		{
			if (hCurrentSelectedItem > 0)
				m_tree.SetItemState(hCurrentSelectedItem, 0, TVIS_BOLD);

			UINT nCode = TVGN_CARET;
			if (hFolder)
			{
				//BOOL retval = m_tree.Select(hFolder, nCode);
				BOOL retval = SelectTreeItem(hFolder);
			}
		}

		if (expand)
			ExpandItem(hFolder, TVE_EXPAND);
		m_bInFillControl = FALSE;
		if (hFolder == 0)
			int deb = 1;
		return hFolder;
	}

	CString tmppath = path;
	path.TrimRight(L"\\");
	std::wstring stdPath = (LPCWSTR)path;
	GlobalFileSizeMap::iterator it = m_gFileSizes.find(stdPath);
	ArchiveFileInfoMap* infoMap;
	if (it == m_gFileSizes.end())
	{
		infoMap = new ArchiveFileInfoMap(path);
		m_gFileSizes.insert(GlobalFileSizeMap::value_type(stdPath, infoMap));
	}
	else
	{
		infoMap = it->second;
		int deb = 1;
	}

	// TODO: globals
	ArchiveFileInfoMap* fileSizesMap = SetupFileSizeMap(path);
	m_fileSizesMap = fileSizesMap;

	if (fileSizesMap == 0)
		int deb = 1;

	hItem = LoadFileSizes(hParent, path, *fileSizes, unconditionalFolderInsert);

	int count = (int)fileSizes->GetCount();
	int deb = 1;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		NListView* pListView = pFrame->GetListView();
		if (pListView)
		{
			pListView->m_path = L"";
			pListView->m_which = NULL;
			// pListView->ResetSize();  // check if needed
			pListView->FillCtrl();
			pListView->CloseMailFile();
		}
	}

	m_tree.SetRedraw();

	m_bInFillControl = FALSE;
	return hItem;
}

void NTreeView::OnSelchanged(NMHDR* pNMHDR, LRESULT* pResult)
{
	CString oldItemName;
	CString newItemName;

	__int64 res = *pResult;
	*pResult = 0;

	NMTREEVIEW* pNm = (LPNMTREEVIEW)pNMHDR;

	HTREEITEM hNewItem = pNm->itemNew.hItem;
	HTREEITEM hOldItem = pNm->itemOld.hItem;

	if (hOldItem)
		oldItemName = m_tree.GetItemText(hOldItem);

	if (hNewItem)
		newItemName = m_tree.GetItemText(hNewItem);

	if (hOldItem == hNewItem)
		return;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame == NULL)
		return;
	NListView* pListView = pFrame->GetListView();
	if (!pListView)
		return;

	NMsgView* pMsgView = pFrame->GetMsgView();
	if (pMsgView)
		pMsgView->DisableMailHeader();

	pListView->CloseMailFile();

	// 	TODO: find better place for below
	pListView->m_path = L"";
	pListView->m_path_label = L"";
	pListView->m_lastSel = -1;

	if (!hNewItem)
	{
		pListView->m_path_label = L"";
		pListView->m_path = L"";
		pListView->m_which = NULL;
		pListView->FillCtrl();

		CString lPath = MboxMail::GetLastPath();
		if (lPath.IsEmpty())
			MboxMail::assert_unexpected();
		return;
	}

	LabelInfo* linfoOld = 0;
	LabelInfo* linfoNew = 0;
	DWORD nIdOld = -1;
	DWORD nIdNew = -1;

	if (hOldItem)
	{
		nIdOld = (DWORD)m_tree.GetItemData(hOldItem);
		linfoOld = m_labelInfoStore.Find(nIdOld);
	}

	nIdNew = (DWORD)m_tree.GetItemData(hNewItem);
	linfoNew = m_labelInfoStore.Find(nIdNew);

	// This is called when tree is closed

	HTREEITEM hOldItemRoot = DetermineRootMboxFile(hOldItem);
	CString oldItemRootName;
	if (hOldItemRoot)
		oldItemRootName = m_tree.GetItemText(hOldItemRoot);

	HTREEITEM hRoot = DetermineRootMboxFile(hNewItem);
	CString rootName;
	if (hRoot)
		rootName = m_tree.GetItemText(hRoot);

	HTREEITEM hParent = m_tree.GetParentItem(hNewItem);
	CString parentName;
	if (hParent)
		parentName = m_tree.GetItemText(hParent);

	HTREEITEM hOldItemParent = 0;
	if (hOldItem)
		hOldItemParent = m_tree.GetParentItem(hOldItem);
	CString oldItemParentName;
	if (hOldItemParent)
		oldItemParentName = m_tree.GetItemText(hOldItemParent);

	//######################################################################################################
	//     Process Folder or SubFolder 
	//######################################################################################################
	if ((linfoNew->m_nodeType == LabelInfo::MailFolder) || (linfoNew->m_nodeType == LabelInfo::MailSubFolder))
	{
		// New item is folder or sub-folder. Need to load corresponding index
		m_labelView = FALSE;

		if (linfoNew->m_nodeType == LabelInfo::MailFolder)
		{
			if (hParent != 0)
				MboxMail::assert_unexpected();
		}
		else // if (linfoNew->m_nodeType == LabelInfo::MailSubFolder)
		{
			if (hParent == 0)
				MboxMail::assert_unexpected();
		}

		if (hOldItem > 0)
			m_tree.SetItemState(hOldItem, 0, TVIS_BOLD);
		m_tree.SetItemState(hNewItem, TVIS_BOLD, TVIS_BOLD);

		CString folderPath;
		DetermineFolderPath(hNewItem, folderPath);
		//folderPath = linfoNew->m_mailFolderPath;

		MboxMail::SetLastPath(folderPath);

		ArchiveFileInfoMap* fileSizesMap = SetupFileSizeMap(folderPath);
		m_fileSizesMap = fileSizesMap;

		int count = (int)fileSizes->GetCount();

		if (fileSizesMap == 0)
			int deb = 1;

		pListView->m_path = L"";
		pListView->m_which = NULL;
		pListView->ResetSize();
		pListView->FillCtrl();

		m_labelView = FALSE;

		CString lPath = MboxMail::GetLastPath();
		if (lPath.IsEmpty())
			MboxMail::assert_unexpected();

		TRACE(L"OnSelchanged_Folder: AssertNTreeView\n");
		BOOL retAssert = AssertNTreeView();

		if (pListView)
		{
			LRESULT lres = pListView->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
		}
		return;
	}  // end of if ((linfoNew->m_nodeType == LabelInfo::MailFolder) || (linfoNew->m_nodeType == LabelInfo::MailSubFolder))

	//CString parentName = m_tree.GetItemText(hParent);
	NamePatternParams* pNamePP = &pFrame->m_NamePatternParams;

	//######################################################################################################
	//     Process Label 
	//######################################################################################################
	if (linfoNew && (linfoNew->m_nodeType == LabelInfo::MailLabel)) // label
	{
		// New item is Label.  May need to load corresponding index file depending 
		// whether new label correspponds to the same mail file or not

		if (hParent == 0)
			return;

		// TODO: set here or see below
		pListView->m_which = hRoot;
		//pListView->m_which = hParent;

		if (nIdNew < 100000)
			int deb = 1; // _ASSERTE ??

		// Set global variables - fix it -:)
		CString mboxFileNamePath = MboxMail::s_path;
		CString lastPath = MboxMail::GetLastPath();
		CString lastDataPath = MboxMail::GetLastDataPath();

		if (linfoNew->m_listFilePath.IsEmpty())
		{
			_ASSERTE(FALSE);  // explain
			// TODO: Should this always be done at the entry of this function ?
			if (hOldItem > 0)
			{
				CString text = m_tree.GetItemText(hOldItem);
				m_tree.SetItemState(hOldItem, 0, TVIS_BOLD);
			}
			m_tree.SetItemState(hNewItem, TVIS_BOLD, TVIS_BOLD);

			CString lPath = MboxMail::GetLastPath();
			if (lPath.IsEmpty())
				MboxMail::assert_unexpected();

			if (pListView)
			{
				LRESULT lres = pListView->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
			}
			return;
		}

		pListView->m_path = linfoNew->m_filePath;
		pListView->m_path_label = linfoNew->m_listFilePath;
		MboxMail::s_path_label = linfoNew->m_listFilePath;

		int paneId = 0;
		CString sText;
		CString fmt = L"Opening Label %s ...";
		ResHelper::TranslateString(fmt);

		sText.Format(fmt, linfoNew->m_label);
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

		CString driveName;
		CString mboxFileDirectory;
		CString mboxFileNameBase;
		CString mboxFileNameExtention;

		FileUtils::SplitFilePath(linfoNew->m_filePath, driveName, mboxFileDirectory, mboxFileNameBase, mboxFileNameExtention);
		CString newLastPath = driveName + mboxFileDirectory;

		CString mboxFileName = mboxFileNameBase + mboxFileNameExtention;

		if ((hOldItemRoot != hRoot) || mboxFileNamePath.Compare(linfoNew->m_filePath))
		{
			// New and old labels correspond to different mail files. Load new index file
			// 
			// or set here. Likely it doesn't matter
			// pListView->m_which = hRoot;
			MboxMail::SetLastPath(newLastPath);
			CString lastPath = MboxMail::GetLastPath();
			CString lastDataPath = MboxMail::GetLastDataPath();

			pListView->m_path = linfoNew->m_filePath;
			//MboxMail::s_path = linfoNew->m_filePath;
			MboxMail::SetMboxFilePath(linfoNew->m_filePath);

			CString errorText;
			int rval = pListView->MailFileFillCtrl(errorText);
			MboxMail::nWhichMailList = IDC_ARCHIVE_LIST;

			pListView->ResetFilterDates();
		}

		lastPath = MboxMail::GetLastPath();
		lastDataPath = MboxMail::GetLastDataPath();

		//MboxMail::nWhichMailList = IDC_FOLDER_LIST;
		int retval = pListView->LoadLabelListFile_v2(linfoNew->m_listFilePath, linfoNew->m_label, linfoNew->m_filePath);
		if (retval < 0)
		{
#if 1
			int retR = NTreeView::RefreshLabelsForSingleMailFile(hParent);  // TODO!!!: verify, hParent should point to mail file and to direct parent whicj could be a label
#else
			HTREEITEM hFolder = FindFolder(hRoot, lastPath);
			if (hFolder)
			{
				HTREEITEM hMboxFolder = FindItem(hFolder, mboxFileName);

				int retR = NTreeView::RefreshLabelsForSingleMailFile(hMboxFolder);
			}
#endif
			if (hOldItem > 0)
			{
				CString text = m_tree.GetItemText(hOldItem);
				m_tree.SetItemState(hOldItem, 0, TVIS_BOLD);
			}
			m_tree.SetItemState(hNewItem, TVIS_BOLD, TVIS_BOLD);

			CString lPath = MboxMail::GetLastPath();
			if (lPath.IsEmpty())
				MboxMail::assert_unexpected();

			if (pListView)
			{
				LRESULT lres = pListView->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
			}
			return;
		}

		if (MboxMail::s_mails.GetCount() > 0)
		{
			pListView->SelectItemFound(0);
		}

		if (hOldItem > 0)
		{
			CString text = m_tree.GetItemText(hOldItem);
			m_tree.SetItemState(hOldItem, 0, TVIS_BOLD);
		}
		m_tree.SetItemState(hNewItem, TVIS_BOLD, TVIS_BOLD);

		m_labelView = TRUE;

		int nIdMailList = IDC_ARCHIVE_LIST;
		pFrame->SetMailList(nIdMailList);

		MboxMail::s_mails_find.SetSizeKeepData(0);
		MboxMail::m_findMails.m_lastSel = -1;
		MboxMail::m_findMails.b_mails_which_sorted = 1;

		sText.Format(L"Ready");
		ResHelper::TranslateString(sText);
		pFrame->SetStatusBarPaneText(paneId, sText, FALSE);

		pNamePP = &pFrame->m_NamePatternParams;

		pFrame->UpdateFilePrintconfig();

		pNamePP = &pFrame->m_NamePatternParams;

		int deb = 1;

		CString lPath = MboxMail::GetLastPath();
		if (lPath.IsEmpty())
			MboxMail::assert_unexpected();

		TRACE(L"OnSelchanged_Label: AssertNTreeView\n");
		BOOL retAssert = AssertNTreeView();

		if (pListView)
		{
			LRESULT lres = pListView->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
		}
		return;
	}  // end of if (linfoNew && (linfoNew->m_nodeType == LabelInfo::MailLabel)) // label

	//######################################################################################################
	//     Process Mail Archive 
	//######################################################################################################

#if 0
	int nId = IDC_ARCHIVE_LIST;
	BOOL enable = TRUE;
	pFrame->EnableMailList(nId, enable);
#endif

	m_labelView = FALSE;

	if (linfoOld && linfoNew)
	{
		if (linfoNew->m_filePath.Compare(linfoOld->m_filePath) == 0)
		{
			m_labelView = FALSE;
			pListView->m_path = linfoNew->m_filePath;

			pFrame->OnBnClickedArchiveList();

			if (hOldItem > 0)
			{
				CString text = m_tree.GetItemText(hOldItem);
				m_tree.SetItemState(hOldItem, 0, TVIS_BOLD);
			}
			m_tree.SetItemState(hNewItem, TVIS_BOLD, TVIS_BOLD);

			MboxMail::s_mails_find.SetSizeKeepData(0);
			MboxMail::m_findMails.m_lastSel = -1;
			MboxMail::m_findMails.b_mails_which_sorted = 1;

			CString lPath = MboxMail::GetLastPath();
			if (lPath.IsEmpty())
				MboxMail::assert_unexpected();

			if (pListView)
			{
				LRESULT lres = pListView->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
			}
			return;
		}
	}

	CString str = m_tree.GetItemText(hNewItem);

	CString folderPath;

	if (linfoNew && (linfoNew->m_nodeType == LabelInfo::MailFile))
	{
		FileUtils::GetFolderPath(linfoNew->m_filePath, folderPath);
	}
	else
		MboxMail::assert_unexpected();

	MboxMail::SetLastPath(folderPath);
	CString path = MboxMail::GetLastPath();

	if (str.IsEmpty() || path.IsEmpty())
		return;

	ArchiveFileInfoMap* fileSizesMap = SetupFileSizeMap(path);
	m_fileSizesMap = fileSizesMap;

	int count = (int)fileSizes->GetCount();

	if (fileSizesMap == 0)
		int deb = 1;

	path.TrimRight(L"\\");
	pListView->m_path = path + L'\\' + str;

	if (hOldItem > 0)
	{
		CString text = m_tree.GetItemText(hOldItem);
		m_tree.SetItemState(hOldItem, 0, TVIS_BOLD);
	}
	m_tree.SetItemState(hNewItem, TVIS_BOLD, TVIS_BOLD);


	// TODO: relying on m_which is risky. 
	// Consider to eliminate and rely on strings and searches, assuming minimum impact on performance
	pListView->m_which = hNewItem;

	pListView->ResetSize();

	pListView->ResetFilterDates();

	int paneId = 0;
	CString sText;
	CString fmt = L"Opening %s ...";
	ResHelper::TranslateString(fmt);
	sText.Format(fmt, str);
	pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	pListView->FillCtrl();
	MboxMail::nWhichMailList = IDC_ARCHIVE_LIST;

	if (MboxMail::s_mails.GetCount() > 0)
	{
		pListView->SelectItemFound(0);
	}

	pFrame->UpdateFilePrintconfig();

	sText = L"Ready";
	ResHelper::TranslateString(sText);
	pFrame->SetStatusBarPaneText(paneId, sText, FALSE);

	MboxMail::ShowHint(HintConfig::MsgWindowPlacementHint, GetSafeHwnd());


	// Just testing and collecting stats
	//pListView->ScanAllMailsInMbox();
	//pListView->ScanAllMailsInMbox_NewParser();

	m_labelView = FALSE;

	CString lPath = MboxMail::GetLastPath();
	if (lPath.IsEmpty())
		MboxMail::assert_unexpected();
	ShowMemStatus();

	TRACE(L"OnSelchanged_MailArchive: AssertNTreeView\n");
	BOOL retAssert = AssertNTreeView();

	if (pListView)
	{
		LRESULT lres = pListView->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
	}
}

void NTreeView::ForceParseMailFile(HTREEITEM hItem)
{
	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame == NULL)
		return;
	NListView* pListView = pFrame->GetListView();
	if (!pListView)
		return;

	pListView->CloseMailFile();

	if (!hItem) {
		pListView->m_path = L"";
		pListView->m_which = NULL;
		pListView->FillCtrl();
		return;
	}
	HTREEITEM hRoot = DetermineRootItem(hItem);
	if (hItem == hRoot) {
		pListView->m_path = L"";
		pListView->m_which = NULL;
		pListView->ResetSize();
		pListView->FillCtrl();
		return;
	}
	CString str = m_tree.GetItemText(hItem);
	CString path = MboxMail::GetLastPath();
	if (str.IsEmpty() || path.IsEmpty())
		return;

	path.TrimRight(L"\\");
	pListView->m_path = path + L'\\' + str;

	CString cache;
	BOOL retval = MboxMail::GetMboxviewFilePath(pListView->m_path, cache);

	FileUtils::DelFile(cache, FALSE);
	// TODO: Should delete more, i.e reletaed data from al Caches

	// Labels are not removed from screen immedietally anyway. Commented out
	int rvalDelete = DeleteLabelsForSingleMailFile(hItem);

	pListView->m_which = hItem;
	pListView->ResetSize();
	pListView->FillCtrl();
	MboxMail::nWhichMailList = IDC_ARCHIVE_LIST;

	// TODO: need to delete Labels and recreate . Done ?
	int rvalRefresh = RefreshLabelsForSingleMailFile(hItem);
}

int NTreeView::FindSelectedItemIndex(MailArray& mails, int referenceIndex, BOOL force)
{
	if ((referenceIndex < 0) || ((referenceIndex >= mails.GetCount()) && (force == FALSE)))
		return -1;

	int i;
	for (i = 0; i < mails.GetCount(); i++)
	{
		if (mails[i]->m_index == referenceIndex)
			return i;
	}
	return -1;
}

HTREEITEM NTreeView::FindLabelItemByLabelPath(HTREEITEM hRoot, CString& path)
{
	HTREEITEM hFolder = 0;
	HTREEITEM hLabel = 0;
	BOOL retval = FALSE;
	CString labelFilePath;
	CString normalizedPath = path;
	normalizedPath.TrimRight(L"\\");

	if (hRoot == 0)
		return 0;

	int index = (DWORD)m_tree.GetItemData(hRoot);
	CString iteName = m_tree.GetItemText(hRoot);
	LabelInfo* linfo = m_labelInfoStore.Find(index);
	if (linfo == 0)
		return 0;

	if (linfo->m_nodeType == LabelInfo::MailLabel)
	{
		labelFilePath = linfo->m_listFilePath;
		if (normalizedPath.CompareNoCase(labelFilePath) == 0)
		{
			hLabel = hRoot;
			return hLabel;
		}
	}
	if (m_tree.ItemHasChildren(hRoot))
	{
		HTREEITEM child = m_tree.GetChildItem(hRoot);
		while (child != NULL)
		{
			CString iteName = m_tree.GetItemText(child);
			hLabel = FindLabelItemByLabelPath(child, path);
			if (hLabel)
			{
				break;
			}
			child = m_tree.GetNextItem(child, TVGN_NEXT);
		}
	}

	return hLabel;
}


void NTreeView::OpenLastSelection(MailSelectionInfo* mailSelection)
{
	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	NListView* pListView = 0;
	if (pFrame == 0)
		return;

	pListView = pFrame->GetListView();

	MailSelectionInfo  mSelection;
	if (mailSelection == 0)
	{
		mailSelection = &mSelection;

		CString section_lastSelection = CString(sz_Software_mboxview) + L"\\LastSelection";

		mailSelection->m_lastPath = pFrame->m_lastPath;
		mailSelection->m_lastMailFilePath = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_lastSelection, L"lastMailFilePath");
		mailSelection->m_lastLabelFilePath = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_lastSelection, L"lastLabelFilePath");
		mailSelection->m_lastMailIndex = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_lastSelection, L"lastMailIndex");
		mailSelection->m_lastWhichSort = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_lastSelection, L"lastWhichSort");
	}

	HTREEITEM hLastPath = FindFolder(mailSelection->m_lastPath);
	HTREEITEM hLastMailFilePath = FindMailFile(mailSelection->m_lastMailFilePath);
	HTREEITEM hLastLabelFilePath = FindLabelFile(mailSelection->m_lastMailFilePath, mailSelection->m_lastLabelFilePath);

	HTREEITEM hItem = hLastPath;
	if (hLastLabelFilePath)
		hItem = hLastLabelFilePath;
	else if (hLastMailFilePath)
		hItem = hLastMailFilePath;

	if (hItem == 0)
		return;

	// This will happen when user tries to select folder which already is inserted into Tree
	HTREEITEM hCurrentSelectedItem = m_tree.GetSelectedItem();
	if (hCurrentSelectedItem)
		CString itemName = m_tree.GetItemText(hCurrentSelectedItem);

	if (hCurrentSelectedItem != hItem)
	{
		if (hCurrentSelectedItem != 0)
			m_tree.SetItemState(hCurrentSelectedItem, 0, TVIS_BOLD);

		BOOL retval = m_tree.SelectItem(hItem);
	}

	if (pListView)
	{
		int col = 1;
		int whichSorted = pListView->MailsWhichColumnSorted();
		if (whichSorted != mailSelection->m_lastWhichSort)
		{
			MboxMail::b_mails_which_sorted = -mailSelection->m_lastWhichSort;
			col = abs(mailSelection->m_lastWhichSort);
			if (col == 99)
				col = 0;
			pListView->SortByColumn(col);
		}
	}

	int selIndex = FindSelectedItemIndex(MboxMail::s_mails, mailSelection->m_lastMailIndex, !mailSelection->m_lastLabelFilePath.IsEmpty());
	if (selIndex < 0)
		selIndex = 0;

	if (pListView && hLastMailFilePath)
	{
		pListView->SelectItemFound(selIndex);
	}

	if (hLastPath)
		ExpandItem(hLastPath, TVE_EXPAND);

	if (hLastMailFilePath)
		ExpandItem(hLastMailFilePath, TVE_EXPAND);

	// TODO: is there a simpler way ?
	PostMessage(WM_CMD_PARAM_RESET_TREE_POS_MESSAGE, 0, 0);


	TRACE(L"PrintFolderNames NEW\n");
	PrintFolderNames(0, TRUE);

#ifdef  _DEBUG
	CString pathTest = "F:\\Thunderbird\\Project.sbd\\Sub-project-1.sbd";
	//pathTest = "F:\\Thunderbird\\";
	//pathTest = "fake";

	TRACE(L"FindFolder OLD\n");
	HTREEITEM hFolder = m_tree.GetRootItem();
	hFolder = FindFolder(hFolder, pathTest);
	CString folderName;
	if (hFolder)
		folderName = m_tree.GetItemText(hFolder);


	TRACE(L"FindFolder NEW\n");
	hFolder = 0;
	BOOL retF = FindFolder(0, pathTest, hFolder, FALSE);

	if (hFolder)
		folderName = m_tree.GetItemText(hFolder);
#endif

	if (pListView)
	{
		LRESULT lres = pListView->PostMessage(WM_CMD_PARAM_ON_SWITCH_WINDOW_MESSAGE, 0, 0);
		int deb = 0;
	}

	int deb = 1;
}

// Called when user specified command line option MAIL_FILE=
// TODO: 
void NTreeView::SelectMailFile(CString* fileNm)
{
	// mailFile can be mailFilePath or mailFileName
	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	CString mailFile;
	if (fileNm)
	{
		mailFile.Append(*fileNm);
		//delete fileNm;
	}
	else
		mailFile = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"mailFile");

	CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_general, L"mailFile", CString(""));

	if (mailFile.IsEmpty())
		return;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame == NULL)
		return;

	NListView* pListView = pFrame->GetListView();
	if (!pListView)
		return;

	CString mboxFilePath;
	CString mboxFileName;
	FileUtils::GetFolderPathAndFileName(mailFile, mboxFilePath, mboxFileName);

	mboxFilePath.TrimRight(L"\\");
	mboxFilePath.Append(L"\\");

	CString txt;
	if (!FileUtils::PathFileExist(mailFile))
	{
		if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
		{
#if 0
			txt = L"Nonexistent Mail File \"" + mailFile;
			txt += L"\".\nTerminating";
#endif

			CString fmt = L"Nonexistent Mail File \"%s\".\nTerminating";
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, mailFile);

			int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
			AfxGetMainWnd()->PostMessage(WM_CLOSE);
		}
		else
		{
#if 0
			txt = L"Nonexistent Mail File \"" + mailFile;
			txt += L"\".\nDo you want to continue?";
#endif
			CString fmt = L"Nonexistent Mail File \"%s\".\nDo you want to continue?";
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, mailFile);
			int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
			if (answer == IDNO)
				AfxGetMainWnd()->PostMessage(WM_CLOSE);
		}
	}
	else
	{
		pListView->m_path = mailFile;

		pListView->CloseMailFile();

		MboxMail::SetLastPath(mboxFilePath);

		HTREEITEM hFolder = HasFolder(mboxFilePath);
		if (hFolder)
		{
			HTREEITEM hItem = NTreeView::FindItem(hFolder, mboxFileName);
			UINT nCode = TVGN_CARET;
			if (hItem)
			{
				//BOOL retval = m_tree.Select(hItem, nCode);
				BOOL retval = SelectTreeItem(hItem);

				if (MboxMail::s_mails.GetCount() > 0)
				{
					pListView->SelectItemFound(0);
					pListView->SetListFocus();
				}
			}
		}
	}
}

// TODO: Works but need to review  all globals  to reduce complexity,
// make reliable and maintainable
// Called when 
void NTreeView::InsertMailFile(CString& mailFile)
{
	// mailFile can be mailFilePath or mailFileName
	if (mailFile.IsEmpty())
		return;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame == NULL)
		return;
	NListView* pListView = pFrame->GetListView();
	if (!pListView)
		return;

	CString fileName;
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;
	FileUtils::SplitFilePath(mailFile, driveName, directory, fileNameBase, fileNameExtention);

	CString mailFileName = fileNameBase + fileNameExtention;

	CString mailFileFolderpath;
	CString path = MboxMail::GetLastPath();
	if (directory.GetLength() > 0)
	{
		mailFileFolderpath = driveName + directory;
		if (mailFileFolderpath.Compare(path))
			MboxMail::assert_unexpected();
		path = mailFileFolderpath;
	}

	if (path.IsEmpty())
		return;

	path.TrimRight(L"\\");
	pListView->m_path = path + L'\\' + mailFileName;

	CString s_path = MboxMail::s_path;
	if (s_path.Compare(pListView->m_path))
	{
		MboxMail::assert_unexpected();
	}

	MboxMail::SetMboxFilePath(pListView->m_path);

	CString txt;
	if (!FileUtils::PathDirExists(path))
		txt = L"Nonexistent Directory \"" + path;
	else if (!FileUtils::PathFileExist(pListView->m_path))
		txt = L"Nonexistent File \"" + pListView->m_path;

	if (!txt.IsEmpty())
	{
		// TODO:
		; //  return;
	}
	else
	{
		path.Append(L"\\");
		hashsum_t hashsum = GlobalFolderInfoDB::GetHashsum(&path);
		FolderInfo* finfo = m_globalFolderInfoDB.Find(&path);
		if (finfo == 0)
		{
			MboxMail::assert_unexpected();
		}

		HTREEITEM hFolder = FindFolder(path);
		if (hFolder)
		{
			HTREEITEM hItem = NTreeView::FindItem(hFolder, mailFileName);
			UINT nCode = TVGN_CARET;
			if (hItem == 0)
			{
				FileSizeMap* sv_fileSizes = fileSizes;

				ArchiveFileInfoMap* fileSizesMap = SetupFileSizeMap(path);
				m_fileSizesMap = fileSizesMap;

				(*fileSizes)[mailFileName].bShow = 1;
				_int64 fSize = FileUtils::FileSize(pListView->m_path);
				(*fileSizes)[mailFileName].fSize = fSize;
				SaveData(hFolder);

				CString mboxIndexFilepath; // = pListView->m_path + ".mboxview";
				BOOL retval = MboxMail::GetMboxviewFilePath(pListView->m_path, mboxIndexFilepath);
				BOOL ret = FileUtils::DelFile(mboxIndexFilepath);


				HTREEITEM hItem = InsertTreeItem(mailFileName, 8, 9, hFolder);
				if (hItem)
				{
					DeleteLabelsForSingleMailFile(hItem);

					int nId = m_labelInfoStore.GetNextId();
					m_tree.SetItemData(hItem, nId);

					LabelInfo* linfo = new LabelInfo(nId, mailFile);  // MailFile
					m_labelInfoStore.Add(linfo, nId);

					BOOL retval = SelectTreeItem(hItem);

					this->SortChildren(hFolder);
				}
				int deb = 1;
			}
			else
			{
				CString mboxIndexFilepath; // = pListView->m_path + ".mboxview";
				BOOL retval = MboxMail::GetMboxviewFilePath(pListView->m_path, mboxIndexFilepath);
				retval = FileUtils::DelFile(mboxIndexFilepath);
				DeleteLabelsForSingleMailFile(hItem);
				HTREEITEM hParent = m_tree.GetParentItem(hItem);
				if (hParent)
				{
					pListView->m_path = L"";
					MboxMail::SetMboxFilePath(pListView->m_path, TRUE);
					retval = SelectTreeItem(hParent);
				}
				else
					retval = SelectTreeItem(hItem);
			}
		}
	}
	TRACE(L"InsertMailFile: AssertNTreeView\n");
	BOOL retval = AssertNTreeView();
}

BOOL NTreeView::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt)
{
	CWnd* wnd = WindowFromPoint(pt);
	if (wnd == NULL)
		return FALSE;
	if (wnd == this) {
		BOOL ret = CWnd::OnMouseWheel(nFlags, zDelta, pt);
		return TRUE;
	}
	if ((GetKeyState(VK_CONTROL) & 0x80) == 0) { // if CTRL key not Down; Do we need to post msg further anyway
		// Commented out, it freezes mbox viewer and and even IE for few seconds when CTRL/SHIFT/etc key are kept down
		; // return wnd->PostMessage(WM_MOUSEWHEEL, MAKELPARAM(nFlags, zDelta), MAKELPARAM(pt.x, pt.y));
	}
	return TRUE;
}

// Write to .mboxview file
void NTreeView::Traverse(HTREEITEM hItem, CFile& fp, FileSizeMap& fileSizes)
{
	CString line;
	CString path;

	ArchiveFileInfo info;
	POSITION pos = fileSizes.GetStartPosition();

	if (pos)
	{
		const char* BOM_UTF16_LI = "\xFF\xFE";
		fp.Write(BOM_UTF16_LI, 2);
	}
	else
		return;

	while (pos)
	{
		fileSizes.GetNextAssoc(pos, path, info);

		// or for better performance; not really crusial here
		// CMap<CString, LPCSTR, ArchiveFileInfo, ArchiveFileInfo>::CPair *infopair = fileSizes.PLookup(path);
		// _int64 fSize = infopair->value.fSize;
		// _int64 bShow = infopair->value.bShow;

		_int64 fSize = fileSizes[path].fSize;
		_int64 bShow = fileSizes[path].bShow;

		line.Format(L"%s\t%lld\t%lld\n", path, fSize, bShow);
		LPCWSTR data = (LPCWSTR)line;
		int datalen = line.GetLength() * 2;
		fp.Write(data, datalen);
	}
}


// Thhis function creates or updates per folder .mboxview index file
// It is called when selcting new mail archive file or by RedrawMails called after sorting
// or touching mail list
void NTreeView::SaveData(HTREEITEM hItem)
{
	CFile fp;
	CString path = MboxMail::GetLastPath();
	if (path.IsEmpty())
		return;
	path.TrimRight(L"\\");

	CString itemName = m_tree.GetItemText(hItem);

	HTREEITEM hParent = m_tree.GetParentItem(hItem);
	HTREEITEM hFolder = 0;
	if (hParent == 0)
	{
		hFolder = HasFolder(path);
		if (hFolder == 0)
			return;
	}
	else
	{
		CString ParentItemName = m_tree.GetItemText(hParent);
		hFolder = hParent;

		DWORD nId = (DWORD)m_tree.GetItemData(hParent);
		LabelInfo* linfo = m_labelInfoStore.Find(nId);
	}

	CString rootName = m_tree.GetItemText(hFolder);

	CString datapath = MboxMail::GetLastDataPath();
	FileUtils::CreateDir(datapath);
	datapath.TrimRight(L"\\");

	CFileException ExError;
	CString viewFile = datapath + L"\\.mboxview";
	if (fp.Open(viewFile, CFile::modeWrite | CFile::modeCreate, &ExError))
	{
		Traverse(hFolder, fp, *fileSizes);
		fp.Close();
	}
	else
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not create \"" + viewFile;
		txt += L"\" file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);
	}
#ifdef _DEBUG
	// Test
	TRACE(L"SaveData: Hash Table fileSizes Count=%d\n", fileSizes->GetCount());
	ArchiveFileInfo info;
	POSITION pos = fileSizes->GetStartPosition();
	while (pos)
	{
		fileSizes->GetNextAssoc(pos, path, info);
		TRACE(L"File=%s FileSize=%lld ShowFile=%d\n", path, info.fSize, info.bShow);
	}
#endif
}

// Clear .mboxview file
void NTreeView::ClearData()
{
	CString datapath = MboxMail::GetLastDataPath();
	datapath.TrimRight(L"\\");

	CString viewFile = datapath + L"\\.mboxview";
	BOOL bRetval = FileUtils::DelFile(viewFile, FALSE);
}


// Not used anymore
void NTreeView::UpdateFileSizesTable(CString& path, _int64 realFSize, FileSizeMap& fileSizes)
{
	_int64 fSize = 0;
	ArchiveFileInfo info;
	BOOL found = fileSizes.Lookup(path, info);
	if (found)
	{
		if (info.fSize != realFSize)
		{
			fileSizes.RemoveKey(path);
			fileSizes[path].fSize = realFSize;
			fileSizes[path].bShow = 1;
		}
	}
	else
	{
		fileSizes[path].fSize = realFSize;
		fileSizes[path].bShow = 1;
	}
}

void NTreeView::OnTreeExpand()
{
	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (pFrame)
		pFrame->UpdateToolsBar();
}

void NTreeView::OnUpdateTreeExpand(CCmdUI* pCmdUI)
{
	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	BOOL isTreeHidden = FALSE;
	if (pFrame)
	{
		isTreeHidden = pFrame->IsTreeHidden();
	}

	BOOL enable = (m_tree.GetRootItem() != NULL) && !isTreeHidden;
	pCmdUI->Enable(enable);
}

void NTreeView::OnUpdateFileRefresh(CCmdUI* pCmdUI)
{
	pCmdUI->Enable(m_tree.GetRootItem() != NULL);
}

void NTreeView::OnFileRefresh()
{
	if (!CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
	{
		m_tree.DeleteAllItems();
		LoadFolders();
	}
	return;
}

void MyAppendMenu(CMenu* menu, int commandId, const wchar_t* commandName, BOOL checkMark)
{
	UINT nFlags = MF_STRING;
	if (checkMark)
		nFlags = MF_STRING | MF_CHECKED;
	menu->AppendMenu(nFlags, commandId, commandName);
	menu->AppendMenu(MF_SEPARATOR);
}

void NTreeView::OnRClick(NMHDR* pNMHDR, LRESULT* pResult)
{

	SCROLLINFO si;
	si.cbSize = sizeof(si);
	si.fMask = SIF_ALL;
	BOOL ret = m_tree.GetScrollInfo(SB_HORZ, &si);
	TRACE(L"GetScrollInfo: pos=%d\n", si.nPos);
	si.nPos = 0;
	ret = m_tree.SetScrollInfo(SB_HORZ, &si);



	// TODO: Add your control notification handler code here
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);

	*pResult = 0;

	CPoint pt;
	::GetCursorPos(&pt);
	CWnd* wnd = WindowFromPoint(pt);

	CPoint ptClient(pt);
	ScreenToClient(&ptClient);

	HTREEITEM hItem = m_tree.GetSelectedItem();
	CString itemTxt;
	if (hItem)
		itemTxt = m_tree.GetItemText(hItem);

	UINT flags = TVHT_TORIGHT | TVHT_ABOVE;  // 1280 = 0x400 + 0x100  - flags are populated by HitTest
	HTREEITEM hTreeItem = m_tree.HitTest(ptClient, &flags);

	if (hTreeItem == 0)
		return;

	CString mailFile = m_tree.GetItemText(hTreeItem);

	if (!(flags & TVHT_ONITEM))
		return;

	HTREEITEM hRoot = DetermineRootItem(hTreeItem);

	if (hRoot == 0)
		return;


	if (hItem == 0)
		return;


	CString roottxt = m_tree.GetItemText(hRoot);

	//if (!(m_tree.GetItemState(hRoot, TVIS_EXPANDED) & TVIS_EXPANDED))
		//return;

	HTREEITEM hParent = m_tree.GetParentItem(hTreeItem);

	if ((hParent == 0) && (hItem != hTreeItem))
	{
		CString errorText = L"Right Click is supported on selected items only.\nUse left click to select mail folder first.";
		ResHelper::TranslateString(errorText);
		HWND h = wnd->GetSafeHwnd();
		int answer = ::MessageBox(h, errorText, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;
	}

	if ((hParent != 0) && (hItem != hTreeItem))
	{
		CString errorText = L"Right Click is supported on selected items only.\nUse left click to select mail file first.";
		ResHelper::TranslateString(errorText);
		HWND h = wnd->GetSafeHwnd();
		int answer = ::MessageBox(h, errorText, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;
	}

	if (hItem)
	{
		DWORD nId = (DWORD)m_tree.GetItemData(hItem);
		LabelInfo* linfo = m_labelInfoStore.Find(nId);
		if (linfo)
		{
			if (linfo->m_nodeType == LabelInfo::MailSubFolder)
			{
				CMenu menu;
				menu.CreatePopupMenu();
				//menu.AppendMenu(MF_SEPARATOR);


				const UINT M_FolderPath_Id = 1;
				MyAppendMenu(&menu, M_FolderPath_Id, L"Show Folder Path");
				const UINT M_FolderLocation_Id = 2;
				MyAppendMenu(&menu, M_FolderLocation_Id, L"Open Folder Location");
				const UINT M_RefreshSubFolder_Id = 3;
				MyAppendMenu(&menu, M_RefreshSubFolder_Id, L"Refresh Folder");
				const UINT M_OpenHiddenFiles_Id = 4;
				MyAppendMenu(&menu, M_OpenHiddenFiles_Id, L"Restore Removed Files");
				const UINT M_MergeMailArchiveFiles_Id = 5;
				MyAppendMenu(&menu, M_MergeMailArchiveFiles_Id, L"Merge Mail Archive Files");

				int index = 1;
				ResHelper::LoadMenuItemsInfo(&menu, index);
				ResHelper::UpdateMenuItemsInfo(&menu, index);

				CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

				UINT command = menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, this);

				UINT nFlags = TPM_RETURNCMD;
				CString menuString;
				int chrCnt = menu.GetMenuString(command, menuString, nFlags);

				switch (command)
				{
				case M_FolderPath_Id:
				{
					CString pathLast = MboxMail::GetLastPath();
					pathLast.TrimRight(L"\\");
					int nId = (DWORD)m_tree.GetItemData(hItem);
					//CString path = m_folderArray.m_array.GetAt(nId);
					CString path = linfo->m_mailFolderPath;
					if (path.Compare(pathLast))
						int deb = 1;
					HWND h = wnd->GetSafeHwnd();
					int answer = ::MessageBox(h, path, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_USERICON);
				}
				break;
				case M_FolderLocation_Id:
				{
					UINT nFlags = MF_BYCOMMAND;
					CString Label;
					int retLabel = menu.GetMenuString(M_FolderLocation_Id, Label, nFlags);

					CString path = MboxMail::GetLastPath();
					//if (FileUtils::BrowseToFile(MboxMail::s_path) == FALSE) {  // TODO: s_path error checking ??
					{
						HWND h = GetSafeHwnd();
						HINSTANCE result = ShellExecute(h, L"open", path, NULL, NULL, SW_SHOWNORMAL);
						CMainFrame::CheckShellExecuteResult(result, h);
					}
				}
				break;
				case M_OpenHiddenFiles_Id:
				{
					if (hItem)
					{
						int ret = OpenHiddenFiles(hItem, *fileSizes);
					}
				}
				break;

				case M_RefreshSubFolder_Id:
				{
					if (hItem)
					{
						int ret = RefreshSubFolder(hItem);
					}
				}
				break;

				case M_MergeMailArchiveFiles_Id:
				{
					if (hItem)
					{
						int ret = MergeMailArchiveFiles(hItem);
					}
				}
				break;

				default: {
					int deb = 1;
				}
					   break;
				}

				return;
			}
		}
	}

	if ((hParent == 0) && (hItem == hTreeItem))
	{
		CMenu menu;
		menu.CreatePopupMenu();
		//menu.AppendMenu(MF_SEPARATOR);

		//
		const UINT M_FolderPath_Id = 1;
		MyAppendMenu(&menu, M_FolderPath_Id, L"Show Folder Path");
		const UINT M_FolderLocation_Id = 2;
		MyAppendMenu(&menu, M_FolderLocation_Id, L"Open Folder Location");
		const UINT M_DeleteItem_Id = 3;
		MyAppendMenu(&menu, M_DeleteItem_Id, L"Remove Folder");
		const UINT M_FolderRefresh_Id = 4;
		MyAppendMenu(&menu, M_FolderRefresh_Id, L"Refresh Folder");
		const UINT M_OpenHiddenFiles_Id = 5;
		MyAppendMenu(&menu, M_OpenHiddenFiles_Id, L"Restore Removed Files");
		const UINT M_MergeMailArchiveFiles_Id = 6;
		MyAppendMenu(&menu, M_MergeMailArchiveFiles_Id, L"Merge Mail Archive Files");

		int index = 1;
		ResHelper::LoadMenuItemsInfo(&menu, index);
		ResHelper::UpdateMenuItemsInfo(&menu, index);

		CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

		UINT command = menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, this);

		UINT nFlags = TPM_RETURNCMD;
		CString menuString;
		int chrCnt = menu.GetMenuString(command, menuString, nFlags);

		switch (command)
		{
		case M_DeleteItem_Id:
		{
			UINT nFlags = MF_BYCOMMAND;
			CString Label;
			int retLabel = menu.GetMenuString(M_DeleteItem_Id, Label, nFlags);

			MboxMail::ShowHint(HintConfig::MessageRemoveFolderHint, GetSafeHwnd());
			if (hItem)
				BOOL ret = DeleteFolder(hItem);
		}
		break;
		case M_FolderPath_Id:
		{
			CString pathLast = MboxMail::GetLastPath();
			pathLast.TrimRight(L"\\");
			int nId = (DWORD)m_tree.GetItemData(hItem);
			CString path = m_folderArray.m_array.GetAt(nId);
			if (path.Compare(pathLast))
				int deb = 1;
			HWND h = wnd->GetSafeHwnd();
			int answer = ::MessageBox(h, path, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_USERICON);
		}
		break;
		case M_FolderLocation_Id:
		{
			UINT nFlags = MF_BYCOMMAND;
			CString Label;
			int retLabel = menu.GetMenuString(M_FolderLocation_Id, Label, nFlags);

			CString path = MboxMail::GetLastPath();
			//if (FileUtils::BrowseToFile(MboxMail::s_path) == FALSE) {  // TODO: s_path error checking ??
			{
				HWND h = GetSafeHwnd();
				HINSTANCE result = ShellExecute(h, L"open", path, NULL, NULL, SW_SHOWNORMAL);
				CMainFrame::CheckShellExecuteResult(result, h);
			}
		}
		break;
		case M_FolderRefresh_Id:
		{
			CString path = MboxMail::GetLastPath();
			path.TrimRight(L"\\");
			HWND h = wnd->GetSafeHwnd();
			//int answer = ::MessageBox(h, path, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_USERICON);
			if (hItem)
			{
				BOOL ret = RefreshFolder(hItem);
			}
		}
		break;

		case M_OpenHiddenFiles_Id:
		{
			if (hItem)
			{
				int ret = OpenHiddenFiles(hItem, *fileSizes);
			}
		}
		break;

		case M_MergeMailArchiveFiles_Id:
		{
			if (hItem)
			{
				int ret = MergeMailArchiveFiles(hItem);
			}
		}
		break;

		default: {
			int deb = 1;
		}
			   break;
		}

		return;
	}

	BOOL isLabel = FALSE;
	LabelInfo* linfo = 0;
	DWORD nId = 0;
	if (hItem)
	{
		nId = (DWORD)m_tree.GetItemData(hItem);
		linfo = m_labelInfoStore.Find(nId);
		if (linfo && !linfo->m_listFilePath.IsEmpty())
			isLabel = TRUE;
	}

	CMenu menu;
	menu.CreatePopupMenu();
	menu.AppendMenu(MF_SEPARATOR);

	//
	CMenu printToSubMenu;
	printToSubMenu.CreatePopupMenu();
	printToSubMenu.AppendMenu(MF_SEPARATOR);

	const UINT S_CSV_Id = 1;
	MyAppendMenu(&printToSubMenu, S_CSV_Id, L"CSV...");

	const UINT S_TEXT_Id = 2;
	MyAppendMenu(&printToSubMenu, S_TEXT_Id, L"Text..");

	const UINT S_HTML_Id = 3;
	MyAppendMenu(&printToSubMenu, S_HTML_Id, L"HTML..");

	const UINT S_PDF_Id = 33;
	MyAppendMenu(&printToSubMenu, S_PDF_Id, L"PDF..");

	menu.AppendMenu(MF_POPUP | MF_STRING, (INT_PTR)printToSubMenu.GetSafeHmenu(), L"Print To");
	menu.AppendMenu(MF_SEPARATOR);

	//
	CMenu sortSubMenu;
	sortSubMenu.CreatePopupMenu();
	sortSubMenu.AppendMenu(MF_SEPARATOR);

	const UINT S_SORT_BY_DATE_Id = 4;
	MyAppendMenu(&sortSubMenu, S_SORT_BY_DATE_Id, L"Date");

	const UINT S_SORT_BY_FROM_Id = 5;
	MyAppendMenu(&sortSubMenu, S_SORT_BY_FROM_Id, L"From");

	const UINT S_SORT_BY_TO_Id = 6;
	MyAppendMenu(&sortSubMenu, S_SORT_BY_TO_Id, L"To");

	const UINT S_SORT_BY_SUBJ_Id = 7;
	MyAppendMenu(&sortSubMenu, S_SORT_BY_SUBJ_Id, L"Subject");

	const UINT S_SORT_BY_DATE_AND_SUBJ_Id = 22;
	//AppendMenu(&sortSubMenu, S_SORT_BY_DATE_AND_SUBJ_Id, L"Subject and Date");

	const UINT S_SORT_BY_SIZE_Id = 8;
	MyAppendMenu(&sortSubMenu, S_SORT_BY_SIZE_Id, L"Size");

	const UINT S_SORT_BY_CONVERSATION_Id = 9;
	MyAppendMenu(&sortSubMenu, S_SORT_BY_CONVERSATION_Id, L"Conversation");

	const UINT S_SORT_BY_POSITION_Id = 10;
	// Sort by position in the archive file. Enabled for debugging only
	//AppendMenu(&sortSubMenu, S_SORT_BY_POSITION_Id, L"Mail ID");

	menu.AppendMenu(MF_POPUP | MF_STRING, (INT_PTR)sortSubMenu.GetSafeHmenu(), L"Sort By");
	menu.AppendMenu(MF_SEPARATOR);
	//

	const UINT M_FileLocation_Id = 11;
	MyAppendMenu(&menu, M_FileLocation_Id, L"Open File Location");

	const UINT M_DatabaseLocation_Id = 12;
	MyAppendMenu(&menu, M_DatabaseLocation_Id, L"Open Database Folder Location");

	const UINT M_Properties_Id = 13;
	MyAppendMenu(&menu, M_Properties_Id, L"Properties");

	const UINT M_AttachmentCache_Id = 14;
	MyAppendMenu(&menu, M_AttachmentCache_Id, L"Export All Mail Attachments");

	const UINT M_EmlCache_Id = 15;
	MyAppendMenu(&menu, M_EmlCache_Id, L"Export All Mails as Eml");

	const UINT M_Reload_Id = 16;
	const UINT M_Remove_Id = 17;
	const UINT S_Labels_Create_Id = 18;
	const UINT S_Labels_Delete_Id = 19;
	const UINT S_Labels_Refresh_Id = 20;
	const UINT M_CreateFolder_Id = 21;
	//
	// const UINT S_SORT_BY_DATE_AND_SUBJ_Id = 22;   // see above


	CMenu labelsSubMenu;
	labelsSubMenu.CreatePopupMenu();

	if (isLabel == FALSE)
	{
		MyAppendMenu(&menu, M_Reload_Id, L"Refresh Index File");

		MyAppendMenu(&menu, M_Remove_Id, L"Remove File");

		MyAppendMenu(&labelsSubMenu, S_Labels_Create_Id, L"Create...");
		MyAppendMenu(&labelsSubMenu, S_Labels_Delete_Id, L"Delete...");
		MyAppendMenu(&labelsSubMenu, S_Labels_Refresh_Id, L"Refresh...");

		menu.AppendMenu(MF_POPUP | MF_STRING, (INT_PTR)labelsSubMenu.GetSafeHmenu(), L"Gmail Labels");
		menu.AppendMenu(MF_SEPARATOR);

		//AppendMenu(&menu, M_CreateFolder_Id, L"Create Folder");  // TODO: later
	}

	int index = 1;
	ResHelper::LoadMenuItemsInfo(&menu, index);
	ResHelper::UpdateMenuItemsInfo(&menu, index);

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	UINT command = menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, this);

	UINT nFlags = TPM_RETURNCMD;
	CString menuString;
	int chrCnt = menu.GetMenuString(command, menuString, nFlags);

	NListView* pListView = 0;
	if (pFrame) {
		pListView = pFrame->GetListView();
	}

	switch (command)
	{
	case M_FileLocation_Id:
	{
		UINT nFlags = MF_BYCOMMAND;
		CString Label;
		int retLabel = menu.GetMenuString(M_FileLocation_Id, Label, nFlags);

		CString path = MboxMail::GetLastPath();
		if (FileUtils::BrowseToFile(MboxMail::s_path) == FALSE) {  // TODO: s_path error checking ??
			HWND h = GetSafeHwnd();
			HINSTANCE result = ShellExecute(h, L"open", path, NULL, NULL, SW_SHOWNORMAL);
			CMainFrame::CheckShellExecuteResult(result, h);
		}
	}
	break;
	case M_DatabaseLocation_Id:
	{
		UINT nFlags = MF_BYCOMMAND;
		CString Label;
		int retLabel = menu.GetMenuString(M_DatabaseLocation_Id, Label, nFlags);

		if (linfo)
		{
			if (linfo->m_nodeType == LabelInfo::MailLabel)
			{
				CString path = linfo->m_listFilePath;
				if (FileUtils::PathFileExist(path))
				{
					if (FileUtils::BrowseToFile(path) == FALSE) {  // TODO: s_path error checking ??
						HWND h = GetSafeHwnd();
						HINSTANCE result = ShellExecute(h, L"open", path, NULL, NULL, SW_SHOWNORMAL);
						CMainFrame::CheckShellExecuteResult(result, h);
					}
				}
			}
			else
			{
				CString path = MboxMail::GetLastDataPath();
				//if (FileUtils::BrowseToFile(MboxMail::s_path) == FALSE) {  // TODO: s_path error checking ??
				{
					HWND h = GetSafeHwnd();
					HINSTANCE result = ShellExecute(h, L"open", path, NULL, NULL, SW_SHOWNORMAL);
					CMainFrame::CheckShellExecuteResult(result, h);
				}
			}
		}
		else
		{
			CString path = MboxMail::GetLastDataPath();
			//if (FileUtils::BrowseToFile(MboxMail::s_path) == FALSE) {  // TODO: s_path error checking ??
			{
				HWND h = GetSafeHwnd();
				HINSTANCE result = ShellExecute(h, L"open", path, NULL, NULL, SW_SHOWNORMAL);
				CMainFrame::CheckShellExecuteResult(result, h);
			}
		}
	}
	break;
	case M_Properties_Id:
	{
		wchar_t sizeStr_inKB[256];
		wchar_t sizeStr_inBytes[256];
		int sizeStrSize = 256;
		CString txt;
		CString tmp;

		_int64 fileSize = FileUtils::FileSize(MboxMail::s_path); // TODO: error checking ??
		LPCWSTR fileSizeStr_inKB = StrFormatKBSize(fileSize, &sizeStr_inKB[0], sizeStrSize);
		if (!fileSizeStr_inKB)
			sizeStr_inKB[0] = 0;

		LPCWSTR fileSizeStr_inBytes = StrFormatByteSize64(fileSize, &sizeStr_inBytes[0], sizeStrSize);
		if (!fileSizeStr_inBytes)
			sizeStr_inBytes[0] = 0;

		int mailCount = MboxMail::s_mails_ref.GetCount();

		txt.Empty();
		CString folder;
		FileUtils::GetFolderPath(MboxMail::s_path, folder);
		txt.Format(L"Folder: %s\n", folder);

		tmp.Format(L"File: %s\n", mailFile);
		txt.Append(tmp);

		CString cstr;
		INT64 numb = fileSize;
		TextUtilsEx::Int2WstrWithCommas(numb, cstr);

		tmp.Format(L"File size:  %s  (%s) (%s)\n", cstr, sizeStr_inKB, sizeStr_inBytes);
		txt += tmp;
		tmp.Empty();
		tmp.Format(L"Mail Count: %d\n", mailCount);
		txt += tmp;

		HWND h = GetSafeHwnd();
		int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		int deb = 1;
	}
	break;
	case S_CSV_Id:
	{
		if (pFrame) {
			pFrame->OnFileExportToCsv();
		}
		int deb = 1;
	}
	break;
	case S_TEXT_Id:
	{
		UINT nFlags = MF_BYCOMMAND;
		CString Label;
		int retLabel = menu.GetMenuString(S_TEXT_Id, Label, nFlags);
		if (pFrame)
		{
			pFrame->PrintMailArchiveToTEXT();
		}

		int deb = 1;
	}
	break;
	case S_HTML_Id:
	{
		if (pFrame)
		{
			pFrame->PrintMailArchiveToHTML();
		}

		int deb = 1;
	}
	break;
	case S_PDF_Id:
	{
		if (pFrame)
		{
			MboxMail::ShowHint(HintConfig::PrintToPDFHint, GetSafeHwnd());
			pFrame->PrintMailArchiveToPDF();
		}
		int deb = 1;
	}
	break;
	case S_SORT_BY_DATE_Id: {
		if (pListView)
			pListView->SortByColumn(1);
	}
	break;
	case S_SORT_BY_FROM_Id: {
		if (pListView)
			pListView->SortByColumn(2);
	}
	break;
	case S_SORT_BY_TO_Id: {
		if (pListView)
			pListView->SortByColumn(3);
	}
	break;
	case S_SORT_BY_SUBJ_Id: {
		if (pListView)
			pListView->SortByColumn(4);
	}
	break;
	case S_SORT_BY_DATE_AND_SUBJ_Id:
	{
		if (pListView)
			pListView->SortBySubjectBasedConversasions();
	}
	break;
	case S_SORT_BY_SIZE_Id: {
		if (pListView)
			pListView->SortByColumn(5);
	}
	break;
	case S_SORT_BY_CONVERSATION_Id: {
		if (pListView)
			pListView->SortByColumn(0);
	}
	break;
								  // Sort by position in the archive file. Enabled for debugging only
	case S_SORT_BY_POSITION_Id: {
		if (pListView)
			pListView->SortByColumn(1, TRUE);
	}
	break;

	case M_Reload_Id:
	{
		CString txt = L"Do you want to refresh index file?";
		ResHelper::TranslateString(txt);

		HWND h = GetSafeHwnd();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDYES)
			ForceParseMailFile(hItem);
		// For internal testing
		//int ret = MboxMail::DumpMailStatsToFile(&MboxMail::s_mails, MboxMail::s_mails.GetCount());
	}
	break;

	case M_AttachmentCache_Id:
	{
		CString txt = L"Do you want to create cache with all attachments?"
			L"\n\nNote that \"File->Attachments Config\" controls whether image attachments embedded into messages will also be exported";
		ResHelper::TranslateString(txt);

		HWND h = GetSafeHwnd();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDYES)
		{
			int firstMail = 0;
			int lastMail = MboxMail::s_mails.GetCount() - 1;

			CString targetPrintSubFolderName = L"";
			if (pListView)
			{
				int retval = pListView->CreateAttachmentCache_Thread(firstMail, lastMail, targetPrintSubFolderName);
			}
		}
	}
	break;

	case M_EmlCache_Id:
	{
		CString txt = L"Do you want to create cache with all Eml files?";
		ResHelper::TranslateString(txt);

		HWND h = GetSafeHwnd();
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDYES)
		{
			int firstMail = 0;
			int lastMail = MboxMail::s_mails.GetCount() - 1;

			CString targetPrintSubFolderName = L"";
			if (pListView)
			{
				int retval = pListView->CreateEmlCache_Thread(firstMail, lastMail, targetPrintSubFolderName);
			}
		}
	}
	break;

	case M_CreateFolder_Id:
	{
		CreateEmptyFolder(hItem);
		CMenu menu;
		CString mboxFilePath = MboxMail::s_path;
		CString mboxFileName = ::PathFindFileName(mboxFilePath);
		HTREEITEM hItem = NTreeView::FindItem(0, mboxFileName);
		TRACE(L"CreateFlatFolderListMenu\n");
		CArray<CString> folderList;
		CreateFlatFolderList(hItem, folderList);
		int deb = 1;
	}
	break;

	case M_Remove_Id:
	{
		CString mailFile = m_tree.GetItemText(hItem);
		RemoveFileFromTreeView(hItem, *fileSizes);
	}
	break;

	case S_Labels_Create_Id:
	{
		//int retval = CreateLabelFilesForSingleMailFile(hItem);
		int retval = RefreshLabelsForSingleMailFile(hItem);
	}
	break;
	case S_Labels_Delete_Id:
	{
		int retval = DeleteLabelsForSingleMailFile(hItem);
	}
	break;
	case S_Labels_Refresh_Id:
	{
		int retval = RefreshLabelsForSingleMailFile(hItem);
	}
	break;

	default: {
		int deb = 1;
	}
		   break;
	}

	Invalidate();
	UpdateWindow();

	*pResult = 0;
}

// When this called?
void NTreeView::OnSelchanging(NMHDR* pNMHDR, LRESULT* pResult)
{
	NMTREEVIEW* pNm = (LPNMTREEVIEW)pNMHDR;
	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	__int64 res = *pResult;
	*pResult = 0;

	if (pFrame == NULL)
		return;
	NListView* pListView = pFrame->GetListView();
	if (!pListView)
		return;
	HTREEITEM hNewItem = pNm->itemNew.hItem;
	HTREEITEM hOldItem = pNm->itemOld.hItem;

	DWORD nIdNew = (DWORD)m_tree.GetItemData(hNewItem);
	LabelInfo* linfoNew = m_labelInfoStore.Find(nIdNew);

	if (hOldItem == 0)
		return;

	DWORD nIdOld = (DWORD)m_tree.GetItemData(hOldItem);
	LabelInfo* linfoOld = m_labelInfoStore.Find(nIdOld);

	if (linfoNew && linfoOld)
	{
		if (linfoNew->m_filePath.Compare(linfoOld->m_filePath) == 0)
			return;
	}

	BOOL isDirty = MboxMail::m_editMails.m_bIsDirty;
	if (!pListView->IsUserSelectedMailListEmpty())
	{
		if (isDirty)
		{
			CString txt = L"User Selected Mails list is not empty.\n"
				L"Content will be lost if you switch to new mail archive.\n"
				L"Do you want to continue?";
			ResHelper::TranslateString(txt);
			int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
			if (answer == IDNO)
			{
				*pResult = 1;
			}
			else if (answer == IDYES)
				MboxMail::m_editMails.m_bIsDirty = FALSE;
		}
	}
}

// TODO: Non Recursive to Search direct children only
// The same name can appear in multiple folders
// Recursive search doesn't make sense
// Since we are looking for the mbox file name and not the full path
HTREEITEM NTreeView::FindItem(HTREEITEM hItem, CString& mailFileName, BOOL recursive)
{
	_ASSERTE(recursive == FALSE);

	CString path;
	if (hItem == TVI_ROOT)
		hItem = 0;

	if (m_tree.ItemHasChildren(hItem))
	{
		HTREEITEM hChild = m_tree.GetChildItem(hItem);
		while (hChild)
		{
			path = m_tree.GetItemText(hChild);
			if (path.Compare(mailFileName) == 0)
				return hChild;
			hChild = m_tree.GetNextSiblingItem(hChild);
		}
	}
	return 0;
}


// TODO: dos't work , fix it, GetItemText doesn't  return full path, just name part
HTREEITEM NTreeView::FindItemByItemPath(HTREEITEM hItem, CString& mailFileName, BOOL recursive)
{
	CString path;
	if (hItem == TVI_ROOT)
		hItem = 0;

	if (m_tree.ItemHasChildren(hItem))
	{
		HTREEITEM hChild = m_tree.GetChildItem(hItem);
		while (hChild)
		{
			path = m_tree.GetItemText(hChild);
			if (path.Compare(mailFileName) == 0)
				return hChild;
			hChild = m_tree.GetNextSiblingItem(hChild);
		}
	}
	return 0;
}


// Non recursive
BOOL NTreeView::DeleteFolder(HTREEITEM hItem, BOOL setNewPath)
{
	if (hItem == 0)
		return FALSE;

	CString itemTxt;
	UINT nCode = TVE_COLLAPSE;
	ExpandItem(hItem, nCode);

	CString datapath = MboxMail::GetLastDataPath();

	CString folderPath;
	DetermineFolderPath(hItem, folderPath);

	MboxMail::SetLastPath(folderPath);
	BOOL ret = m_folderArray.Delete(folderPath);

	BOOL bRetval = RemoveFileSizeMap(folderPath);

	m_globalFolderInfoDB.RemoveRootFolder(folderPath);
	RemoveFolderAsRootFolder(folderPath);

	itemTxt = m_tree.GetItemText(hItem);
	TRACE(L"Deleting %s\n", itemTxt);
	// TODO: DeleteItemChildren doesn't update label table or m_globalFolderInfoDB
	// m_tree.DeleteItem(hItem); delets all children
	// Reenable when we need to do cleanup per Tree item
	// Commented out for now
	// DeleteItemChildren(hItem);
	m_tree.DeleteItem(hItem);

	TRACE(L"=================>>> Delete Folder\n");
	m_globalFolderInfoDB.Print();

	if (setNewPath)
	{
		HTREEITEM hRootItem = m_tree.GetRootItem();
		if (hRootItem)
		{
			CString folderPath;
			itemTxt = m_tree.GetItemText(hRootItem);
			DetermineFolderPath(hRootItem, folderPath);
			MboxMail::SetLastPath(folderPath);

			BOOL selectFolder = TRUE;
			BOOL unconditionalFolderInsert = FALSE;
			BOOL expand = FALSE;
			HTREEITEM hItem = FillCtrl(hRootItem, selectFolder, unconditionalFolderInsert, expand);
		}
		else
		{
			CString empty = L"";
			MboxMail::SetLastPath(empty);
		}
	}
	else
	{
		CString empty = L"";
		MboxMail::SetLastPath(empty);
	}

	//TRACE("DeleteFolder: AssertTreeCtrl\n");
	//BOOL retAssert = AssertTreeCtrl();
	TRACE(L"DeleteFolder: AssertNTreeView\n");
	BOOL retAssert = AssertNTreeView();
	return ret;
}

// Open means installed already under Tree View
// TODO: it looks like redundant code, needs review
HTREEITEM NTreeView::SelectFolderIfAlreadyOpen(CString* path, int folderType, BOOL expand)
{
	CString folderPath = *path;
	folderPath.TrimRight(L"\\");
	folderPath.Append(L"\\");

	HTREEITEM hFolder;
	if (hFolder = IsFolderOpen(*path))
	{
		DWORD nId = (DWORD)m_tree.GetItemData(hFolder);
		LabelInfo* linfo = m_labelInfoStore.Find(nId);

		hashsum_t hashsum = GlobalFolderInfoDB::GetHashsum(&folderPath);
		FolderInfo* rfinfo = m_globalFolderInfoDB.FindRoot(&folderPath);
		if (rfinfo == 0)
		{
			return 0;
		}

#if 0
		FolderInfo* finfo = m_globalFolderInfoDB.Find(&path, hashsum);
		if (finfo == 0)
		{
			return 0;
		}

		if (finfo->m_folderType != FolderInfo::MailFolderNode)
			return 0;
#endif

		//if (rfinfo->m_folderType != FolderInfo::MailFolderNode)
		if (rfinfo->m_folderType != folderType)
			return 0;


		// This will happen when user tries to select folder which already is inserted into Tree
		HTREEITEM hCurrentSelectedItem = m_tree.GetSelectedItem();
		if (hCurrentSelectedItem)
			CString itemName = m_tree.GetItemText(hCurrentSelectedItem);

		if (hCurrentSelectedItem != hFolder)
		{
			if (hCurrentSelectedItem != 0)
				m_tree.SetItemState(hCurrentSelectedItem, 0, TVIS_BOLD);

			BOOL retval = m_tree.SelectItem(hFolder);
		}

		if (expand)
			ExpandItem(hFolder, TVE_EXPAND);
		return hFolder;
	}
	return 0;
}

HTREEITEM NTreeView::SelectTreeItem(CString* path)
{
	HTREEITEM hFolder;
	if (hFolder = HasFolder(*path))
	{
		// This will happen when user tries to select folder which already is inserted into Tree
		HTREEITEM hCurrentSelectedItem = m_tree.GetSelectedItem();
		if (hCurrentSelectedItem)
			CString itemName = m_tree.GetItemText(hCurrentSelectedItem);

		if (hCurrentSelectedItem != hFolder)
		{
			if (hCurrentSelectedItem != 0)
				m_tree.SetItemState(hCurrentSelectedItem, 0, TVIS_BOLD);

			BOOL retval = m_tree.SelectItem(hFolder);
		}

		ExpandItem(hFolder, TVE_EXPAND);
		return hFolder;
	}
	return 0;
}

BOOL NTreeView::SelectTreeItem(HTREEITEM hItem)
{

	if (hItem)
	{
		//UINT nCode = TVGN_CARET;
		//BOOL retval = m_tree.Select(hItem, nCode);
		BOOL retval = m_tree.SelectItem(hItem);
		return retval;
	}
	else
	{
		MboxMail::assert_unexpected();
		return FALSE;
	}
}

void NTreeView::SortChildren(HTREEITEM hItem, BOOL recursive)
{
#if 1
	TVSORTCB tvs;

	// Sort the tree control's items using custom callback procedure.
	tvs.hParent = hItem;
	tvs.lpfnCompare = MyCompareProc;
	tvs.lParam = (LPARAM)this;

	m_tree.SortChildrenCB(&tvs);
#else
	if (!recursive)
	{
		m_tree.SortChildren(hItem);
	}
	else
	{
		if (hItem == NULL)
			hItem = TVI_ROOT;
		if (hItem == TVI_ROOT || m_tree.ItemHasChildren(hItem))
		{
			HTREEITEM child = m_tree.GetChildItem(hItem);

			while (child != NULL)
			{
				SortChildren(child, recursive);
				child = m_tree.GetNextItem(child, TVGN_NEXT);
			}

			m_tree.SortChildren(hItem);
		}
	}
#endif
}

// Root item is not included into the children count
// TODO!!!: verify. Doesn't look good. Unnecessary code. Add recursion
int NTreeView::GetChildrenCount(HTREEITEM hItem, BOOL recursive)
{
	if (hItem == 0)
		return -1;

	int count = 0;
	CString itemTxt;
	itemTxt = m_tree.GetItemText(hItem);

	HTREEITEM hNextNextItem;
	HTREEITEM hNextItem = m_tree.GetChildItem(hItem);

	while (hNextItem != NULL)
	{
		hNextNextItem = m_tree.GetNextSiblingItem(hNextItem);
		if (m_tree.ItemHasChildren(hNextItem))
		{
			HTREEITEM hChild = m_tree.GetChildItem(hNextItem);
			if (hChild)
				count++;
		}
		else
		{
			itemTxt = m_tree.GetItemText(hNextItem);
			count++;
		}
		hNextItem = hNextNextItem;
	}
	return count;
}

BOOL NTreeView::DeleteItemChildren(HTREEITEM hItem)
{
	if (hItem == 0)
		return FALSE;

	UINT nCode = TVE_COLLAPSE;
	ExpandItem(hItem, nCode);

	CString itemTxt;
	itemTxt = m_tree.GetItemText(hItem);

	HTREEITEM hNextNextItem;
	HTREEITEM hNextItem = m_tree.GetChildItem(hItem);

	while (hNextItem != NULL)
	{
		hNextNextItem = m_tree.GetNextSiblingItem(hNextItem);
		if (m_tree.ItemHasChildren(hNextItem))
		{
			NTreeView::DeleteItem(hNextItem);
		}
		else
		{
			itemTxt = m_tree.GetItemText(hNextItem);
			TRACE(L"\tDeleting %s\n", itemTxt);
			m_tree.DeleteItem(hNextItem);
		}
		hNextItem = hNextNextItem;
	}
	return TRUE;
}

BOOL NTreeView::DeleteItem(HTREEITEM hItem)
{
	if (hItem == 0)
		return FALSE;

	CString itemTxt;
	//UINT nCode = TVE_COLLAPSE;
	//ExpandItem(hItem, nCode);

	itemTxt = m_tree.GetItemText(hItem);

	HTREEITEM hNextNextItem;
	HTREEITEM hNextItem = m_tree.GetChildItem(hItem);

	while (hNextItem != NULL)
	{
		hNextNextItem = m_tree.GetNextSiblingItem(hNextItem);
		itemTxt = m_tree.GetItemText(hItem);
		if (hNextItem)
			itemTxt = m_tree.GetItemText(hNextItem);

		if (m_tree.ItemHasChildren(hNextItem))
		{
			NTreeView::DeleteItem(hNextItem);
		}
		else
		{
			itemTxt = m_tree.GetItemText(hNextItem);
			TRACE(L"\tDeleting %s\n", itemTxt);
			m_tree.DeleteItem(hNextItem);
		}
		hNextItem = hNextNextItem;
	}
	itemTxt = m_tree.GetItemText(hItem);
	TRACE(L"Deleting %s\n", itemTxt);
	m_tree.DeleteItem(hItem);
	return TRUE;
}

void NTreeView::StartTimer()
{
	SetTimer(m_nIDEvent, m_nElapse, NULL);
}


void NTreeView::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: Add your message handler code here and/or call default

	if (m_timerTickCnt == 0)
		SelectMailFile();  // will trigger NTreeView::OnSelchanged
	m_timerTickCnt++;

	// Need just one shot timer
	// Is killing timer from the timer callback ok ? Seem to work -:)
	// May need to move to FillCtrl() to be safe
	KillTimer(nIDEvent);

	// But dont send msg after KillTimer
	//CWnd::OnTimer(nIDEvent);
}


void NTreeView::FindAllDirs(LPCWSTR pstr)
{
	CFileFind finder;

	// build a string with wildcards
	CString strWildcard(pstr);
	strWildcard += L"\\*.*";

	// start working for files
	BOOL bWorking = finder.FindFile(strWildcard);

	while (bWorking)
	{
		bWorking = finder.FindNextFile();

		// skip . and .. files; otherwise, we'd
		// recur infinitely!

		if (finder.IsDots())
			continue;

		// if it's a directory, recursively search it

		if (finder.IsDirectory())
		{
			CString str = finder.GetFilePath();
			TRACE(L"%s\n", (LPCWSTR)str);
			NTreeView::FindAllDirs(str);
		}
	}
	finder.Close();
}

void NTreeView::PostMsgCmdParamFileName(CString* fileName)
{
	if (GetSafeHwnd() && !m_bGeneralHintPostMsgDone)
	{
		BOOL res = PostMessage(WM_CMD_PARAM_GENERAL_HINT_MESSAGE, 0, 0);
		m_bGeneralHintPostMsgDone = TRUE;
	}

	if (GetSafeHwnd() && !m_bSelectMailFilePostMsgDone)
	{
		BOOL res = PostMessage(WM_CMD_PARAM_FILE_NAME_MESSAGE, 0, 0);
		m_bSelectMailFilePostMsgDone = TRUE;
	}
}

LRESULT NTreeView::OnCmdParam_FileName(WPARAM wParam, LPARAM lParam)
{
	SelectMailFile();
	// This is just one time message;
	// TO: Fix it
	//m_bGeneralHintPostMsgDone = FALSE;
	return 0;
}

LRESULT NTreeView::OnCmdParam_GeneralHint(WPARAM wParam, LPARAM lParam)
{
	MboxMail::ShowHint(HintConfig::GeneralUsageHint, GetSafeHwnd());
	// This is just one time message;
	// TODO: Fix it
	//m_bGeneralHintPostMsgDone = FALSE;
	return 0;
}

// Mail Folder Helpers

int NTreeView::GetFolderPath(HTREEITEM hItem, CString& mboxName, CString& parentPath)
{
	CString parentName = m_tree.GetItemText(hItem);
	//if (parentName.Compare(mboxName) == 0) return 1;
	parentPath.Append(parentName);
	HTREEITEM hParent = m_tree.GetParentItem(hItem);
	HTREEITEM hNextParent;
	while (hParent)
	{
		hNextParent = m_tree.GetParentItem(hParent);
		if (hNextParent == 0) // root
			break;
		parentName = m_tree.GetItemText(hParent);
		//if (parentName.Compare(mboxName) == 0) break;
		parentPath.Insert(0, L"\\");
		parentPath.Insert(0, parentName);
		hParent = hNextParent;
	}
	return 1;
}

int NTreeView::CreateEmptyFolder(HTREEITEM hItem)
{
	CString mboxFileNamePath = MboxMail::s_path;
	CString driveName;
	CString mboxFileDirectory;
	CString mboxFileNameBase;
	CString mboxFileNameExtention;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	NListView* pListView = 0;
	if (pFrame == 0) {
		return 0;
	}

	pListView = pFrame->GetListView();

	FileUtils::SplitFilePath(mboxFileNamePath, driveName, mboxFileDirectory, mboxFileNameBase, mboxFileNameExtention);

	CString mboxFolderName = mboxFileNameBase + mboxFileNameExtention;
	InputBox dlg(mboxFolderName);

	INT_PTR answer = dlg.DoModal();
	if (answer != IDOK)
	{
		return 0;
	}

	CString newFolderName = dlg.m_input;
	CString subFolderPath;  // path up to mboxFolderName
	int retpath = GetFolderPath(hItem, newFolderName, subFolderPath);
	if (!retpath)
	{
		return 1;
	}

	int retfolder = pListView->CreateEmptyFolder(driveName, mboxFileDirectory, mboxFolderName, subFolderPath, newFolderName);
	if (retfolder != ERROR_SUCCESS)
	{
		return 0;
	}

	//HTREEITEM newItem = m_tree.InsertItem(newFolderName, hItem, TVI_SORT);
	//HTREEITEM newItem = m_tree.InsertItem(newFolderName, 6,7,hItem, TVI_SORT); // TODO: insert into label store
	HTREEITEM newItem = InsertTreeItem(newFolderName, 6, 7, hItem, TVI_SORT);
	if (newItem == 0)
	{
		return 0;
	}

	int nId = m_labelInfoStore.GetNextId();
	m_tree.SetItemData(newItem, nId);

	LabelInfo* linfo = new LabelInfo(nId, newFolderName, newFolderName);  // MailFolder
	m_labelInfoStore.Add(linfo, nId);

	ExpandItem(hItem, TVE_EXPAND);

	CString fileDirPath;
	BOOL ret = FileUtils::CPathGetPath(MboxMail::s_path, fileDirPath);
	fileDirPath += L"\\Folders";
	FindAllDirs(fileDirPath);

	return 1;
}

int NTreeView::GetFolderPath(HTREEITEM hItem, CString& folderPath)
{
	HTREEITEM hNextParent = 0;
	HTREEITEM hNextNextParent = 0;
	HTREEITEM hParent = 0;

	while (hItem)
	{
		hParent = m_tree.GetParentItem(hItem);
		CString parentName;
		if (hParent) {
			parentName = m_tree.GetItemText(hParent);
			hNextParent = m_tree.GetParentItem(hParent);
		}

		CString nextParentName;
		if (hNextParent) {
			nextParentName = m_tree.GetItemText(hNextParent);
			hNextNextParent = m_tree.GetParentItem(hNextParent);
		}

		CString nextNextParentName;
		if (hNextNextParent) {
			nextNextParentName = m_tree.GetItemText(hNextNextParent);
		}

		CString itemName = m_tree.GetItemText(hItem);

		if (!folderPath.IsEmpty())
			folderPath.Insert(0, L"\\");
		folderPath.Insert(0, itemName);

		if (hNextNextParent == 0)
			break;

		hItem = hParent;
	}
	return 1;
}

int NTreeView::CreateFlatFolderList(HTREEITEM hItem, CArray<CString>& folderList)
{
	CString path;
	CString itemName;
	HTREEITEM hLeaveItem = 0;
	//TRACE(L"ItemTree!!\n");
	if (hItem == 0)
	{
		hItem = m_tree.GetRootItem();
	}

	CString rootName = m_tree.GetItemText(hItem);
	if (!m_tree.ItemHasChildren(hItem))
		return 1;

	hItem = m_tree.GetChildItem(hItem);
	CString firstChildName = m_tree.GetItemText(hItem);

	while (hItem != NULL)
	{
		CString folderPath;
		int ret = GetFolderPath(hItem, folderPath);
		folderList.Add(folderPath);
		TRACE(L"TreeFolderPath=%s\n", folderPath);

		if (m_tree.ItemHasChildren(hItem))
		{
			HTREEITEM hChild = m_tree.GetChildItem(hItem);
			CString childName = m_tree.GetItemText(hChild);
			CString parentName = m_tree.GetItemText(hItem);
			CreateFlatFolderList(hItem, folderList);
		}
		hItem = m_tree.GetNextSiblingItem(hItem);
		CString nextChildName = m_tree.GetItemText(hItem);
	}

	return 1;
}

int NTreeView::CreateFlatFolderList(CString& mboxFileName, CArray<CString>& folderList)
{
	int retval = -1;
	HTREEITEM hRoot = 0;
	HTREEITEM hItem = NTreeView::FindItem(hRoot, mboxFileName);
	if (hItem != 0)
	{
		retval = CreateFlatFolderList(hItem, folderList);
	}
	return retval;
}

BOOL NTreeView::OnEraseBkgnd(CDC* pDC)
{
	//BOOL ret = CWnd::OnEraseBkgnd(pDC);

	DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailArchiveList);

	m_tree.SetBkColor(color);

	CRect rect;
	GetClientRect(&rect);
	pDC->FillRect(&rect, &CBrush(color));

	return TRUE;
}

//  class CRegArray

void NTreeView::DetermineFolderPath(HTREEITEM hItem, CString& folderPath)
{
#if 1

	DWORD nId = (DWORD)m_tree.GetItemData(hItem);
	LabelInfo* linfo = m_labelInfoStore.Find(nId);
	if ((linfo->m_nodeType != LabelInfo::MailFolder) && (linfo->m_nodeType != LabelInfo::MailSubFolder))
		MboxMail::assert_unexpected();
	folderPath = linfo->m_mailFolderPath;
	folderPath.TrimRight(L"\\");
	folderPath.Append(L"\\");
#else
	folderPath.Empty();
	HTREEITEM hParent = hItem;
	while (hParent)
	{
		hItem = hParent;
		hParent = m_tree.GetParentItem(hParent);
	}
	if (hItem)
	{
		int retIndex = m_tree.GetItemData(hItem);
		m_folderArray.GetAt(retIndex, folderPath);
		//folderPath.TrimRight("\\");
	}
#endif
}

HTREEITEM NTreeView::DetermineRootFolder(HTREEITEM hItem)
{
	HTREEITEM hRoot = DetermineRootItem(hItem);
	return hRoot;
}


// Must work with folder hierachy such as Thunderbird
// Returns root file HTREEITEM or folder HTREEITEM if hItem represents folder
HTREEITEM NTreeView::DetermineRootMboxFile(HTREEITEM hItem)
{
	if (hItem == 0)
		return hItem;
#if 1
	DWORD nId = -1;
	LabelInfo* linfo = 0;
	while (hItem)
	{
		CString name = m_tree.GetItemText(hItem);
		TRACE(L"DetermineRootFolder: name=%s\n", name);

		nId = (DWORD)m_tree.GetItemData(hItem);
		linfo = m_labelInfoStore.Find(nId);
		_ASSERTE(linfo != 0);
		if (linfo == 0)
			return 0;

		if ((linfo->m_nodeType == LabelInfo::MailFolder) || (linfo->m_nodeType == LabelInfo::MailSubFolder))
			return hItem;

		hItem = m_tree.GetParentItem(hItem);
		_ASSERTE(hItem != 0);
	}
	return hItem;

#else

	DWORD nId = (DWORD)m_tree.GetItemData(hItem);
	LabelInfo *linfo = m_labelInfoStore.Find(nId);
	if ((linfo->m_nodeType == LabelInfo::MailFolder) || (linfo->m_nodeType == LabelInfo::MailSubFolder))
		return hItem;

	HTREEITEM hParentNext;
	HTREEITEM hParent = m_tree.GetParentItem(hItem);
	_ASSERTE(hParent != 0);

	while (hParent)
	{
		CString name = m_tree.GetItemText(hParent);
		TRACE(L"DetermineRootFolder: name=%s\n", name);

		nId = (DWORD)m_tree.GetItemData(hParent);
		linfo = m_labelInfoStore.Find(nId);
		if ((linfo->m_nodeType == LabelInfo::MailFolder) || (linfo->m_nodeType == LabelInfo::MailSubFolder))
			return hItem;

		hParentNext = m_tree.GetParentItem(hParent);
		if (hParentNext == 0)
			break;

		hItem = hParent;
		hParent = hParentNext;
	}
	return hItem;
#endif
}

// It returns root folder
HTREEITEM NTreeView::DetermineRootItem(HTREEITEM hItem)
{
	HTREEITEM hParent = hItem;
	while (hParent)
	{
		CString name = m_tree.GetItemText(hParent);
		TRACE(L"DetermineRootItem: name=%s\n", name);

		hItem = hParent;
		hParent = m_tree.GetParentItem(hParent);
	}
	return hItem;
}

CRegArray::CRegArray()
{
	m_section = CString(sz_Software_mboxview) + L"\\MailFolders";
	m_nMaxSize = 512;
}

CRegArray::CRegArray(CString& section)
{
	m_section = section;
	m_nMaxSize = 512;
}

CRegArray::~CRegArray()
{
	m_array.RemoveAll();
}

void CRegArray::Dump()
{
#ifdef _DEBUG
	CString key;

	TRACE(L"FOLDER LIST:\n");
	int ii = 0;
	for (ii = 0; ii < m_array.GetCount(); ii++)
	{
		TRACE(L"\t%d - \"%s\"\n", ii, m_array.GetAt(ii));
	}
#endif
}

int CRegArray::Find(CString& str)
{
	int ii = 0;
	for (ii = 0; ii < m_array.GetCount(); ii++)
	{
		if (m_array.GetAt(ii).CompareNoCase(str) == 0)
		{
			return ii;
		}
	}
	return -1;
}

int CRegArray::Add(CString& str)
{
	CString key;
	int indexEmpty = -1;
	int ii = 0;
	for (ii = 0; ii < m_array.GetCount(); ii++)
	{
		if (m_array.GetAt(ii).CompareNoCase(str) == 0)
		{
			return ii;
		}
	}
	for (ii = 0; ii < m_array.GetCount(); ii++)
	{
		if (m_array.GetAt(ii).IsEmpty())
		{
			break;
		}
	}

	key.Format(L"%d", ii);
	BOOL retval = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section, key, str);
	m_array.SetAtGrow(ii, str);

	return ii;
}

int CRegArray::Delete(CString& str)
{
	CString empty = L"";
	CString key;
	int ii = 0;
	for (ii = 0; ii < m_array.GetCount(); ii++)
	{
		if (m_array.GetAt(ii).CompareNoCase(str) == 0)
		{
			m_array.SetAt(ii, empty);
			key.Format(L"%d", ii);
			BOOL retval = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section, key, empty);
			return ii;
		}
	}
	return -1;
}

int CRegArray::Delete(int index, CString& str)
{
	CString empty = L"";
	CString key;
	if ((index >= 0) && (index < m_array.GetCount()))
	{
		if (m_array.GetAt(index).CompareNoCase(str) == 0)
		{
			m_array.SetAt(index, empty);
			key.Format(L"%d", index);
			BOOL retval = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section, key, empty);
			return index;
		}
		else
			return -1;
	}
	else
		return -1;
}

LSTATUS CRegArray::DeleteSectionKey()
{
	REGSAM samDesired = KEY_WOW64_32KEY;
	DWORD  Reserved = 0;

	if (m_section.IsEmpty())
		return ERROR_INVALID_HANDLE; // not exactly but should work

	LSTATUS retval = RegDeleteKeyEx(HKEY_CURRENT_USER, m_section, samDesired, Reserved);
	if (retval != ERROR_SUCCESS)
	{
		CString errorText = FileUtils::GetLastErrorAsString(retval);
	}
	return retval;
};

BOOL CRegArray::GetAt(int index, CString& str)
{
	str.Empty();
	if ((index >= 0) && (index < m_array.GetCount()))
	{
		str = m_array.GetAt(index);
		return TRUE;
	}
	else
		return FALSE;
}

int CRegArray::GetCount()
{
	return (int)m_array.GetCount();
}

BOOL CRegArray::CreateKey(CString& section, HKEY& hKey)
{
	DWORD	dwDisposition;
	HKEY	myKey;

	if (ERROR_SUCCESS == RegCreateKeyEx(HKEY_CURRENT_USER,
		(LPCWSTR)m_section, 0, NULL,
		REG_OPTION_NON_VOLATILE,
		KEY_WRITE, NULL, &myKey,
		&dwDisposition))
	{
		return TRUE;
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString();
		return FALSE;
	}
}


// TODO: should we call below upon exit ? see potential issue
BOOL CRegArray::SaveToRegistry()
{
	int ii;
	CString key;
	CString value;
	DWORD	dwDisposition;
	HKEY	myKey;

	//LSTATUS keyDeleted = CRegArray::DeleteSectionKey();
	// TODO: allow to delete or set key value to empty in registry by hand to resolve startup issues, hanging Mbox Viewer
	LSTATUS keyDeleted = ERROR_INVALID_HANDLE;

	if (ERROR_SUCCESS == RegCreateKeyEx(HKEY_CURRENT_USER,
		(LPCWSTR)m_section, 0, NULL,
		REG_OPTION_NON_VOLATILE,
		KEY_WRITE, NULL, &myKey,
		&dwDisposition))
	{
		int index = 0;
		for (ii = 0; ii < m_array.GetCount(); ii++)
		{
			CString value = m_array.GetAt(ii);

			if (keyDeleted == ERROR_SUCCESS)
			{
				if (!value.IsEmpty())
				{
					key.Format(L"%d", index);
					index++;
				}
				else
					continue;
			}
			else
			{
				key.Format(L"%d", ii);
			}

			LSTATUS sts = RegSetValueEx(myKey, key, 0, REG_SZ, (CONST BYTE*)(LPCWSTR)value, value.GetLength() * 2 + 2); // FIXME
			if (sts != ERROR_SUCCESS)
			{
				DWORD err = GetLastError();
				CString errText = FileUtils::GetLastErrorAsString();
				// TODO: break or ignore ??
			}
		}
		RegCloseKey(myKey);
		return TRUE; // ??? FIXME
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString();
		return FALSE;
	}
}

BOOL CRegArray::LoadFromRegistry()
{
	BOOL ret = LoadFromRegistry(m_array);
	return ret;
}

BOOL CRegArray::LoadFromRegistry(CSArray& ar)  // FIXMEFIXME
{
	BOOL ret = TRUE;
	CString fileSection;
	if (CProfile::IsRegistryConfig())
	{
		ret = CRegArray::LoadFromRegistry_registry(ar);
	}
	else
	{
		ret = CRegArray::LoadFromConfigFile(ar);
	}
	return ret;
}


BOOL CRegArray::LoadFromConfigFile(CSArray& ar)  // FIXMEFIXME
{
	CString configSection;
	CString registrySection = m_section;
	ConfigTree::GetSection(registrySection, configSection);
	if (configSection.IsEmpty())  // not UMBoxView
		return FALSE;

	ConfigTree* config = CProfile::GetConfigTree();

	ConfigNode * node = config->ConfigTree::FindNode(configSection);
	if ((node == 0) || !node->m_isSectionNode)
	{
		return 0;
	}

	int ii = 0;
	ConfigNode * childNode;
	INT_PTR indx;
	for (childNode = node->m_configList.first(); childNode != 0; childNode = node->m_configList.next(childNode))
	{
		indx = _ttoi(childNode->m_name);
		ar.SetAtGrow(indx, childNode->m_value);
	}
	return TRUE;
}


BOOL CRegArray::LoadFromRegistry_registry(CSArray& ar)  // FIXMEFIXME
{
	const int datasize = 8196;
	DWORD	size = datasize / 2;
	unsigned char data[datasize];
	int ii;
	data[0] = 0;
	CString key;
	CString val;
	HKEY	myKey;
	CString	result = L"";

	LSTATUS res;
	if (ERROR_SUCCESS == RegOpenKeyEx(HKEY_CURRENT_USER, (LPCWSTR)m_section, NULL, KEY_READ | KEY_QUERY_VALUE, &myKey))
	{
		int missingKeyCnt = 20;  // this is work around
		for (ii = 0; ii < m_nMaxSize; ii++)
		{
			key.Format(L"%d", ii);
			size = datasize / 2;
			res = RegQueryValueEx(myKey, (LPCWSTR)key, NULL, NULL, data, &size);
			if (res == ERROR_SUCCESS)
			{
				result = (wchar_t*)data;
				val = result;
				ar.SetAtGrow(ii, val);
			}
			else if (missingKeyCnt-- < 0)
			{
				break;
			}
		}
		RegCloseKey(myKey);
		return TRUE;  // TODO:: fixit ?? FIXME
	}
	else
	{
		DWORD err = GetLastError();
		CString errText = FileUtils::GetLastErrorAsString();
		return FALSE;
	}
}

void CRegArray::Optimize()
{
	;
}

void NTreeView::RemoveFileFromTreeView(HTREEITEM hItem, FileSizeMap& fileSizes)
{
	HWND h = GetSafeHwnd();
	MboxMail::ShowHint(HintConfig::MessageRemoveFileHint, h);

	ArchiveFileInfo info;
	CString mailFile = m_tree.GetItemText(hItem);
	BOOL found = fileSizes.Lookup(mailFile, info);
	if (found)
	{
		fileSizes[mailFile].bShow = 0;
		SaveData(hItem);
		BOOL ret = m_tree.DeleteItem(hItem);
	}
	return;
}

BOOL NTreeView::RefreshFolder(HTREEITEM hItem)
{
	CString path = MboxMail::GetLastPath();

#if 1
	DWORD index = (DWORD)m_tree.GetItemData(hItem);
	CString iteName = m_tree.GetItemText(hItem);
	LabelInfo* linfo = m_labelInfoStore.Find(index);

	if (linfo->m_nodeType != LabelInfo::MailFolder)
		MboxMail::assert_unexpected();


	FolderInfo* finfo = 0;

	CString folderPath = linfo->m_mailFolderPath;
	folderPath.TrimRight(L"\\");
	folderPath.Append(L"\\");

	finfo = m_globalFolderInfoDB.Find(&folderPath);
	if (finfo == 0)
	{
		MboxMail::assert_unexpected();
	}

	int folderType = 0;
	if (finfo)
		folderType = finfo->m_folderType;

	BOOL ret = DeleteFolder(hItem);
	if (folderType == FolderInfo::MailRootFolderNode)
	{
		CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
		if (pFrame)
		{
			MboxMail::SetLastPath(folderPath);

			pFrame->OpenRootFolderAndSubfolders(folderPath);
		}
	}
	else
	{
		BOOL expand = TRUE;
		DoOpen(path, expand);
	}
#else

	MboxMail::SetLastPath(path);

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	int paneId = 0;
	CString sText;
	CString fname;
	FileUtils::CPathStripPath(path, fname);
	sText.Format("Refreshing mail folder %s ...", fname);
	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	HTREEITEM hParent = TVI_ROOT;
	BOOL selectFolder = TRUE;
	BOOL unconditionalFolderInsert = FALSE;
	BOOL expand = FALSE;
	HTREEITEM hItem2 = FillCtrl(hParent, selectFolder, unconditionalFolderInsert, expand);

	SortChildren(0);

	HTREEITEM hFolder;
	if (hFolder = HasFolder(path))
	{
		SortChildren(hFolder);
		ExpandItem(hFolder, TVE_EXPAND);
	}

	sText.Format("Ready");
	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
#endif
	return TRUE;
}


int NTreeView::OpenHiddenFiles(HTREEITEM hItem, FileSizeMap& fileSizes, BOOL isSubFolder)
{
	CCheckListBoxDlg dlg;
	dlg.m_title = L"Select Files to Restore";
	ResHelper::TranslateString(dlg.m_title);

	FileSizeMap::CPair* pCurVal;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	CString pathLast = MboxMail::GetLastPath();

	DWORD nId = (DWORD)m_tree.GetItemData(hItem);

	LabelInfo* linfo = m_labelInfoStore.Find(nId);
	CString path;

	if (linfo->m_nodeType == LabelInfo::MailSubFolder)
		path = linfo->m_mailFolderPath;
	else
		path = m_folderArray.m_array.GetAt(nId);

	if (path.Compare(pathLast))
		int deb = 1;

	path.TrimRight(L"\\");
	path.AppendChar(L'\\');

	CString mboxFileName;
	pCurVal = fileSizes.PGetFirstAssoc();
	while (pCurVal != NULL)
	{
		if (pCurVal->value.bShow == 0)
		{
			mboxFileName = path + pCurVal->key;
			if (FileUtils::PathFileExist(mboxFileName))
			{
				dlg.m_InList.Add(pCurVal->key);
			}
		}
		pCurVal = fileSizes.PGetNextAssoc(pCurVal);
	}
	//
	int nResponse = (int)dlg.DoModal();
	if (nResponse == IDOK)
	{
		int i;
		for (i = 0; i < dlg.m_OutList.GetCount(); i++)
		{
			CString& s = dlg.m_OutList[i];
			CString filePath = path + s;
			fileSizes[s].bShow = 1;

			//HTREEITEM hItemRet = m_tree.InsertItem(s, 8, 9, hItem);
			HTREEITEM hItemRet = InsertTreeItem(s, 8, 9, hItem);

			int nId = m_labelInfoStore.GetNextId();
			m_tree.SetItemData(hItemRet, nId);


			LabelInfo* linfo = new LabelInfo(nId, filePath);  // MailFile
			m_labelInfoStore.Add(linfo, nId);


			// TODO: resolve later; manual action is required for now
			//DisplayGmailLabels(hItem);
			int deb = 1;
		}
		SaveData(hItem);
		this->SortChildren(hItem);
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: Place code here to handle when the dialog is
		//  dismissed with Cancel
	}
	return 1;
}

BOOL NTreeView::RecreateGmailLabels(HTREEITEM hItem)
{
	int ret = CreateGmailLabelFiles(hItem);
	ret = DisplayGmailLabels(hItem);

	m_labelInfoStore.Print();
	return ret;
}


int NTreeView::IsGmailLabelFile(CString& mboxFile)
{
	static const char* cLabels = "x-gmail-labels";
	static const int cLabelsLen = istrlen(cLabels);

	const int bufflen = 1024;
	unsigned char buff[bufflen + 1];

	CFile fp;
	CFileException* pError = NULL;
	if (!fp.Open(mboxFile, CFile::modeRead | CFile::shareDenyWrite, pError))
	{
		DWORD err = GetLastError();
		TRACE(L"IsGmailLabelFile:: Open Mail File failed err=%ld\n", err);
		return -1;
	}

	UINT readbytes = fp.Read(buff, bufflen);
	buff[readbytes] = 0;

	char* p = strstr((char*)buff, cLabels);

	BOOL bCaseSens = FALSE;
	int pos = g_tu.BMHSearch(buff, readbytes, (unsigned char*)cLabels, cLabelsLen, bCaseSens, 0);
	if (pos >= 0)
		return 1;
	else
		return 0;
}


// Parse mbox file and create label files and Windows folders
// TODO: Current solution is quite complex due to use of files. 
// File names suppose to represent label names, etc. but number of characters are illegal in the file names.
// TODO: May have to recreate label tree first and resolve valid file name issue next  FIXME
//
// Parse mbox  file and create a Gmail label database.
// MBox Viewer doesn't maintain a single global database that would contain all mail content and all relevant relationships.
// Instead, MBox Viewer creates a distributed database consisting of several files.
// Label database is represented as a hierarchy of folders and files.
// Example of label: MyFamily/MySons/John
// MyFamily and MySons are represented by Folders and associated MyFamily.mboxlist (under MyFamily folder) 
// and MySons.mboxlist (under MySons folder) index files.
// John, the leaf node, is represented by the index file John.mboxlist under the MySons folder.
// xxx.mboxlist files are basically arrays of indexes representing mail position in the corresponding mbox file.
//
// Using files introduces complexity since files names can't always be the same as label names and have to be remapped.
// Example of two label paths:
// A label:   A/<home>/kitchen
// B label:   A/"home"/kitchen
// A and B label paths have to be renamed since <,>," chracters are not valid file name characters
// We can replace illegal characters with '_' character for example
// A label:   A/_home_/kitchen
// B label:   A/_home_/kitchen
// Simple replacement created two identical paths that have to be remapped further, say by appending unique ID
// A label:   A/_home_/kitchen
// B label:   A/_home_0/kitchen
// All this can possibly be avoiuded by having global database -:)  Maybe one day it will lhappen
//
// TODO: Inestigate whether the current implementation of the file based approach can be simplified FIXME
//

int  NTreeView::CreateGmailLabelFiles(HTREEITEM hItem)
{
	HWND h = GetSafeHwnd();
	//MboxMail::ShowHint(HintConfig::MessageRemoveFileHint, h);

	static const char* cLabels = "x-gmail-labels";
	static const int cLabelsLen = istrlen(cLabels);
	static const char* cCategory = "category ";
	static const int cCategoryLength = istrlen(cCategory);
	static const char* cOsobiste = "osobiste";
	static const int cOsobisteLen = istrlen(cOsobiste);

	CStringA comma = ",";

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	NListView* pListView = 0;
	if (pFrame)
		pListView = pFrame->GetListView();

	if (pListView == 0) {
		MboxMail::assert_unexpected();
		return -1;
	}

	m_mappedToLabelHT->clear();
	ClearLabelHT();

	CString mailFile = m_tree.GetItemText(hItem);
	int mailCnt = MboxMail::s_mails.GetCount();

	const int buffSmallSize = 2048;  // TODO: make small and enlarge and retry if needed  // FIXME
	SimpleString buff(buffSmallSize + 1);
	CStringA validFileName;
	CStringA rawLabel;
	MboxMail* m;
	CStringA labels;
	CStringA val;
	CStringA labelList;
	CStringA labelsStr;
	CStringA v;
	CStringA el;
	CStringA line;
	int ii; // reusable for not nested loops
	int k; // for mail loop
	int j; // for label loop
	int i; // for sublabel loop
	CStringA delim = "/";
	CStringA mappedEl;  // valid as part of file name
	CStringA mappedV;  // kind of path of mappedEl
	CStringA mappedLabel;  // mapped label path
	CStringA rawLabelFolder;
	CStringA charset;
	//
	CString mappedElW;  // valid as part of file name
	CString mappedVW;  // kind of path of mappedEl
	CString mappedLabelW;  // mapped label path
	CString rawLabelFolderW;
	DWORD error;

	int maxNoLabelCnt = 100;
	BOOL foundLabels = FALSE;

	char* p = buff.Data();
	char* e = p + buff.Count();

	CFile fp;
	CFileException* pError = NULL;
	if (!fp.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite, pError))
	{
		DWORD err = GetLastError();
		TRACE(L"CreateGmailLabelFiles:: Open Mail File failed err=%ld\n", err);
		return -1;
	}

	if (MboxMail::s_mails.GetCount() != MboxMail::s_mails_ref.GetCount())
		int deb = 1;

	CString labelsCachePath;
	CString rootPrintSubFolder = L"LabelCache";
	CString errorText;
	CString targetPrintSubFolder;
	BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, labelsCachePath, errorText);
	if (retval == FALSE) {
		MboxMail::assert_unexpected();
		TRACE(L"CreateGmailLabelFiles: CreateCachePath LabelCache failed\n");
		return -1;
	}

	int noLabelCnt = 0;
	m_labelSeqNumb = 0;
	BOOL TraceLabels = FALSE;
	//CString searchText = "We don't normally send surveys";
	//searchText = "Re: [chromium-dev] Mac builds";
	UINT CP_ASCII = 20127;
	UINT CP_ANSI = 0;  // or 
	CP_ANSI = GetACP();  // local CP
	BOOL ret;
	for (k = 0; k < mailCnt; k++)
	{
		//m_s = MboxMail::s_mails[k];
		m = MboxMail::s_mails_ref[k];

		ret = m->GetBodySS(fp, &buff, buffSmallSize);

		p = buff.Data();
		e = p + buff.Count();

		labelList.Empty();
		DWORD labelCharsetId = CP_UTF8;   // ?? or set to 0, i.e ansi, or 20127 , i.e ascii  // FIXME
		while (p < e)
		{
			// find "x-gmail-labels" header field
			if (TextUtilsEx::strncmpUpper2Lower(p, e, cLabels, cLabelsLen) == 0)
			{
				if (TraceLabels)
					int deb = 1;
				foundLabels = TRUE;
				p = MimeParser::GetMultiLine(p, e, line);
				labelsStr = line.Mid(cLabelsLen);
				labelsStr.TrimLeft(": ");
				labelsStr.TrimRight(" ");

				//TRACE("Labels=%s\n", labelsStr);
				UINT charsetId;
				UINT toCharacterId = 20127;  // us-ascii  toCharacterId is ignored currently
				CStringA decodedLabelsStr = TextUtilsEx::DecodeString(labelsStr, charset, charsetId, toCharacterId, error);

				labelList = decodedLabelsStr;
				labelCharsetId = charsetId;
				break;
			}
			p = MimeParser::EatNewLine(p, e);
		}

		// we expect to find "x-gmail-labels" header field in all messages. Count number of emails wothout "x-gmail-labels" field
		if (foundLabels == FALSE)
		{
			noLabelCnt++;
			int deb = 1;
		}

		if ((foundLabels == FALSE) && val.IsEmpty() && (k > maxNoLabelCnt))  // it looks like it is not gmail  file
		{
			fp.Close();
			TRACE(L"CreateGmailLabelFiles:: Didn't find labels\n");
			return 1;
		}

		if (labelCharsetId != CP_UTF8)
		{
			// Get label as wide char string
			CString labelListW;
			BOOL retW = TextUtilsEx::CodePage2WStr(&labelList, labelCharsetId, &labelListW, error);

			const char* lchar = (LPCSTR)labelList;

			UINT outCodePage = CP_UTF8;
			BOOL retW2A = TextUtilsEx::WStr2CodePage((LPCWSTR)labelListW, labelListW.GetLength(), outCodePage, &labelList, error);
			labelCharsetId = CP_UTF8;
		}

		// tokenize value for "x-gmail-labels" header field
		CStringArrayA a;
		char del = ',';
		int ret = TextUtilsEx::TokenizeA(labelList, a, del);

		// remove duplicate labels
		int firstw = 0;
		while (firstw < a.GetSize())
		{
			el = a.ElementAt(firstw);
			for (ii = firstw + 1; ii < a.GetSize(); ii++)
			{
				v = a.ElementAt(ii);
				if (el.Compare(v) == 0)
				{
					a.RemoveAt(ii);
					break;
				}
			}
			firstw++;
		}

		// iterate all labels and add mail ptr to respectfull list/array
		for (j = 0; j < a.GetSize(); j++)
		{
			v = a.ElementAt(j);
			v.Trim();
			//TRACE("|%s|\n", v);

			// Special handling of Gmail "Category xxx" labels such as
			// "Categor Social", "Category Promotions", "Category Forums"
			// They are mapped to "Category/Social", "Category/Promotions", "Category/Forums"
			//
			if (v.GetLength() >= cCategoryLength)
			{
				char* p = (char*)(LPCSTR)v;
				char* e = p + v.GetLength();
				if (TextUtilsEx::strncmpUpper2Lower(p, e, cCategory, cCategoryLength) == 0)
				{
					// TODO: enhance to remove all spaces
					// Below assumes single space only   // FIXME
					v.SetAt(cCategoryLength - 1, '/');
				}
			}

			CStringArray va;
			TextUtilsEx::SplitStringA(v, delim, va);

			if (TraceLabels) {
				TRACE("CreateGmailLabelFiles: Before Split Label v=\"%s\"\n", v);
				TextUtilsEx::TraceStringArrayW(va);
			}

			// Multiple sequential slash chracters will be replaced with just one
			CString validLabel = v;
			v.Empty();
			int valen = (int)va.GetSize();
			for (ii = 0; ii < valen; ii++)
			{
				el = va.ElementAt(ii);
				el.Trim();
				if (ii == 0)
					v += el;
				else
					v += "/" + el;
			}

			validFileName = v;
			rawLabel = v;

			GmailLabel* gLabel = m_labelHT->find(&rawLabel);
			if (gLabel)
			{
				gLabel->m_ptrList->Add(m);
				continue;
			}

			v.Empty();
			//CString mappedEl;  // valid as part of file name
			//CString mappedV;  // kind of path of mappedEl
			//CString mappedLabel;  // mapped label path

			int sublabelCnt = (int)va.GetSize();
			for (i = 0; i < sublabelCnt; i++)
			{
				el = va.ElementAt(i);
				el.Trim();
				mappedEl = el;
				BOOL bReplaceWhiteWithUnderscore = FALSE;
				BOOL extraValidation = TRUE;
				FileUtils::MakeValidFileNameA(mappedEl, bReplaceWhiteWithUnderscore, extraValidation);
				if (i == 0)
				{
					v = el;
					mappedV = mappedEl;
					mappedLabel = mappedEl;
					gLabel = m_labelHT->find(&v);
					if (gLabel)
					{
						mappedV = gLabel->m_mappedToLabel;
						mappedLabel = gLabel->m_mappedLabelPath;
						continue;
					}
				}
				else
				{
					mappedV = v + "/" + mappedEl;
					mappedLabel += "/" + mappedEl;
					v += "/" + el;
					gLabel = m_labelHT->find(&v);
					if (gLabel)
					{
						mappedV = gLabel->m_mappedToLabel;
						mappedLabel = gLabel->m_mappedLabelPath;
						continue;
					}
				}

				rawLabel = v;

				GmailLabel* gMappedLabel = m_mappedToLabelHT->find(&mappedV);
				if (gMappedLabel)
					int deb = 1;

				GmailLabel* gLabel = m_labelHT->find(&v);
				if (gLabel)
				{
					//_ASSERTE(gMappedLabel);
					if (i == (sublabelCnt - 1))
					{
						gLabel->m_ptrList->Add(m);
						//continue;
					}
					int deb = 1;
				}
				else
				{
					if (gMappedLabel)
					{
						//TRACE("%u\n", glabel);
						if (gLabel == 0)
						{
							CString seqNumb;
							seqNumb.Format(L" %d", m_labelSeqNumb++);
							mappedLabel += seqNumb;

							GmailLabel* mLabel = m_mappedToLabelHT->find(&v);
							if (mLabel)
								MboxMail::assert_unexpected();
						}
						else
							int deb = 1;
					}


					// Currently labelCharsetId is always set to CP_UTF8
					gLabel = new GmailLabel(rawLabel, labelCharsetId, mappedV, mappedLabel);
					m_labelHT->insert(&rawLabel, gLabel);
					m_mappedToLabelHT->insert(&mappedV, gLabel);

					if (i == (sublabelCnt - 1))
					{
						gLabel->m_ptrList->Add(m);
					}
				}
			}
			// va array no longer needed
			int deb = 1;
		}
	}
	fp.Close();

	GmailLableMapType::IHashMapIter iter = m_labelHT->first();

	// find the size of largest list of mails
	int maxCnt = 0;
	for (; !m_labelHT->last(iter); )
	{
		int cnt = iter.element->m_ptrList->Count();
		if (cnt > maxCnt)
			maxCnt = cnt;
		m_labelHT->next(iter);
	}

	MboxMail::s_mails_label.SetSizeKeepData(maxCnt);

	// iterate labels and create folder and .mboxlist files for each label
	iter = m_labelHT->first();
	int totalCnt = 0;
	for (; !m_labelHT->last(iter); )
	{
		iter.element->m_ptrList->Assert();
		TRACE("Name=%s  cnt=%d\n", iter.element->GetLabelA(), iter.element->m_ptrList->Count());
		CString labelW;
		TextUtilsEx::UTF82WStr(&iter.element->GetLabelA(), &labelW, error);

		int cnt = iter.element->m_ptrList->Count();
		totalCnt += cnt;

		MboxMail::s_mails_label.SetSizeKeepData(cnt);
		MboxMail* mb;
		for (ii = 0; ii < cnt; ii++)
		{
			mb = iter.element->m_ptrList->Get(ii);
			MboxMail::s_mails_label[ii] = iter.element->m_ptrList->Get(ii);
		}

		_ASSERTE(MboxMail::s_mails_label.GetCount() <= mailCnt); // FIXMEFIXME it is normal to have multiple labels per one mail


		CStringA labelsCachePathA = labelsCachePath;
		int r = pListView->SaveAsLabelFile(&MboxMail::s_mails_label, labelsCachePath, iter.element->GetLabelA(), iter.element->GetCodePage(), iter.element->m_mappedLabelPath, errorText);
		if (r < 0) {
			MboxMail::assert_unexpected();
			TRACE(L"CreateGmailLabelFiles: SaveAsLabelFile failed\n");
			return -1;
		}

		m_labelHT->next(iter);
	}
	return 1;
}

// Traverse label database, i.e label file hierachy, and insert all labels into CtreeCtrl
int NTreeView::DisplayGmailLabels(HTREEITEM hItem)
{
	CString name;
	LabelInfo* linfo;

	// add labels to Tree
	DWORD nId = (DWORD)m_tree.GetItemData(hItem);

	linfo = m_labelInfoStore.Find(nId);
	if (linfo == 0)
		return -1;

	if (!linfo->m_filePath.IsEmpty())
	{
		name = m_tree.GetItemText(hItem);
		//nId = m_tree.GetItemData(hItem);

		//linfo = m_labelInfoStore.Find(nId);

		CString labelsCachePath;
		CString rootPrintSubFolder = L"LabelCache";
		CString errorText;
		CString targetPrintSubFolder;

		//MboxMail::s_path = linfo->m_filePath;
		MboxMail::SetMboxFilePath(linfo->m_filePath);

		CString driveName;
		CString directory;
		CString fileNameBase;
		CString fileNameExtention;
		FileUtils::SplitFilePath(linfo->m_filePath, driveName, directory, fileNameBase, fileNameExtention);

		CString lastPath = driveName + directory;

		MboxMail::SetLastPath(lastPath);

		BOOL retval = MboxMail::GetCachePath(rootPrintSubFolder, targetPrintSubFolder, labelsCachePath, errorText, &linfo->m_filePath);
		if (retval == FALSE) {
			MboxMail::assert_unexpected();
			return -1;
		}
		if (!labelsCachePath.IsEmpty())
		{
			ShowGmailLabels(hItem, labelsCachePath, linfo->m_filePath);

			CString inbox(L"Inbox");
			CString sent(L"Sent");
			HTREEITEM hParent = m_tree.GetParentItem(hItem);
			this->SortChildren(hItem);
			MoveLabelItem(linfo->m_filePath, sent);
			MoveLabelItem(linfo->m_filePath, inbox);
		}
	}
	//_ASSERTE(totalCnt >= mailCnt);
	return 1;
}

GmailLabel::GmailLabel(CStringA& label, DWORD labelCodePage, CStringA& mappedToLabel, CStringA& mappedLabelPath)
{
	m_label = label;
	m_codePage = labelCodePage;
	m_mappedLabelPath = mappedLabelPath;
	m_mappedToLabel = mappedToLabel;
	m_ptrList = new MySimpleDeque;
}

GmailLabel::~GmailLabel()
{
	delete m_ptrList;
}

CString GmailLabel::GetLabelW()
{
	DWORD error = 0;
	CString label;
	BOOL ret = TextUtilsEx::CodePage2WStr(&m_label, m_codePage, &label, error);
	return label;
}

CStringA& GmailLabel::GetLabelA()
{
	return m_label;
}

MySimpleDeque::MySimpleDeque()
{
	m_ar = 0;
	m_arcnt = 0;
}

MySimpleDeque::~MySimpleDeque()
{
	std::vector <LabelArray*>::iterator it;
	LabelArray* l;
	for (it = m_arList.begin(); it != m_arList.end(); it++)
	{
		l = *it;
		delete l;
	}
	delete m_ar;
}

BOOL MySimpleDeque::Assert()
{
	int cnt = 0;
	if (m_ar == 0)
		return TRUE;

	std::vector <LabelArray*>::iterator it;
	LabelArray* l;
	for (it = m_arList.begin(); it != m_arList.end(); it++)
	{
		l = *it;
		//_ASSERTE(m_arcnt == LSIZE);
	}
	int lcnt = Count();
	int mailCnt = MboxMail::s_mails.GetCount();
	//_ASSERTE(lcnt <= mailCnt);
	return TRUE;
}

void  MySimpleDeque::Clear()
{
	m_arcnt = 0;
	std::vector <LabelArray*>::iterator it;
	LabelArray* l;
	for (it = m_arList.begin(); it != m_arList.end(); it++)
	{
		l = *it;
		delete l;
	}
	m_arList.clear();
}

int MySimpleDeque::Count()
{
	int cnt = 0;
	if (m_ar == 0)
		return 0;

	cnt = (int)m_arList.size() * LSIZE + m_arcnt;

	//this->Assert();

	return cnt;
}

void MySimpleDeque::Add(MboxMail* m)
{
	if (m_ar == 0)
	{
		m_ar = new LabelArray;
		m_arcnt = 0;
	}

	if (m_arcnt >= LSIZE)
	{
		m_arList.push_back(m_ar);
		m_ar = new LabelArray;
		m_arcnt = 0;
	}
	(*m_ar)[m_arcnt++] = m;
}

MboxMail* MySimpleDeque::Get(int position)
{
	LabelArray* ar;
	int arpos;

	int arListCount = (int)m_arList.size() * LSIZE;
	int arCount = arListCount + m_arcnt;
	if (position >= arCount)
		return 0; // error

	if (position < arListCount)
	{
		int index = position / LSIZE;
		ar = m_arList[index];
		arListCount = index * LSIZE;
		arpos = position - arListCount;
	}
	else
	{
		ar = m_ar;
		arpos = position - arListCount;
	}

	if (arpos >= LSIZE)
		int deb = 1;

	MboxMail* m = (*ar)[arpos];

	if (m == 0)
		int deb = 1;

	if (m->m_startOff)
		int deb = 1;

	return m;
}

// recursive
// Iterate folder and sub-folders, Insert Label into TreeCtrl, create xx.mboxlist for folders, add to label store
int NTreeView::ShowGmailLabels(HTREEITEM hItem, CString& listFilePath, CString& dataFilePath)
{
	m_treeCtrlInfoDB.Clear();
	int ret = ShowGmailLabels_internal(hItem, listFilePath, dataFilePath);
	m_treeCtrlInfoDB.Clear();
	return ret;
}


int NTreeView::ShowGmailLabels_internal(HTREEITEM hItem, CString& listFilePath, CString& dataFilePath)
{
	CString errorText;
	CFileFind finder;
	if (hItem == 0)
		return 1;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	NListView* pListView = 0;
	if (pFrame)
		pListView = pFrame->GetListView();

	if (pListView == 0)
	{
		MboxMail::assert_unexpected();
		return -1;
	}

	CString itemName = m_tree.GetItemText(hItem);

	// build a string with wildcards
	CString strWildcard(listFilePath);
	strWildcard += L"\\*.*";

	// start working for files
	BOOL bWorking = finder.FindFile(strWildcard);

	// NTreeView::HasLabel(hItem, label);  used below but it is extremely ineffcient trying to insert lets say 20,000 labels
	// We should not try to built the tree with 20,000 items in the first place but it can happen in case of mistake
	// Implemented m_treeCtrlInfoDB instead to keep track of file path to HTREEITEM

	HTREEITEM hChild = m_tree.GetChildItem(hItem);
	if (hChild)
		int deb = 1;

	while (bWorking)
	{
		bWorking = finder.FindNextFile();

		// skip . and .. files; otherwise, we'd
		// recur infinitely!

		if (finder.IsDots())
			continue;

		CString fPath = finder.GetFilePath();
		CString fName = finder.GetFileName();
		CString label;
		FileUtils::GetFileBaseName(fName, label);
		//TRACE("fPath=|%s| fName=|%s|\n", fPath, fName);

		// if it's a directory, recursively search it
		if (finder.IsDirectory())
		{
			CString folderPath = finder.GetFilePath();
			CString folderName = finder.GetFileName();
			CString label;

			label = folderName;
			//TRACE("%s\n", (LPCWSTR)folderPath);

			CString emptyPath = folderPath + L".mboxlist";

			hashsum_t hashsum = TreeCtrlInfoDB::GetHashsum(&emptyPath);
			TreeCtrlInfo* tinfo = m_treeCtrlInfoDB.Find(&emptyPath, hashsum);

			HTREEITEM found_hItem = 0;
			if (tinfo)
				found_hItem = tinfo->m_hItem;

			//HTREEITEM found_hItem = NTreeView::HasLabel(hItem, label);
			if (found_hItem == 0)
			{
				if (tinfo)
					int deb = 1;

				CString folderNm = folderName;
				DWORD gLabelCodePage = 0;
				if (FileUtils::PathFileExist(emptyPath))
				{
					CStringA gLabel;
					int retv = pListView->GetLabelFromLabelListFile_v2(emptyPath, gLabel, gLabelCodePage, dataFilePath);
					if (retv < 0)
					{
						MboxMail::assert_unexpected();
						finder.Close();
						return -1;
					}

					DWORD error = 0;
					CString label;
					BOOL ret = TextUtilsEx::CodePage2WStr(&gLabel, gLabelCodePage, &folderNm, error);
				}

				HTREEITEM newItem = InsertTreeItem(folderNm, 6, 7, hItem);
				if (newItem == 0)
				{
					MboxMail::assert_unexpected();
					finder.Close();
					return -1;
				}

				int nId = m_labelInfoStore.GetNextId();
				m_tree.SetItemData(newItem, nId);

				//CString emptyPath = folderPath + ".mboxlist";

				tinfo = new TreeCtrlInfo(emptyPath, newItem, TreeCtrlInfo::MailFolder);
				m_treeCtrlInfoDB.Add(hashsum, tinfo);

				if (!FileUtils::PathFileExist(emptyPath))
				{
					// Should never be true
					// All files are created in CreateGmailLabelFiles() now 12/26/2022
					MboxMail::assert_unexpected();
				}

				LabelInfo* linfo = new LabelInfo(nId, dataFilePath, folderNm, emptyPath); // MailLabel
				m_labelInfoStore.Add(linfo, nId);

				int r = NTreeView::ShowGmailLabels_internal(newItem, folderPath, dataFilePath);
				if (r < 0) {
					MboxMail::assert_unexpected();
					finder.Close();
					return -1;
				}
			}
			else
			{
				// Folder related item was created by file
				// It depends on ordre of file and folder discovery while traversing
				// It is possible we never be here
				if (tinfo == 0)
					int deb = 1;

				int r = NTreeView::ShowGmailLabels_internal(found_hItem, folderPath, dataFilePath);
				if (r < 0) {
					MboxMail::assert_unexpected();
					finder.Close();
					return -1;
				}
			}
		}
		else // !finder.IsDirectory())
		{
			CString filePath = finder.GetFilePath();
			CString fileName = finder.GetFileName();
			//TRACE("%s\n", (LPCWSTR)filePath);

			CString label;

			FileUtils::GetFileBaseName(fileName, label);

			CStringA gLabel;
			DWORD gLabelCodePage = 0;
			int retv = pListView->GetLabelFromLabelListFile_v2(filePath, gLabel, gLabelCodePage, dataFilePath);
			if (retv < 0)
			{
				MboxMail::assert_unexpected();
				finder.Close();
				return -1;
			}

			hashsum_t hashsum = TreeCtrlInfoDB::GetHashsum(&filePath);
			TreeCtrlInfo* tinfo = m_treeCtrlInfoDB.Find(&filePath, hashsum);

			HTREEITEM found_hItem = 0;
			if (tinfo)
				found_hItem = tinfo->m_hItem;

			//HTREEITEM found_hItem = NTreeView::HasLabel(hItem, label);
			if (found_hItem == 0)
			{
				if (tinfo)
					int deb = 1;

				//CString gLabelW = gLabel;
				CString gLabelW;

				DWORD error = 0;
				CString label;
				BOOL ret = TextUtilsEx::CodePage2WStr(&gLabel, gLabelCodePage, &gLabelW, error);

				HTREEITEM newItem = InsertTreeItem(gLabelW, 6, 7, hItem);
				if (newItem == 0)
				{
					MboxMail::assert_unexpected();
					finder.Close();
					return -1;
				}

				tinfo = new TreeCtrlInfo(filePath, newItem, TreeCtrlInfo::MailLabel);
				m_treeCtrlInfoDB.Add(hashsum, tinfo);

				int nId = m_labelInfoStore.GetNextId();
				m_tree.SetItemData(newItem, nId);

				LabelInfo* linfo = new LabelInfo(nId, dataFilePath, gLabelW, filePath);  // MailLabel
				m_labelInfoStore.Add(linfo, nId);
			}
			else
			{
				if (tinfo == 0)
					int deb = 1;

				DWORD nId = (DWORD)m_tree.GetItemData(found_hItem);
				LabelInfo* linfo = m_labelInfoStore.Find(nId);

				if (linfo->m_listFilePath.Compare(filePath) == 0) // should be true
					int deb = 1;
				if (linfo)
				{
					linfo->m_listFilePath = filePath;
				}
			}
		}
	}
	finder.Close();

	this->SortChildren(hItem);
	return 1;
}

// only called by NTreeView::ShowGmailLabels; it is expensive if number of children is very large
// 05/14/2022 no longer called
HTREEITEM NTreeView::HasLabel(HTREEITEM hItem, CString& label)
{
	CString name;
	if (hItem != NULL)
	{
		HTREEITEM hChild = m_tree.GetChildItem(hItem);
		while (hChild)
		{
			name = m_tree.GetItemText(hChild);
			if (name.Compare(label) == 0)
				return hChild;
			hChild = m_tree.GetNextSiblingItem(hChild);
		}
	}
	return 0;
}

LabelInfo::LabelInfo(int nId, CString& mailFolderPath, CString& mailDataFolderPath)
{
	m_nodeType = MailFolder;

	m_nId = nId;
	m_mailFolderPath = mailFolderPath;
	m_mailDataFolderPath = mailDataFolderPath;
}

LabelInfo::LabelInfo(int nId, CString& mailFolderPath, CString& mailDataFolderPath, int nodeType)
{
	// Ignore nodeType arg ?? int nodeType is to create different signature than for MailFolder
	m_nodeType = MailSubFolder;

	m_nId = nId;
	m_mailFolderPath = mailFolderPath;
	m_mailDataFolderPath = mailDataFolderPath;
}

LabelInfo::LabelInfo(int nId, CString& mailFilePath)
{
	m_nodeType = MailFile;
	m_nId = nId;
	m_filePath = mailFilePath;
}
LabelInfo::LabelInfo(int nodeType, int nId, CString& mailFilePath, CString& mailFolderPath)
{
	if (nodeType != LabelInfo::MailFile)
		MboxMail::assert_unexpected();

	m_nodeType = MailFile;
	m_nId = nId;
	m_filePath = mailFilePath;
	m_mailFolderPath = mailFolderPath;
}

LabelInfo::LabelInfo(int nId, CString& mailFilePath, CString& label, CString& listFilePath)
{
	m_nodeType = MailLabel;
	m_nId = nId;
	m_label = label;
	m_listFilePath = listFilePath;
	m_filePath = mailFilePath;
}

LabelInfo::~LabelInfo()
{
	;
}


GlobalLabelInfo::GlobalLabelInfo()
{
	m_nId = 100000;  // base id
	m_labelInfoHT = new GlobalLableInfoMapType(311);
}

GlobalLabelInfo::~GlobalLabelInfo()
{
	Clear();
	delete m_labelInfoHT;
}

LabelInfo* GlobalLabelInfo::Find(int key)
{
	LabelInfo* linfo = m_labelInfoHT->find(&key);
	return linfo;
}

int GlobalLabelInfo::GetNextId()
{
	m_nId++;
	if (m_nId == 0x7fffffff)  // that might a problem; try rand() ??
	{
		// _ASSERTE(0);
		m_nId = 100000;
	}

	return m_nId;
}

int GlobalLabelInfo::Add(LabelInfo* linfo, int key)
{
	DWORD_PTR n = 0;

	// No check. User must call Find() first
	m_labelInfoHT->insert(&key, linfo);
	return m_nId;
}

void GlobalLabelInfo::Clear()
{
	GlobalLableInfoMapType::IHashMapIter iter = m_labelInfoHT->first();
	int totalCnt = 0;
	for (; !m_labelInfoHT->last(iter); )
	{
		LabelInfo* l = iter.element;
		m_labelInfoHT->remove(iter);
		delete l;

	}
	m_labelInfoHT->clear();
}

void GlobalLabelInfo::Print()
{
#ifdef _DEBUG
	GlobalLableInfoMapType::IHashMapIter iter = m_labelInfoHT->first();
	int totalCnt = 0;
	CString txt;
	CString nodeType = L"Unknown";

	TRACE(L"GLOBAL LABEL INFO\n");
	for (; !m_labelInfoHT->last(iter); )
	{
		LabelInfo* l = iter.element;
		if (l->m_nodeType == LabelInfo::MailLabel)
			nodeType = L"Label";
		else if (l->m_nodeType == LabelInfo::MailFile)
			nodeType = L"File";
		else if (l->m_nodeType == LabelInfo::MailFolder)
			nodeType = L"Folder";

		txt.Format(L"ID=%d Type=%s\n\tLabel=%s\n\tListFilePath=%s\n\tFilePath=%s\n\tMailFolderpath=%s\n\tMailDataFolderPath=%s\n\n",
			l->m_nId, nodeType, l->m_label, l->m_listFilePath, l->m_filePath, l->m_mailFolderPath, l->m_mailDataFolderPath
		);
		TRACE(L"%s", txt);
		m_labelInfoHT->next(iter);
		int deb = 1;
	}
#endif
}

int NTreeView::LoadLabels()
{
	CString name;
	DWORD nId = 0;
	CString sText;
	int paneId = 0;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame) {
		sText.Format(L"Loading labels ...");
		ResHelper::TranslateString(sText);
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);
	}
	HTREEITEM hRoot = m_tree.GetRootItem();
	if (hRoot)
		name = m_tree.GetItemText(hRoot);

	HTREEITEM hItem = 0;
	HTREEITEM hChild = m_tree.GetChildItem(hItem);
	while (hChild)
	{
		name = m_tree.GetItemText(hChild);
		nId = (DWORD)m_tree.GetItemData(hChild);

		LoadLabels(hChild);
		hChild = m_tree.GetNextSiblingItem(hChild);
	}

	sText.Format(L"Ready"); if (pFrame) pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
	return 1;
}

int NTreeView::LoadLabels(HTREEITEM hItem)
{
	CString sText;
	int paneId = 0;
	CString path;
	CString name;
	DWORD nId;
	LabelInfo* linfo;
	CString lastPath;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	//HTREEITEM hItem = m_tree.GetRootItem();
	if (hItem == 0)
	{
		MboxMail::assert_unexpected();
		return FALSE;
	}

	name = m_tree.GetItemText(hItem);
	nId = (DWORD)m_tree.GetItemData(hItem);

	linfo = m_labelInfoStore.Find(nId);
	if (linfo == 0)
	{
		MboxMail::assert_unexpected();
		return FALSE;
	}

	if ((linfo->m_nodeType != LabelInfo::MailFolder) && (linfo->m_nodeType != LabelInfo::MailSubFolder))
	{
		MboxMail::assert_unexpected();
		return FALSE;
	}

	CString folderlastPath = MboxMail::GetLastPath();
	folderlastPath.TrimRight(L"\\");
	if (folderlastPath.Compare(linfo->m_mailFolderPath))
		MboxMail::assert_unexpected();

	if (pFrame) {
		sText.Format(L"Loading labels ...");
		ResHelper::TranslateString(sText);
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);
	}

	HTREEITEM hChild = m_tree.GetChildItem(hItem);
	while (hChild)
	{
		name = m_tree.GetItemText(hChild);
		nId = (DWORD)m_tree.GetItemData(hChild);

		linfo = m_labelInfoStore.Find(nId);
		if (linfo == 0)
		{
			hChild = m_tree.GetNextSiblingItem(hChild);
			continue;
		}

		if ((linfo->m_nodeType == LabelInfo::MailFolder) || (linfo->m_nodeType == LabelInfo::MailSubFolder))
		{
			MboxMail::assert_unexpected();
			//int bRetval = NTreeView::LoadLabels(hChild);
			hChild = m_tree.GetNextSiblingItem(hChild);
			continue;
		}
		if (linfo->m_nodeType == LabelInfo::MailLabel)
		{
			MboxMail::assert_unexpected();
			hChild = m_tree.GetNextSiblingItem(hChild);
			continue;
		}

		CString labelsCachePath;
		CString rootPrintSubFolder = L"LabelCache";
		CString errorText;
		CString targetPrintSubFolder;

		//MboxMail::s_path = path;
		//MboxMail::SetMboxFilePath(path, TRUE);
		MboxMail::SetMboxFilePath(linfo->m_filePath);

		CString driveName;
		CString directory;
		CString fileNameBase;
		CString fileNameExtention;
		FileUtils::SplitFilePath(linfo->m_filePath, driveName, directory, fileNameBase, fileNameExtention);

		CString lastPath = driveName + directory;

		CString folderlastPath = MboxMail::GetLastPath();
		if (folderlastPath.Compare(lastPath))
			MboxMail::assert_unexpected();

		CString mboxIndexFilepath; // = pListView->m_path + ".mboxview";
		BOOL ret = MboxMail::GetMboxviewFilePath(linfo->m_filePath, mboxIndexFilepath);

		if (!FileUtils::PathFileExist(mboxIndexFilepath))
		{
			//MboxMail::assert_unexpected();
			DeleteLabelsForSingleMailFile(hChild);
			hChild = m_tree.GetNextSiblingItem(hChild);
			continue;
		}

		MboxMail::SetLastPath(lastPath);

		BOOL retval = MboxMail::GetCachePath(rootPrintSubFolder, targetPrintSubFolder, labelsCachePath, errorText, &linfo->m_filePath);
		if (retval == FALSE) {
			MboxMail::assert_unexpected();
			sText.Format(L"Ready"); if (pFrame) pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
			return -1;  // goto instead ?
		}

		if (!labelsCachePath.IsEmpty() && FileUtils::PathDirExists(labelsCachePath))
		{
			int ret = ShowGmailLabels(hChild, labelsCachePath, linfo->m_filePath);
			if (ret < 0) {
				MboxMail::assert_unexpected();
				sText.Format(L"Ready"); 
				ResHelper::TranslateString(sText);
				if (pFrame) pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
				return -1;  // goto instead ?
			}
			CString inbox(L"Inbox");
			CString sent(L"Sent");
			HTREEITEM hParent = m_tree.GetParentItem(hItem);
			this->SortChildren(hChild);
			MoveLabelItem(hChild, linfo->m_filePath, sent);
			MoveLabelItem(hChild, linfo->m_filePath, inbox);
		}
		else
		{
			int deb = 1;
		}

		hChild = m_tree.GetNextSiblingItem(hChild);
	}
	sText.Format(L"Ready");
	ResHelper::TranslateString(sText);
	if (pFrame) pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
	return 1;
}


// TODO: this likely doesn't work anymore due to support for Root Folder Select opion (i.e SubFolders). Verify
HTREEITEM NTreeView::HasMailFile(CString& mailFilePath)
{
	CString name;
	CString name2;
	DWORD nId;
	DWORD nId2;
	LabelInfo* linfo;
	LabelInfo* linfo2;

	HTREEITEM hItem = m_tree.GetRootItem();
	name = m_tree.GetItemText(hItem);

	hItem = 0;
	HTREEITEM hChild = m_tree.GetChildItem(hItem);
	while (hChild)
	{
		name = m_tree.GetItemText(hChild);
		nId = (DWORD)m_tree.GetItemData(hChild);

		linfo = m_labelInfoStore.Find(nId);
		HTREEITEM hChild2 = m_tree.GetChildItem(hChild);
		while (hChild2)
		{
			name2 = m_tree.GetItemText(hChild2);
			nId2 = (DWORD)m_tree.GetItemData(hChild2);

			linfo2 = m_labelInfoStore.Find(nId2);
			if (linfo2 == 0) {
				MboxMail::assert_unexpected();
				hChild2 = m_tree.GetNextSiblingItem(hChild2);
				continue;
			}

			if (linfo2->m_filePath.Compare(mailFilePath) == 0)
				return hChild2;
			hChild2 = m_tree.GetNextSiblingItem(hChild2);
		}
		hChild = m_tree.GetNextSiblingItem(hChild);
	}
	return 0;
}


// TODO: this likely doesn't work anymore due to support for Root Folder Select opion (i.e SubFolders). Verify
void NTreeView::MoveLabelItem(CString& mailFilePath, CString& label)
{
	HTREEITEM hRoot = HasMailFile(mailFilePath);
	HTREEITEM hitemToBeMoved = HasLabel(hRoot, label);
	if (hitemToBeMoved == 0)
		return;

	// TODO: don't move if hitemToBeMoved has children. It would require to recursively
	// move children one by one to new place and then delete hitemToBeMoved

	if (m_tree.ItemHasChildren(hitemToBeMoved))
		return;

	CString name = m_tree.GetItemText(hitemToBeMoved);
	DWORD nId = (DWORD)m_tree.GetItemData(hitemToBeMoved);

	HTREEITEM parent = m_tree.GetParentItem(hitemToBeMoved);
	m_tree.DeleteItem(hitemToBeMoved);
	HTREEITEM hItem = InsertTreeItem(name, 6, 7, parent, TVI_FIRST);
	m_tree.SetItemData(hItem, nId);  // TODO: no need to add to label store ?? alreadi in store ?? check/assert
}

void NTreeView::MoveLabelItem(HTREEITEM hRoot, CString& mailFilePath, CString& label)
{
	//HTREEITEM hRoot = HasMailFile(mailFilePath);
	HTREEITEM hitemToBeMoved = HasLabel(hRoot, label);
	if (hitemToBeMoved == 0)
		return;

	// TODO: don't move if hitemToBeMoved has children. It would require to recursively
	// move children one by one to new place and then delete hitemToBeMoved

	if (m_tree.ItemHasChildren(hitemToBeMoved))
		return;

	CString name = m_tree.GetItemText(hitemToBeMoved);
	DWORD nId = (DWORD)m_tree.GetItemData(hitemToBeMoved);

	HTREEITEM parent = m_tree.GetParentItem(hitemToBeMoved);
	m_tree.DeleteItem(hitemToBeMoved);
	HTREEITEM hItem = InsertTreeItem(name, 6, 7, parent, TVI_FIRST);
	m_tree.SetItemData(hItem, nId);  // TODO: no need to add to label store, already in store
}

// It will create label files and update the Tree
// Label files should be removed before calling this function
int NTreeView::CreateLabelsForSingleMailFile(HTREEITEM hItem)
{
	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	int paneId = 0;
	CString sText;
	CString mailFile = m_tree.GetItemText(hItem);

	CString fmt = L"Creating Gmail Labels for %s ...";
	ResHelper::TranslateString(fmt);
	sText.Format(fmt, mailFile);

	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	RecreateGmailLabels(hItem);

	sText.Format(L"Ready");
	ResHelper::TranslateString(sText);
	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, FALSE);

	return 1;
}


// Deletes all label files
int NTreeView::DeleteLabelsForSingleMailFile(HTREEITEM hItem)
{
	if (!hItem)
		return -1;

	CString labelsCachePath;
	CString rootPrintSubFolder = L"LabelCache";
	CString errorText;
	CString targetPrintSubFolder;

	CString name = m_tree.GetItemText(hItem);
	int nId = (DWORD)m_tree.GetItemData(hItem);

	LabelInfo* linfo = m_labelInfoStore.Find(nId);

	if (linfo)
	{
		BOOL retval = MboxMail::GetCachePath(rootPrintSubFolder, targetPrintSubFolder, labelsCachePath, errorText, &linfo->m_filePath);
		if (retval == FALSE) {
			MboxMail::assert_unexpected();
			return -1;
		}
	}
	else
	{
		BOOL retval = MboxMail::CreateCachePath(rootPrintSubFolder, targetPrintSubFolder, labelsCachePath, errorText);
		if (retval == FALSE) {
			MboxMail::assert_unexpected();
			return -1;
		}
	}

	CString lastDataPath = MboxMail::GetLastDataPath();
	if (FileUtils::PathDirExists(labelsCachePath))
	{
		// TODO: remove files, keep tree item
		if (hItem) {
			DeleteItemChildren(hItem);
		}
		BOOL retval = FileUtils::RemoveDir(labelsCachePath, true, true);
		int deb = 1;
	}
	return 1;
}

int NTreeView::RefreshLabelsForSingleMailFile(HTREEITEM hItem)
{
	CString lastPath = MboxMail::GetLastPath();
	CString lastDataPath = MboxMail::GetLastDataPath();

	int retD = DeleteLabelsForSingleMailFile(hItem);
	int retC = CreateLabelsForSingleMailFile(hItem);

	return retC;
}

int NTreeView::UpdateLabelMailListFile(HTREEITEM hItem)
{
	CString errorText;
	CString mboxFileNamePath = MboxMail::s_path;
	CString labelListFileNamePath = MboxMail::s_path_label;
	CString lastPath = MboxMail::GetLastPath();
	CString lastDataPath = MboxMail::GetLastDataPath();

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	NListView* pListView = 0;
	if (pFrame)
		pListView = pFrame->GetListView();

	if (pListView == 0)
	{
		MboxMail::assert_unexpected();
		return -1;
	}
	if (!labelListFileNamePath.IsEmpty())
	{
		_int64 mailFileSize = MboxMail::s_fSize;
		int wret = pListView->WriteMboxListFile_v2(&MboxMail::s_mails, labelListFileNamePath, mailFileSize, errorText);
		if (wret < 0)
		{
			MboxMail::assert_unexpected();
			return -1;
		}
	}
	else
	{
		MboxMail::assert_unexpected();
		return -1;
	}
	return 1;
}

BOOL NTreeView::DeleteFolderIfEmpty(CString& path)
{
	NTreeView* pTreeView = this;

	return FALSE;

	// HasFolder checks whether folder is a child of CtrlTree so sub-folders should not be affected
	HTREEITEM hFolder = pTreeView->HasFolder(path);
	if (hFolder)
	{
		int cnt = pTreeView->GetChildrenCount(hFolder, FALSE);
		if (cnt == 0)
		{
			pTreeView->DeleteFolder(hFolder);
#if 0
			CString txt = L"No valid mail archive files found in\n\n\"";
			txt.Append(path);
			txt.Append(L"\"\n\nfolder. Ignoring.\n");
#endif

			CString fmt = L"No valid mail archive files found in\n\n\"%s\"\n\nfolder. Ignoring.";
			ResHelper::TranslateString(fmt);
			CString txt;
			txt.Format(fmt, path);


			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
			int deb = 1;
			return TRUE;
		}
	}
	return FALSE;
}

void NTreeView::OnTvnGetInfoTip(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMTVGETINFOTIP pGetInfoTip = (LPNMTVGETINFOTIP)pNMHDR;

	HTREEITEM hItem = pGetInfoTip->hItem;
	if (hItem)
	{
		DWORD nId = (DWORD)m_tree.GetItemData(hItem);
		LabelInfo* linfo = m_labelInfoStore.Find(nId);
		if (linfo && ((linfo->m_nodeType == LabelInfo::MailFolder) || (linfo->m_nodeType == LabelInfo::MailSubFolder)))
		{
			CString csItemTxt = m_tree.GetItemText(hItem);
			CString folderPath = linfo->m_mailFolderPath;
			int len = folderPath.GetLength();
			if (len > (pGetInfoTip->cchTextMax - 2))
				len = pGetInfoTip->cchTextMax - 2;

			_tcsncpy(pGetInfoTip->pszText, folderPath, len);
			pGetInfoTip->pszText[len] = 0;
		}
	}
	*pResult = 0;
}

void NTreeView::OpenFolderAndSubfolders(HTREEITEM hParent, CString& path)
{
	CString rootFolder = path;
	CList<CString, CString&> folderList;
	CString errorText;
	int maxDepth = 10;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	BOOL bRetVal = FileUtils::GetFolderList(rootFolder, folderList, errorText, maxDepth);


	// TODO: Hash Table ??
	CArray<CString> cacheFolderList;
	NTreeView::SetupCacheFolderList(cacheFolderList);

	CString folderPath;
	CString cacheFolderName;
	CString folderName;
	BOOL ignoreFolder;
	while (folderList.GetCount() > 0)
	{
		folderPath = folderList.RemoveHead();
		TRACE(L"Folder Path: \"%s\"\n", folderPath);

		FileUtils::CPathStripPath(folderPath, folderName);
		ignoreFolder = FALSE;
		for (int i = 0; i < cacheFolderList.GetCount(); i++)
		{
			CString& cacheFolderName = cacheFolderList[i];
			if (folderName.Compare(cacheFolderName) == 0)
			{
				ignoreFolder = TRUE;
				break;
			}
		}
		if (ignoreFolder)
			continue;

		BOOL unconditionalFolderInsert = FALSE;
		BOOL selectFolder = FALSE;
		BOOL addToRecentFileList = FALSE;
		BOOL expand = TRUE;
		OpenFolder(hParent, folderPath, selectFolder, unconditionalFolderInsert, addToRecentFileList, expand);
	}
}

HTREEITEM NTreeView::OpenFolder(HTREEITEM hFolder, CString& path, BOOL selectFolder, BOOL unconditionalFolderInsert, BOOL addToRecentFileList, BOOL expand)
{
	CString fileName;
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);


	path.TrimRight(L"\\");
	path.Append(L"\\");

	FileUtils::SplitFilePath(path, driveName, directory, fileNameBase, fileNameExtention);
	if (directory.GetLength() <= 1)
	{
		CString txt = L"The mbox files must be installed under a named folder\n."
			L"Please create folder, move the mbox files to that folder and try again.";

		ResHelper::TranslateString(txt);

		int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
		return 0;
	}

	path.TrimRight(L"\\");
	if (addToRecentFileList)
		AfxGetApp()->AddToRecentFileList(path);

	path.Append(L"\\");
	MboxMail::SetLastPath(path);

	int paneId = 0;
	CString sText;
	CString fname;
	FileUtils::CPathStripPath(path, fname);
	CString fmt = L"Opening new mail folder %s ...";
	ResHelper::TranslateString(fmt);
	sText.Format(fmt, fname);

	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	//BOOL selectFolder = TRUE;
	//BOOL unconditionalFolderInsert = FALSE;
	HTREEITEM hItem = FillCtrl(hFolder, selectFolder, unconditionalFolderInsert, expand);
	if (hItem == 0)
	{
		MboxMail::assert_unexpected();

		sText.Format(L"Ready");
		ResHelper::TranslateString(sText);
		if (pFrame)
			pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
		return hItem;
	}

	if (!DeleteFolderIfEmpty(path))
	{
		CString filePath = path;
		filePath.TrimRight(L"\\");
		AfxGetApp()->AddToRecentFileList(filePath);
	}

	this->SortChildren(hItem);
	CString itemName = m_tree.GetItemText(hItem);

	if (expand)
		ExpandItem(hItem, TVE_EXPAND);

	sText.Format(L"Ready");
	ResHelper::TranslateString(sText);
	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, FALSE);

	TRACE(L"");
	TRACE(L"OpenFolder: done no labels yet\n");
	m_globalFolderInfoDB.Print();

	return hItem;
}

void NTreeView::DoOpenRootFolderAndSubfolders(HTREEITEM hParent, CString& path, BOOL selectFolder)
{
	BOOL expand = TRUE;
	HTREEITEM hFolder = SelectFolderIfAlreadyOpen(&path, FolderInfo::MailRootFolderNode, expand);

	if (hFolder == 0)
	{
		OpenRootFolderAndSubfolders(hParent, path, selectFolder);
	}
}

void NTreeView::OpenRootFolderAndSubfolders(HTREEITEM hParent, CString& path, BOOL selectFolder)
{
	CString rootFolder = path;
	CList<CString, CString&> folderList;
	CString errorText;
	int maxDepth = 10;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	MBoxFolderTree tree(path);

	MBoxFolderNode* rnode = 0;
	BOOL bRet = tree.PopulateFolderTree(path, tree, rnode, errorText, maxDepth);
	if (!errorText.IsEmpty())
	{
		int answer = MessageBox(errorText, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;
	}
	int treeFolderCount = tree.Count();

	TRACE(L"\n\nBefore Prunning\n\n");
	tree.Print(path);

	PruneFolderTree(tree);

	treeFolderCount = tree.Count();

	TRACE(L"\n\nAfter Prunning\n\n");
	tree.Print(path);

	if (treeFolderCount == 0)
	{
		//errorText.Format(L"No folders with mbox files found under the root folder and all sub-folders:\n\n%s", path);

		CString fmt = L"No folders with mbox files found under the root folder and all sub-folders:\n\n%s";
		ResHelper::TranslateString(fmt);
		errorText.Format(fmt, path);

		int answer = MessageBox(errorText, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;
	}

	if (ValidateFolderTree(tree, m_globalFolderInfoDB))
	{
		MboxMail::SetLastPath(path);
		SetFolderAsRoorFolder(path);

		OpenTreeFolders(hParent, tree);

		MboxMail::SetLastPath(path);
		SetFolderAsRoorFolder(path);
		if (selectFolder)
			HTREEITEM hLastFolder = SelectTreeItem(&path);
		TRACE(L"==========>> OpenRootFolderAndSubfolders: ValidateFolderTree TRUE\n");
		m_globalFolderInfoDB.Print();
	}
	else
	{
		// remove root folder and subfolders from m_globalFolderInfoDB
		//m_globalFolderInfoDB.RemoveRootFolder(path);
		TRACE(L"============>> OpenRootFolderAndSubfolders: ValidateFolderTree FALSE\n");
		m_globalFolderInfoDB.Print();

		CString txt;
		txt.Format(L"Can't open root folder \n\n\"%s\"\n"
			L"\n\nAt least one of the root sub-folders is already open\n\n\"%s\"\n"
			L"\nDo you want to remove all conflicting folders from the Tree View and continue to open the root folder?\n",
			path, m_globalFolderInfoDB.m_problemFolderPath);

		HWND h = GetSafeHwnd();
		int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONHAND | MB_YESNO);
		if (answer == IDYES)
		{
			m_globalFolderInfoDB.RemoveRootFolder(path);

			// open root folder and remove all conflicting folders and subfolders
			m_globalFolderInfoDB.m_deleteConflictingFolders = TRUE;
			if (ValidateFolderTree(tree, m_globalFolderInfoDB))
			{
				MboxMail::SetLastPath(path);
				SetFolderAsRoorFolder(path);

				OpenTreeFolders(hParent, tree);

				MboxMail::SetLastPath(path);
				SetFolderAsRoorFolder(path);
				if (selectFolder)
					HTREEITEM hLastFolder = SelectTreeItem(&path);
				TRACE(L"=============>>OpenRootFolderAndSubfolders: Force ValidateFolderTree TRUE\n");
				m_globalFolderInfoDB.Print();
			}
			else
			{
				TRACE(L"============>> OpenRootFolderAndSubfolders: Force ValidateFolderTree FALSE\n");
				m_globalFolderInfoDB.Print();
			}
			m_globalFolderInfoDB.m_deleteConflictingFolders = FALSE;
		}

		int deb = 1;
	}
}

BOOL NTreeView::SetFolderAsRoorFolder(CString& folderPath)
{
	CString datapath = MboxMail::GetLastDataPath();
	CFile fp;
	CFileException ExError;
	CString rootFolderFile = datapath + L".urootfolder";
	if (!fp.Open(rootFolderFile, CFile::modeWrite | CFile::modeCreate, &ExError))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not create \"" + rootFolderFile;
		txt += L"\" file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);
		MboxMail::assert_unexpected();
		return FALSE;
	}
	else
	{
		TRACE(L"Created root file=%s\n", rootFolderFile);
		fp.Close();
		return TRUE;
	}
}

BOOL NTreeView::RemoveFolderAsRootFolder(CString& folderPath)
{
	CString fPath = folderPath;
	fPath.TrimRight(L"\\");
	CString lastPath = MboxMail::GetLastPath();
	lastPath.TrimRight(L"\\");

	CString datapath = MboxMail::GetLastDataPath();
	CString rootFolderFile = datapath + L".urootfolder";
	BOOL retval = FileUtils::DelFile(rootFolderFile, FALSE);
	return retval;
}

BOOL NTreeView::IsFolderARootFolder(CString& folderPath)
{
	CString datapath = MboxMail::GetLastDataPath();
	CString rootFolderFile = datapath + L".urootfolder";
	BOOL retval = FileUtils::PathFileExist(rootFolderFile);
	return retval;
}

BOOL NTreeView::IsFolderOpenAsRootFolder(CString& folderPath)
{
	CString sv_lastPath = MboxMail::GetLastPath();

	MboxMail::SetLastPath(folderPath);
	CString datapath = MboxMail::GetLastDataPath();

	CString rootFolderFile = datapath + L".urootfolder";
	BOOL retval = FileUtils::PathFileExist(rootFolderFile);

	MboxMail::SetLastPath(sv_lastPath);
	return retval;
}

BOOL NTreeView::CanFolderBeOpen(CString& path)
{
	// TODO:: else should catch
	// Do wee this, below using FolderInfo should work
	if (IsFolderARootFolder(path))
		return FALSE;

	hashsum_t hashsum = GlobalFolderInfoDB::GetHashsum(&path);
	FolderInfo* finfo = m_globalFolderInfoDB.Find(&path, hashsum);
	if (finfo == 0)
	{
		return TRUE;
	}
	else
	{
		// TODO: we should not be here, see SelectFolderIfAlreadyOpen in DoOpen
		if (finfo->m_folderType == FolderInfo::MailFolderNode)
			return TRUE;
		else if (finfo->m_folderType == FolderInfo::MailRootFolderNode)
			return FALSE;
		else if (finfo->m_folderType == FolderInfo::MailSubFolderNode)
			return FALSE;
		else
		{
			MboxMail::assert_unexpected();
			return TRUE;
		}
	}
}

BOOL NTreeView::CanRootFolderBeOpen(CString& path)
{
	hashsum_t hashsum = GlobalFolderInfoDB::GetHashsum(&path);
	FolderInfo* finfo = m_globalFolderInfoDB.Find(&path, hashsum);
	if (finfo == 0)
	{
		finfo = new FolderInfo(path, FolderInfo::MailRootFolderNode);
		m_globalFolderInfoDB.Add(hashsum, finfo);
		return TRUE;
	}
	else
	{
		if (finfo->m_folderType == FolderInfo::MailFolderNode)
			return FALSE;
		else if (finfo->m_folderType == FolderInfo::MailRootFolderNode)
			return TRUE;
		else if (finfo->m_folderType == FolderInfo::MailSubFolderNode)
			return FALSE;
		else
		{
			MboxMail::assert_unexpected();
			return TRUE;
		}
	}
}

// TODO: It is basically identical to DoOpenFoldr, consolidate
HTREEITEM NTreeView::OpenRootSubFolder(HTREEITEM hFolder, CString& path, BOOL addToRecentFileList, BOOL expand)
{
#if 1
	CString fileName;
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);


	path.TrimRight(L"\\");
	path.Append(L"\\");

	FileUtils::SplitFilePath(path, driveName, directory, fileNameBase, fileNameExtention);
	if (directory.GetLength() <= 1)
	{
		CString txt = L"The mbox files must be installed under a named folder.\n"
			L"Please create folder, move the mbox files to that folder and try again.";
		int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
		return 0;
	}

	path.TrimRight(L"\\");
	if (addToRecentFileList)
		AfxGetApp()->AddToRecentFileList(path);

	path.Append(L"\\");
	MboxMail::SetLastPath(path);

	int paneId = 0;
	CString sText;
	CString fname;
	FileUtils::CPathStripPath(path, fname);
	CString fmt = L"Opening new mail folder %s ...";
	ResHelper::TranslateString(fmt);
	sText.Format(fmt, fname);
	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	BOOL selectFolder = FALSE;
	BOOL unconditionalFolderInsert = TRUE;
	HTREEITEM hItem = FillCtrl(hFolder, selectFolder, unconditionalFolderInsert, expand);
	if (hItem == 0)
	{
		MboxMail::assert_unexpected();
		sText.Format(L"Ready");
		ResHelper::TranslateString(sText);
		if (pFrame)
			pFrame->SetStatusBarPaneText(paneId, sText, FALSE);
		return hItem;
	}
#if 0
	if (!DeleteFolderIfEmpty(path))
	{
		CString filePath = path;
		filePath.TrimRight(L"\\");
		AfxGetApp()->AddToRecentFileList(filePath);
	}
#endif

	this->SortChildren(hItem);
	CString itemName = m_tree.GetItemText(hItem);

	//ExpandItem(hItem, TVE_EXPAND);

	sText.Format(L"Ready");
	ResHelper::TranslateString(sText);
	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, FALSE);

	TRACE(L"");
	TRACE(L"OpenRootSubFolder: done no labels yet\n");
	m_globalFolderInfoDB.Print();

	return hItem;
#else
	CString fileName;
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	HTREEITEM hItem = 0;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	path.TrimRight(L"\\");
	path.Append(L"\\");

	FileUtils::SplitFilePath(path, driveName, directory, fileNameBase, fileNameExtention);
	if (directory.GetLength() <= 1)
	{
		CString txt = L"The mbox files must be installed under a named folder\n."
			"Please create folder, move the mbox files to that folder and try again.";
		int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
		return hItem;
	}

	path.TrimRight(L"\\");
	if (addToRecentFileList)
		AfxGetApp()->AddToRecentFileList(path);

	path.Append(L"\\");
	MboxMail::SetLastPath(path);

	int paneId = 0;
	CString sText;
	CString fname;
	FileUtils::CPathStripPath(path, fname);
	sText.Format(L"Opening new mail folder %s ...", fname);
	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	BOOL selectFolder = FALSE;
	BOOL unconditionalFolderInsert = TRUE;
	hItem = FillCtrl(hParent, selectFolder, unconditionalFolderInsert, expand);

	SortChildren(hItem);

	sText.Format(L"Ready");
	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, FALSE);

	return hItem;
#endif
}

// MBoxFolderTree

void NTreeView::PruneFolderTree(MBoxFolderTree& tree)
{
	MBoxFolderNode* rootNode = tree.GetRootNode();
	if (rootNode == 0)
		return;

	//TRACE("Pruned Folder List:\n");

	PruneNonMBoxFolders(rootNode);

	if (rootNode && !rootNode->m_valid)
	{
		if (!rootNode->m_nodeList.empty())
		{
			int deb = 1;  // _ASSERTE ??
		}
		tree.EraseRoot();
	}
}

void NTreeView::PruneNonMBoxFolders(MBoxFolderNode* node)
{
	if (node == 0)
		return;

	list<MBoxFolderNode>::iterator it;
	list<MBoxFolderNode>::iterator next_it;
	CString folderName;

	if (node->m_nodeList.empty())
	{
		//TRACE("%s (Leave Node)\n", node->m_folderName);
		if (NTreeView::FolderHasAtLeastOneMboxFile(node->m_folderName) == FALSE)
			node->m_valid = FALSE;
		return;
	}
	else
	{
		//TRACE("%s\n", node->m_folderName);
		for (it = node->m_nodeList.begin(); it != node->m_nodeList.end();)
		{
			next_it = it;
			next_it++;
			const wchar_t* p = it->m_folderName;
			folderName = it->m_folderName;

			PruneNonMBoxFolders(&(*it));
			it = next_it;
			int deb = 1;
		}

		for (it = node->m_nodeList.begin(); it != node->m_nodeList.end();)
		{
			next_it = it;
			next_it++;
			const wchar_t* p = it->m_folderName;
			folderName = it->m_folderName;
			if (!it->m_valid)
				node->m_nodeList.erase(it);

			it = next_it;
			int deb = 1;
		}

		if (node->m_nodeList.empty())
		{
			//TRACE("Leave Node: %s\n", node->m_folderName);
			if (NTreeView::FolderHasAtLeastOneMboxFile(node->m_folderName) == FALSE)
				node->m_valid = FALSE;
		}
		else
		{
#if 0
			// FIXMEFIXME No comments unfortunatelly. Need to waste time to understand before I delete this
			CFile fp;
			CString path = node->m_folderName;
			path.TrimRight("\\");
			path.Append("\\");

			MboxMail::SetLastPath(path);

			CString datapath = MboxMail::GetLastDataPath();
			FileUtils::CreateDir(datapath);
			datapath.TrimRight("\\");

			CFileException ExError;
			CString viewFile = datapath + "\\.folderview";
			if (!fp.Open(viewFile, CFile::modeWrite | CFile::modeCreate, &ExError))
			{
				CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

				CString txt = L"Could not create \"" + viewFile;
				txt += L"\" file.\n";
				txt += exErrorStr;
				MboxMail::assert_unexpected();
				TRACE(L"%s\n", txt);
			}
			else
			{
				CString line;
				for (it = node->m_nodeList.begin(); it != node->m_nodeList.end(); it++)
				{
					folderName = it->m_folderName;
					line.Format("%s\n", folderName);
					fp.Write(line, line.GetLength());

					int deb = 1;
				}
				fp.Close();
			}
#endif
			int dreb = 1;
		}
	}
}

BOOL NTreeView::FolderHasAtLeastOneMboxFile(CString& mboxFilePath)
{
	// build a string with wildcards
	CString strWildcard(mboxFilePath);
	strWildcard += L"\\*.*";

	// start working on files
	CFileFind finder;
	BOOL bWorking = finder.FindFile(strWildcard);

	CString fPath;
	CString fName;

	while (bWorking)
	{
		bWorking = finder.FindNextFile();

		// skip . and .. files; otherwise, we'd
		// recur infinitely!

		if (finder.IsDots())
			continue;

		fPath = finder.GetFilePath();
		fName = finder.GetFileName();

		if (NTreeView::ImboxviewFile(fPath))
		{
			return TRUE;
		}
	}
	return FALSE;
}

HTREEITEM NTreeView::OpenTreeNode(HTREEITEM hParent, MBoxFolderNode* node)
{
	if (node == 0)
		return 0;

	CString parentFolderName;
	if ((hParent != TVI_ROOT) && (hParent != 0))
		parentFolderName = m_tree.GetItemText(hParent);

	CString folderName;
	CString folderPath = node->m_folderName;
	FileUtils::CPathStripPath(folderPath, folderName);

	BOOL addToRecentFileList = FALSE;
	BOOL expand = FALSE;
	HTREEITEM hItem = OpenRootSubFolder(hParent, folderPath, addToRecentFileList, expand);

	CString itemFolderName;
	if (hItem == 0)
		return hItem;

	itemFolderName = m_tree.GetItemText(hItem);

	LoadLabels(hItem);

	if (node->m_nodeList.empty())
	{
		TRACE(L"%s (Leave Node)\n", node->m_folderName);
		return hItem;
	}
	else
	{
		TRACE(L"%s\n", node->m_folderName);
		list<MBoxFolderNode>::iterator it;
		for (it = node->m_nodeList.begin(); it != node->m_nodeList.end(); it++)
		{
			OpenTreeNode(hItem, &(*it));
			int deb = 1;
		}
	}
	//if (hParent == TVI_ROOT)
		//ExpandItem(hItem, TVE_EXPAND);
	return hItem;
}

void NTreeView::OpenTreeFolders(HTREEITEM hParent, MBoxFolderTree& tree)
{
	TRACE(L"MBoxFolderTree: OpenTreeFolders:\n");
	MBoxFolderNode* rootNode = tree.GetRootNode();
	if (rootNode == 0)
		return;
	HTREEITEM hItem = OpenTreeNode(hParent, rootNode);
	//ExpandItem(hParent, TVE_EXPAND);

	BOOL recursive = TRUE;
	int depth = 0;
	TRACE(L"Tree Control:\n");
	PrintTreeCtrl(hItem, recursive, depth);
}

BOOL NTreeView::ValidateFolderTreeNode(MBoxFolderNode* node, GlobalFolderInfoDB& globalFolderInfoDB)
{
	if (node == 0)
		return 0;

	CString folderName;
	CString folderPath = node->m_folderName;
	FileUtils::CPathStripPath(folderPath, folderName);
	folderPath.Append(L"\\");

	hashsum_t hashsum = GlobalFolderInfoDB::GetHashsum(&folderPath);
	FolderInfo* finfo = globalFolderInfoDB.Find(&folderPath, hashsum);

	if (finfo && globalFolderInfoDB.m_deleteConflictingFolders)
	{
		if (HTREEITEM hFolder = IsFolderOpen(finfo->m_RootParent->m_folderPath))
		{
			DeleteFolder(hFolder, FALSE);
		}
		finfo = globalFolderInfoDB.Find(&folderPath, hashsum);
	}
	if (finfo)
	{
		if (globalFolderInfoDB.m_deleteConflictingFolders)
			MboxMail::assert_unexpected();

		globalFolderInfoDB.m_problemFolderPath = finfo->m_folderPath;
		globalFolderInfoDB.m_problemFolderType = finfo->m_folderType;
		globalFolderInfoDB.m_problemFolderRootPath = finfo->m_RootParent->m_folderPath;
		if (globalFolderInfoDB.m_deleteConflictingFolders && folderPath.Compare(finfo->m_RootParent->m_folderPath))
			MboxMail::assert_unexpected();
		return FALSE;
	}
	else
	{
		int folderType = FolderInfo::MailRootFolderNode;
		if (node->m_parent != 0)
		{
			folderType = FolderInfo::MailSubFolderNode;
			RemoveFolderAsRootFolder(folderPath);
		}

		finfo = new FolderInfo(folderPath, folderType);
		m_globalFolderInfoDB.Add(hashsum, finfo);
	}

	if (node->m_nodeList.empty())
	{
		TRACE(L"%s (Leave Node)\n", node->m_folderName);
		return TRUE;
	}
	else
	{
		TRACE(L"%s\n", node->m_folderName);
		list<MBoxFolderNode>::iterator it;
		for (it = node->m_nodeList.begin(); it != node->m_nodeList.end(); it++)
		{
			MBoxFolderNode* subNode = &(*it);
			BOOL retval = ValidateFolderTreeNode(subNode, globalFolderInfoDB);
			if (retval == FALSE)
				return FALSE;
			int deb = 1;
		}
	}
	return TRUE;
}

BOOL NTreeView::ValidateFolderTree(MBoxFolderTree& tree, GlobalFolderInfoDB& globalFolderInfoDB)
{
	TRACE(L"MBoxFolderTree: ValidateTreeFolders:\n");
	MBoxFolderNode* rootNode = tree.GetRootNode();
	if (rootNode == 0)
		return TRUE;

	CString folderPath = rootNode->m_folderName;
	folderPath.Append(L"\\");

	FolderInfo* rfinfo = globalFolderInfoDB.FindRoot(&folderPath);
	if (rfinfo == 0)
	{
		rfinfo = new FolderInfo(folderPath, FolderInfo::MailRootFolderNode);
		globalFolderInfoDB.AddRoot(&rfinfo->m_folderPath, rfinfo);
	}

	BOOL retval = ValidateFolderTreeNode(rootNode, globalFolderInfoDB);
	return retval;
}

void NTreeView::PrintTreeCtrl(HTREEITEM hItem, BOOL recursive, int depth)
{
#ifdef _DEBUG
	CString itemName;
	DWORD nId = 0;

	if (hItem == TVI_ROOT)
		hItem = 0;

	CString depthStr;
	for (int i = 0; i < depth; i++)
	{
		depthStr.Append(L"----");
	}

	if (hItem != 0)
	{
		itemName = m_tree.GetItemText(hItem);
		nId = (DWORD)m_tree.GetItemData(hItem);
		TRACE(L"%s%s  %d\n", depthStr, itemName, nId);
	}

	if (m_tree.ItemHasChildren(hItem))
	{
		HTREEITEM child = m_tree.GetChildItem(hItem);
		while (child != NULL)
		{
			if (recursive)
			{
				PrintTreeCtrl(child, recursive, depth + 1);
			}
			child = m_tree.GetNextItem(child, TVGN_NEXT);
		}
	}
#endif
}

BOOL NTreeView::AssertTreeCtrl(HTREEITEM hItem)
{
#ifdef _DEBUG
	CString lastPath = MboxMail::GetLastPath();
	BOOL retval = NTreeView::AssertTreeCtrl_Internal(hItem);
	MboxMail::SetLastPath(lastPath);
	return retval;
#else
	return TRUE;
#endif
}

BOOL NTreeView::AssertTreeCtrl_Internal(HTREEITEM hItem)
{
	CString itemName;
	CString dataPath;
	CString rootFile;
	DWORD nId = 0;

	if (hItem == TVI_ROOT)
		hItem = 0;

	if (hItem != 0)
	{
		itemName = m_tree.GetItemText(hItem);
		nId = (DWORD)m_tree.GetItemData(hItem);
		//TRACE("%s  %d\n", itemName, nId);

		LabelInfo* linfo = m_labelInfoStore.Find(nId);

		if ((linfo->m_nodeType == LabelInfo::MailFolder) || (linfo->m_nodeType == LabelInfo::MailSubFolder))
		{
			CString folderPath = linfo->m_mailFolderPath;
			folderPath.TrimRight(L"\\");
			folderPath.Append(L"\\");

			TRACE(L"AssertTreeCtrl: FolderPath=%s\n", folderPath);
			hashsum_t hashsum = GlobalFolderInfoDB::GetHashsum(&folderPath);
			FolderInfo* finfo = m_globalFolderInfoDB.Find(&folderPath);
			if (finfo == 0)
			{
				MboxMail::assert_unexpected();
				return FALSE;
			}
			MboxMail::SetLastPath(folderPath);
			dataPath = MboxMail::GetLastDataPath();

			if (IsFolderARootFolder(folderPath))
			{
				if (finfo->m_folderType == FolderInfo::MailFolderNode)
				{
					MboxMail::assert_unexpected();
					return FALSE;
				}
			}
			else if (finfo->m_folderType == FolderInfo::MailRootFolderNode)
			{
				MboxMail::assert_unexpected();
				return FALSE;
			}
			}
		}

	if (m_tree.ItemHasChildren(hItem))
	{
		HTREEITEM child = m_tree.GetChildItem(hItem);
		while (child != NULL)
		{
			AssertTreeCtrl_Internal(child);
			child = m_tree.GetNextItem(child, TVGN_NEXT);
		}
	}
	return TRUE;
	}

BOOL NTreeView::AssertNTreeView()
{
	BOOL retval = TRUE;
#ifdef _DEBUG
	CString lastPath = MboxMail::GetLastPath();
	retval = NTreeView::AssertNTreeView_Internal();
	MboxMail::SetLastPath(lastPath);
#endif
	return retval;
}

BOOL NTreeView::AssertNTreeView_Internal()
{
	BOOL retval = TRUE;
	if (m_treeViewAssertFlags & ASSERT_TREE_CTRL)
		retval = AssertTreeCtrl();

	// Folder Array in Registry and m_folderArray should be in sync
	CString folderPath;
	CRegArray folderArray;

	if (m_treeViewAssertFlags & ASSERT_FOLDERS_IN_REG_AND_MEM)
	{
		folderArray.LoadFromRegistry();
		int ii = 0;
		for (ii = 0; ii < folderArray.m_array.GetCount(); ii++)
		{
			folderPath = folderArray.m_array.GetAt(ii);
			if (m_folderArray.Find(folderPath) < 0)
			{
				MboxMail::assert_unexpected();
				return FALSE;
			}
		}
	}

	// Folder Array m_folderArray and CTreeCtrl hould be in sync
	if (m_treeViewAssertFlags & ASSERT_FOLDERS_IN_MEM_AND_TREE)
	{
		int ii = 0;
		for (ii = 0; ii < m_folderArray.m_array.GetCount(); ii++)
		{
			folderPath = m_folderArray.m_array.GetAt(ii);
			if (!folderPath.IsEmpty())
			{
				if (!IsFolderOpen(folderPath))
				{
					MboxMail::assert_unexpected();
					return FALSE;
				}
			}
		}
	}

	return TRUE;
}

BOOL NTreeView::DeleteOldDataFolder(CString& currentRootDataFolderPath, CString& newRootDataFolderPath, HTREEITEM hItem)
{
	_ASSERTE(!currentRootDataFolderPath.IsEmpty() && !newRootDataFolderPath.IsEmpty());

	BOOL retval = TRUE;

	// TODO: Keep it simple
	CString lastPath = MboxMail::GetLastPath();

	CString fromDataFolderPath = currentRootDataFolderPath + L"UMBoxViewer\\";
	CString toDataFolderPath = newRootDataFolderPath + L"UMBoxViewer\\";

	CStringArray excludeFilter;
	excludeFilter.Add(L"PrintCache");
	excludeFilter.Add(L"MergeCache");

	CStringW errorText = FileUtils::CopyDirectory(fromDataFolderPath, toDataFolderPath, &excludeFilter, FALSE, FALSE);
	//FileUtils::RemoveDir(dataFolderPath, TRUE, TRUE);

	MboxMail::SetLastPath(lastPath);

	return retval;
}


void NTreeView::DoOpen(CString& path, BOOL expand)
{
	HTREEITEM hFolder = SelectFolderIfAlreadyOpen(&path, FolderInfo::MailFolderNode, expand);

	if (hFolder == 0)
	{
		BOOL unconditionalFolderInsert = TRUE;
		BOOL selectFolder = TRUE;
		DoOpenFolder(path, selectFolder, unconditionalFolderInsert, expand);
		this->SortChildren(TVI_ROOT);
	}
}

void NTreeView::DoOpenFilePreview(CString& path)
{
	MboxMail::s_folderContext.m_folderPath.Empty();
	MboxMail::s_folderContext.m_dataFolderPath.Empty();
	MboxMail::s_folderContext.m_rootDataFolderPathSubFolderConfig = L"PView\\";
	if (MboxMail::s_folderContext.m_rootDataFolderPathConfig.IsEmpty())
	{
		_ASSERTE(1);
		MboxMail::s_folderContext.m_rootDataFolderPathConfig = FileUtils::GetMboxviewLocalAppPath();
		MboxMail::s_folderContext.m_rootDataFolderPath = MboxMail::s_folderContext.m_rootDataFolderPathConfig;
	}
	else
	{
		MboxMail::s_folderContext.m_rootDataFolderPathConfig = MboxMail::s_folderContext.m_rootDataFolderPathConfig;
		MboxMail::s_folderContext.m_rootDataFolderPath = MboxMail::s_folderContext.m_rootDataFolderPathConfig;
	}

	CString mboxFilePath;
	CString mboxFileName;
	FileUtils::GetFolderPathAndFileName(path, mboxFilePath, mboxFileName);

	mboxFilePath.TrimRight(L"\\");
	mboxFilePath.Append(L"\\");

	CString mboxFileNameOrPath = path;
	mboxFileNameOrPath.TrimRight(L"\\");


	BOOL unconditionalFolderInsert = FALSE;
	BOOL expand = TRUE;
	BOOL selectFolder = FALSE;
	DoOpenFolder(mboxFilePath, selectFolder, unconditionalFolderInsert, expand);
	this->SortChildren(TVI_ROOT);

	// For now don't delete created folder and files in OnClose()
	CMainFrame::m_commandLineParms.m_bEmlPreviewFolderExisted = TRUE;
	//if (pTreeView->m_folderArray.Find(mboxFilePath) >= 0)
		//m_commandLineParms.m_bEmlPreviewFolderExisted = TRUE;

	if (m_folderArray.Find(mboxFilePath) >= 0)
		HTREEITEM hItem = SelectFolderItem(mboxFileNameOrPath);
}


//  Called when opening non root folder
void NTreeView::DoOpenFolder(CString& folderPath, BOOL selectFolder, BOOL unconditionalFolderInsert, BOOL expand)
{
	HTREEITEM hParent;
	HTREEITEM hItem;

	CString lastPath = MboxMail::GetLastPath();
	CString path = folderPath;
	path.TrimRight(L"\\");
	path.Append(L"\\");
	MboxMail::SetLastPath(path);


	HTREEITEM hOpenFolder = IsFolderOpen(path);
	if (!hOpenFolder)
	{
		// Just in case
		RemoveFolderAsRootFolder(path);
	}

	BOOL retval = CanFolderBeOpen(path);
	if (retval)
	{
		hashsum_t hashsum = GlobalFolderInfoDB::GetHashsum(&path);
		FolderInfo* rfinfo = m_globalFolderInfoDB.FindRoot(&path);
		if (rfinfo == 0)
		{
			rfinfo = new FolderInfo(path, FolderInfo::MailFolderNode);
			m_globalFolderInfoDB.AddRoot(&path, rfinfo);
		}

		FolderInfo* finfo = m_globalFolderInfoDB.Find(&path, hashsum);
		if (finfo == 0)
		{
			finfo = new FolderInfo(path, FolderInfo::MailFolderNode);
			m_globalFolderInfoDB.Add(hashsum, finfo);
		}

		hParent = TVI_ROOT;
		BOOL addToRecentFileList = FALSE;
		hItem = OpenFolder(hParent, path, selectFolder, unconditionalFolderInsert, addToRecentFileList, expand);
		if (hItem)
		{
			LoadLabels(hItem);
		}
		else
		{
			BOOL ret = m_folderArray.Delete(path);

			int index = m_folderArray.Find(path);
			FolderInfo* finfo = m_globalFolderInfoDB.FindRoot(&path);

			hashsum_t hashsum = GlobalFolderInfoDB::GetHashsum(&path);
			if (finfo)
				FolderInfo* fi = m_globalFolderInfoDB.Remove(&path, hashsum);
		}
	}
	else
	{
		hashsum_t hashsum = GlobalFolderInfoDB::GetHashsum(&path);
		FolderInfo* finfo = m_globalFolderInfoDB.Find(&path, hashsum);
		CString existingFolderPath;
		CString existingRootFolderPath;
		CString existingFolderType;
		if (finfo == 0)
		{
			MboxMail::assert_unexpected();
		}
		else
		{
			existingFolderType = finfo->FolderTypeToStr();
			existingFolderPath = finfo->m_folderPath;
			if (finfo->m_RootParent)
				existingRootFolderPath = finfo->m_RootParent->m_folderPath;
		}

		CString txt;
		if (existingFolderPath.Compare(existingRootFolderPath) == 0)
		{
#if 0
			txt.Format(L"Can't open folder \n\n\"%s\"\n"
				L"\n\nFolder already open as %s \n\n \"%s\"\n\n"
				L"You must manually remove the root folder first\n"
				L"Select root folder, right-click and select Remove Folder option\n",
				path, existingFolderType, existingRootFolderPath);
#endif
			CString fmt = L"Can't open folder \n\n\"%s\"\n"
				L"\n\nFolder already open as %s \n\n \"%s\"\n\n"
				L"You must manually remove the root folder first\n"
				L"Select root folder, right-click and select Remove Folder option\n";
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, path, existingFolderType, existingRootFolderPath);
		}
		else
		{
#if 0
			txt.Format(L"Can't open folder \n\n\"%s\"\n"
				L"\n\nFolder already open as %s of the root folder \n\n\"%s\"\n\n"
				L"You must manually remove the root folder first\n"
				L"Select root folder, right-click and select Remove Folder option\n",
				path, existingFolderType, existingRootFolderPath);
#endif

			CString fmt = L"Can't open folder \n\n\"%s\"\n"
				L"\n\nFolder already open as %s of the root folder \n\n\"%s\"\n\n"
				L"You must manually remove the root folder first\n"
				L"Select root folder, right-click and select Remove Folder option\n";
			ResHelper::TranslateString(fmt);
			txt.Format(fmt, path, existingFolderType, existingRootFolderPath);
		}

		HWND h = GetSafeHwnd();
		int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONHAND | MB_OK);

		//ArchiveFileInfoMap *fileSizesMap = SetupFileSizeMap(path);
		//m_fileSizesMap = fileSizesMap;

		if (!hOpenFolder)
		{
			BOOL ret = m_folderArray.Delete(path);
		}

		MboxMail::SetLastPath(lastPath);

		TRACE(L"");
		TRACE(L"%s", txt);
	}
	TRACE(L"DoOpenFolder: AssertNTreeView\n");
	BOOL retAssert = AssertNTreeView();
}

HTREEITEM NTreeView::InsertTreeItem(LPCWSTR lpszItem, int nImage, int nSelectedImage, HTREEITEM hParent, HTREEITEM hInsertAfter)
{
	HTREEITEM hItem = m_tree.InsertItem(lpszItem, nImage, nSelectedImage, hParent, hInsertAfter);
	return hItem;
}

HTREEITEM NTreeView::InsertTreeItem(LPCWSTR lpszItem, HTREEITEM hParent, HTREEITEM hInsertAfter)
{
	HTREEITEM hItem = m_tree.InsertItem(lpszItem, hParent, hInsertAfter);
	return hItem;
}

// Called upon  a right-click Folder option
int NTreeView::MergeMailArchiveFiles(HTREEITEM hFolder)
{
	HWND h = GetSafeHwnd();
	MboxMail::ShowHint(HintConfig::MergeFilesHint, h);

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	// TODO: customize CFileDialog to avoid potential buffer overflow and corruption
#define MAX_CFileDialog_FILE_COUNT 1000
	int FILE_LIST_BUFFER_SIZE = (MAX_CFileDialog_FILE_COUNT * (MAX_PATH + 1)) + 1;

	MboxMail::m_outbuf->ClearAndResize(FILE_LIST_BUFFER_SIZE * 2 + 2);
	wchar_t* fileNameBuffer = (LPWSTR)MboxMail::m_outbuf->Data();

	fileNameBuffer[0] = L'\0';

	//CString  sFilters = "Mail Files (*.mbox)|*.mbox|All Files (*.*)|*.*||";
	CString  sFilters = L"Mail Files All Files (*.*)|*.*|(*.mbox;*.eml)|*.mbox;*eml|(*.mbox)|*.mbox|(*.eml)|*.eml||";

restart:
	{
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
		CFileDialog dlgFile(TRUE, NULL, NULL, dwFlags, sFilters);

		CString inFolderPath = MboxMail::GetLastPath();
#if 0
		// No need to enforce the below
		if (!inFolderPath.IsEmpty())
		{
			if (!FileUtils::PathDirExists(inFolderPath)) {
				return -1;
			}
		}
		else {
			return -1;
		}
#endif

		OPENFILENAME& ofn = dlgFile.GetOFN();
		ofn.Flags |= OFN_ALLOWMULTISELECT;
		ofn.lpstrFile = (LPWSTR)fileNameBuffer;
		ofn.nMaxFile = FILE_LIST_BUFFER_SIZE;
		ofn.lpstrInitialDir = inFolderPath;
		ofn.lpstrTitle = L"Select Mail Archives For Merging";


		INT_PTR ret = dlgFile.DoModal();
		if (ret == IDOK)
		{
			CString folder = dlgFile.GetFolderPath();
			TRACE(L"FOLDER=%s\n", folder);

			CArray<MergeFileInfo> fileList;
			int fileCount = 0;
			CString strFilePath;
			POSITION pos = dlgFile.GetStartPosition();
			while (NULL != pos)
			{
				fileCount++;
				strFilePath = dlgFile.GetNextPathName(pos);
				if (!FileUtils::PathFileExist(strFilePath))
				{
#if 0
					CString txt = L"File path invalid.\n\n";
					txt.Append(strFilePath);
					txt.Append(L"\n\nRetry or Cancel?");
#endif

					CString fmt = L"File path invalid.\n\n%s\n\nRetry or Cancel?";
					ResHelper::TranslateString(fmt);
					CString txt;
					txt.Format(fmt, strFilePath);

					int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_RETRYCANCEL);
					if (answer == IDRETRY)
					{
						goto restart;
					}
				}

				// TODO: no need to check if file is mbox type; it was already verified
				int mboxFileType = NTreeView::ImboxviewFile(strFilePath);
				if (mboxFileType < 1)
				{
					int deb = 1;
				}
				else
				{
					MergeFileInfo mergeFileInfo(strFilePath, mboxFileType);
					fileList.Add(mergeFileInfo);
				}

				TRACE(L"FILE=%s\n", strFilePath);
			}

			// All archive files are valid; merge
			CString title = L"Enter Name for New Archive File";
			CString  fileNameFilter = L"Mail Files (*.mbox)|*.mbox||";
			CString dfltExtension = L".mbox";

			CString outFolderPath;
			CString fileName;
			CString datapath = MboxMail::GetLastDataPath();

			BOOL ret = pFrame->SaveFileDialog(fileName, fileNameFilter, dfltExtension, inFolderPath, outFolderPath, title);
			if (ret == TRUE)
			{
				CString filePath = outFolderPath + L"\\" + fileName;
				CString fileExtension = ::PathFindExtension(fileName);
				CString fileName2 = ::PathFindFileName(filePath);
				//FileUtils::SplitFilePath(CString &fileName, CString &driveName, CString &directory, CString &fileNameBase, CString &fileNameExtention);

			//CString dataFolder = outFolderPath + "\\" + CMainFrame::m_ViewerGlobalDB.GetDBFolder();
				CString dataFolder = outFolderPath;
				CString viewFile = dataFolder + L"\\" + fileName + L".mboxview";
				if (FileUtils::PathFileExist(viewFile))
				{
					if (FileUtils::DelFile(viewFile, FALSE))
					{
						int deb = 1;
					}
					int deb = 1;
				}

				if (pFrame->MergeMboxArchiveFiles(fileList, filePath) < 0)
					return -1;

				//CString txt = L"Created Mbox Mail Archive file \n\n" + filePath;
				CString txt = L"Created Mbox Mail Archive file";
				ResHelper::TranslateString(txt);

				BOOL supressOpenFileOption = FALSE;
				//if (filePath.Compare(datapath) == 0)  // TODO: 
					//supressOpenFileOption = TRUE;
				OpenContainingFolderDlg dlg(txt, filePath, supressOpenFileOption);
				INT_PTR nResponse = dlg.DoModal();
				////////////
				if (nResponse == IDOK)
				{
					if (FileUtils::BrowseToFile(filePath) == FALSE) {
						HWND h = GetSafeHwnd();
						HINSTANCE result = ShellExecute(h, L"open", outFolderPath, NULL, NULL, SW_SHOWNORMAL);
						CMainFrame::CheckShellExecuteResult(result, h);
					}
					int deb = 1;
					return 1;
				}
				else if (nResponse == IDYES)
				{
					CString txt = L"Open Created Archive File \n\n" + filePath;
					OpenArchiveFileDlg dlg;
					dlg.m_sourceFolder = outFolderPath;
					dlg.m_targetFolder = inFolderPath;

					CString driveName;
					CString directory;
					CString fileBaseName;
					CString fileNameExtention;

					FileUtils::GetFileBaseNameAndExtension(fileName, fileBaseName, fileNameExtention);

					dlg.m_archiveFileName = fileBaseName + L"_MERGE.mbox";
					INT_PTR nResponse = dlg.DoModal();
					////////////
					if (nResponse == IDOK)
					{
						CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
						if (pFrame)
						{
							CString archiveFilePath = dlg.m_targetFolder + dlg.m_archiveFileName;

							DWORD nFlags = MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH | MOVEFILE_COPY_ALLOWED;
							BOOL retMove = MoveFileEx(filePath, archiveFilePath, nFlags);
							if (retMove == FALSE)
							{
								CString errorText = FileUtils::GetLastErrorAsString();
								TRACE(L"SaveAsMboxArchiveFile_v2: MoveFileEx failed \"%s\"\n", errorText);
#if 0
								CString txt = L"Failed to move file \"" + strFilePath;
								txt += L"\" file.\n\n";
								txt.Append(errorText);
#endif

								CString fmt = L"Failed to move file \"%s\" file.\n\n%s";
								ResHelper::TranslateString(fmt);
								CString txt;
								txt.Format(fmt, strFilePath, errorText);

								HWND h = NULL; // we don't have any window yet
								int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
								return -1;
							}

							NTreeView* pTreeView = pFrame->GetTreeView();
							if (pTreeView)
							{
								// delete index file to make sure it is not used in case old and new length of new mbox file are the same
								// InsertMailFile will delete index file
								//CString indexFile = archiveFilePath + ".mboxview";
								//DeleteFile(indexFile);

								pTreeView->InsertMailFile(archiveFilePath);
							}
						}
						return 1;
					}
					else if (nResponse == IDCANCEL)
					{
						int deb = 1;
					}
					int deb = 1;
				}
				else if (nResponse == IDCANCEL)
				{
					int deb = 1;
				}
			}
			else
			{
				int deb = 1;
			}
			}
		}
	return -1;
	}

void NTreeView::TreeCtrlTest()
{
	HTREEITEM hParent;
	HTREEITEM hItem;
	HTREEITEM child;
	HTREEITEM hRoot;

	CString itemName;
	DWORD nId = 0;

	// Test a few cases and check TVI_ROOT vs NULL.
	// It looks TVI_ROOT is never returned from any call, just NULL.
	// Not exactly sure why TVI_ROOT is needed as input paramter, NULL seems to work also.
	// TODO: Find cases to invalidae above assupmtions.
	hItem = m_tree.GetRootItem();

	if (hItem == TVI_ROOT)
		int deb = 1;

	if (hItem == NULL)
		int deb = 1;

	child = m_tree.GetChildItem(hItem);
	child = m_tree.GetChildItem(TVI_ROOT);

	CString name = L"First";

	// TVI_ROOT can be used in:
	// GetNextItem(TVI_ROOT, TVGN_CARET);
	// HTREEITEM InsertItem(LPCWSTR lpszItem, HTREEITEM hParent = TVI_ROOT, HTREEITEM hInsertAfter = TVI_LAST);
	// SortChildren(TVI_ROOT);
	// DeleteItem(TVI_ROOT)
	// GetChildItem(TVI_ROOT);


	hParent = TVI_ROOT;

	hItem = m_tree.GetNextItem(TVI_ROOT, TVGN_CHILD);
	//hItem = InsertTreeItem(name, TVI_ROOT);
	hItem = InsertTreeItem(name, 0);  // either 0 or TVI_ROOT work fine
	hItem = m_tree.GetNextItem(TVI_ROOT, TVGN_CHILD);
	hItem = m_tree.GetChildItem(TVI_ROOT);
	hItem = m_tree.GetRootItem();
	if (hItem)
	{
		name = m_tree.GetItemText(hItem);
		hParent = m_tree.GetParentItem(hItem);
		child = m_tree.GetChildItem(hItem);
		hRoot = m_tree.GetRootItem();
		int deb = 1;
	}

	if (hItem == TVI_ROOT)
		int deb = 1;

	if (hItem == NULL)
		int deb = 1;

	hItem = m_tree.GetRootItem();
	name = m_tree.GetItemText(hItem);

	name = L"First.First";
	HTREEITEM hItem2 = m_tree.InsertItem(name, hItem);

	name = L"Second";

	hParent = TVI_ROOT;
	hParent = 0;
	hItem = m_tree.InsertItem(name, hParent);
	if (hItem)
		name = m_tree.GetItemText(hItem);

	HTREEITEM hItem1 = hItem;
	if (hItem == TVI_ROOT)
		int deb = 1;

	if (hItem == NULL)
		int deb = 1;

	hRoot = TVI_ROOT;
	hRoot = 0;
	BOOL bSort = m_tree.SortChildren(hRoot);

	name = L"Second.Second";
	hItem2 = m_tree.InsertItem(name, hItem);


	CString path = L"";
	HTREEITEM hFolder = NTreeView::HasFolder(path);

	// Below returns the first child of the root. Strange that I can use 0 instead of TVI_ROOT and it works
	hItem = m_tree.GetRootItem();
	//
	hItem = m_tree.GetNextItem(TVI_ROOT, TVGN_CHILD);
	hItem = m_tree.GetNextItem(0, TVGN_CHILD);
	//
	hItem = m_tree.GetChildItem(TVI_ROOT);
	hItem = m_tree.GetChildItem(0);

	// BOOL ret = m_tree.ItemHasChildren(TVI_ROOT);  // crash
	// BOOL ret = m_tree.ItemHasChildren(0);  // suspect retuern value


	child = m_tree.GetChildItem(0);
	while (child)
	{
		name = m_tree.GetItemText(child);
		TRACE(L"TreeCtrlTest: name=%s\n", name);
		UINT nCode = TVGN_NEXT;
		child = m_tree.GetNextItem(child, nCode);
	}

	BOOL recursive = TRUE;
	int depth = 0;
	TRACE(L"TreeCtrlTest: Tree Control:\n");
	hRoot = 0;
	PrintTreeCtrl(hRoot, recursive, depth);

	m_tree.DeleteItem(hItem1);

	TRACE(L"TreeCtrlTest: Tree Control:\n");
	PrintTreeCtrl(hRoot, recursive, depth);

	m_tree.DeleteItem(0);

	TRACE(L"TreeCtrlTest: Tree Control:\n");
	PrintTreeCtrl(hRoot, recursive, depth);

	hItem = m_tree.GetRootItem();

	int deb = 1;
}

int NTreeView::RefreshSubFolder(HTREEITEM hFolder)
{
	CString name = m_tree.GetItemText(hFolder);
	DWORD nId = (DWORD)m_tree.GetItemData(hFolder);

	LabelInfo* linfo = m_labelInfoStore.Find(nId);
	if (linfo == 0)
	{
		MboxMail::assert_unexpected();
		return 0;
	}
	if (linfo->m_nodeType != LabelInfo::MailSubFolder)
	{
		MboxMail::assert_unexpected();
		return 0;
	}
	CString folderPath = linfo->m_mailFolderPath;

	int removedCount = PruneFolder(hFolder);
	int newCount = RefreshMboxFilesList(folderPath, hFolder);
	return 1;
}

int NTreeView::RefreshMboxFilesList(CString& folderPath, HTREEITEM hFolder)
{
	CString fw = folderPath + L"\\*.*";
	WIN32_FIND_DATA	wf;
	BOOL found;
	// Iterate all files in the lastPath folder
	// Valid new archives file is addedd  to fileSizes hash table and
	// Index file.mboxview file is removed  to force parsing to create new index file
	//
	HANDLE f = FindFirstFile(fw, &wf);
	if (f != INVALID_HANDLE_VALUE)
	{
		do
		{
			if ((wf.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != FILE_ATTRIBUTE_DIRECTORY && wf.cFileName[0] != '.')
			{
				CString fn = wf.cFileName;
				CString mboxFilePath = folderPath + L"\\" + fn;

				if (!ImboxviewFileFilter(fn))
					continue;

				_int64 realFSize = FileUtils::FileSize(mboxFilePath);

				ArchiveFileInfoMap* fileSizesMap = SetupFileSizeMap(mboxFilePath);
				m_fileSizesMap = fileSizesMap;
				if (fileSizes)
				{
					ArchiveFileInfo info;
					found = (*fileSizes).Lookup(fn, info);
					if (found)
					{
						if (info.fSize != realFSize)
						{
							RemoveMboxAssociatedWorkData(hFolder, mboxFilePath);

							(*fileSizes)[fn].fSize = realFSize;
							(*fileSizes)[fn].bShow = TRUE;
							(*fileSizes)[fn].bNeedsValidation = TRUE;
							m_bIsDataDirty = TRUE;
							int deb = 1;
						}
					}
					else if (NTreeView::ImboxviewFile(mboxFilePath))
					{
#if 0
						// No longer needed
						CString fileNameExtention;
						CString fileNameBase;
						FileUtils::GetFileBaseNameAndExtension(fn, fileNameBase, fileNameExtention);

						wchar_t c = fileNameBase.GetAt(fileNameBase.GetLength() - 1);
						if ((c == ' ') || (c == '\t'))
						{
							CString txt = L"Invalid mail file name \n\n\"" + fn;
							txt += L"\".\n\nBase name of the file (name without the extension) can't have trailing white spaces. Ignoring.\n";
							int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
							continue;
						}
#endif

						(*fileSizes)[fn].fSize = realFSize;
						(*fileSizes)[fn].bShow = TRUE;
						(*fileSizes)[fn].bNeedsValidation = TRUE;
						m_bIsDataDirty = TRUE;

						InsertMailFile(mboxFilePath);
						int deb = 1;
					}
				}
				else
				{
					; // ignore  MboxMail::assert_unexpected();
				}
			}
		} while (FindNextFile(f, &wf));
		FindClose(f);

		if (m_bIsDataDirty)
		{
			SaveData(hFolder);
		}
	}
	return 1;
}

int NTreeView::PruneFolder(HTREEITEM hItem)
{
	CString path;
	CString name;
	DWORD nId;
	LabelInfo* linfo;
	//FolderInfo *finfo;
	CString filePath;

	if (hItem == TVI_ROOT)
		hItem = 0;

	if (m_tree.ItemHasChildren(hItem))
	{
		HTREEITEM hChild = m_tree.GetChildItem(hItem);
		while (hChild)
		{
			name = m_tree.GetItemText(hChild);
			nId = (DWORD)m_tree.GetItemData(hChild);

			linfo = m_labelInfoStore.Find(nId);
			if (linfo == 0)
			{
				MboxMail::assert_unexpected();
				goto NEXT;
}
			if (linfo->m_nodeType != LabelInfo::MailFile)
			{
				goto NEXT;
			}
			filePath = linfo->m_filePath;
			if (FileUtils::PathFileExist(filePath))
			{
				goto NEXT;
			}
			// TODO: remove file from Label Infos
			RemoveMboxAssociatedWorkData(hChild);

			// It will remove mabox file and labels from Tree View
			// TODO: doesn't remove associated LabelInfo
			BOOL retval = NTreeView::DeleteItem(hChild);

		NEXT:		hChild = m_tree.GetNextSiblingItem(hChild);
}
}
	return 0;
}

int NTreeView::RemoveMboxAssociatedWorkData(HTREEITEM hFolder, CString& filePath)
{
	int ret = -1;
	HTREEITEM hItem = 0;
	CString path = filePath;

	path.TrimRight(L"\\");
	path.Append(L"\\");

	CString fileName;
	CString folderPath;
	FileUtils::GetFolderPathAndFileName(path, folderPath, fileName);
	folderPath.TrimRight(L"\\");
	folderPath.Append(L"\\");

	HTREEITEM hFolder2 = FindFolder(path);
	if (hFolder != hFolder)
		MboxMail::assert_unexpected();

	if (hFolder)
	{
		CString fileName;
		CString folderPath;
		CString filePath = path;
		filePath.TrimRight(L"\\");
		//FileUtils::CPathStripPath(path, fileName);
		FileUtils::GetFolderPathAndFileName(filePath, folderPath, fileName);
		hItem = NTreeView::FindItem(hFolder, fileName);
	}
	if (hItem)
		ret = RemoveMboxAssociatedWorkData(hItem);


	return ret;
}

int NTreeView::RemoveMboxAssociatedWorkData(HTREEITEM hItem)
{
	CString name = m_tree.GetItemText(hItem);
	DWORD nId = (DWORD)m_tree.GetItemData(hItem);

	LabelInfo* linfo = m_labelInfoStore.Find(nId);
	if (linfo == 0)
	{
		MboxMail::assert_unexpected();
		return -1;
	}
	if (linfo->m_nodeType != LabelInfo::MailFile)
	{
		MboxMail::assert_unexpected();
		return -1;
	}
	CString filePath = linfo->m_filePath;

	CString mailFile = MboxMail::s_path;

	CString datapath = MboxMail::GetLastDataPath();
	if (datapath.IsEmpty())
	{
		MboxMail::assert_unexpected();
		return -1;
	}

	CString fileName;
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	FileUtils::SplitFilePath(filePath, driveName, directory, fileNameBase, fileNameExtention);

	CString folderPath = driveName + directory;
	fileName = fileNameBase + fileNameExtention;

	CString mboxviewFilePath = datapath + fileName + L".mboxview";
	FileUtils::DelFile(mboxviewFilePath, FALSE);

	CString lastPath = MboxMail::GetLastPath();

	ArchiveFileInfoMap* fileSizesMap = SetupFileSizeMap(lastPath);
	m_fileSizesMap = fileSizesMap;

	if (fileSizes)
	{
		ArchiveFileInfo info;
		BOOL found = (*fileSizes).Lookup(fileName, info);
		if (found)
		{
			(*fileSizes).RemoveKey(fileName);
			SaveData(hItem);
		}
	}
	else
		MboxMail::assert_unexpected();

#if 1
	CString cacheFolderPath;
	cacheFolderPath = datapath + fileNameBase;
	BOOL retval = FileUtils::RemoveDir(cacheFolderPath, true, true);
	if (retval != TRUE)
		_ASSERTE(FALSE);
#else
	CArray<CString> cacheFolderList;
	NTreeView::SetupCacheFolderList(cacheFolderList);

	CString cacheFolderName;
	CString cacheFolderPath;

	for (int i = 0; i < cacheFolderList.GetCount(); i++)
	{
		cacheFolderName = cacheFolderList[i];
		cacheFolderPath = datapath + fileNameBase + L"\\" + cacheFolderName;
		BOOL retval = FileUtils::RemoveDir(cacheFolderPath, true, true);

		int deb = 1;
	}
#endif

	return 1;
}

// Check if item exist in tree folder and select
HTREEITEM NTreeView::SelectFolderItem(CString& mboxFilePath)
{
	CString mboxFolderPath;
	CString mboxFileName;
	FileUtils::GetFolderPathAndFileName(mboxFilePath, mboxFolderPath, mboxFileName);

	HTREEITEM hFolder = HasFolder(mboxFolderPath);
	if (hFolder)
	{
		HTREEITEM hItem = FindItem(hFolder, mboxFileName);
		UINT nCode = TVGN_CARET;
		if (hItem)
		{
			BOOL retval = SelectTreeItem(hItem);

			CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
			NListView* pListView = pFrame->GetListView();
			if (pListView)
			{
				if (MboxMail::s_mails.GetCount() > 0)
				{
					pListView->SelectItemFound(0);
				}
				pListView->SetListFocus();
			}
			return hItem;
		}
	}
	return 0;
}

//////////// MBoxFolderTree 
//////////////////////////////////////////

MBoxFolderNode* MBoxFolderTree::CreateNode(MBoxFolderNode* node)
{
	if (node == 0)
	{
		m_root = &m_rootNode;
		return m_root;
	}
	else
	{
		MBoxFolderNode newNode;
		newNode.m_parent = node;
		node->m_nodeList.push_back(newNode);
		return &node->m_nodeList.back();
	}
}

void MBoxFolderTree::PrintNode(CFile* fp, MBoxFolderNode* node)
{
	if (node == 0)
		return;

	if (node->m_nodeList.empty())
	{
		TRACE(L"%s (Leave Node)\n", node->m_folderName);
		return;
	}
	else
	{
		TRACE(L"%s\n", node->m_folderName);
		list<MBoxFolderNode>::iterator it;
		for (it = node->m_nodeList.begin(); it != node->m_nodeList.end(); it++)
		{
			PrintNode(fp, &(*it));
			int deb = 1;
		}
	}
}

void MBoxFolderTree::Print(CString& filePath)
{
#ifdef _DEBUG
	CFile fp;
	CString path = filePath;
	path.TrimRight(L"\\");
	path.Append(L"\\");

	CString lastPath = MboxMail::GetLastPath();

	// TODO: SetLastPath should not be needed here, revisit
	MboxMail::SetLastPath(path);

	CString datapath = MboxMail::GetLastDataPath();
	FileUtils::CreateDir(datapath);
	datapath.TrimRight(L"\\");

	CFileException ExError;
	CString viewFile = datapath + "\\.folderview";
#if 0
	// TODO: don't write to file to now
	if (!fp.Open(viewFile, CFile::modeWrite | CFile::modeCreate, &ExError))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not create \"" + viewFile;
		txt += L"\" file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);
	}
	else
	{
		TRACE(L"MBoxFolderTree: Folder List:\n");
		PrintNode(&fp, m_root);
		fp.Close();
	}
#else
	TRACE(L"MBoxFolderTree: Folder List:\n");
	PrintNode(&fp, m_root);
#endif

	MboxMail::SetLastPath(lastPath);
#endif
}

int MBoxFolderTree::NodeCount(MBoxFolderNode* node)
{
	int count = 0;

	if (node == 0)
		return count;

	if (node->m_nodeList.empty())
	{
		TRACE(L"MBoxFolderTree NodeCount: %s (Leave Node)\n", node->m_folderName);
		return count;
	}
	else
	{
		TRACE(L"MBoxFolderTree NodeCount: %s\n", node->m_folderName);
		list<MBoxFolderNode>::iterator it;
		for (it = node->m_nodeList.begin(); it != node->m_nodeList.end(); it++)
		{
			count += NodeCount(&(*it));
			int deb = 1;
		}
		count += (int)node->m_nodeList.size();
	}
	return count;
}

int MBoxFolderTree::Count()
{
	int count = NodeCount(m_root);
	if (m_root)
		count++;
	TRACE(L"MBoxFolderTree: Tree Folder count: %d\n", count);
	return count;
}


// Example how to prune to be fully implemented in NTreeView
// Example will remove all folders
void MBoxFolderTree::PruneNonMBoxFolderNode(MBoxFolderNode* node)
{
	if (node == 0)
		return;

	list<MBoxFolderNode>::iterator it;
	list<MBoxFolderNode>::iterator next_it;
	CString folderName;

	if (node->m_nodeList.empty())
	{
		//TRACE("%s (Leave Node)\n", node->m_folderName);
		// Mark node for removal
		node->m_valid = FALSE;
		return;
	}
	else
	{
		//TRACE("%s\n", node->m_folderName);
		for (it = node->m_nodeList.begin(); it != node->m_nodeList.end();)
		{
			next_it = it;
			next_it++;
			const wchar_t* p = it->m_folderName;
			folderName = it->m_folderName;

			PruneNonMBoxFolderNode(&(*it));
			it = next_it;

			int deb = 1;
		}

		for (it = node->m_nodeList.begin(); it != node->m_nodeList.end();)
		{
			next_it = it;
			next_it++;
			const wchar_t* p = it->m_folderName;
			folderName = it->m_folderName;
			if (!it->m_valid)
				node->m_nodeList.erase(it);

			it = next_it;
			int deb = 1;
		}

		if (node->m_nodeList.empty())
		{
			//TRACE("Leave Node: %s\n", node->m_folderName);
			// Mark node for removal
			node->m_valid = FALSE;
		}
		else
		{
			; // Keep the node 
		}
	}
}

void MBoxFolderTree::GetRelativeFolderPath(MBoxFolderNode* rnode, CString& folderPath)
{
	_ASSERTE(rnode);
	MBoxFolderNode* parent = rnode->m_parent;
	MBoxFolderNode* rootParent = parent;

	while (parent)
	{
		rootParent = parent;
		parent = parent->m_parent;
	}
	if (rootParent)
	{
		CString path = rootParent->m_folderName;
		folderPath = rnode->m_folderName.Mid(path.GetLength());
	}
	folderPath.Trim();
	folderPath.Trim(L"\\");
	return;
}

void MBoxFolderTree::PruneNonMBoxFolders()
{
	//TRACE("MBoxFolderTree::Pruned Folder List:\n");

	PruneNonMBoxFolderNode(m_root);

	if (m_root && !m_root->m_valid)
		m_root = 0;
}


// Traverse root and subfolders and populkate tree
BOOL MBoxFolderTree::PopulateFolderTree(CString& rootFolder, MBoxFolderTree& tree, MBoxFolderNode* rnode, CString& errorText, int maxDepth)
{
	BOOL bRetval = TRUE;

	// TODO: Hash Table ??
	CArray<CString> cacheFolderList;
	NTreeView::SetupCacheFolderList(cacheFolderList);

	rootFolder.TrimRight(L"\\");
	//rootFolder.Append("\\");

	CString folderName;
	FileUtils::CPathStripPath(rootFolder, folderName);
	for (int i = 0; i < cacheFolderList.GetCount(); i++)
	{
		CString& cacheFolderName = cacheFolderList[i];
		if (folderName.Compare(cacheFolderName) == 0)
		{
			// Skip folders such as printCache, etc
			return TRUE;
		}
	}

	CString folderPath = rootFolder;
	WIN32_FIND_DATA FileData;
	HANDLE hSearch;
	BOOL	bFinished = FALSE;
	int fileCnt = 0;  // Count of any files under the current folder. Not the global count of all files

	MBoxFolderNode* node = tree.CreateNode(rnode);
	node->m_folderName = rootFolder;

	// Start searching for all folders in the current directory.
	CString searchPath = folderPath + L"\\*.*";
	hSearch = FindFirstFile(searchPath, &FileData);
	if (hSearch == INVALID_HANDLE_VALUE)
	{
		// TODO: should never be true ??
		CString erText = FileUtils::GetLastErrorAsString();
		TRACE(L"PopulateFolderTree: %s.\n", erText);
		return TRUE;  // TODO: TRUE or FALSE
	}

	//CString folderName;
	CString fileName;
	while (!bFinished)
	{
		if (!(_tcscmp(FileData.cFileName, L".") == 0 || _tcscmp(FileData.cFileName, L"..") == 0))
		{
			fileName = CString(FileData.cFileName);
			CString fileFound = folderPath + L"\\" + fileName;

			if (FileData.dwFileAttributes == FILE_ATTRIBUTE_DIRECTORY)
			{
				fileFound.TrimRight(L"\\");
				fileFound.Append(L"\\");
				bRetval = PopulateFolderTree(fileFound, tree, node, errorText, maxDepth - 1);
				if (!errorText.IsEmpty())
					break;
			}
			else
				fileCnt++;
		}
		if (!FindNextFile(hSearch, &FileData)) {
			bFinished = TRUE;
			break;
		}
	}
	FindClose(hSearch);

	return bRetval;
}

//

CString FolderInfo::FolderTypeToStr()
{
	if (m_folderType == MailFolderNode)
		return L"Folder";
	if (m_folderType == MailRootFolderNode)
		return L"Root Folder";
	if (m_folderType == MailSubFolderNode)
		return L"Sub-Folder";
	else
		return L"Unknown Folder Type";
}

GlobalFolderInfoDB::GlobalFolderInfoDB()
{
	m_allFoldersTable = new AllFoldersInfoHashTable(513);
	m_allRootFoldersTable = new AllFoldersInfoHashTable(113);
	m_rootInfo = 0;
	m_deleteConflictingFolders = FALSE;
};

GlobalFolderInfoDB::~GlobalFolderInfoDB()
{
	DeleteAll();
};

void GlobalFolderInfoDB::Print()
{
#ifdef _DEBUG
	AllFoldersInfoHashTable::IHashMapIter iter = m_allRootFoldersTable->first();
	int totalCnt = 0;
	CString txt;
	CString folderType;

	TRACE(L"GLOBAL Folder root Table\n");

	for (; !m_allRootFoldersTable->last(iter); )
	{
		FolderInfo* l = iter.element;

		folderType = l->FolderTypeToStr();

		txt.Format(L"Type=%s FolderPath=%s\n", folderType, l->m_folderPath);
		TRACE(L"%s", txt);
		m_allRootFoldersTable->next(iter);
		int deb = 1;
	}

	//
	folderType.Empty();

	TRACE(L"+++ GLOBAL Folder & SubFolder Table\n");
	iter = m_allFoldersTable->first();
	for (; !m_allFoldersTable->last(iter); )
	{
		folderType = L"Unknown";
		FolderInfo* l = iter.element;

		folderType = l->FolderTypeToStr();

		txt.Format(L"Type=%s FolderPath=%s\n", folderType, l->m_folderPath);
		TRACE(L"%s", txt);
		m_allFoldersTable->next(iter);
		int deb = 1;
	}
#endif
}

hashsum_t GlobalFolderInfoDB::GetHashsum(CString* key)
{
	hashsum_t hashsum = StrHash((const char*)(LPCWSTR)*key, key->GetLength()*2);
	return hashsum;
};

FolderInfo* GlobalFolderInfoDB::Remove(CString* key)
{
	hashsum_t hashsum = GlobalFolderInfoDB::GetHashsum(key);
	FolderInfo* finfo = Remove(key, hashsum);
	return finfo;
}

FolderInfo* GlobalFolderInfoDB::Remove(CString* key, hashsum_t hashsum)
{
	FolderInfo* finfo = m_allFoldersTable->remove(key, hashsum);
	FolderInfo* rfinfo = GlobalFolderInfoDB::FindRoot(key);
	if (rfinfo != finfo->m_RootParent)
		MboxMail::assert_unexpected();

	finfo->m_RootParent->m_folderList.remove(finfo);
	if (finfo->m_RootParent->m_folderList.count() == 0)
		m_allRootFoldersTable->remove(key);

	return finfo;
}

FolderInfo* GlobalFolderInfoDB::Find(CString* key)
{
	FolderInfo* finfo = m_allFoldersTable->find(key);
	return finfo;
}

FolderInfo* GlobalFolderInfoDB::Find(CString* key, hashsum_t hashsum)
{
	//hashsum_t hashsum = GlobalFolderInfoDB::GetHashsum(key);
	FolderInfo* finfo = m_allFoldersTable->find(key, hashsum);
	return finfo;
}

void GlobalFolderInfoDB::Add(CString* key, FolderInfo* info)
{
	if (m_rootInfo == 0) MboxMail::assert_unexpected();
	info->m_RootParent = m_rootInfo;
	m_allFoldersTable->insert(key, info);
	m_rootInfo->m_folderList.insert_head(info);
}


void GlobalFolderInfoDB::Add(hashsum_t hashsum, FolderInfo* info)
{
	if (m_rootInfo == 0)
		MboxMail::assert_unexpected();
	info->m_RootParent = m_rootInfo;
	m_allFoldersTable->insert(hashsum, info);
	m_rootInfo->m_folderList.insert_head(info);
}

FolderInfo* GlobalFolderInfoDB::FindRoot(CString* key)
{
	FolderInfo* finfo = m_allRootFoldersTable->find(key);
	m_rootInfo = finfo;
	return finfo;
}

void GlobalFolderInfoDB::AddRoot(CString* key, FolderInfo* info)
{
	m_allRootFoldersTable->insert(key, info);
	m_rootInfo = info;
}

void GlobalFolderInfoDB::RemoveRootFolder(CString& path)
{
	if ((m_allFoldersTable == 0) || (m_allRootFoldersTable == 0))
	{
		MboxMail::assert_unexpected();
		return;
	}

	Print();


	CString folderPath = path;
	folderPath.TrimRight(L"\\");
	folderPath.Append(L"\\");
	FolderInfo* r = m_allRootFoldersTable->find(&folderPath);
	if (r == 0)
	{
		MboxMail::assert_unexpected();
		return;
	}

	FolderInfo* finfo;
	while (finfo = r->m_folderList.remove_head())
	{
		m_allFoldersTable->remove(&finfo->m_folderPath);
		delete finfo;
	}

	m_allRootFoldersTable->remove(&folderPath);
	delete r;
}

void GlobalFolderInfoDB::Clear()
{
	if ((m_allFoldersTable == 0) || (m_allRootFoldersTable == 0))
	{
		MboxMail::assert_unexpected();
		return;
	}

	AllFoldersInfoHashTable::IHashMapIter riter = m_allRootFoldersTable->first();
	for (; !m_allRootFoldersTable->last(riter); )
	{
		FolderInfo* l = riter.element;
		m_allRootFoldersTable->remove(riter);
		delete l;
	}
	m_allRootFoldersTable->clear();

	AllFoldersInfoHashTable::IHashMapIter iter = m_allFoldersTable->first();
	for (; !m_allFoldersTable->last(iter); )
	{
		FolderInfo* l = iter.element;
		m_allFoldersTable->remove(iter);
		delete l;
	}
	m_allFoldersTable->clear();
}

void GlobalFolderInfoDB::DeleteAll()
{
	Clear();
	delete m_allFoldersTable;
	m_allFoldersTable = 0;
	//
	delete m_allRootFoldersTable;
	m_allRootFoldersTable = 0;
}

/////

//

CString TreeCtrlInfo::TreeCtrlTypeToStr()
{
	if (m_fileType == MailFolder)
		return L"MailFolder";
	if (m_fileType == MailSubFolder)
		return L"MailSubFolder";
	if (m_fileType == MailFile)
		return L"MailFile";
	if (m_fileType == MailLabel)
		return L"MailLabelFile";
	else
		return L"Unknown File Type";
}

TreeCtrlInfoDB::TreeCtrlInfoDB()
{
	m_treeCtrlTable = new TreeCtrlInfoHashTable(50013);
};

TreeCtrlInfoDB::~TreeCtrlInfoDB()
{
	DeleteAll();
};

void TreeCtrlInfoDB::Print()
{
#ifdef _DEBUG
	TreeCtrlInfoHashTable::IHashMapIter iter = m_treeCtrlTable->first();
	int totalCnt = 0;
	CString txt;
	CString fileType;

	//
	fileType.Empty();

	TRACE(L"+++ GLOBAL TreeCtrl Table\n");
	iter = m_treeCtrlTable->first();
	for (; !m_treeCtrlTable->last(iter); )
	{
		fileType = L"Unknown";
		TreeCtrlInfo* l = iter.element;

		fileType = l->TreeCtrlTypeToStr();

		txt.Format(L"Type=%s FolderPath=%s\n", fileType, l->m_filePath);
		TRACE(L"%s", txt);
		m_treeCtrlTable->next(iter);
		int deb = 1;
	}
#endif
}

hashsum_t TreeCtrlInfoDB::GetHashsum(CString* key)
{
	hashsum_t hashsum = StrHash((const char*)(LPCWSTR)*key, key->GetLength()*2);
	return hashsum;
};

TreeCtrlInfo* TreeCtrlInfoDB::Remove(CString* key)
{
	hashsum_t hashsum = GlobalFolderInfoDB::GetHashsum(key);
	TreeCtrlInfo* finfo = Remove(key, hashsum);
	return finfo;
}

TreeCtrlInfo* TreeCtrlInfoDB::Remove(CString* key, hashsum_t hashsum)
{

	TreeCtrlInfo* finfo = m_treeCtrlTable->remove(key);

	return finfo;
}

TreeCtrlInfo* TreeCtrlInfoDB::Find(CString* key)
{
	TreeCtrlInfo* finfo = m_treeCtrlTable->find(key);
	return finfo;
}

TreeCtrlInfo* TreeCtrlInfoDB::Find(CString* key, hashsum_t hashsum)
{
	TreeCtrlInfo* finfo = m_treeCtrlTable->find(key, hashsum);
	return finfo;
}

void TreeCtrlInfoDB::Add(CString* key, TreeCtrlInfo* info)
{
	m_treeCtrlTable->insert(key, info);
}


void TreeCtrlInfoDB::Add(hashsum_t hashsum, TreeCtrlInfo* info)
{
	m_treeCtrlTable->insert(hashsum, info);
}

void TreeCtrlInfoDB::Clear()
{
	if (m_treeCtrlTable == 0)
	{
		MboxMail::assert_unexpected();
		return;
	}

	TreeCtrlInfoHashTable::IHashMapIter iter = m_treeCtrlTable->first();
	for (; !m_treeCtrlTable->last(iter); )
	{
		TreeCtrlInfo* l = iter.element;
		m_treeCtrlTable->remove(iter);
		delete l;
	}
	m_treeCtrlTable->clear();
}

void TreeCtrlInfoDB::DeleteAll()
{
	Clear();
	//
	delete m_treeCtrlTable;
	m_treeCtrlTable = 0;
}


/////


MySaveFileDialog::MySaveFileDialog(
	BOOL bOpenFileDialog, LPCWSTR lpszDefExt, LPCWSTR lpszFileName,
	DWORD dwFlags, LPCWSTR lpszFilter, CWnd* pParentWnd, DWORD dwSize, BOOL bVistaStyle
) : CFileDialog(bOpenFileDialog, lpszDefExt, lpszFileName, dwFlags, lpszFilter, pParentWnd, dwSize, bVistaStyle)
{
	int deb = 1;
}

BOOL MySaveFileDialog::OnFileNameOK()
{
	OPENFILENAME& ofn = GetOFN();

	CString initialFolderPath = ofn.lpstrInitialDir;
	CString folderPath = GetFolderPath();
	if (initialFolderPath.Compare(folderPath))
	{
		CString txt = L"Folder path can't be changed.";
		ResHelper::TranslateString(txt);
		int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);

		UpdateOFNFromShellDialog();
		return TRUE;
	}

	TRACE(L"FOLDER=%s\n", folderPath);

	int fileCount = 0;

	if (fileCount > 0)
		return FALSE;
	else
		return TRUE;
}

LRESULT NTreeView::OnCmdParam_ResetTreePos(WPARAM wParam, LPARAM lParam)
{
	LRESULT lres = m_tree.SendMessage(WM_HSCROLL, SB_LEFT, 0);

	// call DrawMenuBar if it is updated (text, etc)
#if 0
	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	pFrame->DrawMenuBar();
#endif

	return 0;
}

// Labels View
//  create label or no label based tree view
void NTreeView::DoOpenRootFolderAndSubfolders_LabelView(CString& path, BOOL selectFolder)
{
	OpenRootFolderAndSubfolders_LabelView(path, selectFolder);
}

//  create label or no label based tree view
// path == root path
// selectFolder==?? MergeFIXME
void NTreeView::OpenRootFolderAndSubfolders_LabelView(CString& path, BOOL selectFolder)
{
	CString rootFolder = path;
	CList<CString, CString&> folderList;
	CString errorText;
	int maxDepth = 10;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	int paneId = 0;
	CString sText;
	CString fname;
	FileUtils::CPathStripPath(path, fname);
	//sText.Format(L"Merging folder \"%s\" and all sub-folders ...", fname);

	CString fmt = L"Merging folder \"%s\" and all sub-folders ...";
	ResHelper::TranslateString(fmt);
	sText.Format(fmt, fname);

	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	// Create empty tree, set root tree path
	MBoxFolderTree tree(path);

	MBoxFolderNode* rnode = 0;
	errorText.Empty();
	// Traverse root and subfolders and populkate tree
	BOOL bRet = tree.PopulateFolderTree(path, tree, rnode, errorText, maxDepth);
	if (!errorText.IsEmpty())
	{
		int answer = MessageBox(errorText, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;
	}
	int treeFolderCount = tree.Count();

	TRACE(L"\n\nBefore Prunning\n\n");
	tree.Print(path);


	// Delete all folder that don't have any mbox files
	PruneFolderTree(tree);

	treeFolderCount = tree.Count();

	TRACE(L"\n\nAfter Prunning\n\n");
	tree.Print(path);

	if (treeFolderCount == 0)
	{
		//errorText.Format(L"No folders with mbox files found under the root folder and all sub-folders:\n\n%s", path);

		CString fmt = L"No folders with mbox files found under the root folder and all sub-folders:\n\n%s";
		ResHelper::TranslateString(fmt);
		errorText.Format(fmt, path);

		int answer = MessageBox(errorText, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		return;
	}

	if (pFrame)
	{
		NListView* pListView = pFrame->GetListView();
		if (pListView)
		{
			pListView->m_path_label = L"";
			pListView->m_path = L"";
			pListView->m_which = NULL;
			pListView->FillCtrl();
		}
	}

	MboxMail::SetLastPath(path);

	errorText.Empty();
	// TODO: Investigate if separate MergeTreeFolders_NoLabels() would be cleaner solution
	// Merge mbox files, follow user selected configuration 
	int retM = MergeTreeFolders(tree, errorText);
	if (retM < 0)
		return;

	MboxMail::SetLastPath(path);
	if (selectFolder)
		HTREEITEM hLastFolder = SelectTreeItem(&path);

	m_globalFolderInfoDB.Print();

	sText.Format(L"Ready", fname);
	ResHelper::TranslateString(sText);
	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);
}


// Finally, do Merge all mbox files discovered under node and subfolders
int NTreeView::MergeTreeNode(MBoxFolderNode* node)
{
	if (node == 0)
		return 0;

	CString folderName;
	CString folderPath = node->m_folderName;
	FileUtils::CPathStripPath(folderPath, folderName);

	CString relativeFolderPath;
	// get relative path to node->m_folderName path MergeFIXME
	MBoxFolderTree::GetRelativeFolderPath(node, relativeFolderPath);
	relativeFolderPath.Replace(L'\\', L'/');

	BOOL addToRecentFileList = FALSE;
	BOOL expand = FALSE;
	// Merge all mbox file under given path, return list of folders for further processing
	int retM = MergeRootSubFolder(relativeFolderPath, folderPath, addToRecentFileList, expand);
	if (retM < 0)
	{
		return -1;
	}

	CString itemFolderName;

	if (node->m_nodeList.empty())
	{
		TRACE(L"%s (Leave Node)\n", node->m_folderName);
		return 0;
	}
	else
	{
		TRACE(L"%s\n", node->m_folderName);
		list<MBoxFolderNode>::iterator it;
		// process all folder under current path
		for (it = node->m_nodeList.begin(); it != node->m_nodeList.end(); it++)
		{
			MergeTreeNode(&(*it));
			int deb = 1;
		}
	}
	return 0;
}


// Merge all mbox files in root and subfolers  MergeFIXME
//
int NTreeView::MergeTreeFolders(MBoxFolderTree& tree, CString& errorText)
{
	TRACE(L"MergeTreeFolders: OpenTreeFolders:\n");

	MBoxFolderNode* rootNode = tree.GetRootNode();
	if (rootNode == 0)
		return 1;

	CString lastPath = MboxMail::GetLastPath();
	lastPath.TrimRight(L"\\");

	CString lastDataPath = MboxMail::GetLastDataPath();

	CString folderPath;
	CString fileName;

	FileUtils::GetFolderPathAndFileName(lastPath, folderPath, fileName);

	m_rootMboxFilePath = lastPath + L"\\" + fileName + L".mbox";

	CString m_rootMboxFolderPath = lastDataPath + L"MergeCache";   // FIXMEFIXME
	FileUtils::CreateDir(m_rootMboxFolderPath);
	m_rootMboxFilePath = m_rootMboxFolderPath + L"\\" + fileName + L".mbox";


	// Open destination file for merging all mbox files
	CFileException exMergeTo;
	if (!m_rootMboxCfile.Open(m_rootMboxFilePath, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &exMergeTo))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exMergeTo);
#if 0
		CString txt = L"Could not create Merge To File \"" + m_rootMboxFilePath;
		txt += L"\" file.\n";
		txt += exErrorStr;
#endif

		CString fmt = L"Could not create Merge To File \"%s\" file.\n%s";
		ResHelper::TranslateString(fmt);
		CString txt;
		txt.Format(fmt, m_rootMboxFilePath, exErrorStr);

		//TRACE(L"%s\n", txt);

		CFileStatus rStatus;
		BOOL ret = m_rootMboxCfile.GetStatus(rStatus);;

		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);

		return -1;
	}

	// Merge all mbox files discovered under rootNode and subfolders
	// Duplicate mails are not removed yet
	int hIt = MergeTreeNode(rootNode);

	// All done, close new mbox file containing mails from all mbox files
	m_rootMboxCfile.Close();

	lastPath.Append(L"\\");
	m_globalFolderInfoDB.RemoveRootFolder(lastPath);  // FIXMEFIXME

	//CString temp_m_rootMboxFilePath = m_rootMboxFilePath + CString(L".mbox");

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	NListView* pListView = pFrame->GetListView();
	if (!NTreeView::ImboxviewFile(m_rootMboxFilePath))
	{
		_ASSERTE(FALSE);
		int deb = 1;
	}

	int paneId = 0;
	CString sText;
	CString fname;
	FileUtils::CPathStripPath(m_rootMboxFilePath, fname);
	int retval;

	CString fmt = L"Parsing intermediate merge raw mbox file \"%s\" ...";
	ResHelper::TranslateString(fmt);
	sText.Format(fmt, fname);

	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	MboxMail::SetLastPath(lastPath);
	retval = pListView->FillCtrl_ParseMbox(m_rootMboxFilePath);

	if (pFrame->m_labelAssignmentStyle != 0) // FIXMEFIXME why
	{
		int cnt = MboxMail::LinkDuplicateMails(MboxMail::s_mails);  // put all equal mails on seprate collision lists of hash table
	}

	CString txt = L"All mail files are merged.\n\n"
		L"Next, You will be asked to enter the target folder for the merged file.\n"
		L"Be aware that MBox Viewer will open the merged file automatically.\n"
		L"Make sure you select proper target folder.\n";
	ResHelper::TranslateString(txt);

	int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);

	CString saveFileName = fileName + L".mbox";
	CString title = L"Enter Name for Merged Archive File";
	CString  fileNameFilter = L"Mail Files (*.mbox)|*.mbox||";
	CString dfltExtention = L".mbox";
#if 0
	CString inFolderPath;
	CString lPath = lastPath;
	lPath.TrimRight(L"\\");
	FileUtils::GetFolderPath(lPath, inFolderPath);
#else
	// FIXME save dataFolder in CMainFrame and avoid redoing the same code all over !!!!! 
	CString section_general = CString(sz_Software_mboxview) + L"\\General";
	CString inFolderPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"dataFolder");
	inFolderPath += L"MergeFolder";
	BOOL retCreate = FileUtils::CreateDir((LPCWSTR)inFolderPath);
#endif
	CString outFolderPath;

	retval = SaveMergedFileDialog(saveFileName, fileNameFilter, dfltExtention, inFolderPath, outFolderPath, title);
	if (retval == FALSE)
		return -1;

	if (outFolderPath.IsEmpty())
		return -1;

	CString filePath = outFolderPath + saveFileName;

	MboxMail::SetLastPath(outFolderPath);
	CString lDataPath = MboxMail::GetLastDataPath();
	//if (FileUtils::PathFileExist(filePath))
	DeleteMBoxAllWorkFolders(filePath);


	// TODO: this should be moved to ArchiveMailsRemoveDuplicates to avoid confusion
	//
	if (!m_rootMboxCfile.Open(filePath, CFile::modeWrite | CFile::modeCreate | CFile::shareDenyNone, &exMergeTo))
	{
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exMergeTo);

#if 0
		CString txt = L"Could not create Merge To File \"" + filePath;
		txt += L"\" file.\n";
		txt += exErrorStr;
#endif

		CString fmt = L"Could not create Merge To File \"%s\" file.\n%s";
		ResHelper::TranslateString(fmt);
		CString txt;
		txt.Format(fmt, filePath, exErrorStr);

		//TRACE(L"%s\n", txt);

		CFileStatus rStatus;
		BOOL ret = m_rootMboxCfile.GetStatus(rStatus);;

		HWND h = NULL; // we don't have any window yet
		int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);

		return -1;
	}

	FileUtils::CPathStripPath(filePath, fname);
	fmt = "Removing duplicate mails in the final file %s ...";
	ResHelper::TranslateString(fmt);
	sText.Format(fmt, fname);

	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	if (pFrame->m_labelAssignmentStyle != 0)
		MergeMailsRemoveDuplicates();
	else
		ArchiveMailsRemoveDuplicates(m_rootMboxCfile, filePath);

	MboxMail::Destroy(&MboxMail::s_mails);

	if (m_rootMboxCfile.m_hFile != CFile::hFileNull )
		m_rootMboxCfile.Close();

	FileUtils::DelFile(m_rootMboxFilePath);

	MboxMail::SetLastPath(outFolderPath);

	HTREEITEM hFolderItem = FindFolder(outFolderPath);

	if (!hFolderItem)
	{
		BOOL expand = TRUE;
		DoOpen(outFolderPath, expand);
	}
	else
	{
		RefreshFolder(hFolderItem);
	}

	hFolderItem = FindFolder(outFolderPath);

	MailSelectionInfo mailSelection;
	mailSelection.m_lastPath = outFolderPath;
	mailSelection.m_lastMailFilePath = filePath;

	OpenLastSelection(&mailSelection);

	CString fName;
	FileUtils::GetFileName(filePath, fName);

	HTREEITEM hItem = FindItem(hFolderItem, fName);

	if (hItem)
	{
		TRACE(L"==============> \n    RefreshLabelsForSingleMailFile is called\n");
		int retval = RefreshLabelsForSingleMailFile(hItem);
		ExpandItem(hItem, TVE_EXPAND);
	}
	else
	{
		TRACE(L"==============> \n  RefreshLabelsForSingleMailFile not called\n");
		int deb = 1;
	}

#ifdef _DEBUG
	BOOL recursive = TRUE;
	int depth = 0;
	TRACE(L"Tree Control:\n");
	PrintTreeCtrl(hItem, recursive, depth);
#endif

	return 1;
}

long long milliseconds_now()
{
	static LARGE_INTEGER s_frequency;
	static BOOL s_use_qpc = QueryPerformanceFrequency(&s_frequency);
	if (s_use_qpc) {
		LARGE_INTEGER now;
		QueryPerformanceCounter(&now);
		return (1000LL * now.QuadPart) / s_frequency.QuadPart;
	}
	else {
		return GetTickCount64();  // return in mseconds
	}
}

class MergeFileAttr
{
public:
	MergeFileAttr() { m_isDir = FALSE; m_isMbox = FALSE; }
	CString m_path;
	BOOL m_isDir;
	BOOL m_isMbox;
};

int NTreeView::MergeRootSubFolder_NoLabels(CString& relativeFolderPath, CString& path, BOOL addToRecentFileList, BOOL expand)
{
	CString fileName;
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	NListView* pListView = pFrame->GetListView();

	path.TrimRight(L"\\");
	path.Append(L"\\");

	FileUtils::SplitFilePath(path, driveName, directory, fileNameBase, fileNameExtention);
	if (directory.GetLength() <= 1)
	{
		CString txt = L"The mbox files must be installed under a named folder.\n"
			L"Please create folder, move the mbox files to that folder and try again.";
		ResHelper::TranslateString(txt);
		int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
		return -1;
	}

	path.TrimRight(L"\\");

	int paneId = 0;
	CString sText;
	CString folderName;
	FileUtils::CPathStripPath(path, folderName);

	CString fmt = L"Merging mbox files in folder %s ...";
	ResHelper::TranslateString(fmt);
	sText.Format(fmt, (LPCWSTR)folderName);

	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	path.Append(L"\\");
	MboxMail::SetLastPath(path);
	//
	CString mboxFilePath = path;

	// build a string with wildcards
	CString strWildcard(mboxFilePath);
	strWildcard += L"\\*.*";

	// start working on files
	CFileFind finder;
	BOOL bWorking = finder.FindFile(strWildcard);

	CString fPath;
	CString fName;

	MergeFileAttr item;
	int i = 0;
	while (bWorking)
	{
		bWorking = finder.FindNextFile();

		// skip . and .. files; otherwise, we'd
		// recur infinitely!

		if (finder.IsDots())
			continue;

		if (finder.IsDirectory())
			continue;

		fPath = finder.GetFilePath();
		fName = finder.GetFileName();

		if (fPath.Compare(m_rootMboxFilePath) == 0)
			continue;

		if (!NTreeView::ImboxviewFile(fPath))
			continue;

		TRACE(L"FilePath=%s\n", fPath);

		_int64 fileSize = FileUtils::FileSize(fPath);

		CString fmt = L"Merging %s (%lld bytes) ...";
		ResHelper::TranslateString(fmt);
		sText.Format(fmt, (LPCWSTR)fName, fileSize);

		TRACE(L"%s\n", sText);
		if (pFrame)
			pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

		int ret = CMainFrame::MergeMboxArchiveFile(m_rootMboxCfile, fPath, i == 0);
		i++;
	}
	finder.Close();

	TRACE(L"");
	TRACE(L"MergeRootSubFolder: done\n");
	//m_globalFolderInfoDB.Print();

	return 1;
}



int NTreeView::MergeRootSubFolder(CString& relativeFolderPath, CString& path, BOOL addToRecentFileList, BOOL expand)
{
	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);

	if (pFrame && pFrame->m_labelAssignmentStyle == 0)  // no Labels
	{
		int retval = MergeRootSubFolder_NoLabels(relativeFolderPath, path, addToRecentFileList, expand);
		return retval;
	}

	CString fileName;
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;


	NListView* pListView = pFrame->GetListView();

	path.TrimRight(L"\\");
	path.Append(L"\\");

	FileUtils::SplitFilePath(path, driveName, directory, fileNameBase, fileNameExtention);
	if (directory.GetLength() <= 1)
	{
		CString txt = L"The mbox files must be installed under a named folder.\n"
			L"Please create folder, move the mbox files to that folder and try again.";
		ResHelper::TranslateString(txt);
		int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
		return -1;
	}

	path.TrimRight(L"\\");

	int paneId = 0;
	CString sText;
	CString folderName;
	FileUtils::CPathStripPath(path, folderName);

	CString fmt = L"Merging mbox files in folder %s ...";
	ResHelper::TranslateString(fmt);
	sText.Format(fmt, (LPCWSTR)folderName);

	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	path.Append(L"\\");
	MboxMail::SetLastPath(path);
	//
	CString mboxFilePath = path;

	// build a string with wildcards
	CString strWildcard(mboxFilePath);
	strWildcard += L"\\*.*";

	// start working on files
	CFileFind finder;
	BOOL bWorking = finder.FindFile(strWildcard);

	CString fPath;
	CString fName;

	std::vector <MergeFileAttr> mboxFileList;
	mboxFileList.reserve(8);

	MergeFileAttr item;
	int jj = 0;
	while (bWorking)
	{
		jj++;
		bWorking = finder.FindNextFile();

		// skip . and .. files; otherwise, we'd
		// recur infinitely!

		if (finder.IsDots())
			continue;

		fPath = finder.GetFilePath();
		fName = finder.GetFileName();

		if (fPath.Compare(m_rootMboxFilePath) == 0)
			continue;

		item.m_path = fPath;
		item.m_isDir = finder.IsDirectory();
		item.m_isMbox = FALSE;

		if (item.m_isDir)
		{
			mboxFileList.push_back(item);
		}
		else
		{
			if (NTreeView::ImboxviewFile(fPath))
			{
				item.m_isMbox = TRUE;
				mboxFileList.push_back(item);

				if ((jj % 1000) == 0)
				{
					CString fmt = L"Discovered mbox file %s ...";
					ResHelper::TranslateString(fmt);
					sText.Format(fmt, (LPCWSTR)fName);

					if (pFrame)
						pFrame->SetStatusBarPaneText(paneId, sText, TRUE);
				}
			}
		}
	}
	finder.Close();

	MergeFileAttr* ar = mboxFileList.data();
	MergeFileAttr* el;

	int j;

	MessageBeep(MB_OK);

	sText.Format(L"Discovery of mbox files completed ...");
	ResHelper::TranslateString(sText);
	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	// Check for duplication of base file and folder names
	// Keep it simple for now. Stop if any collision is found
	// This is a source o problems. Should we just rely on the full name instead of base name

	CString baseName1;
	CString baseName2;
	MergeFileAttr* el1;
	MergeFileAttr* el2;
	BOOL foundCollision = FALSE;
	if (mboxFileList.size() > 128)
	{
		foundCollision = TRUE;
	}
	// very inefficient for large number of files
	else for (j = 0; j < mboxFileList.size(); j++)
	{
		el1 = &ar[j];

		if ((el1->m_isDir == FALSE) && (el1->m_isMbox == FALSE))
			continue;

		if (el1->m_isDir)
			baseName1 = el1->m_path;
		else
			FileUtils::GetFileBaseName(el1->m_path, baseName1);

		int k;
		for (k = j + 1; k < mboxFileList.size(); k++)
		{
			el2 = &ar[k];

			if ((el2->m_isDir == FALSE) && (el2->m_isMbox == FALSE))
				continue;

			if (el2->m_isDir)
				baseName2 = el2->m_path;
			else
				FileUtils::GetFileBaseName(el2->m_path, baseName2);

			if (baseName1.Compare(baseName2) == 0)
			{
				foundCollision = TRUE;
				break;
			}
		}

		if (foundCollision)
			break;
	}

	MessageBeep(MB_OK);
	sText.Format(L"Detection of name collisions is completed ...");
	ResHelper::TranslateString(sText);
	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	for (j = 0; j < mboxFileList.size(); j++)
	{
		el = &mboxFileList[j];
		TRACE(L"MergeRootSubFolder: found mbox file=%s\n", el->m_path);
	}

	MessageBeep(MB_OK);
	sText.Format(L"Merging mbox files ...");
	ResHelper::TranslateString(sText);
	if (pFrame)
		pFrame->SetStatusBarPaneText(paneId, sText, TRUE);

	CString labelName;
	for (j = 0; j < mboxFileList.size(); j++)
	{
#ifdef _DEBUG
		long long msecStart = milliseconds_now();
#endif
		el = &ar[j];

		if (el->m_isDir)
			continue;

		if (!el->m_isMbox)
			continue;

		fPath = el->m_path;
		TRACE(L"MergeRootSubFolder: merging mbox file=%s\n", fPath);

		int retval = pListView->FillCtrl_ParseMbox(fPath);

		FileUtils::GetFileName(fPath, fName);
		CStringA labels = "X-Gmail-Labels: \"";
		CString label;

		if (pFrame->m_labelAssignmentStyle == 2)  //per folder
		{
			if (relativeFolderPath.IsEmpty())
			{
				label = folderName;
			}
			else
			{
				label += relativeFolderPath;
				//label == relativeFolderPath + "/" + folderName;
			}
		}
		else
		{
			if (foundCollision)
			{
				labelName = fName;
			}
			else
			{
				FileUtils::GetFileBaseName(fName, labelName);
			}

			if (relativeFolderPath.IsEmpty())
			{
				label = labelName;
			}
			else
			{
				label = relativeFolderPath + L"/" + labelName;
			}
		}

		DWORD error;
		CStringA Utf8Label;
		TextUtilsEx::WStr2UTF8(&label, &Utf8Label, error);

		CStringA wordEncodedLabel;
		int encodeType = 'Q';   // 'Q' == quoted (best for asci) 'B' ==  base64 
		TextUtilsEx::WordEncode(Utf8Label, wordEncodedLabel, encodeType);

		labels += wordEncodedLabel + "\"\r\n";

		MergeMails(labels);

#ifdef _DEBUG
		long long msecEnd = milliseconds_now();
		DWORD msecElapsed = (DWORD)(msecEnd - msecStart);
		if (msecElapsed < 20)
		{
			; // Sleep(20);
		}
		TRACE(L"MergeRootSubFolder:: msecElapsed=%d\n", msecElapsed);
#endif
	}

	int deb = 1;

	TRACE(L"");
	TRACE(L"MergeRootSubFolder: done\n");
	//m_globalFolderInfoDB.Print();

	return 1;
}

int NTreeView::MergeMails(CStringA& label)
{
	static const char* cFromMailBegin = "From ";
	static const int cFromMailBeginLen = istrlen(cFromMailBegin);

	MboxMail* m;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (!pFrame)
		return -1;

	SimpleString* bdy = MboxMail::m_workbuf;

	int mailCount = MboxMail::s_mails.GetCount();

	int i = 0;
	char* p;
	char* e;
	char* p_newline;
	char* p_beg;
	int len;

	CStringA FromLine = "From 1513218656940664977@xxx ";

	BOOL ret;
	for (i = 0; i < mailCount; i++)
	{
		m = MboxMail::s_mails[i];

		ret = m->GetBodySS(bdy);
		p = bdy->Data();
		p_beg = p;
		e = p + bdy->Count();

		p = MimeParser::SkipEmptyLines(p, e);

		char* ch_end_line = "\r\n";
		int ch_end_line_len = 2;

		if (TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) != 0)
		{
			// From 1513218656940664977@xxx Thu Sep 24 18:02:48 +0000 2015
			// More parsing would be needed to find first line with Date and Time
			// The below is safe for parsing the created mbox file

			//FromLine = "From 1513218656940664977@xxx\r\n";

			p_newline = MimeParser::EatNewLine(p, e);
			char ch_last = *(p_newline - 1);
			char ch_before_last = *(p_newline - 2);
			if (ch_last == '\n')
			{
				if (ch_before_last != '\r')
				{
					ch_end_line = "\n";
					ch_end_line_len = 1;
				}
			}

			m_rootMboxCfile.Write(FromLine, FromLine.GetLength());
			m_rootMboxCfile.Write(ch_end_line, ch_end_line_len);
			m_rootMboxCfile.Write(label, label.GetLength());

			m_rootMboxCfile.Write(p_beg, UIntPtr2UInt(e - p_beg));
		}
		else
		{
			p_newline = MimeParser::EatNewLine(p, e);
			len = IntPtr2Int(p_newline - p_beg);

			m_rootMboxCfile.Write(p_beg, len);
			m_rootMboxCfile.Write(label, label.GetLength());

			m_rootMboxCfile.Write(p_newline, UIntPtr2UInt(e - p_newline));
		}
		//m_rootMboxCfile.Write("\r\n", 2);  // this should be fine too
		m_rootMboxCfile.Write(ch_end_line, ch_end_line_len);
	}

	int deb = 1;
	return 1;
}

//  FIXME defined as int ArchiveMailsRemoveDuplicates(CFile& fpMerged);
//  Avoid globals such as m_rootMboxCfile
int NTreeView::MergeMailsRemoveDuplicates()   // remove duplicate mails and merge labels
{
	static const char* cLabels = "x-gmail-labels";
	static const int cLabelsLen = istrlen(cLabels);

	static const char* cFromMailBegin = "From ";
	static const int cFromMailBeginLen = istrlen(cFromMailBegin);

	MboxMail* m;
	MboxMail* mc;

	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (!pFrame)
		return -1;

	SimpleString* bdy = MboxMail::m_workbuf;

	int mailCount = MboxMail::s_mails.GetCount();

	CStringA val;
	CStringA line;
	BOOL foundLabels;

	int i = 0;
	char* p;
	char* e;
	char* p_newline;
	char* p_beg;
	int len;

	MailArray* collArray = &MboxMail::s_mails_find;  // caller must assure
	BOOL ret;
	for (i = 0; i < mailCount; i++)
	{
		m = MboxMail::s_mails[i];

		if (m->m_duplicateId)
			continue;

		MboxMail::MboxMailMapCollionList* list = MboxMail::getCollisionList(m);  // uses MboxMail m_pMboxMailMap

		if (list->count() == 1)  // no need to merge labels
		{
			mc = list->first();
			_ASSERT(mc == m);

			ret = m->GetBodySS(bdy);
			p = bdy->Data();
			len = bdy->Count();

			m_rootMboxCfile.Write(p, len);
			m_rootMboxCfile.Write("\r\n", 2);
			continue;
		}

		// Find all equal mails and add to collArray
		collArray->SetCountKeepData(list->count());
		int cnt = 0;
		for (mc = list->first(); mc != 0; mc = list->next(mc))
		{
			bool ret = MboxMail::equal(m, mc);
			if (ret == false)
				continue;
			collArray->SetAt(cnt++, mc);
		}
		collArray->SetCountKeepData(cnt);

		_ASSERT(collArray->GetCount() > 0);
		if (collArray->GetCount() == 1)
		{
			mc = collArray->GetAt(0);
			_ASSERTE(mc == m);
			ret = m->GetBodySS(bdy);
			p = bdy->Data();
			len = bdy->Count();

			m_rootMboxCfile.Write(p, len);
			m_rootMboxCfile.Write("\r\n", 2);
			continue;
		}

		// collArray->count() > 1. Current m can be at any position of collArray
		// a array will contains mails with unique label strings. 
		// This is simplistic approach otherwise decode strings, split and remove duplicate labels
		CStringArrayA a;
		CStringA el;
		int k;
		for (k = 0; k < collArray->GetCount(); k++)
		{
			mc = collArray->GetAt(k);

			if (m == mc)
				continue;

			// Read first 1024 characters. x-gmail-labels from gmail is at very beginning
			ret = mc->GetBodySS(bdy, 1024);  // TODO: make small, enlarge and retry is needed ???
			p = bdy->Data();
			e = p + bdy->Count();

			p = MimeParser::SkipEmptyLines(p, e);

			val.Empty();
			foundLabels = FALSE;
			while (p < e)
			{
				if (TextUtilsEx::strncmpUpper2Lower(p, e, cLabels, cLabelsLen) == 0)
				{
					foundLabels = TRUE;
					p = MimeParser::GetMultiLine(p, e, line, TRUE);
					val = line.Mid(cLabelsLen);
					val.TrimLeft(": ");
					val.TrimRight(" ");
					//
					//TRACE(L"Labels=%s\n", val);

					int j;
					for (j = 0; j < a.GetSize(); j++)
					{
						el = a.ElementAt(j);
						if (el.Compare(val) == 0)
						{
							break;
						}
					}
					if (j == a.GetSize())
						a.Add(val);
					break;
				}
				else
					p = MimeParser::EatNewLine(p, e);
			}
		}

		CStringA extraLabel;
		int ii;
		// a array should contain unique label strings. The label from current m is exluded
		for (ii = 0; ii < a.GetSize(); ii++)
		{
			extraLabel.Append("  ,");
			extraLabel.Append(a.ElementAt(ii));
			extraLabel.Append("\r\n");
		}

		mc = m;

		ret = mc->GetBodySS(bdy);
		p = bdy->Data();
		e = p + bdy->Count();
		p_beg = p;

		p = MimeParser::SkipEmptyLines(p, e);

		while (p < e)
		{
			if (TextUtilsEx::strncmpUpper2Lower(p, e, cLabels, cLabelsLen) == 0)
			{
				foundLabels = TRUE;
				p_newline = MimeParser::GetMultiLine(p, e, line);

				len = IntPtr2Int(p_newline - p_beg);
				m_rootMboxCfile.Write(p_beg, len);
				m_rootMboxCfile.Write((char*)(LPCSTR)extraLabel, extraLabel.GetLength());
				m_rootMboxCfile.Write(p_newline, UIntPtr2UInt(e - p_newline));
				break;
			}
			else
			{
				p = MimeParser::EatNewLine(p, e);
			}
		}
	}
	return 1;
}

int NTreeView::ArchiveMailsRemoveDuplicates(CFile& fp, CString& filePath)
{
	MailArray& s_mails_array = MboxMail::s_mails_all;
	BOOL putDuplicatesOnFindArray = FALSE;
	int dupCnt = MboxMail::RemoveDuplicateMails(s_mails_array, putDuplicatesOnFindArray);
	if (dupCnt == 0)
	{
#if 1
		// TODO: MoveFileEx fails because fp is open
		// Need to close fp or use WCopyFile
		// MoveFileEx avoids copy file if source and destinations are on the the volume
		// More efficient if merge file is very large
		if (m_rootMboxCfile.m_hFile != CFile::hFileNull)  // use of m_hFile not recommended in some cases according to windows doc
			fp.Close();

		DWORD nFlags = MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH | MOVEFILE_COPY_ALLOWED;
		BOOL retMove = MoveFileEx(m_rootMboxFilePath, filePath, nFlags);
		if (retMove == FALSE)
		{
			CString errorText = FileUtils::GetLastErrorAsString();
			TRACE(L"ArchiveMailsRemoveDuplicates: MoveFileEx failed \"%s\"\n", errorText);
#if 0
			CString txt = L"Failed to move file \"" + m_rootMboxFilePath;
			txt += L"\" file.\n\n";
			txt.Append(errorText);
#endif

			CString fmt = L"Failed to move file \"%s\" file.\n\n%s";
			ResHelper::TranslateString(fmt);
			CString txt;
			txt.Format(fmt, m_rootMboxFilePath, errorText);

			HWND h = NULL; // we don't have any window yet
			int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
#else
		// Works, no surprises but less effcient for large merged files
		CString errorText;
		BOOL bFailIfExists = FALSE;
		BOOL retCopy = FileUtils::WCopyFile(m_rootMboxFilePath, filePath, bFailIfExists, errorText);
		if (retCopy == FALSE)
		{
			TRACE(L"ArchiveMailsRemoveDuplicates: WCopyFile failed \"%s\"\n", errorText);

			CString txt = L"Failed to copy file \"" + m_rootMboxFilePath;
			txt += L"\" file.\n\n";
			txt.Append(errorText);

			HWND h = NULL; // we don't have any window yet
			int answer = MessageBox(txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
			return -1;
		}
#endif
	}
	else
	{
		SimpleString* outbuf = MboxMail::m_outbuf;
		outbuf->ClearAndResize(10000);

		BOOL ret;
		MboxMail* m;
		for (int i = 0; i < s_mails_array.GetSize(); i++)
		{
			m = s_mails_array[i];
			outbuf->Clear();
			ret = m->GetBodySS(outbuf);
			// TODO: may need to validate last mail, i.e check for From presence
			fp.Write(outbuf->Data(), outbuf->Count());
		}
	}
	return 1;
}


BOOL NTreeView::SaveMergedFileDialog(CString& fileName, CString& fileNameFilter, CString& dfltExtention, CString& inFolderPath, CString& outFolderPath, CString& title)
{
	CString lastPath = MboxMail::GetLastPath();
	CString datapath = MboxMail::GetLastDataPath();

	CString fileBaseName;
	FileUtils::GetFileBaseName(fileName, fileBaseName);

	for (;;)
	{
		DWORD dwFlags = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
		dwFlags &= ~OFN_NOCHANGEDIR;
		CFileDialog dlgFile(FALSE, dfltExtention, fileBaseName, dwFlags, fileNameFilter);

		OPENFILENAME& ofn = dlgFile.GetOFN();
		//ofn.Flags |= ??;
		ofn.lpstrInitialDir = inFolderPath;
		ofn.lpstrTitle = title;

		INT_PTR ret = dlgFile.DoModal();
		if (ret == IDOK)
		{
			fileName = dlgFile.GetFileName();
			CString fileFilePath = dlgFile.GetPathName();
			BOOL ret = FileUtils::CPathGetPath(fileFilePath, outFolderPath);

			TRACE(L"FOLDER=%s\n", outFolderPath);
			TRACE(L"FILE=%s\n", fileName);

			outFolderPath.TrimRight(L"\\");
			outFolderPath.Append(L"\\");

			CString driveName;
			CString directory;
			CString fileNameBase;
			CString fileNameExtention;

			FileUtils::SplitFilePath(outFolderPath, driveName, directory, fileNameBase, fileNameExtention);
			if (directory.GetLength() <= 1)
			{
				CString txt = L"The mbox files must be installed under a named folder.\n\n"
					L"Please select different folder or create new folder while in Save File Dialog and try again.";
				ResHelper::TranslateString(txt);
				int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
				continue;
			}

			if (outFolderPath.Find(L"UMBoxViewer") >= 0)
			{
				CString txt = L"The \"UMBoxViewer\" name is reserved and it can't appear in the file path.\n"
					L"Please create different name and try again.";
				ResHelper::TranslateString(txt);
				int answer = MessageBox(txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
				continue;
			}

			if (outFolderPath.Compare(lastPath) == 0)
			{
				CString txt = L"Selecting target folder to source root folder for merged file is not recommended or allowed.\n\nPlease try again.";
				ResHelper::TranslateString(txt);
				HWND h = GetSafeHwnd(); // we don't have any window yet
				int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
				continue;
			}

			HTREEITEM hItem = IsFolderOpen(outFolderPath);
			if (hItem)
			{
				BOOL ret = IsFolderOpenAsRootFolder(outFolderPath);
				if (ret)
				{
					CString txt = L"Selected target folder is already open as the root folder.\n\nPlease try again.";
					ResHelper::TranslateString(txt);
					HWND h = GetSafeHwnd(); // we don't have any window yet
					int answer = ::MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
					continue;
				}
			}
			return TRUE;
		}
		break;
	}
	outFolderPath.Empty();
	return FALSE;
}

void NTreeView::SetupCacheFolderList(CArray<CString>& cacheFolderList)
{
	cacheFolderList.Add(L"PrintCache");
	cacheFolderList.Add(L"ImageCache");
	cacheFolderList.Add(L"AttachmentCache");
	cacheFolderList.Add(L"EmlCache");
	cacheFolderList.Add(L"ArchiveCache");
	cacheFolderList.Add(L"ListCache");
	cacheFolderList.Add(L"MergeCache");
	cacheFolderList.Add(L"LabelCache");  // keep as the last
}

LRESULT NTreeView::OnCmdParam_OnSwitchWindow(WPARAM wParam, LPARAM lParam)
{
	CWheelTreeCtrl	*tree = &m_tree;
	CMainFrame::SetWindowFocus(tree);
	return 0;
}
