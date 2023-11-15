//
//////////////////////////////////////////////////////////////////
//
//  Windows Mbox Viewer is a free tool to view, search and print mbox mail archives..
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
#include "TextUtilsEx.h"
#include "CAttachments.h"
#include "mboxview.h"
#include "NMsgView.h"
#include "PictureCtrl.h"
#include "CPictureCtrlDemoDlg.h"

BEGIN_MESSAGE_MAP(CAttachments, CListCtrl)
	//ON_WM_PAINT()
	ON_WM_ERASEBKGND()
	// OWNERDRAW didn't work but PAINT seem to work
	//ON_NOTIFY(NM_CUSTOMDRAW, IDC_ATTACHMENTS, OnCustomDraw)
	ON_NOTIFY_REFLECT(NM_CLICK, OnActivating)
	ON_NOTIFY_REFLECT(NM_DBLCLK, OnDoubleClick)
	ON_NOTIFY_REFLECT(NM_RCLICK, OnRClick)  // Right Click Menu

END_MESSAGE_MAP()


CAttachments::CAttachments(NMsgView *pMsgView)
{
	m_pMsgView = pMsgView;
}


CAttachments::~CAttachments()
{
	int deb = 1;
}

void CAttachments::OnActivating(NMHDR* pNMHDR, LRESULT* pResult)
{
	NMITEMACTIVATE *pnm = (NMITEMACTIVATE *)pNMHDR;

	*pResult = 0;

	int nItem = pnm->iItem;
	if (nItem < 0)
		return;

	// maybe we can do here something usefull in the future -:)
	CString attachmentName = GetItemText(nItem, LVIR_BOUNDS);
}

int CAttachments::FindAttachmentByNameA(CStringA &name)
{
	CStringW nameW;
	DWORD error;
	BOOL ret = TextUtilsEx::Ansi2WStr(name, nameW, error);
	int result = FindAttachmentByNameW(nameW);
	return result;
}

int CAttachments::FindAttachmentByNameW(CStringW &name)
{
	int iCnt = (int)m_attachmentTbl.GetCount();
	for (int ii = 0; ii < iCnt; ii++)
	{
		CStringW attachmentFileName = m_attachmentTbl[ii]->m_nameW;
		if (attachmentFileName.Compare(name) == 0)
			return ii;
	}
	return  -1;
}

void CAttachments::Reset()
{
	if (GetItemCount() != m_attachmentTbl.GetCount()) 
		int deb = 1;

	DeleteAllItems();

	for (int i = 0; i < m_attachmentTbl.GetCount(); i++)
	{
		delete m_attachmentTbl[i];
	}
	m_attachmentTbl.RemoveAll();
}

void CAttachments::ReleaseResources()
{
	// Call to GetItemCount() asserts in debug
	//if (GetItemCount() != m_attachmentTbl.GetCount()) int deb = 1;

	// Call to DeleteAllItems() asserts in debug
	// No memory leaks are not reported when commented out !!!
	//DeleteAllItems();

	for (int i = 0; i <  m_attachmentTbl.GetCount(); i++)
	{
		delete m_attachmentTbl[i];
	}
	m_attachmentTbl.RemoveAll();
	return;
}

int __cdecl AttachmentPred(void const * first, void const * second)
{
	AttachmentInfo *f = *((AttachmentInfo**)first);
	AttachmentInfo *s = *((AttachmentInfo**)second);
	int ret = _wcsicmp(f->m_nameW.operator LPCWSTR(), s->m_nameW.operator LPCWSTR());
	return ret;
}

void CAttachments::Complete()
{
	std::qsort(m_attachmentTbl.GetData(), m_attachmentTbl.GetCount(), sizeof(void*), AttachmentPred);

	//m_attachmentTbl.RemoveAll();

	CString path = CMainFrame::GetMboxviewTempPath();
	for (int i = 0; i < m_attachmentTbl.GetCount(); i++)
	{
		CStringW nameW = m_attachmentTbl[i]->m_nameW;
		CStringW cStrNamePath = path + nameW;

		if (!FileUtils::PathFileExist(cStrNamePath))
			int deb = 1;

		int iIcon = 0;
		SHFILEINFOW shFinfo;
		if (!SHGetFileInfoW(cStrNamePath,
			0,
			&shFinfo,
			sizeof(shFinfo),
			SHGFI_ICON |
			SHGFI_SMALLICON))
		{
			TRACE(L"Error Gettting SystemFileInfo for \"%s\" file!\n", cStrNamePath);
		}
		else {
			iIcon = shFinfo.iIcon;
			// we only need the index from the system image list
			DestroyIcon(shFinfo.hIcon);
		}

		int inserRet = CListCtrl::InsertItem(GetItemCount(), nameW, iIcon);
		int deb = 1;
	}
	return;
}

BOOL CAttachments::InsertItemW(CStringW &cStrName, int id, CMimeBody* pBP)
{
	int index = CAttachments::FindAttachmentByNameW(cStrName);
	// Caller must assure names are unique
	_ASSERTE(index < 0);

	AttachmentInfo *item = new AttachmentInfo;
	item->m_nameW = cStrName;
	m_attachmentTbl.Add(item);

	return TRUE;
}

void CAttachments::OnRClick(NMHDR* pNMHDR, LRESULT* pResult)
{
	const wchar_t *Open = L"Open";
	const wchar_t *Print = L"Print";
	const wchar_t *OpenFileLocation = L"Open File Location";

	LPNMITEMACTIVATE pnm = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);

	*pResult = 0;

	int nItem = pnm->iItem;
	if (nItem < 0)
		return;

	CStringW attachmentName = m_attachmentTbl[nItem]->m_nameW;

	TRACE(L"Selecting %d\n", pnm->iItem);

	// TODO: Add your control notification handler code here
#define MaxShellExecuteErrorCode 32

	CPoint pt;
	::GetCursorPos(&pt);
	CWnd *wnd = WindowFromPoint(pt);

	CMenu menu;
	menu.CreatePopupMenu();
	menu.AppendMenu(MF_SEPARATOR);

	const UINT M_OPEN_Id = 1;
	MyAppendMenu(&menu, M_OPEN_Id, Open);

	const UINT M_PRINT_Id = 2;
	MyAppendMenu(&menu, M_PRINT_Id, Print);

	const UINT M_OpenFileLocation_Id = 3;
	MyAppendMenu(&menu, M_OpenFileLocation_Id, OpenFileLocation);

	CBitmap  printMap;
	int retval = AttachIcon(&menu, Print, IDB_PRINT, printMap);
	CBitmap  foldertMap;
	retval = AttachIcon(&menu, OpenFileLocation, IDB_FOLDER, foldertMap);

	int command = menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, this);

	UINT nFlags = TPM_RETURNCMD;
	CString menuString;
	int chrCnt = menu.GetMenuString(command, menuString, nFlags);

	bool forceOpen = false;
	HINSTANCE result = (HINSTANCE)(MaxShellExecuteErrorCode + 1);  // OK

	CString path = CMainFrame::GetMboxviewTempPath();
	if (command == M_PRINT_Id)
	{
		CStringW filePath = path + attachmentName;
		result = ShellExecuteW(NULL, L"print", attachmentName, NULL, path, SW_SHOW);
		if ((UINT_PTR)result <= MaxShellExecuteErrorCode)
		{
			CString errorText;
			ShellExecuteError2Text((UINT_PTR)result, errorText);
			errorText += L".\nOk to try to open this file ?";
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, errorText, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_YESNO);
			if (answer == IDYES)
				forceOpen = true;
		}
		int deb = 1;
	}
	if ((command == M_OPEN_Id) || forceOpen)
	{
		CStringW filePath = path + attachmentName;

		DWORD binaryType = 0;
		BOOL isExe = GetBinaryTypeW(filePath, &binaryType);

		HWND h = GetSafeHwnd();
		result = ShellExecuteW(h, L"open", attachmentName, NULL, path, SW_SHOWNORMAL);
		CMainFrame::CheckShellExecuteResult(result, h, &filePath);

		int deb = 1;
	}
	else if (command == M_OpenFileLocation_Id)
	{
		CStringW filePath = path + attachmentName;

		if (FileUtils::BrowseToFile(filePath) == FALSE) {
			HWND h = GetSafeHwnd();
			HINSTANCE result = ShellExecuteW(h, L"open", path, NULL, NULL, SW_SHOWNORMAL);
			CMainFrame::CheckShellExecuteResult(result, h);
		}

		int deb = 1;
	}

	Invalidate();
	UpdateWindow();

	*pResult = 0;
}

void CAttachments::OnDoubleClick(NMHDR* pNMHDR, LRESULT* pResult)
{
	NMITEMACTIVATE *pnm = (NMITEMACTIVATE *)pNMHDR;

	*pResult = 0;

	int nItem = pnm->iItem;
	if (nItem < 0)
		return;

	CStringW attachmentNameW = m_attachmentTbl[nItem]->m_nameW;
	CStringW ext = PathFindExtensionW(attachmentNameW);

	CStringA extA;
	DWORD error;
	BOOL retW2A = TextUtilsEx::WStr2Ansi(ext, extA, error);

	bool isSupportedPictureFile = false;

	// Need to create table of extensions so this class and Picture viewer class can share the list
	if (CCPictureCtrlDemoDlg::isSupportedPictureFile(ext))
	{
		isSupportedPictureFile = true;
	}

	if ((nItem >= 0) && (m_pMsgView && m_pMsgView->m_bImageViewer) && isSupportedPictureFile)
	{
		CCPictureCtrlDemoDlg dlg(&attachmentNameW);
		INT_PTR nResponse = dlg.DoModal();
		if (nResponse == IDOK) {
			int deb = 1;
		}
		else if (nResponse == IDCANCEL)
		{
			int deb = 1;
		}
	}
	else
	{
		CStringW path = CMainFrame::GetMboxviewTempPath();
		CStringW filePath = path + attachmentNameW;
		// Windows Photos application doesn't show next/prev photo even with path specified
		// Build in Picture Viewer is set as deafault
		HWND h = GetSafeHwnd();
		HINSTANCE result = ShellExecuteW(h, L"open", attachmentNameW, NULL, path, SW_SHOWNORMAL);
		CMainFrame::CheckShellExecuteResult(result, h, &filePath);

	}

	*pResult = 0;
}

int CAttachments::PreTranslateMessage(MSG* pMsg)
{
	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
	{
		if ((pMsg->message & 0xffff) == WM_MOUSEMOVE)
		{
			int deb = 1;
		}
		else if ((pMsg->message & 0xffff) == WM_KEYDOWN)
		{
			if (pMsg->wParam == VK_ESCAPE)
			{
				AfxGetMainWnd()->PostMessage(WM_CLOSE);
			}
		}
	}

	return CWnd::PreTranslateMessage(pMsg);
}

BOOL CAttachments::OnEraseBkgnd(CDC* pDC)
{
	//BOOL ret = CWnd::OnEraseBkgnd(pDC);

	CRect rect;
	GetClientRect(&rect);

	DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessageAttachments);

	SetBkColor(color);
	BOOL ret = SetTextBkColor(color);
	pDC->FillRect(&rect, &CBrush(color));

	return TRUE;
}
