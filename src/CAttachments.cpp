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
	ON_WM_PAINT()
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

void CAttachments::OnPaint()
{
	CPaintDC dc(this); // device context for painting
					   // TODO: Add your message handler code here
					   // Do not call CListCtrl::OnPaint() for painting messages

	NListView *pListView = 0;
	NMsgView *pMsgView = 0;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		pListView = pFrame->GetListView();
		pMsgView = pFrame->GetMsgView();
	}

	HDC hDC = dc.GetSafeHdc();

	CRect cr;
	GetClientRect(cr);
	DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessageAttachments);
	if (pListView && !pListView->m_bApplyColorStyle)
		color = RGB(255, 255, 255);
	dc.FillRect(&cr, &CBrush(color));

	CFont newFont;
	newFont.CreatePointFont(85, _T("Tahoma"));

	// Set new font. Should reinstall old oldFont?? doesn't seem to matter
	CFont  *pOldFont = dc.SelectObject(&newFont);

	DWORD bkcolor = ::GetSysColor(COLOR_HIGHLIGHT);
	DWORD txcolor = ::GetSysColor(COLOR_HIGHLIGHTTEXT);

	CStringW mboxFileName;
	CStringW mboxFolderPath;
	CStringW mboxFilePath;

	CRect rect;
	int xpos;
	int ypos;
	int rectLen = 0;

	int iCnt = m_attachmentTbl.GetCount();
	for (int ii = 0; ii < iCnt; ii++)
	{
		GetItemRect(ii, &rect, LVIR_BOUNDS);
		if (rect.IsRectNull())
			continue;

		int w = rect.right - rect.left;

		UINT nMask = LVIS_SELECTED | LVIS_FOCUSED;

		UINT nState = GetItemState(ii, nMask);
		if (nState & LVIS_SELECTED)
		{
			dc.FillRect(&rect, &CBrush(bkcolor));
			dc.SetBkMode(TRANSPARENT);
			dc.SetTextColor(txcolor);
		}
		else
		{
			dc.FillRect(&rect, &CBrush(color));
			dc.SetBkMode(TRANSPARENT);
			dc.SetTextColor(RGB(0, 0, 0));
		}

		mboxFileName = m_attachmentTbl[ii]->m_nameW;
		mboxFolderPath = FileUtils::GetmboxviewTempPathW();
		mboxFilePath = mboxFolderPath + mboxFileName;

		int iIcon = 0;
		SHFILEINFOW shFinfo;
		if (!SHGetFileInfoW(mboxFilePath,
			0,
			&shFinfo,
			sizeof(shFinfo),
			SHGFI_ICON |
			SHGFI_SMALLICON))
		{
			if (FileUtils::PathFileExistW(mboxFilePath))
				int deb = 1;
			TRACE("Error Gettting SystemFileInfo!\n");
		}
		else 
		{
			xpos = rect.left + 1;
			ypos = rect.top + 3;

			int w = 14;
			int h = rect.Height() - 3;
			BOOL r = DrawIconEx(hDC, xpos, ypos, shFinfo.hIcon, w, h, 0, 0, DI_NORMAL);

			DestroyIcon(shFinfo.hIcon);
		}
		xpos = rect.left + 4;
		ypos = rect.top + 3;

		xpos += 14;

		BOOL retW = ::ExtTextOutW(hDC, xpos, ypos, ETO_CLIPPED, rect, (LPCWSTR)mboxFileName, mboxFileName.GetLength(), NULL);
	}
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

int CAttachments::FindAttachmentByName(CString &name)
{
	CStringW nameW;
	DWORD error;
	BOOL ret = TextUtilsEx::Ansi2Wide(name, nameW, error);
	int result = FindAttachmentByNameW(nameW);
	return result;
}

int CAttachments::FindAttachmentByNameW(CStringW &name)
{
	int iCnt = m_attachmentTbl.GetCount();
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
	int ret = wcscmp(f->m_nameW.operator LPCWSTR(), s->m_nameW.operator LPCWSTR());
	return ret;
}

void CAttachments::Complete()
{
	std::qsort(m_attachmentTbl.GetData(), m_attachmentTbl.GetCount(), sizeof(void*), AttachmentPred);

	for (int i = 0; i < m_attachmentTbl.GetCount(); i++)
	{
		CStringW nameW = m_attachmentTbl[i]->m_nameW;
		CStringW cStrNamePath = FileUtils::GetmboxviewTempPathW() + nameW;

		int iIcon = 0;
		SHFILEINFOW shFinfo;
		if (!SHGetFileInfoW(cStrNamePath,
			0,
			&shFinfo,
			sizeof(shFinfo),
			SHGFI_ICON |
			SHGFI_SMALLICON))
		{
			TRACE("Error Gettting SystemFileInfo!\n");
		}
		else {
			iIcon = shFinfo.iIcon;
			// we only need the index from the system image list
			DestroyIcon(shFinfo.hIcon);
		}

		CPaintDC dc(this);
		HDC hDC = dc.GetSafeHdc();

		SIZE size;
		int wlen = nameW.GetLength();
		BOOL ret = GetTextExtentPoint32W(hDC, nameW, wlen, &size);

		SIZE size1A;
		CString nameA = "MMM";
		BOOL retA = GetTextExtentPoint32(hDC, nameA, nameA.GetLength(), &size1A);

		int multiplier = size.cx / size1A.cx;

		SIZE size2A;
		for (int k = 0; k < 100; k++)
		{
			BOOL retA = GetTextExtentPoint32(hDC, nameA, nameA.GetLength(), &size2A);
			int fudge = size1A.cx;

			if (multiplier > 10)
				fudge = 3 * size1A.cx;
			else if (multiplier > 5)
				fudge = 2 * size1A.cx;

			if (size2A.cx <= (size.cx + fudge))
				nameA.Append("MMM");
			else
			{
				break;
			}
		}
		int inserRet = CListCtrl::InsertItem(GetItemCount(), nameA, iIcon);
		int deb = 1;
	}
	return;
}

BOOL CAttachments::InsertItemW(CStringW &cStrName, int id, CMimeBody* pBP)
{
	CStringW validNameW;

	BOOL bReplaceWhiteWithUnderscore = FALSE;
	FileUtils::MakeValidFileNameW(cStrName, validNameW, bReplaceWhiteWithUnderscore);

	// Check for duplicate names. Sometimes two or more names can represent diffrent content
	int pos = FindAttachmentByNameW(validNameW);
	if (pos >= 0)
	{
		CStringW fileExtension = ::PathFindExtensionW(validNameW);
		CStringW fileName = ::PathFindFileNameW(validNameW);

		int pos2 = fileName.ReverseFind('.');
		CStringW fileNameBase = fileName;
		if (pos2 >= 0) {
			fileNameBase = fileName.Mid(0, pos2);
		}
		validNameW.Format(L"%s%s%d%s", fileNameBase, L"_", id, fileExtension);
		int deb = 1;
	}
	else
	{
		int deb = 1;
	}

	CStringW cStrNamePathW = FileUtils::GetmboxviewTempPathW() + validNameW;
	CString cStrNamePathA;
	DWORD error;
	BOOL retW2A = TextUtilsEx::Wide2Ansi(cStrNamePathW, cStrNamePathA, error);

	const unsigned char* data = pBP->GetContent();
	int dataLength = pBP->GetContentLength();

	BOOL retWrite = FileUtils::Write2File(cStrNamePathW, data, dataLength);

	AttachmentInfo *item = new AttachmentInfo;
	retW2A = TextUtilsEx::Wide2Ansi(validNameW, item->m_name, error);
	item->m_nameW = validNameW;
	m_attachmentTbl.Add(item);

	return TRUE;
}

BOOL CAttachments::AddInlineAttachment(CString &name)
{
	// name already validated in DetermineImageFileName_SelectedItem
	AttachmentInfo *item = new AttachmentInfo;
	item->m_name.Append(name);
	DWORD error;
	BOOL retW2A = TextUtilsEx::Ansi2Wide(name, item->m_nameW, error);
	m_attachmentTbl.Add(item);

	return TRUE;
}

void CAttachments::OnRClick(NMHDR* pNMHDR, LRESULT* pResult)
{
	const char *Open = _T("Open");
	const char *Print = _T("Print");
	const char *OpenFileLocation = _T("Open File Location");

	LPNMITEMACTIVATE pnm = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);

	*pResult = 0;

	int nItem = pnm->iItem;
	if (nItem < 0)
		return;

	CStringW attachmentName = m_attachmentTbl[nItem]->m_nameW;

	TRACE("Selecting %d\n", pnm->iItem);

	// TODO: Add your control notification handler code here
#define MaxShellExecuteErrorCode 32

	CPoint pt;
	::GetCursorPos(&pt);
	CWnd *wnd = WindowFromPoint(pt);

	CMenu menu;
	menu.CreatePopupMenu();
	menu.AppendMenu(MF_SEPARATOR);

	const UINT M_OPEN_Id = 1;
	AppendMenu(&menu, M_OPEN_Id, Open);

	const UINT M_PRINT_Id = 2;
	AppendMenu(&menu, M_PRINT_Id, Print);

	const UINT M_OpenFileLocation_Id = 3;
	AppendMenu(&menu, M_OpenFileLocation_Id, OpenFileLocation);

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
	if (command == M_PRINT_Id)
	{
		CStringW path = FileUtils::GetmboxviewTempPathW();
		CStringW filePath = path + attachmentName;
		result = ShellExecuteW(NULL, L"print", attachmentName, NULL, path, SW_SHOWNORMAL);
		if ((UINT)result <= MaxShellExecuteErrorCode) {
			CString errorText;
			ShellExecuteError2Text((UINT)result, errorText);
			errorText += _T(".\nOk to try to open this file ?");
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, errorText, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_YESNO);
			if (answer == IDYES)
				forceOpen = true;
		}
		int deb = 1;
	}
	if ((command == M_OPEN_Id) || forceOpen)
	{
		CStringW path = FileUtils::GetmboxviewTempPathW();
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
		CStringW path = FileUtils::GetmboxviewTempPathW();
		CStringW filePath = path + attachmentName;

		if (FileUtils::BrowseToFileW(filePath) == FALSE) {
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

	CString extA;
	DWORD error;
	BOOL retW2A = TextUtilsEx::Wide2Ansi(ext, extA, error);

	bool isSupportedPictureFile = false;

	// Need to create tavble of extensions so this class and Picture viewr class can share the list
	if (CCPictureCtrlDemoDlg::isSupportedPictureFile(extA))
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
		CStringW path = FileUtils::GetmboxviewTempPathW();
		CStringW filePath = path + attachmentNameW;
		// Photos application doesn't show next/prev photo even with path specified
		// Buils in Picture Viewer is set as deafault
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
