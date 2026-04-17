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
#include "ResHelper.h"
#include "MboxMail.h"

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

BOOL CAttachments::InsertItemW(CStringW &cStrName, int id, CStringA& mediaType, CStringA& mediaSubtype)
{
	int index = CAttachments::FindAttachmentByNameW(cStrName);
	// Caller must assure names are unique
	_ASSERTE(index < 0);

	AttachmentInfo *item = new AttachmentInfo;
	item->m_nameW = cStrName;
	item->m_mediaType = mediaType;
	item->m_mediaSubtype = mediaSubtype;
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
	if ((nItem < 0) || (nItem >= m_attachmentTbl.GetCount()))
		return;

	AttachmentInfo* attachmentInfo = m_attachmentTbl[nItem];
	CStringW attachmentName = attachmentInfo->m_nameW;

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

	int index1 = 0;
	ResHelper::UpdateMenuItemsInfo(&menu, index1);

	int command = menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, this);

	int index = 0;
	ResHelper::LoadMenuItemsInfo(&menu, index);
	//ResHelper::UpdateMenuItemsInfo(&menu, index);

	UINT nFlags = TPM_RETURNCMD;
	CString menuString;
	int chrCnt = menu.GetMenuString(command, menuString, nFlags);

	bool forceOpen = false;
	HINSTANCE result = (HINSTANCE)(MaxShellExecuteErrorCode + 1);  // OK

	CString path = CMainFrame::GetMboxviewTempPath();
	if (command == M_PRINT_Id)
	{
		CStringW filePath = path + attachmentName;
		result = ShellExecuteW(NULL, L"print", attachmentName, NULL, path, SW_SHOWNORMAL);
		if ((UINT_PTR)result <= MaxShellExecuteErrorCode)
		{
			CString errorText;
			CString errText;
			ShellExecuteError2Text((UINT_PTR)result, errText);
			CString fmt = L"\"%s\"\n\nOk to try to open this file ?";
			ResHelper::TranslateString(fmt);
			errorText.Format(fmt, errText);


			HWND h = GetSafeHwnd();
			int answer = MyMessageBox(h, errorText, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_YESNO);
			if (answer == IDYES)
				forceOpen = true;
		}
		int deb = 1;
	}
	if ((command == M_OPEN_Id) || forceOpen)
	{
		int retCode = CAttachments::OpenAttachment(attachmentInfo);
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

	//
	int nItem = pnm->iItem;
	if ((nItem < 0) || (nItem >= m_attachmentTbl.GetCount()))
		return;

	AttachmentInfo* attachmentInfo = m_attachmentTbl[nItem];
	CStringW attachmentNameW = attachmentInfo->m_nameW;

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
		int retCode = CAttachments::OpenAttachment(attachmentInfo);
		int deb = 1;
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

BOOL CAttachments::GetWinmail2EmlProcessPath(CString& processPath)
{
	processPath.Empty();
	CString mboxviewPath = CMainFrame::m_processPath;

	CString processDir;
	FileUtils::CPathGetPath(mboxviewPath, processDir);

	// Release path
	CString winmail2emlExePath;

#ifdef USE_STACK_WALKER
	winmail2emlExePath = processDir + LR"(\..\Winmail2Eml\Winmail2Eml.exe)";
#else
	winmail2emlExePath = processDir + LR"(\Winmail2Eml\Winmail2Eml.exe)";
#endif

	if (FileUtils::PathFileExist(winmail2emlExePath))
	{
		processPath = winmail2emlExePath;
		return TRUE;
	}

#ifndef USE_STACK_WALKER
	winmail2emlExePath = processDir + LR"(\..\Winmail2Eml\Winmail2Eml.exe)";
#else
	winmail2emlExePath = processDir + LR"(\Winmail2Eml\Winmail2Eml.exe)";
#endif

	if (FileUtils::PathFileExist(winmail2emlExePath))
	{
		processPath = winmail2emlExePath;
		return TRUE;
	}

	// Development path
	CString baseDir = processDir;
	baseDir.TrimRight(L"\\");
	CString baseDir2;
	FileUtils::GetFolderPath(baseDir, baseDir2);
	baseDir2.TrimRight(L"\\");
	FileUtils::GetFolderPath(baseDir2, baseDir);
	baseDir.TrimRight(L"\\");

	if (!FileUtils::PathFileExist(winmail2emlExePath))
	{
		winmail2emlExePath = baseDir + LR"(\Winmail2Eml\bin\Debug\net8.0\Winmail2Eml.exe)";
	}
	if (!FileUtils::PathFileExist(winmail2emlExePath))
	{
		winmail2emlExePath = baseDir + LR"(\Winmail2Eml\bin\Debug\net8.0-windows\Winmail2Eml.exe)";
	}

	if (!FileUtils::PathFileExist(winmail2emlExePath))
	{
		winmail2emlExePath = baseDir + LR"(\Winmail2Eml\bin\Release\net8.0\Winmail2Eml.exe)";
	}
	if (!FileUtils::PathFileExist(winmail2emlExePath))
	{
		winmail2emlExePath = baseDir + LR"(\Winmail2Eml\bin\Release\net8.0-windows\Winmail2Eml.exe)";
	}

	if (FileUtils::PathFileExist(winmail2emlExePath))
	{
		processPath = winmail2emlExePath;
		return TRUE;
	}
	else
		return FALSE;
}

int CAttachments::OpenAttachment(AttachmentInfo* attachmentInfo)
{
	CString path = CMainFrame::GetMboxviewTempPath();

	CString attachmentName = attachmentInfo->m_nameW;
	CString filePath = path + attachmentName;
	CStringW ext = PathFindExtensionW(attachmentName);

	// Windows Photos application doesn't show next/prev photo even with path specified
	// Build in Picture Viewer is set as deafault

	HWND h = GetSafeHwnd();

#if 0
	// for testing 
	CString traceTxt;
	traceTxt.Format(L"filePath=%s", filePath);
	int answer = MyMessageBox(h, traceTxt, L"Info", MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
#endif

	CString mboxviewPath = CMainFrame::m_processPath;

	HINSTANCE result = (HINSTANCE)(MaxShellExecuteErrorCode + 1);  // OK

	CString  fileNameExtention = ext;

	if ((fileNameExtention.CompareNoCase(L".eml") == 0) ||
		(fileNameExtention.CompareNoCase(L".mbox") == 0) ||
		(fileNameExtention.CompareNoCase(L".mboxrd") == 0) ||
		(fileNameExtention.CompareNoCase(L".mboxcl") == 0) ||
		(fileNameExtention.CompareNoCase(L".mboxcl2") == 0))
	{
		CString fPath = "\"" + filePath + "\"";
		result = ShellExecute(NULL, L"open", mboxviewPath, fPath, path, SW_SHOWNORMAL);
	}
	else if (fileNameExtention.CompareNoCase(L".msg") == 0)
	{
		CString fPath = "\"" + filePath + "\"";
		result = ShellExecute(NULL, L"open", mboxviewPath, fPath, path, SW_HIDE);
	}
	else if (((attachmentName.CompareNoCase(L"winmail.dat") == 0) || (fileNameExtention.CompareNoCase(L".ms-tnef") == 0)) ||
		((attachmentInfo->m_mediaSubtype.CompareNoCase("ms-tnef") == 0) || (attachmentInfo->m_mediaSubtype.CompareNoCase("wnd.ms-tnef") == 0)))
		{
			CString winmail2emlExePath;
			BOOL pathFound = CAttachments::GetWinmail2EmlProcessPath(winmail2emlExePath);

			CString inputWinmailFilePath = filePath;

			BOOL isTnefFileType = IsTnefFile(inputWinmailFilePath);

			CString datapath = MboxMail::GetLastDataPath();
			CString winmail2emlCachePath = datapath + L"Winmail2EmlCache";
			BOOL retCreate = FileUtils::CreateDir(winmail2emlCachePath);
			CString outputEmlFilePath = winmail2emlCachePath + L"\\" + attachmentName + ".eml";

			CString languageName;
			CString targetHtmlLanguageCode = "en";
			if (!ResHelper::IsEnglishConfigured(languageName))
			{
				CString section_general = CString(sz_Software_mboxview) + L"\\General";
				CString language = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"language");
				targetHtmlLanguageCode = ResHelper::GetLanguageCode(languageName);
			}

			CString args = "--mboxview-exe-path \"" + mboxviewPath +
				"\" --input-winmail-file \"" + inputWinmailFilePath +
				"\" --output-eml-file \"" + outputEmlFilePath +
				"\" --target-html-language-code \"" + targetHtmlLanguageCode +
				"\"";

			if (pathFound)
			{
				result = ShellExecute(NULL, L"open", winmail2emlExePath, args, path, SW_HIDE);
			}
			else
				; // alert ? MessageBox ??
	}
	else
	{
		result = ShellExecuteW(h, L"open", attachmentName, NULL, path, SW_SHOWNORMAL);
	}
	CMainFrame::CheckShellExecuteResult(result, h, &filePath);
	return 0;
}

BOOL CAttachments::IsTnefFile(CString& cStrNamePath)
{
	SimpleString txt;
	int bytes2Read = 8;  // or 512 if find "ipm.microsoft"; 

	DWORD dwAccess = GENERIC_READ;
	DWORD dwCreationDisposition = OPEN_EXISTING;

	HANDLE hFile = CreateFile(cStrNamePath, dwAccess, FILE_SHARE_READ, NULL, dwCreationDisposition, 0, NULL);
	if (hFile == INVALID_HANDLE_VALUE)
	{
		int deb = 1;
		return FALSE;
	}
	else
	{
		LPOVERLAPPED lpOverlapped = 0;
		LPVOID       lpBuffer;
		DWORD        nNumberOfBytesToRead;
		DWORD      nNumberOfBytesRead;

		_int64 fsize = FileUtils::FileSize(cStrNamePath);
		if (fsize < bytes2Read)
			return FALSE;

		_int64 bytesLeft = bytes2Read;
		txt.ClearAndResize((int)bytes2Read + 1);
		lpBuffer = txt.Data();
		while (bytesLeft > 0)
		{
			if (bytesLeft > 0xffffffff)
				nNumberOfBytesToRead = 0xffffffff;
			else
				nNumberOfBytesToRead = (DWORD)bytesLeft;

			BOOL retval = ::ReadFile(hFile, lpBuffer, nNumberOfBytesToRead, &nNumberOfBytesRead, lpOverlapped);
			if (retval == FALSE)
				break;

			int newCount = txt.Count() + nNumberOfBytesRead;
			txt.SetCount(newCount);
			lpBuffer = txt.Data(newCount);
			bytesLeft -= nNumberOfBytesRead;
		}
		BOOL retClose = CloseHandle(hFile);

		char* data = txt.Data();

		UINT8 signature0 = (UINT8)data[0];
		UINT8 signature1 = (UINT8)data[1];
		UINT8 signature2 = (UINT8)data[2];
		UINT8 signature3 = (UINT8)data[3];
		if ((signature0 != 0x78) ||
			(signature1 != 0x9F) ||
			(signature2 != 0x3E) ||
			(signature3 != 0x22))
		{
			//_ASSERTE(FALSE);
			return FALSE;
		}

#if 0
		// ignore for now
		char* lowCasePat = "ipm.microsoft";
		int lowCasePatLen = strlen(lowCasePat);
		int isMail = txt.FindNoCase(0, lowCasePat, lowCasePatLen);
		_ASSERTE(isMail >= 0);
		if (isMail < 0)
			return FALSE;
#endif

		return TRUE;
	}
}
