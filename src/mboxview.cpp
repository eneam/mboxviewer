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

// mboxview.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "ctype.h"
#include "locale.h"
#include "FileUtils.h"
#include "TextUtilsEx.h"
#include "mboxview.h"
#include "profile.h"
#include "LinkCursor.h"
#include "MainFrm.h"
#include "MboxMail.h"

#include <afxadv.h> //Has CRecentFileList class definition.
#include "afxlinkctrl.h"
#include "afxwin.h"

#include "ExceptionUtil.h"

#include "winnls.h"
#include "FileConfigurationDB.h"

#include "MimeHelper.h"
#include "SimpleString.h"
#include "ResHelper.h"

#include "Shobjidl.h"
#include "shobjidl_core.h"

int JsonTest();

//#include "afxglobals.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
//#define _CRTDBG_MAP_ALLOC  
//#include <stdlib.h>  
//#include <crtdbg.h> 
#endif

// UpdateData(FALSE) from Dialog members to UI controls 
// UpdateData(TRUE) from UI controls to Dialog members -  be carefull if you call TRUE in dialog followed by CANCEL

static int unhandledExceptionsCnt = 0;

DWORD CmboxviewApp::m_versionMS = 0;
DWORD CmboxviewApp::m_versionLS = 0;
CString CmboxviewApp::m_savedVer = L"";
CString CmboxviewApp::m_processPath = L"";
CString CmboxviewApp::m_startupPath = L"";
CString CmboxviewApp::m_currentPath = L"";


ConfigTree* CmboxviewApp::m_configTree = 0;
BOOL CmboxviewApp::m_registry = TRUE;
CString CmboxviewApp::m_configFilePath;
BOOL CmboxviewApp::m_configFileLoaded = FALSE;
//DebugCString CmboxviewApp::m_configFilePath;

CWnd* CmboxviewApp::wndFocus = 0;

//#pragma warning(3 : 4840 4477)  // useless -:(((

LONG WINAPI MyUnhandledExceptionFilter(PEXCEPTION_POINTERS pExceptionPtrs)
{
	// Do something, for example generate error report
	//..

	//MessageBeep(MB_OK);
	//MessageBeep(MB_OK);
	//MessageBeep(MB_OK);

	UINT seNumb = 0;
	if (pExceptionPtrs->ExceptionRecord)
		seNumb = pExceptionPtrs->ExceptionRecord->ExceptionCode;

	if (MboxMail::ignoreException)
	{
		return EXCEPTION_CONTINUE_SEARCH;
	}

	if (seNumb < 0x80000000)
		return EXCEPTION_CONTINUE_SEARCH;

	// MyUnhandledExceptionFilter() is called as a result of calling AddVectoredExceptionHandler(first, MyUnhandledExceptionFilter);
	// This results invoking MyUnhandledExceptionFilter() for all exceptions
	// We only want to generate the stack trace when exception is caused by mboxviewer code
#ifdef USE_STACK_WALKER
	if (MboxMail::glStackWalker)
	{
		BOOL ret = MboxMail::glStackWalker->ProcessStackTrace(seNumb, pExceptionPtrs);
		if (!ret)
			return EXCEPTION_CONTINUE_SEARCH;
		else
			int deb = 1;
	}
#endif

	if (unhandledExceptionsCnt > 0)
		exit(0);

	if (unhandledExceptionsCnt == 0)
	{

		char const* szCause = seDescription(seNumb);

		wchar_t* stackDumpFileName = L"UnhandledException_StackDump.txt";
		int mailPosition = MboxMail::s_mails.GetCount();
		MyStackWalker* sw = nullptr;
		if (MboxMail::glStackWalker)
		{
			sw = MboxMail::glStackWalker;
		}
		BOOL ret = DumpStack(sw, stackDumpFileName, (CHAR*)szCause, seNumb, pExceptionPtrs->ContextRecord, mailPosition);

		CString progDir;
		BOOL retDir = GetProgramDir(progDir);

		wchar_t* exceptionName = L"UnhandledException";
		CString errorTxt;
		CString szCauseW(szCause);
#ifdef USE_STACK_WALKER
		errorTxt.Format(L"%s: Code=%8.8x Description=%s\n\n"
			"To help to diagnose the problem, created file\n\n%s\n\nin\n\n%s directory.\n\n"
			"Please provide the files to the development team.\n\n",
			exceptionName, seNumb, szCauseW, stackDumpFileName, progDir);
#else
		errorTxt.Format(L"%s: Code=%8.8x Description=%s\n\n", exceptionName, seNumb, szCauseW);
#endif
		AfxMessageBox((LPCWSTR)errorTxt, MB_OK | MB_ICONHAND);

		MboxMail::ignoreException = TRUE;

		if (!MboxMail::runningWorkerThreadType)
		{
			//MessageBeep(MB_OK);
			//MessageBeep(MB_OK);
			//MessageBeep(MB_OK);
			//AfxAbort();
			exit(0);
		}
	}

	unhandledExceptionsCnt++;

	// Execute default exception handler next
	//return EXCEPTION_EXECUTE_HANDLER;
	//return EXCEPTION_CONTINUE_EXECUTION;
	return EXCEPTION_CONTINUE_SEARCH;
}

// CRT invalid parameter handler
void __cdecl InvalidParameterHandler(
	const wchar_t* expression,
	const wchar_t* function,
	const wchar_t* file,
	unsigned int line,
	uintptr_t pReserved)
{
	pReserved;

	// Invalid parameter exception

	// Retrieve exception information
	//EXCEPTION_POINTERS* pExceptionPtrs = GetExceptionInformation();

	CStringW lineW;
	lineW.Format(L"%d", line);
	CStringW txt = L"expression=" + CStringW(expression) + L"function=" + CStringW(function) + L"file=" + CStringW(file) + L"line=" + lineW;

	CStringA txtA;
	DWORD error;
	BOOL retW2A = TextUtilsEx::WStr2Ansi(txt, txtA, error);

	AfxMessageBox(txt, MB_OK | MB_ICONHAND);

	// GetExceptionPointers(0, &pExceptionPtrs);

	// Write minidump file
	//CreateMiniDump(pExceptionPtrs);

	// Terminate process
	TerminateProcess(GetCurrentProcess(), 1);
}


// It looks by AddVectoredExceptionHandler() will catch all kinds of exceptions,
// However, some of them must be ignored and it is difficult to know which exceptions
// must be ignored. Unfortunatelly, in rare unknown cases program will be terminated incorrectly.

PVOID vectorHandle = 0;
void SetMyExceptionHandler()
{
	//Scoped_SE_Translator scoped_se_translator(trans_func);
	//_set_invalid_parameter_handler(InvalidParameterHandler);
	//SetUnhandledExceptionFilter(MyUnhandledExceptionFilter);

	ULONG first = 1;
	vectorHandle = AddVectoredExceptionHandler(first, MyUnhandledExceptionFilter);
}

void UnSetMyExceptionHandler()
{
	//Scoped_SE_Translator scoped_se_translator(trans_func);
	//SetUnhandledExceptionFilter(MyUnhandledExceptionFilter);

	ULONG ret = RemoveVectoredExceptionHandler(vectorHandle);

	//_set_invalid_parameter_handler(InvalidParameterHandler);
}

const wchar_t* sz_Software_mboxview = L"SOFTWARE\\UMBoxViewer";

/////////////////////////////////////////////////////////////////////////////
// CmboxviewApp

BEGIN_MESSAGE_MAP(CmboxviewApp, CWinApp)
	//{{AFX_MSG_MAP(CmboxviewApp)
	ON_COMMAND(ID_APP_ABOUT, OnAppAbout)
	ON_COMMAND(ID_HELP_DONATE, OnHelpDonate)
	// NOTE - the ClassWizard will add and remove mapping macros here.
	//    DO NOT EDIT what you see in these blocks of generated code!
//}}AFX_MSG_MAP
ON_COMMAND_RANGE(ID_FILE_MRU_FIRST, ID_FILE_MRU_LAST, MyMRUFileHandler)
//	ON_UPDATE_COMMAND_UI(ID_FILE_MRU_FILE1, OnUpdateRecentFileMenu)
END_MESSAGE_MAP()


void Com_Initialize()
{
	//DWORD dwCoInit = COINIT_MULTITHREADED;
	DWORD dwCoInit = COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE;

	// call to CoInitializeEx() seem to be required but 
	// Fails anyway with "HRESULT - 0x80010106 - Cannot change thread mode after it is set. "
	HRESULT result = CoInitializeEx(0, dwCoInit);
	int deb = 1;
}


MyCRecentFileList::MyCRecentFileList(UINT nStart, LPCTSTR lpszSection,
	LPCTSTR lpszEntryFormat, int nSize, int nMaxDispLen)
	: CRecentFileList(nStart, lpszSection, lpszEntryFormat, nSize, nMaxDispLen)
{
	int deb = 1;

}

void MyCRecentFileList::DumpMRUsArray(CString &prefix)
{
	CString val;
	for (int iMRU = 0; iMRU < m_nSize; iMRU++)
	{
		val = m_arrNames[iMRU];
		TRACE(L"ReadList: %d%s \"%s\"\n", (iMRU+1), prefix, val);
	}
}

void MyCRecentFileList::ReadList()
{
#if 0
	this->CRecentFileList::ReadList();
#else
	ASSERT(m_arrNames != NULL);
	ASSERT(!m_strSectionName.IsEmpty());
	ASSERT(!m_strEntryFormat.IsEmpty());

	CString section = L"SOFTWARE\\UMBoxViewer\\mboxview\\" + m_strSectionName;
	int nLen = m_strEntryFormat.GetLength() + 10;
	LPTSTR pszEntry = new TCHAR[nLen];

	CString val;
	
	TRACE(L"ReadList:\n");
	for (int iMRU = 0; iMRU < m_nSize; iMRU++)
	{
		_stprintf_s(pszEntry, nLen, m_strEntryFormat, iMRU + 1);
		val = CProfile::_GetProfileString(HKEY_CURRENT_USER, section, pszEntry);
		m_arrNames[iMRU] = val;
	}
	CString prefix;
	DumpMRUsArray(prefix);
	delete[] pszEntry;
#endif
}

void MyCRecentFileList::WriteList()
{
	//
	// this->CRecentFileList::WriteList();
	ASSERT(m_arrNames != NULL);
	ASSERT(!m_strSectionName.IsEmpty());
	ASSERT(!m_strEntryFormat.IsEmpty());

	CString parent_section = L"SOFTWARE\\UMBoxViewer\\mboxview";
	CString section = parent_section + L"\\" + m_strSectionName;

	int nLen = m_strEntryFormat.GetLength() + 10;
	LPTSTR pszEntry = new TCHAR[nLen];
	CWinApp* pApp = AfxGetApp();

	CString val;

	CProfile::_DeleteKey(HKEY_CURRENT_USER, parent_section, m_strSectionName, TRUE);

	TRACE(L"WriteList:\n");
	for (int iMRU = 0; iMRU < m_nSize; iMRU++)
	{
		_stprintf_s(pszEntry, nLen, m_strEntryFormat, iMRU + 1);
		val = m_arrNames[iMRU];
		if (!val.IsEmpty())
		{
			CProfile::_WriteProfileString(HKEY_CURRENT_USER, section, pszEntry, val);
		}
	}
	CString prefix;
	DumpMRUsArray(prefix);
	delete[] pszEntry;
}

void CmboxviewApp::AddToRecentFileList(LPCWSTR lpszPathName)
{
	ASSERT_VALID(this);
	_ASSERTE(m_pRecentFileList != NULL);

	CString pathName = lpszPathName;
	pathName.TrimRight(L"\\");
	CString filePathName;
	int count;
	int i;
	BOOL removed = FALSE;
	count = m_pRecentFileList->GetSize();
	for (i = 0; i < count; i++)
	{
		filePathName = m_pRecentFileList->m_arrNames[i];
		filePathName.TrimRight(L"\\");
		TRACE(L"AddToRecentFileList: \"%s\" \"%s\"\n", lpszPathName, filePathName);
		if (filePathName.Compare(pathName) == 0)
		{
			m_pRecentFileList->Remove(i);
			int newCount = m_pRecentFileList->GetSize();
			removed = TRUE;
			break;
		}
	}
	this->CWinApp::AddToRecentFileList(lpszPathName);
	if (removed)
		m_pRecentFileList->WriteList();
}

void CmboxviewApp::MyMRUFileHandler(UINT i)
{
	ASSERT_VALID(this);
	_ASSERTE(m_pRecentFileList != NULL);

	_ASSERTE(i >= ID_FILE_MRU_FILE1);
	_ASSERTE(i < ID_FILE_MRU_FILE1 + (UINT)m_pRecentFileList->GetSize());

	CString strName, strCurDir, strMessage;
	int nIndex = i - ID_FILE_MRU_FILE1;

	CString folderPath = (*m_pRecentFileList)[nIndex];
	_ASSERTE(folderPath.GetLength() != 0);

	strName.Format(L"MRU: open file (%d) '%s'.\n", (nIndex)+1, (LPCWSTR)folderPath);

	if (!FileUtils::PathDirExists(folderPath)) {
		m_pRecentFileList->Remove(nIndex);
		return;
	}
	// thias will move to top selected file
	else
	{
		m_pRecentFileList->Remove(nIndex);

		this->CWinApp::AddToRecentFileList(folderPath);
		m_pRecentFileList->WriteList();
	}

	((CMainFrame*)AfxGetMainWnd())->DoOpen(folderPath);
}

/////////////////////////////////////////////////////////////////////////////
// CmboxviewApp construction

//#define IHASH_MAP_TEST
#ifdef  IHASH_MAP_TEST
#include "IHashTable.h"
#endif

#include "winerror.h"

BOOL GetErrorMessage(DWORD dwErrorCode, CString& errorMessage)
{
	wchar_t Buffer[1024];
	DWORD cchBufferLength = 1022;

	Buffer[0] = 0;
	DWORD cchMsg = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
		NULL,  /* (not used with FORMAT_MESSAGE_FROM_SYSTEM) */
		dwErrorCode,
		0,  //MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
		Buffer,
		cchBufferLength,
		NULL);
	CString errMessage = CString(Buffer);
	errMessage.TrimRight(L" \t\r\n");
	errorMessage.Format(L"Error Code %u \"%s\"\n", dwErrorCode, errMessage);

	return (cchMsg > 0);
}

#include <iostream>
#include <commctrl.h>

void CreateAnsiToUTF8TableStr(CString& tbl, int firstPos, int lastPos)
{
	DWORD error;
	CString cStr;
	CString cStr2;
	CString cStr3;
	CString cStr4;
	CStringA strA;
	CStringA inStr;
	int i;
	UINT inCodePage = GetACP();
	for (i = firstPos; i <= lastPos; i++)
	{
		cStr.Empty();
		inStr.Format("%c", i);
		//if ((i == 129) || (i == 134)) inStr = " ";
		cStr4 = inStr;
		int len4 = cStr4.GetLength();
		if (cStr4.IsEmpty())
			cStr4 = L" ";
		cStr3.Format(L"%s", CString(inStr));
		int len3 = cStr3.GetLength();
		if (cStr3.IsEmpty())
			cStr3 = L" ";

		strA.Empty();
		TextUtilsEx::Str2UTF8((LPCSTR)inStr, inStr.GetLength(), inCodePage, strA, error);
		cStr.Format(L"%03d 0x%02x 0x%04x %s %s -- ", i, i, cStr4.GetAt(0), cStr4, cStr3);
		cStr.Append(L"0x");
		int j = 0;
		cStr2.Empty();
		for (j = 0; j < strA.GetLength(); j++)
		{
			unsigned char uic = strA.GetAt(j);
			cStr2.Format(L"%02x", uic);
			cStr.Append(cStr2);
		}
		tbl.Append(cStr);
		tbl.Append(L"\n");
	}
}
void CreateCodePageToUTF16TableStr(UINT inCodePage, CString& tbl, int firstPos, int lastPos)
{
	CStringA astr;
	CStringA afmt;
	int i;

	for (i = firstPos; i <= lastPos; i++)
	{
		afmt.Format("%c", i);
		astr.Append(afmt);
	}

	CString wstr;
	CString wfmt;
	DWORD error;
	DWORD dwFlags = 0;

	CString a2wstr;
	CString out;

	BOOL ret = TextUtilsEx::CodePage2WStr(&astr, inCodePage, &wstr, error, dwFlags);
	for (i = firstPos; i <= lastPos; i++)
	{
		a2wstr = CString(astr[i-firstPos]);
		wfmt.Format(L"%c %02x  %lc %04x\n", i, i, wstr[i - firstPos], wstr[i-firstPos]);
		tbl.Append(wfmt);
	}
	tbl.Append(wstr);
	int deb = 1;
}

class ReqFromToInfo
{
public:
	UINT m_type;
	CString m_from;
	CStringW m_to;
};

#if (_WIN32_WINNT >= _WIN32_WINNT_WIN7)

HRESULT CheckIsDefaultApp(const wchar_t* prog, const wchar_t *extension, BOOL* fDefault)
{
	IApplicationAssociationRegistration* pAAR;

	HRESULT hr = CoCreateInstance(CLSID_ApplicationAssociationRegistration,
		NULL, CLSCTX_INPROC,
		__uuidof(IApplicationAssociationRegistration),
		(void**)&pAAR);

	if (SUCCEEDED(hr))
	{
		hr = pAAR->QueryAppIsDefault(extension,
			AT_FILEEXTENSION, AL_EFFECTIVE,
			prog,
			fDefault);

		pAAR->Release();
	}

	return hr;
}
#endif

HWND CmboxviewApp::GetActiveWndGetSafeHwnd()
{
	HWND h = NULL; // we don't have any window yet
	CWnd* wnd = CWnd::GetActiveWindow();
	if (wnd)
		h = wnd->GetSafeHwnd();
	return h;
}

CmboxviewApp::~CmboxviewApp()
{
	if (!CProfile::IsRegistryConfig())
	{
		ConfigTree* confTree = m_configTree;
		_ASSERTE(confTree);
		if (confTree)
		{
			confTree->DumpTree();

			if (CmboxviewApp::m_configFileLoaded)
				confTree->Dump2File();

			confTree->DeleteAllNodes();

			delete confTree;
			m_configTree = 0;
		}
	}

	int deb = 1;
}

CmboxviewApp::CmboxviewApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
	int deb10 = 1;

#ifdef _DEBUG

#if 0
	{
		int cnt = ResHelper::MonitorInfo();
		int deb = 1;
	}
#endif

#if 0

	UINT driveType = GetDriveType(L"Z:");
	if (driveType == DRIVE_REMOTE)
		int deb = 1;
	int debZ = 1;

#endif

#if 0
	// Language file is not loaded yet
	{
		CString filePath = LR"(F:\Documents\GIT1.0.3.42\mboxviewer\x64\Debug\CHANGE_LOG.md.txt)";
		HWND hw = 0;
		//hw = GetSafeHwnd();

		CString badFilePath = L"C:\\Users\\UserA";
		//badFilePath = filePath + L".txt";;

		CStdioFile file;
		CFileException exList;
		UINT nOpenFlags = CFile::modeRead | CFile::typeText | CFile::shareExclusive;
		if (!file.Open(badFilePath, nOpenFlags, &exList))
		{
			DWORD lastErr = ::GetLastError();
			CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exList, lastErr);
			int answer = FileUtils::CheckIfFileLocked(filePath, lastErr, hw);

			CString errorStr = FileUtils::GetErrorAsString(exList, lastErr);
			int deb = 1;
		}
		else
			file.Close();

		int deb = 1;
	}

	int deb = 1;
#endif

#if 0

	CString fmt = L"arg1=%1 arg2=%2 arg3= %1";

	CString mesg = ResHelper::TranslateMsg(&fmt, L"arg1", L"arg2");

	int deb = 1;
#endif

#if 0
//
#pragma warning(1 :  4473 4474 4475 4476 4477 4478 4774 4775 4776 4778 4840 6603)  // useless -:(((
//#pragma warning( push , 1 )  // useless -:(((
// Check whther compiler complains

#if 0
	wchar_t const* str = L" ";// Some string to parse

	char buf[10];

	wchar_t wbf;

	static const wchar_t* fmt = L"%10c %1C";

	swscanf_s(str, fmt, buf, sizeof(buf), &wbf);  // useless ; fmt has to be string literal

#endif
	CString s1;
	int int1 = 7;
	int int2 = 077;
	INT_PTR intptr1 = 0;
	double dble = 0;
	CString fmt = L"string %s integer %d integer2 %00d";
	CString eval;

	// will crash and  generate  exceptions
	eval.Format(L"string %s integer %d integer2 %00d", dble, (LPCWSTR)s1, int1, int2);  
	eval.Format((LPCWSTR)fmt, int1, (LPCWSTR)s1, int1, int2); 
	eval.Format((LPCWSTR)fmt, int1); 

	wprintf(L"string %s integer %d integer2 %00d", (LPCWSTR)s1, int1, int2);  
	wprintf((LPCWSTR)fmt, (LPCWSTR)s1, int1, int2);  /
	wprintf((LPCWSTR)fmt, int1, int2, (LPCWSTR)s1);  
	wprintf((LPCWSTR)fmt, int1);  

//#pragma warning( pop)
#endif


#if 0
	UINT acpCodePage = GetACP();
	UINT inCodePage = 1252;
	//inCodePage = 850;
	//inCodePage = 1253;
	CString tbl;
	int firstPos = 31;
	int lastPos = 255;
	CreateCodePageToUTF16TableStr(inCodePage, tbl, firstPos, lastPos);
	TRACE(L"%s\n", tbl);
	int deb = 1;
#endif

#if 0
	// Doesn't work for my case
	CString szProg = L"mboxview";
	CString szQuery = L".eml";
	szQuery = L".txt";

	//ASSOCIATIONTYPE  atQueryType = 0;
	//ASSOCIATIONLEVEL alQueryLevel = 0;
	//CString pszAppRegistryName = L"";
	BOOL fDefault = FALSE;

	DWORD dwCoInit = COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE;
	CoInitializeEx(NULL, dwCoInit);
	HRESULT ret = CheckIsDefaultApp((LPCWSTR)szProg, (LPCWSTR)szQuery, &fDefault);
	CoUninitialize();
	int deb5 = 1;


#if 0
	{
		IApplicationAssociationRegistration reg =
			(IApplicationAssociationRegistration) new ApplicationAssociationRegistration();

		string progID;

		reg.QueryCurrentDefault(".txt",
			ASSOCIATIONTYPE.AT_FILEEXTENSION,
			ASSOCIATIONLEVEL.AL_EFFECTIVE,
			out progID);
		Console.WriteLine(progID);

		reg.QueryCurrentDefault("http",
			ASSOCIATIONTYPE.AT_URLPROTOCOL,
			ASSOCIATIONLEVEL.AL_EFFECTIVE,
			out progID);
		Console.WriteLine(progID);
	}

	HRESULT ret = IApplicationAssociationRegistration::QueryAppIsDefault(
		szQuery,
		atQueryType,
		alQueryLevel,
		szAppRegistryName,
		&fDefault);
#endif

#endif


#if 0
	CString cstr;
	INT64 numb;
	numb = 9;
	TextUtilsEx::Int2WstrWithCommas(numb, cstr);
	numb = 98;
	TextUtilsEx::Int2WstrWithCommas(numb, cstr);
	numb = 987;
	TextUtilsEx::Int2WstrWithCommas(numb, cstr);
	numb = 9987;
	TextUtilsEx::Int2WstrWithCommas(numb, cstr);
	numb = 98987;
	TextUtilsEx::Int2WstrWithCommas(numb, cstr);
	numb = 987987;
	TextUtilsEx::Int2WstrWithCommas(numb, cstr);
	numb = 9987987;
	TextUtilsEx::Int2WstrWithCommas(numb, cstr);
	numb = 98987987;
	TextUtilsEx::Int2WstrWithCommas(numb, cstr);
	numb = 987987987;
	TextUtilsEx::Int2WstrWithCommas(numb, cstr);
	numb = -987987987;
	TextUtilsEx::Int2WstrWithCommas(numb, cstr);
	numb = +987987987;
	TextUtilsEx::Int2WstrWithCommas(numb, cstr);

	wchar_t* wstr = L"++987987987";
	TextUtilsEx::WStr2WstrWithCommas(wstr, wcslen(wstr), cstr);

	const int deb = 1;
#endif

#if 0
	CArray<ReqFromToInfo> regFromToTable;

	HKEY hKey = HKEY_CURRENT_USER;
	CString section = CString(sz_Software_mboxview) + L"\\General";

	section = CString(sz_Software_mboxview);

	const wchar_t* sz_Software_mboxview_Legacy = L"SOFTWARE\\mboxview";

	CString fromSection = CString(sz_Software_mboxview_Legacy) + L"\\folders";
	CString toSection = CString(sz_Software_mboxview) + L"\\MailFolders";

	TRACE(L"_CopyLegacyFolders\n");
	LSTATUS retCode0 = CProfile::CopyKey(hKey, (LPCWSTR)fromSection, (LPCWSTR)toSection);

	fromSection = CString(sz_Software_mboxview_Legacy) + L"\\MailService";
	toSection = CString(sz_Software_mboxview) + L"\\MailService";

	TRACE(L"_CopyLegacyFolders: MailService\n");
	LSTATUS retCode3 = CProfile::CopyKey(hKey, (LPCWSTR)fromSection, (LPCWSTR)toSection);

	TRACE(L"_CopyLegacyFolders: MailService SubKeys\n");
	LSTATUS retCode4 = CProfile::CopySubKeys(hKey, (LPCWSTR)fromSection, (LPCWSTR)toSection);

	fromSection = CString(sz_Software_mboxview_Legacy) + L"\\Window Placement";
	toSection = CString(sz_Software_mboxview) + L"\\WindowPlacement";

	TRACE(L"_CopyLegacyFolders: Window Placement\n");
	LSTATUS retCode5 = CProfile::CopyKey(hKey, (LPCWSTR)fromSection, (LPCWSTR)toSection);

	fromSection = CString(sz_Software_mboxview_Legacy) + L"\\mboxview\\MRUs";
	toSection = CString(sz_Software_mboxview) + L"\\mboxview\\MRUs";

	TRACE(L"_CopyLegacyFolders: MRUs\n");
	LSTATUS retCode6 = CProfile::CopyKey(hKey, (LPCWSTR)fromSection, (LPCWSTR)toSection);


	fromSection = CString(sz_Software_mboxview_Legacy);
	toSection = CString(sz_Software_mboxview) + L"\\General";

	KeyFromToTable arr;
	arr.SetSize(1);
	arr.RemoveAll();

	arr.SetSize(10);
	arr[0] = { REG_DWORD, L"messageWindowPosition", L"" };
	arr.SetSize(1);

	LSTATUS retCode10 = CProfile::CopyKeyValueList(hKey, fromSection, toSection, arr);

	toSection = CString(sz_Software_mboxview) + L"\\LastSelection";

	arr.RemoveAll();
	arr.SetSize(5);
	arr[0] = { REG_SZ, L"lastLabelFilePath", L"" };
	arr[1] = { REG_SZ, L"lastMailFilePath", L"" };
	arr[2] = { REG_DWORD, L"lastMailIndex", L"" };
	arr[3] = { REG_SZ, L"lastPath", L"" };
	arr[4] = { REG_DWORD, L"lastWhichSort", L"" };


	LSTATUS retCode8 = CProfile::CopyKeyValueList(hKey, fromSection, toSection, arr);

#endif

#if 0
	HKEY hKey = HKEY_CURRENT_USER;
	CString section = CString(sz_Software_mboxview);

	TRACE(L"EnumerateAllSubKeys\n");
	LSTATUS retCode1 = CProfile::EnumerateAllSubKeys(hKey, (LPCWSTR)section);

	section = CString(sz_Software_mboxview) + L"\\WindowPlacement";
	TRACE(L"EnumerateAllSubKeyValues\n");
	LSTATUS retCode2 = CProfile::EnumerateAllSubKeyValues(hKey, (LPCWSTR)section);

	int deb = 1;
#endif

#if 0
	// TEST if file name can be longet than 260 characters

	CFile testFile;
	CString fileName = "\\\\?";

	fileName.AppendChar('\\');

	fileName.Empty();
	fileName.Append("C:\\");
	for (int i = 0; i < 512; i++)
		fileName.Append("x");

	fileName.Append("Y.txt");

	UINT nFlags = CFile::modeWrite | CFile::modeCreate;
	BOOL ret = testFile.Open(fileName, nFlags);
	if (ret)
		testFile.Close();

	DWORD dwAccess = GENERIC_WRITE;
	DWORD dwCreationDisposition = CREATE_ALWAYS;

	CStringW fileNameW;
	DWORD error;

	BOOL r = TextUtilsEx::Ansi2WStr(fileName, fileNameW, error);

	HANDLE hFile = CreateFileW(fileNameW, dwAccess, FILE_SHARE_READ, NULL, dwCreationDisposition, 0, NULL);
	if (hFile != INVALID_HANDLE_VALUE)
	{
		BOOL retClose = CloseHandle(hFile);
	}
#endif


#if 0
	int i;
	for (i = 0; i < 255; i++)
	{
		int c = i;
		if (iscntrl(c))
		{
			TRACE(L"ANSI c=%d %x is Control charcater !!!\n", c, c);
		}
	}
	for (i = 0; i < 255; i++)
	{
		wint_t c = i;
		if (iswcntrl(c))
		{
			TRACE(L"Wide ANSI c=%d %x is Control charcater !!!\n", c, c);
		}
	}
	int deb33 = 1;

#endif

#if 0
	struct NodeData {
		int m_pos;
	};

	using STNode = SimpleTreeNodeEx<NodeData>;

	CString treeName = L"just a test";
	SimpleTreeEx<NodeData> st(treeName);

	STNode* parentNode = 0;
	STNode* node1 = st.InsertNode(parentNode, CString(L"node #1"));
	STNode* node2 = st.InsertNode(parentNode, CString(L"node #2"));

	STNode* node1_1 = st.InsertNode(node1, CString(L"node #1.1"));
	STNode* node2_1 = st.InsertNode(node1, CString(L"node #2.1"));

	int cnt = st.NodeCount(parentNode);


	int deb33 = 1;
#endif

#if 0
	typedef enum _NORM_FORM {
		NormalizationOther = 0,
		NormalizationC = 0x1,
		NormalizationD = 0x2,
		NormalizationKC = 0x5,
		NormalizationKD = 0x6
	} NORM_FORM;
#endif

#if 0
#if (NTDDI_VERSION >= NTDDI_VISTA)
	NORM_FORM form = NormalizationC;
	CString origStr = CString(L"T\u00e8st string \uFF54\uFF4F n\u00f8rm\u00e4lize");
	LPCWSTR   lpSrcString = origStr;
	int       cwSrcLength = origStr.GetLength() + 1;
	wchar_t ooutStr[1024];
	LPWSTR    lpDstString = &ooutStr[0];
	int       cwDstLength = 1024;

	if (IsNormalizedString(form, lpSrcString, cwSrcLength))
	{
		TRACE(L"Already normalized in this form\n");
		int deb = 1;
	}
	else
	{
		TRACE(L"String not normalized in this form\n");
		int deb = 1;
	}


	int retLength = NormalizeString(form, lpSrcString, cwSrcLength, lpDstString, cwDstLength);
	if (retLength <= 0 && GetLastError() != ERROR_SUCCESS)
	{
		if (GetLastError() == ERROR_NO_UNICODE_TRANSLATION)
		{
			TRACE(L"Invalid Unicode found at input character index %d\n", -retLength);
		}
		int deb = 1;
	}
	else
	{
		;// keep the original string. Better than noting ??
	}
#endif

	int deb = 1;
#endif

#if 0
	// TEST CStrinA to CStrinW , TRACE
	CStringA strA = "mama tata";
	CStringW strW = strA;
	const wchar_t* d = (LPCWSTR)strW;  // works

	TRACE("Ani %s\n", strA);
	TRACE("Wide %s\n", strW);   // doesn't work
	TRACE(L"Ani-L %s\n", strA);  // doesn't work and risky looking for two nulls ?
	TRACE(L"Wide-L %s\n", strW);  // works

	strW = "tata mama";
	TRACE(L"Wide-L %s\n", strW); // works

	strA = strW;
	TRACE("Ani %s\n", strA);   // works

	CStringW str2W("siostra");
	TRACE(L"Wide-L %s\n", strW);  // works

	CStringA str2A(L"siostra");
	TRACE("Ansi-L %s\n", str2A);   // works

	UINT myCodePage = GetACP();
	CString tbl;
	tbl.Format(L"ANSI-To-UTF8 TABLE for local Code Page %d\n", myCodePage);
	CreateAnsiToUTF8TableStr(tbl, 32, 127);
	TRACE(L"%s\n", tbl);

	tbl.Empty();
	tbl.Append(L"\n");
	CreateAnsiToUTF8TableStr(tbl, 128, 159);
	TRACE(L"%s\n", tbl);

	tbl.Empty();
	tbl.Append(L"\n");
	CreateAnsiToUTF8TableStr(tbl, 160, 255);
	tbl.Append(L"\n");
	TRACE(L"%s\n", tbl);
	int deb = 1;
#endif



#if 0
	// TaskDialog or MFC wrapper CTaskDialog are supported on Vista and higher. 
	// TODO: Implies XP will not be supported. Do we stil need support for XP ??
	int nButtonPressed = 0;
	HINSTANCE hInst = 0;
	TaskDialog(NULL, hInst, L"Hello", L"Box", L"Mail  xxxx   xxxxxxxxxxxxx   xxxxxx \nmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm  xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx nnnnnnnnnnnnnnnnnnnnnnn",
		//MAKEINTRESOURCE(IDS_APPLICATION_TITLE),
		//MAKEINTRESOURCE(IDS_DOSOMETHING),
		//MAKEINTRESOURCE(IDS_SOMECONTENT),
		TDCBF_OK_BUTTON | TDCBF_CANCEL_BUTTON,
		TD_WARNING_ICON,
		&nButtonPressed);

	if (IDOK == nButtonPressed)
	{
		// OK button pressed
	}
	else if (IDCANCEL == nButtonPressed)
	{
		// Cancel pressed
	}
#endif
#if 0
	int _snwprintf_s(
		wchar_t* buffer,
		size_t sizeOfBuffer,
		size_t count,
		const wchar_t* format[,
		argument] ...
	);
#endif

#if 0
	wchar_t c = L'S';
	CStringW wstr = c;
	//wstr.Format(L"The wide character value 'S' is: %c", c);
	//wstr.Format(L"%c", c);
	//print the character value
	//TRACE("The wide character value 'S' is: %S", c);
	TRACE(L"%s\n", wstr);

	//print the size of wide character
	TRACE(L"Wide character size is %d\n", sizeof(c));
	TRACE(L"Wide character size is %d\n", wstr.GetLength());

	int deb1 = 1;
#endif

#if 0
	CString rootFolder = "G:\\UMBoxViewer";
	CList<CString, CString&> folderList;
	CString errorText;
	int maxDepth = 10;

	BOOL bRetVal = FileUtils::GetFolderList(rootFolder, folderList, errorText, maxDepth);

	// Print all folders, i.e. not empty folder path
	CString folderPath;
	TRACE(L"List of non empty folders:\n");
	while (folderList.GetCount() > 0)
	{
		folderPath = folderList.RemoveHead();
		TRACE(L"Folder Path: \"%s\"\n", folderPath);
	}
#endif

#if 0
	for (int i = 0; i < 1; i++)
	{
		CString errorMessage;
		DWORD dwErrorCode = ERROR_REQ_NOT_ACCEP;
		//dwErrorCode = i;
		BOOL ret = GetErrorMessage(dwErrorCode, errorMessage);
		CString errTextA = FileUtils::GetLastErrorAsString();
		CString errTextB = FileUtils::GetLastErrorAsStringW();

		TRACE(L"%s", errTextA);
		TRACE(L"%s", errTextB);
		TRACE(L"%s", errorMessage);
	}
#endif

#if 0
	void TokenizerTest();
	TokenizerTest();
	int deb = 1;
#endif

	//JsonTest();

	//bool ret = TextUtilities::TestAll();

	//_CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF | _CRTDBG_CHECK_ALWAYS_DF);
	//SetNoOaCache();

#if 0

	CStringA wordEncoded = "=?UTF-8?Q?St=c3=a9phane_Scudeller?= <sscudeller@gmail.com>";

	CStringA inStrUTF8 = "St=c3=a9phane_Scudeller?= <sscudeller@gmail.com>";  // UTF-8 encoded string


	wordEncoded = "=?UTF-8?B?5rWL6K+V5Lit5paH5paH5Lu2LnR4dA==?=";
	CStringA inStrUTF8_base64 = "5rWL6K+V5Lit5paH5paH5Lu2LnR4dA";


	CStringA inStrUTF8_2;
	const char* bodyBegin = (LPCSTR)inStrUTF8_base64;
	int bodyLength = inStrUTF8_base64.GetAllocLength();

	MboxCMimeCodeBase64 d64(bodyBegin, bodyLength);
	int dlength = d64.GetOutputLength();
	int needLength = dlength;

	SimpleString outbuf;
	outbuf.Resize(needLength);

	char* outptr = outbuf.Data();
	int retlen = d64.GetOutput((unsigned char*)outptr, dlength);
	if (retlen > 0)
		outbuf.SetCount(retlen);

	inStrUTF8_2.Append(outbuf.Data(), outbuf.Count());

	CStringA outStr;
	UINT charsetId = CP_UTF8;
	DWORD error = 0;
	int encodeCnt = TextUtilsEx::EncodeString(inStrUTF8, outStr, charsetId, error);


	CStringA outStr2;
	int encodeCnt2 = TextUtilsEx::EncodeString(inStrUTF8_2, outStr2, charsetId, error);


	CStringA encodedTxt_B;
	int encodeType_B = 'B';
	int eCnt_B = TextUtilsEx::WordEncode(inStrUTF8, encodedTxt_B, encodeType_B);

	CStringA encodedTxt_Q;
	int encodeType_Q = 'Q';
	int eCnt_Q = TextUtilsEx::WordEncode(inStrUTF8, encodedTxt_Q, encodeType_Q);


	CStringA decodedTxt;
	CStringA charset;
	UINT dCharsetId = 0;
	int dCnt = TextUtilsEx::DecodeString(wordEncoded, decodedTxt, charset, dCharsetId, error);


	CStringW wstr;
	BOOL retU2W = TextUtilsEx::UTF82WStr(&decodedTxt, &wstr, error);


	CString fileBaseName;
	CString fileNameExtention;
	//FileUtils::GetFileBaseNameAndExtension(decodedTxt, fileBaseName, fileNameExtention);

	decodedTxt.Append(":?>");


	CStringA encodedTxt;
	int encodeType = 'B';
	int eCnt = TextUtilsEx::WordEncode(decodedTxt, encodedTxt, encodeType);



	int debW = 1;
#endif

#ifdef  IHASH_MAP_TEST
	bool IHashMapTest();

	bool ret = IHashMapTest();
	exit(8);
#endif

#if 0
	BOOL CharEncodingTest();

	BOOL ret = CharEncodingTest();
	//ExitInstance();
	int deb = 1;
#endif

#endif  // end of if _DEBUG

	//FileUtils FU;  FU.UnitTest();
	int deb_exit = 1;
}

/////////////////////////////////////////////////////////////////////////////
// The one and only CmboxviewApp object

CmboxviewApp theApp;

void SetBrowserEmulation()
{
	CString processPath;
	CmboxviewApp::GetProcessPath(processPath);

	CString procName;
	FileUtils::GetFileName(processPath, procName);

	if (procName.IsEmpty())
	{
		procName = L"mboxview.exe";
	}

	CString procName64 = L"mboxview64.exe";

	CString ieSvcVer = CProfile::_GetProfileString(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Internet Explorer", L"svcVersion");
	CString ieVer = CProfile::_GetProfileString(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Internet Explorer", L"Version");
	int p = ieVer.Find(L'.');
	int maj = _tstoi(ieVer);
	int min = _tstoi(ieVer.Mid(p + 1));
	DWORD dwEmulation64 = CProfile::_GetProfileInt(HKEY_LOCAL_MACHINE, L"Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", procName64);
	DWORD dwEmulation = CProfile::_GetProfileInt(HKEY_LOCAL_MACHINE, L"Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", procName);
	if (maj == 9 && min >= 10) {
		DWORD dwem = min * 1000 + 1;
		//if (dwEmulation != dwem)  // TODO: check if that was just for testing ???
		{
			BOOL ret1 = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, L"Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", L"mboxview.exe", dwem);
			BOOL ret2 = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, L"Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", L"mboxview64.exe", dwem);
			BOOL ret3 = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, L"Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", procName, dwem);
			BOOL ret4 = CProfile::_WriteProfileInt(HKEY_LOCAL_MACHINE, L"Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", procName, dwem);
			const int deb = 1;
		}
	}
	else
	{
		if (maj == 9) {
			if (dwEmulation != 9999) {
				CProfile::_WriteProfileInt(HKEY_CURRENT_USER, L"Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", procName, 9999);
				CProfile::_WriteProfileInt(HKEY_LOCAL_MACHINE, L"Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", procName, 9999);
			}
		}
		else
			if (maj == 8) {
				if (dwEmulation != 8888) {
					CProfile::_WriteProfileInt(HKEY_CURRENT_USER, L"Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", procName, 8888);
					CProfile::_WriteProfileInt(HKEY_LOCAL_MACHINE, L"Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", procName, 8888);
				}
			}
	}
}

#pragma comment(linker, "/defaultlib:version.lib")

BOOL CmboxviewApp::GetMboxviewLongVersion(CString& version)
{
	DWORD ms, ls;
	if (CmboxviewApp::GetFileVersionInfo((HMODULE)AfxGetInstanceHandle(), ms, ls))
	{
		CmboxviewApp::m_versionMS = ms;
		CmboxviewApp::m_versionLS = ls;

		CString txt;
		// display file version from VS_FIXEDFILEINFO struct
		txt.Format(L"Version %d.%d.%d.%d",
			HIWORD(ms), LOWORD(ms),
			HIWORD(ls), LOWORD(ls));

		if (sizeof(INT_PTR) == 4)
			txt.Append(L" 32bit");
		else if (sizeof(INT_PTR) == 8)
			txt.Append(L" 64bit");
		txt.Append(L" Unicode");
		version = txt;
		return TRUE;
	}
	return FALSE;
}

BOOL CmboxviewApp::GetFileVersionInfo(HMODULE hModule, DWORD& ms, DWORD& ls)
{
	// get module handle
	wchar_t filename[_MAX_PATH];

	// get module file name
	DWORD len = GetModuleFileName(hModule, filename,
		sizeof(filename) / sizeof(filename[0]));
	if (len <= 0)
		return FALSE;

	// read file version info
	DWORD dwDummyHandle; // will always be set to zero
	len = GetFileVersionInfoSize(filename, &dwDummyHandle);
	if (len <= 0)
		return FALSE;

	BYTE* pVersionInfo = new BYTE[len]; // allocate version info
	if (pVersionInfo == NULL)
		return FALSE;
	if (!::GetFileVersionInfo(filename, 0, len, pVersionInfo)) {
		delete[] pVersionInfo;
		return FALSE;
	}

	LPVOID lpvi;
	UINT iLen;
	if (!VerQueryValue(pVersionInfo, L"\\", &lpvi, &iLen)) {
		delete[] pVersionInfo;
		return FALSE;
	}

	ms = ((VS_FIXEDFILEINFO*)lpvi)->dwFileVersionMS;
	ls = ((VS_FIXEDFILEINFO*)lpvi)->dwFileVersionLS;
	BOOL res = (((VS_FIXEDFILEINFO*)lpvi)->dwSignature == VS_FFI_SIGNATURE);

	delete[] pVersionInfo;

	return res;
}

/////////////////////////////////////////////////////////////////////////////
// CmboxviewApp message handlers

int CmboxviewApp::ExitInstance()
{

	MboxMail::ReleaseResources();
	ResHelper::ReleaseResources();

	return CWinApp::ExitInstance();
}

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg(CWnd* pParent = NULL);

	CFont m_linkFont;
	// Dialog Data
		//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	virtual INT_PTR DoModal();
	CWnd* m_pParent;

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnStnClickedDonation();
	CLinkCursor m_link;
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnClose();
};

CAboutDlg::CAboutDlg(CWnd* pParent)
DIALOG_FROM_TEMPLATE(: CDialogEx(CAboutDlg::IDD, pParent))
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT

	m_pParent = pParent;
}

INT_PTR CAboutDlg::DoModal()
{
#ifdef _DIALOG_FROM_TEMPLATE
	INT_PTR ret = CMainFrame::SetTemplate(this, IDD_ABOUTBOX, m_pParent);
#else
	INT_PTR ret = CDialogEx::DoModal();
#endif
	return ret;
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
	DDX_Control(pDX, IDC_DONATION, m_link);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
	//{{AFX_MSG_MAP(CAboutDlg)
	//}}AFX_MSG_MAP
	ON_STN_CLICKED(IDC_DONATION, &CAboutDlg::OnStnClickedDonation)
	ON_WM_CTLCOLOR()
	ON_WM_CLOSE()
END_MESSAGE_MAP()

// App command to run the dialog
void CmboxviewApp::OnAppAbout()
{
	CAboutDlg aboutDlg;
	aboutDlg.DoModal();
}

void CmboxviewApp::OnHelpDonate()
{
	HINSTANCE result = ShellExecute(NULL, L"open", L"https://sourceforge.net/p/mbox-viewer/donate/", NULL, NULL, SW_SHOWNORMAL);
}

/////////////////////////////////////////////////////////////////////////////
// CmboxviewApp message handlers

BOOL CAboutDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	CString version;
	BOOL retGetVersion = CmboxviewApp::GetMboxviewLongVersion(version);
	if (retGetVersion)
	{
		GetDlgItem(IDC_STATIC1)->SetWindowText(version);
	}

	ResHelper::LoadDialogItemsInfo(this);
	ResHelper::UpdateDialogItemsInfo(this);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

class CCmdLine : public CCommandLineInfo
{
public:
	CString m_commandLineOptions;
	BOOL m_hasOptions;
	BOOL m_bOnlyVerifyHasOptions;
	BOOL m_bError;
	BOOL m_bLastPathSet;
	CString m_lastParam;

	CCmdLine::CCmdLine(BOOL bOnlyVerifyHasOptions = FALSE) {
		m_bError = FALSE; m_bLastPathSet = FALSE; m_hasOptions = FALSE;
		m_bOnlyVerifyHasOptions = bOnlyVerifyHasOptions;
	}
	CCmdLine::CCmdLine(CString & commandLineOptions, BOOL bOnlyVerifyHasOptions = FALSE)
	{
		m_commandLineOptions = commandLineOptions;
		m_bError = FALSE; m_bLastPathSet = FALSE; m_hasOptions = FALSE;
		m_bOnlyVerifyHasOptions = bOnlyVerifyHasOptions;
	}
	void ParseParam(LPCWSTR lpszParam, BOOL bFlag, BOOL bLast);
	void ParseCommandLine();
};

void CCmdLine::ParseCommandLine()
{
	CStringArray ar;
	CString delims = L" \t";
	//TextUtilsEx::SplitStringW(m_commandLineOptions, delims, ar);

	//extUtilsEx::TokenizeW(m_commandLineOptions, ar, L' ');

	int nMaxTokens = 100;
	BOOL bTrimToken = TRUE;
	BOOL bEnableEscapedChars = FALSE;
	BOOL bEnableDoubleQuote = TRUE;
	BOOL bReturnEmptyToken = FALSE;
	int cnt = TextUtilsEx::XTokenString(m_commandLineOptions, delims, ar, nMaxTokens, bTrimToken, bEnableEscapedChars, bEnableDoubleQuote, bReturnEmptyToken);

	int argc = (int)ar.GetCount();
	for (int i = 0; i < argc; i++)
	{
		LPCTSTR pszParam = ar[i];
		BOOL bFlag = FALSE;
		BOOL bLast = ((i + 1) == argc);
		if (pszParam[0] == '-' || pszParam[0] == '/')
		{
			// remove flag specifier
			bFlag = TRUE;
			++pszParam;
		}
		this->ParseParam(pszParam, bFlag, bLast);
	}
};

void CCmdLine::ParseParam(LPCWSTR lpszParam, BOOL bFlag, BOOL bLast) // bLast )
{
	HWND h = NULL; // we don't have any window yet  
	//CString ptxt = CString(L"ParseParam() must delete this line:  ") + lpszParam;
	//int answer = ::MessageBox(h, ptxt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);

	CString section_general = CString(sz_Software_mboxview) + L"\\General";
	CString section_options = CString(sz_Software_mboxview) + L"\\Options";

	if (m_bError)
		return;
	if (bFlag)
	{
		m_hasOptions = TRUE;
		CMainFrame::m_commandLineParms.m_hasOptions = TRUE;
	}
	if (m_bOnlyVerifyHasOptions)
		return;

	if (bLast)
	{
		m_lastParam = lpszParam;
	}
	else if (bFlag)
	{
		CMainFrame::m_commandLineParms.m_hasOptions = TRUE;  // avoid this; redo
		if (_tcsncmp(lpszParam, L"FOLDER=", 7) == 0)
		{
			CString openFolder = lpszParam + 7;
			CMainFrame::m_commandLineParms.m_mboxFolderPath = openFolder;
		}
		else if (_tcsncmp(lpszParam, L"MAIL_FILE=", 10) == 0)
		{
			CString mailFile = lpszParam + 10;

			CMainFrame::m_commandLineParms.m_mboxFileNameOrPath = mailFile;
		}
		else if (_tcsncmp(lpszParam, L"EXPORT_EML=", 11) == 0)
		{
			CString exportEML = lpszParam + 11;
			exportEML.MakeLower();
			if (!((exportEML.Compare(L"y") == 0) || (exportEML.Compare(L"n") == 0)))
			{
#if 0
				CString txt = L"Invalid Command Line Option Value \"";
				CString opt = lpszParam;
				txt += opt + L"\". Valid are \"y|n\". Note that once defined valid EXPORT_EML persists in the registry.\nDo you want to continue?";
#endif
				CString txt;
				CString opt = lpszParam;
				CString fmt = L"Invalid Command Line Option Value \"%s\". "
					"Valid are \"y|n\". Note that once defined valid EXPORT_EML persists in the registry.\nDo you want to continue?";

				ResHelper::TranslateString(fmt);
				txt.Format(fmt, lpszParam);

				HWND h = NULL; // we don't have any window yet  
				int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
				if (answer == IDNO)
					m_bError = TRUE;
			}
			else {
				CString section_options = CString(sz_Software_mboxview) + L"\\Options";
				CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_general, L"exportEML", exportEML);

				CMainFrame::m_commandLineParms.m_exportEml = FALSE;
				if (exportEML.Compare(L"y") == 0)
					CMainFrame::m_commandLineParms.m_exportEml = TRUE;
			}
		}
		else if (_tcsncmp(lpszParam, L"PROGRESS_BAR_DELAY=", 19) == 0)
		{
			CString barDelay = lpszParam + 19;
			// Validate
			if (!TextUtilsEx::isNumeric(barDelay))
			{
#if 0
				CString txt = L"Invalid Command Line Option Value \"";
				CString opt = lpszParam;
				txt += opt + L"\".\nDo you want to continue?";
#endif

				CString txt;
				CString opt = lpszParam;
				CString fmt = L"Invalid Command Line Option Value \"%s\".\nDo you want to continue?";
				ResHelper::TranslateString(fmt);
				txt.Format(fmt, lpszParam);

				HWND h = NULL; // we don't have any window yet  
				int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
				if (answer == IDNO)
					m_bError = TRUE;
			}
			else
			{
				DWORD progressBarDelay = _tstoi(barDelay);
				CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"progressBarDelay", progressBarDelay);
				CMainFrame::m_commandLineParms.m_progressBarDelay = progressBarDelay;
			}
		}
		else if (_tcsncmp(lpszParam, L"EML_PREVIEW_MODE", 19) == 0)
		{
			CMainFrame::m_commandLineParms.m_bEmlPreviewMode = TRUE;
		}
		else if (_tcsncmp(lpszParam, L"MBOX_MERGE_LIST_FILE", 20) == 0)
		{
			CString emlListFile = lpszParam + 21;
			CMainFrame::m_commandLineParms.m_mboxListFilePath = emlListFile;
		}
		else if (_tcsncmp(lpszParam, L"MBOX_MERGE_TO_FILE", 18) == 0)
		{
			CString mergeToFilePath = lpszParam + 19;
			FileUtils::NormalizeFilePath(mergeToFilePath);
			CMainFrame::m_commandLineParms.m_mergeToFilePath = mergeToFilePath;
		}
		else if (_tcsncmp(lpszParam, L"TRACE_CASE", 10) == 0)
		{
			CString traceCase = lpszParam + 11;
			DWORD nTraceCase = _tstoi(traceCase);
			CMainFrame::m_commandLineParms.m_traceCase = nTraceCase;
		}
		else
		{
			// Unknown argument
			CMainFrame::m_commandLineParms.m_hasOptions = FALSE;
			CString opt = lpszParam;
			CString txt;
			CString fmt = L"Invalid Command Line Option \"%s\""
				L"\n\nValid options:\n"
				L"  -FOLDER=Folder Path\n"
				L"  -MAIL_FILE=Mbox File Path or Name to open\n"
				L"  -EXPORT_EML=y or n\n"
				L"  -PROGRESS_BAR_DELAY=Seconds\n"
				L"  -EML_PREVIEW_MODE\n"
				L"  -MBOX_MERGE_LIST_FILE=Path to File containing list of mbox files to merge\n"
				L"  -MBOX_MERGE_TO_FILE=Path to File to save merge results\n"
				L"\nDo you want to continue?"
				;

			ResHelper::TranslateString(fmt);
			txt.Format(fmt, lpszParam);

			CMainFrame::m_commandLineParms.Clear();
			HWND h = NULL; // we don't have any window yet
			int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_YESNO);
			if (answer == IDNO)
				m_bError = TRUE;
		}
	}
}
/////////////////////////////////////////////////////////////////////////////
// CmboxviewApp initialization

#if 0
static int tot = 0;
void* __cdecl malloc(size_t size)
{
	int deb = 1;
	char mem[1];
	void* p = new char[size];
	tot += size;
	return p;
}
#endif


#include "afxsock.h"

BOOL CmboxviewApp::GetProcessPath(CString& procressPath)
{
	HMODULE hModule = 0;
	wchar_t procFileName[512];
	procFileName[0] = L'\0';
	DWORD   nSize = 511;
	DWORD retval = GetModuleFileName(hModule, procFileName, nSize);
	if (retval == 0)
	{
		DWORD error = GetLastError();
		CString errorText = FileUtils::GetLastErrorAsString();
		TRACE(L"GetProcessPath: error=%d errorText=%s\n", error, errorText);
		int deb = 1;
	}
	CString processFilePath = procFileName;

	procressPath.Empty();
	wchar_t* pValue;
	errno_t  er = _get_wpgmptr(&pValue);
	if ((er == 0) && pValue)
	{
		procressPath.Append(pValue);

		CString txt;
		txt.Format(L"\nGetModuleFileName=\"%s\"\n_get_wpgmptr     =\"%s\"\n", processFilePath, procressPath);
		TRACE(L"%s", txt);

#ifdef _DEBUG
#if 0
		HWND h = 0;
		int answer = MessageBox(h, txt, L"Info", MB_APPLMODAL | MB_OK);
#endif
#endif

		return TRUE;
	}
	else
		return FALSE;
}


BOOL CmboxviewApp::InitInstance()
{
	CString errorText;

	ResHelper::MyTrace(L"CmboxviewApp::InitInstance\n");


	CString allCommanLineOptions(CWinApp::m_lpCmdLine);
	BOOL bOnlyVerifyHasOptions = TRUE;
	CCmdLine vcmdInfo(bOnlyVerifyHasOptions);

	ParseCommandLine(vcmdInfo);
	if (vcmdInfo.m_bError)
	{
		MboxMail::ReleaseResources();
		return FALSE;
	}
	else if (!vcmdInfo.m_hasOptions && !allCommanLineOptions.IsEmpty())
	{
		BOOL retSetup = CProfile::SetupPreviewConfigurationFile(errorText);
		if (retSetup == FALSE)
		{
			HWND h = NULL; // we don't have any window yet 
			int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);

			MboxMail::ReleaseResources();
			return FALSE;
		}
		CMainFrame::m_commandLineParms.m_bEmlPreviewMode = TRUE;
	}
	else
	{

		BOOL retcheck = CProfile::DetermineConfigurationType(errorText);
		if (retcheck == FALSE)
		{
			HWND h = NULL; // we don't have any window yet 
			int answer = ::MessageBox(h, errorText, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);

			// To stop memory leaks reports by debugger
			MboxMail::ReleaseResources(FALSE);
			return FALSE;
		}
	}


	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	BOOL retExists = TRUE;

	//ConfigTree* config = 0;
	if (!CProfile::IsRegistryConfig())
	{
		DWORD attr = GetFileAttributes(CmboxviewApp::m_configFilePath);

		if (attr & FILE_ATTRIBUTE_READONLY)
		{
			//HWND h = NULL; // we don't have any window yet 
			//HWND h = GetSafeHwnd();
			HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
			CString txt;
			CString fmt = L"Could not write to MBox Viewer configuration file:\n\n%s"
				L"\n\nlikely  due to lack of permission to write.\n"
				L"Please resolve the issue and run the MBox Viewer again\n\n"
				L"Or configure MBox Viewer to use Windows Registry.\n"
				;

			ResHelper::TranslateString(fmt);
			txt.Format(fmt, CmboxviewApp::m_configFilePath);

			int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);

			// To stop memory leaks reports by debugger
			MboxMail::ReleaseResources(FALSE);
			return FALSE;
		}

		ConfigTree* confTree = CProfile::GetConfigTree();
		_ASSERTE(confTree);

		int retload = confTree->LoadConfigFromFile();

		if (retload >= 0)
			CmboxviewApp::m_configFileLoaded = TRUE;

#if 0
		CString configFilePath = CmboxviewApp::m_configFilePath + L"-test";
		confTree->Dump2File(configFilePath);
#endif
	}
	else
	{
		// Check if user can read/write Windows Registry

		CString key = L"MagicNumber";
		int rnumber = rand();

		HKEY hk = HKEY_CURRENT_USER;
		//hk = HKEY_LOCAL_MACHINE;  // enable for testing
		BOOL retwrite = CProfile::_WriteProfileInt(hk, section_general, key, rnumber);

		int number = 0;
		BOOL retget = CProfile::_GetProfileInt(hk, section_general, key, number);

		BOOL ret = CProfile::_DeleteValue(hk, section_general, key);

		if ((retwrite == FALSE) || (retget == FALSE) || (rnumber != number))
		{
			HWND h = NULL; // we don't have any window yet 
			CString txt = L"Could not write to Windows registry likely  due to lack of permission. "
				L"Please resolve the issue and run the MBox Viewer again.\n\n"
				L"Or configure MBox Viewer to use configuration file.\n"
				;

			ResHelper::TranslateString(txt);

			int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);

			// To stop memory leaks reports by debugger
			MboxMail::ReleaseResources(FALSE);
			return FALSE;
		}
	}

	MboxMail::m_HintConfig.LoadHintBitmap();

	CString languageFolderPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"languageFolderPath");
	if (!languageFolderPath.IsEmpty())
	{
		ResHelper::MyTrace(L"CmboxviewApp::InitInstance: %s\n", languageFolderPath);

		CString FolderPath = languageFolderPath;
		FolderPath.TrimRight(L"\\");

		CString folderName;
		FileUtils::CPathStripPath((LPCWSTR)FolderPath, folderName);

		CString translationFileName = folderName + L".txt";
		CString translationFileNamePath = languageFolderPath + L"\\" + translationFileName;

		ResHelper::EnableLanguageLoading();
		ResHelper::LoadLanguageMap(translationFileNamePath);
	}
	else
	{
#ifndef _DEBUG
		ResHelper::DisableLanguageLoading();
#endif
int deb = 1;
	}

	CString dataFolder;
	BOOL ret = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"dataFolder", dataFolder);
	if ((ret == FALSE) || dataFolder.IsEmpty())
	{
		if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode == FALSE)
			NTreeView::ImportLegacyRegistryData();
		int deb = 1;
	}

	HANDLE procHandle = GetCurrentProcess();
	BOOL priClass = SetPriorityClass(procHandle, ABOVE_NORMAL_PRIORITY_CLASS);
	DWORD err = GetLastError();

#ifdef USE_STACK_WALKER
	int options =
		StackWalker::RetrieveSymbol      // Try to get the symbol-name
		| StackWalker::RetrieveLine        // Try to get the line for this symbol
		//| StackWalker::RetrieveModuleInfo  // Try to retrieve the module-infos
		//| StackWalker::RetrieveFileVersion // Also retrieve the version for the DLL/EXE
		| StackWalker::SymBuildPath        // Generate a "good" symbol-search-path
		//| StackWalker::SymUseSymSrv        // Also use the public Microsoft-Symbol-Server
		;

	MboxMail::glStackWalker = new MyStackWalker(options);
	MyStackWalker* sw = MboxMail::glStackWalker;
	if (sw)
	{
		// preload modules, etc to minimize work in MyUnhandledExceptionFilter();
		sw->LoadAllModules();
		sw->ClearBuffer();
	}

#if 0
	char* stackDumpFileName = "UnhandledException_StackDump.txt";
	UINT seNumb = 0;
	CONTEXT* ContextRecord = nullptr;
	int mailPosition = -1;
	BOOL ret = DumpStack(sw, stackDumpFileName, (CHAR*)"Preload", seNumb, ContextRecord, mailPosition);
#endif

	/// Disable "Whole program Optimization" to get accurate Stack Trace  !!!!
	SetMyExceptionHandler();
#endif

	//AfxEnableMemoryTracking(TRUE);
	//afxMemDF = allocMemDF | delayFreeMemDF | checkAlwaysMemDF;


	CString mailFile;
	CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_general, L"mailFile", mailFile);

	CCmdLine cmdInfo(allCommanLineOptions);
	CMainFrame::m_commandLineParms.m_allCommanLineOptions = allCommanLineOptions;
	ParseCommandLine(cmdInfo);
	if (cmdInfo.m_bError)
	{
		MboxMail::ReleaseResources();
		return FALSE;
	}
	else if (!CMainFrame::m_commandLineParms.m_hasOptions && !CMainFrame::m_commandLineParms.m_allCommanLineOptions.IsEmpty())
	{
		CMainFrame::m_commandLineParms.m_bEmlPreviewMode = TRUE;  // m_bDirectFileOpenMode handling is now the same as m_bEmlPreviewMode
		//CMainFrame::m_commandLineParms.m_bDirectFileOpenMode = TRUE;
		CMainFrame::m_commandLineParms.m_mboxFileNameOrPath = CMainFrame::m_commandLineParms.m_allCommanLineOptions;
		CMainFrame::m_commandLineParms.m_mboxFileNameOrPath.Trim(L"\"");
		CMainFrame::m_commandLineParms.m_mboxFileNameOrPath.TrimRight(L"\\");

		CString mailFile = CMainFrame::m_commandLineParms.m_mboxFileNameOrPath;
		CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_general, L"mailFile", mailFile);
	}
	else
	{
		// Do additional Check of command line options
		int retCode = CMainFrame::m_commandLineParms.VerifyParameters();
		if (retCode == -1)  // error and don't continue
		{
			MboxMail::ReleaseResources();
			return FALSE;
		}
		else if (retCode == -2)  // error but continue; probably overkill  // no longer true option
		{
			CMainFrame::m_commandLineParms.Clear();
		}
	}

	CString processPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"processPath");
#if 0
	// Commented out for now since it seem to confuse users. Need to find better solution.
	if (!processPath.IsEmpty())
	{
		CString txt = L"Mbox viewer instance might be running already:\n\n" + processPath;
		txt += L"\n\n";
		txt += L"Only single instance should be running to avoid potential\n";
		txt += L"issues since all instances will share the same data in the registry.\n\n";
		txt += L"In most cases this warning can be ignored.\n"
			"It will be generated if program crashes or it is killed from Task Manager.\n\n";
		txt += L"Do you want to continue?";
		HWND h = NULL; // we don't have any window yet  
		int answer = MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
		{
			MboxMail::ReleaseResources();
			return FALSE;
		}
	}
#endif

	GetProcessPath(m_processPath);

	CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_general, L"processPath", m_processPath);

	AfxEnableControlContainer();

	// Standard initialization
	// If you are not using these features and wish to reduce the size
	//  of your final executable, you should remove from the following
	//  the specific initialization routines you do not need.
	SetBrowserEmulation();

#ifdef _AFXDLL
	Enable3dControls();			// Call this when using MFC in a shared DLL
#else
	//Enable3dControlsStatic();	// Call this when linking to MFC statically
#endif

	// Change the registry key under which our settings are stored.
	// TODO: You should modify this string to be something appropriate
	// such as the name of your company or organization.
	SetRegistryKey(L"UMBoxViewer");  // FIXMEFIXME

	if (CProfile::IsRegistryConfig())
		m_pRecentFileList = new CRecentFileList(0, L"MRUs", L"Path %d", 16);
	else
		m_pRecentFileList = new MyCRecentFileList(0, L"MRUs", L"Path %d", 16);

	m_pRecentFileList->ReadList();
	// Initialize all Managers for usage. They are automatically constructed
	// if not yet present
	// To create the main window, this code creates a new frame window
	// object and then sets it as the application's main window object.

	g_tu.Init();

	DWORD msgViewPosition = 1;
	BOOL retval;
	retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_general, L"messageWindowPosition", msgViewPosition);
	if (retval == TRUE) {
		if ((msgViewPosition < 1) && (msgViewPosition > 3))
			msgViewPosition = 1;  // bottom=1
	}

	CString section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacement";
#if 0
	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
		section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacementPreview";
	else if (CMainFrame::m_commandLineParms.m_bDirectFileOpenMode)
		section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacementDirect";
#endif

	retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_wnd, L"messageWindowPosition", msgViewPosition);
	if (retval == TRUE) {
		if ((msgViewPosition < 1) && (msgViewPosition > 3))
			msgViewPosition = 1;  // bottom=1
	}

	CMainFrame* pFrame = new CMainFrame(msgViewPosition);

	if (pFrame == 0)
	{
		int deb = 1;
	}

	DWORD BUFSIZE = 1024;
	wchar_t Buffer[1024];


	DWORD dwRet;
#if 0
	dwRet = ::GetCurrentDirectory(BUFSIZE, Buffer);
	if (dwRet > 0)
	{
		m_startupPath = CString(Buffer);
	}
#endif

	CString folderPath;
	FileUtils::GetFolderPath(m_processPath, folderPath);

	m_startupPath = folderPath;

	if (!::SetCurrentDirectory((LPCWSTR)folderPath))
	{
		int deb = 1;
	}

	dwRet = ::GetCurrentDirectory(BUFSIZE, Buffer);
	if (dwRet > 0)
	{
		m_currentPath = CString(Buffer);
	}

	pFrame->m_processPath = m_processPath;
	pFrame->m_startupPath = m_startupPath;
	pFrame->m_currentPath = m_currentPath;

	if (pFrame->GetSafeHwnd() == 0)
		int deb = 1;
	m_pMainWnd = pFrame;

	// create and load the frame with its resources

	BOOL retLOad = pFrame->LoadFrame(IDR_MAINFRAME,
		WS_OVERLAPPEDWINDOW | FWS_ADDTOTITLE, NULL,
		NULL);

	// The one and only window has been initialized, so show and update it.
	pFrame->ShowWindow(SW_SHOW);
	pFrame->UpdateWindow();

	DWORD ms, ls;
	if (GetFileVersionInfo((HMODULE)AfxGetInstanceHandle(), ms, ls))
	{
		CmboxviewApp::m_versionMS = ms;
		CmboxviewApp::m_versionLS = ls;

		CString savedPlatform;
		CString newPlatform;
		CString newVer;
		// display file version from VS_FIXEDFILEINFO struct
		newVer.Format(L"%d.%d.%d.%d",
			HIWORD(ms), LOWORD(ms),
			HIWORD(ls), LOWORD(ls));

		if (sizeof(INT_PTR) == 4)
			newPlatform.Append(L"32bit");
		else if (sizeof(INT_PTR) == 8)
			newPlatform.Append(L"64bit");

		savedPlatform = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"platform");
		m_savedVer = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"version");

		if (m_savedVer.Compare(newVer) != 0)
		{
			CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_general, L"version", newVer);
			CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_general, L"platform", newPlatform);
			pFrame->PostMessage(WM_COMMAND, ID_APP_ABOUT, 0);
		}
		else if (savedPlatform.Compare(newPlatform) != 0)
		{
			CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_general, L"platform", newPlatform);
			pFrame->PostMessage(WM_COMMAND, ID_APP_ABOUT, 0);
		}
	}

	MboxMail::LoadHintBitmap();

	if (AfxSocketInit() == FALSE)
	{
		AfxMessageBox(L"Sockets Could Not Be Initialized");
		return FALSE;
	}

	TRACE(L"LoadMenuItemsInfo\n");
	CMenu* menu = pFrame->GetMenu();
	int index = 1;
	ResHelper::LoadMenuItemsInfo(menu, index);
	ResHelper::UpdateMenuItemsInfo(menu, index);

	pFrame->SetNewMenu();

	TRACE(L"LoadDialogItemsInfo\n");
	CDialogBar& dbar = pFrame->GetDialogBar();;
	HWND h = dbar.GetSafeHwnd();
	ResHelper::LoadDialogItemsInfo(&dbar);
	ResHelper::UpdateDialogItemsInfo(&dbar);

	TRACE(L"LoadToolBarItemsInfo\n");
	CToolBar& cbar = pFrame->GetToolBar();
	ResHelper::LoadToolBarItemsInfo(&cbar);
	ResHelper::UpdateToolBarItemsInfo(&cbar);

	TRACE(L"Load ResInfo END\n");

	int deb2 = 1;

	if (pFrame)
	{
		pFrame->UpdateWindow();
		pFrame->DrawMenuBar();

		pFrame->m_createCompleted = TRUE;
	}

	return TRUE;
}

void CAboutDlg::OnStnClickedDonation()
{
	HINSTANCE result = ShellExecute(NULL, L"open", L"https://sourceforge.net/p/mbox-viewer/donate/", NULL, NULL, SW_SHOWNORMAL);
}

HBRUSH CAboutDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);
	if (nCtlColor == CTLCOLOR_STATIC &&
		pWnd->GetSafeHwnd() == m_link.GetSafeHwnd()
		) {
		pDC->SetTextColor(RGB(0, 0, 200));
		if (m_linkFont.m_hObject == NULL) {
			LOGFONT lf;
			CFont *aboutFont = GetFont();
			if (aboutFont)
				aboutFont->GetObject(sizeof(lf), &lf);
			lf.lfUnderline = TRUE;
			m_linkFont.CreateFontIndirect(&lf);
		}
		pDC->SelectObject(&m_linkFont);
	}
	return hbr;
}


void CAboutDlg::OnClose()
{
	// TODO: Add your message handler code here and/or call default

	CDialog::OnClose();
}


#ifdef IHASH_MAP_TEST

#include <string>

int next_fake_prime(int L);

bool IHashMapTest()
{
	class Item
	{
	public:
		Item(std::string& nm) { name = nm; }
		//
		dlink_node<Item> hashNode;
		dlink_node<Item> hashNode2;
		std::string name;
	};

	struct ItemHash {
		unsigned long operator()(const Item* key) const
		{
			hashsum_t  hashsum = std::hash<std::string>{}(key->name);
			return hashsum;
		};
		unsigned long operator()(Item* key) const
		{
			hashsum_t  hashsum = std::hash<std::string>{}(key->name);
			return hashsum;
		};
	};

	struct ItemEqual {
		bool operator()(const Item* key, const Item* item) const
		{
			return (key->name == item->name);
		};
		bool operator()(Item* key, const Item* item) const
		{
			return (key->name == item->name);
		};
	};

	using MyHashMap = IHashMap<Item, Item, ItemHash, ItemEqual, &Item::hashNode>;
	using MyHashMap2 = IHashMap<Item, Item, ItemHash, ItemEqual, &Item::hashNode2>;

	unsigned long sz = 10;

	Item* it = new Item(std::string("John"));
	Item* it2 = new Item(std::string("Irena"));
	Item* it3 = new Item(std::string("Lucyna"));
	Item* it4 = new Item(std::string("Jan"));

	MyHashMap hashMap(sz);
	MyHashMap hashMap2(sz);

	/// Insert  it3 and it4 to second hash map hashMap2
	if (hashMap2.find(it3) == nullptr)
		hashMap2.insert(it3, it3);

	if (hashMap2.find(it4) == nullptr)
		hashMap2.insert(it4, it4);

	hashMap2.clear();
	assert(hashMap2.count() == 0);

	/// Insert it, it2 and it3
	if (hashMap.find(it) == nullptr)
		hashMap.insert(it, it);
	assert(hashMap.count() == 1);

	if (hashMap.find(it2) == nullptr)
		hashMap.insert(it2, it2);
	assert(hashMap.count() == 2);

	if (hashMap.find(it3) == nullptr)
		hashMap.insert(it3, it3);
	assert(hashMap.count() == 3);


	// Iterate and print item names
	MyHashMap::IHashMapIter iter = hashMap.first();
	for (; !hashMap.last(iter); hashMap.next(iter))
	{
		TRACE("%s\n", iter.element->name.c_str());
	}

	// Remove all items by Key one by one
	hashMap.remove(it);
	assert(hashMap.count() == 2);

	hashMap.remove(it2);
	assert(hashMap.count() == 1);

	hashMap.remove(it3);
	assert(hashMap.count() == 0);

	///Insert it, it 2 and i3 again to the same hash map
	if (hashMap.find(it) == nullptr)
		hashMap.insert(it, it);
	assert(hashMap.count() == 1);

	if (hashMap.find(it2) == nullptr)
		hashMap.insert(it2, it2);
	assert(hashMap.count() == 2);

	if (hashMap.find(it3) == nullptr)
		hashMap.insert(it3, it3);
	assert(hashMap.count() == 3);

	// Remove by key and hashsum
	Item* item;
	if ((item = hashMap.find(it2)) != nullptr)
	{
		hashsum_t hashsum = ItemHash{}(it2);
		hashMap.remove(item, hashsum);
	}
	assert(hashMap.count() == 2);

	// Add back
	hashsum_t hashsum = ItemHash{}(it2);
	if (hashMap.find(it2, hashsum) == nullptr)
		hashMap.insert(hashsum, it2);
	assert(hashMap.count() == 3);

	// Iterate and remove selected elements
	iter = hashMap.first();
	for (; !hashMap.last(iter); )
	{
		if (iter.element->name.compare("Lucyna") == 0)
		{
			TRACE("Keep %s\n", iter.element->name.c_str());
			hashMap.next(iter);
		}
		else
		{
			// Remove by Iterator
			TRACE("Remove %s\n", iter.element->name.c_str());
			hashMap.remove(iter);
		}
	}

	hashMap.clear();

	delete it;
	delete it2;
	delete it3;
	delete it4;

	int deb = 1;

	int L = 1;
	int R = 10;
	int p = next_fake_prime(L);
	TRACE("%d,%d=%d\n", L, R, p);
	//
	L = 1000; R = 10;
	p = next_fake_prime(L);
	TRACE("%d,%d=%d\n", L, R, p);
	//
	L = 1197010; R = 100;
	p = next_fake_prime(L);
	TRACE("%d,%d=%d\n", L, R, p);

	return true;
}

int next_fake_prime(int L)
{
	int i, j;
	int R = 10;
	if (L < 3) L = 3;
	int modNumbs = 19;
	if (L <= 19)
		modNumbs = L;
	if (L < 10000)
		R = 10;
	else
		R = 100;

	for (i = L; i <= L + R; i++)
	{
		for (j = 2; j <= modNumbs; ++j)
		{
			if (i % j == 0)
				break;
		}
		if (j == modNumbs)
			break;
	}
	if (i < (L + R))
		return i;
	else
		return L;
}
#endif

#if 0
#include "json/json.h"
#include <iostream>
/** \brief Write a Value object to a string.
 * Example Usage:
 * $g++ stringWrite.cpp -ljsoncpp -std=c++11 -o stringWrite
 * $./stringWrite
 * {
 *     "action" : "run",
 *     "data" :
 *     {
 *         "number" : 1
 *     }
 * }
 */
int JsonTest()
{
	Json::Value root;
	Json::Value data;
	root["action"] = "run";
	data["number"] = 1;
	root["data"] = data;

	Json::StreamWriterBuilder builder;
	const std::string json_file = Json::writeString(builder, root);
	TRACE("XXXXXXXXXXXXXXXXXXXXXXXX\n%s\n", json_file.c_str());

	{
		const std::string rawJson = R"({"Age": 20, "Name": "colin"})";
		const auto rawJsonLength = static_cast<int>(rawJson.length());
		JSONCPP_STRING err;
		Json::Value root;


		Json::CharReaderBuilder builder;
		const std::unique_ptr<Json::CharReader> reader(builder.newCharReader());
		if (!reader->parse(rawJson.c_str(), rawJson.c_str() + rawJsonLength, &root,
			&err)) {
			std::cout << "error" << std::endl;
			return EXIT_FAILURE;
		}

		const std::string name = root["Name"].asString();
		const int age = root["Age"].asInt();

		TRACE("%s\n", name.c_str());
		TRACE("%d\n", age);
	}

	return EXIT_SUCCESS;
}
#endif
#if 0
void TokenizerTest()
{
	CStringArray list;
	CString str;

	str = "\"home \"\"\"\" x\", Home/kitchen \"\" x, \"Inbox\", \"Sent\"";
	list.Add(str);


	str = "Archived,Important,Opened,Category Updates,Osobiste/amex/jan";
	str.Append("\"\"lucyna\"\"");

	str.Append("william,Osobiste/amex/amex,Osobiste/amex/ala/ala,Osobiste/amex/\"\"''' \"\" '''");
	str.Append("' '' ''' \"\" \"\"\"\" \"\"\"\" ''\"\" '''\"\",\"Osobiste/amex/*?<>|:\\//\\\"\"  \\\"\" \'");
	str.Append("\\,\",Osobiste/amex/jan \"\"lucyna robert,Osobiste/amex/a\b\\c a/b/c a\\b\\c");
	str.Append("a;b,\"Osobiste/amex/tata,tata\",Osobiste/amex/mama mama");
	list.Add(str);

	//str = "Osobiste/amex/*?<>|:\\//\""  \\"" \' \,";


	str = "\"my \'\"\' \'\"\', ala/mama,\"alal/ma vs jan, ziggy ero\" , tata, \tlala\",	a \'\"\'  \'\" \'\', \"my \'\"\' \'\"\',\'\"\'";
	//list.Add(str);

	str = "Archived, Important, Starred, \"American Cobalt/American Cobalt vs.Diehl, Hastie, EOI, ERT\", \"Sent\", ala";
	//list.Add(str);

	//CString str = "\'\"\'z";

	int i;
	for (i = 0; i < list.GetCount(); i++)
	{
		CString l = list.GetAt(i);
		TRACE("\nLabel: %s\n", l);

		CStringArray a;
		int retval = TextUtilsEx::Tokenize(l, a);
		TextUtilsEx::TraceStringArray(a);
	}
}
#endif

#if 0
BOOL CharEncodingTest()
{
	SimpleString local1;
	SimpleString local2;
	SimpleString input;


	UINT CP_US_ASCII = 20127;
	UINT CP_ISO_8859_1 = 28591;

	input.Append("Hello Everybody");


	SimpleString* str = &input;
	UINT inCodePage = CP_ISO_8859_1;
	UINT outCodePage = CP_UTF8;
	SimpleString* result = &local1;
	SimpleString* workBuff = &local2;

	int buffLen = str->Count() * 4 + 2;
	workBuff->ClearAndResize(buffLen);
	LPWSTR buff = (LPWSTR)workBuff->Data();  // or  we could call MultiByteToWideChar first to get the required length
	int wlen = MultiByteToWideChar(inCodePage, 0, str->Data(), str->Count(), buff, buffLen);
	if (wlen == 0) {
		result->Clear();
		// error - implement error log file
		const DWORD error = ::GetLastError();
		return FALSE;
	}

	int outLen = wlen * 4 + 2;
	result->ClearAndResize(outLen); // or could  call WideCharToMultiByte first to get the required length
	if (outCodePage == 0)
		outCodePage = CP_UTF8;
	int utf8Len = WideCharToMultiByte(outCodePage, 0, buff, wlen, result->Data(), outLen, NULL, NULL);
	if (utf8Len == 0) {
		result->Clear();
		// error - implement error log file
		const DWORD error = ::GetLastError();
		/*ERROR_INSUFFICIENT_BUFFER.A supplied buffer size was not large enough, or it was incorrectly set to NULL.
		ERROR_INVALID_FLAGS.The values supplied for flags were not valid.
		ERROR_INVALID_PARAMETER.Any of the parameter values was invalid.
		ERROR_NO_UNICODE_TRANSLATION.Invalid Unicode was found in a string.*/
		return FALSE;
	}
	result->SetCount(utf8Len);
	return TRUE;
}
#endif

