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

	char const*  szCause = seDescription(seNumb);

	char *stackDumpFileName = "UnhandledException_StackDump.txt";
	int mailPosition = MboxMail::s_mails.GetCount();
	MyStackWalker *sw  = nullptr;
	if (MboxMail::glStackWalker)
	{
		sw = MboxMail::glStackWalker;
	}
	BOOL ret = DumpStack(sw, stackDumpFileName, (TCHAR*)szCause, seNumb, pExceptionPtrs->ContextRecord, mailPosition);

	CString progDir;
	BOOL retDir = GetProgramDir(progDir);

	char *exceptionName = "UnhandledException";
	CString errorTxt;
#ifdef USE_STACK_WALKER
	errorTxt.Format(_T("%s: Code=%8.8x Description=%s\n\n"
		"To help to diagnose the problem, created file\n\n%s\n\nin\n\n%s directory.\n\n"
		"Please provide the files to the development team.\n\n"),
		exceptionName, seNumb, szCause, stackDumpFileName, progDir);
#else
	errorTxt.Format(_T("%s: Code=%8.8x Description=%s\n\n"), exceptionName, seNumb, szCause);
#endif
	AfxMessageBox((LPCTSTR)errorTxt, MB_OK | MB_ICONHAND);

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

	CString txtA;
	DWORD error;
	BOOL retW2A = TextUtilsEx::Wide2Ansi(txt, txtA, error);

	AfxMessageBox(txtA, MB_OK | MB_ICONHAND);

	// GetExceptionPointers(0, &pExceptionPtrs);

	// Write minidump file
	//CreateMiniDump(pExceptionPtrs);

	// Terminate process
	TerminateProcess(GetCurrentProcess(), 1);
}


// It looks by AddVectoredExceptionHandler() will catch all kinds of exceptions,
// However, some of them must be ignored be ignored and it is difficult to know which exceptions
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

const char *sz_Software_mboxview = "SOFTWARE\\mboxview";

/////////////////////////////////////////////////////////////////////////////
// CmboxviewApp

BEGIN_MESSAGE_MAP(CmboxviewApp, CWinApp)
	//{{AFX_MSG_MAP(CmboxviewApp)
	ON_COMMAND(ID_APP_ABOUT, OnAppAbout)
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


void CmboxviewApp::MyMRUFileHandler(UINT i)
{
	ASSERT_VALID(this);
	ASSERT(m_pRecentFileList != NULL);

	ASSERT(i >= ID_FILE_MRU_FILE1);
	ASSERT(i < ID_FILE_MRU_FILE1 + (UINT)m_pRecentFileList->GetSize());

	CString strName, strCurDir, strMessage;
	int nIndex = i - ID_FILE_MRU_FILE1;
	ASSERT((*m_pRecentFileList)[nIndex].GetLength() != 0);

	strName.Format("MRU: open file (%d) '%s'.\n", (nIndex) + 1,(LPCTSTR)(*m_pRecentFileList)[nIndex]);

	if( ! FileUtils::PathDirExists((*m_pRecentFileList)[nIndex]) ) {
		m_pRecentFileList->Remove(nIndex);
		return;
	}
	((CMainFrame*)AfxGetMainWnd())->DoOpen((*m_pRecentFileList)[nIndex]);
}

/////////////////////////////////////////////////////////////////////////////
// CmboxviewApp construction

//#define IHASH_MAP_TEST
#ifdef  IHASH_MAP_TEST
#include "IHashTable.h"
#endif

CmboxviewApp::CmboxviewApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance

	//bool ret = TextUtilities::TestAll();

	//_CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF | _CRTDBG_CHECK_ALWAYS_DF);
	//SetNoOaCache();

	//int test_enc();
	//test_enc();

#ifdef  IHASH_MAP_TEST
	bool IHashMapTest();

	bool ret = IHashMapTest();
	exit(8);
#endif

	//FileUtils FU;  FU.UnitTest();
}

/////////////////////////////////////////////////////////////////////////////
// The one and only CmboxviewApp object

CmboxviewApp theApp;

void SetBrowserEmulation()
{
	CString ieVer = CProfile::_GetProfileString(HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Internet Explorer", "Version");
	int p = ieVer.Find('.');
	int maj = atoi(ieVer);
	int min = atoi(ieVer.Mid(p + 1));
	DWORD dwEmulation = CProfile::_GetProfileInt(HKEY_LOCAL_MACHINE, "Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", "mboxview.exe");
	if (maj == 9 && min >= 10) {
		DWORD dwem = min * 1000 + 1;
		if (dwEmulation != dwem) {
			CProfile::_WriteProfileInt(HKEY_CURRENT_USER, "Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", "mboxview.exe", dwem);
			CProfile::_WriteProfileInt(HKEY_LOCAL_MACHINE, "Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", "mboxview.exe", dwem);
		}
	}
	else
		if (maj == 9) {
			if (dwEmulation != 9999) {
				CProfile::_WriteProfileInt(HKEY_CURRENT_USER, "Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", "mboxview.exe", 9999);
				CProfile::_WriteProfileInt(HKEY_LOCAL_MACHINE, "Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", "mboxview.exe", 9999);
			}
		}
		else
			if (maj == 8) {
				if (dwEmulation != 8888) {
					CProfile::_WriteProfileInt(HKEY_CURRENT_USER, "Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", "mboxview.exe", 8888);
					CProfile::_WriteProfileInt(HKEY_LOCAL_MACHINE, "Software\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", "mboxview.exe", 8888);
				}
			}
}

#pragma comment(linker, "/defaultlib:version.lib")

BOOL GetFileVersionInfo(HMODULE hModule, DWORD &ms, DWORD &ls)
{
	// get module handle
	TCHAR filename[_MAX_PATH];

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
	if (!VerQueryValue(pVersionInfo, _T("\\"), &lpvi, &iLen)) {
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
	return CWinApp::ExitInstance();
}

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();
	CFont m_linkFont;
// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

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

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
	DDX_Control(pDX, IDC_DONATION, m_link);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
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


/////////////////////////////////////////////////////////////////////////////
// CmboxviewApp message handlers

BOOL CAboutDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	DWORD ms, ls;
    if (GetFileVersionInfo((HMODULE)AfxGetInstanceHandle(), ms, ls)) {
		CString txt;
        // display file version from VS_FIXEDFILEINFO struct
        txt.Format("Version %d.%d.%d.%d", 
                 HIWORD(ms), LOWORD(ms),
                 HIWORD(ls), LOWORD(ls));
		GetDlgItem(IDC_STATIC1)->SetWindowText(txt);
	}
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

class CCmdLine : public CCommandLineInfo
{
public:
	BOOL m_bError;
	BOOL m_bLastPathSet;
	CCmdLine::CCmdLine() {
		m_bError = FALSE; m_bLastPathSet = FALSE;
	}
	void ParseParam(LPCTSTR lpszParam, BOOL bFlag, BOOL bLast);
};

void CCmdLine::ParseParam(LPCTSTR lpszParam, BOOL bFlag, BOOL) // bLast )
{
	if (m_bError)
		return;
	if (bFlag) 
	{
		if (strncmp(lpszParam, _T("FOLDER="), 7) == 0) 
		{
			CString openFolder = lpszParam + 7;
			CMainFrame::m_commandLineParms.m_mboxFolderPath = openFolder;
		}
		else if (strncmp(lpszParam, _T("MAIL_FILE="), 10) == 0) 
		{
			CString mailFile = lpszParam + 10;

			CMainFrame::m_commandLineParms.m_mboxFileNameOrPath = mailFile;
		}
		else if (strncmp(lpszParam, _T("EXPORT_EML="), 11) == 0) 
		{
			CString exportEML = lpszParam + 11;
			exportEML.MakeLower();
			if (!((exportEML.Compare("y") == 0) || (exportEML.Compare("n") == 0))) 
			{
				CString txt = _T("Invalid Command Line Option Value \"");
				CString opt = lpszParam;
				txt += opt + _T("\". Valid are \"y|n\". Note that once defined valid EXPORT_EML persists in the registry.\nDo you want to continue?");
				HWND h = NULL; // we don't have any window yet  
				int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
				if (answer == IDNO)
					m_bError = TRUE;
			}
			else {
				CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("exportEML"), exportEML);

				CMainFrame::m_commandLineParms.m_exportEml = FALSE;
				if (exportEML.Compare("y") == 0)
					CMainFrame::m_commandLineParms.m_exportEml = TRUE;
			}
		}
		else if (strncmp(lpszParam, _T("PROGRESS_BAR_DELAY="), 19) == 0) 
		{
			CString barDelay = lpszParam + 19;
			// Validate
			if (!TextUtilsEx::isNumeric(barDelay))
			{
				CString txt = _T("Invalid Command Line Option Value \"");
				CString opt = lpszParam;
				txt += opt + _T("\".\nDo you want to continue?");
				HWND h = NULL; // we don't have any window yet  
				int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
				if (answer == IDNO)
					m_bError = TRUE;
			}
			else 
			{
				DWORD progressBarDelay = _tstoi(barDelay);
				CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("progressBarDelay"), progressBarDelay);
				CMainFrame::m_commandLineParms.m_progressBarDelay = progressBarDelay;
			}
		}
		else if (strncmp(lpszParam, _T("EML_PREVIEW_MODE"), 19) == 0)
		{
			CMainFrame::m_commandLineParms.m_bEmlPreviewMode = TRUE;
		}
		else if (strncmp(lpszParam, _T("MBOX_MERGE_LIST_FILE"), 20) == 0)
		{
			CString emlListFile = lpszParam + 21;
			CMainFrame::m_commandLineParms.m_mboxListFilePath = emlListFile;
		}
		else if (strncmp(lpszParam, _T("MBOX_MERGE_TO_FILE"), 18) == 0)
		{
			CString mergeToFilePath = lpszParam + 19;
			FileUtils::NormalizeFilePath(mergeToFilePath);
			CMainFrame::m_commandLineParms.m_mergeToFilePath = mergeToFilePath;
		}
		else if (strncmp(lpszParam, _T("TRACE_CASE"), 10) == 0)
		{
			CString traceCase = lpszParam + 11;
			DWORD nTraceCase = _tstoi(traceCase);
			CMainFrame::m_commandLineParms.m_traceCase = nTraceCase;
		}
		else 
		{
			// Unknown argument
			CString txt = _T("Invalid Command Line Option \"");
			CString opt = lpszParam;
			txt += opt 
				+ _T("\"")
				+ _T("\n\nValid options:\n")
				+ _T("  -FOLDER=Folder Path\n")
				+ _T("  -MAIL_FILE=Mbox File Path or Name to open\n")
				+ _T("  -EXPORT_EML=y or n\n")
				+ _T("  -PROGRESS_BAR_DELAY=Seconds\n")
				+ _T("  -EML_PREVIEW_MODE\n")
				+ _T("  -MBOX_MERGE_LIST_FILE=Path to File containing list of mbox files to merge\n")
				+ _T("  -MBOX_MERGE_TO_FILE=Path to File to save merge results\n")
				+ _T("\nDo you want to continue?")
				;

			CMainFrame::m_commandLineParms.Clear();
			HWND h = NULL; // we don't have any window yet
			int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_YESNO);
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
	void *p = new char[size];
	tot += size;
	return p;
}
#endif

BOOL CmboxviewApp::InitInstance()
{
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
	MyStackWalker *sw = MboxMail::glStackWalker;
	if (sw)
	{
		// preload modules, etc to minimize work in MyUnhandledExceptionFilter();
		sw->LoadAllModules();
		sw->ClearBuffer();
	}

#if 0
	char *stackDumpFileName = "UnhandledException_StackDump.txt";
	UINT seNumb = 0;
	CONTEXT *ContextRecord = nullptr;
	int mailPosition = -1;
	BOOL ret = DumpStack(sw, stackDumpFileName, (TCHAR*)"Preload", seNumb, ContextRecord, mailPosition);
#endif

	/// Disable "Whole program Optimization" to get accurate Stack Trace  !!!!
	SetMyExceptionHandler();
#endif

	//AfxEnableMemoryTracking(TRUE);
	//afxMemDF = allocMemDF | delayFreeMemDF | checkAlwaysMemDF;

	CCmdLine cmdInfo;
	CString mailFile;
	CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "mailFile", mailFile);

	CMainFrame::m_commandLineParms.m_allCommanLineOptions = CWinApp::m_lpCmdLine;
	ParseCommandLine(cmdInfo);
	if (cmdInfo.m_bError)
	{
		MboxMail::ReleaseResources();
		return FALSE;
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
		else if (retCode == -2)  // error but continue; probably overkill
		{
			CMainFrame::m_commandLineParms.Clear();
		}
	}

	CString processPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "processPath");
#if 0
	// Commented out for now since it seem to confuse users. Need to find better solution.
	if (!processPath.IsEmpty())
	{
		CString txt = _T("Mbox viewer instance might be running already:\n\n") + processPath;
		txt += _T("\n\n");
		txt += _T("Only single instance should be running to avoid potential\n");
		txt += _T("issues since all instances will share the same data in the registry.\n\n");
		txt += _T("In most cases this warning can be ignored.\n"
			"It will be generated if program crashes or it is killed from Task Manager.\n\n");
		txt += _T("Do you want to continue?");
		HWND h = NULL; // we don't have any window yet  
		int answer = MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
		{
			MboxMail::ReleaseResources();
			return FALSE;
		}
	}
#endif

	char *pValue;
	errno_t  er = _get_pgmptr(&pValue);
	CString procFullPath;
	if ((er == 0) && pValue)
		procFullPath.Append(pValue);

	CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("processPath"), procFullPath);


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
	SetRegistryKey(_T("mboxview"));

	m_pRecentFileList = new
		CRecentFileList(0, "MRUs",
		"Path %d", 16);
	m_pRecentFileList->ReadList();
	// Initialize all Managers for usage. They are automatically constructed
	// if not yet present
	// To create the main window, this code creates a new frame window
	// object and then sets it as the application's main window object.

	g_tu.Init();

	DWORD msgViewPosition = 1;
	BOOL retval;
	retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("messageWindowPosition"), msgViewPosition);
	if (retval == TRUE) {
		if ((msgViewPosition < 1) && (msgViewPosition > 3))
			msgViewPosition = 1;  // bottom=1
	}

	CMainFrame* pFrame = new CMainFrame(msgViewPosition);
	if (pFrame->GetSafeHwnd() == 0)
		int deb = 1;
	m_pMainWnd = pFrame;

	// create and load the frame with its resources

	pFrame->LoadFrame(IDR_MAINFRAME,
		WS_OVERLAPPEDWINDOW | FWS_ADDTOTITLE, NULL,
		NULL);

	// The one and only window has been initialized, so show and update it.
	pFrame->ShowWindow(SW_SHOW);
	pFrame->UpdateWindow();
	
	DWORD ms, ls;
	if (GetFileVersionInfo((HMODULE)AfxGetInstanceHandle(), ms, ls)) {
		CString txt;
		// display file version from VS_FIXEDFILEINFO struct
		txt.Format("%d.%d.%d.%d",
			HIWORD(ms), LOWORD(ms),
			HIWORD(ls), LOWORD(ls));
		CString savedVer = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "version");
		if (savedVer.Compare(txt) != 0) {
			CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "version", txt);
			pFrame->PostMessage(WM_COMMAND, ID_APP_ABOUT, 0);
		}
	}

	MboxMail::LoadHintBitmap();

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

	BOOL r = TextUtilsEx::Ansi2Wide(fileName, fileNameW, error);

	HANDLE hFile = CreateFileW(fileNameW, dwAccess, FILE_SHARE_READ, NULL, dwCreationDisposition, 0, NULL);
	if (hFile != INVALID_HANDLE_VALUE)
	{
		BOOL retClose = CloseHandle(hFile);
	}
#endif

	return TRUE;
}


void CAboutDlg::OnStnClickedDonation()
{
	HINSTANCE result = ShellExecute(NULL, _T("open"), "https://sourceforge.net/p/mbox-viewer/donate/", NULL, NULL, SW_SHOWNORMAL);
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
			GetFont()->GetObject(sizeof(lf), &lf);
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
		Item(std::string &nm) { name = nm; }
		//
		dlink_node<Item> hashNode;
		dlink_node<Item> hashNode2;
		std::string name;
	};

	struct ItemHash {
		unsigned long operator()(const Item *key) const
		{
			unsigned long  hashsum = std::hash<std::string>{}(key->name);
			return hashsum;
		};
		unsigned long operator()(Item *key) const
		{
			unsigned long  hashsum = std::hash<std::string>{}(key->name);
			return hashsum;
		};
	};

	struct ItemEqual {
		bool operator()(const Item *key, const Item *item) const
		{
			return (key->name == item->name);
		};
		bool operator()(Item *key, const Item *item) const
		{
			return (key->name == item->name);
		};
	};

	using MyHashMap = IHashMap<Item, Item, ItemHash, ItemEqual, &Item::hashNode>;
	using MyHashMap2 = IHashMap<Item, Item, ItemHash, ItemEqual, &Item::hashNode2>;

	unsigned long sz = 10;

	Item *it = new Item(std::string("John"));
	Item *it2 = new Item(std::string("Irena"));
	Item *it3 = new Item(std::string("Lucyna"));
	Item *it4 = new Item(std::string("Jan"));

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
	Item *item;
	if ((item = hashMap.find(it2)) != nullptr)
	{
		unsigned long hashsum = ItemHash{}(it2);
		hashMap.remove(item, hashsum);
	}
	assert(hashMap.count() == 2);

	// Add back
	unsigned long hashsum = ItemHash{}(it2);
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
	TRACE("%d,%d=%d\n", L,R,p);
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

	for (i = L; i <= L+R; i++)
	{
		for (j = 2; j <= modNumbs; ++j)
		{
			if (i % j == 0)
				break;
		}
		if (j == modNumbs)
			break;
	}
	if (i < (L+R))
		return i;
	else
		return L;
}
#endif

