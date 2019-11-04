// mboxview.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "ctype.h"
#include "locale.h"
#include "mboxview.h"
#include "profile.h"
#include "LinkCursor.h"

#include "MainFrm.h"
#include "MboxMail.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
//#define _CRTDBG_MAP_ALLOC  
//#include <stdlib.h>  
//#include <crtdbg.h> 
#endif

LONG WINAPI MyUnhandledExceptionFilter(PEXCEPTION_POINTERS pExceptionPtrs)
{
	// Do something, for example generate error report
	//..
	// Execute default exception handler next
	return EXCEPTION_EXECUTE_HANDLER;
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

bool PathExists(LPCSTR file)
{
	WIN32_FIND_DATA fd;
	ZeroMemory(&fd, sizeof(fd));
	HANDLE h = FindFirstFile(file, &fd);
	if( h == INVALID_HANDLE_VALUE )
		return false;
	bool bRes = false;
	bRes = ((fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == FILE_ATTRIBUTE_DIRECTORY);
	FindClose(h);
	return bRes;
}

// duplicated from NListView.cpp , may need to generalize
BOOL _PathFileExist(LPCSTR path)
{
	DWORD fa = GetFileAttributes(path);

	if (fa != (DWORD)0xFFFFFFFF)
		return TRUE;
	return FALSE;
}

BOOL isNumeric(CString &str) {
	int i = 0;
	if (str[i] == '-')
		i++;
	if (i == str.GetLength())
		return FALSE;
	for (; i < str.GetLength(); i++) {
		if (!_istdigit(str[i])) {
			return FALSE;
		}
	}
	return TRUE;
}

BOOL Str2Wide(CString &res, UINT CodePage, CStringW &m_strW)
{
	int len = res.GetLength() * 4 + 2;
	LPWSTR buff = (LPWSTR)malloc(len);  // or  we could call MultiByteToWideChar first to get the required length
	int len1 = MultiByteToWideChar(CodePage, 0, res, res.GetLength(), buff, len);
	if (len1 == 0) {
		free(buff);
		// error - implement error log file
		const DWORD error = ::GetLastError();
		return FALSE;
	}
	buff[len1] = 0;
	m_strW = buff;
	free(buff);
	return TRUE;
}

#include <afxadv.h> //Has CRecentFileList class definition.
#include "afxlinkctrl.h"
#include "afxwin.h"

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

	if( ! PathExists((*m_pRecentFileList)[nIndex]) ) {
		m_pRecentFileList->Remove(nIndex);
		return;
	}
	((CMainFrame*)AfxGetMainWnd())->DoOpen((*m_pRecentFileList)[nIndex]);
}

/////////////////////////////////////////////////////////////////////////////
// CmboxviewApp construction

CmboxviewApp::CmboxviewApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance

	//bool ret = TextUtilities::TestAll();

	//_CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF | _CRTDBG_CHECK_ALWAYS_DF);
	//SetNoOaCache();

	//int test_enc();
	//test_enc();
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

	MboxMail::ShowHint(HintConfig::GeneralUsageHint);
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
	if (bFlag) {
		if (strncmp(lpszParam, _T("FOLDER="), 7) == 0) {
			CString openFolder = lpszParam + 7;
			if (openFolder.Right(1) != _T("\\"))
				openFolder += _T("\\");
			if (m_bLastPathSet == FALSE)
				CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("lastPath"), openFolder);
		}
		else if (strncmp(lpszParam, _T("MAIL_FILE="), 10) == 0) {
			CString mailFile = lpszParam + 10;
			TCHAR ext[_MAX_EXT + 1]; ext[0] = 0;
			TCHAR drive[_MAX_DRIVE + 1]; drive[0] = 0;
			TCHAR dir[_MAX_DIR + 1]; dir[0] = 0;
			TCHAR fname[_MAX_FNAME + 1];

			_tsplitpath(mailFile, drive, dir, fname, ext);
			if (_tcslen(dir) != 0) {
				CString txt;
				if (!_PathFileExist(mailFile)) {
					CString txt = _T("Nonexistent File \"") + mailFile;
					txt += _T("\".\nDo you want to continue?");
					HWND h = NULL; // we don't have any window yet  
					int answer = MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
					if (answer == IDNO)
						m_bError = TRUE;
				}
				else {
					CString lastPath = CString(drive) + dir;
					CString file(fname); file += ext;
					m_bLastPathSet = TRUE;
					CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("lastPath"), lastPath);
					CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("mailFile"), file);

				}
			}
			else {
				CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("mailFile"), mailFile);
			}
		}
		else if (strncmp(lpszParam, _T("EXPORT_EML="), 11) == 0) {
			CString exportEML = lpszParam + 11;
			exportEML.MakeLower();
			if (!((exportEML.Compare("y") == 0) || (exportEML.Compare("n") == 0))) {
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
			}
		}
		else if (strncmp(lpszParam, _T("PROGRESS_BAR_DELAY="), 19) == 0) {
			CString barDelay = lpszParam + 19;
			// Validate
			if (!isNumeric(barDelay))
			{
				CString txt = _T("Invalid Command Line Option Value \"");
				CString opt = lpszParam;
				txt += opt + _T("\".\nDo you want to continue?");
				HWND h = NULL; // we don't have any window yet  
				int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
				if (answer == IDNO)
					m_bError = TRUE;
			}
			else {
				DWORD progressBarDelay = _tstoi(barDelay);
				CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("progressBarDelay"), progressBarDelay);
			}
		}
		else {
			// Unknown argument
			CString txt = _T("Invalid Command Line Option \"");
			CString opt = lpszParam;
			txt += opt + _T("\". \nValid options: -FOLDER=,-MAIL_FILE=,-EXPORT_EML=,\n-PROGRESS_BAR_DELAY=.\nDo you want to continue?");
			HWND h = NULL; // we don't have any window yet
			int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
			if (answer == IDNO)
				m_bError = TRUE;
		}
	}
}
/////////////////////////////////////////////////////////////////////////////
// CmboxviewApp initialization

BOOL CmboxviewApp::InitInstance()
{
	//SetUnhandledExceptionFilter(MyUnhandledExceptionFilter);

	CCmdLine cmdInfo;
	CString mailFile;
	CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "mailFile", mailFile);
	ParseCommandLine(cmdInfo);
	if (cmdInfo.m_bError)
	{
		MboxMail::ReleaseResources();
		return FALSE;
	}
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
		"Path %d", 8);
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
