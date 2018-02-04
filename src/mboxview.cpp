// mboxview.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "mboxview.h"
#include "profile.h"
#include "LinkCursor.h"

#include "MainFrm.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

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
#include <afxadv.h> //Has CRecentFileList class definition.
#include "afxlinkctrl.h"
#include "afxwin.h"

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
	void ParseParam(LPCTSTR lpszParam, BOOL bFlag, BOOL bLast);
};

void CCmdLine::ParseParam(LPCTSTR lpszParam, BOOL bFlag, BOOL) // bLast )
{
	if (bFlag) {
		if (strncmp(lpszParam, "FOLDER=", 5) == 0) {
			CString openFolder = lpszParam + 7;
			if (openFolder.Right(1) != "\\")
				openFolder += "\\";
			CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, "lastPath", openFolder);
		}
	}
}
/////////////////////////////////////////////////////////////////////////////
// CmboxviewApp initialization

BOOL CmboxviewApp::InitInstance()
{
	CCmdLine cmdInfo;
	ParseCommandLine(cmdInfo);
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

	CMainFrame* pFrame = new CMainFrame;
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
