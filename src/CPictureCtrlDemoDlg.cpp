///////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////
// CPictureCtrlDemoDlg.cpp
// 
// Author: Tobias Eiseler
//
// Adapted for Windows MBox Viewer by the mboxview development
// Simplified, added next, previous, rotate and zoom capabilities
//
// E-Mail: tobias.eiseler@sisternicky.com
// 
// https://www.codeproject.com/Articles/24969/An-MFC-picture-control-to-dynamically-show-picture
///////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////


#include "stdafx.h"
#include "afxstr.h"
#include "mboxview.h"
#include "PictureCtrl.h"
#include "CPictureCtrlDemoDlg.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

CString GetmboxviewTempPath(char *name = 0);

// CCPictureCtrlDemoDlg-Dialogfeld

CCPictureCtrlDemoDlg::CCPictureCtrlDemoDlg(CString *attachmentName, CWnd* pParent /*=NULL*/)
	: CDialogEx(CCPictureCtrlDemoDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_ImageFileNameArrayPos = 0;
	CString fpath = GetmboxviewTempPath();
	LoadImageFileNames(fpath);
	for (int i = 0; i < m_ImageFileNameArray.GetSize(); i++)
	{
		const char *fullFilePath = m_ImageFileNameArray[i]->operator LPCSTR();
		char *fileName = new char[m_ImageFileNameArray[i]->GetLength() + 1];
		strcpy(fileName, fullFilePath);
		PathStripPath(fileName);

		if (attachmentName->Compare(fileName) == 0) {
			m_ImageFileNameArrayPos = i;
			delete[] fileName;
			break;
		}
		delete[] fileName;
	}

	m_rotateType = Gdiplus::RotateNoneFlipNone;
	m_Zoom = 0;
	m_ZoomMax = 16;
	m_ZoomMaxForCurrentImage = m_ZoomMax;
}

CCPictureCtrlDemoDlg::~CCPictureCtrlDemoDlg()
{
	for (int i = 0; i < m_ImageFileNameArray.GetSize(); i++)
	{
		delete m_ImageFileNameArray[i];
	}
}

void CCPictureCtrlDemoDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_STATIC_PICTURE, m_picCtrl);
}

BEGIN_MESSAGE_MAP(CCPictureCtrlDemoDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_SIZE()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BUTTON_PREV, &CCPictureCtrlDemoDlg::OnBnClickedPrev)
	ON_BN_CLICKED(IDC_BUTTON_NEXT, &CCPictureCtrlDemoDlg::OnBnClickedNext)
	ON_BN_CLICKED(IDC_BUTTON_ROTATE, &CCPictureCtrlDemoDlg::OnBnClickedRotate)
	ON_BN_CLICKED(IDC_BUTTON_ZOOM, &CCPictureCtrlDemoDlg::OnBnClickedZoom)
	ON_BN_CLICKED(IDC_BUTTON_PRT, &CCPictureCtrlDemoDlg::OnBnClickedButtonPrt)
END_MESSAGE_MAP()


// CCPictureCtrlDemoDlg-Meldungshandler

BOOL CCPictureCtrlDemoDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add the menu command "Info ..." to the system menu.
	// IDM_ABOUTBOX must be in the system commands area.

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
	}

	// Set icon for this dialog box. Will be done automatically
	// if the main window of the application is not a dialog box
	SetIcon(m_hIcon, TRUE);			// Großes Symbol verwenden // use big icon
	SetIcon(m_hIcon, FALSE);		// Kleines Symbol verwenden // Use a small icon

	// TODO: Insert additional initialization here

	//SetBackgroundColor(RGB(0, 0, 0));

	// Return TRUE unless a control is to receive the focus
	return TRUE;  // Geben Sie TRUE zurück, außer ein Steuerelement soll den Fokus erhalten
}

void CCPictureCtrlDemoDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	CDialog::OnSysCommand(nID, lParam);
}

// If you add a Minimize button to the dialog, you will need
// the code below to draw the symbol. For MFC applications using the
// Use Document / View Model, this will be done automatically.

void CCPictureCtrlDemoDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // Gerätekontext zum Zeichnen // device context for drawing

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center symbol in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Symbol zeichnen // draw a symbol
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
		CRect rect;
		GetClientRect(&rect);
		this->GetDC()->FillRect(&rect, &CBrush(RGB(0, 0, 0))); // make dialog box black
#if 0
		if (m_picCtrl.GetSafeHwnd()) {
			m_picCtrl.GetClientRect(&rect);
			m_picCtrl.GetDC()->FillRect(&rect, &CBrush(RGB(0, 0, 0)));
		}
#endif
		if (m_picCtrl.GetSafeHwnd()) {
			LoadImageFromFile();
		}
	}
}

// The system calls this function to query the cursor that is displayed while the user is
// drag the minimized window with the mouse.
HCURSOR CCPictureCtrlDemoDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CCPictureCtrlDemoDlg::LoadImageFromFile()
{
	if (m_ImageFileNameArray.GetSize() <= 0) {
		MessageBeep(MB_OK);
		return;
	}

	if ((m_ImageFileNameArrayPos >= m_ImageFileNameArray.GetCount()) ||
		(m_ImageFileNameArrayPos < 0))
		return;

	// fname != 0
	CString *fname = m_ImageFileNameArray[m_ImageFileNameArrayPos];

	char *fileName = new char[fname->GetLength() + 1];
	strcpy(fileName, fname->GetBuffer());
	PathStripPath(fileName);

	//Load an Image from File
	SetWindowText(fileName);
	delete[] fileName;
	m_picCtrl.LoadFromFile(*fname, m_rotateType, m_Zoom);
}

void CCPictureCtrlDemoDlg::OnBnClickedPrev()
{
	if (m_ImageFileNameArray.GetSize() <= 0) {
		MessageBeep(MB_OK);
		return;
	}

	m_rotateType = Gdiplus::RotateNoneFlipNone;
	m_Zoom = 0;
	m_ZoomMaxForCurrentImage = m_ZoomMax;
	m_picCtrl.SetZoomMaxForCurrentImage(m_ZoomMax);

	m_ImageFileNameArrayPos--;
	if (m_ImageFileNameArrayPos < 0) {
		m_ImageFileNameArrayPos = 0;
		MessageBeep(MB_OK);
	}

	LoadImageFromFile();
}

void CCPictureCtrlDemoDlg::OnBnClickedNext()
{
	if (m_ImageFileNameArray.GetSize() <= 0) {
		MessageBeep(MB_OK);
		return;
	}

	m_rotateType = Gdiplus::RotateNoneFlipNone;
	m_Zoom = 0;
	m_ZoomMaxForCurrentImage = m_ZoomMax;
	m_picCtrl.SetZoomMaxForCurrentImage(m_ZoomMax);

	m_ImageFileNameArrayPos++;
	if (m_ImageFileNameArrayPos >= m_ImageFileNameArray.GetCount()) {
		m_ImageFileNameArrayPos = m_ImageFileNameArray.GetCount() - 1;
		MessageBeep(MB_OK);
	}

	LoadImageFromFile();
}

void CCPictureCtrlDemoDlg::OnBnClickedRotate()
{
	if (m_ImageFileNameArray.GetSize() <= 0) {
		MessageBeep(MB_OK);
		return;
	}

	m_rotateType = (Gdiplus::RotateFlipType)((int)m_rotateType + 1);
	if (m_rotateType > Gdiplus::Rotate270FlipNone)
		m_rotateType = Gdiplus::RotateNoneFlipNone;

	LoadImageFromFile();
}

void CCPictureCtrlDemoDlg::OnBnClickedZoom()
{
	if (m_ImageFileNameArray.GetSize() <= 0) {
		MessageBeep(MB_OK);
		return;
	}

	m_ZoomMaxForCurrentImage = m_ZoomMax;
	int zoomMaxForCurrentImage = m_picCtrl.GetZoomMaxForCurrentImage();
	if ((zoomMaxForCurrentImage > 0) && (zoomMaxForCurrentImage < m_ZoomMax))
		m_ZoomMaxForCurrentImage = zoomMaxForCurrentImage;

	m_Zoom += 1;
	if (m_Zoom >= m_ZoomMaxForCurrentImage)
		m_Zoom = 0;

	LoadImageFromFile();
}

BOOL CCPictureCtrlDemoDlg::isSupportedPictureFile(LPCSTR file)
{
	PTSTR ext = PathFindExtension(file);
	CString cext = ext;
	if ((cext.CompareNoCase(".png") == 0) ||
		(cext.CompareNoCase(".jpg") == 0) ||
		(cext.CompareNoCase(".jpeg") == 0) ||
		(cext.CompareNoCase(".jpe") == 0) ||
		(cext.CompareNoCase(".bmp") == 0) ||
		(cext.CompareNoCase(".tif") == 0) ||
		(cext.CompareNoCase(".tiff") == 0) ||
		(cext.CompareNoCase(".dib") == 0) ||
		(cext.CompareNoCase(".jfif") == 0) ||
		(cext.CompareNoCase(".emf") == 0) ||
		(cext.CompareNoCase(".wmf") == 0) ||
		(cext.CompareNoCase(".ico") == 0) ||
		(cext.CompareNoCase(".gif") == 0))
	{
		return TRUE;
	}
	else
		return FALSE;
}

BOOL CCPictureCtrlDemoDlg::LoadImageFileNames(CString & dir)
{
	WIN32_FIND_DATA FileData;
	HANDLE hSearch;
	CString	appPath;
	BOOL	bFinished = FALSE;

	// Start searching for all files in the current directory.
	CString searchPath = dir + CString("*.*");
	hSearch = FindFirstFile(searchPath, &FileData);
	if (hSearch == INVALID_HANDLE_VALUE) {
		TRACE("No files found.");
		return FALSE;
	}
	while (!bFinished) {
		if (!(strcmp(&FileData.cFileName[0], ".") == 0 || strcmp(&FileData.cFileName[0], "..") == 0)) {
			PTSTR ext = PathFindExtension(&FileData.cFileName[0]);
			CString cext = ext;
			if (isSupportedPictureFile(&FileData.cFileName[0]))
			{
				if (FileData.dwFileAttributes != FILE_ATTRIBUTE_DIRECTORY) {
					CString	*fileFound = new CString;
					*fileFound = dir + "\\" + &FileData.cFileName[0];
					m_ImageFileNameArray.Add(fileFound);
				}
			}
		}
		if (!FindNextFile(hSearch, &FileData))
			bFinished = TRUE;
	}
	FindClose(hSearch);
	return TRUE;
}

void CCPictureCtrlDemoDlg::OnSize(UINT nType, int cx, int cy)
{
	CWnd::OnSize(nType, cx, cy);

	if (m_picCtrl.GetSafeHwnd())
	{
		m_picCtrl.MoveWindow(20, 40, cx-40, cy-60);
	}
	Invalidate();
	RedrawWindow();
}


void CCPictureCtrlDemoDlg::OnBnClickedButtonPrt()
{
#define MaxShellExecuteErrorCode 32

	// TODO: Add your control notification handler code here
	if (m_ImageFileNameArray.GetSize() <= 0) {
		MessageBeep(MB_OK);
		return;
	}

	if ((m_ImageFileNameArrayPos >= 0) && (m_ImageFileNameArrayPos <= m_ImageFileNameArray.GetSize()))
	{
		const char *fullFilePath = m_ImageFileNameArray[m_ImageFileNameArrayPos]->operator LPCSTR();
		HINSTANCE result = ShellExecute(NULL, _T("print"), fullFilePath, NULL, NULL, SW_SHOWNORMAL);
		if ((UINT)result <= MaxShellExecuteErrorCode) {
			CString errorText;
			ShellExecuteError2Text((UINT)result, errorText);
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, errorText, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		}
	}
}
