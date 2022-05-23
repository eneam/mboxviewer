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

///////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////
// CPictureCtrlDemoDlg.cpp
// 
// Author: Tobias Eiseler
//
// Adapted for Windows MBox Viewer by the mboxview development
// Simplified, added re-orientation, added next, previous, rotate, zoom, dragging and print capabilities
// Significant portion of the code likely more than 80% is now custom
//
// E-Mail: tobias.eiseler@sisternicky.com
// 
// https://www.codeproject.com/Articles/24969/An-MFC-picture-control-to-dynamically-show-picture
///////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////


#include "stdafx.h"
#include "afxstr.h"
#include "FileUtils.h"
#include "TextUtilsEx.h"
#include "mboxview.h"
#include "PictureCtrl.h"
#include "CPictureCtrlDemoDlg.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

// CCPictureCtrlDemoDlg-Dialogfeld

CCPictureCtrlDemoDlg::CCPictureCtrlDemoDlg(CStringW *attachmentName, CWnd* pParent /*=NULL*/)
	: CDialogEx(CCPictureCtrlDemoDlg::IDD, pParent), m_picCtrl(this)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_ImageFileNameArrayPos = 0;
	CStringW fpath = FileUtils::GetMboxviewTempPathW();
	LoadImageFileNames(fpath);
	for (int i = 0; i < m_ImageFileNameArray.GetSize(); i++)
	{
		const wchar_t *fullFilePath = m_ImageFileNameArray[i]->operator LPCWSTR();
		CStringW fileName;
		FileUtils::CPathStripPathW(fullFilePath, fileName);

		if (attachmentName->Compare(fileName) == 0) {
			m_ImageFileNameArrayPos = i;
			break;
		}
	}

	m_picCtrl.m_bFixOrientation = TRUE;
	m_rotateType = Gdiplus::RotateNoneFlipNone;
	m_Zoom = 1;
	m_bZoomEnabled = FALSE;
	m_bDrawOnce = TRUE;
	m_sliderRange = 100;
	m_sliderFreq = 5;

	m_hightZoom = 1;
	m_widthZoom = 1;
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
	DDX_Control(pDX, IDC_SLIDER_ZOOM, m_sliderCtrl);
	int deb = 1;
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
	ON_WM_ERASEBKGND()
	ON_WM_MOUSEMOVE()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONUP()
	//ON_WM_RBUTTONDOWN()
	//ON_WM_RBUTTONUP()
	ON_WM_HSCROLL()
	ON_BN_CLICKED(IDCANCEL, &CCPictureCtrlDemoDlg::OnBnClickedCancel)
	ON_STN_CLICKED(IDC_STATIC_PICTURE, &CCPictureCtrlDemoDlg::OnStnClickedStaticPicture)
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

	this->SetBackgroundColor(RGB(0, 0, 0));

	m_sliderCtrl.SetRange(0, m_sliderRange);
	m_sliderCtrl.SetTicFreq(m_sliderFreq);

	EnableZoom(FALSE);

	UpdateData(TRUE);

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
	// DeferWindowPos ?? to reduce flicker
			//ScreenToClient(&rect);

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

	// filePath != 0
	CStringW *filePath = m_ImageFileNameArray[m_ImageFileNameArrayPos];

	CStringW fileName;
	LPCWSTR fPath = filePath->operator LPCWSTR();
	FileUtils::CPathStripPathW((wchar_t*)fPath, fileName);

	//Load an Image from File
	CString fileNameA;
	DWORD error;
	BOOL retW2A = TextUtilsEx::Wide2Ansi(fileName, fileNameA, error);
	SetWindowText(fileNameA);

	m_picCtrl.LoadFromFile(*filePath, m_rotateType, m_Zoom);
}

void CCPictureCtrlDemoDlg::OnBnClickedPrev()
{
	if (m_ImageFileNameArray.GetSize() <= 0) {
		MessageBeep(MB_OK);
		return;
	}

	m_picCtrl.m_bFixOrientation = TRUE;
	m_rotateType = Gdiplus::RotateNoneFlipNone;
	m_Zoom = 1;

	m_ImageFileNameArrayPos--;
	if (m_ImageFileNameArrayPos < 0) {
		m_ImageFileNameArrayPos = 0;
		MessageBeep(MB_OK);
	}

	m_picCtrl.ReleaseImage();
	m_picCtrl.ResetDrag();

	EnableZoom(FALSE);
	LoadImageFromFile();
}

void CCPictureCtrlDemoDlg::OnBnClickedNext()
{
	if (m_ImageFileNameArray.GetSize() <= 0) {
		MessageBeep(MB_OK);
		return;
	}

	m_picCtrl.m_bFixOrientation = TRUE;
	m_rotateType = Gdiplus::RotateNoneFlipNone;
	m_Zoom = 1;

	m_ImageFileNameArrayPos++;
	if (m_ImageFileNameArrayPos >= m_ImageFileNameArray.GetCount()) {
		m_ImageFileNameArrayPos = IntPtr2Int(m_ImageFileNameArray.GetCount() - 1);
		MessageBeep(MB_OK);
	}

	m_picCtrl.ReleaseImage();
	m_picCtrl.ResetDrag();

	EnableZoom(FALSE);
	LoadImageFromFile();
}

void CCPictureCtrlDemoDlg::OnBnClickedRotate()
{
	if (m_ImageFileNameArray.GetSize() <= 0) {
		MessageBeep(MB_OK);
		return;
	}

	m_rotateType = Gdiplus::Rotate90FlipNone;

	LoadImageFromFile();

	m_rotateType = Gdiplus::RotateNoneFlipNone;
}

void CCPictureCtrlDemoDlg::OnBnClickedZoom()
{
	if (m_ImageFileNameArray.GetSize() <= 0) {
		MessageBeep(MB_OK);
		return;
	}

	m_bZoomEnabled = !m_bZoomEnabled;

	EnableZoom(m_bZoomEnabled);

	LoadImageFromFile();
}

void CCPictureCtrlDemoDlg::EnableZoom(BOOL enableZoom)
{
	m_Zoom = 1;
	m_picCtrl.ResetDrag();
	m_sliderCtrl.SetPos(m_sliderRange / 2);
	GetDlgItem(IDC_SLIDER_POS)->SetWindowText(_T("0"));

	GetDlgItem(IDC_SLIDER_ZOOM)->EnableWindow(enableZoom);
	GetDlgItem(IDC_SLIDER_POS)->EnableWindow(enableZoom);

	m_bZoomEnabled = enableZoom;
}

BOOL CCPictureCtrlDemoDlg::isSupportedPictureFile(LPCTSTR file)
{
	PTSTR ext = PathFindExtension(file);
	CString cext = ext;
	if ((cext.CompareNoCase(_T(".png")) == 0) ||
		(cext.CompareNoCase(_T(".jpg")) == 0) ||
		(cext.CompareNoCase(_T(".pjpg")) == 0) ||
		(cext.CompareNoCase(_T(".jpeg")) == 0) ||
		(cext.CompareNoCase(_T(".pjpeg")) == 0) ||
		(cext.CompareNoCase(_T(".jpe")) == 0) ||
		(cext.CompareNoCase(_T(".bmp")) == 0) ||
		(cext.CompareNoCase(_T(".tif")) == 0) ||
		(cext.CompareNoCase(_T(".tiff")) == 0) ||
		(cext.CompareNoCase(_T(".dib")) == 0) ||
		(cext.CompareNoCase(_T(".jfif")) == 0) ||
		(cext.CompareNoCase(_T(".emf")) == 0) ||
		(cext.CompareNoCase(_T(".wmf")) == 0) ||
		(cext.CompareNoCase(_T(".ico")) == 0) ||
		(cext.CompareNoCase(_T(".gif")) == 0))
	{
		return TRUE;
	}
	else
		return FALSE;
}

int __cdecl FooPred(void const * first, void const * second)
{
	CStringW *f = *((CStringW**)first);
	CStringW *s = *((CStringW**)second);
	int ret = wcscmp(f->operator LPCWSTR(), s->operator LPCWSTR());
	return ret;
}

BOOL CCPictureCtrlDemoDlg::LoadImageFileNames(CStringW & dir)
{
	WIN32_FIND_DATAW FileData;
	HANDLE hSearch;
	CStringW	appPath;
	BOOL	bFinished = FALSE;

	// Start searching for all files in the current directory.
	CStringW searchPath = dir + CStringW(L"*.*");
	hSearch = FindFirstFileW(searchPath, &FileData);
	if (hSearch == INVALID_HANDLE_VALUE) {
		TRACE(_T("CCPictureCtrlDemoDlg: No files found.\n"));
		return FALSE;
	}
	while (!bFinished) {
		if (!(wcscmp(&FileData.cFileName[0], L".") == 0 || wcscmp(&FileData.cFileName[0], L"..") == 0)) {
			CStringW ext = PathFindExtensionW(&FileData.cFileName[0]);
			CString extA;
			DWORD error;
			BOOL retW2A = TextUtilsEx::Wide2Ansi(ext, extA, error);
			if (isSupportedPictureFile(extA))
			{
				if (FileData.dwFileAttributes != FILE_ATTRIBUTE_DIRECTORY) {
					CStringW	*fileFound = new CStringW;
					*fileFound = dir + &FileData.cFileName[0];
					m_ImageFileNameArray.Add(fileFound);
				}
			}
		}
		if (!FindNextFileW(hSearch, &FileData))
			bFinished = TRUE;
	}
	FindClose(hSearch);

	qsort(m_ImageFileNameArray.GetData(), m_ImageFileNameArray.GetCount(), sizeof(void*), FooPred);
	return TRUE;
}

void CCPictureCtrlDemoDlg::OnSize(UINT nType, int cx, int cy)
{
	CWnd::OnSize(nType, cx, cy);

	if (m_picCtrl.GetSafeHwnd())
	{
		// Changed repaint and commented out Invalidate to reduce flicker
		m_bDrawOnce = TRUE;
		//BOOL repaint = FALSE;
		BOOL repaint = TRUE;
		m_picCtrl.MoveWindow(20, 40, cx - 40, cy - 60, repaint);
	}
	else
		Invalidate();
	//Invalidate();
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
		const wchar_t *filePath = m_ImageFileNameArray[m_ImageFileNameArrayPos]->operator LPCWSTR();
		HINSTANCE result = ShellExecuteW(NULL, L"print", filePath, NULL, NULL, SW_SHOWNORMAL);
		if ((UINT_PTR)result <= MaxShellExecuteErrorCode) {
			CString errorText;
			ShellExecuteError2Text((UINT_PTR)result, errorText);
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, errorText, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_OK);
		}
	}
}

void CCPictureCtrlDemoDlg::UpdateRotateType(Gdiplus::RotateFlipType rotateType)
{
	m_rotateType = rotateType;
}

void CCPictureCtrlDemoDlg::FillRect(CBrush &brush)
{
	CDialog::OnPaint();
	CRect rect;
	GetClientRect(&rect);
	this->GetDC()->FillRect(&rect, &brush); // make dialog box black
}

BOOL CCPictureCtrlDemoDlg::OnEraseBkgnd(CDC* pDC)
{
	// TODO: Add your message handler code here and/or call default

	if (m_bDrawOnce)
	{
		m_bDrawOnce = FALSE;
		return CDialogEx::OnEraseBkgnd(pDC);
	}
	else
		return TRUE;
}

void CCPictureCtrlDemoDlg::OnMouseMove(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default

	m_picCtrl.ResetDraggedFlag();

	CDialogEx::OnMouseMove(nFlags, point);
}

void CCPictureCtrlDemoDlg::OnLButtonDown(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default

	CDialogEx::OnLButtonDown(nFlags, point);
}

void CCPictureCtrlDemoDlg::OnLButtonUp(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default

	CDialogEx::OnLButtonUp(nFlags, point);
}

void CCPictureCtrlDemoDlg::OnRButtonDown(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default

	CDialogEx::OnRButtonDown(nFlags, point);
}

void CCPictureCtrlDemoDlg::OnRButtonUp(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default

	CDialogEx::OnRButtonUp(nFlags, point);
}


void CCPictureCtrlDemoDlg::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar)
{
	// TODO: Add your message handler code here and/or call default

	nPos = m_sliderCtrl.GetPos();
	int pos = nPos;

	//TRACE(_T("Slider Position: %d\n"), pos);

	if (pos <= (m_sliderRange / 2)) 
	{
		pos = m_sliderRange / 2 - pos;
		m_Zoom = (float)pos /( m_sliderRange / 2);
		m_Zoom = (float)1 - m_Zoom;
		float minZoom = (float)0.05;
		if (m_Zoom < minZoom)
			m_Zoom = minZoom;
	}
	else {
		pos = pos - (m_sliderRange / 2);
		m_Zoom = (float)pos / 10 + 1;
	}

	//TRACE(_T("zoom=%f\n"), m_Zoom);

	CString zoomLevel;
	int level = nPos - (m_sliderRange / 2);
	zoomLevel.Format(_T("%d"), level);
	GetDlgItem(IDC_SLIDER_POS)->SetWindowText(zoomLevel);

	Invalidate();

	CDialogEx::OnHScroll(nSBCode, nPos, pScrollBar);
}


void CCPictureCtrlDemoDlg::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here
	CDialogEx::OnCancel();
}


void CCPictureCtrlDemoDlg::OnStnClickedStaticPicture()
{
	// TODO: Add your control notification handler code here
}
