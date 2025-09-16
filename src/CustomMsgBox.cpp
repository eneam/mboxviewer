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

// CustomMsgBox.cpp : implementation file
//

#include "stdafx.h"
#include "afxdialogex.h"
#include "CustomMsgBox.h"
#include "MainFrm.h"
#include "mboxview.h"

// CustomMsgBox dialog

IMPLEMENT_DYNAMIC(CustomMsgBox, CDialogEx)

CustomMsgBox::CustomMsgBox(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_CUSTOM_MSG_BOX, pParent)
{
	m_textFontHeight = 12;
	m_StatusBarHeight = 12;
}

CustomMsgBox::CustomMsgBox(LPCWSTR lpszText, LPCWSTR lpszCaption, UINT nType, int textFontHeight, CWnd* pParent)
	: CDialogEx(IDD_CUSTOM_MSG_BOX, pParent)
{
	m_textStr.Append(lpszText);
	captionStr.Append(lpszCaption);
	m_textFontHeight = textFontHeight;
	m_nType = nType;

	m_StatusBarHeight = 12;
}

CustomMsgBox::~CustomMsgBox()
{
}

void CustomMsgBox::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, ID_MSG_BOX_BUTTON_1, m_button0);
	DDX_Control(pDX, ID_MSG_BOX_BUTTON_2, m_button1);
	DDX_Control(pDX, ID_MSG_BOX_BUTTON_3, m_button2);
	DDX_Control(pDX, IDC_MSG_BOX_EDIT_TEXT, m_text);
	DDX_Control(pDX, IDC_MSG_BOX_ICON, m_icon);
}

void MB_BUTTON::Set(CButton& button, WORD id, int actionType, CString name, BOOL hidden)
{
	m_id = id;
	m_name = name;
	m_actionType = actionType;
	m_hidden = hidden;

	button.SetWindowText(name);
	button.EnableWindow(TRUE);
	button.ShowWindow(SW_NORMAL);
	if (hidden)
		button.ShowWindow(SW_HIDE);
}


BOOL CustomMsgBox::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here

	SetWindowText(captionStr);

	m_text.EnableWindow(TRUE);
	m_text.SetFont(&CMainFrame::m_dfltFont);

	// Create CStatusBar to create Gripper to make user aware of resing capability
	DWORD dwCtrlStatusStyle = WS_CHILD | WS_VISIBLE | CBRS_BOTTOM | CBRS_GRIPPER;
	if (!m_wndStatusBar.CreateEx(this, 0, dwCtrlStatusStyle))
	{
		TRACE0("Failed to create status bar\n");
		return -1;      // fail to create
	}

	BOOL retM = m_wndStatusBar.ModifyStyle(0, CBRS_GRIPPER);  // just is case

	m_wndStatusBar.GetStatusBarCtrl().SetSimple();  // no harm ??
	m_wndStatusBar.GetStatusBarCtrl().SetMinHeight(m_StatusBarHeight);

	// Calculate text rectangle size, width and height
	// CString fontName = CMainFrame::m_dfltFontName;

	BOOL bRepaint = TRUE; // ??


	// From .rc file
	// --------------

	CRect recDlgRc;
	GetWindowRect(&recDlgRc);

	int wRecDlgRc = recDlgRc.Width();
	int hRecDlgRc = recDlgRc.Height();

	CRect recTextRc;
	m_text.GetRect(&recTextRc);

	int wRecTextRc = recTextRc.Width();
	int hRecTextRc = recTextRc.Height();

	// CStatic icon is used to load standard MessageBox icon
	// ------------------------------------------------------

	m_icon.ModifyStyle(0, WS_CHILD | WS_VISIBLE | SS_ICON | SS_CENTERIMAGE );

	LPWSTR lpIcon = GetIconId(m_nType);

	BOOL winSystemIcon = TRUE;
	int lims = LIM_LARGE;
	CString errorText;
	HICON m_hIcon = LoadMsgBoxIcon(winSystemIcon, lpIcon, lims, errorText);
	if (m_hIcon == 0)
	{
		TRACE(L"%s\n", errorText);
		// ignore ??;
	}

	int cxIcon = GetSystemMetrics(SM_CXICON);
	int cyIcon = GetSystemMetrics(SM_CYICON);

	int SpacingSize = 16;

	CRect iconrect;

	int iconRecLeft = 0;
	int iconRecTop = 0;
	int iconRecRight = 0;
	int iconRecButtom = 0;

	iconRecLeft = SpacingSize;
	iconRecTop = recTextRc.top + SpacingSize;
	iconRecRight = iconRecLeft + cxIcon + SpacingSize;
	iconRecButtom = iconRecTop + cyIcon + SpacingSize;

	SetRect(&iconrect, iconRecLeft, iconRecTop, iconRecRight, iconRecButtom);

	HICON retIcon = m_icon.SetIcon(m_hIcon);

	m_icon.MoveWindow(iconrect, bRepaint);

	// Determine rectangle for translated text. Translation can return original english text
	// ------------------------------------------------------

	CClientDC dc(&m_text);

	HFONT hfntNew = CMainFrame::m_dfltFont.operator HFONT();

	// Determine text total length/width and height
	// Width will not be used

	HFONT hfntOld = (HFONT)dc.SelectObject(hfntNew);
	dc.SetTextColor(RGB(0, 0, 0));

	CSize textSize;
	if (m_textStr.GetLength() != 0)
		textSize = dc.GetTextExtent(m_textStr);
	else
		textSize = dc.GetTextExtent(L"XyZ");  // need text height, text length not used below

	dc.SelectObject(hfntOld);
	//

	int textLen = m_textStr.GetLength();
	CString longestWordText;
	int longetWordLen = LongestWordLength(m_textStr, &longestWordText);
	int wordLen = longestWordText.GetLength();

	CString longestLineText;
	int longetLineLen = LongestLineLength(m_textStr, &longestLineText);
	int lineLen = longestLineText.GetLength();

	ResHelper::TranslateString(m_textStr);
	m_textStr.Replace(L"\n", L"\r\n");

	// Best if we can determine longest word and line for translated text
	// This is not trivial for some languages

	int screenWidth = GetSystemMetrics(SM_CXSCREEN); // the width of the primary display.
	int screenHeight = GetSystemMetrics(SM_CYSCREEN); // the height of the primary display.

	// Decide on new text rectagle width, critical for other  calculations 

	// Need to decide on max width and possibly the longest word
	// User will have option to resize manually if desired

	int maxTextRectagleWidth = screenWidth / 3;
	int configTextRectagleWidth = 400;  // ???


	//Make sure we don't divide by zero
	if (CMainFrame::m_dfltFontSize > 0)
		configTextRectagleWidth = configTextRectagleWidth + (int)(0.15 * ((float)configTextRectagleWidth * CMainFrame::m_cnfFontSize) / CMainFrame::m_dfltFontSize);
	else
		; // 

	int textRectagleWidth = configTextRectagleWidth;
	if (textRectagleWidth > maxTextRectagleWidth)
		textRectagleWidth = maxTextRectagleWidth;

	// Aproximate, may need to do better later
	//Make sure we don't divide by zero
	if (longestWordText.GetLength())
		longetWordLen = (int)(((float)textSize.cx * longetWordLen) / longestWordText.GetLength());

	if (longestLineText.GetLength())
		longetLineLen = (int)(((float)textSize.cx * longetLineLen) / longestLineText.GetLength());


	if ((longetWordLen < textRectagleWidth) && (longetLineLen < textRectagleWidth))
	{
		textRectagleWidth = max(longetWordLen, longetLineLen);
	}

	m_text.SetWindowText(m_textStr);

	int textRecLeft = iconrect.right;
	int textRecTop = recTextRc.top + SpacingSize;
	int textRecRight = textRecLeft + textRectagleWidth;
	int textRecButtom = textRecTop + screenHeight / 2;

	CRect newTextRec;
	newTextRec.SetRect(textRecLeft, textRecTop, textRecRight, textRecButtom);

	m_text.SetRect(&newTextRec);
	m_text.MoveWindow(newTextRec, bRepaint);

	CRect newRecDlgRc;

	int newRecDlgRcLeft = recDlgRc.left;
	int newRecDlgRcTop = recDlgRc.top;
	int newRecDlgRcRight = recDlgRc.right + (newTextRec.Width() - recTextRc.Width()) + 32;
	int newRecDlgRcButtom = recDlgRc.bottom + newTextRec.Height();
	newRecDlgRc.SetRect(newRecDlgRcLeft, newRecDlgRcTop, newRecDlgRcRight, newRecDlgRcButtom);

	if (newRecDlgRc.Width() < recDlgRc.Width())
		newRecDlgRcRight += recDlgRc.Width() - newRecDlgRc.Width();

	newRecDlgRc.SetRect(newRecDlgRcLeft, newRecDlgRcTop, newRecDlgRcRight, newRecDlgRcButtom);

	MoveWindow(&newRecDlgRc, bRepaint);


	
	//if ((CmboxviewApp::m_isRTL == TRUE) && (CmboxviewApp::m_isRTLForDialogs))
	if (m_nType & MB_RTLREADING)
	{
		newTextRec.left -= 2 * SpacingSize;
		newTextRec.right -= 1 * SpacingSize;
		//newTextRec.DeflateRect(-2*SpacingSize, 0);

		m_text.SetRect(&newTextRec);
		m_text.MoveWindow(newTextRec, bRepaint);

		iconRecLeft = newRecDlgRc.Width() - (cxIcon + 2*SpacingSize);
		iconRecTop = newTextRec.top;
		iconRecRight = newRecDlgRc.Width() - 2*SpacingSize;
		iconRecButtom = iconRecTop + cyIcon + SpacingSize;

		SetRect(&iconrect, iconRecLeft, iconRecTop, iconRecRight, iconRecButtom);
		HICON retIcon = m_icon.SetIcon(m_hIcon);
		m_icon.MoveWindow(iconrect, bRepaint);
	}

	int lcnt = m_text.GetLineCount();

	// Need to do extra MoveWindow for proper scaling.
	lcnt++;
	int newTextHeigth = lcnt*textSize.cy;
	newRecDlgRcButtom = recDlgRc.bottom + (newTextHeigth - recTextRc.Height());

	newRecDlgRc.SetRect(newRecDlgRcLeft, newRecDlgRcTop, newRecDlgRcRight, newRecDlgRcButtom);
	MoveWindow(&newRecDlgRc, bRepaint);


	ProcessType(m_nType);
	ProcessType(m_nType);
	//SetFocus(); // focus is already set on one of the buttons

	ResHelper::LoadDialogItemsInfo(this);
	ResHelper::UpdateDialogItemsInfo(this);
	//BOOL retA = ResHelper::ActivateToolTips(this, m_toolTip);

	//return TRUE;  // return TRUE unless you set the focus to a control
	return FALSE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

HICON CustomMsgBox::LoadMsgBoxIcon(BOOL systemIcon, PCWSTR pszName, int lims, CString &errorText)
{
	HICON hIcon = 0;
	HINSTANCE hInst = 0;
	if (!systemIcon)
		hInst = ::GetModuleHandle(0);

	HRESULT retCode = LoadIconMetric(hInst, pszName, lims, &hIcon);
	if (retCode == S_OK)
	{
		_ASSERTE(hIcon);
		return hIcon;
	}
	else
	{
		errorText = FileUtils::GetLastErrorAsString();
		DWORD err = GetLastError();
		TRACE(L"(LoadMsgBoxIcon)error=%ld\n%s\n", err, errorText);
	}
	// Lets try LoadIcon
	// 	
	hIcon = ::LoadIcon(hInst, pszName);
	if (hIcon)
	{
		return hIcon;
	}
	else
	{
		errorText = FileUtils::GetLastErrorAsString();
		DWORD err = GetLastError();
		TRACE(L"(LoadMsgBoxIcon)error=%ld\n%s\n", err, errorText);
	}
	// Lets try LoadImage
	// 	
	int cxx = 0;
	int  cyy = 0;
	UINT fuLoad = LR_DEFAULTSIZE | LR_SHARED;
	hIcon = (HICON)::LoadImage(hInst, pszName, IMAGE_ICON, cxx, cyy, fuLoad);
	if (hIcon == 0)
	{
		errorText = FileUtils::GetLastErrorAsString();
		DWORD err = GetLastError();
		TRACE(L"(LoadMsgBoxIcon)error=%ld\n%s\n", err, errorText);
	}
	return hIcon;
}

BEGIN_MESSAGE_MAP(CustomMsgBox, CDialogEx)
	//ON_WM_CREATE()
	ON_WM_SIZE()
	ON_BN_CLICKED(ID_MSG_BOX_BUTTON_1, &CustomMsgBox::OnBnClickedMsgBoxButton1)
	ON_BN_CLICKED(ID_MSG_BOX_BUTTON_2, &CustomMsgBox::OnBnClickedMsgBoxButton2)
	ON_BN_CLICKED(ID_MSG_BOX_BUTTON_3, &CustomMsgBox::OnBnClickedMsgBoxButton3)
END_MESSAGE_MAP()


// CustomMsgBox message handlers

int CustomMsgBox::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CWnd::OnCreate(lpCreateStruct) == -1)
		return -1;

	return 0;
}

void CustomMsgBox::OnSize(UINT nType, int cx, int cy)
{
	CWnd::OnSize(nType, cx, cy);
	//TRACE(L"cx=%d cy=%d\n", cx, cy);

	// OnSize can be called before CStatusBar is created
	if (m_wndStatusBar.GetSafeHwnd())
	{
		CRect recWin;
		GetWindowRect(&recWin);
		int pos_cy = cy - m_StatusBarHeight;
		UINT nFlags = SWP_NOZORDER | SWP_SHOWWINDOW | SWP_NOSIZE;
		nFlags = SWP_NOZORDER;
		BOOL retPos = m_wndStatusBar.SetWindowPos(0, 0, pos_cy, cx, m_StatusBarHeight, nFlags);
		int deb = 1;
	}
}

int CustomMsgBox::LongestWordLength(CString& str, CString *longestWord)
{
	int n = str.GetLength();
	int res = 0, curr_len = 0, i = 0;
	int begPos = -1;
	int longestBegPos = 0;
	int longestEndPos = 0;
	wchar_t ch;

	for (i = 0; i < n; i++)
	{
		ch = str[i];
		if ((curr_len == 0) && ((ch == L'\n') || (ch == L'\r')))
		{
			// skip leading CR and NL
			continue;
		}
		else if (begPos == -1)
			begPos = i;

		// If current character is 
		// not end of current word.
		if ((ch != L' ') && (ch != L'\n') && (ch != L'\r'))
		{
			curr_len++;
		}
		else // end of word is found
		{
			if (curr_len > res)
			{
				res = curr_len;
				longestBegPos = begPos;
				longestEndPos = i;
				if (longestWord) {
					longestWord->Empty();
					longestWord->Append(&((LPCWSTR)str)[longestBegPos], res);
				}
			}
			begPos = -1;
			curr_len = 0;
		}
	}

	// We do max one more time to consider last word as there won't be any space after last word.
	if (curr_len > res)
	{
		res = curr_len;
		longestBegPos = begPos;
		longestEndPos = i;
		if (longestWord) {
			longestWord->Empty();
			longestWord->Append(&((LPCWSTR)str)[longestBegPos], res);
		}
	}
	return (res);
}

int CustomMsgBox::LongestLineLength(CString& str, CString *longestLine)
{
	int n = str.GetLength();
	int res = 0, curr_len = 0, i = 0;
	int begPos = -1;
	int longestBegPos = 0;
	int longestEndPos = 0;
	wchar_t ch;

	for (i = 0; i < n; i++)
	{
		ch = str[i];
		if ((curr_len == 0) && ((ch == L'\n') || (ch == L'\r')))
		{
			// skip leading CR and NL
			continue;
		}
		else if (begPos == -1)
			begPos = i;

		// If current character is 
		// not end of current line.
		if ((ch != L'\n') && (ch != L'\r'))
		{
			curr_len++;
		}
		else // end of line is found
		{
			if (curr_len > res)
			{
				res = curr_len;
				longestBegPos = begPos;
				longestEndPos = i;
				if (longestLine) {
					longestLine->Empty();
					longestLine->Append(&((LPCWSTR)str)[longestBegPos], res);
				}
			}
			begPos = -1;
			curr_len = 0;
		}
	}

	// We do max one more time to consider last line as there won't be any space after last word.
	if (curr_len > res)
	{
		res = curr_len;
		longestBegPos = begPos;
		longestEndPos = i;
		if (longestLine) {
			longestLine->Empty();
			longestLine->Append(&((LPCWSTR)str)[longestBegPos], res);
		}
	}
	return (res);
}

LPWSTR CustomMsgBox::GetIconId(UINT nType)
{
	LPWSTR iconId = IDI_EXCLAMATION;
	switch (nType & MB_ICONMASK)
	{
	case MB_ICONEXCLAMATION: iconId = IDI_EXCLAMATION; break;
	case MB_ICONHAND:        iconId = IDI_HAND;        break;
	case MB_ICONQUESTION:    iconId = IDI_QUESTION;    break;
	case MB_ICONASTERISK:    iconId = IDI_ASTERISK;    break;
	default:;
	}
	return iconId;
}

void CustomMsgBox::ProcessType(UINT nType)
{
	HWND h = GetSafeHwnd();

	m_buttons[0].Set(m_button0, ID_MSG_BOX_BUTTON_1, IDOK, L"", TRUE);
	m_buttons[1].Set(m_button1, ID_MSG_BOX_BUTTON_2, IDCANCEL, L"", TRUE);
	m_buttons[2].Set(m_button2, ID_MSG_BOX_BUTTON_3, IDRETRY, L"", TRUE);

	switch (nType & MB_TYPEMASK)
	{
	case MB_OK:
		m_buttons[0].Set(m_button0, ID_MSG_BOX_BUTTON_1, IDOK, L"Ok");
		SetDefaultButton(m_button0, ID_MSG_BOX_BUTTON_1, this);
		break;

	case MB_OKCANCEL:
		m_buttons[0].Set(m_button0, ID_MSG_BOX_BUTTON_1, IDCANCEL, L"Cancel");
		SetDefaultButton(m_button0, ID_MSG_BOX_BUTTON_1, this);
		break;

	case MB_YESNO:
		m_buttons[0].Set(m_button0, ID_MSG_BOX_BUTTON_1, IDNO, L"No");
		m_buttons[1].Set(m_button1, ID_MSG_BOX_BUTTON_2, IDYES, L"Yes");
		SetDefaultButton(m_button1, ID_MSG_BOX_BUTTON_2, this);
		break;

	case MB_YESNOCANCEL:
		m_buttons[0].Set(m_button0, ID_MSG_BOX_BUTTON_1, IDCANCEL, L"Cancel");
		m_buttons[1].Set(m_button1, ID_MSG_BOX_BUTTON_2, IDNO, L"No");
		m_buttons[2].Set(m_button2, ID_MSG_BOX_BUTTON_3, IDYES, L"Yes");
		SetDefaultButton(m_button2, ID_MSG_BOX_BUTTON_3, this);
		break;

	case MB_ABORTRETRYIGNORE:
		m_buttons[0].Set(m_button0, ID_MSG_BOX_BUTTON_1, IDIGNORE, L"Ignore");
		m_buttons[1].Set(m_button1, ID_MSG_BOX_BUTTON_2, IDRETRY, L"Retry");
		m_buttons[2].Set(m_button2, ID_MSG_BOX_BUTTON_3, IDABORT, L"Abort");
		SetDefaultButton(m_button2, ID_MSG_BOX_BUTTON_3, this);
		break;

	case MB_RETRYCANCEL:
		m_buttons[0].Set(m_button0, ID_MSG_BOX_BUTTON_1, IDCANCEL, L"Cancel");
		m_buttons[1].Set(m_button1, ID_MSG_BOX_BUTTON_2, IDRETRY, L"Retry");
		SetDefaultButton(m_button1, ID_MSG_BOX_BUTTON_2, this);
		break;

	case MB_CANCELTRYCONTINUE:
		m_buttons[0].Set(m_button0, ID_MSG_BOX_BUTTON_1, IDCONTINUE, L"Continue");
		m_buttons[1].Set(m_button1, ID_MSG_BOX_BUTTON_2, IDTRYAGAIN, L"Try Again");
		m_buttons[2].Set(m_button2, ID_MSG_BOX_BUTTON_3, IDCANCEL, L"Cancel");
		SetDefaultButton(m_button2, ID_MSG_BOX_BUTTON_3, this);
		break;

	default:
		m_buttons[0].Set(m_button0, ID_MSG_BOX_BUTTON_1, IDOK, L"Ok");
		SetDefaultButton(m_button0, ID_MSG_BOX_BUTTON_1, this);
		break;
	}

	//UpdateData(FromVariablesToControls);
}

MB_BUTTON* CustomMsgBox::FindButton(WORD id)
{
	int i = 0;
	for (i = 0; i < 3; i++)
	{
		if (!m_buttons[i].m_hidden && m_buttons[i].m_id == id)
			return &m_buttons[i];
	}
	return 0;
}

BOOL CustomMsgBox::IsDfltCButton(CButton& button)
{
	DWORD style = button.GetStyle();
	if ((style & BS_DEFPUSHBUTTON))
		return TRUE;
	else
		return FALSE;
}

CButton* CustomMsgBox::FindDfltCButton()
{
	// Def array of CButton in .h file ??
	BOOL isDflt = IsDfltCButton(m_button1);
	//_ASSERTE(isDflt); // se .rc file , currently no button set as default
	if (isDflt)
		return(&m_button1);
	isDflt = IsDfltCButton(m_button0);
	if (isDflt)
		return(&m_button0);
	isDflt = IsDfltCButton(m_button2);
	if (isDflt)
		return(&m_button2);

	return 0;
}

void CustomMsgBox::OnBnClickedMsgBoxButton1()
{
	// TODO: Add your control notification handler code here
	MB_BUTTON* button = FindButton(ID_MSG_BOX_BUTTON_1);

	EndDialog(button->m_actionType);
}

void CustomMsgBox::OnBnClickedMsgBoxButton2()
{
	// TODO: Add your control notification handler code here

	MB_BUTTON* button = FindButton(ID_MSG_BOX_BUTTON_2);

	EndDialog(button->m_actionType);
}

void CustomMsgBox::OnBnClickedMsgBoxButton3()
{
	// TODO: Add your control notification handler code here
	MB_BUTTON* button = FindButton(ID_MSG_BOX_BUTTON_3);

	EndDialog(button->m_actionType);
}

void CustomMsgBox::SetDefaultButton(CButton& newDflt, WORD id, CDialogEx *dlg)
{
	BOOL retM = newDflt.ModifyStyle(0, BS_DEFPUSHBUTTON);
	dlg->SetDefID(id);
	newDflt.SetFocus();
	return;
}

void TestCustomMsgBox()
{
	//CustomMsgBox  mbox;
	CString text = L"This dialog enables users to configure the root data folder."
		" MBox Viewer will create a UMBoxViewer folder under the configured root data folder"
		" for all files generated by MBox Viewer such as mails exported in PDF format.\n\n"
		"IMPORTANT: Data Folder path should be as short as possible to avoid truncation of names of files generated by MBox Viewer."
		" Note that the file path is limited to 255 characters.";

	//text = L"";

	CString caption = L"Info";
	//UINT nType = MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO;
	//UINT nType = MB_APPLMODAL | MB_ICONHAND | MB_ABORTRETRYIGNORE;
	//UINT nType = MB_APPLMODAL | MB_ICONHAND | MB_YESNOCANCEL;
	UINT nType = MB_APPLMODAL | MB_ICONHAND | MB_CANCELTRYCONTINUE;
	
	//UINT nType = MB_APPLMODAL | MB_ICONHAND | MB_YESNOCANCEL| MB_RTLREADING;
	

	int textFontHeight = 16;

	CWnd * pParent = 0;
	CustomMsgBox mbox(text, caption, nType, textFontHeight, pParent);

	INT_PTR code = mbox.DoModal();

	if (code == IDOK)
	{
		int deb = 1;
	}
	else
	{
		int deb = 1;
	}

	HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
	int result = MessageBox(h, text, caption, nType);
	if (result == IDOK)
	{
		int deb = 1;
	}
	else
	{
		int deb = 1;
	}
}
