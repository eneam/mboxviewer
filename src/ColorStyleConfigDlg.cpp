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

// ColorStyleConfigDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ColorStyleConfigDlg.h"
#include "afxdialogex.h"
#include "ColorPickerDlg.h"
#include "mboxview.h"
#include "MainFrm.h"
#include "NListView.h"
#include "ResHelper.h"


// ColorStyleConfigDlg dialog

IMPLEMENT_DYNAMIC(ColorStyleConfigDlg, CDialogEx)

static void CSCAssert()
{
#ifdef _DEBUG
	_ASSERTE(FALSE);
#endif
	int deb = 1;
}

// Created in CManFrame and deleted upon MBox Viewer shutdown
// Using SW_SHOW and SW_HIDE

ColorStyleConfigDlg::ColorStyleConfigDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_COLOR_STYLE_DLG, pParent)
{
	m_MainFrameWnd = pParent;
	m_buttonColor = RGB(255, 255, 255);
	m_selectedColorStyle = 0;
	m_lastColorStyle = 0;
	m_selectedPane = 0;
	m_applyColorToSelectedPane = 0;
	m_applyColorToAllPanes = 0;
	for (int i = 0; i < 16; i++)
	{
		m_customColors[i] = RGB(210, 210, 210);
	}
}

ColorStyleConfigDlg::~ColorStyleConfigDlg()
{
	int deb = 1;
}

void ColorStyleConfigDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_listBox);
	DDX_Radio(pDX, IDC_COLOR_DFLT, m_selectedColorStyle);
	DDX_Check(pDX, IDC_COLOR_CHECK, m_applyColorToSelectedPane);
	DDX_Check(pDX, IDC_COLOR_CHECK_ALL, m_applyColorToAllPanes);
}


BEGIN_MESSAGE_MAP(ColorStyleConfigDlg, CDialogEx)
	ON_WM_CLOSE()
	ON_NOTIFY(NM_CUSTOMDRAW, IDC_COLOR_BUTTON, &ColorStyleConfigDlg::OnNMCustomdrawColorButton)
	ON_BN_CLICKED(IDC_COLOR_BUTTON, &ColorStyleConfigDlg::OnBnClickedColorButton)
	ON_BN_CLICKED(IDC_COLOR_CHECK, &ColorStyleConfigDlg::OnBnClickedColorCheck)
	ON_LBN_SELCHANGE(IDC_LIST1, &ColorStyleConfigDlg::OnLbnSelchangeList1)
	ON_BN_CLICKED(IDOK, &ColorStyleConfigDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &ColorStyleConfigDlg::OnBnClickedCancel)
	ON_COMMAND_RANGE(IDC_COLOR_DFLT, IDC_COLOR_STYLE_16, &ColorStyleConfigDlg::OnBnClickedColorStyle)
	ON_BN_CLICKED(IDC_BUTTON_HELP, &ColorStyleConfigDlg::OnBnClickedButtonHelp)
	ON_BN_CLICKED(IDC_BUTTON_CLOSE, &ColorStyleConfigDlg::OnBnClickedButtonClose)
	ON_BN_CLICKED(IDC_COLOR_CHECK_ALL, &ColorStyleConfigDlg::OnBnClickedApplyToAllPanes)
	ON_NOTIFY_EX(TTN_NEEDTEXT, 0, &CFindDlg::OnTtnNeedText)
END_MESSAGE_MAP()


// ColorStyleConfigDlg message handlers

int ColorStyleConfigDlg::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	return 0;
}

void ColorStyleConfigDlg::OnClose()
{
	// TODO: Add your message handler code here and/or call default

	int deb = 1;

	ShowWindow(SW_HIDE);

	//CWnd::OnClose();
}

void ColorStyleConfigDlg::LoadData()
{
	m_listBox.AddString(L"Mail Archive Tree");
	m_listBox.AddString(L"Mail Summary");
	m_listBox.AddString(L"Mail Message");
	m_listBox.AddString(L"Mail Attachments");
	m_listBox.AddString(L"Mail Summary Column Titles");
	m_listBox.AddString(L"Mail MessageHeader Fields");
	m_listBox.AddString(L"Mail Conversation1");
	m_listBox.AddString(L"Mail Conversation2");

	ColorStyleConfig &customColorStyle = CMainFrame::m_ColorStylesDB.m_customColorStyles;

	customColorStyle.SetColor(ColorStyleConfig::MailArchiveList, RGB(0, 255, 0));
	customColorStyle.SetColor(ColorStyleConfig::MailSummaryList, RGB(255, 255, 0));
	customColorStyle.SetColor(ColorStyleConfig::MailMessage, RGB(255, 128, 64));
	customColorStyle.SetColor(ColorStyleConfig::MailMessageAttachments, RGB(255, 128, 255));
	customColorStyle.SetColor(ColorStyleConfig::MailSummaryTitles, RGB(255, 0, 0));
	customColorStyle.SetColor(ColorStyleConfig::MailMessageHeader, RGB(0, 255, 255));
	customColorStyle.SetColor(ColorStyleConfig::MailConversion1, RGB(245, 192, 195));
	customColorStyle.SetColor(ColorStyleConfig::MailConversion2, RGB(205, 192, 196));

	m_selectedColorStyle = ColorStylesDB::ColorDefault;
	m_lastColorStyle = ColorStylesDB::ColorDefault;

	// Will override m_selectedColorStyle and customColorStyle.m_ColorTbl
	LoadFromRegistry();

	int deb = 1;
}

BOOL ColorStyleConfigDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here
	RECT rcDlg;

	HWND hwndOwner = ::GetDesktopWindow();
	CWnd *wndOwner = GetDesktopWindow();
	::GetWindowRect(hwndOwner, &rcDlg);

	CRect rc;
	GetClientRect(&rc);

	SetWindowPos(
		0,
		rcDlg.right - rc.Width() - 200,
		rcDlg.top + 200,
		0, 
		0,
		SWP_SHOWWINDOW|SWP_NOSIZE);

	m_listBox.SetFont(GetFont());

	LoadData();

	WORD nID = ColorStylesDB::NUM2ID(m_selectedColorStyle);

	OnBnClickedColorStyle(nID);

	CheckRadioButton(IDC_COLOR_DFLT, IDC_COLOR_STYLE_16, nID);

	ResHelper::LoadDialogItemsInfo(this);
	ResHelper::UpdateDialogItemsInfo(this);
	BOOL retA = ResHelper::ActivateToolTips(this, m_toolTip);

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}

void ColorStyleConfigDlg::OnNMCustomdrawColorButton(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMCUSTOMDRAW pNMCD = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);
	// TODO: Add your control notification handler code here

	*pResult = 0;

	switch (pNMCD->dwDrawStage)
	{
	case CDDS_PREERASE:
	{
		CDC dc;
		CRect rect = pNMCD->rc;
		HDC hDC = pNMCD->hdc;

		if (dc.Attach(hDC))
		{
			SetBkMode(dc, TRANSPARENT);
			dc.FillRect(&rect, &CBrush(m_buttonColor));
			::FrameRect(hDC, rect, (HBRUSH)GetStockObject(BLACK_BRUSH));

			dc.DrawText(L"Pick Color", &rect, DT_SINGLELINE|DT_CENTER|DT_VCENTER);

			dc.Detach();
		}
		*pResult = CDRF_SKIPDEFAULT;
		int deb = 1;
	}
	break;
	default:
	{
		*pResult = CDDS_PREERASE;
		int deb = 1;  // should not be here
	}
	}
	int deb = 1;
}

void ColorStyleConfigDlg::OnBnClickedColorButton()
{
	// TODO: Add your control notification handler code here

	ColorStyleConfig &colorStyles = CMainFrame::m_ColorStylesDB.m_colorStyles;
	ColorStyleConfig &customColorStyles = CMainFrame::m_ColorStylesDB.m_customColorStyles;

	COLORREF color = RGB(0, 255, 0);
	int nSel = m_listBox.GetCurSel();

	if (nSel != m_selectedPane)
	{
		CSCAssert();
		//return;
	}

	if ((nSel < 0) || (nSel >= ColorStyleConfig::MailMaxPanes) || (nSel >= m_listBox.GetCount()))
	{
		; // done
	}
	else
	{
		color = colorStyles.GetColor(nSel);
	}
	
	DWORD Flags = CC_FULLOPEN | CC_ANYCOLOR;
	ColorPickerDlg dlg(color, Flags, this);

	for (int i = 0; i < ColorStyleConfig::MailMaxPanes; i++)
	{
		m_customColors[i] = colorStyles.GetColor(i);
	}

	dlg.m_cc.lpCustColors = &m_customColors[0];

	INT_PTR ret = dlg.DoModal();
	if (ret == IDCANCEL)
	{
		return;
	}

	DWORD retDlg = CommDlgExtendedError();
	COLORREF *list = dlg.GetSavedCustomColors();
	m_buttonColor = dlg.GetColor();


	CWnd* p;
#if 1
	ClearCustomStyleChecks();
#else
	p = GetDlgItem(IDC_COLOR_CHECK);
	if (p)
	{
		int chk = ((CButton*)p)->GetCheck();
		if (chk)
			OnBnClickedColorCheck();
	}
#endif

	p = GetDlgItem(IDC_COLOR_BUTTON);
	if (p)
	{
		p->RedrawWindow();
	}

	CStringA colorStr;
	NListView::Color2Str(m_buttonColor, colorStr);

	int deb = 1;
}


void ColorStyleConfigDlg::OnBnClickedColorCheck()
{
	// TODO: Add your control notification handler code here


	// CMainFrame::m_ColorStylesDB.m_colorStyles is used as working array
	// until it is saved
	ColorStyleConfig &colorStyles = CMainFrame::m_ColorStylesDB.m_colorStyles;
	ColorStyleConfig &customColorStyles = CMainFrame::m_ColorStylesDB.m_customColorStyles;

	CWnd *p = GetDlgItem(IDC_COLOR_CHECK);
	if (m_buttonColor == customColorStyles.m_ColorTbl[m_selectedPane])
	{
		m_applyColorToSelectedPane = 0;
		if (p)
		{
			((CButton*)p)->SetCheck(m_applyColorToSelectedPane);
		}
		return;
	}

	int chk = 0;
	if (p)
	{
		chk = ((CButton*)p)->GetCheck();
		if (chk)
			m_applyColorToSelectedPane = 1;
		else
			m_applyColorToSelectedPane = 0;
	}
	int nSel = m_listBox.GetCurSel();
	if ((nSel < 0) || (nSel > ColorStyleConfig::MailMaxPanes))
	{
		CSCAssert();
		//return;
	}
	if (nSel != m_selectedPane)
	{
		CSCAssert();
		//return;
	}

	if (m_applyColorToSelectedPane)   // "Apply to Selected Pane" check button set
	{
		colorStyles.SetColor(m_selectedPane, m_buttonColor);
		int deb = 1;
	}
	else   // "Apply to Selected Pane" check button unset
	{
		// restore color from CMainFrame::m_ColorStylesDB.m_customColorStyles
		m_buttonColor = customColorStyles.GetColor(m_selectedPane);
		colorStyles.SetColor(m_selectedPane, m_buttonColor);
		int deb = 1;
	}

	p = GetDlgItem(IDC_COLOR_BUTTON);
	if (p) 
	{
		p->RedrawWindow();
	}
	if (m_MainFrameWnd)
		m_MainFrameWnd->PostMessage(WM_CMD_PARAM_NEW_COLOR_MESSAGE, ColorStylesDB::ColorCustom, 0);

	int deb = 1;
}


void ColorStyleConfigDlg::OnLbnSelchangeList1()
{
	// TODO: Add your control notification handler code here

	ColorStyleConfig &colorStyles = CMainFrame::m_ColorStylesDB.m_colorStyles;
	ColorStyleConfig &customColorStyles = CMainFrame::m_ColorStylesDB.m_customColorStyles;

	int nSel = m_listBox.GetCurSel();

	if ((nSel < 0) || (nSel >= ColorStyleConfig::MailMaxPanes) || (nSel >= m_listBox.GetCount()))
	{
		CSCAssert();
		return;
	}

	CString ItemSelected;
	m_listBox.GetText(nSel, ItemSelected);
	m_selectedPane = nSel;

	m_buttonColor = colorStyles.GetColor(nSel);

	CWnd *p = GetDlgItem(IDC_COLOR_BUTTON);
	if (p) 
	{
		p->RedrawWindow();
	}

#if 1
	ClearCustomStyleChecks();
#else
	CWnd* pall = GetDlgItem(IDC_COLOR_CHECK_ALL);
	p = GetDlgItem(IDC_COLOR_CHECK);
	if (p) 
	{
		if (customColorStyles.GetColor(nSel) == colorStyles.GetColor(nSel))
		{
			m_applyColorToSelectedPane = 0;
			((CButton*)p)->SetCheck(0);
			((CButton*)pall)->SetCheck(0);
		}
		else
		{
			m_applyColorToSelectedPane = 1;
			((CButton*)p)->SetCheck(1);
			((CButton*)pall)->SetCheck(1);
		}
	}
#endif

	int deb = 1;
}

ColorStyleConfig::ColorStyleConfig()
{
	SetDefaults();
}

void ColorStyleConfig::Copy(ColorStyleConfig &colorConfig)
{
	CopyArray(colorConfig.m_ColorTbl, sizeof(colorConfig.m_ColorTbl) / sizeof(DWORD));
	m_ReplaceAllWhiteBackgrounTags = colorConfig.m_ReplaceAllWhiteBackgrounTags;
}

void ColorStyleConfig::SetDefaults()
{
	int i;
	for (i = 0; i < MailMaxPanes; i++)
	{
		m_ColorTbl[i] = COLOR_WHITE;
	}
	m_ColorTbl[MailConversion1] = PeachPuff1;
	m_ColorTbl[MailConversion2] = AntiqueWhite3;
}

void ColorStyleConfig::SetColorStyle(int colorStyle)
{
	int i;
	if (colorStyle == ColorStylesDB::ColorDefault)
	{
		SetDefaults();
	}
	else if (colorStyle == ColorStylesDB::ColorCustom)
	{
		Copy(CMainFrame::m_ColorStylesDB.m_customColorStyles);
	}
	else if (colorStyle == ColorStylesDB::ColorStyle1)
	{
		for (i = 0; i < MailMaxPanes; i++)
		{
			m_ColorTbl[i] = RGB(240, 240, 240);
		}
		//
		m_ColorTbl[MailConversion1] = RGB(245, 192, 195);
		m_ColorTbl[MailConversion2] = RGB(205, 192, 196);
	}
	else if (colorStyle == ColorStylesDB::ColorStyle2)
	{
		for (i = 0; i < MailMaxPanes; i++)
		{
			m_ColorTbl[i] = RGB(230, 230, 230);
		}
		//
		m_ColorTbl[MailConversion1] = RGB(245, 192, 195);
		m_ColorTbl[MailConversion2] = RGB(205, 192, 196);
	}
	else if (colorStyle == ColorStylesDB::ColorStyle3)
	{
		for (i = 0; i < MailMaxPanes; i++)
		{
			m_ColorTbl[i] = RGB(215, 215, 215);
		}
		//
		m_ColorTbl[MailConversion1] = RGB(245, 192, 195);
		m_ColorTbl[MailConversion2] = RGB(205, 192, 196);
	}
	else if (colorStyle == ColorStylesDB::ColorStyle4)
	{
		for (i = 0; i < MailMaxPanes; i++)
		{
			m_ColorTbl[i] = RGB(219, 215, 210);  // good
		}
		//
		m_ColorTbl[MailConversion1] = RGB(245, 192, 195);
		m_ColorTbl[MailConversion2] = RGB(205, 192, 196);
	}
	else if (colorStyle == ColorStylesDB::ColorStyle5)
	{
		for (i = 0; i < MailMaxPanes; i++)
		{
			m_ColorTbl[i] = RGB(205, 192, 186);
		}
		m_ColorTbl[MailConversion1] = RGB(245, 192, 195);
		m_ColorTbl[MailConversion2] = RGB(205, 192, 196);
	}
	else if (colorStyle == ColorStylesDB::ColorStyle6)
	{
		for (i = 0; i < MailMaxPanes; i++)
		{
			m_ColorTbl[i] = RGB(205, 192, 176);
		}
		m_ColorTbl[MailConversion1] = RGB(245, 192, 195);
		m_ColorTbl[MailConversion2] = RGB(205, 192, 196);
	}
	else if (colorStyle == ColorStylesDB::ColorStyle7)
	{
		for (i = 0; i < MailMaxPanes; i++)
		{
			m_ColorTbl[i] = RGB(205, 192, 176);
		}
		m_ColorTbl[MailSummaryTitles] = RGB(205, 192, 196);
		//m_ColorTbl[MailMessageHeader] = RGB(205, 192, 196);
		m_ColorTbl[MailConversion1] = RGB(245, 192, 195);
		m_ColorTbl[MailConversion2] = RGB(205, 192, 196);
	}
	else if (colorStyle == ColorStylesDB::ColorStyle8)
	{
		for (i = 0; i < MailMaxPanes; i++)
		{
			m_ColorTbl[i] = RGB(245, 222, 179);  // wheat
		}
		m_ColorTbl[MailConversion1] = RGB(245, 192, 195);
		m_ColorTbl[MailConversion2] = RGB(205, 192, 196);
	}
	else if (colorStyle == ColorStylesDB::ColorStyle9)
	{
		for (i = 0; i < MailMaxPanes; i++)
		{
			m_ColorTbl[i] = RGB(200, 200, 200);  // wheat
		}
		m_ColorTbl[MailConversion1] = RGB(245, 192, 195);
		m_ColorTbl[MailConversion2] = RGB(205, 192, 196);
	}
	else if (colorStyle == ColorStylesDB::ColorStyle10)
	{
		for (i = 0; i < MailMaxPanes; i++)
		{
			m_ColorTbl[i] = RGB(196, 199, 141);  // wheat
		}
		m_ColorTbl[MailConversion1] = RGB(245, 192, 195);
		m_ColorTbl[MailConversion2] = RGB(205, 192, 196);
	}
	else if (colorStyle == ColorStylesDB::ColorStyle11)
	{
		for (i = 0; i < MailMaxPanes; i++)
		{
			m_ColorTbl[i] = RGB(227, 233, 180);  // wheat
		}
		m_ColorTbl[MailConversion1] = RGB(245, 192, 195);
		m_ColorTbl[MailConversion2] = RGB(205, 192, 196);
	}
	else if (colorStyle == ColorStylesDB::ColorStyle12)
	{
		for (i = 0; i < MailMaxPanes; i++)
		{
			m_ColorTbl[i] = RGB(237, 238, 181);  // wheat
		}
		m_ColorTbl[MailConversion1] = RGB(245, 192, 195);
		m_ColorTbl[MailConversion2] = RGB(205, 192, 196);
	}
	else if (colorStyle == ColorStylesDB::ColorStyle13)
	{
		for (i = 0; i < MailMaxPanes; i++)
		{
			m_ColorTbl[i] = RGB(194, 201, 199);  // wheat
		}
		m_ColorTbl[MailConversion1] = RGB(245, 192, 195);
		m_ColorTbl[MailConversion2] = RGB(205, 192, 196);
	}
	else if (colorStyle == ColorStylesDB::ColorStyle14)
	{
		for (i = 0; i < MailMaxPanes; i++)
		{
			m_ColorTbl[i] = RGB(222, 223, 196);  // wheat
		}
		m_ColorTbl[MailConversion1] = RGB(245, 192, 195);
		m_ColorTbl[MailConversion2] = RGB(205, 192, 196);
	}
	else if (colorStyle == ColorStylesDB::ColorStyle15)
	{
		for (i = 0; i < MailMaxPanes; i++)
		{
			m_ColorTbl[i] = RGB(201, 217, 193);  // wheat
		}
		m_ColorTbl[MailConversion1] = RGB(245, 192, 195);
		m_ColorTbl[MailConversion2] = RGB(205, 192, 196);
	}
	else if (colorStyle == ColorStylesDB::ColorStyle16)
	{
		for (i = 0; i < MailMaxPanes; i++)
		{
			m_ColorTbl[i] = RGB(235, 222, 207);  // wheat
		}
		m_ColorTbl[MailConversion1] = RGB(202, 180, 53);
		m_ColorTbl[MailConversion2] = RGB(205, 192, 196);
	}
	int deb = 1;
}

void ColorStyleConfig::CopyArray(DWORD *table, int tableLength)
{
	SetDefaults();

	if (tableLength < 0)
	{
		CSCAssert();
		return;
	}

	int cnt = tableLength;
	if (tableLength > MailMaxPanes)
		cnt = MailMaxPanes;

	int i;
	for (i = 0; i < cnt; i++)
	{
		m_ColorTbl[i] = table[i];
	}
}
DWORD ColorStyleConfig::GetColor(int index)
{
	if (index < MailMaxPanes)
		return m_ColorTbl[index];
	else
		return COLOR_WHITE;
}

DWORD ColorStyleConfig::SetColor(int index, DWORD color)
{
	if (index < MailMaxPanes)
		return m_ColorTbl[index] = color;
	else
		return COLOR_WHITE;
}

void ColorStyleConfig::ColorArray2String(CString &str)
{
	int i;
	CString colorStr;
	CString colorStrInt;
	DWORD color;

	str.Empty();
	for (i = 0; i < MailMaxPanes; i++)
	{
		color = m_ColorTbl[i];
		colorStr.Format(L"0x%06x", color);
		str.Append(colorStr);
	}
}

void ColorStyleConfig::String2ColorArray(CString &str)
{
	CString colorStr;
	DWORD color;
	int i = 0;
	int j = 0;
	for (j = 0; j < MailMaxPanes; j++)
	{
		if ((i + 8) > str.GetLength())
		{
			CSCAssert();
			break;
		}
		colorStr = str.Mid(i, 8);
		color = _tcstoul(colorStr, 0, 16);
		m_ColorTbl[j] = color;
		i += 8;
	}
}

ColorStylesDB::ColorStylesDB()
{
	m_colorStyles.SetDefaults();
}

void ColorStyleConfigDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	//UpdateData(FALSE); // data -> control
	//UpdateData(TRUE); // control -> data

	ColorStyleConfig &colorStyles = CMainFrame::m_ColorStylesDB.m_colorStyles;
	ColorStyleConfig &customColorStyles = CMainFrame::m_ColorStylesDB.m_customColorStyles;
	
	if (m_selectedColorStyle == ColorStylesDB::ColorCustom)
	{
		customColorStyles.Copy(colorStyles);
		m_applyColorToSelectedPane = 0;		
		ClearCustomStyleChecks();
	}
	m_lastColorStyle = m_selectedColorStyle;

	UpdateRegistry();
}

void ColorStyleConfigDlg::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here

	ColorStyleConfig &colorStyles = CMainFrame::m_ColorStylesDB.m_colorStyles;
	ColorStyleConfig &customColorStyles = CMainFrame::m_ColorStylesDB.m_customColorStyles;

	m_selectedColorStyle = m_lastColorStyle;
	colorStyles.SetColorStyle(m_selectedColorStyle);

	m_applyColorToSelectedPane = 0; // FIXME
	ClearCustomStyleChecks();

	m_buttonColor = colorStyles.GetColor(m_selectedPane);

	WORD nID = ColorStylesDB::NUM2ID(m_selectedColorStyle);

	CheckRadioButton(IDC_COLOR_DFLT, IDC_COLOR_STYLE_16, nID);

	BOOL enableColorSelection = FALSE;
	if (m_selectedColorStyle == ColorStylesDB::ColorCustom)
	{
		enableColorSelection = TRUE;
		m_listBox.SetCurSel(m_selectedPane);
		m_buttonColor = colorStyles.GetColor(m_selectedPane);
	}
	else
	{
		m_listBox.SetCurSel(-1);
		m_buttonColor = RGB(255, 255, 255);
	}

	CWnd *p = GetDlgItem(IDC_COLOR_BUTTON);
	if (p) {
		p->RedrawWindow();
	}

	EnableCustomStyle(enableColorSelection);

	if (m_MainFrameWnd)
		m_MainFrameWnd->PostMessage(WM_CMD_PARAM_NEW_COLOR_MESSAGE, m_selectedColorStyle, 0);
}

void ColorStyleConfigDlg::PostNcDestroy()
{
	// TODO: Add your specialized code here and/or call the base class
	DestroyWindow();
	CDialogEx::PostNcDestroy();
}

int ColorStylesDB::ID2ENUM(WORD nID)
{
	int colorStyle = -1;
	if (nID == IDC_COLOR_DFLT)
		colorStyle = ColorStylesDB::ColorDefault;
	else if (nID == IDC_COLOR_CUSTOM)
		colorStyle = ColorStylesDB::ColorCustom;
	else if (nID == IDC_COLOR_STYLE1)
		colorStyle = ColorStylesDB::ColorStyle1;
	else if (nID == IDC_COLOR_STYLE2)
		colorStyle = ColorStylesDB::ColorStyle2;
	else if (nID == IDC_COLOR_STYLE3)
		colorStyle = ColorStylesDB::ColorStyle3;
	else if (nID == IDC_COLOR_STYLE4)
		colorStyle = ColorStylesDB::ColorStyle4;
	else if (nID == IDC_COLOR_STYLE5)
		colorStyle = ColorStylesDB::ColorStyle5;
	else if (nID == IDC_COLOR_STYLE6)
		colorStyle = ColorStylesDB::ColorStyle6;
	else if (nID == IDC_COLOR_STYLE7)
		colorStyle = ColorStylesDB::ColorStyle7;
	else if (nID == IDC_COLOR_STYLE8)
		colorStyle = ColorStylesDB::ColorStyle8;
	else if (nID == IDC_COLOR_STYLE_9)
		colorStyle = ColorStylesDB::ColorStyle9;
	else if (nID == IDC_COLOR_STYLE_10)
		colorStyle = ColorStylesDB::ColorStyle10;
	else if (nID == IDC_COLOR_STYLE_11)
		colorStyle = ColorStylesDB::ColorStyle11;
	else if (nID == IDC_COLOR_STYLE_12)
		colorStyle = ColorStylesDB::ColorStyle12;
	else if (nID == IDC_COLOR_STYLE_13)
		colorStyle = ColorStylesDB::ColorStyle13;
	else if (nID == IDC_COLOR_STYLE_14)
		colorStyle = ColorStylesDB::ColorStyle14;
	else if (nID == IDC_COLOR_STYLE_15)
		colorStyle = ColorStylesDB::ColorStyle15;
	else if (nID == IDC_COLOR_STYLE_16)
		colorStyle = ColorStylesDB::ColorStyle16;

	return colorStyle;
}

WORD ColorStylesDB::NUM2ID(int colorStyle)
{
	static WORD ColorStyle2ID[] = { 
		IDC_COLOR_DFLT , IDC_COLOR_CUSTOM ,
		IDC_COLOR_STYLE1, IDC_COLOR_STYLE2  ,IDC_COLOR_STYLE3 ,
		IDC_COLOR_STYLE4 , IDC_COLOR_STYLE5 , IDC_COLOR_STYLE6 , IDC_COLOR_STYLE7 , IDC_COLOR_STYLE8,
		IDC_COLOR_STYLE_9, IDC_COLOR_STYLE_10  ,IDC_COLOR_STYLE_11 ,
		IDC_COLOR_STYLE_12 , IDC_COLOR_STYLE_13 , IDC_COLOR_STYLE_14 , IDC_COLOR_STYLE_15 , IDC_COLOR_STYLE_16
	};

	WORD nID = IDC_COLOR_DFLT;

	if (colorStyle == ColorStylesDB::ColorDefault)
		nID = IDC_COLOR_DFLT;
	else if (colorStyle == ColorStylesDB::ColorCustom)
		nID = IDC_COLOR_CUSTOM;
	else if (colorStyle == ColorStylesDB::ColorStyle1)
		nID = IDC_COLOR_STYLE1;
	else if (colorStyle == ColorStylesDB::ColorStyle2)
		nID = IDC_COLOR_STYLE2;
	else if (colorStyle == ColorStylesDB::ColorStyle3)
		nID = IDC_COLOR_STYLE3;
	else if (colorStyle == ColorStylesDB::ColorStyle4)
		nID = IDC_COLOR_STYLE4;
	else if (colorStyle == ColorStylesDB::ColorStyle5)
		nID = IDC_COLOR_STYLE5;
	else if (colorStyle == ColorStylesDB::ColorStyle6)
		nID = IDC_COLOR_STYLE6;
	else if (colorStyle == ColorStylesDB::ColorStyle7)
		nID = IDC_COLOR_STYLE7;
	else if (colorStyle == ColorStylesDB::ColorStyle8)
		nID = IDC_COLOR_STYLE8;
	else if (colorStyle == ColorStylesDB::ColorStyle9)
		nID = IDC_COLOR_STYLE_9;
	else if (colorStyle == ColorStylesDB::ColorStyle10)
		nID = IDC_COLOR_STYLE_10;
	else if (colorStyle == ColorStylesDB::ColorStyle11)
		nID = IDC_COLOR_STYLE_11;
	else if (colorStyle == ColorStylesDB::ColorStyle12)
		nID = IDC_COLOR_STYLE_12;
	else if (colorStyle == ColorStylesDB::ColorStyle13)
		nID = IDC_COLOR_STYLE_13;
	else if (colorStyle == ColorStylesDB::ColorStyle14)
		nID = IDC_COLOR_STYLE_14;
	else if (colorStyle == ColorStylesDB::ColorStyle15)
		nID = IDC_COLOR_STYLE_15;
	else if (colorStyle == ColorStylesDB::ColorStyle16)
		nID = IDC_COLOR_STYLE_16;

	return nID;
}

// Process selected radio button from Default, Custom to ColorStyle8 range
void ColorStyleConfigDlg::OnBnClickedColorStyle(UINT nID)
{
	// TODO: Add your specialized code here and/or call the base class

	ColorStyleConfig &colorStyles = CMainFrame::m_ColorStylesDB.m_colorStyles;
	ColorStyleConfig &customColorStyles = CMainFrame::m_ColorStylesDB.m_customColorStyles;

	int colorStyle = ColorStylesDB::ID2ENUM(nID);

	if (colorStyle < 0)
		return;

	ClearCustomStyleChecks();

	if ((m_selectedColorStyle == ColorStylesDB::ColorCustom) && (colorStyle != ColorStylesDB::ColorCustom))
	{
		int i;
		for (i = 0; i < ColorStyleConfig::MailMaxPanes; i++)
		{
			if (colorStyles.m_ColorTbl[i] != customColorStyles.m_ColorTbl[i])
			{
				CString txt = L"Custom Color Style was changed but not saved! Discard changes?";
				int answer = MessageBox(txt, L"Warning", MB_APPLMODAL | MB_ICONWARNING | MB_YESNO);
				if (answer == IDNO)
				{
					CheckRadioButton(IDC_COLOR_DFLT, IDC_COLOR_STYLE_16, IDC_COLOR_CUSTOM);

					m_applyColorToSelectedPane = 1;
					CWnd *p = GetDlgItem(IDC_COLOR_CHECK);
					if (p)
					{
						((CButton*)p)->SetCheck(m_applyColorToSelectedPane);
					}
					return;
				}
				else
				{
					m_applyColorToSelectedPane = 0;
					CWnd *p = GetDlgItem(IDC_COLOR_CHECK);
					if (p)
					{
						((CButton*)p)->SetCheck(m_applyColorToSelectedPane);
					}
					break;  // all done
				}
			}
		}
		int deb = 1;
	}

	colorStyles.SetColorStyle(colorStyle);

	BOOL enableColorSelection = FALSE;
	if (colorStyle == ColorStylesDB::ColorCustom)
	{
		enableColorSelection = TRUE;
		m_listBox.SetCurSel(m_selectedPane);
		m_buttonColor = colorStyles.GetColor(m_selectedPane);
	}
	else
	{
		m_listBox.SetCurSel(-1);
		m_buttonColor = RGB(255, 255, 255);
	}

	CWnd *p = GetDlgItem(IDC_COLOR_BUTTON);
	if (p) {
		p->RedrawWindow();
	}

	EnableCustomStyle(enableColorSelection);

	m_selectedColorStyle = colorStyle;
	//UpdateData(TRUE);

	if (m_MainFrameWnd)
		m_MainFrameWnd->PostMessage(WM_CMD_PARAM_NEW_COLOR_MESSAGE, colorStyle, 0);

	int deb = 1;
}


// Enable/Disbale Custom Color Style Configuration
void ColorStyleConfigDlg::EnableCustomStyle(BOOL enable)
{
	m_listBox.EnableWindow(enable);

	CWnd *p = GetDlgItem(IDC_COLOR_BUTTON);
	if (p) {
		((CButton*)p)->EnableWindow(enable);
	}

	CWnd* pall = GetDlgItem(IDC_COLOR_CHECK_ALL);
	if (pall) {
		((CButton*)pall)->EnableWindow(enable);
	}

	p = GetDlgItem(IDC_COLOR_CHECK);
	if (p) {
		((CButton*)p)->EnableWindow(enable);
	}
}

void ColorStyleConfigDlg::ClearCustomStyleChecks()
{
	CWnd* p = GetDlgItem(IDC_COLOR_CHECK);
	if (p)
	{
		((CButton*)p)->SetCheck(0);
	}

	CWnd* pall = GetDlgItem(IDC_COLOR_CHECK_ALL);
	if (pall) {
		((CButton*)pall)->SetCheck(0);
	}
}

void ColorStyleConfigDlg::UpdateRegistry()
{
	BOOL retval;
	CString section_pane_color = CString(sz_Software_mboxview) + L"\\PaneColors";

	retval = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_pane_color, L"colorStyle", m_selectedColorStyle);

	if (m_selectedColorStyle == ColorStylesDB::ColorCustom)
	{
		CMainFrame::m_ColorStylesDB.m_colorStyles.ColorArray2String(m_customStyleColorList);
		retval = CProfile::_WriteProfileString(HKEY_CURRENT_USER, section_pane_color, L"customStyleColorList", m_customStyleColorList);
	}
}

void ColorStyleConfigDlg::LoadFromRegistry()
{
	BOOL retval;
	DWORD bColorStyle;
	CString customStyleColorList;
	CString customColorList;

	ColorStyleConfig &colorStyles = CMainFrame::m_ColorStylesDB.m_colorStyles;
	ColorStyleConfig &customColorStyles = CMainFrame::m_ColorStylesDB.m_customColorStyles;

	CString section_pane_color = CString(sz_Software_mboxview) + L"\\PaneColors";

	m_selectedColorStyle = ColorStylesDB::ColorDefault;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_pane_color, L"colorStyle", bColorStyle))
	{
		m_selectedColorStyle = bColorStyle;
		m_lastColorStyle = bColorStyle;
	}

	if (m_selectedColorStyle >= ColorStylesDB::MaxColorStyles)
	{
		m_selectedColorStyle = ColorStylesDB::ColorDefault;
		m_lastColorStyle = ColorStylesDB::ColorDefault;
	}

	if (retval = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_pane_color, L"customStyleColorList", customStyleColorList))
	{
		m_customStyleColorList = customStyleColorList;
		customColorStyles.String2ColorArray(m_customStyleColorList);
		colorStyles.String2ColorArray(m_customStyleColorList);
	}

	for (int i = 0; i < 16; i++)
	{
		m_customColors[i] = RGB(210, 210, 210);
	}
	int deb = 1;
}

void ColorStyleConfigDlg::OnBnClickedButtonHelp()
{
	// TODO: Add your control notification handler code here

	CString helpFileName = L"ColorStyleConfigHelp.pdf";
	HWND h = GetSafeHwnd();
	CMainFrame::OpenHelpFile(helpFileName, h);

	return;
}


void ColorStyleConfigDlg::OnBnClickedButtonClose()
{
	// TODO: Add your control notification handler code here

	ShowWindow(SW_HIDE);
}


void ColorStyleConfigDlg::OnBnClickedApplyToAllPanes()
{
	// TODO: Add your control notification handler code here

	ColorStyleConfig& colorStyles = CMainFrame::m_ColorStylesDB.m_colorStyles;
	ColorStyleConfig& customColorStyles = CMainFrame::m_ColorStylesDB.m_customColorStyles;

	if ((m_selectedPane == ColorStyleConfig::MailConversion1) || (m_selectedPane == ColorStyleConfig::MailConversion1))
	{
		CString txt = L"\"Apply to All Panes\" can't be used when MailConversion1 or MailConversion2 is selected\n";
		int answer = MessageBox(txt, L"Info", MB_ICONHAND|MB_OK);
		return;
	}

	CWnd* p = GetDlgItem(IDC_COLOR_CHECK_ALL);

	int chk = 0;
	if (p)
	{
		chk = ((CButton*)p)->GetCheck();
		if (chk)
			m_applyColorToAllPanes = 1;
		else
			m_applyColorToAllPanes = 0;
	}
	int nSel = m_listBox.GetCurSel();
	if ((nSel < 0) || (nSel > ColorStyleConfig::MailMaxPanes))
	{
		CSCAssert();
		//return;
	}
	if (nSel != m_selectedPane)
	{
		CSCAssert();
		//return;
	}

	int i;
	if (m_applyColorToAllPanes)
	{
		// Don't set MailConversion1 and MailConversion2
		for (i = 0; i < ColorStyleConfig::MailMaxPanes - 2; i++)
		{
			colorStyles.SetColor(i, m_buttonColor);
		}
	}
	else
	{
		// Don't set MailConversion1 and MailConversion2
		for (i = 0; i < ColorStyleConfig::MailMaxPanes - 2; i++)
		{
			colorStyles.SetColor(i, m_buttonColor);
			m_buttonColor = customColorStyles.GetColor(i);
			colorStyles.SetColor(i, m_buttonColor);
		}
	}

	p = GetDlgItem(IDC_COLOR_BUTTON);
	if (p)
	{
		p->RedrawWindow();
	}

	if (m_MainFrameWnd)
		m_MainFrameWnd->PostMessage(WM_CMD_PARAM_NEW_COLOR_MESSAGE, ColorStylesDB::ColorCustom, 0);

	int deb = 1;
}

BOOL ColorStyleConfigDlg::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}
