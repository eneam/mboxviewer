// ColorStyleConfigDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ColorStyleConfigDlg.h"
#include "afxdialogex.h"
#include "ColorPickerDlg.h"
#include "mboxview.h"
#include "MainFrm.h"
#include "NListView.h"


// ColorStyleConfigDlg dialog

IMPLEMENT_DYNAMIC(ColorStyleConfigDlg, CDialogEx)

static void CSCAssert()
{
#ifdef _DEBUG
	ASSERT(0 == 1);
#endif
	int deb = 1;
}

ColorStyleConfigDlg::ColorStyleConfigDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_COLOR_STYLE_DLG, pParent)
{
	m_MainFrameWnd = pParent;
	m_buttonColor = RGB(255, 255, 255);
	m_selectedColorStyle = 0;
	m_lastColorStyle = 0;
	m_selectedPane = 0;
	m_applyColor = 0;
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
	DDX_Check(pDX, IDC_COLOR_CHECK, m_applyColor);
}


BEGIN_MESSAGE_MAP(ColorStyleConfigDlg, CDialogEx)
	ON_WM_CLOSE()
	ON_NOTIFY(NM_CUSTOMDRAW, IDC_COLOR_BUTTON, &ColorStyleConfigDlg::OnNMCustomdrawColorButton)
	ON_BN_CLICKED(IDC_COLOR_BUTTON, &ColorStyleConfigDlg::OnBnClickedColorButton)
	ON_BN_CLICKED(IDC_COLOR_CHECK, &ColorStyleConfigDlg::OnBnClickedColorCheck)
	ON_LBN_SELCHANGE(IDC_LIST1, &ColorStyleConfigDlg::OnLbnSelchangeList1)
	ON_BN_CLICKED(IDOK, &ColorStyleConfigDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &ColorStyleConfigDlg::OnBnClickedCancel)
	ON_COMMAND_RANGE(IDC_COLOR_DFLT, IDC_COLOR_STYLE8, &ColorStyleConfigDlg::OnBnClickedColorStyle)
	ON_BN_CLICKED(IDC_BUTTON_HELP, &ColorStyleConfigDlg::OnBnClickedButtonHelp)
	ON_BN_CLICKED(IDC_BUTTON_CLOSE, &ColorStyleConfigDlg::OnBnClickedButtonClose)
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
	m_listBox.AddString("MailArchiveTree");
	m_listBox.AddString("MailSummary");
	m_listBox.AddString("MailMessage");
	m_listBox.AddString("MailAttachments");
	m_listBox.AddString("MailSummaryColumnTitles");
	m_listBox.AddString("MailMessageHeaderFields");
	m_listBox.AddString("MailConversation1");
	m_listBox.AddString("MailConversation2");

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

	CheckRadioButton(IDC_COLOR_DFLT, IDC_COLOR_STYLE8, nID);

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
			dc.FillRect(&rect, &CBrush(m_buttonColor));
			::FrameRect(hDC, rect, (HBRUSH)GetStockObject(BLACK_BRUSH));

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

	CWnd *p = GetDlgItem(IDC_COLOR_CHECK);
	if (p)
	{
		int chk = ((CButton*)p)->GetCheck();
		if (chk)
			OnBnClickedColorCheck();
	}

	p = GetDlgItem(IDC_COLOR_BUTTON);
	if (p)
	{
		p->RedrawWindow();
	}

	CString colorStr;
	NListView::Color2Str(m_buttonColor, colorStr);

	int deb = 1;
}


void ColorStyleConfigDlg::OnBnClickedColorCheck()
{
	// TODO: Add your control notification handler code here

	ColorStyleConfig &colorStyles = CMainFrame::m_ColorStylesDB.m_colorStyles;
	ColorStyleConfig &customColorStyles = CMainFrame::m_ColorStylesDB.m_customColorStyles;

	CWnd *p = GetDlgItem(IDC_COLOR_CHECK);
	if (m_buttonColor == customColorStyles.m_ColorTbl[m_selectedPane])
	{
		m_applyColor = 0;
		if (p)
		{
			((CButton*)p)->SetCheck(m_applyColor);
		}
		return;
	}

	int chk = 0;
	if (p)
	{
		chk = ((CButton*)p)->GetCheck();
		if (chk)
			m_applyColor = 1;
		else
			m_applyColor = 0;
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

	if (m_applyColor)
	{
		colorStyles.SetColor(m_selectedPane, m_buttonColor);
		int deb = 1;
	}
	else
	{
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
		m_MainFrameWnd->PostMessageA(WM_CMD_PARAM_NEW_COLOR_MESSAGE, ColorStylesDB::ColorCustom, 0);

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

	p = GetDlgItem(IDC_COLOR_CHECK);
	if (p) 
	{
		if (customColorStyles.GetColor(nSel) == colorStyles.GetColor(nSel))
		{
			m_applyColor = 0;
			((CButton*)p)->SetCheck(0);
		}
		else
		{
			m_applyColor = 1;
			((CButton*)p)->SetCheck(1);
		}
	}

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
		m_ColorTbl[MailMessageHeader] = RGB(205, 192, 196);
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
		colorStr.Format("0x%06x", color);
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
		color = strtoul(colorStr, 0, 16);
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
	//UpdateData(FALSE); // control -> data
	//UpdateData(TRUE); // data -> control

	ColorStyleConfig &colorStyles = CMainFrame::m_ColorStylesDB.m_colorStyles;
	ColorStyleConfig &customColorStyles = CMainFrame::m_ColorStylesDB.m_customColorStyles;
	
	if (m_selectedColorStyle == ColorStylesDB::ColorCustom)
	{
		customColorStyles.Copy(colorStyles);
		m_applyColor = 0;		
		CWnd *p = GetDlgItem(IDC_COLOR_CHECK);
		if (p)
		{
			((CButton*)p)->SetCheck(m_applyColor);
		}

	}
	m_lastColorStyle = m_selectedColorStyle;

	UpdateRegistry();
	// Optional ??
	//if (m_MainFrameWnd) m_MainFrameWnd->PostMessageA(WM_CMD_PARAM_NEW_COLOR_MESSAGE, m_selectedColorStyle, 0);
	//ShowWindow(SW_HIDE);
	//CDialogEx::OnOK();
}

void ColorStyleConfigDlg::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here

	ColorStyleConfig &colorStyles = CMainFrame::m_ColorStylesDB.m_colorStyles;
	ColorStyleConfig &customColorStyles = CMainFrame::m_ColorStylesDB.m_customColorStyles;

	m_selectedColorStyle = m_lastColorStyle;
	colorStyles.SetColorStyle(m_selectedColorStyle);

	m_applyColor = 0;
	CWnd *p = GetDlgItem(IDC_COLOR_CHECK);
	if (p)
	{
		((CButton*)p)->SetCheck(m_applyColor);
	}

	m_buttonColor = colorStyles.GetColor(m_selectedPane);

	WORD nID = ColorStylesDB::NUM2ID(m_selectedColorStyle);

	CheckRadioButton(IDC_COLOR_DFLT, IDC_COLOR_STYLE8, nID);

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

	p = GetDlgItem(IDC_COLOR_BUTTON);
	if (p) {
		p->RedrawWindow();
	}

	EnableCustomStyle(enableColorSelection);

	if (m_MainFrameWnd)
		m_MainFrameWnd->PostMessage(WM_CMD_PARAM_NEW_COLOR_MESSAGE, m_selectedColorStyle, 0);

	//ShowWindow(SW_HIDE);
	//CDialogEx::OnCancel();
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

	return colorStyle;
}

WORD ColorStylesDB::NUM2ID(int colorStyle)
{
	static WORD ColorStyle2ID[] = { 
		IDC_COLOR_DFLT , IDC_COLOR_CUSTOM , IDC_COLOR_STYLE1, IDC_COLOR_STYLE2  ,IDC_COLOR_STYLE3 ,
		IDC_COLOR_STYLE4 , IDC_COLOR_STYLE5 , IDC_COLOR_STYLE6 , IDC_COLOR_STYLE7 , IDC_COLOR_STYLE8
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

	return nID;
}

void ColorStyleConfigDlg::OnBnClickedColorStyle(UINT nID)
{
	// TODO: Add your specialized code here and/or call the base class

	ColorStyleConfig &colorStyles = CMainFrame::m_ColorStylesDB.m_colorStyles;
	ColorStyleConfig &customColorStyles = CMainFrame::m_ColorStylesDB.m_customColorStyles;

	int colorStyle = ColorStylesDB::ID2ENUM(nID);

	if (colorStyle < 0)
		return;

	if ((m_selectedColorStyle == ColorStylesDB::ColorCustom) && (colorStyle != ColorStylesDB::ColorCustom))
	{
		int i;
		for (i = 0; i < ColorStyleConfig::MailMaxPanes; i++)
		{
			if (colorStyles.m_ColorTbl[i] != customColorStyles.m_ColorTbl[i])
			{
				CString txt = _T("Custom Color Style Changed! Discard changes?");
				int answer = MessageBox(txt, _T("Warning"), MB_APPLMODAL | MB_ICONWARNING | MB_YESNO);
				if (answer == IDNO)
				{
					CheckRadioButton(IDC_COLOR_DFLT, IDC_COLOR_STYLE8, IDC_COLOR_CUSTOM);

					m_applyColor = 1;
					CWnd *p = GetDlgItem(IDC_COLOR_CHECK);
					if (p)
					{
						((CButton*)p)->SetCheck(m_applyColor);
					}
					return;
				}
				else
				{
					m_applyColor = 0;
					CWnd *p = GetDlgItem(IDC_COLOR_CHECK);
					if (p)
					{
						((CButton*)p)->SetCheck(m_applyColor);
					}
					;  // all done
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

void ColorStyleConfigDlg::EnableCustomStyle(BOOL enable)
{
	m_listBox.EnableWindow(enable);

	CWnd *p = GetDlgItem(IDC_COLOR_BUTTON);
	if (p) {
		((CButton*)p)->EnableWindow(enable);
	}

	p = GetDlgItem(IDC_COLOR_CHECK);
	if (p) {
		((CButton*)p)->EnableWindow(enable);
	}
}

void ColorStyleConfigDlg::UpdateRegistry()
{
	BOOL retval;

	retval = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("colorStyle"), m_selectedColorStyle);

	if (m_selectedColorStyle == ColorStylesDB::ColorCustom)
	{
		CMainFrame::m_ColorStylesDB.m_colorStyles.ColorArray2String(m_customStyleColorList);
		retval = CProfile::_WriteProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("customStyleColorList"), m_customStyleColorList);
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

	m_selectedColorStyle = ColorStylesDB::ColorDefault;
	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("colorStyle"), bColorStyle))
	{
		m_selectedColorStyle = bColorStyle;
		m_lastColorStyle = bColorStyle;
	}

	if (m_selectedColorStyle >= ColorStylesDB::MaxColorStyles)
	{
		m_selectedColorStyle = ColorStylesDB::ColorDefault;
		m_lastColorStyle = ColorStylesDB::ColorDefault;
	}

	if (retval = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("customStyleColorList"), customStyleColorList))
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

	CString HelpPath = FileUtils::GetmboxviewTempPath("MboxHelp");

	BOOL createDirOk = TRUE;
	if (!FileUtils::PathDirExists(HelpPath))
		createDirOk = FileUtils::CreateDirectory(HelpPath);

	CString codePageIdsFile = "ColorStyleConfigHelp.htm";
	CString fullPath = HelpPath + "\\" + codePageIdsFile;

	CFile fp;
	if (!fp.Open(fullPath, CFile::modeWrite | CFile::modeCreate)) {
		CString txt = _T("Could not create \"") + fullPath;
		txt += _T("\" file.\nMake sure file is not open on other applications.");
		HWND h = NULL; // we don't have any window yet
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return;
	}

	CString htmlHdr;

	htmlHdr += "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=US-ASCII\"></head><body style=\'margin-left:0.5cm\'>";
	htmlHdr += "<br><font size=\"+2\"><b>Mbox Viewer Color Style Configuration Help</b></font><br><br>";

	fp.Write((LPCSTR)htmlHdr, htmlHdr.GetLength());

	CString text;

	text.Empty();
	text.Append(
		"Please review the User Guide provided with the package for additional details on the Color Style Confguration feature.<br>"
		"<br>"
		"Color Style Configuration dialog allows users to replace the white background of the 8 panes listed on the dialog box with the predefined colors or custom colors defined by users.<br>"
		"The dialog offers eight predefined color styles users can evaluate before attempting to configure the custom color style."
		" It is helpfull to open one of the mails with attachments to help to evaluate predefined and custom color styles.<br>"
		"<br>"
		"<br><font size=\"+1\"><b>How to create custom color style</b></font><br><br>"
		"Select \"Custom\" radio buttom to start configuration of the custom color style. <br><br>"
		"Configure the same or different color for each pane (MailArchiveTree, etc)."
		"<br><br>"
		"Select one of the panes and left click on the \"Pick Color\" button to start the \"Color Selection Dialog\".<br><br>"
		"The \"Pick Color\" button will initially be populated with the color associated with the selected pane.<br><br>"
		"The top raw of the Custom Colors area will initially be populated with the colors associated with each pane.<br><br>"
		
		"You can change the color in four different areas. The top left section displays a list of 48 predefined colors.<br><br>"
		"If the desired color is not in that section, you can click on the crosshair and drag the mouse in the multi-colored palette.<br><br>"
		"You can also drag the right bar that displays a range based on the color of the palette.<br><br>"
		"For more precision, you can type the Red, Green and Blue values in the custom color editor area. Each uses a integral value that ranges from 1 to 255.<br><br>"
		"Once you are happy with the custom color, left click on the OK button in the \"Color Selection Dialog\". The \"Color Selection Dialog\" will be closed and control will return to the \"Color Style Confguration\" dialog.<br><br>"
		"The \"Pick Color\" button will be painted with the created custom color. Left click on the \"Apply\" button to apply the new color to the selected pane. Left click on the \"Apply\" button again to cancel the change.<br><br>"
		"Left click on the Save button to save the change in order for the change to persist after the mbox viewer shutdown. Or, customize colors of all panes first and then left click on the Save button to save all changes.<br><br>"
		"Left click on the Cancel button to discard all changes made since the last Save or startup.<br><br>"
		"Left click on the Close button or standard Windows Close button to hide the the \"Color Selection Dialog\". Changes made so far are not discarded. To reopen, select File -> Color Style Config menu option.<br><br>"
		"<br>NOTE that the saved color style will be applied to all mails. "
		"You can temporarily disable custom color style for the active mail by right click on the mail header in the mail message window and selecting Enable/Disable Custom Color Style menu option.<br>"

	);
	fp.Write((LPCSTR)text, text.GetLength());


	text.Empty();
	text.Append(
		"<br><br>"
		"<br><br>"
	);
	fp.Write((LPCSTR)text, text.GetLength());

	CString htmlEnd = "\r\n</body></html>";
	fp.Write((LPCSTR)htmlEnd, htmlEnd.GetLength());

	fp.Close();

	ShellExecute(NULL, _T("open"), fullPath, NULL, NULL, SW_SHOWNORMAL);

	int deb = 1;
}


void ColorStyleConfigDlg::OnBnClickedButtonClose()
{
	// TODO: Add your control notification handler code here

	ShowWindow(SW_HIDE);
}
