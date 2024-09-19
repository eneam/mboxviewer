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


#pragma once


// ColorStyleConfigDlg dialog


#define COLOR_NOT_SET 0xFFFFFF
#define COLOR_WHITE RGB(255, 255, 255)
#define PeachPuff1 	RGB(255,218,185)
#define AntiqueWhite3 RGB(205,192,176)

class ColorStyleConfig
{
public:
	enum
	{
		MailArchiveList = 0,
		MailSummaryList = 1,
		MailMessage = 2,
		MailMessageAttachments = 3,
		MailSummaryTitles = 4,
		MailMessageHeader = 5,
		MailConversion1 = 6,
		MailConversion2 = 7,
		MailMaxPanes = 8
	};

	ColorStyleConfig();
	void Copy(ColorStyleConfig &colorConfig);

	void SetDefaults();
	void SetColorStyle(int colorStyle);

	void CopyArray(DWORD *array, int arrayLength);
	DWORD GetColor(int index);
	DWORD SetColor(int index, DWORD color);
	void ColorArray2String(CString &str);
	void String2ColorArray(CString &str);

	DWORD m_ColorTbl[MailMaxPanes];
	BOOL m_ReplaceAllWhiteBackgrounTags;
};


// Defined as Global/static in MainFrame.h. Not defined anywhere else 
class ColorStylesDB
{
public:
	enum
	{
		ColorDefault = 0,
		ColorCustom = 1,
		ColorStyle1 = 2,
		ColorStyle2 = 3,
		ColorStyle3 = 4,
		ColorStyle4 = 5,
		ColorStyle5 = 6,
		ColorStyle6 = 7,
		ColorStyle7 = 8,
		ColorStyle8 = 9,
		ColorStyle9 = 10,
		ColorStyle10 = 11,
		ColorStyle11 = 12,
		ColorStyle12 = 13,
		ColorStyle13 = 14,
		ColorStyle14 = 15,
		ColorStyle15 = 16,
		ColorStyle16 = 17,
		MaxColorStyles = 18
	};
	ColorStylesDB();
	//
	static int ID2ENUM(WORD nID);
	static WORD ColorStylesDB::NUM2ID(int colorStyle);

	ColorStyleConfig m_colorStyles;        // array of current predefined or custom colors for all Panes
	ColorStyleConfig m_customColorStyles;  // array of custom colors/user selected colors for all Panes
};

class ColorStyleConfigDlg : public CDialogEx
{
	DECLARE_DYNAMIC(ColorStyleConfigDlg)

public:
	ColorStyleConfigDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~ColorStyleConfigDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_COLOR_STYLE_DLG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()

	void LoadData();
public:
	CWnd *m_MainFrameWnd;
	CListBox m_listBox;            // List of Panes

	COLORREF m_buttonColor;        // "Pick Color" button color
	CButton m_ColorButton1;        // "Pick Color" button

	int m_lastColorStyle;            // last saved color style.
	int m_selectedColorStyle;        // radio button to select Color style number from Default to ColorStyle8 range
	int m_selectedPane;              // selected Pane number frpm Pane List in Custom config mode
	CString m_customStyleColorList;  // list of Pane custom colors to be written into registry
	CString m_customColorList;       // Not used // FIXME
	int m_applyColorToSelectedPane;  // status of check box "Apply to Selected"
	int m_applyColorToAllPanes;      // status of check box "Apply to All Panes"
	COLORREF m_customColors[16];     // colors selected by user known to Color Picker Dialog
	//
	CToolTipCtrl m_toolTip;

	void EnableCustomStyle(BOOL enable); // enable/disable custom configuration
	void ClearCustomStyleChecks();
	void UpdateRegistry();
	void LoadFromRegistry();

	afx_msg void OnNMCustomdrawColorButton(NMHDR *pNMHDR, LRESULT *pResult);  // Draws "Pick Color" button
	afx_msg void OnBnClickedColorButton();
	afx_msg void OnBnClickedColorCheck(); // Apply selected color to singl/selected Pane
	afx_msg void OnLbnSelchangeList1();  // Called in Custom mode when new pane is selected
	afx_msg void OnBnClickedOk();  // Invoked by Save button. Verifies and updates registry
	afx_msg void OnBnClickedCancel();
	afx_msg void OnBnClickedColorStyle(UINT nID);
	//
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnBnClickedButtonHelp();
	afx_msg void OnClose();
	//
	virtual void PostNcDestroy();
	virtual BOOL OnInitDialog();
	
	afx_msg void OnBnClickedButtonClose();
	afx_msg void OnBnClickedApplyToAllPanes();  // Apply selected color to all Panes
	afx_msg BOOL OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult);
};
