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
		MaxColorStyles = 10
	};
	ColorStylesDB();
	//
	static int ID2ENUM(WORD nID);
	static WORD ColorStylesDB::NUM2ID(int colorStyle);

	ColorStyleConfig m_loadedColorStyles;
	ColorStyleConfig m_colorStyles;
	ColorStyleConfig m_customColorStyles;
	ColorStyleConfig m_savedCustomColorStyles;
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
	//CCheckListBox m_listBox;
	CListBox m_listBox;

	COLORREF m_buttonColor;
	CButton m_ColorButton1;

	int m_lastColorStyle;
	int m_selectedColorStyle;
	int m_selectedPane;
	CString m_customStyleColorList;
	CString m_customColorList;
	int m_applyColor;
	COLORREF m_customColors[16];

	void EnableCustomStyle(BOOL enable);
	void UpdateRegistry();
	void LoadFromRegistry();

	afx_msg void OnNMCustomdrawColorButton(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedColorButton();
	afx_msg void OnBnClickedColorCheck();
	afx_msg void OnLbnSelchangeList1();
	afx_msg void OnBnClickedOk();
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
};
