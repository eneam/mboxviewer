#pragma once


// HtmlPdfHdrConfigDlg dialog

class HdrFldFont
{
public:
	HdrFldFont(CString &id);

	CString m_id;
	CString m_genericFontName;  // font family such as 
	CString m_fontName;  // arial, courier, etc
	CString m_fontStyleName;  // normal, bold, etc
	int m_nFontStyle;  // normal, bold, etc
	BOOL m_bIsBold;
	BOOL m_bIsItalic;
	int m_nFontSize;  // hight

	void SaveToRegistry();
	void LoadFromRegistry();
	void SetDflts();
};

class HdrFldList
{
public:
	HdrFldList();
	~HdrFldList();

	DWORD m_nFldBitmap;

	void ClearFldBitmap();
	void LoadFldBitmap();
	void SaveToRegistry();
	void LoadFromRegistry();

	BOOL IsFldSet(int fldNumber);
	void ClearFld(int fldNumber);
	void SetFld(int fldNumber);

	enum {
		HDR_FLD_SUBJECT = 1,
		HDR_FLD_FROM = 2,
		HDR_FLD_TO = 3,
		HDR_FLD_CC = 4,
		HDR_FLD_BCC = 5,
		HDR_FLD_DATE = 6,
		HDR_FLD_ATTACHMENTS = 7,
		HDR_FLD_MAX = 7
	};
};

class HdrFldConfig
{
public:
	HdrFldConfig();

	BOOL m_bHdrFontDflt;
	BOOL m_bHdrBoldFldName;
	int  m_nHdrFontSize;
	BOOL m_bHdrFldCustomNameFont;
	HdrFldList m_HdrFldList;
	HdrFldFont m_HdrFldFontName;
	HdrFldFont m_HdrFldFontText;

	void SaveToRegistry();
	void LoadFromRegistry();
};


class HtmlPdfHdrConfigDlg : public CDialogEx
{
	DECLARE_DYNAMIC(HtmlPdfHdrConfigDlg)

public:
	HtmlPdfHdrConfigDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~HtmlPdfHdrConfigDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_HTML_PDF_HDR_DLG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()

	void LoadData();

public:
	CCheckListBox m_fldListBox;
	HdrFldConfig m_HdrFldConfig;
	int m_bPickFamilyFont;
	int m_bPickConcreteFont;
	int m_bHdrFldName;

	virtual BOOL OnInitDialog();
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedPickConcreteFont();
	afx_msg void OnBnClickedFldNameBold();
	afx_msg void OnEnChangeFldSize();
	afx_msg void OnBnClickedPickFamilyFont();
	afx_msg void OnBnClickedFontDflt();
	afx_msg void OnBnClickedFontCustom();
	afx_msg void OnBnClickedHdrFldHelp();
};
