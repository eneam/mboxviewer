#pragma once


// finestra di dialogo COptionsDlg

class COptionsDlg : public CDialog
{
	DECLARE_DYNAMIC(COptionsDlg)

public:
	COptionsDlg(CWnd* pParent = NULL);   // costruttore standard
	virtual ~COptionsDlg();

// Dati della finestra di dialogo
	enum { IDD = IDD_OPTIONS };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // Supporto DDX/DDV

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	virtual BOOL OnInitDialog();
	int m_format;
	int m_exportEML;
	int	m_barDelay;
	int m_from_charsetId;
	int m_to_charsetId;
	int m_subj_charsetId;
	int m_show_charsets;
	int m_bImageViewer;
};
