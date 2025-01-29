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


// SMTPMailServerConfigDlg.cpp : implementation file
//

#include "StdAfx.h"
#include "mboxview.h"
#include "SMTPMailServerConfigDlg.h"
#include "afxdialogex.h"
#include "TextUtilsEx.h"
#include "ResHelper.h"


#define GMAIL_MAX_MAIL_SIZE 25590
#define YAHOO_MAX_MAIL_SIZE 40720
#define OUTLOOK_MAX_MAIL_SIZE 35830


// SMTPMailServerConfigDlg dialog

IMPLEMENT_DYNAMIC(SMTPMailServerConfigDlg, CDialogEx)

SMTPMailServerConfigDlg::SMTPMailServerConfigDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_SMTP_DLG, pParent)
{
	;
}


SMTPMailServerConfigDlg::~SMTPMailServerConfigDlg()
{
}

void SMTPMailServerConfigDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Radio(pDX, IDC_GMAIL, m_mailDB.ActiveMailServiceType);
	DDX_Text(pDX, IDC_SMTP_SERVER_ADDR, m_mailDB.SMTPConfig.SmtpServerAddress);
	DDX_Text(pDX, IDC_SMTP_SERVER_PORT, m_mailDB.SMTPConfig.SmtpServerPort);
	DDX_Text(pDX, IDC_USER_ACCOUNT, m_mailDB.SMTPConfig.UserAccount);
	DDX_Text(pDX, IDC_USER_PASSWORD, m_mailDB.SMTPConfig.UserPassword);
	DDX_Text(pDX, IDC_MAX_MAIL_SIZE, m_mailDB.SMTPConfig.MaxMailSize);
	DDX_Control(pDX, IDC_TLS_OPTION, TLSOptions);
}


BEGIN_MESSAGE_MAP(SMTPMailServerConfigDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &SMTPMailServerConfigDlg::OnBnClickedOk)
	ON_COMMAND_RANGE(IDC_GMAIL, IDC_CUSTOM, &SMTPMailServerConfigDlg::OnBnClickedMailServiceType)
	ON_BN_CLICKED(IDC_BBUTTON_SAVE, &SMTPMailServerConfigDlg::OnBnClickedBbuttonSave)
	ON_BN_CLICKED(IDC_SMTP_SERVER_HELP, &SMTPMailServerConfigDlg::OnBnClickedSmtpServerHelp)
	ON_BN_CLICKED(IDCANCEL, &SMTPMailServerConfigDlg::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_RESET_MAX_MAIL_SIZE, &SMTPMailServerConfigDlg::OnBnClickedResetMaxMailSize)
	ON_NOTIFY_EX(TTN_NEEDTEXT, 0, &SMTPMailServerConfigDlg::OnTtnNeedText)
END_MESSAGE_MAP()


// SMTPMailServerConfigDlg message handlers

void SMTPMailServerConfigDlg::OnBnClickedMailServiceType(UINT nID)
{
	UpdateData(TRUE);

	int ActiveMailServiceType = nID;
	m_mailDB.SwitchToNewService(ActiveMailServiceType);

	TLSOptions.SetCurSel(m_mailDB.SMTPConfig.EncryptionType);

	UpdateData(FALSE);
};


// Just Close the dialog
void SMTPMailServerConfigDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here

	//UpdateData(TRUE);

	CDialogEx::OnOK();
}

void SMTPMailServerConfigDlg::OnBnClickedCancel()
{
	// TODO: Add your control notification handler code here

	TLSOptions.SetCurSel(m_mailDB.SMTPConfig.EncryptionType);
	UpdateData(FALSE);
	//CDialogEx::OnCancel();
}

BOOL SMTPMailServerConfigDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	TLSOptions.AddString(L"None");
	TLSOptions.AddString(L"Auto");
	TLSOptions.AddString(L"SslOnConnect");
	TLSOptions.AddString(L"StartTls");
	TLSOptions.AddString(L"StartTlsWhenAvailable");

	// m_mailDB is already set from  void CMainFrame::OnFileSmtpmailserverconfig()
	TLSOptions.SetCurSel(m_mailDB.SMTPConfig.EncryptionType);

	ResHelper::UpdateDialogItemsInfo(this);
	BOOL retA = ResHelper::ActivateToolTips(this, m_toolTip);

	UpdateData(FALSE);

	ResHelper::LoadDialogItemsInfo(this);

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}

void SMTPMailServerConfigDlg::OnBnClickedBbuttonSave()
{
	// TODO: Add your control notification handler code here

	UpdateData();

	int idx = TLSOptions.GetCurSel();

	m_mailDB.SaveChangesToActiveService(idx);

	if (idx < 0)
	{
		// undo changes
		TLSOptions.SetCurSel(m_mailDB.SMTPConfig.EncryptionType);
		UpdateData(FALSE);
	}
}

void SMTPMailServerConfigDlg::OnBnClickedSmtpServerHelp()
{
	// TODO: Add your control notification handler code here
	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	CString processPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, section_general, L"processPath");

	CString processDir;
	FileUtils::CPathGetPath(processPath, processDir);
	CString filePath = processDir + L"\\HelpFiles\\ForwardMails.pdf";

	if (FileUtils::PathFileExist(filePath))
	{
		ShellExecute(NULL, L"open", filePath, NULL, NULL, SW_SHOWNORMAL);
	}
	else
	{
		HWND h = GetSafeHwnd();
		CString txt;
		CString fmt = L"Help file \"%s\" doesn't exist";
		ResHelper::TranslateString(fmt);
		txt.Format(fmt, filePath);
		int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONQUESTION | MB_OK);
	}

	int deb = 1;
}

void MailConfig::Copy(MailConfig &config)
{
	MailServiceName = config.MailServiceName;
	SmtpServerAddress = config.SmtpServerAddress;
	SmtpServerPort = config.SmtpServerPort;
	UserAccount = config.UserAccount;
	UserPassword = config.UserPassword;
	EncryptionType = config.EncryptionType;
	MaxMailSize = config.MaxMailSize;
	UserMailAddress = config.UserMailAddress;
}

void MailConfig::Write2Registry()
{
	BOOL ret;
	CString m_section_mail_forward = CString(sz_Software_mboxview) + "\\MailService\\" + MailServiceName;

	ret = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section_mail_forward, L"MailServiceName", MailServiceName);
	ret = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section_mail_forward, L"SmtpServerAddress", SmtpServerAddress);
	ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section_mail_forward, L"SmtpServerPort", SmtpServerPort);
	ret = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section_mail_forward, L"UserAccount", UserAccount);
	//ret = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section_mail_forward, L"UserPassword", UserPassword);
	ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section_mail_forward, L"EncryptionType", EncryptionType);
	ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section_mail_forward, L"MaxMailSize", MaxMailSize);
	ret = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section_mail_forward, L"UserMailAddress", UserMailAddress);
}

void MailConfig::ReadFromRegistry(CString &serviceName)
{
	CString m_section_mail_forward = CString(sz_Software_mboxview) + "\\MailService\\" + serviceName;

	MailServiceName = CProfile::_GetProfileString(HKEY_CURRENT_USER, m_section_mail_forward, L"MailServiceName");
	SmtpServerAddress = CProfile::_GetProfileString(HKEY_CURRENT_USER, m_section_mail_forward, L"SmtpServerAddress");
	SmtpServerPort = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section_mail_forward, L"SmtpServerPort");
	UserAccount = CProfile::_GetProfileString(HKEY_CURRENT_USER, m_section_mail_forward, L"UserAccount");
	//UserPassword = CProfile::_GetProfileString(HKEY_CURRENT_USER, m_section_mail_forward, L"UserPassword");
	EncryptionType = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section_mail_forward, L"EncryptionType");
	MaxMailSize = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section_mail_forward, L"MaxMailSize");
	UserMailAddress = CProfile::_GetProfileString(HKEY_CURRENT_USER, m_section_mail_forward, L"UserMailAddress");
}

// Not used at the moment. Consider to use if it simplies the code
DWORD MailConfig::WStr2Ansi(MailConfigA& mailConfig)
{
	DWORD error;
	BOOL retW2A;

	retW2A = TextUtilsEx::WStr2Ansi(MailServiceName, mailConfig.MailServiceName, error);
	retW2A = TextUtilsEx::WStr2Ansi(SmtpServerAddress, mailConfig.SmtpServerAddress, error);
	mailConfig.SmtpServerPort = SmtpServerPort;
	retW2A = TextUtilsEx::WStr2Ansi(UserAccount, mailConfig.UserAccount, error);
	retW2A = TextUtilsEx::WStr2Ansi(UserPassword, mailConfig.UserPassword, error);
	mailConfig.EncryptionType = EncryptionType;
	mailConfig.MaxMailSize = MaxMailSize;
	retW2A = TextUtilsEx::WStr2Ansi(UserMailAddress, mailConfig.UserMailAddress, error);

	return 0;
}

MailDB::MailDB()
{
	;
}

MailDB::~MailDB()
{
	;
}

void MailDB::Initialize()
{
	InitConfigsToDflts();
	LoadData();
}

void MailDB::InitConfigsToDflts()
{
	ActiveMailServiceType = 0;
	ActiveMailService = L"Gmail";

	GmailSMTPConfig.MailServiceName = L"Gmail";
	GmailSMTPConfig.SmtpServerAddress = L"smtp.gmail.com";
	GmailSMTPConfig.SmtpServerPort = 587;
	GmailSMTPConfig.UserAccount = L"";
	GmailSMTPConfig.UserPassword = L"";
	GmailSMTPConfig.EncryptionType = StartTls;
	GmailSMTPConfig.MaxMailSize = GMAIL_MAX_MAIL_SIZE;
	GmailSMTPConfig.UserMailAddress = L"";

	YahooSMTPConfig.MailServiceName = L"Yahoo";
	YahooSMTPConfig.SmtpServerAddress = L"smtp.mail.yahoo.com";
	YahooSMTPConfig.SmtpServerPort = 587;
	YahooSMTPConfig.UserAccount = L"";
	YahooSMTPConfig.UserPassword = L"";
	YahooSMTPConfig.EncryptionType = StartTls;
	//YahooSMTPConfig.MaxMailSize = 25590;
	YahooSMTPConfig.MaxMailSize = YAHOO_MAX_MAIL_SIZE; // ~39MB reported by Yahoo SMPT Server Nov/20/2022 and it seems to work
	YahooSMTPConfig.UserMailAddress = L"";

	OutlookSMTPConfig.MailServiceName = L"Outlook";
	OutlookSMTPConfig.SmtpServerAddress = L"smtp-mail.outlook.com";
	OutlookSMTPConfig.SmtpServerPort = 587;
	OutlookSMTPConfig.UserAccount = L"";
	OutlookSMTPConfig.UserPassword = L"";
	OutlookSMTPConfig.EncryptionType = StartTls;
	OutlookSMTPConfig.MaxMailSize = OUTLOOK_MAX_MAIL_SIZE;
	OutlookSMTPConfig.UserMailAddress = L"";

	CustomSMTPConfig.MailServiceName = L"Custom";
	CustomSMTPConfig.SmtpServerAddress = L"";
	CustomSMTPConfig.SmtpServerPort = 0;
	CustomSMTPConfig.UserAccount = L"";
	CustomSMTPConfig.UserPassword = L"";
	CustomSMTPConfig.EncryptionType = None;
	CustomSMTPConfig.MaxMailSize = 0;
	CustomSMTPConfig.UserMailAddress = L"";

}

void MailDB::LoadData()
{
	CString m_section_mail_forward = CString(sz_Software_mboxview) + L"\\MailService";

	CString ActiveMailService;
	BOOL ret = CProfile::_GetProfileString(HKEY_CURRENT_USER, m_section_mail_forward, L"ActiveMailService", ActiveMailService);
	if (ret == FALSE)  // assume this is first time
	{
		CString ActiveMailService = L"Gmail";
		Write2Registry(ActiveMailService);
	}
	ReadFromRegistry();
}

void MailDB::Copy(MailDB &db)
{
	if (this == &db)
		return;

	ActiveMailService = db.ActiveMailService;
	ActiveMailServiceType = db.ActiveMailServiceType;
	GmailSMTPConfig = db.GmailSMTPConfig;
	YahooSMTPConfig = db.YahooSMTPConfig;
	OutlookSMTPConfig = db.OutlookSMTPConfig;
	CustomSMTPConfig = db.CustomSMTPConfig;
	SMTPConfig = db.SMTPConfig;
}

// Only called once to populate registry
void MailDB::Write2Registry(CString &ActiveMailService)
{
	CString m_section_mail_forward = CString(sz_Software_mboxview) + L"\\MailService";

	//CString ActiveMailService = "Gmail";
	BOOL ret = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section_mail_forward, L"ActiveMailService", ActiveMailService);

	GmailSMTPConfig.Write2Registry();;
	YahooSMTPConfig.Write2Registry();
	OutlookSMTPConfig.Write2Registry();
	CustomSMTPConfig.Write2Registry();
}

void MailDB::ReadFromRegistry()
{
	CString m_section_mail_forward = CString(sz_Software_mboxview) + "\\MailService";

	ActiveMailService = L"Gmail";
	BOOL ret = CProfile::_GetProfileString(HKEY_CURRENT_USER, m_section_mail_forward, L"ActiveMailService", ActiveMailService);
	if (ActiveMailService.CompareNoCase(L"Gmail") == 0)
		ActiveMailServiceType = 0; // gmail
	else if (ActiveMailService.CompareNoCase(L"Yahoo") == 0)
		ActiveMailServiceType = 1; // yahoo
	else if (ActiveMailService.CompareNoCase(L"Outlook") == 0)
		ActiveMailServiceType = 2; // outlook
	else if (ActiveMailService.CompareNoCase(L"Custom") == 0)
		ActiveMailServiceType = 3; // custom
	else
		; // ???

	GmailSMTPConfig.ReadFromRegistry(CString(L"Gmail"));
	YahooSMTPConfig.ReadFromRegistry(CString(L"Yahoo"));
	OutlookSMTPConfig.ReadFromRegistry(CString(L"Outlook"));
	CustomSMTPConfig.ReadFromRegistry(CString(L"Custom"));

	if (ActiveMailServiceType == 0)
		SMTPConfig.Copy(GmailSMTPConfig);
	else if (ActiveMailServiceType == 1)
		SMTPConfig.Copy(YahooSMTPConfig);
	else if (ActiveMailServiceType == 2)
		SMTPConfig.Copy(OutlookSMTPConfig);
	else if (ActiveMailServiceType == 3)
		SMTPConfig.Copy(CustomSMTPConfig);
	else
		; // ??
}

void MailDB::SaveChangesToActiveService(int encryptionType)
{
	MailConfig *mail = 0;
	if (ActiveMailServiceType == 0)
		mail = &GmailSMTPConfig;
	else if (ActiveMailServiceType == 1)
		mail = &YahooSMTPConfig;
	else if (ActiveMailServiceType == 2)
		mail = &OutlookSMTPConfig;
	else if (ActiveMailServiceType == 3)
		mail = &CustomSMTPConfig;
	else
		; // ??

	if (encryptionType >= 0)
		SMTPConfig.EncryptionType = encryptionType;

	if (mail)
	{
		mail->Copy(SMTPConfig);
	}

	SMTPConfig.Write2Registry();

	CString m_section_mail_forward = CString(sz_Software_mboxview) + L"\\MailService";

	ActiveMailService = SMTPConfig.MailServiceName;
	BOOL ret = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section_mail_forward, L"ActiveMailService", ActiveMailService);
}

void MailDB::SwitchToNewService(UINT nID)
{
	int ActiveMailServiceType = nID;
	if (ActiveMailServiceType == IDC_GMAIL)
		SMTPConfig.Copy(GmailSMTPConfig);
	else if (ActiveMailServiceType == IDC_YAHOO)
		SMTPConfig.Copy(YahooSMTPConfig);
	else if (ActiveMailServiceType == IDC_OUTLOOK)
		SMTPConfig.Copy(OutlookSMTPConfig);
	else if (ActiveMailServiceType == IDC_CUSTOM)
		SMTPConfig.Copy(CustomSMTPConfig);
	else
		; // ??

	ActiveMailService = SMTPConfig.MailServiceName;
};



void SMTPMailServerConfigDlg::OnBnClickedResetMaxMailSize()
{
	// TODO: Add your control notification handler code here
	int maxMailSize = 0;
	int ActiveMailServiceType = 0;

	if (m_mailDB.ActiveMailService.CompareNoCase(L"Gmail") == 0)
	{
		maxMailSize = GMAIL_MAX_MAIL_SIZE;
		m_mailDB.GmailSMTPConfig.MaxMailSize = maxMailSize;
		ActiveMailServiceType = 0; // gmail
	}
	else if (m_mailDB.ActiveMailService.CompareNoCase(L"Yahoo") == 0)
	{
		maxMailSize = YAHOO_MAX_MAIL_SIZE;
		m_mailDB.YahooSMTPConfig.MaxMailSize = maxMailSize;
		ActiveMailServiceType = 1; // yahoo
	}
	else if (m_mailDB.ActiveMailService.CompareNoCase(L"Outlook") == 0)
	{
		maxMailSize = OUTLOOK_MAX_MAIL_SIZE;
		m_mailDB.OutlookSMTPConfig.MaxMailSize = maxMailSize;
		ActiveMailServiceType = 2; // outlook
	}
	else if (m_mailDB.ActiveMailService.CompareNoCase(L"Custom") == 0)
	{
		ActiveMailServiceType = 3; // custom
	}
	else
		; // ???

	if (maxMailSize > 0)
	{
		m_mailDB.SMTPConfig.MaxMailSize = maxMailSize;
		UpdateData(FALSE);
	}
}

BOOL SMTPMailServerConfigDlg::OnTtnNeedText(UINT id, NMHDR* pNMHDR, LRESULT* pResult)
{
	UNREFERENCED_PARAMETER(id);
	static CString toolTipText;

	CWnd* parentWnd = (CWnd*)this;
	BOOL bRet = ResHelper::OnTtnNeedText(parentWnd, pNMHDR, toolTipText);
	*pResult = 0;

	return bRet;
}
