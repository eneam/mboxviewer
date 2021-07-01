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

	TLSOptions.AddString("None");
	TLSOptions.AddString("Auto");
	TLSOptions.AddString("SslOnConnect");
	TLSOptions.AddString("StartTls");
	TLSOptions.AddString("StartTlsWhenAvailable");

	// m_mailDB is already set from  void CMainFrame::OnFileSmtpmailserverconfig()
	TLSOptions.SetCurSel(m_mailDB.SMTPConfig.EncryptionType);

	UpdateData(FALSE);

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
	CString processPath = CProfile::_GetProfileString(HKEY_CURRENT_USER, sz_Software_mboxview, _T("processPath"));

	CString processDir;
	FileUtils::CPathGetPath(processPath, processDir);
	CString filePath = processDir + "\\ForwardMails.pdf";

	ShellExecute(NULL, _T("open"), filePath, NULL, NULL, SW_SHOWNORMAL);
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
}

void MailConfig::Write2Registry()
{
	BOOL ret;
	CString m_section = CString(sz_Software_mboxview) + "\\MailService\\" + MailServiceName;
	ret = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section, "MailServiceName", MailServiceName);
	ret = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section, "SmtpServerAddress", SmtpServerAddress);
	ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section, "SmtpServerPort", SmtpServerPort);
	ret = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section, "UserAccount", UserAccount);
	//ret = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section, "UserPassword", UserPassword);
	ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section, "EncryptionType", EncryptionType);
	ret = CProfile::_WriteProfileInt(HKEY_CURRENT_USER, m_section, "MaxMailSize", MaxMailSize);
}

void MailConfig::ReadFromRegistry(CString &serviceName)
{
	CString m_section = CString(sz_Software_mboxview) + "\\MailService\\" + serviceName;

	MailServiceName = CProfile::_GetProfileString(HKEY_CURRENT_USER, m_section, "MailServiceName");
	SmtpServerAddress = CProfile::_GetProfileString(HKEY_CURRENT_USER, m_section, "SmtpServerAddress");
	SmtpServerPort = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "SmtpServerPort");
	UserAccount = CProfile::_GetProfileString(HKEY_CURRENT_USER, m_section, "UserAccount");
	//UserPassword = CProfile::_GetProfileString(HKEY_CURRENT_USER, m_section, "UserPassword");
	EncryptionType = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "EncryptionType");
	MaxMailSize = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "MaxMailSize");
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
	ActiveMailService = "Gmail";

	GmailSMTPConfig.MailServiceName = "Gmail";
	GmailSMTPConfig.SmtpServerAddress = "smtp.gmail.com";
	GmailSMTPConfig.SmtpServerPort = 587;
	GmailSMTPConfig.UserAccount = "";
	GmailSMTPConfig.UserPassword = "";
	GmailSMTPConfig.EncryptionType = StartTls;
	GmailSMTPConfig.MaxMailSize = 25590;

	YahooSMTPConfig.MailServiceName = "Yahoo";
	YahooSMTPConfig.SmtpServerAddress = "smtp.mail.yahoo.com";
	YahooSMTPConfig.SmtpServerPort = 587;
	YahooSMTPConfig.UserAccount = "";
	YahooSMTPConfig.UserPassword = "";
	YahooSMTPConfig.EncryptionType = StartTls;
	YahooSMTPConfig.MaxMailSize = 25590;

	OutlookSMTPConfig.MailServiceName = "Outlook";
	OutlookSMTPConfig.SmtpServerAddress = "smtp-mail.outlook.com";
	OutlookSMTPConfig.SmtpServerPort = 587;
	OutlookSMTPConfig.UserAccount = "";
	OutlookSMTPConfig.UserPassword = "";
	OutlookSMTPConfig.EncryptionType = StartTls;
	OutlookSMTPConfig.MaxMailSize = 35830;

	CustomSMTPConfig.MailServiceName = "Custom";
	CustomSMTPConfig.SmtpServerAddress = "";
	CustomSMTPConfig.SmtpServerPort = 0;
	CustomSMTPConfig.UserAccount = "";
	CustomSMTPConfig.UserPassword = "";
	CustomSMTPConfig.EncryptionType = None;
	CustomSMTPConfig.MaxMailSize = 0;

}

void MailDB::LoadData()
{
	CString m_section = CString(sz_Software_mboxview) + "\\MailService";

	CString ActiveMailService;
	BOOL ret = CProfile::_GetProfileString(HKEY_CURRENT_USER, m_section, "AvtiveMailService", ActiveMailService);
	if (ret == FALSE)  // assume this is first time
	{
		Write2Registry();
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
void MailDB::Write2Registry()
{
	CString m_section = CString(sz_Software_mboxview) + "\\MailService";

	CString ActiveMailService = "Gmail";
	BOOL ret = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section, "AvtiveMailService", ActiveMailService);

	GmailSMTPConfig.Write2Registry();;
	YahooSMTPConfig.Write2Registry();
	OutlookSMTPConfig.Write2Registry();
	CustomSMTPConfig.Write2Registry();
}

void MailDB::ReadFromRegistry()
{
	CString m_section = CString(sz_Software_mboxview) + "\\MailService";

	ActiveMailService = "Gmail";
	BOOL ret = CProfile::_GetProfileString(HKEY_CURRENT_USER, m_section, "AvtiveMailService", ActiveMailService);
	if (ActiveMailService.CompareNoCase("Gmail") == 0)
		ActiveMailServiceType = 0; // gmail
	else if (ActiveMailService.CompareNoCase("Yahoo") == 0)
		ActiveMailServiceType = 1; // yahoo
	else if (ActiveMailService.CompareNoCase("Outlook") == 0)
		ActiveMailServiceType = 2; // outlook
	else if (ActiveMailService.CompareNoCase("Custom") == 0)
		ActiveMailServiceType = 3; // custom
	else
		; // ???

	GmailSMTPConfig.ReadFromRegistry(CString("Gmail"));
	YahooSMTPConfig.ReadFromRegistry(CString("Yahoo"));
	OutlookSMTPConfig.ReadFromRegistry(CString("Outlook"));
	CustomSMTPConfig.ReadFromRegistry(CString("Custom"));

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

	CString m_section = CString(sz_Software_mboxview) + "\\MailService";

	ActiveMailService = SMTPConfig.MailServiceName;
	BOOL ret = CProfile::_WriteProfileString(HKEY_CURRENT_USER, m_section, "AvtiveMailService", ActiveMailService);
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

