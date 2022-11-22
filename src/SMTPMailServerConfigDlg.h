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


// SMTPMailServerConfigDlg dialog

enum {
	None = 0,
	Auto = 1,
	SslOnConnect = 2,
	StartTls = 3,
	StartTlsWhenAvailable = 4
};


class MailConfig
{
public:
	CString MailServiceName;
	CString SmtpServerAddress;
	UINT SmtpServerPort;
	CString UserAccount;
	CString UserPassword;
	UINT EncryptionType;
	INT MaxMailSize;
	//CComboBox TLSOptions;
	CString UserMailAddress;

	void Copy(MailConfig &config);
	void Write2Registry();
	void ReadFromRegistry(CString &serviceName);
};


class MailDB
{
public:
	MailDB();
	~MailDB();
	//
	CString ActiveMailService;
	int ActiveMailServiceType;
	MailConfig GmailSMTPConfig;
	MailConfig YahooSMTPConfig;
	MailConfig OutlookSMTPConfig;
	MailConfig CustomSMTPConfig;
	MailConfig SMTPConfig;

	void Initialize();
	void Copy(MailDB &db);
	void LoadData();
	void Write2Registry(CString &ActiveMailService);
	void ReadFromRegistry();
	void SaveChangesToActiveService(int encryptionType);
	void SwitchToNewService(UINT nID);
	void InitConfigsToDflts();
};

class SMTPMailServerConfigDlg : public CDialogEx
{
	DECLARE_DYNAMIC(SMTPMailServerConfigDlg)

public:
	SMTPMailServerConfigDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~SMTPMailServerConfigDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_SMTP_DLG };
#endif

	MailDB m_mailDB;
	CComboBox TLSOptions;

protected:
	virtual BOOL OnInitDialog();
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedMailServiceType(UINT nID);
	DECLARE_MESSAGE_MAP()

public:
	afx_msg void OnBnClickedBbuttonSave();
	afx_msg void OnBnClickedSmtpServerHelp();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnBnClickedResetMaxMailSize();
};
