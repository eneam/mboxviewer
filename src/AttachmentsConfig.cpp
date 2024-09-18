//
//////////////////////////////////////////////////////////////////
//
//  Windows Mbox Viewer is a free tool to view, search and print mbox mail archives..
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

// AttachmentsConfig.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "AttachmentsConfig.h"
#include "afxdialogex.h"
#include "ResHelper.h"


// AttachmentsConfig dialog

IMPLEMENT_DYNAMIC(AttachmentsConfig, CDialogEx)

AttachmentsConfig::AttachmentsConfig(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_ATTACHMENTS_CONFIG, pParent)
{
	;
}

AttachmentsConfig::~AttachmentsConfig()
{
}

void AttachmentsConfig::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_ATTACHMENT_MAX_SIZE, m_attachmentConfigParams.m_attachmentWindowMaxSize);
	DDX_Check(pDX, IDC_ALL_ATTACHMENT_TYPES, m_attachmentConfigParams.m_bShowAllAttachments_Window);
	DDX_Check(pDX, IDC_ALL_ATTACHMENT_TYPES_INDICATOR, m_attachmentConfigParams.m_bAnyAttachment_Indicator);
	DDX_Check(pDX, IDC_APPEND_SINGLE_MAIL_VIEW, m_attachmentConfigParams.m_bSingleMailView);
	DDX_Check(pDX, IDC_APPEND_MULTIPLE_MAILS_VIEW, m_attachmentConfigParams.m_bMultipleMailsView);
	DDX_Check(pDX, IDC_APPEND_SINGLE_MAIL_PRINT, m_attachmentConfigParams.m_bSingleMailPrint);
	DDX_Check(pDX, IDC_APPEND_MULTIPLE_MAILS_PRINT, m_attachmentConfigParams.m_bMultipleMailsPrint);
}


BEGIN_MESSAGE_MAP(AttachmentsConfig, CDialogEx)
	ON_BN_CLICKED(IDOK, &AttachmentsConfig::OnBnClickedOk)
END_MESSAGE_MAP()


// AttachmentsConfig message handlers

AttachmentConfigParams::AttachmentConfigParams()
{
	SetDflts();
}

void AttachmentConfigParams::SetDflts()
{
	m_attachmentWindowMaxSize = 25;

	m_bShowAllAttachments_Window = FALSE;
	m_bAnyAttachment_Indicator = FALSE;
	m_bSingleMailView = FALSE;
	m_bMultipleMailsView = FALSE;
	m_bSingleMailPrint = FALSE;
	m_bMultipleMailsPrint = FALSE;

	int deb = 1;
}

void AttachmentConfigParams::Copy(AttachmentConfigParams &src)
{
	if (this == &src)
		return;

	m_attachmentWindowMaxSize = src.m_attachmentWindowMaxSize;
	m_bShowAllAttachments_Window = src.m_bShowAllAttachments_Window;
	m_bAnyAttachment_Indicator = src.m_bAnyAttachment_Indicator;

	m_bSingleMailView = src.m_bSingleMailView;
	m_bMultipleMailsView = src.m_bMultipleMailsView;
	m_bSingleMailPrint = src.m_bSingleMailPrint;
	m_bMultipleMailsPrint = src.m_bMultipleMailsPrint;
}

void AttachmentConfigParams::UpdateRegistry(AttachmentConfigParams &current, AttachmentConfigParams &updated)
{
	if (&current == &updated)
		return;

	CString section_attachment = CString(sz_Software_mboxview) + L"\\Attachments";

	if (updated.m_attachmentWindowMaxSize != current.m_attachmentWindowMaxSize) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_attachment, L"attachmentWindowMaxSize", updated.m_attachmentWindowMaxSize);
	}

	if (updated.m_bShowAllAttachments_Window != current.m_bShowAllAttachments_Window) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_attachment, L"showAllAttachments_Window", updated.m_bShowAllAttachments_Window);
	}
	if (updated.m_bAnyAttachment_Indicator != current.m_bAnyAttachment_Indicator) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_attachment, L"anyAttachment_Indicator", updated.m_bAnyAttachment_Indicator);
	}

	if (updated.m_bSingleMailView != current.m_bSingleMailView) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_attachment, L"appendToSingleMailView", updated.m_bSingleMailView);
	}
	if (updated.m_bMultipleMailsView != current.m_bMultipleMailsView) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_attachment, L"appendToMultipleMailsView", updated.m_bMultipleMailsView);
	}
	if (updated.m_bSingleMailPrint != current.m_bSingleMailPrint) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_attachment, L"appendToSingleMailPrint", updated.m_bSingleMailPrint);
	}
	if (updated.m_bMultipleMailsPrint != current.m_bMultipleMailsPrint) {
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_attachment, L"appendToMultipleMailsPrint", updated.m_bMultipleMailsPrint);
	}
}


void AttachmentConfigParams::LoadFromRegistry()
{
	BOOL retval;
	DWORD attachmentWindowMaxSize;
	DWORD bShowAllAttachments_Window, bAnyAttachment_Indicator;
	DWORD bSingleMailView, bMultipleMailsView, bSingleMailPrint, bMultipleMailsPrint;

	CString section_attachment = CString(sz_Software_mboxview) + L"\\Attachments";

	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_attachment, L"attachmentWindowMaxSize", attachmentWindowMaxSize))
		m_attachmentWindowMaxSize = attachmentWindowMaxSize;

	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_attachment, L"showAllAttachments_Window", bShowAllAttachments_Window))
		m_bShowAllAttachments_Window = bShowAllAttachments_Window;

	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_attachment, L"anyAttachment_Indicator", bAnyAttachment_Indicator))
		m_bAnyAttachment_Indicator = bAnyAttachment_Indicator;

	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_attachment, L"appendToSingleMailView", bSingleMailView))
		m_bSingleMailView = bSingleMailView;

	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_attachment, L"appendToMultipleMailsView", bMultipleMailsView))
		m_bMultipleMailsView = bMultipleMailsView;

	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_attachment, L"appendToSingleMailPrint", bSingleMailPrint))
		m_bSingleMailPrint = bSingleMailPrint;

	if (retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_attachment, L"appendToMultipleMailsPrint", bMultipleMailsPrint))
		m_bMultipleMailsPrint = bMultipleMailsPrint;

}

void AttachmentsConfig::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here
	UpdateData();

	if ((m_attachmentConfigParams.m_attachmentWindowMaxSize < 0) || (m_attachmentConfigParams.m_attachmentWindowMaxSize > 100))
	{
		CString txt;
			txt.Format(L"Invalid max size for Attachment Window. Valid size is 0-100 !");
			AfxMessageBox(txt, MB_OK | MB_ICONHAND);
		return;
	}
	CDialogEx::OnOK();
}


BOOL AttachmentsConfig::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here

	ResHelper::LoadDialogItemsInfo(this);
	ResHelper::UpdateDialogItemsInfo(this);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}
