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

#pragma once
#include <afxcmn.h>

class CMimeBody;
class NMsgView;

class AttachmentInfo
{
public:
	UINT m_charsetId;
	CString m_name;
	CStringW m_nameW;
};

class CAttachments :
	public CListCtrl
{
public:

	CAttachments(NMsgView *pMsgView);

	~CAttachments();

	void Reset();
	void Complete();
	void ReleaseResources();

	BOOL FindAttachmentByNameW(CStringW &name);
	BOOL FindAttachmentByName(CString &name);
	BOOL AddInlineAttachment(CString &name);
	BOOL InsertItemW(CStringW &cStrName, int id, CMimeBody* pBP);


	NMsgView *m_pMsgView;
	CArray<AttachmentInfo*, AttachmentInfo*> m_attachmentTbl;

	DECLARE_MESSAGE_MAP()
	BOOL PreTranslateMessage(MSG* pMsg);

	afx_msg void OnPaint();
	afx_msg void OnActivating(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnDoubleClick(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnRClick(NMHDR* pNMHDR, LRESULT* pResult);
};

