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
// Based on Jeff Lee, see Mime.h/Mime.cpp
//
//////////////////////////////////////////////////////////////////////
//
// MIME message encoding/decoding
//
// Jeff Lee
// Dec 11, 2000
//
//////////////////////////////////////////////////////////////////////

#pragma once

#include "dllist.h"
#include "Mime.h"
#include "MimeCode.h"

class MboxMail;

class MimeParser
{
public:
	static char *GetMultiLine(char *p, char *e, CString &line);
	static int GetFieldValue(CString &fieldLine, int startPos, CString &value);
	static int GetMessageId(CString &fieldLine, int startPos, CString &value);
	static int GetThreadId(CString &fieldLine, int startPos, CString &value);
	static int GetThreadId(CString &fieldLine, int startPos, unsigned __int64 &value);
	static int GetParamValue(CString &fieldLine, int startPos, const char *param, int paramLen, CString &value);
	static int GetFilenameParamValue(CString &fieldLine, int startPos, const char *param, int paramLen, CString &value, BOOL &hasCharset);
	static int GetFilenameParamPartValue(CString &fieldLine, int startPos, const char *param, int paramLen, CString &value, BOOL &hasCharset);
	static BOOL isEmptyLine(const char* p, const char* e);
	static char* SkipEmptyLines(const char* p, const char* e);
	static char *EatNewLine(char* p, const char* e, BOOL &isEmpty);
	inline static char *EatNewLine(char* p, char*e)
	{
		while ((p < e) && (*p++ != '\n'));
		return p;
	}
};

class MailHeader
{
public:
	MailHeader() { Clear(); }
	~MailHeader() { }
public:
	void Clear();
	char *EatNLine(char* p, char* e) { while (p < e && *p++ != '\n'); return p; }

	int Load(const char* pszData, int nDataSize);
	bool IsMultiPart() { return m_IsMultiPart; }
	bool AssertHdr();

	bool IsAttachment();

	bool m_IsText;
	bool m_IsTextPlain;
	bool m_IsTextHtml;
	bool m_IsMessage;
	bool m_IsAttachment;
	bool m_IsMultiPart;
	CString m_Charset;
	UINT m_PageCode;
	CString m_Description;
	CString m_Disposition;
	CString m_TransferEncoding;
	CString m_SubType;
	CString m_MainType;
	CString m_Boundary;
	CString m_ContentType;
	CString m_ContentId;
	CString m_ContentLocation;
	CMimeHeader::MediaType m_MediaType;
	CString m_Name;   // from Content-Type
	UINT m_NamePageCode;
	CString m_AttachmentName;  // filename From Content-Disposition
	UINT m_AttachmentNamePageCode;
	CString m_AttachmentName2;  // filename From Content-Disposition
	UINT m_AttachmentNamePageCode2;
	CString m_MessageId;
	CString m_ReplyId;
	CString m_InReplyId;
	//unsigned __int64 m_ThreadId;
	CString m_ThreadId;
};

class MailBodyPool;

class MailBody :public MailHeader
{
	friend class MailBodyPool;
public:
	MailBody() { m_bodyDataOffset = 0; };
	~MailBody() { DeleteAll(); };

	int Load(char *& pszDatabase, const char* pszData, int nDataSize);
	BOOL AssertData(MboxMail *mail);

	typedef list<MailBody*> MailBodyList;
	int GetBodyPartList(MailBodyList& rList);
	int m_bodyDataOffset;
	int m_bodyDataLength;

protected:
	char* FindBoundary(const char* pszData, const char* pszDataEnd, const char* boundary, int boundaryLength);
	const char* FindString(const char* pszStr1, const char* pszStr2, const char* pszEnd);

	DLLIST_NODE(MailBody) m_freeList_link;
	DLLIST(MailBody, m_freeList_link) m_listBodies;

	void DeleteAllParts();
	void DeleteAll();

	static MailBody *CreatePart();
	static void ErasePart(MailBody* body);

public:
	static MailBody *CreateMailBody();
	static void FreeMailBody(MailBody* body);
	static MailBodyPool *m_mpool;
};

class MailBodyPool
{
public:
	MailBodyPool() {};
	~MailBodyPool();

	MailBody *AllocPart();
	void FreePart(MailBody* body);

	DLLIST(MailBody, m_freeList_link) m_freeMailBodyList;
};

