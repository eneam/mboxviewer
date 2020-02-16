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


#include "stdafx.h"
#include "MimeParser.h"

#include "MainFrm.h"
#include "MboxMail.h"


int GetMessageId(CString &fieldLine, int startPos, CString &value);
int GetParamValue(CString &fieldLine, int startPos, const char *param, int paramLen, CString &value);
int strncmpUpper2Lower(char *any, char *end, const char *lower, int lowerlength);
int   strncmpUpper2Lower(char *any, char *end, const char *lower, int lowerlength);
int strncmpUpper2Lower(char *any, int anyLength, const char *lower, int lowerlength);
int strncmpExact(char *any, char *end, const char *lower, int lowerlength);
char * GetMultiLine(char *p, char *e, CString &line);
int GetFieldValue(CString &fieldLine, int startPos, CString &value);


void MailHeader::Clear()
{
	m_IsText = false;
	m_IsTextPlain = false;
	m_IsTextHtml = false;
	m_IsMessage = false;
	m_IsAttachment = false;
	m_IsMultiPart = false;
	m_Charset.Empty();
	m_PageCode = 0;
	m_Description.Empty();
	m_Disposition.Empty();
	m_TransferEncoding.Empty();
	m_SubType.Empty();
	m_MainType.Empty();
	m_Boundary.Empty();
	m_ContentType.Empty();
	m_ContentId.Empty();
	m_ContentLocation.Empty();
	m_MediaType = CMimeHeader::MediaType::MEDIA_UNKNOWN;
	m_AttachmentName.Empty();
	m_Name.Empty();
	m_MessageId.Empty();
	m_ReplyId.Empty();
}

int MailHeader::Load(const char* pszData, int nDataSize)
{
	static const char *cType = "content-type:";
	static const int cTypeLen = strlen(cType);
	static const char *cContentLocation = "content-location:";
	static const int cContentLocationLen = strlen(cContentLocation);
	static const char *cContentId = "content-id:";
	static const int cContentIdLen = strlen(cContentId);
	static const char *cTransferEncoding = "content-transfer-encoding:";
	static const int cTransferEncodingLen = strlen(cTransferEncoding);
	static const char *cDisposition = "content-disposition:";
	static const int cDispositionLen = strlen(cDisposition);
	static const char *cAttachment = "attachment";
	static const int cAttachmentLen = strlen(cAttachment);
	static const char *cName = "name";
	static const int cNameLen = strlen(cName);
	static const char *cFileName = "filename";
	static const int cFileNameLen = strlen(cFileName);

	static const char *cMsgId = "message-id:";
	static const int cMsgIdLen = strlen(cMsgId);
	static const char *cReplyId = "in-reply-to:";
	static const int cReplyIdLen = strlen(cReplyId);

	static const char *cAlternative = "alternative";
	static const int cAlternativeLen = strlen(cAlternative);

	static const char *cMultipart = "multipart";
	static const int cMultipartLen = strlen(cMultipart);

	static const char *cBoundary = "boundary";
	static const int cBoundaryLen = strlen(cBoundary);

	static const char *cText = "text";
	static const int cTextLen = strlen(cText);

	static const char *cCharset = "charset";
	static const int cCharsetLen = strlen(cCharset);

	int contentIndex = 0;
	int contentLength = 0;

	CString line;
	char *p = (char*)pszData;
	char *e = (char*)pszData + nDataSize;

	while ((p < e) && ((*p == ' ') || (*p == '\t'))) // eat lines starting with ' ' or '\t'
		p = EatNLine(p, e);


	while ((p < e) && ((*p == '\r') || (*p == '\n')))  // eat empty lines
		p++;

	// make sure p is incremented within a loop or break
	while (p < e)
	{
		if (((*p == '\r') && (*(p + 1) == '\n')) || (*p == '\n')) {

			p = EatNLine(p, e);
			int headLength = p - pszData;
			return headLength;
			break;  // end of header
		}
		else if (strncmpUpper2Lower(p, e, cType, cTypeLen) == 0)
		{
			p = GetMultiLine(p, e, line);
			GetFieldValue(line, cTypeLen, m_ContentType);

			if (strncmpUpper2Lower((char*)(LPCSTR)m_ContentType, m_ContentType.GetLength(), cMultipart, cMultipartLen) == 0)
			{
				m_IsMultiPart = true;
				int ret = GetParamValue(line, cTypeLen, cBoundary, cBoundaryLen, m_Boundary);

			}
			else if (strncmpUpper2Lower((char*)(LPCSTR)m_ContentType, m_ContentType.GetLength(), cText, cTextLen) == 0)
			{
				m_IsText = true;
				if (m_ContentType.Compare("text") == 0)
					m_ContentType += "/plain";
				if (m_ContentType.Compare("text/plain") == 0) {
					m_IsTextPlain = true;
					m_IsTextHtml = false;
				}
				else if (m_ContentType.Compare("text/html") == 0) {
					m_IsTextHtml = true;
					m_IsTextPlain = false;
				}
				int ret = GetParamValue(line, cTypeLen, cCharset, cCharsetLen, m_Charset);
				if (!m_Charset.IsEmpty()) {
					m_PageCode = MboxMail::Str2PageCode(m_Charset);
					if (m_PageCode > CP_UTF8)
						int deb = 1;
					if ((m_PageCode == 0) && m_IsTextHtml)
						int deb = 1;
				}
				ret = GetParamValue(line, cTypeLen, cName, cNameLen, m_Name);
				if (!m_Name.IsEmpty()) {
					CString charset;
					UINT charsetId = 0;
					CString Name = MboxMail::DecodeString(m_Name, charset, charsetId);
					// TODO: what about charset and charsetId :)
					m_Name = Name;
					//m_IsAttachment = true;  // TODO: Caller need to decide
				}
				int deb = 1;
			}
			else
			{
				int ret = GetParamValue(line, cTypeLen, cName, cNameLen, m_Name);
				if (!m_Name.IsEmpty()) {
					CString charset;
					UINT charsetId = 0;
					CString Name = MboxMail::DecodeString(m_Name, charset, charsetId);
					// TODO: what about charset and charsetId :)
					m_Name = Name;
					//m_IsAttachment = true;  // TODO: Caller need to decide
				}
			}
		}
		else if (strncmpUpper2Lower(p, e, cTransferEncoding, cTransferEncodingLen) == 0)
		{
			p = GetMultiLine(p, e, line);
			GetFieldValue(line, cTransferEncodingLen, m_TransferEncoding);
		}
		else if (strncmpUpper2Lower(p, e, cDisposition, cDispositionLen) == 0)
		{
			p = GetMultiLine(p, e, line);
			GetFieldValue(line, cDispositionLen, m_Disposition);

			if (strncmpUpper2Lower((char*)(LPCSTR)m_Disposition, m_Disposition.GetLength(), cAttachment, cAttachmentLen) == 0)
			{
				// TODO:  Maybe. There are plenty of irregular mails and they can be considered as inline and 
				// possibly both inline and attachment
				m_IsAttachment = true;
			}
			// attachment type or not, get the filename/attachment name 
			int ret = GetParamValue(line, cTypeLen, cFileName, cFileNameLen, m_AttachmentName);
			if (!m_AttachmentName.IsEmpty()) {
				CString charset;
				UINT charsetId = 0;
				CString attachmentName = MboxMail::DecodeString(m_AttachmentName, charset, charsetId);
				// TODO: what about charset and charsetId :)
				m_AttachmentName = attachmentName;
			}
		}
		else if (strncmpUpper2Lower(p, e, cMsgId, cMsgIdLen) == 0) {
			p = GetMultiLine(p, e, line);
			GetMessageId(line, cMsgIdLen, m_MessageId);
		}
		else if (strncmpUpper2Lower(p, e, cContentId, cContentIdLen) == 0) {
			p = GetMultiLine(p, e, line);
			GetMessageId(line, cMsgIdLen, m_ContentId);
			m_ContentId.Trim();
			m_ContentId.Trim("<>");

		}
		else if (strncmpUpper2Lower(p, e, cReplyId, cReplyIdLen) == 0) {
			p = GetMultiLine(p, e, line);
			GetMessageId(line, cReplyIdLen, m_ReplyId);
		}
		else if (strncmpUpper2Lower(p, e, cContentLocation, cContentLocationLen) == 0)
		{
			char *p_save = p;
			p = GetMultiLine(p, e, line);
			GetFieldValue(line, cContentLocationLen, m_ContentLocation);
			if (m_ContentLocation.GetLength() > 50)
				int deb = 1;
		}
		else {
			p = EatNLine(p, e);
		}
	}
	if ((m_PageCode == 0) && m_IsTextHtml)
		int deb = 1;
	int headLength = p - pszData;
	return headLength;
}

// Followed the design of CMimeBody::Load in mime.cpp, so ite should be reliable :)
// This parser doesn't decode body parts. It only determines length and offsets to the body parts.
int MailBody::Load(char *& pszDataBase, const char* pszData, int nDataSize)
{
	//
	int nSize = MailHeader::Load(pszData, nDataSize);
	if (nSize <= 0)
		return nSize;

	const char* pszDataBegin = pszData;	// preserve start position
	pszData += nSize;
	nDataSize -= nSize;

	// determine the length of the content
	const char* pszEnd = pszData + nDataSize;

	if (IsMultiPart())
	{
		// find the begin boundary
		if (!m_Boundary.IsEmpty())
		{
			// Should match against "\--" or just "--"
			pszEnd = FindBoundary(pszData, pszEnd, m_Boundary, m_Boundary.GetLength());
			if (!pszEnd)
				pszEnd = pszData + nDataSize;
		}
	}

	// load content
	nSize = (int)(pszEnd - pszData);
	if (nSize > 0)
	{
		// save offset and length of body data
		m_bodyDataOffset = pszData - pszDataBase;
		m_bodyDataLength = nSize;
		pszData += nSize;
		nDataSize -= nSize;

		if (m_ContentType.IsEmpty()) {
			m_ContentType = "text/plain";
			m_IsText = true;
			m_IsTextPlain = true;
			m_IsTextHtml = false;
		}
	}
	if (nDataSize <= 0)
		return (int)(pszData - pszDataBegin);

	// load child body parts
	ASSERT(m_Boundary.GetLength() > 0);

	pszEnd = pszData + nDataSize;
	const char* pszBound1 = FindBoundary(pszData, pszEnd, m_Boundary, m_Boundary.GetLength());

	while (pszBound1 != NULL && pszBound1 < pszEnd)
	{
		const char* pszStart = EatNLine((char*)pszBound1, (char*)pszEnd);
		if (!pszStart)
			break;
		if (pszBound1[m_Boundary.GetLength() + 2] == '-' && pszBound1[m_Boundary.GetLength() + 2 + 1] == '-')
			return (int)(pszStart - pszDataBegin);	// reach the closing boundary

													// look for the next boundary
		const char* pszBound2 = FindBoundary(pszStart, pszEnd, m_Boundary, m_Boundary.GetLength());
		if (!pszBound2)				// overflow, boundary may be truncated
			pszBound2 = pszEnd;
		int nEntitySize = (int)(pszBound2 - pszStart);

		MailBody* pBP = CreatePart();
		m_listBodies.insert_tail(pBP);

		int nInputSize = pBP->Load(pszDataBase, pszStart, nEntitySize);
		if (nInputSize < 0)
		{
			ErasePart(pBP);
			return nInputSize;
		}
		pszBound1 = pszBound2;
	}
	return (int)(pszEnd - pszDataBegin);
}

// return a list of all child body parts belong to this body part
int MailBody::GetBodyPartList(MailBodyList& rList)
{
	int nCount = 0;

	if (!m_IsMultiPart)
	{
		rList.push_back(this);
		nCount++;
	}
	else
	{
		MailBody *iter;
		for (iter = m_listBodies.head(); iter != 0; iter = m_listBodies.next(iter))
		{
			nCount += iter->GetBodyPartList(rList);
		}
	}
	return nCount;
}

char* MailBody::FindBoundary(const char* pszBegin, const char* pszEnd, const char* boundary, int boundaryLength)
{
	static char *bPrefix = "--";
	register char *p = (char*)pszBegin;
	register char *e = (char*)pszEnd;
	while (p < e)
	{
		if (strncmpExact(p, e, bPrefix, 2) == 0)
		{
			if (strncmpExact(&p[2], e, boundary, boundaryLength) == 0)
				return  p;
		}
		p = EatNLine(p, e);
	}
	return 0;
}

MailBody* MailBody::CreatePart()
{
	MailBody* mbody = m_mpool->AllocPart();
	return mbody;
}


void MailBody::ErasePart(MailBody* mbody)
{
	mbody->DeleteAll();
	m_mpool->FreePart(mbody);
}

// public static
MailBody* MailBody::CreateMailBody()
{
	MailBody* mbody = m_mpool->AllocPart();
	return mbody;
}

// public static
void MailBody::FreeMailBody(MailBody* mbody)
{
	mbody->DeleteAll();
	m_mpool->FreePart(mbody);
}

MailBodyPool::~MailBodyPool()
{
	MailBody *body;
	while ((body = m_freeMailBodyList.remove_head()) != 0)
		delete body;
}

MailBody* MailBodyPool::AllocPart()
{
	MailBody* mbody;
	mbody = m_freeMailBodyList.remove_head();
	if (mbody)
		return mbody;
	else
		return new MailBody;
}


void MailBodyPool::FreePart(MailBody* mbody)
{
	m_freeMailBodyList.insert_head(mbody);
	int deb = 1;
}

// delete all child body parts
void MailBody::DeleteAllParts()
{
	MailBody *body;
	while ((body = m_listBodies.remove_head()) != 0)
	{
		body->DeleteAll();
		m_mpool->FreePart(body);
	}
}

void MailBody::DeleteAll()
{
	DeleteAllParts();
	MailHeader::Clear();
	m_bodyDataOffset = 0;
	m_bodyDataLength = 0;
}



// search for string2 in string1 (strstr)
const char* MailBody::FindString(const char* pszStr1, const char* pszStr2, const char* pszEnd)
{
	pszEnd -= ::strlen(pszStr2);
	const char *s1, *s2;
	while (pszStr1 <= pszEnd)
	{
		s1 = pszStr1;
		s2 = pszStr2;
		while (*s1 == *s2 && *s2)
			s1++, s2++;
		if (!*s2)
			return pszStr1;
		pszStr1++;
	}
	return NULL;
}
