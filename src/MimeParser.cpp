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
#include "TextUtilsEx.h"
#include "MimeParser.h"

#include "MainFrm.h"
#include "MboxMail.h"


#ifdef _DEBUG
inline void BreakParser()
{
	int deb = 1;
}
#else
inline void BreakParser() {}
#endif

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
	m_AttachmentName2.Empty();
	m_Name.Empty();
	m_MessageId.Empty();
	m_ReplyId.Empty();
	m_InReplyId.Empty();
	//m_threadId = 0;
	m_ThreadId.Empty();
}

int MailHeader::Load(const char* pszData, int nDataSize)
{
	static const char *cType = "content-type:";
	static const int cTypeLen = istrlen(cType);
	static const char *cContentLocation = "content-location:";
	static const int cContentLocationLen = istrlen(cContentLocation);
	static const char *cContentId = "content-id:";
	static const int cContentIdLen = istrlen(cContentId);
	static const char *cTransferEncoding = "content-transfer-encoding:";
	static const int cTransferEncodingLen = istrlen(cTransferEncoding);
	static const char *cDisposition = "content-disposition:";
	static const int cDispositionLen = istrlen(cDisposition);
	static const char *cAttachment = "attachment";
	static const int cAttachmentLen = istrlen(cAttachment);
	static const char *cName = "name";
	static const int cNameLen = istrlen(cName);
	static const char *cFileName = "filename";
	static const int cFileNameLen = istrlen(cFileName);
	static const char *cFileNameStar = "filename*";
	static const int cFileNameStarLen = istrlen(cFileNameStar);

	static const char *cMsgId = "message-id:";
	static const int cMsgIdLen = istrlen(cMsgId);

	static const char *cReferences = "references:";
	static const int cReferencesLen = istrlen(cReferences);

	static const char *cReplyId = "reply-to:";
	static const int cReplyIdLen = istrlen(cReplyId);

	static const char *cInReplyId = "in-reply-to:";
	static const int cInReplyIdLen = istrlen(cInReplyId);

	static const char *cAlternative = "alternative";
	static const int cAlternativeLen = istrlen(cAlternative);

	static const char *cMultipart = "multipart";
	static const int cMultipartLen = istrlen(cMultipart);

	static const char *cBoundary = "boundary";
	static const int cBoundaryLen = istrlen(cBoundary);

	static const char *cText = "text";
	static const int cTextLen = istrlen(cText);

	static const char *cCharset = "charset";
	static const int cCharsetLen = istrlen(cCharset);

	static const char *cThreadId = "x-gm-thrid:";
	static const int cThreadIdLen = istrlen(cThreadId);

	int contentIndex = 0;
	int contentLength = 0;

	CStringA line;
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
			int headLength = IntPtr2Int(p - pszData);
			return headLength;
			break;  // end of header
		}
		else if (TextUtilsEx::strncmpUpper2Lower(p, e, cType, cTypeLen) == 0)
		{
			BreakParser();
			p = MimeParser::GetMultiLine(p, e, line);
			MimeParser::GetFieldValue(line, cTypeLen, m_ContentType);

			if (TextUtilsEx::strncmpUpper2Lower((char*)(LPCSTR)m_ContentType, m_ContentType.GetLength(), cMultipart, cMultipartLen) == 0)
			{
				m_IsMultiPart = true;
				int ret = MimeParser::GetParamValue(line, cTypeLen, cBoundary, cBoundaryLen, m_Boundary);

			}
			else if (TextUtilsEx::strncmpUpper2Lower((char*)(LPCSTR)m_ContentType, m_ContentType.GetLength(), cText, cTextLen) == 0)
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
				int ret = MimeParser::GetParamValue(line, cTypeLen, cCharset, cCharsetLen, m_Charset);
				if (!m_Charset.IsEmpty())
				{
					m_PageCode = TextUtilsEx::StrPageCodeName2PageCode(m_Charset);
					if (m_PageCode > CP_UTF8)
						int deb = 1;
					if ((m_PageCode == 0) && m_IsTextHtml)
						int deb = 1;
				}
				ret = MimeParser::GetParamValue(line, cTypeLen, cName, cNameLen, m_Name);
				if (!m_Name.IsEmpty())
				{
					UINT toCharsetId = 0; // not used anyway
					CStringA charset;
					UINT charsetId = 0;
					DWORD error;
					CStringA Name = TextUtilsEx::DecodeString(m_Name, charset, charsetId, toCharsetId, error);
					// TODO: what about charset and charsetId :)
					m_Name = Name;
					m_NamePageCode = charsetId;
					//m_IsAttachment = true;  // TODO: Caller need to decide
				}
				int deb = 1;
			}
			else
			{
				BreakParser();
				int ret = MimeParser::GetParamValue(line, cTypeLen, cName, cNameLen, m_Name);
				if (!m_Name.IsEmpty()) 
				{
					UINT toCharsetId = 0; // not used anyway
					CStringA charset;
					UINT charsetId = 0;
					DWORD error;
					CString Name = TextUtilsEx::DecodeString(m_Name, charset, charsetId, toCharsetId, error);
					// TODO: what about charset and charsetId :)
					m_Name = Name;
					m_NamePageCode = charsetId;
					//m_IsAttachment = true;  // TODO: Caller need to decide
				}
			}
		}
		else if (TextUtilsEx::strncmpUpper2Lower(p, e, cTransferEncoding, cTransferEncodingLen) == 0)
		{
			BreakParser();
			p = MimeParser::GetMultiLine(p, e, line);
			MimeParser::GetFieldValue(line, cTransferEncodingLen, m_TransferEncoding);
		}
		else if (TextUtilsEx::strncmpUpper2Lower(p, e, cDisposition, cDispositionLen) == 0)  // Begin Disposition
		{
			BreakParser();
			p = MimeParser::GetMultiLine(p, e, line);
			MimeParser::GetFieldValue(line, cDispositionLen, m_Disposition);
#if 1
			// For testing; remove when no longer needed
			// Maybe we should run this only if both regular Filename and Name are missing/empty?
			int posret = line.Find(cFileNameStar);
			if (posret > 0)
			{
				CStringA fileName;
				BOOL hasCharset;
				// TODO: Verify ; Concatenate multiple filename*0 , filename*1 segments if present; make sure it is safe before you anable 
				int retval = MimeParser::GetFilenameParamValue(line, cDispositionLen, cFileNameStar, cFileNameStarLen, fileName, hasCharset);
				TRACE("NAME  = %s\n", m_Name);
				TRACE("FILENAME  = %s\n", line);
				TRACE("FILENAME* = %s\n", fileName);

				CStringA charset;
				UINT charsetId = 0;

				CStringA outString;
				if (hasCharset)
				{
					// example string UTF-8''Copy%20of%20JOHN%20KOWALSKI%20SECOND%20PRIORITIES%20Rev%201%20with%20effort%20cost.xlsx
					// value looks like URL, may contatin %20 for example
					// can value look like URL but without charset ?? try to verify
					int retdecode = TextUtilsEx::DecodeMimeChunkedString(fileName, charset, charsetId, hasCharset, outString);
					if (retdecode)
					{
						if (charsetId == CP_UTF8)
						{
							fileName.Empty();
							fileName.Append(outString);
						}
						else
						{
							// Convert to UTF-8 or don't touch ??
							CStringA tostr;
							DWORD error;
							BOOL retS2U = TextUtilsEx::Str2UTF8((LPCSTR)outString, outString.GetLength(), charsetId, tostr, error);
							if (retS2U)
							{
								fileName = tostr;
								charsetId = CP_UTF8;
							}

						}
					}
				}
				else
				{
					int retdecode = TextUtilsEx::DecodeMimeChunkedString(fileName, charset, charsetId, hasCharset, outString);
					if (retdecode && fileName.CompareNoCase(outString))
					{
						_ASSERTE(FALSE);

						fileName.Empty();
						fileName.Append(outString);
					}
				}

				// We prefer filename but in this case we may prefer Name if not empty
				// fileName is local here and not used later yet ??
				// Investigate more Concatenate multiple filename*0, etc
				if (!m_Name.IsEmpty())
				{
					if (m_Name.CompareNoCase(fileName) == 0)
						TRACE("Name and Filename match\n");
					else
					{
						_ASSERTE(FALSE);
						TRACE("Name and Filename don't match !!!!!!\n");
					}
				}
				int deb = 1;
			}

#endif
			if (TextUtilsEx::strncmpUpper2Lower((char*)(LPCSTR)m_Disposition, m_Disposition.GetLength(), cAttachment, cAttachmentLen) == 0)
			{
				// TODO:  Maybe. There are plenty of irregular mails and they can be considered as inline and 
				// possibly both inline and attachment
				m_IsAttachment = true;
			}
			// attachment type or not, get the filename/attachment name 
			int ret = MimeParser::GetParamValue(line, cTypeLen, cFileName, cFileNameLen, m_AttachmentName);
			if (!m_AttachmentName.IsEmpty())
			{
				UINT toCharsetId = 0; // not used anyway
				CStringA charset;
				UINT charsetId = 0;
				DWORD error;
				CStringA attachmentName = TextUtilsEx::DecodeString(m_AttachmentName, charset, charsetId, toCharsetId, error);
				// TODO: what about charset and charsetId :)

				m_AttachmentNamePageCode = charsetId;
				m_AttachmentName = attachmentName;
			}
			else if (m_Name.IsEmpty())
			{
				// For now run this only if both regular Filename and Name are empty?
				int posret = line.Find(cFileNameStar);
				if (posret > 0)
				{
					CStringA fileName;
					BOOL hasCharset;
					// TODO: Concatenate multiple filename*0 , filename*1 segments if present;
					int retval = MimeParser::GetFilenameParamValue(line, cDispositionLen, cFileNameStar, cFileNameStarLen, fileName, hasCharset);

					CStringA charset;
					UINT charsetId = 0;

					CStringA attachmentName;
					if (hasCharset)
					{
						// example string UTF-8''Copy%20of%20JOHN%20KOWALSKI%20SECOND%20PRIORITIES%20Rev%201%20with%20effort%20cost.xlsx
						// value looks like URL, may contatin %20 for example
						// can value look like URL but without charset ?? try to verify
						int retdecode = TextUtilsEx::DecodeMimeChunkedString(fileName, charset, charsetId, hasCharset, attachmentName);
						fileName.Empty();
						fileName.Append(attachmentName);

						m_AttachmentNamePageCode2 = charsetId;
						m_AttachmentName2 = attachmentName;
					}
					else
					{
						;// ???
					}
					int deb = 1;
				}
			}
		}  // end of Disposition
		else if (TextUtilsEx::strncmpUpper2Lower(p, e, cMsgId, cMsgIdLen) == 0) {
			BreakParser();
			p = MimeParser::GetMultiLine(p, e, line);
			MimeParser::GetMessageId(line, cMsgIdLen, m_MessageId);
		}
		else if (TextUtilsEx::strncmpUpper2Lower(p, e, cContentId, cContentIdLen) == 0) {
			BreakParser();
			p = MimeParser::GetMultiLine(p, e, line);
			MimeParser::GetMessageId(line, cContentIdLen, m_ContentId);
			m_ContentId.Trim();
			m_ContentId.Trim("<>");

		}
		else if (TextUtilsEx::strncmpUpper2Lower(p, e, cInReplyId, cInReplyIdLen) == 0) {
			BreakParser();
			p = MimeParser::GetMultiLine(p, e, line);
			MimeParser::GetMessageId(line, cInReplyIdLen, m_InReplyId);
		}
		else if (TextUtilsEx::strncmpUpper2Lower(p, e, cContentLocation, cContentLocationLen) == 0)
		{
			BreakParser();
			char *p_save = p;
			p = MimeParser::GetMultiLine(p, e, line);
			MimeParser::GetFieldValue(line, cContentLocationLen, m_ContentLocation);
			if (m_ContentLocation.GetLength() > 50)
				int deb = 1;
		}
		else if (TextUtilsEx::strncmpUpper2Lower(p, e, cThreadId, cThreadIdLen) == 0) {
			BreakParser();
			p = MimeParser::GetMultiLine(p, e, line);
			//imeParser::GetThreadId(line, cThreadIdLen, m_ThreadId);
			m_ThreadId = line.Mid(cThreadIdLen);
			m_ThreadId.Trim();
		}
#if 0
		// Below doesn't always help, commented out for now
		else if (TextUtilsEx::strncmpUpper2Lower(p, e, cReferences, cReferencesLen) == 0) {
			BreakParser();
			// Kludge until we fully process all relevant fields, Message-ID, In-Reply-To and Reply-To
			if (m_InReplyId.IsEmpty())
			{
				p = MimeParser::GetMultiLine(p, e, line);
				MimeParser::GetMessageId(line, cReferencesLen, m_InReplyId);
			}
			else
				p = EatNLine(p, e);

		}
#endif
		else {
			p = EatNLine(p, e);
		}
	}
	if ((m_PageCode == 0) && m_IsTextHtml)
		int deb = 1;
	int headLength = IntPtr2Int(p - pszData);
	return headLength;
}

bool MailHeader::IsAttachment()
{
	if (!m_Name.IsEmpty() ||
		!m_AttachmentName.IsEmpty() ||
		!m_ContentId.IsEmpty() ||
		!m_ContentLocation.IsEmpty() ||
		(m_Disposition.CompareNoCase("attachment") == 0)
		)
		return true;
	else
		return false;
}

// Followed the design of CMimeBody::Load in mime.cpp, so ite should be reliable :)
// This parser doesn't decode body parts. It only determines length and offsets to the body parts.
int MailBody::Load(char *& pszDataBase, const char* pszData, int nDataSize)
{
	//
	int nSize = MailHeader::Load(pszData, nDataSize);
	if (nSize <= 0)
		return nSize;

	if (!m_AttachmentName2.IsEmpty())
	{
		if (m_Name.IsEmpty() && m_AttachmentName.IsEmpty())
		{
			m_AttachmentName.Append(m_AttachmentName2);
			m_AttachmentNamePageCode = m_AttachmentNamePageCode2;
		}
	}

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
		m_bodyDataOffset = IntPtr2Int(pszData - pszDataBase);
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
	_ASSERTE(m_Boundary.GetLength() > 0);

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
		if (TextUtilsEx::strncmpExact(p, e, bPrefix, 2) == 0)
		{
			if (TextUtilsEx::strncmpExact(&p[2], e, boundary, boundaryLength) == 0)
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
	pszEnd -= istrlen(pszStr2);
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

#if 0

August 13, 1982 - 5 - RFC #822

UNFOLDING Long Lines

The process of moving  from  this  folded   multiple - line
representation  of a header field to its single line represen -
tation is called "unfolding".Unfolding  is  accomplished  by
regarding   CRLF   immediately  followed  by  a  LWSP - char  as
equivalent to the LWSP - char.

Note:  While the standard  permits  folding  wherever  linear -
	white - space is permitted, it is recommended that struc -
	tured fields, such as those containing addresses, limit
	folding  to higher - level syntactic breaks.For address
	fields, it  is  recommended  that  such  folding  occur
	between addresses, after the separating comma.
#endif

char *MimeParser::GetMultiLine(char *p, char *e, CStringA &line, BOOL raw)
{
	// TODO: not the most efficient implementation
	char *p_beg = p;
	p = EatNewLine(p, e);

	char *ss = line.GetBufferSetLength(IntPtr2Int(p - p_beg));
	::memcpy(ss, p_beg, p - p_beg);

	line.TrimLeft();
	if (!raw)
		line.TrimRight("\r\n");

	if (line.IsEmpty())
		return p;

	char c = line.GetAt(line.GetLength() - 1);
	while ((p < e) && ((*p == ' ') || (*p == '\t')))
	{
		if ((c == ' ') || (c == '\t'))
			int deb = 1;

		char *p_next_beg = p;
		p = EatNewLine(p, e);

		CStringA nextLine;
		char *ss = nextLine.GetBufferSetLength(IntPtr2Int(p - p_next_beg));
		::memcpy(ss, p_next_beg, p - p_next_beg);
		if (!raw)
		{
			nextLine.TrimLeft();
			nextLine.TrimRight("\r\n");

			line += " " + nextLine;  // this is according to spec but it doesn't work in many cases  FIXME
			//line += nextLine;
		}
		else
		{
			line += nextLine;
		}
	}
	line.Trim();
	return p;
};

int MimeParser::GetFieldValue(CStringA &fieldLine, int startPos, CStringA &value)
{
	int posEnd = fieldLine.FindOneOf(";\n\r");
	if (posEnd < 0)
		value = fieldLine.Mid(startPos);
	else
		value = fieldLine.Mid(startPos, posEnd - startPos);
	value.Trim();
	return 1;
}
int MimeParser::GetMessageId(CStringA& fieldLine, int startPos, CStringA& value)
{
	int next_startPos = MimeParser::GetMessageId((char*)((LPCSTR)fieldLine), startPos, value);
	return next_startPos;
}

int MimeParser::GetMessageId(char* fldLine, int startPos, CStringA &value)
{
	char* str = fldLine;
	str += startPos;
	char ch = *str;
	const char* first = 0;
	const char* last = 0;
	value.Empty();
	while ((*str != 0) && (*str != '<') && (*str != '\n') && (*str != '\r'))
	{
		str++;
	}
	if (*str != '<')
	{
		return -1;
	}
	first = str;
	while ((*str != 0) && (*str != '>') && (*str != '\n') && (*str != '\r'))
	{
		str++;
	}
	if (*str != '>')
	{
		return -1;
	}
	last = str;
	int len = (int)(last - first + 1);
	value.Append(first, len);

	int pos = (int)(last - fldLine + 1);
	return pos;
}

int MimeParser::GetMessageIdList(CStringA& fieldLine, int startPos, MessageIdList &msgIdList)
{
	CStringA value;
	char* buffer = (char*)((LPCSTR)fieldLine);
	while (startPos >= 0)
	{
		startPos = MimeParser::GetMessageId(buffer, startPos, value);
		if (startPos > 0)
			msgIdList.push_back(value);
	}

	return (int)msgIdList.size();
}

int MimeParser::GetThreadId(CStringA &fieldLine, int startPos, CStringA &value)
{
	int posEnd = fieldLine.FindOneOf(";\n\r");
	if (posEnd < 0)  // indicates problem, '>' expected
		value = fieldLine.Mid(startPos);
	else
		value = fieldLine.Mid(startPos, posEnd - startPos);
	value.Trim();
	return 1;
}

int MimeParser::GetThreadId(CStringA &fieldLine, int startPos, unsigned __int64 &value)
{
	const char *pos = (LPCSTR)fieldLine + startPos;
	value = _atoi64(pos);
	return 1;
}

int MimeParser::GetParamValue(CStringA &fieldLine, int startPos, const char *param, int paramLen, CStringA &value)
{
	// TODO: it will break if ';' is part of the value. Need to handle quoted values containing ';' ?
	value.Empty();

	char *pbegin_sv = (char*)(LPCSTR)fieldLine;
	char *pend_sv = pbegin_sv + fieldLine.GetLength();
	char *p = pbegin_sv + startPos;

	// or keep it simple and require that *p == '=' at this point ?
	while (p < pend_sv)
	{
		//p = strstr(p, param);
		p = TextUtilsEx::strnstrUpper2Lower(p, pend_sv, param, paramLen);
		if (p == 0)
			return 0;

		p = p + paramLen;
		// SkipWhite ??
		char c = *p;
		if (c == '=')
			break;
	}

	char *pend = strchr(p, ';');
	if (pend == 0)
		pend = pend_sv;

	while (p < pend)
	{
		if (*p++ == '=')
			break;
	}
	if (p >= pend)
		return 0;

	// handle quotes strings
	char * pEndQoute = 0;
	if (*p == '"')
	{
		pEndQoute = strchr(p+1, '"');
		if (pEndQoute)
			pend = pEndQoute;
		else
			; // ???
	}

	char *posBegin = p;
	char *posEnd = pend;

	value = fieldLine.Mid(IntPtr2Int(posBegin - pbegin_sv), IntPtr2Int(posEnd - posBegin));

	value.Trim("\"\t ");

	return 1;
}

int MimeParser::GetFilenameParamValue(CStringA &fieldLine, int startPos, const char *param, int paramLen, CStringA &value, BOOL &hasCharset)
{
	// TODO: it will break if ';' is part of the value. Need to handle quoted values containing ';' ?
	value.Empty();

	CStringA part;
	hasCharset = FALSE;
	int hasCharsetDone = FALSE;
	int lastPos = fieldLine.GetLength();
	while ((startPos > 0) && (startPos < lastPos))
	{
		BOOL has_charset = FALSE;
		int nextPos = GetFilenameParamPartValue(fieldLine, startPos, param, paramLen, part, has_charset);
		if (!part.IsEmpty())
		{
			value.Append(part);
			startPos = nextPos;
			if (has_charset && !hasCharsetDone)
			{
				hasCharset = TRUE;
				hasCharsetDone = TRUE;
			}
		}
		else
			break;
	}
	value.Trim("\"\t ");
	if (value.IsEmpty())
		return 0;
	else
		return 1;
}

int MimeParser::GetFilenameParamPartValue(CStringA &fieldLine, int startPos, const char *param, int paramLen, CStringA &value, BOOL &hasCharset)
{
	// TODO: it will break if ';' is part of the value. Need to handle quoted values containing ';' ?
	value.Empty();
	hasCharset = FALSE;

	int nextPos = 0;
	char *pbegin_sv = (char*)(LPCSTR)fieldLine;
	char *pend_sv = pbegin_sv + fieldLine.GetLength();
	char *p = pbegin_sv + startPos;

	// or keep it simple and require that *p == '=' at this point ?
	while (p < pend_sv)
	{
		//p = strstr(p, param);
		p = TextUtilsEx::strnstrUpper2Lower(p, pend_sv, param, paramLen);
		if (p == 0)
			return -1;

		p = p + paramLen;
		// TODO: SkipWhite
		char c = *p;
		if (c == '=')
		{
			hasCharset = TRUE;
			break;
		}
		else
		{
			if (isdigit(c))
			{
				p++;
				c = *p;
				if (c == '=')
				{
					break;
				}
				else if (c == '*')
				{
					p++;
					c = *p;
					if (c == '=')
					{
						hasCharset = TRUE;
						break;
					}
					else
						return -1;
				}
				else
					return -1;
			}
			else
				return -1;
		}
	}
	// TODO: SkipWhite

	if (p >= pend_sv)
		return -1;

	char *pend_semicolon = strchr(p, ';');

	// this is hack or work around when segment is not terminated with ;
	char *pend_filename = TextUtilsEx::strnstrUpper2Lower(p, pend_sv, param, paramLen);

	char *pend = 0;
	if (pend_semicolon && pend_filename)
	{
		if (pend_semicolon < pend_filename)
			pend = pend_semicolon;
		else
		{
			pend = pend_filename;
			pend--; // remove single  space due to unfolding
		}
	}
	else if (pend_semicolon)
		pend = pend_semicolon;
	else if (pend_filename)
		pend--; // remove single  space due to unfolding
	else
		pend = pend_sv;

	nextPos = IntPtr2Int(pend - pbegin_sv);

	while (p < pend)
	{
		if (*p++ == '=')
			break;
	}
	if (p >= pend)
		return -1;

	// handle quotes strings
	char * pEndQoute = 0;
	if (*p == '"')
	{
		pEndQoute = strchr(p + 1, '"');
		if (pEndQoute)
			pend = pEndQoute;
		else
			; // ???
	}

	char *posBegin = p;
	char *posEnd = pend;

	CStringA part = fieldLine.Mid(IntPtr2Int(posBegin - pbegin_sv), IntPtr2Int(posEnd - posBegin));
	value.Append(part);

	value.Trim("\"");

	return nextPos;
}

#if 0
char *MimeParser::EatNewLine(char* p, char*e) 
{
	while ((p < e) && (*p++ != '\n'));
	return p;
}
#endif

BOOL MimeParser::isEmptyLine(const char* p, const char* e)
{
	while ((p < e) && ((*p == '\r') || (*p == '\n') || (*p == ' ') || (*p == '\t')))  // eat white
		p++;
	if (p == e)
		return TRUE;
	else
		return FALSE;
}

char* MimeParser::SkipEmptyLines(const char* p, const char* e)
{
	while (p < e)
	{
		if ((*p == '\r') || (*p == '\n') || (*p == ' ') || (*p == '\t'))  // eat white
			p++;
		else
			return (char*)p;
	};
	return (char*)p;
}

char *MimeParser::EatNewLine(char* p, const char* e, bool &isEmpty)
{
	// mail text is not always formatted well
	// below will handle multiple CRs in a row and CR but no NL that follows to support mails not well formatted
	isEmpty = true;
	while (p < e)
	{
		while ((p < e) && (*p != '\n') && (*p != '\r'))
		{
			if (isEmpty)
			{
				if ((*p == ' ') || (*p == '\t') || (*p == '\r') || (*p == '\n'))
					;
				else
					isEmpty = false;
			}
			p++;
		}
		if (p >= e)
			return p;

		if (*p == '\n')  // jst NL nor CR NL
			return ++p;

		// Found CR, skip
		p++;
		if (p >= e)
			return p;

		if (*p == '\n')  // CR NL
			return ++p;

		// handle case abcCRxyz; return pointer to xyz
		// Should caller fix this and always pass abcCRNLxyz FIXME
		if (*p != '\r')  // CR no NL
			return p;

		// *p == '\r'  continue to discard multiple CRs
	}

	return p;
}

char *MimeParser::EatNewLine(char* p, char*e)
{
#if 0
	while ((p < e) && (*p != '\n')) { p++; }
	if (p < e)
		return ++p;
	else
		return p;
#else
	// mail text is not always formatted well
	// below will handle multiple CRs in a row and CR but no NL that follows to support mails not well formatted
	while (p < e)
	{
		while ((p < e) && (*p != '\n') && (*p != '\r')) { p++; }
		if (p >= e)
			return p;

		if (*p == '\n')  // jst NL nor CR NL
			return ++p;

		// Found CR, skip
		p++;
		if (p >= e)
			return p;

		if (*p == '\n')  // CR NL
			return ++p;

		// handle case abcCRxyz; return pointer to xyz
		// Should caller fix this and always pass abcCRNLxyz FIXME
		if (*p != '\r')  // CR no NL
			return p;

		// *p == '\r'  continue to discard multiple CRs
	}

	return p;
#endif
}

char *MimeParser::EatNewLine(char* p, char*e, int &maxLineLength)
{
#if 0
	while ((p < e) && (*p != '\n') && (maxLineLength > 0))
	{
		p++;
		maxLineLength--;
	}
	if (p < e)
		return ++p;
	else
		return p;

#else
	// mail text is not always formatted well
	// below will handle multiple CRs in a row and CR but no NL that follows to support mails not well formatted
	while (p < e)
	{
		while ((p < e) && (*p != '\n') && (*p != '\r') && (maxLineLength > 0))
		{
			p++;
			maxLineLength--;
		}
		if (p >= e)
			return p;

		if (*p == '\n')  // jst NL nor CR NL
			return ++p;

		// Found CR, skip
		p++;
		maxLineLength--;

		if (p >= e)
			return p;

		if (*p == '\n')  // CR NL
			return ++p;

		// handle case abcCRxyz; return pointer to xyz
		// Should caller fix this and always pass abcCRNLxyz FIXME
		if (*p != '\r')  // CR no NL
			return p;

		// *p == '\r'  continue to discard multiple CRs
	}
	if (maxLineLength <= 0)
		int deb = 1;
	return p;
#endif
}
