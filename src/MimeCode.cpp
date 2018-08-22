//////////////////////////////////////////////////////////////////////
//
// MIME Encoding/Decoding:
//	Quoted-printable and Base64 for content encoding;
//	Encoded-word for header field encoding.
//
// Jeff Lee
// Dec 11, 2000
//
//////////////////////////////////////////////////////////////////////
#include "stdafx.h"
#pragma warning (disable : 4786)
#include "MimeCode.h"
#include "MimeChar.h"
#include "Mime.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// CMimeEnvironment - global environment to manage encoding/decoding
//////////////////////////////////////////////////////////////////////
bool CMimeEnvironment::m_bAutoFolding = false;
string CMimeEnvironment::m_strCharset;
list<CMimeEnvironment::CODER_PAIR> CMimeEnvironment::m_listCoders;
list<CMimeEnvironment::FIELD_CODER_PAIR> CMimeEnvironment::m_listFieldCoders;
list<CMimeEnvironment::MEDIA_TYPE_PAIR> CMimeEnvironment::m_listMediaTypes;
CMimeEnvironment CMimeEnvironment::m_globalMgr;

CMimeEnvironment::CMimeEnvironment()
{
	// initialize transfer encoding
	//REGISTER_MIMECODER("7bit", CMimeCode7bit);
	//REGISTER_MIMECODER("8bit", CMimeCode7bit);
	REGISTER_MIMECODER("quoted-printable", CMimeCodeQP);
	REGISTER_MIMECODER("base64", CMimeCodeBase64);

	// initialize header fields encoding
	REGISTER_FIELDCODER("Subject", CFieldCodeText);
	REGISTER_FIELDCODER("Comments", CFieldCodeText);
	REGISTER_FIELDCODER("Content-Description", CFieldCodeText);

	REGISTER_FIELDCODER("From", CFieldCodeAddress);
	REGISTER_FIELDCODER("To", CFieldCodeAddress);
	REGISTER_FIELDCODER("Resent-To", CFieldCodeAddress);
	REGISTER_FIELDCODER("Cc", CFieldCodeAddress);
	REGISTER_FIELDCODER("Resent-Cc", CFieldCodeAddress);
	REGISTER_FIELDCODER("Bcc", CFieldCodeAddress);
	REGISTER_FIELDCODER("Resent-Bcc", CFieldCodeAddress);
	REGISTER_FIELDCODER("Reply-To", CFieldCodeAddress);
	REGISTER_FIELDCODER("Resent-Reply-To", CFieldCodeAddress);

	REGISTER_FIELDCODER("Content-Type", CFieldCodeParameter);
	REGISTER_FIELDCODER("Content-Disposition", CFieldCodeParameter);
}

void CMimeEnvironment::SetAutoFolding(bool bAutoFolding)
{
	m_bAutoFolding = bAutoFolding;
	if (!bAutoFolding)
	{
		DEREGISTER_MIMECODER("7bit");
		DEREGISTER_MIMECODER("8bit");
	}
	else
	{
		REGISTER_MIMECODER("7bit", CMimeCode7bit);
		REGISTER_MIMECODER("8bit", CMimeCode7bit);
	}
}

void CMimeEnvironment::RegisterCoder(const char* pszCodingName, CODER_FACTORY pfnCreateObject/*=NULL*/)
{
	ASSERT(pszCodingName != NULL);
	list<CODER_PAIR>::iterator it = m_listCoders.begin();
	while (it != m_listCoders.end())
	{
		list<CODER_PAIR>::iterator it2 = it;
		it++;
		if (!::_stricmp(pszCodingName, (*it2).first))
			m_listCoders.erase(it2);
	}
	if (pfnCreateObject != NULL)
	{
		CODER_PAIR newPair(pszCodingName, pfnCreateObject);
		m_listCoders.push_front(newPair);
	}
}

CMimeCodeBase* CMimeEnvironment::CreateCoder(const char* pszCodingName)
{
	if (!pszCodingName || !::strlen(pszCodingName))
		pszCodingName = "7bit";

	for (list<CODER_PAIR>::iterator it=m_listCoders.begin(); it!=m_listCoders.end(); it++)
	{
		ASSERT((*it).first != NULL);
		if (!::_stricmp(pszCodingName, (*it).first))
		{
			CODER_FACTORY pfnCreateObject = (*it).second;
			ASSERT(pfnCreateObject != NULL);
			return pfnCreateObject();
		}
	}
	return new CMimeCodeBase;		// default coder for unregistered Content-Transfer-Encoding
}


void CMimeEnvironment::RegisterFieldCoder(const char* pszFieldName, FIELD_CODER_FACTORY pfnCreateObject/*=NULL*/)
{
	ASSERT(pszFieldName != NULL);
	list<FIELD_CODER_PAIR>::iterator it = m_listFieldCoders.begin();
	while (it != m_listFieldCoders.end())
	{
		list<FIELD_CODER_PAIR>::iterator it2 = it;
		it++;
		if (!::_stricmp(pszFieldName, (*it2).first))
			m_listFieldCoders.erase(it2);
	}
	if (pfnCreateObject != NULL)
	{
		FIELD_CODER_PAIR newPair(pszFieldName, pfnCreateObject);
		m_listFieldCoders.push_front(newPair);
	}
}

CFieldCodeBase* CMimeEnvironment::CreateFieldCoder(const char* pszFieldName)
{
	ASSERT(pszFieldName != NULL);
	for (list<FIELD_CODER_PAIR>::iterator it=m_listFieldCoders.begin(); it!=m_listFieldCoders.end(); it++)
	{
		ASSERT((*it).first != NULL);
		if (!::_stricmp(pszFieldName, (*it).first))
		{
			FIELD_CODER_FACTORY pfnCreateObject = (*it).second;
			ASSERT(pfnCreateObject != NULL);
			return pfnCreateObject();
		}
	}
	return new CFieldCodeBase;		// default coder for unregistered header fields
}

void CMimeEnvironment::RegisterMediaType(const char* pszMediaType, BODY_PART_FACTORY pfnCreateObject/*=NULL*/)
{
	ASSERT(pszMediaType != NULL);
	list<MEDIA_TYPE_PAIR>::iterator it = m_listMediaTypes.begin();
	while (it != m_listMediaTypes.end())
	{
		list<MEDIA_TYPE_PAIR>::iterator it2 = it;
		it++;
		if (!::_stricmp(pszMediaType, (*it2).first))
			m_listMediaTypes.erase(it2);
	}
	if (pfnCreateObject != NULL)
	{
		MEDIA_TYPE_PAIR newPair(pszMediaType, pfnCreateObject);
		m_listMediaTypes.push_front(newPair);
	}
}

CMimeBody* CMimeEnvironment::CreateBodyPart(const char* pszMediaType)
{
	if (!pszMediaType || !::strlen(pszMediaType))
		pszMediaType = "text";

	ASSERT(pszMediaType != NULL);
	for (list<MEDIA_TYPE_PAIR>::iterator it=m_listMediaTypes.begin(); it!=m_listMediaTypes.end(); it++)
	{
		ASSERT((*it).first != NULL);
		if (!::_stricmp(pszMediaType, (*it).first))
		{
			BODY_PART_FACTORY pfnCreateObject = (*it).second;
			ASSERT(pfnCreateObject != NULL);
			return pfnCreateObject();
		}
	}
	return new CMimeBody;			// default body part for unregistered media type
}

//////////////////////////////////////////////////////////////////////
// CMimeCode7bit - for 7bit/8bit encoding mechanism (fold long line)
//////////////////////////////////////////////////////////////////////
int CMimeCode7bit::GetEncodeLength() const
{
	int nSize = m_nInputSize + m_nInputSize / MAX_MIME_LINE_LEN * 4;
	//const unsigned char* pbData = m_pbInput;
	//const unsigned char* pbEnd = m_pbInput + m_nInputSize;
	//while (++pbData < pbEnd)
	//	if (*pbData == '.' && *(pbData-1) == '\n')
	//		nSize++;
	nSize += 4;
	return nSize;
}

int CMimeCode7bit::Encode(unsigned char* pbOutput, int nMaxSize) const
{
	const unsigned char* pbData = m_pbInput;
	const unsigned char* pbEnd = m_pbInput + m_nInputSize;
	unsigned char* pbOutStart = pbOutput;
	unsigned char* pbOutEnd = pbOutput + nMaxSize;
	unsigned char* pbSpace = NULL;
	int nLineLen = 0;
	while (pbData < pbEnd)
	{
		if (pbOutput >= pbOutEnd)
			break;

		unsigned char ch = *pbData;
		//if (ch == '.' && pbData-m_pbInput >= 2 && !::memcmp(pbData-2, "\r\n.", 3))
		//{
		//	*pbOutput++ = '.';		// avoid confusing with SMTP end flag
		//	nLineLen++;
		//}

		if (ch == '\r' || ch == '\n')
		{
			nLineLen = -1;
			pbSpace = NULL;
		}
		else if (nLineLen > 0 && CMimeChar::IsSpace(ch))
			pbSpace = pbOutput;

		// fold the line if it's longer than 76
		if (nLineLen >= MAX_MIME_LINE_LEN && pbSpace != NULL &&
			pbOutput+2 <= pbOutEnd)
		{
			int nSize = (int)(pbOutput - pbSpace);
			::memmove(pbSpace+2, pbSpace, nSize);
			*pbSpace++ = '\r';
			*pbSpace = '\n';
			pbSpace = NULL;
			nLineLen = nSize;
			pbOutput += 2;
		}

		*pbOutput++ = ch;
		pbData++;
		nLineLen++;
	}

	return (int)(pbOutput - pbOutStart);
}

//////////////////////////////////////////////////////////////////////
// CMimeCodeQP - for quoted-printable encoding mechanism
//////////////////////////////////////////////////////////////////////
int CMimeCodeQP::GetEncodeLength() const
{
	int nLength = m_nInputSize;
	const unsigned char* pbData = m_pbInput;
	const unsigned char* pbEnd = m_pbInput + m_nInputSize;
	while (pbData < pbEnd)
		if (!CMimeChar::IsPrintable(*pbData++))
			nLength += 2;
	//int nLength = m_nInputSize * 3;
	nLength += nLength / (MAX_MIME_LINE_LEN - 2) * 6;
	return nLength;
}

int CMimeCodeQP::Encode(unsigned char* pbOutput, int nMaxSize) const
{
	static const char* s_QPTable = "0123456789ABCDEF";

	const unsigned char* pbData = m_pbInput;
	const unsigned char* pbEnd = m_pbInput + m_nInputSize;
	unsigned char* pbOutStart = pbOutput;
	unsigned char* pbOutEnd = pbOutput + nMaxSize;
	unsigned char* pbSpace = NULL;
	int nLineLen = 0;
	while (pbData < pbEnd)
	{
		if (pbOutput >= pbOutEnd)
			break;

		unsigned char ch = *pbData;
		bool bQuote = false, bCopy = false;

		// According to RFC 2045, TAB and SPACE MAY be represented as the ASCII characters.
		// But it MUST NOT be so represented at the end of an encoded line.
		if (ch == '\t' || ch == ' ')
		{
			if (pbData == pbEnd-1 || (!m_bQuoteLineBreak && *(pbData+1) == '\r'))
				bQuote = true;		// quote the SPACE/TAB
			else
				bCopy = true;		// copy the SPACE/TAB
			if (nLineLen > 0)
				pbSpace = (unsigned char*) pbOutput;
		}
		else if (!m_bQuoteLineBreak && (ch == '\r' || ch == '\n'))
		{
			bCopy = true;			// keep 'hard' line break
			nLineLen = -1;
			pbSpace = NULL;
		}
		else if (!m_bQuoteLineBreak && ch == '.')
		{
			if (pbData-m_pbInput >= 2 &&
				*(pbData-2) == '\r' && *(pbData-1) == '\n' &&
				*(pbData+1) == '\r' && *(pbData+2) == '\n')
				bQuote = true;		// avoid confusing with SMTP's message end flag
			else
				bCopy = true;
		}
		else if (ch < 33 || ch > 126 || ch == '=')
			bQuote = true;			// quote this character
		else
			bCopy = true;			// copy this character

		if (nLineLen+(bQuote ? 3 : 1) >= MAX_MIME_LINE_LEN && pbOutput+3 <= pbOutEnd)
		{
			if (pbSpace != NULL && pbSpace < pbOutput)
			{
				pbSpace++;
				int nSize = (int)(pbOutput - pbSpace);
				::memmove(pbSpace+3, pbSpace, nSize);
				nLineLen = nSize;
			}
			else
			{
				pbSpace = pbOutput;
				nLineLen = 0;
			}
			::memcpy(pbSpace, "=\r\n", 3);
			pbOutput += 3;
			pbSpace = NULL;
		}

		if (bQuote && pbOutput+3 <= pbOutEnd)
		{
			*pbOutput++ = '=';
			*pbOutput++ = s_QPTable[(ch >> 4) & 0x0f];
			*pbOutput++ = s_QPTable[ch & 0x0f];
			nLineLen += 3;
		}
		else if (bCopy)
		{
			*pbOutput++ = (char) ch;
			nLineLen++;
		}

		pbData++;
	}

	return (int)(pbOutput - pbOutStart);
}

int CMimeCodeQP::Decode(unsigned char* pbOutput, int nMaxSize)
{
	const unsigned char* pbData = m_pbInput;
	const unsigned char* pbEnd = m_pbInput + m_nInputSize;
	unsigned char* pbOutStart = pbOutput;
	unsigned char* pbOutEnd = pbOutput + nMaxSize;

	while (pbData < pbEnd)
	{
		if (pbOutput >= pbOutEnd)
			break;

		unsigned char ch = *pbData++;
		if (ch == '=')
		{
			if (pbData+2 > pbEnd)
				break;				// invalid endcoding
			ch = *pbData++;
			if (CMimeChar::IsHexDigit(ch))
			{
				ch -= ch > '9' ? 0x37 : '0';
				*pbOutput = ch << 4;
				ch = *pbData++;
				ch -= ch > '9' ? 0x37 : '0';
				*pbOutput++ |= ch & 0x0f;
			}
			else if (ch == '\r' && *pbData == '\n')
				pbData++;			// Soft Line Break, eat it
			else					// invalid endcoding, let it go
				*pbOutput++ = ch;
		}
		else// if (ch != '\r' && ch != '\n')
			*pbOutput++ = ch;
	}

	return (int)(pbOutput - pbOutStart);
}

//////////////////////////////////////////////////////////////////////
// CMimeCodeBase64 - for base64 encoding mechanism
//////////////////////////////////////////////////////////////////////
int CMimeCodeBase64::GetEncodeLength() const
{
	int nLength = (m_nInputSize + 2) / 3 * 4;
	if (m_bAddLineBreak)
		nLength += (nLength / MAX_MIME_LINE_LEN + 1) * 2;
	return nLength;
}

int CMimeCodeBase64::GetDecodeLength() const
{
	return m_nInputSize * 3 / 4 + 2;
}

int CMimeCodeBase64::Encode(unsigned char* pbOutput, int nMaxSize) const
{
	static const char* s_Base64Table = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

	unsigned char* pbOutStart = pbOutput;
	unsigned char* pbOutEnd = pbOutput + nMaxSize;
	int nFrom, nLineLen = 0;
	unsigned char chHigh4bits = 0;

	for (nFrom=0; nFrom<m_nInputSize; nFrom++)
	{
		if (pbOutput >= pbOutEnd)
			break;

		unsigned char ch = m_pbInput[nFrom];
		switch (nFrom % 3)
		{
		case 0:
			*pbOutput++ = s_Base64Table[ch >> 2];
			chHigh4bits = (ch << 4) & 0x30;
			break;

		case 1:
			*pbOutput++ = s_Base64Table[chHigh4bits | (ch >> 4)];
			chHigh4bits = (ch << 2) & 0x3c;
			break;

		default:
			*pbOutput++ = s_Base64Table[chHigh4bits | (ch >> 6)];
			if (pbOutput < pbOutEnd)
			{
				*pbOutput++ = s_Base64Table[ch & 0x3f];
				nLineLen++;
			}
		}

		nLineLen++;
		if (m_bAddLineBreak && nLineLen >= MAX_MIME_LINE_LEN && pbOutput+2 <= pbOutEnd)
		{
			*pbOutput++ = '\r';
			*pbOutput++ = '\n';
			nLineLen = 0;
		}
	}

	if (nFrom % 3 != 0 && pbOutput < pbOutEnd)	// 76 = 19 * 4, so the padding wouldn't exceed 76
	{
		*pbOutput++ = s_Base64Table[chHigh4bits];
		int nPad = 4 - (nFrom % 3) - 1;
		if (pbOutput+nPad <= pbOutEnd)
		{
			::memset(pbOutput, '=', nPad);
			pbOutput += nPad;
		}
	}
	if (m_bAddLineBreak && nLineLen != 0 && pbOutput+2 <= pbOutEnd)	// add CRLF
	{
		*pbOutput++ = '\r';
		*pbOutput++ = '\n';
	}
	return (int)(pbOutput - pbOutStart);
}

int CMimeCodeBase64::Decode(unsigned char* pbOutput, int nMaxSize)
{
	const unsigned char* pbData = m_pbInput;
	const unsigned char* pbEnd = m_pbInput + m_nInputSize;
	unsigned char* pbOutStart = pbOutput;
	unsigned char* pbOutEnd = pbOutput + nMaxSize;

	int nFrom = 0;
	unsigned char chHighBits = 0;

	while (pbData < pbEnd)
	{
		if (pbOutput >= pbOutEnd)
			break;

		unsigned char ch = *pbData++;
		if (ch == '\r' || ch == '\n')
			continue;
		ch = (unsigned char) DecodeBase64Char(ch);
		if (ch >= 64)				// invalid encoding, or trailing pad '='
			break;

		switch ((nFrom++) % 4)
		{
		case 0:
			chHighBits = ch << 2;
			break;

		case 1:
			*pbOutput++ = chHighBits | (ch >> 4);
			chHighBits = ch << 4;
			break;

		case 2:
			*pbOutput++ = chHighBits | (ch >> 2);
			chHighBits = ch << 6;
			break;

		default:
			*pbOutput++ = chHighBits | ch;
		}
	}

	return (int)(pbOutput - pbOutStart);
}

//////////////////////////////////////////////////////////////////////
// CMimeEncodedWord - encoded word for non-ascii text (RFC 2047)
//////////////////////////////////////////////////////////////////////
int CMimeEncodedWord::GetEncodeLength() const
{
	if (!m_nInputSize || m_strCharset.empty())
		return CMimeCodeBase::GetEncodeLength();

	int nLength, nCodeLen = (int) m_strCharset.size() + 7;
	if (tolower(m_nEncoding) == 'b')
	{
		CMimeCodeBase64 base64;
		base64.SetInput((const char*)m_pbInput, m_nInputSize, true);
		base64.AddLineBreak(false);
		nLength = base64.GetOutputLength();
	}
	else
	{
		CMimeCodeQP qp;
		qp.SetInput((const char*)m_pbInput, m_nInputSize, true);
		qp.QuoteLineBreak(false);
		nLength = qp.GetOutputLength();
	}

	nCodeLen += 4;
	ASSERT(nCodeLen < MAX_ENCODEDWORD_LEN);
	return (nLength / (MAX_ENCODEDWORD_LEN - nCodeLen) + 1) * nCodeLen + nLength;
}

int CMimeEncodedWord::Encode(unsigned char* pbOutput, int nMaxSize) const
{
	if (m_strCharset.empty())
		return CMimeCodeBase::Encode(pbOutput, nMaxSize);

	if (!m_nInputSize)
		return 0;
	if (tolower(m_nEncoding) == 'b')
		return BEncode(pbOutput, nMaxSize);
	return QEncode(pbOutput, nMaxSize);
}

int CMimeEncodedWord::Decode(unsigned char* pbOutput, int nMaxSize)
{
	m_strCharset.erase();
	const char* pbData = (const char*) m_pbInput;
	const char* pbEnd = pbData + m_nInputSize;
	unsigned char* pbOutStart = pbOutput;
	while (pbData < pbEnd)
	{
		const char* pszHeaderEnd = pbData;
		const char* pszCodeEnd = pbEnd;
		int nCoding = 0, nCodeLen = (int)(pbEnd - pbData);
		if (pbData[0] == '=' && pbData[1] == '?')	// it might be an encoded-word
		{
			pszHeaderEnd = ::strchr(pbData+2, '?');
			if (pszHeaderEnd != NULL && pszHeaderEnd[2] == '?' && pszHeaderEnd+3 < pbEnd)
			{
				nCoding = tolower(pszHeaderEnd[1]);
				pszHeaderEnd += 3;
				pszCodeEnd = ::strstr(pszHeaderEnd, "?=");	// look for the tailer
				if (!pszCodeEnd || pszCodeEnd >= pbEnd)
					pszCodeEnd = pbEnd;
				nCodeLen = (int)(pszCodeEnd - pszHeaderEnd);
				pszCodeEnd += 2;
				if (m_strCharset.empty())
				{
					m_strCharset.assign(pbData+2, pszHeaderEnd-pbData-5);
					m_nEncoding = nCoding;
				}
			}
		}

		int nDecoded;
		if (nCoding == 'b')
		{
			CMimeCodeBase64 base64;
			base64.SetInput(pszHeaderEnd, nCodeLen, false);
			nDecoded = base64.GetOutput(pbOutput, nMaxSize);
		}
		else if (nCoding == 'q')
		{
			CMimeCodeQP qp;
			qp.SetInput(pszHeaderEnd, nCodeLen, false);
			nDecoded = qp.GetOutput(pbOutput, nMaxSize);
		}
		else
		{
			pszCodeEnd = ::strstr(pbData+1, "=?");	// find the next encoded-word
			if (!pszCodeEnd || pszCodeEnd >= pbEnd)
				pszCodeEnd = pbEnd;
			else if (pbData > (const char*) m_pbInput)
			{
				const char* pszSpace = pbData;
				while (CMimeChar::IsSpace((unsigned char)*pszSpace))
					pszSpace++;
				if (pszSpace == pszCodeEnd)	// ignore liner-white-spaces between adjacent encoded words
					pbData = pszCodeEnd;
			}
			nDecoded = min((int)(pszCodeEnd - pbData), nMaxSize);
			::memcpy(pbOutput, pbData, nDecoded);
		}

		pbData = pszCodeEnd;
		pbOutput += nDecoded;
		nMaxSize -= nDecoded;
		if (nMaxSize <= 0)
			break;
	}

	return (int)(pbOutput - pbOutStart);
}

int CMimeEncodedWord::BEncode(unsigned char* pbOutput, int nMaxSize) const
{
	int nCharsetLen = (int)m_strCharset.size();
	int nBlockSize = MAX_ENCODEDWORD_LEN - nCharsetLen - 7;	// a single encoded-word cannot exceed 75 bytes
	nBlockSize = nBlockSize / 4 * 3;
	ASSERT(nBlockSize > 0);

	unsigned char* pbOutStart = pbOutput;
	int nInput = 0;
	for (;;)
	{
		if (nMaxSize < nCharsetLen+7)
			break;
		*pbOutput++ = '=';			// encoded-word header
		*pbOutput++ = '?';
		::memcpy(pbOutput, m_strCharset.c_str(), nCharsetLen);
		pbOutput += nCharsetLen;
		*pbOutput++ = '?';
		*pbOutput++ = 'B';
		*pbOutput++ = '?';

		nMaxSize -= nCharsetLen + 7;
		CMimeCodeBase64 base64;
		base64.SetInput((const char*)m_pbInput+nInput, min(m_nInputSize-nInput, nBlockSize), true);
		base64.AddLineBreak(false);
		int nEncoded = base64.GetOutput(pbOutput, nMaxSize);
		pbOutput += nEncoded;
		*pbOutput++ = '?';			// encoded-word tailer
		*pbOutput++ = '=';

		nInput += nBlockSize;
		nMaxSize -= nEncoded + nCharsetLen + 7;
		if (nInput >= m_nInputSize)
			break;
		*pbOutput++ = ' ';			// add a liner-white-space between adjacent encoded words
		nMaxSize--;
	}
	return (int)(pbOutput - pbOutStart);
}

int CMimeEncodedWord::QEncode(unsigned char* pbOutput, int nMaxSize) const
{
	static const char* s_QPTable = "0123456789ABCDEF";

	const unsigned char* pbData = m_pbInput;
	const unsigned char* pbEnd = m_pbInput + m_nInputSize;
	unsigned char* pbOutStart = pbOutput;
	unsigned char* pbOutEnd = pbOutput + nMaxSize;
	int nCodeLen, nCharsetLen = (int)m_strCharset.size();
	int nLineLen = 0, nMaxLine = MAX_ENCODEDWORD_LEN - nCharsetLen - 7;

	while (pbData < pbEnd)
	{
		unsigned char ch = *pbData++;
		if (ch < 33 || ch > 126 || ch == '=' || ch == '?' || ch == '_')
			nCodeLen = 3;
		else
			nCodeLen = 1;

		if (nLineLen+nCodeLen > nMaxLine)	// add encoded word tailer
		{
			if (pbOutput+3 > pbOutEnd)
				break;
			*pbOutput++ = '?';
			*pbOutput++ = '=';
			*pbOutput++ = ' ';
			nLineLen = 0;
		}

		if (!nLineLen)				// add encoded word header
		{
			if (pbOutput+nCharsetLen+7 > pbOutEnd)
				break;
			*pbOutput++ = '=';
			*pbOutput++ = '?';
			::memcpy(pbOutput, m_strCharset.c_str(), nCharsetLen);
			pbOutput += nCharsetLen;
			*pbOutput++ = '?';
			*pbOutput++ = 'Q';
			*pbOutput++ = '?';
		}

		nLineLen += nCodeLen;
		if (pbOutput+nCodeLen > pbOutEnd)
			break;
		if (nCodeLen > 1)
		{
			*pbOutput++ = '=';
			*pbOutput++ = s_QPTable[(ch >> 4) & 0x0f];
			*pbOutput++ = s_QPTable[ch & 0x0f];
		}
		else
			*pbOutput++ = ch;
	}

	if (pbOutput+2 <= pbOutEnd)
	{
		*pbOutput++ = '?';
		*pbOutput++ = '=';
	}
	return (int)(pbOutput - pbOutStart);
}

//////////////////////////////////////////////////////////////////////
// CFieldCodeBase - base class to encode/decode header fields
// default coder for any unregistered fields
//////////////////////////////////////////////////////////////////////
int CFieldCodeBase::GetEncodeLength() const
{
	// use the global charset if there's no specified charset
	string strCharset = m_strCharset;
	if (strCharset.empty())
		strCharset = CMimeEnvironment::GetGlobalCharset();
	if (strCharset.empty() && !CMimeEnvironment::AutoFolding())
		return CMimeCodeBase::GetEncodeLength();

	int nLength = 0;
	const char* pszData = (const char*) m_pbInput;
	int nInputSize = m_nInputSize;
	int nNonAsciiChars, nDelimeter = GetDelimeter();

	// divide the field into syntactic units to calculate the output length
	do
	{
		int nUnitSize = FindSymbol(pszData, nInputSize, nDelimeter, nNonAsciiChars);
		if (!nNonAsciiChars || strCharset.empty())
			nLength += nUnitSize;
		else
		{
			CMimeEncodedWord coder;
			coder.SetEncoding(SelectEncoding(nUnitSize, nNonAsciiChars), strCharset.c_str());
			coder.SetInput(pszData, nUnitSize, true);
			nLength += coder.GetOutputLength();
		}

		pszData += nUnitSize;
		nInputSize -= nUnitSize;
		if (IsFoldingChar(*pszData))	// the char follows the unit is a delimeter (space or special char)
			nLength += 3;
		nLength++;
		pszData++;
		nInputSize--;
	} while (nInputSize > 0);

	if (CMimeEnvironment::AutoFolding())
		nLength += nLength / MAX_MIME_LINE_LEN * 6;
	return nLength;
}

int CFieldCodeBase::Encode(unsigned char* pbOutput, int nMaxSize) const
{
	// use the global charset if there's no specified charset
	string strCharset = m_strCharset;
	if (strCharset.empty())
		strCharset = CMimeEnvironment::GetGlobalCharset();
	if (strCharset.empty() && !CMimeEnvironment::AutoFolding())
		return CMimeCodeBase::Encode(pbOutput, nMaxSize);

	unsigned char* pbOutBegin = pbOutput;
	unsigned char* pbOutEnd = pbOutput + nMaxSize;
	const char* pszInput = (const char*) m_pbInput;
	int nInputSize = m_nInputSize;
	int nNonAsciiChars, nDelimeter = GetDelimeter();
	int nLineLen = 0;
	unsigned char* pbSpace = NULL;
	string strUnit;
	strUnit.reserve(nInputSize);

	// divide the field into syntactic units to encode
	for (;;)
	{
		int nUnitSize = FindSymbol(pszInput, nInputSize, nDelimeter, nNonAsciiChars);
		if (!nNonAsciiChars || strCharset.empty())
			strUnit.assign(pszInput, nUnitSize);
		else
		{
			CMimeEncodedWord coder;
			coder.SetEncoding(SelectEncoding(nUnitSize, nNonAsciiChars), strCharset.c_str());
			coder.SetInput(pszInput, nUnitSize, true);
			strUnit.resize(coder.GetOutputLength());
			int nEncoded = coder.GetOutput((unsigned char*) strUnit.c_str(), (int) strUnit.capacity());
			strUnit.resize(nEncoded);
		}
		if (nUnitSize < nInputSize)
			strUnit += pszInput[nUnitSize];		// add the following delimeter (space or special char)

		// copy the encoded string to target buffer and perform folding if needed
		if (!CMimeEnvironment::AutoFolding())
		{
			int nSize = min((int) (pbOutEnd - pbOutput), (int) strUnit.size());
			::memcpy(pbOutput, strUnit.c_str(), nSize);
			pbOutput += nSize;
		}
		else
		{
			const char* pszData = strUnit.c_str();
			const char* pszEnd = pszData + strUnit.size();
			while (pszData < pszEnd)
			{
				char ch = *pszData++;
				if (ch == '\r' || ch == '\n')
				{
					nLineLen = -1;
					pbSpace = NULL;
				}
				else if (nLineLen > 0 && CMimeChar::IsSpace(ch))
					pbSpace = pbOutput;

				if (nLineLen >= MAX_MIME_LINE_LEN && pbSpace != NULL &&
					pbOutput+3 <= pbOutEnd)		// fold at the position of the previous space
				{
					int nSize = (int)(pbOutput - pbSpace);
					::memmove(pbSpace+3, pbSpace, nSize);
					::memcpy(pbSpace, "\r\n\t", 3);
					pbOutput += 3;
					pbSpace = NULL;
					nLineLen = nSize + 1;
				}
				if (pbOutput < pbOutEnd)
					*pbOutput++ = ch;
				nLineLen++;
			}
		}

		pszInput += nUnitSize + 1;
		nInputSize -= nUnitSize + 1;
		if (nInputSize <= 0)
			break;

		// fold at the position of the specific char and eat the following spaces
		if (IsFoldingChar(pszInput[-1]) && pbOutput+3 <= pbOutEnd)
		{
			::memcpy(pbOutput, "\r\n\t", 3);
			pbOutput += 3;
			pbSpace = NULL;
			nLineLen = 1;
			while (nInputSize > 0 && CMimeChar::IsSpace(*pszInput))
			{
				pszInput++;
				nInputSize--;
			}
		}
	}
	return (int) (pbOutput - pbOutBegin);
}

int CFieldCodeBase::Decode(unsigned char* pbOutput, int nMaxSize)
{
	CMimeEncodedWord coder;
	coder.SetInput((const char*)m_pbInput, m_nInputSize, false);

	string strField;
	strField.resize(coder.GetOutputLength());
	int nDecoded = coder.GetOutput((unsigned char*) strField.c_str(), (int) strField.capacity());
	strField.resize(nDecoded);
	m_strCharset = coder.GetCharset();

	if (CMimeEnvironment::AutoFolding())
		UnfoldField(strField);
	int nSize = min((int)strField.size(), nMaxSize);
	::memcpy(pbOutput, strField.c_str(), nSize);
	return nSize;
}

void CFieldCodeBase::UnfoldField(string& strField) const
{
	for (;;)
	{
		string::size_type pos = strField.rfind("\r\n");
		if (pos == string::npos)
			break;

		strField.erase(pos, 2);
		//if (strField[pos] == '\t')
		//	strField[pos] = ' ';
		int nSpaces = 0;
		while (CMimeChar::IsSpace((unsigned char)strField[pos+nSpaces]))
			nSpaces++;
		strField.replace(pos, nSpaces, " ");
	}
}

int CFieldCodeBase::FindSymbol(const char* pszData, int nSize, int& nDelimeter, int& nNonAscChars) const
{
	nNonAscChars = 0;
	const char* pszDataStart = pszData;
	const char* pszEnd = pszData + nSize;

	while (pszData < pszEnd)
	{
		char ch = *pszData;
		if (CMimeChar::IsNonAscii((unsigned char)ch))
			nNonAscChars++;
		else
		{
			if (ch == (char) nDelimeter)
			{
				nDelimeter = 0;		// stop at any delimeters (space or specials)
				break;
			}

			if (!nDelimeter && CMimeChar::IsDelimiter(ch))
			{
				switch (ch)
				{
				case '"':
					nDelimeter = '"';	// quoted-string, delimeter is '"'
					break;
				case '(':
					nDelimeter = ')';	// comment, delimeter is ')'
					break;
				case '<':
					nDelimeter = '>';	// address, delimeter is '>'
					break;
				}
				break;
			}
		}
		pszData++;
	}

	return (int)(pszData - pszDataStart);
}
