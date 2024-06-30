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


#include "stdafx.h"
#include "TextUtilsEx.h"
#include "SimpleString.h"
#include <algorithm>
#include "FileUtils.h"
#include "MimeCode.h"

void TextUtilsEx::ReplaceNL2CRNL(const char *in, int inLength, SimpleString *out)
{
	// Assume out is done by caller
	out->ClearAndResize(inLength * 2);
	register char *p = (char*)in;
	register char *e = p + inLength;
	char *p_beg = p;
	int len;

	p_beg = p;
	while (p < e)
	{
		if (*p == '\n')
		{
			if (p == in)
			{
				out->Append("\r\n", 2);
				p++;  // jump over NL
				p_beg = p;
			}
			else if (*(p - 1) != '\r')
			{
				len = IntPtr2Int(p - p_beg);
				if (len > 0)
					out->Append(p_beg, len);

				out->Append("\r\n", 2);
				p++;  // jump over NL
				p_beg = p;
			}
			else // found CR LF; keep going
				p++;
		}
		else
			p++;
	}
	len = IntPtr2Int(p - p_beg);
	if (len > 0)
	{
		out->Append(p_beg, len);
		out->Append("\r\n", 2);
	}
}

void TextUtilsEx::EncodeAsHtml(const char *in, int inLength, SimpleString *out)
{
	// Assume out is done by caller
	out->ClearAndResize(inLength * 2);
	register char *p = (char*)in;
	register char *e = p + inLength;
	char *p_beg;
	int len;
	char c;

	p_beg = p;
	out->Append("<p>");
	while (p < e)
	{
		c = *p;
#if 1
		if (c == ' ')
		{
			if (*(p+1) == ' ')
			{
				len = IntPtr2Int(p - p_beg);
				if (len > 0) out->Append(p_beg, len);
				out->Append("&nbsp;");
				p++;
				p_beg = p;
			}
			else
				p++;
		}
		else
#endif
		if ((c == '>') || (c == '<') || (c == '\"') || (c == '&') || (c == '\''))
		{
			len = IntPtr2Int(p - p_beg);
			if (len > 0) out->Append(p_beg, len);
			switch (c) {
			case '&':  out->Append("&amp;"); break;
			case '\"': out->Append("&quot;"); break;
			case '\'': out->Append("&apos;"); break;
			case '<':  out->Append("&lt;"); break;
			case '>':  out->Append("&gt;"); break;
			default:   break; // we should never be here
			}
			p++;
			p_beg = p;
		}
		else if (*p == '\r')
		{
			len = IntPtr2Int(p - p_beg);
			if (len > 0)  out->Append(p_beg, len);

			p++;  // jump over '\r'
			if (p < e)
			{
				if (*p == '\n') {
					out->Append("<br>\r\n", 6);
					p++;
				}
				else
					out->Append("<br>\r", 5);

			}
			else
				out->Append("<br>\r", 5);

			p_beg = p;
		}
		else if (*p == '\n')
		{
			len = IntPtr2Int(p - p_beg);
			if (len > 0)  out->Append(p_beg, len);
			out->Append("<br>\n", 5);
			p++;
			p_beg = p;
		}
		else
			p++;
	}
	len = IntPtr2Int(p - p_beg);
	if (len > 0)
	{
		out->Append(p_beg, len);
	}
	out->Append("</p>");
}

void TextUtilsEx::EncodeAsHtmlText(const char *in, int inLength, SimpleString *out)
{
	static char *prefixTable[] = { "http:", "https:", "ftp:", "mailto:", "file:", };
	static int prefixLengthTable[] = { istrlen("http:"), istrlen("https:"), istrlen("ftp:"), istrlen("mailto:"), istrlen("file:") };
	static char *pImgTag = "img";
	static int imgTagLen = istrlen(pImgTag);

	// Assume out is done by caller
	out->ClearAndResize(inLength * 3);
	register char *p = (char*)in;
	register char *e = p + inLength;
	char *p_beg;
	int len;
	char c;

	int prefixCnt = sizeof(prefixTable)/sizeof(char*);

	char *p_save = p;
	p_beg = p;  // points to beginning of text to copy
	char *prefixBeg;
	out->Append("<pre>");
	while (p < e)
	{
		c = *p;
		p_save = p;
		if (c == '<')
		{
			char* pBeg = p + 1;
			char* pEnd = pBeg + 3;
			int retlen = TextUtilsEx::strncmpUpper2Lower(pBeg, pEnd, pImgTag, imgTagLen);
			if (retlen == 0)
			{
				p = p + 4;
				while (p < e)
				{
					if (*p == '\r')
						break;
					else
						p++;
				}
				p++;  // skip CR
				_ASSERTE(*p == '\n');
				p++;  // skip NL
				int len = IntPtr2Int(p - p_beg);
				out->Append(p_beg, len);
				p_beg = p;
				continue;
			}
		}
		else if (c == ':')
		{
			int i = 0;
			char *pEnd = p+1;
			for (i; i < prefixCnt; i++)
			{
				char *cPrefix = prefixTable[i];
				int cPrefixLen = prefixLengthTable[i];

				prefixBeg = p - (cPrefixLen -1);
				if (prefixBeg < in)
					continue;

				int retlen = TextUtilsEx::strncmpUpper2Lower(prefixBeg, pEnd, cPrefix, cPrefixLen);
				if (retlen == 0)
				{
					len = IntPtr2Int(prefixBeg - p_beg);
					out->Append(p_beg, len);
					out->Append("<a href=\"");

					p_beg = prefixBeg;
					// find end of link
					c = 0;
					while (p < e)
					{
						c = *p;
						if ((c == ' ') || (c == '\t') || (c == '\r') || (c == '\n') )
							break;
						else
							p++;
					}
					if (c)
					{
						c = *(p-1);
						if ((c == '>') || (c == ')') || (c == '.') || (c == ',') || (c == ']') || (c == '}'))  // inexpensive but not most reliable solution
							p--;
					}
					len = IntPtr2Int(p - p_beg);
					out->Append(p_beg, len);
					out->Append("\">");
					out->Append(p_beg, len);
					out->Append("</a>");
					p_beg = p;
					break;
				}
			}
			if (p == p_save)
			{
				p++;
			}
			continue;
		}
		if ((c == '>') || (c == '<') || (c == '\"') || (c == '&') || (c == '\''))
		{
			len = IntPtr2Int(p - p_beg);
			if (len > 0) out->Append(p_beg, len);
			switch (c) {
			case '&':  out->Append("&amp;"); break;
			case '\"': out->Append("&quot;"); break;
			case '\'': out->Append("&apos;"); break;
			case '<':  out->Append("&lt;"); break;
			case '>':  out->Append("&gt;"); break;
			default:   break; // we should never be here
			}
			p++;
			p_beg = p;
		}
		else
			p++;
	}
	len = IntPtr2Int(p - p_beg);
	if (len > 0)
	{
		out->Append(p_beg, len);
	}
	out->Append("</pre>");
}

BOOL TextUtilsEx::WStr2Ansi(CString &strW, CStringA &strA, DWORD &error)
{
	UINT outCodePage = CP_ACP;
	BOOL ret = WStr2CodePage((wchar_t*)((LPCWSTR)strW), strW.GetLength(), outCodePage, &strA, error);
	return ret;
}

BOOL TextUtilsEx::WStr2Ascii(CString& strW, CStringA& strA, DWORD& error)
{
	UINT CP_US_ASCII = 20127;
	UINT outCodePage = CP_US_ASCII;
	BOOL ret = WStr2CodePage((wchar_t*)((LPCWSTR)strW), strW.GetLength(), outCodePage, &strA, error);
	return ret;
}

BOOL TextUtilsEx::Ansi2WStr(CStringA &strA, CString &strW, DWORD &error, DWORD dwFlags)
{
	UINT strCodePage = CP_ACP;
	BOOL ret = CodePage2WStr(&strA, strCodePage, &strW, error, dwFlags);
	return ret;
}

BOOL TextUtilsEx::Ascii2Wstr(CStringA &strA, CString &strW, DWORD &error, DWORD dwFlags)
{
	UINT CP_US_ASCII = 20127;
	UINT strCodePage = CP_US_ASCII;
	BOOL ret = CodePage2WStr(&strA, strCodePage, &strW, error, dwFlags);
	return ret;
}

UINT TextUtilsEx::StrPageCodeName2PageCode(const  char* PageCodeStr)
{
	UINT CodePage = 0;

	CodePage = TextUtilsEx::charset2Id(PageCodeStr);

	return CodePage;
}

DWORD TextUtilsEx::Str2Ansi(CStringA& str, UINT strCodePage, DWORD& error)
{
	UINT toPageCode = CP_ACP;

	SimpleString result;  // Str2CodePage will resize
	SimpleString workBuff; // Str2CodePage will resize
	BOOL retA2A = TextUtilsEx::Str2CodePage((LPCSTR)str, str.GetLength(), strCodePage, toPageCode, &result, &workBuff, error);
	if (retA2A == TRUE)
	{
		char *buff = str.GetBufferSetLength(result.Count());
		memcpy(buff, result.Data(), result.Count());
		str.ReleaseBuffer();
	}
	return retA2A;
}


BOOL TextUtilsEx::Str2UTF8(const char* instr, int instrlen, UINT inCodePage, CStringA& outstr, DWORD& error)
{
	SimpleString workBuff;
	SimpleString outBuff;
	BOOL retStr2CP = TextUtilsEx::Str2CodePage(instr, instrlen, inCodePage, CP_UTF8,
		&outBuff, &workBuff, error);

	outstr.Empty();
	if (retStr2CP)
	{
		outstr.Append(outBuff.Data(), outBuff.Count());
	}
	else  // return original string. Sould work better than empty string  // FIXME
	{
		outstr.Append(instr, instrlen);
	}
	return retStr2CP;
}

BOOL TextUtilsEx::Str2CodePage(const char *str, int strlen, UINT inCodePage, UINT outCodePage, SimpleString *result, SimpleString *workBuff, DWORD& error)
{
	DWORD dwFlags = 0;
	BOOL retA2W = TextUtilsEx::CodePage2WStr(str, strlen, inCodePage, workBuff, error, dwFlags);

	const wchar_t* wbuff = (wchar_t*)workBuff->Data();
	int wlen = workBuff->Count() / 2;

	BOOL  retW2A = FALSE;
	if (retA2W)
	{
		retW2A = TextUtilsEx::WStr2CodePage(wbuff, wlen, outCodePage, result, error);
	}
	return retW2A;
}

BOOL TextUtilsEx::Str2UTF8(const char* str, int strlen, UINT inCodePage, SimpleString* result, SimpleString* workBuff, DWORD& error)
{
	UINT outCodePage = CP_UTF8;
	return TextUtilsEx::Str2CodePage(str, strlen, inCodePage, outCodePage, result, workBuff, error);
}

BOOL TextUtilsEx::Str2CodePage(SimpleString *str, UINT inCodePage, UINT outCodePage, SimpleString *result, SimpleString *workBuff, DWORD& error)
{
	BOOL retA2A = TextUtilsEx::Str2CodePage(str->Data(), str->Count(), inCodePage, outCodePage, result, workBuff, error);
	return retA2A;
}

BOOL TextUtilsEx::WStr2CodePage(const wchar_t *wbuff, int wbufflen, UINT outCodePage, CStringA *astr, DWORD &error)
{
	error = 0;
	if (wbufflen == 0)
	{
		astr->Empty();
		return TRUE;
	}

	UINT CP_ASCII = 20127;

	// worst case scenario. Estimated length
	int sbufflen = wbufflen * 4 + 2;  // This likely overkill or call WideCharToMultiByte first to get the required length
	sbufflen = ((sbufflen / 128) + 1) * 128;  // may help to reduce memory fragmentation ?

	if (outCodePage == CP_ASCII)
	{
		//_ASSERTE(FALSE);
		outCodePage = CP_UTF8;  // FIXME ?? 
	}

	char* sbuff = astr->GetBuffer(sbufflen);
	int slen = WideCharToMultiByte(outCodePage, 0, wbuff, wbufflen, sbuff, sbufflen-1, NULL, NULL);
	_ASSERTE(slen != 0);
	if (slen == 0)
	{
		error = ::GetLastError();
		CString errorText = FileUtils::GetLastErrorAsString();

		// ERROR_INSUFFICIENT_BUFFER.A supplied buffer size was not large enough, or it was incorrectly set to NULL.
		// ERROR_INVALID_FLAGS.The values supplied for flags were not valid.
		// ERROR_INVALID_PARAMETER. Any of the parameter values was invalid.
		// ERROR_NO_UNICODE_TRANSLATION.Invalid Unicode was found in a string.

		astr->ReleaseBuffer(0);
		return FALSE;
	}
	astr->ReleaseBuffer(slen);  // TODO: check if extra space is freed
	return TRUE;
}

BOOL TextUtilsEx::WStr2CodePage(const wchar_t *wbuff, int wlen, UINT outCodePage, SimpleString *result, DWORD &error)
{
	error = 0;
	if (wlen == 0)
	{
		result->ClearAndResize(1);
		return TRUE;
	}

	UINT CP_ASCII = 20127;

	int outLen = wlen * 4 + 2;  // This likely overkill or call WideCharToMultiByte first to get the required length
	outLen = ((outLen / 128) + 1) * 128;  // may help to reduce memory fragmentation ?

	if (outCodePage == CP_ASCII)
	{
		outCodePage = CP_UTF8;  // FIXME ?? 
	}

	result->ClearAndResize(outLen); 
	int cpLen = WideCharToMultiByte(outCodePage, 0, wbuff, wlen, result->Data(), outLen-1, NULL, NULL);
	_ASSERTE(cpLen != 0);
	if (cpLen == 0)
	{
		error = ::GetLastError();
		CString errorText = FileUtils::GetLastErrorAsString();
		// ERROR_INSUFFICIENT_BUFFER.A supplied buffer size was not large enough, or it was incorrectly set to NULL.
		// ERROR_INVALID_FLAGS.The values supplied for flags were not valid.
		// ERROR_INVALID_PARAMETER. Any of the parameter values was invalid.
		// ERROR_NO_UNICODE_TRANSLATION.Invalid Unicode was found in a string.

		result->Clear();
		return FALSE;
	}
	result->SetCount(cpLen);
	return TRUE;
}

BOOL TextUtilsEx::WStr2UTF8(const wchar_t* wbuff, int wlen, SimpleString* result, DWORD& error)
{
	UINT outCodePage = CP_UTF8;
	return TextUtilsEx::WStr2CodePage(wbuff, wlen, outCodePage, result, error);
}

BOOL TextUtilsEx::CodePage2WStr(SimpleString *str, UINT strCodePage, SimpleString *wstr, DWORD &error, DWORD dwFlags)
{
	BOOL ret = CodePage2WStr(str->Data(), str->Count(), strCodePage, wstr, error, dwFlags);
	return ret;
}
//
BOOL TextUtilsEx::CodePage2WStr(CStringA *str, UINT strCodePage, SimpleString *wstr, DWORD &error, DWORD dwFlags)
{
	BOOL ret = CodePage2WStr((char*)(str->operator LPCSTR()), str->GetLength(), strCodePage, wstr, error, dwFlags);
	return ret;
}

BOOL TextUtilsEx::CodePage2WStr(const char *str, int slen, UINT strCodePage, SimpleString *wstr, DWORD &error, DWORD dwFlags)
{
	/*
	int MultiByteToWideChar(
		[in]            UINT                              CodePage,
		[in]            DWORD                             dwFlags,
		[in]            _In_NLS_string_(cbMultiByte)LPCCH lpMultiByteStr,
		[in]            int                               cbMultiByte,
		[out, optional] LPWSTR                            lpWideCharStr,
		[in]            int                               cchWideChar
	);

	dwFlags:
	0 - output is in KC ??
	MB_COMPOSITE  - output is in NormalizationD 
	MB_ERR_INVALID_CHARS, 
	MB_PRECOMPOSED - output is in NormalizationC 
	MB_USEGLYPHCHARS

	*/
	error = 0;
	if (slen == 0)
	{
		wstr->ClearAndResize(2);
		wchar_t* wbuff = (wchar_t*)wstr->Data();
		wbuff[0] = L'\0';  // will set two 0 trailing bytes
		return TRUE;
	}

	// worst case scenario. Estimated length
	int buffLen = slen * 4 + 2;  // This is likely overkill or call MultiByteToWideChar first to get the required length
	buffLen = ((buffLen / 128) + 1) * 128;  // may help to reduce memory fragmentation ??

	wstr->ClearAndResize(buffLen);
	wchar_t* wbuff = (wchar_t*)wstr->Data();
	int wbuffLen = buffLen/2 - 1;

	int wlen = MultiByteToWideChar(strCodePage, 0, str, slen, wbuff, wbuffLen);
	_ASSERTE(wlen != 0);
	if (wlen == 0)
	{
		error = ::GetLastError();
		CString errorText = FileUtils::GetLastErrorAsString();
		// ERROR_INSUFFICIENT_BUFFER.A supplied buffer size was not large enough, or it was incorrectly set to NULL.
		// ERROR_INVALID_FLAGS.The values supplied for flags were not valid.
		// ERROR_INVALID_PARAMETER. Any of the parameter values was invalid.
		// ERROR_NO_UNICODE_TRANSLATION.Invalid Unicode was found in a string.

		wstr->Clear();
		return FALSE;
	}
	wstr->SetCount(wlen * 2);  // sets one 0 trailing byte
	wbuff[wlen] = L'\0';  // will set two 0 trailing bytes

	return TRUE;
}

BOOL TextUtilsEx::WStr2UTF8(CString* str, CStringA* outstr, DWORD& error)
{
	UINT outCodePage = CP_UTF8;
	BOOL ret = WStr2CodePage(str->operator LPCWSTR(), str->GetLength(), outCodePage, outstr, error);
	return ret;
}


BOOL TextUtilsEx::UTF82WStr(CStringA* str, CString* wstr, DWORD& error, DWORD dwFlags)
{
	BOOL ret = CodePage2WStr((char*)(str->operator LPCSTR()), str->GetLength(), CP_UTF8, wstr, error, dwFlags);
	return ret;
}

BOOL TextUtilsEx::CodePage2WStr(CStringA *str, UINT strCodePage, CString *wstr, DWORD &error, DWORD dwFlags)
{
	BOOL ret = CodePage2WStr((char*)(str->operator LPCSTR()), str->GetLength(), strCodePage, wstr, error, dwFlags);
	return ret;
}

BOOL TextUtilsEx::CodePage2WStr(const char *str, int slen, UINT strCodePage, CString *wstr, DWORD &error, DWORD dwFlags)
{
	/*
	int MultiByteToWideChar(
		[in]            UINT                              CodePage,
		[in]            DWORD                             dwFlags,
		[in]            _In_NLS_string_(cbMultiByte)LPCCH lpMultiByteStr,
		[in]            int                               cbMultiByte,
		[out, optional] LPWSTR                            lpWideCharStr,
		[in]            int                               cchWideChar
	);

	dwFlags:
	0 - output is in KC ??
	MB_COMPOSITE  - output is in NormalizationD 
	MB_ERR_INVALID_CHARS, 
	MB_PRECOMPOSED - output is in NormalizationC 
	MB_USEGLYPHCHARS

	*/
	error = 0;
	wstr->Empty();
	if (slen == 0)
		return TRUE;

	// worst case scenario. Estimated length
	int wbuffLen = slen * 4 + 2;  // This is likely overkill or call MultiByteToWideChar first to get the required length
	wbuffLen = ((wbuffLen / 128) + 1) * 128;  // may help to reduce memory fragmentation ??
	wbuffLen = wbuffLen/2 -2;

	wchar_t* wbuff = wstr->GetBuffer(wbuffLen + 1);
	int wlen = MultiByteToWideChar(strCodePage, 0, str, slen, wbuff, wbuffLen);
	_ASSERTE(wlen != 0);
	if (wlen == 0)
	{
		error = ::GetLastError();
		CString errorText = FileUtils::GetLastErrorAsString();
		// ERROR_INSUFFICIENT_BUFFER.A supplied buffer size was not large enough, or it was incorrectly set to NULL.
		// ERROR_INVALID_FLAGS.The values supplied for flags were not valid.
		// ERROR_INVALID_PARAMETER. Any of the parameter values was invalid.
		// ERROR_NO_UNICODE_TRANSLATION.Invalid Unicode was found in a string.

		wstr->ReleaseBuffer(0);
		return FALSE;
	}
	wstr->ReleaseBuffer(wlen);
	return TRUE;
}

BOOL TextUtilsEx::Str2UTF8(SimpleString *str, UINT strCodePage, SimpleString *result, SimpleString *workBuff, DWORD& error)
{
	BOOL ret = Str2CodePage(str, strCodePage, CP_UTF8, result, workBuff, error);
	return ret;
}

BOOL TextUtilsEx::Str2Ansi(SimpleString *str, UINT strCodePage, SimpleString *result, SimpleString *workBuff, DWORD& error)
{
	UINT currentCodePage = GetACP();
	BOOL ret = Str2CodePage(str, strCodePage, currentCodePage, result, workBuff, error);
	return ret;
}

#include "MimeCode.h"

// TODO: toCharacterId is ignored for now
CStringA TextUtilsEx::DecodeString(CStringA &subj, CStringA &charset, UINT &charsetId, UINT toCharacterId, DWORD& error)
{
	CFieldCodeText tfc;
	//tfc.SetInput(subj.GetBuffer(), subj.GetLength(), false);  // FIXME
	tfc.SetInput((LPCSTR)subj, subj.GetLength(), false);
	int outputLen = tfc.GetOutputLength();
	if (outputLen > 0) 
	{
		int maxOutputLen = 2 * outputLen + 2;
		unsigned char *outBuf = (unsigned char*)malloc(maxOutputLen);
		int decodeLen = tfc.GetOutput(outBuf, maxOutputLen);
		outBuf[decodeLen] = 0;
		CStringA str(outBuf);
		free(outBuf);
		charset = tfc.GetCharset();
		UINT CodePage = TextUtilsEx::StrPageCodeName2PageCode(tfc.GetCharset());
		charsetId = CodePage;
		// No remapping for now. Header of message windows should show text properly
		return str;

		if ((CodePage > 0) && (toCharacterId > 0)) {
			TextUtilsEx::Str2Ansi(str, CodePage, error);
			return str;
		}
		else
			return subj;
	}
	else
		return subj;
}

int TextUtilsEx::DecodeString(CStringA& inStr, CStringA &outStr, CStringA& charset, UINT& charsetId, DWORD& error)
{
	CFieldCodeText tfc;

	tfc.SetInput((LPCSTR)inStr, inStr.GetLength(), false);
	int outputLen = tfc.GetOutputLength();
	if (outputLen > 0)
	{
		int maxOutputLen = 2 * outputLen + 2;

		outStr.GetBuffer(maxOutputLen);
		int decodeLen = tfc.GetOutput((unsigned char*)((LPCSTR)outStr), maxOutputLen);

		if (decodeLen >= 0)
		{
			outStr.ReleaseBuffer(decodeLen);

			charset = tfc.GetCharset();
			UINT CodePage = TextUtilsEx::StrPageCodeName2PageCode(tfc.GetCharset());
			charsetId = CodePage;

			return decodeLen;
		}
		else
		{
			outStr.ReleaseBuffer(0);
			return 0;
		}
	}
	else
		return 0;
}

int TextUtilsEx::EncodeString(CStringA& inStr,  CStringA &outStr, UINT charsetId, DWORD& error)
{
	CFieldCodeText tfc;

	std::string charset;
	BOOL ret = id2charset(charsetId, charset);

	tfc.SetInput((LPCSTR)inStr, inStr.GetLength(), true);
	tfc.SetCharset(charset.c_str());
	int outputLen = tfc.GetOutputLength();

	if (outputLen > 0)
	{
		int maxOutputLen = 2 * outputLen + 2;

		outStr.GetBuffer(maxOutputLen);
		int decodeLen = tfc.GetOutput((unsigned char*)((LPCSTR)outStr), maxOutputLen);
		if (decodeLen >= 0)
		{
			outStr.ReleaseBuffer(decodeLen);
			return decodeLen;
		}
		else
		{
			outStr.ReleaseBuffer(0);
			return 0;
		}
	}
	else
		return 0;
}

int  TextUtilsEx::hextob(char ch)
{
	if (ch >= '0' && ch <= '9') return ch - '0';
	if (ch >= 'A' && ch <= 'F') return ch - 'A' + 10;
	if (ch >= 'a' && ch <= 'f') return ch - 'a' + 10;
	return -1;
}

int TextUtilsEx::DecodeMimeChunkedString(CStringA &inString, CStringA &charset, UINT &charsetId,  BOOL hasCharset, CStringA &outString)
{
	const char *p_beg = (LPCSTR)inString;
	const char *p = p_beg;
	int length = inString.GetLength();
	const char *e = p + length;

	// example string UTF-8''Copy%20of%20JOHN%20KOWALSKI%20SECOND%20PRIORITIES%20Rev%201%20with%20effort%20cost.xlsx
	// value looks like URL, may contatin %20 for example
	// can value look like URL but without charset ?? try to verify

	charsetId = 0;
	if (hasCharset)
	{
		// find charset
		while ((p < e) && (*p != '\''))
		{
			p++;
		}
		if (p >= e)
		{
			//  TODO: can't decode; return untoched string for now
			return 0;
		}

		charset.Append(p_beg, IntPtr2Int(p - p_beg));
		charsetId = TextUtilsEx::StrPageCodeName2PageCode(charset);

		// skip first '\''
		p++;

		// find language
		while ((p < e) && (*p != '\''))
		{
			p++;
		}

		if (p >= e)
		{
			//  TODO: can't decode; return untoched string for now
			return 0;
		}

		p++;
	}

	const char* p_val_beg = p;

	{
		// decode

		SimpleString out(length * 2);

		int c;
		while (p < e)
		{
			c = *p;
			if (c == '%')
			{
				p++;
				c = *p++;
				int first = hextob(c);
				if (first < 0)
					break;
				c = *p++;
				int second = hextob(c);
				if (second < 0)
					break;
				unsigned char uc = first * 16 + second;
				out.Append(uc);
			}
			else
			{
				out.Append(c);
				p++;
			}
		}
		if (p < e)  // not a URL
		{
			outString.Empty();
			int vlen = IntPtr2Int(e - p_val_beg);
			outString.Append(p_val_beg, vlen);
		}
		else
		{
			outString.Empty();
			outString.Append(out.Data(), out.Count());
		}
	}

	return 1;
}


CP2NM cp2name[] = {
	//{ Info.CodePage , "Info.Name" , "Info.DisplayName" },
	{ 37 , "IBM037" , "IBM EBCDIC (US-Canada)" },
	{ 437 , "IBM437" , "OEM United States" },
	{ 500 , "IBM500" , "IBM EBCDIC (International)" },
	{ 708 , "ASMO-708" , "Arabic (ASMO 708)" },
	//{ 709 , "" , "Arabic (ASMO-449+, BCON V4)" },  // no name assigned in Microsoft table
	//{ 710 , "" , "Arabic - Transparent Arabic" },  // no name assigned in Microsoft table
	{ 720 , "DOS-720" , "Arabic (DOS)" },
	{ 737 , "ibm737" , "Greek (DOS)" },
	{ 775 , "ibm775" , "Baltic (DOS)" },
	{ 850 , "ibm850" , "Western European (DOS)" },
	{ 852 , "ibm852" , "Central European (DOS)" },
	{ 855 , "IBM855" , "OEM Cyrillic" },
	{ 857 , "ibm857" , "Turkish (DOS)" },
	{ 858 , "IBM00858" , "OEM Multilingual Latin I" },
	{ 860 , "IBM860" , "Portuguese (DOS)" },
	{ 861 , "ibm861" , "Icelandic (DOS)" },
	{ 862 , "DOS-862" , "Hebrew (DOS)" },
	{ 863 , "IBM863" , "French Canadian (DOS)" },
	{ 864 , "IBM864" , "Arabic (864)" },
	{ 865 , "IBM865" , "Nordic (DOS)" },
	{ 866 , "cp866" , "Cyrillic (DOS)" },
	{ 869 , "ibm869" , "Greek, Modern (DOS)" },
	{ 870 , "IBM870" , "IBM EBCDIC (Multilingual Latin-2)" },
	{ 874 , "windows-874" , "Thai (Windows)" },
	{ 875 , "cp875" , "IBM EBCDIC (Greek Modern)" },
	{ 932 , "shift_jis" , "Japanese (Shift-JIS)" },
	{ 936 , "gb2312" , "Chinese Simplified (GB2312)" },
	{ 949 , "ks_c_5601-1987" , "Korean" },
	{ 950 , "big5" , "Chinese Traditional (Big5)" },
	{ 1026 , "IBM1026" , "IBM EBCDIC (Turkish Latin-5)" },
	{ 1047 , "IBM01047" , "IBM Latin-1" },
	{ 1140 , "IBM01140" , "IBM EBCDIC (US-Canada-Euro)" },
	{ 1141 , "IBM01141" , "IBM EBCDIC (Germany-Euro)" },
	{ 1142 , "IBM01142" , "IBM EBCDIC (Denmark-Norway-Euro)" },
	{ 1143 , "IBM01143" , "IBM EBCDIC (Finland-Sweden-Euro)" },
	{ 1144 , "IBM01144" , "IBM EBCDIC (Italy-Euro)" },
	{ 1145 , "IBM01145" , "IBM EBCDIC (Spain-Euro)" },
	{ 1146 , "IBM01146" , "IBM EBCDIC (UK-Euro)" },
	{ 1147 , "IBM01147" , "IBM EBCDIC (France-Euro)" },
	{ 1148 , "IBM01148" , "IBM EBCDIC (International-Euro)" },
	{ 1149 , "IBM01149" , "IBM EBCDIC (Icelandic-Euro)" },
	{ 1200 , "utf-16" , "Unicode" },
	{ 1201 , "unicodeFFFE" , "Unicode (Big-Endian)" },
	{ 1250 , "windows-1250" , "Central European (Windows)" },
	{ 1251 , "windows-1251" , "Cyrillic (Windows)" },
	{ 1252 , "Windows-1252" , "ANSI Latin1; Western European (Windows)" },
	{ 1253 , "windows-1253" , "Greek (Windows)" },
	{ 1254 , "windows-1254" , "Turkish (Windows)" },
	{ 1255 , "windows-1255" , "Hebrew (Windows)" },
	{ 1256 , "windows-1256" , "Arabic (Windows)" },
	{ 1257 , "windows-1257" , "Baltic (Windows)" },
	{ 1258 , "windows-1258" , "Vietnamese (Windows)" },
	{ 1361 , "Johab" , "Korean (Johab)" },
	{ 10000 , "macintosh" , "Western European (Mac)" },
	{ 10001 , "x-mac-japanese" , "Japanese (Mac)" },
	{ 10002 , "x-mac-chinesetrad" , "Chinese Traditional (Mac)" },
	{ 10003 , "x-mac-korean" , "Korean (Mac)" },
	{ 10004 , "x-mac-arabic" , "Arabic (Mac)" },
	{ 10005 , "x-mac-hebrew" , "Hebrew (Mac)" },
	{ 10006 , "x-mac-greek" , "Greek (Mac)" },
	{ 10007 , "x-mac-cyrillic" , "Cyrillic (Mac)" },
	{ 10008 , "x-mac-chinesesimp" , "Chinese Simplified (Mac)" },
	{ 10010 , "x-mac-romanian" , "Romanian (Mac)" },
	{ 10017 , "x-mac-ukrainian" , "Ukrainian (Mac)" },
	{ 10021 , "x-mac-thai" , "Thai (Mac)" },
	{ 10029 , "x-mac-ce" , "Central European (Mac)" },
	{ 10079 , "x-mac-icelandic" , "Icelandic (Mac)" },
	{ 10081 , "x-mac-turkish" , "Turkish (Mac)" },
	{ 10082 , "x-mac-croatian" , "Croatian (Mac)" },
	{ 12000 , "utf-32" , "Unicode (UTF-32)" },
	{ 12001 , "utf-32BE" , "Unicode (UTF-32 Big-Endian)" },
	{ 20000 , "x-Chinese-CNS" , "Chinese Traditional (CNS)" },
	{ 20001 , "x-cp20001" , "TCA Taiwan" },
	{ 20002 , "x-Chinese-Eten" , "Chinese Traditional (Eten)" },
	{ 20003 , "x-cp20003" , "IBM5550 Taiwan" },
	{ 20004 , "x-cp20004" , "TeleText Taiwan" },
	{ 20005 , "x-cp20005" , "Wang Taiwan" },
	{ 20105 , "x-IA5" , "Western European (IA5)" },
	{ 20106 , "x-IA5-German" , "German (IA5)" },
	{ 20107 , "x-IA5-Swedish" , "Swedish (IA5)" },
	{ 20108 , "x-IA5-Norwegian" , "Norwegian (IA5)" },
	{ 20127 , "us-ascii" , "US-ASCII" },
	{ 20261 , "x-cp20261" , "T.61" },
	{ 20269 , "x-cp20269" , "ISO-6937" },
	{ 20273 , "IBM273" , "IBM EBCDIC (Germany)" },
	{ 20277 , "IBM277" , "IBM EBCDIC (Denmark-Norway)" },
	{ 20278 , "IBM278" , "IBM EBCDIC (Finland-Sweden)" },
	{ 20280 , "IBM280" , "IBM EBCDIC (Italy)" },
	{ 20284 , "IBM284" , "IBM EBCDIC (Spain)" },
	{ 20285 , "IBM285" , "IBM EBCDIC (UK)" },
	{ 20290 , "IBM290" , "IBM EBCDIC (Japanese katakana)" },
	{ 20297 , "IBM297" , "IBM EBCDIC (France)" },
	{ 20420 , "IBM420" , "IBM EBCDIC (Arabic)" },
	{ 20423 , "IBM423" , "IBM EBCDIC (Greek)" },
	{ 20424 , "IBM424" , "IBM EBCDIC (Hebrew)" },
	{ 20833 , "x-EBCDIC-KoreanExtended" , "IBM EBCDIC (Korean Extended)" },
	{ 20838 , "IBM-Thai" , "IBM EBCDIC (Thai)" },
	{ 20866 , "koi8-r" , "Cyrillic (KOI8-R)" },
	{ 20871 , "IBM871" , "IBM EBCDIC (Icelandic)" },
	{ 20880 , "IBM880" , "IBM EBCDIC (Cyrillic Russian)" },
	{ 20905 , "IBM905" , "IBM EBCDIC (Turkish)" },
	{ 20924 , "IBM00924" , "IBM Latin-1" },
	{ 20932 , "EUC-JP" , "Japanese (JIS 0208-1990 and 0212-1990)" },
	{ 20936 , "x-cp20936" , "Chinese Simplified (GB2312-80)" },
	{ 20949 , "x-cp20949" , "Korean Wansung" },
	{ 21025 , "cp1025" , "IBM EBCDIC (Cyrillic Serbian-Bulgarian)" },
	//{ 21027 , "" , "(deprecated)" },  // no name assigned in Microsoft table
	{ 21866 , "koi8-u" , "Cyrillic (KOI8-U)" },
	{ 28591 , "iso-8859-1" , "Western European (ISO)" },
	{ 28592 , "iso-8859-2" , "Central European (ISO)" },
	{ 28593 , "iso-8859-3" , "Latin 3 (ISO)" },
	{ 28594 , "iso-8859-4" , "Baltic (ISO)" },
	{ 28595 , "iso-8859-5" , "Cyrillic (ISO)" },
	{ 28596 , "iso-8859-6" , "Arabic (ISO)" },
	{ 28597 , "iso-8859-7" , "Greek (ISO)" },
	{ 28598 , "iso-8859-8" , "Hebrew (ISO-Visual)" },
	{ 28599 , "iso-8859-9" , "Turkish (ISO)" },
	{ 28603 , "iso-8859-13" , "Estonian (ISO)" },
	{ 28605 , "iso-8859-15" , "Latin 9 (ISO)" },
	{ 29001 , "x-Europa" , "Europa" },
	{ 38598 , "iso-8859-8-i" , "Hebrew (ISO-Logical)" },
	{ 50220 , "iso-2022-jp" , "Japanese (JIS)" },
	{ 50221 , "csISO2022JP" , "Japanese (JIS-Allow 1 byte Kana)" },
	{ 50222 , "iso-2022-jp" , "Japanese (JIS-Allow 1 byte Kana - SO/SI)" },
	{ 50225 , "iso-2022-kr" , "Korean (ISO)" },
	{ 50227 , "x-cp50227" , "Chinese Simplified (ISO-2022)" },
	//{ 50229 , "" , "ISO 2022 Traditional Chinese" },  // per Microsoft table
	{ 50229 , "iso-2022-cn" , "ISO 2022 Traditional Chinese" },  // Z.M my guess
	//{ 50930 , "" , "EBCDIC Japanese (Katakana) Extended" },  // no name assigned in Microsoft table
	//{ 50931 , "" , "EBCDIC US-Canada and Japanese" },  // no name assigned in Microsoft table
	//{ 50933 , "" , "EBCDIC Korean Extended and Korean" },  // no name assigned in Microsoft table
	//{ 50935 , "" , "EBCDIC Simplified Chinese Extended and Simplified Chinese" },  // no name assigned in Microsoft table
	//{ 50936 , "" , "EBCDIC Simplified Chinese" },  // no name assigned in Microsoft table
	//{ 50937 , "" , "EBCDIC US-Canada and Traditional Chinese" },  // no name assigned in Microsoft table
	//{ 50939 , "" , "EBCDIC Japanese (Latin) Extended and Japanese" },  // no name assigned in Microsoft table
	{ 51932 , "euc-jp" , "Japanese (EUC)" },
	{ 51936 , "EUC-CN" , "Chinese Simplified (EUC)" },
	{ 51949 , "euc-kr" , "Korean (EUC)" },
	//{ 51950 , "" , "EUC Traditional Chinese" },  // no name assigned in Microsoft table
	{ 52936 , "hz-gb-2312" , "Chinese Simplified (HZ)" },
	{ 54936 , "GB18030" , "Chinese Simplified (GB18030)" },
	{ 57002 , "x-iscii-de" , "ISCII Devanagari" },
	{ 57003 , "x-iscii-be" , "ISCII Bengali" },
	{ 57004 , "x-iscii-ta" , "ISCII Tamil" },
	{ 57005 , "x-iscii-te" , "ISCII Telugu" },
	{ 57006 , "x-iscii-as" , "ISCII Assamese" },
	{ 57007 , "x-iscii-or" , "ISCII Oriya" },
	{ 57008 , "x-iscii-ka" , "ISCII Kannada" },
	{ 57009 , "x-iscii-ma" , "ISCII Malayalam" },
	{ 57010 , "x-iscii-gu" , "ISCII Gujarati" },
	{ 57011 , "x-iscii-pa" , "ISCII Punjabi" },
	{ 65000 , "utf-7" , "Unicode (UTF-7)" },
	{ 65001 , "utf-8" , "Unicode (UTF-8)" },
};

typedef struct {
	char *m_charset;
	char *m_alt_charset;
} Name2AlternateName;

Name2AlternateName name2altname[] = {
	//{ "Name" , Info.Name },
	{ "gbk" , "GB18030" },    // { 54936 , "GB18030" , "Chinese Simplified (GB18030)" },
	{ "cp1252" , "Windows-1252" },    // { 1252 , "Windows-1252" , "Western European (Windows)" },
	{ "ascii" , "us-ascii" },    // { 20127 , "us-ascii" , "US-ASCII" },
	{ "cp819" , "iso-8859-1" },    // { 28591 , "iso-8859-1" , "Western European (ISO)" },
	{ "latin1" , "iso-8859-1" },    // { 28591 , "iso-8859-1" , "Western European (ISO)" },
	{ "latin2" , "iso-8859-2" },    // { 28592 , "iso-8859-2" , "Central European (ISO)" },
	{ "big-5" , "big5" },    // { 950 , "big5" , "Chinese Traditional (Big5)" },
	{ "SJIS" , "shift_jis" },    // { 932 , "shift_jis" , "Japanese (Shift-JIS)" },
};

#include <unordered_map>

typedef unordered_map<std::string, unsigned int> myMap;
static myMap *cids = 0;

void TextUtilsEx::delete_charset2Id()
{
	delete cids;
	cids = 0;
}

UINT TextUtilsEx::charset2Id(const char *char_set)
{
	UINT id = 0;
	CP2NM *item;
	myMap::iterator it;

	if (!cids) {
		cids = new myMap;

		int cp2name_size = sizeof(cp2name) / sizeof(CP2NM);
		for (int i = 0; i < cp2name_size; i++)
		{
			item = &cp2name[i];
			std::string charset = item->m_charset;
			std::transform(charset.begin(), charset.end(), charset.begin(), ::tolower);
			if (cids->find(charset) == cids->end()) {  // not found, invalid iterator returned
				cids->insert(myMap::value_type(charset, item->m_charsetId));
			}
		}
	}
#if 0
	for (it = cids->begin(); it != cids->end(); it++) {
		TRACE(L"%d %s\n", it->second, it->first);
	}
#endif
	std::string charset = char_set;
	std::transform(charset.begin(), charset.end(), charset.begin(), ::tolower);
	if ((it = cids->find(charset)) != cids->end()) {
		id = it->second;
	}
	else
	{
		int i;
		int nsize = sizeof(name2altname) / sizeof(Name2AlternateName);
		for (i = 0; i < nsize; i++)
		{
			if (strcmp(charset.c_str(), name2altname[i].m_charset) == 0)
			{
				std::string alt_charset = name2altname[i].m_alt_charset;
				std::transform(alt_charset.begin(), alt_charset.end(), alt_charset.begin(), ::tolower);
				if ((it = cids->find(alt_charset)) != cids->end())
					id = it->second;
			}
		}
	}
	return id;
}

typedef unordered_map<unsigned int, std::string> myIdMap;
static myIdMap *ids = 0;

void TextUtilsEx::delete_id2charset()
{
	delete ids;
	ids = 0;
}

BOOL TextUtilsEx::id2name(UINT codePage, CString& codePageName)
{
	CP2NM* item;
	int cp2name_size = sizeof(cp2name) / sizeof(CP2NM);
	for (int i = 0; i < cp2name_size; i++)
	{
		item = &cp2name[i];
		if (codePage == item->m_charsetId)
		{
			CStringA codePageNameA = item->m_charset;
			CString codePageNameW;
			DWORD error = 0;
			UINT cp = GetACP();
			DWORD dwFlags = 0;
			BOOL retA2W = TextUtilsEx::CodePage2WStr(&codePageNameA, cp, &codePageNameW, error, dwFlags);
			if (retA2W)
				codePageName = codePageNameW;
			return TRUE;
		}
	}
	return FALSE;
}

BOOL TextUtilsEx::id2charset(UINT id, std::string &charset)
{
	CP2NM *item;
	myIdMap::iterator it;

	if (!ids) {
		ids = new myIdMap;

		int cp2name_size = sizeof(cp2name) / sizeof(CP2NM);
		for (int i = 0; i < cp2name_size; i++)
		{
			item = &cp2name[i];
			if (ids->find(item->m_charsetId) == ids->end()) {  // not found, invalid iterator returned
				ids->insert(myIdMap::value_type(item->m_charsetId, item->m_charset));
			}
		}
	}
#if 0
	for (it = ids->begin(); it != ids->end(); it++) {
		TRACE(L"%d %s\n", it->first, it->second);
	}
#endif
	if ((it = ids->find(id)) != ids->end()) {
		charset = it->second;
		return TRUE;
	}
	return FALSE;
}

int TextUtilsEx::showCodePageTable(CString &path)
{
	CString codePageIdsFile = "WindowsCodePageIds.htm";
	CString fullPath = path + "\\" + codePageIdsFile;

	CFile fp;
	CFileException ExError;
	if (!fp.Open(fullPath, CFile::modeWrite | CFile::modeCreate, &ExError))
	{
		wchar_t szCause[2048];
		ExError.GetErrorMessage(szCause, 2048);

		CString txt = L"Could not create \"" + fullPath;
		txt += L"\" file.\n";
		txt += szCause;

		HWND h = NULL; // we don't have any window yet ??
		int answer = ::MessageBox(h, txt, L"Error", MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CStringA htmlHdr;

	htmlHdr += "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=US-ASCII\"></head><body>";
	htmlHdr += "<br><br><font size=\"+2\"><b>WINDOWS CODE PAGE IDENTIFIERS TABLE<br><br><br></font>";
	htmlHdr += "For additional information check the Microsoft document <a href=\"https://docs.microsoft.com/en-us/windows/desktop/intl/code-page-identifiers\">Code Page Identifiers</a>";
	htmlHdr += "<div style=\'text-decoration:underline\'><pre>\r\n";

	fp.Write((LPCSTR)htmlHdr, htmlHdr.GetLength());

	CP2NM *item;

	CStringA txt;
	txt.Format(" %22s  %26s  %42s  \n",
		" ", " ", " ");
	fp.Write((LPCSTR)txt, txt.GetLength());

	txt.Empty();
	txt.Format("|%22s |%26s |%42s |\n",
		"Code Page Identifiers", "Name", "Additional Information");
	fp.Write((LPCSTR)txt, txt.GetLength());

	int cp2name_size = sizeof(cp2name) / sizeof(CP2NM);
	for (int i = 0; i < cp2name_size; i++)
	{
		item = &cp2name[i];
		txt.Empty();
		txt.Format("|%22d |%26s |%42s |\n",
			item->m_charsetId, item->m_charset, item->m_info);
		fp.Write((LPCSTR)txt, txt.GetLength());
	}

	CStringA htmlEnd = "\r\n</pre></body></html>";
	fp.Write((LPCSTR)htmlEnd, htmlEnd.GetLength());

	fp.Close();

	ShellExecute(NULL, L"open", fullPath, NULL, NULL, SW_SHOWNORMAL);

	return 1;
}

BOOL TextUtilsEx::Id2LongInfo(UINT codePage, CString& codePageInfo)
{
	CStringA codePageInfoA;
	BOOL ret = TextUtilsEx::Id2LongInfoA(codePage, codePageInfoA);

	DWORD error = 0;
	DWORD dwFlags = 0;
	BOOL retA2W = TextUtilsEx::CodePage2WStr(&codePageInfoA, codePage, &codePageInfo, error, dwFlags);

	return retA2W;
}

BOOL TextUtilsEx::Id2LongInfoA(UINT codePage, CStringA &codePageInfo)
{
	CP2NM* item;
	//CString codePageInfo;
	int cp2name_size = sizeof(cp2name) / sizeof(CP2NM);
	for (int i = 0; i < cp2name_size; i++)
	{
		item = &cp2name[i];
		if (codePage == item->m_charsetId)
		{
			codePageInfo.Empty();
			codePageInfo.Format("%s %s", item->m_charset, item->m_info);
			return TRUE;
		}
	}
	return FALSE;
}

BOOL TextUtilsEx::GetCodePageInfo(UINT codePage, CP2NM& cpInfo)
{
	CP2NM* item;
	//CString codePageInfo;
	int cp2name_size = sizeof(cp2name) / sizeof(CP2NM);
	for (int i = 0; i < cp2name_size; i++)
	{
		item = &cp2name[i];
		if (codePage == item->m_charsetId)
		{
			cpInfo = *item;
			return TRUE;
		}
	}
	return FALSE;
}


BOOL TextUtilsEx::isNumeric(CString &str) {
	int i = 0;
	if (str[i] == '-')
		i++;
	if (i == str.GetLength())
		return FALSE;
	for (; i < str.GetLength(); i++) {
		if (!_istdigit(str[i])) {
			return FALSE;
		}
	}
	return TRUE;
}

BOOL TextUtilsEx::Str2WStr(CStringA &str, UINT strCodePage, CString &wstr, DWORD& error, DWORD dwFlags)
{
	BOOL ret = CodePage2WStr(&str, strCodePage, &wstr, error, dwFlags);
	return ret;
}

char *TextUtilsEx::strchar(char *beg, char *end, char c)
{
	while ((beg < end) && (*beg != c)) { beg++; };
	if (beg < end)
		return beg;
	else
		return 0;
}

char *TextUtilsEx::findOneOf(char *beg, char *end, char *charList)
{
	char oneOf;
	char *p_charList;
	char c;
	while (beg < end)
	{
		p_charList = charList;
		c = *beg;
		while (oneOf = *p_charList)
		{
			if (c == oneOf)
				return beg;
			else
				p_charList++;
		}
		beg++;
	};
	return 0;
}

char *TextUtilsEx::strnstrUpper2Lower(char *any, char *end, const char *lower, int lowerlength)
{
	// TODO: not very efficient; optimize
	char *p;
	for (p = any; p <= (end - lowerlength); p++)
	{
		if (strncmpUpper2Lower(p, end, lower, lowerlength) == 0)
			return p;
	}
	return 0;
}

int TextUtilsEx::strncmpUpper2Lower(char *any, char *end, const char *lower, int lowerlength) {
	// any can be multi line
	// lower is fixed length
	while ((any < end) && (lowerlength > 0) && (*lower++ == tolower(*any++))) { lowerlength--; }
	return lowerlength;
}

int TextUtilsEx::strncmpUpper2Lower(char *any, int anyLength, const char *lower, int lowerlength) {
	// any can be multi line
	// lower is fixed length
	if (anyLength >= lowerlength) {
		while ((lowerlength > 0) && (*lower++ == tolower(*any++))) { lowerlength--; }
	}
	return lowerlength;
}

int TextUtilsEx::strncmpExact(char *any, char *end, const char *lower, int lowerlength) {
	while ((lowerlength > 0) && (any < end) && (*lower++ == *any++)) { lowerlength--; }
	return lowerlength;
}

int TextUtilsEx::findNoCase(const char *input, int count, void const* Str, int  Size)
{
	int i;
	register char *p = (char*)input;
	register int delta_count = count - Size;
	for (i = 0; i < delta_count; i++, p++)
	{
		if (strncmpUpper2Lower(p, (count - i), (char*)Str, Size) == 0)
			return i;
	}
	return -1;
}

char *TextUtilsEx::findNoCaseP(const char *input, int count, void const* Str, int  Size)
{
	int i;
	register char *p = (char*)input;
	register int delta_count = count - Size;
	for (i = 0; i < delta_count; i++, p++)
	{
		if (strncmpUpper2Lower(p, (count - i), (char*)Str, Size) == 0)
			return p;
	}
	return 0;
}

int TextUtilsEx::DecodeURL(char *URL, int urlLen)
{
	int retLen = urlLen;
#if 0
	//	< to   &lt;
	//	> to   &gt;
	//	& to   &amp;
	//	" to   &quote;
	//	' to   &apos;
#endif
	int outLen = 0;
	int escapedCharLen = 0;
	char *p_out = URL;
	char *p_beg = URL;
	char *p = URL;
	int i = 0;
	for (i = 0; i < urlLen; i++)
	{
		escapedCharLen = 0;
		p = &URL[i];
		if (*p == '&')
		{
			char *val = p + 1;
			if (strncmp(val, "amp;", 4) == 0) {
				*p = '&';
				escapedCharLen = 4;
			}
			else if (strncmp(val, "lt;", 3) == 0) {
				*p = '<';
				escapedCharLen = 3;
			}
			else if (strncmp(val, "gt;", 3) == 0) {
				*p = '>';
				escapedCharLen = 3;
			}
			else if (strncmp(val, "quote;", 6) == 0) {
				*p = '"';
				escapedCharLen = 6;
			}
			else if (strncmp(val, "apos;", 5) == 0) {
				*p = '\'';
				escapedCharLen = 5;
			}

			if (escapedCharLen > 0)
			{
				int blockLen = IntPtr2Int(p - p_beg + 1);
				memcpy(p_out, p_beg, blockLen);
				p_out += blockLen;
				outLen += blockLen;
				p += escapedCharLen + 1;
				p_beg = p;
			}
		}
	}

	if (p != p_beg)
	{
		int blockLen = IntPtr2Int(p - p_beg + 1);
		memcpy(p_out, p_beg, blockLen);
		outLen += blockLen;
		p_out += blockLen;

	}

	retLen = IntPtr2Int(p_out - URL);
	if (outLen != retLen)
		int deb = 1;

	return retLen;
}

BOOL TextUtilsEx::isWhiteLine(CString& str)
{
	const wchar_t* p = (LPCWSTR)str;
	const wchar_t* e = p + str.GetLength();
	return isWhiteLine(p, e);
}

BOOL TextUtilsEx::isWhiteLine(const wchar_t* p, const wchar_t* e)
{
	while ((p < e) && ((*p == '\r') || (*p == '\n') || (*p == ' ') || (*p == '\t')))  // eat white
		p++;
	if (p == e)
		return TRUE;
	else
		return FALSE;
}

BOOL TextUtilsEx::isWhiteLineA(CStringA &str)
{
	const char* p = (LPCSTR)str;
	const char* e = p + str.GetLength();
	return isWhiteLineA(p, e);
}

BOOL TextUtilsEx::isWhiteLineA(const char* p, const char* e)
{
	while ((p < e) && ((*p == '\r') || (*p == '\n') || (*p == ' ') || (*p == '\t')))  // eat white
		p++;
	if (p == e)
		return TRUE;
	else
		return FALSE;
}

// Very inefficient
// SplitStringW and SplitStringA candidates for template ??
//
#if 0
template <typename T, typename A, typename C>
static int Tokenize(T& str, A& a, C del)
//void TextUtilsEx::SplitString(const T& strIn, const T& delim, A& a)
{
	enum { inHunt, notInsideDoubleQuotes, insideDoubleQuotes };

	int state = inHunt;
	T token;

	int strLength = str.GetLength();
	for (int i = 0; i < strLength; i++)
	{
		C c = str.GetAt(i);
		if (state == inHunt)
		{
			if ((c == ' ') || (c == '\t'))
				continue;
			else if (c == del)  // separator
				; // state = notInsideDoubleQuotes;  // stay inHunt
			else if (c == '"')
			{
				state = insideDoubleQuotes;
				//token.AppendChar(c);  // doesn't work, best is to ignore
			}
			else if ((c != del) && (c != '"'))  // No need to check just Append
			{
				state = notInsideDoubleQuotes;
				token.AppendChar(c);
			}
		}
		else if (state == notInsideDoubleQuotes)
		{
			int bytesLeft = strLength - i;
			if (c == '"')
			{
				if ((bytesLeft >= 2) && (str.GetAt(i + 1) == '"'))
				{
					token.AppendChar('"');
					i += 1;
				}
				else
					token.AppendChar(c);
			}
			else if (c != del)
			{
				token.AppendChar(c);
			}
			else
			{
				token.Trim(C(" \t"));
				if (!token.IsEmpty())
					a.Add(token);
				token.Empty();
				state = inHunt;
			}
		}
		else  // state == insideDoubleQuotes
		{
			if (c != '"')
			{
				token.AppendChar(c);
			}
			else
			{
				int bytesLeft = strLength - i;
				if ((bytesLeft >= 2) && (str.GetAt(i + 1) == '"'))
				{
					token.AppendChar('"');
					i += 1;
				}
				else
				{
					token.Trim(C(" \t"));
					if (!token.IsEmpty())
						a.Add(token);
					token.Empty();
					state = inHunt;
				}
			}
		}
	}

	token.Trim(C(" \t"));
	if (!token.IsEmpty())
		a.Add(token);
	return 1;
}

#endif

void TextUtilsEx::SplitStringW(const CString &strIn, const CString &delim, CStringArray &a) 
{
	int position = 0;
	CString strToken;

	a.RemoveAll();
	strToken = strIn.Tokenize(delim, position);
	while (!strToken.IsEmpty()) {
		strToken.Trim(L" \t");
		if (!strToken.IsEmpty())
			a.Add(strToken);
		strToken = strIn.Tokenize(delim, position);
	}
}

// Simple tokenizer, heavy
// Will support: a bc d, xy z, "k l, n,'"'","a b", "a,b"
// '"' escapes " character

int TextUtilsEx::TokenizeW(CString &str, CStringArray &a, wchar_t del)
{
	enum { inHunt, notInsideDoubleQuotes, insideDoubleQuotes };

	int state = inHunt;
	CString token;

	int strLength = str.GetLength();
	for (int i = 0; i < strLength; i++)
	{
		wchar_t c = str.GetAt(i);
		if (state == inHunt)
		{
			if ((c == L' ') || (c == L'\t'))
				continue;
			else if (c == del)  // separator
				; // state = notInsideDoubleQuotes;  // stay inHunt
			else if (c == L'"')
			{
				state = insideDoubleQuotes;
				//token.AppendChar(c);  // doesn't work, best is to ignore
			}
			else if ((c != del) && (c != L'"'))  // No need to check just Append
			{
				state = notInsideDoubleQuotes;
				token.AppendChar(c);
			}
		}
		else if (state == notInsideDoubleQuotes)
		{
			int bytesLeft = strLength - i;
			if (c == L'"')
			{
				if ((bytesLeft >= 2) && (str.GetAt(i + 1) == L'"'))
				{
					token.AppendChar(L'"');
					i += 1;
				}
				else
					token.AppendChar(c);
			}
			else if (c != del)
			{
				token.AppendChar(c);
			}
			else
			{
				token.Trim(L" \t");
				if (!token.IsEmpty())
					a.Add(token);
				token.Empty();
				state = inHunt;
			}
		}
		else  // state == insideDoubleQuotes
		{
			if (c != L'"')
			{
				token.AppendChar(c);
			}
			else
			{
				int bytesLeft = strLength - i;
				if ((bytesLeft >= 2) && (str.GetAt(i + 1) == L'"'))
				{
					token.AppendChar(L'"');
					i += 1;
				}
				else
				{
					token.Trim(L" \t");
					if (!token.IsEmpty())
						a.Add(token);
					token.Empty();
					state = inHunt;
				}
			}
		}
	}

	token.Trim(L" \t");
	if (!token.IsEmpty())
		a.Add(token);
	return 1;
}

void TextUtilsEx::TraceStringArrayW(CStringArray &a)
{
	CString el;
	TRACE(L"StringArray:\n");
	for (int i = 0; i < a.GetSize(); i++)
	{
		el = a.ElementAt(i);
		TRACE(L"\t|%s|\n", el);
	}
}

// Very inefficient

void TextUtilsEx::SplitStringA(const CStringA& strIn, const CStringA& delim, CStringArray& a)
{
	int position = 0;
	CStringA strToken;

	a.RemoveAll();
	strToken = strIn.Tokenize(delim, position);
	while (!strToken.IsEmpty())
	{
		strToken.Trim(" \t");
		if (!strToken.IsEmpty())
		{
			CString strTokenW = strToken;
			a.Add(strTokenW);
		}
		strToken = strIn.Tokenize(delim, position);
	}
}

void TextUtilsEx::SplitStringA2A(const CStringA& strIn, const CStringA& delim, CStringArrayA& a)
{
	int position = 0;
	CStringA strToken;

	a.RemoveAll();
	strToken = strIn.Tokenize(delim, position);
	while (!strToken.IsEmpty())
	{
		strToken.Trim(" \t");
		if (!strToken.IsEmpty())
		{
			a.Add(strToken);
		}
		strToken = strIn.Tokenize(delim, position);
	}
}

// Simple tokenizer, heavy
// Will support: a bc d, xy z, "k l, n,'"'","a b", "a,b"
// '"' escapes " character


int TextUtilsEx::TokenizeA(CStringA& str, CStringArrayA& a, char del)
{
	enum { inHunt, notInsideDoubleQuotes, insideDoubleQuotes };

	int state = inHunt;
	CStringA token;

	int strLength = str.GetLength();
	for (int i = 0; i < strLength; i++)
	{
		char c = str.GetAt(i);
		if (state == inHunt)
		{
			if ((c == ' ') || (c == '\t'))
				continue;
			else if (c == del)  // separator
				; // state = notInsideDoubleQuotes;  // stay inHunt
			else if (c == '"')
			{
				state = insideDoubleQuotes;
				//token.AppendChar(c);  // doesn't work, best is to ignore
			}
			else if ((c != del) && (c != '"'))  // No need to check just Append
			{
				state = notInsideDoubleQuotes;
				token.AppendChar(c);
			}
		}
		else if (state == notInsideDoubleQuotes)
		{
			int bytesLeft = strLength - i;
			if (c == '"')
			{
				if ((bytesLeft >= 2) && (str.GetAt(i + 1) == '"'))
				{
					token.AppendChar('"');
					i += 1;
				}
				else
					token.AppendChar(c);
			}
			else if (c != del)
			{
				token.AppendChar(c);
			}
			else
			{
				token.Trim(" \t");
				if (!token.IsEmpty())
					a.Add(token);
				token.Empty();
				state = inHunt;
			}
		}
		else  // state == insideDoubleQuotes
		{
			if (c != '"')
			{
				token.AppendChar(c);
			}
			else
			{
				int bytesLeft = strLength - i;
				if ((bytesLeft >= 2) && (str.GetAt(i + 1) == '"'))
				{
					token.AppendChar('"');
					i += 1;
				}
				else
				{
					token.Trim(" \t");
					if (!token.IsEmpty())
						a.Add(token);
					token.Empty();
					state = inHunt;
				}
			}
		}
	}

	token.Trim(" \t");
	if (!token.IsEmpty())
		a.Add(token);

	return 1;
}

void TextUtilsEx::TraceStringArrayA(CStringArray& a)
{
	CStringA el;
	TRACE(L"StringArray:\n");
	for (int i = 0; i < a.GetSize(); i++)
	{
		el = a.ElementAt(i);
		TRACE("\t|%s|\n", el);
	}
}

// based on MimeCode.cpp

int TextUtilsEx::SelectEncoding(int nLength, int nNonAsciiChars)
{
	int nQEncodeSize = nLength + nNonAsciiChars * 2;
	int nBEncodeSize = (nLength + 2) / 3 * 4;
	return (nQEncodeSize <= nBEncodeSize || nNonAsciiChars * 5 <= nLength) ? 'Q' : 'B';
}

int TextUtilsEx::WordEncode(CStringA& txt, CStringA& encodedTxt, int encodeType)
{
	int nLength = txt.GetLength();

#if 0
	// int SelectEncoding(int nLength, int nNonAsciiChars)
	int nQEncodeSize = nLength + nNonAsciiChars * 2;
	int nBEncodeSize = (nLength + 2) / 3 * 4;
	int encMethos =  ((nQEncodeSize <= nBEncodeSize) || (nNonAsciiChars * 5 <= nLength)) ? 'Q' : 'B';
#endif

	CMimeEncodedWord coder;
	char* charset = "UTF-8";
	const char* pszInput = (LPCSTR)txt;
	//int encType = SelectEncoding(nUnitSize, nNonAsciiChars);
	// Hardcoded
	int encType = encodeType;
	coder.SetEncoding(encType, charset);
	coder.SetInput(pszInput, nLength, true);

	int outLength = coder.GetOutputLength();
	if (outLength <= 0)
		return 0;

	int encodedTxlength = 2 * outLength + 2;

	char* buff = encodedTxt.GetBuffer(encodedTxlength);
	int nEncoded = coder.GetOutput((unsigned char*)buff, encodedTxlength);
	if (nEncoded <= 0)
		nEncoded = 0;
	encodedTxt.ReleaseBuffer(nEncoded);

#if 0
	// just a test
	CStringA strCharset;
	UINT charsetId;
	UINT toCharacterId = 0; // ignored for now
	CStringA result = TextUtilsEx::DecodeString(encodedTxt, strCharset, charsetId, toCharacterId);
#endif

	return nEncoded;
}

void TextUtilsEx::WStr2WstrWithCommas(const wchar_t* wstr, int wstrlen, CString& cstr)
{
	_ASSERTE(wstr[wstrlen] == 0);
	cstr.Empty();
	if (wstrlen <= 0)
		return;

	int wlen = wstrlen;
	int j = 0;
	while (!iswdigit(wstr[j++]))
		wlen--;

#if 0
	int groups = wlen / 3;
	int partial = wlen % 3;

	if ((groups > 0) && (partial == 0))
		groups--;
#else
	int groups = (wlen -1)/3;
#endif

	int i = 0;
	int len = wstrlen - groups * 3;
	cstr.Append(wstr, len);
	for (i = groups; i > 0; i--)
	{
		cstr.AppendChar(L',');
		cstr.Append(&wstr[len], 3);
		len += 3;
	}

	const int deb = 1;
}

// No need to optimize perf so rely on WStr2WstrWithCommas()
void TextUtilsEx::Int2WstrWithCommas(INT64 numb, CString& cstr)
{
	CString str;
	str.Format(L"%lld", numb);
	WStr2WstrWithCommas((LPCWSTR)str, str.GetLength(), cstr);
	const int deb = 1;
}
