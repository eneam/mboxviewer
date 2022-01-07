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

#if 1
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
				len = p - p_beg;
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
	len = p - p_beg;
	if (len > 0)
	{
		out->Append(p_beg, len);
		out->Append("\r\n", 2);
	}
}
#else
void ReplaceNL2CRNL(const char *in, int inLength, SimpleString *out)
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
		while ((p < e) && (*p != '\n')) p++;
		if (p < e)  // found NL
		{
			if (p == in)
			{
				out->Append("\r\n", 2);
				p++;  // jump over NL
				p_beg = p;
			}
			else if (*(p - 1) != '\r')
			{
				len = p - p_beg;
				if (len > 0)
					out->Append(p_beg, len);

				out->Append("\r\n", 2);
				p++;  // jump over NL
				p_beg = p;
			}
			else
				p++;
		}
		else // p >= e didn't found NL done with looping
		{
			// we can be here if input is not terminated by NL
			len = p - p_beg;
			if (len > 0)
			{
				out->Append(p_beg, len);
				out->Append("\r\n", 2);
			}
		}
	}
}
#endif

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
	while (p < e)
	{
		c = *p;
		// TODO: implement as table and evalute performance
		if ((c == '>') || (c == '<') || (c == '\"') || (c == '&') || (c == '\''))
		{
			len = p - p_beg;
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
			len = p - p_beg;
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
			len = p - p_beg;
			if (len > 0)  out->Append(p_beg, len);
			out->Append("<br>\n", 5);
			p++;
			p_beg = p;
		}
		else
			p++;
	}
	len = p - p_beg;
	if (len > 0)
	{
		out->Append(p_beg, len);
	}
}

BOOL TextUtilsEx::Wide2Ansi(CStringW &strW, CString &strA, DWORD &error)
{
	UINT outCodePage = CP_ACP;
	BOOL ret = WStr2CodePage((wchar_t*)((LPCWSTR)strW), strW.GetLength(), outCodePage, &strA, error);
	return ret;
}

BOOL TextUtilsEx::Ansi2Wide(CString &strA, CStringW &strW, DWORD &error)
{
	UINT strCodePage = CP_ACP;
	BOOL ret = CodePage2WStr(&strA, strCodePage, &strW, error);
	return ret;
}

UINT TextUtilsEx::Str2PageCode(const  char* PageCodeStr)
{
	UINT CodePage = 0;

	CodePage = TextUtilsEx::charset2Id(PageCodeStr);

	return CodePage;
}

// TODO: Find time to reduce number of conversion functions, duplicates

DWORD TextUtilsEx::Str2Ansi(CString &str, UINT strCodePage)
{
	int len = str.GetLength() * 4 + 2;
	LPWSTR buff = (LPWSTR)malloc(len);  // or  we could call MultiByteToWideChar first to get the required length
	int len1 = MultiByteToWideChar(strCodePage, 0, str, str.GetLength(), buff, len);
	if (len1 == 0) {
		free(buff);
		// error - implement error log file
		const DWORD error = ::GetLastError();
		return error;
	}
	char * buff1 = (char *)malloc(len1 + 2); // or could  call WideCharToMultiByte first to get the required length
	int len2 = WideCharToMultiByte(CP_ACP, 0, buff, len1, buff1, len1 + 1, NULL, NULL);
	if (len2 == 0) {
		free(buff);
		free(buff1);
		// error - implement error log file
		const DWORD error = ::GetLastError();
		/*ERROR_INSUFFICIENT_BUFFER.A supplied buffer size was not large enough, or it was incorrectly set to NULL.
		ERROR_INVALID_FLAGS.The values supplied for flags were not valid.
		ERROR_INVALID_PARAMETER.Any of the parameter values was invalid.
		ERROR_NO_UNICODE_TRANSLATION.Invalid Unicode was found in a string.*/
		return error;
	}
	buff1[len2] = 0;
	str = buff1;
	free(buff);
	free(buff1);

	return 0;
}

BOOL TextUtilsEx::Str2CodePage(const char *str, int strlen, UINT inCodePage, UINT outCodePage, SimpleString *result, SimpleString *workBuff)
{
	int buffLen = strlen * 4 + 2;
	workBuff->ClearAndResize(buffLen);
	LPWSTR buff = (LPWSTR)workBuff->Data();  // or  we could call MultiByteToWideChar first to get the required length
	int wlen = MultiByteToWideChar(inCodePage, 0, str, strlen, buff, buffLen);
	if (wlen == 0) {
		result->Clear();
		// error - implement error log file
		const DWORD error = ::GetLastError();
		return FALSE;
	}

	int outLen = wlen * 4 + 2;
	result->ClearAndResize(outLen); // or could  call WideCharToMultiByte first to get the required length
	if (outCodePage == 0)
		outCodePage = CP_UTF8;
	//int utf8Len = WideCharToMultiByte(CP_UTF8,     0, buff, wlen, result->Data(), outLen, NULL, NULL);
	int utf8Len = WideCharToMultiByte(outCodePage, 0, buff, wlen, result->Data(), outLen, NULL, NULL);
	if (utf8Len == 0) {
		result->Clear();
		// error - implement error log file
		const DWORD error = ::GetLastError();
		/*ERROR_INSUFFICIENT_BUFFER.A supplied buffer size was not large enough, or it was incorrectly set to NULL.
		ERROR_INVALID_FLAGS.The values supplied for flags were not valid.
		ERROR_INVALID_PARAMETER.Any of the parameter values was invalid.
		ERROR_NO_UNICODE_TRANSLATION.Invalid Unicode was found in a string.*/
		return FALSE;
	}
	result->SetCount(utf8Len);
	return TRUE;
}

BOOL TextUtilsEx::Str2CodePage(SimpleString *str, UINT inCodePage, UINT outCodePage, SimpleString *result, SimpleString *workBuff)
{
	int buffLen = str->Count() * 4 + 2;
	workBuff->ClearAndResize(buffLen);
	LPWSTR buff = (LPWSTR)workBuff->Data();  // or  we could call MultiByteToWideChar first to get the required length
	int wlen = MultiByteToWideChar(inCodePage, 0, str->Data(), str->Count(), buff, buffLen);
	if (wlen == 0) {
		result->Clear();
		// error - implement error log file
		const DWORD error = ::GetLastError();
		return FALSE;
	}

	int outLen = wlen * 4 + 2;
	result->ClearAndResize(outLen); // or could  call WideCharToMultiByte first to get the required length
	if (outCodePage == 0)
		outCodePage = CP_UTF8;
	//int utf8Len = WideCharToMultiByte(CP_UTF8,     0, buff, wlen, result->Data(), outLen, NULL, NULL);
	int utf8Len = WideCharToMultiByte(outCodePage, 0, buff, wlen, result->Data(), outLen, NULL, NULL);
	if (utf8Len == 0) {
		result->Clear();
		// error - implement error log file
		const DWORD error = ::GetLastError();
		/*ERROR_INSUFFICIENT_BUFFER.A supplied buffer size was not large enough, or it was incorrectly set to NULL.
		ERROR_INVALID_FLAGS.The values supplied for flags were not valid.
		ERROR_INVALID_PARAMETER.Any of the parameter values was invalid.
		ERROR_NO_UNICODE_TRANSLATION.Invalid Unicode was found in a string.*/
		return FALSE;
	}
	result->SetCount(utf8Len);
	return TRUE;
}

BOOL TextUtilsEx::WStr2CodePage(wchar_t *wbuff, int wlen, UINT outCodePage, CString *result, DWORD &error)
{
	SimpleString outbuf(4 * wlen + 2);

	BOOL ret = WStr2CodePage(wbuff, wlen, outCodePage, &outbuf, error);

	*result = outbuf.Data();
	return ret;
}

BOOL TextUtilsEx::WStr2CodePage(wchar_t *wbuff, int wlen, UINT outCodePage, SimpleString *result, DWORD &error)
{
	error = 0;
	int outLen = wlen * 4 + 2;
	result->ClearAndResize(outLen); // or could  call WideCharToMultiByte first to get the required length
	if (outCodePage == 0)
		outCodePage = CP_UTF8;

	int utf8Len = WideCharToMultiByte(outCodePage, 0, wbuff, wlen, result->Data(), outLen, NULL, NULL);
	if (utf8Len == 0) 
	{
		result->Clear();
		// error - implement error log file
		error = ::GetLastError();
		/*ERROR_INSUFFICIENT_BUFFER.A supplied buffer size was not large enough, or it was incorrectly set to NULL.
		ERROR_INVALID_FLAGS.The values supplied for flags were not valid.
		ERROR_INVALID_PARAMETER.Any of the parameter values was invalid.
		ERROR_NO_UNICODE_TRANSLATION.Invalid Unicode was found in a string.*/
		return FALSE;
	}
	result->SetCount(utf8Len);
	return TRUE;
}

BOOL TextUtilsEx::CodePage2WStr(SimpleString *str, UINT strCodePage, SimpleString *wstr, DWORD &error)
{
	error = 0;
	// SimpleString doesn't support UTF16 characters so be carfull
	int wbuffLen = str->Count() * 4 + 2;
	wstr->ClearAndResize(wbuffLen);

	int wlen = MultiByteToWideChar(strCodePage, 0, str->Data(), str->Count(), (LPWSTR)((void*)wstr->Data()), wbuffLen);
	if (wlen == 0) {
		wstr->Clear();
		// error - implement error log file
		error = ::GetLastError();
		return FALSE;
	}
	wstr->SetCount(wlen * 2);
	*wstr->Data(wlen * 2 + 1) = 0;
	*wstr->Data(wlen * 2 + 2) = 0;
	wchar_t *wbuf = (LPWSTR)((void*)wstr->Data());
	return TRUE;
}

BOOL TextUtilsEx::CodePage2WStr(CString *str, UINT strCodePage, SimpleString *wstr, DWORD &error)
{
	error = 0;
	int wbuffLen = str->GetLength() * 4 + 2;
	wstr->ClearAndResize(wbuffLen);

	int wlen = MultiByteToWideChar(strCodePage, 0, str->operator LPCSTR(), str->GetLength(), (LPWSTR)((void*)wstr->Data()), wbuffLen);
	if (wlen == 0) {
		wstr->Clear();
		// error - implement error log file
		error = ::GetLastError();
		return FALSE;
	}
	wstr->SetCount(wlen * 2);
	*wstr->Data(wlen * 2 + 1) = 0;
	*wstr->Data(wlen * 2 + 2) = 0;
	wchar_t *wbuf = (LPWSTR)((void*)wstr->Data());
	return TRUE;
}

BOOL TextUtilsEx::CodePage2WStr(CString *str, UINT strCodePage, CStringW *wstr, DWORD &error)
{
	error = 0;
	int wbuffLen = str->GetLength() * 4 + 2;
	SimpleString ss_wstr(wbuffLen);

	int wlen = MultiByteToWideChar(strCodePage, 0, str->operator LPCSTR(), str->GetLength(), (LPWSTR)((void*)ss_wstr.Data()), wbuffLen);
	if (wlen == 0) 
	{
		wstr->Empty();
		// error - implement error log file
		error = ::GetLastError();
		return FALSE;
	}

	ss_wstr.SetCount(wlen * 2);
	ss_wstr.SetAt(wlen * 2 + 1, 0);
	ss_wstr.SetAt(wlen * 2 + 2, 0);

	wchar_t *wbuf = (LPWSTR)((void*)ss_wstr.Data());
	*wstr = wbuf;

	return TRUE;
}

BOOL TextUtilsEx::Str2UTF8(SimpleString *str, UINT strCodePage, SimpleString *result, SimpleString *workBuff)
{
	BOOL ret = Str2CodePage(str, strCodePage, CP_UTF8, result, workBuff);
	return ret;
}

BOOL TextUtilsEx::Str2CurrentCodepage(SimpleString *str, UINT strCodePage, SimpleString *result, SimpleString *workBuff)
{
	UINT currentCodePage = GetACP();
	BOOL ret = Str2CodePage(str, strCodePage, currentCodePage, result, workBuff);
	return ret;
}

#include "MimeCode.h"

CString TextUtilsEx::DecodeString(CString &subj, CString &charset, UINT &charsetId, UINT toCharacterId)
{
	CFieldCodeText tfc;
	tfc.SetInput(subj.GetBuffer(), subj.GetLength(), false);
	int outputLen = tfc.GetOutputLength();
	if (outputLen > 0) 
	{
		int maxOutputLen = 2 * outputLen + 2;
		unsigned char *outBuf = (unsigned char*)malloc(maxOutputLen);
		int decodeLen = tfc.GetOutput(outBuf, maxOutputLen);
		outBuf[decodeLen] = 0;
		CString str(outBuf);
		free(outBuf);
		charset = tfc.GetCharset();
		UINT CodePage = TextUtilsEx::Str2PageCode(tfc.GetCharset());
		charsetId = CodePage;
		// No remapping for now. Header of message windows should show text properly
		return str;

		if ((CodePage > 0) && (toCharacterId > 0)) {
			TextUtilsEx::Str2Ansi(str, CodePage);
			return str;
		}
		else
			return subj;
	}
	else
		return subj;
}

int  TextUtilsEx::hextob(char ch)
{
	if (ch >= '0' && ch <= '9') return ch - '0';
	if (ch >= 'A' && ch <= 'F') return ch - 'A' + 10;
	if (ch >= 'a' && ch <= 'f') return ch - 'a' + 10;
	return -1;
}

int TextUtilsEx::DecodeMimeChunkedString(CString &inString, CString &charset, UINT &charsetId,  BOOL hasCharset, CString &outString)
{
	const char *p_beg = (LPCSTR)inString;
	const char *p = p_beg;
	int length = inString.GetLength();
	const char *e = p + length;

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

		charset.Append(p_beg, p - p_beg);
		charsetId = TextUtilsEx::Str2PageCode(charset);

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
					return 0;
				c = *p++;
				int second = hextob(c);
				if (second < 0)
					return 0;
				unsigned char uc = first * 16 + second;
				out.Append(uc);
			}
			else
			{
				out.Append(c);
				p++;
			}
		}

		outString.Empty();
		outString.Append(out.Data(), out.Count());
	}

	return 1;
}

CP2NM cp2name[] = {
	//{ Info.CodePage , "Info.Name" , "Info.DisplayName" },
	{ 37 , "IBM037" , "IBM EBCDIC (US-Canada)" },
	{ 437 , "IBM437" , "OEM United States" },
	{ 500 , "IBM500" , "IBM EBCDIC (International)" },
	{ 708 , "ASMO-708" , "Arabic (ASMO 708)" },
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
	{ 1252 , "Windows-1252" , "Western European (Windows)" },
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
	{ 51932 , "euc-jp" , "Japanese (EUC)" },
	{ 51936 , "EUC-CN" , "Chinese Simplified (EUC)" },
	{ 51949 , "euc-kr" , "Korean (EUC)" },
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
		TRACE(_T("%d %s\n"), it->second, it->first);
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
		TRACE(_T("%d %s\n"), it->first, it->second);
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
		TCHAR szCause[2048];
		ExError.GetErrorMessage(szCause, 2048);

		CString txt = _T("Could not create \"") + fullPath;
		txt += _T("\" file.\n");
		txt += szCause;

		HWND h = NULL; // we don't have any window yet ??
		int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
		return -1;
	}

	CString htmlHdr;

	htmlHdr += "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=US-ASCII\"></head><body>";
	htmlHdr += "<br><br><font size=\"+2\"><b>WINDOWS CODE PAGE IDENTIFIERS TABLE<br><br><br></font>";
	htmlHdr += "For additional information check the Microsoft document <a href=\"https://docs.microsoft.com/en-us/windows/desktop/intl/code-page-identifiers\">Code Page Identifiers</a>";
	htmlHdr += "<div style=\'text-decoration:underline\'><pre>\r\n";

	fp.Write((LPCSTR)htmlHdr, htmlHdr.GetLength());

	CP2NM *item;

	CString txt;
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

	CString htmlEnd = "\r\n</pre></body></html>";
	fp.Write((LPCSTR)htmlEnd, htmlEnd.GetLength());

	fp.Close();

	ShellExecute(NULL, _T("open"), fullPath, NULL, NULL, SW_SHOWNORMAL);

	int deb = 1;
	return 1;
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

BOOL TextUtilsEx::Str2Wide(CString &res, UINT CodePage, CStringW &m_strW)
{
	int len = res.GetLength() * 4 + 2;
	LPWSTR buff = (LPWSTR)malloc(len);  // or  we could call MultiByteToWideChar first to get the required length
	int len1 = MultiByteToWideChar(CodePage, 0, res, res.GetLength(), buff, len);
	if (len1 == 0) {
		free(buff);
		// error - implement error log file
		const DWORD error = ::GetLastError();
		return FALSE;
	}
	buff[len1] = 0;
	m_strW = buff;
	free(buff);
	return TRUE;
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
				int blockLen = p - p_beg + 1;
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
		int blockLen = p - p_beg + 1;
		memcpy(p_out, p_beg, blockLen);
		outLen += blockLen;
		p_out += blockLen;

	}

	retLen = p_out - URL;
	if (outLen != retLen)
		int deb = 1;

	return retLen;
}

BOOL TextUtilsEx::isWhiteLine(CString &str)
{
	const char* p = (LPCSTR)str;
	const char* e = p + str.GetLength();
	return isWhiteLine(p, e);
}

BOOL TextUtilsEx::isWhiteLine(const char* p, const char* e)
{
	while ((p < e) && ((*p == '\r') || (*p == '\n') || (*p == ' ') || (*p == '\t')))  // eat white
		p++;
	if (p == e)
		return TRUE;
	else
		return FALSE;
}

// Very inefficient
void TextUtilsEx::SplitString(const CString &strIn, const CString &delim, CStringArray &a) 
{
	int position = 0;
	CString strToken;

	a.RemoveAll();
	strToken = strIn.Tokenize(delim, position);
	while (!strToken.IsEmpty()) {
		a.Add(strToken);
		strToken = strIn.Tokenize(delim, position);
	}
}

// Simple tokenizer, heavy
// Will support: a bc d, xy z, "k l, n,'"'","a b", "a,b"
// '"' escapes " character

int TextUtilsEx::Tokenize(CString str, CStringArray &a, char del)
{
	enum { inHunt, notInsideDoubleQuotes, insideDoubleQuotes};

	int state = inHunt;
	CString token;

	int strLength = str.GetLength();
	for (int i = 0; i < strLength; i++)
	{
		char c = str.GetAt(i);
		if (state == inHunt)
		{
			if ((c == ' ') || (c == '\t'))
				continue;
			else if (c == '\'') 
			{
				int bytesLeft = strLength - i;
				if ((bytesLeft >= 3) && (str.GetAt(i + 1) == '\"') && (str.GetAt(i + 2) == '\''))
				{
					token.AppendChar('"');
					i += 2;
				}
				else
					token.AppendChar(c);
				state = notInsideDoubleQuotes;
				continue;
			}
			else if (c == del)
				state = notInsideDoubleQuotes;
			else if (c == '"')
				state = insideDoubleQuotes;
			else if ((c != del) && (c != '"'))
			{
				state = notInsideDoubleQuotes;
				token.AppendChar(c);
			}
			continue;
		}
		if (state == notInsideDoubleQuotes)
		{
			if (c == '\'')
			{
				int bytesLeft = strLength - i;
				if ((bytesLeft >= 3) && (str.GetAt(i + 1) == '\"') && (str.GetAt(i + 2) == '\''))
				{
					token.AppendChar('"');
					i += 2;
				}
				else
					token.AppendChar(c);
				continue;
			}
			else if (c != del)
			{
				token.AppendChar(c);
			}
			else
			{
				token.Trim(" \t");
				a.Add(token);
				token.Empty();
				state = inHunt;
			}
			continue;
		}
		else  // state == insideDoubleQuotes
		{
			if (c == '\'')
			{
				int bytesLeft = strLength - i;
				if ((bytesLeft >= 3) && (str.GetAt(i + 1) == '\"') && (str.GetAt(i + 2) == '\''))
				{
					token.AppendChar('"');
					i += 2;
				}
				else
					token.AppendChar(c);
				continue;
			}
			else if (c != '"')
			{
				token.AppendChar(c);
			}
			else
			{
				token.Trim(" \t");
				a.Add(token);
				token.Empty();
				state = inHunt;
			}
		}
	}
	
	token.Trim(" \t");
	if (!token.IsEmpty())
		a.Add(token);
	return 1;
}

void TextUtilsEx::TraceStringArray(CStringArray &a)
{
	CString el;
	int first = 0;
	TRACE(_T("StringArray:\n"));
	for (int i = 0; i < a.GetSize(); i++)
	{
		el = a.ElementAt(i);
		TRACE(_T("\t|%s|\n"), el);
	}
}
