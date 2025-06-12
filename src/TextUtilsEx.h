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

#include <string>

class SimpleString;

typedef CArray<CStringA> CStringArrayA;

typedef struct {
	UINT m_charsetId;
	char* m_charset;
	char* m_info;
} CP2NM;

class TextUtilsEx
{
public:
	//
	static BOOL WStr2CodePage(const wchar_t* wbuff, int wlen, UINT outCodePage, SimpleString* resultA, DWORD& error);
	static BOOL WStr2CodePage(const wchar_t* wbuff, int wlen, UINT outCodePage, CStringA* resultA, DWORD& error);
	//
	static BOOL WStr2UTF8(const wchar_t* wbuff, int wlen, SimpleString* resultA, DWORD& error);
	static BOOL WStr2UTF8(CString* strW, CStringA* resultA, DWORD& error);
	static BOOL WStr2Ansi(CString& strW, CStringA& resultA, DWORD& error);
	static BOOL WStr2Ascii(CString& strW, CStringA& resultA, DWORD& error);
	//
	static BOOL CodePage2WStr(const char *str, int strlen, UINT strCodePage, SimpleString *wstr, DWORD &error, DWORD dwFlags = 0);
	static BOOL CodePage2WStr(const char* str, int strlen, UINT strCodePage, CString* wstr, DWORD& error, DWORD dwFlags = 0);
	static BOOL CodePage2WStr(SimpleString* str, UINT strCodePage, SimpleString* wstr, DWORD& error, DWORD dwFlags = 0);
	static BOOL CodePage2WStr(CStringA* str, UINT strCodePage, SimpleString* wstr, DWORD& error, DWORD dwFlags = 0);
	static BOOL CodePage2WStr(CStringA* str, UINT strCodePage, CString* wstr, DWORD& error, DWORD dwFlags = 0);
	//
	static BOOL UTF82WStr(CStringA* str, CString* wstr, DWORD& error, DWORD dwFlags = 0);
	static BOOL Ansi2WStr(CStringA& strA, CString& strW, DWORD& error, DWORD dwFlags = 0);
	static BOOL Ascii2Wstr(CStringA& strA, CString& strW, DWORD& error, DWORD dwFlags = 0);
	static BOOL Str2WStr(CStringA& res, UINT CodePage, CString& m_strW, DWORD& error, DWORD dwFlags = 0);
	//
	// PageCode -> PageCode number
	static BOOL Str2CodePage(const char* str, int strlen, UINT inCodePage, UINT outCodePage, SimpleString* result, SimpleString* workBuff, DWORD& error);
	//
	static BOOL Str2UTF8(const char* str, int strlen, UINT inCodePage, SimpleString* result, SimpleString* workBuff, DWORD& error);
	static BOOL Str2UTF8(const char* str, int strlen, UINT inCodePage, CStringA& outstr, DWORD& error);
	//
	static BOOL Str2CodePage(SimpleString* str, UINT inCodePage, UINT outCodePage, SimpleString* result, SimpleString* workBuff, DWORD& error);
	static BOOL Str2UTF8(SimpleString* str, UINT strCodePage, SimpleString* result, SimpleString* workBuff, DWORD& error);
	static BOOL Str2Ansi(SimpleString* str, UINT strCodePage, SimpleString* result, SimpleString* workBuff, DWORD& error);

	static DWORD Str2Ansi(CStringA &str, UINT strCodePage, DWORD& error);
	//
	static CStringA DecodeString(CStringA &subj, CStringA &charset, UINT &charsetId, UINT toCharacterId, DWORD& error);
	static int DecodeString(CStringA& inStr, CStringA &outStr, CStringA& charset, UINT& charsetId, DWORD& error);
	static int EncodeString(CStringA& inStr, CStringA &outStr, UINT charsetId, DWORD& error);

	static UINT StrPageCodeName2PageCode(const  char* PageCodeStr);
	//
	// End
	//
	static int hextob(char ch);
	static int DecodeMimeChunkedString(CStringA &subj, CStringA &charset, UINT &charsetId, BOOL hasCharset, CStringA &outString);

	static void ReplaceNL2CRNL(const char *in, int inLength, SimpleString *out);
	static void EncodeAsHtml(const char *in, int inLength, SimpleString *out);
	static void EncodeAsHtmlText(const char *in, int inLength, SimpleString *out);
	//
	static void delete_charset2Id();
	static UINT charset2Id(const char *char_set);
	static void delete_id2charset();
	static BOOL id2name(UINT id, CString& charsetName);
	static BOOL id2charset(UINT id, std::string &charset);
	static BOOL Id2LongInfoA(UINT id, CStringA& codePageInfo);
	static BOOL Id2LongInfo(UINT id, CString& codePageInfo);
	static BOOL GetCodePageInfo(UINT codePage, CP2NM &cpInfo);
	static int  showCodePageTable(CString &path);
	//
	static BOOL  isNumeric(CString &str);
	static BOOL  isNumericA(CStringA& str);
	static char *strchar(char *beg, char *end, char c);
	static char *findOneOf(char *beg, char *end, char *charList);
	static char *rfindOneOf(char* beg, char* end, char* charList);
	static char* findLastOneOf(char* beg, char* end, char* charList);
	static char* rfindLastOneOf(char* beg, char* end, char* charLis);
	static char *strnstrUpper2Lower(char *any, char *end, const char *lower, int lowerlength);
	static int	 strncmpUpper2Lower(char *any, char *end, const char *lower, int lowerlength);
	static int   strncmpUpper2Lower(char *any, int anyLength, const char *lower, int lowerlength);
	static int   strncmpExact(char *any, char *end, const char *lower, int lowerlength);
	static int   findNoCase(const char *input, int count, void const* Str, int  Size);
	static char *findNoCaseP(const char *input, int count, void const* Str, int  Size);
	static BOOL  isWhiteLine(const wchar_t* p, const wchar_t* e);
	static BOOL  isWhiteLine(CString &str);
	static BOOL  isWhiteLineA(const char* p, const char* e);
	static BOOL  isWhiteLineA(CStringA& str);
	//
	static int DecodeURL(char *URL, int urlLen);

	static void WStr2WstrWithCommas(const wchar_t* wstr, int wstrlen, CString &wstrout);
	static void Int2WstrWithCommas(INT64 numb, CString& wstrout);

	// based on MimeCode.cpp
	static int SelectEncoding(int nLength, int nNonAsciiChars);
	//
	static int WordEncode(CStringA& txt, CStringA& encodedTxt, int encodeType);

	static void CopyLine(char *p, char *e, CStringA &line)
	{
		char c;
		while (p < e)
		{
			c = *p;
			if ((c != '\r') && (c != '\n'))
			{
				line.AppendChar(c);
				p++;
			}
			else
				break;
		}
		if (line.GetLength())
			line.Append("\r\n");
		return;
	}

	static void CopyUpToEndOfLine(char* p, char* e, CStringA& line)
	{
		char c;
		while (p < e)
		{
			c = *p;
			if ((c != '\r') && (c != '\n'))
			{
				line.AppendChar(c);
				p++;
			}
			else
				break;
		}
		return;
	}

	inline static char *SkipNumeric(char *p) {
		while (isdigit(*p)) p++;  // TODO: may trigger exception if p become invalid; unlikely -:)
		return p;
	}

	inline static char *SkipNumericReverse(char *p) {
		while (isdigit(*p)) p--;  // TODO: may trigger exception if p become invalid; unlikely -:)
		return p;
	}

	inline static char *SkipWhiteReverse(char *p) {
		while ((*p == ' ') || (*p == '\t')) p--;  // TODO: may trigger exception if p become invalid; unlikely -:)
		return p;
	}

	inline static char *SkipWhite(char *p) {
		while ((*p == ' ') || (*p == '\t')) p++;
		return p;
	}
	inline static char *SkipWhitePlus(char *p) {
		while ((*p == ' ') || (*p == '\t') || (*p == '"')) p++;
		return p;
	}

	inline static char* SkipNonWhite(char* p) {
		while ((*p != ' ') && (*p != '\t')) p++;
		return p;
	}

	inline static char* SkipEOL(char* p)
	{
		while ((*p != '\n') && (*p != '\r')) p++;
		return p;

	}

	static void SplitStringA(const CStringA &strIn, const CStringA &delim, CStringArray &a);
	static void SplitStringA2A(const CStringA& strIn, const CStringA& delim, CStringArrayA& a);
	static int TokenizeA(CStringA& str, CStringArrayA& a, char del);
	static void TraceStringArrayA(CStringArrayA &a);
	//
	static void SplitStringW(const CString& strIn, const CString& delim, CStringArray& a);
	static int TokenizeW(CString &str, CStringArray& a, wchar_t del);
	static void TraceStringArrayW(CStringArray& a);

	static int XTokenString(LPCTSTR lpszString,
		LPCTSTR lpszDelimiters,
		CStringArray& saTokens,
		int nMaxTokens = 0,
		BOOL bTrimToken = FALSE,
		BOOL bEnableEscapedChars = FALSE,
		BOOL bEnableDoubleQuote = FALSE,
		BOOL bReturnEmptyToken = FALSE);

	static int TextUtilsEx::ReplaceChar(char* str, char c_target, char c_replace);
};


