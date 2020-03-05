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

class TextUtilsEx
{
public:
	//
	// Begin: String conversion functions nightmare -:) FIXME !!
	//
	static BOOL Wide2Ansi(CStringW &strW, CString &strA, DWORD &error);
	static BOOL Ansi2Wide(CString &strA, CStringW &strW, DWORD &error);
	//
	// PageCodeStr -> PageCode number
	static UINT Str2PageCode(const  char* PageCodeStr);
	// Str2Ansi converts str encoded with strCodePage -> CP_ACP i.e. to system default Windows ANSI code page, can be different on different computers.
	static DWORD Str2Ansi(CString &str, UINT strCodePage);  
	// Str2CodePage converts str encoded using strCodePage -> outCodePage or to CP_UTF8 if outCodePage == 0
	static BOOL Str2CodePage(SimpleString *str, UINT inCodePage, UINT outCodePage, SimpleString *result, SimpleString *workBuff);
	// UTF16 -> outCodePage or to UTF8 if outCodePage == 0
	static BOOL WStr2CodePage(wchar_t *wbuff, int wlen, UINT outCodePage, SimpleString *result, DWORD &error);
	static BOOL WStr2CodePage(wchar_t *wbuff, int wlen, UINT outCodePage, CString *result, DWORD &error);
	// strCodePage -> UTF16
	static BOOL CodePage2WStr(SimpleString *str, UINT strCodePage, SimpleString *wstr, DWORD &error);
	//
	static BOOL CodePage2WStr(CString *str, UINT strCodePage, SimpleString *wstr, DWORD &error);
	static BOOL CodePage2WStr(CString *str, UINT strCodePage, CStringW *wstr, DWORD &error);
	// strCodePage -> CP_UTF8
	static BOOL Str2UTF8(SimpleString *str, UINT strCodePage, SimpleString *result, SimpleString *workBuff);
	//  strCodePage -> CP_ACP
	static BOOL Str2CurrentCodepage(SimpleString *str, UINT strCodePage, SimpleString *result, SimpleString *workBuff);
	//
	static CString DecodeString(CString &subj, CString &charset, UINT &charsetId, UINT toCharacterId = 0);
	//
	static BOOL Str2Wide(CString &res, UINT CodePage, CStringW &m_strW);
	//
	// End
	//
	static void ReplaceNL2CRNL(const char *in, int inLength, SimpleString *out);
	static void EncodeAsHtml(const char *in, int inLength, SimpleString *out);
	static void delete_charset2Id();
	static UINT charset2Id(const char *char_set);
	static void delete_id2charset();
	static BOOL id2charset(UINT id, std::string &charset);
	static int showCodePageTable(CString &path);
	static BOOL isNumeric(CString &str);
	static char *strchar(char *beg, char *end, char c);
	static char *findOneOf(char *beg, char *end, char *charList);
	static char *strnstrUpper2Lower(char *any, char *end, const char *lower, int lowerlength);
	static int strncmpUpper2Lower(char *any, char *end, const char *lower, int lowerlength);
	static int strncmpUpper2Lower(char *any, int anyLength, const char *lower, int lowerlength);
	static int strncmpExact(char *any, char *end, const char *lower, int lowerlength);
	static int findNoCase(const char *input, int count, void const* Str, int  Size);
	//
	static int DecodeURL(char *URL, int urlLen);

	inline static char *SkipWhite(char *p) {
		while ((*p == ' ') || (*p == '\t')) p++;
		return p;
	}
	inline static char *SkipWhitePlus(char *p) {
		while ((*p == ' ') || (*p == '\t') || (*p == '"')) p++;
		return p;
	}
};

typedef struct {
	UINT m_charsetId;
	char *m_charset;
	char *m_info;
} CP2NM;
