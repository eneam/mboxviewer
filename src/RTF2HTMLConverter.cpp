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
// The RTF2HTMLConverterRFCCompliant.java was poerted to c++ by Zbigniew Minciel to integrate with free Windows MBox Viewer
//
// https://github.com/bbottema/rtf-to-html/blob/master/src/main/java/org/bbottema/rtftohtml/impl/RTF2HTMLConverterRFCCompliant.java
//  Apache License
//  Version 2.0, January 2004
//  http://www.apache.org/licenses/
//


// Fix vars definitions
#pragma warning(disable : 4267)
#pragma warning(disable : 4244)

#include "ReportMemoryLeaks.h"

#include "stdafx.h"
#include <iostream>
#include <unordered_map>
#include <map>
#include <list>
#include <string>
#include <regex>
#include "TextUtilsEx.h"

#include "RTF2HTMLConverter.h"

std::string UTF16ToUTF8(wchar_t* in, size_t in_len);

static std::string CONTROL_WORD_PATTERN(R"(\\(([^a-zA-Z])|(([a-zA-Z]+)(-?\d*) ?)))");
static std::string ENCODED_CHARACTER_PATTERN(R"(^\\'([0-9a-fA-F]{2}))");


/**
 * The last and most comprehensive converter that follows the RTF RFC and produces the most correct outcome.
 * <p>
 * <strong>Note:</strong> unlike {@link RTF2HTMLConverterClassic}, this converter doesn't wrap the result in
 * basic HTML tags if they're not already present in the RTF source.
 * <p>
 * The resulting source and rendered result is on par with software such as Outlook.
 */


#if 0
 // from wingdi.h
#define ANSI_CHARSET            0
#define DEFAULT_CHARSET         1
#define SYMBOL_CHARSET          2
#define SHIFTJIS_CHARSET        128
#define HANGEUL_CHARSET         129
#define HANGUL_CHARSET          129
#define GB2312_CHARSET          134
#define CHINESEBIG5_CHARSET     136
#define OEM_CHARSET             255
#if(WINVER >= 0x0400)
#define JOHAB_CHARSET           130
#define HEBREW_CHARSET          177
#define ARABIC_CHARSET          178
#define GREEK_CHARSET           161
#define TURKISH_CHARSET         162
#define VIETNAMESE_CHARSET      163
#define THAI_CHARSET            222
#define EASTEUROPE_CHARSET      238
#define RUSSIAN_CHARSET         204

#define MAC_CHARSET             77
#define BALTIC_CHARSET          186
#endif
#endif

//RTF charset number , codepage , Windows/Mac name
static RTFCharset rtfCharsetArr[] =
{
	{ 0 , 1252 , "ANSI" },
	{ 1 , 0 , "Default" },
	{ 2 , 42 , "Symbol" },
	{ 77 , 10000 , "Mac Roman" },
	{ 78 , 10001 , "Mac Shift Jis" },
	{ 79 , 10003 , "Mac Hangul" },
	{ 80 , 10008 , "Mac GB2312" },
	{ 81 , 10002 , "Mac Big5" },
	{ 82 , -1 , "Mac Johab(old)" },
	{ 83 , 10005 , "Mac Hebrew" },
	{ 84 , 10004 , "Mac Arabic" },
	{ 85 , 10006 , "Mac Greek" },
	{ 86 , 10081 , "Mac Turkish" },
	{ 87 , 10021 , "Mac Thai" },
	{ 88 , 10029 , "Mac East Europe" },
	{ 89 , 10007 , "Mac Russian" },
	{ 128 , 932 , "Shift JIS" },
	{ 129 , 949 , "Hangul" },
	{ 130 , 1361 , "Johab" },
	{ 134 , 936 , "GB2312" },
	{ 136 , 950 , "Big5" },
	{ 161 , 1253 , "Greek" },
	{ 162 , 1254 , "Turkish" },
	{ 163 , 1258 , "Vietnamese" },
	{ 177 , 1255 , "Hebrew" },
	{ 178 , 1256 , "Arabic" },
	{ 179 , -1 , "arabic Traditional(old)" },
	{ 180 , -1 , "arabic user(old)" },
	{ 181 , -1 , "Hebrew user(old)" },
	{ 186 , 1257 , "Baltic" },
	{ 204 , 1251 , "Russian" },
	{ 222 , 874 , "Thai" },
	{ 238 , 1250 , "Eastern European" },
	{ 254 , 437 , "PC 437" },
	{ 255 , 850 , "OEM" }
};

int RTFCharsetNumber2CodePage(char* charsetNumberStr)
{
	int charsetNumber = -1;
	sscanf(charsetNumberStr, "%d", &charsetNumber);
	return charsetNumber;
}

int RTFCharsetNumber2CodePage(int charsetNumber)
{
	int pageCode = -1;
	int arraySize = sizeof(rtfCharsetArr) / sizeof(RTFCharset);
	int i = 0;
	for (i = 0; i < arraySize; i++)
	{
		RTFCharset* charSet = &rtfCharsetArr[i];
		if (charSet->charsetNumber == charsetNumber)
		{
			return charSet->pageCode;
		}
	}

	return -1;
}

bool RTF2HTMLConverter::utf16TOutf8(wchar_t inchar, std::string& outstr_utf8, DWORD& error)
{
	CStringA outstr;
	outstr_utf8.clear();

	CString wStr(inchar);
	BOOL ret = TextUtilsEx::WStr2UTF8(&wStr, &outstr, error);
	if (ret) {
		outstr_utf8.assign((LPCSTR)outstr, outstr.GetLength());
	}
	else
		_ASSERTE(true);

	return ret;
}

bool RTF2HTMLConverter::ansipgcTOutf8(std::string& instr, int ansipgc, std::string& outstr_utf8, DWORD& error)
{
	const char* in = instr.c_str();
	int inlen = instr.length();
	bool ret = RTF2HTMLConverter::ansipgcTOutf8(in, inlen, ansipgc, outstr_utf8, error);
	return ret;
}

bool RTF2HTMLConverter::ansipgcTOutf8(const char* in, int inlen, int ansipgc, std::string& outstr_utf8, DWORD& error)
{
	CStringA outstr;

	BOOL ret = TextUtilsEx::Str2UTF8(in, inlen, ansipgc, outstr, error);
	if (ret) {
		outstr_utf8.assign((LPCSTR)outstr, outstr.GetLength());
	}
	else
		_ASSERTE(true);

	return ret;
}

RTF2HTMLConverter::RTF2HTMLConverter()
{
	m_ansicpg = 0;  // windows ansi , 1252
	m_ansicpg_setUTF8 = false;
	m_alwaysEncodeASCII2UTF8 = false;
};

std::string RTF2HTMLConverter::rtf2html(char* crtf, std::string& result)
{
	result.clear();
	FontMap fontTable;
	std::string rtf(crtf);
	result.reserve(rtf.length());
	m_txt7bit.reserve(rtf.length());

	// Determine charset from content. Normally asnsicpg should be present and it will reset
	// what was found. For not set to 1252 windows ansi
	Charset charset = m_CharsetHelper.detectCharsetFromRtfContent(rtf);

	// RTF processing requires stack holding current settings, each group adds new settings to stack
	// '{' creates Group, '}' deletes Group
	std::list<Group> groupStack;
	Group new_group;
	groupStack.push_back(new_group);

	RegexMacher controlWordMatcher(CONTROL_WORD_PATTERN, rtf);
	RegexMacher encodedCharMatcher(ENCODED_CHARACTER_PATTERN, rtf);
	//std::string result;
	int length = rtf.length();
	int charIndex = 0;
	Integer nullControlNumber = 0x7FFFFFFF;

	while (charIndex < length)
	{
		std::string::const_iterator it = rtf.begin() + charIndex;
		char cc = *it;

		char c = rtf.at(charIndex);
		Group& currentGroup = groupStack.front();
		if (c == '\r' || c == '\n') {
			charIndex++;
		}
		else if (c == '{') {  //entering group
			Group new_group(currentGroup);
			new_group.fontTableIndex = 0;  // don't inherit fontTableIndex
			groupStack.push_front(new_group);
			charIndex++;
		}
		else if (c == '}') {  //exiting group
			groupStack.pop_front();
			//Not outputting anything after last closing brace matching opening brace.
			if (groupStack.size() == 1) {
				break;
			}
			charIndex++;
		}
		else if (c == '\\')
		{
			// matching ansi-encoded sequences like \'f5\'93
			int end_index = charIndex + 4;
			encodedCharMatcher.region(charIndex, end_index);  // charIndex == start index, length == end index
			if (encodedCharMatcher.lookingAt())
			{
				std::string hexEncodedSequence;
				while (encodedCharMatcher.lookingAt())
				{
					String group_1 = encodedCharMatcher.group(1);
					hexEncodedSequence.append(group_1);
					charIndex += 4;
					end_index = charIndex + 4;
					encodedCharMatcher.region(charIndex, end_index);
					if (charIndex >= length)
						int deb = 1;
				}

				// restore region
				encodedCharMatcher.region(charIndex, length);

				Charset effectiveCharset = charset;
				if (currentGroup.fontTableIndex != 0)
				{
					FontMap::iterator entry = fontTable.find(currentGroup.fontTableIndex);
					if (entry != fontTable.end() && entry->second.charset != 0) {
						effectiveCharset = entry->second.charset;
					}
				}

				String decoded;
				bool retEncodedUTF8 = RTF2HTMLConverter::hexToString(hexEncodedSequence.c_str(), effectiveCharset, decoded);
				if (!decoded.empty())
					int deb = 1;
				else
					int deb = 1;

				appendIfNotIgnoredGroup(result, decoded, currentGroup, retEncodedUTF8);
				m_ansicpg_setUTF8 = retEncodedUTF8;  // update  m_ansicpg = CP_UTF8 before return from function
				continue;
			}

			// set matcher to current char position and match from it
			controlWordMatcher.region(charIndex, length);
			if (!controlWordMatcher.lookingAt())
			{
				std::string err("RTF file has invalid structure. Failed to match character '" +
					String(&c, 1) + "' at [" + std::to_string(charIndex) + "/" + std::to_string(length) + "] to a control symbol or word.");
				return String(err);  // Indicate failure
			}

			std::string allGropus = controlWordMatcher.GetMatch();  // testing
			int indx = allGropus.find("par");
			if (indx >= 0)
				int deb = 1;

			//checking for control symbol or control word
			//control word can have optional number following it and the optional space as well
			Integer controlNumber = nullControlNumber;
			String controlWord = controlWordMatcher.group(2); // group(2) matches control symbol; !!!! changed from group(2) to group(4) need explanation
			if (controlWord.empty())  // ZMM
			{
				controlWord = controlWordMatcher.group(4); // group(4) matches control word
				String controlNumberString = controlWordMatcher.group(5);
				if (!controlNumberString.empty())
				{
					controlNumber = strtol(controlNumberString.c_str(), 0, 10);
					if (controlNumber < 0)
						controlNumber += 65536;
				}
			}
			int oldIndex = charIndex;
			charIndex += controlWordMatcher.end() - controlWordMatcher.start();

			if (!controlWord.empty())
				int deb = 1;

			if (controlWord == "htmlrtf")
			{
				if (controlNumber > 0)
					int deb = 1;
				else
					int deb = 1;
			}

			//switch (controlWord)  // For now rely on == later on switch once all works !!!
			if (controlWord == "htmlrtf")
			{
				//htmlrtf starts ignored text area, htmlrtf0 ends it
				//Though technically this is not a group, it's easier to treat it as such to ignore everything in between
				if ((controlNumber == nullControlNumber) || (controlNumber == 1))  // nullControlNumber = 0x7FFFFFFF  to indicate ull/invalid number may need to FIX this
					currentGroup.htmlRtf = 1;
				else
					currentGroup.htmlRtf = 0;
			}
			else if (controlWord == "pard")
				appendIfNotIgnoredGroup(result, String("\n"), currentGroup, false);
			else if (controlWord == "par")
				appendIfNotIgnoredGroup(result, String("\n"), currentGroup, false);
			else if (controlWord == "line")
				appendIfNotIgnoredGroup(result, String("\n"), currentGroup, false);
			else if (controlWord == "tab")
				appendIfNotIgnoredGroup(result, String("\t"), currentGroup, false);
			else if (controlWord == "ansicpg")
			{
				//charset definition is important for decoding ansi encoded values
				//charset = CharsetHelper.findCharsetForCodePage(requireNonNull(controlNumber).toString());
				if (controlNumber >= 0)
				{
					charset = controlNumber;
					if (m_ansicpg == 0)  // allow only first "ansicpg" for now
						m_ansicpg = controlNumber;
				}
			}
			else if (controlWord == "fonttbl") // skipping these groups' contents - these are font and color settings
				currentGroup.ignore = true;   // !!!! ignore will remain set until the first/oldest group with ignore set is deleted
			else if (controlWord == "colortbl")
				currentGroup.ignore = true;  // !!!! ignore will remain set until the first/oldest group with ignore set is deleted
			else if (controlWord == "f") {
				// font table index. Might be a new one, or an existing one
				_ASSERTE(controlNumber >= 0);
				if (controlNumber != nullControlNumber)
					currentGroup.fontTableIndex = controlNumber;
			}
			else if (controlWord == "fcharset")
			{
				if (controlNumber != nullControlNumber && currentGroup.fontTableIndex != 0)
				{
					//Charset possibleCharset = CodePage.getCharsetByCodePage(controlNumber);
					Charset possibleCharset = RTFCharsetNumber2CodePage(controlNumber);
					if (possibleCharset != 0)
					{
						FontMap::iterator iter_entry = fontTable.find(currentGroup.fontTableIndex);
						if (iter_entry == fontTable.end())
						{
							FontTableEntry new_entry;
							fontTable.insert({ currentGroup.fontTableIndex, new_entry });
							FontMap::iterator iter = fontTable.find(currentGroup.fontTableIndex);
							iter->second.charset = possibleCharset;
						}
						else
							iter_entry->second.charset = possibleCharset;
					}
				}
			}
			else if (controlWord == "uc")
			{
				// This denotes a number of characters to skip after unicode symbols
				if (controlNumber != nullControlNumber)
					currentGroup.unicodeCharLength = (controlNumber == 0) ? 1 : controlNumber;
			}
			else if (controlWord == "u")
			{
				// Unicode symbols, ZMM RTF doc is confusing; need to investigate more; example msg files would help
#if 0
				// ZMM is this applicable to Mail also? How she know ?
				// Should we check for if Unocde value is between 
				// I don't have enpugh examples
				// From very confusing RTF doc:

				Occasionally Word writes SYMBOL_CHARSET(nonUnicode) characters in the range
				U + F020..U + F0FF instead of U + 0020..U + 00FF.
				Internally Word uses the values U + F020..U + F0FF
				for these characters so that plain - text searches don’t mistakenly match SYMBOL_CHARSET
				characters when searching for Unicode characters in the range U + 0020..U + 00FF.To find out the
				correct symbol font to use, e.g., Wingdings, Symbol, etc., find the last SYMBOL_CHARSET font
				control word \fN used, look up font N in the font table and find the face name.The charset is
				specified by the \fcharsetN control word and SYMBOL_CHARSET is for N = 2. This corresponds
				to codepage 42.
#endif
				if (controlNumber != nullControlNumber)
				{
#if 0
					// Direct port from Java source. Howver, In Java the char type is 2 bytes most of the time
					char unicodeSymbol = (char)controlNumber;  // 
					appendIfNotIgnoredGroup(result, String(&unicodeSymbol, 1), currentGroup);
#endif
					// ZMM Investigate if not Symbol , private Microsoft Unicode range ??
					// 
					wchar_t unicodeSymbol = (wchar_t)controlNumber;
					CString unicodeSymbolStr(unicodeSymbol);

					if ((unicodeSymbol >= 0xF020) && (unicodeSymbol <= 0xF0FF))
						int deb = 1;
					else
						int deb = 1;

					DWORD error;
					std::string str_utf8("?");

					// we must encode alwasy
					bool ret = utf16TOutf8(unicodeSymbol, str_utf8, error);
					if (ret)
					{
						result.append(str_utf8);
						appendIfNotIgnoredGroup(result, str_utf8, currentGroup, true);

						//m_ansicpg_setUTF8 = true;  // update  m_ansicpg = CP_UTF8 before return from function
					}
					else
					{
						_ASSERTE(true);
						str_utf8.assign("?");
						appendIfNotIgnoredGroup(result, str_utf8, currentGroup, false);  // ZMM what would be correct handling ??
					}

					charIndex += currentGroup.unicodeCharLength;
				}
			}
			else if ((controlWord == "{") || (controlWord == "}") || (controlWord == "\\"))  // Escaped characters
				appendIfNotIgnoredGroup(result, controlWord, currentGroup, false);  // bool isUTF8EncodedSymbols = false
			else if (controlWord == "pntext")
				currentGroup.ignore = true;  // !!!! ignore will remain set until the first/oldest group with ignore set is deleted
			else
				int deb = 1;
		}
		else {
			appendIfNotIgnoredGroup(result, String(&c, 1), currentGroup, false);
			charIndex++;
		}
	}

	if (!m_txt7bit.empty())
	{
		const char* str = m_txt7bit.c_str();

		if (m_alwaysEncodeASCII2UTF8 || m_ansicpg_setUTF8)
		{
			DWORD error;
			std::string str_utf8;
			bool ret = RTF2HTMLConverter::ansipgcTOutf8(m_txt7bit, m_ansicpg, str_utf8, error);
			if (ret) {
				m_ansicpg_setUTF8 = true;
				result.append(str_utf8);
			}
			else
			{
				_ASSERTE(true);
				result.append(m_txt7bit);
			}
		}
		else
			result.append(m_txt7bit);

		m_txt7bit.clear();
	}

	if (m_ansicpg_setUTF8)
		m_ansicpg = CP_UTF8;

	return String("");  // Indicate ok;
};

void RTF2HTMLConverter::appendIfNotIgnoredGroup(std::string& result, std::string& symbol, Group& group, bool isUTF8EncodedSymbols)
{
	if (isUTF8EncodedSymbols)
		int deb = 1;

	if (!group.ignore && !group.htmlRtf)
	{
		if (symbol.length() > 4)
			int deb = 1;

		if (!isUTF8EncodedSymbols) {
			m_txt7bit.append(symbol);  // it will appended at some point and may have to be UTF8 encoded.
		}
		else if (!m_txt7bit.empty())  // ASCII text ??
		{
			DWORD error;
			std::string str_utf8;
			bool ret = RTF2HTMLConverter::ansipgcTOutf8(m_txt7bit, m_ansicpg, str_utf8, error);
			if (ret) {
				result.append(str_utf8);
				m_ansicpg_setUTF8 = true;
			}
			else
			{
				_ASSERTE(true);
				result.append(m_txt7bit);
			}

			m_txt7bit.clear();
			result.append(symbol);
		}
		else
			result.append(symbol);

		int deb = 1;
	}
}

bool RTF2HTMLConverter::hexToString(const std::string& hex, Charset charsetNumber, std::string& result)
{
	bool isUTF8EncodedSymbols = false;

	result.clear();
	for (size_t i = 0; i < hex.length(); i += 2)
	{
		std::string _byte = hex.substr(i, 2);
		char chr = (char)(int)strtol(_byte.c_str(), nullptr, 16);
		result.push_back(chr);
		int deb = 1;
	}
	_ASSERT(!result.empty());

	int inCodePage = RTFCharsetNumber2CodePage(charsetNumber);
	if ((charsetNumber == m_ansicpg) || (inCodePage == m_ansicpg))
	{
		int deb = 1;
	}
	else {
		int deb1 = 1;
	}

	if (result.size() > 0)
	{
		DWORD error;
		if (inCodePage < 0)
			inCodePage = charsetNumber;

		if (m_alwaysEncodeASCII2UTF8 || (inCodePage != m_ansicpg) || m_ansicpg_setUTF8)
		{
			std::string str_utf8;
			bool ret = RTF2HTMLConverter::ansipgcTOutf8(result, inCodePage, str_utf8, error);
			if (ret) {
				isUTF8EncodedSymbols = true;
				result = str_utf8;
			}
			else
			{
				_ASSERTE(true);
			}
		}
		else
		{
			int deb = 1;
		}
	}
	return isUTF8EncodedSymbols;
}

bool RegexMacher::lookingAt()
{
	if (m_match != 0) {
		delete m_match;
		m_match = 0;
	}
	if (m_match == 0)
		m_match = new std::smatch;


	std::string::const_iterator b = m_input.begin() + m_start;
	std::string::const_iterator e = m_input.begin() + m_end;

	bool match = std::regex_search(b, e, *m_match, m_pattern, std::regex_constants::match_default);

	all_groups.clear();
	for (unsigned i = 0; i < m_match->size(); ++i)
	{
		std::string grp = m_match->str(i);

		if (i != 0)
			all_groups.append(" \"");
		else
			all_groups.append("\"");
		all_groups.append(grp);
		all_groups.append("\"");
		int deb = 1;
	}
	return match;
};

String RegexMacher::group(int number)
{
	std::string str_group = m_match->str(number);
	return str_group;
};

int RegexMacher::start()
{
	if (m_match->size() > 1)
	{
		int start = m_match->position(0);
		return start;
	}
	else
		return 0; // -1 ??  ZMM
};

int RegexMacher::end()
{
	if (m_match->size() > 1)
	{
		int end = m_match->position(0) + m_match->length(0);
		return end;
	}
	else
		return 0; // -1 ??  ZMM
};

int RegexMacher::GetSize() {
	return m_match->size();
}


