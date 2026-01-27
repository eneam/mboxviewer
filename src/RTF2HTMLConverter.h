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

#pragma once

#include <iostream>
#include <unordered_map>
#include <map>
#include <list>
#include <string>
#include <regex>


/**
 * The last and most comprehensive converter that follows the RTF RFC and produces the most correct outcome.
 * <p>
 * <strong>Note:</strong> unlike {@link RTF2HTMLConverterClassic}, this converter doesn't wrap the result in
 * basic HTML tags if they're not already present in the RTF source.
 * <p>
 * The resulting source and rendered result is on par with software such as Outlook.
 */

//#define WINDOWS_CHARSET WINDOWS_1252
#define WINDOWS_CHARSET 1252

typedef std::string String;
typedef int Integer;
typedef int Charset;

enum {
};

class RegexMacher
{
public:
    RegexMacher(std::string& pattern, std::string& input)
    {
        m_match = 0;
        m_pattern = pattern;
        m_input = input;
        int deb = 1;
    };
    ~RegexMacher()
    {
        delete m_match;
    }

    void region(int start, int end)
    {
        m_start = start;
        m_end = end;
    };

    bool lookingAt();
    String RegexMacher::group(int number);
    int RegexMacher::start();
    int RegexMacher::end();
    int RegexMacher::GetSize();

    std::string GetMatch() {
        return all_groups;
    };

    int m_start;
    int m_end;

    std::regex m_pattern;
    std::string m_input;
    std::smatch* m_match;
    std::string all_groups;
};

class CharsetHelper
{
public:
    Charset detectCharsetFromRtfContent(std::string rtf) { return WINDOWS_CHARSET; }
};

class Group
{
public:
    Group()
    {
        ignore = false;
        unicodeCharLength = 1;
        htmlRtf = false;
        fontTableIndex = 0;
        id = next_id++;
    }

    Group(const Group &group)
    {
        if (group.ignore == false)
            int deb = 1;
        ignore = group.ignore;
        unicodeCharLength = group.unicodeCharLength;
        htmlRtf = group.htmlRtf;
        // Don't inherit fontTableIndex from parent group.   ??? only in one case see below
        fontTableIndex = group.fontTableIndex;
        id = next_id++;
    }

    bool ignore;
    int unicodeCharLength;
    bool htmlRtf;
    Integer fontTableIndex;
    int id;
static int next_id;
};

class FontTableEntry
{
public:
    Charset charset = WINDOWS_CHARSET;
};

typedef std::map<Integer, FontTableEntry> FontMap;

class RTF2HTMLConverter
{
public:
    RTF2HTMLConverter();

    CharsetHelper m_CharsetHelper;
    Charset m_ansicpg;
    bool m_ansicpg_setUTF8;  // set to UTF8 when done conversion
    std::string m_txt7bit;  // ASCII text fragment; characters  0 -127;
    bool m_alwaysEncodeASCII2UTF8;   // can we assume all charsets have Ascii ??

    std::string rtf2html(char* crtf, std::string& result);

    void appendIfNotIgnoredGroup(std::string& result, std::string& symbol, Group& group, bool char7bit=false);
    bool hexToString(const std::string& hex, Charset charset, std::string& result);
    void AppendNL(std::list<Group> & groupStack, std::string& result, Group& currentGroup);
    //
    bool utf16TOutf8(wchar_t inchar, std::string& outstr, DWORD& error);
    //
    bool ansipgcTOutf8(const char* instr, int inlen, int ansipgc, std::string& outstr_utf8, DWORD& error);
    bool ansipgcTOutf8(std::string& instr, int ansipgc, std::string& outstr_utf8, DWORD& error);
};

typedef struct
{
    int charsetNumber;
    int pageCode;
    char* name;
} RTFCharset;

int RTFCharsetNumber2CodePage(char* charsetNumberStr);
int RTFCharsetNumber2CodePage(int charsetNumber);