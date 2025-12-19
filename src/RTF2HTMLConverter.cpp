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

#include "RTF2HTMLConverter.h"

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


RTF2HTMLConverter::RTF2HTMLConverter()
{
	m_ansicpg = 0;  //invalid
};

std::string RTF2HTMLConverter::rtf2html(char* crtf)
{
	FontMap fontTable;
	std::string rtf(crtf);

	Charset charset = m_CharsetHelper.detectCharsetFromRtfContent(rtf);

	// RTF processing requires stack holding current settings, each group adds new settings to stack
	// '{' creates Group, '}' deletes Group
	std::list<Group> groupStack;
	Group new_group;
	groupStack.push_back(new_group);

	RegexMacher controlWordMatcher(CONTROL_WORD_PATTERN, rtf);
	RegexMacher encodedCharMatcher(ENCODED_CHARACTER_PATTERN, rtf);
	std::string result;
	int length = rtf.length();
	int charIndex = 0;

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
				std::string encodedSequence;
				while (encodedCharMatcher.lookingAt()) {
					String group_1 = encodedCharMatcher.group(1);
					encodedSequence.append(group_1);
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
					FontMap::iterator entry = fontTable.find(currentGroup.fontTableIndex);;
					if (entry != fontTable.end() && entry->second.charset != 0) {
						effectiveCharset = entry->second.charset;
					}
				}

				//String decoded = hexToString(encodedSequence.c_str(), effectiveCharset);
				String decoded = RTF2HTMLConverter::hexToString(encodedSequence.c_str());  // ignored effectiveCharset for now  !!!!
				appendIfNotIgnoredGroup(result, decoded, currentGroup);
				continue;
			}

			// set matcher to current char position and match from it
			controlWordMatcher.region(charIndex, length);
			if (!controlWordMatcher.lookingAt())
			{
				std::string err("RTF file has invalid structure. Failed to match character '" +
					String(&c, 1) + "' at [" + std::to_string(charIndex) + "/" + std::to_string(length) + "] to a control symbol or word.");
				return String("");  // Indicate failure
			}

			std::string all1 = controlWordMatcher.GetMatch();
			int indx = all1.find("par");
			if (indx >= 0)
				int deb = 1;

			//checking for control symbol or control word
			//control word can have optional number following it and the optional space as well
			Integer controlNumber = 0x7FFFFFFF;
			String controlWord = controlWordMatcher.group(4); // group(2) matches control symbol; !!!! changed from group(2) to group(4) need explanation
			if (!controlWord.empty())
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
			{
				if (controlWord == "htmlrtf") {
					//htmlrtf starts ignored text area, htmlrtf0 ends it
					//Though technically this is not a group, it's easier to treat it as such to ignore everything in between
					if ((controlNumber == 0x7FFFFFFF) || (controlNumber == 1))  // 0x7FFFFFFF  to indicate no number FIX this
						currentGroup.htmlRtf = 1;
					else
						currentGroup.htmlRtf = 0;
				}
				else if (controlWord == "pard")
					appendIfNotIgnoredGroup(result, String("\n"), currentGroup);
				else if (controlWord == "par")
					appendIfNotIgnoredGroup(result, String("\n"), currentGroup);
				else if (controlWord == "line")
					appendIfNotIgnoredGroup(result, String("\n"), currentGroup);
				else if (controlWord == "tab")
					appendIfNotIgnoredGroup(result, String("\t"), currentGroup);
				else if (controlWord == "ansicpg") {
					//charset definition is important for decoding ansi encoded values
					//charset = CharsetHelper.findCharsetForCodePage(requireNonNull(controlNumber).toString());
					if (controlNumber >= 0)
					{
						charset = controlNumber;
						if (m_ansicpg == 0)  // allow onlu first "ansicpg" for now
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
					if (controlNumber != 0x7FFFFFFF)
						currentGroup.fontTableIndex = controlNumber;
				}
				else if (controlWord == "fcharset")
				{
					if (controlNumber != 0x7FFFFFFF && currentGroup.fontTableIndex != 0)
					{
						//Charset possibleCharset = CodePage.getCharsetByCodePage(controlNumber);
						Charset possibleCharset = 0;
						if (possibleCharset != 0)
						{
							FontMap::iterator iter_entry = fontTable.find(currentGroup.fontTableIndex);
							if (iter_entry != fontTable.end()) {
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
				else if (controlWord == "uc") {
					// This denotes a number of characters to skip after unicode symbols
					if (controlNumber != 0x7FFFFFFF)
						currentGroup.unicodeCharLength = (controlNumber == 0) ? 1 : controlNumber;
				}
				else if (controlWord == "u")
				{
					// Unicode symbols
					if (controlNumber != 0x7FFFFFFF)
					{
						if (controlNumber != 0x7FFFFFFF) {
							char unicodeSymbol = (char)controlNumber;
							appendIfNotIgnoredGroup(result, String(&unicodeSymbol, 1), currentGroup);
							charIndex += currentGroup.unicodeCharLength;
						}
					}
				}
				else if ((controlWord == "{") || (controlWord == "}") || (controlWord == "\\"))  // Escaped characters
					appendIfNotIgnoredGroup(result, controlWord, currentGroup);
				else if (controlWord == "pntext")
					currentGroup.ignore = true;  // !!!! ignore will remain set until the first/oldest group with ignore set is deleted
				else
					int deb = 1;
			}

		}
		else {
			appendIfNotIgnoredGroup(result, String(&c, 1), currentGroup);
			charIndex++;
		}
	}
	return result;
};

std::string RTF2HTMLConverter::hexToString(const std::string& hex)
{
	std::string result;
	for (size_t i = 0; i < hex.length(); i += 2) {
		std::string _byte = hex.substr(i, 2);
		char chr = (char)(int)strtol(_byte.c_str(), nullptr, 16);
		result.push_back(chr);
		int deb = 1;
	}
	return result;
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
		return 0; // -1 ??
};

int RegexMacher::end()
{
	if (m_match->size() > 1)
	{
		int end = m_match->position(0) + m_match->length(0);
		return end;
	}
	else
		return 0; // -1 ??
};

int RegexMacher::GetSize() {
	return m_match->size();
}


