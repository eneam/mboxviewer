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
// The RTF2HTMLConverterRFCCompliant.java was ported to c# by Zbigniew Minciel to integrate with Winmail2Eml.exe and free Windows MBox Viewer
// c# version is a port of c++ version, part of MBox Viewer code base
//
// https://github.com/bbottema/rtf-to-html/blob/master/src/main/java/org/bbottema/rtftohtml/impl/RTF2HTMLConverterRFCCompliant.java
//  Apache License
//  Version 2.0, January 2004
//  http://www.apache.org/licenses/
//
// Note: the ported code works with available MSG files
// However, more complete implementation might be needed based on RTF specification
// Using regular expression is too limited


#pragma warning disable 219

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace RTF2HTMLConversion
{
	using Integer = int;
	using Charset = int;
	using DWORD = uint;

#if COMMENTS
/**
 * The last and most comprehensive converter that follows the RTF RFC and produces the most correct outcome.
 * <p>
 * <strong>Note:</strong> unlike {@link RTF2HTMLConverterClassic}, this converter doesn't wrap the result in
 * basic HTML tags if they're not already present in the RTF source.
 * <p>
 * The resulting source and rendered result is on par with software such as Outlook.
 */
#endif

	//RTF charset number , codepage , Windows/Mac name

	using FontMap = System.Collections.Generic.Dictionary<int, FontTableEntry>;
	using Org.BouncyCastle.Utilities.Encoders;
	using System.Linq;
	using static System.Net.Mime.MediaTypeNames;

	public class RTFCharsetArray
	{
		public static RTFCharset[] rtfCharsetArr = new RTFCharset[]
		{
			new( 0  , 1252  , "ANSI" ),
			new( 1  , 0     , "Default" ),
			new( 2  , 42    , "Symbol" ),
			new(77  , 10000 , "Mac Roman" ),
			new(78  , 10001 , "Mac Shift Jis" ),
			new(79  , 10003 , "Mac Hangul" ),
			new(80  , 10008 , "Mac GB2312" ),
			new(81  , 10002 , "Mac Big5" ),
			new(82  , -1    , "Mac Johab(old)" ),
			new(83  , 10005 , "Mac Hebrew" ),
			new(84  , 10004 , "Mac Arabic" ),
			new(85  , 10006 , "Mac Greek" ),
			new(86  , 10081 , "Mac Turkish" ),
			new(87  , 10021 , "Mac Thai" ),
			new(88  , 10029 , "Mac East Europe" ),
			new(89  , 10007 , "Mac Russian" ),
			new(128 , 932   , "Shift JIS" ),
			new(129 , 949   , "Hangul" ),
			new(130 , 1361  , "Johab" ),
			new(134 , 936   , "GB2312" ),
			new(136 , 950   , "Big5" ),
			new(161 , 1253  , "Greek" ),
			new(162 , 1254  , "Turkish" ),
			new(163 , 1258  , "Vietnamese" ),
			new(177 , 1255  , "Hebrew" ),
			new(178 , 1256  , "Arabic" ),
			new(179 , -1    , "arabic Traditional(old)" ),
			new(180 , -1    , "arabic user(old)" ),
			new(181 , -1    , "Hebrew user(old)" ),
			new(186 , 1257  , "Baltic" ),
			new(204 , 1251  , "Russian" ),
			new(222 , 874   , "Thai" ),
			new(238 , 1250  , "Eastern European" ),
			new(254 , 437   , "PC 437" ),
			new(255 , 850   , "OEM" )
		};
	}

	public class RTF2HTMLConverter
	{
		//public CharsetHelper m_CharsetHelper { get; set; }
		public Charset m_ansicpg { get; set; }
		public bool m_ansicpg_setUTF16 { get; set; }  // set to UTF8 when done conversion
		public StringBuilder m_txt7bit { get; set; }  // ANSI text fragment; characters 0-255;
		public bool m_alwaysEncodeANSI2UTF16 { get; set; }

		public ByteArray m_byteArray;

		public RTF2HTMLConverter()
		{
			m_ansicpg = 0;  // windows ansi , 1252
			m_ansicpg_setUTF16 = false;
			m_alwaysEncodeANSI2UTF16 = false;
			m_txt7bit = new();
			m_byteArray = new(16);
		}

		public int RTFCharsetNumber2CodePage(int charsetNumber)
		{
			int pageCode = -1;
			int arraySize = RTFCharsetArray.rtfCharsetArr.Length;
			int i = 0;
			for (i = 0; i < arraySize; i++)
			{
				RTFCharset charSet = RTFCharsetArray.rtfCharsetArr[i];
				if (charSet.charsetNumber == charsetNumber)
				{
					return charSet.pageCode;
				}
			}
			return pageCode;
		}

		public bool ansipgc2utf16(string from, int fromCodePage, ref string to, ref DWORD error)
		{
			bool ret = true;

			// Needed because we use StringBuilder
			// Use custom ByteArray instead
			byte[] fromBytes = new byte[from.Length];
			for (int i = 0; i < from.Length; i++)
			{
				fromBytes[i] = (byte)from[i];
			}


			Encoding fromEncoding = Encoding.GetEncoding(fromCodePage);
			Encoding toEncoding = Encoding.Unicode;


			string too = fromEncoding.GetString(fromBytes, 0, from.Length);

			// Convert the string into a byte array.

			// Perform the conversion from one encoding to the other.
			byte[] unicodeBytes = Encoding.Convert(fromEncoding, toEncoding, fromBytes);

			to = toEncoding.GetString(unicodeBytes, 0, unicodeBytes.Length);

			return ret;
		}

		public string rtf2html(ref string rtf, ref StringBuilder result)
		{
			int deb = 1;
			string strLF = "\n";
			string strTAB = "\t";
			result.Clear();

			FontMap fontTable = new();

			result = new(rtf.Count());
			result.Capacity = rtf.Count();
			m_txt7bit = new(rtf.Count());
			m_txt7bit.Capacity = rtf.Count();

			m_byteArray.Resize(rtf.Count());  // update to use m_byteArray instead of m_txt7bit

			// Determine charset from content. Normally asnsicpg should be present and it will reset
			// what was found. For not set to 1252 windows ansi
			Charset charset = CharsetHelper.detectCharsetFromRtfContent(rtf);

			// RTF processing requires stack holding current settings, each group adds new settings to stack
			// '{' creates Group, '}' deletes Group
			LinkedList<Group> groupStack = new();
			Group new_group = new();
			groupStack.AddLast(new_group);

			RegexMacher controlWordMatcher = new RegexMacher(RegexMacher.CONTROL_WORD_PATTERN, rtf);
			RegexMacher encodedCharMatcher = new RegexMacher(RegexMacher.ENCODED_CHARACTER_PATTERN, rtf);

			int length = rtf.Count();
			int charIndex = 0;
			Integer nullControlNumber = 0x7FFFFFFF;

			while (charIndex < length)
			{
				char c = rtf[charIndex];
				Group? currentGroup = groupStack.First.Value;
				Debug.Assert(currentGroup != null);
				if (c == '\r' || c == '\n')
				{
					charIndex++;
				}
				else if (c == '{')  //entering group
				{
					Group group = new(currentGroup);
					new_group.fontTableIndex = 0;  // don't inherit fontTableIndex
					groupStack.AddFirst(group);
					charIndex++;
				}
				else if (c == '}')  //exiting group
				{
					// Break very long single HTML line. Need better solution
					// Commented out. Below doesn't work; it breaks for example <pre blocks and pre-wrap style
					//AppendNL(groupStack, result, currentGroup);

					groupStack.RemoveFirst();
					//Not outputting anything after last closing brace matching opening brace.
					if (groupStack.Count() == 1)
					{
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
						string hexEncodedSequence = "";
						while (encodedCharMatcher.lookingAt())
						{
							String group_1 = encodedCharMatcher.group(1);
							hexEncodedSequence += group_1;
							charIndex += 4;
							end_index = charIndex + 4;
							encodedCharMatcher.region(charIndex, end_index);
							if (charIndex >= length)
								deb = 1;
						}

						// restore region
						encodedCharMatcher.region(charIndex, length);

						Charset effectiveCharset = charset;
						if (currentGroup.fontTableIndex != 0)
						{
							bool f = fontTable.ContainsKey(currentGroup.fontTableIndex);
							if (f)
							{
								FontTableEntry entry = fontTable[currentGroup.fontTableIndex];
								if (entry.charset != 0)
								{
									effectiveCharset = entry.charset;
								}
							}
						}

						StringBuilder decodedSB = new("");
						bool retEncodedUTF16 = hexToString(hexEncodedSequence, effectiveCharset, ref decodedSB);
						if (decodedSB.Length == 0)
							deb = 1;
						else
							deb = 1;

						string decoded = decodedSB.ToString();
						appendIfNotIgnoredGroup(result, decoded, currentGroup, retEncodedUTF16);
						if (!m_ansicpg_setUTF16)
							m_ansicpg_setUTF16 = retEncodedUTF16;
						continue;
					}

					// set matcher to current char position and match from it
					controlWordMatcher.region(charIndex, length);
					if (!controlWordMatcher.lookingAt())
					{
						string err = "RTF file has invalid structure. Failed to match character '" +
							c.ToString() + "' at [" + charIndex.ToString() + "/" + length.ToString() + "] to a control symbol or word.";
						return new String(err);  // Indicate failure
					}

					string allGroups = controlWordMatcher.GetMatch();  // testing
					int indx = allGroups.IndexOf(new string("par"));
					if (indx >= 0)
						deb = 1;

					//checking for control symbol or control word
					//control word can have optional number following it and the optional space as well
					Integer controlNumber = nullControlNumber;
					String controlWord = controlWordMatcher.group(2); // group(2) matches control symbol; !!!! changed from group(2) to group(4) need explanation
					if (controlWord.Count() == 0)  // ZMM
					{
						controlWord = controlWordMatcher.group(4); // group(4) matches control word
						String controlNumberString = controlWordMatcher.group(5);
						if (controlNumberString.Count() > 0)
						{
							bool success = int.TryParse(controlNumberString, out int number);
							if (success)
							{
								controlNumber = number;
								if (controlNumber < 0)
									controlNumber += 65536;
							}
							else
								Debug.Assert(success == false);
						}
					}
					int oldIndex = charIndex;
					int posEnd = controlWordMatcher.end();
					int posStart = controlWordMatcher.start();
					int len = posEnd - posStart;
					if (len < 0)
						Debug.Assert(len >= 0);
					charIndex += controlWordMatcher.end() - controlWordMatcher.start();

					if (controlWord.Count() != 0)
						deb = 1;

					if (controlWord == "htmlrtf")
					{
						if (controlNumber > 0)
							deb = 1;
						else
							deb = 1;
					}

					//switch (controlWord)  // For now rely on == later on switch once all works !!!
					if (controlWord == "htmlrtf")
					{
						//htmlrtf starts ignored text area, htmlrtf0 ends it
						//Though technically this is not a group, it's easier to treat it as such to ignore everything in between
						if ((controlNumber == nullControlNumber) || (controlNumber == 1))  // nullControlNumber = 0x7FFFFFFF  to indicate ull/invalid number may need to FIX this
							currentGroup.htmlRtf = true;
						else
							currentGroup.htmlRtf = false;
					}
					else if (controlWord == "pard")
					{
						appendIfNotIgnoredGroup(result, strLF, currentGroup, false);
					}
					else if (controlWord == "par")
					{
						appendIfNotIgnoredGroup(result, strLF, currentGroup, false);
					}
					else if (controlWord == "line")
					{
						appendIfNotIgnoredGroup(result, strLF, currentGroup, false);
					}
					else if (controlWord == "tab")
					{
						appendIfNotIgnoredGroup(result, strTAB, currentGroup, false);
					}
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
					else if (controlWord == "fonttbl")  // skipping these groups' contents - these are font and color settings
					{
						currentGroup.ignore = true;   // !!!! ignore will remain set until the first/oldest group with ignore set is deleted
					}
					else if (controlWord == "colortbl")
					{
						currentGroup.ignore = true;  // !!!! ignore will remain set until the first/oldest group with ignore set is deleted
					}
					else if (controlWord == "f")
					{
						// font table index. Might be a new one, or an existing one
						Debug.Assert(controlNumber >= 0);
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
								bool f = fontTable.ContainsKey(currentGroup.fontTableIndex);
								if (!f)
								{
									FontTableEntry new_entry = new();
									fontTable.Add(currentGroup.fontTableIndex, new_entry);
									bool ff = fontTable.ContainsKey(currentGroup.fontTableIndex);
									if (ff)
									{
										//FontTableEntry entry = fontTable[currentGroup.fontTableIndex];
										//entry.charset = possibleCharset;
										fontTable[currentGroup.fontTableIndex].charset = possibleCharset;
									}
								}
								else
								{
									FontTableEntry entry = fontTable[currentGroup.fontTableIndex];
									fontTable[currentGroup.fontTableIndex].charset = possibleCharset;
								}
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

						// ZMM is this applicable to Mail also? How do we know ?
						// Should we check for if Unicde value is between 
						// I don't have enpugh examples
						// From very confusing RTF doc:
#if COMMENTS
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
#if COMMENTS
							// Direct port from Java source. Howver, In Java the char type is 2 bytes most of the time
							char unicodeSymbol = (char)controlNumber;  // 
							appendIfNotIgnoredGroup(result, String(&unicodeSymbol, 1), currentGroup);
#endif
							// ZMM Investigate if not Symbol , private Microsoft Unicode range ??
							// 
							char unicodeSymbol = (char)controlNumber;
							string unicodeSymbolStr = unicodeSymbol.ToString();

							if ((unicodeSymbol >= 0xF020) && (unicodeSymbol <= 0xF0FF))
								deb = 1;
							else
								deb = 1;

							// Any validation of unicodeSymbolStr is needed ?

							appendIfNotIgnoredGroup(result, unicodeSymbolStr, currentGroup, true);

							charIndex += currentGroup.unicodeCharLength;
						}
					}
					else if ((controlWord == "{") || (controlWord == "}") || (controlWord == "\\"))
					{ // Escaped characters
						appendIfNotIgnoredGroup(result, controlWord, currentGroup, false);  // bool isUTF16EncodedSymbols = false
					}
					else if (controlWord == "pntext")
					{
						currentGroup.ignore = true;  // !!!! ignore will remain set until the first/oldest group with ignore set is deleted
					}
					else
						deb = 1;
				}
				else
				{
					appendIfNotIgnoredGroup(result, c.ToString(), currentGroup, false);
					charIndex++;
				}
			}

			if (m_txt7bit.Length > 0)
			{
				m_alwaysEncodeANSI2UTF16 = true;  // remaining chars must be encode as UT16
				if (m_alwaysEncodeANSI2UTF16 || m_ansicpg_setUTF16)
				{
					DWORD error = 0;
					string from = m_txt7bit.ToString();
					string to = "";
					bool ret = ansipgc2utf16(from, m_ansicpg, ref to, ref error);

					if (ret)
					{
						m_ansicpg_setUTF16 = true;
						result.Append(to);
					}
					else
					{
						Debug.Assert(true);
						result.Append(from);
					}
				}
				else
					result.Append(m_txt7bit);

				m_txt7bit.Clear();
			}

			if (m_ansicpg_setUTF16)
				m_ansicpg = 65001; // CP_UTF8;
			else
				deb = 1;

			return new String("");  // Indicate ok;
		}

		// Rewrite hexToString() and appendIfNotIgnoredGroup();
		void appendIfNotIgnoredGroup(StringBuilder result, string symbol, Group group, bool isUTF16EncodedSymbols)
		{
			int deb = 1;
			if (isUTF16EncodedSymbols)
				deb = 1;

			if (!group.ignore && !group.htmlRtf)
			{
				if (symbol.Count() > 4)
					deb = 1;

				if (!isUTF16EncodedSymbols)
				{
					// store for characters defined by \ansicpg; 7 and 8bit not just 7bit ascii
					m_txt7bit.Append(symbol);  // it will appended at some point and may have to be UNICODE encoded.
				}
				else if (m_txt7bit.Length != 0)
				{
					DWORD error = 0;
					string from = m_txt7bit.ToString();
					string to = "";
					bool ret = ansipgc2utf16(from, m_ansicpg, ref to, ref error);
					if (ret)
					{
						result.Append(to);
						m_ansicpg_setUTF16 = true;
					}
					else
					{
						Debug.Assert(true);
						result.Append(from);
					}

					m_txt7bit.Clear();
					result.Append(symbol);
				}
				else
					result.Append(symbol);

				deb = 1;
			}
		}

		// Looks like ref in ref StringBuilder is not needed since class instance are passed by reference
		// However, struct is passed by value.
		// This is very confusing, ref is not required in all cases to pass arg by ref
		// Microsoft always works hard to create user friendly environment :(
		bool hexToString(string hexInput, Charset charsetNumber, ref StringBuilder res)
		{
			int deb = 1;
			res.Clear();

			// result will contain Unicode characters or raw single byte characters in 2 byte char type
			// update code to use ByteArray to avoid confusion
			string result = "";

			bool isUTF16EncodedSymbols = false;

			if (hexInput.Length > 2)
				deb = 1;

			for (int i = 0; i < hexInput.Length; i += 2)
			{
				string _byte = hexInput.Substring(i, 2);

				char charResult = '?';
				if (int.TryParse(_byte, System.Globalization.NumberStyles.HexNumber, null, out int decOutput))
				{
					charResult = (char)decOutput;
				}
				result += charResult;

				deb = 1;
			}

			// Rewrite hexToString() and appendIfNotIgnoredGroup();
			int inCodePage = RTFCharsetNumber2CodePage(charsetNumber);
			if ((charsetNumber == m_ansicpg) || (inCodePage == m_ansicpg))
				deb = 1;
			else
				deb = 1;

			if (result.Count() > 0)
			{
				DWORD error = 0;
				if (inCodePage < 0)
					inCodePage = charsetNumber;

				if (m_alwaysEncodeANSI2UTF16 || (inCodePage != m_ansicpg) || m_ansicpg_setUTF16)
				{
					string to = "";
					bool ret = ansipgc2utf16(result, inCodePage, ref to, ref error);
					if (ret)
					{
						isUTF16EncodedSymbols = true;
						result = to;
					}
					else
					{
						// will return result
						Debug.Assert(true);
					}
				}
				else
					deb = 1;
			}
			res.Append(result);
			return isUTF16EncodedSymbols;
		}

		public static int GetFirstDiffIndex(string str1, string str2)
		{
			if (str1 == null || str2 == null) return -1;

			int length = Math.Min(str1.Length, str2.Length);

			for (int index = 0; index < length; index++)
			{
				if (str1[index] != str2[index])
				{
					return index;
				}
			}

			return -1;
		}
	}

	//

	public class FontTableEntry
	{
		public Charset charset { get; set; }
	}



	public struct RTFCharset
	{
		public RTFCharset(int in_charsetNumber, int in_pageCode, string in_name)
		{
			charsetNumber = in_charsetNumber;
			pageCode = in_pageCode;
			name = in_name;
		}
		public int charsetNumber { get; set; }
		public int pageCode { get; set; }
		public string name { get; set; }
	}


	public class RegexMacher
	{
		public RegexMacher(string pattern, string input)
		{
			m_regex = new Regex(pattern);
			m_match = m_regex.Match("???");
			m_pattern = pattern;
			m_input = input;
			m_all_groups = "";
			m_start = 0;
			m_end = 0;
		}
		~RegexMacher()
		{
			;/// delete m_match;
		}

		public void region(int start, int end)
		{
			if (start < m_start)
				Debug.Assert(start >= m_start);
			m_start = start;
			m_end = end;
		}
		public bool lookingAt()
		{
			int deb = 1;

			int regionLength = m_end - m_start;
			m_match = m_regex.Match(m_input, m_start, regionLength);

			string matchText = m_match.Value;

			if (m_match.Success)
			{
				deb = 1;

				m_all_groups = "";
				for (int i = 0; i < m_match.Groups.Count; ++i)
				{
					string val = m_match.Groups[i].Value;
					int pos = m_match.Groups[i].Index;
					int length = m_match.Groups[i].Length;

					if (i != 0)
						m_all_groups += " \"";
					else
						m_all_groups += "\"";
					m_all_groups += val;
					m_all_groups += "\"";
					deb = 1;
				}

				return true;
			}
			else
				return false;
		}

		public String group(int number)
		{
			string str_group = m_match.Groups[number].Value;
			return str_group;
		}

		public int start()
		{
			if (m_match.Groups.Count > 1)
			{
				int start = m_match.Groups[0].Index;
				return start;
			}
			else
				return 0; // -1 ??  ZMM
		}

		public int end()
		{
			int mcnt = m_match.Groups.Count;
			if (mcnt > 1)
			{
				int pos = m_match.Groups[0].Index;
				int length = m_match.Groups[0].Length;
				int end = pos + length;
				return end;
			}
			else
				return 0; // -1 ??  ZMM
		}

		public int GetSize()
		{
			return m_match.Groups.Count;
		}


		public string GetMatch()
		{
			return m_all_groups;
		}

		public int m_start { get; set; }
		public int m_end { get; set; }
		public string m_pattern { get; set; }
		public string m_input { get; set; }
		public string m_all_groups { get; set; }

		public static string CONTROL_WORD_PATTERN = """\\(([^a-zA-Z])|(([a-zA-Z]+)(-?\d*) ?))""";
		public static string ENCODED_CHARACTER_PATTERN = """^\\'([0-9a-fA-F]{2})""";

		public Match m_match;
		public Regex m_regex;
	};

	public class Group
	{
		public Group()
		{
			ignore = false;
			unicodeCharLength = 1;
			htmlRtf = false;
			fontTableIndex = 0;
			id = next_id++;
		}

		public Group(Group group)
		{
			if (group.ignore == false)
			{
				int deb = 1;
			}
			ignore = group.ignore;
			unicodeCharLength = group.unicodeCharLength;
			htmlRtf = group.htmlRtf;
			// Don't inherit fontTableIndex from parent group.   ??? only in one case see below
			fontTableIndex = group.fontTableIndex;
			id = next_id++;
		}

		public bool ignore { get; set; }
		public int unicodeCharLength { get; set; }
		public bool htmlRtf { get; set; }
		public Integer fontTableIndex { get; set; }
		public int id { get; set; }
		public static int next_id { get; set; }
	}

	public class CharsetHelper
	{
		public static Charset detectCharsetFromRtfContent(string rtf) { return 1252; }
	}

	public class ByteArray
	{
		public ByteArray(int maxSize)
		{
			m_capacity = maxSize;
			m_length = 0;
			byteArray = new byte[maxSize];
		}

		public ref byte[] Array()
		{
			return ref byteArray;
		}

		public int Resize(int length)  // increase
		{
			if (length > m_capacity)
			{
				byte[] newByteArray = new byte[length];
				System.Array.Copy(byteArray, newByteArray, m_length);
				byteArray = newByteArray;
				m_capacity = length;
			}
			return m_capacity;

		}

		public bool Append(byte val)
		{
			int newCapacity = m_capacity;
			if (m_length >= m_capacity)
			{
				newCapacity = Resize(m_capacity + 128);
			}
			if (newCapacity <= m_capacity)
				return false;

			byteArray[m_length++] = val;
			return true;
		}
		public void Clear()
		{
			m_length = 0;
		}

		public string GetString(int codePage)
		{
			Encoding encoding = Encoding.GetEncoding(codePage);
			if (encoding == null)
				encoding = Encoding.UTF8;

			string str = encoding.GetString(byteArray, 0, m_length);
			return str;
		}

		public string GetString(Encoding encoding)
		{
			Debug.Assert(encoding != null);

			string str = encoding.GetString(byteArray, 0, m_length);
			return str;
		}

		public int Length() { return m_length; }

		private int m_capacity { get; set; }
		private int m_length { get; set; }
		private byte[] byteArray;

	}
}


