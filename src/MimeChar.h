//////////////////////////////////////////////////////////////////////
//
// MIME Character Classification
//
// Jeff Lee
// Dec 16, 2000
//
//////////////////////////////////////////////////////////////////////

#if !defined(_MIME_CHAR_H)
#define _MIME_CHAR_H

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CMimeChar
{
public:
	static bool IsNonAscii(unsigned char ch) { return (ch & 0x80) != 0; }
	static bool IsControl(unsigned char ch) { return !(ch & 0x80) && (m_aCharTbl[ch] & CONTROL) != 0; }
	static bool IsSpace(unsigned char ch) { return !(ch & 0x80) && (m_aCharTbl[ch] & SPACE) != 0; }
	static bool IsPrintable(unsigned char ch) { return !(ch & 0x80) && (m_aCharTbl[ch] & PRINT) != 0; }
	static bool IsSpecial(unsigned char ch) { return !(ch & 0x80) && (m_aCharTbl[ch] & SPECIAL) != 0; }
	static bool IsHexDigit(unsigned char ch) { return !(ch & 0x80) && (m_aCharTbl[ch] & HEXDIGIT) != 0; }
	static bool IsDelimiter(unsigned char ch) { return IsSpace(ch) || IsSpecial(ch); }
	static bool IsToken(unsigned char ch) { return IsNonAscii(ch) || (ch > ' ' && (!IsSpecial(ch)|| ch == '.')); }

private:
	enum { CONTROL=0x01, SPACE=0x02, PRINT=0x04, PUNCT=0x10, SPECIAL=0x20, HEXDIGIT=0x40 };
	static const unsigned char m_aCharTbl[128];
};

#endif // _MIME_CHAR_H