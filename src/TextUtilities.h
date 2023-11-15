//
//////////////////////////////////////////////////////////////////
//
//  Windows Mbox Viewer is a free tool to view, search and print mbox mail archives.
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
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
//  Library General Public License for more details.
//
//  You should have received a copy of the GNU Library General Public
//  License along with this program; if not, write to the
//  Free Software Foundation, Inc., 51 Franklin St, Fifth Floor,
//  Boston, MA  02110 - 1301, USA.
//
//////////////////////////////////////////////////////////////////
//

#ifndef _TEXT_UTILITIES_H_

#define _TEXT_UTILITIES_H_

#ifdef _MSC_VER
#include <afxtempl.h>
#pragma warning (disable : 4018)
#else
#include "CString.h"
#endif


#ifndef _VECTOR_
#include <vector>
#endif

#include "SimpleString.h"

template<class T> class MyCArray : public CArray<T, T>
{
public:
	void SetSizeKeepData(INT_PTR nNewSize, INT_PTR nGrowBy = -1);
	void CopyKeepData(const MyCArray<T>& src);
};

class TextUtilities
{
	void MakeWordSeparatorLookupTable(const char* s, bool* l);
	void MakeLookupTable(const char *s, bool *l);
	void MakeConversionTable(const char* from, const char* to, char* dest);

	CMap<const char *, const char *, char, char> entity_map;

public:
	TextUtilities() { Init(); };
	~TextUtilities();
	bool Init();

	bool IsWordSeparator(const char c);
	char ToLower(char c);

	//
	int StrSearch(unsigned char* text, int n, UINT textCP, unsigned char* pat, int m, BOOL bCaseSens, unsigned char *skipTable);
	int StrSearchW(unsigned char* text, int n, UINT textCP, unsigned char* pat, int m, BOOL bCaseSens, unsigned char *skipTable);
	int BMHSearch( unsigned char *text, int n, unsigned char *pat, int m, BOOL bCaseSens, unsigned char *skipTable);
	int BMHSearchW( unsigned char *text, int n, unsigned char *pat, int m, BOOL bCaseSens, unsigned char *skipTable);
	//
	static UINT32 TextUtilities::CalcCRC32(const char* buf, const UINT length);
	void hexdump(char *title, char *area, int length);

	static CString m_strW;
	static SimpleString m_strA;
	static SimpleString m_workBuff;

	static bool TestAll();
};

extern TextUtilities g_tu;

#endif