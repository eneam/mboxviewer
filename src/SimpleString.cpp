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
#include "SimpleString.h"
#include "TextUtilsEx.h"


// Expensive !!!
int SimpleString::FindNoCase(int offset, char const* Src, int  Size)
{
	int i;
	char *p;
	for (i = offset; i < (m_count - Size); i++)
	{
		p = &m_data[i];
		if (TextUtilsEx::strncmpUpper2Lower(p, (m_count - i), (char*)Src, Size) == 0)
			return i;
	}
	return -1;
}


// Expensive !!!
int SimpleString::FindNoCase(int offset, char const* Src, int  Size, int ncount)
{
	int i;
	char* p;
	if (ncount > m_count)
		ncount = m_count;

	for (i = offset; i < (ncount - Size); i++)
	{
		p = &m_data[i];
		if (TextUtilsEx::strncmpUpper2Lower(p, ncount - i, (char*)Src, Size) == 0)
			return i;
	}
	return -1;
}

int SimpleString::FindAny(int offset, char const* Src)
{
	int i;
	const char *p = m_data;
	int c;
	for (i = offset; i < m_count; i++)
	{
		c = m_data[i];
		p = ::strchr((const char*)Src, c);
		if (p)
			return i;
	}
	return -1;
}

int SimpleString::Find(int offset, char const c)
{
	const char *p = ::strchr((const char*)(m_data + offset), c);
	if (p)
		return(IntPtr2Int(p - m_data));
	else
		return -1;
}

int SimpleString::ReplaceNoCase(int offset, char const* Src, int Size, char const* newSrc, int newSize, int ncount)
{
	int ret = -1;
	int findSrcPos = FindNoCase(offset, Src, Size, ncount);
	if (findSrcPos >= 0)
	{
#if _DEBUG
		// Make local
		SimpleString tmp;
		{
			int needSize = m_count;
			if (newSize > Size)
			{
				needSize = m_count + (newSize - Size);
			}
			tmp.Resize(needSize);
			char* p = this->Data();

			tmp.Append(p, findSrcPos);
			tmp.Append(newSrc, newSize);

			char* pNext = this->Data() + findSrcPos + Size;
			tmp.Append(pNext, this->Count() - (findSrcPos + Size));
		}
#endif
		int needSize = m_count;
		if (newSize > Size)
		{
			int oldCapacity = Capacity();
			needSize = m_count + (newSize - Size);
			int newCapacity = Resize(needSize);

			int i;
			char* pNew = &m_data[needSize];
			char *pOld = &m_data[needSize - Size];
			for (i = m_count; i > findSrcPos; i--)
			{
				*pNew-- = *pOld--;
			}
			char *p = p = &m_data[findSrcPos];
			// Replace inplace
			memcpy(p, newSrc, newSize);
		}
		else
		{
			needSize = m_count - Size + newSize;

			char* p = &m_data[findSrcPos];
			// Replace inplace
			memcpy(p, newSrc, newSize);

			int i;
			char* pNew = &m_data[findSrcPos + newSize];
			char* pOld = &m_data[findSrcPos + Size];
			for (i = findSrcPos + Size; i < m_count; i++)
			{
				*pNew++ = *pOld++;
			}
		}
		SetCount(needSize);

		ret = findSrcPos;
#if _DEBUG
		// Make local
		{
			_ASSERTE(tmp.Count() == this->Count());
			int j = 0;
			for (j = 0; j < tmp.Count(); j++)
			{
				char c1 = tmp.GetAt(j);
				char c2 = this->GetAt(j);
				if (c1 != c2)
					int deb = 1;
			}
			_ASSERTE(memcmp(tmp.Data(), this->Data(), tmp.Count()) == 0);
		}
#endif
		int deb = 1;

	}
	return ret;
}

int SimpleString::Remove(int offsetBegin, int offsetEnd)
{
	_ASSERTE(offsetBegin >= 0);
	_ASSERTE(offsetBegin <= offsetEnd);

	if (offsetBegin <= 0)
		return m_count;

	if (offsetBegin >= offsetEnd)
		return m_count;

	int removeCnt = offsetEnd - offsetBegin;
	int newCnt = m_count - removeCnt;

	char* pDest = &m_data[offsetBegin];
	char* pSrc = &m_data[offsetEnd];
	int i;
	for (i = offsetEnd; i < m_count; i++)
	{
		*pDest++ = *pSrc++;
	}
	SetCount(newCnt);
	return m_count;
}

int SimpleString::Resize(int size)
{
	if (size > m_capacity)
	{
		int new_capacity = size + m_grow_size; // fudge factor ?
		char *new_data = new char[new_capacity + 1]; // one extra to set NULL
		if (new_data) {
			if (m_count > 0)
				::memcpy(new_data, m_data, m_count);
			delete[] m_data;
			m_data = new_data;
			m_data[m_count] = 0;
			m_capacity = new_capacity;
		}
		else
			; // trouble :) caller needs to handle this ?
	}
	return m_capacity;
}

void SimpleString::append_internal(char const* Src, int  Size) {
	int spaceNeeded = m_count + Size;
	if (spaceNeeded > m_capacity)
		Resize(spaceNeeded);
	::memcpy(&m_data[m_count], Src, Size);
	SetCount(spaceNeeded);
}
