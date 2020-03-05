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

int SimpleString::FindNoCase(int offset, void const* Src, int  Size)
{
	int i;
	char *p;
	int count = m_count - offset;
	for (i = offset; i < (count - Size); i++)
	{
		p = &m_data[i];
		if (TextUtilsEx::strncmpUpper2Lower(p, (m_count - i), (char*)Src, Size) == 0)
			return i;
	}
	return -1;
}

int SimpleString::FindAny(int offset, void const* Src)
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
		return(p - m_data);
	else
		return -1;
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

void SimpleString::append_internal(void const* Src, size_t  Size) {
	int spaceNeeded = m_count + Size;
	if (spaceNeeded > m_capacity)
		Resize(spaceNeeded);
	::memcpy(&m_data[m_count], Src, Size);
	SetCount(spaceNeeded);
}
