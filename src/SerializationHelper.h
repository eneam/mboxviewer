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


#pragma once


#define SZBUFFSIZE 1024*1024*100

class SerializerHelper
{
private:
	BOOL m_writing;
	CString m_path;
	HANDLE m_hFile;
	int m_offset;
	char * m_buff;
	int m_buffSize;
public:
	SerializerHelper(LPCSTR fn) {
		m_path = fn;
		m_offset = 0;
		m_buff = NULL;
	}
	~SerializerHelper() {
		close();
	}
	void close();
	BOOL open(BOOL bWrite);
	int GetReadPointer();
	BOOL SetReadPointer(int pos);
	BOOL readN(void *v, int sz);
	BOOL writeN(void *v, int sz);
	BOOL writeString(LPCSTR val);
	BOOL readString(CString &val);

	BOOL writeInt(int val) {
		return writeN(&val, sizeof(int));
	}
	BOOL readInt(int *val) {
		return readN(val, sizeof(int));
	}
	BOOL readUInt(unsigned int *val) {
		return readN(val, sizeof(int));
	}
	BOOL writeInt64(_int64 value) {
		return writeN(&value, sizeof(_int64));
	}
	BOOL readInt64(_int64 *val) {
		return readN(val, sizeof(_int64));
	}
};