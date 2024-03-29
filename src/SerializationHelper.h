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


#define SZBUFFSIZE 1024*100

class SerializerHelper
{
private:
	BOOL m_writing;
	CString m_path;
	HANDLE m_hFile;
	__int64 m_fileSize;
	__int64  m_file_read_offset;  //  read offset into file
	int m_buff_offset;  // offset into m_buff
	char *m_buff;
	int m_buffSize;  // allocated size
	int m_buffCnt;	// bytes in m_buff

public:
	SerializerHelper(LPCWSTR fn)
	{
		m_writing = FALSE;
		m_path = fn;
		m_fileSize = 0;
		m_file_read_offset = 0;
		m_buff_offset = 0;
		m_buffCnt = 0;
		m_buff = NULL;
		m_hFile = INVALID_HANDLE_VALUE;
	}
	~SerializerHelper() {
		close();
	}
	CString GetFileName() {
		return m_path;
	};
	void close();
	BOOL open(BOOL bWrite, int buffSize = 0);
	BOOL GetReadPointer(__int64 &fileReadPosition);
	BOOL SetReadPointer(__int64 fileReadPosition);
	BOOL readN(void *v, int sz);
	BOOL writeN(void *v, int sz);
	BOOL writeString(LPCSTR val);
	BOOL readString(CStringA &val);

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