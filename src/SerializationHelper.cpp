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
#include "FileUtils.h"
#include "SerializationHelper.h"

void SerializerHelper::close() 
{
	if (m_buff != NULL) {
		if (m_writing) {
			m_hFile = CreateFile(m_path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
			if (m_hFile != INVALID_HANDLE_VALUE) {
				DWORD nwritten = 0;
				WriteFile(m_hFile, m_buff, m_offset, &nwritten, NULL);
				CloseHandle(m_hFile);
			}
		}
		free(m_buff);
	}
	m_buff = NULL;
}

BOOL SerializerHelper::open(BOOL bWrite) 
{
	m_writing = bWrite;
	if (bWrite) {
		m_buff = (char *)malloc(m_buffSize = SZBUFFSIZE);
		if (m_buff == NULL)
			return false;
		return TRUE;
		//m_hFile = CreateFile(m_path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	}
	else {
		m_buffSize = (int)FileUtils::FileSize(m_path);
		m_buff = (char *)malloc(m_buffSize);
		if (m_buff == NULL)
			return false;
		m_hFile = CreateFile(m_path, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL);
		DWORD nread;
		BOOL res = ReadFile(m_hFile, m_buff, m_buffSize, &nread, NULL);
		CloseHandle(m_hFile);
		m_hFile = INVALID_HANDLE_VALUE;
		return res;
	}
	return m_hFile != INVALID_HANDLE_VALUE;
}

int SerializerHelper::GetReadPointer()
{
	if (m_writing == TRUE)
		return -1;
	return m_offset;
}

BOOL SerializerHelper::SetReadPointer(int pos)
{
	if (m_writing == TRUE)
		return FALSE;

	m_offset = pos;
	return TRUE;
#if 0
	DWORD pos = SetFilePointer(myFile, iBytePos, NULL, FILE_BEGIN);

	if (pos == INVALID_SET_FILE_POINTER)
		return FALSE;
#endif
}

BOOL SerializerHelper::readN(void *v, int sz)
{
	if (m_buff == 0)
		return FALSE;
	if (m_offset + sz > m_buffSize)
		return FALSE;
	memcpy(v, m_buff + m_offset, sz);
	m_offset += sz;
	return TRUE;
}

BOOL SerializerHelper::writeN(void *v, int sz)
{
	if (m_buff == 0)
		return FALSE;
	if (m_offset + sz > m_buffSize) {
		m_buff = (char *)realloc(m_buff, m_buffSize += SZBUFFSIZE);
		if (m_buff == NULL)
			return FALSE;
	}
	memcpy(m_buff + m_offset, v, sz);
	m_offset += sz;
	return TRUE;
}

BOOL SerializerHelper::writeString(LPCSTR val)
{
	int l = strlen(val);
	if (!writeInt(l))
		return FALSE;
	DWORD written = 0;
	return writeN((void*)val, l);
}
BOOL SerializerHelper::readString(CString &val)
{
	int l = 0;
	if (!readInt(&l))
		return false;
	LPSTR buf = val.GetBufferSetLength(l);
	DWORD nRead = 0;
	return readN(buf, l);
}
