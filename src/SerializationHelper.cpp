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
	if (m_hFile != INVALID_HANDLE_VALUE)
	{
		DWORD nwritten = 0;
		WriteFile(m_hFile, m_buff, m_buff_offset, &nwritten, NULL);
		CloseHandle(m_hFile);
		m_hFile = INVALID_HANDLE_VALUE;
	}

	if (m_buff != NULL)
	{
		free(m_buff);
		m_buff = NULL;
	}
}

#if 0
HANDLE CreateFile(
	LPCSTR                lpFileName,
	DWORD                 dwDesiredAccess,
	DWORD                 dwShareMode,
	LPSECURITY_ATTRIBUTES lpSecurityAttributes,
	DWORD                 dwCreationDisposition,
	DWORD                 dwFlagsAndAttributes,
	HANDLE                hTemplateFile
);
#endif

BOOL SerializerHelper::open(BOOL bWrite, int buffSize)
{
	m_buffCnt = 0;
	m_buff_offset = 0;
	m_file_read_offset = 0;

	if (buffSize > 0)
		m_buffSize = buffSize;
	else
		m_buffSize = SZBUFFSIZE;

	m_writing = bWrite;

	if (bWrite) 
	{
		m_hFile = CreateFile(m_path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
		if (m_hFile == INVALID_HANDLE_VALUE)
		{
			this->close();
			return FALSE;
		}

		m_buff = (char *)malloc(m_buffSize);
		if (m_buff == NULL)
		{
			this->close();
			return false;
		}
		return TRUE;
	}
	else 
	{
		m_buff = (char *)malloc(m_buffSize);
		if (m_buff == NULL)
		{
			this->close();
			return FALSE;
		}

		m_hFile = CreateFile(m_path, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL);
		if (m_hFile == INVALID_HANDLE_VALUE)
		{
			DWORD lError = GetLastError();
			this->close();
			return FALSE;
		}

		DWORD nread;
		BOOL res = ReadFile(m_hFile, m_buff, m_buffSize, &nread, NULL);
		if (!res)
		{
			close();
			return FALSE;
		}
		m_buffCnt = nread;
		return TRUE;
	}
}

#if 0
BOOL SetFilePointerEx(
	HANDLE         hFile,
	LARGE_INTEGER  liDistanceToMove,
	PLARGE_INTEGER lpNewFilePointer,
	DWORD          dwMoveMethod
);
#endif

BOOL SerializerHelper::GetReadPointer(__int64 &fileReadPosition)
{
	if (m_writing == TRUE)
	{
		this->close();
		return -1;
	}

	fileReadPosition = m_file_read_offset;
#if 0
	LARGE_INTEGER   lDistanceToMove;
	lDistanceToMove.QuadPart = 0;
	filePointer = { 0, 0 };
	BOOL ret = SetFilePointerEx(m_hFile, lDistanceToMove, &filePointer, FILE_CURRENT);
	if (ret == FALSE)
	{
		DWORD lError = GetLastError();
		this->close();
		return FALSE;
	}
#endif
	return TRUE;
}

BOOL SerializerHelper::SetReadPointer(__int64 fileReadPosition)
{
	if (m_writing == TRUE)
	{
		this->close();
		return FALSE;
	}

	LARGE_INTEGER filePointer;
	filePointer.QuadPart = fileReadPosition;
	LARGE_INTEGER lNewFilePointer;
	lNewFilePointer.QuadPart = 0;
	BOOL ret = SetFilePointerEx(m_hFile, filePointer, &lNewFilePointer, FILE_BEGIN);
	if (ret == FALSE)
	{
		DWORD lError = GetLastError();
		this->close();
		return FALSE;
	}

	m_buff_offset = 0;
	m_buffCnt = 0;
	m_file_read_offset = lNewFilePointer.QuadPart;
	return ret;
}

BOOL SerializerHelper::readN(void *v, int sz)
{
	if (m_buff == 0)
	{
		this->close();
		return FALSE;
	}

	if (m_buff_offset + sz > m_buffCnt)
	{
		DWORD nread;
		int unreadLength = m_buffCnt - m_buff_offset;
		if (m_buff_offset > 0)
		{
			if (m_buff_offset >= unreadLength)
				memcpy(m_buff, m_buff + m_buff_offset, unreadLength);
			else
			{
				memcpy(m_buff, m_buff + m_buff_offset, m_buff_offset);
				int remainingLen = unreadLength - m_buff_offset;
				memcpy(m_buff + m_buff_offset, m_buff + m_buff_offset + m_buff_offset, remainingLen);
			}
			m_buff_offset = 0;
			m_buffCnt = unreadLength;
		}

		// m_buff_offset == 0
		if (m_buff_offset + sz >= m_buffSize)
		{
			m_buffSize = ((sz / m_buffSize) + 2)*m_buffSize;
			m_buff = (char *)realloc(m_buff, m_buffSize);
			if (m_buff == NULL)
			{
				this->close();
				return FALSE;
			}
		}

		int nBytes2Read = m_buffSize - unreadLength;
		nread = 0;
		BOOL res = ReadFile(m_hFile, m_buff+ unreadLength, nBytes2Read, &nread, NULL);
		if (res == FALSE)
		{
			DWORD lError = GetLastError();
			//this->close();
			//return FALSE;
		}
		m_buffCnt = nread + unreadLength;
		m_buff_offset = 0;
	}

	memcpy(v, m_buff + m_buff_offset, sz);
	m_buff_offset += sz;
	m_file_read_offset += sz;

	return TRUE;
}

BOOL SerializerHelper::writeN(void *v, int sz)
{
	if (m_buff == 0)
	{
		this->close();
		return FALSE;
	}

	if (m_buff_offset + sz > m_buffSize)
	{
		if (m_hFile == INVALID_HANDLE_VALUE)
		{
			this->close();
			return FALSE;
		}

		DWORD nwritten = 0;
		BOOL ret = WriteFile(m_hFile, m_buff, m_buff_offset, &nwritten, NULL);
		if (ret == FALSE)
		{
			DWORD lError = GetLastError();
			//this->close();
			//return FALSE;
		}
		m_buff_offset = 0;
	}

	if (sz > m_buffSize)
	{
		m_buffSize = ((sz / m_buffSize) + 2)*m_buffSize;
		m_buff = (char *)realloc(m_buff, m_buffSize);
		if (m_buff == NULL)
		{
			this->close();
			return FALSE;
		}
	}

	memcpy(m_buff + m_buff_offset, v, sz);
	m_buff_offset += sz;
	m_file_read_offset += sz;
	return TRUE;
}

BOOL SerializerHelper::writeString(LPCSTR val)
{
	int l = istrlen(val);
	if (!writeInt(l))
	{
		this->close();
		return FALSE;
	}
	DWORD written = 0;
	return writeN((void*)val, l);
}

BOOL SerializerHelper::readString(CString &val)
{
	int l = 0;
	if (!readInt(&l))
	{
		this->close();
		return FALSE;
	}
	LPSTR buf = val.GetBufferSetLength(l);
	DWORD nRead = 0;
	BOOL ret = readN(buf, l);
	val.ReleaseBuffer();
	return ret;
}
