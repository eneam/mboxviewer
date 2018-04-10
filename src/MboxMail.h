#pragma once

#include <vector>
#include "UPDialog.h"
#include "Mime.h"
#include "MimeCode.h"

_int64 FileSize(LPCSTR fileName);

class MboxCMimeCodeBase64 : public CMimeCodeBase64
{
public:
	MboxCMimeCodeBase64(const char* pbInput, int nInputSize) {
		SetInput(pbInput, nInputSize, false);
	}
};

class MboxCMimeCodeQP : public CMimeCodeQP
{
public:
	MboxCMimeCodeQP(const char* pbInput, int nInputSize) {
		SetInput(pbInput, nInputSize, false);
	}
};

class MailBodyContent
{
public:
	MailBodyContent() {};
	~MailBodyContent() {};
	CString m_contentType;
	CString m_contentTransferEncoding;
	CString m_contentDisposition;
	CString m_attachmentName;
	int  m_contentOffset;
	int m_contentLength;
};

class MboxMail
{
public:
	MboxMail() {
		m_startOff = m_length = m_hasAttachments = 0;
		m_from_charsetId = m_to_charsetId = m_subj_charsetId = 0;
		m_timeDate = 0;
		m_recv = 1;
	}
	CString GetBody();
	int DumpMailBox(MboxMail *mailBox, int which);

	std::vector <MailBodyContent*> m_ContentDetailsArray;

	_int64 m_startOff;
	int m_hasAttachments;
	int m_length, m_headLength, m_recv;
	time_t m_timeDate;
	CString m_from, m_to, m_subj;
	CString m_from_charset, m_to_charset, m_subj_charset;
	UINT m_from_charsetId, m_to_charsetId, m_subj_charsetId;

	static unsigned char *m_pbOutput;
	static int m_maxOutput;
	static _int64 s_fSize; // current File size
	static _int64 s_oSize; // old file size
	static CString s_path;
	static _int64 s_curmap, s_step;
	static const CUPDUPDATA* pCUPDUPData;
	static void Str2Ansi(CString &res, UINT CodePage);
	static UINT Str2PageCode(const  char* PageCodeStr);
	static void Parse(LPCSTR path);
	static bool Process(char *p, DWORD size, _int64 startOffset, bool bFirstView, bool bLastView, _int64 &lastStartOffset, bool bEml = false);
	static CArray<MboxMail*, MboxMail*> s_mails_ref;  // original cache
	static CArray<MboxMail*, MboxMail*> s_mails;
	static void SortByDate(CArray<MboxMail*, MboxMail*> *s_mails = 0, bool bDesc = false);
	static void SortByFrom(CArray<MboxMail*, MboxMail*> *s_mails = 0, bool bDesc = false);
	static void SortByTo(CArray<MboxMail*, MboxMail*> *s_mails = 0, bool bDesc = false);
	static void SortBySubject(CArray<MboxMail*, MboxMail*> *s_mails = 0, bool bDesc = false);
	static void SortBySize(CArray<MboxMail*, MboxMail*> *s_mails = 0, bool bDesc = false);
	static bool b_mails_sorted;
	static int b_mails_which_sorted;
	static void Destroy() {
		for (int i = 0; i < s_mails.GetSize(); i++)
		{
			for (int j = 0; j < s_mails[i]->m_ContentDetailsArray.size(); j++) {
				delete s_mails[i]->m_ContentDetailsArray[j];
				s_mails[i]->m_ContentDetailsArray[j] = 0;
			}
			delete s_mails[i];
		}
		s_mails.RemoveAll();
		s_mails_ref.RemoveAll();
		b_mails_sorted = false;
	};
};

#define SZBUFFSIZE 1024*1024

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
	void close() {
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
	BOOL open(BOOL bWrite) {
		m_writing = bWrite;
		if (bWrite) {
			m_buff = (char *)malloc(m_buffSize = SZBUFFSIZE);
			if (m_buff == NULL)
				return false;
			return TRUE;
			//m_hFile = CreateFile(m_path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
		}
		else {
			m_buffSize = (int)FileSize(m_path);
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
	BOOL readN(void *v, int sz) {
		if (m_buff == 0)
			return FALSE;
		if (m_offset + sz > m_buffSize)
			return FALSE;
		memcpy(v, m_buff + m_offset, sz);
		m_offset += sz;
		return TRUE;
	}
	BOOL writeN(void *v, int sz) {
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
	BOOL writeString(LPCSTR val) {
		int l = strlen(val);
		if (!writeInt(l))
			return FALSE;
		DWORD written = 0;
		return writeN((void*)val, l);
	}
	BOOL readString(CString &val) {
		int l = 0;
		if (!readInt(&l))
			return false;
		LPSTR buf = val.GetBufferSetLength(l);
		DWORD nRead = 0;
		return readN(buf, l);
	}
};

/* This is too slow
class SerializerHelper
{
private:
CString m_path;
HANDLE m_hFile;
public:
SerializerHelperFs(LPCSTR fn) {
m_path = fn;
}
~SerializerHelperFs() {
close();
}
void close() {
if( m_hFile != INVALID_HANDLE_VALUE )
CloseHandle( m_hFile );
m_hFile = INVALID_HANDLE_VALUE;
}
BOOL open(BOOL bWrite) {
if( bWrite )
m_hFile = CreateFile(m_path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
else
m_hFile = CreateFile(m_path, GENERIC_READ, 0, NULL,	OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL);
return m_hFile != INVALID_HANDLE_VALUE;
}
BOOL writeInt(int value) {
//if( m_hFile == INVALID_HANDLE_VALUE )			return false;
DWORD written = 0;
return WriteFile(m_hFile, &value, sizeof(int), &written, NULL);
}
BOOL readInt(int *val) {
//if( m_hFile == INVALID_HANDLE_VALUE ) return false;
DWORD nRead = 0;
return ReadFile(m_hFile, val, sizeof(int), &nRead, NULL);
}
BOOL writeInt64(_int64 value) {
//if( m_hFile == INVALID_HANDLE_VALUE )	return false;
DWORD written = 0;
return WriteFile(m_hFile, &value, sizeof(_int64), &written, NULL);
}
BOOL readInt64(_int64 *val) {
//if( m_hFile == INVALID_HANDLE_VALUE )	return false;
DWORD nRead = 0;
return ReadFile(m_hFile, val, sizeof(_int64), &nRead, NULL);
}
BOOL writeString(LPCSTR val) {
int l = strlen(val);
if( !writeInt(l) )
return false;
DWORD written = 0;
return WriteFile(m_hFile, val, l, &written, NULL);
}
BOOL readString(CString &val) {
int l = 0;
if( ! readInt(&l) )
return false;
LPSTR buf = val.GetBufferSetLength(l);
DWORD nRead = 0;
return ReadFile(m_hFile, buf, l, &nRead, NULL);
}
};
*/
