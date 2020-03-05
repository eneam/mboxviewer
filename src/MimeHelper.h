#pragma once

#include "afxstr.h"
#include "Mime.h"
#include "MimeCode.h"

class CMimeBody;

class MboxCMimeHelper
{
public:
	static void GetContentType(CMimeBody* pBP, CString &value);
	static void GetContentLocation(CMimeBody* pBP, CString &value);
	static void GetTransferEncoding(CMimeBody* pBP, CString &value);
	static void GetContentID(CMimeBody* pBP, CString &value);
	static void GetContentDescription(CMimeBody* pBP, CString &value);
	static void GetContentDisposition(CMimeBody* pBP, CString &value);
	static void GetCharset(CMimeBody* pBP, CString &value);
	static void Name(CMimeBody* pBP, CString &value);
	static void Filename(CMimeBody* pBP, CString &value);
	//
	static bool IsAttachment(CMimeBody* pBP);
protected:
	static void GetValue(CMimeBody* pBP, const char* fieldName, CString &value);
};

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

