#include "stdafx.h"
#include "MimeHelper.h"

#include "Mime.h"


void MboxCMimeHelper::GetContentType(CMimeBody* pBP, CString &value)
{
	const char *fieldName = CMimeConst::ContentType();
	MboxCMimeHelper::GetValue(pBP, fieldName, value);
}

void MboxCMimeHelper::GetContentLocation(CMimeBody* pBP, CString &value)
{
	//const char *fieldName = CMimeConst::ContentLocation();
	const char *fieldName = "Content-Location";
	MboxCMimeHelper::GetValue(pBP, fieldName, value);
}

void MboxCMimeHelper::GetTransferEncoding(CMimeBody* pBP, CString &value)
{
	const char *fieldName = CMimeConst::TransferEncoding();
	MboxCMimeHelper::GetValue(pBP, fieldName, value);
}
void MboxCMimeHelper::GetContentID(CMimeBody* pBP, CString &value)
{
	const char *fieldName = CMimeConst::ContentID();
	MboxCMimeHelper::GetValue(pBP, fieldName, value);
}
void MboxCMimeHelper::GetContentDescription(CMimeBody* pBP, CString &value)
{
	const char *fieldName = CMimeConst::ContentDescription();
	MboxCMimeHelper::GetValue(pBP, fieldName, value);
}
void MboxCMimeHelper::GetContentDisposition(CMimeBody* pBP, CString &value)
{
	const char *fieldName = CMimeConst::ContentDisposition();
	MboxCMimeHelper::GetValue(pBP, fieldName, value);
}
void MboxCMimeHelper::GetCharset(CMimeBody* pBP, CString &value)
{
	string str = pBP->GetCharset();
	value = str.c_str();
}
void MboxCMimeHelper::Name(CMimeBody* pBP, CString &value)
{
	string str = pBP->GetName();
	value = str.c_str();
}
void MboxCMimeHelper::Filename(CMimeBody* pBP, CString &value)
{
	string str = pBP->GetFilename();
	value = str.c_str();
}
void MboxCMimeHelper::GetValue(CMimeBody* pBP, const char* fieldName, CString &value)
{
	const CMimeField *pFld = pBP->CMimeHeader::GetField(fieldName);
	if (pFld)
	{
		string strValue;
		pFld->GetValue(strValue);
		value = strValue.c_str();
	}
	else
		value.Empty();
}

bool MboxCMimeHelper::IsAttachment(CMimeBody* pBP)
{
	CString name;
	CString fileName;
	CString disposition;

	Name(pBP, name);
	Filename(pBP, fileName);
	GetContentDisposition(pBP, disposition);
	if ((disposition.CompareNoCase("attachment") == 0) || (!name.IsEmpty()) || (!fileName.IsEmpty()))
		return true;
	else
		return false;
}