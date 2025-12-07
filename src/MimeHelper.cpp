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
#include "MimeHelper.h"

#include "Mime.h"


void MboxCMimeHelper::GetContentType(CMimeBody* pBP, CStringA &value)
{
	const char *fieldName = CMimeConst::ContentType();
	MboxCMimeHelper::GetValue(pBP, fieldName, value);
}

void MboxCMimeHelper::GetContentLocation(CMimeBody* pBP, CStringA &value)
{
	//const char *fieldName = CMimeConst::ContentLocation();
	const char *fieldName = "Content-Location";
	MboxCMimeHelper::GetValue(pBP, fieldName, value);
}

void MboxCMimeHelper::GetTransferEncoding(CMimeBody* pBP, CStringA &value)
{
	const char *fieldName = CMimeConst::TransferEncoding();
	MboxCMimeHelper::GetValue(pBP, fieldName, value);
}
void MboxCMimeHelper::GetContentID(CMimeBody* pBP, CStringA &value)
{
	const char *fieldName = CMimeConst::ContentID();
	MboxCMimeHelper::GetValue(pBP, fieldName, value);
}
void MboxCMimeHelper::GetContentDescription(CMimeBody* pBP, CStringA &value)
{
	const char *fieldName = CMimeConst::ContentDescription();
	MboxCMimeHelper::GetValue(pBP, fieldName, value);
}
void MboxCMimeHelper::GetContentDisposition(CMimeBody* pBP, CStringA &value)
{
	const char *fieldName = CMimeConst::ContentDisposition();
	MboxCMimeHelper::GetValue(pBP, fieldName, value);
}
void MboxCMimeHelper::GetCharset(CMimeBody* pBP, CStringA &value)
{
	std::string str = pBP->GetCharset();
	value = str.c_str();
}
void MboxCMimeHelper::Name(CMimeBody* pBP, CStringA &value)
{
	std::string str = pBP->GetName();
	value = str.c_str();
}
void MboxCMimeHelper::Filename(CMimeBody* pBP, CStringA &value)
{
	std::string str = pBP->GetFilename();
	value = str.c_str();
}
void MboxCMimeHelper::GetValue(CMimeBody* pBP, const char* fieldName, CStringA &value)
{
	const CMimeField *pFld = pBP->CMimeHeader::GetField(fieldName);
	if (pFld)
	{
		std::string strValue;
		pFld->GetValue(strValue);
		value = strValue.c_str();
	}
	else
		value.Empty();
}

bool MboxCMimeHelper::IsAttachment(CMimeBody* pBP)
{
	CStringA name;
	CStringA fileName;
	CStringA disposition;

	Name(pBP, name);
	Filename(pBP, fileName);
	GetContentDisposition(pBP, disposition);
	int dispositionMatchRet = disposition.CompareNoCase("attachment");
	if ((dispositionMatchRet == 0) || (!name.IsEmpty()) || (!fileName.IsEmpty()))
		return true;
	else
		return false;
}

bool MboxCMimeHelper::IsInlineAttachment(CMimeBody* pBP)
{
	CStringA name;
	CStringA fileName;
	CStringA disposition;

	Name(pBP, name);
	Filename(pBP, fileName);
	GetContentDisposition(pBP, disposition);
	int dispositionMatchRet = disposition.CompareNoCase("inline");
	if ((dispositionMatchRet == 0) && ((!name.IsEmpty()) || (!fileName.IsEmpty())))
		return true;
	else
		return false;
}


bool MboxCMimeHelper::IsAttachmentDisposition(CMimeBody* pBP)
{
	CStringA name;
	CStringA fileName;
	CStringA disposition;

	Name(pBP, name);
	Filename(pBP, fileName);
	GetContentDisposition(pBP, disposition);
	int dispositionMatchRet = disposition.CompareNoCase("attachment");
	if (dispositionMatchRet == 0)
		return true;
	else
		return false;
}

bool MboxCMimeHelper::IsInlineDisposition(CMimeBody* pBP)
{
	CStringA name;
	CStringA fileName;
	CStringA disposition;

	Name(pBP, name);
	Filename(pBP, fileName);
	GetContentDisposition(pBP, disposition);
	int dispositionMatchRet = disposition.CompareNoCase("inline");
	if (dispositionMatchRet == 0)
		return true;
	else
		return false;
}