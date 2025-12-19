//
//////////////////////////////////////////////////////////////////
//
//  Windows Mbox Viewer is a free tool to view, search and print mbox mail archives..
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

// problem if using ISO c++17 and above
// deleted "using namespace std" in header files and updated code instead of below
//#define _HAS_STD_BYTE 0 

// Fix vars definitions
#pragma warning(disable : 4267) // warning C4267: 'initializing': conversion from 'size_t' to 'int', possible loss of data
#pragma warning(disable : 4244)  // warning C4244: '=': conversion from 'UINT64' to 'UINT', possible loss of data
#pragma warning(disable : 4996)  // warning C4996: 'strnicmp': The POSIX name for this item is deprecated.Instead, use the ISO C and C++ conformant name : _strnicmp.

#include "ReportMemoryLeaks.h"

#include "stdafx.h"

#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <errno.h>
#ifndef _WINDOWS
#include <error.h>
#include <iconv.h>
#endif

#include "cfbf.h"
#include "outlook.h"
#include "RTF2HTMLConverter.h"

#include "MimeParser.h"
#include "FileUtils.h"
#include "TextUtilsEx.h"
#include "mimetypes.h"


extern int maxTextDumpLength;
extern int maxBinaryDumpLength;
extern int maxRtfDumpLength;

static size_t
strlen_utf16(uint16_t *s)
{
    size_t count = 0;
    while (*s) {
        count++;
        s++;
    }
    return count;
}

static int
strncmp_utf16(const uint16_t *s1, const uint16_t *s2, size_t n)
{
    for (size_t i = 0; i < n; ++i) {
        if (s1[i] < s2[i])
            return -1;
        else if (s1[i] > s2[i])
            return 1;
        else if (s1[i] == 0) {
            // found the end of the string
            return 0;
        }
    }
    return 0;
}

static uint16_t *
strchr_utf16(const uint16_t *s, int c)
{
    while (*s) {
        if (*s == c)
            return (uint16_t *) s;
        ++s;
    }
    return NULL;
}

void GetIndentStr(int level, std::string& indentstr)
{
    indentstr.append(level * 4, ' ');

    char levelStrBuff[16];
    char* levelStr = _itoa(level, levelStrBuff, 10);
    if (level <= 1)
        indentstr[0] = levelStr[0];
    else
        indentstr[(level - 1) * 4] = levelStr[0];
}

bool IsSubStream(std::string& name)
{
    int isSubsStream = (strncmp(name.c_str(), "__substg1.0_", strlen("__substg1.0_")) == 0);
    return isSubsStream;
}

bool IsPropertySubStream(std::string& name)
{
    int isPropertySubsStream = (strncmp(name.c_str(), "__properties_", strlen("__properties_")) == 0);
    return isPropertySubsStream;
}

void GetPropertyIdAndType(std::string& name, 
    std::string& propertyIdStr, int& propertyIdNumb, 
    std::string& propertyTypeStr, int& propertyTypeNumb)
{
    propertyIdNumb = 0;
    propertyTypeNumb = 0;

    char propertyId[5];
    char propertyType[5];
    propertyId[0] = 0; propertyType[0] = 0;

    int prefixlen = strlen("__substg1.0_");
    if (strncmp(name.c_str(), "__substg1.0_", prefixlen) != 0)
        return;

    char suffix[9];
    strncpy(suffix, &name[prefixlen], 8);
    suffix[8] = 0;
    strncpy(propertyId, suffix, 4);
    propertyId[4] = 0;
    strncpy(propertyType, &suffix[4], 4);
    propertyType[4] = 0;
    propertyIdStr.append(propertyId);
    propertyTypeStr.append(propertyType);

    propertyTypeNumb = (int)strtol(propertyType, NULL, 16);
    propertyIdNumb = (int)strtol(propertyId, NULL, 16);
}

int DumpMethod(struct DirEntry* entry, int propertyIdNumb, int propertyTypeNumb)
{
    std::string name_utf8 = UTF16ToUTF8((wchar_t*)entry->name, entry->name_length);
    if (strcmp(name_utf8.c_str(), "__properties_version1.0") == 0)
    {
        return 0;
    }
    if ((propertyIdNumb == PropertId::PidTagAttachDataObject) && (propertyTypeNumb == PropertyType::PtypBinary))
        return 0;

    if ((propertyIdNumb == PropertId::PidTagRtfCompressed) || (propertyIdNumb == PropertId::PidTagRtfInSync))
    {
        if (entry->stream_size > 16)
            return 4; // Compressed RTF
        else
            int deb = 0;
    }

    if (propertyIdNumb == PidTagBodyHtml)
        propertyTypeNumb = PropertyType::PtypString8;

    if (propertyTypeNumb == PropertyType::PtypBinary)
        return 1;
    else if (propertyTypeNumb == PropertyType::PtypString)
        return 2;
    else if (propertyTypeNumb == PropertyType::PtypString8)
        return 3;
    else
        return 2;
}

void DumpBuffer(FILE* out, const void* buffer, size_t len);

void DumpRTF(FILE* out, const void* buffer, size_t len)
{
    int maxDumpLength = maxRtfDumpLength;
    if (maxDumpLength == 0)
    {
        fprintf(out, "\n");
        return;
    }

    if (len < 16)   // check PidTagStoreSupportMask to see if RTF is compressed
    {
        DumpBuffer(out, buffer, len);
        return;
    }

    int srclen = 0;
    unsigned char* dst = 0;
    unsigned char** dest = &dst;
    char errorText[256]; errorText[0] = 0;
    int errorTextLen = 256;

    const unsigned char* src = static_cast<const unsigned char*>(buffer);

    int retlen = RtfDecompressor::DecompressRtf(out, (unsigned char*)src, len, dest, errorText, errorTextLen);
    if (retlen > 0)
    {
        dst[retlen] = 0;

        std::string RtfContentType("HTML");
        char* cpos = strstr((char*)dst, "fromhtml1");
        if (cpos == 0)
            RtfContentType = "TEXT";

        if (maxDumpLength < 0)
        {
            // Fix type casting ZMM
            if ((int)len <= maxRtfDumpLength)
                maxDumpLength = len;
        }
        else if (retlen > maxRtfDumpLength)
            retlen = maxRtfDumpLength;

        if (retlen <= maxRtfDumpLength)
            fprintf(out, " ");
        else
            fprintf(out, "\n");

        std::string bin((const char*)dst, retlen);

        fprintf(out, "\"%s\"", bin.c_str());

        RTF2HTMLConverter rtf2html;

        std::string result = rtf2html.rtf2html((char*)dst);
        std::string html((const char*)result.c_str(), retlen);

        //fprintf(stdout, "\n\nRTF2HTMLConverter: rtf2html: !!!!!!!!!!!!!!!\n%s\n\n", html.c_str());
        fprintf(out, " %s \"%s\"", RtfContentType.c_str(), html.c_str());

        free(dst);
    }
    else
    {
        DumpBuffer(out, buffer, len);
    }

    fprintf(out, "\n");
}

void DumpBuffer(FILE* out, const void* buffer, size_t len)
{
    int maxDumpLength = maxBinaryDumpLength;
    if (maxDumpLength == 0)
    {
        fprintf(out, "\n");
        return;
    }

	if (maxDumpLength < 0)
	{
		if (len <= 16)
			maxDumpLength = len;
	}
	else if ((int)len > maxDumpLength)  // Fix type casting ZMM
		len = maxDumpLength;

    if (len <= 16)
        fprintf(out, " ");
    else
        fprintf(out, "\n");

    const unsigned char* str = static_cast<const unsigned char*>(buffer);
    for (size_t i = 0; i < len; i++)
    {
        if (i > 0 && i % 16 == 0)
            fprintf(out, "\n");

        fprintf(out, "%02x ", static_cast<int>(str[i]));
    }
    fprintf(out, "\n");
}

void DumpTextU16(FILE* out, const char* buffer, size_t len)
{
    int maxDumpLength = maxTextDumpLength;
    if (maxDumpLength == 0)
    {
        fprintf(out, "\n");
        return;
    }
    if (maxDumpLength < 0)
    {
        if (len <= 64)
            maxDumpLength = len;
    }
    else if ((int)len > maxDumpLength)  // Fix type casting ZMM
        len = maxDumpLength;

    char* str = (char*)buffer;
    if ((int)len <= maxDumpLength)
    {
        fprintf(out, " ");
        char* buf = (char*)cfbf_malloc(len);
        for (int i = 0; i < (int)len; i++) // Fix type casting ZMM
        {
            if ((buffer[i] == '\n') || (buffer[i] == '\r'))
                buf[i] = ' ';
            else
                buf[i] = buffer[i];
            str = buf;
        }
    }
    else
    {
        fprintf(out, "\n");
    }

	std::string buff8 = UTF16ToUTF8((wchar_t*)str, len / 2);
	char* wstr = (char*)buff8.c_str();
	fprintf(out, "%s\n", wstr);
	if (str != buffer)
		free(str);
}

void DumpTextU8(FILE* out, const char* buffer, size_t len)
{
    int maxDumpLength = maxTextDumpLength;
    if (maxDumpLength == 0)
    {
        fprintf(out, "\n");
        return;
    }
    if (maxDumpLength < 0)
    {
        if (len <= 32)
            maxDumpLength = len;
    }
    else if ((int)len > maxDumpLength)  // Fix type casting ZMM
        len = maxDumpLength;

    char* str = (char*)buffer;
    if ((int)len <= maxDumpLength)
    {
        fprintf(out, " ");
        char* buf = (char*)cfbf_malloc(len);
        for (int i = 0; i < (int)len; i++)  // Fix type casting ZMM
        {
            if ((buffer[i] == '\n') || (buffer[i] == '\r'))
                buf[i] = ' ';
            else
                buf[i] = buffer[i];
            str = buf;
        }
    }
    else
    {
        fprintf(out, "\n");
    }

	std::string txt(str, len);
	fprintf(out, "%s\n", txt.c_str());
	if (str != buffer)
		free(str);
}

int
PrintOutlookObject(void* cookie, struct cfbf* cfbf, DirEntry* e,
    DirEntry* parent, unsigned long entry_id, int depth);

int
ParseOutlookMsg(void* cookie, struct cfbf* cfbf, DirEntry* e,
    DirEntry* parent, unsigned long entry_id, int depth)
{
    int name_length;
    int retval = 1;

    OutlookMsgHelper* msgHelper = (OutlookMsgHelper*)cookie;
    FILE* out = msgHelper->out;

    // Ignore for now. Review later
    if (e->name_length > 64) {
        char txt[512];
        sprintf(txt, "warning: dir entry %lu: name_length is %hu which is > 64", entry_id, (unsigned short)e->name_length);
        name_length = 64;
    }
    else {
        name_length = (int)e->name_length;
    }

    std::string parent_name_utf8;
    std::string name_utf8 = UTF16ToUTF8((wchar_t*)e->name, e->name_length);
    if (parent)
        parent_name_utf8 = UTF16ToUTF8((wchar_t*)parent->name, parent->name_length);

    int isRoot = (e->object_type == CFBF_ROOT_OBJECT_TYPE) ? 1 : 0;
    int isStorage = (e->object_type == CFBF_STORAGE_OBJECT_TYPE) ? 1 : 0;
    int isStream = (e->object_type == CFBF_STREAM_OBJECT_TYPE) ? 1 : 0;

    if (isRoot)
        return 1;

    int retSet = msgHelper->active_msg->SetProperty(parent, e);

    _ASSERTE(parent);
    if (parent)
    {
        if (name_utf8.compare("__substg1.0_3701000D") == 0)
        {
            //fprintf(stdout, "Attach Object Type found __substg1.0_3701000D\n");

            msgHelper->m_msgList.push_back(msgHelper->active_msg);
            msgHelper->active_msg = new OutlookMessage(msgHelper->out, msgHelper->emlFileHandle, cfbf);
            int deb = 1;
        }
    }

    if (msgHelper->out)
    {
        int retPrint = PrintOutlookObject(cookie, cfbf, e, parent, entry_id, depth);
        int deb = 1;
    }
    return retval;
}


int
PrintOutlookObject(void* cookie, struct cfbf* cfbf, DirEntry* e,
    DirEntry* parent, unsigned long entry_id, int depth)
{
    std::string errorText;
    int retval = 1;
    char obj_type_str[10];

    OutlookMsgHelper* msgHelper = (OutlookMsgHelper*)cookie;
    FILE* out = msgHelper->out;

    int indent = (depth - 1) * 4;
    std::string indentstr;
    GetIndentStr(depth, indentstr);

    std::string parent_name_utf8;
    std::string name_utf8 = UTF16ToUTF8((wchar_t*)e->name, e->name_length);
    if (parent)
        parent_name_utf8 = UTF16ToUTF8((wchar_t*)parent->name, parent->name_length);

    int isRoot = (e->object_type == CFBF_ROOT_OBJECT_TYPE) ? 1 : 0;
    int isStorage = (e->object_type == CFBF_STORAGE_OBJECT_TYPE) ? 1 : 0;
    int isStream = (e->object_type == CFBF_STREAM_OBJECT_TYPE) ? 1 : 0;

    int is_substg1 = (strncmp(name_utf8.c_str(), "__substg1.0_", strlen("__substg1.0_")) == 0) ? 1 : 0;
    int is_properties_version1 = (strncmp(name_utf8.c_str(), "__properties_version1", strlen("__properties_version1")) == 0) ? 1 : 0;
    if (is_properties_version1)
        int deb = 1;
    
    std::string start_sector;
    if ((unsigned long)e->start_sector == CFBF_END_OF_CHAIN)
        start_sector = "FFFFFFFF";
    else {
        char buff[64];
        char* sector = _itoa((unsigned long)e->start_sector, buff, 10);
        start_sector.append(sector);
    }

    int isMiniStream = cfbf_dir_stored_in_mini_stream(cfbf, e);
    cfbf_object_type_to_string(e->object_type, obj_type_str, sizeof(obj_type_str));

    std::string propertyIdStr;
    int propertyIdNumb = 0;
    std::string propertyTypeStr;
    int propertyTypeNumb = 0;

    char* propertyName = 0;
    char* type = 0;

    if (isStorage)
	{
		fprintf(out, "%-9s %8s%s %10llu    ", obj_type_str,
			start_sector.c_str(),
			isMiniStream ? "m" : " ",
			(unsigned long long) e->stream_size);

		fprintf(out, "%s", indentstr.c_str());
		fprintf(out, "%s", name_utf8.c_str());

        if (name_utf8.compare("__substg1.0_3701000D") == 0)
		{
			GetPropertyIdAndType(name_utf8, propertyIdStr, propertyIdNumb, propertyTypeStr, propertyTypeNumb);

			propertyName = id2name(propertyIdNumb);
            type = id2type(propertyTypeStr.c_str());

            if (!is_properties_version1)
            {
                if (propertyName)
                    fprintf(out, "  %s", propertyName);
                else
                    fprintf(out, "  Unknown Property Name");

                if (type)
                    fprintf(out, "  %s", type);
                else
                    fprintf(out, "  Uknown Property Type 0x%s", propertyTypeStr.c_str());
            }
		}
	}
    else
	{
		fprintf(out, "%-9s %8s%s %10llu    ", obj_type_str,
			start_sector.c_str(),
			isMiniStream ? "m" : " ",
			(unsigned long long) e->stream_size);

		fprintf(out, "%s", indentstr.c_str());
		fprintf(out, "%s", name_utf8.c_str());

		GetPropertyIdAndType(name_utf8, propertyIdStr, propertyIdNumb, propertyTypeStr, propertyTypeNumb);

        if (propertyTypeNumb == PidTagConversationTopic)
            int deb = 1;

		propertyName = id2name(propertyIdNumb);
        type = id2type(propertyTypeStr.c_str());

        if (!is_properties_version1)
        {
            if (propertyName)
                fprintf(out, "  %s", propertyName);
            else
                fprintf(out, "  Unknown Property Name");

            if (type)
                fprintf(out, "  %s", type);
            else
                fprintf(out, "  Uknown Property Type 0x%s", propertyTypeStr.c_str());
        }
	}

    if (!isStream)
    {
        fprintf(out, "\n");
        goto end;
    }

    if ((propertyName == 0) || (type == 0))
    {
        fprintf(out, "\n");
        goto end;
    }

    int data_len = 0;
    char* data = cfbf_read_entry_data(cfbf, e, &data_len, errorText);

    if (!data)
    {
        fprintf(out, "\n");
        return 1;
    }

    int dumpMethod = DumpMethod(e, propertyIdNumb, propertyTypeNumb);

    bool isPropertySubStream = IsPropertySubStream(name_utf8);
    //if (isPropertySubStream)
        //dumpMethod = 1;

    if (dumpMethod == 1)
    {
        DumpBuffer(out, data, data_len);
    }
    else if (dumpMethod == 2)
    {
        std::string data_utf8 = UTF16ToUTF8((wchar_t*)data, data_len / 2);
        
        maxTextDumpLength = maxTextDumpLength;
        int maxTextDumpLength_sv = maxTextDumpLength;
        if (propertyIdNumb == PidTagTransportMessageHeaders)
            maxTextDumpLength = -1;

        DumpTextU16(out, data, data_len);

        maxTextDumpLength = maxTextDumpLength_sv;
    }
    else if (dumpMethod == 3)
    {
        maxTextDumpLength = maxTextDumpLength;
        int maxTextDumpLength_sv = maxTextDumpLength;
        if (propertyIdNumb == PidTagTransportMessageHeaders)
            maxTextDumpLength = -1;

        DumpTextU8(out, data, data_len);

        maxTextDumpLength = maxTextDumpLength_sv;
    }
    else if (dumpMethod == 4)
    {
        DumpRTF(out, data, data_len);
    }
    else
    {
        fprintf(out, "\n");
        int deb = 1;
    }

    free(data);

end:
    return retval;
}

OutlookMessage::~OutlookMessage()
{
    std::list<RecipientInfo>::iterator itR;
    for (itR = m_recipList.begin(); itR != m_recipList.end(); ++itR)
    {
        int deb = 1;
    }

    itR = m_recipList.begin();
    while (itR != m_recipList.end())
    {
        itR = m_recipList.erase(itR);
    }

    std::list<AttachmentInfo>::iterator itA;
    for (itA = m_attachList.begin(); itA != m_attachList.end(); ++itA)
    {
        int deb = 1;
    }

    itA = m_attachList.begin();
    while (itA != m_attachList.end())
    {
        itA = m_attachList.erase(itA);
    }
}

RecipientInfo* OutlookMessage::FindRecip(std::string& name)
{
    std::list<RecipientInfo>::iterator it;
    for (it = m_recipList.begin(); it != m_recipList.end(); ++it) {
        if (it->m_name == name)
            return &(*it);
    }
    return 0;
}

AttachmentInfo* OutlookMessage::FindAttach(std::string& name)
{
    std::list<AttachmentInfo>::iterator it;
    for (it = m_attachList.begin(); it != m_attachList.end(); ++it) {
        if (it->m_name == name)
            return &(*it);
    }
    return 0;
}

bool OutlookMessage::GetMessageClass(std::string& MessageClass)
{
    std::string errorText;
    MessageClass.clear();

    UINT16 propertyId = PidTagMessageClass;
    bool isRootParent = true;

    std::string value_utf8;
    UINT16 type = 0;
    bool valid = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, m_body.m_PidTagMessageClass, value_utf8, type, errorText);
    if (valid) {
        MessageClass.append(value_utf8);
        return true;
    }
    return false;
}

bool OutlookMessage::GetFromAddress(std::string& From)
{
    std::string errorText;
    From.clear();

    if (m_body.m_PidTagSenderName)
    {
        std::string value_utf8;
        UINT16 type = 0;
        bool valid = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, m_body.m_PidTagSenderName, value_utf8, type, errorText);
        if (valid) {
            From.append(value_utf8);
        }
    }

    if (m_body.m_PidTagSenderEmailAddress)
    {
        std::string value_utf8;
        UINT16 type = 0;
        bool valid = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, m_body.m_PidTagSenderEmailAddress, value_utf8, type, errorText);
        if (valid) {
            From.append(" <");
            From.append(value_utf8);
            From.append(">");
            return true;
        }
    }
    // or check if From is empty
    return false;
}

void OutlookMessage::GetToLists(std::string& To, std::string& CC, std::string& BCC)
{
    std::string errorText;
    std::list<RecipientInfo>::iterator it;
    for (it = m_recipList.begin(); it != m_recipList.end(); ++it)
    {
        UINT16 recipientType = 1;  // 1= PrimaryRecipient, 2 = CopyCarbonRecipient (CC), 3 = Bind CopyCarbonRecipient (BCC)
        Recipient& entry = it->m_recip;
        bool isRootParent = false;
        DWORD propertyId = PidTagRecipientType;
        UINT64 value;
        UINT16 type;
        bool valid = OutlookMessage::GetDirPropertyValueFixedLength(this->m_cfbf, isRootParent, entry.m_properties_version1_0, propertyId, value, type);
        if (!valid)
            recipientType = 1;  // PrimaryRecipient
        else
            recipientType = value & 0x03;

        std::string recip;
        if (entry.m_PidTagDisplayName)
        {
            std::string value_utf8;
            UINT16 type;
            bool valid = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, entry.m_PidTagDisplayName, value_utf8, type, errorText);
            recip.append(value_utf8);
        }
        if (entry.m_PidTagEmailAddress)
        {
            std::string value_utf8;
            UINT16 type = 0;
            bool valid = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, entry.m_PidTagEmailAddress, value_utf8, type, errorText);
            recip.append(" <");
            recip.append(value_utf8);
            recip.append(">");
        }

        std::string recipNotEncode = recip;

        CStringA wordEncodedRecip;
        int encodeType = 'Q';   // 'Q' == quoted (best for asci) 'B' ==  base64
        CStringA txt(recip.c_str(), recip.size());
        CStringA encodedTxt;
        TextUtilsEx::WordEncode(txt, encodedTxt, encodeType);

        recip.clear();
        recip.append((LPCSTR)encodedTxt, encodedTxt.GetLength());

        if (recipientType == 1)
        {
            if (!To.empty())
                To.append(",\r\n    ");
            To.append(recip);
        }
        else if (recipientType == 2)
        {
            if (!CC.empty())
                CC.append(",\r\n    ");
            CC.append(recip);
        }
        else if (recipientType == 3)
        {
            if (!BCC.empty())
                BCC.append(",\r\n    ");
            BCC.append(recip);
        }
        int deb = 1;;
    }
    int deb = 1;
}

bool OutlookMessage::GetSubject(std::string& Subject)
{
    std::string errorText;
    Subject.clear();

#if 0
    // Not needed. Prefix doesn't seem to be applicable
    if (m_body.m_PidTagSubjectPrefix)
    {
        std::string value_utf8;
        bool valid = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, m_body.m_PidTagSubjectPrefix, value_utf8);
        if (valid) {
            Subject.append(value_utf8);
        }
    }
#endif

    if (m_body.m_PidTagSubject)
    {
        std::string value_utf8;
        UINT16 type = 0;
        bool valid = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, m_body.m_PidTagSubject, value_utf8, type, errorText);
        if (valid) {
            Subject.append(value_utf8);
            return true;
        }
    }
    return false;
}

bool OutlookMessage::GetMessageCodepage(UINT64& nMessageCodepage, std::string& MessageCodepage)
{
    UINT16 propertyId = PidTagMessageCodepage;
    bool isRootParent = true;
    UINT16 type = 0;
    bool found = OutlookMessage::GetFixedLengthPropertValue(this->m_cfbf, isRootParent, m_body.m_properties_version1_0, propertyId, nMessageCodepage, type, MessageCodepage, false);
    return found;
}

bool OutlookMessage::GetInternetCodepage(UINT64& nInternetCodepage, std::string& InternetCodepage)
{
    UINT16 propertyId = PidTagInternetCodepage;
    bool isRootParent = true;
    UINT16 type = 0;
    bool found = OutlookMessage::GetFixedLengthPropertValue(this->m_cfbf, isRootParent, m_body.m_properties_version1_0, propertyId, nInternetCodepage, type, InternetCodepage, false);
    return found;
}


bool OutlookMessage::GetStoreSupportMask(UINT64& nStoreSupportMask, std::string& StoreSupportMask)
{
    UINT16 propertyId = PidTagStoreSupportMask;
    bool isRootParent = true;
    UINT16 type = 0;
    bool found = OutlookMessage::GetFixedLengthPropertValue(this->m_cfbf, isRootParent, m_body.m_properties_version1_0, propertyId, nStoreSupportMask, type, StoreSupportMask, true);
    return found;
}

bool OutlookMessage::GetNativeBody(UINT64& nNativeBody, std::string& NativeBody)
{
    UINT16 propertyId = PidTagNativeBody;
    bool isRootParent = true;
    UINT16 type = 0;
    bool found = OutlookMessage::GetFixedLengthPropertValue(this->m_cfbf, isRootParent, m_body.m_properties_version1_0, propertyId, nNativeBody, type, NativeBody, false);
    return found;
}

bool OutlookMessage::GetDate(std::string& Date)
{
    Date.clear();

    UINT16 propertyId = PidTagCreationTime;
    bool isRootParent = true;

    UINT64 value = 0;
    UINT16 type;
    bool found = OutlookMessage::GetDirPropertyValueFixedLength(this->m_cfbf, isRootParent, m_body.m_properties_version1_0, propertyId, value, type);
    if (!found)
    {
        propertyId = PidTagClientSubmitTime;
        found = OutlookMessage::GetDirPropertyValueFixedLength(this->m_cfbf, isRootParent, m_body.m_properties_version1_0, propertyId, value, type);
    }

    if (!found)
        return false;

    if (value > 0x8000000000000000LL)
        int deb = 1;

    FILETIME ft;
    ULARGE_INTEGER ularge;
    ularge.QuadPart = value; // your_uint64_value is the UINT64 you want to convert
    ft.dwLowDateTime = ularge.LowPart;
    ft.dwHighDateTime = ularge.HighPart;

    SYSTEMTIME sysTime;
    BOOL validConversion = FileTimeToSystemTime(&ft, &sysTime);
    UINT16 errCode = GetLastError();

    char dateBuff[128]; dateBuff[0] = 0;
    sprintf(dateBuff, "%d/%d/%d %d:%d\n", sysTime.wMonth, sysTime.wDay, sysTime.wYear, sysTime.wHour, sysTime.wSecond);
    Date.append(dateBuff, strlen(dateBuff) -1);

    if (!Date.empty())
        return true;
    else
        return false;
}

bool OutlookMessage::GetMessageId(std::string& MessageId)
{
    std::string errorText;
    MessageId.clear();

    UINT16 propertyId = PidTagInternetMessageId;
    bool isRootParent = true;

    std::string value_utf8;
    UINT16 type = 0;
    bool valid = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, m_body.m_PidTagInternetMessageId, value_utf8, type, errorText);
    if (valid) {
        MessageId.append(value_utf8);
        return true;
    }
    return false;
}

bool OutlookMessage::GetInReplyTo(std::string& InReplyTo)
{
    std::string errorText;
    InReplyTo.clear();

    UINT16 propertyId = PidTagInReplyToId;
    bool isRootParent = true;

    std::string value_utf8;
    UINT16 type = 0;
    bool valid = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, m_body.m_PidTagInReplyToId, value_utf8, type, errorText);
    if (valid) {
        InReplyTo.append(value_utf8);
        return true;
    }
    return false;
}

bool OutlookMessage::GetContentLocation(std::string& ContentLocation)
{
    std::string errorText;
    ContentLocation.clear();

    UINT16 propertyId = PidTagBodyContentLocation;
    bool isRootParent = true;

    std::string value_utf8;
    UINT16 type = 0;
    bool valid = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, m_body.m_PidTagBodyContentLocation, value_utf8, type, errorText);
    if (valid) {
        ContentLocation.append(value_utf8);
        return true;
    }
    return false;
}

bool OutlookMessage::GetContentId(std::string& ContentId)
{
    std::string errorText;
    ContentId.clear();

    std::string value_utf8;
    UINT16 type = 0;
    bool valid = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, m_body.m_PidTagBodyContentId, value_utf8, type, errorText);
    if (valid) {
        ContentId.append(value_utf8);
        return true;
    }
    return false;
}

bool OutlookMessage::GetPropertyString(struct cfbf* cfbf, DirEntry* entry, std::string& value_utf8, UINT16& type, std::string& errorText)
{
    value_utf8.clear();
    type = 0;

    if (!entry)
        return false;

    bool valid = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, entry, value_utf8, type, errorText);
    return valid;
}

bool OutlookMessage::GetFixedLengthPropertValue(struct cfbf* cfbf, bool isRootParent, DirEntry* entry, UINT16 propertyId, UINT64& nValue, UINT16& type, std::string& sValue,  bool hexFormat)
{
    sValue.clear();
    nValue = 0;

    bool found = OutlookMessage::GetDirPropertyValueFixedLength(this->m_cfbf, isRootParent, m_body.m_properties_version1_0, propertyId, nValue, type);
    if (found) {
        char valStr[32]; valStr[0] = 0;
        if (hexFormat)
            sprintf(valStr, "%08llX", nValue);
        else
            sprintf(valStr, "%llu", nValue);
        sValue.append(valStr);
    }
    else
    {
        sValue.clear();
        nValue = 0;
        type = 0;
    }

    return found;
}

int OutlookMessage::SetProperty(DirEntry* parent, DirEntry* entry)
{
    std::string parent_name_utf8;
    if (parent)
        parent_name_utf8 = UTF16ToUTF8((wchar_t*)parent->name, parent->name_length);
    std::string name_utf8 = UTF16ToUTF8((wchar_t*)entry->name, entry->name_length);
    if (strcmp(name_utf8.c_str(), "__properties_version1.0") == 0)
    {
        int deb = 1;
    }

    std::string propertyIdStr;
    int propertyIdNumb;
    std::string propertyTypeStr;
    int propertyTypeNumb;
    GetPropertyIdAndType(name_utf8, propertyIdStr, propertyIdNumb, propertyTypeStr, propertyTypeNumb);

    std::string bodyDirNamePrefix("Root Entry");
    std::string recipDirNamePrefix("__recip_version1.0_");
    std::string attachDirNamePrefix("__attach_version1.0_");
    if ((strncmp(parent_name_utf8.c_str(), bodyDirNamePrefix.c_str(), bodyDirNamePrefix.length()) == 0) ||
        (parent_name_utf8.compare("__substg1.0_3701000D") == 0))
    {
        m_body.SetProperty(propertyIdNumb, propertyTypeNumb, entry);
    }
    else if (strncmp(parent_name_utf8.c_str(), recipDirNamePrefix.c_str(), recipDirNamePrefix.length()) == 0)
    {
        RecipientInfo* recipInfo = FindRecip(parent_name_utf8);
        if (recipInfo == 0)
        {
            RecipientInfo rInfo;
           
            rInfo.m_name = parent_name_utf8;
            memset(&rInfo.m_recip, 0, sizeof(Recipient));
            m_recipList.push_back(rInfo);
            recipInfo = FindRecip(parent_name_utf8);

        }
        recipInfo->m_recip.SetProperty(propertyIdNumb, propertyTypeNumb, entry);
    }
    else if (strncmp(parent_name_utf8.c_str(), attachDirNamePrefix.c_str(), attachDirNamePrefix.length()) == 0)
    {
        AttachmentInfo* attachInfo = FindAttach(parent_name_utf8);
        if (attachInfo == 0)
        {
            AttachmentInfo aInfo;
            aInfo.m_name = parent_name_utf8;
            memset(&aInfo.m_attach, 0, sizeof(Attachment));
            m_attachList.push_back(aInfo);
            attachInfo = FindAttach(parent_name_utf8);
        }
        attachInfo->m_attach.SetProperty(propertyIdNumb, propertyTypeNumb, entry);
    }
    else
        int deb = 1;

    return 1;
}

int Body::SetProperty(int propertIdNumb, int propertTypeNumb, DirEntry* entry)
{
    switch (propertIdNumb)
    {
    case  PidTagMessageClass: { m_PidTagMessageClass = entry; } break;
    case  PidTagMessageFlags: { m_PidTagMessageFlags = entry; } break;
    case  PidTagBodyHtml: { m_PidTagBodyHtml = entry; } break;
    case  PidTagRtfCompressed: { m_PidTagRtfCompressed = entry; } break;
    case  PidTagNativeBody: { m_PidTagNativeBody = entry; } break;
        //case  PidTagNativeBody: { m_PidTagNativeBody = entry; } break;
    case  PidTagBody: { m_PidTagBody = entry; } break;
        //case  PidTagRtfCompressed: { m_PidTagRtfCompressed = entry; } break;
    case  PidTagInReplyToId: { m_PidTagInReplyToId = entry; } break;
    case  PidTagDisplayName: { m_PidTagDisplayName = entry; } break;
    case  PidTagRecipientDisplayName: { m_PidTagRecipientDisplayName = entry; } break;
    case  PidTagAddressType: { m_PidTagAddressType = entry; } break;
    case  PidTagEmailAddress: { m_PidTagEmailAddress = entry; } break;
    case  PidTagSmtpAddress: { m_PidTagSmtpAddress = entry; } break;
    case  PidTagDisplayBcc: { m_PidTagDisplayBcc = entry; } break;
    case  PidTagDisplayCc: { m_PidTagDisplayCc = entry; } break;
    case  PidTagDisplayTo: { m_PidTagDisplayTo = entry; } break;
    case  PidTagOriginalDisplayBcc: { m_PidTagOriginalDisplayBcc = entry; } break;
    case  PidTagOriginalDisplayCc: { m_PidTagOriginalDisplayCc = entry; } break;
    case  PidTagOriginalDisplayTo: { m_PidTagOriginalDisplayTo = entry; } break;
    case  PidTagRecipientType: { m_PidTagRecipientType = entry; } break;
    case  PidTagSenderName: { m_PidTagSenderName = entry; } break;
    case  PidTagSenderEmailAddress: { m_PidTagSenderEmailAddress = entry; } break;
    case  PidTagSenderSmtpAddress: { m_PidTagSenderSmtpAddress = entry; } break;
    case  PidTagSentRepresentingSmtpAddress: { m_PidTagSentRepresentingSmtpAddress = entry; } break;
    case  PidTagMessageSize: { m_PidTagMessageSize = entry; } break;
    case  PidTagMessageRecipients: { m_PidTagMessageRecipients = entry; } break;
    case  PidTagMessageAttachments: { m_PidTagMessageAttachments = entry; } break;
    case  PidTagHasAttachments: { m_PidTagHasAttachments = entry; } break;
    case  PidTagSubject: { m_PidTagSubject = entry; } break;
    case  PidTagSubjectPrefix: { m_PidTagSubjectPrefix = entry; } break;
    case  PidTagOriginalSubject: { m_PidTagOriginalSubject = entry; } break;
    case  PidTagNormalizedSubject: { m_PidTagNormalizedSubject = entry; } break;
    case  PidTagTransportMessageHeaders: { m_PidTagTransportMessageHeaders = entry; } break;
    case  PidTagLanguage: { m_PidTagLanguage = entry; } break;
    case  PidTagStoreSupportMask: { m_PidTagStoreSupportMask = entry; } break;
    case  PidTagMessageLocaleId: { m_PidTagMessageLocaleId = entry; } break;
    case  PidTagMessageCodepage: { m_PidTagMessageCodepage = entry; } break;
    case  PidTagInternetCodepage: { m_PidTagInternetCodepage = entry; } break;
    case  PidTagInternetMessageId: { m_PidTagInternetMessageId = entry; } break;
    default:
    {
        std::string name_utf8 = UTF16ToUTF8((wchar_t*)entry->name, entry->name_length);
        if (strcmp(name_utf8.c_str(), "__properties_version1.0") == 0)
        {
            m_properties_version1_0 = entry;
        }
        int deb = 1;
    }
    }
    return 1;
}

int Attachment::SetProperty(int propertIdNumb, int propertTypeNumb, DirEntry* entry)
{
    switch (propertIdNumb)
    {
    //case  PidTagAttachDataObject: { m_PidTagAttachDataObject = entry; } break;
    case  PidTagAttachMethod: { m_PidTagAttachMethod = entry; } break;
    case  PidTagAttachFilename: { m_PidTagAttachFilename = entry; } break;
    case  PidTagAttachMimeTag: { m_PidTagAttachMimeTag = entry; } break;
    case  PidTagAttachEncoding: { m_PidTagAttachEncoding = entry; } break;
    case  PidTagAttachExtension: { m_PidTagAttachExtension = entry; } break;
    case  PidTagAttachLongFilename: { m_PidTagAttachLongFilename = entry; } break;
        //case  PidTagAttachDataBinary: { m_PidTagAttachDataBinary = entry; } break;
    case  PidTagAttachSize: { m_PidTagAttachSize = entry; } break;
    case  PidTagLanguage: { m_PidTagLanguage = entry; } break;
    case  PidTagDisplayName: { m_PidTagDisplayName = entry; } break;
    case  PidTagAttachContentId: { m_PidTagAttachContentId = entry; } break;
    default: 
        if (propertIdNumb == PidTagAttachDataObject)  // note that PidTagAttachDataObject = 0x3701 and PidTagAttachDataBinary = 0x3701
        {
            if (propertTypeNumb == PtypObject)
                m_PidTagAttachDataObject = entry;
            else
                m_PidTagAttachDataBinary = entry;
        }
        else
        {
            std::string name_utf8 = UTF16ToUTF8((wchar_t*)entry->name, entry->name_length);
            if (strcmp(name_utf8.c_str(), "__properties_version1.0") == 0)
            {
                m_properties_version1_0 = entry;
            }
        }

    }
    return 1;
}

int Recipient::SetProperty(int propertIdNumb, int propertTypeNumb, DirEntry* entry)
{
    switch (propertIdNumb)
    {
    case  PidTagRecipientType: { m_PidTagRecipientType = entry; } break;
    case  PidTagEmailAddress: { m_PidTagEmailAddress = entry; } break;
    case  PidTagDisplayName: { m_PidTagDisplayName = entry; } break;
    case  PidTagAddressType: { m_PidTagAddressType = entry; } break;
    case  PidTagSmtpAddress: { m_PidTagSmtpAddress = entry; } break;
    case  PidTagRecipientDisplayName: { m_PidTagRecipientDisplayName = entry; } break;
    case  PidTagLanguage: { m_PidTagLanguage = entry; } break;
    case  PidTagRecipientFlags: { m_PidTagRecipientFlags = entry; } break;
    default:
    {
        std::string name_utf8 = UTF16ToUTF8((wchar_t*)entry->name, entry->name_length);
        if (strcmp(name_utf8.c_str(), "__properties_version1.0") == 0)
        {
            m_properties_version1_0 = entry;
        }
        int deb = 1;
    }
    }
    return 1;
}

void PrintProperty(struct cfbf* cfbf, int level, DirEntry* entry)
{
    std::string errorText;
    if (!entry)
        return;

    int indent = (level - 1) * 4;

    FILE* out = cfbf->out;

    std::string name_utf8 = UTF16ToUTF8((wchar_t*)entry->name, entry->name_length);

    std::string propertyIdStr;
    int propertyIdNumb;
    std::string propertyTypeStr;
    int propertyTypeNumb;
    GetPropertyIdAndType(name_utf8, propertyIdStr, propertyIdNumb, propertyTypeStr, propertyTypeNumb);

    char* obj_type_str = "stream";

    int isMiniStream = 0;

    if ((unsigned long)entry->start_sector == CFBF_END_OF_CHAIN) // == 4294967294ULL or 0xFFFFFFFE
    {
        fprintf(out, "%-8s %8s%s %10llu    ", obj_type_str,
            "FFFFFFFE",
            isMiniStream ? "m" : " ",
            (unsigned long long) entry->stream_size);
    }
    else
    {
        fprintf(out, "%-8s %8lu%s %10llu    ", obj_type_str,
            (unsigned long)entry->start_sector,
            isMiniStream ? "m" : " ",
            (unsigned long long) entry->stream_size);
    }

    fprintf(out, "%*s", indent, "");
    fprintf(out, "%s", name_utf8.c_str());

    bool isPropertySubStream = IsPropertySubStream(name_utf8);
    char* propertyName = id2name(propertyIdNumb);
    char* type = id2type(propertyTypeStr.c_str());

    if (!isPropertySubStream)
	{
		if (propertyName)
			fprintf(out, "  %s", propertyName);
		else
			fprintf(out, "  Uknown Property Name");

		if (type)
			fprintf(out, "  %s", type);
		else
			fprintf(out, "  Uknown Property Type 0x%s", propertyTypeStr.c_str());
	}

    if (propertyName == 0)
    {
        fprintf(out, "\n");
        return;
    }

    int data_len = 0;
    char* data = cfbf_read_entry_data(cfbf, entry, &data_len, errorText);

    if (!data)
    {
        fprintf(out, "\n");
        return;
    }

	int dumpMethod = DumpMethod(entry, propertyIdNumb, propertyTypeNumb);

	//if (isPropertySubStream)
		//dumpMethod = 1;

    if (dumpMethod == 1)
    {
        DumpBuffer(out, data, data_len);
    }
    else if (dumpMethod == 2)
    {
        std::string data_utf8 = UTF16ToUTF8((wchar_t*)data, data_len / 2);

        maxTextDumpLength = maxTextDumpLength;
        int maxTextDumpLength_sv = maxTextDumpLength;
        if (propertyIdNumb == PidTagTransportMessageHeaders)
            maxTextDumpLength = -1;

        DumpTextU16(out, data, data_len);

        maxTextDumpLength = maxTextDumpLength_sv;
    }
    else if (dumpMethod == 3)
    {
        maxTextDumpLength = maxTextDumpLength;
        int maxTextDumpLength_sv = maxTextDumpLength;
        if (propertyIdNumb == PidTagTransportMessageHeaders)
            maxTextDumpLength = -1;

        DumpTextU8(out, data, data_len);

        maxTextDumpLength = maxTextDumpLength_sv;
    }
    else if (dumpMethod == 4)
    {
        DumpRTF(out, data, data_len);
    }
    else
    {
        fprintf(out, "\n");
        int deb = 1;
    }

	free(data);
}

void OutlookMessage::PrintDirProperties(struct cfbf* cfbf, int level, DirEntry* entry)
{
    std::string errorText;
    if (entry == 0)
        return;

    FILE* out = cfbf->out;

    std::string name_utf8 = UTF16ToUTF8((wchar_t*)entry->name, entry->name_length);

    if (strcmp(name_utf8.c_str(), "__properties_version1.0") == 0)
    {
        int data_len = 0;
        char* data = cfbf_read_entry_data(cfbf, entry, &data_len, errorText);
        if (!data)
        {
            fprintf(out, "\n");
            return;
        }

        // Print Property Array of  Entries

        int PropertyEntrySize = sizeof(PropertyEntry);

        int arrayLength = 0;
        char* entry_data = 0;
        if (level == 0)
        {
            int topHdrSize = sizeof(PropertyHeaderTop);
            arrayLength = entry->stream_size - topHdrSize;
            PropertyHeaderTop* hdr = (PropertyHeaderTop*)data;
            entry_data = data + topHdrSize;
        }
        else
        {
            int arHdrSize = sizeof(PropertyHeaderOfAttachmentAndRecipient);
            arrayLength = entry->stream_size - arHdrSize;
            PropertyHeaderOfAttachmentAndRecipient* hdr = (PropertyHeaderOfAttachmentAndRecipient*)data;
            entry_data = data + arHdrSize;
        }
        int entryCount = arrayLength / sizeof(PropertyEntry);

        int i = 0;
        for (i = 0; i < entryCount; i++)
        {
            PropertyEntry* e = (PropertyEntry*)(entry_data) + i;
            int tag = e->PropertyTag >> 16;
            int type = e->PropertyTag & 0xFFFF;

            char* propertyName = id2name(tag);
            if (propertyName == 0)
                propertyName = "Unknown";

            char typeStr[32]; typeStr[0] = 0;
            sprintf(typeStr, "%04X", type);
            char* propertyTypeName = id2type(typeStr);
            if (propertyTypeName == 0)
                propertyTypeName = "Unknown";

            fprintf(out, "%08x %04x %s %s %llu\n", e->PropertyTag, e->Flags, propertyName, propertyTypeName, e->Value);
        }
        free(data);
    }
}
bool OutlookMessage::GetDirPropertyValueFixedLength(struct cfbf* cfbf, bool isRootParent, DirEntry* entry, UINT16 propertyId, UINT64 &value, UINT16 &type)
{
    std::string errorText;
    value = 0;
    type = 0;
    if (entry == 0)
        return false;

    FILE* out = cfbf->out;

    std::string name_utf8 = UTF16ToUTF8((wchar_t*)entry->name, entry->name_length);

    if (strcmp(name_utf8.c_str(), "__properties_version1.0") == 0)
    {
        int data_len = 0;
        char* data = cfbf_read_entry_data(cfbf, entry, &data_len, errorText);
        if (!data)
        {
            return false;
        }

        // Print Property Array of  Entries

        int PropertyEntrySize = sizeof(PropertyEntry);

        int arrayLength = 0;
        char* entry_data = 0;
        if (isRootParent)
        {
            int topHdrSize = sizeof(PropertyHeaderTop);
            arrayLength = entry->stream_size - topHdrSize;
            PropertyHeaderTop* hdr = (PropertyHeaderTop*)data;
            entry_data = data + topHdrSize;
        }
        else
        {
            int arHdrSize = sizeof(PropertyHeaderOfAttachmentAndRecipient);
            arrayLength = entry->stream_size - arHdrSize;
            PropertyHeaderOfAttachmentAndRecipient* hdr = (PropertyHeaderOfAttachmentAndRecipient*)data;
            entry_data = data + arHdrSize;
        }
        int entryCount = arrayLength / sizeof(PropertyEntry);

        int i = 0;
        for (i = 0; i < entryCount; i++)
        {
            PropertyEntry* e = (PropertyEntry*)(entry_data)+i;
            UINT16 prop_id = e->PropertyTag >> 16;
            UINT16 prop_type = e->PropertyTag & 0xFFFF;

            if (prop_id == propertyId)
            {
                bool isFixedType = isFixedLengthType(prop_type);
                if (isFixedType)
                {
                    value = e->Value;
                    type = prop_type;
                    free(data);
                    return true;
                }
            }
            //fprintf(out, "%08x %04x %s %s %llu\n", e->PropertyTag, e->Flags, propertyName, propertyTypeName, e->Value);
        }
        free(data);
    }
    return false;
}


bool OutlookMessage::GetStreamDirEntryValueString(struct cfbf* cfbf, DirEntry* entry, std::string& value, UINT16 &type, std::string& errorText)
{
    value.clear();
    type = 0;
    if (!entry)
        return false;

    FILE* out = stdout;

    std::string name_utf8 = UTF16ToUTF8((wchar_t*)entry->name, entry->name_length);

    std::string propertyIdStr;
    int propertyIdNumb;
    std::string propertyTypeStr;
    int propertyTypeNumb;
    GetPropertyIdAndType(name_utf8, propertyIdStr, propertyIdNumb, propertyTypeStr, propertyTypeNumb);

    type = propertyTypeNumb;

    char* propertyName = id2name(propertyIdNumb);

    if ((propertyTypeNumb != PtypString)  && (propertyTypeNumb != PtypString8))
    {
        return false;
    }

    int data_len = 0;
    char* data = cfbf_read_entry_data(cfbf, entry, &data_len, errorText);
    if (!data)
    {
        return false;
    }

    if (propertyTypeNumb == PtypString)
        value = UTF16ToUTF8((wchar_t*)data, data_len / 2);
    else
        value.append(data, data_len);

    free(data);
    return true;
}

char* OutlookMessage::GetStreamDirEntryValue(struct cfbf* cfbf, DirEntry* entry, UINT16& type, int& data_len, std::string& errorText)
{
    type = 0;
    data_len = 0;
    if (!entry)
        return 0;

    std::string name_utf8 = UTF16ToUTF8((wchar_t*)entry->name, entry->name_length);

    std::string propertyIdStr;
    int propertyIdNumb;
    std::string propertyTypeStr;
    int propertyTypeNumb;
    GetPropertyIdAndType(name_utf8, propertyIdStr, propertyIdNumb, propertyTypeStr, propertyTypeNumb);

    char* propertyName = id2name(propertyIdNumb);
    char* propertyTypeName = id2type(propertyTypeStr.c_str());

    type = propertyTypeNumb;
    // Do we need compare with type defined in __properties_version1  ??
    char* data = cfbf_read_entry_data(cfbf, entry, &data_len, errorText);  //  it should return unsigned char*, fix it
    return data;
}

void Body::Print(struct cfbf* cfbf, int level)
{
    PrintProperty(cfbf, level, m_PidTagMessageFlags);
    PrintProperty(cfbf, level, m_PidTagMessageClass);
    PrintProperty(cfbf, level, m_PidTagBodyHtml);
    PrintProperty(cfbf, level, m_PidTagRtfCompressed);
    PrintProperty(cfbf, level, m_PidTagNativeBody);
    //PrintProperty(cfbf, level, m_PidTagNativeBody);
    PrintProperty(cfbf, level, m_PidTagBody);
    //PrintProperty(cfbf, level, m_PidTagRtfCompressed);
    PrintProperty(cfbf, level, m_PidTagInReplyToId);
    PrintProperty(cfbf, level, m_PidTagDisplayName);
    PrintProperty(cfbf, level, m_PidTagRecipientDisplayName);
    PrintProperty(cfbf, level, m_PidTagAddressType);
    PrintProperty(cfbf, level, m_PidTagEmailAddress);
    PrintProperty(cfbf, level, m_PidTagSmtpAddress);
    PrintProperty(cfbf, level, m_PidTagDisplayBcc);
    PrintProperty(cfbf, level, m_PidTagDisplayCc);
    PrintProperty(cfbf, level, m_PidTagDisplayTo);
    PrintProperty(cfbf, level, m_PidTagOriginalDisplayBcc);
    PrintProperty(cfbf, level, m_PidTagOriginalDisplayCc);
    PrintProperty(cfbf, level, m_PidTagOriginalDisplayTo);
    PrintProperty(cfbf, level, m_PidTagRecipientType);
    PrintProperty(cfbf, level, m_PidTagSenderName);
    PrintProperty(cfbf, level, m_PidTagSenderEmailAddress);
    PrintProperty(cfbf, level, m_PidTagSenderSmtpAddress);
    PrintProperty(cfbf, level, m_PidTagSentRepresentingSmtpAddress);
    PrintProperty(cfbf, level, m_PidTagMessageSize);
    PrintProperty(cfbf, level, m_PidTagMessageRecipients);
    PrintProperty(cfbf, level, m_PidTagMessageAttachments);
    PrintProperty(cfbf, level, m_PidTagHasAttachments);
    PrintProperty(cfbf, level, m_PidTagSubject);
    PrintProperty(cfbf, level, m_PidTagSubjectPrefix);
    PrintProperty(cfbf, level, m_PidTagOriginalSubject);
    PrintProperty(cfbf, level, m_PidTagNormalizedSubject);
    PrintProperty(cfbf, level, m_PidTagTransportMessageHeaders);
    PrintProperty(cfbf, level, m_PidTagLanguage);
    PrintProperty(cfbf, level, m_PidTagStoreSupportMask);
    PrintProperty(cfbf, level, m_PidTagMessageLocaleId);
    PrintProperty(cfbf, level, m_PidTagMessageCodepage);
    PrintProperty(cfbf, level, m_PidTagInternetCodepage);
    PrintProperty(cfbf, level, m_PidTagInternetMessageId);
    //
    PrintProperty(cfbf, level, m_properties_version1_0);
    OutlookMessage::PrintDirProperties(cfbf, 0, m_properties_version1_0);
}

void Recipient::Print(struct cfbf* cfbf, int level)
{
    PrintProperty(cfbf, level, m_PidTagRecipientType);
    PrintProperty(cfbf, level, m_PidTagEmailAddress);
    PrintProperty(cfbf, level, m_PidTagDisplayName);
    PrintProperty(cfbf, level, m_PidTagAddressType);
    PrintProperty(cfbf, level, m_PidTagSmtpAddress);
    PrintProperty(cfbf, level, m_PidTagRecipientDisplayName);
    PrintProperty(cfbf, level, m_PidTagLanguage);
    PrintProperty(cfbf, level, m_PidTagRecipientFlags);
    //
    PrintProperty(cfbf, level, m_properties_version1_0);
    OutlookMessage::PrintDirProperties(cfbf, 1, m_properties_version1_0);
}

int Recipient::GetRecipientType(struct cfbf* cfbf, DirEntry* m_properties_version1_0)
{
    // Skip header and iterate array to find PidTagRecipientFlags
    // Root Message handling differs from handling embeded Message
    return 1;
}

void Attachment::Print(struct cfbf* cfbf, int level)
{
    PrintProperty(cfbf, level, m_PidTagAttachDataObject);
    PrintProperty(cfbf, level, m_PidTagAttachMethod);
    PrintProperty(cfbf, level, m_PidTagAttachFilename);
    PrintProperty(cfbf, level, m_PidTagAttachMimeTag);
    PrintProperty(cfbf, level, m_PidTagAttachEncoding);
    PrintProperty(cfbf, level, m_PidTagAttachExtension);
    PrintProperty(cfbf, level, m_PidTagAttachLongFilename);
    PrintProperty(cfbf, level, m_PidTagAttachDataBinary);
    PrintProperty(cfbf, level, m_PidTagAttachSize);
    PrintProperty(cfbf, level, m_PidTagLanguage);
    PrintProperty(cfbf, level, m_PidTagDisplayName);
    PrintProperty(cfbf, level, m_PidTagAttachContentId);
    if (this->m_OutlookMessage) {
        this->m_OutlookMessage->Print();
    }
    //
    PrintProperty(cfbf, level, m_properties_version1_0);
    OutlookMessage::PrintDirProperties(cfbf, 1, m_properties_version1_0);
}

void OutlookMessage::Print()
{
    m_body.Print(m_cfbf, 1);

    FILE* out = m_cfbf->out;

    std::list<AttachmentInfo>::iterator ait;
    for (ait = m_attachList.begin(); ait != m_attachList.end(); ++ait) {
        fprintf(out, "[%s]  !!!!!!!!!!!!!!!\n", ait->m_name.c_str());
        ait->m_attach.Print(m_cfbf, 2);
    }

    std::list<RecipientInfo>::iterator rit;
    for (rit = m_recipList.begin(); rit != m_recipList.end(); ++rit) {
        fprintf(out, "[%s]  !!!!!!!!!!!!!!!\n", rit->m_name.c_str());
       rit->m_recip.Print(m_cfbf, 2);
    }
}

int OutlookMessage::ParseMsg(struct cfbf* cfbf, Parse_Outlook_Msg _ParseOutlookMsg, OutlookMsgHelper& msgHelper, std::string& errorText)
{
    int ret = cfbf_walk_dir_tree(cfbf, _ParseOutlookMsg, &msgHelper, errorText);
    if (ret < 0) {
        return 1;
    }
    else
        return 0;
}

int OutlookMessage::Msg2Eml(std::string& errorText)
{
    std::string msg_utf8;
    int retM2E = OutlookMessage::Msg2Eml(msg_utf8, errorText);
    return retM2E;
}

int OutlookMessage::Msg2Eml(std::string& hdr_utf8, std::string& errorText)
{
    FILE* out = logFile;
    int boundaryLength = 48;
    std::string mixedBoundary("mixed_");
    mixedBoundary.append(GenerateRandomString(boundaryLength));
    std::string relatedBoundary("related_");
    relatedBoundary.append(GenerateRandomString(boundaryLength));
    std::string alternativeBoundary("alternative_");
    alternativeBoundary.append(GenerateRandomString(boundaryLength));

    bool hasMessageHeaders = m_body.m_PidTagTransportMessageHeaders ? true : false;
    bool hasPlainBody = m_body.m_PidTagBody ? true : false;
    bool hasHtmlBody = m_body.m_PidTagBodyHtml ? true : false;
    bool hasRtfBody = m_body.m_PidTagRtfCompressed ? true : false;

    // below may not be needed in case we already have the mime header in .msg
    // it might usefull to verify below with mime header already in .msg
    std::string To;
    std::string CC;
    std::string BCC;
    OutlookMessage::GetToLists(To, CC, BCC);

    // Do we care if property is missing or empty property string will work ??
    std::string Subject;
    bool foundSubject = OutlookMessage::GetSubject(Subject);

    std::string From;
    bool foundFromAddress = OutlookMessage::GetFromAddress(From);

    std::string Date;
    bool foundDate = OutlookMessage::GetDate(Date);
    //long MS_EPOCH_OFFSET = 11644473600000L;

    std::string MessageId;
    bool foundMessageId = OutlookMessage::GetMessageId(MessageId);

    std::string InReplyTo;
    bool foundInReplyTo = OutlookMessage::GetInReplyTo(InReplyTo);

    std::string ContentLocation;
    bool foundContentLocation = OutlookMessage::GetContentLocation(ContentLocation);

    std::string ContentId;
    bool foundContentId = OutlookMessage::GetContentId(ContentId);

    std::string MessageCodepage;
    UINT64 nMessageCodepage = 0;
    bool foundMessageCodePage = OutlookMessage::GetMessageCodepage(nMessageCodepage, MessageCodepage);

    std::string InternetCodepage;
    UINT64 nInternetCodepage = 0;
    bool foundInternetCodepage = OutlookMessage::GetInternetCodepage(nInternetCodepage, InternetCodepage);

    UINT64 nStoreSupportMask = 0;
    std::string StoreSupportMask;

    bool foundStoreSupportMask = OutlookMessage::GetStoreSupportMask(nStoreSupportMask, StoreSupportMask);

    bool STORE_UNICODE_OK = 0x00040000 & nStoreSupportMask ? true : false;

    // PidTagNativeBody  recommended body part. 0 = undefined  1 == plain 2 == rtf 3 = HTML 4 = clear-signed body
    UINT64 nNativeBody = 0;
    std::string NativeBody;
    bool foundNativeBody = OutlookMessage::GetNativeBody(nNativeBody, NativeBody);

    bool isMultipartSigned = false;
    std::string MessageClass;
    bool found_MessageClass = OutlookMessage::GetMessageClass(MessageClass);
    if (!found_MessageClass)
    {
        ;
    }
    else
    {
        if (MessageClass.compare("IPM.Note") == 0)
        {
            ;
        }
        else if (MessageClass.compare("IPM.Note.SMIME.MultipartSigned") == 0)
        {
            // special processing
            // ignore body blocks and merge headers with attachment block
            isMultipartSigned = true;  
        }
        else
        {
            ; // MyMessageBox(); // Mail class:  Supported Mail Class are: 
        }

    }


    // PidTagMessageCcMe Property
    // PidTagMessageRecipientMe Property
    // PidTagReplyRecipientEntries
    // PidTagReplyRecipientNames
    //   MIME header values, the encoding specified in [RFC2047] MUST be used to encode Unicode 


#if 0
    fprintf(stdout, 
        "From: %s\nDate: %s\nTo: %s\nCC: %s\nBCC: %s\nSubject: %s\n"
        "message-id: %s\nin-reply-to: %s\ncontent-location: %s\ncontent-id: %s\n"
        "MessageCodepage: %s\nInternetCodepage: %s\nStoreSupportMask: 0x%s STORE_UNICODE_OK: %d\nNativeBody: %s\n", 
        From.c_str(), Date.c_str(), To.c_str(), CC.c_str(), BCC.c_str(), Subject.c_str(),
        MessageId.c_str(), InReplyTo.c_str(), ContentLocation.c_str(), ContentId.c_str(),
        MessageCodepage.c_str(), InternetCodepage.c_str(), StoreSupportMask.c_str(), STORE_UNICODE_OK, NativeBody.c_str());
#endif


    if (!hasMessageHeaders)
    {
        // hope this is never the case or very rare but forces us to do extra more tricky work
        // will need to create mime header from .msg content other than m_PidTagTransportMessageHeaders

        int encodeType = 'Q';   // 'Q' == quoted (best for asci) 'B' ==  base64

        std::string wordEncodedFrom;
        CStringA txt(From.c_str(), From.size());
        CStringA encodedTxt;
        TextUtilsEx::WordEncode(txt, encodedTxt, encodeType);
        wordEncodedFrom.append((LPCSTR)txt, txt.GetLength());

        //InReplyTo  -> wordEncodedInReplyTo ; InReplyTo is likely comma separated list

        //hdr_utf8.append("\r\n"); to force refresh index
        hdr_utf8.append("Subject: ");
        hdr_utf8.append(Subject);
        hdr_utf8.append("\r\n");
        hdr_utf8.append("Date: ");
        hdr_utf8.append(Date);
        hdr_utf8.append("\r\n");
        hdr_utf8.append("From: ");
        hdr_utf8.append(wordEncodedFrom);
        hdr_utf8.append("\r\n");
        hdr_utf8.append("To: ");
        hdr_utf8.append(To);
        hdr_utf8.append("\r\n");
        hdr_utf8.append("Cc: "); 
        hdr_utf8.append(CC);
        hdr_utf8.append("\r\n");
        hdr_utf8.append("Bcc: "); 
        hdr_utf8.append(BCC);
        hdr_utf8.append("\r\n");
        hdr_utf8.append("In-Reply-To: "); 
        hdr_utf8.append(InReplyTo);
        hdr_utf8.append("\r\n");
        hdr_utf8.append("Message-Id: ");
        hdr_utf8.append(MessageId);
        hdr_utf8.append("\r\n");
        hdr_utf8.append("MIME-Version: 1.0");
        hdr_utf8.append("\r\n");

        int deb = 1;
    }
    else
    {
        DirEntry* entry = m_body.m_PidTagTransportMessageHeaders;
        int data_len = 0;
        char* data = cfbf_read_entry_data(m_cfbf, entry, &data_len, errorText);

        bool isNULTerminated = false;
        int nlCnt = 0;
        int i;
        for (i = (data_len-1); i >= 0; i--)
        {
            char c = data[i];
            if (c == '\0')
            {
                isNULTerminated = true;
                continue;
            }
            if ((c == '\n') || (c == '\r'))
            {
                continue;
            }
            else
                break;
        }

        data_len = i + 1;

        if (!data)
        {
            fprintf(out, "\n");
            m2eReturn();
            return 1;
        }

        if (!OutlookMessage::IsString8(data, data_len))
            hdr_utf8 = UTF16ToUTF8((wchar_t*)data, data_len / 2);
        else
            hdr_utf8.append(data, data_len);

        hdr_utf8.append("\r\n");

        MailHeader mailHdr;
        mailHdr.Load(hdr_utf8.c_str(), hdr_utf8.length());

        // Delete Content-type field and add later
        int lengthBegin = hdr_utf8.length();
        std::string field("Content-Type");
        std::string  deletedContentType = DeleteFieldFromHeader(hdr_utf8, field);
        const char* fld = deletedContentType.c_str();
        int fldlen = deletedContentType.length();
        size_t lengthEnd = hdr_utf8.length();

        field = "Content-Transfer-Encoding";
        std::string deletedContentTransferEncoding = DeleteFieldFromHeader(hdr_utf8, field);
        fld = deletedContentTransferEncoding.c_str();
        fldlen = deletedContentTransferEncoding.length();
        lengthEnd = hdr_utf8.length();

        if (!foundMessageId)
        {
            hdr_utf8.append("Message-Id: \r\n");
        }

        MailHeader mailHdr2;
        mailHdr2.Load(hdr_utf8.c_str(), hdr_utf8.length());

        free(data);
        int deb = 1;
    }

    if (isMultipartSigned)
    {
        // process attachment
        // it should contain mail bdy minus header

        if (m_attachList.size())
        {
            std::list<AttachmentInfo>::iterator it;
            for (it = m_attachList.begin(); it != m_attachList.end(); ++it)
            {
                std::string str_PidTagAttachLongFilename;
                UINT16 type_PidTagAttachLongFilename;
                bool found_PidTagAttachLongFilename = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, it->m_attach.m_PidTagAttachLongFilename, str_PidTagAttachLongFilename, type_PidTagAttachLongFilename, errorText);

                std::string str_PidTagDisplayName;
                UINT16 type_PidTagDisplayName;
                bool found_PidTagDisplayName = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, it->m_attach.m_PidTagDisplayName, str_PidTagDisplayName, type_PidTagDisplayName, errorText);

                std::string str_PidTagAttachExtension;
                UINT16 type_PidTagAttachExtension;
                bool found_PidTagAttachExtension = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, it->m_attach.m_PidTagAttachExtension, str_PidTagAttachExtension, type_PidTagAttachExtension, errorText);

                std::string str_PidTagAttachMimeTag;
                UINT16 type_PidTagAttachMimeTag;
                bool found_PidTagAttachMimeTag = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, it->m_attach.m_PidTagAttachMimeTag, str_PidTagAttachMimeTag, type_PidTagAttachMimeTag, errorText);

                std::string str_PidTagAttachFilename;
                UINT16 type_PidTagAttachFilename;
                bool found_PidTagAttachFilename = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, it->m_attach.m_PidTagAttachFilename, str_PidTagAttachFilename, type_PidTagAttachFilename, errorText);

                std::string str_PidTagAttachContentId;
                UINT16 type_PidTagAttachContentId;
                bool found_PidTagAttachContentId = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, it->m_attach.m_PidTagAttachContentId, str_PidTagAttachContentId, type_PidTagAttachContentId, errorText);

                UINT16 propertyId = PidTagAttachFlags;
                bool isRootParent = false;
                UINT16 type_PidTagAttachFlags = 0;
                std::string str_PidTagAttachFlags;
                UINT64 nPidTagAttachFlags;
                bool found_PidTagAttachFlags = OutlookMessage::GetFixedLengthPropertValue(this->m_cfbf, isRootParent, m_body.m_properties_version1_0, propertyId,
                    nPidTagAttachFlags, type_PidTagAttachFlags, str_PidTagAttachFlags, false);

                //  nPidTagAttachFlags == 4 indicates inline in PidTagBodyHtml

                //m_PidTagAttachEncoding
                isRootParent = false;
                DWORD propertyId_PidTagAttachEncoding = PidTagAttachEncoding;
                UINT64 value_PidTagAttachEncoding;
                UINT16 type_PidTagAttachEncoding;
                bool found_PidTagAttachEncoding = OutlookMessage::GetDirPropertyValueFixedLength(this->m_cfbf, isRootParent, it->m_attach.m_properties_version1_0, propertyId_PidTagAttachEncoding,
                    value_PidTagAttachEncoding, type_PidTagAttachEncoding);

                UINT16 type = 0;
                int data_len;
                char* data = GetStreamDirEntryValue(m_cfbf, it->m_attach.m_PidTagAttachDataBinary, type, data_len, errorText);

                if (!data)
                {
                    fprintf(out, "\n");
                    m2eReturn();
                    return 1;
                }

                bool isString8 = OutlookMessage::IsString8(data, data_len);
                _ASSERTE(isString8);
                if (!isString8)
                {
                    int deb = 1;
                }

                hdr_utf8.append(data, data_len);

                free(data);
                int deb = 1;

            }
        }
        return -1;
    }

    // add one or two body blocks and one or more attachments.

    bool hasValidAttachments = false;
    if (m_attachList.size())
    {
        std::list<AttachmentInfo>::iterator it;
        for (it = m_attachList.begin(); it != m_attachList.end(); ++it)
        {
            if ((it->m_attach.m_PidTagAttachDataBinary) || (it->m_attach.m_PidTagAttachDataObject))
            {
                hasValidAttachments = true; // assume attachments are valid; optimisticly
            }
            int deb = 1;
        }
    }

    if (hasValidAttachments)
    {
        hdr_utf8.append("Content-Type: multipart/mixed;\r\n    boundary=");
        hdr_utf8.append(mixedBoundary);
        hdr_utf8.append("\r\n");

        hdr_utf8.append("\r\n\r\n"); // to make sure

        hdr_utf8.append("--");
        hdr_utf8.append(mixedBoundary);
        hdr_utf8.append("\r\n"); // to make sure
    }

    hdr_utf8.append("Content-Type: multipart/alternative;\r\n    boundary=");
    hdr_utf8.append(alternativeBoundary);
    hdr_utf8.append("\r\n");

    hdr_utf8.append("\r\n"); // to make sure

    hdr_utf8.append("\r\n--");
    hdr_utf8.append(alternativeBoundary);
    hdr_utf8.append("\r\n"); // to make sure

    // Preference order: HtmlBody, hasRtfBody, hasPlainBody or check PidTagNativeBody ??

    if (hasPlainBody)
    {
        DirEntry* entry = m_body.m_PidTagBody;
        int data_len = 0;
        char* data = cfbf_read_entry_data(m_cfbf, entry, &data_len, errorText);

        if (!data)
        {
            fprintf(out, "\n");
            m2eReturn();
            return 1;
        }

        if (foundInternetCodepage)
            ; // must map nInternetCodepage to charset
        else
            ; // ??

        if (hdr_utf8[hdr_utf8.length() - 1] != '\n')
            hdr_utf8.append("\r\n");

        bool isString8 = OutlookMessage::IsString8(data, data_len);

        // Optymistic guessing, may need to enhance once I learn more
        // May examine the entire text to guess charset ??
        UINT pageCode = 65001; // utf-8
        if (isString8 && (nInternetCodepage != 0))
            pageCode = nInternetCodepage;
        else
            ;   // or must quess somehow

        std::string charset;
        BOOL ret = TextUtilsEx::id2charset(pageCode, charset);

        hdr_utf8.append("Content-Type: text/plain; charset=\"");
        hdr_utf8.append(charset);
        hdr_utf8.append("\";");
        hdr_utf8.append("\r\nContent-Transfer-Encoding: quoted-printable");
        hdr_utf8.append("\r\n\r\n"); // to make sure

        const char* plain_data = data;
        int plain_data_len = data_len;

        std::string data_utf8;
        if (!isString8)
        {
            data_utf8 = UTF16ToUTF8((wchar_t*)data, data_len / 2);
            plain_data = data_utf8.c_str();
            plain_data_len = data_utf8.size();
        }

        CMimeCodeQP eQP;
        bool bEncoding = true;
        eQP.SetInput(plain_data, plain_data_len, bEncoding);
        int eLength = eQP.GetOutputLength();

        char* eBuffer = (char*)cfbf_malloc(eLength);
        if (eBuffer == NULL)
            return 1;
        int retlen = eQP.GetOutput((unsigned char*)eBuffer, eLength);
        if (retlen > 0)
            int deb = 1;
        else
            int deb = 1;

        hdr_utf8.append(eBuffer, retlen);

        free(data);
        free(eBuffer);

        //fprintf(stdout, "\n\n%s\n", "Found PlainBody");
    }
    else
        ;// return -1;


    bool bodyFound = false;
    if (hasHtmlBody && !bodyFound)
    {
        DirEntry* entry = m_body.m_PidTagBodyHtml;
        int data_len = 0;
        char* data = cfbf_read_entry_data(m_cfbf, entry, &data_len, errorText);
        if (!data)
        {
            fprintf(out, "\n");
            m2eReturn();
            return 1;
        }

        // data should not be UTF16
        if (OutlookMessage::IsString8(data, data_len))
        {
            std::string html_utf8(data, data_len);
            if (hdr_utf8[hdr_utf8.length() - 1] != '\n')
                hdr_utf8.append("\r\n");

            // Optymistic guessing, may need to enhance one I learn more
            // HTML body may have charset defined in the HTML text
            UINT pageCode = 65001;  //  urf-8
            if (nInternetCodepage != 0)
                pageCode = nInternetCodepage;
            else
                ;   // or must quess somehow

            std::string charset;
            BOOL ret = TextUtilsEx::id2charset(pageCode, charset);

            hdr_utf8.append("\r\n\r\n--");
            hdr_utf8.append(alternativeBoundary);
            hdr_utf8.append("\r\n"); // to make sure

            hdr_utf8.append("Content-Type: text/html; charset=\"");
            hdr_utf8.append(charset);
            hdr_utf8.append("\";");
            hdr_utf8.append("\r\nContent-Transfer-Encoding: quoted-printable");
            hdr_utf8.append("\r\n\r\n"); // to make sure   // ZMMM

            //hdr_utf8.append(html_utf8);

            CMimeCodeQP eQP;
            bool bEncoding = true;
            eQP.SetInput(html_utf8.c_str(), html_utf8.size(), bEncoding);
            int eLength = eQP.GetOutputLength();

            char* eBuffer = (char*)cfbf_malloc(eLength);
            if (eBuffer == NULL)
                return 1;
            int retlen = eQP.GetOutput((unsigned char*)eBuffer, eLength);
            if (retlen > 0)
                int deb = 1;
            else
                int deb = 1;

            hdr_utf8.append(eBuffer, retlen);

            free(data);
            free(eBuffer);

            bodyFound = true;
            //fprintf(stdout, "\n\n%s\n", "Found HtmlBody");
        }
        else
        {
            //fprintf(stdout, "\n\n%s\n", "Found but ignorted HtmlBody encoded as UTF16");
            ; // convert to utf8 ??
        }
    }

    if (hasRtfBody && !bodyFound)
    {
        DirEntry* entry = m_body.m_PidTagRtfCompressed;
        int data_len = 0;
        char* data = cfbf_read_entry_data(m_cfbf, entry, &data_len, errorText);

        if (!data)
        {
            fprintf(out, "\n");
            m2eReturn();
            return 1;
        }

        int srclen = 0;
        unsigned char* dst = 0;
        unsigned char** dest = &dst;
        char errorText[256]; errorText[0] = 0;
        int errorTextLen = 256;

        const unsigned char* src = (const unsigned char*)(data);
        int retlen = RtfDecompressor::DecompressRtf(out, (unsigned char*)src, data_len, dest, errorText, errorTextLen);
        if (retlen > 0)
        {
            dst[retlen] = 0;

            std::string bin((const char*)dst, retlen);

            //fprintf(stdout, "%s\n", bin.c_str());

            char* cpos = strstr((char*)dst, "fromhtml1");  // large effort is required to support fromtext

            std::string result_utf8;

            if (cpos)
            {
                if (hdr_utf8[hdr_utf8.length() - 1] != '\n')
                    hdr_utf8.append("\r\n");

                RTF2HTMLConverter rtf2html;

                std::string result_utf8 = rtf2html.rtf2html((char*)dst);  // simplified extraction of Html, need to do berer later
                //fprintf(stdout, "\n\nRTF2HTMLConverter: rtf2html: !!!!!!!!!!!!!!!\n%s\n\n", result.c_str());

                // Optymistic guessing, may need to enhance once I learn more
                UINT pageCode = 65001;  // utf-8
                if (rtf2html.m_ansicpg != 0)
                {
                    pageCode = rtf2html.m_ansicpg;
                }
                else if (nInternetCodepage != 0)
                    pageCode = nInternetCodepage;
                else
                    ;   // or must quess somehow

                std::string charset;
                BOOL ret = TextUtilsEx::id2charset(pageCode, charset);

                hdr_utf8.append("\r\n\r\n--");
                hdr_utf8.append(alternativeBoundary);
                hdr_utf8.append("\r\n"); // to make sure

                hdr_utf8.append("Content-Type: text/html; charset=\"");
                hdr_utf8.append(charset);
                hdr_utf8.append("\";");
                hdr_utf8.append("\r\nContent-Transfer-Encoding: quoted-printable");;
                hdr_utf8.append("\r\n\r\n"); // to make sure

				//hdr_utf8.append(result_utf8);

                CMimeCodeQP eQP;
                bool bEncoding = true;
                eQP.SetInput(result_utf8.c_str(), result_utf8.size(), bEncoding);
                int eLength = eQP.GetOutputLength();

                char* eBuffer = (char*)cfbf_malloc(eLength);
                if (eBuffer == NULL)
                    return 1;
                int retlen = eQP.GetOutput((unsigned char*)eBuffer, eLength);
                if (retlen > 0)
                    int deb = 1;
                else
                    int deb = 1;

                hdr_utf8.append(eBuffer, retlen);

                free(eBuffer);

                bodyFound = true;
                //fprintf(stdout, "\n\n%s\n", "Found RtfCompressed HtmlBody");
			}
            else
            {
                //fprintf(stdout, "\n\n%s\n", "Found but ignorted RtfCompressed TextBody");
                ; // result_utf8.append((const char*)dst, retlen); // and
                ; // covert text RTF to HTML
            }

            free(dst);
        }
        free(data);
    }

    hdr_utf8.append("\r\n\r\n--");
    hdr_utf8.append(alternativeBoundary);
    hdr_utf8.append("--");
    hdr_utf8.append("\r\n"); // to make sure

    if (hasValidAttachments)
        int deb = 1;

    if (hasValidAttachments && m_attachList.size())
    {
        std::list<AttachmentInfo>::iterator it;
        for (it = m_attachList.begin(); it != m_attachList.end(); ++it)
        {

            std::string str_PidTagAttachLongFilename;
            UINT16 type_PidTagAttachLongFilename;
            bool found_PidTagAttachLongFilename = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, it->m_attach.m_PidTagAttachLongFilename, str_PidTagAttachLongFilename, type_PidTagAttachLongFilename, errorText);

            std::string str_PidTagDisplayName;
            UINT16 type_PidTagDisplayName;
            bool found_PidTagDisplayName = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, it->m_attach.m_PidTagDisplayName, str_PidTagDisplayName, type_PidTagDisplayName, errorText);

            std::string str_PidTagAttachExtension;
            UINT16 type_PidTagAttachExtension;
            bool found_PidTagAttachExtension = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, it->m_attach.m_PidTagAttachExtension, str_PidTagAttachExtension, type_PidTagAttachExtension, errorText);

            std::string str_PidTagAttachMimeTag;
            UINT16 type_PidTagAttachMimeTag;
            bool found_PidTagAttachMimeTag = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, it->m_attach.m_PidTagAttachMimeTag, str_PidTagAttachMimeTag, type_PidTagAttachMimeTag, errorText);

            std::string str_PidTagAttachFilename;
            UINT16 type_PidTagAttachFilename;
            bool found_PidTagAttachFilename = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, it->m_attach.m_PidTagAttachFilename, str_PidTagAttachFilename, type_PidTagAttachFilename, errorText);

            std::string str_PidTagAttachContentId;
            UINT16 type_PidTagAttachContentId;
            bool found_PidTagAttachContentId = OutlookMessage::GetStreamDirEntryValueString(this->m_cfbf, it->m_attach.m_PidTagAttachContentId, str_PidTagAttachContentId, type_PidTagAttachContentId, errorText);

            UINT16 propertyId = PidTagAttachFlags;
            bool isRootParent = false;
            UINT16 type_PidTagAttachFlags = 0;
            std::string str_PidTagAttachFlags;
            UINT64 nPidTagAttachFlags;
            bool found_PidTagAttachFlags = OutlookMessage::GetFixedLengthPropertValue(this->m_cfbf, isRootParent, m_body.m_properties_version1_0, propertyId,
                nPidTagAttachFlags, type_PidTagAttachFlags, str_PidTagAttachFlags, false);

            propertyId = PidTagAttachMethod;
            isRootParent = false;
            UINT16 type_PidTagAttachMethod = 0;
            std::string str_PidTagAttachMethod;
            UINT64 nPidTagAttachMethod;
            bool found_PidTagAttachMethod = OutlookMessage::GetFixedLengthPropertValue(this->m_cfbf, isRootParent, m_body.m_properties_version1_0, propertyId,
                nPidTagAttachMethod, type_PidTagAttachMethod, str_PidTagAttachMethod, false);

            //  nPidTagAttachFlags == 4 indicates inline in PidTagBodyHtml

            //m_PidTagAttachEncoding
            isRootParent = false;
            DWORD propertyId_PidTagAttachEncoding = PidTagAttachEncoding;
            UINT64 value_PidTagAttachEncoding;
            UINT16 type_PidTagAttachEncoding;
            bool found_PidTagAttachEncoding = OutlookMessage::GetDirPropertyValueFixedLength(this->m_cfbf, isRootParent, it->m_attach.m_properties_version1_0, propertyId_PidTagAttachEncoding,
                value_PidTagAttachEncoding, type_PidTagAttachEncoding);

            UINT16 type = 0;
            int data_len;

            if (it->m_attach.m_PidTagAttachDataBinary)
            {
                char* data = GetStreamDirEntryValue(m_cfbf, it->m_attach.m_PidTagAttachDataBinary, type, data_len, errorText);

                if (data)
                {
                    _ASSERTE(type == PtypBinary);

                    CMimeCodeBase64 e64;
                    bool bEncoding = true;
                    e64.SetInput(data, data_len, bEncoding);
                    int eLength = e64.GetOutputLength();

                    char* eBuffer = (char*)cfbf_malloc(eLength);
                    if (eBuffer == NULL)
                        return 1;
                    int retlen = e64.GetOutput((unsigned char*)eBuffer, eLength);
                    if (retlen > 0)
                        int deb = 1;
                    else
                        int deb = 1;

                    if (str_PidTagAttachMimeTag.empty())
                    {
                        std::string mimeType;
                        if (!str_PidTagAttachExtension.empty()) {
                            mimeType = mimetypes::from_extension(str_PidTagAttachExtension);
                        }
                        else if (!str_PidTagDisplayName.empty()) {
                            mimeType = PathFindExtensionA(str_PidTagDisplayName.c_str());
                        }
                        str_PidTagAttachMimeTag = mimeType;
                    }

                    hdr_utf8.append("\r\n--");
                    hdr_utf8.append(mixedBoundary);

                    hdr_utf8.append("\r\nContent-Type: ");
                    hdr_utf8.append(str_PidTagAttachMimeTag);
                    hdr_utf8.append("; name=\"");
                    if (!str_PidTagAttachLongFilename.empty())
                        hdr_utf8.append(str_PidTagAttachLongFilename);
                    else if (!str_PidTagAttachFilename.empty())
                        hdr_utf8.append(str_PidTagAttachFilename);
                    else if (!str_PidTagDisplayName.empty())
                        hdr_utf8.append(str_PidTagDisplayName);


                    hdr_utf8.append("\"\r\nContent-Disposition: ");
                    if (found_PidTagAttachContentId)
                        hdr_utf8.append("inline");
                    else
                        hdr_utf8.append("attachment");
                    hdr_utf8.append("; filename=\"");
                    if (!str_PidTagAttachLongFilename.empty())
                        hdr_utf8.append(str_PidTagAttachLongFilename);
                    else if (!str_PidTagAttachFilename.empty())
                        hdr_utf8.append(str_PidTagAttachFilename);
                    else if (!str_PidTagDisplayName.empty())
                        hdr_utf8.append(str_PidTagDisplayName);

                    hdr_utf8.append("\"\r\nContent-Transfer-Encoding: base64");
                    if (found_PidTagAttachContentId)
                    {
                        hdr_utf8.append("\r\nContent-ID: <");
                        hdr_utf8.append(str_PidTagAttachContentId);
                        hdr_utf8.append(">\r\n");
                    }

                    hdr_utf8.append("\r\n\r\n");

                    hdr_utf8.append(eBuffer, retlen);

                    int deb1 = 1;

                    // PidTagAttachMethod
                    // PidTagDisplayName
                    // PidTagAttachTag   /// may not be set
                    // TNEF {0x2A,86,48,86,F7,14,03,0A,01} 
                    // afStorage {0x2A,86,48,86,F7,14,03,0A,03,02,01}    // application specific
                    // MIME {0x2A,86,48,86,F7,14,03,0A,04} 
                    // PidTagRenderingPosition The value 0xFFFFFFFF indicates a hidden attachment that is not to be rendered in the main text.
                    // PidTagAttachFlags 1,2, 0x00000004 -> inline in PidTagBodyHtml
                    // PidTagAttachEncoding  If the attachment is in MacBinary format, this property is set to  "{0x2A,86,48,86,F7,14,03,0B,01}"; otherwise, it is unset.
                    // PidTagTextAttachmentCharset -> Content-Type
                    // PidTagAttachMimeTag -> Content-Type
                    // PidTagAttachContentId A content identifier unique to this Message  object that matches a corresponding "cid:"URI scheme reference in the HTML body of the Message object.
                    // PidTagAttachContentLocation A relative or full URI that matches a corresponding reference in the HTML
                    // PidNameAttachmentMacContentType The Content-Type header of the Macintosh  attachment.
                    // PidNameAttachmentMacInfo
                    // PidTagBodyContentLocation

                    free(data);
                    free(eBuffer);
                }
                else
                {
                    ;
                }
            }
            else if (it->m_attach.m_PidTagAttachDataObject)
            {
                if (it->m_attach.m_OutlookMessage)
                {
                    OutlookMessage* nested_msg = it->m_attach.m_OutlookMessage;
                    std::string MessageClass;
                    bool found_MessageClass = nested_msg->GetMessageClass(MessageClass);
                    if (found_MessageClass &&
                        ((MessageClass.compare("IPM.Note") == 0) ||
                            (MessageClass.compare("IPM.Note.SMIME.MultipartSigned") == 0)))
                    {
                        std::string errorText;
                        std::string nested_msg_utf8;

                        int retMsg2Eml_2 = nested_msg->Msg2Eml(nested_msg_utf8, errorText);

                        const char* data = nested_msg_utf8.c_str();
                        int data_len = nested_msg_utf8.length();

                        {
                            CMimeCodeBase64 e64;
                            bool bEncoding = true;
                            e64.SetInput(data, data_len, bEncoding);
                            int eLength = e64.GetOutputLength();

                            char* eBuffer = (char*)cfbf_malloc(eLength);
                            if (eBuffer == NULL)
                                return 1;
                            int retlen = e64.GetOutput((unsigned char*)eBuffer, eLength);
                            if (retlen > 0)
                                int deb = 1;
                            else
                                int deb = 1;

                            if (str_PidTagAttachMimeTag.empty())
                            {
#if 0
                                std::string mimeType;
                                if (!str_PidTagAttachExtension.empty()) {
                                    mimeType = mimetypes::from_extension(str_PidTagAttachExtension);
                                }
                                else if (!str_PidTagDisplayName.empty()) {
                                    mimeType = PathFindExtensionA(str_PidTagDisplayName.c_str());
                                }
                                str_PidTagAttachMimeTag = mimeType;
#else
                                str_PidTagAttachMimeTag.append("message/rfc822");
#endif
                            }

                            hdr_utf8.append("\r\n--");
                            hdr_utf8.append(mixedBoundary);

                            hdr_utf8.append("\r\nContent-Type: ");
                            hdr_utf8.append(str_PidTagAttachMimeTag);
                            hdr_utf8.append("; name=\"");
                            hdr_utf8.append(str_PidTagDisplayName);

                            hdr_utf8.append("\"\r\nContent-Disposition: ");
                            if (found_PidTagAttachContentId)
                                hdr_utf8.append("inline");
                            else
                                hdr_utf8.append("attachment");
                            hdr_utf8.append("; filename=\"");
                            hdr_utf8.append(str_PidTagDisplayName);

                            hdr_utf8.append("\"\r\nContent-Transfer-Encoding: base64");
                            if (found_PidTagAttachContentId)
                            {
                                hdr_utf8.append("\r\nContent-ID: <");
                                hdr_utf8.append(str_PidTagAttachContentId);
                                hdr_utf8.append(">\r\n");
                            }

                            hdr_utf8.append("\r\n\r\n");

                            hdr_utf8.append(eBuffer, retlen);

                            int deb1 = 1;

                            free(eBuffer);
                        }
                        int deb = 1;
                    }
                    delete nested_msg;
                }
            }
            int deb = 1;
        }

        if (hasValidAttachments)
        {
            //hdr_utf8.append("\r\n--");
            //hdr_utf8.append(mixedBoundary);

            hdr_utf8.append("\r\n\r\n--");
            hdr_utf8.append(mixedBoundary);
            hdr_utf8.append("--");
            hdr_utf8.append("\r\n\r\n"); // to make sure
        }
    }

    TextUtilsEx::delete_charset2Id();
    TextUtilsEx::delete_id2charset();

    return 1;
}

int  OutlookMessage::WriteToEmlFile(const char* msgFilePath, std::string &emlData_utf8)
{
    size_t outlen = 0;
    wchar_t* filePathW = UTF8ToUTF16((char*)msgFilePath, strlen(msgFilePath), &outlen);

    CStringW fileName;
    FileUtils::CPathStripPath(filePathW, fileName);

    CStringW cStrNamePath = LR"(C:\Users\tata\Downloads\msg2eml\)";
    cStrNamePath.Append((LPCWSTR)fileName);
    cStrNamePath.Append(L".eml");
    free(filePathW);

    unsigned char* msg_data = (unsigned char*)emlData_utf8.c_str();
    int wret = OutlookMessage::Write2File((LPWSTR)((LPCWSTR)cStrNamePath), msg_data, emlData_utf8.length());

    return 1;
}

std::string  OutlookMessage::DeleteFieldFromHeader(std::string& str, std::string& field)
{
    const char* p = str.c_str();
    std::string fld = "\n" + field + ":";
    const char* f = fld.c_str();
    int strLength = str.length();
    int fldLength = fld.length();

    // Expensive, fix later

    // Find first occurence of field at the line beginning
    // May need to check if more occurences??
    const char* beg = 0;
    int pos = 0;
    int pos_begin = 0;
    do
    {
        int result = strnicmp(p, f, fldLength);
        if (result == 0)
        {
            pos_begin = pos;
            break;
        }
        else
        {
            pos++;
            p++;
        }
    } while (pos < strLength);

    if (pos >= strLength)
    {
        std::string empty;
        return empty; // not found
    }

    // find first line that that doesn't start with white characters ' ' and '\t' or end of str

    const char* cpos = str.c_str() + pos + 1;  // skip '\n'
    const char* cpos_last = str.c_str() + strLength;
    while (cpos < cpos_last)
    {
        cpos = strchr(cpos, '\n');
        if (cpos == 0)
        {
            cpos = cpos_last;
            break;
        }
        char c = *cpos;
        if ((c != ' ') && (c != '\t'))
            break;
        else
            cpos++;
    }

    int lengthToErase = cpos - str.c_str() - pos_begin;
    _ASSERTE(cpos >= 0);
    std::string erasedField(&str[pos_begin], lengthToErase);
    std::string& out_str = str.erase(pos_begin, lengthToErase);

    return erasedField;
}

std::string OutlookMessage::GenerateRandomString(int length)
{
	const std::string characters = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
	std::random_device random_device;
	std::mt19937 generator(random_device());

	std::string random_string(characters);
	std::shuffle(random_string.begin(), random_string.end(), generator);

	return random_string.substr(0, length);
}

bool OutlookMessage::Write2File(wchar_t* cStrNamePath, const unsigned char* data, int dataLength)
{
    DWORD dwAccess = GENERIC_WRITE;
    DWORD dwCreationDisposition = CREATE_ALWAYS;

    HANDLE hFile = CreateFileW(cStrNamePath, dwAccess, FILE_SHARE_READ, NULL, dwCreationDisposition, 0, NULL);
    if (hFile == INVALID_HANDLE_VALUE)
    {
        CString errText = FileUtils::GetLastErrorAsString();
        DWORD err = GetLastError();
        TRACE(L"(FileSize)INVALID_HANDLE_VALUE error=%ld file=%s\n%s\n", err, cStrNamePath, errText);
        return false;
    }
    else
    {
        const unsigned char* pszData = data;
        int nLeft = dataLength;

        for (;;)
        {
            DWORD nWritten = 0;

            BOOL retWrite = WriteFile(hFile, pszData, nLeft, &nWritten, 0);
            if (nWritten != nLeft)
                int deb = 1;

            if (nWritten < 0)
                break;

            pszData += nWritten;
            nLeft -= nWritten;
            if (nLeft <= 0)
                break;
        }

        BOOL retClose = CloseHandle(hFile);
        if (retClose == FALSE)
            int deb = 1;
    }
    return true;
}

// Property Data can be set to PtyBinary or PtyString but in reality is PtyString8
// Optymistic check
bool OutlookMessage::IsString8(char* buf, int buflen)
{
    int cnt = 0;
    int len = (buflen > 10000) ? 10000 : buflen;
    for (int i = 0; i < len; i++)
    {
        if (buf[i] == 0)
            cnt++;

        if (cnt > 2)
            return false;
    }
    return true;
}

void OutlookMessage::m2eReturn()
{
    int deb = 1;
}


int OutlookMessage::Write2File(HANDLE hFile, LPCVOID lpBuffer, DWORD nNumberOfBytesToWrite, LPDWORD lpNumberOfBytesWritten)
{
    DWORD bytesToWrite = nNumberOfBytesToWrite;
    const wchar_t* pszDataW = (wchar_t*)lpBuffer;
    const unsigned char* pszData = (unsigned char*)lpBuffer;
    DWORD nwritten = 0;
    while (bytesToWrite > 0)
    {
        nwritten = 0;
        int retW = WriteFile(hFile, pszData, bytesToWrite, &nwritten, NULL);
        if (retW == 0) {
            DWORD retval = GetLastError();
            break;
        }
        pszData += nwritten;
        bytesToWrite -= nwritten;
    }
    *lpNumberOfBytesWritten = nNumberOfBytesToWrite - bytesToWrite;
    if (*lpNumberOfBytesWritten != nNumberOfBytesToWrite)
        return FALSE;
    else
        return TRUE;
}


HANDLE OutlookMessage::FileOpen(wchar_t* cStrNamePath, CString& errorText, BOOL truncate)
{
    DWORD dwAccess = GENERIC_WRITE;
    DWORD dwCreationDisposition = CREATE_ALWAYS;


    HANDLE hFile = CreateFile(cStrNamePath, dwAccess, FILE_SHARE_READ, NULL, dwCreationDisposition, 0, NULL);
    if (hFile == INVALID_HANDLE_VALUE)
    {
        errorText = FileUtils::GetLastErrorAsString();
        DWORD err = GetLastError();
        TRACE(L"(OpenFile)INVALID_HANDLE_VALUE error=%ld file=%s\n%s\n", err, cStrNamePath, errorText);
        return hFile;
    }

    if (truncate)  // dwCreationDisposition = CREATE_ALWAYS;  truncate file suppose truncate but it doesnt seem work ??
    {
        LARGE_INTEGER li = { 0, 0 };
        BOOL ret = SetFilePointerEx(hFile, li, 0, FILE_BEGIN);
        int deb = 1;
    };

    return hFile;
}

void OutlookMessage::FileClose(HANDLE hFile)
{
    if (hFile != INVALID_HANDLE_VALUE)
        CloseHandle(hFile);
}




