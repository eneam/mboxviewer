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

#pragma once

#include <stdio.h>
#include <string.h>
#include <string>
#include <list>
#include <random>
//#include <atlstr.h>

#include "cfbf.h"
#include "property.h"
#include "RtfDecompressor.h"
#include "stdafx.h"

struct DirEntry;
struct StructuredStorageHeader;
class OutlookMessage;

bool IsLittleEndianType();

struct Body
{
    DirEntry* m_PidTagMessageFlags;
    DirEntry* m_PidTagMessageClass;
    DirEntry* m_PidTagBodyHtml;
    DirEntry* m_PidTagRtfCompressed;
    DirEntry* m_PidTagNativeBody;
    //DirEntry* m_PidTagNativeBody;
    DirEntry* m_PidTagBody;
    //DirEntry* m_PidTagRtfCompressed;
    DirEntry* m_PidTagInReplyToId;
    DirEntry* m_PidTagDisplayName;
    DirEntry* m_PidTagRecipientDisplayName;
    DirEntry* m_PidTagAddressType;
    DirEntry* m_PidTagEmailAddress;
    DirEntry* m_PidTagSmtpAddress;
    DirEntry* m_PidTagDisplayBcc;
    DirEntry* m_PidTagDisplayCc;
    DirEntry* m_PidTagDisplayTo;
    DirEntry* m_PidTagOriginalDisplayBcc;
    DirEntry* m_PidTagOriginalDisplayCc;
    DirEntry* m_PidTagOriginalDisplayTo;
    DirEntry* m_PidTagRecipientType;
    DirEntry* m_PidTagSenderName;
    DirEntry* m_PidTagSentRepresentingName;
    DirEntry* m_PidTagSenderEmailAddress;
    DirEntry* m_PidTagSenderSmtpAddress;
    DirEntry* m_PidTagSentRepresentingSmtpAddress;
    DirEntry* m_PidTagMessageSize;
    DirEntry* m_PidTagMessageRecipients;
    DirEntry* m_PidTagMessageAttachments;
    DirEntry* m_PidTagHasAttachments;
    DirEntry* m_PidTagSubject;
    DirEntry* m_PidTagSubjectPrefix;
    DirEntry* m_PidTagOriginalSubject;
    DirEntry* m_PidTagNormalizedSubject;
    DirEntry* m_PidTagTransportMessageHeaders;
    DirEntry* m_PidTagLanguage;
    DirEntry* m_PidTagStoreSupportMask;
    DirEntry* m_PidTagMessageLocaleId;
    DirEntry* m_PidTagMessageCodepage;
    DirEntry* m_PidTagInternetCodepage;
    DirEntry* m_PidTagInternetMessageId;
    DirEntry* m_PidTagBodyContentLocation;
    DirEntry* m_PidTagBodyContentId;
    DirEntry* m_PidTagLastModifierName;

    DirEntry* m_properties_version1_0;
    int SetProperty(int propertIdNumb, int propertTypeNumb, DirEntry* entry);
    void Print(struct cfbf* cfbf, int level);
};

struct Attachment
{
    OutlookMessage* m_OutlookMessage;
    DirEntry* m_PidTagAttachDataObject;
    DirEntry* m_PidTagAttachMethod;
    DirEntry* m_PidTagAttachFilename;
    DirEntry* m_PidTagAttachMimeTag;
    DirEntry* m_PidTagAttachEncoding;
    DirEntry* m_PidTagAttachExtension;
    DirEntry* m_PidTagAttachLongFilename;
    DirEntry* m_PidTagAttachDataBinary;
    DirEntry* m_PidTagAttachSize;
    DirEntry* m_PidTagLanguage;
    DirEntry* m_PidTagDisplayName;
    DirEntry* m_PidTagAttachContentId;
    //
    DirEntry* m_properties_version1_0;
    int SetProperty(int propertIdNumb, int propertTypeNumb, DirEntry* entry);
    void Print(struct cfbf* cfbf, int level);
};

struct AttachmentInfo_
{
    std::string m_name;
    Attachment m_attach;
};

struct Recipient
{
    DirEntry* m_PidTagRecipientType;  // It may just be available under DirEntry* m_properties_version1_0;
    DirEntry* m_PidTagEmailAddress;
    DirEntry* m_PidTagDisplayName;
    DirEntry* m_PidTagAddressType;
    DirEntry* m_PidTagSmtpAddress;
    DirEntry* m_PidTagRecipientDisplayName;
    DirEntry* m_PidTagLanguage;
    DirEntry* m_PidTagRecipientFlags;
    //
    DirEntry* m_properties_version1_0;
    int SetProperty(int propertIdNumb, int propertTypeNumb, DirEntry* entry);
    int GetRecipientType(struct cfbf* cfbf, DirEntry* m_properties_version1_0);  // an item of array of properties
    void Print(struct cfbf* cfbf, int level);
};

struct RecipientInfo
{
    std::string m_name;
    Recipient m_recip;
};

struct PropertyHeaderTop
{
    UINT64 Reserved_1;
    UINT32 NextRecipientID;
    UINT32 NextAttachmentID;
    UINT32 RecipientCount;
    UINT32 AttachmentCount;
    UINT64 Reserved_2;
};

struct PropertyHeaderOfAttachmentAndRecipient
{
    UINT64 Reserved_1;
};

struct PropertyHeaderEmbeded
{
    UINT64 Reserved_1;
    UINT32 NextRecipientID;
    UINT32 NextAttachmentID;
    UINT32 RecipientCount;
    UINT32 AttachmentCount;
};

struct FixedLengthPropertyValue
{
    UINT32 Data;
    UINT32 Reserved;
};

struct PropertyEntry
{
    UINT32 PropertyTag;
    UINT32 Flags;
    UINT64 Value;
};


struct VariableLengthOrMultipleValuedPropertyEntry
{
    UINT32 PropertyTag;
    UINT32 Flags;
    UINT32 Size;
    UINT32 Reserved;
};

//int ParseOutlookMsg(void* cookie, struct cfbf* cfbf, DirEntry* e, DirEntry* parent, unsigned long entry_id, int depth);

typedef int (*Parse_Outlook_Msg)(void* cookie, struct cfbf* cfbf, DirEntry* e, DirEntry* parent, unsigned long entry_id, int depth);
struct OutlookMsgHelper;

class OutlookMessage
{
public:
    OutlookMessage()
    {
        m_baseDepthLevel = 9;
        logFile = 0;
        emlFile = 0;
        m_cfbf = 0;
        memset(&m_body, 0, sizeof(Body));
    }
    OutlookMessage(FILE* outLogFile, HANDLE msg2emlFile, struct cfbf* icfbf)
    {
        _ASSERTE(icfbf);
        m_baseDepthLevel = 0;
        logFile = outLogFile;
        emlFile = msg2emlFile;
        m_cfbf = icfbf;
        memset(&m_body, 0, sizeof(Body));
    }
    ~OutlookMessage();

    int ParseMsg(struct cfbf* cfbf, Parse_Outlook_Msg _ParseOutlookMsg, OutlookMsgHelper& msgHelper, std::string& errorText);
    int Msg2Eml(std::string& errorText);
    int Msg2Eml(std::string& msg_utf8, std::string& errorText);
    int WordEncode(std::string& fld, std::string& wordEncodedFld, int encodeType = 'Q');

    int m_baseDepthLevel;
    struct cfbf* m_cfbf;
    HANDLE emlFile;
    FILE* logFile;
    Body m_body;
    std::list<RecipientInfo> m_recipList;
    std::list<AttachmentInfo_> m_attachList;

    int SetProperty(DirEntry* parent, DirEntry* entry);
    RecipientInfo* FindRecip(std::string& name);
    AttachmentInfo_* FindAttach(std::string& name);

    void GetToLists(std::string& To, std::string& CC, std::string& BCC);
    bool GetMessageClass(std::string& MessageClass);
    bool GetFromAddress(std::string& From);
    bool GetSubject(std::string& Subject);
    bool GetDate(std::string& Date);
    bool GetMessageId(std::string& MessageId);
    bool GetInReplyTo(std::string& InReplyTo);
    bool GetContentLocation(std::string& ContentLocation);
    bool GetContentId(std::string& ContenId);
    bool GetMessageCodepage(UINT64& nMessageCodepage, std::string& MessageCodepage);
    bool GetInternetCodepage(UINT64& nInternetCodepage, std::string& InternetCodepage);
    bool GetStoreSupportMask(UINT64& nStoreSupportMask, std::string& StoreSupportMask);
    bool GetNativeBody(UINT64& nNativeBody, std::string& NativeBody);

    bool GetFixedLengthPropertValue(struct cfbf* cfbf, bool isRootParent, DirEntry* entry, UINT16 propertyId, UINT64& nValue, UINT16& type, std::string& sValue, bool hexFormat = false);
    bool GetPropertyString(struct cfbf* cfbf, DirEntry* entry, std::string& value_utf8, UINT16& type, std::string& errorText);

    static bool Time2Date(UINT64 value, std::string& Date);

    static bool GetDirPropertyValueFixedLength(struct cfbf* cfbf, bool isRootParent, DirEntry* entry, UINT16 propertyId, UINT64& value, UINT16& type);
    static bool GetStreamDirEntryValueString(struct cfbf* cfbf, DirEntry* entry, std::string& value, UINT16& type, std::string& errorText);
    static char* GetStreamDirEntryValue(struct cfbf* cfbf, DirEntry* entry, UINT16& type, int& data_len, std::string& errorText);

    std::string  DeleteFieldFromHeader(std::string &data_utf8, std::string& field);

    static std::string GenerateRandomString(int length);
    static bool IsString8(char* buf, int buflen);

    void Print();
    static void PrintDirProperties(struct cfbf* cfbf, int level, DirEntry* entry, int depth = 0);

    static bool Write2File(wchar_t* cStrNamePath, const unsigned char* data, int dataLength);
    static int WriteToEmlFile(const char* msgFilePath, std::string& emlData_utf8);
    static void m2eReturn();

    static int Write2File(HANDLE hFile, LPCVOID lpBuffer, DWORD nNumberOfBytesToWrite, LPDWORD lpNumberOfBytesWritten);
    static HANDLE FileOpen(wchar_t* cStrNamePath, CString& errorText, BOOL truncate);
    static void FileClose(HANDLE hFile);
};

struct OutlookMsgHelper
{
    OutlookMsgHelper(struct cfbf* _cfbf, HANDLE emlFile, FILE* _out = stdout)
        :msg(_out, emlFile, _cfbf)
    {
        cfbf = _cfbf;
        out = _out;
        emlFileHandle = emlFile;
        active_msg = &msg;
    }
    ~OutlookMsgHelper()
    {
        _ASSERTE(m_msgList.size() == 0);
        int deb = 1;
    }

    std::list<OutlookMessage*> m_msgList;  // MSG attachments
    std::string msg_utf8;
    OutlookMessage msg;
    OutlookMessage* active_msg;
    struct cfbf* cfbf;
    HANDLE emlFileHandle;
    FILE* out;
};

int IsValidOutlookMsgFile(CString& fname);

//int ParseOutlookMsg(void* cookie, struct cfbf* cfbf, DirEntry* e,
    //DirEntry* parent, unsigned long entry_id, int depth);

