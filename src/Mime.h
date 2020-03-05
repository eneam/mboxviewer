//////////////////////////////////////////////////////////////////////
//
// MIME message encoding/decoding
//
// Jeff Lee
// Dec 11, 2000
//
//////////////////////////////////////////////////////////////////////

#if !defined(_MIME_H)
#define _MIME_H

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <string>
#include <list>
using namespace std;

// RFC 1521 - Mechanisms for Specifying and Describing the Format of Internet Message Bodies
// RFC 2045 - Format of Internet Message Bodies
// RFC 2046 - Media Types
// RFC 2047 - Message Header Extensions for Non-ASCII Text
// RFC 2049 - Conformance Criteria and Examples
// RFC 822 - Standard For The Format of ARPA Internet Text Message

class CMimeConst
{
public:
	// field names
	static inline const char* MimeVersion() { return "MIME-Version"; }
	static inline const char* ContentType() { return "Content-Type"; }
	static inline const char* TransferEncoding() { return "Content-Transfer-Encoding"; }
	static inline const char* ContentID() { return "Content-ID"; }
	static inline const char* ContentDescription() { return "Content-Description"; }
	static inline const char* ContentDisposition() { return "Content-Disposition"; }

	// parameter names
	static inline const char* Charset() { return "charset"; }
	static inline const char* Name() { return "name"; }
	static inline const char* Filename() { return "filename"; }
	static inline const char* Boundary() { return "boundary"; }
	static inline const char* Attachment() { return "attachment"; }

	// parameter values
	static inline const char* Encoding7Bit() { return "7bit"; }
	static inline const char* Encoding8Bit() { return "8bit"; }
	static inline const char* EncodingBinary() { return "binary"; }
	static inline const char* EncodingQP() { return "quoted-printable"; }
	static inline const char* EncodingBase64() { return "base64"; }

	static inline const char* MediaText() { return "text"; }
	static inline const char* MediaImage() { return "image"; }
	static inline const char* MediaAudio() { return "audio"; }
	static inline const char* MediaVedio() { return "vedio"; }
	static inline const char* MediaApplication() { return "application"; }
	static inline const char* MediaMultiPart() { return "multipart"; }
	static inline const char* MediaMessage() { return "message"; }
};

//////////////////////////////////////////////////////////////////////
// CMimeField class - Represents a field of a MIME body part header
//////////////////////////////////////////////////////////////////////
class CMimeField
{
public:
	CMimeField() {}
	~CMimeField() {}

public:
	void SetName(const char* pszName);
	const char* GetName() const;
	void SetValue(const char* pszValue);
	const char* GetValue() const;
	void GetValue(string& strValue) const;
	void SetParameter(const char* pszAttr, const char* pszValue);
	bool GetParameter(const char* pszAttr, string& strValue) const;
	void SetCharset(const char* pszCharset);
	const char* GetCharset() const;

	void Clear();
	int GetLength() const;
	int Store(char* pszData, int nMaxSize) const;
	int Load(const char* pszData, int nDataSize);

private:
	string m_strName;				// field name
	string m_strValue;				// field value
	string m_strCharset;			// charset for non-ascii text

private:
	bool FindParameter(const char* pszAttr, int& nPos, int& nSize) const;
};

inline void CMimeField::SetName(const char* pszName)
{ m_strName = pszName; }

inline const char* CMimeField::GetName() const
{ return m_strName.data(); }

inline void CMimeField::SetValue(const char* pszValue)
{ m_strValue = pszValue; }

inline const char* CMimeField::GetValue() const
{ return m_strValue.data(); }

inline void CMimeField::SetCharset(const char* pszCharset)
{ m_strCharset = pszCharset; }

inline const char* CMimeField::GetCharset() const
{ return m_strCharset.c_str(); }

inline void CMimeField::Clear()
{ m_strName.erase(); m_strValue.erase(); m_strCharset.erase(); }

//////////////////////////////////////////////////////////////////////
// CMimeHeader class - Represents the header of a MIME body part
//////////////////////////////////////////////////////////////////////
class CMimeHeader
{
public:
	CMimeHeader() {}
	virtual ~CMimeHeader() { Clear(); }

public:
	enum MediaType
	{
		MEDIA_TEXT, MEDIA_IMAGE, MEDIA_AUDIO, MEDIA_VEDIO, MEDIA_APPLICATION,
		MEDIA_MULTIPART, MEDIA_MESSAGE,
		MEDIA_UNKNOWN
	};
	MediaType GetMediaType() const;

	// set/get the values of header fields
	void SetField(const CMimeField& field);
	const CMimeField* GetField(const char* pszFieldName) const;
	CMimeField* GetField(const char* pszFieldName);
	void SetFieldValue(const char* pszFieldName, const char* pszFieldValue, const char* pszCharset=NULL);
	const char* GetFieldValue(const char* pszFieldName) const;
	bool SetParameter(const char* pszFieldName, const char* pszAttr, const char* pszValue);
	string GetParameter(const char* pszFieldName, const char* pszAttr) const;
	void SetFieldCharset(const char* pszFieldName, const char* pszCharset);
	const char* GetFieldCharset(const char* pszFieldName) const;

	// helper functions for standard body part fields
	const char* GetContentId() const;	
	void SetContentType(const char* pszValue, const char* pszCharset=NULL);
	const char* GetContentType() const;			// Content-Type: mediatype/subtype
	string GetMainType() const;
	string GetSubType() const;
	void SetCharset(const char* pszCharset);	// Content-Type: text/...; charset=...
	string GetCharset() const;
	void SetName(const char* pszName);			// Content-Type: image/...; name=...
	string GetName() const;
	void SetBoundary(const char* pszBoundary=NULL);	// Content-Type: multipart/...; boundary=...
	string GetBoundary() const;

	void SetTransferEncoding(const char* pszValue);
	const char* GetTransferEncoding() const;	// Content-Transfer-Encoding: ...
	void SetDisposition(const char* pszValue, const char* pszCharset=NULL);
	const char* GetDisposition() const;			// Content-Disposition: ...
	string GetFilename() const;					// Content-Disposition: ...; filename=...
	void SetDescription(const char* pszValue, const char* pszCharset=NULL);
	const char* GetDescription() const;			// Content-Description: ...


	string GetNameCharset() const;
	string GetFilenameCharset() const;

	typedef list<CMimeField> CFieldList;
	CFieldList& Fields() { return m_listFields; }

public:
	// overrides
	virtual void Clear();
	virtual int GetLength() const;
	// serialization
	virtual int Store(char* pszData, int nMaxSize) const;
	virtual int Load(const char* pszData, int nDataSize);


protected:
	list<CMimeField> m_listFields;	// list of all header fields
	list<CMimeField>::const_iterator FindField(const char* pszFieldName) const;
	list<CMimeField>::iterator FindField(const char* pszFieldName);


	struct MediaTypeCvt
	{
		int nMediaType;				// media type
		const char* pszSubType;		// subtype
		const char* pszFileExt;		// file extension name
	};
	static const MediaTypeCvt m_TypeCvtTable[];
	static const char* m_TypeTable[];

private:
	CMimeHeader& operator=(const CMimeHeader&);		// forbid operator =
};



// add a new field or update an existing field
inline void CMimeHeader::SetField(const CMimeField& field)
{
	list<CMimeField>::iterator it = FindField(field.GetName());
	if (it != m_listFields.end())
		*it = field;
	else
		m_listFields.push_back(field);
}

// find a field by name
inline const CMimeField* CMimeHeader::GetField(const char* pszFieldName) const
{
	list<CMimeField>::const_iterator it = FindField(pszFieldName);
	if (it != m_listFields.end())
		return &(*it);
	return NULL;
}

inline CMimeField* CMimeHeader::GetField(const char* pszFieldName)
{
	list<CMimeField>::iterator it = FindField(pszFieldName);
	if (it != m_listFields.end())
		return &(*it);
	return NULL;
}

// add a new field or update an existing field
inline void CMimeHeader::SetFieldValue(const char* pszFieldName, const char* pszFieldValue, const char* pszCharset)
{
	CMimeField fd;
	fd.SetName(pszFieldName);
	fd.SetValue(pszFieldValue);
	if (pszCharset != NULL)
		fd.SetCharset(pszCharset);
	SetField(fd);
}

inline const char* CMimeHeader::GetFieldValue(const char* pszFieldName) const
{
	const CMimeField* pfd = GetField(pszFieldName);
	return pfd != NULL ? pfd->GetValue() : NULL;
}

inline void CMimeHeader::SetFieldCharset(const char* pszFieldName, const char* pszCharset)
{
	CMimeField *pfd = GetField(pszFieldName);
	if (pfd)
		pfd->SetCharset(pszCharset);
	else
	{
		CMimeField fd;
		fd.SetName(pszFieldName);
		fd.SetCharset(pszCharset);
		SetField(fd);
	}
}

inline const char* CMimeHeader::GetFieldCharset(const char* pszFieldName) const
{
	const CMimeField* pfd = GetField(pszFieldName);
	return pfd != NULL ? pfd->GetCharset() : NULL;
}

inline bool CMimeHeader::SetParameter(const char* pszFieldName, const char* pszAttr, const char* pszValue)
{
	CMimeField *pfd = GetField(pszFieldName);
	if (pfd)
	{
		pfd->SetParameter(pszAttr, pszValue);
		return true;
	}
	return false;
}

inline string CMimeHeader::GetParameter(const char* pszFieldName, const char* pszAttr) const
{
	string strVal;
	const CMimeField *pfd = GetField(pszFieldName);
	if (pfd)
		pfd->GetParameter(pszAttr, strVal);
	return strVal;
}

inline void CMimeHeader::SetContentType(const char* pszValue, const char* pszCharset)
{ SetFieldValue(CMimeConst::ContentType(), pszValue, pszCharset); }

inline const char* CMimeHeader::GetContentType() const
{ return GetFieldValue(CMimeConst::ContentType()); }

inline const char* CMimeHeader::GetContentId() const
{
	return GetFieldValue(CMimeConst::ContentID());
}

inline string CMimeHeader::GetCharset() const
{ return GetParameter(CMimeConst::ContentType(), CMimeConst::Charset()); }

inline string CMimeHeader::GetName() const
{
	return GetParameter(CMimeConst::ContentType(), CMimeConst::Name());
}
inline string CMimeHeader::GetNameCharset() const
{
	const CMimeField *pfd = GetField(CMimeConst::ContentType());
	return (pfd != 0) ? pfd->GetCharset() : "";
}

inline string CMimeHeader::GetBoundary() const
{ return GetParameter(CMimeConst::ContentType(), CMimeConst::Boundary()); }

inline void CMimeHeader::SetTransferEncoding(const char* pszValue)
{ SetFieldValue(CMimeConst::TransferEncoding(), pszValue); }

inline const char* CMimeHeader::GetTransferEncoding() const
{ return GetFieldValue(CMimeConst::TransferEncoding()); }

// Content-Disposition header field specifies how to present this body part. it could be
// inline or attachment. 'inline' indicates this body part should be displayed in the main body;
// 'attachment' indicates it is separate from the main body. (RFC 2183)

inline void CMimeHeader::SetDisposition(const char* pszValue, const char* pszCharset)
{ SetFieldValue(CMimeConst::ContentDisposition(), pszValue, pszCharset); }

inline const char* CMimeHeader::GetDisposition() const
{ return GetFieldValue(CMimeConst::ContentDisposition()); }

inline string CMimeHeader::GetFilename() const
{ return GetParameter(CMimeConst::ContentDisposition(), CMimeConst::Filename()); }

inline string CMimeHeader::GetFilenameCharset() const
{
	const CMimeField *pfd = GetField(CMimeConst::ContentDisposition());
	return (pfd != 0) ? pfd->GetCharset() : "";
}

inline void CMimeHeader::SetDescription(const char* pszValue, const char* pszCharset)
{ SetFieldValue(CMimeConst::ContentDescription(), pszValue, pszCharset); }

inline const char* CMimeHeader::GetDescription() const
{ return GetFieldValue(CMimeConst::ContentDescription()); }

//////////////////////////////////////////////////////////////////////
// CMimeBody class - Represents a body part in a MIME message
//////////////////////////////////////////////////////////////////////
class CMimeMessage;
class CMimeBody : public CMimeHeader
{
protected:
	CMimeBody() :				// instantiate a CMimeBody object explicitly is not allowed. call CreatePart()
		m_pbText(NULL),
		m_nTextSize(0), m_bIsRelated(false) {}
	virtual ~CMimeBody() { Clear(); }

public:
	int GetContentLength() const;
	const unsigned char* GetContent() const;


	// Parent block is multipart/related block 
	bool IsRelated();

	// operations for 'text' or 'message' media
	bool IsText() const;
	int SetText(const char* pbText, int nLength=0);
	int GetText(char* pbText, int nMaxSize);
	int GetText(string& strText);

	// operations for 'message' media
	bool IsMessage() const;
	bool SetMessage(const CMimeMessage* pMM);
	void GetMessage(CMimeMessage* pMM) const;

	// operations for 'image/audio/vedio/application' (attachment) media
	//bool IsAttachment() const;
	bool ReadFromFile(const char* pszFilename);
	bool WriteToFile(const char* pszFilename);

	// operations for 'multipart' media
	bool IsMultiPart() const;
	void DeleteAll();
	CMimeBody* CreatePart(const char* pszMediaType=NULL, CMimeBody* pWhere=NULL);
	void ErasePart(CMimeBody* pBP);
	CMimeBody* FindFirstPart();
	CMimeBody* FindNextPart();

	typedef list<CMimeBody*> CBodyList;
	int GetBodyPartList(CBodyList& rList) const;
	int GetAttachmentList(CBodyList& rList) const;

public:
	// overrides
	virtual void Clear();
	virtual int GetLength() const;
	// serialization
	virtual int Store(char* pszData, int nMaxSize) const;
	virtual int Load(const char*& pszDataBase, const char* pszData, int nDataSize);

protected:
	bool m_bIsRelated;              // is part of multipart/related block 
	unsigned char* m_pbText;		// content (text) of the body part
	int m_nTextSize;				// length of content
	int m_bodyDataOffset;			// Offset to body data within the mail
	int m_bodyDataLength;			// Lenght of body data within the mail
	CBodyList m_listBodies;			// list of all child body parts
	CBodyList::iterator m_itFind;

protected:
	bool AllocateBuffer(int nBufSize);
	void FreeBuffer();

	friend class CMimeEnvironment;
};

inline bool CMimeBody::IsRelated() {
	return m_bIsRelated;
}

inline int CMimeBody::GetContentLength() const
{ return m_nTextSize; }

inline const unsigned char* CMimeBody::GetContent() const
{ return m_pbText; }

inline bool CMimeBody::IsText() const
{ return GetMediaType() == MEDIA_TEXT; }

inline bool CMimeBody::IsMessage() const
{ return GetMediaType() == MEDIA_MESSAGE; }

#if 0
inline bool CMimeBody::IsAttachment() const
{ return ((GetName().size() > 0) || (GetFilename().size() > 0)); }
#endif

inline bool CMimeBody::IsMultiPart() const
{ return GetMediaType() == MEDIA_MULTIPART; }

inline CMimeBody* CMimeBody::FindFirstPart()
{
	m_itFind = m_listBodies.begin();
	return FindNextPart();
}

inline CMimeBody* CMimeBody::FindNextPart()
{
	if (m_itFind != m_listBodies.end())
		return *m_itFind++;
	return NULL;
}

inline bool CMimeBody::AllocateBuffer(int nBufSize)
{
	FreeBuffer();
	m_pbText = new unsigned char[nBufSize];
	if (!m_pbText)  return false;
	m_nTextSize = nBufSize;
	return true;
}

inline void CMimeBody::FreeBuffer()
{
	delete []m_pbText;
	m_pbText = NULL;
	m_nTextSize = 0;
}

//////////////////////////////////////////////////////////////////////
// CMimeMessage - Represents a MIME message
//////////////////////////////////////////////////////////////////////
class CMimeMessage : public CMimeBody
{
public:
	CMimeMessage() { /*SetVersion();*/ }
	virtual ~CMimeMessage() { Clear(); }

public:
	// set/get RFC 822 message header fields
	void SetFrom(const char* pszFrom, const char* pszCharset=NULL);
	const char* GetFrom() const;
	void SetTo(const char* pszTo, const char* pszCharset=NULL);
	const char* GetTo() const;
	void SetCc(const char* pszCc, const char* pszCharset=NULL);
	const char* GetCc() const;
	void SetBcc(const char* pszBcc, const char* pszCharset=NULL);
	const char* GetBcc() const;

	void SetSubject(const char* pszSubject, const char* pszCharset=NULL);
	const char* GetSubject() const;

	void SetDate();
	void SetDate(int nYear, int nMonth, int nDay, int nHour, int nMinute, int nSecond);
	const char* GetDate() const;
	void SetVersion();
};

inline void CMimeMessage::SetFrom(const char* pszAddr, const char* pszCharset)
{ SetFieldValue("From", pszAddr, pszCharset); }

inline const char* CMimeMessage::GetFrom() const
{ return GetFieldValue("From"); }

inline void CMimeMessage::SetTo(const char* pszAddr, const char* pszCharset)
{ SetFieldValue("To", pszAddr, pszCharset); }

inline const char* CMimeMessage::GetTo() const
{ return GetFieldValue("To"); }

inline void CMimeMessage::SetCc(const char* pszAddr, const char* pszCharset)
{ SetFieldValue("CC", pszAddr, pszCharset); }

inline const char* CMimeMessage::GetCc() const
{ return GetFieldValue("CC"); }

inline void CMimeMessage::SetBcc(const char* pszAddr, const char* pszCharset)
{ SetFieldValue("BCC", pszAddr, pszCharset); }

inline const char* CMimeMessage::GetBcc() const
{ return GetFieldValue("BCC"); }

inline void CMimeMessage::SetSubject(const char* pszSubject, const char* pszCharset)
{ SetFieldValue("Subject", pszSubject, pszCharset); }

inline const char* CMimeMessage::GetSubject() const
{ return GetFieldValue("Subject"); }

inline const char* CMimeMessage::GetDate() const
{ return GetFieldValue("Date"); }

inline void CMimeMessage::SetVersion()
{ SetFieldValue(CMimeConst::MimeVersion(), "1.0"); }

#endif // !defined(_MIME_H)
