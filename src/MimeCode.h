//////////////////////////////////////////////////////////////////////
//
// MIME Encoding/Decoding:
//	Quoted-printable and Base64 for content encoding;
//	Encoded-word for header field encoding.
//
// Jeff Lee
// Dec 11, 2000
//
//////////////////////////////////////////////////////////////////////

#if !defined(_MIME_CODING_H)
#define _MIME_CODING_H

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <utility>
#include <string>
#include <list>
using namespace std;

#pragma warning(disable:4786)	// identifier was truncated to 'number' characters in the debug information

#if !defined(ASSERT)
#if defined(_DEBUG)
	#include <assert.h>
	#define ASSERT(exp)	assert(exp)
#else
	#define ASSERT(exp)	((void)0)
#endif
#endif

#if defined(_DEBUG) && !defined(DEBUG_NEW)
#define DEBUG_NEW new
#endif

// maximum length of an encoded line (RFC 2045)
#define MAX_MIME_LINE_LEN	76
#define MAX_ENCODEDWORD_LEN	75

//////////////////////////////////////////////////////////////////////
// CMimeEnvironment - global environment to manage encoding/decoding
//////////////////////////////////////////////////////////////////////
class CMimeCodeBase;
class CFieldCodeBase;
class CMimeBody;
class CMimeEnvironment
{
public:
	CMimeEnvironment();

public:
	// global options
	static void SetAutoFolding(bool bAutoFolding=true);
	static bool AutoFolding() { return m_bAutoFolding; }
	static void SetGlobalCharset(const char* pszCharset) { m_strCharset = pszCharset; }
	static const char* GetGlobalCharset() { return m_strCharset.c_str(); }

	// Content-Transfer-Encoding coder management
	typedef CMimeCodeBase* (*CODER_FACTORY)();
	static void RegisterCoder(const char* pszCodingName, CODER_FACTORY pfnCreateObject=NULL);
	static CMimeCodeBase* CreateCoder(const char* pszCodingName);

	// header fields encoding/folding management
	typedef CFieldCodeBase* (*FIELD_CODER_FACTORY)();
	static void RegisterFieldCoder(const char* pszFieldName, FIELD_CODER_FACTORY pfnCreateObject=NULL);
	static CFieldCodeBase* CreateFieldCoder(const char* pszFieldName);

	// media type management
	typedef CMimeBody* (*BODY_PART_FACTORY)();
	static void RegisterMediaType(const char* pszMediaType, BODY_PART_FACTORY pfnCreateObject=NULL);
	static CMimeBody* CreateBodyPart(const char* pszMediaType);

private:
	static bool m_bAutoFolding;
	static string m_strCharset;

	typedef pair<const char*, CODER_FACTORY> CODER_PAIR;
	static list<CODER_PAIR> m_listCoders;

	typedef pair<const char*, FIELD_CODER_FACTORY> FIELD_CODER_PAIR;
	static list<FIELD_CODER_PAIR> m_listFieldCoders;

	typedef pair<const char*, BODY_PART_FACTORY> MEDIA_TYPE_PAIR;
	static list<MEDIA_TYPE_PAIR> m_listMediaTypes;

	static CMimeEnvironment m_globalMgr;
};

#define DECLARE_MIMECODER(class_name) \
	public: static CMimeCodeBase* CreateObject() { return new class_name; }

#define REGISTER_MIMECODER(coding_name, class_name) \
	CMimeEnvironment::RegisterCoder(coding_name, class_name::CreateObject)

#define DEREGISTER_MIMECODER(coding_name) \
	CMimeEnvironment::RegisterCoder(coding_name, 0)

#define DECLARE_FIELDCODER(class_name) \
	public: static CFieldCodeBase* CreateObject() { return new class_name; }

#define REGISTER_FIELDCODER(field_name, class_name) \
	CMimeEnvironment::RegisterFieldCoder(field_name, class_name::CreateObject)

#define DEREGISTER_FIELDCODER(field_name) \
	CMimeEnvironment::RegisterFieldCoder(field_name, 0)

#define DECLARE_MEDIATYPE(class_name) \
	public: static CMimeBody* CreateObject() { return new class_name; }

#define REGISTER_MEDIATYPE(media_type, class_name) \
	CMimeEnvironment::RegisterMediaType(media_type, class_name::CreateObject)

#define DEREGISTER_MEDIATYPE(media_type) \
	CMimeEnvironment::RegisterMediaType(media_type, 0)

//////////////////////////////////////////////////////////////////////
// CMimeCodeBase - base class for MIME encoding/decoding
// default implementation for binary/unknown encoding mechanism
//////////////////////////////////////////////////////////////////////
class CMimeCodeBase
{
public:
	CMimeCodeBase() :
		m_pbInput(NULL),
		m_nInputSize(0),
		m_bIsEncoding(false) {}

public:
	void SetInput(const char* pbInput, int nInputSize, bool bEncoding)
	{
		m_pbInput = (const unsigned char*) pbInput;
		m_nInputSize = nInputSize;
		m_bIsEncoding = bEncoding;
	}
	int GetOutputLength() const
	{
		return m_bIsEncoding ? GetEncodeLength() : GetDecodeLength();
	}
	int GetOutput(unsigned char* pbOutput, int nMaxSize)
	{
		return m_bIsEncoding ? Encode(pbOutput, nMaxSize) : Decode(pbOutput, nMaxSize);
	}

protected:
	// overrides
	virtual int GetEncodeLength() const { return m_nInputSize; }
	virtual int GetDecodeLength() const { return m_nInputSize; }
	virtual int Encode(unsigned char* pbOutput, int nMaxSize) const
	{
		int nSize = min(nMaxSize, m_nInputSize);
		::memcpy(pbOutput, m_pbInput, nSize);
		return nSize;
	}
	virtual int Decode(unsigned char* pbOutput, int nMaxSize)
	{
		return CMimeCodeBase::Encode(pbOutput, nMaxSize);
	}

protected:
	const unsigned char* m_pbInput;
	int m_nInputSize;
	bool m_bIsEncoding;
};

//////////////////////////////////////////////////////////////////////
// CMimeCode7bit - for 7bit/8bit encoding mechanism (fold long line)
//////////////////////////////////////////////////////////////////////
class CMimeCode7bit : public CMimeCodeBase
{
	DECLARE_MIMECODER(CMimeCode7bit)

protected:
	virtual int GetEncodeLength() const;
	virtual int Encode(unsigned char* pbOutput, int nMaxSize) const;
};

//////////////////////////////////////////////////////////////////////
// CMimeCodeQP - for quoted-printable encoding mechanism
//////////////////////////////////////////////////////////////////////
class CMimeCodeQP : public CMimeCodeBase
{
public:
	CMimeCodeQP() :
		m_bQuoteLineBreak(false) {}

public:
	DECLARE_MIMECODER(CMimeCodeQP)
	void QuoteLineBreak(bool bQuote=true) { m_bQuoteLineBreak = bQuote; }

protected:
	virtual int GetEncodeLength() const;
	//virtual int GetDecodeLength() const;
	virtual int Encode(unsigned char* pbOutput, int nMaxSize) const;
	virtual int Decode(unsigned char* pbOutput, int nMaxSize);

private:
	bool m_bQuoteLineBreak;
};

//////////////////////////////////////////////////////////////////////
// CMimeCodeBase64 - for base64 encoding mechanism
//////////////////////////////////////////////////////////////////////
class CMimeCodeBase64 : public CMimeCodeBase
{
public:
	CMimeCodeBase64() :
		m_bAddLineBreak(true) {}

public:
	DECLARE_MIMECODER(CMimeCodeBase64)
	void AddLineBreak(bool bAdd=true) { m_bAddLineBreak = bAdd; }

protected:
	virtual int GetEncodeLength() const;
	virtual int GetDecodeLength() const;
	virtual int Encode(unsigned char* pbOutput, int nMaxSize) const;
	virtual int Decode(unsigned char* pbOutput, int nMaxSize);

private:
	bool m_bAddLineBreak;

private:
	static inline int DecodeBase64Char(unsigned int nCode)
	{
		if (nCode >= 'A' && nCode <= 'Z')
			return nCode - 'A';
		if (nCode >= 'a' && nCode <= 'z')
			return nCode - 'a' + 26;
		if (nCode >= '0' && nCode <= '9')
			return nCode - '0' + 52;
		if (nCode == '+')
			return 62;
		if (nCode == '/')
			return 63;
		return 64;
	}
};

//////////////////////////////////////////////////////////////////////
// CMimeEncodedWord - encoded word for non-ascii text (RFC 2047)
//////////////////////////////////////////////////////////////////////
class CMimeEncodedWord : public CMimeCodeBase
{
public:
	CMimeEncodedWord() :
		m_nEncoding(0) {}

	void SetEncoding(int nEncoding, const char* pszCharset)
	{
		m_nEncoding = nEncoding;
		m_strCharset = pszCharset;
	}
	int GetEncoding() const { return m_nEncoding; }
	const char* GetCharset() const { return m_strCharset.c_str(); }

protected:
	virtual int GetEncodeLength() const;
	virtual int Encode(unsigned char* pbOutput, int nMaxSize) const;
	virtual int Decode(unsigned char* pbOutput, int nMaxSize);

private:
	int m_nEncoding;
	string m_strCharset;

	int BEncode(unsigned char* pbOutput, int nMaxSize) const;
	int QEncode(unsigned char* pbOutput, int nMaxSize) const;
};

//////////////////////////////////////////////////////////////////////
// CFieldCodeBase - base class to encode/decode header fields
// default coder for any unregistered fields
//////////////////////////////////////////////////////////////////////
class CFieldCodeBase : public CMimeCodeBase
{
public:
	void SetCharset(const char* pszCharset) { m_strCharset = pszCharset; }
	const char* GetCharset() const { return m_strCharset.c_str(); }

protected:
	string m_strCharset;

	virtual bool IsFoldingChar(char /*ch*/) const { return false; }
	virtual int GetDelimeter() const { return 0; }
	int FindSymbol(const char* pszData, int nSize, int& nDelimeter, int& nNonAscChars) const;
	void UnfoldField(string& strField) const;
	int SelectEncoding(int nLength, int nNonAsciiChars) const
	{
		int nQEncodeSize = nLength + nNonAsciiChars * 2;
		int nBEncodeSize = (nLength + 2) / 3 * 4;
		return (nQEncodeSize <= nBEncodeSize || nNonAsciiChars*5 <= nLength) ? 'Q' : 'B';
	}

protected:
	virtual int GetEncodeLength() const;
	virtual int Encode(unsigned char* pbOutput, int nMaxSize) const;
	virtual int Decode(unsigned char* pbOutput, int nMaxSize);
};

//////////////////////////////////////////////////////////////////////
// CFieldCodeText - encode/decode header fields defined as *text
//////////////////////////////////////////////////////////////////////
class CFieldCodeText : public CFieldCodeBase
{
	DECLARE_FIELDCODER(CFieldCodeText)

protected:
	virtual int GetDelimeter() const { return 0xff; }
};

//////////////////////////////////////////////////////////////////////
// CFieldCodeAddress - encode/decode header fields defined as address
//////////////////////////////////////////////////////////////////////
class CFieldCodeAddress : public CFieldCodeBase
{
	DECLARE_FIELDCODER(CFieldCodeAddress)

protected:
	virtual bool IsFoldingChar(char ch) const { return ch == ',' || ch == ':'; }
};

//////////////////////////////////////////////////////////////////////
// CFieldCodeParameter - encode/decode header fields have parameters
//////////////////////////////////////////////////////////////////////
class CFieldCodeParameter : public CFieldCodeBase
{
	DECLARE_FIELDCODER(CFieldCodeParameter)

protected:
	virtual bool IsFoldingChar(char ch) const { return ch == ';'; }
};

#endif // !defined(_MIME_CODING_H)
