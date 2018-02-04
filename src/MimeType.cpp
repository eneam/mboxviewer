//////////////////////////////////////////////////////////////////////
//
// MIME message encoding/decoding
//
// Jeff Lee
// Dec 22, 2000
//
//////////////////////////////////////////////////////////////////////
#include "stdafx.h"
#include "Mime.h"

// Content-Type header field specifies the media type of a body part. it coule be:
// text/image/audio/vedio/application (discrete type) or message/multipart (composite type).
// the default Content-Type is: text/plain; charset=us-ascii (RFC 2046)

const char* CMimeHeader::m_TypeTable[] =
{
	"text", "image", "audio", "vedio", "application", "multipart", "message", NULL
};

const CMimeHeader::MediaTypeCvt CMimeHeader::m_TypeCvtTable[] =
{
	// media-type, sub-type, file extension
	{ MEDIA_APPLICATION, "xml", "xml" },
	{ MEDIA_APPLICATION, "msword", "doc" },
	{ MEDIA_APPLICATION, "rtf", "rtf" },
	{ MEDIA_APPLICATION, "vnd.ms-excel", "xls" },
	{ MEDIA_APPLICATION, "vnd.ms-powerpoint", "ppt" },
	{ MEDIA_APPLICATION, "pdf", "pdf" },
	{ MEDIA_APPLICATION, "zip", "zip" },

	{ MEDIA_IMAGE, "jpeg", "jpeg" },
	{ MEDIA_IMAGE, "jpeg", "jpg" },
	{ MEDIA_IMAGE, "gif", "gif" },
	{ MEDIA_IMAGE, "tiff", "tif" },
	{ MEDIA_IMAGE, "tiff", "tiff" },

	{ MEDIA_AUDIO, "basic", "wav" },
	{ MEDIA_AUDIO, "basic", "mp3" },

	{ MEDIA_VEDIO, "mpeg", "mpg" },
	{ MEDIA_VEDIO, "mpeg", "mpeg" },

	{ MEDIA_UNKNOWN, "", "" }		// add new subtypes before this line
};
