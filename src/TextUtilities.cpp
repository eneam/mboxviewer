#include "stdafx.h"

#ifdef _MSC_VER
#pragma warning (disable : 4786)
#endif

#include "TextUtilities.h"
#include <algorithm>
#include <map>

TextUtilities	g_tu;

#ifdef _MSC_VER
#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif
#endif

static UINT32 _crc32Table[256] = 
{
	0x00000000, 0xBC5C9CBF, 0x9F32D42D, 0x236E4892, 0xD9EE4509, 0x65B2D9B6,
	0x46DC9124, 0xFA800D9B, 0x54576741, 0xE80BFBFE, 0xCB65B36C, 0x77392FD3,
	0x8DB92248, 0x31E5BEF7, 0x128BF665, 0xAED76ADA, 0xA8AECE82, 0x14F2523D,
	0x379C1AAF, 0x8BC08610, 0x71408B8B, 0xCD1C1734, 0xEE725FA6, 0x522EC319,
	0xFCF9A9C3, 0x40A5357C, 0x63CB7DEE, 0xDF97E151, 0x2517ECCA, 0x994B7075,
	0xBA2538E7, 0x0679A458, 0xB6D67057, 0x0A8AECE8, 0x29E4A47A, 0x95B838C5,
	0x6F38355E, 0xD364A9E1, 0xF00AE173, 0x4C567DCC, 0xE2811716, 0x5EDD8BA9,
	0x7DB3C33B, 0xC1EF5F84, 0x3B6F521F, 0x8733CEA0, 0xA45D8632, 0x18011A8D,
	0x1E78BED5, 0xA224226A, 0x814A6AF8, 0x3D16F647, 0xC796FBDC, 0x7BCA6763,
	0x58A42FF1, 0xE4F8B34E, 0x4A2FD994, 0xF673452B, 0xD51D0DB9, 0x69419106,
	0x93C19C9D, 0x2F9D0022, 0x0CF348B0, 0xB0AFD40F, 0x8A270DFD, 0x367B9142,
	0x1515D9D0, 0xA949456F, 0x53C948F4, 0xEF95D44B, 0xCCFB9CD9, 0x70A70066,
	0xDE706ABC, 0x622CF603, 0x4142BE91, 0xFD1E222E, 0x079E2FB5, 0xBBC2B30A,
	0x98ACFB98, 0x24F06727, 0x2289C37F, 0x9ED55FC0, 0xBDBB1752, 0x01E78BED,
	0xFB678676, 0x473B1AC9, 0x6455525B, 0xD809CEE4, 0x76DEA43E, 0xCA823881,
	0xE9EC7013, 0x55B0ECAC, 0xAF30E137, 0x136C7D88, 0x3002351A, 0x8C5EA9A5,
	0x3CF17DAA, 0x80ADE115, 0xA3C3A987, 0x1F9F3538, 0xE51F38A3, 0x5943A41C,
	0x7A2DEC8E, 0xC6717031, 0x68A61AEB, 0xD4FA8654, 0xF794CEC6, 0x4BC85279,
	0xB1485FE2, 0x0D14C35D, 0x2E7A8BCF, 0x92261770, 0x945FB328, 0x28032F97,
	0x0B6D6705, 0xB731FBBA, 0x4DB1F621, 0xF1ED6A9E, 0xD283220C, 0x6EDFBEB3,
	0xC008D469, 0x7C5448D6, 0x5F3A0044, 0xE3669CFB, 0x19E69160, 0xA5BA0DDF,
	0x86D4454D, 0x3A88D9F2, 0xF3C5F6A9, 0x4F996A16, 0x6CF72284, 0xD0ABBE3B,
	0x2A2BB3A0, 0x96772F1F, 0xB519678D, 0x0945FB32, 0xA79291E8, 0x1BCE0D57,
	0x38A045C5, 0x84FCD97A, 0x7E7CD4E1, 0xC220485E, 0xE14E00CC, 0x5D129C73,
	0x5B6B382B, 0xE737A494, 0xC459EC06, 0x780570B9, 0x82857D22, 0x3ED9E19D,
	0x1DB7A90F, 0xA1EB35B0, 0x0F3C5F6A, 0xB360C3D5, 0x900E8B47, 0x2C5217F8,
	0xD6D21A63, 0x6A8E86DC, 0x49E0CE4E, 0xF5BC52F1, 0x451386FE, 0xF94F1A41,
	0xDA2152D3, 0x667DCE6C, 0x9CFDC3F7, 0x20A15F48, 0x03CF17DA, 0xBF938B65,
	0x1144E1BF, 0xAD187D00, 0x8E763592, 0x322AA92D, 0xC8AAA4B6, 0x74F63809,
	0x5798709B, 0xEBC4EC24, 0xEDBD487C, 0x51E1D4C3, 0x728F9C51, 0xCED300EE,
	0x34530D75, 0x880F91CA, 0xAB61D958, 0x173D45E7, 0xB9EA2F3D, 0x05B6B382,
	0x26D8FB10, 0x9A8467AF, 0x60046A34, 0xDC58F68B, 0xFF36BE19, 0x436A22A6,
	0x79E2FB54, 0xC5BE67EB, 0xE6D02F79, 0x5A8CB3C6, 0xA00CBE5D, 0x1C5022E2,
	0x3F3E6A70, 0x8362F6CF, 0x2DB59C15, 0x91E900AA, 0xB2874838, 0x0EDBD487,
	0xF45BD91C, 0x480745A3, 0x6B690D31, 0xD735918E, 0xD14C35D6, 0x6D10A969,
	0x4E7EE1FB, 0xF2227D44, 0x08A270DF, 0xB4FEEC60, 0x9790A4F2, 0x2BCC384D,
	0x851B5297, 0x3947CE28, 0x1A2986BA, 0xA6751A05, 0x5CF5179E, 0xE0A98B21,
	0xC3C7C3B3, 0x7F9B5F0C, 0xCF348B03, 0x736817BC, 0x50065F2E, 0xEC5AC391,
	0x16DACE0A, 0xAA8652B5, 0x89E81A27, 0x35B48698, 0x9B63EC42, 0x273F70FD,
	0x0451386F, 0xB80DA4D0, 0x428DA94B, 0xFED135F4, 0xDDBF7D66, 0x61E3E1D9,
	0x679A4581, 0xDBC6D93E, 0xF8A891AC, 0x44F40D13, 0xBE740088, 0x02289C37,
	0x2146D4A5, 0x9D1A481A, 0x33CD22C0, 0x8F91BE7F, 0xACFFF6ED, 0x10A36A52,
	0xEA2367C9, 0x567FFB76, 0x7511B3E4, 0xC94D2F5B
};

UINT32 TextUtilities::CalcCRC32(const char* buf, const UINT length)
{
	register UINT32 crc32 = 0;
	register UINT32* crc32Table = _crc32Table;

	register UINT32 len = length;
	while (len-- > 0)
		crc32 = crc32Table[(crc32 ^ *buf++) & 0xFF] ^ (crc32 >> 8);

	return crc32;
}

int TextUtilities::BMHSearchW( unsigned char *text, int n, unsigned char *pat, int m, BOOL bCaseSens )   /* Search Search pat[0..m-1] in text[0..n-1] */
{
    char d[256];
    int i, j, k, lim;
    int res = -1;
    memset(d, m+1, 256);
    for( k=0; k<m; k++ )
        d[pat[k]] = m-k;
    pat[m] = 1;          /* To avoid having code     */ /* for special case n-k+1=m */ /* null terminator overwrite, restore before exit */
	// TODO original lim = n - m + 1; // // out of range and crash under debugger when searching RAW message
	// TODO corrected  lim = n - m; doesn't crash but matching fails in some cases, restored original lim = n-m+1; 
	/* for special case n-k+1=m */
	lim = n - m + 1; // // out of range and may crash under debugger

	if( bCaseSens ) {
		for( k=0; k < lim; k += d[(unsigned char)(text[k+m])] ) /* Searching */
		{
			i=k;
			for( j=0; (unsigned char)(text[i]) == pat[j]; j++ ) {
				i++; /* Could be optimal order */
			}
			if( j == m ) {
				if( k != 0 ) {
					if( IsWordBegin(text[k-1]) )
						continue;
				}
				if( k+m < n ) {
					if( IsWordBegin(text[k+m]) )
						continue;
				}
				res = k;
				break;
			}
		}
	} else {
		for( k=0; k < lim; k += d[(unsigned char)ToLower(text[k+m])] ) /* Searching */
		{
			i=k;
			for( j=0; (unsigned char)ToLower(text[i]) == pat[j]; j++ ) {
				i++; /* Could be optimal order */
			}
			if( j == m ) {
				if( k != 0 ) {
					if( IsWordBegin(text[k-1]) )
						continue;
				}
				if( k+m < n ) {
					if( IsWordBegin(text[k+m]) )
						continue;
				}
				res = k;
				break;
			}
		}
	}
    pat[m] = 0;  // restore null terminator
    return res;
}
int TextUtilities::BMHSearch( unsigned char *text, int n, unsigned char *pat, int m, BOOL bCaseSens )   /* Search Search pat[0..m-1] in text[0..n-1] */
{
    unsigned char d[256];
    int i, j, k, lim;
    int res = -1;
    memset(d, m+1, 256);
    for( k=0; k<m; k++ )
        d[pat[k]] = m-k;
    pat[m] = 1;          /* To avoid having code     */  /* for special case n-k+1=m */
	// TODO original lim = n - m + 1; // // out of range and crash under debugger when searching RAW message
	// TODO corrected  lim = n - m; doesn't crash but matching fails in some cases, restored original lim = n-m+1; 
    /* for special case n-k+1=m */
    lim = n-m+1; // // out of range and may crash under debugger

 	if( bCaseSens ) {
		for( k=0; k < lim; k += d[(unsigned char)(text[k+m])] ) /* Searching */
		{
			i=k;
			for( j=0; (unsigned char)(text[i]) == pat[j]; j++ ) {
				i++; /* Could be optimal order */
			}
			if( j == m ) {
				res = k; 
				break;
			}
		}
	} else {
		for( k=0; k < lim; k += d[(unsigned char)ToLower(text[k+m])] ) /* Searching */
		{
			i=k;
			for (j = 0; (unsigned char)ToLower(text[i]) == pat[j]; j++) {
				i++; /* Could be optimal order */
			}
			if( j == m ) {
				res = k; 
				break;
			}
		}
	}
    pat[m] = 0; 
    return res;
}
#define WordMinSize  2       //  The minimum and maximum lengths a word must be
#define WordMaxSize  40      //  in order even to bother doing more aggressive checks
                                    //  on to determine if it should be indexed.

// I don't think there is a word in English that has more than...
#define WordMaxConsecConsonants  4
#define WordMaxConsecVowels      5
#define WordMaxConsecSame        2
#define WordMaxConsecPuncts      2

static const char *SpaceChars = " \r\n\t";
static bool SpaceCharsTable[256];
static const char AlfaChars[] =      "abcdefghijklmnopqrstuvwxyz_ABCDEFGHIJKLMNOPQRSTUVWXYZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ";
static bool AlfaCharsTable[256];
static const char WordChars[] =      "0123456789abcdefghijklmnopqrstuvwxyz_ABCDEFGHIJKLMNOPQRSTUVWXYZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ";
static bool WordCharsTable[256];
static const char WordBeginChars[] = "0123456789abcdefghijklmnopqrstuvwxyz_ABCDEFGHIJKLMNOPQRSTUVWXYZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ";
static bool WordBeginCharsTable[256];
static const char WordEndChars[] =   "0123456789abcdefghijklmnopqrstuvwxyz_ABCDEFGHIJKLMNOPQRSTUVWXYZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ";
static bool WordEndCharsTable[256];
static const char LowerCaseChars[] =  "abcdefghijklmnopqrstuvwxyzàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ";
static bool LowerCaseCharsTable[256];
static const char UpperCaseChars[] =  "ABCDEFGHIJKLMNOPQRSTUVWXYZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞß";
static bool UpperCaseCharsTable[256];
static const char VowelChars[] =      "aeiouAEIOUÀÁÂÃÄÅÆÈÉÊËÌÍÎÏÒÓÔÕÖØÙÚÛÜàáâãäåæèéêëìíîïðòóôõöøùúûü";
static bool VowelCharsTable[256];
static const char ApoChars[] =      "'`´";
static bool ApoCharsTable[256];

static const char DigitChars[] =      "0123456789";
static bool DigitCharsTable[256];
static const char XDigitChars[] =      "0123456789abcdefABCDEF";
static bool XDigitCharsTable[256];

static char LowerToUpperCharsTable[256];
static char UpperToLowerCharsTable[256];

//
//
// All characters are from the ISO 8859-1 character set mapped to 7-bit ASCII.
//
static char const iso8859_map[] = {
	' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ',	//  0
	' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ',	//  8
	' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ',	//  16
	' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ',	//  24
	' ', '!', '"', '#', '$', '%', '&', '\'',//  32
	'(', ')', '*', '+', ',', '-', '.', '/',	//  40
	'0', '1', '2', '3', '4', '5', '6', '7',	//  48
	'8', '9', ':', ';', '<', '=', '>', '?',	//  56
	'@', 'A', 'B', 'C', 'D', 'E', 'F', 'G',	//  64
	'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O',	//  72
	'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W',	//  80
	'X', 'Y', 'Z', '[', '\\', ']', '^', '_',//  88
	'`', 'a', 'b', 'c', 'd', 'e', 'f', 'g',	//  96
	'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o',	// 104
	'p', 'q', 'r', 's', 't', 'u', 'v', 'w',	// 112
	'x', 'y', 'z', '{', '|', '}', '~', ' ',	// 120
	' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ',	// 128
	' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ',	// 136
	' ', '\'','\'', '"', '"', ' ', ' ', ' ',	// 144
	' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ',	// 152
	' ', '!', ' ', '#', ' ', ' ', '|', ' ',	// 160
	' ', ' ', ' ', '<', ' ', '-', ' ', ' ',	// 168
	' ', ' ', '2', '3','\'', ' ', ' ', '.', // 176
	' ', '1', ' ', '>', ' ', ' ', ' ', '?',	// 184
	'À', 'Á', 'Â', 'Ã', 'Ä', 'Å', 'Æ', 'Ç',	// 192
	'È', 'É', 'Ê', 'Ë', 'Ì', 'Í', 'Î', 'Ï',	// 200
	'Ð', 'Ñ', 'Ò', 'Ó', 'Ô', 'Õ', 'Ö', ' ',	// 208
	'Ø', 'Ù', 'Ú', 'Û', 'Ü', 'Ý', 'Þ', 'ß',	// 216
	'à', 'á', 'â', 'ã', 'ä', 'å', 'æ', 'ç',	// 224
	'è', 'é', 'ê', 'ë', 'ì', 'í', 'î', 'ï',	// 232
	'ð', 'ñ', 'ò', 'ó', 'ô', 'õ', 'ö', ' ',	// 240
	'ø', 'ù', 'ú', 'û', 'ü', 'ý', 'þ', 'ÿ',	// 248
};


void TextUtilities::MakeLookupTable(const char *s, bool *l)
{
	memset(l, false, 256 * sizeof(*l));
	for (; *s; s++)
		l[(int) ((unsigned char) *s)] = true;
}

void TextUtilities::MakeConversionTable( const char *from, const char *to, char *dest )
{
	ASSERT(strlen(from) == strlen(to));
	//memset(dest, 0, 256 * sizeof(*dest));
	for (int i = 0; i < 256; i++)
		dest[i] = i;
	for (; *from; from++, to++)
		dest[(int) ((unsigned char) *from)] = *to;
}

static std::vector<CString>* CreateArray(const char *arr[])
{
	std::vector<CString>* cm = new std::vector<CString>;
	const char **mp = arr;
	while (**mp) {
		cm->push_back(*mp);
		mp++;
	}
	return cm;
}

bool TextUtilities::Init()
{
	MakeLookupTable(SpaceChars, SpaceCharsTable);
	MakeLookupTable(AlfaChars, AlfaCharsTable);
	MakeLookupTable(WordChars, WordCharsTable);
	MakeLookupTable(WordBeginChars, WordBeginCharsTable);
	MakeLookupTable(WordEndChars, WordEndCharsTable);
	MakeLookupTable(LowerCaseChars, LowerCaseCharsTable);
	MakeLookupTable(UpperCaseChars, UpperCaseCharsTable);
	MakeLookupTable(VowelChars, VowelCharsTable);
	MakeLookupTable(ApoChars, ApoCharsTable);
	MakeLookupTable(DigitChars, DigitCharsTable);
	MakeLookupTable(XDigitChars, XDigitCharsTable);
	MakeConversionTable( LowerCaseChars, UpperCaseChars, LowerToUpperCharsTable );
	MakeConversionTable( UpperCaseChars, LowerCaseChars, UpperToLowerCharsTable );
	return true;
}

TextUtilities::~TextUtilities() {

}

#define HASH_LOOKUP(c, tbl) (tbl[(int) (0x000000ff & c)])

bool TextUtilities::IsSpace(const char c) {
	return HASH_LOOKUP(c, SpaceCharsTable);
}

bool TextUtilities::IsAlfa(const char c) {
	return HASH_LOOKUP(c, AlfaCharsTable);
}

bool TextUtilities::IsWord(const char c) {
	return HASH_LOOKUP(c, WordCharsTable);
}

bool TextUtilities::IsWordBegin(const char c) {
	return HASH_LOOKUP(c, WordBeginCharsTable);
}

bool TextUtilities::IsWordEnd(const char c) {
	return HASH_LOOKUP(c, WordEndCharsTable);
}

bool TextUtilities::IsLowerCase(const char c) {
	return HASH_LOOKUP(c, LowerCaseCharsTable);
}

bool TextUtilities::IsUpperCase(const char c) {
	return HASH_LOOKUP(c, UpperCaseCharsTable);
}

bool TextUtilities::IsVowel(const char c) {
	return HASH_LOOKUP(c, VowelCharsTable);
}

bool TextUtilities::IsDigit( const char c ) {
	return HASH_LOOKUP(c, DigitCharsTable);
}

bool TextUtilities::IsXDigit( const char c ) {
	return HASH_LOOKUP(c, XDigitCharsTable);
}

bool TextUtilities::IsAlfaNum( char c ) {
    return IsDigit( c ) || IsAlfa( c );
}

bool TextUtilities::IsApo( char c ) {
    return HASH_LOOKUP(c, ApoCharsTable);
}

bool TextUtilities::IsPunct( const char c ) {
	return ( ! IsAlfaNum(c) && ! IsSpace(c) );
}

char TextUtilities::ToLower( char c ) {
	return HASH_LOOKUP(c, UpperToLowerCharsTable);
}

char TextUtilities::ToUpper( char c ) {
    return HASH_LOOKUP(c, LowerToUpperCharsTable);
}

void TextUtilities::MakeLower( register char *s ) {
	while( *s ) {
		*s = ToLower(*s);
		s++;
	}
}

void TextUtilities::MakeLower( CString &s ) {
	MakeLower(s.GetBuffer(s.GetLength()));
}

void TextUtilities::MakeUpper( register char *s ) {
	while( *s ) {
		*s = ToUpper(*s);
		s++;
	}
}

void TextUtilities::MakeUpper( CString &s ) {
	MakeUpper(s.GetBuffer(s.GetLength()));
}

void TextUtilities::CapitalizeWords( register char *s, int bLower )
{
	while( *s ) {
		while( ! IsWordBegin(*s) )
			s++;
		if( *s ) {
			*s = ToUpper(*s);
			s++;
		}
		while( IsWord(*s) ) {
			if( bLower == 1 )
				*s = ToLower(*s);
			s++;
		}
		s++;
	}
}

bool TextUtilities::IsAllCaps( register const char *s ) {
	while( *s ) {
		if( IsAlfa(*s) && IsLowerCase(*s) )
			break;
		s++;
	}
	return *s == 0;
}

#define N_CHARS		16			/* number of bytes printed in one line */
#define	MAX_BUFF	4094		/* size of print buffer for sprintf() and max TRACE can accept */

void TextUtilities::hexdump(char *title, char *area, int length)
{
	char	buff[MAX_BUFF + 1];
	char	tmp[128];
	size_t	cnt, n, len;
	int	m_length = length;
	int	jj, ii;
	char	ch;

	cnt = 0;

	if (strlen(title) > 0)
	{
		(void *)sprintf(tmp, "%s (length=%d)", title, length);
		(void *)sprintf(&buff[cnt], "%-22s = \n", tmp);
		n = strlen(&buff[cnt]);
		if (n > 0) cnt += n;
	}

	for (jj = 0; jj < m_length; jj += N_CHARS) /* print block of 16 bytes */
	{
		len = m_length - jj;			/* bytes left */
		if (len > N_CHARS)
			len = N_CHARS;

		(void *)sprintf(&buff[cnt], "%04d  ", (jj % 10000));
		n = strlen(&buff[cnt]);
		if (n > 0) cnt += n;

		for (ii = 0; ii < len; ++ii)			/* N_CHARS bytes hex */
		{
			if (ii == N_CHARS / 2)
			{
				(void *)sprintf(&buff[cnt], "| ");
				n = strlen(&buff[cnt]);
				if (n > 0) cnt += n;
			}
			(void *)sprintf(&buff[cnt], "%02x ", (unsigned char)area[jj + ii]);
			n = strlen(&buff[cnt]);
			if (n > 0) cnt += n;
		}

		for (ii = len; ii < N_CHARS; ii++)		/* padd with spaces */
		{
			if (ii == N_CHARS / 2)
			{
				(void *)sprintf(&buff[cnt], "  ");
				n = strlen(&buff[cnt]);
				if (n > 0) cnt += n;
			}
			(void *)sprintf(&buff[cnt], "   ");
			n = strlen(&buff[cnt]);
			if (n > 0) cnt += n;
		}

		(void *)sprintf(&buff[cnt], "      ");		/* 6 spaces */
		n = strlen(&buff[cnt]);
		if (n > 0) cnt += n;

		for (ii = 0; ii < len; ii++)			/* N_CHARS bytes ascii */
		{
			ch = area[jj + ii];
			if (ch < ' ' || ch > '~')
				ch = '_';

			(void *)sprintf(&buff[cnt], "%c", ch);
			n = strlen(&buff[cnt]);
			if (n > 0) cnt += n;
		}

		(void *)sprintf(&buff[cnt], "\n");
		n = strlen(&buff[cnt]);
		if (n > 0) cnt += n;

		/* print accumulated characters */
		if (cnt > (MAX_BUFF - (N_CHARS*20)))
		{
			buff[cnt] = '\0';
			TRACE("%s", buff);
			cnt = 0;
			(void *)sprintf(&buff[cnt], "\n");
			n = strlen(&buff[cnt]);
			if (n > 0) cnt += n;
		}
	}
	/*
		(void *) sprintf(&buff[cnt],"\n");
		n = strlen(&buff[cnt]);
		if ( n > 0 ) cnt += n;
	*/

	if (cnt >= 0)
	{
		buff[cnt] = '\0';
		TRACE("%s", buff);
	}
}



bool TextUtilities::TestAll()
{
	char *s1 = "mbox viewer bigniewZxyz";
	s1 = "mboxview";
	int s1Len = strlen(s1);

	char *pat1 = "ewer";
	char *pat2 = "zbigniewZ";
	pat1 = s1;
	pat2 = "z";

	int pat1Len = strlen(pat1);
	int pat2Len = strlen(pat2);

	char *s11 = new char[s1Len];
	memcpy(s11, s1, s1Len);

	BOOL caseSensitive = TRUE;

	char *pat11 = new char[pat1Len + 1];
	strcpy(pat11, pat1);

	char *pat22 = new char[pat2Len + 1];
	strcpy(pat22, pat2);

	TRACE(_T("pat1=%s pat2=%s\n"), pat11, pat22);

	int pos0 = g_tu.BMHSearch((unsigned char*)s11, s1Len, (unsigned char*)pat11, pat1Len, caseSensitive);

	int pos1 = g_tu.BMHSearch((unsigned char*)s1, s1Len, (unsigned char*)pat11, pat1Len, caseSensitive);

	int pos2 = g_tu.BMHSearch((unsigned char*)s1, s1Len, (unsigned char*)pat22, pat2Len, caseSensitive);

	caseSensitive = FALSE;
	int pos11 = g_tu.BMHSearch((unsigned char*)s1, s1Len, (unsigned char*)pat11, pat1Len, caseSensitive);


	//int BMHSearchW(unsigned char *text, int n, unsigned char *pat, int m, BOOL bCaseSens = FALSE);

	delete[] pat11;
	delete[] pat22;

	int deb = 1;
	return true;
}

// class MyCArray
template<class T> void MyCArray<T>::SetSizeKeepData(INT_PTR nNewSize, INT_PTR nGrowBy)
{
	if (nNewSize == 0)
		m_nSize = 0;
	else
		SetSize(nNewSize, nGrowBy);
}

template<class T> void MyCArray<T>::CopyKeepData(const MyCArray<T>& src)
{
	if (src.GetSize() == 0)
		m_nSize = 0;
	else
		Copy(src);
}

#if 0
////////////////////
extern "C"
{
	typedef char gchar;
	typedef uint32_t guint32;

	//GLIB_VAR const guint16 * const g_ascii_table;

	//#define g_ascii_islower(c) ((g_ascii_table[(guchar) (c)] & G_ASCII_LOWER) != 0)



	gchar
		g_ascii_toupper(gchar c)
	{
		//return g_ascii_islower(c) ? c - 'a' + 'A' : c;
		return  g_tu.ToUpper(c);
	}


	gchar
		g_ascii_tolower(gchar c)
	{
		//return g_ascii_isupper(c) ? c - 'A' + 'a' : c;
		return  g_tu.ToLower(c);
	}



	/* this decodes rfc2047's version of quoted-printable */
	static size_t
		quoted_decode(const unsigned char *in, size_t len, unsigned char *out, int *state, guint32 *save)
	{
		register const unsigned char *inptr;
		register unsigned char *outptr;
		const unsigned char *inend;
		unsigned char c, c1;
		guint32 saved;
		int need;

		if (len == 0)
			return 0;

		inend = in + len;
		outptr = out;
		inptr = in;

		need = *state;
		saved = *save;

		if (need > 0) {
			if (isxdigit((int)*inptr)) {
				if (need == 1) {
					c = g_ascii_toupper((int)(saved & 0xff));
					c1 = g_ascii_toupper((int)*inptr++);
					saved = 0;
					need = 0;

					goto decode;
				}

				saved = 0;
				need = 0;

				goto equals;
			}

			/* last encoded-word ended in a malformed quoted-printable sequence */
			*outptr++ = '=';

			if (need == 1)
				*outptr++ = (char)(saved & 0xff);

			saved = 0;
			need = 0;
		}

		while (inptr < inend) {
			c = *inptr++;
			if (c == '=') {
			equals:
				if (inend - inptr >= 2) {
					if (isxdigit((int)inptr[0]) && isxdigit((int)inptr[1])) {
						c = g_ascii_toupper(*inptr++);
						c1 = g_ascii_toupper(*inptr++);
					decode:
						*outptr++ = (((c >= 'A' ? c - 'A' + 10 : c - '0') & 0x0f) << 4)
							| ((c1 >= 'A' ? c1 - 'A' + 10 : c1 - '0') & 0x0f);
					}
					else {
						/* malformed quoted-printable sequence? */
						*outptr++ = '=';
					}
				}
				else {
					/* truncated payload, maybe it was split across encoded-words? */
					if (inptr < inend) {
						if (isxdigit((int)*inptr)) {
							saved = *inptr;
							need = 1;
							break;
						}
						else {
							/* malformed quoted-printable sequence? */
							*outptr++ = '=';
						}
					}
					else {
						saved = 0;
						need = 2;
						break;
					}
				}
			}
			else if (c == '_') {
				/* _'s are an rfc2047 shortcut for encoding spaces */
				*outptr++ = ' ';
			}
			else {
				*outptr++ = c;
			}
		}

		*state = need;
		*save = saved;

		return (size_t)(outptr - out);
	}
}

int test_enc()
{
	//char *word = "=?UTF-8?Q?St=c3=a9phane_Scudeller?= <sscudeller@gmail.com>";
	char *word = "St=c3=a9phane_Scudeller?= <sscudeller@gmail.com>";
	const unsigned char *in = (const unsigned char*)word;

	guint32 save = 0;
	int state = 0;
	unsigned char buff[10000];
	unsigned char *out = buff;
	size_t len = strlen((char*)in);


	size_t ret = quoted_decode(in, len, out, &state, &save);

	return 1;
}

#endif