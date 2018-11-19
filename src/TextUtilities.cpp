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