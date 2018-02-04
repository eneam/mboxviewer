#include "stdafx.h"

#ifdef _MSC_VER
#pragma warning (disable : 4786)
#endif

//#define _DBG_WORDS

#include "TextUtilities.h"
#include <algorithm>
#include <map>
//std::map<int, std::vector<CString>*> stopwd_map;

TextUtilities	g_tu;

#ifdef _MSC_VER
#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif
#endif

static bool IsValidNumber( CString &w, CString &prevW )
{
	// Check for numbers
//	if( w.GetLength() > 0 && is_digit(w[0]) && ( ( w.GetLength() == 2 && is_digit(w[1]) ) || w.GetLength() == 1 ) ) {
		if( strncmp( prevW, "art", 3 ) == 0 )
			return true;
//	}
	return false;
}

int TextUtilities::BMHSearchW( unsigned char *text, int n, unsigned char *pat, int m, BOOL bCaseSens )   /* Search Search pat[0..m-1] in text[0..n-1] */
{
    char d[256];
    int i, j, k, lim;
    int res = -1;
    memset(d, m+1, 256);
    for( k=0; k<m; k++ )
        d[pat[k]] = m-k;
    pat[m] = 1;          /* To avoid having code     */
    /* for special case n-k+1=m */
    lim = n-m+1;
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
    pat[m] = 0;
    return res;
}
int TextUtilities::BMHSearch( unsigned char *text, int n, unsigned char *pat, int m, BOOL bCaseSens )   /* Search Search pat[0..m-1] in text[0..n-1] */
{   
    char d[256];
    int i, j, k, lim;
    int res = -1;
    memset(d, m+1, 256);
    for( k=0; k<m; k++ )
        d[pat[k]] = m-k;
    pat[m] = 1;          /* To avoid having code     */
    /* for special case n-k+1=m */
    lim = n-m+1;
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
			for( j=0; (unsigned char)ToLower(text[i]) == pat[j]; j++ ) {
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
/*	struct char_entity {
		const char * name;
		char        char_equiv;
	};
	static char_entity const char_entity_table[] = {
		{"amp",    '&'},
		{"Aacute", 'Á'}, {"aacute", 'á'},
		{"Acirc",  'Â'}, {"acirc",  'â'},
		{"AElig",  'Æ'}, {"aelig",  'æ'},
		{"Agrave", 'À'}, {"agrave", 'à'},
		{"Aring",  'Å'}, {"aring",  'å'},
		{"Atilde", 'Ã'}, {"atilde", 'ã'},
		{"Auml",   'Ä'}, {"auml",   'ä'},
		{"Ccedil", 'Ç'}, {"ccedil", 'ç'},
		{"Eacute", 'É'}, {"eacute", 'é'},
		{"Ecirc",  'Ê'}, {"ecirc",  'ê'},
		{"Egrave", 'È'}, {"egrave", 'è'},
		{"ETH",    'Ð'}, {"eth",    'ð'},
		{"Euml",   'Ë'}, {"euml",   'ë'},
		{"Iacute", 'Í'}, {"iacute", 'í'},
		{"Icirc",  'Î'}, {"icirc",  'î'},
		{"Igrave", 'Ì'}, {"igrave", 'ì'},
		{"Iuml",   'Ï'}, {"iuml",   'ï'},
		{"Ntilde", 'Ñ'}, {"ntilde", 'ñ'},
		{"Oacute", 'Ó'}, {"oacute", 'ó'},
		{"Ocirc",  'Ô'}, {"ocirc",  'ô'},
		{"Ograve", 'Ò'}, {"ograve", 'ò'},
		{"Oslash", 'Ø'}, {"slash", 'ø'},
		{"Otilde", 'Õ'}, {"otilde", 'õ'},
		{"Ouml",   'Ö'}, {"ouml",   'ö'},
		{"szlig",  'ß'},
		{"Uacute", 'Ú'}, {"uacute", 'ú'},
		{"Ucirc",  'Û'}, {"ucirc",  'û'},
		{"Ugrave", 'Ù'}, {"ugrave", 'ù'},
		{"Uuml",   'Ü'}, {"uuml",   'ü'},
		{"Yacute", 'Ý'}, {"yacute", 'ý'},
		{"yuml",   'ÿ'},
		{0,   0}
	};
	static const char *it_stopwd[] = { // 1 italian
		"ad", "agli", "ai", "al", "alla", "alle", "allo", "altro", "anche", "che", "ci", "ciò", "come", "con", "così",
		"cui", "da", "dagli", "dai", "dal", "dalla", "dalle", "de", "degli", "dei", "del", "della", "delle", "dello",
		"di", "dopo", "dove", "ed", "era", "fa", "fra", "già", "gli", "ha", "hanno", "ho", "il", "in", "la", "le",
		"lo", "loro", "ma", "mi", "ne", "nei", "nel", "nella", "nelle", "no", "non", "ogni", "per", "perché", "perciò",
		"però", "più", "po", "poi", "pur", "può", "quando", "quegli", "quei", "quel", "quella", "quelle", "quelli",
		"quello", "quest", "questa", "queste", "questi", "questo", "qui", "se", "si", "sia", "solo", "sono", "st",
		"stata", "stesso", "su", "sua", "sue", "sugli", "sui", "sul", "sulla", "sulle", "suo", "sé", "tra", "tutta",
		"tutte", "tutti", "tutto", "un", "una", "uno", "vi", "via", ""
	 };
	static const char *fr_stopwd[] = { // 2 french
		"ans", "au", "aussi", "aux", "avec", "ce", "ces", "cette", "comme", "dans", "de", "des", "deux", "du", "elle",
		"en", "est", "et", "fait", "il", "ils", "je", "la", "le", "les", "leur", "lui", "là", "mais", "même", "ne", "nous",
		"on", "ont", "ou", "par", "pas", "plus", "pour", "qu", "que", "qui", "sa", "sans", "se", "ses", "si", "son", "sont",
		"sur", "un", "une", "été", "être", ""
	};
	static const char *es_stopwd[] = { // 3 spanish
		"al", "como", "con", "de", "del", "el", "en", "es", "esta", "este", "fue", "ha", "la", "las", "le", "lo", "los",
		"más", "no", "para", "pero", "por", "que", "se", "si", "sin", "su", "sus", "un", "una", "ya", ""
	};
	static const char *en_stopwd[] = { // 4 english
		"000", "about", "after", "all", "also", "an", "and", "any", "are", "as", "at", "be", "been", "but", "by", "can",
		"could", "did", "do", "first", "for", "from", "get", "had", "has", "have", "he", "her", "him", "his", "if", "in",
		"into", "is", "it", "its", "just", "last", "lot", "may", "me", "more", "most", "my", "new", "no", "not", "now",
		"of", "on", "one", "or", "other", "our", "out", "over", "said", "she", "so", "some", "than", "that", "the", "their",
		"them", "then", "there", "they", "this", "time", "to", "two", "up", "us", "was", "we", "were", "what", "when",
		"which", "who", "will", "with", "would", "year", "years", "yes", "you", "your", ""
	};
	static const char *de_stopwd[] = { // 6 german
		"aber", "als", "am", "an", "auch", "auf", "aus", "bei", "bis", "da", "das", "dass", "dem", "den", "der",
		"deren", "derer", "dergleiche", "des", "die", "dies", "diese", "ein", "eine", "einem", "einen", "einer",
		"er", "es", "für", "habe", "haben", "habt", "hat", "hatte", "hatten", "hattest", "ihr", "ihre", "ihrem",
		"ihren", "ihrer", "ihres", "im", "in", "ist", "kann", "kannen", "kannst", "konnte", "konnten", "man",
		"mit", "nach", "nicht", "noch", "nur", "oder", "sein", "seind", "seine", "seinem", "seinen", "seiner",
		"seines", "sich", "sie", "sind", "so", "um", "und", "vom", "von", "vor", "werden", "wie", "wird", "zu",
		"zum", "zur", "über", ""
	};
	static const char *se_stopwd[] = { // 7 swedish
		"att", "av", "de", "den", "det", "du", "en", "ett", "från", "får", "för", "han", "har", "hon", "inte", "jag",
		"kan", "man", "med", "men", "mer", "nu", "när", "och", "om", "på", "se", "sig", "ska", "som", "säger", "så",
		"till", "ut", "var", "vi", "än", "är", "år", ""
	};
	static const char *ro_stopwd[] = { // 9 romanian
		"al", "am", "ar", "au", "ca", "care", "ce", "cu", "de", "din", "este", "fi", "fost", "in", "la", "mai", "nu",
		"pe", "pentru", "sa", "se", "si", "un", "va", ""
	};
	static const char *ne_stopwd[] = { // 130 netherlands
		"dat", "de", "die", "een", "en", "er", "het", "ik", "in", "is", "om", "ook", "op", "van", "voor", ""
	};

	for ( register char_entity const *e = char_entity_table; e->name; ++e )
		entity_map[ e->name ] = e->char_equiv & 0x000000ff;

	stopwd_map[1] = CreateArray(it_stopwd);
	stopwd_map[2] = CreateArray(fr_stopwd);
	stopwd_map[3] = CreateArray(es_stopwd);
	stopwd_map[4] = CreateArray(en_stopwd);
	stopwd_map[6] = CreateArray(de_stopwd);
	stopwd_map[7] = CreateArray(se_stopwd);
	stopwd_map[9] = CreateArray(ro_stopwd);
	stopwd_map[130] = CreateArray(ne_stopwd);
*/

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

/*
bool TextUtilities::IsStopWord( const char *word, int language )
{
	std::map<int, std::vector<CString>*>::iterator i = stopwd_map.find(language);
	if( i == stopwd_map.end() ) {
//		TRACE("Unable to find language %d for stopword list!!!!!!!!\n", language);
		return false;
	}
	std::vector<CString>* mp = i->second;
#ifdef _DEBUG
	if( std::binary_search(mp->begin(), mp->end(), CString(word)) ) {
//		TRACE("%s IS STOPWORD for language %d\n", word, language);
		return true;
	}
	return false;
#else
	return std::binary_search(mp->begin(), mp->end(), CString(word));
#endif
}
*/

TextUtilities::~TextUtilities() {
//	for( std::map<int, std::vector<CString>*>::iterator i = stopwd_map.begin(); i != stopwd_map.end(); i ++) {
//		if( i->second )
//			delete i->second;
//	}
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

/*
#define MAX_ENTITY_SIZE 6

char TextUtilities::EntityToAscii( char * &p )
{
	char *orig = p;
	bool const is_num = (*p && *p == '#');
	bool is_hex = false;
	if( is_num ) {
		p++;
		if( *p && (*p == 'x' || *p == 'X') ) {
			is_hex = true;
			p++;
		}
	}
	char	entity_buf[ MAX_ENTITY_SIZE + 1 ];
	int		entity_len = 0;

	while( *p && *p != ';' ) {
		if( ++entity_len > MAX_ENTITY_SIZE ) {
			p = orig;
			return '&';		// Give up and return space
		}
		entity_buf[entity_len-1] = *p++;
	}
	if( ! *p ) {
		p = orig;
		return '&';	// didn't find it
	}
	p++;	// go past ;
	entity_buf[entity_len] = 0;
	if( ! is_num ) {
		unsigned n = 0x000000ff & (unsigned)entity_map[entity_buf];
		//printf("%s -> %u\n", entity_buf, n);
		return	n < sizeof( iso8859_map ) / sizeof( iso8859_map[0] ) ?
		iso8859_map[ n ] : ' ';
	}

	register unsigned n = 0;
	for ( register char const *e = entity_buf; *e; ++e ) {
		if ( is_hex ) {
			if ( !IsXDigit( *e ) ) {		// bad hex num
				return ' ';
			}
			n = (n << 4) | ( IsDigit( *e ) ?
				(*e - '0') : (ToLower( *e ) - 'a' + 10));
		} else {
			if ( !IsDigit( *e ) ) {		// bad dec num
				return ' ';
			}
			n = n * 10 + *e - '0';
		}
	}

	return	n < sizeof( iso8859_map ) / sizeof( iso8859_map[0] ) ?
		iso8859_map[ n ] : ' ';
}


void TextUtilities::increment_skip_p(char * &p, const int count, int &skipcount) {
	p += count;
	skipcount += count;
}
void TextUtilities::increment_buff(char * &b, const int count, int &skipcount, int &curpos, std::vector< std::pair<int, int> > &skip) {
	if( skipcount ) {
		std::pair<int, int> p(curpos, skipcount);
		//printf("%d - %d\n", curpos, skipcount);
		skipcount = 0;
		skip.push_back(p);
	}
	b	   += count;
	curpos += count;
}

int TextUtilities::LowerCompareNoCaseN(register const char *low, register const char *mix, register int count)
{
	int f, l;
	do {
		f = *low++;
		l = ToLower(*mix++);
	} while( --count && f && (f == l) );
	return (f-l);
}

void TextUtilities::StripHtmlWithSkip( char *src, std::vector< std::pair<int, int> > &skip)
{
	static char cc_script[10] = "</script";
	static char cc_style[10] = "</style";
	static char cc_head[10] = "</head";
	register char	*p = src;
	char *buff = p;
	char lastc = 0;

	int curpos = 0;
	int skipcount = 0;
	int rep = 0;
	while( *p ) {
		//printf("%c\n", *p);
		switch( *p ) {
		case	'<':
			increment_skip_p(p, 1, skipcount);
			//p++;
			if( *p == '!' && *(p+1) == '-' && *(p+2) == '-' ) {
				increment_skip_p(p, 3, skipcount);
				//p += 3;
				// Skip html comment
				while( *p ) {
					//printf(" %c\n", *p);
					if( *p != '-' ) {
						increment_skip_p(p, 1, skipcount);
						continue;
					}
					increment_skip_p(p, 1, skipcount);
					while( *p && *p == '-' )
						increment_skip_p(p, 1, skipcount);
						//p++;
					if( *p ) {
						if( *p == '>' ) {
							increment_skip_p(p, 1, skipcount);
							break;
						}
						increment_skip_p(p, 1, skipcount);
					}
					//if( *p && *p++ == '>' )
						//break;
				}

			} else if( ToLower(*p) == 's' ) {
				// strip <script>.*</script>
				// strip <style>.*</style>
				char *tag;int len,l1;
				if( LowerCompareNoCaseN("cript", p+1, 5) == 0) {
					tag = cc_script;
					len = 8,l1=5;
				} else if( LowerCompareNoCaseN("tyle", p+1, 4) == 0) {
					tag = cc_style;
					len = 7,l1=4;
				} else
					goto skiptag;
				int k = BMHSearch((unsigned char*)p+1, strlen(p+1), (unsigned char *)tag, len);
				if( k == -1 ) {
					*buff = 0;
					return;
				}
				//p += k + len + l1 +1;
				//while(*p && *p++ != '>');
				increment_skip_p(p, k + len + l1 +1, skipcount);
				while(*p && *p != '>')
					increment_skip_p(p, 1, skipcount);
				increment_skip_p(p, 1, skipcount);
			} else if( ToLower(*p) == 'h' && LowerCompareNoCaseN("ead", p+1, 3) == 0 ) {
				// strip <head>.*</head>
				char *tag = cc_head;
				int len = 6;
				int k = BMHSearch((unsigned char*)p+4, strlen(p+4), (unsigned char *)tag, len);
				if( k == -1 ) {
					*buff = 0;
					return;
				}
				//p += k + len + 4;
				//while(*p && *p++ != '>');
				increment_skip_p(p, k + len + 4, skipcount);
				while(*p && *p != '>')
					increment_skip_p(p, 1, skipcount);
				increment_skip_p(p, 1, skipcount);
			} else {
				if( ToLower(*p) == 'b' ||
					ToLower(*p) == 't' ||
					ToLower(*p) == 'd' ||
					ToLower(*p) == 'p' ||
					(*p == '/' && 
						(ToLower(*(p+1)) == 't' ||
						 ToLower(*(p+1)) == 'd' ||
						 ToLower(*(p+1)) == 'p')
					) ) {
					if( lastc != ' ' )
						rep = ' ';
				}
skiptag:
				// Skip html tag
				register char quote = 0;
				while( *p ) {
					if( *p == '>' ) {
						increment_skip_p(p, 1, skipcount);
						break;
					}
					if( quote ) {
						if( *p == quote )
							quote = 0;
						increment_skip_p(p, 1, skipcount);
						continue;
					}
					if( *p == '\"' || *p == '\'' ) {
						quote = *p;
						increment_skip_p(p, 1, skipcount);
						continue;
					}
					increment_skip_p(p, 1, skipcount);
				}
				if( rep == ' ' ) {
					lastc = *buff = ' ';
					--skipcount;
					increment_buff(buff, 1, skipcount, curpos, skip);
					rep = 0;
				}
			}

			break;
		case	'&':
			// translate character
			//p++;
			increment_skip_p(p, 1, skipcount);
			{
				char *p1 = p;
				char c = EntityToAscii( p );
				int l = p - p1;
				skipcount += l - 1;
				if( IsSpace(c) ) {
					if( lastc != ' ' ) {
						//lastc = *buff++ = ' ';
						lastc = *buff = ' ';
						increment_buff(buff, 1, skipcount, curpos, skip);
					} else
						skipcount++ ;
				} else {
					//lastc = *buff++ = c;
					lastc = *buff = c;
					increment_buff(buff, 1, skipcount, curpos, skip);
				}
			}
			break;
		default:
			if ( (unsigned char)*p == 146 ){
                *buff = '\'';
				increment_buff(buff, 1, skipcount, curpos, skip);
				p++;
            } else {
				if( IsSpace(*p) ) {
					if( lastc != ' ' ) {
						lastc = *buff = ' ';
						increment_buff(buff, 1, skipcount, curpos, skip);
						p++;
					} else
						increment_skip_p(p, 1, skipcount);
				} else {
					lastc = *buff = *p;
					increment_buff(buff, 1, skipcount, curpos, skip);
					p++;
				}
			}
			//p++;
			break;
		}
	}
	*buff = 0;
}

void TextUtilities::StripHtml( char *src )
{
	static char cc_noscript[15] = "</noscript";
	static char cc_script[10] = "</script";
	static char cc_style[10] = "</style";
	static char cc_head[10] = "</head";
	register char	*p = src;
	char *buff = p;
	char lastc = 0;

	while( *p ) {
		//printf("%c\n", *p);
		switch( *p ) {
		case	'<':
			p++;
			if( *p == '!' && *(p+1) == '-' && *(p+2) == '-' ) {
				p += 3;
				// Skip html comment
				while( *p ) {
					//printf(" %c\n", *p);
					if( *p++ != '-' )
						continue;
					while( *p && *p == '-' )
						p++;
					//while( *p && isspace(*p) )
						//p++;
					if( *p && *p++ == '>' )
						break;
				}
			} else if( ToLower(*p) == 's' ) {
				// strip <script>.*</script>
				// strip <style>.*</style>
				char *tag;int len,l1;
				if( LowerCompareNoCaseN("cript", p+1, 5) == 0) {
					tag = cc_script;
					len = 8,l1=5;
				} else if( LowerCompareNoCaseN("tyle", p+1, 4) == 0) {
					tag = cc_style;
					len = 7,l1=4;
				} else
					goto skiptag;
				int k = BMHSearch((unsigned char*)p+1, strlen(p+1), (unsigned char *)tag, len);
				if( k == -1 ) {
					*buff = 0;
					return;
				}
				p += k + len +1;
				while(*p && *p++ != '>');
			} else if( ToLower(*p) == 'h' && LowerCompareNoCaseN("ead", p+1, 3) == 0 ) {
				// strip <head>.*</head>
				char *tag = cc_head;
				int len = 6;
				int k = BMHSearch((unsigned char*)p+4, strlen(p+4), (unsigned char *)tag, len);
				if( k == -1 ) {
					*buff = 0;
					return;
				}
				p += k + len + 4;
				while(*p && *p++ != '>');
			} else if( ToLower(*p) == 'n' && LowerCompareNoCaseN("oscript", p+1, 7) == 0) {
				// strip <noscript>.*</noscript>
				char *tag = cc_noscript;
				int len = 10;
				int k = BMHSearch((unsigned char*)p+8, strlen(p+8), (unsigned char *)tag, len);
				if( k == -1 ) {
					*buff = 0;
					return;
				}
				p += k + len + 8;
				while(*p && *p++ != '>');
			} else {
				if( ToLower(*p) == 'b' ||
					ToLower(*p) == 't' ||
					ToLower(*p) == 'd' ||
					ToLower(*p) == 'p' ||
					(*p == '/' && 
						(ToLower(*(p+1)) == 't' ||
						 ToLower(*(p+1)) == 'd' ||
						 ToLower(*(p+1)) == 'p')
					) ) {
					if( lastc != ' ' ) {
						lastc = *buff = ' ';
						buff++;
					}
				}
skiptag:
				// Skip html tag
				register char quote = 0;
				while( *p ) {
					if( *p++ == '>' ) {
						break;
					}
					if( quote ) {
						if( *p++ == quote )
							quote = 0;
						continue;
					}
					if( *p == '\"' || *p == '\'' ) {
						quote = *p++;
						continue;
					}
				}
			}
			break;
		case	'&':
			// translate character
			p++;
			{
				char *p1 = p;
				char c = EntityToAscii( p );
//				int l = p - p1;
//				if( l > 0 )
//					p++;
				if( IsSpace(c) ) {
					if( lastc != ' ' ) {
						lastc = *buff = ' ';
						buff++;
					}
				} else {
					lastc = *buff = c;
					buff++;
				}
				if( buff > p )
					p = buff;
			}
			break;

		case	'\303':
			if( (unsigned char)*(p+1) >= 128 && (unsigned char)*(p+1) <= 191 ) {
				p++;
				*buff = iso8859_map[(unsigned char)*p+64];
				buff++;
				p++;
				break;
			}
		default:
			if ( (unsigned char)*p == 146 ){
                *buff = '\'';
				buff++;
            } else {
				if( IsSpace(*p) ) {
					if( lastc != ' ' ) {
						lastc = *buff = ' ';
						buff++;
					}
				} else {
					lastc = *buff = *p;
					buff++;
				}
			}
			p++;
			break;
		}
	}
	*buff = 0;
}


    // <B style="">
static const char *bstyle[] = {
    "color:black;background-color:#ffff66",
    "color:black;background-color:#A0FFFF",
    "color:black;background-color:#99ff99",
    "color:black;background-color:#ff9999",
    "color:black;background-color:#ff66ff",
    "color:white;background-color:#880000",
    "color:white;background-color:#00aa00",
    "color:white;background-color:#886800",
    "color:white;background-color:#004699",
    "color:white;background-color:#990099",
};

int TextUtilities::CalcPos(int w, std::vector< std::pair<int, int> > *skip, bool eq) {
    int res = w;
    if( eq ) {
        for( std::vector< std::pair<int, int> >::iterator i = skip->begin(); i != skip->end(); i ++ ) {
            if( i->first > w )
                break;
            res += i->second;
        }
    } else { 
        for( std::vector< std::pair<int, int> >::iterator i = skip->begin(); i != skip->end(); i ++ ) {
            if( i->first >= w )
                break;
            res += i->second;
        }
    }
    return res;
}

class PairLess {
public:
	bool operator()(const std::pair<int,CString>& x, const std::pair<int,CString>& y) const {
		return x.first > y.first;
	};
};


#include "regexp.h"

extern bool g_bText;

CString TextUtilities::HighlightHtml( char *html, char *wrds, int *nWord, LPCSTR brg)
{
	CString fc = html;
	std::vector< std::pair<int, int> > skip;
	StripHtmlWithSkip(fc.GetBuffer(fc.GetLength()), skip);
	fc.ReleaseBuffer();
	std::vector<CString> ws;
	std::vector<CString> wr;
	char *p = wrds;
    CString curWord;
    while( *p ) {
        if( *p == '+' ) {
			p++;
			continue;
		}
        if( IsSpace(*p) ) {
			g_tu.MakeLower(curWord);
			if( curWord.Find('@') == -1 && curWord.Find('*') == -1)
				ws.push_back(curWord);
			else
				wr.push_back(curWord);
            curWord = "";
            p++;
            continue;
        }
		curWord += *p;
		p++;
    }
	g_tu.MakeLower(curWord);
	if( curWord.Find('@') == -1 && curWord.Find('*') == -1)
		ws.push_back(curWord);
	else
		wr.push_back(curWord);

	
	int cmin = 0;
	int cmax = strlen(html);
	if( brg && *brg && ! g_bText ) {
		CRegexp re(brg, true);
		if( re.Match(html) ) {
			cmin = re.NameSubStart("articolo");
			cmax = cmin + re.NameSubLength("articolo");
		}
	}
	static const char stra[] = "<span id=PTManual style=\"font-weight:bold;%s\">";
    static const char strb[] = "</span>";
    std::vector<std::pair<int, CString> >    inserts;
	CString scol;
	{
		COLORREF	col = GetSysColor(COLOR_HIGHLIGHTTEXT);
		COLORREF	col1 = GetSysColor(COLOR_HIGHLIGHT);
		scol.Format("color: #%02X%02X%02X; background-color: #%02X%02X%02X",
			GetRValue(col), GetGValue(col), GetBValue(col),
			GetRValue(col1), GetGValue(col1), GetBValue(col1));
	}
    int j = 0;
    for( std::vector<CString>::iterator i = ws.begin(); i != ws.end(); i ++, j ++ ) {
        CString word = *i;
        bool bWord = (word.Find('*') == -1);
        if( ! bWord )
            word.Replace("*", "");
         word.Replace("+", "");
		if( word.GetLength() <= 1 )
			continue;
        int w = 0;
        char *p = (char *)(LPCSTR)fc;
        int l = fc.GetLength();
        int cl = 0;
        CString wstra;
		if( nWord == NULL )
			wstra.Format(stra, bstyle[j%10]);
		else
			wstra.Format(stra, scol);
        while( w < l ) {
            if( bWord )
                w = BMHSearchW( (unsigned char *)p, l, (unsigned char *)(LPCSTR)word, word.GetLength() );
            else
                w = BMHSearch( (unsigned char *)p, l, (unsigned char *)(LPCSTR)word, word.GetLength() );
            if( w == -1 ) {
                break;
            }
            int w1 = w + word.GetLength();
            int y1 = CalcPos(w + cl, &skip, 1);
            int y2 = CalcPos(w1 + cl, &skip, 0);
			if( y1 >= cmin && y1 <= cmax && y2 >= cmin && y2 <= cmax ) {
				std::pair<int, CString> pa;
				pa.first = y1;
				pa.second = wstra;
				inserts.push_back(pa);
				pa.first = y2;
				pa.second = strb;
				inserts.push_back(pa);
			}
            w = w + 1;
            cl += w;
            p += w;
            l -= w;
        }
    }
	for( i = wr.begin(); i != wr.end(); i ++, j ++ ) {
		CString word = *i;
		LPSTR	wz = word.GetBuffer(word.GetLength());
		LPSTR	we = wz + word.GetLength();
		CString sreg = "\\<";
		for( ; wz < we; wz++ ) {
			if( *wz == '*' ) {
				sreg += "[a-z0-9àáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ'\\-_&]*";
			} else if( *wz == '@' ) {
				sreg += "\\>";
				int na = wz[1] - '0';
				if( na > 1 && na <= 9 ) {
					wz++;
					while( na > 1 && na <= 9 ) {
						sreg += "[^a-z0-9àáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ]+\\w";
						na--;
					}
				}
				sreg += "[^a-z0-9àáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ]+\\<";
			} else if( *wz == '(') {
				sreg += "\\(";
			} else if( *wz == ')') {
				sreg += "\\)";
			} else {
				sreg += *wz;
			}
		}
		sreg += "\\>";
#ifdef _DBG_WPRDS
		TRACE("sreg= %s\n", sreg);
#endif
		int w = 0;
		char *p = (char *)(LPCSTR)fc;
		int l = fc.GetLength();
		int cl = 0;
		CString wstra;
		if( nWord == NULL )
			wstra.Format(stra, bstyle[j%10]);
		else
			wstra.Format(stra, scol);
		CRegexp re(sreg, TRUE);
		while( w < l ) {
			if( ! re.Match(p) )
				break;
			w = re.SubStart(0);
			int w1 = w + re.SubLength(0);
			int y1 = CalcPos(w + cl, &skip, 1);
			int y2 = CalcPos(w1 + cl, &skip, 0);
			if( y1 >= cmin && y1 <= cmax && y2 >= cmin && y2 <= cmax ) {
				std::pair<int, CString> pa;
				pa.first = y1;
				pa.second = wstra;
				inserts.push_back(pa);
				pa.first = y2;
				pa.second = strb;
				inserts.push_back(pa);
			}
			w = w + 1;
			cl += w;
			p += w;
			l -= w;
		}
    }
    CString res;
    fc = html;
	std::sort(inserts.begin(), inserts.end(), PairLess() );
	int k = inserts.size();
	if( nWord && *nWord*2 > k ) {
		MessageBeep(MB_OK);
		*nWord = 0;
		k = 0;
		return html;
	}
	for( std::vector<std::pair<int,CString> >::iterator ii = inserts.begin(); ii != inserts.end(); ii ++, k-- ) {
		if( nWord == NULL ) {
			res = ii->second + fc.Mid(ii->first) + res;
			fc = fc.Left(ii->first);
		} else if( nWord && *nWord*2 == k ) {
			res = "</a>" + ii->second +fc.Mid(ii->first) + res;
			fc = fc.Left(ii->first);
		} else if( nWord && *nWord*2-1 == k ) {
			res = "<a name='ptfind'>" + ii->second + fc.Mid(ii->first) + res;
			fc = fc.Left(ii->first);
		}
#ifdef _DBG_WPRDS
		else
			TRACE("Skipping %d\n", ii->first);
#endif
	}
	return fc + res;
}

bool TextUtilities::IsOkWord( const char *word, int language )
{
	int digits = 0;
	int puncts = 0;
	int uppers = 0;
	int vowels = 0;
	int alfa = 0;
	int	amp = 0;
	int	apo = 0;
	int at = 0;
	//int dash = 0;
	register const char *c = word;

	for( ; *c; c++ ) {
		if( IsDigit( *c ) ) {
			++ digits;
			continue;
		}
		if( IsPunct( *c ) ) {
			switch( *c ) {
			case '&':
				amp++;
				break;
			case '\'':
				apo++;
				break;
			//case '-':
				//dash++;
				//break;
			case '@':
				at++;
				break;
			}
			puncts++;
			continue;
		}
		if( IsAlfa( *c ) )
			++alfa;
		if( IsUpperCase( *c ) )
			++uppers;
		if( IsVowel( *c ) )
			++vowels;
	}
	int const len = c - word;
	// non esistono parole con più di un apostrofo, e commerciale o chiocciola
	if( amp > 1 || apo > 1 || at > 1 ) {
#ifdef DBG_WORDS
		printf("%s non esistono parole con più di un apostrofo, e commerciale o chiocciola\n", word);
#endif
		return false;
	}

	// Numbers
	if( len == puncts + digits && len > 0)
		return true;

	// se la lung. della parola senza punteggiature minore di WordMinSize
	if( len - puncts < WordMinSize ) {
#ifdef DBG_WORDS
		printf("%s se la lung. della parola senza punteggiature minore di WordMinSize\n", word);
#endif
		return false;
	}

	// Se la lung. della parola minore di WordMinSize
	if( len < WordMinSize ) {
#ifdef DBG_WORDS
		printf("%s Se la lung. della parola minore di WordMinSize\n", word);
#endif
		return false;
	}

	// se  non ci sono minuscole salvala
	if( uppers + puncts + digits == len )
		return true; // do not perform following checks

	//if( IsUpperCase( *word ) )
		//return true; // do not perform following checks

	//if( apo ) // words that start in lower case must not contain an apostrophe
		//return false;

	// italiano1, francese2, inglese4 e spagnolo3 non hanno parole con 0 vocali
	if( (language>0 && language < 5) && vowels == 0 && len > 4 ) {
#ifdef DBG_WORDS
		printf("%s italiano1, francese2, inglese4 e spagnolo3 non hanno parole con 0 vocali\n", word);
#endif
		return false;
	}
	
	// italiano1, francese2, inglese4 e spagnolo3 non hanno parole con tutte vocali da indicizzare
	if( (language>0 && language < 5) && vowels == len && len > 3 ) {
#ifdef DBG_WORDS
		printf("%s italiano1, francese2, inglese4 e spagnolo3 non hanno parole con tutte vocali da indicizzare\n", word);
#endif
		return false;
	}
	
	// words or addresses with more that 6 
	if( puncts > 5 ) {
#ifdef DBG_WORDS
		printf("%s words or addresses with more that 6\n", word);
#endif
		return false;
	}

	// non indicizzare parole che contengono tanti numeri e tante lettere
	if( digits > 6 && alfa > 6 ) {
#ifdef DBG_WORDS
		printf("%s non indicizzare parole che contengono tanti numeri e tante lettere\n", word);
#endif
		return false;
	}

	printf("IsUpperCase(%c) == %s\n", *word, IsUpperCase(*word) ? "TRUE" : "FALSE");
	if( IsUpperCase(*word) ) {
#ifdef DBG_WORDS
		printf("%s Skip consecutive-character checks because first letter is uppercase\n", word);
#endif
		return true;
	}

	////////// Perform consecutive-character checks ///////////////////////
	int consec_consonants = 0;
	int consec_vowels = 0;
	int consec_same = 0;
	int consec_puncts = 0;
	
	register char last_c = '\0';

	for( c = word; *c; ++c ) {

		if ( IsDigit( *c ) ) {
			consec_consonants = 0;
			consec_vowels = 0;
			consec_puncts = 0;
			last_c = '\0';	// consec_same doesn't apply to digits
			continue;
		}

		if ( IsPunct( *c ) ) {
			if ( ++consec_puncts > WordMaxConsecPuncts ) {
#ifdef _DBG_WPRDS
				char buff[256];
				sprintf(buff,  "SKIPPED ( ++consec_puncts(%d) > WordMaxConsecPuncts(%d) )", consec_puncts,WordMaxConsecPuncts);
				TRACE("%s %s\n", word, buff);
#endif
				return false;
			}
			consec_consonants = 0;
			consec_vowels = 0;
			continue;
		}

		if ( *c == last_c ) {
			if ( ++consec_same > WordMaxConsecSame ) {
#ifdef _DBG_WPRDS
				char buff[256];
				sprintf(buff,  "SKIPPED ( ++consec_same(%d) > WordMaxConsecSame(%d) )", consec_same,WordMaxConsecSame);
				TRACE("%s %s\n", word, buff);
#endif
				return false;
			}
		} else {
			consec_same = 0;
			last_c = *c;
		}

		if ( IsVowel( *c ) ) {
			if ( ++consec_vowels > WordMaxConsecVowels ) {
#ifdef _DBG_WPRDS
				char buff[256];
				sprintf(buff,  "SKIPPED ( ++consec_vowels(%d) > WordMaxConsecVowels(%d) )", consec_vowels,WordMaxConsecVowels);
				TRACE("%s %s\n", word, buff);
#endif
				return false;
			}
			consec_consonants = 0;
			consec_puncts = 0;
			continue;
		}

		if ( ++consec_consonants > WordMaxConsecConsonants ) {
#ifdef _DBG_WPRDS
			char buff[256];
			sprintf(buff,  "SKIPPED ( ++consec_consonants(%d) > WordMaxConsecConsonants(%d) )", consec_consonants,WordMaxConsecConsonants);
			TRACE("%s %s\n", word, buff);
#endif
			return false;
		}
		consec_vowels = 0;
		consec_puncts = 0;
	}

	return true;
}

//#include <algorithm>
#ifdef _WIN32
#ifdef _WRITE_WORDS
extern std::map<CString, int>	g_found_words;
#endif
void TextUtilities::GetWords( int language, CMap<CString, LPCSTR, int, int> &words, const char *text )
#else
void TextUtilities::GetWords( int language, std::map<CString, int> &words, const char *text )
#endif
{
	int		wordCount = 0;
	//const char *original_text = text;
	char *current_word = new char[WordMaxSize + 1];
	char *original_word = current_word;
	CString last_word, last_last;
	bool bAdded = false, lastBadded = false;;
	int tl = 0;

	while( *text ) {
		tl++;
		// Ensure word starts correctly
		if( ! IsWordBegin( *text ) ) {
			if( *text == '-' || *text == '+') {
				if( !IsDigit(*(text+1)) ) {
					text++;
					continue;
				}
			} else {
#ifdef _DBG_WPRDS
				TRACE("Skipping because text starts with nowWordBegin %c\n", *text);
#endif
				text++;
				continue;
			}
		}
		*current_word++ = *text++;
		// assign at most WordMaxSize characters to word
		bool bEatText = true;
		char *wMax = WordMaxSize + original_word;
		while( IsWord(*text) ) {
			if( current_word < wMax ) {
				*current_word++ = *text;
				if( language == 1  && *text == '\'' && tl > 4 ) {
					if( IsSpace(*(text+1)) && *(text-1) == 'l' && *(text-2) == 'l' && (IsVowel(*(text+2)) || IsDigit(*(text+2))) )
						text++;
				}
				text++;
			} else {
#ifdef _DBG_WPRDS
				TRACE("word too big!\n");
#endif
				text++;
				bEatText = false;
			}
		}
		// REefuse characters that don't end the word correctly
		*current_word-- = 0;
		while( current_word >= original_word ) {
			if( IsWordEnd( *current_word ) )
				break;
			*current_word-- = 0;
			if( bEatText ) {
				text--;
				tl--;
			}
		}
		current_word = original_word;
		wordCount++;
		// Skip words that are too small
		if( current_word[0] == 0 || current_word[1] == 0 || bEatText == false ) {
#ifdef _DBG_WPRDS
			TRACE("%s Skipping because word is too short1\n", current_word);
#endif
			continue;
		}
#ifdef _DBG_WPRDS
		TRACE("Processing ----> %s\n", current_word);
#endif
		{
			std::vector<CString> pw;
			std::vector<int> ix;
			pw.push_back(CString(current_word));
			CString	word(current_word);
			char *tp = current_word;
			last_last = last_word;
			last_word = current_word;
			last_last.MakeLower();
			lastBadded = bAdded;
			bAdded = false;
			// find all punctuation characters
			while( (tp = strpbrk(tp, "'`´.-&")) != NULL ) {
				//if( language == 1 && current_word[tp] == '\''
				ix.push_back(tp-current_word);
				tp++;
			}
			for( unsigned i = 0; i < ix.size(); i ++ ) {
				if(std::find(pw.begin(),pw.end(),word.Left(ix[i])) == pw.end()) {
#ifdef _DBG_WPRDS
					TRACE("lang=%d, i=%d, pushing back %s%c\n", language, i, word.Left(ix[i]),current_word[ix[i]]);
#endif
					if( language == 1 && i == 0 ) {
						CString w;
						w = word.Left(ix[i]);
						w.MakeLower();
						if( w.Right(2) == "ll" ) {
					   		if( ! IsApo(current_word[ix[i]]) )
								pw.push_back(word.Left(ix[i]));
							else {
#ifdef _DBG_WPRDS
								TRACE("skipping %s because it is followed by '\n", w);
#endif
							}
						} else
							pw.push_back(word.Left(ix[i]));
					} else
						pw.push_back(word.Left(ix[i]));
				}
				if(std::find(pw.begin(),pw.end(),word.Right(word.GetLength()-ix[i]-1)) == pw.end()) {
					pw.push_back(word.Right(word.GetLength()-ix[i]-1));
				}
				for( unsigned j = i+1; j < ix.size(); j ++ ) {
					if( j < ix.size()-1 &&
						std::find(pw.begin(), pw.end(), word.Mid(ix[j]+1,ix[j+1]-ix[j])) == pw.end())
						pw.push_back(word.Mid(ix[j], ix[j+1] - ix[j]));
					if( std::find(pw.begin(), pw.end(), word.Mid(ix[i]+1,ix[j]-ix[i])) == pw.end() )
						pw.push_back(word.Mid(ix[i], ix[j] - ix[i]));
				}
			}
			ix.erase(ix.begin(), ix.end());
			for( i = 0; i < pw.size(); i ++ ) {
				CString    word = pw[i];
				CString	prevword;
				CString	upperWord = word;

				if( i > 0 )
					prevword = pw[i-1];
				if( word.GetLength() < 2 || ! IsWordBegin(word[0])) {
					//if( ! (word.GetLength() == 1 && IsDigit(word[0])) ) {
#ifdef _DBG_WPRDS
						TRACE( "%s length=%d IsWordBegin(%c)=%d SKIPPED (length<2 or !is_word_begin_char)\n", word, word.GetLength(), word[0], IsWordBegin(word[0]) );
#endif
						continue;
					//}
				}
				if( ! IsOkWord( word, language ) ) {
#ifdef _DBG_WPRDS
					TRACE( "%s SKIPPED (!is_ok_word)\n", word );
#endif
					continue;
				}
				bool bAllCaps = IsAllCaps(word);
				bool bIsDell = false;
				MakeLower(prevword);
				MakeLower(word);
				bAllCaps = false;
				if( word.Compare("dell") == 0 && !prevword.IsEmpty() && (bAllCaps || prevword[0] == 'D') )
					bIsDell = true;
				if( bAllCaps || bIsDell || ! IsStopWord( word, language ) ) {
					// Per bypassare le stopword per le parole che precedono es. di pietro
					if( word.GetLength() > 0 && IsDigit(word[0]) && 
									( ( word.GetLength() == 2 && IsDigit(word[1]) ) || word.GetLength() == 1 ) ) {
							if( ! IsValidNumber(word, last_last) ) {
#ifdef _DBG_WPRDS
								TRACE("%s SKIPPING : Is not a valid number!\n", word);
#endif
								continue;
							} else {
#ifdef _DBG_WPRDS
								TRACE("%s ADDING : Is a valid number!\n", word);
#endif
							}
					} else {
#ifdef _DBG_WPRDS
						TRACE("%s ADDED (!stopword)\n", word);
#endif
					}
// Da aggiungere quando serve bypassare le stopword per le parole che seguono es. carta si
				} else {
#ifdef _DBG_WPRDS
					TRACE("%s SKIPPED (stopword)\n", word);
#endif
					continue;
				}
#ifdef _WIN32
				int val = 0;
				if( words.Lookup(word, val) )
					words[word] = words[word] + 1;
				else
					words[word] = 1;

#ifdef _WRITE_WORDS
				std::map<CString, int>::iterator iv = g_found_words.find(word);
				if( iv != g_found_words.end() )
					iv->second = iv->second + 1;
				else
					g_found_words.insert(std::pair<CString, int> (word, 1));
#endif
#else
				std::map<CString, int>::iterator iv = words.find(word);
				if( iv != words.end() )
					*iv = *iv + 1;
				else
					words.insert(std:pair<CString, int> (word, 1));
#endif
				bAdded = true;
			}
			pw.erase(pw.begin(), pw.end());
		}
	}
	delete [] original_word;
	return;
}
*/