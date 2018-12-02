
#ifndef _TEXT_UTILITIES_H_

#define _TEXT_UTILITIES_H_

#ifdef _MSC_VER
#include <afxtempl.h>
#pragma warning (disable : 4018)
#else#include "CString.h"
#endif


#ifndef _VECTOR_
#include <vector>
#endif

class TextUtilities
{
	void MakeLookupTable(const char *s, bool *l);
	void MakeConversionTable( const char *from, const char *to, char *dest );
	CMap<const char *, const char *, char, char> entity_map;
//	void increment_skip_p(char * &p, const int count, int &skipcount);
//	void increment_buff(char * &b, const int count, int &skipcount, int &curpos, std::vector< std::pair<int, int> > &skip);
public:
	TextUtilities() { Init(); };
	~TextUtilities();
	bool Init();

	bool IsSpace(const char c);
	bool IsAlfa(const char c);
	bool IsWord(const char c);
	bool IsWordBegin(const char c);
	bool IsWordEnd(const char c);
	bool IsLowerCase(const char c);
	bool IsUpperCase(const char c);
	bool IsVowel(const char c);
	bool IsDigit( const char c );
	bool IsXDigit( const char c );
	bool IsAlfaNum( char c );
	bool IsApo( char c );
	bool IsPunct( const char c );
	char ToLower( char c );
	char ToUpper( char c );
	void MakeLower( register char *s );
	void MakeLower( CString &s );
	void MakeUpper( register char *s );
	void MakeUpper( CString &s );
	void CapitalizeWords( register char *s, int bLower = true );
	bool IsAllCaps( register const char *s );
	int BMHSearch( unsigned char *text, int n, unsigned char *pat, int m, BOOL bCaseSens = FALSE );
	int BMHSearchW( unsigned char *text, int n, unsigned char *pat, int m, BOOL bCaseSens = FALSE );
	void hexdump(char *title, char *area, int length);
/*
	char EntityToAscii( char * &p );
	int LowerCompareNoCaseN(register const char *low, register const char *mix, register int count);

	void StripHtmlWithSkip( char *src, std::vector< std::pair<int, int> > &skip);
	void StripHtml( char *src );
	int CalcPos(int w, std::vector< std::pair<int, int> > *skip, bool eq);
	CString HighlightHtml( char *html, char *words, int *nWord = NULL, LPCSTR brg = NULL);
	bool IsOkWord( const char *word, int language );
	bool IsStopWord( const char *word, int language );
#ifdef _WIN32
	void GetWords( int language, CMap<CString, LPCSTR, int, int> &words, const char *text );
#else
	void GetWords( int language, std::map<CString, int> &words, const char *text );
#endif
*/

	static bool TestAll();
};

extern TextUtilities g_tu;

#endif