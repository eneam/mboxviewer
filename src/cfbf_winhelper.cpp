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

/*
* Cfbfinfo is written by Graeme Cole
* https://github.com/elocemearg/cfbfinfo
* 
* Helper file by Zbigniew Minciel to port code to Windows 
* 
* Note: For simplification, the code assumes that the target system is little-endian. 
*/

#ifdef _WINDOWS

// Fix vars definitions
#pragma warning(disable : 4267)
#pragma warning(disable : 4244)

#include "ReportMemoryLeaks.h"


#include "stdafx.h"
#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <errno.h>
#include <stdarg.h>
#include "wchar.h"

#include "cfbf_winhelper.h"

void error(int status, int errnum, const char* format, ...) {
    fprintf(stdout, "\nstatus=%d errnum=%d ", status, errnum);

    va_list args;  // Declare a variable argument list
    va_start(args, format);  // Initialize the argument list
    _vprintf_p(format, args);  // Print formatted output
    va_end(args);  // Clean up the argument list

    exit(0);
}


// Caller must call free(utf16buffer) UTF8ToUTF16 utf16buffer 
wchar_t* UTF8ToUTF16(char* in, size_t in_len, size_t* out_len)
{
    *out_len = 0;
    int out_buff_len = in_len * 4;
    wchar_t* out_buff = (wchar_t*)cfbf_malloc(out_buff_len);
    if (out_buff == NULL)
        return 0;
    unsigned int inCodePage = CP_UTF8;
    DWORD err = 0;

    DWORD dwFlags = 0;
    int retlen = CodePage2WStr(in, in_len, inCodePage, out_buff, out_buff_len, &err, dwFlags);
    if (retlen >= 0)
    {
        *out_len = (size_t)retlen;
        return out_buff;
    }
    free(out_buff);
    return 0;
}

char* UTF16ToUTF8(wchar_t* in, size_t in_len, size_t* out_len)
{
    *out_len = 0;
    int out_buff_len = in_len * 4;
    char* out_buff = (char*)cfbf_malloc(out_buff_len);
    if (out_buff == NULL)
        return 0;
    unsigned int outCodePage = CP_UTF8;
    DWORD err = 0;
    int retlen = WStr2CodePage(in, in_len, outCodePage, out_buff, out_buff_len, &err);
    if (retlen >= 0)
    {
        *out_len = (size_t)retlen;
        return out_buff;
    }
    free(out_buff);
    return 0;
}

std::string UTF16ToUTF8(wchar_t* in, size_t in_len)
{
    size_t out_len = 0;
    char* buff8 = UTF16ToUTF8(in, in_len, &out_len);
    if (buff8)
    {
        std::string utf8(buff8);
        free(buff8);
        return utf8;
    }
    else
    {
        std::string utf8;
        return utf8;
    }
}

// Convert dat in UTF16 to specif page code such as UTF8
int WStr2CodePage(const wchar_t* wbuff, int wbufflen, unsigned int outCodePage, char* astr, int astrlen, DWORD* error)
{
    *error = 0;
    _ASSERTE(wbufflen > 0);
    _ASSERTE(astrlen > 0);
    if (wbufflen == 0) {
        return -1;
    }

    unsigned int CP_ASCII = 20127;

    // worst case scenario. Estimated length
    int sbufflen = wbufflen * 4 + 2;  // This likely overkill or call WideCharToMultiByte first to get the required length
    sbufflen = ((sbufflen / 128) + 1) * 128; 

    if (outCodePage == CP_ASCII) {
        outCodePage = CP_UTF8;
    }

    char* sbuff = (char*)cfbf_malloc(sbufflen);
    if (sbuff == NULL)
        return -1;

    int slen = WideCharToMultiByte(outCodePage, 0, wbuff, wbufflen, sbuff, sbufflen - 1, NULL, NULL);
    _ASSERTE(slen != 0);
    if (slen == 0) {
        *error = GetLastError();

        free(sbuff);
        return -1;
    }

    // NULL terminate astr output buffer
    if (sbuff[slen - 1] == 0) { // already terminated by NULL
        slen = strlen(sbuff);
    }

	if ((slen + 1) > astrlen) {
		free(sbuff);
		return -1;
	}

	memcpy(astr, sbuff, slen);
	astr[slen] = 0;

    free(sbuff);
    return slen;
}

// Convert data in specif page code such as UTF8 to UTF16
int CodePage2WStr(const char* str, int slen, UINT strCodePage, wchar_t* wstr, int wstrlen, DWORD* error, DWORD dwFlags)
{
    _ASSERTE(slen > 0);
    _ASSERTE(wstrlen > 0);

    error = 0;
    wstr[0] = 0;
    if (slen == 0)
        return 0;

    // worst case scenario. Estimated length
    int wbuffLen = slen * 4 + 2;  // This is likely overkill or call MultiByteToWideChar first to get the required length
    wbuffLen = ((wbuffLen / 128) + 1) * 128; 

    wchar_t* wbuff = (wchar_t*)cfbf_malloc(wbuffLen*2 + 1);
    if (wbuff == NULL)
        return -1;
    int wlen = MultiByteToWideChar(strCodePage, 0, str, slen, wbuff, wbuffLen);
    _ASSERTE(wlen != 0);
    if (wlen == 0) {
        *error = GetLastError();

        free(wbuff);
        return -1;
    }

    // NULL terminate wstr output buffer
    if (wbuff[wlen - 1] == 0) {  // already terminated by NULL
        wlen = wcslen(wbuff);
    }

	if ((wlen + 1) > wstrlen) {
		free(wbuff);
		return -1;
	}

	wmemcpy(wstr, wbuff, wlen);
	wstr[wlen] = 0;

    free(wbuff);
    return wlen;
}

#ifdef OUTLOOK_MODE
int MyMessageBox(HWND h, LPCTSTR lpszText, LPCTSTR lpszCaption, UINT nType) { return 1; }
int MyMessageBox(LPCTSTR lpszText, LPCTSTR lpszCaption, UINT nType) { return 1; }
int MyAfxMessageBox(LPCTSTR lpszText, UINT nType) { return 1; }
#endif

#endif
