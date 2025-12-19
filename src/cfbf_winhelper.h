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

#ifndef _CFBF_WINHELPER_H
#define _CFBF_WINHELPER_H

#ifdef _WINDOWS

#include "Windows.h"
#include <stdint.h>

#include <errno.h>
#include <sys/types.h>
#include <sys/stat.h>
#include <fcntl.h>
#include "io.h"

#include <stdlib.h>
#include <stdio.h>
#include <string.h>

#include <stdarg.h>
#include "wchar.h"

#include "property.h"


//typedef uint32_t ULONG;    // 4 Bytes
//typedef uint16_t USHORT;  // 2 Bytes
typedef int64_t OFFSET;           // 2 Bytes
typedef ULONG SECT;             // 4 Bytes
typedef ULONG FSINDEX;          // 4 Bytes
typedef USHORT FSOFFSET;        // 2 Bytes
//typedef USHORT WCHAR;           // 2 Bytes
typedef ULONG DFSIGNATURE;      // 4 Bytes
//typedef uint8_t BYTE;     // 1 Byte
//typedef uint16_t WORD;    // 2 Bytes
//typedef uint32_t DWORD;    // 4 Bytes
//typedef ULONG SID;              // 4 Bytes
//typedef unsigned char CLSID[16];// 16 Bytes

void error(int status, int errnum, const char* format, ...);

#define UTF16LE_TO_UTF8 1
#define UTF16_TO_UTF8 1  // investigate
#define UTF8_TO_UTF16LE 2
#define UTF8_TO_UTF16 2 // investigate

typedef int iconv_t;


int WStr2CodePage(const wchar_t* wbuff, int wbufflen, unsigned int outCodePage, char* astr, int astrlen, DWORD* error);
int CodePage2WStr(const char* str, int slen, UINT strCodePage, wchar_t* wstr, int wstrlen, DWORD* error, DWORD dwFlags);

// Caller must call free(utf16buffer) UTF8ToUTF16 utf16buffer 
wchar_t* UTF8ToUTF16(char* in, size_t in_len, size_t* out_len);
std::string UTF16ToUTF8(wchar_t* in, size_t in_len);

// Caller must call free(utf8buffer) UTF16ToUTF8 utf8buffer 
char* UTF16ToUTF8(wchar_t* in, size_t in_len, size_t* out_len);

#endif
#endif
