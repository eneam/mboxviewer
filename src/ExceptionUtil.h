//
//////////////////////////////////////////////////////////////////
//
//  Windows Mbox Viewer is a free tool to view, search and print mbox mail archives.
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


// USE_STACK_WALKER is defined under project property
//

#include <stdexcept>
#include "stdafx.h"
#ifdef USE_STACK_WALKER
#include "StackWalker.h"
#endif
#include "eh.h"


class SimpleString;

class SE_Exception : public std::exception
{
private:
	const unsigned int nSE;
public:
	SE_Exception() noexcept : SE_Exception{ 0 } {}
	SE_Exception(unsigned int n) noexcept : nSE{ n } {}
	unsigned int getSeNumber() const noexcept { return nSE; }
};

class Scoped_SE_Translator
{
private:
	const _se_translator_function old_SE_translator;
public:
	Scoped_SE_Translator(_se_translator_function new_SE_translator) noexcept
		: old_SE_translator{ _set_se_translator(new_SE_translator) } {}
	~Scoped_SE_Translator() noexcept { _set_se_translator(old_SE_translator); }
};

void __cdecl trans_func(unsigned int u, EXCEPTION_POINTERS*);


// Conditional compilation; not sure if it is supported on all OS platforms TBD
#ifdef USE_STACK_WALKER
class MyStackWalker : public StackWalker
{
public:
	MyStackWalker(int  options = OptionsAll, // 'int' is by design, to combine the enum-flags
		LPCSTR szSymPath = NULL,
		DWORD  dwProcessId = GetCurrentProcessId(),
		HANDLE hProcess = GetCurrentProcess());

	static SimpleString *GetBuffer();


protected:
	virtual void OnOutput(LPCSTR szText);

	// No Mutexes ? Why not class member
	static SimpleString *stackWalkerBuffer;
};

#endif

BOOL DumpStack(char *fileName, const char *seText, UINT seNumb, PCONTEXT ContextRecord = 0, int mailPosition = -1);
BOOL DumpStackEx(char *fileName, CException* e);
BOOL DumpMailData(char *fileName, const char *seText, UINT seNumb, int mailPosition, char *data, int datalen);

BOOL GetProgramDir(CString &progDir);
BOOL CreateDumpFilePath(char *fileName, CString &filePath);
const char* seDescription(const int code);



