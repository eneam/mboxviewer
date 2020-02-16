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

#include "stdafx.h"
#include "eh.h"
#include "SimpleString.h"
#include <dbghelp.h>
#include "StackWalker.h"
#include "ExceptionUtil.h"

BOOL CPathGetPath(const char *path, CString &filePath);


void __cdecl trans_func(unsigned int u, EXCEPTION_POINTERS* ep)
{
	char const*  szCause = seDescription(u);
	BOOL ret = DumpStack("TranslationSystemE_StackTrace.txt", (TCHAR*)szCause, u, ep->ContextRecord);

	throw SE_Exception(u);
}

#ifdef USE_STACK_WALKER

SimpleString *MyStackWalker::stackWalkerBuffer = 0;


MyStackWalker::MyStackWalker(int options, LPCSTR szSymPath, DWORD  dwProcessId, HANDLE hProcess) :
	StackWalker(options, szSymPath, dwProcessId, hProcess)
{
	if (stackWalkerBuffer == 0)
		stackWalkerBuffer = new SimpleString;

	if (stackWalkerBuffer)
		MyStackWalker::stackWalkerBuffer->Clear();
};

SimpleString *MyStackWalker::GetBuffer()
{
	if (stackWalkerBuffer == 0)
		stackWalkerBuffer = new SimpleString;

	return stackWalkerBuffer;
}

void MyStackWalker::OnOutput(LPCSTR szText)
{
	if (stackWalkerBuffer == 0)
		stackWalkerBuffer = new SimpleString;

	if (stackWalkerBuffer)
		stackWalkerBuffer->Append((char*)szText);
};

#endif

BOOL GetProgramDir(CString &progDir)
{
	progDir.Empty();

	char *pProgPath = 0;
	errno_t  er = _get_pgmptr(&pProgPath);
	if (pProgPath && (er == 0))
	{
		BOOL ret = CPathGetPath(pProgPath, progDir);
		return ret;
	}
	else
		return FALSE;
}

BOOL CreateDumpFilePath(char *fileName, CString &filePath)
{
	filePath.Empty();

	char *pProgPath = 0;
	errno_t  er = _get_pgmptr(&pProgPath);
	if (pProgPath && (er == 0))
	{
		CString fileDir;
		//BOOL ret = CPathGetPath(pProgPath, fileDir);

		char *noExtension_ProgPath = new char[strlen(pProgPath) + 1];
		strcpy(noExtension_ProgPath, pProgPath);
		PathRemoveExtension(noExtension_ProgPath);
		filePath.Append(noExtension_ProgPath);
		delete noExtension_ProgPath;

		filePath.Append("_");
		filePath.Append(fileName);
	}
	if (filePath.IsEmpty())
	{
		// Should never be here but -:)
		filePath.Append("C:\\");
		filePath.Append(fileName);
	}
	return TRUE;
}

#ifdef USE_STACK_WALKER

BOOL DumpStack(char *fileName, const char *seText, UINT seNumb, PCONTEXT ContextRecord, int mailPosition)
{
	BOOL ret = TRUE;

	int optionsAll = 
		  StackWalker::RetrieveSymbol      // Try to get the symbol-name
		| StackWalker::RetrieveLine        // Try to get the line for this symbol
		| StackWalker::RetrieveModuleInfo  // Try to retrieve the module-infos
		| StackWalker::RetrieveFileVersion // Also retrieve the version for the DLL/EXE
		| StackWalker::SymBuildPath        // Generate a "good" symbol-search-path
		| StackWalker::SymUseSymSrv        // Also use the public Microsoft-Symbol-Server
		; 

	int options =
		StackWalker::RetrieveSymbol      // Try to get the symbol-name
		| StackWalker::RetrieveLine        // Try to get the line for this symbol
		//| StackWalker::RetrieveModuleInfo  // Try to retrieve the module-infos
		//| StackWalker::RetrieveFileVersion // Also retrieve the version for the DLL/EXE
		| StackWalker::SymBuildPath        // Generate a "good" symbol-search-path
		//| StackWalker::SymUseSymSrv        // Also use the public Microsoft-Symbol-Server
		;

	MyStackWalker sw(options);
	//MyStackWalker sw(StackWalker::OptionsAll, NULL, GetCurrentProcessId(), GetCurrentProcess());

	//sw.ShowCallstack();
	sw.ShowCallstack(GetCurrentThread(), ContextRecord);

	SimpleString *buffer = MyStackWalker::GetBuffer();
	if (buffer)
	{
		CString filePath;
		BOOL retC = CreateDumpFilePath(fileName, filePath);

		CFile fp;
		if (!fp.Open(filePath, CFile::modeWrite | CFile::modeCreate))
		{
#if 0
			CString txt = _T("Could not create Stack Dump File\"") + procFullPath;
			txt += _T("\" file.\n");
			HWND h = NULL; // we don't have any window yet
			int answer = ::MessageBox(h, txt, _T("Error"), MB_APPLMODAL | MB_ICONERROR | MB_OK);
#endif
			ret = FALSE;
		}
		else
		{
			CString tmp = CTime::GetCurrentTime().Format(_T("%A, %B %d, %Y %H:%M:%S\n"));
			fp.Write(tmp, tmp.GetLength());

			tmp.Format(_T("Exception: Code %8.8x Description: "), seNumb);
			fp.Write(tmp, tmp.GetLength());

			fp.Write(seText, strlen(seText));

			tmp.Format(_T("\nMail Position: %d\n"), mailPosition);
			fp.Write(tmp, tmp.GetLength());

			fp.Write("\n\n", 2);

			fp.Write(buffer->Data(), buffer->Count());
			fp.Close();
			ret = TRUE;
		}
		buffer->Clear();
	}
	else
		ret = FALSE;

	return ret;
}


BOOL DumpStackEx(char *fileName, CException* e)
{
	TCHAR szCause[256];
	UINT seNumber = 0;
	e->GetErrorMessage(szCause, 255, &seNumber);

	BOOL ret = DumpStack(fileName, szCause, seNumber);
	return ret;
}
#else

BOOL DumpStack(char *fileName, const char *seText, UINT seNumb, PCONTEXT ContextRecord, int mailPosition)
{
	return TRUE;
}

BOOL DumpStackEx(char *fileName, CException* e)
{
	return TRUE;
}

#endif

const char* seDescription(const int code)
{
	switch (code) {
	case EXCEPTION_ACCESS_VIOLATION:         return "ACCESS VIOLATION";
	case EXCEPTION_ARRAY_BOUNDS_EXCEEDED:    return "ARRAY BOUNDS EXCEEDED";
	case EXCEPTION_BREAKPOINT:               return "BREAKPOINT";
	case EXCEPTION_DATATYPE_MISALIGNMENT:    return "DATATYPE MISALIGNMENT";
	case EXCEPTION_FLT_DENORMAL_OPERAND:     return "FLT DENORMAL OPERAND";
	case EXCEPTION_FLT_DIVIDE_BY_ZERO:       return "FLT DIVIDE BY ZERO";
	case EXCEPTION_FLT_INEXACT_RESULT:       return "FLT INEXACT RESULT";
	case EXCEPTION_FLT_INVALID_OPERATION:    return "FLT INVALID OPERATION";
	case EXCEPTION_FLT_OVERFLOW:             return "FLT OVERFLOW";
	case EXCEPTION_FLT_STACK_CHECK:          return "FLT STACK CHECK";
	case EXCEPTION_FLT_UNDERFLOW:            return "FLT UNDERFLOW";
	case EXCEPTION_ILLEGAL_INSTRUCTION:      return "ILLEGAL INSTRUCTION";
	case EXCEPTION_IN_PAGE_ERROR:            return "IN PAGE ERROR";
	case EXCEPTION_INT_DIVIDE_BY_ZERO:       return "INT DIVIDE BY ZERO";
	case EXCEPTION_INT_OVERFLOW:             return "INT OVERFLOW";
	case EXCEPTION_INVALID_DISPOSITION:      return "INVALID DISPOSITION";
	case EXCEPTION_NONCONTINUABLE_EXCEPTION: return "NONCONTINUABLE EXCEPTION";
	case EXCEPTION_PRIV_INSTRUCTION:         return "PRIV INSTRUCTION";
	case EXCEPTION_SINGLE_STEP:              return "SINGLE STEP";
	case EXCEPTION_STACK_OVERFLOW:           return "STACK OVERFLOW";
	default: return "UNKNOWN";
	}
}

BOOL DumpMailData(char *fileName, const char *seText, UINT seNumb, int mailPosition, char *data, int datalen)
{
	BOOL ret = TRUE;

	CString filePath;
	BOOL retC = CreateDumpFilePath(fileName, filePath);

	CFile fp;
	if (!fp.Open(filePath, CFile::modeWrite | CFile::modeCreate))
	{
		ret = FALSE;

	}
	else
	{
		CString tmp = CTime::GetCurrentTime().Format(_T("%A, %B %d, %Y %H:%M:%S\n"));
		fp.Write(tmp, tmp.GetLength());

		tmp.Format(_T("Exception: Code %8.8x Description: "), seNumb);
		fp.Write(tmp, tmp.GetLength());

		fp.Write(seText, strlen(seText));

		tmp.Format(_T("\nMail Position: %d\n"), mailPosition);
		fp.Write(tmp, tmp.GetLength());

		fp.Write("\n\n", 2);

		fp.Write(data, datalen);
		fp.Write("\n", 1);

		fp.Close();
		ret = TRUE;
	}
	return ret;
}

// Need to review and leverage to set or not set SetUnhandledExceptionFilter() 
#ifdef UNHANDLED_EXCEPTION_TEST

/**********************************************************************
 *main.cpp from
 * https://github.com/JochenKalmbach/StackWalker
 **********************************************************************/

// For more info about "PreventSetUnhandledExceptionFilter" see:
// "SetUnhandledExceptionFilter" and VC8
// http://blog.kalmbachnet.de/?postid=75
// and
// Unhandled exceptions in VC8 and above… for x86 and x64
// http://blog.kalmbach-software.de/2008/04/02/unhandled-exceptions-in-vc8-and-above-for-x86-and-x64/
// Even better: http://blog.kalmbach-software.de/2013/05/23/improvedpreventsetunhandledexceptionfilter/

#if defined(_M_X64) || defined(_M_IX86)
static BOOL PreventSetUnhandledExceptionFilter()
{
	HMODULE hKernel32 = LoadLibrary(_T("kernel32.dll"));
	if (hKernel32 == NULL)
		return FALSE;
	void* pOrgEntry = GetProcAddress(hKernel32, "SetUnhandledExceptionFilter");
	if (pOrgEntry == NULL)
		return FALSE;

#ifdef _M_IX86
	// Code for x86:
	// 33 C0                xor         eax,eax
	// C2 04 00             ret         4
	unsigned char szExecute[] = { 0x33, 0xC0, 0xC2, 0x04, 0x00 };
#elif _M_X64
	// 33 C0                xor         eax,eax
	// C3                   ret
	unsigned char szExecute[] = { 0x33, 0xC0, 0xC3 };
#else
#error "The following code only works for x86 and x64!"
#endif

	DWORD dwOldProtect = 0;
	BOOL  bProt = VirtualProtect(pOrgEntry, sizeof(szExecute), PAGE_EXECUTE_READWRITE, &dwOldProtect);

	SIZE_T bytesWritten = 0;
	BOOL   bRet = WriteProcessMemory(GetCurrentProcess(), pOrgEntry, szExecute, sizeof(szExecute),
		&bytesWritten);

	if ((bProt != FALSE) && (dwOldProtect != PAGE_EXECUTE_READWRITE))
	{
		DWORD dwBuf;
		VirtualProtect(pOrgEntry, sizeof(szExecute), dwOldProtect, &dwBuf);
	}
	return bRet;
}
#else
#pragma message("This code works only for x86 and x64!")
#endif

static TCHAR s_szExceptionLogFileName[_MAX_PATH] = _T("\\exceptions.log"); // default
static BOOL  s_bUnhandledExeptionFilterSet = FALSE;
static LONG __stdcall CrashHandlerExceptionFilter(EXCEPTION_POINTERS* pExPtrs)
{
#ifdef _M_IX86
	if (pExPtrs->ExceptionRecord->ExceptionCode == EXCEPTION_STACK_OVERFLOW)
	{
		static char MyStack[1024 * 128]; // be sure that we have enough space...
		// it assumes that DS and SS are the same!!! (this is the case for Win32)
		// change the stack only if the selectors are the same (this is the case for Win32)
		//__asm push offset MyStack[1024*128];
		//__asm pop esp;
		__asm mov eax, offset MyStack[1024 * 128];
		__asm mov esp, eax;
	}
#endif

	StackWalkerToConsole sw; // output to console
	sw.ShowCallstack(GetCurrentThread(), pExPtrs->ContextRecord);
	TCHAR lString[500];
	_stprintf_s(lString,
		_T("*** Unhandled Exception! See console output for more infos!\n")
		_T("   ExpCode: 0x%8.8X\n")
		_T("   ExpFlags: %d\n")
#if _MSC_VER >= 1900
		_T("   ExpAddress: 0x%8.8p\n")
#else
		_T("   ExpAddress: 0x%8.8X\n")
#endif
		_T("   Please report!"),
		pExPtrs->ExceptionRecord->ExceptionCode, pExPtrs->ExceptionRecord->ExceptionFlags,
		pExPtrs->ExceptionRecord->ExceptionAddress);
	FatalAppExit(-1, lString);
	return EXCEPTION_CONTINUE_SEARCH;
}

static void InitUnhandledExceptionFilter()
{
	TCHAR szModName[_MAX_PATH];
	if (GetModuleFileName(NULL, szModName, sizeof(szModName) / sizeof(TCHAR)) != 0)
	{
		_tcscpy_s(s_szExceptionLogFileName, szModName);
		_tcscat_s(s_szExceptionLogFileName, _T(".exp.log"));
	}
	if (s_bUnhandledExeptionFilterSet == FALSE)
	{
		// set global exception handler (for handling all unhandled exceptions)
		SetUnhandledExceptionFilter(CrashHandlerExceptionFilter);
#if defined _M_X64 || defined _M_IX86
		PreventSetUnhandledExceptionFilter();
#endif
		s_bUnhandledExeptionFilterSet = TRUE;
	}
}
#endif // UNHANDLED_EXCEPTION_TEST