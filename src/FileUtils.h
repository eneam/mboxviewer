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


#pragma once

class SimpleString;


class FileUtils
{
public:
	static int fileExists(LPCSTR file);
	static bool PathDirExists(LPCSTR file);
	static bool PathDirExistsW(LPCWSTR file);
	static BOOL PathFileExist(LPCSTR path);
	static BOOL PathFileExistW(LPCWSTR path);
	static __int64 FileSeek(HANDLE hf, __int64 distance, DWORD MoveMethod);
	static CString GetmboxviewTempPath(char *name = 0);
	static CStringW GetmboxviewTempPathW(wchar_t *name = 0);
	BOOL RemoveDirectory(CString &dir, DWORD &error);
	BOOL RemoveDirectoryW(CStringW &dir, DWORD &error);
	static BOOL RemoveDir(CString & dir, bool recursive = false);
	static BOOL RemoveDirW(CString & dir, bool recursive = false);
	static BOOL RemoveDirW(CStringW & dir, bool recursive = false);
	static CString CreateTempFileName(CString ext = "htm");
	static void CPathStripPath(const char *path, CString &fileName);
	static void CPathStripPathW(const wchar_t *path, CStringW &fileName);
	static BOOL CPathGetPath(const char *path, CString &filePath);
	static void SplitFilePath(CString &fileName, CString &driveName, CString &directory, CString &fileNameBase, CString &fileNameExtention);
	static void GetFolderPathAndFileName(CString &fileNamePath, CString &folderPath, CString &fileName);
	static void GetFileBaseNameAndExtension(CString &fileName, CString &fileNameExtention, CString &fileBaseName);
	static void GetFileBaseName(CString &fileName, CString &fileBaseName);
	static void GetFileExtension(CString &fileName, CString &fileNameExtention);
	static void UpdateFileExtension(CString &fileName, CString &newExtension);
	static _int64 FileSize(LPCSTR fileName);
	static BOOL BrowseToFileW(LPCWSTR filename);
	static BOOL BrowseToFile(LPCTSTR filename);
	static void MakeValidFileName(CString &name, BOOL bReplaceWhiteWithUnderscore = TRUE);
	static void MakeValidFileName(SimpleString &name, BOOL bReplaceWhiteWithUnderscore = TRUE);
	static void MakeValidFileNameW(CStringW &name, CStringW &result, BOOL bReplaceWhiteWithUnderscore);
	static BOOL Write2File(CStringW &cStrNamePath, const unsigned char *data, int dataLength);
	static int Write2File(HANDLE hFile, LPCVOID lpBuffer, DWORD nNumberOfBytesToWrite, LPDWORD lpNumberOfBytesWritten);
	static BOOL NormalizeFilePath(CString &filePath);
	//
	void UnitTest();
};

