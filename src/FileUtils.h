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
	static BOOL IsReadonlyFolder(CString &folderPath);
	//
	static int fileExists(LPCSTR file);
	static bool PathDirExists(LPCSTR file);
	static bool PathDirExistsW(CString &dir);
	static bool PathDirExistsW(LPCWSTR file);
	//
	static BOOL PathFileExist(LPCSTR path);
	static BOOL PathFileExistW(LPCWSTR path);
	//
	static __int64 FileSeek(HANDLE hf, __int64 distance, DWORD MoveMethod);
	//
	static CString GetMboxviewTempPath(const char *name = 0);
	static CStringW GetMboxviewTempPathW(const wchar_t *name = 0);
	//
	static CString GetMboxviewLocalAppPath();
	static CString CreateMboxviewLocalAppPath();
	static CString GetMboxviewLocalAppDataPath(const char *name = 0);
	static CStringW GetMboxviewLocalAppDataPathW(const wchar_t *name = 0);
	static CString CreateMboxviewLocalAppDataPath(const char *name = 0);
	static CStringW CreateMboxviewLocalAppDataPathW(const wchar_t *name = 0);
	//
	static BOOL OSRemoveDirectory(CString &dir, CString &errorText);
	static BOOL OSRemoveDirectoryW(CStringW &dir, CStringW &errorText);
	static BOOL RemoveDirectory(CString &dir, DWORD &error);
	static BOOL RemoveDirectoryW(CStringW &dir, DWORD &error);
	//
	static BOOL RemoveDir(CString & dir, bool recursive = false, bool removeFolders = false);
	static BOOL RemoveDirW(CString & dir, bool recursive = false, bool removeFolders = false);
	static BOOL RemoveDirW(CStringW & dir, bool recursive = false, bool removeFolders = false);
	static CString CreateTempFileName(CString ext = "htm");
	//
	static void CPathStripPath(const char *path, CString &fileName);
	static void CPathStripPathW(const wchar_t *path, CStringW &fileName);
	static BOOL CPathGetPath(const char *path, CString &filePath);
	//
	static void SplitFilePath(CString &fileName, CString &driveName, CString &directory, CString &fileNameBase, CString &fileNameExtention);
	static void SplitFilePathW(CStringW &fileName, CStringW &driveName, CStringW &directory, CStringW &fileNameBase, CStringW &fileNameExtention);
	//
	static void GetFolderPathW(CStringW &fileNamePath, CStringW &folderPath);
	static void GetFolderPath(CString &fileNamePath, CString &folderPath);
	static void GetFolderPathAndFileNameW(CStringW &fileNamePath, CStringW &folderPath, CStringW &fileName);
	static void GetFolderPathAndFileName(CString &fileNamePath, CString &folderPath, CString &fileName);
	static void GetFileBaseNameAndExtension(CString &fileName, CString &fileNameExtention, CString &fileBaseName);
	static void GetFileBaseName(CString &fileName, CString &fileBaseName);
	static void GetFileName(CString &fileNamePath, CString &fileName);
	static void GetFileExtension(CString &fileName, CString &fileNameExtention);
	static void UpdateFileExtension(CString &fileName, CString &newExtension);
	static _int64 FileSize(LPCSTR fileName, CString *errorText = 0);
	static _int64 FolderSize(LPCSTR folderPath);
	static int GetFolderFileCount(CString &folderPath, BOOL recursive = FALSE);
	//
	static BOOL BrowseToFileW(LPCWSTR filename);
	static BOOL BrowseToFile(LPCTSTR filename);
	//
	static void MakeValidFilePath(CString &path, BOOL bReplaceWhiteWithUnderscore = TRUE);
	static void MakeValidFilePath(SimpleString &path, BOOL bReplaceWhiteWithUnderscore = TRUE);
	static void MakeValidFileName(CString &name, BOOL bReplaceWhiteWithUnderscore = TRUE);
	static void MakeValidFileName(SimpleString &name, BOOL bReplaceWhiteWithUnderscore = TRUE);
	static void MakeValidFileNameW(CStringW &name, CStringW &result, BOOL bReplaceWhiteWithUnderscore);
	//
	static void MakeValidLabelFilePath(CString &path, BOOL bReplaceWhiteWithUnderscore = TRUE);
	static void MakeValidLabelFilePath(SimpleString &path, BOOL bReplaceWhiteWithUnderscore = TRUE);
	//
	static BOOL Write2File(CStringW &cStrNamePath, const unsigned char *data, int dataLength);
	static BOOL Write2File(CString &cStrNamePath, const unsigned char *data, int dataLength);
	static BOOL ReadEntireFile(CString &cStrNamePath, SimpleString &data);
	static int Write2File(HANDLE hFile, LPCVOID lpBuffer, DWORD nNumberOfBytesToWrite, LPDWORD lpNumberOfBytesWritten);
	static BOOL NormalizeFilePath(CString &filePath);
	// Create all subfolders if they don't exist
	static BOOL CreateDirectory(const char *path);
	static BOOL CreateDirectoryW(const wchar_t *path);
	//
	static BOOL DeleteFile(const char *path, BOOL verify = FALSE);
	static BOOL DeleteFileW(const wchar_t *path, BOOL verify = FALSE);
	static BOOL DeleteFile(CString &path, BOOL verify = FALSE);
	static BOOL DeleteFileW(CStringW &path, BOOL verify = FALSE);
	//
	static CString CopyDirectory(const char *cFromPath, const char* cToPath, BOOL bFailIfExists = FALSE, BOOL removeFolderAfterCopy = FALSE);
	static CStringW CopyDirectoryW(const wchar_t *cFromPath, const wchar_t *cToPath, BOOL bFailIfExists = FALSE, BOOL removeFolderAfterCopy = FALSE);
	static CStringW CopyDirectoryW(const char *cFromPath, const char *cToPath, BOOL bFailIfExists = FALSE, BOOL removeFolderAfterCopy = FALSE);
	//
	static CString MoveDirectory(const char *cFromPath, const char *cToPath);
	static CStringW MoveDirectoryW(const wchar_t *cFromPath, const wchar_t *cToPath);
	//
	static CString GetLastErrorAsString();
	static CStringW GetLastErrorAsStringW();

	static CString GetFileExceptionErrorAsString(CFileException &ExError);
	// Generic text
	static CString GetOpenForReadFileExceptionErrorAsString(CString &fileName, CFileException &exError);

	static BOOL GetFolderList(CString &rootFolder, CList<CString, CString &> &folderList, CString &errorText, int maxDepth);

	static BOOL VerifyName(CString &name);
	static BOOL VerifyNameW(CStringW &name);

	void UnitTest();
};

