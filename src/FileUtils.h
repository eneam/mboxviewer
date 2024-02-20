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
	static bool PathDirExists(CString &dir);
	static bool PathDirExists(LPCWSTR file);
	//
	static BOOL PathFileExist(CString &path);
	static BOOL PathFileExist(LPCWSTR path);
	//
	static __int64 FileSeek(HANDLE hf, __int64 distance, DWORD MoveMethod);
	//
	static CString GetMboxviewTempPath(const wchar_t* folderName, const wchar_t* subfolderName = 0);
	//
	static CString GetMboxviewLocalAppPath();  // Windows created application folder path such as C:\Users\tata\AppData\Local
	static CString CreateMboxviewLocalAppPath();  // Redundant. Windows should already created application folder path such as C:\Users\tata\AppData\Local
	//
	static CString GetMboxviewLocalAppDataPath(const wchar_t* folderName, const wchar_t *subfoldderName = 0);
	static CString CreateMboxviewLocalAppDataPath(const wchar_t *name = 0);
	//
	static BOOL OSRemoveDirectory(CString &dir, CString &errorText);
	static BOOL RemoveDir(CString & dir, bool recursive = false, bool removeFolders = false, CString *errorText = 0);
	//
	static CString CreateTempFileName(const wchar_t *folderName, CString ext = L"htm");
	//
	static void CPathStripPath(const wchar_t *path, CString &fileName);
	static BOOL CPathGetPath(const wchar_t *path, CString &filePath);
	//
	static void SplitFilePath(CString &fileName, CString &driveName, CString &directory, CString &fileNameBase, CString &fileNameExtention);
	static void SplitFileFolder(CString& path, CStringArray& a);
	//
	static void GetFolderPath(CString &fileNamePath, CString &folderPath);
	static void GetFolderPathAndFileName(CString &fileNamePath, CString &folderPath, CString &fileName);
	//
	static void GetFileBaseNameAndExtension(CString &fileName, CString &fileBaseName, CString &fileNameExtention);
	static void GetFileBaseName(CString &fileName, CString &fileBaseName);
	static void GetFileName(CString &fileNamePath, CString &fileName);
	static void GetFileExtension(CString &fileName, CString &fileNameExtention);
	static void UpdateFileExtension(CString &fileName, CString &newExtension);
	//
	static _int64 FileSize(LPCWSTR fileName, CString *errorText = 0);
	static _int64 FolderSize(LPCWSTR folderPath);
	static int GetFolderFileCount(CString &folderPath, BOOL recursive = FALSE);
	//
	static BOOL BrowseToFile(LPCWSTR filename);
	//
	static void MakeValidFileName(const wchar_t* name, int namelen, CString& result, BOOL bReplaceWhiteWithUnderscore, BOOL extraValidation);
	static void MakeValidFileName(CString& name, BOOL bReplaceWhiteWithUnderscore, BOOL extraValidation);
	static void MakeValidFileName(CString& name, CString& result, BOOL bReplaceWhiteWithUnderscore, BOOL extraValidation);
	//
	static void MakeValidFileNameA(const char* name, int namelen, CStringA& result, BOOL bReplaceWhiteWithUnderscore, BOOL extraValidation);
	static void MakeValidFileNameA(CStringA& name, CStringA& result, BOOL bReplaceWhiteWithUnderscore, BOOL extraValidation);
	static void MakeValidFileNameA(CStringA& name, BOOL bReplaceWhiteWithUnderscore, BOOL extraValidation);
	//
	static BOOL Write2File(CString &cStrNamePath, const unsigned char *data, int dataLength);
	static BOOL ReadEntireFile(CString &cStrNamePath, SimpleString &data);
	static int Write2File(HANDLE hFile, LPCVOID lpBuffer, DWORD nNumberOfBytesToWrite, LPDWORD lpNumberOfBytesWritten);
	//
	static BOOL NormalizeFilePath(CString &filePath);
	//
	// Create all subfolders if they don't exist
	static BOOL CreateDir(const wchar_t *path);
	//
	static BOOL DelFile(const wchar_t *path, BOOL verify = TRUE);  // It was FALSE, chnaged to see if that vcan remain TRUE
	static BOOL DelFile(CString &path, BOOL verify = TRUE);  // It was FALSE, chnaged to see if that vcan remain TRUE
	//
	static BOOL WCopyFile(LPCWSTR lpExistingFileName, LPCWSTR lpNewFileName, BOOL bFailIfExists, CString &errorText);
	static BOOL ACopyFile(LPCSTR lpExistingFileName, LPCSTR lpNewFileName, BOOL bFailIfExists, CString &errorText);
	//
	static CStringW CopyDirectory(const wchar_t* cFromPath, const wchar_t* cToPath, CStringArray *excludeFilter = 0, BOOL bFailIfExists = FALSE, BOOL removeFolderAfterCopy = FALSE);
	//
	static CString MoveDirectory(const wchar_t *cFromPath, const wchar_t *cToPath);
	//
	static CString GetLastErrorAsString(DWORD errorCode = 0xFFFF);

	static CString GetFileExceptionErrorAsString(CFileException &ExError);
	static CString GetOpenForReadFileExceptionErrorAsString(CString &fileName, CFileException &exError);

	static BOOL GetFolderList(CString &rootFolder, CList<CString, CString &> &folderList, CString &errorText, int maxDepth);

	static BOOL VerifyName(CString &name);

	static CString SizeToString(_int64 size);

	void UnitTest();
};

