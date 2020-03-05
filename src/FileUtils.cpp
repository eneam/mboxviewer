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


#include "stdafx.h"
#include "FileUtils.h"
#include "TextUtilsEx.h"
#include "SimpleString.h"


bool FileUtils::PathDirExists(LPCSTR path)
{
	DWORD dwFileAttributes = GetFileAttributes(path);
	if (dwFileAttributes != (DWORD)0xFFFFFFFF)
	{
		bool bRes = ((dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == FILE_ATTRIBUTE_DIRECTORY);
		return bRes;
	}
	else
		return false;
}

bool FileUtils::PathDirExistsW(LPCWSTR path)
{
	DWORD dwFileAttributes = GetFileAttributesW(path);
	if (dwFileAttributes != (DWORD)0xFFFFFFFF)
	{
		bool bRes = ((dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == FILE_ATTRIBUTE_DIRECTORY);
		return bRes;
	}
	else
		return false;
}

BOOL FileUtils::PathFileExist(LPCSTR path)
{
	DWORD dwFileAttributes = GetFileAttributes(path);
	if (dwFileAttributes != (DWORD)0xFFFFFFFF)
	{
		BOOL bRes = ((dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != FILE_ATTRIBUTE_DIRECTORY);
		return bRes;
	}
	else
		return FALSE;
}

BOOL FileUtils::PathFileExistW(LPCWSTR path)
{
	DWORD dwFileAttributes = GetFileAttributesW(path);
	if (dwFileAttributes != (DWORD)0xFFFFFFFF)
	{
		BOOL bRes = ((dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != FILE_ATTRIBUTE_DIRECTORY);
		return bRes;
	}
	else
		return FALSE;
}

int FileUtils::fileExists(LPCSTR file)
{
	DWORD dwFileAttributes = GetFileAttributes(file);
	if (dwFileAttributes != (DWORD)0xFFFFFFFF)
	{
		BOOL bRes = ((dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != FILE_ATTRIBUTE_DIRECTORY);
		return bRes;
}
	else
		return FALSE;
}


__int64 FileUtils::FileSeek(HANDLE hf, __int64 distance, DWORD MoveMethod)
{
	LARGE_INTEGER li;

	li.QuadPart = distance;

	li.LowPart = SetFilePointer(hf,
		li.LowPart,
		&li.HighPart,
		MoveMethod);

	if (li.LowPart == INVALID_SET_FILE_POINTER && GetLastError()
		!= NO_ERROR)
	{
		li.QuadPart = -1;
	}

	return li.QuadPart;
}

CString FileUtils::GetmboxviewTempPath(char *name)
{
	char	buf[_MAX_PATH + 1];
	DWORD gtp = ::GetTempPath(_MAX_PATH, buf);
	if (!FileUtils::PathDirExists(buf))
		strcpy(buf, "\\");
	strcat(buf, "mboxview\\");
	if (name) {
		strcat(buf, name);
		strcat(buf, "\\");
	}
	BOOL ret = ::CreateDirectory(buf, NULL);
	return buf;
}

CStringW FileUtils::GetmboxviewTempPathW(wchar_t *name)
{
	wchar_t	buf[_MAX_PATH + 1];
	DWORD gtp = ::GetTempPathW(_MAX_PATH, buf);

	if (!PathIsDirectoryW(buf))
		wcscpy(buf, L"\\");
	wcscat(buf, L"mboxview\\");
	if (name) {
		wcscat(buf, name);
		wcscat(buf, L"\\");
	}
	BOOL ret = ::CreateDirectoryW(buf, NULL);
	return buf;
}

BOOL FileUtils::RemoveDirectory(CString &dir, DWORD &error)
{
	error = 0;

	WIN32_FIND_DATA FileData;
	HANDLE hSearch;
	BOOL	bFinished = FALSE;
	// Start searching for all files in the current directory.
	CString searchPath = dir + "\\*.*";
	hSearch = FindFirstFile(searchPath, &FileData);
	if (hSearch == INVALID_HANDLE_VALUE) {
		TRACE("No files found.");
		return FALSE;
	}
	while (!bFinished)
	{
		if (!(strcmp(FileData.cFileName, ".") == 0 || strcmp(FileData.cFileName, "..") == 0))
		{
			CString fileFound = dir + "\\" + CString(FileData.cFileName);
			if (FileData.dwFileAttributes == FILE_ATTRIBUTE_DIRECTORY)
			{
				BOOL retR = RemoveDirectory(fileFound, error);
				if (retR == FALSE)
					bFinished = TRUE;
			}
		}
		if (!FindNextFile(hSearch, &FileData))
			bFinished = TRUE;
	}
	FindClose(hSearch);
	// dir must be empty
	BOOL retRD = ::RemoveDirectory(dir);
	if (!retRD)
		return FALSE;
	else
		return TRUE;
}

BOOL FileUtils::RemoveDirectoryW(CStringW &dir, DWORD &error)
{
	error = 0;

	WIN32_FIND_DATAW FileData;
	HANDLE hSearch;
	BOOL bFinished = FALSE;
	// Start searching for all files in the current directory.
	CStringW searchPath = dir + L"\\*.*";
	hSearch = FindFirstFileW(searchPath, &FileData);
	if (hSearch == INVALID_HANDLE_VALUE) {
		TRACE("No files found.");
		return FALSE;
	}
	while (!bFinished)
	{
		if (!(wcscmp(FileData.cFileName, L".") == 0 || wcscmp(FileData.cFileName, L"..") == 0))
		{
			CStringW	fileFound = dir + L"\\" + CStringW(FileData.cFileName);
			if (FileData.dwFileAttributes == FILE_ATTRIBUTE_DIRECTORY)
			{
				BOOL retR = RemoveDirectoryW(fileFound, error);
				if (retR == FALSE)
					bFinished = TRUE;
			}
		}
		if (!FindNextFileW(hSearch, &FileData))
			bFinished = TRUE;
	}
	FindClose(hSearch);
	// dir must be empty
	BOOL retRD = ::RemoveDirectoryW( dir );
	if (!retRD)
		return FALSE;
	else
		return TRUE;
}

// Removes files only
BOOL FileUtils::RemoveDir(CString & dir, bool recursive)
{
	WIN32_FIND_DATA FileData;
	HANDLE hSearch;
	BOOL	bFinished = FALSE;
	// Start searching for all files in the current directory.
	CString searchPath = dir + "\\*.*";
	hSearch = FindFirstFile(searchPath, &FileData);
	if (hSearch == INVALID_HANDLE_VALUE) {
		TRACE("No files found.");
		return FALSE;
	}
	while (!bFinished) 
	{
		if (!(strcmp(FileData.cFileName, ".") == 0 || strcmp(FileData.cFileName, "..") == 0)) 
		{
			CString	fileFound = dir + "\\" + FileData.cFileName;
			if (FileData.dwFileAttributes != FILE_ATTRIBUTE_DIRECTORY)
				DeleteFile((LPCSTR)fileFound);
			else if (recursive)
				RemoveDir(fileFound, recursive);
		}
		if (!FindNextFile(hSearch, &FileData))
			bFinished = TRUE;
	}
	FindClose(hSearch);
	//RemoveDirectory( dir );
	return TRUE;
}

BOOL FileUtils::RemoveDirW(CString & dir, bool recursive)
{
	CStringW dirW;
	DWORD error;
	BOOL retA2W = TextUtilsEx::Ansi2Wide(dir, dirW, error);
	BOOL ret = RemoveDirW(dirW, recursive);
	return ret;
}

// Removes files only
BOOL FileUtils::RemoveDirW(CStringW & dir, bool recursive)
{
	WIN32_FIND_DATAW FileData;
	HANDLE hSearch;
	BOOL	bFinished = FALSE;
	// Start searching for all files in the current directory.
	CStringW searchPath = dir + L"\\*.*";
	hSearch = FindFirstFileW(searchPath, &FileData);
	if (hSearch == INVALID_HANDLE_VALUE) {
		TRACE("No files found.");
		return FALSE;
	}
	while (!bFinished) 
	{
		if (!(wcscmp(FileData.cFileName, L".") == 0 || wcscmp(FileData.cFileName, L"..") == 0)) 
		{
			CStringW	fileFound = dir + L"\\" + CStringW(FileData.cFileName);
			if (FileData.dwFileAttributes != FILE_ATTRIBUTE_DIRECTORY)
				DeleteFileW((LPCWSTR)fileFound);
			else if (recursive)
				RemoveDirW(fileFound, recursive);
		}
		if (!FindNextFileW(hSearch, &FileData))
			bFinished = TRUE;
	}
	FindClose(hSearch);
	//RemoveDirectoryW( dir );
	return TRUE;
}

CString FileUtils::CreateTempFileName(CString ext)
{
	CString fmt = GetmboxviewTempPath() + "PTT%05d." + ext;
	CString fileName;
ripeti:
	fileName.Format(fmt, abs((int)(1 + (int)(100000.0*rand() / (RAND_MAX + 1.0)))));
	if (FileUtils::PathFileExist(fileName))
		goto ripeti;

	return fileName; //+_T(".HTM");
}

void FileUtils::CPathStripPath(const char *path, CString &fileName)
{
	int pathlen = strlen(path);
	char *pathbuff = new char[pathlen + 1];
	strcpy(pathbuff, path);
	PathStripPath(pathbuff);
	fileName.Empty();
	fileName.Append(pathbuff);
	delete[] pathbuff;
}

void FileUtils::CPathStripPathW(const wchar_t *path, CStringW &fileName)
{
	int pathlen = wcslen(path);
	wchar_t *pathbuff = new wchar_t[pathlen + 1];
	wcscpy(pathbuff, path);
	PathStripPathW(pathbuff);
	fileName.Empty();
	fileName.Append(pathbuff);
	delete[] pathbuff;
}

BOOL FileUtils::CPathGetPath(const char *path, CString &filePath)
{
	int pathlen = strlen(path);
	char *pathbuff = new char[2 * pathlen + 1];  // to avoid overrun ?? see microsoft docs
	strcpy(pathbuff, path);
	BOOL ret = PathRemoveFileSpec(pathbuff);
	//HRESULT  ret = PathCchRemoveFileSpec(pathbuff, pathlen);  //   pathbuff must be of PWSTR type

	filePath.Empty();
	filePath.Append(pathbuff);
	delete[] pathbuff;
	return ret;  // ret = 0; microsoft doc says if something goes wrong -:) 
}

void FileUtils::SplitFilePath(CString &fileName, CString &driveName, CString &directory, CString &fileNameBase, CString &fileNameExtention)
{
	TCHAR ext[_MAX_EXT + 1]; ext[0] = 0;
	TCHAR drive[_MAX_DRIVE + 1]; drive[0] = 0;
	TCHAR dir[_MAX_DIR + 1]; dir[0] = 0;
	TCHAR fname[_MAX_FNAME + 1]; fname[0] = 0;

	_tsplitpath_s(fileName,
		drive, _MAX_DRIVE + 1,
		dir, _MAX_DIR + 1,
		fname, _MAX_FNAME + 1,
		ext, _MAX_EXT + 1);

	driveName.Append(drive);
	directory.Append(dir);
	fileNameBase.Append(fname);
	fileNameExtention.Append(ext);
}


void FileUtils::UpdateFileExtension(CString &fileName, CString &newExtension)
{
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;
	SplitFilePath(fileName, driveName, directory, fileNameBase, fileNameExtention);

	fileName.Empty();
	fileName.Append(driveName);
	fileName.Append("\\");
	fileName.Append(directory);
	fileName.Append("\\");
	fileName.Append(fileNameBase);
	fileName.Append(newExtension);

}

_int64 FileUtils::FileSize(LPCSTR fileName)
{
	LARGE_INTEGER li = { 0 };
	HANDLE hFile = CreateFile(fileName, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE) {
		DWORD err = GetLastError();
		TRACE(_T("(FileSize)INVALID_HANDLE_VALUE error=%ld file=%s\n"), err, fileName);
		return li.QuadPart;
	}
	else {
		BOOL retval = GetFileSizeEx(hFile, &li);
		CloseHandle(hFile);
		if (retval != TRUE) {
			DWORD err = GetLastError();
			TRACE(_T("(GetFileSizeEx)Failed with error=%ld file=%s\n"), err, fileName);
		}
		else {
			long long fsize = li.QuadPart;
			return li.QuadPart;
		}
	}
	return li.QuadPart;
}

BOOL FileUtils::BrowseToFile(LPCTSTR filename)
{
	BOOL retval = TRUE;

	// It appears m_ie.Create  in CBrowser::OnCreate calls CoInitilize
	//Com_Initialize();

	ITEMIDLIST *pidl = ILCreateFromPath(filename);
	if (pidl) {
		HRESULT  ret = SHOpenFolderAndSelectItems(pidl, 0, 0, 0);
		if (ret != S_OK)
			retval = FALSE;
		ILFree(pidl);
	}
	else
		retval = FALSE;
	return retval;
}

BOOL FileUtils::BrowseToFileW(LPCWSTR filename)
{
	BOOL retval = TRUE;

	// It appears m_ie.Create  in CBrowser::OnCreate calls CoInitilize
	//Com_Initialize();

	ITEMIDLIST *pidl = ILCreateFromPathW(filename);
	if (pidl) {
		HRESULT  ret = SHOpenFolderAndSelectItems(pidl, 0, 0, 0);
		if (ret != S_OK)
			retval = FALSE;
		ILFree(pidl);
	}
	else
		retval = FALSE;
	return retval;
}

void  FileUtils::MakeValidFileName(CString &name, BOOL bReplaceWhiteWithUnderscore)
{
	SimpleString validName(name.GetLength());
	validName.Append((LPCSTR)name, name.GetLength());
	FileUtils::MakeValidFileName(validName, bReplaceWhiteWithUnderscore);
	name.Empty();
	name.Append(validName.Data(), validName.Count());
}

void  FileUtils::MakeValidFileName(SimpleString &name, BOOL bReplaceWhiteWithUnderscore)
{
	// Always merge consecutive underscore characters
	int csLen = name.Count();
	char* str = name.Data();
	int i;
	int j = 0;
	char c;
	BOOL allowUnderscore = TRUE;
	for (i = 0; i < csLen; i++)
	{
		c = str[i];
		if (
			(c == '?') || (c == '/') || (c == '<') || (c == '>') || (c == ':') ||  // not valid file name characters
			(c == '*') || (c == '|') || (c == '"') || (c == '\\')                  // not valid file name characters
			// I seem to remember that browser would not open a file with name with these chars, 
			// but now I can't recreate the case. One more reminder I need to keep track of all test cases
			// 1.0.3.7 || (c == ',') || (c == ';') || (c == '%')
			|| (c == '%')
			)
		{
			if (allowUnderscore)
			{
				allowUnderscore = FALSE;
				str[j++] = '_';
			}
		}
		else if (bReplaceWhiteWithUnderscore && ((c == ' ') || (c == '\t')))
		{
			if (allowUnderscore)
			{
				str[j++] = '_';
				allowUnderscore = FALSE;
			}
		}
		else if ((c < 32))
		{
			if (allowUnderscore)
			{
				str[j++] = '_';
				allowUnderscore = FALSE;
			}
		}
		else
		{
			str[j++] = c;
			allowUnderscore = TRUE;
		}
	}
	name.SetCount(j);
}

void  FileUtils::MakeValidFileNameW(CStringW &name, CStringW &result, BOOL bReplaceWhiteWithUnderscore)
{
	// Always merge consecutive underscore characters
	int csLen = name.GetLength();
	int i;

	wchar_t c;
	BOOL allowUnderscore = TRUE;
	for (i = 0; i < csLen; i++)
	{
		c = name[i];
		if (
			(c == L'?') || (c == L'/') || (c == L'<') || (c == L'>') || (c == L':') ||  // not valid file name characters
			(c == L'*') || (c == L'|') || (c == L'"') || (c == L'\\')                  // not valid file name characters
			// I seem to remember that browser would not open a file with name with these chars, 
			// but now I can't recreate the case. One more reminder I need to keep track of all test cases
			// 1.0.3.7 || (c == ',') || (c == ';') || (c == '%')
			|| (c == L'%')
			)
		{
			if (allowUnderscore)
			{
				allowUnderscore = FALSE;
				result.AppendChar(L'_');
			}
		}
		else if (bReplaceWhiteWithUnderscore && ((c == L' ') || (c == L'\t')))
		{
			if (allowUnderscore)
			{
				result.AppendChar(L'_');
				allowUnderscore = FALSE;
			}
		}
		else if (iswcntrl(c))
		{
			if (allowUnderscore)
			{
				result.AppendChar(L'_');
				allowUnderscore = FALSE;
			}
		}
		else
		{
			result.AppendChar(c);
			allowUnderscore = TRUE;
		}
	}
}

BOOL FileUtils::Write2File(CStringW &cStrNamePath, const unsigned char *data, int dataLength)
{
	DWORD dwAccess = GENERIC_WRITE;
	DWORD dwCreationDisposition = CREATE_ALWAYS;

	HANDLE hFile = CreateFileW(cStrNamePath, dwAccess, FILE_SHARE_READ, NULL, dwCreationDisposition, 0, NULL);
	if (hFile == INVALID_HANDLE_VALUE)
	{
		int deb = 1;
		return FALSE;
	}
	else
	{
		const unsigned char* pszData = data;
		int nLeft = dataLength;

		for (;;)
		{
			DWORD nWritten = 0;

			BOOL retWrite = WriteFile(hFile, pszData, nLeft, &nWritten, 0);
			if (nWritten != nLeft)
				int deb = 1;

			if (nWritten < 0)
				break;

			pszData += nWritten;
			nLeft -= nWritten;
			if (nLeft <= 0)
				break;
		}

		BOOL retClose = CloseHandle(hFile);
		if (retClose == FALSE)
			int deb = 1;
	}
	return TRUE;
}

int FileUtils::Write2File(HANDLE hFile, LPCVOID lpBuffer, DWORD nNumberOfBytesToWrite, LPDWORD lpNumberOfBytesWritten)
{
	DWORD bytesToWrite = nNumberOfBytesToWrite;
	DWORD nwritten = 0;
	while (bytesToWrite > 0)
	{
		nwritten = 0;
		if (!WriteFile(hFile, lpBuffer, bytesToWrite, &nwritten, NULL)) {
			DWORD retval = GetLastError();
			break;
		}
		bytesToWrite -= nwritten;
	}
	*lpNumberOfBytesWritten = nNumberOfBytesToWrite - bytesToWrite;
	if (*lpNumberOfBytesWritten != nNumberOfBytesToWrite)
		return FALSE;
	else
		return TRUE;
}

void FileUtils::UnitTest()
{
	DWORD error;
	bool retPathExists;
	BOOL retRemoveDirectory;
	BOOL retRemoveFiles;

	CString pathA;
	CString subpathA;
	CStringW pathW;
	CStringW subpathW;

	CString safeDir = "AppData\\Local\\Temp";
	CStringW safeDirW = L"AppData\\Local\\Temp";

	/////////////////////////////////////////
	pathW = GetmboxviewTempPathW();
	if (pathW.Find(safeDirW) < 0)
		ASSERT(false);
	retRemoveFiles = RemoveDirW(pathW, true);
	retRemoveDirectory = RemoveDirectoryW(pathW, error);

	pathW = GetmboxviewTempPathW();
	retPathExists = PathDirExistsW(pathW);
	ASSERT(retPathExists);
	retRemoveFiles = RemoveDirW(pathW, true);
	retRemoveDirectory = RemoveDirectoryW(pathW, error);
	ASSERT(retRemoveDirectory);

	pathW = GetmboxviewTempPathW();
	retPathExists = PathDirExistsW(pathW);
	subpathW = GetmboxviewTempPathW(L"Inline");
	retPathExists = PathDirExistsW(subpathW);
	ASSERT(retPathExists);
	retRemoveFiles = RemoveDirW(subpathW, true);
	//retRemoveDirectory = RemoveDirectoryW(subpathW, error);
	retRemoveDirectory = RemoveDirectoryW(pathW, error);
	ASSERT(retRemoveDirectory);

	/////////////////////////////////////////

	/////////////////////////////////
	pathA = GetmboxviewTempPath();
	if (pathA.Find(safeDir) < 0)
		ASSERT(false);
	retRemoveFiles = RemoveDir(pathA, true);
	retRemoveDirectory = RemoveDirectory(pathA, error);

	pathA = GetmboxviewTempPath();
	retPathExists = PathDirExists(pathA);
	ASSERT(retPathExists);
	retRemoveFiles = RemoveDir(pathA, true);
	retRemoveDirectory = RemoveDirectory(pathA, error);
	ASSERT(retRemoveDirectory);


	pathA = GetmboxviewTempPath();
	retPathExists = PathDirExists(pathA);
	subpathA = GetmboxviewTempPath("Inline");
	retPathExists = PathDirExists(subpathA);
	ASSERT(retPathExists);
	//retRemoveDirectory = RemoveDirectory(subpathA, error);
	retRemoveDirectory = RemoveDirectoryA(pathA, error);
	ASSERT(retRemoveDirectory);


}