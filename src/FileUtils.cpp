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

bool FileUtils::PathDirExists(CString &dir)
{
	bool ret = FileUtils::PathDirExists((LPCWSTR)dir);
	return ret;

}

bool FileUtils::PathDirExists(LPCWSTR path)
{
	DWORD dwFileAttributes = GetFileAttributesW(path);
	if (dwFileAttributes == INVALID_FILE_ATTRIBUTES)
	{
		// may return invalid path or invalid file - not sure why
		CStringW retW = GetLastErrorAsString();  
		TRACE(L"%s : \"%s\"\n", retW, path);
		int deb = 1;
	}

	if (dwFileAttributes != INVALID_FILE_ATTRIBUTES)
	{
		bool bRes = ((dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == FILE_ATTRIBUTE_DIRECTORY);
		return bRes;
	}
	else
	{
		return false;
	}
}

BOOL FileUtils::PathFileExist(CString &path)
{
	BOOL ret = PathFileExist((LPCWSTR)path);
	return ret;
}

BOOL FileUtils::PathFileExist(LPCWSTR path)
{
	DWORD dwFileAttributes = GetFileAttributes(path);
	if (dwFileAttributes != INVALID_FILE_ATTRIBUTES)
	{
		BOOL bRes = ((dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != FILE_ATTRIBUTE_DIRECTORY);
		return bRes;
	}
	else
	{
		CString errText = FileUtils::GetLastErrorAsString();
		DWORD err = GetLastError();
		TRACE(L"(FileSize)error=%ld file=%s\n%s\n", err, path, errText);
		return FALSE;
	}
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

CString FileUtils::GetMboxviewTempPath(const wchar_t *folderName, const wchar_t *subfolderName)
{
	wchar_t	buf[_MAX_PATH + 1];
	buf[0] = 0;
	const DWORD gtp = ::GetTempPathW(_MAX_PATH, buf);
	//
	CStringW localTempPath = CStringW(&buf[0]);
	localTempPath.TrimRight(L"\\");
	localTempPath.Append(L"\\");
	localTempPath.Append(folderName);
	localTempPath.Append(L"\\");
	if (subfolderName)
	{
		localTempPath.Append(subfolderName);
		localTempPath.Append(L"\\");
	}
	BOOL ret = TRUE;
	if (!FileUtils::PathDirExists(localTempPath))
		ret = FileUtils::CreateDir(localTempPath);
	return localTempPath;
}

CString FileUtils::GetMboxviewLocalAppPath()
{
	wchar_t	buf[_MAX_PATH + 1];
	buf[0] = 0;
	buf[_MAX_PATH] = 0;
	const DWORD gtp = ::GetTempPathW(_MAX_PATH, buf);

	CString locaAppTempFolderPath = CString(&buf[0]);
	locaAppTempFolderPath.TrimRight(L"\\");

	CString folderPath;
	FileUtils::GetFolderPath(locaAppTempFolderPath, folderPath);

	return folderPath;
}

CString FileUtils::CreateMboxviewLocalAppPath()
{
	CString folderPath = GetMboxviewLocalAppPath();
	BOOL ret = TRUE;
	// It should be created by the OS already. Eliminate below?
	if (!FileUtils::PathDirExists(folderPath)) 
	{
		ret = FileUtils::CreateDir(folderPath);
		if (ret == FALSE)
		{
			folderPath.Empty();
		}
	}
	return folderPath;
}

CString FileUtils::GetMboxviewLocalAppDataPath(const wchar_t* folderName, const wchar_t* subfolderName)
{
	wchar_t	buf[_MAX_PATH + 1];
	buf[0] = 0;
	const DWORD gtp = ::GetTempPathW(_MAX_PATH, buf);

	CStringW folderPath;
	CStringW fileName;

	CStringW localTempPath = CStringW(&buf[0]);
	localTempPath.TrimRight(L"\\");
	FileUtils::GetFolderPathAndFileName(localTempPath, folderPath, fileName);
	folderPath.Append(folderName);
	folderPath.Append(L"\\");
	if (subfolderName) {
		folderPath.Append(subfolderName);
		folderPath.Append(L"\\");
	}
	BOOL ret = TRUE;
	if (!FileUtils::PathDirExists(folderPath))
		ret = FileUtils::CreateDir(folderPath);

	return folderPath;
}

CString FileUtils::CreateMboxviewLocalAppDataPath(const wchar_t *name)
{
	CStringW folderPath = FileUtils::GetMboxviewLocalAppDataPath(name);
	BOOL ret = TRUE;
	if (!FileUtils::PathDirExists(folderPath))
		ret = FileUtils::CreateDir(folderPath);

	return folderPath;
}

BOOL FileUtils::OSRemoveDirectory(CString &dir, CString &errorText)
{
	const BOOL ret = ::RemoveDirectoryW(dir);
	if (!ret)
	{
		errorText = FileUtils::GetLastErrorAsString();
		TRACE(L"OSRemoveDirectory: \"%s\" %s", dir, errorText);
	}
	return ret;
}

// Removes files and optionally folders
BOOL FileUtils::RemoveDir(CString & directory, bool recursive, bool removeFolders, CString *errorText)
{
	CString errText;
	if (!FileUtils::VerifyName(directory))
	{
		TRACE(L"RemoveDir: Failed to delete, VerifyName failed \"%s\"\n", directory);
		_ASSERTE(FALSE);
		return FALSE;
	}

	WIN32_FIND_DATAW FileData;
	HANDLE hSearch;
	BOOL	bFinished = FALSE;

	// Start searching for all files in the current directory.
	CStringW dir = directory;
	dir.TrimRight(L"\\");
	dir.Append(L"\\");

	CStringW searchPath = dir + L"*.*";
	hSearch = FindFirstFileW(searchPath, &FileData);
	if (hSearch == INVALID_HANDLE_VALUE) {
		TRACE(L"RemoveDir: No files found in \"%s\".\n", dir);
		return FALSE;
	}
	while (!bFinished) 
	{
		if (!(wcscmp(FileData.cFileName, L".") == 0 || wcscmp(FileData.cFileName, L"..") == 0)) 
		{
			CStringW	fileFound = dir + CStringW(FileData.cFileName);
			if (FileData.dwFileAttributes != FILE_ATTRIBUTE_DIRECTORY)
				FileUtils::DelFile(fileFound, TRUE);
			else if (recursive)
			{
				CString eText;
				FileUtils::RemoveDir(fileFound, recursive, removeFolders, &eText);
				if (!eText.IsEmpty())
					int deb = 1;
			}
		}
		if (!FindNextFileW(hSearch, &FileData))
			bFinished = TRUE;
	}
	FindClose(hSearch);
	if (removeFolders)
	{
		// Will fail if recursive=false and sub-directories exist
		BOOL retRD = FileUtils::OSRemoveDirectory(dir, errText);
		if (errorText)
		{
			errorText->Empty();
			errorText->Append(errText);
		}
	}
	return TRUE;
}

CString FileUtils::CreateTempFileName(const wchar_t *folderName, CString ext)
{
	CString fmt = GetMboxviewTempPath(folderName) + "PTT%05d." + ext;
	CString fileName;
ripeti:
	fileName.Format(fmt, abs((int)(1 + (int)(100000.0*rand() / (RAND_MAX + 1.0)))));
	if (FileUtils::PathFileExist(fileName))
		goto ripeti;

	return fileName; //+L".HTM";
}

void FileUtils::CPathStripPath(const wchar_t *path, CStringW &fileName)
{
	int pathlen = (int)wcslen(path);
	wchar_t *pathbuff = new wchar_t[pathlen + 1];
	wcscpy(pathbuff, path);
	PathStripPathW(pathbuff);
	fileName.Empty();
	fileName.Append(pathbuff);
	delete[] pathbuff;
}

BOOL FileUtils::CPathGetPath(const wchar_t *path, CString &filePath)
{
	int pathlen = iwstrlen(path);
	wchar_t *pathbuff = new wchar_t[2 * pathlen + 1];  // to avoid overrun ?? see microsoft docs
	_tcscpy(pathbuff, path);
	BOOL ret = PathRemoveFileSpec(pathbuff);
	//HRESULT  ret = PathCchRemoveFileSpec(pathbuff, pathlen);  //   pathbuff must be of PWSTR type

	filePath.Empty();
	filePath.Append(pathbuff);
	delete[] pathbuff;
	return ret;  // ret = 0; microsoft doc says if something goes wrong -:) 
}

void FileUtils::SplitFilePath(CString &fileName, CString &driveName, CString &directory, CString &fileNameBase, CString &fileNameExtention)
{
	wchar_t ext[_MAX_EXT + 1]; ext[0] = 0;
	wchar_t drive[_MAX_DRIVE + 1]; drive[0] = 0;
	wchar_t dir[_MAX_DIR + 1]; dir[0] = 0;
	wchar_t fname[_MAX_FNAME + 1]; fname[0] = 0;

	driveName.Empty();
	directory.Empty();
	fileNameBase.Empty();
	fileNameExtention.Empty();

	_wsplitpath_s(fileName,
		drive, _MAX_DRIVE + 1,
		dir, _MAX_DIR + 1,
		fname, _MAX_FNAME + 1,
		ext, _MAX_EXT + 1);

	driveName.Append(drive);
	directory.Append(dir);
	fileNameBase.Append(fname);
	fileNameExtention.Append(ext);
}

void FileUtils::GetFolderPath(CString &fileNamePath, CString &folderPath)
{
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;
	SplitFilePath(fileNamePath, driveName, directory, fileNameBase, fileNameExtention);
	folderPath = driveName + directory;
}

void FileUtils::GetFolderPathAndFileName(CString &fileNamePath, CString &folderPath, CString &fileName)
{
	CStringW driveName;
	CStringW directory;
	CStringW fileNameBase;
	CStringW fileNameExtention;
	SplitFilePath(fileNamePath, driveName, directory, fileNameBase, fileNameExtention);
	folderPath = driveName + directory;
	fileName = fileNameBase + fileNameExtention;
}

void FileUtils::GetFileBaseNameAndExtension(CString &fileName, CString &fileNameBase, CString &fileNameExtention)
{
	CString driveName;
	CString directory;
	SplitFilePath(fileName, driveName, directory, fileNameBase, fileNameExtention);
}

void FileUtils::GetFileBaseName(CString &fileName, CString &fileBaseName)
{
	CString driveName;
	CString directory;
	//CString fileNameBase;
	CString fileNameExtention;
	SplitFilePath(fileName, driveName, directory, fileBaseName, fileNameExtention);
}

void FileUtils::GetFileName(CString &fileNamePath, CString &fileName)
{
	CString driveName;
	CString directory;
	CString fileBaseName;
	CString fileNameExtention;
	SplitFilePath(fileNamePath, driveName, directory, fileBaseName, fileNameExtention);
	fileName = fileBaseName + fileNameExtention;
}

void FileUtils::GetFileExtension(CString &fileName, CString &fileNameExtention)
{
	CString driveName;
	CString directory;
	CString fileNameBase;
	SplitFilePath(fileName, driveName, directory, fileNameBase, fileNameExtention);
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
	//fileName.Append("\\");
	fileName.Append(directory);
	//fileName.Append("\\");
	fileName.Append(fileNameBase);
	fileName.Append(newExtension);

}

_int64 FileUtils::FileSize(LPCWSTR fileName, CString *errorText)
{
	LARGE_INTEGER li = { 0 };
	HANDLE hFile = CreateFile(fileName, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE)
	{
		CString errText = FileUtils::GetLastErrorAsString();
		DWORD err = GetLastError();
		TRACE(L"(FileSize)INVALID_HANDLE_VALUE error=%ld file=%s\n%s\n", err, fileName, errText);
		if (errorText)
			errorText->Append(errText);
		return li.QuadPart;
	}
	else
	{
		BOOL retval = GetFileSizeEx(hFile, &li);
		CloseHandle(hFile);
		if (retval != TRUE)
		{
			CString errText = FileUtils::GetLastErrorAsString();
			DWORD err = GetLastError();
			TRACE(L"(GetFileSizeEx)Failed with error=%ld file=%s\n%s\n", err, fileName,errText);
			if (errorText)
				errorText->Append(errText);
		}
		else
		{
			long long fsize = li.QuadPart;
			return li.QuadPart;
		}
	}
	return li.QuadPart;
}

BOOL FileUtils::BrowseToFile(LPCWSTR filename)
{
	BOOL retval = TRUE;

	// It appears m_ie.Create  in CBrowser::OnCreate calls CoInitilize
	//Com_Initialize();

	__unaligned ITEMIDLIST *pidl = ILCreateFromPathW(filename);
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

void  FileUtils::MakeValidFileName(CString &name, BOOL bReplaceWhiteWithUnderscore, BOOL extraValidation)
{
	CString result;
	FileUtils::MakeValidFileName(name, result, bReplaceWhiteWithUnderscore, extraValidation);

	name.Empty();
	name.Append(result);
}

void  FileUtils::MakeValidFileNameA(CStringA& name, BOOL bReplaceWhiteWithUnderscore, BOOL extraValidation)
{
	CStringA result;
	FileUtils::MakeValidFileNameA(name, result, bReplaceWhiteWithUnderscore, extraValidation);

	name.Empty();
	name.Append(result);
}

void  FileUtils::MakeValidFileNameA(CStringA& name, CStringA& result, BOOL bReplaceWhiteWithUnderscore, BOOL extraValidation)
{
	FileUtils::MakeValidFileNameA((LPCSTR)name, name.GetLength(), result, bReplaceWhiteWithUnderscore, extraValidation);
}

// name can be UTF8 or ansi or ascii. Replace invalid file name charcater with underscore or space
void  FileUtils::MakeValidFileNameA(const char* name, int namelen, CStringA& result, BOOL bReplaceWhiteWithUnderscore, BOOL extraValidation)
{
	// Always merge consecutive underscore characters
	result.Preallocate(namelen + 2);

	int csLen = namelen;
	int i;

	int c;
	BOOL allowUnderscore = TRUE;
	for (i = 0; i < csLen; i++)
	{
		c = name[i] & 0xFF;
		if (
			(c == '?') || (c == '<') || (c == '>') || (c == ':') ||  // not valid file name characters on Windows
			(c == '*') || (c == '|') || (c == '"')                 // not valid file name characters
			|| (c == '%')     // sometimes pdfbox, Edge and Chrome have problems
			)
		{
			if (allowUnderscore)
			{
				allowUnderscore = FALSE;
				result.AppendChar('_');
			}
		}
		else if (extraValidation &&
			( (c == '#') || (c == '\'')   // sometimes pdfbox, Edge and Chrome have problems
			  || (c == '!') )             // breaks cmd batch files, breaks Print Merge option
			)             
		{
			if (allowUnderscore)
			{
				allowUnderscore = FALSE;
					result.AppendChar('_');
			}
		}
		else if ((c == '/') || (c == '\\')) // reserved file name characters
		{
			if (allowUnderscore)
			{
				allowUnderscore = FALSE;
				result.AppendChar('_');
			}
		}
		else if (bReplaceWhiteWithUnderscore && ((c == ' ') || (c == '\t')))
		{
			if (allowUnderscore)
			{
				result.AppendChar('_');
				allowUnderscore = FALSE;
			}
		}
		else if (iscntrl(c) || (c == 0x7f))
		{
			if (allowUnderscore)
			{
				result.AppendChar('_');
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

void  FileUtils::MakeValidFileName(CString& name, CString& result, BOOL bReplaceWhiteWithUnderscore, BOOL extraValidation)
{
	FileUtils::MakeValidFileName((LPCWSTR)name, name.GetLength(), result, bReplaceWhiteWithUnderscore, extraValidation);
}

void  FileUtils::MakeValidFileName(const wchar_t *name, int namelen, CString &result, BOOL bReplaceWhiteWithUnderscore, BOOL extraValidation)
{
	// Always merge consecutive underscore characters
	result.Preallocate(namelen + 2);

	int csLen = namelen;
	int i;

	wint_t c;
	BOOL allowUnderscore = TRUE;
	for (i = 0; i < csLen; i++)
	{
		c = name[i] & 0xFFFF;
		if (
			(c == L'?') || (c == L'<') || (c == L'>') || (c == L':') ||  // not valid file name characters on Windows
			(c == L'*') || (c == L'|') || (c == L'"')                 // not valid file name characters
			|| (c == L'%')  // sometimes pdfbox, Edge and Chrome have problems
			)
		{
			if (allowUnderscore)
			{
				allowUnderscore = FALSE;
				result.AppendChar(L'_');
			}
		}
		else if (extraValidation &&
			((c == L'#') || (c == L'\'')   // sometimes pdfbox, Edge and Chrome have problems
				|| (c == L'!'))             // breaks cmd batch files, breaks Print Merge option
			)
		{
			if (allowUnderscore)
			{
				allowUnderscore = FALSE;
				result.AppendChar(L'_');
			}
		}
		else if ((c == L'/') || (c == L'\\')) // reserved file name characters
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
		else if (iswcntrl(c) || (c == 0x007f))
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

BOOL FileUtils::Write2File(CString &cStrNamePath, const unsigned char *data, int dataLength)
{
	DWORD dwAccess = GENERIC_WRITE;
	DWORD dwCreationDisposition = CREATE_ALWAYS;

	HANDLE hFile = CreateFileW(cStrNamePath, dwAccess, FILE_SHARE_READ, NULL, dwCreationDisposition, 0, NULL);
	if (hFile == INVALID_HANDLE_VALUE)
	{
		CString errText = FileUtils::GetLastErrorAsString();
		DWORD err = GetLastError();
		TRACE(L"(FileSize)INVALID_HANDLE_VALUE error=%ld file=%s\n%s\n", err, cStrNamePath, errText);
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
	const unsigned char* pszData = (unsigned char*)lpBuffer;
	DWORD nwritten = 0;
	while (bytesToWrite > 0)
	{
		nwritten = 0;
		if (!::WriteFile(hFile, pszData, bytesToWrite, &nwritten, NULL)) {
			DWORD retval = GetLastError();
			break;
		}
		pszData += nwritten;
		bytesToWrite -= nwritten;
	}
	*lpNumberOfBytesWritten = nNumberOfBytesToWrite - bytesToWrite;
	if (*lpNumberOfBytesWritten != nNumberOfBytesToWrite)
		return FALSE;
	else
		return TRUE;
}

BOOL FileUtils::ReadEntireFile(CString &cStrNamePath, SimpleString &txt)
{
	DWORD dwAccess = GENERIC_READ;
	DWORD dwCreationDisposition = OPEN_EXISTING;

	HANDLE hFile = CreateFile(cStrNamePath, dwAccess, FILE_SHARE_READ, NULL, dwCreationDisposition, 0, NULL);
	if (hFile == INVALID_HANDLE_VALUE)
	{
		int deb = 1;
		return FALSE;
	}
	else
	{
		LPOVERLAPPED lpOverlapped = 0;
		LPVOID       lpBuffer;
		DWORD        nNumberOfBytesToRead;
		DWORD      nNumberOfBytesRead;

		_int64 fsize = FileUtils::FileSize(cStrNamePath);
		_int64 bytesLeft = fsize;
		txt.ClearAndResize((int)fsize+1);
		lpBuffer = txt.Data();
		while (bytesLeft > 0)
		{
			if (bytesLeft > 0xffffffff)
				nNumberOfBytesToRead = 0xffffffff;
			else
				nNumberOfBytesToRead = (DWORD)bytesLeft;

			BOOL retval = ::ReadFile(hFile, lpBuffer, nNumberOfBytesToRead, &nNumberOfBytesRead, lpOverlapped);
			if (retval == FALSE)
				break;

			int newCount = txt.Count() + nNumberOfBytesRead;
			txt.SetCount(newCount);
			lpBuffer = txt.Data(newCount);
			bytesLeft -= nNumberOfBytesRead;
		}
		BOOL retClose = CloseHandle(hFile);
	}

	return TRUE;
}

BOOL FileUtils::NormalizeFilePath(CString &filePath)
{
	CString outFilePath;
	filePath.Replace(L"/", L"\\");
	filePath.TrimRight(L"\\");
	int i, k;
	wchar_t c;
	k = 0;
	for (i = 0; i < filePath.GetLength(); )
	{
		c = filePath.GetAt(i);
		outFilePath.AppendChar(c);
		i++;
		if (c == '\\')
		{
			for ( ; i < filePath.GetLength(); i++)
			{
				c = filePath.GetAt(i);
				if (c != '\\')
					break;
			}
		}
	}

	wchar_t buffer[MAX_PATH + 1];
	BOOL ret = PathCanonicalize(buffer, (LPCWSTR)outFilePath);
	if (ret)
	{
		filePath.Empty();
		filePath.Append(buffer);
		return TRUE;
	}
	else
		return FALSE;
}


// create all subfolders if they don't exist

BOOL FileUtils::CreateDir(const wchar_t *path)
{
	CList<CString, CString &> folderList;

	CString folderPath(path);
	CString subFolderPath(path);
	CString folderName;

	int i;
	BOOL foundValidSubFolder = FALSE;
	int maxSubfolderDepth = 128;
	for (i = 0; i < maxSubfolderDepth; i++)
	{
		if (FileUtils::PathDirExists(subFolderPath))
		{
			foundValidSubFolder = TRUE;
			break;
		}

		folderPath = subFolderPath;
		folderPath.TrimRight(L"\\");
		FileUtils::GetFolderPathAndFileName(folderPath, subFolderPath, folderName);
		if (subFolderPath.IsEmpty())
			break;
		folderList.AddHead(folderName);
	}
	if (!foundValidSubFolder)
	{
		return FALSE;
	}
	if (folderList.IsEmpty())
		return TRUE;

	folderPath = subFolderPath;
	folderPath.TrimRight(L"\\");
	folderPath.Append(L"\\");
	while (folderList.GetCount() > 0)
	{
		folderName = folderList.RemoveHead();
		folderPath.Append(folderName);
		folderPath.Append(L"\\");
		if (!::CreateDirectory(folderPath, NULL))  // must be global function
		{
			return FALSE;
		}
		if (!FileUtils::PathDirExists(folderPath))
		{
			return FALSE;
		}
	}
	return TRUE;
}

BOOL FileUtils::VerifyName(CString &name)
{
#if 0
	if ((name.Find(L"Cache") < 0) && // expand "Cache" check PrintCache etc
		(name.Find(L"UMBoxViewer") < 0) &&
		(name.Find(L".WriteSupported") < 0) &&  // should not be needed; MBoxViewer should fail first
		(name.Find(L".rootfolder") < 0) && // should not be needed; MBoxViewer should fail first
		(name.Find(L"AppData\\Local\\Temp\\UMBoxViewer") < 0)
		)
#else
	if (
		(name.Find(L"UMBoxViewer") < 0) &&
		(name.Find(L".WriteSupported") < 0)
		)
#endif
	{
		TRACE(L"VerifyName: Failed \"%s\"\n", name);
		_ASSERTE(FALSE);
		return FALSE;
	}
	return TRUE;
}

BOOL FileUtils::DelFile(CString &path, BOOL verify)
{
	BOOL ret = FileUtils::DelFile(path.operator LPCWSTR(), verify);
	return ret;
}

BOOL FileUtils::DelFile(const wchar_t *path, BOOL verify)
{

	TRACE(L"DelFile: \"%s\"\n", path);

	CStringW filePath = path;
	if (verify)
	{
		if (!VerifyName(filePath))
		{
			TRACE(L"DelFile: VerifyName failed. Failed to delete \"%s\"\n", path);
			_ASSERTE(path);
			return FALSE;
		}
	}

	BOOL ret = ::DeleteFile(path);
	if (!ret)
	{
		CStringW errorText = FileUtils::GetLastErrorAsString();
		TRACE(L"DelFile: \"%s\" %s", path, errorText);
	}
	return ret;
}

// Optimistic Copy, ignores many errors
// TODO: Should Copy make sure all files were copied and delete TO destination in case of any error ?
// Try to copy first and remove from directory only if no errors ??

CString FileUtils::CopyDirectory(const wchar_t* cFromPath, const wchar_t* cToPath, CStringArray * excludeFilter, BOOL bFailIfExists, BOOL removeFolderAfterCopy)
{
	WIN32_FIND_DATAW FileData;
	HANDLE hSearch;
	BOOL	bFinished = FALSE;
	CString errorText;

	CString fromPath = cFromPath;
	if (!FileUtils::PathDirExists(cFromPath))
	{
		errorText.Format(L"CopyDirectory: Invalid from directory path \"%s\"", fromPath);
		//errorText = FileUtils::GetLastErrorAsStringW();
		TRACE(L"%s", errorText);
		return errorText;
	}

	if (excludeFilter)
	{
		int count = (int)excludeFilter->GetCount();
		int i;
		for (i = 0; i < excludeFilter->GetCount(); i++)
		{
			if (fromPath.Find(excludeFilter->GetAt(i)) >= 0)
				break;
		}
		if (i == count)
			return errorText;
	}

	if (!FileUtils::VerifyName(fromPath))
	{
		TRACE(L"CopyDirectory: VerifyName failed. Invalid From Directory Name  \"%s\"\n", fromPath);
		_ASSERTE(FALSE);
		//return TRUE; // FIXME
	}

	CString toPath = cToPath;
	if (!FileUtils::CreateDir(cToPath))
	{
		errorText.Format(L"CopyDirectory: Create To directory \"%s\" failed.", toPath);
		//errorText = FileUtils::GetLastErrorAsStringW();
		TRACE(L"%s", errorText);
		return errorText;
	}

	if (!FileUtils::VerifyName(toPath))
	{
		TRACE(L"CopyDirectory: VerifyName failed. Invalid To Directory Name \"%s\"\n", toPath);
		_ASSERTE(FALSE);
		//return TRUE; // FIXME
	}

	// Start searching for all files in the from directory.
	CString searchPath = fromPath + L"*.*";
	hSearch = FindFirstFileW(searchPath, &FileData);
	if (hSearch == INVALID_HANDLE_VALUE)
	{
		CString errText;
		errText.Format(L"CopyDirectory: Didn't find files in \"%s\"", fromPath);
		//CStringW errText = FileUtils::GetLastErrorAsStringW();
		TRACE(L"%s", errText);
		return errorText;   // Success, return empty error string
	}

	while (!bFinished)
	{
		if (!(wcscmp(FileData.cFileName, L".") == 0 || wcscmp(FileData.cFileName, L"..") == 0))
		{
			CString fileFound = fromPath + FileData.cFileName;
			CString toFilePath = toPath + FileData.cFileName;
			if (FileData.dwFileAttributes != FILE_ATTRIBUTE_DIRECTORY)
			{
				if (!::CopyFileW(fileFound, toFilePath, bFailIfExists))
				{
					CString errText = FileUtils::GetLastErrorAsString();
					TRACE(L"CopyDirectory: \"%s\"\n%s", fileFound, errText);
				}
				else if (removeFolderAfterCopy)
				{
					if (!FileUtils::DelFile(fileFound))
					{
						CString errText = FileUtils::GetLastErrorAsString();
						TRACE(L"CopyDirectory: \"%s\"\n%s", fileFound, errText);
					}
				}
			}
			else // recursive
			{
				CString fromPath = fileFound + L"\\";
				CString toPath = toFilePath + L"\\";
				errorText = FileUtils::CopyDirectory(fromPath, toPath, excludeFilter, bFailIfExists, removeFolderAfterCopy);
				if (!errorText.IsEmpty())
				{
					TRACE(L"CopyDirectory: \"%s\"\n%s", fromPath, errorText);
				}
			}
		}
		if (!FindNextFileW(hSearch, &FileData))
		{
			bFinished = TRUE;
		}
	}

	FindClose(hSearch);
	if (removeFolderAfterCopy)
	{
		errorText.Empty();
		BOOL retRD = FileUtils::OSRemoveDirectory(fromPath, errorText);
		if (!retRD)
		{
			TRACE(L"CopyDirectory: \"%s\"\n%s", fromPath, errorText);
		}
	}
	return errorText;
}

CString FileUtils::MoveDirectory(const wchar_t *cFromPath, const wchar_t *cToPath)
{
	BOOL removeFolderAfterCopy = TRUE;
	BOOL bFailIfExists = FALSE;
	CString errorText = FileUtils::CopyDirectory(cFromPath, cToPath, 0, bFailIfExists, removeFolderAfterCopy);
	return errorText;
}

CString FileUtils::GetLastErrorAsString(DWORD errorCode)
{
	//Get the error message, if any.
	CStringW emptyMsg;
	DWORD errorMessageID = errorCode;
	if (errorCode == 0xFFFF)
		errorMessageID = ::GetLastError();

	//DWORD errorMessageID = ERROR_REQ_NOT_ACCEP;
	if (errorMessageID == 0)
		return emptyMsg; //No error message has been recorded

	LPWSTR messageBuffer = nullptr;
	DWORD size = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
		NULL, errorMessageID, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPWSTR)&messageBuffer, 0, NULL);

	CStringW message(messageBuffer, size);

	CStringW errorMessage;
	message.TrimRight(L" \t\r\n");
	errorMessage.Format(L"Error Code %u \"%s\"\n", errorMessageID, message);

	//Free the buffer.
	LocalFree(messageBuffer);

	return errorMessage;
}

CString FileUtils::GetFileExceptionErrorAsString(CFileException &exError)
{
	CString exErrorStr;
	wchar_t szCause[2048];
	exError.GetErrorMessage(szCause, 2048);

	exErrorStr.Append(szCause);
	return exErrorStr;
}

CString FileUtils::GetOpenForReadFileExceptionErrorAsString(CString &fileName, CFileException &exError)
{
	CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(exError);

	CString txt = L"Could not open \"" + fileName;
	txt += L"\" file.\n";
	txt += exErrorStr;
	return txt;
}

BOOL FileUtils::GetFolderList(CString &rootFolder, CList<CString, CString &> &folderList, CString &errorText, int maxDepth)
{
	// TODO: Hash Table ??
	CArray<CString> cacheFolderList;
	cacheFolderList.Add("PrintCache");
	cacheFolderList.Add("ImageCache");
	cacheFolderList.Add("AttachmentCache");
	cacheFolderList.Add("EmlCache");
	cacheFolderList.Add("LabelCache");
	cacheFolderList.Add("ArchiveCache");
	cacheFolderList.Add("ListCache");
	cacheFolderList.Add("MergeCache");
	//NTreeView::SetupCacheFolderList(cacheFolderList);  // should use this but it would need to include NTreeView.h

	rootFolder.TrimRight(L"\\");
	//rootFolder.Append("\\");

	CString folderName;
	FileUtils::CPathStripPath(rootFolder, folderName);
	for (int i = 0; i < cacheFolderList.GetCount(); i++)
	{
		CString &cacheFolderName = cacheFolderList[i];
		if (folderName.Compare(cacheFolderName) == 0)
		{
			return TRUE;
		}
	}

	CString folderPath = rootFolder;
	WIN32_FIND_DATA FileData;
	HANDLE hSearch;
	BOOL	bFinished = FALSE;
	int fileCnt = 0;

	// Start searching for all folders in the current directory.
	CString searchPath = folderPath + "\\*.*";
	hSearch = FindFirstFile(searchPath, &FileData);
	if (hSearch == INVALID_HANDLE_VALUE) {
		TRACE(L"GetFolderList: No files found.\n");
		return FALSE;
	}


	//CString folderName;
	CString fileName;
	while (!bFinished)
	{
		if (!(wcscmp(FileData.cFileName, L".") == 0 || wcscmp(FileData.cFileName, L"..") == 0))
		{
			fileName = CString(FileData.cFileName);
			CString fileFound = folderPath + "\\" + fileName;

			if (FileData.dwFileAttributes == FILE_ATTRIBUTE_DIRECTORY)
			{
				fileFound.TrimRight(L"\\");
				fileFound.Append(L"\\");
#if 0
				FileUtils::CPathStripPath(folderPath, folderName);
				ignoreFolder = FALSE;
				for (int i = 0; i < cacheFolderList.GetCount(); i++)
				{
					CString &cacheFolderName = cacheFolderList[i];
					if (fileName.Compare(cacheFolderName) == 0)
					{
						ignoreFolder = TRUE;
						break;
					}
				}
				if (!ignoreFolder)
#endif
				{
					//folderList.AddTail(fileFound);
					BOOL bRetval = FileUtils::GetFolderList(fileFound, folderList, errorText, maxDepth - 1);
				}
			}
			else
				fileCnt++;
		}
		if (!FindNextFile(hSearch, &FileData)) {
			bFinished = TRUE;
			break;
		}
	}
	FindClose(hSearch);

	if (fileCnt)
	{
		rootFolder.TrimRight(L"\\");
		rootFolder.Append(L"\\");
		folderList.AddTail(rootFolder);
	}

	return TRUE;
}

int FileUtils::GetFolderFileCount(CString &rootFolderPath, BOOL recursive)
{
	int fileCnt = 0;

	CString folderPath = rootFolderPath;
	folderPath.TrimRight(L"\\");

	WIN32_FIND_DATA FileData;
	HANDLE hSearch;
	BOOL	bFinished = FALSE;

	// Start searching for all folders in the current directory.
	CString searchPath = folderPath + L"\\*.*";
	hSearch = FindFirstFile(searchPath, &FileData);
	if (hSearch == INVALID_HANDLE_VALUE) {
		return 0;
	}

	CString fileName;
	while (!bFinished)
	{
		if (!(_tcscmp(FileData.cFileName, L".") == 0 || _tcscmp(FileData.cFileName, L"..") == 0))
		{
			fileName = CString(FileData.cFileName);
			CString fileFound = folderPath + L"\\" + fileName;

			if (FileData.dwFileAttributes == FILE_ATTRIBUTE_DIRECTORY)
			{
				if (recursive == TRUE)
				{
					BOOL cnt = FileUtils::GetFolderFileCount(fileFound, recursive);
					fileCnt += cnt;
				}
			}
			else
			{
				fileCnt++;
			}
		}

		if (!FindNextFile(hSearch, &FileData)) {
			bFinished = TRUE;
			break;
		}
	}
	FindClose(hSearch);

	return fileCnt;
}

_int64 FileUtils::FolderSize(LPCWSTR fPath)
{
	WIN32_FIND_DATA data;
	__int64 size = 0;
	CString folderPath = fPath;
	folderPath.TrimRight(L"\\");
	folderPath.Append(L"\\)");
	CString filePath = folderPath + L"*.*";
	HANDLE h = FindFirstFile(filePath, &data);
	if (h != INVALID_HANDLE_VALUE)
	{
		do 
		{
			if ((data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
			{
				if (_tcscmp(data.cFileName, L".") != 0 && _tcscmp(data.cFileName, L"..") != 0)
				{
					folderPath = folderPath + data.cFileName;
					size += FolderSize(filePath);
				}

			}
			else
			{
				LARGE_INTEGER sz;
				sz.LowPart = data.nFileSizeLow;
				sz.HighPart = data.nFileSizeHigh;
				size += sz.QuadPart;
			}
		} while (FindNextFile(h, &data) != 0);
		FindClose(h);

	}
	return size;
}

BOOL FileUtils::IsReadonlyFolder(CString &folderPath)
{
	CString path = folderPath;
	path.TrimRight(L"\\");

	CString tempFolderPath = path + L"\\.WriteSupported";

	BOOL retD = FileUtils::RemoveDir(tempFolderPath, true, true);
	if (FileUtils::PathDirExists(tempFolderPath))
	{
		// Assume readonly folder
		return TRUE;
	}

	if (!FileUtils::CreateDir(tempFolderPath))
	{
		// Should never fails
		// Assume readonly folder
		return TRUE;
	}

	retD = FileUtils::RemoveDir(tempFolderPath, true, true);

	return FALSE;
}

void FileUtils::SplitFileFolder(CString& path, CStringArray& a)
{
	CString driveName;
	CString directory;
	CString fileNameBase;
	CString fileNameExtention;

	SplitFilePath(path, driveName, directory, fileNameBase, fileNameExtention);

	int position = 0;
	CString delim = L"\\";
	CString strToken;

	a.RemoveAll();
	directory.TrimRight(L"\\");
	strToken = directory.Tokenize(delim, position);
	while (!strToken.IsEmpty())
	{
		a.Add(strToken);
		strToken = directory.Tokenize(delim, position);
	}
}

CString FileUtils::SizeToString(_int64 size)
{
	//wchar_t sizeStr_inKB[256];
	wchar_t sizeStr_inBytes[256];
	int sizeStrSize = 256;

#if 0
	_int64 size;
	LPCWSTR fileSizeStr_inKB = StrFormatKBSize(fileSize, &sizeStr_inKB[0], sizeStrSize);
	if (!fileSizeStr_inKB)
		sizeStr_inKB[0] = 0;
#endif

	LPCWSTR fileSizeStr_inBytes = StrFormatByteSize64(size, &sizeStr_inBytes[0], sizeStrSize);
	if (!fileSizeStr_inBytes)
		sizeStr_inBytes[0] = 0;

	CString txt;
	//txt.Format(L"File size:  %s (%s)\n", sizeStr_inKB, sizeStr_inBytes);
	txt.Format(L"File size: (%s)", sizeStr_inBytes);

	return txt;
}

// TODO: update test to test chnages to RemoveDirA, new CopyDirectory and MoveDirectory
void FileUtils::UnitTest()
{
#if 0
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
	pathW = GetMboxviewTempPathW();
	if (pathW.Find(safeDirW) < 0)
		_ASSERTE(false);
	retRemoveFiles = RemoveDirW(pathW, true);
	retRemoveDirectory = RemoveDirectoryW(pathW, error);

	pathW = GetMboxviewTempPathW();
	retPathExists = PathDirExistsW(pathW);
	_ASSERTE(retPathExists);
	retRemoveFiles = RemoveDirW(pathW, true);
	retRemoveDirectory = RemoveDirectoryW(pathW, error);
	_ASSERTE(retRemoveDirectory);

	pathW = GetMboxviewTempPathW();
	retPathExists = PathDirExistsW(pathW);
	subpathW = GetMboxviewTempPathW(L"Inline");
	retPathExists = PathDirExistsW(subpathW);
	_ASSERTE(retPathExists);
	retRemoveFiles = RemoveDirW(subpathW, true);
	//retRemoveDirectory = RemoveDirectoryW(subpathW, error);
	retRemoveDirectory = RemoveDirectoryW(pathW, error);
	_ASSERTE(retRemoveDirectory);

	/////////////////////////////////////////

	/////////////////////////////////
	pathA = GetMboxviewTempPath();
	if (pathA.Find(safeDir) < 0)
		_ASSERTE(false);
	retRemoveFiles = RemoveDirA(pathA, true);
	retRemoveDirectory = RemoveDirectoryA(pathA, error);

	pathA = GetMboxviewTempPath();
	retPathExists = PathDirExists(pathA);
	_ASSERTE(retPathExists);
	retRemoveFiles = RemoveDirA(pathA, true);
	retRemoveDirectory = RemoveDirectoryA(pathA, error);
	_ASSERTE(retRemoveDirectory);


	pathA = GetMboxviewTempPath();
	retPathExists = PathDirExists(pathA);
	subpathA = GetMboxviewTempPath("Inline");
	retPathExists = PathDirExists(subpathA);
	_ASSERTE(retPathExists);
	//retRemoveDirectory = RemoveDirectoryA(subpathA, error);
	retRemoveDirectory = RemoveDirectoryA(pathA, error);
	_ASSERTE(retRemoveDirectory);
#endif
}
