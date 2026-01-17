/*
* Cfbfinfo is written by Graeme Cole
* https://github.com/elocemearg/cfbfinfo
*
* Ported by Zbigniew Minciel to Windows to integrate with free Windows MBox Viewer
*
* Note: For simplification, the code assumes that the target system is little-endian.
*/

#pragma warning(disable : 4244)  // warning C4244 : 'argument' : conversion from '__int64' to 'size_t', possible loss of data

#include "ReportMemoryLeaks.h"

#include "stdafx.h"
#include <stdlib.h>
#include <stdio.h>
#include <string.h>

#ifndef _WINDOWS
#include <sys/mman.h>
#include <sys/types.h>
#include <sys/stat.h>
#include <fcntl.h>
#include <unistd.h>
#include <errno.h>
#include <error.h>
#endif

#include "cfbf.h"

void
cfbf_close(struct cfbf* cfbf) {
	cfbf_fat_close(&cfbf->fat);
	cfbf_fat_close(&cfbf->mini_fat);
	free(cfbf->mini_stream);

#ifdef _WINDOWS
	free(cfbf->filepath);
	free(cfbf->file);
	_close(cfbf->fd);
#else
	munmap(cfbf->file, cfbf->file_size);
	close(cfbf->fd);
#endif
}

int
cfbf_open(const wchar_t* wfilename, struct cfbf* cfbf, std::string& errorText)
{
	struct _stat st;
	DirEntry* root;

	if (wfilename == 0) {
		return -1;
	}

	if (cfbf == 0) {
		return -1;
	}

	// ZMM Fix int casting
	int wfilename_len = (int)wcslen(wfilename);
	std::string filename_str8 = UTF16ToUTF8((wchar_t*)wfilename, wfilename_len * 2);
	const char* filename = filename_str8.c_str();

	memset(cfbf, 0, sizeof(*cfbf));

	cfbf->out = 0;
	cfbf->printEnabled = false;

	cfbf->filepath = (wchar_t*)cfbf_malloc(2 * (wcslen(wfilename) + 1));
	if (cfbf->filepath == NULL) {
		char txt[512];
		//error(0, errno, "%s", filename);
		sprintf(txt, "%s", filename);
		errorText = txt;
		return -1;
	}

	wcscpy(cfbf->filepath, wfilename);

	if (_wstat(wfilename, &st) < 0) {
		char txt[512];
		//error(0, errno, "%s", filename);
		sprintf(txt, "%s", filename);
		errorText = txt;
		goto fail;
	}

#ifdef _WINDOWS
	cfbf->fd = _wopen(wfilename, _O_BINARY | _O_RDONLY);
#else
	cfbf->fd = open(filename, O_RDONLY);
#endif
	if (cfbf->fd < 0) {
		char txt[512];
		//error(0, errno, "%s", filename);
		sprintf(txt, "%s", filename);
		errorText = txt;
		goto fail;
	}

	cfbf->file_size = st.st_size;

	if (cfbf->file_size < sizeof(struct StructuredStorageHeader))
	{
		char txt[512];
		sprintf(txt, "%s is too small (%lld bytes) to contain a StructuredStorageHeader (%d bytes)",
			filename, cfbf->file_size, (int)sizeof(struct StructuredStorageHeader));
		errorText = txt;
		goto fail;
	}

#ifdef _WINDOWS
	cfbf->file = cfbf_malloc(cfbf->file_size);
	if (cfbf->file == NULL) {
		char txt[512];
		//error(0, errno, "failed to allocate memory for %s mail file", filename);
		sprintf(txt, "failed to allocate memory for %s mail file", filename);
		errorText = txt;
		goto fail;
	}
	memset(cfbf->file, 0, cfbf->file_size);
	// ZMM Fix int casting
	int bytes_to_read = (int)cfbf->file_size;
	int bytes_read_cnt = 0;
	int len = 0;
	while (bytes_to_read > 0)
	{
		len = _read(cfbf->fd, cfbf->file, bytes_to_read);
		if (len <= 0)
			break;
		bytes_read_cnt += len;
		bytes_to_read -= len;
	}

	if (bytes_read_cnt != cfbf->file_size) {
		char txt[512];
		//error(0, errno, "failed to read %d bytes from %s mail file", cfbf->file_size, filename);
		sprintf(txt, "failed to read %lld bytes from %s mail file", cfbf->file_size, filename);
		errorText = txt;
		goto fail;
	}

#else
	cfbf->file = mmap(NULL, cfbf->file_size, PROT_READ, MAP_SHARED, cfbf->fd, 0);
	if (cfbf->file == NULL) {
		error(0, errno, "failed to mmap %s", filename);
		goto fail;
	}
#endif
	cfbf->header = (struct StructuredStorageHeader*)cfbf->file;

	if (memcmp(cfbf->header->_abSig, "\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1", 8)) {
		char txt[512];
		//error(0, errno, "%s: signature bytes not as expected - this doesn't look like a CFB file", filename);
		sprintf(txt, "%s: signature bytes not as expected - this doesn't look like a CFB file", filename);
		errorText = txt;
		goto fail;
	}

	unsigned long num_start_sectors;
	unsigned long num_fat_sectors = (unsigned long)cfbf->header->_csectFat;

	if (num_fat_sectors > 109)
		num_start_sectors = 109;
	else
		num_start_sectors = num_fat_sectors;

	if (cfbf_fat_open(&cfbf->fat, cfbf, cfbf->header->_sectFat,
		num_start_sectors, cfbf->header->_sectDifStart,
		(unsigned long)cfbf->header->_csectDif,
		num_fat_sectors, errorText) < 0)
	{
		char txt[512];
		sprintf(txt, "%s: failed to load FAT", filename);
		errorText = txt;
		memset(&cfbf->fat, 0, sizeof(cfbf->fat));
		goto fail;
	}

	if (cfbf_mini_fat_open(&cfbf->mini_fat, &cfbf->fat, cfbf,
		cfbf->header->_sectMiniFatStart, cfbf->header->_csectMiniFat, errorText) < 0)
	{
		char txt[512];
		sprintf(txt, "%s: failed to load mini-FAT", filename);
		errorText = txt;
		memset(&cfbf->mini_fat, 0, sizeof(cfbf->mini_fat));
		goto fail;
	}

	/* Load mini-stream - the start sector and length of this is given by the
	 * RootEntry */
	root = (DirEntry*)cfbf_get_sector_ptr(cfbf, cfbf->header->_sectDirStart, errorText);
	if (root == NULL) {
		char txt[512];
		sprintf(txt, "%s: failed to look up root entry", filename);
		errorText = txt;
		goto fail;
	}

	if (memcmp(root->name, "R\0o\0o\0t\0 \0E\0n\0t\0r\0y\0\0\0", 22)) {
		char txt[512];
		sprintf(txt, "%s: first directory entry is not called RootEntry", filename);
		errorText = txt;
		goto fail;
	}

	cfbf->mini_stream = cfbf_alloc_chain_contents_from_fat(cfbf, root->start_sector, root->stream_size, errorText);
	if (cfbf->mini_stream == NULL) {
		char txt[512];
		sprintf(txt, "%s: failed to load mini-stream", filename);
		errorText = txt;
		goto fail;
	}
	cfbf->mini_stream_size = root->stream_size;

	return 0;

fail:
	cfbf_close(cfbf);
	return -1;
}

int
cfbf_get_sector_size(struct cfbf* cfbf) {
	return 1 << cfbf->header->_uSectorShift;
}

int
cfbf_get_mini_fat_sector_size(struct cfbf* cfbf) {
	return 1 << cfbf->header->_uMiniSectorShift;
}

int
cfbf_read_sector(struct cfbf* cfbf, SECT sect, void* dest, std::string& errorText) {
	int sect_size = cfbf_get_sector_size(cfbf);
	void* src = cfbf_get_sector_ptr(cfbf, sect, errorText);
	if (src == NULL)
		return -1;

	memcpy(dest, src, sect_size);

	return 0;
}

void*
cfbf_get_sector_ptr(struct cfbf* cfbf, SECT sect, std::string& errorText)
{
	int sect_size = cfbf_get_sector_size(cfbf);
	unsigned long long offset = (sect + 1) * sect_size;

	if ((long long)offset >= cfbf->file_size) {
		char txt[512];
		sprintf(txt, "can't get sector %lu - it's past the end of the file (file size %lld, sector size %d)", (unsigned long)sect, cfbf->file_size, sect_size);
		errorText = txt;
		return NULL;
	}

	return (BYTE*)cfbf->file + offset;
}

void*
cfbf_get_sector_ptr_in_mini_stream(struct cfbf* cfbf, SECT sector) {
	int mini_sector_size = cfbf_get_mini_fat_sector_size(cfbf);
	uint64_t offset = sector * mini_sector_size;
	if (offset > cfbf->mini_stream_size) {
		return NULL;
	}
	return (void*)((char*)cfbf->mini_stream + offset);
}


int
cfbf_is_sector_in_file(struct cfbf* cfbf, SECT sect) {
	int sect_size = cfbf_get_sector_size(cfbf);
	if ((sect + 1) * sect_size + sect_size > cfbf->file_size)
		return 0;
	else
		return 1;
}
