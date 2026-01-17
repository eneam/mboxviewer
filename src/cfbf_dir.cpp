/*
* Cfbfinfo is written by Graeme Cole
* https://github.com/elocemearg/cfbfinfo
*
* Ported by Zbigniew Minciel to Windows to integrate with free Windows MBox Viewer
*
* Note: For simplification, the code assumes that the target system is little-endian.
*/

// Fix vars definitions
#pragma warning(disable : 4267)
#pragma warning(disable : 4244)

#include "ReportMemoryLeaks.h"


#include "stdafx.h"
#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <errno.h>
#ifndef _WINDOWS
#include <error.h>
#include <iconv.h>
#endif

#include "cfbf.h"
#include "outlook.h"

static size_t
strlen_utf16(uint16_t* s) {
	size_t count = 0;
	while (*s) {
		count++;
		s++;
	}
	return count;
}

static int
strncmp_utf16(const uint16_t* s1, const uint16_t* s2, size_t n)
{
	for (size_t i = 0; i < n; ++i) {
		if (s1[i] < s2[i])
			return -1;
		else if (s1[i] > s2[i])
			return 1;
		else if (s1[i] == 0) {
			// found the end of the string
			return 0;
		}
	}
	return 0;
}

static uint16_t*
strchr_utf16(const uint16_t* s, int c)
{
	while (*s) {
		if (*s == c)
			return (uint16_t*)s;
		++s;
	}
	return NULL;
}

static DirEntry*
cfbf_find_path_in_tree(struct cfbf* cfbf, void** dir_chain,
	int dir_chain_length, int sector_size, int entries_per_sector,
	unsigned long entry_id, uint16_t* sought_path_utf16, std::string& errorText)
{
	uint16_t* path_component_end;
	int path_component_length;
	int last_component = 0;
	SECT sector = entry_id / entries_per_sector;
	int entry_within_sector = entry_id % entries_per_sector;
	DirEntry* e;

	const int sought_path_utf16_len = strlen_utf16(sought_path_utf16);
	size_t out_len = 0;
	char* buff8 = UTF16ToUTF8((wchar_t*)sought_path_utf16, sought_path_utf16_len, &out_len);
	free(buff8);

	if (entry_id == CFBF_NOSTREAM)
		return NULL;

	e = ((DirEntry*)dir_chain[sector]) + entry_within_sector;

	out_len = 0;
	char* name_buff8 = UTF16ToUTF8((wchar_t*)e->name, e->name_length, &out_len);
	free(name_buff8);

	std::string name_utf8 = UTF16ToUTF8((wchar_t*)e->name, e->name_length);

	path_component_end = strchr_utf16(sought_path_utf16, '/');
	if (path_component_end == NULL) {
		path_component_end = sought_path_utf16 + strlen_utf16(sought_path_utf16);
		last_component = 1;
	}

	path_component_length = path_component_end - sought_path_utf16;

	if (e->object_type != CFBF_STORAGE_OBJECT_TYPE &&
		e->object_type != CFBF_STREAM_OBJECT_TYPE &&
		e->object_type != CFBF_ROOT_OBJECT_TYPE)
	{
		if (e->object_type != CFBF_UNKNOWN_OBJECT_TYPE) {
			char txt[512];
			sprintf(txt, "cfbf_find_path_in_tree(): entry ID %lu has invalid object type %d, skipping this entry", entry_id, (int)e->object_type);
			// errorText = txt; ?? ZMMM
		}
		return NULL;
	}
	else
	{
		/* If the name of this node matches sought_path_utf16 up to the first
		 * slash, descend into this node's children. If there is no slash
		 * in sought_path_utf16 and it matches the name of this entry,
		 * return this entry. */
		if ((e->name_length - 1) / 2 == path_component_length &&
			(strncmp_utf16(sought_path_utf16, e->name, path_component_length) == 0))
		{
			if (last_component) {
				/* We've found the entry referred to by the path. */
				return e;
			}
			else {
				return cfbf_find_path_in_tree(cfbf, dir_chain, dir_chain_length,
					sector_size, entries_per_sector, e->child_id,
					path_component_end + 1, errorText);
			}
		}
		else
		{
			/* This entry doesn't match, so try its siblings. */
			DirEntry* found;

			found = cfbf_find_path_in_tree(cfbf, dir_chain, dir_chain_length,
				sector_size, entries_per_sector, e->left_sibling_id,
				sought_path_utf16, errorText);
			if (!found) {
				found = cfbf_find_path_in_tree(cfbf, dir_chain,
					dir_chain_length, sector_size, entries_per_sector,
					e->right_sibling_id, sought_path_utf16, errorText);
			}
			return found;
		}
	}
}

DirEntry*
cfbf_dir_entry_find_path(struct cfbf* cfbf, char* sought_path_utf8, std::string& errorText)
{
	void** dir_chain = NULL;
	int num_dir_secs;
	DirEntry* entry = NULL;
	uint16_t* sought_path_utf16 = NULL;
	//int sought_path_utf16_max;  // ZMM check original file

	dir_chain = cfbf_get_chain_ptrs(cfbf, cfbf->header->_sectDirStart, &num_dir_secs, errorText);

	if (dir_chain == NULL) {
		char txt[512];
		sprintf(txt, "failed to read directory chain");
		errorText = txt;
		return NULL;
	}

	/* Convert sought path to UTF-16 */
	const int sought_path_utf8_len = strlen(sought_path_utf8);
	size_t out_len = 0;
	wchar_t* buff16 = UTF8ToUTF16((char*)sought_path_utf8, sought_path_utf8_len, &out_len);
	if (buff16 == 0)
		goto nomem;

	sought_path_utf16 = (uint16_t*)buff16;

	entry = cfbf_find_path_in_tree(cfbf, dir_chain, num_dir_secs,
		cfbf_get_sector_size(cfbf),
		cfbf_get_sector_size(cfbf) / sizeof(DirEntry), 0,
		sought_path_utf16, errorText);

end:
	free(dir_chain);
	free(sought_path_utf16);

	return entry;

nomem:
	entry = NULL;
	goto end;
}

DirEntry* GetDirEntry(void** dir_chain, int dir_chain_length, unsigned long entry_id, int entries_per_sector, std::string& errorText)
{
	SECT sector = entry_id / entries_per_sector;
	int entry_within_sector = entry_id % entries_per_sector;
	DirEntry* entry;

	// ZMM Fix int casting
	if ((int)sector >= dir_chain_length) {
		char txt[512];
		sprintf(txt, "cfbf_walk_dir_tree_from_chain(): directory entry id %lu not in chain", (unsigned long)entry_id);
		errorText = txt;
		return 0;
	}

	entry = &((DirEntry*)dir_chain[sector])[entry_within_sector];

	return entry;
}

static int
cfbf_walk_dir_tree_from_chain(struct cfbf* cfbf, void** dir_chain,
	int dir_chain_length, int sector_size, int entries_per_sector,
	unsigned long entry_id, DirEntry* parent, int depth,
	int (*callback)(void*, struct cfbf*, DirEntry*,
		DirEntry*, unsigned long, int), void* cookie, std::string& errorText)
{
	DirEntry* entry;
	int ret;

	SECT sector = entry_id / entries_per_sector;
	int entry_within_sector = entry_id % entries_per_sector;


	// ZMM Fix int casting
	if ((int)sector >= dir_chain_length) {
		char txt[512];
		sprintf(txt, "cfbf_walk_dir_tree_from_chain(): directory entry id %lu not in chain", (unsigned long)entry_id);
		errorText = txt;
		return -1;
	}

	entry = &((DirEntry*)dir_chain[sector])[entry_within_sector];

	std::string parent_name_utf8;
	std::string entry_name_utf8 = UTF16ToUTF8((wchar_t*)entry->name, entry->name_length);
	if (parent)
		parent_name_utf8 = UTF16ToUTF8((wchar_t*)parent->name, parent->name_length);

	int deb1 = 1;

	std::string name_utf8 = UTF16ToUTF8((wchar_t*)entry->name, entry->name_length);

	if (name_utf8.compare("__recip_version1.0_#00000000") == 0)
		int deb = 0; // return 0;

	ret = callback(cookie, cfbf, entry, parent, entry_id, depth);
	if (ret == 0) {
		/* Give up but don't fail */
		return 0;
	}
	else if (ret < 0) {
		/* Give up and fail */
		return ret;
	}

	if (entry->object_type == CFBF_STORAGE_OBJECT_TYPE)   // 0x1
		int deb = 1;
	if (entry->object_type == CFBF_STREAM_OBJECT_TYPE) {   // 0x2
		int deb = 1;
	}
	if (entry->object_type == CFBF_ROOT_OBJECT_TYPE) {   // 0x5
		int deb = 1;
	}

	/* Visit children */
	if (entry->child_id != CFBF_NOSTREAM) {
		ret = cfbf_walk_dir_tree_from_chain(cfbf, dir_chain, dir_chain_length,
			sector_size, entries_per_sector, entry->child_id, entry,
			depth + 1, callback, cookie, errorText);
		if (ret <= 0)
			return ret;
	}

	/* Visit left siblings, and visit right siblings. The sibling links seem
	 * to make a tree rather than a doubly-linked list, so do it recursively */
	if (entry->left_sibling_id != CFBF_NOSTREAM) {
		ret = cfbf_walk_dir_tree_from_chain(cfbf, dir_chain, dir_chain_length,
			sector_size, entries_per_sector, entry->left_sibling_id, parent,
			depth, callback, cookie, errorText);
		if (ret <= 0)
			return ret;
	}
	if (entry->right_sibling_id != CFBF_NOSTREAM) {
		ret = cfbf_walk_dir_tree_from_chain(cfbf, dir_chain, dir_chain_length,
			sector_size, entries_per_sector, entry->right_sibling_id, parent,
			depth, callback, cookie, errorText);
		if (ret <= 0)
			return ret;
	}

	if (parent)
	{
		if (name_utf8.compare("__substg1.0_3701000D") == 0)
		{
			//fprintf(stdout, "Attach Object Type END !!!\n");
			OutlookMsgHelper* msgHelper = (OutlookMsgHelper*)cookie;
			OutlookMessage* active_msg = msgHelper->active_msg;
			_ASSERTE(active_msg);
			_ASSERTE(active_msg->emlFile);
			_ASSERTE(active_msg->m_cfbf);
			_ASSERTE(active_msg->m_cfbf->filepath);


			_ASSERTE(!msgHelper->m_msgList.empty());
			_ASSERTE(msgHelper->m_msgList.size() >= 1);
			if (!msgHelper->m_msgList.empty())
			{
				msgHelper->active_msg = msgHelper->m_msgList.back();
				msgHelper->m_msgList.pop_back();
			}
			else
			{
				_ASSERTE(msgHelper->active_msg == &msgHelper->msg);
			}

			AttachmentInfo_* attachInfo = msgHelper->active_msg->FindAttach(parent_name_utf8);
			if (attachInfo)
			{
				attachInfo->m_attach.m_OutlookMessage = active_msg;
			}

			int deb = 1;
		}
	}
	return 1;
}

/* callback should return positive if all is well and to continue the walk,
 * zero to terminate the walk but not fail, and a negative number to terminate
 * the walk and fail.
 * If the callback ever returns zero or negative, cfbf_walk_dir_tree() returns
 * that value. Otherwise we return 1. */
int
cfbf_walk_dir_tree(struct cfbf* cfbf,
	int (*callback)(void* cookie, struct cfbf* cfbf, DirEntry* entry,
		DirEntry* parent, unsigned long entry_id, int depth),
	void* cookie, std::string& errorText)
{
	int ret;
	int num_dir_secs;
	void** dir_chain;

	dir_chain = cfbf_get_chain_ptrs(cfbf, cfbf->header->_sectDirStart, &num_dir_secs, errorText);
	if (dir_chain == NULL) {
		char txt[512];
		sprintf(txt, "failed to read directory chain");
		errorText = txt;
		return -1;
	}

	ret = cfbf_walk_dir_tree_from_chain(cfbf, dir_chain, num_dir_secs,
		cfbf_get_sector_size(cfbf),
		cfbf_get_sector_size(cfbf) / sizeof(DirEntry),
		0, NULL, 0, callback, cookie, errorText);

	free(dir_chain);

	return ret;
}

void
cfbf_object_type_to_string(int object_type, char* dest, int dest_max)
{
	switch (object_type)
	{
	case 0:
		strncpy(dest, "unused", dest_max);
		break;

	case 1:
		strncpy(dest, "[storage]", dest_max);
		break;

	case 2:
		strncpy(dest, "stream", dest_max);
		break;

	case 5:
		strncpy(dest, "root", dest_max);
		break;

	default:
		snprintf(dest, dest_max, "%02X", (unsigned char)object_type);
	}
	if (dest_max > 0)
		dest[dest_max - 1] = '\0';
}

int
cfbf_dir_stored_in_mini_stream(struct cfbf* cfbf, DirEntry* entry)
{
	if (entry->object_type == CFBF_STREAM_OBJECT_TYPE && entry->stream_size > 0 &&
		entry->stream_size < cfbf->header->_ulMiniSectorCutoff) {
		return 1;
	}
	else {
		return 0;
	}
}

void**
cfbf_dir_entry_get_sector_ptrs(struct cfbf* cfbf, DirEntry* entry, int* num_sectors_r, int* sector_size_r, std::string& errorText)
{
	if (cfbf_dir_stored_in_mini_stream(cfbf, entry)) {
		if (sector_size_r)
			*sector_size_r = cfbf_get_mini_fat_sector_size(cfbf);
		return cfbf_get_chain_ptrs_from_mini_stream(cfbf, entry->start_sector, num_sectors_r, errorText);
	}
	else {
		if (sector_size_r)
			*sector_size_r = cfbf_get_sector_size(cfbf);
		return cfbf_get_chain_ptrs(cfbf, entry->start_sector, num_sectors_r, errorText);
	}
}
