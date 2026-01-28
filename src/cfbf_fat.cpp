/*
* Cfbfinfo is written by Graeme Cole
* https://github.com/elocemearg/cfbfinfo
*
* Ported by Zbigniew Minciel to Windows to integrate with free Windows MBox Viewer
*
* Note: For simplification, the code assumes that the target system is little-endian.
*/

#pragma warning(disable : 4244)  // warning C4244: 'argument': conversion from 'int64_t' to 'size_t', possible loss of data

#include "ReportMemoryLeaks.h"

#include "stdafx.h"
#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <errno.h>
#ifndef _WINDOWS
#include <error.h>
#endif

#include "cfbf.h"

void
cfbf_fat_close(struct cfbf_fat* fat) {
	free(fat->sector_entries);
	free(fat->fat_sectors);
}

SECT
cfbf_fat_get_sector_entry(struct cfbf_fat* fat, SECT sect)
{
	// ZMM Fix int casting
	if ((int)sect >= fat->sector_entries_count) {
		return CFBF_FREESECT;
	}
	else {
		return fat->sector_entries[sect];
	}
}

int
cfbf_fat_open(struct cfbf_fat* fat, struct cfbf* cfbf, SECT* start_sectors,
	unsigned long start_sectors_len, SECT difat_first_cont_sector,
	unsigned long num_cont_sectors, unsigned long num_fat_sectors_expected, std::string& errorText)
{
	unsigned long sector_size;
	unsigned long sect_ents_per_sect;
	unsigned long max_fat_sect;
	int retval = 0;
	SECT* difat_sect_data = NULL;

	memset(fat, 0, sizeof(*fat));

	sector_size = cfbf_get_sector_size(cfbf);
	sect_ents_per_sect = sector_size / sizeof(SECT);

	fat->sector_entries = (SECT*)calloc(num_fat_sectors_expected, sector_size);
	if (fat->sector_entries == NULL) {
		errorText = "failed to allocate memory for FAT sector entries";
		goto fail;
	}

	fat->fat_sectors = (SECT**)calloc(num_fat_sectors_expected, sizeof(SECT*));
	if (fat->fat_sectors == NULL) {
		errorText = "failed to allocate memory for FAT sector pointers";
		goto fail;
	}
	fat->num_fat_sectors = num_fat_sectors_expected;
	fat->sector_size = sector_size;

	max_fat_sect = num_fat_sectors_expected * sect_ents_per_sect;

	fat->sector_entries_count = 0;
	// ZMM Fix int casting
	for (int i = 0; i < (int)start_sectors_len; ++i)
	{
		// ZMM Fix int casting
		if (i >= (int)num_fat_sectors_expected) {
			char txt[512];
			sprintf(txt, "cfbf_fat_open: expected %lu sectors in FAT chain, but about to write sector with index %d", num_fat_sectors_expected, i);
			errorText = txt;
			goto fail;
		}
		//fprintf(stderr, "reading sector %lu\n", (unsigned long) start_sectors[i]);
		if (cfbf_read_sector(cfbf, start_sectors[i], fat->sector_entries + i * sect_ents_per_sect, errorText) < 0) {
			goto fail;
		}
		fat->fat_sectors[fat->sector_entries_count / sect_ents_per_sect] = (SECT*)cfbf_get_sector_ptr(cfbf, start_sectors[i], errorText);
		fat->sector_entries_count += sect_ents_per_sect;
	}

	SECT difat_cont_sector = difat_first_cont_sector;
	difat_sect_data = (SECT*)cfbf_malloc(sector_size);
	if (difat_sect_data == NULL)
		goto fail;

	// ZMM Fix int casting
	for (int i = 0; i < (int)num_cont_sectors; ++i)
	{
		if (cfbf_read_sector(cfbf, difat_cont_sector, difat_sect_data, errorText) < 0)
			goto fail;

		/* Each DIFAT sector contains 127 sector numbers (each of which
		 * refer to a FAT sector) followed by a reference to the next
		 * DIFAT sector. */
		 // ZMM Fix int casting
		for (int j = 0; j < (int)sect_ents_per_sect - 1; ++j)
		{
			SECT fat_sector = difat_sect_data[j];

			if (!CFBF_IS_SECTOR(fat_sector))
			{
				if (i == num_cont_sectors - 1 && fat->sector_entries_count == max_fat_sect) {
					/* This is the last DIFAT sector, and we now have the
					 * number of entries we expect, so the rest of this DIFAT
					 * sector need not be full of valid sectors. */
					break;
				}
				else {
					char txt[512];
					sprintf(txt, "cfbf_fat_open: fat_sector 0x%08x (invalid sector) found at slot %d in DIFAT sector, sector %lu", fat_sector, j, (unsigned long)difat_cont_sector);
					errorText = txt;
					goto fail;
				}
			}
			if (fat->sector_entries_count + sect_ents_per_sect > max_fat_sect) {
				char txt[1024];
				sprintf(txt, "cfbf_fat_open: i %d, j %d, fat_sector %lu, expected up to %lu sector entries in FAT chain, but about to write sectors %lu-%lu",
					i, j, (unsigned long)fat_sector, max_fat_sect,
					(unsigned long)fat->sector_entries_count,
					(unsigned long)fat->sector_entries_count + sect_ents_per_sect - 1);
				errorText = txt;
				goto fail;
			}
			fat->fat_sectors[fat->sector_entries_count / sect_ents_per_sect] = (SECT*)cfbf_get_sector_ptr(cfbf, fat_sector, errorText);
			if (cfbf_read_sector(cfbf, fat_sector, fat->sector_entries + fat->sector_entries_count, errorText) < 0) {
				goto fail;
			}
			fat->sector_entries_count += sect_ents_per_sect;
		}

		//fprintf(stderr, "i %d, difat cont sector %lu\n", i, (unsigned long) difat_cont_sector);
		//difat_cont_sector = cfbf_fat_get_sector_entry(fat, difat_cont_sector);
		difat_cont_sector = difat_sect_data[sect_ents_per_sect - 1];
		if (CFBF_IS_SECTOR(difat_cont_sector)) {
			if (i == num_cont_sectors - 1) {
				char txt[512];
				sprintf(txt, "difat continuation sector entry is 0x%08x but this should be the last sector in the chain", (unsigned int)difat_cont_sector);
				errorText = txt;
				goto fail;
			}
		}
		else if (difat_cont_sector == CFBF_END_OF_CHAIN) {
			// ZMM Fix int casting
			if (i < (int)(num_cont_sectors - 1)) {
				char txt[512];
				sprintf(txt, "continuation sector index %d was ENDOFCHAIN, expected new sector", i);
				errorText = txt;
				goto fail;
			}
		}
		else {
			char txt[512];
			sprintf(txt, "invalid sector number in difat chain: 0x%08x", (unsigned int)difat_cont_sector);
			errorText = txt;
			goto fail;
		}
	}

	if (fat->sector_entries_count != max_fat_sect) {
		char txt[512];
		sprintf(txt, "cfbf_fat_open: expected %lu FAT sectors, got %lu", (unsigned long)max_fat_sect, (unsigned long)fat->sector_entries_count);
		errorText = txt;
		goto fail;
	}

end:
	free(difat_sect_data);
	return retval;

fail:
	cfbf_fat_close(fat);
	retval = -1;
	goto end;
}

int
cfbf_mini_fat_open(struct cfbf_fat* mini_fat, struct cfbf_fat* main_fat,
	struct cfbf* cfbf, SECT first_sector, unsigned long num_sectors, std::string& errorText)
{
	int mini_sector_size = cfbf_get_mini_fat_sector_size(cfbf);
	int main_sector_size = cfbf_get_sector_size(cfbf);

	memset(mini_fat, 0, sizeof(*mini_fat));

	mini_fat->sector_size = mini_sector_size;
	if (num_sectors == 0) {
		mini_fat->sector_entries = NULL;
		mini_fat->sector_entries_count = 0;
		return 0;
	}

	mini_fat->sector_entries = (SECT*)calloc(num_sectors, main_sector_size);
	mini_fat->sector_entries_count = 0;

	if (mini_fat->sector_entries == NULL) {
		char txt[512];
		sprintf(txt, "failed to allocate %lu sectors for mini-FAT", (unsigned long)num_sectors);
		errorText = txt;
		goto fail;
	}

	SECT current_sector = first_sector;
	unsigned long sect_entries_per_sect = main_sector_size / sizeof(SECT);
	for (unsigned long sector_index = 0; sector_index < num_sectors; ++sector_index)
	{
		if (cfbf_read_sector(cfbf, current_sector, mini_fat->sector_entries + mini_fat->sector_entries_count, errorText) < 0) {
			goto fail;
		}
		mini_fat->sector_entries_count += sect_entries_per_sect;
		current_sector = cfbf_fat_get_sector_entry(main_fat, current_sector);

		if (current_sector == CFBF_END_OF_CHAIN) {
			if (sector_index != num_sectors - 1) {
				char txt[512];
				sprintf(txt, "found END_OF_CHAIN for sector index %lu of mini-FAT chain, but we're expecting %lu sectors total", sector_index, num_sectors);
				errorText = txt;
				goto fail;
			}
		}
		else if (CFBF_IS_SECTOR(current_sector)) {
			if (sector_index == num_sectors - 1) {
				char txt[512];
				sprintf(txt, "found sector %lu as next sector in sector index %lu of mini-FAT chain, but we're expecting only %lu sectors total", (unsigned long)current_sector, sector_index, num_sectors);
				errorText = txt;
				goto fail;
			}
		}
		else {
			char txt[512];
			sprintf(txt, "Invalid next sector value %lu as next sector in sector index %lu of mini-FAT chain", (unsigned long)current_sector, sector_index);
			errorText = txt;
			goto fail;
		}
	}

	return 0;

fail:
	cfbf_fat_close(mini_fat);
	return -1;
}


static void**
cfbf_get_chain_ptrs_aux(struct cfbf* cfbf, SECT first_sector, int* num_sectors_r, int use_mini_stream, std::string& errorText)
{
	void** sects = NULL;
	int sects_count = 0;
	int sects_size = 4;
	SECT current_sector;
	struct cfbf_fat* fat;

	if (use_mini_stream) {
		fat = &cfbf->mini_fat;
	}
	else {
		fat = &cfbf->fat;
	}

	sects = (void**)cfbf_malloc(sects_size * sizeof(void*));
	if (sects == NULL) {
		goto nomem;
	}

	for (current_sector = first_sector; current_sector != CFBF_END_OF_CHAIN; current_sector = cfbf_fat_get_sector_entry(fat, current_sector))
	{
		void* p;

		if (use_mini_stream)
			p = cfbf_get_sector_ptr_in_mini_stream(cfbf, current_sector);
		else
			p = cfbf_get_sector_ptr(cfbf, current_sector, errorText);

		if (p == NULL) {
			goto fail;
		}
		sects[sects_count++] = p;
		if (sects_count >= sects_size) {
			void** new_sects = (void**)realloc(sects, sects_size * 2 * sizeof(void*));
			if (new_sects == NULL)
				goto nomem;
			sects = new_sects;
			sects_size *= 2;
		}
	}
	sects[sects_count] = NULL;
	if (num_sectors_r)
		*num_sectors_r = sects_count;

	return sects;

nomem:
	char txt[512];
	sprintf(txt, "cfbf_get_chain_ptrs(): out of memory");
	errorText = txt;

fail:
	free(sects);
	return NULL;
}

void**
cfbf_get_chain_ptrs(struct cfbf* cfbf, SECT first_sector, int* num_sectors_r, std::string& errorText) {
	return cfbf_get_chain_ptrs_aux(cfbf, first_sector, num_sectors_r, 0, errorText);
}

void**
cfbf_get_chain_ptrs_from_mini_stream(struct cfbf* cfbf, SECT first_sector, int* num_sectors_r, std::string& errorText) {
	return cfbf_get_chain_ptrs_aux(cfbf, first_sector, num_sectors_r, 1, errorText);
}

void*
cfbf_alloc_chain_contents_from_fat(struct cfbf* cfbf, SECT first_sector, size_t size, std::string& errorText)
{
	int sector_size = cfbf_get_sector_size(cfbf);
	void* data;
	size_t data_pos = 0;
	int read_partial_sector = 0;

	if (size == 0) {
		return NULL;
	}

	data = cfbf_malloc(size);
	if (data == NULL) {
		goto nomem;
	}

	for (SECT sec = first_sector; sec != CFBF_END_OF_CHAIN; sec = cfbf_fat_get_sector_entry(&cfbf->fat, sec))
	{
		size_t to_copy;
		void* p;

		if (read_partial_sector) {
			char txt[512];
			sprintf(txt, "reached where we expected EOF to be, based on file size, but there are more sectors? sec %lu", (unsigned long)sec);
			errorText = txt;
			goto fail;
		}

		to_copy = size - data_pos;
		if ((int)to_copy > sector_size)  // Fix type casting ZMM
			to_copy = sector_size;

		p = cfbf_get_sector_ptr(cfbf, sec, errorText);
		if (p == NULL) {
			char txt[512];
			sprintf(txt, "failed to copy contents, sector %lu not valid", (unsigned long)sec);
			errorText = txt;
			goto fail;
		}

		memcpy((char*)data + data_pos, p, to_copy);
		data_pos += to_copy;

		if ((int)to_copy < sector_size)  // Fix type casting ZMM
			read_partial_sector = 1; // should be the last sector
	}

	if (data_pos != size) {
		char txt[512];
		sprintf(txt, "expected %llu bytes, got %llu", (unsigned long long) size, (unsigned long long) data_pos);
		errorText = txt;
		goto fail;
	}

	return data;

nomem:
	char txt[512];
	sprintf(txt, "cfbf_alloc_chain_contents_from_fat(): no memory");
	errorText = txt;

fail:
	free(data);
	return NULL;
}

void*
cfbf_alloc_chain_content(struct cfbf* cfbf, SECT first_sector, int64_t data_size,
	int use_mini_stream, void* cookie, std::string& errorText)
{
	struct cfbf_fat* fat;
	SECT sector;
	int64_t file_offset = 0;
	FSINDEX sector_index = 0;
	void* data;

	if (data_size == 0) {
		return NULL;
	}

	data = cfbf_malloc(data_size);
	if (data == NULL) {
		char txt[512];
		sprintf(txt, "cfbf_alloc_chain_content(): no memory");
		//error(0, ENOMEM, "cfbf_alloc_chain_content()");
		errorText = txt;
		goto fail;
	}

	if (use_mini_stream)
		fat = &cfbf->mini_fat;
	else
		fat = &cfbf->fat;

	for (sector = first_sector; sector != CFBF_END_OF_CHAIN; sector = cfbf_fat_get_sector_entry(fat, sector))
	{
		void* ptr;
		int this_data_length;

		if (data_size >= 0 && file_offset >= data_size) {
			char txt[512];
			sprintf(txt, "cfbf_alloc_chain_content(): read %lld bytes but there are more sectors? sector %lu", (long long)file_offset, (unsigned long)sector);
			errorText = txt;
			goto fail;
		}

		if (use_mini_stream) {
			ptr = cfbf_get_sector_ptr_in_mini_stream(cfbf, sector);
		}
		else {
			ptr = cfbf_get_sector_ptr(cfbf, sector, errorText);
		}

		if (ptr == NULL) {
			char txt[512];
			sprintf(txt, "cfbf_alloc_chain_content(): failed to fetch pointer for sector %lu", (unsigned long)sector);
			errorText = txt;
			goto fail;
		}

		if (data_size < 0 || data_size - file_offset >= fat->sector_size)
			this_data_length = fat->sector_size;
		else
			this_data_length = (int)(data_size - file_offset);

		memcpy((char*)data + file_offset, ptr, this_data_length);

		file_offset += this_data_length;
		sector_index++;
	}

	if (data_size >= 0 && file_offset != data_size) {
		char txt[512];
		sprintf(txt, "cfbf_alloc_chain_content(): came to end of sector chain but only wrote %lld bytes (expected %lld)", (long long)file_offset, (long long)data_size);
		errorText = txt;
		goto fail;
	}
	_ASSERTE(!(((UINT64)data % 2) || ((UINT64)data % 4) || ((UINT64)data %8)));
	return data;
fail:
	free(data);
	return NULL;
}


char*
cfbf_read_entry_data(struct cfbf* cfbf, DirEntry* entry,
	int (*callback)(void* cookie, const void* sector_data, int length,
		FSINDEX sector_index, int64_t file_offset), void* cookie, std::string& errorText)
{
	SECT first_sector = entry->start_sector;
	int64_t data_size = entry->stream_size;

	int use_mini_stream = 0;
	if (entry->stream_size < cfbf->header->_ulMiniSectorCutoff)
		use_mini_stream = 1;

	char* data = (char*)cfbf_alloc_chain_content(cfbf, first_sector, data_size, use_mini_stream, cookie, errorText);

	return data;

}

char*
cfbf_read_entry_data(struct cfbf* cfbf, DirEntry* entry, int* data_len, std::string& errorText)
{
	void* cookie = 0;
	SECT first_sector = entry->start_sector;
	int64_t data_size = entry->stream_size;

	int use_mini_stream = 0;
	if (entry->stream_size < cfbf->header->_ulMiniSectorCutoff)
		use_mini_stream = 1;

	char* data = (char*)cfbf_alloc_chain_content(cfbf, first_sector, data_size, use_mini_stream, cookie, errorText);
	// ZMM Fix int casting
	*data_len = (int)data_size;

	return data;

}



/* Follow the chain in the FAT (or the mini-FAT, if use_mini_stream is set),
 * starting with first_sector. cfbf_follow_chain() will call callback() once
 * for each sector, passing it the cookie (which has meaning only to
 * callback()), a pointer to the data in the sector, "length" which is the
 * number of meaningful bytes in the sector (this will be no more than the
 * sector size, and could be less if this is the last sector), sector_index
 * (the number of sectors previously read), and file_offset (the number of
 * bytes previously passed to the callback function).
 *
 * If data_size is negative, then it is treated as unknown, we keep
 * following the sector chain until we reach the end, and the last sector
 * is passed in full to the callback (length set to the sector size).
 * Otherwise, if the number of sectors in the chain doesn't tally with
 * data_size, the function fails.
 *
 * The callback function must return 0 to indicate success, or a negative
 * number to indicate that there's been a failure and that cfbf_follow_chain()
 * should fail immediately.
 */
int
cfbf_follow_chain(struct cfbf* cfbf, SECT first_sector, int64_t data_size,
	int use_mini_stream,
	int (*callback)(void* cookie, const void* sector_data, int length,
		FSINDEX sector_index, int64_t file_offset), void* cookie, std::string& errorText)
{
	struct cfbf_fat* fat;
	SECT sector;
	int64_t file_offset = 0;
	FSINDEX sector_index = 0;

	if (use_mini_stream)
		fat = &cfbf->mini_fat;
	else
		fat = &cfbf->fat;

	for (sector = first_sector; sector != CFBF_END_OF_CHAIN; sector = cfbf_fat_get_sector_entry(fat, sector))
	{
		void* ptr;
		int ret;
		int this_data_length;

		if (data_size >= 0 && file_offset >= data_size) {
			char txt[512];
			sprintf(txt, "cfbf_follow_chain(): read %lld bytes but there are more sectors? sector %lu", (long long)file_offset, (unsigned long)sector);
			errorText = txt;
			goto fail;
		}

		if (use_mini_stream) {
			ptr = cfbf_get_sector_ptr_in_mini_stream(cfbf, sector);
		}
		else {
			ptr = cfbf_get_sector_ptr(cfbf, sector, errorText);
		}

		if (ptr == NULL) {
			char txt[512];
			sprintf(txt, "cfbf_follow_chain(): failed to fetch pointer for sector %lu", (unsigned long)sector);
			errorText = txt;
			goto fail;
		}

		if (data_size < 0 || data_size - file_offset >= fat->sector_size)
			this_data_length = fat->sector_size;
		else
			this_data_length = (int)(data_size - file_offset);

		ret = callback(cookie, ptr, this_data_length, sector_index, file_offset);
		if (ret != 0) {
			char txt[512];
			sprintf(txt, "cfbf_follow_chain(): callback returned failure");
			errorText = txt;
			goto fail;
		}

		file_offset += this_data_length;
		sector_index++;
	}

	if (data_size >= 0 && file_offset != data_size) {
		char txt[512];
		sprintf(txt, "cfbf_follow_chain(): came to end of sector chain but only wrote %lld bytes (expected %lld)", (long long)file_offset, (long long)data_size);
		errorText = txt;
		goto fail;
	}

	return 0;

fail:
	return -1;
}

int
cfbf_read_mini_stream(struct cfbf* cfbf, SECT first_sector, int64_t data_size,
	int (*callback)(void* cookie, const void* sector_data, int length,
		FSINDEX sector_index, int64_t file_offset), void* cookie, std::string& errorText)
{
	struct cfbf_fat* fat;
	SECT sector;
	int64_t file_offset = 0;
	FSINDEX sector_index = 0;

	fat = &cfbf->mini_fat;

	for (sector = first_sector; sector != CFBF_END_OF_CHAIN; sector = cfbf_fat_get_sector_entry(fat, sector))
	{
		void* ptr;
		int ret;
		int this_data_length;

		if (data_size >= 0 && file_offset >= data_size) {
			char txt[512];
			sprintf(txt, "cfbf_follow_chain(): read %lld bytes but there are more sectors? sector %lu", (long long)file_offset, (unsigned long)sector);
			errorText = txt;
			goto fail;
		}

		ptr = cfbf_get_sector_ptr_in_mini_stream(cfbf, sector);


		if (ptr == NULL) {
			char txt[512];
			sprintf(txt, "cfbf_follow_chain(): failed to fetch pointer for sector %lu", (unsigned long)sector);
			errorText = txt;
			goto fail;
		}

		if (data_size < 0 || data_size - file_offset >= fat->sector_size)
			this_data_length = fat->sector_size;
		else
			this_data_length = (int)(data_size - file_offset);

		ret = callback(cookie, ptr, this_data_length, sector_index, file_offset);
		if (ret != 0) {
			char txt[512];
			sprintf(txt, "cfbf_follow_chain(): callback returned failure");
			errorText = txt;
			goto fail;
		}

		file_offset += this_data_length;
		sector_index++;
	}

	if (data_size >= 0 && file_offset != data_size) {
		char txt[512];
		sprintf(txt, "cfbf_follow_chain(): came to end of sector chain but only wrote %lld bytes (expected %lld)", (long long)file_offset, (long long)data_size);
		errorText = txt;
		goto fail;
	}

	return 0;

fail:
	return -1;
}


int
cfbf_read_stream(struct cfbf* cfbf, SECT first_sector, int64_t data_size,
	int (*callback)(void* cookie, const void* sector_data, int length,
		FSINDEX sector_index, int64_t file_offset), void* cookie, std::string& errorText)
{
	struct cfbf_fat* fat;
	SECT sector;
	int64_t file_offset = 0;
	FSINDEX sector_index = 0;

	fat = &cfbf->fat;

	for (sector = first_sector; sector != CFBF_END_OF_CHAIN; sector = cfbf_fat_get_sector_entry(fat, sector))
	{
		void* ptr;
		int ret;
		int this_data_length;

		if (data_size >= 0 && file_offset >= data_size) {
			char txt[512];
			sprintf(txt, "cfbf_follow_chain(): read %lld bytes but there are more sectors? sector %lu", (long long)file_offset, (unsigned long)sector);
			errorText = txt;
			goto fail;
		}

		ptr = cfbf_get_sector_ptr(cfbf, sector, errorText);

		if (ptr == NULL) {
			char txt[512];
			sprintf(txt, "cfbf_follow_chain(): failed to fetch pointer for sector %lu", (unsigned long)sector);
			errorText = txt;
			goto fail;
		}

		if (data_size < 0 || data_size - file_offset >= fat->sector_size)
			this_data_length = fat->sector_size;
		else
			this_data_length = (int)(data_size - file_offset);

		ret = callback(cookie, ptr, this_data_length, sector_index, file_offset);
		if (ret != 0) {
			char txt[512];
			sprintf(txt, "cfbf_follow_chain(): callback returned failure");
			errorText = txt;
			goto fail;
		}

		file_offset += this_data_length;
		sector_index++;
	}

	if (data_size >= 0 && file_offset != data_size) {
		char txt[512];
		sprintf(txt, "cfbf_follow_chain(): came to end of sector chain but only wrote %lld bytes (expected %lld)", (long long)file_offset, (long long)data_size);
		errorText = txt;
		goto fail;
	}

	return 0;

fail:
	return -1;
}

void* cfbf_malloc(size_t _Size)
{
	if (_Size == 236544)
		int deb = 1;
	return malloc(_Size);
}