/*
* Cfbfinfo is written by Graeme Cole
* https://github.com/elocemearg/cfbfinfo
*
* Ported by Zbigniew Minciel to Windows to integrate with free Windows MBox Viewer
*
* Note: For simplification, the code assumes that the target system is little-endian.
*/

#ifndef _CFBF_H
#define _CFBF_H

#pragma once

#include "ReportMemoryLeaks.h"

#include <stdint.h>

#ifndef _WINDOWS
// https://en.wikipedia.org/wiki/Compound_File_Binary_Format
// Modified to fix type sizes for Linux

typedef uint32_t ULONG;    // 4 Bytes
typedef uint16_t USHORT;  // 2 Bytes
typedef int64_t OFFSET;           // 2 Bytes
typedef ULONG SECT;             // 4 Bytes
typedef ULONG FSINDEX;          // 4 Bytes
typedef USHORT FSOFFSET;        // 2 Bytes
typedef USHORT WCHAR;           // 2 Bytes
typedef ULONG DFSIGNATURE;      // 4 Bytes
typedef uint8_t BYTE;     // 1 Byte
typedef uint16_t WORD;    // 2 Bytes
typedef uint32_t DWORD;    // 4 Bytes
typedef ULONG SID;              // 4 Bytes
typedef unsigned char CLSID[16];// 16 Bytes

typedef struct filetime {
    uint32_t low;
    uint32_t high;
} FILETIME;
#else
#include "cfbf_winhelper.h"
#endif

#define CFBF_END_OF_CHAIN 0xFFFFFFFE
#define CFBF_FREESECT 0xFFFFFFFF
#define CFBF_FATSECT 0xFFFFFFFD
#define CFBF_DIFSECT 0xFFFFFFFC

#define CFBF_IS_SECTOR(x) (((uint32_t) (x)) < 0xFFFFFFFCUL)

struct StructuredStorageHeader { // [offset from start (bytes), length (bytes)]
    BYTE _abSig[8];             // [00H,08] {0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1,
                                // 0x1a, 0xe1} for current version
    CLSID _clsid;               // [08H,16] reserved must be zero (WriteClassStg/
                                // GetClassFile uses root directory class id)
    USHORT _uMinorVersion;      // [18H,02] minor version of the format: 33 is
                                // written by reference implementation
    USHORT _uDllVersion;        // [1AH,02] major version of the dll/format: 3 for
                                // 512-byte sectors, 4 for 4 KB sectors
    USHORT _uByteOrder;         // [1CH,02] 0xFFFE: indicates Intel byte-ordering
    USHORT _uSectorShift;       // [1EH,02] size of sectors in power-of-two;
                                // typically 9 indicating 512-byte sectors
    USHORT _uMiniSectorShift;   // [20H,02] size of mini-sectors in power-of-two;
                                // typically 6 indicating 64-byte mini-sectors
    USHORT _usReserved;         // [22H,02] reserved, must be zero
    ULONG _ulReserved1;         // [24H,04] reserved, must be zero
    FSINDEX _csectDir;          // [28H,04] must be zero for 512-byte sectors,
                                // number of SECTs in directory chain for 4 KB
                                // sectors
    FSINDEX _csectFat;          // [2CH,04] number of SECTs in the FAT chain
    SECT _sectDirStart;         // [30H,04] first SECT in the directory chain
    DFSIGNATURE _signature;     // [34H,04] signature used for transactions; must
                                // be zero. The reference implementation
                                // does not support transactions
    ULONG _ulMiniSectorCutoff;  // [38H,04] maximum size for a mini stream;
                                // typically 4096 bytes
    SECT _sectMiniFatStart;     // [3CH,04] first SECT in the MiniFAT chain
    FSINDEX _csectMiniFat;      // [40H,04] number of SECTs in the MiniFAT chain
    SECT _sectDifStart;         // [44H,04] first SECT in the DIFAT chain
    FSINDEX _csectDif;          // [48H,04] number of SECTs in the DIFAT chain
    SECT _sectFat[109];         // [4CH,436] the SECTs of first 109 FAT sectors
};

#define CFBF_NOSTREAM (0xffffffff)
#define CFBF_MAXREGSID (0xfffffffa)

#define CFBF_UNKNOWN_OBJECT_TYPE 0x00
#define CFBF_STORAGE_OBJECT_TYPE 0x01
#define CFBF_STREAM_OBJECT_TYPE 0x02
#define CFBF_ROOT_OBJECT_TYPE 0x05

struct DirEntry {
    uint16_t name[32]; // name in UTF-16
    uint16_t name_length;
    uint8_t object_type;
    uint8_t colour;
    uint32_t left_sibling_id;
    uint32_t right_sibling_id;
    uint32_t child_id;
    uint8_t clsid[16];
    uint32_t state;
    FILETIME creation_time;
    FILETIME modified_time;
    uint32_t start_sector;
    uint64_t stream_size;
};


/* Reference:
 * https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/28488197-8193-49d7-84d8-dfd692418ccd
 */

struct cfbf_fat {
    SECT *sector_entries;
    int sector_entries_count;
    int sector_size;

    /* Pointers into the actual cfbf file */
    SECT **fat_sectors;
    int num_fat_sectors;
};

struct cfbf {
    wchar_t* filepath;
    int fd;
    void *file;
    long long file_size;
    struct StructuredStorageHeader *header;
    struct cfbf_fat fat, mini_fat;

    void *mini_stream;
    size_t mini_stream_size;


    // Added until major restructing is implemented
    bool printEnabled;
    FILE* out;  // file handle for printing
};

void
cfbf_fat_close(struct cfbf_fat *fat);

SECT
cfbf_fat_get_sector_entry(struct cfbf_fat *fat, SECT sect);

int
cfbf_fat_open(struct cfbf_fat *fat, struct cfbf *cfbf, SECT *start_sectors,
        unsigned long start_sectors_len, SECT difat_first_cont_sector,
        unsigned long num_cont_sectors, unsigned long num_fat_sectors_expected, std::string& errorText);

int
cfbf_mini_fat_open(struct cfbf_fat *mini_fat, struct cfbf_fat *main_fat,
        struct cfbf *cfbf, SECT first_sector, unsigned long num_sectors, std::string& errorText);

void
cfbf_close(struct cfbf *cfbf);

int
cfbf_open(const wchar_t *filename, struct cfbf *cfbf, std::string& errorText);

int
cfbf_get_sector_size(struct cfbf *cfbf);

int
cfbf_get_mini_fat_sector_size(struct cfbf *cfbf);

int
cfbf_read_sector(struct cfbf* cfbf, SECT sect, void* dest, std::string& errorText);

void *
cfbf_get_sector_ptr(struct cfbf *cfbf, SECT sect, std::string& errorText);

void *
cfbf_get_sector_ptr_in_mini_stream(struct cfbf *cfbf, SECT sector);

int
cfbf_is_sector_in_file(struct cfbf *cfbf, SECT sect);

void **
cfbf_get_chain_ptrs(struct cfbf *cfbf, SECT first_sector, int *num_sectors_r, std::string& errorText);

void **
cfbf_get_chain_ptrs_from_mini_stream(struct cfbf *cfbf, SECT first_sector, int *num_sectors_r, std::string& errorText);

void *
cfbf_alloc_chain_contents_from_fat(struct cfbf *cfbf, SECT first_sector, size_t size, std::string& errorText);

int
cfbf_follow_chain(struct cfbf *cfbf, SECT first_sector, int64_t data_size,
        int use_mini_stream,
        int (*callback)(void *cookie, const void *sector_data, int length,
            FSINDEX sector_index, int64_t file_offset), void *cookie, std::string& errorText);

int
cfbf_walk(struct cfbf *cfbf, FILE *out, int verbosity, std::string& errorText);


int
cfbf_walk_dir_tree(struct cfbf *cfbf,
        int (*callback)(void *cookie, struct cfbf *cfbf, DirEntry *entry,
            DirEntry *parent, unsigned long entry_id, int depth),
        void *cookie, std::string& errorText);

DirEntry *
cfbf_dir_entry_find_path(struct cfbf *cfbf, char *sought_path_utf8, std::string& errorText);

void
cfbf_object_type_to_string(int object_type, char *dest, int dest_max);

int
cfbf_dir_stored_in_mini_stream(struct cfbf *cfbf, DirEntry *entry);

void **
cfbf_dir_entry_get_sector_ptrs(struct cfbf *cfbf, DirEntry *entry, int *num_sectors_r, int *sector_size_r, std::string& errorText);

int
extract_text_from_contents_chain(void **contents_chain, int num_sectors,
        int sector_size, size_t stream_size, int verbosity,
        int (*callback)(void *cookie, const char *text, size_t length),
        void *cookie);

char*
cfbf_read_entry_data(struct cfbf* cfbf, DirEntry* entry, int *data_len, std::string& errorText);

void* cfbf_malloc(size_t _Size);

#endif
