/*
* Cfbfinfo is written by Graeme Cole
* https://github.com/elocemearg/cfbfinfo
*
* Ported by Zbigniew Minciel to Windows to integrate with free Windows MBox Viewer
*
* Note: For simplification, the code assumes that the target system is little-endian.
*/

// Publisher is not used by Windows MBox Viewer. It is just an example code.

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

/* The CONTENTS stream of a Publisher file is a series of 512-byte sectors,
 * as required by the container format.
 *
 * The stream starts with struct pub_contents_header.
 * After that is the first struct pub_contents_segment_list_header, and after
 * that is a number of struct pub_contents_segment_desc structs, the count
 * being given by pub_contents_segment_lsit_header.num_segments. The space for
 * the segment_desc structs may be padded out with zeroes to fill the rest of
 * the sector.
 *
 * If there are more segment descriptors than will fit in a sector, the
 * chain_next_offset member of struct pub_contents_segment_list_header points
 * to the next segment list header in the stream. Otherwise, it is 0xffffffff.
 *
 */

struct pub_contents_header {
    char magic[8]; // "CHNKINK "
    uint32_t crap1;
    uint16_t total_segments; // total count of segment descriptors in the stream
    uint16_t crap2;
    unsigned int unknown_zero : 9; // nine zero bits
    unsigned int num_segment_lists : 23; // count of segment list headers
    uint32_t stream_length;
};

struct pub_contents_segment_list_header {
    uint16_t crap; // usually f8 01

    // number of segment_desc structs that immediately follow this header
    uint16_t num_segments;

    // offset from start of CONTENTS stream of the next
    // pub_contents_segment_list_header, or 0xffffffff if there are no more
    // segment lists after this one
    uint32_t chain_next_offset;
};

struct pub_contents_segment_desc {
    uint16_t tag;  // 0x18 if this is valid
    char data_type[4];
    uint16_t data_type_args[3]; // usually 0 1 0
    char data_format[4];
    uint32_t offset;  // offset from start of CONTENTS stream
    uint32_t length; // length of segment in bytes
};


int
chain_read(void **chain, int num_sectors, int sector_size, size_t stream_size,
        void *dest, size_t offset, size_t size)
{
    size_t bytes_read = 0;

    while (bytes_read < size)
    {
        // ZMM Fix unsigned  casting
        _ASSERTE((offset / sector_size) >= 0);
        FSINDEX sector_index = (ULONG)(offset / sector_size);
        int offset_in_sector = offset % sector_size;
        size_t to_copy;

        // ZMM Fix int casting
        if ((int)sector_index >= num_sectors) {
            error(0, 0, "chain_read(): attempted to read block %lu of a %lu-block object", (unsigned long) sector_index, (unsigned long) num_sectors);
            return -1;
        }

        to_copy = sector_size - offset_in_sector;
        if (to_copy > size - bytes_read) {
            to_copy = size - bytes_read;
        }

        if (offset + to_copy > stream_size) {
            error(0, 0, "chain_read(): attempted to read data from offset %zd to %zd, but stream is only %zd bytes long", offset, offset + to_copy, stream_size);
            return -1;
        }

        memcpy((char *) dest + bytes_read,
                (char *) chain[sector_index] + offset_in_sector,
                to_copy);

        offset += to_copy;
        bytes_read += to_copy;
    }

    return 0;
}

int
extract_text_from_contents_chain(void **contents_chain, int num_sectors,
        int sector_size, size_t stream_size, int verbosity,
        int (*callback)(void *cookie, const char *text, size_t length),
        void *cookie)
{
    struct pub_contents_header header;
    struct pub_contents_segment_desc seg_desc;
    struct pub_contents_segment_list_header seg_list_header;
    size_t seg_list_header_offset;

    if (stream_size < sizeof(struct pub_contents_header)) {
        error(0, 0, "CONTENTS stream size is %zd bytes, too short to contain a header", stream_size);
        return -1;
    }

    if (chain_read(contents_chain, num_sectors, sector_size, stream_size,
                &header, 0, sizeof(header)) < 0) {
        error(0, 0, "failed to read header from CONTENTS stream");
        return -1;
    }

    if (strncmp(header.magic, "CHNKINK ", 8)) {
        error(0, 0, "bad signature at front of CONTENTS header: expected \"CHNKINK \", got \"%.8s\"", header.magic);
        return -1;
    }

    seg_list_header_offset = sizeof(header);
    int num_seg_list_headers = 0;
    int num_segs = 0;

    do
    {
        /* Read the segment list header, which is at the top of a small
         * number of segment descriptors, and which also contains a pointer
         * to the next segment list header */
        if (chain_read(contents_chain, num_sectors, sector_size, stream_size,
                    &seg_list_header, seg_list_header_offset, sizeof(seg_list_header)) < 0) {
            error(0, 0, "failed to read segment list header from contents stream at offset %zd", seg_list_header_offset);
            return -1;
        }
        ++num_seg_list_headers;

        if (seg_list_header.crap != 0x1f8) {
            error(0, 0, "segment list header at offset %zd: expected magic number 0x1f8, got 0x%hx", seg_list_header_offset, (unsigned short) seg_list_header.crap);
            return -1;
        }
 
        for (int seg_index = 0; seg_index < seg_list_header.num_segments; ++seg_index)
        {
            /* Read the segment descriptor, which tells us what kind of segment
             * it is, where it is, and how long it is */
            size_t sd_offset = seg_list_header_offset + sizeof(seg_list_header) + seg_index * sizeof(seg_desc);
            if (chain_read(contents_chain, num_sectors, sector_size,
                        stream_size, &seg_desc,
                        sd_offset, sizeof(seg_desc)) < 0) {
                error(0, 0, "failed to read segment descriptor from contents stream at offset %zd", sd_offset);
                return -1;
            }

            if (seg_desc.tag == 0x18)
            {
                ++num_segs;
                if (!strncmp(seg_desc.data_type, "TEXT", 4) &&
                        !strncmp(seg_desc.data_format, "TEXT", 4))
                {
                    uint32_t offset = seg_desc.offset;
                    char buf[1024];

                    if (verbosity > 0) {
                        fprintf(stderr, "Reading TEXT/TEXT segment, offset %lu, length %lu... ",
                                (unsigned long) seg_desc.offset,
                                (unsigned long) seg_desc.length);
                    }

                    while (offset < seg_desc.offset + seg_desc.length)
                    {
                        size_t to_read = seg_desc.offset + seg_desc.length - offset;
                        if (to_read > sizeof(buf))
                            to_read = sizeof(buf);
                        if (chain_read(contents_chain, num_sectors, sector_size,
                                    stream_size, buf, offset, to_read) < 0) {
                            error(0, 0, "failed to read chunk of text from contents stream at offset %zd", (size_t) offset);
                            return -1;
                        }

                        if (callback(cookie, buf, to_read) < 0) {
                            error(0, 0, "error signalled by callback");
                            return -1;
                        }

                        offset += (uint32_t)to_read;
                    }

                    if (verbosity > 0)
                        fprintf(stderr, "done.\n");
                }
                else
                {
                    if (verbosity > 1) {
                        fprintf(stderr, "Skipping %.4s/%.4s segment, args %hu %hu %hu, offset %lu, length %lu\n",
                                seg_desc.data_type, seg_desc.data_format,
                                (unsigned short) seg_desc.data_type_args[0],
                                (unsigned short) seg_desc.data_type_args[1],
                                (unsigned short) seg_desc.data_type_args[2],
                                (unsigned long) seg_desc.offset,
                                (unsigned long) seg_desc.length);
                    }
                }
            }
        }

        seg_list_header_offset = seg_list_header.chain_next_offset;
    } while (seg_list_header_offset != 0xffffffff);

    if (verbosity > 1) {
        fprintf(stderr, "expected: %u segment descriptors in %u descriptor blocks\n", (unsigned int) header.total_segments, (unsigned int) header.num_segment_lists);
        fprintf(stderr, "observed: %d segment descriptors in %d descriptor blocks\n", num_segs, num_seg_list_headers);
    }

    return 0;
}
