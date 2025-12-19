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
#pragma warning(disable : 4244)  // warning C4244: '=': conversion from 'UINT64' to 'UINT', possible loss of data

#include "ReportMemoryLeaks.h"

#include "stdafx.h"

#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <atlstr.h>

#ifndef _WINDOWS
#include <sys/mman.h>
#include <getopt.h>
#include <errno.h>
#include <error.h>
#include <iconv.h>
#endif

#include "cfbf.h"
#include "outlook.h"
#include "FileUtils.h"



#if 0
void* __cdecl operator new(size_t s)
{
    static int cnt = 0;
    cnt++;
    if (cnt == 234094)
        int deb = 1;
    return ::operator new(s, _NORMAL_BLOCK, __FILE__, __LINE__);
}
#endif


int
ParseOutlookMsg(void* cookie, struct cfbf* cfbf, DirEntry* e,
    DirEntry* parent, unsigned long entry_id, int depth);

typedef enum {
    PROPERTY_NO_FILTER = 0,
    PROPERY_MAIL_FILTER = 1,
    PROPERTY_BODY_FILTER = 2,
    PROPERTY_ATTACHMENT_FILTER = 3,
} PropertyFilter;

PropertyFilter propFilter = PROPERTY_NO_FILTER;
static bool PrintEntryInfo = false;
static bool printEntryData = false;
int maxTextDumpLength = 32;
int maxBinaryDumpLength = 16;
int maxRtfDumpLength = 32;

char* id2name(UINT16 id)
{
    static char* unknown = "Unknown Property";
    if (propFilter == PROPERTY_NO_FILTER)
    {
        char* name = idAll2name(id);
#if 0
        if (name == 0)
            return unknown;
        else
#endif
            return name;
    }
    else if (propFilter == PROPERY_MAIL_FILTER)
    {
        char* name = idMailFilter2name(id);
        return name;
    }
    else if (propFilter == PROPERTY_BODY_FILTER)
    {
        char* name = idBody2name(id);
        return name;
    }
    else
    {
        return "Coding: Unknown Property Name";
    }
};

#ifdef OUTLOOK_MODE

struct write_pub_text_state {
	iconv_t iconv_desc;
	FILE* out;
};


int
print_dir_entry(void *cookie, struct cfbf *cfbf, DirEntry *e,
        DirEntry *parent, unsigned long entry_id, int depth)
{
    char name[129];
    //char *in_p;
    int name_length;
    int indent = (depth-1) * 4;
    int retval = 1;
    char obj_type_str[10];
    FILE *out = (FILE *) cookie;

    if (e->name_length > 64) {
        error(0, 0, "warning: dir entry %lu: name_length is %hu which is > 64", entry_id, (unsigned short) e->name_length);
        name_length = 64;
    }
    else {
        name_length = (int) e->name_length;
    }

    std::string name_utf8 = UTF16ToUTF8((wchar_t*)e->name, e->name_length);

    size_t out_len = 0;
    char* buff8 = UTF16ToUTF8((wchar_t*)e->name, e->name_length, &out_len);
    if (buff8 == 0)
    {
        goto fail;
    }
    strcpy(name, buff8);
    free(buff8);

    cfbf_object_type_to_string(e->object_type, obj_type_str, sizeof(obj_type_str));

   
    int isMiniStream = cfbf_dir_stored_in_mini_stream(cfbf, e);

    if (strncmp(name, "__attach_version1.0_", strlen("__attach_version1.0_")) == 0)
        int deb = 1;

#if 0
    fprintf(out, "%-8s %10lu%s %10llu    ", obj_type_str,
            (unsigned long) e->start_sector,
            isMiniStream ? "m" : " ",
            (unsigned long long) e->stream_size);

    fprintf(out, "%*s", indent, "");
    fprintf(out, "%s", name);
#endif


    int isStream = (e->object_type == CFBF_STREAM_OBJECT_TYPE) ? 1 : CFBF_UNKNOWN_OBJECT_TYPE;
    if (isStream && (strncmp(name, "__substg1.0_", strlen("__substg1.0_")) == 0))
    {
        int prefixlen = strlen("__substg1.0_");
        char suffix[9];
        strncpy(suffix, &name[prefixlen], 8);
        suffix[8] = 0;
        char propertyId[5];
        char propertyType[5];
        strncpy(propertyId, suffix, 4);
        propertyId[4] = 0;
        strncpy(propertyType, &suffix[4], 4);
        propertyType[4] = 0;

        int propertIdNumb = (int)strtol(propertyId, NULL, 16);
        char* propertyName = id2name((UINT32)propertIdNumb);
        if (propertyName)
        {
            if ((unsigned long)e->start_sector == CFBF_END_OF_CHAIN) // == 4294967294ULL or 0xFFFFFFFE
            {
                fprintf(out, "%-8s %8s%s %10llu    ", obj_type_str,
                    "FFFFFFFE",
                    isMiniStream ? "m" : " ",
                    (unsigned long long) e->stream_size);
            }
            else
            {
                fprintf(out, "%-8s %8lu%s %10llu    ", obj_type_str,
                    (unsigned long)e->start_sector,
                    isMiniStream ? "m" : " ",
                    (unsigned long long) e->stream_size);
            }

            fprintf(out, "%*s", indent, "");
            fprintf(out, "%s", name);

            fprintf(out, "  %s", propertyName);
            char* type = id2type(propertyType);
            if (type)
            {
                fprintf(out, "  %s", type);
            }
            else
            {
                fprintf(out, "  Uknown Property Type 0x%s", propertyType);
            }

            fprintf(out, "\n");
        }
#if 0
        else
        {
            fprintf(out, "    Uknown Property Name 0x%04x", propertIdNumb);
        }

        char* type = id2type(propertyType);
        if (type)
        {
            fprintf(out, "    %s", type);
        }
        else
        {
            fprintf(out, "    Uknown Property Type 0x%s", propertyType);
        }
#endif

        int deb = 1;
    }
    else if (strncmp(name, "__", 2) == 0)
    {
        if (strstr(name, "DocumentSummaryInformation") != 0)
            int deb = 1;

        if ((unsigned long)e->start_sector == CFBF_END_OF_CHAIN) // == 4294967294ULL or 0xFFFFFFFE
        {
            fprintf(out, "%-8s %8s%s %10llu    ", obj_type_str,
                "FFFFFFFE",
                isMiniStream ? "m" : " ",
                (unsigned long long) e->stream_size);
        }
        else
        {
            fprintf(out, "%-8s %8lu%s %10llu    ", obj_type_str,
                (unsigned long)e->start_sector,
                isMiniStream ? "m" : " ",
                (unsigned long long) e->stream_size);
        }

        fprintf(out, "%*s", indent, "");
        fprintf(out, "%s", name);
        fprintf(out, "\n");
    }
    else
    {
        ; // ignore
    }

    //fprintf(out, "\n");

end:

    return retval;

fail:
    retval = -1;
    goto end;
}

int
write_publisher_text(void *cookie, const char *data, size_t data_length)
{
    struct write_pub_text_state *state = (struct write_pub_text_state *) cookie;
    if (state->iconv_desc == (iconv_t) -1)
    {
        /* We're not converting the character encoding, we're just writing
         * it all straight out */
        size_t ret = fwrite(data, 1, data_length, state->out);
        if (ret != data_length) {
            error(0, errno, "fwrite()");
            return -1;
        }
    }
    else
    {
        /* Convert text to utf8 before writing it out */

        //while (in_left > 0)
        {
            size_t ret;

            size_t out_len = 0;
            char* buff8 = UTF16ToUTF8((wchar_t*)data, data_length/2, &out_len);
            if (buff8 == 0)
            {
                //error(0, errno, "write_publisher_text()");
                return -1;
            }


            ret = fwrite(buff8, 1, out_len, state->out);
            if (ret != out_len) {
                error(0, errno, "fwrite()");
                free(buff8);
                return -1;
            }
            free(buff8);
        }
    }

    return 0;
}

int
write_sector_to_file(void *cookie, const void *sector_data, int length,
        FSINDEX sector_index, int64_t file_offset)
{
    FILE *out = (FILE *) cookie;
    size_t ret;
    char* data = (char*)sector_data;
    std::string name_utf8 = UTF16ToUTF8((wchar_t*)data, length/2);
    ret = fwrite(data, 1, length, out);
    if (ret != length) {
        error(0, errno, "write_sector_to_file()");
        return -1;
    }

    return 0;
}

void
print_help(FILE *out) {
    fprintf(out, "Compound File Binary File format analyser\n");
    fprintf(out, "Graeme Cole, 2019\n");
    fprintf(out, "Usage: cfbfinfo [action] [options] file.pub\n");
    fprintf(out, "Actions:\n");
    fprintf(out, "    -h         Show this help\n");
    fprintf(out, "    -l         List directory tree\n");
    fprintf(out, "    -r <path>  Dump the object with this path to the output file\n");
    fprintf(out, "               (e.g. -r \"Root Entry/Quill/QuillSub/CONTENTS\")\n");
    fprintf(out, "    -t         Extract TEXT section from CONTENTS object, write to output file\n");
    fprintf(out, "    -w         Walk FAT structure, highlight any problems\n");
    fprintf(out, "Options:\n");
    fprintf(out, "    -c <path>  [with -t] Path to use for CONTENTS object\n");
    fprintf(out, "               (default is \"Root Entry/Quill/QuillSub/CONTENTS\")\n");
    fprintf(out, "    -o <file>  Output file name (default is stderr for -w, stdout otherwise)\n");
    fprintf(out, "    -a         Append to Output file to use with -o option\n");
    fprintf(out, "    -q         Be less verbose\n");
    fprintf(out, "    -u         [with -t] Don't convert text to UTF-8 for output, keep as UTF-16\n");
    fprintf(out, "    -v         Be more verbose\n");
    fprintf(out, "\n");
    fprintf(out, "Use -t to extract text from a Microsoft Publisher file.\n");
    fprintf(out, "If there are no action arguments, print information from the header and exit.\n");
}


int wmain(int argc, wchar_t **argv)
{
    wchar_t *input_filename = NULL;
    std::string input_filename_utf8;
    struct cfbf cfbf;
    int show_header = 0;
    //
    wchar_t *wdump_object_path = NULL;
    char* dump_object_path = NULL;
    std::string dump_object_path_utf8;
    int print_dir_tree = 0;
    //
    wchar_t *woutput_filename = NULL;
    const char* output_filename = NULL;
    std::string output_filename_utf8;
    //
    FILE *out = NULL;
    int walk = 0;
    int extract_publisher_text = 0;
    int exit_status = 0;
    int num_command_options = 0;
    int verbosity = 0;
    //
    wchar_t *wpublisher_contents_path = L"Root Entry/Quill/QuillSub/CONTENTS";
    char* publisher_contents_path = "Root Entry/Quill/QuillSub/CONTENTS";
    std::string publisher_contents_path_utf8;
    //
    int convert_text_to_utf8 = 1;
    int create_mbox = 0;
    bool append_to_output_file = false;

    const char* cmd = nullptr;
    const char* file = nullptr;

    bool dumpraw = false;
    bool escape = false;
    for (int i = 1; i < argc; i++)
    {
        wchar_t* arg = argv[i];
        if (wcscmp(arg, L"-h") == 0)
        {
            print_help(stdout);
            exit(0);
        }
        else if (wcscmp(arg, L"-msg2mbox") == 0)
        {
            create_mbox = 1;
            ++num_command_options;
        }
        else if (wcscmp(arg, L"-l") == 0)
        {
            print_dir_tree = 1;
            ++num_command_options;
        }
        else if (wcscmp(arg, L"-r") == 0)
        {
            i++;
            wdump_object_path = argv[i];
            int wdump_object_path_len = wcslen(wdump_object_path);
            dump_object_path_utf8 = UTF16ToUTF8((wchar_t*)wdump_object_path, wdump_object_path_len * 2);
            dump_object_path = (char*)dump_object_path_utf8.c_str();

            // Skip any leading slashes - we don't want them  ZMM ???
        }
        else if (wcscmp(arg, L"-t") == 0)
        {
            extract_publisher_text = 1;
            ++num_command_options;
        }
        else if (wcscmp(arg, L"-w") == 0)
        {
            walk = 1;
            ++num_command_options;
        }
        else if (wcscmp(arg, L"-c") == 0)
        {
            i++;
            wpublisher_contents_path = argv[i];
            int wpublisher_contents_path_len = wcslen(wpublisher_contents_path);
            publisher_contents_path_utf8 = UTF16ToUTF8((wchar_t*)wpublisher_contents_path, wpublisher_contents_path_len * 2);
            publisher_contents_path = (char*)publisher_contents_path_utf8.c_str();
        }
        else if (wcscmp(arg, L"-o") == 0)
        {
            i++;
            woutput_filename = argv[i];
            int woutput_filename_len = wcslen(woutput_filename);
            output_filename_utf8 = UTF16ToUTF8((wchar_t*)woutput_filename, woutput_filename_len*2);
            output_filename = output_filename_utf8.c_str();
        }
        else if (wcscmp(arg, L"-a") == 0)
        {
            append_to_output_file = true;
        }
        else if (wcscmp(arg, L"-q") == 0)
        {
            verbosity--;
        }
        else if (wcscmp(arg, L"-u") == 0)
        {
            convert_text_to_utf8 = 0;
        }
        else if (wcscmp(arg, L"-i") == 0)
        {
            i++;
            input_filename = argv[i];
        }
        else if (wcscmp(arg, L"-v") == 0)
        {
            verbosity++;
        }
        else
        {
        }
    }



    /* We can only do one action */
    if (num_command_options > 1) {
        error(1, 0, "Only one of -r, -l, -t and -w may be given. Use -h for help.");
    }

    /* If no actions have been specified, print information from the header */
    if (num_command_options == 0) {
        show_header = 1;
    }

    fwprintf(stdout, L"cfbfinfo.exe: Opening file: %s\n", input_filename);

    /* Open the CFB file, which will fail if there's something seriously
     * wrong with it, like it not being a CFB file */
    std::string errorText;
    if (cfbf_open(input_filename, &cfbf, errorText) != 0) {
        exit(1);
    }

    /* If an output file has been specified, open it */
    if (output_filename == NULL || !strcmp(output_filename, "-")) {
        if (walk)
            out = stderr;
        else
            out = stdout;
    }
    else {
        char *fopen_mode = "wb";  // "w" Opens an empty file for writing. If the given file exists, its contents are destroyed.
        if (append_to_output_file)
            fopen_mode = "ab";    // "a" Opens for writing at the end of the file (appending). Creates the file if it doesn't exist.
        out = fopen(output_filename, fopen_mode);
        if (out == NULL)
            error(1, errno, "%s", output_filename);
    }

    int input_filename_len = wcslen(input_filename);

    if ((out != stdout) && (out != stderr)) {
        input_filename_utf8 = UTF16ToUTF8((wchar_t*)input_filename, input_filename_len);
        fprintf(out, "cfbfinfo.exe: Opening file: %s\n", input_filename_utf8.c_str());
    }
    else
        fwprintf(out, L"cfbfinfo.exe: Opening file: %s\n", input_filename);

    /* Do whatever action we've been told to do */
    struct StructuredStorageHeader *header = cfbf.header;
    if (show_header)
    {
        fprintf(out, "DllVersion, MinorVersion:     %hu, %hu\n", (unsigned short) header->_uDllVersion, (unsigned short) header->_uMinorVersion);
        fprintf(out, "Byte-order mark:              %02X %02X\n", ((unsigned char *) header)[0x1c], ((unsigned char *) header)[0x1d]);
        fprintf(out, "Main FAT sector size:         2^%hu (%d)\n", (unsigned short) header->_uSectorShift, cfbf_get_sector_size(&cfbf));
        fprintf(out, "Mini-stream sector size:      2^%hu (%d)\n", (unsigned short) header->_uMiniSectorShift, cfbf_get_mini_fat_sector_size(&cfbf));
        fprintf(out, "FAT chain sector count:       %lu\n", (unsigned long) header->_csectFat);
        if (header->_uSectorShift >= 12)
            fprintf(out, "Directory chain sector count: %lu\n", (unsigned long) header->_csectDir);
        fprintf(out, "Directory chain first sector: %lu\n", (unsigned long) header->_sectDirStart);
        fprintf(out, "Max file size in mini-stream: %lu\n", (unsigned long) header->_ulMiniSectorCutoff);
        fprintf(out, "MiniFAT first sector, count:  %lu, %lu\n", (unsigned long) header->_sectMiniFatStart, (unsigned long) header->_csectMiniFat);
        fprintf(out, "DIFAT first sector, count:    %lu, %lu\n", (unsigned long) header->_sectDifStart, (unsigned long) header->_csectDif);
        fprintf(out, "\n");
    }
    else if (create_mbox)
    {
        FILE* outFile = out;
        if (outFile)
        {
            cfbf.out = outFile;
            cfbf.printEnabled = true;
            fprintf(out, "%-8s %10s  %10s    NAME\n", "TYPE", "START SEC", "SIZE");
        }

        //wchar_t* cStrNamePath = L"";
        CString errText;
        BOOL truncate = TRUE;
        std::string errorText;

        size_t outlen = 0;
        //wchar_t* filePathW = UTF8ToUTF16((char*)cfbf.filepath, strlen(cfbf.filepath), &outlen);
        wchar_t* filePathW = cfbf.filepath;

        CStringW fileName;
        FileUtils::CPathStripPath(filePathW, fileName);

        CString cStrNamePath(LR"(C:\Users\tata\Downloads\msg2eml\)");
        cStrNamePath.Append(fileName);
        cStrNamePath.Append(L".eml");

        wchar_t* emlFilePath = (wchar_t*)(LPCWSTR)cStrNamePath;

        HANDLE emlFileHandle = OutlookMessage::FileOpen(emlFilePath, errText, truncate);

         OutlookMsgHelper msgHelper(&cfbf, emlFileHandle, outFile);

        int ret = msgHelper.msg.ParseMsg(&cfbf, ParseOutlookMsg, msgHelper, errorText);
        if ((ret < 0) || (!errorText.empty()))
            _ASSERTE((ret >= 0) && (errorText.empty()));

        if (outFile)
        {
            fprintf(out, "\nOutlookMessage::Print @@@@@@@@@@@@@@@@@@@@@@@@\n");
            msgHelper.msg.Print();
            fprintf(out, "\nOutlookMessage::Print END ###########################################\n");
        }

        std::string msg_utf8;
        int retcode = msgHelper.msg.Msg2Eml(msg_utf8, errorText);
        if ((retcode < 0) || !errorText.empty())
        {
            _ASSERTE(errorText.empty());
        }

        DWORD NumberOfBytesWritten = 0;
        int retCode = OutlookMessage::Write2File(emlFileHandle, msg_utf8.c_str(), msg_utf8.size(), &NumberOfBytesWritten);

        OutlookMessage::FileClose(emlFileHandle);

        int deb = 1;
    }
    else if (print_dir_tree)
    {
        fprintf(out, "%-8s %10s  %10s    NAME\n", "TYPE", "START SEC", "SIZE");

        int ret = cfbf_walk_dir_tree(&cfbf, print_dir_entry, out, errorText);
        if (ret < 0) {
            exit_status = 1;
        }
    }
    else if (walk)
    {
        if (cfbf_walk(&cfbf, out, verbosity, errorText))
            exit_status = 1;
    }
    else if (dump_object_path)
    {
        DirEntry *entry = cfbf_dir_entry_find_path(&cfbf, dump_object_path, errorText);
        if (entry == NULL) {
            error(0, 0, "object \"%s\" not found in %s", dump_object_path, input_filename);
            exit_status = 1;
        }
        else if (entry->object_type == CFBF_ROOT_OBJECT_TYPE) {
            error(0, 0, "you're not allowed to dump the root entry");
            exit_status = 1;
        }
        else if (entry->object_type != CFBF_STREAM_OBJECT_TYPE) {
            error(0, 0, "%s is not a stream object", dump_object_path);
            exit_status = 1;
        }
        else
        {
            int ret;

            std::string name_utf8 = UTF16ToUTF8((wchar_t*)entry->name, entry->name_length);

            std::string errorText;
            ret = cfbf_follow_chain(&cfbf, entry->start_sector,
                    entry->stream_size,
                    entry->stream_size < header->_ulMiniSectorCutoff,
                    write_sector_to_file, out, errorText);
            if (ret) {
                error(0, 0, "failed to read %s", dump_object_path);
                exit_status = 1;
            }
        }
    }
    else if (extract_publisher_text)
    {
        DirEntry *entry = cfbf_dir_entry_find_path(&cfbf, publisher_contents_path, errorText);

        if (entry == NULL) {
            error(0, 0, "Can't extract text: no entry named \"%s\" in directory", publisher_contents_path);
            exit_status = 1;
        }
        else
        {
            void **contents_chain;
            int sector_size, chain_length;

            contents_chain = cfbf_dir_entry_get_sector_ptrs(&cfbf, entry, &chain_length, &sector_size, errorText);
            if (contents_chain == NULL) {
                exit_status = 1;
            }
            else
            {
                struct write_pub_text_state state;

                memset(&state, 0, sizeof(state));

                if (convert_text_to_utf8)
                {
                    state.iconv_desc = 1;  // must be >= 0
                }
                else {
                    state.iconv_desc = (iconv_t) -1;
                }
                state.out = out;

                if (extract_text_from_contents_chain(contents_chain,
                            chain_length, sector_size, entry->stream_size,
                            verbosity, write_publisher_text, &state) < 0) {
                    exit_status = 1;
                }
            }
            free(contents_chain);
        }
    }

    if (out != NULL && out != stdout && out != stderr) {
        if (fclose(out) == EOF) {
            error(0, errno, "%s", output_filename);
            exit_status = 1;
        }
    }

    cfbf_close(&cfbf);

    return exit_status;
}

#endif
