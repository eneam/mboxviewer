//
// RtfDecompressor.cs
// https://github.com/Sicos1977/MSGReader/blob/master/MsgReader/Outlook/RtfDecompressor.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2025 Kees van Spelde. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
// 
//  Ported to c++ by Zbigniew Minciel to integrate with free Windows MBox Viewer
//

#pragma once

#include <stdio.h>

/// <summary>
///     Class used to decompress compressed RTF
/// </summary>
class RtfDecompressor
{
public:

    /// <summary>
    ///     Calculates the CRC32 of the given bytes.
    ///     The CRC32 calculation is similar to the standard one as demonstrated in RFC 1952,
    ///     but with the inversion (before and after the calculation) omitted.
    /// </summary>
    /// <param name="buf">The byte array to calculate CRC32 on </param>
    /// <param name="off">The offset within buf at which the CRC32 calculation will start </param>
    /// <param name="len">The number of bytes on which to calculate the CRC32</param>
    /// <returns>The CRC32 value</returns>
    static int CalculateCrc32(unsigned char* buf, int off, int len);

    /// <summary>
    ///     Returns an unsigned 32-bit value from little-endian ordered bytes.
    /// </summary>
    /// <param name="buf">Byte array from which byte values are taken</param>
    /// <param name="offset">Offset the offset within buf from which byte values are taken</param>
    /// <returns>An unsigned 32-bit value as a lon</returns>
    static long GetU32(unsigned char* buf, int offset);

    /// <summary>
    ///     Returns an unsigned 8-bit value from a byte array.
    /// </summary>
    /// <param name="buf">A byte array from which byte value is taken</param>
    /// <param name="offset">The offset within buf from which byte value is taken</param>
    /// <returns>An unsigned 8-bit value as an int</returns>
    static int GetU8(unsigned char* buf, int offset);

    /// <summary>
    ///     Decompresses the RTF or returns zero if src does
    ///     not contain valid compressed-RTF bytes.
    /// </summary>
    /// <param name="src">Src the compressed-RTF data bytes</param>
    /// <returns>Length of decompressed buffer. The *dest buffer will contain the decompressed bytes</returns>
    /// Caller should initialize *dest to NULL
    /// Caller is responsible for calling "free *dest"

    static int DecompressRtf(FILE* out, unsigned char* src, int srclen, unsigned char** dest, char* errorText, int errorTextLen);

};
