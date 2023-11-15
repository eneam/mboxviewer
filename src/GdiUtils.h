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

#include "stdafx.h"

class GdiUtils
{
public:

	static int AutoRotateImageFromFile(const wchar_t* imageFile, const wchar_t* newImageFile, int& imageWidth, int& imageHeight);
	static int AutoRotateImageFromMemory(const char* pData, int nSize, const wchar_t* newImageFile, int& imageWidth, int& imageHeight);
	static int AutoRotateImage(Gdiplus::Image* image, const wchar_t* newImageFile, int& imageWidth, int& imageHeight);

	static BOOL loadImage(BYTE* pData, size_t nSize, CStringW& extensionW, CStringA& extension);

	static Gdiplus::RotateFlipType DetermineReorder(Gdiplus::Image& image, CLSID& Clsid, Gdiplus::EncoderParameters& encoderParameters);
	static int GetEncoderClsid(const wchar_t* format, CLSID* pClsid);

	static int GetImageSizeFromFile(const wchar_t* imageFile, int& imageWidth, int& imageHeight);

	static BOOL IsSupportedPictureFileExtension(CString& cext);
	static BOOL IsSupportedPictureFileNameExtension(CString& fileName);

protected:

};