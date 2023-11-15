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

// GdiUtils.cpp : implementation file
//
#include "stdafx.h"
#include "GdiUtils.h"
#include "TextUtilsEx.h"
#include "FileUtils.h"

using namespace Gdiplus;

int GdiUtils::AutoRotateImageFromMemory(const char* pData, int nSize, const wchar_t* newImageFile, int &imageWidth, int &imageHeight)
{
	GdiplusStartupInput gdiplusStartupInput;
	ULONG_PTR m_gdiplusToken;
	int retSts = GdiplusStartup(&m_gdiplusToken, &gdiplusStartupInput, NULL);
	if (retSts != Gdiplus::Status::Ok)
		return -1;

	imageWidth = 0;
	imageHeight = 0;

	IStream* pStream = NULL;
	Gdiplus::Image* pImage = NULL;
	if (CreateStreamOnHGlobal(NULL, TRUE, &pStream) == S_OK)
	{
		if (pStream->Write(pData, (ULONG)nSize, NULL) == S_OK)
		{
			pImage = Image::FromStream(pStream);
			if (pImage == 0)
			{
				pStream->Release();
				GdiplusShutdown(m_gdiplusToken);
				return -1;
			}
			else if ((retSts = pImage->GetLastStatus()) != Gdiplus::Status::Ok)
			{
				delete pImage;
				pStream->Release();
				GdiplusShutdown(m_gdiplusToken);
				return -1;
			}
		}

		int widthB = pImage->GetWidth();
		int heightB = pImage->GetHeight();

		imageWidth = 0;
		imageHeight = 0;

		retSts = GdiUtils::AutoRotateImage(pImage, newImageFile, imageWidth, imageHeight);

		int widthA = pImage->GetWidth();
		int heightA = pImage->GetHeight();

		if (retSts != (Status)-1)
		{
			if (widthB != widthA)
				const int deb = 1;
		}

		pStream->Release();
	}
	else
		retSts = -1;

	GdiplusShutdown(m_gdiplusToken);
	return retSts;
}

int GdiUtils::AutoRotateImageFromFile(const wchar_t* imageFile, const wchar_t* newImageFile, int& imageWidth, int& imageHeight)
{
	// Initialize GDI+.
	GdiplusStartupInput gdiplusStartupInput;
	ULONG_PTR gdiplusToken;
	Status retSts = GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);
	if (retSts != Gdiplus::Status::Ok)
		return -1;

	// Get an image from the disk.
	Image* image = new Image(imageFile);

	Status retImage = image->GetLastStatus();
	if (retImage != Ok)
		goto cleanup;

	imageWidth = 0;
	imageHeight = 0;

	int retR = AutoRotateImage(image, newImageFile, imageWidth, imageHeight);

	_ASSERTE((imageWidth > 0) && (imageHeight > 0));

cleanup:
	delete image;
	GdiplusShutdown(gdiplusToken);
	return 0;
}

int GdiUtils::AutoRotateImage(Gdiplus::Image *image, const wchar_t* newImageFile, int& imageWidth, int& imageHeight)
{
	CLSID             encoderClsid;
	EncoderParameters encoderParameters;
	UINT              width;
	UINT              height;
	Status stat = (Status)-1;

   // Determine whether the width and height of the image 
   // are multiples of 16.
   width = image->GetWidth();
   height = image->GetHeight();

   imageWidth = width;
   imageHeight = height;
    
   CString txt = L"The width of the image is %u ";
   if (width%16 == 0)
      txt.Append(L", which is a multiple of 16.\n");
   else
      txt.Append(L", which is not a multiple of 16.\n");

   TRACE(txt, width);

   txt = L"The height of the image is %u ";
   if (height%16 == 0)
	   txt.Append(L", which is a multiple of 16.\n");
   else
	   txt.Append(L", which is not a multiple of 16.\n");

   TRACE(txt, width);

   // Before we call Image::Save, we must initialize an
   // EncoderParameters object. The EncoderParameters object
   // has an array of EncoderParameter objects. In this
   // case, there is only one EncoderParameter object in the array.
   // The one EncoderParameter object has an array of values.
   // In this case, there is only one value (of type ULONG)
   // in the array. We will set that value to EncoderValueTransformRotate90.

   // Determine rotateType, Get the CLSID of the JPEG encoder, get EncoderParameters
   Gdiplus::RotateFlipType rotateType = GdiUtils::DetermineReorder(*image, encoderClsid, encoderParameters);
   if (rotateType == Gdiplus::RotateNoneFlipNone)
	   goto cleanup;

   PROPID propid;
   UINT propSize;
   Gdiplus::GpStatus stst;
   Gdiplus::Status gStatus = Gdiplus::UnknownImageFormat;

   GUID gFormat = Gdiplus::ImageFormatUndefined;
   gStatus = image->GetRawFormat(&gFormat);

   propid = PropertyTagOrientation;
   propSize = image->GetPropertyItemSize(propid);

   long data[128];
   Gdiplus::PropertyItem* propItem = (Gdiplus::PropertyItem*)&data;

   // proper way is to malloc based on propSize and free before return
   // Gdiplus::PropertyItem* propItem = (Gdiplus::PropertyItem*)malloc(propSize);

   stst = Gdiplus::PropertyNotFound;
   if (gStatus == Gdiplus::Status::Ok) {
	   stst = image->GetPropertyItem(propid, propSize, propItem);
   }

   int newRotateType = Gdiplus::RotateNoneFlipNone;
   USHORT rotateTypeShort = (USHORT)newRotateType;
   ULONG rotateTypeLong = (ULONG)newRotateType;

   if ((gStatus == Gdiplus::Status::Ok) && (stst == Gdiplus::Status::Ok))
   {
	   if (propItem->type == PropertyTagTypeShort)
	   {
		   propItem->value = &rotateTypeShort;
	   }
	   else if (propItem->type == PropertyTagTypeLong)
	   {
		   propItem->value = &rotateTypeLong;
	   }
	   else
		   _ASSERTE((propItem->type == PropertyTagTypeLong) && (propItem->type == PropertyTagTypeShort));
   }
   else
	   goto cleanup;

   Status setSstatus = image->SetPropertyItem(propItem);
   if (setSstatus != Ok)
	   goto cleanup;

   // both approches do work
   
#if 1
   stat = image->RotateFlip(rotateType);
   if (stat == Ok)
   {
	   stat = image->Save(newImageFile, &encoderClsid, NULL);
   }
#else
   stat = image->Save(newImageFile, &encoderClsid, &encoderParameters);
#endif

   if (stat == Ok)
	   TRACE(L"%s saved successfully.\n", newImageFile);
   else
	   TRACE(L"%d  Attempt to save %s failed.\n", stat, newImageFile);

   imageWidth = image->GetWidth();
   imageHeight = image->GetHeight();

cleanup:
   // free((char*)propItem);
   return stat;
}

int GdiUtils::GetEncoderClsid(const wchar_t* format, CLSID* pClsid)
{
   UINT  num = 0;          // number of image encoders
   UINT  size = 0;         // size of the image encoder array in bytes

   ImageCodecInfo* pImageCodecInfo = NULL;

   GetImageEncodersSize(&num, &size);
   if(size == 0)
      return -1;  // Failure

   pImageCodecInfo = (ImageCodecInfo*)(malloc(size));
   if(pImageCodecInfo == NULL)
      return -1;  // Failure

   GetImageEncoders(num, size, pImageCodecInfo);

   for(UINT j = 0; j < num; ++j)
   {
      if( wcscmp(pImageCodecInfo[j].MimeType, format) == 0 )
      {
         *pClsid = pImageCodecInfo[j].Clsid;
         free(pImageCodecInfo);
         return j;  // Success
      }    
   }

   free(pImageCodecInfo);
   return -1;  // Failure
}

Gdiplus::RotateFlipType GdiUtils::DetermineReorder(Gdiplus::Image& image, CLSID& Clsid, EncoderParameters &encoderParameters)
{
	static USHORT rotateTypeShort[2];
	static ULONG rotateTypeLong[2];

	Gdiplus::RotateFlipType rotateType = Gdiplus::RotateNoneFlipNone;

	Gdiplus::Status gStatus = Gdiplus::UnknownImageFormat;

	GUID gFormat = Gdiplus::ImageFormatUndefined;
	gStatus = image.GetRawFormat(&gFormat);
	wchar_t* mimeType = L"";
	if (gStatus == Gdiplus::Status::Ok)
	{
		if (gFormat == Gdiplus::ImageFormatEXIF)
		{
			mimeType = L"image/exif";
			int deb = 1;
		}
		else if (gFormat == Gdiplus::ImageFormatTIFF)
		{
			mimeType = L"image/tiff";
			int deb = 1;
		}
		else if (gFormat == Gdiplus::ImageFormatGIF)
		{
			mimeType = L"image/gif";
			int deb = 1;
		}
		else if (gFormat == Gdiplus::ImageFormatJPEG)
		{
			mimeType = L"image/jpeg";
			int deb = 1;
		}
		else if (gFormat == Gdiplus::ImageFormatMemoryBMP)
		{
			mimeType = L"image/bmp";
			int deb = 1;
		}
		else if (gFormat == Gdiplus::ImageFormatPNG)
		{
			mimeType = L"image/png";
			int deb = 1;
		}
		else if (gFormat == Gdiplus::ImageFormatBMP)
		{
			mimeType = L"image/bmp";
			int deb = 1;
		}
		else if (gFormat == Gdiplus::ImageFormatEMF)
		{
			mimeType = L"image/tiff";
			int deb = 1;
		}
		else if (gFormat == Gdiplus::ImageFormatWMF)
		{
			mimeType = L"image/wmf";
			int deb = 1;
		}
		else if (gFormat == Gdiplus::ImageFormatUndefined)
		{
			mimeType = L"";
			int deb = 1;
		}
		else
		{
			mimeType = L"";
			int deb = 1;
		}
	}

	if (mimeType[0] == L'\0')
		return rotateType;

	// Get the CLSID of the encoder.
	if (GdiUtils::GetEncoderClsid(mimeType, &Clsid) < 0)
		return rotateType;

	long data[128];
	Gdiplus::PropertyItem* propItem = (Gdiplus::PropertyItem*)&data;

	// proper way is to malloc based on propSize and free before return
	// Gdiplus::PropertyItem* propItem = (Gdiplus::PropertyItem*)malloc(propSize);

	Gdiplus::GpStatus stst;

	PROPID propid = PropertyTagOrientation;
	UINT propSize = image.GetPropertyItemSize(propid);
	stst = Gdiplus::PropertyNotFound;
	if (gStatus == Gdiplus::Status::Ok)
	{
		stst = image.GetPropertyItem(propid, propSize, propItem);
	}

	long orientation = 0;
	if ((gStatus == Gdiplus::Status::Ok) && (stst == Gdiplus::Status::Ok) && (gFormat != Gdiplus::ImageFormatUndefined))
	{
		if (propItem->type == PropertyTagTypeShort)
			orientation = static_cast<ULONG>(*((USHORT*)propItem->value));
		else if (propItem->type == PropertyTagTypeLong)
			orientation = static_cast<ULONG>(*((ULONG*)propItem->value));

		int deb = 1;

#if 0
			// Current orientation
            // 0th row == T,   0th column == L
            // View after autorotation
			//     T
			// L       R
			//     B
			1 - The 0th row (T) is at the top of the visual image, and the 0th column (L) is the visual left side.
				//     T
				// L       R   --- Gdiplus::RotateNoneFlipNone  --- encoderParameters.Count = 0; 
				//     B

			2 - The 0th row (T) is at the visual top of the image, and the 0th column (L) is the visual right side.
				//     T
				// R       L   ---  Gdiplus::RotateNoneFlipX --- EncoderValueTransformFlipHorizontal
				//     B

			3 - The 0th row (T) is at the visual bottom of the image, and the 0th column (L) is the visual right side.
				//     B
				// R       L   --- Gdiplus::Rotate180FlipNone  --- EncoderValueTransformRotate180
				//     T
			4 - The 0th row (T) is at the visual bottom of the image, and the 0th column (L) is the visual left side.
				//     B
				// L       R    --- Gdiplus::Rotate180FlipX -- EncoderValueTransformFlipVertical or (EncoderValueTransformRotate180 + EncoderValueTransformFlipHorizontal)
				//     T

			5 - The 0th row (T) is the visual left side of the image, and the 0th column (L) is the visual top.
				//     L
				// T       B   --- Gdiplus::Rotate90FlipX  -- EncoderValueTransformRotate90 + EncoderValueTransformFlipHorizontal
				//     R
#if 0
				// From offical doc, it seems to be incorrect or I am confused
			6 - The 0th row (T) is the visual right side of the image, and the 0th column (L) is the visual top.
				//     L
				// B       T     --- Gdiplus::Rotate270FlipNone    -- EncoderValueTransformRotate270
				//     R
#endif
			6 - The 0th row(T) is the visual left!! side of the image, and the 0th column(L) is the visual bottom!!!.
				//     R
				// T       B     --- Gdiplus::Rotate90FlipNone    -- EncoderValueTransformRotate90
				//     L

			7 - The 0th row (T) is the visual right side of the image, and the 0th column (L) is the visual bottom.
				//     R
				// B       T    --- Gdiplus::Rotate270FlipX  --  EncoderValueTransformRotate270 + EncoderValueTransformFlipHorizontal
				//     L
#if 0
				// From offical doc, it seems to be incorrect or I am confused !!!!
			8 - The 0th row (T) is the visual left side of the image, and the 0th column (L) is the visual bottom.
				//     R
				// T       B    --- Gdiplus::Rotate90FlipNone  -- EncoderValueTransformRotate90
				//     L
#endif
			8 - The 0th row(T) is the visual right!!! side of the image, and the 0th column(L) is the visual bottom.
				//     L
				// B       T    --- Gdiplus::Rotate270FlipNone  -- EncoderValueTransformRotate270
				//     R
				
				// FIXME create small program to test rotation case
#endif

			// Determine required rotation
			encoderParameters.Count = 0;
			encoderParameters.Parameter[0].Guid = Gdiplus::EncoderTransformation;
			encoderParameters.Parameter[0].Type = EncoderParameterValueTypeLong;

			switch (orientation)
			{
			case 1:
				rotateType = Gdiplus::RotateNoneFlipNone;
				break;
			case 2:
				rotateType = Gdiplus::RotateNoneFlipX;
				//
				rotateTypeLong[0] = EncoderValueTransformFlipHorizontal;
				encoderParameters.Count = 1;
				encoderParameters.Parameter[0].NumberOfValues = 1;
				encoderParameters.Parameter[0].Value = &rotateTypeLong[0];
				break;
			case 3:
				rotateType = Gdiplus::Rotate180FlipNone;
				//
				rotateTypeLong[0] = EncoderValueTransformRotate180;
				encoderParameters.Count = 1;
				encoderParameters.Parameter[0].NumberOfValues = 1;
				encoderParameters.Parameter[0].Value = &rotateTypeLong[0];
				break;
			case 4:
				rotateType = Gdiplus::Rotate180FlipX;
				//
				rotateTypeLong[0] = EncoderValueTransformRotate180;
				rotateTypeLong[0] = EncoderValueTransformFlipHorizontal;
				encoderParameters.Count = 1;
				encoderParameters.Parameter[0].NumberOfValues = 2;
				encoderParameters.Parameter[0].Value = &rotateTypeLong[0];
				break;
			case 5:
				rotateType = Gdiplus::Rotate90FlipX;
				//
				rotateTypeLong[0] = EncoderValueTransformRotate90;
				rotateTypeLong[1] = EncoderValueTransformFlipHorizontal;
				encoderParameters.Count = 1;
				encoderParameters.Parameter[0].NumberOfValues = 2;
				encoderParameters.Parameter[0].Value = &rotateTypeLong[0];
				break;
			case 6:
				rotateType = Gdiplus::Rotate90FlipNone;
				//
				rotateTypeLong[0] = EncoderValueTransformRotate90;
				encoderParameters.Count = 1;
				encoderParameters.Parameter[0].NumberOfValues = 1;
				encoderParameters.Parameter[0].Value = &rotateTypeLong[0];
				break;
			case 7:
				rotateType = Gdiplus::Rotate270FlipX;
				//
				rotateTypeLong[0] = EncoderValueTransformRotate270;
				rotateTypeLong[1] = EncoderValueTransformFlipHorizontal;
				encoderParameters.Count = 1;
				encoderParameters.Parameter[0].NumberOfValues = 2;
				encoderParameters.Parameter[0].Value = &rotateTypeLong[0];
				break;
			case 8:
				rotateType = Gdiplus::Rotate270FlipNone;
				//
				rotateTypeLong[0] = EncoderValueTransformRotate270;
				encoderParameters.Count = 1;
				encoderParameters.Parameter[0].NumberOfValues = 1;
				encoderParameters.Parameter[0].Value = &rotateTypeLong[0];
				break;

			default: 
				int deb = 1;
				break;
			}
	}
	return rotateType;
}

BOOL GdiUtils::loadImage(BYTE* pData, size_t nSize, CStringW& extensionW, CStringA& extension)
{
	GdiplusStartupInput gdiplusStartupInput;
	ULONG_PTR m_gdiplusToken;
	Status sts = GdiplusStartup(&m_gdiplusToken, &gdiplusStartupInput, NULL);


	extensionW.Empty();
	extension = ".jpg";

	IStream* pStream = NULL;
	if (CreateStreamOnHGlobal(NULL, TRUE, &pStream) == S_OK)
	{
		if (pStream->Write(pData, (ULONG)nSize, NULL) == S_OK)
		{
			Bitmap* pBmp = Bitmap::FromStream(pStream);
			if (pBmp == 0)
			{
				pStream->Release();
				GdiplusShutdown(m_gdiplusToken);
				return FALSE;
			}
			else if (pBmp->GetLastStatus() != Status::Ok)
			{
				delete pBmp;
				pStream->Release();
				GdiplusShutdown(m_gdiplusToken);
				return FALSE;
			}

			Gdiplus::Status gStatus = Gdiplus::UnknownImageFormat;
			GUID gFormat = Gdiplus::ImageFormatUndefined;
			gStatus = pBmp->GetRawFormat(&gFormat);
			if (gStatus == Gdiplus::Status::Ok)
			{
				if (gFormat == Gdiplus::ImageFormatEXIF) {
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatTIFF) {
					extension = ".tiff";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatGIF) {
					extension = ".gif";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatJPEG) {
					extension = ".jpeg";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatMemoryBMP) {
					extension = ".bmp";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatPNG) {
					extension = ".png";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatBMP) {
					extension = ".bmp";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatEMF) {
					extension = ".emf";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatWMF) {
					extension = ".wmf";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatIcon) {
					extension = ".ico";
					int deb = 1;
				}
				else if (gFormat == Gdiplus::ImageFormatUndefined) {
					int deb = 1;
				}
			}
			int deb = 1;
		}
		pStream->Release();
		GdiplusShutdown(m_gdiplusToken);

		DWORD error;
		BOOL ret = TextUtilsEx::Ansi2WStr(extension, extensionW, error);
		return TRUE;
	}
	GdiplusShutdown(m_gdiplusToken);
	return FALSE;
}

int GdiUtils::GetImageSizeFromFile(const wchar_t* imageFile, int& imageWidth, int& imageHeight)
{
	imageWidth = 0;
	imageHeight = 0;

	// Initialize GDI+.
	GdiplusStartupInput gdiplusStartupInput;
	ULONG_PTR gdiplusToken;
	Status retSts = GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);
	if (retSts != Gdiplus::Status::Ok)
		return -1;

	// Get an image from the disk.

	if (!FileUtils::PathFileExist((LPCWSTR)imageFile))
		_ASSERTE(FALSE);

	Image *image = Image::FromFile((LPCWSTR)imageFile);

	Status retImage = image->GetLastStatus();
	if (retImage != Ok)
		goto cleanup;

	imageWidth = image->GetWidth();
	imageHeight = image->GetHeight();

	_ASSERTE((imageWidth > 0) && (imageHeight > 0));

cleanup:
	delete image;
	GdiplusShutdown(gdiplusToken);
	return 0;
}

BOOL GdiUtils::IsSupportedPictureFileExtension(CString& cext)
{
	if ((cext.CompareNoCase(L".png") == 0) ||
		(cext.CompareNoCase(L".jpg") == 0) ||
		(cext.CompareNoCase(L".gif") == 0) ||
		(cext.CompareNoCase(L".pjpg") == 0) ||
		(cext.CompareNoCase(L".jpeg") == 0) ||
		(cext.CompareNoCase(L".pjpeg") == 0) ||
		(cext.CompareNoCase(L".jpe") == 0) ||
		(cext.CompareNoCase(L".bmp") == 0) ||
		(cext.CompareNoCase(L".tif") == 0) ||
		(cext.CompareNoCase(L".tiff") == 0) ||
		(cext.CompareNoCase(L".dib") == 0) ||
		(cext.CompareNoCase(L".jif") == 0) ||
		(cext.CompareNoCase(L".jfif") == 0) ||
		(cext.CompareNoCase(L".emf") == 0) ||
		(cext.CompareNoCase(L".wmf") == 0) ||
		(cext.CompareNoCase(L".heic") == 0) ||
		(cext.CompareNoCase(L".ico") == 0))
	{
		return TRUE;
	}
	else
		return FALSE;
}

BOOL GdiUtils::IsSupportedPictureFileNameExtension(CString& fileName)
{
	PWSTR ext = PathFindExtension(fileName);
	CString cext = ext;

	BOOL ret = GdiUtils::IsSupportedPictureFileExtension(cext);
	return ret;
}