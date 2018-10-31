///////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////
// PictureCtrl.cpp
// 
// Author: Tobias Eiseler
//
// E-Mail: tobias.eiseler@sisternicky.com
//
// Adapted for Windows MBox Viewer by the mboxview development
// Simplified, added re-orientation, added next, previous, rotate, zoom and print capabilities
// 
// Function: A MFC Picture Control to display
//           an image on a Dialog, etc.
//
// https://www.codeproject.com/Articles/24969/An-MFC-picture-control-to-dynamically-show-picture
///////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////

#include "StdAfx.h"
#include "resource.h"  
#include "PictureCtrl.h"
#include "CPictureCtrlDemoDlg.h"
#include <GdiPlus.h>
#include <atlimage.h>

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

using namespace Gdiplus;

bool PathExists(LPCSTR file);
BOOL PathFileExist(LPCSTR path);

CPictureCtrl::CPictureCtrl(void *pPictureCtrlOwner)
	:CStatic()
	, m_bIsPicLoaded(FALSE)
	, m_gdiplusToken(0)
	, m_rotateType(Gdiplus::RotateNoneFlipNone)
{
	GdiplusStartupInput gdiplusStartupInput;
	GdiplusStartup(&m_gdiplusToken, &gdiplusStartupInput, NULL);

	m_Zoom = 0;
	m_ZoomMax = 16;
	m_ZoomMaxForCurrentImage = 0;
	m_pPictureCtrlOwner = (CCPictureCtrlDemoDlg*)pPictureCtrlOwner;
	m_cimage = 0;
	m_cBitmap = 0;
	m_graphics = 0;
}

CPictureCtrl::~CPictureCtrl(void)
{
	//Tidy up
	FreeData();
	GdiplusShutdown(m_gdiplusToken);
}

BOOL CPictureCtrl::LoadFromFile(CString &szFilePath, Gdiplus::RotateFlipType rotateType, int zoom, BOOL invalidate)
{
	//Set success error state
	SetLastError(ERROR_SUCCESS);
	FreeData();

	m_szFilePath = szFilePath;
	m_rotateType = rotateType;
	m_Zoom = zoom;

	//Mark as Loaded
	m_bIsPicLoaded = TRUE;

	if (invalidate) {
		Invalidate();
		RedrawWindow();
	}

	return TRUE;
}

void CPictureCtrl::FreeData()
{
	m_bIsPicLoaded = FALSE;
}

void CPictureCtrl::PreSubclassWindow()
{
	CStatic::PreSubclassWindow();
	ModifyStyle(0, SS_OWNERDRAW);
}

void CPictureCtrl::DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct)
{
	//Check if pic data is loaded
	if (!PathFileExist(m_szFilePath)) {
		m_bIsPicLoaded = FALSE;
		return;
	}

	if (m_bIsPicLoaded)
	{
		//Get control measures
		CRect rc;
		this->GetClientRect(&rc);

		Gdiplus::Graphics graphicsDev(lpDrawItemStruct->hDC);
		Gdiplus::Graphics *graphics = &graphicsDev;

		CStringW f(m_szFilePath);

		Image fimage(f);
		Image *image = &fimage;

		CString equipMake;
		GetPropertyEquipMake(*image, equipMake);

		CString equipModel;
		GetPropertyEquipModel(*image, equipModel);

		CString pictureDateTime;
		GetPropertyDateTime(*image, pictureDateTime);

		char *fileName = new char[m_szFilePath.GetLength() + 1];
		strcpy(fileName, (LPCSTR)m_szFilePath);
		PathStripPath(fileName);

		CString sep = " ";
		CString largeSep = "   - ";
		CString windowText = fileName + largeSep + pictureDateTime + sep + equipModel + sep + equipMake;
		delete[] fileName;

		m_pPictureCtrlOwner->SetWindowText(windowText);

		if (m_bFixOrientation)
		{
			m_rotateType = DetermineNewOrientation(*image);
			m_pPictureCtrlOwner->UpdateRotateType(m_rotateType);
			m_bFixOrientation = FALSE;
		}

		if (m_rotateType != Gdiplus::RotateNoneFlipNone)
			image->RotateFlip(m_rotateType);

		INT h;
		INT w;

		INT ww = rc.right - rc.left;
		INT hh = rc.bottom - rc.top;

		INT iw = image->GetWidth();
		INT ih = image->GetHeight();

		INT posLeft = 0;
		INT posTop = 0;

		BOOL done = FALSE;
		if (m_Zoom != 0)
		{
			INT zih = ih * m_Zoom;
			INT ziw = iw * m_Zoom;

			int wZoomMaxForCurrentImage = ww / iw + 1;
			int hZoomMaxForCurrentImage = hh / ih + 1;
			if (hZoomMaxForCurrentImage < wZoomMaxForCurrentImage)
				m_ZoomMaxForCurrentImage = hZoomMaxForCurrentImage;
			else
				m_ZoomMaxForCurrentImage = wZoomMaxForCurrentImage;

			if ((ziw > ww) || (zih > hh)) {
				m_ZoomMaxForCurrentImage = m_Zoom;
			}
			else
			{
				h = zih;
				w = ziw;

				posLeft = (ww - ziw) / 2;
				posTop = (hh - zih) / 2;

				//this->GetDC()->FillRect(&rc, &CBrush(RGB(0, 0, 0)));
				graphics->DrawImage(image, posLeft, posTop, w, h);

				done = TRUE;
			}
		}

		if (done == FALSE)
		{
			DOUBLE hratio = (DOUBLE)hh / ih;  // hh = hratio * ih
			w = (int)(hratio * iw);
			h = (int)(hratio * ih);
			if (w > ww) {

				DOUBLE wratio = (DOUBLE)ww / iw;
				h = (int)(wratio * ih);
				w = (int)(wratio * iw);
			}

			if (w < ww) {
				posLeft = (ww - w) / 2;
			}
			if (h < hh) {
				posTop = (hh - h) / 2;
			}

			//this->GetDC()->FillRect(&rc, &CBrush(RGB(0, 0, 0)));
			graphics->DrawImage(image, posLeft, posTop, w, h);
		}
	}
}

Gdiplus::RotateFlipType CPictureCtrl::DetermineNewOrientation(Gdiplus::Image &image)
{
	Gdiplus::RotateFlipType rotateType = Gdiplus::RotateNoneFlipNone;

	Gdiplus::Status gStatus = Gdiplus::UnknownImageFormat;

	GUID gFormat = Gdiplus::ImageFormatUndefined;
	gStatus = image.GetRawFormat(&gFormat);
	if (gStatus == Gdiplus::Status::Ok)
	{
		if (gFormat == Gdiplus::ImageFormatEXIF)
			int deb = 1;
		else if (gFormat == Gdiplus::ImageFormatTIFF)
			int deb = 1;
		else if (gFormat == Gdiplus::ImageFormatGIF)
			int deb = 1;
		else if (gFormat == Gdiplus::ImageFormatJPEG)
			int deb = 1;
		else if (gFormat == Gdiplus::ImageFormatMemoryBMP)
			int deb = 1;
		else if (gFormat == Gdiplus::ImageFormatPNG)
			int deb = 1;
		else if (gFormat == Gdiplus::ImageFormatBMP)
			int deb = 1;
		else if (gFormat == Gdiplus::ImageFormatEMF)
			int deb = 1;
		else if (gFormat == Gdiplus::ImageFormatWMF)
			int deb = 1;
		else if (gFormat == Gdiplus::ImageFormatUndefined)
			int deb = 1;
		else
			int deb = 1;
	}

	long data[128];
	Gdiplus::PropertyItem* item = (Gdiplus::PropertyItem*)&data;

	PROPID propid;
	UINT propSize;
	Gdiplus::GpStatus stst;

	propid = PropertyTagOrientation;
	propSize = image.GetPropertyItemSize(propid);
	stst = Gdiplus::PropertyNotFound;
	if (gStatus == Gdiplus::Status::Ok) {
		stst = image.GetPropertyItem(propid, propSize, item);
	}

	long orientation = 0;
	if ((gStatus == Gdiplus::Status::Ok) && (stst == Gdiplus::Status::Ok) && (gFormat != Gdiplus::ImageFormatUndefined))
	{
		if (item->type == PropertyTagTypeShort)
			orientation = static_cast<ULONG>(*((USHORT*)item->value));
		else if (item->type == PropertyTagTypeLong)
			orientation = static_cast<ULONG>(*((ULONG*)item->value));

		int deb = 1;


#if 0
		1 - The 0th row is at the top of the visual image, and the 0th column is the visual left side.
			2 - The 0th row is at the visual top of the image, and the 0th column is the visual right side.
			3 - The 0th row is at the visual bottom of the image, and the 0th column is the visual right side.
			4 - The 0th row is at the visual bottom of the image, and the 0th column is the visual left side.
			5 - The 0th row is the visual left side of the image, and the 0th column is the visual top.
			6 - The 0th row is the visual right side of the image, and the 0th column is the visual top.
			7 - The 0th row is the visual right side of the image, and the 0th column is the visual bottom.
			8 - The 0th row is the visual left side of the image, and the 0th column is the visual bottom.
#endif


		switch (orientation)
		{
		case 1:
			rotateType = Gdiplus::RotateNoneFlipNone;
			break;
		case 2:
			rotateType = Gdiplus::RotateNoneFlipX;
			break;
		case 3:
			rotateType = Gdiplus::Rotate180FlipNone;
			break;
		case 4:
			rotateType = Gdiplus::Rotate180FlipX;
			break;
		case 5:
			rotateType = Gdiplus::Rotate90FlipX;
			break;
		case 6:
			rotateType = Gdiplus::Rotate90FlipNone;
			break;
		case 7:
			rotateType = Gdiplus::Rotate270FlipX;
			break;
		case 8:
			rotateType = Gdiplus::Rotate270FlipNone;
			break;
		}
	}
	return rotateType;
}

BOOL CPictureCtrl::OnEraseBkgnd(CDC *pDC)
{
	if(m_bIsPicLoaded)
	{
		//Get control measures
		CRect rc;
		this->GetClientRect(&rc);
		pDC->FillRect(rc, &CBrush(RGB(0, 0, 0)));
		return TRUE;
	}
	else
	{
		return CStatic::OnEraseBkgnd(pDC);
	}
}

void CPictureCtrl::GetStringProperty(PROPID propid, Gdiplus::Image &image, CString &str)
{
	long data[128];
	Gdiplus::PropertyItem* item = (Gdiplus::PropertyItem*)&data;

	UINT propSize;
	Gdiplus::GpStatus stst;

	// PROPID Type 	PropertyTagTypeASCII , Count 	Length of the string including the NULL terminator
	propSize = image.GetPropertyItemSize(propid);
	stst = Gdiplus::PropertyNotFound;
	stst = image.GetPropertyItem(propid, propSize, item);
	if (stst == Gdiplus::Status::Ok)
	{
		int valueLength = item->length - 1;
		if (valueLength <= 0)
			return;
		LPCSTR buf = str.GetBuffer(valueLength);
		memcpy((char*)buf, (LPCSTR)item->value, valueLength); // to porotect if not NULL terminated
		str.ReleaseBuffer(valueLength);
		int deb = 1;
	}
}

void CPictureCtrl::GetPropertyEquipMake(Gdiplus::Image &image, CString &equipMake)
{
	long data[128];
	Gdiplus::PropertyItem* item = (Gdiplus::PropertyItem*)&data;

	PROPID propid;
	UINT propSize;
	Gdiplus::GpStatus stst;

	propid = PropertyTagEquipMake; // Type 	PropertyTagTypeASCII , Count 	Length of the string including the NULL terminator
	propSize = image.GetPropertyItemSize(propid);
	stst = Gdiplus::PropertyNotFound;
	stst = image.GetPropertyItem(propid, propSize, item);
	if (stst == Gdiplus::Status::Ok)
	{
		int valueCount = item->length - 1;
		if (valueCount <= 0)
			return;
		LPCSTR buf = equipMake.GetBuffer(item->length-1);
		memcpy((char*)buf, (LPCSTR)item->value, item->length - 1); // to porotect if not NULL terminated
		equipMake.ReleaseBuffer(item->length - 1);
		int deb = 1; 
	}
}

void CPictureCtrl::GetPropertyEquipModel(Gdiplus::Image &image, CString &equipModel)
{
	PROPID propid = PropertyTagEquipModel; // Type 	PropertyTagTypeASCII , Count 	Length of the string including the NULL terminator
	GetStringProperty(propid, image, equipModel);
}

void CPictureCtrl::GetPropertyDateTime(Gdiplus::Image &image, CString &imageDataTime)
{
	PROPID propid = PropertyTagDateTime; // Type 	PropertyTagTypeASCII , Count 	Length of the string including the NULL terminator
	GetStringProperty(propid, image, imageDataTime);
}

void CPictureCtrl::GetPropertyImageTitle(Gdiplus::Image &image, CString &imageTitle)
{
	PROPID propid = PropertyTagImageTitle; // Type 	PropertyTagTypeASCII , Count 	Length of the string including the NULL terminator
	GetStringProperty(propid, image, imageTitle);
}

void CPictureCtrl::GetPropertyImageDescription(Gdiplus::Image &image, CString &imageDescription)
{
	PROPID propid = PropertyTagImageDescription; // Type 	PropertyTagTypeASCII , Count 	Length of the string including the NULL terminator
	GetStringProperty(propid, image, imageDescription);
}



BEGIN_MESSAGE_MAP(CPictureCtrl, CStatic)
	ON_WM_ERASEBKGND()
END_MESSAGE_MAP()
