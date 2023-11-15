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

///////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////
// PictureCtrl.cpp
// 
// Author: Tobias Eiseler
//
// E-Mail: tobias.eiseler@sisternicky.com
//
// Adapted for Windows MBox Viewer by the mboxview development
// Simplified, added re-orientation, added next, previous, rotate, zoom, dragging and print capabilities
// Significant portion of the code likely more than 80% is now new
// 
// Function: A MFC Picture Control to display
//           an image on a Dialog, etc.
//
// https://www.codeproject.com/Articles/24969/An-MFC-picture-control-to-dynamically-show-picture
///////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////

#include "StdAfx.h"
#include <GdiPlus.h>
#include <atlimage.h>
#include "resource.h"  
#include "PictureCtrl.h"
#include "FileUtils.h"
#include "TextUtilsEx.h"
#include "CPictureCtrlDemoDlg.h"
#include "GdiUtils.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

using namespace Gdiplus;

CPictureCtrl::CPictureCtrl(void *pPictureCtrlOwner)
	:CStatic()
	, m_bIsPicLoaded(FALSE)
	, m_gdiplusToken(0)
	, m_rotateType(Gdiplus::RotateNoneFlipNone)
{
	GdiplusStartupInput gdiplusStartupInput;
	GdiplusStartup(&m_gdiplusToken, &gdiplusStartupInput, NULL);

	m_Zoom = 1;
	m_pPictureCtrlOwner = (CCPictureCtrlDemoDlg*)pPictureCtrlOwner;
	m_cimage = 0;
	m_cBitmap = 0;
	m_graphics = 0;
	m_bIsDragged = FALSE;

	m_PointDragBegin.SetPoint(0, 0);;
	m_PointDragEnd.SetPoint(0, 0);
	m_deltaDragX = 0;
	m_deltaDragY = 0;
}

CPictureCtrl::~CPictureCtrl(void)
{
	//Tidy up
	FreeData();
	delete m_cimage;
	GdiplusShutdown(m_gdiplusToken);
}

BOOL CPictureCtrl::LoadFromFile(CStringW &szFilePath, Gdiplus::RotateFlipType rotateType, float zoom)
{
	//Set success error state
	SetLastError(ERROR_SUCCESS);
	FreeData();

	m_szFilePath = szFilePath;
	m_rotateType = rotateType;
	m_Zoom = zoom;

	//Mark as Loaded
	m_bIsPicLoaded = TRUE;

	Invalidate();

	return TRUE;
}

void CPictureCtrl::FreeData()
{
	m_bIsPicLoaded = FALSE;
}

void CPictureCtrl::PreSubclassWindow()
{
	CStatic::PreSubclassWindow();
	ModifyStyle(0, SS_OWNERDRAW | SS_NOTIFY);
}

void CPictureCtrl::DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct)
{
	//Check if pic data is loaded
	if (!PathFileExistsW(m_szFilePath)) {
		m_bIsPicLoaded = FALSE;
		return;
	}

	//TRACE(L"CPictureCtrl::DrawItem\n");

	if (m_bIsPicLoaded)
	{
		CRect rc;
		this->GetClientRect(&rc);
		if ((rc.Width() <= 0) || (rc.Height() <= 0))
			return;
#if 0
		if (rc.Width() <= 0)
		{
			CPoint bottomRight(rc.right+1, rc.bottom);
			rc.SetRect(rc.TopLeft(), bottomRight);
		}
		if (rc.Height() <= 0)
		{
			CPoint bottomRight(rc.right, rc.bottom+1);
			rc.SetRect(rc.TopLeft(), bottomRight);
		}
#endif

		Gdiplus::Graphics graphicsDev(lpDrawItemStruct->hDC);
		Gdiplus::Graphics *graphics = &graphicsDev;

		CStringW f(m_szFilePath);

		if (m_cimage == 0) {
			m_cimage = Image::FromFile(f);
		}

		Status sts = m_cimage->GetLastStatus();
		if (sts != Gdiplus::Status::Ok)
		{
			return;
		}

		if (m_cimage == 0)
			int deb = 1;

		Image *image = m_cimage;

		Bitmap bmp(rc.right, rc.bottom);
		Gdiplus::Graphics gr(&bmp);
		Gdiplus::Graphics *graph = &gr;

		// rectangle for graphics graph
		Gdiplus::Rect rC((int)rc.left, (int)rc.top, (int)rc.Width(), (int)rc.Height());
		
		CString equipMake;
		GetPropertyEquipMake(*image, equipMake);

		CString equipModel;
		GetPropertyEquipModel(*image, equipModel);

		CString pictureDateTime;
		GetPropertyDateTime(*image, pictureDateTime);

		CStringW windowTextW;
		FileUtils::CPathStripPath(m_szFilePath, windowTextW);

#if 0
		CString windowText;
		DWORD error;
		TextUtilsEx::WStr2Ansi(windowTextW, windowText, error);
#endif

		CString sep = L" ";
		CString largeSep = L"   - ";

		if (!(pictureDateTime.IsEmpty() && equipModel.IsEmpty() && equipMake.IsEmpty()))
			windowTextW += largeSep + pictureDateTime + sep + equipModel + sep + equipMake;

		m_pPictureCtrlOwner->SetWindowText(windowTextW);

#if 0
		// Didn't work. Looks may need to handle WM_NCPAINT. Below seem to be overwritten by WM_NCPAINT event
		// Revisit later
		TITLEBARINFO ti;
		memset(&ti, 0, sizeof(ti));
		ti.cbSize = sizeof(ti);
		m_pPictureCtrlOwner->GetTitleBarInfo(&ti);

		int xpos = 5;
		int ypos = 200;
		HDC hDC = lpDrawItemStruct->hDC;
		BOOL retW = ::ExtTextOutW(hDC, xpos, ypos, ETO_CLIPPED, &ti.rcTitleBar, (LPCWSTR)windowTextW, windowTextW.GetLength(), NULL);
#endif


		if (m_bFixOrientation)
		{
			CLSID Clsid;
			EncoderParameters encoderParameters;
			m_rotateType = GdiUtils::DetermineReorder(*image, Clsid, encoderParameters);
			image->RotateFlip(m_rotateType);
			m_rotateType = Gdiplus::RotateNoneFlipNone;
		}

		if (m_rotateType != Gdiplus::RotateNoneFlipNone) {
			image->RotateFlip(m_rotateType);
			m_rotateType = Gdiplus::RotateNoneFlipNone;
		}

		m_bFixOrientation = FALSE;

		INT h;
		INT w;

		INT ww = rc.right - rc.left; // rc.Width()
		INT hh = rc.bottom - rc.top; // rc.Height()

		INT iw = image->GetWidth();  // iw - image width
		INT ih = image->GetHeight();  // ih - image hight

		if ((iw <= 0) || (ih <= 0))
			return;

		INT posLeft = 0;
		INT posTop = 0;

		DOUBLE hratio = (DOUBLE)hh / ih;  // hh = hratio * ih
		w = (int)(hratio * iw);
		h = (int)(hratio * ih);
		if (w > ww) {  
			// image width > client rectagle width , adjust image size
			DOUBLE wratio = (DOUBLE)ww / iw;
			h = (int)(wratio * ih);
			w = (int)(wratio * iw);
		}

		// Keep original dimensions if possible
		if ((iw <= ww) && (ih <= hh))
		{
			w = iw;
			h = ih;
		}

		if (m_Zoom == 1)
		{
			if (w < ww) {
				posLeft = (ww - w) / 2;
			}
			if (h < hh) {
				posTop = (hh - h) / 2;
			}
		}
		else if (m_Zoom != 1)
		{

			// Adjust m_deltaDragX and m_deltaDragY. TODO: needs some work but good enough for now.
			if (m_rect.Height() && m_rect.Width() && rc.Height() && rc.Width())
			{
				m_hightZoom = (float)rc.Height() / m_rect.Height();
				m_widthZoom = (float)rc.Width() / m_rect.Width();

				m_hightZoom = abs(m_hightZoom);
				m_widthZoom = abs(m_widthZoom);

				m_deltaDragX = (int)(m_widthZoom*m_deltaDragX);
				m_deltaDragY = (int)(m_hightZoom*m_deltaDragY);
			}
			if (rc.Height() && rc.Width())
				m_rect = rc;

			//TRACE(L"m_hightZoom=%f m_widthZoom=%f\n", m_hightZoom, m_widthZoom);

			w = (int)(m_Zoom * w);
			h = (int)(m_Zoom * h);

			posLeft = (ww - w) / 2;
			posTop = (hh - h) / 2;

			if (m_Zoom < 1) {
				m_deltaDragX = 0;
				m_deltaDragY = 0;
			}
			else if (m_Zoom > 1)
			{
				posLeft += m_deltaDragX;
				posTop += m_deltaDragY;

				int posRight; posRight = posLeft + w;
				int posBottom; posBottom = posTop + h;

				// Likely logic can be reduced :)
				if (h > hh)  // zommed image hight > client rectagle hight
				{
					// make sure posTop <= 0 and posBottom >= h
					if (posTop > 0) {
						m_deltaDragY -= posTop;
						posTop = 0;
					}
					else if (posBottom < hh) {
						m_deltaDragY += (hh - posBottom);
						posTop += (hh - posBottom);
					}
				}
				else {
					// make sure posTop >= 0 or posBottom <= h
					if (posTop < 0) {
						m_deltaDragY -= posTop;
						posTop = 0;
					}
					else if (posBottom > hh) {
						m_deltaDragY += (hh - posBottom);
						posTop += (hh - posBottom);
					}
				}

				if (w > ww) // zommed image width > client rectagle width
				{
					// make sure posLeft <= 0 and posRight >= w
					if (posLeft > 0) {
						m_deltaDragX -= posLeft;
						posLeft = 0;
					}
					else if (posRight < ww) {
						m_deltaDragX += (ww - posRight);
						posLeft += (ww - posRight);
					}
				}
				else
				{
					// make sure posLeft >= 0 or posRight <= w
					if (posLeft < 0) {
						m_deltaDragX -= posLeft;
						posLeft = 0;
					}
					else if (posRight > ww) {
						m_deltaDragX += (ww - posRight);
						posLeft += (ww - posRight);
					}
				}
			}
		}

		// All below approches work.
		// ExcludeClip would suggest better performance but no significant boost observed
		// TODO: Method 1 & 2 doesn't seem to work in all cases. 
		// Sometimnes new image overlapps old/previous image . Transparent background issue ???
#if 0
		// Method 1
		Gdiplus::SolidBrush bb(Gdiplus::Color(0, 0, 0));
		graphics->DrawImage(image, posLeft, posTop, w, h);
		RectF excludeRect((REAL)posLeft, (REAL)posTop, (REAL)w, (REAL)h);
		graphics->ExcludeClip(excludeRect);
		graphics->FillRectangle(&bb, rC);
#endif

#if 0
		// Method 2
		Gdiplus::SolidBrush bb(Gdiplus::Color(0, 0, 0));
		graph->DrawImage(image, posLeft, posTop, w, h);
		RectF excludeRect((REAL)posLeft, (REAL)posTop, (REAL)w, (REAL)h);
		graph->ExcludeClip(excludeRect);
		graph->FillRectangle(&bb, rC);
		graphics->DrawImage(&bmp, rc.left, rc.top, rc.Width(), rc.Height());
#endif

#if 1
		// Method 3 - initial implementation
		Gdiplus::SolidBrush bb(Gdiplus::Color(0, 0, 0));   // black
		Gdiplus::SolidBrush wb(Gdiplus::Color(255, 255, 255));   // white
		graph->FillRectangle(&bb, rC);
		Gdiplus::Rect rcplus(posLeft, posTop, w, h);
		graph->FillRectangle(&wb, rcplus);
		graph->DrawImage(image, posLeft, posTop, w, h);
		graphics->DrawImage(&bmp, rc.left, rc.top, rc.Width(), rc.Height());  // FIXME
		//graph->FillRectangle(&bb, rC);
#endif
	}
}

BOOL CPictureCtrl::OnEraseBkgnd(CDC *pDC)
{
	return TRUE;
}

void CPictureCtrl::GetStringProperty(PROPID propid, Gdiplus::Image &image, CString &str)
{
	long data[1024];
	Gdiplus::PropertyItem* item = (Gdiplus::PropertyItem*)&data;

	UINT propSize;
	Gdiplus::GpStatus stst;

	// PROPID Type 	PropertyTagTypeASCII , Count 	Length of the string including the NULL terminator
	propSize = image.GetPropertyItemSize(propid);
	stst = Gdiplus::PropertyNotFound;
	stst = image.GetPropertyItem(propid, propSize, item);
	if (stst == Gdiplus::Status::Ok)
	{
		// item->type = ImageTypeBitmap or ImageTypeMetafile
		int valueLength = item->length - 1;
		if (valueLength <= 0)
			return;

		CStringA val = (LPCSTR)item->value;
		DWORD error;
		BOOL ret = TextUtilsEx::CodePage2WStr(&val, 0, &str, error);
		int deb = 1;
	}
}

BOOL CPictureCtrl::GetIntProperty(PROPID propid, Gdiplus::Image &image, long &integer)
{
	long data[1024];
	Gdiplus::PropertyItem* item = (Gdiplus::PropertyItem*)&data;

	UINT propSize;
	Gdiplus::GpStatus stst;

	integer = 0;

	propSize = image.GetPropertyItemSize(propid);
	stst = Gdiplus::PropertyNotFound;
	stst = image.GetPropertyItem(propid, propSize, item);

	if (stst == Gdiplus::Status::Ok)
	{
		if (item->type == PropertyTagTypeByte)
			integer = *((BYTE*)item->value);
		else if (item->type == PropertyTagTypeShort)
			integer = *((SHORT*)item->value);
		else if (item->type == PropertyTagTypeLong)
			integer = *((LONG*)item->value);
		else
			return FALSE;

		int deb = 1;
		return TRUE;
	}
	else
		return FALSE;
}

void CPictureCtrl::GetPropertyEquipMake(Gdiplus::Image &image, CString &equipMake)
{
	PROPID propid = PropertyTagEquipMake; // Type 	PropertyTagTypeASCII , Count 	Length of the string including the NULL terminator
	GetStringProperty(propid, image, equipMake);
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

void CPictureCtrl::ReleaseImage()
{
	if (m_cimage) 
	{
		delete m_cimage;
		m_cimage = 0;
	}
}



BEGIN_MESSAGE_MAP(CPictureCtrl, CStatic)
	ON_WM_ERASEBKGND()
	//ON_WM_NCMOUSEMOVE()
	ON_WM_MOUSEMOVE()
	//ON_WM_MENURBUTTONUP()
	ON_WM_LBUTTONDOWN()
	ON_WM_LBUTTONUP()
	//ON_WM_RBUTTONDOWN()
	//ON_WM_RBUTTONUP()
	ON_WM_MOUSEWHEEL()
	//ON_WM_LBUTTONDBLCLK()
END_MESSAGE_MAP()


void CPictureCtrl::OnMouseMove(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default

	if ((m_bIsDragged == TRUE) && (m_Zoom > 1))
	{
		m_PointDragEnd = point;

		int deltaDragX = m_PointDragEnd.x - m_PointDragBegin.x;
		int deltaDragY = m_PointDragEnd.y - m_PointDragBegin.y;

		m_deltaDragX += deltaDragX;
		m_deltaDragY += deltaDragY;

		m_PointDragBegin = m_PointDragEnd;

		Invalidate();
	}

	CStatic::OnMouseMove(nFlags, point);
}


void CPictureCtrl::OnMenuRButtonUp(UINT nPos, CMenu *pMenu)
{
	// This feature requires Windows 2000 or greater.
	// The symbols _WIN32_WINNT and WINVER must be >= 0x0500.
	// TODO: Add your message handler code here and/or call default

	CStatic::OnMenuRButtonUp(nPos, pMenu);
}


void CPictureCtrl::OnLButtonDown(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default
	m_bIsDragged = TRUE;

	m_PointDragBegin = point;
	m_PointDragEnd.SetPoint(0, 0);

	CStatic::OnLButtonDown(nFlags, point);
}


void CPictureCtrl::OnLButtonUp(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default
	m_bIsDragged = FALSE;
	m_PointDragEnd = point;

	int deltaDragX = m_PointDragEnd.x - m_PointDragBegin.x;
	int deltaDragY = m_PointDragEnd.y - m_PointDragBegin.y;

	m_deltaDragX += deltaDragX;
	m_deltaDragY += deltaDragY;

	m_PointDragBegin = m_PointDragEnd;

	CStatic::OnLButtonUp(nFlags, point);
}


void CPictureCtrl::OnRButtonDown(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default

	CStatic::OnRButtonDown(nFlags, point);
}


void CPictureCtrl::OnRButtonUp(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default

	CStatic::OnRButtonUp(nFlags, point);
}

void CPictureCtrl::ResetDraggedFlag()
{
	if (m_bIsDragged)
		int deb = 1;
	m_bIsDragged = FALSE;
}

void CPictureCtrl::ResetDrag()
{
	m_deltaDragX = 0;
	m_deltaDragY = 0;
	m_rect.SetRectEmpty();
}

BOOL CPictureCtrl::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt)
{
	// TODO: Add your message handler code here and/or call default

	if ((m_bIsDragged == FALSE) && (m_Zoom > 1))
	{
		m_deltaDragX += 0;
		m_deltaDragY += zDelta;

		Invalidate();
	}

	return CStatic::OnMouseWheel(nFlags, zDelta, pt);
}


void CPictureCtrl::OnLButtonDblClk(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default

	CStatic::OnLButtonDblClk(nFlags, point);
}



