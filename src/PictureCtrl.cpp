///////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////
// PictureCtrl.cpp
// 
// Author: Tobias Eiseler
//
// E-Mail: tobias.eiseler@sisternicky.com
//
// Adapted for Windows MBox Viewer by the mboxview development
// Simplified, added next, previous, rotate and zoom capabilities
// 
// Function: A MFC Picture Control to display
//           an image on a Dialog, etc.
//
// https://www.codeproject.com/Articles/24969/An-MFC-picture-control-to-dynamically-show-picture
///////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////

#include "StdAfx.h"
#include "PictureCtrl.h"
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

CPictureCtrl::CPictureCtrl(void)
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
}

CPictureCtrl::~CPictureCtrl(void)
{
	//Tidy up
	FreeData();
	GdiplusShutdown(m_gdiplusToken);
}

BOOL CPictureCtrl::LoadFromFile(CString &szFilePath, Gdiplus::RotateFlipType rotateType, int zoom)
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
	RedrawWindow();

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

	if(m_bIsPicLoaded)
	{
		//Get control measures
		CRect rc;
		this->GetClientRect(&rc);

		Graphics graphics(lpDrawItemStruct->hDC);

		CStringW f(m_szFilePath);
		Image image(f);

		if (m_rotateType != Gdiplus::RotateNoneFlipNone)
			image.RotateFlip(m_rotateType);

		INT h;
		INT w;

		INT ww = rc.right - rc.left;
		INT hh = rc.bottom - rc.top;

		INT iw = image.GetWidth();
		INT ih = image.GetHeight();

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

				this->GetDC()->FillRect(&rc, &CBrush(RGB(0, 0, 0)));
				graphics.DrawImage(&image, posLeft, posTop, w, h);

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

			this->GetDC()->FillRect(&rc, &CBrush(RGB(0, 0, 0)));
			graphics.DrawImage(&image, posLeft, posTop, w, h);
		}
	}
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