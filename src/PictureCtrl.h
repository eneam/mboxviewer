///////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////
// PictureCtrl.h
// 
// Author: Tobias Eiseler
//
// Adapted for Windows MBox Viewer by the mboxview development
// Simplified, added re-orientation, added next, previous, rotate, zoom and print capabilities
// TODO: Resizing by Mouse Move can be slow for large images
//
// E-Mail: tobias.eiseler@sisternicky.com
// 
// Function: A MFC Picture Control to display
//           an image on a Dialog, etc.
//
// https://www.codeproject.com/Articles/24969/An-MFC-picture-control-to-dynamically-show-picture
///////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////

#pragma once
#include "afxstr.h"
#include "afxwin.h"
#include "gdiplusImaging.h"
#include <GdiPlus.h>
#include <atlimage.h>

class CCPictureCtrlDemoDlg;

class CPictureCtrl :
	public CStatic
{
public:

	//Constructor
	CPictureCtrl(void *pPictureCtrlOwner);

	//Destructor
	~CPictureCtrl(void);

public:

	//Loads an image from a file
	BOOL LoadFromFile(CString &szFilePath, Gdiplus::RotateFlipType rotateType, int zoom, BOOL invalidate);

	//Frees the image data
	void FreeData();

	int GetZoomMaxForCurrentImage() { return m_ZoomMaxForCurrentImage; }
	void SetZoomMaxForCurrentImage(int zoomMaxForCurrentImage) { m_ZoomMaxForCurrentImage = zoomMaxForCurrentImage; }

	void ReleaseImage();

	BOOL m_bFixOrientation;

protected:
	virtual void PreSubclassWindow();

	//Draws the Control
	virtual void DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct);
	Gdiplus::RotateFlipType DetermineNewOrientation(Gdiplus::Image &image);

	void GetStringProperty(PROPID propid, Gdiplus::Image &image, CString &str);
	void GetPropertyEquipMake(Gdiplus::Image &image, CString &equipMake);
	void GetPropertyEquipModel(Gdiplus::Image &image, CString &equipModel);
	void GetPropertyDateTime(Gdiplus::Image &image, CString &imageDataTime);
	void GetPropertyImageTitle(Gdiplus::Image &image, CString &imageTitle);
	void GetPropertyImageDescription(Gdiplus::Image &image, CString &imageDescription);

	Gdiplus::RotateFlipType m_rotateType;
	Gdiplus::Image *m_cimage; 
	Gdiplus::CachedBitmap  *m_cBitmap; // not used yet, may help to reduce flicker ??
	Gdiplus::Graphics *m_graphics;  // not used yet

private:

	CString m_szFilePath;
	int m_Zoom;
	int m_ZoomMax;
	int m_ZoomMaxForCurrentImage; 
	CCPictureCtrlDemoDlg *m_pPictureCtrlOwner;


	//Control flag if a pic is loaded
	// TODO: kept from the original code, likely not be needed anymore
	BOOL m_bIsPicLoaded;

	//GDI Plus Token
	ULONG_PTR m_gdiplusToken;
public:
	DECLARE_MESSAGE_MAP()
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
};
