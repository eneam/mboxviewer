///////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////
// PictureCtrl.h
// 
// Author: Tobias Eiseler
//
// Adapted for Windows MBox Viewer by the mboxview development
// Simplified, added next, previous, rotate and zoom capabilities
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

class CPictureCtrl :
	public CStatic
{
public:

	//Constructor
	CPictureCtrl(void);

	//Destructor
	~CPictureCtrl(void);

public:

	//Loads an image from a file
	BOOL LoadFromFile(CString &szFilePath, Gdiplus::RotateFlipType rotateType, int zoom);

	//Frees the image data
	void FreeData();

	int GetZoomMaxForCurrentImage() { return m_ZoomMaxForCurrentImage; }
	void SetZoomMaxForCurrentImage(int zoomMaxForCurrentImage) { m_ZoomMaxForCurrentImage = zoomMaxForCurrentImage; }

protected:
	virtual void PreSubclassWindow();

	//Draws the Control
	virtual void DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct);
	virtual BOOL OnEraseBkgnd(CDC* pDC);

private:

	CString m_szFilePath;
	Gdiplus::RotateFlipType m_rotateType;
	int m_Zoom;
	int m_ZoomMax;
	int m_ZoomMaxForCurrentImage;

	//Control flag if a pic is loaded
	// TODO: kept from the original code, likely not be needed anymore
	BOOL m_bIsPicLoaded;

	//GDI Plus Token
	ULONG_PTR m_gdiplusToken;
};
