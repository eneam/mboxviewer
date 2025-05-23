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
// PictureCtrl.h
// 
// Author: Tobias Eiseler
//
// Adapted for Windows MBox Viewer by the mboxview development
// Simplified, added re-orientation, added next, previous, rotate, zoom, dragging and print capabilities
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

#if !defined(_PICTURE_CTRL_)
#define _PICTURE_CTRL_

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
	CPictureCtrl(void* pPictureCtrlOwner = 0);

	//Destructor
	~CPictureCtrl(void);

public:

	void SetPictureCtrlOwner(void* pPictureCtrlOwner);

	//Loads an image from a file
	BOOL LoadFromFile(CStringW &szFilePath, Gdiplus::RotateFlipType rotateType, float zoom);

	//Frees the image data
	void FreeData();

	void ReleaseImage();
	void ResetDraggedFlag();
	void ResetDrag();

	BOOL m_bFixOrientation;
	CPoint m_PointDragBegin;
	CPoint m_PointDragEnd;
	int m_deltaDragX;
	int m_deltaDragY;

protected:
	virtual void PreSubclassWindow();

	//Draws the Control
	virtual void DrawItem(LPDRAWITEMSTRUCT lpDrawItemStruct);

	void GetStringProperty(PROPID propid, Gdiplus::Image &image, CString &str);
	BOOL GetIntProperty(PROPID propid, Gdiplus::Image &image, long &integer);

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

	CStringW m_szFilePath;
	float m_Zoom;
	CCPictureCtrlDemoDlg *m_pPictureCtrlOwner;
	BOOL m_bIsDragged;
	int m_sliderRange;
	int m_sliderFreq;

	CRect m_rect; // current static rectangle
	float m_hightZoom;
	float m_widthZoom;


	//Control flag if a pic is loaded
	// TODO: kept from the original code, likely not be needed anymore
	BOOL m_bIsPicLoaded;

	//GDI Plus Token
	ULONG_PTR m_gdiplusToken;
public:
	DECLARE_MESSAGE_MAP()
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnMenuRButtonUp(UINT nPos, CMenu *pMenu);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnRButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint point);
};

#endif // _PICTURE_CTRL_
