//
//////////////////////////////////////////////////////////////////
//
//  Windows Mbox Viewer is a free tool to view, search and print mbox mail archives.
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

// MyCButton.cpp : implementation file
// 
// You can set/change the background color
//



#include "stdafx.h"
#include "MyCButton.h"
#include "MainFrm.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

#define BLACK      RGB(0,0,0)
#define YELLOW     RGB(255, 255, 0)
#define RED        RGB(127,0,0)

BEGIN_MESSAGE_MAP(MyCButton, CButton)
	ON_WM_CTLCOLOR_REFLECT()
END_MESSAGE_MAP()

MyCButton::MyCButton()
{

	m_crBkColor = ::GetSysColor(COLOR_3DFACE); // Initializing background color to the system face color.
	//m_crBkColor = YELLOW;
	m_crTextColor = CMainFrame::m_bDarkTheme ? COLOR_DARK_TEXT : BLACK;
	m_brBkgnd.CreateSolidBrush(m_crBkColor); // Creating the Brush Color For the Edit Box Background
}

void MyCButton::SetTextColor(COLORREF crColor)
{
	m_crTextColor = crColor; // Passing the value passed by the dialog to the member varaible for Text Color
	RedrawWindow();
}

void MyCButton::SetBkColor(COLORREF crColor)
{
	m_crBkColor = crColor; // Passing the value passed by the dialog to the member varaible for Backgound Color
	m_brBkgnd.DeleteObject(); // Deleting any Previous Brush Colors if any existed.
	m_brBkgnd.CreateSolidBrush(crColor); // Creating the Brush Color For the Edit Box Background
	RedrawWindow();
}

HBRUSH MyCButton::CtlColor(CDC* pDC, UINT nCtlColor)
{
	HBRUSH hbr;

	hbr = (HBRUSH)m_brBkgnd; // Passing a Handle to the Brush
	pDC->SetBkColor(m_crBkColor); // Setting the Color of the Text Background to the one passed by the Dialog
	BOOL darkBackground = (GetRValue(m_crBkColor) + GetGValue(m_crBkColor) + GetBValue(m_crBkColor)) < 384;
	pDC->SetTextColor(CMainFrame::m_bDarkTheme ?
		(darkBackground ? COLOR_DARK_TEXT : RGB(0, 0, 0)) : m_crTextColor);

	return hbr;
}
