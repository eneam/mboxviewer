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
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
//  Library General Public License for more details.
//
//  You should have received a copy of the GNU Library General Public
//  License along with this program; if not, write to the
//  Free Software Foundation, Inc., 51 Franklin St, Fifth Floor,
//  Boston, MA  02110 - 1301, USA.
//
//////////////////////////////////////////////////////////////////
//

// Splitter.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// C4WaySplitter window

#ifndef __SPLITTER_H__
#define __SPLITTER_H__

class CNSplit : public CSplitterWnd
{
// Construction
public:
	CNSplit();

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(C4WaySplitter)
	virtual void OnDrawSplitter(CDC* pDC, ESplitType nType, const CRect& rect);
	virtual void DeleteView(int row, int col);
	//}}AFX_VIRTUAL

// Implementation
public:
	bool bfirst;
	int m_pos;
	virtual ~CNSplit();

	// Generated message map functions
protected:
//	afx_msg void OnLButtonDblClk(UINT nFlags, CPoint pt);
	//{{AFX_MSG(C4WaySplitter)
	afx_msg void OnPaint();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
private:
	COLORREF m_clrBtnHLit;
	COLORREF m_clrBtnShad;
	COLORREF m_clrBtnFace;
};

/////////////////////////////////////////////////////////////////////////////
#endif
