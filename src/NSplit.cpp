// Splitter.cpp : implementation file
//

#include "stdafx.h"
#include "NSplit.h"

//#include "D:\\Programmi\\Microsoft Visual Studio\\VC98\\MFC\\SRC\\afximpl.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

#define CX_BORDER   1
#define CY_BORDER   1


// HitTest return values (values and spacing between values is important)
enum HitTestValue
{
	noHit                   = 0,
	vSplitterBox            = 1,
	hSplitterBox            = 2,
	bothSplitterBox         = 3,        // just for keyboard
	vSplitterBar1           = 101,
	vSplitterBar15          = 115,
	hSplitterBar1           = 201,
	hSplitterBar15          = 215,
	splitterIntersection1   = 301,
	splitterIntersection225 = 525
};

/////////////////////////////////////////////////////////////////////////////
// CNSplit

CNSplit::CNSplit()
{
	m_clrBtnHLit  = ::GetSysColor(COLOR_BTNHILIGHT);
	m_clrBtnShad  = ::GetSysColor(COLOR_BTNSHADOW);
	m_clrBtnFace  = ::GetSysColor(COLOR_BTNFACE);
	m_pos = 0;
	bfirst = true;
}

CNSplit::~CNSplit()
{
}


BEGIN_MESSAGE_MAP(CNSplit, CSplitterWnd)
	//{{AFX_MSG_MAP(CNSplit)
	ON_WM_PAINT()
//	ON_WM_LBUTTONDBLCLK()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CNSplit message handlers


void CNSplit::OnDrawSplitter(CDC* pDC, ESplitType nType,
	const CRect& rectArg)
{
	// if pDC == NULL, then just invalidate
	if (pDC == NULL)
	{
		RedrawWindow(rectArg, NULL, RDW_INVALIDATE|RDW_NOCHILDREN);
		return;
	}
	ASSERT_VALID(pDC);

	// otherwise, actually draw
	CRect rect = rectArg;
	switch (nType)
	{
	case splitBorder:
		pDC->Draw3dRect(rect, m_clrBtnFace, m_clrBtnFace);
		rect.InflateRect(-CX_BORDER, -CY_BORDER);
		pDC->Draw3dRect(rect, m_clrBtnShad, m_clrBtnHLit);
		return;
	case splitIntersection:
		break;

	case splitBox:
		break;

	case splitBar:
		break;

	default:
		ASSERT(FALSE);  // unknown splitter type
	}
	// fill the middle
	pDC->FillSolidRect(rect, m_clrBtnFace);

}

void CNSplit::OnPaint() 
{
	ASSERT_VALID(this);
	CPaintDC dc(this);
	if( bfirst )
		bfirst = false;
	CRect rectClient;
	GetClientRect(&rectClient);
	rectClient.InflateRect(-m_cxBorder, -m_cyBorder);

	CRect rectInside;
	GetInsideRect(rectInside);

	// draw the splitter boxes
	if (m_bHasVScroll && m_nRows < m_nMaxRows)
	{
		OnDrawSplitter(&dc, splitBox,
			CRect(rectInside.right, rectClient.top,
				rectClient.right, rectClient.top + m_cySplitter));
	}

	if (m_bHasHScroll && m_nCols < m_nMaxCols)
	{
		OnDrawSplitter(&dc, splitBox,
			CRect(rectClient.left, rectInside.bottom,
				rectClient.left + m_cxSplitter, rectClient.bottom));
	}

	// extend split bars to window border (past margins)
	DrawAllSplitBars(&dc, rectInside.right, rectInside.bottom);

	// draw splitter intersections (inside only)
	GetInsideRect(rectInside);
	dc.IntersectClipRect(rectInside);
	CRect rect;
	rect.top = rectInside.top;
	for (int row = 0; row < m_nRows - 1; row++)
	{
		rect.top += m_pRowInfo[row].nCurSize + m_cyBorderShare;
		rect.bottom = rect.top + m_cySplitter;
		rect.left = rectInside.left;
		for (int col = 0; col < m_nCols - 1; col++)
		{
			rect.left += m_pColInfo[col].nCurSize + m_cxBorderShare;
			rect.right = rect.left + m_cxSplitter;
			OnDrawSplitter(&dc, splitIntersection, rect);
			rect.left = rect.right + m_cxBorderShare;
		}
		rect.top = rect.bottom + m_cxBorderShare;
	}
}

void CNSplit::DeleteView(int row, int col)
{
	ASSERT_VALID(this);
	
	// if active child is being deleted - activate next
	CWnd* pPane = GetPane(row, col);

	if( pPane == NULL )
		return;
	if (GetActivePane() == pPane)
		ActivateNext(FALSE);

	// default implementation assumes view will auto delete in PostNcDestroy
	pPane->DestroyWindow();
}

//void CNSplit::OnLButtonDblClk(UINT /*nFlags*/, CPoint pt)
/*{
	int ht = HitTest(pt);
	int	unused;
	CRect rect;

	StopTracking(FALSE);

	if ((GetStyle() & SPLS_DYNAMIC_SPLIT) == 0) {
		if( ht == vSplitterBar1 ) {
			if( m_pos == 0 ) {
				GetRowInfo(0, m_pos, unused);
				SetRowInfo(0, 0, 0);
			} else {
				SetRowInfo(0, m_pos, 0);
				m_pos = 0;
			}
			RecalcLayout();
		}
		else if (ht == hSplitterBar1) {
			if( m_pos == 0 ) {
				GetColumnInfo(0, m_pos, unused);
				SetColumnInfo(0, 0, 0);
			} else {
				SetColumnInfo(0, m_pos, 0);
				m_pos = 0;
			}
			RecalcLayout();
		}
	}
	return;     // do nothing if layout is not static
}
*/
