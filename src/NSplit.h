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
