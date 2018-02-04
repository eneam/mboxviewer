#if !defined(AFX_WHEELTREECTRL_H__768E2058_EFAF_482F_9E9A_A997136A4A1C__INCLUDED_)
#define AFX_WHEELTREECTRL_H__768E2058_EFAF_482F_9E9A_A997136A4A1C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// WheelTreeCtrl.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CWheelTreeCtrl window

class CWheelTreeCtrl : public CTreeCtrl
{
// Construction
public:
	CWheelTreeCtrl();

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CWheelTreeCtrl)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CWheelTreeCtrl();

	// Generated message map functions
protected:
	//{{AFX_MSG(CWheelTreeCtrl)
	afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_WHEELTREECTRL_H__768E2058_EFAF_482F_9E9A_A997136A4A1C__INCLUDED_)
