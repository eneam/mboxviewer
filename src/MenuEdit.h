// MenuEdit.h : header file
// Written by PJ Arends
// pja@telus.net
// http://www3.telus.net/pja/
//
// MBox viewer team: adapted for mbox viewer

#if !defined(AFX_MENUEDIT_H__8EA53611_FD2B_11D4_B625_D04FA07D2222__INCLUDED_)
#define AFX_MENUEDIT_H__8EA53611_FD2B_11D4_B625_D04FA07D2222__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif 

class CMenuEdit : public CEdit
{
public:
    CMenuEdit();

	void SetBkColor(COLORREF crColor); // This Function is to set the BackGround Color for the Text and the Edit Box.
	void SetTextColor(COLORREF crColor); // This Function is to set the Color for the Text.

protected:
    virtual BOOL OnCommand(WPARAM wParam, LPARAM lParam);
	afx_msg void OnClose();
    afx_msg void OnContextMenu(CWnd* pWnd, CPoint point);

	CBrush m_brBkgnd; // Holds Brush Color for the Edit Box
	COLORREF m_crBkColor; // Holds the Background Color for the Text
	COLORREF m_crTextColor; // Holds the Color for the Text
	afx_msg HBRUSH CtlColor(CDC* pDC, UINT nCtlColor); // This Function Gets Called Every Time Your Window Gets Redrawn.

    DECLARE_MESSAGE_MAP()
};

#endif
