// MenuEdit.cpp : implementation file
// Written by PJ Arends
// pja@telus.net
// http://www3.telus.net/pja/

#include "stdafx.h"
#include "MenuEdit.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define MES_UNDO        _T("&Undo")
#define MES_CUT         _T("Cu&t")
#define MES_COPY        _T("&Copy")
#define MES_PASTE       _T("&Paste")
#define MES_DELETE      _T("&Delete")
#define MES_SELECTALL   _T("Select &All")
#define ME_SELECTALL    WM_USER + 0x7000

BEGIN_MESSAGE_MAP(CMenuEdit, CEdit)
    ON_WM_CONTEXTMENU()
	ON_WM_CLOSE()
	ON_WM_CTLCOLOR_REFLECT()
END_MESSAGE_MAP()

// Colorref's to use with your Programs

#define RED        RGB(127,  0,  0)
#define GREEN      RGB(  0,127,  0)
#define BLUE       RGB(  0,  0,127)
#define LIGHTRED   RGB(255,  0,  0)
#define LIGHTGREEN RGB(  0,255,  0)
#define LIGHTBLUE  RGB(  0,  0,255)
#define BLACK      RGB(  0,  0,  0)
#define WHITE      RGB(255,255,255)
#define GRAY       RGB(192,192,192)



CMenuEdit::CMenuEdit()
{
	m_crBkColor = ::GetSysColor(CTLCOLOR_STATIC); // Initializing background color to the system face color.
	m_crBkColor = RGB(230, 230, 230);
	m_crTextColor = BLACK; // Initializing text color to black
	m_brBkgnd.CreateSolidBrush(m_crBkColor); // Creating the Brush Color For the Edit Box Background
}

void CMenuEdit::OnContextMenu(CWnd* pWnd, CPoint point)
{
    SetFocus();
    CMenu menu;
    menu.CreatePopupMenu();
    BOOL bReadOnly = GetStyle() & ES_READONLY;

    DWORD flags = CanUndo() && !bReadOnly ? 0 : MF_GRAYED;
    //menu.InsertMenu(0, MF_BYPOSITION | flags, EM_UNDO, MES_UNDO);
    menu.InsertMenu(1, MF_BYPOSITION | MF_SEPARATOR);

    DWORD sel = GetSel();
    flags = LOWORD(sel) == HIWORD(sel) ? MF_GRAYED : 0;
    menu.InsertMenu(2, MF_BYPOSITION | flags, WM_COPY, MES_COPY);

    flags = (flags == MF_GRAYED || bReadOnly) ? MF_GRAYED : 0;
    //menu.InsertMenu(2, MF_BYPOSITION | flags, WM_CUT, MES_CUT);
    //menu.InsertMenu(4, MF_BYPOSITION | flags, WM_CLEAR, MES_DELETE);

    flags = IsClipboardFormatAvailable(CF_TEXT) && !bReadOnly ? 0 : MF_GRAYED;
    //menu.InsertMenu(4, MF_BYPOSITION | flags, WM_PASTE, MES_PASTE);

    menu.InsertMenu(6, MF_BYPOSITION | MF_SEPARATOR);

    int len = GetWindowTextLength();
    flags = (!len || (LOWORD(sel) == 0 && HIWORD(sel) == len)) ? MF_GRAYED : 0;
    menu.InsertMenu(7, MF_BYPOSITION | flags, ME_SELECTALL, MES_SELECTALL);
	menu.InsertMenu(6, MF_BYPOSITION | MF_SEPARATOR);

    menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON, point.x, point.y, this);
}

BOOL CMenuEdit::OnCommand(WPARAM wParam, LPARAM lParam)
{
    switch (LOWORD(wParam))
    {
    case EM_UNDO:
    case WM_CUT:
    case WM_COPY:
    case WM_CLEAR:
    case WM_PASTE:
        return SendMessage(LOWORD(wParam));
    case ME_SELECTALL:
        return SendMessage (EM_SETSEL, 0, -1);
    default:
        return CEdit::OnCommand(wParam, lParam);
    }
}

void CMenuEdit::OnClose()
{
	//ShowWindow(SW_HIDE);

	CWnd::OnClose();
}

void CMenuEdit::SetTextColor(COLORREF crColor)
{
	m_crTextColor = crColor; // Passing the value passed by the dialog to the member varaible for Text Color
	RedrawWindow();
}

void CMenuEdit::SetBkColor(COLORREF crColor)
{
	m_crBkColor = crColor; // Passing the value passed by the dialog to the member varaible for Backgound Color
	m_brBkgnd.DeleteObject(); // Deleting any Previous Brush Colors if any existed.
	m_brBkgnd.CreateSolidBrush(crColor); // Creating the Brush Color For the Edit Box Background
	RedrawWindow();
}

HBRUSH CMenuEdit::CtlColor(CDC* pDC, UINT nCtlColor)
{
	HBRUSH hbr;
	hbr = (HBRUSH)m_brBkgnd; // Passing a Handle to the Brush
	pDC->SetBkColor(m_crBkColor); // Setting the Color of the Text Background to the one passed by the Dialog
	pDC->SetTextColor(m_crTextColor); // Setting the Text Color to the one Passed by the Dialog

	if (nCtlColor)       // To get rid of compiler warning
		nCtlColor += 0;

	return hbr;
}
