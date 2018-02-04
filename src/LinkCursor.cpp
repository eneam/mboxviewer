#include "stdafx.h"
#include "LinkCursor.h"


CLinkCursor::CLinkCursor()
{
}


CLinkCursor::~CLinkCursor()
{
}
BEGIN_MESSAGE_MAP(CLinkCursor, CStatic)
	ON_WM_SETCURSOR()
END_MESSAGE_MAP()


BOOL CLinkCursor::OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message)
{
	::SetCursor(AfxGetApp()->LoadStandardCursor(IDC_HAND));
	return TRUE;

	//return CStatic::OnSetCursor(pWnd, nHitTest, message);
}
