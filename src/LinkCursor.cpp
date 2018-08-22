#include "stdafx.h"
#include "LinkCursor.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif


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
