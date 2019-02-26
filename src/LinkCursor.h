#if !defined(_LINK_CURSOR_)
#define _LINK_CURSOR_

#pragma once
#include "afxwin.h"
class CLinkCursor :
	public CStatic
{

public:
	CLinkCursor();
	~CLinkCursor();
	DECLARE_MESSAGE_MAP()
	afx_msg BOOL OnSetCursor(CWnd* pWnd, UINT nHitTest, UINT message);
};

#endif // _LINK_CURSOR_

