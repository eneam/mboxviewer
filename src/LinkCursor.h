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

