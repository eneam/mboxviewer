// NMsgView.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "NMsgView.h"
#include "PictureCtrl.h"
#include "CPictureCtrlDemoDlg.h"
#include "MboxMail.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
//#define _CRTDBG_MAP_ALLOC  
//#include <stdlib.h>  
//#include <crtdbg.h> 
#endif

#define CAPT_MAX_HEIGHT	50
#define CAPT_MIN_HEIGHT 5

#define TEXT_BOLD_LEFT	6 // was 10
#define TEXT_NORM_LEFT	85
#define TEXT_LINE_ONE	4
#define TEXT_LINE_TWO	29

#define BSIZE 5

IMPLEMENT_DYNCREATE(NMsgView, CWnd)

/////////////////////////////////////////////////////////////////////////////
// NMsgView

NMsgView::NMsgView()
{
	m_bAttach = FALSE;
	m_nAttachSize = 50;
	m_bMax = TRUE;

	m_searchID = "mboxview_Search";
	m_matchStyle = "color: white; background-color: blue";

	// Get the log font.
	NONCLIENTMETRICS ncm;
	memset(&ncm, 0, sizeof(NONCLIENTMETRICS));
	ncm.cbSize = sizeof(NONCLIENTMETRICS);

	VERIFY(::SystemParametersInfo(SPI_GETNONCLIENTMETRICS,
		sizeof(NONCLIENTMETRICS), &ncm, 0));
	
	ncm.lfMessageFont.lfWeight = 700;
	m_BoldFont.CreateFontIndirect(&ncm.lfMessageFont);

	ncm.lfMessageFont.lfWeight = 400;
	m_NormFont.CreateFontIndirect(&ncm.lfMessageFont);

	HDC hdc = ::GetWindowDC(NULL);
	//ncm.lfMessageFont.lfHeight = -MulDiv(12, GetDeviceCaps(hdc, LOGPIXELSY), 72);
	ncm.lfMessageFont.lfHeight = -MulDiv(11, GetDeviceCaps(hdc, LOGPIXELSY), 72);
	::ReleaseDC(NULL, hdc);

	ncm.lfMessageFont.lfWeight = 400;
	m_BigFont.CreateFontIndirect(&ncm.lfMessageFont);

	ncm.lfMessageFont.lfWeight = 700;
	m_BigBoldFont.CreateFontIndirect(&ncm.lfMessageFont);

	m_strTitleSubject.LoadString(IDS_TITLE_SUBJECT);
	m_strSubject.LoadString(IDS_DESC_NONE);
	m_strTitleFrom.LoadString(IDS_TITLE_FROM);
	m_strFrom.LoadString(IDS_DESC_NONE);
	m_strTitleDate.LoadString(IDS_TITLE_DATE);
	m_strDate.LoadString(IDS_DESC_NONE);
	m_strTitleTo.LoadString(IDS_TITLE_TO);
	m_strTo.LoadString(IDS_DESC_NONE);
	m_strTitleBody.LoadString(IDS_TITLE_BODY);
	m_subj_charsetId = 0;
	m_from_charsetId = 0;
	m_date_charsetId = 0;
	m_to_charsetId = 0;
	m_body_charsetId = 0;

	m_cnf_subj_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "subjCharsetId");
	m_cnf_from_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "fromCharsetId");
	m_cnf_date_charsetId = 0;
	m_cnf_to_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "toCharsetId");
	m_show_charsets = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "showCharsets");

	DWORD bImageViewer;
	BOOL retval;
	retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("imageViewer"), bImageViewer);
	if (retval == TRUE) {
		m_bImageViewer = bImageViewer;
	}
	else {
		bImageViewer = 1;
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("imageViewer"), bImageViewer);
		m_bImageViewer = bImageViewer;
	}
}

NMsgView::~NMsgView()
{
}


BEGIN_MESSAGE_MAP(NMsgView, CWnd)
	//{{AFX_MSG_MAP(NMsgView)
	ON_WM_CREATE()
	ON_WM_SIZE()
	ON_WM_PAINT()
	ON_WM_ERASEBKGND()
	ON_WM_LBUTTONDBLCLK()
	//}}AFX_MSG_MAP
	//ON_NOTIFY(NM_CLICK, IDC_ATTACHMENTS, OnActivating)  // TODO: Had to disable for now to allow OnDoubleClick to work !!
	ON_NOTIFY(NM_DBLCLK, IDC_ATTACHMENTS, OnDoubleClick)
	ON_NOTIFY(NM_RCLICK, IDC_ATTACHMENTS, OnRClick)  // Right Click Menu
	//ON_WM_GETMINMAXINFO()
	//ON_WM_WINDOWPOSCHANGED()
	//ON_WM_WINDOWPOSCHANGING()
	//ON_WM_SIZING()
END_MESSAGE_MAP()

CString GetmboxviewTempPath(char *name = 0);

void NMsgView::OnActivating(NMHDR* pNMHDR, LRESULT* pResult) 
{
	NMITEMACTIVATE *pnm = (NMITEMACTIVATE *)pNMHDR;
	int nItem = pnm->iItem;
	CString attachmentName = m_attachments.GetItemText(nItem, 0);

	*pResult = 0;
}
/////////////////////////////////////////////////////////////////////////////
// NMsgView message handlers

void NMsgView::PostNcDestroy() 
{
	m_font.DeleteObject();
	m_BoldFont.DeleteObject();
	m_NormFont.DeleteObject();
	m_BigFont.DeleteObject();
	m_BigBoldFont.DeleteObject();
	DestroyWindow();
	delete this;
}

int NMsgView::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CWnd::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	if( !m_browser.Create(NULL, NULL, WS_CHILD|WS_VISIBLE, CRect(), this, IDC_BROWSER) )
		return -1;
	if( !m_attachments.Create(WS_CHILD|WS_VISIBLE|LVS_SINGLESEL|LVS_SMALLICON|LVS_SHOWSELALWAYS, CRect(), this, IDC_ATTACHMENTS) )
		return -1;
	m_attachments.SendMessage((CCM_FIRST + 0x7), 5, 0);
	m_attachments.SetTextColor (RGB(0,0,0));
	if( ! m_font.CreatePointFont (85, _T("Tahoma")) )
		if( ! m_font.CreatePointFont (85, _T("Verdana")) )
			m_font.CreatePointFont (85, _T("Arial"));

    m_attachments.SetFont (&m_font);
	CImageList sysImgList;
	SHFILEINFO shFinfo;
	
	sysImgList.Attach((HIMAGELIST)SHGetFileInfo( _T("C:\\"),
							  0,
							  &shFinfo,
							  sizeof( shFinfo ),
							  SHGFI_SYSICONINDEX |
							  SHGFI_SMALLICON ));
	
	m_attachments.SetImageList( &sysImgList, LVSIL_SMALL );
	sysImgList.Detach();
//	m_attachments.SetExtendedStyle(WS_EX_STATICEDGE);
	
	return 0;
}

void NMsgView::OnSize(UINT nType, int cx, int cy) 
{
	int deb1, deb2, deb3, deb4;
	CWnd::OnSize(nType, cx, cy);

	if (nType == SIZE_RESTORED)
		int deb = 1;

	switch (nType)
	{
	case SIZE_MAXIMIZED:
		// window was maximized
		deb1 = 1;
		break;

	case SIZE_MINIMIZED:
		// window was minimized
		deb2 = 1;
		break;

	case SIZE_RESTORED:
		// misleading - this occurs when restored from minimized/maximized AND
		// for normal size operations when already restored
		deb3 = 1;
		break;

	default:
		// you could also deal with SIZE_MAXHIDE and SIZE_MAXSHOW
		// but rarely need to
		deb4 = 1;
		break;
	}

	int nOffset = m_bMax?CAPT_MAX_HEIGHT:CAPT_MIN_HEIGHT;
	cx -= BSIZE*2;
	cy -= BSIZE*2;

	int acy = m_bAttach ? m_nAttachSize : 0;
	
	m_browser.MoveWindow(BSIZE, BSIZE+nOffset, cx, cy - acy - nOffset);
	m_attachments.MoveWindow(BSIZE, cy-acy+BSIZE, cx, acy);

	// TODO: seem to fix resizing issue; it should not be needed by iy seem to work
	Invalidate();
	UpdateWindow();
}

void NMsgView::UpdateLayout()
{
	CRect r;
	GetClientRect(r);

	int cx = r.Width()-1;
	int cy = r.Height()-1;
	int nOffset = m_bMax?CAPT_MAX_HEIGHT:CAPT_MIN_HEIGHT;
	cx -= BSIZE*2;
	cy -= BSIZE*2;

	int acy = m_bAttach ? m_nAttachSize : 0;
	
	m_browser.MoveWindow(BSIZE, BSIZE+nOffset, cx, cy - acy - nOffset);
	m_attachments.MoveWindow(BSIZE, cy-acy+BSIZE, cx, acy);

	Invalidate();
	UpdateWindow();
}

void FrameGradientFill( CDC *dc, CRect rect, int size, COLORREF crstart, COLORREF crend)
{
	// draw black frame around the message window
	// do top
	int r1=GetRValue(crstart),g1=GetGValue(crstart),b1=GetBValue(crstart);
	int r2=GetRValue(crend),g2=GetGValue(crend),b2=GetBValue(crend);
	int height = size-1;
	int left = rect.left;
	int top = rect.top;
	int width = rect.Width();
	int i;
	for(i=0;i<height;i++)
	{ 
		int r,g,b;
		r = r1 + (i * (r2-r1) / height);
		g = g1 + (i * (g2-g1) / height);
		b = b1 + (i * (b2-b1) / height);
		dc->FillSolidRect(left+i,top+i,width-i,1,RGB(r,g,b));
	}
	// bottom
	r2=GetRValue(crstart),g2=GetGValue(crstart),b2=GetBValue(crstart);
	r1=GetRValue(crend),g1=GetGValue(crend),b1=GetBValue(crend);
	height = size-1;
	left = rect.left;
	top = rect.bottom - height;
	width = rect.Width();

	for(i=0;i<height;i++)
	{ 
		int r,g,b;
		r = r1 + (i * (r2-r1) / height);
		g = g1 + (i * (g2-g1) / height);
		b = b1 + (i * (b2-b1) / height);
		dc->FillSolidRect(left+height-i,top+i,width+height-i,1,RGB(r,g,b));
	}
	// left
	r1=GetRValue(crstart),g1=GetGValue(crstart),b1=GetBValue(crstart);
	r2=GetRValue(crend),g2=GetGValue(crend),b2=GetBValue(crend);
	height = rect.Height();
	left = rect.left;
	top = rect.top;
	width = size-1;
	
	for(i=0;i<width;i++)
	{ 
		int r,g,b;
		r = r1 + (i * (r2-r1) / width);
		g = g1 + (i * (g2-g1) / width);
		b = b1 + (i * (b2-b1) / width);
		dc->FillSolidRect(left+i,top+i,1,height-i-i,RGB(r,g,b));
	}
	// right
	r2=GetRValue(crstart),g2=GetGValue(crstart),b2=GetBValue(crstart);
	r1=GetRValue(crend),g1=GetGValue(crend),b1=GetBValue(crend);
	width = size-1;
	height = rect.Height();
	left = rect.right - width;
	top = rect.top;
	
	for(i=0;i<width;i++)
	{ 
		int r,g,b;
		r = r1 + (i * (r2-r1) / width);
		g = g1 + (i * (g2-g1) / width);
		b = b1 + (i * (b2-b1) / width);
		dc->FillSolidRect(left+i,top+width-i,1,height-(width-i)-(width-i),RGB(r,g,b));
	}

	rect.left += size-1;
	rect.right -= size-1;
	rect.bottom -= size-1;
	rect.top += size-1;
	::FrameRect(dc->m_hDC, rect, (HBRUSH)GetStockObject(BLACK_BRUSH));
}

int NMsgView::PaintHdrField(CPaintDC &dc, CRect	&r, int x_pos, int y_pos, BOOL bigFont, CString &FieldTitle,  CString &FieldText, CString &Charset, UINT CharsetId, UINT CnfCharsetId)
{
	HDC hDC = dc.GetSafeHdc();

	int xpos = x_pos;
	int ypos = y_pos;

	if (bigFont)
		dc.SelectObject(&m_BigBoldFont);
	else
		dc.SelectObject(&m_BoldFont);
	xpos += TEXT_BOLD_LEFT;
	dc.ExtTextOut(xpos, ypos, ETO_CLIPPED, r, FieldTitle, NULL);
	CSize szFieldTitle = dc.GetTextExtent(FieldTitle);

	if (bigFont)
		dc.SelectObject(&m_BigFont);
	else
		dc.SelectObject(&m_NormFont);

	UINT charsetId = CharsetId;
	CString StrCharset;
	CString strCharSetId;
	if (m_show_charsets)
	{
		if ((CharsetId == 0) && (CnfCharsetId != 0)) {
			std::string str;
			BOOL ret = id2charset(CnfCharsetId, str);
			CString charset = str.c_str();

			strCharSetId.Format(_T("%u"), CnfCharsetId);
			StrCharset += " (" + charset + "/" + strCharSetId + "*)";

			charsetId = CnfCharsetId;
		}
		else {
			strCharSetId.Format(_T("%u"), CharsetId);
			StrCharset += " (" + Charset + "/" + strCharSetId + ")";
		}
		xpos += szFieldTitle.cx;
		dc.ExtTextOut(xpos, ypos, ETO_CLIPPED, r, StrCharset, NULL);

		CSize szStrCharset = dc.GetTextExtent(StrCharset);
		xpos += szStrCharset.cx;
	}
	else
		xpos += szFieldTitle.cx;

	xpos += TEXT_BOLD_LEFT;
	BOOL done = FALSE;
	if (charsetId) {
		CStringW strW;
		if (Str2Wide(FieldText, charsetId, strW)) {
			::ExtTextOutW(hDC, xpos, ypos, ETO_CLIPPED, r, (LPCWSTR)strW, strW.GetLength(), NULL);
			done = TRUE;
		}
	}
	if (!done) {
		dc.ExtTextOut(xpos, ypos, ETO_CLIPPED, r, FieldText, NULL);
	}
	CSize szFieldText = dc.GetTextExtent(FieldText);

	xpos += szFieldText.cx;
	return xpos;
}

void NMsgView::OnPaint() 
{
	CPaintDC dc(this); // device context for painting
	static bool ft = true;
		
	GetWindowRect(&m_rcCaption);
	ScreenToClient(&m_rcCaption);
	m_rcCaption.DeflateRect(BSIZE, 0);
	m_rcCaption.top = BSIZE;
	if( ft ) {
//		CClientDC cdc(this);
		ft = false;
		CRect cr;
		GetClientRect(cr);
		dc.FillSolidRect(cr, ::GetSysColor(COLOR_WINDOW));
	}
	m_rcCaption.bottom = m_rcCaption.top + (m_bMax?CAPT_MAX_HEIGHT:CAPT_MIN_HEIGHT);

	dc.FillSolidRect(m_rcCaption, ::GetSysColor(COLOR_WINDOW));
	dc.Draw3dRect(m_rcCaption, ::GetSysColor(COLOR_BTNHIGHLIGHT),
		::GetSysColor(COLOR_BTNSHADOW));

	if (m_bMax)
	{
		CRect	r = m_rcCaption;

		if( r.Width() > 120 ) {
			r.right -= 120;
			CFont *pOldFont = dc.SelectObject(&m_BoldFont);

			dc.SetBkMode(TRANSPARENT);

			int xpos = 0;
			int ypos = r.top + TEXT_LINE_TWO;
			UINT localCP = GetACP();
			std::string str;
			BOOL ret = id2charset(localCP, str);
			m_date_charset = str.c_str();

			xpos = PaintHdrField(dc, r, xpos, ypos, FALSE, m_strTitleDate, m_strDate, m_date_charset, localCP, m_cnf_date_charsetId);
			xpos += TEXT_BOLD_LEFT;
			xpos = PaintHdrField(dc, r, xpos, ypos, FALSE, m_strTitleFrom, m_strFrom, m_from_charset, m_from_charsetId, m_cnf_from_charsetId);
			xpos += 3 * TEXT_BOLD_LEFT;
			xpos = PaintHdrField(dc, r, xpos, ypos, FALSE, m_strTitleTo, m_strTo, m_to_charset, m_to_charsetId, m_cnf_to_charsetId);

			if (m_show_charsets) {
				xpos += 3 * TEXT_BOLD_LEFT;
				xpos = PaintHdrField(dc, r, xpos, ypos, FALSE, m_strTitleBody, m_strBody, m_body_charset, m_body_charsetId, m_body_charsetId);
			}

			ypos = r.top + TEXT_LINE_ONE;
			xpos = 0;
			xpos = PaintHdrField(dc, r, xpos, ypos, TRUE, m_strTitleSubject, m_strSubject, m_subj_charset, m_subj_charsetId, m_cnf_subj_charsetId);

			// Rstore font
			dc.SelectObject( pOldFont );
		}
	}
	CRect r;
	GetClientRect(r);
	FrameGradientFill( &dc, r, BSIZE, RGB(0x8a, 0x92, 0xa6), RGB(0x6a, 0x70, 0x80));
	
	// Do not call CWnd::OnPaint() for painting messages
}

BOOL NMsgView::OnEraseBkgnd(CDC* pDC) 
{
	// TODO: Add your message handler code here and/or call default
	
	return CWnd::OnEraseBkgnd(pDC);
}

void NMsgView::OnLButtonDblClk(UINT nFlags, CPoint point) 
{
#if 1
	ClearSearchResultsInIHTMLDocument(m_searchID);
#else
	CString path = GetmboxviewTempPath();
	HINSTANCE result = ShellExecute(NULL, _T("open"), path, NULL, NULL, SW_SHOWNORMAL);
	/*
	if (m_rcCaption.PtInRect(point))
	{

		m_bMax = !m_bMax;
		UpdateLayout();
	}
	CWnd::OnLButtonDblClk(nFlags, point);
	m_browser.SetFocus();
	*/
#endif
}

void NMsgView::OnDoubleClick(NMHDR* pNMHDR, LRESULT* pResult)
{
	NMITEMACTIVATE *pnm = (NMITEMACTIVATE *)pNMHDR;
	int nItem = pnm->iItem;
	CString attachmentName = m_attachments.GetItemText(nItem, 0);

	bool isSupportedPictureFile = false;
	PTSTR ext = PathFindExtension(attachmentName);
	CString cext = ext;
	// Need to create tavble of extensions so this class and Picture viewr class can share the list
	if (CCPictureCtrlDemoDlg::isSupportedPictureFile(attachmentName))
	{
		isSupportedPictureFile = true;
	}

	if (m_bImageViewer && isSupportedPictureFile)
	{
		CCPictureCtrlDemoDlg dlg(&attachmentName);
		INT_PTR nResponse = dlg.DoModal();
		if (nResponse == IDOK) {
			int deb = 1;
		}
		else if (nResponse == IDCANCEL)
		{
			int deb = 1;
		}
	}
	else
	{
		CString path = GetmboxviewTempPath();
		CString fullName = path + attachmentName;
		// Photos application doesn't show next/prev photo even with path specified
		// Buils in Picture Viewer is set as deafault
		HWND h = GetSafeHwnd();
		HINSTANCE result = ShellExecute(h, _T("open"), attachmentName, NULL, path, SW_SHOWNORMAL);
		CheckShellExecuteResult(result, h);

	}

	*pResult = 0;
}

int FindMenuItem(CMenu* Menu, LPCTSTR MenuName) {
	int count = Menu->GetMenuItemCount();
	for (int i = 0; i < count; i++) {
		CString str;
		if (Menu->GetMenuString(i, str, MF_BYPOSITION) &&
			str.Compare(MenuName) == 0)
			return i;
	}
	return -1;
}

int AttachIcon(CMenu* Menu, LPCTSTR MenuName, int resourceId, CBitmap  &cmap)
{
	MENUITEMINFO minfo;
	memset(&minfo, 0, sizeof(minfo));
	minfo.cbSize = sizeof(minfo);
	minfo.fMask = MIIM_BITMAP;
	int pos = -1;
	if (cmap.LoadBitmap(resourceId)) {
		minfo.hbmpItem = cmap.operator HBITMAP();

		pos = FindMenuItem(Menu, MenuName);
		if (pos >= 0)
			Menu->SetMenuItemBitmaps(pos, MF_BYPOSITION, &cmap, &cmap);
	}
	return pos;
}

void NMsgView::OnRClick(NMHDR* pNMHDR, LRESULT* pResult)
{
	const char *Open = _T("Open");
	const char *Print = _T("Print");
	const char *OpenFileLocation = _T("Open File Location");

	LPNMITEMACTIVATE pnm = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);

	int nItem = pnm->iItem;

	if (nItem < 0)
		return;

	CString attachmentName = m_attachments.GetItemText(nItem, 0);

	TRACE("Selecting %d\n", pnm->iItem);

	// TODO: Add your control notification handler code here
#define MaxShellExecuteErrorCode 32

	CPoint pt;
	::GetCursorPos(&pt);
	CWnd *wnd = WindowFromPoint(pt);

	CMenu menu;
	menu.CreatePopupMenu();
	menu.AppendMenu(MF_SEPARATOR);

	const UINT M_OPEN_Id = 1;
	AppendMenu(&menu, M_OPEN_Id, Open);

	const UINT M_PRINT_Id = 2;
	AppendMenu(&menu, M_PRINT_Id, Print);

	const UINT M_OpenFileLocation_Id = 3;
	AppendMenu(&menu, M_OpenFileLocation_Id, OpenFileLocation);

	CBitmap  printMap;
	int retval = AttachIcon(&menu, Print, IDB_PRINT, printMap);
	CBitmap  foldertMap;
	retval = AttachIcon(&menu, OpenFileLocation, IDB_FOLDER, foldertMap);

	int command = menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, this);

	UINT nFlags = TPM_RETURNCMD;
	CString menuString;
	int chrCnt = menu.GetMenuString(command, menuString, nFlags);

	bool forceOpen = false;
	HINSTANCE result = (HINSTANCE)(MaxShellExecuteErrorCode +1);  // OK
	if (command == M_PRINT_Id)
	{
		CString path = GetmboxviewTempPath();
		CString fullName = path + attachmentName;
		result = ShellExecute(NULL, "print", attachmentName, NULL, path, SW_SHOWNORMAL);
		if ((UINT)result <= MaxShellExecuteErrorCode) {
			CString errorText;
			ShellExecuteError2Text((UINT)result, errorText);
			errorText += ".\nOk to try to open this file ?";
			HWND h = GetSafeHwnd();
			int answer = ::MessageBox(h, errorText, _T("Info"), MB_APPLMODAL | MB_ICONINFORMATION | MB_YESNO);
			if (answer == IDYES)
				forceOpen = true;
		}
		int deb = 1;
	}
	if ((command == M_OPEN_Id) || forceOpen)
	{
		CString path = GetmboxviewTempPath();
		CString fullName = path + attachmentName;

		DWORD binaryType = 0;
		BOOL isExe = GetBinaryTypeA(fullName, &binaryType);

		HWND h = GetSafeHwnd();
		result = ShellExecute(h, "open", attachmentName, NULL, path, SW_SHOWNORMAL);
		CheckShellExecuteResult(result, h);

		int deb = 1;
	}
	else if (command == M_OpenFileLocation_Id)
	{
		CString path = GetmboxviewTempPath();
		CString fullName = path + attachmentName;

		if (BrowseToFile(fullName) == FALSE) {
			HWND h = GetSafeHwnd();
			HINSTANCE result = ShellExecute(h, _T("open"), path, NULL, NULL, SW_SHOWNORMAL);
			CheckShellExecuteResult(result, h);
		}

		int deb = 1;
	}

	m_attachments.Invalidate();
	m_attachments.UpdateWindow();

	*pResult = 0;
}

#pragma component(browser, off, references)  // wliminates too many refrences warning but at what price ?
#include <Mshtml.h>
#pragma component(browser, on, references)
#include <atlbase.h>

#ifdef _DEBUG
#define HTML_ASSERT ASSERT
#else
#define HTML_ASSERT(x)
#endif

char *EatNLine(char* p, char* e) { while (p < e && *p++ != '\n'); return p; }

char *SkipEmptyLine(char* p, char* e)
{
	char *p_save = p;
	while ((p < e) && ((*p == '\r') || (*p == ' ') || (*p == '\t')))  // eat empty lines
		p++;
	if (*p == '\n')
		p++;
	else
		p = p_save;
	return p;
}



void BreakBeforeGoingCleanup()
{
	int deb = 1;
}

void NMsgView::MergeWhiteLines(SimpleString *workbuf, int maxOutLines)
{
	if (maxOutLines == 0) {
		workbuf->SetCount(0);
		return;
	}

	// Delete duplicate empty lines
	char *p = workbuf->Data();
	char *e = p + workbuf->Count();

	char *p_end_data;
	char *p_beg_data;
	char *p_save;
	char *p_data;

	unsigned long  len;
	int dataCount = 0;
	int emptyLineCount = 0;
	int outLineCnt = 0;

	while (p < e)
	{
		emptyLineCount = 0;
		p_save = 0;
		while ((p < e) && (p != p_save))
		{
			p_save = p;
			p = SkipEmptyLine(p, e);
			if (p != p_save)
				emptyLineCount++;

		}
		p_beg_data = p;

		while ((p < e) && !((*p == '\r') || (*p == '\n')))
		{
			p = EatNLine(p, e);
			outLineCnt++;
			if ((maxOutLines > 0) && (outLineCnt >= maxOutLines))
			{
				e = p;
				break;
			}
		}

		p_end_data = p;

		if (emptyLineCount > 0)
		{
			p_data = workbuf->Data() + dataCount;
			if ((p_beg_data - p_data) > 1) {
				memcpy(p_data, "\r\n", 2);
				dataCount += 2;
			}
			else
			{
				memcpy(p_data, "\n", 1);
				dataCount += 1;
			}
			outLineCnt++;
		}

		len = p_end_data - p_beg_data;
		//g_tu.hexdump("Data:\n", p_beg_data, len);

		if (len > 0) {
			p_data = workbuf->Data() + dataCount;
			memcpy(p_data, p_beg_data, len);
			dataCount += len;
		}
		if ((maxOutLines > 0) && (outLineCnt >= maxOutLines))
			break;
	}

	workbuf->SetCount(dataCount);

#if 0
	// Verify that outLineCnt is valid
	{
		if (maxOutLines <= 0)
			return;


		char *p = workbuf->Data();
		e = p + workbuf->Count();
		int lineCnt = 0;
		while (p < e)
		{
			p = EatNLine(p, e);
			lineCnt++;
		}
		if (lineCnt != outLineCnt)
			int deb = 1;
	}
#endif
}

void NMsgView::PrintHTMLDocumentToPrinter(SimpleString *inbuf, SimpleString *workbuf, UINT inCodePage)
{
	//CComBSTR cmdID(_T("PRINT"));
	//VARIANT_BOOL vBool;

	IHTMLDocument2 *lpHtmlDocument = 0;
	HRESULT hr;
	VARIANT val;
	VariantInit(&val);
	VARIANT valOut;
	VariantInit(&valOut);

	BOOL retVal = CreateHTMLDocument(&lpHtmlDocument, inbuf, workbuf, inCodePage);
	if ((retVal == FALSE) || (lpHtmlDocument == 0)) {
		return;
	}

#if 0
	CHtmlEditCtrl PrintCtrl;

	if (!PrintCtrl.Create(NULL, WS_CHILD, CRect(0, 0, 0, 0), this, 1))
	{
		ASSERT(FALSE);
		return; // Error!
	}

	//CPrintDialog dlgl(FALSE);  INT_PTR userResult = dlgl.DoModal();
#endif

	//lpHtmlDocument->execCommand(cmdID, VARIANT_TRUE, val, &vBool);

	IOleCommandTarget  *lpOleCommandTarget = 0;
	hr = lpHtmlDocument->QueryInterface(IID_IOleCommandTarget, (VOID**)&lpOleCommandTarget);
	if (FAILED(hr) || !lpOleCommandTarget)
	{
		BreakBeforeGoingCleanup();
		goto cleanup;
	}

	hr = lpOleCommandTarget->Exec(NULL, OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER, &val, &valOut);
	if (FAILED(hr))
	{
		BreakBeforeGoingCleanup();
		goto cleanup;
	}

cleanup:

	hr = VariantClear(&val);
	hr = VariantClear(&valOut);

	if (lpOleCommandTarget)
		lpOleCommandTarget->Release();

	if (lpHtmlDocument)
		lpHtmlDocument->Release();
}

void NMsgView::GetTextFromIHTMLDocument(SimpleString *inbuf, SimpleString *workbuf, UINT inCodePage, UINT outCodePage)
{
	IHTMLDocument2 *lpHtmlDocument = 0;
	LPDISPATCH lpDispatch = 0;
	HRESULT hr;

	if (lpHtmlDocument == 0) 
	{
		BOOL retVal = CreateHTMLDocument(&lpHtmlDocument, inbuf, workbuf, inCodePage);
		if ((retVal == FALSE) || (lpHtmlDocument == 0)){
			return;
		}
	}

	IHTMLElement *lpBodyElm = 0;
	hr = lpHtmlDocument->get_body(&lpBodyElm);
	//HTML_ASSERT(SUCCEEDED(hr) && lpBodyElm);
	if (!lpBodyElm) {
		lpHtmlDocument->Release();
		return;
	}

	// Need to remove STYLE tag otherwise it will be part of text from lpBodyElm->get_innerText(&bstrTxt);
	// TODO: IHTMLDocument2:get_text() is not the greatest, adding or missing text, and slow.
	// Don't think any work is done on MFC and c++ IHTMLDocument.
	// Implemented incomplete PrintIHTMLDocument based on IHTMLDocument framework but is incomplete. It still be slow.
	//
	RemoveStyleTagFromIHTMLDocument(lpBodyElm);

	CComBSTR bstrTxt;
	hr = lpBodyElm->get_innerText(&bstrTxt);
	HTML_ASSERT(SUCCEEDED(hr));
	if (FAILED(hr)) {
		lpBodyElm->Release();
		return;
	}

	int wlen = bstrTxt.Length();
	wchar_t *wstr = bstrTxt;

	//SimpleString *workbuf = MboxMail::m_workbuf;
	workbuf->ClearAndResize(10000);

	BOOL ret = WStr2CodePage(wstr, wlen, outCodePage, workbuf);

	MergeWhiteLines(workbuf, -1);

	if (lpBodyElm)
		lpBodyElm->Release();

	if (lpHtmlDocument)
		lpHtmlDocument->Release();
}

BOOL NMsgView::CreateHTMLDocument(IHTMLDocument2 **lpDocument, SimpleString *inbuf, SimpleString *workbuf, UINT inCodePage)
{
	HRESULT hr;
	BOOL ret = TRUE;
	*lpDocument = 0;

	IHTMLDocument2 *lpDoc = 0;
	SAFEARRAY *psaStrings = 0;

	//CComBSTR bstr(inbuf->Data());
	//int bstrLen = bstr.Length();
	//int deb1 = 1;
	//BSTR  bstrTest = SysAllocString(L"TEST");  // Test heap allocations
	//SysFreeString(bstrTest);
	//int deb = 1;

#if 0
	USES_CONVERSION;
	BSTR  bstr = SysAllocString(A2W(inbuf->Data())); // efficient but relies on stack and not heap; but what is Data() is > stack ?
	int bstrLen = SysStringByteLen(bstr);
	int wlen = wcslen(bstr);
#else
	ret = CodePage2WStr(inbuf, inCodePage, workbuf);
	OLECHAR *oledata = (OLECHAR*)workbuf->Data();
	BSTR  bstr = SysAllocString(oledata); 
	int bstrLen = SysStringByteLen(bstr);
	int wlen = wcslen(bstr);
#endif

	hr = CoCreateInstance(CLSID_HTMLDocument, NULL, CLSCTX_INPROC_SERVER,
		IID_IHTMLDocument2, (void**)&lpDoc);
	HTML_ASSERT(SUCCEEDED(hr) && lpDoc);
	if (FAILED(hr) || !lpDoc) {
		BreakBeforeGoingCleanup();
		ret = FALSE;
		goto cleanup;
	}

	// Creates a new one-dimensional array
	psaStrings = SafeArrayCreateVector(VT_VARIANT, 0, 1);
	HTML_ASSERT(psaStrings);
	if (psaStrings == 0) {
		BreakBeforeGoingCleanup();
		ret = FALSE;
		goto cleanup;
	}
	VARIANT *param = 0;
	hr = SafeArrayAccessData(psaStrings, (LPVOID*)&param);
	HTML_ASSERT(SUCCEEDED(hr) && param);
	if (FAILED(hr) || (param == 0)) {
		BreakBeforeGoingCleanup();
		ret = FALSE;
		goto cleanup;
	}
	param->vt = VT_BSTR;
	param->bstrVal = bstr;

	// put_designMode(L"on");  should disable all activity such as running scripts
	// May need better solution
	hr = lpDoc->put_designMode(L"on");  
	hr = lpDoc->writeln(psaStrings);

	HTML_ASSERT(SUCCEEDED(hr));
	if (FAILED(hr)) {
		BreakBeforeGoingCleanup();
		ret = FALSE;
		goto cleanup;
	}

cleanup:
	// Assume SafeArrayDestroy calls SysFreeString for each BSTR :)
	if (psaStrings != 0) {
		hr = SafeArrayUnaccessData(psaStrings);
		hr = SafeArrayDestroy(psaStrings);
	}
	if (ret == FALSE)
	{
		if (lpDoc) {
			hr = lpDoc->close();
			if (FAILED(hr)) {
				int deb = 1;
			}
			lpDoc->Release();
			lpDoc = 0;
		}
	}
	else
	{
		hr = lpDoc->close();
		if (FAILED(hr)) 
		{
			int deb = 1;
		}
		*lpDocument = lpDoc;
	}
	return ret;
}

void NMsgView::RemoveStyleTagFromIHTMLDocument(IHTMLElement *lpElm)
{
	CComBSTR emptyText("");
	CComBSTR bstrTag;
	//CComBSTR bstrHTML;
	//CComBSTR bstrTEXT;
	VARIANT index;
	HRESULT hr;

	IDispatch *lpCh = 0;
	hr = lpElm->get_children(&lpCh);
	if (FAILED(hr) || !lpCh)
	{
		return;
	}

	IHTMLElementCollection *lpChColl = 0;
	hr = lpCh->QueryInterface(IID_IHTMLElementCollection, (VOID**)&lpChColl);
	if (FAILED(hr) || !lpChColl)
	{
		BreakBeforeGoingCleanup();
		goto cleanup;
	}

	long items = 0;
	IDispatch *ppvDisp = 0;
	IHTMLElement *ppvElement = 0;

	lpChColl->get_length(&items);

	for (long j = 0; j < items; j++)
	{
		index.vt = VT_I4;
		index.lVal = j;
		ppvDisp = 0;
		hr = lpChColl->item(index, index, &ppvDisp);
		if (FAILED(hr) || !ppvDisp) {
			BreakBeforeGoingCleanup();
			goto cleanup;
		}
		if (ppvDisp)
		{
			ppvElement = 0;
			hr = ppvDisp->QueryInterface(IID_IHTMLElement, (void **)&ppvElement);
			if (FAILED(hr) || !ppvElement) {
				BreakBeforeGoingCleanup();
				goto cleanup;
			}
			if (ppvElement)
			{
				bstrTag.Empty();
				//bstrHTML.Empty();
				//bstrTEXT.Empty();

				ppvElement->get_tagName(&bstrTag);
				//ppvElement->get_outerHTML(&bstrHTML);
				//ppvElement->get_innerText(&bstrTEXT);

				CString tag(bstrTag);
				if (tag.CollateNoCase("style") == 0) {
					ppvElement->put_innerText(emptyText);
					int deb = 1;
				}

				RemoveStyleTagFromIHTMLDocument(ppvElement);

				ppvElement->Release();
				ppvElement = 0;

				int deb = 1;
			}
			ppvDisp->Release();
			ppvDisp = 0;
		}
		int deb = 1;
	}
cleanup:

	if (ppvElement)
		ppvElement->Release();

	if (ppvDisp)
		ppvDisp->Release();

	if (lpChColl)
		lpChColl->Release();

	if (lpCh)
		lpCh->Release();
}


BOOL NMsgView::FindElementByTagInIHTMLDocument(IHTMLDocument2 *lpDocument, IHTMLElement **ppvEl, CString &tag)
{
	// browse all elements and return first element found
	IHTMLElementCollection *pAll = 0;
	CComBSTR bstrTag;
	CComBSTR bstrHTML;
	CComBSTR bstrTEXT;
	VARIANT index;

	*ppvEl = 0;
	HRESULT hr = lpDocument->get_all(&pAll);
	if (FAILED(hr) || !pAll) {
		BreakBeforeGoingCleanup();
		goto cleanup;
	}

	long items = 0;
	IDispatch *ppvDisp = 0;
	IHTMLElement *ppvElement = 0;

	pAll->get_length(&items);

	for (long j = 0; j < items; j++)
	{
		index.vt = VT_I4;
		index.lVal = j;
		ppvDisp = 0;
		hr = pAll->item(index, index, &ppvDisp);
		if (FAILED(hr) || !ppvDisp) {
			BreakBeforeGoingCleanup();
			goto cleanup;
		}
		if (ppvDisp) 
		{
			ppvElement = 0;
			hr = ppvDisp->QueryInterface(IID_IHTMLElement, (void **)&ppvElement);
			if (FAILED(hr) || !ppvElement) {
				BreakBeforeGoingCleanup();
				goto cleanup;
			}
			if (ppvElement)
			{
				bstrTag.Empty();
				//bstrHTML.Empty();
				//bstrTEXT.Empty();

				ppvElement->get_tagName(&bstrTag);
				//ppvElement->get_innerHTML(&bstrHTML);
				//ppvElement->get_innerText(&bstrTEXT);

				CString eltag(bstrTag);
				TRACE("TAG=%s\n", (LPCSTR)eltag);

				if (eltag.CompareNoCase(tag))
				{
					ppvElement->Release();
					ppvElement = 0;
				}
				else
					break;

				int deb = 1;
			}
			ppvDisp->Release();
			ppvDisp = 0;
		}
		int deb = 1;
	}

cleanup:
	if (ppvElement)
		*ppvEl = ppvElement;

	if (ppvDisp)
		ppvDisp->Release();

	if (pAll)
		pAll->Release();

	return FALSE;
}

void NMsgView::PrintIHTMLDocument(IHTMLDocument2 *lpDocument)
{
	// Actually this will print BODY
	CString tag("body");
	IHTMLElement *lpElm = 0;
	FindElementByTagInIHTMLDocument(lpDocument, &lpElm, tag);
	if (!lpElm)
	{
		return;
	}

	CComBSTR bstrTEXT;
	lpElm->get_innerText(&bstrTEXT);

	CStringW text;
	PrintIHTMLElement(lpElm, text);

	SimpleString *workbuf = MboxMail::m_workbuf;
	workbuf->ClearAndResize(10000);

	const wchar_t *wstr = text.operator LPCWSTR();
	int wlen = text.GetLength();
	UINT outCodePage = CP_UTF8;
	BOOL ret = WStr2CodePage((wchar_t *)wstr, wlen, outCodePage, workbuf);

	TRACE("TEXT=%s\n", workbuf->Data());
}

void NMsgView::PrintIHTMLElement(IHTMLElement *lpElm, CStringW &text)
{
	CComBSTR bstrTag;
	CComBSTR bstrHTML;
	CComBSTR bstrTEXT;
	VARIANT index;
	HRESULT hr;

	IDispatch *lpCh = 0;
	hr = lpElm->get_children(&lpCh);
	if (FAILED(hr) || !lpCh)
	{
		return;
	}

	IHTMLElementCollection *lpChColl = 0;
	hr = lpCh->QueryInterface(IID_IHTMLElementCollection, (VOID**)&lpChColl);
	if (FAILED(hr) || !lpChColl)
	{
		BreakBeforeGoingCleanup();
		goto cleanup;;
	}

	long items;
	IDispatch *ppvDisp = 0;
	IHTMLElement *ppvElement = 0;

	lpChColl->get_length(&items);

	// TODO: This doesn't quite work, need to find solution.
	if (items <= 0)
	{
		lpElm->get_tagName(&bstrTag);
		lpElm->get_outerHTML(&bstrHTML);
		lpElm->get_innerText(&bstrTEXT);

		text.Append(bstrTEXT);

		CString tag(bstrTag);
		//text.Append("\r\n\r\n");
		TRACE("TAG=%s\n", (LPCSTR)tag);
		//TRACE("TEXT=%s\n", (LPCSTR)innerText);
	}
	else
		text.Append(L"\r\n\r\n");

	for (long j = 0; j < items; j++)
	{
		index.vt = VT_I4;
		index.lVal = j;
		ppvDisp = 0;
		hr = lpChColl->item(index, index, &ppvDisp);
		if (FAILED(hr) || !ppvDisp) {
			BreakBeforeGoingCleanup();
			goto cleanup;
		}
		if (ppvDisp)
		{
			hr = ppvDisp->QueryInterface(IID_IHTMLElement, (void **)&ppvElement);
			if (FAILED(hr) || !ppvElement) {
				BreakBeforeGoingCleanup();
				goto cleanup;
			}
			if (ppvElement)
			{
				bstrTag.Empty();
				bstrHTML.Empty();
				bstrTEXT.Empty();

				ppvElement->get_tagName(&bstrTag);
				ppvElement->get_outerHTML(&bstrHTML);
				ppvElement->get_innerText(&bstrTEXT);


				CString tag(bstrTag);
				TRACE("TAG=%s\n", (LPCSTR)tag);

				if (tag.CollateNoCase("style") == 0)
					int deb = 1;

				if (tag.CollateNoCase("tbody") == 0)
					int deb = 1;

				if (tag.CollateNoCase("td") == 0)
					int deb = 1;

				if (tag.CollateNoCase("body") == 0)
					int deb = 1;

				if (tag.CollateNoCase("style") == 0) {
					CComBSTR emptyText("");
					ppvElement->put_innerText(emptyText);
					int deb = 1;
				}


				PrintIHTMLElement(ppvElement, text);


				ppvElement->Release();
				ppvElement = 0;

				int deb = 1;
			}
			ppvDisp->Release();
			ppvDisp = 0;
		}
		int deb = 1;
	}
cleanup:

	if (ppvElement)
		ppvElement->Release();

	if (ppvDisp)
		ppvDisp->Release();

	if (lpChColl)
		lpChColl->Release();

	if (lpCh)
		lpCh->Release();
}


void NMsgView::FindStringInIHTMLDocument(CString &searchText, BOOL matchWord, BOOL matchCase)
{
	// Based on "Adding a custom search feature to CHtmlViews" on the codeproject by  Marc Richarme, 22 Nov 2000
	// Did resolve some code issues, enhanced and optimized

	// <span id='mboxview_Search' style='color: white; background-color: darkblue'>pa</span>

	HRESULT hr;
	CComBSTR bstrTag;
	CComBSTR bstrHTML;
	CComBSTR bstrTEXT;
	CString htmlPrfix;

	unsigned long matchWordFlag = 2;
	unsigned long matchCaseFlag = 4;
	long  lFlags = 0; // lFlags : 0: match in reverse: 1 , Match partial Words 2=wholeword, 4 = matchcase
	if (matchWord)
		lFlags |= matchWordFlag;
	if (matchCase)
		lFlags |= matchCaseFlag;

	ClearSearchResultsInIHTMLDocument(m_searchID);

	IHTMLDocument2 *lpHtmlDocument = NULL;
	LPDISPATCH lpDispatch = NULL;
	lpDispatch = m_browser.m_ie.GetDocument();
	HTML_ASSERT(lpDispatch);
	if (!lpDispatch)
		return;

	hr = lpDispatch->QueryInterface(IID_IHTMLDocument2, (void**)&lpHtmlDocument);
	HTML_ASSERT(SUCCEEDED(hr) && lpHtmlDocument);
	if (!lpHtmlDocument) {
		lpDispatch->Release();
		return;
	}
	lpDispatch->Release();

	IHTMLElement *lpBodyElm = 0;
	IHTMLBodyElement *lpBody = 0;
	IHTMLTxtRange *lpTxtRange = 0;

	hr = lpHtmlDocument->get_body(&lpBodyElm);
	HTML_ASSERT(SUCCEEDED(hr) && lpBodyElm);
	if (!lpBodyElm) {
		lpHtmlDocument->Release();
		return;
	}
	lpHtmlDocument->Release();

	hr = lpBodyElm->QueryInterface(IID_IHTMLBodyElement, (void**)&lpBody);
	HTML_ASSERT(SUCCEEDED(hr) && lpBody);
	if (!lpBody) {
		lpBodyElm->Release();
		return;
	}
	lpBodyElm->Release();

	hr = lpBody->createTextRange(&lpTxtRange);
	HTML_ASSERT(SUCCEEDED(hr) && lpTxtRange);
	if (!lpTxtRange) {
		lpBody->Release();
		return;
	}
	lpBody->Release();

	CComBSTR html;
	CComBSTR newhtml;
	CComBSTR search(searchText.GetLength() + 1, (LPCTSTR)searchText);

	long t;
	VARIANT_BOOL bFound;
	long count = 0;

	CComBSTR Character(L"Character");
	CComBSTR Textedit(L"Textedit");

	htmlPrfix.Append("<span id='");
	htmlPrfix.Append((LPCTSTR)m_searchID);
	htmlPrfix.Append("' style='");
	htmlPrfix.Append((LPCTSTR)m_matchStyle);
	htmlPrfix.Append("'>");

	bool firstRange = false;
	while ((lpTxtRange->findText(search, count, lFlags, (VARIANT_BOOL*)&bFound) == S_OK) && (VARIANT_TRUE == bFound))
	{
		//IHTMLTxtRange *duplicateRange;
		//lpTxtRange->duplicate(&duplicateRange);

		IHTMLElement *parentText = 0;
		hr = lpTxtRange->parentElement(&parentText);
		if (SUCCEEDED(hr) && parentText)
		{
			bstrTag.Empty();
			bstrHTML.Empty();
			bstrTEXT.Empty();

			parentText->get_tagName(&bstrTag);
			parentText->get_innerHTML(&bstrHTML);
			parentText->get_innerText(&bstrTEXT);

			int deb = 1;
			// Ignore Tags: TITLE, what else ?
			bstrTag.ToLower();
			if (bstrTag == L"title") {
				lpTxtRange->moveStart(Character, 1, &t);
				lpTxtRange->moveEnd(Textedit, 1, &t);

				parentText->Release();
				continue;
			}
			parentText->Release();
		}
		if (parentText)
			parentText->Release();

		if (firstRange == false) {
			firstRange = true;
			if (lpTxtRange->select() == S_OK)
				lpTxtRange->scrollIntoView(VARIANT_TRUE);
		}

		newhtml.Empty();
		lpTxtRange->get_htmlText(&html);
		newhtml.Append((LPCTSTR)htmlPrfix);
		if (searchText == " ")
			newhtml.Append("&nbsp;"); // doesn't work very well, but prevents (some) strange presentation
		else
			newhtml.AppendBSTR(html);
		newhtml.Append("</span>");

		lpTxtRange->pasteHTML(newhtml);

		lpTxtRange->moveStart(Character, searchText.GetLength(), &t);
		lpTxtRange->moveEnd(Textedit, 1, &t);
	}
	if (lpTxtRange)
		lpTxtRange->Release();
}

void NMsgView::ClearSearchResultsInIHTMLDocument(CString searchID)
{
	// Based on "Adding a custom search feature to CHtmlViews" on the codeproject by  Marc Richarme, 22 Nov 2000
	// Did resolve some code issues, enhanced and optimized

	CComBSTR testid(searchID.GetLength() + 1, searchID);
	CComBSTR testtag(5, "SPAN");

	HRESULT hr;
	IHTMLDocument2 *lpHtmlDocument = NULL;
	LPDISPATCH lpDispatch = NULL;
	lpDispatch = m_browser.m_ie.GetDocument();
	HTML_ASSERT(lpDispatch);
	if (!lpDispatch)
		return;

	hr = lpDispatch->QueryInterface(IID_IHTMLDocument2, (void**)&lpHtmlDocument);
	HTML_ASSERT(SUCCEEDED(hr) && lpHtmlDocument);
	if (!lpHtmlDocument) {
		lpDispatch->Release();
		return;
	}
	lpDispatch->Release();

	IHTMLElementCollection *lpAllElements = 0;
	hr = lpHtmlDocument->get_all(&lpAllElements);
	HTML_ASSERT(SUCCEEDED(hr) && lpAllElements);
	if (!lpAllElements) {
		lpHtmlDocument->Release();
		return;
	}
	lpHtmlDocument->Release();

	CComBSTR id;
	CComBSTR tag;
	CComBSTR innerText;

	IUnknown *lpUnk;
	IEnumVARIANT *lpNewEnum;
	if (SUCCEEDED(lpAllElements->get__newEnum(&lpUnk)) && (lpUnk != NULL))
	{
		hr = lpUnk->QueryInterface(IID_IEnumVARIANT, (void**)&lpNewEnum);
		HTML_ASSERT(SUCCEEDED(hr) && lpNewEnum);
		if (!lpNewEnum) {
			lpAllElements->Release();
			return;
		}
		VARIANT varElement;
		IHTMLElement *lpElement;

		VariantInit(&varElement);
		while (lpNewEnum->Next(1, &varElement, NULL) == S_OK)
		{
			HTML_ASSERT(varElement.vt == VT_DISPATCH);
			hr = varElement.pdispVal->QueryInterface(IID_IHTMLElement, (void**)&lpElement);
			HTML_ASSERT(SUCCEEDED(hr) && lpElement);

			if (lpElement)
			{
				id.Empty();
				tag.Empty();
				lpElement->get_id(&id);
				lpElement->get_tagName(&tag);
				if ((id == testid) && (tag == testtag))
				{
					innerText.Empty();
					lpElement->get_innerHTML(&innerText);
					lpElement->put_outerHTML(innerText);
				}
			}
			hr = VariantClear(&varElement);
		}
	}
	lpAllElements->Release();
}


void NMsgView::OnGetMinMaxInfo(MINMAXINFO* lpMMI)
{
	// TODO: Add your message handler code here and/or call default
	int deb = 1;

	CWnd::OnGetMinMaxInfo(lpMMI);
}


void NMsgView::OnWindowPosChanged(WINDOWPOS* lpwndpos)
{
	CWnd::OnWindowPosChanged(lpwndpos);

	// TODO: Add your message handler code here
}


void NMsgView::OnWindowPosChanging(WINDOWPOS* lpwndpos)
{
	CWnd::OnWindowPosChanging(lpwndpos);

	// TODO: Add your message handler code here
}


void NMsgView::OnSizing(UINT fwSide, LPRECT pRect)
{
	CWnd::OnSizing(fwSide, pRect);

	// TODO: Add your message handler code here
}
