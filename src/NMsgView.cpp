// NMsgView.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "NMsgView.h"
#include "PictureCtrl.h"
#include "CPictureCtrlDemoDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
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
	ON_NOTIFY(NM_CLICK, IDC_ATTACHMENTS, OnActivating)
END_MESSAGE_MAP()

CString GetmboxviewTempPath(void);

void NMsgView::OnActivating(NMHDR* pNMHDR, LRESULT* pResult) 
{
	if (m_bImageViewer)
	{
		CCPictureCtrlDemoDlg dlg;
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
		HINSTANCE result = ShellExecute(NULL, _T("open"), path, NULL, NULL, SW_SHOWNORMAL);
	}

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
	CWnd::OnSize(nType, cx, cy);
	int nOffset = m_bMax?CAPT_MAX_HEIGHT:CAPT_MIN_HEIGHT;
	cx -= BSIZE*2;
	cy -= BSIZE*2;

	int acy = m_bAttach ? m_nAttachSize : 0;
	
	m_browser.MoveWindow(BSIZE, BSIZE+nOffset, cx, cy - acy - nOffset);
	m_attachments.MoveWindow(BSIZE, cy-acy+BSIZE, cx, acy);
	
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

			strCharSetId.Format(_T("%d"), CnfCharsetId);
			StrCharset += " (" + charset + "/" + strCharSetId + "*)";

			charsetId = CnfCharsetId;
		}
		else {
			strCharSetId.Format(_T("%d"), CharsetId);
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
}
