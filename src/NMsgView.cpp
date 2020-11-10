//
//////////////////////////////////////////////////////////////////
//
//  Windows Mbox Viewer is a free tool to view, search and print mbox mail archives.
//
// Source code and executable can be downloaded from
//  https://sourceforge.net/projects/mbox-viewer/  and
//  https://github.com/eneam/mboxviewer
//
//  Copyright(C) 2019  Enea Mansutti, Zbigniew Minciel
//
//  This program is free software; you can redistribute it and/or modify
//  it under the terms of the version 3 of GNU Affero General Public License
//  as published by the Free Software Foundation; 
//
//  This program is distributed in the hope that it will be useful,
//  but WITHOUT ANY WARRANTY; without even the implied warranty of
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU
//  Library General Public License for more details.
//
//  You should have received a copy of the GNU Library General Public
//  License along with this program; if not, write to the
//  Free Software Foundation, Inc., 51 Franklin St, Fifth Floor,
//  Boston, MA  02110 - 1301, USA.
//
//////////////////////////////////////////////////////////////////
//

// NMsgView.cpp : implementation file
//

#include "stdafx.h"
#include "FileUtils.h"
#include "TextUtilsEx.h"
#include "HtmlUtils.h"
#include "mboxview.h"
#include "NMsgView.h"
#include "PictureCtrl.h"
#include "CPictureCtrlDemoDlg.h"
#include "MboxMail.h"
#include "MimeParser.h"
#include "MenuEdit.h"

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
	:m_attachments(this), m_hdr()
{
	m_hdrPaneLayout = 0;
	DWORD hdrPaneLayout = 0;

	m_hdrWindowLen = 0;
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
	m_strTitleCC.LoadString(IDS_TITLE_CC);
	m_strTitleBCC.LoadString(IDS_TITLE_BCC);
	m_strTo.LoadString(IDS_DESC_NONE);
	m_strTitleBody.LoadString(IDS_TITLE_BODY);

	// set by SelectItem in NListView
	m_subj_charsetId = 0;
	m_from_charsetId = 0;
	m_date_charsetId = 0;
	m_to_charsetId = 0;
	m_body_charsetId = 0;

	m_cnf_subj_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "subjCharsetId");
	m_cnf_from_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "fromCharsetId");
	m_cnf_date_charsetId = 0;
	m_cnf_to_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "toCharsetId");
	m_cnf_cc_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "ccCharsetId");
	m_cnf_bcc_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "bccCharsetId");
	m_show_charsets = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "showCharsets");

#if 0
	if (m_cnf_from_charsetId == 0)
		m_cnf_from_charsetId = GetACP();

	if (m_cnf_to_charsetId == 0)
		m_cnf_to_charsetId = GetACP();

	if (m_cnf_subj_charsetId == 0)
		m_cnf_subj_charsetId = GetACP();
#endif

	DWORD bImageViewer;
	BOOL retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("imageViewer"), bImageViewer);
	if (retval == TRUE) {
		m_bImageViewer = bImageViewer;
	}
	else {
		bImageViewer = 1;
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("imageViewer"), bImageViewer);
		m_bImageViewer = bImageViewer;
	}

	m_frameCx_TreeNotInHide = 700;
	m_frameCy_TreeNotInHide = 200;
	m_frameCx_TreeInHide = 700;
	m_frameCy_TreeInHide = 200;


	CString m_section = CString(sz_Software_mboxview) + "\\" + "Window Placement";

	BOOL ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "MsgFrameTreeNotHiddenWidth", m_frameCx_TreeNotInHide);
	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "MsgFrameTreeNotHiddenHeight", m_frameCy_TreeNotInHide);
	//
	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "MsgFrameTreeHiddenWidth", m_frameCx_TreeInHide);
	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, m_section, "MsgFrameTreeHiddenHeight", m_frameCy_TreeInHide);
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
	//ON_WM_GETMINMAXINFO()
	//ON_WM_WINDOWPOSCHANGED()
	//ON_WM_WINDOWPOSCHANGING()
	//ON_WM_SIZING()
	ON_WM_CLOSE()
	ON_WM_RBUTTONDOWN()
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// NMsgView message handlers

void NMsgView::PostNcDestroy() 
{
	m_font.DeleteObject();
	m_BoldFont.DeleteObject();
	m_NormFont.DeleteObject();
	m_BigFont.DeleteObject();
	m_BigBoldFont.DeleteObject();
	m_attachments.ReleaseResources();
	DestroyWindow();
	delete this;
}

int NMsgView::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CWnd::OnCreate(lpCreateStruct) == -1)
		return -1;

	if (!m_browser.Create(NULL, NULL, WS_CHILD | WS_VISIBLE, CRect(), this, IDC_BROWSER))
		return -1;

	if (!m_attachments.Create(WS_CHILD | WS_VISIBLE | LVS_SINGLESEL | LVS_SMALLICON | LVS_SHOWSELALWAYS | LVS_AUTOARRANGE, CRect(), this, IDC_ATTACHMENTS))
		return -1;
	m_attachments.SendMessage((CCM_FIRST + 0x7), 5, 0);  // #define CCM_SETVERSION          (CCM_FIRST + 0x7)
	m_attachments.SetTextColor(RGB(0, 0, 0));
	if (!m_font.CreatePointFont(85, _T("Tahoma")))
		if (!m_font.CreatePointFont(85, _T("Verdana")))
			m_font.CreatePointFont(85, _T("Arial"));

	m_attachments.SetFont(&m_font);
	CImageList sysImgList;
	SHFILEINFO shFinfo;

	sysImgList.Attach((HIMAGELIST)SHGetFileInfo(_T("C:\\"),
		0,
		&shFinfo,
		sizeof(shFinfo),
		SHGFI_SYSICONINDEX |
		SHGFI_SMALLICON));

	m_attachments.SetImageList(&sysImgList, LVSIL_SMALL);
	sysImgList.Detach();
	//	m_attachments.SetExtendedStyle(WS_EX_STATICEDGE);

	if (!m_hdr.Create(WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_READONLY | WS_VSCROLL | WS_HSCROLL | WS_SYSMENU , CRect(), this, IDC_EDIT_MAIL_HDR))
	{
		int deb = 1;
	}

	// TODO: Move to CMainFrame

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		if (pFrame->GetMessageWindowPosition() > 1)
			m_hdrPaneLayout = 1;

		if (m_hdrPaneLayout == 0)
		{
			CMenu *menu = pFrame->GetMenu();
			menu->CheckMenuItem(ID_MESSAGEHEADERPANELAYOUT_DEFAULT, MF_CHECKED);
			menu->CheckMenuItem(ID_MESSAGEHEADERPANELAYOUT_EXPANDED, MF_UNCHECKED);
		}
		else
		{
			CMenu *menu = pFrame->GetMenu();
			menu->CheckMenuItem(ID_MESSAGEHEADERPANELAYOUT_DEFAULT, MF_UNCHECKED);
			menu->CheckMenuItem(ID_MESSAGEHEADERPANELAYOUT_EXPANDED, MF_CHECKED);
		}

		pFrame->CheckMessagewindowPositionMenuOption(pFrame->GetMessageWindowPosition());
	}

	return 0;
}

void NMsgView::OnSize(UINT nType, int cx, int cy) 
{
	int deb1, deb2, deb3, deb4;
	CWnd::OnSize(nType, cx, cy);

	CRect r;
	GetClientRect(r);

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		int msgViewPosition = pFrame->MsgViewPosition();
		BOOL bTreeHideVal = pFrame->TreeHideValue();
		BOOL isTreeHidden = pFrame->IsTreeHidden();

		TRACE("NMsgView::OnSize() WinWidth=%d WinHight=%d cx=%d cy=%d viewPos=%d IsTreeHideVal=%d IsTreeHidden=%d\n",
			r.Width(), r.Height(), cx, cy, msgViewPosition, bTreeHideVal, isTreeHidden);

		if (!pFrame->m_bIsTreeHidden)
		{
			m_frameCx_TreeNotInHide = cx;
			m_frameCy_TreeNotInHide = cy;
		}
		else
		{
			m_frameCx_TreeInHide = cx;
			m_frameCy_TreeInHide = cy;
		}
	}

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

	int nOffset = CalculateHigthOfMsgHdrPane();
	cx -= BSIZE*2;
	cy -= BSIZE*2;

	int m_attachmentWindowMaxSize = 25;
	AttachmentConfigParams *attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
	if (attachmentConfigParams)
	{
		m_attachmentWindowMaxSize = attachmentConfigParams->m_attachmentWindowMaxSize;
	}
	if (m_attachmentWindowMaxSize > 0)
	{

		// BEST effort to calculate size of attachment rectangle
		//int scrollVwidth = GetSystemMetrics(SM_CXVSCROLL);
		//int scrollHwidth = GetSystemMetrics(SM_CXHSCROLL);

		int aCnt = m_attachments.GetItemCount();
		int hS = 0;
		int vS = 0;
		m_attachments.GetItemSpacing(TRUE, &hS, &vS);

		RECT irc;
		int rectLen = 0;
		for (int ii = 0; ii < aCnt; ii++)
		{
			m_attachments.GetItemRect(ii, &irc, LVIR_BOUNDS);
			int w = irc.right - irc.left;
			if (w < hS)
				rectLen += hS;
			else
				rectLen += w;
		}

		int m_nAttachLines = 1;
		if (cx > 19)
			m_nAttachLines = rectLen / (cx - 19) + 1;
		else if (cx > 0)
			m_nAttachLines = rectLen / cx + 1;

		m_nAttachSize = m_nAttachLines * 19;
		if (m_nAttachLines > aCnt)
			m_nAttachSize = aCnt * 19;

		if ((m_nAttachSize < 44) && (aCnt > 2))
			m_nAttachSize = 44;
		else if (m_nAttachSize < 23)
			m_nAttachSize = 23;

		if ((cx > 0) && (rectLen > cx))
		{
			m_nAttachSize += 22;
		}

		int nMaxAttachSize = (int)((double)cy * m_attachmentWindowMaxSize / 100);
		if (nMaxAttachSize < 23)
			nMaxAttachSize = 23;

		if (m_nAttachSize > nMaxAttachSize)
		{
			int lines = (nMaxAttachSize - 23) / 17;
			if (lines < 1)
				lines = 1;
			int nEstimatedAttachSize = 23 + lines * 17;

			if ((cx > 0) && (rectLen > cx))
			{
				m_nAttachSize += 22;
			}

			if (nEstimatedAttachSize < m_nAttachSize)
				m_nAttachSize = nEstimatedAttachSize;
		}
}
	else
		m_nAttachSize = 0;

	int acy = m_bAttach ? m_nAttachSize : 0;

	int textWndOffset = nOffset;
	int hdrHight = 0;
	int hdrWidth = 0;
	if (m_hdrWindowLen > 0)
	{
		hdrHight = cy - nOffset;
		hdrWidth = cx + 1;
	}
	m_hdr.MoveWindow(BSIZE, BSIZE + textWndOffset, hdrWidth, hdrHight);

	int browserOffset = nOffset + hdrHight;

	int browserX = BSIZE;
	int browserY = BSIZE + browserOffset;
	int browserWidth = cx;
	int browserHight = cy - acy - browserOffset;
	if (m_hdrWindowLen > 0)
	{
		browserHight = 0;
		browserWidth = 0;
	}
	m_browser.MoveWindow(browserX, browserY, browserWidth, browserHight);

	int attachmentX = BSIZE;
	int attachmentY = cy - acy + BSIZE;
	//attachmentY = cy - acy;
	int attachmentWidth = cx;
	int attachmentHight = acy;
	if (m_hdrWindowLen > 0)
	{
		attachmentHight = 0;
		attachmentWidth = 0;
	}
	m_attachments.MoveWindow(attachmentX, attachmentY, attachmentWidth, attachmentHight);

	// TODO: seem to fix resizing issue; it should not be needed by iy seem to work
	Invalidate();
	UpdateWindow();
}

void NMsgView::UpdateLayout()
{
	CRect r;
	GetClientRect(r);

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		int msgViewPosition = pFrame->MsgViewPosition();
		BOOL bTreeHideVal = pFrame->TreeHideValue();
		BOOL isTreeHidden = pFrame->IsTreeHidden();

		int cx = 0;
		int cy = 0;
		TRACE("NMsgView::UpdateLayout() WinWidth=%d WinHight=%d cx=%d cy=%d viewPos=%d IsTreeHideVal=%d IsTreeHidden=%d\n",
			r.Width(), r.Height(), cx, cy, msgViewPosition, bTreeHideVal, isTreeHidden);
	}

	int cx = r.Width()-1;
	int cy = r.Height()-1;
	int nOffset = CalculateHigthOfMsgHdrPane();
	cx -= BSIZE*2;
	cy -= BSIZE*2;

	//int acy = m_bAttach ? m_nAttachSize : 0;
	//m_browser.MoveWindow(BSIZE, BSIZE+nOffset, cx, cy - acy - nOffset);
	//m_attachments.MoveWindow(BSIZE, cy-acy+BSIZE, cx, acy);

	// BEST effort to calculate size of attachment rectangle
	RECT rc;
	m_attachments.GetViewRect(&rc);
	m_nAttachSize = rc.bottom - rc.top;

	int m_attachmentWindowMaxSize = 25;
	AttachmentConfigParams *attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
	if (attachmentConfigParams)
	{
		m_attachmentWindowMaxSize = attachmentConfigParams->m_attachmentWindowMaxSize;
	}
	if (m_attachmentWindowMaxSize > 0)
	{
		if (m_nAttachSize < 23)
			m_nAttachSize = 23;

		if ((rc.right + 19) > r.Width())
			m_nAttachSize += 22;

		int nMaxAttachSize = (int)((double)cy * m_attachmentWindowMaxSize / 100);
		if (nMaxAttachSize < 23)
			nMaxAttachSize = 23;

		if (m_nAttachSize > nMaxAttachSize)
		{
			int lines = (nMaxAttachSize - 23) / 17;
			if (lines < 1)
				lines = 1;
			int nEstimatedAttachSize = 23 + lines * 17;
			if ((rc.right + 19) > r.Width())
				nEstimatedAttachSize += 22;

			if (nEstimatedAttachSize < m_nAttachSize)
				m_nAttachSize = nEstimatedAttachSize;
		}
	}
	else
		m_nAttachSize = 0;

	int acy = m_bAttach ? (m_nAttachSize) : 0;

	int textWndOffset = nOffset;
	int hdrHight = 0;
	int hdrWidth = 0;
	if (m_hdrWindowLen > 0)
	{
		hdrHight = cy - nOffset;
		hdrWidth = cx + 1;
	}
	m_hdr.MoveWindow(BSIZE, BSIZE + textWndOffset, hdrWidth, hdrHight);

	int browserOffset = nOffset + hdrHight;
	int browserX = BSIZE;
	int browserY = BSIZE + browserOffset;
	int browserWidth = cx;
	int browserHight = cy - acy - browserOffset;
	if (m_hdrWindowLen > 0)
	{
		browserHight = 0;
		browserWidth = 0;
	}
	m_browser.MoveWindow(browserX, browserY, browserWidth, browserHight);

	int attachmentX = BSIZE;
	int attachmentY = cy - acy + BSIZE;
	//attachmentY = cy - acy;
	int attachmentWidth = cx;
	int attachmentHight = acy;
	if (m_hdrWindowLen > 0)
	{
		attachmentHight = 0;
		attachmentWidth = 0;
	}
	m_attachments.MoveWindow(attachmentX, attachmentY, attachmentWidth, attachmentHight);


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
			BOOL ret = TextUtilsEx::id2charset(CnfCharsetId, str);
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

	if (CharsetId == 0)
	{
		if (CnfCharsetId)
			charsetId = CnfCharsetId;
		else
			charsetId = CP_UTF8;
	}

	if (charsetId) {
		CStringW strW;
		if (TextUtilsEx::Str2Wide(FieldText, charsetId, strW)) {
			BOOL retval = ::ExtTextOutW(hDC, xpos, ypos, ETO_CLIPPED, r, (LPCWSTR)strW, strW.GetLength(), NULL);
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
	if (m_hdrWindowLen > 0)
		m_hdr.SetWindowText(m_hdrData.Data());

	CPaintDC dc(this); // device context for painting
	static bool ft = true;

	NListView *pListView = 0;
	NMsgView *pMsgView = 0;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		pListView = pFrame->GetListView();
		pMsgView = pFrame->GetMsgView();
	}

	CRect cr;
	GetClientRect(cr);
	if (cr.IsRectNull())
		return;

	GetWindowRect(&m_rcCaption);
	ScreenToClient(&m_rcCaption);
	m_rcCaption.DeflateRect(BSIZE, 0);
	m_rcCaption.top = BSIZE;
	//if( ft ) 
	{
		ft = false;
		DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);
		if (color == COLOR_WHITE)
			dc.FillSolidRect(cr, ::GetSysColor(COLOR_WINDOW));
		else
		{
			if (pListView && !pListView->m_bApplyColorStyle)
				color = RGB(255, 255, 255);
			dc.FillSolidRect(cr, color);
		}
	}

	int nOffset = CalculateHigthOfMsgHdrPane();
	m_rcCaption.bottom = m_rcCaption.top + nOffset;

	DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessageHeader);
	if (color == COLOR_WHITE)
	{
		dc.FillSolidRect(m_rcCaption, ::GetSysColor(COLOR_WINDOW));
	}
	else
	{
		if (pListView && !pListView->m_bApplyColorStyle)
			color = RGB(255, 255, 255);
		dc.FillSolidRect(m_rcCaption, color);
	}
	dc.Draw3dRect(m_rcCaption, ::GetSysColor(COLOR_BTNHIGHLIGHT), ::GetSysColor(COLOR_BTNSHADOW));

	if (m_bMax)
	{
		CRect	r = m_rcCaption;

		if( r.Width() > 0 ) 
		{
			CFont *pOldFont = dc.SelectObject(&m_BoldFont);

			dc.SetBkMode(TRANSPARENT);

			if (m_hdrPaneLayout == 0)
			{
				int xpos = 0;
				int ypos = r.top + TEXT_LINE_TWO;
				UINT localCP = GetACP();
				std::string str;
				BOOL ret = TextUtilsEx::id2charset(localCP, str);
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
			}
			else
			{
				UINT localCP = GetACP();
				std::string str;
				BOOL ret = TextUtilsEx::id2charset(localCP, str);
				m_date_charset = str.c_str();

				int xpos = 0;
				int ypos = r.top + 4;
				xpos = PaintHdrField(dc, r, xpos, ypos, TRUE, m_strTitleSubject, m_strSubject, m_subj_charset, m_subj_charsetId, m_cnf_subj_charsetId);

				BOOL largeFont = TRUE;
				xpos = 0;
				ypos += 3 + 18;
				xpos = PaintHdrField(dc, r, xpos, ypos, largeFont, m_strTitleFrom, m_strFrom, m_from_charset, m_from_charsetId, m_cnf_from_charsetId);

				xpos = 0;
				ypos += 3 + 18;
				xpos = PaintHdrField(dc, r, xpos, ypos, largeFont, m_strTitleTo, m_strTo, m_to_charset, m_to_charsetId, m_cnf_to_charsetId);

				if (!m_strCC.IsEmpty())
				{
					xpos = 0;
					ypos += 3 + 18;
					xpos = PaintHdrField(dc, r, xpos, ypos, largeFont, m_strTitleCC, m_strCC, m_cc_charset, m_cc_charsetId, m_cnf_cc_charsetId);
				}

				if (!m_strBCC.IsEmpty())
				{
					xpos = 0;
					ypos += 3 + 18;
					xpos = PaintHdrField(dc, r, xpos, ypos, largeFont, m_strTitleBCC, m_strBCC, m_bcc_charset, m_bcc_charsetId, m_cnf_bcc_charsetId);
				}

				xpos = 0;
				ypos += 3 + 18;
				xpos = PaintHdrField(dc, r, xpos, ypos, largeFont, m_strTitleDate, m_strDate, m_date_charset, localCP, m_cnf_date_charsetId);
				xpos += TEXT_BOLD_LEFT;
			}

			// Rstore font
			dc.SelectObject(pOldFont);
		}
	}
	FrameGradientFill( &dc, cr, BSIZE, RGB(0x8a, 0x92, 0xa6), RGB(0x6a, 0x70, 0x80));
	
	// Do not call CWnd::OnPaint() for painting messages
}

BOOL NMsgView::OnEraseBkgnd(CDC* pDC) 
{
	// TODO: Add your message handler code here and/or call default

	return  TRUE;
	//return CWnd::OnEraseBkgnd(pDC);
}

void NMsgView::OnLButtonDblClk(UINT nFlags, CPoint point) 
{
#if 1
	ClearSearchResultsInIHTMLDocument(m_searchID);
#else
	CString path = FileUtils::GetmboxviewTempPath();
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

int FindMenuItem(CMenu* Menu, LPCTSTR MenuName) 
{
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

void NMsgView::ClearSearchResultsInIHTMLDocument(CString &searchID)
{
	HtmlUtils::ClearSearchResultsInIHTMLDocument(m_browser, searchID);
}

void NMsgView::FindStringInIHTMLDocument(CString &searchText, BOOL matchWord, BOOL matchCase)
{
	HtmlUtils::FindStringInIHTMLDocument(m_browser, m_searchID, searchText, matchWord, matchCase, m_matchStyle);
}

void NMsgView::PrintHTMLDocumentToPrinter(SimpleString *inbuf, SimpleString *workbuf, UINT inCodePage)
{
	BOOL bPrintDialogType = 1;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
		bPrintDialogType = pFrame->m_NamePatternParams.m_bPrintDialog;

		HtmlUtils::PrintHTMLDocumentToPrinter(inbuf, workbuf, inCodePage, bPrintDialogType);
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


void NMsgView::OnClose()
{
	// TODO: Add your message handler code here and/or call default

	CWnd::OnClose();
}

void NMsgView::OnRButtonDown(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default

	const char *ClearFindtext = _T("Clear Find Text");
	const char *CustomColor = _T("Custom Color Style");
	const char *ShowMailHdr = _T("View Raw Message Header");
	const char *HdrPaneLayout = _T("Header Pane Layout");

	NListView *pListView = 0;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		pListView = pFrame->GetListView();
	}

	CPoint pt;
	::GetCursorPos(&pt);
	CWnd *wnd = WindowFromPoint(pt);

	CMenu menu;
	menu.CreatePopupMenu();
	menu.AppendMenu(MF_SEPARATOR);

	const UINT M_CLEAR_FIND_TEXT_Id = 1;
	AppendMenu(&menu, M_CLEAR_FIND_TEXT_Id, ClearFindtext);

	const UINT M_ENABLE_DISABLE_COLOR_Id = 2;
	AppendMenu(&menu, M_ENABLE_DISABLE_COLOR_Id, CustomColor, pListView->m_bApplyColorStyle);

	const UINT M_SHOW_MAIL_HEADER_Id = 3;
	AppendMenu(&menu, M_SHOW_MAIL_HEADER_Id, ShowMailHdr, m_hdrWindowLen > 0);

#if 0
	// Added global option under View
	CMenu hdrLayout;
	hdrLayout.CreatePopupMenu();
	hdrLayout.AppendMenu(MF_SEPARATOR);

	const UINT S_DEFAULT_LAYOUT_Id = 4;
	AppendMenu(&hdrLayout, S_DEFAULT_LAYOUT_Id, _T("Default"), m_hdrPaneLayout == 0);

	const UINT S_EXPANDED_LAYOUT_Id = 5;
	AppendMenu(&hdrLayout, S_EXPANDED_LAYOUT_Id, _T("Expanded"), m_hdrPaneLayout == 1);

	menu.AppendMenu(MF_POPUP | MF_STRING, (UINT)hdrLayout.GetSafeHmenu(), _T("Header Fields Layout"));
	menu.AppendMenu(MF_SEPARATOR);
#endif

	int command = menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, this);

	UINT n_Flags = TPM_RETURNCMD;
	CString menuString;
	int chrCnt = menu.GetMenuString(command, menuString, n_Flags);

	if (command == M_CLEAR_FIND_TEXT_Id)
	{
		ClearSearchResultsInIHTMLDocument(m_searchID);
		int deb = 1;
	}
	else if ((command == M_ENABLE_DISABLE_COLOR_Id))
	{
		if (pListView)
		{
			pListView->m_bApplyColorStyle = !pListView->m_bApplyColorStyle;
			int iItem = pListView->m_lastSel;
			if (m_hdrWindowLen)
			{
				this->ShowMailHeader(iItem);
			}
			else
			{
				pListView->Invalidate();
				pListView->SelectItem(pListView->m_lastSel, TRUE);
			}
		}
		int deb = 1;
	}
	else if ((command == M_SHOW_MAIL_HEADER_Id))
	{
		int iItem = pListView->m_lastSel;
		if (m_hdrWindowLen)
		{
			this->HideMailHeader(iItem);
			pListView->Invalidate();
			pListView->SelectItem(pListView->m_lastSel, TRUE);
		}
		else
		{
			this->ShowMailHeader(iItem);
		}
		int deb = 1;
	}
#if 0
	// Added global option under View
	else if ((command == S_DEFAULT_LAYOUT_Id))
	{
		m_hdrPaneLayout = 0;
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("headerPaneLayout"), m_hdrPaneLayout);
		Invalidate();
		UpdateLayout();
		int deb = 1;
	}
	else if ((command == S_EXPANDED_LAYOUT_Id))
	{
		m_hdrPaneLayout = 1;
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("headerPaneLayout"), m_hdrPaneLayout);
		Invalidate();
		UpdateLayout();
		int deb = 1;
	}
#endif

	CWnd::OnRButtonDown(nFlags, point);
}

char * NMsgView::EatFldLine(char *p, char *e)
{
	// Fld found, eat the rest
	while ((p < e) && (*p++ != '\n'));
	while (p < e)
	{
		if ((*p == ' ') || (*p == '\t'))
			while ((p < e) && (*p++ != '\n'));
		else
			break;
	}
	return p;
}

int NMsgView::FindMailHeader(char *data, int datalen)
{
	static const char *cFromMailBegin = "From ";
	static const int cFromMailBeginLen = strlen(cFromMailBegin);

	char *p = data;
	char *e = data + datalen;
	char *fldName = p;
	char *pend = e;

	p = MimeParser::SkipEmptyLines(p, e);

	if (TextUtilsEx::strncmpExact(p, e, cFromMailBegin, cFromMailBeginLen) == 0)
	{
		p = MimeParser::EatNewLine(p, e);
	}
	fldName = p;
	pend = p;

	// Too Simplistic ?? 
	// Parsing mbox file to create index file relies on empty lines for better or worse
	// I think I have seen one case that simplistic approach would fails but can't find any now
	// TODO: add code to compare these two methods
	BOOL simpleApproach = TRUE; 
	if (simpleApproach)
	{
		while (p < e)
		{
			if ((*p == '\n') || ((*p == '\r') && (*(p+1) == '\n')))
				break;
			else
				p = MimeParser::EatNewLine(p, e);
		}
		int headerLength = -1;
		if (p < e)
			headerLength = p - data;
		return headerLength;
	}
	else
	{
		// may find one or two extar lines
		int fldNameCnt = 0;
		char c;
		while (p < e)
		{
			c = *p;
			if ((c == '\r') || (c == '\n'))
			{
				if (fldNameCnt > 0)
				{
					p++;
					break;
				}
				else
				{
					pend = p;
					fldNameCnt = 0;
					p++;
					continue;
				}
			}

			c = *p;
			if (c == ':')
			{
				if (fldNameCnt == 0)
				{
					p = MimeParser::EatNewLine(p, e);
					break;
				}
				else
				{
					p = EatFldLine(p, e);
					pend = p;
					fldName = p;
					fldNameCnt = 0;
					continue;
				}
			}
			else if ((c == ' ') || (c == '\t'))
			{
				while ((p < e) && (*p != '\n') && ((c == ' ') || (c == '\t')))
				{
					p++;
				}

				c = *p;
				if (c == ':')
				{
					if (fldNameCnt == 0)
					{
						p = MimeParser::EatNewLine(p, e);
						break;
					}
					else
					{
						p = EatFldLine(p, e);
						pend = p;
						fldNameCnt = 0;
						fldName = p;
						continue;
					}
				}
				else
				{
					p = MimeParser::EatNewLine(p, e);
					fldNameCnt = 0;
					break;
				}
			}
			else
			{
				fldNameCnt++;
				p++;
			}
		}

		int headerLength = -1;
		if (pend < e)
			headerLength = pend - data;
		return headerLength;
	}
}

int NMsgView::ShowMailHeader(int mailPosition)
{
	if ((mailPosition >= MboxMail::s_mails.GetCount()) || (mailPosition < 0))
		return -1;

	MboxMail *m = MboxMail::s_mails[mailPosition];
	if (m == 0)
	{
		// TODO: Assert ??
		return 1;
	}

	m->GetBody(&m_hdrDataTmp, 128*1024);

	int headerlen = 0;
	if (m->m_headLength > 0)
		headerlen = m->m_headLength;
	else
		headerlen = FindMailHeader(m_hdrDataTmp.Data(), m_hdrDataTmp.Count());

	if (headerlen > 0)
	{
		m->m_headLength = headerlen;

		m_hdrDataTmp.SetCount(headerlen);
		char *p = m_hdrDataTmp.Data();
		// CFile Read removes CR and therfore we need to add CR, great  !!
		char *ms = strchr(p, '\n');
		if (ms) 
		{
			BOOL bAddCR = FALSE;
			if (*(ms - 1) != '\r')
				bAddCR = TRUE;
			int pos = ms - p + 1;
			if (bAddCR)
			{
				SimpleString *tmpbuf = MboxMail::get_tmpbuf();
				TextUtilsEx::ReplaceNL2CRNL((LPCSTR)m_hdrDataTmp.Data(), m_hdrDataTmp.Count(), tmpbuf);
				m_hdrData.Clear();
				m_hdrData.Append("\r\n");
				m_hdrData.Append(tmpbuf->Data(), tmpbuf->Count());
				MboxMail::rel_tmpbuf();
			}
			else
			{
				m_hdrData.Clear();
				m_hdrData.Append("\r\n");
				m_hdrData.Append(m_hdrDataTmp.Data(), m_hdrDataTmp.Count());
			}
		}
	}
	else
		m_hdrData.SetCount(0);

	m_hdr.ShowWindow(SW_SHOW);
	m_hdr.ShowWindow(SW_RESTORE);

	m_hdrWindowLen = 100;
	Invalidate();
	UpdateLayout();

	NListView *pListView = 0;
	NMsgView *pMsgView = 0;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		pListView = pFrame->GetListView();
		if (pListView)
		{
			DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);
			if (m_hdrData.Count() == 0)
				pListView->m_bApplyColorStyle = TRUE;
			if (pListView->m_bApplyColorStyle && (color != COLOR_WHITE))
			{
				m_hdr.SetBkColor(color);
			}
			else
				m_hdr.SetBkColor(COLOR_WHITE);

		}
	}
	return 1;
}

int NMsgView::HideMailHeader(int mailPosition)
{
	m_hdrWindowLen = 0;
	m_hdrData.SetCount(0);
	Invalidate();
	UpdateLayout();
	return 1;
}


void  NMsgView::DisableMailHeader()
{
	m_hdrData.Clear();
	m_hdrWindowLen = 0;
}


int NMsgView::SetMsgHeader(int mailPosition, int gmtTime, CString &format)
{
	if ((mailPosition >= MboxMail::s_mails.GetCount()) || (mailPosition < 0))
		return -1;

	MboxMail *m = MboxMail::s_mails[mailPosition];
	if (m == 0)
	{
		// TODO: Assert ??
		return 1;
	}

	// Set header data
	m_strSubject = m->m_subj;
	m_strFrom = m->m_from;
	m_strTo = m->m_to;
	m_strCC = m->m_cc;
	m_strBCC = m->m_bcc;
	//
	//CTime tt(m->m_timeDate);
	//pMsgView->m_strDate = tt.Format(m_format);
	if (m->m_timeDate >= 0)
	{
		MyCTime tt(m->m_timeDate);
		if (!gmtTime)
		{
			m_strDate = tt.FormatLocalTm(format);
		}
		else {
			m_strDate = tt.FormatGmtTm(format);
		}
	}
	else
		m_strDate = "";

	//
	m_subj_charsetId = m->m_subj_charsetId;
	m_subj_charset = m->m_subj_charset;
	//
	m_from_charsetId = m->m_from_charsetId;
	m_from_charset = m->m_from_charset;
	//
	m_to_charsetId = m->m_to_charsetId;
	m_to_charset = m->m_to_charset;
	//
	m_cc_charsetId = m->m_cc_charsetId;
	m_cc_charset = m->m_cc_charset;
	//
	m_bcc_charsetId = m->m_bcc_charsetId;
	m_bcc_charset = m->m_bcc_charset;
	//
	m_date_charsetId = 0;

	return 1;
}

int NMsgView::CalculateHigthOfMsgHdrPane()
{
	int m_subjectTextHight = 18;
	int m_otherTextHight = 18;

	int pane_Hight = 0;

	if (m_hdrPaneLayout == 0)
	{
		pane_Hight = m_bMax ? CAPT_MAX_HEIGHT : CAPT_MIN_HEIGHT;
	}
	else
	{
		pane_Hight = 
			  4 + m_subjectTextHight  // subject
			+ 3 + m_otherTextHight  // from
			+ 3 + m_otherTextHight  // to
			;
		if (!m_strCC.IsEmpty())
			pane_Hight += 3 + m_otherTextHight;  // cc

		if (!m_strBCC.IsEmpty())
			pane_Hight += 3 + m_otherTextHight;  // bcc

		pane_Hight =
			pane_Hight += 3 + m_subjectTextHight + 4; // date

	}
	return pane_Hight;
}

void NMsgView::OnMessageheaderpanelayoutDefault()
{
	// TODO: Add your command handler code here
	m_hdrPaneLayout = 0;
	CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("headerPaneLayout"), m_hdrPaneLayout);
	Invalidate();
	UpdateLayout();
	int deb = 1;
}


void NMsgView::OnMessageheaderpanelayoutExpanded()
{
	// TODO: Add your command handler code here
	m_hdrPaneLayout = 1;
	CProfile::_WriteProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, _T("headerPaneLayout"), m_hdrPaneLayout);
	Invalidate();
	UpdateLayout();
	int deb = 1;
}

int NMsgView::PreTranslateMessage(MSG* pMsg)
{
	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
	{
		if (pMsg->wParam == VK_ESCAPE)
		{
			AfxGetMainWnd()->PostMessage(WM_CLOSE);            // Do not process further
		}
	}
	return CWnd::PreTranslateMessage(pMsg);
}
