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
#include "ResHelper.h"

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
	:m_attachments(this), m_hdr(), m_menu(CString(L"MailHdr"))
{
	m_hdrPaneLayout = 0;
	DWORD hdrPaneLayout = 0;

	m_hdrWindowLen = 0;
	m_bAttach = FALSE;
	m_nAttachSize = 50;
	m_bMax = TRUE;

	m_searchID = L"mboxview_Search";
	m_matchStyle = L"color: white; background-color: blue";

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
	ncm.lfMessageFont.lfHeight = -MulDiv(12, GetDeviceCaps(hdc, LOGPIXELSY), 72);
	//ncm.lfMessageFont.lfHeight = -MulDiv(11, GetDeviceCaps(hdc, LOGPIXELSY), 72);
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
	m_strMailHeader.LoadString(IDS_TITLE_MAIL_HEADER);

	ResHelper::TranslateString(m_strTitleSubject);
	ResHelper::TranslateString(m_strTitleFrom);
	ResHelper::TranslateString(m_strTitleDate);
	ResHelper::TranslateString(m_strTitleTo);
	ResHelper::TranslateString(m_strTitleCC);
	ResHelper::TranslateString(m_strTitleBCC);
	ResHelper::TranslateString(m_strTitleBody);

	// set by SelectItem in NListView
	m_subj_charsetId = 0;
	m_from_charsetId = 0;
	m_date_charsetId = 0;
	m_to_charsetId = 0;
	m_cc_charsetId = 0;
	m_bcc_charsetId = 0;
	m_body_charsetId = 0;
	//
	m_mail_header_charsetId = 0;
	m_body_text_charsetId = 0;
	m_body_html_charsetId = 0;

	CString section_options = CString(sz_Software_mboxview) + L"\\Options";

	m_cnf_subj_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"subjCharsetId");
	m_cnf_from_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"fromCharsetId");
	m_cnf_date_charsetId = 0;
	m_cnf_to_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"toCharsetId");
	m_cnf_cc_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"ccCharsetId");
	m_cnf_bcc_charsetId = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"bccCharsetId");
	m_show_charsets = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"showCharsets");

#if 0
	if (m_cnf_from_charsetId == 0)
		m_cnf_from_charsetId = GetACP();

	if (m_cnf_to_charsetId == 0)
		m_cnf_to_charsetId = GetACP();

	if (m_cnf_subj_charsetId == 0)
		m_cnf_subj_charsetId = GetACP();
#endif

	DWORD bImageViewer;
	BOOL retval = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_options, L"imageViewer", bImageViewer);
	if (retval == TRUE) {
		m_bImageViewer = bImageViewer;
	}
	else {
		bImageViewer = 1;
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_options, L"imageViewer", bImageViewer);
		m_bImageViewer = bImageViewer;
	}

	m_frameCx_TreeNotInHide = 700;
	m_frameCy_TreeNotInHide = 200;
	m_frameCx_TreeInHide = 700;
	m_frameCy_TreeInHide = 200;

	CString section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacement";
#if 0
	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
		section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacementPreview";
	else if (CMainFrame::m_commandLineParms.m_bDirectFileOpenMode)
		section_wnd = CString(sz_Software_mboxview) + L"\\WindowPlacementDirect";
#endif

	BOOL ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_wnd, L"MsgFrameTreeNotHiddenWidth", m_frameCx_TreeNotInHide);
	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_wnd, L"MsgFrameTreeNotHiddenHeight", m_frameCy_TreeNotInHide);
	//
	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_wnd, L"MsgFrameTreeHiddenWidth", m_frameCx_TreeInHide);
	ret = CProfile::_GetProfileInt(HKEY_CURRENT_USER, section_wnd, L"MsgFrameTreeHiddenHeight", m_frameCy_TreeInHide);
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
	ON_MESSAGE(WM_CMD_PARAM_ON_SIZE_MSGVIEW_MESSAGE, &NMsgView::OnCmdParam_OnSizeMsgView)
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

	if (!m_attachments.Create(WS_CHILD | WS_VISIBLE | LVS_SINGLESEL | LVS_SMALLICON | LVS_SHOWSELALWAYS | LVS_AUTOARRANGE, 
		CRect(), this, IDC_ATTACHMENTS))
		return -1;

	m_attachments.SetExtendedStyle(LVS_EX_LABELTIP | m_attachments.GetExtendedStyle());

	m_attachments.SendMessage((CCM_FIRST + 0x7), 5, 0);  // #define CCM_SETVERSION          (CCM_FIRST + 0x7)
	m_attachments.SetTextColor(RGB(0, 0, 0));
	if (!m_font.CreatePointFont(85, L"Tahoma"))
		if (!m_font.CreatePointFont(85, L"Verdana"))
			m_font.CreatePointFont(85, L"Arial");

	m_attachments.SetFont(&m_font);
	CImageList sysImgList;
	SHFILEINFO shFinfo;

	sysImgList.Attach((HIMAGELIST)SHGetFileInfo(L"C:\\",
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
	// Not working well. When completed, message is send to invoke NMsgView::UpdateLayout()
	// Maybe I will fix that one day
	int deb1, deb2, deb3, deb4;
	CWnd::OnSize(nType, cx, cy);

	TRACE(L"OnSize: c=%d cy=%d\n", cx, cy);

	CRect r;
	GetClientRect(r);

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		int msgViewPosition = pFrame->MsgViewPosition();
		BOOL bTreeHideVal = pFrame->TreeHideValue();
		BOOL isTreeHidden = pFrame->IsTreeHidden();

		TRACE(L"NMsgView::OnSize() WinWidth=%d WinHight=%d cx=%d cy=%d viewPos=%d IsTreeHideVal=%d IsTreeHidden=%d\n",
			r.Width(), r.Height(), cx, cy, msgViewPosition, bTreeHideVal, isTreeHidden);

#if 0
		// No lnger needed, delete.
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
#endif
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

	int m_attachmentWindowMaxSize = 25;  // in %
	AttachmentConfigParams* attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
	if (attachmentConfigParams)
	{
		m_attachmentWindowMaxSize = attachmentConfigParams->m_attachmentWindowMaxSize;
	}
	if (m_attachmentWindowMaxSize > 0)
	{
		// tried to leverage the following:
		// m_attachments.GetViewRect(&rc);
		// m_attachments.ApproximateViewRect();
		// but it didn't work as expected
		// CalculateViewRec(rc, cx, cy); is the best effort solution

		CRect rcview;
		m_attachments.GetViewRect(&rcview);

		int nAttachSize = rcview.Height();
		if (rcview.Width() > cx)
			nAttachSize += 18;

		CSize proposedViewRec(cx, 300);
		CSize approxViewRec = m_attachments.ApproximateViewRect(proposedViewRec, m_attachments.GetItemCount());

		CRect rc;
		CalculateViewRec(rc, cx, cy);
		m_nAttachSize = rc.bottom - rc.top;

		int nMaxAttachSize = (int)((double)cy * m_attachmentWindowMaxSize / 100);
		if (nMaxAttachSize < 23)
			nMaxAttachSize = 23;

		if (m_nAttachSize > nMaxAttachSize)
			m_nAttachSize = nMaxAttachSize;
	}
	else
		m_nAttachSize = 0;

	int acy = m_bAttach ? m_nAttachSize : 0;

	int textWndOffset = nOffset;   // int nOffset = CalculateHigthOfMsgHdrPane();
	int hdrHight = 0;
	int hdrWidth = 0;
	if (m_hdrWindowLen > 0)
	{
		hdrHight = cy - nOffset;  // show header, text block, etc mode
		hdrWidth = cx + 1;
	}
	m_hdr.MoveWindow(BSIZE, BSIZE + textWndOffset, hdrWidth, hdrHight); // show/shrink header, text block, etc area (CMenuEdit)

	int browserOffset = nOffset + hdrHight;

	int browserX = BSIZE;
	int browserY = BSIZE + browserOffset;
	int browserWidth = cx;

	int browserHight = cy - acy - browserOffset;
	if (m_attachments.GetItemCount())  // FIXME
		browserHight -= BSIZE;
	if (m_hdrWindowLen > 0)
	{
		browserHight = 0;
		browserWidth = 0;
	}
	m_browser.MoveWindow(browserX, browserY, browserWidth, browserHight);

	int attachmentX = BSIZE;
	int attachmentY = cy - acy + BSIZE;  // FIXME
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

void NMsgView::CalculateViewRec(CRect& rc, int cx, int cy)
{
	TRACE("\n\nCalculateViewRec\n");

	TRACE("cx=%d cy=%d rc.LT=%d,%d rc.RB=%d,%d rc.w=%d rc.h=%d\n",
		cx, cy, rc.left, rc.top, rc.right, rc.bottom, rc.Width(), rc.Height());

	CRect rcview;
	m_attachments.GetViewRect(&rcview); //  doesn't seem to work well

	TRACE("GetViewRect: cx=%d cy=%d rc.LT=%d,%d rc.RB=%d,%d rc.w=%d rc.h=%d\n",
		cx, cy, rcview.left, rcview.top, rcview.right, rcview.bottom, rcview.Width(), rcview.Height());

	int aCnt = m_attachments.GetItemCount();
	int hS = 0;
	int vS = 0;
	m_attachments.GetItemSpacing(TRUE, &hS, &vS);

	CRect irc;
	m_attachments.GetItemRect(0, &irc, LVIR_BOUNDS);

	int itemHight = irc.Height();

	CPaintDC dc(this);
	HDC hDC = dc.GetSafeHdc();

	int len = 0;
	int totalLen = 0;
	POINT pt;
	int longLineCnt = 0;
	TRACE("\n\nCalculateViewRec\n");
	for (int ii = 0; ii < aCnt; ii++)
	{
		m_attachments.GetItemRect(ii, &irc, LVIR_BOUNDS);
		m_attachments.GetItemPosition(ii, &pt);

		TRACE("cx=%d cy=%d irc.LT=%d,%d irc.RB=%d,%d irc.w=%d irc.h=%d\n",
			cx, cy, irc.left, irc.top, irc.right, irc.bottom, irc.Width(), irc.Height());
		TRACE("pt.x=%d pt.y=%d\n", pt.x, pt.y);

		if ((irc.right + 4) > cx)
			longLineCnt++;

		SIZE sizeItem;
		CString name = m_attachments.GetItemText(ii, 0);
		BOOL retA = GetTextExtentPoint32(hDC, name, name.GetLength(), &sizeItem);

		int sizeExtra = (sizeItem.cx % hS) ? hS : 0;
		int cxItem = (sizeItem.cx / hS) * hS + sizeExtra;

		int itemWidth = sizeItem.cx;
		totalLen += itemWidth;
	}

	int extraLine = (totalLen % (cx-18)) ? 1 : 0;
	int lcnt = totalLen / (cx-18) + extraLine;
	if (longLineCnt)
		lcnt++;

	int height = lcnt * 18 + BSIZE;
	if (height < 23)
		height = 23;

	rc.SetRect(0,0,cx, height);

	// Doesn't work well. Calculation of height are not accurate; force retry. Will invoke  UpdateLayout !!!!
	LRESULT lres = PostMessage(WM_CMD_PARAM_ON_SIZE_MSGVIEW_MESSAGE, 0, 0);
}

void NMsgView::UpdateLayout()
{
	// Works Ok. May need to make small improvements on size and positions of hdr, browser, attachmen panes
	CRect r;
	GetClientRect(r);

	TRACE(L"UpdateLayout: c=%d cy=%d\n", r.Width(), r.Height());

	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		int msgViewPosition = pFrame->MsgViewPosition();
		BOOL bTreeHideVal = pFrame->TreeHideValue();
		BOOL isTreeHidden = pFrame->IsTreeHidden();

		int cx = r.Width();
		int cy = r.Height();
		TRACE(L"NMsgView::UpdateLayout() WinWidth=%d WinHight=%d cx=%d cy=%d viewPos=%d IsTreeHideVal=%d IsTreeHidden=%d\n",
			r.Width(), r.Height(), cx, cy, msgViewPosition, bTreeHideVal, isTreeHidden);
	}

	int cx = r.Width()-1;
	int cy = r.Height()-1;

	int mailHdrFieldsPaneHight = CalculateHigthOfMsgHdrPane();

	cx -= BSIZE*2;
	cy -= BSIZE*2;

	int m_attachmentWindowMaxSize = 25;  // in %
	AttachmentConfigParams *attachmentConfigParams = CMainFrame::GetAttachmentConfigParams();
	if (attachmentConfigParams)
	{
		m_attachmentWindowMaxSize = attachmentConfigParams->m_attachmentWindowMaxSize;
	}
	if (m_attachmentWindowMaxSize > 0)
	{
		CRect rcview;
		m_attachments.GetViewRect(&rcview);  // works best when window is nor resized

		m_nAttachSize = rcview.Height();
		if (rcview.Width() > cx)
			m_nAttachSize += 18;

		int nMaxAttachSize = (int)((double)cy * m_attachmentWindowMaxSize / 100);
		if (nMaxAttachSize < 23)
			nMaxAttachSize = 23;

		if (m_nAttachSize > nMaxAttachSize)
			m_nAttachSize = nMaxAttachSize;
	}
	else
		m_nAttachSize = 0;

	int acy = m_bAttach ? (m_nAttachSize) : 0;

	int hdrHight = 0;
	int hdrWidth = 0;
	if (m_hdrWindowLen > 0)   // show header, text block, etc mode
	{
		hdrHight = cy - mailHdrFieldsPaneHight;
		hdrWidth = cx + 1;
	}
	m_hdr.MoveWindow(BSIZE, BSIZE + mailHdrFieldsPaneHight, hdrWidth, hdrHight);  // show/shrink header, text block, etc area (CMenuEdit)

	int browserOffset = mailHdrFieldsPaneHight + hdrHight;
	int browserX = BSIZE;
	int browserY = BSIZE + browserOffset;
	int browserWidth = cx;
	int browserHight = cy - acy - browserOffset;
	if (m_attachments.GetItemCount())
		browserHight -= 2*BSIZE;
	if (m_hdrWindowLen > 0)
	{
		browserHight = 0;
		browserWidth = 0;
	}
	m_browser.MoveWindow(browserX, browserY, browserWidth, browserHight);

	int attachmentX = BSIZE;
	int attachmentY = cy - acy + BSIZE;
	attachmentY = cy - acy;
	int attachmentWidth = cx;
	int attachmentHight = acy + 2;
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

int NMsgView::PaintHdrField(CPaintDC &dc, CRect	&r, int x_pos, int y_pos, BOOL bigFont, CString &FieldTitle,  CStringA &FieldText, CStringA &Charset, UINT CharsetId, UINT CnfCharsetId)
{
	DWORD error;

	HDC hDC = dc.GetSafeHdc();

	int xpos = x_pos;
	int ypos = y_pos;

	if (bigFont)
		dc.SelectObject(&m_BigBoldFont);
	else
		dc.SelectObject(&m_BoldFont);

	xpos += TEXT_BOLD_LEFT;
	BOOL retTextOut = dc.ExtTextOut(xpos, ypos, ETO_CLIPPED, r, FieldTitle, NULL);
	CSize szFieldTitle = dc.GetTextExtent(FieldTitle);

	if (bigFont)
		dc.SelectObject(&m_BigFont);
	else
		dc.SelectObject(&m_NormFont);

	UINT charsetId = CharsetId;
	CStringA StrCharset;
	CStringA strCharSetId;
	if (m_show_charsets)
	{
		if ((CharsetId == 0) && (CnfCharsetId != 0))
		{
			std::string str;
			BOOL ret = TextUtilsEx::id2charset(CnfCharsetId, str);
			CStringA charset = str.c_str();

			strCharSetId.Format("%u", CnfCharsetId);
			StrCharset += " (" + charset + "/" + strCharSetId + "*)";

			charsetId = CnfCharsetId;
		}
		else
		{
			strCharSetId.Format("%u", CharsetId);
			StrCharset += " (" + Charset + "/" + strCharSetId + ")";
		}
		xpos += szFieldTitle.cx;

		CString StrCharsetW;
		int pageCode = CP_ACP;
		BOOL retS2W = TextUtilsEx::Str2WStr(StrCharset, pageCode, StrCharsetW, error);

		xpos += 0;
		BOOL retval = dc.ExtTextOut(xpos, ypos, ETO_CLIPPED, r, StrCharsetW, NULL);
		CSize szStrCharset = dc.GetTextExtent(StrCharsetW);
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

	CSize szFieldText = { 0,0 };
	if (charsetId)
	{
		CStringW strW;
		if (TextUtilsEx::Str2WStr(FieldText, charsetId, strW, error))
		{
			BOOL retval = dc.ExtTextOut(xpos, ypos, ETO_CLIPPED, r, strW, NULL);
			szFieldText = dc.GetTextExtent(strW);
			done = TRUE;
		}
	}
	_ASSERTE(done);
	if (!done)
	{
		// I guess TextUtilsEx::Str2WStr above failed so we call ExtTextOutA instead of  ExtTextOut
		BOOL retval = ::ExtTextOutA(hDC, xpos, ypos, ETO_CLIPPED, r, (LPCSTR)FieldText, FieldText.GetLength(), NULL);
		CSize szText = { 0,0 };
		BOOL ret = GetTextExtentPoint32A(hDC, (LPCSTR)FieldText, FieldText.GetLength(), &szText);
		if (ret)
			szFieldText = szText;
	}

	xpos += szFieldText.cx;
	return xpos;
}

void NMsgView::OnPaint()
{
	if (m_hdrWindowLen > 0)
		m_hdr.SetWindowText((LPCWSTR)m_hdrData.Data());

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
				//xpos += TEXT_BOLD_LEFT;

				if (m_show_charsets) {
					xpos = 0;
					ypos += 3 + 18;
					xpos = PaintHdrField(dc, r, xpos, ypos, largeFont, m_strTitleBody, m_strBody, m_body_charset, m_body_charsetId, m_body_charsetId);
				}
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
	ClearSearchResultsInIHTMLDocument(m_searchID);
}

int FindMenuItem(CMenu* Menu, LPCWSTR MenuName) 
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

int AttachIcon(CMenu* Menu, LPCWSTR MenuName, int resourceId, CBitmap  &cmap)
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

void NMsgView::ClearSearchResultsInIHTMLDocument(CString& searchID)
{
	HtmlUtils::ClearSearchResultsInIHTMLDocument(m_browser, searchID);
}

void NMsgView::FindStringInIHTMLDocument(CString& searchText, BOOL matchWord, BOOL matchCase)
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

const UINT M_CLEAR_FIND_TEXT_Id = 1;
const UINT M_ENABLE_DISABLE_COLOR_Id = 2;
const UINT M_SHOW_MAIL_HEADER_Id = 3;
const UINT M_SHOW_MAIL_HTML_AS_TEXT_Id = 4;
const UINT M_SHOW_MAIL_TEXT_Id = 5;
const UINT M_SHOW_MAIL_HTML_Id = 6;

static BOOL nMenuInitialed = FALSE;

UINT ToggleMenuCheckState(CMenu *menu, UINT commandID)
{
	UINT state = menu->GetMenuState(commandID, MF_BYCOMMAND);
	ASSERT(state != 0xFFFFFFFF);

	UINT retState;
	if (state & MF_CHECKED)
		retState = menu->CheckMenuItem(commandID, MF_UNCHECKED | MF_BYCOMMAND);
	else
		retState = menu->CheckMenuItem(commandID, MF_CHECKED | MF_BYCOMMAND);
	return retState;
}

void NMsgView::SetMenuToInitialState()
{
	if (!nMenuInitialed)
		return;

	NListView* pListView = 0;
	CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		pListView = pFrame->GetListView();
	}

	if ((pFrame == 0) || (pListView == 0))
		return;

	UINT retState;
	if (pFrame->m_bViewMessageHeaders == TRUE)
		retState = m_menu.CheckMenuItem(M_SHOW_MAIL_HEADER_Id, MF_CHECKED | MF_BYCOMMAND);
	else
		retState = m_menu.CheckMenuItem(M_SHOW_MAIL_HEADER_Id, MF_UNCHECKED | MF_BYCOMMAND);

	DWORD color = CMainFrame::m_ColorStylesDB.m_colorStyles.GetColor(ColorStyleConfig::MailMessage);
	if (color != COLOR_WHITE)
		retState = m_menu.CheckMenuItem(M_ENABLE_DISABLE_COLOR_Id, MF_CHECKED | MF_BYCOMMAND);
	else
		retState = m_menu.CheckMenuItem(M_ENABLE_DISABLE_COLOR_Id, MF_UNCHECKED | MF_BYCOMMAND);
	int deb = 1;
}



void NMsgView::OnRButtonDown(UINT nFlags, CPoint point)
{
	// TODO: Add your message handler code here and/or call default

	const wchar_t *ClearFindtext = L"Clear Find Text";
	const wchar_t*CustomColor = L"Custom Color Style";
	const wchar_t*ShowMailHdr = L"View Raw Message Header";
	const wchar_t*HdrPaneLayout = L"Header Pane Layout";
	const wchar_t*ShowMailHtml = L"View Message Html Block";
	const wchar_t*ShowMailHtmlAsText = L"View Message Html Block as Text";
	const wchar_t*ShowMailText = L"View Message Plain Text Block";

	NListView *pListView = 0;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		pListView = pFrame->GetListView();
	}

	CPoint pt;
	::GetCursorPos(&pt);
	CWnd *wnd = WindowFromPoint(pt);

	// To avoid multiple initializations
	// TODO: Simplify menu  item check state handling
	if (nMenuInitialed == FALSE)
	{
		//CMenu menu;
		m_menu.CreatePopupMenu();
		m_menu.AppendMenu(MF_SEPARATOR);

		//const UINT M_CLEAR_FIND_TEXT_Id = 1;
		MyAppendMenu(&m_menu, M_CLEAR_FIND_TEXT_Id, ClearFindtext);

		//const UINT M_ENABLE_DISABLE_COLOR_Id = 2;
		MyAppendMenu(&m_menu, M_ENABLE_DISABLE_COLOR_Id, CustomColor, MF_UNCHECKED);

		//const UINT M_SHOW_MAIL_HEADER_Id = 3;
		MyAppendMenu(&m_menu, M_SHOW_MAIL_HEADER_Id, ShowMailHdr, MF_UNCHECKED);

		//const UINT M_SHOW_MAIL_HTML_AS_TEXT_Id = 4;
		//const UINT M_SHOW_MAIL_TEXT_Id = 5;
		//const UINT M_SHOW_MAIL_HTML_Id = 6;
		if (MboxMail::developerMode)
		{
			MyAppendMenu(&m_menu, M_SHOW_MAIL_HTML_AS_TEXT_Id, ShowMailHtmlAsText, m_hdrWindowLen == 400);
			MyAppendMenu(&m_menu, M_SHOW_MAIL_TEXT_Id, ShowMailText, m_hdrWindowLen == 200);
			MyAppendMenu(&m_menu, M_SHOW_MAIL_HTML_Id, ShowMailHtml, m_hdrWindowLen == 300);
		}

#if 0
		// Added global option under View
		CMenu hdrLayout;
		hdrLayout.CreatePopupMenu();
		hdrLayout.AppendMenu(MF_SEPARATOR);

		const UINT S_DEFAULT_LAYOUT_Id = 4;
		MyAppendMenu(&hdrLayout, S_DEFAULT_LAYOUT_Id, L"Default", m_hdrPaneLayout == 0);

		const UINT S_EXPANDED_LAYOUT_Id = 5;
		MyAppendMenu(&hdrLayout, S_EXPANDED_LAYOUT_Id, L"Expanded", m_hdrPaneLayout == 1);

		menu.AppendMenu(MF_POPUP | MF_STRING, (UINT)hdrLayout.GetSafeHmenu(), L"Header Fields Layout");
		menu.AppendMenu(MF_SEPARATOR);
#endif

		int index1 = 0;
		ResHelper::UpdateMenuItemsInfo(&m_menu, index1);
		int index = 0;
		ResHelper::LoadMenuItemsInfo(&m_menu, index);

		nMenuInitialed = TRUE;

		SetMenuToInitialState();

		m_menu.SetMenuAsCustom(&m_menu, 0);
	}
	else
	{
		int deb = 1;
	}

	int command = m_menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, this);

	UINT n_Flags = TPM_RETURNCMD;
	CString menuString;
	int chrCnt = m_menu.GetMenuString(command, menuString, n_Flags);

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


			UINT bCheck = ToggleMenuCheckState(&m_menu, command);

			int iItem = pListView->m_lastSel;
			if (m_hdrWindowLen)
			{
				this->ShowMailHeader(iItem);
			}
			else
			{
				pListView->Invalidate();
				// Set ignoreViewMessageHeader to TRUE to prevent SelectItem() from touching m_menu
				pListView->SelectItem(pListView->m_lastSel, TRUE);
			}
		}
		int deb = 1;
	}
	else if ((command == M_SHOW_MAIL_HEADER_Id))
	{
		int iItem = pListView->m_lastSel;
		UINT bCheck = ToggleMenuCheckState(&m_menu, command);

		if (m_hdrWindowLen == 100)
		{
			this->HideMailHeader(iItem);
			pListView->Invalidate();
			// Set ignoreViewMessageHeader to TRUE to prevent SelectItem() from touching m_menu
			pListView->SelectItem(pListView->m_lastSel, TRUE);
		}
		else
		{
			this->ShowMailHeader(iItem);
		}
		int deb = 1;
	}
	else if ((command == M_SHOW_MAIL_HTML_AS_TEXT_Id))
	{
		int iItem = pListView->m_lastSel;
		if (m_hdrWindowLen == 400)
		{
			this->HideMailHeader(iItem);
			pListView->Invalidate();
			// Set ignoreViewMessageHeader to TRUE to prevent SelectItem() from touching m_menu
			pListView->SelectItem(iItem, TRUE);
		}
		else
		{
			this->ShowMailHtmlBlockAsText(iItem);
		}
		int deb = 1;
	}
	else if ((command == M_SHOW_MAIL_TEXT_Id))
	{
		int iItem = pListView->m_lastSel;
		if (m_hdrWindowLen == 200)
		{
			this->HideMailHeader(iItem);
			pListView->Invalidate();
			// Set ignoreViewMessageHeader to TRUE to prevent SelectItem() from touching m_menu
			pListView->SelectItem(iItem, TRUE);
		}
		else
		{
			int textType = 0;
			this->ShowMailTextBlock(iItem, textType);
		}
		int deb = 1;
	}
	else if ((command == M_SHOW_MAIL_HTML_Id))
	{
		int iItem = pListView->m_lastSel;
		if (m_hdrWindowLen == 300)
		{
			this->HideMailHeader(iItem);
			pListView->Invalidate();
			// Set ignoreViewMessageHeader to TRUE to prevent SelectItem() from touching m_menu
			pListView->SelectItem(iItem, TRUE);
		}
		else
		{
			int textType = 1;
			this->ShowMailTextBlock(iItem, textType);
		}
		int deb = 1;
	}

#if 0
	// Added global option under View
	else if ((command == S_DEFAULT_LAYOUT_Id))
	{
		CString section_general = CString(sz_Software_mboxview) + L"\\General";

		m_hdrPaneLayout = 0;
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_general, L"headerPaneLayout", m_hdrPaneLayout);
		Invalidate();
		UpdateLayout();
		int deb = 1;
	}
	else if ((command == S_EXPANDED_LAYOUT_Id))
	{
		CString section_general = CString(sz_Software_mboxview) + L"\\General";

		m_hdrPaneLayout = 1;
		CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_general, L"headerPaneLayout", m_hdrPaneLayout);
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
	static const int cFromMailBeginLen = istrlen(cFromMailBegin);

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
			headerLength = IntPtr2Int(p - data);
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
			headerLength = IntPtr2Int(pend - data);
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

	NListView *pListView = 0;
	NMsgView *pMsgView = 0;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		pListView = pFrame->GetListView();
	}


	BOOL ret;
	if (pListView && pListView->m_bExportEml)
	{
		// Get raw mail body
		ret = m->GetBodySS(&m_hdrDataTmp, 0);
		pListView->SaveAsEmlFile(m_hdrDataTmp.Data(), m_hdrDataTmp.Count());
	}
	else
	{
		ret = m->GetBodySS(&m_hdrDataTmp, 128 * 1024);
	}

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
		// CFile Read removes CR and therefore we need to add CR, great  !!  // FIXMEFIXME
		char *ms = strchr(p, '\n');
		if (ms) 
		{
			BOOL bAddCR = FALSE;
			if (*(ms - 1) != '\r')
				bAddCR = TRUE;

			//bAddCR = TRUE;
			int pos = IntPtr2Int(ms - p + 1);
			if (bAddCR)
			{
				SimpleString *tmpbuf = MboxMail::get_tmpbuf();
				tmpbuf->Clear();
				TextUtilsEx::ReplaceNL2CRNL((LPCSTR)m_hdrDataTmp.Data(), m_hdrDataTmp.Count(), tmpbuf);
				m_hdrData.Clear();
				m_hdrData.Append("\r\n");

				DWORD error;
				UINT CP_US_ASCII = 20127;
				// int codePage = CP_US_ASCII;
				int codePage = CP_ACP;  // default to ANSI code page. or call GetACP() ?? to retrieve the current Windows ANSI code page identifier for the operating system.
				TextUtilsEx::CodePage2WStr(tmpbuf->Data(), tmpbuf->Count(), codePage, &m_hdrData, error);

				wchar_t *data = (wchar_t*)m_hdrData.Data();

				MboxMail::rel_tmpbuf();
			}
			else
			{
				m_hdrData.Clear();
				m_hdrData.Append("\r\n");

				DWORD error;
				UINT CP_US_ASCII = 20127;
				//int codePage = CP_US_ASCII;
				int codePage = CP_ACP;
				TextUtilsEx::CodePage2WStr(m_hdrDataTmp.Data(), m_hdrDataTmp.Count(), codePage, &m_hdrData, error);

				wchar_t* data = (wchar_t*)m_hdrData.Data();
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

	//NListView *pListView = 0;
	//NMsgView *pMsgView = 0;
	//CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
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

		if (m_show_charsets)
			pane_Hight += 3 + m_otherTextHight;  // body

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
	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	m_hdrPaneLayout = 0;
	CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_general, L"headerPaneLayout", m_hdrPaneLayout);
	Invalidate();
	UpdateLayout();
	int deb = 1;
}


void NMsgView::OnMessageheaderpanelayoutExpanded()
{
	// TODO: Add your command handler code here

	CString section_general = CString(sz_Software_mboxview) + L"\\General";

	m_hdrPaneLayout = 1;
	CProfile::_WriteProfileInt(HKEY_CURRENT_USER, section_general, L"headerPaneLayout", m_hdrPaneLayout);
	Invalidate();
	UpdateLayout();
	int deb = 1;
}

int NMsgView::PreTranslateMessage(MSG* pMsg)
{
	if (CMainFrame::m_commandLineParms.m_bEmlPreviewMode)
	{
		if ((pMsg->message & 0xffff) == WM_MOUSEMOVE)
		{
			int deb = 1;
		}
		else if ((pMsg->message & 0xffff) == WM_KEYDOWN)
		{
			if (pMsg->wParam == VK_ESCAPE)
			{
				AfxGetMainWnd()->PostMessage(WM_CLOSE);            // Do not process further
			}
		}
	}
	return CWnd::PreTranslateMessage(pMsg);
}

int NMsgView::ShowMailHtmlBlockAsText(int mailPosition)
{
	ULONGLONG tc_start = GetTickCount64();

	BOOL ret = TRUE;
	if ((mailPosition >= MboxMail::s_mails.GetCount()) || (mailPosition < 0))
		return -1;

	MboxMail *m = MboxMail::s_mails[mailPosition];
	if (m == 0)
	{
		// TODO: Assert ??
		return 1;
	}

	NListView *pListView = 0;
	NMsgView *pMsgView = 0;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		pListView = pFrame->GetListView();
	}


	CFile fp;
	CFileException ExError;
	if (fp.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError) == FALSE)
	{
		DWORD lastErr = ::GetLastError();
#if 1
		HWND h = GetSafeHwnd();
		//HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		CString fmt = L"Could not open mail file:\n\n\"%s\"\n\n%s";  // new format
		CString errorText = FileUtils::ProcessCFileFailure(fmt, MboxMail::s_path, ExError, lastErr, h); 
#else
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not open \"" + MboxMail::s_path;
		txt += L"\" mail file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);
		//errorText = txt;
#endif
		ret = FALSE;

		return 0;
	}

	m_hdrDataTmp.Clear();
	UINT pageCode = 0;
	int textType = 1;
	int plainTextMode = 0;  // no extra img tags; html text has img tags already
	int textlen = MboxMail::GetMailBody_mboxview(fp, mailPosition, &m_hdrDataTmp, pageCode, textType, plainTextMode);  // returns pageCode
	if (textlen != m_hdrDataTmp.Count())
		int deb = 1;

	UINT outPageCode = pageCode;

	if (m_hdrDataTmp.Count() > 0)
	{
		char *p = m_hdrDataTmp.Data();
		// CFile Read removes CR and therfore we need to add CR, great  !!
		char *ms = strchr(p, '\n');
		if (ms)
		{
			BOOL bAddCR = FALSE;
			if (*(ms - 1) != '\r')
				bAddCR = TRUE;
			int pos = IntPtr2Int(ms - p + 1);
			if (bAddCR)
			{
				SimpleString *tmpbuf = MboxMail::get_tmpbuf();
				tmpbuf->Clear();

				TextUtilsEx::ReplaceNL2CRNL((LPCSTR)m_hdrDataTmp.Data(), m_hdrDataTmp.Count(), tmpbuf);
				m_hdrData.Clear();
				m_hdrData.Append("\r\n");

				BOOL bestEffortTextExtract = FALSE;
				if (bestEffortTextExtract)
				{
					// Relative fast but extraction needs to be improved
					HtmlUtils::ExtractTextFromHTML_BestEffort(tmpbuf, &m_hdrDataTmp, pageCode, outPageCode);
				}
				else
				{
					// Below Relies on IHTMLDocument2 // This is could be slow if hyperlinks are present
					// May need to remove hyperlinks first from the document ?
					HtmlUtils::GetTextFromIHTMLDocument(tmpbuf, &m_hdrDataTmp, pageCode, outPageCode);
				}

				DWORD error;
				UINT CP_US_ASCII = 20127;
				int codePage = CP_UTF8;
				// codePage is ignored, pageCode is used instead
				// data in m_hdrData is cleared by CodePage2WStr. Top blank line is lost. FIXME ??
				TextUtilsEx::CodePage2WStr(m_hdrDataTmp.Data(), m_hdrDataTmp.Count(), pageCode, &m_hdrData, error);

				MboxMail::rel_tmpbuf();
			}
			else
			{
				SimpleString* tmpbuf = MboxMail::get_tmpbuf();
				tmpbuf->Clear();

				m_hdrData.Clear();
				m_hdrData.Append("\r\n");

				BOOL bestEffortTextExtract = FALSE;
				if (bestEffortTextExtract)
				{
					// Relative fast but extraction needs to be improved
					HtmlUtils::ExtractTextFromHTML_BestEffort(&m_hdrDataTmp, tmpbuf, pageCode, outPageCode);
				}
				else
				{
					// Below Relies on IHTMLDocument2 // This is very very slow if hyperlinks are present
					// May need to remove hyperlinks first from the document ?
					HtmlUtils::GetTextFromIHTMLDocument(&m_hdrDataTmp, tmpbuf, pageCode, outPageCode);
				}

				DWORD error;
				UINT CP_US_ASCII = 20127;
				int codePage = CP_UTF8;
				// codePage is ignored, pageCode is used instead
				// data in m_hdrData is cleared by CodePage2WStr. Top blank line is lost. FIXME ??
				TextUtilsEx::CodePage2WStr(tmpbuf->Data(), tmpbuf->Count(), pageCode, &m_hdrData, error);

				MboxMail::rel_tmpbuf();
			}
		}
	}
	else
		m_hdrData.SetCount(0);

	fp.Close();

	ULONGLONG tc_end = GetTickCount64();
	DWORD delta = (DWORD)(tc_end - tc_start);
	TRACE(L"ShowMailHtmlBlockAsText:ExtractTextFromHTML_BestEffort extracted text in %ld milliseconds.\n", delta);

	m_hdrWindowLen = 400;

	m_hdr.ShowWindow(SW_SHOW);
	m_hdr.ShowWindow(SW_RESTORE);

	Invalidate();
	UpdateLayout();

	return 1;
}

int NMsgView::ShowMailTextBlock(int mailPosition, int textType)
{
	BOOL ret = TRUE;
	if ((mailPosition >= MboxMail::s_mails.GetCount()) || (mailPosition < 0))
		return -1;

	MboxMail *m = MboxMail::s_mails[mailPosition];
	if (m == 0)
	{
		// TODO: Assert ??
		return 1;
	}

	NListView *pListView = 0;
	NMsgView *pMsgView = 0;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (pFrame)
	{
		pListView = pFrame->GetListView();
	}

	CFile fp;
	CFileException ExError;
	if (fp.Open(MboxMail::s_path, CFile::modeRead | CFile::shareDenyWrite, &ExError) == FALSE)
	{
		DWORD lastErr = ::GetLastError();
#if 1
		HWND h = GetSafeHwnd();
		//HWND h = CmboxviewApp::GetActiveWndGetSafeHwnd();
		CString fmt = L"Could not open mail file:\n\n\"%s\"\n\n%s";  // new format
		CString errorText = FileUtils::ProcessCFileFailure(fmt, MboxMail::s_path, ExError, lastErr, h);
#else
		CString exErrorStr = FileUtils::GetFileExceptionErrorAsString(ExError);

		CString txt = L"Could not open \"" + MboxMail::s_path;
		txt += L"\" mail file.\n";
		txt += exErrorStr;

		TRACE(L"%s\n", txt);
		//errorText = txt;
#endif
		ret = FALSE;
		return 0;
	}

	m_hdrDataTmp.Clear();
	UINT pageCode = 0;
	int plainTextMode = 0;  // no extra img tags;
	int textlen = MboxMail::GetMailBody_mboxview(fp, mailPosition, &m_hdrDataTmp, pageCode, textType, plainTextMode);  // returns pageCode
	if (textlen != m_hdrDataTmp.Count())
		int deb = 1;

	if (m_hdrDataTmp.Count() > 0)
	{
		char *p = m_hdrDataTmp.Data();
		// CFile Read removes CR and therfore we need to add CR, great  !!
		char *ms = strchr(p, '\n');
		if (ms)
		{
			BOOL bAddCR = FALSE;
			if (*(ms - 1) != '\r')
				bAddCR = TRUE;
			int pos = IntPtr2Int(ms - p + 1);
			if (bAddCR)
			{
				SimpleString *tmpbuf = MboxMail::get_tmpbuf();
				TextUtilsEx::ReplaceNL2CRNL((LPCSTR)m_hdrDataTmp.Data(), m_hdrDataTmp.Count(), tmpbuf);
				m_hdrData.Clear();
				m_hdrData.Append("\r\n");

				DWORD error;
				UINT CP_US_ASCII = 20127;
				int codePage = CP_UTF8;
				// codePage is ignored, pageCode is used instead
				// data in m_hdrData is cleared by CodePage2WStr. Top blank line is lost. FIXME ??
				TextUtilsEx::CodePage2WStr(m_hdrDataTmp.Data(), m_hdrDataTmp.Count(), pageCode, &m_hdrData, error);
				MboxMail::rel_tmpbuf();
			}
			else
			{
				m_hdrData.Clear();
				m_hdrData.Append("\r\n");

				DWORD error;
				UINT CP_US_ASCII = 20127;
				int codePage = CP_UTF8;
				// codePage is ignored, pageCode is used instead
				// data in m_hdrData is cleared by CodePage2WStr. Top blank line is lost. FIXME ??
				TextUtilsEx::CodePage2WStr(m_hdrDataTmp.Data(), m_hdrDataTmp.Count(), pageCode, &m_hdrData, error);
			}
		}
	}
	else
		m_hdrData.SetCount(0);

	fp.Close();

	if (textType == 0)
		m_hdrWindowLen = 200;
	else
		m_hdrWindowLen = 300;

	m_hdr.ShowWindow(SW_SHOW);
	m_hdr.ShowWindow(SW_RESTORE);

	Invalidate();
	UpdateLayout();

	return 1;
}

LRESULT NMsgView::OnCmdParam_OnSizeMsgView(WPARAM wParam, LPARAM lParam)
{
	this->UpdateLayout();
	return 0;
}
