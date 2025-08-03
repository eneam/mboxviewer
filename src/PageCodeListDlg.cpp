// PageCodeListDlg.cpp : implementation file
//

#include "stdafx.h"
#include "afxdialogex.h"
#include "TextUtilsEx.h"
#include "PageCodeListDlg.h"
#include "ResHelper.h"
#include "MainFrm.h"


// PageCodeListDlg dialog

//#define OWNERDATA  // to activate LVN_GETDISPINFO, must set Owner Data option

PageCodeListDlg* PageCodeListDlg::pPageCodeListDlgInstance = 0;

IMPLEMENT_DYNAMIC(PageCodeListDlg, CDialogEx)

PageCodeListDlg::PageCodeListDlg(CWnd* pParent /*=nullptr*/)
//DIALOG_FROM_TEMPLATE( : CDialogEx(IDD_PAGE_CODE_LIST_DLG, pParent))
	: CDialogEx(IDD_PAGE_CODE_LIST_DLG, pParent)
{
	pPageCodeListDlgInstance = this;
	m_descendingSort = TRUE;
	m_CodePageSelected = GetACP();
	m_pParent = pParent;
}

PageCodeListDlg::~PageCodeListDlg()
{
	pPageCodeListDlgInstance = 0;
}

INT_PTR PageCodeListDlg::DoModal()
{
#ifdef _DIALOG_FROM_TEMPLATE
	INT_PTR ret = CDialogEx::DoModal();
	//INT_PTR ret = CMainFrame::SetTemplate(this, IDD_PAGE_CODE_LIST_DLG, m_pParent);
#else
	INT_PTR ret = CDialogEx::DoModal();
#endif
	return ret;
}

void PageCodeListDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_CODE_PAGE_ENUMERATION, m_list);
}

BEGIN_MESSAGE_MAP(PageCodeListDlg, CDialogEx)
	ON_NOTIFY(LVN_ITEMACTIVATE, IDC_CODE_PAGE_ENUMERATION, &PageCodeListDlg::OnItemactivateCodePageEnumeration)
	ON_NOTIFY(LVN_GETDISPINFO, IDC_CODE_PAGE_ENUMERATION, &PageCodeListDlg::OnGetdispinfoCodePageEnumeration)
	ON_NOTIFY(HDN_ITEMCLICK, 0, &PageCodeListDlg::OnItemclickCodePageEnumeration)
	ON_WM_CREATE()
	ON_WM_PAINT()
	ON_NOTIFY(NM_CLICK, IDC_CODE_PAGE_ENUMERATION, &PageCodeListDlg::OnClickCodePageEnumeration)
	ON_WM_TIMER()
END_MESSAGE_MAP()


// PageCodeListDlg message handlers

void PageCodeListDlg::OnItemactivateCodePageEnumeration(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMITEMACTIVATE pNMIA = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: Add your control notification handler code here

	*pResult = 0;
}

// Must set OWNERDATA for the list crl
void PageCodeListDlg::OnGetdispinfoCodePageEnumeration(NMHDR* pNMHDR, LRESULT* pResult)
{
	NMLVDISPINFO* pDispInfo = reinterpret_cast<NMLVDISPINFO*>(pNMHDR);

	// TODO: Add your control notification handler code here

	if (pDispInfo->item.mask & LVIF_TEXT)
	{
		CodePageInfo *info = &m_codePageList[pDispInfo->item.iItem];

		switch (pDispInfo->item.iSubItem) {
		case 0:
			m_tmpstr.Format(L"%d", info->codePage);
			break;
		case 1:
			m_tmpstr = info->codePageName;
			break;
		case 2:
			m_tmpstr.Format(L"%d", info->maxCharSize);
			break;

		default:
			break;
		}
		pDispInfo->item.pszText = (wchar_t*)((LPCWSTR)m_tmpstr);
	}
	*pResult = 0;
}

void PageCodeListDlg::OnItemclickCodePageEnumeration(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMHEADER phdr = reinterpret_cast<LPNMHEADER>(pNMHDR);
	// TODO: Add your control notification handler code here

	if (phdr->iItem != 0)
		return;

	m_descendingSort = !m_descendingSort;
	SortCodePageList(phdr->iItem, m_descendingSort);

#ifndef OWNERDATA
	int nItem;
	BOOL t1, t2, t3;
	int codePageCount = (int)m_codePageList.GetCount();

	int ii = 0;
	for (ii = 0; ii < codePageCount; ii++)
	{
		CodePageInfo* info = &m_codePageList.GetData()[ii];

		//errno_t _itoa_s(int value, char* buffer, size_t size, int radix);

		CString codePageStr;
		codePageStr.Format(L"%d", info->codePage);
		CString maxCharSizeStr;
		maxCharSizeStr.Format(L"%d", info->maxCharSize);

		//nItem = m_list.InsertItem(ii, (LPCSTR)codePageStr);
		nItem = ii;

		t1 = m_list.SetItemText(nItem, 0, (LPCWSTR)codePageStr);
		t2 = m_list.SetItemText(nItem, 1, (LPCWSTR)info->codePageName);
		t3 = m_list.SetItemText(nItem, 2, (LPCWSTR)maxCharSizeStr);
	}

	int pos = CodePage2Index(m_CodePageSelected);

	m_list.SetItemState(-1, 0, LVIS_SELECTED);
	m_list.EnsureVisible(pos, FALSE);
	m_list.SetFocus();
	m_list.SetItemState(pos, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);

	//m_list.RedrawItems(0, m_list.GetItemCount());

	m_list.Invalidate();
	SetRedraw();
#else
	UpdateData(TRUE);
	m_list.SetItemCount(m_codePageList.GetCount());
	m_list.Invalidate();
#endif
	*pResult = 0;
}

BOOL PageCodeListDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here

		// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	DWORD dwStyle = m_list.GetStyle();
	DWORD dwExStyle = m_list.GetExtendedStyle();
	dwExStyle |= LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES;
	m_list.SetExtendedStyle(dwExStyle);

	//m_list.SetFont(GetFont());

	PageCodeListDlg::ResetFont();

	CString CodePageId;
	CodePageId.LoadString(IDS_PAGE_CODE_LIST_CODEPAGE_ID);
	CString CodePageName;
	CodePageName.LoadString(IDS_PAGE_CODE_LIST_CODEPAGE_NAME);
	CString MaxCharSize;
	MaxCharSize.LoadString(IDS_PAGE_CODE_LIST_MAX_CHAR_SIZE);

	int r1 = m_list.InsertColumn(0, CodePageId, LVCFMT_LEFT, 100, 0);
	int r2 = m_list.InsertColumn(1, CodePageName, LVCFMT_LEFT, 320);
	int r3 = m_list.InsertColumn(2, MaxCharSize, LVCFMT_LEFT, 120, 0);

	//int r1 = m_list.InsertColumn(0, L"CodePage Id", LVCFMT_LEFT, 100, 0);
	//int r2 = m_list.InsertColumn(1, L"Code Page Name", LVCFMT_LEFT, 320);
	//int r3 = m_list.InsertColumn(2, L"Max Character Size", LVCFMT_LEFT, 120,0);

	BOOL retval = EnumSystemCodePages(EnumCodePagesProc, CP_INSTALLED);

	int column = 0;
	SortCodePageList(column, m_descendingSort);

#ifndef OWNERDATA
	int nItem;
	BOOL t1, t2, t3;
	int codePageCount = (int)m_codePageList.GetCount();

	int ii = 0;
	for (ii = 0; ii < codePageCount; ii++)
	{
		CodePageInfo* info = &m_codePageList.GetData()[ii];

		//errno_t _itoa_s(int value, char* buffer, size_t size, int radix);

		CString codePageStr;
		codePageStr.Format(L"%d", info->codePage);
		CString maxCharSizeStr;
		maxCharSizeStr.Format(L"%d", info->maxCharSize);

		nItem = m_list.InsertItem(ii, (LPCWSTR)codePageStr);
		t1 = m_list.SetItemText(nItem, 0, (LPCWSTR)codePageStr);
		t2 = m_list.SetItemText(nItem, 1, (LPCWSTR)info->codePageName);
		t3 = m_list.SetItemText(nItem, 2, (LPCWSTR)maxCharSizeStr);

		//UINT nId = m_list.MapIndexToID((UINT)nItem);
	}
#endif

	// Resize window. Enhance if needed to be accurate
	int cx = m_list.GetColumnWidth(0);
	cx += m_list.GetColumnWidth(1);
	cx += m_list.GetColumnWidth(2);
	cx += 100;

	int cy = 500;
	SetWindowPos(0, 0, 0, cx, cy, SWP_NOMOVE | SWP_NOZORDER);

	m_list.SetItemCount((int)m_codePageList.GetCount());

	m_list.SetItemState(-1, 0, LVIS_SELECTED);
	int pos = m_CodePageSelected;

	m_list.SetItemState(pos, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
	m_list.EnsureVisible(pos, FALSE);
	m_list.SetFocus();
	
	m_list.Invalidate();
	//m_list.RedrawWindow();
	m_list.SetRedraw();

	UINT_PTR m_nIDEvent = 1;
	UINT m_nElapse = 20;
	this->SetTimer(m_nIDEvent, m_nElapse, NULL);

	NMHEADER nmHdr;
	nmHdr.iItem = 0;
	//m_list.PostMessage(HDN_ITEMCLICK, 0, 0);

	ResHelper::LoadDialogItemsInfo(this);
	ResHelper::UpdateDialogItemsInfo(this);

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

BOOL PageCodeListDlg::EnumCodePagesProc(LPWSTR  pCodePage)
{
	UINT codePage = _ttoi(pCodePage);

	if (pPageCodeListDlgInstance)
		pPageCodeListDlgInstance->AddCodePage(codePage);

	return TRUE;
}

void PageCodeListDlg::AddCodePage(UINT codePage)
{
#if 1
	CodePageInfo cpInfo;
	cpInfo.codePage = codePage;
	cpInfo.maxCharSize = 0;

	CPINFOEX CPInfoEx;
	DWORD dwFlags = 0;

	BOOL ret2 = GetCPInfoEx(codePage, dwFlags, &CPInfoEx);
	//IMultiLanguage::GetCharsetInfo 

	BOOL ret1 = TextUtilsEx::Id2LongInfo(codePage, cpInfo.codePageName);

	int namelen = cpInfo.codePageName.GetLength();

	if (ret1 == TRUE)
	{
		if (ret2 == TRUE)
			cpInfo.maxCharSize = CPInfoEx.MaxCharSize;
	}
	else if (ret2 == TRUE)
	{
		cpInfo.codePageName.Append(CPInfoEx.CodePageName);
		cpInfo.maxCharSize = CPInfoEx.MaxCharSize;
	}
	else
		return;
#else
	CPINFOEX CPInfoEx;
	DWORD dwFlags = 0;

	BOOL ret = GetCPInfoEx(codePage, dwFlags, &CPInfoEx);
	if (ret == FALSE)
		return;

	CodePageInfo cpInfo;
	cpInfo.codePage = codePage;
	cpInfo.codePageName.Append(CPInfoEx.CodePageName);
	cpInfo.maxCharSize = CPInfoEx.MaxCharSize;
#endif
	m_codePageList.Add(cpInfo);


#if 0
	UINT myCodePage = GetACP();

	int ii = 0;
	for (ii = 0; ii < m_codePageList.GetCount(); ii++)
	{
		CodePageInfo* info = &m_codePageList.GetAt(ii);
		if (myCodePage == info->codePage)
		{
		}
	}
#endif
}

void PageCodeListDlg::OnOK()
{
	// TODO: Add your specialized code here and/or call the base class

	CDialogEx::OnOK();
}

void PageCodeListDlg::OnCancel()
{
	// TODO: Add your specialized code here and/or call the base class

	CDialogEx::OnCancel();
}

int PageCodeListDlg::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CDialogEx::OnCreate(lpCreateStruct) == -1)
		return -1;

	// TODO:  Add your specialized creation code here

	return 0;
}

void PageCodeListDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM)dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

void PageCodeListDlg::PrintListCtrl(HTREEITEM hItem, BOOL recursive, int depth)
{
#if 0
	CString itemName;
#endif
}

void PageCodeListDlg::OnClickCodePageEnumeration(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>(pNMHDR);
	// TODO: Add your control notification handler code here

	CString strNewState;
	UINT uNewState = pNMItemActivate->uNewState;
	ItemState2Str(uNewState, strNewState);

	CString strOldState;
	UINT uOldState = pNMItemActivate->uOldState;
	ItemState2Str(uOldState, strOldState);

	int iItem = pNMItemActivate->iItem;

	UINT uMask = 0xffffffff;
	UINT uItemState = m_list.GetItemState(iItem, uMask);
	CString strItemState;
	ItemState2Str(uItemState, strItemState);

	*pResult = 0;
}

#include <algorithm>

bool FooPredAscending(const CodePageInfo& first, const CodePageInfo& second)
{
	if (first.codePage > second.codePage)
		return true;
	return false;
}

bool FooPredDescending(const CodePageInfo& first, const CodePageInfo& second)
{
	if (first.codePage < second.codePage)
		return true;
	return false;
}

void PageCodeListDlg::SortCodePageList(int column, BOOL descendingSort)
{
	if (descendingSort)
		std::sort(m_codePageList.GetData(), m_codePageList.GetData() + m_codePageList.GetSize(), FooPredDescending);
	else
		std::sort(m_codePageList.GetData(), m_codePageList.GetData() + m_codePageList.GetSize(), FooPredAscending);
}

void PageCodeListDlg::ItemState2Str(UINT uState, CString& strState)
{
	strState.Empty();

	if (uState & LVIS_ACTIVATING) {
		strState.Append(L"LVIS_ACTIVATING");
	}
	if (uState & LVIS_CUT) {
		if (!strState.IsEmpty()) { strState.Append(L" | "); }
		strState.Append(L"LVIS_CUT");
	}
	if (uState & LVIS_DROPHILITED) {
		if (!strState.IsEmpty()) { strState.Append(L" | "); }
		strState.Append(L"LVIS_DROPHILITED");
	}
	if (uState & LVIS_FOCUSED) {
		if (!strState.IsEmpty()) { strState.Append(L" | "); }
		strState.Append(L"LVIS_FOCUSED");
	}
	if (uState & LVIS_OVERLAYMASK) {
		if (!strState.IsEmpty()) { strState.Append(L" | "); }
		strState.Append(L"LVIS_OVERLAYMASK");
	}
	if (uState & LVIS_SELECTED) {
		if (!strState.IsEmpty()) { strState.Append(L" | "); }
		strState.Append(L"LVIS_SELECTED");
	}
	if (uState & LVIS_STATEIMAGEMASK) {
		if (!strState.IsEmpty()) { strState.Append(L" | "); }
		strState.Append(L"LVIS_STATEIMAGEMASK");
	}
}

int PageCodeListDlg::CodePage2Index(UINT codePage)
{
	int codePageCount = (int)m_codePageList.GetCount();

	int ii = 0;
	for (ii = 0; ii < codePageCount; ii++)
	{
		CodePageInfo* info = &m_codePageList.GetData()[ii];
		if (info->codePage == codePage)
			return ii;
	}
	return -1;
}


BOOL PageCodeListDlg::PreTranslateMessage(MSG* pMsg)
{
	// TODO: Add your specialized code here and/or call the base class

	if ((pMsg->message & 0xffff) == WM_DRAWITEM)
	{
		int deb = 1;
	}

	return CDialogEx::PreTranslateMessage(pMsg);
}


BOOL PageCodeListDlg::OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult)
{
	// TODO: Add your specialized code here and/or call the base class

	return CDialogEx::OnNotify(wParam, lParam, pResult);
}


void PageCodeListDlg::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: Add your message handler code here and/or call default

	m_list.SetItemCount((int)m_codePageList.GetCount());

	m_list.SetItemState(-1, 0, LVIS_SELECTED);
	int pos = m_CodePageSelected;

	m_list.SetItemState(pos, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
	m_list.EnsureVisible(pos, FALSE);
	m_list.SetFocus();

	m_list.Invalidate();
	m_list.SetRedraw();
	SetRedraw();

	KillTimer(nIDEvent);
	//CDialogEx::OnTimer(nIDEvent);
}



void PageCodeListDlg::ResetFont()
{
	static int g_pointSize = CMainFrame::m_dfltPointFontSize;
	static CString g_fontName = CMainFrame::m_dfltFontName;

	if (CMainFrame::m_cnfFontSize != CMainFrame::m_dfltFontSize)
	{
		g_pointSize = CMainFrame::m_cnfFontSize * 10;
	}

	m_font.DeleteObject();
	if (!m_font.CreatePointFont(g_pointSize, g_fontName))
		m_font.CreatePointFont(g_pointSize, L"Arial");

	LOGFONT	lf;
	m_font.GetLogFont(&lf);
	m_list.SetFont(&m_font);
}
