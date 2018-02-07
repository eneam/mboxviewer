// NListView.cpp : implementation file
//

#include "stdafx.h"
#include "mboxview.h"
#include "NListView.h"
#pragma warning (disable : 4786)
#include "Mime.h"
#include "MimeCode.h"
#include "MboxMail.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


BOOL PathFileExist(LPCSTR path)
{
	DWORD fa = GetFileAttributes(path);

	if( fa != (DWORD)0xFFFFFFFF)
		return TRUE;
	return FALSE;
}

CString GetmboxviewTempPath(void)
{
	char	buf[_MAX_PATH];
	GetTempPath(_MAX_PATH, buf);
	if( ! PathFileExist(buf) )
		strcpy(buf,"\\");
	strcat(buf, "mboxview\\");
	CreateDirectory(buf,NULL);
	return buf;
}

BOOL RemoveDir(CString & dir)
{
	WIN32_FIND_DATA FileData;
	HANDLE hSearch;
	CString	appPath;
	BOOL	bFinished = FALSE;
	// Start searching for all files in the current directory.
	CString searchPath = dir + "\\*.*";
	hSearch = FindFirstFile(searchPath, &FileData); 
	if (hSearch == INVALID_HANDLE_VALUE) {
		TRACE("No files found.");
		return FALSE;
	}
	while ( ! bFinished ) {
		if( ! ( strcmp(FileData.cFileName, ".") == 0 || strcmp(FileData.cFileName, "..") == 0 ) ) {
			CString	fileFound = dir + "\\" + FileData.cFileName;
			if( FileData.dwFileAttributes != FILE_ATTRIBUTE_DIRECTORY )
				DeleteFile( (LPCSTR) fileFound );
			else
				RemoveDir( dir + "\\" + CString(FileData.cFileName) );
		}
		if( ! FindNextFile( hSearch, &FileData ) )
			bFinished = TRUE;
	}
	FindClose(hSearch);
	//RemoveDirectory( dir );
	return TRUE;
}

CString CreateTempFileName(CString ext = "htm")
{
	CString fmt = GetmboxviewTempPath() + "PTT%05d."+ext;
	CString fileName;
ripeti:
	fileName.Format(fmt, abs((int)(1+(int) (100000.0*rand()/(RAND_MAX+1.0)))));
	if( PathFileExist(fileName) )
		goto ripeti;

	return fileName; //+_T(".HTM");
}

IMPLEMENT_DYNCREATE(NListView, CWnd)

/////////////////////////////////////////////////////////////////////////////
// NListView

CString GetDateFormat(int i);

NListView::NListView() : m_lastStartDate((time_t)-1), m_lastEndDate((time_t)-1)
{
	m_searchPos = 0;
	m_searchString.Empty();
	m_bCaseSens = TRUE;
	m_bWholeWord = FALSE;
	m_lastScope = 0;
	m_lastSel = -1;
	m_bInFind = FALSE;
	int iFormat = CProfile::_GetProfileInt(HKEY_CURRENT_USER, sz_Software_mboxview, "format");
	m_format = GetDateFormat(iFormat);

}

NListView::~NListView()
{
}

BEGIN_MESSAGE_MAP(NListView, CWnd)
	//{{AFX_MSG_MAP(NListView)
	ON_WM_CREATE()
	ON_WM_SIZE()
	ON_WM_MOUSEWHEEL()
	//}}AFX_MSG_MAP
	ON_NOTIFY(LVN_ITEMACTIVATE, IDC_LIST, OnActivating)
	ON_NOTIFY(NM_CLICK, IDC_LIST, OnActivating)
	ON_NOTIFY(NM_DBLCLK, IDC_LIST, OnDoubleClick)
	ON_NOTIFY(LVN_COLUMNCLICK, IDC_LIST, OnColumnClick)
	ON_NOTIFY(LVN_KEYDOWN, IDC_LIST, OnKeydown)
	ON_NOTIFY(NM_CUSTOMDRAW, IDC_LIST, OnCustomDraw)
	ON_NOTIFY(LVN_GETDISPINFO, IDC_LIST, OnGetDispInfo)
	ON_UPDATE_COMMAND_UI(ID_EDIT_FIND, OnUpdateEditFind)
	ON_COMMAND(ID_EDIT_FIND, OnEditFind)
	ON_UPDATE_COMMAND_UI(ID_EDIT_FINDAGAIN, OnUpdateEditFindAgain)
	ON_COMMAND(ID_EDIT_FINDAGAIN, OnEditFindAgain)
	ON_COMMAND(ID_EDIT_VIEWEML, &NListView::OnEditVieweml)
	ON_UPDATE_COMMAND_UI(ID_EDIT_VIEWEML, &NListView::OnUpdateEditVieweml)
END_MESSAGE_MAP()

int g_pointSize = 85;
CString g_fontName = "Tahoma";

void NListView::ResetFont()
{
	m_boldFont.DeleteObject();
	m_font.DeleteObject();
	if( ! m_font.CreatePointFont (g_pointSize, g_fontName) )
		m_font.CreatePointFont (85, _T("Arial"));
	LOGFONT	lf;
	m_font.GetLogFont(&lf);
	lf.lfWeight = FW_BOLD;
	m_boldFont.CreateFontIndirect(&lf);
    m_list.SetFont(&m_font);
}

void NListView::OnDoubleClick(NMHDR* pNMHDR, LRESULT* pResult) 
{
	NMITEMACTIVATE *pnm = (NMITEMACTIVATE *)pNMHDR;
	if( m_lastSel == pnm->iItem ) {
		CString path = GetmboxviewTempPath();

		if( PathFileExist(path) )
			ShellExecute(NULL, _T("open"), path, NULL,NULL, SW_SHOWNORMAL );
	}
	*pResult = 0;
}


void NListView::OnActivating(NMHDR* pNMHDR, LRESULT* pResult) 
{
	NMITEMACTIVATE *pnm = (NMITEMACTIVATE *)pNMHDR;
	if( m_lastSel != pnm->iItem ) {
		TRACE("Selecting %d\n", pnm->iItem);
		SelectItem(pnm->iItem);
	}

	*pResult = 0;
}

void NListView::OnColumnClick(NMHDR* pNMHDR, LRESULT* pResult)
{
	NMLISTVIEW * pLV = (NMLISTVIEW*)pNMHDR;
	bool mustSort = false;
	*pResult = 0;

	// to follow OnEditFindAgain approach
	if (m_bInFind)  
		return;

	m_bInFind = true;

	switch (pLV->iSubItem) {
	case 1: // date
		if (abs(MboxMail::b_mails_which_sorted) == 1) {
			MboxMail::b_mails_which_sorted = -MboxMail::b_mails_which_sorted;
		} else {
			MboxMail::b_mails_which_sorted = 1;
		}
		MboxMail::SortByDate(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
		break;
	case 2: // from
		if (abs(MboxMail::b_mails_which_sorted) == 2) {
			MboxMail::b_mails_which_sorted = -MboxMail::b_mails_which_sorted;
		} else {
			MboxMail::b_mails_which_sorted = 2;
		}
		MboxMail::SortByFrom(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
		break;
	case 3: // to
		if (abs(MboxMail::b_mails_which_sorted) == 3) {
			MboxMail::b_mails_which_sorted = -MboxMail::b_mails_which_sorted;
		}
		else {
			MboxMail::b_mails_which_sorted = 3;
		}
		MboxMail::SortByTo(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
		break;
	case 4: // subj
		if (abs(MboxMail::b_mails_which_sorted) == 4) {
			MboxMail::b_mails_which_sorted = -MboxMail::b_mails_which_sorted;
		}
		else {
			MboxMail::b_mails_which_sorted = 4;
		}
		MboxMail::SortBySubject(0, MboxMail::b_mails_which_sorted < 0);
		mustSort = true;
		break;
	}
	if (mustSort)
	{
		MboxMail::b_mails_sorted = true;
		m_searchString.Empty();
		m_lastScope = 0;
		m_lastStartDate = 0;

		RedrawMails();
		// MessageBeep(MB_OK); // too much ??
	}
	m_bInFind = false;
}

void NListView::OnKeydown(NMHDR* pNMHDR, LRESULT* pResult) 
{
//	m_headerCtrl.PostMessage(WM_CANCELMODE);
	NMLVKEYDOWN* pLVKeyDow = (NMLVKEYDOWN*)pNMHDR;
	*pResult = 0;
	int whichSel;
	switch( pLVKeyDow->wVKey ) {
	case VK_HOME:
		SelectItem(0);
		break;
	case VK_END:
		SelectItem(m_list.GetItemCount()-1);
		break;
	case VK_UP:
		if( m_lastSel )
			SelectItem(m_lastSel-1);
		break;
	case VK_DOWN:
		if( m_lastSel < m_list.GetItemCount()-1 )
			SelectItem(m_lastSel+1);
		break;
	case VK_NEXT:
		{
			int curBottom = m_list.GetTopIndex() + m_list.GetCountPerPage() - 1;
			if( m_lastSel >= curBottom )
				whichSel = m_lastSel + m_list.GetCountPerPage() - 1;
			else
				whichSel = curBottom;
		}
		if( whichSel >= m_list.GetItemCount() )
			whichSel = m_list.GetItemCount()-1;
		SelectItem(whichSel);
		break;
	case VK_PRIOR:
		{
			int curTop = m_list.GetTopIndex();
			if( m_lastSel <= curTop ) {
				if( m_lastSel > m_list.GetCountPerPage() )
					whichSel = m_lastSel - m_list.GetCountPerPage();
				else
					whichSel = 0;
			} else
				whichSel = curTop;
		}
		SelectItem(whichSel);
		break;
/*	case VK_BACK:
	case VK_DELETE:
		if( theApp.ShowHidingDlg(GetConfirmDelete(m_listCtrl.GetSelectedCount()), g_hideDeletePopup, "hideDeletePopup", TRUE) ) {
			CWaitCursor cursor;
			DeleteSelectedArticle();
		}
		break;
	*/
	}
}

/////////////////////////////////////////////////////////////////////////////
// NListView message handlers

int NListView::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CWnd::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	if( !m_list.Create(WS_CHILD|WS_VISIBLE|LVS_SINGLESEL|LVS_REPORT|LVS_SHOWSELALWAYS|LVS_OWNERDATA, CRect(), this, IDC_LIST) )
		return -1;

	m_list.SetExtendedStyle(LVS_EX_FULLROWSELECT);	m_list.SendMessage((CCM_FIRST + 0x7), 5, 0);
	m_list.SetTextColor (::GetSysColor(COLOR_WINDOWTEXT));
	ResetFont();
	
	m_list.InsertColumn(0, "!", LVCFMT_LEFT, 18, 0);
	m_list.InsertColumn(1, "date", LVCFMT_LEFT, 100, 0);
	m_list.InsertColumn(2, "from", LVCFMT_LEFT, 150, 0);
	m_list.InsertColumn(3, "to", LVCFMT_LEFT, 150, 0);
	m_list.InsertColumn(4, "subject", LVCFMT_LEFT, 400, 0);

	return 0;
}

void NListView::OnSize(UINT nType, int cx, int cy) 
{
	CWnd::OnSize(nType, cx, cy);
	
	m_list.MoveWindow(0, 0, cx, cy);
}


void NListView::ResetSize()
{
	MboxMail::s_fSize = 0;
	m_list.SetItemCount(0);
	m_list.Invalidate();
}

void NListView::OnGetDispInfo(NMHDR* pNMHDR, LRESULT* pResult) 
{
	LV_DISPINFO* pDispInfo = (LV_DISPINFO*)pNMHDR;
	static const char * ast = "*";
	static const char * nul = "";
	static char datebuff[32];
	static char subjectbuff[1024];
	if (pDispInfo->item.mask & LVIF_TEXT)
	{
		MboxMail *m = MboxMail::s_mails[pDispInfo->item.iItem];
//		MboxMail *m = (MboxMail*)pDispInfo->item.lParam;
		const char *t;
		switch(pDispInfo->item.iSubItem) {
		case 0: // date
			/*
			if( m->m_recv == FALSE )
				t = ast;
			else
				t = nul;*/
			if( m->m_hasAttachments )
				t = ast;
			else
				t = nul;
			break;
		case 1: // date
			datebuff[0] = 0;
			if (m->m_timeDate > 0) {
				CTime tt(m->m_timeDate);
				strcpy(datebuff, (LPCSTR)tt.Format(m_format));
			}
			t = datebuff;
			break;
		case 2: // from
			t = m->m_from.GetBufferSetLength(m->m_from.GetLength());
			break;
		case 3: // to
			t = m->m_to.GetBufferSetLength(m->m_to.GetLength());
			break;
		case 4: // subj
			t = m->m_subj.GetBufferSetLength(m->m_subj.GetLength());
			break;
		default:
			t = nul;
			break;
		}
		pDispInfo->item.pszText = (char*)t;
	}
/*	if( pDispInfo->item.mask & LVIF_PARAM ) {
		pDispInfo->item.lParam = (LPARAM)m;
	}*/
	*pResult = 0;
}


void NListView::OnCustomDraw(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMLVCUSTOMDRAW lplvcd = (LPNMLVCUSTOMDRAW)pNMHDR;
	MboxMail *m;

	switch (lplvcd->nmcd.dwDrawStage)
	{
	case CDDS_PREPAINT:	// Request prepaint notifications for each item.
		*pResult = CDRF_NOTIFYITEMDRAW;
		break;
	case CDDS_ITEMPREPAINT: //Before an item is drawn
		m = MboxMail::s_mails[lplvcd->nmcd.dwItemSpec];
		// Add code here to customise whole line and return:
		if( m->m_startOff >= MboxMail::s_oSize ) {
			CDC dc;
			if( dc.Attach(lplvcd->nmcd.hdc) ) {
				dc.SelectObject(&m_boldFont);
				dc.Detach();
				*pResult = CDRF_NEWFONT;
			}
//			lplvcd->clrText = RGB(0, 0, 255);
		}
		//else lplvcd->clrText = ::GetSysColor(COLOR_WINDOWTEXT);
		//*pResult = CDRF_NEWFONT;
		break;
	default:
		*pResult = CDRF_DODEFAULT;
	}
}

void NListView::PostNcDestroy() 
{
	m_font.DeleteObject();
	m_boldFont.DeleteObject();

	RemoveDir(GetmboxviewTempPath());
	MboxMail::Destroy();
	DestroyWindow();
	delete this;
}

int fileExists(LPCSTR file)
{
   WIN32_FIND_DATA FindFileData;
   HANDLE handle = FindFirstFile(file, &FindFileData) ;
   int found = handle != INVALID_HANDLE_VALUE;
   if(found) 
   {
       //FindClose(&handle); this will crash
       FindClose(handle);
   }
   return found;
}

bool ALongRightProcessProc(const CUPDUPDATA* pCUPDUPData)
{
	CString path = (LPCSTR)pCUPDUPData->GetAppData();
	MboxMail::pCUPDUPData = pCUPDUPData;
	MboxMail::Parse(path);
	return true;
}

#define CACHE_VERSION	6

BOOL SaveMails(LPCSTR cache) {
	int ni = MboxMail::s_mails.GetSize();
	SerializerHelper sz(cache);
	if( ! sz.open(TRUE) )
		return FALSE;
	sz.writeInt(CACHE_VERSION);			// version
	sz.writeInt64(MboxMail::s_fSize = FileSize(MboxMail::s_path));	// file size
	sz.writeInt(ni);					// number of elements
	_int64 lastoff = 0;
	for (int i = 0; i < ni; i++) {
		MboxMail *m = MboxMail::s_mails[i];
		sz.writeInt64(m->m_startOff);
		sz.writeInt(m->m_length);
		sz.writeInt(m->m_hasAttachments);
		sz.writeInt(m->m_headLength);
		sz.writeInt(m->m_recv);
		sz.writeString(m->m_from);
		sz.writeString(m->m_to);
		sz.writeString(m->m_subj);
		sz.writeInt64(m->m_timeDate);
		if (lastoff < m->m_startOff)
			lastoff = m->m_startOff;
	}
	TRACE("lastoff=%lld\n", lastoff);
	sz.close();
	return TRUE;
}

int LoadMails(LPCSTR cache)
{
	SerializerHelper sz(cache);
	if( ! sz.open(FALSE) )
		return -1;
	int version;
	_int64 fSize;
	int ni = 0;
	if( ! sz.readInt(&version) )
		return -1;
	if( version != CACHE_VERSION )
		return -1;
	if( ! sz.readInt64(&fSize) )
		return -1;
	if( fSize != FileSize(MboxMail::s_path) )
		return -1;
	_int64 lastoff = 0;
	if( sz.readInt(&ni) ) {
		for( int i = 0; i < ni; i++ ) {
			MboxMail *m = new MboxMail();
			if( ! sz.readInt64(&m->m_startOff) )
				break;
			if( ! sz.readInt(&m->m_length) )
				break;
			if( ! sz.readInt(&m->m_hasAttachments) )
				break;
			if( ! sz.readInt(&m->m_headLength) )
				break;
			if( ! sz.readInt(&m->m_recv) )
				break;
			if( ! sz.readString(m->m_from) )
				break;
			if( ! sz.readString(m->m_to) )
				break;
			if( ! sz.readString(m->m_subj) )
				break;
			if (!sz.readInt64(&m->m_timeDate))
				break;
			if (lastoff < m->m_startOff)
				lastoff = m->m_startOff;
			MboxMail::s_mails.Add(m);
		}
	}
	TRACE("lastoff=%lld\n", lastoff);
	sz.close();
	MboxMail::s_fSize = MboxMail::s_oSize = fSize;

	MboxMail::s_mails_ref.SetSize(MboxMail::s_mails.GetSize());
	MboxMail::s_mails_ref.Copy(MboxMail::s_mails);
	MboxMail::b_mails_sorted = true;
	MboxMail::SortByDate();

	return ni;
}

void NListView::FillCtrl()
{
	ClearDescView();
	m_searchPos = 0;
	MboxMail::Destroy();
	m_list.SetRedraw(FALSE);
	m_list.DeleteAllItems();
	if( m_path.IsEmpty() )
		return;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if( ! pFrame )
		return;
	NTreeView *pView = pFrame->GetTreeView();
	if( ! pView )
		return;
	CString cache= m_path + ".mboxview";
	int ni = 0;
#ifdef _DEBUG
	DWORD tc = GetTickCount(); // montano@
#endif
	if( fileExists(cache) ) {
		MboxMail::s_path = m_path;
		ni = LoadMails(cache);
		if( ni < 0 ) {
			ni = 0;
			DeleteFile(cache);
		} else
			m_list.SetItemCount(ni);
	}
	if( ! fileExists(cache) ) {
		CUPDialog	Dlg(AfxGetMainWnd()->GetSafeHwnd(), ALongRightProcessProc, (LPVOID)(LPCSTR)m_path);
		if( !Dlg.DoModal() )
			return;
		if( SaveMails(cache) ) {
			ni = MboxMail::s_mails.GetSize();
			m_list.SetItemCount(ni);
		}
	}
#ifdef _DEBUG
	tc = (GetTickCount() - tc);
	TRACE("FillCtrl Took %d:%d %d\n", (tc/1000)/60, (tc/1000)%60, tc);
#endif
	m_list.EnsureVisible(ni, FALSE);
	m_list.SetRedraw(TRUE);
	//pView->m_tree.SetItemData(m_which, (DWORD)FileSize(m_path));
	pView->m_tree.SetItemState(m_which, 0, TVIS_BOLD);
	pView->SaveData();
}

void NListView::SelectItem(int iItem) 
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if( ! pFrame )
		return;
	NMsgView *pView = pFrame->GetMsgView();
	if( ! pView )
		return;
	pView->m_bAttach = FALSE;
	// Sanity check
	if( iItem < 0 || iItem >= MboxMail::s_mails.GetSize() ) {
		return;
		//ClearDescView();
		return;
	}
	m_lastSel = iItem;
	// Erase any files previously saved
	RemoveDir(GetmboxviewTempPath());
	pView->m_attachments.DeleteAllItems();
	// Get cached mail
	MboxMail *m = MboxMail::s_mails[iItem];
	// Set header data
	pView->m_strDescp1 = m->m_from;
	pView->m_strDescp2 = m->m_subj;
	CTime tt(m->m_timeDate);
	pView->m_strDescp3 = tt.Format(m_format);
	// Get raw mail body
	CString bdy = m->GetBody();

	// Save raw message
	{
	// Save mail
		CFile fp(GetmboxviewTempPath()+"message.eml", CFile::modeWrite|CFile::modeCreate);
		fp.Write(bdy, bdy.GetLength());
		fp.Close();
	}
	CString bdycharset = "UTF-8";

	// Decode MIME message
	CMimeMessage mail;
    int nLoadedSize = mail.Load(bdy, bdy.GetLength());
	// Iterate all the descendant body parts
    CMimeBody::CBodyList bodies;
    int nCount = mail.GetBodyPartList(bodies);
    CMimeBody::CBodyList::const_iterator it;
	bdy = "";
	CString ext = "";
	if( bodies.begin() == bodies.end() ) {
		//bdy = mail.GetContent();
		string strText;
		mail.GetText(strText);
		bdy = strText.c_str();
		ext = "txt";
	} else
    for (it=bodies.begin(); it!=bodies.end(); it++)
    {
        CMimeBody* pBP = *it;
		CString curExt;

        // Iterate all the header fields of this body part:
        CMimeHeader::CFieldList& fds = pBP->Fields();
        CMimeHeader::CFieldList::const_iterator itfd;
		CString charset;
        for (itfd=fds.begin(); itfd!=fds.end(); itfd++)
        {
            const CMimeField& fd = *itfd;
			const char *fname = fd.GetName();
			const char *fval = fd.GetValue();
			// Check content type to get mail extension
			if( _stricmp(fname, "Content-Type") == 0 && _strnicmp(fval, "text/", 5) == 0 ) {
				const char *p = fd.GetValue()+5;
				if( _strnicmp(p, "plain", 5) == 0 )
					curExt = "txt";
				else
				if( _strnicmp(p, "htm", 3) == 0 )
					curExt = "htm";
				else
				if( _strnicmp(p, "xml", 3) == 0 )
					curExt = "xml";
				else
					curExt = "txt";
				int pc = CString(fval).Find("charset=");
				if (pc != -1) {
					strcpy(charset.GetBufferSetLength(strlen(fval + pc + 8)), fval + pc + 8);
					charset.Trim('"');
				}
			}
        }

        if ((pBP->IsText() || pBP->IsMessage()) && (bdy.IsEmpty() || !pBP->IsAttachment()) )
        {
			// if message contains alternate parts display last one
			string strText;
			pBP->GetText(strText);
			bdy = strText.c_str();

			if (bdy.GetLength() > 0) {
				UINT CodePage = MboxMail::Str2PageCode(charset);
				if (((CodePage >= 28591) && (CodePage <= 28598)) || (CodePage == 28605)) {
					MboxMail::Str2Ansi(bdy, CodePage);
				}
			}

			ext = curExt;
			bdycharset = charset;
			TRACE("ext=%s charset=%s\n", (LPCSTR)ext, (LPCSTR)charset);
        }
        else if (pBP->IsAttachment())
        {
			// Save all attachments
            string strName = pBP->GetName();
            printf("File name: %s\r\n", strName.c_str());
            printf("File size: %d\r\n", pBP->GetContentLength());
			string name = strName;
            strName = string((LPCSTR)GetmboxviewTempPath()) + strName;
            pBP->WriteToFile(strName.c_str());
			// Set item icon and insert in attachment list
			int iIcon = 0;
			SHFILEINFO shFinfo;
			if ( !SHGetFileInfo( strName.c_str(),
						0,
						&shFinfo,
						sizeof( shFinfo ),
						SHGFI_ICON | 
						SHGFI_SMALLICON ) )
			{
				TRACE("Error Gettting SystemFileInfo!\n");
			} else {
				iIcon = shFinfo.iIcon;
				// we only need the index from the system image list
				DestroyIcon( shFinfo.hIcon );
			}
			pView->m_bAttach = TRUE;
			pView->m_attachments.InsertItem(pView->m_attachments.GetItemCount(), name.c_str(), iIcon);
        }
    }
	// Get temporary file name with correct extension for IE to display
	m_curFile = CreateTempFileName(ext);
	// Save mail
	if (ext.Compare("htm") == 0 && CString(bdy).MakeLower().Find("<body") == -1) {
		bdy = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset="+bdycharset+"\"></head><body>"+bdy;
	}
	CFile fp(m_curFile, CFile::modeWrite|CFile::modeCreate);
	fp.Write(bdy, bdy.GetLength());
	fp.Close();
	// Display mail in IE
	pView->m_browser.Navigate(m_curFile, NULL);
	// Update layou to show/hide attachments
	pView->UpdateLayout();
	return;
}


BOOL NListView::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt) 
{
	CWnd *wnd = WindowFromPoint(pt);
	if( wnd == NULL )
		return FALSE;
	if( wnd == this ) {
		return CWnd::OnMouseWheel(nFlags, zDelta, pt);
	}
	return wnd->PostMessage(WM_MOUSEWHEEL, MAKELPARAM(nFlags,zDelta), MAKELPARAM(pt.x, pt.y));
}

void NListView::ClearDescView()
{
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if( ! pFrame )
		return;
	NMsgView *pView = pFrame->GetMsgView();
	if( ! pView )
		return;
	pView->m_strDescp1.LoadString(IDS_DESC_NONE);
	pView->m_strDescp2.LoadString(IDS_DESC_NONE);
	pView->m_strDescp3.LoadString(IDS_DESC_NONE);
	pView->m_attachments.DeleteAllItems();
	RemoveDir(GetmboxviewTempPath());
	m_curFile.Empty();
	pView->m_browser.Navigate("about:blank", NULL);
	pView->UpdateLayout();
	m_lastSel = -1;
}

void NListView::OnUpdateEditFind(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(! m_bInFind && m_list.GetItemCount()>0 );
	return;
	
}

#include "FindDlg.h"

time_t OleToTime_t(COleDateTime *ot) {
	SYSTEMTIME st;
	ot->GetAsSystemTime(st);
	st.wHour = 0;
	st.wMinute = 0;
	st.wSecond = 0;
	CTime ct(st);
	return ct.GetTime();
}

void NListView::OnEditFind()
{
	if (m_bInFind)
		return;
	m_bInFind = true;
	CFindDlg dlg;

	dlg.m_bWholeWord = m_bWholeWord;
	dlg.m_bCaseSensitive = m_bCaseSens;
	dlg.m_string = m_searchString;
	if (m_lastStartDate <= 0) {
		time_t min = time(NULL);
		time_t max = 0;
		int sz = MboxMail::s_mails.GetSize();
		for (int i = 0; i < sz; i++) {
			time_t t = MboxMail::s_mails[i]->m_timeDate;
			if (t < min && t > 0)
				min = t;
			if (t > max)
				max = t;
		}
		m_lastStartDate = CTime(min);
		if (max == 0)
			m_lastEndDate = CTime(CTime::GetCurrentTime());
		else
			m_lastEndDate = CTime(max);
	}
	dlg.m_startDate = COleDateTime(m_lastStartDate.GetTime());
	dlg.m_endDate = COleDateTime(m_lastEndDate.GetTime());
	dlg.m_scope = m_lastScope;
	dlg.m_filterDates = m_filterDates;

	if (MboxMail::b_mails_sorted == true) {
		MboxMail::s_mails.SetSize(MboxMail::s_mails_ref.GetSize());
		MboxMail::s_mails.Copy(MboxMail::s_mails_ref);
		MboxMail::b_mails_sorted = false;

		RedrawMails();
	}

	if (dlg.DoModal() == IDOK) {
		m_filterDates = dlg.m_filterDates;
		m_searchString = dlg.m_string;
		m_bWholeWord = dlg.m_bWholeWord;
		m_bCaseSens = dlg.m_bCaseSensitive;
		m_lastStartDate = CTime(OleToTime_t(&dlg.m_startDate));
		m_lastEndDate = CTime(OleToTime_t(&dlg.m_endDate));
		m_lastScope = dlg.m_scope;
		m_searchPos = 0;
		m_startoff = m_endoff = 0;
		int sz = MboxMail::s_mails.GetSize();
		if (sz > 0) {
/*			if (m_filterDates) {
				int i = 0;
				TRACE("%s to %s\n", m_lastStartDate.Format("%Y%m%d%H%M%S"), m_lastEndDate.Format("%Y%m%d%H%M%S"));
				_int64 startoff = m_lastStartDate.GetTime(), endoff = m_lastEndDate.GetTime();
				int whichStart = -1, whichEnd = -1;
				for (i = 0; i < sz; i++) {
					MboxMail *m = MboxMail::s_mails[i];
					time_t curt = m->m_timeDate;
					if (curt == 0)
						continue;
					if (whichEnd == -1 && startoff <= curt)
						whichStart = i;
					if (whichEnd == -1 && curt >= endoff)
						whichEnd = i;
				}
				if (whichStart >= 0)
					m_startoff = MboxMail::s_mails[whichStart]->m_startOff;
				else
					m_startoff = 0;
				if (whichEnd >= 0)
					m_endoff = MboxMail::s_mails[whichEnd]->m_startOff + MboxMail::s_mails[whichEnd]->m_length;
				else
					m_endoff = MboxMail::s_fSize;
			}
			else*/
			if (m_lastScope) {
				{
					m_startoff = 0;
					m_endoff = MboxMail::s_fSize;
				}
				_int64 w = DoFind(m_startoff, m_endoff);
				if (w >= 0)
					SelectPos(w);
				else {
					MessageBeep(MB_OK);
					m_searchPos = 0;
				}
			}
			else {
				int which = 0, w = -1;
				if (m_searchPos)
					which = WhichOne(m_searchPos - 1);
				w = DoFastFind(which);
				if (w >= 0)
					SelectItemFound(w);
				else {
					MessageBeep(MB_OK);
					m_searchPos = 0;
				}
			}
		}
		else {
			MessageBeep(MB_OK);
			m_searchPos = 0;
		}
	}
	m_bInFind = false;
}

int NListView::DoFastFind(int which)
{
	int w = -1;
	int sz = MboxMail::s_mails.GetSize();
	time_t sd = m_lastStartDate.GetTime(), ed = m_lastEndDate.GetTime();
	for (int i = which; i < sz; i++) {
		MboxMail *m = MboxMail::s_mails[i];
		bool process = false;
		if (m_filterDates) {
			if (m->m_timeDate >= sd && m->m_timeDate <= ed)
				process = true;
		}
		else
			process = true;
		if (process) {
			if (m_searchString.GetLength() == 1 && m_searchString[0] == '*') {
				w = i;
				break;
			}
			int pos = -1;
			if (m_bWholeWord)
				pos = g_tu.BMHSearchW((unsigned char *)(LPCSTR)m->m_subj, m->m_subj.GetLength(), (unsigned char *)(LPCSTR)m_searchString, m_searchString.GetLength(), m_bCaseSens);
			else
				pos = g_tu.BMHSearch((unsigned char *)(LPCSTR)m->m_subj, m->m_subj.GetLength(), (unsigned char *)(LPCSTR)m_searchString, m_searchString.GetLength(), m_bCaseSens);
			if (pos >= 0) {
				w = i;
				break;
			}
			pos = -1;
			if (m_bWholeWord)
				pos = g_tu.BMHSearchW((unsigned char *)(LPCSTR)m->m_from, m->m_from.GetLength(), (unsigned char *)(LPCSTR)m_searchString, m_searchString.GetLength(), m_bCaseSens);
			else
				pos = g_tu.BMHSearch((unsigned char *)(LPCSTR)m->m_from, m->m_from.GetLength(), (unsigned char *)(LPCSTR)m_searchString, m_searchString.GetLength(), m_bCaseSens);
			if (pos >= 0) {
				w = i;
				break;
			}
		}
	}
	return w;
}
void NListView::OnUpdateEditFindAgain(CCmdUI* pCmdUI) 
{
	pCmdUI->Enable(! m_bInFind && m_list.GetItemCount()>0);
	return;
	
}

void NListView::OnEditFindAgain() 
{
	if( m_bInFind )
		return;
	if( m_searchString.IsEmpty() ) {
		OnEditFind();
		return;
	}
	m_bInFind = true;
	if (m_lastScope) {
		_int64 w = DoFind(m_startoff, m_endoff);
		if (w >= 0)
			SelectPos(w);
		else {
			MessageBeep(MB_OK);
			m_searchPos = 0;
		}
	}
	else {
		int which = 0, w = -1;
		if (m_searchPos)
			which = WhichOne(m_searchPos-1)+1;
		w = DoFastFind(which);
		if (w >= 0)
			SelectItemFound(w);
		else {
			MessageBeep(MB_OK);
			m_searchPos = 0;
		}
	}
	m_bInFind = false;
}

_int64 NListView::DoFind(_int64 searchstart, _int64 searchend)
{
	HANDLE hFile = CreateFile(MboxMail::s_path, GENERIC_READ, 0, NULL,
							OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL);
	if( hFile == INVALID_HANDLE_VALUE )
		return -1;

	LARGE_INTEGER li;
	GetFileSizeEx(hFile, &li);
	_int64 fSize = li.QuadPart;
	if( fSize < 10 ) {
		CloseHandle(hFile);
		return -1;
	}
	int mappingsInFile = (int)(fSize / (_int64)MAPPING_SIZE);
	HANDLE hFileMap = CreateFileMapping(hFile, NULL, PAGE_READONLY, (DWORD)(fSize >> 32), (DWORD)fSize, NULL);
	if( hFileMap == NULL ) {
		CloseHandle( hFile );
		return -1;
	}
#ifdef _DEBUG
	DWORD tc = GetTickCount();
#endif
	TRACE("fSize = %lld\n", fSize);
	_int64 pos;
	_int64 offset = 0;
	for( int i = 0; i <= mappingsInFile; i ++ ) {
		_int64 curmap = (_int64)i * (_int64)MAPPING_SIZE;
		DWORD bufSize = ((fSize - curmap < MAPPING_SIZE) ? (int)(fSize - curmap) : MAPPING_SIZE);
		_int64 curend = curmap + (_int64) bufSize;
		if( curend < m_searchPos )
			continue;
		if ( searchstart > curend )
			continue;
		if (searchend < curmap)
			break;
		TRACE("offset=%lld, bufsize=%ld, fSize-curmap=%lld, end=%lld\n", curmap, bufSize, fSize - curmap, curmap + bufSize);
		char * p = (char *) MapViewOfFileEx(hFileMap, FILE_MAP_READ, (DWORD)(curmap >> 32), (DWORD)curmap, bufSize, NULL);
		char *orig = p;
		char *e = p + bufSize;
		if( p == NULL || fSize < 10) {
			CloseHandle( hFileMap );
			CloseHandle( hFile );
			return -1;
		}
		if( m_searchPos > curmap && curend > m_searchPos ) {
			p += m_searchPos - curmap;
			bufSize -= (DWORD)(m_searchPos - curmap);
			offset = m_searchPos;
		}
		if (searchstart < curend && searchstart - curmap > p - orig) {
			p += (searchstart - curmap) - (p - orig);
			bufSize -= (DWORD)(searchstart - curmap) - (p - orig);
			offset = searchstart;
		}
		if (searchend < curend && searchend > curmap)
			bufSize -= (DWORD)(curend - searchend);
		pos = -1;
		if( m_bWholeWord )
			pos = (_int64)g_tu.BMHSearchW((unsigned char *)p, bufSize, (unsigned char *)(LPCSTR)m_searchString, m_searchString.GetLength(), m_bCaseSens);
		else
			pos = (_int64)g_tu.BMHSearch((unsigned char *)p, bufSize, (unsigned char *)(LPCSTR)m_searchString, m_searchString.GetLength(), m_bCaseSens);

		UnmapViewOfFile( orig );
		if( pos >= 0 ) {
			pos += offset ? offset : curmap;
			ASSERT(pos < curend);
			break;
		}
	}
	CloseHandle( hFileMap );
	CloseHandle( hFile );
	return pos;
}

int NListView::WhichOne( _int64 offset, int hint, int lowhint, int highhint )
{
	TRACE("WhichOne(%d, %d, %d, %d)\n", offset, hint, lowhint, highhint );
	if( highhint == -1 )
		highhint = MboxMail::s_mails.GetSize();
	if( hint == -1 )
		hint = highhint/2;
	if( hint < 0 )
		return -1;
	if( hint >= highhint )
		return -1;
	MboxMail *m = MboxMail::s_mails[hint];
	_int64 eoff = m->m_startOff + m->m_length;
	if( offset >= m->m_startOff && offset <= eoff )
		return hint;
	if( offset < m->m_startOff && hint > lowhint ) {
		return WhichOne(offset, lowhint + (hint-lowhint) / 2, lowhint, hint);
	}
	if( offset >= eoff && hint < highhint ) {
		return WhichOne(offset, hint + (highhint - hint) / 2, hint, highhint);
	}
	return -1;
}

void NListView::SelectPos(_int64 offset)
{
	if (offset < 0) {
		MessageBeep(MB_OK);
		return;
	}

	int which = WhichOne(offset);
	if (which < 0) {
		MessageBeep(MB_OK);
		return;
	}

	POSITION pos = m_list.GetFirstSelectedItemPosition();
	while (pos)
	{
		int nItem = m_list.GetNextSelectedItem(pos);
		m_list.SetItemState(nItem, 0, LVIS_SELECTED | LVIS_FOCUSED);
	}
	m_list.SetItemState(which, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
	m_list.EnsureVisible(which, FALSE);
	SelectItem(which);
	MboxMail *m = MboxMail::s_mails[which];
	m_searchPos = m->m_startOff + m->m_length;
}

void NListView::SelectItemFound(int which)
{
	if (which < 0) {
		MessageBeep(MB_OK);
		return;
	}

	POSITION pos = m_list.GetFirstSelectedItemPosition();
	while (pos)
	{
		int nItem = m_list.GetNextSelectedItem(pos);
		m_list.SetItemState(nItem, 0, LVIS_SELECTED | LVIS_FOCUSED);
	}
	m_list.SetItemState(which, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
	m_list.EnsureVisible(which, FALSE);
	SelectItem(which);
	MboxMail *m = MboxMail::s_mails[which];
	m_searchPos = m->m_startOff + m->m_length;
}


void NListView::OnEditVieweml()
{
	CString path = GetmboxviewTempPath();
	HINSTANCE result = ShellExecute(NULL, _T("open"), path, NULL, NULL, SW_SHOWNORMAL);
}


void NListView::OnUpdateEditVieweml(CCmdUI *pCmdUI)
{
	pCmdUI->Enable(m_list.GetFirstSelectedItemPosition()>0);
}

void NListView::RedrawMails()
{
	// Based on NListView::FillCtrl(). 
	ClearDescView();
	m_searchPos = 0;
	m_list.SetRedraw(FALSE);
	m_list.DeleteAllItems();
	if (m_path.IsEmpty())
		return;
	CMainFrame *pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetApp()->m_pMainWnd);
	if (!pFrame)
		return;
	NTreeView *pView = pFrame->GetTreeView();
	if (!pView)
		return;

	int ni = MboxMail::s_mails.GetSize();
	m_list.SetItemCount(ni);

	m_list.EnsureVisible(ni, FALSE);
	m_list.SetRedraw(TRUE);
	//pView->m_tree.SetItemData(m_which, (DWORD)FileSize(m_path));
	pView->m_tree.SetItemState(m_which, 0, TVIS_BOLD);
	pView->SaveData();
}
