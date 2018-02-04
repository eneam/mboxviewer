// Browser.cpp : implementation file
//

#include "stdafx.h"
#include "Browser.h"
#include "mainfrm.h"
#include "resource.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

void SaveIESettings(bool save = true);
/////////////////////////////////////////////////////////////////////////////
// CBrowser

CBrowser::CBrowser()
{
	m_hWndParent = NULL;
	m_bDidNavigate = false;
	m_bDops = false;
}

BEGIN_MESSAGE_MAP(CBrowser, CWnd)
	//{{AFX_MSG_MAP(CBrowser)
	ON_WM_CREATE()
	ON_WM_SIZE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CBrowser message handlers

BOOL CBrowser::PreCreateWindow(CREATESTRUCT& cs) 
{
	if (!CWnd::PreCreateWindow(cs))
		return FALSE;

	cs.dwExStyle &= ~WS_EX_CLIENTEDGE;
	cs.style &= ~WS_BORDER;
	cs.lpszClass = AfxRegisterWndClass(CS_HREDRAW|CS_VREDRAW|CS_DBLCLKS, //|CS_PARENTDC, 
		::LoadCursor(NULL, IDC_ARROW), HBRUSH(COLOR_WINDOW+1), NULL);

	return TRUE;
}

int CBrowser::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CWnd ::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	if (!m_ie.Create(NULL, WS_VISIBLE | WS_CHILD, CRect(), this, IDC_EXPLORER))
		return -1;

	return 0;
}

void CBrowser::OnSize(UINT nType, int cx, int cy) 
{
	CWnd ::OnSize(nType, cx, cy);
	
	m_ie.MoveWindow( 0, 0, cx, cy );
//	m_ie.UpdateWindow();
}

void CBrowser::Navigate(LPCSTR url, DWORD flags)
{
	COleVariant* pvarURL = new COleVariant( url );
	COleVariant* pvarEmpty = new COleVariant;
	COleVariant* pvarFlags = new COleVariant((long)(flags), VT_I4);
	ASSERT(pvarURL);
	ASSERT(pvarEmpty);
	ASSERT(pvarFlags);
	m_url = url;
	TRACE("CBrowser::Navigate(%s)...", url);
	m_ie.Navigate2( pvarURL, pvarFlags, pvarEmpty, pvarEmpty, pvarEmpty );
	TRACE("Done.\n");
	delete pvarURL;
	delete pvarEmpty;
	delete pvarFlags;
	return;
}

BEGIN_EVENTSINK_MAP(CBrowser, CWnd)
    //{{AFX_EVENTSINK_MAP(CBrowser)
	ON_EVENT(CBrowser, IDC_EXPLORER, 259 /* DocumentComplete */, OnDocumentCompleteExplorer, VTS_DISPATCH VTS_PVARIANT)
	ON_EVENT(CBrowser, IDC_EXPLORER, 250 /* BeforeNavigate */, BeforeNavigate, VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PBOOL)
	//}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()
// DISPID_DOWNLOADCOMPLETE

void CBrowser::OnDocumentCompleteExplorer(LPDISPATCH pDisp, VARIANT FAR* URL) 
{
	IUnknown*  pUnk;
	LPDISPATCH lpWBDisp;
	HRESULT    hr;

	pUnk = m_ie.GetControlUnknown();
	ASSERT(pUnk);

	hr = pUnk->QueryInterface(IID_IDispatch, (void**)&lpWBDisp);
	ASSERT(SUCCEEDED(hr));

	if (pDisp == lpWBDisp )
	{

	}
	lpWBDisp->Release();
}

#include <atlconv.h>

void CBrowser::BeforeNavigate(LPDISPATCH /* pDisp */, VARIANT* URL,
		VARIANT* Flags, VARIANT* TargetFrameName,
		VARIANT* PostData, VARIANT* Headers, BOOL* Cancel)
{
	ASSERT(V_VT(URL) == VT_BSTR);
	ASSERT(Cancel != NULL);
	USES_CONVERSION;

	CString strURL = V_BSTR(URL);
	m_bDidNavigate = true;
	CMainFrame *pFrame = (CMainFrame*)AfxGetMainWnd();
	if( ! pFrame )
		return;
	if( ! ::IsWindow(pFrame->m_hWnd) || ! pFrame->IsKindOf(RUNTIME_CLASS(CMainFrame)) )
		return;
	NMsgView *pView = pFrame->GetMsgView();
	if( ! pView || ! ::IsWindow(pView->m_hWnd))
		return;
	pView->PostMessage(WM_PAINT);
	*Cancel = FALSE;
	return;
}

#include <mshtmdid.h>


BOOL CBrowser::OnAmbientProperty(COleControlSite* pSite, DISPID dispid, VARIANT* pvar) 
{
USES_CONVERSION;
	// Change download properties - no java, no scripts...
	if (dispid == DISPID_AMBIENT_DLCONTROL)
	{
		pvar->vt = VT_I4;
		pvar->lVal = DLCTL_DLIMAGES | DLCTL_VIDEOS | DLCTL_BGSOUNDS | DLCTL_NO_SCRIPTS | DLCTL_NO_JAVA | DLCTL_NO_DLACTIVEXCTLS | DLCTL_NO_RUNACTIVEXCTLS | DLCTL_SILENT; // | DLCTL_NO_CLIENTPULL;

		return TRUE;
	}

	return CWnd::OnAmbientProperty(pSite, dispid, pvar);
}
