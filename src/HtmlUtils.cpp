//
//////////////////////////////////////////////////////////////////
//
//  Windows Mbox Viewer is a free tool to view, search and print mbox mail archives..
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


#include "StdAfx.h"
#include "SimpleString.h"
#include "TextUtilsEx.h"
#include "FileUtils.h"
#include "Browser.h"
#include "HtmlUtils.h"


#pragma component(browser, off, references)  // wliminates too many refrences warning but at what price ?
#include <Mshtml.h>
#pragma component(browser, on, references)
#include <atlbase.h>

#ifdef _DEBUG
#define HTML_ASSERT _ASSERTE
#else
#define HTML_ASSERT(x)
#endif

//char * HtmlUtils::EatNLine(char* p, char* e) { while (p < e && *p++ != '\n'); return p; }

char * HtmlUtils::SkipEmptyLine(char* p, char* e)
{
	char *p_save = p;
	while ((p < e) && ((*p == ' ') || (*p == '\t')))  // eat empty lines
		p++;
	if ((*p == '\r') && (*p == '\n'))
		p += 2;
	else if ((*p == '\r') || (*p == '\n'))
		p++;
	else
		p = p_save;
	return p;
}

void BreakBeforeGoingCleanup()
{
	int deb = 1;
}

void HtmlUtils::MergeWhiteLines(SimpleString *workbuf, int maxOutLines)
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

		len = IntPtr2Int(p_end_data - p_beg_data);
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

void HtmlUtils::PrintToPrinterPageSetup(CWnd *parent)
{
	//CComBSTR cmdID(L"PRINT");
	//VARIANT_BOOL vBool;

	IHTMLDocument2 *lpHtmlDocument = 0;
	HRESULT hr;
	VARIANT val;
	VariantInit(&val);
	VARIANT valOut;
	VariantInit(&valOut);

	SimpleString inbuf;
	inbuf.Append("<html></html>");
	SimpleString workbuf;
	UINT inCodePage = CP_UTF8;

	BOOL retVal = CreateHTMLDocument(&lpHtmlDocument, &inbuf, &workbuf, inCodePage);
	if ((retVal == FALSE) || (lpHtmlDocument == 0)) {
		return;
	}

	IOleCommandTarget  *lpOleCommandTarget = 0;
	hr = lpHtmlDocument->QueryInterface(IID_IOleCommandTarget, (VOID**)&lpOleCommandTarget);
	if (FAILED(hr) || !lpOleCommandTarget)
	{
		BreakBeforeGoingCleanup();
		goto cleanup;
	}

	hr = lpOleCommandTarget->Exec(NULL, OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER, &val, &valOut);

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

void HtmlUtils::PrintHTMLDocumentToPrinter(SimpleString *inbuf, SimpleString *workbuf, UINT inCodePage, int printDialogType)
{
	//CComBSTR cmdID(L"PRINT");
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
		_ASSERTE(FALSE);
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

	DWORD nCmdId = OLECMDID_PRINT;
	DWORD nCmdOption = OLECMDEXECOPT_DONTPROMPTUSER;

	if (printDialogType == 1) {
		nCmdId = OLECMDID_PRINT;
		nCmdOption = OLECMDEXECOPT_PROMPTUSER;
	}
	else if (printDialogType == 2) {
		nCmdId = OLECMDID_PRINTPREVIEW;
		nCmdOption = OLECMDEXECOPT_PROMPTUSER;
	}

#if 0
	// Trying to figure out how to bypass Save On dialog
	// Doesn't work yet (or ever ?)
	nCmdId = OLECMDID_SAVEAS;
	nCmdOption = OLECMDEXECOPT_DONTPROMPTUSER;

	val.vt = VT_BSTR;
	val.bstrVal = CComBSTR(L"F:\\New\\test.pdf");
#endif

	hr = lpOleCommandTarget->Exec(NULL, nCmdId, nCmdOption, &val, &valOut);
	if (FAILED(hr))
	{
		BreakBeforeGoingCleanup();
		DWORD err = GetLastError();
		int deb = 1;
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

void HtmlUtils::GetTextFromIHTMLDocument(SimpleString *inbuf, SimpleString *workbuf, UINT inCodePage, UINT outCodePage)
{
	IHTMLDocument2 *lpHtmlDocument = 0;
	LPDISPATCH lpDispatch = 0;
	HRESULT hr;

	if (lpHtmlDocument == 0)
	{
		BOOL retVal = CreateHTMLDocument(&lpHtmlDocument, inbuf, workbuf, inCodePage);
		if ((retVal == FALSE) || (lpHtmlDocument == 0)) {
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

	DWORD error;
	BOOL ret = TextUtilsEx::WStr2CodePage(wstr, wlen, outCodePage, workbuf, error);

	MergeWhiteLines(workbuf, -1);

	if (lpBodyElm)
		lpBodyElm->Release();

	if (lpHtmlDocument)
		lpHtmlDocument->Release();
}

BOOL HtmlUtils::CreateHTMLDocument(IHTMLDocument2 **lpDocument, SimpleString *inbuf, SimpleString *workbuf, UINT inCodePage)
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
	DWORD error;
	ret = TextUtilsEx::CodePage2WStr(inbuf, inCodePage, workbuf, error);
	OLECHAR *oledata = (OLECHAR*)workbuf->Data();
	BSTR  bstr = SysAllocString(oledata);
	int bstrLen = SysStringByteLen(bstr);
	int wlen = (int)wcslen(bstr);
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
	hr = lpDoc->put_designMode(L"off");
	hr = lpDoc->writeln(psaStrings);

	//Sleep(2000);
	//BSTR* p = 0;
	//HRESULT hready = lpDoc->get_readyState(p);

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

void HtmlUtils::RemoveStyleTagFromIHTMLDocument(IHTMLElement *lpElm)
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

				CStringA tag(bstrTag);
				if (tag.CompareNoCase("style") == 0) {
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


BOOL HtmlUtils::FindElementByTagInIHTMLDocument(IHTMLDocument2 *lpDocument, IHTMLElement **ppvEl, CStringA &tag)
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

				CStringA eltag(bstrTag);
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

void HtmlUtils::PrintIHTMLDocument(IHTMLDocument2 *lpDocument, SimpleString *workbuf)
{
	// Actually this will print BODY
	CStringA tag("body");
	IHTMLElement *lpElm = 0;
	FindElementByTagInIHTMLDocument(lpDocument, &lpElm, tag);
	if (!lpElm)
	{
		return;
	}
#if 0
	IHTMLElement *lpBodyElm = 0;
	hr = lpHtmlDocument->get_body(&lpBodyElm);
	//HTML_ASSERT(SUCCEEDED(hr) && lpBodyElm);
	if (!lpBodyElm) {
		lpHtmlDocument->Release();
		return;
	}
#endif

	CComBSTR bstrTEXT;
	lpElm->get_innerText(&bstrTEXT);

	CStringW text;
	PrintIHTMLElement(lpElm, text);

	workbuf->ClearAndResize(10000);

	const wchar_t *wstr = text.operator LPCWSTR();
	int wlen = text.GetLength();
	UINT outCodePage = CP_UTF8;
	DWORD error;
	BOOL ret = TextUtilsEx::WStr2CodePage((wchar_t *)wstr, wlen, outCodePage, workbuf, error);

	TRACE("TEXT=%s\n", workbuf->Data());
}

void HtmlUtils::PrintIHTMLElement(IHTMLElement *lpElm, CStringW &text)
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

		CStringA tag(bstrTag);
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


				CStringA tag(bstrTag);
				TRACE("TAG=%s\n", (LPCSTR)tag);

				if (tag.CompareNoCase("style") == 0)
					int deb = 1;

				if (tag.CompareNoCase("tbody") == 0)
					int deb = 1;

				if (tag.CompareNoCase("td") == 0)
					int deb = 1;

				if (tag.CompareNoCase("body") == 0)
					int deb = 1;

				if (tag.CompareNoCase("style") == 0) {
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

void HtmlUtils::FindStringInIHTMLDocument(CBrowser &browser, CString searchID, CString &searchText, BOOL matchWord, BOOL matchCase, CString &matchStyle)
{
	// Based on "Adding a custom search feature to CHtmlViews" on the codeproject by  Marc Richarme, 22 Nov 2000
	// without explicit license
	//  https://www.codeproject.com/Articles/832/Adding-a-custom-search-feature-to-CHtmlViews
	// // mboxview development team did resolve some code issues, enhanced and optimized

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

	ClearSearchResultsInIHTMLDocument(browser, searchID);

	IHTMLDocument2 *lpHtmlDocument = NULL;
	LPDISPATCH lpDispatch = NULL;
	lpDispatch = browser.m_ie.GetDocument();
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
	lpDispatch = 0;


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
	lpHtmlDocument = 0;

	hr = lpBodyElm->QueryInterface(IID_IHTMLBodyElement, (void**)&lpBody);
	HTML_ASSERT(SUCCEEDED(hr) && lpBody);
	if (!lpBody) {
		lpBodyElm->Release();
		return;
	}
	lpBodyElm->Release();
	lpBodyElm = 0;

	hr = lpBody->createTextRange(&lpTxtRange);
	HTML_ASSERT(SUCCEEDED(hr) && lpTxtRange);
	if (!lpTxtRange) {
		lpBody->Release();
		return;
	}
	lpBody->Release();
	lpBody = 0;

	CComBSTR html;
	CComBSTR newhtml;
	CComBSTR search(searchText);

	long t;
	VARIANT_BOOL bFound;
	long count = 0;

	CComBSTR Character(L"Character");
	CComBSTR Textedit(L"Textedit");

	htmlPrfix.Append(L"<span id='");
	htmlPrfix.Append(searchID);
	htmlPrfix.Append(L"' style='");
	htmlPrfix.Append(matchStyle);
	htmlPrfix.Append(L"'>");

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
				parentText = 0;
				continue;
			}
			parentText->Release();
			parentText = 0;
		}
		if (parentText)
		{
			parentText->Release();
			parentText = 0;
		}

		if (firstRange == false) {
			firstRange = true;
			if (lpTxtRange->select() == S_OK)
				lpTxtRange->scrollIntoView(VARIANT_TRUE);
		}

		newhtml.Empty();
		lpTxtRange->get_htmlText(&html);
		newhtml.Append(htmlPrfix);
		if (searchText == L" ")
			newhtml.Append(L"&nbsp;"); // doesn't work very well, but prevents (some) strange presentation
		else
			newhtml.AppendBSTR(html);
		newhtml.Append(L"</span>");

		lpTxtRange->pasteHTML(newhtml);

		lpTxtRange->moveStart(Character, searchText.GetLength(), &t);
		lpTxtRange->moveEnd(Textedit, 1, &t);
	}
	if (lpTxtRange)
	{
		lpTxtRange->Release();
		lpTxtRange = 0;
	}
}

void HtmlUtils::ClearSearchResultsInIHTMLDocument(CBrowser& browser, CString& searchID)
{
	// Based on "Adding a custom search feature to CHtmlViews" on the codeproject by  Marc Richarme, 22 Nov 2000
	// without explicit license
	//  https://www.codeproject.com/Articles/832/Adding-a-custom-search-feature-to-CHtmlViews
	// mboxview development team did resolve some code issues, enhanced and optimized
	// TODO: should we just use DOM instead of this hack ??

	CComBSTR testid(searchID);
	CComBSTR testtag(L"SPAN");

	HRESULT hr;
	IHTMLDocument2* lpHtmlDocument = NULL;
	LPDISPATCH lpDispatch = NULL;
	lpDispatch = browser.m_ie.GetDocument();
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

	IHTMLElementCollection* lpAllElements = 0;
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

	IUnknown* lpUnk;
	IEnumVARIANT* lpNewEnum;
	if (SUCCEEDED(lpAllElements->get__newEnum(&lpUnk)) && (lpUnk != NULL))
	{
		hr = lpUnk->QueryInterface(IID_IEnumVARIANT, (void**)&lpNewEnum);
		HTML_ASSERT(SUCCEEDED(hr) && lpNewEnum);
		if (!lpNewEnum) {
			lpAllElements->Release();
			return;
		}
		VARIANT varElement;
		IHTMLElement* lpElement;

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

void HtmlUtils::CommonMimeType2DocumentTypes(CStringA &contentType, CString &documentExtension)
{
	CStringA documentExtensionA;
	CommonMimeType2DocumentTypes(contentType, documentExtensionA);
	DWORD error;
	UINT CP_US_ASCII = 20127;
	UINT strCodePage = CP_US_ASCII;
	BOOL ret = TextUtilsEx::CodePage2WStr(&documentExtensionA, strCodePage, &documentExtension, error);
}

void HtmlUtils::CommonMimeType2DocumentTypes(CStringA &contentType, CStringA &documentExtension)
{
	// use the hash table to optimize

	documentExtension.Empty();
	if (contentType.CompareNoCase("text/plain") == 0)
	{
		documentExtension.Append(".txt");
	}
	else if (contentType.CompareNoCase("text/html") == 0)
	{
		documentExtension.Append(".htm");
	}
	else if (contentType.CompareNoCase("application/msword") == 0)
	{
		documentExtension.Append(".doc");
	}
	else if (contentType.CompareNoCase("application/vnd.amazon.ebook") == 0)
	{
		documentExtension.Append(".azw");
	}
	else if (contentType.CompareNoCase("application/vnd.ms-excel") == 0)
	{
		documentExtension.Append(".xls");
	}
	else if (contentType.CompareNoCase("application/xhtml+xml") == 0)
	{
		documentExtension.Append(".xhtml");
	}
	else if (contentType.CompareNoCase("application/vnd.visio") == 0)
	{
		documentExtension.Append(".vsd");
	}
	else if (contentType.CompareNoCase("audio/webm") == 0)
	{
		documentExtension.Append(".weba");
	}
	else if (contentType.CompareNoCase("application/x-tar") == 0)
	{
		documentExtension.Append(".tar");
	}
	else if (contentType.CompareNoCase("image/svg+xml") == 0)
	{
		documentExtension.Append(".svg");
	}
	else if (contentType.CompareNoCase("application/vnd.ms-powerpoint") == 0)
	{
		documentExtension.Append(".ppt");
	}
	else if (contentType.CompareNoCase("application/java-archive") == 0)
	{
		documentExtension.Append(".jar");
	}
	else if (contentType.CompareNoCase("video/x-msvideo") == 0)
	{
		documentExtension.Append(".avi");
	}
	else if (contentType.CompareNoCase("application/x-7z-compressed") == 0)
	{
		documentExtension.Append(".7z");
	}
	else if (contentType.CompareNoCase("message/rfc822") == 0)
	{
		documentExtension.Append(".eml");
	}
	else if (contentType.CompareNoCase("text/rfc822-headers") == 0)
	{
		documentExtension.Append(".eml");
	}
	else if (contentType.CompareNoCase("application/epub+zip") == 0)
	{
		documentExtension.Append(".epub");
	}
	else if (contentType.CompareNoCase("video/mp2t") == 0)
	{
		documentExtension.Append(".ts");
	}
}

// Best Effort only. Text is not formatted properly but that is fine for searching text
// TODO: explore DOM and innerText. Already tried IHTMLDocument2 and it was very slow. Need to investigate more
void HtmlUtils::ExtractTextFromHTML_BestEffort(SimpleString *inbuf, SimpleString *outbuf, UINT inCodePage, UINT outCodePage)
{
	static char* bodyTag = "body";
	static int bodyTagLen = istrlen(bodyTag);

	static char* styleTag = "<style";
	static int styleTagLen = istrlen(styleTag);

	char *tagBeg;
	int tagLenRet;
	int dataLen = 0;
	int datalen = 0;

	int ret = HtmlUtils::FindHtmlTag(inbuf->Data(), inbuf->Count(), bodyTag, bodyTagLen, tagBeg, tagLenRet);
	if (ret)
	{
		char *inData = tagBeg + tagLenRet;
		datalen = inbuf->Count() - IntPtr2Int(tagBeg + tagLenRet - inbuf->Data());
		_ASSERTE(datalen >= 0);

		if (*inData == '<')
		{
			char *e = inData+datalen;
			char *mark = inData;
			// There could be multiple <style> tags within body. We wil handle just one for now.
			// Other <style. tags wil polute text unfortunately
			if (TextUtilsEx::strncmpUpper2Lower(inData, e, styleTag, styleTagLen) == 0)
			{
				inData += styleTagLen;
				datalen -= styleTagLen;
				while (inData < e)
				{
					if (*inData == '<')
						break;
					else
					{
						inData++;
						datalen--;
					}
				}
				if (inData >= e)
					return;

				dataLen = inbuf->Count() - IntPtr2Int(inData - inbuf->Data());
				_ASSERTE(datalen >= 0);
				_ASSERTE(dataLen == datalen);

				int deb = 1;
			}
		}

		// Replace al tags with the single blank
		int ret = HtmlUtils::ReplaceAllHtmlTags(inData, datalen, outbuf);

		int deb = 1;
	}
	else
	{
		char *inData = inbuf->Data();
		int  datalen = inbuf->Count();

		int ret = HtmlUtils::ReplaceAllHtmlTags(inData, datalen, outbuf);

		int deb = 1;
	}
}

int HtmlUtils::FindHtmlTag(char *inData, int indDataLen, char *tag, int tagLen, char *&tagBeg, int &tagLenRet)
{
	tagBeg = 0;
	tagLenRet = 0;

	char *p = inData;
	char *e = p + indDataLen;

	while (p < e)
	{
		while (p < e)
		{
			if (*p == '<')
				break;
			else
				p++;
		}
		if (p >= e)
			return 0;

		tagBeg = p++;

		if (tag)
		{
			if (TextUtilsEx::strncmpUpper2Lower(p, e, tag, tagLen) != 0)
			{
				continue;
			}
		}

		p += tagLen;

		while (p < e)
		{ 
			if (*p == '>')
				break;
			else if (*p == '<')
				break;
			else
				p++;
		}
		if (p >= e)
			return 0;

		if (*p == '>')
		{
			tagLenRet = IntPtr2Int(p + 1 -tagBeg);
			_ASSERTE(tagLenRet >= 0);
			break;
		}
	}
	return tagLenRet;
}

int HtmlUtils::ReplaceAllHtmlTags(char *inData, int inDataLen, SimpleString *outbuf)
{
	char *tagBeg = 0;
	int tagLenRet = 0;

	char *p = inData;
	char *e = p + inDataLen;

	int datalen = inDataLen;
	int len;
	int ret;

	while (p < e)
	{
		ret = HtmlUtils::FindHtmlTag(p, datalen, 0, 0, tagBeg, tagLenRet);
		if (ret == 0)
			break;

		len = IntPtr2Int(tagBeg - p);
		_ASSERTE(len >= 0);
		outbuf->Append(p, len);
		if (outbuf->Data()[outbuf->Count()-1] != ' ')
			outbuf->Append(' ');

		p = tagBeg + tagLenRet;
		datalen -= len + tagLenRet;
		_ASSERTE(datalen == (e - p));
		int deb = 1;
	}
	if (outbuf->Count())
	{
		len = IntPtr2Int(e - p);
		_ASSERTE(len >= 0);
		outbuf->Append(p, len);
	}

	return 1;
}

// Needs more investigation. More error checking, etc
int HtmlUtils::CreateTranslationHtml(CString& inputFile, CString& targetLanguageCode, CString& outputHtmlFile)
{
	CStringA languageCode = targetLanguageCode;
	SimpleString txt;
	BOOL retval = FileUtils::ReadEntireFile(inputFile, txt);

	char* inData = txt.Data();
	int inDataLen = txt.Count();
	SimpleString encbuff;

	TextUtilsEx::EncodeAsHtmlText(inData, inDataLen, &encbuff);


CStringA htmlDocHdrPlusStyle = LR"----(
<!DOCTYPE html>
<html lang="en-US">
<head>
<script type="text/javascript">

function googleTranslateElementInit()
{
    new google.translate.TranslateElement({ pageLanguage: 'en', 
    includedLanguages: 'es,it,pt,pt-PT,pt-BR,fr,de,pl'},
    'google_translate_element');
}

function myTranslate(lang)
{
    var selectElement = document.querySelector('#google_translate_element select');
    selectElement.value = lang;
    selectElement.dispatchEvent(new Event('change'));
 };

</script>

<script type="text/javascript" src="https://translate.google.com/translate_a/element.js?cb=googleTranslateElementInit"></script>
<script type="text/javascript" src="https://translate.google.com/#en/en/Hello"></script>
</head>

<style>

.goog-te-gadget {
        font-size: 19px !important;
}  

body > .skiptranslate {
    //display: none;
}
.goog-te-banner-frame.skiptranslate {
    display: none !important;
} 

body {
    top: 48px !important; 
	font-size: 18px !important;
}

@media print {
  #google_translate_element {display: none;}
}
</style>

)----";

CStringA htmlBodyDivTranslate = LR"----(

<!-- Text that WILL be translated -->
<div id="google_translate_element"></div>

)----";

CStringA htmlBodyAllsDone = LR"----(
<br>
</body>
</html>
)----";


CString htmlBodyDicNoTranslate = LR"----(

<!-- Text that will NOT be translated -->
<div class="notranslate">

<p>
This text in English and should not be translated.
</p>

</div>

)----";

	CString errorText;
	BOOL truncate = TRUE;
	HANDLE hOutFile = FileUtils::FileOpen(outputHtmlFile, errorText, truncate);
	if (hOutFile == INVALID_HANDLE_VALUE)
		return -1;

	DWORD nNumberOfBytesToWrite = htmlDocHdrPlusStyle.GetLength();;
	DWORD nNumberOfBytesWritten = 0;
	int retcnt = FileUtils::Write2File(hOutFile, (LPCSTR)htmlDocHdrPlusStyle, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
	if (retcnt < 0)
	{
		FileUtils::FileClose(hOutFile);
		return -1;
		// TODO or GOTO
	}

	CStringA bodyTranslate;
	bodyTranslate.Format(R"----(<body onload = "setTimeout(myTranslate, 1000, '%s')" >)----", languageCode);

	nNumberOfBytesToWrite = bodyTranslate.GetLength();
	nNumberOfBytesWritten = 0;
	retcnt = FileUtils::Write2File(hOutFile, (LPCSTR)bodyTranslate, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
	if (retcnt < 0)
	{
		FileUtils::FileClose(hOutFile);
		return -1;
	}

	nNumberOfBytesToWrite = htmlBodyDivTranslate.GetLength();
	nNumberOfBytesWritten = 0;
	retcnt = FileUtils::Write2File(hOutFile, (LPCSTR)htmlBodyDivTranslate, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
	if (retcnt < 0)
	{
		FileUtils::FileClose(hOutFile);
		return -1;
	}

	CString inputFileName;

	FileUtils::GetFileName(inputFile, inputFileName);

	CStringA inputFileNameA;
	DWORD error;
	BOOL retw2utf8 = TextUtilsEx::WStr2UTF8(&inputFileName, &inputFileNameA, error);

	CStringA titleFile;
	titleFile.Format(R"---(<br><pre>---  %s  ---</pre><br>)---", inputFileNameA);

	BOOL WStr2UTF8(CString * strW, CStringA * resultA, DWORD & error);

	nNumberOfBytesToWrite = titleFile.GetLength();
	nNumberOfBytesWritten = 0;
	retcnt = FileUtils::Write2File(hOutFile, (LPCSTR)titleFile, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
	if (retcnt < 0)
	{
		FileUtils::FileClose(hOutFile);
		return -1;
	}

	nNumberOfBytesToWrite = encbuff.Count();;
	nNumberOfBytesWritten = 0;
	retcnt = FileUtils::Write2File(hOutFile, encbuff.Data(), nNumberOfBytesToWrite, &nNumberOfBytesWritten);
	if (retcnt < 0)
	{
		FileUtils::FileClose(hOutFile);
		return -1;
	}

	
	nNumberOfBytesToWrite = htmlBodyAllsDone.GetLength();
	nNumberOfBytesWritten = 0;
	retcnt = FileUtils::Write2File(hOutFile, (LPCSTR)htmlBodyAllsDone, nNumberOfBytesToWrite, &nNumberOfBytesWritten);
	if (retcnt < 0)
	{
		FileUtils::FileClose(hOutFile);
		return -1;
	}

	FileUtils::FileClose(hOutFile);
	return 1;
}

#if 0
// For possible future  as an example
import java.io.*;
import java.util.Scanner;

/** Simplified test of matching tags in an HTML document. */
public class HTML {
	/** Strip the first and last characters off a <tag> string. */
	public static String stripEnds(String t) {
		if (t.length() <= 2) return null; // this is a degenerate tag
		return t.substring(1, t.length() - 1);
	}
	/** Test if a stripped tag string is empty or a true opening tag. */
	public static boolean isOpeningTag(String tag) {
		return (tag.length() == 0) || (tag.charAt(0) != '/');
	}
	/** Test if stripped tag1 matches closing tag2 (first character is '/'). */
	public static boolean areMatchingTags(String tag1, String tag2) {
		return tag1.equals(tag2.substring(1)); // test against name after '/'
	}
	/** Test if every opening tag has a matching closing tag. */
	public static boolean isHTMLMatched(String[] tag) {
		NodeStack<String> S = new NodeStack<String>(); // Stack for matching tags
		for (int i = 0; (i < tag.length) && (tag[i] != null); i++) {
			if (isOpeningTag(tag[i]))
				S.push(tag[i]); // opening tag; push it on the stack
			else {
				if (S.isEmpty())
					return false; // nothing to match
				if (!areMatchingTags(S.pop(), tag[i]))
					return false; // wrong match
			}
		}
		if (S.isEmpty()) return true; // we matched everything
		return false; // we have some tags that never were matched
	}
	public final static int CAPACITY = 1000; // Tag array size
	/* Parse an HTML document into an array of html tags */
	public static String[] parseHTML(Scanner s) {
		String[] tag = new String[CAPACITY]; // our tag array (initially all null)
		int count = 0; // tag counter
		String token; // token returned by the scanner s
		while (s.hasNextLine()) {
			while ((token = s.findInLine("<[^>]*>")) != null) // find the next tag
				tag[count++] = stripEnds(token); // strip the ends off this tag
			s.nextLine(); // go to the next line
		}
		return tag; // our array of (stripped) tags
	}
	public static void main(String[] args) throws IOException { // tester
		if (isHTMLMatched(parseHTML(new Scanner(System.in))))
			System.out.println("The input file is a matched HTML document.");
		else
			System.out.println("The input file is not a matched HTML document.");

#endif
