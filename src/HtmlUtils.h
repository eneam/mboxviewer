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

#pragma once

#include "afxstr.h"

class SimpleString;
class CWnd;
struct IHTMLDocument2;
struct IHTMLElement;
class CBrowser;



class HtmlUtils
{
public:

	static void MergeWhiteLines(SimpleString *workbuf, int maxOutLines);
	static void PrintToPrinterPageSetup(CWnd *parent);
	static void PrintHTMLDocumentToPrinter(SimpleString *inbuf, SimpleString *workbuf, UINT inCodePage, int printDialogType);
	static void GetTextFromIHTMLDocument(SimpleString *inbuf, SimpleString *workbuf, UINT inCodePage, UINT outCodePage);
	static BOOL CreateHTMLDocument(IHTMLDocument2 **lpDocument, SimpleString *inbuf, SimpleString *workbuf, UINT inCodePage);
	static void RemoveStyleTagFromIHTMLDocument(IHTMLElement *lpElm);
	static BOOL FindElementByTagInIHTMLDocument(IHTMLDocument2 *lpDocument, IHTMLElement **ppvEl, CStringA &tag);
	static void PrintIHTMLDocument(IHTMLDocument2 *lpDocument, SimpleString *workbuf);
	static void PrintIHTMLElement(IHTMLElement *lpElm, CStringW &text);
	static void FindStringInIHTMLDocument(CBrowser& browser, CString searchID, CString& searchText, BOOL matchWord, BOOL matchCase, CString& m_matchStyle);
	static void ClearSearchResultsInIHTMLDocument(CBrowser& browser, CString& searchID);
	static void CommonMimeType2DocumentTypes(CStringA &contentType, CStringA &documentExtension);
	static void CommonMimeType2DocumentTypes(CStringA &contentType, CString &documentExtension);
	static int FindHtmlTag(char *inData, int indDataLen, char *tag, int tagLen, char *&tagBeg, int &tagDataLen);
	static void ExtractTextFromHTML_BestEffort(SimpleString *inbuf, SimpleString *outbuf, UINT inCodePage, UINT outCodePage);
	static int ReplaceAllHtmlTags(char *inData, int datalen, SimpleString *outbuf);
	static char* EatNLine(char* p, char* e) { while (p < e && *p++ != '\n'); return p; };
	static char* SkipEmptyLine(char* p, char* e);
	static int CreateTranslationHtml(CString& inputFile, CString &targetLanguage, CString& outputHtmlFile);
	static void SetFontSize_Browser(CBrowser& browser, int fontSize);
};