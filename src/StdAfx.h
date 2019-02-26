// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#if !defined(AFX_STDAFX_H__D51A9CF5_301F_492B_BD6F_CBB4D1516422__INCLUDED_)
#define AFX_STDAFX_H__D51A9CF5_301F_492B_BD6F_CBB4D1516422__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define _AFX_NO_MFC_CONTROLS_IN_DIALOGS

#ifdef WINVER
#undef WINVER
#endif
#define WINVER 0x0501

#ifdef _WIN32_WINNT
#undef _WIN32_WINNT
#endif
#define _WIN32_WINNT 0x0501
             

#ifdef _WIN32_WINDOWS  
#undef _WIN32_WINDOWS
#endif
#define _WIN32_WINDOWS 0x0501


#ifndef _WIN32_IE
#undef _WIN32_IE
#endif
#define _WIN32_IE 0x0500

#define VC_EXTRALEAN		// Exclude rarely-used stuff from Windows headers

#if 0
#include <afxwin.h>         // MFC core and standard components
#include <afxext.h>         // MFC extensions
#include <afxdisp.h>        // MFC Automation classes
#include <afxdtctl.h>		// MFC support for Internet Explorer 4 Common Controls
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>			// MFC support for Windows Common Controls
#endif // _AFX_NO_AFXCMN_SUPPORT

#else

//#include <afxwin.h>         // MFC core and standard components
//#include <afxext.h>         // MFC extensions
#include <afxdisp.h>        // MFC Automation classes
//#include <afxdtctl.h>		// MFC support for Internet Explorer 4 Common Controls
#ifndef _AFX_NO_AFXCMN_SUPPORT
//#include <afxcmn.h>			// MFC support for Windows Common Controls
#endif // _AFX_NO_AFXCMN_SUPPORT
#endif

#include "profile.h"
#include "BrowseForFolder.h"
#include <afxcontrolbars.h>

#include "resource.h"

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__D51A9CF5_301F_492B_BD6F_CBB4D1516422__INCLUDED_)
