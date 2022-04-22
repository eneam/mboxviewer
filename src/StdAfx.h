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

#if 0
// From sdkddkver.h
// _WIN32_WINNT version constants
//
#define _WIN32_WINNT_NT4                    0x0400
#define _WIN32_WINNT_WIN2K                  0x0500
#define _WIN32_WINNT_WINXP                  0x0501
#define _WIN32_WINNT_WS03                   0x0502  // XP + SP1
#define _WIN32_WINNT_WIN6                   0x0600
#define _WIN32_WINNT_VISTA                  0x0600
#define _WIN32_WINNT_WS08                   0x0600
#define _WIN32_WINNT_LONGHORN               0x0600
#define _WIN32_WINNT_WIN7                   0x0601
#define _WIN32_WINNT_WIN8                   0x0602
#define _WIN32_WINNT_WINBLUE                0x0603
#define _WIN32_WINNT_WINTHRESHOLD           0x0A00 /* ABRACADABRA_THRESHOLD*/
#define _WIN32_WINNT_WIN10                  0x0A00 /* ABRACADABRA_THRESHOLD*/
#endif

// The mbox viewer should run on Windows XP and later versions
// That seem to make some MFC functionality not available to us
// Maybe it is time to support Windows 10 only

#ifdef WINVER
#undef WINVER
#endif


#define SUPPORTED_WINVER _WIN32_WINNT_WINXP
//#define SUPPORTED_WINVER _WIN32_WINNT_VISTA

#define WINVER SUPPORTED_WINVER

#ifdef _WIN32_WINNT
#undef _WIN32_WINNT
#endif
#define _WIN32_WINNT SUPPORTED_WINVER
             

#ifdef _WIN32_WINDOWS  
#undef _WIN32_WINDOWS
#endif
#define _WIN32_WINDOWS SUPPORTED_WINVER


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
