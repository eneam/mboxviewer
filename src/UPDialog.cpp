#include "stdafx.h"
#include "updialog.h"
#include "ResHelper.h"

#ifdef _DEBUG
#undef THIS_FILE
#define THIS_FILE __FILE__
#define new DEBUG_NEW
#endif

// Disable the nasty warnings - we know what we are doing !!
#pragma warning(disable : 4244) // disable: conversion from 'LONG_PTR' to 'LONG', possible loss of data
#pragma warning(disable : 4312) // disable: conversion from 'LONG' to 'CUPDialog *' of greater size
#pragma warning(disable : 4996) // disable: 'wcsncpy' was declared deprecated

/////////////////////////////////////////////////////////////////////////////////////////////
// Built-in Dialog Template.
//
// Generated using DlgResToDlgTemplate tool (http://www.codeproject.com/dialog/w32inputbox_1.asp). Thanks to its Author.
//
//	#define IDD_PROGRESS_DIALOG             145
//	#define IDC_PROGRESS_BAR                1018
//	#define IDC_PROGRESS_TEXT               1026
//
//	IDD_PROGRESS_DIALOG DIALOG  0, 0, 187, 70
//	STYLE DS_SETFONT | DS_MODALFRAME | DS_FIXEDSYS | DS_CENTER | WS_POPUP | WS_CLIPCHILDREN | WS_CAPTION | WS_SYSMENU
//	CAPTION "Dialog"
//	FONT 8, "MS Shell Dlg"
//	BEGIN
//		PUSHBUTTON      "&Cancel",IDCANCEL,129,50,50,14
//		CONTROL         "",IDC_PROGRESS_BAR,"msctls_progress32",WS_BORDER,7,25,172,14
//		LTEXT           "Static",IDC_PROGRESS_TEXT,7,7,172,8,SS_CENTERIMAGE
//	END
//

//   The code for Universal Progress Dialog was published by Gopalakrishna Palem on the codeproject under CPOL license
//   https://www.codeproject.com/articles/7700/universal-progress-dialog
//

CString terminationInitiated(L"Termination Initiated..");
CString terminationComplete(L"Termination Complete ..Shutting Down !!");
CString dlgTitle;

static const unsigned char dlg_145[] = 
{
	0xc8,0x08,0xc8,0x82,0x00,0x00,0x00,0x00,0x03,0x00,0x00,0x00,0x00,0x00,0xbb,0x00,0x46,
	0x00,0x00,0x00,0x00,0x00,0x44,0x00,0x69,0x00,0x61,0x00,0x6c,0x00,0x6f,0x00,0x67,
	0x00,0x00,0x00,0x08,0x00,0x4d,0x00,0x53,0x00,0x20,0x00,0x53,0x00,0x68,0x00,0x65,
	0x00,0x6c,0x00,0x6c,0x00,0x20,0x00,0x44,0x00,0x6c,0x00,0x67,0x00,0x00,0x00,0x00,
	0x00,0x01,0x50,0x00,0x00,0x00,0x00,0x81,0x00,0x32,0x00,0x32,0x00,0x0e,0x00,0x02,
	0x00,0xff,0xff,0x80,0x00,0x26,0x00,0x43,0x00,0x61,0x00,0x6e,0x00,0x63,0x00,0x65,
	0x00,0x6c,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x80,0x50,0x00,0x00,0x00,0x00,0x07,
	0x00,0x19,0x00,0xac,0x00,0x0e,0x00,0xfa,0x03,0x6d,0x00,0x73,0x00,0x63,0x00,0x74,
	0x00,0x6c,0x00,0x73,0x00,0x5f,0x00,0x70,0x00,0x72,0x00,0x6f,0x00,0x67,0x00,0x72,
	0x00,0x65,0x00,0x73,0x00,0x73,0x00,0x33,0x00,0x32,0x00,0x00,0x00,0x00,0x00,0x00,
	0x00,0x00,0x00,0x00,0x02,0x02,0x50,0x00,0x00,0x00,0x00,0x07,0x00,0x07,0x00,0xac,
	0x00,0x08,0x00,0x02,0x04,0xff,0xff,0x82,0x00,0x53,0x00,0x74,0x00,0x61,0x00,0x74,
	0x00,0x69,0x00,0x63,0x00,0x00,0x00,0x00,0x00
};
// Do not change these values - they go hand in hand with above template
#define CUPDIALOG_TEXTBOX_ID		(1026)		//Static Control Id
#define CUPDIALOG_PROGRESSBAR_ID	(1018)		//ProgressBar Control Id
#define CUPDIALOG_CANCELBUTTON_ID	(IDCANCEL)	//Cancel Button Control Id
#define CUPDIALOG_TERMINATE_DELAY	(5000)		//Amount of time to wait after signaling the termination, in MilliSeconds.



CUPDialog::CUPDialog(HWND hParentWnd,LP_CUPDIALOG_USERPROC lpUserProc,LPVOID lpUserProcParam,LPCWSTR lpszDlgTitle/*=L"Please Wait.."*/,bool bAllowCancel/*=true*/)
{
	m_hThread = NULL;							//No Thread Yet !!

	m_hParentWnd = hParentWnd;					//Needed to Create the DialogBox - DlgProc asks this as Parameter

	m_hInst = NULL;	

	m_lpszTemplateName = NULL;		//By default we use the built-in template

	m_nStaticControlId = CUPDIALOG_TEXTBOX_ID;		// Default Static Control Id
	
	m_nProgressBarControlId = CUPDIALOG_PROGRESSBAR_ID; // Default Progressbar Control Id

	m_nCancelButtonId = CUPDIALOG_CANCELBUTTON_ID;		// Default Cancel button Control Id

	m_dwTerminateDelay = CUPDIALOG_TERMINATE_DELAY;		//Default Amount of time to wait after signaling the termination, in MilliSeconds.

	m_bAllowCancel = bAllowCancel;				//Is Dialog Terminatable ??

	m_ThreadData.pUserProcParam	= lpUserProcParam;	//We send this as Parameter to the UserProc

	m_ThreadData.m_lpUserProc	= lpUserProc;	//The actual User Procedure

	ZeroMemory(m_szDialogCaption,sizeof(m_szDialogCaption));

	dlgTitle.Empty();
	dlgTitle.Append(lpszDlgTitle);
	ResHelper::TranslateString(dlgTitle);

	_tcsncpy(m_szDialogCaption,(LPCWSTR)dlgTitle,(sizeof(m_szDialogCaption)/sizeof(m_szDialogCaption[0]))-1);
	//_tcsncpy(m_szDialogCaption, lpszDlgTitle, (sizeof(m_szDialogCaption) / sizeof(m_szDialogCaption[0])) - 1);
}

CUPDialog::~CUPDialog(void)
{
	Cleanup();			//It is possible that the Dialog object can be destroyed while Thread is still running..!!
}

void CUPDialog::Cleanup(DWORD dwTerminateDelay)
{
	m_ThreadData.bTerminate = true;
	BOOL forceThreadTermination = FALSE;

	if(m_ThreadData.bAlive)		//If associated Thread is still alive Terminate It
	{
		if (dwTerminateDelay == 0)
			dwTerminateDelay = m_dwTerminateDelay;
		if (WaitForSingleObject(m_hThread, dwTerminateDelay) != WAIT_OBJECT_0)
		{
			TRACE(L"(Cleanup)Terminating Thread\n");
			BOOL terminateStatus = TRUE;
			//_ASSERTE(0);
			if (m_hThread)
			{
				terminateStatus = TerminateThread(m_hThread, IDCANCEL);
				forceThreadTermination = TRUE;
			}
		}
	}

	if(m_hThread)
	{
		CloseHandle(m_hThread);
		m_hThread = NULL;
	}

	BOOL bAlive = m_ThreadData.bAlive;

    m_ThreadData.bAlive = false;
	m_ThreadData.bTerminate = false;
	m_ThreadData.hThreadWnd = NULL;

	if (bAlive)
		Sleep(500);  // hope thread is dead by now

	if (forceThreadTermination)
	{
		// Force Termination leads or may lead to issues
		// Termonated thread doesn't  cleanup heap memory, file handles and ??
		int deb = 1;
	}
}

INT_PTR CUPDialog::DoModal()
{
	INT_PTR retcode = IDCANCEL;
	Cleanup();		//If this is not first time, we had better Terminate any previous instance Threads !!

	if(m_lpszTemplateName == NULL)	// if no custom dialog template is supplied to us, lets use the built-in one
		retcode = DialogBoxIndirectParam(NULL, (LPDLGTEMPLATE) dlg_145, m_hParentWnd, (DLGPROC) ProgressDlgProc, (LPARAM)this); 
	else
		retcode = DialogBoxParam(m_hInst,m_lpszTemplateName,m_hParentWnd,ProgressDlgProc,(LPARAM)this);

	return retcode;
}

static DWORD WINAPI ThreadProc(LPVOID lpThreadParameter)	//Calls the User Progress Procedure
{
    LPPROGRESSTHREADDATA pThreadData = (LPPROGRESSTHREADDATA) lpThreadParameter;

	pThreadData->bAlive = true;

	INT_PTR nResult = IDCANCEL;

	if(pThreadData->bTerminate == true)
		goto TerminateThreadProc;

	nResult = (true == (*pThreadData->m_lpUserProc)((CUPDUPDATA*)lpThreadParameter)) ? IDOK : IDCANCEL;

TerminateThreadProc:

	pThreadData->bAlive = false;

	if(pThreadData->bTerminate == false)
		::PostMessage(pThreadData->hThreadWnd,PROGRESSTHREADDATA::WM_PROGRESSTHREADCOMPLETED,MAKEWPARAM(nResult,0),0);

	return 0;    
}

INT_PTR CALLBACK ProgressDlgProc(HWND hDlg,UINT Message,WPARAM wParam,LPARAM lParam)
{
	BOOL bProcessed = FALSE;
	CUPDialog *pProgressDialog = NULL;

	switch(Message)
	{
	case WM_INITDIALOG:
		{
			pProgressDialog = (CUPDialog*) lParam;
#ifdef GWL_USERDATA
			SetWindowLongPtr(hDlg,GWL_USERDATA,(LONG_PTR)pProgressDialog);
#else
			SetWindowLongPtr(hDlg, GWLP_USERDATA, (LONG_PTR)pProgressDialog);
#endif

			if(pProgressDialog->m_bAllowCancel == false)
				SendMessage(hDlg,PROGRESSTHREADDATA::WM_DISABLECONTROLS,wParam,lParam);

			SendMessage(GetDlgItem(hDlg, pProgressDialog->m_nStaticControlId),WM_SETTEXT,0,(LPARAM)L"");

			SendMessage(GetDlgItem(hDlg, pProgressDialog->m_nProgressBarControlId),PBM_SETPOS,0,0);

			SendMessage(hDlg,WM_SETTEXT,0,(LPARAM)pProgressDialog->m_szDialogCaption);

			((LPPROGRESSTHREADDATA)(LPVOID)(&pProgressDialog->m_ThreadData))->hThreadWnd = hDlg;

			pProgressDialog->OnMessage(hDlg, Message, wParam, lParam, TRUE);

			DWORD dwThreadId = 0;

			pProgressDialog->m_hThread = CreateThread(NULL,NULL,ThreadProc,&pProgressDialog->m_ThreadData,0,&dwThreadId);

			if(pProgressDialog->m_hThread == NULL)	EndDialog(hDlg,IDABORT);

			return TRUE;
		}
	case WM_COMMAND:
		{
#ifdef GWL_USERDATA
		pProgressDialog = (CUPDialog*)GetWindowLongPtr(hDlg, GWL_USERDATA);
#else
		pProgressDialog = (CUPDialog*)GetWindowLongPtr(hDlg,GWLP_USERDATA);
#endif

			if(pProgressDialog->m_nCancelButtonId == LOWORD(wParam))
			{
				SendMessage(hDlg,PROGRESSTHREADDATA::WM_CANCELPROGRESSTHREAD,0,0);
				bProcessed = TRUE;
			}
			break;
		}
	case WM_SYSCOMMAND:
		{
			if(SC_CLOSE == wParam)
			{
				SendMessage(hDlg,PROGRESSTHREADDATA::WM_CANCELPROGRESSTHREAD,0,0);
				bProcessed = TRUE;
			}
			break;
		}
	case PROGRESSTHREADDATA::WM_DISABLECONTROLS:
		{
#ifdef GWL_USERDATA
			pProgressDialog = (CUPDialog*)GetWindowLongPtr(hDlg, GWL_USERDATA);
#else
			pProgressDialog = (CUPDialog*)GetWindowLongPtr(hDlg, GWLP_USERDATA);
#endif

			EnableMenuItem(GetSystemMenu(hDlg,false),SC_CLOSE,MF_DISABLED|MF_GRAYED|MF_BYCOMMAND);
			EnableWindow(GetDlgItem(hDlg, pProgressDialog->m_nCancelButtonId),false);
			bProcessed = TRUE;
			break;
		}
	case PROGRESSTHREADDATA::WM_ENABLECONTROLS:
		{
#ifdef GWL_USERDATA
			pProgressDialog = (CUPDialog*)GetWindowLongPtr(hDlg, GWL_USERDATA);
#else
			pProgressDialog = (CUPDialog*)GetWindowLongPtr(hDlg, GWLP_USERDATA);
#endif
			EnableMenuItem(GetSystemMenu(hDlg,false),SC_CLOSE,MF_ENABLED|MF_BYCOMMAND);
			EnableWindow(GetDlgItem(hDlg, pProgressDialog->m_nCancelButtonId),true);
			bProcessed = TRUE;
			break;
		}
	case PROGRESSTHREADDATA::WM_PROGRESSTHREADCOMPLETED:		//wParam = IDOK or IDCANCEL depending on the Success of Thread
		{
			EndDialog(hDlg, wParam);
			bProcessed = TRUE;
			break;
		}		
	case PROGRESSTHREADDATA::WM_PROGRESSTEXTUPDATE:				//lParam = ProgressText;
		{
#ifdef GWL_USERDATA
		pProgressDialog = (CUPDialog*)GetWindowLongPtr(hDlg, GWL_USERDATA);
#else
				pProgressDialog = (CUPDialog*)GetWindowLongPtr(hDlg, GWLP_USERDATA);
#endif
			SendMessage(GetDlgItem(hDlg, pProgressDialog->m_nStaticControlId),WM_SETTEXT,0,lParam);
			bProcessed = TRUE;
			break;
		}
	case PROGRESSTHREADDATA::WM_PROGRESSBARUPDATE:				//wParam = % Progress; 
		{
#ifdef GWL_USERDATA
			pProgressDialog = (CUPDialog*)GetWindowLongPtr(hDlg, GWL_USERDATA);
#else
			pProgressDialog = (CUPDialog*)GetWindowLongPtr(hDlg, GWLP_USERDATA);
#endif
			SendMessage(GetDlgItem(hDlg, pProgressDialog->m_nProgressBarControlId),PBM_SETPOS,wParam,0);
			bProcessed = TRUE;
			break;
		}
	case PROGRESSTHREADDATA::WM_CANCELPROGRESSTHREAD:			//Enough to Signal the Thread - Actual Handle Would be Closed in the Dialog Destructor
		{
#ifdef GWL_USERDATA
			pProgressDialog = (CUPDialog*)GetWindowLongPtr(hDlg, GWL_USERDATA);
#else
			pProgressDialog = (CUPDialog*)GetWindowLongPtr(hDlg, GWLP_USERDATA);
#endif
			{
				LPPROGRESSTHREADDATA pThreadData = (LPPROGRESSTHREADDATA)(LPVOID)(&pProgressDialog->m_ThreadData);
				pThreadData->bTerminate = true;

				ResHelper::TranslateString(terminationInitiated);

				SendMessage(GetDlgItem(hDlg, pProgressDialog->m_nStaticControlId), WM_SETTEXT, 0, (LPARAM)((LPCWSTR)terminationInitiated));
				//SendMessage(GetDlgItem(hDlg, pProgressDialog->m_nStaticControlId),WM_SETTEXT,0,(LPARAM)(L"Termination Initiated.."));

				SendMessage(hDlg,PROGRESSTHREADDATA::WM_DISABLECONTROLS,wParam,lParam);
				if (pThreadData->bAlive)
				{
					//Sleep(pProgressDialog->m_dwTerminateDelay);
					DWORD sleepTime = 500;
					for (DWORD tm = 0; tm <= pProgressDialog->m_dwTerminateDelay; tm += 500)
					{
						Sleep(sleepTime);
						if (!pThreadData->bAlive)
							break;
					}
				}
				if (pThreadData->bAlive) {
					ResHelper::TranslateString(terminationComplete);
					SendMessage(GetDlgItem(hDlg, pProgressDialog->m_nStaticControlId), WM_SETTEXT, 0, (LPARAM)((LPCWSTR)terminationComplete));
					//SendMessage(GetDlgItem(hDlg, pProgressDialog->m_nStaticControlId), WM_SETTEXT, 0, (LPARAM)(L"Termination Complete ..Shutting Down !!"));
				}
				if (pThreadData->bAlive)
				{
					//Sleep(pProgressDialog->m_dwTerminateDelay);
					DWORD sleepTime = 500;
					for (DWORD tm = 0; tm <= pProgressDialog->m_dwTerminateDelay; tm += 500)
					{
						Sleep(sleepTime);
						if (!pThreadData->bAlive)
							break;
					}
				}
				EndDialog(hDlg, MAKEWPARAM(IDCANCEL,1));
				int deb = 1;
			}
			bProcessed = TRUE;
			break;
		}
	}

#ifdef GWL_USERDATA
	if (pProgressDialog == NULL) // pProgressDialog could be NULL if we have not processed this message (or initialized the value)
		pProgressDialog = (CUPDialog*)GetWindowLongPtr(hDlg, GWL_USERDATA);
#else
	if (pProgressDialog == NULL) // pProgressDialog could be NULL if we have not processed this message (or initialized the value)
		pProgressDialog = (CUPDialog*)GetWindowLongPtr(hDlg, GWLP_USERDATA);
#endif

	if(pProgressDialog != NULL)	// pProgressDialog could be NULL if we have not yet set the GWL_USERDATA
	{
		 //Pass the message to any one who wants to know about it
		 //An Access violation could occur if SetWindowLongPtr has been used on the Dialog.
		INT_PTR RetVal = pProgressDialog->OnMessage(hDlg, Message, wParam, lParam, bProcessed);
			
		if(bProcessed == FALSE) return RetVal;	// if we have not processed the message then return the value we received from OnMessage()
	}

	return bProcessed;	// We have already processed the message
}