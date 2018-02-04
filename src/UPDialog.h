#ifndef __UNIVERSALPROGRESSDIALOG_H_A0B4F977__97BA__43c0__83A2__6B64410E92E2
#define __UNIVERSALPROGRESSDIALOG_H_A0B4F977__97BA__43c0__83A2__6B64410E92E2

#include <TChar.h>
#include "InitCommonControls.h"

#define CUPDIALOG_CONTROL_CLASSES	(ICC_PROGRESS_CLASS)	//We are using Progress bar Control

class CUPDialog;
typedef struct _CUPDialogUserProcData CUPDUPDATA;
typedef bool (*LP_CUPDIALOG_USERPROC)(const CUPDUPDATA*);

typedef struct _ProgressThreadData
{
protected:
	_ProgressThreadData()
	{
		ZeroMemory(this,sizeof(_ProgressThreadData));
	}
public:
	HWND	hThreadWnd;				//The Dialog that Created the Thread !!
	LPVOID	pUserProcParam;			//Parameter that shoud be sent to the UserProc
	bool	bAlive;					//Indicates the Thread State Alive/Dead
	bool	bTerminate;				//Thread Monitors this value to Know if it has to Terminate itself

	LP_CUPDIALOG_USERPROC	m_lpUserProc;	//User Progress Procedure - Called by ProgressDialogBox From the ThreadProc

	enum							//These would be used by Thread - Should be inSync with DlgProc Values !!
	{
		WM_DISABLECONTROLS = (WM_USER+1234),
		WM_ENABLECONTROLS,
		WM_PROGRESSTHREADCOMPLETED,
		WM_PROGRESSBARUPDATE,
		WM_PROGRESSTEXTUPDATE,
		WM_CANCELPROGRESSTHREAD
	};

}PROGRESSTHREADDATA;

typedef PROGRESSTHREADDATA* LPPROGRESSTHREADDATA;

struct _CUPDialogUserProcData: private _ProgressThreadData
{
	friend class CUPDialog;
private:
	_CUPDialogUserProcData()	{	ZeroMemory(this,sizeof(CUPDUPDATA));	}

public:
	// Use this to retreive the lpUserProcParam supplied as part of CUPDialog constructor
	inline LPVOID GetAppData() const	{	return	this->pUserProcParam;	}

	// Call this regularily inside the lengthy procedure to check if user has requested for cancelling the job
	inline bool ShouldTerminate() const	{	return	this->bTerminate;		}

	// Call this frequently inside the lengthy procedure to set the progress text
	inline void SetProgress(LPCTSTR lpszProgressText) const
	{
		if(::IsWindow(this->hThreadWnd) && this->bTerminate == false)
			::SendMessage(this->hThreadWnd,_ProgressThreadData::WM_PROGRESSTEXTUPDATE,0,(LPARAM)lpszProgressText);
	}
	// Call this frequently inside the lengthy procedure to set the progress bar position
	inline void SetProgress(UINT_PTR dwProgressbarPos) const
	{
		if(::IsWindow(this->hThreadWnd) && this->bTerminate == false)
			::SendMessage(this->hThreadWnd,_ProgressThreadData::WM_PROGRESSBARUPDATE,dwProgressbarPos,0);
	}
	// Call this frequently inside the lengthy procedure to set the progress bar position and text
	inline void SetProgress(LPCTSTR lpszProgressText,UINT_PTR dwProgressbarPos) const
	{
		SetProgress(lpszProgressText);
		SetProgress(dwProgressbarPos);
	}
	// Use this to enable or disable the Cancel button on the progress dialog.
	// bAllow = true enables the cancel button (thus user can terminate the job)
	// bAllow = false disables the cancel button (and thus user cannot terminate the job)
	inline void AllowCancel(bool bAllow) const
	{
		if(::IsWindow(this->hThreadWnd) && this->bTerminate == false)
			::SendMessage(this->hThreadWnd,bAllow?_ProgressThreadData::WM_ENABLECONTROLS:_ProgressThreadData::WM_DISABLECONTROLS,0,0);
	}
	// Modifies the Progress Dialog Caption
	inline void SetDialogCaption(LPCTSTR lpszDialogCaption) const
	{
		if(::IsWindow(this->hThreadWnd) && this->bTerminate == false)
			::SendMessage(this->hThreadWnd,WM_SETTEXT,0,(LPARAM)lpszDialogCaption);
	}
};


class CUPDialog : CInitCommonControls<CUPDIALOG_CONTROL_CLASSES>
{
	HWND	m_hParentWnd;			//The Window that Requested the Progress Operation; Needed to Create the DialogBox

	HINSTANCE m_hInst;				//HINSTANCE of the Module that holds the Dialog Template. (If Dialog template is NULL, this will be ignored.)

	LPCTSTR	m_lpszTemplateName;		//Dialog Template that should be used to display the progress dialog. (If NULL, built-in template will be used)

	int		m_nStaticControlId;		//Static Control Id
	
	int		m_nProgressBarControlId;//Progressbar Control Id

	int		m_nCancelButtonId;		//Cancel button Control Id

	DWORD	m_dwTerminateDelay;		//Amount of time to wait after signaling the termination, in MilliSeconds.

	bool	m_bAllowCancel;			//Should the Dialog allow PreEmtption by user before Completion?

	HANDLE	m_hThread;				//Holds the Handle to the Created Thread

	TCHAR	m_szDialogCaption[256];	//Fill up with the Title of the DialogBox

	CUPDUPDATA 	m_ThreadData;

	friend INT_PTR CALLBACK ProgressDlgProc(HWND,UINT,WPARAM,LPARAM);	//The Windows Message Procedure for the DialogBox

	void Cleanup();

public:
	// Constructs a CUPDialog object. But does not yet display the DialogBox. (To display the dialog box use DoModal())
	//		hParentWnd:	Parent window for the dialog box
	//		lpUserProc:	Lengthy procedure that should be executed
	//		lpUserProcParam: Data that should be supplied to lpUserProc. (Accessible through CUPDUPDATA::GetAppData() in lpUserProc)
	//		lpszDlgTitle: Title of the Dialog box
	//		bAllowCancel: Is user allowed to cancel this dialog box? (Enables or Disables the Cancel button based on this value)
	CUPDialog(HWND hParentWnd,LP_CUPDIALOG_USERPROC lpUserProc,LPVOID lpUserProcParam,LPCTSTR lpszDlgTitle=_T("Please Wait.."),bool bAllowCancel=true);

	// SetDialogTemplate allows any custom dialox box to be used as the progress dialog.
	// Make sure that the custom dialog has one static control, one progress bar control and one cancel button.
	//		hInst:				Module instance where the Dialog Template can be found
	//		lpTemplateName:		Template that describes the custom dialog box
	//		StaticControlId:	Identifier of the Static Control on the dialog box. Used for displaying the Progress Messages
	//		ProgressBarControlId: Identifier of the Progressbar Control on the Dialog Box
	//		CancelButtonId:		Identifier of the Cancel Button Control on the Dialog Box.
	inline void SetDialogTemplate(HINSTANCE hInst, LPCTSTR lpTemplateName, int StaticControlId, int ProgressBarControlId, int CancelButtonId)
	{
		m_hInst = hInst;

		m_lpszTemplateName = lpTemplateName;

		m_nStaticControlId = StaticControlId;
		
		m_nProgressBarControlId = ProgressBarControlId;

		m_nCancelButtonId = CancelButtonId;	
	}

	// SetTerminationDelay sets the amount of time the thread should be allowed to run after receiving a termination
	// request from user before it is forcefully destoryed with TerminateThread.
	// Default is 500Ms.
	inline void SetTerminationDelay(DWORD dwMilliSeconds)
	{
		m_dwTerminateDelay = dwMilliSeconds;
	}

	virtual ~CUPDialog();

	// DoModal invokes the Modal Dialog. 
	// If SetDialogTemplate has been called before, then the supplied template will be used. Else in-built template will be used.
	// Return values: 
	//			(ReturnValue==IDABORT)	=> Unable to Create Thread
	//			(ReturnValue==IDOK)		=> Sucessful;
	//			(LOWORD(ReturnValue)==IDCANCEL && HIWORD(ReturnValue)==0) => Some Error in UserProc;	
	//			(LOWORD(ReturnValue)==IDCANCEL && HIWORD(ReturnValue)==1) => User Cancelled the Dialog;
	INT_PTR DoModal();

	// Invoked after CUPDialog has completed its operation on a message.
	//	if bProcessed is TRUE then CUPDialog has processed this message.
	//	if bProcessed is FALSE then CUPDialog did not process this message (and is likely to be sent to the Default DialogProc).
	// When bProcessed is FALSE,
	//		Return TRUE if you have processed the message.
	//		Return FALSE, to let the CUPDialog send it to the default dialg proc.
	//	When bProcessed is TRUE, the return value from this function is ignored (as the message has already been processed by CUPDialog).
	virtual INT_PTR OnMessage(HWND hDlg,UINT Message,WPARAM wParam,LPARAM lParam, BOOL bProcessed)
	{
		// Note that not all messages come here. Especially the WM_SETFONT that comes before the WM_INITDIALOG
		// CUPDialog uses SetWindowLongPtr() on the Dialog handle in WM_INITDIALOG.
		// Take care not to make recursive SendMessage calls.
		return FALSE; // Return false to continue any Default processing
	}
};

#endif