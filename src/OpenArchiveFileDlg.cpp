// OpenArchiveFileDlg.cpp : implementation file
//

#include "stdafx.h"
#include "OpenArchiveFileDlg.h"
#include "FileUtils.h"
#include "afxdialogex.h"


// OpenArchiveFileDlg dialog

IMPLEMENT_DYNAMIC(OpenArchiveFileDlg, CDialogEx)

OpenArchiveFileDlg::OpenArchiveFileDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_OPEN_ARCHIVE_DLG, pParent)
{

}

OpenArchiveFileDlg::~OpenArchiveFileDlg()
{
}

void OpenArchiveFileDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_ARCHIVE_FILE_NAME, m_archiveFileName);
}


BEGIN_MESSAGE_MAP(OpenArchiveFileDlg, CDialogEx)
	ON_BN_CLICKED(IDOK, &OpenArchiveFileDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// OpenArchiveFileDlg message handlers

BOOL OpenArchiveFileDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// TODO:  Add extra initialization here

	if (GetSafeHwnd())
	{
		CWnd *p = GetDlgItem(IDC_STATIC);
#if 0
		if (p)
		{
			CString RuleNumberText;

			RuleNumberText.Format("The created archive file will be moved to the folder housing root archive");

			p->SetWindowText(RuleNumberText);
			p->EnableWindow(TRUE);
		}
#endif
		p = GetDlgItem(IDC_EDIT_DFLT_FOLDER);
		if (p)
		{
			p->SetWindowText(m_sourceFolder);
			p->EnableWindow(FALSE);
		}
		p = GetDlgItem(IDC_EDIT_TARGET_FOLDER);
		if (p)
		{
			CString targetFolder = m_targetFolder;
			targetFolder.TrimRight("\\");
			p->SetWindowText(targetFolder);
			p->EnableWindow(FALSE);
		}
		p = GetDlgItem(IDC_EDIT_ARCHIVE_FILE_NAME);
		if (p)
		{
			p->SetWindowText(m_archiveFileName);
			p->EnableWindow(TRUE);
		}
	}

	//UpdateData(TRUE);

	//SetDlgItemText(IDC_STATIC, "Desired Text String")

	//GetDlgItem(IDC_STATIC)->SetWindowText(m_text);

	;// SetWindowText(caption);

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}

#if 0
void OpenArchiveFileDlg::OnBnClickedYes()
{
	// TODO: Add your control notification handler code here

	CString mboxFilePath = m_targetFolder + "\\" + m_archiveFileName;

	if (FileUtils::PathFileExist(mboxFilePath))
	{
		CString txt = _T("File \"") + mboxFilePath;
		txt += _T("\" exists.\nOverwrite?");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
			return;
	}

	EndDialog(IDYES);
}
#endif



void OpenArchiveFileDlg::OnBnClickedOk()
{
	// TODO: Add your control notification handler code here

	UpdateData(TRUE);

	CString mboxFilePath = m_targetFolder + m_archiveFileName;

	if (FileUtils::PathFileExist(mboxFilePath))
	{
		CString txt = _T("File \"") + mboxFilePath;
		txt += _T("\" exists.\nOverwrite?");
		int answer = MessageBox(txt, _T("Info"), MB_APPLMODAL | MB_ICONQUESTION | MB_YESNO);
		if (answer == IDNO)
			return;
	}

	CDialogEx::OnOK();
}
