// FindFilterRuleDlg.cpp : implementation file
//

#include "StdAfx.h"
#include "FindFilterRuleDlg.h"
#include "afxdialogex.h"


// FindFilterRuleDlg dialog

IMPLEMENT_DYNAMIC(FindFilterRuleDlg, CDialogEx)

FindFilterRuleDlg::FindFilterRuleDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_FIND_FILTER_DLG, pParent)
{
	m_filterNumb = 0;
}

FindFilterRuleDlg::~FindFilterRuleDlg()
{
}

void FindFilterRuleDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);

	DDX_Radio(pDX, IDC_FILTER_NUMBER0, m_filterNumb);
}


BEGIN_MESSAGE_MAP(FindFilterRuleDlg, CDialogEx)
END_MESSAGE_MAP()


// FindFilterRuleDlg message handlers
