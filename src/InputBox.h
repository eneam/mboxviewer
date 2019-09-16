#pragma once


// InputBox dialog

class InputBox : public CDialogEx
{
	DECLARE_DYNAMIC(InputBox)

public:
	InputBox(CWnd* pParent = nullptr);   // standard constructor
	virtual ~InputBox();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_INPUT_BOX_DLG };
#endif

	CString m_input;

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
};
