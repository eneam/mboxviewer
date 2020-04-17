#pragma once


// ColorPickerDlg

class ColorPickerDlg : public CColorDialog
{
	DECLARE_DYNAMIC(ColorPickerDlg)

public:
	ColorPickerDlg(COLORREF clrInit = 0, DWORD dwFlags = 0, CWnd* pParentWnd = nullptr);
	virtual ~ColorPickerDlg();

protected:
	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();

	virtual void OnOK();
	virtual BOOL OnColorOK();
};


