///////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////
// CPictureCtrlDemoDlg.h
// 
// Author: Tobias Eiseler
//
// Adapted for Windows MBox Viewer by the mboxview development
// Simplified, added re-orientation, added next, previous, rotate, zoom and print capabilities
// TODO: Resizing by Mouse Move can be slow for large images
//
// E-Mail: tobias.eiseler@sisternicky.com
// 
// https://www.codeproject.com/Articles/24969/An-MFC-picture-control-to-dynamically-show-picture
///////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////


#pragma once
#include "picturectrl.h"


// CCPictureCtrlDemoDlg-Dialogfeld
class CCPictureCtrlDemoDlg : public CDialogEx
{
// Konstruktion
public:
	CCPictureCtrlDemoDlg(CString *attachmentName, CWnd* pParent = NULL);	// Standardkonstruktor
	~CCPictureCtrlDemoDlg();

	void UpdateRotateType(Gdiplus::RotateFlipType rotateType);
	void FillRect(CBrush &brush);

// Dialogfelddaten
	enum { IDD = IDD_CPICTURECTRLDEMO_DIALOG };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV-Unterstützung
	BOOL LoadImageFileNames(CString & dir);
	void LoadImageFromFile();


// Implementierung
protected:
	HICON m_hIcon;
	CArray<CString*, CString*> m_ImageFileNameArray;
	int m_ImageFileNameArrayPos;
	Gdiplus::RotateFlipType m_rotateType;
	int m_Zoom;  // current zoom multiplier
	int m_ZoomMax;  // maximum zoom iterartions
	int m_ZoomMaxForCurrentImage;  // maximum zoom iterartions determined for the current image

	// Generierte Funktionen für die Meldungstabellen
	// Generated functions for the message tables
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnSize(UINT nType, int cx, int cy);
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedPrev();
	afx_msg void OnBnClickedNext();
	afx_msg void OnBnClickedRotate();
	afx_msg void OnBnClickedZoom();

	CPictureCtrl m_picCtrl;

	static BOOL isSupportedPictureFile(LPCSTR file);
	afx_msg void OnBnClickedButtonPrt();
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
};
