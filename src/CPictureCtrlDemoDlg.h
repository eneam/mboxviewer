///////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////
// CPictureCtrlDemoDlg.h
// 
// Author: Tobias Eiseler
//
// Adapted for Windows MBox Viewer by the mboxview development
// Simplified, added re-orientation, added next, previous, rotate, zoom, dragging and print capabilities
// TODO: Resizing by Mouse Move can be slow for large images
//
// E-Mail: tobias.eiseler@sisternicky.com
// 
// https://www.codeproject.com/Articles/24969/An-MFC-picture-control-to-dynamically-show-picture
///////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////

#if !defined(_PICTURE_CTRL_DEMO_)
#define _PICTURE_CTRL_DEMO_

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
	void EnableZoom(BOOL enableZoom);

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
	float m_Zoom;  // current zoom multiplier
	BOOL m_bZoomEnabled;

	CRect m_rect; // current static rectangle
	float m_hightZoom;
	float m_widthZoom;

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
	CSliderCtrl m_sliderCtrl;
	BOOL m_bDrawOnce;
	int m_sliderRange;
	int m_sliderFreq;

	static BOOL isSupportedPictureFile(LPCSTR file);
	afx_msg void OnBnClickedButtonPrt();
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnMouseMove(UINT nFlags, CPoint point);
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnRButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
	afx_msg void OnBnClickedCancel();
	afx_msg void OnStnClickedStaticPicture();
};

#endif //  _PICTURE_CTRL_DEMO_
