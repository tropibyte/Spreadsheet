// This MFC Samples source code demonstrates using MFC Microsoft Office Fluent User Interface 
// (the "Fluent UI") and is provided only as referential material to supplement the 
// Microsoft Foundation Classes Reference and related electronic documentation 
// included with the MFC C++ library software.  
// License terms to copy, use or distribute the Fluent UI are available separately.  
// To learn more about our Fluent UI licensing program, please visit 
// https://go.microsoft.com/fwlink/?LinkId=238214.
//
// Copyright (C) Microsoft Corporation
// All rights reserved.

// MainFrm.h : interface of the CMainFrame class
//

#pragma once
#include "CalendarBar.h"
#include "Resource.h"
#include "../Grid32/grid32.h"
#include "GridCtrl.h"


class COutlookBar : public CMFCOutlookBar
{
	virtual BOOL AllowShowOnPaneMenu() const { return TRUE; }
	virtual void GetPaneName(CString& strName) const { BOOL bNameValid = strName.LoadString(IDS_OUTLOOKBAR); ASSERT(bNameValid); if (!bNameValid) strName.Empty(); }
};

class CMainFrame : public CFrameWndEx
{
	
protected: // create from serialization only
	CMainFrame() noexcept;
	DECLARE_DYNCREATE(CMainFrame)
        void SetupFontPanel();
        void SetupAlignmentPanel();
        void SetupNumberPanel();
        void SetupInsertPanel();
        void SetupClipboardPanel();
        void SetupSortFillPanel();
        // Convenience: active grid control on the active view, or nullptr.
        CGridCtrl* ActiveGrid();

// Attributes
public:

// Operations
public:

// Overrides
public:
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	
// Implementation
public:
	virtual ~CMainFrame();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:  // control bar embedded members
	CMFCRibbonBar     m_wndRibbonBar;
	CMFCRibbonApplicationButton m_MainButton;
	CMFCToolBarImages m_PanelImages;
	CMFCRibbonStatusBar  m_wndStatusBar;
	COutlookBar       m_wndNavigationBar;
	CMFCShellTreeCtrl m_wndTree;
	CCalendarBar      m_wndCalendar;
    CMFCCaptionBar    m_wndCaptionBar;
    //FONTINFO          m_currFontInfo;
    COLORREF          m_cellBkg;
    CFindReplaceDialog* m_pFindReplace = nullptr;  // modeless; self-deletes on close
    BOOL  m_bFormatPainterArmed = FALSE;
    BOOL  m_bFormatPainterPersistent = FALSE;   // armed by a double-click: stays on until clicked again
    FONTINFO m_capturedFont;
public:
    // Called by CGridView when a cell is clicked. Returns true if the click
    // was consumed (Format Painter armed) and the caller should treat it as
    // already handled.
    bool ConsumeGridClick(CGridCtrl* g, const GRIDPOINT& cell);

// Generated message map functions
protected:
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnApplicationLook(UINT id);
	afx_msg void OnUpdateApplicationLook(CCmdUI* pCmdUI);
	afx_msg void OnViewCaptionBar();
	afx_msg void OnUpdateViewCaptionBar(CCmdUI* pCmdUI);
	afx_msg void OnOptions();
	afx_msg void OnFilePrint();
	afx_msg void OnFilePrintPreview();
	afx_msg void OnUpdateFilePrintPreview(CCmdUI* pCmdUI);
	DECLARE_MESSAGE_MAP()

	BOOL CreateOutlookBar(CMFCOutlookBar& bar, UINT uiID, CMFCShellTreeCtrl& tree, CCalendarBar& calendar, int nInitialWidth);
	BOOL CreateCaptionBar();

	int FindFocusedOutlookWnd(CMFCOutlookBarTabCtrl** ppOutlookWnd);

	CMFCOutlookBarTabCtrl* FindOutlookParent(CWnd* pWnd);
	CMFCOutlookBarTabCtrl* m_pCurrOutlookWnd;
	CMFCOutlookBarPane*    m_pCurrOutlookPage;
public:
	afx_msg void OnCellGoto();
	afx_msg void OnFontName();
	afx_msg void OnFontSize();
	afx_msg void OnFontGrow();
	afx_msg void OnFontShrink();
	afx_msg void OnFontBold();
	afx_msg void OnFontItalic();
	afx_msg void OnFontUnderline();
	afx_msg void OnFontStrikethrough();
	afx_msg void OnFontColor();
	afx_msg void OnFontBackground();
	afx_msg void OnFontSubscript();
	afx_msg void OnFontSuperscript();
	afx_msg void OnInsertPicture();
	afx_msg void OnInsertDateTime();
	afx_msg void OnAlignCmd(UINT nID);
	afx_msg void OnFormatCmd(UINT nID);
	afx_msg void OnUpdateAlignCmd(CCmdUI* pCmdUI);
	afx_msg void OnUpdateFormatCmd(CCmdUI* pCmdUI);
	afx_msg void OnUpdateFontBold(CCmdUI* pCmdUI);
	afx_msg void OnUpdateFontItalic(CCmdUI* pCmdUI);
	afx_msg void OnUpdateFontUnderline(CCmdUI* pCmdUI);
	afx_msg void OnUpdateFontStrikethrough(CCmdUI* pCmdUI);
	afx_msg void OnEditFind();
	afx_msg void OnEditReplace();
	afx_msg void OnEditSelectAll();
	afx_msg void OnEditClear();
	afx_msg void OnFormatPainter();
	afx_msg void OnFormatPainterLock();
	afx_msg void OnUpdateFormatPainter(CCmdUI* pCmdUI);
	afx_msg LRESULT OnFindReplaceMsg(WPARAM wParam, LPARAM lParam);
	afx_msg void OnCellFont();
	afx_msg void OnDlgAlignment();
	afx_msg void OnDlgNumber();
	afx_msg void OnDlgInsert();
	afx_msg void OnDlgSort();
	afx_msg void OnDlgFill();
	afx_msg void OnSortCmd(UINT nID);
	afx_msg void OnFillCmd(UINT nID);
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual BOOL LoadFrame(UINT nIDResource, DWORD dwDefaultStyle = WS_OVERLAPPEDWINDOW | FWS_ADDTOTITLE, CWnd* pParentWnd = NULL, CCreateContext* pContext = NULL);
};

