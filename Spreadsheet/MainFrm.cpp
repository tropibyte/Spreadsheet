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

// MainFrm.cpp : implementation of the CMainFrame class
//

#include "pch.h"
#include "framework.h"
#include "Spreadsheet.h"
#include "MainFrm.h"
#include "GridView.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CMainFrame

IMPLEMENT_DYNCREATE(CMainFrame, CFrameWndEx)

BEGIN_MESSAGE_MAP(CMainFrame, CFrameWndEx)
	ON_WM_CREATE()
	ON_COMMAND_RANGE(ID_VIEW_APPLOOK_WIN_2000, ID_VIEW_APPLOOK_WINDOWS_7, &CMainFrame::OnApplicationLook)
	ON_UPDATE_COMMAND_UI_RANGE(ID_VIEW_APPLOOK_WIN_2000, ID_VIEW_APPLOOK_WINDOWS_7, &CMainFrame::OnUpdateApplicationLook)
	ON_COMMAND(ID_VIEW_CAPTION_BAR, &CMainFrame::OnViewCaptionBar)
	ON_UPDATE_COMMAND_UI(ID_VIEW_CAPTION_BAR, &CMainFrame::OnUpdateViewCaptionBar)
	ON_COMMAND(ID_TOOLS_OPTIONS, &CMainFrame::OnOptions)
	ON_COMMAND(ID_FILE_PRINT, &CMainFrame::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, &CMainFrame::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, &CMainFrame::OnFilePrintPreview)
        ON_UPDATE_COMMAND_UI(ID_FILE_PRINT_PREVIEW, &CMainFrame::OnUpdateFilePrintPreview)
        ON_COMMAND(ID_CELL_GOTO, &CMainFrame::OnCellGoto)
        ON_COMMAND(ID_FONT_NAME, &CMainFrame::OnFontName)
        ON_COMMAND(ID_FONT_SIZE, &CMainFrame::OnFontSize)
        ON_COMMAND(ID_FONT_GROW, &CMainFrame::OnFontGrow)
        ON_COMMAND(ID_FONT_SHRINK, &CMainFrame::OnFontShrink)
        ON_COMMAND(ID_FONT_BOLD, &CMainFrame::OnFontBold)
        ON_COMMAND(ID_FONT_ITALIC, &CMainFrame::OnFontItalic)
        ON_COMMAND(ID_FONT_UNDERLINE, &CMainFrame::OnFontUnderline)
        ON_COMMAND(ID_FONT_STRIKETHROUGH, &CMainFrame::OnFontStrikethrough)
        ON_COMMAND(ID_FONT_COLOR, &CMainFrame::OnFontColor)
        ON_COMMAND(ID_FONT_CELL_BKGND, &CMainFrame::OnFontBackground)
END_MESSAGE_MAP()

// CMainFrame construction/destruction

CMainFrame::CMainFrame() noexcept : m_pCurrOutlookPage(nullptr), m_pCurrOutlookWnd(nullptr)
{
        // TODO: add member initialization code here
        theApp.m_nAppLook = theApp.GetInt(_T("ApplicationLook"), ID_VIEW_APPLOOK_OFF_2007_SILVER);
        m_currFontInfo = FONTINFO();
        m_currFontInfo.m_wsFontFace = L"Arial";
        m_currFontInfo.m_fPointSize = 12.f;
        m_currFontInfo.m_clrTextColor = RGB(0,0,0);
        m_currFontInfo.bItalic = FALSE;
        m_currFontInfo.bUnderline = FALSE;
        m_currFontInfo.bStrikeThrough = FALSE;
        m_currFontInfo.bWeight = 400;
        m_cellBkg = RGB(255,255,255);
}

CMainFrame::~CMainFrame()
{
}

int CMainFrame::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CFrameWndEx::OnCreate(lpCreateStruct) == -1)
		return -1;

	BOOL bNameValid;

	m_wndRibbonBar.Create(this);
	m_wndRibbonBar.LoadFromResource(IDR_RIBBON);
	SetupFontPanel();

	if (!m_wndStatusBar.Create(this))
	{
		TRACE0("Failed to create status bar\n");
		return -1;      // fail to create
	}

	CString strTitlePane1;
	CString strTitlePane2;
	bNameValid = strTitlePane1.LoadString(IDS_STATUS_PANE1);
	ASSERT(bNameValid);
	bNameValid = strTitlePane2.LoadString(IDS_STATUS_PANE2);
	ASSERT(bNameValid);
	m_wndStatusBar.AddElement(new CMFCRibbonStatusBarPane(ID_STATUSBAR_PANE1, strTitlePane1, TRUE), strTitlePane1);
	m_wndStatusBar.AddExtendedElement(new CMFCRibbonStatusBarPane(ID_STATUSBAR_PANE2, strTitlePane2, TRUE), strTitlePane2);

	// enable Visual Studio 2005 style docking window behavior
	CDockingManager::SetDockingMode(DT_SMART);
	// enable Visual Studio 2005 style docking window auto-hide behavior
	EnableAutoHidePanes(CBRS_ALIGN_ANY);

	// Navigation pane will be created at left, so temporary disable docking at the left side:
	EnableDocking(CBRS_ALIGN_TOP | CBRS_ALIGN_BOTTOM | CBRS_ALIGN_RIGHT);

	//// Create and setup "Outlook" navigation bar:
	//if (!CreateOutlookBar(m_wndNavigationBar, ID_VIEW_NAVIGATION, m_wndTree, m_wndCalendar, 250))
	//{
	//	TRACE0("Failed to create navigation pane\n");
	//	return -1;      // fail to create
	//}

	//// Create a caption bar:
	//if (!CreateCaptionBar())
	//{
	//	TRACE0("Failed to create caption bar\n");
	//	return -1;      // fail to create
	//}

	// Outlook bar is created and docking on the left side should be allowed.
	EnableDocking(CBRS_ALIGN_LEFT);
	EnableAutoHidePanes(CBRS_ALIGN_RIGHT);
	// set the visual manager and style based on persisted value
	OnApplicationLook(theApp.m_nAppLook);

	return 0;
}

BOOL CMainFrame::PreCreateWindow(CREATESTRUCT& cs)
{
	if( !CFrameWndEx::PreCreateWindow(cs) )
		return FALSE;
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return TRUE;
}



BOOL CMainFrame::CreateOutlookBar(CMFCOutlookBar& bar, UINT uiID, CMFCShellTreeCtrl& tree, CCalendarBar& calendar, int nInitialWidth)
{
	bar.SetMode2003();

	BOOL bNameValid;
	CString strTemp;
	bNameValid = strTemp.LoadString(IDS_SHORTCUTS);
	ASSERT(bNameValid);
	if (!bar.Create(strTemp, this, CRect(0, 0, nInitialWidth, 32000), uiID, WS_CHILD | WS_VISIBLE | CBRS_LEFT))
	{
		return FALSE; // fail to create
	}

	CMFCOutlookBarTabCtrl* pOutlookBar = (CMFCOutlookBarTabCtrl*)bar.GetUnderlyingWindow();

	if (pOutlookBar == nullptr)
	{
		ASSERT(FALSE);
		return FALSE;
	}

	pOutlookBar->EnableInPlaceEdit(TRUE);

	static UINT uiPageID = 1;

	// can float, can autohide, can resize, CAN NOT CLOSE
	DWORD dwStyle = AFX_CBRS_FLOAT | AFX_CBRS_AUTOHIDE | AFX_CBRS_RESIZE;

	CRect rectDummy(0, 0, 0, 0);
	const DWORD dwTreeStyle = WS_CHILD | WS_VISIBLE | TVS_HASLINES | TVS_LINESATROOT | TVS_HASBUTTONS;

	tree.Create(dwTreeStyle, rectDummy, &bar, 1200);
	bNameValid = strTemp.LoadString(IDS_FOLDERS);
	ASSERT(bNameValid);
	pOutlookBar->AddControl(&tree, strTemp, 2, TRUE, dwStyle);

	calendar.Create(rectDummy, &bar, 1201);
	bNameValid = strTemp.LoadString(IDS_CALENDAR);
	ASSERT(bNameValid);
	pOutlookBar->AddControl(&calendar, strTemp, 3, TRUE, dwStyle);

	bar.SetPaneStyle(bar.GetPaneStyle() | CBRS_TOOLTIPS | CBRS_FLYBY | CBRS_SIZE_DYNAMIC);

	pOutlookBar->SetImageList(theApp.m_bHiColorIcons ? IDB_PAGES_HC : IDB_PAGES, 24);
	pOutlookBar->SetToolbarImageList(theApp.m_bHiColorIcons ? IDB_PAGES_SMALL_HC : IDB_PAGES_SMALL, 16);
	pOutlookBar->RecalcLayout();

	BOOL bAnimation = theApp.GetInt(_T("OutlookAnimation"), TRUE);
	CMFCOutlookBarTabCtrl::EnableAnimation(bAnimation);

	bar.SetButtonsFont(&afxGlobalData.fontBold);

	return TRUE;
}

BOOL CMainFrame::CreateCaptionBar()
{
	if (!m_wndCaptionBar.Create(WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS, this, ID_VIEW_CAPTION_BAR, -1, TRUE))
	{
		TRACE0("Failed to create caption bar\n");
		return FALSE;
	}

	BOOL bNameValid;

	CString strTemp, strTemp2;
	bNameValid = strTemp.LoadString(IDS_CAPTION_BUTTON);
	ASSERT(bNameValid);
	m_wndCaptionBar.SetButton(strTemp, ID_TOOLS_OPTIONS, CMFCCaptionBar::ALIGN_LEFT, FALSE);
	bNameValid = strTemp.LoadString(IDS_CAPTION_BUTTON_TIP);
	ASSERT(bNameValid);
	m_wndCaptionBar.SetButtonToolTip(strTemp);

	bNameValid = strTemp.LoadString(IDS_CAPTION_TEXT);
	ASSERT(bNameValid);
	m_wndCaptionBar.SetText(strTemp, CMFCCaptionBar::ALIGN_LEFT);

	m_wndCaptionBar.SetBitmap(IDB_INFO, RGB(255, 255, 255), FALSE, CMFCCaptionBar::ALIGN_LEFT);
	bNameValid = strTemp.LoadString(IDS_CAPTION_IMAGE_TIP);
	ASSERT(bNameValid);
	bNameValid = strTemp2.LoadString(IDS_CAPTION_IMAGE_TEXT);
	ASSERT(bNameValid);
	m_wndCaptionBar.SetImageToolTip(strTemp, strTemp2);

	return TRUE;
}

// CMainFrame diagnostics

#ifdef _DEBUG
void CMainFrame::AssertValid() const
{
	CFrameWndEx::AssertValid();
}

void CMainFrame::Dump(CDumpContext& dc) const
{
	CFrameWndEx::Dump(dc);
}
#endif //_DEBUG


// CMainFrame message handlers

void CMainFrame::OnApplicationLook(UINT id)
{
	CWaitCursor wait;

	theApp.m_nAppLook = id;

	switch (theApp.m_nAppLook)
	{
	case ID_VIEW_APPLOOK_WIN_2000:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManager));
		m_wndRibbonBar.SetWindows7Look(FALSE);
		break;

	case ID_VIEW_APPLOOK_OFF_XP:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerOfficeXP));
		m_wndRibbonBar.SetWindows7Look(FALSE);
		break;

	case ID_VIEW_APPLOOK_WIN_XP:
		CMFCVisualManagerWindows::m_b3DTabsXPTheme = TRUE;
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows));
		m_wndRibbonBar.SetWindows7Look(FALSE);
		break;

	case ID_VIEW_APPLOOK_OFF_2003:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerOffice2003));
		CDockingManager::SetDockingMode(DT_SMART);
		m_wndRibbonBar.SetWindows7Look(FALSE);
		break;

	case ID_VIEW_APPLOOK_VS_2005:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerVS2005));
		CDockingManager::SetDockingMode(DT_SMART);
		m_wndRibbonBar.SetWindows7Look(FALSE);
		break;

	case ID_VIEW_APPLOOK_VS_2008:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerVS2008));
		CDockingManager::SetDockingMode(DT_SMART);
		m_wndRibbonBar.SetWindows7Look(FALSE);
		break;

	case ID_VIEW_APPLOOK_WINDOWS_7:
		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows7));
		CDockingManager::SetDockingMode(DT_SMART);
		m_wndRibbonBar.SetWindows7Look(TRUE);
		break;

	default:
		switch (theApp.m_nAppLook)
		{
		case ID_VIEW_APPLOOK_OFF_2007_BLUE:
			CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_LunaBlue);
			break;

		case ID_VIEW_APPLOOK_OFF_2007_BLACK:
			CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_ObsidianBlack);
			break;

		case ID_VIEW_APPLOOK_OFF_2007_SILVER:
			CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_Silver);
			break;

		case ID_VIEW_APPLOOK_OFF_2007_AQUA:
			CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_Aqua);
			break;
		}

		CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerOffice2007));
		CDockingManager::SetDockingMode(DT_SMART);
		m_wndRibbonBar.SetWindows7Look(FALSE);
	}

	RedrawWindow(nullptr, nullptr, RDW_ALLCHILDREN | RDW_INVALIDATE | RDW_UPDATENOW | RDW_FRAME | RDW_ERASE);

	theApp.WriteInt(_T("ApplicationLook"), theApp.m_nAppLook);
}

void CMainFrame::OnUpdateApplicationLook(CCmdUI* pCmdUI)
{
	pCmdUI->SetRadio(theApp.m_nAppLook == pCmdUI->m_nID);
}

void CMainFrame::OnViewCaptionBar()
{
	//m_wndCaptionBar.ShowWindow(m_wndCaptionBar.IsVisible() ? SW_HIDE : SW_SHOW);
	RecalcLayout(FALSE);
}

void CMainFrame::OnUpdateViewCaptionBar(CCmdUI* pCmdUI)
{
	//pCmdUI->SetCheck(m_wndCaptionBar.IsVisible());
}

void CMainFrame::OnOptions()
{
	CMFCRibbonCustomizeDialog *pOptionsDlg = new CMFCRibbonCustomizeDialog(this, &m_wndRibbonBar);
	ASSERT(pOptionsDlg != nullptr);

	pOptionsDlg->DoModal();
	delete pOptionsDlg;
}


void CMainFrame::OnFilePrint()
{
	if (IsPrintPreview())
	{
		PostMessage(WM_COMMAND, AFX_ID_PREVIEW_PRINT);
	}
}

void CMainFrame::OnFilePrintPreview()
{
	if (IsPrintPreview())
	{
		PostMessage(WM_COMMAND, AFX_ID_PREVIEW_CLOSE);  // force Print Preview mode closed
	}
}

void CMainFrame::OnUpdateFilePrintPreview(CCmdUI* pCmdUI)
{
	pCmdUI->SetCheck(IsPrintPreview());
}


void CMainFrame::OnCellGoto()
{
        AfxMessageBox(_T("Ctrl+G pressed"));
}

void CMainFrame::OnFontName()
{
        CMFCRibbonFontComboBox* pCombo = DYNAMIC_DOWNCAST(CMFCRibbonFontComboBox, m_wndRibbonBar.FindByID(ID_FONT_NAME));
        if (pCombo)
        {
                m_currFontInfo.m_wsFontFace = pCombo->GetEditText().GetString();
                ApplyCurrentFont();
        }
}

void CMainFrame::OnFontSize()
{
        CMFCRibbonComboBox* pCombo = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, m_wndRibbonBar.FindByID(ID_FONT_SIZE));
        if (pCombo)
        {
                int size = _ttoi(pCombo->GetEditText());
                if (size > 0)
                {
                        m_currFontInfo.m_fPointSize = (float)size;
                        ApplyCurrentFont();
                }
        }
}

void CMainFrame::OnFontGrow()
{
        m_currFontInfo.m_fPointSize += 1.0f;
        ApplyCurrentFont();
}

void CMainFrame::OnFontShrink()
{
        if (m_currFontInfo.m_fPointSize > 1.0f)
        {
                m_currFontInfo.m_fPointSize -= 1.0f;
                ApplyCurrentFont();
        }
}

void CMainFrame::OnFontBold()
{
        m_currFontInfo.bWeight = (m_currFontInfo.bWeight == 400 ? 700 : 400);
        ApplyCurrentFont();
}

void CMainFrame::OnFontItalic()
{
        m_currFontInfo.bItalic = !m_currFontInfo.bItalic;
        ApplyCurrentFont();
}

void CMainFrame::OnFontUnderline()
{
        m_currFontInfo.bUnderline = !m_currFontInfo.bUnderline;
        ApplyCurrentFont();
}

void CMainFrame::OnFontStrikethrough()
{
        m_currFontInfo.bStrikeThrough = !m_currFontInfo.bStrikeThrough;
        ApplyCurrentFont();
}

void CMainFrame::OnFontColor()
{
        CMFCRibbonColorButton* pBtn = DYNAMIC_DOWNCAST(CMFCRibbonColorButton, m_wndRibbonBar.FindByID(ID_FONT_COLOR));
        if (pBtn)
        {
                m_currFontInfo.m_clrTextColor = pBtn->GetColor();
                ApplyCurrentFont();
        }
}

void CMainFrame::OnFontBackground()
{
        CMFCRibbonColorButton* pBtn = DYNAMIC_DOWNCAST(CMFCRibbonColorButton, m_wndRibbonBar.FindByID(ID_FONT_CELL_BKGND));
        if (pBtn)
        {
                m_cellBkg = pBtn->GetColor();
                // Background color would require full cell update - omitted for brevity
        }
}

void CMainFrame::ApplyCurrentFont()
{
        CGridView* pView = DYNAMIC_DOWNCAST(CGridView, GetActiveView());
        if (pView != nullptr)
        {
                pView->ApplyFont(m_currFontInfo);
        }
}


BOOL CMainFrame::PreTranslateMessage(MSG* pMsg)
{
	if (m_hAccelTable != nullptr && ::TranslateAccelerator(m_hWnd, m_hAccelTable, pMsg))
	{
		return TRUE;
	}
	return CFrameWnd::PreTranslateMessage(pMsg);
}


BOOL CMainFrame::LoadFrame(UINT nIDResource, DWORD dwDefaultStyle, CWnd* pParentWnd, CCreateContext* pContext)
{
	// TODO: Add your specialized code here and/or call the base class

	if (CFrameWndEx::LoadFrame(nIDResource, dwDefaultStyle, pParentWnd, pContext))
	{
		if (m_hAccelTable)
			return TRUE;

		m_hAccelTable = LoadAccelerators(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_MAINFRAME));
		if (m_hAccelTable)
			return TRUE;
	}

	return FALSE;
}
void CMainFrame::SetupFontPanel()
{
        CMFCRibbonCategory* pCategory = nullptr;
        for (int i = 0; i < m_wndRibbonBar.GetCategoryCount(); ++i)
        {
                if (0 == _tcsicmp(m_wndRibbonBar.GetCategory(i)->GetName(), _T("Home")))
                {
                        pCategory = m_wndRibbonBar.GetCategory(i);
                        break;
                }
        }
        if (!pCategory)
                return;

        CMFCRibbonPanel* pFontPanel = nullptr;
        for (int i = 0; i < pCategory->GetPanelCount(); ++i)
        {
                if (0 == _tcsicmp(pCategory->GetPanel(i)->GetName(), _T("Font")))
                {
                        pFontPanel = pCategory->GetPanel(i);
                        break;
                }
        }
        if (!pFontPanel)
                pFontPanel = pCategory->AddPanel(_T("Font"));
        if (!pFontPanel)
                return;

        pFontPanel->RemoveAll();
        pFontPanel->SetCenterColumnVert(0);

        CMFCRibbonFontComboBox* pFontComboBox = new CMFCRibbonFontComboBox(ID_FONT_NAME, TRUE, 120);
        pFontPanel->Add(pFontComboBox);

        CMFCRibbonComboBox* pFontSizeComboBox = new CMFCRibbonComboBox(ID_FONT_SIZE, TRUE, 45);
        for (int i = 8; i <= 72; i += (i < 24 ? 2 : 6))
        {
                if (i == 22)
                        continue;
                CString strSize; strSize.Format(_T("%d"), i);
                pFontSizeComboBox->AddItem(strSize);
        }
        pFontPanel->Add(pFontSizeComboBox);
        pFontPanel->Add(new CMFCRibbonButton(ID_FONT_GROW, _T("A+")));
        pFontPanel->Add(new CMFCRibbonButton(ID_FONT_SHRINK, _T("A-")));
        pFontPanel->Add(new CMFCRibbonSeparator(TRUE));

        pFontPanel->Add(new CMFCRibbonButton(ID_FONT_BOLD, _T("B")));
        pFontPanel->Add(new CMFCRibbonButton(ID_FONT_ITALIC, _T("I")));
        pFontPanel->Add(new CMFCRibbonButton(ID_FONT_UNDERLINE, _T("U")));
        pFontPanel->Add(new CMFCRibbonButton(ID_FONT_STRIKETHROUGH, _T("S")));

        pFontPanel->Add(new CMFCRibbonSeparator(TRUE));
        pFontPanel->Add(new CMFCRibbonColorButton(ID_FONT_COLOR, _T("Font Color"), TRUE));
        pFontPanel->Add(new CMFCRibbonColorButton(ID_FONT_CELL_BKGND, _T("Fill"), TRUE));
}

