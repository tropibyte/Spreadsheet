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
#include "GotoDlg.h"
#include <atlconv.h>

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
        ON_COMMAND(ID_FONT_SUBSCRIPT, &CMainFrame::OnFontSubscript)
        ON_COMMAND(ID_FONT_SUPERSCRIPT, &CMainFrame::OnFontSuperscript)
        ON_COMMAND(ID_INSERT_PICTURE, &CMainFrame::OnInsertPicture)
        ON_COMMAND(ID_INSERT_DATETIME, &CMainFrame::OnInsertDateTime)
        ON_COMMAND_RANGE(ID_ALIGN_LEFT, ID_ALIGN_MERGE, &CMainFrame::OnAlignCmd)
        ON_COMMAND_RANGE(ID_FMT_GENERAL, ID_FMT_DATE, &CMainFrame::OnFormatCmd)
END_MESSAGE_MAP()

// CMainFrame construction/destruction

CMainFrame::CMainFrame() noexcept : m_pCurrOutlookPage(nullptr), m_pCurrOutlookWnd(nullptr)
{
        // TODO: add member initialization code here
        theApp.m_nAppLook = theApp.GetInt(_T("ApplicationLook"), ID_VIEW_APPLOOK_WINDOWS_7);
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
	SetupAlignmentPanel();
	SetupNumberPanel();

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

int CMainFrame::FindFocusedOutlookWnd(CMFCOutlookBarTabCtrl** ppOutlookWnd)
{
	return 0;
}

CMFCOutlookBarTabCtrl* CMainFrame::FindOutlookParent(CWnd* pWnd)
{
	return nullptr;
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


// Returns the active grid view's CGridCtrl, or nullptr if no view is active.
CGridCtrl* CMainFrame::ActiveGrid()
{
        CGridView* pView = DYNAMIC_DOWNCAST(CGridView, GetActiveView());
        return pView ? &pView->GetGridCtrl() : nullptr;
}

// Read the current cell's font, apply a mutator, write it back to the selection.
// Centralizes the read-modify-write pattern for all Font-panel toggles.
template <typename Fn>
static void MutateCurrentFont(CGridCtrl* g, Fn mutator)
{
        if (!g) return;
        FONTINFO fi{};
        if (!g->GetCurrentCellFontInfo(fi)) return;
        mutator(fi);
        g->SetSelectionFormat(fi);
}

void CMainFrame::OnCellGoto()
{
        CGotoDlg dlg(this);
        if (dlg.DoModal() != IDOK)
                return;
        CGridCtrl* g = ActiveGrid();
        if (!g) return;
        CT2W wRef(dlg.m_cellRef);
        g->SetCurrentCell(wRef, dlg.nWhich);
}

void CMainFrame::OnFontName()
{
        CMFCRibbonFontComboBox* pCombo = DYNAMIC_DOWNCAST(CMFCRibbonFontComboBox, m_wndRibbonBar.FindByID(ID_FONT_NAME));
        if (!pCombo) return;
        CString face = pCombo->GetEditText();
        if (face.IsEmpty()) return;
        MutateCurrentFont(ActiveGrid(), [&](FONTINFO& fi) { fi.m_wsFontFace = (LPCWSTR)face; });
}

void CMainFrame::OnFontSize()
{
        CMFCRibbonComboBox* pCombo = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, m_wndRibbonBar.FindByID(ID_FONT_SIZE));
        if (!pCombo) return;
        int size = _ttoi(pCombo->GetEditText());
        if (size <= 0) return;
        MutateCurrentFont(ActiveGrid(), [&](FONTINFO& fi) { fi.m_fPointSize = (float)size; });
}

void CMainFrame::OnFontGrow()
{
        MutateCurrentFont(ActiveGrid(), [](FONTINFO& fi) {
                if (fi.m_fPointSize <= 0) fi.m_fPointSize = 10.0f;
                fi.m_fPointSize += 1.0f;
        });
}

void CMainFrame::OnFontShrink()
{
        MutateCurrentFont(ActiveGrid(), [](FONTINFO& fi) {
                if (fi.m_fPointSize > 1.0f) fi.m_fPointSize -= 1.0f;
        });
}

void CMainFrame::OnFontBold()
{
        MutateCurrentFont(ActiveGrid(), [](FONTINFO& fi) {
                fi.bWeight = (fi.bWeight >= 700) ? 400 : 700;
        });
}

void CMainFrame::OnFontItalic()
{
        MutateCurrentFont(ActiveGrid(), [](FONTINFO& fi) { fi.bItalic = !fi.bItalic; });
}

void CMainFrame::OnFontUnderline()
{
        MutateCurrentFont(ActiveGrid(), [](FONTINFO& fi) { fi.bUnderline = !fi.bUnderline; });
}

void CMainFrame::OnFontStrikethrough()
{
        MutateCurrentFont(ActiveGrid(), [](FONTINFO& fi) { fi.bStrikeThrough = !fi.bStrikeThrough; });
}

void CMainFrame::OnFontColor()
{
        CMFCRibbonColorButton* pBtn = DYNAMIC_DOWNCAST(CMFCRibbonColorButton, m_wndRibbonBar.FindByID(ID_FONT_COLOR));
        if (!pBtn) return;
        COLORREF c = pBtn->GetColor();
        MutateCurrentFont(ActiveGrid(), [&](FONTINFO& fi) { fi.m_clrTextColor = c; });
}

void CMainFrame::OnFontBackground()
{
        CMFCRibbonColorButton* pBtn = DYNAMIC_DOWNCAST(CMFCRibbonColorButton, m_wndRibbonBar.FindByID(ID_FONT_CELL_BKGND));
        if (!pBtn) return;
        m_cellBkg = pBtn->GetColor();
        // The cell-background channel doesn't go through FONTINFO — it lives
        // on cell_base::clrBackground. There's no GM_SETBACKGROUND yet, so
        // this stays parked on m_cellBkg until that lands.
}

void CMainFrame::ApplyCurrentFont()
{
        // Legacy entry point. Individual handlers now mutate font state
        // directly through MutateCurrentFont(); this remains for callers
        // that already hold a fully-built FONTINFO they want to apply.
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
// Fluent strip index map — keep in sync with res/fluent_small.bmp + fluent_large.bmp.
// Order: paste(0) cut(1) copy(2) select_all(3) bold(4) italic(5) underline(6) strike(7)
//        subscript(8) superscript(9) font_color(10) highlight(11) font_inc(12) font_dec(13)
//        align_left(14) align_center(15) align_right(16) align_top(17) align_middle(18)
//        align_bottom(19) text_wrap(20) merge(21) picture(22) date(23) object(24)
//        find(25) replace(26) goto(27)
//        fmt_general(28) fmt_number(29) fmt_currency(30) fmt_percent(31) fmt_date(32)
namespace FluentIdx {
        constexpr int Bold        = 4;
        constexpr int Italic      = 5;
        constexpr int Underline   = 6;
        constexpr int Strike      = 7;
        constexpr int Subscript   = 8;
        constexpr int Superscript = 9;
        constexpr int FontColor   = 10;
        constexpr int Highlight   = 11;
        constexpr int FontInc     = 12;
        constexpr int FontDec     = 13;
        constexpr int AlignLeft   = 14;
        constexpr int AlignCenter = 15;
        constexpr int AlignRight  = 16;
        constexpr int AlignTop    = 17;
        constexpr int AlignMiddle = 18;
        constexpr int AlignBottom = 19;
        constexpr int TextWrap    = 20;
        constexpr int Merge       = 21;
        constexpr int FmtGeneral  = 28;
        constexpr int FmtNumber   = 29;
        constexpr int FmtCurrency = 30;
        constexpr int FmtPercent  = 31;
        constexpr int FmtDate     = 32;
}

static CMFCRibbonPanel* FindHomePanel(CMFCRibbonBar& bar, LPCTSTR pszName)
{
        CMFCRibbonCategory* pCategory = nullptr;
        for (int i = 0; i < bar.GetCategoryCount(); ++i)
        {
                if (0 == _tcsicmp(bar.GetCategory(i)->GetName(), _T("Home")))
                {
                        pCategory = bar.GetCategory(i);
                        break;
                }
        }
        if (!pCategory)
                return nullptr;

        for (int i = 0; i < pCategory->GetPanelCount(); ++i)
        {
                if (0 == _tcsicmp(pCategory->GetPanel(i)->GetName(), pszName))
                        return pCategory->GetPanel(i);
        }
        return nullptr;
}

void CMainFrame::SetupFontPanel()
{
        CMFCRibbonPanel* pFontPanel = FindHomePanel(m_wndRibbonBar, _T("Font"));
        if (!pFontPanel)
                return;

        pFontPanel->RemoveAll();

        // Top row: Font name, Font size, A+, A-
        CMFCRibbonButtonsGroup* pTopRow = new CMFCRibbonButtonsGroup();

        CMFCRibbonFontComboBox* pFontComboBox = new CMFCRibbonFontComboBox(ID_FONT_NAME,
                DEVICE_FONTTYPE | RASTER_FONTTYPE | TRUETYPE_FONTTYPE, DEFAULT_CHARSET, DEFAULT_PITCH, 120);
        pTopRow->AddButton(pFontComboBox);

        CMFCRibbonComboBox* pFontSizeComboBox = new CMFCRibbonComboBox(ID_FONT_SIZE, TRUE, 40);
        int fontSizes[] = { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };
        for (int i = 0; i < _countof(fontSizes); ++i)
        {
                CString strSize;
                strSize.Format(_T("%d"), fontSizes[i]);
                pFontSizeComboBox->AddItem(strSize);
        }
        pFontSizeComboBox->SelectItem(4);
        pTopRow->AddButton(pFontSizeComboBox);

        pTopRow->AddButton(new CMFCRibbonButton(ID_FONT_GROW,   _T("Grow font"),   FluentIdx::FontInc));
        pTopRow->AddButton(new CMFCRibbonButton(ID_FONT_SHRINK, _T("Shrink font"), FluentIdx::FontDec));

        pFontPanel->Add(pTopRow);

        // Bottom row: B I U S x2 x^ + Highlight + Color
        CMFCRibbonButtonsGroup* pBottomRow = new CMFCRibbonButtonsGroup();

        pBottomRow->AddButton(new CMFCRibbonButton(ID_FONT_BOLD,          _T("Bold"),          FluentIdx::Bold));
        pBottomRow->AddButton(new CMFCRibbonButton(ID_FONT_ITALIC,        _T("Italic"),        FluentIdx::Italic));
        pBottomRow->AddButton(new CMFCRibbonButton(ID_FONT_UNDERLINE,     _T("Underline"),     FluentIdx::Underline));
        pBottomRow->AddButton(new CMFCRibbonButton(ID_FONT_STRIKETHROUGH, _T("Strikethrough"), FluentIdx::Strike));
        pBottomRow->AddButton(new CMFCRibbonButton(ID_FONT_SUBSCRIPT,     _T("Subscript"),     FluentIdx::Subscript));
        pBottomRow->AddButton(new CMFCRibbonButton(ID_FONT_SUPERSCRIPT,   _T("Superscript"),   FluentIdx::Superscript));

        CMFCRibbonColorButton* pHighlight = new CMFCRibbonColorButton(ID_FONT_CELL_BKGND, _T("Highlight"), TRUE, FluentIdx::Highlight);
        pHighlight->EnableAutomaticButton(_T("No Color"), (COLORREF)-1);
        pHighlight->EnableOtherButton(_T("More Colors..."));
        pHighlight->SetColumns(10);
        pHighlight->SetColor(RGB(255, 255, 0));
        pHighlight->SetDefaultCommand(FALSE);
        pBottomRow->AddButton(pHighlight);

        CMFCRibbonColorButton* pFontColor = new CMFCRibbonColorButton(ID_FONT_COLOR, _T("Font color"), TRUE, FluentIdx::FontColor);
        pFontColor->EnableAutomaticButton(_T("Automatic"), RGB(0, 0, 0));
        pFontColor->EnableOtherButton(_T("More Colors..."));
        pFontColor->SetColumns(10);
        pFontColor->SetColor(RGB(255, 0, 0));
        pFontColor->SetDefaultCommand(FALSE);
        pBottomRow->AddButton(pFontColor);

        pFontPanel->Add(pBottomRow);
}

void CMainFrame::SetupAlignmentPanel()
{
        CMFCRibbonPanel* pPanel = FindHomePanel(m_wndRibbonBar, _T("Alignment"));
        if (!pPanel)
                return;

        pPanel->RemoveAll();

        // Top row: horizontal alignment + wrap text
        CMFCRibbonButtonsGroup* pTopRow = new CMFCRibbonButtonsGroup();
        pTopRow->AddButton(new CMFCRibbonButton(ID_ALIGN_LEFT,   _T("Align left"),   FluentIdx::AlignLeft));
        pTopRow->AddButton(new CMFCRibbonButton(ID_ALIGN_CENTER, _T("Center"),       FluentIdx::AlignCenter));
        pTopRow->AddButton(new CMFCRibbonButton(ID_ALIGN_RIGHT,  _T("Align right"),  FluentIdx::AlignRight));
        pTopRow->AddButton(new CMFCRibbonButton(ID_ALIGN_WRAP,   _T("Wrap text"),    FluentIdx::TextWrap));
        pPanel->Add(pTopRow);

        // Bottom row: vertical alignment + merge cells
        CMFCRibbonButtonsGroup* pBottomRow = new CMFCRibbonButtonsGroup();
        pBottomRow->AddButton(new CMFCRibbonButton(ID_ALIGN_TOP,    _T("Align top"),    FluentIdx::AlignTop));
        pBottomRow->AddButton(new CMFCRibbonButton(ID_ALIGN_MIDDLE, _T("Align middle"), FluentIdx::AlignMiddle));
        pBottomRow->AddButton(new CMFCRibbonButton(ID_ALIGN_BOTTOM, _T("Align bottom"), FluentIdx::AlignBottom));
        pBottomRow->AddButton(new CMFCRibbonButton(ID_ALIGN_MERGE,  _T("Merge cells"),  FluentIdx::Merge));
        pPanel->Add(pBottomRow);
}

void CMainFrame::OnFontSubscript()
{
        // Subscript toggle — depends on a richer FONTINFO model in Grid32. Stubbed for now.
}

void CMainFrame::OnFontSuperscript()
{
        // Superscript toggle — depends on a richer FONTINFO model in Grid32. Stubbed for now.
}

void CMainFrame::OnInsertPicture()
{
        // Awaiting Grid32 picture-cell support.
        AfxMessageBox(_T("Picture insertion will be available once Grid32 supports image cells."));
}

void CMainFrame::OnInsertDateTime()
{
        CGridCtrl* g = ActiveGrid();
        if (!g) return;
        CTime now = CTime::GetCurrentTime();
        CString strDateTime = now.Format(_T("%Y-%m-%d %H:%M"));
        g->SetCurrentCellText((LPCWSTR)CT2W(strDateTime));
}

void CMainFrame::OnAlignCmd(UINT nID)
{
        CGridCtrl* g = ActiveGrid();
        if (!g) return;
        switch (nID)
        {
        case ID_ALIGN_LEFT:   g->SetSelectionHAlign(DT_LEFT);    break;
        case ID_ALIGN_CENTER: g->SetSelectionHAlign(DT_CENTER);  break;
        case ID_ALIGN_RIGHT:  g->SetSelectionHAlign(DT_RIGHT);   break;
        case ID_ALIGN_TOP:    g->SetSelectionVAlign(DT_TOP);     break;
        case ID_ALIGN_MIDDLE: g->SetSelectionVAlign(DT_VCENTER); break;
        case ID_ALIGN_BOTTOM: g->SetSelectionVAlign(DT_BOTTOM);  break;
        case ID_ALIGN_WRAP:
                g->SetSelectionWrap(!g->GetCurrentCellWrap());
                break;
        case ID_ALIGN_MERGE:
                // Awaiting Grid32 merge-cells implementation (mergeRange exists,
                // but no API or paint-time handling yet).
                break;
        }
}

void CMainFrame::SetupNumberPanel()
{
        CMFCRibbonPanel* pPanel = FindHomePanel(m_wndRibbonBar, _T("Number"));
        if (!pPanel)
                return;

        pPanel->RemoveAll();

        // Two rows. Top: General + Number. Bottom: Currency + Percent + Date.
        CMFCRibbonButtonsGroup* pTopRow = new CMFCRibbonButtonsGroup();
        pTopRow->AddButton(new CMFCRibbonButton(ID_FMT_GENERAL, _T("General"), FluentIdx::FmtGeneral));
        pTopRow->AddButton(new CMFCRibbonButton(ID_FMT_NUMBER,  _T("Number"),  FluentIdx::FmtNumber));
        pPanel->Add(pTopRow);

        CMFCRibbonButtonsGroup* pBottomRow = new CMFCRibbonButtonsGroup();
        pBottomRow->AddButton(new CMFCRibbonButton(ID_FMT_CURRENCY, _T("Currency"), FluentIdx::FmtCurrency));
        pBottomRow->AddButton(new CMFCRibbonButton(ID_FMT_PERCENT,  _T("Percent"),  FluentIdx::FmtPercent));
        pBottomRow->AddButton(new CMFCRibbonButton(ID_FMT_DATE,     _T("Date"),     FluentIdx::FmtDate));
        pPanel->Add(pBottomRow);
}

void CMainFrame::OnFormatCmd(UINT nID)
{
        CGridView* pView = DYNAMIC_DOWNCAST(CGridView, GetActiveView());
        if (!pView)
                return;
        UINT fmt = FMT_GENERAL;
        switch (nID) {
        case ID_FMT_GENERAL:  fmt = FMT_GENERAL;  break;
        case ID_FMT_NUMBER:   fmt = FMT_NUMBER;   break;
        case ID_FMT_CURRENCY: fmt = FMT_CURRENCY; break;
        case ID_FMT_PERCENT:  fmt = FMT_PERCENT;  break;
        case ID_FMT_DATE:     fmt = FMT_DATE_ISO; break;
        default: return;
        }
        // SCF_SELECTION on Grid32's side falls back to the current cell when
        // the selection rect is empty, so this works for single-cell edits too.
        pView->GetGridCtrl().SetSelectionNumberFormat(fmt);
}

