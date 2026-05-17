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

// Modeless CFindReplaceDialog posts this registered message on each action.
static const UINT WM_FINDREPLACE_REG = ::RegisterWindowMessage(FINDMSGSTRING);

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
        ON_UPDATE_COMMAND_UI_RANGE(ID_ALIGN_LEFT, ID_ALIGN_MERGE, &CMainFrame::OnUpdateAlignCmd)
        ON_UPDATE_COMMAND_UI_RANGE(ID_FMT_GENERAL, ID_FMT_DATE, &CMainFrame::OnUpdateFormatCmd)
        ON_UPDATE_COMMAND_UI(ID_FONT_BOLD,          &CMainFrame::OnUpdateFontBold)
        ON_UPDATE_COMMAND_UI(ID_FONT_ITALIC,        &CMainFrame::OnUpdateFontItalic)
        ON_UPDATE_COMMAND_UI(ID_FONT_UNDERLINE,     &CMainFrame::OnUpdateFontUnderline)
        ON_UPDATE_COMMAND_UI(ID_FONT_STRIKETHROUGH, &CMainFrame::OnUpdateFontStrikethrough)
        ON_COMMAND(ID_EDIT_FIND,       &CMainFrame::OnEditFind)
        ON_COMMAND(ID_EDIT_REPLACE,    &CMainFrame::OnEditReplace)
        ON_COMMAND(ID_EDIT_SELECT_ALL, &CMainFrame::OnEditSelectAll)
        ON_COMMAND(ID_EDIT_CLEAR,      &CMainFrame::OnEditClear)
        ON_COMMAND(ID_FORMAT_PAINTER,  &CMainFrame::OnFormatPainter)
        ON_UPDATE_COMMAND_UI(ID_FORMAT_PAINTER, &CMainFrame::OnUpdateFormatPainter)
        ON_REGISTERED_MESSAGE(WM_FINDREPLACE_REG, &CMainFrame::OnFindReplaceMsg)
        ON_COMMAND(ID_CELL_FONT,       &CMainFrame::OnCellFont)
        ON_COMMAND(ID_DLG_ALIGNMENT,   &CMainFrame::OnDlgAlignment)
        ON_COMMAND(ID_DLG_NUMBER,      &CMainFrame::OnDlgNumber)
        ON_COMMAND(ID_DLG_INSERT,      &CMainFrame::OnDlgInsert)
        ON_COMMAND(ID_DLG_SORT,        &CMainFrame::OnDlgSort)
        ON_COMMAND(ID_DLG_FILL,        &CMainFrame::OnDlgFill)
        ON_COMMAND_RANGE(ID_SORT_ASC,  ID_SORT_DESC, &CMainFrame::OnSortCmd)
        ON_COMMAND_RANGE(ID_FILL_DOWN, ID_FILL_LEFT, &CMainFrame::OnFillCmd)
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
	SetupInsertPanel();
	SetupClipboardPanel();
	SetupSortFillPanel();

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
//        eraser(33) paint_brush(34)
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
        constexpr int Eraser      = 33;   // Fluent System Icons: ic_fluent_eraser_segment_*_filled
        constexpr int PaintBrush  = 34;   // Fluent System Icons: ic_fluent_paint_brush_*_filled
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

// Helper: ribbon button with tooltip + description in one call.
static CMFCRibbonButton* MakeBtn(UINT id, LPCTSTR text, int idx, LPCTSTR tip, LPCTSTR desc)
{
        auto* b = new CMFCRibbonButton(id, text, idx);
        b->SetToolTipText(tip);
        b->SetDescription(desc);
        return b;
}

void CMainFrame::SetupFontPanel()
{
        CMFCRibbonPanel* pFontPanel = FindHomePanel(m_wndRibbonBar, _T("Font"));
        if (!pFontPanel)
                return;

        pFontPanel->RemoveAll();

        CMFCRibbonButtonsGroup* pTopRow = new CMFCRibbonButtonsGroup();

        CMFCRibbonFontComboBox* pFontComboBox = new CMFCRibbonFontComboBox(ID_FONT_NAME,
                DEVICE_FONTTYPE | RASTER_FONTTYPE | TRUETYPE_FONTTYPE, DEFAULT_CHARSET, DEFAULT_PITCH, 120);
        pFontComboBox->SetToolTipText(_T("Font"));
        pFontComboBox->SetDescription(_T("Pick a font for the selected cells."));
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
        pFontSizeComboBox->SetToolTipText(_T("Font size"));
        pFontSizeComboBox->SetDescription(_T("Change the size of the selected cells' text."));
        pTopRow->AddButton(pFontSizeComboBox);

        pTopRow->AddButton(MakeBtn(ID_FONT_GROW,   _T("Grow font"),   FluentIdx::FontInc,
                _T("Grow font"),   _T("Increase the font size of the selected cells.")));
        pTopRow->AddButton(MakeBtn(ID_FONT_SHRINK, _T("Shrink font"), FluentIdx::FontDec,
                _T("Shrink font"), _T("Decrease the font size of the selected cells.")));

        pFontPanel->Add(pTopRow);

        CMFCRibbonButtonsGroup* pBottomRow = new CMFCRibbonButtonsGroup();

        pBottomRow->AddButton(MakeBtn(ID_FONT_BOLD,          _T("Bold"),          FluentIdx::Bold,
                _T("Bold (Ctrl+B)"),         _T("Make the selected text bold.")));
        pBottomRow->AddButton(MakeBtn(ID_FONT_ITALIC,        _T("Italic"),        FluentIdx::Italic,
                _T("Italic"),                _T("Italicize the selected text.")));
        pBottomRow->AddButton(MakeBtn(ID_FONT_UNDERLINE,     _T("Underline"),     FluentIdx::Underline,
                _T("Underline"),             _T("Underline the selected text.")));
        pBottomRow->AddButton(MakeBtn(ID_FONT_STRIKETHROUGH, _T("Strikethrough"), FluentIdx::Strike,
                _T("Strikethrough"),         _T("Draw a line through the selected text.")));
        pBottomRow->AddButton(MakeBtn(ID_FONT_SUBSCRIPT,     _T("Subscript"),     FluentIdx::Subscript,
                _T("Subscript"),             _T("Type small letters below the line of text.")));
        pBottomRow->AddButton(MakeBtn(ID_FONT_SUPERSCRIPT,   _T("Superscript"),   FluentIdx::Superscript,
                _T("Superscript"),           _T("Type small letters above the line of text.")));

        CMFCRibbonColorButton* pHighlight = new CMFCRibbonColorButton(ID_FONT_CELL_BKGND, _T("Highlight"), TRUE, FluentIdx::Highlight);
        pHighlight->EnableAutomaticButton(_T("No Color"), (COLORREF)-1);
        pHighlight->EnableOtherButton(_T("More Colors..."));
        pHighlight->SetColumns(10);
        pHighlight->SetColor(RGB(255, 255, 0));
        pHighlight->SetDefaultCommand(FALSE);
        pHighlight->SetToolTipText(_T("Highlight color"));
        pHighlight->SetDescription(_T("Make text look like it was marked with a highlighter."));
        pBottomRow->AddButton(pHighlight);

        CMFCRibbonColorButton* pFontColor = new CMFCRibbonColorButton(ID_FONT_COLOR, _T("Font color"), TRUE, FluentIdx::FontColor);
        pFontColor->EnableAutomaticButton(_T("Automatic"), RGB(0, 0, 0));
        pFontColor->EnableOtherButton(_T("More Colors..."));
        pFontColor->SetColumns(10);
        pFontColor->SetColor(RGB(255, 0, 0));
        pFontColor->SetDefaultCommand(FALSE);
        pFontColor->SetToolTipText(_T("Font color"));
        pFontColor->SetDescription(_T("Change the color of the text."));
        pBottomRow->AddButton(pFontColor);

        pFontPanel->Add(pBottomRow);
        pFontPanel->EnableLaunchButton(ID_CELL_FONT);
}

void CMainFrame::SetupAlignmentPanel()
{
        CMFCRibbonPanel* pPanel = FindHomePanel(m_wndRibbonBar, _T("Alignment"));
        if (!pPanel)
                return;

        pPanel->RemoveAll();

        CMFCRibbonButtonsGroup* pTopRow = new CMFCRibbonButtonsGroup();
        pTopRow->AddButton(MakeBtn(ID_ALIGN_LEFT,   _T("Align left"),   FluentIdx::AlignLeft,
                _T("Align left"),   _T("Align text to the left edge of the cell.")));
        pTopRow->AddButton(MakeBtn(ID_ALIGN_CENTER, _T("Center"),       FluentIdx::AlignCenter,
                _T("Center"),       _T("Horizontally center text within the cell.")));
        pTopRow->AddButton(MakeBtn(ID_ALIGN_RIGHT,  _T("Align right"),  FluentIdx::AlignRight,
                _T("Align right"),  _T("Align text to the right edge of the cell.")));
        pTopRow->AddButton(MakeBtn(ID_ALIGN_WRAP,   _T("Wrap text"),    FluentIdx::TextWrap,
                _T("Wrap text"),    _T("Make text visible across multiple lines within a cell.")));
        pPanel->Add(pTopRow);

        CMFCRibbonButtonsGroup* pBottomRow = new CMFCRibbonButtonsGroup();
        pBottomRow->AddButton(MakeBtn(ID_ALIGN_TOP,    _T("Align top"),    FluentIdx::AlignTop,
                _T("Align top"),    _T("Align text to the top of the cell.")));
        pBottomRow->AddButton(MakeBtn(ID_ALIGN_MIDDLE, _T("Align middle"), FluentIdx::AlignMiddle,
                _T("Align middle"), _T("Vertically center text within the cell.")));
        pBottomRow->AddButton(MakeBtn(ID_ALIGN_BOTTOM, _T("Align bottom"), FluentIdx::AlignBottom,
                _T("Align bottom"), _T("Align text to the bottom of the cell.")));
        pBottomRow->AddButton(MakeBtn(ID_ALIGN_MERGE,  _T("Merge cells"),  FluentIdx::Merge,
                _T("Merge cells"),  _T("Combine the selected cells into one larger cell.")));
        pPanel->Add(pBottomRow);
        pPanel->EnableLaunchButton(ID_DLG_ALIGNMENT);
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

        CMFCRibbonButtonsGroup* pTopRow = new CMFCRibbonButtonsGroup();
        pTopRow->AddButton(MakeBtn(ID_FMT_GENERAL, _T("General"), FluentIdx::FmtGeneral,
                _T("General"), _T("No specific number format applied to the cell.")));
        pTopRow->AddButton(MakeBtn(ID_FMT_NUMBER,  _T("Number"),  FluentIdx::FmtNumber,
                _T("Number"),  _T("Display values as numbers with thousands separators and decimals.")));
        pPanel->Add(pTopRow);

        CMFCRibbonButtonsGroup* pBottomRow = new CMFCRibbonButtonsGroup();
        pBottomRow->AddButton(MakeBtn(ID_FMT_CURRENCY, _T("Currency"), FluentIdx::FmtCurrency,
                _T("Currency"), _T("Display values as monetary amounts with a currency symbol.")));
        pBottomRow->AddButton(MakeBtn(ID_FMT_PERCENT,  _T("Percent"),  FluentIdx::FmtPercent,
                _T("Percent"),  _T("Display values as percentages, e.g. 0.25 as 25%.")));
        pBottomRow->AddButton(MakeBtn(ID_FMT_DATE,     _T("Date"),     FluentIdx::FmtDate,
                _T("Date"),     _T("Display values as dates in ISO format.")));
        pPanel->Add(pBottomRow);
        pPanel->EnableLaunchButton(ID_DLG_NUMBER);
}

void CMainFrame::SetupInsertPanel()
{
        CMFCRibbonPanel* pPanel = FindHomePanel(m_wndRibbonBar, _T("Insert"));
        if (!pPanel) return;
        pPanel->EnableLaunchButton(ID_DLG_INSERT);
}

void CMainFrame::SetupClipboardPanel()
{
        CMFCRibbonPanel* pPanel = FindHomePanel(m_wndRibbonBar, _T("Clipboard"));
        if (!pPanel) return;
        pPanel->Add(MakeBtn(ID_FORMAT_PAINTER, _T("Format Painter"), FluentIdx::PaintBrush,
                _T("Format Painter"),
                _T("Copy formatting from one cell and apply it to another by clicking the target.")));
        pPanel->Add(MakeBtn(ID_EDIT_CLEAR, _T("Clear"), FluentIdx::Eraser,
                _T("Clear"), _T("Delete the contents of the selected cells.")));
}

void CMainFrame::SetupSortFillPanel()
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
        if (!pCategory) return;

        // Skip if a previous run already added these panels (ribbon re-layout).
        for (int i = 0; i < pCategory->GetPanelCount(); ++i)
                if (0 == _tcsicmp(pCategory->GetPanel(i)->GetName(), _T("Sort")))
                        return;

        // Sort panel — one launcher arrow, opens the (forthcoming) multi-key Sort dialog.
        CMFCRibbonPanel* pSort = pCategory->AddPanel(_T("Sort"));
        // Use \u2192 (RIGHTWARDS ARROW) explicitly. Literal UTF-8 bytes in source
        // got mis-rendered as "Aâ†Z" when MSVC read the file as Windows-1252.
        pSort->Add(MakeBtn(ID_SORT_ASC,  L"Sort A\u2192Z", -1,
                _T("Sort ascending"),  _T("Sort the current column from smallest to largest.")));
        pSort->Add(MakeBtn(ID_SORT_DESC, L"Sort Z\u2192A", -1,
                _T("Sort descending"), _T("Sort the current column from largest to smallest.")));
        pSort->EnableLaunchButton(ID_DLG_SORT);

        // Fill panel — separate launcher for the (forthcoming) Fill Series dialog.
        CMFCRibbonPanel* pFill = pCategory->AddPanel(_T("Fill"));
        CMFCRibbonButtonsGroup* pTopRow = new CMFCRibbonButtonsGroup();
        pTopRow->AddButton(MakeBtn(ID_FILL_DOWN,  _T("Fill down"),  FluentIdx::AlignBottom,
                _T("Fill down"),  _T("Copy the topmost cell of the selection into every cell below it.")));
        pTopRow->AddButton(MakeBtn(ID_FILL_RIGHT, _T("Fill right"), FluentIdx::AlignRight,
                _T("Fill right"), _T("Copy the leftmost cell of the selection into every cell to its right.")));
        pFill->Add(pTopRow);

        CMFCRibbonButtonsGroup* pBottomRow = new CMFCRibbonButtonsGroup();
        pBottomRow->AddButton(MakeBtn(ID_FILL_UP,    _T("Fill up"),    FluentIdx::AlignTop,
                _T("Fill up"),    _T("Copy the bottommost cell of the selection into every cell above it.")));
        pBottomRow->AddButton(MakeBtn(ID_FILL_LEFT,  _T("Fill left"),  FluentIdx::AlignLeft,
                _T("Fill left"),  _T("Copy the rightmost cell of the selection into every cell to its left.")));
        pFill->Add(pBottomRow);
        pFill->EnableLaunchButton(ID_DLG_FILL);
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

void CMainFrame::OnUpdateAlignCmd(CCmdUI* pCmdUI)
{
        CGridCtrl* g = ActiveGrid();
        if (!g) { pCmdUI->Enable(FALSE); return; }
        switch (pCmdUI->m_nID)
        {
        case ID_ALIGN_LEFT:   pCmdUI->SetCheck(g->GetCurrentCellHAlign() == DT_LEFT);    break;
        case ID_ALIGN_CENTER: pCmdUI->SetCheck(g->GetCurrentCellHAlign() == DT_CENTER);  break;
        case ID_ALIGN_RIGHT:  pCmdUI->SetCheck(g->GetCurrentCellHAlign() == DT_RIGHT);   break;
        case ID_ALIGN_TOP:    pCmdUI->SetCheck(g->GetCurrentCellVAlign() == DT_TOP);     break;
        case ID_ALIGN_MIDDLE: pCmdUI->SetCheck(g->GetCurrentCellVAlign() == DT_VCENTER); break;
        case ID_ALIGN_BOTTOM: pCmdUI->SetCheck(g->GetCurrentCellVAlign() == DT_BOTTOM);  break;
        case ID_ALIGN_WRAP:   pCmdUI->SetCheck(g->GetCurrentCellWrap() ? 1 : 0);         break;
        case ID_ALIGN_MERGE:  pCmdUI->Enable(FALSE); break;  // roadmapped — Grid32 merge not yet implemented
        }
}

void CMainFrame::OnUpdateFormatCmd(CCmdUI* pCmdUI)
{
        CGridCtrl* g = ActiveGrid();
        if (!g) { pCmdUI->Enable(FALSE); return; }
        const UINT cur = g->GetCurrentCellFormat();
        UINT want = FMT_GENERAL;
        switch (pCmdUI->m_nID)
        {
        case ID_FMT_GENERAL:  want = FMT_GENERAL;  break;
        case ID_FMT_NUMBER:   want = FMT_NUMBER;   break;
        case ID_FMT_CURRENCY: want = FMT_CURRENCY; break;
        case ID_FMT_PERCENT:  want = FMT_PERCENT;  break;
        case ID_FMT_DATE:     want = FMT_DATE_ISO; break;
        default: return;
        }
        pCmdUI->SetCheck(cur == want);
}

// Reflect Bold/Italic/Underline/Strikethrough state of the current cell back
// onto the ribbon buttons. Single helper because each button only differs by
// which FONTINFO field it inspects.
static void UpdateFontToggle(CCmdUI* pCmdUI, CGridCtrl* g, bool (*pred)(const FONTINFO&))
{
        if (!g) { pCmdUI->Enable(FALSE); return; }
        FONTINFO fi{};
        if (!g->GetCurrentCellFontInfo(fi)) { pCmdUI->SetCheck(0); return; }
        pCmdUI->SetCheck(pred(fi) ? 1 : 0);
}

void CMainFrame::OnUpdateFontBold(CCmdUI* pCmdUI)
{
        UpdateFontToggle(pCmdUI, ActiveGrid(), [](const FONTINFO& fi) { return fi.bWeight >= 700; });
}

void CMainFrame::OnUpdateFontItalic(CCmdUI* pCmdUI)
{
        UpdateFontToggle(pCmdUI, ActiveGrid(), [](const FONTINFO& fi) { return fi.bItalic != FALSE; });
}

void CMainFrame::OnUpdateFontUnderline(CCmdUI* pCmdUI)
{
        UpdateFontToggle(pCmdUI, ActiveGrid(), [](const FONTINFO& fi) { return fi.bUnderline != FALSE; });
}

void CMainFrame::OnUpdateFontStrikethrough(CCmdUI* pCmdUI)
{
        UpdateFontToggle(pCmdUI, ActiveGrid(), [](const FONTINFO& fi) { return fi.bStrikeThrough != FALSE; });
}

void CMainFrame::OnEditFind()
{
        if (!ActiveGrid()) return;
        if (m_pFindReplace) { m_pFindReplace->SetFocus(); return; }
        m_pFindReplace = new CFindReplaceDialog;
        m_pFindReplace->Create(TRUE, nullptr, nullptr, FR_DOWN, this);
}

void CMainFrame::OnEditReplace()
{
        if (!ActiveGrid()) return;
        if (m_pFindReplace) { m_pFindReplace->SetFocus(); return; }
        m_pFindReplace = new CFindReplaceDialog;
        m_pFindReplace->Create(FALSE, nullptr, nullptr, FR_DOWN, this);
}

void CMainFrame::OnEditSelectAll()
{
        if (CGridCtrl* g = ActiveGrid()) g->SelectAll();
}

void CMainFrame::OnFormatPainter()
{
        // Toggle: a second click while armed disarms without doing anything,
        // matching Excel's behavior.
        if (m_bFormatPainterArmed)
        {
                m_bFormatPainterArmed = FALSE;
                return;
        }
        CGridCtrl* g = ActiveGrid();
        if (!g) return;
        if (!g->GetCurrentCellFontInfo(m_capturedFont)) return;
        m_bFormatPainterArmed = TRUE;
}

void CMainFrame::OnUpdateFormatPainter(CCmdUI* pCmdUI)
{
        pCmdUI->Enable(ActiveGrid() != nullptr);
        pCmdUI->SetCheck(m_bFormatPainterArmed ? 1 : 0);
}

bool CMainFrame::ConsumeGridClick(CGridCtrl* g, const GRIDPOINT& /*cell*/)
{
        if (!m_bFormatPainterArmed || !g)
                return false;
        // Apply the captured font to the cell (or drag-selected range) that
        // the user just clicked on. NM_CLICK fires after OnLButtonUp finalizes
        // the selection, so SetSelectionFormat covers single-click and drag.
        g->SetSelectionFormat(m_capturedFont);
        m_bFormatPainterArmed = FALSE;
        return true;
}

void CMainFrame::OnEditClear()
{
        CGridCtrl* g = ActiveGrid();
        if (!g) return;
        // Clear cell text across the current selection via GM_FILLCELLS with an
        // empty source. FILL_TEXT alone preserves cell formatting (matches the
        // Excel Delete-key behavior); a future "Clear All" can OR in more flags.
        GCFILLSTRUCT fs{};
        fs.m_cbSize          = sizeof(GCFILLSTRUCT);
        fs.m_dwFillHow       = FILLBYTEXT;
        fs.m_dwFillWhat      = FILL_TEXT;
        fs.m_dwFillRange     = FILLBYCURRENTSELECTION;
        fs.m_wsText          = L"";
        ::SendMessage(g->GetSafeHwnd(), GM_FILLCELLS, 0, (LPARAM)&fs);
}

void CMainFrame::OnCellFont()
{
        CGridCtrl* g = ActiveGrid();
        if (!g) return;

        FONTINFO fi{};
        if (!g->GetCurrentCellFontInfo(fi))
        {
                fi.m_wsFontFace = L"Calibri";
                fi.m_fPointSize = 11.0f;
                fi.bWeight      = FW_NORMAL;
        }

        LOGFONT lf{};
        wcscpy_s(lf.lfFaceName, LF_FACESIZE, fi.m_wsFontFace.c_str());
        HDC hdc = ::GetDC(nullptr);
        lf.lfHeight    = -MulDiv((int)(fi.m_fPointSize + 0.5f), GetDeviceCaps(hdc, LOGPIXELSY), 72);
        ::ReleaseDC(nullptr, hdc);
        lf.lfWeight    = (LONG)(fi.bWeight ? fi.bWeight : FW_NORMAL);
        lf.lfItalic    = fi.bItalic        ? 1 : 0;
        lf.lfUnderline = fi.bUnderline     ? 1 : 0;
        lf.lfStrikeOut = fi.bStrikeThrough ? 1 : 0;
        lf.lfCharSet   = DEFAULT_CHARSET;

        CFontDialog dlg(&lf, CF_EFFECTS | CF_SCREENFONTS | CF_INITTOLOGFONTSTRUCT, nullptr, this);
        dlg.m_cf.rgbColors = fi.m_clrTextColor;
        if (dlg.DoModal() != IDOK)
                return;

        LOGFONT out{};
        dlg.GetCurrentFont(&out);
        fi.m_wsFontFace    = out.lfFaceName;
        // CHOOSEFONT.iPointSize is in tenths of a point.
        fi.m_fPointSize    = dlg.m_cf.iPointSize / 10.0f;
        fi.bWeight         = (UINT)out.lfWeight;
        fi.bItalic         = out.lfItalic    != 0;
        fi.bUnderline      = out.lfUnderline != 0;
        fi.bStrikeThrough  = out.lfStrikeOut != 0;
        fi.m_clrTextColor  = dlg.GetColor();
        g->SetSelectionFormat(fi);
}

void CMainFrame::OnDlgAlignment()
{
        AfxMessageBox(_T("A full Cell Alignment dialog is coming soon. ")
                      _T("For now, use the Alignment panel buttons."),
                      MB_OK | MB_ICONINFORMATION);
}

void CMainFrame::OnDlgNumber()
{
        AfxMessageBox(_T("A full Number Format dialog is coming soon. ")
                      _T("For now, use the Number panel buttons."),
                      MB_OK | MB_ICONINFORMATION);
}

void CMainFrame::OnDlgInsert()
{
        AfxMessageBox(_T("A full Insert dialog is coming soon. ")
                      _T("For now, use the Insert panel buttons."),
                      MB_OK | MB_ICONINFORMATION);
}

void CMainFrame::OnDlgSort()
{
        AfxMessageBox(_T("A full multi-key Sort dialog is coming soon. ")
                      _T("For now, use the Sort panel buttons."),
                      MB_OK | MB_ICONINFORMATION);
}

void CMainFrame::OnDlgFill()
{
        AfxMessageBox(_T("A full Fill (Series / Justify) dialog is coming soon. ")
                      _T("For now, use the Fill panel buttons."),
                      MB_OK | MB_ICONINFORMATION);
}

void CMainFrame::OnSortCmd(UINT nID)
{
        CGridCtrl* g = ActiveGrid();
        if (!g) return;
        GRIDSELECTION sel{};
        ::SendMessage(g->GetSafeHwnd(), GM_GETSELECTION, 0, (LPARAM)&sel);
        GCSORTSTRUCT ss{};
        ss.m_cbSize     = sizeof(GCSORTSTRUCT);
        ss.m_sortSel    = sel;
        ss.m_bAscending = (nID == ID_SORT_ASC);
        ::SendMessage(g->GetSafeHwnd(), GM_SORTCELLS, 0, (LPARAM)&ss);
}

void CMainFrame::OnFillCmd(UINT nID)
{
        CGridCtrl* g = ActiveGrid();
        if (!g) return;
        GRIDSELECTION sel{};
        ::SendMessage(g->GetSafeHwnd(), GM_GETSELECTION, 0, (LPARAM)&sel);

        // Pick the source corner of the selection. OnFillCells doesn't use the
        // direction flag itself; we encode direction here by choosing refCell.
        GRIDPOINT refCell{};
        switch (nID)
        {
        case ID_FILL_DOWN:  refCell = sel.start; break;                                // top-left
        case ID_FILL_RIGHT: refCell = sel.start; break;                                // top-left
        case ID_FILL_UP:    refCell = { sel.end.nRow, sel.start.nCol }; break;         // bottom-left
        case ID_FILL_LEFT:  refCell = { sel.start.nRow, sel.end.nCol }; break;         // top-right
        default: return;
        }

        GCFILLSTRUCT fs{};
        fs.m_cbSize        = sizeof(GCFILLSTRUCT);
        fs.m_dwFillHow     = FILLBYREFCELL;
        fs.m_dwFillWhat    = FILL_TEXT | FILL_COLOR | FILL_FONT | FILL_POINTSIZE | FILL_EFFECTS;
        fs.m_dwFillRange   = FILLBYCURRENTSELECTION;
        fs.m_refCell       = refCell;
        ::SendMessage(g->GetSafeHwnd(), GM_FILLCELLS, 0, (LPARAM)&fs);
}

LRESULT CMainFrame::OnFindReplaceMsg(WPARAM, LPARAM lParam)
{
        CFindReplaceDialog* dlg = CFindReplaceDialog::GetNotifier(lParam);
        if (!dlg) return 0;
        if (dlg->IsTerminating())
        {
                m_pFindReplace = nullptr;   // dialog self-deletes via PostNcDestroy
                return 0;
        }
        CGridCtrl* g = ActiveGrid();
        if (!g) return 0;

        // Start the next search just past the current cell so repeated Find Next
        // moves forward instead of re-matching the cell we're already on.
        GRIDPOINT cur{};
        g->GetCurrentCell(cur);
        GRIDPOINT next = cur;
        if (dlg->SearchDown()) { next.nCol += 1; }
        else                    { if (next.nCol == 0) { if (next.nRow) { --next.nRow; next.nCol = 0xFFFF; } } else --next.nCol; }

        CString findStr = dlg->GetFindString();
        CT2W wFind(findStr);

        if (dlg->FindNext())
        {
                GCFINDSTRUCT fs{};
                fs.m_cbSize = sizeof(fs);
                fs.m_wsFindText = (LPCWSTR)wFind;
                fs.m_startCell = next;
                fs.m_bMatchCase = dlg->MatchCase();
                fs.m_bSearchForward = dlg->SearchDown();
                if (!g->FindText(fs))
                        AfxMessageBox(_T("No more matches."), MB_OK | MB_ICONINFORMATION);
                return 0;
        }

        if (dlg->ReplaceCurrent() || dlg->ReplaceAll())
        {
                CT2W wRepl(dlg->GetReplaceString());
                GCREPLACESTRUCT rs{};
                rs.m_cbSize = sizeof(rs);
                rs.m_wsFindText = (LPCWSTR)wFind;
                rs.m_wsReplaceText = (LPCWSTR)wRepl;
                rs.m_bMatchCase = dlg->MatchCase();
                rs.m_bSearchForward = TRUE;

                if (dlg->ReplaceCurrent())
                {
                        rs.m_startCell = cur;       // operate on the current match, don't skip past it
                        g->ReplaceText(rs);
                        return 0;
                }

                rs.m_startCell = { 0, 0 };
                int count = 0;
                while (g->ReplaceText(rs))
                {
                        ++count;
                        g->GetCurrentCell(rs.m_startCell);
                        rs.m_startCell.nCol += 1;   // advance past last replacement
                }
                CString msg; msg.Format(_T("%d replacement(s) made."), count);
                AfxMessageBox(msg, MB_OK | MB_ICONINFORMATION);
        }
        return 0;
}

