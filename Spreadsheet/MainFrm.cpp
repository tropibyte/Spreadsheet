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
#include "AlignmentDlg.h"
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
        ON_COMMAND(ID_FORMAT_PAINTER_LOCK, &CMainFrame::OnFormatPainterLock)
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

// Read the current cell's font, apply a mutator, and write back only the
// attribute the command changes (mask, GCFM_*). Writing the whole FONTINFO
// made Bold on an Arial-12 cell and a Times-20 cell turn both Arial 12 (H7):
// the active cell decides the new value, each cell keeps the rest of its font.
template <typename Fn>
static void MutateCurrentFont(CGridCtrl* g, DWORD mask, Fn mutator)
{
        if (!g) return;
        FONTINFO fi{};
        if (!g->GetCurrentCellFontInfo(fi)) return;
        mutator(fi);
        g->SetSelectionFormat(fi, mask);
}

void CMainFrame::OnCellGoto()
{
        CGotoDlg dlg(this);
        if (dlg.DoModal() != IDOK)
                return;
        CGridCtrl* g = ActiveGrid();
        if (!g) return;
        CT2W wRef(dlg.m_cellRef);
        if (!g->SetCurrentCell(wRef, dlg.nWhich))
        {
                CString msg;
                msg.Format(dlg.nWhich == SETBYNAME ? _T("No cell is named \"%s\".")
                                                   : _T("\"%s\" is not a cell on this sheet."),
                           (LPCTSTR)dlg.m_cellRef);
                AfxMessageBox(msg, MB_OK | MB_ICONINFORMATION);
        }
}

void CMainFrame::OnFontName()
{
        CMFCRibbonFontComboBox* pCombo = DYNAMIC_DOWNCAST(CMFCRibbonFontComboBox, m_wndRibbonBar.FindByID(ID_FONT_NAME));
        if (!pCombo) return;
        CString face = pCombo->GetEditText();
        if (face.IsEmpty()) return;
        MutateCurrentFont(ActiveGrid(), GCFM_FACE, [&](FONTINFO& fi) { wcsncpy_s(fi.m_szFontFace, (LPCWSTR)face, _TRUNCATE); });
}

void CMainFrame::OnFontSize()
{
        CMFCRibbonComboBox* pCombo = DYNAMIC_DOWNCAST(CMFCRibbonComboBox, m_wndRibbonBar.FindByID(ID_FONT_SIZE));
        if (!pCombo) return;
        int size = _ttoi(pCombo->GetEditText());
        if (size <= 0) return;
        MutateCurrentFont(ActiveGrid(), GCFM_SIZE, [&](FONTINFO& fi) { fi.m_fPointSize = (float)size; });
}

void CMainFrame::OnFontGrow()
{
        MutateCurrentFont(ActiveGrid(), GCFM_SIZE, [](FONTINFO& fi) {
                if (fi.m_fPointSize <= 0) fi.m_fPointSize = 10.0f;
                fi.m_fPointSize += 1.0f;
        });
}

void CMainFrame::OnFontShrink()
{
        MutateCurrentFont(ActiveGrid(), GCFM_SIZE, [](FONTINFO& fi) {
                if (fi.m_fPointSize > 1.0f) fi.m_fPointSize -= 1.0f;
        });
}

void CMainFrame::OnFontBold()
{
        MutateCurrentFont(ActiveGrid(), GCFM_BOLD, [](FONTINFO& fi) {
                fi.bWeight = (fi.bWeight >= 700) ? 400 : 700;
        });
}

void CMainFrame::OnFontItalic()
{
        MutateCurrentFont(ActiveGrid(), GCFM_ITALIC, [](FONTINFO& fi) { fi.bItalic = !fi.bItalic; });
}

void CMainFrame::OnFontUnderline()
{
        MutateCurrentFont(ActiveGrid(), GCFM_UNDERLINE, [](FONTINFO& fi) { fi.bUnderline = !fi.bUnderline; });
}

void CMainFrame::OnFontStrikethrough()
{
        MutateCurrentFont(ActiveGrid(), GCFM_STRIKE, [](FONTINFO& fi) { fi.bStrikeThrough = !fi.bStrikeThrough; });
}

void CMainFrame::OnFontColor()
{
        CMFCRibbonColorButton* pBtn = DYNAMIC_DOWNCAST(CMFCRibbonColorButton, m_wndRibbonBar.FindByID(ID_FONT_COLOR));
        if (!pBtn) return;
        COLORREF c = pBtn->GetColor();
        MutateCurrentFont(ActiveGrid(), GCFM_COLOR, [&](FONTINFO& fi) { fi.m_clrTextColor = c; });
}

void CMainFrame::OnFontBackground()
{
        CMFCRibbonColorButton* pBtn = DYNAMIC_DOWNCAST(CMFCRibbonColorButton, m_wndRibbonBar.FindByID(ID_FONT_CELL_BKGND));
        if (!pBtn) return;
        m_cellBkg = pBtn->GetColor();
        // The cell background lives on the cell, not in FONTINFO, so it has
        // its own message (H11).
        if (CGridCtrl* g = ActiveGrid())
                g->SetSelectionBackground(m_cellBkg);
}


// True for the grid's inline cell editor: an EDIT whose parent is a Grid32.
static bool IsGridCellEditor(HWND hWnd)
{
        TCHAR cls[32];
        if (!hWnd || !::GetClassName(hWnd, cls, _countof(cls)) || _tcsicmp(cls, _T("Edit")) != 0)
                return false;
        HWND hParent = ::GetParent(hWnd);
        return hParent && ::GetClassName(hParent, cls, _countof(cls)) &&
               _tcsicmp(cls, GRID_WNDCLASS_NAME) == 0;
}

// The editing keys the cell editor handles for its own text.
static bool IsTextEditingKey(const MSG* pMsg)
{
        if (pMsg->message != WM_KEYDOWN && pMsg->message != WM_SYSKEYDOWN)
                return false;
        const bool ctrl = (::GetKeyState(VK_CONTROL) & 0x8000) != 0;
        const bool shift = (::GetKeyState(VK_SHIFT) & 0x8000) != 0;
        const bool alt = (::GetKeyState(VK_MENU) & 0x8000) != 0;
        switch (pMsg->wParam)
        {
        case 'C': case 'X': case 'V': case 'Z': case 'Y': case 'A':
                return ctrl && !alt;
        case VK_INSERT:
                return ctrl || shift;
        case VK_DELETE:
                return shift;
        case VK_BACK:
                return alt;
        }
        return false;
}

BOOL CMainFrame::PreTranslateMessage(MSG* pMsg)
{
        // While a cell is being edited, Ctrl+C/X/V/Z and friends act on the
        // text in the editor, not on the grid's cells, so they must not be
        // turned into commands by the accelerator table.
        if (IsGridCellEditor(pMsg->hwnd) && IsTextEditingKey(pMsg))
                return FALSE;
        // CFrameWndEx, not CFrameWnd: the ribbon's KeyTips, Alt navigation and
        // tooltips live there, and it translates the accelerators too (M15).
        return CFrameWndEx::PreTranslateMessage(pMsg);
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
//        eraser(33) paint_brush(34) sort_up(35) sort_down(36)
//        fill_down(37) fill_right(38) fill_up(39) fill_left(40)
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
        constexpr int SortUp      = 35;   // Fluent: ic_fluent_arrow_sort_up_*_filled
        constexpr int SortDown    = 36;   // Fluent: ic_fluent_arrow_sort_down_*_filled
        constexpr int FillDown    = 37;   // Fluent: ic_fluent_arrow_down_*_filled
        constexpr int FillRight   = 38;
        constexpr int FillUp      = 39;
        constexpr int FillLeft    = 40;
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
// A ribbon element's OnLButtonDblClk is a no-op, so a double-click on a
// button produces only the first click's command. This one also posts
// ID_FORMAT_PAINTER_LOCK for the double-click, which makes the Format
// Painter persistent as in Excel (S9).
class CFormatPainterButton : public CMFCRibbonButton
{
public:
        CFormatPainterButton(UINT nID, LPCTSTR lpszText, int nSmallImageIndex)
                : CMFCRibbonButton(nID, lpszText, nSmallImageIndex) {}
        void OnLButtonDblClk(CPoint point) override
        {
                CMFCRibbonButton::OnLButtonDblClk(point);
                if (CWnd* pMain = AfxGetMainWnd())
                        pMain->PostMessage(WM_COMMAND, ID_FORMAT_PAINTER_LOCK);
        }
};

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
        // Match Grid32's default cell font (set in CGrid32Mgr ctor).
        pFontComboBox->SetEditText(_T("Arial"));
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
                // Toggle: unmerge the block under the active cell, otherwise
                // merge the selection (H10).
                if (g->IsCurrentCellMerged())
                        g->UnmergeSelection();
                else
                        g->MergeSelection();
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
        auto* pPainter = new CFormatPainterButton(ID_FORMAT_PAINTER, _T("Format Painter"), FluentIdx::PaintBrush);
        pPainter->SetToolTipText(_T("Format Painter"));
        pPainter->SetDescription(_T("Copy formatting from one cell and apply it to another by clicking the target. ")
                                 _T("Double-click to keep painting until you click it again."));
        pPanel->Add(pPainter);
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
        pSort->Add(MakeBtn(ID_SORT_ASC,  _T("Sort ascending"),  FluentIdx::SortUp,
                _T("Sort A\u2192Z"),  _T("Sort the current column from smallest to largest.")));
        pSort->Add(MakeBtn(ID_SORT_DESC, _T("Sort descending"), FluentIdx::SortDown,
                _T("Sort Z\u2192A"),  _T("Sort the current column from largest to smallest.")));
        pSort->EnableLaunchButton(ID_DLG_SORT);

        // Fill panel — separate launcher for the (forthcoming) Fill Series dialog.
        CMFCRibbonPanel* pFill = pCategory->AddPanel(_T("Fill"));
        CMFCRibbonButtonsGroup* pTopRow = new CMFCRibbonButtonsGroup();
        pTopRow->AddButton(MakeBtn(ID_FILL_DOWN,  _T("Fill down"),  FluentIdx::FillDown,
                _T("Fill down"),  _T("Copy the topmost cell of the selection into every cell below it.")));
        pTopRow->AddButton(MakeBtn(ID_FILL_RIGHT, _T("Fill right"), FluentIdx::FillRight,
                _T("Fill right"), _T("Copy the leftmost cell of the selection into every cell to its right.")));
        pFill->Add(pTopRow);

        CMFCRibbonButtonsGroup* pBottomRow = new CMFCRibbonButtonsGroup();
        pBottomRow->AddButton(MakeBtn(ID_FILL_UP,    _T("Fill up"),    FluentIdx::FillUp,
                _T("Fill up"),    _T("Copy the bottommost cell of the selection into every cell above it.")));
        pBottomRow->AddButton(MakeBtn(ID_FILL_LEFT,  _T("Fill left"),  FluentIdx::FillLeft,
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
        case ID_ALIGN_MERGE:  pCmdUI->SetCheck(g->IsCurrentCellMerged() ? 1 : 0); break;
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

// Excel semantics: a click arms the painter for one application, a click
// while armed disarms it, and a double-click (ID_FORMAT_PAINTER_LOCK, from
// CFormatPainterButton) keeps it armed until it is clicked again.
void CMainFrame::OnFormatPainter()
{
        if (m_bFormatPainterArmed)
        {
                m_bFormatPainterArmed = FALSE;
                m_bFormatPainterPersistent = FALSE;
                return;
        }
        CGridCtrl* g = ActiveGrid();
        if (!g) return;
        if (!g->GetCurrentCellFontInfo(m_capturedFont)) return;
        m_bFormatPainterArmed = TRUE;
        m_bFormatPainterPersistent = FALSE;
}

// The ribbon never re-sends a button's command for the second click of a
// double-click, so promoting by timing two ID_FORMAT_PAINTER commands could
// never happen (S9). The double-click arrives here instead, after the first
// click has already armed (or disarmed) the painter.
void CMainFrame::OnFormatPainterLock()
{
        CGridCtrl* g = ActiveGrid();
        if (!g) return;
        if (!m_bFormatPainterArmed && !g->GetCurrentCellFontInfo(m_capturedFont))
                return;
        m_bFormatPainterArmed = TRUE;
        m_bFormatPainterPersistent = TRUE;
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
        // Persistent mode (entered via double-click of the button) keeps the
        // painter armed across applications; single-shot disarms here.
        if (!m_bFormatPainterPersistent)
                m_bFormatPainterArmed = FALSE;
        return true;
}

void CMainFrame::OnEditClear()
{
        CGridCtrl* g = ActiveGrid();
        if (!g) return;
        // WM_CLEAR clears the contents of the selection, formulas included,
        // keeps the formatting, recalculates once and is one undo step. Going
        // through GM_FILLCELLS with empty text left formulas behind to come
        // back on the next recalc, created a cell for every empty position,
        // and could not be undone (S4).
        g->SendMessage(WM_CLEAR);
}

void CMainFrame::OnCellFont()
{
        CGridCtrl* g = ActiveGrid();
        if (!g) return;

        FONTINFO fi{};
        if (!g->GetCurrentCellFontInfo(fi))
        {
                wcscpy_s(fi.m_szFontFace, L"Calibri");
                fi.m_fPointSize = 11.0f;
                fi.bWeight      = FW_NORMAL;
        }

        LOGFONT lf{};
        wcscpy_s(lf.lfFaceName, LF_FACESIZE, fi.m_szFontFace);
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
        wcsncpy_s(fi.m_szFontFace, out.lfFaceName, _TRUNCATE);
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
        CGridCtrl* g = ActiveGrid();
        if (!g) return;

        // DT_LEFT and DT_TOP are both 0, so a 0 here is a real setting, not
        // "unset": treating it as unset turned Top into Middle (S8).
        CAlignmentDlg dlg(this);
        const UINT hAlign = g->GetCurrentCellHAlign();
        const UINT vAlign = g->GetCurrentCellVAlign();
        const BOOL bWrap  = g->GetCurrentCellWrap();
        dlg.m_hAlign = hAlign;
        dlg.m_vAlign = vAlign;
        dlg.m_bWrap  = bWrap;

        if (dlg.DoModal() != IDOK)
                return;

        // Write only what the user changed, so OK doesn't flatten a selection
        // whose cells are aligned differently.
        if (dlg.m_hAlign != hAlign)
                g->SetSelectionHAlign(dlg.m_hAlign);
        if (dlg.m_vAlign != vAlign)
                g->SetSelectionVAlign(dlg.m_vAlign);
        if (!dlg.m_bWrap != !bWrap)
                g->SetSelectionWrap(dlg.m_bWrap);
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
        g->GetSelection(sel);
        // A single cell means "sort the sheet by this column": the rows of
        // the used range, keyed on the active cell's column.
        if (sel.start.nRow == sel.end.nRow && sel.start.nCol == sel.end.nCol)
        {
                GRIDSELECTION used{};
                if (!g->GetUsedRange(used))
                        return;
                sel = used;
        }
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
        // The grid fills each column (down/up) or row (right/left) of the
        // selection from its own edge cell. Choosing one corner here made Fill
        // Down and Fill Right copy the same cell everywhere, and an upward
        // drag fill in reverse (S7).
        GCFILLSTRUCT fs{};
        fs.m_cbSize        = sizeof(GCFILLSTRUCT);
        fs.m_dwFillHow     = FILLBYREFCELL;
        fs.m_dwFillWhat    = FILL_ALL;
        fs.m_dwFillRange   = FILLBYCURRENTSELECTION;
        switch (nID)
        {
        case ID_FILL_DOWN:  fs.m_dwFillDirection = FILL_DOWN;  break;
        case ID_FILL_RIGHT: fs.m_dwFillDirection = FILL_RIGHT; break;
        case ID_FILL_UP:    fs.m_dwFillDirection = FILL_UP;    break;
        case ID_FILL_LEFT:  fs.m_dwFillDirection = FILL_LEFT;  break;
        default: return;
        }
        ::SendMessage(g->GetSafeHwnd(), GM_FILLCELLS, 0, (LPARAM)&fs);
}

// Find Next wraps around the sheet once. Replace changes every occurrence in
// the current match and moves on to the next one; it used to change only the
// first occurrence and stay put (S6). Replace All is one grid operation and
// one undo step: looping on GM_REPLACETEXT never finished when the
// replacement contained the search text (H5), and counted cells, not
// replacements (S6).
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

        CString findStr = dlg->GetFindString();
        CT2W wFind(findStr);
        const BOOL bDown = dlg->SearchDown();

        // Search from the cell after (or before) the active one, then from the
        // other end of the sheet.
        auto findNext = [&]() -> BOOL {
                GRIDPOINT cur{};
                g->GetCurrentCell(cur);
                GCFINDSTRUCT fs{};
                fs.m_cbSize = sizeof(fs);
                fs.m_pwszFindText = wFind;
                fs.m_bMatchCase = dlg->MatchCase();
                fs.m_bSearchForward = bDown;
                fs.m_startCell = cur;
                if (bDown)
                        fs.m_startCell.nCol += 1;
                else if (cur.nCol > 0)
                        fs.m_startCell.nCol -= 1;
                else if (cur.nRow > 0)
                        fs.m_startCell = { cur.nRow - 1, (UINT)-1 };
                else
                        fs.m_startCell = { (UINT)-1, (UINT)-1 };
                if ((bDown || cur.nRow > 0 || cur.nCol > 0) && g->FindText(fs))
                        return TRUE;
                fs.m_startCell = bDown ? GRIDPOINT{ 0, 0 } : GRIDPOINT{ (UINT)-1, (UINT)-1 };
                return g->FindText(fs);
        };

        if (dlg->FindNext())
        {
                if (!findNext())
                        AfxMessageBox(_T("No matches."), MB_OK | MB_ICONINFORMATION);
                return 0;
        }

        if (dlg->ReplaceCurrent() || dlg->ReplaceAll())
        {
                CT2W wRepl(dlg->GetReplaceString());
                GCREPLACESTRUCT rs{};
                rs.m_cbSize = sizeof(rs);
                rs.m_pwszFindText = wFind;
                rs.m_pwszReplaceText = wRepl;
                rs.m_bMatchCase = dlg->MatchCase();
                rs.m_bSearchForward = TRUE;

                if (dlg->ReplaceCurrent())
                {
                        // Replace in the active cell if it matches, otherwise
                        // in the next match; then show the match after it.
                        g->GetCurrentCell(rs.m_startCell);
                        if (!g->ReplaceText(rs))
                        {
                                rs.m_startCell = { 0, 0 };
                                if (!g->ReplaceText(rs))
                                {
                                        AfxMessageBox(_T("No matches."), MB_OK | MB_ICONINFORMATION);
                                        return 0;
                                }
                        }
                        findNext();
                        return 0;
                }

                const UINT count = g->ReplaceAll(rs);
                CString msg;
                msg.Format(_T("%u replacement(s) made."), count);
                AfxMessageBox(msg, MB_OK | MB_ICONINFORMATION);
        }
        return 0;
}

