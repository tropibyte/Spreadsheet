// GridView.cpp : implementation file
//

#include "pch.h"
#include "Spreadsheet.h"
#include "GridView.h"
#include "MainFrm.h"
#include <commctrl.h>


// CGridView
#define GRIDCTRL_ID 1
#define FORMULABAR_ID 2

IMPLEMENT_DYNCREATE(CGridView, CView)

CGridView::CGridView()
{

}

CGridView::~CGridView()
{
}

// Cut, Copy, Paste, Undo and Redo live here, next to the grid they act on.
// CSpreadsheetView used to bind the same IDs to MessageBox stubs, which won
// because MFC routes commands to the most-derived view first (H1, H3).
BEGIN_MESSAGE_MAP(CGridView, CView)
	ON_WM_CREATE()
	ON_WM_SIZE()
	ON_WM_SETFOCUS()
	ON_COMMAND(ID_EDIT_CUT, &CGridView::OnEditCut)
	ON_COMMAND(ID_EDIT_COPY, &CGridView::OnEditCopy)
	ON_COMMAND(ID_EDIT_PASTE, &CGridView::OnEditPaste)
	ON_COMMAND(ID_EDIT_UNDO, &CGridView::OnEditUndo)
	ON_COMMAND(ID_EDIT_REDO, &CGridView::OnEditRedo)
	ON_UPDATE_COMMAND_UI(ID_EDIT_CUT, &CGridView::OnUpdateEditCutCopy)
	ON_UPDATE_COMMAND_UI(ID_EDIT_COPY, &CGridView::OnUpdateEditCutCopy)
	ON_UPDATE_COMMAND_UI(ID_EDIT_PASTE, &CGridView::OnUpdateEditPaste)
	ON_UPDATE_COMMAND_UI(ID_EDIT_UNDO, &CGridView::OnUpdateEditUndo)
	ON_UPDATE_COMMAND_UI(ID_EDIT_REDO, &CGridView::OnUpdateEditRedo)
    ON_NOTIFY(NM_CLICK, GRIDCTRL_ID, &CGridView::OnGridClick)
    ON_NOTIFY(GN_CONTENTCHANGED, GRIDCTRL_ID, &CGridView::OnGridContentChanged)
    ON_NOTIFY(GN_SELCHANGE, GRIDCTRL_ID, &CGridView::OnGridRefreshBar)
    ON_NOTIFY(GN_CELLCHANGED, GRIDCTRL_ID, &CGridView::OnGridRefreshBar)
    ON_NOTIFY(GN_EDITCOMMIT, GRIDCTRL_ID, &CGridView::OnGridRefreshBar)
    ON_NOTIFY(GN_EDITCANCEL, GRIDCTRL_ID, &CGridView::OnGridRefreshBar)
    ON_NOTIFY(GN_EDITSTART, GRIDCTRL_ID, &CGridView::OnGridEditText)
    ON_NOTIFY(GN_EDITCHANGE, GRIDCTRL_ID, &CGridView::OnGridEditText)
END_MESSAGE_MAP()


// CGridView drawing

void CGridView::OnDraw(CDC* /*pDC*/)
{
	// The grid is a child window and paints itself; printing goes through
	// OnPrint, which renders cells into the printer DC.
}


// CGridView diagnostics

#ifdef _DEBUG
void CGridView::AssertValid() const
{
	CView::AssertValid();
}

#ifndef _WIN32_WCE
void CGridView::Dump(CDumpContext& dc) const
{
	CView::Dump(dc);
}
#endif
#endif //_DEBUG

void CGridView::CommitEdit()
{
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
        m_wndGridCtrl.CommitEdit(TRUE);
    // Text typed into the formula bar but not yet entered, likewise.
    if (m_formulaBar.GetSafeHwnd() != NULL)
        m_formulaBar.CommitInput();
}

void CGridView::OnEditCut()
{
    if (CEdit* pEdit = m_formulaBar.FocusedEdit())
    {
        pEdit->Cut();
        return;
    }
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
        m_wndGridCtrl.SendMessage(WM_CUT);
}

void CGridView::OnEditCopy()
{
    if (CEdit* pEdit = m_formulaBar.FocusedEdit())
    {
        pEdit->Copy();
        return;
    }
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
        m_wndGridCtrl.SendMessage(WM_COPY);
}

void CGridView::OnEditPaste()
{
    if (CEdit* pEdit = m_formulaBar.FocusedEdit())
    {
        pEdit->Paste();
        return;
    }
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
        m_wndGridCtrl.SendMessage(WM_PASTE);
}

void CGridView::OnEditUndo()
{
    if (CEdit* pEdit = m_formulaBar.FocusedEdit())
    {
        pEdit->Undo();
        return;
    }
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
        m_wndGridCtrl.Undo();
}

void CGridView::OnEditRedo()
{
    if (m_formulaBar.FocusedEdit())
        return;     // an edit box has no redo
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
        m_wndGridCtrl.Redo();
}

// Undo, Redo and Paste are greyed out when there is nothing to do (M17).
void CGridView::OnUpdateEditCutCopy(CCmdUI* pCmdUI)
{
    if (m_formulaBar.FocusedEdit())
    {
        pCmdUI->Enable(TRUE);
        return;
    }
    pCmdUI->Enable(m_wndGridCtrl.GetSafeHwnd() != NULL);
}

void CGridView::OnUpdateEditPaste(CCmdUI* pCmdUI)
{
    if (m_formulaBar.FocusedEdit())
    {
        pCmdUI->Enable(::IsClipboardFormatAvailable(CF_UNICODETEXT));
        return;
    }
    static const UINT cfGridCells = ::RegisterClipboardFormat(_T("Grid32.Cells"));
    pCmdUI->Enable(m_wndGridCtrl.GetSafeHwnd() != NULL &&
        (::IsClipboardFormatAvailable(CF_UNICODETEXT) || ::IsClipboardFormatAvailable(cfGridCells)));
}

void CGridView::OnUpdateEditUndo(CCmdUI* pCmdUI)
{
    if (CEdit* pEdit = m_formulaBar.FocusedEdit())
    {
        pCmdUI->Enable(pEdit->CanUndo());
        return;
    }
    pCmdUI->Enable(m_wndGridCtrl.GetSafeHwnd() != NULL && m_wndGridCtrl.CanUndo());
}

void CGridView::OnUpdateEditRedo(CCmdUI* pCmdUI)
{
    if (m_formulaBar.FocusedEdit())
    {
        pCmdUI->Enable(FALSE);
        return;
    }
    pCmdUI->Enable(m_wndGridCtrl.GetSafeHwnd() != NULL && m_wndGridCtrl.CanRedo());
}


// CGridView printing (H12)

BOOL CGridView::OnPreparePrinting(CPrintInfo* pInfo)
{
    CommitEdit();
    return DoPreparePrinting(pInfo);
}

// Lay the used part of the sheet out on pages: columns left to right until
// the page is full, then a new page across; rows likewise down. Pages are
// numbered down first, then across, as Excel does.
void CGridView::OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo)
{
    m_printColStarts.assign(1, 0);
    m_printRowStarts.assign(1, 0);
    GRIDSELECTION used{};
    if (!m_wndGridCtrl.GetUsedRange(used))
        used = GRIDSELECTION{};
    m_printRange = GRIDSELECTION{ { 0, 0 }, used.end };

    const int dpiX = pDC->GetDeviceCaps(LOGPIXELSX);
    const int dpiY = pDC->GetDeviceCaps(LOGPIXELSY);
    // Half an inch of margin on every side of the printable area.
    const long pageW = pDC->GetDeviceCaps(HORZRES) - dpiX;
    const long pageH = pDC->GetDeviceCaps(VERTRES) - dpiY;

    long width = 0;
    for (UINT c = 0; c <= used.end.nCol; ++c)
    {
        const long w = MulDiv((int)m_wndGridCtrl.GetColumnWidth(c), dpiX, 96);
        if (width > 0 && width + w > pageW)
        {
            m_printColStarts.push_back(c);
            width = 0;
        }
        width += w;
    }
    long height = 0;
    for (UINT r = 0; r <= used.end.nRow; ++r)
    {
        const long h = MulDiv((int)m_wndGridCtrl.GetRowHeight(r), dpiY, 96);
        if (h == 0)
            continue;   // hidden by a filter
        if (height > 0 && height + h > pageH)
        {
            m_printRowStarts.push_back(r);
            height = 0;
        }
        height += h;
    }
    pInfo->SetMaxPage((UINT)(m_printColStarts.size() * m_printRowStarts.size()));
}

void CGridView::OnPrint(CDC* pDC, CPrintInfo* pInfo)
{
    const size_t rowPages = m_printRowStarts.size();
    const size_t page = pInfo->m_nCurPage - 1;
    const size_t across = page / rowPages, down = page % rowPages;
    if (across >= m_printColStarts.size())
        return;

    GCRENDERSTRUCT rs{};
    rs.m_cbSize = sizeof(GCRENDERSTRUCT);
    // In print preview m_hDC is the screen and the attribute DC the printer;
    // the page is laid out at the printer's resolution either way.
    rs.m_hDC = pDC->m_hDC;
    rs.m_nLogPixelsX = pDC->GetDeviceCaps(LOGPIXELSX);
    rs.m_nLogPixelsY = pDC->GetDeviceCaps(LOGPIXELSY);
    rs.m_ptOrigin = { rs.m_nLogPixelsX / 2, rs.m_nLogPixelsY / 2 };
    rs.m_range.start = { m_printRowStarts[down], m_printColStarts[across] };
    rs.m_range.end.nRow = down + 1 < rowPages ? m_printRowStarts[down + 1] - 1 : m_printRange.end.nRow;
    rs.m_range.end.nCol = across + 1 < m_printColStarts.size() ? m_printColStarts[across + 1] - 1 : m_printRange.end.nCol;
    rs.m_bGridLines = TRUE;
    m_wndGridCtrl.RenderRange(rs);
}


// CGridView message handlers


int CGridView::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
    GRIDCREATESTRUCT gcs;
    memset(&gcs, 0, sizeof(GRIDCREATESTRUCT));
    gcs.cbSize = sizeof(GRIDCREATESTRUCT);
    gcs.nWidth = 702;
    gcs.nHeight = 4096;
    gcs.nDefColWidth = 135;
    gcs.nDefRowHeight = 30;
    gcs.clrCurrentCellHighlightBox = RGB(255, 0, 0);

    if (CView::OnCreate(lpCreateStruct) == -1)
        return -1;

    if (!m_wndGridCtrl.Create(GS_SPREADSHEET | GRID_DEFAULT_STYLE, CRect(0, 0, 50, 50), this, GRIDCTRL_ID, gcs))
        return -1;

    // After the grid: the bar reads the active cell as it is created.
    if (!m_formulaBar.Create(this, FORMULABAR_ID, &m_wndGridCtrl))
        return -1;

    return 0;
}


void CGridView::OnSize(UINT nType, int cx, int cy)
{
    CView::OnSize(nType, cx, cy);

    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
    {
        // The formula bar across the top, the grid below it.
        int barH = 0;
        if (m_formulaBar.GetSafeHwnd() != NULL)
        {
            barH = m_formulaBar.GetHeight();
            if (barH > cy)
                barH = cy;
            m_formulaBar.MoveWindow(0, 0, cx, barH);
        }
        m_wndGridCtrl.MoveWindow(0, barH, cx, cy - barH);
    }
}

// The view is only a frame for the grid; focus belongs to the grid, which is
// where typing starts an edit. The view used to keep focus and forward key
// presses, so characters (WM_CHAR) never reached the grid at all.
void CGridView::OnSetFocus(CWnd* pOldWnd)
{
    CView::OnSetFocus(pOldWnd);
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
        m_wndGridCtrl.SetFocus();
}

void CGridView::OnGridClick(NMHDR* pNMHDR, LRESULT* pResult)
{
    *pResult = 0;
    LPGRIDNMHDR p = reinterpret_cast<LPGRIDNMHDR>(pNMHDR);
    // App-level handlers live on CMainFrame (Format Painter etc.); CGridView
    // just forwards. If/when more click consumers appear, generalize this
    // into a dispatch table.
    if (CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetMainWnd()))
        pFrame->ConsumeGridClick(&m_wndGridCtrl, p->cellPt);
}

// The grid owns the data, so the document has no other way to know it changed:
// nothing routes edits through CDocument. Without this the modified flag was
// never set, and closing after editing threw the work away without asking.
void CGridView::OnGridContentChanged(NMHDR* /*pNMHDR*/, LRESULT* pResult)
{
    if (CDocument* pDoc = GetDocument())
        pDoc->SetModifiedFlag(TRUE);
    // Paste, fill, undo, a load: whatever changed may be the active cell.
    if (m_formulaBar.GetSafeHwnd() != NULL)
        m_formulaBar.Refresh();
    if (pResult)
        *pResult = 0;
}

void CGridView::OnGridRefreshBar(NMHDR* /*pNMHDR*/, LRESULT* pResult)
{
    if (m_formulaBar.GetSafeHwnd() != NULL)
        m_formulaBar.Refresh();
    if (pResult)
        *pResult = 0;
}

void CGridView::OnGridEditText(NMHDR* pNMHDR, LRESULT* pResult)
{
    LPGRIDNMHDR p = reinterpret_cast<LPGRIDNMHDR>(pNMHDR);
    if (m_formulaBar.GetSafeHwnd() != NULL)
        m_formulaBar.ShowEditorText(p->cellPt, p->text ? p->text : L"");
    if (pResult)
        *pResult = 0;
}


// CFormulaBar

namespace {
enum { IDC_FB_NAMEBOX = 101, IDC_FB_INPUT = 102 };

// "A".."Z", "AA".."ZZ", ... for a 0-based column.
CStringW ColumnName(UINT col)
{
    CStringW name;
    unsigned long long n = (unsigned long long)col + 1;
    while (n > 0)
    {
        name.Insert(0, (wchar_t)(L'A' + (int)((n - 1) % 26)));
        n = (n - 1) / 26;
    }
    return name;
}

CStringW CellRef(const GRIDPOINT& cell)
{
    CStringW ref;
    ref.Format(L"%s%u", (LPCWSTR)ColumnName(cell.nCol), cell.nRow + 1);
    return ref;
}

CStringW WindowTextW(HWND hWnd)
{
    CStringW text;
    const int len = ::GetWindowTextLengthW(hWnd);
    if (len > 0)
    {
        ::GetWindowTextW(hWnd, text.GetBuffer(len + 1), len + 1);
        text.ReleaseBuffer();
    }
    return text;
}

// Keys an edit box handles for its own text that the frame's accelerator
// table would otherwise turn into grid commands: copying or pasting cells,
// undoing the last cell edit. CMainFrame::PreTranslateMessage keeps the same
// list for the in-cell editor.
bool IsBoxEditingKey(WPARAM vk, bool ctrl, bool shift, bool alt)
{
    switch (vk)
    {
    case 'C': case 'X': case 'V': case 'Z': case 'Y':
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
}

BEGIN_MESSAGE_MAP(CFormulaBar, CWnd)
    ON_WM_SIZE()
    ON_WM_PAINT()
    ON_WM_ERASEBKGND()
    ON_EN_CHANGE(IDC_FB_INPUT, &CFormulaBar::OnInputChange)
    ON_EN_KILLFOCUS(IDC_FB_INPUT, &CFormulaBar::OnInputKillFocus)
    ON_EN_KILLFOCUS(IDC_FB_NAMEBOX, &CFormulaBar::OnNameKillFocus)
END_MESSAGE_MAP()

BOOL CFormulaBar::Create(CWnd* pParent, UINT nID, CGridCtrl* pGrid)
{
    m_pGrid = pGrid;

    // The system's message font, at about 11 pt rather than 9, as Excel's
    // bar is; the "fx" label in its italic.
    NONCLIENTMETRICS ncm{};
    ncm.cbSize = sizeof(ncm);
    ::SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0);
    LOGFONT lf = ncm.lfMessageFont;
    lf.lfHeight = lf.lfHeight * 11 / 9;
    if (!m_font.CreateFontIndirect(&lf))
        return FALSE;
    lf.lfItalic = TRUE;
    if (!m_fxFont.CreateFontIndirect(&lf))
        return FALSE;

    // Tall enough for one line of that font in a bordered box, with a margin
    // above and below and a separating line along the bottom.
    {
        CClientDC dc(pParent);
        CFont* pOld = dc.SelectObject(&m_font);
        TEXTMETRIC tm{};
        dc.GetTextMetrics(&tm);
        dc.SelectObject(pOld);
        m_nHeight = tm.tmHeight + tm.tmExternalLeading + 13;
    }

    const CString cls = AfxRegisterWndClass(CS_HREDRAW | CS_VREDRAW,
        ::LoadCursor(nullptr, IDC_ARROW), (HBRUSH)(COLOR_BTNFACE + 1));
    if (!CWnd::Create(cls, _T(""), WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN,
            CRect(0, 0, 0, m_nHeight), pParent, nID))
        return FALSE;

    const DWORD editStyle = WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL;
    if (!m_nameBox.Create(editStyle, CRect(0, 0, 0, 0), this, IDC_FB_NAMEBOX) ||
        !m_input.Create(editStyle, CRect(0, 0, 0, 0), this, IDC_FB_INPUT))
        return FALSE;
    m_nameBox.SetFont(&m_font);
    m_input.SetFont(&m_font);
    m_nameBox.SetLimitText(255);
    m_input.SetLimitText(0);        // no 32K cap on a formula

    Refresh();
    return TRUE;
}

bool CFormulaBar::GridAlive() const
{
    return m_pGrid && m_pGrid->GetSafeHwnd() && ::IsWindow(m_pGrid->GetSafeHwnd());
}

void CFormulaBar::Refresh()
{
    if (!GridAlive() || !m_input.GetSafeHwnd())
        return;
    // Text being typed into the box goes to the cell it was typed for
    // before the box moves on (a click on another cell, say).
    CommitInput();

    GRIDPOINT cell{};
    if (!m_pGrid->GetCurrentCell(cell))
        return;
    CStringW text;
    m_pGrid->GetCellInput(cell, text);
    m_target = cell;
    m_original = text;
    SetInputText(text);
    ShowName(cell);
}

void CFormulaBar::ShowEditorText(const GRIDPOINT& cell, LPCWSTR text)
{
    if (!m_input.GetSafeHwnd())
        return;
    m_target = cell;
    SetInputText(text);
    ShowName(cell);
}

BOOL CFormulaBar::CommitInput()
{
    if (!m_bDirty)
        return FALSE;
    m_bDirty = false;       // first: writing the cell refreshes the bar
    const CStringW text = WindowTextW(m_input.GetSafeHwnd());
    if (text == m_original || !GridAlive())
        return FALSE;
    m_original = text;
    // The cell the text was typed for, whatever is active now.
    return m_pGrid->SetCellText(m_target, text);
}

CEdit* CFormulaBar::FocusedEdit()
{
    const HWND hFocus = ::GetFocus();
    if (!hFocus)
        return nullptr;
    if (hFocus == m_input.GetSafeHwnd())
        return &m_input;
    if (hFocus == m_nameBox.GetSafeHwnd())
        return &m_nameBox;
    return nullptr;
}

void CFormulaBar::SetInputText(LPCWSTR text)
{
    // Leave the box alone when it already says this, so a refresh while it
    // has the focus does not move the caret.
    if (WindowTextW(m_input.GetSafeHwnd()) != text)
    {
        m_bLoading = true;
        ::SetWindowTextW(m_input.GetSafeHwnd(), text);
        m_bLoading = false;
    }
    m_bDirty = false;
}

void CFormulaBar::ShowName(const GRIDPOINT& cell)
{
    // Not while someone is typing a reference into it.
    if (::GetFocus() != m_nameBox.GetSafeHwnd())
        ::SetWindowTextW(m_nameBox.GetSafeHwnd(), CellRef(cell));
}

void CFormulaBar::LeaveInput(int dRow, int dCol)
{
    const GRIDPOINT from = m_target;
    CommitInput();
    if (!GridAlive())
        return;
    // Enter moves down a row and Tab right a column, as in Excel; with Shift
    // they go back. A move off the grid's edge stays where it is.
    const long long row = (long long)from.nRow + dRow;
    const long long col = (long long)from.nCol + dCol;
    if (row >= 0 && col >= 0)
        m_pGrid->SetCurrentCell((UINT)row, (UINT)col);
    m_pGrid->SetFocus();
}

void CFormulaBar::CancelInput()
{
    SetInputText(m_original);
    if (GridAlive())
        m_pGrid->SetFocus();
}

void CFormulaBar::GoToName()
{
    if (!GridAlive())
        return;
    CStringW ref = WindowTextW(m_nameBox.GetSafeHwnd());
    ref.Trim();
    // A reference first ("B7", "A1:C3"), then the name of a cell.
    if (!ref.IsEmpty() &&
        (m_pGrid->SetCurrentCell((LPCWSTR)ref, (short)SETBYCOORDINATE) ||
         m_pGrid->SetCurrentCell((LPCWSTR)ref, (short)SETBYNAME)))
    {
        m_pGrid->SetFocus();
        Refresh();
        return;
    }
    ::MessageBeep(MB_ICONWARNING);
    m_nameBox.SetSel(0, -1);
}

BOOL CFormulaBar::PreTranslateMessage(MSG* pMsg)
{
    const bool inInput = pMsg->hwnd == m_input.GetSafeHwnd();
    const bool inName = pMsg->hwnd == m_nameBox.GetSafeHwnd();
    if ((inInput || inName) && (pMsg->message == WM_KEYDOWN || pMsg->message == WM_SYSKEYDOWN))
    {
        const bool ctrl = (::GetKeyState(VK_CONTROL) & 0x8000) != 0;
        const bool shift = (::GetKeyState(VK_SHIFT) & 0x8000) != 0;
        const bool alt = (::GetKeyState(VK_MENU) & 0x8000) != 0;
        switch (pMsg->wParam)
        {
        case VK_RETURN:
            if (alt)
                break;
            if (inName)
                GoToName();
            else
                LeaveInput(shift ? -1 : 1, 0);
            return TRUE;
        case VK_TAB:
            if (ctrl || alt)
                break;
            if (inInput)
                LeaveInput(0, shift ? -1 : 1);
            return TRUE;
        case VK_ESCAPE:
            if (inName)
            {
                if (GridAlive())
                    m_pGrid->SetFocus();    // OnNameKillFocus puts the name back
            }
            else
                CancelInput();
            return TRUE;
        case 'A':
            if (ctrl && !alt)
            {
                ::SendMessage(pMsg->hwnd, EM_SETSEL, 0, -1);
                return TRUE;
            }
            break;
        }
        // Clipboard and undo keys belong to the box, not to the frame's
        // accelerators, which would copy or paste cells instead.
        if (IsBoxEditingKey(pMsg->wParam, ctrl, shift, alt))
        {
            ::TranslateMessage(pMsg);
            ::DispatchMessage(pMsg);
            return TRUE;
        }
    }
    return CWnd::PreTranslateMessage(pMsg);
}

void CFormulaBar::OnInputChange()
{
    if (!m_bLoading)
        m_bDirty = true;
}

void CFormulaBar::OnInputKillFocus()
{
    // Clicking a cell, the name box or another window enters what was typed,
    // as leaving the in-cell editor does.
    CommitInput();
}

void CFormulaBar::OnNameKillFocus()
{
    GRIDPOINT cell{};
    if (GridAlive() && m_pGrid->GetCurrentCell(cell))
        ShowName(cell);
}

void CFormulaBar::OnSize(UINT nType, int cx, int cy)
{
    CWnd::OnSize(nType, cx, cy);
    if (!m_input.GetSafeHwnd())
        return;

    const int margin = 4;
    int boxH = cy - 2 * margin - 1;         // -1: the line along the bottom
    if (boxH < 0)
        boxH = 0;
    int nameW = 0, fxW = 0;
    {
        CClientDC dc(this);
        CFont* pOld = dc.SelectObject(&m_font);
        nameW = dc.GetTextExtent(_T("WWW00000")).cx + 12;     // fits "ZZ4096" with room
        fxW = dc.GetTextExtent(_T("fx")).cx + 24;
        dc.SelectObject(pOld);
    }
    int x = margin;
    m_nameBox.MoveWindow(x, margin, nameW, boxH);
    x += nameW;
    m_fxRect.SetRect(x, margin, x + fxW, margin + boxH);
    x += fxW;
    int inputW = cx - margin - x;
    if (inputW < 0)
        inputW = 0;
    m_input.MoveWindow(x, margin, inputW, boxH);
}

BOOL CFormulaBar::OnEraseBkgnd(CDC* /*pDC*/)
{
    return TRUE;        // OnPaint fills the whole bar
}

void CFormulaBar::OnPaint()
{
    CPaintDC dc(this);
    CRect rc;
    GetClientRect(&rc);
    dc.FillSolidRect(rc, ::GetSysColor(COLOR_BTNFACE));
    dc.FillSolidRect(rc.left, rc.bottom - 1, rc.Width(), 1, ::GetSysColor(COLOR_3DSHADOW));
    CFont* pOld = dc.SelectObject(&m_fxFont);
    dc.SetBkMode(TRANSPARENT);
    dc.SetTextColor(::GetSysColor(COLOR_BTNTEXT));
    dc.DrawText(_T("fx"), m_fxRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    dc.SelectObject(pOld);
}
