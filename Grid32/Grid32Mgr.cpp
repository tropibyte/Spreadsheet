#include "pch.h"
#include <algorithm>
#include <CommCtrl.h>
#include <sstream>
#include <cwchar>
#include <cwctype>
#include <set>
#include <tuple>
#include <climits>
#include "Grid32Mgr.h"
#include "grid32_internal.h"
#include "Formulator.h"
#include <new>
#include <string_view>


// Cell-type inference helpers (DateToSerial, TryParseDate, TryParseNumber,
// IEquals, InferType) are defined inline in grid32_internal.h under the
// Grid32Detail namespace so the test project can call them without
// pulling in this entire TU. Callers in this file go through Grid32Detail::.

namespace {
// The empty cell every new grid starts with. Whole-sheet formatting
// (Select All + a format command) changes the grid's copy of it, which SSF
// saves as its DF record; ClearAll puts this one back.
GRIDCELL BuiltInDefaultCell()
{
    GRIDCELL cell;
    cell.clrBackground = RGB(255, 255, 255);
    cell.m_clrBorderColor = 0;
    cell.m_nBorderWidth = 1;
    cell.penStyle = PS_DOT;
    cell.fontInfo.m_wsFontFace = L"Arial";
    cell.fontInfo.m_fPointSize = 12.0f;
    cell.fontInfo.m_clrTextColor = RGB(0, 0, 0);
    cell.fontInfo.bItalic = FALSE;
    cell.fontInfo.bUnderline = FALSE;
    cell.fontInfo.bWeight = 400;                    // Normal weight (400)
    cell.justification = DT_CENTER | DT_VCENTER;
    return cell;
}
}

CGrid32Mgr::CGrid32Mgr() :
m_hWndGrid(NULL), nColHeaderHeight(40), nRowHeaderWidth(70), m_editWndProc(nullptr),
m_nMouseHoverDelay(400), m_npHoverDelaySet(0), m_nToBeSized(~0), m_resizeStartSize(0),
m_bEditing(false), m_editCell{ 0, 0 }, m_hEditFont(NULL),
m_rgbSizingLine(RGB(0, 0, 64)), m_nSizingLine(0), m_bResizable(true), m_bUndoRecordEnabled(TRUE),
m_bDeferRecalc(FALSE), m_bSuppressChangeNotify(FALSE),
m_bSizing(false), m_bSelecting(false), m_hDefaultFont(NULL),
m_clientRect{ 0, 0, 0, 0 }, m_paintRect{ 0, 0, 0, 0 }, m_gdiplusToken(0), m_gridHitTest(0),
m_hWndEdit(NULL), m_pBatch(nullptr), m_bMergesDirty(true), m_cfGridCells(0)
{
    memset(&gcs, 0, sizeof(GRIDCREATESTRUCT));
    m_defaultGridCell = BuiltInDefaultCell();
    m_cornerCell.clrBackground = RGB(234, 234, 240);
    m_cornerCell.m_clrBorderColor = RGB(0, 0, 0);
    m_cornerCell.m_nBorderWidth = 1;
    m_cornerCell.penStyle = PS_SOLID;
    m_cornerCell.fontInfo.m_wsFontFace = L"Arial";         // Font face
    m_cornerCell.fontInfo.m_fPointSize = 11.50f;            // Point size
    m_cornerCell.fontInfo.m_clrTextColor = RGB(0, 0, 0);   // Black text color
    m_cornerCell.fontInfo.bItalic = FALSE;                 // No italic
    m_cornerCell.fontInfo.bUnderline = FALSE;              // No underline
    m_cornerCell.fontInfo.bWeight = 400;                   // Normal weight (400)
    m_cornerCell.justification = DT_CENTER | DT_VCENTER;
    m_currentCell = m_visibleTopLeft = m_HoverCell = { 0,0 };
    m_bRedraw = TRUE;
    dwError = 0;
    memset(&m_selectionRect, 0, sizeof(GRIDSELECTION));
    m_bSelecting = FALSE;
    m_mouseDraggingStartPoint = { 0, 0 };
}


CGrid32Mgr::~CGrid32Mgr()
{
    // OnDestroy is idempotent — covers the WM_CREATE-fail path where the
    // window never received WM_DESTROY so OnDestroy was never called.
    OnDestroy();
    DeleteAllCells();
}


namespace {
// Sets a BOOL member for the lifetime of the scope and puts the old value
// back on the way out, including when the body throws.
class FlagScope
{
public:
    FlagScope(BOOL& flag, BOOL value) : m_flag(flag), m_previous(flag) { m_flag = value; }
    ~FlagScope() { m_flag = m_previous; }
    FlagScope(const FlagScope&) = delete;
    FlagScope& operator=(const FlagScope&) = delete;
private:
    BOOL& m_flag;
    BOOL  m_previous;
};

// Marks a region that applies many cells as one user-visible operation.
//
// It suspends two things: per-edit recalculation, which would otherwise be
// quadratic over the edit, and the content-changed notification, which the
// caller sends once when the scope closes instead of once per cell.
//
// Both flags are restored to their previous value rather than cleared, so a
// nested bulk operation can't re-enable either while an outer one is still
// running, and an exception can't leave them stuck on.
class BulkEditScope
{
public:
    BulkEditScope(BOOL& deferRecalc, BOOL& suppressNotify)
        : m_deferRecalc(deferRecalc), m_prevDefer(deferRecalc),
          m_suppressNotify(suppressNotify), m_prevSuppress(suppressNotify)
    {
        m_deferRecalc = TRUE;
        m_suppressNotify = TRUE;
    }
    ~BulkEditScope()
    {
        m_deferRecalc = m_prevDefer;
        m_suppressNotify = m_prevSuppress;
    }
    BulkEditScope(const BulkEditScope&) = delete;
    BulkEditScope& operator=(const BulkEditScope&) = delete;
private:
    BOOL& m_deferRecalc;    BOOL m_prevDefer;
    BOOL& m_suppressNotify; BOOL m_prevSuppress;
};

// Pixel height of a font at the given DPI, as a negative (character-height)
// LOGFONT value.
int CellFontHeight(float pointSize, int logPixelsY)
{
    return -(int)(pointSize * logPixelsY / 72.0 + 0.5);
}
}

// ---- batch undo --------------------------------------------------------------

CGrid32Mgr::BatchUndo::BatchUndo(CGrid32Mgr& mgr)
    : m_mgr(mgr), m_pOuter(mgr.m_pBatch), m_prevRecord(mgr.m_bUndoRecordEnabled),
      m_bCollect(false)
{
    if (!m_pOuter)
    {
        // Nothing to collect when recording is already off (an undo or redo
        // is replaying, or a caller suspended it on purpose).
        m_bCollect = m_prevRecord != FALSE;
        m_mgr.m_pBatch = this;
        m_mgr.m_bUndoRecordEnabled = FALSE;
    }
}

void CGrid32Mgr::BatchUndo::Touch(UINT nRow, UINT nCol)
{
    if (m_pOuter)
    {
        m_pOuter->Touch(nRow, nCol);
        return;
    }
    if (!m_bCollect)
        return;
    auto key = std::make_pair(nRow, nCol);
    if (m_cells.count(key))
        return;     // only the state before the first change matters
    CellChange& ch = m_cells[key];
    ch.pt = { nRow, nCol };
    if (PGRIDCELL p = m_mgr.GetCell(nRow, nCol))
    {
        ch.bHadOld = true;
        ch.oldCell = *p;
    }
}

void CGrid32Mgr::BatchUndo::TouchDefault()
{
    if (m_pOuter)
    {
        m_pOuter->TouchDefault();
        return;
    }
    if (!m_bCollect || m_bDefault)
        return;
    m_bDefault = true;
    m_oldDefault = m_mgr.m_defaultGridCell;
}

CGrid32Mgr::BatchUndo::~BatchUndo()
{
    if (m_pOuter)
        return;     // the outermost scope records for everyone
    m_mgr.m_pBatch = nullptr;
    m_mgr.m_bUndoRecordEnabled = m_prevRecord;
    if (!m_bCollect)
        return;
    try
    {
        GridEditOperation op;
        op.type = EditOperationType::Batch;
        for (auto& entry : m_cells)
        {
            CellChange& ch = entry.second;
            PGRIDCELL p = m_mgr.GetCell(ch.pt.nRow, ch.pt.nCol);
            ch.bHasNew = p != nullptr;
            if (p)
                ch.newCell = *p;
            const bool unchanged = ch.bHadOld == ch.bHasNew &&
                (!ch.bHadOld || CellsEqual(ch.oldCell, ch.newCell));
            if (!unchanged)
                op.changes.push_back(std::move(ch));
        }
        if (m_bDefault && !CellsEqual(m_oldDefault, m_mgr.m_defaultGridCell))
        {
            op.bDefaultChanged = true;
            op.oldDefault = m_oldDefault;
            op.newDefault = m_mgr.m_defaultGridCell;
        }
        // A command that changed nothing leaves no empty step behind.
        if (!op.changes.empty() || op.bDefaultChanged)
            m_mgr.RecordUndoOperation(op);
    }
    catch (...)
    {
        // Out of memory while building the record: the edit stands, it just
        // can't be undone. Never let that escape a destructor.
        m_mgr.SetLastError(GRID_ERROR_OUT_OF_MEMORY);
    }
}

bool CGrid32Mgr::CellsEqual(const GRIDCELL& a, const GRIDCELL& b)
{
    const CELLFONT& fa = a.fontInfo;
    const CELLFONT& fb = b.fontInfo;
    const bool sameValue = a.m_dValue == b.m_dValue ||
        (std::isnan(a.m_dValue) && std::isnan(b.m_dValue));
    return fa.m_wsFontFace == fb.m_wsFontFace && fa.m_fPointSize == fb.m_fPointSize &&
        fa.m_clrTextColor == fb.m_clrTextColor && fa.bItalic == fb.bItalic &&
        fa.bUnderline == fb.bUnderline && fa.bStrikeThrough == fb.bStrikeThrough &&
        fa.bWeight == fb.bWeight &&
        a.clrBackground == b.clrBackground && a.m_clrBorderColor == b.m_clrBorderColor &&
        a.m_nBorderWidth == b.m_nBorderWidth && a.penStyle == b.penStyle &&
        a.justification == b.justification && a.m_nFormat == b.m_nFormat &&
        a.m_wsName == b.m_wsName && a.m_wsText == b.m_wsText &&
        a.m_wsFormula == b.m_wsFormula && a.m_bFormula == b.m_bFormula &&
        a.m_eType == b.m_eType && sameValue &&
        a.mergeRange.start.nRow == b.mergeRange.start.nRow &&
        a.mergeRange.start.nCol == b.mergeRange.start.nCol &&
        a.mergeRange.end.nRow == b.mergeRange.end.nRow &&
        a.mergeRange.end.nCol == b.mergeRange.end.nCol;
}

// Bijective base 26: 0 -> A, 25 -> Z, 26 -> AA, 701 -> ZZ, 702 -> AAA. The
// loop ends when the quotient runs out, not by unsigned wrap-around (M20).
std::wstring CGrid32Mgr::GetColumnLabel(size_t index)
{
    std::wstring label;
    size_t n = index + 1;
    while (n > 0)
    {
        label.insert(label.begin(), (wchar_t)(L'A' + (n - 1) % 26));
        n = (n - 1) / 26;
    }
    return label;
}

bool CGrid32Mgr::Create(PGRIDCREATESTRUCT pGCS)
{
    // Store the provided GRIDCREATESTRUCT in the member variable
    gcs = *pGCS;

    if (gcs.nWidth == 0 || gcs.nHeight == 0 ||
        gcs.nWidth > (size_t)UINT_MAX || gcs.nHeight > (size_t)UINT_MAX)
        return false;

    // If the host left nDefRowHeight at 0, keep the constructor's
    // default (40) so the column header doesn't collapse.
    if (gcs.style & GS_COLHEADER)
    {
        if (gcs.nDefRowHeight > 0)
            nColHeaderHeight = static_cast<long>(gcs.nDefRowHeight);
        // else keep the existing default initialized in the constructor.
    }
    else
    {
        nColHeaderHeight = 0;
    }

    // Row and column sizes are stored only where they differ from these
    // defaults, so nothing is allocated per row or column here (P7).

    m_hWndEdit = CreateWindowEx(0, _T("EDIT"), _T(""), WS_CHILD | ES_AUTOHSCROLL,
        0, 0, 100, 20, m_hWndGrid, (HMENU)ID_CELL_EDIT, GetModuleHandle(NULL), NULL);

    m_editWndProc = (WNDPROC)SetWindowLongPtr(m_hWndEdit, GWLP_WNDPROC,
        (LONG_PTR)EditCtrl_WndProc);

    if (!m_editWndProc) {
        return false; // Fail if subclassing edit control fails
    }

    // Formulas and formatting travel between grids in this format;
    // CF_UNICODETEXT alongside it carries what other programs can read.
    m_cfGridCells = RegisterClipboardFormatW(L"Grid32.Cells");

    SetLastError(0);

    GdiplusStartupInput gdiplusStartupInput;
    if (GdiplusStartup(&m_gdiplusToken, &gdiplusStartupInput, NULL) != Ok) {
        return false; // Initialization failed
    }

    HDC hDC = GetDC(m_hWndGrid);
    m_hDefaultFont = CreateFontW(
        CellFontHeight(m_defaultGridCell.fontInfo.m_fPointSize, GetDeviceCaps(hDC, LOGPIXELSY)), 0, 0, 0,
        m_defaultGridCell.fontInfo.bWeight, m_defaultGridCell.fontInfo.bItalic,
        m_defaultGridCell.fontInfo.bUnderline, 0, 0, 0, 0, 0, 0,
        m_defaultGridCell.fontInfo.m_wsFontFace.c_str()
    );

    ReleaseDC(m_hWndGrid, hDC);
    SendMessage(m_hWndEdit, WM_SETFONT, (WPARAM)m_hDefaultFont, 0);

    return true; // Return true if creation succeeded, false otherwise
}

void CGrid32Mgr::OnDestroy()
{
    if (m_hWndEdit != NULL && IsWindow(m_hWndEdit))
    {
        SendMessage(m_hWndEdit, WM_SETFONT, 0, 0);
        // Restore the original WNDPROC so late-queue messages to the edit
        // window (e.g. EM_GETSEL from background tooltip/IME work) don't
        // route through EditCtrl_WndProc with a stale CGrid32Mgr pointer.
        if (m_editWndProc != nullptr)
        {
            SetWindowLongPtr(m_hWndEdit, GWLP_WNDPROC, (LONG_PTR)m_editWndProc);
            m_editWndProc = nullptr;
        }
    }
    m_bEditing = false;
    m_hEditFont = NULL;
    ClearFontCache();
    if (m_hDefaultFont != NULL)
    {
        DeleteObject(m_hDefaultFont);
        m_hDefaultFont = NULL;
    }
    if (m_gdiplusToken != 0)
    {
        GdiplusShutdown(m_gdiplusToken);
        m_gdiplusToken = 0;
    }
}

// ---- geometry ----------------------------------------------------------------

size_t CGrid32Mgr::ColW(size_t nCol) const
{
    if (nCol >= gcs.nWidth)
        return 0;
    auto it = m_colWidths.find((UINT)nCol);
    return it != m_colWidths.end() ? it->second : gcs.nDefColWidth;
}

size_t CGrid32Mgr::RowH(size_t nRow) const
{
    if (nRow >= gcs.nHeight || m_hiddenRows.count((UINT)nRow))
        return 0;
    auto it = m_rowHeights.find((UINT)nRow);
    return it != m_rowHeights.end() ? it->second : gcs.nDefRowHeight;
}

// Width of columns [start, end): the default width times the count,
// corrected for the few columns that have their own. O(sized columns), not
// O(columns), which is what lets the grid be any size.
size_t CGrid32Mgr::CalculatedColumnDistance(size_t start, size_t end) const
{
    if (end > gcs.nWidth)
        end = gcs.nWidth;
    if (start >= end)
        return 0;
    long long total = (long long)(end - start) * (long long)gcs.nDefColWidth;
    for (auto it = m_colWidths.lower_bound((UINT)start); it != m_colWidths.end() && it->first < end; ++it)
        total += (long long)it->second - (long long)gcs.nDefColWidth;
    return total > 0 ? (size_t)total : 0;
}

// Height of rows [start, end), hidden rows counting as zero.
size_t CGrid32Mgr::CalculatedRowDistance(size_t start, size_t end) const
{
    if (end > gcs.nHeight)
        end = gcs.nHeight;
    if (start >= end)
        return 0;
    long long total = (long long)(end - start) * (long long)gcs.nDefRowHeight;
    for (auto it = m_rowHeights.lower_bound((UINT)start); it != m_rowHeights.end() && it->first < end; ++it)
        total += (long long)it->second - (long long)gcs.nDefRowHeight;
    for (auto it = m_hiddenRows.lower_bound((UINT)start); it != m_hiddenRows.end() && *it < end; ++it)
    {
        auto sized = m_rowHeights.find(*it);
        total -= (long long)(sized != m_rowHeights.end() ? sized->second : gcs.nDefRowHeight);
    }
    return total > 0 ? (size_t)total : 0;
}

// Client x of a column's left edge. Columns scrolled off to the left get a
// negative position, and nCol == nWidth gives the grid's right edge.
long long CGrid32Mgr::GetColumnStartPos(UINT nCol)
{
    long long pos = GetActualRowHeaderWidth();
    if (nCol >= m_visibleTopLeft.nCol)
        return pos + (long long)CalculatedColumnDistance(m_visibleTopLeft.nCol, nCol);
    return pos - (long long)CalculatedColumnDistance(nCol, m_visibleTopLeft.nCol);
}

long long CGrid32Mgr::GetRowStartPos(UINT nRow)
{
    long long pos = GetActualColHeaderHeight();
    if (nRow >= m_visibleTopLeft.nRow)
        return pos + (long long)CalculatedRowDistance(m_visibleTopLeft.nRow, nRow);
    return pos - (long long)CalculatedRowDistance(nRow, m_visibleTopLeft.nRow);
}

// The rectangle a cell occupies, or the whole merged block it belongs to.
// Returns whether any of it is inside the area being painted.
bool CGrid32Mgr::GetCellRect(UINT nRow, UINT nCol, RECT& rect)
{
    GRIDPOINT start{ nRow, nCol }, end{ nRow, nCol };
    if (const MERGERANGE* m = FindMerge(nRow, nCol))
    {
        start = m->start;
        end = m->end;
    }
    rect.left = (long)GetColumnStartPos(start.nCol);
    rect.top = (long)GetRowStartPos(start.nRow);
    rect.right = (long)GetColumnStartPos(end.nCol + 1);
    rect.bottom = (long)GetRowStartPos(end.nRow + 1);
    return CanDrawRect(rect);
}

void CGrid32Mgr::GetCurrentCellRect(RECT &rect)
{
    GetCellRect(m_currentCell.nRow, m_currentCell.nCol, rect);
}

bool CGrid32Mgr::CanDrawRect(const RECT& rect)
{
    // m_paintRect is the update region while a paint is running and the
    // client area otherwise; empty rects (hidden rows) never draw.
    if (rect.right <= rect.left || rect.bottom <= rect.top)
        return false;
    return !(rect.left >= m_paintRect.right ||
        rect.right <= m_paintRect.left ||
        rect.top >= m_paintRect.bottom ||
        rect.bottom <= m_paintRect.top);
}

bool CGrid32Mgr::IsCellVisible(UINT nRow, UINT nCol)
{
    if (nRow >= (UINT)gcs.nHeight || nCol >= (UINT)gcs.nWidth)
        return false;
    if (m_visibleTopLeft.nRow > nRow || m_visibleTopLeft.nCol > nCol)
        return false;
    const long long right = GetColumnStartPos(nCol + 1);
    const long long bottom = GetRowStartPos(nRow + 1);
    return right <= m_clientRect.right && bottom <= m_clientRect.bottom;
}

// The block of columns and rows that fits entirely in the client area,
// starting at the top-left visible cell (always at least that one cell).
void CGrid32Mgr::CalculatePageStats(PAGESTAT& pageStat)
{
    memset(&pageStat, 0, sizeof(PAGESTAT));
    pageStat.start = m_visibleTopLeft;
    pageStat.end = m_visibleTopLeft;
    const long long availW = (long long)m_clientRect.right - GetActualRowHeaderWidth();
    const long long availH = (long long)m_clientRect.bottom - GetActualColHeaderHeight();

    long long used = 0;
    for (size_t idx = m_visibleTopLeft.nCol; idx < gcs.nWidth; ++idx)
    {
        long long w = (long long)ColW(idx);
        if (used + w > availW && idx > m_visibleTopLeft.nCol)
            break;
        used += w;
        pageStat.end.nCol = (UINT)idx;
    }
    pageStat.nWidth = (size_t)used;

    used = 0;
    for (size_t idx = m_visibleTopLeft.nRow; idx < gcs.nHeight; ++idx)
    {
        long long h = (long long)RowH(idx);
        if (used + h > availH && idx > m_visibleTopLeft.nRow)
            break;
        used += h;
        pageStat.end.nRow = (UINT)idx;
    }
    pageStat.nHeight = (size_t)used;
}

// ---- font cache --------------------------------------------------------------

bool CGrid32Mgr::FontKey::operator<(const FontKey& o) const
{
    return std::tie(height, weight, italic, underline, strike, face) <
        std::tie(o.height, o.weight, o.italic, o.underline, o.strike, o.face);
}

HFONT CGrid32Mgr::GetCachedFont(const CELLFONT& fi, int lfHeight)
{
    FontKey key{ fi.m_wsFontFace, lfHeight, fi.bWeight,
        fi.bItalic ? TRUE : FALSE, fi.bUnderline ? TRUE : FALSE, fi.bStrikeThrough ? TRUE : FALSE };
    auto it = m_fontCache.find(key);
    if (it != m_fontCache.end())
        return it->second;
    HFONT hFont = CreateFontW(lfHeight, 0, 0, 0, fi.bWeight, key.italic, key.underline,
        key.strike, 0, 0, 0, 0, 0, fi.m_wsFontFace.c_str());
    if (hFont)
        m_fontCache.emplace(std::move(key), hFont);
    return hFont;
}

void CGrid32Mgr::ClearFontCache()
{
    for (auto& entry : m_fontCache)
        DeleteObject(entry.second);
    m_fontCache.clear();
}

// ---- painting ----------------------------------------------------------------

void CGrid32Mgr::OnPaint(PAINTSTRUCT& ps)
{
    HDC hDC = ps.hdc;

    // Only the update region is drawn and copied to the screen (P6); every
    // Draw* helper skips what CanDrawRect says lies outside it.
    RECT rcPaint = ps.rcPaint;
    if (rcPaint.right <= rcPaint.left || rcPaint.bottom <= rcPaint.top)
        rcPaint = m_clientRect;

    // A sheet that has used many distinct fonts drops them all between
    // paints, never during one (fonts are selected into DCs while drawing),
    // and never while the editor holds one.
    if (m_fontCache.size() > 256 && !m_bEditing)
        ClearFontCache();

    // Resources declared outside try so cleanup is guaranteed even on throw.
    HDC memDC = NULL;
    HBITMAP memBitmap = NULL;
    HBITMAP oldBitmap = NULL;

    try
    {
        int width = m_clientRect.right - m_clientRect.left;
        int height = m_clientRect.bottom - m_clientRect.top;

        memDC = CreateCompatibleDC(hDC);
        memBitmap = CreateCompatibleBitmap(hDC, width, height);
        oldBitmap = (HBITMAP)SelectObject(memDC, memBitmap);

        m_paintRect = rcPaint;

        HBRUSH hBrush = CreateSolidBrush(m_defaultGridCell.clrBackground);
        FillRect(memDC, &rcPaint, hBrush);
        DeleteObject(hBrush);

        DrawGrid(memDC);
        DrawCells(memDC);
        DrawSelectionSurround(memDC);
        DrawCurrentCellBox(memDC);
        DrawHeader(memDC);
        DrawVoidSpace(memDC);
        DrawSizingLine(memDC);

        BitBlt(hDC, rcPaint.left, rcPaint.top, rcPaint.right - rcPaint.left,
            rcPaint.bottom - rcPaint.top, memDC, rcPaint.left, rcPaint.top, SRCCOPY);
    }
    catch (const std::exception&)
    {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
    }
    m_paintRect = m_clientRect;

    // Unconditional cleanup — exception or not.
    if (memDC != NULL && oldBitmap != NULL)
        SelectObject(memDC, oldBitmap);
    if (memBitmap != NULL)
        DeleteObject(memBitmap);
    if (memDC != NULL)
        DeleteDC(memDC);
}

void CGrid32Mgr::DrawCurrentCellBox(HDC hDC)
{
    if (!(gcs.style & GS_HIGHLIGHTSELECTION))
        return;

    RECT rect;
    if (!GetCellRect(m_currentCell.nRow, m_currentCell.nCol, rect))
        return;

    HPEN hPen = CreatePen(PS_SOLID, 3, gcs.clrCurrentCellHighlightBox);
    HPEN hOldPen = (HPEN)SelectObject(hDC, hPen);
    HBRUSH hOldBrush = (HBRUSH)SelectObject(hDC, GetStockObject(NULL_BRUSH));
    Rectangle(hDC, rect.left, rect.top, rect.right, rect.bottom);
    SelectObject(hDC, hOldPen);
    SelectObject(hDC, hOldBrush);
    DeleteObject(hPen);
}


void CGrid32Mgr::DrawHeader(HDC hDC)
{
    if (gcs.style & GS_ROWHEADER)
        DrawRowHeaders(hDC);
    if (gcs.style & GS_COLHEADER)
        DrawColHeaders(hDC);
    if ((gcs.style & GS_ROWHEADER) && (gcs.style & GS_COLHEADER))
    {
        DrawHeaderButton(hDC, m_cornerCell.clrBackground, m_cornerCell.m_clrBorderColor,
            RECT{ 0, 0, (long)nRowHeaderWidth, (long)nColHeaderHeight }, true);
    }
}

void CGrid32Mgr::DrawHeaderButton(HDC hDC, COLORREF background, COLORREF border, RECT coordinates, bool bRaised)
{
    // Create pens and brushes for drawing
    HPEN hBorderPen = CreatePen(PS_SOLID, 1, border);
    HBRUSH hBackgroundBrush = CreateSolidBrush(background);

    // Select the border pen and draw the rectangle border
    HPEN hOldPen = (HPEN)SelectObject(hDC, hBorderPen);
    HBRUSH hOldBrush = (HBRUSH)SelectObject(hDC, hBackgroundBrush);

    Rectangle(hDC, coordinates.left, coordinates.top, coordinates.right, coordinates.bottom);

    // Clean up GDI objects
    SelectObject(hDC, hOldPen);
    SelectObject(hDC, hOldBrush);
    DeleteObject(hBorderPen);
    DeleteObject(hBackgroundBrush);

    if (bRaised)
    {
        // Calculate highlight (average with white) and shadow (average with black) colors
        COLORREF highlight = RGB(GetRValue(background) <= 192 ? (GetRValue(background) + 255) / 2 : 255,
            GetGValue(background) <= 192 ? (GetGValue(background) + 255) / 2 : 255,
            GetBValue(background) <= 192 ? (GetBValue(background) + 255) / 2 : 255);

        COLORREF shadow = RGB(GetRValue(background) / 2,
            GetGValue(background) / 2,
            GetBValue(background) / 2);

        // Create pens for highlight and shadow
        HPEN hHighlightPen = CreatePen(PS_SOLID, 1, highlight);
        HPEN hShadowPen = CreatePen(PS_SOLID, 1, shadow);

        // Draw the top and left highlight
        HPEN hOldPen = (HPEN)SelectObject(hDC, hHighlightPen);
        // Top highlight
        MoveToEx(hDC, coordinates.left, coordinates.bottom - 2, NULL);
        LineTo(hDC, coordinates.left, coordinates.top);
        LineTo(hDC, coordinates.right - 2, coordinates.top);
        // Bottom right shadow
        SelectObject(hDC, hShadowPen);
        LineTo(hDC, coordinates.right - 2, coordinates.bottom - 2);
        LineTo(hDC, coordinates.left, coordinates.bottom - 2);

        // Draw the second highlight (top and left inner)
        SelectObject(hDC, hHighlightPen);
        MoveToEx(hDC, coordinates.left + 1, coordinates.bottom - 3, NULL);
        LineTo(hDC, coordinates.left + 1, coordinates.top + 1);
        LineTo(hDC, coordinates.right - 3, coordinates.top + 1);
        // Bottom right inner shadow
        SelectObject(hDC, hShadowPen);
        LineTo(hDC, coordinates.right - 3, coordinates.bottom - 3);
        LineTo(hDC, coordinates.left + 1, coordinates.bottom - 3);

        // Clean up GDI objects
        SelectObject(hDC, hOldPen);
        DeleteObject(hHighlightPen);
        DeleteObject(hShadowPen);
    }
}


void CGrid32Mgr::DrawColHeaders(HDC hDC)
{
    // Header labels are positive (cell-height) font sizes, as they always were.
    HFONT hFont = GetCachedFont(m_cornerCell.fontInfo,
        (int)((long double)m_cornerCell.fontInfo.m_fPointSize * GetDeviceCaps(hDC, LOGPIXELSY) / 72.0));
    HFONT hOldFont = (HFONT)SelectObject(hDC, hFont);
    SetBkMode(hDC, TRANSPARENT);
    SetTextColor(hDC, m_cornerCell.fontInfo.m_clrTextColor);

    RECT rect{ GetActualRowHeaderWidth(), 0, GetActualRowHeaderWidth(), GetActualColHeaderHeight() };
    for (size_t idx = m_visibleTopLeft.nCol; idx < gcs.nWidth && rect.left < m_clientRect.right; ++idx)
    {
        rect.right = rect.left + (long)ColW(idx);
        if (CanDrawRect(rect))
        {
            DrawHeaderButton(hDC, m_cornerCell.clrBackground, m_cornerCell.m_clrBorderColor, rect, true);
            if (gcs.style & GS_SPREADSHEETHEADER)
            {
                std::wstring text = GetColumnLabel(idx);
                DrawTextW(hDC, text.c_str(), -1, &rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
            }
        }
        rect.left = rect.right;
    }
    SelectObject(hDC, hOldFont);
}

void CGrid32Mgr::DrawRowHeaders(HDC hDC)
{
    HFONT hFont = GetCachedFont(m_cornerCell.fontInfo,
        (int)((long double)m_cornerCell.fontInfo.m_fPointSize * GetDeviceCaps(hDC, LOGPIXELSY) / 72.0));
    HFONT hOldFont = (HFONT)SelectObject(hDC, hFont);
    SetBkMode(hDC, TRANSPARENT);
    SetTextColor(hDC, m_cornerCell.fontInfo.m_clrTextColor);

    RECT rect{ 0, GetActualColHeaderHeight(), GetActualRowHeaderWidth(), GetActualColHeaderHeight() };
    for (size_t idx = m_visibleTopLeft.nRow; idx < gcs.nHeight && rect.top < m_clientRect.bottom; ++idx)
    {
        rect.bottom = rect.top + (long)RowH(idx);
        if (CanDrawRect(rect))
        {
            DrawHeaderButton(hDC, m_cornerCell.clrBackground, m_cornerCell.m_clrBorderColor, rect, true);
            if (gcs.style & GS_SPREADSHEETHEADER)
            {
                std::wstring text = std::to_wstring(idx + 1);
                DrawTextW(hDC, text.c_str(), -1, &rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
            }
        }
        rect.top = rect.bottom;
    }
    SelectObject(hDC, hOldFont);
}

// One line per visible row and column boundary. This used to walk every
// column from 0 and, inside it, every row from 0, creating and deleting a
// brush for each cell it passed on the way to the visible ones.
void CGrid32Mgr::DrawGrid(HDC hDC)
{
    if (!(gcs.style & GS_DRAWLINES))
        return;

    const long left = GetActualRowHeaderWidth();
    const long top = GetActualColHeaderHeight();
    const long right = (long)min((long long)m_clientRect.right, GetColumnStartPos((UINT)gcs.nWidth));
    const long bottom = (long)min((long long)m_clientRect.bottom, GetRowStartPos((UINT)gcs.nHeight));

    HPEN hBorderPen = CreatePen(m_defaultGridCell.penStyle, m_defaultGridCell.m_nBorderWidth, m_defaultGridCell.m_clrBorderColor);
    HPEN hOldPen = (HPEN)SelectObject(hDC, hBorderPen);

    long y = top;
    for (size_t row = m_visibleTopLeft.nRow; row < gcs.nHeight && y <= m_clientRect.bottom; ++row)
    {
        long h = (long)RowH(row);
        if (h == 0)
            continue;
        y += h;
        if (y - 1 >= m_paintRect.top && y - 1 < m_paintRect.bottom)
        {
            MoveToEx(hDC, max(left, m_paintRect.left), y - 1, NULL);
            LineTo(hDC, min(right, m_paintRect.right), y - 1);
        }
    }

    long x = left;
    for (size_t col = m_visibleTopLeft.nCol; col < gcs.nWidth && x <= m_clientRect.right; ++col)
    {
        x += (long)ColW(col);
        if (x - 1 >= m_paintRect.left && x - 1 < m_paintRect.right)
        {
            MoveToEx(hDC, x - 1, max(top, m_paintRect.top), NULL);
            LineTo(hDC, x - 1, min(bottom, m_paintRect.bottom));
        }
    }
    SelectObject(hDC, hOldPen);
    DeleteObject(hBorderPen);
}

void CGrid32Mgr::DrawSelectionOverlay(HDC hDC, RECT &gridCellRect)
{
    BYTE alpha = 255 / 5;
    DrawSemiTransparentRectangle(hDC, gridCellRect, gcs.clrSelectionOverlay, alpha);
}

// Background, border and text of one cell (or merged block) in rect.
void CGrid32Mgr::DrawCellContent(HDC hDC, const GRIDCELL& cell, const RECT& rect, int logPixelsY, bool bFillBackground)
{
    if (bFillBackground || cell.penStyle != m_defaultGridCell.penStyle ||
        cell.clrBackground != m_defaultGridCell.clrBackground)
    {
        HBRUSH hBackgroundBrush = CreateSolidBrush(cell.clrBackground);
        HPEN hBorderPen = CreatePen(cell.penStyle, cell.m_nBorderWidth, cell.m_clrBorderColor);
        HPEN hWhiteOutPen = CreatePen(PS_SOLID, 1, cell.clrBackground);
        HPEN hOldPen = (HPEN)SelectObject(hDC, hWhiteOutPen);
        HBRUSH hOldBrush = (HBRUSH)SelectObject(hDC, hBackgroundBrush);
        Rectangle(hDC, rect.left, rect.top, rect.right, rect.bottom);
        SelectObject(hDC, hBorderPen);
        Rectangle(hDC, rect.left, rect.top, rect.right, rect.bottom);
        SelectObject(hDC, hOldPen);
        SelectObject(hDC, hOldBrush);
        DeleteObject(hBorderPen);
        DeleteObject(hWhiteOutPen);
        DeleteObject(hBackgroundBrush);
    }
    if (cell.m_wsText.empty())
        return;

    SetBkMode(hDC, TRANSPARENT);
    SetTextColor(hDC, cell.fontInfo.m_clrTextColor);
    HFONT hFont = GetCachedFont(cell.fontInfo, CellFontHeight(cell.fontInfo.m_fPointSize, logPixelsY));
    HFONT hOldFont = (HFONT)SelectObject(hDC, hFont);

    // DT_WORDBREAK only takes effect when DT_SINGLELINE is off, so cells with
    // the wrap bit drop the single-line flag. DT_NOPREFIX keeps "A&B" from
    // rendering as an underlined B.
    UINT j = cell.justification;
    UINT flags = ((j & DT_WORDBREAK) ? j : (j | DT_SINGLELINE)) | DT_NOPREFIX;
    RECT textRect = rect;
    InflateRect(&textRect, -2, 0);
    DrawTextW(hDC, cell.m_wsText.c_str(), -1, &textRect, flags);

    SelectObject(hDC, hOldFont);
}

// Draws the populated cells in view. Empty cells are the background OnPaint
// already filled, so only cells that exist are visited, and each merged block
// is drawn once at its full size.
void CGrid32Mgr::DrawCells(HDC hDC)
{
    const int logPixelsY = GetDeviceCaps(hDC, LOGPIXELSY);

    // Column and row edges of the visible block, computed once per paint.
    std::vector<long> colX, rowY;
    {
        long x = GetActualRowHeaderWidth();
        colX.push_back(x);
        for (size_t c = m_visibleTopLeft.nCol; c < gcs.nWidth && x < m_clientRect.right; ++c)
            colX.push_back(x += (long)ColW(c));
        long y = GetActualColHeaderHeight();
        rowY.push_back(y);
        for (size_t r = m_visibleTopLeft.nRow; r < gcs.nHeight && y < m_clientRect.bottom; ++r)
            rowY.push_back(y += (long)RowH(r));
    }
    const UINT firstCol = m_visibleTopLeft.nCol, endCol = firstCol + (UINT)colX.size() - 1;
    const UINT firstRow = m_visibleTopLeft.nRow, endRow = firstRow + (UINT)rowY.size() - 1;

    RefreshMergeIndex();
    for (const MERGERANGE& m : m_merges)
    {
        if (m.end.nRow < firstRow || m.start.nRow >= endRow || m.end.nCol < firstCol || m.start.nCol >= endCol)
            continue;
        RECT rect;
        if (GetCellRect(m.start.nRow, m.start.nCol, rect))
            DrawCellContent(hDC, *GetCellOrDefault(m.start.nRow, m.start.nCol), rect, logPixelsY, true);
    }

    for (UINT row = firstRow; row < endRow; ++row)
    {
        const long top = rowY[row - firstRow], bottom = rowY[row - firstRow + 1];
        if (bottom <= top)
            continue;   // hidden by a filter
        for (auto it = mapCells.lower_bound({ row, firstCol });
            it != mapCells.end() && it->first.first == row && it->first.second < endCol; ++it)
        {
            const UINT col = it->first.second;
            if (!m_merges.empty() && FindMerge(row, col))
                continue;   // drawn above as part of its block
            RECT rect{ colX[col - firstCol], top, colX[col - firstCol + 1], bottom };
            if (CanDrawRect(rect))
                DrawCellContent(hDC, *it->second, rect, logPixelsY, false);
        }
    }

    // One translucent overlay for the selection, with a hole where the
    // active cell is, instead of a GDI+ fill per selected cell.
    if (IsSelectionUnicellular())
        return;
    GRIDSELECTION sel = GetNormalizedSelection();
    RECT selRect{ (long)GetColumnStartPos(sel.start.nCol), (long)GetRowStartPos(sel.start.nRow),
                  (long)GetColumnStartPos(sel.end.nCol + 1), (long)GetRowStartPos(sel.end.nRow + 1) };
    RECT area{ GetActualRowHeaderWidth(), GetActualColHeaderHeight(), m_clientRect.right, m_clientRect.bottom };
    RECT visibleSel;
    if (!IntersectRect(&visibleSel, &selRect, &area))
        return;
    RECT cur;
    GetCurrentCellRect(cur);
    RECT hole;
    if (!IntersectRect(&hole, &cur, &visibleSel))
    {
        if (CanDrawRect(visibleSel))
            DrawSelectionOverlay(hDC, visibleSel);
        return;
    }
    RECT bands[4] = {
        { visibleSel.left, visibleSel.top, visibleSel.right, hole.top },
        { visibleSel.left, hole.bottom, visibleSel.right, visibleSel.bottom },
        { visibleSel.left, hole.top, hole.left, hole.bottom },
        { hole.right, hole.top, visibleSel.right, hole.bottom },
    };
    for (RECT& band : bands)
        if (CanDrawRect(band))
            DrawSelectionOverlay(hDC, band);
}

void CGrid32Mgr::DrawVoidSpace(HDC hDC)
{
    bool bDrawBorder = false;
    long nMaxWidth = static_cast<long>(CalculatedColumnDistance(m_visibleTopLeft.nCol, gcs.nWidth));
    long nMaxHeight = static_cast<long>(CalculatedRowDistance(m_visibleTopLeft.nRow, gcs.nHeight));
    if (nMaxWidth < m_clientRect.right)
    {
        HBRUSH hBrush = CreateSolidBrush(RGB(128, 128, 128));
        HPEN hPen = CreatePen(PS_SOLID, 1, RGB(128, 128, 128));

        HPEN oldPen = (HPEN)SelectObject(hDC, hPen);
        HBRUSH oldBrush = (HBRUSH)SelectObject(hDC, hBrush);

        Rectangle(hDC, nMaxWidth + GetActualRowHeaderWidth(), m_clientRect.top, m_clientRect.right, m_clientRect.bottom);

        SelectObject(hDC, oldPen);
        SelectObject(hDC, oldBrush);
        DeleteObject(hPen);
        DeleteObject(hBrush);

        bDrawBorder = true;
    }
    if (nMaxHeight < m_clientRect.bottom)
    {
        HBRUSH hBrush = CreateSolidBrush(RGB(128, 128, 128));
        HPEN hPen = CreatePen(PS_SOLID, 1, RGB(128, 128, 128));

        HPEN oldPen = (HPEN)SelectObject(hDC, hPen);
        HBRUSH oldBrush = (HBRUSH)SelectObject(hDC, hBrush);

        Rectangle(hDC, 0, nMaxHeight + GetActualColHeaderHeight(), m_clientRect.right, m_clientRect.bottom);

        SelectObject(hDC, oldPen);
        SelectObject(hDC, oldBrush);
        DeleteObject(hPen);
        DeleteObject(hBrush);

        bDrawBorder = true;
    }

    if (bDrawBorder)
    {
        int x = nMaxWidth < m_clientRect.right ? nMaxWidth : m_clientRect.right;
        int y = nMaxHeight < m_clientRect.bottom ? nMaxHeight : m_clientRect.bottom;

        x += GetActualRowHeaderWidth() + 1;
        y += GetActualColHeaderHeight() + 1;

        HPEN hPen = CreatePen(PS_SOLID, 3, 0);
        HPEN oldPen = (HPEN)SelectObject(hDC, hPen);

        if (x < m_clientRect.right)
        {
            MoveToEx(hDC, x, 0, NULL);
            LineTo(hDC, x, y < m_clientRect.bottom ? y : m_clientRect.bottom);
        }

        if (y < m_clientRect.bottom)
        {
            MoveToEx(hDC, 0, y, NULL);
            LineTo(hDC, x < m_clientRect.right ? x : m_clientRect.right, y);
        }


        SelectObject(hDC, oldPen);
        DeleteObject(hPen);
    }
}

void CGrid32Mgr::DrawSizingLine(HDC hDC)
{
    if (!m_bResizable || !m_bSizing)
        return;

    HPEN hOldPen = NULL, hPen = CreatePen(PS_DASHDOT, 1, m_rgbSizingLine);
    POINT pt;
    GetCursorPos(&pt);
    ScreenToClient(m_hWndGrid, &pt);
    hOldPen = (HPEN)SelectObject(hDC, hPen);
    if (m_gridHitTest == GRID_HTCOLDIVIDER)
    {
        MoveToEx(hDC, pt.x, 0, NULL);
        LineTo(hDC, pt.x, m_clientRect.bottom);
    }
    else if (m_gridHitTest == GRID_HTROWDIVIDER)
    {
        MoveToEx(hDC, 0, pt.y, NULL);
        LineTo(hDC, m_clientRect.right, pt.y);
    }
    SelectObject(hDC, hOldPen);
    DeleteObject(hPen);
}

void CGrid32Mgr::DrawSemiTransparentRectangle(HDC hdc, RECT rect, COLORREF color, BYTE alpha)
{
    Graphics graphics(hdc);
    Color fillColor(alpha, GetRValue(color), GetGValue(color), GetBValue(color));
    SolidBrush brush(fillColor);
    Gdiplus::Rect rect2;
    rect2.X = rect.left;
    rect2.Y = rect.top;
    rect2.Width = rect.right - rect.left;
    rect2.Height = rect.bottom - rect.top;
    graphics.FillRectangle(&brush, rect2);
}

void CGrid32Mgr::DrawSelectionSurround(HDC hDC)
{
    if (IsSelectionUnicellular())
        return;

    GRIDSELECTION sel = GetNormalizedSelection();
    RECT rect{ (long)GetColumnStartPos(sel.start.nCol), (long)GetRowStartPos(sel.start.nRow),
               (long)GetColumnStartPos(sel.end.nCol + 1), (long)GetRowStartPos(sel.end.nRow + 1) };
    if (!CanDrawRect(rect))
        return;

    HPEN hPen = CreatePen(PS_SOLID, 3, gcs.clrSelectionOverlay), hOldPen;
    HBRUSH hOldBrush;
    hOldPen = (HPEN)SelectObject(hDC, hPen);
    hOldBrush = (HBRUSH)SelectObject(hDC, GetStockObject(NULL_BRUSH));
    Rectangle(hDC, rect.left, rect.top, rect.right, rect.bottom);
    SelectObject(hDC, hOldBrush);
    SelectObject(hDC, hOldPen);
    DeleteObject(hPen);
}

// ---- merged cells ------------------------------------------------------------
// A merge is stored on its anchor (top-left) cell as mergeRange, which is the
// form SSF already persists; m_merges is the index painting and hit-testing
// read, rebuilt whenever a mutator that could have moved an anchor ran.

void CGrid32Mgr::RefreshMergeIndex()
{
    if (!m_bMergesDirty)
        return;
    m_merges.clear();
    for (const auto& entry : mapCells)
    {
        const MERGERANGE& m = entry.second->mergeRange;
        const bool isAnchor = m.start.nRow == entry.first.first && m.start.nCol == entry.first.second;
        const bool spans = m.end.nRow > m.start.nRow || m.end.nCol > m.start.nCol;
        if (isAnchor && spans && m.end.nRow >= m.start.nRow && m.end.nCol >= m.start.nCol &&
            m.end.nRow < gcs.nHeight && m.end.nCol < gcs.nWidth)
            m_merges.push_back(m);
    }
    m_bMergesDirty = false;
}

const MERGERANGE* CGrid32Mgr::FindMerge(UINT nRow, UINT nCol)
{
    RefreshMergeIndex();
    for (const MERGERANGE& m : m_merges)
        if (nRow >= m.start.nRow && nRow <= m.end.nRow && nCol >= m.start.nCol && nCol <= m.end.nCol)
            return &m;
    return nullptr;
}

bool CGrid32Mgr::IsCellMerged(UINT row, UINT col)
{
    return FindMerge(row, col) != nullptr;
}

// Grow sel until no merged block straddles its edge.
void CGrid32Mgr::ExpandForMerges(GRIDSELECTION& sel)
{
    NormalizeSelectionRect(sel);
    RefreshMergeIndex();
    bool grown = true;
    while (grown)
    {
        grown = false;
        for (const MERGERANGE& m : m_merges)
        {
            const bool overlaps = !(m.end.nRow < sel.start.nRow || m.start.nRow > sel.end.nRow ||
                                    m.end.nCol < sel.start.nCol || m.start.nCol > sel.end.nCol);
            if (!overlaps)
                continue;
            if (m.start.nRow < sel.start.nRow) { sel.start.nRow = m.start.nRow; grown = true; }
            if (m.start.nCol < sel.start.nCol) { sel.start.nCol = m.start.nCol; grown = true; }
            if (m.end.nRow > sel.end.nRow)     { sel.end.nRow = m.end.nRow;     grown = true; }
            if (m.end.nCol > sel.end.nCol)     { sel.end.nCol = m.end.nCol;     grown = true; }
        }
    }
}

// Merge sel into one block shown with its top-left cell's content. Merges it
// overlaps are absorbed. The other cells keep their content (hidden while
// merged), so unmerging gives it back.
BOOL CGrid32Mgr::MergeCells(const GRIDSELECTION& selIn)
{
    GRIDSELECTION sel = selIn;
    ExpandForMerges(sel);
    if (sel.end.nRow >= gcs.nHeight || sel.end.nCol >= gcs.nWidth)
    {
        SetLastError(GRID_ERROR_OUT_OF_RANGE);
        return FALSE;
    }
    if (sel.start.nRow == sel.end.nRow && sel.start.nCol == sel.end.nCol)
    {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
        return FALSE;
    }
    {
        BatchUndo batch(*this);
        for (const GRIDPOINT& pt : CellsInRange(sel))
        {
            PGRIDCELL p = GetCellOrCreate(pt.nRow, pt.nCol);
            if (p)
                p->mergeRange = MERGERANGE();
        }
        PGRIDCELL anchor = GetCellOrCreate(sel.start.nRow, sel.start.nCol);
        if (!anchor)
            return FALSE;
        anchor->mergeRange.start = sel.start;
        anchor->mergeRange.end = sel.end;
        m_bMergesDirty = true;
    }
    if (m_currentCell.nRow >= sel.start.nRow && m_currentCell.nRow <= sel.end.nRow &&
        m_currentCell.nCol >= sel.start.nCol && m_currentCell.nCol <= sel.end.nCol)
        m_currentCell = sel.start;
    m_selectionRect = sel;
    Invalidate();
    NotifyContentChanged();
    SetLastError(0);
    return TRUE;
}

BOOL CGrid32Mgr::UnmergeCells(const GRIDSELECTION& selIn)
{
    GRIDSELECTION sel = selIn;
    ExpandForMerges(sel);
    bool any = false;
    {
        BatchUndo batch(*this);
        for (const GRIDPOINT& pt : CellsInRange(sel))
        {
            PGRIDCELL p = GetCell(pt.nRow, pt.nCol);
            if (p && (p->mergeRange.end.nRow != p->mergeRange.start.nRow ||
                      p->mergeRange.end.nCol != p->mergeRange.start.nCol))
            {
                TouchCell(pt.nRow, pt.nCol);
                p->mergeRange = MERGERANGE();
                any = true;
            }
        }
        m_bMergesDirty = true;
    }
    if (!any)
    {
        SetLastError(GRID_ERROR_NOT_EXIST);
        return FALSE;
    }
    Invalidate();
    NotifyContentChanged();
    SetLastError(0);
    return TRUE;
}

// ---- cell store --------------------------------------------------------------

// Retrieve an existing cell pointer or nullptr if absent
PGRIDCELL CGrid32Mgr::GetCell(UINT nRow, UINT nCol)
{
    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);
    return (it != mapCells.end()) ? it->second : nullptr;
}

// Retrieve an existing cell or the default cell if absent. The result is for
// reading only; writes go through GetCellOrCreate.
PGRIDCELL CGrid32Mgr::GetCellOrDefault(UINT nRow, UINT nCol)
{
    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);
    return (it != mapCells.end()) ? it->second : &m_defaultGridCell;
}

// Retrieve an existing cell or create one. On failure this returns nullptr:
// it used to return &m_defaultGridCell, and a caller that then styled "the
// cell" restyled every empty cell in the grid (AUDIT_2026-10 S10).
PGRIDCELL CGrid32Mgr::GetCellOrCreate(UINT nRow, UINT nCol, bool bRecord)
{
    if (nRow >= gcs.nHeight || nCol >= gcs.nWidth)
    {
        SetLastError(GRID_ERROR_OUT_OF_RANGE);
        return nullptr;
    }
    // The caller is about to write through the pointer.
    TouchCell(nRow, nCol);
    m_bMergesDirty = true;

    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);
    if (it != mapCells.end())
        return it->second;

    PGRIDCELL pCell = new (std::nothrow) GRIDCELL(m_defaultGridCell);
    if (!pCell)
    {
        SetLastError(GRID_ERROR_OUT_OF_MEMORY);
        SendGridNotification(NM_OUTOFMEMORY);
        return nullptr;
    }
    try
    {
        mapCells[key] = pCell;
    }
    catch (const std::bad_alloc&)
    {
        delete pCell;
        SetLastError(GRID_ERROR_OUT_OF_MEMORY);
        SendGridNotification(NM_OUTOFMEMORY);
        return nullptr;
    }

    if (bRecord)
    {
        GridEditOperation op;
        op.type = EditOperationType::SetFullCell;
        op.row = nRow;
        op.col = nCol;
        op.oldState = m_defaultGridCell;
        op.bOldMissing = true;
        op.newState = *pCell;
        RecordUndoOperation(op);
    }

    return pCell;
}

PGRIDCELL CGrid32Mgr::CreateCell(UINT nRow, UINT nCol)
{
    if (nRow >= gcs.nHeight || nCol >= gcs.nWidth)
    {
        SetLastError(GRID_ERROR_OUT_OF_RANGE);
        return nullptr;
    }
    PGRIDCELL pCell = new (std::nothrow) GRIDCELL(m_defaultGridCell);
    if (!pCell)
    {
        SetLastError(GRID_ERROR_OUT_OF_MEMORY);
        SendGridNotification(NM_OUTOFMEMORY);
        return nullptr;
    }
    TouchCell(nRow, nCol);
    m_bMergesDirty = true;
    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);
    if (it != mapCells.end())
    {
        delete it->second;
        it->second = pCell;
        return pCell;
    }
    try
    {
        mapCells[key] = pCell;
    }
    catch (const std::bad_alloc&)
    {
        delete pCell;
        SetLastError(GRID_ERROR_OUT_OF_MEMORY);
        return nullptr;
    }
    return pCell;
}

// Replace a cell wholesale. The new cell is allocated before the old one is
// freed, so a failed allocation leaves the old cell in place rather than a
// dangling pointer in the map (M1).
void CGrid32Mgr::SetCell(UINT nRow, UINT nCol, const GRIDCELL& gc)
{
    if (nRow >= gcs.nHeight || nCol >= gcs.nWidth)
    {
        SetLastError(GRID_ERROR_OUT_OF_RANGE);
        return;
    }
    PGRIDCELL pNew = new (std::nothrow) GRIDCELL(gc);
    if (!pNew)
    {
        SetLastError(GRID_ERROR_OUT_OF_MEMORY);
        SendGridNotification(NM_OUTOFMEMORY);
        return;
    }
    TouchCell(nRow, nCol);

    GridEditOperation op;
    op.type = EditOperationType::SetFullCell;
    op.row = nRow;
    op.col = nCol;
    op.newState = gc;

    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);
    if (it != mapCells.end())
    {
        op.oldState = *it->second;
        delete it->second;
        it->second = pNew;
    }
    else
    {
        op.oldState = m_defaultGridCell;
        op.bOldMissing = true;
        try
        {
            mapCells[key] = pNew;
        }
        catch (const std::bad_alloc&)
        {
            delete pNew;
            SetLastError(GRID_ERROR_OUT_OF_MEMORY);
            SendGridNotification(NM_OUTOFMEMORY);
            return;
        }
    }
    m_bMergesDirty = true;

    RecordUndoOperation(op);
    RecalculateFormulas();
    NotifyContentChanged();
    SetLastError(0);
}

std::vector<GRIDPOINT> CGrid32Mgr::CellsInRange(const GRIDSELECTION& selIn)
{
    GRIDSELECTION sel = selIn;
    NormalizeSelectionRect(sel);
    std::vector<GRIDPOINT> pts;
    auto it = mapCells.lower_bound({ sel.start.nRow, sel.start.nCol });
    while (it != mapCells.end() && it->first.first <= sel.end.nRow)
    {
        const UINT r = it->first.first, c = it->first.second;
        if (c < sel.start.nCol)
        {
            it = mapCells.lower_bound({ r, sel.start.nCol });
            continue;
        }
        if (c > sel.end.nCol)
        {
            if (r == UINT_MAX)
                break;
            it = mapCells.lower_bound({ r + 1, sel.start.nCol });
            continue;
        }
        pts.push_back({ r, c });
        ++it;
    }
    return pts;
}

bool CGrid32Mgr::IsWholeGrid(const GRIDSELECTION& selIn) const
{
    GRIDSELECTION sel = selIn;
    if (sel.start.nRow > sel.end.nRow) std::swap(sel.start.nRow, sel.end.nRow);
    if (sel.start.nCol > sel.end.nCol) std::swap(sel.start.nCol, sel.end.nCol);
    return sel.start.nRow == 0 && sel.start.nCol == 0 &&
        sel.end.nRow + 1 >= gcs.nHeight && sel.end.nCol + 1 >= gcs.nWidth;
}

namespace {
bool AnyFormulaCells(const std::map<std::pair<UINT, UINT>, PGRIDCELL>& cells)
{
    for (const auto& entry : cells)
        if (entry.second->m_bFormula)
            return true;
    return false;
}
}

UINT CGrid32Mgr::GetCurrentCellTextLen()
{
    return GetCellTextLen(m_currentCell.nRow, m_currentCell.nCol);
}

UINT CGrid32Mgr::GetCellTextLen(const GRIDPOINT& gridPt)
{
    return GetCellTextLen(gridPt.nRow, gridPt.nCol);
}

UINT CGrid32Mgr::GetCellTextLen(UINT nRow, UINT nCol)
{
    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);

    if (it != mapCells.end()) {
        return (UINT)it->second->m_wsText.length();
    }

    return 0;
}

void CGrid32Mgr::GetCurrentCellText(LPWSTR pText, UINT nLen)
{
    return GetCellText(m_currentCell.nRow, m_currentCell.nCol, pText, nLen);
}

void CGrid32Mgr::GetCellText(const GRIDPOINT& gridPt, LPWSTR pText, UINT nLen)
{
    return GetCellText(gridPt.nRow, gridPt.nCol, pText, nLen);
}

void CGrid32Mgr::GetCellText(UINT nRow, UINT nCol, LPWSTR pText, UINT nLen)
{
    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);

    if (pText && nLen > 0)
    {
        if (it != mapCells.end()) {
            // _TRUNCATE: write up to nLen-1 chars and always null-terminate.
            // Passing src.length() invokes the invalid-parameter handler
            // (debug abort / release crash) when the source is too long for
            // the buffer; _TRUNCATE silently truncates instead.
            _tcsncpy_s(pText, nLen,
                it->second->m_wsText.c_str(), _TRUNCATE);
            return;
        }

        *pText = 0;
        return;
    }
}

void CGrid32Mgr::SetCurrentCellText(LPCWSTR newText)
{
    SetCellText(m_currentCell.nRow, m_currentCell.nCol, newText);
}

void CGrid32Mgr::SetCellText(const GRIDPOINT& gridPt, LPCWSTR newText)
{
    SetCellText(gridPt.nRow, gridPt.nCol, newText);
}


void CGrid32Mgr::SetCellText(UINT nRow, UINT nCol, LPCWSTR newText)
{
    if (nRow >= gcs.nHeight || nCol >= gcs.nWidth)
    {
        SetLastError(GRID_ERROR_OUT_OF_RANGE);
        return;
    }
    TouchCell(nRow, nCol);

    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);

    PGRIDCELL pCell = nullptr;
    GridEditOperation op;
    op.type = EditOperationType::SetText;
    op.row = nRow;
    op.col = nCol;

    if (it != mapCells.end()) {
        pCell = it->second;
        op.oldState = *pCell;
    }
    else {
        pCell = new (std::nothrow) GRIDCELL(m_defaultGridCell);
        if (!pCell)
        {
            SetLastError(GRID_ERROR_OUT_OF_MEMORY);
            SendGridNotification(NM_OUTOFMEMORY);
            return;
        }
        try
        {
            mapCells[key] = pCell;
        }
        catch (const std::bad_alloc&)
        {
            delete pCell;
            SetLastError(GRID_ERROR_OUT_OF_MEMORY);
            SendGridNotification(NM_OUTOFMEMORY);
            return;
        }
        op.oldState = m_defaultGridCell;
        op.bOldMissing = true;
    }

    // Apply the new text or formula
    if (newText && newText[0] == L'=') {
        pCell->m_bFormula = true;
        pCell->m_wsFormula = newText + 1;
        RefreshFormulaCell(pCell);
    } else {
        pCell->m_bFormula = false;
        pCell->m_wsFormula.clear();
        std::wstring raw = newText ? newText : L"";
        double v = 0.0;
        pCell->m_eType = Grid32Detail::InferType(raw, v);
        pCell->m_dValue = v;
        // Apply per-cell number format. For FMT_GENERAL or text cells the
        // user's raw input is the display text; otherwise format the
        // numeric value (currency/percent/date/time/etc.).
        pCell->m_wsText = Grid32Detail::FormatCellValue(pCell->m_eType,
            pCell->m_dValue, raw, pCell->m_nFormat);
    }
    op.newState = *pCell;

    // Record undo operation
    RecordUndoOperation(op);

    // Every other mutator (SetCell, ClearCellText, DeleteCell) recalculates;
    // this one didn't, and it is the path the edit control and GM_SETCELLTEXT
    // both take. That left =A1*2 showing a stale result after A1 was typed
    // into. Bulk callers suppress this with BulkEditScope and recalculate
    // once at the end.
    //
    // The undo record is captured above on purpose: it stores the cell's
    // source, and the display text of every dependent is derived state that
    // undo re-derives by recalculating again.
    RecalculateFormulas();
    NotifyContentChanged();

    // Only this cell changed on screen unless formulas could have moved too.
    if (AnyFormulaCells(mapCells))
        Invalidate();
    else
        InvalidateCell(nRow, nCol);
    SetLastError(0);
}


// Clear a cell's contents — text, formula, type and value — and keep its
// formatting. A cell left identical to an empty one is removed outright, so
// clearing doesn't leave the map populated with blanks.
void CGrid32Mgr::ClearCellText(UINT nRow, UINT nCol)
{
    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);
    if (it == mapCells.end())
        return;

    TouchCell(nRow, nCol);
    PGRIDCELL pCell = it->second;
    GridEditOperation op;
    op.type = EditOperationType::SetText;
    op.row = nRow;
    op.col = nCol;
    op.oldState = *pCell;

    pCell->m_wsText.clear();
    pCell->m_wsFormula.clear();
    pCell->m_bFormula = false;
    // Formulas read m_dValue now, so a cleared cell must stop reporting its
    // old number (AUDIT_2026-10 S3/S4).
    pCell->m_eType = CT_Text;
    pCell->m_dValue = 0.0;

    if (CellsEqual(*pCell, m_defaultGridCell))
    {
        delete pCell;
        mapCells.erase(it);
        m_bMergesDirty = true;
        op.type = EditOperationType::Delete;
        op.newState = m_defaultGridCell;
    }
    else
    {
        op.newState = *pCell;
    }

    RecordUndoOperation(op);
    RecalculateFormulas();
    NotifyContentChanged();
    if (AnyFormulaCells(mapCells))
        Invalidate();
    else
        InvalidateCell(nRow, nCol);
    SetLastError(0);
}

// Delete a cell and support undo
void CGrid32Mgr::DeleteCell(UINT nRow, UINT nCol)
{
    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);
    if (it != mapCells.end())
    {
        TouchCell(nRow, nCol);
        GridEditOperation op;
        op.type = EditOperationType::Delete;
        op.row = nRow;
        op.col = nCol;
        op.oldState = *it->second;
        op.newState = m_defaultGridCell;

        delete it->second;
        mapCells.erase(it);
        m_bMergesDirty = true;

        RecordUndoOperation(op);
        RecalculateFormulas();
        NotifyContentChanged();
        SetLastError(0);
    }
}


void CGrid32Mgr::DeleteAllCells()
{
    for (auto& cellPair : mapCells)
    {
        delete cellPair.second;
    }
    mapCells.clear();
    m_merges.clear();
    m_bMergesDirty = true;
    RecalculateFormulas();
}

// ---- formatting --------------------------------------------------------------

void CGrid32Mgr::ApplyFontMask(CELLFONT& dst, const CELLFONT& src, DWORD mask)
{
    if (mask & GCFM_FACE)      dst.m_wsFontFace = src.m_wsFontFace;
    if (mask & GCFM_SIZE)      dst.m_fPointSize = src.m_fPointSize;
    if (mask & GCFM_COLOR)     dst.m_clrTextColor = src.m_clrTextColor;
    if (mask & GCFM_BOLD)      dst.bWeight = src.bWeight;
    if (mask & GCFM_ITALIC)    dst.bItalic = src.bItalic;
    if (mask & GCFM_UNDERLINE) dst.bUnderline = src.bUnderline;
    if (mask & GCFM_STRIKE)    dst.bStrikeThrough = src.bStrikeThrough;
}

// Every formatting command funnels through here, so all of them are one
// undo step each (H9, M24), none recalculates (M8), and none allocates a
// cell per grid position when the whole sheet is selected (S10).
template <typename Fn>
void CGrid32Mgr::ForEachFormattedCell(const GRIDSELECTION& selIn, Fn fn)
{
    GRIDSELECTION sel = selIn;
    NormalizeSelectionRect(sel);
    if (sel.start.nRow >= gcs.nHeight || sel.start.nCol >= gcs.nWidth)
    {
        SetLastError(GRID_ERROR_OUT_OF_RANGE);
        return;
    }
    if (sel.end.nRow >= gcs.nHeight) sel.end.nRow = (UINT)gcs.nHeight - 1;
    if (sel.end.nCol >= gcs.nWidth) sel.end.nCol = (UINT)gcs.nWidth - 1;

    DWORD error = 0;
    {
        BatchUndo batch(*this);
        if (IsWholeGrid(sel))
        {
            // Restyle what exists, plus the template every new cell starts as.
            batch.TouchDefault();
            fn(m_defaultGridCell);
            for (auto& entry : mapCells)
            {
                TouchCell(entry.first.first, entry.first.second);
                fn(*entry.second);
            }
        }
        else
        {
            for (UINT r = sel.start.nRow; r <= sel.end.nRow && !error; ++r)
            {
                for (UINT c = sel.start.nCol; c <= sel.end.nCol; ++c)
                {
                    PGRIDCELL p = GetCellOrCreate(r, c);
                    if (!p)
                    {
                        error = GetLastError();
                        break;
                    }
                    fn(*p);
                }
            }
        }
    }
    Invalidate();
    NotifyContentChanged();
    SetLastError(error);
}

void CGrid32Mgr::SetCellFormat(UINT nRow, UINT nCol, const CELLFONT& fi, DWORD mask)
{
    GRIDSELECTION one{ { nRow, nCol }, { nRow, nCol } };
    ForEachFormattedCell(one, [&](GRIDCELL& cell) { ApplyFontMask(cell.fontInfo, fi, mask); });
}

void CGrid32Mgr::SetSelectionFormat(const CELLFONT& fi, DWORD mask)
{
    SetRangeFormat(GetNormalizedSelection(), fi, mask);
}

void CGrid32Mgr::SetRangeFormat(const GRIDSELECTION& sel, const CELLFONT& fi, DWORD mask)
{
    ForEachFormattedCell(sel, [&](GRIDCELL& cell) { ApplyFontMask(cell.fontInfo, fi, mask); });
}

// Re-render a cell's display text after its number format changed. Text
// cells stay literal; General shows the value's plain edit form rather than
// keeping whatever the previous format rendered.
static void RenderNumberFormat(GRIDCELL& cell)
{
    if (cell.m_bFormula || cell.m_eType == CT_Text || cell.m_eType == CT_Formula || cell.m_eType == CT_Error)
        return;
    cell.m_wsText = cell.m_nFormat == FMT_GENERAL
        ? Grid32Detail::FormatEditValue(cell.m_eType, cell.m_dValue, cell.m_wsText)
        : Grid32Detail::FormatCellValue(cell.m_eType, cell.m_dValue, cell.m_wsText, cell.m_nFormat);
}

void CGrid32Mgr::SetCellNumberFormat(UINT nRow, UINT nCol, UINT format)
{
    GRIDSELECTION one{ { nRow, nCol }, { nRow, nCol } };
    ForEachFormattedCell(one, [&](GRIDCELL& cell) {
        cell.m_nFormat = format;
        if (cell.m_bFormula)
            RefreshFormulaCell(&cell);      // the evaluator renders the result in m_nFormat
        else
            RenderNumberFormat(cell);
    });
}

void CGrid32Mgr::SetSelectionNumberFormat(UINT format)
{
    ForEachFormattedCell(GetNormalizedSelection(), [&](GRIDCELL& cell) {
        cell.m_nFormat = format;
        if (cell.m_bFormula)
            RefreshFormulaCell(&cell);
        else
            RenderNumberFormat(cell);
    });
}

void CGrid32Mgr::SetSelectionBackground(COLORREF clr)
{
    SetRangeBackground(GetNormalizedSelection(), clr);
}

void CGrid32Mgr::SetRangeBackground(const GRIDSELECTION& sel, COLORREF clr)
{
    ForEachFormattedCell(sel, [&](GRIDCELL& cell) { cell.clrBackground = clr; });
}

bool CGrid32Mgr::GetCurrentCellFontInfo(CELLFONT& out)
{
    out = GetCellOrDefault(m_currentCell.nRow, m_currentCell.nCol)->fontInfo;
    return true;
}

// Apply an alignment-axis update across the current selection. axis selects
// which DT_* bits get cleared before lParam's value is OR'd in.
static void ApplyAlignBits(UINT& just, UINT axis, UINT value)
{
    switch (axis)
    {
    case GA_HORIZ:
        just &= ~(DT_LEFT | DT_CENTER | DT_RIGHT);
        just |= (value & (DT_LEFT | DT_CENTER | DT_RIGHT));
        break;
    case GA_VERT:
        just &= ~(DT_TOP | DT_VCENTER | DT_BOTTOM);
        just |= (value & (DT_TOP | DT_VCENTER | DT_BOTTOM));
        break;
    case GA_WRAP:
        if (value) just |= DT_WORDBREAK;
        else       just &= ~DT_WORDBREAK;
        break;
    }
}

void CGrid32Mgr::SetSelectionAlignBits(UINT axis, UINT value)
{
    ForEachFormattedCell(GetNormalizedSelection(),
        [&](GRIDCELL& cell) { ApplyAlignBits(cell.justification, axis, value); });
}

void CGrid32Mgr::SetSelectionHAlign(UINT halign)
{
    SetSelectionAlignBits(GA_HORIZ, halign);
}

void CGrid32Mgr::SetSelectionVAlign(UINT valign)
{
    SetSelectionAlignBits(GA_VERT, valign);
}

void CGrid32Mgr::SetSelectionWrap(BOOL wrap)
{
    SetSelectionAlignBits(GA_WRAP, wrap ? 1 : 0);
}

// ---- scrolling ---------------------------------------------------------------

// The scroll bars move the view, not the active cell. Their line arrows
// used to move the current cell, which also let a click on one move it out
// from under an open editor (AUDIT_2026-10 S2).
void CGrid32Mgr::OnHScroll(UINT nSBCode, UINT nPos, HWND hScrollBar)
{
    UNREFERENCED_PARAMETER(hScrollBar);
    PAGESTAT pageStat;
    CalculatePageStats(pageStat);
    const long long page = std::max<long long>(1, (long long)pageStat.end.nCol - pageStat.start.nCol);
    long long left = m_visibleTopLeft.nCol;
    switch (nSBCode)
    {
    case SB_LINELEFT:  left -= 1; break;
    case SB_LINERIGHT: left += 1; break;
    case SB_PAGELEFT:  left -= page; break;
    case SB_PAGERIGHT: left += page; break;
    case SB_THUMBTRACK:
    case SB_THUMBPOSITION:
    {
        // nPos is only 16 bits; the 32-bit track position comes from the bar.
        SCROLLINFO si{};
        si.cbSize = sizeof(SCROLLINFO);
        si.fMask = SIF_TRACKPOS;
        left = GetScrollInfo(m_hWndGrid, SB_HORZ, &si) ? si.nTrackPos : (long long)nPos;
        break;
    }
    case SB_LEFT:  left = 0; break;
    case SB_RIGHT: left = (long long)gcs.nWidth - 1; break;
    default: return;
    }
    if (left < 0) left = 0;
    if (left >= (long long)gcs.nWidth) left = (long long)gcs.nWidth - 1;
    if ((UINT)left != m_visibleTopLeft.nCol)
    {
        m_visibleTopLeft.nCol = (UINT)left;
        SetScrollRanges();
        Invalidate();
    }
}

void CGrid32Mgr::OnVScroll(UINT nSBCode, UINT nPos, HWND hScrollBar)
{
    UNREFERENCED_PARAMETER(hScrollBar);
    PAGESTAT pageStat;
    CalculatePageStats(pageStat);
    const long page = std::max<long>(1, (long)pageStat.end.nRow - (long)pageStat.start.nRow);
    UINT top = m_visibleTopLeft.nRow;
    switch (nSBCode)
    {
    case SB_LINEUP:   top = StepVisibleRow(top, -1); break;
    case SB_LINEDOWN: top = StepVisibleRow(top, 1); break;
    case SB_PAGEUP:   top = StepVisibleRow(top, -page); break;
    case SB_PAGEDOWN: top = StepVisibleRow(top, page); break;
    case SB_THUMBTRACK:
    case SB_THUMBPOSITION:
    {
        SCROLLINFO si{};
        si.cbSize = sizeof(SCROLLINFO);
        si.fMask = SIF_TRACKPOS;
        long long pos = GetScrollInfo(m_hWndGrid, SB_VERT, &si) ? si.nTrackPos : (long long)nPos;
        if (pos < 0) pos = 0;
        if (pos >= (long long)gcs.nHeight) pos = (long long)gcs.nHeight - 1;
        top = (UINT)pos;
        break;
    }
    case SB_TOP:    top = 0; break;
    case SB_BOTTOM: top = (UINT)gcs.nHeight - 1; break;
    default: return;
    }
    if (top != m_visibleTopLeft.nRow)
    {
        m_visibleTopLeft.nRow = top;
        SetScrollRanges();
        Invalidate();
    }
}

void CGrid32Mgr::OnMouseWheel(int zDelta, UINT nFlags, int x, int y)
{
    UNREFERENCED_PARAMETER(nFlags); UNREFERENCED_PARAMETER(x); UNREFERENCED_PARAMETER(y);
    UINT lines = 3;
    SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, &lines, 0);
    if (lines == 0)
        return;
    PAGESTAT pageStat;
    CalculatePageStats(pageStat);
    long rows = (lines == WHEEL_PAGESCROLL)
        ? std::max<long>(1, (long)pageStat.end.nRow - (long)pageStat.start.nRow)
        : (long)lines;
    long steps = -(zDelta * rows) / WHEEL_DELTA;
    if (steps == 0)
        steps = zDelta > 0 ? -1 : 1;
    UINT top = StepVisibleRow(m_visibleTopLeft.nRow, steps);
    if (top != m_visibleTopLeft.nRow)
    {
        m_visibleTopLeft.nRow = top;
        SetScrollRanges();
        Invalidate();
    }
}


void CGrid32Mgr::SetScrollRanges()
{
    PAGESTAT pageStat;
    CalculatePageStats(pageStat);

    if (gcs.style & WS_HSCROLL)
    {
        SCROLLINFO siH{};
        siH.cbSize = sizeof(SCROLLINFO);
        siH.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
        siH.nMin = 0;
        siH.nMax = static_cast<int>(gcs.nWidth - 1);
        siH.nPage = pageStat.end.nCol - pageStat.start.nCol + 1;
        siH.nPos = m_visibleTopLeft.nCol;
        SetScrollInfo(m_hWndGrid, SB_HORZ, &siH, TRUE);
    }

    if (gcs.style & WS_VSCROLL)
    {
        SCROLLINFO siV{};
        siV.cbSize = sizeof(SCROLLINFO);
        siV.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
        siV.nMin = 0;
        siV.nMax = static_cast<int>(gcs.nHeight - 1);
        siV.nPage = pageStat.end.nRow - pageStat.start.nRow + 1;
        siV.nPos = m_visibleTopLeft.nRow;
        SetScrollInfo(m_hWndGrid, SB_VERT, &siV, TRUE);
    }
}

// ---- navigation --------------------------------------------------------------

// Move delta rows from row, stepping over rows a filter hid. Stops at the
// last visible row in that direction.
UINT CGrid32Mgr::StepVisibleRow(UINT row, long delta)
{
    long long r = row;
    const int dir = delta > 0 ? 1 : -1;
    long steps = delta > 0 ? delta : -delta;
    while (steps-- > 0)
    {
        long long next = r + dir;
        while (next >= 0 && next < (long long)gcs.nHeight && IsRowHidden((UINT)next))
            next += dir;
        if (next < 0 || next >= (long long)gcs.nHeight)
            break;
        r = next;
    }
    return (UINT)r;
}

// Scroll just enough to bring the cell fully into view.
void CGrid32Mgr::EnsureCellVisible(UINT nRow, UINT nCol)
{
    if (nRow >= gcs.nHeight || nCol >= gcs.nWidth)
        return;
    const long long availW = (long long)m_clientRect.right - GetActualRowHeaderWidth();
    const long long availH = (long long)m_clientRect.bottom - GetActualColHeaderHeight();
    const GRIDPOINT before = m_visibleTopLeft;

    if (nCol < m_visibleTopLeft.nCol)
        m_visibleTopLeft.nCol = nCol;
    else if (availW > 0 && (long long)CalculatedColumnDistance(m_visibleTopLeft.nCol, (size_t)nCol + 1) > availW)
    {
        // The leftmost column that still lets nCol fit entirely (nCol itself
        // when it is wider than the view).
        UINT left = nCol;
        long long w = (long long)ColW(nCol);
        while (left > 0 && w + (long long)ColW(left - 1) <= availW)
            w += (long long)ColW(--left);
        m_visibleTopLeft.nCol = left;
    }

    if (nRow < m_visibleTopLeft.nRow)
        m_visibleTopLeft.nRow = nRow;
    else if (availH > 0 && (long long)CalculatedRowDistance(m_visibleTopLeft.nRow, (size_t)nRow + 1) > availH)
    {
        UINT top = nRow;
        long long h = (long long)RowH(nRow);
        while (top > 0 && h + (long long)RowH(top - 1) <= availH)
            h += (long long)RowH(--top);
        m_visibleTopLeft.nRow = top;
    }

    if (before.nRow != m_visibleTopLeft.nRow || before.nCol != m_visibleTopLeft.nCol)
    {
        SetScrollRanges();
        Invalidate();
    }
}

void CGrid32Mgr::ClampTopLeft()
{
    if (gcs.nHeight && m_visibleTopLeft.nRow >= gcs.nHeight)
        m_visibleTopLeft.nRow = (UINT)gcs.nHeight - 1;
    if (gcs.nWidth && m_visibleTopLeft.nCol >= gcs.nWidth)
        m_visibleTopLeft.nCol = (UINT)gcs.nWidth - 1;
}

void CGrid32Mgr::NotifySelChange()
{
    GRIDNMHDR h{};
    h.cellPt = m_currentCell;
    h.range = GetNormalizedSelection();
    SendGridNotification(GN_SELCHANGE, &h);
}

// The one place the active cell and the selection move from the keyboard or
// the API, so the selection can never be left behind on a cell the user has
// moved away from (AUDIT_2026-10 S1), and GN_SELCHANGE always reports a range
// that contains cellPt.
void CGrid32Mgr::MoveCurrentCell(UINT nRow, UINT nCol, bool bExtend)
{
    if (gcs.nWidth == 0 || gcs.nHeight == 0)
        return;
    if (nRow >= gcs.nHeight) nRow = (UINT)gcs.nHeight - 1;
    if (nCol >= gcs.nWidth) nCol = (UINT)gcs.nWidth - 1;

    // Moving away commits whatever is being typed, into its own cell.
    if (m_bEditing)
        CommitEdit();

    const GRIDPOINT prevCell = m_currentCell;
    const GRIDSELECTION prevSel = m_selectionRect;

    if (bExtend)
    {
        // Shift+move: the anchor (the active cell) stays, the far corner moves.
        m_selectionRect.start = m_currentCell;
        m_selectionRect.end = { nRow, nCol };
    }
    else
    {
        GRIDPOINT target{ nRow, nCol };
        if (const MERGERANGE* m = FindMerge(nRow, nCol))
            target = m->start;
        m_currentCell = target;
        m_selectionRect.start = m_selectionRect.end = target;
    }

    EnsureCellVisible(nRow, nCol);

    const bool moved = !AreGridPointsEqual(prevCell, m_currentCell) ||
        !AreGridPointsEqual(prevSel.start, m_selectionRect.start) ||
        !AreGridPointsEqual(prevSel.end, m_selectionRect.end);
    if (moved)
    {
        Invalidate();
        NotifySelChange();
    }
}

void CGrid32Mgr::IncrementSelectedCell(long nRow, long nCol, bool bExtend)
{
    GRIDPOINT base = bExtend ? m_selectionRect.end : m_currentCell;
    if (!bExtend)
    {
        // Leaving a merged block moves from its far edge, not its anchor.
        if (const MERGERANGE* m = FindMerge(base.nRow, base.nCol))
        {
            if (nRow > 0) base.nRow = m->end.nRow;
            if (nCol > 0) base.nCol = m->end.nCol;
        }
    }
    UINT row = nRow ? StepVisibleRow(base.nRow, nRow) : base.nRow;
    long long col = (long long)base.nCol + nCol;
    if (col < 0) col = 0;
    if (col >= (long long)gcs.nWidth) col = (long long)gcs.nWidth - 1;
    MoveCurrentCell(row, (UINT)col, bExtend);
}

// Page Up / Page Down: move the view and the cell by one screenful, so the
// cell keeps its place on screen.
void CGrid32Mgr::IncrementSelectedPage(long nRow, long nCol, bool bExtend)
{
    PAGESTAT pageStat;
    CalculatePageStats(pageStat);
    const long rowsPerPage = std::max<long>(1, (long)pageStat.end.nRow - (long)pageStat.start.nRow + 1);
    const long colsPerPage = std::max<long>(1, (long)pageStat.end.nCol - (long)pageStat.start.nCol + 1);

    GRIDPOINT base = bExtend ? m_selectionRect.end : m_currentCell;
    UINT row = nRow ? StepVisibleRow(base.nRow, nRow * rowsPerPage) : base.nRow;
    long long col = (long long)base.nCol + (long long)nCol * colsPerPage;
    if (col < 0) col = 0;
    if (col >= (long long)gcs.nWidth) col = (long long)gcs.nWidth - 1;

    if (nRow)
        m_visibleTopLeft.nRow = StepVisibleRow(m_visibleTopLeft.nRow, nRow * rowsPerPage);
    if (nCol)
    {
        long long left = (long long)m_visibleTopLeft.nCol + (long long)nCol * colsPerPage;
        m_visibleTopLeft.nCol = (UINT)std::max<long long>(0, std::min<long long>(left, (long long)gcs.nWidth - 1));
    }
    SetScrollRanges();
    Invalidate();
    MoveCurrentCell(row, (UINT)col, bExtend);
}

// Ctrl+arrow: to the first or last row / column of the grid.
void CGrid32Mgr::IncrementSelectEdge(long nRow, long nCol, bool bExtend)
{
    if (gcs.nWidth == 0 || gcs.nHeight == 0)
        return;
    GRIDPOINT base = bExtend ? m_selectionRect.end : m_currentCell;
    UINT row = base.nRow, col = base.nCol;
    if (nRow < 0)
        row = IsRowHidden(0) ? StepVisibleRow(0, 1) : 0;
    else if (nRow > 0)
    {
        row = (UINT)gcs.nHeight - 1;
        if (IsRowHidden(row))
            row = StepVisibleRow(row, -1);
    }
    if (nCol < 0)
        col = 0;
    else if (nCol > 0)
        col = (UINT)gcs.nWidth - 1;
    MoveCurrentCell(row, col, bExtend);
}

void CGrid32Mgr::OnMove(int x, int y)
{
    UNREFERENCED_PARAMETER(x);
    UNREFERENCED_PARAMETER(y);
}

void CGrid32Mgr::OnSize(UINT nType, int cx, int cy)
{
    UNREFERENCED_PARAMETER(nType);
    UNREFERENCED_PARAMETER(cx);
    UNREFERENCED_PARAMETER(cy);
    // The client rect first: the scroll bars' page size is computed from it.
    GetClientRect(m_hWndGrid, &m_clientRect);
    m_paintRect = m_clientRect;
    SetScrollRanges();
    if (m_bEditing)
        ShowEditControl();      // keep the editor over its cell
}

bool CGrid32Mgr::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
    UNREFERENCED_PARAMETER(nRepCnt);
    UNREFERENCED_PARAMETER(nFlags);
    // GetKeyState, not GetAsyncKeyState: the modifier state that belongs to
    // this message, not whatever the keyboard is doing now.
    const bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;

    long dRow = 0, dCol = 0;
    switch (nChar)
    {
    case VK_LEFT:  dCol = -1; break;
    case VK_RIGHT: dCol = 1;  break;
    case VK_UP:    dRow = -1; break;
    case VK_DOWN:  dRow = 1;  break;
    }
    if (dRow || dCol)
    {
        // Shift extends the selection (S1); Ctrl jumps to the edge.
        if (ctrl)
            IncrementSelectEdge(dRow, dCol, shift);
        else
            IncrementSelectedCell(dRow, dCol, shift);
        return true;
    }

    switch (nChar)
    {
    case VK_PRIOR:
        IncrementSelectedPage(-1, 0, shift);
        return true;
    case VK_NEXT:
        IncrementSelectedPage(1, 0, shift);
        return true;
    case VK_HOME:
    {
        GRIDPOINT base = shift ? m_selectionRect.end : m_currentCell;
        MoveCurrentCell(ctrl ? 0 : base.nRow, 0, shift);
        return true;
    }
    case VK_END:
        if (ctrl)
        {
            GRIDSELECTION used{};
            if (OnGetUsedRange(&used))
                MoveCurrentCell(used.end.nRow, used.end.nCol, shift);
            return true;
        }
        break;
    case VK_TAB:
        // Tab from the grid (no edit open) navigates like Excel: right by
        // default, left with Shift held.
        IncrementSelectedCell(0, shift ? -1 : 1);
        return true;
    case VK_DELETE:
        // Clears the whole selection, not just the active cell.
        OnClear();
        return true;
    case VK_F2:
        BeginEdit(true);
        return true;
    case VK_RETURN:
        // Nothing to commit; the host already had NM_RETURN. Enter used to
        // open the editor (iswspace(VK_RETURN)) and commit an untouched
        // cell, dirtying the document (AUDIT_2026-10 S11).
        return true;
    }

    // Standard editing shortcuts, for hosts without an accelerator table.
    if (ctrl)
    {
        switch (nChar)
        {
        case 'A': OnSetSelection(nullptr); return true;
        case 'C': OnCopy();  return true;
        case 'X': OnCut();   return true;
        case 'V': OnPaste(); return true;
        case 'Z': OnUndo();  return true;
        case 'Y': OnRedo();  return true;
        }
    }

    // Printable input arrives as WM_CHAR (OnChar), where the keyboard layout,
    // dead keys and the IME have already produced the real character. It
    // used to be guessed from the virtual-key code here, so F1 (0x70, 'p')
    // opened the editor (H14).
    return false;
}

bool CGrid32Mgr::OnKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags)
{
    // Handle key releases here
    return false;
}

// Typing over a cell opens the editor empty and hands it the character, as
// Excel does; F2 or a double-click edits the existing content instead.
bool CGrid32Mgr::OnChar(WCHAR ch)
{
    if (m_bEditing)
    {
        // A character queued to the grid before focus moved to the editor
        // (the second half of a surrogate pair, a fast typist) belongs there.
        if (ch >= 0x20 || ch == L'\b')
            SendMessage(m_hWndEdit, WM_CHAR, ch, 1);
        return true;
    }
    if (ch == L'\b')
    {
        // Backspace clears the cell into an empty editor; Escape keeps it.
        BeginEdit(false);
        return true;
    }
    if (ch < 0x20 || ch == 0x7F)
        return false;   // Enter, Tab, Escape and Ctrl+letter combinations
    BeginEdit(false);
    SendMessage(m_hWndEdit, WM_CHAR, ch, 1);
    return true;
}

void CGrid32Mgr::GetCellByPoint(int x, int y, GRIDPOINT& gridPt)
{
    if (x < GetActualRowHeaderWidth() || x >= m_clientRect.right || gcs.nWidth == 0)
        gridPt.nCol = ~0;
    else
    {
        size_t nWidth = (size_t)GetActualRowHeaderWidth();
        size_t idx = m_visibleTopLeft.nCol;
        if (idx >= (size_t)gcs.nWidth)
            idx = m_visibleTopLeft.nCol = 0;
        while (idx + 1 < (size_t)gcs.nWidth && nWidth + ColW(idx) < (size_t)x)
        {
            nWidth += ColW(idx);
            ++idx;
        }
        gridPt.nCol = (UINT)idx;
    }

    if (y < GetActualColHeaderHeight() || y >= m_clientRect.bottom || gcs.nHeight == 0)
    {
        gridPt.nRow = ~0;
    }
    else
    {
        size_t nHeight = (size_t)GetActualColHeaderHeight();
        size_t idx = m_visibleTopLeft.nRow;
        if (idx >= (size_t)gcs.nHeight)
            idx = m_visibleTopLeft.nRow = 0;
        while (idx + 1 < (size_t)gcs.nHeight && nHeight + RowH(idx) < (size_t)y)
        {
            nHeight += RowH(idx);
            ++idx;
        }
        gridPt.nRow = (UINT)idx;
    }
}

void CGrid32Mgr::StartMouseTracking()
{
    TRACKMOUSEEVENT tme;
    memset(&tme, 0, sizeof(TRACKMOUSEEVENT));
    tme.cbSize = sizeof(TRACKMOUSEEVENT);
    tme.dwFlags = TME_HOVER | TME_LEAVE;
    tme.hwndTrack = m_hWndGrid;
    tme.dwHoverTime = 100;
    TrackMouseEvent(&tme);
}

void CGrid32Mgr::CancelDragState()
{
    // Called when capture is lost or a drag is canceled by external means
    // (Alt+Tab, modal dialog, WM_CANCELMODE, debugger break, etc.).
    m_bSelecting = FALSE;
    m_bSizing = FALSE;
    m_gridHitTest = 0;
    m_mouseDraggingStartPoint = { ~0, ~0 };
    ClipCursor(NULL);
    CancelMouseTracking();
}

void CGrid32Mgr::CancelMouseTracking()
{
    TRACKMOUSEEVENT tme;
    memset(&tme, 0, sizeof(TRACKMOUSEEVENT));
    tme.cbSize = sizeof(TRACKMOUSEEVENT);
    tme.dwFlags = TME_CANCEL | TME_HOVER | TME_LEAVE;
    tme.hwndTrack = m_hWndGrid;
    TrackMouseEvent(&tme);
}

// Drag-selecting past an edge of the view: grow the selection by one cell in
// that direction and scroll it into view.
void CGrid32Mgr::CalcSelectionCoordinatesWithMouse(int x, int y, bool& bChangeX, bool& bChangeY)
{
    bChangeX = false, bChangeY = false;
    GRIDPOINT end = m_selectionRect.end;
    GRIDPOINT gridPt;
    GetCellByPoint(x, y, gridPt);

    if (x <= GetActualRowHeaderWidth())
    {
        if (end.nCol > 0) { --end.nCol; bChangeX = true; }
    }
    else if (x >= m_clientRect.right)
    {
        if (end.nCol + 1 < gcs.nWidth) { ++end.nCol; bChangeX = true; }
    }
    else if (gridPt.nCol != (UINT)~0 && gridPt.nCol != end.nCol)
    {
        end.nCol = gridPt.nCol;
        bChangeX = true;
    }

    if (y <= GetActualColHeaderHeight())
    {
        UINT r = StepVisibleRow(end.nRow, -1);
        if (r != end.nRow) { end.nRow = r; bChangeY = true; }
    }
    else if (y >= m_clientRect.bottom)
    {
        UINT r = StepVisibleRow(end.nRow, 1);
        if (r != end.nRow) { end.nRow = r; bChangeY = true; }
    }
    else if (gridPt.nRow != (UINT)~0 && gridPt.nRow != end.nRow)
    {
        end.nRow = gridPt.nRow;
        bChangeY = true;
    }

    if (bChangeX || bChangeY)
    {
        m_selectionRect.end = end;
        EnsureCellVisible(end.nRow, end.nCol);
    }
}

void CGrid32Mgr::GetCurrentMouseCoordinates(int& x, int& y)
{
    POINT pt;
    GetCursorPos(&pt);
    ScreenToClient(m_hWndGrid, &pt);

    //Get the mouse coordinates ourselves to fix bug with negative coordinates
    x = pt.x;
    y = pt.y;
}

void CGrid32Mgr::OnLButtonDown(UINT nFlags, int x, int y)
{
    GetCurrentMouseCoordinates(x, y);

    // An open edit commits into the cell it was opened on before the click
    // moves anything (AUDIT_2026-10 S2), and the grid takes focus so the
    // keyboard follows the click.
    if (m_bEditing)
        CommitEdit();
    SetFocus(m_hWndGrid);

    switch (m_gridHitTest)
    {
    case GRID_HTCELL:
        if (!m_bSelecting && !m_bSizing)
        {
            GRIDPOINT gridPt;
            GetCellByPoint(x, y, gridPt);
            if (gridPt.nRow == (UINT)~0 || gridPt.nCol == (UINT)~0)
                break;
            if (nFlags & MK_SHIFT)
            {
                // Shift+click extends from the active cell.
                m_selectionRect.start = m_currentCell;
                m_selectionRect.end = gridPt;
            }
            else
            {
                if (const MERGERANGE* m = FindMerge(gridPt.nRow, gridPt.nCol))
                    gridPt = m->start;
                m_currentCell = gridPt;
                m_selectionRect.start = m_selectionRect.end = gridPt;
            }
            EnsureCellVisible(gridPt.nRow, gridPt.nCol);
            m_mouseDraggingStartPoint.x = x;
            m_mouseDraggingStartPoint.y = y;

            SetCapture(m_hWndGrid);
            m_bSelecting = TRUE;
            StartMouseTracking();
            Invalidate();
        }
        break;
    case GRID_HTCOLDIVIDER:
    case GRID_HTROWDIVIDER:
        if (m_bResizable && !m_bSelecting)
        {
            SetCapture(m_hWndGrid);
            m_bSizing = TRUE;
            // Remember pre-resize size for the resize-completion notification.
            if (m_gridHitTest == GRID_HTCOLDIVIDER && m_nToBeSized < gcs.nWidth)
                m_resizeStartSize = ColW(m_nToBeSized);
            else if (m_gridHitTest == GRID_HTROWDIVIDER && m_nToBeSized < gcs.nHeight)
                m_resizeStartSize = RowH(m_nToBeSized);
        }
        break;
    case GRID_HTCOLHEADER:
    case GRID_HTROWHEADER:
    case GRID_HTCORNER:
        {
            GRIDPOINT cellPt{};
            GetCellByPoint(x, y, cellPt);
            GRIDNMHDR h{};
            h.pt = { x, y };
            h.cellPt = cellPt;
            INT code = m_gridHitTest == GRID_HTCOLHEADER ? GN_COLHEADERCLICK
                     : m_gridHitTest == GRID_HTROWHEADER ? GN_ROWHEADERCLICK
                     :                                    GN_CORNERCLICK;
            SendGridNotification(code, &h);
        }
        break;
    }
}

// The class has CS_DBLCLKS, so the second click of a double-click arrives
// here instead of as WM_LBUTTONDOWN; NM_DBLCLK never fired while the grid
// tried to time clicks itself. Unless the host handles NM_DBLCLK (returns
// nonzero), a double-click edits the cell, as F2 does.
void CGrid32Mgr::OnLButtonDblClk(UINT nFlags, int x, int y)
{
    UNREFERENCED_PARAMETER(nFlags);
    GetCurrentMouseCoordinates(x, y);
    if (m_gridHitTest != GRID_HTCELL)
        return;
    GRIDNMHDR h{};
    h.pt = { x, y };
    GetCellByPoint(x, y, h.cellPt);
    if (h.cellPt.nRow == (UINT)~0 || h.cellPt.nCol == (UINT)~0)
        return;
    if (SendGridNotification(NM_DBLCLK, &h) == 0)
        BeginEdit(true);
}

void CGrid32Mgr::OnLButtonUp(UINT nFlags, int x, int y)
{
    GetCurrentMouseCoordinates(x, y);

    // Track which event types to fire after the existing logic completes, so
    // SendGridNotification calls happen on a stable selection rect / sizes.
    bool fireSelChange = false;
    bool fireColResize = false, fireRowResize = false;
    size_t resizeOld = m_resizeStartSize;
    size_t resizeIdx = m_nToBeSized;

    if (m_bSelecting)
    {
        CancelMouseTracking();
        Invalidate();
        fireSelChange = true;
    }
    else if (m_bResizable && m_bSizing)
    {
        if (m_gridHitTest == GRID_HTCOLDIVIDER && m_nToBeSized < gcs.nWidth)
        {
            long long nNewSize = x - GetColumnStartPos((UINT)m_nToBeSized);
            if (nNewSize > 0)
            {
                SetColumnWidth(m_nToBeSized, (size_t)nNewSize);
                fireColResize = true;
            }
            Invalidate();
        }
        else if (m_gridHitTest == GRID_HTROWDIVIDER && m_nToBeSized < gcs.nHeight)
        {
            long long nNewSize = y - GetRowStartPos((UINT)m_nToBeSized);
            if (nNewSize > 0)
            {
                SetRowHeight(m_nToBeSized, (size_t)nNewSize);
                fireRowResize = true;
            }
            Invalidate();
        }
    }

    // NM_CLICK fires after every left-button release, even if the gesture was a
    // drag-select. cellPt is the released-on cell. Host code (e.g. Format
    // Painter) uses this as the "user picked a target" hook.
    bool wasCellClick = (m_gridHitTest == GRID_HTCELL);

    m_bSelecting = m_bSizing = 0;
    ReleaseCapture();
    ClipCursor(NULL);
    m_gridHitTest = 0;
    m_mouseDraggingStartPoint = { ~0, ~0 };

    if (wasCellClick)
    {
        GRIDNMHDR h{};
        h.pt = { x, y };
        GetCellByPoint(x, y, h.cellPt);
        SendGridNotification(NM_CLICK, &h);
    }
    if (fireSelChange)
        NotifySelChange();
    if (fireColResize)
    {
        GRIDNMHDR h{};
        h.cellPt = { 0, (UINT)resizeIdx };
        h.oldSize = resizeOld;
        h.newSize = ColW(resizeIdx);
        SendGridNotification(GN_COLWIDTHCHANGED, &h);
    }
    if (fireRowResize)
    {
        GRIDNMHDR h{};
        h.cellPt = { (UINT)resizeIdx, 0 };
        h.oldSize = resizeOld;
        h.newSize = RowH(resizeIdx);
        SendGridNotification(GN_ROWHEIGHTCHANGED, &h);
    }
}


void CGrid32Mgr::OnMouseMove(UINT nFlags, int x, int y)
{
    GetCurrentMouseCoordinates(x, y);

    if (m_bResizable && m_bSizing)
    {
        if (m_gridHitTest == GRID_HTCOLDIVIDER)
        {
            Invalidate();
            m_nSizingLine = x;
        }
        else if (m_gridHitTest == GRID_HTROWDIVIDER)
        {
            Invalidate();
            m_nSizingLine = y;
        }
    }
    else if (m_bSelecting)
    {
        bool bChangeX = false, bChangeY = false;
        if (x > GetActualRowHeaderWidth() && x < m_clientRect.right &&
            y > GetActualColHeaderHeight() && y < m_clientRect.bottom)
        {
            GRIDPOINT gridPt;
            GetCellByPoint(x, y, gridPt);
            if (gridPt.nRow != (UINT)~0 && gridPt.nCol != (UINT)~0)
                m_selectionRect.end = gridPt;
        }
        else
        {
            CalcSelectionCoordinatesWithMouse(x, y, bChangeX, bChangeY);
        }
        Invalidate();
    }
    else
    {
        GetCellByPoint(x, y, m_HoverCell);
        if (!m_npHoverDelaySet)
        {
            SendGridNotification(NM_HOVER);
            m_npHoverDelaySet = SetTimer(m_hWndGrid, HOVER_TIMER, m_nMouseHoverDelay, NULL);
        }
    }
}

void CGrid32Mgr::OnMouseHover(UINT nFlags, int x, int y)
{
    GetCurrentMouseCoordinates(x, y);

    bool bChangeX = false, bChangeY = false;
    if (m_bSelecting)
    {
        CalcSelectionCoordinatesWithMouse(x, y, bChangeX, bChangeY);
        StartMouseTracking();
        if (bChangeX || bChangeY)
            Invalidate();
    }
}

void CGrid32Mgr::OnMouseLeave(UINT nFlags, int x, int y)
{
    OnMouseHover(nFlags, x, y);
}

void CGrid32Mgr::OnNcMouseMove(UINT nFlags, int x, int y)
{
    GetCurrentMouseCoordinates(x, y);
}

LRESULT CGrid32Mgr::OnNcHitTest(UINT nFlags, int x, int y, bool& bHitTested)
{
    if (m_bSelecting || m_bSizing)
    {
        bHitTested = true;
        return m_gridHitTest;
    }

    if (IsCoordinateInClientArea(x, y))
    {
        // Convert screen coordinates to client coordinates
        POINT pt = { x, y };
        ScreenToClient(m_hWndGrid, &pt);

        long rowHeaderWidth = GetActualRowHeaderWidth();
        long colHeaderHeight = GetActualColHeaderHeight();

        // Check if the point is in the corner cell
        if (pt.x < rowHeaderWidth + 2 && pt.y < colHeaderHeight + 2)
        {
            bHitTested = true;
            return GRID_HTCORNER;
        }

        // Check if the point is in the column header
        if (pt.y < colHeaderHeight)
        {
            bHitTested = true;
            return IsOverColumnDivider(pt.x, pt.y) ?
                GRID_HTCOLDIVIDER : GRID_HTCOLHEADER;
        }

        // Check if the point is in the row header
        if (pt.x < rowHeaderWidth)
        {
            bHitTested = true;
            return IsOverRowDivider(pt.x, pt.y) ?
                GRID_HTROWDIVIDER : GRID_HTROWHEADER;
        }

        // Check if the point is over a column divider
        if (IsOverColumnDivider(pt.x, pt.y))
        {
            bHitTested = true;
            return GRID_HTCOLDIVIDER;
        }

        // Check if the point is over a row divider
        if (IsOverRowDivider(pt.x, pt.y))
        {
            bHitTested = true;
            return GRID_HTROWDIVIDER;
        }

        // Otherwise, the point is over a cell
        bHitTested = true;
        return GRID_HTCELL;
    }

    bHitTested = false;
    return GRID_HTOUTSIDE;
}

// A right-click outside the selection moves the active cell there first, so
// the context menu's Cut/Copy/Paste act on the cell under the mouse.
void CGrid32Mgr::OnRButtonDown(UINT nFlags, int x, int y)
{
    UNREFERENCED_PARAMETER(nFlags);
    GetCurrentMouseCoordinates(x, y);
    if (m_bEditing)
        CommitEdit();
    SetFocus(m_hWndGrid);

    GRIDNMHDR h{};
    h.pt = { x, y };
    h.cellPt = m_currentCell;
    if (m_gridHitTest == GRID_HTCELL)
    {
        GRIDPOINT pt;
        GetCellByPoint(x, y, pt);
        if (pt.nRow != (UINT)~0 && pt.nCol != (UINT)~0)
        {
            if (!IsCellPointInsideSelection(pt))
                MoveCurrentCell(pt.nRow, pt.nCol, false);
            h.cellPt = pt;
        }
    }
    h.range = GetNormalizedSelection();
    SendGridNotification(NM_RCLICK, &h);
}

void CGrid32Mgr::OnRButtonUp(UINT nFlags, int x, int y)
{
    // Handle right mouse button up events here
    // This is often used in conjunction with WM_RBUTTONDOWN for drag-and-drop operations
}

bool CGrid32Mgr::IsCoordinateInClientArea(int x, int y)
{
    RECT r;
    GetClientRect(m_hWndGrid, &r);

    // Convert screen coordinates to client coordinates
    POINT pt = { x, y };
    ScreenToClient(m_hWndGrid, &pt);

    // Check if the converted coordinates are within the client area
    return (r.left <= pt.x && pt.x <= r.right && r.top <= pt.y && pt.y <= r.bottom);
}

void CGrid32Mgr::SetCursorBasedOnNcHitTest()
{
    switch (m_gridHitTest)
    {
    case GRID_HTCOLDIVIDER:
        SetCursor(LoadCursor(nullptr, IDC_SIZEWE));
        break;
    case GRID_HTROWDIVIDER:
        SetCursor(LoadCursor(nullptr, IDC_SIZENS));
        break;
    default:
        SetCursor(LoadCursor(nullptr, IDC_ARROW));
        break;
    }
}

bool CGrid32Mgr::IsOverColumnDivider(int x, int y)
{
    UNREFERENCED_PARAMETER(y);
    long long pos = GetActualRowHeaderWidth();
    for (size_t idx = m_visibleTopLeft.nCol; idx < gcs.nWidth && pos < m_clientRect.right && pos < (long long)x + 5; ++idx)
    {
        pos += (long long)ColW(idx);
        if (pos >= (long long)x - 1 && pos <= (long long)x + 1)
        {
            m_nToBeSized = idx;
            return true;
        }
    }
    m_nToBeSized = ~0;
    return false;
}

bool CGrid32Mgr::IsOverRowDivider(int x, int y)
{
    UNREFERENCED_PARAMETER(x);
    long long pos = GetActualColHeaderHeight();
    for (size_t idx = m_visibleTopLeft.nRow; idx < gcs.nHeight && pos < m_clientRect.bottom && pos < (long long)y + 5; ++idx)
    {
        const size_t h = RowH(idx);
        if (h == 0)
            continue;   // a filtered-out row has no divider of its own
        pos += (long long)h;
        if (pos >= (long long)y - 1 && pos <= (long long)y + 1)
        {
            m_nToBeSized = idx;
            return true;
        }
    }
    m_nToBeSized = ~0;
    return false;
}

// ---- selection ---------------------------------------------------------------

// Set the selection. The far corner is clamped to the grid, so "to the end"
// can be written as (UINT)-1 (H4); a selection whose near corner is off the
// grid is rejected. A null selection selects everything. The active cell
// stays put if it lies inside the new selection.
bool CGrid32Mgr::OnSetSelection(const GRIDSELECTION* pGridSel)
{
    if (gcs.nWidth == 0 || gcs.nHeight == 0)
    {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
        return false;
    }
    GRIDSELECTION sel{};
    if (pGridSel)
        sel = *pGridSel;
    else
        sel = { { 0, 0 }, { (UINT)gcs.nHeight - 1, (UINT)gcs.nWidth - 1 } };

    if (min(sel.start.nRow, sel.end.nRow) >= gcs.nHeight || min(sel.start.nCol, sel.end.nCol) >= gcs.nWidth)
    {
        SetLastError(GRID_ERROR_INVALID_CELL_REFERENCE);
        return false;
    }
    auto clamp = [&](GRIDPOINT& pt) {
        if (pt.nRow >= gcs.nHeight) pt.nRow = (UINT)gcs.nHeight - 1;
        if (pt.nCol >= gcs.nWidth) pt.nCol = (UINT)gcs.nWidth - 1;
    };
    clamp(sel.start);
    clamp(sel.end);

    if (m_bEditing)
        CommitEdit();
    m_selectionRect = sel;
    if (!IsCellPointInsideSelection(m_currentCell))
    {
        GRIDSELECTION norm = sel;
        NormalizeSelectionRect(norm);
        m_currentCell = norm.start;
        if (const MERGERANGE* m = FindMerge(m_currentCell.nRow, m_currentCell.nCol))
            m_currentCell = m->start;
    }
    EnsureCellVisible(m_currentCell.nRow, m_currentCell.nCol);
    Invalidate();
    SetLastError(0);
    return true;
}

// The selection as a top-left / bottom-right rectangle, clamped to the grid
// and grown to include any merged block it cuts through.
void CGrid32Mgr::OnGetSelection(GRIDSELECTION* pGridSel)
{
    if (pGridSel) {
        *pGridSel = GetNormalizedSelection();
        SetLastError(0);
    }
    else {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
    }
}

GRIDSELECTION CGrid32Mgr::GetNormalizedSelection()
{
    GRIDSELECTION sel = m_selectionRect;
    NormalizeSelectionRect(sel);
    if (gcs.nHeight && sel.end.nRow >= gcs.nHeight) sel.end.nRow = (UINT)gcs.nHeight - 1;
    if (gcs.nWidth && sel.end.nCol >= gcs.nWidth) sel.end.nCol = (UINT)gcs.nWidth - 1;
    if (sel.start.nRow > sel.end.nRow) sel.start.nRow = sel.end.nRow;
    if (sel.start.nCol > sel.end.nCol) sel.start.nCol = sel.end.nCol;
    ExpandForMerges(sel);
    return sel;
}

// Accepts "B7", "$B$7", "b7" and ranges "A1:C3" (either corner first). The
// letters are the bijective base-26 column (A = 0, Z = 25, AA = 26) and the
// number is 1-based, so "A1" is (0, 0). Anything else, or a reference past
// the grid's edge, is rejected rather than silently mapped somewhere.
bool CGrid32Mgr::ParseCellReference(LPCWSTR pwszRef, GRIDSELECTION& out)
{
    if (!pwszRef)
        return false;
    std::wstring ref = pwszRef;
    const std::wstring ws = L" \t\r\n";
    size_t b = ref.find_first_not_of(ws), e = ref.find_last_not_of(ws);
    if (b == std::wstring::npos)
        return false;
    ref = ref.substr(b, e - b + 1);

    auto parseOne = [&](const std::wstring& s, GRIDPOINT& pt) -> bool {
        size_t i = 0;
        if (i < s.size() && s[i] == L'$') ++i;
        unsigned long long col = 0;
        size_t letters = 0;
        while (i < s.size() && iswalpha(s[i]))
        {
            const wchar_t up = (wchar_t)towupper(s[i]);
            if (up < L'A' || up > L'Z')
                return false;
            col = col * 26 + (up - L'A' + 1);
            if (++letters > 7 || col > UINT_MAX)
                return false;
            ++i;
        }
        if (letters == 0)
            return false;
        if (i < s.size() && s[i] == L'$') ++i;
        unsigned long long row = 0;
        size_t digits = 0;
        while (i < s.size() && iswdigit(s[i]))
        {
            row = row * 10 + (s[i] - L'0');
            if (++digits > 10 || row > UINT_MAX)
                return false;
            ++i;
        }
        if (digits == 0 || i != s.size() || row == 0)
            return false;
        if (col - 1 >= gcs.nWidth || row - 1 >= gcs.nHeight)
            return false;
        pt = { (UINT)(row - 1), (UINT)(col - 1) };
        return true;
    };

    size_t colon = ref.find(L':');
    if (colon == std::wstring::npos)
    {
        GRIDPOINT pt;
        if (!parseOne(ref, pt))
            return false;
        out.start = out.end = pt;
        return true;
    }
    GRIDPOINT a, z;
    if (!parseOne(ref.substr(0, colon), a) || !parseOne(ref.substr(colon + 1), z))
        return false;
    out.start = a;
    out.end = z;
    return true;
}

// Go To. SETBYNAME looks a cell up by the name given to it; SETBYCOORDINATE
// takes a reference ParseCellReference accepts. A range reference selects
// the range with the active cell at its first corner. It used to build the
// row and column 1-based against 0-based keys, mangle "$A$1", only find cells
// that already existed, and return TRUE when it found nothing (S16).
BOOL CGrid32Mgr::OnSetCurrentCell(LPCWSTR pwszRef, short nWhich)
{
    if (!pwszRef || !*pwszRef)
    {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
        return FALSE;
    }
    if (nWhich == SETBYNAME)
    {
        for (const auto& cellEntry : mapCells)
        {
            PGRIDCELL pCell = cellEntry.second;
            if (pCell && !pCell->m_wsName.empty() && _wcsicmp(pCell->m_wsName.c_str(), pwszRef) == 0)
            {
                MoveCurrentCell(cellEntry.first.first, cellEntry.first.second, false);
                SetLastError(0);
                return TRUE;
            }
        }
        SetLastError(GRID_ERROR_NOT_EXIST);
        return FALSE;
    }
    if (nWhich == SETBYCOORDINATE)
    {
        GRIDSELECTION sel{};
        if (!ParseCellReference(pwszRef, sel))
        {
            SetLastError(GRID_ERROR_INVALID_CELL_REFERENCE);
            return FALSE;
        }
        MoveCurrentCell(sel.start.nRow, sel.start.nCol, false);
        if (!AreGridPointsEqual(sel.start, sel.end))
        {
            m_selectionRect = sel;
            EnsureCellVisible(sel.start.nRow, sel.start.nCol);
            Invalidate();
            NotifySelChange();
        }
        SetLastError(0);
        return TRUE;
    }
    SetLastError(GRID_ERROR_INVALID_PARAMETER);
    return FALSE;
}

BOOL CGrid32Mgr::OnSetCurrentCell(UINT nRow, UINT nCol)
{
    return SetCurrentCell(nRow, nCol) ? TRUE : FALSE;
}

bool CGrid32Mgr::SetCurrentCell(UINT nRow, UINT nCol)
{
    if (nRow >= gcs.nHeight || nCol >= gcs.nWidth)
    {
        SetLastError(GRID_ERROR_OUT_OF_RANGE);
        return false;
    }
    MoveCurrentCell(nRow, nCol, false);
    SetLastError(0);
    return true;
}

bool CGrid32Mgr::AreGridPointsEqual(const GRIDPOINT& point1, const GRIDPOINT& point2)
{
    return (point1.nRow == point2.nRow && point1.nCol == point2.nCol);
}

bool CGrid32Mgr::IsCellPointInsideSelection(const GRIDPOINT& gridPt)
{
    GRIDSELECTION sel = GetNormalizedSelection();
    return (gridPt.nRow >= sel.start.nRow && gridPt.nRow <= sel.end.nRow &&
        gridPt.nCol >= sel.start.nCol && gridPt.nCol <= sel.end.nCol);
}

// True when the selection is just the active cell (or the merged block the
// active cell anchors).
bool CGrid32Mgr::IsSelectionUnicellular()
{
    GRIDSELECTION sel = GetNormalizedSelection();
    GRIDSELECTION cur{ m_currentCell, m_currentCell };
    ExpandForMerges(cur);
    return AreGridPointsEqual(sel.start, cur.start) && AreGridPointsEqual(sel.end, cur.end);
}

void CGrid32Mgr::ConstrainMaxPosition()
{
    if (gcs.nWidth && m_selectionRect.end.nCol >= gcs.nWidth)
        m_selectionRect.end.nCol = (UINT)(gcs.nWidth - 1);
    if (gcs.nHeight && m_selectionRect.end.nRow >= gcs.nHeight)
        m_selectionRect.end.nRow = (UINT)(gcs.nHeight - 1);
    ClampTopLeft();
}


void CGrid32Mgr::NormalizeSelectionRect(GRIDSELECTION& selection)
{
    if (selection.start.nRow > selection.end.nRow)
        std::swap(selection.start.nRow, selection.end.nRow);
    if (selection.start.nCol > selection.end.nCol)
        std::swap(selection.start.nCol, selection.end.nCol);
}

void CGrid32Mgr::OnIncrementCell(UINT direction, GRIDPOINT* pGridPoint)
{
    if (!pGridPoint) return;
    // GRIDPOINT carries UINT fields, but the Increment* APIs take signed
    // long deltas. The host signals negative direction by writing a UINT
    // whose bit-pattern is a negative int (e.g. (UINT)-1). Cast through
    // int first so sign-extension into long is explicit and portable.
    long nRow = (long)(int)pGridPoint->nRow;
    long nCol = (long)(int)pGridPoint->nCol;
    switch (direction)
    {
    case INCREMENT_SINGLE:
        IncrementSelectedCell(nRow, nCol);
        break;
    case INCREMENT_PAGE:
        IncrementSelectedPage(nRow, nCol);
        break;
    case INCREMENT_EDGE:
        IncrementSelectEdge(nRow, nCol);
        break;
    }
}

// ---- inline editor -----------------------------------------------------------

namespace {
// The text a cell is edited as: a formula's source, or a formatted number's
// plain value ("$1,234.50" edits as "1234.5"), else the text itself.
std::wstring EditTextOf(const GRIDCELL& cell)
{
    if (cell.m_bFormula)
        return L"=" + cell.m_wsFormula;
    if (cell.m_eType != CT_Text && cell.m_nFormat != FMT_GENERAL)
        return Grid32Detail::FormatEditValue(cell.m_eType, cell.m_dValue, cell.m_wsText);
    return cell.m_wsText;
}
}

void CGrid32Mgr::BeginEdit(bool bKeepText)
{
    if (m_bEditing)
        CommitEdit();

    // The editor always opens on the active cell, which is a merged block's
    // anchor when the cell is merged.
    m_editCell = m_currentCell;
    PGRIDCELL pCell = GetCell(m_editCell.nRow, m_editCell.nCol);
    // What the cell holds now. A commit that leaves it unchanged writes
    // nothing, so opening and closing the editor never dirties the document.
    m_editOriginal = pCell ? EditTextOf(*pCell) : std::wstring();
    const std::wstring text = bKeepText ? m_editOriginal : std::wstring();

    SetWindowText(m_hWndEdit, text.c_str());
    SendMessage(m_hWndEdit, EM_SETSEL, (WPARAM)text.length(), (LPARAM)text.length());
    m_bEditing = true;

    GRIDNMHDR h{};
    h.cellPt = m_editCell;
    SendGridNotification(GN_EDITSTART, &h);
    ShowEditControl();
}

// Position the editor over its cell, in the cell's font, and give it focus.
// Safe to call again (on resize) while an edit is open.
void CGrid32Mgr::ShowEditControl()
{
    if (!m_bEditing)
        return;

    EnsureCellVisible(m_editCell.nRow, m_editCell.nCol);
    RECT cellRect{};
    GetCellRect(m_editCell.nRow, m_editCell.nCol, cellRect);
    InflateRect(&cellRect, -1, -1);

    // Edit caret position follows the cell's vertical alignment so the
    // user's text appears where the committed text will render: the edit is
    // sized to one line of its font and slid within the cell.
    PGRIDCELL pCell = GetCellOrDefault(m_editCell.nRow, m_editCell.nCol);
    const int cellH = cellRect.bottom - cellRect.top;
    int editHeight = cellH;

    HDC hdc = ::GetDC(m_hWndGrid);
    if (hdc) {
        const int logPixelsY = GetDeviceCaps(hdc, LOGPIXELSY);
        HFONT hFont = GetCachedFont(pCell->fontInfo, CellFontHeight(pCell->fontInfo.m_fPointSize, logPixelsY));
        if (hFont)
        {
            if (hFont != m_hEditFont)
            {
                SendMessage(m_hWndEdit, WM_SETFONT, (WPARAM)hFont, FALSE);
                m_hEditFont = hFont;
            }
            // The font's real line height (ascent + descent), not its em
            // size, which clipped descenders at larger sizes (S14).
            HGDIOBJ hOld = SelectObject(hdc, hFont);
            TEXTMETRICW tm{};
            if (GetTextMetricsW(hdc, &tm))
                editHeight = tm.tmHeight + tm.tmExternalLeading + 4;
            SelectObject(hdc, hOld);
        }
        ::ReleaseDC(m_hWndGrid, hdc);
    }
    if (editHeight > cellH)
        editHeight = cellH;

    const UINT vAlign = pCell->justification & (DT_TOP | DT_VCENTER | DT_BOTTOM);
    int top = cellRect.top;
    switch (vAlign)
    {
    case DT_VCENTER: top = cellRect.top + (cellH - editHeight) / 2; break;
    case DT_BOTTOM:  top = cellRect.bottom - editHeight;            break;
    case DT_TOP:
    default:         top = cellRect.top;                            break;
    }

    MoveWindow(m_hWndEdit, cellRect.left, top, cellRect.right - cellRect.left, editHeight, TRUE);
    BringWindowToTop(m_hWndEdit);
    SendMessage(m_hWndEdit, EM_SCROLLCARET, 0, 0);
    ShowWindow(m_hWndEdit, SW_SHOW);
    SetFocus(m_hWndEdit);
}

namespace {
// Hide the editor and hand focus back to the grid if the editor had it. The
// caller has already cleared m_bEditing, so the WM_KILLFOCUS this causes
// finds no open edit and commits nothing.
void HideEditor(HWND hWndEdit, HWND hWndGrid, HFONT hDefaultFont)
{
    const bool hadFocus = GetFocus() == hWndEdit;
    ShowWindow(hWndEdit, SW_HIDE);
    SendMessage(hWndEdit, WM_SETFONT, (WPARAM)hDefaultFont, FALSE);
    if (hadFocus)
        SetFocus(hWndGrid);
}
}

bool CGrid32Mgr::CommitEdit()
{
    if (!m_bEditing)
        return false;
    std::wstring text;
    const int len = GetWindowTextLength(m_hWndEdit);
    if (len > 0)
    {
        text.resize((size_t)len + 1);
        const int got = GetWindowText(m_hWndEdit, &text[0], len + 1);
        text.resize(got > 0 ? (size_t)got : 0);
    }
    CommitEditText(text);
    return true;
}

// Write text into the cell the editor was opened on — never "whatever is
// current now", which a click or a scroll could have changed while the
// editor still had focus (AUDIT_2026-10 S2).
void CGrid32Mgr::CommitEditText(const std::wstring& text)
{
    if (!m_bEditing)
        return;
    // Cleared first: anything below that moves focus must find no edit open.
    m_bEditing = false;
    const GRIDPOINT cell = m_editCell;
    HideEditor(m_hWndEdit, m_hWndGrid, m_hDefaultFont);
    m_hEditFont = NULL;

    const bool changed = text != m_editOriginal;
    if (changed)
        SetCellText(cell.nRow, cell.nCol, text.c_str());

    GRIDNMHDR h{};
    h.cellPt = cell;
    h.range = { cell, cell };
    h.text = text.c_str();
    SendGridNotification(GN_EDITCOMMIT, &h);
    if (changed)
        SendGridNotification(GN_CELLCHANGED, &h);
    InvalidateCell(cell.nRow, cell.nCol);
}

// Escape. The edit is closed before GN_EDITCANCEL goes out, so a handler that
// moves focus can't commit the text being thrown away, and nothing is left
// armed to swallow the next real commit (AUDIT_2026-10 S12).
void CGrid32Mgr::CancelEdit()
{
    if (!m_bEditing)
        return;
    m_bEditing = false;
    const GRIDPOINT cell = m_editCell;
    HideEditor(m_hWndEdit, m_hWndGrid, m_hDefaultFont);
    m_hEditFont = NULL;

    GRIDNMHDR h{};
    h.cellPt = cell;
    SendGridNotification(GN_EDITCANCEL, &h);
    InvalidateCell(cell.nRow, cell.nCol);
}

// ---- clipboard ---------------------------------------------------------------

namespace {
// Excel's convention for a text cell on the clipboard: a field holding a tab,
// a line break or a leading quote is wrapped in quotes, with quotes doubled.
std::wstring QuoteClipField(const std::wstring& s)
{
    if (s.find_first_of(L"\t\r\n") == std::wstring::npos && (s.empty() || s[0] != L'"'))
        return s;
    std::wstring out = L"\"";
    for (wchar_t ch : s)
    {
        if (ch == L'"')
            out += L'"';
        out += ch;
    }
    out += L'"';
    return out;
}

void AppendClipField(std::wstring& out, const wchar_t* tag, const std::wstring& value)
{
    out += tag;
    out += L':';
    out += std::to_wstring(value.size());
    out += L':';
    out += value;
    out += L';';
}

// Doubles and floats travel as their bit patterns: exact, and immune to the
// locale's decimal separator.
std::wstring BitsOf(double d)
{
    unsigned long long bits = 0;
    memcpy(&bits, &d, sizeof(bits));
    wchar_t buf[17];
    swprintf_s(buf, 17, L"%016llx", bits);
    return buf;
}

double DoubleFromBits(const std::wstring& s)
{
    unsigned long long bits = wcstoull(s.c_str(), nullptr, 16);
    double d = 0;
    memcpy(&d, &bits, sizeof(d));
    return d;
}

bool PutClipboardString(UINT format, const std::wstring& data)
{
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, (data.size() + 1) * sizeof(wchar_t));
    if (!hMem)
        return false;
    LPWSTR p = (LPWSTR)GlobalLock(hMem);
    if (!p)
    {
        GlobalFree(hMem);
        return false;
    }
    memcpy(p, data.c_str(), (data.size() + 1) * sizeof(wchar_t));
    GlobalUnlock(hMem);
    if (!SetClipboardData(format, hMem))
    {
        // SetClipboardData failed — ownership not transferred, free.
        GlobalFree(hMem);
        return false;
    }
    return true;
}

bool GetClipboardString(UINT format, std::wstring& data)
{
    HANDLE hData = GetClipboardData(format);
    if (!hData)
        return false;
    LPCWSTR p = (LPCWSTR)GlobalLock(hData);
    if (!p)
        return false;
    // Bounded by the block's size: clipboard data from another program is
    // not guaranteed to be terminated.
    const size_t maxChars = GlobalSize(hData) / sizeof(wchar_t);
    data.assign(p, wcsnlen(p, maxChars));
    GlobalUnlock(hData);
    return true;
}
}

// Tab-separated display text, CRLF between rows, as other programs expect.
std::wstring CGrid32Mgr::BuildCopyText(const GRIDSELECTION& selIn)
{
    GRIDSELECTION sel = selIn;
    NormalizeSelectionRect(sel);
    std::wstring data;
    for (UINT r = sel.start.nRow; r <= sel.end.nRow && r < gcs.nHeight; ++r)
    {
        if (IsRowHidden(r))
            continue;
        if (!data.empty())
            data += L"\r\n";
        for (UINT c = sel.start.nCol; c <= sel.end.nCol && c < gcs.nWidth; ++c)
        {
            if (PGRIDCELL cell = GetCell(r, c))
                data += QuoteClipField(cell->m_wsText);
            if (c < sel.end.nCol)
                data += L'\t';
        }
    }
    return data;
}

// The private format: every populated cell in the block with its formula
// source, typed value and formatting, at offsets from the block's corner.
std::wstring CGrid32Mgr::BuildCopyRecords(const GRIDSELECTION& selIn)
{
    GRIDSELECTION sel = selIn;
    NormalizeSelectionRect(sel);
    std::wstring out = L"G32C1\n";
    out += std::to_wstring(sel.end.nRow - sel.start.nRow + 1) + L"," +
           std::to_wstring(sel.end.nCol - sel.start.nCol + 1) + L"\n";
    for (const GRIDPOINT& pt : CellsInRange(sel))
    {
        const GRIDCELL& c = *GetCell(pt.nRow, pt.nCol);
        out += std::to_wstring(pt.nRow - sel.start.nRow) + L"," + std::to_wstring(pt.nCol - sel.start.nCol) + L";";
        AppendClipField(out, L"T", c.m_wsText);
        if (c.m_bFormula)
            AppendClipField(out, L"F", c.m_wsFormula);
        AppendClipField(out, L"Y", std::to_wstring((int)c.m_eType));
        AppendClipField(out, L"V", BitsOf(c.m_dValue));
        AppendClipField(out, L"NF", std::to_wstring(c.m_nFormat));
        AppendClipField(out, L"FF", c.fontInfo.m_wsFontFace);
        AppendClipField(out, L"FS", BitsOf(c.fontInfo.m_fPointSize));
        AppendClipField(out, L"FC", std::to_wstring(c.fontInfo.m_clrTextColor));
        AppendClipField(out, L"WT", std::to_wstring(c.fontInfo.bWeight));
        AppendClipField(out, L"IT", std::to_wstring(c.fontInfo.bItalic ? 1 : 0));
        AppendClipField(out, L"UN", std::to_wstring(c.fontInfo.bUnderline ? 1 : 0));
        AppendClipField(out, L"ST", std::to_wstring(c.fontInfo.bStrikeThrough ? 1 : 0));
        AppendClipField(out, L"BG", std::to_wstring(c.clrBackground));
        AppendClipField(out, L"J", std::to_wstring(c.justification));
        out += L"\n";
    }
    return out;
}

void CGrid32Mgr::OnCopy()
{
    GRIDSELECTION sel = GetNormalizedSelection();
    const std::wstring text = BuildCopyText(sel);
    const std::wstring records = BuildCopyRecords(sel);

    if (OpenClipboard(m_hWndGrid))
    {
        EmptyClipboard();
        PutClipboardString(CF_UNICODETEXT, text);
        if (m_cfGridCells)
            PutClipboardString(m_cfGridCells, records);
        CloseClipboard();
    }
}

void CGrid32Mgr::OnCut()
{
    OnCopy();
    OnClear();
}

void CGrid32Mgr::OnPaste()
{
    std::wstring records, text;
    bool haveRecords = false, haveText = false;
    if (!OpenClipboard(m_hWndGrid))
        return;
    if (m_cfGridCells && IsClipboardFormatAvailable(m_cfGridCells))
        haveRecords = GetClipboardString(m_cfGridCells, records);
    if (!haveRecords)
        haveText = GetClipboardString(CF_UNICODETEXT, text);
    CloseClipboard();

    if (haveRecords && PasteRecords(records))
        return;
    if (haveText)
        PasteText(text);
}

// Paste tab-separated text at the selection's top-left corner. Lines end in
// CRLF, LF or CR — the clipboard's CRLF used to leave a '\r' on every row's
// last field, so "42\r" landed as text (B6). Quoted fields may hold tabs,
// line breaks and doubled quotes. A single value pasted over a larger
// selection fills it.
void CGrid32Mgr::PasteText(const std::wstring& data)
{
    std::vector<std::vector<std::wstring>> rows(1);
    std::wstring field;
    bool inQuotes = false, quotedField = false;
    for (size_t i = 0; i < data.size(); ++i)
    {
        const wchar_t ch = data[i];
        if (inQuotes)
        {
            if (ch == L'"')
            {
                if (i + 1 < data.size() && data[i + 1] == L'"') { field += L'"'; ++i; }
                else inQuotes = false;
            }
            else
                field += ch;
            continue;
        }
        if (ch == L'"' && field.empty() && !quotedField)
        {
            inQuotes = quotedField = true;
        }
        else if (ch == L'\t')
        {
            rows.back().push_back(field);
            field.clear();
            quotedField = false;
        }
        else if (ch == L'\r' || ch == L'\n')
        {
            if (ch == L'\r' && i + 1 < data.size() && data[i + 1] == L'\n')
                ++i;
            rows.back().push_back(field);
            field.clear();
            quotedField = false;
            rows.emplace_back();
        }
        else
            field += ch;
    }
    if (!field.empty() || quotedField || !rows.back().empty())
        rows.back().push_back(field);
    // The line break after the last row ends it; it doesn't start another.
    if (rows.size() > 1 && rows.back().empty())
        rows.pop_back();
    if (rows.empty() || (rows.size() == 1 && rows[0].empty()))
        return;

    GRIDSELECTION sel = GetNormalizedSelection();
    GRIDSELECTION target{ sel.start, sel.start };
    {
        BatchUndo batch(*this);
        BulkEditScope bulkEdit(m_bDeferRecalc, m_bSuppressChangeNotify);
        if (rows.size() == 1 && rows[0].size() == 1 && !IsWholeGrid(sel))
        {
            target = sel;
            for (UINT r = sel.start.nRow; r <= sel.end.nRow; ++r)
                for (UINT c = sel.start.nCol; c <= sel.end.nCol; ++c)
                    SetCellText(r, c, rows[0][0].c_str());
        }
        else
        {
            for (size_t dr = 0; dr < rows.size() && sel.start.nRow + dr < gcs.nHeight; ++dr)
            {
                for (size_t dc = 0; dc < rows[dr].size() && sel.start.nCol + dc < gcs.nWidth; ++dc)
                {
                    const UINT r = sel.start.nRow + (UINT)dr, c = sel.start.nCol + (UINT)dc;
                    if (rows[dr][dc].empty())
                        ClearCellText(r, c);
                    else
                        SetCellText(r, c, rows[dr][dc].c_str());
                    target.end.nRow = max(target.end.nRow, r);
                    target.end.nCol = max(target.end.nCol, c);
                }
            }
        }
    }
    RecalculateFormulas();
    NotifyContentChanged();
    m_selectionRect = target;
    m_currentCell = target.start;
    Invalidate();
    NotifySelChange();
}

// Paste the private format: formulas stay formulas and formatting comes
// along. Empty cells of the copied block clear their targets, as in Excel.
bool CGrid32Mgr::PasteRecords(const std::wstring& data)
{
    std::wistringstream in(data);
    std::wstring line;
    if (!std::getline(in, line) || line != L"G32C1")
        return false;
    if (!std::getline(in, line))
        return false;
    UINT nRows = 0, nCols = 0;
    if (swscanf_s(line.c_str(), L"%u,%u", &nRows, &nCols) != 2 || nRows == 0 || nCols == 0)
        return false;

    struct Record { UINT dr, dc; GRIDCELL cell; };
    std::vector<Record> records;
    while (std::getline(in, line))
    {
        if (line.empty())
            continue;
        UINT dr = 0, dc = 0;
        const size_t semi = line.find(L';');
        if (semi == std::wstring::npos || swscanf_s(line.substr(0, semi).c_str(), L"%u,%u", &dr, &dc) != 2)
            return false;
        Record rec{ dr, dc, m_defaultGridCell };
        GRIDCELL& c = rec.cell;
        c.mergeRange = MERGERANGE();
        size_t p = semi + 1;
        while (p < line.size())
        {
            const size_t tagEnd = line.find(L':', p);
            if (tagEnd == std::wstring::npos) return false;
            const size_t lenEnd = line.find(L':', tagEnd + 1);
            if (lenEnd == std::wstring::npos) return false;
            const std::wstring tag = line.substr(p, tagEnd - p);
            const size_t len = wcstoul(line.substr(tagEnd + 1, lenEnd - tagEnd - 1).c_str(), nullptr, 10);
            if (lenEnd + 1 + len > line.size()) return false;
            const std::wstring value = line.substr(lenEnd + 1, len);
            p = lenEnd + 1 + len;
            if (p < line.size() && line[p] == L';') ++p;

            if (tag == L"T") c.m_wsText = value;
            else if (tag == L"F") { c.m_bFormula = true; c.m_wsFormula = value; }
            else if (tag == L"Y") c.m_eType = (CellType)wcstol(value.c_str(), nullptr, 10);
            else if (tag == L"V") c.m_dValue = DoubleFromBits(value);
            else if (tag == L"NF") c.m_nFormat = (UINT)wcstoul(value.c_str(), nullptr, 10);
            else if (tag == L"FF") c.fontInfo.m_wsFontFace = value;
            else if (tag == L"FS") c.fontInfo.m_fPointSize = (float)DoubleFromBits(value);
            else if (tag == L"FC") c.fontInfo.m_clrTextColor = (COLORREF)wcstoul(value.c_str(), nullptr, 10);
            else if (tag == L"WT") c.fontInfo.bWeight = (UINT)wcstoul(value.c_str(), nullptr, 10);
            else if (tag == L"IT") c.fontInfo.bItalic = value == L"1";
            else if (tag == L"UN") c.fontInfo.bUnderline = value == L"1";
            else if (tag == L"ST") c.fontInfo.bStrikeThrough = value == L"1";
            else if (tag == L"BG") c.clrBackground = (COLORREF)wcstoul(value.c_str(), nullptr, 10);
            else if (tag == L"J") c.justification = (UINT)wcstoul(value.c_str(), nullptr, 10);
        }
        records.push_back(std::move(rec));
    }

    const GRIDSELECTION sel = GetNormalizedSelection();
    GRIDSELECTION target{ sel.start, { sel.start.nRow + nRows - 1, sel.start.nCol + nCols - 1 } };
    if (target.end.nRow >= gcs.nHeight) target.end.nRow = (UINT)gcs.nHeight - 1;
    if (target.end.nCol >= gcs.nWidth) target.end.nCol = (UINT)gcs.nWidth - 1;
    {
        BatchUndo batch(*this);
        BulkEditScope bulkEdit(m_bDeferRecalc, m_bSuppressChangeNotify);
        for (const GRIDPOINT& pt : CellsInRange(target))
            DeleteCell(pt.nRow, pt.nCol);
        for (Record& rec : records)
        {
            const unsigned long long r = (unsigned long long)sel.start.nRow + rec.dr;
            const unsigned long long c = (unsigned long long)sel.start.nCol + rec.dc;
            if (r >= gcs.nHeight || c >= gcs.nWidth)
                continue;
            SetCell((UINT)r, (UINT)c, rec.cell);
        }
    }
    RecalculateFormulas();
    NotifyContentChanged();
    m_selectionRect = target;
    m_currentCell = target.start;
    Invalidate();
    NotifySelChange();
    return true;
}

// Clear the contents of every cell in the selection, keeping formatting, as
// one undo step with one recalculation. Only cells that exist are visited.
// The ribbon's Clear used to go through GM_FILLCELLS, which blanked the text
// but kept formulas (they came back on the next recalc), created a cell for
// every empty position, and could not be undone (S4); WM_CLEAR's undo record
// had no coordinates at all (S17).
void CGrid32Mgr::OnClear()
{
    if (m_bEditing)
        CommitEdit();
    const std::vector<GRIDPOINT> cells = CellsInRange(GetNormalizedSelection());
    if (cells.empty())
        return;
    {
        BatchUndo batch(*this);
        BulkEditScope bulkEdit(m_bDeferRecalc, m_bSuppressChangeNotify);
        for (const GRIDPOINT& pt : cells)
            ClearCellText(pt.nRow, pt.nCol);
    }
    RecalculateFormulas();
    NotifyContentChanged();
    Invalidate();
    SetLastError(0);
}

// ---- undo --------------------------------------------------------------------

// Put back the cell snapshot an edit operation recorded. SetCell replaces the
// whole GRIDCELL, so the formula source, cell type, cached numeric value,
// number format and character formatting all come back together; a side
// recorded as "no cell" is restored by deleting the cell.
void CGrid32Mgr::RestoreCells(const GridEditOperation& op, bool bUndo, GRIDSELECTION* pAffected)
{
    bool any = false;
    GRIDSELECTION box{};
    auto grow = [&](UINT r, UINT c) {
        if (!any) { box.start = box.end = { r, c }; any = true; return; }
        box.start.nRow = min(box.start.nRow, r); box.start.nCol = min(box.start.nCol, c);
        box.end.nRow = max(box.end.nRow, r);     box.end.nCol = max(box.end.nCol, c);
    };
    {
        // Restoring must not itself be recorded, and a multi-cell restore is
        // one user action: one recalc and one change notification, not one
        // per cell.
        FlagScope noRecord(m_bUndoRecordEnabled, FALSE);
        BulkEditScope bulkEdit(m_bDeferRecalc, m_bSuppressChangeNotify);

        switch (op.type)
        {
        case EditOperationType::Batch:
            for (const CellChange& ch : op.changes)
            {
                if (bUndo ? ch.bHadOld : ch.bHasNew)
                    SetCell(ch.pt.nRow, ch.pt.nCol, bUndo ? ch.oldCell : ch.newCell);
                else
                    DeleteCell(ch.pt.nRow, ch.pt.nCol);
                grow(ch.pt.nRow, ch.pt.nCol);
            }
            if (op.bDefaultChanged)
            {
                m_defaultGridCell = bUndo ? op.oldDefault : op.newDefault;
                grow(0, 0);
                grow((UINT)gcs.nHeight - 1, (UINT)gcs.nWidth - 1);
            }
            break;
        case EditOperationType::Delete:
            if (bUndo)
                SetCell(op.row, op.col, op.oldState);
            else
                DeleteCell(op.row, op.col);
            grow(op.row, op.col);
            break;
        default:
            if (bUndo && op.bOldMissing)
                DeleteCell(op.row, op.col);
            else
                SetCell(op.row, op.col, bUndo ? op.oldState : op.newState);
            grow(op.row, op.col);
            break;
        }
    }
    m_bMergesDirty = true;

    RecalculateFormulas();
    NotifyContentChanged();
    if (pAffected)
        *pAffected = box;
}

bool CGrid32Mgr::OnUndo(GRIDSELECTION* pAffected)
{
    // Text being typed is committed first, so undo takes back that edit.
    if (m_bEditing)
        CommitEdit();
    if (m_undoStack.empty())
        return false;

    // Copy the top op but don't pop yet — if the restore throws, we want
    // the undo stack to retain the entry so the user can retry rather
    // than silently losing the operation.
    GridEditOperation op = m_undoStack.top();
    GRIDSELECTION box{};
    try
    {
        RestoreCells(op, /*bUndo*/ true, &box);
    }
    catch (...)
    {
        SetLastError(GRID_ERROR_NOT_IMPLEMENTED);
        return false;
    }

    m_undoStack.pop();
    m_redoStack.push(op);
    // Move to what changed, so the user sees where the undo happened.
    if (!IsWholeGrid(box))
        MoveCurrentCell(box.start.nRow, box.start.nCol, false);
    if (pAffected)
        *pAffected = box;
    Invalidate();
    SetLastError(0);
    return true;
}

bool CGrid32Mgr::OnRedo(GRIDSELECTION* pAffected)
{
    if (m_bEditing)
        CommitEdit();
    if (m_redoStack.empty())
        return false;

    // Held until the restore succeeds, mirroring OnUndo.
    GridEditOperation op = m_redoStack.top();
    GRIDSELECTION box{};
    try
    {
        RestoreCells(op, /*bUndo*/ false, &box);
    }
    catch (...)
    {
        SetLastError(GRID_ERROR_NOT_IMPLEMENTED);
        return false;
    }

    m_redoStack.pop();
    m_undoStack.push(op);
    if (!IsWholeGrid(box))
        MoveCurrentCell(box.start.nRow, box.start.nCol, false);
    if (pAffected)
        *pAffected = box;
    Invalidate();
    SetLastError(0);
    return true;
}

size_t CGrid32Mgr::OnCanUndo()
{
    return m_undoStack.size();
}

size_t CGrid32Mgr::OnCanRedo()
{
    return m_redoStack.size();
}

void CGrid32Mgr::RecordUndoOperation(const GridEditOperation& op)
{
	if (!m_bUndoRecordEnabled)
		return; // If recording is disabled, do not record the operation
    m_undoStack.push(op);
    // Clear redo stack
    while (!m_redoStack.empty()) {
        m_redoStack.pop();
    }
}

// ---- commands ----------------------------------------------------------------

BOOL CGrid32Mgr::Invalidate(bool bErasebkgnd)
{
    if (m_bRedraw)
        return ::InvalidateRect(m_hWndGrid, NULL, bErasebkgnd);

    return 0;
}

// Repaint one cell (or its merged block), with room for the active-cell box.
void CGrid32Mgr::InvalidateCell(UINT nRow, UINT nCol)
{
    if (!m_bRedraw)
        return;
    RECT rect;
    if (GetCellRect(nRow, nCol, rect))
    {
        InflateRect(&rect, 3, 3);
        ::InvalidateRect(m_hWndGrid, &rect, FALSE);
    }
}

bool CGrid32Mgr::SetColumnWidth(size_t nCol, size_t newWidth)
{
    if (nCol < gcs.nWidth) {
        if (newWidth == gcs.nDefColWidth)
            m_colWidths.erase((UINT)nCol);
        else
            m_colWidths[(UINT)nCol] = newWidth;
        SetScrollRanges();
        Invalidate();
        return true;
    }

    SetLastError(GRID_ERROR_OUT_OF_RANGE);
    return false;
}

size_t CGrid32Mgr::GetColumnWidth(size_t nCol)
{
    if (nCol < gcs.nWidth)
        return ColW(nCol);

    SetLastError(GRID_ERROR_OUT_OF_RANGE);
    return 0;
}

// A hidden row keeps the height it is given and gets it back when the filter
// that hid it is cleared.
bool CGrid32Mgr::SetRowHeight(size_t nRow, size_t newHeight)
{
    if (nRow < gcs.nHeight) {
        if (newHeight == gcs.nDefRowHeight)
            m_rowHeights.erase((UINT)nRow);
        else
            m_rowHeights[(UINT)nRow] = newHeight;
        SetScrollRanges();
        Invalidate();
        return true;
    }

    SetLastError(GRID_ERROR_OUT_OF_RANGE);
    return false;
}

// The displayed height: 0 for a row a filter has hidden.
size_t CGrid32Mgr::GetRowHeight(size_t nRow)
{
    if (nRow < gcs.nHeight)
        return RowH(nRow);

    SetLastError(GRID_ERROR_OUT_OF_RANGE);
    return 0;
}

// An empty cell reads as an empty string: that is a successful read, not an
// error. It used to return TRUE and set GRID_ERROR_NOT_EXIST at once (M9).
BOOL CGrid32Mgr::OnGetCellText(const GRIDPOINT& point, GRID_GETTEXT* text) {
    if (!text || !text->wszBuff || text->nLen == 0 ||
        point.nRow >= gcs.nHeight || point.nCol >= gcs.nWidth) {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    PGRIDCELL pCell = GetCell(point.nRow, point.nCol);
    if (pCell) {
        if (text->nLen < pCell->m_wsText.length() + 1) {
            SetLastError(GRID_ERROR_INVALID_PARAMETER);
            return FALSE;
        }
        wcscpy_s(text->wszBuff, text->nLen, pCell->m_wsText.c_str());
        SetLastError(0);
        return TRUE;
    }

    text->wszBuff[0] = L'\0';
    SetLastError(0);
    return TRUE;
}

LRESULT CGrid32Mgr::SendGridNotification(INT code, GRIDNMHDR* pNMHDRInfo)
{
    GRIDNMHDR nmhdr;
    memset(&nmhdr, 0, sizeof(GRIDNMHDR));
    if (pNMHDRInfo)
        nmhdr = *pNMHDRInfo;
    nmhdr.hwndFrom = m_hWndGrid;
    nmhdr.idFrom = GetWindowLongPtr(m_hWndGrid, GWLP_ID);
    nmhdr.code = code;
    if (!pNMHDRInfo)
    {
        GetMousePosition(nmhdr.pt);
        nmhdr.cellPt = code == NM_HOVER ? m_HoverCell : m_currentCell;
    }

    HWND parentWnd = GetParent(m_hWndGrid);
    if (parentWnd)
        return SendMessage(parentWnd, WM_NOTIFY, nmhdr.idFrom, (LPARAM)&nmhdr);
    return 0;
}


void CGrid32Mgr::NotifyContentChanged()
{
    // Inside a bulk edit the caller sends one notification when the scope
    // closes, so individual cell writes stay quiet.
    if (m_bSuppressChangeNotify)
        return;
    // DeleteAllCells runs from the destructor, by which point the window may
    // already be going away; never notify through a dead HWND.
    if (!m_hWndGrid || !IsWindow(m_hWndGrid))
        return;
    SendGridNotification(GN_CONTENTCHANGED);
}

void CGrid32Mgr::GetMousePosition(POINT& pt)
{
    GetCursorPos(&pt);
    ScreenToClient(m_hWndGrid, &pt);
}


void CGrid32Mgr::OnTimer(UINT_PTR idEvent)
{
    if (idEvent == HOVER_TIMER)
    {
        KillTimer(m_hWndGrid, idEvent);
        m_npHoverDelaySet = 0;
    }
}

bool CGrid32Mgr::IsCellFontDefault(GRIDCELL& gc)
{
    const CELLFONT& cellFont = gc.fontInfo;
    const CELLFONT& defaultFont = m_defaultGridCell.fontInfo;

    return (cellFont.m_wsFontFace == defaultFont.m_wsFontFace &&
        cellFont.m_fPointSize == defaultFont.m_fPointSize &&
        cellFont.m_clrTextColor == defaultFont.m_clrTextColor &&
        cellFont.bItalic == defaultFont.bItalic &&
        cellFont.bUnderline == defaultFont.bUnderline &&
        cellFont.bStrikeThrough == defaultFont.bStrikeThrough &&
        cellFont.bWeight == defaultFont.bWeight);
}

// The positions of the populated cells, in row-major order. With a null
// array, how many there are. It used to copy GRIDCELLs, std::wstrings and
// all, into the host's memory.
size_t CGrid32Mgr::OnEnumCells(GRIDPOINT* pts, UINT numElements)
{
    if (!pts)
        return mapCells.size();
    size_t nCount = 0;
    for (auto it = mapCells.begin(); it != mapCells.end() && nCount < numElements; ++it, ++nCount)
        pts[nCount] = { it->first.first, it->first.second };
    return nCount;
}

namespace {
// The parts of a GCCELLINFO its mask names, applied to a cell's formatting.
void ApplyCellInfoFormat(GRIDCELL& cell, const GCCELLINFO& info)
{
    if (info.m_dwMask & GCI_FONT)
        cell.fontInfo = CellFontFrom(info.m_font);
    if (info.m_dwMask & GCI_BACKGROUND)
        cell.clrBackground = info.m_clrBackground;
    if (info.m_dwMask & GCI_ALIGN)
        cell.justification = info.m_nJustification;
    if (info.m_dwMask & GCI_FORMAT)
        cell.m_nFormat = info.m_nFormat;
}
}

BOOL CGrid32Mgr::OnSetCellInfo(const GRIDPOINT& pt, const GCCELLINFO& info)
{
    if (info.m_cbSize != sizeof(GCCELLINFO))
    {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
        return FALSE;
    }
    if (pt.nRow >= gcs.nHeight || pt.nCol >= gcs.nWidth)
    {
        SetLastError(GRID_ERROR_OUT_OF_RANGE);
        return FALSE;
    }
    {
        BatchUndo batch(*this);
        if (info.m_dwMask & ~GCI_TEXT)
        {
            PGRIDCELL p = GetCellOrCreate(pt.nRow, pt.nCol);
            if (!p)
                return FALSE;
            ApplyCellInfoFormat(*p, info);
            if (info.m_dwMask & GCI_FORMAT)
            {
                if (p->m_bFormula)
                    RefreshFormulaCell(p);
                else
                    RenderNumberFormat(*p);
            }
        }
        if (info.m_dwMask & GCI_TEXT)
            SetCellText(pt.nRow, pt.nCol, info.m_pwszText ? info.m_pwszText : L"");
    }
    InvalidateCell(pt.nRow, pt.nCol);
    NotifyContentChanged();
    SetLastError(0);
    return TRUE;
}

// The bounding box of every cell that exists (content or formatting) and of
// every merged block.
BOOL CGrid32Mgr::OnGetUsedRange(GRIDSELECTION* pRange)
{
    if (!pRange)
    {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
        return FALSE;
    }
    if (mapCells.empty())
    {
        SetLastError(0);
        return FALSE;
    }
    GRIDSELECTION box{ { UINT_MAX, UINT_MAX }, { 0, 0 } };
    for (const auto& entry : mapCells)
    {
        box.start.nRow = min(box.start.nRow, entry.first.first);
        box.start.nCol = min(box.start.nCol, entry.first.second);
        box.end.nRow = max(box.end.nRow, entry.first.first);
        box.end.nCol = max(box.end.nCol, entry.first.second);
    }
    RefreshMergeIndex();
    for (const MERGERANGE& m : m_merges)
    {
        box.end.nRow = max(box.end.nRow, m.end.nRow);
        box.end.nCol = max(box.end.nCol, m.end.nCol);
    }
    *pRange = box;
    SetLastError(0);
    return TRUE;
}

// Draw a block of cells into any DC at its own resolution: screen pixels
// (96 DPI) are scaled to the DC's LOGPIXELSX/Y. CGridView's printing and
// print preview paginate with GM_GETUSEDRANGE and the column widths, then
// call this once per page (H12).
BOOL CGrid32Mgr::OnRenderRange(const GCRENDERSTRUCT& rs)
{
    if (rs.m_cbSize != sizeof(GCRENDERSTRUCT) || !rs.m_hDC || gcs.nWidth == 0 || gcs.nHeight == 0)
    {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
        return FALSE;
    }
    GRIDSELECTION range = rs.m_range;
    NormalizeSelectionRect(range);
    if (range.start.nRow >= gcs.nHeight || range.start.nCol >= gcs.nWidth)
    {
        SetLastError(GRID_ERROR_OUT_OF_RANGE);
        return FALSE;
    }
    if (range.end.nRow >= gcs.nHeight) range.end.nRow = (UINT)gcs.nHeight - 1;
    if (range.end.nCol >= gcs.nWidth) range.end.nCol = (UINT)gcs.nWidth - 1;

    HDC hdc = rs.m_hDC;
    int dpiX = rs.m_nLogPixelsX > 0 ? rs.m_nLogPixelsX : GetDeviceCaps(hdc, LOGPIXELSX);
    int dpiY = rs.m_nLogPixelsY > 0 ? rs.m_nLogPixelsY : GetDeviceCaps(hdc, LOGPIXELSY);
    if (dpiX <= 0) dpiX = 96;
    if (dpiY <= 0) dpiY = 96;
    auto sx = [&](long long v) { return (long)(v * dpiX / 96); };
    auto sy = [&](long long v) { return (long)(v * dpiY / 96); };

    const long hdrW = rs.m_bHeaders ? sx(nRowHeaderWidth) : 0;
    const long hdrH = rs.m_bHeaders ? sy(nColHeaderHeight > 0 ? nColHeaderHeight : 20) : 0;
    std::vector<long> xs, ys;
    long x = rs.m_ptOrigin.x + hdrW, y = rs.m_ptOrigin.y + hdrH;
    xs.push_back(x);
    for (UINT c = range.start.nCol; c <= range.end.nCol; ++c)
        xs.push_back(x += sx((long long)ColW(c)));
    ys.push_back(y);
    for (UINT r = range.start.nRow; r <= range.end.nRow; ++r)
        ys.push_back(y += sy((long long)RowH(r)));

    // Rendering into a foreign DC must not use the screen paint's clipping.
    const RECT savedPaint = m_paintRect;
    m_paintRect = { LONG_MIN / 2, LONG_MIN / 2, LONG_MAX / 2, LONG_MAX / 2 };
    const int saved = SaveDC(hdc);

    if (rs.m_bGridLines)
    {
        HPEN hPen = CreatePen(PS_SOLID, 1, RGB(192, 192, 192));
        HGDIOBJ hOld = SelectObject(hdc, hPen);
        for (size_t i = 0; i < ys.size(); ++i)
        {
            if (i > 0 && ys[i] == ys[i - 1])
                continue;   // a hidden row
            MoveToEx(hdc, xs.front(), ys[i], NULL);
            LineTo(hdc, xs.back(), ys[i]);
        }
        for (long xe : xs)
        {
            MoveToEx(hdc, xe, ys.front(), NULL);
            LineTo(hdc, xe, ys.back());
        }
        SelectObject(hdc, hOld);
        DeleteObject(hPen);
    }

    auto cellRect = [&](GRIDPOINT a, GRIDPOINT z) {
        // Clip a block to the range being rendered.
        a.nRow = max(a.nRow, range.start.nRow); a.nCol = max(a.nCol, range.start.nCol);
        z.nRow = min(z.nRow, range.end.nRow);   z.nCol = min(z.nCol, range.end.nCol);
        return RECT{ xs[a.nCol - range.start.nCol], ys[a.nRow - range.start.nRow],
                     xs[z.nCol - range.start.nCol + 1], ys[z.nRow - range.start.nRow + 1] };
    };

    RefreshMergeIndex();
    for (const MERGERANGE& m : m_merges)
    {
        if (m.end.nRow < range.start.nRow || m.start.nRow > range.end.nRow ||
            m.end.nCol < range.start.nCol || m.start.nCol > range.end.nCol)
            continue;
        RECT rect = cellRect(m.start, m.end);
        DrawCellContent(hdc, *GetCellOrDefault(m.start.nRow, m.start.nCol), rect, dpiY, true);
    }
    for (const GRIDPOINT& pt : CellsInRange(range))
    {
        if (IsRowHidden(pt.nRow) || (!m_merges.empty() && FindMerge(pt.nRow, pt.nCol)))
            continue;
        RECT rect = cellRect(pt, pt);
        DrawCellContent(hdc, *GetCell(pt.nRow, pt.nCol), rect, dpiY, false);
    }

    if (rs.m_bHeaders)
    {
        HFONT hFont = GetCachedFont(m_cornerCell.fontInfo, CellFontHeight(10.0f, dpiY));
        HGDIOBJ hOldFont = SelectObject(hdc, hFont);
        HBRUSH hBrush = CreateSolidBrush(m_cornerCell.clrBackground);
        HPEN hPen = CreatePen(PS_SOLID, 1, RGB(128, 128, 128));
        HGDIOBJ hOldBrush = SelectObject(hdc, hBrush);
        HGDIOBJ hOldPen = SelectObject(hdc, hPen);
        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, m_cornerCell.fontInfo.m_clrTextColor);
        Rectangle(hdc, rs.m_ptOrigin.x, rs.m_ptOrigin.y, xs.front() + 1, ys.front() + 1);
        for (UINT c = range.start.nCol; c <= range.end.nCol; ++c)
        {
            RECT rect{ xs[c - range.start.nCol], rs.m_ptOrigin.y, xs[c - range.start.nCol + 1] + 1, ys.front() + 1 };
            Rectangle(hdc, rect.left, rect.top, rect.right, rect.bottom);
            std::wstring label = GetColumnLabel(c);
            DrawTextW(hdc, label.c_str(), -1, &rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
        }
        for (UINT r = range.start.nRow; r <= range.end.nRow; ++r)
        {
            if (IsRowHidden(r))
                continue;
            RECT rect{ rs.m_ptOrigin.x, ys[r - range.start.nRow], xs.front() + 1, ys[r - range.start.nRow + 1] + 1 };
            Rectangle(hdc, rect.left, rect.top, rect.right, rect.bottom);
            std::wstring label = std::to_wstring(r + 1);
            DrawTextW(hdc, label.c_str(), -1, &rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
        }
        SelectObject(hdc, hOldPen);
        SelectObject(hdc, hOldBrush);
        SelectObject(hdc, hOldFont);
        DeleteObject(hPen);
        DeleteObject(hBrush);
    }

    RestoreDC(hdc, saved);
    m_paintRect = savedPaint;
    SetLastError(0);
    return TRUE;
}

// ---- fill, sort, filter ------------------------------------------------------

// Copy the parts of src that m_dwFillWhat names into each target cell.
// Content goes through SetCellText, so a filled "42" is a number and a
// formula stays a formula (M12); the direction and range fields are honoured
// (M12, S7), and the whole fill is one undo step (M24). wParam is reserved
// and must be 0 (M13).
BOOL CGrid32Mgr::OnFillCells(WPARAM wParam, const GCFILLSTRUCT& fillStruct)
{
    if (fillStruct.m_cbSize != sizeof(GCFILLSTRUCT) || wParam != 0 || fillStruct.m_dwFillWhat == 0)
    {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    GRIDSELECTION range{};
    switch (fillStruct.m_dwFillRange)
    {
    case FILLBYRANGE:
        range = fillStruct.m_range;
        NormalizeSelectionRect(range);
        break;
    case FILLBYCURRENTSELECTION:
    case 0:
        range = GetNormalizedSelection();
        break;
    default:
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
        return FALSE;
    }
    if (range.end.nRow >= gcs.nHeight || range.end.nCol >= gcs.nWidth)
    {
        SetLastError(GRID_ERROR_OUT_OF_RANGE);
        return FALSE;
    }

    const DWORD what = fillStruct.m_dwFillWhat;
    DWORD error = 0;
    auto apply = [&](const GRIDCELL& src, UINT r, UINT c) {
        if (what & ~FILL_TEXT)
        {
            PGRIDCELL p = GetCellOrCreate(r, c);
            if (!p)
            {
                error = GetLastError();
                return;
            }
            if (what & FILL_COLOR)
                p->clrBackground = src.clrBackground;
            if (what & FILL_FONT)
                p->fontInfo.m_wsFontFace = src.fontInfo.m_wsFontFace;
            if (what & FILL_POINTSIZE)
                p->fontInfo.m_fPointSize = src.fontInfo.m_fPointSize;
            if (what & FILL_EFFECTS)
            {
                p->fontInfo.bItalic = src.fontInfo.bItalic;
                p->fontInfo.bUnderline = src.fontInfo.bUnderline;
                p->fontInfo.bStrikeThrough = src.fontInfo.bStrikeThrough;
                p->fontInfo.bWeight = src.fontInfo.bWeight;
            }
            if (what & FILL_FORMAT)
            {
                p->fontInfo.m_clrTextColor = src.fontInfo.m_clrTextColor;
                p->justification = src.justification;
                p->m_nFormat = src.m_nFormat;
                RenderNumberFormat(*p);
            }
        }
        // Content last, so it renders in the number format just copied.
        if (what & FILL_TEXT)
        {
            const std::wstring text = EditTextOf(src);
            if (text.empty())
                ClearCellText(r, c);
            else
                SetCellText(r, c, text.c_str());
        }
    };

    {
        BatchUndo batch(*this);
        BulkEditScope bulkEdit(m_bDeferRecalc, m_bSuppressChangeNotify);
        switch (fillStruct.m_dwFillDirection)
        {
        case FILL_DOWN:
            for (UINT c = range.start.nCol; c <= range.end.nCol && !error; ++c)
            {
                const GRIDCELL src = *GetCellOrDefault(range.start.nRow, c);
                for (UINT r = range.start.nRow + 1; r <= range.end.nRow && !error; ++r)
                    apply(src, r, c);
            }
            break;
        case FILL_UP:
            for (UINT c = range.start.nCol; c <= range.end.nCol && !error; ++c)
            {
                const GRIDCELL src = *GetCellOrDefault(range.end.nRow, c);
                for (UINT r = range.start.nRow; r < range.end.nRow && !error; ++r)
                    apply(src, r, c);
            }
            break;
        case FILL_RIGHT:
            for (UINT r = range.start.nRow; r <= range.end.nRow && !error; ++r)
            {
                const GRIDCELL src = *GetCellOrDefault(r, range.start.nCol);
                for (UINT c = range.start.nCol + 1; c <= range.end.nCol && !error; ++c)
                    apply(src, r, c);
            }
            break;
        case FILL_LEFT:
            for (UINT r = range.start.nRow; r <= range.end.nRow && !error; ++r)
            {
                const GRIDCELL src = *GetCellOrDefault(r, range.end.nCol);
                for (UINT c = range.start.nCol; c < range.end.nCol && !error; ++c)
                    apply(src, r, c);
            }
            break;
        case 0:
        {
            GRIDCELL src = m_defaultGridCell;
            if (fillStruct.m_dwFillHow == FILLBYGCS)
            {
                ApplyCellInfoFormat(src, fillStruct.m_cell);
                if ((fillStruct.m_cell.m_dwMask & GCI_TEXT) && fillStruct.m_cell.m_pwszText)
                    src.m_wsText = fillStruct.m_cell.m_pwszText;
            }
            else if (fillStruct.m_dwFillHow == FILLBYTEXT)
            {
                if (fillStruct.m_cell.m_pwszText)
                    src.m_wsText = fillStruct.m_cell.m_pwszText;
            }
            else if (fillStruct.m_dwFillHow == FILLBYREFCELL &&
                     fillStruct.m_refCell.nRow < gcs.nHeight && fillStruct.m_refCell.nCol < gcs.nWidth)
                src = *GetCellOrDefault(fillStruct.m_refCell.nRow, fillStruct.m_refCell.nCol);
            else
            {
                error = GRID_ERROR_INVALID_PARAMETER;
                break;
            }
            for (UINT r = range.start.nRow; r <= range.end.nRow && !error; ++r)
                for (UINT c = range.start.nCol; c <= range.end.nCol && !error; ++c)
                    apply(src, r, c);
            break;
        }
        default:
            error = GRID_ERROR_INVALID_PARAMETER;
            break;
        }
    }
    RecalculateFormulas();
    NotifyContentChanged();
    Invalidate();
    SetLastError(error);
    return error == 0;
}

// Sort the rows of m_sortSel by one key column. Only cells that exist are
// moved: it used to copy every column of every row, so sorting 100 rows of a
// 702-column grid left ~70,000 new empty cells behind (P3). m_dwWhich names
// the key and wParam carries GCSF_* flags (M13). Blanks sort last in both
// directions, numbers before text, text case-insensitively unless
// GCSF_MATCHCASE; equal keys keep their order. One undo step.
BOOL CGrid32Mgr::OnSortCells(WPARAM wParam, const GCSORTSTRUCT& sortStruct)
{
    if (sortStruct.m_cbSize != sizeof(GCSORTSTRUCT) ||
        (wParam & ~(WPARAM)(GCSF_SELECTIONONLY | GCSF_HEADERROW | GCSF_MATCHCASE)))
    {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
        return FALSE;
    }
    if (m_bEditing)
        CommitEdit();

    GRIDSELECTION sel = sortStruct.m_sortSel;
    NormalizeSelectionRect(sel);
    if (sel.start.nRow >= gcs.nHeight || sel.start.nCol >= gcs.nWidth)
    {
        SetLastError(GRID_ERROR_OUT_OF_RANGE);
        return FALSE;
    }
    if (sel.end.nRow >= gcs.nHeight) sel.end.nRow = (UINT)gcs.nHeight - 1;
    if (sel.end.nCol >= gcs.nWidth) sel.end.nCol = (UINT)gcs.nWidth - 1;

    const bool selectionOnly = (wParam & GCSF_SELECTIONONLY) != 0;
    const bool matchCase = (wParam & GCSF_MATCHCASE) != 0;
    const UINT firstRow = sel.start.nRow + ((wParam & GCSF_HEADERROW) ? 1 : 0);
    if (firstRow > sel.end.nRow)
    {
        SetLastError(0);
        return TRUE;    // nothing below the header
    }

    UINT keyCol;
    if (sortStruct.m_dwWhich)
        keyCol = (UINT)(sortStruct.m_dwWhich - 1);
    else if (m_currentCell.nCol >= sel.start.nCol && m_currentCell.nCol <= sel.end.nCol)
        keyCol = m_currentCell.nCol;
    else
        keyCol = sel.start.nCol;
    const UINT c0 = selectionOnly ? sel.start.nCol : 0;
    const UINT c1 = selectionOnly ? sel.end.nCol : (UINT)gcs.nWidth - 1;
    if (keyCol >= gcs.nWidth || keyCol < c0 || keyCol > c1)
    {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    // A merged block can't be split across the rows being reordered.
    RefreshMergeIndex();
    for (const MERGERANGE& m : m_merges)
    {
        if (!(m.end.nRow < firstRow || m.start.nRow > sel.end.nRow || m.end.nCol < c0 || m.start.nCol > c1))
        {
            SetLastError(GRID_ERROR_INVALID_PARAMETER);
            return FALSE;
        }
    }

    struct SortKey {
        bool isEmpty;
        bool isNumeric;
        double num;
        std::wstring text;
    };
    struct Row { UINT src; SortKey key; };
    std::vector<Row> rows;
    for (UINT r = firstRow; r <= sel.end.nRow; ++r)
    {
        Row row{ r, { true, false, 0.0, std::wstring() } };
        if (PGRIDCELL cell = GetCell(r, keyCol))
        {
            double n = 0.0;
            row.key.text = cell->m_wsText;
            row.key.num = cell->m_dValue;
            row.key.isEmpty = cell->m_wsText.empty() && !cell->m_bFormula && cell->m_eType == CT_Text;
            row.key.isNumeric = cell->m_eType == CT_Number || cell->m_eType == CT_Date ||
                cell->m_eType == CT_Boolean ||
                (cell->m_bFormula && std::isfinite(cell->m_dValue) &&
                 (cell->m_nFormat != FMT_GENERAL || Grid32Detail::TryParseNumber(cell->m_wsText, n)));
        }
        rows.push_back(std::move(row));
    }

    const bool asc = sortStruct.m_bAscending != FALSE;
    std::stable_sort(rows.begin(), rows.end(), [asc, matchCase](const Row& a, const Row& b) {
        const SortKey& ka = a.key;
        const SortKey& kb = b.key;
        if (ka.isEmpty != kb.isEmpty)
            return kb.isEmpty;      // blanks last, whichever the direction
        if (ka.isEmpty)
            return false;
        if (ka.isNumeric != kb.isNumeric)
            return asc ? ka.isNumeric : kb.isNumeric;
        if (ka.isNumeric)
            return asc ? ka.num < kb.num : ka.num > kb.num;
        const int cmp = matchCase ? wcscmp(ka.text.c_str(), kb.text.c_str())
                                  : _wcsicmp(ka.text.c_str(), kb.text.c_str());
        return asc ? cmp < 0 : cmp > 0;
    });

    bool inOrder = true;
    std::vector<UINT> newRowOf(sel.end.nRow - firstRow + 1);
    for (size_t i = 0; i < rows.size(); ++i)
    {
        newRowOf[rows[i].src - firstRow] = firstRow + (UINT)i;
        if (rows[i].src != firstRow + i)
            inOrder = false;
    }
    if (inOrder)
    {
        SetLastError(0);
        return TRUE;
    }

    const GRIDSELECTION block{ { firstRow, c0 }, { sel.end.nRow, c1 } };
    std::vector<std::pair<GRIDPOINT, GRIDCELL>> moved;
    for (const GRIDPOINT& pt : CellsInRange(block))
        moved.push_back({ pt, *GetCell(pt.nRow, pt.nCol) });
    {
        BatchUndo batch(*this);
        BulkEditScope bulkEdit(m_bDeferRecalc, m_bSuppressChangeNotify);
        for (const auto& entry : moved)
            DeleteCell(entry.first.nRow, entry.first.nCol);
        for (const auto& entry : moved)
            SetCell(newRowOf[entry.first.nRow - firstRow], entry.first.nCol, entry.second);
    }
    RecalculateFormulas();
    NotifyContentChanged();
    Invalidate();
    SetLastError(0);
    return TRUE;
}

namespace {
std::wstring Lowered(const std::wstring& s)
{
    std::wstring out = s;
    for (wchar_t& ch : out)
        ch = (wchar_t)towlower(ch);
    return out;
}
}

// Hide the rows whose cell in the filter column does not match. It used to
// select a rectangle from the first match to the last, non-matching rows
// included, and hide nothing (H16). Filtering is a view: it is not undoable
// and not saved, and FILTER_NONE shows every row again.
BOOL CGrid32Mgr::OnFilterCells(WPARAM wParam, const GCFILTERSTRUCT& filterStruct)
{
    if (filterStruct.m_cbSize != sizeof(GCFILTERSTRUCT) || wParam != 0)
    {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
        return FALSE;
    }
    const DWORD type = filterStruct.m_dwFilterType & 0xFF;
    const bool matchCase = (filterStruct.m_dwFilterType & FILTER_MATCHCASE) != 0;

    if (type == FILTER_NONE)
    {
        if (!m_hiddenRows.empty())
        {
            m_hiddenRows.clear();
            SetScrollRanges();
            Invalidate();
        }
        SetLastError(0);
        return TRUE;
    }
    if (type != FILTER_CONTAINS && type != FILTER_EQUALS &&
        type != FILTER_STARTS_WITH && type != FILTER_ENDS_WITH)
    {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
        return FALSE;
    }
    const GRIDPOINT header = filterStruct.m_filterColumn;
    if (header.nRow >= gcs.nHeight || header.nCol >= gcs.nWidth)
    {
        SetLastError(GRID_ERROR_OUT_OF_RANGE);
        return FALSE;
    }
    if (m_bEditing)
        CommitEdit();

    GRIDSELECTION used{};
    if (!OnGetUsedRange(&used))
    {
        SetLastError(0);
        return TRUE;    // nothing to filter
    }

    const std::wstring filterText = filterStruct.m_pwszFilterText ? filterStruct.m_pwszFilterText : L"";
    const std::wstring want = matchCase ? filterText : Lowered(filterText);
    for (UINT r = header.nRow + 1; r <= used.end.nRow && r < gcs.nHeight; ++r)
    {
        PGRIDCELL cell = GetCell(r, header.nCol);
        const std::wstring text = cell ? (matchCase ? cell->m_wsText : Lowered(cell->m_wsText)) : std::wstring();
        bool match = false;
        switch (type)
        {
        case FILTER_CONTAINS:    match = text.find(want) != std::wstring::npos; break;
        case FILTER_EQUALS:      match = text == want; break;
        case FILTER_STARTS_WITH: match = text.compare(0, want.size(), want) == 0; break;
        case FILTER_ENDS_WITH:
            match = text.size() >= want.size() &&
                text.compare(text.size() - want.size(), want.size(), want) == 0;
            break;
        }
        if (!match)
            m_hiddenRows.insert(r);
    }

    // The active cell can't stay on a row that just disappeared.
    if (IsRowHidden(m_currentCell.nRow))
        MoveCurrentCell(header.nRow, m_currentCell.nCol, false);
    ClampTopLeft();
    if (IsRowHidden(m_visibleTopLeft.nRow))
        m_visibleTopLeft.nRow = StepVisibleRow(m_visibleTopLeft.nRow, -1);
    SetScrollRanges();
    Invalidate();
    SetLastError(0);
    return TRUE;
}

// ---- find and replace --------------------------------------------------------

namespace {
size_t ReplaceAllIn(std::wstring& s, const std::wstring& find, const std::wstring& repl, bool matchCase)
{
    if (find.empty())
        return 0;
    const std::wstring hay = matchCase ? s : Lowered(s);
    const std::wstring needle = matchCase ? find : Lowered(find);
    std::wstring out;
    size_t count = 0, pos = 0, hit;
    while ((hit = hay.find(needle, pos)) != std::wstring::npos)
    {
        out.append(s, pos, hit - pos);
        out += repl;
        pos = hit + find.size();
        ++count;
    }
    if (count)
    {
        out.append(s, pos, std::wstring::npos);
        s.swap(out);
    }
    return count;
}
}

// Walks only the cells that exist, in reading order from m_startCell
// (inclusive), skipping rows a filter hid. It used to visit every position of
// the dense grid (P2). A cell matches on its displayed text or, for a
// formula, its source (M11).
bool CGrid32Mgr::FindNextMatch(const GCFINDSTRUCT& findStruct, GRIDPOINT& found)
{
    if (!findStruct.m_pwszFindText || !*findStruct.m_pwszFindText)
        return false;
    const bool matchCase = findStruct.m_bMatchCase != FALSE;
    const std::wstring findText = findStruct.m_pwszFindText;
    const std::wstring target = matchCase ? findText : Lowered(findText);
    auto matches = [&](const GRIDCELL& cell) {
        auto has = [&](const std::wstring& s) {
            return (matchCase ? s : Lowered(s)).find(target) != std::wstring::npos;
        };
        return has(cell.m_wsText) || (cell.m_bFormula && has(L"=" + cell.m_wsFormula));
    };
    const auto start = std::make_pair(findStruct.m_startCell.nRow, findStruct.m_startCell.nCol);

    if (findStruct.m_bSearchForward)
    {
        for (auto it = mapCells.lower_bound(start); it != mapCells.end(); ++it)
        {
            if (it->first.second >= gcs.nWidth || IsRowHidden(it->first.first))
                continue;
            if (matches(*it->second))
            {
                found = { it->first.first, it->first.second };
                return true;
            }
        }
        return false;
    }
    for (auto it = std::make_reverse_iterator(mapCells.upper_bound(start)); it != mapCells.rend(); ++it)
    {
        if (IsRowHidden(it->first.first))
            continue;
        if (matches(*it->second))
        {
            found = { it->first.first, it->first.second };
            return true;
        }
    }
    return false;
}

bool CGrid32Mgr::OnFindText(const GCFINDSTRUCT& findStruct)
{
    if (findStruct.m_cbSize != sizeof(GCFINDSTRUCT))
        return false;
    GRIDPOINT found{};
    if (!FindNextMatch(findStruct, found))
    {
        SetLastError(GRID_ERROR_NOT_EXIST);
        return false;
    }
    MoveCurrentCell(found.nRow, found.nCol, false);
    SetLastError(0);
    return true;
}

// Replace every occurrence in the first cell at or after m_startCell whose
// content (a formula's source, not its result) contains the text, and make
// that cell current. Replacing inside the displayed text used to overwrite a
// formula with its own result (M11), only the first occurrence was replaced
// (S6), and the cell was created and recorded twice for undo (M6).
bool CGrid32Mgr::OnReplaceText(const GCREPLACESTRUCT& replaceStruct)
{
    if (replaceStruct.m_cbSize != sizeof(GCREPLACESTRUCT) || !replaceStruct.m_pwszFindText || !*replaceStruct.m_pwszFindText)
        return false;
    const bool matchCase = replaceStruct.m_bMatchCase != FALSE;
    const std::wstring findText = replaceStruct.m_pwszFindText;
    const std::wstring replaceText = replaceStruct.m_pwszReplaceText ? replaceStruct.m_pwszReplaceText : L"";
    const auto start = std::make_pair(replaceStruct.m_startCell.nRow, replaceStruct.m_startCell.nCol);

    auto visit = [&](auto it) -> bool {
        if (IsRowHidden(it->first.first))
            return false;
        std::wstring source = EditTextOf(*it->second);
        if (!ReplaceAllIn(source, findText, replaceText, matchCase))
            return false;
        const GRIDPOINT pt{ it->first.first, it->first.second };
        SetCellText(pt.nRow, pt.nCol, source.c_str());
        MoveCurrentCell(pt.nRow, pt.nCol, false);
        return true;
    };

    if (replaceStruct.m_bSearchForward)
    {
        for (auto it = mapCells.lower_bound(start); it != mapCells.end(); ++it)
            if (visit(it))
            {
                SetLastError(0);
                return true;
            }
    }
    else
    {
        for (auto it = std::make_reverse_iterator(mapCells.upper_bound(start)); it != mapCells.rend(); ++it)
            if (visit(std::prev(it.base())))
            {
                SetLastError(0);
                return true;
            }
    }
    SetLastError(GRID_ERROR_NOT_EXIST);
    return false;
}

// Replace All, inside the grid: every occurrence in every visible cell, one
// undo step, one recalculation. The hosts used to loop on GM_REPLACETEXT,
// which re-matched the cell it had just changed whenever the replacement
// contained the search text and never finished (H5).
UINT CGrid32Mgr::OnReplaceAll(const GCREPLACESTRUCT& replaceStruct)
{
    if (replaceStruct.m_cbSize != sizeof(GCREPLACESTRUCT) || !replaceStruct.m_pwszFindText || !*replaceStruct.m_pwszFindText)
    {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
        return 0;
    }
    if (m_bEditing)
        CommitEdit();
    const bool matchCase = replaceStruct.m_bMatchCase != FALSE;
    const std::wstring findText = replaceStruct.m_pwszFindText;
    const std::wstring replaceText = replaceStruct.m_pwszReplaceText ? replaceStruct.m_pwszReplaceText : L"";

    std::vector<std::pair<GRIDPOINT, std::wstring>> edits;
    UINT count = 0;
    for (const auto& entry : mapCells)
    {
        if (IsRowHidden(entry.first.first))
            continue;
        std::wstring source = EditTextOf(*entry.second);
        const size_t n = ReplaceAllIn(source, findText, replaceText, matchCase);
        if (n)
        {
            count += (UINT)n;
            edits.push_back({ { entry.first.first, entry.first.second }, std::move(source) });
        }
    }
    if (edits.empty())
    {
        SetLastError(GRID_ERROR_NOT_EXIST);
        return 0;
    }
    {
        BatchUndo batch(*this);
        BulkEditScope bulkEdit(m_bDeferRecalc, m_bSuppressChangeNotify);
        for (const auto& edit : edits)
            SetCellText(edit.first.nRow, edit.first.nCol, edit.second.c_str());
    }
    RecalculateFormulas();
    NotifyContentChanged();
    Invalidate();
    SetLastError(0);
    return count;
}



// ===========================================================================
// GM_STREAMOUT / GM_STREAMIN
//
// Everything from here to the end of ClearAll serves those messages and
// nothing else. What each SF_* format carries -- and why only SF_SSF keeps a
// formula's source -- is documented next to the constants in grid32.h.
//
// The headers are included here rather than at the top of the file so that
// the stream code stays one self-contained block.
// ===========================================================================
#include <charconv>
#include <climits>
#include <locale>
#include <memory>

namespace {

// ---- numbers ---------------------------------------------------------------
// std::to_chars and std::from_chars ignore the locale by definition, and
// to_chars with no precision writes the shortest text that parses back to the
// identical value. std::to_wstring, which SSF used before, writes six fixed
// decimals with the C locale's decimal separator (AUDIT_2026-09.md M14).
template <class T>
std::wstring FormatExact(T value)
{
    char buf[64];
    const std::to_chars_result res = std::to_chars(buf, buf + sizeof(buf), value);
    if (res.ec != std::errc())
        return L"0";
    return std::wstring(buf, res.ptr);
}

// Reads what FormatExact writes, and what older SSF files hold: to_wstring's
// "1.500000", or "1,500000" when the writer ran under a comma-decimal locale.
// Like wcstod, it takes the longest valid prefix.
template <class T>
bool ParseExact(std::wstring_view text, T& out)
{
    const bool hasDot = text.find(L'.') != std::wstring_view::npos;
    std::string narrow;
    narrow.reserve(text.size());
    for (wchar_t ch : text)
    {
        if (narrow.empty() && (ch == L' ' || ch == L'\t'))
            continue;
        if (ch < 0x20 || ch > 0x7E)
            break;
        narrow.push_back(ch == L',' && !hasDot ? '.' : static_cast<char>(ch));
    }
    const char* first = narrow.data();
    const char* last = first + narrow.size();
    if (first != last && *first == '+')     // from_chars takes no '+'
        ++first;
    T value{};
    const std::from_chars_result res = std::from_chars(first, last, value);
    if (res.ec != std::errc() || res.ptr == first)
        return false;
    out = value;
    return true;
}

// A non-negative decimal integer that spans the whole of `text`.
bool ParseIndex(std::wstring_view text, unsigned long long& out)
{
    if (text.empty() || text.size() > 19)
        return false;
    unsigned long long value = 0;
    for (wchar_t ch : text)
    {
        if (ch < L'0' || ch > L'9')
            return false;
        value = value * 10 + static_cast<unsigned long long>(ch - L'0');
    }
    out = value;
    return true;
}

// ---- cells -----------------------------------------------------------------
// A populated cell, in the row-major order std::map keeps mapCells in. Writers
// walk a vector of these -- the cells that exist, in a deterministic order --
// instead of the grid's whole nHeight x nWidth area (AUDIT_2026-09.md P1).
struct StreamCell
{
    UINT row, col;
    const GRIDCELL* cell;
};

// A cell parsed from an interchange format, waiting to be applied to the grid.
// A `literal` cell keeps its text as text (a String in XML or ODF); any other
// goes through the keystroke typing path, so "42" becomes a number and "=A1"
// a formula. `format` is an FMT_* code to apply afterwards.
struct StagedCell
{
    UINT row, col;
    std::wstring text;
    bool literal;
    UINT format;
};

UINT LastColumn(const std::vector<StreamCell>& cells)
{
    UINT last = 0;
    for (const StreamCell& sc : cells)
        if (sc.col > last)
            last = sc.col;
    return last;
}

// What an interchange format exports for a cell: the value it displays, as the
// type it holds. A formula contributes its result, never its source.
enum class ExportKind { String, Number, Boolean, Date };

ExportKind ClassifyForExport(const GRIDCELL& cell, double& number)
{
    number = cell.m_dValue;
    switch (cell.m_eType)
    {
    case CT_Number:
        return std::isfinite(number) ? ExportKind::Number : ExportKind::String;
    case CT_Boolean:
        return ExportKind::Boolean;
    case CT_Date:
        return ExportKind::Date;
    case CT_Formula:
    {
        // A formula's result is a number exactly when its display text is one.
        double v = 0.0;
        if (Grid32Detail::TryParseNumber(cell.m_wsText, v) && std::isfinite(v))
        {
            number = v;
            return ExportKind::Number;
        }
        return ExportKind::String;
    }
    default:
        return ExportKind::String;
    }
}

// ISO 8601 date (and time, when the serial has a fraction or `alwaysTime` is
// set) for an Excel-style serial. False for serials outside 1900..9999, which
// neither Excel nor LibreOffice reads back as a date.
bool SerialToIso(double serial, bool alwaysTime, std::wstring& out)
{
    if (!std::isfinite(serial) || serial < 2.0 || serial >= 2958466.0)
        return false;
    const double whole = std::floor(serial);
    int y = 0, m = 0, d = 0;
    Grid32Detail::SerialToYMD(whole, y, m, d);
    long secs = std::lround((serial - whole) * 86400.0);
    if (secs > 86399)
        secs = 86399;
    wchar_t buf[40];
    if (secs == 0 && !alwaysTime)
        swprintf_s(buf, 40, L"%04d-%02d-%02d", y, m, d);
    else
        swprintf_s(buf, 40, L"%04d-%02d-%02dT%02ld:%02ld:%02ld", y, m, d,
                   secs / 3600, (secs / 60) % 60, secs % 60);
    out = buf;
    return true;
}

// ---- XML text --------------------------------------------------------------
// Escapes for element content and attribute values alike. CR, LF and tab are
// written as character references: a parser normalises a literal CR to LF and
// a literal tab or newline in an attribute to a space. The other C0 controls
// are not allowed in XML 1.0 even as references, so they are dropped -- a
// document that no parser accepts is worse than a missing control character.
void AppendXmlChar(std::wstring& out, wchar_t ch)
{
    switch (ch)
    {
    case L'&':  out += L"&amp;";  break;
    case L'<':  out += L"&lt;";   break;
    case L'>':  out += L"&gt;";   break;
    case L'"':  out += L"&quot;"; break;
    case L'\'': out += L"&apos;"; break;
    case L'\r': out += L"&#13;";  break;
    case L'\n': out += L"&#10;";  break;
    case L'\t': out += L"&#9;";   break;
    default:
        if (ch >= 0x20 && ch != 0xFFFE && ch != 0xFFFF)
            out.push_back(ch);
        break;
    }
}

std::wstring EscapeXML(const std::wstring& text)
{
    std::wstring out;
    out.reserve(text.size());
    for (wchar_t ch : text)
        AppendXmlChar(out, ch);
    return out;
}

void AppendCodePoint(std::wstring& out, unsigned long cp)
{
    if constexpr (sizeof(wchar_t) == 2)
    {
        if (cp > 0xFFFF)
        {
            cp -= 0x10000;
            out.push_back(static_cast<wchar_t>(0xD800 + (cp >> 10)));
            out.push_back(static_cast<wchar_t>(0xDC00 + (cp & 0x3FF)));
            return;
        }
    }
    out.push_back(static_cast<wchar_t>(cp));
}

// Decodes the five predefined entities and numeric character references.
// Anything unrecognised is kept as written.
std::wstring UnescapeXML(std::wstring_view text)
{
    std::wstring out;
    out.reserve(text.size());
    for (size_t i = 0; i < text.size(); ++i)
    {
        if (text[i] != L'&')
        {
            out.push_back(text[i]);
            continue;
        }
        const size_t semi = text.find(L';', i);
        if (semi == std::wstring_view::npos || semi - i > 12)
        {
            out.push_back(L'&');
            continue;
        }
        const std::wstring_view ent = text.substr(i + 1, semi - i - 1);
        if (ent == L"amp")       out.push_back(L'&');
        else if (ent == L"lt")   out.push_back(L'<');
        else if (ent == L"gt")   out.push_back(L'>');
        else if (ent == L"quot") out.push_back(L'"');
        else if (ent == L"apos") out.push_back(L'\'');
        else if (ent.size() > 1 && ent[0] == L'#')
        {
            const bool hex = ent[1] == L'x' || ent[1] == L'X';
            size_t k = hex ? 2 : 1;
            bool ok = k < ent.size();
            unsigned long cp = 0;
            for (; ok && k < ent.size(); ++k)
            {
                const wchar_t ch = ent[k];
                unsigned long digit = 0;
                if (ch >= L'0' && ch <= L'9')                digit = static_cast<unsigned long>(ch - L'0');
                else if (hex && ch >= L'a' && ch <= L'f')    digit = static_cast<unsigned long>(ch - L'a' + 10);
                else if (hex && ch >= L'A' && ch <= L'F')    digit = static_cast<unsigned long>(ch - L'A' + 10);
                else                                         ok = false;
                cp = cp * (hex ? 16 : 10) + digit;
                if (cp > 0x10FFFF)
                    ok = false;
            }
            if (!ok)
            {
                out.push_back(L'&');
                continue;
            }
            AppendCodePoint(out, cp);
        }
        else
        {
            out.push_back(L'&');
            continue;
        }
        i = semi;
    }
    return out;
}

// ---- XML scanning ----------------------------------------------------------
// The readers pull cells out of SpreadsheetML and ODF without a full XML
// parser: they find elements by name, read attributes and decode text. That
// covers what OnStreamOut, Excel and LibreOffice write.
constexpr size_t kNpos = std::wstring_view::npos;

bool IsXmlSpace(wchar_t ch)
{
    return ch == L' ' || ch == L'\t' || ch == L'\r' || ch == L'\n';
}

// True if a start tag (or empty-element tag) <name ...> begins at `pos`. The
// name must match whole: "Row" does not match <RowBreaks>.
bool IsElementAt(std::wstring_view xml, size_t pos, std::wstring_view name)
{
    if (pos >= xml.size() || xml[pos] != L'<')
        return false;
    const size_t end = pos + 1 + name.size();
    if (end >= xml.size() || xml.compare(pos + 1, name.size(), name) != 0)
        return false;
    const wchar_t next = xml[end];
    return IsXmlSpace(next) || next == L'>' || next == L'/';
}

// The next <name ...> at or after `from` and before `limit`, or kNpos.
size_t FindElement(std::wstring_view xml, std::wstring_view name, size_t from, size_t limit)
{
    while (from < limit)
    {
        const size_t p = xml.find(L'<', from);
        if (p == kNpos || p >= limit)
            return kNpos;
        if (IsElementAt(xml, p, name))
            return p;
        from = p + 1;
    }
    return kNpos;
}

// The '>' that closes the tag starting at `tagStart`, stepping over quoted
// attribute values; kNpos if unterminated. `selfClosing` reports "/>".
size_t FindTagEnd(std::wstring_view xml, size_t tagStart, bool& selfClosing)
{
    wchar_t quote = 0;
    for (size_t i = tagStart + 1; i < xml.size(); ++i)
    {
        const wchar_t ch = xml[i];
        if (quote)
        {
            if (ch == quote)
                quote = 0;
        }
        else if (ch == L'"' || ch == L'\'')
            quote = ch;
        else if (ch == L'>')
        {
            selfClosing = xml[i - 1] == L'/';
            return i;
        }
    }
    return kNpos;
}

// The decoded value of attribute `name` in the tag xml[tagStart, tagEnd).
bool GetAttr(std::wstring_view xml, size_t tagStart, size_t tagEnd,
             std::wstring_view name, std::wstring& value)
{
    const std::wstring_view tag = xml.substr(tagStart, tagEnd - tagStart);
    size_t i = 1;
    while (i < tag.size() && !IsXmlSpace(tag[i]) && tag[i] != L'/')
        ++i;                                        // the element name
    while (i < tag.size())
    {
        while (i < tag.size() && (IsXmlSpace(tag[i]) || tag[i] == L'/'))
            ++i;
        const size_t nameStart = i;
        while (i < tag.size() && tag[i] != L'=' && !IsXmlSpace(tag[i]) && tag[i] != L'/')
            ++i;
        const std::wstring_view attrName = tag.substr(nameStart, i - nameStart);
        while (i < tag.size() && IsXmlSpace(tag[i]))
            ++i;
        if (i >= tag.size() || tag[i] != L'=')
        {
            if (attrName.empty())
                break;
            continue;                               // an attribute with no value
        }
        ++i;
        while (i < tag.size() && IsXmlSpace(tag[i]))
            ++i;
        if (i >= tag.size() || (tag[i] != L'"' && tag[i] != L'\''))
            return false;
        const wchar_t quote = tag[i++];
        const size_t valueEnd = tag.find(quote, i);
        if (valueEnd == kNpos)
            return false;
        if (attrName == name)
        {
            value = UnescapeXML(tag.substr(i, valueEnd - i));
            return true;
        }
        i = valueEnd + 1;
    }
    return false;
}

// A repeat count attribute (ODF number-*-repeated); 1 when absent or invalid.
unsigned long long RepeatCount(std::wstring_view xml, size_t tagStart, size_t tagEnd,
                               std::wstring_view name)
{
    std::wstring attr;
    unsigned long long n = 0;
    if (GetAttr(xml, tagStart, tagEnd, name, attr) && ParseIndex(attr, n) && n > 0)
        return n;
    return 1;
}

// The character data of xml[begin, end) with any markup dropped.
std::wstring XmlTextContent(std::wstring_view xml, size_t begin, size_t end)
{
    std::wstring out;
    size_t i = begin;
    while (i < end)
    {
        if (xml[i] == L'<')
        {
            bool selfClosing = false;
            const size_t tagEnd = FindTagEnd(xml, i, selfClosing);
            if (tagEnd == kNpos || tagEnd >= end)
                break;
            i = tagEnd + 1;
            continue;
        }
        size_t next = xml.find(L'<', i);
        if (next == kNpos || next > end)
            next = end;
        out += UnescapeXML(xml.substr(i, next - i));
        i = next;
    }
    return out;
}

// ---- CSV / TSV / plain text --------------------------------------------------
void AppendDelimitedField(std::wstring& out, const std::wstring& text,
                          wchar_t delim, bool quoting)
{
    if (!quoting)
    {
        // Plain text has no escape mechanism. A tab or line break inside a
        // cell would change the table's shape, so each becomes one space.
        for (size_t i = 0; i < text.size(); ++i)
        {
            const wchar_t ch = text[i];
            if (ch == L'\r' && i + 1 < text.size() && text[i + 1] == L'\n')
                continue;                           // CRLF -> one space
            out.push_back(ch == L'\t' || ch == L'\r' || ch == L'\n' ? L' ' : ch);
        }
        return;
    }
    // RFC 4180 section 2: a field containing the delimiter, a quote, CR or LF
    // is enclosed in quotes, and each quote inside it is doubled.
    const wchar_t specials[] = { delim, L'"', L'\r', L'\n', L'\0' };
    if (text.find_first_of(specials) == std::wstring::npos)
    {
        out += text;
        return;
    }
    out.push_back(L'"');
    for (wchar_t ch : text)
    {
        if (ch == L'"')
            out.push_back(L'"');
        out.push_back(ch);
    }
    out.push_back(L'"');
}

// Rows 0 to the last populated row, every one as wide as the widest, so each
// record has the same number of fields; CRLF after every record.
std::wstring WriteDelimited(const std::vector<StreamCell>& cells, wchar_t delim, bool quoting)
{
    std::wstring out;
    if (cells.empty())
        return out;
    const UINT lastCol = LastColumn(cells);
    const UINT lastRow = cells.back().row;
    size_t i = 0;
    for (UINT r = 0; r <= lastRow; ++r)
    {
        for (UINT c = 0; c <= lastCol; ++c)
        {
            if (c > 0)
                out.push_back(delim);
            if (i < cells.size() && cells[i].row == r && cells[i].col == c)
                AppendDelimitedField(out, cells[i++].cell->m_wsText, delim, quoting);
        }
        out += L"\r\n";
    }
    return out;
}

void ParseDelimited(std::wstring_view input, wchar_t delim, bool quoting,
                    size_t maxRows, size_t maxCols, std::vector<StagedCell>& out)
{
    size_t row = 0, col = 0;
    std::wstring field;
    bool inQuotes = false;
    auto endField = [&]()
    {
        // An empty field is an empty cell, and the grid has been cleared, so
        // there is nothing to write. Fields past the grid's width are dropped.
        if (!field.empty() && col < maxCols)
            out.push_back({ static_cast<UINT>(row), static_cast<UINT>(col), field, false, FMT_GENERAL });
        field.clear();
        ++col;
    };
    for (size_t i = 0; i < input.size() && row < maxRows; ++i)
    {
        const wchar_t ch = input[i];
        if (inQuotes)
        {
            if (ch != L'"')
                field.push_back(ch);
            else if (i + 1 < input.size() && input[i + 1] == L'"')
            {
                field.push_back(L'"');              // RFC 4180: "" inside quotes is one "
                ++i;
            }
            else
                inQuotes = false;
            continue;
        }
        if (quoting && ch == L'"' && field.empty())
            inQuotes = true;                        // a field that starts with a quote is quoted
        else if (ch == delim)
            endField();
        else if (ch == L'\n')
        {
            endField();
            col = 0;
            ++row;
        }
        else if (ch != L'\r')                       // the CR of a CRLF line ending
            field.push_back(ch);
    }
    if (row < maxRows)
        endField();                                 // a last record with no line break
}

// ---- RTF -------------------------------------------------------------------
void AppendRtfUnit(std::wstring& out, unsigned long unit)
{
    // \uN takes a signed 16-bit N, one per UTF-16 unit; the '?' after it is
    // what a reader without Unicode support shows instead (\uc1).
    const long n = unit > 0x7FFF ? static_cast<long>(unit) - 0x10000 : static_cast<long>(unit);
    out += L"\\u";
    out += std::to_wstring(n);
    out.push_back(L'?');
}

void AppendRtfText(std::wstring& out, const std::wstring& text)
{
    for (size_t i = 0; i < text.size(); ++i)
    {
        const wchar_t ch = text[i];
        if (ch == L'\\' || ch == L'{' || ch == L'}')
        {
            out.push_back(L'\\');
            out.push_back(ch);
        }
        else if (ch == L'\t')
            out += L"\\tab ";
        else if (ch == L'\r' || ch == L'\n')
        {
            if (ch == L'\r' && i + 1 < text.size() && text[i + 1] == L'\n')
                ++i;
            out += L"\\line ";
        }
        else if (ch >= 0x20 && ch < 0x7F)
            out.push_back(ch);
        else if (ch >= 0x80)
        {
            unsigned long cp = static_cast<unsigned long>(ch);
            if (cp > 0xFFFF)                        // only where wchar_t is UTF-32
            {
                cp -= 0x10000;
                AppendRtfUnit(out, 0xD800 + (cp >> 10));
                AppendRtfUnit(out, 0xDC00 + (cp & 0x3FF));
            }
            else
                AppendRtfUnit(out, cp);
        }
        // Other control characters mean nothing inside a cell: dropped.
    }
}

// One RTF table row per grid row, 0 to the last populated row; every row
// restates its definition, as RTF requires. Column edges come from the grid's
// column widths at 15 twips per pixel (96 DPI).
std::wstring WriteRtf(const std::vector<StreamCell>& cells, const std::vector<size_t>& colWidthsPx)
{
    std::wstring out = L"{\\rtf1\\ansi\\ansicpg1252\\deff0\\uc1"
                       L"{\\fonttbl{\\f0\\fswiss\\fcharset0 Arial;}}\\f0\\fs24\r\n";
    if (!cells.empty())
    {
        const UINT lastCol = LastColumn(cells);
        const UINT lastRow = cells.back().row;
        std::wstring rowDef = L"\\trowd\\trgaph108";
        unsigned long long edge = 0;
        for (UINT c = 0; c <= lastCol; ++c)
        {
            const size_t px = c < colWidthsPx.size() && colWidthsPx[c] > 0 ? colWidthsPx[c] : 64;
            edge += static_cast<unsigned long long>(px) * 15;
            rowDef += L"\\cellx" + std::to_wstring(edge);
        }
        rowDef += L"\r\n\\pard\\intbl ";
        size_t i = 0;
        for (UINT r = 0; r <= lastRow; ++r)
        {
            out += rowDef;
            for (UINT c = 0; c <= lastCol; ++c)
            {
                if (i < cells.size() && cells[i].row == r && cells[i].col == c)
                    AppendRtfText(out, cells[i++].cell->m_wsText);
                out += L"\\cell ";
            }
            out += L"\\row\r\n";
        }
    }
    out += L"\\pard\\par\r\n}\r\n";
    return out;
}

// ---- SpreadsheetML 2003 (SF_XML) ---------------------------------------------
// Excel number-format codes for FMT_NUMBER..FMT_THOUSANDS. They are written as
// styles "fmt1".."fmt8", which the reader maps straight back.
const wchar_t* XmlNumberFormat(UINT fmt)
{
    switch (fmt)
    {
    case FMT_NUMBER:    return L"0.00";
    case FMT_INTEGER:   return L"0";
    case FMT_CURRENCY:  return L"$#,##0.00";
    case FMT_PERCENT:   return L"0.00%";
    case FMT_DATE_ISO:  return L"yyyy-mm-dd";
    case FMT_DATE_US:   return L"m/d/yyyy";
    case FMT_TIME:      return L"hh:mm:ss";
    case FMT_THOUSANDS: return L"#,##0.00";
    default:            return nullptr;
    }
}

std::wstring WriteSpreadsheetML(const std::vector<StreamCell>& cells)
{
    // Rows first, so the <Styles> block can list just the styles in use.
    bool usedFormat[FMT_THOUSANDS + 1] = {};
    bool usedDate = false;
    std::wstring rows;
    size_t i = 0;
    UINT nextRow = 0;
    while (i < cells.size())
    {
        // Rows and cells are sparse: ss:Index (1-based) skips the gaps.
        const UINT r = cells[i].row;
        rows += L"   <Row";
        if (r != nextRow)
            rows += L" ss:Index=\"" + std::to_wstring(r + 1ULL) + L"\"";
        rows += L">";
        UINT nextCol = 0;
        for (; i < cells.size() && cells[i].row == r; ++i)
        {
            const GRIDCELL& cell = *cells[i].cell;
            double number = 0.0;
            ExportKind kind = ClassifyForExport(cell, number);
            std::wstring iso;
            if (kind == ExportKind::Date && !SerialToIso(number, true, iso))
                kind = ExportKind::String;

            rows += L"<Cell";
            if (cells[i].col != nextCol)
                rows += L" ss:Index=\"" + std::to_wstring(cells[i].col + 1ULL) + L"\"";
            if (XmlNumberFormat(cell.m_nFormat))
            {
                usedFormat[cell.m_nFormat] = true;
                rows += L" ss:StyleID=\"fmt" + std::to_wstring(cell.m_nFormat) + L"\"";
            }
            else if (kind == ExportKind::Date)
            {
                // A General-format date would otherwise show as its serial number.
                usedDate = true;
                rows += L" ss:StyleID=\"dt\"";
            }
            rows += L"><Data ss:Type=\"";
            switch (kind)
            {
            case ExportKind::Number:  rows += L"Number\">" + FormatExact(number); break;
            case ExportKind::Boolean: rows += number != 0.0 ? L"Boolean\">1" : L"Boolean\">0"; break;
            case ExportKind::Date:    rows += L"DateTime\">" + iso + L".000"; break;   // Excel's own form
            default:                  rows += L"String\">" + EscapeXML(cell.m_wsText); break;
            }
            rows += L"</Data></Cell>";
            nextCol = cells[i].col + 1;
        }
        rows += L"</Row>\r\n";
        nextRow = r + 1;
    }

    std::wstring out =
        L"<?xml version=\"1.0\"?>\r\n"
        L"<?mso-application progid=\"Excel.Sheet\"?>\r\n"
        L"<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\""
        L" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">\r\n";
    bool anyStyle = usedDate;
    for (bool used : usedFormat)
        anyStyle = anyStyle || used;
    if (anyStyle)
    {
        out += L" <Styles>\r\n";
        if (usedDate)
            out += L"  <Style ss:ID=\"dt\"><NumberFormat ss:Format=\"Short Date\"/></Style>\r\n";
        for (UINT f = 0; f <= FMT_THOUSANDS; ++f)
            if (usedFormat[f])
                out += L"  <Style ss:ID=\"fmt" + std::to_wstring(f) + L"\"><NumberFormat ss:Format=\"" +
                       EscapeXML(XmlNumberFormat(f)) + L"\"/></Style>\r\n";
        out += L" </Styles>\r\n";
    }
    out += L" <Worksheet ss:Name=\"Sheet1\">\r\n  <Table>\r\n";
    out += rows;
    out += L"  </Table>\r\n </Worksheet>\r\n</Workbook>\r\n";
    return out;
}

// The <Data> of one <Cell>, as typed text. False if the cell has none.
bool XmlCellData(std::wstring_view xml, size_t begin, size_t end,
                 std::wstring& text, bool& literal)
{
    size_t data = FindElement(xml, L"Data", begin, end);
    std::wstring_view endTag = L"</Data>";
    const size_t prefixed = FindElement(xml, L"ss:Data", begin, end);
    if (prefixed < data)
    {
        data = prefixed;
        endTag = L"</ss:Data>";
    }
    if (data == kNpos)
        return false;
    bool selfClosing = false;
    const size_t openEnd = FindTagEnd(xml, data, selfClosing);
    if (openEnd == kNpos || openEnd >= end)
        return false;
    std::wstring type;
    GetAttr(xml, data, openEnd, L"ss:Type", type);
    std::wstring content;
    if (!selfClosing)
    {
        size_t close = xml.find(endTag, openEnd);
        if (close == kNpos || close > end)
            close = end;
        content = XmlTextContent(xml, openEnd + 1, close);  // drops rich-text markup
    }
    literal = false;
    text.clear();
    if (content.empty())
        return true;
    if (type == L"Boolean")
        text = content == L"0" ? L"FALSE" : L"TRUE";
    else if (type == L"DateTime")
        text = content.substr(0, 10);               // the date; the grid's dates have no time
    else if (type == L"Number")
        text = content;
    else
    {
        text = content;                             // String, Error: kept as written
        literal = true;
    }
    return true;
}

// Cells of the first worksheet of an Excel 2003 XML Spreadsheet: <Row> and
// <Cell> with their optional 1-based ss:Index, ss:MergeAcross spans, typed
// <Data>, and ss:StyleID "fmtN" mapped back to FMT_N. False if the input is
// not a SpreadsheetML workbook.
bool ParseSpreadsheetML(std::wstring_view xml, size_t maxRows, size_t maxCols,
                        std::vector<StagedCell>& out)
{
    if (FindElement(xml, L"Workbook", 0, xml.size()) == kNpos &&
        FindElement(xml, L"ss:Workbook", 0, xml.size()) == kNpos)
        return false;
    const size_t table = FindElement(xml, L"Table", 0, xml.size());
    if (table == kNpos)
        return true;                                // a workbook with an empty sheet
    size_t tableEnd = xml.find(L"</Table>", table);
    if (tableEnd == kNpos)
        tableEnd = xml.size();

    const std::wstring_view rowEndTag = L"</Row>";
    const std::wstring_view cellEndTag = L"</Cell>";
    std::wstring attr;
    unsigned long long index = 0;
    unsigned long long row = 0;
    size_t p = table + 1;
    while (row < maxRows)
    {
        const size_t rowStart = FindElement(xml, L"Row", p, tableEnd);
        if (rowStart == kNpos)
            break;
        bool rowSelfClosing = false;
        const size_t rowOpenEnd = FindTagEnd(xml, rowStart, rowSelfClosing);
        if (rowOpenEnd == kNpos)
            break;
        if (GetAttr(xml, rowStart, rowOpenEnd, L"ss:Index", attr) && ParseIndex(attr, index) && index > 0)
            row = index - 1;
        size_t rowEnd = rowSelfClosing ? rowOpenEnd : xml.find(rowEndTag, rowOpenEnd);
        if (rowEnd == kNpos || rowEnd > tableEnd)
            rowEnd = tableEnd;

        unsigned long long col = 0;
        size_t c = rowOpenEnd + 1;
        while (!rowSelfClosing && row < maxRows)
        {
            const size_t cellStart = FindElement(xml, L"Cell", c, rowEnd);
            if (cellStart == kNpos)
                break;
            bool cellSelfClosing = false;
            const size_t cellOpenEnd = FindTagEnd(xml, cellStart, cellSelfClosing);
            if (cellOpenEnd == kNpos || cellOpenEnd > rowEnd)
                break;
            if (GetAttr(xml, cellStart, cellOpenEnd, L"ss:Index", attr) && ParseIndex(attr, index) && index > 0)
                col = index - 1;
            unsigned long long span = 0;
            if (!GetAttr(xml, cellStart, cellOpenEnd, L"ss:MergeAcross", attr) || !ParseIndex(attr, span))
                span = 0;
            UINT format = FMT_GENERAL;
            unsigned long long f = 0;
            if (GetAttr(xml, cellStart, cellOpenEnd, L"ss:StyleID", attr) && attr.size() > 3 &&
                attr.compare(0, 3, L"fmt") == 0 && ParseIndex(std::wstring_view(attr).substr(3), f) &&
                XmlNumberFormat(static_cast<UINT>(f)))
                format = static_cast<UINT>(f);

            size_t next = cellOpenEnd + 1;
            if (!cellSelfClosing)
            {
                size_t close = xml.find(cellEndTag, cellOpenEnd);
                if (close == kNpos || close > rowEnd)
                    close = rowEnd;
                std::wstring text;
                bool literal = false;
                if (XmlCellData(xml, cellOpenEnd + 1, close, text, literal) && !text.empty() && col < maxCols)
                    out.push_back({ static_cast<UINT>(row), static_cast<UINT>(col), std::move(text), literal, format });
                next = close + cellEndTag.size();
            }
            col += 1 + span;
            c = next;
        }
        ++row;
        p = rowSelfClosing ? rowOpenEnd + 1 : rowEnd + rowEndTag.size();
    }
    return true;
}

// ---- OpenDocument flat spreadsheet (SF_ODF) ---------------------------------
// One line of cell text as a <text:p>. ODF collapses runs of white space
// inside a paragraph and drops it at either end, so a tab is <text:tab/> and
// every space that would be lost is a <text:s/>.
void AppendOdfParagraph(std::wstring& out, std::wstring_view line)
{
    out += L"<text:p>";
    size_t i = 0;
    while (i < line.size())
    {
        const wchar_t ch = line[i];
        if (ch == L' ')
        {
            size_t run = 0;
            while (i + run < line.size() && line[i + run] == L' ')
                ++run;
            // Inside the line a run keeps its first space as a real one.
            const size_t literal = (i == 0 || i + run == line.size()) ? 0 : 1;
            if (literal)
                out.push_back(L' ');
            const size_t encoded = run - literal;
            if (encoded == 1)
                out += L"<text:s/>";
            else if (encoded > 1)
                out += L"<text:s text:c=\"" + std::to_wstring(encoded) + L"\"/>";
            i += run;
        }
        else
        {
            if (ch == L'\t')
                out += L"<text:tab/>";
            else
                AppendXmlChar(out, ch);
            ++i;
        }
    }
    out += L"</text:p>";
}

void AppendOdfCell(std::wstring& out, const GRIDCELL& cell)
{
    double number = 0.0;
    ExportKind kind = ClassifyForExport(cell, number);
    std::wstring iso;
    if (kind == ExportKind::Date && !SerialToIso(number, false, iso))
        kind = ExportKind::String;
    out += L"<table:table-cell office:value-type=\"";
    switch (kind)
    {
    case ExportKind::Number:  out += L"float\" office:value=\"" + FormatExact(number); break;
    case ExportKind::Boolean: out += number != 0.0 ? L"boolean\" office:boolean-value=\"true"
                                                   : L"boolean\" office:boolean-value=\"false"; break;
    case ExportKind::Date:    out += L"date\" office:date-value=\"" + iso; break;
    default:                  out += L"string"; break;
    }
    out += L"\">";
    // One paragraph per line of the cell's display text.
    const std::wstring& text = cell.m_wsText;
    size_t start = 0;
    for (;;)
    {
        const size_t eol = text.find_first_of(L"\r\n", start);
        AppendOdfParagraph(out, std::wstring_view(text).substr(start, eol == std::wstring::npos ? std::wstring::npos : eol - start));
        if (eol == std::wstring::npos)
            break;
        start = eol + (text[eol] == L'\r' && eol + 1 < text.size() && text[eol + 1] == L'\n' ? 2 : 1);
    }
    out += L"</table:table-cell>";
}

std::wstring WriteFlatOdf(const std::vector<StreamCell>& cells)
{
    std::wstring out =
        L"<?xml version=\"1.0\"?>\r\n"
        L"<office:document"
        L" xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\""
        L" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\""
        L" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\""
        L" office:version=\"1.2\""
        L" office:mimetype=\"application/vnd.oasis.opendocument.spreadsheet\">\r\n"
        L" <office:body>\r\n  <office:spreadsheet>\r\n   <table:table table:name=\"Sheet1\">\r\n";
    // A table declares its columns before its rows, and holds at least one row.
    const UINT columns = cells.empty() ? 1 : LastColumn(cells) + 1;
    out += L"    <table:table-column table:number-columns-repeated=\"" + std::to_wstring(columns) + L"\"/>\r\n";
    if (cells.empty())
        out += L"    <table:table-row><table:table-cell/></table:table-row>\r\n";

    // Sparse: runs of empty rows and cells are written once, with a repeat count.
    size_t i = 0;
    UINT nextRow = 0;
    while (i < cells.size())
    {
        const UINT r = cells[i].row;
        if (r > nextRow)
        {
            out += L"    <table:table-row";
            if (r - nextRow > 1)
                out += L" table:number-rows-repeated=\"" + std::to_wstring(r - nextRow) + L"\"";
            out += L"><table:table-cell/></table:table-row>\r\n";
        }
        out += L"    <table:table-row>";
        UINT nextCol = 0;
        for (; i < cells.size() && cells[i].row == r; ++i)
        {
            const UINT c = cells[i].col;
            if (c > nextCol)
            {
                out += L"<table:table-cell";
                if (c - nextCol > 1)
                    out += L" table:number-columns-repeated=\"" + std::to_wstring(c - nextCol) + L"\"";
                out += L"/>";
            }
            AppendOdfCell(out, *cells[i].cell);
            nextCol = c + 1;
        }
        out += L"</table:table-row>\r\n";
        nextRow = r + 1;
    }
    out += L"   </table:table>\r\n  </office:spreadsheet>\r\n </office:body>\r\n</office:document>\r\n";
    return out;
}

// The text a cell shows: its <text:p> paragraphs joined by line breaks, with
// <text:s/>, <text:tab/> and <text:line-break/> expanded and any other markup
// (spans, links) dropped from around its text.
std::wstring OdfCellText(std::wstring_view xml, size_t begin, size_t end)
{
    std::wstring out;
    bool first = true;
    size_t p = begin;
    while ((p = FindElement(xml, L"text:p", p, end)) != kNpos)
    {
        bool selfClosing = false;
        const size_t openEnd = FindTagEnd(xml, p, selfClosing);
        if (openEnd == kNpos || openEnd >= end)
            break;
        if (!first)
            out.push_back(L'\n');
        first = false;
        if (selfClosing)
        {
            p = openEnd + 1;
            continue;
        }
        size_t paraEnd = xml.find(L"</text:p>", openEnd);
        if (paraEnd == kNpos || paraEnd > end)
            paraEnd = end;
        size_t i = openEnd + 1;
        while (i < paraEnd)
        {
            if (xml[i] != L'<')
            {
                size_t next = xml.find(L'<', i);
                if (next == kNpos || next > paraEnd)
                    next = paraEnd;
                out += UnescapeXML(xml.substr(i, next - i));
                i = next;
                continue;
            }
            bool tagSelfClosing = false;
            const size_t tagEnd = FindTagEnd(xml, i, tagSelfClosing);
            if (tagEnd == kNpos || tagEnd >= paraEnd)
                break;
            if (IsElementAt(xml, i, L"text:s"))
            {
                std::wstring count;
                unsigned long long n = 1;
                if (!GetAttr(xml, i, tagEnd, L"text:c", count) || !ParseIndex(count, n) || n == 0)
                    n = 1;
                out.append(static_cast<size_t>(n < 4096 ? n : 4096), L' ');
            }
            else if (IsElementAt(xml, i, L"text:tab"))
                out.push_back(L'\t');
            else if (IsElementAt(xml, i, L"text:line-break"))
                out.push_back(L'\n');
            i = tagEnd + 1;
        }
        p = paraEnd;
    }
    return out;
}

// Typed value of one ODF cell. Numbers, booleans and dates come from their
// office: attributes and are typed as if entered; a string is kept literally.
void OdfCellValue(std::wstring_view xml, size_t cellStart, size_t openEnd, size_t contentEnd,
                  std::wstring& text, bool& literal)
{
    std::wstring type, value;
    GetAttr(xml, cellStart, openEnd, L"office:value-type", type);
    literal = false;
    if ((type == L"float" || type == L"percentage" || type == L"currency") &&
        GetAttr(xml, cellStart, openEnd, L"office:value", value))
    {
        text = value;
        return;
    }
    if (type == L"boolean" && GetAttr(xml, cellStart, openEnd, L"office:boolean-value", value))
    {
        text = value == L"true" ? L"TRUE" : L"FALSE";
        return;
    }
    if (type == L"date" && GetAttr(xml, cellStart, openEnd, L"office:date-value", value))
    {
        text = value.substr(0, 10);
        return;
    }
    text = contentEnd > openEnd + 1 ? OdfCellText(xml, openEnd + 1, contentEnd) : std::wstring();
    literal = type == L"string";
}

// Cells of the first table of an ODF spreadsheet (a flat .fods, or the
// content.xml of a package). Honours table:number-rows-repeated and
// table:number-columns-repeated, which LibreOffice writes for every run of
// empty or identical cells. False if there is no table.
bool ParseOdf(std::wstring_view xml, size_t maxRows, size_t maxCols, std::vector<StagedCell>& out)
{
    const size_t table = FindElement(xml, L"table:table", 0, xml.size());
    if (table == kNpos)
        return false;
    size_t tableEnd = xml.find(L"</table:table>", table);
    if (tableEnd == kNpos)
        tableEnd = xml.size();

    struct RowCell { unsigned long long col, repeat; std::wstring text; bool literal; };
    std::vector<RowCell> rowCells;
    const std::wstring_view rowEndTag = L"</table:table-row>";
    unsigned long long row = 0;
    size_t p = table + 1;
    while (row < maxRows)
    {
        const size_t rowStart = FindElement(xml, L"table:table-row", p, tableEnd);
        if (rowStart == kNpos)
            break;
        bool rowSelfClosing = false;
        const size_t rowOpenEnd = FindTagEnd(xml, rowStart, rowSelfClosing);
        if (rowOpenEnd == kNpos)
            break;
        const unsigned long long rowsRepeated =
            RepeatCount(xml, rowStart, rowOpenEnd, L"table:number-rows-repeated");
        size_t rowEnd = rowSelfClosing ? rowOpenEnd : xml.find(rowEndTag, rowOpenEnd);
        if (rowEnd == kNpos || rowEnd > tableEnd)
            rowEnd = tableEnd;

        rowCells.clear();
        unsigned long long col = 0;
        size_t c = rowOpenEnd + 1;
        // Covered cells (under a merge) are rare: remember where the next one
        // is rather than rescanning the rest of the row for every cell.
        size_t covered = FindElement(xml, L"table:covered-table-cell", c, rowEnd);
        while (!rowSelfClosing && col < maxCols)
        {
            size_t cellStart = FindElement(xml, L"table:table-cell", c, rowEnd);
            std::wstring_view cellEndTag = L"</table:table-cell>";
            if (covered != kNpos && covered < c)
                covered = FindElement(xml, L"table:covered-table-cell", c, rowEnd);
            if (covered < cellStart)
            {
                cellStart = covered;
                cellEndTag = L"</table:covered-table-cell>";
            }
            if (cellStart == kNpos)
                break;
            bool cellSelfClosing = false;
            const size_t cellOpenEnd = FindTagEnd(xml, cellStart, cellSelfClosing);
            if (cellOpenEnd == kNpos || cellOpenEnd > rowEnd)
                break;
            const unsigned long long repeat =
                RepeatCount(xml, cellStart, cellOpenEnd, L"table:number-columns-repeated");
            size_t contentEnd = cellOpenEnd + 1;
            size_t next = cellOpenEnd + 1;
            if (!cellSelfClosing)
            {
                contentEnd = xml.find(cellEndTag, cellOpenEnd);
                if (contentEnd == kNpos || contentEnd > rowEnd)
                    contentEnd = rowEnd;
                next = contentEnd + cellEndTag.size();
            }
            std::wstring text;
            bool literal = false;
            OdfCellValue(xml, cellStart, cellOpenEnd, contentEnd, text, literal);
            if (!text.empty())
                rowCells.push_back({ col, repeat, std::move(text), literal });
            col += repeat;
            c = next;
        }

        if (rowCells.empty())
            row += rowsRepeated < maxRows - row ? rowsRepeated : maxRows - row;
        for (unsigned long long k = 0; !rowCells.empty() && k < rowsRepeated && row < maxRows; ++k, ++row)
            for (const RowCell& rc : rowCells)
                for (unsigned long long j = 0; j < rc.repeat && rc.col + j < maxCols; ++j)
                    out.push_back({ static_cast<UINT>(row), static_cast<UINT>(rc.col + j), rc.text, rc.literal, FMT_GENERAL });
        p = rowSelfClosing ? rowOpenEnd + 1 : rowEnd + rowEndTag.size();
    }
    return true;
}

// ---- SSF (SF_SSF) ----------------------------------------------------------
void AppendTLV(std::wostream& out, const wchar_t* tag, const std::wstring& value)
{
    out << L';' << tag << L':' << value.size() << L':' << value;
}

// The TAG:LEN:VALUE fields of one cell (or of the sheet default).
void AppendSsfTags(std::wostream& ss, const GRIDCELL* cell)
{
    if (cell->m_eType != CT_Text)
    {
        const wchar_t* typeStr = L"t";
        switch (cell->m_eType) {
            case CT_Number:  typeStr = L"n"; break;
            case CT_Date:    typeStr = L"d"; break;
            case CT_Boolean: typeStr = L"b"; break;
            case CT_Formula: typeStr = L"f"; break;
            case CT_Error:   typeStr = L"e"; break;
            default:         typeStr = L"t"; break;
        }
        AppendTLV(ss, L"Y", typeStr);
        // V carries the numeric form for any non-text type.
        AppendTLV(ss, L"V", FormatExact(cell->m_dValue));
    }
    // FM always: a sheet default with a number format would otherwise turn
    // every General cell read back into that format.
    AppendTLV(ss, L"FM", std::to_wstring(cell->m_nFormat));
    AppendTLV(ss, L"T", cell->m_wsText);
    if (cell->m_bFormula)
        AppendTLV(ss, L"F", cell->m_wsFormula);
    AppendTLV(ss, L"FF", cell->fontInfo.m_wsFontFace);
    AppendTLV(ss, L"FS", FormatExact(cell->fontInfo.m_fPointSize));
    AppendTLV(ss, L"FC", std::to_wstring(cell->fontInfo.m_clrTextColor));
    AppendTLV(ss, L"IT", std::to_wstring(cell->fontInfo.bItalic));
    AppendTLV(ss, L"UN", std::to_wstring(cell->fontInfo.bUnderline));
    AppendTLV(ss, L"ST", std::to_wstring(cell->fontInfo.bStrikeThrough));
    AppendTLV(ss, L"WT", std::to_wstring(cell->fontInfo.bWeight));
    AppendTLV(ss, L"BG", std::to_wstring(cell->clrBackground));
    AppendTLV(ss, L"BC", std::to_wstring(cell->m_clrBorderColor));
    AppendTLV(ss, L"BW", std::to_wstring(cell->m_nBorderWidth));
    AppendTLV(ss, L"PS", std::to_wstring(cell->penStyle));
    AppendTLV(ss, L"J", std::to_wstring(cell->justification));
    AppendTLV(ss, L"NM", cell->m_wsName);
    AppendTLV(ss, L"MR", std::to_wstring(cell->mergeRange.start.nRow) + L"," +
                          std::to_wstring(cell->mergeRange.start.nCol) + L"," +
                          std::to_wstring(cell->mergeRange.end.nRow) + L"," +
                          std::to_wstring(cell->mergeRange.end.nCol));
}

std::wstring WriteSsf(const std::map<std::pair<UINT, UINT>, PGRIDCELL>& cells,
                      size_t nHeight, size_t nWidth, const GRIDCELL* sheetDefault)
{
    // SSF v2 -- Spreadsheet File format, v2.
    //   Line 1:  "SSF2,1"               -- format marker + revision
    //   Line 2:  "WB,<sheetCount>"      -- workbook envelope
    //   Per sheet:
    //     "SH,<name>,<rows>,<cols>"     -- sheet header
    //     One record per cell, in row-major order:
    //       "<row>,<col>;TAG:LEN:VALUE[;TAG:LEN:VALUE]*"
    //     LEN counts the characters of VALUE, which may therefore contain
    //     anything, line breaks included.
    //     Tags:
    //       Y   cell type char (n/d/b/f/e -- t omitted)
    //       V   numeric value (for n/d/b/f/e types): shortest round-trip
    //           decimal, '.' separator, whatever the locale
    //       T   display text
    //       F   formula source (no leading '=')
    //       FS  font point size, written like V
    //       FM/FF/FC/IT/UN/ST/WT/BG/BC/BW/PS/J/NM/MR -- as in v1
    //     "DF;TAG:LEN:VALUE..." before the cells, when the sheet's empty
    //     cells are formatted (whole-sheet formatting): the template every
    //     cell of the sheet starts from.
    //     Every cell that exists is written, including one that holds only
    //     formatting: a highlighted empty cell, or the anchor of a merged
    //     block. Skipping those lost the formatting on save.
    std::wostringstream ss;
    // A host that installs a global C++ locale with digit grouping would
    // otherwise get row 1234 written as "1,234", corrupting the record.
    ss.imbue(std::locale::classic());
    ss << L"SSF2,1\n";
    ss << L"WB,1\n";
    ss << L"SH,Sheet1," << nHeight << L"," << nWidth << L"\n";
    if (sheetDefault)
    {
        ss << L"DF";
        AppendSsfTags(ss, sheetDefault);
        ss << L"\n";
    }
    for (const auto& entry : cells)
    {
        ss << entry.first.first << L"," << entry.first.second;
        AppendSsfTags(ss, entry.second);
        ss << L"\n";
    }
    return ss.str();
}

// One TAG:LEN:VALUE field of an SSF record, applied to `cell`.
void ApplySsfTag(GRIDCELL& cell, std::wstring_view tag, const std::wstring& value,
                 bool& hadExplicitType)
{
    if (tag == L"T")
        cell.m_wsText = value;
    else if (tag == L"Y" && !value.empty())
    {
        // v2: explicit cell type. Maps t/n/d/b/f/e to CellType enum.
        hadExplicitType = true;
        switch (value[0]) {
            case L'n': cell.m_eType = CT_Number;  break;
            case L'd': cell.m_eType = CT_Date;    break;
            case L'b': cell.m_eType = CT_Boolean; break;
            case L'f': cell.m_eType = CT_Formula; break;
            case L'e': cell.m_eType = CT_Error;   break;
            default:   cell.m_eType = CT_Text;    break;
        }
    }
    else if (tag == L"V")
        ParseExact(value, cell.m_dValue);
    else if (tag == L"F")
    {
        cell.m_bFormula = true;
        cell.m_wsFormula = value;
    }
    else if (tag == L"FF")
        cell.fontInfo.m_wsFontFace = value;
    else if (tag == L"FS")
        ParseExact(value, cell.fontInfo.m_fPointSize);
    else if (tag == L"FC")
        cell.fontInfo.m_clrTextColor = (COLORREF)std::stoul(value);
    else if (tag == L"IT")
        cell.fontInfo.bItalic = (BOOL)std::stoul(value);
    else if (tag == L"UN")
        cell.fontInfo.bUnderline = (BOOL)std::stoul(value);
    else if (tag == L"ST")
        cell.fontInfo.bStrikeThrough = (BOOL)std::stoul(value);
    else if (tag == L"WT")
        cell.fontInfo.bWeight = (UINT)std::stoul(value);
    else if (tag == L"BG")
        cell.clrBackground = (COLORREF)std::stoul(value);
    else if (tag == L"BC")
        cell.m_clrBorderColor = (COLORREF)std::stoul(value);
    else if (tag == L"BW")
        cell.m_nBorderWidth = (UINT)std::stoul(value);
    else if (tag == L"PS")
        cell.penStyle = std::stoi(value);
    else if (tag == L"J")
        cell.justification = (UINT)std::stoul(value);
    else if (tag == L"FM")
        cell.m_nFormat = (UINT)std::stoul(value);
    else if (tag == L"NM")
        cell.m_wsName = value;
    else if (tag == L"MR")
    {
        size_t p1 = 0, p2 = value.find(L',', p1);
        if (p2 != std::wstring::npos)
        {
            cell.mergeRange.start.nRow = (UINT)std::stoul(value.substr(p1, p2 - p1));
            p1 = p2 + 1; p2 = value.find(L',', p1);
            if (p2 != std::wstring::npos)
            {
                cell.mergeRange.start.nCol = (UINT)std::stoul(value.substr(p1, p2 - p1));
                p1 = p2 + 1; p2 = value.find(L',', p1);
                if (p2 != std::wstring::npos)
                {
                    cell.mergeRange.end.nRow = (UINT)std::stoul(value.substr(p1, p2 - p1));
                    cell.mergeRange.end.nCol = (UINT)std::stoul(value.substr(p2 + 1));
                }
            }
        }
    }
}

// Parses SSF v1 ("SSF,<rev>") and v2 ("SSF2,<rev>") into finished cells.
// Records are read by position, not split into lines first: every value is
// length-prefixed, so a cell whose text holds a line break reads back intact
// (splitting on lines used to cut such a record in two). Malformed numbers in
// a value throw std::invalid_argument, which the caller turns into an error.
DWORD ParseSsf(std::wstring_view in, const GRIDCELL& defaultCell, size_t maxRows, size_t maxCols,
               std::map<std::pair<UINT, UINT>, GRIDCELL>& out,
               GRIDCELL& sheetDefault, bool& hasSheetDefault)
{
    // Cells start from the file's DF record when it has one.
    GRIDCELL templateCell = defaultCell;
    hasSheetDefault = false;
    auto lineEnd = [&in](size_t from) {
        const size_t e = in.find(L'\n', from);
        return e == kNpos ? in.size() : e;
    };
    auto startsWith = [&in](size_t pos, std::wstring_view prefix) {
        return in.compare(pos, prefix.size(), prefix) == 0;
    };

    const bool isV2 = startsWith(0, L"SSF2");
    const bool isV1 = !isV2 && startsWith(0, L"SSF");
    if (!isV1 && !isV2)
        return GRID_ERROR_INVALID_PARAMETER;
    size_t pos = lineEnd(0) + 1;
    if (isV1 && pos < in.size())
        pos = lineEnd(pos) + 1;                     // v1 dimensions line, unused on load

    while (pos < in.size())
    {
        const size_t eol = lineEnd(pos);
        if (in[pos] == L'\n' || in[pos] == L'\r' ||
            (isV2 && (startsWith(pos, L"WB,") || startsWith(pos, L"SH,"))))
        {
            pos = eol + 1;                          // blank, or envelope / sheet header
            continue;
        }
        const size_t semi = in.find(L';', pos);
        if (semi == kNpos || semi > eol)
        {
            pos = eol + 1;                          // no record on this line
            continue;
        }
        const std::wstring_view coords = in.substr(pos, semi - pos);
        const bool isDefault = isV2 && coords == L"DF";
        const size_t comma = coords.find(L',');
        unsigned long long r = 0, c = 0;
        if (!isDefault && (comma == kNpos || !ParseIndex(coords.substr(0, comma), r) ||
            !ParseIndex(coords.substr(comma + 1), c)))
            return GRID_ERROR_INVALID_PARAMETER;

        GRIDCELL cell = templateCell;
        bool hadExplicitType = false;
        size_t p = semi + 1;
        while (p < in.size() && in[p] != L'\n' && in[p] != L'\r')
        {
            const size_t tagEnd = in.find(L':', p);
            const size_t lenEnd = tagEnd == kNpos ? kNpos : in.find(L':', tagEnd + 1);
            unsigned long long len = 0;
            if (lenEnd == kNpos || !ParseIndex(in.substr(tagEnd + 1, lenEnd - tagEnd - 1), len) ||
                len > in.size() - (lenEnd + 1))
                return GRID_ERROR_INVALID_PARAMETER;
            const std::wstring value(in.substr(lenEnd + 1, static_cast<size_t>(len)));
            ApplySsfTag(cell, in.substr(p, tagEnd - p), value, hadExplicitType);
            p = lenEnd + 1 + static_cast<size_t>(len);
            if (p < in.size() && in[p] == L';')
                ++p;
            else if (p < in.size() && in[p] != L'\n' && in[p] != L'\r')
                return GRID_ERROR_INVALID_PARAMETER;
        }
        pos = p;
        if (pos < in.size() && in[pos] == L'\r')
            ++pos;
        if (pos < in.size() && in[pos] == L'\n')
            ++pos;

        if (isDefault)
        {
            templateCell = cell;
            sheetDefault = cell;
            hasSheetDefault = true;
            continue;
        }
        if (r >= maxRows || c >= maxCols)
            continue;                               // outside this grid

        // v1 files don't carry Y/V tags -- infer typing from the loaded text
        // so opened-from-v1 sheets behave like freshly-typed ones (proper
        // sort, formula-result numeric form).
        if (!hadExplicitType)
        {
            if (cell.m_bFormula)
            {
                cell.m_eType = CT_Formula;
                double n = 0.0;
                if (Grid32Detail::TryParseNumber(cell.m_wsText, n)) cell.m_dValue = n;
            }
            else
            {
                double v = 0.0;
                cell.m_eType = Grid32Detail::InferType(cell.m_wsText, v);
                cell.m_dValue = v;
            }
        }
        // Regenerate display text from the format code so the file stays
        // canonical even if the format renderer changes between revisions.
        if (cell.m_nFormat != FMT_GENERAL && cell.m_eType != CT_Text && !cell.m_bFormula)
        {
            cell.m_wsText = Grid32Detail::FormatCellValue(
                cell.m_eType, cell.m_dValue, cell.m_wsText, cell.m_nFormat);
        }
        out[{ static_cast<UINT>(r), static_cast<UINT>(c) }] = std::move(cell);
    }
    return 0;
}

} // namespace

void CGrid32Mgr::OnStreamOut(LPGCSTREAM pStream)
{
    if (!pStream || pStream->m_cbSize != sizeof(GCSTREAM))
        return;

    // Every exit from here reports through m_dwError and the callback, and
    // none reports success without having written the whole output.
    auto finish = [pStream](DWORD error)
    {
        pStream->m_dwError = error;
        if (pStream->m_pfnCallback)
            pStream->m_pfnCallback(pStream);
    };

    pStream->m_cbBuffOut = 0;
    const size_t bufWchars = pStream->m_cbBuffSize / sizeof(wchar_t);
    if (!pStream->m_pwszBuff || bufWchars == 0)
    {
        finish(GRID_ERROR_INVALID_PARAMETER);
        return;
    }
    pStream->m_pwszBuff[0] = L'\0';

    DWORD fmt = pStream->m_dwFormat;
    if (fmt == 0) fmt = SF_CSV; // default

    // The interchange formats write what cells display; one that displays
    // nothing is an empty position. mapCells is already in row-major order.
    std::vector<StreamCell> cells;
    if (fmt != SF_SSF)
    {
        cells.reserve(mapCells.size());
        for (const auto& entry : mapCells)
            if (!entry.second->m_wsText.empty())
                cells.push_back({ entry.first.first, entry.first.second, entry.second });
    }

    std::wstring data;
    switch (fmt)
    {
    case SF_CSV:  data = WriteDelimited(cells, L',', true);   break;
    case SF_TSV:  data = WriteDelimited(cells, L'\t', true);  break;
    case SF_TEXT: data = WriteDelimited(cells, L'\t', false); break;
    case SF_SSF:
    {
        const GRIDCELL builtIn = BuiltInDefaultCell();
        data = WriteSsf(mapCells, gcs.nHeight, gcs.nWidth,
                        CellsEqual(m_defaultGridCell, builtIn) ? nullptr : &m_defaultGridCell);
        break;
    }
    case SF_XML:  data = WriteSpreadsheetML(cells); break;
    case SF_ODF:  data = WriteFlatOdf(cells); break;
    case SF_RTF:
    {
        std::vector<size_t> widths(gcs.nWidth);
        for (size_t c = 0; c < widths.size(); ++c)
            widths[c] = ColW(c);
        data = WriteRtf(cells, widths);
        break;
    }
    default:
        // SF_XLSX -- a genuine .xlsx is a ZIP package, which this
        // wide-character buffer cannot carry (see grid32.h) -- and any unknown
        // or combined format flags.
        finish(GRID_ERROR_NOT_IMPLEMENTED);
        return;
    }

    // Too small a buffer gets nothing rather than a truncated prefix, plus the
    // exact size needed, so the caller can retry once at the right size.
    if (data.size() >= bufWchars)
    {
        const unsigned long long needed = (data.size() + 1ULL) * sizeof(wchar_t);
        pStream->m_cbBuffOut = needed <= UINT_MAX ? static_cast<UINT>(needed) : 0;
        finish(GRID_ERROR_OUT_OF_RANGE);
        return;
    }
    wmemcpy(pStream->m_pwszBuff, data.c_str(), data.size());
    pStream->m_pwszBuff[data.size()] = L'\0';
    pStream->m_cbBuffOut = static_cast<UINT>(data.size() * sizeof(wchar_t));
    finish(0);
}

// Load grid contents, replacing what the grid held.
void CGrid32Mgr::OnStreamIn(LPGCSTREAM pStream)
{
    if (!pStream || pStream->m_cbSize != sizeof(GCSTREAM))
        return;

    auto finish = [pStream](DWORD error)
    {
        pStream->m_dwError = error;
        if (pStream->m_pfnCallback)
            pStream->m_pfnCallback(pStream);
    };
    if (!pStream->m_pwszBuff)
    {
        finish(GRID_ERROR_INVALID_PARAMETER);
        return;
    }
    const size_t maxChars = pStream->m_cbBuffSize / sizeof(wchar_t);
    const std::wstring_view input(pStream->m_pwszBuff, wcsnlen(pStream->m_pwszBuff, maxChars));

    // Parse everything before touching the grid, so a stream that fails to
    // parse leaves the current content exactly as it was.
    std::vector<StagedCell> staged;                         // interchange formats
    std::map<std::pair<UINT, UINT>, GRIDCELL> ssfCells;     // SF_SSF: finished cells
    GRIDCELL sheetDefault;                                  // SF_SSF: the DF record
    bool hasSheetDefault = false;
    DWORD error = 0;
    try
    {
        switch (pStream->m_dwFormat)
        {
        case SF_CSV:  ParseDelimited(input, L',', true, gcs.nHeight, gcs.nWidth, staged);   break;
        case SF_TSV:  ParseDelimited(input, L'\t', true, gcs.nHeight, gcs.nWidth, staged);  break;
        case SF_TEXT: ParseDelimited(input, L'\t', false, gcs.nHeight, gcs.nWidth, staged); break;
        case SF_SSF:  error = ParseSsf(input, BuiltInDefaultCell(), gcs.nHeight, gcs.nWidth, ssfCells,
                                       sheetDefault, hasSheetDefault); break;
        case SF_XML:
            if (!ParseSpreadsheetML(input, gcs.nHeight, gcs.nWidth, staged))
                error = GRID_ERROR_INVALID_PARAMETER;
            break;
        case SF_ODF:
            if (!ParseOdf(input, gcs.nHeight, gcs.nWidth, staged))
                error = GRID_ERROR_INVALID_PARAMETER;
            break;
        default:
            // SF_RTF, SF_XLSX (see grid32.h), and unknown or combined flags.
            error = GRID_ERROR_NOT_IMPLEMENTED;
            break;
        }
    }
    catch (const std::bad_alloc&)
    {
        error = GRID_ERROR_OUT_OF_MEMORY;
    }
    catch (const std::exception&)
    {
        error = GRID_ERROR_INVALID_PARAMETER;       // e.g. a non-numeric SSF value
    }
    if (error != 0)
    {
        finish(error);
        return;
    }

    // Replace: the load starts from an empty grid with no undo history and is
    // not itself undoable. Recalculation waits for the whole load, and one
    // GN_CONTENTCHANGED covers both the clear and the load.
    try
    {
        {
            FlagScope quiet(m_bSuppressChangeNotify, TRUE);
            ClearAll();
        }
        if (hasSheetDefault)
            m_defaultGridCell = sheetDefault;
        BulkEditScope bulkEdit(m_bDeferRecalc, m_bSuppressChangeNotify);
        FlagScope noUndo(m_bUndoRecordEnabled, FALSE);
        for (auto& entry : ssfCells)
        {
            std::unique_ptr<GRIDCELL> pCell(new GRIDCELL(std::move(entry.second)));
            PGRIDCELL& slot = mapCells[entry.first];
            delete slot;
            slot = pCell.release();
        }
        for (const StagedCell& sc : staged)
        {
            if (sc.literal)
            {
                std::unique_ptr<GRIDCELL> pCell(new GRIDCELL(m_defaultGridCell));
                pCell->m_wsText = sc.text;
                PGRIDCELL& slot = mapCells[std::make_pair(sc.row, sc.col)];
                delete slot;
                slot = pCell.release();
            }
            else
                SetCellText(sc.row, sc.col, sc.text.c_str());
            if (sc.format != FMT_GENERAL)
            {
                PGRIDCELL pCell = GetCell(sc.row, sc.col);
                if (pCell)
                {
                    pCell->m_nFormat = sc.format;
                    if (pCell->m_eType != CT_Text && !pCell->m_bFormula)
                        pCell->m_wsText = Grid32Detail::FormatCellValue(
                            pCell->m_eType, pCell->m_dValue, pCell->m_wsText, sc.format);
                }
            }
        }
    }
    catch (const std::bad_alloc&)
    {
        error = GRID_ERROR_OUT_OF_MEMORY;           // a partial load; say so
    }

    // Refresh evaluated text for any formula cells loaded above.
    RecalculateFormulas();
    NotifyContentChanged();
    Invalidate();
    finish(error);
}

void CGrid32Mgr::ClearAll()
{
    // An edit in progress belongs to the content being discarded. Cancel it
    // the way Escape does: committing it later would write its text into
    // the cleared sheet at the reset cursor.
    CancelEdit();

    const bool removedAny = !mapCells.empty() || !CellsEqual(m_defaultGridCell, BuiltInDefaultCell());
    DeleteAllCells();
    std::stack<GridEditOperation>().swap(m_undoStack);
    std::stack<GridEditOperation>().swap(m_redoStack);
    // The sheet's own sizes, filter and whole-sheet formatting go with it.
    m_colWidths.clear();
    m_rowHeights.clear();
    m_hiddenRows.clear();
    m_defaultGridCell = BuiltInDefaultCell();

    m_currentCell = { 0, 0 };
    m_selectionRect.start = m_selectionRect.end = m_currentCell;
    m_visibleTopLeft = { 0, 0 };
    SetScrollRanges();
    Invalidate();
    SetLastError(0);

    if (removedAny)
        NotifyContentChanged();
}


// The formula engine's text-result slot is private to Formulator.cpp, and
// CFormulator::Evaluate reports a text result in its Result. This file used to
// declare the slot's helpers extern and call them, while Formulator.cpp
// defined them inline: no out-of-line symbol was guaranteed, and at -O2 the
// link failed (AUDIT M23).
std::wstring CGrid32Mgr::EvaluateFormula(const std::wstring& expr)
{
    std::wstring clean = expr;
    if (!clean.empty() && clean[0] == L'=')
        clean = clean.substr(1);
    CFormulator::Result r = CFormulator::Evaluate(this, clean);
    return r.isText ? r.text : CFormulator::FormatGeneral(r.value);
}

// Store one evaluation result in its formula cell.
//
// m_dValue is the number itself. It used to be the display text parsed back,
// which lost everything past six significant digits ("1.23457e+06") and would
// read 0 the moment a number format turned 1234.5 into "$1,234.50".
//
// The display text is that number under the cell's own number format (H8):
// a formula cell formatted as currency shows "$1,234.50", and keeps doing so
// as it recalculates. A text result (LEFT, CONCAT, ...) is shown as-is, since
// number formats do not apply to text. So is a value FormatFixed cannot
// render: an infinity or NaN (=1/0), or a magnitude past what its integer
// arithmetic holds, keeps its general text rather than becoming garbage.
static void ApplyFormulaResult(PGRIDCELL pCell, const CFormulator::Result& r)
{
    pCell->m_eType = CT_Formula;
    pCell->m_dValue = r.value;
    if (r.isText)
    {
        pCell->m_wsText = r.text;
        return;
    }
    std::wstring general = CFormulator::FormatGeneral(r.value);
    const bool renderable = std::isfinite(r.value) && std::fabs(r.value) < 1e14;
    pCell->m_wsText = renderable
        ? Grid32Detail::FormatCellValue(CT_Formula, r.value, general, pCell->m_nFormat)
        : general;
}

// Re-evaluate a formula cell from its source. Both the display text and the
// cached numeric value are refreshed together: OnSortCells compares CT_Formula
// cells by m_dValue, so refreshing only the text would sort by whatever the
// value happened to be when the formula was first entered. Also re-renders the
// result when only the cell's number format changed.
void CGrid32Mgr::RefreshFormulaCell(PGRIDCELL pCell)
{
    if (!pCell || !pCell->m_bFormula)
        return;

    // The cell goes in as `self`, so a reference back to it is the same cycle
    // it is when RecalculateFormulas evaluates the cell by position.
    ApplyFormulaResult(pCell, CFormulator::Evaluate(this, pCell->m_wsFormula, pCell));
}

void CGrid32Mgr::RecalculateFormulas()
{
    // Bulk operations (sort, paste, stream-in) hold a BulkEditScope to
    // avoid a full recalc per cell; they invoke RecalculateFormulas()
    // themselves once the scope has closed.
    if (m_bDeferRecalc)
        return;

    // Single pass, order-independent: each formula is evaluated from its
    // source, and a reference to another formula cell recurses into that
    // cell's source rather than reading its (possibly stale) display text.
    //
    // The pass memoizes (P4). Without it a formula referenced N times was
    // evaluated N times — once as itself and once more inside every formula
    // that read it — so a chain whose cells each read the one before twice
    // cost 2^length. Now each formula cell is evaluated once per pass, and a
    // dependency chain longer than the parser's depth guard still resolves.
    //
    // Every formula is still recalculated on every call: the mutators call
    // this without saying which cell changed, so there is nothing to limit
    // the pass to.
    CFormulator::RecalcPass pass(this);
    for (auto& entry : mapCells)
    {
        PGRIDCELL pCell = entry.second;
        if (pCell && pCell->m_bFormula)
            ApplyFormulaResult(pCell,
                CFormulator::EvaluateCell(this, entry.first.first, entry.first.second));
    }
}
