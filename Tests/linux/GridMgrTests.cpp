// CGrid32Mgr tests — drives the real grid manager headless.
//
// Grid32Mgr.cpp is Win32 code, so these tests only build against the shim in
// win32shim/ (see README.md). They link the real Grid32Mgr.cpp, Formulator.cpp
// and EditWndProc.cpp against no-op Win32 stubs: no window is created, nothing
// is painted, and every GDI handle is null. That is enough to exercise the
// parts worth testing — the cell store, typing, recalculation, formats and
// import — which is what Tests/Grid32Tests.cpp calls out as deferred because
// it needs "a message-only host window".
//
// What this cannot cover: anything whose behaviour lives in Win32 itself
// (painting, hit-testing, scrolling, the clipboard, the edit subclass).

#define _CRT_SECURE_NO_WARNINGS
#include <pch.h>
#include <grid32_internal.h>
#include "win32shim/win32stubs.h"

#include <iostream>
#include <functional>
#include <cmath>
#include <string>
#include "Formulator.h"

// ---- tiny test harness (same shape as Tests/Grid32Tests.cpp) ---------------
struct TestCase {
    const char* name;
    std::function<void()> fn;
};
static std::vector<TestCase>& Tests() { static std::vector<TestCase> v; return v; }
static int g_failCount = 0;
static const char* g_currentTest = nullptr;

struct TestReg { TestReg(const char* n, std::function<void()> f) { Tests().push_back({n, f}); } };

#define TEST(name) \
    static void Test_##name(); \
    static TestReg Reg_##name(#name, Test_##name); \
    static void Test_##name()

#define FAIL_HERE(msg) do { \
    ++g_failCount; \
    std::cerr << "  FAIL [" << g_currentTest << "] " << __FILE__ << ":" << __LINE__ << "  " << msg << "\n"; \
    return; \
} while (0)

#define ASSERT_TRUE(cond) do { if (!(cond)) FAIL_HERE("expected true: " #cond); } while (0)
#define ASSERT_NEAR(a, b, eps) do { double _da = (a), _db = (b); \
    if (std::fabs(_da - _db) > (eps)) { std::cerr << "  (got " << _da << ", want " << _db << ")\n"; \
    FAIL_HERE("expected " #a " near " #b); } } while (0)
// Wide text goes to std::cerr as narrow text: writing to std::wcerr as well
// mixed stream orientations, and the FAIL lines went missing.
static std::string Narrow(const std::wstring& w)
{
    std::string out;
    for (wchar_t ch : w)
        out += (ch > 0 && ch < 128) ? (char)ch : '?';
    return out;
}
#define ASSERT_WSTR(expr, want) do { std::wstring _g = (expr); std::wstring _w = (want); \
    if (_g != _w) { std::cerr << "  (got \"" << Narrow(_g) << "\", want \"" << Narrow(_w) << "\")\n"; \
    FAIL_HERE("expected " #expr " == " #want); } } while (0)

// ---- fixture ---------------------------------------------------------------
// A grid with no real window behind it. Create() succeeds because every Win32
// call it makes is stubbed. Derives from CGrid32Mgr so the fixture can give
// m_hWndGrid a non-null value: notifications are suppressed for a null or
// dead HWND, so without it nothing would be observable.
class Sheet : private CGrid32Mgr
{
public:
    Sheet(size_t cols = 16, size_t rows = 64)
    {
        static int fakeHwnd = 0;
        m_hWndGrid = &fakeHwnd;
        GRIDCREATESTRUCT gcs{};
        gcs.cbSize = sizeof(GRIDCREATESTRUCT);
        gcs.nWidth = cols;
        gcs.nHeight = rows;
        gcs.nDefColWidth = 100;
        gcs.nDefRowHeight = 20;
        gcs.style = GS_SPREADSHEET;
        m_ok = Create(&gcs);
    }
    bool ok() const { return m_ok; }
    CGrid32Mgr& mgr() { return *this; }

    void Set(UINT row, UINT col, const wchar_t* text) { SetCellText(row, col, text); }
    std::wstring Text(UINT row, UINT col)
    {
        PGRIDCELL c = GetCell(row, col);
        return c ? c->m_wsText : std::wstring();
    }
    double Value(UINT row, UINT col)
    {
        PGRIDCELL c = GetCell(row, col);
        return c ? c->m_dValue : 0.0;
    }
    // State the public API doesn't expose.
    GRIDPOINT Cur() const { return m_currentCell; }
    GRIDSELECTION Sel() { return GetNormalizedSelection(); }
    void Select(UINT r0, UINT c0, UINT r1, UINT c1)
    {
        GRIDSELECTION sel{ { r0, c0 }, { r1, c1 } };
        OnSetSelection(&sel);
    }
    bool Editing() const { return m_bEditing; }
    size_t CellCount() const { return mapCells.size(); }
    const GRIDCELL& Default() const { return m_defaultGridCell; }
    // Moves the active cell without going through MoveCurrentCell, the way a
    // click or a scroll arrow used to while the editor still had focus.
    void PokeCurrentCell(UINT row, UINT col) { m_currentCell = { row, col }; }
private:
    bool m_ok;
};

// ---- typing and type inference ---------------------------------------------
TEST(SetCellText_StoresTextAndInfersType)
{
    Sheet s;
    ASSERT_TRUE(s.ok());
    s.Set(0, 0, L"hello");
    ASSERT_WSTR(s.Text(0, 0), L"hello");
    ASSERT_TRUE(s.mgr().GetCell(0, 0)->m_eType == CT_Text);

    s.Set(1, 0, L"42");
    ASSERT_TRUE(s.mgr().GetCell(1, 0)->m_eType == CT_Number);
    ASSERT_NEAR(s.Value(1, 0), 42.0, 1e-9);

    s.Set(2, 0, L"TRUE");
    ASSERT_TRUE(s.mgr().GetCell(2, 0)->m_eType == CT_Boolean);

    s.Set(3, 0, L"2026-05-15");
    ASSERT_TRUE(s.mgr().GetCell(3, 0)->m_eType == CT_Date);
}

TEST(SetCellText_FormulaStoresSourceAndResult)
{
    Sheet s;
    s.Set(0, 0, L"6");
    s.Set(1, 0, L"=A1*7");
    PGRIDCELL f = s.mgr().GetCell(1, 0);
    ASSERT_TRUE(f->m_bFormula);
    ASSERT_WSTR(f->m_wsFormula, L"A1*7");       // stored without the '='
    ASSERT_WSTR(f->m_wsText, L"42");
    ASSERT_TRUE(f->m_eType == CT_Formula);
    ASSERT_NEAR(f->m_dValue, 42.0, 1e-9);
}

// ---- recalculation (AUDIT_2026-09.md B2) -----------------------------------
TEST(EditingASource_RefreshesItsDependents)
{
    Sheet s;
    s.Set(0, 0, L"5");
    s.Set(1, 0, L"=A1*2");
    ASSERT_WSTR(s.Text(1, 0), L"10");

    s.Set(0, 0, L"7");                 // the regression: this used to leave "10"
    ASSERT_WSTR(s.Text(1, 0), L"14");
}

TEST(Recalc_RefreshesCachedNumericValue)
{
    // OnSortCells compares CT_Formula cells by m_dValue, so a recalc that
    // updated only the display text would sort by stale keys.
    Sheet s;
    s.Set(0, 0, L"5");
    s.Set(1, 0, L"=A1*2");
    s.Set(0, 0, L"7");
    ASSERT_NEAR(s.Value(1, 0), 14.0, 1e-9);
}

TEST(Recalc_PropagatesThroughAChain)
{
    Sheet s;
    s.Set(0, 0, L"5");
    s.Set(1, 0, L"=A1*2");
    s.Set(2, 0, L"=A2+1");
    ASSERT_WSTR(s.Text(2, 0), L"11");

    s.Set(0, 0, L"10");                // A1 -> A2 -> A3
    ASSERT_WSTR(s.Text(1, 0), L"20");
    ASSERT_WSTR(s.Text(2, 0), L"21");
}

TEST(Recalc_RefreshesRangeAggregates)
{
    Sheet s;
    s.Set(0, 1, L"1");
    s.Set(1, 1, L"2");
    s.Set(2, 1, L"3");
    s.Set(3, 1, L"=SUM(B1:B3)");
    ASSERT_WSTR(s.Text(3, 1), L"6");

    s.Set(1, 1, L"20");
    ASSERT_WSTR(s.Text(3, 1), L"24");
}

TEST(Recalc_RunsWhenASourceIsCleared)
{
    Sheet s;
    s.Set(0, 1, L"1");
    s.Set(1, 1, L"2");
    s.Set(2, 1, L"3");
    s.Set(3, 1, L"=SUM(B1:B3)");
    s.mgr().ClearCellText(1, 1);
    ASSERT_WSTR(s.Text(3, 1), L"4");
}

TEST(Recalc_RunsWhenASourceIsDeleted)
{
    Sheet s;
    s.Set(0, 0, L"4");
    s.Set(1, 0, L"=A1+1");
    ASSERT_WSTR(s.Text(1, 0), L"5");
    s.mgr().DeleteCell(0, 0);
    ASSERT_WSTR(s.Text(1, 0), L"1");
}

// ---- number formats --------------------------------------------------------
TEST(NumberFormat_RerendersTheDisplayText)
{
    Sheet s;
    s.Set(0, 0, L"1234.5");
    s.mgr().SetCellNumberFormat(0, 0, FMT_CURRENCY);
    ASSERT_WSTR(s.Text(0, 0), L"$1,234.50");
    ASSERT_NEAR(s.Value(0, 0), 1234.5, 1e-9);   // underlying value preserved
}

// ---- stream-in -------------------------------------------------------------
static void StreamInCsv(Sheet& s, const std::wstring& csv, DWORD fmt = SF_CSV)
{
    std::vector<wchar_t> buf(csv.begin(), csv.end());
    buf.push_back(L'\0');
    GCSTREAM st{};
    st.m_cbSize = sizeof(GCSTREAM);
    st.m_pwszBuff = buf.data();
    st.m_cbBuffSize = (UINT)(buf.size() * sizeof(wchar_t));
    st.m_dwFormat = fmt;
    s.mgr().OnStreamIn(&st);
}

TEST(StreamIn_Csv_LoadsCells)
{
    Sheet s;
    StreamInCsv(s, L"1,2\n3,4\n");
    ASSERT_WSTR(s.Text(0, 0), L"1");
    ASSERT_WSTR(s.Text(0, 1), L"2");
    ASSERT_WSTR(s.Text(1, 0), L"3");
    ASSERT_WSTR(s.Text(1, 1), L"4");
}

TEST(StreamIn_Csv_HonoursQuotedFields)
{
    Sheet s;
    StreamInCsv(s, L"\"a,b\",plain\n");
    ASSERT_WSTR(s.Text(0, 0), L"a,b");
    ASSERT_WSTR(s.Text(0, 1), L"plain");
}

TEST(StreamIn_RecalculatesOnce_AndFormulasSeeTheImport)
{
    // The import defers per-cell recalculation and runs one pass at the end,
    // so a formula ahead of its inputs in the file still sees them. Stream-in
    // replaces the sheet (H15), so the formula comes from the import itself.
    Sheet s;
    StreamInCsv(s, L"=A2+B2\n1,2\n3,4\n");
    ASSERT_WSTR(s.Text(0, 0), L"3");
}

// ---- undo / redo (AUDIT_2026-09.md B5) -------------------------------------
TEST(Undo_RestoresAFormula)
{
    Sheet s;
    s.Set(0, 0, L"4");            // A1
    s.Set(1, 0, L"=A1+1");        // A2 = 5
    ASSERT_WSTR(s.Text(1, 0), L"5");

    s.Set(1, 0, L"99");           // overwrite the formula with a literal
    ASSERT_WSTR(s.Text(1, 0), L"99");
    ASSERT_TRUE(!s.mgr().GetCell(1, 0)->m_bFormula);

    s.mgr().OnUndo();
    // Used to come back as the literal text "5" with the formula gone.
    ASSERT_TRUE(s.mgr().GetCell(1, 0)->m_bFormula);
    ASSERT_WSTR(s.mgr().GetCell(1, 0)->m_wsFormula, L"A1+1");
    ASSERT_WSTR(s.Text(1, 0), L"5");
}

TEST(Undo_RestoredFormulaIsLive)
{
    // The real test of B5: a restored formula must recalculate, not just look
    // like a formula. Restoring the display text would pass the check above
    // if it happened to match, but would fail this one.
    Sheet s;
    s.Set(0, 0, L"4");
    s.Set(1, 0, L"=A1+1");
    s.Set(1, 0, L"99");
    s.mgr().OnUndo();

    s.Set(0, 0, L"10");           // move the source
    ASSERT_WSTR(s.Text(1, 0), L"11");
    ASSERT_NEAR(s.Value(1, 0), 11.0, 1e-9);
}

TEST(Redo_ReappliesTheOverwrite)
{
    Sheet s;
    s.Set(0, 0, L"4");
    s.Set(1, 0, L"=A1+1");
    s.Set(1, 0, L"99");
    s.mgr().OnUndo();
    ASSERT_TRUE(s.mgr().GetCell(1, 0)->m_bFormula);

    s.mgr().OnRedo();
    ASSERT_TRUE(!s.mgr().GetCell(1, 0)->m_bFormula);
    ASSERT_WSTR(s.Text(1, 0), L"99");
}

TEST(Undo_RestoresNumberFormat)
{
    // SetFormat undo restored fontInfo only, so the number format stayed put.
    Sheet s;
    s.Set(0, 0, L"1234.5");
    ASSERT_WSTR(s.Text(0, 0), L"1234.5");

    s.mgr().SetCellNumberFormat(0, 0, FMT_CURRENCY);
    ASSERT_WSTR(s.Text(0, 0), L"$1,234.50");

    s.mgr().OnUndo();
    ASSERT_TRUE(s.mgr().GetCell(0, 0)->m_nFormat == FMT_GENERAL);
    ASSERT_WSTR(s.Text(0, 0), L"1234.5");
}

TEST(Undo_RestoresDeletedCell)
{
    Sheet s;
    s.Set(0, 0, L"=1+2");
    ASSERT_WSTR(s.Text(0, 0), L"3");
    s.mgr().DeleteCell(0, 0);
    ASSERT_TRUE(s.mgr().GetCell(0, 0) == nullptr);

    s.mgr().OnUndo();
    ASSERT_TRUE(s.mgr().GetCell(0, 0) != nullptr);
    ASSERT_TRUE(s.mgr().GetCell(0, 0)->m_bFormula);
    ASSERT_WSTR(s.Text(0, 0), L"3");
}

TEST(Undo_RefreshesDependents)
{
    Sheet s;
    s.Set(0, 0, L"5");
    s.Set(1, 0, L"=A1*2");        // 10
    s.Set(0, 0, L"7");            // dependent becomes 14
    ASSERT_WSTR(s.Text(1, 0), L"14");

    s.mgr().OnUndo();             // A1 back to 5
    ASSERT_WSTR(s.Text(0, 0), L"5");
    ASSERT_WSTR(s.Text(1, 0), L"10");
}

TEST(Undo_StackDepthTracksEdits)
{
    Sheet s;
    ASSERT_TRUE(s.mgr().OnCanUndo() == 0);
    s.Set(0, 0, L"a");
    s.Set(0, 0, L"b");
    ASSERT_TRUE(s.mgr().OnCanUndo() == 2);

    s.mgr().OnUndo();
    ASSERT_TRUE(s.mgr().OnCanUndo() == 1);
    ASSERT_TRUE(s.mgr().OnCanRedo() == 1);

    s.mgr().OnRedo();
    ASSERT_TRUE(s.mgr().OnCanUndo() == 2);
    ASSERT_TRUE(s.mgr().OnCanRedo() == 0);
}

TEST(Undo_SendsOneNotification)
{
    Sheet s;
    s.Set(0, 0, L"x");
    StubMessageLogClear();
    s.mgr().OnUndo();
    ASSERT_TRUE(StubCountNotifications(GN_CONTENTCHANGED) == 1);
}

// ---- content-changed notification (AUDIT_2026-09.md B4) --------------------
// The host marks its document dirty off GN_CONTENTCHANGED, so every edit that
// alters what StreamOut would write has to send one.
static size_t Notifications() { return StubCountNotifications(GN_CONTENTCHANGED); }

TEST(Notify_TypingSendsOne)
{
    Sheet s;
    StubMessageLogClear();
    s.Set(0, 0, L"hello");
    ASSERT_TRUE(Notifications() == 1);
}

TEST(Notify_EachEditSendsOne)
{
    Sheet s;
    StubMessageLogClear();
    s.Set(0, 0, L"1");
    s.Set(1, 0, L"2");
    s.Set(2, 0, L"3");
    ASSERT_TRUE(Notifications() == 3);
}

TEST(Notify_ClearAndDeleteSend)
{
    Sheet s;
    s.Set(0, 0, L"x");
    StubMessageLogClear();
    s.mgr().ClearCellText(0, 0);
    ASSERT_TRUE(Notifications() == 1);

    s.Set(1, 0, L"y");
    StubMessageLogClear();
    s.mgr().DeleteCell(1, 0);
    ASSERT_TRUE(Notifications() == 1);
}

TEST(Notify_FormattingSends)
{
    Sheet s;
    s.Set(0, 0, L"1234.5");
    StubMessageLogClear();
    s.mgr().SetCellNumberFormat(0, 0, FMT_CURRENCY);
    ASSERT_TRUE(Notifications() == 1);

    CELLFONT fi{};
    fi.m_wsFontFace = L"Arial";
    fi.m_fPointSize = 12.0f;
    StubMessageLogClear();
    s.mgr().SetCellFormat(0, 0, fi);
    ASSERT_TRUE(Notifications() == 1);
}

TEST(Notify_BulkImportSendsExactlyOne)
{
    // 4 cells imported, but it is one user action: the bulk scope suppresses
    // the per-cell notifications and the import sends a single one.
    Sheet s;
    StubMessageLogClear();
    StreamInCsv(s, L"1,2\n3,4\n");
    ASSERT_TRUE(Notifications() == 1);
}

TEST(Notify_NavigationDoesNotSend)
{
    // Moving the cursor or the selection is not a document change.
    Sheet s;
    s.Set(0, 0, L"x");
    StubMessageLogClear();
    s.mgr().SetCurrentCell(2, 2);
    GRIDSELECTION sel{ {0, 0}, {3, 3} };
    s.mgr().OnSetSelection(&sel);
    ASSERT_TRUE(Notifications() == 0);
}

TEST(Notify_FailedReadsDoNotSend)
{
    Sheet s;
    s.Set(0, 0, L"5");
    StubMessageLogClear();
    s.Text(0, 0);
    s.Value(0, 0);
    s.mgr().GetCell(9, 9);
    ASSERT_TRUE(Notifications() == 0);
}

// ---- AUDIT_2026-10: selection and keyboard (S1) -----------------------------
static bool SameSel(const GRIDSELECTION& a, UINT r0, UINT c0, UINT r1, UINT c1)
{
    return a.start.nRow == r0 && a.start.nCol == c0 && a.end.nRow == r1 && a.end.nCol == c1;
}

TEST(Selection_KeyboardMoveCollapsesTheSelection)
{
    // Tab from a multi-cell selection used to leave the selection behind, so
    // the ribbon formatted cells the cursor had left (S1).
    Sheet s;
    s.Select(0, 0, 3, 3);
    s.mgr().IncrementSelectedCell(0, 1);
    ASSERT_TRUE(s.Cur().nRow == 0 && s.Cur().nCol == 1);
    ASSERT_TRUE(SameSel(s.Sel(), 0, 1, 0, 1));
}

TEST(Selection_ShiftMoveExtendsFromTheActiveCell)
{
    Sheet s;
    s.mgr().SetCurrentCell(1, 1);
    s.mgr().IncrementSelectedCell(1, 0, true);
    s.mgr().IncrementSelectedCell(1, 0, true);
    s.mgr().IncrementSelectedCell(0, 1, true);
    ASSERT_TRUE(s.Cur().nRow == 1 && s.Cur().nCol == 1);     // the anchor stays
    ASSERT_TRUE(SameSel(s.Sel(), 1, 1, 3, 2));
}

TEST(Selection_EachMoveSendsOneSelChange)
{
    Sheet s;
    StubMessageLogClear();
    s.mgr().IncrementSelectedCell(1, 0);
    ASSERT_TRUE(StubCountNotifications(GN_SELCHANGE) == 1);
    StubMessageLogClear();
    s.mgr().IncrementSelectedCell(-1, 0);
    s.mgr().IncrementSelectedCell(-1, 0);       // already at row 0: no move, no notification
    ASSERT_TRUE(StubCountNotifications(GN_SELCHANGE) == 1);
}

TEST(Selection_SetSelectionClampsInsteadOfRejecting)
{
    // Select All sent (UINT)-1 for the far corner and was rejected (H4).
    Sheet s(16, 64);
    s.Select(0, 0, (UINT)-1, (UINT)-1);
    ASSERT_TRUE(SameSel(s.Sel(), 0, 0, 63, 15));
}

// ---- AUDIT_2026-10: the inline editor (S2, S11, S12, H14) -------------------
TEST(Edit_CommitsIntoTheCellItWasOpenedOn)
{
    // A click moved the current cell without taking focus from the editor,
    // and Enter then wrote the text into the clicked cell (S2).
    Sheet s;
    s.mgr().SetCurrentCell(0, 0);
    s.mgr().BeginEdit(false);
    s.PokeCurrentCell(5, 5);
    s.mgr().CommitEditText(L"hello");
    ASSERT_WSTR(s.Text(0, 0), L"hello");
    ASSERT_TRUE(s.mgr().GetCell(5, 5) == nullptr);
    ASSERT_TRUE(!s.Editing());
}

TEST(Edit_EnterOnTheGridDoesNotTouchTheCell)
{
    // Enter opened the editor and committed an untouched cell: a new empty
    // cell, an undo step and a dirty document (S11).
    Sheet s;
    StubMessageLogClear();
    ASSERT_TRUE(s.mgr().OnKeyDown(VK_RETURN, 1, 0));
    ASSERT_TRUE(!s.Editing());
    ASSERT_TRUE(s.CellCount() == 0);
    ASSERT_TRUE(s.mgr().OnCanUndo() == 0);
    ASSERT_TRUE(StubCountNotifications(GN_CONTENTCHANGED) == 0);
}

TEST(Edit_UnchangedCommitWritesNothing)
{
    Sheet s;
    s.Set(0, 0, L"x");
    StubMessageLogClear();
    s.mgr().BeginEdit(true);                    // F2
    s.mgr().CommitEditText(L"x");
    ASSERT_TRUE(s.mgr().OnCanUndo() == 1);
    ASSERT_TRUE(StubCountNotifications(GN_CONTENTCHANGED) == 0);
}

TEST(Edit_FunctionKeysDoNotOpenTheEditor)
{
    // F1 is VK 0x70, iswalnum(L'p'): it opened the editor (H14). Editing now
    // starts from WM_CHAR.
    Sheet s;
    ASSERT_TRUE(!s.mgr().OnKeyDown(VK_F1, 1, 0));
    ASSERT_TRUE(!s.Editing());
    ASSERT_TRUE(s.mgr().OnChar(L'p'));
    ASSERT_TRUE(s.Editing());
}

TEST(Edit_TypingReplacesBackspaceClears)
{
    Sheet s;
    s.Set(0, 0, L"old");
    s.mgr().OnChar(L'n');
    s.mgr().CommitEditText(L"new");
    ASSERT_WSTR(s.Text(0, 0), L"new");

    s.mgr().OnChar(L'\b');                      // Backspace + Enter clears
    s.mgr().CommitEditText(L"");
    ASSERT_WSTR(s.Text(0, 0), L"");
}

TEST(Edit_EscapeCancelsWithoutArmingAnything)
{
    // Escape set its suppress flag after notifying, and the flag then
    // swallowed the next real commit (S12).
    Sheet s;
    s.Set(0, 0, L"keep");
    s.mgr().BeginEdit(false);
    StubMessageLogClear();
    s.mgr().CancelEdit();
    ASSERT_TRUE(!s.Editing());
    ASSERT_TRUE(StubCountNotifications(GN_EDITCANCEL) == 1);
    ASSERT_WSTR(s.Text(0, 0), L"keep");

    s.mgr().BeginEdit(false);
    s.mgr().CommitEditText(L"next");
    ASSERT_WSTR(s.Text(0, 0), L"next");
}

// ---- AUDIT_2026-10: Clear (S4, S17) -------------------------------------------
TEST(Clear_RemovesContentAndFormulas_KeepsFormatting)
{
    Sheet s;
    s.Set(0, 0, L"1");
    s.Set(1, 0, L"=A1*2");
    s.Set(2, 0, L"x");
    CELLFONT bold{};
    bold.bWeight = 700;
    s.mgr().SetCellFormat(2, 0, bold, GCFM_BOLD);
    const size_t undoBefore = s.mgr().OnCanUndo();

    s.Select(0, 0, 9, 2);                       // mostly empty positions
    s.mgr().OnClear();
    ASSERT_TRUE(s.mgr().GetCell(0, 0) == nullptr);
    ASSERT_TRUE(s.mgr().GetCell(1, 0) == nullptr);      // the formula is gone too
    ASSERT_TRUE(s.mgr().GetCell(2, 0) != nullptr);      // formatting survives
    ASSERT_WSTR(s.Text(2, 0), L"");
    ASSERT_TRUE(s.mgr().GetCell(2, 0)->fontInfo.bWeight == 700);
    ASSERT_TRUE(s.CellCount() == 1);                    // nothing created
    ASSERT_TRUE(s.mgr().OnCanUndo() == undoBefore + 1); // one step
}

TEST(Clear_UndoRestoresEveryCell)
{
    // WM_CLEAR's undo record had no coordinates and restored nothing (S17).
    Sheet s;
    s.Set(0, 0, L"1");
    s.Set(1, 0, L"=A1*2");
    s.Select(0, 0, 1, 0);
    s.mgr().OnClear();
    s.mgr().OnUndo();
    ASSERT_WSTR(s.Text(0, 0), L"1");
    ASSERT_TRUE(s.mgr().GetCell(1, 0) && s.mgr().GetCell(1, 0)->m_bFormula);
    ASSERT_WSTR(s.Text(1, 0), L"2");
    s.mgr().OnRedo();
    ASSERT_TRUE(s.CellCount() == 0);
}

TEST(Clear_DependentsRecalculate)
{
    Sheet s;
    s.Set(0, 0, L"5");
    s.Set(0, 1, L"=A1+1");
    s.Select(0, 0, 0, 0);
    s.mgr().OnClear();
    ASSERT_WSTR(s.Text(0, 1), L"1");
}

// ---- AUDIT_2026-10: find and replace (S6, H5, M6, M11, P2) ------------------
static GCREPLACESTRUCT Replace(const wchar_t* find, const wchar_t* with)
{
    GCREPLACESTRUCT rs{};
    rs.m_cbSize = sizeof(rs);
    rs.m_pwszFindText = find;
    rs.m_pwszReplaceText = with;
    rs.m_bSearchForward = TRUE;
    return rs;
}

TEST(ReplaceAll_ReplacesEveryOccurrenceAndFinishes)
{
    // Replacing "a" with "ab" re-matched the cell it had just changed and
    // never finished (H5); only the first occurrence per cell changed (S6).
    Sheet s;
    s.Set(0, 0, L"a a");
    s.Set(1, 0, L"ba");
    s.Set(2, 0, L"none");
    const size_t undoBefore = s.mgr().OnCanUndo();
    ASSERT_TRUE(s.mgr().OnReplaceAll(Replace(L"a", L"ab")) == 3);
    ASSERT_WSTR(s.Text(0, 0), L"ab ab");
    ASSERT_WSTR(s.Text(1, 0), L"bab");
    ASSERT_TRUE(s.mgr().OnCanUndo() == undoBefore + 1);
    s.mgr().OnUndo();
    ASSERT_WSTR(s.Text(0, 0), L"a a");
    ASSERT_WSTR(s.Text(1, 0), L"ba");
}

TEST(Replace_EditsTheFormulaNotItsResult)
{
    // Replacing in a formula cell wrote its result over the formula (M11).
    Sheet s;
    s.Set(0, 0, L"2");
    s.Set(0, 1, L"=A1*3");
    const size_t undoBefore = s.mgr().OnCanUndo();
    ASSERT_TRUE(s.mgr().OnReplaceText(Replace(L"3", L"4")));
    ASSERT_TRUE(s.mgr().GetCell(0, 1)->m_bFormula);
    ASSERT_WSTR(s.mgr().GetCell(0, 1)->m_wsFormula, L"A1*4");
    ASSERT_WSTR(s.Text(0, 1), L"8");
    ASSERT_TRUE(s.mgr().OnCanUndo() == undoBefore + 1);     // not two (M6)
}

TEST(Find_SearchesFormulaSourceAndMovesTheCell)
{
    Sheet s(16, 64);
    s.Set(40, 3, L"=SUM(A1:A2)");
    GCFINDSTRUCT fs{};
    fs.m_cbSize = sizeof(fs);
    fs.m_pwszFindText = L"sum(";
    fs.m_bSearchForward = TRUE;
    ASSERT_TRUE(s.mgr().OnFindText(fs));
    ASSERT_TRUE(s.Cur().nRow == 40 && s.Cur().nCol == 3);
    ASSERT_TRUE(SameSel(s.Sel(), 40, 3, 40, 3));
    fs.m_startCell = { 40, 4 };
    ASSERT_TRUE(!s.mgr().OnFindText(fs));       // nothing after it
}

// ---- AUDIT_2026-10: fill (S7, M12, M13, M24) --------------------------------
static GCFILLSTRUCT FillDir(DWORD dir)
{
    GCFILLSTRUCT fs{};
    fs.m_cbSize = sizeof(GCFILLSTRUCT);
    fs.m_dwFillHow = FILLBYREFCELL;
    fs.m_dwFillWhat = FILL_ALL;
    fs.m_dwFillRange = FILLBYCURRENTSELECTION;
    fs.m_dwFillDirection = dir;
    return fs;
}

TEST(Fill_DownCopiesEachColumnsTopCell)
{
    Sheet s;
    s.Set(0, 0, L"1");
    s.Set(0, 1, L"x");
    s.Select(2, 1, 0, 0);                       // dragged upward: not normalised by the host
    const size_t undoBefore = s.mgr().OnCanUndo();
    ASSERT_TRUE(s.mgr().OnFillCells(0, FillDir(FILL_DOWN)));
    ASSERT_WSTR(s.Text(2, 0), L"1");
    ASSERT_TRUE(s.mgr().GetCell(2, 0)->m_eType == CT_Number);   // typed, not text (M12)
    ASSERT_WSTR(s.Text(2, 1), L"x");
    ASSERT_TRUE(s.mgr().OnCanUndo() == undoBefore + 1);         // one step (M24)
    s.mgr().OnUndo();
    ASSERT_TRUE(s.mgr().GetCell(1, 0) == nullptr && s.mgr().GetCell(2, 1) == nullptr);
}

TEST(Fill_RightAndFormulas)
{
    Sheet s;
    s.Set(0, 0, L"=1+1");
    s.Select(0, 0, 0, 2);
    ASSERT_TRUE(s.mgr().OnFillCells(0, FillDir(FILL_RIGHT)));
    ASSERT_TRUE(s.mgr().GetCell(0, 2)->m_bFormula);
    ASSERT_WSTR(s.Text(0, 2), L"2");
}

TEST(Fill_ReservedWParamIsRejected)
{
    Sheet s;
    s.Set(0, 0, L"1");
    s.Select(0, 0, 3, 0);
    ASSERT_TRUE(!s.mgr().OnFillCells(1, FillDir(FILL_DOWN)));
    ASSERT_TRUE(s.mgr().GetCell(3, 0) == nullptr);
}

// ---- AUDIT_2026-10: sort (P3, M13, M24) ---------------------------------------
TEST(Sort_MovesOnlyCellsThatExist)
{
    // Sorting copied every column of every row back, filling the map with
    // empty cells (P3), and could not be undone (M24).
    Sheet s;
    s.Set(0, 0, L"3"); s.Set(0, 1, L"c");
    s.Set(1, 0, L"1"); s.Set(1, 1, L"a");
    s.Set(2, 0, L"2"); s.Set(2, 1, L"b");
    GCSORTSTRUCT ss{};
    ss.m_cbSize = sizeof(ss);
    ss.m_sortSel = { { 0, 0 }, { 2, 0 } };
    ss.m_bAscending = TRUE;
    ASSERT_TRUE(s.mgr().OnSortCells(0, ss));
    ASSERT_WSTR(s.Text(0, 0), L"1"); ASSERT_WSTR(s.Text(0, 1), L"a");
    ASSERT_WSTR(s.Text(2, 0), L"3"); ASSERT_WSTR(s.Text(2, 1), L"c");
    ASSERT_TRUE(s.CellCount() == 6);
    s.mgr().OnUndo();
    ASSERT_WSTR(s.Text(0, 0), L"3"); ASSERT_WSTR(s.Text(0, 1), L"c");
}

TEST(Sort_KeyColumnAndBlanksLast)
{
    Sheet s;
    s.Set(0, 1, L"b"); s.Set(0, 0, L"row0");
    s.Set(1, 0, L"row1");                       // B2 blank
    s.Set(2, 1, L"A"); s.Set(2, 0, L"row2");
    GCSORTSTRUCT ss{};
    ss.m_cbSize = sizeof(ss);
    ss.m_sortSel = { { 0, 0 }, { 2, 1 } };
    ss.m_bAscending = FALSE;
    ss.m_dwWhich = GCS_KEYCOLUMN(1);            // M13: the key is honoured
    ASSERT_TRUE(s.mgr().OnSortCells(0, ss));
    ASSERT_WSTR(s.Text(0, 0), L"row0");         // "b" > "A" case-insensitively
    ASSERT_WSTR(s.Text(1, 0), L"row2");
    ASSERT_WSTR(s.Text(2, 0), L"row1");         // the blank stays last descending
}

// ---- AUDIT_2026-10 S10 and AUDIT_2026-09 H7/H9/M6/M8: formatting --------------
TEST(Format_WholeGridStylesTheDefaultCellNotEveryPosition)
{
    Sheet s(702, 4096);
    s.Set(0, 0, L"x");
    s.mgr().OnSetSelection(nullptr);            // Select All
    CELLFONT bold{};
    bold.bWeight = 700;
    s.mgr().SetSelectionFormat(bold, GCFM_BOLD);
    ASSERT_TRUE(s.CellCount() == 1);
    ASSERT_TRUE(s.Default().fontInfo.bWeight == 700);
    ASSERT_TRUE(s.mgr().GetCell(0, 0)->fontInfo.bWeight == 700);
    s.mgr().OnUndo();
    ASSERT_TRUE(s.Default().fontInfo.bWeight == 400);
    ASSERT_TRUE(s.mgr().GetCell(0, 0)->fontInfo.bWeight == 400);
}

TEST(Format_MaskKeepsTheRestOfEachCellsFont)
{
    // Bold over a mixed selection made every cell the active cell's font (H7).
    Sheet s;
    CELLFONT times{};
    times.m_wsFontFace = L"Times";
    times.m_fPointSize = 20.0f;
    s.mgr().SetCellFormat(0, 0, times, GCFM_FACE | GCFM_SIZE);
    s.Select(0, 0, 1, 0);
    CELLFONT bold{};
    bold.bWeight = 700;
    s.mgr().SetSelectionFormat(bold, GCFM_BOLD);
    ASSERT_WSTR(s.mgr().GetCell(0, 0)->fontInfo.m_wsFontFace, L"Times");
    ASSERT_NEAR(s.mgr().GetCell(0, 0)->fontInfo.m_fPointSize, 20.0, 1e-6);
    ASSERT_WSTR(s.mgr().GetCell(1, 0)->fontInfo.m_wsFontFace, L"Arial");
    ASSERT_TRUE(s.mgr().GetCell(0, 0)->fontInfo.bWeight == 700);
    ASSERT_TRUE(s.mgr().GetCell(1, 0)->fontInfo.bWeight == 700);
}

TEST(Format_AlignmentIsOneUndoStep)
{
    Sheet s;
    s.Set(0, 0, L"x");
    s.Select(0, 0, 2, 2);
    const size_t undoBefore = s.mgr().OnCanUndo();
    s.mgr().SetSelectionHAlign(DT_RIGHT);
    ASSERT_TRUE(s.mgr().OnCanUndo() == undoBefore + 1);     // H9
    s.mgr().OnUndo();
    ASSERT_TRUE((s.mgr().GetCell(0, 0)->justification & DT_RIGHT) == 0);
    ASSERT_TRUE(s.CellCount() == 1);            // the cells it created are gone again
}

TEST(Format_NumberFormatIsOneUndoStep)
{
    // GetCellOrCreate(..., true) recorded a second step (M6).
    Sheet s;
    s.Set(0, 0, L"5");
    const size_t undoBefore = s.mgr().OnCanUndo();
    s.mgr().SetCellNumberFormat(0, 0, FMT_PERCENT);
    ASSERT_TRUE(s.mgr().OnCanUndo() == undoBefore + 1);
}

TEST(Format_FontChangeDoesNotRecalculate)
{
    Sheet s;
    s.Set(0, 0, L"1");
    s.Set(1, 0, L"=A1+1");
    const size_t before = CFormulator::CellEvaluationCount();
    CELLFONT bold{};
    bold.bWeight = 700;
    s.mgr().SetCellFormat(0, 0, bold, GCFM_BOLD);           // M8
    ASSERT_TRUE(CFormulator::CellEvaluationCount() == before);
}

TEST(Format_BackgroundAppliesAndUndoes)
{
    Sheet s;
    s.Select(1, 1, 2, 2);
    s.mgr().SetSelectionBackground(RGB(255, 255, 0));       // H11
    ASSERT_TRUE(s.mgr().GetCell(2, 2)->clrBackground == RGB(255, 255, 0));
    s.mgr().OnUndo();
    ASSERT_TRUE(s.CellCount() == 0);
}

// ---- AUDIT_2026-09 H10: merged cells ------------------------------------------
TEST(Merge_BlockActsAsOneCell)
{
    Sheet s;
    s.Set(0, 0, L"title");
    ASSERT_TRUE(s.mgr().MergeCells({ { 0, 0 }, { 1, 2 } }));
    const MERGERANGE* m = s.mgr().FindMerge(1, 2);
    ASSERT_TRUE(m && m->start.nRow == 0 && m->start.nCol == 0 && m->end.nRow == 1 && m->end.nCol == 2);
    s.mgr().SetCurrentCell(1, 1);               // inside the block: lands on its anchor
    ASSERT_TRUE(s.Cur().nRow == 0 && s.Cur().nCol == 0);
    ASSERT_TRUE(SameSel(s.Sel(), 0, 0, 1, 2));
    s.mgr().IncrementSelectedCell(0, 1);        // leaves from the block's far edge
    ASSERT_TRUE(s.Cur().nRow == 0 && s.Cur().nCol == 3);
    s.mgr().OnUndo();
    ASSERT_TRUE(s.mgr().FindMerge(1, 2) == nullptr);
}

TEST(Merge_UnmergeAndPersistOnTheAnchor)
{
    Sheet s;
    s.mgr().MergeCells({ { 2, 2 }, { 3, 3 } });
    ASSERT_TRUE(s.mgr().GetCell(2, 2)->mergeRange.end.nRow == 3);   // what SSF saves
    ASSERT_TRUE(s.mgr().UnmergeCells({ { 3, 3 }, { 3, 3 } }));
    ASSERT_TRUE(!s.mgr().IsCellMerged(2, 2));
}

// ---- AUDIT_2026-09 H16: filtering ---------------------------------------------
TEST(Filter_HidesNonMatchingRowsAndShowsThemAgain)
{
    Sheet s;
    s.Set(0, 0, L"Fruit");
    s.Set(1, 0, L"apple");
    s.Set(2, 0, L"banana");
    s.Set(3, 0, L"Apricot");
    GCFILTERSTRUCT fs{};
    fs.m_cbSize = sizeof(fs);
    fs.m_filterColumn = { 0, 0 };
    fs.m_pwszFilterText = L"ap";
    fs.m_dwFilterType = FILTER_STARTS_WITH;
    ASSERT_TRUE(s.mgr().OnFilterCells(0, fs));
    ASSERT_TRUE(!s.mgr().IsRowHidden(0));       // the header
    ASSERT_TRUE(!s.mgr().IsRowHidden(1));
    ASSERT_TRUE(s.mgr().IsRowHidden(2));
    ASSERT_TRUE(!s.mgr().IsRowHidden(3));       // case-insensitive
    ASSERT_TRUE(s.mgr().RowH(2) == 0);
    s.mgr().SetCurrentCell(1, 0);
    s.mgr().IncrementSelectedCell(1, 0);        // steps over the hidden row
    ASSERT_TRUE(s.Cur().nRow == 3);
    fs.m_dwFilterType = FILTER_NONE;
    ASSERT_TRUE(s.mgr().OnFilterCells(0, fs));
    ASSERT_TRUE(!s.mgr().IsRowHidden(2));
}

// ---- AUDIT_2026-10 S16: Go To -------------------------------------------------
TEST(GoTo_AcceptsCellReferences)
{
    Sheet s(16, 64);
    ASSERT_TRUE(s.mgr().OnSetCurrentCell(L"$C$5", SETBYCOORDINATE));
    ASSERT_TRUE(s.Cur().nRow == 4 && s.Cur().nCol == 2);
    ASSERT_TRUE(s.mgr().OnSetCurrentCell(L" b2 ", SETBYCOORDINATE));
    ASSERT_TRUE(s.Cur().nRow == 1 && s.Cur().nCol == 1);
    ASSERT_TRUE(s.mgr().OnSetCurrentCell(L"A1:C3", SETBYCOORDINATE));
    ASSERT_TRUE(SameSel(s.Sel(), 0, 0, 2, 2));
}

TEST(GoTo_RejectsBadReferences)
{
    Sheet s(16, 64);
    s.mgr().SetCurrentCell(3, 3);
    ASSERT_TRUE(!s.mgr().OnSetCurrentCell(L"A0", SETBYCOORDINATE));
    ASSERT_TRUE(!s.mgr().OnSetCurrentCell(L"Q1", SETBYCOORDINATE));     // column 17 of 16
    ASSERT_TRUE(!s.mgr().OnSetCurrentCell(L"A65", SETBYCOORDINATE));
    ASSERT_TRUE(!s.mgr().OnSetCurrentCell(L"1A", SETBYCOORDINATE));
    ASSERT_TRUE(!s.mgr().OnSetCurrentCell(L"nobody", SETBYNAME));
    ASSERT_TRUE(s.Cur().nRow == 3 && s.Cur().nCol == 3);
}

// ---- AUDIT_2026-09 B6, M10: clipboard payloads ----------------------------------
TEST(Paste_CrLfTextLandsTyped)
{
    // "42\r" landed as text: the clipboard's CRLF left the \r behind (B6).
    Sheet s;
    s.mgr().PasteText(L"1\t2\r\n3\t42\r\n");
    ASSERT_WSTR(s.Text(1, 1), L"42");
    ASSERT_TRUE(s.mgr().GetCell(1, 1)->m_eType == CT_Number);
    ASSERT_TRUE(s.mgr().GetCell(2, 0) == nullptr);      // the trailing break is not a row
    ASSERT_TRUE(SameSel(s.Sel(), 0, 0, 1, 1));
}

TEST(Paste_QuotedFieldsAndOneUndoStep)
{
    Sheet s;
    const size_t undoBefore = s.mgr().OnCanUndo();
    s.mgr().PasteText(L"\"a\tb\"\t\"say \"\"hi\"\"\"\n");
    ASSERT_WSTR(s.Text(0, 0), L"a\tb");
    ASSERT_WSTR(s.Text(0, 1), L"say \"hi\"");
    ASSERT_TRUE(s.mgr().OnCanUndo() == undoBefore + 1);
}

TEST(Copy_TextIsTabbedAndQuoted)
{
    Sheet s;
    s.Set(0, 0, L"a");
    s.Set(0, 1, L"x\ty");
    s.Set(1, 1, L"z");
    ASSERT_WSTR(s.mgr().BuildCopyText({ { 0, 0 }, { 1, 1 } }), L"a\t\"x\ty\"\r\n\tz");
}

TEST(Copy_PrivateFormatKeepsFormulasAndFormats)
{
    // Copy carried only the displayed text (M10).
    Sheet s;
    s.Set(0, 0, L"=1+2");
    s.mgr().SetCellNumberFormat(0, 0, FMT_CURRENCY);
    const std::wstring records = s.mgr().BuildCopyRecords({ { 0, 0 }, { 0, 0 } });
    s.mgr().SetCurrentCell(4, 2);
    ASSERT_TRUE(s.mgr().PasteRecords(records));
    PGRIDCELL c = s.mgr().GetCell(4, 2);
    ASSERT_TRUE(c && c->m_bFormula);
    ASSERT_WSTR(c->m_wsFormula, L"1+2");
    ASSERT_TRUE(c->m_nFormat == FMT_CURRENCY);
    ASSERT_WSTR(s.Text(4, 2), L"$3.00");
}

// ---- AUDIT_2026-10 S13, AUDIT_2026-09 M9 / M20 / P7 ----------------------------
TEST(Undo_ReportsWhatChangedAndNothingWhenEmpty)
{
    Sheet s;
    GRIDSELECTION box{};
    ASSERT_TRUE(!s.mgr().OnUndo(&box));         // nothing to undo: no GN_CELLCHANGED
    s.Set(3, 4, L"x");
    ASSERT_TRUE(s.mgr().OnUndo(&box));
    ASSERT_TRUE(SameSel(box, 3, 4, 3, 4));
}

TEST(GetCellText_EmptyCellIsSuccess)
{
    Sheet s;
    wchar_t buf[8] = L"junk";
    GRID_GETTEXT gt{ buf, 8 };
    ASSERT_TRUE(s.mgr().OnGetCellText({ 2, 2 }, &gt) == TRUE);
    ASSERT_TRUE(s.mgr().GetLastError() == 0);
    ASSERT_WSTR(std::wstring(buf), L"");
}

TEST(ColumnLabels_AreBijectiveBase26)
{
    ASSERT_WSTR(CGrid32Mgr::GetColumnLabel(0), L"A");
    ASSERT_WSTR(CGrid32Mgr::GetColumnLabel(25), L"Z");
    ASSERT_WSTR(CGrid32Mgr::GetColumnLabel(26), L"AA");
    ASSERT_WSTR(CGrid32Mgr::GetColumnLabel(701), L"ZZ");
    ASSERT_WSTR(CGrid32Mgr::GetColumnLabel(702), L"AAA");
}

TEST(Geometry_SparseSizesAndHiddenRows)
{
    // Row and column metadata is stored only where it differs (P7).
    Sheet s(16, 1000000);
    s.mgr().SetColumnWidth(5, 250);
    ASSERT_TRUE(s.mgr().CalculatedColumnDistance(0, 10) == 9 * 100 + 250);
    s.mgr().SetRowHeight(10, 45);
    ASSERT_TRUE(s.mgr().CalculatedRowDistance(0, 1000000) == 999999ull * 20 + 45);
    s.mgr().SetColumnWidth(5, 100);             // back to the default: nothing stored
    ASSERT_TRUE(s.mgr().CalculatedColumnDistance(0, 16) == 1600);
}

// ---- runner ----------------------------------------------------------------
int main()
{
    std::cout << "Running " << Tests().size() << " CGrid32Mgr tests\n";
    for (auto& t : Tests())
    {
        g_currentTest = t.name;
        int before = g_failCount;
        t.fn();
        if (g_failCount == before)
            std::cout << "  ok   " << t.name << "\n";
    }
    if (g_failCount == 0) { std::cout << "\nAll tests passed.\n"; return 0; }
    std::cout << "\n" << g_failCount << " failure(s).\n";
    return 1;
}
