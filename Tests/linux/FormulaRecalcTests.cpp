// Formula engine tests against the real grid manager.
//
// FormulatorTests drives the parser against a stub sheet; this suite drives
// it the way the control does — SetCellText, SetCellNumberFormat,
// RecalculateFormulas, undo — through the real Grid32Mgr.cpp linked against
// the Win32 shim (see README.md). Linux-only, like GridMgrTests, and for the
// same reason: the shim defines Win32 symbols that would collide with the
// real ones.
//
// Covers AUDIT_2026-10 S3 (references read values, not display text), H8
// (number formats apply to formula cells), H21 (references bounded by the
// grid) and P4 (one evaluation per formula per recalculation). M23 is covered
// by this suite existing at all: it links Grid32Mgr.cpp and Formulator.cpp at
// -O2, which failed with an undefined Grid32FormulaText::ResetText.

#define _CRT_SECURE_NO_WARNINGS
#include <pch.h>
#include <grid32_internal.h>
#include <Formulator.h>
#include "win32shim/win32stubs.h"

#include <iostream>
#include <functional>
#include <cmath>
#include <string>
#include <chrono>

// ---- tiny test harness (same shape as the other suites) ---------------------
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
#define ASSERT_WSTR(expr, want) do { std::wstring _g = (expr); std::wstring _w = (want); \
    if (_g != _w) { std::wcerr << L"  (got \"" << _g << L"\", want \"" << _w << L"\")\n"; \
    FAIL_HERE("expected " #expr " == " #want); } } while (0)

// ---- fixture ---------------------------------------------------------------
// A grid with no real window behind it; see GridMgrTests.cpp. Derives from
// CGrid32Mgr so a test can hold recalculation off while it builds a sheet,
// then run exactly one pass.
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
    void Format(UINT row, UINT col, UINT fmt) { SetCellNumberFormat(row, col, fmt); }
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
    void HoldRecalc(bool hold) { m_bDeferRecalc = hold ? TRUE : FALSE; }
private:
    bool m_ok;
};

static double ElapsedSeconds(std::chrono::steady_clock::time_point since)
{
    return std::chrono::duration<double>(std::chrono::steady_clock::now() - since).count();
}

// ---- S3: references read the value, not the display text --------------------
TEST(Reference_ReadsTypedValueNotDisplayText)
{
    Sheet s;
    ASSERT_TRUE(s.ok());
    s.Set(0, 0, L"1234.5");
    s.Format(0, 0, FMT_CURRENCY);                 // A1 shows "$1,234.50"
    s.Set(1, 0, L"2026-05-15");                   // A2 is a date
    s.Set(2, 0, L"TRUE");                         // A3 is a boolean
    s.Set(3, 0, L"0.25");
    s.Format(3, 0, FMT_PERCENT);                  // A4 shows "25.00%"
    ASSERT_WSTR(s.Text(0, 0), L"$1,234.50");
    ASSERT_WSTR(s.Text(3, 0), L"25.00%");

    s.Set(0, 1, L"=A1+1");
    ASSERT_WSTR(s.Text(0, 1), L"1235.5");         // was "1": stod("$1,234.50") threw
    s.Set(1, 1, L"=YEAR(A2)");
    ASSERT_WSTR(s.Text(1, 1), L"2026");           // was "1905": the date read as 2026
    s.Set(2, 1, L"=A3*10");
    ASSERT_WSTR(s.Text(2, 1), L"10");             // was "0"
    s.Set(3, 1, L"=A4*100");
    ASSERT_WSTR(s.Text(3, 1), L"25");             // was "2500"
}

TEST(Reference_FollowsAFormatChangeWithoutChangingValue)
{
    // Formatting the source re-renders its text; what a formula reads from it
    // must not move.
    Sheet s;
    s.Set(0, 0, L"1000");
    s.Set(0, 1, L"=A1*2");
    ASSERT_WSTR(s.Text(0, 1), L"2000");
    s.Format(0, 0, FMT_THOUSANDS);                // A1 shows "1,000.00"
    s.Set(5, 5, L"x");                            // any edit: a full recalculation
    ASSERT_WSTR(s.Text(0, 1), L"2000");           // was "2": stod read "1"
}

// ---- H8: number formats apply to formula cells --------------------------------
TEST(FormulaCell_TakesItsNumberFormat)
{
    Sheet s;
    s.Set(0, 0, L"1234.5");
    s.Set(0, 1, L"=A1");
    ASSERT_WSTR(s.Text(0, 1), L"1234.5");
    s.Format(0, 1, FMT_CURRENCY);
    ASSERT_WSTR(s.Text(0, 1), L"$1,234.50");      // was "1234.5": formula cells were skipped
    ASSERT_NEAR(s.Value(0, 1), 1234.5, 1e-9);
}

TEST(FormulaCell_FormatSurvivesRecalculation)
{
    Sheet s;
    s.Set(0, 0, L"1");
    s.Set(0, 1, L"=A1/4");
    s.Format(0, 1, FMT_PERCENT);
    ASSERT_WSTR(s.Text(0, 1), L"25.00%");
    s.Set(0, 0, L"3");                            // the source moves: recalculated
    ASSERT_WSTR(s.Text(0, 1), L"75.00%");         // was "0.75"
    ASSERT_NEAR(s.Value(0, 1), 0.75, 1e-12);
}

TEST(FormulaCell_FormattedBeforeTheFormulaIsTyped)
{
    Sheet s;
    s.Format(2, 2, FMT_CURRENCY);
    s.Set(2, 2, L"=1000*2");
    ASSERT_WSTR(s.Text(2, 2), L"$2,000.00");
}

TEST(FormulaCell_EveryFormatAndBackToGeneral)
{
    Sheet s;
    s.Set(0, 0, L"=1234.5");
    s.Format(0, 0, FMT_NUMBER);    ASSERT_WSTR(s.Text(0, 0), L"1234.50");
    s.Format(0, 0, FMT_INTEGER);   ASSERT_WSTR(s.Text(0, 0), L"1235");
    s.Format(0, 0, FMT_THOUSANDS); ASSERT_WSTR(s.Text(0, 0), L"1,234.50");
    s.Format(0, 0, FMT_GENERAL);   ASSERT_WSTR(s.Text(0, 0), L"1234.5");

    s.Set(1, 0, L"=DATE(2026,5,15)");
    s.Format(1, 0, FMT_DATE_ISO);  ASSERT_WSTR(s.Text(1, 0), L"2026-05-15");
    s.Format(1, 0, FMT_DATE_US);   ASSERT_WSTR(s.Text(1, 0), L"5/15/2026");
}

TEST(FormulaCell_SelectionFormatReachesFormulas)
{
    Sheet s;
    s.Set(0, 0, L"=0.5");
    s.Set(1, 0, L"=0.25");
    GRIDSELECTION sel{ {0, 0}, {1, 0} };
    s.mgr().OnSetSelection(&sel);
    s.mgr().SetSelectionNumberFormat(FMT_PERCENT);
    ASSERT_WSTR(s.Text(0, 0), L"50.00%");
    ASSERT_WSTR(s.Text(1, 0), L"25.00%");
}

TEST(FormulaCell_UndoOfAFormatReRenders)
{
    Sheet s;
    s.Set(0, 0, L"=1234.5");
    s.Format(0, 0, FMT_CURRENCY);
    ASSERT_WSTR(s.Text(0, 0), L"$1,234.50");
    s.mgr().OnUndo();
    ASSERT_TRUE(s.mgr().GetCell(0, 0)->m_nFormat == FMT_GENERAL);
    ASSERT_WSTR(s.Text(0, 0), L"1234.5");
    s.mgr().OnRedo();
    ASSERT_WSTR(s.Text(0, 0), L"$1,234.50");
}

TEST(FormulaCell_TextResultIgnoresNumberFormat)
{
    Sheet s;
    s.Set(0, 0, L"=UPPER(\"abc\")");
    s.Format(0, 0, FMT_CURRENCY);
    ASSERT_WSTR(s.Text(0, 0), L"ABC");
}

TEST(FormulaCell_UnrenderableResultKeepsGeneralText)
{
    // FormatFixed's integer arithmetic cannot hold an infinity or a huge
    // magnitude; those keep their general text instead of turning to garbage.
    Sheet s;
    s.Set(0, 0, L"=1/0");
    s.Format(0, 0, FMT_CURRENCY);
    ASSERT_WSTR(s.Text(0, 0), CFormulator::FormatGeneral(1.0 / 0.0));
    s.Set(1, 0, L"=POWER(10,20)");
    s.Format(1, 0, FMT_PERCENT);
    ASSERT_WSTR(s.Text(1, 0), CFormulator::FormatGeneral(1e20));
}

TEST(FormulaCell_CachedValueIsTheNumberNotItsText)
{
    // m_dValue was the display text parsed back: six significant digits, and
    // 0 as soon as a format put a "$" in front.
    Sheet s;
    s.Set(0, 0, L"=1234567.89");
    ASSERT_NEAR(s.Value(0, 0), 1234567.89, 1e-6);  // was 1234570
    s.Format(0, 0, FMT_CURRENCY);
    ASSERT_WSTR(s.Text(0, 0), L"$1,234,567.89");
    ASSERT_NEAR(s.Value(0, 0), 1234567.89, 1e-6);
}

TEST(FormulaCell_ReadersOfAFormattedFormulaSeeItsValue)
{
    Sheet s;
    s.Set(0, 0, L"1234.5");
    s.Set(0, 1, L"=A1");
    s.Format(0, 1, FMT_CURRENCY);                 // B1 shows "$1,234.50"
    s.Set(0, 2, L"=B1*2");
    ASSERT_WSTR(s.Text(0, 2), L"2469");
    s.Set(0, 3, L"=CONCAT(B1)");
    ASSERT_WSTR(s.Text(0, 3), L"1234.5");         // the value, not "$1,234.50"
}

TEST(FormulaCell_BareReferenceToTextShowsTheText)
{
    Sheet s;
    s.Set(0, 0, L"=UPPER(\"abc\")");
    s.Set(1, 0, L"hello");
    s.Set(0, 1, L"=A1");
    ASSERT_WSTR(s.Text(0, 1), L"ABC");
    s.Set(1, 1, L"=A2");
    ASSERT_WSTR(s.Text(1, 1), L"hello");          // was "0"
    s.Set(2, 1, L"=A1+1");
    ASSERT_WSTR(s.Text(2, 1), L"1");              // was "ABC": the text leaked out
}

// ---- H21: references are bounded by the grid ------------------------------------
TEST(Reference_OutsideTheGridReadsAsZero)
{
    Sheet s(16, 64);                              // A..P, rows 1..64
    s.Set(63, 0, L"7");                           // A64, the last row
    s.Set(0, 15, L"3");                           // P1, the last column
    s.Set(0, 1, L"=A64");
    ASSERT_WSTR(s.Text(0, 1), L"7");
    s.Set(1, 1, L"=P1");
    ASSERT_WSTR(s.Text(1, 1), L"3");
    s.Set(2, 1, L"=A65");
    ASSERT_WSTR(s.Text(2, 1), L"0");
    s.Set(3, 1, L"=Q1");
    ASSERT_WSTR(s.Text(3, 1), L"0");
}

TEST(Range_OffTheGridIsRejectedPromptly)
{
    // =SUM(A1:A4294967296) spun forever on the UI thread: row 4294967296
    // parses as UINT_MAX, and a UINT counter can never pass it.
    Sheet s(16, 64);
    s.Set(0, 0, L"5");
    auto start = std::chrono::steady_clock::now();
    s.Set(0, 1, L"=SUM(A1:A4294967296)");
    ASSERT_WSTR(s.Text(0, 1), L"0");
    s.Set(1, 1, L"=SUM(A1:A65)");                 // one row past the end: rejected, not clamped
    ASSERT_WSTR(s.Text(1, 1), L"0");
    s.Set(2, 1, L"=SUM(A1:A64)");
    ASSERT_WSTR(s.Text(2, 1), L"5");
    ASSERT_TRUE(ElapsedSeconds(start) < 2.0);
}

// ---- P4: one evaluation per formula per pass ------------------------------------
// Builds a column of n cells in which every formula reads its neighbour twice.
static void BuildDoubledChain(Sheet& s, UINT n, bool upward)
{
    s.HoldRecalc(true);
    if (upward)
    {
        // A(n) = 1, A(k) = (A(k+1)+A(k+1))/2+1: A1 = n. Visited top-down, the
        // first cell's evaluation has to descend the whole chain.
        s.Set(n - 1, 0, L"1");
        for (UINT r = 0; r + 1 < n; ++r)
        {
            wchar_t buf[64];
            swprintf_s(buf, 64, L"=(A%u+A%u)/2+1", r + 2, r + 2);
            s.Set(r, 0, buf);
        }
    }
    else
    {
        // A1 = 1, A(k) = A(k-1)+A(k-1): A(k) = 2^(k-1).
        s.Set(0, 0, L"1");
        for (UINT r = 1; r < n; ++r)
        {
            wchar_t buf[64];
            swprintf_s(buf, 64, L"=A%u+A%u", r, r);
            s.Set(r, 0, buf);
        }
    }
    s.HoldRecalc(false);
}

TEST(Recalc_EvaluatesEachFormulaOnce)
{
    // Without the memo one pass evaluated A50 2^48 times over.
    Sheet s(4, 128);
    BuildDoubledChain(s, 50, false);
    size_t before = CFormulator::CellEvaluationCount();
    auto start = std::chrono::steady_clock::now();
    s.mgr().RecalculateFormulas();
    ASSERT_TRUE(CFormulator::CellEvaluationCount() - before == 49);
    ASSERT_NEAR(s.Value(49, 0), std::ldexp(1.0, 49), 1.0);
    ASSERT_TRUE(ElapsedSeconds(start) < 1.0);
}

TEST(Recalc_EvaluatesEachFormulaOnce_InAnyOrder)
{
    // Same, but the first formula visited depends on all the others.
    Sheet s(4, 128);
    BuildDoubledChain(s, 40, true);
    size_t before = CFormulator::CellEvaluationCount();
    s.mgr().RecalculateFormulas();
    ASSERT_TRUE(CFormulator::CellEvaluationCount() - before == 39);
    ASSERT_WSTR(s.Text(0, 0), L"40");
    ASSERT_WSTR(s.Text(38, 0), L"2");
}

TEST(Recalc_ChainPastTheDepthGuardResolvesExactly)
{
    // 3000 cells deep, visited from the top, each reading the next twice.
    // Before: exponential, and truncated where the depth guard tripped.
    const UINT n = 3000;
    Sheet s(4, 4096);
    BuildDoubledChain(s, n, true);
    size_t before = CFormulator::CellEvaluationCount();
    auto start = std::chrono::steady_clock::now();
    s.mgr().RecalculateFormulas();
    ASSERT_WSTR(s.Text(0, 0), L"3000");
    ASSERT_WSTR(s.Text(1500, 0), L"1500");
    ASSERT_TRUE(CFormulator::CellEvaluationCount() - before <= 3 * (size_t)n);
    ASSERT_TRUE(ElapsedSeconds(start) < 5.0);
}

TEST(Recalc_TypingIntoALongChainStaysFast)
{
    // Each keystroke recalculates the whole sheet; with the memo that is
    // linear in the sheet, so building and then editing the chain is cheap.
    Sheet s(4, 512);
    BuildDoubledChain(s, 400, false);
    s.mgr().RecalculateFormulas();
    auto start = std::chrono::steady_clock::now();
    s.Set(0, 0, L"0.5");
    ASSERT_NEAR(s.Value(399, 0), std::ldexp(1.0, 398), std::ldexp(1.0, 360));
    ASSERT_TRUE(ElapsedSeconds(start) < 2.0);
}

TEST(Recalc_CyclesStillTerminateAndAgree)
{
    // A self-reference reads as 0 whether the cell is refreshed on its own
    // (by pointer, as typing does) or as part of a pass (by position).
    Sheet s;
    s.Set(0, 0, L"=A1+1");
    ASSERT_WSTR(s.Text(0, 0), L"1");
    s.Set(5, 5, L"x");                            // another full pass
    ASSERT_WSTR(s.Text(0, 0), L"1");

    s.Set(1, 0, L"=B2+1");                        // A2 <-> B2
    s.Set(1, 1, L"=A2+1");
    ASSERT_TRUE(std::isfinite(s.Value(1, 0)));
    ASSERT_TRUE(std::isfinite(s.Value(1, 1)));
}

// ---- runner ----------------------------------------------------------------
int main()
{
    std::cout << "Running " << Tests().size() << " formula recalculation tests\n";
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
