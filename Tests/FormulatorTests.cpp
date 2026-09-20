// Formulator test harness — standalone console exe, separate from
// Grid32Tests.exe. Build: see Tests/run_formulator_tests.bat.
//
// Formulator.cpp's only dependency on the grid is CGrid32Mgr::GetCell, so this
// exe links the parser against the stub definitions at the bottom of this file
// instead of Grid32Mgr.cpp — which would drag in GDI+, the window procedure
// and the edit subclass. Nothing here links Grid32Mgr.cpp, so the stubs are
// the only definitions in the image and there is no ODR conflict.
//
// If Formulator.cpp ever calls a second CGrid32Mgr member, this exe stops
// linking until a stub for it is added below. That is deliberate: it keeps the
// parser's coupling to the manager visible and minimal.

#define _CRT_SECURE_NO_WARNINGS
#include "../Grid32/pch.h"
#include "../Grid32/grid32_internal.h"
#include "../Grid32/Formulator.h"

#include <iostream>
#include <functional>
#include <cmath>
#include <cstdio>
#include <cwchar>

// ---- tiny test harness (same shape as Grid32Tests.cpp) ---------------------
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
#define ASSERT_FALSE(cond) do { if ((cond)) FAIL_HERE("expected false: " #cond); } while (0)
#define ASSERT_NEAR(a, b, eps) do { double _da = (a), _db = (b); if (std::fabs(_da - _db) > (eps)) { \
    std::cerr << "  (got " << _da << ", want " << _db << ")\n"; FAIL_HERE("expected " #a " near " #b); } } while (0)

// ---- stub sheet ------------------------------------------------------------
// mapCells is protected, so a derived class is how the test populates it.
class TestSheet : public CGrid32Mgr
{
public:
    std::map<std::pair<UINT, UINT>, PGRIDCELL>& Cells() { return mapCells; }
};

// A CGrid32Mgr whose cell map is populated directly. Text beginning with '='
// becomes a formula cell, mirroring SetCellText's rule.
static TestSheet* g_sheet = nullptr;
static std::map<std::pair<UINT, UINT>, PGRIDCELL>* g_cells = nullptr;

static void SheetReset()
{
    if (g_cells) { for (auto& kv : *g_cells) delete kv.second; g_cells->clear(); }
}

static void SheetSet(UINT row, UINT col, const wchar_t* text)
{
    PGRIDCELL cell = new GRIDCELL();
    if (text[0] == L'=') { cell->m_bFormula = true; cell->m_wsFormula = text + 1; cell->m_eType = CT_Formula; }
    else                 { cell->m_wsText = text; }
    auto key = std::make_pair(row, col);
    auto it = g_cells->find(key);
    if (it != g_cells->end()) { delete it->second; it->second = cell; }
    else                      { (*g_cells)[key] = cell; }
}

// Evaluate a formula against the stub sheet, the way EvaluateFormula does.
static double Eval(const wchar_t* formula)
{
    std::wstring e = formula;
    if (!e.empty() && e[0] == L'=') e = e.substr(1);
    std::set<std::pair<UINT, UINT>> visited;
    return CFormulator::EvalArg(g_sheet, e, visited);
}

// A1=5 A2=7 A3=9 B1=2 B2=4
static void StandardSheet()
{
    SheetReset();
    SheetSet(0, 0, L"5"); SheetSet(1, 0, L"7"); SheetSet(2, 0, L"9");
    SheetSet(0, 1, L"2"); SheetSet(1, 1, L"4");
}

// ---- ParseCellRef / ParseRange (B1) ---------------------------------------
TEST(ParseCellRef_AcceptsWholeReference)
{
    UINT r = 0, c = 0;
    ASSERT_TRUE(CFormulator::ParseCellRef(L"A1", r, c));   ASSERT_TRUE(r == 0 && c == 0);
    ASSERT_TRUE(CFormulator::ParseCellRef(L"B3", r, c));   ASSERT_TRUE(r == 2 && c == 1);
    ASSERT_TRUE(CFormulator::ParseCellRef(L"ZZ100", r, c));ASSERT_TRUE(r == 99 && c == 701);
}

TEST(ParseCellRef_RejectsTrailingJunk)
{
    // The bug that made every range collapse to its first cell: std::stoul
    // stops at ':' without throwing, so "A1:B10" parsed as a valid ref to A1.
    UINT r = 0, c = 0;
    ASSERT_FALSE(CFormulator::ParseCellRef(L"A1:B10", r, c));
    ASSERT_FALSE(CFormulator::ParseCellRef(L"A1B", r, c));
    ASSERT_FALSE(CFormulator::ParseCellRef(L"A 1", r, c));
    ASSERT_FALSE(CFormulator::ParseCellRef(L"A0", r, c));   // rows are 1-based
    ASSERT_FALSE(CFormulator::ParseCellRef(L"SUM", r, c));
    ASSERT_FALSE(CFormulator::ParseCellRef(L"123", r, c));
    ASSERT_FALSE(CFormulator::ParseCellRef(L"", r, c));
}

TEST(ParseRange_SingleAndSpan)
{
    UINT sr = 9, sc = 9, er = 9, ec = 9;
    ASSERT_TRUE(CFormulator::ParseRange(L"A1", sr, sc, er, ec));
    ASSERT_TRUE(sr == 0 && sc == 0 && er == 0 && ec == 0);

    ASSERT_TRUE(CFormulator::ParseRange(L"A1:B3", sr, sc, er, ec));
    ASSERT_TRUE(sr == 0 && sc == 0 && er == 2 && ec == 1);

    // Reversed and whitespace-padded ranges normalize.
    ASSERT_TRUE(CFormulator::ParseRange(L"B3:A1", sr, sc, er, ec));
    ASSERT_TRUE(sr == 0 && sc == 0 && er == 2 && ec == 1);
    ASSERT_TRUE(CFormulator::ParseRange(L" A1 : B3 ", sr, sc, er, ec));
    ASSERT_TRUE(sr == 0 && sc == 0 && er == 2 && ec == 1);

    ASSERT_FALSE(CFormulator::ParseRange(L"A1:", sr, sc, er, ec));
    ASSERT_FALSE(CFormulator::ParseRange(L"zzz", sr, sc, er, ec));
}

// ---- Range aggregates (B1) -------------------------------------------------
TEST(Aggregates_SpanTheWholeRange)
{
    StandardSheet();
    ASSERT_NEAR(Eval(L"=SUM(A1:A3)"), 21.0, 1e-9);
    ASSERT_NEAR(Eval(L"=SUM(A1:B2)"), 18.0, 1e-9);   // 5+2+7+4
    ASSERT_NEAR(Eval(L"=SUM(A3:A1)"), 21.0, 1e-9);   // reversed
    ASSERT_NEAR(Eval(L"=SUM( A1 : A3 )"), 21.0, 1e-9);
    ASSERT_NEAR(Eval(L"=SUM(A2)"), 7.0, 1e-9);       // single cell
    ASSERT_NEAR(Eval(L"=SUM(A1:A3)+1"), 22.0, 1e-9);
    ASSERT_NEAR(Eval(L"=MIN(A1:A3)"), 5.0, 1e-9);
    ASSERT_NEAR(Eval(L"=MAX(A1:A3)"), 9.0, 1e-9);
    ASSERT_NEAR(Eval(L"=COUNT(A1:A10)"), 3.0, 1e-9);
    ASSERT_NEAR(Eval(L"=AVERAGE(A1:A3)"), 7.0, 1e-9);
}

TEST(Aggregates_SkipBlanks)
{
    // Ranges now really do span their whole rectangle, so blank cells must not
    // drag MIN to 0 or inflate AVERAGE's denominator.
    StandardSheet();
    ASSERT_NEAR(Eval(L"=MIN(A1:A10)"), 5.0, 1e-9);
    ASSERT_NEAR(Eval(L"=AVERAGE(A1:A10)"), 7.0, 1e-9);
    ASSERT_NEAR(Eval(L"=SUM(A1:A10)"), 21.0, 1e-9);  // blanks add 0, harmless
    // An all-blank range has no min/max; report 0, not an infinity.
    ASSERT_NEAR(Eval(L"=MAX(D1:D9)"), 0.0, 1e-9);
    ASSERT_NEAR(Eval(L"=MIN(D1:D9)"), 0.0, 1e-9);
}

// ---- Repeated references and cycles (B3) -----------------------------------
TEST(RepeatedReference_EvaluatesEachTime)
{
    // The visited set used to accumulate every cell touched, so the second
    // mention of A1 resolved to 0.
    StandardSheet();
    ASSERT_NEAR(Eval(L"=A1+A1"), 10.0, 1e-9);
    ASSERT_NEAR(Eval(L"=A1+A1+A1"), 15.0, 1e-9);
    ASSERT_NEAR(Eval(L"=A1*A1"), 25.0, 1e-9);
    ASSERT_NEAR(Eval(L"=SUM(A1:A3)+A1"), 26.0, 1e-9);
}

TEST(SharedDependency_ReachesBothLegs)
{
    SheetReset();
    SheetSet(0, 0, L"10");        // A1
    SheetSet(1, 0, L"=A1*2");     // A2 = 20
    SheetSet(2, 0, L"=A1*3");     // A3 = 30
    ASSERT_NEAR(Eval(L"=A2+A3"), 50.0, 1e-9);
}

TEST(Cycles_StillTerminate)
{
    // Direct self-reference: A1 = A1+1 -> the inner A1 is on the path, yields 0.
    SheetReset();
    SheetSet(0, 0, L"=A1+1");
    ASSERT_NEAR(Eval(L"=A1"), 1.0, 1e-9);

    // Mutual: A1 = B1+1, B1 = A1+1. B1 sees A1 on the path (0), so B1 = 1
    // and A1 = 2. The point is that it terminates with a finite value.
    SheetReset();
    SheetSet(0, 0, L"=B1+1");
    SheetSet(0, 1, L"=A1+1");
    ASSERT_NEAR(Eval(L"=A1"), 2.0, 1e-9);
}

TEST(DeepChain_DoesNotOverflowStack)
{
    // 2000-cell dependency chain; the depth guard caps recursion well before
    // the stack runs out.
    SheetReset();
    SheetSet(0, 0, L"1");
    for (UINT i = 1; i < 2000; ++i)
    {
        wchar_t buf[32];
        swprintf_s(buf, 32, L"=A%u+1", i);
        SheetSet(i, 0, buf);
    }
    double v = Eval(L"=A2000");
    ASSERT_TRUE(std::isfinite(v));
}

TEST(LongUnaryChain_DoesNotOverflowStack)
{
    // Unary +/- recurses straight back into ParseFactor. Without a guard there
    // this segfaults.
    StandardSheet();
    std::wstring expr(100000, L'-');
    expr += L"1";
    std::set<std::pair<UINT, UINT>> visited;
    double v = CFormulator::EvalArg(g_sheet, expr, visited);
    ASSERT_TRUE(std::isfinite(v));
}

// ---- Comparison operators (H6) ---------------------------------------------
TEST(Comparisons_ProduceOneOrZero)
{
    StandardSheet();
    ASSERT_NEAR(Eval(L"=A1>3"), 1.0, 1e-9);
    ASSERT_NEAR(Eval(L"=A1<3"), 0.0, 1e-9);
    ASSERT_NEAR(Eval(L"=A1>=5"), 1.0, 1e-9);
    ASSERT_NEAR(Eval(L"=A1<=4"), 0.0, 1e-9);
    ASSERT_NEAR(Eval(L"=A1=5"), 1.0, 1e-9);
    ASSERT_NEAR(Eval(L"=A1<>5"), 0.0, 1e-9);
    ASSERT_NEAR(Eval(L"=A1<>4"), 1.0, 1e-9);
}

TEST(Comparisons_BindLooserThanArithmetic)
{
    StandardSheet();
    ASSERT_NEAR(Eval(L"=A1+1>B1*2"), 1.0, 1e-9);  // 6 > 4
    ASSERT_NEAR(Eval(L"=1+1=2"), 1.0, 1e-9);
    ASSERT_NEAR(Eval(L"=2*3<5"), 0.0, 1e-9);
}

TEST(Comparisons_UseTolerance)
{
    // Binary rounding must not make =0.1+0.2=0.3 false.
    StandardSheet();
    ASSERT_NEAR(Eval(L"=0.1+0.2=0.3"), 1.0, 1e-9);
}

TEST(Conditionals_ReadComparisons)
{
    StandardSheet();
    ASSERT_NEAR(Eval(L"=IF(A1>3,100,200)"), 100.0, 1e-9);
    ASSERT_NEAR(Eval(L"=IF(A1<3,100,200)"), 200.0, 1e-9);
    ASSERT_NEAR(Eval(L"=IF(A1>A2,1,0)"), 0.0, 1e-9);
    ASSERT_NEAR(Eval(L"=IFS(A1>10,1,A1>3,2,1,3)"), 2.0, 1e-9);
    ASSERT_NEAR(Eval(L"=AND(A1>1,A2>1)"), 1.0, 1e-9);
    ASSERT_NEAR(Eval(L"=AND(A1>1,A2>99)"), 0.0, 1e-9);
    ASSERT_NEAR(Eval(L"=OR(A1>99,A2>1)"), 1.0, 1e-9);
    ASSERT_NEAR(Eval(L"=NOT(A1>99)"), 1.0, 1e-9);
}

// ---- Arithmetic must be unchanged ------------------------------------------
TEST(Arithmetic_Unchanged)
{
    StandardSheet();
    ASSERT_NEAR(Eval(L"=2+3*4"), 14.0, 1e-9);
    ASSERT_NEAR(Eval(L"=(2+3)*4"), 20.0, 1e-9);
    ASSERT_NEAR(Eval(L"=10/4"), 2.5, 1e-9);
    ASSERT_NEAR(Eval(L"=-A1"), -5.0, 1e-9);
    ASSERT_NEAR(Eval(L"=--A1"), 5.0, 1e-9);
    ASSERT_NEAR(Eval(L"=2*-3"), -6.0, 1e-9);
    ASSERT_NEAR(Eval(L"=ROUND(3.14159,2)"), 3.14, 1e-9);
    ASSERT_NEAR(Eval(L"=ABS(0-7)"), 7.0, 1e-9);
    ASSERT_NEAR(Eval(L"=SQRT(16)"), 4.0, 1e-9);
    ASSERT_NEAR(Eval(L"=POWER(2,10)"), 1024.0, 1e-9);
}

// ---- runner ----------------------------------------------------------------
int main()
{
    TestSheet sheet;
    g_sheet = &sheet;
    g_cells = &sheet.Cells();

    std::cout << "Running " << Tests().size() << " Formulator tests\n";
    for (auto& t : Tests())
    {
        g_currentTest = t.name;
        int before = g_failCount;
        t.fn();
        if (g_failCount == before)
            std::cout << "  ok   " << t.name << "\n";
    }
    SheetReset();

    if (g_failCount == 0) { std::cout << "\nAll tests passed.\n"; return 0; }
    std::cout << "\n" << g_failCount << " failure(s).\n";
    return 1;
}

// ---- link seam -------------------------------------------------------------
// The only CGrid32Mgr members Formulator.cpp needs at link time. Grid32Mgr.cpp
// is deliberately not part of this exe, so these are the sole definitions in
// the image. GetCell reads nothing but mapCells, so the stub constructor does
// not need to reproduce the real one's initialization.
CGrid32Mgr::CGrid32Mgr()
    : pRowInfoArray(nullptr), pColInfoArray(nullptr),
      m_hWndGrid(nullptr), m_hWndEdit(nullptr),
      nColHeaderHeight(0), nRowHeaderWidth(0),
      m_bRedraw(TRUE), m_bResizable(FALSE), m_bUndoRecordEnabled(FALSE),
      m_bDeferRecalc(FALSE), dwError(0),
      m_bSelecting(FALSE), m_bSizing(FALSE), m_editWndProc(nullptr),
      m_nMouseHoverDelay(0), m_npHoverDelaySet(0), m_lastClickTime(0),
      m_gridHitTest(0), m_nToBeSized(0), m_nSizingLine(0), m_rgbSizingLine(0),
      m_hDefaultFont(nullptr), m_gdiplusToken(0)
{
    memset(&gcs, 0, sizeof(gcs));
    memset(&m_selectionRect, 0, sizeof(m_selectionRect));
    m_currentCell = m_visibleTopLeft = m_HoverCell = { 0, 0 };
    m_scrollDifference = { 0, 0 };
    m_mouseDraggingStartPoint = { 0, 0 };
    totalGridCellRect = m_clientRect = { 0, 0, 0, 0 };
    visibleGrid = { 0, 0 };
}

CGrid32Mgr::~CGrid32Mgr() {}

PGRIDCELL CGrid32Mgr::GetCell(UINT nRow, UINT nCol)
{
    auto it = mapCells.find(std::make_pair(nRow, nCol));
    return (it != mapCells.end()) ? it->second : nullptr;
}
