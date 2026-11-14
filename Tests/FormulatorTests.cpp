// Formulator test harness — standalone console exe, separate from
// Grid32Tests.exe. Build: see Tests/run_formulator_tests.bat.
//
// Formulator.cpp's only out-of-line dependency on the grid is
// CGrid32Mgr::GetCell (it also reads the grid's extent and cell map, through
// accessors defined inline in Grid32Mgr.h), so this exe links the parser
// against the stub definitions at the bottom of this file instead of
// Grid32Mgr.cpp — which would drag in GDI+, the window procedure and the edit
// subclass. Nothing here links Grid32Mgr.cpp, so the stubs are the only
// definitions in the image and there is no ODR conflict.
//
// If Formulator.cpp ever calls a second out-of-line CGrid32Mgr member, this
// exe stops linking until a stub for it is added below. That is deliberate: it
// keeps the parser's coupling to the manager visible and minimal.

#define _CRT_SECURE_NO_WARNINGS
#include "../Grid32/pch.h"
#include "../Grid32/grid32_internal.h"
#include "../Grid32/Formulator.h"

#include <iostream>
#include <functional>
#include <cmath>
#include <cstdio>
#include <cwchar>
#include <chrono>
#include <clocale>
#include <string>

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
// Fails on NaN too: "> eps" is false for NaN, so a NaN (an error result)
// used to pass every ASSERT_NEAR.
#define ASSERT_NEAR(a, b, eps) do { double _da = (a), _db = (b); if (!(std::fabs(_da - _db) <= (eps))) { \
    std::cerr << "  (got " << _da << ", want " << _db << ")\n"; FAIL_HERE("expected " #a " near " #b); } } while (0)

// ---- stub sheet ------------------------------------------------------------
// mapCells and gcs are protected, so a derived class is how the test populates
// them. The grid is the host's size: 702 columns (A..ZZ) by 4096 rows.
static const UINT kSheetCols = 702;
static const UINT kSheetRows = 4096;

class TestSheet : public CGrid32Mgr
{
public:
    TestSheet() { gcs.nWidth = kSheetCols; gcs.nHeight = kSheetRows; }
    std::map<std::pair<UINT, UINT>, PGRIDCELL>& Cells() { return mapCells; }
};

// A CGrid32Mgr whose cell map is populated directly. Text beginning with '='
// becomes a formula cell; anything else is typed and rendered the way
// SetCellText does it.
static TestSheet* g_sheet = nullptr;
static std::map<std::pair<UINT, UINT>, PGRIDCELL>* g_cells = nullptr;

static void SheetReset()
{
    if (g_cells) { for (auto& kv : *g_cells) delete kv.second; g_cells->clear(); }
}

static void SheetPut(UINT row, UINT col, PGRIDCELL cell)
{
    auto key = std::make_pair(row, col);
    auto it = g_cells->find(key);
    if (it != g_cells->end()) { delete it->second; it->second = cell; }
    else                      { (*g_cells)[key] = cell; }
}

// As typed into the grid, under number format `format`: the cell keeps its
// inferred type and value, and shows the formatted rendering.
static void SheetSetFormatted(UINT row, UINT col, const wchar_t* text, UINT format)
{
    PGRIDCELL cell = new GRIDCELL();
    cell->m_nFormat = format;
    if (text[0] == L'=')
    {
        cell->m_bFormula = true;
        cell->m_wsFormula = text + 1;
        cell->m_eType = CT_Formula;
    }
    else
    {
        double v = 0.0;
        cell->m_eType = Grid32Detail::InferType(text, v);
        cell->m_dValue = v;
        cell->m_wsText = Grid32Detail::FormatCellValue(cell->m_eType, v, text, format);
    }
    SheetPut(row, col, cell);
}

static void SheetSet(UINT row, UINT col, const wchar_t* text)
{
    SheetSetFormatted(row, col, text, FMT_GENERAL);
}

// A CT_Text cell holding `text` whatever it looks like — what GM_SETCELL or a
// fill stores, since neither infers a type.
static void SheetSetText(UINT row, UINT col, const wchar_t* text)
{
    PGRIDCELL cell = new GRIDCELL();
    cell->m_wsText = text;
    SheetPut(row, col, cell);
}

static PGRIDCELL SheetCell(UINT row, UINT col)
{
    auto it = g_cells->find(std::make_pair(row, col));
    return it != g_cells->end() ? it->second : nullptr;
}

static double ElapsedSeconds(std::chrono::steady_clock::time_point since)
{
    return std::chrono::duration<double>(std::chrono::steady_clock::now() - since).count();
}

// Evaluate a formula against the stub sheet, the way EvaluateFormula does.
static double Eval(const wchar_t* formula)
{
    std::wstring e = formula;
    if (!e.empty() && e[0] == L'=') e = e.substr(1);
    std::set<std::pair<UINT, UINT>> visited;
    return CFormulator::EvalArg(g_sheet, e, visited);
}

// A formula's result is the text `want` (not an error).
static bool IsText(const wchar_t* formula, const wchar_t* want)
{
    CFormulator::Result r = CFormulator::Evaluate(g_sheet, formula);
    return r.isText && !r.isError && r.text == want;
}

// A formula's result is the error `code`, with a NaN value.
static bool IsError(const wchar_t* formula, const wchar_t* code)
{
    CFormulator::Result r = CFormulator::Evaluate(g_sheet, formula);
    return r.isError && !r.isText && r.text == code && std::isnan(r.value);
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
    // this segfaults. Past the guard the expression cannot be evaluated, and
    // says so: it used to come out as whatever had been read when the guard
    // tripped.
    StandardSheet();
    std::wstring expr(100000, L'-');
    expr += L"1";
    CFormulator::Result r = CFormulator::Evaluate(g_sheet, expr);
    ASSERT_TRUE(r.isError && r.text == CFormulator::kErrorSyntax);
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

// ---- Typed cells are read by value, not display text (S3) ------------------
TEST(TypedCells_ReadTheirValueNotTheirDisplay)
{
    SheetReset();
    SheetSetFormatted(0, 0, L"1234.5", FMT_CURRENCY);    // shows "$1,234.50"
    SheetSetFormatted(1, 0, L"0.25", FMT_PERCENT);       // shows "25.00%"
    SheetSetFormatted(2, 0, L"1234.5", FMT_THOUSANDS);   // shows "1,234.50"
    SheetSet(3, 0, L"2024-01-15");                       // CT_Date
    SheetSet(4, 0, L"TRUE");                             // CT_Boolean
    SheetSet(5, 0, L"FALSE");
    ASSERT_TRUE(SheetCell(0, 0)->m_wsText == L"$1,234.50");
    ASSERT_TRUE(SheetCell(1, 0)->m_wsText == L"25.00%");

    // std::stod on the display text gave 0, 25, 1, 2024 and 0.
    ASSERT_NEAR(Eval(L"=A1"), 1234.5, 1e-9);
    ASSERT_NEAR(Eval(L"=A2"), 0.25, 1e-12);
    ASSERT_NEAR(Eval(L"=A3"), 1234.5, 1e-9);
    ASSERT_NEAR(Eval(L"=A4"), Grid32Detail::DateToSerial(2024, 1, 15), 1e-9);
    ASSERT_NEAR(Eval(L"=YEAR(A4)"), 2024.0, 1e-9);
    ASSERT_NEAR(Eval(L"=MONTH(A4)"), 1.0, 1e-9);
    ASSERT_NEAR(Eval(L"=A4+1-A4"), 1.0, 1e-9);
    ASSERT_NEAR(Eval(L"=A5+1"), 2.0, 1e-9);
    ASSERT_NEAR(Eval(L"=IF(A5,10,20)"), 10.0, 1e-9);
    ASSERT_NEAR(Eval(L"=IF(A6,10,20)"), 20.0, 1e-9);
}

TEST(TypedCells_AggregatesUseValues)
{
    SheetReset();
    SheetSetFormatted(0, 0, L"1000", FMT_CURRENCY);     // "$1,000.00"
    SheetSetFormatted(1, 0, L"2000", FMT_THOUSANDS);    // "2,000.00"
    SheetSet(2, 0, L"3000");
    ASSERT_NEAR(Eval(L"=SUM(A1:A3)"), 6000.0, 1e-9);    // was 0 + 2 + 3000
    ASSERT_NEAR(Eval(L"=AVERAGE(A1:A3)"), 2000.0, 1e-9);
    ASSERT_NEAR(Eval(L"=MIN(A1:A3)"), 1000.0, 1e-9);
    ASSERT_NEAR(Eval(L"=MAX(A1:A3)"), 3000.0, 1e-9);
    ASSERT_NEAR(Eval(L"=COUNT(A1:A3)"), 3.0, 1e-9);     // was 2: "$1,000.00" didn't parse
}

TEST(TextCells_NumericLookingTextReadsAsANumber)
{
    // CT_Text cells: typed with separators or symbols InferType doesn't take,
    // or stored by a path that doesn't infer types at all (GM_SETCELL, fill).
    SheetReset();
    SheetSetText(0, 0, L"1,234.50");
    SheetSetText(1, 0, L"$20");
    SheetSetText(2, 0, L"25%");
    SheetSetText(3, 0, L" 42 ");
    SheetSetText(4, 0, L"12abc");     // std::stod read the prefix: 12
    SheetSetText(5, 0, L"1,2");       // malformed grouping stays text
    SheetSetText(6, 0, L"-$5");
    ASSERT_NEAR(Eval(L"=A1+0"), 1234.5, 1e-9);   // was 1
    ASSERT_NEAR(Eval(L"=A2+0"), 20.0, 1e-9);     // was 0
    ASSERT_NEAR(Eval(L"=A3+0"), 0.25, 1e-12);    // was 25
    ASSERT_NEAR(Eval(L"=A4+0"), 42.0, 1e-9);
    // Text that is not a number is #VALUE! in arithmetic, as in Excel (it
    // read as 0 until the #VALUE! rule; before that std::stod read "12abc"
    // as 12).
    ASSERT_TRUE(IsError(L"A5+0", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"A6+0", CFormulator::kErrorValue));
    ASSERT_NEAR(Eval(L"=A7+0"), -5.0, 1e-9);
    ASSERT_NEAR(Eval(L"=COUNT(A1:A7)"), 5.0, 1e-9);
}

TEST(Aggregates_SkipNonNumericTextConsistently)
{
    // SUM, AVERAGE, MIN, MAX and COUNT agree on which cells are numbers. Text
    // that doesn't read as one is skipped like a blank (E6), not averaged in
    // as 0.
    SheetReset();
    SheetSet(0, 0, L"10");
    SheetSet(1, 0, L"hello");
    SheetSet(2, 0, L"20");
    ASSERT_NEAR(Eval(L"=SUM(A1:A3)"), 30.0, 1e-9);
    ASSERT_NEAR(Eval(L"=AVERAGE(A1:A3)"), 15.0, 1e-9);  // was 30/3
    ASSERT_NEAR(Eval(L"=MIN(A1:A3)"), 10.0, 1e-9);      // was 0
    ASSERT_NEAR(Eval(L"=MAX(A1:A3)"), 20.0, 1e-9);
    ASSERT_NEAR(Eval(L"=COUNT(A1:A3)"), 2.0, 1e-9);
}

TEST(ClearedCell_IsBlankWhateverItsStaleType)
{
    // ClearCellText empties the text but leaves m_eType and m_dValue behind.
    SheetReset();
    SheetSet(0, 0, L"5");
    SheetSet(1, 0, L"7");
    SheetCell(1, 0)->m_wsText.clear();          // still CT_Number, m_dValue 7
    ASSERT_NEAR(Eval(L"=A2+0"), 0.0, 1e-9);
    ASSERT_NEAR(Eval(L"=SUM(A1:A2)"), 5.0, 1e-9);
    ASSERT_NEAR(Eval(L"=AVERAGE(A1:A2)"), 5.0, 1e-9);
    ASSERT_NEAR(Eval(L"=COUNT(A1:A2)"), 1.0, 1e-9);
}

TEST(TextFunctions_SeeTheValueNotTheRendering)
{
    SheetReset();
    SheetSetFormatted(0, 0, L"1234.5", FMT_CURRENCY);
    ASSERT_NEAR(Eval(L"=LEN(A1)"), 6.0, 1e-9);      // "1234.5", not "$1,234.50"
    SheetSet(1, 0, L"hello");
    ASSERT_NEAR(Eval(L"=LEN(A2)"), 5.0, 1e-9);
}

TEST(TextResult_DoesNotLeakIntoArithmetic)
{
    // A referenced cell's text result used to leak out through the shared
    // slot as the referencing formula's own result: =A1+1 showed "42".
    SheetReset();
    SheetSet(0, 0, L"=LEFT(\"42abc\",2)");     // A1 = "42"
    SheetSet(0, 1, L"=UPPER(\"x\")");          // B1 = "X"
    CFormulator::Result r = CFormulator::Evaluate(g_sheet, L"A1+1");
    ASSERT_FALSE(r.isText);
    ASSERT_NEAR(r.value, 43.0, 1e-9);           // "42" reads as 42
    r = CFormulator::Evaluate(g_sheet, L"B1+1");
    ASSERT_FALSE(r.isText);
    ASSERT_TRUE(r.isError && r.text == CFormulator::kErrorValue);   // "X" is no number
}

TEST(BareReference_ShowsTheCellAsItIs)
{
    SheetReset();
    SheetSet(0, 0, L"=UPPER(\"abc\")");
    SheetSet(1, 0, L"hello");
    SheetSet(2, 0, L"7");
    CFormulator::Result r = CFormulator::Evaluate(g_sheet, L"A1");
    ASSERT_TRUE(r.isText && r.text == L"ABC");
    r = CFormulator::Evaluate(g_sheet, L"A2");
    ASSERT_TRUE(r.isText && r.text == L"hello");   // was 0
    r = CFormulator::Evaluate(g_sheet, L"A3");
    ASSERT_TRUE(!r.isText && r.value == 7.0);
    r = CFormulator::Evaluate(g_sheet, L"A9");     // blank
    ASSERT_TRUE(!r.isText && r.value == 0.0);
}

TEST(SelfReference_SameWhetherByPointerOrPosition)
{
    // RefreshFormulaCell hands the cell over by pointer, RecalculateFormulas
    // by position; a reference back to the cell must be the same cycle both
    // ways, or the display would change depending on which ran last.
    SheetReset();
    SheetSet(0, 0, L"=A1+1");
    CFormulator::Result byPointer = CFormulator::Evaluate(g_sheet, L"A1+1", SheetCell(0, 0));
    CFormulator::Result byPosition = CFormulator::EvaluateCell(g_sheet, 0, 0);
    ASSERT_NEAR(byPointer.value, 1.0, 1e-9);
    ASSERT_NEAR(byPosition.value, 1.0, 1e-9);
}

// ---- References are bounded by the grid (H21) -------------------------------
TEST(References_OutsideTheGridDoNotResolve)
{
    // A reference outside the grid reads as 0, like a malformed one ("A0").
    SheetReset();
    SheetSet(0, 0, L"5");
    SheetSet(kSheetRows - 1, 0, L"7");          // A4096, the last row
    SheetSet(0, kSheetCols - 1, L"3");          // ZZ1, the last column
    SheetSet(kSheetRows, 0, L"100");            // A4097: off the grid, reachable only by writing the map
    SheetSet(0, kSheetCols, L"100");            // AAA1
    ASSERT_NEAR(Eval(L"=A4096"), 7.0, 1e-9);
    ASSERT_NEAR(Eval(L"=ZZ1"), 3.0, 1e-9);
    ASSERT_NEAR(Eval(L"=A4097"), 0.0, 1e-9);    // was 100
    ASSERT_NEAR(Eval(L"=AAA1"), 0.0, 1e-9);     // was 100
    ASSERT_NEAR(Eval(L"=A4294967296"), 0.0, 1e-9);
    ASSERT_NEAR(Eval(L"=LEN(A4097)"), 1.0, 1e-9);   // reads as "0"
}

TEST(Ranges_OutsideTheGridAreRejectedNotWalked)
{
    SheetReset();
    SheetSet(0, 0, L"5");
    SheetSet(kSheetRows, 0, L"100");            // A4097, off the grid
    ASSERT_NEAR(Eval(L"=SUM(A1:A4096)"), 5.0, 1e-9);
    ASSERT_NEAR(Eval(L"=SUM(A1:A4097)"), 0.0, 1e-9);    // was 105
    ASSERT_NEAR(Eval(L"=SUM(A4097:A1)"), 0.0, 1e-9);
    ASSERT_NEAR(Eval(L"=SUM(A1:AAA1)"), 0.0, 1e-9);

    // Row 4294967296 parses as UINT_MAX. Walking to it never ended: the UINT
    // loop counter wraps before it can pass UINT_MAX, so this spun forever.
    auto start = std::chrono::steady_clock::now();
    ASSERT_NEAR(Eval(L"=SUM(A1:A4294967296)"), 0.0, 1e-9);
    ASSERT_TRUE(IsError(L"AVERAGE(A1:A4294967296)", CFormulator::kErrorDiv0));   // an average of nothing
    ASSERT_NEAR(Eval(L"=MAX(A4294967296:A1)"), 0.0, 1e-9);
    ASSERT_NEAR(Eval(L"=COUNT(A1:A4294967296)"), 0.0, 1e-9);
    ASSERT_TRUE(ElapsedSeconds(start) < 1.0);
}

TEST(WholeGridRange_SumsWhatIsThere)
{
    // A range spanning the whole grid visits the populated cells, not all
    // 2.9 million positions.
    SheetReset();
    SheetSet(0, 0, L"1");
    SheetSet(100, 300, L"2");
    SheetSet(kSheetRows - 1, kSheetCols - 1, L"4");
    SheetSet(5, 5, L"=SUM(B1:B2)");             // 0, and a formula inside the range
    ASSERT_NEAR(Eval(L"=SUM(A1:ZZ4096)"), 7.0, 1e-9);
    ASSERT_NEAR(Eval(L"=COUNT(A1:ZZ4096)"), 4.0, 1e-9);
    ASSERT_NEAR(Eval(L"=SUM(B2:ZY4095)"), 2.0, 1e-9);   // excludes row 1, col A, the last row and column
}

// ---- Memoization (P4) -------------------------------------------------------
TEST(Memo_DoubledChainEvaluatesEachCellOnce)
{
    // Every cell reads the one before twice. Without a memo, evaluating A40
    // walks 2^39 paths; with one, each formula cell is evaluated once.
    SheetReset();
    SheetSet(0, 0, L"1");
    for (UINT i = 1; i < 40; ++i)
    {
        wchar_t buf[32];
        swprintf_s(buf, 32, L"=A%u+A%u", i, i);
        SheetSet(i, 0, buf);
    }
    size_t before = CFormulator::CellEvaluationCount();
    auto start = std::chrono::steady_clock::now();
    ASSERT_NEAR(Eval(L"=A40"), std::ldexp(1.0, 39), 1e-3);
    ASSERT_TRUE(CFormulator::CellEvaluationCount() - before == 39);
    ASSERT_TRUE(ElapsedSeconds(start) < 1.0);
}

TEST(DeepChain_ResolvesExactly)
{
    // Far longer than the depth guard allows in one descent. It used to come
    // out truncated (the guard read the rest of the chain as 0); the deepest
    // cell reached is now evaluated first and the descent retried.
    SheetReset();
    SheetSet(0, 0, L"1");
    for (UINT i = 1; i < 2000; ++i)
    {
        wchar_t buf[32];
        swprintf_s(buf, 32, L"=A%u+1", i);
        SheetSet(i, 0, buf);
    }
    ASSERT_NEAR(Eval(L"=A2000"), 2000.0, 1e-9);
}

TEST(DeepDoubledChain_ResolvesInLinearWork)
{
    // Both at once: deep, and every cell read twice. A(k) = k.
    const UINT n = 3000;
    SheetReset();
    SheetSet(0, 0, L"1");
    for (UINT i = 1; i < n; ++i)
    {
        wchar_t buf[48];
        swprintf_s(buf, 48, L"=(A%u+A%u)/2+1", i, i);
        SheetSet(i, 0, buf);
    }
    size_t before = CFormulator::CellEvaluationCount();
    auto start = std::chrono::steady_clock::now();
    ASSERT_NEAR(Eval(L"=A3000"), (double)n, 1e-6);
    ASSERT_TRUE(CFormulator::CellEvaluationCount() - before <= 3 * (size_t)n);
    ASSERT_TRUE(ElapsedSeconds(start) < 5.0);
}

// ---- Parser fixes (2026-11) -------------------------------------------------
TEST(Mod_TakesTheDivisorsSign)
{
    // std::fmod takes the dividend's sign; Excel's MOD takes the divisor's.
    StandardSheet();
    ASSERT_NEAR(Eval(L"=MOD(7,3)"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=MOD(-3,2)"), 1.0, 1e-12);      // was -1
    ASSERT_NEAR(Eval(L"=MOD(3,-2)"), -1.0, 1e-12);     // was 1
    ASSERT_NEAR(Eval(L"=MOD(-3,-2)"), -1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=MOD(5.5,-2)"), -0.5, 1e-12);
    ASSERT_TRUE(IsError(L"MOD(1,0)", CFormulator::kErrorDiv0));
    // An exact multiple is +0, not -0 (which displays as "-0").
    CFormulator::Result r = CFormulator::Evaluate(g_sheet, L"MOD(-6,3)");
    ASSERT_TRUE(r.value == 0.0 && !std::signbit(r.value));
}

TEST(NumberLiterals_SignedExponent)
{
    // The token scan stopped at the exponent's sign: "1e-3" read as 1 - 3.
    StandardSheet();
    ASSERT_NEAR(Eval(L"=1e-3"), 0.001, 1e-15);         // was -2
    ASSERT_NEAR(Eval(L"=2.5e+2"), 250.0, 1e-12);       // was 4.5
    ASSERT_NEAR(Eval(L"=2E-1*10"), 2.0, 1e-12);
    ASSERT_NEAR(Eval(L"=-1.5e-7"), -1.5e-7, 1e-20);
    ASSERT_NEAR(Eval(L"=1e3"), 1000.0, 1e-12);
    ASSERT_NEAR(Eval(L"=A1-1"), 4.0, 1e-12);           // references untouched
}

TEST(PowerAndRoot_SplitAtTopLevelCommas)
{
    // arg.find(L',') cut POWER(ROUND(2.5,0),2) inside ROUND's argument list.
    StandardSheet();
    ASSERT_NEAR(Eval(L"=POWER(ROUND(2.5,0),2)"), 9.0, 1e-12);   // was 1
    ASSERT_NEAR(Eval(L"=ROOT(POWER(3,4),4)"), 3.0, 1e-12);      // was 1
    ASSERT_NEAR(Eval(L"=POWER(2,10)"), 1024.0, 1e-12);
    ASSERT_NEAR(Eval(L"=ROOT(16)"), 4.0, 1e-12);
    ASSERT_NEAR(Eval(L"=ROOT(27,3)"), 3.0, 1e-12);
}

TEST(FunctionCall_ParenthesesInStringsAreText)
{
    // The scan for the closing ')' ignored quotes: LEN("a)b") measured "a.
    StandardSheet();
    ASSERT_NEAR(Eval(L"=LEN(\"a)b\")"), 3.0, 1e-12);           // was 1
    ASSERT_NEAR(Eval(L"=LEN(\"a(b\")+1"), 4.0, 1e-12);         // was 3
    ASSERT_NEAR(Eval(L"=LEN(\"a\"\"b)\")"), 4.0, 1e-12);       // escaped quote
    CFormulator::Result r = CFormulator::Evaluate(g_sheet, L"UPPER(\"x)y\")");
    ASSERT_TRUE(r.isText && r.text == L"X)Y");                  // was "X"
    r = CFormulator::Evaluate(g_sheet, L"CONCAT(\"(\",\")\")");
    ASSERT_TRUE(r.isText && r.text == L"()");
}

TEST(TextFunction_DoesNotLeakThroughArithmetic)
{
    // The same leak TextResult_DoesNotLeakIntoArithmetic closed for a
    // referenced cell, but with the text function inline: whatever consumed
    // its result, the formula still showed the text.
    StandardSheet();
    CFormulator::Result r = CFormulator::Evaluate(g_sheet, L"1+LEFT(\"5\",1)");
    ASSERT_FALSE(r.isText);
    ASSERT_NEAR(r.value, 6.0, 1e-12);                           // was "5"
    r = CFormulator::Evaluate(g_sheet, L"ABS(LEFT(\"-5\",2))");
    ASSERT_FALSE(r.isText);
    ASSERT_NEAR(r.value, 5.0, 1e-12);                           // was "-5"
    r = CFormulator::Evaluate(g_sheet, L"-LEFT(\"5\",1)");
    ASSERT_FALSE(r.isText);
    ASSERT_NEAR(r.value, -5.0, 1e-12);
    r = CFormulator::Evaluate(g_sheet, L"LEFT(\"x\",1)*2");
    ASSERT_FALSE(r.isText);
    ASSERT_TRUE(r.isError && r.text == CFormulator::kErrorValue);   // "x" is no number
    r = CFormulator::Evaluate(g_sheet, L"IF(LEFT(\"1\",1),5,0)");
    ASSERT_FALSE(r.isText);                                     // the condition is consumed
    ASSERT_NEAR(r.value, 5.0, 1e-12);
    ASSERT_NEAR(Eval(L"=LEN(1+LEFT(\"5\",1))"), 1.0, 1e-12);    // "6"

    // What is text still is.
    r = CFormulator::Evaluate(g_sheet, L"IF(0,5,UPPER(\"n\"))");
    ASSERT_TRUE(r.isText && r.text == L"N");
    r = CFormulator::Evaluate(g_sheet, L"(UPPER(LEFT(\"abc\",2)))");
    ASSERT_TRUE(r.isText && r.text == L"AB");
    r = CFormulator::Evaluate(g_sheet, L"CONCAT(LEFT(\"ab\",1),1+1)");
    ASSERT_TRUE(r.isText && r.text == L"a2");
}

TEST(FormatGeneral_ShowsFifteenSignificantDigits)
{
    // The stream default of six digits showed 1234567 as "1.23457e+06", and
    // the text functions saw that rendering.
    StandardSheet();
    ASSERT_TRUE(CFormulator::FormatGeneral(1234567.0) == L"1234567");
    ASSERT_TRUE(CFormulator::FormatGeneral(1234567.89) == L"1234567.89");
    ASSERT_TRUE(CFormulator::FormatGeneral(0.1 + 0.2) == L"0.3");
    ASSERT_TRUE(CFormulator::FormatGeneral(-2.5) == L"-2.5");
    ASSERT_NEAR(Eval(L"=LEN(123456789)"), 9.0, 1e-12);          // was 11
    CFormulator::Result r = CFormulator::Evaluate(g_sheet, L"CONCAT(\"a\",1234567)");
    ASSERT_TRUE(r.isText && r.text == L"a1234567");
}

// ---- Rounding on the decimal value -------------------------------------------
TEST(Round_RoundsTheDecimalTheUserWrote)
{
    // 1.005 is stored as 1.00499999999999989..., so a plain round gave 1.
    StandardSheet();
    ASSERT_NEAR(Eval(L"=ROUND(1.005,2)"), 1.01, 1e-12);     // was 1
    ASSERT_NEAR(Eval(L"=ROUND(1.015,2)"), 1.02, 1e-12);     // was 1.01
    ASSERT_NEAR(Eval(L"=ROUND(0.285,2)"), 0.29, 1e-12);     // was 0.28
    ASSERT_NEAR(Eval(L"=ROUND(-1.005,2)"), -1.01, 1e-12);
    ASSERT_NEAR(Eval(L"=ROUND(2.5,0)"), 3.0, 1e-12);        // half away from zero
    ASSERT_NEAR(Eval(L"=ROUND(-2.5,0)"), -3.0, 1e-12);
    ASSERT_NEAR(Eval(L"=ROUND(1250,-2)"), 1300.0, 1e-9);
    ASSERT_NEAR(Eval(L"=ROUND(3.14159,2)"), 3.14, 1e-12);
    ASSERT_NEAR(Eval(L"=ROUND(1.25,1.9)"), 1.3, 1e-12);     // digits truncate to 1; was 1.2463...
    ASSERT_NEAR(Eval(L"=ROUND(1,400)"), 1.0, 1e-12);        // was nan
    ASSERT_NEAR(Eval(L"=ROUND(123,-400)"), 0.0, 1e-12);     // was nan
}

TEST(CeilingFloorTrunc_RoundTheDecimalToo)
{
    StandardSheet();
    ASSERT_NEAR(Eval(L"=FLOOR(0.3,0.1)"), 0.3, 1e-12);      // was 0.2
    ASSERT_NEAR(Eval(L"=CEILING(2.1,0.1)"), 2.1, 1e-12);
    ASSERT_NEAR(Eval(L"=CEILING(4.42,0.05)"), 4.45, 1e-12);
    ASSERT_NEAR(Eval(L"=FLOOR(-2.1,1)"), -3.0, 1e-12);
    ASSERT_NEAR(Eval(L"=TRUNC(0.29,2)"), 0.29, 1e-12);      // the digits argument was ignored
    ASSERT_NEAR(Eval(L"=TRUNC(-1.79,1)"), -1.7, 1e-12);
    ASSERT_NEAR(Eval(L"=TRUNC(1234,-2)"), 1200.0, 1e-9);
    ASSERT_NEAR(Eval(L"=TRUNC(-1.7)"), -1.0, 1e-12);
}

// ---- Text values: literals and comparisons ----------------------------------

TEST(StringLiterals_AreTextValues)
{
    // A quoted literal outside a text function read as 0.
    StandardSheet();
    ASSERT_TRUE(IsText(L"\"abc\"", L"abc"));
    ASSERT_TRUE(IsText(L"\"a\"\"b\"", L"a\"b"));
    ASSERT_TRUE(IsText(L"IF(A1>3,\"yes\",\"no\")", L"yes"));    // was 0
    ASSERT_TRUE(IsText(L"IF(A1<3,\"yes\",\"no\")", L"no"));
    ASSERT_TRUE(IsText(L"IF(A1>3,UPPER(\"y\"),0)", L"Y"));
    ASSERT_NEAR(Eval(L"=\"5\"+1"), 6.0, 1e-12);                 // arithmetic reads it as a number
    ASSERT_NEAR(Eval(L"=LEN(\"a\"\"b\")"), 3.0, 1e-12);
    // Everything after a literal's closing quote used to be ignored in a
    // text function's argument.
    ASSERT_TRUE(IsError(L"UPPER(\"a\"+1)", CFormulator::kErrorValue));   // "a" is no number
    CFormulator::Result r = CFormulator::Evaluate(g_sheet, L"\"abc");
    ASSERT_TRUE(r.isError && r.text == CFormulator::kErrorSyntax);
}

TEST(References_ToTextCarryTheText)
{
    SheetReset();
    SheetSet(0, 0, L"hello");               // A1
    SheetSet(1, 0, L"=UPPER(\"abc\")");     // A2
    ASSERT_TRUE(IsText(L"(A1)", L"hello"));
    ASSERT_TRUE(IsText(L"IF(1,A1,0)", L"hello"));
    ASSERT_TRUE(IsText(L"IF(1,A2,0)", L"ABC"));
    ASSERT_TRUE(IsError(L"A1+1", CFormulator::kErrorValue));     // text in arithmetic
}

TEST(Comparisons_CompareText)
{
    SheetReset();
    SheetSet(0, 0, L"hello");               // A1
    SheetSet(1, 0, L"=\"ABC\"");            // A2, a text result
    SheetSetText(2, 0, L"42");              // A3, text that reads as a number
    SheetSet(3, 0, L"5");                   // A4
    ASSERT_NEAR(Eval(L"=A1=\"hello\""), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=A1=\"HELLO\""), 1.0, 1e-12);             // case-insensitive, as in Excel
    ASSERT_NEAR(Eval(L"=A1<>\"hello\""), 0.0, 1e-12);
    ASSERT_NEAR(Eval(L"=A1=\"world\""), 0.0, 1e-12);             // was 1: both read as 0
    ASSERT_NEAR(Eval(L"=A1>\"abc\""), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=A1<\"abc\""), 0.0, 1e-12);
    ASSERT_NEAR(Eval(L"=A2=\"abc\""), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=LEFT(A1,1)=\"h\""), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=IF(A1=\"hello\",A4*2,0)"), 10.0, 1e-12);
    // Text is greater than any number, unless it reads as one.
    ASSERT_NEAR(Eval(L"=A1>99999"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=A4<\"x\""), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=A3=42"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=A3>41"), 1.0, 1e-12);
    // Text in a comparison inside arithmetic is not mistaken for a sibling's.
    ASSERT_TRUE(IsError(L"LEFT(\"a\",1)*(A4=5)", CFormulator::kErrorValue));
    ASSERT_NEAR(Eval(L"=2*(A4=5)"), 2.0, 1e-12);
}

TEST(Comparisons_BlankIsEmptyTextAndZero)
{
    SheetReset();
    SheetSet(0, 0, L"0");                   // A1
    SheetSet(1, 0, L"x");                   // A2
    // C1 is blank.
    ASSERT_NEAR(Eval(L"=C1=\"\""), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=C1=0"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=C1<>\"\""), 0.0, 1e-12);
    ASSERT_NEAR(Eval(L"=\"\"=C1"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=C1=D1"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=A1=\"\""), 0.0, 1e-12);                  // 0 is not empty
    ASSERT_NEAR(Eval(L"=A2=\"\""), 0.0, 1e-12);
    ASSERT_TRUE(IsText(L"IF(C1=\"\",\"empty\",\"full\")", L"empty"));
    ASSERT_TRUE(IsText(L"IF(A2=\"\",\"empty\",\"full\")", L"full"));
}

// ---- Error values (E2, part) --------------------------------------------------

TEST(Errors_MalformedFormulasAreSyntaxErrors)
{
    // Each of these used to show whatever had been parsed before the problem.
    StandardSheet();
    const wchar_t* kSyntax = CFormulator::kErrorSyntax;
    ASSERT_TRUE(IsError(L"12abc", kSyntax));        // was 12
    ASSERT_TRUE(IsError(L"1.2.3", kSyntax));
    ASSERT_TRUE(IsError(L"2(3)", kSyntax));         // was 0
    ASSERT_TRUE(IsError(L"1 2", kSyntax));          // was 1
    ASSERT_TRUE(IsError(L"1)", kSyntax));
    ASSERT_TRUE(IsError(L"1+", kSyntax));           // was 1
    ASSERT_TRUE(IsError(L"", kSyntax));             // a cell holding just "="
    ASSERT_TRUE(IsError(L"*2", kSyntax));
    ASSERT_TRUE(IsError(L"(1+2", kSyntax));         // was 3
    ASSERT_TRUE(IsError(L"SUM(A1:A3", kSyntax));    // was 21
    ASSERT_TRUE(IsError(L"A1:A3", kSyntax));        // a range outside a function
    ASSERT_TRUE(IsError(L"ABS(1,2)", kSyntax));     // was 1
    ASSERT_TRUE(IsError(L"POWER(2)", kSyntax));     // was 1
    // Syntax the engine does not have yet is flagged rather than misread.
    ASSERT_TRUE(IsError(L"2^3", kSyntax));          // was 2
    ASSERT_TRUE(IsError(L"50%", kSyntax));          // was 50
    ASSERT_TRUE(IsError(L"$A$1", kSyntax));         // was 0
}

TEST(Errors_UnknownNames)
{
    StandardSheet();
    ASSERT_TRUE(IsError(L"abc", CFormulator::kErrorName));       // was 0
    ASSERT_TRUE(IsError(L"FOO(1)", CFormulator::kErrorName));    // was 0
    ASSERT_TRUE(IsError(L"A1+SUMM(A1:A3)", CFormulator::kErrorName));
    // TRUE and FALSE are not unknown names.
    ASSERT_NEAR(Eval(L"=TRUE"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=FALSE+1"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=IF(TRUE(),2,3)"), 2.0, 1e-12);
    // A reference off the grid still reads as 0 (H21), not as a name.
    ASSERT_NEAR(Eval(L"=A4097"), 0.0, 1e-12);
}

TEST(Errors_ArithmeticWithNoResult)
{
    StandardSheet();
    ASSERT_TRUE(IsError(L"1/0", CFormulator::kErrorDiv0));          // was inf
    ASSERT_TRUE(IsError(L"A1/(B1-2)", CFormulator::kErrorDiv0));
    ASSERT_TRUE(IsError(L"MOD(1,0)", CFormulator::kErrorDiv0));     // was 0
    ASSERT_TRUE(IsError(L"AVERAGE(D1:D9)", CFormulator::kErrorDiv0)); // was 0
    ASSERT_TRUE(IsError(L"SQRT(-1)", CFormulator::kErrorNum));      // was 0
    ASSERT_TRUE(IsError(L"LOG(0)", CFormulator::kErrorNum));        // was -inf
    ASSERT_TRUE(IsError(L"LN(-1)", CFormulator::kErrorNum));        // was nan
    ASSERT_TRUE(IsError(L"POWER(10,400)", CFormulator::kErrorNum)); // was inf
    ASSERT_TRUE(IsError(L"1e999", CFormulator::kErrorNum));
    // The first error wins.
    ASSERT_TRUE(IsError(L"1/0+abc", CFormulator::kErrorDiv0));
    ASSERT_TRUE(IsError(L"abc+1/0", CFormulator::kErrorName));
    // MIN and MAX of nothing are still 0, as in Excel.
    ASSERT_NEAR(Eval(L"=MAX(D1:D9)"), 0.0, 1e-12);
}

TEST(Errors_Propagate)
{
    StandardSheet();
    SheetSet(0, 2, L"=1/0");                // C1
    SheetSet(1, 2, L"=C1+1");               // C2
    ASSERT_TRUE(IsError(L"C1", CFormulator::kErrorDiv0));
    ASSERT_TRUE(IsError(L"C2*2", CFormulator::kErrorDiv0));
    ASSERT_TRUE(IsError(L"SUM(A1:C3)", CFormulator::kErrorDiv0));
    ASSERT_TRUE(IsError(L"MAX(C1:C2)", CFormulator::kErrorDiv0));
    ASSERT_TRUE(IsError(L"LEN(C1)", CFormulator::kErrorDiv0));
    ASSERT_TRUE(IsError(L"C1=1", CFormulator::kErrorDiv0));
    ASSERT_TRUE(IsError(L"SUM(1/0,1)", CFormulator::kErrorDiv0));
    // Only what is evaluated counts: IF does not evaluate the other branch.
    ASSERT_NEAR(Eval(L"=IF(1,5,C1)"), 5.0, 1e-12);
    ASSERT_TRUE(IsError(L"IF(0,5,C1)", CFormulator::kErrorDiv0));
    // COUNT counts numbers; an error is not one, and does not make COUNT fail.
    ASSERT_NEAR(Eval(L"=COUNT(A1:C3)"), 5.0, 1e-12);
    ASSERT_NEAR(Eval(L"=COUNT(1/0,1)"), 1.0, 1e-12);
    // A formula that reads an error is not itself reported as the cause.
    ASSERT_NEAR(Eval(L"=A1+1"), 6.0, 1e-12);
}

TEST(ChainedComparison_IsAnErrorNotIgnored)
{
    // Only one comparison is read. The rest used to be dropped silently,
    // so =1=1=2 was TRUE; now it is flagged until chaining is implemented.
    StandardSheet();
    ASSERT_TRUE(IsError(L"1=1=2", CFormulator::kErrorSyntax));
    ASSERT_TRUE(IsError(L"3>2>1", CFormulator::kErrorSyntax));
    ASSERT_NEAR(Eval(L"=(1=1)=1"), 1.0, 1e-12);     // parenthesized, it is two comparisons
}

// ---- Aggregates take several arguments (E4) -----------------------------------
TEST(Aggregates_TakeSeveralArguments)
{
    StandardSheet();                        // A1=5 A2=7 A3=9 B1=2 B2=4
    ASSERT_NEAR(Eval(L"=SUM(A1:A3,B1)"), 23.0, 1e-12);                 // was 0
    ASSERT_NEAR(Eval(L"=SUM(A1:A3, B1:B2, 100)"), 127.0, 1e-12);
    ASSERT_NEAR(Eval(L"=SUM(A1,A2)"), 12.0, 1e-12);
    ASSERT_NEAR(Eval(L"=SUM(A1,,A2)"), 12.0, 1e-12);                   // an omitted argument
    ASSERT_NEAR(Eval(L"=SUM(A1*2,1)"), 11.0, 1e-12);
    ASSERT_NEAR(Eval(L"=AVERAGE(A1:A3,10)"), 7.75, 1e-12);
    ASSERT_NEAR(Eval(L"=MIN(A1:A3,0)"), 0.0, 1e-12);
    ASSERT_NEAR(Eval(L"=MAX(A1:A3,B1*10)"), 20.0, 1e-12);
    ASSERT_NEAR(Eval(L"=COUNT(A1:A3,1,\"x\",\"5\")"), 5.0, 1e-12);      // "x" is not a number
    ASSERT_TRUE(IsError(L"SUM(\"abc\",1)", CFormulator::kErrorValue));   // given as text: no number
    ASSERT_NEAR(Eval(L"=SUM(\"5\",1)"), 6.0, 1e-12);
    ASSERT_NEAR(Eval(L"=SUM(A1:A4097,1)"), 1.0, 1e-12);               // off the grid: nothing
    ASSERT_NEAR(Eval(L"=SUM(SUM(A1:A2),A3)"), 21.0, 1e-12);
}

// ---- Quick fixes (2026-11, round 4) ---------------------------------------------
TEST(Comparisons_ExactAtFifteenDigits)
{
    // The tolerance was 1e-9 of the larger side, so it grew with the numbers.
    StandardSheet();
    ASSERT_NEAR(Eval(L"=1000000000=1000000001"), 0.0, 1e-12);      // was 1
    ASSERT_NEAR(Eval(L"=1000000000<1000000001"), 1.0, 1e-12);      // was 0
    ASSERT_NEAR(Eval(L"=1000000001>1000000000"), 1.0, 1e-12);      // was 0
    ASSERT_NEAR(Eval(L"=1000000000<>1000000001"), 1.0, 1e-12);     // was 0
    ASSERT_NEAR(Eval(L"=123456789012=123456789000"), 0.0, 1e-12);  // was 1
    ASSERT_NEAR(Eval(L"=0.1234567890=0.1234567891"), 0.0, 1e-12);  // was 1
    // One second apart on a date: 45306 + 1/86400. The tolerance was ~4 s.
    ASSERT_NEAR(Eval(L"=45306+1/86400>45306"), 1.0, 1e-12);        // was 0
    ASSERT_NEAR(Eval(L"=45306+1/86400=45306"), 0.0, 1e-12);        // was 1
    // Binary rounding still does not make equal decimals unequal.
    ASSERT_NEAR(Eval(L"=0.1+0.2=0.3"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=1-0.9=0.1"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=0.1*3=0.3"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=0.1+0.2<=0.3"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=0.1+0.2>0.3"), 0.0, 1e-12);
    // A leftover that is not 0 at 15 digits is not 0, as in Excel.
    ASSERT_NEAR(Eval(L"=0.1+0.2-0.3=0"), 0.0, 1e-12);
}

// Run `body` with a locale whose iswalpha knows letters beyond ASCII, as
// MSVC's does, so the test means the same on Linux as on Windows.
template <class Fn>
static void WithUnicodeLocale(Fn body)
{
    const char* old = std::setlocale(LC_ALL, nullptr);
    const std::string saved = old ? old : "C";
    if (!std::setlocale(LC_ALL, "C.UTF-8") && !std::setlocale(LC_ALL, ".UTF8"))
        std::cout << "  note: no UTF-8 locale; running under the current one\n";
    body();
    std::setlocale(LC_ALL, saved.c_str());
}

TEST(ParseCellRef_ColumnLettersAreAsciiOnly)
{
    // Under a Unicode-aware iswalpha, "É1" parsed as column 'É'-'A'+1 = 137.
    WithUnicodeLocale([] {
        UINT r = 0, c = 0;
        ASSERT_FALSE(CFormulator::ParseCellRef(L"\u00C91", r, c));      // É1, was EG1
        ASSERT_FALSE(CFormulator::ParseCellRef(L"\u00E91", r, c));      // é1
        ASSERT_FALSE(CFormulator::ParseCellRef(L"A\u00C91", r, c));
        ASSERT_FALSE(CFormulator::ParseCellRef(L"A\u0661", r, c));      // an Arabic-Indic digit
        ASSERT_TRUE(CFormulator::ParseCellRef(L"eg1", r, c));
        ASSERT_TRUE(r == 0 && c == 136);

        SheetReset();
        SheetSet(0, 136, L"7");                                       // EG1
        CFormulator::Result res = CFormulator::Evaluate(g_sheet, L"\u00C91");
        ASSERT_TRUE(res.isError);                                     // was 7
    });
}

TEST(TextFunctions_HugeCountsDoNotOverflow)
{
    // The counts were cast to int, undefined past INT_MAX.
    StandardSheet();
    ASSERT_TRUE(IsText(L"LEFT(\"abc\",1e20)", L"abc"));           // was ""
    ASSERT_TRUE(IsText(L"RIGHT(\"abc\",3e9)", L"abc"));           // was ""
    ASSERT_TRUE(IsText(L"MID(\"abcdef\",2,1e10)", L"bcdef"));     // was ""
    ASSERT_TRUE(IsText(L"MID(\"abcdef\",3e9,2)", L""));           // was "ab"
    // In range they behave as before: counts truncate, a start below 1 is 1.
    ASSERT_TRUE(IsText(L"LEFT(\"abc\",2.9)", L"ab"));
    ASSERT_TRUE(IsText(L"LEFT(\"abc\")", L"a"));
    ASSERT_TRUE(IsText(L"LEFT(\"abc\",0)", L""));
    ASSERT_TRUE(IsText(L"RIGHT(\"abc\",2)", L"bc"));
    ASSERT_TRUE(IsText(L"MID(\"abcdef\",2.9,2.9)", L"bc"));
    ASSERT_TRUE(IsError(L"MID(\"abcdef\",-5,2)", CFormulator::kErrorValue));   // no position -5
    ASSERT_TRUE(IsText(L"MID(\"abcdef\",7,2)", L""));
}

TEST(DateParts_OutOfRangeAreNumErrors)
{
    StandardSheet();
    ASSERT_TRUE(IsError(L"DATE(2024,1,1e10)", CFormulator::kErrorNum));   // undefined cast
    ASSERT_TRUE(IsError(L"DATE(1e300,1,1)", CFormulator::kErrorNum));
    ASSERT_TRUE(IsError(L"YEAR(1e15)", CFormulator::kErrorNum));          // was 2012841336
    ASSERT_TRUE(IsError(L"DAY(-1e15)", CFormulator::kErrorNum));
    ASSERT_NEAR(Eval(L"=DATE(2024,1,15)"), 45306.0, 1e-9);
    ASSERT_NEAR(Eval(L"=YEAR(DATE(2024,1,15))"), 2024.0, 1e-9);
    ASSERT_NEAR(Eval(L"=MONTH(45306)"), 1.0, 1e-9);
}

TEST(Date_FollowsExcel)
{
    StandardSheet();
    // Years 0..1899 are offsets from 1900.
    ASSERT_NEAR(Eval(L"=DATE(24,1,1)"), 8767.0, 1e-9);               // 1924-01-01; was -685193
    ASSERT_NEAR(Eval(L"=YEAR(DATE(24,1,1))"), 1924.0, 1e-9);         // was 24
    ASSERT_NEAR(Eval(L"=DATE(0,1,1)"), 1.0, 1e-9);
    ASSERT_NEAR(Eval(L"=YEAR(DATE(1899,12,31))"), 3799.0, 1e-9);
    // Excel's serials, 1900-02-29 included.
    ASSERT_NEAR(Eval(L"=DATE(1900,1,1)"), 1.0, 1e-9);                // was 2
    ASSERT_NEAR(Eval(L"=DATE(1900,2,28)"), 59.0, 1e-9);              // was 60
    ASSERT_NEAR(Eval(L"=DATE(1900,2,29)"), 60.0, 1e-9);
    ASSERT_NEAR(Eval(L"=DATE(1900,3,0)"), 60.0, 1e-9);
    ASSERT_NEAR(Eval(L"=DATE(1900,3,1)"), 61.0, 1e-9);
    ASSERT_NEAR(Eval(L"=DATE(1900,1,0)"), 0.0, 1e-9);
    ASSERT_NEAR(Eval(L"=DATE(9999,12,31)"), 2958465.0, 1e-9);
    // Months and days roll over, however far.
    ASSERT_NEAR(Eval(L"=DATE(2024,15,1)-DATE(2025,3,1)"), 0.0, 1e-9);   // was 2
    ASSERT_NEAR(Eval(L"=DATE(2024,-10,1)-DATE(2023,2,1)"), 0.0, 1e-9);  // was -3
    ASSERT_NEAR(Eval(L"=DATE(2024,1,0)-DATE(2023,12,31)"), 0.0, 1e-9);
    ASSERT_NEAR(Eval(L"=DATE(2024.9,1,1)-DATE(2024,1,1)"), 0.0, 1e-9);  // parts truncate
    // Outside 1900-01-00 .. 9999-12-31 there is no date.
    ASSERT_TRUE(IsError(L"DATE(1900,1,-1)", CFormulator::kErrorNum));
    ASSERT_TRUE(IsError(L"DATE(-1,1,1)", CFormulator::kErrorNum));
    ASSERT_TRUE(IsError(L"DATE(9999,12,32)", CFormulator::kErrorNum));
    ASSERT_TRUE(IsError(L"DATE(10000,1,1)", CFormulator::kErrorNum));      // was 2958466
}

TEST(DateParts_FollowExcel)
{
    StandardSheet();
    ASSERT_NEAR(Eval(L"=YEAR(0)"), 1900.0, 1e-9);                    // 1900-01-00; was 1899
    ASSERT_NEAR(Eval(L"=MONTH(0)"), 1.0, 1e-9);
    ASSERT_NEAR(Eval(L"=DAY(0)"), 0.0, 1e-9);
    ASSERT_NEAR(Eval(L"=DAY(1)"), 1.0, 1e-9);                        // 1900-01-01
    ASSERT_NEAR(Eval(L"=MONTH(60)"), 2.0, 1e-9);                     // 1900-02-29
    ASSERT_NEAR(Eval(L"=DAY(60)"), 29.0, 1e-9);
    ASSERT_NEAR(Eval(L"=DAY(61)"), 1.0, 1e-9);
    ASSERT_NEAR(Eval(L"=YEAR(2958465)"), 9999.0, 1e-9);
    ASSERT_NEAR(Eval(L"=DAY(45306.75)"), 15.0, 1e-9);                // the time of day is ignored
    ASSERT_TRUE(IsError(L"YEAR(-5)", CFormulator::kErrorNum));        // was 1899
    ASSERT_TRUE(IsError(L"YEAR(2958466)", CFormulator::kErrorNum));
    ASSERT_TRUE(IsError(L"MONTH(-0.5)", CFormulator::kErrorNum));
}

TEST(NegativeZero_IsZero)
{
    // =-A1 with A1 blank showed "-0".
    StandardSheet();
    CFormulator::Result r = CFormulator::Evaluate(g_sheet, L"-D9");
    ASSERT_TRUE(!r.isError && r.value == 0.0 && !std::signbit(r.value));
    r = CFormulator::Evaluate(g_sheet, L"0*-1");
    ASSERT_TRUE(r.value == 0.0 && !std::signbit(r.value));
    ASSERT_TRUE(CFormulator::FormatGeneral(-0.0) == L"0");
    ASSERT_NEAR(Eval(L"=LEN(-0)"), 1.0, 1e-12);                         // "0", was "-0"
    ASSERT_TRUE(CFormulator::FormatGeneral(-2.5) == L"-2.5");
}

TEST(Trim_TouchesSpacesOnly)
{
    // Excel's TRIM removes spaces (U+0020) only. iswspace also took tabs and
    // line breaks, and turned each into a space.
    StandardSheet();
    ASSERT_TRUE(IsText(L"TRIM(\"  a   b  \")", L"a b"));
    ASSERT_TRUE(IsText(L"TRIM(\"a\tb\")", L"a\tb"));                  // was "a b"
    ASSERT_TRUE(IsText(L"TRIM(\"\tx \")", L"\tx"));                  // was "x"
    ASSERT_TRUE(IsText(L"TRIM(\"a\nb\")", L"a\nb"));
    ASSERT_TRUE(IsText(L"TRIM(\"a \t b\")", L"a \t b"));
    ASSERT_TRUE(IsText(L"TRIM(\"a\u00A0b\")", L"a\u00A0b"));          // no-break space is text
    ASSERT_TRUE(IsText(L"TRIM(\"   \")", L""));
}

// ---- VALUE ---------------------------------------------------------------------
TEST(Value_ReadsNumericTextOrFails)
{
    SheetReset();
    SheetSetText(0, 0, L"42");              // A1..A9: text cells, as `text A1 42` stores them
    SheetSetText(1, 0, L"abc");
    SheetSetText(2, 0, L"$1,234.50");
    SheetSetText(3, 0, L"  -3.5  ");
    SheetSetText(4, 0, L"25%");
    SheetSetText(5, 0, L"2024-01-15");
    SheetSetText(6, 0, L"1,2");
    SheetSetText(7, 0, L"12abc");
    SheetSetText(8, 0, L"0");
    SheetSet(0, 1, L"7");                   // B1: a number
    SheetSet(1, 1, L"2024-01-15");          // B2: a typed date
    SheetSet(2, 1, L"=UPPER(\"1e3\")");     // B3: a text result
    SheetSet(3, 1, L"=1/0");                // B4: an error

    ASSERT_NEAR(Eval(L"=VALUE(A1)"), 42.0, 1e-12);
    ASSERT_NEAR(Eval(L"=VALUE(A3)"), 1234.5, 1e-12);
    ASSERT_NEAR(Eval(L"=VALUE(A4)"), -3.5, 1e-12);
    ASSERT_NEAR(Eval(L"=VALUE(A5)"), 0.25, 1e-12);
    ASSERT_NEAR(Eval(L"=VALUE(A6)"), 45306.0, 1e-12);        // date text reads as its serial
    ASSERT_NEAR(Eval(L"=VALUE(A9)"), 0.0, 1e-12);            // "0" is a number...
    ASSERT_TRUE(IsError(L"VALUE(A2)", CFormulator::kErrorValue));   // ...and "abc" is not
    ASSERT_TRUE(IsError(L"VALUE(A7)", CFormulator::kErrorValue));   // badly grouped
    ASSERT_TRUE(IsError(L"VALUE(A8)", CFormulator::kErrorValue));   // trailing junk
    ASSERT_TRUE(IsError(L"VALUE(\"\")", CFormulator::kErrorValue));
    // Literals and text results read the same way.
    ASSERT_NEAR(Eval(L"=VALUE(\"5/15/2026\")"), 46157.0, 1e-12);
    ASSERT_NEAR(Eval(L"=VALUE(\" 2024-01-15 \")"), 45306.0, 1e-12);
    ASSERT_NEAR(Eval(L"=VALUE(B3)"), 1000.0, 1e-12);
    ASSERT_NEAR(Eval(L"=VALUE(LEFT(\"123abc\",3))*2"), 246.0, 1e-12);
    // Numbers pass through, a blank cell is 0, an error stays itself.
    ASSERT_NEAR(Eval(L"=VALUE(B1)"), 7.0, 1e-12);
    ASSERT_NEAR(Eval(L"=VALUE(B2)"), 45306.0, 1e-12);
    ASSERT_NEAR(Eval(L"=VALUE(42)"), 42.0, 1e-12);
    ASSERT_NEAR(Eval(L"=VALUE(Z99)"), 0.0, 1e-12);
    ASSERT_TRUE(IsError(L"VALUE(B4)", CFormulator::kErrorDiv0));
    // The result is a number, and the error travels like any other.
    CFormulator::Result r = CFormulator::Evaluate(g_sheet, L"VALUE(A1)");
    ASSERT_TRUE(!r.isText && !r.isError && r.value == 42.0);
    ASSERT_TRUE(IsError(L"VALUE(A2)+1", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"SUM(VALUE(A1),VALUE(A2))", CFormulator::kErrorValue));
    ASSERT_NEAR(Eval(L"=COUNT(VALUE(A1),VALUE(A2))"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=IF(1,5,VALUE(A2))"), 5.0, 1e-12);   // not evaluated, no error
    // Arithmetic reads text the same way now.
    ASSERT_TRUE(IsError(L"A2+0", CFormulator::kErrorValue));
    ASSERT_NEAR(Eval(L"=A1+0"), 42.0, 1e-12);
    // One argument, and not none.
    ASSERT_TRUE(IsError(L"VALUE(A1,A2)", CFormulator::kErrorSyntax));
    ASSERT_TRUE(IsError(L"VALUE()", CFormulator::kErrorSyntax));
}

TEST(Value_ErrorPropagatesThroughCells)
{
    SheetReset();
    SheetSetText(0, 0, L"abc");             // A1
    SheetSet(0, 1, L"=VALUE(A1)");          // B1
    SheetSet(0, 2, L"=B1*2");               // C1
    ASSERT_TRUE(IsError(L"C1", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"MAX(B1:C1)", CFormulator::kErrorValue));
}

// ---- #6: date results ------------------------------------------------------------
static bool IsDate(const wchar_t* formula, double serial)
{
    CFormulator::Result r = CFormulator::Evaluate(g_sheet, formula);
    return r.isDate && !r.isText && !r.isError && r.value == serial;
}
static bool IsPlainNumber(const wchar_t* formula, double value)
{
    CFormulator::Result r = CFormulator::Evaluate(g_sheet, formula);
    return !r.isDate && !r.isText && !r.isError && std::fabs(r.value - value) < 1e-9;
}

TEST(DateResults_AreMarkedAsDates)
{
    SheetReset();
    SheetSet(0, 0, L"2024-01-15");                       // A1: a typed date, 45306
    SheetSet(1, 0, L"2024-03-01");                       // A2: 45352
    SheetSetFormatted(2, 0, L"45306", FMT_DATE_ISO);     // A3: a number shown as a date
    SheetSet(3, 0, L"=A1+30");                           // A4: a formula date
    SheetSetFormatted(4, 0, L"=45306", FMT_DATE_US);     // A5: a formula formatted as a date
    SheetSet(5, 0, L"7");                                // A6: a plain number

    // Where dates come from.
    ASSERT_TRUE(IsDate(L"DATE(2024,1,15)", 45306.0));
    ASSERT_TRUE(IsDate(L"A1", 45306.0));
    ASSERT_TRUE(IsDate(L"(A1)", 45306.0));
    ASSERT_TRUE(IsDate(L"A3", 45306.0));
    ASSERT_TRUE(IsDate(L"A4", 45336.0));
    ASSERT_TRUE(IsDate(L"A5", 45306.0));
    CFormulator::Result today = CFormulator::Evaluate(g_sheet, L"TODAY()");
    ASSERT_TRUE(today.isDate && today.value > 46000.0);
    ASSERT_TRUE(CFormulator::Evaluate(g_sheet, L"NOW()").isDate);
    // A date moved by days is a date; the days between two dates are not.
    ASSERT_TRUE(IsDate(L"A1+7", 45313.0));
    ASSERT_TRUE(IsDate(L"7+A1", 45313.0));
    ASSERT_TRUE(IsDate(L"A1-7", 45299.0));
    ASSERT_TRUE(IsDate(L"A1+A6", 45313.0));
    ASSERT_TRUE(IsPlainNumber(L"A2-A1", 46.0));
    ASSERT_TRUE(IsPlainNumber(L"A1+A2", 90658.0));
    ASSERT_TRUE(IsPlainNumber(L"7-A1", 7.0 - 45306.0));
    ASSERT_TRUE(IsPlainNumber(L"A1*1", 45306.0));
    ASSERT_TRUE(IsPlainNumber(L"-A1", -45306.0));
    ASSERT_TRUE(IsPlainNumber(L"A6+1", 8.0));
    // A date on the left of * is not mistaken for the right operand's.
    ASSERT_TRUE(IsPlainNumber(L"A1*A6", 45306.0 * 7.0));
    // Functions: date functions and choices pass a date on; the rest do not.
    ASSERT_TRUE(IsDate(L"MAX(A1:A2)", 45352.0));
    ASSERT_TRUE(IsDate(L"MIN(A1,A2)", 45306.0));
    ASSERT_TRUE(IsPlainNumber(L"MAX(A1:A2,100000)", 100000.0));   // not all dates
    ASSERT_TRUE(IsDate(L"IF(A1>A2,A1,A2)", 45352.0));
    ASSERT_TRUE(IsDate(L"IFERROR(A1,0)", 45306.0));
    ASSERT_TRUE(IsPlainNumber(L"YEAR(A1)", 2024.0));
    ASSERT_TRUE(IsPlainNumber(L"ROUND(A1,0)", 45306.0));
    ASSERT_TRUE(IsPlainNumber(L"SUM(A1:A2)", 90658.0));
    ASSERT_TRUE(IsPlainNumber(L"A1=A3", 1.0));
    ASSERT_TRUE(IsPlainNumber(L"VALUE(\"2024-01-15\")", 45306.0));   // VALUE gives a number
}

TEST(DateParts_ReadDateText)
{
    StandardSheet();
    ASSERT_NEAR(Eval(L"=YEAR(\"2024-01-15\")"), 2024.0, 1e-9);     // was #NUM!
    ASSERT_NEAR(Eval(L"=MONTH(\"5/15/2026\")"), 5.0, 1e-9);
    ASSERT_NEAR(Eval(L"=DAY(\" 2024-01-15 \")"), 15.0, 1e-9);
    ASSERT_NEAR(Eval(L"=YEAR(\"45306\")"), 2024.0, 1e-9);          // numeric text too
    ASSERT_TRUE(IsError(L"DAY(\"abc\")", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"YEAR(\"\")", CFormulator::kErrorValue));
}

// ---- #7: ranges where Excel takes them -------------------------------------------
TEST(AndOr_TakeRangesAndSeeEveryArgument)
{
    SheetReset();
    SheetSet(0, 0, L"1");  SheetSet(1, 0, L"2");  SheetSet(2, 0, L"0");  SheetSet(3, 0, L"hello");
    SheetSet(0, 2, L"=1/0");                                       // C1
    ASSERT_NEAR(Eval(L"=AND(A1:A2)"), 1.0, 1e-12);                 // was #ERROR!
    ASSERT_NEAR(Eval(L"=AND(A1:A3)"), 0.0, 1e-12);
    ASSERT_NEAR(Eval(L"=AND(A1:A4)"), 0.0, 1e-12);                 // text is skipped
    ASSERT_NEAR(Eval(L"=AND(A1,A2)"), 1.0, 1e-12);                 // blanks are skipped
    ASSERT_NEAR(Eval(L"=OR(A1:A3)"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=OR(A3)"), 0.0, 1e-12);
    ASSERT_NEAR(Eval(L"=OR(A3, A1>0)"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=AND(A1>0,A2>1,A3<1)"), 1.0, 1e-12);
    // Every argument is looked at, so an error is the answer.
    ASSERT_TRUE(IsError(L"AND(0,1/0)", CFormulator::kErrorDiv0));   // was 0
    ASSERT_TRUE(IsError(L"OR(1,C1)", CFormulator::kErrorDiv0));     // was 1
    // Nothing to look at: no answer.
    ASSERT_TRUE(IsError(L"AND(A4)", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"OR(Z1:Z9)", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"AND()", CFormulator::kErrorValue));
}

TEST(Concat_TakesRanges)
{
    SheetReset();
    SheetSet(0, 0, L"1");  SheetSet(1, 0, L"2");  SheetSet(2, 0, L"x");
    SheetSet(0, 1, L"a");  SheetSet(1, 1, L"b");
    SheetSet(0, 2, L"=1/0");                                       // C1
    ASSERT_TRUE(IsText(L"CONCAT(A1:A3)", L"12x"));                  // was #ERROR!
    ASSERT_TRUE(IsText(L"CONCAT(A1:B2)", L"1a2b"));                 // row by row
    ASSERT_TRUE(IsText(L"CONCAT(A1:A9, \"-\", B1:B2)", L"12x-ab")); // blanks add nothing
    ASSERT_TRUE(IsText(L"CONCATENATE(\"<\",A1:A2,\">\")", L"<12>"));
    ASSERT_TRUE(IsText(L"CONCAT(A:A)", L"12x"));
    ASSERT_TRUE(IsText(L"CONCAT(A1,B1)", L"1a"));
    ASSERT_TRUE(IsText(L"CONCAT(A1:A4097)", L""));                  // off the grid: nothing
    ASSERT_TRUE(IsError(L"CONCAT(A1:C1)", CFormulator::kErrorDiv0));
}

TEST(Ranges_WholeColumnsAndRows)
{
    SheetReset();
    SheetSet(0, 0, L"1");  SheetSet(1, 0, L"2");  SheetSet(4095, 0, L"4");   // A1, A2, A4096
    SheetSet(0, 1, L"10"); SheetSet(1, 2, L"100");                           // B1, C2
    ASSERT_NEAR(Eval(L"=SUM(A:A)"), 7.0, 1e-12);                   // was #NAME?
    ASSERT_NEAR(Eval(L"=SUM(a:a)"), 7.0, 1e-12);
    ASSERT_NEAR(Eval(L"=SUM(A:B)"), 17.0, 1e-12);
    ASSERT_NEAR(Eval(L"=SUM(B:A)"), 17.0, 1e-12);                  // either end first
    ASSERT_NEAR(Eval(L"=SUM(1:1)"), 11.0, 1e-12);
    ASSERT_NEAR(Eval(L"=SUM(2:1)"), 113.0, 1e-12);
    ASSERT_NEAR(Eval(L"=SUM(4096:4096)"), 4.0, 1e-12);
    ASSERT_NEAR(Eval(L"=COUNT(A:C)"), 5.0, 1e-12);
    ASSERT_NEAR(Eval(L"=MAX(1:2)"), 100.0, 1e-12);
    ASSERT_NEAR(Eval(L"=SUM(A:A, 10)"), 17.0, 1e-12);
    ASSERT_NEAR(Eval(L"=SUM( A : A )"), 7.0, 1e-12);
    // Columns or rows past the grid are rejected whole, as a cell range is.
    ASSERT_NEAR(Eval(L"=SUM(A:AAA)"), 0.0, 1e-12);
    ASSERT_NEAR(Eval(L"=SUM(4097:4097)"), 0.0, 1e-12);
    ASSERT_TRUE(CFormulator::Evaluate(g_sheet, L"SUM(1:4294967297)").isError);   // no such row
    // Mixed forms are not ranges; read as an expression, "A" is an unknown name.
    ASSERT_TRUE(IsError(L"SUM(A:1)", CFormulator::kErrorName));
    // A formula inside its own column reads itself as 0 (a cycle), as before.
    SheetSet(9, 0, L"=SUM(A:A)");                                  // A10
    ASSERT_NEAR(Eval(L"=A10"), 7.0, 1e-12);
}

// ---- #8: IFERROR, ISERROR, ISBLANK -------------------------------------------------
TEST(IfError_CatchesValueErrorsOnly)
{
    SheetReset();
    SheetSet(0, 0, L"1");                                          // A1
    SheetSet(0, 2, L"=1/0");                                       // C1
    SheetSetText(1, 0, L"abc");                                    // A2
    ASSERT_TRUE(IsText(L"IFERROR(1/0,\"oops\")", L"oops"));
    ASSERT_NEAR(Eval(L"=IFERROR(C1,0)"), 0.0, 1e-12);
    ASSERT_NEAR(Eval(L"=IFERROR(C1*2,-1)"), -1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=IFERROR(A1*2,0)"), 2.0, 1e-12);            // no error: the value
    ASSERT_TRUE(IsText(L"IFERROR(UPPER(\"ok\"),0)", L"OK"));       // text passes through
    ASSERT_TRUE(IsText(L"IFERROR(FOO(1),\"unknown\")", L"unknown"));   // #NAME? is caught
    ASSERT_NEAR(Eval(L"=IFERROR(VALUE(A2),-1)"), -1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=IFERROR(SQRT(-1),0)"), 0.0, 1e-12);
    ASSERT_TRUE(IsError(L"IFERROR(1/0,A1/0)", CFormulator::kErrorDiv0));   // the fallback's own
    ASSERT_NEAR(Eval(L"=IFERROR(1/0,0)+5"), 5.0, 1e-12);
    // A malformed formula is not an error value: it is not hidden.
    ASSERT_TRUE(IsError(L"IFERROR(12abc,0)", CFormulator::kErrorSyntax));
    ASSERT_TRUE(IsError(L"IFERROR(1)", CFormulator::kErrorSyntax));
    ASSERT_TRUE(IsError(L"IFERROR(1,2,3)", CFormulator::kErrorSyntax));
    // An error before IFERROR is still the formula's.
    ASSERT_TRUE(IsError(L"1/0+IFERROR(1,2)", CFormulator::kErrorDiv0));
}

TEST(IsError_And_IsBlank)
{
    SheetReset();
    SheetSet(0, 0, L"1");                                          // A1
    SheetSet(0, 2, L"=1/0");                                       // C1
    SheetSet(1, 0, L"=\"\"");                                      // A2: a formula whose result is ""
    SheetSet(2, 0, L"0");                                          // A3
    ASSERT_NEAR(Eval(L"=ISERROR(C1)"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=ISERROR(A1)"), 0.0, 1e-12);
    ASSERT_NEAR(Eval(L"=ISERROR(VALUE(\"x\"))"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=ISERROR(FOO())"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=IF(ISERROR(C1),\"bad\",C1)") , 0.0, 1e-12);   // text "bad" reads as 0
    ASSERT_TRUE(IsText(L"IF(ISERROR(C1),\"bad\",C1)", L"bad"));
    ASSERT_TRUE(IsError(L"ISERROR(1 2)", CFormulator::kErrorSyntax));

    ASSERT_NEAR(Eval(L"=ISBLANK(Z1)"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=ISBLANK(A1)"), 0.0, 1e-12);
    ASSERT_NEAR(Eval(L"=ISBLANK(A2)"), 0.0, 1e-12);                // "" from a formula is not blank
    ASSERT_NEAR(Eval(L"=ISBLANK(A3)"), 0.0, 1e-12);                // 0 is not blank
    ASSERT_NEAR(Eval(L"=ISBLANK(\"\")"), 0.0, 1e-12);
    ASSERT_NEAR(Eval(L"=ISBLANK(A4097)"), 1.0, 1e-12);             // off the grid: nothing there
    ASSERT_NEAR(Eval(L"=ISBLANK(1/0)"), 0.0, 1e-12);               // not an error either
    ASSERT_TRUE(IsText(L"IF(ISBLANK(Z5),\"empty\",\"full\")", L"empty"));
    ASSERT_TRUE(IsError(L"ISBLANK()", CFormulator::kErrorSyntax));
    // A cleared cell keeps its old type but is blank.
    SheetSet(3, 0, L"7");
    SheetCell(3, 0)->m_wsText.clear();
    ASSERT_NEAR(Eval(L"=ISBLANK(A4)"), 1.0, 1e-12);
}

// ---- #11: text where a number is needed -------------------------------------------
TEST(TextInArithmetic_IsANumberOrValueError)
{
    SheetReset();
    SheetSet(0, 0, L"hello");               // A1
    SheetSetText(1, 0, L"42");              // A2: text "42"
    SheetSet(2, 0, L"=\"\"");               // A3: a formula whose result is ""
    // Text that reads as a number (or a date) is that number, as in Excel.
    ASSERT_NEAR(Eval(L"=\"42\"+1"), 43.0, 1e-12);
    ASSERT_NEAR(Eval(L"=\"$1,234.50\"*2"), 2469.0, 1e-12);
    ASSERT_NEAR(Eval(L"=\"25%\"*4"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=\"2024-01-15\"+1"), 45307.0, 1e-12);  // date text, as in Excel
    ASSERT_NEAR(Eval(L"=A2+1"), 43.0, 1e-12);
    ASSERT_NEAR(Eval(L"=-A2"), -42.0, 1e-12);
    // Other text is #VALUE!. It read as 0.
    ASSERT_TRUE(IsError(L"\"abc\"+1", CFormulator::kErrorValue));   // was 1
    ASSERT_TRUE(IsError(L"1+\"abc\"", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"A1*2", CFormulator::kErrorValue));        // was 0
    ASSERT_TRUE(IsError(L"A1/2", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"-A1", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"A3+1", CFormulator::kErrorValue));        // "" is no number
    ASSERT_TRUE(IsError(L"\"\"+1", CFormulator::kErrorValue));
    // A blank cell is 0; unary plus and comparisons are not arithmetic.
    ASSERT_NEAR(Eval(L"=Z9+1"), 1.0, 1e-12);
    ASSERT_TRUE(IsText(L"+A1", L"hello"));
    ASSERT_NEAR(Eval(L"=A1=\"hello\""), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=A1>5"), 1.0, 1e-12);
    // The error travels like any other, and IFERROR catches it.
    SheetSet(0, 1, L"=A1+1");               // B1
    ASSERT_TRUE(IsError(L"B1*2", CFormulator::kErrorValue));
    ASSERT_NEAR(Eval(L"=IFERROR(A1+1,0)"), 0.0, 1e-12);
}

TEST(NumericFunctions_ReadTextArguments)
{
    SheetReset();
    SheetSet(0, 0, L"hello");               // A1
    ASSERT_NEAR(Eval(L"=ABS(\"-5\")"), 5.0, 1e-12);
    ASSERT_NEAR(Eval(L"=ROUND(\"3.14159\",2)"), 3.14, 1e-12);
    ASSERT_NEAR(Eval(L"=DATE(\"2024\",\"1\",\"15\")"), 45306.0, 1e-12);
    ASSERT_TRUE(IsError(L"ABS(\"abc\")", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"SQRT(A1)", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"ROUND(1.5,\"x\")", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"MOD(A1,2)", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"SIN(A1)", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"POWER(2,A1)", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"LEFT(\"abc\",\"x\")", CFormulator::kErrorValue));
    // Given as arguments to an aggregate, text counts when it is a number.
    ASSERT_NEAR(Eval(L"=SUM(\"5\",1)"), 6.0, 1e-12);
    ASSERT_NEAR(Eval(L"=SUM(\"2024-01-15\")"), 45306.0, 1e-12);
    ASSERT_TRUE(IsError(L"SUM(\"abc\",1)", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"AVERAGE(\"x\")", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"MAX(A1+0)", CFormulator::kErrorValue));
    ASSERT_NEAR(Eval(L"=COUNT(\"abc\",1,\"5\")"), 2.0, 1e-12);    // COUNT just does not count it
    // In a range, text is skipped as before.
    ASSERT_NEAR(Eval(L"=SUM(A1:A3,1)"), 1.0, 1e-12);
}

TEST(Conditions_ReadTrueFalseTextAndRefuseOtherText)
{
    SheetReset();
    SheetSet(0, 0, L"hello");               // A1
    SheetSetText(1, 0, L"TRUE");            // A2: the text "TRUE"
    ASSERT_NEAR(Eval(L"=IF(\"TRUE\",1,2)"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=IF(\"false\",1,2)"), 2.0, 1e-12);
    ASSERT_NEAR(Eval(L"=IF(A2,1,2)"), 1.0, 1e-12);
    ASSERT_NEAR(Eval(L"=IF(\"1\",1,2)"), 1.0, 1e-12);               // numeric text, as elsewhere
    ASSERT_NEAR(Eval(L"=IF(Z9,1,2)"), 2.0, 1e-12);                  // blank is false
    ASSERT_TRUE(IsError(L"IF(\"abc\",1,2)", CFormulator::kErrorValue));   // took the false branch
    ASSERT_TRUE(IsError(L"IF(A1,1,2)", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"IFS(A1,1)", CFormulator::kErrorValue));
    ASSERT_NEAR(Eval(L"=NOT(\"FALSE\")"), 1.0, 1e-12);
    ASSERT_TRUE(IsError(L"NOT(\"abc\")", CFormulator::kErrorValue));
    // An error in the condition is the answer.
    ASSERT_TRUE(IsError(L"IF(1/0,1,2)", CFormulator::kErrorDiv0));
}

TEST(IfAndIfs_ArgumentCountsAndNoMatch)
{
    StandardSheet();
    ASSERT_TRUE(IsError(L"IF(1)", CFormulator::kErrorSyntax));           // read as 0
    ASSERT_TRUE(IsError(L"IF(1,2,3,4)", CFormulator::kErrorSyntax));     // the 4th was ignored
    ASSERT_NEAR(Eval(L"=IF(0,1)"), 0.0, 1e-12);                          // no else: 0 (FALSE)
    ASSERT_TRUE(IsError(L"IFS()", CFormulator::kErrorSyntax));
    ASSERT_TRUE(IsError(L"IFS(1)", CFormulator::kErrorSyntax));
    ASSERT_TRUE(IsError(L"IFS(0,1,1)", CFormulator::kErrorSyntax));      // a condition with no value
    ASSERT_TRUE(IsError(L"NOT()", CFormulator::kErrorSyntax));
    // No condition true: #N/A, as in Excel. It read 0.
    ASSERT_TRUE(IsError(L"IFS(0,1)", CFormulator::kErrorNA));
    ASSERT_TRUE(IsError(L"IFS(A1>10,1,A1>9,2)", CFormulator::kErrorNA));
    ASSERT_NEAR(Eval(L"=IFS(A1>10,1,A1>3,2)"), 2.0, 1e-12);
    // #N/A is an error like the others.
    ASSERT_TRUE(IsText(L"IFERROR(IFS(0,1),\"none\")", L"none"));
    ASSERT_NEAR(Eval(L"=ISERROR(IFS(0,1))"), 1.0, 1e-12);
    SheetSet(0, 2, L"=IFS(0,1)");                                        // C1
    ASSERT_TRUE(IsError(L"C1+1", CFormulator::kErrorNA));
}

TEST(TextFunctions_RefuseNegativeCountsAndPositions)
{
    StandardSheet();
    ASSERT_TRUE(IsError(L"LEFT(\"abc\",-1)", CFormulator::kErrorValue));     // was ""
    ASSERT_TRUE(IsError(L"RIGHT(\"abc\",-1)", CFormulator::kErrorValue));
    ASSERT_TRUE(IsError(L"MID(\"abcdef\",0,2)", CFormulator::kErrorValue));  // was "ab"
    ASSERT_TRUE(IsError(L"MID(\"abcdef\",2,-1)", CFormulator::kErrorValue)); // was ""
    ASSERT_TRUE(IsError(L"MID(\"abcdef\",2)", CFormulator::kErrorSyntax));   // MID takes three
    ASSERT_TRUE(IsText(L"LEFT(\"abc\",0)", L""));
    ASSERT_TRUE(IsText(L"MID(\"abcdef\",2,0)", L""));
    ASSERT_TRUE(IsText(L"MID(\"abcdef\",2,3)", L"bcd"));
    ASSERT_TRUE(IsText(L"LEFT(\"abc\",\"2\")", L"ab"));                     // numeric text counts
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
    : m_hWndGrid(nullptr), m_hWndEdit(nullptr),
      nColHeaderHeight(0), nRowHeaderWidth(0),
      m_bRedraw(TRUE), m_bResizable(FALSE), m_bUndoRecordEnabled(FALSE),
      m_bDeferRecalc(FALSE), m_bSuppressChangeNotify(FALSE), dwError(0),
      m_bSelecting(FALSE), m_bSizing(FALSE), m_editWndProc(nullptr),
      m_nMouseHoverDelay(0), m_npHoverDelaySet(0),
      m_gridHitTest(0), m_nToBeSized(0), m_resizeStartSize(0),
      m_bEditing(false), m_editCell{ 0, 0 }, m_hEditFont(nullptr),
      m_nSizingLine(0), m_rgbSizingLine(0),
      m_hDefaultFont(nullptr), m_pBatch(nullptr), m_bMergesDirty(true), m_cfGridCells(0),
      m_gdiplusToken(0)
{
    memset(&gcs, 0, sizeof(gcs));
    memset(&m_selectionRect, 0, sizeof(m_selectionRect));
    m_currentCell = m_visibleTopLeft = m_HoverCell = { 0, 0 };
    m_mouseDraggingStartPoint = { 0, 0 };
    m_clientRect = m_paintRect = { 0, 0, 0, 0 };
}

CGrid32Mgr::~CGrid32Mgr() {}

PGRIDCELL CGrid32Mgr::GetCell(UINT nRow, UINT nCol)
{
    auto it = mapCells.find(std::make_pair(nRow, nCol));
    return (it != mapCells.end()) ? it->second : nullptr;
}
