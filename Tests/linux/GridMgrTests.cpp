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

#include <iostream>
#include <functional>
#include <cmath>
#include <string>

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
#define ASSERT_WSTR(expr, want) do { std::wstring _g = (expr); std::wstring _w = (want); \
    if (_g != _w) { std::wcerr << L"  (got \"" << _g << L"\", want \"" << _w << L"\")\n"; \
    FAIL_HERE("expected " #expr " == " #want); } } while (0)

// ---- fixture ---------------------------------------------------------------
// A grid with no window behind it. Create() succeeds because every Win32 call
// it makes is stubbed.
class Sheet
{
public:
    Sheet(size_t cols = 16, size_t rows = 64)
    {
        GRIDCREATESTRUCT gcs{};
        gcs.cbSize = sizeof(GRIDCREATESTRUCT);
        gcs.nWidth = cols;
        gcs.nHeight = rows;
        gcs.nDefColWidth = 100;
        gcs.nDefRowHeight = 20;
        gcs.style = GS_SPREADSHEET;
        m_ok = m_mgr.Create(&gcs);
    }
    bool ok() const { return m_ok; }
    CGrid32Mgr& mgr() { return m_mgr; }

    void Set(UINT row, UINT col, const wchar_t* text) { m_mgr.SetCellText(row, col, text); }
    std::wstring Text(UINT row, UINT col)
    {
        PGRIDCELL c = m_mgr.GetCell(row, col);
        return c ? c->m_wsText : std::wstring();
    }
    double Value(UINT row, UINT col)
    {
        PGRIDCELL c = m_mgr.GetCell(row, col);
        return c ? c->m_dValue : 0.0;
    }
private:
    CGrid32Mgr m_mgr;
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
    // The import defers per-cell recalculation and runs one pass at the end;
    // a formula placed beforehand must still end up correct.
    Sheet s;
    s.Set(5, 0, L"=A1+B1");
    StreamInCsv(s, L"1,2\n3,4\n");
    ASSERT_WSTR(s.Text(5, 0), L"3");
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
