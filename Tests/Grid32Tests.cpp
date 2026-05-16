// Grid32 test harness — standalone console exe.
// Build: see Tests/run_tests.bat.
//
// Initial coverage focuses on the pure helpers added during the typing
// refactor (Grid32Detail::* in grid32_internal.h). Later passes should add
// fixture-driven tests against CGrid32Mgr through a message-only host
// window — that requires more bootstrapping and is intentionally deferred.

#define _CRT_SECURE_NO_WARNINGS
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <string>
#include <vector>
#include <iostream>
#include <functional>
#include <cmath>
#include <climits>

// Pull in the helpers under test. The header includes Grid32Mgr.h which
// pulls in <map>, <vector>, <stack> — fine for a test exe.
#include "../Grid32/grid32_internal.h"

// ---- tiny test harness -----------------------------------------------------
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
#define ASSERT_EQ(a, b) do { if (!((a) == (b))) FAIL_HERE("expected " #a " == " #b); } while (0)
#define ASSERT_NEAR(a, b, eps) do { double _da = (a), _db = (b); if (std::fabs(_da - _db) > (eps)) FAIL_HERE("expected " #a " near " #b); } while (0)

using namespace Grid32Detail;

// ---- DateToSerial ----------------------------------------------------------
// Excel reference: 1899-12-30 = 0, 1900-01-01 = 2 (Excel quirk), 1900-03-01 = 61,
// 2026-05-15 = 46157. Our impl doesn't carry the 1900-leap-year-quirk so
// values past Feb 1900 differ by 1 from Excel.
TEST(DateToSerial_KnownValues)
{
    // 1899-12-31 = 1 day after epoch in our calendar
    ASSERT_EQ((int)DateToSerial(1899, 12, 31), 1);
    // 1900-01-01 = 2 in our calendar
    ASSERT_EQ((int)DateToSerial(1900, 1, 1), 2);
    // 2000-01-01 known to be 36526 (Excel says 36526)
    ASSERT_EQ((int)DateToSerial(2000, 1, 1), 36526);
}

// ---- TryParseDate ----------------------------------------------------------
TEST(TryParseDate_ISO)
{
    double s = 0;
    ASSERT_TRUE(TryParseDate(L"2026-05-15", s));
    ASSERT_TRUE(s > 46000 && s < 47000);
}

TEST(TryParseDate_US)
{
    double s = 0;
    ASSERT_TRUE(TryParseDate(L"5/15/2026", s));
    double sIso = 0;
    ASSERT_TRUE(TryParseDate(L"2026-05-15", sIso));
    ASSERT_NEAR(s, sIso, 0.001);
}

TEST(TryParseDate_TwoDigitYear)
{
    double s = 0;
    ASSERT_TRUE(TryParseDate(L"5/15/26", s));   // -> 2026
    double sFull = 0;
    ASSERT_TRUE(TryParseDate(L"5/15/2026", sFull));
    ASSERT_NEAR(s, sFull, 0.001);
}

TEST(TryParseDate_LeapYearFeb29Valid)
{
    double s = 0;
    ASSERT_TRUE(TryParseDate(L"2024-02-29", s));   // 2024 is leap
}

TEST(TryParseDate_NonLeapYearFeb29Invalid)
{
    double s = 0;
    ASSERT_FALSE(TryParseDate(L"2023-02-29", s));  // 2023 not leap
}

TEST(TryParseDate_RejectsGarbage)
{
    double s = 0;
    ASSERT_FALSE(TryParseDate(L"hello", s));
    ASSERT_FALSE(TryParseDate(L"", s));
    ASSERT_FALSE(TryParseDate(L"2026-13-01", s));     // month 13
    ASSERT_FALSE(TryParseDate(L"2026-02-30", s));     // bad day for Feb
    ASSERT_FALSE(TryParseDate(L"1/2/3/4", s));        // too many slashes
}

// ---- TryParseNumber --------------------------------------------------------
TEST(TryParseNumber_Integers)
{
    double v = 0;
    ASSERT_TRUE(TryParseNumber(L"42", v));      ASSERT_NEAR(v, 42.0, 1e-9);
    ASSERT_TRUE(TryParseNumber(L"-7", v));      ASSERT_NEAR(v, -7.0, 1e-9);
    ASSERT_TRUE(TryParseNumber(L"0", v));       ASSERT_NEAR(v, 0.0, 1e-9);
}

TEST(TryParseNumber_Decimals)
{
    double v = 0;
    ASSERT_TRUE(TryParseNumber(L"3.14", v));    ASSERT_NEAR(v, 3.14, 1e-9);
    ASSERT_TRUE(TryParseNumber(L"-0.5", v));    ASSERT_NEAR(v, -0.5, 1e-9);
}

TEST(TryParseNumber_Scientific)
{
    double v = 0;
    ASSERT_TRUE(TryParseNumber(L"1e3", v));     ASSERT_NEAR(v, 1000.0, 1e-9);
    ASSERT_TRUE(TryParseNumber(L"6.022e23", v)); ASSERT_NEAR(v, 6.022e23, 1e15);
}

TEST(TryParseNumber_RejectsTrailingJunk)
{
    double v = 0;
    ASSERT_FALSE(TryParseNumber(L"42x", v));    // stod consumes "42", we reject
    ASSERT_FALSE(TryParseNumber(L"42 ", v));    // trailing space
    ASSERT_FALSE(TryParseNumber(L"", v));
    ASSERT_FALSE(TryParseNumber(L"abc", v));
}

// ---- IEquals ---------------------------------------------------------------
TEST(IEquals_CaseInsensitive)
{
    ASSERT_TRUE(IEquals(L"TRUE", L"true"));
    ASSERT_TRUE(IEquals(L"True", L"TRUE"));
    ASSERT_TRUE(IEquals(L"false", L"FALSE"));
    ASSERT_FALSE(IEquals(L"truth", L"TRUE"));
    ASSERT_FALSE(IEquals(L"TR", L"TRUE"));
    ASSERT_FALSE(IEquals(L"", L"TRUE"));
}

// ---- InferType -------------------------------------------------------------
TEST(InferType_EmptyIsText)
{
    double v = 0;
    ASSERT_EQ((int)InferType(L"", v), (int)CT_Text);
}

TEST(InferType_NumberLiterals)
{
    double v = 0;
    ASSERT_EQ((int)InferType(L"42", v), (int)CT_Number);
    ASSERT_NEAR(v, 42.0, 1e-9);
    ASSERT_EQ((int)InferType(L"-3.14", v), (int)CT_Number);
    ASSERT_NEAR(v, -3.14, 1e-9);
}

TEST(InferType_Dates)
{
    double v = 0;
    ASSERT_EQ((int)InferType(L"2026-05-15", v), (int)CT_Date);
    ASSERT_TRUE(v > 46000 && v < 47000);
    v = 0;
    ASSERT_EQ((int)InferType(L"5/15/2026", v), (int)CT_Date);
    ASSERT_TRUE(v > 46000 && v < 47000);
}

TEST(InferType_Booleans)
{
    double v = 0;
    ASSERT_EQ((int)InferType(L"TRUE", v), (int)CT_Boolean);
    ASSERT_NEAR(v, 1.0, 1e-9);
    v = 0;
    ASSERT_EQ((int)InferType(L"false", v), (int)CT_Boolean);
    ASSERT_NEAR(v, 0.0, 1e-9);
}

TEST(InferType_TextFallback)
{
    double v = 0;
    ASSERT_EQ((int)InferType(L"hello world", v), (int)CT_Text);
    ASSERT_EQ((int)InferType(L"42 widgets", v), (int)CT_Text);  // trailing text
    ASSERT_EQ((int)InferType(L"Maybe", v), (int)CT_Text);       // not bool keyword
}

// Specifically verifies the audit bug class: a "date-shaped" string with
// invalid components should NOT be classified as a date.
TEST(InferType_NotDateButLooksLikeOne)
{
    double v = 0;
    // 2026-13-01 fails month check -> falls through to Number? "2026-13-01"
    // isn't a valid number either, so should be Text.
    ASSERT_EQ((int)InferType(L"2026-13-01", v), (int)CT_Text);
}

// ---- FormatWithThousands ---------------------------------------------------
TEST(FormatWithThousands_Basic)
{
    ASSERT_TRUE(FormatWithThousands(0) == L"0");
    ASSERT_TRUE(FormatWithThousands(1) == L"1");
    ASSERT_TRUE(FormatWithThousands(999) == L"999");
    ASSERT_TRUE(FormatWithThousands(1000) == L"1,000");
    ASSERT_TRUE(FormatWithThousands(1234567) == L"1,234,567");
    ASSERT_TRUE(FormatWithThousands(-1234) == L"-1,234");
    // LLONG_MIN edge: -(LLONG_MIN) overflows; our impl guards via unsigned math.
    ASSERT_TRUE(FormatWithThousands(LLONG_MIN).front() == L'-');
}

// ---- FormatFixed -----------------------------------------------------------
TEST(FormatFixed_NoDecimals)
{
    ASSERT_TRUE(FormatFixed(42.0, 0, false) == L"42");
    ASSERT_TRUE(FormatFixed(42.7, 0, false) == L"43");      // half-away-from-zero
    ASSERT_TRUE(FormatFixed(-42.5, 0, false) == L"-43");
    ASSERT_TRUE(FormatFixed(0.0, 0, false) == L"0");
}

TEST(FormatFixed_TwoDecimals)
{
    ASSERT_TRUE(FormatFixed(3.14, 2, false) == L"3.14");
    ASSERT_TRUE(FormatFixed(3.1, 2, false) == L"3.10");     // trailing zero kept
    ASSERT_TRUE(FormatFixed(3.0, 2, false) == L"3.00");
    ASSERT_TRUE(FormatFixed(-3.14, 2, false) == L"-3.14");
}

TEST(FormatFixed_Thousands)
{
    ASSERT_TRUE(FormatFixed(1234.5, 2, true) == L"1,234.50");
    ASSERT_TRUE(FormatFixed(1234567.0, 0, true) == L"1,234,567");
}

// ---- FormatCellValue -------------------------------------------------------
TEST(FormatCellValue_TextAlwaysRaw)
{
    // Text cells ignore format — raw text wins regardless.
    ASSERT_TRUE(FormatCellValue(CT_Text, 42.0, L"hello", FMT_CURRENCY) == L"hello");
    ASSERT_TRUE(FormatCellValue(CT_Text, 0.0, L"world", FMT_NUMBER) == L"world");
}

TEST(FormatCellValue_GeneralReturnsRaw)
{
    // FMT_GENERAL preserves the user's typed input even for typed cells.
    ASSERT_TRUE(FormatCellValue(CT_Number, 3.14, L"3.14", FMT_GENERAL) == L"3.14");
    ASSERT_TRUE(FormatCellValue(CT_Date,   46157.0, L"2026-05-15", FMT_GENERAL) == L"2026-05-15");
}

TEST(FormatCellValue_Number)
{
    ASSERT_TRUE(FormatCellValue(CT_Number, 3.14, L"3.14", FMT_NUMBER) == L"3.14");
    ASSERT_TRUE(FormatCellValue(CT_Number, 1234.5, L"1234.5", FMT_THOUSANDS) == L"1,234.50");
}

TEST(FormatCellValue_Integer)
{
    ASSERT_TRUE(FormatCellValue(CT_Number, 3.14, L"3.14", FMT_INTEGER) == L"3");
    ASSERT_TRUE(FormatCellValue(CT_Number, 3.7, L"3.7", FMT_INTEGER) == L"4");
}

TEST(FormatCellValue_Currency)
{
    ASSERT_TRUE(FormatCellValue(CT_Number, 1234.5, L"", FMT_CURRENCY) == L"$1,234.50");
    ASSERT_TRUE(FormatCellValue(CT_Number, -42.0, L"", FMT_CURRENCY) == L"-$42.00");
    ASSERT_TRUE(FormatCellValue(CT_Number, 0.0, L"", FMT_CURRENCY) == L"$0.00");
}

TEST(FormatCellValue_Percent)
{
    // 0.5 -> "50.00%". The format multiplies by 100.
    ASSERT_TRUE(FormatCellValue(CT_Number, 0.5, L"", FMT_PERCENT) == L"50.00%");
    ASSERT_TRUE(FormatCellValue(CT_Number, 1.0, L"", FMT_PERCENT) == L"100.00%");
}

TEST(FormatCellValue_DateISO)
{
    double s = DateToSerial(2026, 5, 15);
    ASSERT_TRUE(FormatCellValue(CT_Date, s, L"", FMT_DATE_ISO) == L"2026-05-15");
}

TEST(FormatCellValue_DateUS)
{
    double s = DateToSerial(2026, 5, 15);
    ASSERT_TRUE(FormatCellValue(CT_Date, s, L"", FMT_DATE_US) == L"5/15/2026");
}

TEST(FormatCellValue_Time)
{
    // 0.5 of a day = 12:00:00
    ASSERT_TRUE(FormatCellValue(CT_Number, 0.5, L"", FMT_TIME) == L"12:00:00");
    // 0.25 = 06:00:00
    ASSERT_TRUE(FormatCellValue(CT_Number, 0.25, L"", FMT_TIME) == L"06:00:00");
}

// ---- FormatEditValue -------------------------------------------------------
TEST(FormatEditValue_BooleanRoundTrip)
{
    ASSERT_TRUE(FormatEditValue(CT_Boolean, 1.0, L"TRUE") == L"TRUE");
    ASSERT_TRUE(FormatEditValue(CT_Boolean, 0.0, L"FALSE") == L"FALSE");
}

TEST(FormatEditValue_DateISO)
{
    double s = DateToSerial(2026, 5, 15);
    ASSERT_TRUE(FormatEditValue(CT_Date, s, L"") == L"2026-05-15");
}

TEST(FormatEditValue_NumberNoTrailingZeros)
{
    // 42 → "42" (decimal point and trailing zeros stripped from to_wstring's %.6f output).
    ASSERT_TRUE(FormatEditValue(CT_Number, 42.0, L"") == L"42");
    ASSERT_TRUE(FormatEditValue(CT_Number, 0.0, L"") == L"0");
    ASSERT_TRUE(FormatEditValue(CT_Number, 3.14, L"") == L"3.14");
}

// ---- SerialToYMD round-trip --------------------------------------------
TEST(SerialToYMD_RoundTrip)
{
    struct { int y, m, d; } cases[] = {
        { 1900, 1, 1 },
        { 1999, 12, 31 },
        { 2000, 1, 1 },
        { 2024, 2, 29 },     // leap day
        { 2026, 5, 15 },
        { 2100, 7, 4 },
    };
    for (auto& tc : cases) {
        double s = DateToSerial(tc.y, tc.m, tc.d);
        int y = 0, m = 0, d = 0;
        SerialToYMD(s, y, m, d);
        if (y != tc.y || m != tc.m || d != tc.d) {
            std::cerr << "  mismatch for " << tc.y << "-" << tc.m << "-" << tc.d
                << " serial=" << s << " got " << y << "-" << m << "-" << d << "\n";
            ++g_failCount;
        }
    }
}

// ---- driver ----------------------------------------------------------------
int main()
{
    int total = (int)Tests().size();
    std::cout << "Running " << total << " Grid32 tests...\n";
    for (auto& t : Tests())
    {
        g_currentTest = t.name;
        int before = g_failCount;
        t.fn();
        if (g_failCount == before)
            std::cout << "  PASS " << t.name << "\n";
    }
    int passed = total - g_failCount;
    std::cout << "\n" << passed << "/" << total << " passed";
    if (g_failCount > 0) std::cout << "  (" << g_failCount << " failed)";
    std::cout << "\n";
    return g_failCount == 0 ? 0 : 1;
}
