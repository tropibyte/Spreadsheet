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
// Fails on NaN too: "> eps" is false for NaN, so a NaN (an error result)
// used to pass every ASSERT_NEAR.
#define ASSERT_NEAR(a, b, eps) do { double _da = (a), _db = (b); if (!(std::fabs(_da - _db) <= (eps))) FAIL_HERE("expected " #a " near " #b); } while (0)

using namespace Grid32Detail;

// ---- DateToSerial ----------------------------------------------------------
// Excel's 1900 date system: 1900-01-01 = 1, and 1900 is treated as a leap year,
// so 60 is "1900-02-29" and 1900-03-01 = 61; from there a serial is the days
// since 1899-12-30. 0 is "1900-01-00". (This test used to pin 1900-01-01 = 2,
// under a comment calling that the Excel quirk; it is not what Excel does.)
TEST(DateToSerial_KnownValues)
{
    ASSERT_EQ((int)DateToSerial(1900, 1, 1), 1);
    ASSERT_EQ((int)DateToSerial(1900, 1, 31), 31);
    ASSERT_EQ((int)DateToSerial(1900, 2, 1), 32);
    ASSERT_EQ((int)DateToSerial(1900, 2, 28), 59);
    ASSERT_EQ((int)DateToSerial(1900, 2, 29), 60);      // Excel's phantom day
    ASSERT_EQ((int)DateToSerial(1900, 3, 1), 61);
    ASSERT_EQ((int)DateToSerial(1900, 3, 0), 60);       // the day before March 1st
    ASSERT_EQ((int)DateToSerial(1900, 1, 0), 0);        // "1900-01-00"
    ASSERT_EQ((int)DateToSerial(2000, 1, 1), 36526);
    ASSERT_EQ((int)DateToSerial(2026, 5, 15), 46157);
    ASSERT_EQ((int)DateToSerial(9999, 12, 31), 2958465);
}

TEST(DateToSerial_RollsMonthsAndDaysOver)
{
    // Months outside 1..12 roll into the years around them. The old formula
    // only got months 0..14 right: (2024,15,1) came out as 2025-03-03.
    ASSERT_EQ((int)DateToSerial(2024, 13, 1), (int)DateToSerial(2025, 1, 1));
    ASSERT_EQ((int)DateToSerial(2024, 15, 1), (int)DateToSerial(2025, 3, 1));
    ASSERT_EQ((int)DateToSerial(2024, 25, 1), (int)DateToSerial(2026, 1, 1));
    ASSERT_EQ((int)DateToSerial(2024, 0, 1), (int)DateToSerial(2023, 12, 1));
    ASSERT_EQ((int)DateToSerial(2024, -10, 1), (int)DateToSerial(2023, 2, 1));
    ASSERT_EQ((int)DateToSerial(2024, -11, 1), (int)DateToSerial(2023, 1, 1));
    ASSERT_EQ((int)DateToSerial(2024, -12, 1), (int)DateToSerial(2022, 12, 1));
    ASSERT_EQ((int)DateToSerial(2024, 1, 0), (int)DateToSerial(2023, 12, 31));
    ASSERT_EQ((int)DateToSerial(2024, 3, 0), (int)DateToSerial(2024, 2, 29));
    ASSERT_EQ((int)DateToSerial(2023, 3, 0), (int)DateToSerial(2023, 2, 28));
    ASSERT_EQ((int)DateToSerial(2024, 1, 366), (int)DateToSerial(2024, 12, 31));
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
        { 1900, 1, 0 },      // serial 0, as Excel shows it
        { 1900, 2, 28 },
        { 1900, 2, 29 },     // serial 60, Excel's phantom day
        { 1900, 3, 1 },
        { 1999, 12, 31 },
        { 2000, 2, 29 },     // a leap century
        { 2100, 2, 28 },     // not a leap century
        { 9999, 12, 31 },
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

TEST(SerialToYMD_EveryDayRoundTrips)
{
    // Every serial Excel can hold maps to a day that maps back to it, and
    // consecutive serials are consecutive days.
    int py = 1900, pm = 1, pd = 0;
    for (int s = 1; s <= 2958465; ++s)
    {
        int y = 0, m = 0, d = 0;
        SerialToYMD((double)s, y, m, d);
        if ((int)DateToSerial(y, m, d) != s) { FAIL_HERE("no round trip at serial " << s); }
        const bool next = (y == py && m == pm && d == pd + 1) ||
                          (y == py && m == pm + 1 && d == 1) ||
                          (y == py + 1 && m == 1 && pm == 12 && d == 1);
        if (!next) { FAIL_HERE("serial " << s << " is not the day after " << py << "-" << pm << "-" << pd); }
        py = y; pm = m; pd = d;
    }
    // The time of day does not change the date.
    int y = 0, m = 0, d = 0;
    SerialToYMD(45306.999, y, m, d);
    ASSERT_TRUE(y == 2024 && m == 1 && d == 15);
}

TEST(FormatCellValue_DatesOutsideTheRangeKeepTheirText)
{
    // No date before 1900-01-00 or after 9999-12-31: the value's text stays.
    ASSERT_TRUE(FormatCellValue(CT_Number, -5.0, L"-5", FMT_DATE_ISO) == L"-5");
    ASSERT_TRUE(FormatCellValue(CT_Number, 3e6, L"3000000", FMT_DATE_US) == L"3000000");
    ASSERT_TRUE(FormatCellValue(CT_Number, 1.0, L"1", FMT_DATE_ISO) == L"1900-01-01");
    ASSERT_TRUE(FormatCellValue(CT_Number, 60.0, L"60", FMT_DATE_ISO) == L"1900-02-29");
    ASSERT_TRUE(FormatCellValue(CT_Number, 0.0, L"0", FMT_DATE_ISO) == L"1900-01-00");
    ASSERT_TRUE(FormatCellValue(CT_Number, 2958465.0, L"2958465", FMT_DATE_ISO) == L"9999-12-31");
}

TEST(FormatCellValue_LogicalsIgnoreNumberFormats)
{
    ASSERT_TRUE(FormatCellValue(CT_Boolean, 1.0, L"TRUE", FMT_NUMBER) == L"TRUE");     // was "1.00"
    ASSERT_TRUE(FormatCellValue(CT_Boolean, 0.0, L"FALSE", FMT_CURRENCY) == L"FALSE"); // was "$0.00"
    ASSERT_TRUE(FormatCellValue(CT_Boolean, 1.0, L"true", FMT_PERCENT) == L"TRUE");
    ASSERT_TRUE(FormatCellValue(CT_Boolean, 1.0, L"true", FMT_GENERAL) == L"true");    // General shows what was typed
}

// ---- Case mapping (grid32_case.h) -------------------------------------------
TEST(CaseTable_MapsLettersAndNothingElse)
{
    // ASCII, by the fast path.
    for (wchar_t c = L'a'; c <= L'z'; ++c)
    {
        ASSERT_TRUE(ToUpperChar(c) == c - 32);
        ASSERT_TRUE(ToLowerChar((wchar_t)(c - 32)) == c);
    }
    ASSERT_TRUE(ToUpperChar(L'1') == L'1' && ToUpperChar(L'@') == L'@' && ToLowerChar(L'[') == L'[');
    // Runs of each shape: a block (à..þ, skipping ÷), alternating pairs
    // (Latin Extended-A), single far mappings (ÿ, µ, ı), Greek, Cyrillic.
    ASSERT_TRUE(ToUpperChar(0x00E0) == 0x00C0 && ToUpperChar(0x00FE) == 0x00DE);
    ASSERT_TRUE(ToUpperChar(0x00F7) == 0x00F7);          // ÷ is no letter
    ASSERT_TRUE(ToUpperChar(0x0101) == 0x0100 && ToUpperChar(0x0100) == 0x0100);
    ASSERT_TRUE(ToLowerChar(0x0100) == 0x0101 && ToLowerChar(0x0101) == 0x0101);
    ASSERT_TRUE(ToUpperChar(0x00FF) == 0x0178 && ToLowerChar(0x0178) == 0x00FF);
    ASSERT_TRUE(ToUpperChar(0x00B5) == 0x039C);          // µ -> Greek capital mu
    ASSERT_TRUE(ToUpperChar(0x0131) == L'I');            // dotless ı
    ASSERT_TRUE(ToUpperChar(0x03C9) == 0x03A9 && ToLowerChar(0x03A9) == 0x03C9);
    ASSERT_TRUE(ToUpperChar(0x044F) == 0x042F && ToUpperChar(0x0450) == 0x0400);
    ASSERT_TRUE(ToUpperChar(0x00DF) == 0x00DF);          // ß: no one-character capital
    // Not letters, surrogate halves, and the last code points: unchanged.
    ASSERT_TRUE(ToUpperChar(0x20AC) == 0x20AC && ToUpperChar(0x4E2D) == 0x4E2D);
    ASSERT_TRUE(ToUpperChar(0xD801) == 0xD801 && ToLowerChar(0xDC28) == 0xDC28);
    ASSERT_TRUE(ToUpperChar(0xFFFF) == 0xFFFF && ToLowerChar(0x0000) == 0x0000);
    // Every letter's case pair maps back, across the whole BMP: upper(lower(c))
    // is c for every capital whose lowercase is a different letter that maps
    // back to it.
    int checked = 0;
    for (unsigned long c = 0; c < 0x10000; ++c)
    {
        const wchar_t lo = ToLowerChar((wchar_t)c);
        if (lo != (wchar_t)c && ToUpperChar(lo) == (wchar_t)c) ++checked;
    }
    ASSERT_TRUE(checked > 1000);
}

// ---- driver ----------------------------------------------------------------// ---- driver ----------------------------------------------------------------
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
