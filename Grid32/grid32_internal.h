#pragma once
#include "Grid32Mgr.h"
#include <cwctype>
#include <cmath>
#include <cstdlib>

bool InitializeGridLibrary(HMODULE hModule);
bool UnInitializeGridLibrary(HMODULE hModule);
bool RegisterPhotoGalleryClass(HINSTANCE hInstance);      // photogallery_wndproc.cpp
bool UnregisterPhotoGalleryClass(HINSTANCE hInstance);

// ---- Cell type inference (header-defined so tests can call without
//      linking against the manager impl) -----------------------------------
namespace Grid32Detail {

// ---- Dates -----------------------------------------------------------------
// A date is a serial number of days in Excel's 1900 date system:
//
//   1 = 1900-01-01 ... 59 = 1900-02-28, 60 = 1900-02-29, 61 = 1900-03-01 ...
//   2958465 = 9999-12-31, and 0 = "1900-01-00", the day before 1900-01-01.
//
// Excel treats 1900 as a leap year (for Lotus 1-2-3 compatibility), so 60 is a
// day that never existed, every serial from 61 on is the number of days since
// 1899-12-30, and those before it are one less. These helpers used to count
// from 1899-12-30 throughout, which put 1900-01-01 at 2 and every date before
// 1900-03-01 a day later than Excel does.
//
// The helpers carry on below 0 (-1 = 1899-12-30) so the arithmetic stays
// continuous; what is a date a sheet can hold (Excel's 0..2958465) is for the
// callers to decide, and kMaxDateSerial is the top of it.
constexpr double kMaxDateSerial = 2958465.0;    // 9999-12-31

namespace DateCalc {
    // Days from 1970-01-01 to y-m-d (m 1..12) in the proleptic Gregorian
    // calendar, exactly, for any y (H. Hinnant's days_from_civil).
    inline long long DaysFromCivil(long long y, long long m, long long d)
    {
        if (m <= 2) y -= 1;
        const long long era = (y >= 0 ? y : y - 399) / 400;
        const long long yoe = y - era * 400;
        const long long doy = (153 * (m + (m > 2 ? -3 : 9)) + 2) / 5 + d - 1;
        const long long doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
        return era * 146097 + doe - 719468;
    }

    // The inverse (civil_from_days).
    inline void CivilFromDays(long long z, long long& y, long long& m, long long& d)
    {
        z += 719468;
        const long long era = (z >= 0 ? z : z - 146096) / 146097;
        const long long doe = z - era * 146097;
        const long long yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
        const long long doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
        const long long mp = (5 * doy + 2) / 153;
        d = doy - (153 * mp + 2) / 5 + 1;
        m = mp < 10 ? mp + 3 : mp - 9;
        y = yoe + era * 400 + (m <= 2 ? 1 : 0);
    }

    // 1970-01-01 is serial 25569.
    constexpr long long kUnixEpochSerial = 25569;
}

// Calendar Y/M/D to a serial. A month outside 1..12 rolls into the years
// around it and a day outside the month into the months around it, as Excel's
// DATE does: (2024,13,1) is 2025-01-01, (2024,1,0) is 2023-12-31, and
// (1900,3,0) is Excel's 1900-02-29, 60. (The old formula only got months
// 0..14 right: (2024,15,1) came out two days late.)
inline double DateToSerial(int y, int m, int d)
{
    long long yy = y;
    long long mm = (long long)m - 1;                    // 0-based, then into 0..11
    long long carry = mm >= 0 ? mm / 12 : -((11 - mm) / 12);
    yy += carry;
    mm -= carry * 12;
    long long first = DateCalc::DaysFromCivil(yy, mm + 1, 1) + DateCalc::kUnixEpochSerial;
    if (first < 61) first -= 1;                         // before 1900-03-01: Excel's count
    return (double)(first + (long long)d - 1);
}

// Inverse of DateToSerial for the day `serial` falls on (its fraction, the
// time of day, is ignored). 0 is 1900-01-00 and 60 is 1900-02-29, as Excel
// shows them. Serials too large for any calendar arithmetic (beyond +-1e11
// days) come back as 0-00-00.
inline void SerialToYMD(double serial, int& y, int& m, int& d)
{
    const double day = std::floor(serial);
    if (!(std::fabs(day) <= 1e11)) { y = m = d = 0; return; }
    const long long s = (long long)day;
    if (s == 0)  { y = 1900; m = 1; d = 0;  return; }
    if (s == 60) { y = 1900; m = 2; d = 29; return; }
    long long yy = 0, mm = 0, dd = 0;
    DateCalc::CivilFromDays((s < 60 ? s + 1 : s) - DateCalc::kUnixEpochSerial, yy, mm, dd);
    y = (int)yy; m = (int)mm; d = (int)dd;
}

inline bool TryParseDate(const std::wstring& text, double& serialOut)
{
    if (text.empty()) return false;
    int y = 0, m = 0, d = 0;
    if (text.size() == 10 && (text[4] == L'-' || text[4] == L'/') &&
        text[4] == text[7] && iswdigit(text[0]) && iswdigit(text[5]) && iswdigit(text[8]))
    {
        try {
            y = std::stoi(text.substr(0, 4));
            m = std::stoi(text.substr(5, 2));
            d = std::stoi(text.substr(8, 2));
        } catch (...) { return false; }
    }
    else
    {
        size_t slash1 = text.find(L'/');
        if (slash1 == std::wstring::npos) return false;
        size_t slash2 = text.find(L'/', slash1 + 1);
        if (slash2 == std::wstring::npos) return false;
        if (text.find(L'/', slash2 + 1) != std::wstring::npos) return false;
        try {
            m = std::stoi(text.substr(0, slash1));
            d = std::stoi(text.substr(slash1 + 1, slash2 - slash1 - 1));
            y = std::stoi(text.substr(slash2 + 1));
        } catch (...) { return false; }
        if (y < 100) y += (y < 30 ? 2000 : 1900);
    }
    if (y < 1900 || y > 9999 || m < 1 || m > 12 || d < 1 || d > 31) return false;
    static const int dim[] = { 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
    int maxD = dim[m - 1];
    bool leap = (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);
    if (m == 2 && leap) maxD = 29;
    if (d > maxD) return false;
    serialOut = DateToSerial(y, m, d);
    return true;
}

inline bool TryParseNumber(const std::wstring& text, double& valueOut)
{
    if (text.empty()) return false;
    try {
        size_t consumed = 0;
        double v = std::stod(text, &consumed);
        if (consumed != text.size()) return false;
        valueOut = v;
        return true;
    } catch (...) { return false; }
}

inline bool IEquals(const std::wstring& a, const wchar_t* b)
{
    size_t i = 0;
    while (b[i] != 0 && i < a.size())
    {
        if (towlower(a[i]) != towlower(b[i])) return false;
        ++i;
    }
    return i == a.size() && b[i] == 0;
}

inline CellType InferType(const std::wstring& text, double& valueOut)
{
    valueOut = 0.0;
    if (text.empty()) return CT_Text;
    double n = 0.0;
    if (TryParseDate(text, n)) { valueOut = n; return CT_Date; }
    if (IEquals(text, L"TRUE"))  { valueOut = 1.0; return CT_Boolean; }
    if (IEquals(text, L"FALSE")) { valueOut = 0.0; return CT_Boolean; }
    if (TryParseNumber(text, n)) { valueOut = n; return CT_Number; }
    return CT_Text;
}

// ---- Number formatting --------------------------------------------------
// Render a value with thousands separators. helper: stringify integer with
// commas every three digits.
inline std::wstring FormatWithThousands(long long whole)
{
    std::wstring s;
    bool neg = whole < 0;
    unsigned long long n = neg ? (unsigned long long)(-(whole + 1)) + 1ULL : (unsigned long long)whole;
    if (n == 0) s = L"0";
    int digits = 0;
    while (n > 0)
    {
        if (digits > 0 && digits % 3 == 0) s.insert(s.begin(), L',');
        s.insert(s.begin(), (wchar_t)(L'0' + (n % 10)));
        n /= 10;
        ++digits;
    }
    if (neg) s.insert(s.begin(), L'-');
    return s;
}

// Render `value` with `decimals` fractional digits (rounded half-away-from-zero)
// and optional thousands separators. Returns the formatted body without sign;
// sign is prepended by callers as needed.
inline std::wstring FormatFixed(double value, int decimals, bool useThousands)
{
    bool neg = value < 0;
    double abs = neg ? -value : value;
    double mult = 1.0;
    for (int i = 0; i < decimals; ++i) mult *= 10.0;
    long long scaled = (long long)(abs * mult + 0.5);
    long long whole = scaled / (long long)mult;
    long long frac = scaled - whole * (long long)mult;

    std::wstring out = useThousands ? FormatWithThousands(whole) : std::to_wstring(whole);
    if (decimals > 0)
    {
        out.push_back(L'.');
        std::wstring fracStr = std::to_wstring(frac);
        while ((int)fracStr.size() < decimals) fracStr.insert(fracStr.begin(), L'0');
        out += fracStr;
    }
    if (neg) out.insert(out.begin(), L'-');
    return out;
}

// Format a typed cell value for display. CT_Text cells always return rawText.
// For numeric / date / boolean / formula cells, the format code drives the
// representation and m_dValue carries the canonical scalar.
inline std::wstring FormatCellValue(CellType type, double value,
                                    const std::wstring& rawText, UINT format)
{
    if (type == CT_Text || format == FMT_GENERAL) return rawText;

    auto formatDate = [&rawText](double serial, bool iso) -> std::wstring {
        // Outside Excel's date range (negative, or past 9999-12-31) there is
        // no date to show; keep the value's own text rather than invent one.
        if (!(serial >= 0.0 && serial < kMaxDateSerial + 1.0)) return rawText;
        int y = 0, m = 0, d = 0;
        SerialToYMD(serial, y, m, d);
        wchar_t buf[16];
        if (iso) swprintf_s(buf, 16, L"%04d-%02d-%02d", y, m, d);
        else     swprintf_s(buf, 16, L"%d/%d/%04d", m, d, y);
        return buf;
    };

    switch (format)
    {
    case FMT_NUMBER:    return FormatFixed(value, 2, false);
    case FMT_INTEGER:   return FormatFixed(value, 0, false);
    case FMT_THOUSANDS: return FormatFixed(value, 2, true);
    case FMT_CURRENCY:  {
        std::wstring body = FormatFixed(value < 0 ? -value : value, 2, true);
        return (value < 0) ? (L"-$" + body) : (L"$" + body);
    }
    case FMT_PERCENT:   return FormatFixed(value * 100.0, 2, false) + L"%";
    case FMT_DATE_ISO:  return formatDate(value, true);
    case FMT_DATE_US:   return formatDate(value, false);
    case FMT_TIME: {
        double frac = value - std::floor(value);
        if (frac < 0) frac += 1.0;
        int totalSec = (int)(frac * 86400.0 + 0.5);
        int hh = (totalSec / 3600) % 24;
        int mm = (totalSec / 60) % 60;
        int ss = totalSec % 60;
        wchar_t buf[16];
        swprintf_s(buf, 16, L"%02d:%02d:%02d", hh, mm, ss);
        return buf;
    }
    default: return rawText;
    }
}

// How a formula's date result (CFormulator::Result::isDate) shows in a General
// cell: "2024-01-15", with " 13:45" (or " 13:45:07") when it has a time of
// day. A serial outside Excel's dates has no date to show and keeps `general`,
// its number's text. Other formats render it as they render any number.
inline std::wstring FormatGeneralDate(double serial, const std::wstring& general)
{
    if (!(serial >= 0.0 && serial < kMaxDateSerial + 1.0)) return general;
    double day = std::floor(serial);
    long secs = std::lround((serial - day) * 86400.0);
    if (secs >= 86400) { day += 1.0; secs = 0; }            // 23:59:59.7 rounds to midnight
    if (day > kMaxDateSerial) return general;
    int y = 0, m = 0, d = 0;
    SerialToYMD(day, y, m, d);
    wchar_t buf[40];
    if (secs == 0)
        swprintf_s(buf, 40, L"%04d-%02d-%02d", y, m, d);
    else if (secs % 60 == 0)
        swprintf_s(buf, 40, L"%04d-%02d-%02d %02ld:%02ld", y, m, d, secs / 3600, (secs / 60) % 60);
    else
        swprintf_s(buf, 40, L"%04d-%02d-%02d %02ld:%02ld:%02ld", y, m, d,
                   secs / 3600, (secs / 60) % 60, secs % 60);
    return buf;
}

// Render the value the user should see when re-entering edit mode on a
// formatted typed cell. Numbers come back as their natural form (no trailing
// zeros, no thousands separators, no currency symbol), dates as ISO, booleans
// as TRUE/FALSE. Text cells return their literal text.
inline std::wstring FormatEditValue(CellType type, double value, const std::wstring& rawText)
{
    switch (type)
    {
    case CT_Boolean:
        return value != 0.0 ? L"TRUE" : L"FALSE";
    case CT_Date: {
        int y = 0, m = 0, d = 0;
        SerialToYMD(value, y, m, d);
        wchar_t buf[16];
        swprintf_s(buf, 16, L"%04d-%02d-%02d", y, m, d);
        return buf;
    }
    case CT_Number: {
        // Trim a trailing ".000000" / ".5000" tail from to_wstring's fixed
        // 6-digit output. Integers come back without a dot.
        std::wstring s = std::to_wstring(value);
        if (s.find(L'.') != std::wstring::npos)
        {
            while (!s.empty() && s.back() == L'0') s.pop_back();
            if (!s.empty() && s.back() == L'.') s.pop_back();
        }
        return s;
    }
    default:
        return rawText;
    }
}

} // namespace Grid32Detail

