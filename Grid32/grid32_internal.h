#pragma once
#include "Grid32Mgr.h"
#include <cwctype>
#include <cmath>
#include <cstdlib>

bool InitializeGridLibrary(HMODULE hModule);
bool UnInitializeGridLibrary(HMODULE hModule);

// ---- Cell type inference (header-defined so tests can call without
//      linking against the manager impl) -----------------------------------
namespace Grid32Detail {

// Convert calendar Y/M/D to Excel-style serial (days since 1899-12-30).
inline double DateToSerial(int y, int m, int d)
{
    if (m <= 2) { y -= 1; m += 12; }
    int a = y / 100;
    int b = 2 - a + a / 4;
    long long jdn = (long long)(365.25 * (y + 4716)) + (long long)(30.6001 * (m + 1)) + d + b - 1524;
    return (double)(jdn - 2415019);
}

// Inverse of DateToSerial. Valid for serial >= 1.
inline void SerialToYMD(double serial, int& y, int& m, int& d)
{
    long long jdn = (long long)serial + 2415019;
    long long a = jdn + 32044;
    long long b = (4 * a + 3) / 146097;
    long long c = a - (146097 * b) / 4;
    long long dd = (4 * c + 3) / 1461;
    long long e = c - (1461 * dd) / 4;
    long long mm = (5 * e + 2) / 153;
    d = (int)(e - (153 * mm + 2) / 5 + 1);
    m = (int)(mm + 3 - 12 * (mm / 10));
    y = (int)(100 * b + dd - 4800 + mm / 10);
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

    auto formatDate = [](double serial, bool iso) -> std::wstring {
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

