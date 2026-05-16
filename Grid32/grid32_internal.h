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

} // namespace Grid32Detail

