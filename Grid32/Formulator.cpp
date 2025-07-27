#include "pch.h"
#include "grid32mgr.h"
#include "Formulator.h"
#include <cmath>
#include <limits>


// --- Formula evaluation helpers ---


void CFormulator::SkipSpaces(const std::wstring& s, size_t& pos) {
    while (pos < s.size() && iswspace(s[pos])) ++pos;
}

bool CFormulator::ParseCellRef(const std::wstring& token, UINT& row, UINT& col) {
    size_t i = 0;
    col = 0;
    while (i < token.size() && iswalpha(token[i])) {
        col = col * 26 + (towupper(token[i]) - L'A' + 1);
        ++i;
    }
    if (i == 0 || i >= token.size()) return false;
    try {
        row = (UINT)std::stoul(token.substr(i)) - 1;
    }
    catch (...) {
        return false;
    }
    col -= 1;
    return true;
}

double CFormulator::SumRange(CGrid32Mgr* mgr, const std::wstring& arg,
    std::set<std::pair<UINT, UINT>>& visited)
{
    UINT sr, sc, er, ec;
    if (!ParseCellRef(arg, sr, sc))
    {
        size_t colon = arg.find(L':');
        if (colon == std::wstring::npos)
            return 0.0;
        std::wstring s1 = arg.substr(0, colon);
        std::wstring s2 = arg.substr(colon + 1);
        if (!ParseCellRef(s1, sr, sc) || !ParseCellRef(s2, er, ec))
            return 0.0;
    }
    else
    {
        er = sr; ec = sc;
    }

    double total = 0.0;
    for (UINT r = sr; r <= er; ++r)
        for (UINT c = sc; c <= ec; ++c)
            total += GetCellValue(mgr, r, c, visited);
    return total;
}

double CFormulator::MinMaxRange(CGrid32Mgr* mgr, const std::wstring& arg,
    std::set<std::pair<UINT, UINT>>& visited, bool bMax)
{
    UINT sr, sc, er, ec;
    if (!ParseCellRef(arg, sr, sc))
    {
        size_t colon = arg.find(L':');
        if (colon == std::wstring::npos)
            return 0.0;
        std::wstring s1 = arg.substr(0, colon);
        std::wstring s2 = arg.substr(colon + 1);
        if (!ParseCellRef(s1, sr, sc) || !ParseCellRef(s2, er, ec))
            return 0.0;
    }
    else
    {
        er = sr; ec = sc;
    }

    double best = bMax ? -std::numeric_limits<double>::infinity()
        : std::numeric_limits<double>::infinity();
    for (UINT r = sr; r <= er; ++r)
        for (UINT c = sc; c <= ec; ++c)
        {
            double v = GetCellValue(mgr, r, c, visited);
            if (bMax)
                best = max(best, v);
            else
                best = min(best, v);
        }
    return best;
}

double CFormulator::CountRange(CGrid32Mgr* mgr, const std::wstring& arg,
    std::set<std::pair<UINT, UINT>>& visited)
{
    UINT sr, sc, er, ec;
    if (!ParseCellRef(arg, sr, sc))
    {
        size_t colon = arg.find(L':');
        if (colon == std::wstring::npos)
            return 0.0;
        std::wstring s1 = arg.substr(0, colon);
        std::wstring s2 = arg.substr(colon + 1);
        if (!ParseCellRef(s1, sr, sc) || !ParseCellRef(s2, er, ec))
            return 0.0;
    }
    else
    {
        er = sr; ec = sc;
    }

    double count = 0.0;
    for (UINT r = sr; r <= er; ++r)
        for (UINT c = sc; c <= ec; ++c)
        {
            auto key = std::make_pair(r, c);
            if (visited.count(key))
                continue;
            PGRIDCELL cell = mgr->GetCell(r, c);
            if (!cell)
                continue;
            if (cell->m_bFormula)
            {
                visited.insert(key);
                size_t p = 0;
                ParseExpression(mgr, cell->m_wsFormula, p, visited);
                ++count;
            }
            else if (!cell->m_wsText.empty())
            {
                try { UNREFERENCED_PARAMETER( std::stod(cell->m_wsText)); ++count; }
                catch (...) {}
            }
        }
    return count;
}

double CFormulator::ParseTerm(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos,
    std::set<std::pair<UINT, UINT>>& visited) {
    double value = ParseFactor(mgr, expr, pos, visited);
    while (true) {
        SkipSpaces(expr, pos);
        if (pos >= expr.size()) break;
        wchar_t op = expr[pos];
        if (op != L'*' && op != L'/') break;
        ++pos;
        double rhs = ParseFactor(mgr, expr, pos, visited);
        if (op == L'*') value *= rhs; else if (rhs != 0) value /= rhs;
    }
    return value;
}

double CFormulator::ParseExpression(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos,
    std::set<std::pair<UINT, UINT>>& visited) {
    double value = ParseTerm(mgr, expr, pos, visited);
    while (true) {
        SkipSpaces(expr, pos);
        if (pos >= expr.size()) break;
        wchar_t op = expr[pos];
        if (op != L'+' && op != L'-') break;
        ++pos;
        double rhs = ParseTerm(mgr, expr, pos, visited);
        if (op == L'+') value += rhs; else value -= rhs;
    }
    return value;
}

double CFormulator::ParseFactor(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos,
    std::set<std::pair<UINT, UINT>>& visited) {
    SkipSpaces(expr, pos);
    if (pos >= expr.size()) return 0.0;
    if (expr[pos] == L'(') {
        ++pos;
        double val = ParseExpression(mgr, expr, pos, visited);
        if (pos < expr.size() && expr[pos] == L')') ++pos;
        return val;
    }
    if (expr[pos] == L'+' || expr[pos] == L'-') {
        bool neg = (expr[pos] == L'-');
        ++pos;
        double val = ParseFactor(mgr, expr, pos, visited);
        return neg ? -val : val;
    }

    size_t start = pos;
    while (pos < expr.size() && (iswalnum(expr[pos]) || expr[pos] == L'.')) ++pos;
    std::wstring token = expr.substr(start, pos - start);
    if (token.empty()) return 0.0;

    if (pos < expr.size() && expr[pos] == L'(')
    {
        ++pos; // skip '('
        size_t argStart = pos;
        int depth = 1;
        while (pos < expr.size() && depth > 0)
        {
            if (expr[pos] == L'(') depth++;
            else if (expr[pos] == L')') depth--;
            if (depth > 0) ++pos;
        }
        std::wstring arg = expr.substr(argStart, pos - argStart);
        if (pos < expr.size() && expr[pos] == L')') ++pos;

        double res = 0.0;
        if (_wcsicmp(token.c_str(), L"SUM") == 0)
        {
            res = SumRange(mgr, arg, visited);
        }
        else if (_wcsicmp(token.c_str(), L"AVERAGE") == 0)
        {
            double sum = SumRange(mgr, arg, visited);
            UINT sr, sc, er, ec;
            if (ParseCellRef(arg, sr, sc))
            {
                er = sr; ec = sc;
            }
            else
            {
                size_t colon = arg.find(L':');
                if (colon != std::wstring::npos && ParseCellRef(arg.substr(0, colon), sr, sc) &&
                    ParseCellRef(arg.substr(colon + 1), er, ec))
                {
                }
                else
                {
                    sr = sc = er = ec = 0;
                }
            }
            UINT count = (er - sr + 1) * (ec - sc + 1);
            res = count ? sum / count : 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"MIN") == 0)
        {
            res = MinMaxRange(mgr, arg, visited, false);
        }
        else if (_wcsicmp(token.c_str(), L"MAX") == 0)
        {
            res = MinMaxRange(mgr, arg, visited, true);
        }
        else if (_wcsicmp(token.c_str(), L"COUNT") == 0)
        {
            res = CountRange(mgr, arg, visited);
        }
        else if (_wcsicmp(token.c_str(), L"SIN") == 0)
        {
            size_t p = 0;
            res = sin(ParseExpression(mgr, arg, p, visited));
        }
        else if (_wcsicmp(token.c_str(), L"COS") == 0)
        {
            size_t p = 0;
            res = cos(ParseExpression(mgr, arg, p, visited));
        }
        else if (_wcsicmp(token.c_str(), L"TAN") == 0)
        {
            size_t p = 0;
            res = tan(ParseExpression(mgr, arg, p, visited));
        }
        else if (_wcsicmp(token.c_str(), L"POWER") == 0)
        {
            size_t comma = arg.find(L',');
            double base = 0.0, exp = 0.0;
            if (comma != std::wstring::npos)
            {
                size_t p = 0;
                base = ParseExpression(mgr, arg.substr(0, comma), p, visited);
                p = 0;
                exp = ParseExpression(mgr, arg.substr(comma + 1), p, visited);
            }
            res = pow(base, exp);
        }
        else if (_wcsicmp(token.c_str(), L"ROOT") == 0)
        {
            size_t comma = arg.find(L',');
            double val = 0.0, n = 2.0;
            size_t p = 0;
            if (comma != std::wstring::npos)
            {
                val = ParseExpression(mgr, arg.substr(0, comma), p, visited);
                p = 0;
                n = ParseExpression(mgr, arg.substr(comma + 1), p, visited);
            }
            else
            {
                val = ParseExpression(mgr, arg, p, visited);
            }
            res = n != 0.0 ? pow(val, 1.0 / n) : 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"LOG") == 0)
        {
            size_t p = 0;
            res = log10(ParseExpression(mgr, arg, p, visited));
        }
        else if (_wcsicmp(token.c_str(), L"LN") == 0)
        {
            size_t p = 0;
            res = log(ParseExpression(mgr, arg, p, visited));
        }
        else
        {
            // Unknown function, treat as 0
        }
        return res;
    }
    else
    {
        UINT row = 0, col = 0;
        if (ParseCellRef(token, row, col))
            return GetCellValue(mgr, row, col, visited);

        try {
            return std::stod(token);
        }
        catch (...) {
            return 0.0;
        }
    }
}

double CFormulator::GetCellValue(CGrid32Mgr* mgr, UINT row, UINT col,
    std::set<std::pair<UINT, UINT>>& visited) {
    auto key = std::make_pair(row, col);
    if (visited.count(key)) return 0.0;
    visited.insert(key);
    PGRIDCELL cell = mgr->GetCell(row, col);
    if (!cell) return 0.0;
    if (cell->m_bFormula) {
        size_t p = 0;
        return ParseExpression(mgr, cell->m_wsFormula, p, visited);
    }
    try {
        return std::stod(cell->m_wsText);
    }
    catch (...) {
        return 0.0;
    }
}