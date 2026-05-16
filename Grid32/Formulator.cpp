#include "pch.h"
#include "grid32mgr.h"
#include "Formulator.h"
#include "grid32_internal.h"
#include <cmath>
#include <limits>
#include <chrono>
#include <ctime>
#include <algorithm>


// --- Formula evaluation helpers ---

// Text-returning functions stash their result in a thread-local slot and
// return 0.0 as their numeric placeholder. Callers (the top-level
// EvaluateFormula, or another text function consuming this as input) check
// g_textResultValid after parsing and prefer the string when set.
namespace Grid32FormulaText {
    thread_local std::wstring g_textResult;
    thread_local bool g_textResultValid = false;

    inline void SetText(const std::wstring& s) {
        g_textResult = s;
        g_textResultValid = true;
    }
    inline void ResetText() {
        g_textResult.clear();
        g_textResultValid = false;
    }
}

namespace {
    // Cap formula recursion to defend against pathological cycles and deeply
    // nested expressions that the visited-set alone can't bound (e.g.,
    // long alternating chains across multiple cells).
    constexpr int kMaxFormulaDepth = 256;
    thread_local int g_formulaDepth = 0;

    struct FormulaDepthGuard {
        bool overflowed;
        FormulaDepthGuard() : overflowed(false) {
            if (++g_formulaDepth > kMaxFormulaDepth)
                overflowed = true;
        }
        ~FormulaDepthGuard() { --g_formulaDepth; }
    };

    // Split a function's argument string at top-level commas, respecting
    // nested parentheses and quoted string literals. "" inside a quoted run
    // is treated as a literal " and does not exit quote mode (Excel's
    // string-literal escape).
    std::vector<std::wstring> SplitArgs(const std::wstring& arg) {
        std::vector<std::wstring> out;
        int depth = 0;
        bool inQuote = false;
        size_t start = 0;
        for (size_t i = 0; i < arg.size(); ++i) {
            wchar_t c = arg[i];
            if (inQuote) {
                if (c == L'"') {
                    if (i + 1 < arg.size() && arg[i + 1] == L'"') ++i;
                    else inQuote = false;
                }
                continue;
            }
            if (c == L'"') { inQuote = true; }
            else if (c == L'(') ++depth;
            else if (c == L')') { if (depth > 0) --depth; }
            else if (c == L',' && depth == 0) {
                out.push_back(arg.substr(start, i - start));
                start = i + 1;
            }
        }
        out.push_back(arg.substr(start));
        return out;
    }

    inline std::wstring TrimWs(const std::wstring& s) {
        size_t a = 0, b = s.size();
        while (a < b && iswspace(s[a])) ++a;
        while (b > a && iswspace(s[b - 1])) --b;
        return s.substr(a, b - a);
    }

    // Resolve an argument expression to a string. Recognizes quoted literals
    // ("hello", "with ""embedded"" quotes"), cell references, and falls back
    // to numeric evaluation (consulting the text-result TLS in case a
    // nested call produced a string).
    std::wstring EvalAsString(CGrid32Mgr* mgr, const std::wstring& arg,
                              std::set<std::pair<UINT, UINT>>& visited) {
        std::wstring s = TrimWs(arg);
        if (s.empty()) return std::wstring();

        // Quoted literal
        if (s.front() == L'"') {
            std::wstring out;
            for (size_t i = 1; i < s.size(); ++i) {
                if (s[i] == L'"') {
                    if (i + 1 < s.size() && s[i + 1] == L'"') {
                        out.push_back(L'"');
                        ++i;
                    } else break;
                } else {
                    out.push_back(s[i]);
                }
            }
            return out;
        }

        // Bare cell reference (e.g. A1) — resolve to cell text.
        UINT row = 0, col = 0;
        if (mgr != nullptr && CFormulator::ParseCellRef(s, row, col)) {
            PGRIDCELL cell = mgr->GetCell(row, col);
            return cell ? cell->m_wsText : std::wstring();
        }

        // Generic expression: evaluate; prefer TLS string if the call chain
        // set one, otherwise stringify the numeric result.
        Grid32FormulaText::ResetText();
        double num = CFormulator::EvalArg(mgr, s, visited);
        if (Grid32FormulaText::g_textResultValid) {
            std::wstring r = Grid32FormulaText::g_textResult;
            Grid32FormulaText::ResetText();
            return r;
        }
        std::wstringstream ss;
        ss << num;
        return ss.str();
    }

    // ---- Date helpers ---------------------------------------------------
    // SerialToYMD now lives in grid32_internal.h alongside DateToSerial so
    // both are unit-testable. Use the qualified name where needed.
    using Grid32Detail::SerialToYMD;

    double TodaySerial() {
        std::time_t t = std::time(nullptr);
        std::tm lt{};
#ifdef _MSC_VER
        localtime_s(&lt, &t);
#else
        lt = *std::localtime(&t);
#endif
        return Grid32Detail::DateToSerial(lt.tm_year + 1900, lt.tm_mon + 1, lt.tm_mday);
    }

    double NowSerial() {
        std::time_t t = std::time(nullptr);
        std::tm lt{};
#ifdef _MSC_VER
        localtime_s(&lt, &t);
#else
        lt = *std::localtime(&t);
#endif
        double day = Grid32Detail::DateToSerial(lt.tm_year + 1900, lt.tm_mon + 1, lt.tm_mday);
        double frac = (lt.tm_hour * 3600 + lt.tm_min * 60 + lt.tm_sec) / 86400.0;
        return day + frac;
    }

}


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
        if (op == L'*')
        {
            value *= rhs;
        }
        else
        {
            // Division: NaN/inf on zero divisor — surfaces the error to the
            // caller via the result rather than silently returning the
            // dividend. Spreadsheet host will format these distinctly.
            value /= rhs;
        }
    }
    return value;
}

double CFormulator::EvalArg(CGrid32Mgr* mgr, const std::wstring& s,
    std::set<std::pair<UINT, UINT>>& visited) {
    size_t p = 0;
    return ParseExpression(mgr, s, p, visited);
}

double CFormulator::ParseExpression(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos,
    std::set<std::pair<UINT, UINT>>& visited) {
    FormulaDepthGuard depth;
    if (depth.overflowed) return 0.0;
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
        else if (_wcsicmp(token.c_str(), L"IF") == 0)
        {
            // IF(cond, then [, else]) — non-zero condition selects `then`.
            auto args = SplitArgs(arg);
            if (args.size() >= 2)
            {
                double cond = EvalArg(mgr, args[0], visited);
                if (cond != 0.0)
                    res = EvalArg(mgr, args[1], visited);
                else if (args.size() >= 3)
                    res = EvalArg(mgr, args[2], visited);
                else
                    res = 0.0;
            }
        }
        else if (_wcsicmp(token.c_str(), L"IFS") == 0)
        {
            // IFS(cond1, val1, cond2, val2, ...) — first truthy condition wins.
            auto args = SplitArgs(arg);
            for (size_t i = 0; i + 1 < args.size(); i += 2)
            {
                if (EvalArg(mgr, args[i], visited) != 0.0)
                {
                    res = EvalArg(mgr, args[i + 1], visited);
                    break;
                }
            }
        }
        else if (_wcsicmp(token.c_str(), L"AND") == 0)
        {
            auto args = SplitArgs(arg);
            res = 1.0;
            for (auto& a : args) {
                if (EvalArg(mgr, a, visited) == 0.0) { res = 0.0; break; }
            }
            if (args.empty()) res = 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"OR") == 0)
        {
            auto args = SplitArgs(arg);
            res = 0.0;
            for (auto& a : args) {
                if (EvalArg(mgr, a, visited) != 0.0) { res = 1.0; break; }
            }
        }
        else if (_wcsicmp(token.c_str(), L"NOT") == 0)
        {
            res = (EvalArg(mgr, arg, visited) == 0.0) ? 1.0 : 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"ABS") == 0)
        {
            res = std::fabs(EvalArg(mgr, arg, visited));
        }
        else if (_wcsicmp(token.c_str(), L"ROUND") == 0)
        {
            // ROUND(value, digits)
            auto args = SplitArgs(arg);
            double v = args.size() >= 1 ? EvalArg(mgr, args[0], visited) : 0.0;
            double d = args.size() >= 2 ? EvalArg(mgr, args[1], visited) : 0.0;
            double mult = std::pow(10.0, d);
            res = std::floor(v * mult + (v >= 0 ? 0.5 : -0.5)) / mult;
        }
        else if (_wcsicmp(token.c_str(), L"CEILING") == 0)
        {
            auto args = SplitArgs(arg);
            double v = args.size() >= 1 ? EvalArg(mgr, args[0], visited) : 0.0;
            double sig = args.size() >= 2 ? EvalArg(mgr, args[1], visited) : 1.0;
            res = sig != 0.0 ? std::ceil(v / sig) * sig : 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"FLOOR") == 0)
        {
            auto args = SplitArgs(arg);
            double v = args.size() >= 1 ? EvalArg(mgr, args[0], visited) : 0.0;
            double sig = args.size() >= 2 ? EvalArg(mgr, args[1], visited) : 1.0;
            res = sig != 0.0 ? std::floor(v / sig) * sig : 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"MOD") == 0)
        {
            // MOD(dividend, divisor)
            auto args = SplitArgs(arg);
            double d1 = args.size() >= 1 ? EvalArg(mgr, args[0], visited) : 0.0;
            double d2 = args.size() >= 2 ? EvalArg(mgr, args[1], visited) : 0.0;
            res = d2 != 0.0 ? std::fmod(d1, d2) : 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"INT") == 0)
        {
            // INT rounds toward -inf, like Excel.
            res = std::floor(EvalArg(mgr, arg, visited));
        }
        else if (_wcsicmp(token.c_str(), L"TRUNC") == 0)
        {
            // TRUNC drops the fractional part toward zero.
            double v = EvalArg(mgr, arg, visited);
            res = (v >= 0) ? std::floor(v) : std::ceil(v);
        }
        else if (_wcsicmp(token.c_str(), L"SQRT") == 0)
        {
            double v = EvalArg(mgr, arg, visited);
            res = (v >= 0) ? std::sqrt(v) : 0.0;
        }
        // -------- Text functions -----------------------------------------
        else if (_wcsicmp(token.c_str(), L"LEN") == 0)
        {
            std::wstring s = EvalAsString(mgr, arg, visited);
            res = (double)s.size();
        }
        else if (_wcsicmp(token.c_str(), L"LEFT") == 0)
        {
            auto args = SplitArgs(arg);
            std::wstring s = args.size() >= 1 ? EvalAsString(mgr, args[0], visited) : std::wstring();
            int n = args.size() >= 2 ? (int)EvalArg(mgr, args[1], visited) : 1;
            if (n < 0) n = 0;
            if ((size_t)n > s.size()) n = (int)s.size();
            Grid32FormulaText::SetText(s.substr(0, n));
            res = 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"RIGHT") == 0)
        {
            auto args = SplitArgs(arg);
            std::wstring s = args.size() >= 1 ? EvalAsString(mgr, args[0], visited) : std::wstring();
            int n = args.size() >= 2 ? (int)EvalArg(mgr, args[1], visited) : 1;
            if (n < 0) n = 0;
            if ((size_t)n > s.size()) n = (int)s.size();
            Grid32FormulaText::SetText(s.substr(s.size() - n));
            res = 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"MID") == 0)
        {
            // MID(text, start, length) — start is 1-based per Excel
            auto args = SplitArgs(arg);
            std::wstring s = args.size() >= 1 ? EvalAsString(mgr, args[0], visited) : std::wstring();
            int start = args.size() >= 2 ? (int)EvalArg(mgr, args[1], visited) : 1;
            int n = args.size() >= 3 ? (int)EvalArg(mgr, args[2], visited) : 0;
            if (start < 1) start = 1;
            size_t idx = (size_t)(start - 1);
            if (idx >= s.size() || n <= 0) Grid32FormulaText::SetText(std::wstring());
            else
            {
                if ((size_t)n > s.size() - idx) n = (int)(s.size() - idx);
                Grid32FormulaText::SetText(s.substr(idx, n));
            }
            res = 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"CONCAT") == 0 ||
                 _wcsicmp(token.c_str(), L"CONCATENATE") == 0)
        {
            auto args = SplitArgs(arg);
            std::wstring out;
            for (auto& a : args) out += EvalAsString(mgr, a, visited);
            Grid32FormulaText::SetText(out);
            res = 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"UPPER") == 0)
        {
            std::wstring s = EvalAsString(mgr, arg, visited);
            std::transform(s.begin(), s.end(), s.begin(), ::towupper);
            Grid32FormulaText::SetText(s);
            res = 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"LOWER") == 0)
        {
            std::wstring s = EvalAsString(mgr, arg, visited);
            std::transform(s.begin(), s.end(), s.begin(), ::towlower);
            Grid32FormulaText::SetText(s);
            res = 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"TRIM") == 0)
        {
            // Excel TRIM removes leading/trailing spaces and collapses
            // internal runs of spaces to single spaces.
            std::wstring s = EvalAsString(mgr, arg, visited);
            std::wstring out;
            bool inSpace = true;
            for (wchar_t c : s) {
                if (iswspace(c)) {
                    if (!inSpace && !out.empty()) out.push_back(L' ');
                    inSpace = true;
                } else {
                    out.push_back(c);
                    inSpace = false;
                }
            }
            // Strip a possible trailing space added before final non-space.
            while (!out.empty() && iswspace(out.back())) out.pop_back();
            Grid32FormulaText::SetText(out);
            res = 0.0;
        }
        // -------- Date functions -----------------------------------------
        else if (_wcsicmp(token.c_str(), L"TODAY") == 0)
        {
            res = TodaySerial();
        }
        else if (_wcsicmp(token.c_str(), L"NOW") == 0)
        {
            res = NowSerial();
        }
        else if (_wcsicmp(token.c_str(), L"DATE") == 0)
        {
            // DATE(y, m, d) — construct serial
            auto args = SplitArgs(arg);
            int y = args.size() >= 1 ? (int)EvalArg(mgr, args[0], visited) : 1900;
            int m = args.size() >= 2 ? (int)EvalArg(mgr, args[1], visited) : 1;
            int d = args.size() >= 3 ? (int)EvalArg(mgr, args[2], visited) : 1;
            res = Grid32Detail::DateToSerial(y, m, d);
        }
        else if (_wcsicmp(token.c_str(), L"YEAR") == 0 ||
                 _wcsicmp(token.c_str(), L"MONTH") == 0 ||
                 _wcsicmp(token.c_str(), L"DAY") == 0)
        {
            double serial = EvalArg(mgr, arg, visited);
            int y, m, d;
            SerialToYMD(serial, y, m, d);
            if (_wcsicmp(token.c_str(), L"YEAR") == 0)       res = (double)y;
            else if (_wcsicmp(token.c_str(), L"MONTH") == 0) res = (double)m;
            else                                              res = (double)d;
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
    FormulaDepthGuard depth;
    if (depth.overflowed) return 0.0;
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