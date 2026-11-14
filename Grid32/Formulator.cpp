#include "pch.h"
#include "grid32mgr.h"
#include "Formulator.h"
#include "grid32_internal.h"
#include <cmath>
#include <limits>
#include <chrono>
#include <ctime>
#include <algorithm>
#include <map>
#include <vector>
#include <locale>
#include <sstream>
#include <charconv>
#include <system_error>


// State for one evaluation: one top-level formula, or a whole recalculation
// pass (CFormulator::RecalcPass). Declared in Formulator.h so RecalcPass can
// hold one; nothing outside this file looks inside.
struct Grid32FormulaSession
{
    static constexpr size_t kNoCut = ~(size_t)0;

    CGrid32Mgr* mgr;
    // The formula's own cell when it was handed over by pointer
    // (CFormulator::Evaluate). It is on the path at depth 0 for the whole
    // evaluation, as a cell evaluated by position is.
    const GRIDCELL* self;
    // Results of formula cells evaluated so far (P4).
    std::map<std::pair<UINT, UINT>, CFormulator::Result> memo;
    // Formula cells being evaluated, outermost first; path[i] is at depth i+1.
    // Mirrors the `visited` set the parser threads through, but in order.
    std::vector<std::pair<UINT, UINT>> path;

    // What the subtree being evaluated has run into. A result is memoized only
    // when nothing below it was cut short: not by a cycle cut on a cell above
    // it on the path (the value would then depend on the route taken to the
    // cell), and not by the depth guard (the value would be truncated).
    size_t minCutDepth;
    bool overflowed;

    // One attempt is one evaluation from the top of a fresh path; see
    // Engine::ResolvePending. `attemptCell` is the cell it evaluates, if it is
    // a cell; `pending` the cells already waiting on other attempts.
    bool haveAttemptCell;
    std::pair<UINT, UINT> attemptCell;
    const std::vector<std::pair<UINT, UINT>>* pending;

    // The formula cell innermost on the path the first time the depth guard
    // tripped during this attempt. When evaluating that cell on its own first
    // would let a retry get further, the attempt is `abandoned`: the rest of
    // it evaluates no more formula cells, since its result will be discarded.
    // Without that, a cell that reads a truncated (so unmemoized) neighbour
    // twice re-walks it twice, and a doubled chain explodes before the
    // attempt can end.
    bool haveOverflowCell;
    std::pair<UINT, UINT> overflowCell;
    bool abandoned;

    Grid32FormulaSession(CGrid32Mgr* m, const GRIDCELL* s)
        : mgr(m), self(s), minCutDepth(kNoCut), overflowed(false),
          haveAttemptCell(false), attemptCell(0, 0), pending(nullptr),
          haveOverflowCell(false), overflowCell(0, 0), abandoned(false) {}

    void BeginAttempt(const std::pair<UINT, UINT>* cell,
        const std::vector<std::pair<UINT, UINT>>* waiting)
    {
        minCutDepth = kNoCut;
        overflowed = false;
        haveAttemptCell = cell != nullptr;
        if (cell) attemptCell = *cell;
        pending = waiting;
        haveOverflowCell = false;
        abandoned = false;
    }

    // Called once per attempt, when the depth guard first trips.
    void NoteOverflow()
    {
        overflowed = true;
        if (haveOverflowCell || path.empty())
            return;     // already placed, or the root expression itself is the deep part
        overflowCell = path.back();
        haveOverflowCell = true;
        abandoned = !(haveAttemptCell && overflowCell == attemptCell)   // its own formula is the deep part
            && memo.count(overflowCell) == 0
            && !(pending && std::find(pending->begin(), pending->end(), overflowCell) != pending->end());
    }
};

namespace {
    typedef std::pair<UINT, UINT> CellKey;

    // The session the current thread is evaluating in, if any.
    thread_local Grid32FormulaSession* g_session = nullptr;
    thread_local size_t g_cellEvaluations = 0;

    struct SessionScope {
        Grid32FormulaSession* outer;
        explicit SessionScope(Grid32FormulaSession* s) : outer(g_session) { g_session = s; }
        ~SessionScope() { g_session = outer; }
        SessionScope(const SessionScope&) = delete;
        SessionScope& operator=(const SessionScope&) = delete;
    };

    // Cap formula recursion to defend against pathological cycles and deeply
    // nested expressions that the visited-set alone can't bound (e.g.,
    // long alternating chains across multiple cells).
    constexpr int kMaxFormulaDepth = 256;
    thread_local int g_formulaDepth = 0;

    struct FormulaDepthGuard {
        bool overflowed;
        FormulaDepthGuard() : overflowed(false) {
            if (++g_formulaDepth > kMaxFormulaDepth)
            {
                overflowed = true;
                if (g_session) g_session->NoteOverflow();
            }
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

    inline bool IsAsciiDigit(wchar_t c) { return c >= L'0' && c <= L'9'; }

    inline void SkipSpaces(const std::wstring& s, size_t& pos) {
        while (pos < s.size() && iswspace(s[pos])) ++pos;
    }

    // Read text as a number the way a person would: "42", " -3.5 ", "1e3",
    // "1,234.50", "$1,234.50", "-$5", "25%" (= 0.25). Grouping commas must be
    // well formed, so "1,2" stays text; anything with trailing junk ("12abc")
    // is not a number at all, where std::stod would have read its prefix.
    // Used for cells whose text is all there is (CT_Text, and text results of
    // formulas). Typed cells carry their value in m_dValue and never get here.
    bool ParseNumericText(const std::wstring& s, double& out)
    {
        size_t i = 0, n = s.size();
        while (i < n && iswspace(s[i])) ++i;
        while (n > i && iswspace(s[n - 1])) --n;
        if (i == n) return false;

        bool neg = false, signed_ = false, currency = false, percent = false;
        if (s[i] == L'+' || s[i] == L'-') { neg = (s[i] == L'-'); signed_ = true; ++i; }
        if (i < n && s[i] == L'$')
        {
            currency = true;
            ++i;
            if (!signed_ && i < n && (s[i] == L'+' || s[i] == L'-'))   // "$-5"
            {
                neg = (s[i] == L'-');
                ++i;
            }
        }
        if (n > i && s[n - 1] == L'%') { percent = true; --n; }
        if (currency && percent) return false;

        std::string core;     // plain ASCII mantissa/exponent for strtod
        size_t intDigits = 0, groupDigits = 0;
        bool grouped = false;
        while (i < n && (IsAsciiDigit(s[i]) || s[i] == L','))
        {
            if (s[i] == L',')
            {
                // A separator needs 1-3 digits before the first one and
                // exactly 3 between later ones.
                if (intDigits == 0 || (grouped ? groupDigits != 3 : groupDigits > 3))
                    return false;
                grouped = true;
                groupDigits = 0;
            }
            else
            {
                core.push_back((char)s[i]);
                ++intDigits;
                ++groupDigits;
            }
            ++i;
        }
        if (grouped && groupDigits != 3) return false;

        size_t fracDigits = 0;
        if (i < n && s[i] == L'.')
        {
            core.push_back('.');
            ++i;
            while (i < n && IsAsciiDigit(s[i])) { core.push_back((char)s[i]); ++i; ++fracDigits; }
        }
        if (intDigits + fracDigits == 0) return false;

        if (i < n && (s[i] == L'e' || s[i] == L'E'))
        {
            core.push_back('e');
            ++i;
            if (i < n && (s[i] == L'+' || s[i] == L'-')) { core.push_back((char)s[i]); ++i; }
            size_t expDigits = 0;
            while (i < n && IsAsciiDigit(s[i])) { core.push_back((char)s[i]); ++i; ++expDigits; }
            if (expDigits == 0) return false;
        }
        if (i != n) return false;

        // from_chars rather than strtod: strtod follows the C locale, and
        // under a comma-decimal one read "1.5" as 1.
        double v = 0.0;
        const std::from_chars_result fc = std::from_chars(core.data(), core.data() + core.size(), v);
        if (fc.ec != std::errc() || !std::isfinite(v)) return false;   // "1e999" is not a number a cell can hold
        if (percent) v /= 100.0;
        out = neg ? -v : v;
        return true;
    }

    // `x` rounded to 15 significant digits, the precision Excel works in.
    // Binary doubles hold most decimals only approximately: 1.005 is stored
    // as 1.00499999999999989..., so 1.005*100 is 100.49999999999999 and a
    // plain round gives 100. Snapping first gives the 100.5 the user wrote.
    // to_chars/from_chars are locale-independent, unlike printf and streams.
    double Snap15(double x)
    {
        if (!std::isfinite(x) || x == 0.0) return x;
        char buf[64];
        const std::to_chars_result w = std::to_chars(buf, buf + sizeof(buf), x,
            std::chars_format::general, 15);
        if (w.ec != std::errc()) return x;
        double snapped = x;
        const std::from_chars_result r = std::from_chars(buf, w.ptr, snapped);
        return r.ec == std::errc() ? snapped : x;
    }

    enum class RoundMode { HalfAwayFromZero, TowardZero };

    // ROUND(value, digits) and TRUNC(value, digits). `digits` is truncated to
    // an integer, as Excel does, and may be negative (ROUND(1250,-2) = 1300).
    // The scaled value is snapped to 15 digits before rounding (see Snap15).
    double RoundToDigits(double value, double digits, RoundMode mode)
    {
        if (!std::isfinite(value) || !std::isfinite(digits)) return value;
        const double d = std::trunc(digits);
        if (std::fabs(d) > 400.0)               // past any double's range
            return d > 0 ? value : 0.0;
        const double scale = std::pow(10.0, std::fabs(d));
        if (!std::isfinite(scale))
            return d > 0 ? value : 0.0;
        double x = d >= 0 ? value * scale : value / scale;
        if (!std::isfinite(x)) return value;    // already finer than 10^-d
        x = Snap15(x);
        x = mode == RoundMode::HalfAwayFromZero ? std::round(x) : std::trunc(x);
        const double out = d >= 0 ? x / scale : x * scale;
        return out == 0.0 ? 0.0 : out;          // no "-0"
    }

    // CEILING / FLOOR: the quotient is snapped before it is rounded, so
    // FLOOR(0.3,0.1) is 0.3 (0.3/0.1 is 2.9999999999999996), and the result
    // is snapped so it reads as the multiple it is.
    double RoundToMultiple(double value, double significance, bool up)
    {
        if (significance == 0.0) return 0.0;
        const double q = Snap15(value / significance);
        const double out = Snap15((up ? std::ceil(q) : std::floor(q)) * significance);
        return out == 0.0 ? 0.0 : out;
    }

    // A count or position argument (LEFT, RIGHT, MID) as a size_t. Casting
    // the double straight to int is undefined once it is out of int's range:
    // LEFT("abc",1e20) came out as "" and MID("abcdef",3e9,2) as "ab".
    // Truncates toward zero, as Excel does; NaN and anything below 1 give 0,
    // and anything past `limit` gives `limit`.
    size_t ClampCount(double v, size_t limit)
    {
        if (!(v >= 1.0)) return 0;
        if (v >= (double)limit) return limit;
        return (size_t)v;
    }

    // A whole numeric literal as the parser sees it: digits, an optional
    // fraction, an optional exponent ("12", ".5", "1.5e-3"). "12abc" and
    // "1.2.3" are not, where std::stod would have read their prefix.
    bool IsNumberLiteral(const std::wstring& t)
    {
        size_t i = 0, digits = 0;
        while (i < t.size() && IsAsciiDigit(t[i])) { ++i; ++digits; }
        if (i < t.size() && t[i] == L'.')
        {
            ++i;
            while (i < t.size() && IsAsciiDigit(t[i])) { ++i; ++digits; }
        }
        if (digits == 0) return false;
        if (i < t.size() && (t[i] == L'e' || t[i] == L'E'))
        {
            ++i;
            if (i < t.size() && (t[i] == L'+' || t[i] == L'-')) ++i;
            size_t exp = 0;
            while (i < t.size() && IsAsciiDigit(t[i])) { ++i; ++exp; }
            if (exp == 0) return false;
        }
        return i == t.size();
    }

    // Text comparison for the comparison operators: case-insensitive, as in
    // Excel ("abc" = "ABC"). Returns <0, 0 or >0.
    int CompareTextNoCase(const std::wstring& a, const std::wstring& b)
    {
        const size_t n = a.size() < b.size() ? a.size() : b.size();
        for (size_t i = 0; i < n; ++i)
        {
            const wint_t ca = towupper(a[i]), cb = towupper(b[i]);
            if (ca != cb) return ca < cb ? -1 : 1;
        }
        return a.size() == b.size() ? 0 : (a.size() < b.size() ? -1 : 1);
    }

    // Text as VALUE reads it: a number by ParseNumericText, or a date by
    // what typing into a cell accepts (TryParseDate: ISO and US), as its
    // serial. False for anything else.
    bool TextToNumber(const std::wstring& text, double& out)
    {
        return ParseNumericText(text, out) || Grid32Detail::TryParseDate(TrimWs(text), out);
    }

    // A reference names a cell only if that cell is inside the grid. ParseCellRef
    // accepts any syntactically valid reference up to row 2^32, so without this
    // =SUM(A1:A4294967296) walked four billion rows on the UI thread (H21).
    // A reference outside the grid is treated exactly like a malformed one
    // ("A0"): it does not resolve, and the expression reads it as 0 — the
    // engine has no error values yet (E2).
    bool InGrid(CGrid32Mgr* mgr, UINT row, UINT col)
    {
        return mgr != nullptr
            && (size_t)row < mgr->GetGridRowCount()
            && (size_t)col < mgr->GetGridColCount();
    }

    bool ResolveCellRef(CGrid32Mgr* mgr, const std::wstring& token, UINT& row, UINT& col)
    {
        return CFormulator::ParseCellRef(token, row, col) && InGrid(mgr, row, col);
    }

    // What a function argument is as a range.
    enum class RangeArg {
        NotARange,  // an expression (or nothing range-like): evaluate it
        OffGrid,    // range syntax, but not inside the grid: contributes nothing
        Ok,         // sr..er x sc..ec, inclusive and inside the grid
    };

    // Column letters only ("A", "zz"), as a 0-based column index.
    bool ParseColumnLetters(const std::wstring& t, UINT& col)
    {
        if (t.empty() || t.size() > 7) return false;
        unsigned long long c = 0;
        for (wchar_t ch : t)
        {
            if (ch >= L'a' && ch <= L'z') ch = (wchar_t)(ch - L'a' + L'A');
            if (ch < L'A' || ch > L'Z') return false;
            c = c * 26 + (unsigned long long)(ch - L'A' + 1);
        }
        if (c == 0 || c > 0xFFFFFFFFull) return false;
        col = (UINT)(c - 1);
        return true;
    }

    // Row digits only ("7"), as a 0-based row index.
    bool ParseRowDigits(const std::wstring& t, UINT& row)
    {
        if (t.empty() || t.size() > 10) return false;
        unsigned long long r = 0;
        for (wchar_t ch : t)
        {
            if (!IsAsciiDigit(ch)) return false;
            r = r * 10 + (unsigned long long)(ch - L'0');
        }
        if (r == 0 || r > 0x100000000ull) return false;
        row = (UINT)(r - 1);
        return true;
    }

    // Resolve a range argument against the grid: a cell ("A1"), a rectangle
    // ("A1:B3", either corner first), whole columns ("A:A", "B:D") or whole
    // rows ("2:2", "1:5"). Whole columns and rows span the grid's extent;
    // walking them costs what the sheet holds (ForEachPopulatedCell), not
    // what they span. A cell or rectangle that runs off the grid is rejected
    // whole rather than clamped, as a single out-of-grid reference is (H21);
    // so are columns or rows beyond it.
    RangeArg ResolveRangeArg(CGrid32Mgr* mgr, const std::wstring& arg,
        UINT& sr, UINT& sc, UINT& er, UINT& ec)
    {
        if (CFormulator::ParseRange(arg, sr, sc, er, ec))
            return InGrid(mgr, sr, sc) && InGrid(mgr, er, ec) ? RangeArg::Ok : RangeArg::OffGrid;

        const size_t colon = arg.find(L':');
        if (colon == std::wstring::npos || !mgr) return RangeArg::NotARange;
        const std::wstring a = TrimWs(arg.substr(0, colon)), b = TrimWs(arg.substr(colon + 1));
        const UINT rows = (UINT)mgr->GetGridRowCount(), cols = (UINT)mgr->GetGridColCount();
        UINT lo = 0, hi = 0;
        if (ParseColumnLetters(a, lo) && ParseColumnLetters(b, hi))
        {
            if (lo > hi) std::swap(lo, hi);
            if (rows == 0 || hi >= cols) return RangeArg::OffGrid;
            sr = 0; er = rows - 1; sc = lo; ec = hi;
            return RangeArg::Ok;
        }
        if (ParseRowDigits(a, lo) && ParseRowDigits(b, hi))
        {
            if (lo > hi) std::swap(lo, hi);
            if (cols == 0 || hi >= rows) return RangeArg::OffGrid;
            sc = 0; ec = cols - 1; sr = lo; er = hi;
            return RangeArg::Ok;
        }
        return RangeArg::NotARange;
    }

    // A number format that shows a value as a date. A number under one reads
    // as a date in a formula (=A1+1 on a date-formatted 45306 is a date).
    bool IsDateFormat(UINT fmt)
    {
        return fmt == FMT_DATE_ISO || fmt == FMT_DATE_US;
    }

    // Visit the populated cells of an inclusive rectangle in row-major order,
    // the order the old dense loops used. Blank cells contribute nothing to
    // any aggregate, so walking the sparse map costs what the sheet holds
    // rather than what the range spans.
    template <class Fn>
    void ForEachPopulatedCell(CGrid32Mgr* mgr, UINT sr, UINT sc, UINT er, UINT ec, Fn fn)
    {
        const auto& cells = mgr->GetCellMap();
        auto it = cells.lower_bound(CellKey(sr, sc));
        while (it != cells.end() && it->first.first <= er)
        {
            const UINT r = it->first.first;
            const UINT c = it->first.second;
            if (c < sc)
            {
                it = cells.lower_bound(CellKey(r, sc));
                continue;
            }
            if (c > ec)
            {
                if (r == er) break;
                it = cells.lower_bound(CellKey(r + 1, sc));
                continue;
            }
            if (it->second)
                fn(it->first, it->second);
            ++it;
        }
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

    // Keeps a cell on the current evaluation path for the guard's lifetime.
    // The path set is what detects genuine cycles (a cell that references
    // itself, directly or through a chain); it must NOT accumulate every cell
    // ever seen, or a value referenced twice in one formula evaluates as 0 the
    // second time. Erasing in the destructor also unwinds correctly if an
    // evaluation throws. The session's ordered copy of the path is kept in
    // step, so a cycle cut can be placed by depth.
    struct PathGuard {
        std::set<std::pair<UINT, UINT>>& path;
        std::pair<UINT, UINT> key;
        Grid32FormulaSession* session;
        PathGuard(std::set<std::pair<UINT, UINT>>& p, const std::pair<UINT, UINT>& k)
            : path(p), key(k), session(g_session)
        {
            path.insert(key);
            if (session) session->path.push_back(key);
        }
        ~PathGuard()
        {
            if (session) session->path.pop_back();
            path.erase(key);
        }
        PathGuard(const PathGuard&) = delete;
        PathGuard& operator=(const PathGuard&) = delete;
    };

    // Gives one formula cell's evaluation a fresh record of what its subtree
    // runs into, then folds that record into the enclosing subtree's.
    struct SubtreeScope {
        Grid32FormulaSession* s;
        size_t outerMinCut;
        bool outerOverflow;
        explicit SubtreeScope(Grid32FormulaSession* session)
            : s(session), outerMinCut(Grid32FormulaSession::kNoCut), outerOverflow(false)
        {
            if (!s) return;
            outerMinCut = s->minCutDepth;
            outerOverflow = s->overflowed;
            s->minCutDepth = Grid32FormulaSession::kNoCut;
            s->overflowed = false;
        }
        // True when the result computed at `depth` does not depend on how the
        // cell was reached: every cycle cut below it was on itself or on a
        // cell below it, and the depth guard never tripped.
        bool Clean(size_t depth) const
        {
            return s && !s->overflowed && s->minCutDepth >= depth;
        }
        ~SubtreeScope()
        {
            if (!s) return;
            if (outerMinCut < s->minCutDepth) s->minCutDepth = outerMinCut;
            s->overflowed = s->overflowed || outerOverflow;
        }
        SubtreeScope(const SubtreeScope&) = delete;
        SubtreeScope& operator=(const SubtreeScope&) = delete;
    };

    // Record a cycle cut on `key` against the subtree being evaluated.
    void NoteCycleCut(const CellKey& key, const GRIDCELL* cell)
    {
        Grid32FormulaSession* s = g_session;
        if (!s) return;
        size_t depth = 0;    // the session's own cell, or one we cannot place
        if (cell != s->self)
        {
            for (size_t i = 0; i < s->path.size(); ++i)
                if (s->path[i] == key) { depth = i + 1; break; }
        }
        if (depth < s->minCutDepth) s->minCutDepth = depth;
    }

    // Comparison operators, lowest precedence in the grammar.
    enum class CompareOp { None, Lt, Gt, Le, Ge, Eq, Ne };
}


// ---- Evaluation internals ---------------------------------------------------
// Nested in CFormulator so it can reach the private parts of the class (and
// Result); nothing outside this file sees it.
//
// The parser passes Values. It used to pass a double and keep everything else
// about a value in thread-local side channels: a text slot (is this value
// text, and which), a date flag, an error slot (the first error so far), and
// a check of the source text for "is this operand a bare reference to an
// empty cell". Every operator and function had to set, clear, park and
// restore those in the right order, and the blank check only worked when the
// reference was written bare, so =(A1)="" was FALSE while =A1="" was TRUE
// (review item #13). A Value carries all of it, so each rule lives in one
// place: ToNumber, ToCondition and ToText say how a value reads in each role,
// and errors travel as values, the first one met (left to right) winning.
struct CFormulator::Engine
{
    typedef std::set<std::pair<UINT, UINT>> Path;

    // ---- Values -------------------------------------------------------------
    struct Value
    {
        enum Kind : unsigned char {
            Number,     // num; isDate when it is a date (a General cell shows it as one)
            Text,       // text
            Blank,      // a reference to an empty cell: 0 as a number, "" as text
            Error,      // error, a CFormulator::kError* code
        };
        Kind kind = Number;
        bool isDate = false;
        double num = 0.0;
        std::wstring text;
        const wchar_t* error = nullptr;

        static Value Num(double v, bool date = false) { Value x; x.num = v; x.isDate = date; return x; }
        static Value Str(std::wstring t) { Value x; x.kind = Text; x.text = std::move(t); return x; }
        static Value Empty() { Value x; x.kind = Blank; return x; }
        static Value Err(const wchar_t* code) { Value x; x.kind = Error; x.error = code; return x; }
        bool IsError() const { return kind == Error; }
    };

    // A value where a number is needed — an operand of + - * / or of a minus
    // sign, a numeric function's argument. A number is itself (a date stays a
    // date, for the + and - rules); a blank cell is 0; text is the number it
    // reads as ("42", "$1,234.50", "25%", or a date such as "2024-01-15"), as
    // in Excel, and #VALUE! if it reads as none; an error is itself.
    static Value ToNumber(const Value& v)
    {
        switch (v.kind)
        {
        case Value::Number: return v;
        case Value::Blank:  return Value::Num(0.0);
        case Value::Text:
        {
            double n = 0.0;
            return TextToNumber(v.text, n) ? Value::Num(n) : Value::Err(kErrorValue);
        }
        default:            return v;
        }
    }

    // A value as a condition (IF, IFS, NOT): true when not 0. Text "TRUE" and
    // "FALSE" are true and false, as in Excel; other text reads as ToNumber
    // reads it, so "abc" is #VALUE!. A blank cell is false.
    static Value ToCondition(const Value& v)
    {
        if (v.kind != Value::Text) return ToNumber(v);
        const std::wstring text = TrimWs(v.text);
        if (Grid32Detail::IEquals(text, L"TRUE")) return Value::Num(1.0);
        if (Grid32Detail::IEquals(text, L"FALSE")) return Value::Num(0.0);
        return ToNumber(v);
    }

    // A value where text is needed (LEN, LEFT, CONCAT, ...): text is itself, a
    // number its General form (a date its serial, as in Excel), a blank cell
    // "", an error itself.
    static Value ToText(const Value& v)
    {
        switch (v.kind)
        {
        case Value::Number: return Value::Str(FormatGeneral(v.num));
        case Value::Blank:  return Value::Str(std::wstring());
        default:            return v;
        }
    }

    // The static code for an error's text (a Result carries it as a string).
    static const wchar_t* ErrorCode(const std::wstring& text)
    {
        static const wchar_t* const kCodes[] = {
            kErrorSyntax, kErrorName, kErrorDiv0, kErrorNum, kErrorValue, kErrorNA };
        for (const wchar_t* code : kCodes)
            if (text == code) return code;
        return kErrorSyntax;
    }

    // A formula cell's (or a reference's) Result as a Value, and back.
    static Value FromResult(const Result& r)
    {
        if (r.isError) return Value::Err(ErrorCode(r.text));
        if (r.isText) return Value::Str(r.text);
        return Value::Num(r.value, r.isDate);
    }

    static Result ErrorResult(const wchar_t* code)
    {
        Result r;
        r.isError = true;
        r.text = code;
        r.value = std::numeric_limits<double>::quiet_NaN();
        return r;
    }

    // What a formula shows. A text result keeps its text read as a number in
    // `value` for callers that need one; a blank is 0; a number with no
    // finite value — an overflow (POWER(10,400)) or a domain error that
    // reached the result — is #NUM! ("inf" and "nan" used to show); -0 is 0.
    static Result ToResult(const Value& v)
    {
        Result r;
        switch (v.kind)
        {
        case Value::Error:
            return ErrorResult(v.error);
        case Value::Text:
        {
            r.isText = true;
            r.text = v.text;
            double n = 0.0;
            r.value = ParseNumericText(r.text, n) ? n : 0.0;
            return r;
        }
        case Value::Blank:
            return r;
        default:
            if (!std::isfinite(v.num)) return ErrorResult(kErrorNum);
            r.value = v.num == 0.0 ? 0.0 : v.num;   // -0 (=-A1 with A1 blank) showed as "-0"
            r.isDate = v.isDate;
            return r;
        }
    }

    // ---- Sources and references ----------------------------------------------

    // Evaluate formula source (a stored formula never carries its '=', but
    // one is tolerated rather than read as a comparison against nothing).
    static Result EvalSource(CGrid32Mgr* mgr, const std::wstring& formula, Path& visited)
    {
        size_t pos = 0;
        if (!formula.empty() && formula[0] == L'=') pos = 1;

        Result passthrough;
        if (BareReference(mgr, formula.substr(pos), visited, passthrough))
            return passthrough;
        return ToResult(ParseWhole(mgr, formula, pos, visited));
    }

    // Parse `expr` from `pos` as one whole expression. Anything left over is
    // a syntax error: "2(3)", "1 2", "A1:B2" outside a function, a stray ')'.
    // (An error already met comes first.)
    static Value ParseWhole(CGrid32Mgr* mgr, const std::wstring& expr, size_t pos, Path& visited)
    {
        Value v = ParseExpression(mgr, expr, pos, visited);
        SkipSpaces(expr, pos);
        if (pos < expr.size() && !v.IsError())
            return Value::Err(kErrorSyntax);
        return v;
    }

    // A function's argument. An omitted one, as in IF(A1,,5), is 0; anything
    // else must parse whole: ABS(1 2) is a syntax error, not 1.
    static Value EvalValue(CGrid32Mgr* mgr, const std::wstring& arg, Path& visited)
    {
        if (TrimWs(arg).empty()) return Value::Num(0.0);
        return ParseWhole(mgr, arg, 0, visited);
    }

    static Value EvalNumber(CGrid32Mgr* mgr, const std::wstring& arg, Path& visited)
    {
        return ToNumber(EvalValue(mgr, arg, visited));
    }

    static Value EvalCondition(CGrid32Mgr* mgr, const std::wstring& arg, Path& visited)
    {
        return ToCondition(EvalValue(mgr, arg, visited));
    }

    // A text function's argument. A bare cell reference reads the cell as
    // CellAsText does ("1234.5" for a currency cell, not "$1,234.50");
    // anything else is evaluated and read as ToText reads it.
    static Value EvalText(CGrid32Mgr* mgr, const std::wstring& arg, Path& visited)
    {
        const std::wstring s = TrimWs(arg);
        if (s.empty()) return Value::Str(std::wstring());
        UINT row = 0, col = 0;
        if (ResolveCellRef(mgr, s, row, col))
            return CellAsText(mgr, CellKey(row, col), mgr->GetCell(row, col), visited);
        return ToText(ParseWhole(mgr, s, 0, visited));
    }

    // A formula that is nothing but a reference (=A1) takes the referenced
    // cell as it is: a formula's result, text included, or a text cell's
    // text. Returns false when `expr` is not a bare reference.
    static bool BareReference(CGrid32Mgr* mgr, const std::wstring& expr, Path& visited, Result& out)
    {
        UINT row = 0, col = 0;
        if (!ResolveCellRef(mgr, TrimWs(expr), row, col)) return false;
        out = ReferenceValue(mgr, row, col, visited);
        return true;
    }

    // A reference (inside the grid) as a Result: a formula's result (text or
    // an error included), a text cell's text, or a typed or blank cell's
    // number. A cycle reads as an empty numeric result.
    static Result ReferenceValue(CGrid32Mgr* mgr, UINT row, UINT col, Path& visited)
    {
        const CellKey key(row, col);
        const GRIDCELL* cell = mgr->GetCell(row, col);
        Result out;

        if (cell && cell->m_bFormula)
        {
            FormulaCellResult(mgr, key, cell, visited, out);
            // A formula cell the user formatted as a date reads as one too.
            if (!out.isText && !out.isError && IsDateFormat(cell->m_nFormat))
                out.isDate = true;
            return out;
        }
        if (IsTextCell(cell))
        {
            out.isText = true;
            out.text = cell->m_wsText;
            double n = 0.0;
            out.value = ParseNumericText(out.text, n) ? n : 0.0;
            return out;
        }
        // Blank or typed: its value, as any other reference reads it.
        double v = 0.0;
        bool isNumber = false;
        ReadCell(mgr, key, cell, visited, v, isNumber, nullptr, &out.isDate);
        out.value = isNumber ? v : 0.0;
        return out;
    }

    // A typed cell whose value is a date: a typed date, or a number shown
    // under a date format.
    static bool IsDateValueCell(const GRIDCELL* cell)
    {
        return cell && !cell->m_bFormula && !cell->m_wsText.empty() &&
            (cell->m_eType == CT_Date ||
             ((cell->m_eType == CT_Number) && IsDateFormat(cell->m_nFormat)));
    }

    // A non-formula cell whose content is its text: CT_Text, and anything
    // else that is not a typed value. Empty text is a blank cell, whatever
    // type it kept when it was cleared.
    static bool IsTextCell(const GRIDCELL* cell)
    {
        return cell && !cell->m_bFormula && !cell->m_wsText.empty() &&
            cell->m_eType != CT_Number && cell->m_eType != CT_Date && cell->m_eType != CT_Boolean;
    }

    // A cell that holds nothing: no cell at all, or one with no formula and
    // no text. Blank compares equal to both 0 and "".
    static bool IsBlankCell(const GRIDCELL* cell)
    {
        return !cell || (!cell->m_bFormula && cell->m_wsText.empty());
    }

    // A formula cell's result: from the memo when it is there, otherwise by
    // evaluating its source and memoizing the result if it is clean. Returns
    // false, with `out` an empty numeric result, when the cell is already on
    // the evaluation path (a cycle) or the depth guard tripped.
    static bool FormulaCellResult(CGrid32Mgr* mgr, const CellKey& key, const GRIDCELL* cell,
        Path& visited, Result& out)
    {
        out = Result();
        FormulaDepthGuard depth;
        if (depth.overflowed) return false;

        Grid32FormulaSession* s = g_session;
        // `visited` is the chain of formula cells currently being evaluated,
        // not a log of every cell touched. Finding this cell already on the
        // chain is a real cycle (A1 -> B1 -> A1) and yields 0; finding it a
        // second time *alongside* the first (=A1+A1, or A1 read by two
        // different formulas in one recalc) must evaluate normally — and with
        // the memo, a second reading costs a lookup rather than a re-walk.
        if (visited.count(key) || (s && cell == s->self))
        {
            NoteCycleCut(key, cell);
            return false;
        }

        const bool memoize = s && s->mgr == mgr;
        if (memoize)
        {
            auto it = s->memo.find(key);
            if (it != s->memo.end()) { out = it->second; return true; }
        }
        if (s && s->abandoned)
        {
            // This attempt will be retried; don't spend anything on it.
            s->overflowed = true;
            return false;
        }

        ++g_cellEvaluations;
        const size_t myDepth = s ? s->path.size() + 1 : 0;
        // Only formula cells can recurse, so only they need to go on the
        // chain. PathGuard removes the entry again on the way out, including
        // when the nested evaluation throws.
        PathGuard guard(visited, key);
        SubtreeScope subtree(s);
        out = EvalSource(mgr, cell->m_wsFormula, visited);
        if (memoize && subtree.Clean(myDepth))
            s->memo[key] = out;
        return true;
    }

    // What a cell in a range contributes to an aggregate (S3). Returns false
    // for a blank cell. `isNumber` says whether it takes part in SUM /
    // AVERAGE / MIN / MAX / COUNT / AND / OR; `isLogical` that it is a TRUE or
    // FALSE cell, which AND and OR count and the others, as in Excel, skip.
    //
    // Typed cells are read from m_dValue, never from their display text: a
    // currency cell shows "$1,234.50", a percentage "25.00%", a date
    // "2024-01-15", a boolean "TRUE", and std::stod read those as 0, 25, 2024
    // and 0. Text is not a number here, however it reads — a text cell, or a
    // formula's text result — as in Excel: '42 is typed as text precisely so
    // that it is not one. (Review #9: it counted when it read as a number.)
    // Arithmetic on a reference is another matter: =A1+1 reads "42" as 42.
    //
    // A formula cell whose result is an error sets `error` to its code and
    // reads as not a number; the aggregate decides what the error means.
    // `isDate`, when given, says whether the number is a date (MIN and MAX of
    // dates are dates).
    static bool ReadCell(CGrid32Mgr* mgr, const CellKey& key, const GRIDCELL* cell,
        Path& visited, double& value, bool& isNumber, const wchar_t** error = nullptr,
        bool* isDate = nullptr, bool* isLogical = nullptr)
    {
        value = 0.0;
        isNumber = false;
        if (isDate) *isDate = false;
        if (isLogical) *isLogical = false;
        if (!cell) return false;

        if (cell->m_bFormula)
        {
            Result r;
            if (!FormulaCellResult(mgr, key, cell, visited, r))
            {
                // A cycle or a truncated chain reads as 0 and still counts as
                // a number, as a formula cell always has.
                isNumber = true;
                return true;
            }
            if (r.isError)
            {
                if (error) *error = ErrorCode(r.text);
                return true;
            }
            if (r.isText) return true;                  // text: not a number
            value = r.value;
            isNumber = true;
            if (isDate) *isDate = r.isDate || IsDateFormat(cell->m_nFormat);
            return true;
        }

        // Blank. This also covers a cleared cell, which keeps the m_eType and
        // m_dValue it had before ClearCellText emptied its text.
        if (cell->m_wsText.empty()) return false;

        switch (cell->m_eType)
        {
        case CT_Number:
        case CT_Date:
        case CT_Boolean:
            value = cell->m_dValue;
            isNumber = true;
            if (isDate) *isDate = IsDateValueCell(cell);
            if (isLogical) *isLogical = cell->m_eType == CT_Boolean;
            return true;
        default:
            return true;                                // text: not a number
        }
    }

    // A cell read in a text function (LEN, LEFT, CONCAT, ...). Formula cells
    // give their result as it would display under FMT_GENERAL; typed cells
    // under a number format give their value the way the user would edit it
    // ("1234.5", not "$1,234.50"), consistent with ReadCell reading the value
    // rather than the rendering.
    static Value CellAsText(CGrid32Mgr* mgr, const CellKey& key, const GRIDCELL* cell, Path& visited)
    {
        if (!cell) return Value::Str(std::wstring());
        if (cell->m_bFormula)
        {
            Result r;
            FormulaCellResult(mgr, key, cell, visited, r);   // a cycle reads as 0 here too
            if (r.isError) return Value::Err(ErrorCode(r.text));
            return Value::Str(r.isText ? r.text : FormatGeneral(r.value));
        }
        if (cell->m_nFormat != FMT_GENERAL && !cell->m_wsText.empty() &&
            (cell->m_eType == CT_Number || cell->m_eType == CT_Date || cell->m_eType == CT_Boolean))
        {
            return Value::Str(Grid32Detail::FormatEditValue(cell->m_eType, cell->m_dValue, cell->m_wsText));
        }
        return Value::Str(cell->m_wsText);
    }

    // ---- Comparisons ---------------------------------------------------------

    static bool ApplyCompare(CompareOp op, int cmp)
    {
        switch (op)
        {
        case CompareOp::Lt: return cmp < 0;
        case CompareOp::Gt: return cmp > 0;
        case CompareOp::Le: return cmp <= 0;
        case CompareOp::Ge: return cmp >= 0;
        case CompareOp::Eq: return cmp == 0;
        case CompareOp::Ne: return cmp != 0;
        default:            return false;
        }
    }

    // How two values (neither an error) compare:
    //  - numbers, at the 15 significant digits Excel works in (Snap15), so
    //    0.1+0.2 = 0.3 while 1000000000 and 1000000001 differ;
    //  - text, case-insensitively, as in Excel ("abc" = "ABC");
    //  - a blank cell as whatever the other side is: "" against text, 0
    //    against a number, so =A1="" and =A1=0 are both true for empty A1 —
    //    and so are =(A1)="" and =IF(TRUE,A1)="", which the old source-text
    //    check could not see;
    //  - text against a number: numerically when the text reads as a number
    //    (the engine reads "42" as 42 everywhere else too); otherwise text is
    //    greater than any number, as Excel orders them.
    static bool Compare(CompareOp op, Value a, Value b)
    {
        if (a.kind == Value::Blank && b.kind == Value::Blank)
            return ApplyCompare(op, 0);
        if (a.kind == Value::Blank)
            a = b.kind == Value::Text ? Value::Str(std::wstring()) : Value::Num(0.0);
        if (b.kind == Value::Blank)
            b = a.kind == Value::Text ? Value::Str(std::wstring()) : Value::Num(0.0);

        if (a.kind == Value::Text && b.kind == Value::Text)
            return ApplyCompare(op, CompareTextNoCase(a.text, b.text));
        if (a.kind == Value::Text)
        {
            if (!ParseNumericText(a.text, a.num)) return ApplyCompare(op, 1);
        }
        else if (b.kind == Value::Text)
        {
            if (!ParseNumericText(b.text, b.num)) return ApplyCompare(op, -1);
        }

        const double lhs = Snap15(a.num), rhs = Snap15(b.num);
        switch (op)
        {
        case CompareOp::Lt: return lhs < rhs;
        case CompareOp::Gt: return lhs > rhs;
        case CompareOp::Le: return lhs <= rhs;
        case CompareOp::Ge: return lhs >= rhs;
        case CompareOp::Eq: return lhs == rhs;
        case CompareOp::Ne: return lhs != rhs;
        default:            return false;
        }
    }

    // ---- The grammar ---------------------------------------------------------
    //   expression := additive [ ( = <> < > <= >= ) additive ]
    //   additive   := term { ( + - ) term }
    //   term       := factor { ( * / ) factor }
    //   factor     := ( expression ) | "text" | +factor | -factor
    //               | number | reference | TRUE | FALSE | name( arguments )
    //
    // Every level parses all of its input even after an error, so the
    // position is right for what follows, and returns the first error met.

    // Comparisons bind least tightly, so "A1+1 > B1*2" compares the two
    // arithmetic results. Only one comparison is read; a second ("=1=1=1")
    // is left over, which ParseWhole reports as a syntax error. The result
    // is 1 (true) or 0 (false).
    static Value ParseExpression(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos, Path& visited)
    {
        FormulaDepthGuard depth;
        if (depth.overflowed)
            return Value::Err(kErrorSyntax);    // nested too deeply to evaluate

        Value lhs = ParseAdditive(mgr, expr, pos, visited);

        SkipSpaces(expr, pos);
        if (pos >= expr.size()) return lhs;

        // Two-character operators first, so "<=" isn't read as "<" followed
        // by a stray "=".
        CompareOp op = CompareOp::None;
        if (expr.compare(pos, 2, L"<>") == 0)      { op = CompareOp::Ne; pos += 2; }
        else if (expr.compare(pos, 2, L"<=") == 0) { op = CompareOp::Le; pos += 2; }
        else if (expr.compare(pos, 2, L">=") == 0) { op = CompareOp::Ge; pos += 2; }
        else if (expr[pos] == L'<')                { op = CompareOp::Lt; ++pos; }
        else if (expr[pos] == L'>')                { op = CompareOp::Gt; ++pos; }
        else if (expr[pos] == L'=')                { op = CompareOp::Eq; ++pos; }
        else return lhs;

        Value rhs = ParseAdditive(mgr, expr, pos, visited);
        if (lhs.IsError()) return lhs;
        if (rhs.IsError()) return rhs;
        return Value::Num(Compare(op, lhs, rhs) ? 1.0 : 0.0);
    }

    // + and -: both operands as numbers (ToNumber). A date moved by a number
    // of days is a date (A1+7, 7+A1, A1-7); two dates added, or one
    // subtracted from another, are not: A2-A1 is a number of days.
    static Value ParseAdditive(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos, Path& visited)
    {
        Value value = ParseTerm(mgr, expr, pos, visited);
        for (;;)
        {
            SkipSpaces(expr, pos);
            if (pos >= expr.size()) break;
            const wchar_t op = expr[pos];
            if (op != L'+' && op != L'-') break;
            ++pos;
            const Value a = ToNumber(value);
            const Value b = ToNumber(ParseTerm(mgr, expr, pos, visited));
            if (a.IsError()) { value = a; continue; }
            if (b.IsError()) { value = b; continue; }
            const bool date = op == L'+' ? a.isDate != b.isDate : (a.isDate && !b.isDate);
            value = Value::Num(op == L'+' ? a.num + b.num : a.num - b.num, date);
        }
        return value;
    }

    // * and /: both operands as numbers; neither a product nor a quotient is
    // a date. Dividing by zero is #DIV/0!, as in Excel (it was inf, and
    // before that the dividend, L7).
    static Value ParseTerm(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos, Path& visited)
    {
        Value value = ParseFactor(mgr, expr, pos, visited);
        for (;;)
        {
            SkipSpaces(expr, pos);
            if (pos >= expr.size()) break;
            const wchar_t op = expr[pos];
            if (op != L'*' && op != L'/') break;
            ++pos;
            const Value a = ToNumber(value);
            const Value b = ToNumber(ParseFactor(mgr, expr, pos, visited));
            if (a.IsError()) { value = a; continue; }
            if (b.IsError()) { value = b; continue; }
            if (op == L'*')          value = Value::Num(a.num * b.num);
            else if (b.num == 0.0)   value = Value::Err(kErrorDiv0);
            else                     value = Value::Num(a.num / b.num);
        }
        return value;
    }

    static Value ParseFactor(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos, Path& visited)
    {
        // Unary +/- below recurses straight back into ParseFactor, which would
        // otherwise sidestep ParseExpression's guard and let "=-----...-1" run
        // the stack out.
        FormulaDepthGuard depth;
        if (depth.overflowed)
            return Value::Err(kErrorSyntax);

        SkipSpaces(expr, pos);
        if (pos >= expr.size())
            return Value::Err(kErrorSyntax);    // an operand is missing: "=", "=1+", "=("

        if (expr[pos] == L'(')
        {
            ++pos;
            Value v = ParseExpression(mgr, expr, pos, visited);
            SkipSpaces(expr, pos);
            if (pos < expr.size() && expr[pos] == L')') ++pos;
            else if (!v.IsError()) v = Value::Err(kErrorSyntax);   // "=(1+2"
            return v;    // text, a date or a blank cell stays what it is: =(A1)=""
        }
        if (expr[pos] == L'"')
        {
            // A string literal; "" inside it is one quote. Its value is text,
            // which IF can hand on and a comparison can compare (="abc",
            // =IF(A1>3,"yes","no"), =A1="abc").
            std::wstring literal;
            ++pos;
            while (pos < expr.size())
            {
                if (expr[pos] == L'"')
                {
                    if (pos + 1 < expr.size() && expr[pos + 1] == L'"') { literal.push_back(L'"'); pos += 2; continue; }
                    ++pos;
                    return Value::Str(std::move(literal));
                }
                literal.push_back(expr[pos++]);
            }
            return Value::Err(kErrorSyntax);    // no closing quote
        }
        if (expr[pos] == L'+' || expr[pos] == L'-')
        {
            const bool neg = expr[pos] == L'-';
            ++pos;
            Value v = ParseFactor(mgr, expr, pos, visited);
            if (!neg) return v;                  // +A1 is A1, text included
            const Value n = ToNumber(v);        // -"abc" is #VALUE!; -LEFT("5",1) is -5
            return n.IsError() ? n : Value::Num(-n.num);
        }

        const size_t start = pos;
        while (pos < expr.size() && (iswalnum(expr[pos]) || expr[pos] == L'.')) ++pos;
        // A numeric literal with a signed exponent ("1e-3", "2.5E+2"): the scan
        // above stops at the sign, which left "1e" to be read as 1 and the "-3"
        // to be subtracted from it. A token that starts with a digit or '.' is
        // never a cell reference or a function name, so this cannot steal one.
        if (pos > start && pos + 1 < expr.size() &&
            (expr[pos] == L'+' || expr[pos] == L'-') && IsAsciiDigit(expr[pos + 1]) &&
            (IsAsciiDigit(expr[start]) || expr[start] == L'.') &&
            (expr[pos - 1] == L'e' || expr[pos - 1] == L'E'))
        {
            bool mantissa = true;
            for (size_t i = start; i + 1 < pos; ++i)
                if (!IsAsciiDigit(expr[i]) && expr[i] != L'.') { mantissa = false; break; }
            if (mantissa)
            {
                pos += 2;   // the sign and the first exponent digit
                while (pos < expr.size() && IsAsciiDigit(expr[pos])) ++pos;
            }
        }
        const std::wstring token = expr.substr(start, pos - start);
        if (token.empty())
            return Value::Err(kErrorSyntax);    // something no operand starts with: "=*2", "=)", "=$A$1"

        if (pos < expr.size() && expr[pos] == L'(')
        {
            ++pos; // skip '('
            const size_t argStart = pos;
            // Find the matching ')'. Parentheses inside a string literal are
            // text, not structure: without tracking quotes, LEN("a)b") closed
            // at the quoted ')' and measured "a. "" inside a literal is an
            // escaped quote and leaves quote mode unchanged, as in SplitArgs.
            int level = 1;
            bool inQuote = false;
            while (pos < expr.size() && level > 0)
            {
                const wchar_t c = expr[pos];
                if (c == L'"') inQuote = !inQuote;
                else if (!inQuote && c == L'(') level++;
                else if (!inQuote && c == L')') level--;
                if (level > 0) ++pos;
            }
            if (pos >= expr.size())
            {
                // No closing ')' ("=SUM(A1:A3"), or an unterminated literal
                // inside the call swallowed it.
                return Value::Err(kErrorSyntax);
            }
            const std::wstring arg = expr.substr(argStart, pos - argStart);
            ++pos;  // the ')'
            return CallFunction(mgr, token, arg, visited);
        }

        // A number must be one whole literal: =12abc read as 12.
        if (IsAsciiDigit(token[0]) || token[0] == L'.')
        {
            if (!IsNumberLiteral(token))
                return Value::Err(kErrorSyntax);
            // IsNumberLiteral has checked it is ASCII. from_chars, unlike
            // stod, ignores the C locale's decimal separator.
            std::string narrow;
            narrow.reserve(token.size());
            for (wchar_t c : token) narrow.push_back(static_cast<char>(c));
            double v = 0.0;
            const std::from_chars_result r = std::from_chars(narrow.data(), narrow.data() + narrow.size(), v);
            if (r.ec != std::errc() || !std::isfinite(v))
                return Value::Err(kErrorNum);   // 1e999
            return Value::Num(v);
        }

        UINT row = 0, col = 0;
        if (ParseCellRef(token, row, col))
        {
            // A reference outside the grid does not resolve (see InGrid) and
            // reads as 0, as it always has.
            if (!InGrid(mgr, row, col)) return Value::Num(0.0);
            // An empty cell is Blank: 0 as a number, "" as text, and equal to
            // both in a comparison. Anything else is the cell as a value:
            // text stays text, an error is the error, a date is a date.
            if (IsBlankCell(mgr->GetCell(row, col))) return Value::Empty();
            return FromResult(ReferenceValue(mgr, row, col, visited));
        }

        if (_wcsicmp(token.c_str(), L"TRUE") == 0) return Value::Num(1.0);
        if (_wcsicmp(token.c_str(), L"FALSE") == 0) return Value::Num(0.0);

        // A name that is neither a reference nor a keyword. It used to read
        // as 0.
        return Value::Err(kErrorName);
    }

    // ---- Aggregates ------------------------------------------------------------

    // Feed every number an aggregate's arguments hold to `fn(number,
    // isDate)`: SUM, AVERAGE, MIN, MAX, COUNT, AND and OR all take any mix of
    // ranges, references and expressions (E4), and agree on what counts.
    //
    //  - A range or reference contributes its populated cells that hold
    //    numbers (ReadCell): blanks and text are skipped, numeric-looking
    //    text included, and so are TRUE/FALSE cells unless `countLogical`
    //    (AND, OR), as in Excel. One that runs off the grid contributes
    //    nothing. Ranges may be whole columns or rows (SUM(A:A), SUM(2:2);
    //    see ResolveRangeArg).
    //  - Any other argument is evaluated and contributes one value. Text
    //    counts when it reads as a number or a date ("5", "2024-01-15"), as
    //    in Excel; other text is #VALUE!. An omitted argument ("A1,,A2")
    //    contributes nothing.
    //
    // Returns the first error met, in a cell or an argument, or null. With
    // `skipErrors` (COUNT) errors and non-numeric text simply do not count,
    // and the result is always null.
    template <class Fn>
    static const wchar_t* ForEachAggregateValue(CGrid32Mgr* mgr, const std::wstring& arglist, Path& visited,
        bool skipErrors, Fn fn, bool countLogical = false)
    {
        const wchar_t* firstError = nullptr;
        auto note = [&](const wchar_t* code) { if (!skipErrors && !firstError) firstError = code; };

        for (const std::wstring& raw : SplitArgs(arglist))
        {
            const std::wstring a = TrimWs(raw);
            if (a.empty()) continue;

            UINT sr = 0, sc = 0, er = 0, ec = 0;
            const RangeArg kind = ResolveRangeArg(mgr, a, sr, sc, er, ec);
            if (kind == RangeArg::OffGrid) continue;
            if (kind == RangeArg::Ok)
            {
                ForEachPopulatedCell(mgr, sr, sc, er, ec, [&](const CellKey& key, PGRIDCELL cell) {
                    double v = 0.0;
                    bool isNumber = false, isDate = false, isLogical = false;
                    const wchar_t* cellError = nullptr;
                    if (!ReadCell(mgr, key, cell, visited, v, isNumber, &cellError, &isDate, &isLogical))
                        return;
                    if (cellError) note(cellError);
                    else if (isNumber && (countLogical || !isLogical)) fn(v, isDate);
                });
                continue;
            }

            const Value v = ParseWhole(mgr, a, 0, visited);
            switch (v.kind)
            {
            case Value::Error:
                note(v.error);
                break;
            case Value::Text:
            {
                double n = 0.0;
                if (TextToNumber(v.text, n)) fn(n, false);
                else note(kErrorValue);
                break;
            }
            case Value::Blank:
                fn(0.0, false);
                break;
            default:
                fn(v.num, v.isDate);
                break;
            }
        }
        return firstError;
    }

    // The aggregates agree with each other about what a cell is worth. Blank
    // cells, text and TRUE/FALSE cells in a range are skipped — matching
    // Excel, and E6: treating them as 0 would drag MIN to 0 and inflate
    // AVERAGE's denominator.
    static Value Sum(CGrid32Mgr* mgr, const std::wstring& arg, Path& visited)
    {
        double total = 0.0;
        if (const wchar_t* e = ForEachAggregateValue(mgr, arg, visited, false,
                [&](double v, bool) { total += v; }))
            return Value::Err(e);
        return Value::Num(total);
    }

    static Value Average(CGrid32Mgr* mgr, const std::wstring& arg, Path& visited)
    {
        double total = 0.0;
        UINT count = 0;
        if (const wchar_t* e = ForEachAggregateValue(mgr, arg, visited, false,
                [&](double v, bool) { total += v; ++count; }))
            return Value::Err(e);
        // The average of nothing is undefined, as in Excel. It read 0 while
        // the engine had no error values (E6).
        if (count == 0) return Value::Err(kErrorDiv0);
        return Value::Num(total / count);
    }

    // MIN / MAX. The latest or earliest of some dates is a date. An all-blank
    // range has no minimum or maximum; Excel reports 0.
    static Value MinMax(CGrid32Mgr* mgr, const std::wstring& arg, Path& visited, bool bMax)
    {
        double best = 0.0;
        bool any = false, allDates = true;
        if (const wchar_t* e = ForEachAggregateValue(mgr, arg, visited, false, [&](double v, bool isDate) {
                allDates = allDates && isDate;
                if (!any) { best = v; any = true; }
                else if (bMax) { if (v > best) best = v; }
                else           { if (v < best) best = v; }
            }))
            return Value::Err(e);
        return Value::Num(any ? best : 0.0, any && allDates);
    }

    // COUNT tallies the values the other aggregates would use. A formula cell
    // is evaluated to find out (a text result is not a number); the memo
    // makes that a lookup during a recalculation, and a formula that is on
    // the path (COUNT over a range that includes its own cell) still counts,
    // as it always has. Errors are not numbers and, as in Excel, do not make
    // COUNT an error.
    static Value Count(CGrid32Mgr* mgr, const std::wstring& arg, Path& visited)
    {
        double count = 0.0;
        ForEachAggregateValue(mgr, arg, visited, true, [&](double, bool) { count += 1.0; });
        return Value::Num(count);
    }

    // AND / OR take conditions, values and ranges alike: AND(A1:A3),
    // OR(A1>5, B1:B9). They read their arguments the way the aggregates do,
    // except that TRUE/FALSE cells in a range count (text is skipped), and
    // like Excel they look at every argument, so an error anywhere is the
    // result (AND(0,1/0) is #DIV/0!). With nothing to look at there is no
    // answer: #VALUE!, as in Excel.
    static Value AndOr(CGrid32Mgr* mgr, const std::wstring& arg, Path& visited, bool isAnd)
    {
        bool any = false, all = true, some = false;
        if (const wchar_t* e = ForEachAggregateValue(mgr, arg, visited, false, [&](double v, bool) {
                any = true;
                if (v != 0.0) some = true; else all = false;
            }, true))
            return Value::Err(e);
        if (!any) return Value::Err(kErrorValue);
        return Value::Num((isAnd ? all : some) ? 1.0 : 0.0);
    }

    // ---- Functions ---------------------------------------------------------------

    // The first error among some values, or null.
    static const Value* FirstError(std::initializer_list<const Value*> values)
    {
        for (const Value* v : values)
            if (v->IsError()) return v;
        return nullptr;
    }

    // A text function's result: the text, or the error met on the way.
    static Value TextOr(const Value& failed, std::wstring text)
    {
        return failed.IsError() ? failed : Value::Str(std::move(text));
    }

    // Call the function `name` on the argument text `arg` (between its
    // parentheses). Arguments are evaluated left to right; the first error
    // met is the result, unless the function is the kind that catches them.
    static Value CallFunction(CGrid32Mgr* mgr, const std::wstring& name, const std::wstring& arg, Path& visited)
    {
        std::wstring fn = name;
        for (wchar_t& c : fn) if (c >= L'a' && c <= L'z') c = (wchar_t)(c - L'a' + L'A');

        // -------- Aggregates and logic --------------------------------------
        if (fn == L"SUM")     return Sum(mgr, arg, visited);
        if (fn == L"AVERAGE") return Average(mgr, arg, visited);
        if (fn == L"MIN")     return MinMax(mgr, arg, visited, false);
        if (fn == L"MAX")     return MinMax(mgr, arg, visited, true);
        if (fn == L"COUNT")   return Count(mgr, arg, visited);
        if (fn == L"AND")     return AndOr(mgr, arg, visited, true);
        if (fn == L"OR")      return AndOr(mgr, arg, visited, false);

        if (fn == L"IF")
        {
            // IF(cond, then [, else]): a true condition selects `then`, which
            // is handed on as it is (text, a date, a blank cell). Two or three
            // arguments: IF(1) read as 0, and extra ones were ignored.
            const auto args = SplitArgs(arg);
            if (args.size() < 2 || args.size() > 3) return Value::Err(kErrorSyntax);
            const Value cond = EvalCondition(mgr, args[0], visited);
            if (cond.IsError()) return cond;
            if (cond.num != 0.0) return EvalValue(mgr, args[1], visited);
            if (args.size() == 3) return EvalValue(mgr, args[2], visited);
            return Value::Num(0.0);
        }
        if (fn == L"IFS")
        {
            // IFS(cond1, val1, cond2, val2, ...): the first true condition
            // wins. Pairs only (IFS(1) read as 0), and with no condition true
            // there is no value: #N/A, as in Excel (it read 0).
            const auto args = SplitArgs(arg);
            if (args.size() < 2 || args.size() % 2 != 0) return Value::Err(kErrorSyntax);
            for (size_t i = 0; i + 1 < args.size(); i += 2)
            {
                const Value cond = EvalCondition(mgr, args[i], visited);
                if (cond.IsError()) return cond;
                if (cond.num != 0.0) return EvalValue(mgr, args[i + 1], visited);
            }
            return Value::Err(kErrorNA);
        }
        if (fn == L"NOT")
        {
            if (TrimWs(arg).empty()) return Value::Err(kErrorSyntax);
            const Value cond = EvalCondition(mgr, arg, visited);
            if (cond.IsError()) return cond;
            return Value::Num(cond.num == 0.0 ? 1.0 : 0.0);
        }

        // -------- Errors and blanks -----------------------------------------
        if (fn == L"IFERROR")
        {
            // IFERROR(value, value_if_error): `value`, text and dates
            // included, unless it is an error; then the fallback. A syntax
            // error is not caught: it is a mistake in the formula, not a
            // value, and hiding it would hide the typo.
            const auto args = SplitArgs(arg);
            if (args.size() != 2) return Value::Err(kErrorSyntax);
            const Value v = EvalValue(mgr, args[0], visited);
            if (v.IsError() && v.error != kErrorSyntax) return EvalValue(mgr, args[1], visited);
            return v;
        }
        if (fn == L"ISERROR")
        {
            // 1 when the argument is an error (#DIV/0!, #NAME?, #NUM!,
            // #VALUE!, #N/A), else 0. Never an error itself, but for a
            // malformed argument, as with IFERROR.
            if (TrimWs(arg).empty()) return Value::Err(kErrorSyntax);
            const Value v = EvalValue(mgr, arg, visited);
            if (v.IsError() && v.error == kErrorSyntax) return v;
            return Value::Num(v.IsError() ? 1.0 : 0.0);
        }
        if (fn == L"ISBLANK")
        {
            // 1 only for an empty cell: not for "", not for a formula whose
            // result is "", not for 0. A reference outside the grid has
            // nothing in it either. Anything else is evaluated (so a malformed
            // argument is still a syntax error); it is blank when it is a
            // reference to an empty cell, as (A1) is.
            const std::wstring t = TrimWs(arg);
            if (t.empty()) return Value::Err(kErrorSyntax);
            UINT row = 0, col = 0;
            if (ParseCellRef(t, row, col))
                return Value::Num((!InGrid(mgr, row, col) || IsBlankCell(mgr->GetCell(row, col))) ? 1.0 : 0.0);
            const Value v = EvalValue(mgr, t, visited);
            if (v.IsError() && v.error == kErrorSyntax) return v;
            return Value::Num(v.kind == Value::Blank ? 1.0 : 0.0);
        }

        // -------- Maths --------------------------------------------------------
        if (fn == L"SIN" || fn == L"COS" || fn == L"TAN")
        {
            const Value v = EvalNumber(mgr, arg, visited);
            if (v.IsError()) return v;
            return Value::Num(fn == L"SIN" ? std::sin(v.num) : fn == L"COS" ? std::cos(v.num) : std::tan(v.num));
        }
        if (fn == L"POWER")
        {
            // Split at top-level commas only: arg.find(L',') cut
            // POWER(ROUND(2.5,0),2) inside ROUND's own argument list. Exactly
            // two: POWER(2) read as 0^0 = 1. No finite result (POWER(0,-1))
            // is #NUM!.
            const auto args = SplitArgs(arg);
            if (args.size() != 2) return Value::Err(kErrorSyntax);
            const Value base = EvalNumber(mgr, args[0], visited);
            const Value exp = EvalNumber(mgr, args[1], visited);
            if (const Value* e = FirstError({ &base, &exp })) return *e;
            return Value::Num(std::pow(base.num, exp.num));
        }
        if (fn == L"ROOT")
        {
            const auto args = SplitArgs(arg);
            const Value val = args.size() >= 1 ? EvalNumber(mgr, args[0], visited) : Value::Num(0.0);
            const Value n = args.size() >= 2 ? EvalNumber(mgr, args[1], visited) : Value::Num(2.0);
            if (const Value* e = FirstError({ &val, &n })) return *e;
            return Value::Num(n.num != 0.0 ? std::pow(val.num, 1.0 / n.num) : 0.0);
        }
        if (fn == L"LOG" || fn == L"LN")
        {
            // Logarithms of zero or a negative number have no value; they
            // showed "-inf" and "nan".
            const Value v = EvalNumber(mgr, arg, visited);
            if (v.IsError()) return v;
            if (!(v.num > 0.0)) return Value::Err(kErrorNum);
            return Value::Num(fn == L"LOG" ? std::log10(v.num) : std::log(v.num));
        }
        if (fn == L"ABS")
        {
            const Value v = EvalNumber(mgr, arg, visited);
            return v.IsError() ? v : Value::Num(std::fabs(v.num));
        }
        if (fn == L"ROUND" || fn == L"TRUNC")
        {
            // ROUND: half away from zero, on the decimal the user sees: 1.005
            // rounds to 1.01, not the 1.00 its binary value (1.00499999...)
            // gives. TRUNC(value [, digits]) drops digits toward zero.
            const auto args = SplitArgs(arg);
            const Value v = args.size() >= 1 ? EvalNumber(mgr, args[0], visited) : Value::Num(0.0);
            const Value d = args.size() >= 2 ? EvalNumber(mgr, args[1], visited) : Value::Num(0.0);
            if (const Value* e = FirstError({ &v, &d })) return *e;
            return Value::Num(RoundToDigits(v.num, d.num,
                fn == L"ROUND" ? RoundMode::HalfAwayFromZero : RoundMode::TowardZero));
        }
        if (fn == L"CEILING" || fn == L"FLOOR")
        {
            const auto args = SplitArgs(arg);
            const Value v = args.size() >= 1 ? EvalNumber(mgr, args[0], visited) : Value::Num(0.0);
            const Value sig = args.size() >= 2 ? EvalNumber(mgr, args[1], visited) : Value::Num(1.0);
            if (const Value* e = FirstError({ &v, &sig })) return *e;
            return Value::Num(RoundToMultiple(v.num, sig.num, fn == L"CEILING"));
        }
        if (fn == L"MOD")
        {
            // The result takes the divisor's sign, as in Excel: MOD(-3,2) = 1
            // and MOD(3,-2) = -1 (std::fmod takes the dividend's). MOD(x,0)
            // is #DIV/0!.
            const auto args = SplitArgs(arg);
            const Value d1 = args.size() >= 1 ? EvalNumber(mgr, args[0], visited) : Value::Num(0.0);
            const Value d2 = args.size() >= 2 ? EvalNumber(mgr, args[1], visited) : Value::Num(0.0);
            if (const Value* e = FirstError({ &d1, &d2 })) return *e;
            if (d2.num == 0.0) return Value::Err(kErrorDiv0);
            double r = std::fmod(d1.num, d2.num);
            if (r == 0.0) r = 0.0;                        // fmod(-6,3) is -0.0, which shows as "-0"
            else if ((r < 0.0) != (d2.num < 0.0)) r += d2.num;
            return Value::Num(r);
        }
        if (fn == L"INT")
        {
            // INT rounds toward -inf, like Excel.
            const Value v = EvalNumber(mgr, arg, visited);
            return v.IsError() ? v : Value::Num(std::floor(v.num));
        }
        if (fn == L"SQRT")
        {
            const Value v = EvalNumber(mgr, arg, visited);
            if (v.IsError()) return v;
            return v.num >= 0 ? Value::Num(std::sqrt(v.num)) : Value::Err(kErrorNum);
        }

        // -------- Text -----------------------------------------------------------
        if (fn == L"LEN")
        {
            const Value s = EvalText(mgr, arg, visited);
            return s.IsError() ? s : Value::Num((double)s.text.size());
        }
        if (fn == L"LEFT" || fn == L"RIGHT")
        {
            // A negative count has no meaning: #VALUE!, as in Excel.
            const auto args = SplitArgs(arg);
            const Value s = args.size() >= 1 ? EvalText(mgr, args[0], visited) : Value::Str(std::wstring());
            const Value count = args.size() >= 2 ? EvalNumber(mgr, args[1], visited) : Value::Num(1.0);
            if (const Value* e = FirstError({ &s, &count })) return *e;
            if (count.num < 0.0) return Value::Err(kErrorValue);
            const size_t n = ClampCount(count.num, s.text.size());
            return Value::Str(fn == L"LEFT" ? s.text.substr(0, n) : s.text.substr(s.text.size() - n));
        }
        if (fn == L"MID")
        {
            // MID(text, start, length), start 1-based. Three arguments; a
            // start below 1 or a negative length is #VALUE!, as in Excel.
            // Past the end there is nothing.
            const auto args = SplitArgs(arg);
            const Value s = args.size() >= 1 ? EvalText(mgr, args[0], visited) : Value::Str(std::wstring());
            if (s.IsError()) return s;
            if (args.size() != 3) return Value::Err(kErrorSyntax);
            const Value start = EvalNumber(mgr, args[1], visited);
            const Value length = EvalNumber(mgr, args[2], visited);
            if (const Value* e = FirstError({ &start, &length })) return *e;
            if (start.num < 1.0 || length.num < 0.0) return Value::Err(kErrorValue);
            size_t pos1 = ClampCount(start.num, s.text.size() + 1);
            if (pos1 == 0) pos1 = 1;
            const size_t idx = pos1 - 1;
            if (idx >= s.text.size()) return Value::Str(std::wstring());
            return Value::Str(s.text.substr(idx, ClampCount(length.num, s.text.size() - idx)));
        }
        if (fn == L"CONCAT" || fn == L"CONCATENATE")
        {
            // A range argument contributes its cells' text in row-major
            // order, blanks adding nothing: CONCAT(A1:C1). A single reference
            // reads as EvalText does.
            std::wstring out;
            for (const std::wstring& a : SplitArgs(arg))
            {
                const std::wstring t = TrimWs(a);
                UINT sr = 0, sc = 0, er = 0, ec = 0;
                const RangeArg kind = t.find(L':') != std::wstring::npos
                    ? ResolveRangeArg(mgr, t, sr, sc, er, ec) : RangeArg::NotARange;
                if (kind == RangeArg::NotARange)
                {
                    const Value s = EvalText(mgr, a, visited);
                    if (s.IsError()) return s;
                    out += s.text;
                }
                else if (kind == RangeArg::Ok)
                {
                    Value failed;
                    ForEachPopulatedCell(mgr, sr, sc, er, ec, [&](const CellKey& key, PGRIDCELL cell) {
                        if (failed.IsError()) return;
                        const Value s = CellAsText(mgr, key, cell, visited);
                        if (s.IsError()) failed = s;
                        else out += s.text;
                    });
                    if (failed.IsError()) return failed;
                }
            }
            return Value::Str(std::move(out));
        }
        if (fn == L"UPPER" || fn == L"LOWER")
        {
            Value s = EvalText(mgr, arg, visited);
            if (s.IsError()) return s;
            std::transform(s.text.begin(), s.text.end(), s.text.begin(),
                fn == L"UPPER" ? ::towupper : ::towlower);
            return s;
        }
        if (fn == L"TRIM")
        {
            // Excel TRIM removes leading and trailing spaces and collapses
            // internal runs of spaces to one. Spaces only (U+0020): tabs,
            // line breaks and non-breaking spaces are text and stay.
            const Value s = EvalText(mgr, arg, visited);
            if (s.IsError()) return s;
            std::wstring out;
            bool pendingSpace = false;
            for (wchar_t c : s.text)
            {
                if (c == L' ')
                {
                    pendingSpace = !out.empty();
                }
                else
                {
                    if (pendingSpace) out.push_back(L' ');
                    pendingSpace = false;
                    out.push_back(c);
                }
            }
            return Value::Str(std::move(out));
        }
        if (fn == L"VALUE")
        {
            // VALUE(text): text read as a number, #VALUE! when it is not one
            // — read as arithmetic reads it (ToNumber): " 42 ", "1e3",
            // "1,234.50", "$1,234.50", "25%" = 0.25, and date text as its
            // serial, as in Excel. A number passes through (a plain number,
            // not a date) and a blank cell is 0.
            if (TrimWs(arg).empty()) return Value::Err(kErrorSyntax);  // VALUE() has nothing to read
            const Value v = ToNumber(EvalValue(mgr, arg, visited));
            return v.IsError() ? v : Value::Num(v.num);
        }

        // -------- Dates -----------------------------------------------------------
        if (fn == L"TODAY") return Value::Num(TodaySerial(), true);
        if (fn == L"NOW")   return Value::Num(NowSerial(), true);
        if (fn == L"DATE")
        {
            // As Excel's DATE: the parts are truncated to whole numbers; a
            // year 0..1899 is that many years after 1900 (DATE(24,1,1) is
            // 1924-01-01); a month or day out of range rolls over
            // (DATE(2024,13,1) is 2025-01-01); and the result must be a date
            // Excel can hold, 0 (1900-01-00) to 9999-12-31, or it is #NUM!.
            // The size check comes first: the parts go to int, and out of
            // its range that cast is undefined.
            const auto args = SplitArgs(arg);
            const Value y = args.size() >= 1 ? EvalNumber(mgr, args[0], visited) : Value::Num(1900.0);
            const Value m = args.size() >= 2 ? EvalNumber(mgr, args[1], visited) : Value::Num(1.0);
            const Value d = args.size() >= 3 ? EvalNumber(mgr, args[2], visited) : Value::Num(1.0);
            if (const Value* e = FirstError({ &y, &m, &d })) return *e;
            const double kMaxPart = 1e9;
            if (!(std::fabs(y.num) <= kMaxPart && std::fabs(m.num) <= kMaxPart && std::fabs(d.num) <= kMaxPart))
                return Value::Err(kErrorNum);
            int year = (int)y.num;
            if (year >= 0 && year < 1900) year += 1900;
            const double serial = (year < 0 || year > 9999)
                ? -1.0 : Grid32Detail::DateToSerial(year, (int)m.num, (int)d.num);
            if (serial < 0.0 || serial > Grid32Detail::kMaxDateSerial) return Value::Err(kErrorNum);
            return Value::Num(serial, true);
        }
        if (fn == L"YEAR" || fn == L"MONTH" || fn == L"DAY")
        {
            // Date text reads as its date, as VALUE reads it:
            // YEAR("2024-01-15") is 2024. A serial outside Excel's dates,
            // 0 (1900-01-00) to 2958465 (9999-12-31), has no year, month or
            // day: #NUM!, as in Excel. 0 and 60 are the days Excel shows as
            // 1900-01-00 and 1900-02-29.
            const Value v = EvalNumber(mgr, arg, visited);
            if (v.IsError()) return v;
            if (!(v.num >= 0.0 && v.num < Grid32Detail::kMaxDateSerial + 1.0)) return Value::Err(kErrorNum);
            int y = 0, m = 0, d = 0;
            SerialToYMD(v.num, y, m, d);
            return Value::Num(fn == L"YEAR" ? (double)y : fn == L"MONTH" ? (double)m : (double)d);
        }

        if (fn == L"TRUE")  return Value::Num(1.0);
        if (fn == L"FALSE") return Value::Num(0.0);

        // Not a function this engine knows. It used to read as 0.
        return Value::Err(iswalpha(name[0]) ? kErrorName : kErrorSyntax);
    }

    // ---- Deep chains -----------------------------------------------------------

    // Evaluate every cell on `pending` (innermost last) from the top of a fresh
    // path, memoizing each.
    //
    // When an attempt trips the depth guard, the formula cell that was
    // innermost at the trip is pushed and evaluated on its own first — with
    // the whole depth budget to itself — so the retry finds it memoized and
    // gets further. A chain far longer than the guard allows therefore
    // resolves exactly, whatever order the cells are visited in, instead of
    // reading 0 wherever the guard tripped. A cell that cannot get further
    // (its own expression is the deep part) keeps the result it reached.
    //
    // Terminates: every push is a cell that is neither memoized nor already
    // pending (Grid32FormulaSession::NoteOverflow), and every pop memoizes one.
    static void ResolvePending(CGrid32Mgr* mgr, Grid32FormulaSession& s, std::vector<CellKey>& pending)
    {
        while (!pending.empty())
        {
            const CellKey key = pending.back();
            if (s.memo.count(key)) { pending.pop_back(); continue; }

            s.BeginAttempt(&key, &pending);
            Result r;
            PGRIDCELL cell = mgr ? mgr->GetCell(key.first, key.second) : nullptr;
            if (cell && cell->m_bFormula)
            {
                Path visited;
                FormulaCellResult(mgr, key, cell, visited, r);
            }
            if (s.abandoned)
            {
                pending.push_back(s.overflowCell);
                continue;
            }
            s.memo[key] = r;
            pending.pop_back();
        }
    }

    // Run `attempt` — one evaluation of an expression that is not itself a
    // cell — until it completes without tripping the depth guard, resolving
    // the cell it tripped in first each time (see ResolvePending).
    template <class Attempt>
    static void DriveRoot(CGrid32Mgr* mgr, Grid32FormulaSession& s, Attempt attempt)
    {
        for (;;)
        {
            s.BeginAttempt(nullptr, nullptr);
            attempt();
            if (!s.abandoned) return;
            std::vector<CellKey> pending(1, s.overflowCell);
            ResolvePending(mgr, s, pending);
        }
    }
};


bool CFormulator::ParseCellRef(const std::wstring& token, UINT& row, UINT& col) {
    // Column letters are A-Z only. iswalpha accepts any letter the CRT
    // knows, and MSVC's knows Unicode: "É1" went on to be read as column
    // 'É' - 'A' + 1 = 137, so =É1 silently read EG1.
    auto asciiLetter = [](wchar_t ch) {
        return (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z');
    };
    size_t i = 0;
    unsigned long long c = 0;
    while (i < token.size() && asciiLetter(token[i])) {
        const wchar_t upper = (token[i] >= L'a') ? (wchar_t)(token[i] - L'a' + L'A') : token[i];
        c = c * 26 + (unsigned long long)(upper - L'A' + 1);
        if (c > 0xFFFFFFFFull) return false;   // column index overflow
        ++i;
    }
    if (i == 0 || i >= token.size()) return false;

    // The remainder must be digits and nothing else. std::stoul stops at the
    // first non-digit without complaining, so without this check "A1:B10"
    // parses as a valid reference to A1 and every range collapses to its
    // first cell; "A1B" would likewise be accepted as A1.
    for (size_t j = i; j < token.size(); ++j)
        if (!IsAsciiDigit(token[j])) return false;

    unsigned long long r = 0;
    try {
        r = std::stoull(token.substr(i));
    }
    catch (...) {
        return false;
    }
    // Rows are 1-based in reference syntax; "A0" has no zero-based index.
    if (r == 0 || r > 0x100000000ull) return false;

    row = (UINT)(r - 1);
    col = (UINT)(c - 1);
    return true;
}

bool CFormulator::ParseRange(const std::wstring& arg, UINT& sr, UINT& sc,
    UINT& er, UINT& ec)
{
    size_t colon = arg.find(L':');
    if (colon == std::wstring::npos)
    {
        if (!ParseCellRef(TrimWs(arg), sr, sc)) return false;
        er = sr; ec = sc;
    }
    else
    {
        if (!ParseCellRef(TrimWs(arg.substr(0, colon)), sr, sc)) return false;
        if (!ParseCellRef(TrimWs(arg.substr(colon + 1)), er, ec)) return false;
    }
    // Normalize so callers can always iterate start..end inclusive; a range
    // written "A3:A1" is the same range as "A1:A3".
    if (sr > er) std::swap(sr, er);
    if (sc > ec) std::swap(sc, ec);
    return true;
}

// Evaluate an expression and return it as a number: a text result as the
// number it reads as (0 if none), a blank as 0, an error as NaN (Evaluate
// says which). Called from outside an evaluation, as the tests do, it is an
// entry point in its own right, with a memo and deep chains resolved.
double CFormulator::EvalArg(CGrid32Mgr* mgr, const std::wstring& s,
    std::set<std::pair<UINT, UINT>>& visited)
{
    auto asNumber = [](const Engine::Value& v) -> double {
        switch (v.kind)
        {
        case Engine::Value::Error: return std::numeric_limits<double>::quiet_NaN();
        case Engine::Value::Blank: return 0.0;
        case Engine::Value::Text:
        {
            double n = 0.0;
            return TextToNumber(v.text, n) ? n : 0.0;
        }
        default:                   return v.num;
        }
    };
    if (g_session)
        return asNumber(Engine::EvalValue(mgr, s, visited));

    Grid32FormulaSession session(mgr, nullptr);
    SessionScope scope(&session);
    double result = 0.0;
    Engine::DriveRoot(mgr, session, [&]() {
        result = asNumber(Engine::ParseWhole(mgr, s, 0, visited));
    });
    return result;
}

// ---- Entry points -----------------------------------------------------------

CFormulator::Result CFormulator::Evaluate(CGrid32Mgr* mgr, const std::wstring& formula,
    const GRIDCELL* self)
{
    // Always a session of its own: `self` belongs to this evaluation only.
    Grid32FormulaSession session(mgr, self);
    SessionScope scope(&session);
    Result result;
    Engine::DriveRoot(mgr, session, [&]() {
        std::set<std::pair<UINT, UINT>> visited;
        result = Engine::EvalSource(mgr, formula, visited);
    });
    return result;
}

CFormulator::Result CFormulator::EvaluateCell(CGrid32Mgr* mgr, UINT row, UINT col)
{
    const CellKey key(row, col);

    // Use the open RecalcPass's session when there is one for this grid and
    // it is idle; otherwise this call is a one-off evaluation of its own.
    Grid32FormulaSession local(mgr, nullptr);
    Grid32FormulaSession* s = g_session;
    const bool reuse = s && s->mgr == mgr && s->self == nullptr
        && s->path.empty() && g_formulaDepth == 0;
    if (!reuse) s = &local;
    SessionScope scope(s);

    auto it = s->memo.find(key);
    if (it != s->memo.end())
        return it->second;

    std::vector<CellKey> pending(1, key);
    Engine::ResolvePending(mgr, *s, pending);
    return s->memo[key];
}

CFormulator::RecalcPass::RecalcPass(CGrid32Mgr* mgr)
    : m_pSession(new Grid32FormulaSession(mgr, nullptr)), m_pOuter(g_session)
{
    g_session = m_pSession;
}

CFormulator::RecalcPass::~RecalcPass()
{
    g_session = m_pOuter;
    delete m_pSession;
}

// Up to 15 significant digits, as Excel shows a General number. The stream's
// default of 6 showed =1234567 as "1.23457e+06" and fed that text to LEN,
// CONCAT and the other text functions (=LEN(123456789) was 11). 15 rather
// than the 17 a round trip needs, so =0.1+0.2 still reads 0.3. The classic
// locale keeps a global C++ locale from adding grouping separators or a
// comma decimal point. Negative zero is written as 0.
std::wstring CFormulator::FormatGeneral(double value)
{
    if (value == 0.0) value = 0.0;     // -0 reads "0", as it does in Excel
    std::wostringstream ss;
    ss.imbue(std::locale::classic());
    ss.precision(15);
    ss << value;
    return ss.str();
}

size_t CFormulator::CellEvaluationCount()
{
    return g_cellEvaluations;
}
