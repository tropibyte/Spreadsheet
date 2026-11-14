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


// --- Formula evaluation helpers ---

namespace {
    // Text-returning functions stash their result in a thread-local slot and
    // return 0.0 as their numeric placeholder. Whoever started the evaluation
    // checks g_textResultValid after parsing and prefers the string when set.
    //
    // The slot is private to this file. Grid32Mgr.cpp used to declare these
    // helpers extern and call them, while they were defined inline here, so
    // nothing guaranteed an out-of-line definition existed: at -O2 GCC inlined
    // every use, emitted no symbol, and the link failed (AUDIT M23).
    // CFormulator::Evaluate now hands a text result back in its Result.
    namespace Grid32FormulaText {
        thread_local std::wstring g_textResult;
        thread_local bool g_textResultValid = false;

        void SetText(const std::wstring& s) {
            g_textResult = s;
            g_textResultValid = true;
        }
        void ResetText() {
            g_textResult.clear();
            g_textResultValid = false;
        }

        // The first error the evaluation ran into (a CFormulator::kError*
        // code), or null. Errors work like the text slot: whatever produces
        // one sets it, and whoever started the evaluation turns it into the
        // result. The first error wins, as evaluation runs left to right.
        thread_local const wchar_t* g_error = nullptr;

        void SetError(const wchar_t* code) {
            if (!g_error) g_error = code;
        }

        // Parks the caller's slots while a nested formula is evaluated and
        // puts them back afterwards. Without it a referenced cell's text
        // result leaked out as the referencing formula's own: with
        // A1 = LEFT("abc",1), =A1+1 displayed "a". A referenced cell's error
        // reaches the referencing formula only through its Result, which
        // decides whether to take it on (COUNT, for one, does not).
        struct Scope {
            std::wstring text;
            bool valid;
            const wchar_t* error;
            Scope() : text(g_textResult), valid(g_textResultValid), error(g_error)
            {
                ResetText();
                g_error = nullptr;
            }
            ~Scope() { g_textResult = text; g_textResultValid = valid; g_error = error; }
            Scope(const Scope&) = delete;
            Scope& operator=(const Scope&) = delete;
        };
    }
}

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

    // A text function's result. The text goes in the slot, where it becomes
    // the formula's result if nothing consumes it; the return value is the
    // text read as a number, which is what arithmetic around the call sees
    // (as it does for a text cell): =1+LEFT("5",1) is 6, not the text "5".
    double TextResult(const std::wstring& text)
    {
        Grid32FormulaText::SetText(text);
        double n = 0.0;
        return ParseNumericText(text, n) ? n : 0.0;
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

    // A range resolves only if both corners are inside the grid; a range that
    // runs off the grid is rejected whole rather than clamped, as a single
    // out-of-grid reference is.
    bool ResolveRange(CGrid32Mgr* mgr, const std::wstring& arg,
        UINT& sr, UINT& sc, UINT& er, UINT& ec)
    {
        return CFormulator::ParseRange(arg, sr, sc, er, ec)
            && InGrid(mgr, sr, sc) && InGrid(mgr, er, ec);
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

    // The functions whose result is text (see TextResult).
    bool IsTextFunction(const std::wstring& name)
    {
        static const wchar_t* const kNames[] = {
            L"LEFT", L"RIGHT", L"MID", L"CONCAT", L"CONCATENATE",
            L"UPPER", L"LOWER", L"TRIM" };
        for (const wchar_t* n : kNames)
            if (_wcsicmp(name.c_str(), n) == 0) return true;
        return false;
    }

    // Comparison operators, lowest precedence in the grammar.
    enum class CompareOp { None, Lt, Gt, Le, Ge, Eq, Ne };
}


// ---- Evaluation internals ---------------------------------------------------
// Nested in CFormulator so it can reach the parser; nothing outside this file
// sees it.
struct CFormulator::Engine
{
    typedef std::set<std::pair<UINT, UINT>> Path;

    // Evaluate formula source with the caller's text slot parked, and report
    // whether the result was text.
    static Result EvalSource(CGrid32Mgr* mgr, const std::wstring& formula, Path& visited)
    {
        Grid32FormulaText::Scope textScope;
        size_t pos = 0;
        // A stored formula never carries its '=', but tolerate one rather than
        // read it as a comparison against nothing.
        if (!formula.empty() && formula[0] == L'=') pos = 1;

        Result passthrough;
        if (BareReference(mgr, formula.substr(pos), visited, passthrough))
            return passthrough;

        double v = ParseWhole(mgr, formula, pos, visited);
        Result r;
        if (Grid32FormulaText::g_error)
        {
            r = ErrorResult(Grid32FormulaText::g_error);
        }
        else if (Grid32FormulaText::g_textResultValid)
        {
            r.isText = true;
            r.text = Grid32FormulaText::g_textResult;
            double n = 0.0;
            r.value = ParseNumericText(r.text, n) ? n : 0.0;
        }
        else if (!std::isfinite(v))
        {
            // An overflow (POWER(10,400)) or a domain error that reached the
            // result: there is no number to show. "inf" and "nan" used to.
            r = ErrorResult(kErrorNum);
        }
        else
        {
            // -0 (=-A1 with A1 blank, =0*-1) is 0: it showed as "-0".
            r.value = v == 0.0 ? 0.0 : v;
        }
        return r;
    }

    static Result ErrorResult(const wchar_t* code)
    {
        Result r;
        r.isError = true;
        r.text = code;
        r.value = std::numeric_limits<double>::quiet_NaN();
        return r;
    }

    // Parse `expr` from `pos` as one whole expression. Anything left over is
    // a syntax error: "2(3)", "1 2", "A1:B2" outside a function, a stray ')'.
    // The parser used to stop at the first thing it did not understand and
    // show what it had read so far, so =12abc showed 12.
    static double ParseWhole(CGrid32Mgr* mgr, const std::wstring& expr, size_t pos, Path& visited)
    {
        const double v = ParseExpression(mgr, expr, pos, visited);
        SkipSpaces(expr, pos);
        if (pos < expr.size())
            Grid32FormulaText::SetError(kErrorSyntax);
        return v;
    }

    // A formula that is nothing but a reference (=A1) takes the referenced
    // cell as it is: a formula's result, text included, or a text cell's
    // text. Text used to reach =A1 from a formula only by accident — it leaked
    // through the shared slot, into =A1+1 as well — and never from a text
    // cell, which showed 0. Arithmetic on text still reads it as a number.
    // Returns false when `expr` is not a bare reference.
    static bool BareReference(CGrid32Mgr* mgr, const std::wstring& expr, Path& visited, Result& out)
    {
        UINT row = 0, col = 0;
        if (!ResolveCellRef(mgr, TrimWs(expr), row, col)) return false;
        out = ReferenceValue(mgr, row, col, visited);
        return true;
    }

    // A reference (inside the grid) as a value: a formula's result (text or an
    // error included), a text cell's text, or a typed or blank cell's number.
    // A cycle reads as an empty numeric result.
    static Result ReferenceValue(CGrid32Mgr* mgr, UINT row, UINT col, Path& visited)
    {
        const CellKey key(row, col);
        const GRIDCELL* cell = mgr->GetCell(row, col);
        Result out;

        if (cell && cell->m_bFormula)
        {
            FormulaCellResult(mgr, key, cell, visited, out);
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
        ReadCell(mgr, key, cell, visited, v, isNumber);
        out.value = isNumber ? v : 0.0;
        return out;
    }

    // Take on a referenced cell's error as this evaluation's own.
    static void PropagateError(const Result& r)
    {
        static const wchar_t* const kCodes[] = { kErrorSyntax, kErrorName, kErrorDiv0, kErrorNum, kErrorValue };
        for (const wchar_t* code : kCodes)
            if (r.text == code) { Grid32FormulaText::SetError(code); return; }
        Grid32FormulaText::SetError(kErrorSyntax);
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

    // What a cell contributes to a formula (S3). Returns false for a blank
    // cell. `isNumber` says whether it takes part in SUM / AVERAGE / MIN / MAX
    // / COUNT; the same reading serves a reference in arithmetic, so =A1+0
    // and =SUM(A1) always agree.
    //
    // Typed cells are read from m_dValue, never from their display text: a
    // currency cell shows "$1,234.50", a percentage "25.00%", a date
    // "2024-01-15", a boolean "TRUE", and std::stod read those as 0, 25, 2024
    // and 0. Text cells are numbers only when the whole text reads as one.
    //
    // A formula cell whose result is an error makes the reading formula an
    // error too (=A1+1, SUM over it). A caller that passes `skippedError`
    // (COUNT) is told instead, and the cell reads as not a number.
    static bool ReadCell(CGrid32Mgr* mgr, const CellKey& key, const GRIDCELL* cell,
        Path& visited, double& value, bool& isNumber, bool* skippedError = nullptr)
    {
        value = 0.0;
        isNumber = false;
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
                if (skippedError) *skippedError = true;
                else PropagateError(r);
                return true;
            }
            double n = 0.0;
            value = r.value;
            isNumber = !r.isText || ParseNumericText(r.text, n);
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
            return true;
        default:
            isNumber = ParseNumericText(cell->m_wsText, value);
            if (!isNumber) value = 0.0;
            return true;
        }
    }

    // A cell read in a text function (LEN, LEFT, CONCAT, ...). Formula cells
    // give their result as it would display under FMT_GENERAL; typed cells
    // under a number format give their value the way the user would edit it
    // ("1234.5", not "$1,234.50"), consistent with ReadCell reading the value
    // rather than the rendering.
    static std::wstring CellAsString(CGrid32Mgr* mgr, const CellKey& key, const GRIDCELL* cell,
        Path& visited)
    {
        if (!cell) return std::wstring();
        if (cell->m_bFormula)
        {
            Result r;
            FormulaCellResult(mgr, key, cell, visited, r);   // a cycle reads as 0 here too
            if (r.isError)
            {
                PropagateError(r);
                return std::wstring();
            }
            return r.isText ? r.text : FormatGeneral(r.value);
        }
        if (cell->m_nFormat != FMT_GENERAL && !cell->m_wsText.empty() &&
            (cell->m_eType == CT_Number || cell->m_eType == CT_Date || cell->m_eType == CT_Boolean))
        {
            return Grid32Detail::FormatEditValue(cell->m_eType, cell->m_dValue, cell->m_wsText);
        }
        return cell->m_wsText;
    }

    // Resolve an argument expression to a string. A bare cell reference reads
    // the cell as CellAsString does; anything else — a quoted literal
    // ("hello", "with ""embedded"" quotes") included — is evaluated, and the
    // text it produced is preferred to its number.
    //
    // A quoted literal used to be cut out here and everything after its
    // closing quote ignored, so UPPER("a"+1) read "a". The parser reads
    // literals itself now, so they go through it like any other expression.
    static std::wstring EvalAsString(CGrid32Mgr* mgr, const std::wstring& arg, Path& visited)
    {
        std::wstring s = TrimWs(arg);
        if (s.empty()) return std::wstring();

        // Bare cell reference (e.g. A1).
        UINT row = 0, col = 0;
        if (ResolveCellRef(mgr, s, row, col))
            return CellAsString(mgr, CellKey(row, col), mgr->GetCell(row, col), visited);

        // Generic expression: evaluate; prefer TLS string if the call chain
        // set one, otherwise stringify the numeric result.
        Grid32FormulaText::ResetText();
        double num = CFormulator::EvalArg(mgr, s, visited);
        if (Grid32FormulaText::g_textResultValid) {
            std::wstring r = Grid32FormulaText::g_textResult;
            Grid32FormulaText::ResetText();
            return r;
        }
        return FormatGeneral(num);
    }

    // One side of a comparison.
    struct Operand
    {
        enum Kind { Number, Text, Blank } kind;
        double num;
        std::wstring text;
    };

    // What the operand just parsed from expr[start, end) is: text if it left
    // text in the slot, blank if it is nothing but a reference to an empty
    // cell, otherwise the number `value`.
    static Operand MakeOperand(CGrid32Mgr* mgr, double value, const std::wstring& expr,
        size_t start, size_t end)
    {
        Operand o{ Operand::Number, value, std::wstring() };
        if (Grid32FormulaText::g_textResultValid)
        {
            o.kind = Operand::Text;
            o.text = Grid32FormulaText::g_textResult;
            return o;
        }
        UINT row = 0, col = 0;
        if (end > start && ResolveCellRef(mgr, TrimWs(expr.substr(start, end - start)), row, col) &&
            IsBlankCell(mgr->GetCell(row, col)))
        {
            o.kind = Operand::Blank;
            o.num = 0.0;
        }
        return o;
    }

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

    // How two values compare:
    //  - numbers, at the 15 significant digits Excel works in (Snap15): an
    //    exact == would make =A1=0.3 false whenever A1 was computed as
    //    0.1+0.2. This used to be a tolerance of 1e-9 times the larger side,
    //    which grew with the numbers: 1000000000 and 1000000001 compared
    //    equal, and two date-times within four seconds of each other;
    //  - text, case-insensitively, as in Excel ("abc" = "ABC");
    //  - a blank cell as whatever the other side is: "" against text, 0
    //    against a number, so =A1="" and =A1=0 are both true for empty A1;
    //  - text against a number: numerically when the text reads as a number
    //    (the engine reads "42" as 42 everywhere else too); otherwise text is
    //    greater than any number, as Excel orders them.
    static bool Compare(CompareOp op, Operand a, Operand b)
    {
        if (a.kind == Operand::Blank && b.kind == Operand::Blank)
            return ApplyCompare(op, 0);
        if (a.kind == Operand::Blank)
            a.kind = b.kind == Operand::Text ? Operand::Text : Operand::Number;   // text "" / number 0
        if (b.kind == Operand::Blank)
            b.kind = a.kind == Operand::Text ? Operand::Text : Operand::Number;

        if (a.kind == Operand::Text && b.kind == Operand::Text)
            return ApplyCompare(op, CompareTextNoCase(a.text, b.text));
        if (a.kind == Operand::Text)
        {
            if (!ParseNumericText(a.text, a.num)) return ApplyCompare(op, 1);
        }
        else if (b.kind == Operand::Text)
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

    // Feed every number an aggregate's arguments hold to `fn`: SUM, AVERAGE,
    // MIN, MAX and COUNT all take any mix of ranges, references and
    // expressions (E4), and agree on what counts.
    //
    //  - A range or reference contributes its populated cells that read as
    //    numbers (ReadCell); blanks and non-numeric text are skipped. One
    //    that runs off the grid contributes nothing, as before.
    //  - Any other argument is evaluated and contributes one value, unless it
    //    is text that does not read as a number. An omitted one ("A1,,A2")
    //    contributes nothing.
    //
    // Errors, in a cell or an expression, make the aggregate an error; with
    // `skipErrors` (COUNT) they are simply not numbers.
    template <class Fn>
    static void ForEachAggregateValue(CGrid32Mgr* mgr, const std::wstring& arglist, Path& visited,
        bool skipErrors, Fn fn)
    {
        for (const std::wstring& raw : SplitArgs(arglist))
        {
            const std::wstring a = TrimWs(raw);
            if (a.empty()) continue;

            UINT sr = 0, sc = 0, er = 0, ec = 0;
            if (CFormulator::ParseRange(a, sr, sc, er, ec))
            {
                if (!ResolveRange(mgr, a, sr, sc, er, ec)) continue;
                ForEachPopulatedCell(mgr, sr, sc, er, ec, [&](const CellKey& key, PGRIDCELL cell) {
                    double v = 0.0;
                    bool isNumber = false, skipped = false;
                    if (ReadCell(mgr, key, cell, visited, v, isNumber, skipErrors ? &skipped : nullptr) &&
                        isNumber && !skipped)
                        fn(v);
                });
                continue;
            }

            // Evaluate with the error slot parked, to see this argument's own.
            const wchar_t* const outerError = Grid32FormulaText::g_error;
            Grid32FormulaText::g_error = nullptr;
            Grid32FormulaText::ResetText();
            const double v = CFormulator::EvalArg(mgr, a, visited);
            const wchar_t* const argError = Grid32FormulaText::g_error;
            double n = 0.0;
            const bool number = !argError && (!Grid32FormulaText::g_textResultValid ||
                ParseNumericText(Grid32FormulaText::g_textResult, n));
            Grid32FormulaText::ResetText();
            Grid32FormulaText::g_error = outerError;
            if (argError && !skipErrors) Grid32FormulaText::SetError(argError);
            if (number) fn(v);
        }
    }

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


void CFormulator::SkipSpaces(const std::wstring& s, size_t& pos) {
    while (pos < s.size() && iswspace(s[pos])) ++pos;
}

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

// The aggregates read their arguments through Engine::ForEachAggregateValue,
// so they agree with each other and with a direct reference about what a cell
// is worth. Blank cells, and text that does not read as a number, are
// skipped — matching Excel, and E6: treating them as 0 would drag MIN to 0 and
// inflate AVERAGE's denominator.
double CFormulator::SumRange(CGrid32Mgr* mgr, const std::wstring& arg,
    std::set<std::pair<UINT, UINT>>& visited)
{
    double total = 0.0;
    Engine::ForEachAggregateValue(mgr, arg, visited, false, [&](double v) { total += v; });
    return total;
}

double CFormulator::AverageRange(CGrid32Mgr* mgr, const std::wstring& arg,
    std::set<std::pair<UINT, UINT>>& visited)
{
    double total = 0.0;
    UINT count = 0;
    Engine::ForEachAggregateValue(mgr, arg, visited, false, [&](double v) { total += v; ++count; });
    // The average of nothing is undefined, as in Excel. It read 0 while the
    // engine had no error values (E6).
    if (count == 0)
    {
        Grid32FormulaText::SetError(kErrorDiv0);
        return 0.0;
    }
    return total / count;
}

double CFormulator::MinMaxRange(CGrid32Mgr* mgr, const std::wstring& arg,
    std::set<std::pair<UINT, UINT>>& visited, bool bMax)
{
    double best = 0.0;
    bool any = false;
    Engine::ForEachAggregateValue(mgr, arg, visited, false, [&](double v) {
        if (!any) { best = v; any = true; }
        else if (bMax) { if (v > best) best = v; }
        else           { if (v < best) best = v; }
    });
    // An all-blank range has no minimum or maximum; Excel reports 0.
    return any ? best : 0.0;
}

double CFormulator::CountRange(CGrid32Mgr* mgr, const std::wstring& arg,
    std::set<std::pair<UINT, UINT>>& visited)
{
    // COUNT tallies the values the other aggregates would use. A formula cell
    // is evaluated to find out (a text result is not a number); the memo
    // makes that a lookup during a recalculation, and a formula that is on
    // the path (COUNT over a range that includes its own cell) still counts,
    // as it always has. Errors are not numbers and, as in Excel, do not make
    // COUNT an error.
    double count = 0.0;
    Engine::ForEachAggregateValue(mgr, arg, visited, true, [&](double) { count += 1.0; });
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
        else if (rhs == 0.0)
        {
            // Not inf or nan, which the cell used to show: an error, as in
            // Excel. (Returning the dividend, before that, was L7.)
            Grid32FormulaText::SetError(kErrorDiv0);
            value = 0.0;
        }
        else
        {
            value /= rhs;
        }
        // The operands were read as numbers; a text operand is not the result.
        Grid32FormulaText::ResetText();
    }
    return value;
}

double CFormulator::EvalArg(CGrid32Mgr* mgr, const std::wstring& s,
    std::set<std::pair<UINT, UINT>>& visited) {
    if (g_session)
    {
        // Part of an evaluation already under way (a function argument). An
        // omitted argument, as in IF(A1,,5), is 0. Anything else must parse
        // whole: ABS(1 2) is a syntax error, not 1.
        if (TrimWs(s).empty()) return 0.0;
        return Engine::ParseWhole(mgr, s, 0, visited);
    }

    // The outermost call: someone evaluating an expression directly, as the
    // tests do. It gets what the grid's own entry points get — a memo, and
    // deep chains resolved rather than truncated. An error comes back as NaN;
    // Evaluate reports which one.
    Grid32FormulaSession session(mgr, nullptr);
    SessionScope scope(&session);
    const std::wstring callerText = Grid32FormulaText::g_textResult;
    const bool callerTextValid = Grid32FormulaText::g_textResultValid;
    const wchar_t* const callerError = Grid32FormulaText::g_error;
    double result = 0.0;
    bool failed = false;
    Engine::DriveRoot(mgr, session, [&]() {
        // A retry starts from the slots as the caller left them, minus any
        // error: this expression's errors are its own.
        Grid32FormulaText::g_textResult = callerText;
        Grid32FormulaText::g_textResultValid = callerTextValid;
        Grid32FormulaText::g_error = nullptr;
        result = Engine::ParseWhole(mgr, s, 0, visited);
        failed = Grid32FormulaText::g_error != nullptr;
    });
    Grid32FormulaText::g_error = callerError;
    return failed ? std::numeric_limits<double>::quiet_NaN() : result;
}

double CFormulator::ParseAdditive(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos,
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
        Grid32FormulaText::ResetText();   // as in ParseTerm
    }
    return value;
}

// Top of the grammar: an additive expression optionally compared against
// another. Comparisons bind less tightly than + - * /, so "A1+1 > B1*2"
// compares the two arithmetic results. Only one comparison is read; a second
// one ("=1=1=1") is left over, which ParseWhole reports as a syntax error
// rather than ignoring it as it used to.
//
// The result is 1.0 (true) or 0.0 (false), which is exactly what IF / IFS /
// AND / OR / NOT already test with "!= 0.0".
//
// Either side may be text (a literal, a text function, a reference to text)
// or a blank cell; Engine::Compare says how those compare.
double CFormulator::ParseExpression(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos,
    std::set<std::pair<UINT, UINT>>& visited) {
    FormulaDepthGuard depth;
    if (depth.overflowed)
    {
        Grid32FormulaText::SetError(kErrorSyntax);   // nested too deeply to evaluate
        return 0.0;
    }

    // Text a sibling operand left in the slot is not this expression's: in
    // LEFT("a",1)*(B1="x"), B1 must not be read as "a".
    Grid32FormulaText::ResetText();
    SkipSpaces(expr, pos);
    const size_t lhsStart = pos;
    double lhs = ParseAdditive(mgr, expr, pos, visited);
    const size_t lhsEnd = pos;

    SkipSpaces(expr, pos);
    if (pos >= expr.size()) return lhs;

    // Two-character operators first, so "<=" isn't read as "<" followed by a
    // stray "=".
    CompareOp op = CompareOp::None;
    if (expr.compare(pos, 2, L"<>") == 0)      { op = CompareOp::Ne; pos += 2; }
    else if (expr.compare(pos, 2, L"<=") == 0) { op = CompareOp::Le; pos += 2; }
    else if (expr.compare(pos, 2, L">=") == 0) { op = CompareOp::Ge; pos += 2; }
    else if (expr[pos] == L'<')                { op = CompareOp::Lt; ++pos; }
    else if (expr[pos] == L'>')                { op = CompareOp::Gt; ++pos; }
    else if (expr[pos] == L'=')                { op = CompareOp::Eq; ++pos; }
    else return lhs;

    const Engine::Operand a = Engine::MakeOperand(mgr, lhs, expr, lhsStart, lhsEnd);
    Grid32FormulaText::ResetText();
    SkipSpaces(expr, pos);
    const size_t rhsStart = pos;
    double rhs = ParseAdditive(mgr, expr, pos, visited);
    const Engine::Operand b = Engine::MakeOperand(mgr, rhs, expr, rhsStart, pos);
    Grid32FormulaText::ResetText();   // a comparison's result is a number

    return Engine::Compare(op, a, b) ? 1.0 : 0.0;
}

double CFormulator::ParseFactor(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos,
    std::set<std::pair<UINT, UINT>>& visited) {
    // Unary +/- below recurses straight back into ParseFactor, which would
    // otherwise sidestep ParseExpression's guard and let "=-----...-1" run the
    // stack out.
    FormulaDepthGuard depth;
    if (depth.overflowed)
    {
        Grid32FormulaText::SetError(kErrorSyntax);
        return 0.0;
    }

    SkipSpaces(expr, pos);
    if (pos >= expr.size())
    {
        // An operand is missing: "=", "=1+", "=(".
        Grid32FormulaText::SetError(kErrorSyntax);
        return 0.0;
    }
    if (expr[pos] == L'(') {
        ++pos;
        double val = ParseExpression(mgr, expr, pos, visited);
        SkipSpaces(expr, pos);
        if (pos < expr.size() && expr[pos] == L')') ++pos;
        else Grid32FormulaText::SetError(kErrorSyntax);   // "=(1+2"
        return val;
    }
    if (expr[pos] == L'"') {
        // A string literal; "" inside it is one quote. Its value is text,
        // which IF can hand on and a comparison can compare (="abc",
        // =IF(A1>3,"yes","no"), =A1="abc"); arithmetic reads it as a number.
        std::wstring literal;
        ++pos;
        bool closed = false;
        while (pos < expr.size())
        {
            if (expr[pos] == L'"')
            {
                if (pos + 1 < expr.size() && expr[pos + 1] == L'"') { literal.push_back(L'"'); pos += 2; continue; }
                ++pos;
                closed = true;
                break;
            }
            literal.push_back(expr[pos++]);
        }
        if (!closed) Grid32FormulaText::SetError(kErrorSyntax);
        return TextResult(literal);
    }
    if (expr[pos] == L'+' || expr[pos] == L'-') {
        bool neg = (expr[pos] == L'-');
        ++pos;
        double val = ParseFactor(mgr, expr, pos, visited);
        if (!neg) return val;
        Grid32FormulaText::ResetText();   // -LEFT("5",1) is the number -5
        return -val;
    }

    size_t start = pos;
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
    std::wstring token = expr.substr(start, pos - start);
    if (token.empty())
    {
        // Something no operand starts with: "=*2", "=)", "=$A$1".
        Grid32FormulaText::SetError(kErrorSyntax);
        return 0.0;
    }

    if (pos < expr.size() && expr[pos] == L'(')
    {
        ++pos; // skip '('
        size_t argStart = pos;
        // Find the matching ')'. Parentheses inside a string literal are
        // text, not structure: without tracking quotes, LEN("a)b") closed at
        // the quoted ')' and measured "a. "" inside a literal is an escaped
        // quote and leaves quote mode unchanged, as in SplitArgs.
        int depth = 1;
        bool inQuote = false;
        while (pos < expr.size() && depth > 0)
        {
            const wchar_t c = expr[pos];
            if (c == L'"') inQuote = !inQuote;
            else if (!inQuote && c == L'(') depth++;
            else if (!inQuote && c == L')') depth--;
            if (depth > 0) ++pos;
        }
        if (pos >= expr.size())
        {
            // No closing ')' ("=SUM(A1:A3"), or an unterminated literal
            // inside the call swallowed it.
            Grid32FormulaText::SetError(kErrorSyntax);
            return 0.0;
        }
        std::wstring arg = expr.substr(argStart, pos - argStart);
        ++pos;  // the ')'

        double res = 0.0;
        if (_wcsicmp(token.c_str(), L"SUM") == 0)
        {
            res = SumRange(mgr, arg, visited);
        }
        else if (_wcsicmp(token.c_str(), L"AVERAGE") == 0)
        {
            res = AverageRange(mgr, arg, visited);
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
            res = sin(EvalArg(mgr, arg, visited));
        }
        else if (_wcsicmp(token.c_str(), L"COS") == 0)
        {
            res = cos(EvalArg(mgr, arg, visited));
        }
        else if (_wcsicmp(token.c_str(), L"TAN") == 0)
        {
            res = tan(EvalArg(mgr, arg, visited));
        }
        else if (_wcsicmp(token.c_str(), L"POWER") == 0)
        {
            // Split at top-level commas only: arg.find(L',') cut
            // POWER(ROUND(2.5,0),2) inside ROUND's own argument list.
            auto args = SplitArgs(arg);
            double base = 0.0, exp = 0.0;
            if (args.size() == 2)
            {
                base = EvalArg(mgr, args[0], visited);
                exp = EvalArg(mgr, args[1], visited);
                res = pow(base, exp);   // no finite result (POWER(0,-1)) is #NUM!
            }
            else
            {
                Grid32FormulaText::SetError(kErrorSyntax);   // POWER(2) read as 0^0 = 1
            }
        }
        else if (_wcsicmp(token.c_str(), L"ROOT") == 0)
        {
            auto args = SplitArgs(arg);
            double val = args.size() >= 1 ? EvalArg(mgr, args[0], visited) : 0.0;
            double n = args.size() >= 2 ? EvalArg(mgr, args[1], visited) : 2.0;
            res = n != 0.0 ? pow(val, 1.0 / n) : 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"LOG") == 0 || _wcsicmp(token.c_str(), L"LN") == 0)
        {
            // Logarithms of zero or a negative number have no value; they
            // showed "-inf" and "nan".
            const double v = EvalArg(mgr, arg, visited);
            if (v > 0.0)
                res = _wcsicmp(token.c_str(), L"LOG") == 0 ? log10(v) : log(v);
            else
                Grid32FormulaText::SetError(kErrorNum);
        }
        else if (_wcsicmp(token.c_str(), L"IF") == 0)
        {
            // IF(cond, then [, else]) — non-zero condition selects `then`.
            auto args = SplitArgs(arg);
            if (args.size() >= 2)
            {
                double cond = EvalArg(mgr, args[0], visited);
                // The condition is consumed here; only the branch taken may
                // hand a text result on.
                Grid32FormulaText::ResetText();
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
                const double cond = EvalArg(mgr, args[i], visited);
                Grid32FormulaText::ResetText();
                if (cond != 0.0)
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
            // Half away from zero, on the decimal the user sees: 1.005 rounds
            // to 1.01, not the 1.00 its binary value (1.00499999...) gives.
            res = RoundToDigits(v, d, RoundMode::HalfAwayFromZero);
        }
        else if (_wcsicmp(token.c_str(), L"CEILING") == 0)
        {
            auto args = SplitArgs(arg);
            double v = args.size() >= 1 ? EvalArg(mgr, args[0], visited) : 0.0;
            double sig = args.size() >= 2 ? EvalArg(mgr, args[1], visited) : 1.0;
            res = RoundToMultiple(v, sig, true);
        }
        else if (_wcsicmp(token.c_str(), L"FLOOR") == 0)
        {
            auto args = SplitArgs(arg);
            double v = args.size() >= 1 ? EvalArg(mgr, args[0], visited) : 0.0;
            double sig = args.size() >= 2 ? EvalArg(mgr, args[1], visited) : 1.0;
            res = RoundToMultiple(v, sig, false);
        }
        else if (_wcsicmp(token.c_str(), L"MOD") == 0)
        {
            // MOD(dividend, divisor)
            auto args = SplitArgs(arg);
            double d1 = args.size() >= 1 ? EvalArg(mgr, args[0], visited) : 0.0;
            double d2 = args.size() >= 2 ? EvalArg(mgr, args[1], visited) : 0.0;
            // The result takes the divisor's sign, as in Excel:
            // MOD(-3,2) = 1 and MOD(3,-2) = -1. std::fmod takes the
            // dividend's, and gave -1 and 1.
            res = 0.0;
            if (d2 == 0.0)
            {
                Grid32FormulaText::SetError(kErrorDiv0);   // as in Excel; it read 0
            }
            else
            {
                res = std::fmod(d1, d2);
                if (res == 0.0)
                    res = 0.0;      // fmod(-6,3) is -0.0, which shows as "-0"
                else if ((res < 0.0) != (d2 < 0.0))
                    res += d2;
            }
        }
        else if (_wcsicmp(token.c_str(), L"INT") == 0)
        {
            // INT rounds toward -inf, like Excel.
            res = std::floor(EvalArg(mgr, arg, visited));
        }
        else if (_wcsicmp(token.c_str(), L"TRUNC") == 0)
        {
            // TRUNC(value [, digits]) drops digits toward zero. The digits
            // argument used to be ignored; now that an argument must parse
            // whole it would have been a syntax error instead.
            auto args = SplitArgs(arg);
            double v = args.size() >= 1 ? EvalArg(mgr, args[0], visited) : 0.0;
            double d = args.size() >= 2 ? EvalArg(mgr, args[1], visited) : 0.0;
            res = RoundToDigits(v, d, RoundMode::TowardZero);
        }
        else if (_wcsicmp(token.c_str(), L"SQRT") == 0)
        {
            double v = EvalArg(mgr, arg, visited);
            if (v >= 0) res = std::sqrt(v);
            else Grid32FormulaText::SetError(kErrorNum);   // it read 0
        }
        // -------- Text functions -----------------------------------------
        else if (_wcsicmp(token.c_str(), L"LEN") == 0)
        {
            std::wstring s = Engine::EvalAsString(mgr, arg, visited);
            res = (double)s.size();
        }
        else if (_wcsicmp(token.c_str(), L"LEFT") == 0)
        {
            auto args = SplitArgs(arg);
            std::wstring s = args.size() >= 1 ? Engine::EvalAsString(mgr, args[0], visited) : std::wstring();
            const size_t n = args.size() >= 2 ? ClampCount(EvalArg(mgr, args[1], visited), s.size())
                                              : (s.empty() ? 0 : 1);
            res = TextResult(s.substr(0, n));
        }
        else if (_wcsicmp(token.c_str(), L"RIGHT") == 0)
        {
            auto args = SplitArgs(arg);
            std::wstring s = args.size() >= 1 ? Engine::EvalAsString(mgr, args[0], visited) : std::wstring();
            const size_t n = args.size() >= 2 ? ClampCount(EvalArg(mgr, args[1], visited), s.size())
                                              : (s.empty() ? 0 : 1);
            res = TextResult(s.substr(s.size() - n));
        }
        else if (_wcsicmp(token.c_str(), L"MID") == 0)
        {
            // MID(text, start, length) — start is 1-based per Excel
            auto args = SplitArgs(arg);
            std::wstring s = args.size() >= 1 ? Engine::EvalAsString(mgr, args[0], visited) : std::wstring();
            const double start = args.size() >= 2 ? EvalArg(mgr, args[1], visited) : 1.0;
            const double length = args.size() >= 3 ? EvalArg(mgr, args[2], visited) : 0.0;
            // A start below 1 (or NaN) is read as 1; past the end there is
            // nothing.
            size_t pos1 = ClampCount(start, s.size() + 1);
            if (pos1 == 0) pos1 = 1;
            const size_t idx = pos1 - 1;
            if (idx >= s.size()) res = TextResult(std::wstring());
            else res = TextResult(s.substr(idx, ClampCount(length, s.size() - idx)));
        }
        else if (_wcsicmp(token.c_str(), L"CONCAT") == 0 ||
                 _wcsicmp(token.c_str(), L"CONCATENATE") == 0)
        {
            auto args = SplitArgs(arg);
            std::wstring out;
            for (auto& a : args) out += Engine::EvalAsString(mgr, a, visited);
            res = TextResult(out);
        }
        else if (_wcsicmp(token.c_str(), L"UPPER") == 0)
        {
            std::wstring s = Engine::EvalAsString(mgr, arg, visited);
            std::transform(s.begin(), s.end(), s.begin(), ::towupper);
            res = TextResult(s);
        }
        else if (_wcsicmp(token.c_str(), L"LOWER") == 0)
        {
            std::wstring s = Engine::EvalAsString(mgr, arg, visited);
            std::transform(s.begin(), s.end(), s.begin(), ::towlower);
            res = TextResult(s);
        }
        else if (_wcsicmp(token.c_str(), L"TRIM") == 0)
        {
            // Excel TRIM removes leading and trailing spaces and collapses
            // internal runs of spaces to one. Spaces only (U+0020): tabs,
            // line breaks and non-breaking spaces are text and stay, where
            // iswspace used to turn each into a space ("a<tab>b" -> "a b").
            std::wstring s = Engine::EvalAsString(mgr, arg, visited);
            std::wstring out;
            bool pendingSpace = false;
            for (wchar_t c : s) {
                if (c == L' ') {
                    pendingSpace = !out.empty();
                } else {
                    if (pendingSpace) out.push_back(L' ');
                    pendingSpace = false;
                    out.push_back(c);
                }
            }
            res = TextResult(out);
        }
        else if (_wcsicmp(token.c_str(), L"VALUE") == 0)
        {
            // VALUE(text): text read as a number, #VALUE! when it is not one.
            // Numeric text is what the engine reads as a number everywhere
            // (ParseNumericText: " 42 ", "1e3", "1,234.50", "$1,234.50",
            // "25%" = 0.25); date text is what typing into a cell accepts
            // (TryParseDate: ISO and US), and reads as its serial, as in
            // Excel. A number passes through and a blank cell is 0. Unlike
            // =A1+0, which reads any text that is not a number as 0, text
            // that is neither fails: "abc", "", "12abc", "1,2".
            if (TrimWs(arg).empty())
                Grid32FormulaText::SetError(kErrorSyntax);  // VALUE() has nothing to read
            Grid32FormulaText::ResetText();
            const double v = EvalArg(mgr, arg, visited);
            res = v;
            if (Grid32FormulaText::g_textResultValid)
            {
                const std::wstring text = Grid32FormulaText::g_textResult;
                double n = 0.0;
                if (ParseNumericText(text, n) || Grid32Detail::TryParseDate(TrimWs(text), n))
                    res = n;
                else
                {
                    res = 0.0;
                    Grid32FormulaText::SetError(kErrorValue);
                }
            }
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
            const double y = args.size() >= 1 ? EvalArg(mgr, args[0], visited) : 1900.0;
            const double m = args.size() >= 2 ? EvalArg(mgr, args[1], visited) : 1.0;
            const double d = args.size() >= 3 ? EvalArg(mgr, args[2], visited) : 1.0;
            // As Excel's DATE: the parts are truncated to whole numbers; a
            // year 0..1899 is that many years after 1900 (DATE(24,1,1) is
            // 1924-01-01 — it read as the year 24, a negative serial); a
            // month or day out of range rolls over (DATE(2024,13,1) is
            // 2025-01-01); and the result must be a date Excel can hold,
            // 0 (1900-01-00) to 9999-12-31, or it is #NUM!. The size check
            // comes first: the parts go to int, and out of its range that
            // cast is undefined.
            const double kMaxPart = 1e9;
            if (!(std::fabs(y) <= kMaxPart && std::fabs(m) <= kMaxPart && std::fabs(d) <= kMaxPart))
            {
                Grid32FormulaText::SetError(kErrorNum);
            }
            else
            {
                int year = (int)y;
                if (year >= 0 && year < 1900) year += 1900;
                const double serial = (year < 0 || year > 9999)
                    ? -1.0 : Grid32Detail::DateToSerial(year, (int)m, (int)d);
                if (serial < 0.0 || serial > Grid32Detail::kMaxDateSerial)
                    Grid32FormulaText::SetError(kErrorNum);
                else
                    res = serial;
            }
        }
        else if (_wcsicmp(token.c_str(), L"YEAR") == 0 ||
                 _wcsicmp(token.c_str(), L"MONTH") == 0 ||
                 _wcsicmp(token.c_str(), L"DAY") == 0)
        {
            double serial = EvalArg(mgr, arg, visited);
            // A serial outside Excel's dates, 0 (1900-01-00) to 2958465
            // (9999-12-31), has no year, month or day: #NUM!, as in Excel.
            // YEAR(1e15) used to read 2012841336 and YEAR(-5) 1899. 0 and
            // 60 are the days Excel shows as 1900-01-00 and 1900-02-29.
            if (!(serial >= 0.0 && serial < Grid32Detail::kMaxDateSerial + 1.0))
            {
                Grid32FormulaText::SetError(kErrorNum);
            }
            else
            {
                int y, m, d;
                SerialToYMD(serial, y, m, d);
                if (_wcsicmp(token.c_str(), L"YEAR") == 0)       res = (double)y;
                else if (_wcsicmp(token.c_str(), L"MONTH") == 0) res = (double)m;
                else                                              res = (double)d;
            }
        }
        else if (_wcsicmp(token.c_str(), L"TRUE") == 0)
        {
            res = 1.0;
        }
        else if (_wcsicmp(token.c_str(), L"FALSE") == 0)
        {
            res = 0.0;
        }
        else
        {
            // Not a function this engine knows. It used to read as 0.
            Grid32FormulaText::SetError(iswalpha(token[0]) ? kErrorName : kErrorSyntax);
        }

        // A function that returns a number consumes any text its arguments
        // produced: =ABS(LEFT("-5",2)) is 5, not the text "-5". Only the text
        // functions, and IF/IFS handing on the branch they took, leave a
        // text result in the slot.
        if (!IsTextFunction(token) &&
            _wcsicmp(token.c_str(), L"IF") != 0 && _wcsicmp(token.c_str(), L"IFS") != 0)
        {
            Grid32FormulaText::ResetText();
        }
        return res;
    }
    else
    {
        // A number must be one whole literal: =12abc read as 12.
        if (IsAsciiDigit(token[0]) || token[0] == L'.')
        {
            if (!IsNumberLiteral(token))
            {
                Grid32FormulaText::SetError(kErrorSyntax);
                return 0.0;
            }
            // IsNumberLiteral has checked it is ASCII. from_chars, unlike
            // stod, ignores the C locale's decimal separator.
            std::string narrow;
            narrow.reserve(token.size());
            for (wchar_t c : token) narrow.push_back(static_cast<char>(c));
            double v = 0.0;
            const std::from_chars_result r = std::from_chars(narrow.data(), narrow.data() + narrow.size(), v);
            if (r.ec != std::errc() || !std::isfinite(v))
            {
                Grid32FormulaText::SetError(kErrorNum);   // 1e999
                return 0.0;
            }
            return v;
        }

        UINT row = 0, col = 0;
        if (ParseCellRef(token, row, col))
        {
            // A reference outside the grid does not resolve (see InGrid) and
            // reads as 0, as it always has.
            if (!InGrid(mgr, row, col)) return 0.0;

            // The cell as a value: text stays text (for a comparison, or for
            // IF to hand on), an error is taken on, a number is a number.
            const Result r = Engine::ReferenceValue(mgr, row, col, visited);
            if (r.isError)
            {
                Engine::PropagateError(r);
                return 0.0;
            }
            return r.isText ? TextResult(r.text) : r.value;
        }

        if (_wcsicmp(token.c_str(), L"TRUE") == 0) return 1.0;
        if (_wcsicmp(token.c_str(), L"FALSE") == 0) return 0.0;

        // A name that is neither a reference nor a keyword. It used to read
        // as 0.
        Grid32FormulaText::SetError(kErrorName);
        return 0.0;
    }
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
