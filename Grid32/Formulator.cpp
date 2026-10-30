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

        // Parks the caller's slot while a nested formula is evaluated and puts
        // it back afterwards. Without it a referenced cell's text result leaked
        // out as the referencing formula's own: with A1 = LEFT("abc",1),
        // =A1+1 displayed "a".
        struct Scope {
            std::wstring text;
            bool valid;
            Scope() : text(g_textResult), valid(g_textResultValid) { ResetText(); }
            ~Scope() { g_textResult = text; g_textResultValid = valid; }
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

        double v = std::strtod(core.c_str(), nullptr);
        if (!std::isfinite(v)) return false;   // "1e999" is not a number a cell can hold
        if (percent) v /= 100.0;
        out = neg ? -v : v;
        return true;
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

        double v = ParseExpression(mgr, formula, pos, visited);
        Result r;
        if (Grid32FormulaText::g_textResultValid)
        {
            r.isText = true;
            r.text = Grid32FormulaText::g_textResult;
            double n = 0.0;
            r.value = ParseNumericText(r.text, n) ? n : 0.0;
        }
        else
        {
            r.value = v;
        }
        return r;
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
        const CellKey key(row, col);
        const GRIDCELL* cell = mgr->GetCell(row, col);
        out = Result();

        if (cell && cell->m_bFormula)
        {
            FormulaCellResult(mgr, key, cell, visited, out);   // a cycle: an empty numeric result
            return true;
        }
        if (cell && !cell->m_wsText.empty() && cell->m_eType != CT_Number &&
            cell->m_eType != CT_Date && cell->m_eType != CT_Boolean)
        {
            out.isText = true;
            out.text = cell->m_wsText;
            double n = 0.0;
            out.value = ParseNumericText(out.text, n) ? n : 0.0;
            return true;
        }
        // Blank or typed: its value, as any other reference reads it.
        double v = 0.0;
        bool isNumber = false;
        ReadCell(mgr, key, cell, visited, v, isNumber);
        out.value = isNumber ? v : 0.0;
        return true;
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
    static bool ReadCell(CGrid32Mgr* mgr, const CellKey& key, const GRIDCELL* cell,
        Path& visited, double& value, bool& isNumber)
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
            return r.isText ? r.text : FormatGeneral(r.value);
        }
        if (cell->m_nFormat != FMT_GENERAL && !cell->m_wsText.empty() &&
            (cell->m_eType == CT_Number || cell->m_eType == CT_Date || cell->m_eType == CT_Boolean))
        {
            return Grid32Detail::FormatEditValue(cell->m_eType, cell->m_dValue, cell->m_wsText);
        }
        return cell->m_wsText;
    }

    // Resolve an argument expression to a string. Recognizes quoted literals
    // ("hello", "with ""embedded"" quotes"), cell references, and falls back
    // to numeric evaluation (consulting the text-result TLS in case a
    // nested call produced a string).
    static std::wstring EvalAsString(CGrid32Mgr* mgr, const std::wstring& arg, Path& visited)
    {
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
    size_t i = 0;
    unsigned long long c = 0;
    while (i < token.size() && iswalpha(token[i])) {
        c = c * 26 + (unsigned long long)(towupper(token[i]) - L'A' + 1);
        if (c > 0xFFFFFFFFull) return false;   // column index overflow
        ++i;
    }
    if (i == 0 || i >= token.size()) return false;

    // The remainder must be digits and nothing else. std::stoul stops at the
    // first non-digit without complaining, so without this check "A1:B10"
    // parses as a valid reference to A1 and every range collapses to its
    // first cell; "A1B" would likewise be accepted as A1.
    for (size_t j = i; j < token.size(); ++j)
        if (!iswdigit(token[j])) return false;

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

// The range aggregates all read cells through Engine::ReadCell, so they agree
// with each other and with a direct reference about what a cell is worth.
// Blank cells, and text that does not read as a number, are skipped — matching
// Excel, and E6: treating them as 0 would drag MIN to 0 and inflate AVERAGE's
// denominator. (SUM adding 0 for them would be harmless, but it skips them by
// the same rule.)
double CFormulator::SumRange(CGrid32Mgr* mgr, const std::wstring& arg,
    std::set<std::pair<UINT, UINT>>& visited)
{
    UINT sr = 0, sc = 0, er = 0, ec = 0;
    if (!ResolveRange(mgr, arg, sr, sc, er, ec))
        return 0.0;

    double total = 0.0;
    ForEachPopulatedCell(mgr, sr, sc, er, ec, [&](const CellKey& key, PGRIDCELL cell) {
        double v = 0.0;
        bool isNumber = false;
        if (Engine::ReadCell(mgr, key, cell, visited, v, isNumber) && isNumber)
            total += v;
    });
    return total;
}

double CFormulator::AverageRange(CGrid32Mgr* mgr, const std::wstring& arg,
    std::set<std::pair<UINT, UINT>>& visited)
{
    UINT sr = 0, sc = 0, er = 0, ec = 0;
    if (!ResolveRange(mgr, arg, sr, sc, er, ec))
        return 0.0;

    double total = 0.0;
    UINT count = 0;
    ForEachPopulatedCell(mgr, sr, sc, er, ec, [&](const CellKey& key, PGRIDCELL cell) {
        double v = 0.0;
        bool isNumber = false;
        if (Engine::ReadCell(mgr, key, cell, visited, v, isNumber) && isNumber)
        {
            total += v;
            ++count;
        }
    });
    return count ? total / count : 0.0;
}

double CFormulator::MinMaxRange(CGrid32Mgr* mgr, const std::wstring& arg,
    std::set<std::pair<UINT, UINT>>& visited, bool bMax)
{
    UINT sr = 0, sc = 0, er = 0, ec = 0;
    if (!ResolveRange(mgr, arg, sr, sc, er, ec))
        return 0.0;

    double best = 0.0;
    bool any = false;
    ForEachPopulatedCell(mgr, sr, sc, er, ec, [&](const CellKey& key, PGRIDCELL cell) {
        double v = 0.0;
        bool isNumber = false;
        if (!Engine::ReadCell(mgr, key, cell, visited, v, isNumber) || !isNumber)
            return;
        if (!any) { best = v; any = true; }
        else if (bMax) { if (v > best) best = v; }
        else           { if (v < best) best = v; }
    });
    // An all-blank range has no minimum or maximum; report 0 rather than
    // leaking an infinity into the sheet.
    return any ? best : 0.0;
}

double CFormulator::CountRange(CGrid32Mgr* mgr, const std::wstring& arg,
    std::set<std::pair<UINT, UINT>>& visited)
{
    UINT sr = 0, sc = 0, er = 0, ec = 0;
    if (!ResolveRange(mgr, arg, sr, sc, er, ec))
        return 0.0;

    // COUNT tallies the cells the other aggregates would use. A formula cell
    // is evaluated to find out (a text result is not a number); the memo
    // makes that a lookup during a recalculation, and a formula that is on
    // the path (COUNT over a range that includes its own cell) still counts,
    // as it always has.
    double count = 0.0;
    ForEachPopulatedCell(mgr, sr, sc, er, ec, [&](const CellKey& key, PGRIDCELL cell) {
        double v = 0.0;
        bool isNumber = false;
        if (Engine::ReadCell(mgr, key, cell, visited, v, isNumber) && isNumber)
            count += 1.0;
    });
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
    if (g_session)
    {
        // Part of an evaluation already under way (a function argument).
        size_t p = 0;
        return ParseExpression(mgr, s, p, visited);
    }

    // The outermost call: someone evaluating an expression directly, as the
    // tests do. It gets what the grid's own entry points get — a memo, and
    // deep chains resolved rather than truncated.
    Grid32FormulaSession session(mgr, nullptr);
    SessionScope scope(&session);
    const std::wstring callerText = Grid32FormulaText::g_textResult;
    const bool callerTextValid = Grid32FormulaText::g_textResultValid;
    double result = 0.0;
    Engine::DriveRoot(mgr, session, [&]() {
        // A retry starts from the text slot as the caller left it.
        Grid32FormulaText::g_textResult = callerText;
        Grid32FormulaText::g_textResultValid = callerTextValid;
        size_t p = 0;
        result = ParseExpression(mgr, s, p, visited);
    });
    return result;
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
    }
    return value;
}

// Top of the grammar: an additive expression optionally compared against
// another. Comparisons bind less tightly than + - * /, so "A1+1 > B1*2"
// compares the two arithmetic results. Excel's comparisons don't chain, so a
// single optional operator is the whole rule.
//
// The result is 1.0 (true) or 0.0 (false), which is exactly what IF / IFS /
// AND / OR / NOT already test with "!= 0.0".
double CFormulator::ParseExpression(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos,
    std::set<std::pair<UINT, UINT>>& visited) {
    FormulaDepthGuard depth;
    if (depth.overflowed) return 0.0;

    double lhs = ParseAdditive(mgr, expr, pos, visited);

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

    double rhs = ParseAdditive(mgr, expr, pos, visited);

    // Compare with a small relative tolerance. Cell values arrive as doubles
    // parsed from text, so an exact == would make =A1=0.3 false whenever A1
    // was computed as 0.1+0.2.
    double scale = std::fabs(lhs) > std::fabs(rhs) ? std::fabs(lhs) : std::fabs(rhs);
    if (!(scale > 1.0)) scale = 1.0;          // also pins NaN to 1.0
    const double eps = 1e-9 * scale;

    switch (op)
    {
    case CompareOp::Lt: return (lhs < rhs - eps) ? 1.0 : 0.0;
    case CompareOp::Gt: return (lhs > rhs + eps) ? 1.0 : 0.0;
    case CompareOp::Le: return (lhs <= rhs + eps) ? 1.0 : 0.0;
    case CompareOp::Ge: return (lhs >= rhs - eps) ? 1.0 : 0.0;
    case CompareOp::Eq: return (std::fabs(lhs - rhs) <= eps) ? 1.0 : 0.0;
    case CompareOp::Ne: return (std::fabs(lhs - rhs) > eps) ? 1.0 : 0.0;
    default:            return lhs;
    }
}

double CFormulator::ParseFactor(CGrid32Mgr* mgr, const std::wstring& expr, size_t& pos,
    std::set<std::pair<UINT, UINT>>& visited) {
    // Unary +/- below recurses straight back into ParseFactor, which would
    // otherwise sidestep ParseExpression's guard and let "=-----...-1" run the
    // stack out.
    FormulaDepthGuard depth;
    if (depth.overflowed) return 0.0;

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
            std::wstring s = Engine::EvalAsString(mgr, arg, visited);
            res = (double)s.size();
        }
        else if (_wcsicmp(token.c_str(), L"LEFT") == 0)
        {
            auto args = SplitArgs(arg);
            std::wstring s = args.size() >= 1 ? Engine::EvalAsString(mgr, args[0], visited) : std::wstring();
            int n = args.size() >= 2 ? (int)EvalArg(mgr, args[1], visited) : 1;
            if (n < 0) n = 0;
            if ((size_t)n > s.size()) n = (int)s.size();
            Grid32FormulaText::SetText(s.substr(0, n));
            res = 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"RIGHT") == 0)
        {
            auto args = SplitArgs(arg);
            std::wstring s = args.size() >= 1 ? Engine::EvalAsString(mgr, args[0], visited) : std::wstring();
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
            std::wstring s = args.size() >= 1 ? Engine::EvalAsString(mgr, args[0], visited) : std::wstring();
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
            for (auto& a : args) out += Engine::EvalAsString(mgr, a, visited);
            Grid32FormulaText::SetText(out);
            res = 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"UPPER") == 0)
        {
            std::wstring s = Engine::EvalAsString(mgr, arg, visited);
            std::transform(s.begin(), s.end(), s.begin(), ::towupper);
            Grid32FormulaText::SetText(s);
            res = 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"LOWER") == 0)
        {
            std::wstring s = Engine::EvalAsString(mgr, arg, visited);
            std::transform(s.begin(), s.end(), s.begin(), ::towlower);
            Grid32FormulaText::SetText(s);
            res = 0.0;
        }
        else if (_wcsicmp(token.c_str(), L"TRIM") == 0)
        {
            // Excel TRIM removes leading/trailing spaces and collapses
            // internal runs of spaces to single spaces.
            std::wstring s = Engine::EvalAsString(mgr, arg, visited);
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
        // A reference outside the grid does not resolve (see InGrid) and
        // falls through to the numeric parse below, which reads it as 0.
        UINT row = 0, col = 0;
        if (ResolveCellRef(mgr, token, row, col))
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
    if (!mgr) return 0.0;
    double value = 0.0;
    bool isNumber = false;
    Engine::ReadCell(mgr, CellKey(row, col), mgr->GetCell(row, col), visited, value, isNumber);
    return isNumber ? value : 0.0;
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

std::wstring CFormulator::FormatGeneral(double value)
{
    std::wstringstream ss;
    ss << value;
    return ss.str();
}

size_t CFormulator::CellEvaluationCount()
{
    return g_cellEvaluations;
}
