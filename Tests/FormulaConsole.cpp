// Formula console — an interactive shell over the formula engine, for trying
// formulas by hand and for writing self-checking scripts.
//
//   FormulaConsole                     interactive
//   FormulaConsole script.txt ...      run scripts, exit 1 if an `expect` failed
//   FormulaConsole -i script.txt       run scripts, then stay interactive
//   FormulaConsole < script.txt        piped input runs as a script
//
// Type `help` at the prompt for the commands.
//
// Like FormulatorTests, this links Formulator.cpp against stub CGrid32Mgr
// members (at the bottom of this file) rather than Grid32Mgr.cpp, so it builds
// the same way with g++ and with cl.exe: Tests/linux/Makefile and
// Tests/run_formula_console.bat.
//
// What it mirrors from the grid, so what you see here is what a cell shows:
//   - typing into a cell:   Grid32Detail::InferType + FormatCellValue, as
//                           CGrid32Mgr::SetCellText does;
//   - recalculation:        one CFormulator::RecalcPass over every formula
//                           cell, as CGrid32Mgr::RecalculateFormulas does;
//   - a formula's display:  RenderFormulaResult below, a copy of
//                           ApplyFormulaResult in Grid32Mgr.cpp (static
//                           there, so it cannot be called from here). Keep the
//                           two in step.
// What it does not have: undo, merges, sorting, fills, files — anything that
// is the grid rather than the engine.

#define _CRT_SECURE_NO_WARNINGS
#include "../Grid32/pch.h"
#include "../Grid32/grid32_internal.h"
#include "../Grid32/Formulator.h"

#include <cstdio>
#include <cstring>
#include <fstream>
#include <iostream>
#include <map>
#include <string>
#include <vector>

#ifdef _WIN32
#include <io.h>
#define CONSOLE_ISATTY(fd) _isatty(fd)
#define CONSOLE_FILENO(f) _fileno(f)
#else
#include <unistd.h>
#define CONSOLE_ISATTY(fd) isatty(fd)
#define CONSOLE_FILENO(f) fileno(f)
#endif

namespace {

typedef std::pair<UINT, UINT> CellKey;

// The host's grid: 702 columns (A..ZZ) by 4096 rows, as in FormulatorTests.
const UINT kSheetCols = 702;
const UINT kSheetRows = 4096;

// ---- UTF-8 <-> wide --------------------------------------------------------
// The console speaks UTF-8 (on Windows, main() switches the code pages); the
// engine speaks wchar_t, which is UTF-16 on Windows and UTF-32 elsewhere.

std::wstring Widen(const std::string& s)
{
    std::wstring out;
    for (size_t i = 0; i < s.size();)
    {
        unsigned char c = (unsigned char)s[i];
        unsigned long cp = 0xFFFD;
        size_t len = 1;
        if (c < 0x80) { cp = c; }
        else if ((c >> 5) == 0x6 && i + 1 < s.size()) { cp = ((c & 0x1Ful) << 6) | (s[i + 1] & 0x3F); len = 2; }
        else if ((c >> 4) == 0xE && i + 2 < s.size()) { cp = ((c & 0x0Ful) << 12) | ((s[i + 1] & 0x3Ful) << 6) | (s[i + 2] & 0x3F); len = 3; }
        else if ((c >> 3) == 0x1E && i + 3 < s.size()) { cp = ((c & 0x07ul) << 18) | ((s[i + 1] & 0x3Ful) << 12) | ((s[i + 2] & 0x3Ful) << 6) | (s[i + 3] & 0x3F); len = 4; }
        i += len;
        if (sizeof(wchar_t) == 2 && cp > 0xFFFF)
        {
            cp -= 0x10000;
            out.push_back((wchar_t)(0xD800 + (cp >> 10)));
            out.push_back((wchar_t)(0xDC00 + (cp & 0x3FF)));
        }
        else
        {
            out.push_back((wchar_t)cp);
        }
    }
    return out;
}

std::string Narrow(const std::wstring& s)
{
    std::string out;
    for (size_t i = 0; i < s.size(); ++i)
    {
        unsigned long cp = (unsigned long)s[i];
        if (sizeof(wchar_t) == 2 && cp >= 0xD800 && cp < 0xDC00 && i + 1 < s.size())
        {
            cp = 0x10000 + ((cp - 0xD800) << 10) + ((unsigned long)s[i + 1] - 0xDC00);
            ++i;
        }
        if (cp < 0x80) out.push_back((char)cp);
        else if (cp < 0x800) { out.push_back((char)(0xC0 | (cp >> 6))); out.push_back((char)(0x80 | (cp & 0x3F))); }
        else if (cp < 0x10000) { out.push_back((char)(0xE0 | (cp >> 12))); out.push_back((char)(0x80 | ((cp >> 6) & 0x3F))); out.push_back((char)(0x80 | (cp & 0x3F))); }
        else { out.push_back((char)(0xF0 | (cp >> 18))); out.push_back((char)(0x80 | ((cp >> 12) & 0x3F))); out.push_back((char)(0x80 | ((cp >> 6) & 0x3F))); out.push_back((char)(0x80 | (cp & 0x3F))); }
    }
    return out;
}

// Display width of UTF-8 text, for lining up the table (one column per code
// point; good enough for a test console).
size_t Width(const std::string& s)
{
    size_t n = 0;
    for (unsigned char c : s)
        if ((c & 0xC0) != 0x80) ++n;
    return n;
}

std::string Trim(const std::string& s)
{
    size_t a = 0, b = s.size();
    while (a < b && (s[a] == ' ' || s[a] == '\t' || s[a] == '\r' || s[a] == '\n')) ++a;
    while (b > a && (s[b - 1] == ' ' || s[b - 1] == '\t' || s[b - 1] == '\r' || s[b - 1] == '\n')) --b;
    return s.substr(a, b - a);
}

std::string Lower(std::string s)
{
    for (char& c : s) if (c >= 'A' && c <= 'Z') c = (char)(c - 'A' + 'a');
    return s;
}

// Split off the first whitespace-delimited word.
std::string FirstWord(const std::string& s, std::string& rest)
{
    size_t i = 0;
    while (i < s.size() && s[i] != ' ' && s[i] != '\t') ++i;
    rest = Trim(s.substr(i));
    return s.substr(0, i);
}

std::string ColumnName(UINT col)
{
    std::string name;
    unsigned long long c = (unsigned long long)col + 1;
    while (c > 0)
    {
        --c;
        name.insert(name.begin(), (char)('A' + c % 26));
        c /= 26;
    }
    return name;
}

std::string RefName(const CellKey& k)
{
    return ColumnName(k.second) + std::to_string(k.first + 1);
}

// ---- number formats ------------------------------------------------------------

struct FormatName { const char* name; UINT fmt; };
const FormatName kFormats[] = {
    { "general", FMT_GENERAL }, { "number", FMT_NUMBER }, { "integer", FMT_INTEGER },
    { "currency", FMT_CURRENCY }, { "percent", FMT_PERCENT }, { "date", FMT_DATE_ISO },
    { "usdate", FMT_DATE_US }, { "time", FMT_TIME }, { "thousands", FMT_THOUSANDS },
};

const char* FormatToName(UINT fmt)
{
    for (const FormatName& f : kFormats) if (f.fmt == fmt) return f.name;
    return "?";
}

bool NameToFormat(const std::string& name, UINT& fmt)
{
    for (const FormatName& f : kFormats)
        if (Lower(name) == f.name) { fmt = f.fmt; return true; }
    return false;
}

const char* TypeName(CellType t)
{
    switch (t)
    {
    case CT_Text:    return "text";
    case CT_Number:  return "number";
    case CT_Date:    return "date";
    case CT_Boolean: return "boolean";
    case CT_Formula: return "formula";
    case CT_Error:   return "error";
    default:         return "?";
    }
}

// What kind of value a Result is, as the console reports it.
const char* ResultKind(const CFormulator::Result& r)
{
    return r.isError ? "error" : r.isText ? "text" : r.isBoolean ? "boolean" : r.isDate ? "date" : "number";
}

// What a cell shows for a Result. A copy of ApplyFormulaResult in
// Grid32Mgr.cpp — keep the two in step.
void RenderFormulaResult(PGRIDCELL pCell, const CFormulator::Result& r)
{
    // A date result makes the cell a date (CT_Date, still m_bFormula), so
    // sorting and export treat it as the date it shows; any other result
    // keeps CT_Formula.
    const bool isDate = r.isDate && !r.isText && !r.isError;
    const bool isBoolean = r.isBoolean && !r.isText && !r.isError;
    pCell->m_eType = isDate ? CT_Date : isBoolean ? CT_Boolean : CT_Formula;
    pCell->m_dValue = r.isText ? std::numeric_limits<double>::quiet_NaN() : r.value;
    if (r.isText || r.isError)
    {
        pCell->m_wsText = r.text;
        return;
    }
    if (isBoolean)
    {
        // TRUE or FALSE, whatever the number format (as in Excel); m_dValue
        // is 1 or 0. A comparison showed 1 or 0.
        pCell->m_wsText = r.value != 0.0 ? L"TRUE" : L"FALSE";
        return;
    }
    std::wstring general = CFormulator::FormatGeneral(r.value);
    if (isDate && pCell->m_nFormat == FMT_GENERAL)
    {
        // A date shows as a date in General, as a typed one does
        // (=TODAY() showed 46340); a Number format shows the serial.
        pCell->m_wsText = Grid32Detail::FormatGeneralDate(r.value, general);
        return;
    }
    const bool renderable = std::isfinite(r.value) && std::fabs(r.value) < 1e14;
    pCell->m_wsText = renderable
        ? Grid32Detail::FormatCellValue(CT_Formula, r.value, general, pCell->m_nFormat)
        : general;
}

// How an expression typed at the prompt displays: as a General formula cell.
std::string DisplayOf(const CFormulator::Result& r)
{
    GRIDCELL cell;
    RenderFormulaResult(&cell, r);
    return Narrow(cell.m_wsText);
}

} // namespace

// ---- the sheet --------------------------------------------------------------
// mapCells and gcs are protected in CGrid32Mgr, so the sheet derives from it.
class ConsoleSheet : public CGrid32Mgr
{
public:
    ConsoleSheet() { gcs.nWidth = kSheetCols; gcs.nHeight = kSheetRows; }
    ~ConsoleSheet() { ClearAll(); }

    std::map<CellKey, PGRIDCELL>& Cells() { return mapCells; }

    PGRIDCELL Find(const CellKey& k)
    {
        auto it = mapCells.find(k);
        return it != mapCells.end() ? it->second : nullptr;
    }

    PGRIDCELL GetOrCreate(const CellKey& k)
    {
        PGRIDCELL& slot = mapCells[k];
        if (!slot) slot = new GRIDCELL();
        return slot;
    }

    // As typed into the cell: "=..." is a formula, "'..." is text (the
    // apostrophe dropped), anything else is typed by InferType and rendered
    // under the cell's number format (SetCellText).
    void Type(const CellKey& k, const std::wstring& text)
    {
        PGRIDCELL cell = GetOrCreate(k);
        if (!text.empty() && text[0] == L'=')
        {
            cell->m_bFormula = true;
            cell->m_wsFormula = text.substr(1);
            cell->m_eType = CT_Formula;
        }
        else if (!text.empty() && text[0] == Grid32Detail::kTextPrefix)
        {
            // '42 is the text "42", as in the grid: the apostrophe is dropped.
            Literal(k, text.substr(1));
            return;
        }
        else
        {
            cell->m_bFormula = false;
            cell->m_wsFormula.clear();
            double v = 0.0;
            cell->m_eType = Grid32Detail::InferType(text, v);
            cell->m_dValue = v;
            cell->m_wsText = Grid32Detail::FormatCellValue(cell->m_eType, v, text, cell->m_nFormat);
        }
        Recalculate();
    }

    // A CT_Text cell holding `text` whatever it looks like — what GM_SETCELL
    // and fills store, since neither infers a type.
    void Literal(const CellKey& k, const std::wstring& text)
    {
        PGRIDCELL cell = GetOrCreate(k);
        cell->m_bFormula = false;
        cell->m_wsFormula.clear();
        cell->m_eType = CT_Text;
        cell->m_dValue = 0.0;
        cell->m_wsText = text;
        Recalculate();
    }

    // Apply a number format, re-rendering as SetCellNumberFormat does.
    void Format(const CellKey& k, UINT fmt)
    {
        PGRIDCELL cell = GetOrCreate(k);
        cell->m_nFormat = fmt;
        if (!cell->m_bFormula && cell->m_eType != CT_Text && cell->m_eType != CT_Error)
        {
            cell->m_wsText = fmt == FMT_GENERAL
                ? Grid32Detail::FormatEditValue(cell->m_eType, cell->m_dValue, cell->m_wsText)
                : Grid32Detail::FormatCellValue(cell->m_eType, cell->m_dValue, cell->m_wsText, fmt);
        }
    }

    bool Clear(const CellKey& k)
    {
        auto it = mapCells.find(k);
        if (it == mapCells.end()) return false;
        delete it->second;
        mapCells.erase(it);
        return true;
    }

    void ClearAll()
    {
        for (auto& kv : mapCells) delete kv.second;
        mapCells.clear();
        results.clear();
    }

    // One pass over every formula cell, as RecalculateFormulas does. Returns
    // the number of formula evaluations it took.
    size_t Recalculate()
    {
        const size_t before = CFormulator::CellEvaluationCount();
        results.clear();
        CFormulator::RecalcPass pass(this);
        for (auto& kv : mapCells)
        {
            PGRIDCELL cell = kv.second;
            if (cell && cell->m_bFormula)
            {
                CFormulator::Result r = CFormulator::EvaluateCell(this, kv.first.first, kv.first.second);
                RenderFormulaResult(cell, r);
                results[kv.first] = r;
            }
        }
        return CFormulator::CellEvaluationCount() - before;
    }

    bool InGrid(const CellKey& k) const
    {
        return k.first < kSheetRows && k.second < kSheetCols;
    }

    // The last recalculation's result for a formula cell.
    const CFormulator::Result* ResultOf(const CellKey& k) const
    {
        auto it = results.find(k);
        return it != results.end() ? &it->second : nullptr;
    }

private:
    std::map<CellKey, CFormulator::Result> results;
};

namespace {

// ---- the console ------------------------------------------------------------

const char* const kHelp =
"Evaluate\n"
"  <expr>  or  =<expr>        evaluate against the sheet:  =SUM(A1:A3)*2\n"
"\n"
"Cells\n"
"  A1: <content>              type into A1, as in the grid:  A1: 5   A2: =A1*2\n"
"  set A1 <content>           the same                       A3: hello   A4: '42 (text)\n"
"  text A1 <content>          store as literal text, as GM_SETCELL does (no inference,\n"
"                             a leading ' kept)\n"
"  A1                         inspect a cell: input, type, value, display, format\n"
"  clear A1 | A1:B3 | all     empty cells\n"
"  format A1[:B3] <name>      general number integer currency percent\n"
"                             date usdate time thousands\n"
"\n"
"Look\n"
"  show [A1:D10]              the sheet as the grid would display it\n"
"  list                       every cell: input, type, value, display\n"
"  dump                       the sheet as a script that recreates it\n"
"  recalc                     recalculate; reports how many formulas it evaluated\n"
"\n"
"Scripts\n"
"  load <file>                run a script (the same commands, one per line)\n"
"  expect <A1|=expr> => <text>   check what a cell or expression displays;\n"
"                             quote the text for spaces or empty: => \"\"\n"
"  echo <text>                print text\n"
"  # comment                  ignored\n"
"\n"
"  help, quit\n";

struct Console
{
    ConsoleSheet sheet;
    bool scriptMode = false;     // echo commands, no prompt
    int passed = 0, failed = 0;
    int loadDepth = 0;

    void Out(const std::string& s) { std::cout << s << "\n"; }

    // "A1" -> key; in the grid only.
    bool Ref(const std::string& text, CellKey& k)
    {
        UINT r = 0, c = 0;
        if (!CFormulator::ParseCellRef(Widen(text), r, c)) return false;
        k = CellKey(r, c);
        return sheet.InGrid(k);
    }

    bool Range(const std::string& text, UINT& sr, UINT& sc, UINT& er, UINT& ec)
    {
        if (!CFormulator::ParseRange(Widen(text), sr, sc, er, ec)) return false;
        return sheet.InGrid(CellKey(sr, sc)) && sheet.InGrid(CellKey(er, ec));
    }

    std::string InputOf(PGRIDCELL cell)
    {
        if (!cell) return std::string();
        if (cell->m_bFormula) return "=" + Narrow(cell->m_wsFormula);
        // What the grid's editor shows: text that needs its apostrophe has it.
        if (cell->m_eType == CT_Text && Grid32Detail::NeedsTextPrefix(cell->m_wsText))
            return "'" + Narrow(cell->m_wsText);
        return Narrow(cell->m_wsText);
    }

    std::string ValueOf(const CellKey& k, PGRIDCELL cell)
    {
        if (!cell) return "(blank)";
        if (cell->m_bFormula)
        {
            const CFormulator::Result* r = sheet.ResultOf(k);
            if (!r) return "?";
            if (r->isError) return std::string("error ") + Narrow(r->text);
            if (r->isText) return "text \"" + Narrow(r->text) + "\"";
            if (r->isDate) return "date, serial " + Narrow(CFormulator::FormatGeneral(r->value));
            if (r->isBoolean) return std::string("boolean ") + (r->value != 0.0 ? "TRUE" : "FALSE");
            return "number " + Narrow(CFormulator::FormatGeneral(r->value));
        }
        if (cell->m_wsText.empty()) return "(blank)";
        const std::string v = Narrow(CFormulator::FormatGeneral(cell->m_dValue));
        switch (cell->m_eType)
        {
        case CT_Number:  return "number " + v;
        case CT_Date:    return "date, serial " + v;
        case CT_Boolean: return "boolean " + v;
        default:         return "text \"" + Narrow(cell->m_wsText) + "\"";
        }
    }

    void Inspect(const CellKey& k)
    {
        PGRIDCELL cell = sheet.Find(k);
        if (!cell) { Out(RefName(k) + "  (empty)"); return; }
        Out(RefName(k));
        Out("  input    " + InputOf(cell));
        Out("  type     " + std::string(cell->m_bFormula ? "formula" : TypeName(cell->m_eType)));
        Out("  value    " + ValueOf(k, cell));
        Out("  display  " + Narrow(cell->m_wsText));
        Out("  format   " + std::string(FormatToName(cell->m_nFormat)));
    }

    void Evaluate(const std::string& expr)
    {
        std::wstring w = Widen(expr);
        if (!w.empty() && w[0] == L'=') w = w.substr(1);
        const CFormulator::Result r = CFormulator::Evaluate(&sheet, w);
        std::string line = "  " + DisplayOf(r);
        const size_t pad = 24;
        if (Width(line) < pad) line.append(pad - Width(line), ' ');
        line += "  (" + std::string(ResultKind(r));
        if (!r.isError && r.isText)
            line += ", reads as " + Narrow(CFormulator::FormatGeneral(r.value));
        line += ")";
        Out(line);
    }

    // The used range, or `range` if given, as the grid shows it.
    void Show(const std::string& range)
    {
        UINT sr = 0, sc = 0, er = 0, ec = 0;
        if (!range.empty())
        {
            if (!Range(range, sr, sc, er, ec)) { Out("  not a range in the grid: " + range); return; }
        }
        else
        {
            if (sheet.Cells().empty()) { Out("  (the sheet is empty)"); return; }
            for (auto& kv : sheet.Cells())
            {
                if (kv.first.first > er) er = kv.first.first;
                if (kv.first.second > ec) ec = kv.first.second;
            }
        }
        const UINT maxRows = 50, maxCols = 12;
        bool clipped = false;
        if (er - sr + 1 > maxRows) { er = sr + maxRows - 1; clipped = true; }
        if (ec - sc + 1 > maxCols) { ec = sc + maxCols - 1; clipped = true; }

        const size_t maxWidth = 20;
        std::vector<size_t> width(ec - sc + 1, 0);
        std::vector<std::vector<std::string>> text(er - sr + 1, std::vector<std::string>(ec - sc + 1));
        std::vector<std::vector<bool>> right(er - sr + 1, std::vector<bool>(ec - sc + 1, false));
        for (UINT c = sc; c <= ec; ++c) width[c - sc] = ColumnName(c).size();
        for (UINT r = sr; r <= er; ++r)
        {
            for (UINT c = sc; c <= ec; ++c)
            {
                PGRIDCELL cell = sheet.Find(CellKey(r, c));
                if (!cell) continue;
                std::string t = Narrow(cell->m_wsText);
                if (Width(t) > maxWidth)
                {
                    // Cut at a character boundary and mark the cut.
                    size_t n = 0, i = 0;
                    for (; i < t.size(); ++i)
                        if (((unsigned char)t[i] & 0xC0) != 0x80 && ++n > maxWidth - 1) break;
                    t = t.substr(0, i) + "~";
                }
                // Numbers right, text and errors left, as the grid aligns them.
                bool isRight = false;
                if (cell->m_bFormula)
                {
                    const CFormulator::Result* res = sheet.ResultOf(CellKey(r, c));
                    isRight = res && !res->isText && !res->isError && !res->isBoolean;
                }
                else
                {
                    isRight = cell->m_eType == CT_Number || cell->m_eType == CT_Date || cell->m_eType == CT_Boolean;
                }
                text[r - sr][c - sc] = t;
                right[r - sr][c - sc] = isRight;
                if (Width(t) > width[c - sc]) width[c - sc] = Width(t);
            }
        }

        const size_t rowHeader = std::to_string(er + 1).size();
        std::string line(rowHeader + 1, ' ');
        for (UINT c = sc; c <= ec; ++c)
        {
            std::string name = ColumnName(c);
            line += "| " + name + std::string(width[c - sc] - name.size() + 1, ' ');
        }
        Out(line);
        std::string rule(rowHeader + 1, '-');
        for (size_t w : width) rule += "+" + std::string(w + 2, '-');
        Out(rule);
        for (UINT r = sr; r <= er; ++r)
        {
            std::string num = std::to_string(r + 1);
            line = std::string(rowHeader - num.size(), ' ') + num + " ";
            for (UINT c = sc; c <= ec; ++c)
            {
                const std::string& t = text[r - sr][c - sc];
                const std::string pad(width[c - sc] - Width(t), ' ');
                line += "| " + (right[r - sr][c - sc] ? pad + t : t + pad) + " ";
            }
            Out(line);
        }
        if (clipped)
            Out("  (clipped to " + std::to_string(maxRows) + " rows x " + std::to_string(maxCols) +
                " columns; give a range to see more)");
    }

    void List()
    {
        if (sheet.Cells().empty()) { Out("  (the sheet is empty)"); return; }
        for (auto& kv : sheet.Cells())
        {
            PGRIDCELL cell = kv.second;
            std::string line = RefName(kv.first);
            line.append(line.size() < 7 ? 7 - line.size() : 1, ' ');
            line += InputOf(cell);
            line.append(Width(line) < 32 ? 32 - Width(line) : 1, ' ');
            line += "-> " + Narrow(cell->m_wsText);
            line.append(Width(line) < 52 ? 52 - Width(line) : 1, ' ');
            line += "[" + std::string(cell->m_bFormula ? "formula, " : "") + ValueOf(kv.first, cell);
            if (cell->m_nFormat != FMT_GENERAL) line += ", " + std::string(FormatToName(cell->m_nFormat));
            line += "]";
            Out(line);
        }
    }

    // The sheet as commands that rebuild it: paste into a bug report.
    void Dump()
    {
        Out("clear all");
        for (auto& kv : sheet.Cells())
        {
            PGRIDCELL cell = kv.second;
            const std::string ref = RefName(kv.first);
            if (cell->m_bFormula)
                Out(ref + ": " + InputOf(cell));
            else if (cell->m_eType == CT_Text && !cell->m_wsText.empty())
            {
                // Text that typing would read as something else keeps its
                // apostrophe, as the grid's editor shows it.
                Out(ref + ": " + (Grid32Detail::NeedsTextPrefix(cell->m_wsText) ? "'" : "") +
                    Narrow(cell->m_wsText));
            }
            else if (!cell->m_wsText.empty())
                Out(ref + ": " + Narrow(Grid32Detail::FormatEditValue(cell->m_eType, cell->m_dValue, cell->m_wsText)));
            if (cell->m_nFormat != FMT_GENERAL)
                Out("format " + ref + " " + FormatToName(cell->m_nFormat));
        }
    }

    // Unquote an `expect` value: "..." with "" for a quote; otherwise as is.
    static std::string Unquote(const std::string& s)
    {
        if (s.size() >= 2 && s.front() == '"' && s.back() == '"')
        {
            std::string out;
            for (size_t i = 1; i + 1 < s.size(); ++i)
            {
                out.push_back(s[i]);
                if (s[i] == '"' && i + 2 < s.size() && s[i + 1] == '"') ++i;
            }
            return out;
        }
        return s;
    }

    void Expect(const std::string& args)
    {
        const size_t arrow = args.find("=>");
        if (arrow == std::string::npos) { Out("  usage: expect <A1|=expr> => <text>"); return; }
        const std::string target = Trim(args.substr(0, arrow));
        const std::string want = Unquote(Trim(args.substr(arrow + 2)));

        std::string got, kind;
        CellKey k;
        if (Ref(target, k))
        {
            PGRIDCELL cell = sheet.Find(k);
            got = cell ? Narrow(cell->m_wsText) : std::string();
            kind = cell ? (cell->m_bFormula ? "formula" : TypeName(cell->m_eType)) : "empty";
        }
        else
        {
            std::wstring w = Widen(target);
            if (!w.empty() && w[0] == L'=') w = w.substr(1);
            const CFormulator::Result r = CFormulator::Evaluate(&sheet, w);
            got = DisplayOf(r);
            kind = ResultKind(r);
        }

        if (got == want) { ++passed; if (!scriptMode) Out("  ok"); return; }
        ++failed;
        Out("  FAIL  expect " + target + " => \"" + want + "\"   got \"" + got + "\" (" + kind + ")");
    }

    void Load(const std::string& path)
    {
        if (loadDepth > 8) { Out("  load: nested too deeply"); return; }
        std::ifstream in(path.c_str(), std::ios::binary);
        if (!in) { Out("  load: cannot open " + path); return; }
        ++loadDepth;
        const bool outer = scriptMode;
        scriptMode = true;
        std::string line;
        bool first = true;
        while (std::getline(in, line))
        {
            if (first && line.size() >= 3 && line.compare(0, 3, "\xEF\xBB\xBF") == 0)
                line = line.substr(3);     // a UTF-8 BOM, as Notepad writes
            first = false;
            if (!Run(line)) break;
        }
        scriptMode = outer;
        --loadDepth;
    }

    // Run one line. Returns false to quit.
    bool Run(const std::string& raw)
    {
        const std::string line = Trim(raw);
        if (line.empty()) return true;
        if (line[0] == '#')
        {
            if (scriptMode) Out(line);
            return true;
        }
        if (scriptMode) Out("> " + line);

        std::string rest;
        const std::string word = FirstWord(line, rest);
        const std::string cmd = Lower(word);

        if (cmd == "quit" || cmd == "exit") return false;
        if (cmd == "help" || cmd == "?") { std::cout << kHelp; return true; }
        if (cmd == "echo") { Out(rest); return true; }
        if (cmd == "show") { Show(rest); return true; }
        if (cmd == "list") { List(); return true; }
        if (cmd == "dump") { Dump(); return true; }
        if (cmd == "load") { Load(rest); return true; }
        if (cmd == "expect") { Expect(rest); return true; }
        if (cmd == "recalc")
        {
            Out("  " + std::to_string(sheet.Recalculate()) + " formula evaluation(s)");
            return true;
        }
        if (cmd == "set" || cmd == "text")
        {
            std::string content;
            const std::string ref = FirstWord(rest, content);
            CellKey k;
            if (!Ref(ref, k)) { Out("  not a cell in the grid: " + ref); return true; }
            // Content is taken as typed; only the single separating space goes.
            const size_t at = line.find(ref, word.size()) + ref.size();
            std::string typed = at < line.size() ? line.substr(at) : std::string();
            if (!typed.empty() && (typed[0] == ' ' || typed[0] == '\t')) typed = typed.substr(1);
            if (cmd == "set") sheet.Type(k, Widen(typed)); else sheet.Literal(k, Widen(typed));
            return true;
        }
        if (cmd == "clear")
        {
            if (rest.empty() || Lower(rest) == "all") { sheet.ClearAll(); return true; }
            UINT sr = 0, sc = 0, er = 0, ec = 0;
            if (!Range(rest, sr, sc, er, ec)) { Out("  not a cell or range in the grid: " + rest); return true; }
            for (UINT r = sr; r <= er; ++r)
                for (UINT c = sc; c <= ec; ++c)
                    sheet.Clear(CellKey(r, c));
            sheet.Recalculate();
            return true;
        }
        if (cmd == "format")
        {
            std::string name;
            const std::string target = FirstWord(rest, name);
            UINT fmt = 0, sr = 0, sc = 0, er = 0, ec = 0;
            if (!Range(target, sr, sc, er, ec)) { Out("  not a cell or range in the grid: " + target); return true; }
            if (!NameToFormat(name, fmt)) { Out("  unknown format '" + name + "'; see help"); return true; }
            for (UINT r = sr; r <= er; ++r)
                for (UINT c = sc; c <= ec; ++c)
                    sheet.Format(CellKey(r, c), fmt);
            sheet.Recalculate();    // formula cells re-render in their format
            return true;
        }

        // "A1: content" types into A1; "A1" alone inspects it. The colon must
        // be followed by a space, '=' or the end, so "A1:B2" is not taken for
        // one.
        const size_t colon = line.find(':');
        if (colon != std::string::npos)
        {
            CellKey k;
            const bool sep = colon + 1 == line.size() || line[colon + 1] == ' ' ||
                line[colon + 1] == '\t' || line[colon + 1] == '=';
            const std::string target = Trim(line.substr(0, colon));
            UINT r = 0, c = 0;
            if (sep && !Ref(target, k) && CFormulator::ParseCellRef(Widen(target), r, c))
            {
                Out("  " + target + " is outside the grid (A1:" + ColumnName(kSheetCols - 1) +
                    std::to_string(kSheetRows) + ")");
                return true;
            }
            if (sep && Ref(target, k))
            {
                std::string typed = line.substr(colon + 1);
                if (!typed.empty() && (typed[0] == ' ' || typed[0] == '\t')) typed = typed.substr(1);
                sheet.Type(k, Widen(typed));
                return true;
            }
        }
        CellKey k;
        if (rest.empty() && Ref(word, k)) { Inspect(k); return true; }

        Evaluate(line);
        return true;
    }
};

// One line from the console window. On Windows, UTF-8 read through std::cin
// from a console is unreliable for anything outside ASCII, so an interactive
// session reads UTF-16 with ReadConsoleW and converts it. Piped input and
// scripts are read as UTF-8 bytes on every platform.
bool ReadLine(bool tty, std::string& out)
{
#ifdef _WIN32
    if (tty)
    {
        HANDLE h = GetStdHandle(STD_INPUT_HANDLE);
        std::wstring w;
        wchar_t buf[512];
        for (;;)
        {
            DWORD n = 0;
            if (!ReadConsoleW(h, buf, 512, &n, nullptr) || n == 0)
                break;
            w.append(buf, n);
            if (w.back() == L'\n')
                break;
        }
        if (w.empty() || w[0] == 0x1A)      // Ctrl+Z, Enter: end of input
            return false;
        while (!w.empty() && (w.back() == L'\n' || w.back() == L'\r'))
            w.pop_back();
        out = Narrow(w);
        return true;
    }
#else
    (void)tty;
#endif
    return static_cast<bool>(std::getline(std::cin, out));
}

} // namespace

int main(int argc, char** argv)
{
#ifdef _WIN32
    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);
#endif
    Console console;
    bool interactive = argc == 1;
    std::vector<std::string> scripts;
    for (int i = 1; i < argc; ++i)
    {
        const std::string a = argv[i];
        if (a == "-i") interactive = true;
        else if (a == "-h" || a == "--help") { std::cout << "usage: FormulaConsole [-i] [script ...]\n\n" << kHelp; return 0; }
        else scripts.push_back(a);
    }

    for (const std::string& s : scripts) console.Load(s);

    const bool tty = CONSOLE_ISATTY(CONSOLE_FILENO(stdin)) != 0;
    if (interactive)
    {
        console.scriptMode = !tty;      // piped input reads like a script
        if (tty)
            std::cout << "Formula console — " << kSheetCols << " columns x " << kSheetRows
                      << " rows. Type help for commands.\n";
        std::string line;
        for (;;)
        {
            if (tty) std::cout << "fx> " << std::flush;
            if (!ReadLine(tty, line)) break;
            if (!console.Run(line)) break;
        }
        if (tty) std::cout << "\n";
    }

    if (console.passed + console.failed > 0)
        std::cout << "\nexpect: " << console.passed << " passed, " << console.failed << " failed\n";
    return console.failed ? 1 : 0;
}

// ---- link seam -------------------------------------------------------------
// The only CGrid32Mgr members Formulator.cpp needs at link time; see the note
// at the top of FormulatorTests.cpp. Grid32Mgr.cpp is not part of this exe.
CGrid32Mgr::CGrid32Mgr()
    : m_hWndGrid(nullptr), m_hWndEdit(nullptr),
      nColHeaderHeight(0), nRowHeaderWidth(0),
      m_bRedraw(TRUE), m_bResizable(FALSE), m_bUndoRecordEnabled(FALSE),
      m_bDeferRecalc(FALSE), m_bSuppressChangeNotify(FALSE), dwError(0),
      m_bSelecting(FALSE), m_bSizing(FALSE), m_editWndProc(nullptr),
      m_nMouseHoverDelay(0), m_npHoverDelaySet(0),
      m_gridHitTest(0), m_nToBeSized(0), m_resizeStartSize(0),
      m_bEditing(false), m_editCell{ 0, 0 }, m_hEditFont(nullptr),
      m_nSizingLine(0), m_rgbSizingLine(0),
      m_hDefaultFont(nullptr), m_pBatch(nullptr), m_bMergesDirty(true), m_cfGridCells(0),
      m_gdiplusToken(0)
{
    memset(&gcs, 0, sizeof(gcs));
    memset(&m_selectionRect, 0, sizeof(m_selectionRect));
    m_currentCell = m_visibleTopLeft = m_HoverCell = { 0, 0 };
    m_mouseDraggingStartPoint = { 0, 0 };
    m_clientRect = m_paintRect = { 0, 0, 0, 0 };
}

CGrid32Mgr::~CGrid32Mgr() {}

PGRIDCELL CGrid32Mgr::GetCell(UINT nRow, UINT nCol)
{
    auto it = mapCells.find(std::make_pair(nRow, nCol));
    return (it != mapCells.end()) ? it->second : nullptr;
}
