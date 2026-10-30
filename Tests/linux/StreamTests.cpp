// File import/export tests -- GM_STREAMOUT, GM_STREAMIN and GM_CLEARALL,
// driven through the real CGrid32Mgr against the Win32 shim (see README.md).
//
// Covers AUDIT_2026-09.md B7 (CSV/TSV quoting), H15 (stream-in replaces, the
// grid can be cleared), H19 (well-formed ODF), H20 (SF_TEXT, SF_RTF, a real
// SF_XML, an explicit refusal for SF_XLSX), M14 (exact, locale-independent
// SSF numbers) and P1 (stream-out walks the cells that exist, and reports the
// exact buffer size it needs). Round trips are the backbone: write a sheet in
// a format, read it into a fresh grid, and compare.

#define _CRT_SECURE_NO_WARNINGS
#include <pch.h>
#include <grid32_internal.h>
#include "win32shim/win32stubs.h"

#include <algorithm>
#include <chrono>
#include <clocale>
#include <cmath>
#include <cstring>
#include <functional>
#include <iostream>
#include <locale>
#include <map>
#include <string>
#include <vector>

// ---- tiny test harness (same shape as GridMgrTests.cpp) --------------------
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
#define ASSERT_EQ(a, b) do { auto _a = (a); auto _b = (b); if (!(_a == _b)) { \
    std::cerr << "  (got " << _a << ", want " << _b << ")\n"; \
    FAIL_HERE("expected " #a " == " #b); } } while (0)
// Failure output is narrow only: once std::wcerr writes to stderr, glibc
// makes the stream wide-oriented and silently drops later std::cerr output.
#define ASSERT_WSTR(expr, want) do { std::wstring _g = (expr); std::wstring _w = (want); \
    if (_g != _w) { std::cerr << "  (got \"" << Visible(_g) << "\", want \"" << Visible(_w) << "\")\n"; \
    FAIL_HERE("expected " #expr " == " #want); } } while (0)
#define ASSERT_CONTAINS(hay, needle) do { std::wstring _h = (hay); std::wstring _n = (needle); \
    if (_h.find(_n) == std::wstring::npos) { std::cerr << "  (\"" << Visible(_n) << "\" not in \"" \
    << Visible(_h.substr(0, 600)) << "\")\n"; FAIL_HERE("expected " #hay " to contain " #needle); } } while (0)
#define ASSERT_NOT_CONTAINS(hay, needle) do { std::wstring _h = (hay); std::wstring _n = (needle); \
    if (_h.find(_n) != std::wstring::npos) { std::cerr << "  (\"" << Visible(_n) << "\" found)\n"; \
    FAIL_HERE("expected " #hay " not to contain " #needle); } } while (0)

static std::string Narrow(const std::wstring& s)
{
    std::string out;
    for (wchar_t ch : s) out.push_back(ch < 0x80 ? (char)ch : '?');
    return out;
}

// A wide string as printable ASCII: control characters and non-ASCII escaped.
static std::string Visible(const std::wstring& s)
{
    std::string out;
    for (wchar_t ch : s)
    {
        if (ch == L'\r') out += "\\r";
        else if (ch == L'\n') out += "\\n";
        else if (ch == L'\t') out += "\\t";
        else if (ch < 0x20 || ch > 0x7E) out += "\\u" + std::to_string((unsigned long)ch);
        else out.push_back((char)ch);
    }
    return out;
}

// ---- fixture ---------------------------------------------------------------
// As in GridMgrTests.cpp: a grid with no real window behind it, deriving from
// CGrid32Mgr so it can reach the state a stream operation is meant to reset.
class Sheet : private CGrid32Mgr
{
public:
    Sheet(size_t cols = 16, size_t rows = 64)
    {
        static int fakeHwnd = 0;
        m_hWndGrid = &fakeHwnd;
        GRIDCREATESTRUCT gcs{};
        gcs.cbSize = sizeof(GRIDCREATESTRUCT);
        gcs.nWidth = cols;
        gcs.nHeight = rows;
        gcs.nDefColWidth = 100;
        gcs.nDefRowHeight = 20;
        gcs.style = GS_SPREADSHEET;
        m_ok = Create(&gcs);
    }
    bool ok() const { return m_ok; }
    CGrid32Mgr& mgr() { return *this; }

    void Set(UINT row, UINT col, const std::wstring& text) { SetCellText(row, col, text.c_str()); }
    std::wstring Text(UINT row, UINT col)
    {
        PGRIDCELL c = GetCell(row, col);
        return c ? c->m_wsText : std::wstring();
    }
    PGRIDCELL Cell(UINT row, UINT col) { return GetCell(row, col); }
    size_t CellCount() const { return mapCells.size(); }
    GRIDPOINT Current() const { return m_currentCell; }
    GRIDSELECTION Selection() const { return m_selectionRect; }
private:
    bool m_ok;
};

// ---- stream helpers --------------------------------------------------------
static int g_callbacks = 0;
static DWORD CALLBACK CountCallback(tagGCSTREAM*) { ++g_callbacks; return 0; }

// Stream out into a buffer of `bufWchars` (big enough by default).
static std::wstring Out(Sheet& s, DWORD fmt, DWORD* error = nullptr, size_t bufWchars = 4u << 20)
{
    std::vector<wchar_t> buf(bufWchars, L'#');
    GCSTREAM st{};
    st.m_cbSize = sizeof(GCSTREAM);
    st.m_pwszBuff = buf.data();
    st.m_cbBuffSize = (UINT)(buf.size() * sizeof(wchar_t));
    st.m_dwFormat = fmt;
    st.m_dwError = 0xDEAD;
    s.mgr().OnStreamOut(&st);
    if (error) *error = st.m_dwError;
    if (st.m_dwError != 0)
        return std::wstring();
    return std::wstring(buf.data(), st.m_cbBuffOut / sizeof(wchar_t));
}

static DWORD In(Sheet& s, const std::wstring& text, DWORD fmt)
{
    std::vector<wchar_t> buf(text.begin(), text.end());
    buf.push_back(L'\0');
    GCSTREAM st{};
    st.m_cbSize = sizeof(GCSTREAM);
    st.m_pwszBuff = buf.data();
    st.m_cbBuffSize = (UINT)(buf.size() * sizeof(wchar_t));
    st.m_dwFormat = fmt;
    st.m_dwError = 0xDEAD;
    s.mgr().OnStreamIn(&st);
    return st.m_dwError;
}

// Write `src` in `fmt` and load the result into `dst`.
static bool RoundTrip(Sheet& src, DWORD fmt, Sheet& dst)
{
    DWORD err = 1;
    std::wstring data = Out(src, fmt, &err);
    if (err != 0) { std::cerr << "  (stream-out error " << err << ")\n"; return false; }
    err = In(dst, data, fmt);
    if (err != 0) { std::cerr << "  (stream-in error " << err << ")\n"; return false; }
    return true;
}

// ---- a small XML checker -----------------------------------------------------
// XML 1.0 well-formedness for what the writers produce -- balanced tags, one
// root, quoted and unique attributes, only known entity references -- plus
// namespace well-formedness: every element and attribute prefix must be
// declared in scope. The last is exactly what H19 broke. On success it returns
// the root element's namespace URI and local name.
struct XmlRoot { std::wstring ns, local; };

static bool XmlCheck(const std::wstring& x, XmlRoot& root, std::string& why)
{
    const size_t npos = std::wstring::npos;
    std::vector<std::wstring> open;
    std::vector<std::map<std::wstring, std::wstring>> scopes;
    bool sawRoot = false;
    size_t i = 0;
    auto fail = [&](const std::string& msg) { why = msg + " at offset " + std::to_string(i); return false; };
    auto lookup = [&](const std::wstring& prefix, std::wstring& uri) {
        if (prefix == L"xml") { uri = L"http://www.w3.org/XML/1998/namespace"; return true; }
        for (auto it = scopes.rbegin(); it != scopes.rend(); ++it)
        {
            auto f = it->find(prefix);
            if (f != it->end()) { uri = f->second; return true; }
        }
        uri.clear();
        return prefix.empty();
    };
    auto prefixOf = [](const std::wstring& q) {
        size_t c = q.find(L':');
        return c == std::wstring::npos ? std::wstring() : q.substr(0, c);
    };
    auto goodText = [&](size_t from, size_t to) {
        for (size_t k = from; k < to; ++k)
        {
            if (x[k] == L'<') return false;
            if (x[k] < 0x20 && x[k] != L'\t' && x[k] != L'\r' && x[k] != L'\n') return false;
            if (x[k] != L'&') continue;
            size_t semi = x.find(L';', k);
            if (semi == npos || semi > to) return false;
            std::wstring e = x.substr(k + 1, semi - k - 1);
            if (!(e == L"amp" || e == L"lt" || e == L"gt" || e == L"quot" || e == L"apos" ||
                  (e.size() > 1 && e[0] == L'#')))
                return false;
            k = semi;
        }
        return true;
    };

    while (i < x.size())
    {
        if (x[i] != L'<')
        {
            size_t next = x.find(L'<', i);
            if (next == npos) next = x.size();
            if (open.empty())
            {
                for (size_t k = i; k < next; ++k)
                    if (!iswspace(x[k])) return fail("text outside the root element");
            }
            else if (!goodText(i, next))
                return fail("bad character data");
            i = next;
            continue;
        }
        if (x.compare(i, 4, L"<!--") == 0)
        {
            size_t e = x.find(L"-->", i);
            if (e == npos) return fail("unterminated comment");
            i = e + 3;
            continue;
        }
        if (x.compare(i, 2, L"<?") == 0)
        {
            if (i != 0 && x.compare(i, 6, L"<?xml ") == 0) return fail("XML declaration not at the start");
            size_t e = x.find(L"?>", i);
            if (e == npos) return fail("unterminated processing instruction");
            i = e + 2;
            continue;
        }
        if (x.compare(i, 2, L"</") == 0)
        {
            size_t e = x.find(L'>', i);
            if (e == npos) return fail("unterminated end tag");
            std::wstring name = x.substr(i + 2, e - i - 2);
            while (!name.empty() && iswspace(name.back())) name.pop_back();
            if (open.empty() || open.back() != name) return fail("mismatched end tag </" + Narrow(name) + ">");
            open.pop_back();
            scopes.pop_back();
            i = e + 1;
            continue;
        }
        size_t k = i + 1;
        while (k < x.size() && !iswspace(x[k]) && x[k] != L'>' && x[k] != L'/') ++k;
        std::wstring qname = x.substr(i + 1, k - i - 1);
        if (qname.empty()) return fail("empty element name");
        if (open.empty() && sawRoot) return fail("a second root element");
        std::map<std::wstring, std::wstring> decls;
        std::vector<std::wstring> attrs;
        bool selfClosing = false;
        for (;;)
        {
            while (k < x.size() && iswspace(x[k])) ++k;
            if (k >= x.size()) return fail("unterminated start tag");
            if (x[k] == L'>') { ++k; break; }
            if (x[k] == L'/' && k + 1 < x.size() && x[k + 1] == L'>') { selfClosing = true; k += 2; break; }
            size_t an = k;
            while (k < x.size() && x[k] != L'=' && !iswspace(x[k]) && x[k] != L'>') ++k;
            std::wstring aname = x.substr(an, k - an);
            while (k < x.size() && iswspace(x[k])) ++k;
            if (k >= x.size() || x[k] != L'=') return fail("attribute without a value");
            ++k;
            while (k < x.size() && iswspace(x[k])) ++k;
            if (k >= x.size() || (x[k] != L'"' && x[k] != L'\'')) return fail("unquoted attribute");
            wchar_t q = x[k++];
            size_t ve = x.find(q, k);
            if (ve == npos) return fail("unterminated attribute value");
            if (!goodText(k, ve)) return fail("bad attribute value");
            if (std::find(attrs.begin(), attrs.end(), aname) != attrs.end()) return fail("duplicate attribute");
            attrs.push_back(aname);
            if (aname == L"xmlns") decls[L""] = x.substr(k, ve - k);
            else if (aname.compare(0, 6, L"xmlns:") == 0) decls[aname.substr(6)] = x.substr(k, ve - k);
            k = ve + 1;
        }
        scopes.push_back(decls);
        std::wstring uri;
        if (!lookup(prefixOf(qname), uri)) return fail("undeclared prefix on <" + Narrow(qname) + ">");
        for (const std::wstring& a : attrs)
        {
            if (a == L"xmlns" || a.compare(0, 6, L"xmlns:") == 0) continue;
            std::wstring p = prefixOf(a), u;
            if (!p.empty() && !lookup(p, u)) return fail("undeclared prefix on attribute " + Narrow(a));
        }
        if (open.empty())
        {
            sawRoot = true;
            root.ns = uri;
            size_t c = qname.find(L':');
            root.local = c == npos ? qname : qname.substr(c + 1);
        }
        if (selfClosing) scopes.pop_back();
        else open.push_back(qname);
        i = k;
    }
    if (!open.empty()) return fail("unclosed <" + Narrow(open.back()) + ">");
    if (!sawRoot) return fail("no root element");
    return true;
}

#define ASSERT_XML(doc, root) do { std::string _why; if (!XmlCheck((doc), (root), _why)) { \
    std::cerr << "  (" << _why << ")\n"; FAIL_HERE("expected well-formed XML: " #doc); } } while (0)

// =============================================================================
// B7 -- CSV / TSV quoting
// =============================================================================
TEST(Csv_QuotesFieldsThatNeedIt)
{
    Sheet s;
    s.Set(0, 0, L"a,b");
    s.Set(0, 1, L"say \"hi\"");
    s.Set(0, 2, L"two\nlines");
    s.Set(0, 3, L"plain");
    s.Set(1, 0, L"cr\rhere");
    s.Set(1, 1, L" lead");
    ASSERT_WSTR(Out(s, SF_CSV),
        L"\"a,b\",\"say \"\"hi\"\"\",\"two\nlines\",plain\r\n"
        L"\"cr\rhere\", lead,,\r\n");
}

TEST(Csv_EveryRowHasTheSameNumberOfFields)
{
    Sheet s;
    s.Set(0, 0, L"x");
    s.Set(2, 3, L"y");
    ASSERT_WSTR(Out(s, SF_CSV), L"x,,,\r\n,,,\r\n,,,y\r\n");
}

static const wchar_t* const kAwkward[] = {
    L"a,b", L"\"quoted\"", L"\"", L"\"\"", L"x\"\"y", L"line1\r\nline2", L"trailing\n",
    L"\nleading", L",", L" spaced ", L"tab\there", L"h\u00e9llo \u2713", L"semi;colon",
};

TEST(Csv_RoundTripsAwkwardText)
{
    Sheet src;
    const size_t n = sizeof(kAwkward) / sizeof(kAwkward[0]);
    for (size_t i = 0; i < n; ++i)
        src.Set((UINT)i, (UINT)(i % 3) * 2, kAwkward[i]);   // gaps between columns too
    Sheet dst;
    ASSERT_TRUE(RoundTrip(src, SF_CSV, dst));
    for (size_t i = 0; i < n; ++i)
        ASSERT_WSTR(dst.Text((UINT)i, (UINT)(i % 3) * 2), kAwkward[i]);
    ASSERT_EQ(dst.CellCount(), n);                         // the gaps stay empty
}

TEST(Tsv_RoundTripsAwkwardText)
{
    Sheet src;
    const size_t n = sizeof(kAwkward) / sizeof(kAwkward[0]);
    for (size_t i = 0; i < n; ++i)
        src.Set((UINT)i, 1, kAwkward[i]);
    std::wstring tsv = Out(src, SF_TSV);
    ASSERT_CONTAINS(tsv, L"\t\"tab\there\"\r\n");           // a tab inside a field is quoted
    ASSERT_CONTAINS(tsv, L"\ta,b\r\n");                     // a comma is not special in TSV
    Sheet dst;
    ASSERT_EQ(In(dst, tsv, SF_TSV), 0ul);
    for (size_t i = 0; i < n; ++i)
        ASSERT_WSTR(dst.Text((UINT)i, 1), kAwkward[i]);
}

TEST(Csv_ExportsFormulaResultsNotSource)
{
    // The documented choice (grid32.h): interchange formats carry what a
    // cell displays, so a formula exports its result.
    Sheet s;
    s.Set(0, 0, L"5");
    s.Set(0, 1, L"=A1*2");
    ASSERT_WSTR(Out(s, SF_CSV), L"5,10\r\n");
    Sheet back;
    ASSERT_EQ(In(back, Out(s, SF_CSV), SF_CSV), 0ul);
    ASSERT_WSTR(back.Text(0, 1), L"10");
    ASSERT_TRUE(!back.Cell(0, 1)->m_bFormula);
    ASSERT_TRUE(back.Cell(0, 1)->m_eType == CT_Number);
}

TEST(Csv_FieldsBeyondTheGridAreDropped)
{
    // A record wider than the grid used to wrap its overflow into the next
    // row, where the next record then overwrote it.
    Sheet s(3, 8);
    ASSERT_EQ(In(s, L"a,b,c,d,e\nf,g\n", SF_CSV), 0ul);
    ASSERT_WSTR(s.Text(0, 2), L"c");
    ASSERT_WSTR(s.Text(1, 0), L"f");
    ASSERT_WSTR(s.Text(1, 1), L"g");
    ASSERT_EQ(s.CellCount(), (size_t)5);
}

// =============================================================================
// H15 -- stream-in replaces; GM_CLEARALL
// =============================================================================
TEST(StreamIn_ReplacesExistingContent)
{
    Sheet s;
    s.Set(10, 10, L"old");
    s.Set(0, 0, L"zzz");
    ASSERT_EQ(In(s, L"1,2\n", SF_CSV), 0ul);
    ASSERT_TRUE(s.Cell(10, 10) == nullptr);                 // used to survive: a merge, not a load
    ASSERT_WSTR(s.Text(0, 0), L"1");
    ASSERT_WSTR(s.Text(0, 1), L"2");
    ASSERT_EQ(s.CellCount(), (size_t)2);
}

TEST(StreamIn_ReplacesForEveryFormat)
{
    const DWORD formats[] = { SF_CSV, SF_TSV, SF_TEXT, SF_SSF, SF_XML, SF_ODF };
    for (DWORD fmt : formats)
    {
        Sheet src;
        src.Set(0, 0, L"new");
        Sheet dst;
        dst.Set(5, 5, L"stale");
        ASSERT_TRUE(RoundTrip(src, fmt, dst));
        ASSERT_TRUE(dst.Cell(5, 5) == nullptr);
        ASSERT_WSTR(dst.Text(0, 0), L"new");
    }
}

TEST(StreamIn_LeavesUndoAndRedoEmpty)
{
    Sheet s;
    s.Set(0, 0, L"a");
    s.Set(0, 0, L"b");
    s.mgr().OnUndo();
    ASSERT_TRUE(s.mgr().OnCanUndo() == 1);
    ASSERT_TRUE(s.mgr().OnCanRedo() == 1);

    ASSERT_EQ(In(s, L"1,2\n3,4\n", SF_CSV), 0ul);
    // The old history described a sheet that is gone, and the import used to
    // push one undo record per imported cell on top of it.
    ASSERT_TRUE(s.mgr().OnCanUndo() == 0);
    ASSERT_TRUE(s.mgr().OnCanRedo() == 0);
    s.mgr().OnUndo();                                       // nothing to undo
    ASSERT_WSTR(s.Text(1, 1), L"4");
}

TEST(StreamIn_FormulasSeeTheWholeImport)
{
    // The formula comes before the cells it reads; the single recalculation
    // after the load must still see them.
    Sheet s;
    ASSERT_EQ(In(s, L"=A2+B2,\n1,2\n", SF_CSV), 0ul);
    ASSERT_TRUE(s.Cell(0, 0)->m_bFormula);
    ASSERT_WSTR(s.Text(0, 0), L"3");
}

TEST(StreamIn_SendsOneNotificationWhenReplacing)
{
    Sheet s;
    s.Set(3, 3, L"x");
    StubMessageLogClear();
    ASSERT_EQ(In(s, L"1,2\n3,4\n", SF_CSV), 0ul);
    ASSERT_EQ(StubCountNotifications(GN_CONTENTCHANGED), (size_t)1);   // the clear does not add one
}

TEST(StreamIn_ParseFailureLeavesTheGridUntouched)
{
    const wchar_t* bad[] = {
        L"not an ssf file",
        L"SSF2,1\nWB,1\nSH,Sheet1,64,16\n0,0;T:99:short\n",     // length past the end
        L"SSF2,1\nWB,1\nSH,Sheet1,64,16\nx,0;T:1:a\n",          // bad coordinates
        L"SSF2,1\nWB,1\nSH,Sheet1,64,16\n0,0;Y:1:n;V:1:1;FC:3:red\n",   // non-numeric value
    };
    for (const wchar_t* text : bad)
    {
        Sheet s;
        s.Set(2, 2, L"keep");
        s.Set(2, 3, L"me");
        StubMessageLogClear();
        ASSERT_EQ(In(s, text, SF_SSF), (unsigned long)GRID_ERROR_INVALID_PARAMETER);
        ASSERT_WSTR(s.Text(2, 2), L"keep");
        ASSERT_WSTR(s.Text(2, 3), L"me");
        ASSERT_EQ(s.CellCount(), (size_t)2);
        ASSERT_TRUE(s.mgr().OnCanUndo() == 2);
        ASSERT_EQ(StubCountNotifications(GN_CONTENTCHANGED), (size_t)0);
    }
    Sheet x;
    x.Set(0, 0, L"keep");
    ASSERT_EQ(In(x, L"<html/>", SF_XML), (unsigned long)GRID_ERROR_INVALID_PARAMETER);
    ASSERT_EQ(In(x, L"<html/>", SF_ODF), (unsigned long)GRID_ERROR_INVALID_PARAMETER);
    ASSERT_WSTR(x.Text(0, 0), L"keep");
}

TEST(ClearAll_RemovesCellsAndResetsState)
{
    Sheet s;
    s.Set(0, 0, L"a");
    s.Set(4, 5, L"=1+1");
    s.Set(4, 5, L"b");
    s.mgr().OnUndo();
    s.mgr().SetCurrentCell(3, 4);
    GRIDSELECTION sel{ { 1, 1 }, { 6, 7 } };
    s.mgr().OnSetSelection(&sel);
    ASSERT_TRUE(s.mgr().OnCanUndo() > 0 && s.mgr().OnCanRedo() > 0);

    s.mgr().ClearAll();
    ASSERT_EQ(s.CellCount(), (size_t)0);
    ASSERT_TRUE(s.mgr().OnCanUndo() == 0);
    ASSERT_TRUE(s.mgr().OnCanRedo() == 0);
    ASSERT_EQ(s.Current().nRow, 0u);
    ASSERT_EQ(s.Current().nCol, 0u);
    GRIDSELECTION now = s.Selection();
    ASSERT_TRUE(now.start.nRow == 0 && now.start.nCol == 0 && now.end.nRow == 0 && now.end.nCol == 0);
}

TEST(ClearAll_NotifiesOnceOnlyWhenSomethingWasRemoved)
{
    Sheet s;
    s.Set(0, 0, L"a");
    s.Set(1, 0, L"b");
    StubMessageLogClear();
    s.mgr().ClearAll();
    ASSERT_EQ(StubCountNotifications(GN_CONTENTCHANGED), (size_t)1);

    StubMessageLogClear();
    s.mgr().ClearAll();                                     // already empty
    ASSERT_EQ(StubCountNotifications(GN_CONTENTCHANGED), (size_t)0);
}

TEST(ClearAll_GridIsUsableAfterwards)
{
    Sheet s;
    s.Set(0, 0, L"2");
    s.mgr().ClearAll();
    s.Set(0, 0, L"5");
    s.Set(0, 1, L"=A1*3");
    ASSERT_WSTR(s.Text(0, 1), L"15");
    ASSERT_TRUE(s.mgr().OnCanUndo() == 2);
}

// =============================================================================
// H19 -- ODF is well-formed
// =============================================================================
TEST(Odf_IsWellFormedWithEveryPrefixDeclared)
{
    Sheet s;
    s.Set(0, 0, L"<tag> & \"quotes\" 'apos'");
    s.Set(1, 2, L"1234.5");
    s.Set(3, 0, L"TRUE");
    std::wstring odf = Out(s, SF_ODF);
    XmlRoot root;
    ASSERT_XML(odf, root);                                  // <text:p> used to have no text: declaration
    ASSERT_WSTR(root.local, L"document");
    ASSERT_WSTR(root.ns, L"urn:oasis:names:tc:opendocument:xmlns:office:1.0");
    ASSERT_CONTAINS(odf, L"xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"");
    ASSERT_CONTAINS(odf, L"office:mimetype=\"application/vnd.oasis.opendocument.spreadsheet\"");
    ASSERT_CONTAINS(odf, L"office:value-type=\"float\" office:value=\"1234.5\"");
    ASSERT_CONTAINS(odf, L"office:value-type=\"boolean\" office:boolean-value=\"true\"");
    ASSERT_NOT_CONTAINS(odf, L"encoding=");                 // the buffer is UTF-16; the host picks bytes
}

TEST(Odf_EmptySheetIsStillAValidDocument)
{
    Sheet s;
    std::wstring odf = Out(s, SF_ODF);
    XmlRoot root;
    ASSERT_XML(odf, root);
    ASSERT_CONTAINS(odf, L"<table:table-row>");
    Sheet back;
    back.Set(0, 0, L"x");
    ASSERT_EQ(In(back, odf, SF_ODF), 0ul);
    ASSERT_EQ(back.CellCount(), (size_t)0);
}

TEST(Odf_RoundTripsTextAndTypedValues)
{
    Sheet src;
    src.Set(0, 0, L"  two leading, inner   three, trailing  ");
    src.Set(0, 3, L"tab\there");
    src.Set(2, 1, L"line one\nline two\r\nline three");
    src.Set(2, 2, L"<&>\"'");
    src.Set(5, 0, L"0.30000000000000004");
    src.Set(5, 1, L"FALSE");
    src.Set(5, 2, L"2026-05-15");
    src.Set(5, 3, L"=A6*2");
    Sheet dst;
    ASSERT_TRUE(RoundTrip(src, SF_ODF, dst));
    ASSERT_WSTR(dst.Text(0, 0), src.Text(0, 0));
    ASSERT_WSTR(dst.Text(0, 3), L"tab\there");
    ASSERT_WSTR(dst.Text(2, 1), L"line one\nline two\nline three");   // ODF paragraphs: LF
    ASSERT_WSTR(dst.Text(2, 2), L"<&>\"'");
    ASSERT_TRUE(dst.Cell(5, 0)->m_eType == CT_Number);
    ASSERT_TRUE(dst.Cell(5, 0)->m_dValue == src.Cell(5, 0)->m_dValue);   // exact
    ASSERT_TRUE(dst.Cell(5, 1)->m_eType == CT_Boolean && dst.Cell(5, 1)->m_dValue == 0.0);
    ASSERT_TRUE(dst.Cell(5, 2)->m_eType == CT_Date);
    ASSERT_TRUE(dst.Cell(5, 2)->m_dValue == src.Cell(5, 2)->m_dValue);
    ASSERT_TRUE(!dst.Cell(5, 3)->m_bFormula);                // a result, by design
    ASSERT_WSTR(dst.Text(5, 3), src.Text(5, 3));
    ASSERT_EQ(dst.CellCount(), src.CellCount());
}

TEST(Odf_ReadsLibreOfficeRepeatsAndMarkup)
{
    // The shape LibreOffice writes: repeat counts for runs of cells and rows,
    // covered cells under a merge, spans, and encoded spaces.
    const wchar_t* fods =
        L"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
        L"<office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\""
        L" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\""
        L" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" office:version=\"1.2\">\n"
        L"<office:body><office:spreadsheet><table:table table:name=\"Sheet1\">\n"
        L"<table:table-column table:number-columns-repeated=\"1024\"/>\n"
        L"<table:table-row>"
        L"<table:table-cell office:value-type=\"string\"><text:p>a<text:s text:c=\"2\"/>b<text:span text:style-name=\"T1\">c</text:span></text:p></table:table-cell>"
        L"<table:table-cell table:number-columns-repeated=\"2\"/>"
        L"<table:table-cell table:number-columns-spanned=\"2\" office:value-type=\"float\" office:value=\"42\"><text:p>42</text:p></table:table-cell>"
        L"<table:covered-table-cell/>"
        L"<table:table-cell office:value-type=\"string\"><text:p>after</text:p></table:table-cell>"
        L"<table:table-cell table:number-columns-repeated=\"1018\"/>"
        L"</table:table-row>\n"
        L"<table:table-row table:number-rows-repeated=\"3\"><table:table-cell table:number-columns-repeated=\"1024\"/></table:table-row>\n"
        L"<table:table-row table:number-rows-repeated=\"2\">"
        L"<table:table-cell table:number-columns-repeated=\"3\" office:value-type=\"float\" office:value=\"7\"><text:p>7</text:p></table:table-cell>"
        L"</table:table-row>\n"
        L"<table:table-row table:number-rows-repeated=\"1048570\"><table:table-cell table:number-columns-repeated=\"1024\"/></table:table-row>\n"
        L"</table:table></office:spreadsheet></office:body></office:document>\n";
    Sheet s;
    ASSERT_EQ(In(s, fods, SF_ODF), 0ul);
    ASSERT_WSTR(s.Text(0, 0), L"a  bc");
    ASSERT_WSTR(s.Text(0, 3), L"42");
    ASSERT_TRUE(s.Cell(0, 3)->m_eType == CT_Number);
    ASSERT_TRUE(s.Cell(0, 4) == nullptr);                   // the covered cell
    ASSERT_WSTR(s.Text(0, 5), L"after");
    for (UINT r = 4; r <= 5; ++r)
        for (UINT c = 0; c < 3; ++c)
            ASSERT_WSTR(s.Text(r, c), L"7");
    ASSERT_EQ(s.CellCount(), (size_t)(3 + 6));
}

// =============================================================================
// H20 -- SF_TEXT, SF_RTF, SpreadsheetML, and no fake .xlsx
// =============================================================================
TEST(Text_IsPlainTabSeparated)
{
    Sheet s;
    s.Set(0, 0, L"a\"b");
    s.Set(0, 1, L"x\ty");
    s.Set(1, 0, L"two\r\nlines");
    DWORD err = 1;
    std::wstring text = Out(s, SF_TEXT, &err);
    ASSERT_EQ(err, 0ul);                                    // used to "succeed" with an empty string
    ASSERT_WSTR(text, L"a\"b\tx y\r\ntwo lines\t\r\n");
}

TEST(Text_RoundTripsQuotesLiterally)
{
    // Plain text has no quoting, so a leading quote is just a character.
    Sheet src;
    src.Set(0, 0, L"\"quoted\"");
    src.Set(0, 1, L"it's \"fine\"");
    Sheet dst;
    ASSERT_TRUE(RoundTrip(src, SF_TEXT, dst));
    ASSERT_WSTR(dst.Text(0, 0), L"\"quoted\"");
    ASSERT_WSTR(dst.Text(0, 1), L"it's \"fine\"");
}

// Braces balance once escapes (\\, \{, \}) are skipped.
static bool RtfBalanced(const std::wstring& rtf)
{
    int depth = 0;
    for (size_t i = 0; i < rtf.size(); ++i)
    {
        if (rtf[i] == L'\\') { ++i; continue; }
        if (rtf[i] == L'{') ++depth;
        if (rtf[i] == L'}' && --depth < 0) return false;
    }
    return depth == 0;
}

static size_t Count(const std::wstring& hay, const std::wstring& needle)
{
    size_t n = 0;
    for (size_t p = hay.find(needle); p != std::wstring::npos; p = hay.find(needle, p + needle.size()))
        ++n;
    return n;
}

TEST(Rtf_IsATable)
{
    Sheet s;
    s.Set(0, 0, L"{braces}");
    s.Set(0, 2, L"back\\slash");
    s.Set(2, 1, L"caf\u00e9\nnext");
    DWORD err = 1;
    std::wstring rtf = Out(s, SF_RTF, &err);
    ASSERT_EQ(err, 0ul);                                    // used to "succeed" with an empty string
    ASSERT_TRUE(rtf.compare(0, 6, L"{\\rtf1") == 0);
    ASSERT_TRUE(RtfBalanced(rtf));
    ASSERT_EQ(Count(rtf, L"\\trowd"), (size_t)3);           // rows 0..2, the empty one too
    ASSERT_EQ(Count(rtf, L"\\row\r\n"), (size_t)3);
    ASSERT_EQ(Count(rtf, L"\\cell "), (size_t)9);           // 3 x 3
    ASSERT_EQ(Count(rtf, L"\\cellx"), (size_t)9);
    ASSERT_CONTAINS(rtf, L"\\cellx1500\\cellx3000\\cellx4500");   // 100 px columns
    ASSERT_CONTAINS(rtf, L"\\{braces\\}\\cell ");
    ASSERT_CONTAINS(rtf, L"back\\\\slash\\cell ");
    ASSERT_CONTAINS(rtf, L"caf\\u233?\\line next\\cell ");
}

TEST(Xml_IsSpreadsheetML2003)
{
    Sheet s;
    s.Set(0, 0, L"name & <co>");
    s.Set(0, 2, L"1234.5");
    s.mgr().SetCellNumberFormat(0, 2, FMT_CURRENCY);
    s.Set(3, 0, L"TRUE");
    s.Set(3, 1, L"2026-05-15");
    DWORD err = 1;
    std::wstring xml = Out(s, SF_XML, &err);
    ASSERT_EQ(err, 0ul);
    XmlRoot root;
    ASSERT_XML(xml, root);
    ASSERT_WSTR(root.local, L"Workbook");                   // not a bare <worksheet> fragment
    ASSERT_WSTR(root.ns, L"urn:schemas-microsoft-com:office:spreadsheet");
    ASSERT_CONTAINS(xml, L"<?mso-application progid=\"Excel.Sheet\"?>");
    ASSERT_CONTAINS(xml, L"<Worksheet ss:Name=\"Sheet1\">");
    ASSERT_CONTAINS(xml, L"<Data ss:Type=\"String\">name &amp; &lt;co&gt;</Data>");
    ASSERT_CONTAINS(xml, L"<Cell ss:Index=\"3\" ss:StyleID=\"fmt3\"><Data ss:Type=\"Number\">1234.5</Data>");
    ASSERT_CONTAINS(xml, L"<Style ss:ID=\"fmt3\"><NumberFormat ss:Format=\"$#,##0.00\"/></Style>");
    ASSERT_CONTAINS(xml, L"<Row ss:Index=\"4\">");
    ASSERT_CONTAINS(xml, L"<Data ss:Type=\"Boolean\">1</Data>");
    ASSERT_CONTAINS(xml, L"ss:StyleID=\"dt\"><Data ss:Type=\"DateTime\">2026-05-15T00:00:00.000</Data>");
    ASSERT_NOT_CONTAINS(xml, L"encoding=");
}

TEST(Xml_RoundTripsValuesTypesAndFormats)
{
    Sheet src;
    src.Set(0, 0, L"0.1");
    src.Set(0, 1, L"1234.5");
    src.mgr().SetCellNumberFormat(0, 1, FMT_CURRENCY);
    src.Set(0, 2, L"0.000000001");
    src.Set(1, 0, L"FALSE");
    src.Set(1, 1, L"5/15/2026");
    src.mgr().SetCellNumberFormat(1, 1, FMT_DATE_ISO);
    src.Set(2, 0, L"multi\nline & <markup>");
    GRIDCELL text42;                                        // a number-looking *text* cell
    text42.m_wsText = L"42";
    src.mgr().SetCell(2, 1, text42);
    GRIDCELL formulaText;                                   // text that looks like a formula
    formulaText.m_wsText = L"=not a formula";
    src.mgr().SetCell(2, 2, formulaText);
    Sheet dst;
    ASSERT_TRUE(RoundTrip(src, SF_XML, dst));
    ASSERT_TRUE(dst.Cell(0, 0)->m_dValue == 0.1);
    ASSERT_TRUE(dst.Cell(0, 1)->m_dValue == 1234.5);
    ASSERT_EQ(dst.Cell(0, 1)->m_nFormat, (UINT)FMT_CURRENCY);
    ASSERT_WSTR(dst.Text(0, 1), L"$1,234.50");
    ASSERT_TRUE(dst.Cell(0, 2)->m_dValue == 1e-9);
    ASSERT_TRUE(dst.Cell(1, 0)->m_eType == CT_Boolean);
    ASSERT_WSTR(dst.Text(1, 0), L"FALSE");
    ASSERT_TRUE(dst.Cell(1, 1)->m_eType == CT_Date);
    ASSERT_WSTR(dst.Text(1, 1), L"2026-05-15");
    ASSERT_WSTR(dst.Text(2, 0), L"multi\nline & <markup>");
    ASSERT_TRUE(dst.Cell(2, 1)->m_eType == CT_Text);        // a String stays a string
    ASSERT_WSTR(dst.Text(2, 1), L"42");
    ASSERT_TRUE(!dst.Cell(2, 2)->m_bFormula);
    ASSERT_WSTR(dst.Text(2, 2), L"=not a formula");
}

TEST(Xml_ReadsExcelStyleInput)
{
    // Excel writes CRLF-indented XML, an encoding declaration, styles and
    // options blocks, ss:Index gaps, MergeAcross and rich-text ss:Data.
    const wchar_t* xml =
        L"<?xml version=\"1.0\"?>\r\n<?mso-application progid=\"Excel.Sheet\"?>\r\n"
        L"<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\r\n"
        L" xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n"
        L" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\r\n"
        L" xmlns:html=\"http://www.w3.org/TR/REC-html40\">\r\n"
        L" <Worksheet ss:Name=\"Data\">\r\n"
        L"  <Table ss:ExpandedColumnCount=\"5\" ss:ExpandedRowCount=\"4\">\r\n"
        L"   <Column ss:Width=\"60\"/>\r\n"
        L"   <Row>\r\n"
        L"    <Cell ss:MergeAcross=\"1\"><Data ss:Type=\"String\">merged</Data></Cell>\r\n"
        L"    <Cell><Data ss:Type=\"Number\">1.0000000000000001E-9</Data></Cell>\r\n"
        L"    <Cell ss:Index=\"5\"><ss:Data ss:Type=\"String\" xmlns=\"http://www.w3.org/TR/REC-html40\"><B>bold</B> text</ss:Data></Cell>\r\n"
        L"   </Row>\r\n"
        L"   <Row ss:Index=\"4\"><Cell><Data ss:Type=\"DateTime\">2026-05-15T00:00:00.000</Data></Cell></Row>\r\n"
        L"  </Table>\r\n"
        L"  <WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><Panes><Pane><ActiveRow>3</ActiveRow></Pane></Panes></WorksheetOptions>\r\n"
        L" </Worksheet>\r\n"
        L"</Workbook>\r\n";
    Sheet s;
    ASSERT_EQ(In(s, xml, SF_XML), 0ul);
    ASSERT_WSTR(s.Text(0, 0), L"merged");
    ASSERT_TRUE(s.Cell(0, 1) == nullptr);                   // under the merge
    ASSERT_TRUE(s.Cell(0, 2)->m_dValue == 1e-9);
    ASSERT_WSTR(s.Text(0, 4), L"bold text");
    ASSERT_TRUE(s.Cell(3, 0)->m_eType == CT_Date);
    ASSERT_EQ(s.CellCount(), (size_t)4);
}

TEST(Xlsx_FailsExplicitlyInBothDirections)
{
    Sheet s;
    s.Set(0, 0, L"keep");
    std::vector<wchar_t> buf(1024, L'#');
    GCSTREAM st{};
    st.m_cbSize = sizeof(GCSTREAM);
    st.m_pwszBuff = buf.data();
    st.m_cbBuffSize = (UINT)(buf.size() * sizeof(wchar_t));
    st.m_dwFormat = SF_XLSX;
    st.m_pfnCallback = CountCallback;
    g_callbacks = 0;
    s.mgr().OnStreamOut(&st);
    // It used to report success for a bare <worksheet> fragment no program opens.
    ASSERT_EQ(st.m_dwError, (DWORD)GRID_ERROR_NOT_IMPLEMENTED);
    ASSERT_EQ(st.m_cbBuffOut, 0u);
    ASSERT_TRUE(buf[0] == L'\0');
    ASSERT_EQ(g_callbacks, 1);

    ASSERT_EQ(In(s, L"<worksheet><sheetData/></worksheet>", SF_XLSX), (unsigned long)GRID_ERROR_NOT_IMPLEMENTED);
    ASSERT_WSTR(s.Text(0, 0), L"keep");
}

TEST(UnsupportedFormatsAreRefusedNotFaked)
{
    Sheet s;
    s.Set(0, 0, L"keep");
    DWORD err = 0;
    Out(s, 0x1000, &err);
    ASSERT_EQ(err, (DWORD)GRID_ERROR_NOT_IMPLEMENTED);
    Out(s, SF_CSV | SF_TSV, &err);
    ASSERT_EQ(err, (DWORD)GRID_ERROR_NOT_IMPLEMENTED);
    ASSERT_EQ(In(s, L"{\\rtf1 x}", SF_RTF), (unsigned long)GRID_ERROR_NOT_IMPLEMENTED);
    ASSERT_WSTR(s.Text(0, 0), L"keep");
}

TEST(StreamOut_NullBufferIsAnError)
{
    Sheet s;
    GCSTREAM st{};
    st.m_cbSize = sizeof(GCSTREAM);
    st.m_dwFormat = SF_CSV;
    st.m_cbBuffSize = 100;
    st.m_pfnCallback = CountCallback;
    g_callbacks = 0;
    s.mgr().OnStreamOut(&st);                               // used to return silently, "success"
    ASSERT_EQ(st.m_dwError, (DWORD)GRID_ERROR_INVALID_PARAMETER);
    ASSERT_EQ(g_callbacks, 1);
}

// =============================================================================
// M14 -- SSF numbers are exact and locale-independent
// =============================================================================
TEST(Ssf_RoundTripsDoublesExactly)
{
    const wchar_t* values[] = {
        L"0.30000000000000004", L"0.000000001", L"123456789.12345678", L"-2.5e300",
        L"3.141592653589793", L"1e-300", L"0.1",
    };
    Sheet src;
    for (UINT i = 0; i < sizeof(values) / sizeof(values[0]); ++i)
        src.Set(i, 0, values[i]);
    Sheet dst;
    ASSERT_TRUE(RoundTrip(src, SF_SSF, dst));
    for (UINT i = 0; i < sizeof(values) / sizeof(values[0]); ++i)
    {
        // to_wstring wrote six decimals: 1e-9 came back as 0, and so on.
        if (!(dst.Cell(i, 0)->m_dValue == src.Cell(i, 0)->m_dValue))
        {
            std::cerr << "  (" << Narrow(values[i]) << ")\n";
            FAIL_HERE("value changed in an SSF round trip");
        }
    }
}

TEST(Ssf_WritesShortestRoundTripNumbers)
{
    Sheet s;
    s.Set(0, 0, L"0.1");
    s.Set(1, 0, L"1e-9");
    std::wstring ssf = Out(s, SF_SSF);
    ASSERT_CONTAINS(ssf, L";V:3:0.1;");
    ASSERT_CONTAINS(ssf, L";V:5:1e-09;");
    ASSERT_CONTAINS(ssf, L";FS:2:12;");                     // a float, written the same way
    ASSERT_NOT_CONTAINS(ssf, L"0.100000");
}

TEST(Ssf_FontPointSizeRoundTripsExactly)
{
    Sheet src;
    src.Set(0, 0, L"x");
    CELLFONT fi = src.Cell(0, 0)->fontInfo;
    fi.m_fPointSize = 10.3f;
    src.mgr().SetCellFormat(0, 0, fi);
    Sheet dst;
    ASSERT_TRUE(RoundTrip(src, SF_SSF, dst));
    ASSERT_TRUE(dst.Cell(0, 0)->fontInfo.m_fPointSize == 10.3f);
}

namespace {
// A C++ locale that writes 1234.5 as "1.234,5" -- what a host gets from
// std::locale::global(std::locale("de_DE.UTF-8")), built here from facets so
// the test does not depend on which locales the machine has installed.
struct GermanPunctW : std::numpunct<wchar_t> {
protected:
    wchar_t do_decimal_point() const override { return L','; }
    wchar_t do_thousands_sep() const override { return L'.'; }
    std::string do_grouping() const override { return "\3"; }
};
struct GermanPunctA : std::numpunct<char> {
protected:
    char do_decimal_point() const override { return ','; }
    char do_thousands_sep() const override { return '.'; }
    std::string do_grouping() const override { return "\3"; }
};
struct GlobalLocaleScope {
    std::locale previous;
    explicit GlobalLocaleScope(const std::locale& loc) : previous(std::locale::global(loc)) {}
    ~GlobalLocaleScope() { std::locale::global(previous); }
};
}

TEST(Ssf_IgnoresAGlobalCppLocale)
{
    // Row 1500 written through a grouping locale came out as "1.500", and the
    // record landed on row 1 -- or, read with stricter parsing, not at all.
    std::locale german(std::locale(std::locale::classic(), new GermanPunctW), new GermanPunctA);
    GlobalLocaleScope scope(german);
    Sheet src(16, 3000);
    src.Set(1500, 2, L"1234.5");
    std::wstring ssf = Out(src, SF_SSF);
    ASSERT_CONTAINS(ssf, L"\n1500,2;");
    ASSERT_CONTAINS(ssf, L";V:6:1234.5;");
    Sheet dst(16, 3000);
    ASSERT_EQ(In(dst, ssf, SF_SSF), 0ul);
    ASSERT_TRUE(dst.Cell(1500, 2) != nullptr);
    ASSERT_TRUE(dst.Cell(1500, 2)->m_dValue == 1234.5);
    ASSERT_EQ(dst.CellCount(), (size_t)1);
}

TEST(Ssf_IgnoresACommaDecimalCLocale)
{
    // std::to_wstring follows the C locale's LC_NUMERIC. Only checkable where
    // such a locale is installed; the facet test above covers the rest.
    const char* names[] = { "de_DE.UTF-8", "de_DE.utf8", "fr_FR.UTF-8", "fr_FR.utf8", "de_DE", "nl_NL.UTF-8" };
    const char* found = nullptr;
    for (const char* n : names)
        if (std::setlocale(LC_NUMERIC, n)) { found = n; break; }
    if (!found)
    {
        std::cout << "  note: no comma-decimal C locale installed; C-locale check skipped\n";
        return;
    }
    Sheet src;
    src.Set(0, 0, L"1.5");
    std::wstring ssf = Out(src, SF_SSF);
    Sheet dst;
    DWORD err = In(dst, ssf, SF_SSF);
    std::setlocale(LC_NUMERIC, "C");
    ASSERT_CONTAINS(ssf, L";V:3:1.5;");
    ASSERT_EQ(err, 0ul);
    ASSERT_TRUE(dst.Cell(0, 0)->m_dValue == 1.5);
}

TEST(Ssf_ReadsFilesFromTheOldWriter)
{
    // to_wstring output: six decimals, and a comma under a comma-decimal locale.
    const wchar_t* old =
        L"SSF2,1\nWB,1\nSH,Sheet1,64,16\n"
        L"0,0;Y:1:n;V:8:1.500000;T:3:1.5;FF:5:Arial;FS:9:12.000000\n"
        L"1,0;Y:1:n;V:8:2,500000;T:3:2.5;FF:5:Arial;FS:9:11,500000\n";
    Sheet s;
    ASSERT_EQ(In(s, old, SF_SSF), 0ul);
    ASSERT_TRUE(s.Cell(0, 0)->m_dValue == 1.5);
    ASSERT_TRUE(s.Cell(1, 0)->m_dValue == 2.5);
    ASSERT_TRUE(s.Cell(0, 0)->fontInfo.m_fPointSize == 12.0f);
    ASSERT_TRUE(s.Cell(1, 0)->fontInfo.m_fPointSize == 11.5f);
}

TEST(Ssf_RoundTripsLineBreaksInText)
{
    // SSF values are length-prefixed, but the reader split the file into
    // lines first, cutting a multi-line cell's record in two.
    Sheet src;
    src.Set(0, 0, L"line1\nline2");
    src.Set(0, 1, L"crlf\r\nhere");
    src.Set(1, 0, L"after");
    Sheet dst;
    ASSERT_TRUE(RoundTrip(src, SF_SSF, dst));
    ASSERT_WSTR(dst.Text(0, 0), L"line1\nline2");
    ASSERT_WSTR(dst.Text(0, 1), L"crlf\r\nhere");
    ASSERT_WSTR(dst.Text(1, 0), L"after");
    ASSERT_EQ(dst.CellCount(), (size_t)3);
}

TEST(Ssf_RoundTripsFormulasAndFormats)
{
    Sheet src;
    src.Set(0, 0, L"4");
    src.Set(0, 1, L"=A1*2.5");
    src.Set(1, 0, L"1234.5");
    src.mgr().SetCellNumberFormat(1, 0, FMT_CURRENCY);
    Sheet dst;
    ASSERT_TRUE(RoundTrip(src, SF_SSF, dst));
    ASSERT_TRUE(dst.Cell(0, 1)->m_bFormula);                // SSF keeps the source
    ASSERT_WSTR(dst.Cell(0, 1)->m_wsFormula, L"A1*2.5");
    ASSERT_WSTR(dst.Text(0, 1), L"10");
    ASSERT_WSTR(dst.Text(1, 0), L"$1,234.50");
    ASSERT_TRUE(dst.mgr().OnCanUndo() == 0);
}

// =============================================================================
// P1 -- stream-out walks the cells that exist; one retry at most
// =============================================================================
TEST(StreamOut_CostFollowsTheCellsNotTheGrid)
{
    // 20,000 x 5,000 = 100 million positions. The old writers visited every
    // one of them, per format; three cells should take next to no time.
    Sheet s(5000, 20000);
    ASSERT_TRUE(s.ok());
    s.Set(0, 0, L"a");
    s.Set(1, 2, L"b");
    s.Set(3, 1, L"=1+1");
    const DWORD formats[] = { SF_CSV, SF_TSV, SF_TEXT, SF_RTF, SF_SSF, SF_XML, SF_ODF };
    auto start = std::chrono::steady_clock::now();
    for (DWORD fmt : formats)
    {
        DWORD err = 1;
        Out(s, fmt, &err, 1u << 16);
        ASSERT_EQ(err, 0ul);
    }
    double seconds = std::chrono::duration<double>(std::chrono::steady_clock::now() - start).count();
    if (seconds > 1.0)
    {
        std::cerr << "  (" << seconds << " s for 3 cells)\n";
        FAIL_HERE("stream-out cost scales with the grid, not the cells");
    }
    // ...and the output is the populated block, not the whole grid.
    ASSERT_WSTR(Out(s, SF_CSV), L"a,,\r\n,,b\r\n,,\r\n,2,\r\n");
}

TEST(StreamOut_IsRowMajorWhateverTheInsertionOrder)
{
    Sheet s;
    s.Set(2, 1, L"c");
    s.Set(0, 3, L"b");
    s.Set(1, 0, L"x");
    s.Set(0, 0, L"a");
    std::wstring ssf = Out(s, SF_SSF);
    size_t p00 = ssf.find(L"\n0,0;"), p03 = ssf.find(L"\n0,3;"), p10 = ssf.find(L"\n1,0;"), p21 = ssf.find(L"\n2,1;");
    ASSERT_TRUE(p00 != std::wstring::npos && p00 < p03 && p03 < p10 && p10 < p21);
    ASSERT_WSTR(Out(s, SF_SSF), ssf);                       // deterministic
    ASSERT_WSTR(Out(s, SF_XML), Out(s, SF_XML));
}

TEST(StreamOut_ReportsTheExactSizeNeeded)
{
    Sheet s;
    for (UINT r = 0; r < 40; ++r)
        s.Set(r, r % 5, L"some text " + std::to_wstring(r));
    const std::wstring full = Out(s, SF_SSF);
    ASSERT_TRUE(full.size() > 100);

    // Too small -- including exactly one wchar_t short, with no room for the
    // terminator: an error, an empty buffer, and the size in bytes needed.
    for (size_t small : { (size_t)8, full.size() })
    {
        std::vector<wchar_t> buf(small, L'#');
        GCSTREAM st{};
        st.m_cbSize = sizeof(GCSTREAM);
        st.m_pwszBuff = buf.data();
        st.m_cbBuffSize = (UINT)(buf.size() * sizeof(wchar_t));
        st.m_dwFormat = SF_SSF;
        s.mgr().OnStreamOut(&st);
        ASSERT_EQ(st.m_dwError, (DWORD)GRID_ERROR_OUT_OF_RANGE);
        ASSERT_TRUE(buf[0] == L'\0');
        ASSERT_EQ((size_t)st.m_cbBuffOut, (full.size() + 1) * sizeof(wchar_t));
    }

    // One retry at exactly that size succeeds.
    std::vector<wchar_t> buf(full.size() + 1, L'#');
    GCSTREAM st{};
    st.m_cbSize = sizeof(GCSTREAM);
    st.m_pwszBuff = buf.data();
    st.m_cbBuffSize = (UINT)(buf.size() * sizeof(wchar_t));
    st.m_dwFormat = SF_SSF;
    s.mgr().OnStreamOut(&st);
    ASSERT_EQ(st.m_dwError, 0ul);
    ASSERT_EQ((size_t)st.m_cbBuffOut, full.size() * sizeof(wchar_t));
    ASSERT_WSTR(std::wstring(buf.data()), full);
}

// ---- runner ----------------------------------------------------------------
int main()
{
    std::cout << "Running " << Tests().size() << " stream tests\n";
    for (auto& t : Tests())
    {
        g_currentTest = t.name;
        int before = g_failCount;
        t.fn();
        if (g_failCount == before)
            std::cout << "  ok   " << t.name << "\n";
    }
    if (g_failCount == 0) { std::cout << "\nAll tests passed.\n"; return 0; }
    std::cout << "\n" << g_failCount << " failure(s).\n";
    return 1;
}
