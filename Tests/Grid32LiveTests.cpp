// Grid32 live-window tests: a real Grid32 window, its real inline EDIT
// control and its real paint code, driven through the public message API.
//
// The other suites run the manager headless (Tests/linux/GridMgrTests.cpp
// against a Win32 shim), so nothing there exercises the window procedures,
// the edit-control subclass, or GDI. This one does. It includes only
// grid32.h, the header a host compiles against, which also proves that header
// carries no C++ library types (AUDIT_2026-09 M4).
//
// Usage:  Grid32LiveTests.exe <path\to\Grid32.dll> [outdir]
// Exit code = number of failures. With an outdir, the rendered images are
// written there as BMPs for a person to look at.

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <commctrl.h>
#include <cstdio>
#include <cwchar>
#include <string>
#include <vector>
#include "../Grid32/grid32.h"

static int g_failures = 0;
static int g_checks = 0;
static std::vector<UINT> g_notifications;

#define CHECK(cond, what) do { ++g_checks; if (cond) { wprintf(L"  PASS  %s\n", what); } \
    else { ++g_failures; wprintf(L"  FAIL  %s  (%hs:%d)\n", what, __FILE__, __LINE__); } } while (0)

static LRESULT CALLBACK HostProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    if (msg == WM_NOTIFY)
        g_notifications.push_back(reinterpret_cast<NMHDR*>(lParam)->code);
    return DefWindowProcW(hWnd, msg, wParam, lParam);
}

static size_t Count(UINT code)
{
    size_t n = 0;
    for (UINT c : g_notifications)
        if (c == code)
            ++n;
    return n;
}

static HWND CreateGrid(HWND host, size_t cols, size_t rows)
{
    GRIDCREATESTRUCT gcs{};
    gcs.cbSize = sizeof(gcs);
    gcs.nWidth = cols;
    gcs.nHeight = rows;
    gcs.nDefColWidth = 100;
    gcs.nDefRowHeight = 24;
    gcs.clrCurrentCellHighlightBox = RGB(255, 0, 0);
    gcs.clrSelectionOverlay = RGB(0, 0, 255);
    HWND grid = CreateWindowExW(0, L"Grid32", L"", WS_CHILD | WS_VISIBLE | GS_SPREADSHEET,
        0, 0, 820, 520, host, (HMENU)1, GetModuleHandleW(nullptr), &gcs);
    return grid;
}

static std::wstring CellText(HWND grid, UINT row, UINT col)
{
    wchar_t buf[256] = L"";
    GRIDPOINT pt{ row, col };
    GRID_GETTEXT gt{ buf, 256 };
    SendMessageW(grid, GM_GETCELLTEXT, (WPARAM)&pt, (LPARAM)&gt);
    return buf;
}

static void SetText(HWND grid, UINT row, UINT col, const wchar_t* text)
{
    GRIDPOINT pt{ row, col };
    SendMessageW(grid, GM_SETCELLTEXT, (WPARAM)&pt, (LPARAM)text);
}

static void Select(HWND grid, UINT r0, UINT c0, UINT r1, UINT c1)
{
    GRIDSELECTION sel{ { r0, c0 }, { r1, c1 } };
    SendMessageW(grid, GM_SETSELECTION, 0, (LPARAM)&sel);
}

static HWND Editor(HWND grid)
{
    return FindWindowExW(grid, nullptr, L"Edit", nullptr);
}

// IsWindowVisible is false for every child of the (never shown) host, so look
// at the editor's own WS_VISIBLE bit.
static bool EditorOpen(HWND edit)
{
    return (GetWindowLongW(edit, GWL_STYLE) & WS_VISIBLE) != 0;
}

static std::wstring EditorText(HWND grid)
{
    wchar_t buf[256] = L"";
    GetWindowTextW(Editor(grid), buf, 256);
    return buf;
}

// A 32-bit top-down DIB section and a DC selected into it.
struct Canvas
{
    HDC dc = nullptr;
    HBITMAP bmp = nullptr;
    HGDIOBJ old = nullptr;
    DWORD* bits = nullptr;
    int w = 0, h = 0;
    Canvas(int width, int height) : w(width), h(height)
    {
        BITMAPINFO bi{};
        bi.bmiHeader.biSize = sizeof(bi.bmiHeader);
        bi.bmiHeader.biWidth = w;
        bi.bmiHeader.biHeight = -h;
        bi.bmiHeader.biPlanes = 1;
        bi.bmiHeader.biBitCount = 32;
        bi.bmiHeader.biCompression = BI_RGB;
        HDC screen = GetDC(nullptr);
        dc = CreateCompatibleDC(screen);
        ReleaseDC(nullptr, screen);
        bmp = CreateDIBSection(dc, &bi, DIB_RGB_COLORS, (void**)&bits, nullptr, 0);
        old = SelectObject(dc, bmp);
        RECT r{ 0, 0, w, h };
        FillRect(dc, &r, (HBRUSH)GetStockObject(WHITE_BRUSH));
    }
    ~Canvas()
    {
        SelectObject(dc, old);
        DeleteObject(bmp);
        DeleteDC(dc);
    }
    // Pixels of exactly this colour (COLORREF) in the rectangle.
    size_t CountColor(COLORREF c, RECT r) const
    {
        const DWORD want = (GetRValue(c) << 16) | (GetGValue(c) << 8) | GetBValue(c);
        size_t n = 0;
        for (int y = max(0L, r.top); y < min((LONG)h, r.bottom); ++y)
            for (int x = max(0L, r.left); x < min((LONG)w, r.right); ++x)
                if ((bits[y * w + x] & 0xFFFFFF) == want)
                    ++n;
        return n;
    }
    size_t CountNotWhite() const
    {
        size_t n = 0;
        for (int i = 0; i < w * h; ++i)
            if ((bits[i] & 0xFFFFFF) != 0xFFFFFF)
                ++n;
        return n;
    }
    void Save(const std::wstring& path) const
    {
        BITMAPFILEHEADER fh{};
        BITMAPINFOHEADER ih{};
        ih.biSize = sizeof(ih);
        ih.biWidth = w;
        ih.biHeight = -h;
        ih.biPlanes = 1;
        ih.biBitCount = 32;
        ih.biCompression = BI_RGB;
        fh.bfType = 0x4D42;
        fh.bfOffBits = sizeof(fh) + sizeof(ih);
        fh.bfSize = fh.bfOffBits + (DWORD)(w * h * 4);
        FILE* f = nullptr;
        if (_wfopen_s(&f, path.c_str(), L"wb") || !f)
            return;
        fwrite(&fh, sizeof(fh), 1, f);
        fwrite(&ih, sizeof(ih), 1, f);
        fwrite(bits, 4, (size_t)w * h, f);
        fclose(f);
    }
};

int wmain(int argc, wchar_t** argv)
{
    if (argc < 2)
    {
        wprintf(L"usage: Grid32LiveTests <Grid32.dll> [outdir]\n");
        return 2;
    }
    const std::wstring outDir = argc >= 3 ? argv[2] : L"";
    HMODULE hGrid = LoadLibraryW(argv[1]);
    if (!hGrid)
    {
        wprintf(L"cannot load %s (error %lu)\n", argv[1], GetLastError());
        return 2;
    }

    WNDCLASSW wc{};
    wc.lpfnWndProc = HostProc;
    wc.hInstance = GetModuleHandleW(nullptr);
    wc.lpszClassName = L"Grid32LiveTestsHost";
    RegisterClassW(&wc);
    // Never shown: everything below runs through messages and WM_PRINTCLIENT.
    HWND host = CreateWindowExW(0, wc.lpszClassName, L"host", WS_OVERLAPPEDWINDOW,
        0, 0, 860, 600, nullptr, nullptr, wc.hInstance, nullptr);
    HWND grid = CreateGrid(host, 26, 200);
    CHECK(grid != nullptr, L"a Grid32 window is created");
    if (!grid)
        return 1;
    HWND edit = Editor(grid);
    CHECK(edit != nullptr, L"the grid owns an inline EDIT control");

    wprintf(L"Typing\n");
    g_notifications.clear();
    SendMessageW(grid, WM_CHAR, L'h', 1);
    CHECK(EditorOpen(edit), L"a printable WM_CHAR opens the editor (H14)");
    CHECK(EditorText(grid) == L"h", L"...and the editor received the character");
    SendMessageW(edit, WM_CHAR, L'i', 1);
    SendMessageW(edit, WM_KEYDOWN, VK_RETURN, 1);
    CHECK(!EditorOpen(edit), L"Enter in the editor closes it");
    CHECK(CellText(grid, 0, 0) == L"hi", L"...and commits \"hi\" into A1");
    CHECK(Count(GN_EDITSTART) == 1 && Count(GN_EDITCOMMIT) == 1, L"GN_EDITSTART and GN_EDITCOMMIT are sent once each");
    CHECK(Count(GN_CONTENTCHANGED) == 1, L"one GN_CONTENTCHANGED for the edit");

    g_notifications.clear();
    SendMessageW(grid, WM_KEYDOWN, VK_F1, 1);
    CHECK(!EditorOpen(edit), L"F1 does not open the editor (H14)");
    SendMessageW(grid, WM_KEYDOWN, VK_RETURN, 1);
    CHECK(!EditorOpen(edit) && Count(GN_CONTENTCHANGED) == 0, L"Enter on the grid neither edits nor dirties (S11)");
    CHECK(Count(NM_RETURN) == 1, L"...but the host still gets NM_RETURN");

    wprintf(L"Escape\n");
    g_notifications.clear();
    SendMessageW(grid, WM_CHAR, L'z', 1);
    SendMessageW(edit, WM_KEYDOWN, VK_ESCAPE, 1);
    CHECK(!EditorOpen(edit), L"Escape closes the editor");
    CHECK(CellText(grid, 0, 0) == L"hi", L"...and leaves A1 as it was");
    CHECK(Count(GN_EDITCANCEL) == 1 && Count(GN_EDITCOMMIT) == 0, L"GN_EDITCANCEL, and no commit (S12)");
    SendMessageW(grid, WM_CHAR, L'k', 1);
    SendMessageW(edit, WM_KEYDOWN, VK_RETURN, 1);
    CHECK(CellText(grid, 0, 0) == L"k", L"the next edit commits normally (nothing left armed)");

    wprintf(L"F2, the formula source and the editor's own cell\n");
    SetText(grid, 1, 0, L"=1+2");
    GRIDPOINT a2{ 1, 0 };
    SendMessageW(grid, GM_SETCURRENTCELL, 0, (LPARAM)&a2);
    SendMessageW(grid, WM_KEYDOWN, VK_F2, 1);
    CHECK(EditorOpen(edit) && EditorText(grid) == L"=1+2", L"F2 edits the formula, not its result");
    SetWindowTextW(edit, L"=2+2");
    GRIDPOINT away{ 10, 5 };
    SendMessageW(grid, GM_SETCURRENTCELL, 0, (LPARAM)&away);     // moving away commits first
    CHECK(CellText(grid, 1, 0) == L"4", L"moving the cell commits into the cell being edited (S2)");
    CHECK(CellText(grid, 10, 5) == L"", L"...not into the cell moved to");
    GRIDPOINT cur{};
    SendMessageW(grid, GM_GETCURRENTCELL, 0, (LPARAM)&cur);
    CHECK(cur.nRow == 10 && cur.nCol == 5, L"GM_GETCURRENTCELL reports the new cell");

    SendMessageW(grid, WM_CHAR, L'7', 1);
    CHECK(SendMessageW(grid, GM_COMMITEDIT, TRUE, 0) == TRUE, L"GM_COMMITEDIT reports an open edit (S5)");
    CHECK(CellText(grid, 10, 5) == L"7", L"...and commits it");
    CHECK(SendMessageW(grid, GM_COMMITEDIT, TRUE, 0) == FALSE, L"...and FALSE when nothing is open");

    wprintf(L"Formatting through the plain-struct API (M4, H7, H11, H10)\n");
    SetText(grid, 3, 1, L"Bold me");
    Select(grid, 3, 1, 3, 2);
    GCCELLCHARFORMAT cf{};
    cf.m_cbSize = sizeof(cf);
    cf.m_format.bWeight = 700;
    cf.m_dwMask = GCFM_BOLD;
    CHECK(SendMessageW(grid, GM_SETCHARFORMAT, SCF_SELECTION | SCF_MASKED, (LPARAM)&cf) == TRUE,
          L"GM_SETCHARFORMAT reports success (H17)");
    FONTINFO fi{};
    SendMessageW(grid, GM_GETCHARFORMAT, 0, (LPARAM)&fi);
    CHECK(fi.bWeight == 700 && wcscmp(fi.m_szFontFace, L"Arial") == 0, L"bold applied, face kept (masked)");
    CHECK(SendMessageW(grid, GM_SETBACKGROUND, SCF_SELECTION, RGB(255, 255, 0)) == TRUE, L"GM_SETBACKGROUND");
    CHECK((COLORREF)SendMessageW(grid, GM_GETBACKGROUND, 0, 0) == RGB(255, 255, 0), L"GM_GETBACKGROUND reads it back");
    SetText(grid, 6, 1, L"Merged title");
    GRIDSELECTION block{ { 6, 1 }, { 7, 3 } };
    CHECK(SendMessageW(grid, GM_MERGECELLS, TRUE, (LPARAM)&block) == TRUE, L"GM_MERGECELLS merges B7:D8");
    GRIDPOINT inside{ 7, 2 };
    SendMessageW(grid, GM_SETCURRENTCELL, 0, (LPARAM)&inside);
    MERGERANGE mr{};
    CHECK(SendMessageW(grid, GM_GETMERGE, 0, (LPARAM)&mr) == TRUE && mr.end.nCol == 3,
          L"a cell inside the block reports the merge");
    GCCELLINFO info{};
    info.m_cbSize = sizeof(info);
    info.m_dwMask = GCI_TEXT | GCI_FORMAT;
    info.m_pwszText = L"1234.5";
    info.m_nFormat = FMT_CURRENCY;
    GRIDPOINT c10{ 9, 2 };
    CHECK(SendMessageW(grid, GM_SETCELL, (WPARAM)&c10, (LPARAM)&info) == TRUE, L"GM_SETCELL takes a GCCELLINFO");
    CHECK(CellText(grid, 9, 2) == L"$1,234.50", L"...text and number format applied");

    wprintf(L"Painting\n");
    SendMessageW(grid, WM_SIZE, 0, MAKELPARAM(820, 520));
    {
        Canvas canvas(820, 520);
        SendMessageW(grid, WM_PRINTCLIENT, (WPARAM)canvas.dc, PRF_CLIENT);
        // B4:C4 is yellow: row 4 starts at 24 (header) + 3 * 24, column B at 70 + 100.
        RECT b4{ 170 + 4, 24 + 72 + 4, 270 - 4, 24 + 96 - 4 };
        CHECK(canvas.CountColor(RGB(255, 255, 0), b4) > 100, L"the highlighted cell paints yellow");
        CHECK(canvas.CountNotWhite() > 10000, L"headers, grid lines and text were drawn");
        if (!outDir.empty())
            canvas.Save(outDir + L"\\live_paint.bmp");
    }

    wprintf(L"Rendering for print (H12)\n");
    {
        Canvas page(1700, 1000);
        GCRENDERSTRUCT rs{};
        rs.m_cbSize = sizeof(rs);
        rs.m_hDC = page.dc;
        rs.m_range = { { 0, 0 }, { 10, 5 } };
        rs.m_bGridLines = TRUE;
        rs.m_bHeaders = TRUE;
        rs.m_nLogPixelsX = rs.m_nLogPixelsY = 192;      // twice the screen
        CHECK(SendMessageW(grid, GM_RENDERRANGE, 0, (LPARAM)&rs) == TRUE, L"GM_RENDERRANGE succeeds");
        // B4 at 2x: x from (70 + 100) * 2, y from (24 + 72) * 2.
        RECT b4{ 340 + 8, 192 + 8, 540 - 8, 240 - 8 };
        CHECK(page.CountColor(RGB(255, 255, 0), b4) > 1000, L"the print render scales the cells");
        GRIDSELECTION used{};
        CHECK(SendMessageW(grid, GM_GETUSEDRANGE, 0, (LPARAM)&used) == TRUE && used.end.nRow == 10,
              L"GM_GETUSEDRANGE covers the populated cells");
        if (!outDir.empty())
            page.Save(outDir + L"\\live_render.bmp");
    }

    wprintf(L"Undo notifications (S13)\n");
    HWND fresh = CreateGrid(host, 10, 10);
    g_notifications.clear();
    CHECK(SendMessageW(fresh, WM_UNDO, 0, 0) == FALSE && Count(GN_CELLCHANGED) == 0,
          L"nothing to undo: FALSE and no GN_CELLCHANGED");
    SetText(fresh, 2, 2, L"x");
    g_notifications.clear();
    CHECK(SendMessageW(fresh, WM_UNDO, 0, 0) == TRUE && Count(GN_CELLCHANGED) == 1, L"an undo sends one GN_CELLCHANGED");
    GRIDSELECTION bad{ { 50, 0 }, { 60, 0 } };
    g_notifications.clear();
    CHECK(SendMessageW(fresh, GM_SETSELECTION, 0, (LPARAM)&bad) == FALSE && Count(GN_SELCHANGE) == 0,
          L"a rejected selection sends no GN_SELCHANGE");
    DestroyWindow(fresh);

    wprintf(L"Coordinates past 65535 (M2, M3)\n");
    HWND tall = CreateGrid(host, 4, 200000);
    SetText(tall, 150000, 3, L"deep");
    CHECK(CellText(tall, 150000, 3) == L"deep", L"GM_SETCELLTEXT / GM_GETCELLTEXT at row 150001");
    GRIDPOINT deep{ 199999, 2 };
    SendMessageW(tall, GM_SETCURRENTCELL, 0, (LPARAM)&deep);
    SendMessageW(tall, GM_GETCURRENTCELL, 0, (LPARAM)&cur);
    CHECK(cur.nRow == 199999 && cur.nCol == 2, L"GM_SETCURRENTCELL / GM_GETCURRENTCELL at the last row");
    CHECK(SendMessageW(tall, GM_SETCURRENTCELL, SETBYCOORDINATE, (LPARAM)L"$D$150001") == TRUE, L"Go To a deep reference (S16)");
    SendMessageW(tall, GM_GETCURRENTCELL, 0, (LPARAM)&cur);
    CHECK(cur.nRow == 150000 && cur.nCol == 3, L"...lands on it");
    CHECK(SendMessageW(tall, GM_ENUMCELLS, 0, 0) == 1, L"GM_ENUMCELLS counts populated cells (H18)");
    DestroyWindow(tall);

    wprintf(L"Streams (H17)\n");
    {
        std::vector<wchar_t> buf(1 << 16);
        GCSTREAM st{};
        st.m_cbSize = sizeof(st);
        st.m_pwszBuff = buf.data();
        st.m_cbBuffSize = (UINT)(buf.size() * sizeof(wchar_t));
        st.m_dwFormat = SF_SSF;
        CHECK(SendMessageW(grid, GM_STREAMOUT, 0, (LPARAM)&st) == TRUE && st.m_dwError == 0,
              L"GM_STREAMOUT reports success");
        CHECK(wcsstr(buf.data(), L"Merged title") != nullptr, L"...and wrote the sheet");
    }

    DestroyWindow(host);
    wprintf(L"\n%d checks, %d failed\n", g_checks, g_failures);
    return g_failures;
}
