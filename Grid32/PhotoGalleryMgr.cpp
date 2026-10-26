#include "pch.h"
#include "PhotoGalleryMgr.h"
#include <windowsx.h>
#include <cmath>

#pragma comment(lib, "comctl32.lib")

namespace {

const UINT_PTR IDT_SETTLE = 1;
const UINT_PTR IDT_RETRY = 2;
const UINT kSettleMs = 200;                             // quiet time before tiles are re-sharpened
const UINT kRetryMs = 400;                              // cloud preview still on its way: ask again
const BYTE kMaxRetries = 15;
const ULONGLONG kAutoKeyBit = 0x8000000000000000ull;    // keys the control invents (internal mode)
const ULONGLONG kIndexKeyBit = 0x4000000000000000ull;   // owner-data items the host gave no key
const int kBulkLimit = 32;                              // more changes than this: one PGGN_ITEMCHANGED, iItem -1

Gdiplus::Color GpColor(COLORREF c, BYTE a = 255)
{
    return Gdiplus::Color(a, GetRValue(c), GetGValue(c), GetBValue(c));
}

COLORREF Blend(COLORREF a, COLORREF b, int pctB)
{
    auto mix = [pctB](int x, int y) { return (x * (100 - pctB) + y * pctB) / 100; };
    return RGB(mix(GetRValue(a), GetRValue(b)), mix(GetGValue(a), GetGValue(b)), mix(GetBValue(a), GetBValue(b)));
}

int Clamp(int v, int lo, int hi)
{
    return v < lo ? lo : (v > hi ? hi : v);
}

UINT DpiOf(HWND hWnd)
{
    UINT dpi = GetDpiForWindow(hWnd);
    if (dpi == 0)
    {
        HDC hdc = GetDC(hWnd);
        dpi = (UINT)GetDeviceCaps(hdc, LOGPIXELSX);
        ReleaseDC(hWnd, hdc);
    }
    return dpi ? dpi : 96;
}

} // namespace

// ---- Lifetime -------------------------------------------------------------------

CPhotoGalleryMgr::CPhotoGalleryMgr()
{
    m_metrics.cbSize = sizeof(PGGTILEMETRICS);
    m_metrics.mask = PGGTM_ALL;
    m_metrics.mode = PGGTS_AUTOFIT;
    m_metrics.cxTile = 160;
    m_metrics.cxMin = 64;
    m_metrics.cxMax = 480;
    m_metrics.nColumns = 3;
    m_metrics.nAspectX = 1;
    m_metrics.nAspectY = 1;
    m_metrics.cxGap = 4;
    m_metrics.cyGap = 4;
    m_hostMetrics = m_metrics;
    m_dZoomTile = m_metrics.cxTile;

    m_colors.cbSize = sizeof(PGGCOLORS);
    m_colors.clrBk = m_colors.clrPlaceholder = m_colors.clrSelection = m_colors.clrFocus = m_colors.clrHot = CLR_DEFAULT;

    m_loaderSettings.cbSize = sizeof(PGGLOADERSETTINGS);
    m_loaderSettings.flags = PGGLF_SHELLCACHE | PGGLF_DECODE | PGGLF_ORIENTATION | PGGLF_CLOUDTHUMBS;
    m_loaderSettings.nThreads = 0;
    m_loaderSettings.nCacheBudgetMB = 256;
    m_loaderSettings.nPrefetchScreens = 1;
    m_cache.SetBudget((size_t)m_loaderSettings.nCacheBudgetMB * 1024 * 1024);

    m_layout = CreateUniformLayout();
    m_pathBuf.assign(32768, L'\0');
}

CPhotoGalleryMgr::~CPhotoGalleryMgr()
{
    StopLoader();
    FreeBackBuffer();
}

bool CPhotoGalleryMgr::OnNcCreate(HWND hWnd)
{
    m_hWnd = hWnd;
    return m_gdiplus.Ok();
}

int CPhotoGalleryMgr::OnCreate(const CREATESTRUCT* pcs)
{
    m_dwStyle = pcs->style;
    m_dpi = (int)DpiOf(m_hWnd);
    RECT rc;
    GetClientRect(m_hWnd, &rc);
    m_cxClient = rc.right;
    m_cyClient = rc.bottom;
    if (m_dwStyle & PGGS_AUTOLOAD)
        StartLoader();
    // The scroll bar is always present (disabled when not needed) so its
    // appearing and disappearing can never fight the layout's width.
    UpdateScrollBar();
    return 0;
}

void CPhotoGalleryMgr::OnDestroy()
{
    StopLoader();
    KillTimer(m_hWnd, IDT_SETTLE);
    KillTimer(m_hWnd, IDT_RETRY);
    // Let the host free per-item data, as ListView does.
    if (!IsOwner())
    {
        for (int i = 0; i < (int)m_items.size() && !m_bDead; ++i)
        {
            NMPGGITEM nm = MakeItemNotify(i);
            Notify(PGGN_DELETEITEM, &nm.hdr);
        }
    }
    FreeBackBuffer();
}

void CPhotoGalleryMgr::OnStyleChanging(STYLESTRUCT* pss)
{
    // Owner-data changes what the items are; it is fixed at creation.
    pss->styleNew = (pss->styleNew & ~PGGS_OWNERDATA) | (pss->styleOld & PGGS_OWNERDATA);
}

void CPhotoGalleryMgr::OnStyleChanged(DWORD dwOld, DWORD dwNew)
{
    m_dwStyle = dwNew;
    DWORD changed = dwOld ^ dwNew;
    if (changed & PGGS_AUTOLOAD)
    {
        if (dwNew & PGGS_AUTOLOAD)
            StartLoader();
        else
            StopLoader();
    }
    if ((changed & PGGS_MULTISELECT) && !(dwNew & PGGS_MULTISELECT) && m_nSelected > 1)
    {
        int keep = (m_iFocus >= 0 && (m_states[m_iFocus] & PGGIS_SELECTED)) ? m_iFocus : OnGetNextItem(-1, PGGNI_SELECTED);
        SelectOnly(keep);
    }
    if (changed & PGGS_FIT)
        m_cache.DropAllDisplays();
    if (!(dwNew & (PGGS_HOTTRACK | PGGS_CHECKBOXES)))
        m_iHot = -1;
    InvalidateRect(m_hWnd, nullptr, FALSE);
}

void CPhotoGalleryMgr::OnDpiChanged()
{
    int dpi = (int)DpiOf(m_hWnd);
    if (dpi == m_dpi)
        return;
    Anchor a = CaptureAnchor(nullptr);
    m_dpi = dpi;
    m_bLayoutDirty = true;
    EnsureLayout();
    RestoreAnchor(a);
    m_cache.DropAllDisplays();
}

// ---- Basics -----------------------------------------------------------------------

LRESULT CPhotoGalleryMgr::Notify(UINT code, NMHDR* pHdr)
{
    if (m_bDead)
        return 0;
    pHdr->hwndFrom = m_hWnd;
    pHdr->idFrom = (UINT_PTR)GetDlgCtrlID(m_hWnd);
    pHdr->code = code;
    HWND hParent = GetParent(m_hWnd);
    if (!hParent)
        return 0;
    return SendMessage(hParent, WM_NOTIFY, pHdr->idFrom, (LPARAM)pHdr);
}

void CPhotoGalleryMgr::ItemsChanged()
{
    ++m_itemsGen;
    m_frameInfo.clear();
    m_visibleKeys.clear();
    m_hintFrom = m_hintTo = -1;
    m_bLayoutDirty = true;
    if (m_iHot >= Count())
        m_iHot = -1;
    InvalidateRect(m_hWnd, nullptr, FALSE);
}

bool CPhotoGalleryMgr::GetItemInfo(int i, ItemInfo& info)
{
    if (i < 0 || i >= Count())
        return false;
    if (!IsOwner())
    {
        const Item& it = m_items[i];
        info.key = it.key;
        info.lParam = it.lParam;
        info.iOverlay = it.iOverlay;
        info.bCallback = it.bThumbCallback;
        info.path = it.path;
        return true;
    }

    if (m_frameInfoFrame != m_frame)
    {
        m_frameInfo.clear();
        m_frameInfoFrame = m_frame;
    }
    auto found = m_frameInfo.find(i);
    if (found != m_frameInfo.end())
    {
        info = found->second;
        return true;
    }

    NMPGGDISPINFO di{};
    di.item.cbSize = sizeof(PGGITEM);
    di.item.mask = PGGIF_KEY | PGGIF_PARAM | PGGIF_OVERLAY;
    di.item.iItem = i;
    di.item.key = kIndexKeyBit | (ULONGLONG)i;
    di.item.iOverlay = -1;
    UINT gen = m_itemsGen;
    Notify(PGGN_GETDISPINFO, &di.hdr);
    if (Changed(gen))
        return false;
    info.key = di.item.key;
    info.lParam = di.item.lParam;
    info.iOverlay = di.item.iOverlay;
    info.bCallback = true;
    m_frameInfo[i] = info;
    return true;
}

UINT CPhotoGalleryMgr::StateOf(int i) const
{
    UINT s = 0;
    if (i >= 0 && i < (int)m_states.size() && (m_states[i] & PGGIS_SELECTED))
        s |= PGGIS_SELECTED;
    if (i >= 0 && i == m_iFocus)
        s |= PGGIS_FOCUSED;
    return s;
}

NMPGGITEM CPhotoGalleryMgr::MakeItemNotify(int i, const POINT* ptClient, UINT uHit, UINT uKeyFlags)
{
    NMPGGITEM nm{};
    nm.iItem = i;
    nm.iPrevItem = -1;
    nm.uHitFlags = uHit;
    nm.uKeyFlags = uKeyFlags;
    if (i >= 0 && i < Count())
    {
        if (!IsOwner())
        {
            nm.key = m_items[i].key;
            nm.lParam = m_items[i].lParam;
        }
        else
        {
            // No round trip to the host here: use what the last paint learned.
            auto found = m_frameInfo.find(i);
            if (found != m_frameInfo.end())
            {
                nm.key = found->second.key;
                nm.lParam = found->second.lParam;
            }
        }
        nm.uOldState = nm.uNewState = StateOf(i);
    }
    if (ptClient)
    {
        nm.ptAction = *ptClient;
        nm.ptScreen = *ptClient;
        ClientToScreen(m_hWnd, &nm.ptScreen);
    }
    else
    {
        GetCursorPos(&nm.ptScreen);
        nm.ptAction = nm.ptScreen;
        ScreenToClient(m_hWnd, &nm.ptAction);
    }
    return nm;
}

int CPhotoGalleryMgr::KeyToIndex(ULONGLONG key)
{
    auto vis = m_visibleKeys.find(key);
    if (vis != m_visibleKeys.end() && vis->second < Count())
        return vis->second;
    if (IsOwner())
        return -1;
    if (m_keyIndexGen != m_itemsGen)
    {
        m_keyIndex.clear();
        for (int i = (int)m_items.size() - 1; i >= 0; --i)
            m_keyIndex[m_items[i].key] = i;
        m_keyIndexGen = m_itemsGen;
    }
    auto found = m_keyIndex.find(key);
    return found == m_keyIndex.end() ? -1 : found->second;
}

UINT CPhotoGalleryMgr::KeyFlagsFromMouse(WPARAM keys)
{
    UINT f = 0;
    if (keys & MK_CONTROL)
        f |= PGGKF_CONTROL;
    if (keys & MK_SHIFT)
        f |= PGGKF_SHIFT;
    if (GetKeyState(VK_MENU) < 0)
        f |= PGGKF_ALT;
    return f;
}

UINT CPhotoGalleryMgr::KeyFlagsFromKeyboard()
{
    UINT f = 0;
    if (GetKeyState(VK_CONTROL) < 0)
        f |= PGGKF_CONTROL;
    if (GetKeyState(VK_SHIFT) < 0)
        f |= PGGKF_SHIFT;
    if (GetKeyState(VK_MENU) < 0)
        f |= PGGKF_ALT;
    return f;
}

// ---- Colors -----------------------------------------------------------------------

COLORREF CPhotoGalleryMgr::ColorBk() const
{
    return m_colors.clrBk == CLR_DEFAULT ? GetSysColor(COLOR_WINDOW) : m_colors.clrBk;
}

COLORREF CPhotoGalleryMgr::ColorPlaceholder() const
{
    if (m_colors.clrPlaceholder != CLR_DEFAULT)
        return m_colors.clrPlaceholder;
    COLORREF bk = ColorBk();
    int lum = (GetRValue(bk) * 299 + GetGValue(bk) * 587 + GetBValue(bk) * 114) / 1000;
    return lum > 128 ? Blend(bk, RGB(0, 0, 0), 7) : Blend(bk, RGB(255, 255, 255), 12);
}

COLORREF CPhotoGalleryMgr::ColorSelection() const
{
    return m_colors.clrSelection == CLR_DEFAULT ? GetSysColor(COLOR_HIGHLIGHT) : m_colors.clrSelection;
}

COLORREF CPhotoGalleryMgr::ColorHot() const
{
    return m_colors.clrHot == CLR_DEFAULT ? RGB(255, 255, 255) : m_colors.clrHot;
}

// ---- Geometry -----------------------------------------------------------------------

void CPhotoGalleryMgr::ResolveGeometry(PGGeometry& geo) const
{
    const PGGTILEMETRICS& m = m_metrics;
    geo.cxGap = Px(m.cxGap);
    geo.cyGap = Px(m.cyGap);
    geo.cxMargin = Px(m.cxMargin);
    geo.cyMargin = Px(m.cyMargin);
    const int avail = (std::max)(1, m_cxClient - 2 * geo.cxMargin);
    const int minPx = (std::max)(8, Px(m.cxMin));
    const int maxPx = (std::max)(minPx, Px(m.cxMax));
    auto cellFor = [&](int n) { return (avail - (n - 1) * geo.cxGap) / n; };

    int n = 1;
    switch (m.mode)
    {
    case PGGTS_FIXEDCOLUMNS:
        n = (std::max)(1, m.nColumns);
        while (n > 1 && cellFor(n) < 8)
            --n;
        geo.stretch = true;
        break;
    case PGGTS_FIXEDSIZE:
        geo.cellCx = (std::max)(8, Px(m.cxTile));
        n = (std::max)(1, (avail + geo.cxGap) / (geo.cellCx + geo.cxGap));
        geo.stretch = false;
        break;
    default:    // PGGTS_AUTOFIT
    {
        const int pref = Clamp(Px(m.cxTile), minPx, maxPx);
        n = (std::max)(1, (int)std::floor((double)(avail + geo.cxGap) / (pref + geo.cxGap) + 0.5));
        while (cellFor(n) > maxPx && cellFor(n + 1) >= 8)
            ++n;
        while (n > 1 && cellFor(n) < minPx)
            --n;
        geo.stretch = true;
        break;
    }
    }
    geo.nColumns = n;
    geo.avail = avail;
    if (geo.stretch)
        geo.cellCx = (std::max)(1, cellFor(n));
    geo.cellCy = (std::max)(1, MulDiv(geo.cellCx, m.nAspectY, m.nAspectX));
}

void CPhotoGalleryMgr::EnsureLayout()
{
    if (!m_bLayoutDirty || m_bInLayout)
        return;
    m_bInLayout = true;
    m_bLayoutDirty = false;
    ResolveGeometry(m_geo);
    if (m_layout->NeedsSpans())
    {
        if (IsOwner() && !EnsureOwnerSpans())
        {
            // The host changed the items while we asked; lay out next time.
            m_bLayoutDirty = true;
            m_bInLayout = false;
            return;
        }
        m_layout->Arrange(m_geo, Count(), [this](int i) -> PGSpan {
            if (IsOwner())
                return i < (int)m_ownerSpans.size() ? m_ownerSpans[i] : PGSpan{};
            return m_items[i].span;
        });
    }
    else
    {
        m_layout->Arrange(m_geo, Count(), nullptr);
    }
    m_bInLayout = false;
    ClampScroll();
    UpdateScrollBar();
}

bool CPhotoGalleryMgr::EnsureOwnerSpans()
{
    if (m_bOwnerSpansValid && (int)m_ownerSpans.size() == m_nOwnerCount)
        return true;
    const UINT gen = m_itemsGen;
    std::vector<PGSpan> spans(m_nOwnerCount);
    for (int i = 0; i < m_nOwnerCount; ++i)
    {
        NMPGGDISPINFO di{};
        di.item.cbSize = sizeof(PGGITEM);
        di.item.mask = PGGIF_SPAN;
        di.item.iItem = i;
        di.item.cxSpan = di.item.cySpan = 1;
        Notify(PGGN_GETDISPINFO, &di.hdr);
        if (Changed(gen))
            return false;
        spans[i].cx = (BYTE)Clamp(di.item.cxSpan, 1, PGG_MAXSPAN);
        spans[i].cy = (BYTE)Clamp(di.item.cySpan, 1, PGG_MAXSPAN);
    }
    m_ownerSpans.swap(spans);
    m_bOwnerSpansValid = true;
    return true;
}

int CPhotoGalleryMgr::MaxColumns() const
{
    int gap = Px(m_metrics.cxGap);
    int avail = m_cxClient - 2 * Px(m_metrics.cxMargin);
    int n = (avail + gap) / ((std::max)(8, Px(m_metrics.cxMin)) + gap);
    return (std::max)((std::max)(1, n), m_metrics.nColumns);
}

PGGTILEMETRICS CPhotoGalleryMgr::Normalize(const PGGTILEMETRICS& in) const
{
    PGGTILEMETRICS m = in;
    m.cbSize = sizeof(PGGTILEMETRICS);
    m.mask = PGGTM_ALL;
    if (m.mode > PGGTS_AUTOFIT)
        m.mode = PGGTS_AUTOFIT;
    m.cxMin = Clamp(m.cxMin, 8, 4096);
    m.cxMax = Clamp(m.cxMax, m.cxMin, 8192);
    m.cxTile = Clamp(m.cxTile, m.cxMin, m.cxMax);
    m.nColumns = Clamp(m.nColumns, 1, 256);
    m.nAspectX = Clamp(m.nAspectX, 1, 1000);
    m.nAspectY = Clamp(m.nAspectY, 1, 1000);
    m.cxGap = Clamp(m.cxGap, 0, 256);
    m.cyGap = Clamp(m.cyGap, 0, 256);
    m.cxMargin = Clamp(m.cxMargin, 0, 1024);
    m.cyMargin = Clamp(m.cyMargin, 0, 1024);
    return m;
}

void CPhotoGalleryMgr::FillActual(PGGTILEMETRICS& m) const
{
    m.cxActual = m_geo.cellCx;
    m.cyActual = m_geo.cellCy;
    m.nColumnsActual = m_geo.nColumns;
}

static bool SameConfig(const PGGTILEMETRICS& a, const PGGTILEMETRICS& b)
{
    return a.mode == b.mode && a.cxTile == b.cxTile && a.cxMin == b.cxMin && a.cxMax == b.cxMax &&
           a.nColumns == b.nColumns && a.nAspectX == b.nAspectX && a.nAspectY == b.nAspectY &&
           a.cxGap == b.cxGap && a.cyGap == b.cyGap && a.cxMargin == b.cxMargin && a.cyMargin == b.cyMargin;
}

bool CPhotoGalleryMgr::ApplyMetrics(const PGGTILEMETRICS& proposedIn, UINT uCause, const POINT* ptAnchor)
{
    PGGTILEMETRICS proposed = Normalize(proposedIn);
    if (uCause == PGGMC_USERZOOM)
    {
        NMPGGTILEMETRICS nm{};
        nm.metrics = proposed;
        FillActual(nm.metrics);
        nm.uCause = uCause;
        LRESULT veto = Notify(PGGN_TILEMETRICSCHANGING, &nm.hdr);
        if (m_bDead)
            return false;
        if (veto)
        {
            m_dZoomTile = m_metrics.cxTile;
            return false;
        }
        proposed = Normalize(nm.metrics);
    }
    if (SameConfig(proposed, m_metrics))
        return false;

    Anchor a = CaptureAnchor(ptAnchor);
    const PGGeometry old = m_geo;
    m_metrics = proposed;
    if (uCause == PGGMC_HOST)
        m_hostMetrics = proposed;
    if (uCause != PGGMC_USERZOOM || m_metrics.mode == PGGTS_FIXEDCOLUMNS)
        m_dZoomTile = m_metrics.cxTile;
    m_bLayoutDirty = true;
    EnsureLayout();
    RestoreAnchor(a);
    if (old.cellCx != m_geo.cellCx || old.cellCy != m_geo.cellCy)
        BeginInteractive();
    UpdateHotFromCursor();
    if (m_bDead)
        return false;

    NMPGGTILEMETRICS nm{};
    nm.metrics = m_metrics;
    FillActual(nm.metrics);
    nm.uCause = uCause;
    Notify(PGGN_TILEMETRICSCHANGED, &nm.hdr);
    return true;
}

bool CPhotoGalleryMgr::ZoomSteps(int steps, const POINT* ptAnchor, UINT uCause)
{
    PGGTILEMETRICS m = m_metrics;
    if (steps == 0)
    {
        m = m_hostMetrics;
        m_dZoomTile = m.cxTile;
    }
    else if (m.mode == PGGTS_FIXEDCOLUMNS)
    {
        m.nColumns = Clamp(m.nColumns - steps, 1, MaxColumns());   // zooming in = fewer columns
    }
    else
    {
        m_dZoomTile = (std::min)((double)m.cxMax, (std::max)((double)m.cxMin, m_dZoomTile * std::pow(1.15, steps)));
        m.cxTile = (int)std::lround(m_dZoomTile);
    }
    return ApplyMetrics(m, uCause, ptAnchor);
}

bool CPhotoGalleryMgr::ZoomFactor(double factor, const POINT* ptAnchor, UINT uCause)
{
    PGGTILEMETRICS m = m_metrics;
    m_dZoomTile = (std::min)((double)m.cxMax, (std::max)((double)m.cxMin, m_dZoomTile * factor));
    m.cxTile = (int)std::lround(m_dZoomTile);
    if (m.cxTile == m_metrics.cxTile)
        return false;   // keep accumulating small steps
    return ApplyMetrics(m, uCause, ptAnchor);
}

void CPhotoGalleryMgr::BeginInteractive()
{
    m_bInteractive = true;
    SetTimer(m_hWnd, IDT_SETTLE, kSettleMs, nullptr);
}

CPhotoGalleryMgr::Anchor CPhotoGalleryMgr::CaptureAnchor(const POINT* ptClient)
{
    Anchor a;
    EnsureLayout();
    if (Count() == 0 || m_cyClient <= 0)
        return a;
    RECT rc;
    std::vector<int> band;
    if (ptClient)
    {
        POINT pc{ ptClient->x, ptClient->y + m_scrollY };
        int i = m_layout->HitTest(pc);
        if (i < 0)
        {
            m_layout->ItemsInBand(pc.y, pc.y + 1, band);
            if (!band.empty())
                i = band.front();
        }
        if (i >= 0 && m_layout->ItemRect(i, rc))
        {
            a.iItem = i;
            a.frac = (std::min)(1.0, (std::max)(0.0, (pc.y - rc.top) / (double)(std::max)(1L, rc.bottom - rc.top)));
            a.yClient = ptClient->y;
            return a;
        }
    }
    if (m_scrollY == 0)
    {
        a.iItem = -2;
        return a;
    }
    m_layout->ItemsInBand(m_scrollY, m_scrollY + m_cyClient, band);
    for (int i : band)
    {
        if (m_layout->ItemRect(i, rc) && rc.bottom > m_scrollY)
        {
            a.iItem = i;
            a.yClient = rc.top - m_scrollY;
            return a;
        }
    }
    return a;
}

void CPhotoGalleryMgr::RestoreAnchor(const Anchor& a)
{
    EnsureLayout();
    RECT rc;
    if (a.iItem == -2)
        m_scrollY = 0;
    else if (a.iItem >= 0 && m_layout->ItemRect(a.iItem, rc))
        m_scrollY = (int)std::lround(rc.top + a.frac * (rc.bottom - rc.top) - a.yClient);
    ClampScroll();
    UpdateScrollBar();
    InvalidateRect(m_hWnd, nullptr, FALSE);
}

void CPhotoGalleryMgr::ClampScroll()
{
    int maxY = (std::max)(0, m_layout->ContentHeight() - m_cyClient);
    m_scrollY = Clamp(m_scrollY, 0, maxY);
}

void CPhotoGalleryMgr::UpdateScrollBar()
{
    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask = SIF_ALL | SIF_DISABLENOSCROLL;
    si.nMin = 0;
    si.nMax = (std::max)(0, m_layout->ContentHeight() - 1);
    si.nPage = (UINT)(std::max)(0, m_cyClient);
    si.nPos = m_scrollY;
    SetScrollInfo(m_hWnd, SB_VERT, &si, TRUE);
}

void CPhotoGalleryMgr::ScrollTo(int y)
{
    EnsureLayout();
    int maxY = (std::max)(0, m_layout->ContentHeight() - m_cyClient);
    y = Clamp(y, 0, maxY);
    if (y == m_scrollY)
        return;
    int dy = m_scrollY - y;
    m_scrollY = y;
    UpdateScrollBar();
    // Move what is already on screen and paint only the strip that appears.
    if (dy < m_cyClient && -dy < m_cyClient)
        ScrollWindowEx(m_hWnd, 0, dy, nullptr, nullptr, nullptr, nullptr, SW_INVALIDATE);
    else
        InvalidateRect(m_hWnd, nullptr, FALSE);
    UpdateHotFromCursor();
}

int CPhotoGalleryMgr::LinePx() const
{
    return (std::max)(Px(16), m_geo.cellCy / 3);
}

int CPhotoGalleryMgr::PrefetchPx() const
{
    return (int)(std::min)(m_loaderSettings.nPrefetchScreens, 4u) * m_cyClient;
}

bool CPhotoGalleryMgr::GetItemClientRect(int i, RECT& rc)
{
    EnsureLayout();
    if (!m_layout->ItemRect(i, rc))
        return false;
    OffsetRect(&rc, 0, -m_scrollY);
    return true;
}

void CPhotoGalleryMgr::InvalidateItem(int i)
{
    RECT rc;
    if (i >= 0 && i < Count() && GetItemClientRect(i, rc) && rc.bottom > 0 && rc.top < m_cyClient)
        InvalidateRect(m_hWnd, &rc, FALSE);
}

void CPhotoGalleryMgr::InvalidateKey(ULONGLONG key)
{
    auto found = m_visibleKeys.find(key);
    if (found != m_visibleKeys.end())
        InvalidateItem(found->second);
}

bool CPhotoGalleryMgr::CheckRect(const RECT& rcTile, RECT& rcCheck) const
{
    if ((m_dwStyle & (PGGS_CHECKBOXES | PGGS_MULTISELECT)) != (PGGS_CHECKBOXES | PGGS_MULTISELECT))
        return false;
    const int d = Px(22), m = Px(6);
    if (rcTile.right - rcTile.left < 2 * d || rcTile.bottom - rcTile.top < 2 * d)
        return false;
    rcCheck = RECT{ rcTile.left + m, rcTile.top + m, rcTile.left + m + d, rcTile.top + m + d };
    return true;
}

RECT CPhotoGalleryMgr::OverlayRect(const RECT& rcTile) const
{
    int cx = 0, cy = 0;
    if (m_hOverlay)
        ImageList_GetIconSize(m_hOverlay, &cx, &cy);
    const int m = Px(6);
    return RECT{ rcTile.right - m - cx, rcTile.top + m, rcTile.right - m, rcTile.top + m + cy };
}

int CPhotoGalleryMgr::HitTestClient(POINT pt, UINT* pFlags)
{
    EnsureLayout();
    int i = -1;
    if (pt.x >= 0 && pt.y >= 0 && pt.x < m_cxClient && pt.y < m_cyClient)
        i = m_layout->HitTest(POINT{ pt.x, pt.y + m_scrollY });
    UINT f = PGGHT_NOWHERE;
    if (i >= 0)
    {
        f = PGGHT_ONIMAGE;
        RECT rc, rcPart;
        GetItemClientRect(i, rc);
        int iOverlay = -1;
        if (!IsOwner())
            iOverlay = m_items[i].iOverlay;
        else
        {
            auto found = m_frameInfo.find(i);
            if (found != m_frameInfo.end())
                iOverlay = found->second.iOverlay;
        }
        if (CheckRect(rc, rcPart) && PtInRect(&rcPart, pt))
            f = PGGHT_ONCHECK;
        else if (m_hOverlay && iOverlay >= 0 && (rcPart = OverlayRect(rc), PtInRect(&rcPart, pt)))
            f = PGGHT_ONOVERLAY;
    }
    if (pFlags)
        *pFlags = f;
    return i;
}

bool CPhotoGalleryMgr::EnsureVisible(int i, bool bPartialOK)
{
    EnsureLayout();
    RECT rc;
    if (!m_layout->ItemRect(i, rc))
        return false;
    const int top = m_scrollY, bottom = m_scrollY + m_cyClient;
    if (bPartialOK && rc.bottom > top && rc.top < bottom)
        return true;
    if (rc.top < top)
        ScrollTo(rc.top - m_geo.cyGap);
    else if (rc.bottom > bottom)
        ScrollTo(rc.bottom - m_cyClient + m_geo.cyGap);
    return true;
}

int CPhotoGalleryMgr::TopIndex()
{
    EnsureLayout();
    std::vector<int> band;
    m_layout->ItemsInBand(m_scrollY, m_scrollY + (std::max)(1, m_cyClient), band);
    RECT rc;
    for (int i : band)
        if (m_layout->ItemRect(i, rc) && rc.bottom > m_scrollY)
            return i;
    return -1;
}

// ---- Painting -----------------------------------------------------------------------------

void CPhotoGalleryMgr::OnPaint()
{
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(m_hWnd, &ps);
    if (hdc)
        Render(hdc, ps.rcPaint);
    EndPaint(m_hWnd, &ps);
    if (!m_bDead)
        AfterPaint();
}

bool CPhotoGalleryMgr::EnsureBackBuffer(int cx, int cy)
{
    cx = (std::max)(cx, 1);
    cy = (std::max)(cy, 1);
    if (m_hdcBack && m_cxBack >= cx && m_cyBack >= cy)
        return true;
    FreeBackBuffer();

    // Grow in steps so a live resize does not reallocate on every pixel.
    const int w = (cx + 127) & ~127, h = (cy + 127) & ~127;
    HDC hdcScreen = GetDC(m_hWnd);
    m_hdcBack = CreateCompatibleDC(hdcScreen);
    BITMAPINFO bi{};
    bi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bi.bmiHeader.biWidth = w;
    bi.bmiHeader.biHeight = -h;
    bi.bmiHeader.biPlanes = 1;
    bi.bmiHeader.biBitCount = 32;
    bi.bmiHeader.biCompression = BI_RGB;
    void* pBits = nullptr;
    m_hbmBack = CreateDIBSection(hdcScreen, &bi, DIB_RGB_COLORS, &pBits, nullptr, 0);
    ReleaseDC(m_hWnd, hdcScreen);
    if (!m_hdcBack || !m_hbmBack)
    {
        FreeBackBuffer();
        return false;
    }
    m_hbmBackOld = (HBITMAP)SelectObject(m_hdcBack, m_hbmBack);
    m_cxBack = w;
    m_cyBack = h;
    // Cached tiles were made for the old surface.
    m_cache.DropAllDisplays();
    return true;
}

void CPhotoGalleryMgr::FreeBackBuffer()
{
    if (m_hdcBack)
    {
        if (m_hbmBackOld)
            SelectObject(m_hdcBack, m_hbmBackOld);
        DeleteDC(m_hdcBack);
    }
    if (m_hbmBack)
        DeleteObject(m_hbmBack);
    m_hdcBack = nullptr;
    m_hbmBack = m_hbmBackOld = nullptr;
    m_cxBack = m_cyBack = 0;
}

void CPhotoGalleryMgr::Render(HDC hdcTarget, const RECT& rcPaintIn)
{
    EnsureLayout();
    if (m_bDead)
        return;
    RECT rcClient{ 0, 0, m_cxClient, m_cyClient }, rcPaint;
    if (!IntersectRect(&rcPaint, &rcPaintIn, &rcClient))
        return;
    if (!EnsureBackBuffer(m_cxClient, m_cyClient))
    {
        HBRUSH hbr = CreateSolidBrush(ColorBk());
        FillRect(hdcTarget, &rcPaint, hbr);
        DeleteObject(hbr);
        return;
    }
    ++m_frame;
    SendCacheHint();
    if (m_bDead)
        return;

    std::vector<int> items;
    m_layout->ItemsInBand(m_scrollY + rcPaint.top, m_scrollY + rcPaint.bottom, items);

    const int w = rcPaint.right - rcPaint.left, h = rcPaint.bottom - rcPaint.top;
    bool bAborted = false;
    {
        Gdiplus::Graphics g(m_hdcBack);
        g.SetClip(Gdiplus::Rect(rcPaint.left, rcPaint.top, w, h));
        Gdiplus::SolidBrush bk(GpColor(ColorBk()));
        g.FillRectangle(&bk, rcPaint.left, rcPaint.top, w, h);

        DWORD dwPre = CustomDraw(g, CDDS_PREPAINT, -1, rcClient, nullptr, 0, FALSE);
        if (m_bDead)
            return;
        if (!(dwPre & CDRF_SKIPDEFAULT))
        {
            for (int i : items)
            {
                if (!DrawTile(g, i, dwPre))
                {
                    bAborted = true;
                    break;
                }
            }
        }
        if (m_bDead)
            return;
        if (!bAborted && (dwPre & CDRF_NOTIFYPOSTPAINT))
            CustomDraw(g, CDDS_POSTPAINT, -1, rcClient, nullptr, 0, FALSE);
        if (m_bDead)
            return;
    }
    BitBlt(hdcTarget, rcPaint.left, rcPaint.top, w, h, m_hdcBack, rcPaint.left, rcPaint.top, SRCCOPY);
    if (bAborted)
        InvalidateRect(m_hWnd, nullptr, FALSE);     // the items changed mid-paint: paint again
}

bool CPhotoGalleryMgr::DrawTile(Gdiplus::Graphics& g, int i, DWORD dwPrePaint)
{
    RECT rc;
    if (!GetItemClientRect(i, rc))
        return true;
    const UINT gen = m_itemsGen;
    ItemInfo info;
    if (!GetItemInfo(i, info))
        return false;

    const int w = rc.right - rc.left, h = rc.bottom - rc.top;
    const bool bHasFocus = GetFocus() == m_hWnd;
    const bool bSel = (m_states[i] & PGGIS_SELECTED) != 0;
    const bool bFocus = i == m_iFocus && bHasFocus;
    const bool bHot = i == m_iHot && (m_dwStyle & (PGGS_HOTTRACK | PGGS_CHECKBOXES));
    const UINT uItemState = (bSel ? CDIS_SELECTED : 0) | (bFocus ? CDIS_FOCUS : 0) | (bHot ? CDIS_HOT : 0);

    PGThumb* pFound = m_cache.Find(info.key);
    DWORD dwItem = CDRF_DODEFAULT;
    if (dwPrePaint & CDRF_NOTIFYITEMDRAW)
    {
        dwItem = CustomDraw(g, CDDS_ITEMPREPAINT, i, rc, &info, uItemState, pFound && !pFound->src.empty());
        if (Changed(gen))
            return false;
    }

    BOOL fThumbReady = FALSE;
    if (!(dwItem & CDRF_SKIPDEFAULT))
    {
        PGThumb* t = EnsureThumb(i, info, w, h);
        if (Changed(gen))
            return false;
        bool bDrawn = false;
        if (t)
        {
            t->lastUsed = m_frame;
            if (!t->src.empty() && EnsureDisplay(*t, w, h, g))
            {
                g.DrawCachedBitmap(t->disp.get(), rc.left, rc.top);
                bDrawn = true;
                fThumbReady = TRUE;
            }
        }
        if (!bDrawn)
        {
            Gdiplus::SolidBrush ph(GpColor(ColorPlaceholder()));
            g.FillRectangle(&ph, rc.left, rc.top, w, h);
            if (t && t->state == PGThumb::Failed)
                DrawFailedGlyph(g, rc, t->uFailReason);
        }
        DrawChrome(g, rc, bSel, bFocus, bHot);
        if (m_hOverlay && info.iOverlay >= 0)
        {
            RECT ro = OverlayRect(rc);
            HDC hdc = g.GetHDC();
            ImageList_Draw(m_hOverlay, info.iOverlay, hdc, ro.left, ro.top, ILD_TRANSPARENT);
            g.ReleaseHDC(hdc);
        }
    }
    if (dwItem & CDRF_NOTIFYPOSTPAINT)
    {
        CustomDraw(g, CDDS_ITEMPOSTPAINT, i, rc, &info, uItemState, fThumbReady);
        if (Changed(gen))
            return false;
    }
    return true;
}

void CPhotoGalleryMgr::DrawChrome(Gdiplus::Graphics& g, const RECT& rc, bool bSel, bool bFocus, bool bHot)
{
    const Gdiplus::RectF r((Gdiplus::REAL)rc.left, (Gdiplus::REAL)rc.top,
                           (Gdiplus::REAL)(rc.right - rc.left), (Gdiplus::REAL)(rc.bottom - rc.top));
    const bool bActive = GetFocus() == m_hWnd || (m_dwStyle & PGGS_SHOWSELALWAYS);

    if (bHot && !bSel)
    {
        Gdiplus::SolidBrush hot(GpColor(ColorHot(), 48));
        g.FillRectangle(&hot, r);
    }
    if (bSel)
    {
        // Without focus the selection goes grey, as in Explorer.
        COLORREF c = bActive ? ColorSelection() : GetSysColor(COLOR_BTNSHADOW);
        Gdiplus::SolidBrush tint(GpColor(c, 64));
        g.FillRectangle(&tint, r);
        Gdiplus::Pen border(GpColor(c), (Gdiplus::REAL)Px(3));
        border.SetAlignment(Gdiplus::PenAlignmentInset);
        g.DrawRectangle(&border, r);
    }

    RECT rcCheck;
    if (CheckRect(rc, rcCheck) && (bSel || bHot))
    {
        g.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
        const Gdiplus::RectF e((Gdiplus::REAL)rcCheck.left, (Gdiplus::REAL)rcCheck.top,
                               (Gdiplus::REAL)(rcCheck.right - rcCheck.left), (Gdiplus::REAL)(rcCheck.bottom - rcCheck.top));
        Gdiplus::Pen ring(Gdiplus::Color(235, 255, 255, 255), (Gdiplus::REAL)Px(2));
        if (bSel)
        {
            Gdiplus::SolidBrush fill(GpColor(ColorSelection()));
            g.FillEllipse(&fill, e);
            g.DrawEllipse(&ring, e);
            Gdiplus::Pen tick(Gdiplus::Color(255, 255, 255, 255), (Gdiplus::REAL)Px(2));
            tick.SetLineJoin(Gdiplus::LineJoinRound);
            tick.SetStartCap(Gdiplus::LineCapRound);
            tick.SetEndCap(Gdiplus::LineCapRound);
            Gdiplus::PointF pts[3] = {
                { e.X + e.Width * 0.28f, e.Y + e.Height * 0.52f },
                { e.X + e.Width * 0.44f, e.Y + e.Height * 0.68f },
                { e.X + e.Width * 0.74f, e.Y + e.Height * 0.36f } };
            g.DrawLines(&tick, pts, 3);
        }
        else
        {
            Gdiplus::SolidBrush shade(Gdiplus::Color(70, 0, 0, 0));
            g.FillEllipse(&shade, e);
            g.DrawEllipse(&ring, e);
        }
        g.SetSmoothingMode(Gdiplus::SmoothingModeNone);
    }

    if (bFocus)
    {
        RECT rf = rc;
        InflateRect(&rf, -Px(2), -Px(2));
        if (m_colors.clrFocus != CLR_DEFAULT)
        {
            Gdiplus::Pen dot(GpColor(m_colors.clrFocus), 1.0f);
            dot.SetDashStyle(Gdiplus::DashStyleDot);
            g.DrawRectangle(&dot, (INT)rf.left, (INT)rf.top, (INT)(rf.right - rf.left - 1), (INT)(rf.bottom - rf.top - 1));
        }
        else
        {
            // The classic XOR focus rectangle shows on any photo.
            HDC hdc = g.GetHDC();
            DrawFocusRect(hdc, &rf);
            g.ReleaseHDC(hdc);
        }
    }
}

void CPhotoGalleryMgr::DrawFailedGlyph(Gdiplus::Graphics& g, const RECT& rc, UINT uReason)
{
    const int w = rc.right - rc.left, h = rc.bottom - rc.top;
    const float s = (float)(std::min)(w, h) / 5.0f;
    if (s < Px(6))
        return;
    const float cx = rc.left + w / 2.0f, cy = rc.top + h / 2.0f;
    const COLORREF c = Blend(ColorPlaceholder(), GetSysColor(COLOR_WINDOWTEXT), 30);
    g.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
    if (uReason == PGGTR_CLOUDONLY)
    {
        Gdiplus::SolidBrush b(GpColor(c));
        g.FillEllipse(&b, cx - s * 1.0f, cy - s * 0.15f, s * 0.9f, s * 0.9f);
        g.FillEllipse(&b, cx - s * 0.5f, cy - s * 0.65f, s * 1.1f, s * 1.1f);
        g.FillEllipse(&b, cx + s * 0.15f, cy - s * 0.25f, s * 0.9f, s * 0.9f);
        g.FillRectangle(&b, cx - s * 0.55f, cy + s * 0.25f, s * 1.15f, s * 0.5f);
    }
    else
    {
        Gdiplus::Pen pen(GpColor(c), (Gdiplus::REAL)(std::max)(1, Px(2)));
        pen.SetLineJoin(Gdiplus::LineJoinRound);
        g.DrawRectangle(&pen, cx - s, cy - s * 0.75f, 2 * s, 1.5f * s);
        Gdiplus::PointF hill[3] = { { cx - s * 0.75f, cy + s * 0.5f }, { cx - s * 0.15f, cy - s * 0.2f }, { cx + s * 0.4f, cy + s * 0.5f } };
        g.DrawLines(&pen, hill, 3);
        g.DrawEllipse(&pen, cx + s * 0.35f, cy - s * 0.5f, s * 0.35f, s * 0.35f);
    }
    g.SetSmoothingMode(Gdiplus::SmoothingModeNone);
}

DWORD CPhotoGalleryMgr::CustomDraw(Gdiplus::Graphics& g, DWORD dwStage, int i, const RECT& rc,
                                   const ItemInfo* pInfo, UINT uItemState, BOOL fThumbReady)
{
    NMPGGCUSTOMDRAW cd{};
    cd.nmcd.dwDrawStage = dwStage;
    cd.nmcd.rc = rc;
    cd.nmcd.dwItemSpec = i >= 0 ? (DWORD_PTR)i : 0;
    cd.nmcd.uItemState = uItemState;
    cd.nmcd.lItemlParam = pInfo ? pInfo->lParam : 0;
    cd.key = pInfo ? pInfo->key : 0;
    cd.iOverlay = pInfo ? pInfo->iOverlay : -1;
    cd.fThumbReady = fThumbReady;
    // GDI+ must hand the DC over so the host's GDI calls land in order.
    HDC hdc = g.GetHDC();
    cd.nmcd.hdc = hdc;
    LRESULT r = Notify(NM_CUSTOMDRAW, &cd.nmcd.hdr);
    g.ReleaseHDC(hdc);
    return (DWORD)r;
}

void CPhotoGalleryMgr::SendCacheHint()
{
    if (!IsOwner() || Count() == 0)
        return;
    const int pre = PrefetchPx();
    std::vector<int> band;
    m_layout->ItemsInBand((std::max)(0, m_scrollY - pre), m_scrollY + m_cyClient + pre, band);
    if (band.empty())
        return;
    if (band.front() == m_hintFrom && band.back() == m_hintTo)
        return;
    m_hintFrom = band.front();
    m_hintTo = band.back();
    NMPGGCACHEHINT hint{};
    hint.iFrom = m_hintFrom;
    hint.iTo = m_hintTo;
    Notify(PGGN_ODCACHEHINT, &hint.hdr);
}

// After every paint: remember what is on screen, queue the loader with the
// visible tiles first and the prefetch ring after (nearest first), and trim
// the cache.
void CPhotoGalleryMgr::AfterPaint()
{
    m_visibleKeys.clear();
    const bool bAuto = m_loader && (m_dwStyle & PGGS_AUTOLOAD);
    if (Count() == 0)
    {
        if (bAuto)
            m_loader->SetWanted({});
        return;
    }

    const int pre = bAuto ? PrefetchPx() : 0;
    const int viewTop = m_scrollY, viewBottom = m_scrollY + m_cyClient;
    std::vector<int> band;
    m_layout->ItemsInBand((std::max)(0, viewTop - pre), viewBottom + pre, band);

    std::vector<std::pair<int, int>> order;     // (distance from the viewport, index)
    order.reserve(band.size());
    for (int i : band)
    {
        RECT rc;
        m_layout->ItemRect(i, rc);
        int dist = (rc.bottom > viewTop && rc.top < viewBottom) ? 0
                 : (rc.top >= viewBottom ? rc.top - viewBottom + 1 : viewTop - rc.bottom + 1);
        order.emplace_back(dist, i);
    }
    std::stable_sort(order.begin(), order.end(),
                     [](const std::pair<int, int>& a, const std::pair<int, int>& b) { return a.first < b.first; });

    const UINT gen = m_itemsGen;
    const bool bFit = (m_dwStyle & PGGS_FIT) != 0;
    std::vector<PGLoadJob> jobs;
    for (const auto& o : order)
    {
        const int i = o.second;
        if (o.first > 0 && !bAuto)
            break;
        ItemInfo info;
        if (!GetItemInfo(i, info))
            return;
        if (o.first == 0)
            m_visibleKeys[info.key] = i;
        if (!bAuto)
            continue;

        RECT rc;
        m_layout->ItemRect(i, rc);
        const int w = rc.right - rc.left, h = rc.bottom - rc.top;
        PGThumb* t = EnsureThumb(i, info, w, h);
        if (Changed(gen))
            return;
        if (!t || t->path.empty() || t->state == PGThumb::Pending || t->state == PGThumb::Failed)
            continue;
        if (t->retryTick && (int)(GetTickCount() - t->retryTick) < 0)
            continue;       // the cloud provider is still fetching it
        const bool bNeed = t->src.empty() || (!m_bInteractive && (t->bProvisional || NeedsUpgrade(*t, w, h)));
        if (!bNeed)
            continue;
        PGLoadJob job;
        job.key = info.key;
        job.epoch = t->epoch;
        job.path = t->path;
        job.cx = w;
        job.cy = h;
        job.bFit = bFit;
        jobs.push_back(std::move(job));
    }
    if (bAuto)
        m_loader->SetWanted(std::move(jobs));
    m_cache.Trim(m_frame);
}

// ---- Thumbnails -----------------------------------------------------------------------------

bool CPhotoGalleryMgr::NeedsUpgrade(const PGThumb& t, int w, int h) const
{
    if (t.src.empty() || t.bNative)
        return false;
    const bool bFit = (m_dwStyle & PGGS_FIT) != 0;
    return ScaleNeeded(t.src.w, t.src.h, w, h, bFit) > 1.25 &&
           (std::max)(w, h) > t.requestedEdge + t.requestedEdge / 4;
}

// Makes sure the cache has an entry for this item and, if it has no pixels
// (or too few for the tile), asks the host. Under PGGS_AUTOLOAD the host is
// asked once and the loader does the rest (AfterPaint queues it).
// Returns nullptr if the notification changed the items or killed the control.
PGThumb* CPhotoGalleryMgr::EnsureThumb(int i, const ItemInfo& info, int w, int h)
{
    const bool bAuto = (m_dwStyle & PGGS_AUTOLOAD) != 0;
    PGThumb* t = &m_cache.Get(info.key);
    if (bAuto && !IsOwner() && t->path != info.path)
        t->path = info.path;
    if (t->state == PGThumb::Pending || t->state == PGThumb::Failed)
        return t;

    const bool bHave = !t->src.empty();
    if (bHave)
    {
        if (t->bProvisional && bAuto)
            return t;               // the loader fetches the sharper one
        if (!NeedsUpgrade(*t, w, h) || m_bInteractive)
            return t;               // good enough, or sharpen once the zoom settles
    }

    const bool bAskHost = IsOwner() || info.bCallback;
    if (!bAskHost || (bAuto && t->bHostAsked))
        return t;

    t->bHostAsked = true;
    NMPGGDISPINFO di{};
    di.item.cbSize = sizeof(PGGITEM);
    di.item.mask = PGGIF_THUMB;
    di.item.iItem = i;
    di.item.key = info.key;
    di.item.iOverlay = info.iOverlay;
    di.sizeWanted = SIZE{ w, h };
    if (bAuto && IsOwner())
    {
        di.item.mask |= PGGIF_PATH;
        m_pathBuf[0] = L'\0';
        di.item.pszPath = &m_pathBuf[0];
        di.item.cchPathMax = (int)m_pathBuf.size();
    }
    const UINT gen = m_itemsGen;
    Notify(PGGN_GETDISPINFO, &di.hdr);
    if (Changed(gen))
        return nullptr;

    // The handler may have added cache entries (PGGM_SETTHUMBNAIL), so look again.
    t = m_cache.Find(info.key);
    if (!t)
        return nullptr;
    const int edge = (std::max)(w, h);
    if (di.item.hbmThumb && di.item.hbmThumb != PGG_THUMBCALLBACK)
    {
        PGPixels px;
        if (PixelsFromHBITMAP(di.item.hbmThumb, px))
        {
            CapHostPixels(px);
            StoreSource(*t, std::move(px), edge);
            t->bProvisional = (di.item.thumbFlags & PGGTF_PROVISIONAL) != 0;
        }
    }
    else if (di.item.thumbFlags & PGGTF_PENDING)
    {
        t->state = PGThumb::Pending;
        t->requestedEdge = edge;
    }
    if ((di.item.mask & PGGIF_PATH) && di.item.pszPath && di.item.pszPath[0])
        t->path = di.item.pszPath;
    if (!bAuto && t->src.empty() && t->state != PGThumb::Pending)
    {
        t->state = PGThumb::Failed;     // the host had nothing; PGGM_SETTHUMBNAIL or an invalidate revives it
        t->uFailReason = 0;
    }
    if (!bAuto && bHave)
        t->requestedEdge = (std::max)(t->requestedEdge, edge);  // asked at this size; do not ask again
    return t;
}

bool CPhotoGalleryMgr::EnsureDisplay(PGThumb& t, int w, int h, Gdiplus::Graphics& target)
{
    const bool bFit = (m_dwStyle & PGGS_FIT) != 0;
    if (t.disp && t.dispW == w && t.dispH == h && t.dispFit == bFit && !(t.dispLowQ && !m_bInteractive))
        return true;
    t.DropDisplay();

    const bool bHQ = !m_bInteractive;
    Gdiplus::Bitmap src(t.src.w, t.src.h, t.src.w * 4, PixelFormat32bppPARGB, (BYTE*)t.src.px.data());
    Gdiplus::Bitmap tile(w, h, PixelFormat32bppPARGB);
    {
        Gdiplus::Graphics tg(&tile);
        tg.Clear(GpColor(ColorPlaceholder()));
        tg.SetInterpolationMode(bHQ ? Gdiplus::InterpolationModeHighQualityBicubic : Gdiplus::InterpolationModeBilinear);
        tg.SetPixelOffsetMode(Gdiplus::PixelOffsetModeHalf);
        Gdiplus::ImageAttributes ia;
        ia.SetWrapMode(Gdiplus::WrapModeTileFlipXY);
        const double sw = t.src.w, sh = t.src.h;
        if (!bFit)
        {
            // Center-crop: cover the tile, cut the overhang evenly.
            double s = ScaleNeeded(t.src.w, t.src.h, w, h, false);
            double cw = w / s, ch = h / s;
            tg.DrawImage(&src, Gdiplus::RectF(0, 0, (Gdiplus::REAL)w, (Gdiplus::REAL)h),
                         (Gdiplus::REAL)((sw - cw) / 2), (Gdiplus::REAL)((sh - ch) / 2),
                         (Gdiplus::REAL)cw, (Gdiplus::REAL)ch, Gdiplus::UnitPixel, &ia);
        }
        else
        {
            double s = ScaleNeeded(t.src.w, t.src.h, w, h, true);
            double dw = sw * s, dh = sh * s;
            tg.DrawImage(&src, Gdiplus::RectF((Gdiplus::REAL)((w - dw) / 2), (Gdiplus::REAL)((h - dh) / 2),
                                              (Gdiplus::REAL)dw, (Gdiplus::REAL)dh),
                         0, 0, (Gdiplus::REAL)sw, (Gdiplus::REAL)sh, Gdiplus::UnitPixel, &ia);
        }
    }
    t.disp.reset(new Gdiplus::CachedBitmap(&tile, &target));
    if (!t.disp || t.disp->GetLastStatus() != Gdiplus::Ok)
    {
        t.disp.reset();
        return false;
    }
    t.dispW = w;
    t.dispH = h;
    t.dispFit = bFit;
    t.dispLowQ = !bHQ;
    return true;
}

// Host bitmaps can be any size; keep no more than the largest tile can use.
void CPhotoGalleryMgr::CapHostPixels(PGPixels& px) const
{
    const int limit = 2 * (std::max)(Px(m_metrics.cxMax), TileEdge());
    const int shortEdge = (std::min)(px.w, px.h);
    if (shortEdge <= limit + limit / 2)
        return;
    const double s = (double)limit / shortEdge;
    PGPixels scaled;
    if (ScalePixels(px, (std::max)(1, (int)(px.w * s + 0.5)), (std::max)(1, (int)(px.h * s + 0.5)), scaled))
        px = std::move(scaled);
}

void CPhotoGalleryMgr::StoreSource(PGThumb& t, PGPixels&& px, int edge)
{
    t.src = std::move(px);
    t.state = PGThumb::Ready;
    t.uFailReason = 0;
    t.bNative = false;
    t.bProvisional = false;
    t.requestedEdge = edge;
    t.DropDisplay();
}

bool CPhotoGalleryMgr::SetItemThumb(Item& it, HBITMAP hbm, UINT flags)
{
    if (hbm == PGG_THUMBCALLBACK)
    {
        Unpin(it);
        it.bThumbCallback = true;
        PGThumb& t = m_cache.Get(it.key);
        if (t.nPins == 0)
            t.ResetSource();
        return true;
    }
    it.bThumbCallback = false;
    if (!hbm)
    {
        Unpin(it);
        PGThumb* t = m_cache.Find(it.key);
        if (t && t->nPins == 0)
            t->ResetSource();
        return true;
    }
    PGPixels px;
    if (!PixelsFromHBITMAP(hbm, px))
        return false;
    CapHostPixels(px);
    PGThumb& t = m_cache.Get(it.key);
    if (!it.bPinned)
    {
        ++t.nPins;
        it.bPinned = true;
    }
    StoreSource(t, std::move(px), TileEdge());
    t.bProvisional = (flags & PGGTF_PROVISIONAL) != 0;
    return true;
}

void CPhotoGalleryMgr::Unpin(Item& it)
{
    if (!it.bPinned)
        return;
    PGThumb* t = m_cache.Find(it.key);
    if (t && t->nPins > 0)
        --t->nPins;
    it.bPinned = false;
}

void CPhotoGalleryMgr::OnLoaderDone()
{
    if (!m_loader)
        return;
    std::vector<PGLoadResult> results;
    m_loader->TakeResults(results);
    for (PGLoadResult& r : results)
    {
        PGThumb* t = m_cache.Find(r.key);
        if (!t || t->epoch != r.epoch)
            continue;   // evicted or invalidated while loading
        if (!r.bOk && r.uReason == PGG_REASON_PENDING && t->nRetries < kMaxRetries)
        {
            // The provider is fetching its preview. Quietly ask again soon.
            ++t->nRetries;
            t->retryTick = GetTickCount() + kRetryMs;
            SetTimer(m_hWnd, IDT_RETRY, kRetryMs, nullptr);
            continue;
        }
        if (r.uReason == PGG_REASON_PENDING)
            r.uReason = PGGTR_CLOUDONLY;    // gave up waiting
        t->retryTick = 0;

        NMPGGTHUMB nm{};
        nm.iItem = KeyToIndex(r.key);
        nm.key = r.key;
        nm.pszPath = r.path.c_str();
        nm.uSource = r.uSource;
        nm.hr = r.hr;
        if (r.bOk)
        {
            t->src = std::move(r.pixels);
            t->state = PGThumb::Ready;
            t->bNative = r.bNative;
            t->bProvisional = false;
            t->uFailReason = 0;
            t->requestedEdge = (std::max)(r.cx, r.cy);
            t->DropDisplay();

            nm.bmih.biSize = sizeof(BITMAPINFOHEADER);
            nm.bmih.biWidth = t->src.w;
            nm.bmih.biHeight = -t->src.h;
            nm.bmih.biPlanes = 1;
            nm.bmih.biBitCount = 32;
            nm.bmih.biCompression = BI_RGB;
            nm.bmih.biSizeImage = (DWORD)t->src.bytes();
            nm.pvBits = t->src.px.data();
            nm.cbStride = t->src.w * 4;
            Notify(PGGN_THUMBLOADED, &nm.hdr);
        }
        else
        {
            if (t->src.empty())
                t->state = PGThumb::Failed;
            t->bProvisional = false;
            t->requestedEdge = (std::max)(r.cx, r.cy);  // do not retry at this size
            t->uFailReason = r.uReason;
            nm.uReason = r.uReason;
            Notify(PGGN_THUMBFAILED, &nm.hdr);
        }
        if (m_bDead)
            return;
        InvalidateKey(r.key);
    }
}

void CPhotoGalleryMgr::StartLoader()
{
    if (!m_loader)
        m_loader = std::make_unique<CThumbLoader>(m_hWnd, m_loaderSettings);
    InvalidateRect(m_hWnd, nullptr, FALSE);
}

void CPhotoGalleryMgr::StopLoader()
{
    if (m_loader)
    {
        m_loader->Detach();
        m_loader.reset();
    }
}

// ---- Selection --------------------------------------------------------------------------------

void CPhotoGalleryMgr::ApplyState(int i, UINT state)
{
    const bool bWas = (m_states[i] & PGGIS_SELECTED) != 0, bNow = (state & PGGIS_SELECTED) != 0;
    if (bWas != bNow)
    {
        m_states[i] = bNow ? PGGIS_SELECTED : 0;
        m_nSelected += bNow ? 1 : -1;
    }
    if (state & PGGIS_FOCUSED)
        m_iFocus = i;
    else if (m_iFocus == i)
        m_iFocus = -1;
}

// Changes one item's state with PGGN_ITEMCHANGING / PGGN_ITEMCHANGED. A veto
// is not an error. Returns false only when a handler changed the items or
// destroyed the control, so the caller must stop.
bool CPhotoGalleryMgr::SetItemStateNotify(int i, UINT state, UINT mask)
{
    if (i < 0 || i >= Count())
        return true;
    mask &= (PGGIS_SELECTED | PGGIS_FOCUSED);
    const UINT gen = m_itemsGen;

    // Single selection: the previous item lets go first.
    if (!IsMulti() && (mask & state & PGGIS_SELECTED))
    {
        for (int j = OnGetNextItem(-1, PGGNI_SELECTED); j >= 0; j = OnGetNextItem(j, PGGNI_SELECTED))
        {
            if (j != i && !SetItemStateNotify(j, 0, PGGIS_SELECTED))
                return false;
        }
    }
    if ((mask & state & PGGIS_FOCUSED) && m_iFocus >= 0 && m_iFocus != i)
    {
        if (!SetItemStateNotify(m_iFocus, 0, PGGIS_FOCUSED))
            return false;
    }
    if (Changed(gen))
        return false;

    const UINT oldS = StateOf(i), newS = (oldS & ~mask) | (state & mask);
    if (oldS == newS)
        return true;
    NMPGGITEM nm = MakeItemNotify(i);
    nm.uOldState = oldS;
    nm.uNewState = newS;
    nm.uChanged = oldS ^ newS;
    LRESULT veto = Notify(PGGN_ITEMCHANGING, &nm.hdr);
    if (Changed(gen))
        return false;
    if (veto)
        return true;
    ApplyState(i, newS);
    InvalidateItem(i);
    Notify(PGGN_ITEMCHANGED, &nm.hdr);
    return !Changed(gen);
}

// Selects [a, b] (a < 0: nothing) and, unless bKeepOthers, deselects the
// rest. Up to kBulkLimit changes are reported item by item; more come as a
// single PGGN_ITEMCHANGING / PGGN_ITEMCHANGED with iItem -1.
bool CPhotoGalleryMgr::SelectRange(int a, int b, bool bKeepOthers)
{
    const int n = Count();
    if (a > b)
        std::swap(a, b);
    if (a >= 0 && !IsMulti())
        a = b;      // single selection: the item under the cursor wins
    if (b >= n)
        b = n - 1;

    std::vector<int> toClear, toSet;
    for (int i = 0; i < n; ++i)
    {
        const bool bIn = a >= 0 && i >= a && i <= b;
        const bool bSel = (m_states[i] & PGGIS_SELECTED) != 0;
        if (bIn && !bSel)
            toSet.push_back(i);
        else if (!bIn && bSel && !bKeepOthers)
            toClear.push_back(i);
    }
    if (toSet.empty() && toClear.empty())
        return true;

    const UINT gen = m_itemsGen;
    if ((int)(toSet.size() + toClear.size()) <= kBulkLimit)
    {
        for (int i : toClear)
            if (!SetItemStateNotify(i, 0, PGGIS_SELECTED))
                return false;
        for (int i : toSet)
            if (!SetItemStateNotify(i, PGGIS_SELECTED, PGGIS_SELECTED))
                return false;
        return true;
    }

    NMPGGITEM nm = MakeItemNotify(-1);
    nm.uChanged = PGGIS_SELECTED;
    nm.uNewState = toSet.empty() ? 0 : PGGIS_SELECTED;
    nm.uOldState = toClear.empty() ? 0 : PGGIS_SELECTED;
    LRESULT veto = Notify(PGGN_ITEMCHANGING, &nm.hdr);
    if (Changed(gen))
        return false;
    if (veto)
        return true;
    for (int i : toClear)
        ApplyState(i, StateOf(i) & ~PGGIS_SELECTED);
    for (int i : toSet)
        ApplyState(i, StateOf(i) | PGGIS_SELECTED);
    InvalidateRect(m_hWnd, nullptr, FALSE);
    Notify(PGGN_ITEMCHANGED, &nm.hdr);
    return !Changed(gen);
}

bool CPhotoGalleryMgr::ToggleItem(int i)
{
    const bool bSel = (StateOf(i) & PGGIS_SELECTED) != 0;
    return SetItemStateNotify(i, bSel ? 0 : PGGIS_SELECTED, PGGIS_SELECTED);
}

void CPhotoGalleryMgr::SetHot(int i)
{
    if (i == m_iHot)
        return;
    const int prev = m_iHot;
    m_iHot = i;
    InvalidateItem(prev);
    InvalidateItem(i);
    if (m_dwStyle & PGGS_HOTTRACK)
    {
        NMPGGITEM nm = MakeItemNotify(i);
        nm.iPrevItem = prev;
        Notify(PGGN_HOTITEMCHANGE, &nm.hdr);
    }
}

void CPhotoGalleryMgr::UpdateHotFromCursor()
{
    if (!(m_dwStyle & (PGGS_HOTTRACK | PGGS_CHECKBOXES)) || m_bDead)
        return;
    POINT pt;
    GetCursorPos(&pt);
    if (WindowFromPoint(pt) != m_hWnd)
    {
        SetHot(-1);
        return;
    }
    ScreenToClient(m_hWnd, &pt);
    SetHot(HitTestClient(pt, nullptr));
}

// ---- Input ------------------------------------------------------------------------------------

void CPhotoGalleryMgr::OnLButtonDown(WPARAM keys, POINT pt)
{
    if (GetFocus() != m_hWnd)
        SetFocus(m_hWnd);
    if (m_bDead)
        return;
    UINT uHit = 0;
    const int i = HitTestClient(pt, &uHit);
    CancelPress();
    m_press.bActive = true;
    m_press.iItem = i;
    m_press.pt = pt;
    SetCapture(m_hWnd);

    const bool bCtrl = (keys & MK_CONTROL) != 0, bShift = (keys & MK_SHIFT) != 0;
    if (i < 0)
    {
        if (!bCtrl && !bShift)
            ClearSelection();
        return;
    }
    if ((uHit & PGGHT_ONCHECK) && IsMulti())
    {
        m_press.bOnCheck = true;
        if (ToggleItem(i) && SetFocusItem(i))
            m_iAnchor = i;
        return;
    }
    if (IsMulti() && bShift)
    {
        int a = (m_iAnchor >= 0 && m_iAnchor < Count()) ? m_iAnchor : i;
        if (SelectRange(a, i, bCtrl))
            SetFocusItem(i);
    }
    else if (IsMulti() && bCtrl)
    {
        if (ToggleItem(i) && SetFocusItem(i))
            m_iAnchor = i;
    }
    else if ((StateOf(i) & PGGIS_SELECTED) && m_nSelected > 1)
    {
        // Pressing inside a multi-selection may start a drag of all of it;
        // collapse to this item only if the button comes up without one.
        m_press.bDeferredSelect = true;
        if (SetFocusItem(i))
            m_iAnchor = i;
    }
    else
    {
        if (SelectOnly(i) && SetFocusItem(i))
            m_iAnchor = i;
    }
}

void CPhotoGalleryMgr::OnLButtonUp(WPARAM keys, POINT pt)
{
    if (!m_press.bActive || m_press.bRight)
        return;
    const Press press = m_press;
    CancelPress();
    if (GetCapture() == m_hWnd)
        ReleaseCapture();

    UINT uHit = 0;
    const int i = HitTestClient(pt, &uHit);
    if (press.bDeferredSelect && i == press.iItem)
    {
        if (!SelectOnly(i))
            return;
    }
    NMPGGITEM nm = MakeItemNotify(i, &pt, uHit, KeyFlagsFromMouse(keys));
    Notify(NM_CLICK, &nm.hdr);
}

void CPhotoGalleryMgr::OnLButtonDblClk(WPARAM keys, POINT pt)
{
    if (GetFocus() != m_hWnd)
        SetFocus(m_hWnd);
    if (m_bDead)
        return;
    UINT uHit = 0;
    const int i = HitTestClient(pt, &uHit);
    if (i >= 0 && (uHit & PGGHT_ONCHECK) && IsMulti())
    {
        ToggleItem(i);      // a quick second click on the check circle is just another toggle
        return;
    }
    const UINT gen = m_itemsGen;
    NMPGGITEM nm = MakeItemNotify(i, &pt, uHit, KeyFlagsFromMouse(keys));
    Notify(NM_DBLCLK, &nm.hdr);
    if (Changed(gen) || i < 0)
        return;
    Notify(PGGN_ITEMACTIVATE, &nm.hdr);
}

void CPhotoGalleryMgr::OnRButtonDown(WPARAM keys, POINT pt)
{
    if (GetFocus() != m_hWnd)
        SetFocus(m_hWnd);
    if (m_bDead)
        return;
    const int i = HitTestClient(pt, nullptr);
    CancelPress();
    m_press.bActive = true;
    m_press.bRight = true;
    m_press.iItem = i;
    m_press.pt = pt;

    // As in Explorer: right-clicking outside the selection selects that item
    // alone, inside it keeps the selection, so the menu acts on what is selected.
    if (i >= 0)
    {
        if (!(StateOf(i) & PGGIS_SELECTED))
        {
            if (!SelectOnly(i))
                return;
        }
        if (SetFocusItem(i))
            m_iAnchor = i;
    }
    else if (!(keys & MK_CONTROL))
    {
        ClearSelection();
    }
}

bool CPhotoGalleryMgr::OnRButtonUp(WPARAM keys, POINT pt)
{
    if (!m_press.bActive || !m_press.bRight)
        return false;
    CancelPress();
    UINT uHit = 0;
    const int i = HitTestClient(pt, &uHit);
    NMPGGITEM nm = MakeItemNotify(i, &pt, uHit, KeyFlagsFromMouse(keys));
    return Notify(NM_RCLICK, &nm.hdr) != 0 || m_bDead;
}

bool CPhotoGalleryMgr::OnContextMenu(int xScreen, int yScreen)
{
    const bool bKeyboard = xScreen == -1 && yScreen == -1;
    POINT pt{};
    int i = -1;
    UINT uHit = PGGHT_NOWHERE;
    if (bKeyboard)
    {
        // Shift+F10 or the Menu key: the focused item, menu at its center.
        i = m_iFocus;
        RECT rc;
        if (i >= 0 && EnsureVisible(i, true) && GetItemClientRect(i, rc))
        {
            pt = POINT{ (rc.left + rc.right) / 2, (rc.top + rc.bottom) / 2 };
            uHit = PGGHT_ONIMAGE;
        }
        else
        {
            i = -1;
            pt = POINT{ Px(8), Px(8) };
        }
    }
    else
    {
        pt = POINT{ xScreen, yScreen };
        ScreenToClient(m_hWnd, &pt);
        i = HitTestClient(pt, &uHit);
    }
    NMPGGITEM nm = MakeItemNotify(i, &pt, uHit, KeyFlagsFromKeyboard() | (bKeyboard ? PGGKF_KEYBOARD : 0));
    return Notify(PGGN_CONTEXTMENU, &nm.hdr) != 0 || m_bDead;
}

void CPhotoGalleryMgr::OnMouseMove(WPARAM keys, POINT pt)
{
    if (m_dwStyle & (PGGS_HOTTRACK | PGGS_CHECKBOXES))
    {
        if (!m_bTrackingLeave)
        {
            TRACKMOUSEEVENT tme{ sizeof(tme), TME_LEAVE, m_hWnd, 0 };
            m_bTrackingLeave = TrackMouseEvent(&tme) != FALSE;
        }
        SetHot(HitTestClient(pt, nullptr));
        if (m_bDead)
            return;
    }

    if (m_press.bActive && !m_press.bRight && !m_press.bOnCheck && m_press.iItem >= 0 &&
        (abs(pt.x - m_press.pt.x) > GetSystemMetrics(SM_CXDRAG) ||
         abs(pt.y - m_press.pt.y) > GetSystemMetrics(SM_CYDRAG)))
    {
        const Press press = m_press;
        CancelPress();
        if (GetCapture() == m_hWnd)
            ReleaseCapture();
        NMPGGITEM nm = MakeItemNotify(press.iItem, &press.pt, PGGHT_ONIMAGE, KeyFlagsFromMouse(keys));
        Notify(PGGN_BEGINDRAG, &nm.hdr);
    }
}

void CPhotoGalleryMgr::OnMouseLeave()
{
    m_bTrackingLeave = false;
    SetHot(-1);
}

void CPhotoGalleryMgr::OnMouseWheel(WPARAM wParam, POINT ptScreen)
{
    const int delta = GET_WHEEL_DELTA_WPARAM(wParam);
    const WORD keys = GET_KEYSTATE_WPARAM(wParam);
    POINT pt = ptScreen;
    ScreenToClient(m_hWnd, &pt);

    if ((keys & MK_CONTROL) && (m_dwStyle & PGGS_USERZOOM))
    {
        // Precision touchpads deliver pinch as Ctrl+wheel with small deltas.
        if (m_metrics.mode == PGGTS_FIXEDCOLUMNS)
        {
            m_zoomWheelAccum += delta;
            const int steps = m_zoomWheelAccum / WHEEL_DELTA;
            if (steps)
            {
                m_zoomWheelAccum -= steps * WHEEL_DELTA;
                ZoomSteps(steps, &pt, PGGMC_USERZOOM);
            }
        }
        else
        {
            ZoomFactor(std::pow(1.15, delta / (double)WHEEL_DELTA), &pt, PGGMC_USERZOOM);
        }
        return;
    }

    UINT lines = 3;
    SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, &lines, 0);
    const int perNotch = lines == WHEEL_PAGESCROLL ? m_cyClient : (int)lines * LinePx();
    m_wheelAccum += delta * perNotch;
    const int dy = m_wheelAccum / WHEEL_DELTA;
    m_wheelAccum -= dy * WHEEL_DELTA;
    if (dy)
        ScrollTo(m_scrollY - dy);
}

bool CPhotoGalleryMgr::OnGesture(HGESTUREINFO hgi)
{
    if (!(m_dwStyle & PGGS_USERZOOM))
        return false;
    GESTUREINFO gi{};
    gi.cbSize = sizeof(gi);
    if (!GetGestureInfo(hgi, &gi) || gi.dwID != GID_ZOOM)
        return false;

    POINT pt{ gi.ptsLocation.x, gi.ptsLocation.y };
    ScreenToClient(m_hWnd, &pt);
    if (gi.dwFlags & GF_BEGIN)
    {
        m_gestureStartDist = gi.ullArguments;
        m_gestureBaseTile = m_dZoomTile;
        m_gestureBaseCols = m_metrics.nColumns;
    }
    else if (m_gestureStartDist)
    {
        // ullArguments is the distance between the two fingers.
        const double ratio = (double)gi.ullArguments / (double)m_gestureStartDist;
        if (m_metrics.mode == PGGTS_FIXEDCOLUMNS)
        {
            int cols = Clamp((int)std::lround(m_gestureBaseCols / ratio), 1, MaxColumns());
            if (cols != m_metrics.nColumns)
            {
                PGGTILEMETRICS m = m_metrics;
                m.nColumns = cols;
                ApplyMetrics(m, PGGMC_USERZOOM, &pt);
            }
        }
        else if (m_dZoomTile > 0)
        {
            ZoomFactor(m_gestureBaseTile * ratio / m_dZoomTile, &pt, PGGMC_USERZOOM);
        }
    }
    if (gi.dwFlags & GF_END)
        m_gestureStartDist = 0;
    CloseGestureInfoHandle(hgi);
    return true;
}

bool CPhotoGalleryMgr::OnKeyDown(UINT vk)
{
    const UINT gen = m_itemsGen;
    NMPGGKEYDOWN kd{};
    kd.wVKey = (WORD)vk;
    kd.flags = KeyFlagsFromKeyboard();
    if (Notify(PGGN_KEYDOWN, &kd.hdr) || Changed(gen))
        return true;

    const bool bCtrl = (kd.flags & PGGKF_CONTROL) != 0, bShift = (kd.flags & PGGKF_SHIFT) != 0;
    switch (vk)
    {
    case VK_LEFT: case VK_RIGHT: case VK_UP: case VK_DOWN:
    case VK_HOME: case VK_END: case VK_PRIOR: case VK_NEXT:
    {
        EnsureLayout();
        const int next = m_layout->Navigate(m_iFocus, vk, m_cyClient);
        if (next < 0)
            return true;
        if (IsMulti() && bShift)
        {
            int a = (m_iAnchor >= 0 && m_iAnchor < Count()) ? m_iAnchor : next;
            if (!SelectRange(a, next, bCtrl) || !SetFocusItem(next))
                return true;
        }
        else if (IsMulti() && bCtrl)
        {
            if (!SetFocusItem(next))
                return true;
        }
        else
        {
            if (!SelectOnly(next) || !SetFocusItem(next))
                return true;
            m_iAnchor = next;
        }
        EnsureVisible(next, false);
        return true;
    }
    case VK_SPACE:
        if (m_iFocus >= 0)
        {
            if (IsMulti() && bCtrl)
                ToggleItem(m_iFocus);
            else if (SelectOnly(m_iFocus))
                m_iAnchor = m_iFocus;
        }
        return true;
    case VK_RETURN:
        if (m_iFocus >= 0)
        {
            NMPGGITEM nm = MakeItemNotify(m_iFocus, nullptr, PGGHT_ONIMAGE, kd.flags | PGGKF_KEYBOARD);
            RECT rc;
            if (GetItemClientRect(m_iFocus, rc))
            {
                nm.ptAction = POINT{ (rc.left + rc.right) / 2, (rc.top + rc.bottom) / 2 };
                nm.ptScreen = nm.ptAction;
                ClientToScreen(m_hWnd, &nm.ptScreen);
            }
            Notify(PGGN_ITEMACTIVATE, &nm.hdr);
        }
        return true;
    case 'A':
        if (bCtrl && IsMulti())
        {
            SelectAll();
            return true;
        }
        break;
    case VK_OEM_PLUS: case VK_ADD:
    case VK_OEM_MINUS: case VK_SUBTRACT:
    case '0': case VK_NUMPAD0:
        if (bCtrl && (m_dwStyle & PGGS_USERZOOM))
        {
            POINT center{ m_cxClient / 2, m_cyClient / 2 };
            int steps = (vk == VK_OEM_PLUS || vk == VK_ADD) ? 1 : (vk == VK_OEM_MINUS || vk == VK_SUBTRACT) ? -1 : 0;
            ZoomSteps(steps, &center, PGGMC_USERZOOM);
            return true;
        }
        break;
    }
    return false;
}

void CPhotoGalleryMgr::OnVScroll(UINT code)
{
    EnsureLayout();
    switch (code)
    {
    case SB_LINEUP:     ScrollTo(m_scrollY - LinePx()); break;
    case SB_LINEDOWN:   ScrollTo(m_scrollY + LinePx()); break;
    case SB_PAGEUP:     ScrollTo(m_scrollY - (std::max)(LinePx(), m_cyClient - LinePx())); break;
    case SB_PAGEDOWN:   ScrollTo(m_scrollY + (std::max)(LinePx(), m_cyClient - LinePx())); break;
    case SB_TOP:        ScrollTo(0); break;
    case SB_BOTTOM:     ScrollTo(MAXINT); break;
    case SB_THUMBTRACK:
    case SB_THUMBPOSITION:
    {
        // nTrackPos is 32-bit; HIWORD(wParam) would cap the content at 65,535 pixels.
        SCROLLINFO si{ sizeof(si), SIF_TRACKPOS };
        GetScrollInfo(m_hWnd, SB_VERT, &si);
        ScrollTo(si.nTrackPos);
        break;
    }
    }
}

void CPhotoGalleryMgr::OnSize(int cx, int cy)
{
    if (cx == m_cxClient && cy == m_cyClient)
        return;
    Anchor a = CaptureAnchor(nullptr);
    const PGGeometry old = m_geo;
    const bool bWidth = cx != m_cxClient;
    m_cxClient = cx;
    m_cyClient = cy;
    if (bWidth)
        m_bLayoutDirty = true;
    EnsureLayout();
    RestoreAnchor(a);
    if (old.cellCx != m_geo.cellCx || old.cellCy != m_geo.cellCy)
        BeginInteractive();
    m_hintFrom = m_hintTo = -1;
}

void CPhotoGalleryMgr::OnTimer(UINT_PTR id)
{
    if (id == IDT_RETRY)
    {
        KillTimer(m_hWnd, IDT_RETRY);
        InvalidateRect(m_hWnd, nullptr, FALSE);     // the next paint queues the retries
    }
    else if (id == IDT_SETTLE)
    {
        KillTimer(m_hWnd, IDT_SETTLE);
        m_bInteractive = false;
        InvalidateRect(m_hWnd, nullptr, FALSE);     // re-sharpen, and queue upgrades
    }
}

void CPhotoGalleryMgr::OnFocusChange(bool bGot)
{
    NMHDR hdr{};
    Notify(bGot ? NM_SETFOCUS : NM_KILLFOCUS, &hdr);
    if (!m_bDead)
        InvalidateRect(m_hWnd, nullptr, FALSE);
}

// ---- Messages ---------------------------------------------------------------------------------

int CPhotoGalleryMgr::OnInsertItem(const PGGITEM* p)
{
    if (IsOwner() || !p || p->cbSize < sizeof(PGGITEM))
        return -1;
    const int idx = Clamp(p->iItem, 0, (int)m_items.size());

    Item it;
    it.key = (p->mask & PGGIF_KEY) ? p->key : (kAutoKeyBit | m_nextAutoKey++);
    if (p->mask & PGGIF_PARAM)
        it.lParam = p->lParam;
    if ((p->mask & PGGIF_TEXT) && p->pszText)
        it.text = p->pszText;
    if ((p->mask & PGGIF_PATH) && p->pszPath)
        it.path = p->pszPath;
    if (p->mask & PGGIF_OVERLAY)
        it.iOverlay = p->iOverlay;
    if (p->mask & PGGIF_SPAN)
    {
        it.span.cx = (BYTE)Clamp(p->cxSpan, 1, PGG_MAXSPAN);
        it.span.cy = (BYTE)Clamp(p->cySpan, 1, PGG_MAXSPAN);
    }
    if (p->mask & PGGIF_THUMB)
        SetItemThumb(it, p->hbmThumb, p->thumbFlags);

    m_items.insert(m_items.begin() + idx, std::move(it));
    m_states.insert(m_states.begin() + idx, 0);
    auto shift = [idx](int& v) { if (v >= idx) ++v; };
    shift(m_iFocus);
    shift(m_iAnchor);
    m_iHot = -1;
    if ((p->mask & PGGIF_STATE) && (p->stateMask & p->state & PGGIS_SELECTED))
    {
        if (!IsMulti())
            for (int i = 0; i < (int)m_states.size(); ++i)
                if (m_states[i])
                    ApplyState(i, StateOf(i) & ~PGGIS_SELECTED);
        ApplyState(idx, StateOf(idx) | PGGIS_SELECTED);
    }
    if ((p->mask & PGGIF_STATE) && (p->stateMask & p->state & PGGIS_FOCUSED))
        m_iFocus = idx;
    ItemsChanged();
    return idx;
}

BOOL CPhotoGalleryMgr::OnDeleteItem(int i)
{
    if (IsOwner() || i < 0 || i >= (int)m_items.size())
        return FALSE;
    const UINT gen = m_itemsGen;
    NMPGGITEM nm = MakeItemNotify(i);
    Notify(PGGN_DELETEITEM, &nm.hdr);
    if (Changed(gen) || i >= (int)m_items.size())
        return FALSE;

    Unpin(m_items[i]);
    if (m_states[i] & PGGIS_SELECTED)
        --m_nSelected;
    m_items.erase(m_items.begin() + i);
    m_states.erase(m_states.begin() + i);
    const int n = (int)m_items.size();
    auto fix = [i, n](int& v) {
        if (v == i)
            v = n ? (std::min)(i, n - 1) : -1;
        else if (v > i)
            --v;
    };
    fix(m_iFocus);
    fix(m_iAnchor);
    m_iHot = -1;
    CancelPress();
    ItemsChanged();
    return TRUE;
}

BOOL CPhotoGalleryMgr::OnDeleteAllItems()
{
    if (IsOwner())
        return OnSetItemCount(0);
    const UINT gen = m_itemsGen;
    for (int i = 0; i < (int)m_items.size(); ++i)
    {
        NMPGGITEM nm = MakeItemNotify(i);
        Notify(PGGN_DELETEITEM, &nm.hdr);
        if (Changed(gen))
            return FALSE;
    }
    for (Item& it : m_items)
        Unpin(it);
    m_items.clear();
    m_states.clear();
    m_nSelected = 0;
    m_iFocus = m_iAnchor = m_iHot = -1;
    m_scrollY = 0;
    CancelPress();
    ItemsChanged();
    return TRUE;
}

BOOL CPhotoGalleryMgr::OnSetItem(const PGGITEM* p)
{
    if (!p || p->cbSize < sizeof(PGGITEM) || p->iItem < 0 || p->iItem >= Count())
        return FALSE;
    const int i = p->iItem;
    BOOL bOk = TRUE;
    if (!IsOwner())
    {
        Item& it = m_items[i];
        if ((p->mask & PGGIF_KEY) && p->key != it.key)
        {
            // Carry host pixels over to the new key.
            PGPixels px;
            bool bMove = false;
            if (it.bPinned)
            {
                PGThumb* old = m_cache.Find(it.key);
                if (old && !old->src.empty())
                {
                    px = old->src;
                    bMove = true;
                }
                Unpin(it);
            }
            it.key = p->key;
            if (bMove)
            {
                PGThumb& t = m_cache.Get(it.key);
                ++t.nPins;
                it.bPinned = true;
                StoreSource(t, std::move(px), TileEdge());
            }
            m_keyIndexGen = (UINT)-1;
        }
        if (p->mask & PGGIF_THUMB)
            bOk = SetItemThumb(it, p->hbmThumb, p->thumbFlags) ? TRUE : FALSE;
        if (p->mask & PGGIF_TEXT)
            it.text = p->pszText ? p->pszText : L"";
        if ((p->mask & PGGIF_PATH) && it.path != (p->pszPath ? p->pszPath : L""))
        {
            it.path = p->pszPath ? p->pszPath : L"";
            PGThumb* t = m_cache.Find(it.key);
            if (t && t->nPins == 0)
                t->ResetSource();   // a different file: load it again
        }
        if (p->mask & PGGIF_PARAM)
            it.lParam = p->lParam;
        if (p->mask & PGGIF_OVERLAY)
            it.iOverlay = p->iOverlay;
        if (p->mask & PGGIF_SPAN)
        {
            PGSpan s{ (BYTE)Clamp(p->cxSpan, 1, PGG_MAXSPAN), (BYTE)Clamp(p->cySpan, 1, PGG_MAXSPAN) };
            if (s.cx != it.span.cx || s.cy != it.span.cy)
            {
                it.span = s;
                if (m_layout->NeedsSpans())
                {
                    Anchor a = CaptureAnchor(nullptr);
                    m_bLayoutDirty = true;
                    EnsureLayout();
                    RestoreAnchor(a);
                }
            }
        }
        InvalidateItem(i);
    }
    // State last: its notifications may change everything above.
    if (p->mask & PGGIF_STATE)
        SetItemStateNotify(i, p->state, p->stateMask);
    return bOk;
}

BOOL CPhotoGalleryMgr::OnGetItem(PGGITEM* p)
{
    if (!p || p->cbSize < sizeof(PGGITEM) || p->iItem < 0 || p->iItem >= Count())
        return FALSE;
    const int i = p->iItem;
    if (p->mask & PGGIF_STATE)
        p->state = StateOf(i) & p->stateMask;
    if (IsOwner())
        return TRUE;    // everything else belongs to the host
    const Item& it = m_items[i];
    if (p->mask & PGGIF_KEY)
        p->key = it.key;
    if (p->mask & PGGIF_PARAM)
        p->lParam = it.lParam;
    if (p->mask & PGGIF_OVERLAY)
        p->iOverlay = it.iOverlay;
    if (p->mask & PGGIF_SPAN)
    {
        p->cxSpan = it.span.cx;
        p->cySpan = it.span.cy;
    }
    if ((p->mask & PGGIF_TEXT) && p->pszText && p->cchTextMax > 0)
        wcsncpy_s(p->pszText, p->cchTextMax, it.text.c_str(), _TRUNCATE);
    if ((p->mask & PGGIF_PATH) && p->pszPath && p->cchPathMax > 0)
        wcsncpy_s(p->pszPath, p->cchPathMax, it.path.c_str(), _TRUNCATE);
    if (p->mask & PGGIF_THUMB)
    {
        // The control never hands out bitmaps; report whether pixels are ready.
        p->hbmThumb = it.bThumbCallback ? PGG_THUMBCALLBACK : nullptr;
        PGThumb* t = m_cache.Find(it.key);
        p->thumbFlags = (t && !t->src.empty()) ? 0 : PGGTF_PENDING;
    }
    return TRUE;
}

BOOL CPhotoGalleryMgr::OnSetItemCount(int n)
{
    if (!IsOwner() || n < 0)
        return FALSE;
    m_nOwnerCount = n;
    m_states.resize(n, 0);
    m_nSelected = 0;
    for (BYTE s : m_states)
        if (s & PGGIS_SELECTED)
            ++m_nSelected;
    if (m_iFocus >= n)
        m_iFocus = -1;
    if (m_iAnchor >= n)
        m_iAnchor = -1;
    m_bOwnerSpansValid = false;
    CancelPress();
    ItemsChanged();
    EnsureLayout();
    return TRUE;
}

BOOL CPhotoGalleryMgr::OnSetThumbnail(const PGGTHUMBNAIL* p)
{
    if (!p || p->cbSize < sizeof(PGGTHUMBNAIL))
        return FALSE;
    PGThumb& t = m_cache.Get(p->key);
    if (p->hbm)
    {
        PGPixels px;
        if (!PixelsFromHBITMAP(p->hbm, px))
            return FALSE;
        CapHostPixels(px);
        StoreSource(t, std::move(px), TileEdge());
        t.bProvisional = (p->flags & PGGTF_PROVISIONAL) != 0;
        t.bHostAsked = true;
    }
    else
    {
        if (t.nPins == 0)
            t.ResetSource();
        t.state = PGThumb::Failed;
        t.uFailReason = 0;
    }
    InvalidateKey(p->key);
    return TRUE;
}

BOOL CPhotoGalleryMgr::OnInvalidateThumbnail(const ULONGLONG* pKey)
{
    if (!pKey)
    {
        m_cache.ResetAll();
        InvalidateRect(m_hWnd, nullptr, FALSE);
        return TRUE;
    }
    PGThumb* t = m_cache.Find(*pKey);
    if (!t)
        return TRUE;
    if (t->nPins > 0)
        t->DropDisplay();
    else
        t->ResetSource();
    InvalidateKey(*pKey);
    return TRUE;
}

BOOL CPhotoGalleryMgr::OnRedrawItems(int iFirst, int iLast)
{
    if (iFirst > iLast)
        std::swap(iFirst, iLast);
    iFirst = (std::max)(0, iFirst);
    iLast = (std::min)(Count() - 1, iLast);
    if (iFirst > iLast)
        return TRUE;
    if (IsOwner())
    {
        for (int i = iFirst; i <= iLast; ++i)
            m_frameInfo.erase(i);
        if (m_layout->NeedsSpans())
        {
            // Spans may have changed: ask again for this range.
            m_bOwnerSpansValid = false;
            Anchor a = CaptureAnchor(nullptr);
            m_bLayoutDirty = true;
            EnsureLayout();
            RestoreAnchor(a);
            return TRUE;
        }
    }
    for (int i = iFirst; i <= iLast; ++i)
        InvalidateItem(i);
    return TRUE;
}

BOOL CPhotoGalleryMgr::OnSetTileMetrics(const PGGTILEMETRICS* p)
{
    if (!p || p->cbSize < sizeof(PGGTILEMETRICS))
        return FALSE;
    PGGTILEMETRICS m = m_metrics;
    if (p->mask & PGGTM_MODE)
        m.mode = p->mode;
    if (p->mask & PGGTM_SIZE)
        m.cxTile = p->cxTile;
    if (p->mask & PGGTM_LIMITS)
    {
        m.cxMin = p->cxMin;
        m.cxMax = p->cxMax;
    }
    if (p->mask & PGGTM_COLUMNS)
        m.nColumns = p->nColumns;
    if (p->mask & PGGTM_ASPECT)
    {
        m.nAspectX = p->nAspectX;
        m.nAspectY = p->nAspectY;
    }
    if (p->mask & PGGTM_SPACING)
    {
        m.cxGap = p->cxGap;
        m.cyGap = p->cyGap;
    }
    if (p->mask & PGGTM_MARGINS)
    {
        m.cxMargin = p->cxMargin;
        m.cyMargin = p->cyMargin;
    }
    m_hostMetrics = Normalize(m);
    ApplyMetrics(m, PGGMC_HOST, nullptr);
    return TRUE;
}

BOOL CPhotoGalleryMgr::OnGetTileMetrics(PGGTILEMETRICS* p)
{
    if (!p || p->cbSize < sizeof(PGGTILEMETRICS))
        return FALSE;
    EnsureLayout();
    *p = m_metrics;
    FillActual(*p);
    return TRUE;
}

BOOL CPhotoGalleryMgr::OnSetLayout(UINT kind, UINT flags)
{
    if (kind != PGGLAYOUT_UNIFORM && kind != PGGLAYOUT_MIXED)
        return FALSE;
    if (kind == PGGLAYOUT_UNIFORM)
        flags = 0;
    if (kind == m_layoutKind && flags == m_layoutFlags)
        return TRUE;
    Anchor a = CaptureAnchor(nullptr);
    m_layoutKind = kind;
    m_layoutFlags = flags;
    m_layout = kind == PGGLAYOUT_MIXED ? CreateMixedLayout((flags & PGGMIX_DENSE) != 0) : CreateUniformLayout();
    m_bOwnerSpansValid = false;
    m_bLayoutDirty = true;
    EnsureLayout();
    if (a.iItem >= 0 || a.iItem == -2)
        RestoreAnchor(a);
    m_hintFrom = m_hintTo = -1;
    InvalidateRect(m_hWnd, nullptr, FALSE);
    return TRUE;
}

BOOL CPhotoGalleryMgr::OnSetLoaderSettings(const PGGLOADERSETTINGS* p)
{
    if (!p || p->cbSize < sizeof(PGGLOADERSETTINGS))
        return FALSE;
    PGGLOADERSETTINGS s = m_loaderSettings;
    if (p->mask & PGGLS_FLAGS)
        s.flags = p->flags & (PGGLF_SHELLCACHE | PGGLF_DECODE | PGGLF_ORIENTATION | PGGLF_CLOUDRECALL | PGGLF_CLOUDTHUMBS);
    if (p->mask & PGGLS_THREADS)
        s.nThreads = (std::min)(p->nThreads, 16u);
    if (p->mask & PGGLS_CACHEBUDGET)
        s.nCacheBudgetMB = (std::max)(16u, (std::min)(p->nCacheBudgetMB, 8192u));
    if (p->mask & PGGLS_PREFETCH)
        s.nPrefetchScreens = (std::min)(p->nPrefetchScreens, 4u);

    const bool bFlags = s.flags != m_loaderSettings.flags;
    const bool bThreads = s.nThreads != m_loaderSettings.nThreads;
    m_loaderSettings = s;
    m_cache.SetBudget((size_t)s.nCacheBudgetMB * 1024 * 1024);
    if (m_loader)
    {
        if (bThreads)
        {
            StopLoader();
            StartLoader();
        }
        else if (bFlags)
        {
            m_loader->SetFlags(s.flags);
        }
    }
    if (bFlags)
    {
        // A different policy may succeed where the old one failed (e.g. cloud recall).
        m_cache.ResetFailed();
        InvalidateRect(m_hWnd, nullptr, FALSE);
    }
    return TRUE;
}

BOOL CPhotoGalleryMgr::OnGetLoaderSettings(PGGLOADERSETTINGS* p)
{
    if (!p || p->cbSize < sizeof(PGGLOADERSETTINGS))
        return FALSE;
    *p = m_loaderSettings;
    p->mask = PGGLS_FLAGS | PGGLS_THREADS | PGGLS_CACHEBUDGET | PGGLS_PREFETCH;
    return TRUE;
}

BOOL CPhotoGalleryMgr::OnSetColors(const PGGCOLORS* p)
{
    if (!p || p->cbSize < sizeof(PGGCOLORS))
        return FALSE;
    if (p->mask & PGGCLR_BK)
        m_colors.clrBk = p->clrBk;
    if (p->mask & PGGCLR_PLACEHOLDER)
        m_colors.clrPlaceholder = p->clrPlaceholder;
    if (p->mask & PGGCLR_SELECTION)
        m_colors.clrSelection = p->clrSelection;
    if (p->mask & PGGCLR_FOCUS)
        m_colors.clrFocus = p->clrFocus;
    if (p->mask & PGGCLR_HOT)
        m_colors.clrHot = p->clrHot;
    m_cache.DropAllDisplays();      // letterbox bars use the placeholder color
    InvalidateRect(m_hWnd, nullptr, FALSE);
    return TRUE;
}

BOOL CPhotoGalleryMgr::OnGetColors(PGGCOLORS* p)
{
    if (!p || p->cbSize < sizeof(PGGCOLORS))
        return FALSE;
    p->mask = PGGCLR_BK | PGGCLR_PLACEHOLDER | PGGCLR_SELECTION | PGGCLR_FOCUS | PGGCLR_HOT;
    p->clrBk = ColorBk();
    p->clrPlaceholder = ColorPlaceholder();
    p->clrSelection = ColorSelection();
    p->clrFocus = m_colors.clrFocus;
    p->clrHot = ColorHot();
    return TRUE;
}

BOOL CPhotoGalleryMgr::OnSetItemState(int i, const PGGITEM* p)
{
    if (!p)
        return FALSE;
    if (i == -1)
    {
        // All items: only selection makes sense.
        if (p->stateMask & PGGIS_SELECTED)
            return ((p->state & PGGIS_SELECTED) ? SelectAll() : ClearSelection()) ? TRUE : FALSE;
        return TRUE;
    }
    if (i < 0 || i >= Count())
        return FALSE;
    SetItemStateNotify(i, p->state, p->stateMask);
    return TRUE;
}

int CPhotoGalleryMgr::OnGetNextItem(int iStart, UINT flags)
{
    const int n = Count();
    if (flags & PGGNI_FOCUSED)
        return (m_iFocus > iStart && m_iFocus < n) ? m_iFocus : -1;
    for (int i = (std::max)(0, iStart + 1); i < n; ++i)
    {
        if (!(flags & PGGNI_SELECTED) || (m_states[i] & PGGIS_SELECTED))
            return i;
    }
    return -1;
}
