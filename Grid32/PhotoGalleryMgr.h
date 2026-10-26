#pragma once

#include <vector>
#include <string>
#include <memory>
#include <unordered_map>
#include "photogallery.h"
#include "PhotoGalleryLayout.h"
#include "PhotoGalleryThumbs.h"
#include "PhotoGalleryLoader.h"

// One GDI+ session per control. It is the first member of CPhotoGalleryMgr,
// so it is constructed first and destroyed last: every GDI+ object the
// manager owns (cached tiles) is released before GdiplusShutdown runs.
class CGdiplusSession
{
    ULONG_PTR m_token = 0;

public:
    CGdiplusSession()
    {
        Gdiplus::GdiplusStartupInput in;
        if (Gdiplus::GdiplusStartup(&m_token, &in, nullptr) != Gdiplus::Ok)
            m_token = 0;
    }
    ~CGdiplusSession()
    {
        if (m_token)
            Gdiplus::GdiplusShutdown(m_token);
    }
    CGdiplusSession(const CGdiplusSession&) = delete;
    CGdiplusSession& operator=(const CGdiplusSession&) = delete;
    bool Ok() const { return m_token != 0; }
};

class CPhotoGalleryMgr
{
    CGdiplusSession m_gdiplus;      // must stay the first member

    struct Item
    {
        ULONGLONG    key = 0;
        LPARAM       lParam = 0;
        std::wstring text, path;
        int          iOverlay = -1;
        PGSpan       span;
        bool         bThumbCallback = false;
        bool         bPinned = false;   // host pixels live in the cache under key
    };

    // What painting and notifications need about one item. In owner-data
    // mode it comes from PGGN_GETDISPINFO, once per frame.
    struct ItemInfo
    {
        ULONGLONG    key = 0;
        LPARAM       lParam = 0;
        int          iOverlay = -1;
        bool         bCallback = false;
        std::wstring path;
    };

    struct Press
    {
        bool  bActive = false, bRight = false;
        bool  bDeferredSelect = false;  // clicked inside a multi-selection: collapse on release
        bool  bOnCheck = false;
        int   iItem = -1;
        POINT pt{};
    };

    // Keeps the user's place across relayouts (zoom, resize, layout change).
    struct Anchor
    {
        int    iItem = -1;      // -2 = pinned to the top
        double frac = 0.0;      // position inside the item, 0 = top edge
        int    yClient = 0;     // where that point sits in the window
    };

    HWND  m_hWnd = nullptr;
    DWORD m_dwStyle = 0;

    // Notifications are synchronous and a handler may destroy the control.
    // The window procedure counts its nesting and deletes the manager only
    // when the outermost call returns; m_bDead tells inner frames to stop.
    // m_itemsGen tells them the items moved under their feet.
    int   m_nDepth = 0;
    bool  m_bDead = false;
    UINT  m_itemsGen = 0;

    // Items
    std::vector<Item> m_items;          // internal mode
    int   m_nOwnerCount = 0;            // PGGS_OWNERDATA
    std::vector<BYTE> m_states;         // PGGIS_SELECTED per index, both modes
    int   m_nSelected = 0, m_iFocus = -1, m_iAnchor = -1, m_iHot = -1;
    ULONGLONG m_nextAutoKey = 1;
    std::vector<PGSpan> m_ownerSpans;
    bool  m_bOwnerSpansValid = false;
    std::unordered_map<int, ItemInfo> m_frameInfo;
    DWORD m_frameInfoFrame = 0;
    std::unordered_map<ULONGLONG, int> m_keyIndex;  // internal mode, rebuilt lazily
    UINT  m_keyIndexGen = (UINT)-1;

    // Geometry
    PGGTILEMETRICS m_metrics{};          // in effect (user zoom included)
    PGGTILEMETRICS m_hostMetrics{};      // last set by the host: what Ctrl+0 returns to
    double m_dZoomTile = 160.0;         // fractional tile width so small zoom steps add up
    UINT  m_layoutKind = PGGLAYOUT_UNIFORM, m_layoutFlags = 0;
    std::unique_ptr<CGalleryLayout> m_layout;
    PGGeometry m_geo;
    bool  m_bLayoutDirty = true, m_bInLayout = false;
    int   m_dpi = 96, m_cxClient = 0, m_cyClient = 0, m_scrollY = 0;
    int   m_wheelAccum = 0, m_zoomWheelAccum = 0;
    bool  m_bInteractive = false;       // zoom/resize in progress: cheap scaling, sharpen later
    ULONGLONG m_gestureStartDist = 0;
    double m_gestureBaseTile = 0.0;
    int   m_gestureBaseCols = 0;

    PGGCOLORS   m_colors{};
    HIMAGELIST m_hOverlay = nullptr;

    // Back buffer: one 32bpp DIB section, kept across paints.
    HDC     m_hdcBack = nullptr;
    HBITMAP m_hbmBack = nullptr, m_hbmBackOld = nullptr;
    int     m_cxBack = 0, m_cyBack = 0;

    // Thumbnails
    CThumbCache m_cache;
    DWORD m_frame = 0;
    std::unordered_map<ULONGLONG, int> m_visibleKeys;   // key -> index on screen at the last paint
    int   m_hintFrom = -1, m_hintTo = -1;

    PGGLOADERSETTINGS m_loaderSettings{};
    std::unique_ptr<CThumbLoader> m_loader;
    std::wstring m_pathBuf;

    Press m_press;
    bool  m_bTrackingLeave = false;

public:
    CPhotoGalleryMgr();
    ~CPhotoGalleryMgr();

    static LRESULT CALLBACK PhotoGallery_WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

private:
    LRESULT Dispatch(UINT uMsg, WPARAM wParam, LPARAM lParam);

    // Lifetime
    bool OnNcCreate(HWND hWnd);
    int  OnCreate(const CREATESTRUCT* pcs);
    void OnDestroy();
    void OnStyleChanging(STYLESTRUCT* pss);
    void OnStyleChanged(DWORD dwOld, DWORD dwNew);
    void OnDpiChanged();

    // Basics
    bool IsOwner() const { return (m_dwStyle & PGGS_OWNERDATA) != 0; }
    bool IsMulti() const { return (m_dwStyle & PGGS_MULTISELECT) != 0; }
    int  Count() const { return IsOwner() ? m_nOwnerCount : (int)m_items.size(); }
    int  Px(int dip) const { return MulDiv(dip, m_dpi, 96); }
    LRESULT Notify(UINT code, NMHDR* pHdr);
    bool Changed(UINT gen) const { return m_bDead || gen != m_itemsGen; }
    void ItemsChanged();
    bool GetItemInfo(int i, ItemInfo& info);
    NMPGGITEM MakeItemNotify(int i, const POINT* ptClient = nullptr, UINT uHit = 0, UINT uKeyFlags = 0);
    UINT StateOf(int i) const;
    int  KeyToIndex(ULONGLONG key);
    static UINT KeyFlagsFromMouse(WPARAM keys);
    static UINT KeyFlagsFromKeyboard();

    // Colors
    COLORREF ColorBk() const;
    COLORREF ColorPlaceholder() const;
    COLORREF ColorSelection() const;
    COLORREF ColorHot() const;

    // Geometry and scrolling
    void ResolveGeometry(PGGeometry& geo) const;
    void EnsureLayout();
    bool EnsureOwnerSpans();
    int  MaxColumns() const;
    PGGTILEMETRICS Normalize(const PGGTILEMETRICS& m) const;
    void FillActual(PGGTILEMETRICS& m) const;
    bool ApplyMetrics(const PGGTILEMETRICS& proposed, UINT uCause, const POINT* ptAnchor);
    bool ZoomSteps(int steps, const POINT* ptAnchor, UINT uCause);
    bool ZoomFactor(double factor, const POINT* ptAnchor, UINT uCause);
    void BeginInteractive();
    Anchor CaptureAnchor(const POINT* ptClient);
    void RestoreAnchor(const Anchor& a);
    void ClampScroll();
    void UpdateScrollBar();
    void ScrollTo(int y);
    int  LinePx() const;
    int  PrefetchPx() const;
    bool GetItemClientRect(int i, RECT& rc);
    void InvalidateItem(int i);
    void InvalidateKey(ULONGLONG key);
    bool CheckRect(const RECT& rcTile, RECT& rcCheck) const;
    RECT OverlayRect(const RECT& rcTile) const;
    int  HitTestClient(POINT pt, UINT* pFlags);
    bool EnsureVisible(int i, bool bPartialOK);
    int  TopIndex();

    // Painting
    void OnPaint();
    void Render(HDC hdcTarget, const RECT& rcPaint);
    bool EnsureBackBuffer(int cx, int cy);
    void FreeBackBuffer();
    bool DrawTile(Gdiplus::Graphics& g, int i, DWORD dwPrePaint);
    void DrawChrome(Gdiplus::Graphics& g, const RECT& rc, bool bSel, bool bFocus, bool bHot);
    void DrawFailedGlyph(Gdiplus::Graphics& g, const RECT& rc, UINT uReason);
    DWORD CustomDraw(Gdiplus::Graphics& g, DWORD dwStage, int i, const RECT& rc, const ItemInfo* pInfo,
                     UINT uItemState, BOOL fThumbReady);
    void SendCacheHint();
    void AfterPaint();

    // Thumbnails
    int  TileEdge() const { return (std::max)(m_geo.cellCx, m_geo.cellCy); }
    bool NeedsUpgrade(const PGThumb& t, int w, int h) const;
    PGThumb* EnsureThumb(int i, const ItemInfo& info, int w, int h);
    bool EnsureDisplay(PGThumb& t, int w, int h, Gdiplus::Graphics& target);
    void CapHostPixels(PGPixels& px) const;
    void StoreSource(PGThumb& t, PGPixels&& px, int edge);
    bool SetItemThumb(Item& it, HBITMAP hbm, UINT flags);
    void Unpin(Item& it);
    void OnLoaderDone();
    void StartLoader();
    void StopLoader();

    // Selection
    void ApplyState(int i, UINT state);
    bool SetItemStateNotify(int i, UINT state, UINT mask);
    bool SelectRange(int a, int b, bool bKeepOthers);
    bool SelectOnly(int i) { return SelectRange(i, i, false); }
    bool ClearSelection() { return SelectRange(-1, -1, false); }
    bool SelectAll() { return Count() > 0 ? SelectRange(0, Count() - 1, false) : true; }
    bool ToggleItem(int i);
    bool SetFocusItem(int i) { return SetItemStateNotify(i, PGGIS_FOCUSED, PGGIS_FOCUSED); }
    void SetHot(int i);
    void UpdateHotFromCursor();

    // Input
    void OnLButtonDown(WPARAM keys, POINT pt);
    void OnLButtonUp(WPARAM keys, POINT pt);
    void OnLButtonDblClk(WPARAM keys, POINT pt);
    void OnRButtonDown(WPARAM keys, POINT pt);
    bool OnRButtonUp(WPARAM keys, POINT pt);       // true = handled, no WM_CONTEXTMENU
    bool OnContextMenu(int xScreen, int yScreen);  // true = handled
    void OnMouseMove(WPARAM keys, POINT pt);
    void OnMouseLeave();
    void OnMouseWheel(WPARAM wParam, POINT ptScreen);
    bool OnGesture(HGESTUREINFO hgi);
    bool OnKeyDown(UINT vk);
    void OnVScroll(UINT code);
    void OnSize(int cx, int cy);
    void OnTimer(UINT_PTR id);
    void OnFocusChange(bool bGot);
    void CancelPress() { m_press = Press(); }

    // Messages
    int  OnInsertItem(const PGGITEM* p);
    BOOL OnDeleteItem(int i);
    BOOL OnDeleteAllItems();
    BOOL OnSetItem(const PGGITEM* p);
    BOOL OnGetItem(PGGITEM* p);
    BOOL OnSetItemCount(int n);
    BOOL OnSetThumbnail(const PGGTHUMBNAIL* p);
    BOOL OnInvalidateThumbnail(const ULONGLONG* pKey);
    BOOL OnRedrawItems(int iFirst, int iLast);
    BOOL OnSetTileMetrics(const PGGTILEMETRICS* p);
    BOOL OnGetTileMetrics(PGGTILEMETRICS* p);
    BOOL OnSetLayout(UINT kind, UINT flags);
    BOOL OnSetLoaderSettings(const PGGLOADERSETTINGS* p);
    BOOL OnGetLoaderSettings(PGGLOADERSETTINGS* p);
    BOOL OnSetColors(const PGGCOLORS* p);
    BOOL OnGetColors(PGGCOLORS* p);
    BOOL OnSetItemState(int i, const PGGITEM* p);
    int  OnGetNextItem(int iStart, UINT flags);
};
