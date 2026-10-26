#pragma once

// PhotoGalleryCtrl.h : MFC wrapper for the PhotoGalleryGrid control in Grid32.dll.
//
// Header-only, so Grid32.dll itself stays MFC-free: include it after
// <afxwin.h> in an MFC project. Notifications reach the parent the usual
// way (ON_NOTIFY(NM_CLICK, IDC_GALLERY, ...)). In owner-data mode you can
// instead hand the control an IPhotoGalleryDataSource, which answers
// PGGN_GETDISPINFO before the parent sees it.

#include "photogallery.h"

class IPhotoGalleryDataSource
{
public:
    virtual ~IPhotoGalleryDataSource() = default;
    // Fill whatever pItem->mask asks for. sizeWanted is the tile size in pixels.
    virtual void GetDispInfo(PGGITEM* pItem, SIZE sizeWanted) = 0;
    // The items about to be drawn (visible range plus the prefetch ring).
    virtual void CacheHint(int /*iFrom*/, int /*iTo*/) {}
};

class CPhotoGalleryCtrl : public CWnd
{
public:
    // Grid32.dll registers the class when it loads.
    static HMODULE InitPhotoGalleryControl() { return ::LoadLibrary(_T("Grid32.dll")); }

    BOOL Create(DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID)
    {
        return CWnd::Create(PHOTOGALLERY_WNDCLASS_NAME, nullptr, dwStyle | WS_CHILD, rect, pParentWnd, nID);
    }
    BOOL CreateEx(DWORD dwExStyle, DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID)
    {
        return CWnd::CreateEx(dwExStyle, PHOTOGALLERY_WNDCLASS_NAME, nullptr, dwStyle | WS_CHILD, rect, pParentWnd, nID);
    }

    // ---- Items ----
    int InsertItem(const PGGITEM& item) { return (int)Send(PGGM_INSERTITEM, 0, (LPARAM)&item); }
    int InsertItem(int iItem, HBITMAP hbmThumb, LPCWSTR pszPath = nullptr, LPARAM lParam = 0,
                   ULONGLONG key = 0, BYTE cxSpan = 1, BYTE cySpan = 1)
    {
        PGGITEM it = MakeItem(iItem);
        it.mask = PGGIF_THUMB | PGGIF_PARAM | PGGIF_SPAN | (key ? PGGIF_KEY : 0) | (pszPath ? PGGIF_PATH : 0);
        it.hbmThumb = hbmThumb;
        it.pszPath = const_cast<LPWSTR>(pszPath);
        it.lParam = lParam;
        it.key = key;
        it.cxSpan = cxSpan;
        it.cySpan = cySpan;
        return InsertItem(it);
    }
    BOOL DeleteItem(int iItem) { return (BOOL)Send(PGGM_DELETEITEM, (WPARAM)iItem, 0); }
    BOOL DeleteAllItems() { return (BOOL)Send(PGGM_DELETEALLITEMS, 0, 0); }
    BOOL SetItem(const PGGITEM& item) { return (BOOL)Send(PGGM_SETITEM, 0, (LPARAM)&item); }
    BOOL GetItem(PGGITEM& item) const { return (BOOL)Send(PGGM_GETITEM, 0, (LPARAM)&item); }
    int GetItemCount() const { return (int)Send(PGGM_GETITEMCOUNT, 0, 0); }
    BOOL SetItemCount(int nItems, DWORD dwFlags = 0) { return (BOOL)Send(PGGM_SETITEMCOUNT, (WPARAM)nItems, (LPARAM)dwFlags); }
    LPARAM GetItemData(int iItem) const
    {
        PGGITEM it = MakeItem(iItem);
        it.mask = PGGIF_PARAM;
        return GetItem(it) ? it.lParam : 0;
    }
    BOOL SetItemData(int iItem, LPARAM lParam)
    {
        PGGITEM it = MakeItem(iItem);
        it.mask = PGGIF_PARAM;
        it.lParam = lParam;
        return SetItem(it);
    }
    BOOL SetItemSpan(int iItem, BYTE cx, BYTE cy)
    {
        PGGITEM it = MakeItem(iItem);
        it.mask = PGGIF_SPAN;
        it.cxSpan = cx;
        it.cySpan = cy;
        return SetItem(it);
    }

    // ---- Thumbnails ----
    BOOL SetThumbnail(ULONGLONG key, HBITMAP hbm, UINT flags = 0)
    {
        PGGTHUMBNAIL th{ sizeof(PGGTHUMBNAIL), key, hbm, flags };
        return (BOOL)Send(PGGM_SETTHUMBNAIL, 0, (LPARAM)&th);
    }
    BOOL InvalidateThumbnail(ULONGLONG key) { return (BOOL)Send(PGGM_INVALIDATETHUMBNAIL, 0, (LPARAM)&key); }
    BOOL InvalidateAllThumbnails() { return (BOOL)Send(PGGM_INVALIDATETHUMBNAIL, 0, 0); }
    BOOL RedrawItems(int iFirst, int iLast) { return (BOOL)Send(PGGM_REDRAWITEMS, (WPARAM)iFirst, (LPARAM)iLast); }

    // ---- Geometry and look ----
    BOOL SetTileMetrics(const PGGTILEMETRICS& m) { return (BOOL)Send(PGGM_SETTILEMETRICS, 0, (LPARAM)&m); }
    BOOL GetTileMetrics(PGGTILEMETRICS& m) const
    {
        m.cbSize = sizeof(PGGTILEMETRICS);
        return (BOOL)Send(PGGM_GETTILEMETRICS, 0, (LPARAM)&m);
    }
    BOOL SetFixedColumns(int nColumns)
    {
        PGGTILEMETRICS m{ sizeof(PGGTILEMETRICS), PGGTM_MODE | PGGTM_COLUMNS, PGGTS_FIXEDCOLUMNS };
        m.nColumns = nColumns;
        return SetTileMetrics(m);
    }
    BOOL SetTileSize(int cxTile, UINT mode = PGGTS_AUTOFIT)
    {
        PGGTILEMETRICS m{ sizeof(PGGTILEMETRICS), PGGTM_MODE | PGGTM_SIZE, mode };
        m.cxTile = cxTile;
        return SetTileMetrics(m);
    }
    BOOL SetTileAspect(int x, int y)
    {
        PGGTILEMETRICS m{ sizeof(PGGTILEMETRICS), PGGTM_ASPECT };
        m.nAspectX = x;
        m.nAspectY = y;
        return SetTileMetrics(m);
    }
    BOOL SetLayout(UINT layout, UINT flags = 0) { return (BOOL)Send(PGGM_SETLAYOUT, layout, flags); }
    UINT GetLayout() const { return (UINT)Send(PGGM_GETLAYOUT, 0, 0); }
    BOOL SetColors(const PGGCOLORS& c) { return (BOOL)Send(PGGM_SETCOLORS, 0, (LPARAM)&c); }
    BOOL GetColors(PGGCOLORS& c) const
    {
        c.cbSize = sizeof(PGGCOLORS);
        return (BOOL)Send(PGGM_GETCOLORS, 0, (LPARAM)&c);
    }
    HIMAGELIST SetOverlayImageList(HIMAGELIST himl) { return (HIMAGELIST)Send(PGGM_SETIMAGELIST, 0, (LPARAM)himl); }
    HIMAGELIST GetOverlayImageList() const { return (HIMAGELIST)Send(PGGM_GETIMAGELIST, 0, 0); }
    BOOL Zoom(int steps) { return (BOOL)Send(PGGM_ZOOM, (WPARAM)steps, 0); }

    // ---- Self-loading ----
    void EnableAutoLoad(BOOL bEnable) { ModifyStyle(bEnable ? 0 : PGGS_AUTOLOAD, bEnable ? PGGS_AUTOLOAD : 0); }
    BOOL SetLoaderSettings(const PGGLOADERSETTINGS& s) { return (BOOL)Send(PGGM_SETLOADERSETTINGS, 0, (LPARAM)&s); }
    BOOL GetLoaderSettings(PGGLOADERSETTINGS& s) const
    {
        s.cbSize = sizeof(PGGLOADERSETTINGS);
        return (BOOL)Send(PGGM_GETLOADERSETTINGS, 0, (LPARAM)&s);
    }

    // ---- Selection, hit testing, scrolling ----
    BOOL EnsureVisible(int iItem, BOOL bPartialOK = FALSE) { return (BOOL)Send(PGGM_ENSUREVISIBLE, (WPARAM)iItem, bPartialOK); }
    int HitTest(CPoint pt, UINT* pFlags = nullptr) const
    {
        PGGHITTESTINFO hti{ pt, 0, -1 };
        int i = (int)Send(PGGM_HITTEST, 0, (LPARAM)&hti);
        if (pFlags)
            *pFlags = hti.flags;
        return i;
    }
    BOOL GetItemRect(int iItem, LPRECT lpRect) const { return (BOOL)Send(PGGM_GETITEMRECT, (WPARAM)iItem, (LPARAM)lpRect); }
    int GetNextItem(int iStart, UINT flags) const { return (int)Send(PGGM_GETNEXTITEM, (WPARAM)iStart, (LPARAM)flags); }
    BOOL SetItemState(int iItem, UINT state, UINT mask)
    {
        PGGITEM it = MakeItem(iItem);
        it.state = state;
        it.stateMask = mask;
        return (BOOL)Send(PGGM_SETITEMSTATE, (WPARAM)iItem, (LPARAM)&it);
    }
    UINT GetItemState(int iItem, UINT mask) const { return (UINT)Send(PGGM_GETITEMSTATE, (WPARAM)iItem, (LPARAM)mask); }
    int GetSelectedCount() const { return (int)Send(PGGM_GETSELECTEDCOUNT, 0, 0); }
    int GetFocusedItem() const { return GetNextItem(-1, PGGNI_FOCUSED); }
    void SelectAll() { SetItemState(-1, PGGIS_SELECTED, PGGIS_SELECTED); }
    void SelectNone() { SetItemState(-1, 0, PGGIS_SELECTED); }
    int GetTopIndex() const { return (int)Send(PGGM_GETTOPINDEX, 0, 0); }

    // Answer PGGN_GETDISPINFO / PGGN_ODCACHEHINT here instead of in the parent.
    // The source must outlive the control (or be reset to nullptr first).
    void SetDataSource(IPhotoGalleryDataSource* pSource) { m_pDataSource = pSource; }

protected:
    BOOL OnChildNotify(UINT message, WPARAM wParam, LPARAM lParam, LRESULT* pResult) override
    {
        if (message == WM_NOTIFY && m_pDataSource)
        {
            NMHDR* pHdr = reinterpret_cast<NMHDR*>(lParam);
            if (pHdr->code == PGGN_GETDISPINFO)
            {
                NMPGGDISPINFO* pDi = reinterpret_cast<NMPGGDISPINFO*>(lParam);
                m_pDataSource->GetDispInfo(&pDi->item, pDi->sizeWanted);
                if (pResult)
                    *pResult = 0;
                return TRUE;    // answered: the parent does not see it
            }
            if (pHdr->code == PGGN_ODCACHEHINT)
            {
                NMPGGCACHEHINT* pHint = reinterpret_cast<NMPGGCACHEHINT*>(lParam);
                m_pDataSource->CacheHint(pHint->iFrom, pHint->iTo);
                // Fall through: the parent may want the hint too.
            }
        }
        return CWnd::OnChildNotify(message, wParam, lParam, pResult);
    }

private:
    IPhotoGalleryDataSource* m_pDataSource = nullptr;

    LRESULT Send(UINT msg, WPARAM wParam, LPARAM lParam) const
    {
        ASSERT(::IsWindow(m_hWnd));
        return ::SendMessage(m_hWnd, msg, wParam, lParam);
    }
    static PGGITEM MakeItem(int iItem)
    {
        PGGITEM it{};
        it.cbSize = sizeof(PGGITEM);
        it.iItem = iItem;
        it.iOverlay = -1;
        it.cxSpan = it.cySpan = 1;
        return it;
    }
};
