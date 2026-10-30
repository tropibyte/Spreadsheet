// photogallery.h : public interface of the PhotoGalleryGrid control in Grid32.dll.
//
// An Instagram-style grid of photo tiles. Load Grid32.dll (LoadLibrary is
// enough; the class registers itself) and create a window of class
// PHOTOGALLERY_WNDCLASS_NAME. Everything is configured with PGGM_* messages
// after creation, so the control also works from a dialog template.
//
// The interface is C on purpose: no STL types cross it, every struct starts
// with cbSize, and strings travel as pointer + length.
//
// Data modes
//   Internal (default)  PGGM_INSERTITEM / PGGM_SETITEM / PGGM_DELETEITEM. The
//                       control copies each thumbnail you hand it.
//   PGGS_OWNERDATA       PGGM_SETITEMCOUNT only; the control asks for what it
//                       is about to draw through PGGN_GETDISPINFO.
//   PGGS_AUTOLOAD        (either mode, off by default) items carry a file
//                       path and the control builds thumbnails itself on
//                       worker threads.
//
// Keys: give every item a stable PGGIF_KEY (a database id, a path hash); the
// thumbnail cache is keyed by it. Keep keys below 2^62: the control makes up
// keys above that for items that have none. Owner-data items without a key
// are cached by index, which says nothing about which photo is there, so
// PGGM_SETITEMCOUNT forgets all of those thumbnails and PGGM_REDRAWITEMS
// those of its range, and the host is asked for them again.
//
// Bitmaps: the control never takes ownership of an HBITMAP. It copies the
// pixels before the call (or notification) returns; free yours afterwards.
// The bitmap must not be selected into a DC. 32bpp bitmaps whose alpha
// channel is not all zero are treated as premultiplied (the AlphaBlend
// convention); anything else is treated as opaque.

#pragma once
#ifndef _PHOTOGALLERY_H
#define _PHOTOGALLERY_H

#include <windows.h>
#include <commctrl.h>

#define PHOTOGALLERY_WNDCLASS_NAME  L"PhotoGalleryGrid"

// ---- Window styles (control-specific low word) ----------------------------

#define PGGS_OWNERDATA      0x0001  // items come from PGGN_GETDISPINFO; fixed at creation
#define PGGS_MULTISELECT    0x0002  // Ctrl/Shift extend the selection
#define PGGS_FIT            0x0004  // letterbox each photo in its tile (default: center-crop)
#define PGGS_HOTTRACK       0x0008  // highlight the tile under the mouse, send PGGN_HOTITEMCHANGE
#define PGGS_USERZOOM       0x0010  // Ctrl+wheel, Ctrl +/-/0 and pinch change the tile size
#define PGGS_AUTOLOAD       0x0020  // control opens files and builds thumbnails itself
#define PGGS_CHECKBOXES     0x0040  // check circle on hover toggles selection (with PGGS_MULTISELECT)
#define PGGS_SHOWSELALWAYS  0x0080  // the selection keeps its color without focus (otherwise it turns grey)

// ---- Items ------------------------------------------------------------------

#define PGGIF_THUMB         0x0001  // hbmThumb (+ thumbFlags)
#define PGGIF_TEXT          0x0002  // pszText: the item's tooltip (owner-data: asked for on hover);
                                    // no accessible name yet
#define PGGIF_PATH          0x0004  // pszPath: file loaded under PGGS_AUTOLOAD
#define PGGIF_PARAM         0x0008  // lParam
#define PGGIF_STATE         0x0010  // state / stateMask
#define PGGIF_KEY           0x0020  // key
#define PGGIF_OVERLAY       0x0040  // iOverlay
#define PGGIF_SPAN          0x0080  // cxSpan / cySpan (PGGLAYOUT_MIXED)

#define PGGIS_SELECTED      0x0001
#define PGGIS_FOCUSED       0x0002

// thumbFlags (with PGGIF_THUMB) and PGGTHUMBNAIL.flags
#define PGGTF_PENDING       0x0001  // GETDISPINFO answer: pixels will arrive later via PGGM_SETTHUMBNAIL
#define PGGTF_PROVISIONAL   0x0002  // a stand-in (e.g. a small saved thumbnail); a sharper one may follow

// Internal mode: store this in hbmThumb to have the control ask for the
// pixels through PGGN_GETDISPINFO when it needs them.
#define PGG_THUMBCALLBACK   ((HBITMAP)(LONG_PTR)-1)

#define PGG_MAXSPAN         4

typedef struct tagPGGITEM {
    UINT      cbSize;
    UINT      mask;          // PGGIF_*: which fields are valid (set) or wanted (get)
    int       iItem;
    UINT      state;         // PGGIS_*
    UINT      stateMask;
    HBITMAP   hbmThumb;      // copied by the control; the caller keeps ownership
    UINT      thumbFlags;    // PGGTF_*
    LPWSTR    pszText;       // set: null-terminated; get: buffer of cchTextMax
    int       cchTextMax;
    LPWSTR    pszPath;       // set: null-terminated; get: buffer of cchPathMax
    int       cchPathMax;
    LPARAM    lParam;
    ULONGLONG key;           // stable identity (e.g. a database id); keys the thumbnail cache
    int       iOverlay;      // index into the overlay image list, -1 = none
    BYTE      cxSpan;        // tile size in grid cells for PGGLAYOUT_MIXED, 1..PGG_MAXSPAN
    BYTE      cySpan;
} PGGITEM, *LPPGGITEM;

typedef struct tagPGGTHUMBNAIL {
    UINT      cbSize;
    ULONGLONG key;
    HBITMAP   hbm;           // copied; NULL marks the thumbnail as unavailable
    UINT      flags;         // PGGTF_PROVISIONAL
} PGGTHUMBNAIL, *LPPGGTHUMBNAIL;

// ---- Tile metrics -------------------------------------------------------------
// Sizes are in device-independent pixels (96 per inch) and scale with the
// window's DPI. PGGM_GETTILEMETRICS also reports the geometry in effect.
//
// The scroll position is an int, so the scrolled content stops at about one
// billion physical pixels (INT_MAX / 2): some 3.5 million rows of 300-pixel
// tiles. Rows past that are laid out but cannot be scrolled to.

#define PGGTM_MODE          0x0001
#define PGGTM_SIZE          0x0002  // cxTile
#define PGGTM_LIMITS        0x0004  // cxMin, cxMax
#define PGGTM_COLUMNS       0x0008  // nColumns
#define PGGTM_ASPECT        0x0010  // nAspectX, nAspectY
#define PGGTM_SPACING       0x0020  // cxGap, cyGap
#define PGGTM_MARGINS       0x0040  // cxMargin, cyMargin
#define PGGTM_ALL           0x007F

#define PGGTS_FIXEDCOLUMNS  0   // nColumns columns; tiles stretch with the window (Instagram)
#define PGGTS_FIXEDSIZE     1   // tiles exactly cxTile wide; column count floats (Explorer)
#define PGGTS_AUTOFIT       2   // about cxTile wide; column count floats, tiles stretch to fill (Photos)

typedef struct tagPGGTILEMETRICS {
    UINT cbSize;
    UINT mask;               // PGGTM_*
    UINT mode;               // PGGTS_*
    int  cxTile;             // preferred / exact tile width
    int  cxMin, cxMax;       // clamp for PGGTS_AUTOFIT and for zooming
    int  nColumns;           // PGGTS_FIXEDCOLUMNS
    int  nAspectX, nAspectY; // tile height = width * nAspectY / nAspectX
    int  cxGap, cyGap;       // space between tiles
    int  cxMargin, cyMargin; // space around the grid
    // Filled by PGGM_GETTILEMETRICS: what is in effect now, physical pixels.
    int  cxActual, cyActual, nColumnsActual;
} PGGTILEMETRICS, *LPPGGTILEMETRICS;

// ---- Layout ---------------------------------------------------------------------

#define PGGLAYOUT_UNIFORM   0   // every tile one cell
#define PGGLAYOUT_MIXED     1   // tiles span cxSpan x cySpan cells (PGGIF_SPAN)

#define PGGMIX_DENSE        0x0001  // PGGM_SETLAYOUT lParam: back-fill holes (order may jump)

// ---- Colors -----------------------------------------------------------------------

#define PGGCLR_BK           0x0001
#define PGGCLR_PLACEHOLDER  0x0002
#define PGGCLR_SELECTION    0x0004
#define PGGCLR_FOCUS        0x0008
#define PGGCLR_HOT          0x0010

typedef struct tagPGGCOLORS {
    UINT     cbSize;
    UINT     mask;           // PGGCLR_*
    COLORREF clrBk;          // CLR_DEFAULT restores the system color
    COLORREF clrPlaceholder; // tile shown while a thumbnail is loading
    COLORREF clrSelection;
    COLORREF clrFocus;
    COLORREF clrHot;
} PGGCOLORS, *LPPGGCOLORS;

// ---- Self-loading (PGGS_AUTOLOAD) -------------------------------------------------

#define PGGLF_SHELLCACHE    0x0001  // try the Windows thumbnail cache first           (default on)
#define PGGLF_DECODE        0x0002  // otherwise decode local files with WIC             (default on)
#define PGGLF_ORIENTATION   0x0004  // apply EXIF orientation when decoding              (default on)
#define PGGLF_CLOUDRECALL   0x0008  // may open cloud-only files, which downloads them  (default OFF)
#define PGGLF_CLOUDTHUMBS   0x0010  // ask the cloud provider (OneDrive) for its preview: a small
                                    // network fetch, never a file download            (default on)

#define PGGLS_FLAGS         0x0001
#define PGGLS_THREADS       0x0002
#define PGGLS_CACHEBUDGET   0x0004
#define PGGLS_PREFETCH      0x0008

typedef struct tagPGGLOADERSETTINGS {
    UINT cbSize;
    UINT mask;               // PGGLS_*
    UINT flags;              // PGGLF_*
    UINT nThreads;           // 0 = automatic
    UINT nCacheBudgetMB;     // decoded pixels kept in memory (all modes)
    UINT nPrefetchScreens;   // how far beyond the viewport to load ahead
} PGGLOADERSETTINGS, *LPPGGLOADERSETTINGS;

// ---- Hit testing ----------------------------------------------------------------------

#define PGGHT_NOWHERE       0x0001  // empty space or a gap between tiles
#define PGGHT_ONIMAGE       0x0002
#define PGGHT_ONOVERLAY     0x0004  // the overlay badge
#define PGGHT_ONCHECK       0x0008  // the check circle (PGGS_CHECKBOXES)

typedef struct tagPGGHITTESTINFO {
    POINT pt;                // client coordinates
    UINT  flags;             // out: PGGHT_*
    int   iItem;             // out: -1 = none
} PGGHITTESTINFO, *LPPGGHITTESTINFO;

// PGGM_GETNEXTITEM lParam
#define PGGNI_ALL           0x0000
#define PGGNI_SELECTED      0x0001
#define PGGNI_FOCUSED       0x0002

// PGGM_SETITEMCOUNT lParam
#define PGGSICF_NOSCROLL    0x0001

// ---- Messages ---------------------------------------------------------------------------
// Grid32's GM_* messages use WM_USER+1..; these start well clear of them.

#define PGGM_FIRST              (WM_USER + 0x200)
#define PGGM_INSERTITEM         (PGGM_FIRST + 0)   // lParam LPPGGITEM (iItem = position) -> index or -1
#define PGGM_DELETEITEM         (PGGM_FIRST + 1)   // wParam index -> BOOL
#define PGGM_DELETEALLITEMS     (PGGM_FIRST + 2)   // -> BOOL
#define PGGM_SETITEM            (PGGM_FIRST + 3)   // lParam LPPGGITEM -> BOOL
#define PGGM_GETITEM            (PGGM_FIRST + 4)   // lParam LPPGGITEM -> BOOL
#define PGGM_GETITEMCOUNT       (PGGM_FIRST + 5)   // -> int
#define PGGM_SETITEMCOUNT       (PGGM_FIRST + 6)   // wParam count, lParam PGGSICF_* -> BOOL (PGGS_OWNERDATA)
#define PGGM_SETTHUMBNAIL       (PGGM_FIRST + 7)   // lParam LPPGGTHUMBNAIL -> BOOL
#define PGGM_INVALIDATETHUMBNAIL (PGGM_FIRST + 8)   // lParam const ULONGLONG* key, NULL = all -> BOOL
#define PGGM_REDRAWITEMS        (PGGM_FIRST + 9)   // wParam first, lParam last -> BOOL
#define PGGM_SETTILEMETRICS     (PGGM_FIRST + 10)  // lParam LPPGGTILEMETRICS -> BOOL
#define PGGM_GETTILEMETRICS     (PGGM_FIRST + 11)  // lParam LPPGGTILEMETRICS -> BOOL
#define PGGM_SETLAYOUT          (PGGM_FIRST + 12)  // wParam PGGLAYOUT_*, lParam PGGMIX_* -> BOOL
#define PGGM_GETLAYOUT          (PGGM_FIRST + 13)  // -> PGGLAYOUT_*
#define PGGM_SETLOADERSETTINGS  (PGGM_FIRST + 14)  // lParam LPPGGLOADERSETTINGS -> BOOL
#define PGGM_GETLOADERSETTINGS  (PGGM_FIRST + 15)  // lParam LPPGGLOADERSETTINGS -> BOOL
#define PGGM_SETCOLORS          (PGGM_FIRST + 16)  // lParam LPPGGCOLORS -> BOOL
#define PGGM_GETCOLORS          (PGGM_FIRST + 17)  // lParam LPPGGCOLORS -> BOOL
#define PGGM_SETIMAGELIST       (PGGM_FIRST + 18)  // lParam HIMAGELIST (overlay badges; not owned) -> previous
#define PGGM_GETIMAGELIST       (PGGM_FIRST + 19)  // -> HIMAGELIST
#define PGGM_ENSUREVISIBLE      (PGGM_FIRST + 20)  // wParam index, lParam fPartialOK -> BOOL
#define PGGM_HITTEST            (PGGM_FIRST + 21)  // lParam LPPGGHITTESTINFO -> index or -1
#define PGGM_GETITEMRECT        (PGGM_FIRST + 22)  // wParam index, lParam LPRECT (client) -> BOOL
#define PGGM_GETNEXTITEM        (PGGM_FIRST + 23)  // wParam start (-1 = from the top), lParam PGGNI_* -> index or -1
#define PGGM_SETITEMSTATE       (PGGM_FIRST + 24)  // wParam index (-1 = all), lParam LPPGGITEM -> BOOL
#define PGGM_GETITEMSTATE       (PGGM_FIRST + 25)  // wParam index, lParam mask -> state
#define PGGM_GETSELECTEDCOUNT   (PGGM_FIRST + 26)  // -> int
#define PGGM_ZOOM               (PGGM_FIRST + 27)  // wParam steps (+ in, - out, 0 = reset) -> BOOL
#define PGGM_GETTOPINDEX        (PGGM_FIRST + 28)  // -> first visible index or -1

// ---- Notifications (WM_NOTIFY) ---------------------------------------------------------
// The common codes NM_CLICK, NM_DBLCLK, NM_RCLICK carry an NMPGGITEM;
// NM_SETFOCUS and NM_KILLFOCUS a plain NMHDR; NM_CUSTOMDRAW an
// NMPGGCUSTOMDRAW. Those codes are shared by every control, so check
// hdr.idFrom / hwndFrom before casting.
//
// Notifications are sent synchronously. A handler may delete items or even
// destroy the control; the control re-checks its state when the call returns.

#define PGGN_FIRST              (0U - 2100U)      // Grid32's GN_* use -2000..-2009
#define PGGN_GETDISPINFO        (PGGN_FIRST - 0)   // NMPGGDISPINFO
#define PGGN_ODCACHEHINT        (PGGN_FIRST - 1)   // NMPGGCACHEHINT
#define PGGN_ITEMCHANGING       (PGGN_FIRST - 2)   // NMPGGITEM; return TRUE to veto
#define PGGN_ITEMCHANGED        (PGGN_FIRST - 3)   // NMPGGITEM; iItem -1 = many items
#define PGGN_ITEMACTIVATE       (PGGN_FIRST - 4)   // NMPGGITEM: double-click or Enter
#define PGGN_HOTITEMCHANGE      (PGGN_FIRST - 5)   // NMPGGITEM; iPrevItem = old hot item
#define PGGN_CONTEXTMENU        (PGGN_FIRST - 6)   // NMPGGITEM: mouse or keyboard; return TRUE if handled
#define PGGN_BEGINDRAG          (PGGN_FIRST - 7)   // NMPGGITEM
#define PGGN_DELETEITEM         (PGGN_FIRST - 8)   // NMPGGITEM (internal mode): free lParam here
#define PGGN_KEYDOWN            (PGGN_FIRST - 9)   // NMPGGKEYDOWN; return TRUE to skip default handling
#define PGGN_TILEMETRICSCHANGING (PGGN_FIRST - 10)  // NMPGGTILEMETRICS: user zoom; edit or return TRUE to veto
#define PGGN_TILEMETRICSCHANGED (PGGN_FIRST - 11)  // NMPGGTILEMETRICS
#define PGGN_THUMBLOADED        (PGGN_FIRST - 12)  // NMPGGTHUMB: the control built a thumbnail (copy it to persist)
#define PGGN_THUMBFAILED        (PGGN_FIRST - 13)  // NMPGGTHUMB: uReason says why
#define PGGN_LAST               (PGGN_FIRST - 99)

// NMPGGITEM.uKeyFlags
#define PGGKF_CONTROL       0x0001
#define PGGKF_SHIFT         0x0002
#define PGGKF_ALT           0x0004
#define PGGKF_KEYBOARD      0x0008  // raised from the keyboard (Enter, Shift+F10, Menu key)

typedef struct tagNMPGGITEM {
    NMHDR     hdr;
    int       iItem;         // -1 = empty space (or many items for PGGN_ITEMCHANGED)
    int       iPrevItem;     // PGGN_HOTITEMCHANGE: the previous hot item
    ULONGLONG key;
    LPARAM    lParam;
    UINT      uHitFlags;     // PGGHT_*
    POINT     ptAction;      // client coordinates
    POINT     ptScreen;      // screen coordinates, ready for TrackPopupMenu
    UINT      uKeyFlags;     // PGGKF_*
    UINT      uOldState;     // PGGN_ITEMCHANGING / PGGN_ITEMCHANGED
    UINT      uNewState;
    UINT      uChanged;      // PGGIS_* bits that change
} NMPGGITEM, *LPNMPGGITEM;

typedef struct tagNMPGGDISPINFO {
    NMHDR     hdr;
    PGGITEM    item;          // item.mask says what is wanted; fill those fields
    SIZE      sizeWanted;    // tile size in pixels the thumbnail will be drawn at
} NMPGGDISPINFO, *LPNMPGGDISPINFO;

typedef struct tagNMPGGCACHEHINT {
    NMHDR     hdr;
    int       iFrom;         // visible range plus the prefetch ring
    int       iTo;
} NMPGGCACHEHINT, *LPNMPGGCACHEHINT;

typedef struct tagNMPGGKEYDOWN {
    NMHDR     hdr;
    WORD      wVKey;
    UINT      flags;         // PGGKF_CONTROL / PGGKF_SHIFT / PGGKF_ALT
} NMPGGKEYDOWN, *LPNMPGGKEYDOWN;

#define PGGMC_USERZOOM      1
#define PGGMC_HOST          2

typedef struct tagNMPGGTILEMETRICS {
    NMHDR          hdr;
    PGGTILEMETRICS  metrics;  // CHANGING: the proposal, editable; CHANGED: now in effect
    UINT           uCause;   // PGGMC_*
} NMPGGTILEMETRICS, *LPNMPGGTILEMETRICS;

// NMPGGTHUMB.uSource
#define PGGTSRC_SHELLCACHE  1   // Windows thumbnail cache
#define PGGTSRC_DECODE      2   // decoded from the file with WIC
#define PGGTSRC_PROVIDER    3   // the cloud provider's preview of a file that stays in the cloud

// NMPGGTHUMB.uReason (PGGN_THUMBFAILED)
#define PGGTR_NOTFOUND      1
#define PGGTR_CLOUDONLY     2   // not downloaded and PGGLF_CLOUDRECALL is off
#define PGGTR_UNSUPPORTED   3   // no codec / not an image
#define PGGTR_DECODEERROR   4
#define PGGTR_ACCESSDENIED  5

typedef struct tagNMPGGTHUMB {
    NMHDR             hdr;
    int               iItem;     // -1 if the item is not currently known
    ULONGLONG         key;
    LPCWSTR           pszPath;
    UINT              uSource;   // PGGTSRC_*
    UINT              uReason;   // PGGTR_* (PGGN_THUMBFAILED)
    HRESULT           hr;
    BITMAPINFOHEADER  bmih;      // 32bpp, top-down (biHeight < 0), premultiplied BGRA
    const void*       pvBits;    // valid only during the notification: copy to keep
    int               cbStride;
} NMPGGTHUMB, *LPNMPGGTHUMB;

typedef struct tagNMPGGCUSTOMDRAW {
    NMCUSTOMDRAW  nmcd;      // dwItemSpec = index, rc = tile rect, lItemlParam = lParam
    ULONGLONG     key;
    int           iOverlay;
    BOOL          fThumbReady;
} NMPGGCUSTOMDRAW, *LPNMPGGCUSTOMDRAW;

#endif // _PHOTOGALLERY_H
