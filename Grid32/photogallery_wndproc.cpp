#include "pch.h"
#include "PhotoGalleryMgr.h"
#include <new>
#include <exception>
#include <windowsx.h>

bool RegisterPhotoGalleryClass(HINSTANCE hInstance)
{
    WNDCLASSEX wc = { 0 };

    wc.cbSize = sizeof(WNDCLASSEX);
    wc.style = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS | CS_GLOBALCLASS;
    wc.lpfnWndProc = CPhotoGalleryMgr::PhotoGallery_WndProc;
    wc.cbClsExtra = 0;
    wc.cbWndExtra = sizeof(PVOID);
    wc.hInstance = hInstance;
    wc.hIcon = NULL;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = NULL;
    wc.lpszMenuName = NULL;
    wc.lpszClassName = PHOTOGALLERY_WNDCLASS_NAME;
    wc.hIconSm = NULL;

    return RegisterClassEx(&wc) != NULL;
}

bool UnregisterPhotoGalleryClass(HINSTANCE hInstance)
{
    return UnregisterClass(PHOTOGALLERY_WNDCLASS_NAME, hInstance) != NULL;
}

LRESULT CALLBACK CPhotoGalleryMgr::PhotoGallery_WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    CPhotoGalleryMgr* pMgr = reinterpret_cast<CPhotoGalleryMgr*>(GetWindowLongPtr(hWnd, 0));
    if (uMsg == WM_NCCREATE)
    {
        // nothrow: an exception must never cross the Win32 callback boundary.
        pMgr = new (std::nothrow) CPhotoGalleryMgr();
        if (!pMgr)
            return FALSE;
        if (!pMgr->OnNcCreate(hWnd))
        {
            delete pMgr;
            return FALSE;
        }
        SetWindowLongPtr(hWnd, 0, reinterpret_cast<LONG_PTR>(pMgr));
        return DefWindowProc(hWnd, uMsg, wParam, lParam);
    }
    if (pMgr == nullptr)
        return DefWindowProc(hWnd, uMsg, wParam, lParam);

    // A notification handler may destroy the window while we are several
    // frames deep. Count the frames; the outermost one deletes the manager.
    ++pMgr->m_nDepth;
    LRESULT lr = 0;
    try
    {
        lr = pMgr->Dispatch(uMsg, wParam, lParam);
    }
    catch (const std::bad_alloc&)
    {
        lr = 0;
    }
    catch (...)
    {
        lr = 0;
    }
    if (--pMgr->m_nDepth == 0 && pMgr->m_bDead)
        delete pMgr;
    return lr;
}

LRESULT CPhotoGalleryMgr::Dispatch(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
    case WM_CREATE:
        return OnCreate(reinterpret_cast<const CREATESTRUCT*>(lParam));
    case WM_DESTROY:
        OnDestroy();
        break;
    case WM_NCDESTROY:
        m_bDead = true;
        SetWindowLongPtr(m_hWnd, 0, 0);
        return DefWindowProc(m_hWnd, uMsg, wParam, lParam);
    case WM_STYLECHANGING:
        if (wParam == (WPARAM)GWL_STYLE)
            OnStyleChanging(reinterpret_cast<STYLESTRUCT*>(lParam));
        return 0;
    case WM_STYLECHANGED:
        if (wParam == (WPARAM)GWL_STYLE)
        {
            const STYLESTRUCT* pss = reinterpret_cast<const STYLESTRUCT*>(lParam);
            OnStyleChanged(pss->styleOld, pss->styleNew);
        }
        return 0;
    case WM_DPICHANGED_AFTERPARENT:
        OnDpiChanged();
        return 0;
    case WM_SIZE:
        OnSize((int)(short)LOWORD(lParam), (int)(short)HIWORD(lParam));
        return 0;
    case WM_ERASEBKGND:
        return TRUE;
    case WM_PAINT:
        OnPaint();
        return 0;
    case WM_PRINTCLIENT:
    {
        // PrintWindow and hidden windows get no WM_PAINT; do the after-paint
        // work (queue the loader, trim the cache) here too.
        RECT rc{ 0, 0, m_cxClient, m_cyClient };
        Render(reinterpret_cast<HDC>(wParam), rc);
        if (!m_bDead)
            AfterPaint();
        return 0;
    }
    case WM_VSCROLL:
        OnVScroll(LOWORD(wParam));
        return 0;
    case WM_MOUSEWHEEL:
        OnMouseWheel(wParam, POINT{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) });
        return 0;
    case WM_LBUTTONDOWN:
        RelayToTip(uMsg, wParam, lParam);     // a click hides the tooltip
        OnLButtonDown(wParam, POINT{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) });
        return 0;
    case WM_LBUTTONUP:
        RelayToTip(uMsg, wParam, lParam);
        OnLButtonUp(wParam, POINT{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) });
        return 0;
    case WM_LBUTTONDBLCLK:
        RelayToTip(uMsg, wParam, lParam);
        OnLButtonDblClk(wParam, POINT{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) });
        return 0;
    case WM_RBUTTONDOWN:
        RelayToTip(uMsg, wParam, lParam);
        OnRButtonDown(wParam, POINT{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) });
        return 0;
    case WM_RBUTTONUP:
        RelayToTip(uMsg, wParam, lParam);
        // Unhandled, DefWindowProc turns the release into WM_CONTEXTMENU.
        if (OnRButtonUp(wParam, POINT{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) }))
            return 0;
        break;
    case WM_MBUTTONDOWN:
    case WM_MBUTTONUP:
        RelayToTip(uMsg, wParam, lParam);
        break;
    case WM_CONTEXTMENU:
        // Unhandled, DefWindowProc passes it on to the parent.
        if (OnContextMenu(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)))
            return 0;
        break;
    case WM_MOUSEMOVE:
        OnMouseMove(wParam, POINT{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) });
        RelayToTip(uMsg, wParam, lParam);     // after: the first move creates the tooltip
        return 0;
    case WM_MOUSELEAVE:
        OnMouseLeave();
        return 0;
    case WM_CAPTURECHANGED:
        // Someone else has the mouse now: the press is over (nothing to release).
        if (reinterpret_cast<HWND>(lParam) != m_hWnd)
            m_press = Press();
        return 0;
    case WM_CANCELMODE:
        CancelPress();      // and the capture
        break;
    case WM_GESTURE:
        // Handled gestures close their handle; the rest go to DefWindowProc,
        // which turns a touch pan into WM_VSCROLL.
        if (OnGesture(reinterpret_cast<HGESTUREINFO>(lParam)))
            return 0;
        break;
    case WM_KEYDOWN:
        if (OnKeyDown((UINT)wParam))
            return 0;
        break;
    case WM_GETDLGCODE:
    {
        // In a dialog, Enter on a focused item activates it (PGGN_ITEMACTIVATE)
        // instead of pressing the default button. lParam is the message, or NULL.
        LRESULT code = DLGC_WANTARROWS | DLGC_WANTCHARS;
        const MSG* pMsg = reinterpret_cast<const MSG*>(lParam);
        if (pMsg && pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_RETURN && m_iFocus >= 0 && m_iFocus < Count())
            code |= DLGC_WANTMESSAGE;
        return code;
    }
    case WM_SETFOCUS:
        OnFocusChange(true);
        return 0;
    case WM_KILLFOCUS:
        OnFocusChange(false);
        return 0;
    case WM_TIMER:
        OnTimer((UINT_PTR)wParam);
        return 0;
    case WM_PGG_LOADERDONE:
        OnLoaderDone();
        return 0;

    // ---- PGGM_* ----
    case PGGM_INSERTITEM:
        return OnInsertItem(reinterpret_cast<const PGGITEM*>(lParam));
    case PGGM_DELETEITEM:
        return OnDeleteItem((int)wParam);
    case PGGM_DELETEALLITEMS:
        return OnDeleteAllItems();
    case PGGM_SETITEM:
        return OnSetItem(reinterpret_cast<const PGGITEM*>(lParam));
    case PGGM_GETITEM:
        return OnGetItem(reinterpret_cast<PGGITEM*>(lParam));
    case PGGM_GETITEMCOUNT:
        return Count();
    case PGGM_SETITEMCOUNT:
        return OnSetItemCount((int)wParam);
    case PGGM_SETTHUMBNAIL:
        return OnSetThumbnail(reinterpret_cast<const PGGTHUMBNAIL*>(lParam));
    case PGGM_INVALIDATETHUMBNAIL:
        return OnInvalidateThumbnail(reinterpret_cast<const ULONGLONG*>(lParam));
    case PGGM_REDRAWITEMS:
        return OnRedrawItems((int)wParam, (int)lParam);
    case PGGM_SETTILEMETRICS:
        return OnSetTileMetrics(reinterpret_cast<const PGGTILEMETRICS*>(lParam));
    case PGGM_GETTILEMETRICS:
        return OnGetTileMetrics(reinterpret_cast<PGGTILEMETRICS*>(lParam));
    case PGGM_SETLAYOUT:
        return OnSetLayout((UINT)wParam, (UINT)lParam);
    case PGGM_GETLAYOUT:
        return m_layoutKind;
    case PGGM_SETLOADERSETTINGS:
        return OnSetLoaderSettings(reinterpret_cast<const PGGLOADERSETTINGS*>(lParam));
    case PGGM_GETLOADERSETTINGS:
        return OnGetLoaderSettings(reinterpret_cast<PGGLOADERSETTINGS*>(lParam));
    case PGGM_SETCOLORS:
        return OnSetColors(reinterpret_cast<const PGGCOLORS*>(lParam));
    case PGGM_GETCOLORS:
        return OnGetColors(reinterpret_cast<PGGCOLORS*>(lParam));
    case PGGM_SETIMAGELIST:
    {
        HIMAGELIST hOld = m_hOverlay;
        m_hOverlay = reinterpret_cast<HIMAGELIST>(lParam);
        InvalidateRect(m_hWnd, nullptr, FALSE);
        return reinterpret_cast<LRESULT>(hOld);
    }
    case PGGM_GETIMAGELIST:
        return reinterpret_cast<LRESULT>(m_hOverlay);
    case PGGM_ENSUREVISIBLE:
        return EnsureVisible((int)wParam, lParam != 0) ? TRUE : FALSE;
    case PGGM_HITTEST:
    {
        LPPGGHITTESTINFO p = reinterpret_cast<LPPGGHITTESTINFO>(lParam);
        if (!p)
            return -1;
        p->iItem = HitTestClient(p->pt, &p->flags);
        return p->iItem;
    }
    case PGGM_GETITEMRECT:
        return (lParam && GetItemClientRect((int)wParam, *reinterpret_cast<RECT*>(lParam))) ? TRUE : FALSE;
    case PGGM_GETNEXTITEM:
        return OnGetNextItem((int)wParam, (UINT)lParam);
    case PGGM_SETITEMSTATE:
        return OnSetItemState((int)wParam, reinterpret_cast<const PGGITEM*>(lParam));
    case PGGM_GETITEMSTATE:
        return StateOf((int)wParam) & (UINT)lParam;
    case PGGM_GETSELECTEDCOUNT:
        return m_nSelected;
    case PGGM_ZOOM:
    {
        POINT center{ m_cxClient / 2, m_cyClient / 2 };
        return ZoomSteps((int)wParam, &center, PGGMC_USERZOOM) ? TRUE : FALSE;
    }
    case PGGM_GETTOPINDEX:
        return TopIndex();
    }
    return DefWindowProc(m_hWnd, uMsg, wParam, lParam);
}
