#include "pch.h"
#include "grid32.h"
#include "grid32_internal.h"
#include <new>
#include <exception>
#include <cstddef>
#include <windowsx.h>
#include <commctrl.h>   // NM_CLICK / NM_KEYDOWN / NM_SETFOCUS / NM_KILLFOCUS / NM_RETURN

// Cell coordinates arrive by pointer (see grid32.h). They used to be packed
// into one WPARAM as 16-bit halves, which capped every message at 65535 rows
// and columns, and GM_SETCURRENTCELL packed them the other way round from
// everything else (M2, M3).
static const GRIDPOINT* GridPointArg(WPARAM w)
{
    return reinterpret_cast<const GRIDPOINT*>(w);
}

LRESULT CALLBACK CGrid32Mgr::Grid32_WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    CGrid32Mgr* pMgr = reinterpret_cast<CGrid32Mgr*>(GetWindowLongPtr(hWnd, 0));
    if (pMgr == nullptr && uMsg != WM_NCCREATE)
    {
        // Messages can arrive before WM_NCCREATE (WM_GETMINMAXINFO) and after
        // WM_NCDESTROY. There is nothing to route them to; a message box here
        // used to re-enter the message pump from inside the window procedure
        // (H22).
        return DefWindowProc(hWnd, uMsg, wParam, lParam);
    }

    try
    {
        switch (uMsg)
        {
        case WM_NCCREATE:
        {
            // nothrow new: throwing across the Win32 callback boundary is
            // implementation-defined and unsafe on most ABIs. The outer
            // catch(bad_alloc) below requires a valid pMgr for SetLastError,
            // and at WM_NCCREATE we don't have one yet.
            pMgr = new (std::nothrow) CGrid32Mgr();
            if (!pMgr)
                return FALSE;

            pMgr->m_hWndGrid = hWnd;
            SetWindowLongPtr(hWnd, 0, reinterpret_cast<LONG_PTR>(pMgr));
            break;
        }
        case WM_CREATE:
        {
            CREATESTRUCT* pCreateStruct = reinterpret_cast<CREATESTRUCT*>(lParam);
            if (pCreateStruct->lpCreateParams == NULL)
            {
                return -1; // Fail creation
            }

            PGRIDCREATESTRUCT pGridCreateStruct = reinterpret_cast<PGRIDCREATESTRUCT>(pCreateStruct->lpCreateParams);
            if (pGridCreateStruct->cbSize != sizeof(GRIDCREATESTRUCT))
            {
                return -1; // Fail creation
            }

            pGridCreateStruct->style = pCreateStruct->style;

            if (!pMgr->Create(pGridCreateStruct))
            {
                return -1; // Fail creation
            }
        }
        break;
        case WM_NCDESTROY:
            delete pMgr;
            pMgr = nullptr;
            SetWindowLongPtr(hWnd, 0, 0);
            break;
        case WM_DESTROY:
            pMgr->OnDestroy();
            break;
        case WM_MOVE:
        {
            int x = (int)(short)LOWORD(lParam);
            int y = (int)(short)HIWORD(lParam);
            pMgr->OnMove(x, y);
        }
        break;
        case WM_SIZE:
        {
            UINT nType = (UINT)wParam;
            int cx = (int)LOWORD(lParam);
            int cy = (int)HIWORD(lParam);
            pMgr->OnSize(nType, cx, cy);
        }
        break;
        case WM_HSCROLL:
            pMgr->OnHScroll(LOWORD(wParam), HIWORD(wParam), (HWND)lParam);
            return 0;
        case WM_VSCROLL:
            pMgr->OnVScroll(LOWORD(wParam), HIWORD(wParam), (HWND)lParam);
            return 0;
        case WM_KEYDOWN:
            {
                GRIDNMHDR h{};
                h.cellPt = pMgr->m_currentCell;
                h.vkey   = (UINT)wParam;
                h.flags  = (UINT)lParam;
                pMgr->SendGridNotification(NM_KEYDOWN, &h);
                if (wParam == VK_RETURN)
                    pMgr->SendGridNotification(NM_RETURN, &h);
            }
            if (pMgr->OnKeyDown((UINT)wParam, (UINT)lParam, 0))
            {
                pMgr->Invalidate(false);
                return 0;
            }

            break;
        case WM_CHAR:
            // Typing starts an edit here, from the character the keyboard
            // layout and IME produced (H14).
            if (pMgr->OnChar((WCHAR)wParam))
                return 0;
            break;
        case WM_SETFOCUS:
            pMgr->SendGridNotification(NM_SETFOCUS);
            break;
        case WM_KILLFOCUS:
            pMgr->SendGridNotification(NM_KILLFOCUS);
            break;
        case WM_KEYUP:
            if (pMgr->OnKeyUp((UINT)wParam, (UINT)lParam, 0))
            {
                // Handle key up actions if necessary
                return 0;
            }
            break;

        case WM_LBUTTONDOWN:
            pMgr->OnLButtonDown((UINT)wParam, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
            break;
        case WM_LBUTTONDBLCLK:
            pMgr->OnLButtonDblClk((UINT)wParam, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
            return 0;

        case WM_LBUTTONUP:
            pMgr->OnLButtonUp((UINT)wParam, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
            break;
        case WM_MOUSEMOVE:
            pMgr->OnMouseMove((UINT)wParam, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
            break;
        case WM_MOUSEHOVER:
            pMgr->OnMouseHover((UINT)wParam, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
            break;
        case WM_MOUSELEAVE:
            pMgr->OnMouseLeave((UINT)wParam, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
            break;
        case WM_NCMOUSEMOVE:
            pMgr->OnNcMouseMove((UINT)wParam, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
            break;
        case WM_NCHITTEST:
        {
            bool bHitTested = false;
            pMgr->m_gridHitTest = pMgr->OnNcHitTest((UINT)wParam, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), bHitTested);
            if (bHitTested)
            {
                pMgr->SetCursorBasedOnNcHitTest();
                return HTCLIENT;
            }
            break;
        }
        case WM_MOUSEWHEEL:
            pMgr->OnMouseWheel((short)GET_WHEEL_DELTA_WPARAM(wParam), (UINT)wParam, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
            return 0;
        case WM_RBUTTONDOWN:
            pMgr->OnRButtonDown((UINT)wParam, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
            break;

        case WM_RBUTTONUP:
            pMgr->OnRButtonUp((UINT)wParam, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
            break;
        case WM_TIMER:
            pMgr->OnTimer((UINT_PTR)wParam);
            break;
        case WM_CAPTURECHANGED:
        case WM_CANCELMODE:
            pMgr->CancelDragState();
            break;
        case WM_NCPAINT:
            // Handle non-client area painting
            break;

        case WM_ERASEBKGND:
            return TRUE;

        case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            if (hdc)
                pMgr->OnPaint(ps);

            EndPaint(hWnd, &ps);
            return 0; // Indicate that we have handled the message
        }
        case WM_PRINTCLIENT:
        {
            // Paint the whole client area into the DC we are given, so
            // PrintWindow, AnimateWindow and screenshots see the grid.
            PAINTSTRUCT ps{};
            ps.hdc = reinterpret_cast<HDC>(wParam);
            GetClientRect(hWnd, &ps.rcPaint);
            if (ps.hdc)
                pMgr->OnPaint(ps);
            return 0;
        }
        case WM_CLEAR:
            pMgr->OnClear();
            return 0;

        case WM_COPY:
            pMgr->OnCopy();
            return 0;

        case WM_CUT:
            pMgr->OnCut();
            return 0;

        case WM_PASTE:
            pMgr->OnPaste();
            return 0;

        case WM_UNDO:
        case GM_REDO:
            {
                // GN_CELLCHANGED names what the undo actually changed, and is
                // not sent at all when there was nothing to undo (S13).
                GRIDSELECTION changed{};
                const bool ok = uMsg == WM_UNDO ? pMgr->OnUndo(&changed) : pMgr->OnRedo(&changed);
                if (ok)
                {
                    GRIDNMHDR h{};
                    h.cellPt = changed.start;
                    h.range = changed;
                    pMgr->SendGridNotification(GN_CELLCHANGED, &h);
                }
                return ok;
            }
        case GM_CANUNDO:
            return pMgr->OnCanUndo();
        case GM_CANREDO:
            return pMgr->OnCanRedo();
        case GM_SETCELL:
        {
            const GRIDPOINT* pt = GridPointArg(wParam);
            if (!pt || !lParam)
            {
                pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
                return FALSE;
            }
            if (!pMgr->OnSetCellInfo(*pt, *reinterpret_cast<const GCCELLINFO*>(lParam)))
                return FALSE;
            GRIDNMHDR h{}; h.cellPt = *pt; h.range = { *pt, *pt };
            pMgr->SendGridNotification(GN_CELLCHANGED, &h);
            return TRUE;
        }
        case GM_GETLASTERROR:
            return pMgr->dwError;

        case GM_SETCOLUMNWIDTH:
            {
                size_t oldW = pMgr->GetColumnWidth((size_t)wParam);
                bool ok = pMgr->SetColumnWidth((size_t)wParam, (size_t)lParam);
                if (ok)
                {
                    GRIDNMHDR h{}; h.cellPt = { 0, (UINT)wParam };
                    h.oldSize = oldW; h.newSize = (size_t)lParam;
                    pMgr->SendGridNotification(GN_COLWIDTHCHANGED, &h);
                }
                return ok;
            }

        case GM_GETCOLUMNWIDTH:
            return pMgr->GetColumnWidth((size_t)wParam);

        case GM_SETROWHEIGHT:
            {
                size_t oldH = pMgr->GetRowHeight((size_t)wParam);
                bool ok = pMgr->SetRowHeight((size_t)wParam, (size_t)lParam);
                if (ok)
                {
                    GRIDNMHDR h{}; h.cellPt = { (UINT)wParam, 0 };
                    h.oldSize = oldH; h.newSize = (size_t)lParam;
                    pMgr->SendGridNotification(GN_ROWHEIGHTCHANGED, &h);
                }
                return ok;
            }

        case GM_GETROWHEIGHT:
            return pMgr->GetRowHeight((size_t)wParam);

        case GM_SETSELECTION:
            // No GN_SELCHANGE for a selection that was rejected (S13).
            if (!lParam || !pMgr->OnSetSelection(reinterpret_cast<const GRIDSELECTION*>(lParam)))
            {
                if (!lParam)
                    pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
                return FALSE;
            }
            pMgr->NotifySelChange();
            return TRUE;

        case GM_GETSELECTION:
            pMgr->OnGetSelection(reinterpret_cast<GRIDSELECTION*>(lParam));
            return lParam != 0;

        case GM_SETCHARFORMAT:
            {
                // These used to fall through to DefWindowProc and return 0,
                // so every CGridCtrl wrapper reported failure (H17).
                if (!lParam)
                {
                    pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
                    return FALSE;
                }
                const WPARAM which = wParam & ~(WPARAM)SCF_MASKED;
                CELLFONT fi;
                DWORD mask = GCFM_ALL;
                GRIDSELECTION range{};
                if (which == SCF_RANGE || (wParam & SCF_MASKED))
                {
                    LPGCCELLCHARFORMAT pCF = reinterpret_cast<LPGCCELLCHARFORMAT>(lParam);
                    if (pCF->m_cbSize >= sizeof(GCCELLCHARFORMAT))
                        mask = pCF->m_dwMask;
                    else if (pCF->m_cbSize != offsetof(GCCELLCHARFORMAT, m_dwMask))
                    {
                        pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
                        return FALSE;
                    }
                    fi = CellFontFrom(pCF->m_format);
                    range = pCF->m_range;
                }
                else
                {
                    fi = CellFontFrom(*reinterpret_cast<PFONTINFO>(lParam));
                }

                GRIDNMHDR h{};
                switch (which)
                {
                case SCF_RANGE:
                    pMgr->SetRangeFormat(range, fi, mask);
                    pMgr->NormalizeSelectionRect(range);
                    h.range = range;
                    h.cellPt = range.start;
                    break;
                case SCF_SELECTION:
                    pMgr->SetSelectionFormat(fi, mask);
                    h.range = pMgr->GetNormalizedSelection();
                    h.cellPt = pMgr->m_currentCell;
                    break;
                case SCF_CURRENTCELL:
                    pMgr->SetCellFormat(pMgr->m_currentCell.nRow, pMgr->m_currentCell.nCol, fi, mask);
                    h.cellPt = pMgr->m_currentCell;
                    h.range = { h.cellPt, h.cellPt };
                    break;
                default:
                    pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
                    return FALSE;
                }
                if (pMgr->dwError)
                    return FALSE;
                pMgr->SendGridNotification(GN_CELLCHANGED, &h);
                return TRUE;
            }

        case GM_SETCELLTEXT:
            {
                const GRIDPOINT* pt = GridPointArg(wParam);
                if (!pt)
                {
                    pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
                    return FALSE;
                }
                pMgr->SetCellText(*pt, reinterpret_cast<LPCWSTR>(lParam));
                if (pMgr->dwError)
                    return FALSE;
                GRIDNMHDR h{}; h.cellPt = *pt; h.range = { *pt, *pt };
                pMgr->SendGridNotification(GN_CELLCHANGED, &h);
                return TRUE;
            }

        case GM_GETCELLTEXT:
        {
            const GRIDPOINT* pt = GridPointArg(wParam);
            if (!pt || !lParam)
            {
                pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
                return 0;
            }

            return pMgr->OnGetCellText(*pt, reinterpret_cast<GRID_GETTEXT*>(lParam));
        }
        case GM_SETREDRAW:
            pMgr->m_bRedraw = wParam != 0 ? TRUE : FALSE;
            if (pMgr->m_bRedraw)
                pMgr->Invalidate();
            return 0;

        case GM_GETREDRAW:
            return pMgr->m_bRedraw ? TRUE : FALSE;

        case GM_SETCURRENTCELL:
            // Both forms move through MoveCurrentCell, which collapses the
            // selection onto the cell and sends the one GN_SELCHANGE.
            if (wParam)
                return pMgr->OnSetCurrentCell(reinterpret_cast<LPCWSTR>(lParam), (short)wParam);
            if (!lParam)
            {
                pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
                return FALSE;
            }
            {
                const GRIDPOINT* pt = reinterpret_cast<const GRIDPOINT*>(lParam);
                return pMgr->OnSetCurrentCell(pt->nRow, pt->nCol);
            }
        case GM_GETCURRENTCELL:
            if (!lParam)
            {
                pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
                return FALSE;
            }
            *reinterpret_cast<GRIDPOINT*>(lParam) = pMgr->m_currentCell;
            return TRUE;
        case GM_INCREMENTCURRENTCELL:
            pMgr->OnIncrementCell((UINT)wParam, (GRIDPOINT *)lParam);
            return 0;

        case GM_FILLCELLS:
            if (lParam)
            {
                const GCFILLSTRUCT& fs = *(LPGCFILLSTRUCT)lParam;
                GRIDSELECTION range = fs.m_dwFillRange == FILLBYRANGE ? fs.m_range : pMgr->GetNormalizedSelection();
                pMgr->NormalizeSelectionRect(range);
                if (!pMgr->OnFillCells(wParam, fs))
                    return FALSE;
                GRIDNMHDR h{}; h.range = range; h.cellPt = range.start;
                pMgr->SendGridNotification(GN_CELLCHANGED, &h);
                return TRUE;
            }
            pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
            return FALSE;
        case GM_SORTCELLS:
            if (lParam)
            {
                LPGCSORTSTRUCT ss = (GCSORTSTRUCT*)lParam;
                if (!pMgr->OnSortCells(wParam, *ss))
                    return FALSE;
                GRIDSELECTION range = ss->m_sortSel;
                pMgr->NormalizeSelectionRect(range);
                GRIDNMHDR h{}; h.range = range; h.cellPt = range.start;
                pMgr->SendGridNotification(GN_CELLCHANGED, &h);
                return TRUE;
            }
            pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
            return FALSE;

        case GM_FILTERCELLS:
            if (lParam)
                return pMgr->OnFilterCells(wParam, *(GCFILTERSTRUCT*)lParam);
            pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
            return FALSE;
        case GM_FINDTEXT:
            if (lParam)
                return pMgr->OnFindText(*(LPGCFINDSTRUCT)lParam);
            return FALSE;
        case GM_REPLACETEXT:
            if (lParam)
            {
                BOOL ok = pMgr->OnReplaceText(*(LPGCREPLACESTRUCT)lParam);
                if (ok)
                {
                    GRIDNMHDR h{}; h.cellPt = pMgr->m_currentCell;
                    h.range = { h.cellPt, h.cellPt };
                    pMgr->SendGridNotification(GN_CELLCHANGED, &h);
                }
                return ok;
            }
            return FALSE;
        case GM_REPLACEALL:
            if (lParam)
            {
                UINT count = pMgr->OnReplaceAll(*(LPGCREPLACESTRUCT)lParam);
                if (count)
                {
                    GRIDSELECTION used{};
                    pMgr->OnGetUsedRange(&used);
                    GRIDNMHDR h{}; h.cellPt = used.start; h.range = used;
                    pMgr->SendGridNotification(GN_CELLCHANGED, &h);
                }
                return count;
            }
            return 0;

        case GM_SETFORMAT:
            // wParam: SCF_CURRENTCELL / SCF_SELECTION
            // lParam: UINT format code (FMT_*).
            {
                GRIDNMHDR h{};
                if (wParam == SCF_SELECTION)
                {
                    pMgr->SetSelectionNumberFormat((UINT)lParam);
                    h.range = pMgr->GetNormalizedSelection(); h.cellPt = pMgr->m_currentCell;
                }
                else if (wParam == SCF_CURRENTCELL)
                {
                    pMgr->SetCellNumberFormat(pMgr->m_currentCell.nRow,
                                              pMgr->m_currentCell.nCol,
                                              (UINT)lParam);
                    h.cellPt = pMgr->m_currentCell;
                    h.range = { h.cellPt, h.cellPt };
                }
                else
                {
                    pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
                    return FALSE;
                }
                if (pMgr->dwError)
                    return FALSE;
                pMgr->SendGridNotification(GN_CELLCHANGED, &h);
                return TRUE;
            }

        case GM_GETCHARFORMAT:
            // lParam: PFONTINFO out-param. Reads the current cell's font.
            if (lParam) {
                CELLFONT fi;
                if (pMgr->GetCurrentCellFontInfo(fi)) {
                    *reinterpret_cast<PFONTINFO>(lParam) = FontInfoFrom(fi);
                    return TRUE;
                }
            }
            return FALSE;

        case GM_SETALIGN:
            // wParam: GA_HORIZ / GA_VERT / GA_WRAP
            // lParam: DT_* value for horiz/vert axis, or BOOL for wrap.
            if (wParam != GA_HORIZ && wParam != GA_VERT && wParam != GA_WRAP)
            {
                pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
                return FALSE;
            }
            pMgr->SetSelectionAlignBits((UINT)wParam, wParam == GA_WRAP ? (lParam ? 1u : 0u) : (UINT)lParam);
            if (pMgr->dwError)
                return FALSE;
            {
                GRIDNMHDR h{}; h.range = pMgr->GetNormalizedSelection(); h.cellPt = pMgr->m_currentCell;
                pMgr->SendGridNotification(GN_CELLCHANGED, &h);
            }
            return TRUE;

        case GM_GETALIGN:
            // wParam: GA_HORIZ / GA_VERT / GA_WRAP
            // Returns the matching DT_* value (horiz/vert) or 0/1 (wrap)
            // for the current cell.
            {
                PGRIDCELL pCell = pMgr->GetCellOrDefault(pMgr->m_currentCell.nRow,
                                                         pMgr->m_currentCell.nCol);
                switch (wParam) {
                case GA_HORIZ: return (LRESULT)(pCell->justification & (DT_LEFT | DT_CENTER | DT_RIGHT));
                case GA_VERT:  return (LRESULT)(pCell->justification & (DT_TOP | DT_VCENTER | DT_BOTTOM));
                case GA_WRAP:  return (pCell->justification & DT_WORDBREAK) ? 1 : 0;
                }
            }
            return 0;

        case GM_GETFORMAT:
            // Returns the current cell's FMT_* number-format code.
            return (LRESULT)pMgr->GetCellOrDefault(pMgr->m_currentCell.nRow,
                                                   pMgr->m_currentCell.nCol)->m_nFormat;

        case GM_SELECTALL:
            if (!pMgr->OnSetSelection(nullptr))
                return FALSE;
            pMgr->NotifySelChange();
            return TRUE;

        case GM_ENUMCELLS:
            // The count copied, or with a null array the count available.
            // The value was discarded before (H18).
            return (LRESULT)pMgr->OnEnumCells(reinterpret_cast<GRIDPOINT*>(lParam), (UINT)wParam);

        case GM_COMMITEDIT:
            if (wParam)
                return pMgr->CommitEdit();
            if (!pMgr->IsEditing())
                return FALSE;
            pMgr->CancelEdit();
            return TRUE;

        case GM_SETBACKGROUND:
            {
                GRIDNMHDR h{};
                if (wParam == SCF_SELECTION)
                {
                    pMgr->SetSelectionBackground((COLORREF)lParam);
                    h.range = pMgr->GetNormalizedSelection();
                }
                else if (wParam == SCF_CURRENTCELL)
                {
                    h.range = { pMgr->m_currentCell, pMgr->m_currentCell };
                    pMgr->SetRangeBackground(h.range, (COLORREF)lParam);
                }
                else
                {
                    pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
                    return FALSE;
                }
                if (pMgr->dwError)
                    return FALSE;
                h.cellPt = pMgr->m_currentCell;
                pMgr->SendGridNotification(GN_CELLCHANGED, &h);
                return TRUE;
            }
        case GM_GETBACKGROUND:
            return (LRESULT)pMgr->GetCellOrDefault(pMgr->m_currentCell.nRow,
                                                   pMgr->m_currentCell.nCol)->clrBackground;

        case GM_MERGECELLS:
            {
                GRIDSELECTION sel = lParam ? *reinterpret_cast<const GRIDSELECTION*>(lParam)
                                           : pMgr->GetNormalizedSelection();
                const BOOL ok = wParam ? pMgr->MergeCells(sel) : pMgr->UnmergeCells(sel);
                if (ok)
                {
                    pMgr->ExpandForMerges(sel);
                    GRIDNMHDR h{}; h.range = sel; h.cellPt = sel.start;
                    pMgr->SendGridNotification(GN_CELLCHANGED, &h);
                    pMgr->NotifySelChange();
                }
                return ok;
            }
        case GM_GETMERGE:
            {
                const MERGERANGE* m = pMgr->FindMerge(pMgr->m_currentCell.nRow, pMgr->m_currentCell.nCol);
                if (m && lParam)
                    *reinterpret_cast<MERGERANGE*>(lParam) = *m;
                return m != nullptr;
            }

        case GM_GETUSEDRANGE:
            return pMgr->OnGetUsedRange(reinterpret_cast<GRIDSELECTION*>(lParam));
        case GM_RENDERRANGE:
            if (!lParam)
            {
                pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
                return FALSE;
            }
            return pMgr->OnRenderRange(*reinterpret_cast<const GCRENDERSTRUCT*>(lParam));

        case GM_STREAMIN:
        case GM_STREAMOUT:
            // TRUE when the stream's m_dwError came back 0. These used to fall
            // through to DefWindowProc, so CGridCtrl::StreamIn/StreamOut
            // reported failure every time (H17).
            if (!lParam)
            {
                pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
                return FALSE;
            }
            {
                LPGCSTREAM pStream = reinterpret_cast<LPGCSTREAM>(lParam);
                pStream->m_dwError = GRID_ERROR_INVALID_PARAMETER;   // until the call says otherwise
                if (uMsg == GM_STREAMIN)
                    pMgr->OnStreamIn(pStream);
                else
                    pMgr->OnStreamOut(pStream);
                pMgr->SetLastError(pStream->m_dwError);
                return pStream->m_dwError == 0;
            }

        case GM_CLEARALL:
            pMgr->ClearAll();
            return 0;


        }
    }
    catch (const std::bad_alloc&)
    {
        _ASSERT(pMgr != nullptr);
        if (pMgr)
            pMgr->SetLastError(GRID_ERROR_OUT_OF_MEMORY);
        return -1;
    }
	catch (const std::exception&)
	{
        _ASSERT(pMgr != nullptr);
		if (pMgr)
			pMgr->SetLastError(GRID_ERROR_NOT_IMPLEMENTED);
		return -1;
	}
	catch (...)
    {
        _ASSERT(pMgr != nullptr);
        if (pMgr)
            pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
        return -1;
    }

    return DefWindowProc(hWnd, uMsg, wParam, lParam);
}


