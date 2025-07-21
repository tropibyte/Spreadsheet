#include "pch.h"
#include "grid32.h"
#include "grid32_internal.h"
#include <new>
#include <exception>

GRIDPOINT MakeGridPointFromWPARAM(WPARAM wParam)
{
    GRIDPOINT pt;
    pt.nCol = LOWORD(wParam);
    pt.nRow = HIWORD(wParam);
    return pt;
}

LRESULT CALLBACK CGrid32Mgr::Grid32_WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    bool bCallDefault = true;
    CGrid32Mgr* pMgr = reinterpret_cast<CGrid32Mgr*>(GetWindowLongPtr(hWnd, 0));
    if (pMgr == nullptr && uMsg != WM_NCCREATE)
    {
        // Catastrophic failure
        MessageBox(hWnd, _T("Catastrophic failure: CGrid32Mgr instance is nullptr."), _T("Error"), MB_ICONERROR);
        return -1; // or an appropriate error code
    }

    try
    {
        switch (uMsg)
        {
        case WM_NCCREATE:
        {
            pMgr = new CGrid32Mgr();
            if (!pMgr)
                return -1;

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
            break;
        case WM_VSCROLL:
            pMgr->OnVScroll(LOWORD(wParam), HIWORD(wParam), (HWND)lParam);
            break;
        case WM_KEYDOWN:
            if (pMgr->OnKeyDown((UINT)wParam, (UINT)lParam, 0))
            {
                pMgr->Invalidate(false);
                return 0;
            }

            break;
        case WM_KEYUP:
            if (pMgr->OnKeyUp((UINT)wParam, (UINT)lParam, 0))
            {
                // Handle key up actions if necessary
                return 0;
            }
            break;

        case WM_LBUTTONDOWN:
            pMgr->OnLButtonDown((UINT)wParam, (int)LOWORD(lParam), (int)HIWORD(lParam));
            break;

        case WM_LBUTTONUP:
            pMgr->OnLButtonUp((UINT)wParam, (int)LOWORD(lParam), (int)HIWORD(lParam));
            break;
        case WM_MOUSEMOVE:
            pMgr->OnMouseMove((UINT)wParam, (int)LOWORD(lParam), (int)HIWORD(lParam));
            break;
        case WM_MOUSEHOVER:
            pMgr->OnMouseHover((UINT)wParam, (int)LOWORD(lParam), (int)HIWORD(lParam));
            break;
        case WM_MOUSELEAVE:
            pMgr->OnMouseLeave((UINT)wParam, (int)LOWORD(lParam), (int)HIWORD(lParam));
            break;
        case WM_NCMOUSEMOVE:
            pMgr->OnNcMouseMove((UINT)wParam, (int)LOWORD(lParam), (int)HIWORD(lParam));
            break;
        case WM_NCHITTEST:
        {
            bool bHitTested = false;
            pMgr->m_gridHitTest = pMgr->OnNcHitTest((UINT)wParam, LOWORD(lParam), HIWORD(lParam), bHitTested);
            if (bHitTested)
            {
                pMgr->SetCursorBasedOnNcHitTest();
                return HTCLIENT;
            }
            break;
        }
        case WM_MOUSEWHEEL:
            pMgr->OnMouseWheel((short)GET_WHEEL_DELTA_WPARAM(wParam), (UINT)wParam, LOWORD(lParam), HIWORD(lParam));
            break;
        case WM_RBUTTONDOWN:
            pMgr->OnRButtonDown((UINT)wParam, LOWORD(lParam), HIWORD(lParam));
            break;

        case WM_RBUTTONUP:
            pMgr->OnRButtonUp((UINT)wParam, LOWORD(lParam), HIWORD(lParam));
            break;
        case WM_TIMER:
            pMgr->OnTimer((UINT_PTR)wParam);
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
        case WM_CLEAR:
            pMgr->OnClear();
            break;

        case WM_COPY:
            pMgr->OnCopy();
            break;

        case WM_CUT:
            pMgr->OnCut();
            break;

        case WM_PASTE:
            pMgr->OnPaste();
            break;

        case WM_UNDO:
            pMgr->OnUndo();
            break;
        case GM_REDO:
            pMgr->OnRedo();
            break;
        case GM_CANUNDO:
            return pMgr->OnCanUndo();
            //break;
        case GM_CANREDO:
            return pMgr->OnCanRedo();
            //break;
        case GM_SETCELL:
        {
            if (lParam)
            {
                GRIDPOINT pt = MakeGridPointFromWPARAM(wParam);
                pMgr->SetCell(pt.nRow, pt.nCol, *reinterpret_cast<PGRIDCELL>(lParam));
            }
            else
            {
                pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
            }
            break;
        }
        case GM_GETLASTERROR:
            return pMgr->dwError;

        case GM_SETCOLUMNWIDTH:
            return pMgr->SetColumnWidth((size_t)wParam, (size_t)lParam);

        case GM_GETCOLUMNWIDTH:
            return pMgr->GetColumnWidth((size_t)wParam);

        case GM_SETROWHEIGHT:
            return pMgr->SetRowHeight((size_t)wParam, (size_t)lParam);

        case GM_GETROWHEIGHT:
            return pMgr->GetRowHeight((size_t)wParam);

        case GM_SETSELECTION:
            pMgr->OnSetSelection(reinterpret_cast<GRIDSELECTION*>(lParam));
            break;

        case GM_GETSELECTION:
            pMgr->OnGetSelection(reinterpret_cast<GRIDSELECTION*>(lParam));
            break;

        case GM_SETCHARFORMAT:
            if (lParam)
            {
                if (wParam == SCF_RANGE)
                {
                    LPGCCELLCHARFORMAT pCF = reinterpret_cast<LPGCCELLCHARFORMAT>(lParam);
                    if (pCF->m_cbSize == sizeof(GCCELLCHARFORMAT))
                        pMgr->SetRangeFormat(pCF->m_range, pCF->m_format);
                }
                else if (wParam == SCF_SELECTION)
                {
                    PFONTINFO pInfo = reinterpret_cast<PFONTINFO>(lParam);
                    pMgr->SetSelectionFormat(*pInfo);
                }
                else
                {
                    PFONTINFO pInfo = reinterpret_cast<PFONTINFO>(lParam);
                    pMgr->SetCellFormat(pMgr->m_currentCell.nRow, pMgr->m_currentCell.nCol, *pInfo);
                }
            }
            break;

        case GM_SETCELLTEXT:
            pMgr->SetCellText(MakeGridPointFromWPARAM(wParam), reinterpret_cast<LPCWSTR>(lParam));
            break;

        case GM_GETCELLTEXT:
        {
            if (!lParam)
            {
                pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
                return 0;
            }

            return pMgr->OnGetCellText(MakeGridPointFromWPARAM(wParam), reinterpret_cast<GRID_GETTEXT*>(lParam));
        }
        case GM_SETREDRAW:
            pMgr->m_bRedraw = wParam != 0 ? TRUE : FALSE;
            break;

        case GM_GETREDRAW:
            return pMgr->m_bRedraw ? TRUE : FALSE;

        case GM_SETCURRENTCELL:
            if (wParam)
            {
                LPCWSTR pwszRef = (LPCWSTR)lParam;
                return pMgr->OnSetCurrentCell(pwszRef, (short)wParam);
            }
            else
            {
                UINT nRow = LOWORD(lParam);
                UINT nCol = HIWORD(lParam);
                return pMgr->OnSetCurrentCell(nRow, nCol);
            }
            break;
        case GM_GETCURRENTCELL:
        {
            return MAKELRESULT(pMgr->m_currentCell.nCol, pMgr->m_currentCell.nRow);
        }
        case GM_INCREMENTCURRENTCELL:
            pMgr->OnIncrementCell((UINT)wParam, (GRIDPOINT *)lParam);
            break;

        case GM_FILLCELLS:
            if (lParam)
                pMgr->OnFillCells(wParam, *(LPGCFILLSTRUCT)lParam);
            break;
        case GM_SORTCELLS:
            if(lParam)
                pMgr->OnSortCells(wParam, *(GCSORTSTRUCT*)lParam);
            break;

        case GM_FILTERCELLS:
            if(lParam)
                pMgr->OnFilterCells(wParam, *(GCFILTERSTRUCT*)lParam);
            break;
        case GM_FINDTEXT:
            // Handle GM_FINDTEXT
            break;

        case GM_ENUMCELLS:
            // Handle GM_ENUMCELLS
            pMgr->OnEnumCells((GRIDCELL*)lParam, (UINT)wParam);
            break;

        case GM_STREAMIN:
            if (lParam)
                pMgr->OnStreamIn(reinterpret_cast<LPGCSTREAM>(lParam));
            break;

        case GM_STREAMOUT:
            if (lParam)
                pMgr->OnStreamOut(reinterpret_cast<LPGCSTREAM>(lParam));
            break;


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


