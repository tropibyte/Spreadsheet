#include "pch.h"
#include "grid32.h"
#include "grid32_internal.h"

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
            InvalidateRect(hWnd, NULL, false);
            return 0;
        }

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
        if(hdc)
            pMgr->Paint(ps);

        EndPaint(hWnd, &ps);
        return 0; // Indicate that we have handled the message
    }


    }

    return DefWindowProc(hWnd, uMsg, wParam, lParam);
}


