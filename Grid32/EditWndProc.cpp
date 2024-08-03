#include "pch.h"
#include "grid32.h"
#include "grid32_internal.h"


LRESULT CALLBACK CGrid32Mgr::EditCtrl_WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    CGrid32Mgr* pMgr = reinterpret_cast<CGrid32Mgr*>(GetWindowLongPtr(GetParent(hWnd), 0));
    if (!pMgr)
        return 0;

    auto setcurcelltext = [&]() {
        int len = GetWindowTextLength(hWnd);
        WCHAR* wsBuff = new WCHAR[(size_t)len + 10];
        memset(wsBuff, 0, ((size_t)len + 10) * sizeof(WCHAR));
        GetWindowText(hWnd, wsBuff, (len + 10));
        pMgr->SetCurrentCellText(wsBuff);
        delete[] wsBuff;
        pMgr->Invalidate();
    };

    switch (message) {
    case WM_KEYDOWN:
        // Handle keydown events
        switch (wParam)
        {
        case VK_RETURN:
        case VK_DOWN:
            SetFocus(GetParent(hWnd));
            pMgr->IncrementSelectedCell(1, 0);
            break;
        case VK_LEFT:
            SetFocus(GetParent(hWnd));
            pMgr->IncrementSelectedCell(0, -1);
            break;
        case VK_RIGHT:
            SetFocus(GetParent(hWnd));
            pMgr->IncrementSelectedCell(0, 1);
            break;
        case VK_UP:
            SetFocus(GetParent(hWnd));
            pMgr->IncrementSelectedCell(-1, 0);
            break;
        }
        break;

    case WM_KILLFOCUS:
    {
        setcurcelltext();
        ShowWindow(hWnd, SW_HIDE);
        break;
    }

    default:
        // Pass other messages to the original edit control window procedure
        return CallWindowProc(pMgr->m_editWndProc, hWnd, message, wParam, lParam);
    }

    return 0; // Indicate that the message was handled
}