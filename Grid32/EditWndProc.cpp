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
        {
            DWORD startPos, endPos;
            SendMessage(hWnd, EM_GETSEL, (WPARAM)&startPos, (LPARAM)&endPos);
            if (startPos == 0) {
                SetFocus(GetParent(hWnd));
                pMgr->IncrementSelectedCell(0, -1);
            }
            else {
                return CallWindowProc(pMgr->m_editWndProc, hWnd, message, wParam, lParam);
            }
        }
            break;
        case VK_RIGHT:
        {
            int textLength = GetWindowTextLength(hWnd);
            DWORD startPos, endPos;
            SendMessage(hWnd, EM_GETSEL, (WPARAM)&startPos, (LPARAM)&endPos);
            if (endPos == textLength) {
                SetFocus(GetParent(hWnd));
                pMgr->IncrementSelectedCell(0, 1);
            }
            else {
                return CallWindowProc(pMgr->m_editWndProc, hWnd, message, wParam, lParam);
            }
        }
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
    case WM_SHOWWINDOW:
        if (wParam) { // Being shown
            // Get the current cell font info and set it to the edit control
            PGRIDCELL pCell = pMgr->GetCurrentCell();
            if (pCell) {
                LOGFONT lf = { 0 };
                if (!pMgr->IsCellFontDefault(*pCell))
                {
                    wcscpy_s(lf.lfFaceName, pCell->fontInfo.m_wsFontFace.c_str());
                    lf.lfHeight = (LONG)(pCell->fontInfo.m_fPointSize * GetDeviceCaps(GetDC(hWnd), LOGPIXELSY) / -72.0);
                    lf.lfItalic = pCell->fontInfo.bItalic;
                    lf.lfUnderline = pCell->fontInfo.bUnderline;
                    lf.lfStrikeOut = pCell->fontInfo.bStrikeThrough;
                    lf.lfWeight = pCell->fontInfo.bWeight;
                    HFONT hNewFont = CreateFontIndirect(&lf);
                    SendMessage(hWnd, WM_SETFONT, (WPARAM)hNewFont, TRUE);
                }
            }
        }
        else { // Being hidden
         // Restore the original font
            HFONT hFont = (HFONT)SendMessage(hWnd, WM_GETFONT, 0, 0);
            if (hFont != pMgr->m_hDefaultFont)
            {
                SendMessage(hWnd, WM_SETFONT, (WPARAM)pMgr->m_hDefaultFont, TRUE);
                DeleteObject(hFont);
            }
        }
        break;
    default:
        // Pass other messages to the original edit control window procedure
        return CallWindowProc(pMgr->m_editWndProc, hWnd, message, wParam, lParam);
    }

    return CallWindowProc(pMgr->m_editWndProc, hWnd, message, wParam, lParam); 
}