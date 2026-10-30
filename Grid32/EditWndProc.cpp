#include "pch.h"
#include "grid32.h"
#include "grid32_internal.h"
#include <new>


// The inline cell editor. The grid owns the edit state (BeginEdit /
// CommitEdit / CancelEdit); this subclass only turns keys and focus changes
// into those calls.
LRESULT CALLBACK CGrid32Mgr::EditCtrl_WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    CGrid32Mgr* pMgr = reinterpret_cast<CGrid32Mgr*>(GetWindowLongPtr(GetParent(hWnd), 0));
    if (!pMgr)
        return 0;

    try {
        switch (message) {
        case WM_KEYDOWN:
            // Navigation keys commit into the cell being edited, then move.
            switch (wParam)
            {
            case VK_RETURN:
                // Per host spec: Enter commits in place, marker stays.
                pMgr->CommitEdit();
                return 0;
            case VK_DOWN:
                pMgr->CommitEdit();
                pMgr->IncrementSelectedCell(1, 0);
                return 0;
            case VK_UP:
                pMgr->CommitEdit();
                pMgr->IncrementSelectedCell(-1, 0);
                return 0;
            case VK_TAB:
                // Tab = commit + right, Shift+Tab = commit + left.
                pMgr->CommitEdit();
                pMgr->IncrementSelectedCell(0, (GetKeyState(VK_SHIFT) & 0x8000) ? -1 : 1);
                return 0;
            case VK_LEFT:
            {
                DWORD startPos = 0, endPos = 0;
                SendMessage(hWnd, EM_GETSEL, (WPARAM)&startPos, (LPARAM)&endPos);
                if (startPos == 0 && endPos == 0) {
                    pMgr->CommitEdit();
                    pMgr->IncrementSelectedCell(0, -1);
                    return 0;
                }
                return CallWindowProc(pMgr->m_editWndProc, hWnd, message, wParam, lParam);
            }
            case VK_RIGHT:
            {
                DWORD startPos = 0, endPos = 0;
                SendMessage(hWnd, EM_GETSEL, (WPARAM)&startPos, (LPARAM)&endPos);
                if (startPos == endPos && endPos == (DWORD)GetWindowTextLength(hWnd)) {
                    pMgr->CommitEdit();
                    pMgr->IncrementSelectedCell(0, 1);
                    return 0;
                }
                return CallWindowProc(pMgr->m_editWndProc, hWnd, message, wParam, lParam);
            }
            case VK_ESCAPE:
                // The cell keeps its pre-edit content.
                pMgr->CancelEdit();
                return 0;
            }
            break;
        case WM_CHAR:
            // Enter, Tab and Escape were handled on key-down; the characters
            // they translate to would only make the edit control beep.
            if (wParam == L'\r' || wParam == L'\n' || wParam == L'\t' || wParam == 0x1B)
                return 0;
            break;
        case WM_KILLFOCUS:
            // Focus left for somewhere else — the ribbon, another window. The
            // text goes into the cell the edit was opened on. A commit or
            // cancel that is moving focus itself has already closed the edit,
            // so this finds nothing to do.
            pMgr->CommitEdit();
            break;
        default:
            break;
        }
    }
    catch (const std::bad_alloc&)
    {
        pMgr->SetLastError(GRID_ERROR_OUT_OF_MEMORY);
    }
    catch (const std::exception&)
    {
        pMgr->SetLastError(GRID_ERROR_INVALID_PARAMETER);
    }

    return CallWindowProc(pMgr->m_editWndProc, hWnd, message, wParam, lParam);
}
