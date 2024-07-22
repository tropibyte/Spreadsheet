#include "pch.h"
#include "grid32.h"
#include "grid32_internal.h"

LRESULT CGrid32Mgr::Grid32_WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    CGrid32Mgr* pMgr = reinterpret_cast<CGrid32Mgr*>(GetWindowLongPtr(hWnd, 0)); 
    if (pMgr == nullptr && uMsg != WM_NCCREATE)
    {
        //Catastrophic failure

    }
    switch (uMsg)
    {
    case WM_NCCREATE:
    {
        pMgr = new CGrid32Mgr();
        SetWindowLongPtr(hWnd, 0, reinterpret_cast<LONG_PTR>(pMgr));
    }
        break;
    case WM_NCDESTROY:
        delete pMgr;
        pMgr = nullptr;
        SetWindowLongPtr(hWnd, 0, 0);
        break;
    }
    return DefWindowProc(hWnd, uMsg, wParam, lParam);
}
