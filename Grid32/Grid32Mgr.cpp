#include "pch.h"
#include "Grid32Mgr.h"

CGrid32Mgr::CGrid32Mgr()
{
    memset(&gcs, 0, sizeof(GRIDCREATESTRUCT));
}


CGrid32Mgr::~CGrid32Mgr()
{
}

bool CGrid32Mgr::Create(PGRIDCREATESTRUCT pGCS)
{
    // Store the provided GRIDCREATESTRUCT in the member variable
    gcs = *pGCS;

    // Perform any other initialization needed for your grid

    return true; // Return true if creation succeeded, false otherwise
}

void CGrid32Mgr::Paint(HWND hWnd, PAINTSTRUCT& ps)
{
    HDC& hDC = ps.hdc;
    // Retrieve client area dimensions
    RECT rect;
    GetClientRect(hWnd, &rect);

    // Draw centered text "Grid32"
    SetTextColor(hDC, RGB(0, 0, 0));
    SetBkMode(hDC, TRANSPARENT);
    const TCHAR* text = _T("Grid32");
    DrawText(hDC, text, -1, &rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

}