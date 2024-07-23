#include "pch.h"
#include "Grid32Mgr.h"

CGrid32Mgr::CGrid32Mgr()
{
    memset(&gcs, 0, sizeof(GRIDCREATESTRUCT));
}


CGrid32Mgr::~CGrid32Mgr()
{
    DeleteAllCells();
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

PGRIDCELL CGrid32Mgr::GetCell(UINT nRow, UINT nCol)
{
    auto it = mapCells.find(std::make_pair(nRow, nCol));
    if (it != mapCells.end())
    {
        return it->second;
    }
    return &m_defaultGridCell;
}

void CGrid32Mgr::SetCell(UINT nRow, UINT nCol, const GRIDCELL& gc)
{
    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);

    // If the cell exists, delete the old one
    if (it != mapCells.end())
    {
        delete it->second;
        it->second = new GRIDCELL(gc);
    }
    else
    {
        mapCells[key] = new GRIDCELL(gc);
    }
}

void CGrid32Mgr::DeleteCell(UINT nRow, UINT nCol)
{
    auto it = mapCells.find(std::make_pair(nRow, nCol));
    if (it != mapCells.end())
    {
        delete it->second;
        mapCells.erase(it);
    }
}

void CGrid32Mgr::DeleteAllCells()
{
    for (auto& cellPair : mapCells)
    {
        delete cellPair.second;
    }
    mapCells.clear();
}