#include "pch.h"
#include "Grid32Mgr.h"

CGrid32Mgr::CGrid32Mgr() : pRowInfoArray(nullptr), pColInfoArray(nullptr), 
m_hWndGrid(NULL), nRowHeaderHeight(50), nColHeaderWidth(50)
{
    memset(&gcs, 0, sizeof(GRIDCREATESTRUCT));
    memset(&m_cornerCell, 0, sizeof(m_cornerCell));
    memset(&m_defaultGridCell, 0, sizeof(m_cornerCell));
    m_defaultGridCell.clrBackground = RGB(255, 255, 255);
    m_defaultGridCell.m_clrBorderColor = 0;
    m_defaultGridCell.m_nBorderWidth = 1;
    m_defaultGridCell.penStyle = PS_DOT;
    m_defaultGridCell.fontInfo.m_wsFontFace = L"Arial";         // Font face
    m_defaultGridCell.fontInfo.m_fPointSize = 12.0f;            // Point size
    m_defaultGridCell.fontInfo.m_clrTextColor = RGB(0, 0, 0);   // Black text color
    m_defaultGridCell.fontInfo.bItalic = FALSE;                 // No italic
    m_defaultGridCell.fontInfo.bUnderline = FALSE;              // No underline
    m_defaultGridCell.fontInfo.bWeight = 400;                   // Normal weight (400)
    m_cornerCell.clrBackground = RGB(234, 234, 240);
    m_cornerCell.m_clrBorderColor = RGB(0, 0, 0);
    m_cornerCell.m_nBorderWidth = 1;
    m_cornerCell.penStyle = PS_SOLID;
    m_cornerCell.fontInfo.m_wsFontFace = L"Arial";         // Font face
    m_cornerCell.fontInfo.m_fPointSize = 11.0f;            // Point size
    m_cornerCell.fontInfo.m_clrTextColor = RGB(0, 0, 0);   // Black text color
    m_cornerCell.fontInfo.bItalic = FALSE;                 // No italic
    m_cornerCell.fontInfo.bUnderline = FALSE;              // No underline
    m_cornerCell.fontInfo.bWeight = 400;                   // Normal weight (400)
}


CGrid32Mgr::~CGrid32Mgr()
{
    DeleteAllCells();
    delete[] pRowInfoArray;
    delete[] pColInfoArray;
}

std::wstring CGrid32Mgr::GetColumnLabel(size_t index, size_t maxWidth)
{
    std::wstring label;
    while (index >= 0 && index < maxWidth)
    {
        label.insert(label.begin(), L'A' + (index % 26));
        index = index / 26 - 1;
    }
    return label;
}

bool CGrid32Mgr::Create(PGRIDCREATESTRUCT pGCS)
{
    // Store the provided GRIDCREATESTRUCT in the member variable
    gcs = *pGCS;

    if (gcs.nWidth == 0 || gcs.nHeight == 0)
        return false;
    // Perform any other initialization needed for your grid
    pColInfoArray = new COLINFO[gcs.nWidth];
    pRowInfoArray = new ROWINFO[gcs.nHeight];

    if (!pRowInfoArray || !pColInfoArray)
        return false;

    nColHeaderWidth = nRowHeaderHeight = (long)gcs.nDefRowHeight;

    for (size_t idx = 0; idx < gcs.nWidth; ++idx)
    {
        pColInfoArray[idx].nWidth = gcs.nDefColWidth;
        pColInfoArray[idx].clrBackground = m_cornerCell.clrBackground;
        pColInfoArray[idx].fontInfo = m_cornerCell.fontInfo;
        pColInfoArray[idx].m_clrBorderColor = m_cornerCell.m_clrBorderColor;
        pColInfoArray[idx].m_nBorderWidth = m_cornerCell.m_nBorderWidth;
        pColInfoArray[idx].penStyle = m_cornerCell.penStyle;

        // Check if the GS_SPREADSHEETHEADER bit is set
        if (gcs.style & GS_SPREADSHEETHEADER)
        {
            pColInfoArray[idx].m_wsName = GetColumnLabel(idx, gcs.nWidth);
        }
    }
    for (size_t idx = 0; idx < gcs.nHeight; ++idx)
    {
        pRowInfoArray[idx].nHeight = gcs.nDefRowHeight;
        pRowInfoArray[idx].clrBackground = m_cornerCell.clrBackground;
        pRowInfoArray[idx].fontInfo = m_cornerCell.fontInfo;
        pRowInfoArray[idx].m_clrBorderColor = m_cornerCell.m_clrBorderColor;
        pRowInfoArray[idx].m_nBorderWidth = m_cornerCell.m_nBorderWidth;
        pRowInfoArray[idx].penStyle = m_cornerCell.penStyle;

        if (gcs.style & GS_SPREADSHEETHEADER)
        {
            pRowInfoArray[idx].m_wsName = std::to_wstring(idx + 1);
        }
    }

    return true; // Return true if creation succeeded, false otherwise
}

void CGrid32Mgr::Paint(PAINTSTRUCT& ps)
{
    HDC& hDC = ps.hdc;
    // Retrieve client area dimensions
    RECT rect;
    GetClientRect(m_hWndGrid, &rect);

    DrawGrid(hDC, rect);
    DrawCells(hDC, rect);
    DrawHeader(hDC, rect);
}

void CGrid32Mgr::CalculateTotalGridRect()
{
    memset(&totalGridCellRect, 0, sizeof(RECT));
    for (size_t idx = 0; idx < gcs.nWidth; ++idx)
        totalGridCellRect.right += (long)pColInfoArray[idx].nWidth;
    for (size_t idx = 0; idx < gcs.nHeight; ++idx)
        totalGridCellRect.bottom += (long)pRowInfoArray[idx].nHeight;
}

void CGrid32Mgr::OffsetRectByScroll(RECT& rect)
{
    if ((gcs.style & GS_ROWHEADER) && (gcs.style & GS_COLHEADER))
    {
        OffsetRect(&rect, nColHeaderWidth, nRowHeaderHeight);
    }
}

bool CGrid32Mgr::CanDrawRect(RECT clientRect, RECT rect)
{
    // Check if the rectangles intersect at all
    return !(rect.left > clientRect.right ||
        rect.right < clientRect.left ||
        rect.top > clientRect.bottom ||
        rect.bottom < clientRect.top);
}

void CGrid32Mgr::DrawHeader(HDC hDC, const RECT &rect)
{
    // Check if gcs.style has both GS_ROWHEADER and GS_COLHEADER bits set
    CalculateTotalGridRect();
    if (gcs.style & GS_ROWHEADER)
    {
        DrawRowHeaders(hDC, rect);
    }
    if (gcs.style & GS_COLHEADER)
    {
        DrawColHeaders(hDC, rect);
    }
    if ((gcs.style & GS_ROWHEADER) && (gcs.style & GS_COLHEADER))
    {
        DrawCornerCell(hDC);
    }
}

void CGrid32Mgr::DrawCornerCell(HDC hDC)
{
    DrawHeaderButton(hDC, m_cornerCell.clrBackground, m_cornerCell.m_clrBorderColor, RECT{ 0, 0, (long)nColHeaderWidth, (long)nRowHeaderHeight }, true);
}

void CGrid32Mgr::DrawHeaderButton(HDC hDC, COLORREF background, COLORREF border, RECT coordinates, bool bRaised)
{
    // Create pens and brushes for drawing
    HPEN hBorderPen = CreatePen(PS_SOLID, 1, border);
    HBRUSH hBackgroundBrush = CreateSolidBrush(background);

    // Select the border pen and draw the rectangle border
    HPEN hOldPen = (HPEN)SelectObject(hDC, hBorderPen);
    HBRUSH hOldBrush = (HBRUSH)SelectObject(hDC, hBackgroundBrush);

    Rectangle(hDC, coordinates.left, coordinates.top, coordinates.right, coordinates.bottom);

    //// Fill the rectangle with the background color
    //FillRect(hDC, &coordinates, hBackgroundBrush);

    // Clean up GDI objects
    SelectObject(hDC, hOldPen);
    SelectObject(hDC, hOldBrush);
    DeleteObject(hBorderPen);
    DeleteObject(hBackgroundBrush);

    if (bRaised)
    {
        // Calculate highlight (average with white) and shadow (average with black) colors
        COLORREF highlight = RGB((GetRValue(background) + 255) / 2,
            (GetGValue(background) + 255) / 2,
            (GetBValue(background) + 255) / 2);

        COLORREF shadow = RGB(GetRValue(background) / 2,
            GetGValue(background) / 2,
            GetBValue(background) / 2);

        // Create pens for highlight and shadow
        HPEN hHighlightPen = CreatePen(PS_SOLID, 1, highlight);
        HPEN hShadowPen = CreatePen(PS_SOLID, 1, shadow);

        // Draw the top and left highlight
        HPEN hOldPen = (HPEN)SelectObject(hDC, hHighlightPen);
        // Top highlight
        MoveToEx(hDC, coordinates.left, coordinates.bottom - 2, NULL);
        LineTo(hDC, coordinates.left, coordinates.top);
        LineTo(hDC, coordinates.right - 2, coordinates.top);
        // Bottom right shadow
        SelectObject(hDC, hShadowPen);
        LineTo(hDC, coordinates.right - 2, coordinates.bottom - 2);
        LineTo(hDC, coordinates.left, coordinates.bottom - 2);

        // Draw the second highlight (top and left inner)
        SelectObject(hDC, hHighlightPen);
        MoveToEx(hDC, coordinates.left + 1, coordinates.bottom - 3, NULL);
        LineTo(hDC, coordinates.left + 1, coordinates.top + 1);
        LineTo(hDC, coordinates.right - 3, coordinates.top + 1);
        // Bottom right inner shadow
        SelectObject(hDC, hShadowPen);
        LineTo(hDC, coordinates.right - 3, coordinates.bottom - 3);
        LineTo(hDC, coordinates.left + 1, coordinates.bottom - 3);

        // Clean up GDI objects
        SelectObject(hDC, hOldPen);
        DeleteObject(hHighlightPen);
        DeleteObject(hShadowPen);
    }
}


void CGrid32Mgr::DrawColHeaders(HDC hDC, const RECT& rect)
{
    RECT colHeaderRect = totalGridCellRect;
    OffsetRectByScroll(colHeaderRect);
    colHeaderRect.right = colHeaderRect.left;
    colHeaderRect.top = 0;
    colHeaderRect.bottom = nRowHeaderHeight;

    for (size_t idx = 0; idx < gcs.nWidth; ++idx)
    {
        colHeaderRect.right += (long)pColInfoArray[idx].nWidth;

        if (CanDrawRect(rect, colHeaderRect))
        {
            DrawHeaderButton(hDC, pColInfoArray[idx].clrBackground, pColInfoArray[idx].m_clrBorderColor, colHeaderRect, true);

            // Draw text in the center
            std::wstring &text = pColInfoArray[idx].m_wsName;
            SetBkMode(hDC, TRANSPARENT);
            SetTextColor(hDC, pColInfoArray[idx].fontInfo.m_clrTextColor);

            HFONT hFont = CreateFontW(
                (int)((long double)pColInfoArray[idx].fontInfo.m_fPointSize * GetDeviceCaps(hDC, LOGPIXELSY) / 72.0), 0, 0, 0,
                pColInfoArray[idx].fontInfo.bWeight, pColInfoArray[idx].fontInfo.bItalic,
                pColInfoArray[idx].fontInfo.bUnderline, 0, 0, 0, 0, 0, 0,
                pColInfoArray[idx].fontInfo.m_wsFontFace.c_str()
            );

            HFONT hOldFont = (HFONT)SelectObject(hDC, hFont);

            DrawTextW(hDC, text.c_str(), -1, &colHeaderRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

            SelectObject(hDC, hOldFont);
            DeleteObject(hFont);

        }
        else if (rect.right < colHeaderRect.left)
            break;
        colHeaderRect.left = colHeaderRect.right;
    }
}

void CGrid32Mgr::DrawRowHeaders(HDC hDC, const RECT& rect)
{
    RECT rowHeaderRect = totalGridCellRect;
    OffsetRectByScroll(rowHeaderRect);
    rowHeaderRect.left = 0;
    rowHeaderRect.right = nColHeaderWidth;
    rowHeaderRect.bottom = rowHeaderRect.top;

    for (size_t idx = 0; idx < gcs.nWidth; ++idx)
    {
        rowHeaderRect.bottom += (long)pRowInfoArray[idx].nHeight;

        if (CanDrawRect(rect, rowHeaderRect))
        {
            DrawHeaderButton(hDC, pRowInfoArray[idx].clrBackground, pRowInfoArray[idx].m_clrBorderColor, rowHeaderRect, true);

            // Draw text in the center
            std::wstring& text = pRowInfoArray[idx].m_wsName;
            SetBkMode(hDC, TRANSPARENT);
            SetTextColor(hDC, pRowInfoArray[idx].fontInfo.m_clrTextColor);

            HFONT hFont = CreateFontW(
                (int)((long double)pRowInfoArray[idx].fontInfo.m_fPointSize * GetDeviceCaps(hDC, LOGPIXELSY) / 72.0), 0, 0, 0,
                pRowInfoArray[idx].fontInfo.bWeight, pRowInfoArray[idx].fontInfo.bItalic,
                pRowInfoArray[idx].fontInfo.bUnderline, 0, 0, 0, 0, 0, 0,
                pRowInfoArray[idx].fontInfo.m_wsFontFace.c_str()
            );

            HFONT hOldFont = (HFONT)SelectObject(hDC, hFont);

            DrawTextW(hDC, text.c_str(), -1, &rowHeaderRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

            SelectObject(hDC, hOldFont);
            DeleteObject(hFont);

        }
        else if (rect.bottom < rowHeaderRect.top)
            break;
        rowHeaderRect.top = rowHeaderRect.bottom;
    }
}

void CGrid32Mgr::DrawGrid(HDC hDC, const RECT& clientRect)
{
    RECT gridCellRect = totalGridCellRect;
    OffsetRectByScroll(gridCellRect);

    gridCellRect.right = gridCellRect.left;
    gridCellRect.bottom = gridCellRect.top;

    HPEN hBorderPen = CreatePen(m_defaultGridCell.penStyle, m_defaultGridCell.m_nBorderWidth, m_defaultGridCell.m_clrBorderColor);
    HPEN hOldPen = (HPEN)SelectObject(hDC, hBorderPen);
    
    for (size_t col = 0; col < gcs.nWidth; ++col)
    {
        gridCellRect.right += (long)pColInfoArray[col].nWidth;
        for (size_t row = 0; row < gcs.nHeight; ++row)
        {
            gridCellRect.bottom += (long)pRowInfoArray[row].nHeight;
            // Draw the cell background
            HBRUSH hBackgroundBrush = CreateSolidBrush(m_defaultGridCell.clrBackground);
            FillRect(hDC, &gridCellRect, hBackgroundBrush);
            DeleteObject(hBackgroundBrush);

            MoveToEx(hDC, 0, gridCellRect.top - 1, NULL);
            LineTo(hDC, clientRect.right, gridCellRect.top - 1);

            gridCellRect.top = gridCellRect.bottom;

            if (clientRect.bottom < gridCellRect.top)
                break;
        }

        MoveToEx(hDC, gridCellRect.left - 1, 0, NULL);
        LineTo(hDC, gridCellRect.left - 1, clientRect.bottom);

        gridCellRect.left = gridCellRect.right;
        if (clientRect.right < gridCellRect.left)
            break;

    }
    SelectObject(hDC, hOldPen);
    DeleteObject(hBorderPen);
}

void CGrid32Mgr::DrawCells(HDC hDC, const RECT& clientRect)
{
    RECT gridCellRect = totalGridCellRect;
    OffsetRectByScroll(gridCellRect);

    gridCellRect.right = gridCellRect.left;
    gridCellRect.bottom = gridCellRect.top;

    for (size_t col = 0; col < gcs.nWidth; ++col)
    {
        gridCellRect.right += (long)pColInfoArray[col].nWidth;
        for (size_t row = 0; row < gcs.nHeight; ++row)
        {
            PGRIDCELL pCell = GetCell((UINT)row, (UINT)col);

            gridCellRect.bottom += (long)pRowInfoArray[row].nHeight;

            if (CanDrawRect(clientRect, gridCellRect) && pCell != &m_defaultGridCell)
            {
                // Draw the cell base
                HBRUSH hBackgroundBrush = CreateSolidBrush(pCell->clrBackground);
                HPEN hBorderPen = CreatePen(pCell->penStyle, pCell->m_nBorderWidth, pCell->m_clrBorderColor);
                HPEN hOldPen = (HPEN)SelectObject(hDC, hBorderPen);
                HBRUSH hOldBrush = (HBRUSH)SelectObject(hDC, hBackgroundBrush);
                Rectangle(hDC, gridCellRect.left, gridCellRect.top, gridCellRect.right, gridCellRect.bottom);
                SelectObject(hDC, hOldPen);
                SelectObject(hDC, hOldBrush);
                DeleteObject(hBorderPen);
                DeleteObject(hBackgroundBrush);

                // Draw the cell text
                SetBkMode(hDC, TRANSPARENT);
                SetTextColor(hDC, pCell->fontInfo.m_clrTextColor);

                HFONT hFont = CreateFontW(
                    (int)pCell->fontInfo.m_fPointSize, 0, 0, 0,
                    pCell->fontInfo.bWeight, pCell->fontInfo.bItalic,
                    pCell->fontInfo.bUnderline, 0, 0, 0, 0, 0, 0,
                    pCell->fontInfo.m_wsFontFace.c_str()
                );

                HFONT hOldFont = (HFONT)SelectObject(hDC, hFont);

                DrawTextW(hDC, pCell->m_wsText.c_str(), -1, &gridCellRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

                SelectObject(hDC, hOldFont);
                DeleteObject(hFont);
            }
            else if (clientRect.right < gridCellRect.left || clientRect.bottom < gridCellRect.top)
                break;

            gridCellRect.top = gridCellRect.bottom;
        }

        gridCellRect.left = gridCellRect.right;
    }
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

//Unused function, but it keeps asking me to create it and says "Create" needs to create it, so...
void CGrid32Mgr::DrawButton(HDC hDC, COLORREF background, COLORREF border, RECT coordinates, bool bRaised)
{
    UNREFERENCED_PARAMETER(hDC);
    UNREFERENCED_PARAMETER(background);
    UNREFERENCED_PARAMETER(border);
    UNREFERENCED_PARAMETER(coordinates);
    UNREFERENCED_PARAMETER(bRaised);
}