#include "pch.h"
#include "Grid32Mgr.h"

CGrid32Mgr::CGrid32Mgr() : pRowInfoArray(nullptr), pColInfoArray(nullptr), 
m_hWndGrid(NULL), nRowHeaderHeight(40), nColHeaderWidth(70)
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
    m_selectionPoint = m_visibleTopLeft = { 0,0 };
    m_scrollDifference = { 0, 0 };
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

    //nColHeaderWidth = nRowHeaderHeight = (long)gcs.nDefRowHeight;

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
    HDC hDC = ps.hdc;

    // Retrieve client area dimensions
    RECT rect;
    GetClientRect(m_hWndGrid, &rect);
    int width = rect.right - rect.left;
    int height = rect.bottom - rect.top;

    // Create a memory DC compatible with the screen DC
    HDC memDC = CreateCompatibleDC(hDC);

    // Create a bitmap compatible with the screen DC
    HBITMAP memBitmap = CreateCompatibleBitmap(hDC, width, height);

    // Select the bitmap into the memory DC
    HBITMAP oldBitmap = (HBITMAP)SelectObject(memDC, memBitmap);

    // Fill the background with white (or any desired background color)
    HBRUSH hBrush = CreateSolidBrush(m_defaultGridCell.clrBackground);
    FillRect(memDC, &rect, hBrush);
    DeleteObject(hBrush);

    // Perform all drawing operations on the memory DC
    DrawGrid(memDC, rect);
    DrawCells(memDC, rect);
    DrawSelectionBox(memDC, rect);
    DrawHeader(memDC, rect);
    DrawVoidSpace(memDC, rect);

    // Copy the contents of the memory DC to the screen
    BitBlt(hDC, 0, 0, width, height, memDC, 0, 0, SRCCOPY);

    // Clean up
    SelectObject(memDC, oldBitmap);
    DeleteObject(memBitmap);
    DeleteDC(memDC);
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
    OffsetRect(&rect, 0-m_scrollDifference.x, 0 - m_scrollDifference.y);

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

void CGrid32Mgr::DrawSelectionBox(HDC hDC, const RECT& clientRect)
{
    if (!(gcs.style & GS_HIGHLIGHTSELECTION))
        return;

    RECT gridCellRect = totalGridCellRect;
    OffsetRectByScroll(gridCellRect);

    gridCellRect.right = gridCellRect.left;
    gridCellRect.bottom = gridCellRect.top;

    for (size_t col = 0; col < gcs.nWidth; ++col)
    {
        gridCellRect.right += (long)pColInfoArray[col].nWidth;
        if (col >= m_selectionPoint.nCol)
        {
            for (size_t row = 0; row < gcs.nHeight; ++row)
            {
                gridCellRect.bottom += (long)pRowInfoArray[row].nHeight;
                if (row >= m_selectionPoint.nRow)
                {
                    if (CanDrawRect(clientRect, gridCellRect))
                    {
                        // Create a pen with the specified color and a width of 3
                        HPEN hPen = CreatePen(PS_SOLID, 3, gcs.clrSelectBox);
                        HPEN hOldPen = (HPEN)SelectObject(hDC, hPen);

                        // Use a NULL brush to only draw the border
                        HBRUSH hOldBrush = (HBRUSH)SelectObject(hDC, GetStockObject(NULL_BRUSH));

                        // Draw the rectangle
                        Rectangle(hDC, gridCellRect.left, gridCellRect.top, gridCellRect.right, gridCellRect.bottom);

                        // Restore the old pen and brush
                        SelectObject(hDC, hOldPen);
                        SelectObject(hDC, hOldBrush);

                        // Delete the created pen
                        DeleteObject(hPen);

                        return;
                    }
                    else if (clientRect.right < gridCellRect.left || clientRect.bottom < gridCellRect.top)
                        break;
                }
                gridCellRect.top = gridCellRect.bottom;
                if (clientRect.bottom < gridCellRect.top)
                    break;
            }
        }
        gridCellRect.left = gridCellRect.right;
        if (clientRect.right < gridCellRect.left)
            break;
    }
}



void CGrid32Mgr::DrawHeader(HDC hDC, const RECT &rect)
{
    // Check if gcs.style has both GS_ROWHEADER and GS_COLHEADER bits set
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
        COLORREF highlight = RGB(GetRValue(background) <= 192 ? (GetRValue(background) + 255) / 2 : 255,
            GetGValue(background) <= 192 ? (GetGValue(background) + 255) / 2 : 255,
            GetGValue(background) <= 192 ? (GetBValue(background) + 255) / 2 : 255);

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

    for (size_t idx = 0; idx < gcs.nHeight; ++idx)
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
            else if (clientRect.bottom < gridCellRect.top)
                break;

            gridCellRect.top = gridCellRect.bottom;
        }

        gridCellRect.left = gridCellRect.right;
        if (clientRect.right < gridCellRect.left)
            break;
    }
}

void CGrid32Mgr::DrawVoidSpace(HDC hDC, const RECT& rect)
{
    RECT clientRect;
    GetClientRect(m_hWndGrid, &clientRect);
    bool bDrawBorder = false;
    long nMaxWidth = static_cast<long>(CalculatedColumnDistance(m_visibleTopLeft.nCol, gcs.nWidth));
    long nMaxHeight = static_cast<long>(CalculatedRowDistance(m_visibleTopLeft.nRow, gcs.nHeight));
    if (nMaxWidth < clientRect.right)
    {
        HBRUSH hBrush = CreateSolidBrush(RGB(128, 128, 128));
        HPEN hPen = CreatePen(PS_SOLID, 1, RGB(128, 128, 128));

        HPEN oldPen = (HPEN)SelectObject(hDC, hPen);
        HBRUSH oldBrush = (HBRUSH)SelectObject(hDC, hBrush);

        Rectangle(hDC, nMaxWidth + nColHeaderWidth, clientRect.top, clientRect.right, clientRect.bottom);

        DeleteObject(SelectObject(hDC, oldPen));
        DeleteObject(SelectObject(hDC, oldBrush));

        bDrawBorder = true;
    }
    if (nMaxHeight < clientRect.bottom)
    {
        HBRUSH hBrush = CreateSolidBrush(RGB(128, 128, 128));
        HPEN hPen = CreatePen(PS_SOLID, 1, RGB(128, 128, 128));

        HPEN oldPen = (HPEN)SelectObject(hDC, hPen);
        HBRUSH oldBrush = (HBRUSH)SelectObject(hDC, hBrush);

        Rectangle(hDC, 0, nMaxHeight + nRowHeaderHeight, clientRect.right, clientRect.bottom);

        DeleteObject(SelectObject(hDC, oldPen));
        DeleteObject(SelectObject(hDC, oldBrush));

        bDrawBorder = true;
    }

    if (bDrawBorder)
    {
        int x = nMaxWidth < clientRect.right ? nMaxWidth : clientRect.right;
        int y = nMaxHeight < clientRect.bottom ? nMaxHeight : clientRect.bottom;

        x += nColHeaderWidth + 1;
        y += nRowHeaderHeight + 1;

        HPEN hPen = CreatePen(PS_SOLID, 3, 0);
        HPEN oldPen = (HPEN)SelectObject(hDC, hPen);

        if (x < clientRect.right)
        {
            MoveToEx(hDC, x, 0, NULL);
            LineTo(hDC, x, y < clientRect.bottom ? y : clientRect.bottom);
        }

        if (y < clientRect.bottom)
        {
            MoveToEx(hDC, 0, y, NULL);
            LineTo(hDC, x < clientRect.right ? x : clientRect.right, y);
        }


        DeleteObject(SelectObject(hDC, oldPen));
    }
}

size_t CGrid32Mgr::CalculatedColumnDistance(size_t start, size_t end)
{
    size_t nWidth = 0;
    if (start >= gcs.nWidth || end > gcs.nWidth)
        return 0;
    for (size_t idx = start; idx < end; ++idx)
        nWidth += pColInfoArray[idx].nWidth;

    return nWidth;
}

size_t CGrid32Mgr::CalculatedRowDistance(size_t start, size_t end)
{
    size_t nHeight = 0;
    if (start >= gcs.nHeight || end > gcs.nHeight)
        return 0;
    for (size_t idx = start; idx < end; ++idx)
        nHeight += pRowInfoArray[idx].nHeight;

    return nHeight;
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

void CGrid32Mgr::OnHScroll(UINT nSBCode, UINT nPos, HWND hScrollBar)
{
    switch (nSBCode)
    {
    case SB_LINELEFT:
        IncrementSelectedCell(-1, 0, INCREMENT_SINGLE);
        break;
    case SB_LINERIGHT:
        IncrementSelectedCell(1, 0, INCREMENT_SINGLE);
        break;
    case SB_PAGELEFT:
        // Handle left page scroll
        break;
    case SB_PAGERIGHT:
        // Handle right page scroll
        break;
    case SB_THUMBTRACK:
        m_visibleTopLeft.nCol = (LONG)nPos;
        break;
    case SB_THUMBPOSITION:
        m_visibleTopLeft.nCol = (LONG)nPos;
        SetScrollPos(m_hWndGrid, SB_HORZ, nPos, TRUE);
        break;
    case SB_LEFT:
        m_scrollDifference.x = 0;
        break;
    case SB_RIGHT:
        // Handle scroll to far right
        break;
    case SB_ENDSCROLL:
        // Handle end of scroll action
        break;
    default:
        break;
    }

    InvalidateRect(m_hWndGrid, NULL, true);
}

void CGrid32Mgr::OnVScroll(UINT nSBCode, UINT nPos, HWND hScrollBar)
{
    switch (nSBCode)
    {
    case SB_LINEUP:
        IncrementSelectedCell(0, -1, INCREMENT_SINGLE);
        break;
    case SB_LINEDOWN:
        IncrementSelectedCell(0, 1, INCREMENT_SINGLE);
        break;
    case SB_PAGEUP:
        // Handle page scroll up
        break;
    case SB_PAGEDOWN:
        // Handle page scroll down
        break;
    case SB_THUMBTRACK:
        m_visibleTopLeft.nCol = (LONG)nPos;
        break;
    case SB_THUMBPOSITION:
        m_visibleTopLeft.nCol = (LONG)nPos;
        break;
    case SB_TOP:
        m_visibleTopLeft.nCol = 0;
        break;
    case SB_BOTTOM:
        // Handle scroll to bottom
        break;
    case SB_ENDSCROLL:
        // Handle end of scroll action
        break;
    default:
        break;
    }
    InvalidateRect(m_hWndGrid, NULL, false);
}


void CGrid32Mgr::SetScrollRanges()
{
    if (gcs.style & WS_HSCROLL)
    {
        // Set the horizontal scroll range
        SCROLLINFO siH = { sizeof(SCROLLINFO) };
        siH.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
        siH.nMin = 0;
        siH.nMax = static_cast<int>(gcs.nWidth); // Assuming each column is of uniform width
        siH.nPage = 25; // Page size can be the width of one column or more
        siH.nPos = 0;
        SetScrollInfo(m_hWndGrid, SB_HORZ, &siH, TRUE);
    }
    if (gcs.style & WS_VSCROLL)
    {
        // Set the vertical scroll range
        SCROLLINFO siV = { sizeof(SCROLLINFO) };
        siV.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
        siV.nMin = 0;
        siV.nMax = static_cast<int>(gcs.nHeight); // Assuming each row is of uniform height
        siV.nPage = 25; // Page size can be the height of one row or more
        siV.nPos = 0;
        SetScrollInfo(m_hWndGrid, SB_VERT, &siV, TRUE);
    }
}

void CGrid32Mgr::IncrementSelectedCell(long nRow, long nCol, short nWhich)
{
    // Increment or decrement the selection point
    UINT absRow = (UINT)abs(nRow), absCol = (UINT)abs(nCol);
    GRIDPOINT oldVisibleTopLeft = m_visibleTopLeft;

    if (nRow < 0)
    {
        if (absRow >= m_selectionPoint.nRow)
            m_selectionPoint.nRow = 0;
        else
            m_selectionPoint.nRow -= absRow;
    }
    else
    {
        m_selectionPoint.nRow += absRow;
        if (m_selectionPoint.nRow >= (UINT)gcs.nHeight)
            m_selectionPoint.nRow = (UINT)gcs.nHeight - 1;
    }

    if (nCol < 0)
    {
        if (absCol >= m_selectionPoint.nCol)
            m_selectionPoint.nCol = 0;
        else
            m_selectionPoint.nCol -= absCol;
    }
    else
    {
        m_selectionPoint.nCol += absCol;
        if (m_selectionPoint.nCol >= (UINT)gcs.nWidth)
            m_selectionPoint.nCol = (UINT)gcs.nWidth - 1;
    }

    // Get the dimensions of the client area
    RECT clientRect, cell = { nColHeaderWidth, nRowHeaderHeight, 0, 0 };
    GetClientRect(m_hWndGrid, &clientRect);

    // Calculate the total height and width of the visible grid area

    if (m_visibleTopLeft.nRow <= m_selectionPoint.nRow)
    {
        for (size_t i = m_visibleTopLeft.nRow; i < m_selectionPoint.nRow; ++i)
        {
            cell.top += static_cast<int>(pRowInfoArray[i].nHeight);
        }
    }
    else
    {
        int nHeight = 0;
        for (size_t i = m_selectionPoint.nRow; i != ~0 && i >= 0; --i)
        {
            if (i == 0)
            {
                m_visibleTopLeft.nRow = 0;
                cell.top = -100;
                break;
            }
            if (nHeight + static_cast<int>(pRowInfoArray[i].nHeight) > clientRect.bottom)
            {
                m_visibleTopLeft.nRow = i;
                cell.top -= nHeight;
                break;
            }
            nHeight += static_cast<int>(pRowInfoArray[i].nHeight);
        }
    }

    if (m_visibleTopLeft.nCol <= m_selectionPoint.nCol)
    {
        for (size_t i = m_visibleTopLeft.nCol; i < m_selectionPoint.nCol; ++i)
        {
            cell.left += static_cast<int>(pColInfoArray[i].nWidth);
        }
    }
    else
    {
        int nWidth = 0;
        for (size_t i = m_selectionPoint.nCol; i != ~0 && i >= 0; --i)
        {
            if (nWidth + static_cast<int>(pColInfoArray[i].nWidth) > clientRect.right)
            {
                m_visibleTopLeft.nCol = i;
                break;
            }
            nWidth += static_cast<int>(pColInfoArray[i].nWidth);
            if (i == 0)
            {
                m_visibleTopLeft.nCol = 0;
                break;
            }
        }
        cell.left -= nWidth;
    }

    cell.bottom = cell.top + static_cast<int>(pRowInfoArray[m_selectionPoint.nRow].nHeight);
    cell.right = cell.left + static_cast<int>(pColInfoArray[m_selectionPoint.nCol].nWidth);

    // Determine the scroll direction based on visibility
    int scrollFlags = 0;

    if (cell.top < clientRect.top || cell.bottom > clientRect.bottom)
        scrollFlags |= SCROLL_VERT;

    if (cell.left < clientRect.left || cell.right > clientRect.right)
        scrollFlags |= SCROLL_HORZ;

    if (scrollFlags != 0)
    {
        // Scroll to make the selected cell visible
        ScrollToCell(m_selectionPoint.nRow, m_selectionPoint.nCol, oldVisibleTopLeft, scrollFlags);
    }

    // Force the grid to redraw
    InvalidateRect(m_hWndGrid, NULL, false);
}

void CGrid32Mgr::ScrollToCell(size_t row, size_t col, GRIDPOINT& oldVisibleTopLeft, int scrollFlags)
{
    int scrollY = 0, scrollX = 0;
    int newTop = 0, newLeft = 0;
    RECT clientRect;
    GetClientRect(m_hWndGrid, &clientRect);

    if (scrollFlags & SCROLL_VERT)
    {
        for (size_t i = m_visibleTopLeft.nRow; i < row; ++i)
        {
            scrollY += static_cast<int>(pRowInfoArray[i].nHeight);
        }
        SetScrollPos(m_hWndGrid, SB_VERT, row, TRUE);

        // Update the top row index
        if (row > m_visibleTopLeft.nRow)
            m_scrollDifference.y += scrollY;
        else
            m_scrollDifference.y -= scrollY;

        newTop = static_cast<int>(row);
    }
    else
    {
        // Use current visible top row if no vertical scrolling
        newTop = m_visibleTopLeft.nRow;
    }

    if (scrollFlags & SCROLL_HORZ)
    {
        if (oldVisibleTopLeft.nCol)
        {

        }
         
        if (col > m_visibleTopLeft.nCol && m_visibleTopLeft.nCol >= oldVisibleTopLeft.nCol)
        {
            for (size_t i = m_visibleTopLeft.nCol; i < col; ++i)
            {
                scrollX += static_cast<int>(pColInfoArray[i].nWidth);
            }
            m_scrollDifference.x += scrollX;
            // Update the left column index
            newLeft = static_cast<int>(col);
        }
        else if (m_visibleTopLeft.nCol == 0)
        {
            m_scrollDifference.x = 0;
        }
        else
        {
            //for(size_t i = m_visibleTopLeft.nCol; i <= m_selectionPoint.nCol; ++i)
            for(size_t i = m_selectionPoint.nCol; i != ~0 && i >= m_visibleTopLeft.nCol; --i)
            {
                if (scrollX + static_cast<int>(pColInfoArray[i].nWidth) > clientRect.right)
                {
                    m_visibleTopLeft.nCol = i+1;
                    break;
                }
                scrollX += static_cast<int>(pColInfoArray[i].nWidth);

            }

            m_scrollDifference.x -= scrollX;
            // Update the left column index
            newLeft = m_visibleTopLeft.nCol;

        }
        SetScrollPos(m_hWndGrid, SB_HORZ, col, TRUE);
        if (m_scrollDifference.x < 0)
            m_scrollDifference.x = 0;
    }
    else
    {
        // Use current visible left column if no horizontal scrolling
        newLeft = m_visibleTopLeft.nCol;
    }

    // Update visible top-left position
    m_visibleTopLeft.nRow = newTop;
    m_visibleTopLeft.nCol = newLeft;

}

void CGrid32Mgr::OnMove(int x, int y)
{
    UNREFERENCED_PARAMETER(x);
    UNREFERENCED_PARAMETER(y);
}

void CGrid32Mgr::OnSize(UINT nType, int cx, int cy)
{
    UNREFERENCED_PARAMETER(nType);
    UNREFERENCED_PARAMETER(cx);
    UNREFERENCED_PARAMETER(cy);
    CalculateTotalGridRect();
    SetScrollRanges();
}

bool CGrid32Mgr::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
    switch (nChar)
    {
    case VK_LEFT:
        IncrementSelectedCell(0, -1, INCREMENT_SINGLE);
        return true;

    case VK_RIGHT:
        IncrementSelectedCell(0, 1, INCREMENT_SINGLE);
        return true;

    case VK_UP:
        IncrementSelectedCell(-1, 0, INCREMENT_SINGLE);
        return true;

    case VK_DOWN:
        IncrementSelectedCell(1, 0, INCREMENT_SINGLE);
        return true;

    default:
        // Handle other keys
        break;
    }

    return false;
}
