#include "pch.h"
#include <CommCtrl.h>
#include "Grid32Mgr.h"


const std::set<char> allowedPunctuation = {
'!', '@', '#', '$', '%', '^', '&', '*', '(', ')',
'-', '_', '=', '+', '[', ']', '{', '}', '\\', '|',
';', ':', '\'', '\"', ',', '.', '/', '<', '>', '?',
'`', '~'
};

CGrid32Mgr::CGrid32Mgr() : pRowInfoArray(nullptr), pColInfoArray(nullptr),
m_hWndGrid(NULL), nColHeaderHeight(40), nRowHeaderWidth(70), m_editWndProc(nullptr),
m_nMouseHoverDelay(400), m_npHoverDelaySet(0), m_nToBeSized(~0),
m_rgbSizingLine(RGB(0, 0, 64)), m_nSizingLine(0), m_bResizable(true), m_bUndoRecordEnabled(TRUE),
m_bSizing(false), m_bSelecting(false), m_hDefaultFont(NULL),
m_clientRect{ 0, 0, 0, 0 }, m_gdiplusToken(0), m_gridHitTest(0),
m_hWndEdit(NULL), m_lastClickTime(0), totalGridCellRect{ 0, 0, 0, 0 }
{
    memset(&gcs, 0, sizeof(GRIDCREATESTRUCT));
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
    m_defaultGridCell.justification = DT_CENTER | DT_VCENTER;
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
    m_cornerCell.justification = DT_CENTER | DT_VCENTER;
    m_currentCell = m_visibleTopLeft = m_HoverCell = { 0,0 };
    m_scrollDifference = { 0, 0 };
    visibleGrid = { 0,0 };
    m_bRedraw = TRUE;
    dwError = 0;
    memset(&m_selectionRect, 0, sizeof(GRIDSELECTION));
    m_bSelecting = FALSE;
    m_mouseDraggingStartPoint = { 0, 0 };
}


CGrid32Mgr::~CGrid32Mgr()
{
    DeleteAllCells();
    delete[] pRowInfoArray;
    delete[] pColInfoArray;
}


bool IsRectInside(const RECT& r1, const RECT& r2) {
    return (r1.left >= r2.left &&
        r1.right <= r2.right &&
        r1.top >= r2.top &&
        r1.bottom <= r2.bottom);
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
    {
        SetLastError(GRID_ERROR_OUT_OF_MEMORY);
        SendGridNotification(NM_OUTOFMEMORY);
        return false;
    }


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

    // **Edit Control Initialization (stub):**
    m_hWndEdit = CreateWindowEx(0, _T("EDIT"), _T(""), WS_CHILD | ES_AUTOHSCROLL,
        0, 0, 100, 20, m_hWndGrid, (HMENU)ID_CELL_EDIT, GetModuleHandle(NULL), NULL);

    // **Stub for edit control WndProc:**
    // Replace this with your actual WndProc implementation for the edit control
    m_editWndProc = (WNDPROC)SetWindowLongPtr(m_hWndEdit, GWLP_WNDPROC,
        (LONG_PTR)EditCtrl_WndProc);

    if (!m_editWndProc) {
        return false; // Fail if subclassing edit control fails
    }

    SetLastError(0);

    GdiplusStartupInput gdiplusStartupInput;
    if (GdiplusStartup(&m_gdiplusToken, &gdiplusStartupInput, NULL) != Ok) {
        return false; // Initialization failed
    }

    HDC hDC = GetDC(m_hWndGrid);
    m_hDefaultFont = CreateFontW(
        (int)(m_defaultGridCell.fontInfo.m_fPointSize * GetDeviceCaps(hDC, LOGPIXELSY)) / -72, 0, 0, 0,
        m_defaultGridCell.fontInfo.bWeight, m_defaultGridCell.fontInfo.bItalic,
        m_defaultGridCell.fontInfo.bUnderline, 0, 0, 0, 0, 0, 0,
        m_defaultGridCell.fontInfo.m_wsFontFace.c_str()
    );

    ReleaseDC(m_hWndGrid, hDC);
    SendMessage(m_hWndEdit, WM_SETFONT, (WPARAM)m_hDefaultFont, 0);

    return true; // Return true if creation succeeded, false otherwise
}

void CGrid32Mgr::OnDestroy()
{
    SendMessage(m_hWndEdit, WM_SETFONT, 0, 0);
    DeleteObject(m_hDefaultFont);
    m_hDefaultFont = NULL;
    GdiplusShutdown(m_gdiplusToken);
}

void CGrid32Mgr::OnPaint(PAINTSTRUCT& ps)
{
    HDC hDC = ps.hdc;

    // Retrieve client area dimensions
    int width = m_clientRect.right - m_clientRect.left;
    int height = m_clientRect.bottom - m_clientRect.top;

    // Create a memory DC compatible with the screen DC
    HDC memDC = CreateCompatibleDC(hDC);

    // Create a bitmap compatible with the screen DC
    HBITMAP memBitmap = CreateCompatibleBitmap(hDC, width, height);

    // Select the bitmap into the memory DC
    HBITMAP oldBitmap = (HBITMAP)SelectObject(memDC, memBitmap);

    // Fill the background with white (or any desired background color)
    HBRUSH hBrush = CreateSolidBrush(m_defaultGridCell.clrBackground);
    FillRect(memDC, &m_clientRect, hBrush);
    DeleteObject(hBrush);

    // Perform all drawing operations on the memory DC
    DrawGrid(memDC);
    DrawCells(memDC);
    DrawSelectionSurround(memDC); 
    DrawCurrentCellBox(memDC);
    DrawHeader(memDC);
    DrawVoidSpace(memDC);
    DrawSizingLine(memDC);

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
        OffsetRect(&rect, GetActualRowHeaderWidth(), GetActualColHeaderHeight());
    }
}

bool CGrid32Mgr::CanDrawRect(RECT rect)
{
    // Check if the rectangles intersect at all
    return !(rect.left > m_clientRect.right ||
        rect.right < m_clientRect.left ||
        rect.top > m_clientRect.bottom ||
        rect.bottom < m_clientRect.top);
}

void CGrid32Mgr::DrawCurrentCellBox(HDC hDC)
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
        if (col >= m_currentCell.nCol)
        {
            for (size_t row = 0; row < gcs.nHeight; ++row)
            {
                gridCellRect.bottom += (long)pRowInfoArray[row].nHeight;
                if (row >= m_currentCell.nRow)
                {
                    if (CanDrawRect(gridCellRect))
                    {
                        // Create a pen with the specified color and a width of 3
                        HPEN hPen = CreatePen(PS_SOLID, 3, gcs.clrCurrentCellHighlightBox);
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
                    else if (m_clientRect.right < gridCellRect.left || m_clientRect.bottom < gridCellRect.top)
                        break;
                }
                gridCellRect.top = gridCellRect.bottom;
                if (m_clientRect.bottom < gridCellRect.top)
                    break;
            }
        }
        gridCellRect.left = gridCellRect.right;
        if (m_clientRect.right < gridCellRect.left)
            break;
    }
}



void CGrid32Mgr::DrawHeader(HDC hDC)
{
    // Check if gcs.style has both GS_ROWHEADER and GS_COLHEADER bits set
    if (gcs.style & GS_ROWHEADER)
    {
        DrawRowHeaders(hDC);
    }
    if (gcs.style & GS_COLHEADER)
    {
        DrawColHeaders(hDC);
    }
    if ((gcs.style & GS_ROWHEADER) && (gcs.style & GS_COLHEADER))
    {
        DrawCornerCell(hDC);
    }
}

void CGrid32Mgr::DrawCornerCell(HDC hDC)
{
    DrawHeaderButton(hDC, m_cornerCell.clrBackground, m_cornerCell.m_clrBorderColor, RECT{ 0, 0, (long)nRowHeaderWidth, (long)nColHeaderHeight }, true);
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


void CGrid32Mgr::DrawColHeaders(HDC hDC)
{
    RECT colHeaderRect = totalGridCellRect;
    OffsetRectByScroll(colHeaderRect);
    colHeaderRect.right = colHeaderRect.left;
    colHeaderRect.top = 0;
    colHeaderRect.bottom = GetActualColHeaderHeight();

    for (size_t idx = 0; idx < gcs.nWidth; ++idx)
    {
        colHeaderRect.right += (long)pColInfoArray[idx].nWidth;

        if (CanDrawRect(colHeaderRect))
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
        else if (m_clientRect.right < colHeaderRect.left)
            break;
        colHeaderRect.left = colHeaderRect.right;
    }
}

void CGrid32Mgr::DrawRowHeaders(HDC hDC)
{
    RECT rowHeaderRect = totalGridCellRect;
    OffsetRectByScroll(rowHeaderRect);
    rowHeaderRect.left = 0;
    rowHeaderRect.right = GetActualRowHeaderWidth();
    rowHeaderRect.bottom = rowHeaderRect.top;

    for (size_t idx = 0; idx < gcs.nHeight; ++idx)
    {
        rowHeaderRect.bottom += (long)pRowInfoArray[idx].nHeight;

        if (CanDrawRect(rowHeaderRect))
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
        else if (m_clientRect.bottom < rowHeaderRect.top)
            break;
        rowHeaderRect.top = rowHeaderRect.bottom;
    }
}

void CGrid32Mgr::DrawGrid(HDC hDC)
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
            LineTo(hDC, m_clientRect.right, gridCellRect.top - 1);

            gridCellRect.top = gridCellRect.bottom;

            if (m_clientRect.bottom < gridCellRect.top)
                break;
        }

        MoveToEx(hDC, gridCellRect.left - 1, 0, NULL);
        LineTo(hDC, gridCellRect.left - 1, m_clientRect.bottom);

        gridCellRect.left = gridCellRect.right;
        if (m_clientRect.right < gridCellRect.left)
            break;

    }
    SelectObject(hDC, hOldPen);
    DeleteObject(hBorderPen);
}

void CGrid32Mgr::DrawSelectionOverlay(HDC hDC, RECT &gridCellRect)
{
    BYTE alpha = 255 / 5;
    DrawSemiTransparentRectangle(hDC, gridCellRect, gcs.clrSelectionOverlay, alpha);
}


void CGrid32Mgr::DrawCells(HDC hDC)
{
    RECT gridCellRect = { 0, 0, 0, 0 };

    for (size_t col = m_visibleTopLeft.nCol;
        col < gcs.nWidth && gridCellRect.left <= m_clientRect.right; ++col)
    {
        gridCellRect.left = (long)GetColumnStartPos((UINT)col);
        gridCellRect.right = gridCellRect.left + (long)pColInfoArray[col].nWidth;
        for (size_t row = m_visibleTopLeft.nRow;
            row < gcs.nHeight && gridCellRect.top <= m_clientRect.bottom; ++row)
        {
            gridCellRect.top = (long)GetRowStartPos((UINT)row);
            gridCellRect.bottom = gridCellRect.top + (long)pRowInfoArray[row].nHeight;

            GRIDPOINT gridPt = { (UINT)row, (UINT)col };
            PGRIDCELL pCell = GetCellOrDefault((UINT)row, (UINT)col);

            if (CanDrawRect(gridCellRect) && pCell != &m_defaultGridCell)
            {
                // Draw the cell base
                if (pCell->penStyle != m_defaultGridCell.penStyle ||
                    pCell->clrBackground != m_defaultGridCell.clrBackground)
                {
                    HBRUSH hBackgroundBrush = CreateSolidBrush(pCell->clrBackground);
                    HPEN hBorderPen = CreatePen(pCell->penStyle, pCell->m_nBorderWidth, pCell->m_clrBorderColor);
                    HPEN hWhiteOutPen = CreatePen(PS_SOLID, 1, pCell->clrBackground);
                    HPEN hOldPen = (HPEN)SelectObject(hDC, hWhiteOutPen);
                    HBRUSH hOldBrush = (HBRUSH)SelectObject(hDC, hBackgroundBrush);
                    Rectangle(hDC, gridCellRect.left, gridCellRect.top, gridCellRect.right, gridCellRect.bottom);
                    SelectObject(hDC, hBorderPen);
                    //SelectObject(hDC, hBackgroundBrush);
                    Rectangle(hDC, gridCellRect.left, gridCellRect.top, gridCellRect.right, gridCellRect.bottom);
                    SelectObject(hDC, hOldPen);
                    SelectObject(hDC, hOldBrush);
                    DeleteObject(hBorderPen);
                    DeleteObject(hWhiteOutPen);
                    DeleteObject(hBackgroundBrush);
                }
                // Draw the cell text
                SetBkMode(hDC, TRANSPARENT);
                SetTextColor(hDC, pCell->fontInfo.m_clrTextColor);

                HFONT hFont = CreateFontW(
                    (int)(pCell->fontInfo.m_fPointSize * GetDeviceCaps(hDC, LOGPIXELSY)) / -72, 0, 0, 0,
                    pCell->fontInfo.bWeight, pCell->fontInfo.bItalic,
                    pCell->fontInfo.bUnderline, 0, 0, 0, 0, 0, 0,
                    pCell->fontInfo.m_wsFontFace.c_str()
                );

                HFONT hOldFont = (HFONT)SelectObject(hDC, hFont);

                DrawTextW(hDC, pCell->m_wsText.c_str(), -1, &gridCellRect, pCell->justification | DT_SINGLELINE);

                SelectObject(hDC, hOldFont);
                DeleteObject(hFont);

            }

            if (!IsSelectionUnicellular() && 
                IsCellPointInsideSelection(gridPt) && 
                !AreGridPointsEqual(gridPt, m_currentCell))
            {
                DrawSelectionOverlay(hDC, gridCellRect);
            }
        }

        gridCellRect = { 0, 0, 0, 0 };
    }
}

void CGrid32Mgr::DrawVoidSpace(HDC hDC)
{
    bool bDrawBorder = false;
    long nMaxWidth = static_cast<long>(CalculatedColumnDistance(m_visibleTopLeft.nCol, gcs.nWidth));
    long nMaxHeight = static_cast<long>(CalculatedRowDistance(m_visibleTopLeft.nRow, gcs.nHeight));
    if (nMaxWidth < m_clientRect.right)
    {
        HBRUSH hBrush = CreateSolidBrush(RGB(128, 128, 128));
        HPEN hPen = CreatePen(PS_SOLID, 1, RGB(128, 128, 128));

        HPEN oldPen = (HPEN)SelectObject(hDC, hPen);
        HBRUSH oldBrush = (HBRUSH)SelectObject(hDC, hBrush);

        Rectangle(hDC, nMaxWidth + GetActualRowHeaderWidth(), m_clientRect.top, m_clientRect.right, m_clientRect.bottom);

        DeleteObject(SelectObject(hDC, oldPen));
        DeleteObject(SelectObject(hDC, oldBrush));

        bDrawBorder = true;
    }
    if (nMaxHeight < m_clientRect.bottom)
    {
        HBRUSH hBrush = CreateSolidBrush(RGB(128, 128, 128));
        HPEN hPen = CreatePen(PS_SOLID, 1, RGB(128, 128, 128));

        HPEN oldPen = (HPEN)SelectObject(hDC, hPen);
        HBRUSH oldBrush = (HBRUSH)SelectObject(hDC, hBrush);

        Rectangle(hDC, 0, nMaxHeight + GetActualColHeaderHeight(), m_clientRect.right, m_clientRect.bottom);

        DeleteObject(SelectObject(hDC, oldPen));
        DeleteObject(SelectObject(hDC, oldBrush));

        bDrawBorder = true;
    }

    if (bDrawBorder)
    {
        int x = nMaxWidth < m_clientRect.right ? nMaxWidth : m_clientRect.right;
        int y = nMaxHeight < m_clientRect.bottom ? nMaxHeight : m_clientRect.bottom;

        x += GetActualRowHeaderWidth() + 1;
        y += GetActualColHeaderHeight() + 1;

        HPEN hPen = CreatePen(PS_SOLID, 3, 0);
        HPEN oldPen = (HPEN)SelectObject(hDC, hPen);

        if (x < m_clientRect.right)
        {
            MoveToEx(hDC, x, 0, NULL);
            LineTo(hDC, x, y < m_clientRect.bottom ? y : m_clientRect.bottom);
        }

        if (y < m_clientRect.bottom)
        {
            MoveToEx(hDC, 0, y, NULL);
            LineTo(hDC, x < m_clientRect.right ? x : m_clientRect.right, y);
        }


        DeleteObject(SelectObject(hDC, oldPen));
    }
}

void CGrid32Mgr::DrawSizingLine(HDC hDC)
{
    if (!m_bResizable || !m_bSizing)
        return;

    HPEN hOldPen = NULL, hPen = CreatePen(PS_DASHDOT, 1, m_rgbSizingLine);
    POINT pt;
    GetCursorPos(&pt);
    ScreenToClient(m_hWndGrid, &pt);
    hOldPen = (HPEN)SelectObject(hDC, hPen);
    if (m_gridHitTest == GRID_HTCOLDIVIDER)
    {
        MoveToEx(hDC, pt.x, 0, NULL);
        LineTo(hDC, pt.x, m_clientRect.bottom);
    }
    else if (m_gridHitTest == GRID_HTROWDIVIDER)
    {
        MoveToEx(hDC, 0, pt.y, NULL);
        LineTo(hDC, m_clientRect.right, pt.y);
    }
    SelectObject(hDC, hOldPen);
    DeleteObject(hPen);
}

void CGrid32Mgr::DrawSemiTransparentRectangle(HDC hdc, RECT rect, COLORREF color, BYTE alpha) 
{
    Graphics graphics(hdc);
    Color fillColor(alpha, GetRValue(color), GetGValue(color), GetBValue(color));
    SolidBrush brush(fillColor);
    //graphics.FillRectangle(&brush, rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top);
    Gdiplus::Rect rect2;
    rect2.X = rect.left;
    rect2.Y = rect.top;
    rect2.Width = rect.right - rect.left;
    rect2.Height = rect.bottom - rect.top;
    graphics.FillRectangle(&brush, rect2);
}

void CGrid32Mgr::DrawSelectionSurround(HDC hDC)
{
    if (IsSelectionUnicellular())
        return;

    GRIDSELECTION normalizedSelection = m_selectionRect;

    NormalizeSelectionRect(normalizedSelection);
    POINT startPt = { (long)GetColumnStartPos((UINT)normalizedSelection.start.nCol),
                        (long)GetRowStartPos((UINT)normalizedSelection.start.nRow) },
          endPt = { (long)GetColumnStartPos((UINT)normalizedSelection.end.nCol),
                        (long)GetRowStartPos((UINT)normalizedSelection.end.nRow) };

    UINT endCol = normalizedSelection.end.nCol < gcs.nWidth ?
        normalizedSelection.end.nCol : (UINT)(gcs.nWidth - 1);
    UINT endRow = normalizedSelection.end.nRow < gcs.nHeight ?
        normalizedSelection.end.nRow : (UINT)(gcs.nHeight - 1);

    endPt.x += (long)pColInfoArray[endCol].nWidth;
    endPt.y += (long)pRowInfoArray[endRow].nHeight;

    if (normalizedSelection.start.nCol < m_visibleTopLeft.nCol)
        startPt.x = -10;
    if (normalizedSelection.start.nRow < m_visibleTopLeft.nRow)
        startPt.y = -10;

    HPEN hPen = CreatePen(PS_SOLID, 3, gcs.clrSelectionOverlay), hOldPen;
    HBRUSH hOldBrush;
    hOldPen = (HPEN)SelectObject(hDC, hPen);
    hOldBrush = (HBRUSH)SelectObject(hDC, GetStockObject(NULL_BRUSH));
    Rectangle(hDC, startPt.x, startPt.y, endPt.x, endPt.y);
    SelectObject(hDC, hOldBrush);
    DeleteObject(SelectObject(hDC, hOldPen));
    hPen = NULL;
}

bool CGrid32Mgr::IsCellMerged(UINT row, UINT col) 
{
    return false;
}

void CGrid32Mgr::DrawMergedCell(HDC hDC, const MERGERANGE& mergeRange) {
    // Calculate the combined rectangle of the merged cells
    // Draw the cell content (only in the top-left cell)
    // Handle borders and other visual elements
}

void CGrid32Mgr::SelectMergedRange(UINT row, UINT col)
{
    // Find the merged range and select all cells within it
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

void CGrid32Mgr::GetCurrentCellRect(RECT &rect)
{
    memset(&rect, 0, sizeof(RECT));
    if (!(gcs.style & GS_HIGHLIGHTSELECTION))
        return;

    RECT gridCellRect = totalGridCellRect;
    OffsetRectByScroll(gridCellRect);

    gridCellRect.right = gridCellRect.left;
    gridCellRect.bottom = gridCellRect.top;

    for (size_t col = 0; col < gcs.nWidth; ++col)
    {
        gridCellRect.right += (long)pColInfoArray[col].nWidth;
        if (col == m_currentCell.nCol)
        {
            for (size_t row = 0; row < gcs.nHeight; ++row)
            {
                gridCellRect.bottom += (long)pRowInfoArray[row].nHeight;
                if (row == m_currentCell.nRow)
                {
                    if (CanDrawRect(gridCellRect))
                    {
                        rect = gridCellRect;
                        return;
                    }
                }
                gridCellRect.top = gridCellRect.bottom;
                if (m_clientRect.bottom < gridCellRect.top)
                    break;
            }
        }
        gridCellRect.left = gridCellRect.right;
        if (m_clientRect.right < gridCellRect.left)
            break;
    }
}

// Retrieve an existing cell pointer or nullptr if absent
PGRIDCELL CGrid32Mgr::GetCell(UINT nRow, UINT nCol)
{
    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);
    return (it != mapCells.end()) ? it->second : nullptr;
}

// Retrieve an existing cell or the default cell if absent
PGRIDCELL CGrid32Mgr::GetCellOrDefault(UINT nRow, UINT nCol)
{
    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);
    return (it != mapCells.end()) ? it->second : &m_defaultGridCell;
}

// Retrieve an existing cell or create a new one with undo support
PGRIDCELL CGrid32Mgr::GetCellOrCreate(UINT nRow, UINT nCol, bool bRecord)
{
    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);
    if (it != mapCells.end())
        return it->second;

    // Create new cell
    PGRIDCELL pCell = new GRIDCELL(m_defaultGridCell);
    if (!pCell) {
        SetLastError(GRID_ERROR_OUT_OF_MEMORY);
        SendGridNotification(NM_OUTOFMEMORY);
        return &m_defaultGridCell;
    }
    mapCells[key] = pCell;

    if (bRecord)
    {
        // Record creation as full-cell set
        GridEditOperation op;
        op.type = EditOperationType::SetFullCell;
        op.row = nRow;
        op.col = nCol;
        op.oldState = m_defaultGridCell;
        op.newState = *pCell;
        RecordUndoOperation(op);
    }

    return pCell;
}

PGRIDCELL CGrid32Mgr::CreateCell(UINT nRow, UINT nCol)
{
	auto key = std::make_pair(nRow, nCol);
	auto it = mapCells.find(key);
	if (it != mapCells.end()) {
		// Cell already exists
		delete it->second;
		mapCells.erase(it); // Remove the existing cell
	}
	// Create new cell
	PGRIDCELL pCell = new GRIDCELL(m_defaultGridCell);
	if (!pCell) {
		SetLastError(GRID_ERROR_OUT_OF_MEMORY);
		SendGridNotification(NM_OUTOFMEMORY);
		return nullptr;
	}
	mapCells[key] = pCell;
	return pCell;
}

// Refactored SetCell method with clean undo/redo handling
void CGrid32Mgr::SetCell(UINT nRow, UINT nCol, const GRIDCELL& gc)
{
    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);

    GridEditOperation op;
    op.type = EditOperationType::SetFullCell;
    op.row = nRow;
    op.col = nCol;
    op.newState = gc; // Always set new state

    if (it != mapCells.end()) {
        // Existing cell: capture old state
        op.oldState = *it->second;

        // Replace cell
        delete it->second;
        it->second = new GRIDCELL(gc);
        if (!it->second) {
            SetLastError(GRID_ERROR_OUT_OF_MEMORY);
            SendGridNotification(NM_OUTOFMEMORY);
            return;
        }
    }
    else {
        // New cell: old state is default
        op.oldState = m_defaultGridCell;

        // Create cell
        PGRIDCELL pCell = new GRIDCELL(gc);
        if (!pCell) {
            SetLastError(GRID_ERROR_OUT_OF_MEMORY);
            SendGridNotification(NM_OUTOFMEMORY);
            return;
        }
        mapCells[key] = pCell;
    }

    // Record undo/redo
    RecordUndoOperation(op);
    SetLastError(0);
}


UINT CGrid32Mgr::GetCurrentCellTextLen()
{
    return GetCellTextLen(m_currentCell.nRow, m_currentCell.nCol);
}

UINT CGrid32Mgr::GetCellTextLen(const GRIDPOINT& gridPt)
{
    return GetCellTextLen(gridPt.nRow, gridPt.nCol);
}

UINT CGrid32Mgr::GetCellTextLen(UINT nRow, UINT nCol)
{
    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);

    if (it != mapCells.end()) {
        return (UINT)(it->second->m_wsText.length() + 2);
    }

    return 0;
}

void CGrid32Mgr::GetCurrentCellText(LPWSTR pText, UINT nLen)
{
    return GetCellText(m_currentCell.nRow, m_currentCell.nCol, pText, nLen);
}

void CGrid32Mgr::GetCellText(const GRIDPOINT& gridPt, LPWSTR pText, UINT nLen)
{
    return GetCellText(gridPt.nRow, gridPt.nCol, pText, nLen);
}

// Direct FONTINFO overload
void CGrid32Mgr::SetCellFormat(UINT nRow, UINT nCol, const FONTINFO& fi)
{
    // Prepare undo operation
    GridEditOperation op;
    op.type = EditOperationType::SetFormat;
    op.row = nRow;
    op.col = nCol;

    // Capture old state
    PGRIDCELL pCell = GetCell(nRow, nCol);
    if (pCell)
    {
        op.oldState = *pCell;
    }
    else
    {
        op.oldState = m_defaultGridCell;
        pCell = CreateCell(nRow, nCol);
        if (!pCell) {
            SetLastError(GRID_ERROR_OUT_OF_MEMORY);
            SendGridNotification(NM_OUTOFMEMORY);
            return;
        }
    }

    // Apply new formatting
    pCell->fontInfo = fi;
    op.newState = *pCell;

    // Record undo operation
    RecordUndoOperation(op);

    // Redraw updated cell
    RECT r;
    GetCurrentCellRect(r);
    Invalidate();
    SetLastError(0);
}




void CGrid32Mgr::GetCellText(UINT nRow, UINT nCol, LPWSTR pText, UINT nLen)
{
    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);

    if (pText)
    {
        if (it != mapCells.end()) {
            _tcsncpy_s(pText, nLen, 
                it->second->m_wsText.c_str(), it->second->m_wsText.length()); 
            return;
        }

        *pText = 0;
        return;
    }
}

void CGrid32Mgr::SetCurrentCellText(LPCWSTR newText)
{
    SetCellText(m_currentCell.nRow, m_currentCell.nCol, newText);
}

void CGrid32Mgr::SetCellText(const GRIDPOINT& gridPt, LPCWSTR newText)
{
    SetCellText(gridPt.nRow, gridPt.nCol, newText);
}


void CGrid32Mgr::SetCellText(UINT nRow, UINT nCol, LPCWSTR newText)
{
    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);

    PGRIDCELL pCell = nullptr;
    GridEditOperation op;
    op.type = EditOperationType::SetText;
    op.row = nRow;
    op.col = nCol;

    if (it != mapCells.end()) {
        pCell = it->second;
        op.oldState = *pCell;
    }
    else {
        pCell = new GRIDCELL(m_defaultGridCell);
        if (!pCell) {
            SetLastError(GRID_ERROR_OUT_OF_MEMORY);
            SendGridNotification(NM_OUTOFMEMORY);
            return;
        }
        mapCells[key] = pCell;
        op.oldState = m_defaultGridCell;
    }

    // Apply the new text
    pCell->m_wsText = newText;
    op.newState = *pCell;

    // Record undo operation
    RecordUndoOperation(op);

    // Redraw the cell
    RECT r;
    GetCurrentCellRect(r);
    Invalidate();
    SetLastError(0);
}


// Clear the text of a cell and support undo
void CGrid32Mgr::ClearCellText(UINT nRow, UINT nCol)
{
    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);

    if (it != mapCells.end()) {
        // Setup undo operation
        GridEditOperation op;
        op.type = EditOperationType::SetText;
        op.row = nRow;
        op.col = nCol;
        // Capture old state
        op.oldState = *it->second;

        // Clear the text
        it->second->m_wsText.clear();
        // Capture new state
        op.newState = *it->second;

        // Record undo operation
        RecordUndoOperation(op);

        // Redraw the cell
        RECT r;
        GetCurrentCellRect(r);
        Invalidate();
        SetLastError(0);
    }
}

// Delete a cell and support undo
void CGrid32Mgr::DeleteCell(UINT nRow, UINT nCol)
{
    auto key = std::make_pair(nRow, nCol);
    auto it = mapCells.find(key);
    if (it != mapCells.end())
    {
        // Setup undo operation
        GridEditOperation op;
        op.type = EditOperationType::Delete;
        op.row = nRow;
        op.col = nCol;
        // Capture old state
        op.oldState = *it->second;
        // New state is default (no cell)
        op.newState = m_defaultGridCell;

        // Perform deletion
        delete it->second;
        mapCells.erase(it);

        // Record undo operation
        RecordUndoOperation(op);
        SetLastError(0);
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
        IncrementSelectedCell(-1, 0);
        break;
    case SB_LINERIGHT:
        IncrementSelectedCell(1, 0);
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

    Invalidate(true);
}

void CGrid32Mgr::OnVScroll(UINT nSBCode, UINT nPos, HWND hScrollBar)
{
    switch (nSBCode)
    {
    case SB_LINEUP:
        IncrementSelectedCell(0, -1);
        break;
    case SB_LINEDOWN:
        IncrementSelectedCell(0, 1);
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
    Invalidate(false);
}


void CGrid32Mgr::SetScrollRanges()
{
    if (gcs.style & WS_HSCROLL)
    {
        // Set the horizontal scroll range
        SCROLLINFO siH = { sizeof(SCROLLINFO) };
        siH.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
        siH.nMin = 0;
        siH.nMax = static_cast<int>(gcs.nWidth); 
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

void Nop(){}

bool CGrid32Mgr::IsCellIncrementWithinCurrentPage(long nRow, long nCol, PAGESTAT& pageStat)
{
    if (nRow < 0 && (size_t)abs(nRow) > m_currentCell.nRow)
        return false;
    if (nCol < 0 && (size_t)abs(nCol) > m_currentCell.nCol)
        return false;

    return((m_currentCell.nCol + nCol) <= pageStat.end.nCol &&
            (m_currentCell.nCol + nCol) >= pageStat.start.nCol &&
            (m_currentCell.nRow + nRow) <= pageStat.end.nRow &&
            (m_currentCell.nRow + nRow) >= pageStat.start.nRow);
}

void CGrid32Mgr::IncrementSelectedCell(long nRow, long nCol)
{
    PAGESTAT pageStat;
    CalculatePageStats(pageStat);
    _ASSERT(pageStat.nWidth < ((size_t)m_clientRect.right - GetActualRowHeaderWidth()));
    _ASSERT(pageStat.nHeight < ((size_t)m_clientRect.bottom - GetActualColHeaderHeight()));
    Nop();
    
    if (nRow > 0 && m_currentCell.nRow == gcs.nHeight - 1)
        nRow = 0;
    if (nCol > 0 && m_currentCell.nCol == gcs.nWidth - 1)
        nCol = 0;
    if (IsCellIncrementWithinCurrentPage(nRow, nCol, pageStat))
    {
        m_currentCell.nRow += nRow;
        m_currentCell.nCol += nCol;
    }
    else
    {
        if (nRow < 0)
        {
            size_t nHeight = 0;
            if ((long long)(m_visibleTopLeft.nRow + (size_t)nRow) > 0)
            {
                m_currentCell.nRow = m_visibleTopLeft.nRow + (size_t)nRow;
                for (size_t idx = m_visibleTopLeft.nRow - 1; idx != ~0; --idx)
                {
                    if (idx == 0)
                    {
                        m_visibleTopLeft.nRow = 0;
                        break;
                    }
                    else if((nHeight + pRowInfoArray[idx].nHeight) < (size_t)m_clientRect.bottom - GetActualColHeaderHeight())
                        nHeight += pRowInfoArray[idx].nHeight;
                    else
                    {
                        m_visibleTopLeft.nRow = (UINT)idx + 1;
                        break;
                    }
                }
            }
            else
            {
                m_visibleTopLeft.nRow = m_currentCell.nRow = 0;
            }
        }
        else if (nRow > 0)
        {
            if (pageStat.end.nRow + (size_t)nRow < gcs.nHeight - 1)
            {
                auto rowDiff = pageStat.end.nRow - m_visibleTopLeft.nRow;
                m_visibleTopLeft.nRow = m_currentCell.nRow = pageStat.end.nRow + (size_t)nRow;
            }
        }

        if (nCol < 0)
        {
            size_t nWidth = 0;
            if ((long long)(m_visibleTopLeft.nCol + (size_t)nCol) > 0)
            {
                m_currentCell.nCol = m_visibleTopLeft.nCol + (size_t)nCol;
                for (size_t idx = m_visibleTopLeft.nCol - 1; idx != ~0; --idx)
                {
                    if (idx == 0)
                    {
                        m_visibleTopLeft.nCol = 0;
                        break;
                    }
                    else if ((nWidth + pColInfoArray[idx].nWidth) < (size_t)m_clientRect.right - GetActualRowHeaderWidth())
                        nWidth += pColInfoArray[idx].nWidth;
                    else
                    {
                        m_visibleTopLeft.nCol = (UINT)idx + 1;
                        break;
                    }
                }
            }
            else
            {
                m_visibleTopLeft.nCol = m_currentCell.nCol = 0;
            }
        }
        else if (nCol > 0)
        {
            if (pageStat.end.nCol + (size_t)nCol < gcs.nWidth - 1)
            {
                auto colDiff = pageStat.end.nCol - m_visibleTopLeft.nCol;
                m_visibleTopLeft.nCol = m_currentCell.nCol = pageStat.end.nCol + (size_t)nCol;
            }
        }
    }

    m_scrollDifference.x = (long)CalculatedColumnDistance(0, (size_t)m_visibleTopLeft.nCol);
    m_scrollDifference.y = (long)CalculatedRowDistance(0, (size_t)m_visibleTopLeft.nRow);

    Invalidate(false);
}

void CGrid32Mgr::IncrementSelectedPage(long nRow, long nCol)
{
    if (nRow < 0)
    {
        size_t nHeight = 0;
        auto rowDiff = m_currentCell.nRow - m_visibleTopLeft.nRow;
        for (size_t idx = m_visibleTopLeft.nRow - 1; idx != ~0; --idx)
        {
            if (idx == 0 || m_visibleTopLeft.nRow == 0)
            {
                m_visibleTopLeft.nRow = 0;
                m_currentCell.nRow = rowDiff;
                break;
            }
            else if ((nHeight + pRowInfoArray[idx].nHeight) < (size_t)m_clientRect.bottom - GetActualColHeaderHeight())
                nHeight += pRowInfoArray[idx].nHeight;
            else
            {
                ++idx;
                auto rowDiff = m_visibleTopLeft.nRow - (UINT)idx;
                m_currentCell.nRow -= rowDiff;
                m_visibleTopLeft.nRow = (UINT)idx;
                break;
            }
        }
    }
    else if (nRow > 0)
    {
        for (size_t idx = 0; idx < abs(nRow); ++idx)
        {
            PAGESTAT pageStat;
            CalculatePageStats(pageStat);
            UINT nRowDiff = pageStat.end.nRow + 1 - pageStat.start.nRow;
            if (((size_t)m_visibleTopLeft.nRow + nRowDiff) > gcs.nHeight ||
                ((size_t)m_currentCell.nRow + nRowDiff) > gcs.nHeight)
            {
                IncrementSelectEdge(nRow, nCol);
                return;
            }
            m_visibleTopLeft.nRow += nRowDiff;
            m_currentCell.nRow += nRowDiff;
        }
    }

    if (nCol < 0)
    {
        size_t nWidth = 0;
        auto colDiff = m_currentCell.nCol - m_visibleTopLeft.nCol;
        for (size_t idx = m_visibleTopLeft.nCol - 1; idx != ~0; --idx)
        {
            if (idx == 0 || m_visibleTopLeft.nCol == 0)
            {
                m_visibleTopLeft.nCol = 0;
                m_currentCell.nCol = colDiff;
                break;
            }
            else if ((nWidth + pColInfoArray[idx].nWidth) < (size_t)m_clientRect.right - GetActualRowHeaderWidth())
                nWidth += pColInfoArray[idx].nWidth;
            else
            {
                ++idx;
                auto colDiff = m_visibleTopLeft.nCol - idx;
                m_currentCell.nCol -= (UINT)colDiff;
                m_visibleTopLeft.nCol = (UINT)idx;
                break;
            }
        }
    }
    else if (nCol > 0)
    {
        for (size_t idx = 0; idx < abs(nCol); ++idx)
        {
            PAGESTAT pageStat;
            CalculatePageStats(pageStat);
            UINT nRowDiff = pageStat.end.nCol + 1 - pageStat.start.nCol;
            if (((size_t)m_visibleTopLeft.nCol + nRowDiff) > gcs.nWidth ||
                ((size_t)m_currentCell.nCol + nRowDiff) > gcs.nWidth)
            {
                IncrementSelectEdge(nRow, nCol);
                return;
            }
            m_visibleTopLeft.nCol += nRowDiff;
            m_currentCell.nCol += nRowDiff;
        }
    }

    m_scrollDifference.x = (long)CalculatedColumnDistance(0, (size_t)m_visibleTopLeft.nCol);
    m_scrollDifference.y = (long)CalculatedRowDistance(0, (size_t)m_visibleTopLeft.nRow);

    Invalidate(false);
}

void CGrid32Mgr::IncrementSelectEdge(long nRow, long nCol)
{
    if (nRow < 0)
    {
        m_visibleTopLeft.nRow = 0;
        m_currentCell.nRow = 0;
    }
    else if (nRow > 0)
    {
        size_t nHeight = 0;
        m_currentCell.nRow = (UINT)gcs.nHeight - 1;
        size_t idx = m_currentCell.nRow - 1;
        for (; nHeight < ((size_t)m_clientRect.bottom * 2) / 3; --idx)
        {
            nHeight += pRowInfoArray[idx].nHeight;
        }

        m_visibleTopLeft.nRow = (UINT)idx;
    }
    if (nCol < 0)
    {
        m_visibleTopLeft.nCol = 0;
        m_currentCell.nCol = 0;
    }
    else if (nCol > 0)
    {
        size_t nWidth = 0;
        m_currentCell.nCol = (UINT)gcs.nWidth - 1;
        size_t idx = m_currentCell.nCol - 1;
        for (; nWidth < (m_clientRect.right * 2) / 3; --idx)
        {
            nWidth += pColInfoArray[idx].nWidth;
        }

        m_visibleTopLeft.nCol = (UINT)idx;
    }
    m_scrollDifference.x = (long)CalculatedColumnDistance(0, (size_t)m_visibleTopLeft.nCol);
    m_scrollDifference.y = (long)CalculatedRowDistance(0, (size_t)m_visibleTopLeft.nRow);

    Invalidate(false);

}


size_t CGrid32Mgr::MaxBeginRow()
{
    return size_t();
}

size_t CGrid32Mgr::MaxBeginColumn()
{
    return size_t();
}

void CGrid32Mgr::CalculatePageStats(PAGESTAT& pageStat)
{
    memset(&pageStat, 0, sizeof(PAGESTAT));
    pageStat.start = m_visibleTopLeft;
    for (size_t idx = (size_t)m_visibleTopLeft.nCol; idx < gcs.nWidth; ++idx)
    {
        if (idx == gcs.nWidth - 1)
            pageStat.end.nCol = (UINT)gcs.nWidth;
        if (pageStat.nWidth + pColInfoArray[idx].nWidth <= ((size_t)m_clientRect.right - GetActualRowHeaderWidth()))
            pageStat.nWidth += pColInfoArray[idx].nWidth;
        else
        {
            pageStat.end.nCol = (UINT)idx - 1;
            break;
        }
    }
    for (size_t idx = m_visibleTopLeft.nRow; idx < gcs.nHeight; ++idx)
    {
        if (idx == gcs.nHeight - 1)
        {
            pageStat.end.nRow = (UINT)gcs.nHeight - 1;
            break;
        }
        else if (pageStat.nHeight + pRowInfoArray[idx].nHeight <= ((size_t)m_clientRect.bottom - GetActualColHeaderHeight()))
            pageStat.nHeight += pRowInfoArray[idx].nHeight;
        else
        {
            pageStat.end.nRow = (UINT)idx - 1;
            break;
        }
    }


}


bool CGrid32Mgr::IsCellVisible(UINT nRow, UINT nCol)
{
    RECT cell = { 0, 0, 0, 0 };
    RECT visibleGridRect = { GetActualRowHeaderWidth(), GetActualColHeaderHeight(), m_clientRect.right, m_clientRect.bottom };

    if (m_visibleTopLeft.nRow > nRow ||
        m_visibleTopLeft.nCol > nCol)
        return false;

    if (m_visibleTopLeft.nRow == nRow &&
        m_visibleTopLeft.nCol == nCol)
        return true;

    GRIDSIZE cellPos = { 0,0 };
    cell = { GetActualRowHeaderWidth(), GetActualColHeaderHeight(), (long)pColInfoArray[m_visibleTopLeft.nCol].nWidth, (long)pRowInfoArray[m_visibleTopLeft.nRow].nHeight };

    for (UINT idx = m_visibleTopLeft.nRow; idx < nRow; ++idx)
       OffsetRect(&cell, 0, (int)pRowInfoArray[idx].nHeight);
    for (UINT idx = m_visibleTopLeft.nCol; idx < nCol; ++idx)
        OffsetRect(&cell, (int)pColInfoArray[idx].nWidth, 0);

    if (cellPos.width < m_clientRect.right && cellPos.height < m_clientRect.bottom)
        return true;

    return false;
}

void CGrid32Mgr::ScrollToCell(size_t row, size_t col, int scrollFlags)
{

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
    GetClientRect(m_hWndGrid, &m_clientRect);
    visibleGrid.width = m_clientRect.right - GetActualRowHeaderWidth();
    visibleGrid.height = m_clientRect.bottom - GetActualColHeaderHeight();
}

bool CGrid32Mgr::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
    switch (nChar)
    {
    case VK_LEFT:
        if (GetAsyncKeyState(VK_CONTROL) & 0x8000)
            IncrementSelectEdge(0, -1);
        else if (GetAsyncKeyState(VK_SHIFT) & 0x8000)
            IncrementSelectedPage(0, -1);
        else
            IncrementSelectedCell(0, -1);
        return true;

    case VK_RIGHT:
        if (GetAsyncKeyState(VK_CONTROL) & 0x8000)
            IncrementSelectEdge(0, 1);
        else if (GetAsyncKeyState(VK_SHIFT) & 0x8000)
            IncrementSelectedPage(0, 1);
        else
            IncrementSelectedCell(0, 1);
        return true;

    case VK_UP:
        if (GetAsyncKeyState(VK_CONTROL) & 0x8000)
            IncrementSelectEdge(-1, 0);
        else if (GetAsyncKeyState(VK_SHIFT) & 0x8000)
            IncrementSelectedPage(-1, 0);
        else
            IncrementSelectedCell(-1, 0);
        return true;

    case VK_DOWN:
        if (GetAsyncKeyState(VK_CONTROL) & 0x8000)
            IncrementSelectEdge(1, 0);
        else if (GetAsyncKeyState(VK_SHIFT) & 0x8000)
            IncrementSelectedPage(1, 0);
        else
            IncrementSelectedCell(1, 0);
        return true;
    case VK_PRIOR:
        IncrementSelectedPage(-1, 0);
        return true;
    case VK_NEXT:
        IncrementSelectedPage(1, 0);
        return true;
    case VK_DELETE:
        ClearCellText(m_currentCell.nRow, m_currentCell.nCol);
        Invalidate();
        break;
    default:
        // Handle alphanumeric characters, space, and specific punctuation
        if (VK_BACK == nChar || std::isalnum(nChar) || std::isspace(nChar) || allowedPunctuation.count(nChar))
        {
            // Show edit control at current cell position
            UINT nLen = GetCurrentCellTextLen();
            if (nLen)
            {
                TCHAR* buff = new TCHAR[(size_t)nLen + 2];
                GetCurrentCellText(buff, nLen);
                SetWindowText(m_hWndEdit, buff);
                PostMessage(m_hWndEdit, EM_SETSEL, nLen, nLen);
                delete[] buff;
            }
            else
                SetWindowText(m_hWndEdit, L"");

            ShowEditControl();

            // Send the character to the edit control
            PostMessage(m_hWndEdit, WM_KEYDOWN, nChar, nRepCnt);
            return true;
        }
        break;
    }

    return false;
}

bool CGrid32Mgr::OnKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags)
{
    // Handle key releases here
    return false;
}

void CGrid32Mgr::GetCellByPoint(int x, int y, GRIDPOINT& gridPt)
{
    if (x < GetActualRowHeaderWidth() || x >= m_clientRect.right)
        gridPt.nCol = ~0;
    else
    {
        size_t nWidth = (size_t)GetActualRowHeaderWidth(); 
        size_t idx = m_visibleTopLeft.nCol;
        if (idx > (size_t)(~0)/2)
            idx = m_visibleTopLeft.nCol = 0;
        for (; nWidth + pColInfoArray[idx].nWidth < x; ++idx, nWidth += pColInfoArray[idx].nWidth);
        gridPt.nCol = (UINT)idx;

    }

    if(y < GetActualColHeaderHeight() || y >= m_clientRect.bottom)
    {
        gridPt.nRow = ~0;
    }
    else
    {
        size_t nHeight = (size_t)GetActualColHeaderHeight();
        size_t idx = m_visibleTopLeft.nRow;
        if (idx > (size_t)-100)
            idx = m_visibleTopLeft.nRow = 0;
        for (; nHeight + pRowInfoArray[idx].nHeight < y; ++idx, nHeight += pRowInfoArray[idx].nHeight);
        gridPt.nRow = (UINT)idx;
    }
}

void CGrid32Mgr::SetCellByPoint(int x, int y, GRIDPOINT &gridPt)
{
    if (x < GetActualRowHeaderWidth() || x >= m_clientRect.right ||
        y < GetActualColHeaderHeight() || y >= m_clientRect.bottom)
    {
        gridPt.nRow = ~0;
        gridPt.nCol = ~0;
    }
    else
    {
        size_t nWidth = (size_t)GetActualRowHeaderWidth(), nHeight = (size_t)GetActualColHeaderHeight();
        size_t idx = m_visibleTopLeft.nCol;
        for (; nWidth + pColInfoArray[idx].nWidth < x; ++idx, nWidth += pColInfoArray[idx].nWidth);
        m_currentCell.nCol = (UINT)idx;
        if (nWidth + pColInfoArray[idx].nWidth > m_clientRect.right)
        {
            m_visibleTopLeft.nCol = (UINT)idx;
            m_scrollDifference.x = (long)CalculatedColumnDistance(0, (size_t)m_visibleTopLeft.nCol);
        }
        idx = m_visibleTopLeft.nRow;
        for (; nHeight + pRowInfoArray[idx].nHeight < y; ++idx, nHeight += pRowInfoArray[idx].nHeight);
        m_currentCell.nRow = (UINT)idx;
        if (nHeight + pRowInfoArray[idx].nHeight > m_clientRect.bottom)
        {
            m_visibleTopLeft.nRow = (UINT)idx;
            m_scrollDifference.y = (long)CalculatedRowDistance(0, (size_t)m_visibleTopLeft.nRow);
        }

        gridPt = m_currentCell;
        Invalidate(false);
    }
}

void CGrid32Mgr::StartMouseTracking()
{
    // Get the default hover time-out from system settings
    UINT hoverTime = 0;
    SystemParametersInfo(SPI_GETMOUSEHOVERTIME, 0, &hoverTime, 0);

    // Set up the TRACKMOUSEEVENT structure
    TRACKMOUSEEVENT tme;
    memset(&tme, 0, sizeof(TRACKMOUSEEVENT));
    tme.cbSize = sizeof(TRACKMOUSEEVENT);
    tme.dwFlags = TME_HOVER | TME_LEAVE;
    tme.hwndTrack = m_hWndGrid;
    //tme.dwHoverTime = hoverTime; // Use the system hover time
    tme.dwHoverTime = 100;

    // Start tracking
    TrackMouseEvent(&tme);
}

void CGrid32Mgr::CancelMouseTracking()
{
    // Set up the TRACKMOUSEEVENT structure to cancel tracking
    TRACKMOUSEEVENT tme;
    memset(&tme, 0, sizeof(TRACKMOUSEEVENT));
    tme.cbSize = sizeof(TRACKMOUSEEVENT);
    tme.dwFlags = TME_CANCEL | TME_HOVER | TME_LEAVE;
    tme.hwndTrack = m_hWndGrid;

    // Cancel tracking
    TrackMouseEvent(&tme);
}

void CGrid32Mgr::ConstrainMouseToClientArea()
{
    RECT clientRect, screenRect;
    GetClientRect(m_hWndGrid, &clientRect);
    GetWindowRect(GetDesktopWindow(), &screenRect);

    MapWindowPoints(m_hWndGrid, GetDesktopWindow(), (LPPOINT)&clientRect, 2);

    screenRect.left = clientRect.left - 1;
    if (screenRect.left < 1)
        screenRect.left = 1;
    screenRect.top = clientRect.top - 1;
    if (screenRect.top < 1)
        screenRect.top = 1;

    // Constrain the cursor to the screenRect
    ClipCursor(&screenRect);
}

void CGrid32Mgr::CalcSelectionCoordinatesWithMouse(int x, int y, bool& bChangeX, bool bChangeY)
{
    bChangeX = false, bChangeY = false;
    POINT pt;
    GRIDPOINT gridPt;
    GetCursorPos(&pt);
    ScreenToClient(m_hWndGrid, &pt);
    GetCellByPoint(x, y, gridPt);

    if (pt.x <= GetActualRowHeaderWidth())
    {
        if (m_visibleTopLeft.nCol > (UINT)-100)
            m_visibleTopLeft.nCol = 0;
        else if (m_visibleTopLeft.nCol > 0)
            --m_visibleTopLeft.nCol;

        if (m_selectionRect.end.nCol > (UINT)-100)
            m_selectionRect.end.nCol = 0;
        else if (m_selectionRect.end.nCol > 0)
            --m_selectionRect.end.nCol;

        bChangeX = true;
    }
    else if (pt.x >= m_clientRect.right)
    {
        ++m_visibleTopLeft.nCol;
        ++m_selectionRect.end.nCol;
        ConstrainMaxPosition();
        bChangeX = true;
    }
    else
    {
        if (gridPt.nCol != ~0)
        {
            if (m_selectionRect.end.nCol != gridPt.nCol)
            {
                m_selectionRect.end.nCol = gridPt.nCol;
                bChangeX = true;
            }
        }
    }

    if (pt.y <= GetActualColHeaderHeight())
    {
        if (m_visibleTopLeft.nRow > 0)
            --m_visibleTopLeft.nRow;
        else
            m_visibleTopLeft.nRow = 0;
        if (m_selectionRect.end.nRow > 0)
            --m_selectionRect.end.nRow;
        else
            m_selectionRect.end.nRow = 0;

        bChangeY = true;
    }
    else if (pt.y >= m_clientRect.bottom)
    {
        ++m_visibleTopLeft.nRow;
        ++m_selectionRect.end.nRow;
        ConstrainMaxPosition();
        bChangeY = true;
    }
    else
    {
        if (gridPt.nRow != ~0)
        {
            if (m_selectionRect.end.nRow != gridPt.nRow)
            {
                m_selectionRect.end.nRow = gridPt.nRow;
                bChangeY = true;
            }
        }
    }

    if (bChangeX)
        m_scrollDifference.x = (long)CalculatedColumnDistance(0, (size_t)m_visibleTopLeft.nCol);
    if (bChangeY)
        m_scrollDifference.y = (long)CalculatedRowDistance(0, (size_t)m_visibleTopLeft.nRow);

}

void CGrid32Mgr::GetCurrentMouseCoordinates(int& x, int& y)
{
    POINT pt;
    GetCursorPos(&pt);
    ScreenToClient(m_hWndGrid, &pt);

    //Get the mouse coordinates ourselves to fix bug with negative coordinates
    x = pt.x;
    y = pt.y;
}

void CGrid32Mgr::OnLButtonDown(UINT nFlags, int x, int y)
{
    GetCurrentMouseCoordinates(x, y);
    // Handle left mouse button down events here
    GRIDPOINT gridPt;
    switch (m_gridHitTest)
    {
    case GRID_HTCELL:
        if (!m_bSelecting && !m_bSizing/* &&
            m_mouseDraggingStartPoint.x <= 0 && m_mouseDraggingStartPoint.y <= 0*/)
        {
            SetCellByPoint(x, y, gridPt);
            m_mouseDraggingStartPoint.x = x;
            m_mouseDraggingStartPoint.y = y;
            if(gridPt.nRow != ~0 && gridPt.nCol != ~0)
                m_selectionRect.start = m_selectionRect.end = gridPt;
            else
            {
                GetCellByPoint(x, y, gridPt);
                _ASSERT(gridPt.nRow != ~0 && gridPt.nCol != ~0);
                m_selectionRect.end = gridPt;
            }

            SetCapture(m_hWndGrid);
            m_bSelecting = TRUE;
            StartMouseTracking();
//            ConstrainMouseToClientArea();

        }
        else if (m_bSelecting)
        {
            GetCellByPoint(x, y, gridPt);
            m_selectionRect.end = gridPt;
            StartMouseTracking();
//            ConstrainMouseToClientArea();
        }

        break;
    case GRID_HTCOLDIVIDER:
    case GRID_HTROWDIVIDER:
        if (m_bResizable && !m_bSelecting)
        {
            SetCapture(m_hWndGrid);
//            ConstrainMouseToClientArea();
            m_bSizing = TRUE;
        }
        break;
    }

    if (GetTickCount64() - m_lastClickTime < GetDoubleClickTime()) {
        // Double-click detected
        SendGridNotification(NM_DBLCLK);
    }

    m_lastClickTime = GetTickCount64();
}

void CGrid32Mgr::OnLButtonUp(UINT nFlags, int x, int y)
{
    GetCurrentMouseCoordinates(x, y);

    if (m_bSelecting)
    {
        GRIDPOINT gridPt;
        GetCellByPoint(x, y, gridPt);
  //      _ASSERT(gridPt.nRow != ~0 && gridPt.nCol != ~0);
        m_selectionRect.end = gridPt;
        ConstrainMaxPosition();
        CancelMouseTracking();
        Invalidate();
    }
    else if (m_bResizable && m_bSizing)
    {
        if (m_gridHitTest == GRID_HTCOLDIVIDER)
        {
            long long nWidth = (signed)GetColumnStartPos((UINT)m_nToBeSized);

            long long nNewSize = x - nWidth;

            if (nNewSize > 0)
                pColInfoArray[m_nToBeSized].nWidth = (size_t)nNewSize;

            Invalidate();
        }
        else if (m_gridHitTest == GRID_HTROWDIVIDER)
        {
            long long nHeight = (signed)GetRowStartPos((UINT)m_nToBeSized);

            long long nNewSize = y - nHeight;

            if (nNewSize > 0)
                pRowInfoArray[m_nToBeSized].nHeight = (size_t)nNewSize;

            Invalidate();
        }
    }
    m_bSelecting = m_bSizing = 0;
    ReleaseCapture();
    ClipCursor(NULL);
    m_gridHitTest = 0;
    m_mouseDraggingStartPoint = { ~0, ~0 };
}


void CGrid32Mgr::OnMouseMove(UINT nFlags, int x, int y)
{
    GetCurrentMouseCoordinates(x, y);

    if (m_bResizable && m_bSizing)
    {
        if (m_gridHitTest == GRID_HTCOLDIVIDER)
        {
            RECT r = { x - 5, 0, x + 5, m_clientRect.bottom };
            Invalidate();
            r.left = m_nSizingLine - 5;
            r.right = r.left + 10;
            Invalidate();
            m_nSizingLine = x;
        }
        else if (m_gridHitTest == GRID_HTROWDIVIDER)
        {
            RECT r = { 0, y - 5, m_clientRect.right, y + 5 };
            Invalidate();
            r.top = m_nSizingLine - 5;
            r.bottom = r.top + 10;
            Invalidate();
            m_nSizingLine = y;
        }
    }
    else if (m_bSelecting)
    {
        bool bChangeX = false, bChangeY = false;
        if (x > GetActualRowHeaderWidth() && x < m_clientRect.right && 
            y > GetActualColHeaderHeight() && y < m_clientRect.bottom)
        {
            GRIDPOINT gridPt;
            GetCellByPoint(x, y, gridPt);
            m_selectionRect.end = gridPt;
        }
        else
        {
            CalcSelectionCoordinatesWithMouse(x, y, bChangeX, bChangeY);
        }

        if (bChangeX || bChangeY)
        {
            ConstrainMaxPosition();
        }
        Invalidate();
    }
    else
    {
        GetCellByPoint(x, y, m_HoverCell);
        if (!m_npHoverDelaySet)
        {
            SendGridNotification(NM_HOVER);
            m_npHoverDelaySet = SetTimer(m_hWndGrid, HOVER_TIMER, m_nMouseHoverDelay, NULL);
        }
    }
}

void CGrid32Mgr::OnMouseHover(UINT nFlags, int x, int y)
{
    GetCurrentMouseCoordinates(x, y);

    bool bChangeX = false, bChangeY = false;
    if (m_bSelecting)
    {
        CalcSelectionCoordinatesWithMouse(x, y, bChangeX, bChangeY);
        StartMouseTracking();
        if (bChangeX || bChangeY)
        {
            m_scrollDifference.x = (long)CalculatedColumnDistance(0, (size_t)m_visibleTopLeft.nCol);
            m_scrollDifference.y = (long)CalculatedRowDistance(0, (size_t)m_visibleTopLeft.nRow);

            Invalidate();
        }
    }
}

void CGrid32Mgr::OnMouseLeave(UINT nFlags, int x, int y)
{
    OnMouseHover(nFlags, x, y);
}

void CGrid32Mgr::OnNcMouseMove(UINT nFlags, int x, int y)
{
    GetCurrentMouseCoordinates(x, y);
}

LRESULT CGrid32Mgr::OnNcHitTest(UINT nFlags, int x, int y, bool& bHitTested)
{
    if (m_bSelecting || m_bSizing)
    {
        bHitTested = true;
        return m_gridHitTest;
    }

    if (IsCoordinateInClientArea(x, y))
    {
        // Convert screen coordinates to client coordinates
        POINT pt = { x, y };
        ScreenToClient(m_hWndGrid, &pt);

        long rowHeaderWidth = GetActualRowHeaderWidth();
        long colHeaderHeight = GetActualColHeaderHeight();

        // Check if the point is in the corner cell
        if (pt.x < rowHeaderWidth + 2 && pt.y < colHeaderHeight + 2)
        {
            bHitTested = true;
            return GRID_HTCORNER;
        }

        // Check if the point is in the column header
        if (pt.y < colHeaderHeight)
        {
            bHitTested = true;        
            return IsOverColumnDivider(pt.x, pt.y) ? 
                GRID_HTCOLDIVIDER : GRID_HTCOLHEADER;
        }

        // Check if the point is in the row header
        if (pt.x < rowHeaderWidth)
        {
            bHitTested = true;
            return IsOverRowDivider(pt.x, pt.y) ? 
                GRID_HTROWDIVIDER : GRID_HTROWHEADER;
        }

        // Check if the point is over a column divider
        if (IsOverColumnDivider(pt.x, pt.y))
        {
            bHitTested = true;
            return GRID_HTCOLDIVIDER;
        }

        // Check if the point is over a row divider
        if (IsOverRowDivider(pt.x, pt.y))
        {
            bHitTested = true;
            return GRID_HTROWDIVIDER;
        }

        // Otherwise, the point is over a cell
        bHitTested = true;
        return GRID_HTCELL;
    }

    bHitTested = false;
    return GRID_HTOUTSIDE;
}


void CGrid32Mgr::OnMouseWheel(int zDelta, UINT nFlags, int x, int y)
{
    // Handle mouse wheel scrolling here
}

void CGrid32Mgr::OnRButtonDown(UINT nFlags, int x, int y)
{
    GRIDNMHDR my_hdr;
    memset(&my_hdr, 0, sizeof(GRIDNMHDR));
    SendGridNotification(NM_RCLICK);
}

void CGrid32Mgr::OnRButtonUp(UINT nFlags, int x, int y)
{
    // Handle right mouse button up events here
    // This is often used in conjunction with WM_RBUTTONDOWN for drag-and-drop operations
}

bool CGrid32Mgr::IsCoordinateInClientArea(int x, int y)
{
    RECT r;
    GetClientRect(m_hWndGrid, &r);

    // Convert screen coordinates to client coordinates
    POINT pt = { x, y };
    ScreenToClient(m_hWndGrid, &pt);

    // Check if the converted coordinates are within the client area
    return (r.left <= pt.x && pt.x <= r.right && r.top <= pt.y && pt.y <= r.bottom);
}

void CGrid32Mgr::SetCursorBasedOnNcHitTest()
{
    switch (m_gridHitTest)
    {
    case GRID_HTCOLDIVIDER:
        SetCursor(LoadCursor(nullptr, IDC_SIZEWE));
        break;
    case GRID_HTROWDIVIDER:
        SetCursor(LoadCursor(nullptr, IDC_SIZENS));
        break;
    default:
        SetCursor(LoadCursor(nullptr, IDC_ARROW));
        break;
    }
}

void CGrid32Mgr::OnClear()
{
    // Clear the grid content
    // This might involve deleting all cells or rows/columns
}

void CGrid32Mgr::OnCopy()
{
    // Copy the selected cells to the clipboard
    // You'll need to implement logic to determine the selected cells and format the data for the clipboard
}

void CGrid32Mgr::OnCut()
{
    // Cut the selected cells to the clipboard
    // Similar to OnCopy, but also remove the cells from the grid
}

void CGrid32Mgr::OnPaste()
{
    // Paste the clipboard content into the grid
    // You'll need to parse the clipboard data and insert it into the grid appropriately
}

// Undo the last edit operation
void CGrid32Mgr::OnUndo()
{
    if (m_undoStack.empty())
        return;

    GridEditOperation op = m_undoStack.top();
    m_undoStack.pop();

    switch (op.type)
    {
    case EditOperationType::SetText:
        // Revert text change
        // Use SetCellText which records undo, so bypass recording here
    {
        // Temporary disable recording
        bool record = m_bUndoRecordEnabled;
        m_bUndoRecordEnabled = false;
        SetCellText(op.row, op.col, op.oldState.m_wsText.c_str());
        m_bUndoRecordEnabled = record;
    }
    break;

    case EditOperationType::SetFormat:
    {
        bool record = m_bUndoRecordEnabled;
        m_bUndoRecordEnabled = false;
        SetCellFormat(op.row, op.col, op.oldState.fontInfo);
        m_bUndoRecordEnabled = record;
    }
    break;

    case EditOperationType::SetFullCell:
    {
        bool record = m_bUndoRecordEnabled;
        m_bUndoRecordEnabled = false;
        SetCell(op.row, op.col, op.oldState);
        m_bUndoRecordEnabled = record;
    }
    break;

    case EditOperationType::Delete:
    {
        // Undo delete: recreate the cell
        bool record = m_bUndoRecordEnabled;
        m_bUndoRecordEnabled = false;
        SetCell(op.row, op.col, op.oldState);
        m_bUndoRecordEnabled = record;
    }
    break;

    case EditOperationType::Clipboard:
        // TODO: implement clipboard undo
        break;
    }

    // Push to redo stack
    m_redoStack.push(op);
    Invalidate();
    SetLastError(0);
}

// Redo the last undone operation
void CGrid32Mgr::OnRedo()
{
    if (m_redoStack.empty())
        return;

    GridEditOperation op = m_redoStack.top();
    m_redoStack.pop();

    switch (op.type)
    {
    case EditOperationType::SetText:
    {
        bool record = m_bUndoRecordEnabled;
        m_bUndoRecordEnabled = false;
        SetCellText(op.row, op.col, op.newState.m_wsText.c_str());
        m_bUndoRecordEnabled = record;
    }
    break;

    case EditOperationType::SetFormat:
    {
        bool record = m_bUndoRecordEnabled;
        m_bUndoRecordEnabled = false;
        SetCellFormat(op.row, op.col, op.newState.fontInfo);
        m_bUndoRecordEnabled = record;
    }
    break;

    case EditOperationType::SetFullCell:
    {
        bool record = m_bUndoRecordEnabled;
        m_bUndoRecordEnabled = false;
        SetCell(op.row, op.col, op.newState);
        m_bUndoRecordEnabled = record;
    }
    break;

    case EditOperationType::Delete:
    {
        bool record = m_bUndoRecordEnabled;
        m_bUndoRecordEnabled = false;
        // Redo delete: remove cell
        DeleteCell(op.row, op.col);
        m_bUndoRecordEnabled = record;
    }
    break;

    case EditOperationType::Clipboard:
        // TODO: implement clipboard redo
        break;
    }

    // After redo, push back onto undo stack
    m_undoStack.push(op);
    Invalidate();
    SetLastError(0);
}

size_t CGrid32Mgr::OnCanUndo()
{
    // Enable or disable UI undo based on stack state
    return m_undoStack.size();
}

size_t CGrid32Mgr::OnCanRedo()
{
    return m_redoStack.size();
}


void CGrid32Mgr::OnSetSelection(GRIDSELECTION* pGridSel) 
{
    if (pGridSel) {
        // Validate selection coordinates
        if (pGridSel->start.nRow < 0 || pGridSel->start.nCol < 0 ||
            pGridSel->end.nRow >= gcs.nHeight || pGridSel->end.nCol >= gcs.nWidth) {
            SetLastError(GRID_ERROR_INVALID_CELL_REFERENCE);
            return;
        }

        // Ensure start is before end
        if (pGridSel->start.nRow > pGridSel->end.nRow ||
            (pGridSel->start.nRow == pGridSel->end.nRow && pGridSel->start.nCol > pGridSel->end.nCol)) {
            std::swap(pGridSel->start, pGridSel->end);
        }

        m_selectionRect = *pGridSel;
        Invalidate();
        SetLastError(0);
    }
    else {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
    }
}

void CGrid32Mgr::OnGetSelection(GRIDSELECTION* pGridSel) 
{
    if (pGridSel) {
        *pGridSel = m_selectionRect;
        SetLastError(0);
    }
    else {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
    }
}

BOOL CGrid32Mgr::OnSetCurrentCell(LPCWSTR pwszRef, short nWhich)
{
    if (nWhich == SETBYNAME)
    {
        for (const auto& cellEntry : mapCells)
        {
            PGRIDCELL pCell = reinterpret_cast<PGRIDCELL>(cellEntry.second);
            if (pCell && _wcsicmp(pCell->m_wsName.c_str(), pwszRef) == 0)
            {
                UINT nRow = cellEntry.first.first;
                UINT nCol = cellEntry.first.second;
                SetCurrentCell(nRow, nCol);
                return TRUE;
            }
        }
    }
    else if (nWhich == SETBYCOORDINATE)
    {
        std::wstring wsCellRef = FormatCellReference(pwszRef);

        // Parse the formattedRef to get nRow and nCol
        // Assuming the format is "$COL$ROW"
        std::wstring colName = wsCellRef.substr(1, wsCellRef.find(L'$') - 1);
        std::wstring rowName = wsCellRef.substr(wsCellRef.find_last_of(L'$') + 1);

        UINT nCol = 0;
        for (wchar_t ch : colName)
        {
            nCol = nCol * 26 + (ch - L'A' + 1);
        }
        UINT nRow = std::stoi(rowName);

        auto it = mapCells.find({ nRow, nCol });
        if (it != mapCells.end() && it->second != nullptr)
        {
            SetCurrentCell(nRow, nCol);
            return TRUE;
        }
    }
    std::wstring wsCellRef = FormatCellReference(pwszRef);
    if (wsCellRef.length() == 0 && dwError > 0)
        return FALSE;


    return TRUE;
}

BOOL CGrid32Mgr::OnSetCurrentCell(UINT nRow, UINT nCol)
{
    if (nRow >= gcs.nHeight || nCol >= gcs.nWidth)
        return FALSE;

    PAGESTAT pageStat;
    CalculatePageStats(pageStat);

    m_currentCell.nRow = nRow;
    m_currentCell.nCol = nCol;

    if (!IsCellVisible(nRow, nCol))
    {
        m_visibleTopLeft.nRow = nRow - ((pageStat.end.nRow - pageStat.start.nRow) / 2);
        m_visibleTopLeft.nCol = nCol - ((pageStat.end.nCol - pageStat.start.nCol) / 2);
    }

    return TRUE;
}

// Function to trim whitespace from both ends of a string
std::wstring Trim(const std::wstring& str)
{
    const std::wstring whitespace = L" \t\n\r\f\v";
    size_t start = str.find_first_not_of(whitespace);
    size_t end = str.find_last_not_of(whitespace);

    return (start == std::wstring::npos || end == std::wstring::npos) ? L"" : str.substr(start, end - start + 1);
}

// Function to parse cell reference and return it in the format $COL$ROW
std::wstring CGrid32Mgr::FormatCellReference(LPCWSTR pwszRef)
{
    std::wstring input = Trim(pwszRef);
    if (input.empty()) {
        SetLastError(GRID_ERROR_INVALID_CELL_REFERENCE);
        return std::wstring();
    }

    std::wstring column;
    std::wstring row;

    // Parse the input to separate column and row
    bool foundColumn = false;
    for (wchar_t ch : input) {
        if (isalpha(ch)) {
            if (foundColumn) {
                SetLastError(GRID_ERROR_INVALID_CELL_REFERENCE);
                return std::wstring();
            }
            column += toupper(ch);
        }
        else if (isdigit(ch)) {
            foundColumn = true;
            row += ch;
        }
        else if (ch == ':' || ch == ',' || ch == '.' || ch == ';' || ch == ' ' || ch == '-') {
            // Ignore these characters
        }
        else {
            SetLastError(GRID_ERROR_INVALID_CELL_REFERENCE);
            return std::wstring();
        }
    }

    if (column.empty() || row.empty()) {
        SetLastError(GRID_ERROR_INVALID_CELL_REFERENCE);
        return std::wstring();
    }

    std::wstringstream result;
    result << L"$" << column << L"$" << row;

    return result.str();
}

bool CGrid32Mgr::SetCurrentCell(UINT nRow, UINT nCol)
{
    SetLastError(GRID_ERROR_NOT_IMPLEMENTED);
    return false;
}

BOOL CGrid32Mgr::Invalidate(bool bErasebkgnd)
{
    if (m_bRedraw)
        return ::InvalidateRect(m_hWndGrid, NULL, bErasebkgnd);

    return 0;
}

bool CGrid32Mgr::SetColumnWidth(size_t nCol, size_t newWidth) 
{
    if (nCol >= 0 && nCol < gcs.nWidth) {
        pColInfoArray[nCol].nWidth = newWidth;
        Invalidate();
        return true;
    }

    SetLastError(GRID_ERROR_OUT_OF_RANGE);
    return false;
}

size_t CGrid32Mgr::GetColumnWidth(size_t nCol)
{
    if (nCol >= 0 && nCol < gcs.nWidth)
        return pColInfoArray[nCol].nWidth;

    SetLastError(GRID_ERROR_OUT_OF_RANGE);
    return 0;
}

bool CGrid32Mgr::SetRowHeight(size_t nRow, size_t newHeight) 
{
    if (nRow >= 0 && nRow < gcs.nHeight) {
        pRowInfoArray[nRow].nHeight = newHeight;
        Invalidate();
        return true;
    }

    SetLastError(GRID_ERROR_OUT_OF_RANGE);
    return false;
}


size_t CGrid32Mgr::GetRowHeight(size_t nRow)
{
    if (nRow >= 0 && nRow < gcs.nHeight)
        return pRowInfoArray[nRow].nHeight;

    SetLastError(GRID_ERROR_OUT_OF_RANGE);
    return 0;
}

BOOL CGrid32Mgr::OnGetCellText(const GRIDPOINT& point, GRID_GETTEXT* text) {
    // Input validation
    if (!text || point.nRow >= gcs.nHeight || point.nCol >= gcs.nWidth) {
        SetLastError(GRID_ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    // Check if a cell exists at the given point
    PGRIDCELL pCell = GetCell(point.nRow, point.nCol);

    // If cell exists, copy its text to the provided buffer
    if (pCell) {
        // Check if buffer size is sufficient
        if (text->nLen < pCell->m_wsText.length() + 1) {
            // Set error (buffer too small)
            SetLastError(GRID_ERROR_INVALID_PARAMETER);
            return FALSE;
        }
        // Copy cell text to buffer
        wcscpy_s(text->wszBuff, text->nLen, pCell->m_wsText.c_str());
        SetLastError(0);
        return TRUE;
    }

    // If no cell exists, return empty string
    text->wszBuff[0] = L'\0';
    SetLastError(GRID_ERROR_NOT_EXIST);
    return TRUE;
}

void CGrid32Mgr::ShowEditControl()
{
    // This should NEVER happen because we fail the 
    // whole creation of the grid if m_hWndEdit is NULL in WM_CREATE
    if (!m_hWndEdit)  
    {
        return;
    }

    RECT currentCellRect;
    memset(&currentCellRect, 0, sizeof(RECT));
    GetCurrentCellRect(currentCellRect);
    InflateRect(&currentCellRect, -1, -1);
    // Set the position and size of the edit control
    MoveWindow(m_hWndEdit,
        currentCellRect.left,
        currentCellRect.top,
        currentCellRect.right - currentCellRect.left,
        currentCellRect.bottom - currentCellRect.top,
        TRUE);

    // Show the edit control
    BringWindowToTop(m_hWndEdit); 
    SendMessage(m_hWndEdit, EM_SCROLLCARET, 0, 0);
    ShowWindow(m_hWndEdit, SW_SHOW);
    SetFocus(m_hWndEdit);
}

void CGrid32Mgr::SendGridNotification(INT code, GRIDNMHDR* pNMHDRInfo)
{
    GRIDNMHDR nmhdr;
    memset(&nmhdr, 0, sizeof(GRIDNMHDR));
    if (pNMHDRInfo)
        nmhdr = *pNMHDRInfo;
    nmhdr.hwndFrom = m_hWndGrid;
    nmhdr.idFrom = GetWindowLongPtr(m_hWndGrid, GWLP_ID);
    nmhdr.code = code;
    if (!pNMHDRInfo)
    {
        GetMousePosition(nmhdr.pt);
        nmhdr.cellPt = code == NM_HOVER ? m_HoverCell : m_currentCell;
    }

    HWND parentWnd = GetParent(m_hWndGrid);
    if (parentWnd) {
        SendMessage(parentWnd, WM_NOTIFY, nmhdr.idFrom, (LPARAM)&nmhdr);
    }
}


void CGrid32Mgr::GetMousePosition(POINT& pt)
{
    GetCursorPos(&pt);
    ScreenToClient(m_hWndGrid, &pt);
}


void CGrid32Mgr::OnTimer(UINT_PTR idEvent)
{
    if (idEvent == HOVER_TIMER)
    {
        KillTimer(m_hWndGrid, idEvent);
        m_npHoverDelaySet = 0;
    }
}

bool CGrid32Mgr::IsOverColumnDivider(int x, int y) 
{
    size_t nPixWidth = GetActualRowHeaderWidth(), idx = m_visibleTopLeft.nCol;
    for (; nPixWidth < m_clientRect.right && idx < gcs.nWidth && (long long)nPixWidth < (long long)x + 5; nPixWidth += pColInfoArray[idx++].nWidth)
    {
        if ((long long)nPixWidth == (long long)x - 1 || (long long)nPixWidth == (long long)x || (long long)nPixWidth == (long long)x + 1)
        {
            m_nToBeSized = idx - 1;
            return true;
        }
    }

    m_nToBeSized = ~0;
    return false;
}

bool CGrid32Mgr::IsOverRowDivider(int x, int y) 
{
    size_t nPixHeight = GetActualColHeaderHeight(), idx = m_visibleTopLeft.nRow;
    for (; nPixHeight < m_clientRect.bottom && idx < gcs.nHeight && (long long)nPixHeight < (long long)y + 5; nPixHeight += pRowInfoArray[idx++].nHeight)
    {
        if ((long long)nPixHeight == (long long)y - 1 || (long long)nPixHeight == (long long)y || (long long)nPixHeight == (long long)y + 1)
        {
            m_nToBeSized = idx - 1;
            return true;
        }
    }

    m_nToBeSized = ~0;
    return false;
}

long long CGrid32Mgr::GetColumnStartPos(UINT nCol)
{
    long long startPos = GetActualRowHeaderWidth();

    // If nCol is the first visible column, return the row header width
    if (nCol == m_visibleTopLeft.nCol)
        return startPos;

    if (nCol >= gcs.nWidth)
        nCol = (UINT)(gcs.nWidth - 1);
    // Sum up the widths of all visible columns up to nCol
    for (UINT col = m_visibleTopLeft.nCol; col < nCol; ++col)
    {
        startPos += (signed)pColInfoArray[col].nWidth;
    }

    return startPos;
}

long long CGrid32Mgr::GetRowStartPos(UINT nRow)
{
    long long startPos = GetActualColHeaderHeight();

    // If nRow is the first visible row, return the column header height
    if (nRow == m_visibleTopLeft.nRow)
        return startPos;

    if (nRow >= gcs.nHeight)
        nRow = (UINT)(gcs.nHeight - 1);

    // Sum up the heights of all visible rows up to nRow
    for (UINT row = m_visibleTopLeft.nRow; row < nRow; ++row)
    {
        startPos += (signed)pRowInfoArray[row].nHeight;
    }

    return startPos;
}


bool CGrid32Mgr::IsCellFontDefault(GRIDCELL& gc)
{
    const FONTINFO& cellFont = gc.fontInfo;
    const FONTINFO& defaultFont = m_defaultGridCell.fontInfo;

    return (cellFont.m_wsFontFace == defaultFont.m_wsFontFace &&
        cellFont.m_fPointSize == defaultFont.m_fPointSize &&
        cellFont.m_clrTextColor == defaultFont.m_clrTextColor &&
        cellFont.bItalic == defaultFont.bItalic &&
        cellFont.bUnderline == defaultFont.bUnderline &&
        cellFont.bStrikeThrough == defaultFont.bStrikeThrough &&
        cellFont.bWeight == defaultFont.bWeight);
}

bool CGrid32Mgr::AreGridPointsEqual(const GRIDPOINT& point1, const GRIDPOINT& point2)
{
    return (point1.nRow == point2.nRow && point1.nCol == point2.nCol);
}


bool CGrid32Mgr::IsCellPointInsideSelection(const GRIDPOINT& gridPt)
{
    // Normalize the selection rectangle to ensure start is always the top-left
    // and end is always the bottom-right.
    UINT startRow = min(m_selectionRect.start.nRow, m_selectionRect.end.nRow);
    UINT endRow = max(m_selectionRect.start.nRow, m_selectionRect.end.nRow);
    UINT startCol = min(m_selectionRect.start.nCol, m_selectionRect.end.nCol);
    UINT endCol = max(m_selectionRect.start.nCol, m_selectionRect.end.nCol);

    // Check if the gridPt is inside the normalized selection rectangle
    return (gridPt.nRow >= startRow && gridPt.nRow <= endRow &&
        gridPt.nCol >= startCol && gridPt.nCol <= endCol);
}

bool CGrid32Mgr::IsSelectionUnicellular()
{
    return AreGridPointsEqual(m_selectionRect.start, m_selectionRect.end);
}

void CGrid32Mgr::ConstrainMaxPosition()
{
    size_t nPos = 0, idx = 0;
    if (m_selectionRect.end.nCol >= (UINT)-100)  //decremented once too many
        m_selectionRect.end.nCol = 0;
    else if (m_selectionRect.end.nCol >= gcs.nWidth)
        m_selectionRect.end.nCol = (UINT)(gcs.nWidth - 1);
    
    if (m_selectionRect.end.nRow >= (UINT)-100)
        m_selectionRect.end.nRow = 0;
    else if (m_selectionRect.end.nRow >= gcs.nHeight)
        m_selectionRect.end.nRow = (UINT)(gcs.nHeight - 1);

    for (idx = gcs.nWidth - 1;
        idx > 0 && nPos < (m_clientRect.right * 2) / 3;
        nPos += pColInfoArray[idx--].nWidth);

    if (m_visibleTopLeft.nCol >= (UINT)-100)
        m_visibleTopLeft.nCol = 0;
    else if (m_visibleTopLeft.nCol > idx)
        m_visibleTopLeft.nCol = (UINT)(idx);

    for (idx = gcs.nHeight - 1;
        idx > 0 && nPos < (m_clientRect.bottom * 2) / 3;
        nPos += pRowInfoArray[idx--].nHeight);

    if (m_visibleTopLeft.nRow >= (UINT)-100)
        m_visibleTopLeft.nRow = 0;
    else if (m_visibleTopLeft.nRow > idx)
        m_visibleTopLeft.nRow = (UINT)(idx);

}


void CGrid32Mgr::NormalizeSelectionRect(GRIDSELECTION& selection)
{
    // Create a new GRIDPOINT to hold the normalized start and end points
    GRIDPOINT normalizedStart, normalizedEnd;

    // Normalize the rows
    if (selection.start.nRow <= selection.end.nRow)
    {
        normalizedStart.nRow = selection.start.nRow;
        normalizedEnd.nRow = selection.end.nRow;
    }
    else
    {
        normalizedStart.nRow = selection.end.nRow;
        normalizedEnd.nRow = selection.start.nRow;
    }

    // Normalize the columns
    if (selection.start.nCol <= selection.end.nCol)
    {
        normalizedStart.nCol = selection.start.nCol;
        normalizedEnd.nCol = selection.end.nCol;
    }
    else
    {
        normalizedStart.nCol = selection.end.nCol;
        normalizedEnd.nCol = selection.start.nCol;
    }

    // Assign the normalized points back to the selection
    selection.start = normalizedStart;
    selection.end = normalizedEnd;
}

void CGrid32Mgr::OnIncrementCell(UINT direction, GRIDPOINT* pGridPoint)
{
    switch (direction)
    {
    case INCREMENT_SINGLE:
        IncrementSelectedCell(pGridPoint->nRow, pGridPoint->nCol);
        break;
    case INCREMENT_PAGE:
        IncrementSelectedPage(pGridPoint->nRow, pGridPoint->nCol);
        break;
    case INCREMENT_EDGE:
        IncrementSelectEdge(pGridPoint->nRow, pGridPoint->nCol);
        break;
    }
}


size_t CGrid32Mgr::OnEnumCells(GRIDCELL* gcArray, UINT numElements)
{
    // If gcArray is nullptr, return the number of cells in the map
    if (!gcArray)
        return mapCells.size();

    auto itor = mapCells.begin();
    size_t nCount = 0;

    // Determine the number of elements to copy
    size_t elementsToCopy = (numElements <= mapCells.size()) ? numElements : mapCells.size();

    // Copy elements from the map to the provided array
    for (; nCount < elementsToCopy && itor != mapCells.end(); ++itor, ++nCount)
        CopyGridCell(gcArray[nCount], *(itor->second));

    return nCount;
}

void CGrid32Mgr::OnFillCells(WPARAM wParam, const GCFILLSTRUCT& fillStruct)
{
}

void CGrid32Mgr::OnSortCells(WPARAM wParam, const GCSORTSTRUCT& sortStruct)
{
}

void CGrid32Mgr::OnFilterCells(WPARAM wParam, const GCFILTERSTRUCT& filterStruct)
{
}

// Helper to record and apply edit operations
void CGrid32Mgr::RecordUndoOperation(const GridEditOperation& op)
{
	if (!m_bUndoRecordEnabled)
		return; // If recording is disabled, do not record the operation
    m_undoStack.push(op);
    // Clear redo stack
    while (!m_redoStack.empty()) {
        m_redoStack.pop();
    }
}


void CGrid32Mgr::CopyGridCell(GRIDCELL& dest, GRIDCELL& src)
{
    // Copy the base cell attributes
    dest.fontInfo.m_wsFontFace = src.fontInfo.m_wsFontFace;
    dest.fontInfo.m_fPointSize = src.fontInfo.m_fPointSize;
    dest.fontInfo.m_clrTextColor = src.fontInfo.m_clrTextColor;
    dest.fontInfo.bItalic = src.fontInfo.bItalic;
    dest.fontInfo.bUnderline = src.fontInfo.bUnderline;
    dest.fontInfo.bStrikeThrough = src.fontInfo.bStrikeThrough;
    dest.fontInfo.bWeight = src.fontInfo.bWeight;

    dest.clrBackground = src.clrBackground;
    dest.m_clrBorderColor = src.m_clrBorderColor;
    dest.m_nBorderWidth = src.m_nBorderWidth;
    dest.penStyle = src.penStyle;
    dest.justification = src.justification;
    dest.m_wsName = src.m_wsName;

    // Copy the specific GRIDCELL attributes
    dest.m_wsText = src.m_wsText;
    dest.mergeRange = src.mergeRange;
}
