#pragma once

// Tile layout for the PhotoGalleryGrid control. Pure geometry: no HWND, no
// GDI, so the math can be tested on its own. All coordinates are content
// coordinates in physical pixels (y = 0 is the top of the scrolled content).

#include <vector>
#include <memory>
#include <functional>
#include <climits>

// Content coordinates are ints (the scroll position is one), so the content
// stops at this height, with room left for a viewport and a prefetch ring
// below it. Rows past it pile up at the edge and cannot be scrolled to.
const int kPGMaxContent = INT_MAX / 2;

inline int PGClampContent(long long y)
{
    return y > kPGMaxContent ? kPGMaxContent : (y < 0 ? 0 : (int)y);
}

// Geometry resolved from PGGTILEMETRICS for the current client width and DPI.
// In the stretch modes column widths differ by at most one pixel; ColLeft
// spreads the remainder so the right edge stays flush.
struct PGGeometry
{
    int  nColumns = 1;
    int  cellCx = 1;          // nominal cell width
    int  cellCy = 1;          // row height
    int  cxGap = 0, cyGap = 0;
    int  cxMargin = 0, cyMargin = 0;
    int  avail = 1;           // width the columns share (stretch modes)
    bool stretch = true;

    int ColLeft(int c) const;
    int RowTop(int r) const { return PGClampContent(cyMargin + (long long)r * RowPitch()); }
    int RowPitch() const { return cellCy + cyGap; }
    int SpanWidth(int c, int cx) const { return ColLeft(c + cx) - cxGap - ColLeft(c); }
    int SpanHeight(int cy) const { return cy * cellCy + (cy - 1) * cyGap; }
    int ColumnAt(int x) const;          // -1 in a gap or outside the grid
    int RowIndexAt(int y) const;        // row whose pitch band holds y (gap included), clamped at 0
    int ContentHeight(int nRows) const; // clamped to kPGMaxContent
};

struct PGSpan
{
    BYTE cx = 1, cy = 1;
};

class CGalleryLayout
{
public:
    virtual ~CGalleryLayout() = default;

    // getSpan is only called when NeedsSpans() is true.
    virtual void Arrange(const PGGeometry& geo, int nItems, const std::function<PGSpan(int)>& getSpan) = 0;
    virtual bool NeedsSpans() const = 0;
    virtual int  ContentHeight() const = 0;
    virtual bool ItemRect(int i, RECT& rc) const = 0;
    // Items intersecting [y0, y1), ascending.
    virtual void ItemsInBand(int y0, int y1, std::vector<int>& out) const = 0;
    virtual int  HitTest(POINT pt) const = 0;
    // Arrow keys, Home/End, PgUp/PgDn. i < 0 means "nothing focused yet".
    virtual int  Navigate(int i, UINT vk, int cyPage) const = 0;
};

std::unique_ptr<CGalleryLayout> CreateUniformLayout();
std::unique_ptr<CGalleryLayout> CreateMixedLayout(bool bDense);
