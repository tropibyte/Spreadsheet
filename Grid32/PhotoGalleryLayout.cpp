#include "pch.h"
#include "photogallery.h"
#include "PhotoGalleryLayout.h"

// ---- PGGeometry ---------------------------------------------------------------

int PGGeometry::ColLeft(int c) const
{
    if (stretch)
        return cxMargin + MulDiv(c, avail + cxGap, nColumns);
    return cxMargin + c * (cellCx + cxGap);
}

int PGGeometry::ColumnAt(int x) const
{
    if (x < cxMargin || x >= ColLeft(nColumns) - cxGap)
        return -1;

    int c = stretch ? MulDiv(x - cxMargin, nColumns, avail + cxGap) : (x - cxMargin) / (cellCx + cxGap);
    c = (std::max)(0, (std::min)(c, nColumns - 1));
    while (c > 0 && ColLeft(c) > x)
        --c;
    while (c < nColumns - 1 && ColLeft(c + 1) <= x)
        ++c;
    return x < ColLeft(c) + SpanWidth(c, 1) ? c : -1;
}

int PGGeometry::RowIndexAt(int y) const
{
    if (y < cyMargin)
        return 0;
    return (y - cyMargin) / RowPitch();
}

int PGGeometry::ContentHeight(int nRows) const
{
    if (nRows <= 0)
        return 0;
    return 2 * cyMargin + nRows * cellCy + (nRows - 1) * cyGap;
}

namespace {

int NearestInRow(const std::vector<int>& candidates, int iExclude, int cx,
                 const std::function<void(int, RECT&)>& rectOf)
{
    int best = -1;
    long bestDist = MAXLONG;
    for (int j : candidates)
    {
        if (j == iExclude)
            continue;
        RECT rc;
        rectOf(j, rc);
        long dist = (cx < rc.left) ? rc.left - cx : (cx >= rc.right ? cx - rc.right + 1 : 0);
        if (dist < bestDist || (dist == bestDist && j < best))
        {
            bestDist = dist;
            best = j;
        }
    }
    return best;
}

// ---- Uniform: every tile one cell; everything is arithmetic ---------------------

class CUniformLayout : public CGalleryLayout
{
    PGGeometry m_geo;
    int m_nItems = 0;

    int Rows() const { return (m_nItems + m_geo.nColumns - 1) / m_geo.nColumns; }

public:
    void Arrange(const PGGeometry& geo, int nItems, const std::function<PGSpan(int)>&) override
    {
        m_geo = geo;
        m_nItems = (std::max)(0, nItems);
    }

    bool NeedsSpans() const override { return false; }

    int ContentHeight() const override { return m_geo.ContentHeight(Rows()); }

    bool ItemRect(int i, RECT& rc) const override
    {
        if (i < 0 || i >= m_nItems)
            return false;
        int r = i / m_geo.nColumns, c = i % m_geo.nColumns;
        rc.left = m_geo.ColLeft(c);
        rc.top = m_geo.RowTop(r);
        rc.right = rc.left + m_geo.SpanWidth(c, 1);
        rc.bottom = rc.top + m_geo.cellCy;
        return true;
    }

    void ItemsInBand(int y0, int y1, std::vector<int>& out) const override
    {
        out.clear();
        if (y1 <= y0 || m_nItems == 0)
            return;
        int r0 = m_geo.RowIndexAt(y0);
        int r1 = (std::min)(m_geo.RowIndexAt(y1 - 1), Rows() - 1);
        int iEnd = (std::min)(m_nItems, (r1 + 1) * m_geo.nColumns);
        for (int i = r0 * m_geo.nColumns; i < iEnd; ++i)
            out.push_back(i);
    }

    int HitTest(POINT pt) const override
    {
        int c = m_geo.ColumnAt(pt.x);
        if (c < 0 || pt.y < m_geo.cyMargin)
            return -1;
        int r = m_geo.RowIndexAt(pt.y);
        if (pt.y >= m_geo.RowTop(r) + m_geo.cellCy)
            return -1;  // in the gap below the row
        int i = r * m_geo.nColumns + c;
        return i < m_nItems ? i : -1;
    }

    int Navigate(int i, UINT vk, int cyPage) const override
    {
        if (m_nItems == 0)
            return -1;
        if (i < 0 || i >= m_nItems)
            return 0;
        const int cols = m_geo.nColumns;
        const int last = m_nItems - 1;
        const int rowsPerPage = (std::max)(1, cyPage / (std::max)(1, m_geo.RowPitch()));
        switch (vk)
        {
        case VK_LEFT:  return (std::max)(0, i - 1);
        case VK_RIGHT: return (std::min)(last, i + 1);
        case VK_UP:    return i - cols >= 0 ? i - cols : i;
        case VK_DOWN:
            if (i + cols <= last)
                return i + cols;
            return (i / cols < last / cols) ? last : i;   // short last row: land on the last tile
        case VK_HOME:  return 0;
        case VK_END:   return last;
        case VK_PRIOR:
        {
            int t = i - rowsPerPage * cols;
            return t >= 0 ? t : i % cols;
        }
        case VK_NEXT:  return (std::min)(last, i + rowsPerPage * cols);
        }
        return i;
    }
};

// ---- Mixed: tiles span cx x cy cells, packed in reading order -----------------
// Sparse packing (the default) never places an item before the previous one,
// so reading order is preserved and holes may remain. Dense packing back-fills
// holes, like CSS grid's "auto-flow: dense".

class CMixedLayout : public CGalleryLayout
{
    struct Placed { int row, col; BYTE cx, cy; };

    bool m_bDense;
    PGGeometry m_geo;
    std::vector<Placed> m_placed;
    int m_nRows = 0;
    // Items covering each row, compressed: m_rowItems[m_rowStart[r] .. m_rowStart[r + 1]).
    std::vector<int> m_rowStart;
    std::vector<int> m_rowItems;

public:
    explicit CMixedLayout(bool bDense) : m_bDense(bDense) {}

    void Arrange(const PGGeometry& geo, int nItems, const std::function<PGSpan(int)>& getSpan) override
    {
        m_geo = geo;
        const int cols = geo.nColumns;
        nItems = (std::max)(0, nItems);
        m_placed.assign(nItems, Placed{ 0, 0, 1, 1 });

        std::vector<BYTE> occ;  // rows * cols, grows as needed
        auto cell = [&](int r, int c) -> BYTE& {
            size_t need = (size_t)(r + 1) * cols;
            if (occ.size() < need)
                occ.resize(need, 0);
            return occ[(size_t)r * cols + c];
        };
        auto fits = [&](int r, int c, int cx, int cy) {
            for (int rr = r; rr < r + cy; ++rr)
                for (int cc = c; cc < c + cx; ++cc)
                    if ((size_t)rr * cols + cc < occ.size() && occ[(size_t)rr * cols + cc])
                        return false;
            return true;
        };

        int curR = 0, curC = 0;     // sparse cursor
        int firstOpenRow = 0;       // dense: rows above this are full
        m_nRows = 0;
        for (int i = 0; i < nItems; ++i)
        {
            PGSpan s = getSpan ? getSpan(i) : PGSpan{};
            int cx = (std::max)(1, (std::min)((int)s.cx, (std::min)(cols, PGG_MAXSPAN)));
            int cy = (std::max)(1, (std::min)((int)s.cy, PGG_MAXSPAN));

            int startR = m_bDense ? firstOpenRow : curR;
            int startC = m_bDense ? 0 : curC;
            int pr = -1, pc = -1;
            for (int r = startR; pr < 0; ++r)
            {
                for (int c = (r == startR ? startC : 0); c + cx <= cols; ++c)
                {
                    if (fits(r, c, cx, cy))
                    {
                        pr = r;
                        pc = c;
                        break;
                    }
                }
            }
            for (int rr = pr; rr < pr + cy; ++rr)
                for (int cc = pc; cc < pc + cx; ++cc)
                    cell(rr, cc) = 1;
            m_placed[i] = Placed{ pr, pc, (BYTE)cx, (BYTE)cy };
            m_nRows = (std::max)(m_nRows, pr + cy);
            curR = pr;
            curC = pc + cx;
            if (m_bDense)
            {
                auto rowFull = [&](int r) {
                    if ((size_t)(r + 1) * cols > occ.size())
                        return false;
                    for (int c = 0; c < cols; ++c)
                        if (!occ[(size_t)r * cols + c])
                            return false;
                    return true;
                };
                while (rowFull(firstOpenRow))
                    ++firstOpenRow;
            }
        }

        // Row index.
        m_rowStart.assign(m_nRows + 1, 0);
        for (const Placed& p : m_placed)
            for (int r = p.row; r < p.row + p.cy; ++r)
                ++m_rowStart[r + 1];
        for (int r = 0; r < m_nRows; ++r)
            m_rowStart[r + 1] += m_rowStart[r];
        m_rowItems.assign(m_rowStart[m_nRows], 0);
        std::vector<int> fill(m_rowStart.begin(), m_rowStart.end() - 1);
        for (int i = 0; i < nItems; ++i)
            for (int r = m_placed[i].row; r < m_placed[i].row + m_placed[i].cy; ++r)
                m_rowItems[fill[r]++] = i;
    }

    bool NeedsSpans() const override { return true; }

    int ContentHeight() const override { return m_geo.ContentHeight(m_nRows); }

    bool ItemRect(int i, RECT& rc) const override
    {
        if (i < 0 || i >= (int)m_placed.size())
            return false;
        const Placed& p = m_placed[i];
        rc.left = m_geo.ColLeft(p.col);
        rc.top = m_geo.RowTop(p.row);
        rc.right = rc.left + m_geo.SpanWidth(p.col, p.cx);
        rc.bottom = rc.top + m_geo.SpanHeight(p.cy);
        return true;
    }

    void ItemsInBand(int y0, int y1, std::vector<int>& out) const override
    {
        out.clear();
        if (y1 <= y0 || m_nRows == 0 || m_geo.RowIndexAt(y0) >= m_nRows)
            return;     // nothing, or a band entirely below the content
        int r0 = m_geo.RowIndexAt(y0);
        int r1 = (std::min)(m_geo.RowIndexAt(y1 - 1), m_nRows - 1);
        for (int r = r0; r <= r1; ++r)
        {
            for (int k = m_rowStart[r]; k < m_rowStart[r + 1]; ++k)
            {
                int i = m_rowItems[k];
                // Report each item once: at its first row inside the band.
                if ((std::max)(m_placed[i].row, r0) == r)
                    out.push_back(i);
            }
        }
        std::sort(out.begin(), out.end());
    }

    int HitTest(POINT pt) const override
    {
        if (m_nRows == 0 || pt.y < m_geo.cyMargin)
            return -1;
        int r = m_geo.RowIndexAt(pt.y);
        if (r >= m_nRows)
            return -1;
        for (int k = m_rowStart[r]; k < m_rowStart[r + 1]; ++k)
        {
            RECT rc;
            ItemRect(m_rowItems[k], rc);
            if (PtInRect(&rc, pt))
                return m_rowItems[k];
        }
        return -1;
    }

    int Navigate(int i, UINT vk, int cyPage) const override
    {
        const int n = (int)m_placed.size();
        if (n == 0)
            return -1;
        if (i < 0 || i >= n)
            return 0;

        RECT cur;
        ItemRect(i, cur);
        const int cx = (cur.left + cur.right) / 2;
        auto rectOf = [this](int j, RECT& rc) { ItemRect(j, rc); };
        auto rowItems = [this](int r) {
            return std::vector<int>(m_rowItems.begin() + m_rowStart[r], m_rowItems.begin() + m_rowStart[r + 1]);
        };
        auto nearestFrom = [&](int r, int step) {
            for (; r >= 0 && r < m_nRows; r += step)
            {
                int j = NearestInRow(rowItems(r), i, cx, rectOf);
                if (j >= 0)
                    return j;
            }
            return i;
        };

        switch (vk)
        {
        case VK_LEFT:  return (std::max)(0, i - 1);
        case VK_RIGHT: return (std::min)(n - 1, i + 1);
        case VK_HOME:  return 0;
        case VK_END:   return n - 1;
        case VK_UP:    return nearestFrom(m_placed[i].row - 1, -1);
        case VK_DOWN:  return nearestFrom(m_placed[i].row + m_placed[i].cy, +1);
        case VK_PRIOR:
        case VK_NEXT:
        {
            int y = (cur.top + cur.bottom) / 2 + (vk == VK_NEXT ? cyPage : -cyPage);
            int r = (std::max)(0, (std::min)(m_geo.RowIndexAt((std::max)(0, y)), m_nRows - 1));
            int j = NearestInRow(rowItems(r), -1, cx, rectOf);
            return j >= 0 ? j : i;
        }
        }
        return i;
    }
};

} // namespace

std::unique_ptr<CGalleryLayout> CreateUniformLayout()
{
    return std::make_unique<CUniformLayout>();
}

std::unique_ptr<CGalleryLayout> CreateMixedLayout(bool bDense)
{
    return std::make_unique<CMixedLayout>(bDense);
}
