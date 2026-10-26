#include "pch.h"
#include "PhotoGalleryThumbs.h"

bool PixelsFromHBITMAP(HBITMAP hbm, PGPixels& out)
{
    BITMAP bm{};
    if (!hbm || !GetObject(hbm, sizeof(bm), &bm) || bm.bmWidth <= 0 || bm.bmHeight == 0)
        return false;

    const int w = bm.bmWidth, h = bm.bmHeight < 0 ? -bm.bmHeight : bm.bmHeight;
    BITMAPINFO bi{};
    bi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bi.bmiHeader.biWidth = w;
    bi.bmiHeader.biHeight = -h;     // top-down
    bi.bmiHeader.biPlanes = 1;
    bi.bmiHeader.biBitCount = 32;
    bi.bmiHeader.biCompression = BI_RGB;

    PGPixels px;
    px.w = w;
    px.h = h;
    px.px.resize((size_t)w * h);
    HDC hdc = GetDC(NULL);
    int lines = GetDIBits(hdc, hbm, 0, h, px.px.data(), &bi, DIB_RGB_COLORS);
    ReleaseDC(NULL, hdc);
    if (lines != h)
        return false;

    // GDI never writes alpha, so an all-zero alpha channel means "opaque".
    bool bHasAlpha = false;
    if (bm.bmBitsPixel == 32)
    {
        for (UINT32 p : px.px)
        {
            if (p & 0xFF000000)
            {
                bHasAlpha = true;
                break;
            }
        }
    }
    if (!bHasAlpha)
    {
        for (UINT32& p : px.px)
            p |= 0xFF000000;
    }
    out = std::move(px);
    return true;
}

bool ScalePixels(const PGPixels& in, int w, int h, PGPixels& out)
{
    if (in.empty() || w <= 0 || h <= 0)
        return false;

    PGPixels res;
    res.w = w;
    res.h = h;
    res.px.assign((size_t)w * h, 0);
    {
        Gdiplus::Bitmap src(in.w, in.h, in.w * 4, PixelFormat32bppPARGB, (BYTE*)in.px.data());
        Gdiplus::Bitmap dst(w, h, w * 4, PixelFormat32bppPARGB, (BYTE*)res.px.data());
        Gdiplus::Graphics g(&dst);
        g.SetCompositingMode(Gdiplus::CompositingModeSourceCopy);
        g.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
        g.SetPixelOffsetMode(Gdiplus::PixelOffsetModeHalf);
        Gdiplus::ImageAttributes ia;
        ia.SetWrapMode(Gdiplus::WrapModeTileFlipXY);    // no dark fringe at the edges
        if (g.DrawImage(&src, Gdiplus::Rect(0, 0, w, h), 0, 0, in.w, in.h, Gdiplus::UnitPixel, &ia) != Gdiplus::Ok)
            return false;
    }
    out = std::move(res);
    return true;
}

double ScaleNeeded(int w, int h, int cx, int cy, bool bFit)
{
    if (w <= 0 || h <= 0)
        return 1e9;
    double sx = (double)cx / w, sy = (double)cy / h;
    return bFit ? (std::min)(sx, sy) : (std::max)(sx, sy);
}

void PGThumb::ResetSource()
{
    src = PGPixels();
    state = Empty;
    bProvisional = bNative = bHostAsked = false;
    requestedEdge = 0;
    uFailReason = 0;
    retryTick = 0;
    nRetries = 0;
    ++epoch;
    DropDisplay();
}

PGThumb* CThumbCache::Find(ULONGLONG key)
{
    auto it = m_map.find(key);
    return it == m_map.end() ? nullptr : &it->second;
}

PGThumb& CThumbCache::Get(ULONGLONG key)
{
    return m_map[key];
}

void CThumbCache::DropAllDisplays()
{
    for (auto& kv : m_map)
        kv.second.DropDisplay();
}

void CThumbCache::ResetAll()
{
    for (auto& kv : m_map)
    {
        if (kv.second.nPins > 0)
        {
            kv.second.DropDisplay();
            continue;
        }
        kv.second.ResetSource();
    }
}

void CThumbCache::ResetFailed()
{
    for (auto& kv : m_map)
    {
        if (kv.second.state == PGThumb::Failed && kv.second.nPins == 0)
            kv.second.ResetSource();
    }
}

size_t CThumbCache::TotalBytes() const
{
    size_t total = 0;
    for (const auto& kv : m_map)
        total += kv.second.Bytes();
    return total;
}

void CThumbCache::Trim(DWORD frameNow, bool bForce)
{
    DWORD now = GetTickCount();
    if (!bForce && now - m_lastTrimTick < 250)
        return;
    m_lastTrimTick = now;

    // Entries that hold nothing are cheap but not free; drop the stale ones
    // once there are a lot of them (owner-data galleries can touch many keys).
    if (m_map.size() > 20000)
    {
        for (auto it = m_map.begin(); it != m_map.end();)
        {
            const PGThumb& t = it->second;
            if (t.Bytes() == 0 && t.nPins == 0 && t.state != PGThumb::Pending && t.lastUsed != frameNow)
                it = m_map.erase(it);
            else
                ++it;
        }
    }

    size_t total = TotalBytes();
    if (total <= m_budget)
        return;

    std::vector<std::pair<DWORD, ULONGLONG>> lru;
    lru.reserve(m_map.size());
    for (const auto& kv : m_map)
        if (kv.second.lastUsed != frameNow && kv.second.Bytes() > 0)
            lru.emplace_back(kv.second.lastUsed, kv.first);
    std::sort(lru.begin(), lru.end());

    // Displays first (cheap to rebuild), then unpinned sources.
    for (const auto& e : lru)
    {
        if (total <= m_budget)
            return;
        PGThumb& t = m_map[e.second];
        if (t.disp)
        {
            total -= (size_t)t.dispW * t.dispH * 4;
            t.DropDisplay();
        }
    }
    for (const auto& e : lru)
    {
        if (total <= m_budget)
            return;
        PGThumb& t = m_map[e.second];
        if (t.nPins == 0 && !t.src.empty())
        {
            total -= t.src.bytes();
            t.ResetSource();
        }
    }
}
