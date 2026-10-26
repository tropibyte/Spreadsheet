#pragma once

// Thumbnail pixels and the thumbnail cache for the PhotoGalleryGrid control.
//
// Pixels live in plain memory as premultiplied BGRA (GDI+'s PARGB layout),
// never as one GDI handle per photo: the default per-process GDI quota is
// 10,000 handles. Only tiles actually drawn get a GDI+ CachedBitmap, and the
// cache trims itself to a byte budget, least recently drawn first.

#include <vector>
#include <memory>
#include <string>
#include <unordered_map>

struct PGPixels
{
    std::vector<UINT32> px;   // premultiplied BGRA, top-down, stride = w * 4
    int w = 0, h = 0;

    bool empty() const { return px.empty() || w <= 0 || h <= 0; }
    size_t bytes() const { return px.size() * sizeof(UINT32); }
};

// Copies any HBITMAP into premultiplied BGRA. A 32bpp bitmap whose alpha is
// not all zero is taken as premultiplied; anything else becomes opaque. The
// bitmap must not be selected into a DC. Safe on any thread.
bool PixelsFromHBITMAP(HBITMAP hbm, PGPixels& out);

// High-quality resample with GDI+. UI thread.
bool ScalePixels(const PGPixels& in, int w, int h, PGPixels& out);

// Scale (>1 = enlarge) needed to draw w x h pixels into a cx x cy tile:
// covering it (center-crop) or fitting inside it (letterbox).
double ScaleNeeded(int w, int h, int cx, int cy, bool bFit);

struct PGThumb
{
    enum State : BYTE { Empty, Pending, Ready, Failed };

    PGPixels src;
    State    state = Empty;
    bool     bProvisional = false;  // host said a sharper one may follow
    bool     bNative = false;       // src is the whole image: asking again cannot help
    bool     bHostAsked = false;    // PGGS_AUTOLOAD: the host-first GETDISPINFO was tried
    int      nPins = 0;             // internal-mode items whose host pixels live here
    UINT     uFailReason = 0;
    UINT     epoch = 0;             // bumped on reset; stale loader results are dropped
    int      requestedEdge = 0;     // tile long edge the current src was produced for
    DWORD    retryTick = 0;         // a cloud preview is being fetched: not before this tick
    BYTE     nRetries = 0;
    std::wstring path;              // PGGS_AUTOLOAD source file

    std::unique_ptr<Gdiplus::CachedBitmap> disp;   // src scaled to one tile, device format
    int      dispW = 0, dispH = 0;
    bool     dispFit = false, dispLowQ = false;
    DWORD    lastUsed = 0;          // paint frame that last drew it

    size_t Bytes() const { return src.bytes() + (disp ? (size_t)dispW * dispH * 4 : 0); }
    void DropDisplay() { disp.reset(); dispW = dispH = 0; dispLowQ = false; }
    void ResetSource();
};

class CThumbCache
{
    std::unordered_map<ULONGLONG, PGThumb> m_map;
    size_t m_budget = 256u * 1024 * 1024;
    DWORD  m_lastTrimTick = 0;

public:
    PGThumb* Find(ULONGLONG key);
    PGThumb& Get(ULONGLONG key);
    void Erase(ULONGLONG key) { m_map.erase(key); }
    void Clear() { m_map.clear(); }
    void DropAllDisplays();
    void ResetAll();                // forget every source except pinned host pixels
    void ResetFailed();             // let failed loads try again (e.g. settings changed)
    void SetBudget(size_t bytes) { m_budget = bytes; }
    size_t Budget() const { return m_budget; }
    size_t Count() const { return m_map.size(); }
    size_t TotalBytes() const;
    // Evict least-recently-drawn entries until under budget. Never touches
    // what frameNow drew. Rate-limited unless bForce.
    void Trim(DWORD frameNow, bool bForce = false);
};
