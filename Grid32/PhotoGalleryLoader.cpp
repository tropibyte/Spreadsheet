#include "pch.h"
#include "photogallery.h"
#include "PhotoGalleryLoader.h"
#include <deque>
#include <unordered_set>
#include <mutex>
#include <condition_variable>
#include <new>
#include <wincodec.h>
#include <shobjidl.h>
#include <wrl/client.h>

#pragma comment(lib, "windowscodecs.lib")

using Microsoft::WRL::ComPtr;

struct CThumbLoader::Shared
{
    std::mutex                    mtx;
    std::condition_variable       cv;
    std::deque<PGLoadJob>         queue;
    std::unordered_set<ULONGLONG> inFlight;
    std::vector<PGLoadResult>     done;
    bool                          bPostPending = false;
    bool                          bStop = false;
    HWND                          hwnd = NULL;
    UINT                          flags = 0;
};

namespace {

// Cloud placeholders (OneDrive Files On-Demand and friends). Reading the
// attributes does not download anything; opening the file does.
const DWORD kCloudAttrs = FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS | FILE_ATTRIBUTE_RECALL_ON_OPEN | FILE_ATTRIBUTE_OFFLINE;

struct WorkerStart
{
    std::shared_ptr<CThumbLoader::Shared> shared;
    HMODULE hSelf;
};

UINT ReadOrientation(IWICBitmapFrameDecode* pFrame)
{
    ComPtr<IWICMetadataQueryReader> q;
    if (FAILED(pFrame->GetMetadataQueryReader(&q)))
        return 1;
    // The photo-metadata policy name covers every container WIC knows; the
    // raw EXIF paths are fallbacks for JPEG and TIFF.
    static const wchar_t* const kNames[] = {
        L"System.Photo.Orientation", L"/app1/ifd/{ushort=274}", L"/ifd/{ushort=274}" };
    for (const wchar_t* name : kNames)
    {
        PROPVARIANT v;
        PropVariantInit(&v);
        UINT o = 0;
        if (SUCCEEDED(q->GetMetadataByName(name, &v)))
        {
            if (v.vt == VT_UI2)
                o = v.uiVal;
            else if (v.vt == VT_UI4)
                o = v.ulVal;
        }
        PropVariantClear(&v);
        if (o >= 1 && o <= 8)
            return o;
    }
    return 1;
}

// EXIF orientation -> the clockwise rotation (applied before the flip) that
// shows the image upright.
WICBitmapTransformOptions OrientationTransform(UINT o)
{
    switch (o)
    {
    case 2: return WICBitmapTransformFlipHorizontal;
    case 3: return WICBitmapTransformRotate180;
    case 4: return WICBitmapTransformFlipVertical;
    case 5: return (WICBitmapTransformOptions)(WICBitmapTransformRotate90 | WICBitmapTransformFlipHorizontal);
    case 6: return WICBitmapTransformRotate90;
    case 7: return (WICBitmapTransformOptions)(WICBitmapTransformRotate270 | WICBitmapTransformFlipHorizontal);
    case 8: return WICBitmapTransformRotate270;
    }
    return WICBitmapTransformRotate0;
}

HRESULT DecodeWithWIC(IWICImagingFactory* f, const PGLoadJob& job, UINT flags, PGLoadResult& r)
{
    if (!f)
        return E_NOINTERFACE;

    ComPtr<IWICBitmapDecoder> dec;
    HRESULT hr = f->CreateDecoderFromFilename(job.path.c_str(), nullptr, GENERIC_READ,
                                              WICDecodeMetadataCacheOnDemand, &dec);
    if (FAILED(hr))
        return hr;
    ComPtr<IWICBitmapFrameDecode> frame;
    if (FAILED(hr = dec->GetFrame(0, &frame)))
        return hr;
    UINT w = 0, h = 0;
    if (FAILED(hr = frame->GetSize(&w, &h)))
        return hr;
    if (w == 0 || h == 0)
        return WINCODEC_ERR_BADIMAGE;

    const UINT orient = (flags & PGGLF_ORIENTATION) ? ReadOrientation(frame.Get()) : 1;
    const bool bSwap = orient >= 5;     // 90/270-degree turns swap width and height
    const UINT dispW = bSwap ? h : w, dispH = bSwap ? w : h;
    double s = ScaleNeeded(dispW, dispH, job.cx, job.cy, job.bFit);
    if (s > 1.0)
        s = 1.0;                        // never enlarge
    const UINT tw = (std::max)(1u, (UINT)(dispW * s + 0.5)), th = (std::max)(1u, (UINT)(dispH * s + 0.5));
    const UINT sw = bSwap ? th : tw, sh = bSwap ? tw : th;     // size before rotation

    ComPtr<IWICBitmapSource> src = frame;

    // Let the codec decode at a reduced size when it can; for JPEG that
    // scales in the DCT domain and is several times faster than a full decode.
    ComPtr<IWICBitmapSourceTransform> xf;
    if (sw < w && SUCCEEDED(frame.As(&xf)))
    {
        UINT cw = sw, ch = sh;
        WICPixelFormatGUID pf = GUID_WICPixelFormat32bppPBGRA;
        if (SUCCEEDED(xf->GetClosestSize(&cw, &ch)) && cw >= sw && ch >= sh && cw < w &&
            SUCCEEDED(xf->GetClosestPixelFormat(&pf)))
        {
            ComPtr<IWICBitmap> tmp;
            if (SUCCEEDED(f->CreateBitmap(cw, ch, pf, WICBitmapCacheOnLoad, &tmp)))
            {
                WICRect rc{ 0, 0, (INT)cw, (INT)ch };
                ComPtr<IWICBitmapLock> lk;
                if (SUCCEEDED(tmp->Lock(&rc, WICBitmapLockWrite, &lk)))
                {
                    UINT cb = 0, stride = 0;
                    BYTE* p = nullptr;
                    HRESULT hrx = lk->GetStride(&stride);
                    if (SUCCEEDED(hrx))
                        hrx = lk->GetDataPointer(&cb, &p);
                    if (SUCCEEDED(hrx))
                        hrx = xf->CopyPixels(nullptr, cw, ch, &pf, WICBitmapTransformRotate0, stride, cb, p);
                    lk.Reset();
                    if (SUCCEEDED(hrx))
                        src = tmp;
                }
            }
        }
    }

    UINT cw = 0, ch = 0;
    src->GetSize(&cw, &ch);
    if (cw != sw || ch != sh)
    {
        ComPtr<IWICBitmapScaler> sc;
        if (FAILED(hr = f->CreateBitmapScaler(&sc)) ||
            FAILED(hr = sc->Initialize(src.Get(), sw, sh, WICBitmapInterpolationModeFant)))
            return hr;
        src = sc;
    }

    ComPtr<IWICFormatConverter> cv;
    if (FAILED(hr = f->CreateFormatConverter(&cv)) ||
        FAILED(hr = cv->Initialize(src.Get(), GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherTypeNone,
                                   nullptr, 0.0, WICBitmapPaletteTypeCustom)))
        return hr;
    src = cv;

    if (orient != 1)
    {
        ComPtr<IWICBitmapFlipRotator> fr;
        if (FAILED(hr = f->CreateBitmapFlipRotator(&fr)) ||
            FAILED(hr = fr->Initialize(src.Get(), OrientationTransform(orient))))
            return hr;
        src = fr;
    }

    UINT ow = 0, oh = 0;
    src->GetSize(&ow, &oh);
    PGPixels px;
    px.w = (int)ow;
    px.h = (int)oh;
    px.px.resize((size_t)ow * oh);
    if (FAILED(hr = src->CopyPixels(nullptr, ow * 4, ow * oh * 4, (BYTE*)px.px.data())))
        return hr;
    r.pixels = std::move(px);
    r.bNative = s >= 1.0;
    return S_OK;
}

HRESULT ShellThumb(const std::wstring& path, int size, UINT siigbf, PGPixels& out)
{
    ComPtr<IShellItemImageFactory> fac;
    HRESULT hr = SHCreateItemFromParsingName(path.c_str(), nullptr, IID_PPV_ARGS(&fac));
    if (FAILED(hr))
        return hr;
    HBITMAP hbm = nullptr;
    if (FAILED(hr = fac->GetImage(SIZE{ size, size }, (SIIGBF)siigbf, &hbm)))
        return hr;
    if (!hbm)
        return E_FAIL;
    bool bOk = PixelsFromHBITMAP(hbm, out);
    DeleteObject(hbm);
    return bOk ? S_OK : E_FAIL;
}

// The provider is still fetching the preview; a later request returns it.
const HRESULT kExtractionPending = (HRESULT)0x8004B205L;   // WTS_E_EXTRACTIONPENDING

void LoadOne(IWICImagingFactory* f, const PGLoadJob& job, UINT flags, PGLoadResult& r)
{
    r.key = job.key;
    r.epoch = job.epoch;
    r.path = job.path;
    r.cx = job.cx;
    r.cy = job.cy;

    DWORD attrs = GetFileAttributesW(job.path.c_str());
    if (attrs == INVALID_FILE_ATTRIBUTES)
    {
        DWORD e = GetLastError();
        r.hr = HRESULT_FROM_WIN32(e);
        r.uReason = e == ERROR_ACCESS_DENIED ? PGGTR_ACCESSDENIED : PGGTR_NOTFOUND;
        return;
    }
    if (attrs & FILE_ATTRIBUTE_DIRECTORY)
    {
        r.uReason = PGGTR_UNSUPPORTED;
        return;
    }
    const bool bCloud = (attrs & kCloudAttrs) != 0;
    const bool bMayOpen = !bCloud || (flags & PGGLF_CLOUDRECALL);

    auto useShell = [&](PGPixels& px) {
        r.pixels = std::move(px);
        r.bOk = true;
        r.uSource = PGGTSRC_SHELLCACHE;
    };

    // 1. The Windows thumbnail cache: instant for anything Explorer has shown,
    //    and already oriented. Cache-only unless we may open a cloud file, in
    //    which case the provider's own thumbnail handler may run.
    const int edge = (std::max)(job.cx, job.cy);
    const int size = (std::min)(1024, (std::max)(32, edge * 3 / 2));
    PGPixels shell;
    bool bShell = false;
    if (flags & PGGLF_SHELLCACHE)
    {
        UINT siigbf = SIIGBF_THUMBNAILONLY | SIIGBF_BIGGERSIZEOK;
        if (!(bCloud && bMayOpen))
            siigbf |= SIIGBF_INCACHEONLY;
        bShell = SUCCEEDED(ShellThumb(job.path, size, siigbf, shell));
        if (bShell && ScaleNeeded(shell.w, shell.h, job.cx, job.cy, job.bFit) <= 1.25)
        {
            useShell(shell);
            return;
        }
    }

    if (!bMayOpen)
    {
        // 2. Cloud-only and not allowed to download: the provider's own preview.
        //    OneDrive answers "pending" the first time and fetches it in the
        //    background; the file itself never comes down.
        if (flags & PGGLF_CLOUDTHUMBS)
        {
            PGPixels preview;
            HRESULT hr = ShellThumb(job.path, size, SIIGBF_THUMBNAILONLY | SIIGBF_BIGGERSIZEOK, preview);
            if (SUCCEEDED(hr) && (!bShell || preview.w * preview.h > shell.w * shell.h))
            {
                r.pixels = std::move(preview);
                r.bOk = true;
                r.uSource = PGGTSRC_PROVIDER;
                return;
            }
            if (hr == kExtractionPending && !bShell)
            {
                r.uReason = PGG_REASON_PENDING;
                r.hr = hr;
                return;
            }
        }
        if (bShell)
            useShell(shell);    // the best we can do without downloading the file
        else
            r.uReason = PGGTR_CLOUDONLY;
        return;
    }

    // 2. Decode the file.
    if (flags & PGGLF_DECODE)
    {
        HRESULT hr = DecodeWithWIC(f, job, flags, r);
        if (SUCCEEDED(hr))
        {
            r.bOk = true;
            r.uSource = PGGTSRC_DECODE;
            return;
        }
        r.hr = hr;
        r.pixels = PGPixels();
        if (bShell)
        {
            useShell(shell);
            return;
        }
        if (hr == WINCODEC_ERR_COMPONENTNOTFOUND || hr == WINCODEC_ERR_UNKNOWNIMAGEFORMAT)
            r.uReason = PGGTR_UNSUPPORTED;
        else if (hr == E_ACCESSDENIED || hr == HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION))
            r.uReason = PGGTR_ACCESSDENIED;
        else
            r.uReason = PGGTR_DECODEERROR;
        return;
    }

    if (bShell)
        useShell(shell);
    else
        r.uReason = PGGTR_UNSUPPORTED;
}

DWORD WINAPI LoaderWorker(LPVOID pv)
{
    HMODULE hSelf = nullptr;
    {
        std::unique_ptr<WorkerStart> start(static_cast<WorkerStart*>(pv));
        hSelf = start->hSelf;
        std::shared_ptr<CThumbLoader::Shared> shared = std::move(start->shared);
        start.reset();

        HRESULT hrCo = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
        {
            ComPtr<IWICImagingFactory> f;
            CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&f));
            for (;;)
            {
                PGLoadJob job;
                UINT flags = 0;
                {
                    std::unique_lock<std::mutex> lk(shared->mtx);
                    shared->cv.wait(lk, [&] { return shared->bStop || !shared->queue.empty(); });
                    if (shared->bStop)
                        break;
                    job = std::move(shared->queue.front());
                    shared->queue.pop_front();
                    shared->inFlight.insert(job.key);
                    flags = shared->flags;
                }

                PGLoadResult r;
                LoadOne(f.Get(), job, flags, r);

                std::lock_guard<std::mutex> lk(shared->mtx);
                if (!shared->bStop && shared->hwnd)
                {
                    // The key stays "in flight" until the window has taken the
                    // result; otherwise a paint in between would queue it again.
                    shared->done.push_back(std::move(r));
                    if (!shared->bPostPending)
                    {
                        // Posted under the lock: once Detach() has run, nothing is posted.
                        shared->bPostPending = true;
                        PostMessage(shared->hwnd, WM_PGG_LOADERDONE, 0, 0);
                    }
                }
                else
                {
                    shared->inFlight.erase(job.key);
                }
            }
        }
        if (SUCCEEDED(hrCo))
            CoUninitialize();
    }
    // Everything above has been destroyed; the DLL may go once this returns.
    FreeLibraryAndExitThread(hSelf, 0);
}

} // namespace

UINT CThumbLoader::AutoThreadCount()
{
    SYSTEM_INFO si{};
    GetSystemInfo(&si);
    UINT n = si.dwNumberOfProcessors > 1 ? si.dwNumberOfProcessors - 1 : 1;
    return (std::min)(n, 4u);
}

CThumbLoader::CThumbLoader(HWND hwndNotify, const PGGLOADERSETTINGS& settings)
    : m_shared(std::make_shared<Shared>())
{
    m_shared->hwnd = hwndNotify;
    m_shared->flags = settings.flags;

    UINT n = settings.nThreads ? (std::min)(settings.nThreads, 16u) : AutoThreadCount();
    for (UINT i = 0; i < n; ++i)
    {
        // Pin the DLL before the thread exists, so it cannot be unloaded
        // between CreateThread and the worker's first instruction.
        HMODULE hSelf = nullptr;
        if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,
                                reinterpret_cast<LPCWSTR>(&LoaderWorker), &hSelf))
            continue;
        WorkerStart* start = new (std::nothrow) WorkerStart{ m_shared, hSelf };
        if (!start)
        {
            FreeLibrary(hSelf);
            continue;
        }
        HANDLE hThread = CreateThread(nullptr, 0, LoaderWorker, start, 0, nullptr);
        if (!hThread)
        {
            delete start;
            FreeLibrary(hSelf);
            continue;
        }
        SetThreadPriority(hThread, THREAD_PRIORITY_BELOW_NORMAL);
        CloseHandle(hThread);
    }
}

CThumbLoader::~CThumbLoader()
{
    Detach();
}

void CThumbLoader::SetWanted(std::vector<PGLoadJob>&& jobs)
{
    {
        std::lock_guard<std::mutex> lk(m_shared->mtx);
        m_shared->queue.clear();
        std::unordered_set<ULONGLONG> seen;
        for (PGLoadJob& job : jobs)
        {
            if (m_shared->inFlight.count(job.key) || !seen.insert(job.key).second)
                continue;
            m_shared->queue.push_back(std::move(job));
        }
    }
    m_shared->cv.notify_all();
}

void CThumbLoader::TakeResults(std::vector<PGLoadResult>& out)
{
    std::lock_guard<std::mutex> lk(m_shared->mtx);
    out.swap(m_shared->done);
    m_shared->done.clear();
    for (const PGLoadResult& r : out)
        m_shared->inFlight.erase(r.key);
    m_shared->bPostPending = false;
}

void CThumbLoader::SetFlags(UINT flags)
{
    std::lock_guard<std::mutex> lk(m_shared->mtx);
    m_shared->flags = flags;
}

void CThumbLoader::Detach()
{
    {
        std::lock_guard<std::mutex> lk(m_shared->mtx);
        m_shared->bStop = true;
        m_shared->hwnd = NULL;
        m_shared->queue.clear();
        m_shared->done.clear();
    }
    m_shared->cv.notify_all();
}
