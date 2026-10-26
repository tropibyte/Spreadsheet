#pragma once

// Background thumbnail loader for PGGS_AUTOLOAD.
//
// Workers never touch the control or the host's data: each job carries a
// copy of the path. Results go into a locked list and the window gets one
// private message to drain it, so nothing leaks if the window dies first.
// Workers share state with the control through a shared_ptr and pin
// Grid32.dll while they run, so destroying the control (or FreeLibrary)
// never waits on a slow network read and never unloads running code.

#include <string>
#include <vector>
#include <memory>
#include "photogallery.h"
#include "PhotoGalleryThumbs.h"

// Posted to the control when results are ready. Private to the control.
#define WM_PGG_LOADERDONE    (WM_USER + 0x3F0)

// PGLoadResult.uReason when a cloud provider is still fetching its preview
// (WTS_E_EXTRACTIONPENDING): ask again shortly. Never reported to the host.
#define PGG_REASON_PENDING   0x100

struct PGLoadJob
{
    ULONGLONG    key = 0;
    UINT         epoch = 0;
    std::wstring path;
    int          cx = 0, cy = 0;    // tile size the thumbnail is wanted for
    bool         bFit = false;
};

struct PGLoadResult
{
    ULONGLONG    key = 0;
    UINT         epoch = 0;
    std::wstring path;
    int          cx = 0, cy = 0;
    bool         bOk = false;
    bool         bNative = false;   // decoded at the image's full size
    UINT         uSource = 0;       // PGGTSRC_*
    UINT         uReason = 0;       // PGGTR_* when !bOk
    HRESULT      hr = S_OK;
    PGPixels     pixels;
};

class CThumbLoader
{
public:
    struct Shared;

    CThumbLoader(HWND hwndNotify, const PGGLOADERSETTINGS& settings);
    ~CThumbLoader();

    // Replace the queue. Front = most urgent. Jobs already running are kept.
    void SetWanted(std::vector<PGLoadJob>&& jobs);
    void TakeResults(std::vector<PGLoadResult>& out);
    void SetFlags(UINT flags);
    // Stop posting to the window and let the workers wind down on their own.
    void Detach();

    static UINT AutoThreadCount();

private:
    std::shared_ptr<Shared> m_shared;
};
