# Linux test build

Grid32 is Win32 code built with MSVC, but almost none of what is worth testing
in it is actually Windows-specific: the cell store, type inference, the formula
parser, recalculation, number formats and the stream-in parsers are all plain
C++. This directory lets those parts build and run on Linux, so CI can catch
regressions without a Windows runner.

```
make check     # build and run all three suites
make clean
```

## How it works

`win32shim/` supplies just enough of `<windows.h>`, `<tchar.h>`, `<CommCtrl.h>`,
`<windowsx.h>`, `<objidl.h>` and `<gdiplus.h>` for the real Grid32 headers and
sources to compile: the basic typedefs (`HWND`, `RECT`, `COLORREF`, …), the
constants Grid32 references (`WM_*`, `DT_*`, `PS_*`, `VK_*`, …), the structs it
passes around (`SCROLLINFO`, `TRACKMOUSEEVENT`, `PAINTSTRUCT`, …), and
declarations for the ~90 API functions it calls.

`win32shim/win32stubs.cpp` defines those functions as no-ops returning zero or
null. Nothing draws, no window exists, every GDI handle is null. `Create()`
still succeeds, which is all `GridMgrTests` needs.

`win32shim/grid32mgr.h` exists because `Formulator.cpp` includes
`"grid32mgr.h"` while the file on disk is `Grid32Mgr.h` — fine on Windows,
not on a case-sensitive filesystem.

## The suites

| Suite | Sources | Covers |
|---|---|---|
| `Grid32Tests` | `../Grid32Tests.cpp` | `Grid32Detail::*` header helpers (dates, number parsing, formatting) |
| `FormulatorTests` | `../FormulatorTests.cpp` + `Formulator.cpp` | the formula parser, against stub `CGrid32Mgr` members |
| `GridMgrTests` | `GridMgrTests.cpp` + stubs + `Grid32Mgr.cpp`, `Formulator.cpp`, `EditWndProc.cpp` | the real manager: typing, recalculation, formats, CSV import |

The first two also build on Windows via `../run_tests.bat` and
`../run_formulator_tests.bat`. `GridMgrTests` is Linux-only: it defines Win32
symbols that would collide with the real ones.

## What this does and does not prove

It proves the logic is right and that the code compiles as C++17 under GCC's
warning set — which is stricter than the project's `/W3` in places, and has
already caught real bugs.

It does **not** prove anything about behaviour that lives in Win32 itself:
painting, hit-testing, scrolling, the clipboard, or the edit-control subclass.
Those still need a Windows build and a human at the keyboard. Treat a green run
here as "the logic is sound", not "it ships".

The MSVC build is covered separately by the `windows-build` job in
`.github/workflows/ci.yml`, which compiles Grid32 for x64 and Win32 and runs
the two portable suites under `cl.exe`. That job deliberately does not build
the MFC host — see the comment at the end of the workflow for why, and what it
would cost to add.

## Note on `-O0`

The Makefile builds at `-O0` on purpose. `Grid32FormulaText::ResetText` is
defined `inline` in `Formulator.cpp` but declared `extern` in `Grid32Mgr.cpp`
and called from there; at `-O2` GCC inlines every use, drops the symbol, and
the link fails with an undefined reference. That is `AUDIT_2026-09.md` **M23**.
Raise the optimisation level once it is fixed.
