# Linux test build

Grid32 is Win32 code built with MSVC, but almost none of what is worth testing
in it is actually Windows-specific: the cell store, type inference, the formula
parser, recalculation, number formats and the stream-in parsers are all plain
C++. This directory lets those parts build and run on Linux, so CI can catch
regressions without a Windows runner.

```
make check     # build and run all five suites
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
| `GridMgrTests` | `GridMgrTests.cpp` + stubs + `Grid32Mgr.cpp`, `Formulator.cpp`, `EditWndProc.cpp` | the real manager: typing and the editor's state, selection, undo, clear, fill, sort, filter, merge, find/replace, clipboard payloads, formats |
| `FormulaRecalcTests` | `FormulaRecalcTests.cpp` + the same sources | the formula engine through the manager: recalculation, number formats on formula cells, grid-bounded references |
| `StreamTests` | `StreamTests.cpp` + the same sources | import and export in every format, and `GM_CLEARALL` |

The first two also build on Windows via `../run_tests.bat` and
`../run_formulator_tests.bat`. The other three are Linux-only: they define
Win32 symbols that would collide with the real ones.

## What this does and does not prove

It proves the logic is right and that the code compiles as C++17 under GCC's
warning set — which is stricter than the project's `/W3` in places, and has
already caught real bugs.

It does **not** prove anything about behaviour that lives in Win32 itself:
painting, hit-testing, scrolling, the clipboard, or the edit-control subclass.
`../Grid32LiveTests.cpp` (`../run_live_tests.bat`) covers part of that on
Windows — a real Grid32 window, its inline editor, its paint and print
rendering, driven through the message API — but mouse hit-testing and the
clipboard still need a human at the keyboard. Treat a green run here as "the
logic is sound", not "it ships".

The MSVC build is covered separately by the `windows-build` job in
`.github/workflows/ci.yml`, which compiles Grid32 for x64 and Win32, runs the
two portable suites under `cl.exe`, and runs the live-window tests against the
x64 DLL. That job deliberately does not build
the MFC host — see the comment at the end of the workflow for why, and what it
would cost to add.
