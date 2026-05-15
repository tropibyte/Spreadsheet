# Grid32 Bug Audit

Findings from the audit performed on 2026-05-15. Each item carries a checkbox; commits referencing the fix should tick the box.

## Critical

- [x] **C1** `IsCellVisible` always returns true (`Grid32Mgr.cpp:1696-1707`) — `cellPos` initialized to `{0,0}` and never mutated; programmatic scroll-to-cell silently broken.
- [x] **C2** Unsigned underflow in `IncrementSelectEdge` / `IncrementSelectedPage` (`Grid32Mgr.cpp:1507-1567, 1604-1625`) — `size_t idx = nRow - 1` wraps to `~0` at row 0; `pRowInfoArray[~0]` is a wild read.
- [x] **C3** `OnDestroy` never runs when `WM_CREATE` fails (`grid32_wndproc.cpp:40-66`, `Grid32Mgr.cpp:65-70`) — GDI+ token, default `HFONT`, edit subclass leak on failed grid creates.
- [x] **C4** `GetCellByPoint` loop reads past `pColInfoArray[gcs.nWidth]` (`Grid32Mgr.cpp:1834, 1849`, also `SetCellByPoint` 1866/1874) — comma-operator order bug + missing upper-bound check.
- [x] **C5** `OnPaint` exception path leaks memory DC + bitmap (`Grid32Mgr.cpp:185-231`) — cleanup inside `try`, jumped past on throw.
- [x] **C6** Mouse coords extracted as unsigned (`grid32_wndproc.cpp:108, 112, 115, 118, 121, 124, 138, 141, 145`) — drag-above-window yields ~65000 instead of negatives.

## High

- [x] **H1** `CopyGridCell` drops `m_bFormula` and `m_wsFormula` (`Grid32Mgr.cpp:3474-3495`) — `GM_ENUMCELLS` returns broken formula data.
- [x] **H2** `OnSetCurrentCell(UINT,UINT)` underflows `m_visibleTopLeft` (`Grid32Mgr.cpp:2745-2748`) — missing the `> halfRows` ternary guard that `SetCurrentCell` has.
- [x] **H3** `m_bSelecting` / `m_bSizing` never reset on capture loss — add `WM_CAPTURECHANGED` / `WM_CANCELMODE` handler.
- [x] **H4** `OnGetCellText` writes `wszBuff[0]` without checking `nLen` (`Grid32Mgr.cpp:2916-2918`).
- [x] **H5** `EM_SETSEL` caret position uses inflated `GetCellTextLen` value (`Grid32Mgr.cpp:1015-1035, 1798-1800`) — `+2` is buffer headroom, not a caret offset.
- [x] **H6** `SetClipboardData` failure leaks `HGLOBAL` (`Grid32Mgr.cpp:2389-2401`).
- [x] **H7** CSV/TSV stream-in has no quote handling (`Grid32Mgr.cpp:3673-3705`).
- [x] **H8** Stream-in trusts `m_cbBuffSize` as bytes-of-payload (`Grid32Mgr.cpp:3677, 3710, 3771`) — reads garbage past terminator.
- [x] **H9** Stream-out `wmemcpy` does not null-terminate; writes through `const_cast<LPWSTR>(LPCWSTR)` (`Grid32Mgr.cpp:3645-3652`).
- [x] **H10** SSF stream-in doesn't recognize `=` prefix as formulas (`Grid32Mgr.cpp:3796-3811`) — round-trip loses formulas.
- [x] **H11** `GetCellText` buffer size: `_tcsncpy_s(..., src.length())` invokes invalid-parameter handler on overflow (`Grid32Mgr.cpp:1126-1142`).
- [x] **H12** Formulator has no recursion-depth limit; `CountRange`/`SumRange` don't bail on visited cells (`Formulator.cpp:115-134`).
- [x] **H13** Sort recalculates all formulas O(N) times during a single sort (`Grid32Mgr.cpp:3263-3306`, root cause `SetCell` at `1010`).

## Medium

- [x] **M1** `OnPaste` calls `SetCellText` in a tight loop → recalc thrash (`Grid32Mgr.cpp:2478-2497`).
- [x] **M2** `OnSetSelection` signed/unsigned comparison is a no-op (`Grid32Mgr.cpp:2655`) — `pGridSel->start.nRow < 0` always false on UINT.
- [x] **M3** `CalculatePageStats`: `pageStat.end.nCol = gcs.nWidth` is one past end (`Grid32Mgr.cpp:1653-1654`).
- [x] **M4** `OnIncrementCell` can't express negative deltas (`Grid32Mgr.cpp:3166-3180`) — UINT struct passed where signed deltas needed.
- [x] **M5** Edit subclass is never restored on destroy (`Grid32Mgr.cpp:149-150`, `EditWndProc.cpp:8`) — late-queue messages dispatched against stale parent.
- [x] **M6** `DrawCornerCell` highlight RGB uses `GetGValue` where it should use `GetBValue` (`Grid32Mgr.cpp:365-367`).
- [x] **M7** `CalcSelectionCoordinatesWithMouse` decrements UINT past zero with `(UINT)-100` sentinel masking the bug (`Grid32Mgr.cpp:1949-1957, 1982-1989`).

## Low / Also noted

- [x] **L1** `Grid32Mgr.cpp:1411-1414` — `gcs.nHeight - 1` wraps when `gcs.nHeight == 0`.
- [x] **L2** `Grid32Mgr.cpp:1392` — `abs(LONG_MIN)` is UB.
- [x] **L3** `Grid32Mgr.cpp:3634` — XLSX writes column letter as `(wchar_t)(L'A' + c)`; columns ≥ 26 produce invalid letters.
- [x] **L4** `Grid32Mgr.cpp:2566` — Undo dispatch can throw mid-restore, pushing partial state to redo stack.
- [x] **L5** `Grid32Mgr.cpp:111` — `nColHeaderHeight` overwritten with caller's `nDefRowHeight = 0`, collapsing the header.
- [ ] **L6** `Grid32Mgr.cpp:2705-2717` — `OnSetCurrentCell` with non-numeric `SETBYCOORDINATE` returns `GRID_ERROR_NOT_IMPLEMENTED` instead of `INVALID_PARAMETER`.
- [ ] **L7** `Formulator.cpp:147` — `=1/0` silently returns the dividend instead of an error.
- [ ] **L8** `grid32_wndproc.cpp:32-38` — `new CGrid32Mgr()` in `WM_NCCREATE` without `nothrow`; bad_alloc through Win32 callback is UB on most ABIs.
- [ ] **L9** `Grid32Mgr.cpp:2147 / 2155` — local `RECT r` in `OnMouseMove` computed but never used.
