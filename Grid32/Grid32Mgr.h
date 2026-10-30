#pragma once

#include <string>
#include <vector>
#include <map>
#include <set>
#include <stack>

#include "grid32_cell.h"

#define ID_CELL_EDIT		8
#define HOVER_TIMER			1

enum class EditOperationType {
	SetText,
	SetFullCell,
	SetFormat,
	Delete,
	Batch       // any number of cells, each restored from its own snapshot
};

// One cell's before/after snapshot inside a Batch operation. A side with no
// cell (bHadOld / bHasNew false) is restored by deleting the cell, so undoing
// a command that created cells leaves the map as sparse as it found it.
struct CellChange {
	GRIDPOINT pt{};
	bool bHadOld = false;
	bool bHasNew = false;
	GRIDCELL oldCell;
	GRIDCELL newCell;
};

struct GridEditOperation {
        EditOperationType type = EditOperationType::SetFullCell;
        UINT row = 0, col = 0;
        GRIDCELL oldState;
        GRIDCELL newState;
        // Single-cell ops: the cell did not exist before the edit, so undo
        // deletes it instead of writing a default cell in its place.
        bool bOldMissing = false;
        // Batch ops.
        std::vector<CellChange> changes;
        // Whole-sheet formatting changes the default cell rather than
        // creating a cell for every position in the grid.
        bool bDefaultChanged = false;
        GRIDCELL oldDefault, newDefault;
};

class CGrid32Mgr
{
protected:
	// Records every cell a multi-cell command touches as one undo step.
	// While a scope is open, per-cell recording is off and the cell mutators
	// (SetCell, SetCellText, ClearCellText, DeleteCell, CreateCell,
	// GetCellOrCreate) snapshot each cell the first time they touch it, so a
	// command only has to open the scope to become undoable. Nested scopes
	// fold into the outermost one, which records when it closes.
	class BatchUndo
	{
	public:
		explicit BatchUndo(CGrid32Mgr& mgr);
		~BatchUndo();
		BatchUndo(const BatchUndo&) = delete;
		BatchUndo& operator=(const BatchUndo&) = delete;
		void Touch(UINT nRow, UINT nCol);
		void TouchDefault();
	private:
		CGrid32Mgr& m_mgr;
		BatchUndo* m_pOuter;
		BOOL m_prevRecord;
		bool m_bCollect;
		std::map<std::pair<UINT, UINT>, CellChange> m_cells;
		bool m_bDefault = false;
		GRIDCELL m_oldDefault;
	};
	friend class BatchUndo;

	// HFONTs are cached by everything that shapes the glyphs, so painting
	// creates each distinct font once instead of once per cell per paint.
	struct FontKey {
		std::wstring face;
		int height;
		UINT weight;
		BOOL italic, underline, strike;
		bool operator<(const FontKey& o) const;
	};

	GRIDCREATESTRUCT gcs;
	std::map<std::pair<UINT, UINT>, PGRIDCELL> mapCells;
	GRIDCELL m_defaultGridCell, m_cornerCell;
	// Row and column metadata is sparse: only sizes that differ from the
	// default are stored, so a grid of any size costs nothing up front.
	std::map<UINT, size_t> m_colWidths, m_rowHeights;
	std::set<UINT> m_hiddenRows;        // rows hidden by GM_FILTERCELLS

	HWND m_hWndGrid, m_hWndEdit;
	long nColHeaderHeight, nRowHeaderWidth;
	RECT m_clientRect;
	RECT m_paintRect;                   // the update region of the paint in progress
	GRIDPOINT m_currentCell, m_visibleTopLeft, m_HoverCell;
	BOOL m_bRedraw, m_bResizable, m_bUndoRecordEnabled;
	BOOL m_bDeferRecalc;
	BOOL m_bSuppressChangeNotify;
	DWORD dwError;
	GRIDSELECTION m_selectionRect;      // start = anchor, end = the moving corner
	BOOL m_bSelecting, m_bSizing;
	WNDPROC m_editWndProc;
	UINT m_nMouseHoverDelay;
	UINT_PTR m_npHoverDelaySet;
	LONG_PTR m_gridHitTest;
	size_t m_nToBeSized;
	size_t m_resizeStartSize;   // pre-resize column width / row height; captured on divider down
	// Inline editor state. The edit always commits into m_editCell, the cell
	// it was opened on, whatever the current cell has become since.
	bool m_bEditing;
	GRIDPOINT m_editCell;
	std::wstring m_editOriginal;        // text the editor opened with
	HFONT m_hEditFont;                  // cached font the editor is using, not owned
	long m_nSizingLine;
	COLORREF m_rgbSizingLine;
	HFONT m_hDefaultFont;
	POINT m_mouseDraggingStartPoint;
	std::stack<GridEditOperation> m_undoStack;
	std::stack<GridEditOperation> m_redoStack;
	BatchUndo* m_pBatch;
	std::map<FontKey, HFONT> m_fontCache;
	// Derived from the anchor cells' mergeRange (the SSF-persisted form) and
	// rebuilt lazily whenever a mutator that could have moved one runs.
	std::vector<MERGERANGE> m_merges;
	bool m_bMergesDirty;
	UINT m_cfGridCells;                 // private clipboard format (formulas + formats)

	ULONG_PTR m_gdiplusToken;

public:
	CGrid32Mgr();
	virtual ~CGrid32Mgr();
	static std::wstring GetColumnLabel(size_t index);
	bool Create(PGRIDCREATESTRUCT pGCS);
	void OnDestroy();
	static LRESULT CALLBACK Grid32_WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	static LRESULT CALLBACK EditCtrl_WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam);

	// ---- geometry ----------------------------------------------------------
	size_t ColW(size_t nCol) const;     // 0 for out-of-range columns
	size_t RowH(size_t nRow) const;     // 0 for hidden or out-of-range rows
	bool IsRowHidden(UINT nRow) const { return m_hiddenRows.count(nRow) != 0; }
	size_t CalculatedColumnDistance(size_t start, size_t end) const;
	size_t CalculatedRowDistance(size_t start, size_t end) const;
	long long GetColumnStartPos(UINT nCol);
	long long GetRowStartPos(UINT nRow);
	bool GetCellRect(UINT nRow, UINT nCol, RECT& rect);   // merge-aware, client coords
	void GetCurrentCellRect(RECT& rect);
	bool IsCellVisible(UINT nRow, UINT nCol);
	void CalculatePageStats(PAGESTAT& pageStat);
	long GetActualColHeaderHeight() { return gcs.style & GS_COLHEADER ? nColHeaderHeight : 0; }
	long GetActualRowHeaderWidth() { return gcs.style & GS_ROWHEADER ? nRowHeaderWidth : 0; }

	// ---- painting ----------------------------------------------------------
	void OnPaint(PAINTSTRUCT& ps);
	bool CanDrawRect(const RECT& rect);
	void DrawCurrentCellBox(HDC hDC);
	void DrawHeader(HDC hDC);
	void DrawHeaderButton(HDC hDC, COLORREF background, COLORREF border, RECT coordinates, bool bRaised);
	void DrawRowHeaders(HDC hDC);
	void DrawColHeaders(HDC hDC);
	void DrawGrid(HDC hDC);
	void DrawSelectionOverlay(HDC hDC, RECT& gridCellRect);
	void DrawCells(HDC hDC);
	void DrawCellContent(HDC hDC, const GRIDCELL& cell, const RECT& rect, int logPixelsY, bool bFillBackground);
	void DrawVoidSpace(HDC hDC);
	void DrawSizingLine(HDC hDC);
	void DrawSemiTransparentRectangle(HDC hdc, RECT rect, COLORREF color, BYTE alpha);
	void DrawSelectionSurround(HDC hDC);
	HFONT GetCachedFont(const CELLFONT& fi, int lfHeight);
	void ClearFontCache();
	// Render a block of cells to any DC (printing, print preview, export).
	BOOL OnRenderRange(const GCRENDERSTRUCT& rs);
	BOOL OnGetUsedRange(GRIDSELECTION* pRange);

	// ---- merged cells ------------------------------------------------------
	void RefreshMergeIndex();
	const MERGERANGE* FindMerge(UINT nRow, UINT nCol);
	bool IsCellMerged(UINT row, UINT col);
	void ExpandForMerges(GRIDSELECTION& sel);
	BOOL MergeCells(const GRIDSELECTION& sel);
	BOOL UnmergeCells(const GRIDSELECTION& sel);

	// ---- cell store --------------------------------------------------------
	PGRIDCELL GetCell(UINT nRow, UINT nCol);
	PGRIDCELL GetCellOrDefault(UINT nRow, UINT nCol);
	// Returns nullptr (and sets GRID_ERROR_OUT_OF_MEMORY) if the cell could
	// not be allocated; never a shared placeholder a caller could write to.
	PGRIDCELL GetCellOrCreate(UINT nRow, UINT nCol, bool bRecord = false);
	PGRIDCELL CreateCell(UINT nRow, UINT nCol);
	auto GetCurrentCell() { return GetCellOrDefault(m_currentCell.nRow, m_currentCell.nCol); }
	static bool CellsEqual(const GRIDCELL& a, const GRIDCELL& b);
	// Positions of the cells that exist inside sel, in row-major order.
	std::vector<GRIDPOINT> CellsInRange(const GRIDSELECTION& sel);
	bool IsWholeGrid(const GRIDSELECTION& sel) const;
	void TouchCell(UINT nRow, UINT nCol) { if (m_pBatch) m_pBatch->Touch(nRow, nCol); }
	UINT GetCurrentCellTextLen();
	UINT GetCellTextLen(const GRIDPOINT& gridPt);
	UINT GetCellTextLen(UINT nRow, UINT nCol);
	void GetCurrentCellText(LPWSTR pText, UINT nLen);
	void GetCellText(const GRIDPOINT& gridPt, LPWSTR pText, UINT nLen);
	void GetCellText(UINT nRow, UINT nCol, LPWSTR pText, UINT nLen);
	void SetCell(UINT nRow, UINT nCol, const GRIDCELL& gc);
	void SetCurrentCellText(LPCWSTR newText);
	void SetCellText(UINT nRow, UINT nCol, LPCWSTR newText);
	void SetCellText(const GRIDPOINT& gridPt, LPCWSTR newText);
	void ClearCellText(UINT nRow, UINT nCol);
	void DeleteCell(UINT nRow, UINT nCol);
	void DeleteAllCells();

	// ---- formatting --------------------------------------------------------
	// mask is a GCFM_* set: only those attributes are copied from fontInfo,
	// so toggling Bold on a mixed selection leaves each cell's face alone.
	void SetCellFormat(UINT nRow, UINT nCol, const CELLFONT& fontInfo, DWORD mask = GCFM_ALL);
	void SetSelectionFormat(const CELLFONT& fontInfo, DWORD mask = GCFM_ALL);
	void SetRangeFormat(const GRIDSELECTION& sel, const CELLFONT& fontInfo, DWORD mask = GCFM_ALL);
	static void ApplyFontMask(CELLFONT& dst, const CELLFONT& src, DWORD mask);
	// Set the per-cell number format (FMT_*). Re-renders display text.
	void SetCellNumberFormat(UINT nRow, UINT nCol, UINT format);
	void SetSelectionNumberFormat(UINT format);
	void SetSelectionBackground(COLORREF clr);
	void SetRangeBackground(const GRIDSELECTION& sel, COLORREF clr);
	// Read the current cell's font info (the default font for an empty cell).
	bool GetCurrentCellFontInfo(CELLFONT& out);
	// Alignment / wrap. Each setter updates only the bits it owns
	// (horiz: DT_LEFT/CENTER/RIGHT; vert: DT_TOP/VCENTER/BOTTOM; wrap: DT_WORDBREAK).
	void SetSelectionHAlign(UINT halign);
	void SetSelectionVAlign(UINT valign);
	void SetSelectionWrap(BOOL wrap);
	void SetSelectionAlignBits(UINT axis, UINT value);
	// Apply fn(cell) to every cell of the selection as one undoable step.
	// A whole-grid selection restyles the existing cells and the default
	// cell instead of allocating every position (AUDIT_2026-10 S10).
	template <typename Fn> void ForEachFormattedCell(const GRIDSELECTION& sel, Fn fn);

	// ---- navigation and selection -----------------------------------------
	void OnHScroll(UINT nSBCode, UINT nPos, HWND hScrollBar);
	void OnVScroll(UINT nSBCode, UINT nPos, HWND hScrollBar);
	void SetScrollRanges();
	// Move the active cell (bExtend = false: the selection collapses onto it)
	// or the selection's moving corner (bExtend = true: the anchor and active
	// cell stay). Scrolls the target into view and sends one GN_SELCHANGE.
	void MoveCurrentCell(UINT nRow, UINT nCol, bool bExtend);
	void IncrementSelectedCell(long nRow, long nCol, bool bExtend = false);
	void IncrementSelectedPage(long nRow, long nCol, bool bExtend = false);
	void IncrementSelectEdge(long nRow, long nCol, bool bExtend = false);
	void EnsureCellVisible(UINT nRow, UINT nCol);
	// Step delta rows from row, skipping rows a filter hid.
	UINT StepVisibleRow(UINT row, long delta);
	void ClampTopLeft();
	void NotifySelChange();
	void OnMove(int x, int y);
	void OnSize(UINT nType, int cx, int cy);
	bool OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	bool OnKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags);
	bool OnChar(WCHAR ch);
	void GetCellByPoint(int x, int y, GRIDPOINT& gridPt);
	void StartMouseTracking();
	void CancelMouseTracking();
	void CancelDragState();
	void CalcSelectionCoordinatesWithMouse(int x, int y, bool& bChangeX, bool& bChangeY);
	void GetCurrentMouseCoordinates(int& x, int& y);
	void OnLButtonDown(UINT nFlags, int x, int y);
	void OnLButtonDblClk(UINT nFlags, int x, int y);
	void OnLButtonUp(UINT nFlags, int x, int y);
	void OnMouseMove(UINT nFlags, int x, int y);
	void OnMouseHover(UINT nFlags, int x, int y);
	void OnMouseLeave(UINT nFlags, int x, int y);
	void OnNcMouseMove(UINT nFlags, int x, int y);
	LRESULT OnNcHitTest(UINT nFlags, int x, int y, bool& bHitTested);
	void OnMouseWheel(int zDelta, UINT nFlags, int x, int y);
	void OnRButtonDown(UINT nFlags, int x, int y);
	void OnRButtonUp(UINT nFlags, int x, int y);
	bool IsCoordinateInClientArea(int x, int y);
	void SetCursorBasedOnNcHitTest();
	bool OnSetSelection(const GRIDSELECTION* pGridSel);
	void OnGetSelection(GRIDSELECTION* pGridSel);
	BOOL OnSetCurrentCell(LPCWSTR pwszRef, short nWhich);
	BOOL OnSetCurrentCell(UINT nRow, UINT nCol);
	// Parse "B7", "$B$7" or "A1:C3" (case-insensitive) into 0-based
	// coordinates. False for malformed or out-of-range references.
	bool ParseCellReference(LPCWSTR pwszRef, GRIDSELECTION& out);
	bool SetCurrentCell(UINT nRow, UINT nCol);
	GRIDSELECTION GetNormalizedSelection();
	bool AreGridPointsEqual(const GRIDPOINT& point1, const GRIDPOINT& point2);
	bool IsCellPointInsideSelection(const GRIDPOINT& gridPt);
	bool IsSelectionUnicellular();
	void ConstrainMaxPosition();
	void NormalizeSelectionRect(GRIDSELECTION& selection);
	void OnIncrementCell(UINT incUnit, GRIDPOINT* pGridPoint);

	// ---- inline editor -----------------------------------------------------
	// bKeepText: F2 / double-click (edit the existing content, caret at the
	// end). Otherwise the editor opens empty, as typing over a cell does.
	void BeginEdit(bool bKeepText);
	void ShowEditControl();
	bool IsEditing() const { return m_bEditing; }
	// Write the editor's text into m_editCell and close it. False if no edit
	// was open.
	bool CommitEdit();
	// Close the editor without writing. GN_EDITCANCEL goes out after the
	// edit is already closed, so a handler that moves focus can't commit it.
	void CancelEdit();
	void CommitEditText(const std::wstring& text);

	// ---- clipboard, undo ---------------------------------------------------
	void OnClear();
	void OnCopy();
	void OnCut();
	void OnPaste();
	// The clipboard payloads, separated from the clipboard so they can be
	// tested: tab-separated display text, and the private format that keeps
	// formulas and formatting.
	std::wstring BuildCopyText(const GRIDSELECTION& sel);
	std::wstring BuildCopyRecords(const GRIDSELECTION& sel);
	void PasteText(const std::wstring& data);
	bool PasteRecords(const std::wstring& data);
	bool OnUndo(GRIDSELECTION* pAffected = nullptr);
	bool OnRedo(GRIDSELECTION* pAffected = nullptr);
	size_t OnCanUndo();
	size_t OnCanRedo();
	void RecordUndoOperation(const GridEditOperation& op);
	// Re-apply one edit operation's recorded cell snapshot. bUndo selects
	// the old state, otherwise the new one. pAffected receives the bounding
	// box of the cells it touched.
	void RestoreCells(const GridEditOperation& op, bool bUndo, GRIDSELECTION* pAffected = nullptr);

	// ---- commands ----------------------------------------------------------
	DWORD GetLastError() { return dwError; }
	void SetLastError(DWORD dwNewError) { dwError = dwNewError; }
	BOOL Invalidate(bool bErasebkgnd = false);
	void InvalidateCell(UINT nRow, UINT nCol);
	bool SetColumnWidth(size_t nCol, size_t newWidth);
	size_t GetColumnWidth(size_t nCol);
	bool SetRowHeight(size_t nRow, size_t newHeight);
	size_t GetRowHeight(size_t nRow);
	BOOL OnGetCellText(const GRIDPOINT& point, GRID_GETTEXT* text);
	LRESULT SendGridNotification(INT code, GRIDNMHDR* pNMHDRInfo = nullptr);
	// Tell the host that persisted content changed, so it can mark its
	// document dirty. Every mutator that alters what StreamOut would write
	// must call this; bulk operations hold a BulkEditScope and send one
	// notification when the scope closes instead of one per cell.
	void NotifyContentChanged();
	void GetMousePosition(POINT& pt);
	void OnTimer(UINT_PTR idEvent);
	bool IsOverColumnDivider(int x, int y);
	bool IsOverRowDivider(int x, int y);
	bool IsCellFontDefault(GRIDCELL& gc);
	size_t OnEnumCells(GRIDPOINT* pts, UINT numElements);
	// GM_SETCELL: the fields of info that its mask names, as one undo step.
	BOOL OnSetCellInfo(const GRIDPOINT& pt, const GCCELLINFO& info);
	BOOL OnFillCells(WPARAM wParam, const GCFILLSTRUCT& fillStruct);
	BOOL OnSortCells(WPARAM wParam, const GCSORTSTRUCT& sortStruct);
        BOOL OnFilterCells(WPARAM wParam, const GCFILTERSTRUCT& filterStruct);
        bool OnFindText(const GCFINDSTRUCT& findStruct);
        // Locate the next cell matching findStruct without moving the
        // selection. Searches populated cells only, and a formula cell's
        // source as well as its displayed result.
        bool FindNextMatch(const GCFINDSTRUCT& findStruct, GRIDPOINT& found);
        bool OnReplaceText(const GCREPLACESTRUCT& replaceStruct);
        // Replace every match in the grid as one undoable step; returns the
        // number of occurrences replaced.
        UINT OnReplaceAll(const GCREPLACESTRUCT& replaceStruct);
        void OnStreamOut(LPGCSTREAM pStream);
        // Replaces the content: clears the grid (as ClearAll does) and loads
        // the stream, recording no undo. On a parse error the grid is untouched.
        void OnStreamIn(LPGCSTREAM pStream);
        // GM_CLEARALL: remove every cell, empty the undo and redo stacks,
        // cancel any in-progress cell edit and put the cursor, selection and
        // scroll position back at A1. Sends one GN_CONTENTCHANGED if any cell
        // was removed, none otherwise.
        void ClearAll();
        void RecalculateFormulas();
        // Read-only views for the formula engine: the grid's extent, so a
        // reference outside it is rejected instead of walked, and the
        // populated cells, so a range visits what exists rather than its area.
        size_t GetGridRowCount() const { return gcs.nHeight; }
        size_t GetGridColCount() const { return gcs.nWidth; }
        const std::map<std::pair<UINT, UINT>, PGRIDCELL>& GetCellMap() const { return mapCells; }
        // Re-evaluate one formula cell, refreshing both its display text and
        // the cached numeric value that sort/compare read. No-op for cells
        // that aren't formulas.
        void RefreshFormulaCell(PGRIDCELL pCell);
        std::wstring EvaluateFormula(const std::wstring& expr);
};

