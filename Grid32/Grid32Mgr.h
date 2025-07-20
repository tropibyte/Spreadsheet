#pragma once

#include <string>
#include <vector>
#include <map>
#include <stack>

#include "grid32.h"

#define ID_CELL_EDIT		8
#define HOVER_TIMER			1

enum class EditOperationType {
	SetText,
	SetFullCell,
	SetFormat,
	Delete,
	Clipboard
};

struct GridEditOperation {
        EditOperationType type;
        UINT row, col;
        GRIDCELL oldState;
        GRIDCELL newState;
        std::vector<std::pair<GRIDPOINT, GRIDCELL>> oldCells;
        std::vector<std::pair<GRIDPOINT, GRIDCELL>> newCells;
};

class CGrid32Mgr
{
protected:
	GRIDCREATESTRUCT gcs;
	std::map<std::pair<UINT, UINT>, PGRIDCELL> mapCells;
	GRIDCELL m_defaultGridCell, m_cornerCell;
	ROWINFO* pRowInfoArray;
	COLINFO* pColInfoArray;

	HWND m_hWndGrid, m_hWndEdit;
	long nColHeaderHeight, nRowHeaderWidth;
	RECT totalGridCellRect, m_clientRect;
	struct {
		long width, height;
	}visibleGrid;
	GRIDPOINT m_currentCell, m_visibleTopLeft, m_HoverCell;
	POINT m_scrollDifference;
	BOOL m_bRedraw, m_bResizable, m_bUndoRecordEnabled;
	DWORD dwError;
	GRIDSELECTION m_selectionRect;
	BOOL m_bSelecting, m_bSizing;
	WNDPROC m_editWndProc;
	UINT m_nMouseHoverDelay;
	UINT_PTR m_npHoverDelaySet;
	unsigned long long m_lastClickTime;
	LONG_PTR m_gridHitTest;
	size_t m_nToBeSized;
	long m_nSizingLine;
	COLORREF m_rgbSizingLine;
	HFONT m_hDefaultFont;
	POINT m_mouseDraggingStartPoint;
	std::stack<GridEditOperation> m_undoStack;
	std::stack<GridEditOperation> m_redoStack;

	ULONG_PTR m_gdiplusToken;

public:
	CGrid32Mgr();
	virtual ~CGrid32Mgr();
	std::wstring GetColumnLabel(size_t index, size_t maxWidth);
	bool Create(PGRIDCREATESTRUCT pGCS);
	void OnDestroy();
	static LRESULT CALLBACK Grid32_WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	static LRESULT CALLBACK EditCtrl_WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam);
	void OnPaint(PAINTSTRUCT& ps);
	void CalculateTotalGridRect();
	void OffsetRectByScroll(RECT&/*rect*/);
	bool CanDrawRect(RECT rect);
	void DrawCurrentCellBox(HDC hDC);
	void DrawHeader(HDC hDC);
	void DrawCornerCell(HDC hDC);
	void DrawHeaderButton(HDC hDC, COLORREF background, COLORREF border, RECT coordinates, bool bRaised);
	void DrawRowHeaders(HDC hDC);
	void DrawColHeaders(HDC hDC);
	void DrawGrid(HDC hDC);
	void DrawSelectionOverlay(HDC hDC, RECT& gridCellRect);
	void DrawCells(HDC hDC);
	void DrawVoidSpace(HDC hDC);
	void DrawSizingLine(HDC hDC);
	void DrawSemiTransparentRectangle(HDC hdc, RECT rect, COLORREF color, BYTE alpha);
	void DrawSelectionSurround(HDC hDC);
	bool IsCellMerged(UINT row, UINT col);
	void DrawMergedCell(HDC hDC, const MERGERANGE& mergeRange);
	void SelectMergedRange(UINT row, UINT col);
	size_t CalculatedColumnDistance(size_t start, size_t end);
	size_t CalculatedRowDistance(size_t start, size_t end);
	void GetCurrentCellRect(RECT& rect);
	PGRIDCELL GetCell(UINT nRow, UINT nCol);
	PGRIDCELL GetCellOrDefault(UINT nRow, UINT nCol);
	PGRIDCELL GetCellOrCreate(UINT nRow, UINT nCol, bool bRecord = false);
	PGRIDCELL CreateCell(UINT nRow, UINT nCol);
	auto GetCurrentCell() { return GetCellOrDefault(m_currentCell.nRow, m_currentCell.nCol); }
	UINT GetCurrentCellTextLen();
	UINT GetCellTextLen(const GRIDPOINT& gridPt);
	UINT GetCellTextLen(UINT nRow, UINT nCol);
	void GetCurrentCellText(LPWSTR pText, UINT nLen);
	void GetCellText(const GRIDPOINT& gridPt, LPWSTR pText, UINT nLen);
	void SetCellFormat(UINT nRow, UINT nCol, const FONTINFO& fontInfo);
	void GetCellText(UINT nRow, UINT nCol, LPWSTR pText, UINT nLen);
	void SetCell(UINT nRow, UINT nCol, const GRIDCELL& gc);
	void SetCurrentCellText(LPCWSTR newText);
	void SetCellText(UINT nRow, UINT nCol, LPCWSTR newText);
	void SetCellText(const GRIDPOINT& gridPt, LPCWSTR newText);
	void ClearCellText(UINT nRow, UINT nCol);
	void DeleteCell(UINT nRow, UINT nCol);
	void DeleteAllCells();
	void OnHScroll(UINT nSBCode, UINT nPos, HWND hScrollBar);
	void OnVScroll(UINT nSBCode, UINT nPos, HWND hScrollBar);
	void SetScrollRanges();
	bool IsCellIncrementWithinCurrentPage(long nRow, long nCol, PAGESTAT& pageStat);
	void IncrementSelectedCell(long nRow, long nCol);
	void IncrementSelectedPage(long nRow, long nCol);
	void IncrementSelectEdge(long nRow, long nCol);
	void ScrollToCell(size_t row, size_t col, int scrollFlags);
	void OnMove(int x, int y);
	void OnSize(UINT nType, int cx, int cy);
	bool OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	bool OnKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags);
	void GetCellByPoint(int x, int y, GRIDPOINT& gridPt);
	void SetCellByPoint(int x, int y, GRIDPOINT& gridPt);
	void StartMouseTracking();
	void CancelMouseTracking();
	void ConstrainMouseToClientArea();
	void CalcSelectionCoordinatesWithMouse(int x, int y, bool& bChangeX, bool bChangeY);
	void GetCurrentMouseCoordinates(int& x, int& y);
	void OnLButtonDown(UINT nFlags, int x, int y);
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
	void OnClear();
	void OnCopy();
	void OnCut();
	void OnPaste();
	void OnUndo();
	void OnRedo();
	size_t OnCanUndo();
	size_t OnCanRedo();
	void OnSetSelection(GRIDSELECTION* pGridSel);
	void OnGetSelection(GRIDSELECTION* pGridSel);
	bool IsCellVisible(UINT nRow, UINT nCol);
	size_t MaxBeginRow();
	size_t MaxBeginColumn();
	void CalculatePageStats(PAGESTAT& pageStat);
	long GetActualColHeaderHeight() { return gcs.style & GS_COLHEADER ? nColHeaderHeight : 0; }
	long GetActualRowHeaderWidth() { return gcs.style & GS_ROWHEADER ? nRowHeaderWidth : 0; }
	BOOL OnSetCurrentCell(LPCWSTR pwszRef, short nWhich);
	BOOL OnSetCurrentCell(UINT nRow, UINT nCol);
	DWORD GetLastError() { return dwError; }
	void SetLastError(DWORD dwNewError) { dwError = dwNewError; }
	std::wstring FormatCellReference(LPCWSTR pwszRef);
	bool SetCurrentCell(UINT nRow, UINT nCol);
	BOOL Invalidate(bool bErasebkgnd = false);
	bool SetColumnWidth(size_t nCol, size_t newWidth);
	size_t GetColumnWidth(size_t nCol);
	bool SetRowHeight(size_t nRow, size_t newHeight);
	size_t GetRowHeight(size_t nRow);
	BOOL OnGetCellText(const GRIDPOINT& point, GRID_GETTEXT* text);
	void ShowEditControl();
	void SendGridNotification(INT code, GRIDNMHDR* pNMHDRInfo = nullptr);
	void GetMousePosition(POINT& pt);
	void OnTimer(UINT_PTR idEvent);
	bool IsOverColumnDivider(int x, int y);
	bool IsOverRowDivider(int x, int y);
	long long GetColumnStartPos(UINT nCol);
	long long GetRowStartPos(UINT nRow);
	bool IsCellFontDefault(GRIDCELL& gc);
	bool AreGridPointsEqual(const GRIDPOINT& point1, const GRIDPOINT& point2);
	bool IsCellPointInsideSelection(const GRIDPOINT& gridPt);
	bool IsSelectionUnicellular();
	void ConstrainMaxPosition();
	void NormalizeSelectionRect(GRIDSELECTION& selection);
	void OnIncrementCell(UINT incUnit, GRIDPOINT* pGridPoint);
	size_t OnEnumCells(GRIDCELL* gcArray, UINT numElements);
	void OnFillCells(WPARAM wParam, const GCFILLSTRUCT& fillStruct);
	void OnSortCells(WPARAM wParam, const GCSORTSTRUCT& sortStruct);
	void OnFilterCells(WPARAM wParam, const GCFILTERSTRUCT& filterStruct);
	void RecordUndoOperation(const GridEditOperation& op);
        void CopyGridCell(GRIDCELL& dest, GRIDCELL& src);
        void OnStreamOut(LPGCSTREAM pStream);
        void OnStreamIn(LPGCSTREAM pStream);
};

