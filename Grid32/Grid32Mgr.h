#pragma once

#include <string>
#include <vector>
#include <map>

#include "grid32.h"

#define ID_CELL_EDIT		8
#define HOVER_TIMER			1

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
	BOOL m_bRedraw;
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

public:
	CGrid32Mgr();
	virtual ~CGrid32Mgr();
	std::wstring GetColumnLabel(size_t index, size_t maxWidth);
	bool Create(PGRIDCREATESTRUCT pGCS);
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
	void DrawCells(HDC hDC);
	void DrawVoidSpace(HDC hDC);
	void DrawSizingLine(HDC hDC);
	bool IsCellMerged(UINT row, UINT col);
	void DrawMergedCell(HDC hDC, const MERGERANGE& mergeRange);
	void SelectMergedRange(UINT row, UINT col);
	size_t CalculatedColumnDistance(size_t start, size_t end);
	size_t CalculatedRowDistance(size_t start, size_t end);
	void GetCurrentCellRect(RECT& rect);
	PGRIDCELL GetCell(UINT nRow, UINT nCol);
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
	void OnLButtonDown(UINT nFlags, int x, int y);
	void OnLButtonUp(UINT nFlags, int x, int y);
	void OnMouseMove(UINT nFlags, int x, int y);
	void OnNcMouseMove(UINT nFlags, int x, int y);
	LRESULT OnNcHitTest(UINT nFlags, int x, int y, bool &bHitTested);
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
	void OnCanUndo();
	void OnCanRedo();
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
	void SetLastError(DWORD dwNewError) {	dwError = dwNewError;	}
	std::wstring FormatCellReference(LPCWSTR pwszRef);
	bool SetCurrentCell(UINT nRow, UINT nCol);
	BOOL Invalidate(bool bErasebkgnd = false);
	bool SetColumnWidth(size_t nCol, size_t newWidth);
	size_t GetColumnWidth(size_t nCol);
	bool SetRowHeight(size_t nRow, size_t newHeight);
	size_t GetRowHeight(size_t nRow);
	BOOL OnGetCellText(const GRIDPOINT& point, GRID_GETTEXT* text);
	void ShowEditControl();
	void SendGridNotification(INT code, GRIDNMHDR *pNMHDRInfo = nullptr);
	void GetMousePosition(POINT& pt);
	void OnTimer(UINT_PTR idEvent);
	bool IsOverColumnDivider(int x, int y);
	bool IsOverRowDivider(int x, int y);
	long GetColumnStartPos(UINT nCol);
	long GetRowStartPos(UINT nRow);
};

