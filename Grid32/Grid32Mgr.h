#pragma once

#include <string>
#include <vector>
#include <map>

#include "grid32.h"

class CGrid32Mgr
{
protected:
	GRIDCREATESTRUCT gcs;
	std::map<std::pair<UINT, UINT>, PGRIDCELL> mapCells;
	GRIDCELL m_defaultGridCell, m_cornerCell;
	ROWINFO* pRowInfoArray;
	COLINFO* pColInfoArray;
	HWND m_hWndGrid;
	long nRowHeaderHeight, nColHeaderWidth;
	RECT totalGridCellRect, m_clientRect;
	struct {
		long width, height;
	}visibleGrid;
	GRIDPOINT m_selectionPoint, m_visibleTopLeft;
	POINT m_scrollDifference;

public:
	CGrid32Mgr();
	virtual ~CGrid32Mgr();
	std::wstring GetColumnLabel(size_t index, size_t maxWidth);
	bool Create(PGRIDCREATESTRUCT pGCS);
	static LRESULT CALLBACK Grid32_WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	void Paint(PAINTSTRUCT& ps);
	void CalculateTotalGridRect();
	void OffsetRectByScroll(RECT&/*rect*/);
	bool CanDrawRect(RECT rect);
	void DrawSelectionBox(HDC hDC);
	void DrawHeader(HDC hDC);
	void DrawCornerCell(HDC hDC);
	void DrawHeaderButton(HDC hDC, COLORREF background, COLORREF border, RECT coordinates, bool bRaised);
	void DrawRowHeaders(HDC hDC);
	void DrawColHeaders(HDC hDC);
	void DrawGrid(HDC hDC);
	void DrawCells(HDC hDC);
	void DrawVoidSpace(HDC hDC);
	size_t CalculatedColumnDistance(size_t start, size_t end);
	size_t CalculatedRowDistance(size_t start, size_t end);
	PGRIDCELL GetCell(UINT nRow, UINT nCol);
	void SetCell(UINT nRow, UINT nCol, const GRIDCELL& gc);
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
	bool IsCellVisible(UINT nRow, UINT nCol);
	size_t MaxBeginRow();
	size_t MaxBeginColumn();
	void CalculatePageStats(PAGESTAT& pageStat);
	long GetActualRowHeaderHeight() { return gcs.style & GS_ROWHEADER ? nRowHeaderHeight : 0; }
	long GetActualColHeaderWidth() { return gcs.style & GS_COLHEADER ? nColHeaderWidth : 0; }
};

