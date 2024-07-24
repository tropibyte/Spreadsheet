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
	RECT totalGridCellRect;
	GRIDPOINT m_selectionPoint, m_beginDrawPoint;
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
	bool CanDrawRect(RECT clientRect, RECT rect);
	void DrawSelectionBox(HDC hDC, const RECT& clientRect);
	void DrawHeader(HDC hDC, const RECT& rect);
	void DrawCornerCell(HDC hDC);
	void DrawHeaderButton(HDC hDC, COLORREF background, COLORREF border, RECT coordinates, bool bRaised);
	void DrawRowHeaders(HDC hDC, const RECT& rect);
	void DrawColHeaders(HDC hDC, const RECT& rect);
	void DrawGrid(HDC hDC, const RECT& clientRect);
	void DrawCells(HDC hDC, const RECT& clientRect);
	PGRIDCELL GetCell(UINT nRow, UINT nCol);
	void SetCell(UINT nRow, UINT nCol, const GRIDCELL& gc);
	void DeleteCell(UINT nRow, UINT nCol);
	void DeleteAllCells();
	void OnHScroll(UINT nSBCode, UINT nPos, HWND hScrollBar);
	void OnVScroll(UINT nSBCode, UINT nPos, HWND hScrollBar);
	void SetScrollRanges();
	void IncrementSelectedCell(long nRow, long nCol, short nWhich);
	void OnMove(int x, int y);
	void OnSize(UINT nType, int cx, int cy);
	bool OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
};

