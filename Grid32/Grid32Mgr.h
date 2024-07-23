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
	GRIDCELL m_defaultGridCell;

public:
	CGrid32Mgr();
	virtual ~CGrid32Mgr();
	bool Create(PGRIDCREATESTRUCT pGCS);
	static LRESULT CALLBACK Grid32_WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	void Paint(HWND hWnd, PAINTSTRUCT& ps);
	PGRIDCELL GetCell(UINT nRow, UINT nCol);
	void SetCell(UINT nRow, UINT nCol, const GRIDCELL& gc);
	void DeleteCell(UINT nRow, UINT nCol);
	void DeleteAllCells();
};

