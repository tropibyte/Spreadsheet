#pragma once

#include "grid32.h"

class CGrid32Mgr
{
	GRIDCREATESTRUCT gcs;
public:
	CGrid32Mgr();
	virtual ~CGrid32Mgr();
	bool Create(PGRIDCREATESTRUCT pGCS);
	static LRESULT CALLBACK Grid32_WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	void Paint(HWND hWnd, PAINTSTRUCT& ps);
};

