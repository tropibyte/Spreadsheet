#pragma once

#ifndef _GRID32
#define _GRID32

#define GRID_WNDCLASS_NAME _T("Grid32")

// Grid32 Window styles

#define GS_TABULAR					0x0000
#define GS_DRAWLINES				0x0001
#define GS_AUTOHSCROLL				0x0002
#define GS_AUTOVSCROLL				0x0004
#define GS_ROWHEADER				0x0008
#define GS_COLHEADER				0x0010
#define GS_SPREADSHEETHEADER		0x0020

#define GS_SPREADSHEET				(GS_SPREADSHEETHEADER | GS_ROWHEADER | GS_COLHEADER | GS_AUTOHSCROLL | GS_AUTOVSCROLL | GS_DRAWLINES | GS_TABULAR)

typedef struct {
	size_t cbSize;
	size_t nWidth, nHeight;
}GRIDCREATESTRUCT, *PGRIDCREATESTRUCT;



#endif // !_GRID32
