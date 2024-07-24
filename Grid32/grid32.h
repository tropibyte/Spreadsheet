#pragma once

#include <string>

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
	size_t nDefRowHeight, nDefColWidth;
	LONG   style;
}GRIDCREATESTRUCT, *PGRIDCREATESTRUCT;

typedef struct __FONTINFO {
	std::wstring m_wsFontFace;
	float m_fPointSize;
	COLORREF m_clrTextColor;
	BOOL bItalic, bUnderline;
	UINT bWeight;
	__FONTINFO();
}FONTINFO;

struct cell_base
{
	FONTINFO fontInfo;
	COLORREF clrBackground, m_clrBorderColor;
	UINT m_nBorderWidth;
	int penStyle;
	std::wstring m_wsName;
	cell_base();
};

typedef struct __ROWINFO : public cell_base {
	size_t nHeight;
	__ROWINFO() : nHeight(0) {}
}ROWINFO;

typedef struct __COLINFO : public cell_base {
	size_t nWidth;
	__COLINFO() : nWidth(0) {}
}COLINFO;

typedef struct __GRIDCELL : public cell_base {
	std::wstring m_wsText;
}GRIDCELL, *PGRIDCELL;

#endif // !_GRID32
