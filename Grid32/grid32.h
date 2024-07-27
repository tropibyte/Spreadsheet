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
#define GS_HIGHLIGHTSELECTION		0x0040

#define GS_SPREADSHEET				(GS_SPREADSHEETHEADER | GS_HIGHLIGHTSELECTION | GS_ROWHEADER | GS_COLHEADER | GS_AUTOHSCROLL | GS_AUTOVSCROLL | GS_DRAWLINES | GS_TABULAR)

// Grid32 constants

#define INCREMENT_SINGLE			1
#define INCREMENT_PAGE				10

#define SCROLL_VERT					1
#define SCROLL_HORZ					2

#define SETBYNAME					1
#define SETBYCOORDINATE				2

// Grid32 Messages

#define GM_REDO						(WM_USER + 1)
#define GM_CANUNDO					(WM_USER + 2)
#define GM_CANREDO					(WM_USER + 3)
#define GM_SETCELL					(WM_USER + 4)
#define GM_GETLASTERROR				(WM_USER + 5)

// Grid32 Error

#define GRID_ERROR_INVALID_CELL_REFERENCE   30001
#define GRID_ERROR_NOT_IMPLEMENTED			30002


// Grid32 structures

typedef struct {
	size_t cbSize;
	size_t nWidth, nHeight;
	size_t nDefRowHeight, nDefColWidth;
	LONG   style;
	COLORREF clrSelectBox;
}GRIDCREATESTRUCT, *PGRIDCREATESTRUCT;

typedef struct __FONTINFO {
	std::wstring m_wsFontFace;
	float m_fPointSize;
	COLORREF m_clrTextColor;
	BOOL bItalic, bUnderline, bStrikeThrough;
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

typedef struct __GRIDPOINT
{
	UINT nRow, nCol;
}GRIDPOINT;

typedef struct{
	long width, height;
}GRIDSIZE;

typedef struct {
	GRIDPOINT start, end;
	size_t nWidth, nHeight;
}PAGESTAT;

typedef struct {
	LPCWSTR pwszRef;
	short nWhich;
}GRIDSETCELL;
#endif // !_GRID32
