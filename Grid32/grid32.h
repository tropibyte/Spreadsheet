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

// Grid32 WM_NCHITTEST constants

#define GRID_HTCORNER			    (HTHELP + 11)
#define GRID_HTCOLHEADER		    (HTHELP + 12)
#define GRID_HTROWHEADER			(HTHELP + 13)
#define GRID_HTCOLDIVIDER			(HTHELP + 14)
#define GRID_HTROWDIVIDER			(HTHELP + 15)
#define GRID_HTCELL					(HTHELP + 16)
#define GRID_HTOUTSIDE				(HTHELP + 17)


// Grid32 Messages

#define GM_REDO						(WM_USER + 1)
#define GM_CANUNDO					(WM_USER + 2)
#define GM_CANREDO					(WM_USER + 3)
#define GM_SETCELL					(WM_USER + 4)
#define GM_GETLASTERROR				(WM_USER + 5)
#define GM_SETCOLUMNWIDTH			(WM_USER + 6)
#define GM_GETCOLUMNWIDTH			(WM_USER + 7)
#define GM_SETROWHEIGHT				(WM_USER + 8)
#define GM_GETROWHEIGHT				(WM_USER + 9)
#define GM_SETSELECTION				(WM_USER + 10)
#define GM_GETSELECTION				(WM_USER + 11)
#define GM_SETCELLTEXT				(WM_USER + 12)
#define GM_GETCELLTEXT				(WM_USER + 13)
#define GM_SETREDRAW				(WM_SETREDRAW)
#define GM_GETREDRAW				(WM_USER + 14)

// Grid32 Error

#define GRID_ERROR_INVALID_CELL_REFERENCE			30001
#define GRID_ERROR_NOT_IMPLEMENTED					30002
#define GRID_ERROR_OUT_OF_RANGE						30003
#define GRID_ERROR_INVALID_PARAMETER				30004
#define GRID_ERROR_NOT_EXIST						30005
#define GRID_ERROR_OUT_OF_MEMORY					30006


// Grid32 structures

typedef struct {
	size_t cbSize;
	size_t nWidth, nHeight;
	size_t nDefRowHeight, nDefColWidth;
	LONG   style;
	COLORREF clrSelectBox;
}GRIDCREATESTRUCT, *PGRIDCREATESTRUCT;

typedef struct __GRIDPOINT
{
	UINT nRow, nCol;
}GRIDPOINT;

typedef struct __MergeRange {
	GRIDPOINT start, end;
}MERGERANGE;

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
	UINT justification;
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
	MERGERANGE mergeRange;
}GRIDCELL, *PGRIDCELL;

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
}GRIDSETCURCELL;

typedef struct {
	GRIDPOINT start, end;
}GRIDSELECTION;

typedef struct {
	LPWSTR wszBuff;
	UINT nLen;
}GRID_GETTEXT;

typedef struct tagGRIDNMHDR : tagNMHDR {
	// Add custom members for your notification data
	// For example:
	POINT pt; // Mouse position
	GRIDPOINT cellPt; // Grid cell coordinates
	void* pExtra; //Extra info as needed
} GRIDNMHDR, * LPGRIDNMHDR;

#endif // !_GRID32
