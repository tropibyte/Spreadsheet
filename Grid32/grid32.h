#pragma once

#include <string>

struct tagGCSTREAM;

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
#define INCREMENT_EDGE				100

#define SCROLL_VERT					1
#define SCROLL_HORZ					2

#define SETBYNAME					1
#define SETBYCOORDINATE				2

// Stream format
#define SF_TEXT 					0x0001
#define SF_RTF                                          0x0002
#define SF_CSV                                          0x0004
#define SF_TSV                                          0x0008
#define SF_SSF                                          0x0010
#define SF_XML                                          0x0020
#define SF_ODF                                          0x0040
#define SF_XLSX                                         0x0080

//Fill constants
#define FILL_UP						0x0001
#define FILL_DOWN					0x0002
#define FILL_LEFT					0x0004
#define FILL_RIGHT					0x0008

#define FILLBYREFCELL				1
#define FILLBYTEXT					2
#define FILLBYGCS					3

#define FILLBYRANGE					0x0001
#define FILLBYCURRENTSELECTION		0x0002

#define FILL_TEXT					0x0001
#define FILL_COLOR					0x0002
#define FILL_FONT					0x0004
#define FILL_POINTSIZE				0x0008
#define FILL_EFFECTS				0x0010

//Sorting constants


// Filtering constants

#define FILTER_CONTAINS				0x0001
#define FILTER_EQUALS				0x0002
#define FILTER_STARTS_WITH			0x0004
#define FILTER_ENDS_WITH			0x0008

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
#define GM_SETCURRENTCELL			(WM_USER + 15)
#define GM_GETCURRENTCELL			(WM_USER + 16)
#define GM_INCREMENTCURRENTCELL		(WM_USER + 17)
#define GM_FILLCELLS				(WM_USER + 18)
#define GM_SORTCELLS                (WM_USER + 19)
#define GM_FILTERCELLS              (WM_USER + 20)
#define GM_FINDTEXT					(WM_USER + 21)
#define GM_ENUMCELLS				(WM_USER + 22)
#define GM_SETCHARFORMAT			(WM_USER + 23)
#define GM_STREAMIN					(WM_USER + 24)
#define GM_STREAMOUT                            (WM_USER + 25)
#define GM_REPLACETEXT                          (WM_USER + 26)
// GM_SETCHARFORMAT wParam flags
#define SCF_CURRENTCELL 0x0000
#define SCF_SELECTION 0x0001
#define SCF_RANGE 0x0002


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
	COLORREF clrCurrentCellHighlightBox, clrSelectionOverlay;
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
}FONTINFO, *PFONTINFO;

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
        std::wstring m_wsFormula;
        bool m_bFormula;
        MERGERANGE mergeRange;
        __GRIDCELL() : m_bFormula(false) {}
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
typedef struct tagGCCELLCHARFORMAT
{
    UINT m_cbSize;
    GRIDSELECTION m_range;
    FONTINFO m_format;
} GCCELLCHARFORMAT, *LPGCCELLCHARFORMAT;


typedef struct tagGRIDNMHDR : tagNMHDR {
	// Add custom members for your notification data
	// For example:
	POINT pt; // Mouse position
	GRIDPOINT cellPt; // Grid cell coordinates
	void* pExtra; //Extra info as needed
} GRIDNMHDR, * LPGRIDNMHDR;

typedef DWORD(CALLBACK* GCSTREAMCALLBACK)(struct tagGCSTREAM *);

typedef struct tagGCSTREAM {
	DWORD_PTR          m_dwCookie;
	DWORD              m_dwError;
	LPCWSTR			   m_pwszBuff;
	UINT			   m_cbSize, m_cbBuffSize, m_cbBuffOut;
	GCSTREAMCALLBACK   m_pfnCallback;
        DWORD             m_dwFormat;
} GCSTREAM, *LPGCSTREAM;

typedef struct tagGCFILLSTRUCT
{
	UINT  m_cbSize;
	DWORD m_dwFillDirection, m_dwFillHow;
	DWORD m_dwFillRange, m_dwFillWhat;
	GRIDPOINT m_refCell;
	GRIDCELL refGC;
	std::wstring& m_wsText = refGC.m_wsText;
}GCFILLSTRUCT, *LPGCFILLSTRUCT;

typedef struct tagGCSORTSTRUCT
{
	UINT m_cbSize;        // Size of this structure
	GRIDSELECTION m_sortSel; // Column to be sorted
	BOOL m_bAscending;    // Sort order: true for ascending, false for descending
	DWORD m_dwWhich;
} GCSORTSTRUCT, *LPGCSORTSTRUCT;

typedef struct tagGCFILTERSTRUCT
{
        UINT m_cbSize;         // Size of this structure
        GRIDPOINT m_filterColumn;  // Column to apply the filter
        std::wstring m_wsFilterText; // Text to filter by
        DWORD m_dwFilterType;        // Type of filtering
} GCFILTERSTRUCT, *LPGCFILTERSTRUCT;

typedef struct tagGCFINDSTRUCT
{
        UINT m_cbSize;                // Size of this structure
        std::wstring m_wsFindText;    // Text to search for
        GRIDPOINT m_startCell;        // Starting cell for the search
        BOOL m_bMatchCase;            // Case sensitive search flag
        BOOL m_bSearchForward;        // Direction flag: TRUE=forward, FALSE=backward
} GCFINDSTRUCT, *LPGCFINDSTRUCT;

typedef struct tagGCREPLACESTRUCT
{
        UINT m_cbSize;                // Size of this structure
        std::wstring m_wsFindText;    // Text to search for
        std::wstring m_wsReplaceText; // Replacement text
        GRIDPOINT m_startCell;        // Starting cell for the search
        BOOL m_bMatchCase;            // Case sensitive search flag
        BOOL m_bSearchForward;        // Direction flag
} GCREPLACESTRUCT, *LPGCREPLACESTRUCT;

#endif // !_GRID32

