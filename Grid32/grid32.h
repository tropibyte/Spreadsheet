#pragma once


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

// Stream formats (GCSTREAM::m_dwFormat).
//
// Formulas: SF_SSF is the native format and keeps everything -- formula
// source, cached result, number format and cell styling. Every other format
// is an interchange format and carries what each cell *displays*: a formula
// cell exports its evaluated result, never its source, so the file reads the
// same in programs that cannot evaluate Grid32 formulas (Excel and
// LibreOffice export CSV the same way). Reading one back therefore yields
// values, not formulas.
//
//   SF_TEXT  out, in  Tab-separated plain text, CRLF after every row, no
//                     quoting. A tab or line break inside a cell is written
//                     as a space, so the table keeps its shape.
//   SF_RTF   out      An RTF table. Stream-in: GRID_ERROR_NOT_IMPLEMENTED.
//   SF_CSV   out, in  RFC 4180: comma-separated, CRLF after every row, every
//                     row as wide as the widest. A field containing the
//                     delimiter, a quote, CR or LF is quoted, and quotes
//                     inside it are doubled.
//   SF_TSV   out, in  As SF_CSV, with a tab delimiter.
//   SF_SSF   out, in  Native line-based format; see CGrid32Mgr::OnStreamOut.
//                     Numbers are written locale-independently and exactly.
//   SF_XML   out, in  Excel 2003 XML Spreadsheet (SpreadsheetML), which Excel
//                     opens directly: typed values plus the FMT_* number
//                     format of each cell.
//   SF_ODF   out, in  OpenDocument flat spreadsheet (.fods), which LibreOffice
//                     opens directly: typed values and display text.
//   SF_XLSX  GRID_ERROR_NOT_IMPLEMENTED in both directions. A genuine .xlsx is
//                     a ZIP package, and a ZIP cannot travel through the
//                     wide-character GCSTREAM buffer.
//
// The XML formats carry no encoding declaration: the buffer is UTF-16 and the
// host picks the byte encoding, so write it as UTF-16 with a BOM, or convert
// it to UTF-8 and write it without one.
#define SF_TEXT 					0x0001
#define SF_RTF                                          0x0002
#define SF_CSV                                          0x0004
#define SF_TSV                                          0x0008
#define SF_SSF                                          0x0010
#define SF_XML                                          0x0020
#define SF_ODF                                          0x0040
#define SF_XLSX                                         0x0080

//Fill constants
// GCFILLSTRUCT::m_dwFillDirection. With a direction, each column (down/up)
// or row (right/left) of the range copies its own edge cell — the top row
// for FILL_DOWN, the bottom row for FILL_UP, the left column for FILL_RIGHT,
// the right column for FILL_LEFT — and m_dwFillHow is ignored. Without one
// (0), the whole range receives the single source m_dwFillHow names.
#define FILL_UP						0x0001
#define FILL_DOWN					0x0002
#define FILL_LEFT					0x0004
#define FILL_RIGHT					0x0008

#define FILLBYREFCELL				1
#define FILLBYTEXT					2
#define FILLBYGCS					3

// GCFILLSTRUCT::m_dwFillRange: m_range, or the current selection.
#define FILLBYRANGE					0x0001
#define FILLBYCURRENTSELECTION		0x0002

// GCFILLSTRUCT::m_dwFillWhat. FILL_TEXT copies the content (a formula's
// source, or text that is re-typed so numbers stay numbers).
#define FILL_TEXT					0x0001
#define FILL_COLOR					0x0002      // background
#define FILL_FONT					0x0004      // face
#define FILL_POINTSIZE				0x0008
#define FILL_EFFECTS				0x0010      // weight, italic, underline, strikethrough
#define FILL_FORMAT					0x0020      // text colour, alignment, number format
#define FILL_ALL					0x003F

//Sorting constants
// GCSORTSTRUCT::m_dwWhich names the key column: 0 = the active cell's column
// if it lies in the sort range, else the range's first column;
// GCS_KEYCOLUMN(n) = column n (0-based).
#define GCS_KEYCOLUMN(col)			((DWORD)(col) + 1)
// GM_SORTCELLS wParam flags.
#define GCSF_SELECTIONONLY			0x0001      // move only the selected columns, not whole rows
#define GCSF_HEADERROW				0x0002      // the range's first row is a header and stays put
#define GCSF_MATCHCASE				0x0004


// Filtering constants
// GM_FILTERCELLS hides the rows below m_filterColumn.nRow (the header row,
// which stays visible) through the last used row whose cell in column
// m_filterColumn.nCol does not match. Filters accumulate; FILTER_NONE shows
// every row again.
#define FILTER_NONE					0x0000
#define FILTER_CONTAINS				0x0001
#define FILTER_EQUALS				0x0002
#define FILTER_STARTS_WITH			0x0004
#define FILTER_ENDS_WITH			0x0008
#define FILTER_MATCHCASE			0x0100      // OR into m_dwFilterType

// Grid32 WM_NCHITTEST constants

#define GRID_HTCORNER			    (HTHELP + 11)
#define GRID_HTCOLHEADER		    (HTHELP + 12)
#define GRID_HTROWHEADER			(HTHELP + 13)
#define GRID_HTCOLDIVIDER			(HTHELP + 14)
#define GRID_HTROWDIVIDER			(HTHELP + 15)
#define GRID_HTCELL					(HTHELP + 16)
#define GRID_HTOUTSIDE				(HTHELP + 17)


// Grid32 Messages
//
// Cell coordinates travel by pointer, never packed into a WPARAM/LPARAM, so
// no message caps the grid at 65535 rows or columns:
//   GM_SETCELL        wParam = const GRIDPOINT*, lParam = const GCCELLINFO*
//   GM_SETCELLTEXT    wParam = const GRIDPOINT*, lParam = LPCWSTR
//   GM_GETCELLTEXT    wParam = const GRIDPOINT*, lParam = GRID_GETTEXT*
//   GM_SETCURRENTCELL wParam = 0,                lParam = const GRIDPOINT*
//                     wParam = SETBYNAME / SETBYCOORDINATE, lParam = LPCWSTR
//   GM_GETCURRENTCELL lParam = GRIDPOINT* (out); returns TRUE

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
// wParam: capacity, lParam: GRIDPOINT* (out), or NULL to count. Fills in the
// positions of the populated cells in row-major order; returns how many.
#define GM_ENUMCELLS				(WM_USER + 22)
#define GM_SETCHARFORMAT			(WM_USER + 23)
#define GM_STREAMIN					(WM_USER + 24)
#define GM_STREAMOUT                            (WM_USER + 25)
#define GM_REPLACETEXT                          (WM_USER + 26)
#define GM_SETFORMAT                            (WM_USER + 27)
#define GM_GETCHARFORMAT                        (WM_USER + 28)
#define GM_SETALIGN                             (WM_USER + 29)
#define GM_GETALIGN                             (WM_USER + 30)
#define GM_GETFORMAT                            (WM_USER + 31)
#define GM_SELECTALL                            (WM_USER + 32)
#define GM_CLEARALL                             (WM_USER + 33)  // remove every cell; reset undo/redo, selection and current cell
// Close the inline editor. wParam: TRUE commits its text, FALSE discards it.
// Returns TRUE if an edit was open. Hosts send it before saving or closing
// so text still being typed is not lost.
#define GM_COMMITEDIT                           (WM_USER + 34)
// wParam: SCF_SELECTION / SCF_CURRENTCELL; lParam: COLORREF.
#define GM_SETBACKGROUND                        (WM_USER + 35)
// Returns the current cell's background COLORREF.
#define GM_GETBACKGROUND                        (WM_USER + 36)
// wParam: TRUE merges, FALSE unmerges; lParam: const GRIDSELECTION*, or NULL
// for the current selection. Returns TRUE on success.
#define GM_MERGECELLS                           (WM_USER + 37)
// lParam: MERGERANGE* (out, optional). Returns TRUE if the current cell is
// part of a merged range.
#define GM_GETMERGE                             (WM_USER + 38)
// lParam: GRIDSELECTION* (out). The bounding box of every non-empty cell;
// returns FALSE if the grid is empty.
#define GM_GETUSEDRANGE                         (WM_USER + 39)
// lParam: const GCRENDERSTRUCT*. Draws a block of cells into any DC
// (printing, print preview). Returns TRUE on success.
#define GM_RENDERRANGE                          (WM_USER + 40)
// lParam: const GCREPLACESTRUCT*. Replaces every match in the grid as one
// undo step; returns the number of occurrences replaced.
#define GM_REPLACEALL                           (WM_USER + 41)
// GM_SETCHARFORMAT / GM_SETFORMAT wParam flags
#define SCF_CURRENTCELL 0x0000
#define SCF_SELECTION 0x0001
#define SCF_RANGE 0x0002
// GM_SETCHARFORMAT: OR into wParam to pass an LPGCCELLCHARFORMAT whose
// m_dwMask names the attributes to apply (for SCF_CURRENTCELL/SCF_SELECTION
// too; m_range is read only for SCF_RANGE). Without it, SCF_CURRENTCELL and
// SCF_SELECTION take a PFONTINFO and apply every attribute.
#define SCF_MASKED 0x0100
// GCCELLCHARFORMAT::m_dwMask
#define GCFM_FACE      0x0001
#define GCFM_SIZE      0x0002
#define GCFM_COLOR     0x0004
#define GCFM_BOLD      0x0008      // bWeight
#define GCFM_ITALIC    0x0010
#define GCFM_UNDERLINE 0x0020
#define GCFM_STRIKE    0x0040
#define GCFM_ALL       0x007F
// GM_SETALIGN wParam — which axis the lParam value updates.
#define GA_HORIZ                                0
#define GA_VERT                                 1
#define GA_WRAP                                 2


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
	__MergeRange() : start{ 0, 0 }, end{ 0, 0 } {}
}MERGERANGE;

// Everything in this header is a plain struct: no C++ library types cross
// the DLL boundary, so a host need not share Grid32's compiler, CRT or
// iterator-debug level (M4). Zero-initialize before use (FONTINFO fi{};).
typedef struct __FONTINFO {
	WCHAR m_szFontFace[LF_FACESIZE];    // face name, null-terminated
	float m_fPointSize;
	COLORREF m_clrTextColor;
	BOOL bItalic, bUnderline, bStrikeThrough;
	UINT bWeight;
}FONTINFO, *PFONTINFO;

// Cell number-format codes. Determines how a typed value renders to display
// text. Text cells ignore the format.
#define FMT_GENERAL    0   // raw text (the user-entered representation)
#define FMT_NUMBER     1   // 0.00
#define FMT_INTEGER    2   // 0
#define FMT_CURRENCY   3   // $#,##0.00
#define FMT_PERCENT    4   // 0.00%
#define FMT_DATE_ISO   5   // yyyy-mm-dd
#define FMT_DATE_US    6   // m/d/yyyy
#define FMT_TIME       7   // hh:mm:ss
#define FMT_THOUSANDS  8   // #,##0.00 (no currency symbol)

// A cell as the API describes it (GM_SETCELL, and GCFILLSTRUCT's FILLBYGCS).
// m_dwMask (GCI_*) names the fields that are meaningful.
#define GCI_TEXT        0x0001  // m_pwszText: the content as typed; "=..." is a formula
#define GCI_FONT        0x0002  // m_font, every attribute
#define GCI_BACKGROUND  0x0004  // m_clrBackground
#define GCI_ALIGN       0x0008  // m_nJustification: DT_* alignment and wrap bits
#define GCI_FORMAT      0x0010  // m_nFormat: FMT_*
#define GCI_ALL         0x001F
typedef struct tagGCCELLINFO
{
	UINT     m_cbSize;
	DWORD    m_dwMask;
	LPCWSTR  m_pwszText;
	FONTINFO m_font;
	COLORREF m_clrBackground;
	UINT     m_nJustification;
	UINT     m_nFormat;
} GCCELLINFO, *LPGCCELLINFO;

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
    DWORD m_dwMask;     // GCFM_*; GCFM_ALL if m_cbSize predates the field
} GCCELLCHARFORMAT, *LPGCCELLCHARFORMAT;

// GM_RENDERRANGE. The block's column widths and row heights are scaled from
// screen pixels (96 DPI) to the target resolution, so the same call serves
// the screen, print preview and a 600 DPI printer. In print preview the DC
// drawn on is the screen's while the page is the printer's, so the host
// passes the printer's resolution; 0 means "the DC's own".
typedef struct tagGCRENDERSTRUCT
{
    UINT          m_cbSize;
    HDC           m_hDC;
    POINT         m_ptOrigin;     // top-left of the block in logical units
    GRIDSELECTION m_range;        // inclusive; hidden rows are skipped
    BOOL          m_bHeaders;     // draw column letters and row numbers
    BOOL          m_bGridLines;
    int           m_nLogPixelsX;  // target resolution, or 0
    int           m_nLogPixelsY;
} GCRENDERSTRUCT, *LPGCRENDERSTRUCT;


// Notification payload. SendGridNotification() zero-fills the struct, so
// handlers only need to read the fields relevant to the notification code:
//   NM_CLICK / NM_DBLCLK / NM_RCLICK / NM_HOVER  -> pt, cellPt
//   NM_KEYDOWN / NM_RETURN                       -> vkey, flags
//   NM_SETFOCUS / NM_KILLFOCUS                   -> (header only)
//   GN_SELCHANGE                                 -> cellPt (current), range (selection rect)
//   GN_CELLCHANGED                               -> cellPt or range (caller's choice)
//   GN_EDITSTART                                 -> cellPt
//   GN_EDITCOMMIT                                -> cellPt, text (transient, valid only for callback)
//   GN_EDITCANCEL                                -> cellPt
//   GN_COLWIDTHCHANGED                           -> cellPt.nCol, oldSize, newSize
//   GN_ROWHEIGHTCHANGED                          -> cellPt.nRow, oldSize, newSize
//   GN_COLHEADERCLICK                            -> cellPt.nCol, pt
//   GN_ROWHEADERCLICK                            -> cellPt.nRow, pt
//   GN_CORNERCLICK                               -> pt
//   GN_CONTENTCHANGED                            -> (header only) any edit; hosts mark the document dirty
typedef struct tagGRIDNMHDR : tagNMHDR {
	POINT         pt;          // mouse position (client coords)
	GRIDPOINT     cellPt;      // primary cell the event refers to
	GRIDSELECTION range;       // for events with a range payload
	UINT          vkey;        // virtual-key code for key events
	UINT          flags;       // key state flags / event-specific bits
	size_t        oldSize;     // pre-resize column width or row height
	size_t        newSize;     // post-resize column width or row height
	LPCWSTR       text;        // edit-commit text (valid only during callback)
	void*         pExtra;      // forward-compat slot
} GRIDNMHDR, * LPGRIDNMHDR;

// Grid notification codes. Below NM_LAST and below LVN_/TVN_/etc. common-control
// ranges so there's no collision with anything a host already handles. The 'U'
// matters: NMHDR::code is UINT, so the arithmetic must stay unsigned for the
// value the grid sends to match the value a handler compares against.
#define GN_FIRST              (0U - 2000U)
#define GN_SELCHANGE          (GN_FIRST -  0)
#define GN_CELLCHANGED        (GN_FIRST -  1)
#define GN_EDITSTART          (GN_FIRST -  2)
#define GN_EDITCOMMIT         (GN_FIRST -  3)
#define GN_EDITCANCEL         (GN_FIRST -  4)
#define GN_COLWIDTHCHANGED    (GN_FIRST -  5)
#define GN_ROWHEIGHTCHANGED   (GN_FIRST -  6)
#define GN_COLHEADERCLICK     (GN_FIRST -  7)
#define GN_ROWHEADERCLICK     (GN_FIRST -  8)
#define GN_CORNERCLICK        (GN_FIRST -  9)
#define GN_CONTENTCHANGED     (GN_FIRST - 10)  // any edit, paste, fill, sort, undo, load

typedef DWORD(CALLBACK* GCSTREAMCALLBACK)(struct tagGCSTREAM *);

// GM_STREAMIN replaces the grid's content: on success every existing cell is
// gone, the undo and redo stacks are empty, and the cursor is at A1. If the
// input cannot be parsed, m_dwError is set and the grid is left unchanged.
// (GRID_ERROR_OUT_OF_MEMORY part-way through a load is the one exception: the
// grid then holds what had been loaded.)
//
// GM_STREAMOUT writes the whole output and a terminating null. On success
// m_cbBuffOut is the number of bytes written, not counting the terminator. If
// the buffer is too small, m_dwError is GRID_ERROR_OUT_OF_RANGE, nothing is
// written (the buffer holds an empty string), and m_cbBuffOut is the size in
// bytes, terminator included, that the output needs -- so a caller retries at
// most once.
//
// Either message reports through m_dwError, and calls m_pfnCallback (if set)
// once the operation has finished.
typedef struct tagGCSTREAM {
	DWORD_PTR          m_dwCookie;
	DWORD              m_dwError;
	// Stream-in reads from this buffer; stream-out writes into it.
	// Must point to writable memory for GM_STREAMOUT.
	LPWSTR             m_pwszBuff;
	UINT			   m_cbSize, m_cbBuffSize, m_cbBuffOut;
	GCSTREAMCALLBACK   m_pfnCallback;
        DWORD             m_dwFormat;
} GCSTREAM, *LPGCSTREAM;

typedef struct tagGCFILLSTRUCT
{
	UINT  m_cbSize;
	DWORD m_dwFillDirection, m_dwFillHow;
	DWORD m_dwFillRange, m_dwFillWhat;
	GRIDPOINT m_refCell;        // FILLBYREFCELL
	GCCELLINFO m_cell;          // FILLBYGCS; FILLBYTEXT reads only m_cell.m_pwszText
	GRIDSELECTION m_range;      // FILLBYRANGE
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
        LPCWSTR m_pwszFilterText;    // Text to filter by
        DWORD m_dwFilterType;        // Type of filtering
} GCFILTERSTRUCT, *LPGCFILTERSTRUCT;

typedef struct tagGCFINDSTRUCT
{
        UINT m_cbSize;                // Size of this structure
        LPCWSTR m_pwszFindText;       // Text to search for
        GRIDPOINT m_startCell;        // Starting cell for the search
        BOOL m_bMatchCase;            // Case sensitive search flag
        BOOL m_bSearchForward;        // Direction flag: TRUE=forward, FALSE=backward
} GCFINDSTRUCT, *LPGCFINDSTRUCT;

typedef struct tagGCREPLACESTRUCT
{
        UINT m_cbSize;                // Size of this structure
        LPCWSTR m_pwszFindText;       // Text to search for
        LPCWSTR m_pwszReplaceText;    // Replacement text
        GRIDPOINT m_startCell;        // Starting cell for the search
        BOOL m_bMatchCase;            // Case sensitive search flag
        BOOL m_bSearchForward;        // Direction flag
} GCREPLACESTRUCT, *LPGCREPLACESTRUCT;

#endif // !_GRID32

