#pragma once

// The cell store's own types. They hold std::wstring, so they stay inside
// Grid32.dll: grid32.h, the header hosts compile against, describes cells
// and fonts with plain structs (FONTINFO, GCCELLINFO) and LPCWSTR instead.
// A host built with a different compiler, CRT or _ITERATOR_DEBUG_LEVEL used
// to share these std::wstring layouts and heaps through SendMessage
// (AUDIT_2026-09 M4).

#include <string>
#include "grid32.h"

// A cell's font as the grid stores it. FONTINFO is the same information with
// a fixed-size face name, for the public API.
struct CELLFONT {
	std::wstring m_wsFontFace;
	float m_fPointSize;
	COLORREF m_clrTextColor;
	BOOL bItalic, bUnderline, bStrikeThrough;
	UINT bWeight;
	CELLFONT() : m_fPointSize(0), m_clrTextColor(0), bItalic(FALSE), bUnderline(FALSE), bStrikeThrough(FALSE), bWeight(0)
	{}
};

inline CELLFONT CellFontFrom(const FONTINFO& fi)
{
	CELLFONT cf;
	size_t len = 0;
	while (len < LF_FACESIZE && fi.m_szFontFace[len])
		++len;
	cf.m_wsFontFace.assign(fi.m_szFontFace, len);
	cf.m_fPointSize = fi.m_fPointSize;
	cf.m_clrTextColor = fi.m_clrTextColor;
	cf.bItalic = fi.bItalic;
	cf.bUnderline = fi.bUnderline;
	cf.bStrikeThrough = fi.bStrikeThrough;
	cf.bWeight = fi.bWeight;
	return cf;
}

inline FONTINFO FontInfoFrom(const CELLFONT& cf)
{
	FONTINFO fi{};
	// Face names are at most LF_FACESIZE - 1 characters, as in a LOGFONT.
	const size_t len = cf.m_wsFontFace.size() < LF_FACESIZE - 1 ? cf.m_wsFontFace.size() : LF_FACESIZE - 1;
	for (size_t i = 0; i < len; ++i)
		fi.m_szFontFace[i] = cf.m_wsFontFace[i];
	fi.m_szFontFace[len] = 0;
	fi.m_fPointSize = cf.m_fPointSize;
	fi.m_clrTextColor = cf.m_clrTextColor;
	fi.bItalic = cf.bItalic;
	fi.bUnderline = cf.bUnderline;
	fi.bStrikeThrough = cf.bStrikeThrough;
	fi.bWeight = cf.bWeight;
	return fi;
}

struct cell_base
{
	CELLFONT fontInfo;
	COLORREF clrBackground, m_clrBorderColor;
	UINT m_nBorderWidth;
	int penStyle;
	UINT justification;
	UINT m_nFormat;             // FMT_* constant — display formatting
	std::wstring m_wsName;
	cell_base() : clrBackground(0), m_clrBorderColor(0), m_nBorderWidth(0), penStyle(0), justification(0), m_nFormat(FMT_GENERAL)
	{}
};

// Underlying value type for a cell. m_wsText always carries the displayed
// representation; cells that are typed (Number/Date/Boolean/Formula) also
// populate m_dValue with the canonical scalar value. CT_Text cells leave
// m_dValue at 0.0 and treat m_wsText as authoritative.
enum CellType {
        CT_Text    = 0,
        CT_Number  = 1,
        CT_Date    = 2,   // m_dValue is days since 1899-12-30 (Excel serial)
        CT_Boolean = 3,
        CT_Formula = 4,   // m_wsFormula carries the source; m_dValue is the last evaluated number; m_wsText is the displayed result
        CT_Error   = 5,
};

typedef struct __GRIDCELL : public cell_base {
        std::wstring m_wsText;
        std::wstring m_wsFormula;
        bool m_bFormula;
        CellType m_eType;
        double m_dValue;
        MERGERANGE mergeRange;
        __GRIDCELL() : m_bFormula(false), m_eType(CT_Text), m_dValue(0.0) {}
}GRIDCELL, *PGRIDCELL;
