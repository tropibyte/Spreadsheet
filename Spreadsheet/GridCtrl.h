#pragma once

#include "../Grid32/grid32.h"

// CGridCtrl

class CGridCtrl : public CWnd
{
	DECLARE_DYNAMIC(CGridCtrl)

public:
	CGridCtrl();
	virtual ~CGridCtrl();
	static HMODULE InitGridControl();
	BOOL Create(DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, GRIDCREATESTRUCT &gcs);
	BOOL SetCurrentCell(LPCWSTR pwszRef, short nWhich);
    BOOL SetCurrentCell(UINT nRow, UINT nCol);
    BOOL SetCurrentCellFormat(const FONTINFO& fi);
    BOOL SetSelectionFormat(const FONTINFO& fi);
    BOOL SetRangeFormat(const GRIDSELECTION& sel, const FONTINFO& fi);
    BOOL SetCurrentCellNumberFormat(UINT format);
    BOOL SetSelectionNumberFormat(UINT format);
    BOOL GetCurrentCellFontInfo(FONTINFO& fi);
    BOOL SetSelectionHAlign(UINT halign);
    BOOL SetSelectionVAlign(UINT valign);
    BOOL SetSelectionWrap(BOOL wrap);
    UINT GetCurrentCellHAlign();    // DT_LEFT/CENTER/RIGHT, or 0 if none.
    UINT GetCurrentCellVAlign();    // DT_TOP/VCENTER/BOTTOM, or 0 if none.
    BOOL GetCurrentCellWrap();
    BOOL SetCurrentCellText(LPCWSTR text);
    BOOL FindText(const GCFINDSTRUCT& fs);
    BOOL ReplaceText(const GCREPLACESTRUCT& rs);
    BOOL GetCurrentCell(GRIDPOINT& pt) const;
    BOOL StreamIn(GCSTREAM& stream);
    BOOL StreamOut(GCSTREAM& stream);
    DWORD GetLastError();

protected:
	DECLARE_MESSAGE_MAP()
};


