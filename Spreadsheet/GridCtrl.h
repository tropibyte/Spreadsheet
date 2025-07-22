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
    BOOL FindText(const GCFINDSTRUCT& fs);
    BOOL ReplaceText(const GCREPLACESTRUCT& rs);
    BOOL GetCurrentCell(GRIDPOINT& pt) const;
    DWORD GetLastError();

protected:
	DECLARE_MESSAGE_MAP()
};


