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
    // mask (GCFM_*) names the attributes to apply; the rest of each cell's
    // font is left as it is.
    BOOL SetCurrentCellFormat(const FONTINFO& fi, DWORD mask = GCFM_ALL);
    BOOL SetSelectionFormat(const FONTINFO& fi, DWORD mask = GCFM_ALL);
    BOOL SetRangeFormat(const GRIDSELECTION& sel, const FONTINFO& fi, DWORD mask = GCFM_ALL);
    BOOL SetCurrentCellNumberFormat(UINT format);
    BOOL SetSelectionNumberFormat(UINT format);
    BOOL GetCurrentCellFontInfo(FONTINFO& fi);
    BOOL SetSelectionHAlign(UINT halign);
    BOOL SetSelectionVAlign(UINT valign);
    BOOL SetSelectionWrap(BOOL wrap);
    UINT GetCurrentCellHAlign();    // DT_LEFT/CENTER/RIGHT
    UINT GetCurrentCellVAlign();    // DT_TOP/VCENTER/BOTTOM
    BOOL GetCurrentCellWrap();
    UINT GetCurrentCellFormat();    // FMT_* (FMT_GENERAL if no cell).
    BOOL SetSelectionBackground(COLORREF clr);
    COLORREF GetCurrentCellBackground();
    BOOL MergeSelection();
    BOOL UnmergeSelection();
    BOOL IsCurrentCellMerged();
    BOOL SetCurrentCellText(LPCWSTR text);
    BOOL FindText(const GCFINDSTRUCT& fs);
    BOOL ReplaceText(const GCREPLACESTRUCT& rs);
    UINT ReplaceAll(const GCREPLACESTRUCT& rs);    // occurrences replaced
    void SelectAll();
    BOOL GetCurrentCell(GRIDPOINT& pt) const;
    BOOL GetSelection(GRIDSELECTION& sel) const;
    BOOL GetUsedRange(GRIDSELECTION& range) const;
    size_t GetColumnWidth(UINT nCol) const;
    size_t GetRowHeight(UINT nRow) const;     // 0 for a row a filter hid
    BOOL RenderRange(const GCRENDERSTRUCT& rs);
    // Close the inline editor: bCommit writes its text, otherwise it is
    // discarded. Returns TRUE if an edit was open.
    BOOL CommitEdit(BOOL bCommit = TRUE);
    BOOL CanUndo() const;
    BOOL CanRedo() const;
    BOOL Undo();
    BOOL Redo();
    BOOL StreamIn(GCSTREAM& stream);
    BOOL StreamOut(GCSTREAM& stream);
    // GM_CLEARALL: remove every cell, empty undo/redo, cursor back to A1.
    void ClearAll();
    DWORD GetLastError();

protected:
	DECLARE_MESSAGE_MAP()
};


