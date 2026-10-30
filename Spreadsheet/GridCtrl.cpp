// GridCtrl.cpp : implementation file
//

#include "pch.h"
#include "Spreadsheet.h"
#include "GridCtrl.h"


// CGridCtrl

IMPLEMENT_DYNAMIC(CGridCtrl, CWnd)

CGridCtrl::CGridCtrl()
{

}

CGridCtrl::~CGridCtrl()
{
}

HMODULE CGridCtrl::InitGridControl()
{
	return LoadLibrary(_T("Grid32.dll"));
}

BOOL CGridCtrl::Create(DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, GRIDCREATESTRUCT& gcs)
{
    // Ensure the style includes WS_CHILD
    dwStyle |= WS_CHILD;
    if (pParentWnd == nullptr)
        return FALSE;

    // Call CWnd::Create with the required parameters
    return CWnd::Create(
        GRID_WNDCLASS_NAME,       // Class name
        NULL,           // No window name
        dwStyle,        // Window style
        rect,           // Position and size
        pParentWnd,     // Parent window
        nID,            // Control ID
        reinterpret_cast<CCreateContext *>(&gcs)            // Additional creation parameters
    );
}

BOOL CGridCtrl::SetCurrentCell(LPCWSTR pwszRef, short nWhich)
{
    return ::SendMessage(m_hWnd, GM_SETCURRENTCELL, nWhich, (LPARAM)pwszRef) != 0;
}

BOOL CGridCtrl::SetCurrentCell(UINT nRow, UINT nCol)
{
    GRIDPOINT pt{ nRow, nCol };
    return ::SendMessage(m_hWnd, GM_SETCURRENTCELL, 0, (LPARAM)&pt) != 0;
}

static GCCELLCHARFORMAT MakeCharFormat(const GRIDSELECTION& sel, const FONTINFO& fi, DWORD mask)
{
    GCCELLCHARFORMAT cf{};
    cf.m_cbSize = sizeof(GCCELLCHARFORMAT);
    cf.m_range = sel;
    cf.m_format = fi;
    cf.m_dwMask = mask;
    return cf;
}

BOOL CGridCtrl::SetCurrentCellFormat(const FONTINFO& fi, DWORD mask)
{
    GCCELLCHARFORMAT cf = MakeCharFormat(GRIDSELECTION{}, fi, mask);
    return ::SendMessage(m_hWnd, GM_SETCHARFORMAT, SCF_CURRENTCELL | SCF_MASKED, (LPARAM)&cf) != 0;
}

BOOL CGridCtrl::SetSelectionFormat(const FONTINFO& fi, DWORD mask)
{
    GCCELLCHARFORMAT cf = MakeCharFormat(GRIDSELECTION{}, fi, mask);
    return ::SendMessage(m_hWnd, GM_SETCHARFORMAT, SCF_SELECTION | SCF_MASKED, (LPARAM)&cf) != 0;
}

BOOL CGridCtrl::SetRangeFormat(const GRIDSELECTION& sel, const FONTINFO& fi, DWORD mask)
{
    GCCELLCHARFORMAT cf = MakeCharFormat(sel, fi, mask);
    return ::SendMessage(m_hWnd, GM_SETCHARFORMAT, SCF_RANGE, (LPARAM)&cf) != 0;
}

BOOL CGridCtrl::SetCurrentCellNumberFormat(UINT format)
{
    return ::SendMessage(m_hWnd, GM_SETFORMAT, SCF_CURRENTCELL, (LPARAM)format) != 0;
}

BOOL CGridCtrl::SetSelectionNumberFormat(UINT format)
{
    return ::SendMessage(m_hWnd, GM_SETFORMAT, SCF_SELECTION, (LPARAM)format) != 0;
}

BOOL CGridCtrl::GetCurrentCellFontInfo(FONTINFO& fi)
{
    return (BOOL)::SendMessage(m_hWnd, GM_GETCHARFORMAT, SCF_CURRENTCELL, (LPARAM)&fi);
}

BOOL CGridCtrl::SetSelectionHAlign(UINT halign)
{
    return ::SendMessage(m_hWnd, GM_SETALIGN, GA_HORIZ, (LPARAM)halign) != 0;
}

BOOL CGridCtrl::SetSelectionVAlign(UINT valign)
{
    return ::SendMessage(m_hWnd, GM_SETALIGN, GA_VERT, (LPARAM)valign) != 0;
}

BOOL CGridCtrl::SetSelectionWrap(BOOL wrap)
{
    return ::SendMessage(m_hWnd, GM_SETALIGN, GA_WRAP, (LPARAM)wrap) != 0;
}

UINT CGridCtrl::GetCurrentCellHAlign()
{
    return (UINT)::SendMessage(m_hWnd, GM_GETALIGN, GA_HORIZ, 0);
}

UINT CGridCtrl::GetCurrentCellVAlign()
{
    return (UINT)::SendMessage(m_hWnd, GM_GETALIGN, GA_VERT, 0);
}

BOOL CGridCtrl::GetCurrentCellWrap()
{
    return (BOOL)::SendMessage(m_hWnd, GM_GETALIGN, GA_WRAP, 0);
}

UINT CGridCtrl::GetCurrentCellFormat()
{
    return (UINT)::SendMessage(m_hWnd, GM_GETFORMAT, 0, 0);
}

BOOL CGridCtrl::SetSelectionBackground(COLORREF clr)
{
    return ::SendMessage(m_hWnd, GM_SETBACKGROUND, SCF_SELECTION, (LPARAM)clr) != 0;
}

COLORREF CGridCtrl::GetCurrentCellBackground()
{
    return (COLORREF)::SendMessage(m_hWnd, GM_GETBACKGROUND, 0, 0);
}

BOOL CGridCtrl::MergeSelection()
{
    return ::SendMessage(m_hWnd, GM_MERGECELLS, TRUE, 0) != 0;
}

BOOL CGridCtrl::UnmergeSelection()
{
    return ::SendMessage(m_hWnd, GM_MERGECELLS, FALSE, 0) != 0;
}

BOOL CGridCtrl::IsCurrentCellMerged()
{
    return ::SendMessage(m_hWnd, GM_GETMERGE, 0, 0) != 0;
}

BOOL CGridCtrl::SetCurrentCellText(LPCWSTR text)
{
    GRIDPOINT pt{};
    GetCurrentCell(pt);
    return ::SendMessage(m_hWnd, GM_SETCELLTEXT, (WPARAM)&pt, (LPARAM)text) != 0;
}

BOOL CGridCtrl::FindText(const GCFINDSTRUCT& fs)
{
    return ::SendMessage(m_hWnd, GM_FINDTEXT, 0, (LPARAM)&fs) != 0;
}

BOOL CGridCtrl::ReplaceText(const GCREPLACESTRUCT& rs)
{
    return ::SendMessage(m_hWnd, GM_REPLACETEXT, 0, (LPARAM)&rs) != 0;
}

UINT CGridCtrl::ReplaceAll(const GCREPLACESTRUCT& rs)
{
    return (UINT)::SendMessage(m_hWnd, GM_REPLACEALL, 0, (LPARAM)&rs);
}

void CGridCtrl::SelectAll()
{
    ::SendMessage(m_hWnd, GM_SELECTALL, 0, 0);
}

BOOL CGridCtrl::GetCurrentCell(GRIDPOINT& pt) const
{
    return ::SendMessage(m_hWnd, GM_GETCURRENTCELL, 0, (LPARAM)&pt) != 0;
}

BOOL CGridCtrl::GetSelection(GRIDSELECTION& sel) const
{
    return ::SendMessage(m_hWnd, GM_GETSELECTION, 0, (LPARAM)&sel) != 0;
}

BOOL CGridCtrl::GetUsedRange(GRIDSELECTION& range) const
{
    return ::SendMessage(m_hWnd, GM_GETUSEDRANGE, 0, (LPARAM)&range) != 0;
}

size_t CGridCtrl::GetColumnWidth(UINT nCol) const
{
    return (size_t)::SendMessage(m_hWnd, GM_GETCOLUMNWIDTH, nCol, 0);
}

size_t CGridCtrl::GetRowHeight(UINT nRow) const
{
    return (size_t)::SendMessage(m_hWnd, GM_GETROWHEIGHT, nRow, 0);
}

BOOL CGridCtrl::RenderRange(const GCRENDERSTRUCT& rs)
{
    return ::SendMessage(m_hWnd, GM_RENDERRANGE, 0, (LPARAM)&rs) != 0;
}

BOOL CGridCtrl::CommitEdit(BOOL bCommit)
{
    if (!m_hWnd)
        return FALSE;
    return ::SendMessage(m_hWnd, GM_COMMITEDIT, bCommit, 0) != 0;
}

BOOL CGridCtrl::CanUndo() const
{
    return ::SendMessage(m_hWnd, GM_CANUNDO, 0, 0) != 0;
}

BOOL CGridCtrl::CanRedo() const
{
    return ::SendMessage(m_hWnd, GM_CANREDO, 0, 0) != 0;
}

BOOL CGridCtrl::Undo()
{
    return ::SendMessage(m_hWnd, WM_UNDO, 0, 0) != 0;
}

BOOL CGridCtrl::Redo()
{
    return ::SendMessage(m_hWnd, GM_REDO, 0, 0) != 0;
}

BOOL CGridCtrl::StreamIn(GCSTREAM& stream)
{
    return ::SendMessage(m_hWnd, GM_STREAMIN, 0, (LPARAM)&stream) != 0;
}

BOOL CGridCtrl::StreamOut(GCSTREAM& stream)
{
    return ::SendMessage(m_hWnd, GM_STREAMOUT, 0, (LPARAM)&stream) != 0;
}

void CGridCtrl::ClearAll()
{
    ::SendMessage(m_hWnd, GM_CLEARALL, 0, 0);
}

DWORD CGridCtrl::GetLastError()
{
    return (DWORD)::SendMessage(m_hWnd, GM_GETLASTERROR, 0, 0);
}


BEGIN_MESSAGE_MAP(CGridCtrl, CWnd)
END_MESSAGE_MAP()



// CGridCtrl message handlers


