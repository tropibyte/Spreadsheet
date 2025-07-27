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
    return ::SendMessage(m_hWnd, GM_SETCURRENTCELL, 0, MAKELPARAM(nRow, nCol)) != 0;
}

BOOL CGridCtrl::SetCurrentCellFormat(const FONTINFO& fi)
{
    return ::SendMessage(m_hWnd, GM_SETCHARFORMAT, 0, (LPARAM)&fi) != 0;
}

BOOL CGridCtrl::SetSelectionFormat(const FONTINFO& fi)
{
    return ::SendMessage(m_hWnd, GM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&fi) != 0;
}

BOOL CGridCtrl::SetRangeFormat(const GRIDSELECTION& sel, const FONTINFO& fi)
{
    GCCELLCHARFORMAT cf{};
    cf.m_cbSize = sizeof(GCCELLCHARFORMAT);
    cf.m_range = sel;
    cf.m_format = fi;
    return ::SendMessage(m_hWnd, GM_SETCHARFORMAT, SCF_RANGE, (LPARAM)&cf) != 0;
}

BOOL CGridCtrl::FindText(const GCFINDSTRUCT& fs)
{
    return ::SendMessage(m_hWnd, GM_FINDTEXT, 0, (LPARAM)&fs) != 0;
}

BOOL CGridCtrl::ReplaceText(const GCREPLACESTRUCT& rs)
{
    return ::SendMessage(m_hWnd, GM_REPLACETEXT, 0, (LPARAM)&rs) != 0;
}

BOOL CGridCtrl::GetCurrentCell(GRIDPOINT& pt) const
{
    LRESULT res = ::SendMessage(m_hWnd, GM_GETCURRENTCELL, 0, 0);
    pt.nCol = LOWORD(res);
    pt.nRow = HIWORD(res);
    return TRUE;
}

BOOL CGridCtrl::StreamIn(GCSTREAM& stream)
{
    return ::SendMessage(m_hWnd, GM_STREAMIN, 0, (LPARAM)&stream) != 0;
}

BOOL CGridCtrl::StreamOut(GCSTREAM& stream)
{
    return ::SendMessage(m_hWnd, GM_STREAMOUT, 0, (LPARAM)&stream) != 0;
}

DWORD CGridCtrl::GetLastError()
{
    return (DWORD)::SendMessage(m_hWnd, GM_GETLASTERROR, 0, 0);
}


BEGIN_MESSAGE_MAP(CGridCtrl, CWnd)
END_MESSAGE_MAP()



// CGridCtrl message handlers


