// GridView.cpp : implementation file
//

#include "pch.h"
#include "Spreadsheet.h"
#include "GridView.h"


// CGridView
#define GRIDCTRL_ID 1

IMPLEMENT_DYNCREATE(CGridView, CView)

CGridView::CGridView()
{

}

CGridView::~CGridView()
{
}

BEGIN_MESSAGE_MAP(CGridView, CView)
	ON_WM_CREATE()
	ON_WM_SIZE()
    ON_WM_KEYDOWN()
    ON_WM_KEYUP()
    ON_WM_INPUT()
    ON_WM_SYSKEYUP()
    ON_WM_SYSKEYDOWN()
END_MESSAGE_MAP()


// CGridView drawing

void CGridView::OnDraw(CDC* pDC)
{
	CDocument* pDoc = GetDocument();
	// TODO: add draw code here
}


// CGridView diagnostics

#ifdef _DEBUG
void CGridView::AssertValid() const
{
	CView::AssertValid();
}

#ifndef _WIN32_WCE
void CGridView::Dump(CDumpContext& dc) const
{
	CView::Dump(dc);
}

BOOL CGridView::SetCurrentCell(LPCWSTR pwszRef, short nWhich)
{
    return m_wndGridCtrl.SetCurrentCell(pwszRef, nWhich);
}

BOOL CGridView::SetCurrentCell(UINT nRow, UINT nCol)
{
    return m_wndGridCtrl.SetCurrentCell(nRow, nCol);
}

#endif
#endif //_DEBUG

void CGridView::OnFindNext(LPCTSTR lpszFind, BOOL bNext, BOOL bCase)
{
    if (!lpszFind)
        return;
    GCFINDSTRUCT fs{};
    fs.m_cbSize = sizeof(GCFINDSTRUCT);
    fs.m_wsFindText = lpszFind;
    fs.m_bMatchCase = bCase;
    fs.m_bSearchForward = bNext;
    m_wndGridCtrl.GetCurrentCell(fs.m_startCell);
    if (!m_wndGridCtrl.FindText(fs))
        OnTextNotFound(lpszFind);
}
void CGridView::OnReplaceSel(LPCTSTR lpszFind, BOOL bNext, BOOL bCase, LPCTSTR lpszReplace)
{
    if (!lpszFind || !lpszReplace)
        return;
    GCREPLACESTRUCT rs{};
    rs.m_cbSize = sizeof(GCREPLACESTRUCT);
    rs.m_wsFindText = lpszFind;
    rs.m_wsReplaceText = lpszReplace;
    rs.m_bMatchCase = bCase;
    rs.m_bSearchForward = bNext;
    m_wndGridCtrl.GetCurrentCell(rs.m_startCell);
    if (!m_wndGridCtrl.ReplaceText(rs))
        OnTextNotFound(lpszFind);
}
void CGridView::OnReplaceAll(LPCTSTR lpszFind, LPCTSTR lpszReplace, BOOL bCase)
{
    if (!lpszFind || !lpszReplace)
        return;
    GCREPLACESTRUCT rs{};
    rs.m_cbSize = sizeof(GCREPLACESTRUCT);
    rs.m_wsFindText = lpszFind;
    rs.m_wsReplaceText = lpszReplace;
    rs.m_bMatchCase = bCase;
    rs.m_bSearchForward = TRUE;
    rs.m_startCell = {0,0};
    bool replaced = false;
    while (m_wndGridCtrl.ReplaceText(rs))
    {
        replaced = true;
        m_wndGridCtrl.GetCurrentCell(rs.m_startCell);
    }
    if (!replaced)
        OnTextNotFound(lpszFind);
}

void CGridView::OnTextNotFound(LPCTSTR lpszFind)
{
}

void CGridView::OnEditCut()
{
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
        m_wndGridCtrl.SendMessage(WM_CUT);
}

void CGridView::OnEditCopy()
{
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
        m_wndGridCtrl.SendMessage(WM_COPY);
}

void CGridView::OnEditPaste()
{
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
        m_wndGridCtrl.SendMessage(WM_PASTE);
}

void CGridView::OnEditSelectAll()
{
    GRIDSELECTION sel{ {0,0}, { (UINT)-1, (UINT)-1 } };
    m_wndGridCtrl.SetRangeFormat(sel, FONTINFO());
}

void CGridView::OnEditClear()
{
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
        m_wndGridCtrl.SendMessage(WM_CLEAR);
}

void CGridView::OnEditUndo()
{
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
        m_wndGridCtrl.SendMessage(WM_UNDO);
}

void CGridView::OnEditFind()
{
    OnFindNext(_T(""), TRUE, FALSE);
}

void CGridView::OnEditReplace()
{
    OnReplaceSel(_T(""), TRUE, FALSE, _T(""));
}

void CGridView::OnEditRepeat()
{
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
        m_wndGridCtrl.SendMessage(GM_REDO);
}


// CGridView message handlers


int CGridView::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
    GRIDCREATESTRUCT gcs;
    memset(&gcs, 0, sizeof(GRIDCREATESTRUCT));
    gcs.cbSize = sizeof(GRIDCREATESTRUCT);
    gcs.nWidth = 702;
    gcs.nHeight = 4096;
    gcs.nDefColWidth = 135;
    gcs.nDefRowHeight = 30;
    gcs.clrCurrentCellHighlightBox = RGB(255, 0, 0);

    if (CView::OnCreate(lpCreateStruct) == -1)
        return -1;

    if (!m_wndGridCtrl.Create(GS_SPREADSHEET | GRID_DEFAULT_STYLE, CRect(0, 0, 50, 50), this, GRIDCTRL_ID, gcs))
        return -1;

    return 0;
}


void CGridView::OnSize(UINT nType, int cx, int cy)
{
    CView::OnSize(nType, cx, cy);

    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
    {
        m_wndGridCtrl.MoveWindow(0, 0, cx, cy);
    }
}



void CGridView::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
    {
        // Forward the WM_KEYDOWN message to the grid control
        m_wndGridCtrl.SendMessage(WM_KEYDOWN, nChar, MAKELPARAM(nRepCnt, nFlags));
    }
    CView::OnKeyDown(nChar, nRepCnt, nFlags);
}


void CGridView::OnKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags)
{
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
    {
        // Forward the WM_KEYUP message to the grid control
        m_wndGridCtrl.SendMessage(WM_KEYUP, nChar, MAKELPARAM(nRepCnt, nFlags));
    }
    CView::OnKeyUp(nChar, nRepCnt, nFlags);
}


void CGridView::OnRawInput(UINT nInputcode, HRAWINPUT hRawInput)
{
    // This feature requires Windows XP or greater.
    // The symbol _WIN32_WINNT must be >= 0x0501.
    // TODO: Add your message handler code here and/or call default

    CView::OnRawInput(nInputcode, hRawInput);
}


void CGridView::OnSysKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags)
{
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
    {
        // Forward the WM_SYSKEYUP message to the grid control
        m_wndGridCtrl.SendMessage(WM_SYSKEYUP, nChar, MAKELPARAM(nRepCnt, nFlags));
    }
    CView::OnSysKeyUp(nChar, nRepCnt, nFlags);
}


void CGridView::OnSysKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
    {
        // Forward the WM_SYSKEYDOWN message to the grid control
        m_wndGridCtrl.SendMessage(WM_SYSKEYDOWN, nChar, MAKELPARAM(nRepCnt, nFlags));
    }
    CView::OnSysKeyDown(nChar, nRepCnt, nFlags);
}

void CGridView::ApplyFont(const FONTINFO& fi)
{
    m_wndGridCtrl.SetCurrentCellFormat(fi);
}
