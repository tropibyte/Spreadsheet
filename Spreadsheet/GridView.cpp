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

BOOL CGridView::SetCell(LPCWSTR pwszRef, short nWhich)
{
    return m_wndGridCtrl.SetCell(pwszRef, nWhich);
}

BOOL CGridView::SetCell(UINT nRow, UINT nCol)
{
    return m_wndGridCtrl.SetCell(nRow, nCol);
}

#endif
#endif //_DEBUG

void CGridView::OnFindNext(LPCTSTR lpszFind, BOOL bNext, BOOL bCase)
{
}
void CGridView::OnReplaceSel(LPCTSTR lpszFind, BOOL bNext, BOOL bCase, LPCTSTR lpszReplace)
{
}
void CGridView::OnReplaceAll(LPCTSTR lpszFind, LPCTSTR lpszReplace, BOOL bCase)
{
}

void CGridView::OnTextNotFound(LPCTSTR lpszFind)
{
}

void CGridView::OnEditCut()
{
}

void CGridView::OnEditCopy()
{
}

void CGridView::OnEditPaste()
{
}

void CGridView::OnEditSelectAll()
{
}

void CGridView::OnEditClear()
{
}

void CGridView::OnEditUndo()
{
}

void CGridView::OnEditFind()
{
}

void CGridView::OnEditReplace()
{
}

void CGridView::OnEditRepeat()
{
}


// CGridView message handlers


int CGridView::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
    GRIDCREATESTRUCT gcs;
    memset(&gcs, 0, sizeof(GRIDCREATESTRUCT));
    gcs.cbSize = sizeof(GRIDCREATESTRUCT);
    gcs.nWidth = 702;
    gcs.nHeight = 4096;
    gcs.nDefColWidth = 150;
    gcs.nDefRowHeight = 40;
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
