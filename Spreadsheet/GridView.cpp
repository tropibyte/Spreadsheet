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
#endif
#endif //_DEBUG


// CGridView message handlers


int CGridView::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
    GRIDCREATESTRUCT gcs;
    memset(&gcs, 0, sizeof(GRIDCREATESTRUCT));
    gcs.cbSize = sizeof(GRIDCREATESTRUCT);
    gcs.nWidth = 702;
    gcs.nHeight = 32768;

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

