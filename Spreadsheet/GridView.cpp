// GridView.cpp : implementation file
//

#include "pch.h"
#include "Spreadsheet.h"
#include "GridView.h"


// CGridView

IMPLEMENT_DYNCREATE(CGridView, CCtrlView)

CGridView::CGridView() : CCtrlView(GRID_WNDCLASS_NAME, 0)
{

}

CGridView::~CGridView()
{
}

BEGIN_MESSAGE_MAP(CGridView, CView)
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
