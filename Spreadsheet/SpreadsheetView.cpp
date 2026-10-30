// This MFC Samples source code demonstrates using MFC Microsoft Office Fluent User Interface
// (the "Fluent UI") and is provided only as referential material to supplement the
// Microsoft Foundation Classes Reference and related electronic documentation
// included with the MFC C++ library software.
// License terms to copy, use or distribute the Fluent UI are available separately.
// To learn more about our Fluent UI licensing program, please visit
// https://go.microsoft.com/fwlink/?LinkId=238214.
//
// Copyright (C) Microsoft Corporation
// All rights reserved.

// SpreadsheetView.cpp : implementation of the CSpreadsheetView class
//

#include "pch.h"
#include "framework.h"
// SHARED_HANDLERS can be defined in an ATL project implementing preview, thumbnail
// and search filter handlers and allows sharing of document code with that project.
#ifndef SHARED_HANDLERS
#include "Spreadsheet.h"
#endif

#include "SpreadsheetDoc.h"
#include "SpreadsheetView.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CSpreadsheetView

IMPLEMENT_DYNCREATE(CSpreadsheetView, CGridView)

BEGIN_MESSAGE_MAP(CSpreadsheetView, CGridView)
	// Standard printing commands
	ON_COMMAND(ID_FILE_PRINT, &CGridView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, &CGridView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, &CSpreadsheetView::OnFilePrintPreview)
	ON_WM_CONTEXTMENU()
	ON_WM_RBUTTONUP()
	// Cut/Copy/Paste/Undo/Redo are CGridView's; Select All, Go To, Cell Font,
	// Find and Replace are CMainFrame's. The stubs this map used to bind to
	// those IDs shadowed the working handlers (H1, H2, M16).
END_MESSAGE_MAP()

// CSpreadsheetView construction/destruction

CSpreadsheetView::CSpreadsheetView() noexcept
{
	// TODO: add construction code here

}

CSpreadsheetView::~CSpreadsheetView()
{
}

BOOL CSpreadsheetView::PreCreateWindow(CREATESTRUCT& cs)
{
	// The grid's GS_* styles belong on the grid child (CGridView::OnCreate);
	// OR'd into the view's own style they only collided with the low-word
	// control style bits (M21).
	cs.style |= WS_VISIBLE;

	return CGridView::PreCreateWindow(cs);
}


// CSpreadsheetView printing


void CSpreadsheetView::OnFilePrintPreview()
{
#ifndef SHARED_HANDLERS
	AFXPrintPreview(this);
#endif
}

void CSpreadsheetView::OnRButtonUp(UINT /* nFlags */, CPoint point)
{
	ClientToScreen(&point);
	OnContextMenu(this, point);
}

void CSpreadsheetView::OnContextMenu(CWnd* /* pWnd */, CPoint point)
{
#ifndef SHARED_HANDLERS
	theApp.GetContextMenuManager()->ShowPopupMenu(IDR_POPUP_EDIT, point.x, point.y, this, TRUE);
#endif
}


// CSpreadsheetView diagnostics

#ifdef _DEBUG
void CSpreadsheetView::AssertValid() const
{
	CGridView::AssertValid();
}

void CSpreadsheetView::Dump(CDumpContext& dc) const
{
	CGridView::Dump(dc);
}

CSpreadsheetDoc* CSpreadsheetView::GetDocument() const // non-debug version is inline
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CSpreadsheetDoc)));
	return (CSpreadsheetDoc*)m_pDocument;
}
#endif //_DEBUG


// CSpreadsheetView message handlers
