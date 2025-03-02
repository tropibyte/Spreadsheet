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

// SpreadsheetView.h : interface of the CSpreadsheetView class
//

#pragma once
#include "GridView.h"  // Include the header for CGridView

class CSpreadsheetView : public CGridView
{
protected: // create from serialization only
	CSpreadsheetView() noexcept;
	DECLARE_DYNCREATE(CSpreadsheetView)

// Attributes
public:
	CSpreadsheetDoc* GetDocument() const;

// Operations
public:

// Overrides
public:
	virtual void OnDraw(CDC* pDC);  // overridden to draw this view
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
protected:
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrinting(CDC* pDC, CPrintInfo* pInfo);

// Implementation
public:
	virtual ~CSpreadsheetView();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	afx_msg void OnFilePrintPreview();
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnContextMenu(CWnd* pWnd, CPoint point);
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnEditPaste();
	afx_msg void OnEditCut();
	afx_msg void OnEditCopy();
	afx_msg void OnEditSelectAll();
	afx_msg void OnCellGoto();
	afx_msg void OnCellFont();
};

#ifndef _DEBUG  // debug version in SpreadsheetView.cpp
inline CSpreadsheetDoc* CSpreadsheetView::GetDocument() const
   { return reinterpret_cast<CSpreadsheetDoc*>(m_pDocument); }
#endif

