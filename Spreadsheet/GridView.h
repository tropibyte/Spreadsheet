#pragma once

#include "GridCtrl.h"

// CGridView view
const DWORD GRID_DEFAULT_STYLE = WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | WS_HSCROLL;

class CGridView : public CView
{
	DECLARE_DYNCREATE(CGridView)

protected:
	CGridView();           // protected constructor used by dynamic creation
	virtual ~CGridView();
	CGridCtrl m_wndGridCtrl;

public:
	virtual void OnDraw(CDC* pDC);      // overridden to draw this view
#ifdef _DEBUG
	virtual void AssertValid() const;
#ifndef _WIN32_WCE
	virtual void Dump(CDumpContext& dc) const;
#endif
#endif

protected:
	DECLARE_MESSAGE_MAP()
public:
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSize(UINT nType, int cx, int cy);
};


