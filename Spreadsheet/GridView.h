#pragma once

#include <vector>
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
	// Page breaks worked out in OnBeginPrinting: the first column of each
	// page across and the first row of each page down.
	std::vector<UINT> m_printColStarts, m_printRowStarts;
	GRIDSELECTION m_printRange{};

public:
        CGridCtrl& GetGridCtrl() { return m_wndGridCtrl; }
        // Commit text still in the cell editor (before saving or printing).
        void CommitEdit();
        virtual void OnDraw(CDC* pDC);      // overridden to draw this view
#ifdef _DEBUG
	virtual void AssertValid() const;
#ifndef _WIN32_WCE
	virtual void Dump(CDumpContext& dc) const;
#endif
#endif
protected:
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnPrint(CDC* pDC, CPrintInfo* pInfo);

	DECLARE_MESSAGE_MAP()
public:
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnSetFocus(CWnd* pOldWnd);
	afx_msg void OnEditCut();
	afx_msg void OnEditCopy();
	afx_msg void OnEditPaste();
	afx_msg void OnEditUndo();
	afx_msg void OnEditRedo();
	afx_msg void OnUpdateEditCutCopy(CCmdUI* pCmdUI);
	afx_msg void OnUpdateEditPaste(CCmdUI* pCmdUI);
	afx_msg void OnUpdateEditUndo(CCmdUI* pCmdUI);
	afx_msg void OnUpdateEditRedo(CCmdUI* pCmdUI);
        afx_msg void OnGridClick(NMHDR* pNMHDR, LRESULT* pResult);
        // GN_CONTENTCHANGED from the grid: marks the document dirty so the
        // framework offers to save on close.
        afx_msg void OnGridContentChanged(NMHDR* pNMHDR, LRESULT* pResult);
};


