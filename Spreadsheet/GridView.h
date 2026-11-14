#pragma once

#include <vector>
#include "GridCtrl.h"

// CGridView view
const DWORD GRID_DEFAULT_STYLE = WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | WS_HSCROLL;

// The formula bar above the grid, as in Excel. The name box shows the active
// cell's reference; type a reference ("B7", "A1:C3") or a cell's name there
// and press Enter to go to it. The input box shows the active cell's content
// as it was typed -- a formula's source, not its result -- and follows the
// in-cell editor while a cell is being edited. Change it there and press
// Enter (or Tab) to write it back to the cell, or Escape to put it back.
class CFormulaBar : public CWnd
{
public:
	BOOL Create(CWnd* pParent, UINT nID, CGridCtrl* pGrid);
	int GetHeight() const { return m_nHeight; }
	// Reload both boxes from the grid's active cell. Text typed into the
	// input box is written to its cell first.
	void Refresh();
	// Show the in-cell editor's text while a cell is being edited.
	void ShowEditorText(const GRIDPOINT& cell, LPCWSTR text);
	// Write text typed into the input box to the cell it was typed for.
	// Returns TRUE if anything was written.
	BOOL CommitInput();
	// The bar's edit box that has the keyboard focus, or nullptr.
	CEdit* FocusedEdit();
	virtual BOOL PreTranslateMessage(MSG* pMsg);

protected:
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnPaint();
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	afx_msg void OnInputChange();
	afx_msg void OnInputKillFocus();
	afx_msg void OnNameKillFocus();
	DECLARE_MESSAGE_MAP()

private:
	void SetInputText(LPCWSTR text);
	void ShowName(const GRIDPOINT& cell);
	void LeaveInput(int dRow, int dCol);   // Enter, Tab: commit and move
	void CancelInput();                    // Escape
	void GoToName();                       // Enter in the name box
	bool GridAlive() const;

	CGridCtrl* m_pGrid = nullptr;
	CEdit m_nameBox, m_input;
	CFont m_font, m_fxFont;
	CRect m_fxRect;
	int m_nHeight = 26;
	GRIDPOINT m_target{};       // the cell the input box belongs to
	CStringW m_original;        // what the input box held when it was loaded
	bool m_bDirty = false;      // the user has changed the input box since
	bool m_bLoading = false;    // EN_CHANGE caused by SetInputText itself
};

class CGridView : public CView
{
	DECLARE_DYNCREATE(CGridView)

protected:
	CGridView();           // protected constructor used by dynamic creation
	virtual ~CGridView();
	CGridCtrl m_wndGridCtrl;
	CFormulaBar m_formulaBar;   // name box and input box above the grid
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
        // Keep the formula bar on the active cell (GN_SELCHANGE,
        // GN_CELLCHANGED, GN_EDITCOMMIT, GN_EDITCANCEL), and on the in-cell
        // editor's text while a cell is edited (GN_EDITSTART, GN_EDITCHANGE).
        afx_msg void OnGridRefreshBar(NMHDR* pNMHDR, LRESULT* pResult);
        afx_msg void OnGridEditText(NMHDR* pNMHDR, LRESULT* pResult);
};


