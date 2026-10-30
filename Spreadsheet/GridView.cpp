// GridView.cpp : implementation file
//

#include "pch.h"
#include "Spreadsheet.h"
#include "GridView.h"
#include "MainFrm.h"
#include <commctrl.h>


// CGridView
#define GRIDCTRL_ID 1

IMPLEMENT_DYNCREATE(CGridView, CView)

CGridView::CGridView()
{

}

CGridView::~CGridView()
{
}

// Cut, Copy, Paste, Undo and Redo live here, next to the grid they act on.
// CSpreadsheetView used to bind the same IDs to MessageBox stubs, which won
// because MFC routes commands to the most-derived view first (H1, H3).
BEGIN_MESSAGE_MAP(CGridView, CView)
	ON_WM_CREATE()
	ON_WM_SIZE()
	ON_WM_SETFOCUS()
	ON_COMMAND(ID_EDIT_CUT, &CGridView::OnEditCut)
	ON_COMMAND(ID_EDIT_COPY, &CGridView::OnEditCopy)
	ON_COMMAND(ID_EDIT_PASTE, &CGridView::OnEditPaste)
	ON_COMMAND(ID_EDIT_UNDO, &CGridView::OnEditUndo)
	ON_COMMAND(ID_EDIT_REDO, &CGridView::OnEditRedo)
	ON_UPDATE_COMMAND_UI(ID_EDIT_CUT, &CGridView::OnUpdateEditCutCopy)
	ON_UPDATE_COMMAND_UI(ID_EDIT_COPY, &CGridView::OnUpdateEditCutCopy)
	ON_UPDATE_COMMAND_UI(ID_EDIT_PASTE, &CGridView::OnUpdateEditPaste)
	ON_UPDATE_COMMAND_UI(ID_EDIT_UNDO, &CGridView::OnUpdateEditUndo)
	ON_UPDATE_COMMAND_UI(ID_EDIT_REDO, &CGridView::OnUpdateEditRedo)
    ON_NOTIFY(NM_CLICK, GRIDCTRL_ID, &CGridView::OnGridClick)
    ON_NOTIFY(GN_CONTENTCHANGED, GRIDCTRL_ID, &CGridView::OnGridContentChanged)
END_MESSAGE_MAP()


// CGridView drawing

void CGridView::OnDraw(CDC* /*pDC*/)
{
	// The grid is a child window and paints itself; printing goes through
	// OnPrint, which renders cells into the printer DC.
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

void CGridView::CommitEdit()
{
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
        m_wndGridCtrl.CommitEdit(TRUE);
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

void CGridView::OnEditUndo()
{
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
        m_wndGridCtrl.Undo();
}

void CGridView::OnEditRedo()
{
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
        m_wndGridCtrl.Redo();
}

// Undo, Redo and Paste are greyed out when there is nothing to do (M17).
void CGridView::OnUpdateEditCutCopy(CCmdUI* pCmdUI)
{
    pCmdUI->Enable(m_wndGridCtrl.GetSafeHwnd() != NULL);
}

void CGridView::OnUpdateEditPaste(CCmdUI* pCmdUI)
{
    static const UINT cfGridCells = ::RegisterClipboardFormat(_T("Grid32.Cells"));
    pCmdUI->Enable(m_wndGridCtrl.GetSafeHwnd() != NULL &&
        (::IsClipboardFormatAvailable(CF_UNICODETEXT) || ::IsClipboardFormatAvailable(cfGridCells)));
}

void CGridView::OnUpdateEditUndo(CCmdUI* pCmdUI)
{
    pCmdUI->Enable(m_wndGridCtrl.GetSafeHwnd() != NULL && m_wndGridCtrl.CanUndo());
}

void CGridView::OnUpdateEditRedo(CCmdUI* pCmdUI)
{
    pCmdUI->Enable(m_wndGridCtrl.GetSafeHwnd() != NULL && m_wndGridCtrl.CanRedo());
}


// CGridView printing (H12)

BOOL CGridView::OnPreparePrinting(CPrintInfo* pInfo)
{
    CommitEdit();
    return DoPreparePrinting(pInfo);
}

// Lay the used part of the sheet out on pages: columns left to right until
// the page is full, then a new page across; rows likewise down. Pages are
// numbered down first, then across, as Excel does.
void CGridView::OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo)
{
    m_printColStarts.assign(1, 0);
    m_printRowStarts.assign(1, 0);
    GRIDSELECTION used{};
    if (!m_wndGridCtrl.GetUsedRange(used))
        used = GRIDSELECTION{};
    m_printRange = GRIDSELECTION{ { 0, 0 }, used.end };

    const int dpiX = pDC->GetDeviceCaps(LOGPIXELSX);
    const int dpiY = pDC->GetDeviceCaps(LOGPIXELSY);
    // Half an inch of margin on every side of the printable area.
    const long pageW = pDC->GetDeviceCaps(HORZRES) - dpiX;
    const long pageH = pDC->GetDeviceCaps(VERTRES) - dpiY;

    long width = 0;
    for (UINT c = 0; c <= used.end.nCol; ++c)
    {
        const long w = MulDiv((int)m_wndGridCtrl.GetColumnWidth(c), dpiX, 96);
        if (width > 0 && width + w > pageW)
        {
            m_printColStarts.push_back(c);
            width = 0;
        }
        width += w;
    }
    long height = 0;
    for (UINT r = 0; r <= used.end.nRow; ++r)
    {
        const long h = MulDiv((int)m_wndGridCtrl.GetRowHeight(r), dpiY, 96);
        if (h == 0)
            continue;   // hidden by a filter
        if (height > 0 && height + h > pageH)
        {
            m_printRowStarts.push_back(r);
            height = 0;
        }
        height += h;
    }
    pInfo->SetMaxPage((UINT)(m_printColStarts.size() * m_printRowStarts.size()));
}

void CGridView::OnPrint(CDC* pDC, CPrintInfo* pInfo)
{
    const size_t rowPages = m_printRowStarts.size();
    const size_t page = pInfo->m_nCurPage - 1;
    const size_t across = page / rowPages, down = page % rowPages;
    if (across >= m_printColStarts.size())
        return;

    GCRENDERSTRUCT rs{};
    rs.m_cbSize = sizeof(GCRENDERSTRUCT);
    // In print preview m_hDC is the screen and the attribute DC the printer;
    // the page is laid out at the printer's resolution either way.
    rs.m_hDC = pDC->m_hDC;
    rs.m_nLogPixelsX = pDC->GetDeviceCaps(LOGPIXELSX);
    rs.m_nLogPixelsY = pDC->GetDeviceCaps(LOGPIXELSY);
    rs.m_ptOrigin = { rs.m_nLogPixelsX / 2, rs.m_nLogPixelsY / 2 };
    rs.m_range.start = { m_printRowStarts[down], m_printColStarts[across] };
    rs.m_range.end.nRow = down + 1 < rowPages ? m_printRowStarts[down + 1] - 1 : m_printRange.end.nRow;
    rs.m_range.end.nCol = across + 1 < m_printColStarts.size() ? m_printColStarts[across + 1] - 1 : m_printRange.end.nCol;
    rs.m_bGridLines = TRUE;
    m_wndGridCtrl.RenderRange(rs);
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

// The view is only a frame for the grid; focus belongs to the grid, which is
// where typing starts an edit. The view used to keep focus and forward key
// presses, so characters (WM_CHAR) never reached the grid at all.
void CGridView::OnSetFocus(CWnd* pOldWnd)
{
    CView::OnSetFocus(pOldWnd);
    if (m_wndGridCtrl.GetSafeHwnd() != NULL)
        m_wndGridCtrl.SetFocus();
}

void CGridView::OnGridClick(NMHDR* pNMHDR, LRESULT* pResult)
{
    *pResult = 0;
    LPGRIDNMHDR p = reinterpret_cast<LPGRIDNMHDR>(pNMHDR);
    // App-level handlers live on CMainFrame (Format Painter etc.); CGridView
    // just forwards. If/when more click consumers appear, generalize this
    // into a dispatch table.
    if (CMainFrame* pFrame = DYNAMIC_DOWNCAST(CMainFrame, AfxGetMainWnd()))
        pFrame->ConsumeGridClick(&m_wndGridCtrl, p->cellPt);
}

// The grid owns the data, so the document has no other way to know it changed:
// nothing routes edits through CDocument. Without this the modified flag was
// never set, and closing after editing threw the work away without asking.
void CGridView::OnGridContentChanged(NMHDR* /*pNMHDR*/, LRESULT* pResult)
{
    if (CDocument* pDoc = GetDocument())
        pDoc->SetModifiedFlag(TRUE);
    if (pResult)
        *pResult = 0;
}
