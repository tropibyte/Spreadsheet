// GridCtrl.cpp : implementation file
//

#include "pch.h"
#include "Spreadsheet.h"
#include "GridCtrl.h"


// CGridCtrl

IMPLEMENT_DYNAMIC(CGridCtrl, CWnd)

CGridCtrl::CGridCtrl()
{

}

CGridCtrl::~CGridCtrl()
{
}

HMODULE CGridCtrl::InitGridControl()
{
	return LoadLibrary(_T("Grid32.dll"));
}

BOOL CGridCtrl::Create(DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, const GRIDCREATESTRUCT& gcs)
{
    // Ensure the style includes WS_CHILD
    dwStyle |= WS_CHILD;
    if (pParentWnd == nullptr)
        return FALSE;

    // Call CWnd::Create with the required parameters
    return CWnd::Create(
        GRID_WNDCLASS_NAME,       // Class name
        NULL,           // No window name
        dwStyle,        // Window style
        rect,           // Position and size
        pParentWnd,     // Parent window
        nID,            // Control ID
        NULL            // Additional creation parameters
    );
}


BEGIN_MESSAGE_MAP(CGridCtrl, CWnd)
END_MESSAGE_MAP()



// CGridCtrl message handlers


