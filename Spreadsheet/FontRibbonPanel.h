#pragma once

// FontRibbonPanel.h
#pragma once
#include <afxribbonbar.h>

class CFontRibbonPanel : public CMFCRibbonPanel
{
public:
    CFontRibbonPanel(LPCTSTR lpszName, HICON hIcon);

protected:
    virtual void OnLayout();
    virtual void OnDraw(CDC* pDC);
};
