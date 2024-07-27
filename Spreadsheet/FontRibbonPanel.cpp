#include "pch.h"
#include "FontRibbonPanel.h"

CFontRibbonPanel::CFontRibbonPanel(LPCTSTR lpszName, HICON hIcon)
    : CMFCRibbonPanel(lpszName, hIcon)
{
}

void x() {}
void CFontRibbonPanel::OnLayout()
{
    x();
}

void CFontRibbonPanel::OnDraw(CDC* /*pDC*/)
{
    
}