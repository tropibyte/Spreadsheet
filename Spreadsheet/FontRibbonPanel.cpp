#include "pch.h"
#include "FontRibbonPanel.h"

CFontRibbonPanel::CFontRibbonPanel(LPCTSTR lpszName, HICON hIcon)
    : CMFCRibbonPanel(lpszName, hIcon)
{
}

void CFontRibbonPanel::OnLayout()
{
    CMFCRibbonPanel::OnLayout();
    int x = m_rect.left + 2;
    int y = m_rect.top + 2;

    for (int i = 0; i < m_arElements.GetSize(); ++i)
    {
        CMFCRibbonBaseElement* pElem = m_arElements[i];
        if (!pElem)
            continue;
        CSize sz = pElem->GetRegularSize(NULL);
        pElem->SetRect(CRect(x, y, x + sz.cx, y + sz.cy));
        x += sz.cx + 2;
    }
}

void CFontRibbonPanel::OnDraw(CDC* pDC)
{
    CMFCRibbonPanel::OnDraw(pDC);
    for (int i = 0; i < m_arElements.GetSize(); ++i)
    {
        CMFCRibbonBaseElement* pElem = m_arElements[i];
        if (pElem)
            pElem->OnDraw(pDC);
    }
}