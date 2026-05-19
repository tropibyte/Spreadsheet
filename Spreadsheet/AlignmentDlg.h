#pragma once

#include "../Grid32/grid32.h"

// CAlignmentDlg — Excel-style "Format Cells > Alignment" dialog.
// Wraps GM_SETALIGN / GM_GETALIGN. Roadmapped controls (indent, orientation,
// shrink-to-fit, merge) are present and disabled until the Grid32 model
// supports them; see [[grid32-roadmap]] item #12.
class CAlignmentDlg : public CDialogEx
{
        DECLARE_DYNAMIC(CAlignmentDlg)

public:
        CAlignmentDlg(CWnd* pParent = nullptr);
        virtual ~CAlignmentDlg();

#ifdef AFX_DESIGN_TIME
        enum { IDD = IDD_DLG_ALIGNMENT };
#endif

        UINT m_hAlign;   // DT_LEFT / DT_CENTER / DT_RIGHT
        UINT m_vAlign;   // DT_TOP / DT_VCENTER / DT_BOTTOM
        BOOL m_bWrap;

protected:
        virtual void DoDataExchange(CDataExchange* pDX);
        virtual BOOL OnInitDialog();

        CComboBox m_cbHAlign;
        CComboBox m_cbVAlign;

        DECLARE_MESSAGE_MAP()
public:
        afx_msg void OnBnClickedOk();
};
