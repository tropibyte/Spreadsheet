// AlignmentDlg.cpp — Cell Alignment dialog.

#include "pch.h"
#include "Spreadsheet.h"
#include "AlignmentDlg.h"
#include "afxdialogex.h"

IMPLEMENT_DYNAMIC(CAlignmentDlg, CDialogEx)

CAlignmentDlg::CAlignmentDlg(CWnd* pParent)
        : CDialogEx(IDD_DLG_ALIGNMENT, pParent)
        , m_hAlign(DT_LEFT)
        , m_vAlign(DT_VCENTER)
        , m_bWrap(FALSE)
{
}

CAlignmentDlg::~CAlignmentDlg() = default;

void CAlignmentDlg::DoDataExchange(CDataExchange* pDX)
{
        CDialogEx::DoDataExchange(pDX);
        DDX_Control(pDX, IDC_COMBO_HALIGN, m_cbHAlign);
        DDX_Control(pDX, IDC_COMBO_VALIGN, m_cbVAlign);
}

BEGIN_MESSAGE_MAP(CAlignmentDlg, CDialogEx)
        ON_BN_CLICKED(IDOK, &CAlignmentDlg::OnBnClickedOk)
END_MESSAGE_MAP()

namespace {
        // Combo index <-> DT_* mapping. Keep in lockstep with the AddString
        // order in OnInitDialog.
        constexpr UINT kHAlign[] = { DT_LEFT,    DT_CENTER,  DT_RIGHT };
        constexpr UINT kVAlign[] = { DT_TOP,     DT_VCENTER, DT_BOTTOM };

        int findIndex(const UINT* arr, size_t n, UINT v)
        {
                for (size_t i = 0; i < n; ++i)
                        if (arr[i] == v) return (int)i;
                return 0;
        }
}

BOOL CAlignmentDlg::OnInitDialog()
{
        CDialogEx::OnInitDialog();

        m_cbHAlign.AddString(_T("Left"));
        m_cbHAlign.AddString(_T("Center"));
        m_cbHAlign.AddString(_T("Right"));
        m_cbHAlign.SetCurSel(findIndex(kHAlign, _countof(kHAlign), m_hAlign));

        m_cbVAlign.AddString(_T("Top"));
        m_cbVAlign.AddString(_T("Middle"));
        m_cbVAlign.AddString(_T("Bottom"));
        m_cbVAlign.SetCurSel(findIndex(kVAlign, _countof(kVAlign), m_vAlign));

        CheckDlgButton(IDC_CHECK_WRAP, m_bWrap ? BST_CHECKED : BST_UNCHECKED);

        return TRUE;
}

void CAlignmentDlg::OnBnClickedOk()
{
        int hi = m_cbHAlign.GetCurSel();
        int vi = m_cbVAlign.GetCurSel();
        if (hi >= 0 && hi < (int)_countof(kHAlign)) m_hAlign = kHAlign[hi];
        if (vi >= 0 && vi < (int)_countof(kVAlign)) m_vAlign = kVAlign[vi];
        m_bWrap = (IsDlgButtonChecked(IDC_CHECK_WRAP) == BST_CHECKED);
        CDialogEx::OnOK();
}
