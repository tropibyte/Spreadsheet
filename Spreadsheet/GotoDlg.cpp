// GotoDlg.cpp : implementation file
//

#include "pch.h"
#include "Spreadsheet.h"
#include "GotoDlg.h"
#include "afxdialogex.h"

// CGotoDlg dialog

IMPLEMENT_DYNAMIC(CGotoDlg, CDialogEx)

CGotoDlg::CGotoDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_DLG_GOTOCELL, pParent), nWhich(0)
	, m_cellRef(_T(""))
{

}

CGotoDlg::~CGotoDlg()
{
}

void CGotoDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT_WHERE, m_EditReference);
	DDX_Text(pDX, IDC_EDIT_WHERE, m_cellRef);
}


BEGIN_MESSAGE_MAP(CGotoDlg, CDialogEx)
	ON_BN_CLICKED(IDC_RADIO_COORDINATES, &CGotoDlg::OnBnClickedRadioCoordinates)
	ON_BN_CLICKED(IDC_RADIO_NAME, &CGotoDlg::OnBnClickedRadioName)
	ON_BN_CLICKED(IDOK, &CGotoDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// CGotoDlg message handlers


BOOL CGotoDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	((CButton*)GetDlgItem(IDC_RADIO_COORDINATES))->SetCheck(BST_CHECKED);
	((CButton*)GetDlgItem(IDC_RADIO_NAME))->SetCheck(BST_UNCHECKED);

	GetDlgItem(IDC_EDIT_WHERE)->SetFocus();

	return TRUE;  // return TRUE unless you set the focus to a control
				  // EXCEPTION: OCX Property Pages should return FALSE
}


void CGotoDlg::OnBnClickedRadioCoordinates()
{
	((CButton*)GetDlgItem(IDC_RADIO_COORDINATES))->SetCheck(BST_CHECKED);
	((CButton*)GetDlgItem(IDC_RADIO_NAME))->SetCheck(BST_UNCHECKED);
}


void CGotoDlg::OnBnClickedRadioName()
{
	((CButton*)GetDlgItem(IDC_RADIO_COORDINATES))->SetCheck(BST_UNCHECKED);
	((CButton*)GetDlgItem(IDC_RADIO_NAME))->SetCheck(BST_CHECKED);
}


void CGotoDlg::OnBnClickedOk()
{
	if (((CButton*)GetDlgItem(IDC_RADIO_COORDINATES))->GetCheck() == BST_CHECKED)
		nWhich = SETBYCOORDINATE;
	else if (((CButton*)GetDlgItem(IDC_RADIO_NAME))->GetCheck() == BST_CHECKED)
		nWhich = SETBYNAME;

	CDialogEx::OnOK();
}
