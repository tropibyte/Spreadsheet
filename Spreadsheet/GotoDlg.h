#pragma once


// CGotoDlg dialog

class CGotoDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CGotoDlg)

public:
	CGotoDlg(CWnd* pParent = nullptr);   // standard constructor
	virtual ~CGotoDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DLG_GOTOCELL };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedRadioCoordinates();
	afx_msg void OnBnClickedRadioName();
	CEdit m_EditReference;
	short nWhich;
	afx_msg void OnBnClickedOk();
	CString m_cellRef;
};
