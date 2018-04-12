#pragma once


// CAddFolderDlg dialog

class CAddFolderDlg : public CDialog
{
	DECLARE_DYNAMIC(CAddFolderDlg)

public:
	CAddFolderDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CAddFolderDlg();

// Dialog Data
	enum { IDD = IDD_DIALOG_ADD };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	CString folderName;
	afx_msg void OnBnClickedOk();
};
